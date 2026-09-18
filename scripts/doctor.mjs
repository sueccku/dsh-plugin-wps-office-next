// Read-only environment check for dsh-plugin-wps-office-next.
// Run: node scripts/doctor.mjs
// Prints OK / WARN / ERROR lines and exits non-zero when any ERROR is found.
import { existsSync, readFileSync, readdirSync, openSync, readSync, closeSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { homedir } from "node:os";

const root = dirname(dirname(fileURLToPath(import.meta.url)));
let errors = 0;
let warnings = 0;
function report(level, name, detail) {
  if (level === "ERROR") errors++;
  if (level === "WARN") warnings++;
  console.log(level.padEnd(5) + " " + name + (detail ? "  " + detail : ""));
}

/** Numeric dotted-version compare; missing components count as 0. */
function cmpVersion(a, b) {
  const pa = String(a).split(".").map(Number);
  const pb = String(b).split(".").map(Number);
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    const x = pa[i] || 0;
    const y = pb[i] || 0;
    if (x !== y) return x - y;
  }
  return 0;
}

/** PE COFF Machine: 0x8664 = x64, 0x014c = x86. Reads 4 KB, never executes anything. */
function peMachine(file) {
  try {
    const fd = openSync(file, "r");
    const buf = Buffer.alloc(4096);
    const n = readSync(fd, buf, 0, 4096, 0);
    closeSync(fd);
    if (n < 0x40) return 0;
    const peOff = buf.readUInt32LE(0x3c);
    if (peOff + 6 > n || buf.toString("ascii", peOff, peOff + 2) !== "PE") return 0;
    return buf.readUInt16LE(peOff + 4);
  } catch { return 0; }
}

// Platform and runtime
if (process.platform === "win32" && process.arch === "x64") report("OK", "platform", process.platform + "/" + process.arch);
else report("ERROR", "platform", "this bundle is Windows x64 only, found " + process.platform + "/" + process.arch);

const nodeMajor = Number(process.versions.node.split(".")[0]);
const nodeMinor = Number(process.versions.node.split(".")[1]);
if (nodeMajor > 22 || (nodeMajor === 22 && nodeMinor >= 19)) report("OK", "node", process.versions.node);
else report("ERROR", "node", "need 22.19+ or 24+, found " + process.versions.node);

// Windows PowerShell with an STA apartment, which WPS COM requires
const ps = spawnSync("powershell", ["-NoProfile", "-NonInteractive", "-STA", "-Command", "$PSVersionTable.PSVersion.ToString(); [System.Threading.Thread]::CurrentThread.GetApartmentState().ToString()"], { encoding: "utf8", windowsHide: true });
if (ps.error) {
  report("ERROR", "powershell", "not runnable: " + ps.error.message);
} else if (ps.status !== 0) {
  report("ERROR", "powershell", "exited with " + ps.status + ": " + String(ps.stderr || "").trim().slice(0, 120));
} else {
  const lines = String(ps.stdout || "").trim().split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const version = lines[0] || "?";
  const apartment = lines[1] || "?";
  if (apartment === "STA") report("OK", "powershell STA", "version=" + version);
  else report("ERROR", "powershell STA", "apartment is " + apartment + ", WPS COM needs STA");
}

// Bundle contents
const mcpEntry = join(root, "mcp", "dist", "index.js");
if (existsSync(mcpEntry)) report("OK", "mcp entry", mcpEntry);
else report("ERROR", "mcp entry missing", mcpEntry + " (run npm run build in mcp/)");

const hostScript = join(root, "host", "wps-com-host.ps1");
if (existsSync(hostScript)) {
  const bytes = readFileSync(hostScript).subarray(0, 3);
  const bom = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf;
  if (bom) report("OK", "com host", hostScript);
  else report("ERROR", "com host", "missing UTF-8 BOM; PowerShell 5.1 would misread its Chinese text and the host would not parse");
} else {
  report("ERROR", "com host missing", hostScript);
}

const actions = join(root, "host", "wps-actions.ps1");
if (existsSync(actions)) {
  const bytes = readFileSync(actions).subarray(0, 3);
  const bom = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf;
  if (bom) report("OK", "actions module", "UTF-8 BOM present");
  else report("ERROR", "actions module", "missing UTF-8 BOM; PowerShell 5.1 will misread Chinese text");
} else {
  report("ERROR", "actions module missing", actions + " (run scripts/build-host-actions.ps1)");
}

const skillDirs = ["wps-office-next", "wps-excel", "wps-word", "wps-ppt"];
for (const name of skillDirs) {
  const file = join(root, "skills", name, "SKILL.md");
  if (existsSync(file)) report("OK", "skill " + name);
  else report("WARN", "skill " + name + " missing", file);
}

// WPS COM registration: read-only, never launches WPS
const probe = spawnSync("powershell", ["-NoProfile", "-NonInteractive", "-STA", "-Command", [
  "$found = @()",
  "foreach ($id in 'Ket.Application','Kwps.Application','Kwpp.Application') {",
  "  try { $null = [System.Runtime.InteropServices.Marshal]::GetActiveObject($id); $found += $id } catch { }",
  "}",
  "if ($found.Count -eq 0) { 'NONE' } else { $found -join ',' }"
].join("\n")], { encoding: "utf8", windowsHide: true });
const running = String(probe.stdout || "").trim();
if (running === "NONE" || running === "") report("WARN", "WPS COM", "no running WPS instance found; start WPS 12.1+ and open a document");
else report("OK", "WPS COM", "active: " + running);

// WPS Office version and architecture precheck (S8): read-only, never launches WPS.
{
  const wpsBase = join(process.env.LOCALAPPDATA || join(homedir(), "AppData", "Local"), "Kingsoft", "WPS Office");
  const installed = existsSync(wpsBase) ? readdirSync(wpsBase).filter((d) => /^\d+\.\d+/.test(d)).sort(cmpVersion) : [];
  if (installed.length === 0) {
    report("WARN", "WPS version", "no WPS Office install found under " + wpsBase + "; install WPS 12.1+ 64-bit");
  } else {
    const latest = installed[installed.length - 1];
    const machine = peMachine(join(wpsBase, latest, "office6", "wps.exe"));
    const arch = machine === 0x8664 ? "x64" : machine === 0x14c ? "x86" : "unknown";
    if (cmpVersion(latest, "12.1") < 0) report("ERROR", "WPS version", latest + " is older than 12.1; upgrade WPS Office");
    else if (arch === "x86") report("ERROR", "WPS version", latest + " is 32-bit; this bundle needs the 64-bit WPS");
    else if (arch !== "x64") report("WARN", "WPS version", latest + " (architecture could not be confirmed)");
    else report("OK", "WPS version", latest + " " + arch);
  }
}

// Bundle wiring (S9): the published patch must declare both entries a profile expects.
{
  let patchRel = "cordis.patch.yml";
  try {
    const pkg = JSON.parse(readFileSync(join(root, "package.json"), "utf8"));
    if (pkg.dsh && pkg.dsh.bundle && pkg.dsh.bundle.patch) patchRel = pkg.dsh.bundle.patch;
  } catch { /* fall back to the default name */ }
  const patchFile = join(root, patchRel);
  if (!existsSync(patchFile)) {
    report("ERROR", "plugin wiring", "missing " + patchFile);
  } else {
    const text = readFileSync(patchFile, "utf8");
    const missing = ["wps-office-next-plugin", "mcp-wps-office-next"].filter((id) => !text.includes(id));
    if (missing.length) report("ERROR", "plugin wiring", patchRel + " is missing " + missing.join(", "));
    else report("OK", "plugin wiring", patchRel + " declares wps-office-next-plugin + mcp-wps-office-next");
  }
}

// Profile wiring, informational
const profilesDir = join(homedir(), ".dsh", "profiles");
if (existsSync(profilesDir)) {
  const wired = [];
  for (const name of readdirSync(profilesDir)) {
    const manifest = join(profilesDir, name, "package.json");
    if (!existsSync(manifest)) continue;
    try {
      const text = readFileSync(manifest, "utf8");
      if (text.includes("dsh-plugin-wps-office-next")) wired.push(name);
    } catch { /* unreadable profile */ }
  }
  if (wired.length) report("OK", "wired profiles", wired.join(", "));
  else report("WARN", "wired profiles", "no profile lists dsh-plugin-wps-office-next yet");
} else {
  report("WARN", "DSH profiles", "not found at " + profilesDir);
}

console.log("");
console.log(errors === 0 ? "DOCTOR OK" + (warnings ? " (" + warnings + " warning)" : "") : "DOCTOR FAILED (" + errors + " error, " + warnings + " warning)");
process.exit(errors === 0 ? 0 : 1);
