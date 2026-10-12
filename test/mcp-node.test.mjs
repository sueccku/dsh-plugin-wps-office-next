// Static: the MCP stdio child must never be launched with a bare `node`.
// Background (verified against app.asar): DSH Desktop starts the Host with
// desktopNodeEnvironment(node, void 0, ...), so DSH_DESKTOP_NODE_EXECUTABLE is unset; the mcp-client
// child env is then built from scrubbedParentEnv(), which drops every DSH_* name anyway. On a machine
// with no Node.js on the host PATH, `spawn('node')` therefore fails with ENOENT, and because the row
// sets failOnStartupError: true the whole MCP entry dies with zero WPS tools registered.
// Run: node test/mcp-node.test.mjs
import { existsSync, readFileSync, writeFileSync, rmSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { homedir, tmpdir } from "node:os";
import { join } from "node:path";

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }

// --- 1. the published patch must resolve `command`, never hard-code a bare `node` ---
const patch = readFileSync("cordis.patch.yml", "utf8");
check(
  "cordis.patch.yml does not spawn a bare `node`",
  !/^\s*command:\s*node\s*$/mu.test(patch),
  (patch.split(/\r?\n/).find((l) => /^\s*command:/.test(l)) || "(no command line)").trim()
);
check(
  "cordis.patch.yml resolves `command` through !!js WPS_OFFICE_MCP_NODE",
  /^\s*command:\s*!!js\s+"process\.env\.WPS_OFFICE_MCP_NODE/mu.test(patch),
  ""
);

// --- 2. the resolver prefers the DSH-managed runtime Node, then process.execPath ---
const mcpNode = await import("../scripts/lib/mcp-node.mjs");
const candidates = mcpNode.nodeExecutableCandidates();
check("candidate list ends with process.execPath", candidates[candidates.length - 1] === process.execPath, candidates.join(" | "));
check(
  "candidate list includes the DSH-managed runtime node",
  candidates.some((c) => /dsh-runtimes[\\/]dsh-primary-runtime[\\/]dependencies[\\/]node[\\/]bin[\\/]/u.test(c)),
  candidates[0] ?? "(none)"
);
const resolved = mcpNode.resolveNodeExecutable();
check("resolveNodeExecutable() returns an absolute existing path on this machine", existsSync(resolved), resolved);

// DSH_HOME is used when set, so a non-default install layout still resolves.
{
  const saved = process.env.DSH_HOME;
  process.env.DSH_HOME = "D:\\dsh-home-probe";
  const withHome = mcpNode.nodeExecutableCandidates();
  if (saved === undefined) delete process.env.DSH_HOME; else process.env.DSH_HOME = saved;
  check(
    "DSH_HOME is honoured when set",
    withHome.some((c) => c.startsWith("D:\\dsh-home-probe")),
    withHome[0] ?? "(none)"
  );
}
{
  const saved = process.env.DSH_HOME;
  delete process.env.DSH_HOME;
  const noHome = mcpNode.nodeExecutableCandidates();
  if (saved !== undefined) process.env.DSH_HOME = saved;
  check(
    "falls back to ~/.dsh/dsh-runtimes when DSH_HOME is unset",
    noHome.some((c) => c.startsWith(join(homedir(), ".dsh", "dsh-runtimes"))),
    noHome[0] ?? "(none)"
  );
}

// WPS_OFFICE_MCP_NODE is an explicit override and must win.
{
  const saved = process.env.WPS_OFFICE_MCP_NODE;
  process.env.WPS_OFFICE_MCP_NODE = "C:\\probe\\node.exe";
  const overridden = mcpNode.nodeExecutableCandidates();
  check("WPS_OFFICE_MCP_NODE wins over the bundled runtime", overridden[0] === "C:\\probe\\node.exe", overridden[0] ?? "(none)");
  const savedExists = existsSync("C:\\probe\\node.exe");
  check(
    "a non-existent override is not selected by resolveNodeExecutable()",
    savedExists || mcpNode.resolveNodeExecutable() !== "C:\\probe\\node.exe",
    mcpNode.resolveNodeExecutable()
  );
  if (saved === undefined) delete process.env.WPS_OFFICE_MCP_NODE; else process.env.WPS_OFFICE_MCP_NODE = saved;
}

// --- 3. plugin.js must publish WPS_OFFICE_MCP_NODE for the patch to read ---
const plugin = readFileSync("plugin.js", "utf8");
check("plugin.js publishes WPS_OFFICE_MCP_NODE", /process\.env\.WPS_OFFICE_MCP_NODE\s*=/.test(plugin), "");
check("plugin.js publishes WPS_OFFICE_MCP_ENTRY", /process\.env\.WPS_OFFICE_MCP_ENTRY\s*=/.test(plugin), "");

// --- 4. the doctor's mcp node verdict must follow what the resolver actually returns ---
// Case A: no DSH-managed runtime, but the interpreter IS a real node (the documented "no node on
// the host PATH" invocation runs doctor with the runtime's own node) — that is a healthy OK.
{
  const doctor = spawnSync(process.execPath, ["scripts/doctor.mjs"], {
    encoding: "utf8",
    windowsHide: true,
    env: { ...process.env, DSH_HOME: join(tmpdir(), "dsh-doctor-no-runtime-" + process.pid) },
  });
  const out = String(doctor.stdout || "");
  const line = (out.split(/\r?\n/).find((l) => /mcp node/.test(l)) || "(missing)").trim();
  const interpreterIsNode = /(^|[/\\])node(\.exe)?$/iu.test(process.execPath);
  check("doctor still exits 0 when no DSH-managed runtime is found", doctor.status === 0, "status=" + doctor.status);
  check(
    interpreterIsNode
      ? "a real-node interpreter is reported OK (not a bogus missing-runtime warning)"
      : "an Electron interpreter with no managed runtime is WARNed",
    interpreterIsNode ? /^OK\s+mcp node/m.test(out) : /WARN\s+mcp node/.test(out),
    line
  );
}

// A user-pinned override that exists but is not a Node must be reported as a WARN, not a healthy OK.
{
  const bogus = join(tmpdir(), "dsh-not-a-node-" + process.pid + ".exe");
  writeFileSync(bogus, "not a node runtime\n");
  try {
    const doctor = spawnSync(process.execPath, ["scripts/doctor.mjs"], {
      encoding: "utf8",
      windowsHide: true,
      env: { ...process.env, WPS_OFFICE_MCP_NODE: bogus },
    });
    const out = String(doctor.stdout || "");
    const line = (out.split(/\r?\n/).find((l) => /mcp node/.test(l)) || "(missing)").trim();
    check("doctor exits 0 with a user-pinned non-Node override", doctor.status === 0, "status=" + doctor.status);
    check(
      "a user-pinned path that exists but is not Node is WARNed, not accepted as OK",
      /WARN\s+mcp node/.test(out) && !/^OK\s+mcp node/m.test(out),
      line
    );
  } finally {
    rmSync(bogus, { force: true });
  }
}

// --- 5. the shipped one-liner must satisfy the "no bare node" contract too ---
check("package.json still points dsh at the patch", (() => {
  const pkg = JSON.parse(readFileSync("package.json", "utf8"));
  return !!(pkg.dsh && pkg.dsh.bundle && pkg.dsh.bundle.patch) && pkg.files.includes("scripts");
})(), "the patch reads scripts/lib/mcp-node.mjs, so scripts/ must ship");

const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "MCP NODE TESTS OK (" + results.length + ")" : "MCP NODE TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);
