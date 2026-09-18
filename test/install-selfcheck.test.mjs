// S8/S9: install-time self-checks. Runs the doctor and asserts the version/wiring lines it must
// print, plus the static facts those checks rely on. WPS is optional (doctor warns when absent), so
// this file is safe for the CI static gates.
// Run: node test/install-selfcheck.test.mjs
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }

const doctor = spawnSync(process.execPath, ["scripts/doctor.mjs"], { encoding: "utf8", windowsHide: true });
const out = String(doctor.stdout || "");
check("doctor exits 0 (warnings are allowed)", doctor.status === 0, "status=" + doctor.status + " " + String(doctor.stderr || "").trim().slice(0, 80));
check("doctor prints a WPS version + architecture line", /WPS version\s+\S+/.test(out), (out.split(/\r?\n/).find((l) => l.includes("WPS version")) || "(missing)").trim());
check("doctor prints the plugin wiring line", /plugin wiring/.test(out) && out.includes("wps-office-next-plugin") && out.includes("mcp-wps-office-next"), (out.split(/\r?\n/).find((l) => l.includes("plugin wiring")) || "(missing)").trim());

const patch = readFileSync("cordis.patch.yml", "utf8");
check("cordis.patch.yml declares both bundle ids", patch.includes("wps-office-next-plugin") && patch.includes("mcp-wps-office-next"), "");
check("the plugin entry stays before the mcp entry", patch.indexOf("wps-office-next-plugin") < patch.indexOf("mcp-wps-office-next"), "");
const pkg = JSON.parse(readFileSync("package.json", "utf8"));
check("package.json points dsh at the patch", !!(pkg.dsh && pkg.dsh.bundle && pkg.dsh.bundle.patch), JSON.stringify(pkg.dsh && pkg.dsh.bundle));

const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "INSTALL SELF-CHECK TESTS OK (" + results.length + ")" : "INSTALL SELF-CHECK TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);
