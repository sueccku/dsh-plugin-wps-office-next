// Verifies the DSH entry: path publication and skill registration, without booting DSH.
import { existsSync } from "node:fs";
import { dirname, join } from "node:path";

const mod = await import("../plugin.js");
const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }

check("module exports name", mod.name === "wps-office-next", String(mod.name));
check("module injects skills", Array.isArray(mod.inject) && mod.inject.includes("skills"));
check("MCP entry resolved inside the package", existsSync(mod.mcpEntry), mod.mcpEntry);
check("COM host script resolved inside the package", existsSync(mod.hostScript), mod.hostScript);
check("env WPS_OFFICE_MCP_ENTRY published", process.env.WPS_OFFICE_MCP_ENTRY === mod.mcpEntry, process.env.WPS_OFFICE_MCP_ENTRY);

const registered = [];
const ctx = { skills: { register(skill) { registered.push(skill); return () => {}; } } };
const dispose = mod.apply(ctx);
check("registers four skills", registered.length === 4, registered.map((s) => s.name).join(", "));
check("returns a disposer", typeof dispose === "function");
for (const skill of registered) {
  check("skill " + skill.name + " has description", typeof skill.description === "string" && skill.description.length > 10);
  check("skill " + skill.name + " has whenToUse", typeof skill.whenToUse === "string" && skill.whenToUse.length > 5);
  check("skill " + skill.name + " has body", typeof skill.content === "string" && skill.content.length > 200, "len=" + (skill.content || "").length);
  check("skill " + skill.name + " frontmatter stripped", !String(skill.content).startsWith("---"));
  const dir = skill.resourceBase && skill.resourceBase.path;
  check("skill " + skill.name + " resourceBase is a directory", skill.resourceBase && skill.resourceBase.kind === "directory" && existsSync(dir));
  check("skill " + skill.name + " reference.md sits beside SKILL.md", existsSync(join(dirname(join(dir, "SKILL.md")), "reference.md")));
}
check("no skill body leaks frontmatter name", !registered.some((s) => s.content.includes("name: wps-")));

dispose();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "PLUGIN TESTS OK (" + results.length + ")" : "PLUGIN TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);
