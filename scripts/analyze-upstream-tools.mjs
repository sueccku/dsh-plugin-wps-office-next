// Fixed pass: map each upstream tool to the COM action(s) it drives.
import { readFileSync, readdirSync, statSync } from "node:fs";
import { join, relative } from "node:path";

const root = "mcp/src/tools";
function walk(dir) {
  const out = [];
  for (const e of readdirSync(dir)) {
    const p = join(dir, e);
    if (statSync(p).isDirectory()) out.push(...walk(p));
    else if (p.endsWith(".ts")) out.push(p);
  }
  return out;
}
const files = walk(root).sort();
const toolRe = /name:\s*'(wps_[a-z0-9_]+)'/g;
const actionRe = /(?:executeMethod|invokeAction)[^(]{0,300}\(\s*'([A-Za-z0-9_]+)'/g;

const tools = [];
for (const f of files) {
  const text = readFileSync(f, "utf8");
  const events = [];
  let m;
  toolRe.lastIndex = 0;
  while ((m = toolRe.exec(text)) !== null) events.push({ i: m.index, kind: "tool", name: m[1] });
  actionRe.lastIndex = 0;
  while ((m = actionRe.exec(text)) !== null) events.push({ i: m.index, kind: "action", name: m[1] });
  events.sort((a, b) => a.i - b.i);
  let cur = null;
  for (const ev of events) {
    if (ev.kind === "tool") { cur = { name: ev.name, file: relative(root, f).replace(/\\/g, "/"), actions: new Set() }; tools.push(cur); }
    else if (cur) cur.actions.add(ev.name);
  }
}
const byAction = {};
for (const t of tools) for (const a of t.actions) (byAction[a] = byAction[a] || []).push(t.name);
const reached = new Set(Object.keys(byAction));

console.log("tools=" + tools.length + "  distinct actions reached=" + reached.size);
const mapped = tools.filter((t) => t.actions.size > 0).length;
console.log("tools with a mapped action=" + mapped + "  without=" + (tools.length - mapped));
console.log("");

const shared = Object.entries(byAction).filter(([, v]) => v.length > 1).sort((a, b) => b[1].length - a[1].length);
console.log("actions driven by >1 tool (" + shared.length + "):");
for (const [a, v] of shared.slice(0, 18)) console.log("  " + a.padEnd(24) + v.length + "  " + v.join(", "));
console.log("");

const bridge = JSON.parse(readFileSync("baseline/upstream-0.1.0/com-actions.json", "utf8"));
const bridgeNames = bridge.actions.map((a) => a.action);
const unwrapped = bridgeNames.filter((a) => !reached.has(a));
console.log("bridge actions=" + bridgeNames.length + "  never driven by any tool=" + unwrapped.length);
console.log("  " + unwrapped.slice(0, 60).join(", "));
console.log("");

const dead = [...reached].filter((a) => !bridgeNames.includes(a));
console.log("actions referenced by tools but absent from the bridge=" + dead.length + (dead.length ? "  " + dead.slice(0, 20).join(", ") : ""));
console.log("");

const multi = tools.filter((t) => t.actions.size > 1);
console.log("tools driving more than one action=" + multi.length);
for (const t of multi.slice(0, 12)) console.log("  " + t.name + " -> " + [...t.actions].join(", "));
console.log("");

const byFile = {};
for (const t of tools) { const a = /^wps_(excel|word|ppt)_/.exec(t.name)?.[1] || "common"; (byFile[a] = byFile[a] || []).push(t); }
console.log("distinct actions per app:");
for (const app of ["excel", "word", "ppt", "common"]) {
  const s = new Set();
  for (const t of byFile[app] || []) for (const a of t.actions) s.add(a);
  console.log("  " + app.padEnd(7) + "tools=" + String((byFile[app] || []).length).padStart(3) + "  actions=" + s.size);
}
