// Duplicate report: tools that drive the same COM action, with their parameter shapes.
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

const tools = new Map();
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
    if (ev.kind === "tool") {
      cur = { name: ev.name, file: relative(root, f).replace(/\\/g, "/"), actions: new Set() };
      // capture the definition block for parameter extraction
      const after = text.slice(ev.i);
      const end = after.indexOf("};");
      const block = after.slice(0, end < 0 ? 900 : end + 2);
      const props = [...block.matchAll(/([A-Za-z_][A-Za-z0-9_]*):\s*\{\s*\n?\s*type:\s*'(\w+)'/g)].map((x) => x[1]);
      const req = /required:\s*\[([^\]]*)\]/.exec(block);
      cur.params = props.filter((p) => p !== "inputSchema").join(",");
      cur.required = (req ? req[1] : "").replace(/'/g, "").replace(/\s+/g, " ").trim();
      tools.set(ev.name, cur);
    } else if (cur) cur.actions.add(ev.name);
  }
}

const appOf = (n) => (/^wps_(excel|word|ppt)_/.exec(n) || [])[1] || "common";
const byAction = {};
for (const t of tools.values()) for (const a of t.actions) (byAction[a] = byAction[a] || []).push(t.name);

console.log("SAME-APP duplicates (same action, same app) — candidates for merging:");
let sameApp = 0;
for (const [action, names] of Object.entries(byAction).sort()) {
  if (names.length < 2) continue;
  const apps = new Set(names.map(appOf));
  if (apps.size !== 1) continue;
  sameApp++;
  console.log("");
  console.log("  action " + action + "  [" + [...apps][0] + "]");
  for (const n of names) {
    const t = tools.get(n);
    console.log("    " + n);
    console.log("        file=" + t.file + "  params=(" + t.params + ")  required=(" + t.required + ")");
  }
}
console.log("");
console.log("same-app duplicate groups: " + sameApp);
console.log("");

console.log("CROSS-APP shared actions (legitimate — different application):");
for (const [action, names] of Object.entries(byAction).sort()) {
  if (names.length < 2) continue;
  const apps = new Set(names.map(appOf));
  if (apps.size === 1) continue;
  console.log("  " + action + " -> " + names.join(", "));
}
