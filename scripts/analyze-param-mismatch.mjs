// Static comparison: the keys a tool sends vs the keys the bridge actually reads.
import { readFileSync, readdirSync, statSync } from "node:fs";
import { join, relative } from "node:path";

// --- bridge: action -> params read ---
const bridgeText = readFileSync("mcp/scripts/wps-com.ps1", "utf8");
const bridgeLines = bridgeText.split(/\r?\n/);
const reads = new Map();
for (let i = 0; i < bridgeLines.length; i++) {
  const m = /^    "([A-Za-z][A-Za-z0-9_]*)" \{$/.exec(bridgeLines[i]);
  if (!m) continue;
  let end = i + 1;
  while (end < bridgeLines.length && bridgeLines[end] !== "    }") end++;
  const body = bridgeLines.slice(i + 1, end).join("\n");
  const keys = new Set([...body.matchAll(/\$p\.([A-Za-z_][A-Za-z0-9_]*)/g)].map((x) => x[1]));
  reads.set(m[1], keys);
}

// --- tools: name -> [{action, keys, source}] ---
function walk(dir) {
  const out = [];
  for (const e of readdirSync(dir)) {
    const p = join(dir, e);
    if (statSync(p).isDirectory()) out.push(...walk(p));
    else if (p.endsWith(".ts")) out.push(p);
  }
  return out;
}
function literalKeys(text, start) {
  // start points just after the object's "{"
  const keys = [];
  let depth = 1;
  let i = start;
  while (i < text.length && depth > 0) {
    const ch = text[i];
    if (ch === "{") depth++;
    else if (ch === "}") depth--;
    else if (depth === 1) {
      const keyMatch = /^([A-Za-z_$][A-Za-z0-9_$]*)\s*(?::|,|\})/.exec(text.slice(i, i + 60));
      if (keyMatch) { keys.push(keyMatch[1]); i += keyMatch[1].length; continue; }
    }
    i++;
  }
  return keys;
}
const tools = [];
for (const f of walk("mcp/src/tools").sort()) {
  const text = readFileSync(f, "utf8");
  const toolRe = /name:\s*'(wps_[a-z0-9_]+)'/g;
  const callRe = /(?:executeMethod|invokeAction)[^(]{0,300}\(\s*'([A-Za-z0-9_]+)'\s*,/g;
  const events = [];
  let m;
  toolRe.lastIndex = 0;
  while ((m = toolRe.exec(text)) !== null) {
    const after = text.slice(m.index);
    const end = after.indexOf("};");
    const block = after.slice(0, end < 0 ? 900 : end + 2);
    const props = [...block.matchAll(/([A-Za-z_][A-Za-z0-9_]*):\s*\{\s*\n?\s*type:\s*'(\w+)'/g)].map((x) => x[1]).filter((p) => p !== "inputSchema");
    events.push({ i: m.index, kind: "tool", name: m[1], props });
  }
  callRe.lastIndex = 0;
  while ((m = callRe.exec(text)) !== null) events.push({ i: m.index, kind: "call", action: m[1], at: m.index + m[0].length });
  events.sort((a, b) => a.i - b.i);
  let cur = null;
  for (const ev of events) {
    if (ev.kind === "tool") { cur = { name: ev.name, file: relative("mcp/src/tools", f).replace(/\\/g, "/"), props: ev.props, calls: [] }; tools.push(cur); continue; }
    if (!cur) continue;
    const rest = text.slice(ev.at, ev.at + 600);
    let keys;
    const objStart = rest.search(/\{/);
    const varStart = rest.search(/^\s*args\b/);
    if (varStart === 0) keys = cur.props.slice();
    else if (objStart >= 0 && objStart < 40) keys = literalKeys(rest, objStart + 1);
    else keys = cur.props.slice();
    cur.calls.push({ action: ev.action, keys });
  }
}

const problems = [];
for (const t of tools) {
  for (const c of t.calls) {
    const bridgeKeys = reads.get(c.action);
    if (!bridgeKeys) { problems.push({ kind: "NO-ACTION", tool: t.name, action: c.action, detail: "action not present in bridge" }); continue; }
    const dead = c.keys.filter((k) => !bridgeKeys.has(k));
    const never = [...bridgeKeys].filter((k) => !c.keys.includes(k));
    if (dead.length) problems.push({ kind: "DEAD-PARAM", tool: t.name, action: c.action, detail: "tool sends " + dead.join(",") + " but bridge reads " + [...bridgeKeys].join(",") });
    else if (never.length) problems.push({ kind: "UNSENT", tool: t.name, action: c.action, detail: "bridge reads " + never.join(",") + " which this tool never sends" });
  }
}
console.log("tools analysed: " + tools.length);
console.log("suspicious: " + problems.length);
console.log("");
for (const p of problems) {
  console.log("[" + p.kind + "] " + p.tool + "  -> " + p.action);
  console.log("    " + p.detail);
}
