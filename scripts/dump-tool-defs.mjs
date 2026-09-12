import { readFileSync, readdirSync, statSync } from "node:fs";
import { join } from "node:path";

const targets = process.argv.slice(2);
function walk(dir) {
  const out = [];
  for (const e of readdirSync(dir)) {
    const p = join(dir, e);
    if (statSync(p).isDirectory()) out.push(...walk(p));
    else if (p.endsWith(".ts")) out.push(p);
  }
  return out;
}
const files = walk("mcp/src/tools");
for (const name of targets) {
  let found = false;
  for (const f of files) {
    const text = readFileSync(f, "utf8");
    const i = text.indexOf("name: '" + name + "'");
    if (i < 0) continue;
    const after = text.slice(i);
    const end = after.indexOf("};");
    const block = after.slice(0, end < 0 ? 900 : end + 2);
    const props = [...block.matchAll(/([A-Za-z_][A-Za-z0-9_]*):\s*\{\s*\n?\s*type:\s*'(\w+)'/g)].map((m) => m[1] + ":" + m[2]);
    const req = /required:\s*\[([^\]]*)\]/.exec(block);
    console.log("=== " + name + "  (" + f.replace(/\\/g, "/") + ")");
    console.log("    props: " + props.join(", "));
    console.log("    required: " + (req ? req[1].trim() : "(none)"));
    console.log("    desc: " + (/(description:\s*'([^']{0,120}))/.exec(block)?.[2] || "?"));
    found = true;
    break;
  }
  if (!found) console.log("=== " + name + " NOT FOUND");
}
