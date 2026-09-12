import { readFileSync, readdirSync, statSync } from "node:fs";
import { join } from "node:path";
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
for (const name of process.argv.slice(2)) {
  for (const f of files) {
    const text = readFileSync(f, "utf8");
    const i = text.indexOf("name: '" + name + "'");
    if (i < 0) continue;
    const slice = text.slice(i, i + 2600);
    // print only the handler part: from the first "Handler" or "executeMethod/invokeAction" onwards
    const h = slice.search(/(Handler|executeMethod|invokeAction)/);
    const body = slice.slice(Math.max(0, h - 120));
    const lines = body.split(/\r?\n/).slice(0, 30);
    console.log("=== " + name + "  (" + f.replace(/\\/g, "/") + ")");
    console.log(lines.map((l) => "    " + l.trim()).join("\n"));
    console.log("");
    break;
  }
}
