import { readFileSync } from "node:fs";
const lines = readFileSync("mcp/scripts/wps-com.ps1", "utf8").split(/\r?\n/);
for (const a of process.argv.slice(2)) {
  const s = lines.findIndex((l) => l === '    "' + a + '" {');
  if (s < 0) { console.log("=== " + a + " NOT FOUND"); continue; }
  let e = s + 1;
  while (e < lines.length && lines[e] !== "    }") e++;
  console.log("=== " + a + "  (lines " + (s + 1) + "-" + (e + 1) + ")");
  console.log(lines.slice(s + 1, e).map((l) => "    " + l.trim()).join("\n"));
  console.log("");
}
