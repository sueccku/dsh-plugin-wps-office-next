import { readFileSync } from "node:fs";
const text = readFileSync("mcp/scripts/wps-com.ps1", "utf8");
const lines = text.split(/\r?\n/);
const actions = process.argv.slice(2);
for (const a of actions) {
  const start = lines.findIndex((l) => l === '    "' + a + '" {');
  if (start < 0) { console.log("=== " + a + "  NOT FOUND"); continue; }
  let end = start + 1;
  while (end < lines.length && lines[end] !== "    }") end++;
  const body = lines.slice(start + 1, end);
  const used = [...new Set(body.join("\n").match(/\$p\.[A-Za-z_][A-Za-z0-9_]*/g) || [])].map((s) => s.slice(3));
  console.log("=== " + a + "  (lines " + (start + 1) + "-" + (end + 1) + ")");
  console.log("    reads params: " + (used.length ? used.join(", ") : "(none)"));
  if (body.length <= 14) { console.log(body.map((l) => "      " + l.trim()).join("\n")); }
}
