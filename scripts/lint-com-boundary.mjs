// Lists raw COM member writes/calls still bypassing the PSObject boundary helpers.
// The PowerShell COM binder is cached per member after its first use, so a call site that can see
// different value types must go through Set-ComValue / Invoke-ComMethod.
// Run: node scripts/lint-com-boundary.mjs
import { readFileSync } from "node:fs";

const text = readFileSync("mcp/scripts/wps-com.ps1", "utf8").split(/\r?\n/);
const patterns = [
  { name: "value-write", re: /\.(Value2|Formula|NumberFormat|Text)\s*=(?!=)/, risk: "HIGH - value type varies" },
  { name: "com-call", re: /\$[A-Za-z_][A-Za-z0-9_.]*\.(Replace|Find|FindNext)\s*\(/, risk: "HIGH - argument types vary" },
  { name: "colour-write", re: /\.(Color|ForeColor\.RGB|RGB)\s*=/, risk: "LOW - always an integer" }
];
const counts = {};
const samples = {};
for (const { name, re, risk } of patterns) {
  for (let i = 0; i < text.length; i++) {
    const line = text[i];
    if (line.trim().startsWith("#")) continue;
    if (re.test(line)) {
      counts[name] = (counts[name] || 0) + 1;
      samples[name] = samples[name] || [];
      if (samples[name].length < 5) samples[name].push((i + 1) + ": " + line.trim().slice(0, 100));
    }
  }
}
let total = 0;
for (const { name, risk } of patterns) {
  total += counts[name] || 0;
  console.log(name.padEnd(14) + String(counts[name] || 0).padStart(4) + "   " + risk);
  for (const s of samples[name] || []) console.log("      " + s);
}
console.log("");
console.log("");
console.log("candidates for review (value-write + com-call, i.e. type-varying): " + (counts["value-write"] || 0) + " + " + (counts["com-call"] || 0) + " = " + (((counts["value-write"] || 0) + (counts["com-call"] || 0))));
console.log("colour-write sites are always integers and are not a binder-cache risk: " + (counts["colour-write"] || 0));
