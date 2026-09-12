// Budget analysis: real advertised-payload cost, marginal cost per tool, and
// name confusability (the accuracy cost of a large surface).
import { createRequire } from "node:module";
import { readFileSync } from "node:fs";
const require = createRequire(import.meta.url);
const toolset = require("../mcp/dist/server/toolset.js");

const snap = JSON.parse(readFileSync(process.argv[2], "utf8"));
const byName = new Map(snap.tools.map((t) => [t.name, t]));

// Reproduce exactly what tools/list sends for a non-full mode.
function payloadBytes(names, compact) {
  let n = 0;
  for (const name of names) {
    const t = byName.get(name);
    if (!t) continue;
    const projected = {
      name: t.name,
      description: compact ? toolset.compactDescription(t.description) : t.description,
      inputSchema: t.inputSchema
    };
    n += Buffer.byteLength(JSON.stringify(projected), "utf8");
  }
  return n;
}
const tok = (b) => Math.round(b / 3.5);

const facade = [...toolset.FACADE_TOOLS].filter((n) => byName.has(n));
const curated = [...toolset.STANDARD_TOOLS];
const order = [...facade, ...curated];

const allNames = snap.tools.map((t) => t.name);
console.log("catalog: " + allNames.length + " tools, full-description payload " + payloadBytes(allNames, false) + " B / " + tok(payloadBytes(allNames, false)) + " tokens");
console.log("");

console.log("corrected cumulative curve (compaction applied, as tools/list really sends it):");
for (const n of [4, 6, 8, 10, 12, 15, 20, 25, 30, 35, 39, order.length]) {
  const names = order.slice(0, n);
  const b = payloadBytes(names, true);
  console.log("  " + String(n).padStart(3) + " tools -> " + String(b).padStart(6) + " B, " + String(tok(b)).padStart(5) + " tokens, marginal " + (n > 4 ? Math.round((b - payloadBytes(order.slice(0, n - 1), true))) : 0) + " B");
}
const standardBytes = payloadBytes(order, true);
console.log("  standard total = " + standardBytes + " B / " + tok(standardBytes) + " tokens");
console.log("");

function appOf(name) {
  const m = /^wps_(excel|word|ppt)_/.exec(name);
  if (m) return m[1];
  if (/^wps_(common|convert)_/.test(name)) return "common";
  return "other";
}
console.log("app-scoped preset sizes (facade + that app curated list):");
for (const app of ["excel", "word", "ppt", "common"]) {
  const names = [...facade, ...curated.filter((n) => appOf(n) === app)];
  const b = payloadBytes(names, true);
  console.log("  " + app.padEnd(7) + " " + String(names.length).padStart(3) + " tools -> " + String(b).padStart(6) + " B, " + String(tok(b)).padStart(5) + " tokens");
}
console.log("");

// Confusability: how many tool pairs look alike, at different surface sizes.
function tokensOf(name) {
  return name.replace(/^wps_/, "").split("_");
}
function jaccard(a, b) {
  const A = new Set(tokensOf(a));
  const B = new Set(tokensOf(b));
  let inter = 0;
  for (const x of A) if (B.has(x)) inter++;
  const union = A.size + B.size - inter;
  return union === 0 ? 0 : inter / union;
}
function confusable(names, threshold) {
  let pairs = 0;
  for (let i = 0; i < names.length; i++) {
    for (let j = i + 1; j < names.length; j++) {
      if (jaccard(names[i], names[j]) >= threshold) pairs++;
    }
  }
  return pairs;
}
console.log("confusable tool pairs (name-token Jaccard >= 0.5):");
for (const n of [4, 10, 20, 30, order.length, allNames.length]) {
  const names = n <= order.length ? order.slice(0, n) : allNames;
  const pairs = confusable(names, 0.5);
  const possible = (names.length * (names.length - 1)) / 2;
  console.log("  " + String(n).padStart(3) + " tools -> " + String(pairs).padStart(4) + " pairs (" + (possible ? (100 * pairs / possible).toFixed(1) : "0") + "% of all pairs)");
}
console.log("");

// Examples of the most confusable advertised groups.
const adv = [...order];
const hits = [];
for (let i = 0; i < adv.length; i++) {
  for (let j = i + 1; j < adv.length; j++) {
    const s = jaccard(adv[i], adv[j]);
    if (s >= 0.5) hits.push({ a: adv[i], b: adv[j], s });
  }
}
hits.sort((x, y) => y.s - x.s);
console.log("most confusable advertised pairs:");
for (const h of hits.slice(0, 8)) console.log("  " + h.s.toFixed(2) + "  " + h.a + "  <->  " + h.b);
