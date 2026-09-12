import { readFileSync } from "node:fs";

const snap = JSON.parse(readFileSync(process.argv[2], "utf8"));

function project(tools, propMax, toolMax) {
  return tools.map((t) => {
    const schema = JSON.parse(JSON.stringify(t.inputSchema || {}));
    const walk = (node) => {
      if (!node || typeof node !== "object") return;
      if (Array.isArray(node)) { node.forEach(walk); return; }
      if (typeof node.description === "string") {
        if (propMax <= 0) delete node.description;
        else if (node.description.length > propMax) node.description = node.description.slice(0, propMax - 1) + ".";
      }
      for (const key of Object.keys(node)) walk(node[key]);
    };
    walk(schema);
    let desc = t.description || "";
    if (toolMax > 0 && desc.length > toolMax) desc = desc.slice(0, toolMax - 1) + ".";
    return { name: t.name, description: desc, inputSchema: schema };
  });
}
function bytes(tools) { let n = 0; for (const t of tools) n += Buffer.byteLength(JSON.stringify(t), "utf8"); return n; }

console.log("baseline bytes=" + bytes(snap.tools) + " tools=" + snap.tools.length);
for (const [propMax, toolMax] of [[0, 0], [200, 200], [90, 110], [60, 100], [48, 100], [32, 90], [24, 80], [0, 80], [0, 60]]) {
  const p = project(snap.tools, propMax, toolMax);
  const b = bytes(p);
  console.log("propMax=" + propMax + " toolMax=" + toolMax + " bytes=" + b + " approxTokens=" + Math.round(b / 3.5));
}
