import { spawn } from "node:child_process";
import { writeFileSync, mkdirSync } from "node:fs";
import { dirname } from "node:path";
const [, , entry, outPath] = process.argv;
const child = spawn(process.execPath, [entry], { stdio: ["pipe", "pipe", "pipe"], windowsHide: true });
let buf = "";
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + "\n"); }
function req(id, method, params) { return new Promise((resolve) => { pending.set(id, resolve); send({ jsonrpc: "2.0", id, method, params }); }); }
child.stdout.on("data", (d) => { buf += d.toString(); let i; while ((i = buf.indexOf("\n")) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on("data", () => {});
await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "snapshot", version: "1.0.0" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
const res = await req(2, "tools/list", {});
const tools = (res.result && res.result.tools) || [];
let bytes = 0;
for (const t of tools) bytes += Buffer.byteLength(JSON.stringify(t), "utf8");
mkdirSync(dirname(outPath), { recursive: true });
writeFileSync(outPath, JSON.stringify({ capturedAt: new Date().toISOString(), entry, toolCount: tools.length, schemaBytes: bytes, approxTokens: Math.round(bytes / 3.5), tools }, null, 2) + "\n");
console.log("tools=" + tools.length + " schemaBytes=" + bytes + " approxTokens=" + Math.round(bytes / 3.5) + " -> " + outPath);
child.kill();
process.exit(0);
