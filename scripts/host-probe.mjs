import { spawn } from "node:child_process";

const host = process.argv[2] || "host/wps-com-host.ps1";
const child = spawn("powershell", ["-NoProfile", "-NoLogo", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass", "-File", host], { stdio: ["pipe", "pipe", "pipe"], windowsHide: true });

let buf = "";
let stderr = "";
const pending = new Map();
let seq = 0;
const ready = { at: 0 };

child.stdout.on("data", (d) => {
  buf += d.toString();
  let i;
  while ((i = buf.indexOf("\n")) >= 0) {
    const line = buf.slice(0, i).trim();
    buf = buf.slice(i + 1);
    if (!line) continue;
    let m;
    try { m = JSON.parse(line); } catch { console.log("non-json stdout: " + line.slice(0, 200)); continue; }
    if (m.ready) { ready.at = Date.now(); console.log("ready pid=" + m.pid + " in " + (ready.at - t0) + "ms"); continue; }
    if (m.id && pending.has(m.id)) { const p = pending.get(m.id); pending.delete(m.id); p(m); }
  }
});
child.stderr.on("data", (d) => { stderr += d.toString(); });

function send(action, params) {
  const id = ++seq;
  return new Promise((resolve) => {
    pending.set(id, resolve);
    const started = Date.now();
    child.stdin.write(JSON.stringify({ id, action, params: params || {} }) + "\n");
    const timer = setTimeout(() => { if (pending.has(id)) { pending.delete(id); resolve({ __timeout: true, ms: Date.now() - started }); } }, 60000);
    const wrapped = (v) => { clearTimeout(timer); v.__wall = Date.now() - started; resolve(v); };
    pending.set(id, wrapped);
  });
}

const t0 = Date.now();
const cases = [
  ["ping", {}],
  ["wireCheck", {}],
  ["ping", {}],
  ["getAppInfo", {}],
  ["ping", {}],
  ["getAppInfo", {}]
];
for (const [action, params] of cases) {
  const r = await send(action, params);
  const txt = r.result ? JSON.stringify(r.result).slice(0, 150) : JSON.stringify(r).slice(0, 150);
  console.log(action.padEnd(12) + " wall=" + String(r.__wall).padStart(5) + "ms host=" + String(r.ms).padStart(5) + "ms ok=" + r.ok + "  " + txt);
}
await send("__shutdown", {});
await new Promise((r) => setTimeout(r, 300));
if (stderr.trim()) console.log("stderr tail: " + stderr.trim().split("\n").slice(-4).join(" | ").slice(0, 400));
console.log("probe done, total=" + (Date.now() - t0) + "ms");
try { child.kill(); } catch {}
process.exit(0);
