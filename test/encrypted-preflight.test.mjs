// FIXES 69: an encrypted OOXML file is recognised from its file header BEFORE any COM call reaches
// .Open(). Opening one is the last way to freeze a session: WPS pops a modal password box that blocks
// the STA host, and Presentations.Open has no password parameter at all. The guard is pure byte
// inspection, so this test needs no WPS and can run in CI.
// Run: node test/encrypted-preflight.test.mjs
import { spawn } from "node:child_process";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";

const ARTIFACTS = "test/.artifacts/encrypted-preflight";
mkdirSync(ARTIFACTS, { recursive: true });

// The OLE/CFB magic: what every encrypted OOXML file starts with (a plain one starts with "PK").
const CFB = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
function encryptedLike(name) {
  const file = join(ARTIFACTS, name);
  const bytes = Buffer.alloc(1024, 0x11);
  Buffer.from(CFB).copy(bytes, 0);
  writeFileSync(file, bytes);
  return file;
}

const cases = [
  { app: "word", tool: "wps_word_open_document", file: encryptedLike("locked.docx") },
  { app: "excel", tool: "wps_excel_open_workbook", file: encryptedLike("locked.xlsx") },
  { app: "ppt", tool: "wps_ppt_open_presentation", file: encryptedLike("locked.pptx") },
];

const child = spawn(process.execPath, ["mcp/dist/index.js"], { stdio: ["pipe", "pipe", "pipe"], windowsHide: true });
let buf = "";
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + "\n"); }
function req(id, method, params) { return new Promise((resolve) => { pending.set(id, resolve); send({ jsonrpc: "2.0", id, method, params }); }); }
child.stdout.on("data", (d) => { buf += d.toString(); let i; while ((i = buf.indexOf("\n")) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on("data", () => {});

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }
function text(res) { return res && res.result && res.result.content ? String(res.result.content[0].text) : JSON.stringify((res && res.error) || {}); }

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "encrypted-preflight", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 10;
function call(name, args, ms) {
  return new Promise((resolve) => {
    const timer = setTimeout(() => resolve({ __timeout: true }), ms);
    req(id++, "tools/call", { name, arguments: args }).then((r) => { clearTimeout(timer); resolve(r); });
  });
}

for (const c of cases) {
  const started = Date.now();
  const res = await call(c.tool, { filePath: c.file }, 30000);
  const ms = Date.now() - started;
  const body = text(res);
  check(c.app + ": an encrypted file returns instead of hanging", !res.__timeout, ms + "ms");
  check(c.app + ": it fails", !!(res.result && res.result.isError), body.replace(/\s+/g, " ").slice(0, 120));
  check(c.app + ": the message names the encryption and the next step", /加密/.test(body) && /密码/.test(body) && /另存为/.test(body), body.replace(/\s+/g, " ").slice(0, 170));
  check(c.app + ": it fails immediately (no COM, no modal)", ms < 10000, ms + "ms");
}

// The guard must not be widened to legacy binary formats: .doc/.xls/.ppt are CFB too and open fine.
const bridge = readFileSync("mcp/scripts/wps-com.ps1", "utf8");
const list = /WpsOoxmlExtensions = @\(([^)]*)\)/.exec(bridge);
check("the guard declares an extension list", !!list, list ? list[1] : "not found");
if (list) {
  const exts = [...list[1].matchAll(/'([^']+)'/g)].map((m) => m[1]);
  check("the guard covers the OOXML extensions", [".docx", ".docm", ".xlsx", ".xlsm", ".pptx", ".pptm"].every((e) => exts.includes(e)), exts.join(" "));
  check("the guard leaves legacy binary formats alone", ![".doc", ".xls", ".ppt", ".dps", ".et", ".wps"].some((e) => exts.includes(e)), exts.join(" "));
}

try { child.kill(); } catch { }
const failed = results.filter((r) => !r.ok);
console.log("");
console.log(failed.length ? "ENCRYPTED PREFLIGHT TESTS FAILED" : "ENCRYPTED PREFLIGHT TESTS OK (" + results.length + ")");
process.exit(failed.length ? 1 : 0);
