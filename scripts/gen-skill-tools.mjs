// 从**实时 MCP 工具目录**生成 Skill 的按需参考文件与 SKILL.md 的短入口块。
//
// Run:
//   node scripts/gen-skill-tools.mjs            # 写入 skills/**
//   node scripts/gen-skill-tools.mjs --check    # 只在内存里生成并与磁盘比对，不写文件（CI 用）
//
// 产出：
//   1. skills/<skill>/reference.md —— 完整工具目录 + 按任务分组索引（整份由本脚本覆盖）。
//   2. 各 SKILL.md 里 `GENERATED:advertised` 标记块 —— **只写计数与参考入口，不再注入工具名清单**。
//
// 为什么改成这样（改进指南 §8.1）：以前这里把几万个字符的工具名重新打进 SKILL.md，
// 于是"读一次技能"就等于读一份工具目录，而这份目录还必须在生成器里再维护一遍。
// 现在 SKILL.md 只负责流程与入口，目录留在 reference.md，需要时再读。
//
// 只做 initialize + tools/list：生成任务**不发起任何 COM 业务操作**（不许因为跑生成器而启动 WPS）。
// 请求有超时、子进程退出会被当作失败、stdout 用 StringDecoder 收帧（中文描述不会因多字节被切断）。
import { spawn } from "node:child_process";
import { writeFileSync, readFileSync, mkdirSync } from "node:fs";
import { createRequire } from "node:module";
import { StringDecoder } from "node:string_decoder";

const argv = process.argv.slice(2);
const CHECK = argv.includes("--check");
const entry = argv.find((arg) => !arg.startsWith("--")) || "mcp/dist/index.js";

/** 整个生成过程的硬超时：卡住的 server 不该让 CI 挂在那里。 */
const GEN_TIMEOUT_MS = Number(process.env.WPS_OFFICE_GEN_TIMEOUT_MS || 60000);
/** 单次请求（initialize / tools/list）的超时。 */
const REQ_TIMEOUT_MS = Number(process.env.WPS_OFFICE_GEN_REQ_TIMEOUT_MS || 30000);

const child = spawn(process.execPath, [entry], {
  stdio: ["pipe", "pipe", "pipe"],
  windowsHide: true,
  env: { ...process.env, WPS_OFFICE_TOOLSET: "full" },
});

let buf = "";
const decoder = new StringDecoder("utf8");
const pending = new Map();
let childError = null;

function failAll(error) {
  childError = childError || error;
  for (const [, entry] of pending) {
    clearTimeout(entry.timer);
    entry.reject(error);
  }
  pending.clear();
}

function send(o) {
  if (!child.stdin || child.stdin.destroyed) throw new Error("生成器：MCP server 的标准输入已关闭");
  child.stdin.write(JSON.stringify(o) + "\n");
}

function req(id, method, params) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      pending.delete(id);
      reject(new Error("生成器：" + method + " 超过 " + REQ_TIMEOUT_MS + "ms 未返回，已放弃"));
    }, REQ_TIMEOUT_MS);
    pending.set(id, { resolve, reject, timer });
    send({ jsonrpc: "2.0", id, method, params });
  });
}

child.stdout.on("data", (chunk) => {
  buf += decoder.write(chunk);
  let index;
  while ((index = buf.indexOf("\n")) >= 0) {
    const line = buf.slice(0, index).trim();
    buf = buf.slice(index + 1);
    if (!line) continue;
    let message;
    try { message = JSON.parse(line); } catch { continue; }
    if (message.id && pending.has(message.id)) {
      const entry = pending.get(message.id);
      pending.delete(message.id);
      clearTimeout(entry.timer);
      entry.resolve(message);
    }
  }
});
child.stderr.on("data", () => { /* 生成器不消费 server 日志 */ });
child.on("error", (error) => failAll(new Error("生成器：无法启动 MCP server：" + error.message)));
child.on("exit", (code, signal) => {
  if (pending.size) failAll(new Error("生成器：MCP server 提前退出（code=" + String(code) + ", signal=" + String(signal) + "）"));
});

const overall = setTimeout(() => {
  failAll(new Error("生成器：整体超过 " + GEN_TIMEOUT_MS + "ms 未完成"));
  try { child.kill(); } catch { /* already gone */ }
  process.exitCode = 1;
}, GEN_TIMEOUT_MS);
overall.unref?.();

/** 收尾：无论成败都要把子进程收干净，避免留下孤儿进程。 */
function shutdown() {
  clearTimeout(overall);
  try { child.kill(); } catch { /* already gone */ }
}

let tools;
try {
  await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "gen-skills", version: "1.0.0" } });
  send({ jsonrpc: "2.0", method: "notifications/initialized" });
  const res = await req(2, "tools/list", {});
  tools = (res.result && res.result.tools) || [];
  if (!tools.length) throw new Error("生成器：tools/list 返回了 0 个工具（mcp/dist 是否已构建？）");
} catch (error) {
  console.error(String(error && error.message ? error.message : error));
  shutdown();
  process.exit(1);
}
shutdown();

// 广告面来自 server 自己用的那个模块，所以不会与真实 tools/list 漂移。
const require = createRequire(import.meta.url);
const toolset = require("../mcp/dist/server/toolset.js");
const standard = new Set([...(toolset.FACADE_TOOLS || []), ...(toolset.STANDARD_TOOLS || [])]);

function firstSentence(text, max) {
  const flat = (text || "").replace(/\s+/g, " ").trim();
  const stop = flat.search(/[。.!?]/);
  const cut = stop >= 0 && stop + 1 <= max ? flat.slice(0, stop + 1) : flat.slice(0, max);
  return cut.replace(/\|/g, "/");
}

function bucket(name) {
  const m = /^wps_(excel|word|ppt)_/.exec(name);
  if (m) return m[1];
  if (/^wps_(common|convert)_/.test(name)) return "common";
  if (["wps_status", "wps_help", "wps_call", "wps_batch"].includes(name)) return "facade";
  return "builtin";
}

const groups = { facade: [], common: [], excel: [], word: [], ppt: [], builtin: [] };
for (const t of tools) groups[bucket(t.name)].push(t);

const catalog = new Map(tools.map((t) => [t.name, t]));

// ---- 分类元数据：只写工具名，参数 schema 不在这里重复维护 -------------------------------
// 校验是硬的：未知名字、重复归类、以及某个注册工具没有被归入任何组，都会让生成失败。
// 这样新增工具时必须做一次有意的归类，索引不会静默漏项。
const groupSpec = JSON.parse(readFileSync("spec/skill-tool-groups.json", "utf8"));
const groupErrors = [];
const grouped = new Map();
const plan = {};
for (const [app, sections] of Object.entries(groupSpec)) {
  if (app.startsWith("_")) continue;
  plan[app] = [];
  for (const section of sections) {
    const names = [];
    for (const name of section.tools || []) {
      if (!catalog.has(name)) { groupErrors.push(app + " / " + section.title + "：目录里没有 " + name); continue; }
      if (grouped.has(name)) { groupErrors.push(name + " 被归入两组：" + grouped.get(name) + " 与 " + section.title); continue; }
      grouped.set(name, section.title);
      names.push(name);
    }
    plan[app].push({ title: section.title, names });
  }
  const registered = new Set(tools.filter((t) => app === "office-next"
    ? bucket(t.name) !== "excel" && bucket(t.name) !== "word" && bucket(t.name) !== "ppt"
    : bucket(t.name) === app).map((t) => t.name));
  const missing = [...registered].filter((n) => !grouped.has(n));
  if (missing.length) groupErrors.push(app + "：这些工具没有被归入任何任务组 -> " + missing.join(", "));
  const extra = plan[app].flatMap((s) => s.names).filter((n) => !registered.has(n));
  if (extra.length) groupErrors.push(app + "：这些名字不属于本应用 -> " + extra.join(", "));
}
if (groupErrors.length) {
  console.error("spec/skill-tool-groups.json 与实时目录不一致：");
  for (const error of groupErrors) console.error("  x " + error);
  process.exit(1);
}

const titles = {
  facade: "门面工具（始终广告）",
  common: "通用与转换工具",
  excel: "WPS 表格工具",
  word: "WPS 文字工具",
  ppt: "WPS 演示工具",
  builtin: "内置兼容工具",
};

function fullTable(names, heading) {
  const rows = names.map((name) => {
    const t = catalog.get(name);
    const mark = standard.has(name) ? "直达" : "wps_call";
    return "| " + name + " | " + mark + " | " + firstSentence(t.description, 90) + " |";
  });
  return ["## " + heading + "（" + names.length + "）", "", "| 工具 | 在 standard 档 | 说明 |", "|---|---|---|", ...rows, ""].join("\n");
}

function taskIndex(app) {
  const lines = [];
  for (const section of plan[app]) {
    if (!section.names.length) continue;
    lines.push("### " + section.title + "（" + section.names.length + "）");
    lines.push("");
    for (const name of section.names) lines.push("- `" + name + "`");
    lines.push("");
  }
  return lines.join("\n");
}

const header = [
  "<!-- GENERATED by scripts/gen-skill-tools.mjs - do not edit by hand. -->",
  "",
  "本文件由实时工具目录生成：完整工具表 + 按任务分组索引，供选工具和 `wps_call` 派发时查名。",
  "「直达」表示 standard 档直接广告，模型可以直接调用；「wps_call」表示未广告，",
  "必须先 `wps_help {tool:\"...\"}` 取参数 schema，再 `wps_call {tool, args}` 执行。",
  "参数类型、单位、默认值与枚举以运行时 schema（`wps_help`）为准，本文件不复制它们。",
  "",
].join("\n");

function referenceFor(app) {
  if (app === "office-next") {
    const names = [...groups.facade, ...groups.common, ...groups.builtin].map((t) => t.name);
    const advertised = names.filter((n) => standard.has(n));
    return [
      header,
      "## 按任务查找（" + plan["office-next"].filter((s) => s.names.length).length + " 组）",
      "",
      taskIndex("office-next"),
      fullTable(names, "门面、通用内置与转换工具（" + names.length + "，其中直达 " + advertised.length + "）"),
    ].join("\n");
  }
  const names = groups[app].map((t) => t.name);
  const advertised = names.filter((n) => standard.has(n));
  return [
    header,
    "## 按任务查找（" + plan[app].filter((s) => s.names.length).length + " 组）",
    "",
    taskIndex(app),
    fullTable(names, titles[app] + "（" + names.length + "，其中直达 " + advertised.length + "）"),
  ].join("\n");
}

const out = {
  "skills/wps-office-next/reference.md": referenceFor("office-next"),
  "skills/wps-excel/reference.md": referenceFor("excel"),
  "skills/wps-word/reference.md": referenceFor("word"),
  "skills/wps-ppt/reference.md": referenceFor("ppt"),
};

// ---- SKILL.md 的生成块：计数 + 参考入口，**不含工具名清单** ------------------------------
const MARK_START = "<!-- GENERATED:advertised:start -->";
const MARK_END = "<!-- GENERATED:advertised:end -->";

const countsByApp = {
  facade: groups.facade.filter((t) => standard.has(t.name)).length,
  common: groups.common.filter((t) => standard.has(t.name)).length,
  excel: groups.excel.filter((t) => standard.has(t.name)).length,
  word: groups.word.filter((t) => standard.has(t.name)).length,
  ppt: groups.ppt.filter((t) => standard.has(t.name)).length,
};
const totalAdvertised = Object.values(countsByApp).reduce((a, b) => a + b, 0);
const registeredTotal = tools.length;

const blocks = {
  "skills/wps-office-next/SKILL.md":
    "standard 档默认直接广告 **" + totalAdvertised + "** 个工具（门面 " + countsByApp.facade +
    " / 通用 " + countsByApp.common + " / 表格 " + countsByApp.excel + " / 文字 " + countsByApp.word +
    " / 演示 " + countsByApp.ppt + "），注册目录共 **" + registeredTotal + "** 个；\n" +
    "未广告的工具**只能用 `wps_call` 调用**（先用 `wps_help` 取 schema）。\n" +
    "完整工具表与按任务分组索引见同目录 `reference.md`。",
  "skills/wps-excel/SKILL.md":
    "本应用注册 **" + groups.excel.length + "** 个工具，standard 档直接广告 **" + countsByApp.excel +
    "** 个；\n完整工具表与按任务分组索引见同目录 `reference.md`（未广告的用 `wps_call`）。",
  "skills/wps-word/SKILL.md":
    "本应用注册 **" + groups.word.length + "** 个工具，standard 档直接广告 **" + countsByApp.word +
    "** 个；\n完整工具表与按任务分组索引见同目录 `reference.md`（未广告的用 `wps_call`）。",
  "skills/wps-ppt/SKILL.md":
    "本应用注册 **" + groups.ppt.length + "** 个工具，standard 档直接广告 **" + countsByApp.ppt +
    "** 个；\n完整工具表与按任务分组索引见同目录 `reference.md`（未广告的用 `wps_call`）。",
};

const SKILL_MARKERS = {
  "skills/wps-office-next/SKILL.md": {},
  "skills/wps-excel/SKILL.md": {},
  "skills/wps-word/SKILL.md": {},
  "skills/wps-ppt/SKILL.md": {},
};

function withBlock(path) {
  const text = readFileSync(path, "utf8");
  const start = text.indexOf(MARK_START);
  const end = text.indexOf(MARK_END);
  if (start < 0 || end < start) throw new Error(path + " 缺少 GENERATED:advertised 标记");
  return text.slice(0, start) + [MARK_START, blocks[path], MARK_END].join("\n") + text.slice(end + MARK_END.length);
}

const writes = [];
for (const [path, content] of Object.entries(out)) writes.push({ path, content });
for (const path of Object.keys(SKILL_MARKERS)) writes.push({ path, content: withBlock(path) });

// ---- --check：只在内存里比，不落盘；有差异就非零退出 ---------------------------------------
const onDisk = (path) => {
  try { return readFileSync(path, "utf8"); } catch { return null; }
};

if (CHECK) {
  const drift = writes.filter((w) => onDisk(w.path) !== w.content);
  for (const w of writes) console.log((drift.includes(w) ? "DRIFT " : "ok    ") + w.path);
  console.log("check 模式：未写任何文件。tools=" + tools.length + " advertised=" + totalAdvertised);
  if (drift.length) {
    console.error("生成物与磁盘不一致（" + drift.length + " 个文件）：" + drift.map((w) => w.path).join(", "));
    console.error("跑不带 --check 的 scripts/gen-skill-tools.mjs 刷新。");
    process.exit(1);
  }
  process.exit(0);
}

for (const w of writes) {
  mkdirSync(w.path.split("/").slice(0, -1).join("/"), { recursive: true });
  writeFileSync(w.path, w.content);
  console.log("wrote " + w.path + " (" + w.content.split("\n").length + " lines)");
}
console.log("total tools=" + tools.length + " advertised=" + totalAdvertised);
