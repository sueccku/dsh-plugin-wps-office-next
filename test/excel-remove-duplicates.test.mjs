// 回归：wps_excel_remove_duplicates 的**语义**契约。
//
// 契约（工具自己的描述）：columns 不填 → 「根据所有列判断」。
// 旧行为（本次修复前，实测）：不填时桥里写死 `$cols = @(1)`，只按**第 1 列**去重；
// 而且行数用 `$range.Rows.Count` 算 —— 那是**按地址**算的，删完也不变，于是删了 69 行却回报「删除行数: 0」。
// 实测证据（WPS 12.1，131 行数据的真实表）：被去重成 62 行，回报 0/132。
//
// 因此本文件断言的是**真值**，不是"文案里含'删除'二字"：
//   ① 只删「整行完全相同」的行，"第 1 列相同、其余不同"的行必须留下；
//   ② 回报的三个数字必须与表格实际状态一致；
//   ③ 显式给 columns 时必须真的按那几列判重；
//   ④ 越界的列号必须明确报错，不能静默等于"没做"。
//
// 首列故意用**纯文本编码**（D-01/D-02）而不是日期：日期串在 WPS 里可能被规范化，
// 会让"整行比对"依赖格式细节；编码列没有这个歧义，断言才钉得住。
// Run: node test/excel-remove-duplicates.test.mjs
import { spawn } from "node:child_process";

const child = spawn(process.execPath, ["mcp/dist/index.js"], { stdio: ["pipe", "pipe", "pipe"], windowsHide: true });
let buf = "";
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + "\n"); }
function req(id, method, params) { return new Promise((r) => { pending.set(id, r); send({ jsonrpc: "2.0", id, method, params }); }); }
child.stdout.on("data", (d) => {
  buf += d.toString();
  let i;
  while ((i = buf.indexOf("\n")) >= 0) {
    const line = buf.slice(0, i).trim();
    buf = buf.slice(i + 1);
    if (!line) continue;
    let m; try { m = JSON.parse(line); } catch { continue; }
    if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); }
  }
});
child.stderr.on("data", () => {});

const results = [];
function check(name, ok, detail) {
  results.push({ name, ok });
  console.log((ok ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : ""));
}
function text(res) { return res && res.result && res.result.content ? String(res.result.content[0].text) : JSON.stringify((res && res.error) || {}); }
function ok(res) { return !!(res && res.result && !res.result.isError); }
let id = 100;
async function call(name, args, ms) {
  const res = await Promise.race([
    req(id++, "tools/call", { name, arguments: args }),
    new Promise((r) => setTimeout(() => r({ __timeout: true }), ms || 30000)),
  ]);
  if (res && res.__timeout) console.log("      TIMEOUT " + name);
  return res;
}
// 「原始行数: 5 / 删除行数: 1 / 剩余行数: 4」→ "5,1,4"
const triple = (s) => (String(s).match(/\d+/g) || []).map(Number).slice(0, 3).join(",");
// read_range 的「第N行: a | b | c」→ 只保留非空行
function rowsOf(resText) {
  const out = [];
  for (const line of String(resText).split(/\r?\n/)) {
    const m = /^第(\d+)行:\s*(.*)$/.exec(line.trim());
    if (!m) continue;
    const cells = m[2].split("|").map((c) => c.trim());
    if (cells.some((c) => c !== "")) out.push({ row: Number(m[1]), cells });
  }
  return out;
}
const lineOf = (r) => r.cells.join("/");

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "dedup", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });

// 关键构造：第 3 行与第 2 行**第 1 列相同、其余不同**（旧行为会把第 3 行当重复删掉），
// 第 4 行与第 2 行**完全相同**（必须删）。
const BLOCK = [
  ["编码", "门店", "品类", "金额"],
  ["D-01", "城西店", "面包", 100],
  ["D-01", "城东店", "蛋糕", 200],
  ["D-01", "城西店", "面包", 100],
  ["D-02", "城西店", "面包", 150],
];

check("create_workbook", ok(await call("wps_excel_create_workbook", {})), "");
// 记下本文件创建的那个工作簿名：收尾要关掉它。宿主是常驻的，留着会让**后面的测试文件**
// 在"多个工作簿打开"的状态下跑，未指定 sheet 的调用会落到活动窗口上 → 与本次改动无关的假红（实测踩过）。
const bookList = text(await call("wps_excel_get_open_workbooks", {}));
const bookName = ((bookList.split(/\r?\n/).find((l) => l.includes("|")) || "").split("|")[0] || "").trim();
check("拿到本次创建的工作簿名", bookName.length > 0, bookName);

// ---- 用例 1：省略 columns → 必须按所有列判重 ----------------------------------------------
await call("wps_excel_write_range", { range: "A1", data: BLOCK });
const r1 = await call("wps_excel_remove_duplicates", { range: "A1:D5", hasHeader: true });
const t1 = text(r1);
check("默认(省略 columns)调用成功", ok(r1), t1.replace(/\s+/g, " ").slice(0, 100));
check("默认按所有列判重：5 行只删掉 1 行完全重复的行", triple(t1) === "5,1,4", t1.replace(/\s+/g, " ").slice(0, 100));

const after1 = rowsOf(text(await call("wps_excel_read_range", { range: "A1:D6" })));
check("回报的剩余行数与表里实际非空行数一致", after1.length === 4, "实际非空行=" + after1.length);
check("第 1 列相同、其余不同的行被保留（旧行为会删掉它）",
  after1.length === 4 && after1[2].cells[1] === "城东店" && after1[2].cells[2] === "蛋糕",
  JSON.stringify(after1.map(lineOf)));
check("保留下来的正是「表头 + 第2行 + 第3行 + 第5行」",
  after1.length === 4
  && lineOf(after1[1]) === "D-01/城西店/面包/100"
  && lineOf(after1[2]) === "D-01/城东店/蛋糕/200"
  && lineOf(after1[3]) === "D-02/城西店/面包/150",
  JSON.stringify(after1.map(lineOf)));

// ---- 用例 2：显式给 columns=["A"] → 必须真的按第 1 列判重（删 2 行） ----------------------
await call("wps_excel_write_range", { range: "A10", data: BLOCK });
const r2 = await call("wps_excel_remove_duplicates", { range: "A10:D14", hasHeader: true, columns: ["A"] });
const t2 = text(r2);
check("显式 columns=[\"A\"] 调用成功", ok(r2), t2.replace(/\s+/g, " ").slice(0, 100));
check("显式按第 1 列判重：删掉 2 行同编码行", triple(t2) === "5,2,3", t2.replace(/\s+/g, " ").slice(0, 100));
const after2 = rowsOf(text(await call("wps_excel_read_range", { range: "A10:D15" })));
check("显式 columns 的结果与表里实际一致", after2.length === 3, "实际非空行=" + after2.length);

// ---- 用例 3：hasHeader=false → 3 行完全相同只留 1 行 --------------------------------------
await call("wps_excel_write_range", { range: "A20", data: [["x", 1], ["x", 1], ["x", 1]] });
const r3 = await call("wps_excel_remove_duplicates", { range: "A20:B22", hasHeader: false });
check("hasHeader=false：3 行完全相同 → 删 2 留 1", triple(text(r3)) === "3,2,1", text(r3).replace(/\s+/g, " ").slice(0, 100));

// ---- 用例 4：越界列号必须明确报错（不能静默"什么都没做"） --------------------------------
await call("wps_excel_write_range", { range: "A30", data: [["a", 1], ["a", 1]] });
const r4 = await call("wps_excel_remove_duplicates", { range: "A30:B31", columns: [9] });
const t4 = text(r4);
check("越界列号明确报错而不是静默通过", !ok(r4) && /out of range/.test(t4), t4.replace(/\s+/g, " ").slice(0, 120));

// ---- 收尾：只关掉本文件创建的工作簿（不保存），不给后面的测试文件留状态 --------------------
const closeRes = await call("wps_excel_close_workbook", { name: bookName, save: false });
const restList = text(await call("wps_excel_get_open_workbooks", {}));
check("收尾关闭本文件创建的工作簿", ok(closeRes) && !restList.includes(bookName),
  text(closeRes).replace(/\s+/g, " ").slice(0, 80));

console.log("\n" + results.filter((r) => r.ok).length + "/" + results.length + " passed");
child.kill();
process.exit(results.every((r) => r.ok) ? 0 : 1);
