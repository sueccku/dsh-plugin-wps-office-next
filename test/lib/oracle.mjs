// 独立 oracle：**不经过插件的工具层**，直接读回真值。
//
// 为什么需要（W1-3 / W1-4）：本项目 49 个测试文件里，绝大多数只用"插件自己的回报"验证自己 ——
// 而读写同源会一起错。P5 就是在全绿套件下漏过去的：断言查的是 styles[0]（第 1 段），
// 缺陷打的是紧邻的上一段。覆盖率早就 268/268 了，缺的是**证明力**。
//
// 三条独立路径，按强度排序（能用上一条就别退到下一条）：
//   1) 外部库读**落盘后的文件**：openpyxl / python-docx / python-pptx —— 能看出"到底写没写进文件"
//   2) 裸 COM 直接问 WPS：不经过 mcp/dist —— 能看出"活文档的真实状态"
//   3) 文件头：PNG / PDF 的魔数与尺寸
//
// 任何一条都**禁止**用被验证工具自己的读接口（那是自证，不是独立验证）。
//
// ⚠️ 两个已经踩过的坑（否则会得到"永远通过"的假断言）：
//   1. **ProgID 分应用**：WPS 里 Word=Kwps.Application、表格=Ket.Application、演示=Kwpp.Application。
//      拿 Word 的 ProgID 去连表格会得到 MK_E_UNAVAILABLE（0x800401E3）。
//   2. **失败必须返回 null，不能返回空数组** —— 否则 before=[] / after=[] 会静默通过。
import { spawnSync } from "node:child_process";
import { existsSync, readFileSync, statSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";

/** WPS 各应用在 ROT 里的 ProgID（与桥里的表一致）。 */
export const PROGID = { excel: "Ket.Application", word: "Kwps.Application", ppt: "Kwpp.Application" };

/** 裸 COM：脚本的返回值经 base64 回传。
 *  为什么绕一圈：中文经 PowerShell → 管道的编码链会被改坏（FIXES 87 时代实测过）。 */
export function com(script) {
  const b64 = Buffer.from(script, "utf8").toString("base64");
  const r = spawnSync("powershell", ["-NoProfile", "-STA", "-Command",
    `$s = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${b64}')); $out = Invoke-Expression $s; [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([string]$out))`],
    { encoding: "utf8", windowsHide: true });
  try { return Buffer.from((r.stdout || "").trim(), "base64").toString("utf8"); } catch { return ""; }
}

function sleep(ms) { const end = Date.now() + ms; while (Date.now() < end) { /* 同步等待，测试里最多几百毫秒 */ } }

/** 带重试的 com：WPS 刚启动时实例可能还没进 ROT（MK_E_UNAVAILABLE）。
 *  失败时返回 null（**不是空串**），调用方据此把断言判失败 —— 绝不能让"读不到"被当成"是空的"。 */
export function comStrict(script) {
  let last = "";
  for (let i = 0; i < 4; i++) {
    const out = com(script);
    if (out) return out;
    last = out;
    sleep(250);
  }
  return last === "" && /MK_E_UNAVAILABLE|0x800401E3/.test(com(script)) ? null : null;
}

/** 裸 COM 读活动工作簿的一块区域，返回 string[][]；**读失败返回 null**。 */
export function comGrid(range, sheetName) {
  const sheetExpr = sheetName ? `$w.Worksheets.Item(${JSON.stringify(sheetName)})` : "$w.ActiveSheet";
  let out = "";
  for (let i = 0; i < 4 && !out; i++) {
    out = com(`
$ErrorActionPreference = 'Stop'
try {
  $w = [Runtime.InteropServices.Marshal]::GetActiveObject('${PROGID.excel}')
  $s = ${sheetExpr}
  $rows = @()
  foreach ($r in $s.Range('${range}').Rows) {
    $cells = @()
    foreach ($c in $r.Cells) { $t = [string]$c.Text; $t = $t.Replace('\\','\\\\').Replace('"','\\"'); $cells += ('"' + $t + '"') }
    $rows += ('[' + ($cells -join ',') + ']')
  }
  ('[' + ($rows -join ',') + ']')
} catch { '' }
`);
    if (!out) sleep(250);
  }
  if (!out) return null;
  try {
    const v = JSON.parse(out);
    if (!Array.isArray(v)) return null;
    return Array.isArray(v[0]) ? v : [v];
  } catch { return null; }
}

/** 裸 COM 读工作表名单；**读失败返回 null**。 */
export function comSheetNames() {
  let out = "";
  for (let i = 0; i < 4 && !out; i++) {
    out = com(`$ErrorActionPreference='Stop'; try { $w=[Runtime.InteropServices.Marshal]::GetActiveObject('${PROGID.excel}'); ($w.Worksheets | ForEach-Object { $_.Name }) -join ',' } catch { '' }`);
    if (!out) sleep(250);
  }
  return out ? out.trim().split(",") : null;
}

/** 裸 COM 读活动 Word 文档每一段的文本；**读失败返回 null**。 */
export function comParagraphs() {
  let out = "";
  for (let i = 0; i < 4 && !out; i++) {
    out = com(`$ErrorActionPreference='Stop'; try { $w=[Runtime.InteropServices.Marshal]::GetActiveObject('${PROGID.word}'); $d=$w.ActiveDocument; $t=@(); for ($i=1; $i -le $d.Paragraphs.Count; $i++) { $x=([string]$d.Paragraphs.Item($i).Range.Text).TrimEnd([char]13,[char]10); $x=$x.Replace('\\','\\\\').Replace('"','\\"'); $t += ('"' + $x + '"') }; ('[' + ($t -join ',') + ']') } catch { '' }`);
    if (!out) sleep(250);
  }
  if (!out) return null;
  try { const v = JSON.parse(out); return Array.isArray(v) ? v : [v]; } catch { return null; }
}

// ---- 外部库（tier 1）：需要 python + openpyxl / python-docx / python-pptx ----
let PY_CACHE;
/** 找 python：WPS_TEST_PYTHON 覆盖 > DSH 自带运行时 > PATH。找不到返回 null（调用方据此降级）。 */
export function findPython() {
  if (PY_CACHE !== undefined) return PY_CACHE;
  const candidates = [
    process.env.WPS_TEST_PYTHON,
    join(homedir(), ".dsh", "dsh-runtimes", "dsh-primary-runtime", "dependencies", "python", "python.exe"),
    "python", "py",
  ].filter(Boolean);
  for (const c of candidates) {
    const r = spawnSync(c, ["-c", "import sys"], { encoding: "utf8", windowsHide: true });
    if (!r.error && r.status === 0) { PY_CACHE = c; return c; }
  }
  PY_CACHE = null;
  return null;
}

/** 跑一段 python：code 里给 result 赋值（可 JSON 序列化），返回解析后的对象；失败返回 null。 */
export function py(code) {
  const p = findPython();
  if (!p) return null;
  const script = `import json\n${code}\nprint(json.dumps(result, default=str))`;
  const r = spawnSync(p, ["-c", script], { encoding: "utf8", windowsHide: true });
  const out = (r.stdout || "").trim().split("\n").pop();
  try { return JSON.parse(out); } catch { return null; }
}

/** 用 openpyxl 读**落盘后**的 xlsx：值 / 数字格式 / 粗体 / 条件格式 / 表名，全部来自文件本身；失败返回 null。 */
export function xlsx(file, range, sheetName) {
  return py(`
import openpyxl
wb = openpyxl.load_workbook(${JSON.stringify(file)}, data_only=False)
ws = wb[${sheetName ? JSON.stringify(sheetName) : "wb.sheetnames[0]"}]
cells = list(ws[${JSON.stringify(range)}])
result = {
  "sheetnames": wb.sheetnames,
  "values": [[("" if c.value is None else c.value) for c in row] for row in cells],
  "formats": [[c.number_format for c in row] for row in cells],
  "bold": [[bool(c.font.bold) for c in row] for row in cells],
  "conditional": [str(r) for r in ws.conditional_formatting],
}
`);
}

/** 用 openpyxl 读**落盘后**的 xlsx 的页面/打印/外观属性（按表名返回）。
 *  为什么需要：打印设置类是"回报写了 ≠ 文件里真是那样"的重灾区 —— 只信插件自己的回读会漏。
 *  注意 openpyxl 的单位：页边距是**英寸**（36 磅 = 0.5 英寸），paperSize 是**数字**（A3 = 8）。 */
export function xlsxProps(file) {
  return py(`
import openpyxl
wb = openpyxl.load_workbook(${JSON.stringify(file)})

def props(name):
    ws = wb[name]
    setupPr = ws.sheet_properties.pageSetUpPr
    return {
        "orientation": ws.page_setup.orientation,
        "paperSize": ws.page_setup.paperSize,
        "scale": ws.page_setup.scale,
        "fitToWidth": ws.page_setup.fitToWidth,
        "fitToPage": (bool(setupPr.fitToPage) if setupPr is not None else None),
        "marginTop": ws.page_margins.top,
        "marginLeft": ws.page_margins.left,
        "horizontalCentered": bool(ws.print_options.horizontalCentered),
        "printTitleRows": ws.print_title_rows,
        "headerLeft": (ws.oddHeader.left.text if ws.oddHeader.left is not None else None),
        "footerCenter": (ws.oddFooter.center.text if ws.oddFooter.center is not None else None),
        "tabColor": (ws.sheet_properties.tabColor.rgb if ws.sheet_properties.tabColor is not None else None),
        "sheetState": ws.sheet_state,
    }

result = {name: props(name) for name in wb.sheetnames}
`);
}

/** 读**落盘后**的 xlsx 里定义的表（ListObject）：名字/范围/列名/样式。失败返回 null。
 *  直接解析 \`xl/tables/*.xml\`（zip 里的 part）而不是走 openpyxl 的表 API —— 实测 WPS 写的表
 *  openpyxl 能读懂，但遍历 \`ws.tables\` 这条路径在本项目里拿不到东西；直接读 part 更稳，
 *  而且同样是"以文件为准"，不经过插件的任何接口。 */
export function xlsxTables(file) {
  // 用 ElementTree 而不是正则：多层模板字符串会把正则里的反斜杠吃掉（我第一版就栽在这）。
  return py(`
import zipfile
import xml.etree.ElementTree as ET
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
z = zipfile.ZipFile(${JSON.stringify(file)})
out = {}
for part in z.namelist():
    if not (part.startswith("xl/tables/") and part.endswith(".xml")):
        continue
    root = ET.fromstring(z.read(part))
    key = root.get("name") or part
    cols = [c.get("name") for c in root.findall(NS + "tableColumns/" + NS + "tableColumn")]
    si = root.find(NS + "tableStyleInfo")
    out[key] = {"ref": root.get("ref"), "columns": cols, "style": (si.get("name") if si is not None else None), "part": part}
result = out
`);
}


/** 用 python-docx 读**落盘后**的 docx 段落文本；失败返回 null。 */
export function docxText(file) {
  return py(`
import docx
result = [p.text for p in docx.Document(${JSON.stringify(file)}).paragraphs]
`);
}

/** 用 python-docx 读**落盘后**的 docx 里的表格：每张表的行数/列数/单元格文本。失败返回 null。 */
export function docxTables(file) {
  return py(`
import docx
d = docx.Document(${JSON.stringify(file)})
out = []
for t in d.tables:
    rows = []
    for r in t.rows:
        rows.append([c.text for c in r.cells])
    out.append({"rows": len(t.rows), "cols": len(t.columns), "cells": rows})
result = out
`);
}

/** 读**落盘后**的 docx 的结构性事实（直接读 XML part，不走 python-docx 的高层 API）：
 *  分栏标记 / 修订标记数 / 页眉页脚里有没有 PAGE 域 / 批注条数。
 *  这些恰恰是"回报写了 ≠ 文件里真是那样"最典型的地方。失败返回 null。 */
export function docxFacts(file) {
  // 不用正则：多层模板字符串会把反斜杠吃掉（W1-4 踩过）。
  return py(`
import zipfile
z = zipfile.ZipFile(${JSON.stringify(file)})
names = z.namelist()
doc = z.read("word/document.xml").decode("utf-8", "ignore")
footers = [n for n in names if n.startswith("word/footer") and n.endswith(".xml")]
headers = [n for n in names if n.startswith("word/header") and n.endswith(".xml")]

def has_page(parts):
    for p in parts:
        if "PAGE" in z.read(p).decode("utf-8", "ignore"):
            return True
    return False

cols = ["<w:cols" + c.split(">")[0] + ">" for c in doc.split("<w:cols")[1:]]
comments = 0
if "word/comments.xml" in names:
    comments = z.read("word/comments.xml").decode("utf-8", "ignore").count("<w:comment ")
result = {
    "colsTags": cols,
    "insertions": doc.count("<w:ins "),
    "deletions": doc.count("<w:del "),
    "footerHasPageField": has_page(footers),
    "headerHasPageField": has_page(headers),
    "comments": comments,
}
`);
}

/** 用 python-docx 数**落盘后**的 docx 里的内嵌图片数量（"插入图片"到底有没有进文件）。失败返回 null。 */
export function docxInlineShapes(file) {
  return py("import docx\nresult = len(docx.Document(" + JSON.stringify(file) + ").inline_shapes)");
}

/** 用 python-pptx 读**落盘后**的 pptx 每页文本；失败返回 null。 */
export function pptxText(file) {
  return py(`
import pptx
out = []
for s in pptx.Presentation(${JSON.stringify(file)}).slides:
    out.append([sh.text_frame.text for sh in s.shapes if sh.has_text_frame])
result = out
`);
}

/** 读**落盘后**的 xlsx 里那些"只断言调用没报错"根本验证不到的特性：
 *  数据验证规则 / 超链接 / 边框 / 条件格式规则。失败返回 null。 */
export function xlsxFeatures(file, sheetName, borderCells) {
  const code = [
    "import openpyxl",
    "wb = openpyxl.load_workbook(" + JSON.stringify(file) + ")",
    sheetName ? "ws = wb[" + JSON.stringify(sheetName) + "]" : "ws = wb[wb.sheetnames[0]]",
    "dv = []",
    "for d in ws.data_validations.dataValidation:",
    '    dv.append({"type": d.type, "formula1": d.formula1, "sqref": str(d.sqref)})',
    "hl = []",
    "for row in ws.iter_rows():",
    "    for c in row:",
    "        if c.hyperlink is not None:",
    '            hl.append({"cell": c.coordinate, "target": c.hyperlink.target})',
    "borders = {}",
    "for coord in " + JSON.stringify(borderCells || ["A1"]) + ":",
    "    b = ws[coord].border",
    '    borders[coord] = {"top": b.top.style, "bottom": b.bottom.style, "left": b.left.style, "right": b.right.style}',
    'result = {"dv": dv, "hyperlinks": hl, "borders": borders, "conditional": [str(r) for r in ws.conditional_formatting]}',
  ].join("\n");
  return py(code);
}

/** 在**落盘后**的 OOXML（xlsx/docx/pptx）里数某个记号出现的次数与所在 part。
 *  例：迷你图数 `<x14:sparklineGroup `、图表标题文本、修订标记。失败返回 null。
 *  用字符串 count 而不是正则：多层模板字符串会把反斜杠吃掉（W1-4 踩过）。
 *  这里刻意用**普通字符串拼接**而不是模板字面量 —— 少一层转义就少一类坑。 */
export function zipScan(file, needle, prefix) {
  const code = [
    'import zipfile',
    "needle = " + JSON.stringify(needle),
    "prefix = " + JSON.stringify(prefix || ""),
    "z = zipfile.ZipFile(" + JSON.stringify(file) + ")",
    "count = 0",
    "parts = []",
    "for n in z.namelist():",
    '    if not n.endswith(".xml"):',
    "        continue",
    "    if prefix and not n.startswith(prefix):",
    "        continue",
    '    c = z.read(n).decode("utf-8", "ignore").count(needle)',
    "    if c:",
    "        count += c",
    "        parts.append(n)",
    'result = {"count": count, "parts": parts}',
  ].join("\n");
  return py(code);
}
/** 文件头：PNG 的字节数与真实像素尺寸（导出类工具最容易被"写了个空图"骗过去）。 */
export function pngInfo(file) {
  if (!existsSync(file)) return { exists: false };
  const b = readFileSync(file);
  if (b.length < 24 || b[0] !== 0x89 || b[1] !== 0x50) return { exists: true, bytes: b.length, png: false };
  return { exists: true, bytes: statSync(file).size, png: true, width: b.readUInt32BE(16), height: b.readUInt32BE(20) };
}
