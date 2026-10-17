# 生成固定评估场景的输入文件（docs/skill-evaluation.md §1）。
#
# 用法： python fixtures.py <fixture-key> <outdir>
#   fixture-key 必须在 test/skilleval/scenarios.mjs 的 FIXTURE_KEYS 里（run.mjs 会对账）。
# 只写输入文件，不碰任何真实文档；输出目录由调用方保证是干净的空目录。
import os
import sys


def two_word_target(out):
    from docx import Document

    a = Document()
    a.add_paragraph("A-原始内容")
    a.save(os.path.join(out, "报告A.docx"))
    b = Document()
    b.add_paragraph("B-原始内容")
    b.save(os.path.join(out, "报告B.docx"))


def single_word(out):
    from docx import Document

    doc = Document()
    doc.add_paragraph("A-原始内容")
    doc.save(os.path.join(out, "报告A.docx"))


def excel_row_formulas(out):
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["行号", "单价", "数量", "金额"])
    for i in range(2, 7):
        ws.append([i - 1, 100 * (i - 1), i - 1, None])   # D 列留空，由被测 Skill 填公式
    wb.save(os.path.join(out, "输入.xlsx"))


def small_change(out):
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = 1
    ws["A2"] = "不要改我"
    wb.save(os.path.join(out, "输入.xlsx"))


FIXTURES = {
    "two-word-target": two_word_target,
    "single-word": single_word,
    "excel-row-formulas": excel_row_formulas,
    "small-change": small_change,
}


def main(argv):
    if len(argv) != 3:
        print("usage: fixtures.py <fixture-key> <outdir>", file=sys.stderr)
        return 2
    key, out = argv[1], argv[2]
    build = FIXTURES.get(key)
    if build is None:
        print("unknown fixture key: " + key, file=sys.stderr)
        return 2
    os.makedirs(out, exist_ok=True)
    build(out)
    for name in sorted(os.listdir(out)):
        print(name)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
