# 只读检查一个产物文件的内容（verify.mjs 调它，避免在 Node 里手写 OOXML 解析）。
#
# ⚠️ 这个文件**不能叫 inspect.py**：脚本所在目录会被放进 sys.path，那样 openpyxl / lxml 里的
# `from inspect import isgenerator` 会拿到本文件，直接 ImportError（实测：夹具全部生成失败）。
#
# 用法：
#   python ooxml_probe.py docx-text <file>
#   python ooxml_probe.py xlsx-cell <file> <sheet> <cell>
# 输出：一行 JSON（UTF-8）。读不出来就以非零退出并把原因写到 stderr，
# **不要**用空结果冒充"文件里没有内容"。
import json
import sys


def docx_text(path):
    from docx import Document

    doc = Document(path)
    parts = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                parts.append(cell.text)
    return {"text": "\n".join(parts)}


def xlsx_cell(path, sheet, cell):
    from openpyxl import load_workbook

    # data_only=False：要读**公式文本**（判定"是否保留了公式"必须看这个）。
    # 计算值另用 data_only=True 读缓存值；缓存可能不存在，由调用方决定是否可靠。
    wb = load_workbook(path, data_only=False)
    try:
        if sheet not in wb.sheetnames:
            raise SystemExit("sheet not found: " + sheet + " (have: " + ", ".join(wb.sheetnames) + ")")
        ws = wb[sheet]
        raw = ws[cell].value
        return {"formula": raw if isinstance(raw, str) else None, "value": raw}
    finally:
        wb.close()


def main(argv):
    if len(argv) < 3:
        print("usage: inspect.py docx-text <file> | xlsx-cell <file> <sheet> <cell>", file=sys.stderr)
        return 2
    kind = argv[1]
    try:
        if kind == "docx-text":
            result = docx_text(argv[2])
        elif kind == "xlsx-cell":
            if len(argv) != 5:
                print("usage: inspect.py xlsx-cell <file> <sheet> <cell>", file=sys.stderr)
                return 2
            result = xlsx_cell(argv[2], argv[3], argv[4])
        else:
            print("unknown kind: " + kind, file=sys.stderr)
            return 2
    except SystemExit:
        raise
    except Exception as error:   # noqa: BLE001 - 任何解析失败都要原样报出去
        print(type(error).__name__ + ": " + str(error), file=sys.stderr)
        return 1
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
