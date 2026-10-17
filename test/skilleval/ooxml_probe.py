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


def docx_xml(path):
    """docx 的 word/document.xml 原文（用于数 TOC 域、查节属性这类 python-docx 不暴露的东西）。"""
    import zipfile
    with zipfile.ZipFile(path) as zf:
        return {"xml": zf.read("word/document.xml").decode("utf-8", "replace")}


def docx_sections(path):
    """每节的页方向（python-docx 的 section.orientation 会按 pgSz 的宽高推断）。"""
    from docx import Document

    doc = Document(path)
    orientations = []
    for section in doc.sections:
        try:
            value = section.orientation
            orientations.append("landscape" if int(value) == 1 else "portrait")
        except Exception:   # noqa: BLE001
            orientations.append("unknown")
    return {"count": len(orientations), "orientations": orientations}


def pptx_slides(path):
    """按页序返回每页的文字（用于判页序、判某页有没有某个标题）。"""
    from pptx import Presentation

    pres = Presentation(path)
    slides = []
    for index, slide in enumerate(pres.slides, start=1):
        texts = []
        for shape in slide.shapes:
            try:
                if shape.has_text_frame and shape.text_frame.text.strip():
                    texts.append(shape.text_frame.text)
            except Exception:   # noqa: BLE001
                continue
        slides.append({"index": index, "texts": texts})
    return {"count": len(slides), "slides": slides}


def pptx_slide_xml(path, index):
    """某一页的 slideN.xml 原文（用于查动画 timing 这类 python-pptx 不暴露的东西）。"""
    import zipfile
    with zipfile.ZipFile(path) as zf:
        names = sorted(n for n in zf.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml"))
        if index < 1 or index > len(names):
            raise SystemExit("slide index out of range: " + str(index) + " (have " + str(len(names)) + ")")
        return {"xml": zf.read(names[index - 1]).decode("utf-8", "replace")}


def main(argv):
    if len(argv) < 3:
        print("usage: ooxml_probe.py <kind> <file> [args...]", file=sys.stderr)
        return 2
    kind = argv[1]
    try:
        if kind == "docx-text":
            result = docx_text(argv[2])
        elif kind == "docx-xml":
            result = docx_xml(argv[2])
        elif kind == "docx-sections":
            result = docx_sections(argv[2])
        elif kind == "pptx-slides":
            result = pptx_slides(argv[2])
        elif kind == "pptx-slide-xml":
            if len(argv) != 4:
                print("usage: ooxml_probe.py pptx-slide-xml <file> <index>", file=sys.stderr)
                return 2
            result = pptx_slide_xml(argv[2], int(argv[3]))
        elif kind == "xlsx-cell":
            if len(argv) != 5:
                print("usage: ooxml_probe.py xlsx-cell <file> <sheet> <cell>", file=sys.stderr)
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
