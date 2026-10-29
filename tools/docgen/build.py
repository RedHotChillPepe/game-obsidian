"""Сборка Word-документов проекта из Markdown-подобных исходников в tools/docgen/src.

Запуск:  python tools/docgen/build.py
Результат: папка docs/ в корне проекта.

Поддерживаемая разметка:
  # / ## / ###        заголовки (попадают в оглавление)
  - текст / "  - "    маркированный список (2 уровня)
  1. текст            нумерованный список
  | a | b |           таблица (первая строка — шапка, строку-разделитель |---| можно писать)
  > текст             выделенный блок-примечание
  **жирный**, *курсив*
  !title Заголовок | Подзаголовок     титульная страница
  !toc                оглавление (обновляется в Word: F9)
  !pagebreak          разрыв страницы
  !widths 3,1,1      относительные ширины колонок следующей таблицы
  !img путь | подпись картинка (если файла нет — рамка-заглушка со ссылкой)
"""
import datetime
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[2]
SRC = Path(__file__).resolve().parent / "src"
OUT = ROOT / "docs"

ACCENT = RGBColor(0x7A, 0x12, 0x12)
DARK = RGBColor(0x22, 0x22, 0x22)
MUTED = RGBColor(0x66, 0x66, 0x66)
HEADER_FILL = "3A1010"
ZEBRA_FILL = "F4EFEA"
NOTE_FILL = "F6EDE4"


def shade(element, fill):
    pr = element.get_or_add_tcPr() if hasattr(element, "get_or_add_tcPr") else element.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pr.append(shd)


def left_border(paragraph, color="7A1212"):
    ppr = paragraph._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    for k, v in {"w:val": "single", "w:sz": "24", "w:space": "8", "w:color": color}.items():
        left.set(qn(k), v)
    bdr.append(left)
    ppr.append(bdr)


def add_inline(paragraph, text, color=None, size=None):
    # **жирный** и *курсив*
    for part in re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*)", text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            run = paragraph.add_run(part[1:-1])
            run.italic = True
        else:
            run = paragraph.add_run(part)
        if color:
            run.font.color.rgb = color
        if size:
            run.font.size = size


def setup_styles(doc):
    st = doc.styles["Normal"]
    st.font.name = "Calibri"
    st.font.size = Pt(11)
    st.font.color.rgb = DARK
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    st.paragraph_format.space_after = Pt(6)
    st.paragraph_format.line_spacing = 1.15
    for name, size, color in (("Heading 1", 18, ACCENT), ("Heading 2", 14, ACCENT), ("Heading 3", 12, DARK)):
        h = doc.styles[name]
        h.font.name = "Calibri"
        h.font.size = Pt(size)
        h.font.bold = True
        h.font.color.rgb = color
        h.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
        h.element.rPr.rFonts.set(qn("w:ascii"), "Calibri")
        h.element.rPr.rFonts.set(qn("w:hAnsi"), "Calibri")
        h.paragraph_format.space_before = Pt(16 if name == "Heading 1" else 10)
        h.paragraph_format.space_after = Pt(6)
        h.paragraph_format.keep_with_next = True


def setup_page(doc, title):
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
    sec.left_margin = sec.right_margin = Cm(2.2)
    sec.top_margin = sec.bottom_margin = Cm(2.0)
    hp = sec.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_inline(hp, title, color=MUTED, size=Pt(8))
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = fp.add_run()
    for tag, text in (("begin", None), (None, "PAGE"), ("end", None)):
        if tag:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), tag)
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = text
        run._r.append(el)
    run.font.size = Pt(8)
    run.font.color.rgb = MUTED


def add_toc(doc):
    p = doc.add_paragraph()
    add_inline(p, "Содержание", color=ACCENT, size=Pt(16))
    p = doc.add_paragraph()
    run = p.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")  # Word обновит оглавление при открытии
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = 'TOC \\o "1-2" \\h \\z \\u'
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    txt = OxmlElement("w:t")
    txt.text = "Оглавление обновится при открытии (если нет — правый клик → «Обновить поле» или F9)."
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for el in (begin, instr, sep, txt, end):
        run._r.append(el)
    p.add_run().add_break(WD_BREAK.PAGE)


def add_title(doc, title, subtitle):
    for _ in range(6):
        doc.add_paragraph()
    p = doc.add_paragraph()
    add_inline(p, title, color=ACCENT, size=Pt(30))
    p.runs[0].bold = True
    if subtitle:
        p = doc.add_paragraph()
        add_inline(p, subtitle, color=MUTED, size=Pt(14))
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(30)
    add_inline(p, f"Версия 0.1 · {datetime.date.today():%d.%m.%Y} · рабочий документ", color=MUTED, size=Pt(10))
    p.add_run().add_break(WD_BREAK.PAGE)


def add_table(doc, rows, widths=None):
    rows = [r for r in rows if not re.fullmatch(r"\|?\s*:?-{2,}.*", r.strip())]
    cells = [[c.strip() for c in r.strip().strip("|").split("|")] for r in rows]
    ncol = max(len(r) for r in cells)
    table = doc.add_table(rows=len(cells), cols=ncol)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    total = Cm(16.6)
    for i, row in enumerate(cells):
        for j in range(ncol):
            cell = table.cell(i, j)
            share = widths[j] / sum(widths) if widths and len(widths) == ncol else 1 / ncol
            cell.width = int(total * share)
            cell.paragraphs[0].paragraph_format.space_after = Pt(2)
            text = row[j] if j < len(row) else ""
            if i == 0:
                if j == 0:
                    trpr = table.rows[0]._tr.get_or_add_trPr()
                    hdr = OxmlElement("w:tblHeader")
                    hdr.set(qn("w:val"), "true")
                    trpr.append(hdr)
                add_inline(cell.paragraphs[0], text, color=RGBColor(0xFF, 0xFF, 0xFF), size=Pt(10))
                for r in cell.paragraphs[0].runs:
                    r.bold = True
                shade(cell._tc, HEADER_FILL)
            else:
                add_inline(cell.paragraphs[0], text, size=Pt(10))
                if i % 2 == 0:
                    shade(cell._tc, ZEBRA_FILL)
    doc.add_paragraph()


def add_image(doc, spec, base):
    path, _, caption = spec.partition("|")
    path, caption = path.strip(), caption.strip()
    f = (base / path) if not Path(path).is_absolute() else Path(path)
    if not f.suffix:
        # путь без расширения: берём первый найденный файл с подходящим расширением
        f = next((f.with_suffix(e) for e in (".png", ".jpg", ".jpeg") if f.with_suffix(e).exists()), f)
    if f.exists() and f.is_file():
        try:
            doc.add_picture(str(f), width=Cm(7.5))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        except Exception:
            f = None
    else:
        f = None
    if f is None:
        p = doc.add_paragraph()
        shade(p._p, "EEEEEE")
        add_inline(p, f"[Изображение не встроено. Положите файл {path}.png или .jpg и пересоберите документ]", color=MUTED, size=Pt(9))
    if caption:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_inline(p, caption, color=MUTED, size=Pt(9))


def build(src_file):
    lines = src_file.read_text(encoding="utf-8").splitlines()
    doc = Document()
    setup_styles(doc)
    doc_title = src_file.stem
    for ln in lines:
        if ln.startswith("!title"):
            doc_title = ln[6:].split("|")[0].strip()
            break
    setup_page(doc, doc_title)
    i = 0
    para_buf = []
    widths = None

    def flush():
        if para_buf:
            p = doc.add_paragraph()
            add_inline(p, " ".join(para_buf))
            para_buf.clear()

    while i < len(lines):
        ln = lines[i].rstrip()
        s = ln.strip()
        if not s:
            flush()
        elif s.startswith("!title"):
            flush()
            t, _, sub = s[6:].partition("|")
            add_title(doc, t.strip(), sub.strip())
        elif s == "!toc":
            flush()
            add_toc(doc)
        elif s == "!pagebreak":
            flush()
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        elif s.startswith("!img"):
            flush()
            add_image(doc, s[4:], ROOT)
        elif s.startswith("!widths"):
            widths = [float(x) for x in s[7:].split(",")]
        elif s.startswith("|"):
            flush()
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                rows.append(lines[i])
                i += 1
            add_table(doc, rows, widths)
            widths = None
            continue
        elif s.startswith("### "):
            flush()
            doc.add_heading(s[4:], level=3)
        elif s.startswith("## "):
            flush()
            doc.add_heading(s[3:], level=2)
        elif s.startswith("# "):
            flush()
            doc.add_heading(s[2:], level=1)
        elif s.startswith("> "):
            flush()
            p = doc.add_paragraph()
            shade(p._p, NOTE_FILL)
            left_border(p)
            add_inline(p, s[2:])
        elif re.match(r"^\s*[-*] ", ln):
            flush()
            level = 2 if ln.startswith("  ") else 1
            p = doc.add_paragraph(style="List Bullet" if level == 1 else "List Bullet 2")
            p.paragraph_format.space_after = Pt(2)
            add_inline(p, re.sub(r"^\s*[-*] ", "", ln))
        elif re.match(r"^\d+\. ", s):
            flush()
            p = doc.add_paragraph(style="List Number")
            p.paragraph_format.space_after = Pt(2)
            add_inline(p, re.sub(r"^\d+\. ", "", s))
        else:
            para_buf.append(s)
        i += 1
    flush()
    OUT.mkdir(exist_ok=True)
    out = OUT / (src_file.stem + ".docx")
    doc.save(out)
    return out


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    targets = [Path(a) for a in sys.argv[1:]] or sorted(SRC.glob("*.md"))
    for t in targets:
        print("собран:", build(t))
