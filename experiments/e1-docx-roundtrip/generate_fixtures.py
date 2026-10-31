"""Generate the Experiment 1 representative DOCX corpus (fixtures A-D).

Every paragraph and table cell carries a unique marker such as [A-012] so the
round-trip checker can detect lost, duplicated or reordered content precisely.
Output is deterministic: re-running produces the same documents.

Usage:  python generate_fixtures.py [--out fixtures]
"""
import argparse
import io
import json
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from PIL import Image, ImageDraw

LOREM = (
    "Prime Word must preserve the meaning and structure of professional documents "
    "through every import, edit and export, without silently losing any content. "
    "Reliable editing builds the trust that every later feature depends on."
)


class Markers:
    """Hands out sequential markers: [A-001], [A-002], ..."""

    def __init__(self, prefix):
        self.prefix = prefix
        self.n = 0

    def __call__(self):
        self.n += 1
        return f"[{self.prefix}-{self.n:03d}]"


# --- low-level helpers for features python-docx has no API for -------------

def add_hyperlink(paragraph, url, text):
    r_id = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), r_id)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rpr.extend([color, underline])
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    run.extend([rpr, t])
    link.append(run)
    paragraph._p.append(link)


def add_field(paragraph, instr, placeholder):
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), instr)
    run = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = placeholder
    run.append(t)
    fld.append(run)
    paragraph._p.append(fld)


def page_x_of_y(paragraph, label):
    paragraph.add_run(f"{label} — Page ")
    add_field(paragraph, " PAGE ", "1")
    paragraph.add_run(" of ")
    add_field(paragraph, " NUMPAGES ", "1")


def repeat_header_row(row):
    tr_pr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    tr_pr.append(el)


def shade(cell, fill):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(shd)


def chart_png(seed, size=(600, 300)):
    """A deterministic bar-chart-like image; no text so no font dependency."""
    w, h = size
    img = Image.new("RGB", size, (248, 250, 252))
    d = ImageDraw.Draw(img)
    d.line([(40, h - 30), (w - 20, h - 30)], fill=(100, 116, 139), width=2)
    d.line([(40, 20), (40, h - 30)], fill=(100, 116, 139), width=2)
    palette = [(37, 99, 235), (13, 148, 136), (234, 88, 12), (147, 51, 234)]
    for i in range(8):
        bar_h = 40 + ((i * 37 + seed * 53) % (h - 90))
        x0 = 60 + i * ((w - 90) // 8)
        d.rectangle([x0, h - 30 - bar_h, x0 + 40, h - 31], fill=palette[(i + seed) % 4])
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


def new_document(title):
    doc = Document()
    doc.core_properties.title = title
    doc.core_properties.author = "Prime Productivity — Experiment 1"
    return doc


def fill_table(table, rows, m, header_fill=None):
    for r, values in enumerate(rows):
        for c, value in enumerate(values):
            cell = table.cell(r, c)
            cell.text = f"{value} {m()}"
            if r == 0 and header_fill:
                shade(cell, header_fill)
                for run in cell.paragraphs[0].runs:
                    run.bold = True


# --- fixtures ----------------------------------------------------------------

def fixture_a():
    """A — Basic formatting."""
    m = Markers("A")
    doc = new_document("Fixture A — Basic Formatting")
    doc.add_heading(f"Fixture A — Basic Formatting {m()}", level=0)
    doc.add_heading(f"1. Headings and paragraphs {m()}", level=1)
    doc.add_paragraph(f"{LOREM} {m()}")
    doc.add_heading(f"1.1 Second-level heading {m()}", level=2)
    doc.add_paragraph(f"A plain body paragraph under a level-two heading. {m()}")
    doc.add_heading(f"1.1.1 Third-level heading {m()}", level=3)
    doc.add_paragraph(f"A plain body paragraph under a level-three heading. {m()}")

    doc.add_heading(f"2. Character formatting {m()}", level=1)
    p = doc.add_paragraph()
    p.add_run("Bold").bold = True
    p.add_run(", ")
    p.add_run("italic").italic = True
    p.add_run(", ")
    p.add_run("underlined").underline = True
    p.add_run(", ")
    r = p.add_run("struck through")
    r.font.strike = True
    p.add_run(", ")
    r = p.add_run("bold italic")
    r.bold = r.italic = True
    p.add_run(", H")
    p.add_run("2").font.subscript = True
    p.add_run("O and E = mc")
    p.add_run("2").font.superscript = True
    p.add_run(f". {m()}")

    p = doc.add_paragraph()
    for name, rgb in [("Red", "C00000"), ("Green", "00B050"), ("Blue", "0070C0"), ("Purple", "7030A0")]:
        p.add_run(f"{name} text").font.color.rgb = RGBColor.from_string(rgb)
        p.add_run(" · ")
    r = p.add_run("highlighted text")
    r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    p.add_run(f" {m()}")

    p = doc.add_paragraph()
    for font, size in [("Arial", 11), ("Times New Roman", 12), ("Georgia", 14), ("Courier New", 10), ("Calibri", 18)]:
        r = p.add_run(f"{font} {size}pt")
        r.font.name = font
        r.font.size = Pt(size)
        p.add_run(" | ")
    p.add_run(m())

    doc.add_heading(f"3. Paragraph alignment {m()}", level=1)
    for label, align in [("Left aligned", WD_ALIGN_PARAGRAPH.LEFT), ("Centered", WD_ALIGN_PARAGRAPH.CENTER),
                         ("Right aligned", WD_ALIGN_PARAGRAPH.RIGHT), ("Justified", WD_ALIGN_PARAGRAPH.JUSTIFY)]:
        p = doc.add_paragraph(f"{label}: {LOREM} {m()}")
        p.alignment = align

    p = doc.add_paragraph(f"Indented paragraph with extra spacing before and after. {m()}")
    p.paragraph_format.left_indent = Inches(0.5)
    p.paragraph_format.first_line_indent = Inches(0.25)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(12)

    doc.add_heading(f"4. Lists {m()}", level=1)
    for text in ["First bullet", "Second bullet", "Third bullet"]:
        doc.add_paragraph(f"{text} {m()}", style="List Bullet")
        if text == "Second bullet":
            doc.add_paragraph(f"Nested bullet {m()}", style="List Bullet 2")
            doc.add_paragraph(f"Another nested bullet {m()}", style="List Bullet 2")
    for text in ["Step one", "Step two", "Step three"]:
        doc.add_paragraph(f"{text} {m()}", style="List Number")
        if text == "Step two":
            doc.add_paragraph(f"Sub-step {m()}", style="List Number 2")

    doc.add_heading(f"5. Special characters {m()}", level=1)
    doc.add_paragraph(f"Punctuation: “curly quotes” ‘single’ — em dash – en dash … ellipsis {m()}")
    doc.add_paragraph(f"Symbols: © ® ™ € £ ¥ § ¶ ° ± × ÷ ≤ ≥ ≠ → ✓ {m()}")
    doc.add_paragraph(f"Accents and scripts: café naïve Zürich São Paulo Ελληνικά Русский 中文 {m()}")
    doc.add_paragraph(f"Non-breaking space and a\ttab character {m()}")
    p = doc.add_paragraph(f"Line one of a paragraph with a manual line break {m()}")
    p.add_run().add_break()
    p.add_run(f"Line two after the break {m()}")
    return doc


def fixture_b():
    """B — Structured content: tables, hyperlinks, images, captions."""
    m = Markers("B")
    doc = new_document("Fixture B — Structured Content")
    doc.add_heading(f"Fixture B — Structured Content {m()}", level=0)
    doc.add_paragraph(f"This document exercises tables, hyperlinks and images. {m()}")

    doc.add_heading(f"1. Simple table {m()}", level=1)
    t = doc.add_table(rows=4, cols=3)
    t.style = "Table Grid"
    fill_table(t, [["Feature", "Priority", "Owner"],
                   ["DOCX import", "Critical", "Editor team"],
                   ["Autosave", "Critical", "Platform team"],
                   ["Templates", "High", "Design team"]], m, header_fill="D9E2F3")
    doc.add_paragraph(f"Paragraph after the simple table. {m()}")

    doc.add_heading(f"2. Merged cells and alignment {m()}", level=1)
    t = doc.add_table(rows=4, cols=4)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    fill_table(t, [["Quarter", "Plan", "Actual", "Variance"],
                   ["Q1", "100", "96", "-4"],
                   ["Q2", "120", "131", "+11"],
                   ["Total", "220", "227", "+7"]], m, header_fill="E2EFDA")
    merged = t.cell(0, 1).merge(t.cell(0, 2))  # horizontal merge
    merged.paragraphs[0].runs[0].text = "Plan and actual"
    t.cell(1, 3).merge(t.cell(2, 3))           # vertical merge
    for row in t.rows[1:]:
        row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_paragraph(f"Paragraph after the merged table. {m()}")

    doc.add_heading(f"3. Hyperlinks {m()}", level=1)
    p = doc.add_paragraph("Visit ")
    add_hyperlink(p, "https://example.com/prime-word", "the Prime Word project page")
    p.add_run(f" for details. {m()}")
    p = doc.add_paragraph("A second link: ")
    add_hyperlink(p, "https://example.org/docs?section=tables&lang=en", "documentation with a query string")
    p.add_run(f" {m()}")

    doc.add_heading(f"4. Images {m()}", level=1)
    doc.add_paragraph(f"An inline image follows, centered, with a caption. {m()}")
    doc.add_picture(chart_png(1), width=Inches(5))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    try:
        cap = doc.add_paragraph(f"Figure 1: Synthetic chart image {m()}", style="Caption")
    except KeyError:
        cap = doc.add_paragraph(f"Figure 1: Synthetic chart image {m()}")
        cap.runs[0].italic = True
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading(f"5. Image inside a table {m()}", level=1)
    t = doc.add_table(rows=2, cols=2)
    t.style = "Table Grid"
    t.cell(0, 0).text = f"Thumbnail {m()}"
    t.cell(0, 1).text = f"Description {m()}"
    t.cell(1, 0).paragraphs[0].add_run().add_picture(chart_png(2, (300, 150)), width=Inches(1.5))
    t.cell(1, 1).text = f"A smaller image placed inside a table cell. {m()}"
    doc.add_paragraph(f"Closing paragraph of fixture B. {m()}")
    return doc


def fixture_c():
    """C — Multi-page layout: margins, page breaks, headers/footers, sections."""
    m = Markers("C")
    doc = new_document("Fixture C — Multi-page Layout")
    s1 = doc.sections[0]
    s1.top_margin = Inches(0.75)
    s1.bottom_margin = Inches(0.75)
    s1.left_margin = Inches(1.25)
    s1.right_margin = Inches(1.0)
    s1.different_first_page_header_footer = True
    s1.first_page_header.paragraphs[0].text = f"Cover page header {m()}"
    s1.header.paragraphs[0].text = f"Prime Word — Fixture C {m()}"
    s1.header.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    page_x_of_y(s1.footer.paragraphs[0], "Portrait section")
    s1.footer.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading(f"Fixture C — Multi-page Layout {m()}", level=0)
    doc.add_paragraph(f"Cover page. The first page has its own header. {m()}")
    doc.add_page_break()

    for chapter in range(1, 4):
        doc.add_heading(f"Chapter {chapter} {m()}", level=1)
        for _ in range(4):
            doc.add_paragraph(f"{LOREM} {LOREM} {m()}")
        if chapter < 3:
            doc.add_page_break()

    doc.add_heading(f"Table in a paginated document {m()}", level=1)
    t = doc.add_table(rows=6, cols=3)
    t.style = "Table Grid"
    fill_table(t, [["Page element", "Expected", "Notes"]] +
               [[f"Element {i}", "Preserved", "Check visually"] for i in range(1, 6)], m, header_fill="FCE4D6")

    s2 = doc.add_section(WD_SECTION.NEW_PAGE)
    s2.orientation = WD_ORIENT.LANDSCAPE
    s2.page_width, s2.page_height = s1.page_height, s1.page_width
    s2.different_first_page_header_footer = False
    s2.header.is_linked_to_previous = False
    s2.header.paragraphs[0].text = f"Landscape section header {m()}"
    s2.footer.is_linked_to_previous = False
    page_x_of_y(s2.footer.paragraphs[0], "Landscape section")

    doc.add_heading(f"Landscape section {m()}", level=1)
    doc.add_paragraph(f"This section is landscape with its own header and footer. {m()}")
    t = doc.add_table(rows=3, cols=6)
    t.style = "Table Grid"
    fill_table(t, [[f"Col {c}" for c in range(1, 7)]] +
               [[f"R{r}C{c}" for c in range(1, 7)] for r in (1, 2)], m)
    return doc


def fixture_d():
    """D — Complex report: everything combined, tables crossing page boundaries."""
    m = Markers("D")
    doc = new_document("Fixture D — Complex Report")
    sec = doc.sections[0]
    sec.header.paragraphs[0].text = f"Annual Operations Report — Confidential {m()}"
    page_x_of_y(sec.footer.paragraphs[0], "Fixture D")
    sec.footer.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading(f"Annual Operations Report {m()}", level=0)
    p = doc.add_paragraph()
    p.add_run("Prepared for: ").bold = True
    p.add_run(f"Executive Committee {m()}")
    p = doc.add_paragraph()
    p.add_run("Classification: ").bold = True
    r = p.add_run("Internal")
    r.font.color.rgb = RGBColor.from_string("C00000")
    p.add_run(f" {m()}")

    doc.add_heading(f"1. Executive summary {m()}", level=1)
    for _ in range(3):
        p = doc.add_paragraph(f"{LOREM} {LOREM} {m()}")
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    for text in ["Revenue grew in every region", "Operating costs fell", "Customer satisfaction improved"]:
        doc.add_paragraph(f"{text} {m()}", style="List Bullet")

    doc.add_heading(f"2. Regional performance {m()}", level=1)
    doc.add_paragraph(f"The table below is long enough to cross a page boundary; its header row "
                      f"is marked to repeat on each page. {m()}")
    t = doc.add_table(rows=1, cols=5)
    t.style = "Table Grid"
    for c, h in enumerate(["Region", "Q1", "Q2", "Q3", "Q4"]):
        t.cell(0, c).text = f"{h} {m()}"
        shade(t.cell(0, c), "D9E2F3")
    repeat_header_row(t.rows[0])
    for i in range(1, 46):
        cells = t.add_row().cells
        cells[0].text = f"Region {i:02d} {m()}"
        for q in range(1, 5):
            cells[q].text = f"{(i * 17 + q * 31) % 900 + 100} {m()}"

    doc.add_heading(f"3. Analysis {m()}", level=1)
    for section in ["Market conditions", "Operational efficiency", "Risk outlook"]:
        doc.add_heading(f"{section} {m()}", level=2)
        p = doc.add_paragraph()
        p.add_run("Key point: ").bold = True
        p.add_run(f"{LOREM} ")
        p.add_run("This sentence is italic for emphasis. ").italic = True
        p.add_run(f"{LOREM} {m()}")
        for step in ["Assess", "Decide", "Act"]:
            doc.add_paragraph(f"{step} — {section.lower()} {m()}", style="List Number")

    doc.add_picture(chart_png(3), width=Inches(6))
    doc.add_paragraph(f"Figure 1: Quarterly trend {m()}").alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading(f"4. Budget detail {m()}", level=1)
    t = doc.add_table(rows=1, cols=4)
    t.style = "Table Grid"
    for c, h in enumerate(["Line item", "Budget", "Actual", "Comment"]):
        t.cell(0, c).text = f"{h} {m()}"
    repeat_header_row(t.rows[0])
    for i in range(1, 31):
        cells = t.add_row().cells
        cells[0].text = f"Item {i:02d} {m()}"
        cells[1].text = f"{i * 1000:,} {m()}"
        cells[2].text = f"{i * 1000 + (i % 7 - 3) * 150:,} {m()}"
        cells[3].text = f"{'Over' if i % 7 > 3 else 'Within'} budget {m()}"

    doc.add_heading(f"5. Conclusion {m()}", level=1)
    p = doc.add_paragraph(f"{LOREM} Further reading: ")
    add_hyperlink(p, "https://example.com/annual-report", "full annual report")
    p.add_run(f". {m()}")
    doc.add_paragraph(f"End of report. {m()}")
    return doc


FIXTURES = {
    "A_basic_formatting": (fixture_a, ["headings", "character formatting", "colors", "fonts", "alignment",
                                       "indentation", "bullets", "numbered lists", "nested lists",
                                       "special characters", "line breaks"]),
    "B_structured_content": (fixture_b, ["tables", "cell shading", "merged cells", "hyperlinks",
                                         "inline images", "captions", "image in table"]),
    "C_multipage_layout": (fixture_c, ["margins", "page breaks", "headers", "footers", "page numbers",
                                       "different first page", "sections", "landscape"]),
    "D_complex_report": (fixture_d, ["mixed formatting", "long tables across pages", "repeating header rows",
                                     "lists", "images", "hyperlinks", "headers/footers", "page numbers"]),
}


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--out", default=str(Path(__file__).parent / "fixtures"))
    args = ap.parse_args()
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    manifest = {}
    for name, (build, features) in FIXTURES.items():
        path = out / f"{name}.docx"
        build().save(path)
        manifest[name] = {"file": path.name, "description": build.__doc__.strip(), "features": features}
        print(f"wrote {path}")
    (out / "manifest.json").write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")


if __name__ == "__main__":
    main()
