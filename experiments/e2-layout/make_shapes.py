"""Has Word make Wave 6's shapes fixture (N_shapes.docx): a theme-styled rectangle, a red oval, a green
right arrow, an unfilled triangle, a line, and a text box with text, fill and border; and its PDF."""
import sys
from pathlib import Path

ROOT = Path(r"C:\dev\The Solution Prime")
sys.path.insert(0, str(ROOT / "experiments" / "e1-docx-roundtrip"))
from word import Word  # noqa: E402

out = ROOT / "experiments" / "e1-docx-roundtrip" / "fixtures-wave6" / "N_shapes.docx"
with Word() as w:
    doc = w.app.Documents.Add()
    try:
        doc.Content.Text = "Shapes and text boxes [N-001]\rThis paragraph anchors the shapes below. [N-002]\r" + ("Body text. " * 40) + "[N-003]\r"
        anchor = doc.Paragraphs(2).Range
        s = doc.Shapes
        rect = s.AddShape(1, 72, 140, 120, 70, anchor)          # msoShapeRectangle, theme style
        oval = s.AddShape(9, 220, 140, 100, 70, anchor)         # msoShapeOval
        oval.Fill.ForeColor.RGB = 0x0000C0                      # BGR: red C00000
        oval.Line.ForeColor.RGB = 0x000000
        arrow = s.AddShape(33, 350, 150, 110, 50, anchor)       # msoShapeRightArrow
        arrow.Fill.ForeColor.RGB = 0x50B000                     # 00B050 green
        tri = s.AddShape(7, 480, 140, 70, 70, anchor)           # msoShapeIsoscelesTriangle
        tri.Fill.Visible = 0
        tri.Line.ForeColor.RGB = 0x9A572B                       # 2B579A
        tri.Line.Weight = 2
        line = s.AddLine(72, 240, 300, 260, anchor)
        line.Line.Weight = 1.5
        box = s.AddTextbox(1, 330, 230, 200, 70, anchor)
        box.TextFrame.TextRange.Text = "A text box with a yellow fill and a blue border. [N-004]"
        box.Fill.ForeColor.RGB = 0xCCF2FF                       # FFF2CC
        box.Line.ForeColor.RGB = 0xC47244                       # 4472C4
        box.Line.Weight = 1.5
        for i in range(1, s.Count + 1):
            sh = s(i)
            sh.WrapFormat.Type = 3                              # wdWrapNone: in front of text
        doc.SaveAs2(str(out), 16)
        doc.ExportAsFixedFormat(str(out.with_suffix(".word.pdf")), 17)
        print("saved", out.name, "shapes", s.Count)
    finally:
        doc.Close(0)
