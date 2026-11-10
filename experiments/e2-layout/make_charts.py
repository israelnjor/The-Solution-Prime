"""Has Word make Wave 6's charts fixture (O_charts.docx): a clustered column, a clustered bar, a line and
a pie chart (Word's sample data, titles set), inline; and its PDF."""
import sys
from pathlib import Path

ROOT = Path(r"C:\dev\The Solution Prime")
sys.path.insert(0, str(ROOT / "experiments" / "e1-docx-roundtrip"))
from word import Word  # noqa: E402

out = ROOT / "experiments" / "e1-docx-roundtrip" / "fixtures-wave6" / "O_charts.docx"
KINDS = [(51, "Column chart [O-002]"), (57, "Bar chart [O-003]"), (4, "Line chart [O-004]"), (5, "Pie chart [O-005]")]
with Word() as w:
    doc = w.app.Documents.Add()
    try:
        doc.Content.Text = "Charts [O-001]\r"
        for kind, title in KINDS:
            r = doc.Content
            r.InsertParagraphAfter()
            r = doc.Paragraphs(doc.Paragraphs.Count).Range
            shape = doc.InlineShapes.AddChart2(-1, kind, r)
            chart = shape.Chart
            chart.ChartData.Workbook.Close()
            chart.HasTitle = True
            chart.ChartTitle.Text = title
            shape.Width = 360
            shape.Height = 200
        doc.SaveAs2(str(out), 16)
        doc.ExportAsFixedFormat(str(out.with_suffix(".word.pdf")), 17)
        print("saved", out.name, "charts", doc.InlineShapes.Count)
    finally:
        doc.Close(0)
