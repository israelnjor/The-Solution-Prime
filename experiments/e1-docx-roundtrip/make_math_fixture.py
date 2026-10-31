"""Create fixtures-preservation/G_math.docx with Microsoft Word: common equation structures.

Each equation is typed in Word's linear format and built up by Word itself, so the OMML is exactly
what Word produces: fractions, radicals, scripts, n-ary operators with limits, a matrix, accents.
"""
from pathlib import Path

from word import Word

HERE = Path(__file__).resolve().parent
OUT = HERE / "fixtures-preservation" / "G_math.docx"
EQUATIONS = [
    "Quadratic formula [G-001]", "x=(-b±√(b^2-4ac))/(2a)",
    "Sum with limits [G-002]", "∑_(i=1)^n i^2=(n(n+1)(2n+1))/6",
    "Matrix [G-003]", "A=[■(1&0@0&1)]",
    "Radical with degree [G-004]", "√(3&x+1)+e^(iπ)=0",
    "Integral and subscript [G-005]", "∫_0^1 x_k dx",
]

with Word() as w:
    doc = w.app.Documents.Add()
    for i in range(0, len(EQUATIONS), 2):
        label, linear = EQUATIONS[i], EQUATIONS[i + 1]
        end = doc.Content.End - 1
        r = doc.Range(end, end)
        r.InsertAfter(label)
        r.InsertParagraphAfter()
        end = doc.Content.End - 1
        r = doc.Range(end, end)
        r.InsertAfter(linear)
        eq = doc.Range(r.Start, r.End)
        r.InsertParagraphAfter()
        doc.OMaths.Add(eq)
    for k in range(1, doc.OMaths.Count + 1):
        doc.OMaths(k).BuildUp()
    doc.SaveAs2(FileName=str(OUT), FileFormat=16)
    doc.Close(0)
print(f"wrote {OUT}")
