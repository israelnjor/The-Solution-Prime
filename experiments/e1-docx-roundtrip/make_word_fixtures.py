"""Create fixtures-word/: the fixtures opened and re-saved by Microsoft Word.

Word-authored XML is far messier than python-docx output (rsids, w14 attributes, proofing marks,
compatibility blocks), so it is the more realistic input for an import engine.
"""
from pathlib import Path

from word import Word

HERE = Path(__file__).resolve().parent
out = HERE / "fixtures-word"
out.mkdir(exist_ok=True)
with Word() as w:
    for src in sorted((HERE / "fixtures").glob("*.docx")):
        w.resave(src, out / src.name)
        print(f"wrote {out / src.name}")
