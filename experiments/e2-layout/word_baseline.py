"""Record where Microsoft Word places every content marker: page number and vertical position.

This is the ground truth for Experiment 2: our layout engine must put each paragraph on the same
page as Word, at nearly the same height. Output: baselines/<fixture>.json
Usage: python word_baseline.py [fixtures_dir]
"""
import json
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "e1-docx-roundtrip"))
from word import WD_DO_NOT_SAVE_CHANGES, WD_STATISTIC_PAGES, Word  # noqa: E402

MARKER = re.compile(r"\[[A-Z]-\d{3}\]")
WD_ACTIVE_END_PAGE_NUMBER = 3
WD_VERTICAL_POSITION_RELATIVE_TO_PAGE = 6
WD_COLLAPSE_START = 1


def baseline(word, path):
    doc = word._open(path)
    try:
        doc.Repaginate()
        markers = {}
        for para in doc.Paragraphs:
            rng = para.Range
            found = MARKER.findall(rng.Text)
            if not found:
                continue
            start = rng.Duplicate
            start.Collapse(WD_COLLAPSE_START)
            pos = {"page": start.Information(WD_ACTIVE_END_PAGE_NUMBER),
                   "y": round(start.Information(WD_VERTICAL_POSITION_RELATIVE_TO_PAGE), 1)}
            for m in found:
                markers.setdefault(m, pos)
        return {"pages": doc.ComputeStatistics(WD_STATISTIC_PAGES), "markers": markers}
    finally:
        doc.Close(WD_DO_NOT_SAVE_CHANGES)


def main():
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE.parent / "e1-docx-roundtrip" / "fixtures"
    out = HERE / "baselines"
    out.mkdir(exist_ok=True)
    with Word() as w:
        for f in sorted(src.glob("*.docx")):
            data = baseline(w, f)
            (out / f"{f.stem}.json").write_text(json.dumps(data, indent=1), encoding="utf-8")
            print(f"{f.stem}: {data['pages']} pages, {len(data['markers'])} markers")


if __name__ == "__main__":
    main()
