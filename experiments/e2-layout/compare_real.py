"""Compare our layout with Word's on ANY document (no test markers needed).

Word and our engine each list every paragraph (text, page, top of first line). Paragraphs are
matched by text, then we report agreement and the places where the two start to diverge.
Usage: python compare_real.py <file.docx> [--show N]
"""
import difflib
import json
import re
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
sys.path.insert(0, str(ROOT / "experiments" / "e1-docx-roundtrip"))
from word import WD_DO_NOT_SAVE_CHANGES, WD_STATISTIC_PAGES, Word  # noqa: E402

CLEAN = re.compile(r"[\r\x07\x0b\x0c\x01\x02\x13\x14\x15]")


def key(text):
    return re.sub(r"\s+", " ", CLEAN.sub("", text)).strip()[:40]


def word_paragraphs(path):
    with Word() as w:
        doc = w._open(path)
        try:
            doc.Repaginate()
            out = []
            for p in doc.Paragraphs:
                text = key(p.Range.Text)
                if not text:
                    continue
                r = p.Range.Duplicate
                r.Collapse(1)
                out.append({"text": text, "page": r.Information(3), "y": round(r.Information(6), 1)})
            return doc.ComputeStatistics(WD_STATISTIC_PAGES), out
        finally:
            doc.Close(WD_DO_NOT_SAVE_CHANGES)


def our_paragraphs(path):
    cli = ROOT / "packages" / "layout" / "dist" / "cli.js"
    data = json.loads(subprocess.run(["node", str(cli), "paragraphs", str(path)], capture_output=True, text=True,
                                     check=True, encoding="utf-8").stdout)
    return data["pages"], [{**p, "text": key(p["text"])} for p in data["paragraphs"] if key(p["text"])]


def main():
    path = Path(sys.argv[1]).resolve()
    show = int(sys.argv[sys.argv.index("--show") + 1]) if "--show" in sys.argv else 15
    sys.stdout.reconfigure(encoding="utf-8")
    wp, wl = word_paragraphs(path)
    op, ol = our_paragraphs(path)
    sm = difflib.SequenceMatcher(a=[p["text"] for p in wl], b=[p["text"] for p in ol], autojunk=False)
    pairs = [(wl[b.a + k], ol[b.b + k]) for b in sm.get_matching_blocks() for k in range(b.size)]
    same = [(w, o) for w, o in pairs if w["page"] == o["page"]]
    close = [(w, o) for w, o in same if abs(w["y"] - o["y"]) <= 2]
    print(f"pages: Word {wp}, ours {op}")
    print(f"paragraphs: Word {len(wl)}, ours {len(ol)}, matched {len(pairs)}")
    print(f"same page {len(same)}/{len(pairs)}, within 2pt {len(close)}/{len(pairs)}")
    print("\nwhere the error changes (Word page/y -> ours page/y, delta):")
    prev, shown = None, 0
    for w, o in pairs:
        d = round(o["y"] - w["y"], 1) if w["page"] == o["page"] else f"page {w['page']}->{o['page']}"
        changed = prev is None or (isinstance(d, str) != isinstance(prev, str)) or \
            (not isinstance(d, str) and abs(d - prev) > 2)
        if changed and shown < show:
            print(f"  {w['page']}/{w['y']:>6} -> {o['page']}/{o['y']:>6}  {d!s:>12}  {w['text']}")
            shown += 1
        prev = d


if __name__ == "__main__":
    main()
