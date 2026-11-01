"""Real-document sweep: every .docx in a folder against Microsoft Word.

For each document: our layout vs Word's (pages; paragraphs on the same page / within 2pt), and a round
trip through the editor (open -> edit model -> save) that Word must reopen with the same page count.
Writes sweep-report.md into the folder (keep personal documents in a corpus-private folder).

Usage: python sweep.py <folder>
"""
import difflib
import os
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "e1-docx-roundtrip"))
import compare_real  # noqa: E402
from word import Word  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
ROUNDTRIP = ROOT / "packages" / "editor" / "test" / "edit-roundtrip.mjs"


def layout(path: Path):
    wp, wl = compare_real.word_paragraphs(path)
    op, ol = compare_real.our_paragraphs(path)
    sm = difflib.SequenceMatcher(a=[p["text"] for p in wl], b=[p["text"] for p in ol], autojunk=False)
    pairs = [(wl[b.a + k], ol[b.b + k]) for b in sm.get_matching_blocks() for k in range(b.size)]
    same = [(w, o) for w, o in pairs if w["page"] == o["page"]]
    close = [(w, o) for w, o in same if abs(w["y"] - o["y"]) <= 2]
    n = max(len(pairs), 1)
    return wp, op, f"{100 * len(same) // n}%", f"{100 * len(close) // n}%"


def roundtrip(path: Path, pages: int) -> str:
    out = Path(tempfile.gettempdir()) / f"prime-sweep-{os.getpid()}.docx"
    r = subprocess.run(["node", str(ROUNDTRIP), str(path), str(out)], capture_output=True, text=True, timeout=300)
    if r.returncode != 0:
        return "save failed: " + (r.stderr.strip().splitlines() or ["?"])[-1][:80]
    try:
        # Its own Word session: Word shares one hidden instance, which the layout step closes.
        with Word() as word:
            again = word.inspect(str(out))["pages"]
        return "ok" if again == pages else f"Word reopens with {again} pages (was {pages})"
    except Exception as e:  # Word could not open our file
        return f"Word cannot open the saved file: {e}"[:120]
    finally:
        out.unlink(missing_ok=True)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    folder = Path(sys.argv[1]).resolve()
    docs = sorted(p for p in folder.rglob("*.docx") if not p.name.startswith("~$"))
    rows = []
    for doc in docs:
        try:
            wp, op, same, close = layout(doc)
            trip = roundtrip(doc, wp)
            rows.append(f"| {doc.name} | {wp} → {op} | {same} | {close} | {trip} |")
        except Exception as e:
            rows.append(f"| {doc.name} | — | — | — | error: {str(e)[:80]} |")
        print(rows[-1], flush=True)
    report = folder / "sweep-report.md"
    report.write_text("# Real-document sweep\n\n| Document | Pages (Word → ours) | Same page | Within 2pt | Round trip |\n"
                      "|---|---|---|---|---|\n" + "\n".join(rows) + "\n", encoding="utf-8")
    print("report:", report)


if __name__ == "__main__":
    main()
