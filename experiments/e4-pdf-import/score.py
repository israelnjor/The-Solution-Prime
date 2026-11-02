"""Experiment 4 scorer: PDF -> Word conversion against the original Word documents.

Each corpus PDF was exported by Word from a fixture document, so the original .docx is the answer key:
  python score.py              # converts every corpus PDF with prime-pdf-import, then scores
  python score.py --no-word    # skip Word page counts

Measures, per document: markers recovered ([X-NNN] tags in the text), text similarity, original
paragraphs recovered intact, formatting at each marker (bold, italic, size, heading, list, table,
alignment), tables and pictures, and pages when Word opens the result.
"""
import difflib
import json
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(ROOT / "experiments" / "e1-docx-roundtrip"))
from docxfp import MARKER_RE, PROPS, Fingerprint  # noqa: E402

CLI = ROOT / "packages" / "pdf-import" / "dist" / "cli.js"
SOURCES = [ROOT / "experiments" / "e1-docx-roundtrip" / "fixtures", ROOT / "experiments" / "e1-docx-roundtrip" / "fixtures-preservation"]
OUT = HERE / "results"


def original_for(pdf: Path) -> Path:
    for d in SOURCES:
        if (d / f"{pdf.stem}.docx").exists():
            return d / f"{pdf.stem}.docx"
    raise FileNotFoundError(pdf.stem)


def role(p):
    style = (p["style"] or "").lower()
    if "table" in p["where"]:
        return "table"
    if p["list"]:
        return "list"
    if style.startswith("heading") or style == "title":
        return "heading"
    return "body"


def marker_index(fp):
    """marker -> (paragraph, formatting of the marker's first character)."""
    out = {}
    for p in fp.paragraphs:
        for m in MARKER_RE.finditer(p["text"]):
            chars = p["chars"]
            i = m.start() if len(chars) == len(p["text"]) else 0
            props = chars[i] if i < len(chars) else (chars[0] if chars else ())
            out.setdefault(m.group(0), (p, dict(zip(PROPS, props)) if isinstance(props, tuple) else dict(props)))
    return out


def body_text(fp):
    return " ".join(p["text"] for p in fp.paragraphs if p["text"])


def score(orig_path: Path, conv_path: Path):
    a, b = Fingerprint(str(orig_path)), Fingerprint(str(conv_path))
    ma, mb = marker_index(a), marker_index(b)
    markers = list(ma)
    found = [m for m in markers if m in mb]
    agree = {"bold": 0, "italic": 0, "size": 0, "role": 0, "align": 0}
    misses = []
    for m in found:
        (pa, ca), (pb, cb) = ma[m], mb[m]
        ok = {
            "bold": bool(ca.get("b")) == bool(cb.get("b")),
            "italic": bool(ca.get("i")) == bool(cb.get("i")),
            "size": abs(float(ca.get("sz") or 0) - float(cb.get("sz") or 0)) <= 0.5,
            "role": role(pa) == role(pb),
            "align": pa["align"] == pb["align"],
        }
        for k, v in ok.items():
            agree[k] += v
        if not all(ok.values()) and len(misses) < 6:
            misses.append(f"{m} " + ", ".join(f"{k}: {role(pa) if k == 'role' else pa['align'] if k == 'align' else ca.get({'bold': 'b', 'italic': 'i', 'size': 'sz'}[k])}"
                                              f" -> {role(pb) if k == 'role' else pb['align'] if k == 'align' else cb.get({'bold': 'b', 'italic': 'i', 'size': 'sz'}[k])}"
                                              for k, v in ok.items() if not v))
    texts_b = {p["text"] for p in b.paragraphs}
    paras = [p for p in a.paragraphs if p["text"].strip() and "table" not in p["where"]]
    intact = sum(p["text"] in texts_b for p in paras)
    sim = difflib.SequenceMatcher(None, body_text(a), body_text(b), autojunk=False).ratio()
    pct = lambda n, d: round(100 * n / d) if d else 100  # noqa: E731
    return {
        "markers": f"{len(found)}/{len(markers)}",
        "text": f"{round(100 * sim)}%",
        "paragraphs": f"{pct(intact, len(paras))}%",
        **{k: f"{pct(v, len(found))}%" for k, v in agree.items()},
        "tables": f"{len(a.tables)} -> {len(b.tables)}",
        "pictures": f"{sum(len(p['images']) for p in a.paragraphs)} -> {sum(len(p['images']) for p in b.paragraphs)}",
        "misses": misses,
    }


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    use_word = "--no-word" not in sys.argv
    OUT.mkdir(exist_ok=True)
    rows, details, scored = [], [], []
    pages = {}
    pdfs = sorted((HERE / "corpus").glob("*.pdf"))
    for pdf in pdfs:
        out = OUT / f"{pdf.stem}.docx"
        r = subprocess.run(["node", str(CLI), str(pdf), str(out)], capture_output=True, text=True, encoding="utf-8")
        if r.returncode:
            rows.append(f"| {pdf.stem} | conversion failed: {r.stderr.strip().splitlines()[-1][:80]} |")
            continue
        scored.append((pdf, out, score(original_for(pdf), out)))
    # Page counts in one Word session for all files (several sessions in a row leave Word instances behind).
    if use_word and scored:
        sys.path.insert(0, str(ROOT / "experiments" / "e1-docx-roundtrip"))
        from word import Word
        with Word() as w:
            for pdf, out, _ in scored:
                try:
                    pages[pdf.stem] = (w.inspect(str(original_for(pdf)))["pages"], w.inspect(str(out))["pages"])
                except Exception as e:  # Word could not open our file
                    pages[pdf.stem] = ("?", f"Word error: {str(e)[:40]}")
    for pdf, out, s in scored:
        pg = f"{pages[pdf.stem][0]} -> {pages[pdf.stem][1]}" if pdf.stem in pages else "—"
        rows.append(f"| {pdf.stem} | {s['markers']} | {s['text']} | {s['paragraphs']} | {s['bold']} | {s['italic']} | {s['size']} | {s['role']} | {s['align']} | {s['tables']} | {s['pictures']} | {pg} |")
        if s["misses"]:
            details.append(f"### {pdf.stem}\n" + "\n".join(f"- {m}" for m in s["misses"]))
    report = ("# Experiment 4 — PDF → Word\n\n"
              "| Document | Markers | Text | Paragraphs intact | Bold | Italic | Size | Role | Align | Tables | Pictures | Pages (orig → ours) |\n"
              "|---|---|---|---|---|---|---|---|---|---|---|---|\n" + "\n".join(rows) + "\n\n## First mismatches\n\n" + "\n\n".join(details) + "\n")
    (HERE / "report.md").write_text(report, encoding="utf-8")
    print(report)


if __name__ == "__main__":
    main()
