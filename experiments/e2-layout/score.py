"""Experiment 2 scorer: our layout engine vs Microsoft Word, marker by marker.

For each fixture with a Word baseline: page count, % of markers on the same page as Word, and the
vertical error (pt) of markers that are on the same page. Writes report-<fixtures dir>.md here.
Usage: python score.py [fixtures_dir]
"""
import json
import statistics
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
CLI = ROOT / "packages" / "layout" / "dist" / "cli.js"
TOLERANCE = 2.0  # pt


def ours(path):
    out = subprocess.run(["node", str(CLI), "markers", str(path)], capture_output=True, text=True, check=True)
    return json.loads(out.stdout)


def main():
    fixtures = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "experiments" / "e1-docx-roundtrip" / "fixtures"
    lines = ["# Experiment 2 — layout vs Microsoft Word", "",
             f"Vertical tolerance: ±{TOLERANCE}pt. Position = top of a paragraph's first line.", "",
             "| Fixture | Pages (Word → ours) | Same page | Within ±2pt | Median error | Worst |",
             "|---|---|---|---|---|---|"]
    details = []
    for base_file in sorted((HERE / "baselines").glob("*.json")):
        docx = fixtures / f"{base_file.stem}.docx"
        if not docx.exists():
            continue
        word = json.loads(base_file.read_text(encoding="utf-8"))
        mine = ours(docx)
        rows = []
        for m, w in word["markers"].items():
            o = mine["markers"].get(m)
            rows.append((m, w, o))
        same = [(m, w, o) for m, w, o in rows if o and o["page"] == w["page"]]
        errors = [abs(o["y"] - w["y"]) for _, w, o in same]
        within = sum(e <= TOLERANCE for e in errors)
        worst = max(errors, default=0)
        n = len(rows)
        lines.append(f"| {base_file.stem} | {word['pages']} → {mine['pages']} | {len(same)}/{n} ({100 * len(same) / n:.0f}%) "
                     f"| {within}/{n} ({100 * within / n:.0f}%) | {statistics.median(errors) if errors else 0:.1f}pt | {worst:.1f}pt |")
        bad = [(m, w, o) for m, w, o in rows if not o or o["page"] != w["page"] or abs(o["y"] - w["y"]) > TOLERANCE]
        if bad:
            details.append(f"\n## {base_file.stem}: first mismatches\n")
            details.append("| Marker | Word (page, y) | Ours (page, y) |\n|---|---|---|")
            for m, w, o in bad[:12]:
                details.append(f"| {m} | {w['page']}, {w['y']} | {o['page'] if o else '—'}, {o['y'] if o else '—'} |")
    report = "\n".join(lines + details) + "\n"
    (HERE / f"report-{fixtures.name}.md").write_text(report, encoding="utf-8")
    sys.stdout.reconfigure(encoding="utf-8")
    print(report)


if __name__ == "__main__":
    main()
