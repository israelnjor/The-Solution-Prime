"""Experiment 1 runner: round-trip every fixture through every candidate and score the results.

For each candidate x fixture:
  1. run the candidate's command:  <cmd> <in.docx> <out.docx>
  2. fingerprint original and output, compare (compare.py)
  3. open both in Microsoft Word: page count + PDF render for visual review
Writes results/<timestamp>/report.md, report.json, and the outputs/PDFs.

Usage:
  python run.py                          # everything
  python run.py --candidates identity,html-pipeline --fixtures A,D
  python run.py --no-word                # skip Word page counts / PDFs
"""
import argparse
import json
import shutil
import subprocess
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path

from compare import AREAS, SEVERITY_ORDER, area_status, compare
from docxfp import Fingerprint

HERE = Path(__file__).resolve().parent
ICON = {"PASS": "✅", "WARN": "⚠️", "FAIL": "❌", "SKIP": "⏭️"}
GATE_QUESTIONS = [
    "Can the candidate preserve the document features in the agreed MVP?",
    "Can it provide acceptable page layout without building a pagination engine from scratch?",
    "Can we use it within a sustainable budget and licensing model?",
    "Can core editing remain offline?",
    "What happens to features it cannot represent faithfully — reported, or silently lost?",
]


def load_candidates(selected):
    cands = json.loads((HERE / "candidates.json").read_text(encoding="utf-8"))
    if selected:
        names = set(selected.split(","))
        cands = [c for c in cands if c["name"] in names]
    return cands


def load_fixtures(selected, fixtures_dir):
    paths = sorted(fixtures_dir.glob("*.docx"))
    if not paths and fixtures_dir == HERE / "fixtures":
        subprocess.run([sys.executable, str(HERE / "generate_fixtures.py")], check=True)
        paths = sorted((HERE / "fixtures").glob("*.docx"))
    if selected:
        prefixes = tuple(s.strip() for s in selected.split(","))
        paths = [p for p in paths if p.stem.startswith(prefixes)]
    return paths


def run_candidate(cand, src, dst, timeout):
    cmd = [part.replace("{python}", sys.executable).replace("{in}", str(src)).replace("{out}", str(dst))
           for part in cand["cmd"]]
    started = time.perf_counter()
    proc = subprocess.run(cmd, cwd=HERE, capture_output=True, text=True, timeout=timeout, shell=False)
    elapsed = time.perf_counter() - started
    return proc, elapsed


def evaluate(cand, fixture, out_dir, word, baseline, timeout):
    result = {"candidate": cand["name"], "fixture": fixture.stem, "findings": [], "metrics": {}}
    dst = out_dir / f"{fixture.stem}.docx"
    try:
        proc, elapsed = run_candidate(cand, fixture, dst, timeout)
        result["metrics"]["seconds"] = round(elapsed, 2)
        if proc.returncode != 0 or not dst.exists():
            result["findings"].append({"severity": "CRITICAL", "area": "Reliability",
                                       "message": f"candidate failed (exit {proc.returncode})",
                                       "examples": [(proc.stderr or proc.stdout)[-500:]]})
            return result
    except subprocess.TimeoutExpired:
        result["findings"].append({"severity": "CRITICAL", "area": "Reliability",
                                   "message": f"candidate timed out after {timeout}s", "examples": []})
        return result

    try:
        findings, metrics = compare(Fingerprint(fixture), Fingerprint(dst))
        result["findings"].extend(findings)
        result["metrics"].update(metrics)
    except Exception as e:  # unreadable output is itself a finding
        result["findings"].append({"severity": "CRITICAL", "area": "Reliability",
                                   "message": f"output is not a readable DOCX: {e}", "examples": []})
        return result

    if word:
        try:
            info = word.inspect(dst, out_dir / f"{fixture.stem}.pdf")
            result["metrics"]["pages"] = info["pages"]
            base = baseline[fixture.stem]["pages"]
            result["metrics"]["pages_original"] = base
            if info["pages"] != base:
                result["findings"].append({"severity": "MAJOR", "area": "Pagination",
                                           "message": f"page count in Word changed ({base} → {info['pages']})",
                                           "examples": []})
        except Exception as e:
            result["findings"].append({"severity": "CRITICAL", "area": "Reliability",
                                       "message": f"Microsoft Word could not open the output: {e}", "examples": []})

    warnings_file = Path(str(dst) + ".warnings.json")
    warnings = json.loads(warnings_file.read_text(encoding="utf-8")) if warnings_file.exists() else []
    result["candidate_warnings"] = warnings
    losses = [x for x in result["findings"] if x["severity"] in ("CRITICAL", "MAJOR")]
    if losses and not warnings:
        result["findings"].append({"severity": "MAJOR", "area": "Transparency",
                                   "message": f"{len(losses)} critical/major changes with no warning reported by the candidate",
                                   "examples": []})
    elif losses:
        result["findings"].append({"severity": "MINOR", "area": "Transparency",
                                   "message": f"{len(losses)} critical/major changes; candidate reported {len(warnings)} "
                                              "warnings — review whether they cover the losses", "examples": warnings})
    elif warnings:
        result["findings"].append({"severity": "MINOR", "area": "Transparency",
                                   "message": f"candidate reported {len(warnings)} warnings", "examples": warnings})
    return result


def write_report(run_dir, results, candidates, fixtures, baseline):
    by_key = {(r["candidate"], r["fixture"]): r for r in results}
    lines = [f"# Experiment 1 — DOCX round-trip report", "",
             f"Run: `{run_dir.name}` · fixtures: {len(fixtures)} · candidates: {len(candidates)}", "",
             "Legend: ✅ pass · ⚠️ minor drift · ❌ critical/major loss · ⏭️ skipped", "",
             "## Summary", ""]
    header = "| Candidate | Fixture | " + " | ".join(AREAS) + " | Words kept | Format identical | Pages (orig → out) |"
    lines += [header, "|" + "---|" * (len(AREAS) + 5)]
    for cand in candidates:
        for fx in fixtures:
            r = by_key.get((cand["name"], fx.stem))
            if r is None:
                continue
            if r.get("skipped"):
                lines.append(f"| {cand['name']} | {fx.stem} | " + " | ".join(ICON["SKIP"] for _ in AREAS)
                             + f" | — | — | {r['skipped']} |")
                continue
            st = area_status(r["findings"])
            m = r["metrics"]
            pages = f"{m['pages_original']} → {m['pages']}" if "pages" in m else "—"
            lines.append(f"| {cand['name']} | {fx.stem} | " + " | ".join(ICON[st[a]] for a in AREAS)
                         + f" | {m.get('words_preserved_pct', '—')}% | {m.get('formatting_identical_pct', '—')}% | {pages} |")
    lines += ["", "## Candidates", ""]
    for cand in candidates:
        lines.append(f"- **{cand['name']}** ({cand['role']}): {cand['description']}")
    lines += ["", "## Findings", ""]
    for cand in candidates:
        for fx in fixtures:
            r = by_key.get((cand["name"], fx.stem))
            if r is None or r.get("skipped"):
                continue
            lines.append(f"### {cand['name']} · {fx.stem}")
            lines.append("")
            if not r["findings"]:
                lines += ["No differences detected.", ""]
                continue
            for x in sorted(r["findings"], key=lambda x: SEVERITY_ORDER[x["severity"]]):
                lines.append(f"- **{x['severity']}** [{x['area']}] {x['message']}")
                for ex in x["examples"]:
                    lines.append(f"  - `{str(ex)[:200]}`")
            lines.append("")
    lines += ["## Decision gate (answer per candidate, from the evidence above)", ""]
    lines += [f"{i}. {q}" for i, q in enumerate(GATE_QUESTIONS, 1)]
    lines += ["", "Visual review: compare `baseline/*.pdf` with `<candidate>/*.pdf` side by side.", ""]
    (run_dir / "report.md").write_text("\n".join(lines), encoding="utf-8")
    (run_dir / "report.json").write_text(json.dumps({"baseline": baseline, "results": results}, indent=2,
                                                    ensure_ascii=False, default=str), encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(description="Experiment 1 DOCX round-trip runner")
    ap.add_argument("--candidates", help="comma-separated candidate names (default: all)")
    ap.add_argument("--fixtures", help="comma-separated fixture prefixes, e.g. A,D (default: all)")
    ap.add_argument("--no-word", action="store_true", help="skip Microsoft Word page counts and PDF renders")
    ap.add_argument("--timeout", type=int, default=300)
    ap.add_argument("--fixtures-dir", default=str(HERE / "fixtures"),
                    help="directory of input .docx files (e.g. fixtures-word, the Word-saved variants)")
    args = ap.parse_args()

    candidates = load_candidates(args.candidates)
    fixtures = load_fixtures(args.fixtures, Path(args.fixtures_dir).resolve())
    run_dir = HERE / "results" / datetime.now().strftime("%Y%m%d-%H%M%S")
    (run_dir / "baseline").mkdir(parents=True)

    word_ctx = None
    if not args.no_word:
        from word import Word
        word_ctx = Word()
    results, baseline = [], {}
    word = word_ctx.__enter__() if word_ctx else None
    try:
        for fx in fixtures:
            shutil.copy(fx, run_dir / "baseline" / fx.name)
            baseline[fx.stem] = word.inspect(fx, run_dir / "baseline" / f"{fx.stem}.pdf") if word else {}
            print(f"baseline {fx.stem}: {baseline[fx.stem]}")
        for cand in candidates:
            out_dir = run_dir / cand["name"]
            out_dir.mkdir()
            if cand.get("requires") and not (HERE / cand["requires"]).exists():
                for fx in fixtures:
                    results.append({"candidate": cand["name"], "fixture": fx.stem,
                                    "skipped": f"setup needed: {cand.get('setup', '')}"})
                print(f"skip {cand['name']}: run `{cand.get('setup')}`")
                continue
            for fx in fixtures:
                try:
                    r = evaluate(cand, fx, out_dir, word, baseline, args.timeout)
                except Exception:
                    r = {"candidate": cand["name"], "fixture": fx.stem, "metrics": {},
                         "findings": [{"severity": "CRITICAL", "area": "Reliability",
                                       "message": "runner error", "examples": [traceback.format_exc()[-500:]]}]}
                results.append(r)
                st = area_status(r["findings"])
                print(f"{cand['name']:>14} · {fx.stem:<22} " + " ".join(f"{a[:4]}:{st[a]}" for a in AREAS))
    finally:
        if word_ctx:
            word_ctx.__exit__(None, None, None)

    write_report(run_dir, results, candidates, fixtures, baseline)
    print(f"\nreport: {run_dir / 'report.md'}")


if __name__ == "__main__":
    main()
