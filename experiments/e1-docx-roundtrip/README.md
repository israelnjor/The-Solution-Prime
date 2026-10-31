# Experiment 1 — DOCX round-trip fidelity

Question: which editing engine can import → export our representative documents without losing
content, structure or layout? (ADR-002)

## How it works

1. `generate_fixtures.py` builds four deterministic DOCX files in `fixtures/`. Every paragraph,
   table cell and header carries a unique marker like `[A-012]`, so loss and duplication are exact.
   - **A** basic formatting · **B** tables/merges/links/images · **C** margins, sections, landscape,
     headers/footers, page numbers · **D** long report with tables crossing pages
2. Each candidate in `candidates.json` is a command: `<cmd> <in.docx> <out.docx>`. It may write
   `<out.docx>.warnings.json` (list of strings) to report what it could not preserve.
3. `run.py` round-trips every fixture through every candidate, then:
   - fingerprints both files (`docxfp.py`: text, effective formatting per character, lists,
     tables, merges, images, links, sections, headers/footers, fields)
   - scores the differences (`compare.py`) against the rubric: Content, Formatting, Structure,
     Pagination, Reliability, Transparency
   - opens both in Microsoft Word for page counts and PDF renders (`word.py`)
4. Output: `results/<timestamp>/report.md` (+ `report.json`, outputs, PDFs for visual review).

Severity: **CRITICAL** content lost/changed · **MAJOR** visible structure/layout loss ·
**MINOR** formatting drift. Losses with no candidate warning are flagged as *silent loss*.

```powershell
..\..\.venv\Scripts\python run.py                                   # all
..\..\.venv\Scripts\python run.py --candidates html-pipeline --fixtures C,D
```

## Adding a candidate

Write a script that converts `in.docx → [your editor's model] → out.docx`, add it to
`candidates.json`, and run. Controls (`identity`, `python-docx`) must stay all-✅; if they don't,
the checker is wrong, not the candidate.

## Fixture sets

| Directory | Made by | Purpose |
|---|---|---|
| `fixtures/` | `generate_fixtures.py` (deterministic) | A–D, the core feature corpus |
| `fixtures-word/` | `make_word_fixtures.py` (needs Word) | A–D re-saved by Word: realistic, messier XML |
| `fixtures-preservation/` | `make_preservation_fixture.py` (needs Word) | E: unsupported features that must survive untouched (footnotes, comments, tracked changes, content control, TOC, equation, bookmark link, floating image) |

Select one with `--fixtures-dir`, e.g. `run.py --candidates prime-docx --fixtures-dir fixtures-word`.

## Phase 1 result — `prime-docx` (2026-10-31)

Our own engine (`packages/docx`) passes the Experiment 1 gate on every fixture set:

| Fixtures | Result |
|---|---|
| A–D (generated) | All ✅ — 100% words, 100% identical formatting, Word page counts 2/2/5/6 unchanged |
| A–D (Word-saved) | All ✅ |
| E (preservation) | ✅ everything, ⚠️ Transparency: each unsupported feature reported by name (footnote, comment, tracked insertion/deletion, content control, equation, floating image) |

The output goes through a JSON snapshot of the model before writing, so the model itself is
proven complete — not just the XML pass-through.

**Known limits** (not yet covered by fixtures): footnote/comment *content* is carried as whole
parts and is not fingerprinted; revision ids (`w:rsid*`, `w14:paraId`) are intentionally
dropped. Inline pictures are regenerated only when they are plain (image, size, name, alt text);
pictures with borders, effects, crops or links are preserved verbatim and reported. Fixture F
(`fixtures-preservation/F_pictures.docx`) covers this, including a header logo, and the checker
fingerprints picture decoration and header pictures, so such a loss cannot go unnoticed.

## Earlier results (2026-10-31)

| Candidate | Result |
|---|---|
| identity, python-docx (controls) | All ✅ on all fixtures — the checker reports no false differences. |
| word-resave (reference) | All ✅ except one extra empty paragraph in C. |
| html-pipeline (mammoth → html-to-docx) | ❌ Loses headers/footers, sections, landscape page, first-page header, margins and page breaks (C: 5 → 3 pages); loses repeating table headers, alignment, colors, fonts, highlight and strikethrough; corrupts a hyperlink URL (`&` double-escaped) and writes an invalid `gridSpan=0`. |

Takeaway: an HTML-based import/export path cannot meet the MVP. The editor engine needs a
DOCX-aware model with page and section semantics. That is the bar for the next candidates.
