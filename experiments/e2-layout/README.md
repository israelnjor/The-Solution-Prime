# Experiment 2 — Layout engine vs Microsoft Word

Question (ADR-007): can our own layout engine paginate documents the way Word does?

## Method

1. `word_baseline.py` asks Word where every content marker lands — page number and vertical
   position of the paragraph's first line — and stores it in `baselines/`.
2. `@prime/layout` lays out the same `.docx` (via our own DOCX reader) using real font metrics.
3. `score.py` compares marker by marker: same page? within ±2pt?

```powershell
npm run build
..\..\.venv\Scripts\python word_baseline.py ..\e1-docx-roundtrip\fixtures      # needs Word
..\..\.venv\Scripts\python score.py ..\e1-docx-roundtrip\fixtures               # -> report-fixtures.md
```

## Results (2026-10-31)

Calibrated on A–D, then tested on E–F, which were **not** used for tuning (made natively by
modern Word: different compatibility mode, Aptos font, header logo, pictures, TOC).

| Fixture | Pages (Word → ours) | Same page | Within ±2pt | Worst |
|---|---|---|---|---|
| A basic formatting | 2 → 2 | 100% | 100% | 0.4pt |
| B tables, links, images | 2 → 2 | 100% | 100% | 1.8pt |
| C sections, landscape | 5 → 5 | 100% | 100% | 0.8pt |
| D 6-page report | 6 → 6 | 100% | 99% (385/387) | 2.4pt |
| **E** (held out) | 1 → 1 | 93% | 64% | content control missing, see below |
| **F** (held out) | 2 → 2 | 100% | 88% (7/8) | 2.8pt |

## Word behaviours discovered and implemented

Each was found from a measured mismatch, not assumed:

- Single line height = Windows ascent + descent, **except** fonts flagged `USE_TYPO_METRICS`
  (e.g. Aptos), which use typographic metrics.
- Adjacent paragraph spacing collapses to **max(space after, space before)**.
- Space-before is dropped at the top of a page after a page break, but kept (collapsed) at the
  start of a new section and on the first page.
- A paragraph ending in a page break leaves its space-after on the previous page.
- Table-style paragraph properties apply to cell paragraphs (defaults → table style → paragraph style).
- Widow/orphan control is on unless a document turns it off.
- Table rows split across pages at line boundaries unless "can't split"; header rows repeat.
- Vertically merged cells stretch the last row they span.
- Picture lines: picture + descent + line-spacing extra (not scaled); borders add the effect extent.
- Missing glyphs fall back to Segoe UI Symbol (symbols) and Yu Gothic (CJK) — matching Word here.
- Office cloud fonts (Aptos etc.) live in `%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts`.
- Headers/footers taller than their margin push the body.

Later additions, each from a measured mismatch:

- Content controls and tracked changes are model containers now, so their text is laid out
  (deleted text included, as in Word's markup view) — E went from 64% to 100% within ±2pt.
- Widow/orphan control works between manual page breaks, not across them; a picture line fits if
  the picture fits (its descent may spill into the margin). Found on a real 9-page assignment
  (`compare_real.py`, which matches paragraphs by text): 11 pages → 9, same as Word.
- Older documents (compatibility mode < 15) outdent tables by the cell margin; centred/right tables
  use the slack.

## Any real document

`compare_real.py <file.docx>` needs no markers: Word and our engine list every paragraph, they are
matched by text, and the report shows where the two start to diverge.

## PDF export

`node packages/pdf/dist/cli.js in.docx out.pdf` draws the layout engine's pages (alignment,
justification, colours, underline/strike/highlight, super/subscript, pictures, table borders and
shading, paragraph borders, headers/footers with live page numbers). Visual comparison against
Word's own PDF of fixtures A–C: near-identical pages.

## Equations

`packages/layout/src/math.ts` typesets Office Math (OMML) in Cambria Math: runs (math italics,
operator spacing, unary minus), accents and overlines, scripts, fractions, delimiters, radicals with
degrees, n-ary operators (∑ ∫ with limits), functions, limits, equation arrays and matrices; display
equations are centred. Checked visually against Word's own PDF on a real assignment (37 equations)
and on fixture G (`make_math_fixture.py`: quadratic formula, sum with limits, matrix, radical,
integral) — near-identical. Equations stay preserved verbatim in the DOCX model; editing them is
a later phase.

Remaining math differences: matrix brackets are scaled glyphs (Word uses stretched variants, so
ours look heavier); display equations get slightly more vertical space; lines with overlines are
~0.4pt shorter than Word's.

## Known gaps (next)

- Equation editing (display is done; the model still holds equations as preserved OMML).
- Not yet: custom tab stops, floating images/text wrapping, footnote areas, columns, hyphenation,
  even/odd headers, picture borders in PDF, EMF/WMF/GIF pictures in PDF (drawn as boxes).
- Fonts come from the machine. A document using a font that is not installed lays out with a
  substitute — the product will need to warn about that.

## Findings, 2026-11-01 (a real 11-page Times New Roman document)

Measured against Word with `compare_real.py` and small Word experiments (bullet lists in four fonts):

- **Line height includes the font's external leading** (Windows/GDI: the hhea line gap beyond the
  Windows ascent + descent). Times New Roman and Arial 12pt lines are 13.8pt in Word; Calibri and
  Cambria have none. Not for fallback fonts (a line with a few CJK characters is not spaced by Yu Gothic's gap).
- **A list label adds its ascent only**: Symbol bullet + Times New Roman 12pt = 12.06 (Symbol ascent)
  + 2.60 (Times descent) = 14.66pt. Line spacing (e.g. 1.16) scales the text's height; the label's
  extra is added unscaled.
- **`w:hideMark`**: an empty cell's paragraph mark adds nothing to the row height.
- Word breaks lines after hyphens inside words (never after "/").
- Word reports positions (`Range.Information`) on a 0.3pt grid: average pitches over many lines.

Result for that document: 17 → 186 of 199 paragraphs within 2pt of Word, all on the same page, 11 pages
like Word (was 12). Fixtures A–F unchanged or better.
