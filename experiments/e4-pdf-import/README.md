# Experiment 4 — PDF → Word

Prime Word's own PDF → Word conversion (ADR-009, `packages/pdf-import`): pdf.js (Mozilla, Apache-2.0)
reads the PDF; everything after that — lines, paragraphs, tables, lists, headings, headers and footers,
pictures, styles — is ours. No services, no subscriptions; it runs offline in the app.

## Answer key

`corpus/` holds PDFs that Word exported from the round-trip fixtures, so the original `.docx` is the
answer key. `score.py` converts each PDF and compares the result with the original:

```bash
npm run build                                   # from the repo root (builds packages/pdf-import)
python experiments/e4-pdf-import/score.py       # add --no-word to skip Word page counts
node packages/pdf-import/dist/cli.js in.pdf out.docx   # one file
```

## Results (2026-11-01)

| Document | Markers | Text | Paragraphs intact | Bold | Italic | Size | Role | Align | Tables | Pictures | Pages |
|---|---|---|---|---|---|---|---|---|---|---|---|
| A basic formatting | 34/34 | 100% | 94% | 100% | 100% | 100% | 94% | 100% | 0 → 0 | 0 → 0 | 2 → 2 |
| B tables, links, images | 44/44 | 98% | 100% | 100% | 100% | 100% | 100% | 98% | 3 → 3 | 2 → 2 | 2 → 2 |
| C multi-page, sections | 56/56 | 97% | 85% | 100% | 100% | 100% | 100% | 100% | 2 → 2 | 0 → 0 | 5 → 5 |
| D complex report | 385/387 | 99% | 100% | 100% | 100% | 100% | 100% | 99% | 2 → 2 | 1 → 1 | 6 → 6 |
| E footnotes, comments, changes | 13/13 | 74% | 71% | 100% | 100% | 100% | 100% | 69% | 0 → 0 | 1 → 1 | 1 → 2 |
| F pictures | 8/8 | 89% | 43%* | 100% | 100% | 100% | 100% | 100% | 1 → 1 | 5 → 5 | 2 → 2 |
| G equations | 5/5 | 70%† | 100% | 100% | 100% | 100% | 100% | 100% | 0 → 0 | 0 → 0 | 1 → 1 |

\* Pictures that sat inside a line of text come back on a line of their own. † Equations come back as
text (𝐴̅ + 𝐵 …), not as Word equations; the scorer counts the original's equations as empty text.

Recognised: fonts (incl. Light/Semibold), bold, italic, size, colour, underline, super/subscript,
links; paragraphs (joined across lines and pages), alignment, indents, spacing, line spacing; numbered
headings; bullet and numbered lists; Title and Heading 1–3 by size; ruled tables (merged cells, shading,
cell alignment, pictures in cells, tables that continue across pages, repeated header rows); headers
and footers (page numbers become PAGE fields; repeated logos); pictures; page size and margins.

## Added 2026-11-02 (real PDFs: a letter, a LaTeX slide deck, a Word report with contents)

- **Line breaks**: a line that ends early at normal line spacing becomes a line break (Shift+Enter) in
  the same paragraph — also for centred and right-aligned text.
- **Tables of contents**: three or more lines ending in a dot leader (or gap) and a page number become
  Word contents entries (TOC 1–3 by indent, right tab with dots) inside a TOC field; lists of figures
  become a table of figures (TOC \c "Figure"). Word recognises both and can update them.
- **LaTeX PDFs**: Computer/Latin Modern fonts mapped (italic and bold from the font name); accents
  composed (m¨ussen → müssen); blackboard bold (ℝ, ℂ, 𝔽); CMEX brackets; list labels keep their style
  and number ((a), (ii), 3)); dashes are no longer taken for bullets.
- **Colours and backgrounds**: text colour per run (pdf.js merges runs of different colours); filled
  boxes behind text (incl. rounded ones, ignoring see-through shadows) become paragraph shading.
- Underlines drawn in pieces; numeric list ids (a non-numeric id made Word reject the file); text never
  contains characters XML forbids.

The layout engine gained tab stops (right/centre/decimal, dot leaders), paragraph shading, and PAGEREF
page numbers computed from the layout (contents pages show current numbers).

## Not yet

- **Scanned pages** (no text, one full-page picture) are reported; text recognition (Tesseract) is next.
- Bullets drawn as shapes (slide decks); equations as Word equations (large operators break apart); footnotes; text beside pictures (pictures go on their own line);
  multi-column pages; tables without rules; paragraph borders and shading.
