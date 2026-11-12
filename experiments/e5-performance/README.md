# Experiment 5 — Long-document performance

Prime Word must stay responsive in long documents (requirements: *ordinary editing stays responsive in
long documents*). This measures typing in a ~145-page document in the real editor (Print Layout).

```bash
node make-long.mjs long.docx 18     # 1,512 blocks, ~145 pages (fixtures D + A + B repeated)
npm run dev                         # from the repo root, in another terminal
./run-selftest.sh                   # T=perf (typing) — or T=align F=<docx> for page alignment only
```

`perf` types 40 characters in the middle of the document at ~16 characters a second, then waits 4 s.
It reports keystroke times (until the browser has painted), main-thread tasks over 50 ms ("long
tasks": the moments typing would stick), and time per step (`pw:` performance entries, recorded only
in self-test runs). `align` checks that every page's content starts at its sheet's body top.

## Results (2026-11-01, 145 pages)

|                           | Before                                         | After                                                          |
| ------------------------- | ---------------------------------------------- | -------------------------------------------------------------- |
| Keystroke, median / worst | 659 ms / 2.6 s                                 | **28 ms / 46 ms**                                        |
| Long tasks while typing   | ~250 (100–450 ms each)                        | **none**                                                 |
| After a pause             | repeated 2 s freezes for seconds               | re-layout and page corrections, 50–270 ms each, ~0.8 s in all |
| Page alignment            | pages after ~56 misaligned (up to 4 pages off) | all within 1 px                                                |

Causes fixed:

1. **Page-gap fitting** re-measured the whole document once per page (a style change between every
   read). Now one measuring pass; later positions follow arithmetically from earlier sizes.
2. **Paragraph formatting** was rebuilt for all 1,500 paragraphs on every keystroke; text-only edits now
   just map the existing formatting.
3. **Overflow repair** searched the whole document for every overflowing page, and retried pages that
   cannot be fixed (one block taller than the page) on every keystroke. Now it looks only at that page,
   and remembers unfixable pages until the next layout.
4. **Corrections** (breaks moved earlier where the screen runs longer than the layout) were capped at
   50 per layout and made one per redraw — long documents ran out, and pages jumped for seconds. Now
   all breaks move in one round per redraw, up to 10 rounds.
5. React redrew all page sheets on every keystroke; they are now memoised, as is the word count.

## Still open

- After each pause, the layout engine re-lays out the whole document (~110–180 ms at 145 pages,
  growing with length) and the corrections are found again from scratch (several rounds). Next:
  re-use corrections across layouts, and lay out incrementally from the edited page.
- In this generated document the screen runs longer than the layout on almost every page (140
  corrections for 145 pages), against 0–1 per fixture on its own: on screen, page breaks come earlier
  than in the PDF. To investigate next — the fixtures' content in another document's styles is
  probably hitting a difference between the browser's line breaking and the engine's.

## Regression found and fixed (2026-11-03)

Typing had slowed to a median of ~70 ms per keystroke (from 28 ms). Switching plugins off one at a
time pointed at the paragraph-formatting plugin, and inside it at two loops in ProseMirror that are
cheap only when there are few decorations at the top level of the document. It had one on every
paragraph (1,476 here):

1. Mapping a decoration set re-checks each node decoration with `content.findIndex`, a walk from the
   first child: ~1,500 × 1,500 steps per keystroke (4.7 ms in Node, ~13 ms in the browser).
2. Updating the view looks up each top-level node's decorations by scanning all top-level ones.

Paragraphs now show their own formatting (a node view); the plugin only keeps list labels and floats
as decorations (378 in all, none at the top level), and restyles paragraphs after a rebuild.
List-label widgets were also keyed by position, so every label after the cursor was redrawn on each
keystroke; they are keyed by their content now.

| 145 pages                                 | Before      | After               |
| ----------------------------------------- | ----------- | ------------------- |
| Applying a keystroke (`dispatch`)       | ~31 ms      | **~3.6 ms**   |
| Fitting the pages, per keystroke (median) | —          | 2 ms                |
| Keystroke until painted, median           | 63–75 ms   | **43–45 ms** |
| Page alignment                            | within 1 px | within 1 px         |

The rest of each keystroke is the browser laying out and painting a 145-page editable page (our
code now accounts for ~11 ms of the 44). Timings on this computer vary by ±20% between runs (and a
run is spoiled when a re-layout lands during typing), so compare three runs. Further gains would need
pages kept out of the browser's layout while off screen, a larger change.

### Later the same day: tab stops on screen

With tabs drawn at their stops (as the layout places them) instead of the browser's fixed tab width,
the screen and the pages agree on far more lines, so fitting has less to correct:

| 145 pages                               | Before              | After                       |
| --------------------------------------- | ------------------- | --------------------------- |
| Keystroke until painted, median / worst | 43–45 ms / ~150 ms | **28–29 ms / 40 ms** |
| One-off fitting pass after a layout     | ~425 ms             | ~160 ms                     |
| Page alignment                          | within 1 px         | within 1 px                 |

## A 480-page converted book (2026-11-04)

A 13 MB PDF book (474 pages, 531 pictures) converted to 7,471 blocks. Typing had slowed to 71 ms a
keystroke, and every pause froze the app for about a second. Four causes, each now fixed:

1. **Whole-document work on every keystroke**: ProseMirror's DecorationSet.forChild scanned all
   top-level decorations (the 481 page gaps) and all decorated children (paragraphs with spelling
   marks) for every top-level node; the word count and the tab-stop pass walked the whole document.
   forChild now finds a child by binary search (same results); the word count is cached per block;
   tab stops are sized only where the document changed.
2. **The layout measured every paragraph again after each pause.** Measured lines are now kept per
   paragraph across layouts (the editor gives an unchanged paragraph the same model object), keyed by
   everything else they depend on (width, styles, lists, table formatting, list and note numbers).
3. **Spell checking re-read the whole document after each pause**, and forgot the words it had
   checked. It now re-checks only the paragraphs that changed.
4. **Pictures**: converted pictures were all RGBA PNG without filtering (68 MB for this book), and in
   the app they were not converted at all (pdf.js hands over bitmaps there). Now: photographs as JPEG,
   others as filtered PNG, at most 200 dpi at their printed size — 19.9 MB, all 531 kept.

| 480-page book                   | Before                                  | After               |
| ------------------------------- | --------------------------------------- | ------------------- |
| Applying a keystroke            | ~21 ms                                  | **~3.7 ms**   |
| Keystroke until painted, median | 71 ms                                   | **50–53 ms** |
| Layout after a pause            | ~584 ms                                 | **~180 ms**   |
| Spell check after a pause       | ~282 ms                                 | **3 ms**      |
| Pictures in the converted file  | 68 MB (none, when converted in the app) | **19.9 MB**   |

The 145-page document gained too: 18–19 ms a keystroke (was 26–29). What remains on the book is the
browser repainting a 480-page editable page; keeping off-screen pages out of its layout is the next step.

## A 944-page converted book: page breaks that settle (2026-11-04)

An 8 MB PDF textbook converted to 9,778 blocks (506 tables), 855 pages. Typing took over a second a
keystroke, and the pages never settled. Screen and layout now agree closely on line heights (runs no
taller than the text set their own line; superscripts at 65 %; tab boxes no longer raise lines), and
the screen's corrections to the page breaks ("fitting", paged.ts) were reworked:

1. **Corrections cascaded.** A table row 2 px over moved the whole table; a paragraph that ran over
   moved whole; the heading that had started a page got its space-before back. Each left part of a
   page empty or pushed more onto the next page than planned, so nearly every later page overflowed
   in turn, round after round, until the 10-round limit left pages overflowing. Now a break moves
   before the crossing row or the crossing *line* (found from the text's own rectangles, not
   hit-testing, which only works on screen), the space a page-starting paragraph gains or loses is
   planned for, and finding the break a page already has is not a correction.
2. **Every layout started all corrections again**, and a long fit blocking the page made the next
   layout fire mid-sentence. Now a page keeps its correction when the layout put its start (and the
   previous page's) where it was; the next layout waits for a real pause after our own work, longer
   when layouts are slow.

| 855-page book                           | Before                    | After                           |
| --------------------------------------- | ------------------------- | ------------------------------- |
| Keystroke until painted, median / worst | 1,043 ms / seconds        | **120–133 ms / ~190 ms** |
| Layouts while typing 40 characters      | 22                        | **1**                     |
| Corrections when opened (rounds)        | 726 (never settled)       | **680 in 1 round**        |
| Pages off their sheet afterwards, worst | 823 pages, hundreds of px | **within 2.5 px**         |

Corrections on the generated 145-page document fell from 455 to 56 (it is back to 20 ms a keystroke).

### The same book, the next night: where each keystroke goes

Measured with the DevTools protocol (a JS CPU profile, then a full browser trace) while typing:

| Cause                                                                                                     | Fix                                                                                                           | Book, per keystroke                                  |
| --------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------- |
| React measured the selection across the whole focused editor on every render (status bar, ribbon)         | in documents over 300,000 characters, the ribbon and status bar update after a 200 ms pause in typing         | −30 ms                                              |
| ProseMirror's DecorationSet.create scanned every decoration once per top-level block (9,778 × thousands) | a linear, sorted build at the top level (render.ts)                                                           | ~1 s off each rebuild after a pause                  |
| Fitting measured all 855 page spacers after every keystroke                                               | after an edit, only the first spacer after it is resized (fastFit); a full fit only if the page now overflows | −25 ms                                              |
| 855 page sheets, masks and header/footer boxes, all positioned                                            | only the pages near the screen are drawn (PageSheets)                                                         | no change measured (kept: less work while scrolling) |
| Superscripts as `position:relative` (a layer each)                                                      | `vertical-align` (runs have no line height of their own)                                                    | no change measured (kept)                            |

Result: the book went from ~120 to **~80–90 ms** a keystroke (windowed Edge with GPU, as in the app: 88 ms;
the 145-page document: 17–21 ms). Tried and dropped: `content-visibility: auto` on the editor's blocks
(made both documents *slower*: the browser checks every skippable block's visibility each frame).

What remains is the browser's own per-frame work, which grows with the document: after an edit, it
re-checks the position of every element below it (~90,000 in this book; commit ~39 ms, compositing
inputs ~14 ms a frame with GPU). Going further means the editor no longer holding the whole document in
one page of the browser (e.g. one editing surface per section or chapter, or drawing only the pages on
screen) — a structural change.
