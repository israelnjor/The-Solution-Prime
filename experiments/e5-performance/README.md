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

| | Before | After |
|---|---|---|
| Keystroke, median / worst | 659 ms / 2.6 s | **28 ms / 46 ms** |
| Long tasks while typing | ~250 (100–450 ms each) | **none** |
| After a pause | repeated 2 s freezes for seconds | re-layout and page corrections, 50–270 ms each, ~0.8 s in all |
| Page alignment | pages after ~56 misaligned (up to 4 pages off) | all within 1 px |

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

| 145 pages | Before | After |
|---|---|---|
| Applying a keystroke (`dispatch`) | ~31 ms | **~3.6 ms** |
| Fitting the pages, per keystroke (median) | — | 2 ms |
| Keystroke until painted, median | 63–75 ms | **43–45 ms** |
| Page alignment | within 1 px | within 1 px |

The rest of each keystroke is the browser laying out and painting a 145-page editable page (our
code now accounts for ~11 ms of the 44). Timings on this computer vary by ±20% between runs (and a
run is spoiled when a re-layout lands during typing), so compare three runs. Further gains would need
pages kept out of the browser's layout while off screen, a larger change.

### Later the same day: tab stops on screen

With tabs drawn at their stops (as the layout places them) instead of the browser's fixed tab width,
the screen and the pages agree on far more lines, so fitting has less to correct:

| 145 pages | Before | After |
|---|---|---|
| Keystroke until painted, median / worst | 43–45 ms / ~150 ms | **28–29 ms / 40 ms** |
| One-off fitting pass after a layout | ~425 ms | ~160 ms |
| Page alignment | within 1 px | within 1 px |
