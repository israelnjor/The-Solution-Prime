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
