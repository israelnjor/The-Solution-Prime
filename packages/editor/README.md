# @prime/editor

The editing core of Prime Word (ADR-005): ProseMirror, used directly, with a schema derived from the
Prime Document Model.

- `schema.ts` — one node or mark per model construct. Properties travel as attributes, verbatim.
  Content the editor does not interpret (equations, floating drawings, bookmarks…) is an atom node
  carrying its preserved XML, so it is displayed and saved back unchanged.
- `convert.ts` — model ⇄ editor document, lossless. Container marks (tracked change, content
  control, link, field) are rebuilt into nesting on the way back.
- `commands.ts` — formatting edits change the model's own properties (run props, paragraph props).
- `render.ts` — browser rendering. Paragraph styling and list numbers come from the layout
  engine's style resolver in one document-order pass; equations are drawn by the math typesetter.
- `blank.ts` — a new document built from the model alone.

## Gate E3 — opening a document must not change it

`node dist/cli.js roundtrip in.docx out.docx` runs DOCX → model → editor (serialised to JSON and
back) → model → DOCX. It is the `prime-editor` candidate in Experiment 1. Result (2026-10-31): all
✅ on fixtures A–D (generated and Word-saved), E, F, G and a real 9-page assignment.

`test/edit-roundtrip.mjs` applies real edits (typing, bold, style, alignment) and saves; Word
confirms the result.

## Print Layout (Phase 4)

`paged.ts` + `pages.ts`: the editor keeps one flowing document; the **layout engine** decides where
pages start (it reports each page start as a paragraph path + character offset, or a table row), and
page-gap spacers are inserted there. After each redraw the spacers are sized from real screen
positions (via a stylesheet outside the editor — changing the editor's DOM would make ProseMirror
redraw it), so each page's content starts exactly at the top of its sheet's body area. If the
screen's content would overflow a page, that page's break moves to the last row/paragraph that
fits: the engine gives the plan, the screen has the last word, and a page never overflows.

The app draws page sheets behind the editor, margin masks above it (hiding e.g. table borders
running through the space between pages; skipped where content overflowed, so text is never
hidden) and headers/footers with live page numbers on top — all from the engine's layout.

Re-layout runs 250ms after typing stops: ~16ms for a 6-page report (Node, warm), ~3ms/page.
`test/page-starts.mjs` checks that engine page starts land on the right editor text.

## Known gaps

- Print Layout: a table row the engine splits across pages moves whole onto the next page on
  screen; repeated table header rows are not shown on continuation pages; a paragraph whose page
  break falls mid-paragraph re-applies its first-line indent after the break; sections with
  different page sizes are shown at the first page's size (with a notice).
- Fonts the browser cannot see (Office cloud fonts such as Aptos) display with a substitute; the
  desktop shell will load them from Office's font cache.
- Vertically merged table cells show as separate cells; table structure cannot be edited yet.
- Lists: numbering is displayed; creating or changing lists is not available yet.
- Autosave and crash recovery (ADR-010) are not wired into the app yet.
