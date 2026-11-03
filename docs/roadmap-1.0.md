# Prime Word 1.0 — feature coverage and roadmap

Goal for 1.0: cover **70–80% of the Word features people use in everyday documents** (assignments,
reports, theses, letters, CVs, minutes) — not 80% of Word's ~1,500 commands, most of which almost
nobody uses. Coverage below is counted over the everyday features in each area.

Legend: ✅ done · 🟡 partial (shown/kept but not editable, or basic only) · ❌ missing

## Where we are (2026-11-03, after Waves 0–2): about 55%

| Area | Feature | Status |
|---|---|---|
| **File** | New from templates, Open .docx, Recent, Save / Save As, autosave + crash recovery | ✅ |
| | Export PDF, Print (via PDF), PDF → Word converter | ✅ |
| | Print dialog (printer, pages, copies), print preview | ❌ |
| | Open/save .doc, .rtf, .odt, .txt | ❌ |
| | Document properties, password protection, version history | ❌ |
| **Home: font** | Font, size, colour, highlight, bold/italic/underline/strike | ✅ |
| | Superscript/subscript, grow/shrink font, change case, clear formatting, underline styles, character spacing, font dialog | ✅ |
| | Small caps (saved, not drawn), text effects (shadow, outline, glow) | ❌ |
| **Home: clipboard** | Cut/copy/paste (Word and web formatting kept) | ✅ |
| | Paste options (keep source, merge, text only), format painter, Ctrl+Shift+C / V | ✅ |
| **Home: paragraph** | Alignment, line/paragraph spacing, indents, bullets, numbering, list levels | ✅ |
| | Borders (basic) | 🟡 |
| | Show ¶ marks, ruler (indents, tab stops), tab stops on screen, line and page breaks, tabs dialog | ✅ |
| | Multilevel list library, shading, sort | ❌ |
| **Home: styles** | Apply paragraph styles | ✅ |
| | Create / modify styles, style gallery | ❌ |
| **Home: editing** | Find / Replace | ✅ |
| | Find options (formatting, wildcards), Go To, navigation pane | ❌ |
| **Insert** | Page break, table, picture, link, header/footer, page number fields, comment | ✅ |
| | Equations: insert and edit (LaTeX notation, live preview, saved as Word equations) | ✅ |
| | Text boxes (shown, moved; not editable) | 🟡 |
| | Symbol (with ₵ and ₦), date & time (as text or a DATE field) | ✅ |
| | Section/column breaks, shapes, charts, SmartArt, WordArt, drop cap, bookmarks, cross-references | ❌ |
| **Design** | Themes, style sets, watermark, page colour, page borders | ❌ |
| **Layout** | Margins, orientation, size, indents/spacing | ✅ |
| | Floating objects: move/resize | ✅ |
| | Columns, line numbers, hyphenation, wrap-text options, arrange | ❌ |
| **References** | Insert/update table of contents, footnotes and endnotes (laid out at the foot of the page / after the text, editable), captions, table of figures, cross-references, citations & bibliography (APA, Harvard, IEEE; Word's own sources), Update Field (F9) | ✅ |
| | Index, table of authorities, footnote options (numbering format, restart per page) | ❌ |
| **Review** | Spelling, word count, comments, show/accept/reject tracked changes | ✅ |
| | Record tracked changes, grammar, thesaurus, compare, read aloud, protect | ❌ |
| **View** | Print / Web layout, zoom, ruler, formatting marks | ✅ |
| | Read mode, outline, navigation pane, vertical ruler, multiple pages, dark mode | ❌ |
| **Tables** | Insert, add/delete rows and columns, Tab navigation | ✅ |
| | Merge/split cells, borders & shading, table styles, cell alignment, column resize, header row repeat, sort, text↔table | ❌ |
| **Pictures** | Insert, inline/floating display, wrap options, drag anywhere, crop, rotate, alt text | ✅ |
| | Picture styles, free rotation, adjustments (brightness, colour) | ❌ |
| **Make it yours** | My templates, letterhead designer, signatures (photo/scan clean-up, draw, type), blocks (signature blocks, addresses, stamps, saved text; left/centre/right; cursor, header or footer), drag paragraphs (handle, Alt+Shift+↑/↓) | ✅ |
| **Everyday behaviour** | Undo/redo, keyboard shortcuts (Word's formatting keys too), accessibility basics | ✅ |
| | Right-click menu, AutoCorrect (own replacements), smart quotes, auto-lists ("1. " / "- "), status bar "Page X of Y" | ✅ |

## Where we are ahead (value propositions)

- **PDF → Word built in**, offline, free — Word needs Acrobat or online services for good results.
- **Offline and private**: nothing leaves the computer; no account, no subscription.
- **Fast on long documents** (145 pages: ~28 ms per keystroke until painted, no freezes while typing)
  and a small installer (5.6 MB, PDF converter included).
- **Faithful to Word's layout**: same pages, same line positions (measured against Word itself).
- **Make it yours**: design a letterhead in a minute; turn a phone photo of a signature into a clean
  e-signature; keep signature blocks, addresses and company stamps and insert them left, centre or
  right with one click — Word has no letterhead designer, no signature clean-up and no stamp maker.
- **Ambient AI** (planned, after 1.0): help that works offline where possible.

## Roadmap to ~75% (in order)

Each wave is checked against Word (layout comparison, round trip, Word reopening our files).

0. **Make it yours** ✅ (2026-11-03) — pictures that move freely (wrap, drag, crop, rotate, alt
   text); drag paragraphs; my templates and the letterhead designer; signatures; blocks and stamps.

1. **Feels like Word** ✅ (2026-11-03) — right-click menu; format painter; clear formatting; change case; grow/shrink
   font; superscript/subscript; paste options; show ¶ marks; ruler (indents, tab stops); AutoCorrect,
   smart quotes, automatic lists; symbol; date & time; "Page X of Y" in the status bar; font and
   paragraph dialogs complete.
2. **References** ✅ (2026-11-03) (the student/academic market) — insert and update table of contents; footnotes and
   endnotes; captions and table of figures; cross-references; citations and bibliography (APA,
   Harvard, IEEE); equation editor.
3. **Tables complete** — merge/split, borders and shading, table styles, cell alignment, column
   resizing, header rows, sort, text↔table, autofit.
4. **Layout and design** — section and column breaks, columns, page borders, watermark, page colour,
   line numbers, hyphenation, create/modify styles, themes and style sets (basic), drop cap.
5. **Review, view and files** — record tracked changes, compare documents, navigation pane, read
   mode, outline view, multiple pages, print dialog, document properties, password protection,
   .doc/.rtf/.odt/.txt open and save, thesaurus, read aloud, grammar (basic).
6. **Objects** — shapes and text boxes (insert and edit), simple charts.
7. **Release** — text recognition for scanned PDFs, code signing, automatic updates, About/version,
   user guide.

Later (1.x): mailings and mail merge, SmartArt, index, macros, collaboration, AI features.
