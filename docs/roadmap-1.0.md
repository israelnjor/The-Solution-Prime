# Prime Word 1.0 — feature coverage and roadmap

Goal for 1.0: cover **70–80% of the Word features people use in everyday documents** (assignments,
reports, theses, letters, CVs, minutes) — not 80% of Word's ~1,500 commands, most of which almost
nobody uses. Coverage below is counted over the everyday features in each area.

Legend: ✅ done · 🟡 partial (shown/kept but not editable, or basic only) · ❌ missing

## Where we are (2026-11-13, after Waves 0–6 and a round of polish from real use): about 75–78%

The 1.0 target is reached. **Feature freeze for 1.0**: from here the work is shipping (see Release
below); a new feature goes in only if a test user is blocked without it.

| Area | Feature | Status |
|---|---|---|
| **File** | New from templates, Open .docx, Recent, Save / Save As, autosave + crash recovery | ✅ |
| | Export PDF, PDF → Word converter; Print dialog (pages, preview; printer, copies and two-sided in the system's window) | ✅ |
| | Open .doc, .rtf, .odt, .txt; save as .rtf, .odt, .txt (and .docx) | ✅ |
| | Document properties (File → Info: title, author, tags, statistics), password protection (Word's encryption: opens Word's, Word opens ours) | ✅ |
| | Share (Windows' share panel: a Word copy or a PDF, to WhatsApp, email, Nearby Share…), delete pages | ✅ |
| | Drop files on the window (documents open in a tab; a PDF goes to the converter, confirmed with its first page, pages and size) | ✅ |
| | Save as .doc, version history | ❌ |
| **Home: font** | Font, size, colour, bold/italic/underline/strike; highlight in any colour (Word's highlighter colours, any other as shading) | ✅ |
| | Superscript/subscript, grow/shrink font, change case, clear formatting, underline styles, character spacing, font dialog | ✅ |
| | Small caps (saved, not drawn), text effects (shadow, outline, glow) | ❌ |
| **Home: clipboard** | Cut/copy/paste (Word and web formatting kept) | ✅ |
| | Paste options (keep source, merge, text only), format painter, Ctrl+Shift+C / V | ✅ |
| **Home: paragraph** | Alignment, line/paragraph spacing, indents (buttons, Tab/Shift+Tab), bullets, numbering, list levels | ✅ |
| | Sub-lists (Bullets/Numbering in a list item; bullets under numbers and the reverse), bullet library (9 symbols), numbering library (7 styles); Enter keeps the paragraph's formatting, as Word does | ✅ |
| | Borders (basic) | 🟡 |
| | Show ¶ marks, ruler (indents, tab stops), tab stops on screen, line and page breaks, tabs dialog | ✅ |
| | Multilevel numbering (1, 1.1, 1.1.1), paragraph shading, sort | ❌ |
| **Home: styles** | Apply paragraph styles | ✅ |
| | Create / modify styles (Styles dialog), style sets, theme fonts and colours (Design tab) | ✅ |
| **Home: editing** | Find / Replace | ✅ |
| | Search in the navigation pane (Ctrl+F: results list, matches marked in the outline) | ✅ |
| | Find options (formatting, wildcards), Go To | ❌ |
| **Insert** | Page break, table, picture, link, header/footer, page number fields, comment | ✅ |
| | Equations: insert and edit (LaTeX notation, live preview, saved as Word equations) | ✅ |
| | Text boxes (insert, type in them, move, resize, fill and border), shapes (13 everyday shapes: insert, fill, outline, wrap) | ✅ |
| | Charts: column, bar, line, area, pie, doughnut, scatter, stacked; colour sets, legend, data labels, gridlines, title; a Chart Design tab; data typed, pasted from Excel, or taken from a table in the document; Word's charts shown and editable | ✅ |
| | Symbol (with ₵ and ₦), date & time (as text or a DATE field) | ✅ |
| | Section and column breaks, drop cap | ✅ |
| | SmartArt, WordArt, 3-D charts, shape effects | ❌ |
| **Design** | Style sets, theme fonts and colours, watermark, page colour, page borders | ✅ |
| **Layout** | Margins, orientation, size, indents/spacing | ✅ |
| | Floating objects: move/resize | ✅ |
| | Columns (any number, line between, balanced), line numbers, automatic hyphenation, sections with their own page size and orientation | ✅ |
| | Wrap-text options for text boxes, arrange | ❌ |
| **References** | Insert/update table of contents, footnotes and endnotes (laid out at the foot of the page / after the text, editable), captions, table of figures, cross-references, citations & bibliography (APA, Harvard, IEEE; Word's own sources), Update Field (F9) | ✅ |
| | Index, table of authorities, footnote options (numbering format, restart per page) | ❌ |
| **Review** | Spelling, word count, comments, track changes (record insertions, deletions, paragraph breaks and formatting; show/accept/reject) | ✅ |
| | Grammar (basic rules), thesaurus (WordNet, offline), compare documents (as tracked changes), read aloud (the computer's own voices) | ✅ |
| | Restrict editing, full grammar and style checking | ❌ |
| **View** | Print / Web layout, zoom, ruler, formatting marks, navigation pane (headings: jump, collapse, drag to move a section) | ✅ |
| | Read mode (one or two pages side by side), outline view (levels, promote/demote, show level, fold), One Page / Page Width zoom | ✅ |
| | Pages side by side while editing, vertical ruler, dark mode; on screen, a section with a different column count or page width shows its text in one column (it prints right) | ❌ |
| **Tables** | Insert, add/delete rows and columns, Tab navigation, merge/split cells, borders & shading per cell, table styles (Word's conditional formats: header/total row, banding, first/last column) with options, cell alignment, column resize by dragging, distribute/AutoFit, repeat header rows, sort, text↔table | ✅ |
| | Row height, cell margins, nested-table tools, formulas, draw table | ❌ |
| **Pictures** | Insert, inline/floating display, wrap options, drag anywhere, crop, rotate, alt text | ✅ |
| | Picture styles, free rotation, adjustments (brightness, colour) | ❌ |
| **Make it yours** | My templates, letterhead designer, signatures (photo/scan clean-up, draw, type), blocks (signature blocks, addresses, stamps, saved text; left/centre/right; cursor, header or footer), drag paragraphs (handle, Alt+Shift+↑/↓) | ✅ |
| **Everyday behaviour** | Undo/redo, keyboard shortcuts (Word's formatting keys too), accessibility basics | ✅ |
| | Right-click menu, AutoCorrect (own replacements), smart quotes, auto-lists ("1. " / "- "), status bar "Page X of Y" | ✅ |

## Where we are ahead (value propositions)

- **PDF → Word built in**, offline, free — Word needs Acrobat or online services for good results.
- **Offline and private**: nothing leaves the computer; no account, no subscription.
- **Fast on long documents** (145 pages: ~28 ms per keystroke until painted, no freezes while typing)
  and a small installer (6.3 MB, PDF converter and thesaurus included).
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
3. **Tables complete** ✅ (2026-11-04) — merge/split, borders and shading, table styles, cell alignment, column
   resizing, header rows, sort, text↔table, autofit.
4. **Layout and design** ✅ (2026-11-07) — section breaks (next page, continuous, even/odd page) and
   column breaks; sections with their own page size and orientation on screen (landscape pages in a
   portrait document); text columns (any number, equal or own widths, line between; balanced before a
   continuous break, as Word does); page colour, page borders, watermark; line numbers; automatic
   hyphenation; create/modify styles, style sets, theme fonts and colours; drop cap. Each checked
   against Word (same pages and positions; Word reads ours as its own). On screen, a document shows one
   column count (its first): a section with a different count shows as one column (it prints right).
5. **Review, view and files** ✅ (2026-11-08) — record tracked changes (insertions, deletions, paragraph
   breaks and formatting, by Word's rules); navigation pane (headings: jump, collapse, drag to move a
   section; Ctrl+F search with results); compare documents (the differences as tracked changes: accepting
   all gives the revised text, rejecting all the original); read mode (one or two pages side by side);
   outline view; One Page / Page Width zoom; Print dialog with preview; document properties; password
   protection (Word's Agile Encryption, AES-256: each opens the other's); open .doc, .rtf, .odt, .txt and
   save .rtf, .odt, .txt (checked: Word opens what we write, and we read what Word writes); thesaurus
   (WordNet 3.0, offline); read aloud (installed voices only); basic grammar. Not done: saving .doc
   (.docx is read everywhere), pages side by side while editing (read mode has them).
6. **Objects** ✅ (2026-11-10) — text boxes (insert; double-click to type, with formatting; fill and border),
   shapes (rectangle, oval, arrows, triangle, star, line…: insert, fill, outline, wrap), drawn on screen and in
   the PDF from Word's own description (theme colours included); charts (column, bar, line, pie) inserted
   from a small table and edited by double-click, saved as real Word charts with their spreadsheet (Word's
   Edit Data works), and Word's charts shown. Checked against Word (same shapes, colours, sizes; Word's
   charts drawn alike; Word reads ours as its own).
7. **Projects** — phase 1 ✅ (2026-11-10): a project is an ordinary folder (files stay where the person and
   their backups can see them; a hidden .primeword/project.json keeps name and colour and travels with it).
   Home shows project cards; a project page lists its documents, PDFs, pictures and other files; new
   documents and converted PDFs are saved straight into it; files can be copied in; tabs and the window
   title show the project. Next: version history per document (restore, compare) and zip backup (phase 2);
   notes, checklist, activity and search across a project (phase 3); project templates (phase 4).
   **Polish from real use** ✅ (2026-11-11 to 11-13) — share (Windows' share panel), delete pages; lists
   (Enter continues a list, sub-lists, bullet and numbering libraries, Tab and the indent buttons on any
   paragraph); highlight in any colour; charts completed (see Insert); format painter on a single click;
   files dropped on the window; the PDF converter confirms before converting. Fixed: typing after a page
   break (and between two breaks), lines overlapping after the text was made bigger, list formatting lost
   on screen, landscape sections in a portrait document (a new page, as Word does; margins kept while
   typing), drop caps beside the text where the page has text columns, the ribbon at half-screen width.
8. **Release** (next — feature freeze) — text recognition for PDFs without text ✅ (2026-11-10: scanned
   pages and pages printed as outlines, e.g. by Microsoft Print to PDF, read by Windows' own offline
   recogniser — about half a second a page; macOS and Linux not yet). To do, in order:
   1. Code signing (without it Windows SmartScreen warns people off the installer).
   2. Automatic updates (so fixes reach people after they install).
   3. About / version.
   4. User guide (short: the everyday tasks and what is ours — PDF → Word, letterheads, signatures).
   5. Bug-fix pass: real documents (reports, papers, company documents); every problem found is fixed
      and re-checked.
   6. Test users: 5–10 people from the target market (colleagues, researchers, small offices) on the
      installer before a wide release.

Later (1.x), in order of value to the target market: multilevel numbering (1, 1.1, 1.1.1), restrict
editing, fuller grammar and style checking, index and table of authorities, paragraph shading, Find
options and Go To, row height and cell margins, picture styles and adjustments, mailings and mail
merge, SmartArt, macros, version history (Projects phase 2), real-time collaboration, AI features.

## Business: users, licences and payments (decided 2026-11-13, details pending)

- **Who it is for:** professionals — workers, researchers, publishers and corporate staff (not mainly
  students).
- **Who pays:** individuals, and companies buying licences (seats) for their workers.
- **How:** yearly licences.
- **What is free and what is paid:** to be decided.
- **Proposed design** (to confirm before building):
  - No account to use Prime Word. A licence key carries its expiry date and a signature that the app
    checks offline. Renewal: the app fetches the renewed key when online, or the person pastes it; a
    grace period after expiry. When a licence lapses, documents stay openable, readable and exportable.
  - Companies: an admin page for the buyer (seats bought, who has one, assign / move / revoke), signed in
    with an emailed link. The only online service, and it holds licences, never documents.
  - Payments: a merchant of record (Paddle or Lemon Squeezy) for international buyers (VAT, yearly
    billing, invoices); Paystack for Ghana and Africa (MoMo, cards, cedis); bank transfer against an
    invoice for larger companies (by hand at first). Payout support and fees to be checked for each.
- **Waits for:** the company's registration (planned within 2026); the Microsoft Store account is then
  opened as the company.

Other systems (after 1.0, which is Windows only): **Linux** first (no signing cost; AppImage/.deb, Flathub),
then **macOS** (Apple Developer Program, $99/year, for signing and notarisation; built and tested on a Mac).
Both need: the share panel (Windows-only today), text recognition for scanned PDFs (Windows' recogniser
today; macOS has its own), metric-compatible fonts where Calibri/Cambria are missing (layout fidelity),
and a check of the system web view's differences.
