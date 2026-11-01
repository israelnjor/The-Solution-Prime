# Prime Productivity

A professional productivity suite. The first product is **Prime Word**: a Windows-first,
offline-first word processor with high DOCX compatibility, PDF→DOCX conversion, and subtle,
optional AI assistance.

**Current phase:** Phase 4 — paginated editing. Prime Word is a Windows desktop app (Tauri) that
opens, edits and saves `.docx` files in Print Layout: real pages, laid out by our own engine, with
headers, footers and page numbers; Web Layout and zoom (status bar) are available too.

## Run Prime Word

Desktop app (needs Rust — installed with the minimal rustup profile):

```powershell
npm install
npm run build                          # engine packages
cd apps\word
npm run desktop                        # development window, reloads on code changes
npm run desktop:build                  # release build -> src-tauri\target\release\prime-word.exe
```

Browser (no Rust needed): `npm run dev` from the repo root, then open http://localhost:5173.

**Open** (Ctrl+O) and **Save** (Ctrl+S) use native Windows dialogs; Save writes back to the same
file. To rebuild the desktop app while Prime Word is open, use `scripts\build-desktop.ps1` (it moves
the running .exe aside; restart Prime Word to get the new build).

**Installer:** `scripts\build-installer.ps1` builds
`apps\word\src-tauri\target\release\bundle\nsis\Prime Word_<version>_x64-setup.exe`. It installs for
the current user (no administrator prompt), with a Start menu entry under *Prime Productivity*,
*Open with → Prime Word* for `.docx` files and an uninstaller in Settings → Apps. The first build
downloads the NSIS tools from Tauri's official releases; Tauri checks their hashes.

Editing (ribbon tabs Home / Insert / Layout / Table):

| | |
|---|---|
| Text | Font, size, colour, highlight; Bold/Italic/Underline/Strike (Ctrl+B/I/U); paragraph styles |
| Paragraphs | Bullets, numbering, list levels (Tab / Shift+Tab; Enter on an empty item ends the list); alignment; line spacing (1.0–3.0, at least, exactly), spacing before/after, indents and hanging indent (Paragraph dialog); borders (bottom / top / box) |
| Insert | Page break (Ctrl+Enter), table, picture, hyperlink (Ctrl+K) |
| Tables | Rows above/below, columns left/right, delete row/column/table; Tab / Shift+Tab move between cells (Tab in the last cell adds a row) |
| Find | Find (Ctrl+F) and Replace (Ctrl+H) with match case / whole words |
| Layout | Margins (Normal / Narrow / Moderate / Wide / custom in cm), orientation, paper size (Letter, A4, Legal, A5, Executive) — whole document |
| Header & Footer | Per section (Insert tab): different first page, "Same as previous section", page number and "Page X of Y" fields |
| New | Templates: Blank, Letter, Report, CV, Invoice, Memo |
| Spelling | Offline, British English by default (US English for documents marked en-US), from the SCOWL Hunspell dictionaries (MIT/BSD). Red wavy underline; right-click for suggestions, Ignore All, Add to Dictionary. Skips acronyms, web and e-mail addresses. Toggle in the status bar |
| Clipboard | Copy/paste within Prime Word is lossless (model properties travel in `data-pw-*` attributes). From Word and the web: headings, alignment, bold/italic/underline/strike, font, size, colour, highlight, super/subscript, links, bulleted/numbered and nested lists (incl. Word's own list markup), tables (borders, shading, merged columns) and embedded pictures |
| View | Print Layout / Web Layout, zoom 25–200% |
| Output | Export PDF (fonts embedded from the installed fonts), Print (Ctrl+P: opens the PDF in the default viewer) |
| Safety | Autosave for recovery 3 s after typing stops; after a crash, the next start offers Restore / Discard. Saves are atomic (temp file → flush → rename), so a crash mid-save cannot corrupt the document. Closing with unsaved changes asks first. |

Desktop automation (tests, batch use): `prime-word.exe <file.docx> --pw-export-pdf=<out.pdf>` exports
and exits; `--pw-type=<text>` types into the opened document.

`packages/editor/test/features.mjs` drives all of these commands headlessly and saves a .docx that
Word verifies (lists, fonts, colours, highlight, link, page break, table shape, picture, replace).

The desktop app may only touch the files you pick in its dialogs, plus read-only access to Office's
cloud-font folder (so documents using Aptos and other Office cloud fonts display correctly) —
see `apps/word/src-tauri/capabilities/default.json`.

| Path | What |
|---|---|
| [docs/requirements.md](docs/requirements.md) | MVP scope (v0.2, locked), quality requirements, acceptance criteria |
| [docs/architecture-decisions.md](docs/architecture-decisions.md) | ADR-000…013: own-the-stack strategy, permissive licences, build order |
| [packages/pdm/](packages/pdm/) | `@prime/pdm` — Prime Document Model types (ADR-004) |
| [packages/docx/](packages/docx/) | `@prime/docx` — our DOCX reader/writer + `prime-docx` CLI (ADR-006) |
| [packages/layout/](packages/layout/) | `@prime/layout` — layout engine: pages and positioned lines (ADR-007) |
| [packages/editor/](packages/editor/) | `@prime/editor` — ProseMirror schema derived from the model, lossless conversion, commands, rendering (ADR-005) |
| [apps/word/](apps/word/) | The Prime Word application (Vite + React) |
| [packages/pdf/](packages/pdf/) | `@prime/pdf` — PDF export from the layout engine: `node packages/pdf/dist/cli.js in.docx out.pdf` (ADR-008) |
| [experiments/e1-docx-roundtrip/](experiments/e1-docx-roundtrip/) | Experiment 1: DOCX round-trip fidelity harness, fixtures and candidates |
| [experiments/e2-layout/](experiments/e2-layout/) | Experiment 2: layout engine vs Word, marker by marker |
| [scripts/try-docx.ps1](scripts/try-docx.ps1) | Test your own `.docx` files: `.\scripts\try-docx.ps1 <file-or-folder> [-OpenInWord]` |

## Setup (Windows)

```powershell
npm install
npm run build                     # packages/pdm, packages/docx -> dist/
node packages/docx/dist/cli.js roundtrip in.docx out.docx
node packages/docx/dist/cli.js dump in.docx model.json

python -m venv .venv
.venv\Scripts\python -m pip install -r experiments\e1-docx-roundtrip\requirements.txt
cd experiments\e1-docx-roundtrip\candidates\html-pipeline; npm install; cd ..\..
..\..\.venv\Scripts\python run.py
```

Microsoft Word is used (via COM) for baseline page counts and PDF renders; pass `--no-word` to skip.
