// Quick formatting: change case, grow/shrink font, superscript, clear formatting, format painter.
import { blankDocument, changeCase, clearFormatting, copyFormat, createSchema, growFont, pasteFormat, toggleVertAlign, toProseMirror } from "@prime/editor";
import { EditorState, TextSelection } from "prosemirror-state";

const base = blankDocument();
base.sections[0].blocks = [
  { type: "paragraph", props: { align: "center", spacing: { after: 24 } }, content: [{ type: "text", text: "hello world. this is it", props: { bold: true, size: 11 } }] },
  { type: "paragraph", props: {}, content: [{ type: "text", text: "Plain second paragraph" }] },
];
let state = EditorState.create({ doc: toProseMirror(base, createSchema()) });
const run = (cmd) => { cmd(state, (tr) => { state = state.apply(tr); }); };
const select = (from, to) => { state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to))); };
const p1 = () => state.doc.child(0);
const runOf = (n) => n.marks.find((m) => m.type.name === "run")?.attrs.props ?? {};
const end1 = 1 + p1().content.size;
let ok = true;
const check = (label, got, want) => { const pass = JSON.stringify(got) === JSON.stringify(want); ok &&= pass; console.log(pass ? "ok  " : "FAIL", label, JSON.stringify(got)); };

select(1, end1);
run(changeCase("sentence")); check("sentence", p1().textContent, "Hello world. This is it");
run(changeCase("upper")); check("upper", p1().textContent, "HELLO WORLD. THIS IS IT");
run(changeCase("title")); check("title", p1().textContent, "Hello World. This Is It");
run(changeCase("cycle")); check("cycle (mixed → lower)", p1().textContent, "hello world. this is it");
run(changeCase("cycle")); check("cycle (lower → UPPER)", p1().textContent, "HELLO WORLD. THIS IS IT");
check("bold kept", runOf(p1().firstChild).bold, true);
run(growFont(() => base, 1)); check("grow 11 → 12", runOf(p1().firstChild).size, 12);
run(growFont(() => base, 1)); check("grow 12 → 14", runOf(p1().firstChild).size, 14);
run(growFont(() => base, -1, true)); check("shrink by a point", runOf(p1().firstChild).size, 13);
run(toggleVertAlign("superscript")); check("superscript on", runOf(p1().firstChild).vertAlign, "superscript");
run(toggleVertAlign("superscript")); check("superscript off", runOf(p1().firstChild).vertAlign, undefined);
// Painter: copy the first paragraph's look, apply it to the whole second paragraph.
select(2, 2);
const f = copyFormat(state);
const s2 = p1().nodeSize + 1;
select(s2, s2 + state.doc.child(1).content.size);
run(pasteFormat(f));
check("painter: run", runOf(state.doc.child(1).firstChild).bold, true);
check("painter: paragraph", state.doc.child(1).attrs.props.align, "center");
select(1, end1);
run(clearFormatting);
check("clear: run", runOf(p1().firstChild), {});
check("clear: paragraph", p1().attrs.props, {});
console.log(ok ? "all passed" : "FAILED");
