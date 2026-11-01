// Review commands: list/accept/reject tracked changes, add and delete comments; Word verifies.
// Usage: node review.mjs <in.docx (with tracked changes and a comment)> <accepted.docx> <rejected.docx> <commented.docx>
import { readFileSync, writeFileSync } from "node:fs";
import { readDocx, writeDocx } from "@prime/docx";
import { acceptOrReject, addComment, anchoredComments, changes, createSchema, deleteComment, fromProseMirror, toProseMirror } from "@prime/editor";
import { EditorState, TextSelection } from "prosemirror-state";

const [input, accepted, rejected, commented] = process.argv.slice(2);
const schema = createSchema();
const open = () => {
  const { doc } = readDocx(readFileSync(input));
  return { doc, state: EditorState.create({ doc: toProseMirror(doc, schema) }) };
};
const apply = (s, cmd) => { let out = s; if (!cmd(s, (tr) => { out = s.apply(tr); })) throw new Error("refused"); return out; };

let { doc, state } = open();
console.log("changes", JSON.stringify(changes(state.doc).map((c) => [c.kind, c.author, state.doc.textBetween(c.from, c.to)])));
console.log("comments", JSON.stringify(anchoredComments(state.doc, doc).map((c) => [c.id, c.author, c.text, state.doc.textBetween(c.from, c.to)])));
writeFileSync(accepted, writeDocx(fromProseMirror(apply(state, acceptOrReject(true, true)).doc, doc)));
({ doc, state } = open());
writeFileSync(rejected, writeDocx(fromProseMirror(apply(state, acceptOrReject(false, true)).doc, doc)));

// Add a comment on the first word of the first non-empty paragraph; delete the original comment.
({ doc, state } = open());
let at;
state.doc.descendants((n, pos) => { if (at === undefined && n.isTextblock && n.textContent.length > 8) at = pos + 1; return at === undefined; });
state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, at, at + 5)));
state = apply(state, addComment(() => doc, "Prime Tester", "Please check this word."));
const old = anchoredComments(state.doc, doc).find((c) => c.author !== "Prime Tester");
if (old) state = apply(state, deleteComment(old.id));
console.log("after", JSON.stringify(anchoredComments(state.doc, doc).map((c) => [c.id, c.author, c.text, state.doc.textBetween(c.from, c.to)])));
writeFileSync(commented, writeDocx(fromProseMirror(state.doc, doc)));
