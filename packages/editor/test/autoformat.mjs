// AutoCorrect / AutoFormat as you type: each case is typed one character at a time into an empty
// paragraph, exactly as the plugin sees it, and the paragraph's text (and list) is compared.
import { AUTOFORMAT_DEFAULTS, autoFormatAfter, blankDocument, createSchema, listKind, toProseMirror } from "@prime/editor";
import { EditorState, TextSelection } from "prosemirror-state";

const cases = [
  ['"Hello," she said. "It\'s fine." ', "“Hello,” she said. “It’s fine.” "],
  ["teh cat adn the dog ", "The cat and the dog "],
  ["i think so. i do. ", "I think so. I do. "],
  ["see e.g. this one ", "See e.g. this one "],
  ["THe end ", "The end "],
  ["(c) 2026 Prime(tm) wait... ok --> next ", "© 2026 Prime™ wait… ok → next "],
  ["pages 3 -- 5 and well--known ", "Pages 3 – 5 and well—known "],
  ["visit www.example.com now ", "Visit www.example.com now "],
  ["NASA and USA stay ", "NASA and USA stay "],
  ["he said 'yes' ", "He said ‘yes’ "],
  ["1. first item", "First item", "number"],
  ["- a point", "A point", "bullet"],
  ["* another one", "Another one", "bullet"],
  ["2. not a list start", "2. Not a list start"],
  ["I use pw daily ", "I use Prime Word daily "],
];
let failed = 0;
for (const [input, expected, list] of cases) {
  const base = blankDocument();
  let state = EditorState.create({ doc: toProseMirror(base, createSchema()) });
  state = state.apply(state.tr.setSelection(TextSelection.atStart(state.doc)));
  for (const ch of input) {
    const at = state.selection.from;
    state = state.apply(state.tr.insertText(ch, at));
    const change = state.tr;
    if (autoFormatAfter(change, at + 1, ch, { ...AUTOFORMAT_DEFAULTS, custom: { pw: "Prime Word" } }, () => base)) state = state.apply(change);
  }
  const p = state.doc.child(0);
  const kind = listKind(base, p.attrs.props.list?.numId);
  const ok = p.textContent === expected && kind === list;
  if (!ok) failed++;
  console.log(ok ? "ok  " : "FAIL", JSON.stringify(input), "→", JSON.stringify(p.textContent), kind ?? "");
}
console.log(failed ? `${failed} failed` : "all passed");
