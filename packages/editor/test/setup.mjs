// Page setup + header/footer editing through the real commands; saves a .docx for Word to verify.
// Usage: node setup.mjs <in.docx> <out.docx>
import { readFileSync, writeFileSync } from "node:fs";
import { readDocx, writeDocx } from "@prime/docx";
import {
  createSchema, fromProseMirror, insertField, PAPER, pageSetup, setPageSetup, setStory, storyFromProseMirror,
  storyToProseMirror, toProseMirror,
} from "@prime/editor";
import { EditorState } from "prosemirror-state";

const [input, output] = process.argv.slice(2);
const { doc } = readDocx(readFileSync(input));
const schema = createSchema();
let state = EditorState.create({ doc: toProseMirror(doc, schema) });
const run = (s, cmd, set) => { if (!cmd(s, (tr) => set(s.apply(tr)))) throw new Error("refused"); };

const cm = (v) => (v * 72) / 2.54;
run(state, setPageSetup({ paper: PAPER.A4, orientation: "landscape", margins: { top: cm(2), bottom: cm(2), left: cm(2), right: cm(2) } }), (s) => { state = s; });

// Edit the header and footer in their own little editors, as the app does.
const editStory = (text, withFields) => {
  let story = EditorState.create({ doc: storyToProseMirror([], schema) });
  story = story.apply(story.tr.insertText(text, 1));
  if (withFields) {
    const set = (s) => { story = s; };
    run(story, insertField("PAGE"), set);
    story = story.apply(story.tr.insertText(" of "));
    run(story, insertField("NUMPAGES"), set);
  }
  return { blocks: storyFromProseMirror(story.doc) };
};
run(state, setStory("header", editStory("Prime Word test header", false)), (s) => { state = s; });
run(state, setStory("footer", editStory("Page ", true)), (s) => { state = s; });

console.log("page setup now:", JSON.stringify(pageSetup(state)));
writeFileSync(output, writeDocx(fromProseMirror(state.doc, doc)));
