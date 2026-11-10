import { DOMParser as PmDOMParser } from "prosemirror-model";
import {
  createRender, createSchema, insertField, ownStory, paragraphStyles, sectionInfo, setStory, setTitlePage, storyAt,
  storyFromProseMirror, storyToProseMirror, toggleRun, typingParser } from "@prime/editor";
import type { Measurer } from "@prime/layout";
import type { Block, HeaderFooterKind, PrimeDocument, Story } from "@prime/pdm";
import { baseKeymap } from "prosemirror-commands";
import { history, redo, undo } from "prosemirror-history";
import { keymap } from "prosemirror-keymap";
import { type Command, EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { useEffect, useRef, useState } from "react";

type Which = "header" | "footer";
type Key = `${Which}/${HeaderFooterKind}`;

/** A small editor for one header or footer. Reports its blocks on every change. */
function StoryEditor({ label, doc, measurer, blocks, onChange, disabled }: {
  label: string; doc: PrimeDocument; measurer: Measurer; blocks: Block[]; onChange(blocks: Block[]): void; disabled: boolean;
}) {
  const mount = useRef<HTMLDivElement>(null);
  const viewRef = useRef<EditorView | null>(null);
  useEffect(() => {
    if (!mount.current) return;
    const schema = createSchema(createRender(doc, measurer));
    const v = new EditorView(mount.current, {
      // Typing read with the document's own rules; pasting with all of them (web formatting too).
      domParser: typingParser(schema), clipboardParser: PmDOMParser.fromSchema(schema),
      state: EditorState.create({
        doc: storyToProseMirror(blocks, schema),
        plugins: [
          history(),
          keymap({ "Mod-z": undo, "Mod-y": redo, "Mod-b": toggleRun("bold"), "Mod-i": toggleRun("italic"), "Mod-u": toggleRun("underline") }),
          keymap(baseKeymap),
          paragraphStyles(doc, measurer),
        ],
      }),
      editable: () => !disabled,
      dispatchTransaction(tr) {
        v.updateState(v.state.apply(tr));
        if (tr.docChanged) onChange(storyFromProseMirror(v.state.doc));
      },
    });
    viewRef.current = v;
    return () => v.destroy();
  }, [disabled]);

  const run = (cmd: Command) => () => {
    const v = viewRef.current;
    if (v && !disabled) { cmd(v.state, v.dispatch, v); v.focus(); }
  };
  const pageXofY: Command = (_state, dispatch, view) => {
    if (!dispatch || !view) return true;
    view.dispatch(view.state.tr.insertText("Page "));
    insertField("PAGE")(view.state, view.dispatch, view);
    view.dispatch(view.state.tr.insertText(" of "));
    insertField("NUMPAGES")(view.state, view.dispatch, view);
    return true;
  };

  return (
    <section className={`hf-story ${disabled ? "linked" : ""}`}>
      <div className="hf-bar">
        <strong>{label}</strong>
        <button onClick={run(insertField("PAGE"))} disabled={disabled}>Page number</button>
        <button onClick={run(pageXofY)} disabled={disabled}>Page X of Y</button>
      </div>
      <div ref={mount} className="pw-editor hf-editor" />
    </section>
  );
}

const isEmpty = (blocks: Block[]) => blocks.every((b) => b.type === "paragraph" && b.content.length === 0);

/**
 * Header & Footer for the section at the cursor: "Different first page" (separate first-page header
 * and footer) and, from the second section on, "Same as previous section" (Word's Link to Previous).
 */
export function HeaderFooterEditor({ view, doc, measurer, onClose }: {
  view: EditorView; doc: PrimeDocument; measurer: Measurer; onClose(): void;
}) {
  const info = sectionInfo(view.state);
  const [differentFirst, setDifferentFirst] = useState(info.titlePage);
  const [linked, setLinked] = useState(info.number > 1 && !ownStory(view.state, "header") && !ownStory(view.state, "footer"));
  const [kind, setKind] = useState<HeaderFooterKind>("default");
  // Edited content per header/footer kind; starts from what the section shows today (own or inherited).
  const edits = useRef(new Map<Key, Block[]>());
  const initial = (which: Which, k: HeaderFooterKind): Story | undefined => storyAt(view.state, which, k);
  const blocksFor = (which: Which, k: HeaderFooterKind) => edits.current.get(`${which}/${k}`) ?? initial(which, k)?.blocks ?? [];

  const apply = () => {
    const kinds: HeaderFooterKind[] = differentFirst ? ["default", "first"] : ["default"];
    for (const which of ["header", "footer"] as const) {
      for (const k of ["default", "first"] as const) {
        let story: Story | undefined;
        if (!linked && kinds.includes(k)) {
          const blocks = blocksFor(which, k);
          const before = initial(which, k);
          story = isEmpty(blocks) ? undefined : { blocks, ...(before?.preservedRels ? { preservedRels: before.preservedRels } : {}) };
        }
        // Only write what changed: a linked section, or a kind no longer used, has no story of its own.
        if (linked || !kinds.includes(k) || edits.current.has(`${which}/${k}`) || story !== ownStory(view.state, which, k)) {
          setStory(which, story, k)(view.state, view.dispatch, view);
        }
      }
    }
    if (differentFirst !== info.titlePage) setTitlePage(differentFirst)(view.state, view.dispatch, view);
    onClose();
    view.focus();
  };

  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); }}>
      <div className="hf-dialog" role="dialog" aria-label="Header and footer">
        <h3>Header &amp; Footer{info.count > 1 ? ` — Section ${info.number} of ${info.count}` : ""}</h3>
        <div className="hf-options">
          <label><input type="checkbox" checked={differentFirst} onChange={(e) => { setDifferentFirst(e.target.checked); if (!e.target.checked) setKind("default"); }} /> Different first page</label>
          {info.number > 1 && (
            <label><input type="checkbox" checked={linked} onChange={(e) => setLinked(e.target.checked)} /> Same as previous section</label>
          )}
          {differentFirst && (
            <span className="hf-tabs">
              <button className={kind === "first" ? "on" : ""} onClick={() => setKind("first")}>First page</button>
              <button className={kind === "default" ? "on" : ""} onClick={() => setKind("default")}>Other pages</button>
            </span>
          )}
        </div>
        {linked && <p className="hf-hint">This section uses the previous section's header and footer. Untick "Same as previous section" to give it its own.</p>}
        <StoryEditor key={`header/${kind}/${linked}`} label={kind === "first" ? "First-page header" : "Header"} doc={doc} measurer={measurer}
          blocks={blocksFor("header", kind)} disabled={linked} onChange={(b) => edits.current.set(`header/${kind}`, b)} />
        <StoryEditor key={`footer/${kind}/${linked}`} label={kind === "first" ? "First-page footer" : "Footer"} doc={doc} measurer={measurer}
          blocks={blocksFor("footer", kind)} disabled={linked} onChange={(b) => edits.current.set(`footer/${kind}`, b)} />
        <div className="hf-actions">
          <button onClick={onClose}>Cancel</button>
          <button className="primary" onClick={apply}>Done</button>
        </div>
      </div>
    </div>
  );
}
