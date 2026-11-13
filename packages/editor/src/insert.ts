/** Insert commands: page break, hyperlink, picture. */
import type { PrimeDocument } from "@prime/pdm";
import { type Command, Plugin, TextSelection } from "prosemirror-state";

/** Inserts a page break: the text after the cursor starts on a new page. */
export const insertPageBreak: Command = (state, dispatch) => {
  const { hard_break } = state.schema.nodes;
  if (!dispatch) return true;
  const tr = state.tr.replaceSelectionWith(hard_break.create({ kind: "page" }), false);
  // Word puts the following text in a new paragraph.
  tr.split(tr.selection.from);
  dispatch(tr.scrollIntoView());
  return true;
};

/**
 * Typing where the browser cannot hold a caret: between two page breaks (a blank page in Word files,
 * which keep several breaks in one paragraph) or between other non-text inline nodes. Page breaks are
 * drawn as non-editable boxes, so the browser moves typed text to the next text it finds — onto a
 * later page. Here the editor inserts the text at its own (correct) position instead.
 */
export function caretBetweenAtoms(): Plugin {
  return new Plugin({
    props: {
      handleDOMEvents: {
        beforeinput: (view, e) => {
          const ev = e as InputEvent;
          if (ev.inputType !== "insertText" || !ev.data || view.composing) return false;
          const { state } = view;
          const { $from, empty, from, to } = state.selection;
          if (!empty || !$from.parent.inlineContent) return false;
          const before = $from.nodeBefore, after = $from.nodeAfter;
          const breakNear = [before, after].some((n) => n?.type.name === "hard_break" && n.attrs.kind !== "line");
          if (!breakNear || before?.isText || after?.isText) return false;
          ev.preventDefault();
          const handled = view.someProp("handleTextInput", (f) => f(view, from, to, ev.data!));
          if (!handled) {
            const marks = state.storedMarks ?? $from.marks();
            view.dispatch(state.tr.replaceWith(from, to, state.schema.text(ev.data!, marks)).scrollIntoView());
          }
          return true;
        },
      },
    },
  });
}

/**
 * Links the selection to `href`; with an empty selection, inserts `text` (default: the address) as
 * a link. An empty href removes links from the selection.
 */
export function setLink(href: string, text?: string): Command {
  return (state, dispatch) => {
    const link = state.schema.marks.link;
    if (!dispatch) return true;
    const { from, to, empty } = state.selection;
    const tr = state.tr;
    if (!href) {
      tr.removeMark(from, to, link);
    } else if (empty) {
      const label = text || href;
      tr.insertText(label, from);
      // (Typed inside another link: this one replaces it there, rather than nesting in it.)
      tr.removeMark(from, from + label.length, link).addMark(from, from + label.length, link.create({ href }));
      tr.setSelection(TextSelection.create(tr.doc, from + label.length));
    } else {
      tr.removeMark(from, to, link).addMark(from, to, link.create({ href }));
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** The link address at the cursor, if any (to edit an existing link). */
export function linkAt(state: { selection: { $from: { marks(): readonly { type: { name: string }; attrs: Record<string, unknown> }[] } } }): string | undefined {
  return state.selection.$from.marks().find((m) => m.type.name === "link")?.attrs.href as string | undefined;
}

const EXT: Record<string, string> = { "image/png": "png", "image/jpeg": "jpeg", "image/gif": "gif", "image/bmp": "bmp" };

/**
 * Inserts a picture. The image is added to the document's image resources; its display size is
 * its natural size (at 96 dpi), scaled down to fit the text width.
 */
export function insertImage(doc: () => PrimeDocument, image: { data: string; contentType: string; width: number; height: number; name?: string }): Command {
  return (state, dispatch) => {
    const ext = EXT[image.contentType];
    if (!ext) return false;
    if (!dispatch) return true;
    const d = doc();
    let n = Object.keys(d.images).length + 1;
    while (d.images[`word/media/prime-image${n}.${ext}`] || d.package.parts.some((p) => p.name === `word/media/prime-image${n}.${ext}`)) n++;
    const imageId = `word/media/prime-image${n}.${ext}`;
    d.images[imageId] = { contentType: image.contentType, extension: ext, data: image.data };
    const section = d.sections[0]?.props ?? {};
    const maxWidth = (section.pageWidth ?? 612) - (section.margins?.left ?? 72) - (section.margins?.right ?? 72);
    let width = (image.width * 72) / 96;
    let height = (image.height * 72) / 96;
    if (width > maxWidth) {
      height *= maxWidth / width;
      width = maxWidth;
    }
    const node = state.schema.nodes.image.create({ imageId, width, height, name: image.name ?? `Picture ${n}`, description: image.name ?? null });
    dispatch(state.tr.replaceSelectionWith(node).scrollIntoView());
    return true;
  };
}
