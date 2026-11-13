/**
 * The ProseMirror schema, derived from the Prime Document Model (ADR-004/005).
 *
 * Every PDM construct has a node or mark; properties travel as attributes, verbatim, so that
 * document -> editor -> document is lossless. How things look is decided by `render` (render.ts),
 * which is optional so the schema also works headless (conversion, tests, Node).
 *
 * Clipboard: everything Prime Word writes carries its model properties in data-pw-* attributes, so
 * copying within Prime Word is lossless. HTML from elsewhere (Word, the web) is parsed into
 * temporary "x_" marks which paste.ts merges into run formatting after the paste.
 */
import { type DOMOutputSpec, DOMParser, type Mark, type Node, Schema, type TagParseRule } from "prosemirror-model";
import { cellFromHtml, paragraphFromHtml, tableFromHtml } from "./paste.js";

/** Rendering hooks supplied by the app (browser). Without them, nodes render as plain elements. */
export interface Render {
  paragraph(node: Node): DOMOutputSpec;
  /** Inline CSS for a run's formatting. */
  run(mark: Mark): string;
  table(node: Node): DOMOutputSpec;
  cell(node: Node): DOMOutputSpec;
  image(node: Node): DOMOutputSpec;
  runPreserved(node: Node): DOMOutputSpec;
  inlinePreserved(node: Node): DOMOutputSpec;
  blockPreserved(node: Node): DOMOutputSpec;
  sectionBreak(node: Node): DOMOutputSpec;
}

const props = { props: { default: {} } };
/** A tracked formatting change on props (the formatting before, kept as Word keeps it). */
const changed = (props: { extra?: { name: string }[] } | null, name: string) => !!props?.extra?.some((x) => x.name === name);
const json = (v: unknown) => JSON.stringify(v);
const parse = (el: HTMLElement, name: string) => JSON.parse(el.getAttribute(name) ?? "null");

/** Our own elements first (lossless), then generic HTML. */
const own = (tag: string, attr: string, to: (v: any) => Record<string, unknown>): TagParseRule => ({
  tag: `${tag}[${attr}]`, priority: 100, getAttrs: (el) => to(parse(el as HTMLElement, attr)),
});

export function createSchema(render?: Render): Schema {
  return new Schema({
    nodes: {
      doc: {
        content: "block+",
        // The last section's properties, headers and footers (earlier sections end in section_break nodes);
        // the footnotes' and endnotes' text (PrimeDocument.notes), so editing a note can be undone.
        // The page colour (PrimeDocument.background), so changing it can be undone.
        // Automatic hyphenation (PrimeDocument.hyphenation), likewise.
        attrs: { section: { default: null }, notes: { default: null }, background: { default: null }, hyphenation: { default: null }, trackRevisions: { default: false },
          // The stylesheet once a style is created or modified (it then overrides the file's), likewise.
          styles: { default: null } },
      },
      paragraph: {
        group: "block",
        content: "inline*",
        attrs: { ...props },
        toDOM: (n) => ["p", { "data-pw-props": json(n.attrs.props), ...(changed(n.attrs.props, "w:pPrChange") ? { class: "pw-fmt-para" } : {}) }, 0],
        parseDOM: [
          own("p", "data-pw-props", (props) => ({ props })),
          ...[1, 2, 3, 4, 5, 6].map((level): TagParseRule => ({ tag: `h${level}`, attrs: { props: { style: `Heading${level}` } } })),
          { tag: "li", getAttrs: (el) => ({ props: paragraphFromHtml(el as HTMLElement) }) },
          { tag: "p", getAttrs: (el) => ({ props: paragraphFromHtml(el as HTMLElement) }) },
          { tag: "blockquote", attrs: { props: { indent: { left: 36 } } } },
        ],
      },
      table: {
        group: "block",
        content: "table_row+",
        isolating: true,
        attrs: { ...props, grid: { default: [] } },
        toDOM: (n) => render?.table(n) ?? ["table", { "data-pw-props": json(n.attrs.props), "data-pw-grid": json(n.attrs.grid) }, ["tbody", 0]],
        parseDOM: [
          { tag: "table[data-pw-props]", priority: 100, getAttrs: (el) => ({ props: parse(el as HTMLElement, "data-pw-props"), grid: parse(el as HTMLElement, "data-pw-grid") ?? [] }) },
          { tag: "table", getAttrs: (el) => tableFromHtml(el as HTMLElement) },
        ],
      },
      table_row: {
        content: "table_cell+",
        attrs: { ...props, extra: { default: null } },
        toDOM: (n) => {
          // A set row height (w:trHeight, in twips): at least that tall ("exact" too: content may still
          // grow it on screen, never shrink it below the layout's).
          const h = n.attrs.props?.height?.val ? Number(n.attrs.props.height.val) / 20 : 0;
          return ["tr", {
            "data-pw-props": json(n.attrs.props), ...(n.attrs.extra ? { "data-pw-extra": json(n.attrs.extra) } : {}),
            ...(h ? { style: `height:${h}pt` } : {}),
          }, 0];
        },
        parseDOM: [
          { tag: "tr[data-pw-props]", priority: 100, getAttrs: (el) => ({ props: parse(el as HTMLElement, "data-pw-props"), extra: parse(el as HTMLElement, "data-pw-extra") }) },
          { tag: "tr" },
        ],
      },
      table_cell: {
        content: "block+",
        isolating: true,
        attrs: { ...props },
        toDOM: (n) => render?.cell(n) ?? ["td", { "data-pw-props": json(n.attrs.props) }, 0],
        parseDOM: [
          own("td", "data-pw-props", (props) => ({ props })),
          { tag: "td", getAttrs: (el) => ({ props: cellFromHtml(el as HTMLElement) }) },
          { tag: "th", getAttrs: (el) => ({ props: cellFromHtml(el as HTMLElement) }) },
        ],
      },
      content_control: {
        group: "block",
        content: "block+",
        attrs: { settings: { default: [] } },
        toDOM: (n) => ["div", { class: "pw-content-control", "data-pw-control": json(n.attrs.settings) }, 0],
        parseDOM: [own("div", "data-pw-control", (settings) => ({ settings }))],
      },
      block_preserved: {
        group: "block",
        atom: true,
        attrs: { preserved: {} },
        toDOM: (n) => render?.blockPreserved(n) ?? ["div", { class: "pw-preserved", "data-pw-preserved": json(n.attrs.preserved) }],
        parseDOM: [own("div", "data-pw-preserved", (preserved) => ({ preserved }))],
      },
      section_break: {
        group: "block",
        atom: true,
        selectable: false,
        attrs: { section: {} },
        toDOM: (n) => render?.sectionBreak(n) ?? ["hr", { class: "pw-section-break", "data-pw-section": json(n.attrs.section) }],
        parseDOM: [{ tag: "[data-pw-section]", priority: 100, getAttrs: (el) => ({ section: parse(el as HTMLElement, "data-pw-section") }) }],
      },
      text: { group: "inline" },
      tab: {
        group: "inline", inline: true, atom: true, attrs: { ...props },
        toDOM: (n) => ["span", { class: "pw-tab", "data-pw-tab": json(n.attrs.props) }, "\t"],
        parseDOM: [own("span", "data-pw-tab", (props) => ({ props }))],
      },
      hard_break: {
        group: "inline", inline: true, atom: true,
        attrs: { ...props, kind: { default: "line" }, clear: { default: null } },
        toDOM: (n) => (n.attrs.kind === "line"
          ? ["br", { "data-pw-break": json(n.attrs) }]
          : ["span", { class: n.attrs.kind === "column" ? "pw-page-break pw-column-break" : "pw-page-break", contenteditable: "false", "data-pw-break": json(n.attrs) }]),
        parseDOM: [
          { tag: "[data-pw-break]", priority: 100, getAttrs: (el) => parse(el as HTMLElement, "data-pw-break") },
          { tag: "br" },
        ],
      },
      image: {
        group: "inline", inline: true, atom: true, draggable: true,
        // `float`: wrap and position of a floating picture (FloatInfo); null: in line with the text.
        attrs: { ...props, imageId: {}, width: {}, height: {}, name: { default: null }, description: { default: null }, float: { default: null } },
        toDOM: (n) => render?.image(n) ?? ["img", { "data-pw-image": json(n.attrs) }],
        parseDOM: [
          { tag: "img[data-pw-image]", priority: 100, getAttrs: (el) => parse(el as HTMLElement, "data-pw-image") },
          // Pictures embedded in pasted HTML; paste.ts turns them into image resources.
          {
            tag: "img[src^='data:image/']",
            getAttrs: (el) => {
              const img = el as HTMLImageElement;
              const w = img.width || Number(img.getAttribute("width")) || 300;
              const h = img.height || Number(img.getAttribute("height")) || 200;
              return { imageId: `pending:${img.getAttribute("src")}`, width: (w * 72) / 96, height: (h * 72) / 96, description: img.alt || null };
            },
          },
        ],
      },
      // A footnote or endnote reference: its number comes from CSS counters (references in order).
      note_ref: {
        group: "inline", inline: true, atom: true, attrs: { ...props, kind: { default: "footnote" }, id: { default: "" } },
        toDOM: (n) => ["sup", { class: `pw-noteref pw-${n.attrs.kind}`, contenteditable: "false", "data-pw-noteref": json(n.attrs) }],
        parseDOM: [{ tag: "sup[data-pw-noteref]", priority: 100, getAttrs: (el) => parse(el as HTMLElement, "data-pw-noteref") }],
      },
      // The note's own number at the start of its text (in the Notes pane, shown by the pane).
      note_mark: {
        group: "inline", inline: true, atom: true, attrs: { ...props, kind: { default: "footnote" } },
        toDOM: (n) => ["span", { class: "pw-notemark", contenteditable: "false", "data-pw-notemark": json(n.attrs) }],
        parseDOM: [{ tag: "span[data-pw-notemark]", priority: 100, getAttrs: (el) => parse(el as HTMLElement, "data-pw-notemark") }],
      },
      field_char: {
        group: "inline", inline: true, atom: true, attrs: { ...props, kind: {} },
        toDOM: (n) => ["span", { class: "pw-hidden", "data-pw-fieldchar": json(n.attrs) }],
        parseDOM: [{ tag: "span[data-pw-fieldchar]", priority: 100, getAttrs: (el) => parse(el as HTMLElement, "data-pw-fieldchar") }],
      },
      instr: {
        group: "inline", inline: true, atom: true, attrs: { ...props, text: { default: "" } },
        toDOM: (n) => ["span", { class: "pw-hidden", "data-pw-instr": json(n.attrs) }],
        parseDOM: [{ tag: "span[data-pw-instr]", priority: 100, getAttrs: (el) => parse(el as HTMLElement, "data-pw-instr") }],
      },
      run_preserved: {
        // `float`: where a floating picture or text box sits (FloatInfo), for display; null otherwise.
        group: "inline", inline: true, atom: true, attrs: { ...props, preserved: {}, float: { default: null } },
        toDOM: (n) => render?.runPreserved(n) ?? ["span", { class: "pw-preserved", "data-pw-runpreserved": json(n.attrs) }],
        parseDOM: [{ tag: "[data-pw-runpreserved]", priority: 100, getAttrs: (el) => parse(el as HTMLElement, "data-pw-runpreserved") }],
      },
      inline_preserved: {
        group: "inline", inline: true, atom: true, attrs: { preserved: {} },
        toDOM: (n) => render?.inlinePreserved(n) ?? ["span", { class: "pw-preserved", "data-pw-inlinepreserved": json(n.attrs) }],
        parseDOM: [{ tag: "[data-pw-inlinepreserved]", priority: 100, getAttrs: (el) => parse(el as HTMLElement, "data-pw-inlinepreserved") }],
      },
    },
    marks: {
      // Containers, outermost first: the converter rebuilds nesting in this order. Each can hold another
      // of its kind (excludes ""), as in Word: a link inside a link (tables of contents), a field inside a
      // field, a content control inside another. `again` tells apart one inside an identical one (Word
      // writes a mailto link inside the same link): marks are a set, and equal marks would merge.
      revision: {
        attrs: { kind: {}, attrs: { default: {} }, again: { default: 0 } },
        inclusive: false,
        excludes: "",
        toDOM: (m) => [m.attrs.kind === "del" ? "del" : "ins", { class: `pw-${m.attrs.kind}`, "data-pw-revision": json(m.attrs) }, 0],
        parseDOM: [{ tag: "[data-pw-revision]", priority: 100, getAttrs: (el) => parse(el as HTMLElement, "data-pw-revision") }],
      },
      inline_control: {
        attrs: { settings: { default: [] }, again: { default: 0 } },
        excludes: "",
        toDOM: (m) => ["span", { class: "pw-content-control", "data-pw-inlinecontrol": json(m.attrs.settings) }, 0],
        parseDOM: [own("span", "data-pw-inlinecontrol", (settings) => ({ settings }))],
      },
      link: {
        attrs: { href: { default: null }, anchor: { default: null }, tooltip: { default: null }, again: { default: 0 } },
        inclusive: false,
        excludes: "",
        toDOM: (m) => ["a", {
          href: m.attrs.href ?? `#${m.attrs.anchor ?? ""}`, title: m.attrs.tooltip ?? m.attrs.href ?? "", "data-pw-link": json(m.attrs),
        }, 0],
        parseDOM: [
          { tag: "a[data-pw-link]", priority: 100, getAttrs: (el) => parse(el as HTMLElement, "data-pw-link") },
          {
            tag: "a[href]",
            getAttrs: (el) => {
              const href = (el as HTMLElement).getAttribute("href") ?? "";
              return href.startsWith("#") ? { anchor: href.slice(1) } : /^(https?:|mailto:)/i.test(href) ? { href } : false;
            },
          },
        ],
      },
      field: {
        attrs: { instr: { default: "" }, again: { default: 0 } },
        inclusive: false,
        excludes: "",
        toDOM: (m) => ["span", { class: "pw-field", "data-pw-field": m.attrs.instr }, 0],
        parseDOM: [{ tag: "span[data-pw-field]", priority: 100, getAttrs: (el) => ({ instr: (el as HTMLElement).getAttribute("data-pw-field") }) }],
      },
      // Character formatting: the run's RunProps, verbatim.
      run: {
        attrs: { ...props },
        toDOM: (m) => ["span", { style: render?.run(m) ?? "", "data-pw-run": json(m.attrs.props), ...(changed(m.attrs.props, "w:rPrChange") ? { class: "pw-fmt" } : {}) }, 0],
        parseDOM: [own("span", "data-pw-run", (props) => ({ props }))],
      },
      // Formatting from pasted HTML; merged into `run` by paste.ts and never kept in a document.
      x_bold: {
        parseDOM: [{ tag: "b" }, { tag: "strong" }, { style: "font-weight", getAttrs: (v) => (/^(bold|[6-9]00)$/.test(v as string) ? null : false) }],
        toDOM: () => ["b", 0],
      },
      x_italic: {
        parseDOM: [{ tag: "i" }, { tag: "em" }, { style: "font-style=italic" }],
        toDOM: () => ["i", 0],
      },
      x_underline: {
        parseDOM: [{ tag: "u" }, { style: "text-decoration", getAttrs: (v) => ((v as string).includes("underline") ? null : false) }],
        toDOM: () => ["u", 0],
      },
      x_strike: {
        parseDOM: [{ tag: "s" }, { tag: "strike" }, { tag: "del:not([data-pw-revision])" }, { style: "text-decoration", getAttrs: (v) => ((v as string).includes("line-through") ? null : false) }],
        toDOM: () => ["s", 0],
      },
      x_style: {
        attrs: { css: { default: "" } },
        excludes: "",
        parseDOM: [
          // Word's list bullets/numbers are drawn from the list definition; its literal copy is dropped.
          { tag: "span[style*='mso-list:Ignore']", ignore: true, priority: 60 },
          { tag: "span[style]:not([data-pw-run])", getAttrs: (el) => ({ css: (el as HTMLElement).getAttribute("style") ?? "" }) },
          { tag: "font", getAttrs: (el) => ({ css: [(el as HTMLElement).getAttribute("face") ? `font-family:${(el as HTMLElement).getAttribute("face")}` : "", (el as HTMLElement).getAttribute("color") ? `color:${(el as HTMLElement).getAttribute("color")}` : ""].join(";") }) }],
        toDOM: (m) => ["span", { style: m.attrs.css }, 0],
      },
      x_super: { parseDOM: [{ tag: "sup" }], toDOM: () => ["sup", 0] },
      x_sub: { parseDOM: [{ tag: "sub" }], toDOM: () => ["sub", 0] },
    },
  });
}

/** Container marks, outermost first. */
export const CONTAINER_MARKS = ["revision", "inline_control", "link", "field"] as const;

/**
 * Reading what is typed: the document's own HTML only, without the rules for pasted web formatting
 * (x_bold, x_style…). Those match our own runs' inline styles (font-weight: 700…): typing re-read a
 * whole formatted word as different, replaced it, and the formatting for the next text (bold turned off)
 * spread over the word before. Pasting keeps every rule (clipboardParser).
 */
export function typingParser(schema: Schema): DOMParser {
  const all = DOMParser.fromSchema(schema);
  return new DOMParser(schema, all.rules.filter((r) => !("mark" in r && r.mark?.startsWith("x_"))));
}
