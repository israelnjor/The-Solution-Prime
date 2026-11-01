/**
 * Prime Document Model <-> ProseMirror document. Lossless for everything the schema represents:
 * gate E3 runs DOCX -> PDM -> editor -> PDM -> DOCX through the Experiment 1 checker.
 */
import type { Block, Inline, Paragraph, PrimeDocument, Section, Table, TableRow } from "@prime/pdm";
import type { Mark, Node, Schema } from "prosemirror-model";
import { CONTAINER_MARKS } from "./schema.js";

// --- PDM -> ProseMirror --------------------------------------------------------------------

export function toProseMirror(doc: PrimeDocument, schema: Schema): Node {
  const blocks: Node[] = [];
  doc.sections.forEach((section, i) => {
    blocks.push(...section.blocks.flatMap((b) => block(b, schema)));
    if (i < doc.sections.length - 1) blocks.push(schema.nodes.section_break.create({ section: sectionAttrs(section) }));
  });
  if (!blocks.length) blocks.push(schema.nodes.paragraph.create());
  const last = doc.sections[doc.sections.length - 1];
  return schema.nodes.doc.create({ section: last ? sectionAttrs(last) : null }, blocks);
}

const sectionAttrs = (s: Section) => ({ props: s.props, headers: s.headers, footers: s.footers });

function block(b: Block, schema: Schema): Node[] {
  switch (b.type) {
    case "paragraph":
      return [schema.nodes.paragraph.create({ props: b.props }, inlines(b.content, schema, []))];
    case "table":
      return [table(b, schema)];
    case "contentControl": {
      const inner = b.blocks.flatMap((x) => block(x, schema));
      return [schema.nodes.content_control.create({ settings: b.settings }, inner.length ? inner : [schema.nodes.paragraph.create()])];
    }
    case "preserved":
      return [schema.nodes.block_preserved.create({ preserved: b.preserved })];
  }
}

function table(t: Table, schema: Schema): Node {
  const rows = t.rows.map((row) => schema.nodes.table_row.create(
    { props: row.props, extra: row.extra ?? null },
    row.cells.map((cell) => {
      const content = cell.blocks.flatMap((x) => block(x, schema));
      return schema.nodes.table_cell.create({ props: cell.props }, content.length ? content : [schema.nodes.paragraph.create()]);
    }),
  ));
  return schema.nodes.table.create({ props: t.props, grid: t.grid }, rows);
}

function inlines(items: Inline[], schema: Schema, marks: readonly Mark[]): Node[] {
  const out: Node[] = [];
  const n = schema.nodes;
  for (const i of items) {
    const runMarks = "props" in i && i.props ? [...marks, schema.marks.run.create({ props: i.props })] : marks;
    switch (i.type) {
      case "text":
        if (i.text) out.push(schema.text(i.text, runMarks));
        break;
      case "tab": out.push(n.tab.create({ props: i.props ?? {} }, null, marks)); break;
      case "break": out.push(n.hard_break.create({ props: i.props ?? {}, kind: i.kind, clear: i.clear ?? null }, null, marks)); break;
      case "image":
        out.push(n.image.create({
          props: i.props ?? {}, imageId: i.imageId, width: i.width, height: i.height,
          name: i.name ?? null, description: i.description ?? null,
        }, null, marks));
        break;
      case "fieldChar": out.push(n.field_char.create({ props: i.props ?? {}, kind: i.kind }, null, marks)); break;
      case "instr": out.push(n.instr.create({ props: i.props ?? {}, text: i.text }, null, marks)); break;
      case "runPreserved": out.push(n.run_preserved.create({ props: i.props ?? {}, preserved: i.preserved, float: i.float ?? null }, null, marks)); break;
      case "preserved": out.push(n.inline_preserved.create({ preserved: i.preserved }, null, marks)); break;
      case "link":
        out.push(...inlines(i.content, schema, [...marks, schema.marks.link.create({
          href: i.href ?? null, anchor: i.anchor ?? null, tooltip: i.tooltip ?? null,
        })]));
        break;
      case "field": out.push(...inlines(i.content, schema, [...marks, schema.marks.field.create({ instr: i.instr })])); break;
      case "contentControl":
        out.push(...inlines(i.content, schema, [...marks, schema.marks.inline_control.create({ settings: i.settings })]));
        break;
      case "revision":
        out.push(...inlines(i.content, schema, [...marks, schema.marks.revision.create({ kind: i.kind, attrs: i.attrs })]));
        break;
    }
  }
  return out;
}

// --- ProseMirror -> PDM --------------------------------------------------------------------

/** Rebuilds the document: content and sections from the editor, everything else from `base`. */
export function fromProseMirror(pm: Node, base: PrimeDocument): PrimeDocument {
  const sections: Section[] = [];
  let blocks: Block[] = [];
  pm.forEach((node) => {
    if (node.type.name === "section_break") {
      const s = node.attrs.section;
      sections.push({ props: s.props, headers: s.headers, footers: s.footers, blocks });
      blocks = [];
    } else blocks.push(...toBlocks(node));
  });
  const last = pm.attrs.section ?? { props: {}, headers: {}, footers: {} };
  sections.push({ props: last.props, headers: last.headers, footers: last.footers, blocks });
  return { ...base, sections };
}

function toBlocks(node: Node): Block[] {
  switch (node.type.name) {
    case "paragraph":
      return [paragraph(node)];
    case "table":
      return [toTable(node)];
    case "content_control": {
      const inner: Block[] = [];
      node.forEach((c) => inner.push(...toBlocks(c)));
      return [{ type: "contentControl", settings: node.attrs.settings, blocks: inner }];
    }
    case "block_preserved":
      return [{ type: "preserved", preserved: node.attrs.preserved }];
    default:
      return [];
  }
}

function toTable(node: Node): Table {
  const rows: TableRow[] = [];
  node.forEach((r) => {
    const cells: TableRow["cells"] = [];
    r.forEach((c) => {
      const blocks: Block[] = [];
      c.forEach((b) => blocks.push(...toBlocks(b)));
      cells.push({ props: c.attrs.props, blocks });
    });
    rows.push({ props: r.attrs.props, cells, ...(r.attrs.extra ? { extra: r.attrs.extra } : {}) });
  });
  return { type: "table", props: node.attrs.props, grid: node.attrs.grid, rows };
}

function paragraph(node: Node): Paragraph {
  const children: Node[] = [];
  node.forEach((c) => children.push(c));
  return { type: "paragraph", props: node.attrs.props, content: nest(children, 0) };
}

const sameMark = (a: Mark | undefined, b: Mark | undefined) => (!a && !b) || (!!a && !!b && a.eq(b));

/** Rebuilds container nesting (revision > content control > link > field) from marks. */
function nest(nodes: Node[], level: number): Inline[] {
  if (level === CONTAINER_MARKS.length) return nodes.map(leaf);
  const type = CONTAINER_MARKS[level];
  const out: Inline[] = [];
  let i = 0;
  while (i < nodes.length) {
    const mark = nodes[i].marks.find((m) => m.type.name === type);
    let j = i + 1;
    while (j < nodes.length && sameMark(nodes[j].marks.find((m) => m.type.name === type), mark)) j++;
    const group = nodes.slice(i, j);
    const inner = nest(group, level + 1);
    if (!mark) out.push(...inner);
    else if (type === "revision") out.push({ type: "revision", kind: mark.attrs.kind, attrs: mark.attrs.attrs, content: inner });
    else if (type === "inline_control") out.push({ type: "contentControl", settings: mark.attrs.settings, content: inner });
    else if (type === "link") {
      out.push({
        type: "link",
        ...(mark.attrs.href !== null ? { href: mark.attrs.href } : {}),
        ...(mark.attrs.anchor !== null ? { anchor: mark.attrs.anchor } : {}),
        ...(mark.attrs.tooltip !== null ? { tooltip: mark.attrs.tooltip } : {}),
        content: inner,
      });
    } else out.push({ type: "field", instr: mark.attrs.instr, content: inner });
    i = j;
  }
  return out;
}

function leaf(node: Node): Inline {
  const run = node.marks.find((m) => m.type.name === "run")?.attrs.props;
  const props = (p: Record<string, unknown> | undefined) => (p && Object.keys(p).length ? { props: p } : {});
  const a = node.attrs;
  switch (node.type.name) {
    case "text": return { type: "text", text: node.text ?? "", ...props(run) };
    case "tab": return { type: "tab", ...props(a.props) };
    case "hard_break": return { type: "break", kind: a.kind, ...(a.clear ? { clear: a.clear } : {}), ...props(a.props) };
    case "image":
      return {
        type: "image", imageId: a.imageId, width: a.width, height: a.height,
        ...(a.name !== null ? { name: a.name } : {}), ...(a.description !== null ? { description: a.description } : {}),
        ...props(a.props),
      };
    case "field_char": return { type: "fieldChar", kind: a.kind, ...props(a.props) };
    case "instr": return { type: "instr", text: a.text, ...props(a.props) };
    case "run_preserved": return { type: "runPreserved", preserved: a.preserved, ...(a.float ? { float: a.float } : {}), ...props(a.props) };
    default: return { type: "preserved", preserved: a.preserved };
  }
}

// --- headers and footers ---------------------------------------------------------------------

/** A header or footer's blocks as a standalone editor document (for editing it). */
export function storyToProseMirror(blocks: Block[], schema: Schema): Node {
  const nodes = blocks.flatMap((b) => block(b, schema));
  return schema.nodes.doc.create({ section: null }, nodes.length ? nodes : [schema.nodes.paragraph.create()]);
}

/** The blocks of an edited header or footer. */
export function storyFromProseMirror(pm: Node): Block[] {
  const out: Block[] = [];
  pm.forEach((node) => out.push(...toBlocks(node)));
  return out;
}
