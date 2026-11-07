/**
 * Tracked formatting changes, as Word keeps them: a run's earlier character formatting inside its
 * properties (w:rPrChange), a paragraph's earlier paragraph formatting inside its own (w:pPrChange).
 * Recorded by trackchanges.ts; listed, accepted and rejected by review.ts.
 */
import { readParagraphProps, readRunProps, writeParagraphProps, writeRunProps } from "@prime/docx";
import type { ParagraphProps, RunProps } from "@prime/pdm";
import { DOMParser } from "@xmldom/xmldom";
import type { Node } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const RUN_CHANGE = "w:rPrChange";
export const PARA_CHANGE = "w:pPrChange";

type Extra = { name: string; xml: string }[];
type Props = { extra?: Extra; markRun?: unknown; [k: string]: unknown };

/** JSON with keys sorted (props compared by value). */
function stable(x: unknown): string {
  if (Array.isArray(x)) return `[${x.map(stable).join(",")}]`;
  if (x && typeof x === "object") {
    return `{${Object.keys(x).filter((k) => (x as Record<string, unknown>)[k] !== undefined).sort().map((k) => `${JSON.stringify(k)}:${stable((x as Record<string, unknown>)[k])}`).join(",")}}`;
  }
  return JSON.stringify(x);
}

/** Props without one preserved element (and without an empty extra list). */
function without<T extends Props>(props: T, name: string): T {
  const extra = (props.extra ?? []).filter((x) => x.name !== name);
  const out = { ...props, extra } as T;
  if (!extra.length) delete out.extra;
  return out;
}

const changeOf = (props: Props, name: string) => props.extra?.find((x) => x.name === name);

/** The formatting a change element holds (the formatting before the change). */
export function baselineRun(xml: string): RunProps {
  const root = new DOMParser().parseFromString(`<w:x xmlns:w="${W}">${xml}</w:x>`, "application/xml").documentElement;
  const rPr = root.getElementsByTagNameNS(W, "rPr")[0] as unknown as Element | undefined;
  return readRunProps(rPr) ?? {};
}
export function baselineParagraph(xml: string): ParagraphProps {
  const root = new DOMParser().parseFromString(`<w:x xmlns:w="${W}">${xml}</w:x>`, "application/xml").documentElement;
  const pPr = root.getElementsByTagNameNS(W, "pPr")[0] as unknown as Element | undefined;
  return readParagraphProps(pPr).props;
}

const escAttr = (s: string) => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
let changeId = 850000 + Math.floor(Math.random() * 40000);
const attrs = (author: string, date: string) => `w:id="${changeId++}" w:author="${escAttr(author)}" w:date="${date}"`;

/** A run's character formatting (its `run` mark's props; none: {}). */
export const runPropsOf = (node: Node): RunProps => (node.marks.find((m) => m.type.name === "run")?.attrs.props ?? {}) as RunProps;

/** Gives the inline content between `from` and `to` these run props (none: no run mark). */
export function setRunProps(tr: Transaction, from: number, to: number, props: RunProps) {
  const run = tr.doc.type.schema.marks.run;
  tr.removeMark(from, to, run);
  if (Object.keys(props).length) tr.addMark(from, to, run.create({ props }));
}

/** The props a format change leaves after it is accepted (true) or rejected (false). */
export function resolvedRun(props: RunProps, accept: boolean): RunProps {
  const c = changeOf(props as Props, RUN_CHANGE);
  if (!c) return props;
  return accept ? without(props as Props, RUN_CHANGE) as RunProps : baselineRun(c.xml);
}
export function resolvedParagraph(props: ParagraphProps, accept: boolean): ParagraphProps {
  const c = changeOf(props as Props, PARA_CHANGE);
  if (!c) return props;
  if (accept) return without(props as Props, PARA_CHANGE) as ParagraphProps;
  // The paragraph mark's own formatting and revision are not part of the paragraph's (they stay).
  const base = baselineParagraph(c.xml);
  return { ...base, ...(props.markRun ? { markRun: props.markRun } : {}) };
}

export const hasRunChange = (node: Node) => !!changeOf(runPropsOf(node) as Props, RUN_CHANGE);
export const hasParagraphChange = (node: Node) => !!changeOf((node.attrs.props ?? {}) as Props, PARA_CHANGE);

/**
 * Records the formatting changes a transaction made between `from` and `to` (positions in `tr.doc`):
 * each run whose character formatting differs from before keeps the earlier formatting in a
 * w:rPrChange (the first one, if it already had one: the formatting before any change); a run changed
 * back loses it. Likewise paragraphs (w:pPrChange). `back` maps `tr.doc` positions to `before`'s.
 * Text inserted in this transaction or a tracked insertion is not a formatting change.
 */
export function recordFormatting(tr: Transaction, before: Node, back: { mapResult(pos: number, assoc?: number): { pos: number; deleted: boolean } },
  from: number, to: number, author: string, date: string) {
  const ops: (() => void)[] = [];
  tr.doc.nodesBetween(from, to, (node, pos) => {
    if (node.isTextblock && node.type.name === "paragraph") {
      // (Its start, biased back: a step that only replaced its opening token does not count as deleting it.)
      const r = back.mapResult(pos, -1);
      const old = r.deleted ? null : before.nodeAt(r.pos);
      const props = (node.attrs.props ?? {}) as ParagraphProps;
      const markRev = (props.markRun?.extra ?? []).some((x) => x.name === "w:ins");
      if (old && old.type === node.type && !markRev) {
        const oldProps = (old.attrs.props ?? {}) as ParagraphProps;
        const kept = changeOf(oldProps as Props, PARA_CHANGE);
        const plain = (p: ParagraphProps) => { const { markRun: _m, ...rest } = without(p as Props, PARA_CHANGE) as ParagraphProps; return rest; };
        const baseline = kept ? plain(baselineParagraph(kept.xml)) : plain(oldProps);
        const now = plain(props);
        let next: ParagraphProps;
        if (stable(now) === stable(baseline)) next = { ...now, ...(props.markRun ? { markRun: props.markRun } : {}) };
        else {
          const item = kept ?? { name: PARA_CHANGE, xml: `<w:pPrChange ${attrs(author, date)}>${writeParagraphProps(baseline) || "<w:pPr/>"}</w:pPrChange>` };
          next = { ...now, ...(props.markRun ? { markRun: props.markRun } : {}), extra: [...(now.extra ?? []), item] };
        }
        if (stable(next) !== stable(props)) ops.push(() => tr.setNodeMarkup(pos, undefined, { ...node.attrs, props: next }));
      }
      return true;
    }
    if (!node.isInline) return true;
    if (node.marks.some((m) => m.type.name === "revision" && m.attrs.kind === "ins")) return false;
    const s = Math.max(from, pos);
    const e = Math.min(to, pos + node.nodeSize);
    if (s >= e) return false;
    const r = back.mapResult(s, 1);
    if (r.deleted) return false;
    const props = runPropsOf(node);
    // The text may have had different formatting in parts before (made bold across bold and plain).
    before.nodesBetween(r.pos, r.pos + (e - s), (oldNode, oldPos) => {
      if (!oldNode.isInline) return true;
      const os = Math.max(r.pos, oldPos);
      const oe = Math.min(r.pos + (e - s), oldPos + oldNode.nodeSize);
      if (os >= oe) return false;
      const oldProps = runPropsOf(oldNode);
      const kept = changeOf(oldProps as Props, RUN_CHANGE);
      const baseline = kept ? baselineRun(kept.xml) : without(oldProps as Props, RUN_CHANGE) as RunProps;
      const now = without(props as Props, RUN_CHANGE) as RunProps;
      let next: RunProps;
      if (stable(now) === stable(baseline)) next = now;
      else {
        const item = kept ?? { name: RUN_CHANGE, xml: `<w:rPrChange ${attrs(author, date)}>${writeRunProps(baseline) || "<w:rPr/>"}</w:rPrChange>` };
        next = { ...now, extra: [...(now.extra ?? []), item] };
      }
      if (stable(next) !== stable(props)) {
        const a = s + (os - r.pos);
        const b = s + (oe - r.pos);
        ops.push(() => setRunProps(tr, a, b, next));
      }
      return false;
    });
    return false;
  });
  for (const op of ops) op();
}
