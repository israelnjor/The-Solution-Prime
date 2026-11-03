/**
 * Tab stops on screen. Browsers put a tab at the next multiple of a fixed width; Word puts it at the
 * paragraph's next tab stop (left, centre, right, decimal), else at the next default stop every 36pt.
 * After drawing, each tab in a paragraph that has tabs gets the width that takes the text to its stop —
 * the same rule the layout engine uses, so the screen and the pages agree.
 */
import { Resolver, type TabStop } from "@prime/layout";
import type { Paragraph, ParagraphProps, PrimeDocument } from "@prime/pdm";
import type { Node } from "prosemirror-model";
import { type Command, Plugin } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

const DEFAULT_TAB = 36;
const PX = 96 / 72;
const LEADERS = ["dot", "hyphen", "underscore", "middleDot", "heavy"];

function hasTab(node: Node): boolean {
  let found = false;
  node.forEach((n) => { if (n.type.name === "tab") found = true; });
  return found;
}

/** Sizes the tabs of one paragraph (its DOM `p`), first to last. */
function sizeTabs(view: EditorView, node: Node, p: HTMLElement, resolver: Resolver) {
  if (p.classList.contains("pw-righttab")) return; // a contents line: laid out by CSS
  const container = p.parentElement;
  if (!container) return;
  const rp = resolver.paragraph({ type: "paragraph", props: node.attrs.props, content: [] } as Paragraph);
  const tabs = Array.from(p.querySelectorAll<HTMLElement>(":scope > .pw-tab, :scope > a > .pw-tab, :scope > span > .pw-tab"));
  if (!tabs.length) return;
  const box = view.dom.getBoundingClientRect();
  const cssWidth = parseFloat(getComputedStyle(view.dom).width);
  const zoom = cssWidth ? box.width / cssWidth : 1;
  const cs = getComputedStyle(container);
  const left = container.getBoundingClientRect().left + (parseFloat(cs.paddingLeft) + parseFloat(cs.borderLeftWidth || "0")) * zoom;
  for (let i = 0; i < tabs.length; i++) {
    const tab = tabs[i];
    const start = (tab.getBoundingClientRect().left - left) / zoom / PX; // pt from the column's left edge
    const stop = rp.tabs.find((t) => t.pos > start + 0.01 && t.align !== "bar");
    let end: number;
    if (stop && stop.align !== "left") {
      // The text up to the next tab (or the end of the line) ends at, or centres on, the stop.
      const range = document.createRange();
      range.setStartAfter(tab);
      if (tabs[i + 1]) range.setEndBefore(tabs[i + 1]);
      else range.setEnd(p, p.childNodes.length);
      const w = range.getBoundingClientRect().width / zoom / PX;
      end = Math.max(start, stop.align === "center" ? stop.pos - w / 2 : stop.pos - w);
    } else end = stop ? stop.pos : (Math.floor(start / DEFAULT_TAB + 1e-6) + 1) * DEFAULT_TAB;
    const width = Math.max(0, end - start);
    const leader = stop?.leader && LEADERS.includes(stop.leader) ? `pw-leader-${stop.leader}` : "";
    tab.style.cssText = `display:inline-block;width:${width.toFixed(2)}pt;overflow:hidden;vertical-align:baseline`;
    tab.className = `pw-tab${leader ? ` ${leader}` : ""}`;
  }
}

/** Tab stops on screen, for the paragraphs that have tabs (each again only when it changes). */
export function tabStops(doc: () => PrimeDocument): Plugin {
  return new Plugin({
    view(view) {
      let done = new WeakSet<Node>();
      let resolver = new Resolver(doc());
      const run = (all: boolean) => {
        if (all) { done = new WeakSet(); resolver = new Resolver(doc()); }
        view.state.doc.descendants((node, pos) => {
          if (node.type.name !== "paragraph") return true;
          if (done.has(node) || !hasTab(node)) return false;
          done.add(node);
          const p = view.nodeDOM(pos);
          if (p instanceof HTMLElement) sizeTabs(view, node, p, resolver);
          return false;
        });
      };
      // After the first drawing (fonts and styles applied).
      const first = setTimeout(() => run(true), 0);
      const fonts = () => run(true);
      document.fonts?.addEventListener?.("loadingdone", fonts);
      return {
        update(v, prev) { if (v.state.doc !== prev.doc) run(false); },
        destroy() { clearTimeout(first); document.fonts?.removeEventListener?.("loadingdone", fonts); },
      };
    },
  });
}

/**
 * Sets the selected paragraphs' tab stops to `stops` (pt from the column's left edge): written as the
 * paragraph's own w:tabs, with "clear" entries for any stops its style has that are not wanted.
 */
export function setTabStops(doc: () => PrimeDocument, stops: TabStop[]): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const resolver = new Resolver(doc());
    const tr = state.tr;
    state.doc.nodesBetween(state.selection.from, state.selection.to, (node, pos) => {
      if (node.type.name !== "paragraph") return true;
      const props = node.attrs.props as ParagraphProps;
      const others = props.extra?.filter((e) => e.name !== "w:tabs") ?? [];
      // What the style alone gives this paragraph.
      const inherited = resolver.paragraph({ type: "paragraph", props: { ...props, extra: others }, content: [] } as Paragraph).tabs;
      const near = (a: number, b: number) => Math.abs(a - b) < 0.5;
      const tw = (pt: number) => Math.round(pt * 20);
      const items = [
        ...inherited.filter((t) => !stops.some((s) => near(s.pos, t.pos) && s.align === t.align && s.leader === t.leader))
          .map((t) => `<w:tab w:val="clear" w:pos="${tw(t.pos)}"/>`),
        ...stops.filter((s) => !inherited.some((t) => near(s.pos, t.pos) && s.align === t.align && s.leader === t.leader))
          .map((s) => `<w:tab w:val="${s.align}"${s.leader ? ` w:leader="${s.leader}"` : ""} w:pos="${tw(s.pos)}"/>`),
      ];
      const extra = items.length ? [...others, { name: "w:tabs", xml: `<w:tabs>${items.join("")}</w:tabs>` }] : others;
      const next: ParagraphProps = { ...props, extra };
      if (!extra.length) delete next.extra;
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, props: next });
      return false;
    });
    dispatch(tr.scrollIntoView());
    return true;
  };
}
