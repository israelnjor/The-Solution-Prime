import { useEffect, useRef } from "react";

/** An open document; `project` and `color` when it is in one of the person's projects. */
export interface TabInfo { id: number; name: string; dirty: boolean; project?: string; color?: string }

/**
 * The open documents, as tabs above the ribbon (shown once more than one is open): click one to switch
 * to it; × or a middle click closes it; + starts a new document. Ctrl+Tab / Ctrl+Shift+Tab move between
 * them and Ctrl+W closes the current one (App.tsx).
 */
export function DocTabs({ tabs, active, onSelect, onClose, onNew }: {
  tabs: TabInfo[]; active: number; onSelect(id: number): void; onClose(id: number): void; onNew(): void;
}) {
  const strip = useRef<HTMLDivElement>(null);
  // The current tab in sight (many tabs scroll sideways).
  useEffect(() => {
    strip.current?.querySelector<HTMLElement>(`[data-tab="${active}"]`)?.scrollIntoView({ block: "nearest", inline: "nearest" });
  }, [active]);
  return (
    <div className="doc-tabs" role="tablist" aria-label="Open documents" ref={strip}>
      {tabs.map((t) => (
        <div key={t.id} data-tab={t.id} role="tab" aria-selected={t.id === active} tabIndex={t.id === active ? 0 : -1}
          className={`doc-tab${t.id === active ? " on" : ""}${t.color ? " in-project" : ""}`}
          title={`${t.name}${t.project ? ` — project ${t.project}` : ""}${t.dirty ? " (unsaved changes)" : ""}`}
          style={t.color ? { ["--project" as string]: `#${t.color}` } : undefined}
          onClick={() => onSelect(t.id)}
          onAuxClick={(e) => { if (e.button === 1) { e.preventDefault(); onClose(t.id); } }}
          onMouseDown={(e) => { if (e.button === 1) e.preventDefault(); }}
          onKeyDown={(e) => {
            if (e.key === "ArrowRight" || e.key === "ArrowLeft") {
              const i = tabs.findIndex((x) => x.id === t.id);
              const next = tabs[(i + (e.key === "ArrowRight" ? 1 : tabs.length - 1)) % tabs.length];
              onSelect(next.id);
              strip.current?.querySelector<HTMLElement>(`[data-tab="${next.id}"]`)?.focus();
            }
            if (e.key === "Delete") onClose(t.id);
          }}>
          <span className="doc-tab-name">{t.name.replace(/\.docx$/i, "")}</span>
          {t.dirty && <span className="doc-tab-dirty" aria-label="unsaved changes">•</span>}
          <button className="doc-tab-close" aria-label={`Close ${t.name}`} title="Close (Ctrl+W)"
            onClick={(e) => { e.stopPropagation(); onClose(t.id); }}>×</button>
        </div>
      ))}
      <button className="doc-tab-new" onClick={onNew} title="New document" aria-label="New document">+</button>
    </div>
  );
}
