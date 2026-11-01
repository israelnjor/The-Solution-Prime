import { useEffect, useRef, useState } from "react";
import { forgetRecent, isDesktop, type RecentDocument, recentDocuments } from "./platform";

const folderOf = (path: string) => path.slice(0, Math.max(path.lastIndexOf("\\"), path.lastIndexOf("/")));

/** Recently opened or saved documents. Missing files are shown greyed out, with a way to remove them. */
export function RecentList({ onOpen }: { onOpen(path: string): void }) {
  const [items, setItems] = useState<RecentDocument[]>();
  const refresh = () => void recentDocuments().then(setItems).catch(() => setItems([]));
  useEffect(refresh, []);

  if (!items) return null;
  if (!items.length) return <p className="recent-empty">Documents you open or save will appear here.</p>;
  return (
    <ul className="recent-list">
      {items.map((d) => (
        <li key={d.path} className={d.exists ? "" : "missing"}>
          <button className="recent-open" disabled={!d.exists} onClick={() => onOpen(d.path)}
            title={d.exists ? d.path : `${d.path} (not found: moved, renamed or deleted)`}>
            <strong>{d.name}</strong>
            <span>{d.exists ? folderOf(d.path) : "Not found"}</span>
          </button>
          <button className="recent-remove" title="Remove from list" aria-label={`Remove ${d.name} from the list`}
            onClick={() => void forgetRecent(d.path).then(refresh)}>×</button>
        </li>
      ))}
    </ul>
  );
}

/** Ribbon button: "Recent ▾" with the list in a drop-down (desktop only). */
export function RecentMenu({ onOpen }: { onOpen(path: string): void }) {
  const [open, setOpen] = useState(false);
  const box = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!open) return;
    const close = (e: MouseEvent) => { if (!box.current?.contains(e.target as Node)) setOpen(false); };
    const esc = (e: KeyboardEvent) => { if (e.key === "Escape") setOpen(false); };
    window.addEventListener("mousedown", close);
    window.addEventListener("keydown", esc);
    return () => { window.removeEventListener("mousedown", close); window.removeEventListener("keydown", esc); };
  }, [open]);
  if (!isDesktop) return null;
  return (
    <div className="recent-menu" ref={box}>
      <button onClick={() => setOpen((o) => !o)} aria-expanded={open} title="Recent documents">Recent ▾</button>
      {open && (
        <div className="recent-popup" role="menu">
          <RecentList onOpen={(p) => { setOpen(false); onOpen(p); }} />
        </div>
      )}
    </div>
  );
}
