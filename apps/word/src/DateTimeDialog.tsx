import { insertFieldResult } from "@prime/editor";
import type { EditorView } from "prosemirror-view";
import { useState } from "react";

/** Word's date pictures (the \@ switch of a DATE field). */
const FORMATS = [
  "d MMMM yyyy", "dd/MM/yyyy", "d/M/yyyy", "dddd, d MMMM yyyy", "MMMM d, yyyy", "dd MMM yyyy", "yyyy-MM-dd",
  "dd.MM.yyyy", "MMMM yyyy", "d MMM yy", "HH:mm", "h:mm am/pm", "dd/MM/yyyy HH:mm",
];
const STORE = "prime-word:datetime";

/** A date in one of Word's date pictures: d dd ddd dddd, M MM MMM MMMM, yy yyyy, h hh H HH, mm, ss, am/pm. */
export function formatDate(date: Date, picture: string): string {
  const day = (style: "long" | "short") => date.toLocaleDateString(undefined, { weekday: style });
  const month = (style: "long" | "short") => date.toLocaleDateString(undefined, { month: style });
  const pad = (n: number) => String(n).padStart(2, "0");
  const h12 = date.getHours() % 12 || 12;
  const tokens: Record<string, () => string> = {
    dddd: () => day("long"), ddd: () => day("short"), dd: () => pad(date.getDate()), d: () => String(date.getDate()),
    MMMM: () => month("long"), MMM: () => month("short"), MM: () => pad(date.getMonth() + 1), M: () => String(date.getMonth() + 1),
    yyyy: () => String(date.getFullYear()), yy: () => pad(date.getFullYear() % 100),
    HH: () => pad(date.getHours()), H: () => String(date.getHours()), hh: () => pad(h12), h: () => String(h12),
    mm: () => pad(date.getMinutes()), ss: () => pad(date.getSeconds()),
    "am/pm": () => (date.getHours() < 12 ? "am" : "pm"), "AM/PM": () => (date.getHours() < 12 ? "AM" : "PM"),
  };
  return picture.replace(/dddd|ddd|dd|d|MMMM|MMM|MM|M|yyyy|yy|HH|H|hh|h|mm|ss|am\/pm|AM\/PM/g, (t) => tokens[t]());
}

/** Insert → Date & Time: today's date or the time in a chosen format, as text or as a field Word updates. */
export function DateTimeDialog({ view, onClose }: { view: EditorView; onClose(): void }) {
  const saved = (() => { try { return JSON.parse(localStorage.getItem(STORE) ?? "{}"); } catch { return {}; } })();
  const [picture, setPicture] = useState<string>(saved.picture ?? FORMATS[0]);
  const [update, setUpdate] = useState<boolean>(!!saved.update);
  const now = new Date();
  const insert = () => {
    const text = formatDate(now, picture);
    if (update) insertFieldResult(`DATE \\@ "${picture}"`, text)(view.state, view.dispatch);
    else view.dispatch(view.state.tr.insertText(text).scrollIntoView());
    try { localStorage.setItem(STORE, JSON.stringify({ picture, update })); } catch { /* storage unavailable */ }
    onClose();
    view.focus();
  };
  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); if (e.key === "Enter") insert(); }}>
      <div className="hf-dialog datetime-dialog" role="dialog" aria-label="Date and Time">
        <h3>Date and Time</h3>
        <div className="datetime-list" role="listbox" aria-label="Formats">
          {FORMATS.map((p) => (
            <button key={p} role="option" aria-selected={p === picture} className={p === picture ? "on" : ""}
              onClick={() => setPicture(p)} onDoubleClick={() => { setPicture(p); setTimeout(insert, 0); }}>{formatDate(now, p)}</button>
          ))}
        </div>
        <label className="choice"><input type="checkbox" checked={update} onChange={(e) => setUpdate(e.target.checked)} /> Update automatically</label>
        <div className="hf-actions">
          <button onClick={onClose}>Cancel</button>
          <button className="primary" onClick={insert}>OK</button>
        </div>
      </div>
    </div>
  );
}
