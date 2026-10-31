import { type Template, TEMPLATES } from "@prime/editor";

/** File → New: pick a starting point. */
export function TemplatePicker({ onPick, onClose }: { onPick(t: Template): void; onClose(): void }) {
  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); }} onClick={onClose}>
      <div className="hf-dialog templates" role="dialog" aria-label="New document" onClick={(e) => e.stopPropagation()}>
        <h3>New document</h3>
        <div className="template-grid">
          {TEMPLATES.map((t) => (
            <button key={t.id} className="template-card" onClick={() => onPick(t)} autoFocus={t.id === "blank"}>
              <span className={`template-thumb thumb-${t.id}`} aria-hidden="true">
                <i /><i /><i /><i /><i />
              </span>
              <strong>{t.name}</strong>
              <span>{t.description}</span>
            </button>
          ))}
        </div>
        <div className="hf-actions"><button onClick={onClose}>Cancel</button></div>
      </div>
    </div>
  );
}
