import { AUTOFORMAT_DEFAULTS, type AutoFormatOptions } from "@prime/editor";
import { useState } from "react";

const STORE = "prime-word:autocorrect";

/** The AutoCorrect options saved on this computer (the defaults where none are saved). */
export function loadAutoCorrect(): AutoFormatOptions {
  try {
    const saved = localStorage.getItem(STORE);
    if (saved) return { ...AUTOFORMAT_DEFAULTS, ...JSON.parse(saved) };
  } catch { /* storage unavailable */ }
  return { ...AUTOFORMAT_DEFAULTS };
}

const RULES: [keyof Omit<AutoFormatOptions, "custom">, string][] = [
  ["quotes", "\"Straight quotes\" with “smart quotes”"],
  ["dashes", "Hyphens (--) with dashes (– or —)"],
  ["symbols", "Symbols: (c) → ©, (tm) → ™, ... → …, --> → →"],
  ["corrections", "Common misspellings (teh → the, recieve → receive)"],
  ["capitalise", "Capitalise the first letter of sentences"],
  ["twoCapitals", "Correct TWo INitial CApitals"],
  ["lists", "Automatic lists (\"1. \" or \"- \" at the start of a line)"],
];

/** AutoCorrect options: which automatic changes happen as you type, and your own replacements. */
export function AutoCorrectDialog({ onChange, onClose }: { onChange(o: AutoFormatOptions): void; onClose(): void }) {
  const [opts, setOpts] = useState(loadAutoCorrect);
  const [replace, setReplace] = useState("");
  const [by, setBy] = useState("");
  const update = (next: AutoFormatOptions) => {
    setOpts(next);
    onChange(next);
    try { localStorage.setItem(STORE, JSON.stringify(next)); } catch { /* storage unavailable */ }
  };
  const custom = Object.entries(opts.custom ?? {});
  const add = () => {
    if (!replace.trim() || !by) return;
    update({ ...opts, custom: { ...opts.custom, [replace.trim()]: by } });
    setReplace("");
    setBy("");
  };

  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); }}>
      <div className="hf-dialog autocorrect" role="dialog" aria-label="AutoCorrect">
        <h3>AutoCorrect</h3>
        <fieldset><legend>As you type, replace</legend>
          {RULES.map(([k, label]) => (
            <label key={k} className="choice"><input type="checkbox" checked={opts[k]} onChange={(e) => update({ ...opts, [k]: e.target.checked })} /> {label}</label>
          ))}
        </fieldset>
        <fieldset><legend>Your own replacements</legend>
          <form className="ac-add" onSubmit={(e) => { e.preventDefault(); add(); }}>
            <input value={replace} onChange={(e) => setReplace(e.target.value)} placeholder="Replace (e.g. pw)" aria-label="Replace" />
            <input value={by} onChange={(e) => setBy(e.target.value)} placeholder="With (e.g. Prime Word)" aria-label="With" />
            <button type="submit" disabled={!replace.trim() || !by}>Add</button>
          </form>
          {custom.length > 0 && (
            <ul className="ac-list">
              {custom.map(([k, v]) => (
                <li key={k}><code>{k}</code> → {v}
                  <button className="link" aria-label={`Remove ${k}`} onClick={() => {
                    const rest = { ...opts.custom };
                    delete rest[k];
                    update({ ...opts, custom: rest });
                  }}>Remove</button>
                </li>
              ))}
            </ul>
          )}
        </fieldset>
        <p className="hint">Undo (Ctrl+Z), or Backspace straight away, takes back a change you did not want.</p>
        <div className="hf-actions"><button className="primary" onClick={onClose}>Done</button></div>
      </div>
    </div>
  );
}
