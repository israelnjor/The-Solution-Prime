import { useEffect, useRef, useState } from "react";

/**
 * Passwords: asked for to open a protected document ("open"), or chosen to protect one ("set": typed
 * twice; with a way to remove one already set).
 */
export function PasswordDialog({ mode, name, hasPassword, onSubmit, onRemove, onClose }: {
  mode: "open" | "set"; name: string; hasPassword?: boolean;
  /** Resolves to an error message to show, or undefined when done. */
  onSubmit(password: string): Promise<string | undefined>;
  onRemove?(): void;
  onClose(): void;
}) {
  const [password, setPassword] = useState("");
  const [again, setAgain] = useState("");
  const [error, setError] = useState<string>();
  const [busy, setBusy] = useState(false);
  const first = useRef<HTMLInputElement>(null);
  useEffect(() => first.current?.focus(), []);
  const submit = async () => {
    if (busy) return;
    if (!password) { setError("Type a password."); return; }
    if (mode === "set" && password !== again) { setError("The two passwords are not the same."); return; }
    setBusy(true);
    setError(undefined);
    const e = await onSubmit(password);
    setBusy(false);
    if (e) { setError(e); first.current?.select(); }
  };
  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape" && !busy) onClose(); if (e.key === "Enter") void submit(); }}>
      <div className="hf-dialog password-dialog" role="dialog" aria-label={mode === "open" ? "Password" : "Protect document"}>
        <h3>{mode === "open" ? "Password" : "Protect with a password"}</h3>
        <p className="hint">{mode === "open"
          ? `${name} is protected with a password.`
          : "Anyone opening this document will need the password, in Prime Word or in Word. Keep it somewhere safe: a forgotten password cannot be recovered."}</p>
        <label>{mode === "open" ? "Password" : "New password"}
          <input ref={first} type="password" value={password} autoComplete={mode === "open" ? "current-password" : "new-password"} aria-label="Password" onChange={(e) => setPassword(e.target.value)} />
        </label>
        {mode === "set" && (
          <label>Type it again
            <input type="password" value={again} autoComplete="new-password" aria-label="Confirm password" onChange={(e) => setAgain(e.target.value)} />
          </label>
        )}
        {error && <p className="error" role="alert">{error}</p>}
        <div className="buttons">
          <button className="primary" disabled={busy} onClick={() => void submit()}>{busy ? (mode === "open" ? "Unlocking…" : "Protecting…") : mode === "open" ? "Open" : "Protect"}</button>
          {mode === "set" && hasPassword && onRemove && <button onClick={onRemove} disabled={busy} title="Save the document without a password from now on">Remove Password</button>}
          <button onClick={onClose} disabled={busy}>Cancel</button>
        </div>
      </div>
    </div>
  );
}
