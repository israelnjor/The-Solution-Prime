import { useEffect, useRef, useState } from "react";
import {
  type Attachment, CATEGORIES, deliver, type Delivery, type FeedbackCategory, type FeedbackContext, keepReport, MAX_FILE, MAX_FILES,
  onPending, rememberedSender,
} from "./feedback";

const size = (n: number) => (n < 1024 * 1024 ? `${Math.max(1, Math.round(n / 1024))} KB` : `${(n / 1024 / 1024).toFixed(1)} MB`);
const PLACEHOLDER: Record<FeedbackCategory, string> = {
  bug: "What happened, and what did you expect to happen?",
  difficulty: "What were you trying to do, and what made it hard?",
  idea: "What would you like Prime Word to do?",
  general: "Tell us what you think.",
};

/**
 * Send feedback: a category, a sentence or two, and the tester's email (remembered), sent in about
 * thirty seconds. Screenshots (pasted or chosen) and the open document are attached only when the
 * tester adds them. The report is kept on disk first, so it is never lost: offline, it goes later.
 */
export function FeedbackDialog({ context, currentDocument, onClose }: {
  context: FeedbackContext;
  /** The open document as a Word file, if one is open (attached only when the tester ticks it). */
  currentDocument?: { name: string; make(): Promise<Uint8Array> };
  onClose(): void;
}) {
  const remembered = rememberedSender();
  const [category, setCategory] = useState<FeedbackCategory>();
  const [description, setDescription] = useState("");
  const [email, setEmail] = useState(remembered.email);
  const [followUp, setFollowUp] = useState(remembered.followUp);
  const [files, setFiles] = useState<Attachment[]>([]);
  const [withDocument, setWithDocument] = useState(false);
  const [problem, setProblem] = useState("");
  const [state, setState] = useState<"editing" | "sending" | Delivery>("editing");
  const [waiting, setWaiting] = useState(0);
  const text = useRef<HTMLTextAreaElement>(null);
  // Sending already (set at once: a double click must not make a second report).
  const busy = useRef(false);
  useEffect(() => onPending(setWaiting), []);

  const emailOk = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim());
  const ready = !!category && description.trim().length >= 3 && emailOk;

  const add = async (list: File[]) => {
    const next: Attachment[] = [];
    for (const f of list) {
      if (files.length + next.length >= MAX_FILES) { setProblem(`At most ${MAX_FILES} files.`); break; }
      if (f.size > MAX_FILE) { setProblem(`${f.name} is ${size(f.size)}: files can be up to 10 MB.`); continue; }
      next.push({ name: f.name || "screenshot.png", type: f.type || "application/octet-stream", bytes: new Uint8Array(await f.arrayBuffer()) });
    }
    if (next.length) { setFiles((fs) => [...fs, ...next]); setProblem(""); }
  };
  // A screenshot pasted (Win+Shift+S, then Ctrl+V): attached.
  const paste = (e: React.ClipboardEvent) => {
    const images = [...e.clipboardData.files].filter((f) => f.type.startsWith("image/"));
    if (!images.length) return;
    e.preventDefault();
    void add(images.map((f, i) => new File([f], `screenshot-${files.length + i + 1}.png`, { type: f.type })));
  };
  const choose = () => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/*";
    input.multiple = true;
    input.onchange = () => void add([...(input.files ?? [])]);
    input.click();
  };

  const send = async () => {
    if (!ready || !category || busy.current) return;
    busy.current = true;
    setState("sending");
    setProblem("");
    try {
      const attachments = [...files];
      if (withDocument && currentDocument) {
        const bytes = await currentDocument.make();
        if (bytes.length <= MAX_FILE) attachments.push({ name: currentDocument.name, type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", bytes });
      }
      // Kept on disk first: from here the report cannot be lost.
      const id = await keepReport({ category, description, email, followUp, attachments, context });
      setState(await deliver(id));
    } catch (e) {
      busy.current = false;
      setState("editing");
      setProblem(`Prime Word could not keep your feedback: ${(e as Error).message}`);
    }
  };

  const done = state === "sent" || state === "offline" || state === "unconfigured" || state === "refused";
  return (
    <div className="hf-overlay" onKeyDown={(e) => {
      if (e.key === "Escape" && state !== "sending") onClose();
      if (e.key === "Enter" && (e.ctrlKey || e.metaKey) && !done) { e.preventDefault(); void send(); }
    }}>
      <div className="hf-dialog feedback-dialog" role="dialog" aria-label="Send feedback" onPaste={paste}>
        <h3>Send feedback</h3>
        {done ? (
          <div className="feedback-result" role="status">
            {state === "sent" && <p><strong>Thank you!</strong> Your feedback was sent.</p>}
            {state === "offline" && <p>You're offline. Your feedback has been saved and will automatically be sent when you're connected again.</p>}
            {state === "unconfigured" && <p>Your feedback has been saved. It will be sent automatically once sending feedback is switched on in an update.</p>}
            {state === "refused" && <p>Your feedback has been saved, but it could not be accepted as it is. It stays on this computer; please tell us another way.</p>}
            <div className="hf-actions"><button className="primary" onClick={onClose} autoFocus>Done</button></div>
          </div>
        ) : (
          <>
            <div className="feedback-kinds" role="radiogroup" aria-label="What kind of feedback">
              {CATEGORIES.map((c) => (
                <button key={c.id} role="radio" aria-checked={category === c.id} className={category === c.id ? "on" : ""} title={c.hint}
                  onClick={() => { setCategory(c.id); text.current?.focus(); }}>{c.label}</button>
              ))}
            </div>
            <textarea ref={text} rows={4} value={description} onChange={(e) => setDescription(e.target.value)} aria-label="Your feedback"
              placeholder={category ? PLACEHOLDER[category] : "Choose a kind of feedback above, then tell us in a sentence or two."} autoFocus />
            <div className="feedback-row">
              <label className="field">Your email <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} placeholder="name@example.com" aria-label="Your email"
                aria-invalid={!!email && !emailOk} /></label>
              <label className="choice"><input type="checkbox" checked={followUp} onChange={(e) => setFollowUp(e.target.checked)} /> You may contact me about this</label>
            </div>
            <div className="feedback-files">
              <button onClick={choose} disabled={files.length >= MAX_FILES}>Add a screenshot…</button>
              <span className="hint">or press Win+Shift+S, then Ctrl+V here</span>
              {currentDocument && (
                <label className="choice"><input type="checkbox" checked={withDocument} onChange={(e) => setWithDocument(e.target.checked)} />
                  Attach the document I'm working on ({currentDocument.name})</label>
              )}
              {files.length > 0 && (
                <ul>{files.map((f, i) => (
                  <li key={i}>{f.name} <span className="hint">{size(f.bytes.length)}</span>
                    <button className="link" aria-label={`Remove ${f.name}`} onClick={() => setFiles((fs) => fs.filter((_, j) => j !== i))}>Remove</button></li>
                ))}</ul>
              )}
            </div>
            <details className="feedback-included">
              <summary>Also included</summary>
              <p className="hint">Prime Word's version, your Windows version, the time, your language and screen size, and what was open
                (the home screen or a document, its number of pages and words). Never your document's text unless you attach it.</p>
            </details>
            {problem && <p className="hint error">{problem}</p>}
            <div className="hf-actions">
              {waiting > 0 && <span className="hint feedback-waiting">{waiting === 1 ? "1 earlier report" : `${waiting} earlier reports`} waiting to be sent</span>}
              <button onClick={onClose} disabled={state === "sending"}>Cancel</button>
              <button className="primary" onClick={() => void send()} disabled={!ready || state === "sending"} title="Send (Ctrl+Enter)">
                {state === "sending" ? "Sending…" : "Send"}
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
