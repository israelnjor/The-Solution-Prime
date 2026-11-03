import {
  addressBlock, blankDocument, insertPart, type Part, type PartAlign, partFromSelection, type PartPlace, picturePart, signatureBlock,
} from "@prime/editor";
import type { PrimeDocument } from "@prime/pdm";
import type { EditorView } from "prosemirror-view";
import { useEffect, useMemo, useRef, useState } from "react";
import { deleteEntry, type LibraryEntry, listEntries, newId, saveEntry } from "./library";
import { PageThumb } from "./MyTemplates";
import type { SavedSignature } from "./SignatureDialog";
import { pngOf } from "./signature";

/** A saved block: a part (paragraphs and their pictures) and what kind of block it is. */
export type SavedBlock = LibraryEntry<{ part: Part; kind: Maker | "selection" }>;
type Maker = "signature" | "address" | "stamp";

const today = () => new Date().toLocaleDateString(undefined, { day: "numeric", month: "long", year: "numeric" });
const CLOSINGS = ["", "Yours faithfully,", "Yours sincerely,", "Kind regards,", "Best regards,", "Signed,"];
const KIND_LABEL: Record<SavedBlock["data"]["kind"], string> = { signature: "Signature block", address: "Address", stamp: "Stamp", selection: "Saved text" };

/** A small page showing a part, as it would look aligned. */
function PartThumb({ part, align, width = 160 }: { part: Part; align?: PartAlign; width?: number }) {
  const doc = useMemo<PrimeDocument>(() => {
    const d = blankDocument();
    d.images = { ...part.images };
    d.sections[0].blocks = structuredClone(part.blocks).map((b) => (b.type === "paragraph" && align ? { ...b, props: { ...b.props, align } } : b));
    return d;
  }, [part, align]);
  return <PageThumb doc={doc} width={width} />;
}

/**
 * Blocks: signature blocks, addresses, stamps and saved text, kept on this computer; inserted at the
 * cursor or into the header or footer, aligned left, centred or right.
 */
export function BlocksDialog({ view, base, onClose, initial }: { view: EditorView; base: () => PrimeDocument; onClose(): void; initial?: "selection" }) {
  const [saved, setSaved] = useState<SavedBlock[]>();
  const [making, setMaking] = useState<Maker | "selection" | undefined>(initial);
  const [place, setPlace] = useState<PartPlace>("cursor");
  const [align, setAlign] = useState<PartAlign>("left");
  const selection = useMemo(() => partFromSelection(base)(view.state), [view, base]);
  const load = () => void listEntries<SavedBlock["data"]>("blocks").then(setSaved).catch(() => setSaved([]));
  useEffect(load, []);

  const insert = (part: Part) => {
    insertPart(part, base, place, align)(view.state, view.dispatch, view);
    onClose();
    view.focus();
  };
  const save = async (name: string, kind: SavedBlock["data"]["kind"], part: Part) => {
    await saveEntry("blocks", { id: newId(name), name, created: Date.now(), data: { part, kind } } satisfies SavedBlock);
    setMaking(undefined);
    load();
  };

  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); }}>
      <div className="hf-dialog blocks" role="dialog" aria-label="Blocks">
        <h3>{making ? `New ${KIND_LABEL[making].toLowerCase()}` : "Blocks"}</h3>
        {making === "signature" && <SignatureBlockMaker onSave={(name, part) => void save(name, "signature", part)} onCancel={() => setMaking(undefined)} />}
        {making === "address" && <AddressMaker onSave={(name, part) => void save(name, "address", part)} onCancel={() => setMaking(undefined)} />}
        {making === "stamp" && <StampMaker onSave={(name, part) => void save(name, "stamp", part)} onCancel={() => setMaking(undefined)} />}
        {making === "selection" && selection && <SelectionSaver part={selection} onSave={(name) => void save(name, "selection", selection)} onCancel={() => setMaking(undefined)} />}
        {!making && (
          <>
            <div className="blocks-place">
              <label>Insert <select value={place} onChange={(e) => setPlace(e.target.value as PartPlace)}>
                <option value="cursor">at the cursor</option>
                <option value="header">in the header</option>
                <option value="footer">in the footer</option>
              </select></label>
              <div className="blocks-align" role="radiogroup" aria-label="Alignment">
                {(["left", "center", "right"] as const).map((a) => (
                  <button key={a} role="radio" aria-checked={align === a} className={align === a ? "on" : ""} onClick={() => setAlign(a)}>
                    {a === "left" ? "Left" : a === "center" ? "Centre" : "Right"}
                  </button>
                ))}
              </div>
            </div>
            {saved && !saved.length && <p className="hint">Blocks you make appear here, ready to insert into any document.</p>}
            <div className="blocks-list">
              {saved?.map((b) => (
                <div key={b.id} className="template-card mine">
                  <button className="mine-open" onClick={() => insert(b.data.part)} title={`Insert ${b.name}`}>
                    <PartThumb part={b.data.part} align={align} />
                    <strong>{b.name}</strong>
                    <span>{KIND_LABEL[b.data.kind]}</span>
                  </button>
                  <button className="mine-delete" aria-label={`Delete ${b.name}`}
                    onClick={() => { if (window.confirm(`Delete the block "${b.name}"?`)) void deleteEntry("blocks", b.id).then(load); }}>×</button>
                </div>
              ))}
            </div>
            <div className="hf-actions">
              <span className="blocks-new">New:
                <button onClick={() => setMaking("signature")}>Signature block…</button>
                <button onClick={() => setMaking("address")}>Address…</button>
                <button onClick={() => setMaking("stamp")}>Stamp…</button>
                <button disabled={!selection} title={selection ? "Keep the selected text as a block" : "Select some text first"} onClick={() => setMaking("selection")}>Save selection…</button>
              </span>
              <button onClick={onClose}>Close</button>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

function Actions({ ready, name, setName, onSave, onCancel }: { ready: boolean; name: string; setName(n: string): void; onSave(): void; onCancel(): void }) {
  return (
    <div className="hf-actions">
      <label className="sig-name">Name <input value={name} onChange={(e) => setName(e.target.value)} /></label>
      <button onClick={onCancel}>Back</button>
      <button className="primary" disabled={!ready || !name.trim()} onClick={onSave}>Save block</button>
    </div>
  );
}

function SignatureBlockMaker({ onSave, onCancel }: { onSave(name: string, part: Part): void; onCancel(): void }) {
  const [signatures, setSignatures] = useState<SavedSignature[]>([]);
  const [signatureId, setSignatureId] = useState("");
  const [closing, setClosing] = useState("Yours faithfully,");
  const [person, setPerson] = useState("");
  const [title, setTitle] = useState("");
  const [organisation, setOrganisation] = useState("");
  const [line, setLine] = useState(true);
  const [dated, setDated] = useState(false);
  const [name, setName] = useState("My signature block");
  useEffect(() => void listEntries<SavedSignature["data"]>("signatures").then((list) => {
    setSignatures(list);
    if (list[0]) setSignatureId(list[0].id);
  }).catch(() => {}), []);
  const sig = signatures.find((s) => s.id === signatureId)?.data;
  const part = useMemo(() => signatureBlock({
    closing: closing || undefined, signature: sig, line, name: person || "Your name", title: title || undefined,
    organisation: organisation || undefined, date: dated ? today() : undefined,
  }), [closing, sig, line, person, title, organisation, dated]);
  return (
    <div className="blocks-maker">
      <form className="letterhead-form" onSubmit={(e) => e.preventDefault()}>
        <label>Closing <select value={closing} onChange={(e) => setClosing(e.target.value)}>{CLOSINGS.map((c) => <option key={c} value={c}>{c || "(none)"}</option>)}</select></label>
        <label>Signature <select value={signatureId} onChange={(e) => setSignatureId(e.target.value)}>
          <option value="">(room to sign by hand)</option>
          {signatures.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
        </select></label>
        <label>Name<input value={person} onChange={(e) => setPerson(e.target.value)} placeholder="e.g. Ama Mensah" autoFocus /></label>
        <label>Title (optional)<input value={title} onChange={(e) => setTitle(e.target.value)} placeholder="e.g. Head Librarian" /></label>
        <label>Organisation (optional)<input value={organisation} onChange={(e) => setOrganisation(e.target.value)} /></label>
        <label className="choice"><input type="checkbox" checked={line} onChange={(e) => setLine(e.target.checked)} /> A line to sign on</label>
        <label className="choice"><input type="checkbox" checked={dated} onChange={(e) => setDated(e.target.checked)} /> Today's date</label>
      </form>
      <figure className="letterhead-preview"><PartThumb part={part} width={240} /><figcaption>Preview</figcaption></figure>
      <Actions ready={!!person.trim()} name={name} setName={setName} onSave={() => onSave(name.trim(), part)} onCancel={onCancel} />
    </div>
  );
}

function AddressMaker({ onSave, onCancel }: { onSave(name: string, part: Part): void; onCancel(): void }) {
  const [lines, setLines] = useState("");
  const [bold, setBold] = useState(true);
  const [name, setName] = useState("");
  const list = lines.split("\n");
  const part = useMemo(() => addressBlock(list, { bold }), [lines, bold]);
  return (
    <div className="blocks-maker">
      <form className="letterhead-form" onSubmit={(e) => e.preventDefault()}>
        <label>Address (one line each)<textarea rows={6} value={lines} onChange={(e) => setLines(e.target.value)} autoFocus
          placeholder={"Accra Community Library\n12 Independence Avenue\nAccra, Ghana"} /></label>
        <label className="choice"><input type="checkbox" checked={bold} onChange={(e) => setBold(e.target.checked)} /> First line in bold</label>
      </form>
      <figure className="letterhead-preview"><PartThumb part={part} width={240} /><figcaption>Preview</figcaption></figure>
      <Actions ready={!!lines.trim()} name={name || list[0]?.trim() || ""} setName={setName}
        onSave={() => onSave((name || list[0]).trim(), part)} onCancel={onCancel} />
    </div>
  );
}

/** A rubber stamp, drawn: round (text around the edge) or oblong (text in lines), with a centre line and date. */
export function drawStamp(c: HTMLCanvasElement, s: { shape: "round" | "oblong"; outer: string; centre: string; date: boolean; color: string }) {
  const ctx = c.getContext("2d")!;
  const round = s.shape === "round";
  c.width = round ? 600 : 800;
  c.height = round ? 600 : 360;
  ctx.clearRect(0, 0, c.width, c.height);
  ctx.strokeStyle = ctx.fillStyle = s.color;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  const centre = s.centre.toUpperCase();
  const date = s.date ? today().toUpperCase() : "";
  if (round) {
    const r = 290;
    ctx.lineWidth = 10;
    ctx.beginPath(); ctx.arc(300, 300, r, 0, Math.PI * 2); ctx.stroke();
    ctx.lineWidth = 4;
    ctx.beginPath(); ctx.arc(300, 300, r - 78, 0, Math.PI * 2); ctx.stroke();
    // The outer text around the top of the ring, letter by letter (smaller until it fits).
    const text = s.outer.toUpperCase();
    const rr = r - 42;
    let size = 46;
    ctx.font = `bold ${size}px Arial`;
    while (ctx.measureText(text).width / rr > Math.PI * 1.6 && size > 20) ctx.font = `bold ${(size -= 2)}px Arial`;
    let a = -Math.PI / 2 - ctx.measureText(text).width / rr / 2;
    for (const ch of text) {
      const w = ctx.measureText(ch).width / rr;
      ctx.save();
      ctx.translate(300 + Math.cos(a + w / 2) * rr, 300 + Math.sin(a + w / 2) * rr);
      ctx.rotate(a + w / 2 + Math.PI / 2);
      ctx.fillText(ch, 0, 0);
      ctx.restore();
      a += w;
    }
    ctx.font = "40px Arial";
    ctx.fillText("★", 300, 300 + rr); // the bottom of the ring, as on a seal
    ctx.font = "bold 58px Arial";
    ctx.fillText(centre, 300, date ? 270 : 300, 380);
    if (date) { ctx.font = "bold 34px Arial"; ctx.fillText(date, 300, 340, 360); }
  } else {
    ctx.lineWidth = 10;
    ctx.strokeRect(8, 8, 784, 344);
    ctx.lineWidth = 4;
    ctx.strokeRect(26, 26, 748, 308);
    ctx.font = "bold 44px Arial";
    ctx.fillText(s.outer.toUpperCase(), 400, 90, 700);
    ctx.font = "bold 76px Arial";
    ctx.fillText(centre, 400, date ? 185 : 215, 700);
    if (date) { ctx.font = "bold 40px Arial"; ctx.fillText(date, 400, 275, 700); }
  }
}

function StampMaker({ onSave, onCancel }: { onSave(name: string, part: Part): void; onCancel(): void }) {
  const [shape, setShape] = useState<"round" | "oblong">("round");
  const [outer, setOuter] = useState("");
  const [centre, setCentre] = useState("Received");
  const [date, setDate] = useState(true);
  const [color, setColor] = useState("#1c3f94");
  const [name, setName] = useState("");
  const canvas = useRef<HTMLCanvasElement>(null);
  useEffect(() => { if (canvas.current) drawStamp(canvas.current, { shape, outer: outer || "Your organisation", centre, date, color }); }, [shape, outer, centre, date, color]);
  const make = (): Part => {
    const png = pngOf(canvas.current!);
    return picturePart(png, { width: shape === "round" ? 100 : 150, name: name || "Stamp", float: true });
  };
  return (
    <div className="blocks-maker">
      <form className="letterhead-form" onSubmit={(e) => e.preventDefault()}>
        <fieldset className="row"><legend>Shape</legend>
          <label className="choice"><input type="radio" checked={shape === "round"} onChange={() => setShape("round")} /> Round</label>
          <label className="choice"><input type="radio" checked={shape === "oblong"} onChange={() => setShape("oblong")} /> Oblong</label>
        </fieldset>
        <label>Organisation<input value={outer} onChange={(e) => setOuter(e.target.value)} placeholder="e.g. Accra Community Library" autoFocus /></label>
        <label>Centre text<input value={centre} onChange={(e) => setCentre(e.target.value)} placeholder="e.g. Received, Approved, Paid" /></label>
        <label className="choice"><input type="checkbox" checked={date} onChange={(e) => setDate(e.target.checked)} /> Today's date</label>
        <label className="choice">Colour <input type="color" value={color} onChange={(e) => setColor(e.target.value)} /></label>
      </form>
      <figure className="letterhead-preview"><canvas ref={canvas} className="stamp-preview" /><figcaption>Preview</figcaption></figure>
      <Actions ready={!!outer.trim()} name={name || `${centre} stamp`} setName={setName} onSave={() => onSave((name || `${centre} stamp`).trim(), make())} onCancel={onCancel} />
    </div>
  );
}

function SelectionSaver({ part, onSave, onCancel }: { part: Part; onSave(name: string): void; onCancel(): void }) {
  const first = part.blocks.find((b) => b.type === "paragraph");
  const guess = first && first.type === "paragraph" ? first.content.map((i) => (i.type === "text" ? i.text : "")).join("").slice(0, 40) : "";
  const [name, setName] = useState(guess || "Saved text");
  return (
    <div className="blocks-maker">
      <p className="hint">The selected text, with its formatting and pictures, is kept as a block.</p>
      <figure className="letterhead-preview"><PartThumb part={part} width={240} /><figcaption>Preview</figcaption></figure>
      <Actions ready name={name} setName={setName} onSave={() => onSave(name.trim())} onCancel={onCancel} />
    </div>
  );
}
