import { useEffect, useRef, useState } from "react";
import { deleteEntry, type LibraryEntry, listEntries, newId, saveEntry } from "./library";
import { cleanSignature, cropToInk, type InkColour, pngOf, toCanvas } from "./signature";

/** A saved signature: a transparent PNG and its pixel size. */
export type SavedSignature = LibraryEntry<{ png: string; width: number; height: number }>;

type Source = "photo" | "draw" | "type";
const SCRIPT_FONTS = ["Segoe Script", "Ink Free", "Lucida Handwriting", "Brush Script MT", "Segoe Print"];

/** A picture file the person chooses (a photo or scan of their signature). */
function choosePicture(): Promise<HTMLImageElement | undefined> {
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/*";
    input.onchange = () => {
      const f = input.files?.[0];
      if (!f) return resolve(undefined);
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = () => resolve(undefined);
      img.src = URL.createObjectURL(f);
    };
    input.click();
  });
}

/**
 * Signatures: the ones saved on this computer (insert, delete), and making a new one from a photo or
 * scan (cleaned up automatically), by drawing it, or by typing it in a handwriting font.
 */
export function SignatureDialog({ onInsert, onClose }: { onInsert(s: SavedSignature): void; onClose(): void }) {
  const [saved, setSaved] = useState<SavedSignature[]>();
  const [creating, setCreating] = useState(false);
  const load = () => void listEntries<SavedSignature["data"]>("signatures").then((list) => {
    setSaved(list);
    if (!list.length) setCreating(true);
  }).catch(() => setSaved([]));
  useEffect(load, []);

  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); }}>
      <div className="hf-dialog signatures" role="dialog" aria-label="Signatures">
        <h3>{creating ? "New signature" : "Signatures"}</h3>
        {creating ? (
          <SignatureMaker onSaved={() => { setCreating(false); load(); }} onCancel={() => (saved?.length ? setCreating(false) : onClose())} />
        ) : (
          <>
            <p className="hint">Saved on this computer only. Click a signature to insert it; then drag it onto the signature line.</p>
            <div className="sig-list">
              {saved?.map((s) => (
                <div key={s.id} className="sig-card">
                  <button className="sig-open" onClick={() => onInsert(s)} title={`Insert ${s.name}`}>
                    <img src={`data:image/png;base64,${s.data.png}`} alt={s.name} />
                    <span>{s.name}</span>
                  </button>
                  <button className="mine-delete" aria-label={`Delete ${s.name}`}
                    onClick={() => { if (window.confirm(`Delete the signature "${s.name}"?`)) void deleteEntry("signatures", s.id).then(load); }}>×</button>
                </div>
              ))}
            </div>
            <div className="hf-actions">
              <button onClick={onClose}>Close</button>
              <button className="primary" onClick={() => setCreating(true)}>New signature…</button>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

function SignatureMaker({ onSaved, onCancel }: { onSaved(): void; onCancel(): void }) {
  const [source, setSource] = useState<Source>("photo");
  const [photo, setPhoto] = useState<HTMLImageElement>();
  const [turns, setTurns] = useState(0);
  const [sensitivity, setSensitivity] = useState(0);
  const [colour, setColour] = useState<InkColour>("original");
  const [typed, setTyped] = useState("");
  const [font, setFont] = useState(SCRIPT_FONTS[0]);
  const [name, setName] = useState("My signature");
  const [result, setResult] = useState<HTMLCanvasElement>();
  const [warning, setWarning] = useState("");
  const pad = useRef<HTMLCanvasElement>(null);
  const preview = useRef<HTMLCanvasElement>(null);
  const strokes = useRef(0);

  // Photo: clean up whenever the picture or a setting changes.
  useEffect(() => {
    if (source !== "photo" || !photo) return;
    const cleaned = cleanSignature(toCanvas(photo, turns), { sensitivity, colour });
    setResult(cleaned.canvas);
    setWarning(cleaned.coverage < 0.005 ? "Hardly any ink was found: try a higher sensitivity, or a sharper photo on plain paper." :
      cleaned.coverage > 0.45 ? "A lot of the picture was kept: try a lower sensitivity, or a photo with plain paper around the signature." : "");
  }, [source, photo, turns, sensitivity, colour]);

  // Typed: draw the name in the chosen handwriting font.
  useEffect(() => {
    if (source !== "type") return;
    if (!typed.trim()) { setResult(undefined); return; }
    const c = document.createElement("canvas");
    c.width = 1200;
    c.height = 300;
    const ctx = c.getContext("2d")!;
    ctx.font = `140px "${font}"`;
    ctx.fillStyle = colour === "blue" ? "#102d8c" : "#141419";
    ctx.textBaseline = "middle";
    ctx.fillText(typed, 20, 150, 1160);
    setResult(cropToInk(c));
    setWarning("");
  }, [source, typed, font, colour]);

  // Drawn: a pad to sign on with a mouse, pen or finger.
  const draw = (e: React.PointerEvent<HTMLCanvasElement>) => {
    const c = pad.current!;
    const ctx = c.getContext("2d")!;
    const r = c.getBoundingClientRect();
    const at = (ev: { clientX: number; clientY: number }) => [((ev.clientX - r.left) * c.width) / r.width, ((ev.clientY - r.top) * c.height) / r.height] as const;
    c.setPointerCapture(e.pointerId);
    ctx.lineCap = ctx.lineJoin = "round";
    ctx.strokeStyle = colour === "blue" ? "#102d8c" : "#141419";
    ctx.beginPath();
    ctx.moveTo(...at(e));
    const move = (ev: PointerEvent) => {
      // Pen pressure, where the device reports it, makes the line thicker.
      ctx.lineWidth = 3 + (ev.pressure && ev.pointerType === "pen" ? ev.pressure * 5 : 2);
      ctx.lineTo(...at(ev));
      ctx.stroke();
    };
    const up = () => {
      c.removeEventListener("pointermove", move);
      c.removeEventListener("pointerup", up);
      strokes.current++;
      setResult(cropToInk(c));
    };
    c.addEventListener("pointermove", move);
    c.addEventListener("pointerup", up);
  };
  const clearPad = () => {
    const c = pad.current;
    c?.getContext("2d")!.clearRect(0, 0, c.width, c.height);
    strokes.current = 0;
    setResult(undefined);
  };

  // The result on a checkerboard (it is transparent).
  useEffect(() => {
    const p = preview.current;
    if (!p) return;
    const ctx = p.getContext("2d")!;
    ctx.clearRect(0, 0, p.width, p.height);
    if (!result) return;
    const k = Math.min(p.width / result.width, p.height / result.height);
    ctx.drawImage(result, (p.width - result.width * k) / 2, (p.height - result.height * k) / 2, result.width * k, result.height * k);
  }, [result]);

  const save = async () => {
    if (!result) return;
    const png = pngOf(result);
    await saveEntry("signatures", { id: newId(name), name: name.trim() || "Signature", created: Date.now(), data: { png: png.data, width: png.width, height: png.height } });
    onSaved();
  };

  return (
    <div className="sig-maker">
      <div className="hf-tabs">
        {(["photo", "draw", "type"] as const).map((s) => (
          <button key={s} className={source === s ? "on" : ""} onClick={() => { setSource(s); setResult(undefined); setWarning(""); }}>
            {s === "photo" ? "From a photo or scan" : s === "draw" ? "Draw" : "Type"}
          </button>
        ))}
      </div>
      {source === "photo" && (
        <div className="sig-controls">
          <button onClick={() => void choosePicture().then((img) => img && setPhoto(img))}>{photo ? "Choose another picture…" : "Choose a photo or scan…"}</button>
          {photo && (
            <>
              <label>Sensitivity <input type="range" min={-1} max={1} step={0.1} value={sensitivity} onChange={(e) => setSensitivity(Number(e.target.value))} /></label>
              <button onClick={() => setTurns((t) => (t + 1) % 4)} title="Rotate a quarter turn">⟳ Rotate</button>
            </>
          )}
          <p className="hint">Sign on plain white paper and take the photo straight from above, in good light. The paper and shadows are removed automatically.</p>
        </div>
      )}
      {source === "draw" && (
        <div className="sig-controls">
          <canvas ref={pad} className="sig-pad" width={900} height={300} onPointerDown={draw} />
          <button onClick={clearPad}>Clear</button>
        </div>
      )}
      {source === "type" && (
        <div className="sig-controls">
          <input value={typed} onChange={(e) => setTyped(e.target.value)} placeholder="Type your name" autoFocus />
          <select value={font} onChange={(e) => setFont(e.target.value)}>{SCRIPT_FONTS.map((f) => <option key={f} style={{ fontFamily: `"${f}"` }}>{f}</option>)}</select>
        </div>
      )}
      <div className="sig-controls">
        <label>Ink <select value={colour} onChange={(e) => setColour(e.target.value as InkColour)}>
          {source === "photo" && <option value="original">As in the photo</option>}
          <option value="black">Black</option>
          <option value="blue">Blue</option>
        </select></label>
      </div>
      <canvas ref={preview} className="sig-preview" width={600} height={200} aria-label="Signature preview" />
      {warning && <p className="sig-warning">{warning}</p>}
      <p className="hint">A signature picture is a simple electronic signature: it shows who signed, but anyone with the file could copy it. Keep signed documents as PDFs you send directly.</p>
      <div className="hf-actions">
        <label className="sig-name">Name <input value={name} onChange={(e) => setName(e.target.value)} /></label>
        <button onClick={onCancel}>Cancel</button>
        <button className="primary" disabled={!result} onClick={() => void save()}>Save signature</button>
      </div>
    </div>
  );
}
