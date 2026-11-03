import { pictureAt, replacePicture } from "@prime/editor";
import type { PrimeDocument } from "@prime/pdm";
import type { EditorView } from "prosemirror-view";
import { useEffect, useRef, useState } from "react";

type Box = { x: number; y: number; w: number; h: number }; // fractions (0–1) of the rotated picture
const FULL: Box = { x: 0, y: 0, w: 1, h: 1 };
const MAX = { w: 560, h: 400 }; // preview size, px

/**
 * Crop & Rotate: the picture itself is edited (a new image replaces it), so it looks the same in Word,
 * in the PDF and on screen. Rotation in quarter turns; the crop frame is dragged by its edges and corners.
 */
export function PictureEditor({ view, base, onClose }: { view: EditorView; base: () => PrimeDocument; onClose(): void }) {
  const pic = pictureAt(view.state);
  const res = pic && !pic.kept ? base().images[pic.node.attrs.imageId] : undefined;
  const [img, setImg] = useState<HTMLImageElement>();
  const [turns, setTurns] = useState(0);
  const [crop, setCrop] = useState<Box>(FULL);
  const canvas = useRef<HTMLCanvasElement>(null);
  const frame = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!res) return;
    const el = new Image();
    el.onload = () => setImg(el);
    el.src = `data:${res.contentType};base64,${res.data}`;
  }, [res]);

  const sideways = turns % 2 === 1;
  const size = img ? { w: sideways ? img.naturalHeight : img.naturalWidth, h: sideways ? img.naturalWidth : img.naturalHeight } : { w: 1, h: 1 };
  const scale = Math.min(MAX.w / size.w, MAX.h / size.h, 1);
  const view_ = { w: Math.round(size.w * scale), h: Math.round(size.h * scale) };

  /** The rotated picture drawn on `c` at `k` canvas pixels per picture pixel. */
  const drawRotated = (c: HTMLCanvasElement, k: number) => {
    if (!img) return;
    c.width = Math.round(size.w * k);
    c.height = Math.round(size.h * k);
    const ctx = c.getContext("2d")!;
    ctx.save();
    ctx.translate(c.width / 2, c.height / 2);
    ctx.rotate((turns * Math.PI) / 2);
    ctx.drawImage(img, (-img.naturalWidth * k) / 2, (-img.naturalHeight * k) / 2, img.naturalWidth * k, img.naturalHeight * k);
    ctx.restore();
  };

  useEffect(() => { if (canvas.current) drawRotated(canvas.current, scale); });

  // Dragging the frame's edges and corners (or the whole frame).
  const drag = (handle: string) => (e: React.PointerEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const start = { x: e.clientX, y: e.clientY, crop };
    const move = (m: PointerEvent) => {
      const dx = (m.clientX - start.x) / view_.w;
      const dy = (m.clientY - start.y) / view_.h;
      let { x, y, w, h } = start.crop;
      const min = 0.05;
      if (handle === "move") {
        x = Math.min(Math.max(0, x + dx), 1 - w);
        y = Math.min(Math.max(0, y + dy), 1 - h);
      } else {
        if (handle.includes("w")) { const nx = Math.min(Math.max(0, x + dx), x + w - min); w += x - nx; x = nx; }
        if (handle.includes("e")) w = Math.min(Math.max(min, w + dx), 1 - x);
        if (handle.includes("n")) { const ny = Math.min(Math.max(0, y + dy), y + h - min); h += y - ny; y = ny; }
        if (handle.includes("s")) h = Math.min(Math.max(min, h + dy), 1 - y);
      }
      setCrop({ x, y, w, h });
    };
    const up = () => { window.removeEventListener("pointermove", move); window.removeEventListener("pointerup", up); };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  };

  const rotate = (by: number) => {
    setTurns((t) => (t + by + 4) % 4);
    setCrop(FULL); // a crop frame does not survive a quarter turn meaningfully
  };

  const apply = () => {
    if (!img || !pic || !res) return;
    // Full resolution: rotate, then cut out the frame.
    const whole = document.createElement("canvas");
    drawRotated(whole, 1);
    const sx = Math.round(crop.x * whole.width);
    const sy = Math.round(crop.y * whole.height);
    const sw = Math.max(1, Math.round(crop.w * whole.width));
    const sh = Math.max(1, Math.round(crop.h * whole.height));
    const out = document.createElement("canvas");
    out.width = sw;
    out.height = sh;
    out.getContext("2d")!.drawImage(whole, sx, sy, sw, sh, 0, 0, sw, sh);
    const type = res.contentType === "image/jpeg" ? "image/jpeg" : "image/png";
    const data = out.toDataURL(type, 0.92).split(",")[1];
    // Same scale as before: points per picture pixel.
    const ptPerPx = pic.node.attrs.width / img.naturalWidth;
    replacePicture(base, { data, contentType: type, width: sw * ptPerPx, height: sh * ptPerPx })(view.state, view.dispatch);
    onClose();
    view.focus();
  };

  if (!pic || pic.kept) return null;
  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); }}>
      <div className="hf-dialog picedit" role="dialog" aria-label="Crop and rotate picture">
        <h3>Crop &amp; Rotate</h3>
        <div className="picedit-stage" style={{ width: view_.w, height: view_.h }}>
          <canvas ref={canvas} style={{ width: view_.w, height: view_.h }} />
          <div ref={frame} className="picedit-frame" onPointerDown={drag("move")}
            style={{ left: crop.x * view_.w, top: crop.y * view_.h, width: crop.w * view_.w, height: crop.h * view_.h }}>
            {["n", "s", "e", "w", "nw", "ne", "sw", "se"].map((h) => (
              <span key={h} className={`picedit-handle h-${h}`} onPointerDown={drag(h)} />
            ))}
          </div>
        </div>
        <div className="picedit-tools">
          <button onClick={() => rotate(-1)} title="Rotate left 90°">⟲ Rotate left</button>
          <button onClick={() => rotate(1)} title="Rotate right 90°">⟳ Rotate right</button>
          <button onClick={() => setCrop(FULL)}>Reset crop</button>
          <span className="hint">Drag the frame's edges or corners to crop.</span>
        </div>
        <div className="hf-actions">
          <button onClick={onClose}>Cancel</button>
          <button className="primary" onClick={apply} disabled={!img}>Apply</button>
        </div>
      </div>
    </div>
  );
}
