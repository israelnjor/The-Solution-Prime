import type { ChartData } from "@prime/docx";
import { chartSvg } from "@prime/layout";
import { useMemo, useState } from "react";

const KINDS: [ChartData["kind"], string][] = [["column", "Column"], ["bar", "Bar"], ["line", "Line"], ["pie", "Pie"]];

/**
 * Insert → Chart (and double-click a chart to edit it): its kind, title and data — categories down the
 * side, a series per column — with a preview drawn as the page draws it. Pie charts use the first series.
 */
export function ChartDialog({ initial, editing, onOk, onClose }: {
  initial: ChartData; editing: boolean; onOk(data: ChartData): void; onClose(): void;
}) {
  const [kind, setKind] = useState(initial.kind);
  const [title, setTitle] = useState(initial.title ?? "");
  const [categories, setCategories] = useState(initial.categories);
  const [series, setSeries] = useState(initial.series.map((s) => ({ name: s.name, values: s.values.map((v) => (Number.isFinite(v) ? String(v) : "")) })) as { name: string; values: string[] }[]);
  const data: ChartData = useMemo(() => ({
    kind, ...(title.trim() ? { title: title.trim() } : {}), categories,
    series: series.map((s) => ({ name: s.name, values: categories.map((_, i) => (s.values[i] === undefined || s.values[i].trim() === "" ? NaN : Number(s.values[i]))) })),
  }), [kind, title, categories, series]);
  const svg = useMemo(() => chartSvg(data, 360, 216), [data]);
  const bad = data.series.some((s) => s.values.some((v, i) => series[data.series.indexOf(s)].values[i]?.trim() && !Number.isFinite(v)));
  const setValue = (j: number, i: number, v: string) => setSeries((ss) => ss.map((s, k) => (k === j ? { ...s, values: categories.map((_, n) => (n === i ? v : s.values[n] ?? "")) } : s)));
  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); }}>
      <div className="hf-dialog chart-dialog" role="dialog" aria-label={editing ? "Edit chart" : "Insert chart"}>
        <h3>{editing ? "Edit Chart" : "Insert Chart"}</h3>
        <div className="chart-grid">
          <div className="chart-form">
            <div className="chart-kinds" role="radiogroup" aria-label="Chart type">
              {KINDS.map(([k, label]) => (
                <button key={k} role="radio" aria-checked={kind === k} className={kind === k ? "on" : ""} onClick={() => setKind(k)}>{label}</button>
              ))}
            </div>
            <label>Title<input value={title} onChange={(e) => setTitle(e.target.value)} placeholder="Chart title (optional)" aria-label="Chart title" /></label>
            <div className="chart-data" role="table" aria-label="Chart data">
              <table>
                <thead>
                  <tr>
                    <th />
                    {series.map((s, j) => (
                      <th key={j}><input value={s.name} aria-label={`Series ${j + 1} name`} onChange={(e) => setSeries((ss) => ss.map((x, k) => (k === j ? { ...x, name: e.target.value } : x)))} /></th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {categories.map((c, i) => (
                    <tr key={i}>
                      <th><input value={c} aria-label={`Category ${i + 1}`} onChange={(e) => setCategories((cs) => cs.map((x, k) => (k === i ? e.target.value : x)))} /></th>
                      {series.map((s, j) => (
                        <td key={j}><input value={s.values[i] ?? ""} inputMode="decimal" aria-label={`${s.name}, ${c}`}
                          className={s.values[i]?.trim() && !Number.isFinite(Number(s.values[i])) ? "bad" : ""}
                          onChange={(e) => setValue(j, i, e.target.value)} /></td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div className="chart-rows">
              <button onClick={() => setCategories((cs) => [...cs, `Category ${cs.length + 1}`])}>+ Category</button>
              <button disabled={categories.length <= 1} onClick={() => { setCategories((cs) => cs.slice(0, -1)); setSeries((ss) => ss.map((s) => ({ ...s, values: s.values.slice(0, categories.length - 1) }))); }}>− Category</button>
              <button disabled={kind === "pie"} onClick={() => setSeries((ss) => [...ss, { name: `Series ${ss.length + 1}`, values: categories.map(() => "") }])}>+ Series</button>
              <button disabled={series.length <= 1} onClick={() => setSeries((ss) => ss.slice(0, -1))}>− Series</button>
            </div>
            {bad && <p className="error">Some values are not numbers; they are left out.</p>}
          </div>
          <div className="chart-preview" aria-label="Preview" dangerouslySetInnerHTML={{ __html: svg }} />
        </div>
        <div className="buttons">
          <button className="primary" onClick={() => onOk(data)}>{editing ? "Update Chart" : "Insert Chart"}</button>
          <button onClick={onClose}>Cancel</button>
        </div>
      </div>
    </div>
  );
}
