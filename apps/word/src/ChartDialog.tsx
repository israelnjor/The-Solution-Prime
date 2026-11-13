import type { ChartData } from "@prime/docx";
import { CHART_PALETTES, chartSvg } from "@prime/layout";
import { useMemo, useState } from "react";

const KINDS: [ChartData["kind"], string][] = [
  ["column", "Column"], ["bar", "Bar"], ["line", "Line"], ["area", "Area"], ["pie", "Pie"], ["doughnut", "Doughnut"], ["scatter", "Scatter"],
];
/** Kinds whose series can be stacked. */
const STACKS: ChartData["kind"][] = ["column", "bar", "area"];

/**
 * Insert → Chart (and double-click a chart to edit it): its kind, title and data — categories down the
 * side, a series per column — with a preview drawn as the page draws it. Pie and doughnut charts use the
 * first series; a scatter chart's first column holds the x values.
 */
export function ChartDialog({ initial, editing, onOk, onClose }: {
  initial: ChartData; editing: boolean; onOk(data: ChartData): void; onClose(): void;
}) {
  const [kind, setKind] = useState(initial.kind);
  const [title, setTitle] = useState(initial.title ?? "");
  const [stacked, setStacked] = useState(!!initial.stacked);
  const [legend, setLegend] = useState(initial.legend ?? "bottom");
  const [labels, setLabels] = useState(!!initial.labels);
  const [gridlines, setGridlines] = useState(initial.gridlines !== false);
  const [colors, setColors] = useState(initial.colors);
  const [categories, setCategories] = useState(initial.categories);
  const [series, setSeries] = useState(initial.series.map((s) => ({ name: s.name, ...(s.color ? { color: s.color } : {}), values: s.values.map((v) => (Number.isFinite(v) ? String(v) : "")) })) as { name: string; color?: string; values: string[] }[]);
  const round = kind === "pie" || kind === "doughnut";
  const data: ChartData = useMemo(() => ({
    kind, ...(title.trim() ? { title: title.trim() } : {}), ...(stacked && STACKS.includes(kind) ? { stacked: true } : {}),
    ...(legend !== "bottom" ? { legend } : {}), ...(labels ? { labels: true } : {}), ...(!gridlines && !round ? { gridlines: false } : {}),
    ...(colors?.length ? { colors } : {}), categories,
    series: series.map((s) => ({ name: s.name, ...(s.color ? { color: s.color } : {}), values: categories.map((_, i) => (s.values[i] === undefined || s.values[i].trim() === "" ? NaN : Number(s.values[i]))) })),
  }), [kind, title, stacked, legend, labels, gridlines, colors, round, categories, series]);
  // The colour set in use: the one chosen, or the one the series' own colours match.
  const own = series.map((s) => s.color).filter(Boolean).join(",");
  const activePalette = CHART_PALETTES.find((p) => (colors?.length ? p.colors.slice(0, colors.length).join(",") === colors.join(",") : own ? p.colors.slice(0, series.length).join(",") === own : p.id === "office"))?.id;
  const choosePalette = (c: string[]) => { setColors(c); setSeries((ss) => ss.map(({ color: _, ...s }) => s)); };
  const svg = useMemo(() => chartSvg(data, 360, 216), [data]);
  const badX = kind === "scatter" && categories.some((c) => !c.trim() || !Number.isFinite(Number(c)));
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
            {STACKS.includes(kind) && (
              <label className="choice" title="Each category's series on top of each other, showing their total">
                <input type="checkbox" checked={stacked} onChange={(e) => setStacked(e.target.checked)} /> Stacked
              </label>
            )}
            <fieldset className="chart-style">
              <legend>Style</legend>
              <div className="chart-palettes" role="radiogroup" aria-label="Colours">
                {CHART_PALETTES.map((p) => (
                  <button key={p.id} role="radio" aria-checked={activePalette === p.id} className={activePalette === p.id ? "on" : ""} title={p.name} aria-label={`${p.name} colours`}
                    onClick={() => choosePalette(p.colors)}>
                    {p.colors.slice(0, 5).map((c) => <span key={c} style={{ background: `#${c}` }} />)}
                  </button>
                ))}
              </div>
              <div className="chart-options">
                <label>Legend
                  <select value={legend} onChange={(e) => setLegend(e.target.value as typeof legend)} aria-label="Legend">
                    <option value="bottom">Bottom</option><option value="top">Top</option><option value="right">Right</option><option value="none">None</option>
                  </select>
                </label>
                <label className="choice"><input type="checkbox" checked={labels} onChange={(e) => setLabels(e.target.checked)} /> Data labels</label>
                <label className="choice" title={round ? "Pie and doughnut charts have no gridlines" : undefined}>
                  <input type="checkbox" checked={gridlines && !round} disabled={round} onChange={(e) => setGridlines(e.target.checked)} /> Gridlines
                </label>
              </div>
            </fieldset>
            <label>Title<input value={title} onChange={(e) => setTitle(e.target.value)} placeholder="Chart title (optional)" aria-label="Chart title" /></label>
            <div className="chart-data" role="table" aria-label="Chart data">
              <table>
                <thead>
                  <tr>
                    <th>{kind === "scatter" ? "X values" : ""}</th>
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
              <button disabled={round} onClick={() => setSeries((ss) => [...ss, { name: `Series ${ss.length + 1}`, values: categories.map(() => "") }])}>+ Series</button>
              <button disabled={series.length <= 1} onClick={() => setSeries((ss) => ss.slice(0, -1))}>− Series</button>
            </div>
            {bad && <p className="error">Some values are not numbers; they are left out.</p>}
            {badX && <p className="error">A scatter chart's x values (the first column) must be numbers.</p>}
            {round && series.length > 1 && <p className="hint">A {kind} chart shows the first series only.</p>}
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
