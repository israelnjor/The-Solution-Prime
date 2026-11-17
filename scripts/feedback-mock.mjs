// A stand-in for the Prime Productivity feedback API (docs/feedback-api.md), for developing and testing
// the in-app form and its outbox without the platform. Keeps reports in memory and attachments in a folder.
//
//   node scripts/feedback-mock.mjs [port] [folder]        (default 8787, a temp folder)
//   then run the app with VITE_FEEDBACK_API=http://localhost:8787
//
// GET /reports lists what arrived (with how many times each was announced, to see retries are not duplicated).
import { createServer } from "node:http";
import { mkdirSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createHash } from "node:crypto";

const port = Number(process.argv[2] ?? 8787);
const folder = process.argv[3] ?? join(tmpdir(), "prime-feedback-mock");
mkdirSync(folder, { recursive: true });
const reports = new Map(); // id -> { report, announced, uploaded: Set, complete }
const CATEGORIES = new Set(["bug", "difficulty", "idea", "general"]);
const MAX_FILE = 10 * 1024 * 1024;

const body = (req) => new Promise((resolve) => { const parts = []; req.on("data", (c) => parts.push(c)); req.on("end", () => resolve(Buffer.concat(parts))); });
const send = (res, status, data) => {
  res.writeHead(status, { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" });
  res.end(JSON.stringify(data));
};
const invalid = (r) => {
  if (!r || typeof r !== "object") return "not a report";
  if (!/^[A-Za-z0-9-]{8,64}$/.test(r.id ?? "")) return "bad id";
  if (!CATEGORIES.has(r.category)) return "bad category";
  if (typeof r.description !== "string" || r.description.trim().length < 3 || r.description.length > 10000) return "bad description";
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(r.email ?? "")) return "bad email";
  if (!Array.isArray(r.attachments) || r.attachments.length > 5) return "bad attachments";
  if (r.attachments.some((a) => !/^a\d{1,3}$/.test(a.file) || !(a.size > 0 && a.size <= MAX_FILE))) return "attachment too large or badly named";
  return "";
};

createServer(async (req, res) => {
  if (req.method === "OPTIONS") {
    res.writeHead(204, { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Methods": "POST, PUT, GET, OPTIONS", "Access-Control-Allow-Headers": "Content-Type, X-Prime-Install, X-Prime-Product" });
    return res.end();
  }
  const url = new URL(req.url, `http://localhost:${port}`);
  if (req.method === "GET" && url.pathname === "/reports") {
    return send(res, 200, [...reports.values()].map((r) => ({ id: r.report.id, category: r.report.category, description: r.report.description, email: r.report.email,
      announced: r.announced, attachments: r.report.attachments.map((a) => `${a.name} (${a.size})`), uploaded: [...r.uploaded], complete: r.complete, diagnostics: r.report.diagnostics })));
  }
  if (req.method === "POST" && url.pathname === "/v1/feedback") {
    let report;
    try { report = JSON.parse((await body(req)).toString("utf8")); } catch { return send(res, 400, { error: "not JSON" }); }
    const why = invalid(report);
    if (why) return send(res, 400, { error: why });
    // The same id again: the same record (a retry), never a second one.
    const known = reports.get(report.id) ?? { report, announced: 0, uploaded: new Set(), complete: false };
    known.announced++;
    reports.set(report.id, known);
    const uploads = report.attachments.filter((a) => !known.uploaded.has(a.file))
      .map((a) => ({ file: a.file, url: `http://localhost:${port}/upload/${report.id}/${a.file}` }));
    return send(res, 200, { status: known.complete ? "complete" : "received", uploads });
  }
  const up = /^\/upload\/([A-Za-z0-9-]+)\/(a\d+)$/.exec(url.pathname);
  if (req.method === "PUT" && up) {
    const known = reports.get(up[1]);
    const meta = known?.report.attachments.find((a) => a.file === up[2]);
    if (!known || !meta) return send(res, 404, { error: "unknown upload" });
    const bytes = await body(req);
    if (bytes.length !== meta.size || createHash("sha256").update(bytes).digest("hex") !== meta.sha256) return send(res, 400, { error: "file does not match" });
    writeFileSync(join(folder, `${up[1]}-${meta.name.replace(/[^A-Za-z0-9._-]/g, "_")}`), bytes);
    known.uploaded.add(up[2]);
    return send(res, 200, { ok: true });
  }
  const done = /^\/v1\/feedback\/([A-Za-z0-9-]+)\/complete$/.exec(url.pathname);
  if (req.method === "POST" && done) {
    const known = reports.get(done[1]);
    if (!known) return send(res, 404, { error: "unknown report" });
    if (known.report.attachments.some((a) => !known.uploaded.has(a.file))) return send(res, 409, { error: "attachments missing" });
    if (!known.complete) { known.complete = true; console.log(`[feedback] ${known.report.category}: ${known.report.description.slice(0, 60)} (${known.report.email})`); }
    return send(res, 200, { status: "complete" });
  }
  send(res, 404, { error: "not found" });
}).listen(port, () => console.log(`feedback mock on http://localhost:${port} (attachments in ${folder})`));
