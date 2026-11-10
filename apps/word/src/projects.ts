/**
 * Projects (desktop): folders on the person's disk that Prime Word keeps together — the files stay in
 * the folder, where they can be seen, backed up and synced; Prime Word keeps a list of the folders and,
 * in each, a small hidden .primeword/project.json (name, colour) that travels with it. See projects.rs.
 *
 * Self-tests supply a stand-in at globalThis.primeProjectsBackend (an in-memory project store).
 */
import { invoke } from "@tauri-apps/api/core";
import { documentDir } from "@tauri-apps/api/path";
import { open as openDialog } from "@tauri-apps/plugin-dialog";
import { isDesktop } from "./platform";

/** What Prime Word keeps in a project's folder about it. */
export interface ProjectMeta {
  version?: number;
  name: string;
  /** Card and tab colour (RRGGBB). */
  color: string;
  /** When it became a project (ms since 1970). */
  created?: number;
}

export interface ProjectSummary {
  path: string;
  exists: boolean;
  meta: ProjectMeta;
  files: number;
  /** When a file in it last changed. */
  modified: number;
}

export interface ProjectFile {
  path: string;
  /** Its folder within the project ("" at the top). */
  folder: string;
  name: string;
  size: number;
  modified: number;
}

/**
 * Project colours to choose from: soft ones first (cream, lavender, mint, sky, blush, sand, sage, mist —
 * cream is the default), then strong ones (blue, green, red, purple, amber, teal, slate, pink).
 */
export const PROJECT_COLORS = [
  "F4EEDD", "ECE5F6", "E1F1E7", "E0ECF8", "F8E4E9", "F1E5D1", "E5ECDD", "E9ECF0",
  "2B579A", "1E7D45", "B4371F", "8E44AD", "C77C02", "0E7C86", "5D6D7E", "C2185B",
];

const rgbOf = (hex: string) => [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) || 0);
const hexOf = (rgb: number[]) => rgb.map((v) => Math.round(Math.max(0, Math.min(255, v))).toString(16).padStart(2, "0")).join("");
const mix = (hex: string, toward: number, amount: number) => hexOf(rgbOf(hex).map((v) => v + (toward - v) * amount));

/**
 * How a project's colour is shown: `band` for its card's top band, swatch and dot (a strong colour as it
 * is; a soft one a shade deeper, so it shows on white), `edge` for hover borders and its tabs.
 */
export function projectTone(hex: string | undefined): { band: string; edge: string } {
  const h = /^[0-9a-f]{6}$/i.test(hex ?? "") ? hex! : PROJECT_COLORS[0];
  const [r, g, b] = rgbOf(h);
  const light = (0.299 * r + 0.587 * g + 0.114 * b) / 255 > 0.72;
  return light ? { band: `#${mix(h, 0, 0.18)}`, edge: `#${mix(h, 0, 0.32)}` } : { band: `#${h}`, edge: `#${h}` };
}

/** CSS variables for a project's colour (--project-band, --project-edge). */
export function projectStyle(hex: string | undefined): Record<string, string> {
  const t = projectTone(hex);
  return { "--project-band": t.band, "--project-edge": t.edge };
}

type Backend = (command: string, args?: Record<string, unknown>) => Promise<unknown>;
const backend = (): Backend | undefined =>
  (globalThis as { primeProjectsBackend?: Backend }).primeProjectsBackend ?? (isDesktop ? (c, a) => invoke(c, a) : undefined);

/** Whether projects can be used here (the desktop app). */
export const projectsAvailable = () => !!backend();

const call = <T>(command: string, args?: Record<string, unknown>) => {
  const b = backend();
  if (!b) return Promise.reject(new Error("Projects are available in the desktop app."));
  return b(command, args) as Promise<T>;
};

export const listProjects = () => (backend() ? call<ProjectSummary[]>("project_list") : Promise.resolve([]));
export const projectFiles = (folder: string) => call<ProjectFile[]>("project_files", { folder });
export const saveProject = (folder: string, meta: ProjectMeta) => call<void>("project_save", { folder, meta });
export const forgetProject = (folder: string) => call<void>("project_forget", { folder });
/** A path in the project for a new file (with "(2)" added when the name is taken). */
export const newPathIn = (folder: string, name: string) => call<string>("project_new_path", { folder, name });
export const openExternal = (path: string) => call<void>("project_open_external", { path });

/** Makes `folder` a project (creating it when `create`); a folder that already is one keeps its settings. */
export const addProject = (folder: string, meta: ProjectMeta, create: boolean) =>
  call<ProjectSummary>("project_add", { folder, meta: { version: 1, created: Date.now(), ...meta }, create });

const sep = (path: string) => (path.includes("\\") ? "\\" : "/");
export const joinPath = (folder: string, name: string) => `${folder.replace(/[\\/]+$/, "")}${sep(folder)}${name}`;

/** Where new projects go unless the person picks somewhere else: their Documents folder. */
export async function defaultProjectsFolder(): Promise<string> {
  const stand = (globalThis as { primeProjectsHome?: string }).primeProjectsHome;
  if (stand) return stand;
  try { return await documentDir(); } catch { return ""; }
}

/** Asks for a folder (to hold a new project, or to become one). */
export async function chooseFolder(title: string, defaultPath?: string): Promise<string | undefined> {
  const stand = (globalThis as { primeChooseFolder?: () => Promise<string | undefined> }).primeChooseFolder;
  if (stand) return stand();
  const picked = await openDialog({ directory: true, multiple: false, title, ...(defaultPath ? { defaultPath } : {}) });
  return typeof picked === "string" ? picked : undefined;
}

/** Asks for files to copy into a project, and copies them; returns their new paths. */
export async function importFiles(folder: string): Promise<string[]> {
  const stand = (globalThis as { primeChooseFiles?: () => Promise<string[]> }).primeChooseFiles;
  const picked = stand ? await stand() : await openDialog({ multiple: true, directory: false, title: "Add files to the project" });
  const paths = Array.isArray(picked) ? picked : typeof picked === "string" ? [picked] : [];
  return paths.length ? call<string[]>("project_import", { folder, paths }) : [];
}

/** What kind of file it is, for its icon and what opening it does. */
export type FileKind = "document" | "pdf" | "image" | "sheet" | "slides" | "other";
export function kindOf(name: string): FileKind {
  const ext = /\.([a-z0-9]+)$/i.exec(name)?.[1]?.toLowerCase() ?? "";
  if (["docx", "doc", "rtf", "odt", "txt"].includes(ext)) return "document";
  if (ext === "pdf") return "pdf";
  if (["png", "jpg", "jpeg", "gif", "bmp", "webp", "tif", "tiff", "heic", "svg"].includes(ext)) return "image";
  if (["xlsx", "xls", "csv", "ods"].includes(ext)) return "sheet";
  if (["pptx", "ppt", "odp"].includes(ext)) return "slides";
  return "other";
}

/** The project a file belongs to (the innermost project folder holding it), if any. */
export function projectOf(projects: ProjectSummary[], path: string | undefined): ProjectSummary | undefined {
  if (!path) return undefined;
  const norm = (p: string) => p.replace(/\//g, "\\").replace(/\\+$/, "").toLowerCase();
  const p = norm(path);
  return projects
    .filter((x) => p.startsWith(`${norm(x.path)}\\`))
    .sort((a, b) => b.path.length - a.path.length)[0];
}
