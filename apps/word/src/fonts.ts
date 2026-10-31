/**
 * Office "cloud fonts" (Aptos, the default since 2023, and others) are downloaded by Office into
 * its own cache instead of being installed, so web views cannot see them. On the desktop we read
 * the ones a document uses from that cache and register them with the page.
 */
import { readDir, readFile } from "@tauri-apps/plugin-fs";
import { BaseDirectory } from "@tauri-apps/api/path";
import { isDesktop } from "./platform";

const CLOUD = "Microsoft/FontCache/4/CloudFonts";
const loaded = new Set<string>();

/** Weight and italic flag from the font's OS/2 table. */
function faceOf(bytes: Uint8Array): { weight: string; style: string } {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const tables = dv.getUint16(4);
  for (let i = 0; i < tables; i++) {
    const r = 12 + 16 * i;
    if (String.fromCharCode(...bytes.subarray(r, r + 4)) === "OS/2") {
      const at = dv.getUint32(r + 8);
      return { weight: String(dv.getUint16(at + 4)), style: dv.getUint16(at + 62) & 1 ? "italic" : "normal" };
    }
  }
  return { weight: "400", style: "normal" };
}

/**
 * Registers the cloud fonts whose family names appear in `usage` (style names, theme fonts, runs).
 * Returns true when anything new was loaded, so the caller can re-measure text.
 */
export async function loadCloudFonts(usage: string): Promise<boolean> {
  if (!isDesktop) return false;
  let families: string[];
  try {
    families = (await readDir(CLOUD, { baseDir: BaseDirectory.LocalData })).filter((e) => e.isDirectory).map((e) => e.name);
  } catch {
    return false; // no Office cloud fonts on this machine
  }
  let added = false;
  for (const family of families) {
    if (loaded.has(family) || !usage.includes(family)) continue;
    loaded.add(family);
    const files = await readDir(`${CLOUD}/${family}`, { baseDir: BaseDirectory.LocalData });
    for (const f of files.filter((e) => /\.(ttf|otf)$/i.test(e.name))) {
      const bytes = await readFile(`${CLOUD}/${family}/${f.name}`, { baseDir: BaseDirectory.LocalData });
      const face = new FontFace(family, bytes, faceOf(bytes));
      document.fonts.add(await face.load());
      added = true;
    }
  }
  return added;
}
