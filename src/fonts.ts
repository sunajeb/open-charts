// Node font loader: finds the font files of any family on this computer and builds Measurers.
// Lives outside src/core so the engine stays free of file-system access.
//
// Where fonts are looked for (fontDirs): the system's and the user's font folders, and Microsoft
// Office's own fonts, which PowerPoint draws with but the system does not list:
//   macOS    /System/Library/Fonts (+ Supplemental), /Library/Fonts, ~/Library/Fonts,
//            /Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts (Calibri, Aptos, …),
//            ~/Library/Group Containers/UBF8T346G9.Office/FontCache (cloud fonts Office downloads;
//            macOS may deny other apps access to it, then it is skipped)
//   Windows  %WINDIR%\Fonts, %LOCALAPPDATA%\Microsoft\Windows\Fonts (fonts installed per user),
//            %LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts (cloud fonts Office downloads)
//   Linux    /usr/share/fonts, /usr/local/share/fonts, ~/.local/share/fonts, ~/.fonts
// and last open-charts' own Liberation Sans (vendor/fonts, metric-compatible with Arial).
// $OPEN_CHARTS_FONT_DIRS (path-list) replaces the platform's folders (tests, unusual set-ups).
// Files are read only, never changed.
//
// A family is matched against the files' name tables (family / typographic family / full name, in
// every language), so cloud fonts with hashed file names are found too. The index of names is kept
// in the cache directory and rebuilt when a font folder changes (its modification time).
//
// Parsing a font's metrics takes milliseconds, most of the CLI's start-up after Node itself, so the
// metrics measuring needs (advance widths per code point, line height) are cached as JSON, keyed by
// the font file's path, face, weight, size and modification time: in $OPEN_CHARTS_CACHE, else
// $XDG_CACHE_HOME/open-charts (~/.cache/open-charts).
// The caches are an optimisation only: unreadable or unwritable, fonts are read as before.
import { closeSync, existsSync, mkdirSync, openSync, readdirSync, readFileSync, readSync, renameSync, statSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { homedir } from 'node:os';
import { delimiter, join, posix, win32 } from 'node:path';
import { fileURLToPath } from 'node:url';
import { fontMetrics, measurerFromMetrics, type FontMetrics, type Measurer } from './core/text.ts';
import { readFaces, type FaceInfo, type ReadAt } from './core/sfnt.ts';
import { fontSet } from './core/fonts.ts';

// Windows keeps its fonts in %WINDIR%\Fonts (C:\Windows\Fonts on almost every PC).
const WINDOWS_FONTS = `${process.env.WINDIR ?? process.env.SystemRoot ?? 'C:\\Windows'}\\Fonts`;
const VENDOR_FONTS = (() => { try { return fileURLToPath(new URL('../vendor/fonts', import.meta.url)); } catch { return ''; } })();
// Arial or a metric-compatible substitute (Liberation Sans, Arimo) gives widths that match PowerPoint.
// Found by path first, so the default font costs no index.
const REGULAR = [
  '/System/Library/Fonts/Supplemental/Arial.ttf',
  `${WINDOWS_FONTS}\\arial.ttf`,
  '/usr/share/fonts/truetype/msttcorefonts/Arial.ttf',
  '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
  '/usr/share/fonts/liberation-sans/LiberationSans-Regular.ttf',
  join(VENDOR_FONTS, 'LiberationSans-Regular.ttf'),
];
const BOLD = [
  '/System/Library/Fonts/Supplemental/Arial Bold.ttf',
  `${WINDOWS_FONTS}\\arialbd.ttf`,
  '/usr/share/fonts/truetype/msttcorefonts/Arial_Bold.ttf',
  '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf',
  '/usr/share/fonts/liberation-sans/LiberationSans-Bold.ttf',
  join(VENDOR_FONTS, 'LiberationSans-Bold.ttf'),
];

/** The folders fonts are looked for in, in order of preference (see the top of this file). */
export function fontDirs(platform = process.platform, env = process.env, home = homedir()): string[] {
  if (env.OPEN_CHARTS_FONT_DIRS) return env.OPEN_CHARTS_FONT_DIRS.split(delimiter).filter(Boolean);
  const own = VENDOR_FONTS ? [VENDOR_FONTS] : [];
  // Paths follow the target platform's rules, not the host's (the tests ask for every platform).
  if (platform === 'darwin') {
    return [
      '/System/Library/Fonts', '/Library/Fonts', posix.join(home, 'Library/Fonts'),
      '/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts',
      posix.join(home, 'Library/Group Containers/UBF8T346G9.Office/FontCache'),
      ...own,
    ];
  }
  if (platform === 'win32') {
    const local = env.LOCALAPPDATA ?? win32.join(home, 'AppData\\Local');
    const win = `${env.WINDIR ?? env.SystemRoot ?? 'C:\\Windows'}\\Fonts`;
    return [win, `${local}\\Microsoft\\Windows\\Fonts`, `${local}\\Microsoft\\FontCache\\4\\CloudFonts`, ...own];
  }
  return ['/usr/share/fonts', '/usr/local/share/fonts', posix.join(home, '.local/share/fonts'), posix.join(home, '.fonts'), ...own];
}

function load(p: string): ArrayBuffer {
  const buf = readFileSync(p);
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
}

const cacheDir = () => process.env.OPEN_CHARTS_CACHE || join(process.env.XDG_CACHE_HOME || join(homedir(), '.cache'), 'open-charts');

function writeCache(file: string, data: unknown) {
  try {
    mkdirSync(cacheDir(), { recursive: true });
    const tmp = `${file}.${process.pid}.tmp`;
    writeFileSync(tmp, JSON.stringify(data));
    renameSync(tmp, file); // atomic: concurrent CLIs never read half a file
  } catch { /* read-only home: fine, just slower next time */ }
}

/**
 * Metrics of one font file (face `index` of a collection; `wght`: an instance of a variable font),
 * from the cache when it is current, else read (and cached).
 */
export function cachedMetrics(path: string, index = 0, wght?: number): FontMetrics {
  let file: string | null = null;
  try {
    const st = statSync(path);
    const key = createHash('sha1').update(`metrics-v2|${path}|${index}|${wght ?? ''}|${st.size}|${st.mtimeMs}`).digest('hex').slice(0, 16);
    file = join(cacheDir(), `font-${key}.json`);
    const m = JSON.parse(readFileSync(file, 'utf8')) as FontMetrics;
    if (m.v === 1 && m.codes.length === m.advances.length) return m;
  } catch { /* no cache yet */ }
  const m = fontMetrics(load(path), index, wght);
  if (file) writeCache(file, m);
  return m;
}

// ---- Font index: family names → files --------------------------------------------------------

/** One face of a font file in the index. */
export interface IndexedFace extends FaceInfo { path: string }

interface IndexFile {
  v: 2;
  roots: string[];
  /** Every folder scanned, with its modification time: a change in any of them rebuilds the index. */
  dirs: Record<string, number>;
  files: Record<string, { size: number; mtime: number; faces: FaceInfo[] }>;
}

const FONT_FILE = /\.(ttf|otf|ttc|otc)$/i;
const MAX_DEPTH = 5;

/** Reads a file by ranges, so a 50 MB collection costs only its name tables. */
function fileReader(fd: number): ReadAt {
  return (offset, length) => {
    const buf = Buffer.alloc(length);
    const n = readSync(fd, buf, 0, length, offset);
    return new Uint8Array(buf.buffer, buf.byteOffset, n);
  };
}

function facesOf(path: string): FaceInfo[] {
  let fd: number | undefined;
  try {
    fd = openSync(path, 'r');
    return readFaces(fileReader(fd));
  } catch {
    return []; // not a font we can read (damaged, a bitmap-only font, no permission)
  } finally {
    if (fd !== undefined) closeSync(fd);
  }
}

const mtime = (p: string) => { try { return statSync(p).mtimeMs; } catch { return -1; } };

function scan(roots: string[], previous: IndexFile | null): IndexFile {
  const out: IndexFile = { v: 2, roots, dirs: {}, files: {} };
  const walk = (dir: string, depth: number) => {
    let entries;
    try { entries = readdirSync(dir, { withFileTypes: true }); } catch { out.dirs[dir] = -1; return; }
    out.dirs[dir] = mtime(dir);
    for (const e of entries) {
      const p = join(dir, e.name);
      if (e.isDirectory()) { if (depth < MAX_DEPTH) walk(p, depth + 1); continue; }
      if (!FONT_FILE.test(e.name) || out.files[p]) continue;
      let st;
      try { st = statSync(p); } catch { continue; }
      const old = previous?.files[p];
      out.files[p] = old && old.size === st.size && old.mtime === st.mtimeMs ? old : { size: st.size, mtime: st.mtimeMs, faces: facesOf(p) };
    }
  };
  for (const r of roots) walk(r, 0);
  return out;
}

let memoIndex: { key: string; idx: IndexFile; faces: IndexedFace[] } | null = null;

/**
 * Every face of every font file in `roots` (default: fontDirs()), from the cached index while no font
 * folder has changed (checked on every call: a long-running add-in server sees fonts installed since).
 */
export function fontIndex(roots = fontDirs()): IndexedFace[] {
  const key = `${cacheDir()}\n${roots.join('\n')}`;
  const fresh = (i: IndexFile | null) => i?.v === 2 && i.roots.join('\n') === roots.join('\n') && Object.entries(i.dirs).every(([d, t]) => mtime(d) === t);
  if (memoIndex?.key === key && fresh(memoIndex.idx)) return memoIndex.faces;
  const file = join(cacheDir(), `font-index-${createHash('sha1').update(roots.join('\n')).digest('hex').slice(0, 12)}.json`);
  let idx: IndexFile | null = memoIndex?.key === key ? memoIndex.idx : null;
  if (!idx) try { idx = JSON.parse(readFileSync(file, 'utf8')) as IndexFile; } catch { /* first run */ }
  if (!fresh(idx)) {
    idx = scan(roots, idx?.v === 2 ? idx : null);
    writeCache(file, idx);
  }
  // Root order is preference order (system before Office's copies, open-charts' own last).
  const faces = Object.entries(idx!.files).flatMap(([path, f]) => f.faces.map((face) => ({ ...face, path })));
  memoIndex = { key, idx: idx!, faces };
  return faces;
}

const norm = (s: string) => s.trim().toLowerCase().replace(/\s+/g, ' ');

/** A face of an installed font file; `wght` is set for a variable font's instance. */
export interface FontFace { path: string; index: number; wght?: number; family: string }

/**
 * The installed face of `family` closest to the style asked for, or null when the family is not
 * installed (or, for bold / italic, has no such face and no weight axis to make one).
 * Names are matched as PowerPoint lists them (name ID 1 in any language), then by typographic family,
 * then by full name.
 */
export function findFont(family: string, opts: { bold?: boolean; italic?: boolean } = {}, faces = fontIndex()): FontFace | null {
  const key = norm(family);
  const tiers: ((f: IndexedFace) => boolean)[] = [
    (f) => f.families.some((n) => norm(n) === key),
    (f) => f.typoFamilies.some((n) => norm(n) === key),
    (f) => f.fullNames.some((n) => norm(n) === key),
  ];
  let found: IndexedFace[] = [];
  for (const t of tiers) { found = faces.filter(t); if (found.length) break; }
  if (!found.length) return null;
  const bold = !!opts.bold, italic = !!opts.italic, want = bold ? 700 : 400;
  const covers = (f: IndexedFace) => !!f.wght && f.wght.min <= want && want <= f.wght.max;
  const score = (f: IndexedFace) => (f.italic !== italic ? 1000 : 0) + (covers(f) ? 1 : Math.abs(f.weight - want) + (f.bold !== bold ? 50 : 0));
  let best = found[0];
  for (const f of found) if (score(f) < score(best)) best = f;
  if (best.italic !== italic && italic) return null;
  if (covers(best)) {
    const at = best.wght!.def === want ? undefined : want;
    return { path: best.path, index: best.index, family: best.families[0] ?? family, ...(at !== undefined ? { wght: at } : {}) };
  }
  // A bold request answered by a regular-weight face: there is no bold; the caller synthesises it.
  // (A regular request answered by a heavy face is right: a single-weight family such as Arial Black.)
  if (bold && best.weight < 600 && !best.bold) return null;
  return { path: best.path, index: best.index, family: best.families[0] ?? family };
}

function measurerForFace(regular: FontFace, bold: FontFace | null): Measurer {
  const r = cachedMetrics(regular.path, regular.index, regular.wght);
  const b = bold ? cachedMetrics(bold.path, bold.index, bold.wght) : null;
  return measurerFromMetrics({ ...r, family: regular.family }, b, () => ({
    regular: load(regular.path), index: regular.index,
    bold: bold ? load(bold.path) : null, boldIndex: bold?.index,
  }));
}

/** Measurer of an installed family (regular + bold faces), or null when it is not installed. */
export function familyMeasurer(family: string, faces?: IndexedFace[]): Measurer | null {
  const all = faces ?? fontIndex();
  const regular = findFont(family, {}, all);
  if (!regular) return null;
  return measurerForFace(regular, findFont(family, { bold: true }, all));
}

/**
 * True the first time `message` is reported on this computer (kept in the cache directory), so the
 * add-in tells the user about a missing font once, not on every chart. Without a cache: once per process.
 */
const reported = new Set<string>();
export function reportOnce(message: string): boolean {
  if (reported.has(message)) return false;
  reported.add(message);
  const file = join(cacheDir(), 'reported.json');
  let seen: string[] = [];
  try { seen = JSON.parse(readFileSync(file, 'utf8')) as string[]; } catch { /* first time */ }
  if (Array.isArray(seen) && seen.includes(message)) return false;
  writeCache(file, [...(Array.isArray(seen) ? seen : []), message].slice(-200));
  return true;
}

/** The metrics of an installed family (the Office.js pane's measurer fetches them), or null. */
export function familyMetrics(family: string): { family: string; regular: FontMetrics; bold: FontMetrics | null } | null {
  const all = fontIndex();
  const r = findFont(family, {}, all);
  if (!r) return null;
  const b = findFont(family, { bold: true }, all);
  return { family: r.family, regular: cachedMetrics(r.path, r.index, r.wght), bold: b ? cachedMetrics(b.path, b.index, b.wght) : null };
}

/**
 * The measurer the CLI, editor and add-in lay out with. With no arguments: every font family on this
 * computer (`forFont`), Arial by default; a missing family is measured with a metric-compatible
 * substitute, else a wider stand-in that the scene reports (core/fonts.ts). With a font file (and
 * optionally its bold), that one file measures every font, as before font lookup existed.
 */
export function createMeasurer(fontFile?: string, boldFile?: string): Measurer {
  const regular = fontFile ?? REGULAR.find(existsSync);
  if (!regular || !existsSync(regular)) throw new Error('No Arial-compatible font found. Pass --font <path to .ttf> (Liberation Sans works).');
  const bold = boldFile ? (existsSync(boldFile) ? boldFile : null) : fontFile ? null : BOLD.find(existsSync) ?? null;
  const base = measurerFromMetrics(cachedMetrics(regular), bold ? cachedMetrics(bold) : null, () => ({ regular: load(regular), bold: bold ? load(bold) : null }));
  if (fontFile) return base;
  return fontSet(base, (family) => familyMeasurer(family));
}
