// Node font loader: finds Arial or a metric-compatible substitute on disk and builds a Measurer.
// Lives outside src/core so the engine stays free of file-system access.
//
// Parsing a font takes tens of milliseconds, most of the CLI's start-up after Node itself, so the
// metrics measuring needs (advance widths per code point, line height) are cached as JSON, keyed by
// the font file's path, size and modification time: in $OPEN_CHARTS_CACHE, else $XDG_CACHE_HOME/open-charts
// (~/.cache/open-charts).
// The cache is an optimisation only: unreadable or unwritable, the font is parsed as before.
import { existsSync, mkdirSync, readFileSync, renameSync, statSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { fontMetrics, measurerFromMetrics, type FontMetrics, type Measurer } from './core/text.ts';

// Windows keeps its fonts in %WINDIR%\Fonts (C:\Windows\Fonts on almost every PC).
const WINDOWS_FONTS = `${process.env.WINDIR ?? process.env.SystemRoot ?? 'C:\\Windows'}\\Fonts`;
// Arial or a metric-compatible substitute (Liberation Sans, Arimo) gives widths that match PowerPoint.
const REGULAR = [
  '/System/Library/Fonts/Supplemental/Arial.ttf',
  `${WINDOWS_FONTS}\\arial.ttf`,
  '/usr/share/fonts/truetype/msttcorefonts/Arial.ttf',
  '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
  '/usr/share/fonts/liberation-sans/LiberationSans-Regular.ttf',
];
const BOLD = [
  '/System/Library/Fonts/Supplemental/Arial Bold.ttf',
  `${WINDOWS_FONTS}\\arialbd.ttf`,
  '/usr/share/fonts/truetype/msttcorefonts/Arial_Bold.ttf',
  '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf',
  '/usr/share/fonts/liberation-sans/LiberationSans-Bold.ttf',
];

function load(p: string): ArrayBuffer {
  const buf = readFileSync(p);
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
}

const cacheDir = () => process.env.OPEN_CHARTS_CACHE || join(process.env.XDG_CACHE_HOME || join(homedir(), '.cache'), 'open-charts');

/** Metrics of one font file, from the cache when it is current, else parsed (and cached). */
export function cachedMetrics(path: string): FontMetrics {
  let file: string | null = null;
  try {
    const st = statSync(path);
    const key = createHash('sha1').update(`metrics-v1|${path}|${st.size}|${st.mtimeMs}`).digest('hex').slice(0, 16);
    file = join(cacheDir(), `font-${key}.json`);
    const m = JSON.parse(readFileSync(file, 'utf8')) as FontMetrics;
    if (m.v === 1 && m.codes.length === m.advances.length) return m;
  } catch { /* no cache yet */ }
  const m = fontMetrics(load(path));
  if (file) {
    try {
      mkdirSync(cacheDir(), { recursive: true });
      const tmp = `${file}.${process.pid}.tmp`;
      writeFileSync(tmp, JSON.stringify(m));
      renameSync(tmp, file); // atomic: concurrent CLIs never read half a file
    } catch { /* read-only home: fine, just slower next time */ }
  }
  return m;
}

export function createMeasurer(fontFile?: string, boldFile?: string): Measurer {
  const regular = fontFile ?? REGULAR.find(existsSync);
  if (!regular || !existsSync(regular)) throw new Error('No Arial-compatible font found. Pass --font <path to .ttf> (Liberation Sans works).');
  const bold = boldFile ? (existsSync(boldFile) ? boldFile : null) : BOLD.find(existsSync) ?? null;
  return measurerFromMetrics(cachedMetrics(regular), bold ? cachedMetrics(bold) : null, () => ({ regular: load(regular), bold: bold ? load(bold) : null }));
}
