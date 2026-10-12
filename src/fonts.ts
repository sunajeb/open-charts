// Node font loader: finds Arial or a metric-compatible substitute on disk and builds a Measurer.
// Lives outside src/core so the engine stays free of file-system access.
import { existsSync, readFileSync } from 'node:fs';
import { measurerFromFonts, type Measurer } from './core/text.ts';

// Arial or a metric-compatible substitute (Liberation Sans, Arimo) gives widths that match PowerPoint.
const REGULAR = [
  '/System/Library/Fonts/Supplemental/Arial.ttf',
  'C:\\Windows\\Fonts\\arial.ttf',
  '/usr/share/fonts/truetype/msttcorefonts/Arial.ttf',
  '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
  '/usr/share/fonts/liberation-sans/LiberationSans-Regular.ttf',
];
const BOLD = [
  '/System/Library/Fonts/Supplemental/Arial Bold.ttf',
  'C:\\Windows\\Fonts\\arialbd.ttf',
  '/usr/share/fonts/truetype/msttcorefonts/Arial_Bold.ttf',
  '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf',
  '/usr/share/fonts/liberation-sans/LiberationSans-Bold.ttf',
];

function load(paths: string[]): ArrayBuffer | null {
  const p = paths.find(existsSync);
  if (!p) return null;
  const buf = readFileSync(p);
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
}

export function createMeasurer(fontFile?: string, boldFile?: string): Measurer {
  const regular = load(fontFile ? [fontFile] : REGULAR);
  if (!regular) throw new Error('No Arial-compatible font found. Pass --font <path to .ttf> (Liberation Sans works).');
  return measurerFromFonts(regular, load(boldFile ? [boldFile] : BOLD));
}
