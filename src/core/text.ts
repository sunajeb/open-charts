import { existsSync, readFileSync } from 'node:fs';
import { parse } from '../../vendor/opentype.module.js';

interface Font { getAdvanceWidth(text: string, size: number, options?: { kerning?: boolean }): number }

export interface Measurer {
  width(text: string, size: number, bold?: boolean): number;
  lineHeight(size: number): number;
}

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

function load(paths: string[]): Font | null {
  const p = paths.find(existsSync);
  if (!p) return null;
  const buf = readFileSync(p);
  return parse(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
}

export function createMeasurer(fontFile?: string, boldFile?: string): Measurer {
  const regular = load(fontFile ? [fontFile] : REGULAR);
  const bold = load(boldFile ? [boldFile] : BOLD);
  if (!regular) throw new Error('No Arial-compatible font found. Pass --font <path to .ttf> (Liberation Sans works).');
  const lines = (text: string) => text.split('\n');
  return {
    width(text, size, isBold = false) {
      const font = isBold ? bold ?? regular : regular;
      const scale = isBold && !bold ? 1.06 : 1;
      return Math.max(...lines(text).map((l) => font.getAdvanceWidth(l, size, { kerning: false }) * scale));
    },
    // PowerPoint single spacing for Arial ≈ (ascent + descent) / em = 1.149; kerning is off by default.
    lineHeight: (size) => size * 1.149,
  };
}

/** Greedy word wrap into at most maxLines lines; the last line keeps any overflow. */
export function wrap(text: string, maxWidth: number, size: number, m: Measurer, maxLines = 3): string[] {
  const words = text.split(/\s+/).filter(Boolean);
  const out: string[] = [];
  let line = '';
  for (const word of words) {
    const next = line ? `${line} ${word}` : word;
    if (line && m.width(next, size) > maxWidth && out.length < maxLines - 1) {
      out.push(line);
      line = word;
    } else {
      line = next;
    }
  }
  if (line) out.push(line);
  return out.length ? out : [''];
}
