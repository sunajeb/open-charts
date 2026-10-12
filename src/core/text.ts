// Text measurement and line breaking. Pure: fonts arrive as bytes (see src/fonts.ts for the Node loader).
import { parse } from '../../vendor/opentype.module.js';

interface FontTables {
  hhea?: { ascender: number; descender: number; lineGap: number };
  os2?: { usWinAscent: number; usWinDescent: number; sTypoAscender: number; sTypoDescender: number; sTypoLineGap: number; fsSelection: number };
}
interface Font {
  unitsPerEm?: number;
  tables?: FontTables;
  names?: { fontFamily?: { en?: string } };
  getAdvanceWidth(text: string, size: number, options?: { kerning?: boolean }): number;
}

export interface Measurer {
  width(text: string, size: number, bold?: boolean): number;
  lineHeight(size: number): number;
  /** Family name read from the font file, when known. */
  family?: string;
}

/**
 * Single-spaced line height as a multiple of the font size, the way PowerPoint (GDI / DirectWrite)
 * computes it: typo metrics when OS/2 asks for them (USE_TYPO_METRICS), else win ascent + descent
 * plus whatever extra leading hhea.lineGap adds on top. Arial: (1854 + 434 + 67) / 2048 ≈ 1.1499.
 */
export function lineHeightFactor(font: Font): number {
  const em = font.unitsPerEm || 2048;
  const { hhea, os2 } = font.tables ?? {};
  if (os2 && os2.fsSelection & 0x80) {
    return (os2.sTypoAscender - os2.sTypoDescender + os2.sTypoLineGap) / em;
  }
  if (os2 && os2.usWinAscent) {
    const win = os2.usWinAscent + os2.usWinDescent;
    const gap = hhea ? Math.max(0, hhea.lineGap - (win - (hhea.ascender - hhea.descender))) : 0;
    return (win + gap) / em;
  }
  if (hhea) return (hhea.ascender - hhea.descender + hhea.lineGap) / em;
  return 1.2;
}

/** Measurer over parsed font bytes. Widths are measured without kerning, as PowerPoint draws them. */
export function measurerFromFonts(regularBytes: ArrayBuffer, boldBytes?: ArrayBuffer | null): Measurer {
  const regular = parse(regularBytes) as Font;
  const bold = boldBytes ? (parse(boldBytes) as Font) : null;
  const factor = lineHeightFactor(regular);
  return {
    family: regular.names?.fontFamily?.en,
    width(text, size, isBold = false) {
      const font = isBold ? bold ?? regular : regular;
      const scale = isBold && !bold ? 1.06 : 1;
      return Math.max(...text.split('\n').map((l) => font.getAdvanceWidth(l, size, { kerning: false }) * scale));
    },
    lineHeight: (size) => size * factor,
  };
}

// Break opportunities (a small UAX #14 subset): after spaces (consumed), after hyphens, slashes and
// dashes (kept on the line), and between two CJK ideographs / kana.
const CJK = /[぀-ヿ㐀-䶿一-鿿豈-﫿가-힯]/;
const CLOSING = /[)\]}»”’,.;:!?、。，．）」』】]/;

/** Split text into unbreakable pieces; each piece records whether a space preceded it. */
export function breakPieces(text: string): { s: string; space: boolean }[] {
  const out: { s: string; space: boolean }[] = [];
  let cur = '';
  let space = false;
  const flush = () => { if (cur) out.push({ s: cur, space }); cur = ''; space = false; };
  const chars = [...text];
  for (let i = 0; i < chars.length; i++) {
    const ch = chars[i];
    if (ch === ' ' || ch === '\t' || ch === '\n') { flush(); space = out.length > 0; continue; }
    const next = chars[i + 1];
    cur += ch;
    const breakAfter = (/[-/–—]/.test(ch) && cur.length > 1 && next !== undefined && !CLOSING.test(next) && next !== ' ')
      || (CJK.test(ch) && next !== undefined && CJK.test(next) && !CLOSING.test(next));
    if (breakAfter) flush();
  }
  flush();
  return out;
}

/**
 * Greedy line breaking into at most maxLines lines; the last line keeps any overflow.
 * Use `fits` to check whether the result fits: a single piece wider than the slot never does.
 */
export function wrap(text: string, maxWidth: number, size: number, m: Measurer, maxLines = 3, bold = false): string[] {
  const pieces = breakPieces(text);
  const out: string[] = [];
  let line = '';
  for (const p of pieces) {
    const next = line ? `${line}${p.space ? ' ' : ''}${p.s}` : p.s;
    if (line && m.width(next, size, bold) > maxWidth && out.length < maxLines - 1) {
      out.push(line);
      line = p.s;
    } else {
      line = next;
    }
  }
  if (line) out.push(line);
  return out.length ? out : [''];
}

/** True when every wrapped line fits the width. */
export const fits = (lines: string[], maxWidth: number, size: number, m: Measurer, bold = false) =>
  lines.every((l) => m.width(l, size, bold) <= maxWidth + 0.01);
