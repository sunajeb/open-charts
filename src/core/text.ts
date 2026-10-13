// Text measurement and line breaking. Pure: fonts arrive as bytes (see src/fonts.ts for the Node loader).
// Metrics come from the font's own tables (core/sfnt.ts: TrueType and CFF outlines, collections,
// variable fonts); opentype.js only shapes the scripts that need it (Arabic, Hebrew, ...).
import { parse } from '../../vendor/opentype.module.js';
import { bytesReader, extractFace, faceMetrics } from './sfnt.ts';

interface FontTables {
  hhea?: { ascender: number; descender: number; lineGap: number };
  os2?: { usWinAscent: number; usWinDescent: number; sTypoAscender: number; sTypoDescender: number; sTypoLineGap: number; fsSelection: number };
}
/** The parts of an opentype.js font this module reads. */
interface Font {
  unitsPerEm?: number;
  tables?: FontTables;
  getAdvanceWidth(text: string, size: number, options?: { kerning?: boolean }): number;
}

export interface Measurer {
  width(text: string, size: number, bold?: boolean): number;
  lineHeight(size: number): number;
  /** Family name read from the font file, when known. */
  family?: string;
  /**
   * The measurer for another font family (theme fonts already resolved: core/fonts.ts). Absent:
   * this measurer stands for every font (one font file given by path, older callers).
   */
  forFont?(family: string): Measurer;
  /** Set when the family asked for was not found and this measurer stands in for it. */
  fallback?: FontFallback;
}

/** A font that could not be found, and what measures it instead (core/fonts.ts). */
export interface FontFallback {
  requested: string;
  /** The font measured instead. */
  used: string;
  /** No metric-compatible substitute: a conservative estimate (wider and taller than most fonts). */
  wide: boolean;
}

/**
 * Single-spaced line height of an opentype.js font as a multiple of the font size (the same rule
 * core/sfnt.ts applies to the tables): typo metrics when OS/2 asks for them (USE_TYPO_METRICS), else
 * win ascent + descent plus whatever extra leading hhea.lineGap adds on top. Arial ≈ 1.1499.
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

/**
 * What measuring needs from a font file, small enough to cache as JSON (src/fonts.ts keeps it on disk
 * so the CLI does not parse the font on every start): advance width per code point, in font units.
 */
export interface FontMetrics {
  v: 1;
  family?: string;
  unitsPerEm: number;
  lineHeight: number;
  notdef: number;
  /** Code points and their advance widths, as two parallel arrays. */
  codes: number[];
  advances: number[];
}

/**
 * Metrics of face `index` of a font file (a collection holds several); `wght` picks an instance of
 * a variable font (default: its default instance). The same numbers opentype.js reads (test/fonts).
 */
export function fontMetrics(bytes: ArrayBuffer, index = 0, wght?: number): FontMetrics {
  return faceMetrics(bytesReader(bytes), index, wght);
}

/** opentype.js over face `index` of a font file; null when it cannot read it (e.g. CFF2 outlines). */
function shaperFont(bytes: ArrayBuffer, index = 0): Font | null {
  try { return parse(extractFace(bytes, index)) as Font; } catch { return null; }
}

// Scripts whose glyphs change with context (Hebrew, Arabic, Syriac, ... and their presentation forms):
// measured by the full shaper (opentype.js), parsed on first use.
// East Asian ideographs, kana, Hangul and full-width forms missing from a Latin font: PowerPoint draws
// them in a fallback East Asian font, where they are one em wide (not the .notdef box's width).
const FULL_WIDTH = /[\u1100-\u115F\u2E80-\u9FFF\uA960-\uA97F\uAC00-\uD7AF\uF900-\uFAFF\uFE30-\uFE4F\uFF01-\uFF60\uFFE0-\uFFE6]/u;

const SHAPED = /[֐-ࣿיִ-﷿ﹰ-﻿]/;

/**
 * Measurer over font metrics: the sum of the glyph advances, without kerning or ligatures, as
 * PowerPoint draws text by default (identical to opentype.js for Latin, Greek and Cyrillic — checked
 * over every pair of Latin characters in test/text.test.ts; East Asian characters the font lacks are one em). Text in shaped scripts goes through
 * opentype.js, from `fontBytes` when given. Widths are memoised (bounded), so repeated labels cost nothing.
 */
export function measurerFromMetrics(
  regular: FontMetrics, bold?: FontMetrics | null,
  fontBytes?: () => { regular: ArrayBuffer; bold?: ArrayBuffer | null; index?: number; boldIndex?: number } | null,
): Measurer {
  return build(regular, bold, () => {
    const b = fontBytes?.();
    return b ? [shaperFont(b.regular, b.index), b.bold ? shaperFont(b.bold, b.boldIndex) : null] : [];
  });
}

function build(regular: FontMetrics, bold: FontMetrics | null | undefined, fonts: () => (Font | null)[]): Measurer {
  const tables = [regular, bold ?? regular].map((fm) => ({ fm, adv: new Map(fm.codes.map((c, i) => [c, fm.advances[i]])) }));
  let shaper: (Font | null)[] | null = null;
  const shaped = (line: string, size: number, isBold: boolean) => {
    shaper ??= fonts();
    const font = isBold ? shaper[1] ?? shaper[0] : shaper[0];
    if (!font) return null;
    // Some system fonts (e.g. Windows' Arial) use GSUB lookups opentype.js can't apply; never fail a
    // layout over it — fall back to the plain advance-width table for that line.
    try { return font.getAdvanceWidth(line, size, { kerning: false }); } catch { return null; }
  };
  const lineWidth = (line: string, size: number, isBold: boolean) => {
    if (SHAPED.test(line)) { const w = shaped(line, size, isBold); if (w != null) return w; }
    const { fm, adv } = tables[isBold ? 1 : 0];
    // Same arithmetic, in the same order, as opentype.js's forEachGlyph, so widths match to the bit.
    const scale = 1 / fm.unitsPerEm * size;
    let x = 0;
    for (const ch of line) {
      const a = adv.get(ch.codePointAt(0)!) ?? (FULL_WIDTH.test(ch) ? fm.unitsPerEm : fm.notdef);
      if (a) x += a * scale;
    }
    return x;
  };
  const memo = new Map<string, number>();
  return {
    family: regular.family,
    width(text, size, isBold = false) {
      const key = `${isBold ? 'b' : 'r'}${size}|${text}`;
      const hit = memo.get(key);
      if (hit !== undefined) return hit;
      const scale = isBold && !bold ? 1.06 : 1;
      const w = Math.max(...text.split('\n').map((l) => lineWidth(l, size, isBold) * scale));
      if (memo.size >= 20000) memo.delete(memo.keys().next().value!);
      memo.set(key, w);
      return w;
    },
    lineHeight: (size) => size * regular.lineHeight,
  };
}

/** Measurer over font bytes. Widths are measured without kerning, as PowerPoint draws them. */
export function measurerFromFonts(regularBytes: ArrayBuffer, boldBytes?: ArrayBuffer | null): Measurer {
  return measurerFromMetrics(fontMetrics(regularBytes), boldBytes ? fontMetrics(boldBytes) : null, () => ({ regular: regularBytes, bold: boldBytes }));
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
