// Text measurement in the browser. Arial (and its metric twins) is measured with the bundled
// Liberation Sans (SIL OFL, vendor/fonts), parsed by the same code the CLI uses, so those scenes
// are byte-identical to Node's (test/e2e/officejs.spec.ts). Any other font, in this order:
//   1. exact metrics from the server the pane is loaded from (GET fonts/metrics?family=…, answered by
//      src/officejs/server.ts with the Node font resolver, src/fonts.ts): the same scene as Node;
//   2. the browser's canvas, when the font is installed where the pane runs: advance widths
//      without kerning, exactly as Node reads them; line spacing estimated (browsers do not report
//      a font's line gap) as ascent + descent + 3 % (Arial's gap);
//   3. a metric-compatible substitute, else the conservative wide stand-in (core/fonts.ts), which
//      the scene reports.
// Fonts are loaded before a layout with `load` (the pane knows a chart's fonts from its spec).
import { measurerFromFonts, measurerFromMetrics, type FontMetrics, type Measurer } from '../core/text.ts';
import { fontSet, isThemeFont } from '../core/fonts.ts';

export const FONT_FILES = { regular: 'fonts/LiberationSans-Regular.ttf', bold: 'fonts/LiberationSans-Bold.ttf' };

/** What the server answers for a family it has installed. */
export interface FontMetricsResponse { family: string; regular: FontMetrics; bold: FontMetrics | null }

export interface BrowserMeasurer extends Measurer {
  /** Fetches exact metrics for these families (once each) before a layout; never fails. */
  load(families: (string | undefined)[]): Promise<void>;
}

/** Fonts a chart (with its style) is drawn in: the style's font and any element's own font. */
export function specFonts(spec: { style?: { font?: string }; overrides?: Record<string, { font?: string } | undefined> } | undefined): string[] {
  const out = new Set<string>();
  if (spec?.style?.font) out.add(spec.style.font);
  for (const o of Object.values(spec?.overrides ?? {})) if (o && typeof o.font === 'string') out.add(o.font);
  return [...out];
}

type Ctx2D = { font: string; fontKerning?: string; measureText(t: string): { width: number; fontBoundingBoxAscent?: number; fontBoundingBoxDescent?: number } };

function canvas2d(): Ctx2D | null {
  try {
    const g = globalThis as { OffscreenCanvas?: new (w: number, h: number) => { getContext(k: '2d'): unknown }; document?: { createElement(t: 'canvas'): { getContext(k: '2d'): unknown } } };
    const c = (g.OffscreenCanvas ? new g.OffscreenCanvas(1, 1).getContext('2d') : g.document?.createElement('canvas').getContext('2d')) as Ctx2D | null;
    if (c && 'fontKerning' in c) c.fontKerning = 'none';
    return c ?? null;
  } catch { return null; }
}

const SIZE = 1000; // measure at a large size: canvas widths are exact; ascent / descent are whole pixels
const PROBE = 'mmmmmlli WAV 0123 ÄÖ';
const quote = (f: string) => `"${f.replace(/["\\]/g, '')}"`;

/** A measurer over the browser's own copy of `family`, or null when the browser does not have it. */
export function canvasMeasurer(family: string, ctx = canvas2d()): Measurer | null {
  if (!ctx) return null;
  const f = quote(family);
  // Installed = text in "family, generic" differs from the generic font alone (for two generics).
  const installed = ['monospace', 'serif'].some((g) => {
    ctx.font = `100px ${f}, ${g}`;
    const a = ctx.measureText(PROBE).width;
    ctx.font = `100px ${g}`;
    return a !== ctx.measureText(PROBE).width;
  });
  if (!installed) return null;
  ctx.font = `${SIZE}px ${f}`;
  const tm = ctx.measureText('Hg');
  const box = (tm.fontBoundingBoxAscent ?? 0) + (tm.fontBoundingBoxDescent ?? 0);
  const factor = box > 0 ? box / SIZE + 0.03 : 1.2;
  const memo = new Map<string, number>();
  return {
    family,
    width(text, size, bold = false) {
      const key = `${bold ? 'b' : 'r'}|${text}`;
      let w = memo.get(key);
      if (w === undefined) {
        ctx.font = `${bold ? 'bold ' : ''}${SIZE}px ${f}`;
        w = Math.max(0, ...text.split('\n').map((l) => ctx.measureText(l).width)) / SIZE;
        if (memo.size >= 20000) memo.delete(memo.keys().next().value!);
        memo.set(key, w);
      }
      return w * size;
    },
    lineHeight: (size) => size * factor,
  };
}

/**
 * Fetches both bundled faces relative to `base` (the task pane's URL) and builds the measurer. `opts.canvas`
 * false: no canvas measuring (tests of the fallback).
 */
export async function loadMeasurer(base: string | URL = (globalThis as { location?: Location }).location?.href ?? '', opts: { canvas?: boolean } = {}): Promise<BrowserMeasurer> {
  const get = async (p: string) => {
    const r = await fetch(new URL(p, base));
    if (!r.ok) throw new Error(`Could not load ${p} (HTTP ${r.status})`);
    return r.arrayBuffer();
  };
  const [regular, bold] = await Promise.all([get(FONT_FILES.regular), get(FONT_FILES.bold)]);
  const bundled = measurerFromFonts(regular, bold);
  const exact = new Map<string, Measurer | null>();
  const ctx = opts.canvas === false ? null : canvas2d();
  // Arial's exact twins (Helvetica shares the widths but not the line spacing: measured on its own).
  const twin = new Set(['liberation sans', 'arial', 'arimo']);
  const find = (family: string): Measurer | null => {
    const key = family.toLowerCase();
    if (twin.has(key)) return null; // the bundled font measures it (fontSet's substitute)
    return exact.get(key) ?? canvasMeasurer(family, ctx);
  };
  let set = fontSet(bundled, find);
  // The service answers 200 with metrics, or with null for a font it does not have; anything else
  // (a static host's 404) means there is no service: it is not asked again.
  let service = true;
  return {
    family: bundled.family,
    width: (t, s, b) => set.width(t, s, b),
    lineHeight: (s) => set.lineHeight(s),
    forFont: (f) => set.forFont!(f),
    async load(families) {
      const todo = [...new Set(families.filter((f): f is string => !!f && !isThemeFont(f)).map((f) => f.trim()))]
        .filter((f) => !twin.has(f.toLowerCase()) && !exact.has(f.toLowerCase()));
      if (!todo.length || !service) return;
      await Promise.all(todo.map(async (family) => {
        let m: Measurer | null = null;
        try {
          const r = await fetch(new URL(`fonts/metrics?family=${encodeURIComponent(family)}`, base));
          const j = r.ok ? await r.json() as FontMetricsResponse | null : undefined;
          if (j === undefined) service = false;
          else if (j?.regular?.v === 1) m = { ...measurerFromMetrics(j.regular, j.bold), family };
        } catch { service = false; }
        exact.set(family.toLowerCase(), m);
      }));
      set = fontSet(bundled, find); // forget what was measured before the exact metrics arrived
    },
  };
}
