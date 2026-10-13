// The engine as the Office.js add-in runs it in the browser: lay a chart out for a frame on a slide,
// produce the writer's one-slide deck for Path A, and render previews for the datasheet editor.
// Pure (no DOM, no Office): bundled by scripts/build-officejs.ts and unit-tested in Node.
import { defaultStyle, layoutChart, themeOf, type ChartSpec, type Frame, type Measurer, type Scene, type Style, type ThemeColors } from '../core/index.ts';
import { buildDeck, chartContent, CHART_FRAME, type DeckSpec } from '../deck.ts';
import { writePptx, type SceneContent } from '../render/pptx.ts';
import { shapeBounds } from '../addin/format.ts';

export { CHART_FRAME };

/** A chart as stored in OC_SPEC: the ChartSpec with the style it was drawn with. */
export type StoredSpec = ChartSpec & { style?: Partial<Style> };

export function splitSpec(spec: StoredSpec): { chart: ChartSpec; style?: Partial<Style> } {
  const { style, ...chart } = spec;
  return { chart, style: style && Object.keys(style).length ? style : undefined };
}

/**
 * Lays the chart out at `frame`. `slideTheme` (the slide's own theme colours, when the host can read
 * them) wins over the spec's `style.theme`, like the .ppam add-in's `theme=`: theme slots then
 * resolve to the colours PowerPoint shows, so label contrast is decided on them. It is not stored.
 */
export function layoutFor(spec: StoredSpec, frame: Frame, m: Measurer, slideTheme?: Partial<ThemeColors>): Scene {
  const { chart, style } = splitSpec(spec);
  const theme = slideTheme ? { ...style?.theme, ...slideTheme } : style?.theme;
  return layoutChart(chart, frame, m, { ...defaultStyle, ...style, ...(theme ? { theme } : {}) });
}

export interface BuiltChart { scene: Scene; content: SceneContent; font: string }

/** Layout plus what the writer and the host need: group name, alt text, tag contract v1. */
export function buildChart(spec: StoredSpec, frame: Frame, m: Measurer, slideTheme?: Partial<ThemeColors>): BuiltChart {
  const { chart, style } = splitSpec(spec);
  const scene = layoutFor(spec, frame, m, slideTheme);
  const font = style?.font ?? defaultStyle.font;
  return { scene, content: chartContent(chart, scene, font, style), font };
}

/** Path A: the chart on a slide of its own, as the CLI writes it (theme slots stay slots). */
export function chartDeck(built: BuiltChart, spec: StoredSpec, slideTheme?: Partial<ThemeColors>): Uint8Array {
  const { style } = splitSpec(spec);
  return writePptx([{ scenes: [built.content] }], built.font, themeOf({ theme: { ...style?.theme, ...slideTheme } }));
}

/** Previews for the editor: the deck as `/api/render` answers it, but laid out in the browser. */
export function renderDeck(deck: DeckSpec, m: Measurer): { slides: { svg: string; issues: string[] }[] } {
  const { slides } = buildDeck(deck, m);
  return { slides: slides.map((s) => ({ svg: s.svg, issues: s.issues })) };
}

/** Bounds of a chart group drawn by any host: the frame (hosts add an invisible frame shape) plus every shape. */
export function sceneBounds(scene: Scene): Frame {
  const f = scene.frame;
  let x0 = f.x, y0 = f.y, x1 = f.x + f.w, y1 = f.y + f.h;
  for (const s of scene.shapes) {
    const b = shapeBounds(s);
    x0 = Math.min(x0, b.x0); y0 = Math.min(y0, b.y0); x1 = Math.max(x1, b.x1); y1 = Math.max(y1, b.y1);
  }
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
}

const close = (a: number, b: number) => Math.abs(a - b) < 0.75;
export const sameBox = (a: Frame, b: Frame) => close(a.x, b.x) && close(a.y, b.y) && close(a.w, b.w) && close(a.h, b.h);

/**
 * The frame to re-lay a chart out in, from its OC_FRAME and the group's current bounds on the slide
 * (same rule as the .ppam add-in, src/addin/engine.ts resolveFrame): bounds hug frame + shapes, so
 * when they still match what the stored frame lays out to, nothing moved; otherwise the frame is
 * mapped by the move/scale that takes the laid-out bounds to the current ones. `resized` drives the
 * pane's banner.
 */
export function resolveFrame(spec: StoredSpec, stored: Frame | undefined, current: Frame | undefined, m: Measurer, slideTheme?: Partial<ThemeColors>): { frame: Frame; resized: boolean } {
  const f0 = stored ?? current ?? CHART_FRAME;
  if (!current || !stored) return { frame: f0, resized: false };
  const b0 = sceneBounds(layoutFor(spec, f0, m, slideTheme));
  if (b0.w < 1 || b0.h < 1 || sameBox(b0, current)) return { frame: f0, resized: false };
  const sx = current.w / b0.w, sy = current.h / b0.h;
  const frame = { x: current.x + (f0.x - b0.x) * sx, y: current.y + (f0.y - b0.y) * sy, w: f0.w * sx, h: f0.h * sy };
  // A move keeps the size: only a change of size needs a new layout (a move is just re-placed).
  return { frame, resized: !close(frame.w, f0.w) || !close(frame.h, f0.h) };
}

/** Base64 of bytes, in chunks (String.fromCharCode on a whole deck overflows the stack). */
export function toBase64(bytes: Uint8Array): string {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

export function fromBase64(b64: string): Uint8Array {
  const s = atob(b64);
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
  return out;
}
