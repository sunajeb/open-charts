// Engine side of the PowerPoint add-in. VBA sends a small text request (through AppleScriptTask on
// the Mac), this lays the chart out with the same engine as the CLI and answers in the line format
// (format.ts) that VBA turns into native shapes.
//
// Request, one field per line:
//   1  op       insert | update | relayout | edit | types
//   2  arg      gallery item (insert) or edit-session id (update, optional)
//   3  frame    "x,y,w,h" pt: target frame (insert) or the chart's OC_FRAME tag
//   4  bounds   "x,y,w,h" pt: the group's current bounds on the slide (update, relayout)
//   then, optional, one "key=value" per line:
//      theme=RRGGBB,… the slide's theme colours in ThemeColorScheme order (dk1 lt1 dk2 lt2
//                     accent1…accent6 hlink folHlink): theme slots resolve to these, so label
//                     contrast is decided on the colours PowerPoint will show
//      scale=N        same-scale group: the smallest own scale (OC_SCALE, pt per unit) of the
//                     other charts in this chart's group on the slide
//   5… spec     the chart's OC_SPEC tag, reassembled (update, relayout, edit); always starts with "{"
//
// update   re-lays the chart out where it now is, with the applied edit session if there is one
// relayout the same, but answers "N" (nothing to do) when neither the frame nor the data changed
// edit     opens the chart in the editor (an edit session, docs/ux-editor.md)
import { createHash } from 'node:crypto';
import { defaultStyle, layoutChart, THEME_SLOTS, type ChartSpec, type Frame, type LayoutHints, type Measurer, type Scene, type Style, type ThemeColors } from '../core/index.ts';
import { chartContent } from '../deck.ts';
import { shapeName } from '../render/pptx.ts';
import { parseFrame } from '../render/tags.ts';
import type { Session } from '../server.ts';
import { FORMAT_VERSION, num, record, shapeBounds, shapeRecord } from './format.ts';
import { galleryItems, sampleFor, typeLabel } from './registry.ts';
import { reinstallCommand } from './launcher.ts';

export interface AddinRequest {
  op: string; arg: string; frame?: Frame; bounds?: Frame; spec: string;
  /** The slide's theme colours (from `theme=`); they win over the spec's `style.theme`. */
  theme?: Partial<ThemeColors>;
  /** Same-scale group: the smallest own scale of the other charts in the group (from `scale=`). */
  ptPerUnit?: number;
}

/** Context of one layout in PowerPoint: the slide's theme and the chart's same-scale group. */
interface Env { theme?: Partial<ThemeColors>; hints: LayoutHints }

/** ThemeColorScheme order (msoThemeColorSchemeIndex 1–12); the last two are not slots. */
const SCHEME_ORDER = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'] as const;

function parseTheme(v: string): Partial<ThemeColors> | undefined {
  const hex = v.split(',').map((x) => x.trim());
  const out: Partial<ThemeColors> = {};
  SCHEME_ORDER.forEach((k, i) => { if (/^[0-9A-Fa-f]{6}$/.test(hex[i] ?? '') && (THEME_SLOTS as readonly string[]).includes(k)) out[k] = hex[i].toUpperCase(); });
  return Object.keys(out).length ? out : undefined;
}

/** Edit sessions on the add-in's local server (src/addin/server.ts). */
export interface Sessions {
  get(id: string): Promise<Session | undefined>;
  /** Stores the chart as an open session and shows it in the editor. */
  open(id: string, chart: object, title: string): Promise<void>;
  /**
   * Marks an applied session as laid out (back to "open", chart kept; the add-in server sets it to
   * "drawn" once PowerPoint has finished drawing). Sessions are keyed by the chart's OC_SPEC, so
   * without this an edit would later be applied to any other chart that still has the same data (a
   * fresh sample of the same type, a copy) once the editor was closed.
   */
  settle?(id: string): Promise<void>;
}

export function parseRequest(text: string): AddinRequest {
  const lines = text.replace(/\r\n?/g, '\n').split('\n');
  let i = 4;
  const params: Record<string, string> = {};
  for (; i < lines.length && /^[a-z]+=/.test(lines[i]); i++) {
    const at = lines[i].indexOf('=');
    params[lines[i].slice(0, at)] = lines[i].slice(at + 1).trim();
  }
  const scale = Number(params.scale);
  return {
    op: (lines[0] ?? '').trim(),
    arg: (lines[1] ?? '').trim(),
    frame: parseFrame(lines[2]?.trim()),
    bounds: parseFrame(lines[3]?.trim()),
    spec: lines.slice(i).join('\n').trim(),
    ...(params.theme ? { theme: parseTheme(params.theme) } : {}),
    ...(params.scale && scale > 0 && Number.isFinite(scale) ? { ptPerUnit: scale } : {}),
  };
}

/** Same frame as a chart on a generated slide: the area below the title. */
export const DEFAULT_FRAME: Frame = { x: 48, y: 96, w: 864, h: 408 };

/** Edit-session id for a chart: derived from its OC_SPEC, so Update chart finds an applied edit. */
export const specKey = (spec: string) => createHash('sha1').update(spec).digest('hex').slice(0, 16);

const SELECT_FIRST = 'Select an open-charts chart first, then try again.';

export async function handleRequest(req: AddinRequest, m: Measurer, sessions?: Sessions): Promise<string> {
  const env: Env = { theme: req.theme, hints: req.ptPerUnit ? { ptPerUnit: req.ptPerUnit } : {} };
  try {
    switch (req.op) {
      case 'insert': {
        const sample = sampleFor(req.arg);
        if ('error' in sample) return error(sample.error);
        return respond(sample.spec, undefined, req.frame ?? DEFAULT_FRAME, m, env);
      }
      case 'update':
      case 'relayout': {
        if (!req.spec) return error(SELECT_FIRST);
        const id = req.arg || specKey(req.spec);
        const session = await sessions?.get(id);
        // An Apply that did not reach the slide ("error") is still pending for an explicit Update
        // chart on that chart, but never for the automatic re-layout of another chart with the same data.
        const pending = session?.state === 'applied' || (session?.state === 'error' && req.op === 'update');
        const edited = pending && session!.chart && typeof session!.chart === 'object';
        const { chart, style } = edited ? splitSpec(session!.chart as ChartSpec & { style?: Partial<Style> }) : readSpec(req.spec);
        const frame = resolveFrame(chart, style, req.frame, req.bounds, m, env);
        if (req.op === 'relayout' && !edited && req.frame && sameFrame(frame, req.frame)) return `${header()}\n${record('N')}\n`;
        const out = respond(chart, style, frame, m, env);
        if (edited) await sessions?.settle?.(id);
        return out;
      }
      case 'edit': {
        if (!req.spec) return error(SELECT_FIRST);
        const { chart, style } = readSpec(req.spec); // fail early on a damaged tag
        if (!sessions) return error('Editing is not available here.');
        const id = specKey(req.spec);
        await sessions.open(id, style ? { ...chart, style } : chart, `${typeLabel(chart.type)} chart`);
        return `${header()}\n${record('I', id)}\n`;
      }
      case 'types':
        return [header(), ...galleryItems().map((g) => record('M', g.id, 'spec' in sampleFor(g.id) ? 'yes' : 'no'))].join('\n') + '\n';
      default:
        return error(`This add-in and its engine do not match ("${req.op}"). In Terminal, run: ${reinstallCommand()}`);
    }
  } catch (err) {
    return error((err as Error).message);
  }
}

const sameFrame = (a: Frame, b: Frame) =>
  Math.abs(a.x - b.x) < 0.5 && Math.abs(a.y - b.y) < 0.5 && Math.abs(a.w - b.w) < 0.5 && Math.abs(a.h - b.h) < 0.5;

const header = () => record('OC', FORMAT_VERSION);
const error = (msg: string) => `${header()}\n${record('E', msg)}\n`;

function readSpec(json: string): { chart: ChartSpec; style?: Partial<Style> } {
  let parsed: ChartSpec & { style?: Partial<Style> };
  try {
    parsed = JSON.parse(json);
  } catch {
    throw new Error('This chart\'s data (OC_SPEC tag) is damaged and cannot be read.');
  }
  return splitSpec(parsed);
}

function splitSpec(parsed: ChartSpec & { style?: Partial<Style> }): { chart: ChartSpec; style?: Partial<Style> } {
  const { style, ...chart } = parsed;
  return { chart, style: style && Object.keys(style).length ? style : undefined };
}

function layout(chart: ChartSpec, style: Partial<Style> | undefined, frame: Frame, m: Measurer, env: Env = { hints: {} }): Scene {
  // Theme slots resolve to the slide's own theme colours: PowerPoint draws them in those.
  const theme = env.theme ? { ...style?.theme, ...env.theme } : style?.theme;
  return layoutChart(chart, frame, m, { ...defaultStyle, ...style, ...(theme ? { theme } : {}) }, env.hints);
}

/** Bounds of the chart group as drawn by the add-in: the shapes plus its invisible frame rectangle. */
export function sceneBounds(scene: Scene): Frame | undefined {
  const f = scene.frame;
  let x0 = f.x, y0 = f.y, x1 = f.x + f.w, y1 = f.y + f.h;
  for (const s of scene.shapes) {
    const b = shapeBounds(s);
    x0 = Math.min(x0, b.x0); y0 = Math.min(y0, b.y0); x1 = Math.max(x1, b.x1); y1 = Math.max(y1, b.y1);
  }
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
}

/**
 * The frame to re-lay the chart out in. VBA reports the group's bounds, which hug the shapes, not
 * the frame; so lay out in the stored frame, and if the user has since moved or resized the group,
 * map the frame by the same move/scale that takes the laid-out bounds to the current ones.
 */
export function resolveFrame(chart: ChartSpec, style: Partial<Style> | undefined, stored: Frame | undefined, current: Frame | undefined, m: Measurer, env?: Env): Frame {
  const f0 = stored ?? current ?? DEFAULT_FRAME;
  if (!current) return f0;
  const b0 = sceneBounds(layout(chart, style, f0, m, env));
  if (!b0 || b0.w < 1 || b0.h < 1) return f0;
  const close = (a: number, b: number) => Math.abs(a - b) < 0.75;
  if (close(b0.x, current.x) && close(b0.y, current.y) && close(b0.w, current.w) && close(b0.h, current.h)) return f0;
  const sx = current.w / b0.w, sy = current.h / b0.h;
  return { x: current.x + (f0.x - b0.x) * sx, y: current.y + (f0.y - b0.y) * sy, w: f0.w * sx, h: f0.h * sy };
}

function respond(chart: ChartSpec, style: Partial<Style> | undefined, frame: Frame, m: Measurer, env?: Env): string {
  const scene = layout(chart, style, frame, m, env);
  const font = style?.font ?? defaultStyle.font;
  const content = chartContent(chart, scene, font, style);
  const lines = [header(), record('G', content.name, content.altText), record('F', [frame.x, frame.y, frame.w, frame.h].map(num).join(','))];
  for (const [k, v] of content.tags ?? []) lines.push(record('T', k, v));
  scene.shapes.forEach((s, i) => {
    const role = scene.roles?.[i] ?? '';
    lines.push(shapeRecord(s, font, role, shapeName(s, i + 1, role || undefined)));
  });
  return lines.join('\n') + '\n';
}
