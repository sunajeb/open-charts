// Engine side of the PowerPoint add-in. VBA sends a small text request (through AppleScriptTask on
// the Mac), this lays the chart out with the same engine as the CLI and answers in the line format
// (format.ts) that VBA turns into native shapes.
//
// Request, one field per line:
//   1  op       insert | update | edit | types
//   2  arg      chart type (insert)
//   3  frame    "x,y,w,h" pt: target frame (insert) or the chart's OC_FRAME tag (update/edit)
//   4  bounds   "x,y,w,h" pt: the group's current bounds on the slide (update)
//   5… spec     the chart's OC_SPEC tag, reassembled (update/edit)
import { createHash } from 'node:crypto';
import { defaultStyle, layoutChart, type ChartSpec, type Frame, type Measurer, type Scene, type Style } from '../core/index.ts';
import { chartContent } from '../deck.ts';
import { shapeName } from '../render/pptx.ts';
import { parseFrame } from '../render/tags.ts';
import { FORMAT_VERSION, record, shapeBounds, shapeRecord } from './format.ts';
import { galleryItems, sampleFor } from './registry.ts';

export interface AddinRequest { op: string; arg: string; frame?: Frame; bounds?: Frame; spec: string }

export interface PendingEdits {
  /** The edited spec for a chart whose OC_SPEC hashes to `key`, if the user edited it. */
  take(key: string): string | undefined;
  /** Remember that the chart with this OC_SPEC is being edited; returns the edit URL or message. */
  start(key: string, spec: string): string;
}

export function parseRequest(text: string): AddinRequest {
  const lines = text.replace(/\r\n?/g, '\n').split('\n');
  return {
    op: (lines[0] ?? '').trim(),
    arg: (lines[1] ?? '').trim(),
    frame: parseFrame(lines[2]?.trim()),
    bounds: parseFrame(lines[3]?.trim()),
    spec: lines.slice(4).join('\n').trim(),
  };
}

/** Same frame as a chart on a generated slide: the area below the title. */
export const DEFAULT_FRAME: Frame = { x: 48, y: 96, w: 864, h: 408 };

export const specKey = (spec: string) => createHash('sha1').update(spec).digest('hex').slice(0, 16);

export function handleRequest(req: AddinRequest, m: Measurer, edits?: PendingEdits): string {
  try {
    switch (req.op) {
      case 'insert': {
        const sample = sampleFor(req.arg);
        if ('error' in sample) return error(sample.error);
        return respond(sample.spec, undefined, req.frame ?? DEFAULT_FRAME, m);
      }
      case 'update': {
        if (!req.spec) return error('Select an open-charts chart first.');
        const edited = edits?.take(specKey(req.spec));
        const { chart, style } = readSpec(edited ?? req.spec);
        const frame = resolveFrame(chart, style, req.frame, req.bounds, m);
        return respond(chart, style, frame, m);
      }
      case 'edit': {
        if (!req.spec) return error('Select an open-charts chart first.');
        readSpec(req.spec); // fail early on a damaged tag
        if (!edits) return error('Editing is not available here.');
        return [header(), record('M', edits.start(specKey(req.spec), req.spec))].join('\n') + '\n';
      }
      case 'types':
        return [header(), ...galleryItems().map((g) => record('M', g.id, 'spec' in sampleFor(g.id) ? 'yes' : 'no'))].join('\n') + '\n';
      default:
        return error(`Unknown request "${req.op}". Reinstall the add-in: npx open-charts addin install`);
    }
  } catch (err) {
    return error((err as Error).message);
  }
}

const header = () => record('OC', FORMAT_VERSION);
const error = (msg: string) => `${header()}\n${record('E', msg)}\n`;

function readSpec(json: string): { chart: ChartSpec; style?: Partial<Style> } {
  let parsed: ChartSpec & { style?: Partial<Style> };
  try {
    parsed = JSON.parse(json);
  } catch {
    throw new Error('This chart\'s data (OC_SPEC tag) is damaged and cannot be read.');
  }
  const { style, ...chart } = parsed;
  return { chart, style: style && Object.keys(style).length ? style : undefined };
}

function layout(chart: ChartSpec, style: Partial<Style> | undefined, frame: Frame, m: Measurer): Scene {
  return layoutChart(chart, frame, m, { ...defaultStyle, ...style });
}

export function sceneBounds(scene: Scene): Frame | undefined {
  if (!scene.shapes.length) return undefined;
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
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
export function resolveFrame(chart: ChartSpec, style: Partial<Style> | undefined, stored: Frame | undefined, current: Frame | undefined, m: Measurer): Frame {
  const f0 = stored ?? current ?? DEFAULT_FRAME;
  if (!current) return f0;
  const b0 = sceneBounds(layout(chart, style, f0, m));
  if (!b0 || b0.w < 1 || b0.h < 1) return f0;
  const close = (a: number, b: number) => Math.abs(a - b) < 0.75;
  if (close(b0.x, current.x) && close(b0.y, current.y) && close(b0.w, current.w) && close(b0.h, current.h)) return f0;
  const sx = current.w / b0.w, sy = current.h / b0.h;
  return { x: current.x + (f0.x - b0.x) * sx, y: current.y + (f0.y - b0.y) * sy, w: f0.w * sx, h: f0.h * sy };
}

function respond(chart: ChartSpec, style: Partial<Style> | undefined, frame: Frame, m: Measurer): string {
  const scene = layout(chart, style, frame, m);
  const font = style?.font ?? defaultStyle.font;
  const content = chartContent(chart, scene, font, style);
  const lines = [header(), record('G', content.name, content.altText)];
  for (const [k, v] of content.tags ?? []) lines.push(record('T', k, v));
  scene.shapes.forEach((s, i) => {
    const role = scene.roles?.[i] ?? '';
    lines.push(shapeRecord(s, font, role, shapeName(s, i + 1, role || undefined)));
  });
  return lines.join('\n') + '\n';
}
