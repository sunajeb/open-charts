// Drag handles: the chart's directly manipulable properties, found in a laid-out scene, and what a
// drag does to the spec. Host-independent (no screen, no PowerPoint): the add-in maps the handles to
// the screen (src/addin/handles.ts) and the Mac helper's overlay catches the mouse.
//
//   gap      the right edge of the first column (bottom edge of the first bar when horizontal): its
//            width, kept centred in its slot -> options.gap
//   axisMax  the top of the value axis (its right end when horizontal) -> options.axisMax (headroom
//            only: dragged below the top of the data, the option is removed)
//   level    the bubble of level arrow k -> options.levelArrows[k].position, whichever of "beside"
//            and "between" puts the bubble nearest the pointer
//   label    a selected label: anywhere -> overrides[role].pin (centre as fractions of the frame);
//            the label is then placed by hand until reset (unpinElement)
//
// Every handle resets with `resetSpec` (the add-in: double-click it). All coordinates are slide points.
import type { ChartSpec, Frame, Scene, TextShape } from './types.ts';
import { appliesTo, LIMITS } from './spec.ts';
import { kindOf, unpinElement } from './overrides.ts';

export type HandleKind = 'gap' | 'axisMax' | 'level' | 'label';

export interface DragHandle {
  /** Stable within one chart: `gap`, `axisMax`, `level:<k>`, `label:<role>`. */
  id: string;
  kind: HandleKind;
  /** The marker's centre on the slide (an area handle: the element's centre). */
  at: { x: number; y: number };
  /** Area handle (a selected label): the whole box catches the mouse, no marker is drawn. */
  box?: Frame;
  /** Mouse cursor over the handle. */
  cursor: 'ew' | 'ns' | 'move';
  role?: string;
  index?: number;
}

/** Lays a spec out at the chart's frame and context (level arrows try their positions). */
export type Relayout = (spec: ChartSpec) => Scene;

const horizontal = (spec: ChartSpec) => spec.options?.orientation === 'horizontal' && appliesTo('orientation', spec.type);
const centre = (b: { x: number; y: number; w: number; h: number }) => ({ x: b.x + b.w / 2, y: b.y + b.h / 2 });

/** The bars of the first category that has any: their union on the slide. */
function firstCluster(scene: Scene): Frame | undefined {
  let best = Infinity;
  let box: { x0: number; y0: number; x1: number; y1: number } | undefined;
  scene.shapes.forEach((s, i) => {
    const m = /^bar:s\d+:c(\d+)$/.exec(scene.roles?.[i] ?? '');
    if (!m || s.kind !== 'rect' || !(s.w > 0) || !(s.h > 0)) return;
    const c = Number(m[1]);
    if (c > best) return;
    if (c < best) { best = c; box = undefined; }
    box = box
      ? { x0: Math.min(box.x0, s.x), y0: Math.min(box.y0, s.y), x1: Math.max(box.x1, s.x + s.w), y1: Math.max(box.y1, s.y + s.h) }
      : { x0: s.x, y0: s.y, x1: s.x + s.w, y1: s.y + s.h };
  });
  return box && { x: box.x0, y: box.y0, w: box.x1 - box.x0, h: box.y1 - box.y0 };
}

/** The text shape of a role (a label, a level arrow's bubble). */
function textOf(scene: Scene, role: string): TextShape | undefined {
  const i = scene.shapes.findIndex((s, k) => s.kind === 'text' && scene.roles?.[k] === role);
  return i >= 0 ? (scene.shapes[i] as TextShape) : undefined;
}

/** The handles a laid-out chart offers; `selected` (OC_ROLE values) adds the selected labels. */
export function dragHandles(spec: ChartSpec, scene: Scene, selected: readonly string[] = []): DragHandle[] {
  const out: DragHandle[] = [];
  if (!scene.roles || scene.roles.includes('placeholder')) return out;
  const hz = horizontal(spec);
  if (appliesTo('gap', spec.type)) {
    const c = firstCluster(scene);
    if (c) out.push({ id: 'gap', kind: 'gap', at: hz ? { x: c.x + c.w / 2, y: c.y + c.h } : { x: c.x + c.w, y: c.y + c.h / 2 }, cursor: hz ? 'ns' : 'ew' });
  }
  const at = scene.valueScale?.at;
  // Only where a value axis is drawn: the handle sits on it.
  if (at && appliesTo('axisMax', spec.type) && spec.options?.valueAxis !== false && scene.roles.includes('axis')
      && scene.roles.some((r) => r.startsWith('axis:tick:'))) {
    const along = Math.abs(at[2] - at[0]) > Math.abs(at[3] - at[1]);
    out.push({ id: 'axisMax', kind: 'axisMax', at: { x: at[2], y: at[3] }, cursor: along ? 'ew' : 'ns' });
  }
  (spec.options?.levelArrows ?? []).forEach((_, k) => {
    const t = textOf(scene, `arrow:level:${k}`);
    if (t) out.push({ id: `level:${k}`, kind: 'level', at: { x: t.x + t.w, y: t.y + t.h / 2 }, cursor: 'move', index: k });
  });
  for (const role of new Set(selected)) {
    if (kindOf(role) !== 'label') continue;
    const t = textOf(scene, role);
    if (t) out.push({ id: `label:${role}`, kind: 'label', at: centre(t), box: { x: t.x, y: t.y, w: t.w, h: t.h }, cursor: 'move', role });
  }
  return out;
}

const round = (v: number, step: number) => Math.round(v / step) * step;
/** A round step near `x`: 1, 2 or 5 times a power of ten. */
export function niceStep(x: number): number {
  if (!(x > 0) || !Number.isFinite(x)) return 1;
  const p = 10 ** Math.floor(Math.log10(x));
  return [1, 2, 5, 10].map((f) => f * p).find((s) => s >= x) ?? 10 * p;
}
const clean = (v: number) => Number(v.toPrecision(12));

function withOptions(spec: ChartSpec, patch: Record<string, unknown>): ChartSpec {
  const options: Record<string, unknown> = { ...(spec.options ?? {}) };
  for (const [k, v] of Object.entries(patch)) { if (v === undefined) delete options[k]; else options[k] = v; }
  const out = { ...spec, options } as ChartSpec;
  if (!Object.keys(options).length) delete out.options;
  return out;
}

/** The gap that makes the first column's edge follow the pointer (`dx`, `dy`: the drag so far). */
export function gapFor(spec: ChartSpec, scene: Scene, dx: number, dy: number): number | undefined {
  const c = firstCluster(scene);
  if (!c) return undefined;
  const hz = horizontal(spec);
  const width = hz ? c.h : c.w;
  const gap0 = typeof spec.options?.gap === 'number' ? spec.options.gap : 0.5;
  const slot = width * (1 + gap0);
  // The column stays centred: its edge moves by d, so its width by 2d.
  const w = Math.max(slot / (1 + LIMITS.gap.max), Math.min(slot, width + 2 * (hz ? dy : dx)));
  return clean(Math.min(LIMITS.gap.max, Math.max(LIMITS.gap.min, round(slot / w - 1, 0.05))));
}

/** The value under a point on the value axis (ValueScale.at), or undefined without one. */
export function valueAt(scene: Scene, x: number, y: number): number | undefined {
  const vs = scene.valueScale;
  if (!vs?.at) return undefined;
  const [x0, y0, x1, y1] = vs.at;
  const len2 = (x1 - x0) ** 2 + (y1 - y0) ** 2;
  if (!(len2 > 0)) return undefined;
  const t = ((x - x0) * (x1 - x0) + (y - y0) * (y1 - y0)) / len2;
  return vs.min + t * (vs.max - vs.min);
}

/**
 * The spec after dragging handle `h` by (dx, dy) slide points from where the drag started, on the
 * chart laid out as `scene` (the layout at drag start). `relayout` lays a candidate out (level
 * arrows choose a position; axisMax finds the data's own top).
 */
export function dragSpec(spec: ChartSpec, scene: Scene, h: DragHandle, dx: number, dy: number, relayout?: Relayout): ChartSpec {
  switch (h.kind) {
    case 'gap': {
      const gap = gapFor(spec, scene, dx, dy);
      return gap === undefined ? spec : withOptions(spec, { gap });
    }
    case 'axisMax': {
      const v = valueAt(scene, h.at.x + dx, h.at.y + dy);
      const vs = scene.valueScale;
      if (v === undefined || !vs) return spec;
      const step = niceStep((vs.max - vs.min) / 100);
      const want = clean(round(v, step));
      // Below the data's own top the option would be ignored: leave it out instead.
      const auto = relayout?.(withOptions(spec, { axisMax: undefined })).valueScale?.max;
      return withOptions(spec, { axisMax: auto !== undefined && want <= auto + step / 2 ? undefined : want });
    }
    case 'level': {
      const k = h.index ?? -1;
      const arrows = spec.options?.levelArrows;
      if (!arrows?.[k] || !relayout) return spec;
      const from = textOf(scene, `arrow:level:${k}`);
      if (!from) return spec;
      const target = { x: from.x + from.w / 2 + dx, y: from.y + from.h / 2 + dy };
      let best = spec, bestD = Infinity;
      for (const position of ['beside', 'between'] as const) {
        const next = withOptions(spec, { levelArrows: arrows.map((a, i) => (i === k ? { ...a, position } : a)) });
        const t = textOf(relayout(next), `arrow:level:${k}`);
        if (!t) continue;
        const d = Math.hypot(t.x + t.w / 2 - target.x, t.y + t.h / 2 - target.y);
        if (d < bestD - 0.01) { bestD = d; best = next; }
      }
      return best;
    }
    case 'label': {
      if (!h.role) return spec;
      const f = scene.frame;
      const fx = Math.min(1, Math.max(0, (h.at.x + dx - f.x) / f.w));
      const fy = Math.min(1, Math.max(0, (h.at.y + dy - f.y) / f.h));
      const overrides = { ...(spec.overrides ?? {}) };
      overrides[h.role] = { ...(overrides[h.role] ?? {}), pin: { x: round(fx, 1e-4), y: round(fy, 1e-4) } };
      return { ...spec, overrides };
    }
  }
}

/** The spec with what handle `h` sets removed (double-click: back to automatic). */
export function resetSpec(spec: ChartSpec, h: DragHandle): ChartSpec {
  switch (h.kind) {
    case 'gap': return spec.options?.gap === undefined ? spec : withOptions(spec, { gap: undefined });
    case 'axisMax': return spec.options?.axisMax === undefined ? spec : withOptions(spec, { axisMax: undefined });
    case 'level': {
      const arrows = spec.options?.levelArrows;
      const a = arrows?.[h.index ?? -1];
      if (!a || a.position === undefined) return spec;
      return withOptions(spec, { levelArrows: arrows!.map((x, i) => { if (i !== h.index) return x; const { position: _, ...rest } = x; return rest; }) });
    }
    case 'label': return h.role ? unpinElement(spec, h.role) : spec;
  }
}

/** A short note on what the drag sets, shown next to the pointer. */
export function describeDrag(spec: ChartSpec, h: DragHandle): string {
  const o = spec.options ?? {};
  switch (h.kind) {
    case 'gap': return `Gap ${Math.round((typeof o.gap === 'number' ? o.gap : 0.5) * 100)}%`;
    case 'axisMax': return typeof o.axisMax === 'number' ? `Axis max ${o.axisMax.toLocaleString('en-US', { maximumFractionDigits: 6 })}` : 'Axis max: auto';
    case 'level': return o.levelArrows?.[h.index ?? -1]?.position === 'between' ? 'Arrow: between' : 'Arrow: beside';
    case 'label': return 'Label placed by hand (double-click to reset)';
  }
}
