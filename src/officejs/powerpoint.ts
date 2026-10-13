// The PowerPoint side of the Office.js add-in: a thin adapter over PowerPoint JS. It reads the
// selected chart (its OC_* tags and bounds), draws a Path B plan as native shapes in one group, and
// does Path A (insertSlidesFromBase64) — on a new slide, or in place on the user's slide via
// exportAsBase64 (1.8). Batching rule (docs/plan/03 §4.2): never sync inside a loop; insert = 2
// syncs, update = 3.
import type { Frame } from '../core/types.ts';
import type { Box, HostCaps, Op } from './plan.ts';
import type { Tags } from '../render/tags.ts';
import { THEME_SLOTS, type ThemeColors } from '../core/theme.ts';

/* eslint-disable @typescript-eslint/no-explicit-any */
type PP = any;
declare const PowerPoint: PP;

export interface SelectedChart {
  slideId: string;
  /** Id of the chart group (the selected shape, or the group a selected child belongs to). */
  id: string;
  name: string;
  bounds: Frame;
  tags: Record<string, string>;
}

export type Selection =
  | { kind: 'none'; slideId?: string }
  | { kind: 'chart'; chart: SelectedChart }
  /** A shape that is not an open-charts chart (or several shapes). */
  | { kind: 'other'; slideId?: string; bounds?: Frame; count: number };

const frameOf = (s: { left: number; top: number; width: number; height: number }): Frame => ({ x: s.left, y: s.top, w: s.width, h: s.height });

const tagMap = (items: { key: string; value: string }[]) => Object.fromEntries(items.map((t) => [t.key.toUpperCase(), t.value]));
const isChart = (tags: Record<string, string>) => tags.OC_V !== undefined && (tags.OC_SPEC !== undefined || tags.OC_SPEC_N !== undefined);

/** The selected open-charts chart, if one (or one of its shapes) is selected. Two or three syncs. */
export async function readSelection(caps: HostCaps): Promise<Selection> {
  if (!caps.selection) return { kind: 'none' };
  return PowerPoint.run(async (ctx: PP) => {
    const slides = ctx.presentation.getSelectedSlides();
    slides.load('items/id');
    const shapes = ctx.presentation.getSelectedShapes();
    shapes.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height');
    await ctx.sync();
    const slideId: string | undefined = slides.items[0]?.id;
    if (shapes.items.length === 0) return { kind: 'none', slideId };
    if (shapes.items.length > 1) return { kind: 'other', slideId, count: shapes.items.length };
    const shape = shapes.items[0];
    shape.tags.load('items/key,items/value');
    // A shape inside a chart group: the chart is its parent group (1.8).
    let parent: PP = null;
    if (caps.groups && shape.type !== 'Group') {
      parent = shape.parentGroup;
      parent.load('id,name,left,top,width,height');
      parent.tags.load('items/key,items/value');
    }
    try { await ctx.sync(); } catch { parent = null; await syncTagsOnly(ctx, shape); }
    for (const s of [shape, parent]) {
      if (!s) continue;
      const tags = tagMap(s.tags.items ?? []);
      if (isChart(tags)) return { kind: 'chart', chart: { slideId: slideId ?? '', id: s.id, name: s.name, bounds: frameOf(s), tags } };
    }
    return { kind: 'other', slideId, bounds: frameOf(shape), count: 1 };
  });
}

/** After a failed parentGroup read (the shape is in no group), read the shape's tags alone. */
async function syncTagsOnly(ctx: PP, shape: PP) {
  shape.tags.load('items/key,items/value');
  await ctx.sync();
}

const THEME_NAMES: Record<string, string> = {
  dk1: 'Dark1', lt1: 'Light1', dk2: 'Dark2', lt2: 'Light2',
  accent1: 'Accent1', accent2: 'Accent2', accent3: 'Accent3', accent4: 'Accent4', accent5: 'Accent5', accent6: 'Accent6',
};

/** The current slide's theme colours (1.10), so theme slots lay out against what PowerPoint shows. */
export async function slideTheme(caps: HostCaps, slideId?: string): Promise<Partial<ThemeColors> | undefined> {
  if (!caps.rotation) return undefined;
  try {
    return await PowerPoint.run(async (ctx: PP) => {
      const slide = slideId ? ctx.presentation.slides.getItem(slideId) : ctx.presentation.getSelectedSlides().getItemAt(0);
      const results = THEME_SLOTS.map((k) => [k, slide.themeColorScheme.getThemeColor(THEME_NAMES[k])] as const);
      await ctx.sync();
      const out: Partial<ThemeColors> = {};
      for (const [k, r] of results) if (/^#?[0-9A-Fa-f]{6}$/.test(r.value ?? '')) out[k] = r.value.replace('#', '').toUpperCase();
      return Object.keys(out).length ? out : undefined;
    });
  } catch {
    return undefined;
  }
}

function addOp(shapes: PP, op: Op): PP {
  const opt = (b: Box) => ({ left: b.left, top: b.top, width: b.width, height: b.height });
  let s: PP;
  if (op.kind === 'line') {
    s = shapes.addLine('Straight', opt(op.box));
  } else if (op.kind === 'text') {
    s = shapes.addTextBox(op.text.text, opt(op.box));
  } else {
    s = shapes.addGeometricShape(op.geom, opt(op.box));
  }
  s.name = op.name;
  if (op.kind === 'line') {
    applyLine(s, op.line);
  } else {
    if (op.fill) s.fill.setSolidColor(op.fill); else s.fill.clear();
    applyLine(s, op.line);
  }
  if (op.kind !== 'line' && op.text) {
    if (op.kind === 'geom') s.textFrame.textRange.text = op.text.text;
    const tf = s.textFrame;
    tf.wordWrap = false;
    tf.autoSizeSetting = 'AutoSizeNone';
    tf.leftMargin = 0; tf.rightMargin = 0; tf.topMargin = 0; tf.bottomMargin = 0;
    tf.verticalAlignment = op.text.valign;
    const font = tf.textRange.font;
    font.name = op.text.font; font.size = op.text.size; font.bold = op.text.bold; font.color = op.text.color;
    tf.textRange.paragraphFormat.horizontalAlignment = op.text.align;
  }
  if (op.kind !== 'text' && op.rotation) s.rotation = op.rotation;
  if (op.kind === 'geom' && op.adjustments) op.adjustments.forEach((v, i) => s.adjustments.set(i, v));
  if (op.role) s.tags.add('OC_ROLE', op.role);
  return s;
}

function applyLine(s: PP, line: { color: string; weight: number; dash?: boolean } | null) {
  if (!line) { s.lineFormat.visible = false; return; }
  s.lineFormat.visible = true;
  s.lineFormat.color = line.color;
  s.lineFormat.weight = line.weight;
  s.lineFormat.dashStyle = line.dash ? 'Dash' : 'Solid';
}

export interface DrawChart {
  ops: Op[];
  frame: Frame;
  name: string;
  altText: string;
  tags: Tags;
}

/**
 * Path B: draws the chart as one group on the slide, the invisible frame rectangle first (so the
 * group's bounds are the frame, as the .ppam add-in draws it), then every shape with its OC_ROLE.
 * With `replaceId` the old chart is deleted in the same batch and the new group is sent back to
 * the old one's z-order position. Returns the new group's id.
 */
export async function drawPathB(caps: HostCaps, chart: DrawChart, target: { slideId?: string; replaceId?: string }): Promise<string> {
  return PowerPoint.run(async (ctx: PP) => {
    const slide = target.slideId ? ctx.presentation.slides.getItem(target.slideId) : ctx.presentation.getSelectedSlides().getItemAt(0);
    const shapes = slide.shapes;
    let old: PP = null;
    let z = -1;
    if (target.replaceId) {
      old = shapes.getItem(target.replaceId);
      old.load('zOrderPosition');
      await ctx.sync(); // sync 1 of 3 (update only)
      z = old.zOrderPosition;
    }
    const frame = shapes.addGeometricShape('Rectangle', { left: chart.frame.x, top: chart.frame.y, width: chart.frame.w, height: chart.frame.h });
    frame.name = 'Frame';
    frame.fill.clear();
    frame.lineFormat.visible = false;
    frame.tags.add('OC_ROLE', 'frame');
    const created = [frame, ...chart.ops.map((op) => addOp(shapes, op))];
    const group = shapes.addGroup(created);
    group.name = chart.name;
    if (caps.rotation) group.altTextDescription = chart.altText;
    for (const [k, v] of chart.tags) group.tags.add(k, v);
    if (old) old.delete();
    group.load('id,zOrderPosition');
    await ctx.sync();
    if (old && z >= 0) for (let k = group.zOrderPosition; k > z; k--) group.setZOrder('SendBackward');
    slide.setSelectedShapes([group.id]);
    await ctx.sync();
    return group.id as string;
  });
}

/** Path A, new slide: inserts the one-slide deck after `afterSlideId` (or the selected slide). */
export async function insertSlide(b64: string, afterSlideId?: string): Promise<void> {
  await PowerPoint.run(async (ctx: PP) => {
    let target = afterSlideId;
    if (!target) {
      const sel = ctx.presentation.getSelectedSlides();
      sel.load('items/id');
      await ctx.sync();
      target = sel.items[0]?.id;
    }
    ctx.presentation.insertSlidesFromBase64(b64, { formatting: 'UseDestinationTheme', ...(target ? { targetSlideId: target } : {}) });
    await ctx.sync();
  });
}

/** A name no user shape has: marks the chart to replace in the exported slide's XML. */
export const REPLACE_MARKER = 'open-charts:replace-me';

/**
 * Path A in place (1.8): exports the slide (renaming the old chart to the marker first), lets
 * `patch` put the new chart into it, inserts the patched slide after the original, deletes the
 * original and selects the new slide. The slide gets a new id; everything else on it is kept.
 */
export async function replaceSlide(slideId: string, replace: { id: string; name: string } | undefined, chart: { name: string; frameTag: string }, patch: (b64: string) => Promise<string>): Promise<void> {
  const exported: string = await PowerPoint.run(async (ctx: PP) => {
    const slide = ctx.presentation.slides.getItem(slideId);
    if (replace) slide.shapes.getItem(replace.id).name = REPLACE_MARKER;
    const r = slide.exportAsBase64();
    await ctx.sync();
    return r.value;
  });
  let patched: string;
  try {
    patched = await patch(exported);
  } catch (e) {
    if (replace) {
      await PowerPoint.run(async (ctx: PP) => {
        ctx.presentation.slides.getItem(slideId).shapes.getItem(replace.id).name = replace.name;
        await ctx.sync();
      });
    }
    throw e;
  }
  await PowerPoint.run(async (ctx: PP) => {
    const slides = ctx.presentation.slides;
    slides.load('items/id');
    await ctx.sync();
    const at = slides.items.findIndex((s: PP) => s.id === slideId);
    ctx.presentation.insertSlidesFromBase64(patched, { formatting: 'UseDestinationTheme', targetSlideId: slideId });
    ctx.presentation.slides.getItem(slideId).delete();
    await ctx.sync();
    const after = ctx.presentation.slides;
    after.load('items/id');
    await ctx.sync();
    const fresh = after.items[Math.max(0, at)];
    if (!fresh) return;
    ctx.presentation.setSelectedSlides([fresh.id]);
    // Select the chart again, so the pane stays bound to it: the group with its name and OC_FRAME.
    const shapes = fresh.shapes;
    shapes.load('items/id,items/name');
    await ctx.sync();
    const named = shapes.items.filter((s: PP) => s.name === chart.name);
    const frames = named.map((s: PP) => s.tags.getItemOrNullObject('OC_FRAME').load('value'));
    await ctx.sync();
    const hit = named.find((_: PP, i: number) => !frames[i].isNullObject && frames[i].value === chart.frameTag);
    if (hit) fresh.setSelectedShapes([hit.id]);
    await ctx.sync();
  });
}
