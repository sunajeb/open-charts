// Per-element formatting (DATA-FORMAT.md "Overrides"), after the reference product's floating toolbar
// or, as a convenience, a selector (`label:value:s1:*`, `bar:*:c2`, `label:category:*`) — to an
// ElementStyle. Which properties an element takes depends on its kind (ELEMENT_PROPERTIES, the one
// applicability table behind validation, the docs and, later, the toolbar). Overrides live in the
// model (OC_SPEC) and the engine applies them while it lays the chart out — a label is measured,
// placed and collision-checked at its overridden content, size and direction — so they survive every
// re-layout (data edits, resize, Update). Formatting done with PowerPoint's own tools is re-asserted
// on the next update; `detach` ("convert to plain text") hands an element over to PowerPoint for good.
// An override that matches nothing drawn (a category removed since) is a warning, never an error.
import type { ChartSpec, Shape, TextShape } from './types.ts';
import type { Measurer } from './text.ts';
import { isElementColor, resolveElementColor, type LinkColors, type ThemeColors } from './theme.ts';

// ---- The model --------------------------------------------------------------------------------

/** Data a label field shows: value, share, series / category name, or the total of its column. */
export const LABEL_FIELDS = ['value', 'percent', 'series', 'category', 'total'] as const;
export type LabelField = typeof LABEL_FIELDS[number];
/** Number format of one numeric field (value, percent, total). */
export interface NumberField { field: 'value' | 'percent' | 'total'; decimals?: number; thousands?: boolean; sign?: boolean }
/** One field of a label: data-linked (a name, or a number with its own format) or custom text. */
export type ContentField = LabelField | NumberField | { text: string };

export type Dash = 'solid' | 'dash' | 'dot';
export type Arrowheads = 'none' | 'start' | 'end' | 'both';
export type MarkerStyle = 'circle' | 'square' | 'diamond' | 'triangle' | 'none';

/** Formatting of one element. Every key is optional; ELEMENT_PROPERTIES says which element kinds read each. */
export interface ElementStyle {
  /** Label content: an ordered list of fields, joined with `separator`, e.g. ["category", "value", {"field": "percent", "decimals": 1}]. */
  content?: ContentField[];
  /** Between fields; default one space. */
  separator?: string;
  /** Number precision of the label's numbers (a field's own `decimals` wins). */
  decimals?: number;
  /** Font: any installed font name, or the theme's heading (`+mj-lt`) / body (`+mn-lt`) font. */
  font?: string;
  /** Font size, points 4–72 (this element only). */
  size?: number;
  /** Text colour. */
  color?: string;
  bold?: boolean;
  align?: 'left' | 'center' | 'right';
  /** Text direction: -90 = turned 90° left (reads bottom to top), 0 = horizontal, 90 = turned 90° right. */
  rotate?: -90 | 0 | 90;
  /** Fill: a segment, bar, area, slice or marker; the swatch of a legend entry; the box behind a label. "none" = no fill. */
  fill?: string;
  /** Outline / border colour, or "none". */
  outline?: string;
  outlineWidth?: number;
  outlineDash?: Dash;
  /** Line colour (lines, connectors, arrows, axes, gridlines). */
  line?: string;
  lineWidth?: number;
  dash?: Dash;
  arrowheads?: Arrowheads;
  /** Marker shape of a line / scatter point. */
  marker?: MarkerStyle;
  /** Convert to plain text: the engine stops drawing the element; the host leaves a plain copy that PowerPoint owns. */
  detach?: boolean;
}
export type Overrides = Record<string, ElementStyle>;

/** What a role is, for the property table: the first segment of the role decides. */
export type ElementKind = 'label' | 'legend' | 'segment' | 'marker' | 'line' | 'decoration' | 'axis';
export const ROLE_KINDS: Record<string, ElementKind> = {
  label: 'label', legend: 'legend',
  bar: 'segment', wedge: 'segment', area: 'segment', ball: 'segment', shade: 'segment',
  marker: 'marker',
  line: 'line', connector: 'line', leader: 'line', today: 'line',
  arrow: 'decoration', valueline: 'decoration', indicator: 'decoration', bracket: 'decoration',
  axis: 'axis', gridline: 'axis',
};
export const ELEMENT_KINDS = ['label', 'legend', 'segment', 'marker', 'line', 'decoration', 'axis'] as const satisfies readonly ElementKind[];

export interface PropertyInfo {
  /** Element kinds that read the property; on any other it is a validation error. */
  kinds: readonly ElementKind[];
  /** Toolbar group (the reference's MiniToolbarInfo factory the control belongs to). */
  group: 'content' | 'precision' | 'font' | 'fontColor' | 'alignment' | 'rotation' | 'fill' | 'outline' | 'lineStyle' | 'marker' | 'plainText';
  /** The value's shape, as documented. */
  value: string;
}

const TEXT_KINDS = ['label', 'legend', 'decoration', 'axis'] as const;
/**
 * Every element property and the element kinds that read it: the one table behind validation, the
 * schema, DATA-FORMAT.md and the toolbar the editor and add-in will build from it.
 */
export const ELEMENT_PROPERTIES = {
  content: { kinds: ['label'], group: 'content', value: 'ordered array of fields: `value` `percent` `series` `category` `total`, `{field, decimals?, thousands?, sign?}`, `{text}`' },
  separator: { kinds: ['label'], group: 'content', value: 'string between fields (default one space)' },
  decimals: { kinds: ['label'], group: 'precision', value: 'whole number 0–10' },
  font: { kinds: TEXT_KINDS, group: 'font', value: 'font name, or `+mj-lt` / `+mn-lt` (theme heading / body font)' },
  size: { kinds: TEXT_KINDS, group: 'font', value: 'points 4–72' },
  bold: { kinds: TEXT_KINDS, group: 'font', value: 'boolean' },
  color: { kinds: TEXT_KINDS, group: 'fontColor', value: 'colour' },
  align: { kinds: ['label', 'axis'], group: 'alignment', value: '`left` `center` `right`' },
  rotate: { kinds: ['label', 'axis'], group: 'rotation', value: '`-90` (90° left) `0` `90` (90° right)' },
  fill: { kinds: ['segment', 'marker', 'label', 'legend', 'decoration'], group: 'fill', value: 'colour or `none`' },
  outline: { kinds: ['segment', 'marker', 'label', 'decoration'], group: 'outline', value: 'colour or `none`' },
  outlineWidth: { kinds: ['segment', 'marker', 'label', 'decoration'], group: 'outline', value: 'points 0–20' },
  outlineDash: { kinds: ['segment', 'marker'], group: 'outline', value: '`solid` `dash` `dot`' },
  line: { kinds: ['line', 'decoration', 'axis'], group: 'lineStyle', value: 'colour' },
  lineWidth: { kinds: ['line', 'decoration', 'axis'], group: 'lineStyle', value: 'points 0.25–20' },
  dash: { kinds: ['line', 'decoration', 'axis'], group: 'lineStyle', value: '`solid` `dash` `dot`' },
  arrowheads: { kinds: ['line', 'decoration'], group: 'lineStyle', value: '`none` `start` `end` `both`' },
  marker: { kinds: ['marker'], group: 'marker', value: '`circle` `square` `diamond` `triangle` `none`' },
  detach: { kinds: ['label'], group: 'plainText', value: '`true`: convert to plain text' },
} as const satisfies Record<keyof ElementStyle, PropertyInfo>;
export type ElementProperty = keyof typeof ELEMENT_PROPERTIES;
export const ELEMENT_STYLE_KEYS = Object.keys(ELEMENT_PROPERTIES) as ElementProperty[];

/** The kind of element a role (or selector) names, from its first segment. */
export const kindOf = (role: string): ElementKind | undefined => ROLE_KINDS[role.split(':')[0]];

/** Properties an element kind reads (the toolbar for a selection is the union over the selected elements). */
export const propertiesFor = (kind: ElementKind): ElementProperty[] =>
  ELEMENT_STYLE_KEYS.filter((k) => (ELEMENT_PROPERTIES[k].kinds as readonly string[]).includes(kind));

// ---- Validation --------------------------------------------------------------------------------

/** A role segment: a word, `s<i>`, `c<j>`, a number (ticks, gridlines) or `*` (any one segment; trailing: the rest). */
const SEGMENT = /^([a-z]+|s\d+|c\d+|\d+|\*)$/;

export interface OverrideProblem { path: (string | number)[]; message: string }

const isObj = (v: unknown): v is Record<string, unknown> => !!v && typeof v === 'object' && !Array.isArray(v);
const isInt = (v: unknown): v is number => typeof v === 'number' && Number.isInteger(v);
const describe = (v: unknown) =>
  v === null ? 'null' : typeof v === 'number' ? String(v) : typeof v === 'string' ? `string "${v}"` : Array.isArray(v) ? 'array' : typeof v;
const COLOUR = 'a hex colour, a theme colour (accent1…accent6, tx1, tx2, bg1, bg2, dk1, lt1, dk2, lt2, hlink, folHlink) or one with a tint / shade like "accent1-lighter40" / "accent1-darker25"';
const DASHES = ['solid', 'dash', 'dot'];

/**
 * Problems in an `overrides` object; paths relative to the chart ("overrides", key, prop, …). Keys are
 * checked for syntax and element kind only: an index past the data is a warning at layout time (data
 * edits must not make a chart invalid).
 */
export function checkOverrides(o: unknown): OverrideProblem[] {
  const out: OverrideProblem[] = [];
  const add = (path: (string | number)[], message: string) => out.push({ path: ['overrides', ...path], message });
  if (!isObj(o)) { add([], 'must be an object of {role: style}'); return out; }
  for (const [key, st] of Object.entries(o)) {
    const segs = key.split(':');
    const kind = kindOf(key);
    if (!segs.every((s) => SEGMENT.test(s))) {
      add([key], 'must be a role like "label:value:s1:c2", or a selector with "*", e.g. "label:value:s1:*", "bar:*:c2", "label:category:*"');
    } else if (!kind) add([key], `"${segs[0]}" elements cannot be formatted (only ${Object.keys(ROLE_KINDS).join(', ')})`);
    if (!isObj(st)) { add([key], 'must be an object, e.g. {"color": "C0504D", "size": 14}'); continue; }
    for (const k of Object.keys(st)) {
      const info = (ELEMENT_PROPERTIES as Record<string, PropertyInfo>)[k];
      if (!info) add([key, k], `unknown property "${k}"`);
      else if (kind && !info.kinds.includes(kind)) add([key, k], `not used by ${kind} elements (only ${info.kinds.join(', ')})`);
    }
    const p = (k: string, ...rest: (string | number)[]) => [key, k, ...rest];
    if (st.content !== undefined) {
      if (!Array.isArray(st.content) || !st.content.length) add(p('content'), 'must be a non-empty array of fields, e.g. ["value"] or ["category", "value"]');
      else st.content.forEach((f, i) => {
        if (typeof f === 'string') { if (!(LABEL_FIELDS as readonly string[]).includes(f)) add(p('content', i), `must be one of ${LABEL_FIELDS.join(', ')}, {field, decimals, thousands, sign} or {text}, got ${describe(f)}`); return; }
        if (!isObj(f)) { add(p('content', i), `must be a field name, {field, …} or {text}, got ${describe(f)}`); return; }
        if ('text' in f) {
          if (typeof f.text !== 'string') add(p('content', i, 'text'), 'must be a string');
          for (const k of Object.keys(f)) if (k !== 'text') add(p('content', i, k), `unknown property "${k}"`);
          return;
        }
        for (const k of Object.keys(f)) if (!['field', 'decimals', 'thousands', 'sign'].includes(k)) add(p('content', i, k), `unknown property "${k}"`);
        if (f.field !== 'value' && f.field !== 'percent' && f.field !== 'total') add(p('content', i, 'field'), `must be value, percent or total, got ${describe(f.field)}`);
        if (f.decimals !== undefined && !(isInt(f.decimals) && f.decimals >= 0 && f.decimals <= 10)) add(p('content', i, 'decimals'), `must be a whole number 0–10, got ${describe(f.decimals)}`);
        for (const k of ['thousands', 'sign']) if (f[k] !== undefined && typeof f[k] !== 'boolean') add(p('content', i, k), `must be true or false, got ${describe(f[k])}`);
      });
    }
    if (st.separator !== undefined && typeof st.separator !== 'string') add(p('separator'), 'must be a string');
    if (st.decimals !== undefined && !(isInt(st.decimals) && st.decimals >= 0 && st.decimals <= 10)) add(p('decimals'), `must be a whole number 0–10, got ${describe(st.decimals)}`);
    if (st.font !== undefined && (typeof st.font !== 'string' || !st.font)) add(p('font'), 'must be a font name, or +mj-lt / +mn-lt');
    if (st.size !== undefined && (typeof st.size !== 'number' || !(st.size >= 4 && st.size <= 72))) add(p('size'), `must be a number of points 4–72, got ${describe(st.size)}`);
    for (const k of ['bold', 'detach']) if (st[k] !== undefined && typeof st[k] !== 'boolean') add(p(k), `must be true or false, got ${describe(st[k])}`);
    if (st.align !== undefined && !['left', 'center', 'right'].includes(st.align as string)) add(p('align'), `must be left, center or right, got ${describe(st.align)}`);
    if (st.rotate !== undefined && st.rotate !== -90 && st.rotate !== 0 && st.rotate !== 90) add(p('rotate'), `must be -90 (90° left), 0 or 90 (90° right), got ${describe(st.rotate)}`);
    for (const k of ['color', 'line']) if (st[k] !== undefined && !isElementColor(st[k])) add(p(k), `must be ${COLOUR}, got ${describe(st[k])}`);
    for (const k of ['fill', 'outline']) if (st[k] !== undefined && st[k] !== 'none' && !isElementColor(st[k])) add(p(k), `must be "none" or ${COLOUR}, got ${describe(st[k])}`);
    if (st.outlineWidth !== undefined && (typeof st.outlineWidth !== 'number' || !(st.outlineWidth >= 0 && st.outlineWidth <= 20))) add(p('outlineWidth'), `must be a number of points 0–20, got ${describe(st.outlineWidth)}`);
    if (st.lineWidth !== undefined && (typeof st.lineWidth !== 'number' || !(st.lineWidth >= 0.25 && st.lineWidth <= 20))) add(p('lineWidth'), `must be a number of points 0.25–20, got ${describe(st.lineWidth)}`);
    for (const k of ['dash', 'outlineDash']) if (st[k] !== undefined && !DASHES.includes(st[k] as string)) add(p(k), `must be solid, dash or dot, got ${describe(st[k])}`);
    if (st.arrowheads !== undefined && !['none', 'start', 'end', 'both'].includes(st.arrowheads as string)) add(p('arrowheads'), `must be none, start, end or both, got ${describe(st.arrowheads)}`);
    if (st.marker !== undefined && !['circle', 'square', 'diamond', 'triangle', 'none'].includes(st.marker as string)) add(p('marker'), `must be circle, square, diamond, triangle or none, got ${describe(st.marker)}`);
  }
  return out;
}

// ---- Matching ----------------------------------------------------------------------------------

/** Whether a key (role or selector) matches a role. `*` matches one segment; a trailing `*` the rest. */
export function matches(key: string, role: string): boolean {
  if (key === role) return true;
  if (!key.includes('*')) return false;
  const k = key.split(':'), r = role.split(':');
  for (let i = 0; i < k.length; i++) {
    if (k[i] === '*' && i === k.length - 1) return r.length >= k.length;
    if (i >= r.length || (k[i] !== '*' && k[i] !== r[i])) return false;
  }
  return k.length === r.length;
}

/** Precedence: an exact role beats any selector; among selectors, more named segments win (series / category over kind). */
export const specificity = (key: string): number =>
  key.includes('*') ? key.split(':').filter((s) => s !== '*').length : Infinity;

/**
 * Lookup of the merged style of a role: every matching key contributes, least specific first (ties:
 * the key written later wins), so `label:value:*` {size} and `label:value:s1:c2` {color} both apply.
 */
export function compileOverrides(o: Overrides | undefined): (role: string) => ElementStyle | undefined {
  const keys = Object.keys(o ?? {}).map((k, i) => ({ k, i, rank: specificity(k) })).sort((a, b) => a.rank - b.rank || a.i - b.i);
  if (!keys.length) return () => undefined;
  const cache = new Map<string, ElementStyle | undefined>();
  return (role) => {
    if (cache.has(role)) return cache.get(role);
    let st: ElementStyle | undefined;
    for (const { k } of keys) if (matches(k, role)) st = { ...st, ...o![k] };
    cache.set(role, st);
    return st;
  };
}

// ---- Applying overrides during layout ---------------------------------------------------------

/** What the layout needs to apply overrides (`Ctx.ov`). Roles it is asked about use visible series indices. */
export interface OverrideCtx {
  get(role: string): ElementStyle | undefined;
  theme: ThemeColors & LinkColors;
}

/** Data a label's content is computed from: the chart as laid out, and its number formatting. */
export interface ContentSource {
  spec: ChartSpec;
  /** Format an (unscaled) value as the chart does, with optional decimals / format / sign. */
  fmt(v: number, o?: { decimals?: number; format?: string; signed?: boolean }): string;
  pct(v: number, decimals: number, signed?: boolean): string;
  /** The chart's decimals for values, and for shares. */
  decimals: number;
  pctDecimals: number;
}

const num = (v: unknown): number | null => (typeof v === 'number' && Number.isFinite(v) ? v : null);
const idx = (role: string, p: 's' | 'c') => { const m = role.match(new RegExp(`(?:^|:)${p}(\\d+)(?=:|$)`)); return m ? +m[1] : null; };

/** The fields a label shows by default (what the chart type draws), for `decimals` without `content`. */
function defaultFields(spec: ChartSpec, role: string): ContentField[] | null {
  if (role.startsWith('label:category')) return ['category'];
  if (role.startsWith('label:series')) return ['series'];
  if (role.startsWith('label:total')) return ['total'];
  if (!role.startsWith('label:value')) return null;
  const pie = spec.type === 'pie' || spec.type === 'doughnut';
  const lc = spec.options?.labelContent ?? 'percent';
  if (pie) return lc === 'value' ? ['value'] : lc === 'both' ? ['value', 'percent'] : ['percent'];
  if (spec.type === 'percent' && !spec.options?.percentData) return ['percent'];
  return ['value'];
}

/**
 * The label text for an override (null: keep the chart's own text). Numeric fields take the field's
 * format, else the override's `decimals`, else the chart's; a field without data for the element
 * (a series name on a category label) is left out. `own` stands in for a value the data cannot give
 * (a waterfall total cell).
 */
export function labelContent(src: ContentSource, role: string, st: ElementStyle, own: string): string | null {
  const fields = st.content ?? (st.decimals !== undefined ? defaultFields(src.spec, role) : null);
  if (!fields) return null;
  const { spec } = src;
  const s = idx(role, 's'), k = idx(role, 'c');
  const cell = (i: number, j: number) => num(spec.series[i]?.values[j]);
  const sum = (xs: (number | null)[]) => xs.reduce<number>((a, v) => a + (v ?? 0), 0);
  const pie = spec.type === 'pie' || spec.type === 'doughnut';
  const colTotal = (j: number) => sum(spec.series.map((_, i) => cell(i, j)));
  const grand = sum(spec.series.flatMap((x) => x.values.map(num)));
  let value: number | null, whole: number | null;
  if (s != null && k != null) { value = cell(s, k); whole = pie ? sum(spec.categories.map((_, j) => cell(s, j))) : colTotal(k); }
  else if (k != null) { value = colTotal(k); whole = grand; }
  else if (s != null) { value = sum(spec.categories.map((_, j) => cell(s, j))); whole = grand; }
  else { value = grand; whole = grand; }
  const total = k != null ? (pie && s != null ? whole : colTotal(k)) : grand;
  const out: string[] = [];
  for (const f of fields) {
    if (typeof f === 'object' && 'text' in f) { out.push(f.text); continue; }
    const nf: NumberField | null = typeof f === 'object' ? f : f === 'value' || f === 'percent' || f === 'total' ? { field: f } : null;
    if (!nf) {
      if (f === 'series' && s != null && spec.series[s]) out.push(spec.series[s].name);
      else if (f === 'category' && k != null && spec.categories[k] !== undefined) out.push(spec.categories[k]);
      continue;
    }
    const percent = nf.field === 'percent';
    const v = nf.field === 'total' ? total : percent ? (value != null && whole ? (value / whole) * 100 : null) : value;
    if (v == null) { if (nf.field === 'value') out.push(own); continue; }
    const d = nf.decimals ?? st.decimals;
    if (percent) { out.push(src.pct(v, d ?? src.pctDecimals, nf.sign)); continue; }
    if (d === undefined && nf.thousands === undefined && !nf.sign) { out.push(src.fmt(v)); continue; }
    // An explicit precision or separator replaces the chart's number format for this field.
    const dd = d ?? src.decimals;
    const format = nf.thousands === undefined ? undefined : `${nf.thousands ? '#,##0' : '0'}${dd > 0 ? `.${'0'.repeat(dd)}` : ''}`;
    out.push(src.fmt(v, { decimals: dd, format, ...(nf.sign ? { signed: true } : {}) }));
  }
  return out.length ? out.join(st.separator ?? ' ') : null;
}

/** Text measured as drawn: width × height of its lines, before any rotation. */
export function textExtent(t: Pick<TextShape, 'text' | 'size' | 'bold' | 'lineHeight'>, m: Measurer) {
  const lines = t.text.split('\n');
  return { w: Math.max(...lines.map((l) => m.width(l, t.size, t.bold))), h: lines.length * (t.lineHeight ?? m.lineHeight(t.size)) };
}

const paint = (c: string, ov: OverrideCtx) => resolveElementColor(c, ov.theme);

/**
 * A text shape with its role's override applied: content, font, size, bold, colour, alignment,
 * direction, and — unless the role also has a mark (a legend's swatch takes it) — the box's fill and
 * outline. The box grows to the new text from the side its alignment anchors (left-aligned text grows
 * right, bottom-anchored text grows up, centred text both ways); turned text gets the box of its footprint.
 */
export function styleText(t: TextShape, st: ElementStyle, role: string, src: ContentSource, ov: OverrideCtx, m: Measurer, hasMark: boolean): TextShape {
  const out: TextShape = { ...t };
  const content = labelContent(src, role, st, t.text);
  if (content != null) out.text = content;
  if (st.size !== undefined) { out.size = st.size; out.lineHeight = m.lineHeight(st.size); }
  if (st.bold !== undefined) { if (st.bold) out.bold = true; else delete out.bold; }
  if (st.font !== undefined) out.font = st.font;
  if (st.color !== undefined) out.color = paint(st.color, ov);
  if (!hasMark) {
    if (st.fill !== undefined) { if (st.fill === 'none') delete out.fill; else out.fill = paint(st.fill, ov); }
    if (st.outline !== undefined) { if (st.outline === 'none') delete out.stroke; else out.stroke = paint(st.outline, ov); }
    if (st.outlineWidth !== undefined && out.stroke) out.strokeWidth = st.outlineWidth;
  }
  const ext = textExtent(out, m);
  if (st.rotate === 90 || st.rotate === -90) {
    out.rotate = st.rotate;
    // The footprint of the turned text, anchored where the text was: its top for top-anchored text, …
    const w = ext.h, h = ext.w + 2;
    out.x = t.align === 'left' ? t.x : t.align === 'right' ? t.x + t.w - w : t.x + t.w / 2 - w / 2;
    out.y = t.valign === 'top' ? t.y : t.valign === 'bottom' ? t.y + t.h - h : t.y + t.h / 2 - h / 2;
    out.w = w; out.h = h;
    out.align = 'center'; out.valign = 'middle';
    return out;
  }
  if (st.rotate === 0) delete out.rotate;
  const w = Math.max(t.w, ext.w + 2), h = Math.max(t.h, ext.h);
  if (w > t.w) out.x = t.align === 'left' ? t.x : t.align === 'right' ? t.x + t.w - w : t.x + t.w / 2 - w / 2;
  if (h > t.h) out.y = t.valign === 'top' ? t.y : t.valign === 'bottom' ? t.y + t.h - h : t.y + t.h / 2 - h / 2;
  out.w = w; out.h = h;
  if (st.align) out.align = st.align;
  return out;
}

const GEOM: Record<Exclude<MarkerStyle, 'none'>, 'ellipse' | 'rect' | 'diamond' | 'triangle'> = { circle: 'ellipse', square: 'rect', diamond: 'diamond', triangle: 'triangle' };

/**
 * A mark with its role's override applied: fill, outline and marker shape on rects and paths; line
 * colour, width, dash and arrowheads on lines (a label's leader line keeps its look).
 */
export function styleMark(s: Shape, st: ElementStyle, role: string, ov: OverrideCtx): Shape {
  if (s.kind === 'text') return s;
  if (s.kind === 'line') {
    if (role.startsWith('label:') || role.startsWith('legend:')) return s;
    const out = { ...s };
    if (st.line !== undefined) out.color = paint(st.line, ov);
    if (st.lineWidth !== undefined) out.width = st.lineWidth;
    if (st.dash !== undefined) { if (st.dash === 'solid') delete out.dash; else out.dash = st.dash === 'dot' ? 'dot' : true; }
    if (st.arrowheads !== undefined) {
      delete out.arrowEnd; delete out.arrowStart;
      if (st.arrowheads === 'end' || st.arrowheads === 'both') out.arrowEnd = true;
      if (st.arrowheads === 'start' || st.arrowheads === 'both') out.arrowStart = true;
    }
    return out;
  }
  const out = { ...s };
  if (st.fill !== undefined) { if (st.fill === 'none') out.fill = ''; else out.fill = paint(st.fill, ov); }
  if (st.outline !== undefined) {
    if (st.outline === 'none') { delete out.stroke; delete out.strokeWidth; } else out.stroke = paint(st.outline, ov);
  }
  if (st.outlineWidth !== undefined && out.stroke) out.strokeWidth = st.outlineWidth;
  if (st.outlineDash !== undefined) { if (st.outlineDash === 'solid') delete out.dash; else out.dash = st.outlineDash === 'dot' ? 'dot' : true; }
  if (out.kind === 'rect' && st.marker !== undefined && role.startsWith('marker:')) {
    if (st.marker === 'none') { out.fill = ''; delete out.stroke; delete out.strokeWidth; }
    else { out.geom = GEOM[st.marker]; delete out.ellipse; }
  }
  return out as Shape;
}

/**
 * Overrides that matched nothing drawn, one warning each: an index past the data (a series or
 * category removed since), or an element the chart does not draw. Keys on hidden series and detached
 * elements are not reported. `seen` holds every role the layout asked about.
 */
export function unmatched(o: Overrides | undefined, seen: Set<string>, spec: ChartSpec): string[] {
  const out: string[] = [];
  const roles = [...seen];
  for (const [key, st] of Object.entries(o ?? {})) {
    if (st?.detach || roles.some((r) => matches(key, r))) continue;
    const s = idx(key, 's'), c = idx(key, 'c');
    if (s != null && spec.series[s]?.hidden) continue;
    const why = s != null && s >= spec.series.length ? `there is no series ${s} (the chart has ${spec.series.length})`
      : c != null && c >= spec.categories.length ? `there is no category ${c} (the chart has ${spec.categories.length})`
        : 'the chart draws no such element';
    out.push(`override "${key}" matches nothing: ${why}`);
  }
  return out;
}

/**
 * Convert an element to plain text: the spec with `detach: true` on the role (the engine no longer
 * draws it, so re-layouts leave the plain copy alone), plus the element's shapes as last drawn in
 * `scene`, for the host to place outside the chart group once. The input spec is not changed.
 */
export function detachElement<T extends ChartSpec>(spec: T, scene: { shapes: Shape[]; roles?: string[] }, role: string): { spec: T; shapes: Shape[] } {
  const shapes = scene.shapes.filter((_, i) => scene.roles?.[i] === role);
  const overrides = { ...(spec.overrides ?? {}), [role]: { ...(spec.overrides?.[role] ?? {}), detach: true } };
  return { spec: { ...spec, overrides }, shapes };
}

/**
 * Overrides after a series (`s`) or category (`c`) was inserted, deleted or moved: every key naming
 * index i is renamed to map(i); keys whose element was deleted (map → -1) are dropped, so each override
 * stays on its element through data edits. Selectors with `*` in that position are kept as they are.
 * (The editor runs the same rule on its datasheet's insert / delete.)
 */
export function remapOverrides(o: Overrides | undefined, p: 's' | 'c', map: (i: number) => number): Overrides | undefined {
  if (!o) return o;
  const out: Overrides = {};
  const re = new RegExp(`(^|:)${p}(\\d+)(?=:|$)`);
  for (const [key, st] of Object.entries(o)) {
    const m = re.exec(key);
    if (!m) { out[key] = st; continue; }
    const j = map(+m[2]);
    if (j >= 0) out[key.replace(re, `$1${p}${j}`)] = st;
  }
  return out;
}
