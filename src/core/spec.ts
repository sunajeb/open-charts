// Spec versioning and validation. Pure functions; the JSON Schema in schema/*.schema.json mirrors
// these rules (a sync test keeps the key lists identical).
import { AGENDA_OPTION_KEYS, AGENDA_PRESETS, AGENDA_STATES, HEADING_STYLE_KEYS } from './agenda.ts';
import type { ChartOptions, ChartSpec } from './types.ts';
import { LOCALES } from './numbers.ts';
import { COLOR_NAMES, isColor } from './theme.ts';
import { parseCategoryDate } from './dates.ts';
import { checkOverrides } from './overrides.ts';

export const SPEC_VERSION = 1;

export interface SpecError {
  /** JSON Pointer into the spec, e.g. "/series/1/values/3" ("" = the root). */
  path: string;
  message: string;
}

export const CHART_TYPES = [
  'stacked', 'clustered', 'percent', 'waterfall', 'mekko', 'line',
  'area', 'area100', 'pie', 'doughnut', 'harvey',
  'combo', 'mekkoxy', 'scatter', 'bubble', 'table', 'gantt',
] as const;
export const CHART_KEYS = ['$schema', 'specVersion', 'type', 'categories', 'series', 'options', 'style', 'overrides'] as const;
export const SERIES_KEYS = ['name', 'values', 'color', 'as', 'hidden'] as const;
export type ChartTypeName = typeof CHART_TYPES[number];
const except = (...types: ChartTypeName[]) => CHART_TYPES.filter((t) => !types.includes(t));

/**
 * The direction a chart grows in (the reference's insert arrows ▲ ▼ ◀ ▶, manual: Rotate and flip
 * elements): `up` columns from the baseline (default), `down` columns hanging from the top, `right`
 * bars growing to the right (categories on the left), `left` bars growing to the left (categories on
 * the right). A line chart turned sideways is a profile chart; lines and areas are never hung upside
 * down (a value axis growing downwards reads as a loss).
 */
export const DIRECTIONS = {
  stacked: ['up', 'down', 'left', 'right'],
  clustered: ['up', 'down', 'left', 'right'],
  percent: ['up', 'down', 'left', 'right'],
  waterfall: ['up', 'down', 'left', 'right'],
  mekko: ['up', 'down', 'left', 'right'],
  line: ['up', 'left', 'right'],
  area: ['up', 'left', 'right'],
  area100: ['up', 'left', 'right'],
  combo: ['up', 'down', 'left', 'right'],
} as const satisfies Partial<Record<ChartTypeName, readonly Direction[]>>;
export type Direction = 'up' | 'down' | 'left' | 'right';
/** Directions a chart type can be drawn in (`['up']` for the types without arrows). */
export const directionsOf = (type: string): readonly Direction[] => (DIRECTIONS as Record<string, readonly Direction[]>)[type] ?? ['up'];

/** Inclusive numeric limits, shared by the validator, the JSON Schema and the editor's inputs. */
export const LIMITS = {
  decimals: { min: 0, max: 10, step: 1 },
  gap: { min: 0, max: 10, step: 0.1 },
} as const;

export interface OptionInfo {
  /** Chart types that read the option; on any other type it is a validation error. */
  types: readonly ChartTypeName[];
  /** The value's shape, as documented (DATA-FORMAT.md options table). */
  value: string;
  /** What the engine does when the option is left out and `optionDefaults` has no fixed value for it. */
  auto?: string;
}

/**
 * Every chart option and the chart types that read it: the one applicability table behind
 * validation, the schema descriptions, the editor's controls and the DATA-FORMAT.md options table
 * (tests keep the schema and the doc in sync with it).
 */
export const OPTIONS = {
  unit: { types: except('harvey', 'gantt'), value: 'string or `{text, position, space}`', auto: 'none' },
  decimals: { types: except('harvey', 'gantt'), value: `whole number ${LIMITS.decimals.min}–${LIMITS.decimals.max}`, auto: 'inferred from the data' },
  format: { types: except('harvey', 'gantt'), value: 'Excel number format', auto: 'none' },
  locale: { types: except('harvey'), value: '`en-US` `en-GB` `en-IN` `de-DE` `fr-FR`', auto: '`en-US`' },
  scale: { types: except('harvey', 'gantt'), value: '`none` `auto` `k` `m` `bn`' },
  direction: { types: Object.keys(DIRECTIONS) as ChartTypeName[], value: '`up` `down` `left` `right`' },
  valueAxis: { types: ['stacked', 'clustered', 'percent', 'waterfall', 'mekko', 'line', 'area', 'area100', 'combo', 'mekkoxy'], value: 'boolean' },
  zeroBased: { types: ['line', 'scatter', 'bubble'], value: 'boolean', auto: 'on, unless the data sits in a narrow band far from zero' },
  gap: { types: ['stacked', 'clustered', 'percent', 'waterfall', 'combo'], value: `number ${LIMITS.gap.min}–${LIMITS.gap.max}` },
  seriesLabels: { types: except('harvey', 'table'), value: '`right` `legend` `none`', auto: 'chosen per type and data' },
  totals: { types: ['stacked', 'percent', 'mekko', 'area', 'doughnut', 'combo', 'mekkoxy'], value: 'boolean' },
  cagr: { types: ['stacked', 'percent'], value: '`{from, to, years?}` or `null`', auto: 'none' },
  difference: { types: ['stacked', 'percent'], value: '`{from, to, mode?}` or `null`', auto: 'none' },
  totalsAt: { types: ['waterfall'], value: 'array of category indices', auto: 'the first category and the `"e"` cells, else first and last' },
  step: { types: ['line'], value: 'boolean' },
  percentData: { types: ['percent'], value: 'boolean' },
  combo: { types: ['combo'], value: '`stacked` `clustered`' },
  mekko: { types: ['mekko'], value: '`percent` `unit`' },
  sort: { types: ['mekkoxy'], value: '`sheet` `reverse` `ascending` `descending`' },
  axisTitles: { types: ['scatter', 'bubble'], value: '`{x?, y?}`', auto: 'none' },
  sizeLegend: { types: ['bubble'], value: 'boolean or caption' },
  labelContent: { types: ['pie', 'doughnut'], value: '`percent` `value` `both`' },
  scaleGroup: { types: ['stacked', 'clustered', 'waterfall', 'line', 'area', 'combo'], value: 'group name', auto: 'none (own scale)' },
  axisMax: { types: ['stacked', 'clustered', 'waterfall', 'line', 'area', 'combo'], value: 'number: the value axis reaches at least this high', auto: 'fits the data' },
  shades: { types: ['gantt'], value: 'array of `{from, to, label?}`', auto: 'none' },
  brackets: { types: ['gantt'], value: 'array of `{from, to, label?}`', auto: 'none' },
  today: { types: ['gantt'], value: 'date', auto: 'none' },
  levelArrows: { types: ['stacked', 'clustered', 'percent', 'line'], value: 'array of `{from, to, mode?, position?}`', auto: 'none' },
  valueLines: { types: ['stacked', 'clustered', 'line'], value: 'array of `{value, label?, position?}`', auto: 'none' },
  percentIndicator: { types: ['percent'], value: '`left` `right` `none`', auto: 'right; left when series names are on the right, none if a value axis is there too' },
  percentBase: { types: ['percent'], value: 'array of numbers or `null`, one per category', auto: 'each column\'s sum' },
  categoryGaps: { types: ['stacked', 'clustered', 'percent', 'waterfall', 'combo'], value: 'array of category indices', auto: 'none' },
  categoryBreaks: { types: ['stacked', 'clustered', 'percent', 'waterfall', 'combo'], value: 'array of category indices', auto: 'none' },
  axisBreak: { types: ['stacked', 'clustered'], value: '`{from, to, style?}` or `null`', auto: 'none' },
  secondaryAxis: { types: ['combo'], value: 'array of series indices', auto: 'none' },
} as const satisfies Record<keyof ChartOptions, OptionInfo>;
export type OptionKey = keyof typeof OPTIONS;
export const OPTION_KEYS = Object.keys(OPTIONS) as OptionKey[];

/** Whether a chart type reads an option. */
export const appliesTo = (key: string, type: string): boolean =>
  !!(OPTIONS as Record<string, OptionInfo>)[key]?.types.includes(type as ChartTypeName);

/**
 * The value the engine uses for each option a chart type reads when the spec leaves it out. Options
 * whose default depends on the data (`decimals`, `zeroBased`, `seriesLabels`, `totalsAt`) are absent;
 * `OPTIONS[k].auto` describes them. The engine lays out with these values and the editor shows them,
 * so a control always shows what is drawn.
 */
export function optionDefaults(type: string): ChartOptions {
  const all: ChartOptions = {
    scale: 'none',
    direction: 'up',
    valueAxis: type === 'line',
    gap: 0.5,
    totals: type === 'stacked' || type === 'mekko' || type === 'mekkoxy',
    step: false,
    percentData: false,
    combo: 'stacked',
    mekko: 'percent',
    sort: 'sheet',
    sizeLegend: false,
    labelContent: 'percent',
  };
  return Object.fromEntries(Object.entries(all).filter(([k]) => appliesTo(k, type))) as ChartOptions;
}

/** What the editor needs from this module; the server sends it with the page (`OC_SPEC`). */
export const editorSpec = () => ({
  options: Object.fromEntries(OPTION_KEYS.map((k) => [k, OPTIONS[k].types])),
  defaults: Object.fromEntries(CHART_TYPES.map((t) => [t, optionDefaults(t)])),
  limits: LIMITS,
  directions: DIRECTIONS,
});
export const DECK_KEYS = ['$schema', 'specVersion', 'style', 'agenda', 'slides'] as const;
export const SLIDE_KEYS = ['title', 'chart', 'charts', 'chapter', 'level', 'toc', 'split'] as const;
/** Keys of a deck's `agenda`: the agenda options plus the title of agenda slides. */
export const DECK_AGENDA_KEYS = [...AGENDA_OPTION_KEYS, 'title'] as const;
/** Keys of one entry of a slide's `charts` (several charts on one slide). */
export const SLIDE_CHART_KEYS = ['frame', 'chart'] as const;
export const STYLE_KEYS = ['font', 'themeFonts', 'fontSize','text', 'muted', 'axis', 'background', 'palette', 'positive', 'negative', 'theme', 'numbers', 'elements'] as const;
/** Keys of `style.numbers`: number defaults for every chart. */
export const STYLE_NUMBER_KEYS = ['locale', 'decimals', 'scale'] as const;

const clone = <T>(v: T): T => (v === undefined ? v : JSON.parse(JSON.stringify(v)));

/**
 * Upgrade any supported spec version to the current one. v0 (no `specVersion`) has the same shape
 * as v1; the upgrade stamps the version and removes options the chart type does not use (v0 accepted
 * them silently, v1 rejects them), adding one "/options/<key>: …" line per removal to `warnings`.
 * Specs newer than this engine are returned unchanged so `validateSpec` can report them.
 */
export function migrate<T extends object>(spec: T, warnings?: string[], base = ''): T & { specVersion: number } {
  if (!spec || typeof spec !== 'object' || Array.isArray(spec)) return spec as T & { specVersion: number };
  const v = (spec as { specVersion?: unknown }).specVersion;
  if (v !== undefined) return renameOrientation(spec, warnings, base) as T & { specVersion: number };
  const out = { specVersion: SPEC_VERSION, ...clone(spec) } as T & { specVersion: number; type?: unknown; options?: Record<string, unknown> };
  renameOrientation(out, warnings, base, true);
  const { type, options } = out;
  if (isObj(options) && CHART_TYPES.includes(type as never)) {
    for (const k of Object.keys(options)) {
      if (k in OPTIONS && !appliesTo(k, type as string)) {
        delete options[k];
        warnings?.push(`${base}/options/${esc(k)}: not used by ${type} charts; removed`);
      }
    }
  }
  return out;
}

/**
 * `options.orientation` (`vertical` | `horizontal`, before chart directions) is now `direction`
 * (`up` | `right`): rewritten wherever it appears, with a warning. Returns a copy when it changes
 * anything (in place with `inPlace`, for a copy `migrate` already made).
 */
function renameOrientation<T extends object>(spec: T, warnings: string[] | undefined, base: string, inPlace = false): T {
  const o = (spec as { options?: unknown }).options;
  if (!isObj(o) || !('orientation' in o)) return spec;
  const out = inPlace ? spec : clone(spec);
  const opts = (out as { options: Record<string, unknown> }).options;
  const { orientation, ...rest } = opts;
  const dir = orientation === 'horizontal' ? 'right' : orientation === 'vertical' ? 'up' : undefined;
  if (dir === undefined) return spec; // an invalid value: the validator reports it under its own name
  const next: Record<string, unknown> = {};
  // Same position in the options, so the spec still reads the same way.
  for (const k of Object.keys(opts)) {
    if (k === 'orientation') { if (dir !== 'up' && rest.direction === undefined) next.direction = dir; }
    else next[k] = opts[k];
  }
  (out as { options: Record<string, unknown> }).options = next;
  warnings?.push(`${base}/options/orientation: renamed to "direction"${dir === 'up' ? ' (up is the default; removed)' : ` ("${dir}")`}`);
  return out;
}

/** Upgrade a deck and every chart in it (warnings as in `migrate`, with paths from the deck root). */
export function migrateDeck<T extends { slides?: { chart?: object; charts?: { chart?: object }[] }[] }>(deck: T, warnings?: string[]): T & { specVersion: number } {
  const out = migrate(deck);
  if (Array.isArray(out.slides)) {
    out.slides = out.slides.map((s, i) => {
      if (!s || typeof s !== 'object') return s;
      if (Array.isArray(s.charts)) {
        return { ...s, charts: s.charts.map((c, k) => (c && typeof c === 'object' && c.chart ? { ...c, chart: migrate(c.chart, warnings, `/slides/${i}/charts/${k}/chart`) } : c)) };
      }
      return s.chart ? { ...s, chart: migrate(s.chart, warnings, `/slides/${i}/chart`) } : s;
    });
  }
  return out;
}

// ---- Validation --------------------------------------------------------------------------------

const isObj = (v: unknown): v is Record<string, unknown> => !!v && typeof v === 'object' && !Array.isArray(v);
const isInt = (v: unknown): v is number => typeof v === 'number' && Number.isInteger(v);
const HEX = /^#?[0-9A-Fa-f]{6}$/;
const esc = (k: string | number) => String(k).replace(/~/g, '~0').replace(/\//g, '~1');
const describe = (v: unknown) =>
  v === null ? 'null' : typeof v === 'number' ? String(v) : typeof v === 'string' ? `string "${v}"` : Array.isArray(v) ? 'array' : typeof v;

class Errors {
  list: SpecError[] = [];
  base: string;
  constructor(base: string) { this.base = base; }
  add(path: (string | number)[], message: string) {
    this.list.push({ path: this.base + path.map((p) => `/${esc(p)}`).join(''), message });
  }
  unknown(obj: Record<string, unknown>, known: readonly string[], path: (string | number)[]) {
    for (const k of Object.keys(obj)) if (!known.includes(k)) this.add([...path, k], `unknown property "${k}"`);
  }
}

function checkColor(e: Errors, v: unknown, path: (string | number)[]) {
  if (!isColor(v)) e.add(path, `must be a hex colour like "1F3A5F" or a theme colour (${COLOR_NAMES.join(', ')}), got ${describe(v)}`);
}

/** Every problem in a chart spec, path-addressed. Empty array = valid. Never throws. */
export function validateSpec(spec: unknown, base = ''): SpecError[] {
  const e = new Errors(base);
  if (!isObj(spec)) { e.add([], 'chart spec must be an object'); return e.list; }
  e.unknown(spec, CHART_KEYS, []);
  if (spec.specVersion !== undefined && spec.specVersion !== SPEC_VERSION) {
    e.add(['specVersion'], `unsupported specVersion ${describe(spec.specVersion)}; this engine reads version ${SPEC_VERSION} (made with a newer open-charts?)`);
  }
  if (spec.$schema !== undefined && typeof spec.$schema !== 'string') e.add(['$schema'], 'must be a string');
  const type = spec.type;
  if (!CHART_TYPES.includes(type as never)) e.add(['type'], `must be one of ${CHART_TYPES.join(', ')}, got ${describe(type)}`);

  let nCat = -1;
  if (!Array.isArray(spec.categories)) e.add(['categories'], 'must be an array of strings');
  else {
    nCat = spec.categories.length;
    if (!nCat) e.add(['categories'], 'chart needs at least one category');
    spec.categories.forEach((c, i) => { if (typeof c !== 'string') e.add(['categories', i], `must be a string, got ${describe(c)}`); });
  }

  if (!Array.isArray(spec.series)) e.add(['series'], 'must be an array of series');
  else {
    if (!spec.series.length) e.add(['series'], 'chart needs at least one series');
    const seen = new Map<string, number>();
    spec.series.forEach((s, i) => {
      if (!isObj(s)) { e.add(['series', i], 'must be an object with name and values'); return; }
      e.unknown(s, SERIES_KEYS, ['series', i]);
      if (typeof s.name !== 'string') e.add(['series', i, 'name'], `must be a string, got ${describe(s.name)}`);
      else if (seen.has(s.name)) e.add(['series', i, 'name'], `duplicate series name "${s.name}" (also series ${seen.get(s.name)})`);
      else seen.set(s.name, i);
      if (s.color !== undefined) checkColor(e, s.color, ['series', i, 'color']);
      if (!Array.isArray(s.values)) { e.add(['series', i, 'values'], 'must be an array'); return; }
      if (nCat >= 0 && s.values.length !== nCat) {
        e.add(['series', i, 'values'], `has ${s.values.length} values but there are ${nCat} categories`);
      }
      if (s.as !== undefined && s.as !== 'column' && s.as !== 'line') e.add(['series', i, 'as'], `must be "column" or "line", got ${describe(s.as)}`);
      if (s.hidden !== undefined && typeof s.hidden !== 'boolean') e.add(['series', i, 'hidden'], `must be true or false, got ${describe(s.hidden)}`);
      s.values.forEach((v, j) => {
        const at = ['series', i, 'values', j];
        if (v === null) return;
        if (v === 'e' && type === 'waterfall') return;
        if (type === 'scatter' || type === 'bubble') return checkPoint(e, v, at, type === 'bubble' ? 3 : 2);
        if (type === 'gantt') { if (!isDate(v)) e.add(at, `must be a date like "2026-03-15", "Mar 2026" or "Q2 2026", or null, got ${describe(v)}`); return; }
        if (type === 'table' && typeof v === 'string') return;
        if (typeof v !== 'number') e.add(at, `must be a number or null, got ${describe(v)}`);
        else if (!Number.isFinite(v)) e.add(at, `must be a finite number, got ${v}`);
        else if (type === 'mekkoxy' && i === 0 && v < 0) e.add(at, `width must not be negative, got ${v}`);
      });
    });
    checkTypeSeries(e, type as string, spec.series);
  }

  if (spec.options !== undefined) {
    if (!isObj(spec.options)) e.add(['options'], 'must be an object');
    else validateOptions(e, spec.options, type as string, nCat, Array.isArray(spec.series) ? spec.series : []);
  }
  if (spec.style !== undefined) checkStyle(e, spec.style);
  if (spec.overrides !== undefined) for (const p of checkOverrides(spec.overrides)) e.add(p.path, p.message);
  return e.list;
}

/** A style (on a deck or a chart): fonts, colours (hex or theme slots), theme colours, number defaults. */
function checkStyle(e: Errors, st: unknown) {
  if (!isObj(st)) { e.add(['style'], 'must be an object'); return; }
  e.unknown(st, STYLE_KEYS, ['style']);
  if (st.font !== undefined && (typeof st.font !== 'string' || !st.font)) e.add(['style', 'font'], 'must be a font name, or +mj-lt / +mn-lt');
  if (st.themeFonts !== undefined) {
    if (!isObj(st.themeFonts)) e.add(['style', 'themeFonts'], 'must be an object, e.g. {"major": "Aptos Display", "minor": "Aptos"}');
    else {
      e.unknown(st.themeFonts, ['major', 'minor'], ['style', 'themeFonts']);
      for (const k of ['major', 'minor']) {
        const v = st.themeFonts[k];
        if (v !== undefined && (typeof v !== 'string' || !v || v.startsWith('+'))) e.add(['style', 'themeFonts', k], 'must be a font name');
      }
    }
  }
  if (st.fontSize !== undefined && (typeof st.fontSize !== 'number' || !(st.fontSize >= 4 && st.fontSize <= 72))) e.add(['style', 'fontSize'], 'must be a number of points 4–72');
  for (const k of ['text', 'muted', 'axis', 'background', 'positive', 'negative']) if (st[k] !== undefined) checkColor(e, st[k], ['style', k]);
  if (st.palette !== undefined) {
    if (!Array.isArray(st.palette) || !st.palette.length) e.add(['style', 'palette'], 'must be a non-empty array of colours');
    else st.palette.forEach((c, i) => checkColor(e, c, ['style', 'palette', i]));
  }
  if (st.theme !== undefined) {
    if (!isObj(st.theme)) e.add(['style', 'theme'], 'must be an object of theme colours, e.g. {"accent1": "1F3A5F"}');
    else {
      e.unknown(st.theme, THEME_SLOT_KEYS, ['style', 'theme']);
      for (const k of THEME_SLOT_KEYS) {
        const v = st.theme[k];
        if (v !== undefined && (typeof v !== 'string' || !HEX.test(v))) e.add(['style', 'theme', k], `must be a hex colour like "1F3A5F", got ${describe(v)}`);
      }
    }
  }
  if (st.numbers !== undefined) {
    if (!isObj(st.numbers)) e.add(['style', 'numbers'], 'must be an object with locale, decimals and/or scale');
    else {
      const sub = new Errors(`${e.base}/style`);
      // The same rules as the chart options of the same names (paths "/style/numbers/<key>").
      validateOptions(sub, Object.fromEntries(Object.entries(st.numbers).filter(([k]) => (STYLE_NUMBER_KEYS as readonly string[]).includes(k))), '', -1, [], 'numbers');
      e.list.push(...sub.list);
      e.unknown(st.numbers, STYLE_NUMBER_KEYS, ['style', 'numbers']);
    }
  }
  if (st.elements !== undefined) {
    // Default element styles: the overrides vocabulary (role or selector → ElementStyle), minus what
    // only makes sense for one element of one chart (a pinned position, plain-text conversion).
    for (const p of checkOverrides(st.elements)) e.add(['style', 'elements', ...p.path.slice(1)], p.message);
    if (isObj(st.elements)) {
      for (const [k, v] of Object.entries(st.elements)) {
        for (const only of ['pin', 'detach'] as const) if (isObj(v) && v[only] !== undefined) e.add(['style', 'elements', k, only], `"${only}" is per element: set it in the chart's overrides, not in a style`);
      }
    }
  }
}

/** Every problem in a style object (deck, chart or style file), paths under `base`. Never throws. */
export function validateStyle(style: unknown, base = ''): SpecError[] {
  const e = new Errors(base);
  checkStyle(e, style);
  return e.list.map((x) => ({ ...x, path: x.path.replace(`${base}/style`, base) || '/' }));
}
const THEME_SLOT_KEYS = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'dk1', 'lt1', 'dk2', 'lt2'] as const;

function validateOptions(e: Errors, o: Record<string, unknown>, type: string, nCat: number, series: unknown[], at = 'options') {
  const p = (k: string, ...rest: (string | number)[]) => [at, k, ...rest];
  e.unknown(o, OPTION_KEYS, [at]);
  // One rule for every option: the chart type reads it, or it is an error (null = "none" is accepted).
  if (CHART_TYPES.includes(type as never)) {
    for (const k of OPTION_KEYS) {
      if (o[k] !== undefined && o[k] !== null && !appliesTo(k, type)) e.add(p(k), `not used by ${type} charts (only ${OPTIONS[k].types.join(', ')})`);
    }
  }
  const range = (k: keyof typeof LIMITS, whole: boolean) => {
    const v = o[k], { min, max } = LIMITS[k];
    if (v !== undefined && (typeof v !== 'number' || !Number.isFinite(v) || (whole && !isInt(v)) || v < min || v > max)) {
      e.add(p(k), `must be a ${whole ? 'whole ' : ''}number ${min}–${max}, got ${describe(v)}`);
    }
  };
  range('decimals', true);
  range('gap', false);
  for (const k of ['totals', 'valueAxis', 'zeroBased', 'step'] as const) {
    if (o[k] !== undefined && typeof o[k] !== 'boolean') e.add(p(k), `must be true or false, got ${describe(o[k])}`);
  }
  const index = (v: unknown, path: (string | number)[]) => {
    if (!isInt(v)) { e.add(path, `must be a category index, got ${describe(v)}`); return false; }
    if (nCat >= 0 && (v < 0 || v >= nCat)) { e.add(path, `category index ${v} is out of range 0–${nCat - 1}`); return false; }
    return true;
  };
  for (const k of ['cagr', 'difference'] as const) {
    const r = o[k];
    if (r === undefined || r === null) continue;
    if (!isObj(r)) { e.add(p(k), 'must be {from, to} or null'); continue; }
    e.unknown(r, k === 'cagr' ? ['from', 'to', 'years'] : ['from', 'to', 'mode'], p(k));
    if (k === 'difference' && r.mode !== undefined && !DIFF_MODES.includes(r.mode as never)) e.add(p(k, 'mode'), `must be one of ${DIFF_MODES.join(', ')}, got ${describe(r.mode)}`);
    const ok = index(r.from, p(k, 'from')) && index(r.to, p(k, 'to'));
    if (ok && r.from === r.to) e.add(p(k, 'to'), '"from" and "to" must be different categories');
    else if (ok && (r.from as number) > (r.to as number)) e.add(p(k, 'to'), '"to" must come after "from"');
    if (k === 'cagr' && r.years !== undefined && (typeof r.years !== 'number' || !Number.isFinite(r.years) || r.years <= 0)) {
      e.add(p(k, 'years'), `must be a positive number, got ${describe(r.years)}`);
    }
  }
  if (o.unit !== undefined && typeof o.unit !== 'string') {
    if (!isObj(o.unit)) e.add(p('unit'), `must be a string or {text, position, space}, got ${describe(o.unit)}`);
    else {
      e.unknown(o.unit, ['text', 'position', 'space'], p('unit'));
      if (typeof o.unit.text !== 'string') e.add(p('unit', 'text'), 'must be a string');
      if (o.unit.position !== undefined && o.unit.position !== 'prefix' && o.unit.position !== 'suffix') e.add(p('unit', 'position'), 'must be "prefix" or "suffix"');
      if (o.unit.space !== undefined && typeof o.unit.space !== 'boolean') e.add(p('unit', 'space'), 'must be true or false');
    }
  }
  if (o.format !== undefined && (typeof o.format !== 'string' || !/[0#?]/.test(o.format))) {
    e.add(p('format'), `must be an Excel number format with at least one digit placeholder (0 or #), got ${describe(o.format)}`);
  }
  if (o.locale !== undefined) {
    const lang = typeof o.locale === 'string' ? o.locale.split('-')[0].toLowerCase() : '';
    if (typeof o.locale !== 'string' || !Object.keys(LOCALES).some((k) => k === o.locale || k.toLowerCase().startsWith(`${lang}-`))) {
      e.add(p('locale'), `must be one of ${Object.keys(LOCALES).join(', ')}, got ${describe(o.locale)}`);
    }
  }
  const enumOf = (k: string, values: string[]) => {
    if (o[k] !== undefined && !values.includes(o[k] as string)) e.add(p(k), `must be one of ${values.join(', ')}, got ${describe(o[k])}`);
  };
  enumOf('scale', ['none', 'auto', 'k', 'm', 'bn']);
  if (o.direction !== undefined && appliesTo('direction', type)) {
    const ok = directionsOf(type);
    if (!ok.includes(o.direction as Direction)) e.add(p('direction'), `must be one of ${ok.join(', ')} for ${type} charts, got ${describe(o.direction)}`);
  }
  enumOf('mekko', ['percent', 'unit']);
  enumOf('seriesLabels', ['right', 'legend', 'none']);
  enumOf('labelContent', ['percent', 'value', 'both']);
  if (o.totalsAt !== undefined) {
    if (!Array.isArray(o.totalsAt)) e.add(p('totalsAt'), 'must be an array of category indices');
    else o.totalsAt.forEach((v, i) => index(v, p('totalsAt', i)));
  }
  if (o.percentData !== undefined && typeof o.percentData !== 'boolean') e.add(p('percentData'), `must be true or false, got ${describe(o.percentData)}`);
  enumOf('combo', ['stacked', 'clustered']);
  enumOf('sort', ['sheet', 'reverse', 'ascending', 'descending']);
  if (o.axisTitles !== undefined) {
    if (!isObj(o.axisTitles)) e.add(p('axisTitles'), 'must be {x, y}');
    else {
      e.unknown(o.axisTitles, ['x', 'y'], p('axisTitles'));
      for (const k of ['x', 'y']) if (o.axisTitles[k] !== undefined && typeof o.axisTitles[k] !== 'string') e.add(p('axisTitles', k), 'must be a string');
    }
  }
  for (const key of ['shades', 'brackets'] as const) {
    const list = o[key];
    if (list === undefined) continue;
    if (!Array.isArray(list)) e.add(p(key), 'must be an array of {from, to, label}');
    else list.forEach((sh, i) => {
      if (!isObj(sh)) { e.add(p(key, i), 'must be {from, to, label}'); return; }
      e.unknown(sh, ['from', 'to', 'label'], p(key, i));
      for (const k of ['from', 'to'] as const) if (!isDate(sh[k])) e.add(p(key, i, k), `must be a date like "2026-03-15", got ${describe(sh[k])}`);
      if (isDate(sh.from) && isDate(sh.to) && dayOf(sh.to) <= dayOf(sh.from)) e.add(p(key, i, 'to'), '"to" must come after "from"');
      if (sh.label !== undefined && typeof sh.label !== 'string') e.add(p(key, i, 'label'), 'must be a string');
    });
  }
  if (o.today !== undefined && !isDate(o.today)) e.add(p('today'), `must be a date like "2026-03-15", got ${describe(o.today)}`);
  validateDecorations(e, o, type, nCat, series, index);
  if (o.sizeLegend !== undefined && typeof o.sizeLegend !== 'boolean' && typeof o.sizeLegend !== 'string') e.add(p('sizeLegend'), `must be true, false or a caption, got ${describe(o.sizeLegend)}`);
  if (o.axisMax !== undefined && o.axisMax !== null && (typeof o.axisMax !== 'number' || !Number.isFinite(o.axisMax))) e.add(p('axisMax'), `must be a number, got ${describe(o.axisMax)}`);
  if (o.scaleGroup !== undefined && o.scaleGroup !== null && (typeof o.scaleGroup !== 'string' || !o.scaleGroup)) e.add(p('scaleGroup'), `must be a group name, got ${describe(o.scaleGroup)}`);
}

const DIFF_MODES = ['relative', 'absolute', 'reverse'] as const;

/** levelArrows, valueLines, percentIndicator, percentBase, categoryGaps / Breaks, axisBreak, secondaryAxis. */
function validateDecorations(e: Errors, o: Record<string, unknown>, type: string, nCat: number, series: unknown[], index: (v: unknown, path: (string | number)[]) => boolean) {
  const p = (k: string, ...rest: (string | number)[]) => ['options', k, ...rest];
  const nSer = series.length;
  const finite = (v: unknown): v is number => typeof v === 'number' && Number.isFinite(v);
  const nVl = Array.isArray(o.valueLines) ? o.valueLines.length : 0;
  if (o.valueLines !== undefined) {
    if (!Array.isArray(o.valueLines)) e.add(p('valueLines'), 'must be an array of {value, label, position}');
    else o.valueLines.forEach((vl, i) => {
      if (!isObj(vl)) { e.add(p('valueLines', i), 'must be {value, label, position}'); return; }
      e.unknown(vl, ['value', 'label', 'position'], p('valueLines', i));
      if (vl.value === 'mean') { if (type === 'line') e.add(p('valueLines', i, 'value'), 'line charts: give a number (the mean is of column totals)'); }
      else if (!finite(vl.value)) e.add(p('valueLines', i, 'value'), `must be a number or "mean", got ${describe(vl.value)}`);
      if (vl.label !== undefined && typeof vl.label !== 'string') e.add(p('valueLines', i, 'label'), 'must be a string');
      if (vl.position !== undefined && vl.position !== 'left' && vl.position !== 'right') e.add(p('valueLines', i, 'position'), `must be "left" or "right", got ${describe(vl.position)}`);
    });
  }
  if (o.levelArrows !== undefined) {
    if (!Array.isArray(o.levelArrows)) e.add(p('levelArrows'), 'must be an array of {from, to, mode, position}');
    else o.levelArrows.forEach((a, i) => {
      const at = p('levelArrows', i);
      if (!isObj(a)) { e.add(at, 'must be {from, to, mode, position}'); return; }
      e.unknown(a, ['from', 'to', 'mode', 'position'], at);
      const ends = (['from', 'to'] as const).map((k) => {
        const end = a[k];
        if (!isObj(end)) { e.add([...at, k], 'must be {category, series} or {valueLine}'); return null; }
        if (end.valueLine !== undefined) {
          e.unknown(end, ['valueLine'], [...at, k]);
          if (!isInt(end.valueLine) || end.valueLine < 0 || end.valueLine >= nVl) e.add([...at, k, 'valueLine'], nVl ? `must be a value line index 0–${nVl - 1}, got ${describe(end.valueLine)}` : 'the chart has no value lines (options.valueLines)');
          return 'vl';
        }
        e.unknown(end, ['category', 'series'], [...at, k]);
        index(end.category, [...at, k, 'category']);
        if (end.series === undefined) { if (type === 'clustered' || type === 'line') e.add([...at, k, 'series'], `${type} charts: say which series (a bar / point)`); }
        else if (!isInt(end.series) || end.series < 0 || end.series >= nSer) e.add([...at, k, 'series'], `must be a series index 0–${nSer - 1}, got ${describe(end.series)}`);
        return `${end.category}:${end.series}`;
      });
      if (ends[0] === 'vl' && ends[1] === 'vl') e.add(at, 'at least one end must be a segment or point');
      else if (ends[0] && ends[0] !== 'vl' && ends[0] === ends[1]) e.add([...at, 'to'], '"from" and "to" must be different');
      if (a.mode !== undefined && !DIFF_MODES.includes(a.mode as never)) e.add([...at, 'mode'], `must be one of ${DIFF_MODES.join(', ')}, got ${describe(a.mode)}`);
      if (a.position !== undefined && a.position !== 'beside' && a.position !== 'between') e.add([...at, 'position'], `must be "beside" or "between", got ${describe(a.position)}`);
    });
  }
  if (o.percentIndicator !== undefined && !['left', 'right', 'none'].includes(o.percentIndicator as string)) e.add(p('percentIndicator'), `must be one of left, right, none, got ${describe(o.percentIndicator)}`);
  if (o.percentBase !== undefined) {
    if (!Array.isArray(o.percentBase)) e.add(p('percentBase'), 'must be an array with one number (or null) per category');
    else {
      if (nCat >= 0 && o.percentBase.length !== nCat) e.add(p('percentBase'), `has ${o.percentBase.length} values but there are ${nCat} categories`);
      o.percentBase.forEach((v, i) => { if (v !== null && !(finite(v) && v > 0)) e.add(p('percentBase', i), `must be a positive number or null, got ${describe(v)}`); });
    }
  }
  for (const k of ['categoryGaps', 'categoryBreaks'] as const) {
    const list = o[k];
    if (list === undefined) continue;
    if (!Array.isArray(list)) { e.add(p(k), 'must be an array of category indices'); continue; }
    list.forEach((v, i) => {
      if (index(v, p(k, i)) && nCat >= 0 && v === nCat - 1) e.add(p(k, i), `goes after a category, so ${nCat - 1} (the last) has nothing to separate`);
    });
  }
  if (o.axisBreak !== undefined && o.axisBreak !== null) {
    const b = o.axisBreak;
    if (!isObj(b)) e.add(p('axisBreak'), 'must be {from, to, style} or null');
    else {
      e.unknown(b, ['from', 'to', 'style'], p('axisBreak'));
      for (const k of ['from', 'to']) if (!finite(b[k])) e.add(p('axisBreak', k), `must be a number, got ${describe(b[k])}`);
      if (finite(b.from) && finite(b.to) && b.to <= b.from) e.add(p('axisBreak', 'to'), '"to" must be greater than "from"');
      if (b.style !== undefined && b.style !== 'straight' && b.style !== 'wiggle') e.add(p('axisBreak', 'style'), `must be "straight" or "wiggle", got ${describe(b.style)}`);
    }
  }
  if (o.secondaryAxis !== undefined) {
    if (!Array.isArray(o.secondaryAxis)) e.add(p('secondaryAxis'), 'must be an array of series indices');
    else o.secondaryAxis.forEach((v, i) => {
      if (!isInt(v) || v < 0 || v >= nSer) e.add(p('secondaryAxis', i), `must be a series index 0–${nSer - 1}, got ${describe(v)}`);
      else if (!isObj(series[v]) || (series[v] as { as?: unknown }).as !== 'line') e.add(p('secondaryAxis', i), `series ${v} is drawn as columns; only a line ("as": "line") can use the secondary axis`);
    });
  }
}

const isDate = (v: unknown): v is string => typeof v === 'string' && parseCategoryDate(v) !== null;
const dayOf = (v: string) => { const d = parseCategoryDate(v)!; return Date.UTC(d.y, d.m - 1, d.d); };

/** Scatter / bubble cell: [x, y] or [x, y, size] of finite numbers, size not negative. */
function checkPoint(e: Errors, v: unknown, at: (string | number)[], n: number) {
  const what = n === 3 ? '[x, y, size]' : '[x, y]';
  if (!Array.isArray(v) || v.length !== n) { e.add(at, `must be ${what} or null, got ${describe(v)}`); return; }
  v.forEach((x, k) => { if (typeof x !== 'number' || !Number.isFinite(x)) e.add([...at, k], `must be a finite number, got ${describe(x)}`); });
  if (n === 3 && typeof v[2] === 'number' && v[2] < 0) e.add([...at, 2], `bubble size must not be negative, got ${v[2]}`);
}

/** Rules on the series as a whole that depend on the chart type. */
function checkTypeSeries(e: Errors, type: string, series: unknown[]) {
  // Hidden rows: something must stay visible, and rows a chart type reads by position cannot hide.
  const hidden = (s: unknown) => isObj(s) && s.hidden === true;
  if (series.length && series.every(hidden)) e.add(['series'], 'every series is hidden; show at least one');
  const fixed = type === 'gantt' ? ['Start', 'End'] : type === 'mekkoxy' ? ['width'] : [];
  fixed.forEach((what, i) => { if (hidden(series[i])) e.add(['series', i, 'hidden'], `the ${what} series of a ${type === 'gantt' ? 'Gantt' : 'Mekko XY'} chart cannot be hidden`); });
  if (type === 'mekkoxy' && series.length >= 2 && !hidden(series[0]) && series.slice(1).every(hidden)) e.add(['series'], 'Mekko XY needs at least one visible height series');
  if (type === 'mekkoxy' && series.length < 2) e.add(['series'], 'Mekko XY needs a width series (the first) and at least one height series');
  if (type === 'gantt') {
    if (series.length < 2) { e.add(['series'], 'Gantt needs a Start series (the first) and an End series (the second); further series are milestones'); return; }
    const [a, b] = series as { values?: unknown[] }[];
    if (!Array.isArray(a?.values) || !Array.isArray(b?.values)) return;
    a.values.forEach((s, j) => {
      const f = b.values![j];
      if (isDate(s) !== isDate(f)) e.add(['series', isDate(s) ? 1 : 0, 'values', j], 'a bar needs both a start and an end date');
      else if (isDate(s) && isDate(f) && dayOf(f) <= dayOf(s)) e.add(['series', 1, 'values', j], 'end must come after start');
    });
  }
}

/** Every problem in a deck: its own keys, the style, and each slide's chart (paths start at the deck root). */
export function validateDeck(deck: unknown): SpecError[] {
  const e = new Errors('');
  if (!isObj(deck)) { e.add([], 'deck must be an object with a "slides" array'); return e.list; }
  e.unknown(deck, DECK_KEYS, []);
  if (deck.specVersion !== undefined && deck.specVersion !== SPEC_VERSION) {
    e.add(['specVersion'], `unsupported specVersion ${describe(deck.specVersion)}; this engine reads version ${SPEC_VERSION}`);
  }
  if (deck.style !== undefined) checkStyle(e, deck.style);
  if (deck.agenda !== undefined) validateAgenda(e, deck.agenda);
  if (!Array.isArray(deck.slides)) { e.add(['slides'], 'must be an array of slides'); return e.list; }
  const out = e.list;
  deck.slides.forEach((s, i) => {
    if (!isObj(s)) { out.push({ path: `/slides/${i}`, message: 'must be an object with a "chart"' }); return; }
    const se = new Errors(`/slides/${i}`);
    se.unknown(s, SLIDE_KEYS, []);
    if (s.title !== undefined && typeof s.title !== 'string') se.add(['title'], 'must be a string');
    if (s.split !== undefined && typeof s.split !== 'boolean') se.add(['split'], 'must be true or false');
    if (s.chapter !== undefined || s.toc !== undefined || s.level !== undefined) {
      // An agenda slide: the agenda draws it, so it has no chart.
      if (s.chapter !== undefined && (typeof s.chapter !== 'string' || !s.chapter.trim())) se.add(['chapter'], 'must be the chapter\'s name');
      if (s.toc !== undefined && s.toc !== true) se.add(['toc'], 'must be true (a table-of-contents slide)');
      if (s.chapter !== undefined && s.toc !== undefined) se.add(['toc'], 'a slide is either a chapter or a table of contents, not both');
      if (s.level !== undefined && (s.chapter === undefined || !Number.isInteger(s.level) || (s.level as number) < 1 || (s.level as number) > 9)) se.add(['level'], 'must be a whole number 1–9 on a chapter slide');
      if (s.chapter === undefined && s.toc === undefined) se.add([], 'a slide with "level" needs a "chapter"');
      if (s.chart !== undefined || s.charts !== undefined) se.add([s.chart !== undefined ? 'chart' : 'charts'], 'an agenda slide (chapter or toc) has no chart');
      out.push(...se.list);
      return;
    }
    if (s.split && s.chart === undefined && s.charts === undefined) { out.push(...se.list); return; }
    if (s.charts === undefined) { out.push(...se.list, ...validateSpec(migrate(s.chart as object), `/slides/${i}/chart`)); return; }
    // Several charts on one slide: `charts` instead of `chart`, each with an optional frame.
    if (s.chart !== undefined) se.add(['chart'], 'a slide has either "chart" or "charts", not both');
    if (!Array.isArray(s.charts) || !s.charts.length) { se.add(['charts'], 'must be a non-empty array of {frame, chart}'); out.push(...se.list); return; }
    out.push(...se.list);
    s.charts.forEach((c, k) => {
      const ce = new Errors(`/slides/${i}/charts/${k}`);
      if (!isObj(c)) { ce.add([], 'must be an object with a "chart"'); out.push(...ce.list); return; }
      ce.unknown(c, SLIDE_CHART_KEYS, []);
      if (c.frame !== undefined) {
        const f = c.frame;
        if (!isObj(f)) ce.add(['frame'], 'must be {x, y, w, h} in points');
        else {
          ce.unknown(f, ['x', 'y', 'w', 'h'], ['frame']);
          for (const key of ['x', 'y', 'w', 'h']) {
            const v = f[key];
            if (typeof v !== 'number' || !Number.isFinite(v) || ((key === 'w' || key === 'h') && v <= 0)) ce.add(['frame', key], `must be a ${key === 'w' || key === 'h' ? 'positive ' : ''}number of points, got ${describe(v)}`);
          }
        }
      }
      out.push(...ce.list, ...validateSpec(migrate(c.chart as object), `/slides/${i}/charts/${k}/chart`));
    });
  });
  return out;
}

function validateAgenda(e: Errors, a: unknown) {
  if (!isObj(a)) { e.add(['agenda'], 'must be an object of agenda options'); return; }
  e.unknown(a, DECK_AGENDA_KEYS, ['agenda']);
  if (a.title !== undefined && typeof a.title !== 'string') e.add(['agenda', 'title'], 'must be a string');
  if (a.preset !== undefined && !(AGENDA_PRESETS as readonly unknown[]).includes(a.preset)) e.add(['agenda', 'preset'], `must be one of ${AGENDA_PRESETS.join(', ')}`);
  if (a.numbering !== undefined && typeof a.numbering !== 'boolean') e.add(['agenda', 'numbering'], 'must be true or false');
  if (a.fontSize !== undefined && (typeof a.fontSize !== 'number' || !(a.fontSize >= 6 && a.fontSize <= 72))) e.add(['agenda', 'fontSize'], 'must be a number of points 6–72');
  if (a.tocLevels !== undefined && (!Number.isInteger(a.tocLevels) || (a.tocLevels as number) < 1 || (a.tocLevels as number) > 9)) e.add(['agenda', 'tocLevels'], 'must be a whole number 1–9');
  if (a.styles === undefined) return;
  if (!isObj(a.styles)) { e.add(['agenda', 'styles'], 'must be an object of heading styles by state'); return; }
  e.unknown(a.styles, AGENDA_STATES, ['agenda', 'styles']);
  for (const [k, v] of Object.entries(a.styles)) {
    const p = ['agenda', 'styles', k];
    if (!isObj(v)) { e.add(p, 'must be an object with fill, line, color, bold and/or size'); continue; }
    e.unknown(v, HEADING_STYLE_KEYS, p);
    for (const c of ['fill', 'line', 'color']) if (v[c] !== undefined && v[c] !== '') checkColor(e, v[c], [...p, c]);
    if (v.bold !== undefined && typeof v.bold !== 'boolean') e.add([...p, 'bold'], 'must be true or false');
    if (v.size !== undefined && (typeof v.size !== 'number' || !(v.size >= 4 && v.size <= 96))) e.add([...p, 'size'], 'must be a number of points 4–96');
  }
}

/** Error thrown by layoutChart for an invalid spec; `errors` holds the path-addressed list. */
export class InvalidSpecError extends Error {
  errors: SpecError[];
  constructor(errors: SpecError[]) {
    super(`Invalid chart spec:\n${errors.map((x) => `  ${x.path || '/'}: ${x.message}`).join('\n')}`);
    this.name = 'InvalidSpecError';
    this.errors = errors;
  }
}

/** Migrate, validate and return the current-version spec, or throw InvalidSpecError. */
export function prepareSpec(spec: ChartSpec, warnings?: string[]): ChartSpec {
  const s = migrate(spec, warnings);
  const errors = validateSpec(s);
  if (errors.length) throw new InvalidSpecError(errors);
  return s;
}
