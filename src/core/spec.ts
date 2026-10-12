// Spec versioning and validation. Pure functions; the JSON Schema in schema/*.schema.json mirrors
// these rules (a sync test keeps the key lists identical).
import type { ChartSpec } from './types.ts';
import { LOCALES } from './numbers.ts';

export const SPEC_VERSION = 1;

export interface SpecError {
  /** JSON Pointer into the spec, e.g. "/series/1/values/3" ("" = the root). */
  path: string;
  message: string;
}

export const CHART_TYPES = [
  'stacked', 'clustered', 'percent', 'waterfall', 'mekko', 'line',
  'area', 'area100',
] as const;
export const CHART_KEYS = ['$schema', 'specVersion', 'type', 'categories', 'series', 'options'] as const;
export const SERIES_KEYS = ['name', 'values', 'color'] as const;
export const OPTION_KEYS = [
  'totals', 'cagr', 'difference', 'valueAxis', 'zeroBased', 'decimals', 'unit', 'format', 'locale',
  'scale', 'orientation', 'mekko', 'gap', 'seriesLabels', 'totalsAt',
  'step',
] as const;
export const DECK_KEYS = ['$schema', 'specVersion', 'style', 'slides'] as const;
export const SLIDE_KEYS = ['title', 'chart'] as const;
export const STYLE_KEYS = ['font', 'fontSize', 'text', 'muted', 'axis', 'background', 'palette', 'positive', 'negative'] as const;

const clone = <T>(v: T): T => (v === undefined ? v : JSON.parse(JSON.stringify(v)));

/**
 * Upgrade any supported spec version to the current one. v0 (no `specVersion`) has the same shape
 * as v1, so the upgrade only stamps the version: nothing is dropped or renamed. Specs newer than this
 * engine are returned unchanged so `validateSpec` can report them.
 */
export function migrate<T extends object>(spec: T): T & { specVersion: number } {
  if (!spec || typeof spec !== 'object' || Array.isArray(spec)) return spec as T & { specVersion: number };
  const v = (spec as { specVersion?: unknown }).specVersion;
  if (v === undefined) return { specVersion: SPEC_VERSION, ...clone(spec) } as T & { specVersion: number };
  return spec as T & { specVersion: number };
}

/** Upgrade a deck and every chart in it. */
export function migrateDeck<T extends { slides?: { chart?: object }[] }>(deck: T): T & { specVersion: number } {
  const out = migrate(deck);
  if (Array.isArray(out.slides)) {
    out.slides = out.slides.map((s) => (s && typeof s === 'object' && s.chart ? { ...s, chart: migrate(s.chart) } : s));
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
  if (typeof v !== 'string' || !HEX.test(v)) e.add(path, `must be a hex colour like "1F3A5F" or "#1F3A5F", got ${describe(v)}`);
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
      s.values.forEach((v, j) => {
        if (v === null) return;
        if (v === 'e' && type === 'waterfall') return;
        if (typeof v !== 'number') e.add(['series', i, 'values', j], `must be a number or null, got ${describe(v)}`);
        else if (!Number.isFinite(v)) e.add(['series', i, 'values', j], `must be a finite number, got ${v}`);
      });
    });
  }

  if (spec.options !== undefined) {
    if (!isObj(spec.options)) e.add(['options'], 'must be an object');
    else validateOptions(e, spec.options, type as string, nCat);
  }
  return e.list;
}

function validateOptions(e: Errors, o: Record<string, unknown>, type: string, nCat: number) {
  const p = (k: string, ...rest: (string | number)[]) => ['options', k, ...rest];
  e.unknown(o, OPTION_KEYS, ['options']);
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
    e.unknown(r, k === 'cagr' ? ['from', 'to', 'years'] : ['from', 'to'], p(k));
    const ok = index(r.from, p(k, 'from')) && index(r.to, p(k, 'to'));
    if (ok && r.from === r.to) e.add(p(k, 'to'), '"from" and "to" must be different categories');
    else if (ok && (r.from as number) > (r.to as number)) e.add(p(k, 'to'), '"to" must come after "from"');
    if (k === 'cagr' && r.years !== undefined && (typeof r.years !== 'number' || !Number.isFinite(r.years) || r.years <= 0)) {
      e.add(p(k, 'years'), `must be a positive number, got ${describe(r.years)}`);
    }
  }
  if (o.decimals !== undefined && (!isInt(o.decimals) || o.decimals < 0 || o.decimals > 10)) {
    e.add(p('decimals'), `must be a whole number 0–10, got ${describe(o.decimals)}`);
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
  enumOf('orientation', ['vertical', 'horizontal']);
  enumOf('mekko', ['percent', 'unit']);
  enumOf('seriesLabels', ['right', 'legend', 'none']);
  if (o.orientation === 'horizontal' && !['stacked', 'clustered', 'percent'].includes(type)) {
    e.add(p('orientation'), `horizontal orientation is available for stacked, clustered and percent charts, not ${type}`);
  }
  if (o.step && type !== 'line') e.add(p('step'), `step lines are available for line charts, not ${type}`);
  if (o.gap !== undefined && (typeof o.gap !== 'number' || !Number.isFinite(o.gap) || o.gap < 0 || o.gap > 10)) {
    e.add(p('gap'), `must be a number 0–10, got ${describe(o.gap)}`);
  }
  if (o.totalsAt !== undefined) {
    if (!Array.isArray(o.totalsAt)) e.add(p('totalsAt'), 'must be an array of category indices');
    else o.totalsAt.forEach((v, i) => index(v, p('totalsAt', i)));
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
  if (deck.style !== undefined) {
    if (!isObj(deck.style)) e.add(['style'], 'must be an object');
    else {
      const st = deck.style;
      e.unknown(st, STYLE_KEYS, ['style']);
      if (st.font !== undefined && (typeof st.font !== 'string' || !st.font)) e.add(['style', 'font'], 'must be a font name');
      if (st.fontSize !== undefined && (typeof st.fontSize !== 'number' || !(st.fontSize >= 4 && st.fontSize <= 72))) e.add(['style', 'fontSize'], 'must be a number of points 4–72');
      for (const k of ['text', 'muted', 'axis', 'background', 'positive', 'negative']) if (st[k] !== undefined) checkColor(e, st[k], ['style', k]);
      if (st.palette !== undefined) {
        if (!Array.isArray(st.palette) || !st.palette.length) e.add(['style', 'palette'], 'must be a non-empty array of hex colours');
        else st.palette.forEach((c, i) => checkColor(e, c, ['style', 'palette', i]));
      }
    }
  }
  if (!Array.isArray(deck.slides)) { e.add(['slides'], 'must be an array of slides'); return e.list; }
  const out = e.list;
  deck.slides.forEach((s, i) => {
    if (!isObj(s)) { out.push({ path: `/slides/${i}`, message: 'must be an object with a "chart"' }); return; }
    const se = new Errors(`/slides/${i}`);
    se.unknown(s, SLIDE_KEYS, []);
    if (s.title !== undefined && typeof s.title !== 'string') se.add(['title'], 'must be a string');
    out.push(...se.list, ...validateSpec(migrate(s.chart as object), `/slides/${i}/chart`));
  });
  return out;
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
export function prepareSpec(spec: ChartSpec): ChartSpec {
  const s = migrate(spec);
  const errors = validateSpec(s);
  if (errors.length) throw new InvalidSpecError(errors);
  return s;
}
