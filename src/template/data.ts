// The fill data format (`open-charts fill template.pptx data.json`): values by slot name, and
// optionally which template slides to output, in which order, repeated per item.
//
//   {
//     "values": { "client": "Acme", "revenue": { "type": "stacked", ... } },
//     "slides": [ { "slide": 1 }, { "slide": 3, "each": [ { "region": "EMEA", "sales": {...} }, ... ] } ]
//   }
//
// A value is text (string or number) for a `{{name}}` field, or a chart spec for a `chart:name` shape.
import type { ChartSpec, Style } from '../core/types.ts';
import { migrate, validateSpec, validateStyle, type SpecError } from '../core/spec.ts';

export type FillValue = string | number | ChartSpec;
export type FillValues = Record<string, FillValue>;

export interface FillSlide {
  /** Template slide number, 1-based. */
  slide: number;
  /** Values for this slide, over the data's `values`. */
  values?: FillValues;
  /** One copy of the slide per item, each item's values over the slide's. */
  each?: FillValues[];
}

export interface FillData {
  $schema?: string;
  /** Style overrides for every chart (over a `--style` file, under each chart's own `style`). */
  style?: Partial<Style>;
  values?: FillValues;
  /** Output slides in order. Omitted: every template slide once, in template order. */
  slides?: FillSlide[];
}

export const FILL_KEYS = ['$schema', 'style', 'values', 'slides'] as const;
export const FILL_SLIDE_KEYS = ['slide', 'values', 'each'] as const;
/** Slot names: letters, digits, `_`, `-`, `.`; `{{name}}` in text, `chart:name` as a shape name. */
export const SLOT_NAME = /^[A-Za-z0-9_.-]{1,100}$/;
export const FILL_LIMITS = { slides: 2000, each: 1000, text: 10_000 } as const;

const isObj = (v: unknown): v is Record<string, unknown> => typeof v === 'object' && v !== null && !Array.isArray(v);

export const isChartValue = (v: unknown): v is ChartSpec => isObj(v);

/** Every problem in fill data, path-addressed. Never throws. */
export function validateFillData(d: unknown): SpecError[] {
  const errors: SpecError[] = [];
  const add = (path: string, message: string) => errors.push({ path, message });
  if (!isObj(d)) { add('/', 'fill data must be a JSON object with "values" and optional "slides"'); return errors; }
  for (const k of Object.keys(d)) if (!(FILL_KEYS as readonly string[]).includes(k)) add(`/${k}`, `unknown property "${k}" (expected ${FILL_KEYS.join(', ')})`);
  if (d.style !== undefined) errors.push(...validateStyle(d.style, '/style'));
  if (d.values !== undefined) checkValues(d.values, '/values', errors);
  if (d.slides !== undefined) {
    if (!Array.isArray(d.slides)) add('/slides', 'must be an array of {slide, values?, each?}');
    else {
      if (d.slides.length > FILL_LIMITS.slides) add('/slides', `at most ${FILL_LIMITS.slides} slides`);
      d.slides.forEach((s, i) => {
        const at = `/slides/${i}`;
        if (!isObj(s)) { add(at, 'must be an object {slide, values?, each?}'); return; }
        for (const k of Object.keys(s)) if (!(FILL_SLIDE_KEYS as readonly string[]).includes(k)) add(`${at}/${k}`, `unknown property "${k}"`);
        if (!Number.isInteger(s.slide) || (s.slide as number) < 1) add(`${at}/slide`, 'must be a template slide number (1 = the first slide)');
        if (s.values !== undefined) checkValues(s.values, `${at}/values`, errors);
        if (s.each !== undefined) {
          if (!Array.isArray(s.each) || !s.each.length) add(`${at}/each`, 'must be a non-empty array of values, one per copy of the slide');
          else if (s.each.length > FILL_LIMITS.each) add(`${at}/each`, `at most ${FILL_LIMITS.each} items`);
          else s.each.forEach((v, j) => checkValues(v, `${at}/each/${j}`, errors));
        }
      });
    }
  }
  return errors;
}

export function checkValues(v: unknown, at: string, errors: SpecError[]) {
  if (!isObj(v)) { errors.push({ path: at, message: 'must be an object of values by slot name' }); return; }
  for (const [k, x] of Object.entries(v)) {
    const p = `${at}/${k}`;
    if (!SLOT_NAME.test(k)) errors.push({ path: p, message: 'slot names use letters, digits, "_", "-" and "." (at most 100)' });
    if (typeof x === 'string') { if (x.length > FILL_LIMITS.text) errors.push({ path: p, message: `text longer than ${FILL_LIMITS.text} characters` }); }
    else if (typeof x === 'number') { if (!Number.isFinite(x)) errors.push({ path: p, message: 'must be a finite number' }); }
    else if (isObj(x)) errors.push(...validateSpec(migrate(x), p));
    else errors.push({ path: p, message: 'must be text, a number or a chart spec' });
  }
}
