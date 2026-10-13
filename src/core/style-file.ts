// Style files: a client deck's look as one JSON file (`--style style.json` on build and fill). A style
// file is a style (the same keys as `style` in a deck) plus a name and per-chart-type styles under
// `types`: fonts, colours, number defaults and default element styles (`elements`, the overrides
// vocabulary). Style files are treated as untrusted input: size-limited, every key allowed by name,
// every value checked, before anything reads them.
import type { ChartSpec, Style } from './types.ts';
import { CHART_TYPES, validateStyle, type SpecError } from './spec.ts';

export interface StyleFile extends Partial<Style> {
  $schema?: string;
  /** Shown in messages and stored nowhere else. */
  name?: string;
  description?: string;
  /** Overrides for one chart type over the file's own keys, e.g. `{"line": {"palette": [...]}}`. */
  types?: Partial<Record<ChartSpec['type'], Partial<Style>>>;
}

export const STYLE_FILE_KEYS = ['$schema', 'name', 'description', 'types'] as const;
/** Hard limits on an untrusted style file. */
export const STYLE_FILE_LIMITS = { bytes: 256 * 1024, string: 200, font: 64, palette: 64 } as const;

const isObj = (v: unknown): v is Record<string, unknown> => typeof v === 'object' && v !== null && !Array.isArray(v);

export class InvalidStyleFileError extends Error {
  errors: SpecError[];
  constructor(errors: SpecError[], name = 'style file') {
    super(`Invalid ${name}:\n${errors.map((e) => `  ${e.path}: ${e.message}`).join('\n')}`);
    this.errors = errors;
  }
}

/** Every problem in a parsed style file, path-addressed. Empty = valid. Never throws. */
export function validateStyleFile(v: unknown): SpecError[] {
  if (!isObj(v)) return [{ path: '/', message: 'a style file must be a JSON object' }];
  const errors: SpecError[] = [];
  const own = Object.fromEntries(Object.entries(v).filter(([k]) => !(STYLE_FILE_KEYS as readonly string[]).includes(k)));
  errors.push(...validateStyle(own));
  for (const k of ['$schema', 'name', 'description'] as const) {
    const s = v[k];
    if (s !== undefined && (typeof s !== 'string' || s.length > (k === 'description' ? 2000 : STYLE_FILE_LIMITS.string))) errors.push({ path: `/${k}`, message: `must be a string of at most ${k === 'description' ? 2000 : STYLE_FILE_LIMITS.string} characters` });
  }
  const limits = (st: Record<string, unknown>, base: string) => {
    if (typeof st.font === 'string' && st.font.length > STYLE_FILE_LIMITS.font) errors.push({ path: `${base}/font`, message: `must be at most ${STYLE_FILE_LIMITS.font} characters` });
    if (typeof st.font === 'string' && /[\u0000-\u001f<>"&]/.test(st.font)) errors.push({ path: `${base}/font`, message: 'must be a plain font name' });
    if (Array.isArray(st.palette) && st.palette.length > STYLE_FILE_LIMITS.palette) errors.push({ path: `${base}/palette`, message: `must have at most ${STYLE_FILE_LIMITS.palette} colours` });
  };
  limits(own, '');
  if (v.types !== undefined) {
    if (!isObj(v.types)) errors.push({ path: '/types', message: 'must be an object of styles by chart type, e.g. {"line": {"palette": ["accent2"]}}' });
    else {
      for (const [t, st] of Object.entries(v.types)) {
        if (!(CHART_TYPES as readonly string[]).includes(t)) { errors.push({ path: `/types/${t}`, message: `unknown chart type "${t}" (one of ${CHART_TYPES.join(', ')})` }); continue; }
        errors.push(...validateStyle(st, `/types/${t}`));
        if (isObj(st)) limits(st, `/types/${t}`);
      }
    }
  }
  return errors;
}

/** Parses and checks an untrusted style file; throws InvalidStyleFileError listing every problem. */
export function parseStyleFile(text: string, name = 'style file'): StyleFile {
  if (text.length > STYLE_FILE_LIMITS.bytes) throw new InvalidStyleFileError([{ path: '/', message: `larger than ${STYLE_FILE_LIMITS.bytes / 1024} KB` }], name);
  let v: unknown;
  try { v = JSON.parse(text); } catch (err) { throw new InvalidStyleFileError([{ path: '/', message: `not JSON: ${(err as Error).message}` }], name); }
  const errors = validateStyleFile(v);
  if (errors.length) throw new InvalidStyleFileError(errors, name);
  return v as StyleFile;
}

const NESTED = ['theme', 'numbers'] as const;

/**
 * Layers styles, later wins. `theme` and `numbers` merge one level deeper, and `elements` per key
 * (each key's properties merge), so an override can change one colour or one line's width.
 */
export function mergeStyles(...styles: (Partial<Style> | undefined)[]): Partial<Style> {
  const out: Partial<Style> = {};
  for (const st of styles) {
    if (!st) continue;
    for (const [k, v] of Object.entries(st) as [keyof Style, never][]) {
      if (v === undefined) continue;
      if ((NESTED as readonly string[]).includes(k)) (out as Record<string, unknown>)[k] = { ...(out[k] as object), ...(v as object) };
      else if (k === 'elements') {
        const el: Record<string, object> = { ...out.elements };
        for (const [ek, es] of Object.entries(v as object)) el[ek] = { ...el[ek], ...es };
        out.elements = el;
      } else (out as Record<string, unknown>)[k] = v;
    }
  }
  return out;
}

/** A style file's own style keys (no name, no per-type overrides). */
export function fileStyle(file: StyleFile | undefined): Partial<Style> {
  if (!file) return {};
  const { $schema: _s, name: _n, description: _d, types: _t, ...own } = file;
  return own;
}

/** The style a style file gives one chart type: its own keys, then `types[type]` over them. */
export function styleFor(file: StyleFile | undefined, type: ChartSpec['type']): Partial<Style> {
  return mergeStyles(fileStyle(file), file?.types?.[type]);
}
