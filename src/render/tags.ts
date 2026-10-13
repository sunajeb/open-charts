// Tag contract v1 (docs/MASTER-PLAN.md §3): the keys every host writes on an open-charts group so
// the chart can be found, re-opened and re-laid out later. Written as real PowerPoint tags
// (p:custDataLst → ppt/tags/tagN.xml), which VBA sees as Shape.Tags and Office.js as shape.tags.
import type { Frame } from '../core/types.ts';
import { ENGINE_VERSION } from '../version.ts';

export const TAG_CONTRACT_VERSION = '1';
/** OC_SPEC values longer than this (UTF-16 code units) are split into OC_SPEC_0..n + OC_SPEC_N. */
export const SPEC_CHUNK = 30_000;

export type Tags = [name: string, value: string][];

const num = (v: number) => String(+v.toFixed(2));

/** Group-level tags for one chart. `spec` is the ChartSpec plus optional `style` overrides. */
export function chartTags(spec: object, frame: Frame, font: string, chunk = SPEC_CHUNK): Tags {
  const json = JSON.stringify(spec);
  const tags: Tags = [['OC_V', TAG_CONTRACT_VERSION], ['OC_ENGINE', ENGINE_VERSION]];
  if (json.length <= chunk) {
    tags.push(['OC_SPEC', json]);
  } else {
    const parts = splitChunks(json, chunk);
    parts.forEach((p, i) => tags.push([`OC_SPEC_${i}`, p]));
    tags.push(['OC_SPEC_N', String(parts.length)]);
  }
  tags.push(['OC_FRAME', [frame.x, frame.y, frame.w, frame.h].map(num).join(',')], ['OC_FONT', font]);
  return tags;
}

/**
 * A long tag value as tags: `name` when it fits, else `name_0`..`name_n` plus `name_N` (the count),
 * like OC_SPEC. Used for OC_BASE (last-rendered formatting per role, core/overrides.ts).
 */
export function chunkedTags(name: string, value: string, chunk = SPEC_CHUNK): Tags {
  if (value.length <= chunk) return [[name, value]];
  const parts = splitChunks(value, chunk);
  return [...parts.map((p, i): [string, string] => [`${name}_${i}`, p]), [`${name}_N`, String(parts.length)]];
}

/** Reassembles a chunked tag (see chunkedTags); undefined when absent. Names match case-insensitively. */
export function chunkedValue(tags: Record<string, string>, name: string): string | undefined {
  const t = upper(tags), key = name.toUpperCase();
  if (t[key] !== undefined) return t[key];
  if (t[`${key}_N`] === undefined) return undefined;
  const n = Number(t[`${key}_N`]);
  const parts: string[] = [];
  for (let i = 0; i < n; i++) {
    const p = t[`${key}_${i}`];
    if (p === undefined) throw new Error(`${name} chunk ${i} of ${n} is missing`);
    parts.push(p);
  }
  return parts.join('');
}

function splitChunks(s: string, size: number): string[] {
  const out: string[] = [];
  for (let i = 0; i < s.length;) {
    let j = Math.min(i + size, s.length);
    const c = s.charCodeAt(j - 1);
    if (j < s.length && c >= 0xd800 && c <= 0xdbff) j--; // never split a surrogate pair
    out.push(s.slice(i, j));
    i = j;
  }
  return out;
}

/** Reassembles OC_SPEC (plain or chunked). Tag names are matched case-insensitively, as PowerPoint upper-cases them. */
export function specJson(tags: Record<string, string>): string | undefined {
  const t = upper(tags);
  if (t.OC_SPEC !== undefined) return t.OC_SPEC;
  if (t.OC_SPEC_N === undefined) return undefined;
  const n = Number(t.OC_SPEC_N);
  const parts: string[] = [];
  for (let i = 0; i < n; i++) {
    const p = t[`OC_SPEC_${i}`];
    if (p === undefined) throw new Error(`OC_SPEC chunk ${i} of ${n} is missing`);
    parts.push(p);
  }
  return parts.join('');
}

export function parseFrame(v: string | undefined): Frame | undefined {
  const n = v?.split(',').map(Number);
  return n && n.length === 4 && n.every(Number.isFinite) ? { x: n[0], y: n[1], w: n[2], h: n[3] } : undefined;
}

export function upper(tags: Record<string, string>): Record<string, string> {
  return Object.fromEntries(Object.entries(tags).map(([k, v]) => [k.toUpperCase(), v]));
}
