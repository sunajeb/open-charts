// The element mini toolbar: when the user selects pieces of an open-charts chart on the slide (a
// label, a bar), a small floating toolbar appears next to them (Mac: the pop-up helper's "toolbar"
// panel), offering the formatting those elements take; a change writes per-element overrides into the
// chart's spec (`overrides`, src/core/overrides.ts on the overrides branch; DATA-FORMAT "Overrides")
// and re-lays the chart out through the same flow as Apply.
//
//   PowerPoint (AppEvents → OnSelectionChange → NotifySelection, only when the selected elements of
//   an open-charts chart change) → AppleScriptTask "elementSelection" → background curl, returns at
//   once → POST /api/addin/selection  (an engine request: op "select", line 2 = the OC_ROLE values,
//   tab-separated, empty = nothing selected; then the chart's context, `screen=` of the selection,
//   `seq=` (VBA Timer, orders the background posts), and the chart's OC_SPEC)
//   → the server opens a session for the chart (as Edit data does) and shows the toolbar panel
//   → GET /toolbar (src/web/toolbar.html), ocPopupOpen({id, roles}) → GET /api/addin/properties
//   → a control changes: POST /api/addin/format {id, roles, style} → overrides merged into the
//   session's chart, state "applied" → applyInPowerPoint → ApplyEdit draws it.
//
// The controls come from the property-applicability table (ELEMENT_PROPERTIES / propertiesFor in
// src/core/overrides.ts): the page renders the properties it has a control for (fill, size, bold,
// text colour, rotate today; one renderer each in src/web/toolbar.html CONTROLS) and skips the rest.
// To add a control: add a renderer there keyed by the property name. A built-in table with the same
// names (FALLBACK) stands in if that module cannot be loaded (an older engine copy).
import { readFileSync } from 'node:fs';
import type { IncomingMessage, ServerResponse } from 'node:http';
import { getSession, setSession, type Session } from '../server.ts';
import { parseRequest, rememberContext, specKey, type AddinRequest } from './engine.ts';
import { chartRect, clampInto, screenOf, type PopupHelper, type Rect, type ScreenInfo } from './popup.ts';

export interface ToolbarProperty { prop: string; group: string; value: string }
interface Table { kindOf(role: string): string | undefined; propertiesFor(kind: string): string[]; info: Record<string, { group: string; value: string }> }

/** Stand-in for src/core/overrides.ts (same property names, groups and role kinds). */
const FALLBACK: Table = {
  kindOf: (role) => ({ label: 'label', legend: 'legend', bar: 'segment', wedge: 'segment', area: 'segment', ball: 'segment', marker: 'marker' } as Record<string, string>)[role.split(':')[0]],
  propertiesFor: (kind) => ({ label: ['size', 'rotate', 'fill'], legend: ['size', 'fill'], segment: ['fill'], marker: ['fill'] } as Record<string, string[]>)[kind] ?? [],
  info: {
    size: { group: 'font', value: 'points 4–72' },
    rotate: { group: 'rotation', value: '`-90` (90° left) `0` `90` (90° right)' },
    fill: { group: 'fill', value: 'colour or `none`' },
  },
};

let table: Promise<Table> | undefined;
/** The property table: src/core/overrides.ts when present, else FALLBACK. */
export function propertyTable(): Promise<Table> {
  const path = '../core/overrides.ts';
  return (table ??= import(path).then(
    (m: { kindOf(r: string): string | undefined; propertiesFor(k: string): string[]; ELEMENT_PROPERTIES: Record<string, { group: string; value: string }> }) =>
      ({ kindOf: m.kindOf, propertiesFor: m.propertiesFor, info: m.ELEMENT_PROPERTIES }),
    () => FALLBACK,
  ));
}

/** The properties every selected element takes (the toolbar shows what applies to all of them). */
export function controlsFor(roles: string[], t: Table = FALLBACK): ToolbarProperty[] {
  let common: string[] | undefined;
  for (const role of roles) {
    const kind = t.kindOf(role);
    const props = kind ? t.propertiesFor(kind) : [];
    common = common ? common.filter((p) => props.includes(p)) : props;
  }
  return (common ?? []).map((prop) => ({ prop, group: t.info[prop]?.group ?? '', value: t.info[prop]?.value ?? '' }));
}

/**
 * The chart with `style` merged into the overrides of each role (a null value removes that property;
 * an override left empty is removed).
 */
export function withOverrides<C extends { overrides?: Record<string, Record<string, unknown>> }>(chart: C, roles: string[], style: Record<string, unknown>): C {
  const overrides = { ...(chart.overrides ?? {}) };
  for (const role of roles) {
    const next: Record<string, unknown> = { ...(overrides[role] ?? {}) };
    for (const [k, v] of Object.entries(style)) { if (v === null) delete next[k]; else next[k] = v; }
    if (Object.keys(next).length) overrides[role] = next; else delete overrides[role];
  }
  const out = { ...chart, overrides };
  if (!Object.keys(overrides).length) delete (out as { overrides?: unknown }).overrides;
  return out;
}

/** The current value of each property where every selected element has the same one. */
export function currentStyle(chart: { overrides?: Record<string, Record<string, unknown>> }, roles: string[]): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  const first = chart.overrides?.[roles[0]] ?? {};
  for (const [k, v] of Object.entries(first)) if (roles.every((r) => JSON.stringify(chart.overrides?.[r]?.[k]) === JSON.stringify(v))) out[k] = v;
  return out;
}

export const TOOLBAR_SIZE = { w: 300, h: 44 };

/** Above the selection, left-aligned with it; below when there is no room above; inside the display. */
export function placeToolbar(sel: Rect | undefined, windows: Rect[], screens: ScreenInfo[], size = TOOLBAR_SIZE): Rect | undefined {
  if (!sel) return undefined;
  const box = screenOf(sel, screens)?.visible ?? { x: 0, y: 0, w: 1440, h: 900 };
  const gap = 8;
  const above = { x: sel.x, y: sel.y - gap - size.h, ...size };
  const r = above.y >= box.y ? above : { x: sel.x, y: sel.y + sel.h + gap, ...size };
  return clampInto(r, box);
}

export interface ToolbarDeps {
  base: string;
  helper: PopupHelper;
  /** Runs an applied session in PowerPoint (the server's onSession for "applied"). */
  apply: (s: Session) => void;
  log?: (line: string) => void;
  /** Every selection report, in order (the drag handles, src/addin/handles.ts). */
  onSelect?: (req: AddinRequest) => void;
}

const TOOLBAR_HTML = new URL('../web/toolbar.html', import.meta.url);
const read = async (req: IncomingMessage) => { const c: Buffer[] = []; for await (const d of req) c.push(d as Buffer); return Buffer.concat(c).toString('utf8'); };
const json = (res: ServerResponse, code: number, v: unknown) => res.writeHead(code, { 'content-type': 'application/json', 'cache-control': 'no-store' }).end(JSON.stringify(v));

/** The toolbar's routes on the add-in server. */
export function toolbarRoutes(deps: ToolbarDeps) {
  let lastSeq = -1;
  let current: { id: string; roles: string[] } | undefined;
  return async (req: IncomingMessage, res: ServerResponse): Promise<boolean> => {
    const url = req.url ?? '';
    if (req.method === 'GET' && /^\/toolbar(\?|$)/.test(url)) {
      res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' }).end(readFileSync(TOOLBAR_HTML, 'utf8'));
      return true;
    }
    if (req.method === 'GET' && url.startsWith('/api/addin/properties')) {
      const q = new URL(url, 'http://x').searchParams;
      const roles = (q.get('roles') ?? '').split(',').filter(Boolean);
      const s = getSession(q.get('id') ?? '');
      json(res, 200, { controls: controlsFor(roles, await propertyTable()), style: s ? currentStyle(s.chart as never, roles) : {} });
      return true;
    }
    if (req.method === 'POST' && url === '/api/addin/selection') {
      const text = await read(req);
      res.writeHead(204).end();
      const seq = Number(/^seq=([\d.]+)$/m.exec(text)?.[1] ?? NaN);
      // Background posts can overtake each other: drop an older one (VBA Timer restarts at midnight).
      if (Number.isFinite(seq)) { if (seq < lastSeq && lastSeq - seq < 3600) return true; lastSeq = seq; }
      const req2 = parseRequest(text.replace(/^seq=.*\n/m, ''));
      deps.onSelect?.(req2);
      const roles = req2.arg.split('\t').filter(Boolean);
      if (!roles.length || !req2.spec) { current = undefined; deps.helper.hide('toolbar'); return true; }
      let chart: object;
      try { chart = JSON.parse(req2.spec); } catch { return true; }
      const id = specKey(req2.spec);
      if (!getSession(id)) setSession(id, { chart, title: 'chart', state: 'open' } as never);
      if (req2.context) rememberContext(id, req2.context, req2.spec);
      current = { id, roles };
      const info = await deps.helper.info();
      if (!info || current?.id !== id) return true;
      const frame = placeToolbar(chartRect(req2.where, info.windows, info.screens), info.windows, info.screens);
      if (!frame) { deps.helper.hide('toolbar'); return true; }
      deps.helper.showPanel?.('toolbar', `${deps.base}/toolbar`, { id, roles }, frame);
      return true;
    }
    if (req.method === 'POST' && url === '/api/addin/format') {
      let b: { id?: string; roles?: string[]; style?: Record<string, unknown> };
      try { b = JSON.parse(await read(req)); } catch { json(res, 400, { error: 'bad JSON' }); return true; }
      const s = b.id ? getSession(b.id) : undefined;
      if (!s || !Array.isArray(b.roles) || !b.roles.length || !b.style) { json(res, 404, { error: 'Select the chart element again.' }); return true; }
      const saved = setSession(s.id, { chart: withOverrides(s.chart as never, b.roles, b.style), state: 'applied' } as never);
      deps.log?.(`TOOLBAR ${s.id} ${b.roles.join(',')} ${JSON.stringify(b.style)}`);
      json(res, 200, { rev: saved.rev });
      deps.apply(saved);
      return true;
    }
    return false;
  };
}
