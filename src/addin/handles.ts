// Drag handles on the slide (Mac): the server side of direct manipulation. When an open-charts
// chart (or an element of one) is selected, the Mac helper's transparent overlay (addin/popup-mac,
// HandleOverlay) shows small handles over it — column gap, value-axis top, level arrows, the
// selected labels (src/core/handles.ts) — and catches the mouse only while the pointer is on one, so
// every other click goes to PowerPoint. Dragging one is live: each move comes here, the chart is laid
// out with the changed spec (< 1 ms warm) and its shapes go back to the overlay, which draws them
// over the chart at once (a ghost); the mouse-up commits the change through the Apply flow (one
// ApplyEdit, one undo step in the session's chart). A double-click resets the handle.
//
//   PowerPoint selection -> NotifySelection (VBA) -> POST /api/addin/selection (with view=: the chart
//   group's corners on screen) -> select(): layout, handles -> helper {"cmd":"handles"}
//   helper {"event":"handle","phase":"down|move|up|reset","id","x","y"} -> layout -> {"cmd":"ghost"}
//   up -> session "applied" -> applyInPowerPoint -> re-laid-out handles once drawn
//   helper {"event":"remeasure"} (scroll, zoom, window moved: the overlay hid itself) -> the
//   SelectionInfo macro answers the same request as NotifySelection -> select()
//
// Positions: slide points map to global screen points (top left of the main display, as CGWindowList
// reports windows) by one zoom factor and offset, measured from the chart group's corners through
// PointsToScreenPixelsX/Y (which may answer in pixels: popup.ts chartRect works out which).
import type { ChartSpec, DragHandle, Frame, Measurer, Scene, Shape } from '../core/index.ts';
import { describeDrag, dragHandles, dragSpec, resetSpec } from '../core/index.ts';
import { flattenPath, rectGeom } from '../core/path.ts';
import { getSession, setSession, type Session } from '../server.ts';
import { contextOf, parseRequest, rememberContext, sceneBounds, sceneFor, specKey, type AddinRequest } from './engine.ts';
import { chartRect, type PopupHelper, type Rect, type ScreenInfo } from './popup.ts';

/** Slide points -> global screen points: screen = slide * k + offset. */
export interface View { k: number; ox: number; oy: number }

/**
 * The view of the chart's slide from the chart group's bounds (slide pt) and its corners on screen
 * (`view=`): undefined when they cannot be trusted (outside every PowerPoint window at any display
 * scale, or not one zoom factor across and down).
 */
export function viewFrom(corners: [number, number, number, number] | undefined, bounds: Frame | undefined, windows: Rect[], screens: ScreenInfo[]): View | undefined {
  if (!corners || !bounds || !(bounds.w > 0) || !(bounds.h > 0)) return undefined;
  const r = chartRect({ screen: corners }, windows, screens);
  if (!r) return undefined;
  const kx = r.w / bounds.w, ky = r.h / bounds.h;
  // Corners come in whole pixels: allow a pixel's worth of difference between the two factors.
  if (!(kx > 0) || Math.abs(kx - ky) > Math.max(0.02 * kx, 2 / Math.min(bounds.w, bounds.h))) return undefined;
  const k = (kx + ky) / 2;
  return { k, ox: r.x - bounds.x * k, oy: r.y - bounds.y * k };
}

export const toScreen = (v: View, x: number, y: number) => ({ x: x * v.k + v.ox, y: y * v.k + v.oy });
export const toSlide = (v: View, x: number, y: number) => ({ x: (x - v.ox) / v.k, y: (y - v.oy) / v.k });

/** Size of a handle marker on screen (pt), whatever the zoom. */
export const MARKER = 9;
/** The overlay reaches this far (screen pt) beyond the chart, for markers on its edge and the ghost. */
const MARGIN = 28;

export interface ScreenHandle { id: string; rect: [number, number, number, number]; marker: boolean; cursor: DragHandle['cursor'] }

/** Handles on screen: a marker of MARKER points centred on its point, or the element's box. */
export function screenHandles(handles: DragHandle[], v: View): ScreenHandle[] {
  return handles.map((h) => {
    if (h.box) {
      const p = toScreen(v, h.box.x, h.box.y);
      const w = Math.max(8, h.box.w * v.k), hh = Math.max(8, h.box.h * v.k);
      return { id: h.id, rect: [p.x - (w - h.box.w * v.k) / 2, p.y - (hh - h.box.h * v.k) / 2, w, hh], marker: false, cursor: h.cursor };
    }
    const c = toScreen(v, h.at.x, h.at.y);
    return { id: h.id, rect: [c.x - MARKER / 2, c.y - MARKER / 2, MARKER, MARKER], marker: true, cursor: h.cursor };
  });
}

const intersect = (a: Rect, b: Rect): Rect | undefined => {
  const x = Math.max(a.x, b.x), y = Math.max(a.y, b.y);
  const w = Math.min(a.x + a.w, b.x + b.w) - x, h = Math.min(a.y + a.h, b.y + b.h) - y;
  return w > 0 && h > 0 ? { x, y, w, h } : undefined;
};
const inside = (x: number, y: number, r: Rect) => x >= r.x && x <= r.x + r.w && y >= r.y && y <= r.y + r.h;

/** The overlay's frame: the chart (and its frame) on screen plus a margin, inside its PowerPoint window. */
export function overlayFrame(scene: Scene, v: View, windows: Rect[]): Rect | undefined {
  const b = sceneBounds(scene)!;
  const p = toScreen(v, b.x, b.y);
  const r = { x: p.x - MARGIN, y: p.y - MARGIN, w: b.w * v.k + 2 * MARGIN, h: b.h * v.k + 2 * MARGIN };
  const win = windows.find((w) => inside(p.x + (b.w * v.k) / 2, p.y + (b.h * v.k) / 2, w));
  return win ? intersect(r, win) : windows.length ? undefined : r;
}

// ---- The ghost: the re-laid-out chart as the overlay draws it ------------------------------------

const hex = (c: string | undefined): string | null => {
  const m = /^#?([0-9A-Fa-f]{6})/.exec(c ?? '');
  return m ? m[1].toUpperCase() : null;
};
const r1 = (n: number) => Math.round(n * 10) / 10;

/**
 * Shapes in screen points for the overlay (short keys: one JSON line per frame). The first is a
 * backdrop in the slide's background colour over `cover` (the chart as PowerPoint still shows it).
 */
export function ghostShapes(scene: Scene, v: View, cover: Frame, background = 'FFFFFF'): object[] {
  const P = (x: number, y: number) => { const p = toScreen(v, x, y); return [r1(p.x), r1(p.y)]; };
  const k = v.k;
  const [bx, by] = P(cover.x - 2, cover.y - 2);
  const out: object[] = [{ t: 'r', x: bx, y: by, w: r1((cover.w + 4) * k), h: r1((cover.h + 4) * k), f: hex(background) ?? 'FFFFFF' }];
  for (const s of scene.shapes as Shape[]) {
    if (s.kind === 'rect') {
      const [x, y] = P(s.x, s.y);
      out.push({ t: 'r', x, y, w: r1(s.w * k), h: r1(s.h * k), f: hex(s.fill), s: hex(s.stroke), sw: r1((s.strokeWidth ?? 0.75) * k), g: rectGeom(s) });
    } else if (s.kind === 'line') {
      const [x1, y1] = P(s.x1, s.y1), [x2, y2] = P(s.x2, s.y2);
      out.push({ t: 'l', x1, y1, x2, y2, c: hex(s.color), w: r1(Math.max(0.5, s.width * k)), ...(s.dash ? { d: 1 } : {}), ...(s.arrowEnd ? { ae: 1 } : {}), ...(s.arrowStart ? { as: 1 } : {}) });
    } else if (s.kind === 'path') {
      for (const sub of flattenPath(s.commands)) {
        out.push({ t: 'p', p: sub.points.flatMap((q) => P(q.x, q.y)), z: sub.closed ? 1 : 0, f: hex(s.fill), s: hex(s.stroke), sw: r1((s.strokeWidth ?? 0.75) * k) });
      }
    } else if (s.kind === 'text') {
      const [x, y] = P(s.x, s.y);
      out.push({
        t: 'x', x, y, w: r1(s.w * k), h: r1(s.h * k), s: s.text, z: r1(s.size * k), c: hex(s.color), f: hex(s.fill), k: hex(s.stroke),
        a: s.align[0], v: s.valign[0], ...(s.bold ? { b: 1 } : {}), ...(s.rotate ? { r: s.rotate } : {}), ...(s.rounded ? { o: 1 } : {}),
      });
    }
  }
  return out;
}

// ---- The controller ------------------------------------------------------------------------------

export interface HandlesDeps {
  helper: PopupHelper;
  measurer: () => Measurer;
  /** Runs an applied session in PowerPoint (applyInPowerPoint); resolves once it has drawn or failed. */
  apply: (s: Session) => Promise<void> | void;
  /** The selection request as NotifySelection builds it, now (the SelectionInfo macro): re-measuring. */
  measure?: () => Promise<string | undefined>;
  /** Steps and timings (the add-in server: verbose logging only). */
  log?: (line: string) => void;
  /** Failures (default: `log`; the add-in server: always logged). */
  warn?: (line: string) => void;
  /** Clock for the drag timing in the log (tests). */
  now?: () => number;
}

interface Current {
  id: string;
  chart: ChartSpec;
  context: string;
  roles: string[];
  view: View;
  windows: Rect[];
  scene: Scene;
  relayout: (spec: ChartSpec) => Scene;
  handles: DragHandle[];
  background: string;
}

interface Drag { h: DragHandle; x0: number; y0: number; last: string; spec: ChartSpec; t0: number; frames: number }

type Msg = Record<string, unknown> & { event?: string };

/** Drag handles for the selected chart: what the selection route and the helper's events call. */
export function handlesController(deps: HandlesDeps) {
  const { helper } = deps;
  const log = deps.log ?? (() => {});
  const warn = deps.warn ?? log;
  const now = deps.now ?? (() => performance.now());
  let cur: Current | undefined;
  let drag: Drag | undefined;
  let pending: { x: number; y: number } | undefined;
  let scheduled = false;
  let committing = false;
  let seq = 0;

  const send = (m: object) => helper.send?.(m);
  const hide = () => { cur = undefined; drag = undefined; send({ cmd: 'handles', hide: true }); };

  /** Shows the current chart's handles (and drops any ghost). */
  const show = () => {
    if (!cur) return;
    const frame = overlayFrame(cur.scene, cur.view, cur.windows);
    const handles = cur.handles.length && frame ? screenHandles(cur.handles, cur.view).filter((h) => inside(h.rect[0] + h.rect[2] / 2, h.rect[1] + h.rect[3] / 2, frame)) : [];
    if (!frame || !handles.length) { send({ cmd: 'handles', hide: true }); return; }
    send({ cmd: 'handles', frame: [frame.x, frame.y, frame.w, frame.h], handles, k: cur.view.k });
  };

  const backgroundOf = (context: string) => {
    // theme= is dk1 lt1 …: the slide's background is lt1 on a light slide.
    const theme = /^theme=([^\n]*)$/m.exec(context)?.[1]?.split(',');
    return theme?.[1] && /^[0-9A-Fa-f]{6}$/.test(theme[1]) ? theme[1] : 'FFFFFF';
  };

  /** A selection (NotifySelection, or SelectionInfo after scrolling): the handles for it, or none. */
  async function select(req: AddinRequest): Promise<void> {
    const mine = ++seq;
    if (drag || committing) return; // the overlay owns the mouse: PowerPoint did not see a click
    const corners = req.where?.view;
    if (!req.spec || !corners || !req.context) { hide(); return; }
    let chart: ChartSpec;
    try { chart = JSON.parse(req.spec); } catch { hide(); return; }
    const info = await helper.info();
    if (mine !== seq || drag) return;
    if (!info) { hide(); return; }
    const view = viewFrom(corners, req.bounds, info.windows, info.screens);
    if (!view) { log(`HANDLES no trusted view (view=${corners.join(',')}, bounds ${JSON.stringify(req.bounds)})`); hide(); return; }
    const id = specKey(req.spec);
    try {
      const { scene, relayout } = sceneFor(chart, req.context, deps.measurer());
      const roles = req.arg.split('\t').filter(Boolean);
      cur = { id, chart, context: req.context, roles, view, windows: info.windows, scene, relayout, handles: dragHandles(chart, scene, roles), background: backgroundOf(req.context) };
      if (!getSession(id)) setSession(id, { chart, title: 'chart', state: 'open' } as never);
      rememberContext(id, req.context, req.spec);
    } catch (err) {
      warn(`HANDLES layout failed: ${(err as Error).message}`);
      hide();
      return;
    }
    show();
  }

  const ghostFor = (spec: ChartSpec, c: Current) => {
    const scene = c.relayout(spec);
    const a = sceneBounds(c.scene)!, b = sceneBounds(scene)!;
    const x = Math.min(a.x, b.x), y = Math.min(a.y, b.y);
    const cover = { x, y, w: Math.max(a.x + a.w, b.x + b.w) - x, h: Math.max(a.y + a.h, b.y + b.h) - y };
    return ghostShapes(scene, c.view, cover, c.background);
  };

  /** The latest pointer position of a drag: re-laid out and sent as a ghost (coalesced: newest wins). */
  const step = () => {
    scheduled = false;
    const c = cur, d = drag, p = pending;
    pending = undefined;
    if (!c || !d || !p) return;
    const a = toSlide(c.view, d.x0, d.y0), b = toSlide(c.view, p.x, p.y);
    const spec = dragSpec(c.chart, c.scene, d.h, b.x - a.x, b.y - a.y, c.relayout);
    const key = JSON.stringify(spec);
    if (key === d.last) return;
    d.last = key;
    d.spec = spec;
    d.frames++;
    send({ cmd: 'ghost', shapes: ghostFor(spec, c), tip: describeDrag(spec, d.h) });
  };

  async function commit(c: Current, spec: ChartSpec, what: string) {
    committing = true;
    try {
      if (!getSession(c.id)) setSession(c.id, { chart: c.chart, title: 'chart', state: 'open' } as never);
      const saved = setSession(c.id, { chart: spec, state: 'applied' } as never);
      log(`HANDLES ${what}`);
      c.chart = spec;
      await deps.apply(saved);
      const after = getSession(c.id);
      if (after?.state === 'drawn') {
        // Drawn: the chart's new context (bounds may have changed) and its new layout.
        c.context = contextOf(c.id) ?? c.context;
        const { scene, relayout } = sceneFor(spec, c.context, deps.measurer());
        Object.assign(c, { scene, relayout, handles: dragHandles(spec, scene, c.roles) });
      } else {
        warn(`HANDLES apply did not draw: ${after?.state ?? 'no session'} ${after?.message ?? ''}`);
      }
    } finally {
      committing = false;
    }
    if (cur === c) show(); else send({ cmd: 'ghost', clear: true });
  }

  /** An event from the helper. */
  let stateWaiter: ((m: Msg | undefined) => void) | undefined;

  /** What is shown now, for tests and the live check (GET /api/addin/handles): handles, labels and the overlay's own state. */
  async function state() {
    const overlay = await new Promise<Msg | undefined>((resolve) => {
      const t = setTimeout(() => { stateWaiter = undefined; resolve(undefined); }, 1500);
      stateWaiter = (msg) => { clearTimeout(t); stateWaiter = undefined; resolve(msg); };
      send({ cmd: 'handlesState' });
    });
    if (!cur) return { overlay };
    const c = cur;
    const labels = c.scene.shapes.flatMap((s, i) => {
      const role = c.scene.roles?.[i] ?? '';
      if (s.kind !== 'text' || !role.startsWith('label:')) return [];
      const p = toScreen(c.view, s.x, s.y);
      return [{ role, rect: [p.x, p.y, s.w * c.view.k, s.h * c.view.k] }];
    });
    return { id: c.id, view: c.view, roles: c.roles, options: c.chart.options ?? {}, overrides: c.chart.overrides ?? {}, handles: screenHandles(c.handles, c.view), labels, overlay };
  }

  function event(m: Msg): void {
    if (m.event === 'handlesState') { stateWaiter?.(m); return; }
    if (m.event === 'remeasure') {
      if (!deps.measure || drag || committing) return;
      const measure = deps.measure;
      // Only while PowerPoint shows a window (the helper lists them; no Apple event): a window that
      // closed as PowerPoint quits also asks to re-measure, and a macro sent then starts it again.
      void helper.info().then((info) => (info?.windows.length ? measure() : undefined))
        .then((text) => (text ? select(parseRequest(text.replace(/^seq=.*\n/m, ''))) : hide()));
      return;
    }
    if (m.event !== 'handle') return;
    const x = Number(m.x), y = Number(m.y), phase = String(m.phase);
    const c = cur;
    if (!c) { send({ cmd: 'ghost', clear: true }); return; }
    if (phase === 'down') {
      const h = c.handles.find((k) => k.id === m.id);
      if (!h || !Number.isFinite(x) || !Number.isFinite(y)) return;
      drag = { h, x0: x, y0: y, last: JSON.stringify(c.chart), spec: c.chart, t0: now(), frames: 0 };
      return;
    }
    if (phase === 'move') {
      if (!drag || !Number.isFinite(x) || !Number.isFinite(y)) return;
      pending = { x, y };
      if (!scheduled) { scheduled = true; setImmediate(step); }
      return;
    }
    if (phase === 'up' || phase === 'cancel') {
      const d = drag;
      if (!d) return;
      if (phase === 'up' && Number.isFinite(x) && Number.isFinite(y)) { pending = { x, y }; step(); }
      drag = undefined;
      pending = undefined;
      const ms = now() - d.t0;
      const changed = phase === 'up' && d.last !== JSON.stringify(c.chart);
      if (!changed) { send({ cmd: 'ghost', clear: true }); return; }
      void commit(c, d.spec, `${d.h.id} -> ${describeDrag(d.spec, d.h)} (${d.frames} frames in ${Math.round(ms)} ms)`);
      return;
    }
    if (phase === 'reset') {
      const h = c.handles.find((k) => k.id === m.id);
      if (!h || drag) return;
      const spec = resetSpec(c.chart, h);
      if (spec === c.chart) return;
      void commit(c, spec, `${h.id} reset`);
    }
  }

  return { select, event, state, get current() { return cur; }, get dragging() { return !!drag; } };
}
