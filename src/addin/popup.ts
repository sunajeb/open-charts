// The "Edit data" pop-up: the editor's datasheet (`?session=<id>&popup=1`, src/web/editor.html) in a
// small window over PowerPoint, next to the chart, instead of a browser tab.
//
//   Mac      a resident native helper (addin/popup-mac/main.swift, built by `npm run build:popup`,
//            installed by `addin install` as OpenChartsPopup.app in the add-in's home): an
//            always-on-top NSPanel hosting a WKWebView whose page is preloaded, so showing a chart is
//            one JavaScript call (no page load, no browser start). The add-in server starts it and
//            talks to it over stdin/stdout, one JSON object per line (see HELPER PROTOCOL below).
//   Windows  Microsoft Edge in app mode (`msedge --app=URL --window-size --window-position`): a
//            borderless window without tabs or address bar; ships with Windows 10/11.
//   else     (or when the helper cannot start) the default browser, as before.
//
// Positions are global screen coordinates with the origin at the top left of the main display, in
// points on the Mac (as CGWindowList reports windows) and in pixels on Windows.
//
// HELPER PROTOCOL (Mac), server -> helper:
//   {"cmd":"load","panel":"datasheet","url":U}       preload U in that panel (hidden)
//   {"cmd":"info"}                                   -> {"event":"info","screens":[…],"windows":[…]}
//   {"cmd":"show","panel":P,"id":ID,"url":U,"frame":[x,y,w,h],"focus":true}
//        calls window.ocPopupOpen(ID) in the preloaded page (or loads U if the page cannot), orders
//        the panel front once the page reports ready -> {"event":"shown","panel":P,"ms":N,"ready":bool}
//   {"cmd":"hide","panel":P}                         -> {"event":"closed","panel":P}
//   {"cmd":"eval","panel":P,"js":S}                  -> {"event":"eval","result":…} (tests)
// helper -> server, unsolicited: {"event":"loaded","panel":P}, {"event":"closed","panel":P} (the page
// or the user closed it), {"event":"error","message":…}.
// Panels are named so the helper can host more than one (the datasheet; the element mini toolbar).
import { spawn, type ChildProcess } from 'node:child_process';
import { existsSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { addinHome } from './launcher.ts';

export interface Rect { x: number; y: number; w: number; h: number }
export interface ScreenInfo { frame: Rect; visible: Rect; scale?: number }

/**
 * Where Edit data says the chart is, from the request's `screen=` and `window=` lines (VBA):
 *   screen  the chart's corners through PowerPoint's PointsToScreenPixelsX/Y (left, top, right, bottom)
 *   window  PowerPoint's document window: left, top, width, height (pt) and zoom (%)
 */
export interface ChartWhere {
  screen?: [number, number, number, number]; window?: [number, number, number, number, number];
  /** The chart group's corners (line 4 of the request: its bounds) the same way: drag handles (handles.ts). */
  view?: [number, number, number, number];
}

/** The default pop-up size (pt): a datasheet of ~6 columns × 6 rows with the compact toolbar. */
export const POPUP_SIZE = { w: 620, h: 360 };
const GAP = 12;

export function parseWhere(params: Record<string, string>): ChartWhere | undefined {
  const nums = (v: string | undefined, n: number) => {
    const a = (v ?? '').split(',').map(Number);
    return a.length === n && a.every(Number.isFinite) ? a : undefined;
  };
  const screen = nums(params.screen, 4) as ChartWhere['screen'];
  const window = nums(params.window, 5) as ChartWhere['window'];
  const view = nums(params.view, 4) as ChartWhere['view'];
  return screen || window || view ? { ...(screen ? { screen } : {}), ...(window ? { window } : {}), ...(view ? { view } : {}) } : undefined;
}

const inside = (x: number, y: number, r: Rect) => x >= r.x && x <= r.x + r.w && y >= r.y && y <= r.y + r.h;
const area = (r: Rect) => Math.max(0, r.w) * Math.max(0, r.h);
const overlap = (a: Rect, b: Rect): Rect => {
  const x = Math.max(a.x, b.x), y = Math.max(a.y, b.y);
  return { x, y, w: Math.min(a.x + a.w, b.x + b.w) - x, h: Math.min(a.y + a.h, b.y + b.h) - y };
};

/**
 * The chart's rectangle on screen, or undefined when it cannot be trusted. PowerPoint for Mac's
 * PointsToScreenPixels may answer in pixels (Retina: 2 per point) or not at all, so the corners are
 * accepted only when, at scale 1 or at a display's scale, the chart's centre falls inside one of
 * PowerPoint's windows (from the helper; front-most first).
 */
export function chartRect(where: ChartWhere | undefined, windows: Rect[], screens: ScreenInfo[]): Rect | undefined {
  const s = where?.screen;
  if (!s || !(s[2] > s[0]) || !(s[3] > s[1])) return undefined;
  const scales = [...new Set([1, ...screens.map((x) => x.scale ?? 1)])];
  for (const k of scales) {
    const r = { x: s[0] / k, y: s[1] / k, w: (s[2] - s[0]) / k, h: (s[3] - s[1]) / k };
    if (!windows.length || windows.some((w) => inside(r.x + r.w / 2, r.y + r.h / 2, w))) return r;
  }
  return undefined;
}

/** The display a rectangle is mostly on (or the main one). */
export function screenOf(r: Rect | undefined, screens: ScreenInfo[]): ScreenInfo | undefined {
  if (!screens.length) return undefined;
  if (!r) return screens[0];
  let best = screens[0], most = -1;
  for (const s of screens) {
    const a = area(overlap(r, s.frame)) || (inside(r.x + r.w / 2, r.y + r.h / 2, s.frame) ? 1 : 0);
    if (a > most) { most = a; best = s; }
  }
  return best;
}

/** Moves (and if need be shrinks) `r` into `box`. */
export function clampInto(r: Rect, box: Rect): Rect {
  const w = Math.min(r.w, box.w), h = Math.min(r.h, box.h);
  return {
    x: Math.min(Math.max(r.x, box.x), box.x + box.w - w),
    y: Math.min(Math.max(r.y, box.y), box.y + box.h - h),
    w, h,
  };
}

/**
 * Where the pop-up goes: next to the chart — right of it, else left, below, above (the first that
 * fits the display's visible area; else the one that shows most, moved inside), top-aligned with the
 * chart. Without a trusted chart rectangle: inside PowerPoint's front window at its right edge, below
 * the ribbon. Without either: centred on the main display.
 */
export function placePopup(chart: Rect | undefined, windows: Rect[], screens: ScreenInfo[], size = POPUP_SIZE): Rect {
  const anchor = chart ?? windows[0];
  const screen = screenOf(anchor, screens);
  const box = screen?.visible ?? { x: 0, y: 0, w: 1440, h: 900 };
  const fit = (r: Rect) => clampInto(r, box);
  if (chart) {
    const { w, h } = size;
    const top = chart.y, left = chart.x;
    const candidates: Rect[] = [
      { x: chart.x + chart.w + GAP, y: top, w, h },
      { x: chart.x - GAP - w, y: top, w, h },
      { x: left, y: chart.y + chart.h + GAP, w, h },
      { x: left, y: chart.y - GAP - h, w, h },
    ];
    const fits = candidates.find((c) => area(overlap(c, box)) >= area(c) - 0.5);
    if (fits) return fits;
    // Nothing fits whole: the side with most room, moved inside, so it covers as little as possible.
    const best = candidates.reduce((a, b) => (area(overlap(b, box)) > area(overlap(a, box)) ? b : a));
    return fit(best);
  }
  const win = windows[0];
  if (win) return fit({ x: win.x + win.w - size.w - 24, y: win.y + Math.min(170, win.h / 4), ...size });
  return fit({ x: box.x + (box.w - size.w) / 2, y: box.y + (box.h - size.h) / 2, ...size });
}

/**
 * Once the user has dragged or resized the pop-up, it opens where they left it (as the reference's
 * datasheet does), moved back onto a display if that one is gone.
 */
export function keepSaved(saved: Rect, screens: ScreenInfo[]): Rect {
  const s = screens.find((x) => saved.x + saved.w / 2 >= x.frame.x && saved.x + saved.w / 2 <= x.frame.x + x.frame.w && saved.y >= x.frame.y && saved.y <= x.frame.y + x.frame.h) ?? screens[0];
  const w = Math.max(360, saved.w), h = Math.max(220, saved.h);
  return s ? clampInto({ ...saved, w, h }, s.visible) : { ...saved, w, h };
}

/** The pop-up page for an edit session. */
export const popupUrl = (base: string, id: string) => `${base}/?session=${encodeURIComponent(id)}&popup=1`;
/** The page the Mac helper preloads (no session yet: it waits for ocPopupOpen). */
export const preloadUrl = (base: string) => `${base}/?popup=1`;

// ---- Windows: Edge app mode ---------------------------------------------------------------------

/** Edge's usual install locations (it ships with Windows 10/11). */
export function edgePaths(env: NodeJS.ProcessEnv = process.env): string[] {
  const roots = [env['ProgramFiles(x86)'], env.ProgramFiles, env.LOCALAPPDATA].filter(Boolean) as string[];
  return roots.map((r) => `${r}\\Microsoft\\Edge\\Application\\msedge.exe`);
}

/**
 * msedge arguments for a borderless pop-up at `rect` (pixels). Its own profile folder, so Edge starts
 * a separate instance that honours --window-size/--window-position (an existing Edge window would
 * otherwise take the URL as a tab and ignore both) and keeps the pop-up out of the user's history.
 */
export function edgeArgs(url: string, rect: Rect, profile: string): string[] {
  const r = (n: number) => String(Math.round(n));
  return [
    `--app=${url}`,
    `--window-size=${r(rect.w)},${r(rect.h)}`,
    `--window-position=${r(rect.x)},${r(rect.y)}`,
    `--user-data-dir=${profile}`,
    '--no-first-run', '--no-default-browser-check', '--disable-features=msEdgeSidebarV2,Translate',
  ];
}

// ---- Mac: the native helper ---------------------------------------------------------------------

const ROOT = fileURLToPath(new URL('../../', import.meta.url));
export const HELPER_APP = 'OpenChartsPopup.app';
export const HELPER_EXE = 'Contents/MacOS/open-charts-popup';

/** The helper's executable: the installed copy, else a development build in dist/. */
export function helperPath(home = addinHome()): string | undefined {
  for (const app of [join(home, HELPER_APP), join(ROOT, 'dist', HELPER_APP)]) {
    const exe = join(app, HELPER_EXE);
    if (existsSync(exe)) return exe;
  }
  return undefined;
}

type Msg = Record<string, unknown> & { event?: string };

export interface PopupHelper {
  /** Starts the helper (if it is not running) and preloads the datasheet page. */
  warm(): boolean;
  /** Shows session `id` next to the chart; resolves with the helper's "shown" answer, or undefined. */
  show(base: string, id: string, where?: ChartWhere, opts?: { focus?: boolean; size?: { w: number; h: number } }): Promise<Msg | undefined>;
  info(): Promise<{ windows: Rect[]; screens: ScreenInfo[]; saved?: Rect } | undefined>;
  /** Shows another panel (the element toolbar): ocPopupOpen(arg) in its preloaded page, at `frame`, without taking the keyboard. */
  showPanel?(panel: string, url: string, arg: unknown, frame: Rect): void;
  hide(panel?: string): void;
  eval(js: string, panel?: string): Promise<Msg | undefined>;
  /** Sends one command as is (the drag-handle overlay: src/addin/handles.ts). */
  send?(m: object): void;
  /** Every message the helper sends (the overlay's drag events). */
  onEvent?(listener: (m: Msg) => void): void;
  stop(): void;
  readonly running: boolean;
}

const toRect = (a: unknown): Rect | undefined =>
  Array.isArray(a) && a.length === 4 && a.every((n) => typeof n === 'number') ? { x: a[0], y: a[1], w: a[2], h: a[3] } : undefined;

/**
 * The resident Mac helper: one process for the server's lifetime (it exits when its stdin closes,
 * i.e. with the server). `exe` undefined = not built/installed: show() resolves undefined and the
 * caller falls back to the browser.
 */
export function popupHelper(base: string, exe: string | undefined, log: (line: string) => void = () => {}): PopupHelper {
  let child: ChildProcess | undefined;
  let buffer = '';
  const waiters: Array<{ event: string; resolve: (m: Msg | undefined) => void; timer: NodeJS.Timeout }> = [];
  const listeners: Array<(m: Msg) => void> = [];
  const send = (m: object) => { try { child?.stdin?.write(`${JSON.stringify(m)}\n`); } catch { /* reported by exit */ } };
  const expect = (event: string, ms: number) => new Promise<Msg | undefined>((resolve) => {
    const w = { event, resolve, timer: setTimeout(() => { waiters.splice(waiters.indexOf(w), 1); resolve(undefined); }, ms) };
    waiters.push(w);
  });
  const start = () => {
    if (child && child.exitCode === null && !child.killed) return true;
    if (!exe) return false;
    let c: ChildProcess;
    try { c = spawn(exe, [], { stdio: ['pipe', 'pipe', 'ignore'] }); } catch { return false; }
    buffer = '';
    c.stdout!.setEncoding('utf8');
    c.stdout!.on('data', (d: string) => {
      buffer += d;
      for (let nl = buffer.indexOf('\n'); nl >= 0; nl = buffer.indexOf('\n')) {
        const line = buffer.slice(0, nl);
        buffer = buffer.slice(nl + 1);
        let m: Msg;
        try { m = JSON.parse(line); } catch { continue; }
        if (m.event === 'error') log(`POPUP helper: ${String(m.message)}`);
        const i = waiters.findIndex((w) => w.event === m.event);
        if (i >= 0) { const [w] = waiters.splice(i, 1); clearTimeout(w.timer); w.resolve(m); }
        for (const l of listeners) { try { l(m); } catch (err) { log(`POPUP listener: ${(err as Error).message}`); } }
      }
    });
    const gone = () => {
      if (child !== c) return;
      child = undefined;
      for (const w of waiters.splice(0)) { clearTimeout(w.timer); w.resolve(undefined); }
    };
    c.on('exit', gone);
    c.on('error', gone);
    c.stdin!.on('error', () => {});
    c.unref();
    (c.stdout as unknown as { unref?: () => void }).unref?.();
    child = c;
    send({ cmd: 'load', panel: 'datasheet', url: preloadUrl(base) });
    send({ cmd: 'load', panel: 'toolbar', url: `${base}/toolbar` });
    return true;
  };
  const helper: PopupHelper = {
    get running() { return !!child && child.exitCode === null; },
    warm: start,
    async info() {
      if (!start()) return undefined;
      const answer = expect('info', 1000);
      send({ cmd: 'info' });
      const m = await answer;
      if (!m) return undefined;
      const windows = ((m.windows as unknown[]) ?? []).map(toRect).filter(Boolean) as Rect[];
      const screens = ((m.screens as Array<{ frame: unknown; visible: unknown; scale?: number }>) ?? [])
        .map((s) => ({ frame: toRect(s.frame)!, visible: toRect(s.visible)!, scale: s.scale }))
        .filter((s) => s.frame && s.visible);
      const saved = toRect((m.saved as Record<string, unknown> | undefined)?.datasheet);
      return { windows, screens, ...(saved ? { saved } : {}) };
    },
    async show(base2, id, where, opts = {}) {
      if (!start()) return undefined;
      const info = await helper.info();
      if (!info) return undefined;
      const frame = info.saved ? keepSaved(info.saved, info.screens) : placePopup(chartRect(where, info.windows, info.screens), info.windows, info.screens, opts.size);
      const answer = expect('shown', 3000);
      send({ cmd: 'show', panel: 'datasheet', id, url: popupUrl(base2, id), frame: [frame.x, frame.y, frame.w, frame.h], focus: opts.focus !== false });
      return answer;
    },
    hide(panel = 'datasheet') { send({ cmd: 'hide', panel }); },
    showPanel(panel, url, arg, frame) {
      if (start()) send({ cmd: 'show', panel, id: arg, url, frame: [frame.x, frame.y, frame.w, frame.h], focus: false });
    },
    async eval(js, panel = 'datasheet') {
      if (!start()) return undefined;
      const answer = expect('eval', 3000);
      send({ cmd: 'eval', panel, js });
      return answer;
    },
    send(m) { if (start()) send(m); },
    onEvent(listener) { listeners.push(listener); },
    stop() { child?.stdin?.end(); child = undefined; },
  };
  return helper;
}

// ---- One entry point per platform -----------------------------------------------------------------

export interface OpenPopupDeps {
  platform?: NodeJS.Platform;
  /** Mac: the resident helper (the add-in server's). */
  helper?: PopupHelper;
  /** Starts a process detached (Windows: msedge). */
  launch?: (cmd: string, args: string[]) => boolean;
  /** The default browser (the fallback). */
  browser?: (url: string) => void;
  exists?: (path: string) => boolean;
  log?: (line: string) => void;
  /** Windows: Edge's profile folder for pop-ups. */
  profile?: string;
  env?: NodeJS.ProcessEnv;
}

const defaultLaunch = (cmd: string, args: string[]) => {
  try {
    const c = spawn(cmd, args, { detached: true, stdio: 'ignore' });
    c.on('error', () => {});
    c.unref();
    return true;
  } catch { return false; }
};
const defaultBrowser = (url: string) => {
  const [cmd, args] = process.platform === 'win32' ? ['cmd', ['/c', 'start', '', url]] : ['open', [url]];
  defaultLaunch(cmd as string, args as string[]);
};

/**
 * Shows edit session `id` (served at `base`) in a pop-up next to the chart, on any platform: the Mac
 * helper, Edge app mode on Windows, the default browser otherwise or when those fail. `where` is the
 * chart's position as Edit data reported it (Windows: `screen` in pixels from PointsToScreenPixels).
 * Resolves with what was used.
 */
export async function openPopup(base: string, id: string, where: ChartWhere | undefined, deps: OpenPopupDeps = {}): Promise<'helper' | 'edge' | 'browser'> {
  const platform = deps.platform ?? process.platform;
  const browser = deps.browser ?? defaultBrowser;
  const log = deps.log ?? (() => {});
  if (platform === 'darwin' && deps.helper) {
    const t = Date.now();
    const shown = await deps.helper.show(base, id, where).catch(() => undefined);
    if (shown) {
      log(`POPUP shown in ${Date.now() - t} ms (page ${shown.ready ? 'ready' : 'not ready'} after ${String(shown.ms)} ms)`);
      return 'helper';
    }
    log('POPUP helper did not answer: opening the editor in the browser');
  } else if (platform === 'win32') {
    const exists = deps.exists ?? existsSync;
    const edge = edgePaths(deps.env).find(exists);
    const s = where?.screen;
    const chart = s && s[2] > s[0] && s[3] > s[1] ? { x: s[0], y: s[1], w: s[2] - s[0], h: s[3] - s[1] } : undefined;
    // Without the display's work area, place by the chart alone (Windows keeps a window on screen).
    const screens: ScreenInfo[] = chart ? [{ frame: { x: -1e5, y: -1e5, w: 2e5, h: 2e5 }, visible: { x: 0, y: 0, w: 1e5, h: 1e5 } }] : [];
    const rect = placePopup(chart, [], screens);
    const profile = deps.profile ?? join(addinHome(), 'edge-popup');
    if (edge && (deps.launch ?? defaultLaunch)(edge, edgeArgs(popupUrl(base, id), rect, profile))) return 'edge';
    log('POPUP Edge not found: opening the editor in the browser');
  }
  browser(`${base}/?session=${encodeURIComponent(id)}`);
  return 'browser';
}
