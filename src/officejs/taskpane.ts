// Task pane entry of the Office.js add-in (bundled to dist/officejs/taskpane.js, loaded by
// taskpane.html before the datasheet editor's own script). It gives the editor a host
// (`window.OC_HOST`, used in `?host=office` mode instead of the local server's /api/render and edit
// sessions) and binds the pane to PowerPoint:
//   nothing / another shape selected → insert mode: Insert draws the chart on the current slide
//   an open-charts chart selected    → edit mode: the datasheet shows its OC_SPEC; Apply replaces it
//                                      in place (position, size, z-order); Revert undoes the last Apply
//   the chart was resized            → banner: "Re-layout to new size"
// Path B (native shapes in a group) where the PowerPoint can draw the scene, else Path A (the
// writer's slide via insertSlidesFromBase64), see plan.ts.
import type { ChartSpec, Frame, Measurer, ThemeColors } from '../core/index.ts';
import type { DeckSpec } from '../deck.ts';
import { parseFrame, specJson } from '../render/tags.ts';
import { ENGINE_VERSION } from '../version.ts';
import { buildChart, chartDeck, CHART_FRAME, fromBase64, renderDeck, resolveFrame, toBase64, type StoredSpec } from './engine.ts';
import { loadMeasurer } from './measure.ts';
import { hostCaps, NO_CAPS, planScene, type HostCaps } from './plan.ts';
import { drawPathB, insertSlide, readSelection, replaceSlide, REPLACE_MARKER, slideTheme, type SelectedChart } from './powerpoint.ts';
import { browserInflate, patchSlide } from './slide-patch.ts';

/* eslint-disable @typescript-eslint/no-explicit-any */
declare const Office: any;
const w = window as any;

// ---- Log: to the console, an in-page list (tests) and the dev server's /api/log (live checks) ----
const LOG: string[] = (w.OC_LOG = []);
export function log(event: string, data: Record<string, unknown> = {}) {
  const line = `${event} ${JSON.stringify(data)}`;
  LOG.push(line);
  console.log(`[open-charts] ${line}`);
  try { void fetch(new URL('api/log', location.href), { method: 'POST', body: line, keepalive: true }).catch(() => {}); } catch { /* no server: fine */ }
}

// ---- State --------------------------------------------------------------------------------------
let caps: HostCaps = NO_CAPS;
let officeReady = false;
const measurer: Promise<Measurer> = loadMeasurer().catch((e) => { log('fonts-error', { message: String(e) }); throw e; });

interface Bound { chart: SelectedChart; spec: StoredSpec; frame: Frame; resized: boolean }
let bound: Bound | null = null;
let slideId: string | undefined;
let theme: Partial<ThemeColors> | undefined;
/** Previous spec + frame of the bound chart before each Apply, for Revert. */
let history: { spec: StoredSpec; frame: Frame }[] = [];
let busy = false;
/** Set while the pane re-binds to the chart it just drew (keeps history and the editor's undo). */
let rebinding = false;

const editor = () => w.OC_EDITOR as undefined | {
  load(chart: StoredSpec, info: { title: string }): void;
  setOffice(s: { mode: 'insert' | 'edit' | 'unavailable'; title?: string; resized?: boolean; canRevert?: boolean; note?: string }): void;
};

function publish(note?: string) {
  const e = editor();
  if (!e) return;
  if (!officeReady) e.setOffice({ mode: 'unavailable', note: note ?? 'Open this pane from PowerPoint (Home ▸ open-charts) to insert charts.' });
  else if (bound) e.setOffice({ mode: 'edit', title: bound.chart.name, resized: bound.resized, canRevert: history.length > 0, note });
  else e.setOffice({ mode: 'insert', note });
}

// ---- Selection → pane ----------------------------------------------------------------------------
async function refresh(force = false) {
  if (!officeReady || busy) return;
  try {
    const sel = await readSelection(caps);
    slideId = sel.kind === 'chart' ? sel.chart.slideId : sel.slideId ?? slideId;
    if (sel.kind !== 'chart') {
      if (bound) { bound = null; history = []; log('unbound'); }
      publish();
      return;
    }
    const json = specJson(sel.chart.tags);
    if (!json) return;
    let spec: StoredSpec;
    try { spec = JSON.parse(json); } catch { publish('This chart\'s data (OC_SPEC tag) is damaged and cannot be read.'); return; }
    theme = await slideTheme(caps, sel.chart.slideId);
    const m = await measurer;
    const { frame, resized } = resolveFrame(spec, parseFrame(sel.chart.tags.OC_FRAME), sel.chart.bounds, m, theme);
    const same = bound?.chart.id === sel.chart.id || rebinding;
    if (!same) history = [];
    bound = { chart: sel.chart, spec, frame, resized };
    if (!same || (force && !rebinding)) editor()?.load(structuredClone(spec), { title: sel.chart.name });
    log('bound', { id: sel.chart.id, type: spec.type, resized, engine: sel.chart.tags.OC_ENGINE });
    const older = sel.chart.tags.OC_ENGINE && sel.chart.tags.OC_ENGINE !== ENGINE_VERSION;
    publish(older ? `Made with open-charts ${sel.chart.tags.OC_ENGINE}; Apply re-lays it out with ${ENGINE_VERSION} (labels may move).` : undefined);
  } catch (e) {
    log('selection-error', { message: String((e as Error).message ?? e) });
  }
}

// ---- Drawing ---------------------------------------------------------------------------------------
async function draw(spec: StoredSpec, frame: Frame, replace?: SelectedChart): Promise<string> {
  const t0 = performance.now();
  const m = await measurer;
  const target = replace?.slideId ?? slideId;
  theme = await slideTheme(caps, target);
  const built = buildChart(spec, frame, m, theme);
  const plan = planScene(built.scene, caps, built.font);
  const shapes = built.scene.shapes.length;
  if (plan.path === 'B') {
    await drawPathB(caps, { ops: plan.ops, frame: built.scene.frame, name: built.content.name, altText: built.content.altText, tags: built.content.tags ?? [] }, { slideId: target, replaceId: replace?.id });
    log(replace ? 'apply' : 'insert', { path: 'B', type: spec.type, shapes, ops: plan.ops.length, ms: Math.round(performance.now() - t0) });
    return 'Drawn on the slide';
  }
  if (caps.groups && target) {
    // Path A on this slide: export it, put the chart into its XML, put it back (1.8).
    const frameTag = (built.content.tags ?? []).find(([k]) => k === 'OC_FRAME')?.[1] ?? '';
    await replaceSlide(target, replace && { id: replace.id, name: replace.name }, { name: built.content.name, frameTag }, async (b64) =>
      toBase64(await patchSlide(fromBase64(b64), built.content, { font: built.font, ...(replace ? { replaceName: REPLACE_MARKER } : {}) }, browserInflate)));
    log(replace ? 'apply' : 'insert', { path: 'A-inplace', reason: plan.reason, type: spec.type, shapes, ms: Math.round(performance.now() - t0) });
    return 'Drawn on the slide (the slide was re-inserted, so PowerPoint can draw every shape)';
  }
  if (!replace && caps.insertSlides) {
    await insertSlide(toBase64(chartDeck(built, spec, theme)), target);
    log('insert', { path: 'A-newslide', reason: plan.reason, type: spec.type, shapes, ms: Math.round(performance.now() - t0) });
    return `Inserted on a new slide: ${plan.reason}`;
  }
  throw new Error(`This PowerPoint cannot update the chart in place: ${plan.reason}. Update PowerPoint, or insert the chart again.`);
}

/** Binds the pane to the chart just drawn (selected by the drawing step). */
async function rebind() {
  rebinding = true;
  try { await refresh(true); } finally { rebinding = false; }
}

async function run<T>(what: string, f: () => Promise<T>): Promise<T> {
  if (busy) throw new Error('Busy — try again in a moment.');
  busy = true;
  try {
    return await f();
  } catch (e) {
    log(`${what}-error`, { message: String((e as Error).message ?? e) });
    throw e;
  } finally {
    busy = false;
  }
}

// ---- The host the editor talks to (src/web/editor.html, ?host=office) -------------------------------
w.OC_HOST = {
  async render(payload: DeckSpec & { slides: { title?: string; chart: ChartSpec }[] }) {
    const m = await measurer;
    const frame = bound?.frame ?? CHART_FRAME;
    const deck: DeckSpec = {
      style: payload.style,
      slides: payload.slides.map((s) => {
        const chart = theme ? { ...s.chart, style: { ...s.chart.style, theme: { ...s.chart.style?.theme, ...theme } } } : s.chart;
        return { title: s.title, charts: [{ frame, chart }] };
      }),
    };
    return renderDeck(deck, m);
  },
  async apply(chart: StoredSpec): Promise<{ message: string }> {
    if (!officeReady) throw new Error('Not running in PowerPoint.');
    const message = await run('apply', async () => {
      if (bound) {
        history.push({ spec: bound.spec, frame: bound.frame });
        return draw(chart, bound.frame, bound.chart);
      }
      return draw(chart, CHART_FRAME);
    });
    await rebind();
    return { message };
  },
  async revert(): Promise<{ message: string }> {
    const prev = history.at(-1);
    if (!bound || !prev) throw new Error('Nothing to revert.');
    const message = await run('revert', async () => {
      await draw(prev.spec, prev.frame, bound!.chart);
      history.pop();
      return 'Reverted to the chart before the last Apply';
    });
    await rebind();
    return { message };
  },
  /** Re-lays the chart out at its new size with its stored data. */
  async relayout(): Promise<{ message: string }> {
    if (!bound) throw new Error('Select an open-charts chart first.');
    return w.OC_HOST.apply(bound.spec);
  },
  refresh: () => refresh(true),
  caps: () => caps,
  log: () => [...LOG],
};

// ---- Office ---------------------------------------------------------------------------------------
function start() {
  if (typeof Office === 'undefined') { log('no-office'); publish(); return; }
  Office.onReady((info: { host?: string; platform?: string }) => {
    if (info.host !== Office.HostType?.PowerPoint && info.host !== 'PowerPoint') { log('not-powerpoint', { host: info.host }); publish(); return; }
    const isSet = (v: string) => !!Office.context.requirements.isSetSupported('PowerPointApi', v);
    caps = hostCaps(isSet);
    officeReady = true;
    const matrix = Object.fromEntries(['1.1', '1.2', '1.3', '1.4', '1.5', '1.6', '1.7', '1.8', '1.9', '1.10'].map((v) => [v, isSet(v)]));
    log('ready', { platform: info.platform, engine: ENGINE_VERSION, PowerPointApi: matrix });
    try {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => { void refresh(); });
    } catch (e) { log('selection-handler-error', { message: String(e) }); }
    publish();
    void refresh(true);
  });
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
