// Engine side of the PowerPoint add-in. VBA sends a small text request (through AppleScriptTask on
// the Mac), this lays the chart out with the same engine as the CLI and answers in the line format
// (format.ts) that VBA turns into native shapes.
//
// Request, one field per line:
//   1  op       insert | update | relayout | edit | detach | direction | command | types | agenda (src/addin/agenda.ts)
//   2  arg      gallery item (insert; "<item>@<direction>" inserts in that direction), edit-session id
//               (update, optional), element role (detach), the direction (direction: up | down | left
//               | right) or chart command (command: core/commands.ts CHART_COMMANDS)
//   3  frame    "x,y,w,h" pt: target frame (insert) or the chart's OC_FRAME tag
//   4  bounds   "x,y,w,h" pt: the group's current bounds on the slide (update, relayout)
//   then, optional, one "key=value" per line:
//      theme=RRGGBB,… the slide's theme colours in ThemeColorScheme order (dk1 lt1 dk2 lt2
//                     accent1…accent6 hlink folHlink): theme slots resolve to these, so label
//                     contrast is decided on the colours PowerPoint will show
//      majorfont=F, minorfont=F  the deck's theme fonts (headings, body): +mj-lt / +mn-lt in the
//                     style or an element's override measure as these
//      scale=N        same-scale group: the smallest own scale (OC_SCALE, pt per unit) of the
//                     other charts in this chart's group on the slide
//      base=JSON      the chart's OC_BASE tag (update, relayout): its formatting as last drawn
//      native=JSON    what its shapes look like now: [{role, kind, props}] (core/overrides.ts
//                     RenderedEntry); with base, formatting changed in PowerPoint since is kept as
//                     overrides (mergeNativeEdits) instead of being drawn over
//      screen=…, window=…  (edit) where the chart is on screen, for the Edit data pop-up
//                     (src/addin/popup.ts ChartWhere); not part of the layout context
//      roles=a,b,…    (command) OC_ROLE values of the selected elements; none = the whole chart
//      group=NAME     (command scaleGroup) the same-scale group to join; empty = leave it
//   (edit sends the same lines 3–4 and params, so the server can lay an Apply out by itself:
//    layoutForApply)
//   5… spec     the chart's OC_SPEC tag, reassembled (update, relayout, edit); always starts with "{"
//
// update   re-lays the chart out where it now is, with the applied edit session if there is one
// relayout the same, but answers "N" (nothing to do) when neither the frame nor the data changed
// edit     opens the chart in the editor (an edit session, docs/ux-editor.md)
// direction lays the chart out again where it is, growing in the direction `arg` (spec.ts DIRECTIONS)
// detach   converts one label to plain text ("Convert to plain text"): its text as drawn now is
//          frozen in the spec (custom content, `detach: true`), so data edits no longer rewrite it;
//          it stays in the chart group and moves with the chart
// command  a ribbon / right-click command that changes the chart's options or overrides, not its
//          data (totals, cagr, difference, unpin, scaleGroup: core/commands.ts runCommand); laid
//          out like update, keeping formatting changed in PowerPoint (base= / native=)
import { createHash } from 'node:crypto';
import { defaultStyle, detachElement, InvalidSpecError, runCommand, kindOf, layoutChart, mergeNativeEdits, migrate, renderedOf, SPEC_VERSION, THEME_SLOTS, type ChartSpec, type RenderedEntry, type RenderedMap, type Frame, type LayoutHints, type Measurer, type Scene, type Style, type ThemeColors, type ThemeFonts } from '../core/index.ts';
import { chartContent } from '../deck.ts';
import { directionsOf, type Direction } from '../core/spec.ts';
import { shapeName } from '../render/pptx.ts';
import { parseFrame } from '../render/tags.ts';
import type { Session } from '../server.ts';
import { FORMAT_VERSION, num, record, shapeBounds, shapeRecord } from './format.ts';
import { galleryItems, sampleFor, typeLabel } from './registry.ts';
import { reinstallCommand } from './launcher.ts';
import { handleAgenda } from './agenda.ts';
import { parseWhere, type ChartWhere } from './popup.ts';

export interface AddinRequest {
  op: string; arg: string; frame?: Frame; bounds?: Frame; spec: string;
  /** The slide's theme colours (from `theme=`); they win over the spec's `style.theme`. */
  theme?: Partial<ThemeColors>;
  /** The deck's theme fonts (`majorfont=`, `minorfont=`): what +mj-lt / +mn-lt stand for. */
  themeFonts?: ThemeFonts;
  /** Same-scale group: the smallest own scale of the other charts in the group (from `scale=`). */
  ptPerUnit?: number;
  /** The chart's OC_BASE and its shapes as they are now (`base=`, `native=`): native edits to keep. */
  base?: RenderedMap;
  native?: RenderedEntry[];
  /**
   * Lines 3 up to the spec as sent (frame, bounds, theme=, scale=), each ending in a newline: what the
   * layout depends on besides the data. VBA compares it with the chart's own to reuse a layout.
   */
  context?: string;
  /** Edit data: where the chart is on screen (`screen=`, `window=`), to put the pop-up next to it. */
  where?: ChartWhere;
  /** Command: the selected elements' roles (`roles=`) and the same-scale group (`group=`). */
  roles?: string[];
  group?: string;
}

/** Context of one layout in PowerPoint: the slide's theme and the chart's same-scale group. */
interface Env {
  theme?: Partial<ThemeColors>; themeFonts?: ThemeFonts; hints: LayoutHints;
  /** True when a message should be shown now (a missing font is reported once): see handleRequest. */
  notify?: (message: string) => boolean;
}

/** ThemeColorScheme order (msoThemeColorSchemeIndex 1–12); the last two are not slots. */
const SCHEME_ORDER = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'] as const;

function parseTheme(v: string): Partial<ThemeColors> | undefined {
  const hex = v.split(',').map((x) => x.trim());
  const out: Partial<ThemeColors> = {};
  SCHEME_ORDER.forEach((k, i) => { if (/^[0-9A-Fa-f]{6}$/.test(hex[i] ?? '') && (THEME_SLOTS as readonly string[]).includes(k)) out[k] = hex[i].toUpperCase(); });
  return Object.keys(out).length ? out : undefined;
}

/** Edit sessions on the add-in's local server (src/addin/server.ts). */
export interface Sessions {
  get(id: string): Promise<Session | undefined>;
  /** Stores the chart as an open session and shows it in the editor (`where`: next to the chart). */
  open(id: string, chart: object, title: string, where?: ChartWhere): Promise<void>;
  /**
   * Marks an applied session as laid out (back to "open", chart kept; the add-in server sets it to
   * "drawn" once PowerPoint has finished drawing). Sessions are keyed by the chart's OC_SPEC, so
   * without this an edit would later be applied to any other chart that still has the same data (a
   * fresh sample of the same type, a copy) once the editor was closed.
   */
  settle?(id: string): Promise<void>;
}

export function parseRequest(text: string): AddinRequest {
  const lines = text.replace(/\r\n?/g, '\n').split('\n');
  let i = 4;
  const params: Record<string, string> = {};
  for (; i < lines.length && /^[a-z]+=/.test(lines[i]); i++) {
    const at = lines[i].indexOf('=');
    params[lines[i].slice(0, at)] = lines[i].slice(at + 1).trim();
  }
  const scale = Number(params.scale);
  // base= / native= describe the shapes, screen= / window= where the chart is on screen (Edit data
  // pop-up): neither is the layout's context.
  const context = lines.length > 2 ? lines.slice(2, i).filter((l) => !/^(base|native|screen|window|view|roles|group)=/.test(l)).map((l) => `${l}\n`).join('') : '';
  const json = (v?: string) => { try { return v ? JSON.parse(v) : undefined; } catch { return undefined; } };
  const base = json(params.base), native = json(params.native);
  const where = parseWhere(params);
  const themeFonts: ThemeFonts = {
    ...(params.majorfont ? { major: params.majorfont } : {}), ...(params.minorfont ? { minor: params.minorfont } : {}),
  };
  return {
    op: (lines[0] ?? '').trim(),
    arg: (lines[1] ?? '').trim(),
    frame: parseFrame(lines[2]?.trim()),
    bounds: parseFrame(lines[3]?.trim()),
    spec: lines.slice(i).join('\n').trim(),
    ...(params.theme ? { theme: parseTheme(params.theme) } : {}),
    ...(themeFonts.major || themeFonts.minor ? { themeFonts } : {}),
    ...(params.scale && scale > 0 && Number.isFinite(scale) ? { ptPerUnit: scale } : {}),
    ...(context ? { context } : {}),
    ...(base && typeof base === 'object' && !Array.isArray(base) && Array.isArray(native) ? { base, native } : {}),
    ...(where ? { where } : {}),
    ...(params.roles ? { roles: params.roles.split(',').filter(Boolean) } : {}),
    ...('group' in params ? { group: params.group } : {}),
  };
}

/**
 * Apply without a round trip through PowerPoint's AppleScriptTask (about 0.2 s each): the add-in
 * server lays the applied chart out itself and hands the response to ApplyEdit with the context it
 * assumed. The context of a session's chart comes from Edit data and then from each Apply's answer.
 */
const applyContexts = new Map<string, { context: string; spec: string }>();
const MAX_CONTEXTS = 64;

export function rememberContext(id: string, context: string, spec?: string) {
  const prev = applyContexts.get(id);
  applyContexts.delete(id);
  applyContexts.set(id, { context, spec: spec ?? prev?.spec ?? '' });
  if (applyContexts.size > MAX_CONTEXTS) applyContexts.delete(applyContexts.keys().next().value!);
}

export const contextOf = (id: string): string | undefined => applyContexts.get(id)?.context;

/**
 * The response ApplyEdit draws for applied session `s` on a chart whose context is `context`:
 * exactly the `update` request VBA would send for it. The session is not settled here: the server
 * does that once it hands the response to PowerPoint.
 */
export async function layoutForApply(s: Session, context: string, m: Measurer): Promise<string> {
  const req = parseRequest(`update\n${s.id}\n${context}`);
  const spec = applyContexts.get(s.id)?.spec || JSON.stringify(s.chart);
  const applied: Sessions = { get: async () => ({ ...s, state: 'applied' }), open: async () => {} };
  return handleRequest({ ...req, spec }, m, applied);
}

/**
 * A chart (its OC_SPEC object, style included) laid out for a chart context exactly as `update`
 * would draw it, and a function that lays a changed spec out the same way (drag handles,
 * src/addin/handles.ts: < 1 ms a layout once warm).
 */
export function sceneFor(chart: object, context: string, m: Measurer): { scene: Scene; frame: Frame; relayout: (spec: ChartSpec) => Scene } {
  const req = parseRequest(`update\n\n${context}`);
  const env: Env = { theme: req.theme, themeFonts: req.themeFonts, hints: req.ptPerUnit ? { ptPerUnit: req.ptPerUnit } : {} };
  const first = splitSpec(chart as ChartSpec & { style?: Partial<Style> });
  const frame = resolveFrame(first.chart, first.style, req.frame, req.bounds, m, env);
  const relayout = (spec: ChartSpec) => { const s = splitSpec(spec as ChartSpec & { style?: Partial<Style> }); return layout(s.chart, s.style, frame, m, env); };
  return { scene: layout(first.chart, first.style, frame, m, env), frame, relayout };
}

/** Same frame as a chart on a generated slide: the area below the title. */
export const DEFAULT_FRAME: Frame = { x: 48, y: 96, w: 864, h: 408 };

/** Edit-session id for a chart: derived from its OC_SPEC, so Update chart finds an applied edit. */
export const specKey = (spec: string) => createHash('sha1').update(spec).digest('hex').slice(0, 16);

const SELECT_FIRST = 'Select an open-charts chart first, then try again.';

export async function handleRequest(req: AddinRequest, m: Measurer, sessions?: Sessions, notify?: (message: string) => boolean): Promise<string> {
  // A missing font is reported when the user acts on a chart, not on the automatic re-layout.
  const tell = req.op === 'relayout' ? undefined : notify;
  const env: Env = { theme: req.theme, themeFonts: req.themeFonts, hints: req.ptPerUnit ? { ptPerUnit: req.ptPerUnit } : {}, notify: tell };
  try {
    switch (req.op) {
      case 'insert': {
        const sample = sampleFor(req.arg);
        if ('error' in sample) return error(sample.error);
        // A new chart takes the deck's body font (as PowerPoint text does): stored as the theme token,
        // so it follows the deck if its theme fonts change later. Without theme fonts: the default style.
        const style = req.themeFonts?.minor ? { font: '+mn-lt' } : undefined;
        return respond(sample.spec, style, req.frame ?? DEFAULT_FRAME, m, env);
      }
      case 'update':
      case 'relayout': {
        if (!req.spec) return error(SELECT_FIRST);
        const id = req.arg || specKey(req.spec);
        const session = await sessions?.get(id);
        // An Apply that did not reach the slide ("error") is still pending for an explicit Update
        // chart on that chart, but never for the automatic re-layout of another chart with the same data.
        const pending = session?.state === 'applied' || (session?.state === 'error' && req.op === 'update');
        const edited = pending && session!.chart && typeof session!.chart === 'object';
        const read = edited ? splitSpec(session!.chart as ChartSpec & { style?: Partial<Style> }) : readSpec(req.spec);
        const { style } = read;
        let { chart } = read;
        // Formatting changed with PowerPoint's own tools since the last layout becomes overrides.
        let kept = false;
        if (!edited && req.base && req.native) {
          const merged = mergeNativeEdits(chart, req.base, renderedOf(req.native));
          chart = merged.spec;
          kept = merged.changes.length > 0;
        }
        const frame = resolveFrame(chart, style, req.frame, req.bounds, m, env);
        if (req.op === 'relayout' && !edited && !kept && req.frame && sameFrame(frame, req.frame)) return `${header()}\n${record('N')}\n`;
        const out = respond(chart, style, frame, m, env);
        if (edited) await sessions?.settle?.(id);
        return out;
      }
      case 'edit': {
        if (!req.spec) return error(SELECT_FIRST);
        const { chart, style } = readSpec(req.spec); // fail early on a damaged tag
        if (!sessions) return error('Editing is not available here.');
        const id = specKey(req.spec);
        await sessions.open(id, style ? { ...chart, style } : chart, `${typeLabel(chart.type)} chart`, req.where);
        if (req.context) rememberContext(id, req.context, req.spec);
        return `${header()}\n${record('I', id)}\n`;
      }
      case 'detach': {
        if (!req.spec) return error(SELECT_FIRST);
        const role = req.arg;
        if (kindOf(role) !== 'label') return error('Select one label of the chart, then convert it to plain text.');
        const { chart, style } = readSpec(req.spec);
        const frame = resolveFrame(chart, style, req.frame, req.bounds, m, env);
        const before = layout(chart, style, frame, m, env);
        if (!before.roles?.includes(role)) return error('This label is not drawn any more; update the chart and try again.');
        const spec = detachElement(chart, before, role);
        if (!spec) return error('This label is not drawn any more; update the chart and try again.');
        return respond(spec, style, frame, m, env);
      }
      case 'command': {
        if (!req.spec) return error(SELECT_FIRST);
        const read = readSpec(req.spec);
        let chart = read.chart;
        if (req.base && req.native) chart = mergeNativeEdits(chart, req.base, renderedOf(req.native)).spec;
        const done = runCommand(chart, req.arg, req.roles ?? [], req.group ?? '');
        if ('error' in done) return error(done.error);
        // The frame as the chart is now (before the command): a command never moves the chart.
        const frame = resolveFrame(chart, read.style, req.frame, req.bounds, m, env);
        return respond(done.spec, read.style, frame, m, env);
      }
      case 'agenda': {
        // Agenda commands (src/addin/agenda.ts): frame = the slide size, bounds = the agenda
        // placeholder of the "open-charts agenda" layout, the rest a snapshot of the deck.
        const size = req.frame ?? { x: 0, y: 0, w: 960, h: 540 };
        return handleAgenda({ command: req.arg, slide: { w: size.w, h: size.h }, placeholder: req.bounds, theme: req.theme, snapshot: req.spec }, m);
      }
      case 'direction': {
        // The ribbon's direction buttons: the selected chart, laid out again in its place growing `arg`.
        if (!req.spec) return error(SELECT_FIRST);
        const { chart, style } = readSpec(req.spec);
        const dir = req.arg as Direction;
        const ok = directionsOf(chart.type);
        if (!ok.includes(dir)) {
          return error(ok.length > 1 ? `${typeLabel(chart.type)} charts can grow ${ok.join(', ')}, not ${dir}.` : `${typeLabel(chart.type)} charts have no direction to change.`);
        }
        const frame = resolveFrame(chart, style, req.frame, req.bounds, m, env);
        const { direction: _, ...rest } = chart.options ?? {};
        return respond({ ...chart, options: dir === 'up' ? rest : { ...rest, direction: dir } }, style, frame, m, env);
      }
      case 'types':
        return [header(), ...galleryItems().map((g) => record('M', g.id, 'spec' in sampleFor(g.id) ? 'yes' : 'no'))].join('\n') + '\n';
      default:
        return error(`This add-in and its engine do not match ("${req.op}"). In Terminal, run: ${reinstallCommand()}`);
    }
  } catch (err) {
    // A chart whose stored data was edited into something the engine cannot lay out: say where.
    if (err instanceof InvalidSpecError) return error(`This chart's data (OC_SPEC tag) is not valid, so it cannot be laid out:\n${err.errors.map((x) => `${x.path || '/'}: ${x.message}`).join('\n')}`);
    return error((err as Error).message);
  }
}

const sameFrame = (a: Frame, b: Frame) =>
  Math.abs(a.x - b.x) < 0.5 && Math.abs(a.y - b.y) < 0.5 && Math.abs(a.w - b.w) < 0.5 && Math.abs(a.h - b.h) < 0.5;

const header = () => record('OC', FORMAT_VERSION);
const error = (msg: string) => `${header()}\n${record('E', msg)}\n`;

const DAMAGED = 'This chart\'s data (OC_SPEC tag) is damaged and cannot be read.';
export const NEWER_SPEC = 'This chart was made with a newer version of open-charts. Install the latest open-charts to change it; until then it stays as it is.';

/**
 * The chart in an OC_SPEC tag, upgraded to the current spec version (migrate): a chart from an older
 * version is laid out, edited and stored again as a current one. Throws a plain message for a tag
 * that is not a chart object or that a newer open-charts wrote.
 */
function readSpec(json: string): { chart: ChartSpec; style?: Partial<Style> } {
  let parsed: unknown;
  try {
    parsed = JSON.parse(json);
  } catch {
    throw new Error(DAMAGED);
  }
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed) || typeof (parsed as { type?: unknown }).type !== 'string') throw new Error(DAMAGED);
  const v = (parsed as { specVersion?: unknown }).specVersion;
  if (typeof v === 'number' && v > SPEC_VERSION) throw new Error(NEWER_SPEC);
  return splitSpec(migrate(parsed as ChartSpec & { style?: Partial<Style> }));
}

function splitSpec(parsed: ChartSpec & { style?: Partial<Style> }): { chart: ChartSpec; style?: Partial<Style> } {
  const { style, ...chart } = parsed;
  return { chart, style: style && Object.keys(style).length ? style : undefined };
}

function layout(chart: ChartSpec, style: Partial<Style> | undefined, frame: Frame, m: Measurer, env: Env = { hints: {} }): Scene {
  // Theme slots resolve to the slide's own theme colours: PowerPoint draws them in those.
  const theme = env.theme ? { ...style?.theme, ...env.theme } : style?.theme;
  // Theme fonts: the deck's, as PowerPoint will draw them (not stored with the chart).
  const themeFonts = env.themeFonts ? { ...style?.themeFonts, ...env.themeFonts } : style?.themeFonts;
  return layoutChart(chart, frame, m, { ...defaultStyle, ...style, ...(theme ? { theme } : {}), ...(themeFonts ? { themeFonts } : {}) }, env.hints);
}

/** Bounds of the chart group as drawn by the add-in: the shapes plus its invisible frame rectangle. */
export function sceneBounds(scene: Scene): Frame | undefined {
  const f = scene.frame;
  let x0 = f.x, y0 = f.y, x1 = f.x + f.w, y1 = f.y + f.h;
  for (const s of scene.shapes) {
    const b = shapeBounds(s);
    x0 = Math.min(x0, b.x0); y0 = Math.min(y0, b.y0); x1 = Math.max(x1, b.x1); y1 = Math.max(y1, b.y1);
  }
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
}

/**
 * The frame to re-lay the chart out in. VBA reports the group's bounds, which hug the shapes, not
 * the frame; so lay out in the stored frame, and if the user has since moved or resized the group,
 * map the frame by the same move/scale that takes the laid-out bounds to the current ones.
 */
export function resolveFrame(chart: ChartSpec, style: Partial<Style> | undefined, stored: Frame | undefined, current: Frame | undefined, m: Measurer, env?: Env): Frame {
  const f0 = stored ?? current ?? DEFAULT_FRAME;
  if (!current) return f0;
  const b0 = sceneBounds(layout(chart, style, f0, m, env));
  if (!b0 || b0.w < 1 || b0.h < 1) return f0;
  const close = (a: number, b: number) => Math.abs(a - b) < 0.75;
  if (close(b0.x, current.x) && close(b0.y, current.y) && close(b0.w, current.w) && close(b0.h, current.h)) return f0;
  const sx = current.w / b0.w, sy = current.h / b0.h;
  return { x: current.x + (f0.x - b0.x) * sx, y: current.y + (f0.y - b0.y) * sy, w: f0.w * sx, h: f0.h * sy };
}

function respond(chart: ChartSpec, style: Partial<Style> | undefined, frame: Frame, m: Measurer, env?: Env): string {
  const scene = layout(chart, style, frame, m, env);
  const font = style?.font ?? defaultStyle.font;
  const content = chartContent(chart, scene, font, style);
  const lines = [header(), record('G', content.name, content.altText), record('F', [frame.x, frame.y, frame.w, frame.h].map(num).join(','))];
  // A font that is not installed here: tell the user (once on this computer: the caller's notify).
  for (const w of scene.warnings ?? []) if (w.startsWith('font "') && env?.notify?.(w)) lines.push(record('M', w));
  for (const [k, v] of content.tags ?? []) lines.push(record('T', k, v));
  scene.shapes.forEach((s, i) => {
    const role = scene.roles?.[i] ?? '';
    lines.push(shapeRecord(s, font, role, shapeName(s, i + 1, role || undefined)));
  });
  return lines.join('\n') + '\n';
}
