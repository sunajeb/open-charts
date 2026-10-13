// The add-in's local server (127.0.0.1:4174): the same `serve()` as `open-charts serve` (editor,
// render API, edit sessions) plus
//   POST /api/addin/run    the warm engine: a request in, the line format out (no node start-up)
//   GET  /api/addin/ping   liveness
//   POST /api/addin/quit   stop (install restarts it with the new engine)
// When the editor applies a session (docs/ux-editor.md), it tells PowerPoint to re-lay out the
// chart: VBA on the Mac cannot poll, so the server runs the add-in's ApplyEdit macro through its
// long-lived AppleScript helper (runner.ts; on Windows a PowerShell helper calling it through COM,
// win.ts) and records the outcome on the session ("drawn" or
// "error" with a message) for the editor to show.
//
// The server belongs to the PowerPoint that started it (the bridge spawns it): macOS attributes the
// server, and so every osascript it runs, to that PowerPoint process. Once that PowerPoint has quit,
// an Apple event from the server to a new PowerPoint is held by PowerPoint's own permission check and
// freezes it (no reply, ApplyEdit never runs). So a server that outlived its PowerPoint never sends
// one: it answers the new PowerPoint's ping and engine calls with 503 and exits, and the bridge
// starts a fresh server from the new PowerPoint.
import { execFileSync, spawn } from 'node:child_process';
import { appendFileSync, existsSync, mkdirSync, readFileSync } from 'node:fs';
import { addinHome, reinstallCommand } from './launcher.ts';
import { join } from 'node:path';
import type { IncomingMessage, ServerResponse } from 'node:http';
import { getSession, serve, setSession, type Session } from '../server.ts';
import { createMeasurer } from '../fonts.ts';
import type { Measurer } from '../core/index.ts';
import { contextOf, handleRequest, layoutForApply, parseRequest, rememberContext, type Sessions } from './engine.ts';
import { macroRunner, type MacroResult } from './runner.ts';
import { openUrl, powershellRunner, selfCommand } from './win.ts';
import { helperPath, openPopup, popupHelper, type ChartWhere } from './popup.ts';
import { toolbarRoutes } from './toolbar.ts';
import { handlesController } from './handles.ts';

export const ADDIN_PORT = 4174;
const BASE = `http://127.0.0.1:${ADDIN_PORT}`;
const CLI = new URL('../cli.ts', import.meta.url);
const EDITOR = new URL('../web/editor.html', import.meta.url);

export { addinHome };

export function log(line: string, home = addinHome()) {
  try {
    mkdirSync(home, { recursive: true });
    const d = new Date(), local = new Date(d.getTime() - d.getTimezoneOffset() * 60_000).toISOString().slice(0, 19);
    appendFileSync(join(home, 'addin.log'), `${local} ${line.replace(/[\r\n\t]+/g, ' ')}\n`);
  } catch { /* logging must never break the add-in */ }
}

const ping = (base = BASE) => fetch(`${base}/api/addin/ping`, { signal: AbortSignal.timeout(500) }).then((r) => r.ok, () => false);

/** Starts the server in the background unless it is already up. */
export async function ensureServer(base = BASE): Promise<boolean> {
  if (await ping(base)) return true;
  const [command, args] = selfCommand(CLI, ['addin', 'serve']);
  spawn(command, args, { detached: true, stdio: 'ignore', windowsHide: true }).unref();
  for (let i = 0; i < 60; i++) {
    await new Promise((r) => setTimeout(r, 100));
    if (await ping(base)) return true;
  }
  return false;
}

/**
 * Edit sessions over HTTP (works from the server itself and from a one-off `addin run`). `open` shows
 * the session: the default browser here; the add-in server passes the pop-up (src/addin/popup.ts).
 */
export type OpenEditor = (url: string, id: string, where?: ChartWhere) => void;
export function httpSessions(base = BASE, open: OpenEditor = (url) => openUrl(url)): Sessions {
  return {
    async get(id) {
      try {
        const r = await fetch(`${base}/api/session/${id}`, { signal: AbortSignal.timeout(2000) });
        return r.ok ? ((await r.json()) as Session) : undefined;
      } catch { return undefined; }
    },
    async open(id, chart, title, where) {
      if (!(await ensureServer(base))) throw new Error('The open-charts editor could not be started.');
      const r = await fetch(`${base}/api/session/${id}`, { method: 'POST', body: JSON.stringify({ chart, title, state: 'open' }) });
      if (!r.ok) throw new Error(`The editor refused the chart: ${await r.text()}`);
      open(`http://localhost:${ADDIN_PORT}/?session=${id}`, id, where);
    },
    async settle(id) {
      try {
        await fetch(`${base}/api/session/${id}`, { method: 'POST', body: JSON.stringify({ state: 'open' }), signal: AbortSignal.timeout(2000) });
      } catch { /* the chart is drawn either way */ }
    },
  };
}

/**
 * The editor in src/web supports `?session=<id>` once it fetches /api/session itself. Until then
 * (older editor), the page is preloaded with the chart and posts each change back as an Apply.
 */
export function editorPage(id: string, chart: unknown): string | undefined {
  const html = readFileSync(EDITOR, 'utf8');
  if (html.includes('/api/session/')) return undefined;
  const { style, ...spec } = chart as { style?: unknown };
  const deck = style ? { style, slides: [{ chart: spec }] } : { slides: [{ chart: spec }] };
  const store = /const STORE = '([^']+)'/.exec(html)?.[1] ?? 'open-charts-deck-v1';
  const boot = `<script>(function () {
  var id = ${JSON.stringify(id)}, store = ${JSON.stringify(store)}, last = null, timer = 0;
  try { localStorage.setItem(store, ${JSON.stringify(JSON.stringify(deck)).replace(/</g, '\\u003c')}); } catch (e) {}
  setInterval(function () {
    var d = null;
    try { d = localStorage.getItem(store); } catch (e) {}
    if (!d || d === last) return;
    var first = last === null;
    last = d;
    if (first) return;
    clearTimeout(timer);
    timer = setTimeout(function () {
      var deck = JSON.parse(d), chart = deck.slides && deck.slides[0] && deck.slides[0].chart;
      if (deck.style) chart.style = deck.style;
      fetch('/api/session/' + id, { method: 'POST', body: JSON.stringify({ chart: chart, state: 'applied' }) });
    }, 600);
  }, 300);
})();</script>`;
  const at = html.indexOf('<script');
  return at < 0 ? html + boot : html.slice(0, at) + boot + html.slice(at);
}

/** PowerPoint's process id while it is running (pgrep: no Apple event, no permission prompt). */
export function powerPointPid(): number | undefined {
  if (process.platform === 'win32') return undefined; // no Apple events on Windows: nothing to outlive
  try {
    return Number(execFileSync('pgrep', ['-x', 'Microsoft PowerPoint'], { encoding: 'utf8' }).trim().split('\n')[0]) || undefined;
  } catch { return undefined; }
}

const alive = (pid: number) => {
  try { process.kill(pid, 0); return true; } catch (err) { return (err as NodeJS.ErrnoException).code === 'EPERM'; }
};

/**
 * True when the PowerPoint that started the server (`owner`) has quit and another PowerPoint is
 * running: the server must not send that one Apple events (see the top of this file).
 */
export const outlivedPowerPoint = (owner: number | undefined, running = powerPointPid): boolean =>
  owner !== undefined && !alive(owner) && running() !== undefined;

export const RESTARTED_MESSAGE = 'PowerPoint was restarted after this editor window opened. Close this window, select the chart and click Edit data again.';

/** How long an Apply may take in PowerPoint before the editor is told it did not get through. */
export const APPLY_TIMEOUT_MS = 20_000;

/**
 * ApplyEdit's answer: log lines (SYNC …, APPLIED …), and the chart's context ("CONTEXT", tab for
 * newline): after drawing, the context of the new chart; without an APPLIED line, the chart's
 * context did not match the one the response was laid out for, so lay it out again for this one.
 */
export function parseApplyAnswer(out: string): { lines: string[]; context?: string; applied?: string; timer?: [number, number] } {
  const lines: string[] = [];
  let context: string | undefined, applied: string | undefined, timer: [number, number] | undefined;
  for (const line of out.replace(/\r/g, '\n').split('\n')) {
    if (line.startsWith('CONTEXT\t')) context = line.slice(8).replace(/\t/g, '\n');
    else if (line.startsWith('TIMER\t')) {
      const [a, b] = line.slice(6).split('\t').map(Number);
      if (Number.isFinite(a) && Number.isFinite(b)) timer = [a, b];
    } else if (line.trim()) {
      lines.push(line.trim());
      if (line.startsWith('APPLIED ')) applied = line.trim();
    }
  }
  return { lines, context, applied, ...(timer ? { timer } : {}) };
}

/** Seconds since local midnight, as VBA's Timer counts. */
const sinceMidnight = (ms: number) => { const d = new Date(ms); return (ms - new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime()) / 1000; };

/** What the user should be told when `run VB macro` failed (the Apple event's error, or a timeout). */
export function applyErrorMessage(err: { message: string; killed?: boolean }, platform = process.platform): string {
  const text = err.message;
  if (platform === 'win32' && (err.killed || /timed out/i.test(text))) {
    return 'PowerPoint did not respond. Close any open dialog in PowerPoint (and leave text editing), then Apply again.';
  }
  if (err.killed || /-1712\b|timed out/i.test(text)) {
    return 'PowerPoint did not respond. If macOS is asking whether PowerPoint may control PowerPoint, click Allow in that dialog, then Apply again.';
  }
  if (/-1743\b|not authori[sz]ed/i.test(text)) {
    return 'macOS does not allow PowerPoint to be controlled (Don\'t Allow was chosen). Turn it on in System Settings › Privacy & Security › Automation › Microsoft PowerPoint, then Apply again.';
  }
  if (/-600\b|not running/i.test(text)) return 'PowerPoint isn\'t running. Open the presentation, click the chart once, then Apply again.';
  const why = /execution error: (.*)$/m.exec(text)?.[1] ?? text.split('\n').filter(Boolean).at(-1) ?? text;
  return `PowerPoint could not update the chart (${why.trim()}). Click the chart once, then Apply again.`;
}

/** Runs a macro of the add-in in PowerPoint (`run VB macro`): its return value, or the error. */
export type Run = (macro: string, args: string[], timeoutMs: number) => Promise<MacroResult>;
// Windows: PowerShell + COM (win.ts); Mac: JavaScript for Automation.
const helper = process.platform === 'win32' ? powershellRunner() : macroRunner();
const helperRun: Run = (macro, args, timeoutMs) => helper.run(macro, args, timeoutMs);

/** Lays an applied session's chart out for a chart context (layoutForApply with the server's measurer). */
export type Prepare = (s: Session, context: string) => Promise<string>;

/**
 * Runs the add-in's ApplyEdit macro for an applied session and records the outcome. The server lays
 * the chart out itself for the chart's last known context (`prepare`) and hands the response to the
 * macro, which only draws: no engine round trip from PowerPoint. As it does, the session goes back
 * to "open" (rev + 1, as the engine's settle would), so the edit is never applied to another chart
 * with the same data; once the macro has drawn, that becomes "drawn". If the chart's context changed
 * (moved, resized, another theme), the macro answers with its context and is run again with a
 * layout for it. If the macro failed, timed out or found no chart, the session becomes "error" with
 * what to do. A newer Apply or a Cancel meanwhile is left alone.
 */
export async function applyInPowerPoint(s: Session, home = addinHome(), run: Run = helperRun, outlived = () => false, prepare?: Prepare): Promise<void> {
  const ours = () => {
    const now = getSession(s.id);
    return now && ((now.state === 'applied' && now.rev === s.rev) || (now.state === 'open' && now.rev === s.rev + 1)) ? now : undefined;
  };
  const fail = (message: string) => { if (ours()) setSession(s.id, { state: 'error', message }); };
  let prefix = '';
  try { prefix = JSON.parse(readFileSync(join(home, 'addin.json'), 'utf8')).prefix; } catch { /* not installed */ }
  if (!prefix) {
    log(`APPLY ${s.id} skipped: add-in not installed`, home);
    return fail(`The open-charts add-in is not installed. In Terminal, run: ${reinstallCommand(home)}`);
  }
  if (outlived()) {
    log(`APPLY ${s.id} skipped: this server was started by a PowerPoint that has quit`, home);
    return fail(RESTARTED_MESSAGE);
  }
  const t = Date.now();
  const macro = `${prefix}_ApplyEdit`;
  let layoutMs = 0;
  const lay = async (context: string | undefined) => {
    if (context === undefined || !prepare) return '';
    const t0 = Date.now();
    try { return await prepare(s, context); } catch { return ''; } finally { layoutMs += Date.now() - t0; }
  };
  const settle = () => { if (ours()?.state === 'applied') setSession(s.id, { state: 'open' }); };
  let context = contextOf(s.id);
  let response = await lay(context);
  if (response) settle();
  let sent = Date.now();
  let result = await run(macro, [s.id, context ?? '', response], APPLY_TIMEOUT_MS);
  let answer = 'out' in result ? parseApplyAnswer(result.out) : undefined;
  if (answer && !answer.applied && answer.context !== undefined && prepare) {
    // The chart is not where (or as) it was when the response was laid out: lay it out for now.
    context = answer.context;
    response = await lay(context);
    if (response) settle();
    sent = Date.now();
    result = await run(macro, [s.id, '*', response], APPLY_TIMEOUT_MS);
    answer = 'out' in result ? parseApplyAnswer(result.out) : undefined;
  }
  const back = Date.now(), ms = back - t;
  // Where the time went: until the macro started (the Apple event), in it, and back.
  const [t0, t1] = answer?.timer ?? [];
  const parts = t0 !== undefined && t1 !== undefined && t1 >= t0
    ? `, to macro ${Math.round((t0 - sinceMidnight(sent)) * 1000)} ms, macro ${Math.round((t1 - t0) * 1000)} ms, back ${Math.round((sinceMidnight(back) - t1) * 1000)} ms`
    : '';
  if ('error' in result) {
    log(`APPLY ${s.id} failed after ${ms} ms: ${result.error.killed ? 'timed out' : result.error.message}`, home);
    return fail(applyErrorMessage(result.error));
  }
  for (const line of answer?.lines ?? []) log(line, home);
  if (answer?.applied && answer.context !== undefined) rememberContext(s.id, answer.context);
  const now = ours();
  // An older add-in (no answer) settles through the engine instead: "open" means it drew.
  const drew = answer?.applied ? / OK\b/.test(answer.applied) : now?.state === 'open';
  log(`APPLY ${s.id} ${drew ? 'done' : 'ran, but no chart was updated'} (${ms} ms${prepare ? `, layout ${layoutMs} ms` : ''}${parts})`, home);
  if (drew && now) setSession(s.id, { state: 'drawn' });
  else fail('PowerPoint could not find the chart. Click it once on the slide, then Apply again (or click Update chart).');
}

/**
 * The selection as NotifySelection reports it, asked now (the add-in's SelectionInfo macro): the drag
 * handles re-measure the chart on screen after a scroll, zoom or window move. Undefined when the
 * add-in is not installed, PowerPoint is not the server's, or the macro fails.
 */
export async function selectionNow(outlived = () => false, home = addinHome(), run: Run = helperRun): Promise<string | undefined> {
  let prefix = '';
  try { prefix = JSON.parse(readFileSync(join(home, 'addin.json'), 'utf8')).prefix; } catch { return undefined; }
  if (!prefix || outlived()) return undefined;
  const r = await run(`${prefix}_SelectionInfo`, [], 5000);
  return 'out' in r && r.out.startsWith('select\n') ? r.out : undefined;
}

/** `sessions` that add up the time spent in their calls (the engine timing in the log). */
function timedSessions(sessions: Sessions): Sessions & { ms: number } {
  const out = { ms: 0 } as Sessions & { ms: number };
  const time = <A extends unknown[], R>(f: (...a: A) => Promise<R>) => async (...a: A) => {
    const t = performance.now();
    try { return await f(...a); } finally { out.ms += performance.now() - t; }
  };
  out.get = time(sessions.get.bind(sessions));
  out.open = time(sessions.open.bind(sessions));
  if (sessions.settle) out.settle = time(sessions.settle.bind(sessions));
  return out;
}

/**
 * The add-in's own routes. `outlived` (see outlivedPowerPoint) retires the server: the new
 * PowerPoint's ping and engine calls get 503, so the bridge starts a fresh server, and `retire` stops
 * this one.
 */
export function addinRoutes(m: () => Measurer, sessions: Sessions, outlived = () => false, retire = () => process.exit(0), onEdit = () => {}) {
  return async (req: IncomingMessage, res: ServerResponse): Promise<boolean> => {
    const url = req.url ?? '';
    if ((url === '/api/addin/ping' || url === '/api/addin/run') && outlived()) {
      res.writeHead(503, { connection: 'close' }).end('open-charts: restarting for the new PowerPoint');
      setImmediate(retire);
      return true;
    }
    if (req.method === 'GET' && url === '/api/addin/ping') {
      res.writeHead(200).end('open-charts');
      return true;
    }
    if (req.method === 'POST' && url === '/api/addin/run') {
      const t0 = performance.now();
      const chunks: Buffer[] = [];
      for await (const c of req) chunks.push(c as Buffer);
      const request = parseRequest(Buffer.concat(chunks).toString('utf8'));
      const timed = timedSessions(sessions);
      const out = await handleRequest(request, m(), timed);
      res.writeHead(200, { 'content-type': 'text/plain; charset=utf-8' }).end(out);
      const ms = performance.now() - t0;
      if (request.op !== 'types') log(`ENGINE ${request.op} ${Math.round(ms)} ms (sessions ${Math.round(timed.ms)} ms, layout ${Math.round(ms - timed.ms)} ms)`);
      if (request.op === 'edit') onEdit();
      return true;
    }
    if (req.method === 'POST' && url === '/api/addin/quit') {
      res.writeHead(204).end();
      setImmediate(() => process.exit(0));
      return true;
    }
    const s = /^\/\?session=([0-9A-Za-z_-]+)$/.exec(url);
    if (req.method === 'GET' && s) {
      const session = await sessions.get(s[1]);
      const page = session && editorPage(s[1], session.chart);
      if (!page) return false; // the editor handles sessions itself (or shows "not found")
      res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' }).end(page);
      return true;
    }
    return false;
  };
}

/** `open-charts addin serve`: the add-in's server. Exits after an idle hour. */
export function serveAddin(port = ADDIN_PORT) {
  let measurer: Measurer | undefined;
  const m = () => (measurer ??= createMeasurer());
  m();
  let timer: NodeJS.Timeout | undefined;
  const idle = () => { clearTimeout(timer); timer = setTimeout(() => process.exit(0), 60 * 60 * 1000); };
  idle();
  const owner = powerPointPid();
  const outlived = () => outlivedPowerPoint(owner);
  log(`SERVER started (pid ${process.pid}, PowerPoint ${owner ?? 'not running'})`);
  helper.warm(); // starts the AppleScript helper (no Apple event yet), so the first Apply does not wait for it
  const retire = () => { log(`SERVER retired: PowerPoint ${owner} has quit, a new one is running`); process.exit(0); };
  // Edit data: the user is about to Apply, so open the helper's connection to PowerPoint now (one
  // read-only Apple event; the first one costs up to seconds while the editor's browser starts).
  const connect = () => {
    if (outlived()) return;
    const t = Date.now();
    void helper.connect(10_000).then((r) => log(`HELPER ${'error' in r ? `not connected: ${r.error.message}` : r.out} (${Date.now() - t} ms)`));
  };
  // Edit data opens the pop-up next to the chart (Mac: the resident helper, page preloaded now;
  // Windows: Edge app mode), or the default browser if that fails.
  const base = `http://127.0.0.1:${port}`;
  const popup = popupHelper(base, process.platform === 'darwin' ? helperPath() : undefined, log);
  if (process.platform === 'darwin' && popup.warm()) log('POPUP helper started');
  const showEditor: OpenEditor = (_url, id, where) => { void openPopup(base, id, where, { helper: popup, log }); };
  const routes = addinRoutes(m, httpSessions(base, showEditor), outlived, retire, connect);
  const applyNow = (s: Session) => { void applyInPowerPoint(s, undefined, undefined, outlived, (a, context) => layoutForApply(a, context, m())); };
  // Drag handles over the selected chart (Mac helper overlay → live layout → Apply on mouse-up).
  const handles = handlesController({
    helper: popup,
    measurer: m,
    apply: (s) => applyInPowerPoint(s, undefined, undefined, outlived, (a, context) => layoutForApply(a, context, m())),
    measure: () => selectionNow(outlived),
    log,
  });
  popup.onEvent?.((msg) => handles.event(msg));
  // The element mini toolbar (selection → toolbar panel → per-element overrides → Apply).
  const toolbar = toolbarRoutes({ base, helper: popup, apply: applyNow, log, onSelect: (r) => { void handles.select(r); } });
  return serve(port, {}, {
    routes: async (req, res) => { idle(); return (await toolbar(req, res)) || routes(req, res); },
    onSession: (s) => { if (s.state === 'applied') applyNow(s); },
    confirmsApply: true,
  });
}

/** Stops a running add-in server (after an install, so the next call loads the new engine). */
export async function stopServer(base = BASE) {
  try { await fetch(`${base}/api/addin/quit`, { method: 'POST', signal: AbortSignal.timeout(1000) }); } catch { /* not running */ }
}

export const installed = (home = addinHome()) => existsSync(join(home, 'addin.json'));
