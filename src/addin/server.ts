// The add-in's local server (127.0.0.1:4174): the same `serve()` as `open-charts serve` (editor,
// render API, edit sessions) plus
//   POST /api/addin/run    the warm engine: a request in, the line format out (no node start-up)
//   GET  /api/addin/ping   liveness
//   POST /api/addin/quit   stop (install restarts it with the new engine)
// When the editor applies a session (docs/ux-editor.md), it tells PowerPoint to re-lay out the
// chart: VBA on the Mac cannot poll, so the server runs the add-in's ApplyEdit macro via AppleScript
// and records the outcome on the session ("drawn" or "error" with a message) for the editor to show.
import { execFile, spawn } from 'node:child_process';
import { appendFileSync, existsSync, mkdirSync, readFileSync } from 'node:fs';
import { addinHome, reinstallCommand } from './launcher.ts';
import { join } from 'node:path';
import type { IncomingMessage, ServerResponse } from 'node:http';
import { getSession, serve, setSession, type Session } from '../server.ts';
import { createMeasurer } from '../fonts.ts';
import type { Measurer } from '../core/index.ts';
import { handleRequest, parseRequest, type Sessions } from './engine.ts';

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
  spawn(process.execPath, ['--experimental-strip-types', '--no-warnings', CLI.pathname, 'addin', 'serve'], { detached: true, stdio: 'ignore' }).unref();
  for (let i = 0; i < 60; i++) {
    await new Promise((r) => setTimeout(r, 100));
    if (await ping(base)) return true;
  }
  return false;
}

/** Edit sessions over HTTP (works from the server itself and from a one-off `addin run`). */
export function httpSessions(base = BASE, open = (url: string) => { spawn('open', [url], { stdio: 'ignore', detached: true }).unref(); }): Sessions {
  return {
    async get(id) {
      try {
        const r = await fetch(`${base}/api/session/${id}`, { signal: AbortSignal.timeout(2000) });
        return r.ok ? ((await r.json()) as Session) : undefined;
      } catch { return undefined; }
    },
    async open(id, chart, title) {
      if (!(await ensureServer(base))) throw new Error('The open-charts editor could not be started.');
      const r = await fetch(`${base}/api/session/${id}`, { method: 'POST', body: JSON.stringify({ chart, title, state: 'open' }) });
      if (!r.ok) throw new Error(`The editor refused the chart: ${await r.text()}`);
      open(`http://localhost:${ADDIN_PORT}/?session=${id}`);
    },
    async settle(id) {
      try {
        await fetch(`${base}/api/session/${id}`, { method: 'POST', body: JSON.stringify({ state: 'drawn' }), signal: AbortSignal.timeout(2000) });
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

/** How long an Apply may take in PowerPoint before the editor is told it did not get through. */
export const APPLY_TIMEOUT_MS = 20_000;

/**
 * The AppleScript that runs ApplyEdit. It never starts PowerPoint (a `tell` would): a closed
 * PowerPoint has no chart to update.
 */
export const applyScript = (prefix: string, id: string) =>
  `if application "Microsoft PowerPoint" is not running then error "PowerPoint is not running." number -600
tell application "Microsoft PowerPoint" to run VB macro macro name "${prefix}_ApplyEdit" list of parameters {"${id}"}`;

/** What the user should be told when `run VB macro` failed (osascript's error, or a timeout). */
export function applyErrorMessage(err: { message: string; killed?: boolean }): string {
  const text = err.message;
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

type Run = (script: string, timeoutMs: number) => Promise<{ message: string; killed?: boolean } | undefined>;
const osascriptRun: Run = (script, timeoutMs) =>
  new Promise((resolve) => execFile('osascript', ['-e', script], { timeout: timeoutMs }, (err, _out, stderr) =>
    resolve(err ? { message: `${err.message}\n${stderr}`, killed: err.killed } : undefined)));

/**
 * Runs the add-in's ApplyEdit macro for an applied session and records the outcome: the engine
 * marks the session "drawn" when it lays the chart out; if the macro failed, timed out or found no
 * chart, the session (still at this Apply's rev) becomes "error" with what to do.
 */
export async function applyInPowerPoint(s: Session, home = addinHome(), run: Run = osascriptRun): Promise<void> {
  const fail = (message: string) => {
    const now = getSession(s.id);
    if (now?.state === 'applied' && now.rev === s.rev) setSession(s.id, { state: 'error', message });
  };
  let prefix = '';
  try { prefix = JSON.parse(readFileSync(join(home, 'addin.json'), 'utf8')).prefix; } catch { /* not installed */ }
  if (!prefix) {
    log(`APPLY ${s.id} skipped: add-in not installed`, home);
    return fail(`The open-charts add-in is not installed. In Terminal, run: ${reinstallCommand(home)}`);
  }
  const t = Date.now();
  const err = await run(applyScript(prefix, s.id), APPLY_TIMEOUT_MS);
  const ms = Date.now() - t;
  if (err) {
    log(`APPLY ${s.id} failed after ${ms} ms: ${err.killed ? 'timed out' : err.message}`, home);
    return fail(applyErrorMessage(err));
  }
  log(`APPLY ${s.id} ${getSession(s.id)?.state === 'applied' ? 'ran, but no chart was updated' : 'done'} (${ms} ms)`, home);
  fail('PowerPoint could not find the chart. Click it once on the slide, then Apply again (or click Update chart).');
}

export function addinRoutes(m: () => Measurer, sessions: Sessions) {
  return async (req: IncomingMessage, res: ServerResponse): Promise<boolean> => {
    const url = req.url ?? '';
    if (req.method === 'GET' && url === '/api/addin/ping') {
      res.writeHead(200).end('open-charts');
      return true;
    }
    if (req.method === 'POST' && url === '/api/addin/run') {
      const chunks: Buffer[] = [];
      for await (const c of req) chunks.push(c as Buffer);
      const out = await handleRequest(parseRequest(Buffer.concat(chunks).toString('utf8')), m(), sessions);
      res.writeHead(200, { 'content-type': 'text/plain; charset=utf-8' }).end(out);
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
  const routes = addinRoutes(m, httpSessions(`http://127.0.0.1:${port}`));
  return serve(port, {}, {
    routes: (req, res) => { idle(); return routes(req, res); },
    onSession: (s) => { if (s.state === 'applied') void applyInPowerPoint(s); },
    confirmsApply: true,
  });
}

/** Stops a running add-in server (after an install, so the next call loads the new engine). */
export async function stopServer(base = BASE) {
  try { await fetch(`${base}/api/addin/quit`, { method: 'POST', signal: AbortSignal.timeout(1000) }); } catch { /* not running */ }
}

export const installed = (home = addinHome()) => existsSync(join(home, 'addin.json'));
