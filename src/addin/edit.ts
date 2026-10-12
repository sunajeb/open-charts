// "Edit data" for the PowerPoint add-in: opens the chart's spec in the local editor
// (src/web/editor.html, served by `open-charts serve`) and keeps what the user types there, so the
// next "Update chart" on that chart picks it up. Charts are matched by a hash of their OC_SPEC,
// so no extra tag is needed: an edit applies to the chart whose data it started from.
import { spawn } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';
import type { IncomingMessage, ServerResponse } from 'node:http';
import { serve } from '../server.ts';
import { handleRequest, parseRequest, type PendingEdits } from './engine.ts';
import { createMeasurer } from '../fonts.ts';
import type { Measurer } from '../core/index.ts';

export const EDIT_PORT = 4174;
const EDITOR = new URL('../web/editor.html', import.meta.url);
const CLI = new URL('../cli.ts', import.meta.url);

/** Where the add-in keeps its engine copy, edits and log. */
export function addinHome(): string {
  return process.env.OPEN_CHARTS_HOME ?? join(homedir(), 'Library', 'Application Support', 'open-charts');
}

const editsDir = (home: string) => join(home, 'edits');
const keyOk = (key: string) => /^[0-9a-f]{16}$/.test(key);

export function fileEdits(home = addinHome(), launch = launchEditor): PendingEdits {
  return {
    take(key) {
      const file = join(editsDir(home), `${key}.json`);
      if (!existsSync(file)) return undefined;
      const spec = readFileSync(file, 'utf8');
      rmSync(file, { force: true });
      rmSync(join(editsDir(home), `${key}.base.json`), { force: true });
      return spec;
    },
    start(key, spec) {
      mkdirSync(editsDir(home), { recursive: true });
      writeFileSync(join(editsDir(home), `${key}.base.json`), spec);
      launch(`http://localhost:${EDIT_PORT}/edit/${key}`);
      return 'The chart\'s data is open in your browser. Edit it there, then select the chart in PowerPoint and click Update chart.';
    },
  };
}

/**
 * Starts the edit server in the background unless it is already running, then opens the page.
 * Runs in a detached helper process so the PowerPoint call returns at once.
 */
export function launchEditor(url: string) {
  const script = `
    import { spawn } from 'node:child_process';
    const ping = () => fetch('http://127.0.0.1:${EDIT_PORT}/api/addin/ping').then((r) => r.ok, () => false);
    if (!(await ping())) {
      spawn(process.execPath, ['--experimental-strip-types', '--no-warnings', ${JSON.stringify(CLI.pathname)}, 'addin', 'serve'],
        { detached: true, stdio: 'ignore' }).unref();
      for (let i = 0; i < 50 && !(await ping()); i++) await new Promise((r) => setTimeout(r, 100));
    }
    spawn('open', [${JSON.stringify(url)}], { stdio: 'ignore' });`;
  spawn(process.execPath, ['--input-type=module', '-e', script], { detached: true, stdio: 'ignore' }).unref();
}

const STORE = (() => {
  try { return /const STORE = '([^']+)'/.exec(readFileSync(EDITOR, 'utf8'))?.[1] ?? 'open-charts-deck-v1'; } catch { return 'open-charts-deck-v1'; }
})();

/** The editor page preloaded with one chart; changes are posted back as the user edits. */
export function editPage(key: string, spec: string): string {
  const parsed = JSON.parse(spec);
  const { style, ...chart } = parsed;
  const deck = style ? { style, slides: [{ chart }] } : { slides: [{ chart }] };
  const boot = `<script>(function () {
  var key = ${JSON.stringify(key)}, store = ${JSON.stringify(STORE)};
  try { localStorage.setItem(store, ${JSON.stringify(JSON.stringify(deck)).replace(/</g, '\\u003c')}); } catch (e) {}
  var last = null;
  setInterval(function () {
    var d = null;
    try { d = localStorage.getItem(store); } catch (e) {}
    if (d && d !== last) {
      if (last !== null) fetch('/api/addin/save/' + key, { method: 'POST', body: d });
      last = d;
    }
  }, 500);
})();</script>`;
  const html = readFileSync(EDITOR, 'utf8');
  const at = html.indexOf('<script');
  return at < 0 ? html + boot : html.slice(0, at) + boot + html.slice(at);
}

/** Routes added to the editor server for edit sessions. */
export function editRoutes(home = addinHome()) {
  return async (req: IncomingMessage, res: ServerResponse): Promise<boolean> => {
    const url = req.url ?? '';
    if (req.method === 'GET' && url === '/api/addin/ping') {
      res.writeHead(200).end('open-charts');
      return true;
    }
    if (req.method === 'POST' && url === '/api/addin/run') {
      // The warm engine: same request and answer as `open-charts addin run`, without node's start-up.
      const chunks: Buffer[] = [];
      for await (const c of req) chunks.push(c as Buffer);
      res.writeHead(200, { 'content-type': 'text/plain; charset=utf-8' });
      res.end(handleRequest(parseRequest(Buffer.concat(chunks).toString('utf8')), measurer(), fileEdits(home)));
      return true;
    }
    if (req.method === 'POST' && url === '/api/addin/quit') {
      res.writeHead(204).end();
      setImmediate(() => process.exit(0));
      return true;
    }
    let m = /^\/edit\/([0-9a-f]+)$/.exec(url);
    if (req.method === 'GET' && m && keyOk(m[1])) {
      const base = join(editsDir(home), `${m[1]}.base.json`);
      const pending = join(editsDir(home), `${m[1]}.json`);
      const spec = existsSync(pending) ? readFileSync(pending, 'utf8') : existsSync(base) ? readFileSync(base, 'utf8') : undefined;
      if (!spec) { res.writeHead(404).end('This edit session has ended. Click Edit data in PowerPoint again.'); return true; }
      res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' });
      res.end(editPage(m[1], spec));
      return true;
    }
    m = /^\/api\/addin\/save\/([0-9a-f]+)$/.exec(url);
    if (req.method === 'POST' && m && keyOk(m[1])) {
      const chunks: Buffer[] = [];
      for await (const c of req) chunks.push(c as Buffer);
      const deck = JSON.parse(Buffer.concat(chunks).toString('utf8'));
      const chart = deck?.slides?.[0]?.chart;
      if (!chart) { res.writeHead(400).end('no chart'); return true; }
      const spec = deck.style && Object.keys(deck.style).length ? { ...chart, style: deck.style } : chart;
      mkdirSync(editsDir(home), { recursive: true });
      writeFileSync(join(editsDir(home), `${m[1]}.json`), JSON.stringify(spec));
      res.writeHead(204).end();
      return true;
    }
    return false;
  };
}

let cachedMeasurer: Measurer | undefined;
const measurer = () => (cachedMeasurer ??= createMeasurer());

/** The add-in's local server: warm engine, edit sessions and the editor. Exits after an idle hour. */
export function serveEdits(port = EDIT_PORT) {
  const routes = editRoutes();
  let timer: NodeJS.Timeout | undefined;
  const idle = () => { clearTimeout(timer); timer = setTimeout(() => process.exit(0), 60 * 60 * 1000); timer.unref(); };
  idle();
  measurer();
  return serve(port, {}, (req, res) => { idle(); return routes(req, res); });
}

/** Stops a running add-in server (after an install, so the next call loads the new engine). */
export async function stopServer(port = EDIT_PORT) {
  try { await fetch(`http://127.0.0.1:${port}/api/addin/quit`, { method: 'POST', signal: AbortSignal.timeout(1000) }); } catch { /* not running */ }
}
