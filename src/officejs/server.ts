// Static host for the Office.js add-in during development and for air-gapped use
// (npm run officejs:serve): serves dist/officejs over https://localhost:4175 with the dev
// certificate from scripts/officejs-cert.ts, plus a log endpoint the task pane writes to:
//   POST /api/log   one line per event (ready, insert, apply, errors), appended to officejs.log
//   GET  /api/log   the lines so far (used by the live check, tools/officejs-live.ts)
//   GET  /api/selftest  {run:true} once when started with --selftest (the live check), else false
// Without a certificate it serves plain http (tests, Office on the web through a tunnel: name the
// tunnel's host with --allow-host, since only requests addressed to this server are answered;
// src/guard.ts).
import { createServer as createHttps } from 'node:https';
import { createServer as createHttp, type IncomingMessage, type ServerResponse } from 'node:http';
import { appendFileSync, existsSync, mkdirSync, readFileSync, statSync } from 'node:fs';
import { extname, join, normalize, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { addinHome } from '../addin/launcher.ts';
import { refusal } from '../guard.ts';

export const OFFICEJS_PORT = 4175;
export const certDir = (home = addinHome()) => join(home, 'officejs-cert');
export const logFile = (home = addinHome()) => join(home, 'officejs.log');

const TYPES: Record<string, string> = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.xml': 'application/xml',
  '.png': 'image/png', '.ttf': 'font/ttf', '.json': 'application/json', '.LICENSE': 'text/plain; charset=utf-8',
};

export interface OfficeServerOptions { root: string; port?: number; host?: string; tls?: { key: string; cert: string } | null; log?: string;
  /** Answer GET /api/selftest with {run:true} once, so the pane runs its live self-test. */
  selftest?: boolean;
  /** Receives each slide the self-test exports (base64 .pptx, POST /api/selftest/slide). */
  onSlide?: (b64: string) => void;
  /** Receives each log line as it arrives. */
  onLog?: (line: string) => void;
  /** Host names besides localhost/127.0.0.1 this server answers to (a tunnel). */
  allowHosts?: string[] }

export function serveOfficeJs(opts: OfficeServerOptions) {
  const lines: string[] = [];
  let selftest = !!opts.selftest;
  const guard = { port: () => (server.address() as { port: number } | null)?.port ?? port, scheme: opts.tls ? 'https' as const : 'http' as const, hosts: opts.allowHosts };
  const handler = async (req: IncomingMessage, res: ServerResponse) => {
    const why = refusal(req, guard);
    if (why) { req.resume(); res.writeHead(403, { 'content-type': 'text/plain; charset=utf-8', connection: 'close' }).end(`open-charts: request refused (${why})`); return; }
    const url = new URL(req.url ?? '/', 'http://x');
    if (url.pathname === '/api/selftest') {
      res.writeHead(200, { 'content-type': 'application/json', 'cache-control': 'no-store' }).end(JSON.stringify({ run: selftest }));
      selftest = false; // once: a reloaded pane does not run it again
      return;
    }
    if (url.pathname === '/api/selftest/slide' && req.method === 'POST' && opts.onSlide) {
      const chunks: Buffer[] = [];
      for await (const c of req) chunks.push(c as Buffer);
      opts.onSlide(Buffer.concat(chunks).toString('utf8'));
      res.writeHead(204).end();
      return;
    }
    if (url.pathname === '/api/log') {
      if (req.method === 'POST') {
        const chunks: Buffer[] = [];
        for await (const c of req) chunks.push(c as Buffer);
        const d = new Date(), local = new Date(d.getTime() - d.getTimezoneOffset() * 60_000).toISOString().slice(0, 19);
        const line = `${local} ${Buffer.concat(chunks).toString('utf8').replace(/[\r\n]+/g, ' ').slice(0, 4000)}`;
        lines.push(line);
        opts.onLog?.(line);
        if (lines.length > 2000) lines.shift();
        if (opts.log) { try { appendFileSync(opts.log, `${line}\n`); } catch { /* logging never breaks the add-in */ } }
        res.writeHead(204).end();
      } else {
        res.writeHead(200, { 'content-type': 'text/plain; charset=utf-8', 'cache-control': 'no-store' }).end(lines.join('\n') + (lines.length ? '\n' : ''));
      }
      return;
    }
    if (url.pathname === '/fonts/metrics') {
      // Exact metrics of an installed font for the pane's measurer (src/officejs/measure.ts).
      const family = url.searchParams.get('family') ?? '';
      const found = family && family.length <= 64 ? (await import('../fonts.ts')).familyMetrics(family) : null;
      // null = not installed here (the pane then measures it itself); a 404 would mean "no service".
      res.writeHead(200, { 'content-type': 'application/json', 'cache-control': 'no-store' }).end(JSON.stringify(found));
      return;
    }
    const rel = normalize(decodeURIComponent(url.pathname === '/' ? '/taskpane.html' : url.pathname)).replace(/^([/\\])+/, '');
    const root = resolve(opts.root), file = join(root, rel);
    if (rel.startsWith('..') || !file.startsWith(root) || !existsSync(file) || !statSync(file).isFile()) { res.writeHead(404).end('Not found'); return; }
    res.writeHead(200, { 'content-type': TYPES[extname(file)] ?? 'application/octet-stream', 'cache-control': 'no-store' });
    res.end(readFileSync(file));
  };
  const server = opts.tls ? createHttps({ key: opts.tls.key, cert: opts.tls.cert }, handler) : createHttp(handler);
  const port = opts.port ?? OFFICEJS_PORT;
  server.listen(port, opts.host ?? '127.0.0.1', () => {
    console.log(`open-charts Office.js add-in: ${opts.tls ? 'https' : 'http'}://localhost:${(server.address() as { port: number }).port}/taskpane.html`);
  });
  return server;
}

/** The dev certificate, when scripts/officejs-cert.ts has made one. */
export function devTls(dir = certDir()): { key: string; cert: string } | null {
  const key = join(dir, 'localhost.key'), cert = join(dir, 'localhost.crt');
  return existsSync(key) && existsSync(cert) ? { key: readFileSync(key, 'utf8'), cert: readFileSync(cert, 'utf8') } : null;
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const arg = (k: string) => { const i = process.argv.indexOf(k); return i > 0 ? process.argv[i + 1] : undefined; };
  const root = resolve(arg('--root') ?? fileURLToPath(new URL('../../dist/officejs', import.meta.url)));
  if (!existsSync(join(root, 'taskpane.html'))) { console.error(`No add-in build in ${root}. Run: npm run build:officejs`); process.exit(1); }
  const tls = process.argv.includes('--http') ? null : devTls();
  if (!tls && !process.argv.includes('--http')) { console.error('No dev certificate. Run: npm run officejs:cert   (or pass --http)'); process.exit(1); }
  mkdirSync(addinHome(), { recursive: true });
  const allowHosts = process.argv.flatMap((a, i) => (a === '--allow-host' && process.argv[i + 1] ? [process.argv[i + 1]] : []));
  serveOfficeJs({ root, port: Number(arg('--port') ?? OFFICEJS_PORT), tls, log: logFile(), selftest: process.argv.includes('--selftest'), allowHosts });
}
