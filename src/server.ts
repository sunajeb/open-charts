import { createServer, type IncomingMessage, type ServerResponse } from 'node:http';
import { readFileSync } from 'node:fs';
import { createMeasurer } from './fonts.ts';
import { buildDeck, type DeckSpec } from './deck.ts';

const EDITOR = new URL('./web/editor.html', import.meta.url);

async function body(req: IncomingMessage): Promise<DeckSpec> {
  const chunks: Buffer[] = [];
  for await (const c of req) chunks.push(c as Buffer);
  return JSON.parse(Buffer.concat(chunks).toString('utf8'));
}

// Edit sessions: the PowerPoint add-in stores a chart here, opens the editor with ?session=<id>, and
// reads the chart back when the editor applies it (contract: docs/ux-editor.md). In memory only.
export interface Session { id: string; rev: number; state: 'open' | 'applied' | 'cancelled'; chart: unknown; title?: string; updatedAt: string }
const SESSION_ID = /^[A-Za-z0-9_-]{1,64}$/;
const SESSION_STATES = ['open', 'applied', 'cancelled'];
const MAX_SESSIONS = 200;
const sessions = new Map<string, Session>();

function postSession(id: string, b: { chart?: unknown; title?: unknown; state?: unknown }): Session {
  const prev = sessions.get(id);
  const state = b.state ?? 'open';
  if (!SESSION_STATES.includes(state as string)) throw new Error(`state must be one of ${SESSION_STATES.join(', ')}`);
  const chart = b.chart ?? prev?.chart;
  if (!chart || typeof chart !== 'object') throw new Error('a new session needs a "chart" object');
  if (b.title !== undefined && typeof b.title !== 'string') throw new Error('title must be a string');
  const s: Session = {
    id, rev: (prev?.rev ?? 0) + 1, state: state as Session['state'], chart,
    title: (b.title as string | undefined) ?? prev?.title, updatedAt: new Date().toISOString(),
  };
  sessions.delete(id);
  sessions.set(id, s);
  if (sessions.size > MAX_SESSIONS) sessions.delete(sessions.keys().next().value!);
  return s;
}

/** Local editor: datasheet + live preview, renders with the same engine as the CLI. */
export interface ServeOptions {
  /** Extra routes checked before the built-in ones (the add-in's local endpoints); return true when handled. */
  routes?: (req: IncomingMessage, res: ServerResponse) => boolean | Promise<boolean>;
  /** Called after every successful POST /api/session/<id> (e.g. the add-in re-lays out on "applied"). */
  onSession?: (s: Session) => void;
}

export function serve(port: number, fonts: { font?: string; bold?: string } = {}, opts: ServeOptions = {}) {
  const m = createMeasurer(fonts.font, fonts.bold);
  const server = createServer(async (req, res) => {
    try {
      if (opts.routes && (await opts.routes(req, res))) return;
      if (req.method === 'GET' && /^\/(index\.html)?(\?|$)/.test(req.url ?? '')) {
        res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' });
        res.end(readFileSync(EDITOR));
        return;
      }
      const sm = req.url?.match(/^\/api\/session\/([^/?]+)$/);
      if (sm) {
        const id = decodeURIComponent(sm[1]);
        const cors = { 'access-control-allow-origin': '*', 'access-control-allow-methods': 'GET, POST, OPTIONS', 'access-control-allow-headers': 'content-type' };
        if (req.method === 'OPTIONS') { res.writeHead(204, cors).end(); return; }
        if (!SESSION_ID.test(id)) { res.writeHead(400, { ...cors, 'content-type': 'application/json' }).end(JSON.stringify({ error: 'session id must be 1–64 letters, digits, - or _' })); return; }
        const json = (code: number, v: unknown) => res.writeHead(code, { ...cors, 'content-type': 'application/json', 'cache-control': 'no-store' }).end(JSON.stringify(v));
        if (req.method === 'GET') {
          const s = sessions.get(id);
          if (s) json(200, s); else json(404, { error: `no edit session "${id}"` });
          return;
        }
        if (req.method === 'POST') {
          const saved = postSession(id, (await body(req)) as never);
          json(200, saved);
          opts.onSession?.(saved);
          return;
        }
      }
      if (req.method === 'POST' && req.url === '/api/render') {
        const { slides } = buildDeck(await body(req), m);
        res.writeHead(200, { 'content-type': 'application/json' });
        res.end(JSON.stringify({ slides: slides.map((s) => ({ svg: s.svg, issues: s.issues })) }));
        return;
      }
      if (req.method === 'POST' && req.url === '/api/pptx') {
        const { pptx } = buildDeck(await body(req), m);
        res.writeHead(200, {
          'content-type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
          'content-disposition': 'attachment; filename="charts.pptx"',
        });
        res.end(Buffer.from(pptx));
        return;
      }
      res.writeHead(404).end('Not found');
    } catch (err) {
      res.writeHead(400, { 'content-type': 'application/json', ...(req.url?.startsWith('/api/session/') ? { 'access-control-allow-origin': '*' } : {}) });
      res.end(JSON.stringify({ error: (err as Error).message }));
    }
  });
  server.listen(port, '127.0.0.1', () => console.log(`open-charts editor: http://localhost:${port}`));
  return server;
}
