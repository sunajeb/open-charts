import { createServer, type IncomingMessage, type ServerResponse } from 'node:http';
import { readFileSync } from 'node:fs';
import { createMeasurer } from './fonts.ts';
import { buildDeck, type DeckSpec } from './deck.ts';
import { editorSpec } from './core/spec.ts';
import { DEFAULT_THEME, defaultStyle } from './core/index.ts';
import { editorGallery } from './addin/registry.ts';
import { refusal } from './guard.ts';
import { withTokenBoot } from './addin/token.ts';

const EDITOR = new URL('./web/editor.html', import.meta.url);

/**
 * The editor page with the engine's option table, per-type defaults and limits (`window.OC_SPEC`,
 * from src/core/spec.ts) set before its script runs, so its controls never drift from the engine.
 */
export function editorHtml(): string {
  const html = readFileSync(EDITOR, 'utf8');
  // The default palette (theme slots) and the colours the slots stand for, for the colour swatches.
  // The chart gallery (tiles, icons, directions) comes from the add-in's registry: one catalogue for both.
  const spec = { ...editorSpec(), palette: defaultStyle.palette, theme: DEFAULT_THEME, gallery: editorGallery() };
  const boot = `<script>window.OC_SPEC = ${JSON.stringify(spec).replace(/</g, '\\u003c')};</script>\n`;
  const at = html.indexOf('<script');
  return html.slice(0, at) + boot + html.slice(at);
}

async function body(req: IncomingMessage): Promise<DeckSpec> {
  const chunks: Buffer[] = [];
  for await (const c of req) chunks.push(c as Buffer);
  return JSON.parse(Buffer.concat(chunks).toString('utf8'));
}

// Edit sessions: the PowerPoint add-in stores a chart here, opens the editor with ?session=<id>, and
// reads the chart back when the editor applies it (contract: docs/ux-editor.md). In memory only.
//   open → applied (editor's Apply) → drawn (the chart on the slide was updated)
//                                   → error (it was not; `message` says why and what to do)
//   any → cancelled (editor's Cancel)
export type SessionState = 'open' | 'applied' | 'drawn' | 'error' | 'cancelled';
export interface Session {
  id: string; rev: number; state: SessionState; chart: unknown; title?: string; updatedAt: string;
  /** Why an Apply did not reach the slide (state "error" only). */
  message?: string;
  /** In responses only: this host confirms every Apply by moving the session to "drawn" or "error". */
  confirms?: true;
}
const SESSION_ID = /^[A-Za-z0-9_-]{1,64}$/;
const SESSION_STATES: SessionState[] = ['open', 'applied', 'drawn', 'error', 'cancelled'];
const MAX_SESSIONS = 200;
const sessions = new Map<string, Session>();

/** Creates or updates a session (the body of POST /api/session/<id>); throws on a bad body. */
export function setSession(id: string, b: { chart?: unknown; title?: unknown; state?: unknown; message?: unknown }): Session {
  const prev = sessions.get(id);
  const state = b.state ?? 'open';
  if (!SESSION_STATES.includes(state as SessionState)) throw new Error(`state must be one of ${SESSION_STATES.join(', ')}`);
  const chart = b.chart ?? prev?.chart;
  if (!chart || typeof chart !== 'object') throw new Error('a new session needs a "chart" object');
  if (b.title !== undefined && typeof b.title !== 'string') throw new Error('title must be a string');
  if (b.message !== undefined && typeof b.message !== 'string') throw new Error('message must be a string');
  const s: Session = {
    id, rev: (prev?.rev ?? 0) + 1, state: state as SessionState, chart,
    title: (b.title as string | undefined) ?? prev?.title, updatedAt: new Date().toISOString(),
  };
  if (state === 'error') s.message = (b.message as string | undefined) || 'The chart on the slide was not updated.';
  sessions.delete(id);
  sessions.set(id, s);
  if (sessions.size > MAX_SESSIONS) sessions.delete(sessions.keys().next().value!);
  return s;
}

export const getSession = (id: string): Session | undefined => sessions.get(id);

/** Local editor: datasheet + live preview, renders with the same engine as the CLI. */
export interface ServeOptions {
  /** Extra routes checked before the built-in ones (the add-in's local endpoints); return true when handled. */
  routes?: (req: IncomingMessage, res: ServerResponse) => boolean | Promise<boolean>;
  /** Called after every successful POST /api/session/<id> (e.g. the add-in re-lays out on "applied"). */
  onSession?: (s: Session) => void;
  /**
   * The host confirms every Apply: after "applied" it sets the session to "drawn" or "error"
   * (setSession, or a POST). Sessions are then answered with `confirms: true`, and the editor waits
   * for that outcome before it says the slide was updated.
   */
  confirmsApply?: boolean;
  /**
   * The add-in server's per-install token (src/addin/token.ts): every request but the public ones
   * (the editor page; `isPublic`) must carry it. Host and Origin are checked either way (src/guard.ts).
   */
  token?: string;
  /** More requests that need no token (the add-in's toolbar page and ping). */
  isPublic?: (req: IncomingMessage) => boolean;
}

/** The editor page (GET /, /index.html, with any query string): static, no session data in it. */
export const isEditorPage = (req: IncomingMessage) => req.method === 'GET' && /^\/(index\.html)?(\?|$)/.test(req.url ?? '');

export function serve(port: number, fonts: { font?: string; bold?: string } = {}, opts: ServeOptions = {}) {
  const m = createMeasurer(fonts.font, fonts.bold);
  const view = (s: Session): Session => (opts.confirmsApply ? { ...s, confirms: true } : s);
  const guard = {
    port: () => (server.address() as { port: number } | null)?.port ?? port,
    token: opts.token,
    isPublic: (req: IncomingMessage) => isEditorPage(req) || (opts.isPublic?.(req) ?? false),
  };
  const server = createServer(async (req, res) => {
    try {
      const why = refusal(req, guard);
      if (why) {
        req.resume();
        res.writeHead(403, { 'content-type': 'application/json', connection: 'close' }).end(JSON.stringify({ error: `open-charts: request refused (${why})` }));
        return;
      }
      if (opts.routes && (await opts.routes(req, res))) return;
      if (isEditorPage(req)) {
        res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' });
        res.end(opts.token ? withTokenBoot(editorHtml()) : editorHtml());
        return;
      }
      const sm = req.url?.match(/^\/api\/session\/([^/?]+)$/);
      if (sm) {
        const id = decodeURIComponent(sm[1]);
        if (!SESSION_ID.test(id)) { res.writeHead(400, { 'content-type': 'application/json' }).end(JSON.stringify({ error: 'session id must be 1–64 letters, digits, - or _' })); return; }
        const json = (code: number, v: unknown) => res.writeHead(code, { 'content-type': 'application/json', 'cache-control': 'no-store' }).end(JSON.stringify(v));
        if (req.method === 'GET') {
          const s = sessions.get(id);
          if (s) json(200, view(s)); else json(404, { error: `no edit session "${id}"` });
          return;
        }
        if (req.method === 'POST') {
          const saved = setSession(id, (await body(req)) as never);
          json(200, view(saved));
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
      res.writeHead(400, { 'content-type': 'application/json' });
      res.end(JSON.stringify({ error: (err as Error).message }));
    }
  });
  server.listen(port, '127.0.0.1', () => console.log(`open-charts editor: http://localhost:${port}`));
  return server;
}
