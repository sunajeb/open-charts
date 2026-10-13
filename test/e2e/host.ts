// Stand-in for the PowerPoint add-in's server in the editor E2E (`node host.ts <port>`): it confirms
// every Apply as the add-in does (src/addin/server.ts), through the same onSession hook, with the
// outcome picked by the session id's prefix:
//   perm…  PowerPoint does not answer (macOS permission prompt), error after 3 s
//   gone…  PowerPoint is not running, error at once
//   other  the chart is drawn after 300 ms
// Like PowerPoint, it draws one apply at a time, in order. GET /test/log/<id> lists what reached
// "PowerPoint" for a session: { state, chart, at } per POST (the pop-up's live-update tests).
// Guarded like the add-in's server: every request but the pages needs the token (E2E_TOKEN); pages
// get it in the URL fragment (src/addin/token.ts).
import { serve, setSession, type Session } from '../../src/server.ts';
import { applyErrorMessage } from '../../src/addin/server.ts';
import { isPublicAddinRequest } from '../../src/addin/server.ts';
import { toolbarRoutes } from '../../src/addin/toolbar.ts';
import type { PopupHelper } from '../../src/addin/popup.ts';

const port = Number(process.argv[2]);
const token = process.env.E2E_TOKEN;
if (!token) throw new Error('E2E_TOKEN is not set');
const log = new Map<string, Array<{ state: string; chart: unknown; at: number }>>();
let queue = Promise.resolve();
// The element toolbar's page and routes (no helper: nothing is shown; applies are logged below).
const toolbar = toolbarRoutes({ base: '', token, helper: { info: async () => undefined, hide: () => {} } as unknown as PopupHelper, apply: (s) => onSession(s) });
serve(port, {}, {
  confirmsApply: true,
  token,
  isPublic: isPublicAddinRequest,
  routes: async (req, res) => {
    if (await toolbar(req, res)) return true;
    const m = /^\/test\/log\/([0-9A-Za-z_-]+)$/.exec(req.url ?? '');
    if (!m) return false;
    res.writeHead(200, { 'content-type': 'application/json' }).end(JSON.stringify(log.get(m[1]) ?? []));
    return true;
  },
  onSession: (s: Session) => onSession(s),
});
function onSession(s: Session) {
  if (!log.has(s.id)) log.set(s.id, []);
  log.get(s.id)!.push({ state: s.state, chart: s.chart, at: Date.now() });
  if (s.state !== 'applied') return;
  const later = (ms: number, body: { state: string; message?: string }) => setTimeout(() => setSession(s.id, body), ms);
  if (s.id.startsWith('perm')) later(3000, { state: 'error', message: applyErrorMessage({ message: 'timed out', killed: true }) });
  else if (s.id.startsWith('gone')) later(0, { state: 'error', message: applyErrorMessage({ message: 'execution error: PowerPoint is not running. (-600)' }) });
  else queue = queue.then(() => new Promise((r) => setTimeout(() => { setSession(s.id, { state: 'drawn' }); r(); }, 300)));
}
