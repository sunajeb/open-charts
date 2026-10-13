// Edit-session API used by the PowerPoint add-in (contract: docs/ux-editor.md).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import type { AddressInfo } from 'node:net';
import { editorHtml, getSession, serve, type Session } from '../src/server.ts';
import { editorSpec } from '../src/core/spec.ts';
import { DEFAULT_THEME, defaultStyle } from '../src/core/index.ts';

test('the editor page carries the engine option table, per-type defaults and limits before its script', () => {
  const html = editorHtml();
  const m = /<script>window\.OC_SPEC = (.*?);<\/script>/.exec(html);
  assert.ok(m, 'OC_SPEC is set');
  assert.ok(html.indexOf(m[0]) < html.indexOf('<script>', html.indexOf(m[0]) + m[0].length), 'before the editor script');
  assert.deepEqual(JSON.parse(m[1]), JSON.parse(JSON.stringify({ ...editorSpec(), palette: defaultStyle.palette, theme: DEFAULT_THEME })));
  const spec = editorSpec();
  assert.equal(spec.defaults.mekko.totals, true);
  assert.equal(spec.defaults.line.valueAxis, true);
  assert.deepEqual(spec.options.cagr, ['stacked', 'percent']);
  assert.equal(spec.limits.gap.max, 10);
});

const chart = { type: 'stacked', categories: ['a', 'b'], series: [{ name: 'S', values: [1, 2] }] };

test('session: create, read, apply, cancel; bad ids and unknown sessions are rejected', async () => {
  const log = console.log; console.log = () => {};
  const seen: Session[] = [];
  const server = serve(0, {}, {
    routes: (req, res) => {
      if (req.url !== '/api/addin/ping') return false;
      res.writeHead(200).end('pong');
      return true;
    },
    onSession: (s) => seen.push(s),
  });
  await new Promise((r) => server.once('listening', r));
  console.log = log;
  const base = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
  const post = (id: string, b: unknown) => fetch(`${base}/api/session/${id}`, { method: 'POST', body: JSON.stringify(b) });
  try {
    assert.equal((await fetch(`${base}/api/session/nope`)).status, 404);
    assert.equal((await fetch(`${base}/api/session/bad%20id`)).status, 400);
    assert.equal((await post('s1', {})).status, 400, 'a new session needs a chart');

    let r = await post('s1', { chart, title: 'Revenue' });
    assert.equal(r.status, 200);
    assert.equal(r.headers.get('access-control-allow-origin'), '*');
    let s = await r.json();
    assert.deepEqual([s.rev, s.state, s.title], [1, 'open', 'Revenue']);

    s = await (await fetch(`${base}/api/session/s1`)).json();
    assert.deepEqual(s.chart, chart);

    const edited = { ...chart, series: [{ name: 'S', values: [5, 6] }] };
    s = await (await post('s1', { chart: edited, state: 'applied' })).json();
    assert.deepEqual([s.rev, s.state, s.title], [2, 'applied', 'Revenue']);
    s = await (await post('s1', { state: 'cancelled' })).json();
    assert.deepEqual([s.rev, s.state], [3, 'cancelled']);
    assert.deepEqual(s.chart, edited, 'cancel keeps the last applied chart');
    assert.equal((await post('s1', { state: 'done' })).status, 400);

    assert.deepEqual(seen.map((x) => [x.rev, x.state]), [[1, 'open'], [2, 'applied'], [3, 'cancelled']], 'onSession after each successful POST');
    assert.equal(s.confirms, undefined, 'a plain host does not confirm Apply');
    assert.equal(await (await fetch(`${base}/api/addin/ping`)).text(), 'pong', 'extra routes run first');

    // The host's outcome of an Apply: drawn, or error with a message (cleared by the next state).
    s = await (await post('s1', { chart: edited, state: 'applied' })).json();
    s = await (await post('s1', { state: 'error', message: 'PowerPoint isn\'t running.' })).json();
    assert.deepEqual([s.rev, s.state, s.message], [5, 'error', 'PowerPoint isn\'t running.']);
    s = await (await post('s1', { state: 'drawn' })).json();
    assert.deepEqual([s.rev, s.state, s.message], [6, 'drawn', undefined]);
    assert.deepEqual(s.chart, edited);
    assert.equal((await post('s1', { state: 'error', message: 42 })).status, 400);
  } finally {
    server.close();
  }
});

test('session: a host that confirms Apply says so on every session', async () => {
  const log = console.log; console.log = () => {};
  const server = serve(0, {}, { confirmsApply: true });
  await new Promise((r) => server.once('listening', r));
  console.log = log;
  const base = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
  try {
    const s = await (await fetch(`${base}/api/session/c1`, { method: 'POST', body: JSON.stringify({ chart }) })).json();
    assert.equal(s.confirms, true);
    assert.equal((await (await fetch(`${base}/api/session/c1`)).json()).confirms, true);
    assert.equal(getSession('c1')!.confirms, undefined, 'not stored');

    // The editor page is served with the session query string.
    const page = await fetch(`${base}/?session=s1`);
    assert.equal(page.status, 200);
    assert.match(await page.text(), /<title>/);
  } finally {
    server.close();
  }
});
