// Edit-session API used by the PowerPoint add-in (contract: docs/ux-editor.md).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import type { AddressInfo } from 'node:net';
import { serve, type Session } from '../src/server.ts';

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
    assert.equal(await (await fetch(`${base}/api/addin/ping`)).text(), 'pong', 'extra routes run first');

    // The editor page is served with the session query string.
    const page = await fetch(`${base}/?session=s1`);
    assert.equal(page.status, 200);
    assert.match(await page.text(), /<title>/);
  } finally {
    server.close();
  }
});
