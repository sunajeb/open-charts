// The real Mac pop-up helper (dist/OpenChartsPopup.app, `npm run build:popup`) against the editor
// E2E's stand-in host: preload, show a session, time it, edit live, cancel (revert). The panel is
// placed off screen and never takes the keyboard, so the test does not disturb a user at the Mac.
// Runs only on a Mac with the helper built and OPEN_CHARTS_POPUP_TEST=1 (it opens a real window).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { spawn, type ChildProcess } from 'node:child_process';
import { existsSync } from 'node:fs';
import { createServer } from 'node:net';
import { fileURLToPath } from 'node:url';
import { HELPER_APP, HELPER_EXE } from '../src/addin/popup.ts';

const ROOT = fileURLToPath(new URL('..', import.meta.url));
const EXE = `${ROOT}dist/${HELPER_APP}/${HELPER_EXE}`;
const skip = process.platform !== 'darwin' || !existsSync(EXE) || process.env.OPEN_CHARTS_POPUP_TEST !== '1';

const freePort = () => new Promise<number>((resolve) => {
  const s = createServer();
  s.listen(0, '127.0.0.1', () => { const { port } = s.address() as { port: number }; s.close(() => resolve(port)); });
});

function lines(c: ChildProcess) {
  const got: Array<Record<string, unknown>> = [];
  let buf = '';
  c.stdout!.setEncoding('utf8');
  c.stdout!.on('data', (d: string) => {
    buf += d;
    for (let nl = buf.indexOf('\n'); nl >= 0; nl = buf.indexOf('\n')) { got.push(JSON.parse(buf.slice(0, nl))); buf = buf.slice(nl + 1); }
  });
  const next = async (event: string, ms = 5000) => {
    const t = Date.now();
    for (;;) {
      const i = got.findIndex((m) => m.event === event);
      if (i >= 0) return got.splice(i, 1)[0];
      if (Date.now() - t > ms) throw new Error(`no ${event} from the helper within ${ms} ms (got ${JSON.stringify(got)})`);
      await new Promise((r) => setTimeout(r, 5));
    }
  };
  return next;
}

test('Mac helper: preloaded page opens a session in < 300 ms, edits reach the host live, Cancel reverts', { skip }, async () => {
  const port = await freePort();
  const host = spawn(process.execPath, ['--experimental-strip-types', '--no-warnings', 'test/e2e/host.ts', String(port)], { cwd: ROOT, stdio: ['ignore', 'pipe', 'inherit'] });
  const base = `http://127.0.0.1:${port}`;
  const helper = spawn(EXE, [], { stdio: ['pipe', 'pipe', 'inherit'] });
  try {
    for (let i = 0; i < 100; i++) { try { if ((await fetch(base)).ok) break; } catch { /* starting */ } await new Promise((r) => setTimeout(r, 100)); }
    const next = lines(helper);
    const send = (m: object) => helper.stdin!.write(`${JSON.stringify(m)}\n`);
    await next('started');
    const t0 = Date.now();
    send({ cmd: 'load', panel: 'datasheet', url: `${base}/?popup=1` });
    await next('loaded', 15_000);
    const preload = Date.now() - t0;
    const info = await (send({ cmd: 'info' }), next('info'));
    assert.ok(Array.isArray(info.screens) && (info.screens as unknown[]).length >= 1);

    const timings: number[] = [];
    let id = '';
    for (let n = 0; n < 3; n++) {
      id = `helper${Date.now().toString(16)}${n}`;
      const chart = { type: 'clustered', categories: ['A', 'B'], series: [{ name: 'S', values: [10, 20] }] };
      await fetch(`${base}/api/session/${id}`, { method: 'POST', body: JSON.stringify({ chart, title: 'Sales' }) });
      const t = Date.now();
      send({ cmd: 'show', panel: 'datasheet', id, url: `${base}/?session=${id}&popup=1`, frame: [-4000, 200, 620, 360], focus: false });
      const shown = await next('shown');
      timings.push(Date.now() - t);
      assert.equal(shown.ready, true, 'the page said it was ready (no timeout)');
      if (n < 2) { send({ cmd: 'eval', js: "ocPopupClose('cancel')" }); await next('eval'); await next('closed'); }
    }
    console.log(`# preload ${preload} ms; show -> ready: ${timings.join(', ')} ms`);
    assert.ok(Math.max(...timings) < 300, `opening took ${timings} ms`);

    // An edit, as typing would make it: the slide (host) gets it ~400 ms later.
    send({ cmd: 'eval', js: `(() => { const i = document.querySelector('#sheet input[data-r="0"][data-c="0"]'); i.focus(); i.value = '55'; i.dispatchEvent(new Event('input', { bubbles: true })); i.dispatchEvent(new Event('change', { bubbles: true })); return true; })()` });
    await next('eval');
    const log = async () => (await (await fetch(`${base}/test/log/${id}`)).json()) as Array<{ state: string; chart: { series: Array<{ values: number[] }> } }>;
    for (let i = 0; i < 50 && !(await log()).some((e) => e.state === 'applied'); i++) await new Promise((r) => setTimeout(r, 100));
    assert.deepEqual((await log()).filter((e) => e.state === 'applied').map((e) => e.chart.series[0].values), [[55, 20]]);
    await new Promise((r) => setTimeout(r, 500)); // drawn
    send({ cmd: 'eval', js: "ocPopupClose('cancel')" });
    await next('closed');
    for (let i = 0; i < 30 && !(await log()).some((e) => e.state === 'cancelled'); i++) await new Promise((r) => setTimeout(r, 100));
    const states = await log();
    assert.deepEqual(states.map((e) => e.state), ['open', 'applied', 'applied', 'cancelled']);
    assert.deepEqual(states[2].chart.series[0].values, [10, 20], 'the opening chart is put back');
  } finally {
    helper.stdin!.end();
    host.kill();
  }
});
