// The local servers accept only their own callers (src/guard.ts, src/addin/token.ts): the add-in
// server needs the per-install token on everything but its pages and ping; every server refuses a
// foreign Host (DNS rebinding) or Origin (another site in the browser) and sends no CORS headers.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync, statSync, writeFileSync, mkdirSync } from 'node:fs';
import { request as httpRequest, type Server } from 'node:http';
import type { AddressInfo } from 'node:net';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { serve } from '../src/server.ts';
import { debug, isPublicAddinRequest, log, saveSettings, settings } from '../src/addin/server.ts';
import { ensureToken, readToken, sameToken, TOKEN_BOOT, tokenHeaders, tokenPath, withToken } from '../src/addin/token.ts';
import { edgeArgs, openPopup, popupUrl, preloadUrl, toolbarUrl } from '../src/addin/popup.ts';
import { serveOfficeJs } from '../src/officejs/server.ts';

const chart = { type: 'stacked', categories: ['a', 'b'], series: [{ name: 'S', values: [1, 2] }] };
const TOKEN = 'ab'.repeat(32);

/** One request with any headers (fetch may not set Host). */
function call(port: number, path: string, o: { method?: string; headers?: Record<string, string>; body?: string } = {}): Promise<{ status: number; headers: Record<string, unknown>; body: string }> {
  return new Promise((resolve, reject) => {
    const req = httpRequest({ host: '127.0.0.1', port, path, method: o.method ?? 'GET', headers: { host: `127.0.0.1:${port}`, ...o.headers } }, (res) => {
      const chunks: Buffer[] = [];
      res.on('data', (c) => chunks.push(c));
      res.on('end', () => resolve({ status: res.statusCode!, headers: res.headers, body: Buffer.concat(chunks).toString('utf8') }));
    });
    req.on('error', reject);
    req.end(o.body);
  });
}

async function listening(server: Server): Promise<number> {
  if (!server.listening) await new Promise((r) => server.once('listening', r));
  return (server.address() as AddressInfo).port;
}

/** Starts a server without its "listening on" line. */
async function quiet<S extends { listening: boolean; once(e: 'listening', f: () => void): unknown }>(f: () => S): Promise<S> {
  const l = console.log;
  console.log = () => {};
  try { const s = f(); if (!s.listening) await new Promise<void>((r) => s.once('listening', r)); return s; } finally { console.log = l; }
}

test('token file: 32 random bytes as a header line, user-only, kept, repaired when damaged', () => {
  const home = mkdtempSync(join(tmpdir(), 'oc-token-'));
  try {
    assert.equal(readToken(home), undefined);
    assert.deepEqual(tokenHeaders(home), {});
    const t = ensureToken(home);
    assert.match(t, /^[0-9a-f]{64}$/);
    assert.equal(readFileSync(tokenPath(home), 'utf8'), `X-OC-Token: ${t}\n`, 'curl -H @file reads it as is');
    if (process.platform !== 'win32') assert.equal(statSync(tokenPath(home)).mode & 0o777, 0o600);
    assert.equal(ensureToken(home), t, 'a reinstall keeps it');
    assert.notEqual(ensureToken(mkdtempSync(join(tmpdir(), 'oc-token-'))), t, 'random per install');
    assert.deepEqual(tokenHeaders(home), { 'X-OC-Token': t });
    writeFileSync(tokenPath(home), `\uFEFFX-OC-Token: ${t}\r\n`); // as PowerShell might write it
    assert.equal(readToken(home), t);
    writeFileSync(tokenPath(home), 'garbage');
    assert.equal(readToken(home), undefined);
    assert.notEqual(ensureToken(home), t);
    assert.ok(sameToken(t, t) && !sameToken(t.slice(1), t) && !sameToken(undefined, t) && !sameToken('0'.repeat(64), t));
  } finally { rmSync(home, { recursive: true, force: true }); }
});

test('add-in server: every route but the pages and ping needs the token; Host and Origin must be its own', async () => {
  const server = await quiet(() => serve(0, {}, {
    token: TOKEN,
    isPublic: isPublicAddinRequest,
    confirmsApply: true,
    routes: (req, res) => {
      if (req.url === '/api/addin/ping') { res.writeHead(200).end('open-charts'); return true; }
      if (req.url === '/toolbar') { res.writeHead(200).end('<html>'); return true; }
      return false;
    },
  }));
  const port = await listening(server);
  const auth = { 'x-oc-token': TOKEN };
  const post = (headers: Record<string, string>) => call(port, '/api/session/s1', { method: 'POST', headers, body: JSON.stringify({ chart }) });
  try {
    // Public: the pages (static) and liveness.
    const page = await call(port, '/?session=s1&popup=1');
    assert.equal(page.status, 200);
    assert.ok(page.body.includes(TOKEN_BOOT), 'the editor page moves the fragment token into its requests');
    assert.ok(!page.body.includes(TOKEN), 'the token is never in a page');
    assert.equal((await call(port, '/toolbar')).status, 200);
    assert.equal((await call(port, '/api/addin/ping')).status, 200);

    // Without or with a wrong token: refused, nothing changes.
    assert.equal((await post({})).status, 403);
    assert.equal((await post({ 'x-oc-token': 'cd'.repeat(32) })).status, 403);
    assert.equal((await post({ 'x-oc-token': TOKEN.slice(2) })).status, 403);
    assert.equal((await call(port, '/api/session/s1')).status, 403, 'reading a session needs it too');
    assert.equal((await call(port, '/api/render', { method: 'POST', body: '{}' })).status, 403);
    assert.equal((await call(port, '/api/addin/quit', { method: 'POST' })).status, 403);
    assert.equal((await call(port, '/api/session/s1', { headers: auth })).status, 404, 'no session was made');

    // With it.
    const ok = await post(auth);
    assert.equal(ok.status, 200);
    assert.equal(ok.headers['access-control-allow-origin'], undefined, 'no CORS');
    assert.equal(JSON.parse((await call(port, '/api/session/s1', { headers: auth })).body).rev, 1);
    assert.equal((await post({ ...auth, origin: `http://127.0.0.1:${port}` })).status, 200, 'the page itself');
    assert.equal((await post({ ...auth, origin: `http://localhost:${port}`, host: `localhost:${port}` })).status, 200);

    // DNS rebinding: a page on evil.example re-pointed at 127.0.0.1 sends its own name as Host.
    for (const host of ['evil.example', `evil.example:${port}`, `127.0.0.1:${port + 1}`, `localhost.evil.example:${port}`]) {
      assert.equal((await call(port, '/api/session/s1', { headers: { ...auth, host } })).status, 403, host);
      assert.equal((await call(port, '/', { headers: { host } })).status, 403, `page via ${host}`);
    }
    // Another site in the browser (a form post or fetch carries its Origin), and the preflight.
    for (const origin of ['http://evil.example', 'null', `https://127.0.0.1:${port}`, `http://127.0.0.1:${port}.evil.example`]) {
      assert.equal((await post({ ...auth, origin })).status, 403, origin);
    }
    const pre = await call(port, '/api/session/s1', { method: 'OPTIONS', headers: { origin: 'http://evil.example', 'access-control-request-method': 'POST' } });
    assert.equal(pre.status, 403);
    assert.equal(pre.headers['access-control-allow-origin'], undefined);
    assert.equal(JSON.parse((await call(port, '/api/session/s1', { headers: auth })).body).rev, 3, 'only the two accepted posts');
  } finally { server.close(); }
});

test('plain editor (open-charts serve): no token, but the same Host and Origin checks', async () => {
  const server = await quiet(() => serve(0));
  const port = await listening(server);
  try {
    assert.equal((await call(port, '/api/session/p1', { method: 'POST', body: JSON.stringify({ chart }) })).status, 200);
    assert.equal((await call(port, '/api/render', { method: 'POST', headers: { origin: 'http://evil.example' }, body: '{}' })).status, 403);
    assert.equal((await call(port, '/api/session/p1', { headers: { host: 'evil.example' } })).status, 403);
    const page = await call(port, '/');
    assert.equal(page.status, 200);
    assert.ok(!page.body.includes(TOKEN_BOOT), 'no token boot without a token');
  } finally { server.close(); }
});

test('the Mac bridge sends the token with curl -H @file (never on a command line)', { skip: !existsSync('/usr/bin/curl') }, async () => {
  const home = mkdtempSync(join(tmpdir(), 'oc-curl-'));
  const token = ensureToken(home);
  const server = await quiet(() => serve(0, {}, { token }));
  const port = await listening(server);
  try {
    // Not execFileSync: the server answers on this thread.
    const curl = (extra: string[]) => new Promise<string>((resolve) => {
      const c = execFile('/usr/bin/curl', ['-sf', '--max-time', '5', '--data-binary', '@-', ...extra, `http://127.0.0.1:${port}/api/session/m1`], { encoding: 'utf8' },
        (err, out) => resolve(err ? 'refused' : out));
      c.stdin!.end(JSON.stringify({ chart }));
    });
    assert.equal(await curl([]), 'refused');
    assert.equal(JSON.parse(await curl(['-H', `@${tokenPath(home)}`])).rev, 1);
  } finally { server.close(); rmSync(home, { recursive: true, force: true }); }
  const script = readFileSync(new URL('../addin/applescript/open-charts.applescript', import.meta.url), 'utf8');
  for (const route of ['run', 'selection']) assert.match(script, new RegExp(`& my tokenHeader\\(\\) & " http://127\\.0\\.0\\.1:4174/api/addin/${route}`), route);
  assert.match(script, /return " -H @" & quoted form of \(homeDir & "\/token"\)/);
});

test('pages the add-in opens carry the token in the fragment, never the query string', async () => {
  const base = 'http://127.0.0.1:4174';
  for (const url of [popupUrl(base, 'abc', TOKEN), preloadUrl(base, TOKEN), toolbarUrl(base, TOKEN), withToken(`${base}/?session=x`, TOKEN)]) {
    const u = new URL(url);
    assert.equal(u.hash, `#oc-token=${TOKEN}`, url);
    assert.ok(!u.search.includes(TOKEN));
  }
  assert.equal(popupUrl(base, 'abc'), `${base}/?session=abc&popup=1`, 'unchanged without a token');
  // Windows Edge app mode and the browser fallback.
  const launched: string[][] = [], opened: string[] = [];
  await openPopup(base, 'abc', undefined, { platform: 'win32', token: TOKEN, exists: () => true, env: { LOCALAPPDATA: 'C:\\x' }, launch: (_c, a) => { launched.push(a); return true; } });
  assert.ok(launched[0].includes(`--app=${base}/?session=abc&popup=1#oc-token=${TOKEN}`));
  assert.ok(edgeArgs('u', { x: 0, y: 0, w: 1, h: 1 }, 'p').length > 0);
  await openPopup(base, 'abc', undefined, { platform: 'linux', token: TOKEN, browser: (u) => opened.push(u) });
  assert.deepEqual(opened, [`${base}/?session=abc#oc-token=${TOKEN}`]);
});

test('Office.js dev host: refuses a foreign Host or Origin', async () => {
  const root = mkdtempSync(join(tmpdir(), 'oc-ojs-'));
  mkdirSync(root, { recursive: true });
  writeFileSync(join(root, 'taskpane.html'), '<html>');
  const server = await quiet(() => serveOfficeJs({ root, port: 0, tls: null, allowHosts: ['tunnel.example'] }));
  const port = await listening(server as unknown as Server);
  try {
    assert.equal((await call(port, '/taskpane.html')).status, 200);
    assert.equal((await call(port, '/api/log', { method: 'POST', headers: { origin: `http://127.0.0.1:${port}` }, body: 'ready' })).status, 204);
    assert.equal((await call(port, '/api/log', { method: 'POST', headers: { origin: 'http://evil.example' }, body: 'x' })).status, 403);
    assert.equal((await call(port, '/api/log', { headers: { host: 'evil.example' } })).status, 403);
    assert.equal((await call(port, '/api/log', { headers: { host: 'tunnel.example', origin: 'http://tunnel.example' } })).status, 200, '--allow-host');
  } finally { (server as unknown as Server).close(); rmSync(root, { recursive: true, force: true }); }
});

test('VBA: Windows sends the token from the token file; settings turn verbose logging on', () => {
  const bas = readFileSync(new URL('../addin/vba/OpenCharts.bas', import.meta.url), 'utf8');
  assert.match(bas, /http\.setRequestHeader "X-OC-Token", WinToken\(\)/);
  assert.match(bas, /Private Const TOKEN_FILE As String = "\\open-charts\\token"/);
  assert.match(bas, /If Left\$\(t, 12\) = "X-OC-Token: " Then WinToken = Trim\$\(Mid\$\(t, 13\)\)/);
  assert.match(bas, /gVerbose = \(ignored = "OK verbose"\)/);
});

test('logging: results and errors always, diagnostics only with verbose on', () => {
  const home = mkdtempSync(join(tmpdir(), 'oc-log-'));
  const lines = () => { try { return readFileSync(join(home, 'addin.log'), 'utf8').trim().split('\n').map((l) => l.slice(20)); } catch { return []; } };
  try {
    assert.equal(settings(home).verbose, undefined);
    log('APPLY x done', home);
    debug('ENGINE insert 12 ms', home);
    assert.deepEqual(lines(), ['APPLY x done']);
    saveSettings({ verbose: true }, home);
    debug('ENGINE insert 9 ms', home);
    assert.deepEqual(lines(), ['APPLY x done', 'ENGINE insert 9 ms']);
    saveSettings({ verbose: false }, home);
    debug('ENGINE insert 7 ms', home);
    assert.equal(lines().length, 2);
  } finally { rmSync(home, { recursive: true, force: true }); }
});
