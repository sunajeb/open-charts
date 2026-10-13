// Help: the first-run welcome, the sample deck and "Report a problem" (src/addin/support.ts), and
// their ribbon / VBA / bridge wiring.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createMeasurer } from '../src/fonts.ts';
import { handleRequest, parseRequest } from '../src/addin/engine.ts';
import { fileCommand, handleSupport, LOG_TAIL, WELCOME, writeSampleDeck, type Shell } from '../src/addin/support.ts';
import { buildAddin, HELP_BUTTONS } from '../src/addin/package.ts';
import { insertableItems } from '../src/addin/registry.ts';
import { unzip } from '../src/ooxml/unzip.ts';
import { ENGINE_VERSION } from '../src/version.ts';

const m = createMeasurer();
const tempHome = () => mkdtempSync(join(tmpdir(), 'oc-support-'));
const fakeShell = () => {
  const calls: string[] = [];
  return { calls, shell: { open: (p) => calls.push(`open ${p}`), reveal: (p) => calls.push(`reveal ${p}`) } as Shell };
};
const SPEC = '{"type":"stacked","categories":["A"],"series":[{"name":"S","values":[1]}]}';

test('welcome: the engine answers the text with the sample question, and remembers it was shown', async () => {
  const home = tempHome(), was = process.env.OPEN_CHARTS_HOME;
  process.env.OPEN_CHARTS_HOME = home;
  try {
    writeFileSync(join(home, 'settings.json'), '{"verbose": true}\n');
    const out = await handleRequest(parseRequest('welcome'), m);
    const lines = out.trim().split('\n');
    assert.equal(lines[0], 'OC\t2');
    assert.ok(lines[1].startsWith('M\tWelcome to open-charts'));
    assert.match(out, /Insert chart gallery/);
    assert.match(out, /USER-GUIDE\.md/);
    assert.match(out, /Open a sample deck now\?\n$/);
    assert.deepEqual(JSON.parse(readFileSync(join(home, 'settings.json'), 'utf8')), { verbose: true, welcomed: true });
  } finally {
    if (was === undefined) delete process.env.OPEN_CHARTS_HOME; else process.env.OPEN_CHARTS_HOME = was;
    rmSync(home, { recursive: true, force: true });
  }
});

test('sample deck: one slide per insertable chart type, opened with its app', () => {
  const home = tempHome();
  try {
    const { calls, shell } = fakeShell();
    const out = handleSupport('sample', '', {}, m, home, shell);
    assert.equal(out, 'OC\t2\n', 'nothing to tell on success');
    const path = join(home, 'open-charts-sample.pptx');
    assert.deepEqual(calls, [`open ${path}`]);
    const slides = [...unzip(readFileSync(path)).keys()].filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));
    assert.equal(slides.length, insertableItems().length);
    assert.equal(writeSampleDeck(m, home), path, 'written again in place');
  } finally { rmSync(home, { recursive: true, force: true }); }
});

test('report: log, settings, addin.json and versions in a zip; never the token; the chart only when agreed', () => {
  const home = tempHome();
  try {
    writeFileSync(join(home, 'addin.log'), `${'x'.repeat(LOG_TAIL)}last line\n`);
    writeFileSync(join(home, 'settings.json'), '{"verbose": true}\n');
    writeFileSync(join(home, 'addin.json'), '{"prefix": "OCX"}\n');
    writeFileSync(join(home, 'token'), 'X-OC-Token: secret\n');
    const { calls, shell } = fakeShell();
    const req = parseRequest(['report', 'spec', '', '', 'app=16.89 (build 25091420) Macintosh', SPEC].join('\n'));
    assert.equal(req.app, '16.89 (build 25091420) Macintosh');
    const out = handleSupport(req.op, req.arg, { app: req.app, spec: req.spec }, m, home, shell);
    const path = /^M\tSaved a problem report:\\n(.+?\.zip)\\n/m.exec(out)?.[1];
    assert.ok(path && path.startsWith(join(home, 'reports', 'open-charts-report-')), out);
    assert.match(out, /and the selected chart's data\. No slide content\./);
    assert.deepEqual(calls, [`reveal ${path}`]);
    const zip = unzip(readFileSync(path));
    assert.deepEqual([...zip.keys()].sort(), ['addin.json', 'addin.log', 'chart-spec.json', 'settings.json', 'versions.json']);
    const log = zip.get('addin.log')!.text();
    assert.equal(log.length, LOG_TAIL, 'only the log\'s tail');
    assert.ok(log.endsWith('last line\n'));
    assert.equal(zip.get('chart-spec.json')!.text(), `${SPEC}\n`);
    const versions = JSON.parse(zip.get('versions.json')!.text());
    assert.equal(versions.openCharts, ENGINE_VERSION);
    assert.equal(versions.powerPoint, '16.89 (build 25091420) Macintosh');
    assert.equal(versions.node, process.versions.node);
    for (const [, e] of zip) assert.doesNotMatch(e.text(), /secret/);

    // Without "spec" (the user said no) the chart stays out, whatever follows the params.
    const no = handleSupport('report', '', { spec: SPEC }, m, home, shell);
    const noPath = /Saved a problem report:\\n(.+?\.zip)/.exec(no)![1];
    assert.ok(!unzip(readFileSync(noPath)).has('chart-spec.json'));
    assert.match(no, /version numbers\. No slide content\./);
  } finally { rmSync(home, { recursive: true, force: true }); }
});

test('report from the CLI: `addin report --spec <file>` prints where the zip is', () => {
  const home = tempHome();
  try {
    const spec = join(home, 'chart.json');
    writeFileSync(spec, SPEC);
    const out = execFileSync(process.execPath, ['--experimental-strip-types', '--no-warnings', new URL('../src/cli.ts', import.meta.url).pathname, 'addin', 'report', '--spec', spec],
      { encoding: 'utf8', env: { ...process.env, OPEN_CHARTS_HOME: home } });
    const path = /Saved a problem report:\n(.+\.zip)\n/.exec(out)?.[1];
    assert.ok(path && existsSync(path), out);
    assert.ok(unzip(readFileSync(path)).has('chart-spec.json'));
  } finally { rmSync(home, { recursive: true, force: true }); }
});

test('open / reveal a file: Finder on the Mac, Explorer on Windows (quotes as written)', () => {
  assert.deepEqual(fileCommand('/a b/r.zip', true, 'darwin'), ['open', ['-R', '/a b/r.zip']]);
  assert.deepEqual(fileCommand('/a b/s.pptx', false, 'darwin'), ['open', ['/a b/s.pptx']]);
  assert.deepEqual(fileCommand('C:\\a b\\r.zip', true, 'win32'), ['explorer.exe', ['/select,"C:\\a b\\r.zip"']]);
  assert.deepEqual(fileCommand('C:\\a b\\s.pptx', false, 'win32'), ['explorer.exe', ['"C:\\a b\\s.pptx"']]);
});

test('ribbon, VBA and bridge: Help group, first-run welcome once, report asks before adding the chart', () => {
  const built = buildAddin();
  const group = /<group id="grpHelp" label="Help"[^>]*>([\s\S]*?)<\/group>/.exec(built.ribbon)?.[1] ?? '';
  for (const [id, label] of HELP_BUTTONS) assert.match(group, new RegExp(`<button id="${id}" label="${label}"[^>]* onAction="${built.prefix}_onAction"/>`));
  assert.ok(built.ribbon.indexOf('grpHelp') > built.ribbon.indexOf('grpAgenda'), 'last on the tab, as the welcome says');
  assert.match(WELCOME, /Welcome, on the right of the open-charts tab/);

  const src = built.source;
  const body = (name: string) => { const a = src.search(new RegExp(`(Sub|Function) ${name}\\(`)); return src.slice(a, src.slice(a).search(/End (Sub|Function)/) + a); };
  assert.match(body('Dispatch'), /id = "btnWelcome" Then\s+ShowWelcome/);
  assert.match(body('Dispatch'), /id = "btnReport" Then\s+ReportProblem/);
  assert.match(body('ShowWelcome'), /Engine\("welcome"\)/);
  assert.match(body('ShowWelcome'), /MsgBox\(text, vbYesNo[^)]*\) = vbYes Then text = AnswerText\(Engine\("sample"\)\)/);
  // The chart's data only after a Yes; the request carries PowerPoint's version.
  assert.match(body('ReportProblem'), /= vbYes Then\s+arg = "spec"\s+spec = SpecOf\(chart\)\s+End If/);
  assert.match(body('ReportProblem'), /Engine\("report" & vbLf & arg & vbLf & vbLf & vbLf & "app=" & app & vbLf & spec\)/);
  // First run: the flag comes with the server start (Mac bridge / Windows settings.json), the dialog
  // waits for the first selection change (never onLoad, which an install's Apple event triggers).
  assert.match(body('StartEngine'), /gWelcome = \(InStr\(ignored, " welcome"\) > 0\)/);
  assert.match(body('StartEngine'), /gWelcome = InStr\(ignored, """welcomed"":true"\) = 0/);
  assert.match(body('OnSelectionChange'), /If gWelcome And Not gSilent Then\s+gWelcome = False/);
  assert.doesNotMatch(body('OCPREFIX_onLoad') + body('Auto_Open'), /ShowWelcome/);
  const bridge = readFileSync(new URL('../addin/applescript/open-charts.applescript', import.meta.url), 'utf8');
  assert.match(bridge, /grep -Eq '\\"welcomed\\": \*true' " & settingsFile\n\t\ton error\n\t\t\tset flags to flags & " welcome"/);
});

test('the bridge answers "OK welcome" until the welcome was shown', { skip: process.platform !== 'darwin' }, () => {
  const dir = tempHome();
  try {
    // A bridge whose engine is `true`: startServer only checks the settings.
    const src = join(dir, 'b.applescript'), scpt = join(dir, 'b.scpt');
    writeFileSync(src, readFileSync(new URL('../addin/applescript/open-charts.applescript', import.meta.url), 'utf8')
      .replace('__NODE__', '').replace('__CLI__', '/usr/bin/true').replace('__HOME__', dir));
    execFileSync('osacompile', ['-o', scpt, src]);
    const start = () => execFileSync('osascript', ['-e', `set s to load script (POSIX file ${JSON.stringify(scpt)})`, '-e', 'tell s to startServer("")'], { encoding: 'utf8' }).trim();
    assert.equal(start(), 'OK welcome');
    writeFileSync(join(dir, 'settings.json'), '{\n  "verbose": true,\n  "welcomed": true\n}\n');
    assert.equal(start(), 'OK verbose');
  } finally { rmSync(dir, { recursive: true, force: true }); }
});
