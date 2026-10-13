// The Office.js add-in in a real browser (Chromium), against a recording mock of office.js
// (office-mock.js): the browser engine's parity with Node, and the task pane's contract with
// PowerPoint — insert (Path B), select → edit → Apply in place (position, size, z-order), Revert,
// the resize banner, Path A fallbacks. Live PowerPoint is covered by tools/officejs-live.ts.
import { expect, test, webkit, type Page } from '@playwright/test';
import { mkdtempSync, readFileSync, rmSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { extname, join } from 'node:path';
import { buildOfficeJs } from '../../scripts/build-officejs.ts';
import { measurerFromFonts } from '../../src/core/text.ts';
import { samples } from '../../src/core/samples.ts';
import type { ChartSpec } from '../../src/core/types.ts';
import { buildChart, chartDeck, CHART_FRAME, toBase64 } from '../../src/officejs/engine.ts';
import { writePptx } from '../../src/render/pptx.ts';
import { readCharts } from '../../src/ooxml/pptx-read.ts';

const ORIGIN = 'https://oc.test';
const OFFICE_JS = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js';
const font = (f: string) => { const b = readFileSync(new URL(`../../vendor/fonts/${f}`, import.meta.url)); return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer; };
const m = measurerFromFonts(font('LiberationSans-Regular.ttf'), font('LiberationSans-Bold.ttf'));
const MOCK = readFileSync(new URL('./office-mock.js', import.meta.url), 'utf8');
const TYPES: Record<string, string> = { '.html': 'text/html', '.js': 'text/javascript', '.ttf': 'font/ttf', '.png': 'image/png', '.xml': 'application/xml' };

let dist = '';
test.beforeAll(async () => {
  dist = mkdtempSync(join(tmpdir(), 'oc-officejs-e2e-'));
  await buildOfficeJs({ out: dist, quiet: true, base: ORIGIN });
});
test.afterAll(() => { rmSync(dist, { recursive: true, force: true }); });

let errors: string[] = [];
async function serve(page: Page, config: object = {}) {
  errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('console', (msg) => { if (msg.type() === 'error') errors.push(msg.text()); });
  await page.route(OFFICE_JS, (r) => r.fulfill({ contentType: 'text/javascript', body: `window.__OC_MOCK_CONFIG = ${JSON.stringify(config)};\n${MOCK}` }));
  await page.route(`${ORIGIN}/**`, async (r) => {
    const url = new URL(r.request().url());
    if (url.pathname === '/api/log') return r.fulfill({ status: 204 });
    if (url.pathname === '/parity.html') return r.fulfill({ contentType: 'text/html', body: '<!doctype html><title>parity</title>' });
    const file = join(dist, url.pathname);
    if (!existsSync(file)) return r.fulfill({ status: 404 });
    return r.fulfill({ contentType: TYPES[extname(file)] ?? 'application/octet-stream', body: readFileSync(file) });
  });
}

const log = (page: Page) => page.evaluate(() => (window as unknown as { OC_LOG: string[] }).OC_LOG);
const calls = (page: Page) => page.evaluate(() => (window as unknown as { __OC_CALLS: string[] }).__OC_CALLS);
const waitLog = (page: Page, prefix: string) => expect.poll(async () => (await log(page)).find((l) => l.startsWith(prefix)) ?? '', { timeout: 10_000 }).not.toBe('');

async function openPane(page: Page, config: object = {}) {
  await serve(page, config);
  await page.goto(`${ORIGIN}/taskpane.html?host=office`);
  await waitLog(page, 'ready ');
  await expect(page.locator('#preview svg')).toBeVisible();
}

/** A chart group as the CLI / .ppam writes it: tags of tag contract v1, an invisible frame rect, its shapes. */
function chartGroup(spec: ChartSpec, frame = CHART_FRAME, id = '7') {
  const b = buildChart(spec, frame, m);
  const tags = Object.fromEntries(b.content.tags!);
  return { id, name: b.content.name, type: 'Group', left: frame.x, top: frame.y, width: frame.w, height: frame.h, tags, children: [{ id: `${id}1`, name: 'Frame', left: frame.x, top: frame.y, width: frame.w, height: frame.h, tags: { OC_ROLE: 'frame' } }] };
}

test.afterEach(() => { expect(errors, 'page errors').toEqual([]); });

/** Every sample laid out and written in the page by dist/officejs/engine.js, compared with Node. */
async function parity(page: Page) {
  await serve(page);
  await page.goto(`${ORIGIN}/parity.html`);
  const browser = await page.evaluate(async ({ frame }) => {
    const e = await import('/engine.js');
    const m = await e.loadMeasurer(location.href);
    const out: Record<string, { scene: string; pptx: string }> = {};
    for (const [type, spec] of Object.entries(e.samples)) {
      const b = e.buildChart(spec, frame, m);
      out[type] = { scene: JSON.stringify(b.scene), pptx: e.toBase64(e.chartDeck(b, spec)) };
    }
    return out;
  }, { frame: CHART_FRAME });
  const types = Object.keys(samples);
  expect(Object.keys(browser).sort()).toEqual(types.sort());
  for (const type of types) {
    const spec = samples[type as keyof typeof samples]!;
    const b = buildChart(spec, CHART_FRAME, m);
    expect(browser[type].scene, `${type} scene`).toBe(JSON.stringify(b.scene));
    expect(browser[type].pptx, `${type} .pptx`).toBe(toBase64(chartDeck(b, spec)));
  }
}

test('the browser bundle lays out every sample byte-identically to Node (Chromium, same Liberation Sans)', async ({ page }) => parity(page));

// PowerPoint for Mac runs add-ins in WKWebView: the same checks in WebKit, when Playwright's WebKit is installed.
const hasWebKit = (() => { try { return existsSync(webkit.executablePath()); } catch { return false; } })();
test('WebKit: byte-identical layout, and the task pane inserts a chart', async () => {
  test.skip(!hasWebKit, 'WebKit not installed (npx playwright install webkit)');
  const browser = await webkit.launch();
  try {
    await parity(await browser.newPage());
    const page = await browser.newPage();
    await openPane(page);
    await page.click('#applyBtn');
    await waitLog(page, 'insert ');
    expect((await log(page)).find((l) => l.startsWith('insert '))).toContain('"path":"B"');
  } finally {
    await browser.close();
  }
});

test('outside PowerPoint the pane previews but cannot insert', async ({ page }) => {
  await serve(page);
  await page.route(OFFICE_JS, (r) => r.fulfill({ contentType: 'text/javascript', body: '' }));
  await page.goto(`${ORIGIN}/taskpane.html?host=office`);
  await expect(page.locator('#preview svg')).toBeVisible();
  await expect(page.locator('#officeBannerText')).toContainText('Open this pane from PowerPoint');
});

test('Insert draws the chart on the current slide as one tagged group (Path B, 2 syncs)', async ({ page }) => {
  await openPane(page);
  await expect(page.locator('#applyLabel')).toHaveText('Insert on slide');
  const syncs0 = await page.evaluate(() => (window as unknown as { __OC_SYNCS: () => number }).__OC_SYNCS());
  const before = (await calls(page)).length;
  await page.click('#applyBtn');
  await waitLog(page, 'insert ');
  const line = (await log(page)).find((l) => l.startsWith('insert '))!;
  expect(line).toContain('"path":"B"');
  const made = (await calls(page)).slice(before);
  expect(made.filter((c) => c === 'addGroup ' + made.filter((x) => /^add(GeometricShape|Line|TextBox)/.test(x)).length)).toHaveLength(1);
  const shapes = await page.evaluate(() => (window as any).__OC_MOCK.shapes().map((s: any) => ({ type: s.type, name: s.name, tags: Object.fromEntries(s.tags.items.map((t: any) => [t.key, t.value])), kids: s.children?.length, frame: [s.left, s.top, s.width, s.height], first: s.children?.[0]?.name })));
  expect(shapes).toHaveLength(1);
  const g = shapes[0];
  expect(g.type).toBe('Group');
  expect(g.name).toBe('Stacked column chart');
  expect(g.first).toBe('Frame');
  expect(Object.keys(g.tags)).toEqual(expect.arrayContaining(['OC_V', 'OC_ENGINE', 'OC_SPEC', 'OC_FRAME', 'OC_FONT']));
  expect(g.tags.OC_FRAME).toBe('48,96,864,408');
  expect(g.frame).toEqual([48, 96, 864, 408]);
  // Draw = 2 syncs (create + group + tags, then select); the theme read and re-binding add reads.
  expect(made.filter((c) => c === 'sync').length).toBeLessThanOrEqual(8);
  expect(await page.evaluate(() => (window as unknown as { __OC_SYNCS: () => number }).__OC_SYNCS()) - syncs0).toBeGreaterThan(0);
  // The pane is now bound to the chart it drew.
  await expect(page.locator('#applyLabel')).toHaveText('Apply to slide');
  await expect(page.locator('#sessionTitle')).toContainText('Editing Stacked column chart');
});

test('select a chart → its data loads; Apply replaces it in place (position, size, z-order); Revert restores it', async ({ page }) => {
  const spec = { ...samples.clustered!, categories: ['North', 'South', 'East', 'West'] };
  const frame = { x: 100, y: 120, w: 500, h: 300 };
  const chart = chartGroup(spec, frame);
  const above = { id: '9', name: 'Sticker', left: 0, top: 0, width: 50, height: 50 };
  await openPane(page, { slides: [{ id: '256#', shapes: [{ id: '2', name: 'Title', left: 0, top: 0, width: 900, height: 40 }, chart, above] }] });
  await page.evaluate(() => (window as any).__OC_MOCK.select(['7']));
  await waitLog(page, 'bound ');
  await expect(page.locator('#sessionTitle')).toContainText('Editing Clustered column chart');
  await expect.poll(() => page.locator('#sheet input').evaluateAll((els) => els.map((e) => (e as HTMLInputElement).value))).toContain('North');
  const cell = page.locator('#sheet input[data-r="0"][data-c="0"]');
  await cell.fill('99');
  await cell.press('Enter');
  await page.click('#applyBtn');
  await waitLog(page, 'apply ');
  const after = await page.evaluate(() => (window as any).__OC_MOCK.shapes().map((s: any) => ({ id: s.id, name: s.name, frame: s.tags.items.find((t: any) => t.key === 'OC_FRAME')?.value, spec: s.tags.items.find((t: any) => t.key === 'OC_SPEC')?.value })));
  // Old group gone, new one at the same z-order (between title and sticker), same frame and name.
  expect(after.map((s: any) => s.name)).toEqual(['Title', 'Clustered column chart', 'Sticker']);
  expect(after[1].id).not.toBe('7');
  expect(after[1].frame).toBe('100,120,500,300');
  expect(JSON.parse(after[1].spec).series[0].values[0]).toBe(99);
  expect(await page.evaluate(() => (window as any).__OC_MOCK.selected().shapeIds)).toEqual([after[1].id]);
  // Revert: the chart as it was before the Apply.
  await expect(page.locator('#revertBtn')).toBeVisible();
  await page.click('#revertBtn');
  await waitLog(page, 'apply ');
  await expect.poll(async () => page.evaluate(() => JSON.parse((window as any).__OC_MOCK.shapes()[1].tags.items.find((t: any) => t.key === 'OC_SPEC').value).series[0].values[0])).toBe(spec.series[0].values[0]);
  const names = await page.evaluate(() => (window as any).__OC_MOCK.shapes().map((s: any) => s.name));
  expect(names).toEqual(['Title', 'Clustered column chart', 'Sticker']);
  await expect(page.locator('#revertBtn')).toBeHidden();
});

test('a resized chart shows the banner; Re-layout plans it for the new size', async ({ page }) => {
  const chart = chartGroup(samples.waterfall!, { x: 60, y: 100, w: 600, h: 360 });
  await openPane(page, { slides: [{ id: '256#', shapes: [chart] }] });
  // The user drags the group to half its width.
  await page.evaluate(() => { const g = (window as any).__OC_MOCK.shapes()[0]; g.width = 300; (window as any).__OC_MOCK.select([g.id]); });
  await waitLog(page, 'bound ');
  await expect(page.locator('#officeBanner')).toBeVisible();
  await expect(page.locator('#relayoutBtn')).toBeVisible();
  await page.click('#relayoutBtn');
  await waitLog(page, 'apply ');
  const g = await page.evaluate(() => { const s = (window as any).__OC_MOCK.shapes()[0]; return { frame: s.tags.items.find((t: any) => t.key === 'OC_FRAME').value, width: s.width }; });
  expect(g.frame).toBe('60,100,300,360');
  await expect(page.locator('#officeBanner')).toBeHidden();
});

test('a chart PowerPoint JS cannot draw goes in through the slide export (Path A in place, 1.8+)', async ({ page }) => {
  const chart = chartGroup(samples.area!, { x: 60, y: 100, w: 600, h: 360 });
  // What exportAsBase64 returns: the slide with the old chart renamed to the marker.
  const exported = writePptx([{ title: 'Area', scenes: [{ ...buildChart(samples.area!, { x: 60, y: 100, w: 600, h: 360 }, m).content, name: 'open-charts:replace-me' }] }]);
  await openPane(page, { slides: [{ id: '256#', shapes: [chart] }], exportBase64: toBase64(exported) });
  await page.evaluate(() => (window as any).__OC_MOCK.select(['7']));
  await waitLog(page, 'bound ');
  await page.click('#applyBtn');
  await waitLog(page, 'apply ');
  expect((await log(page)).find((l) => l.startsWith('apply '))).toContain('"path":"A-inplace"');
  const c = await calls(page);
  expect(c).toContain('exportAsBase64');
  expect(c).toContain('insertSlidesFromBase64 UseDestinationTheme 256#');
  expect(c).toContain('deleteSlide 256#');
  expect(await page.evaluate(() => (window as any).__OC_EXPORTED_NAMES)).toEqual(['open-charts:replace-me']);
  const inserted = await page.evaluate(() => (window as any).__OC_INSERTED[0]);
  const [slide] = readCharts(Uint8Array.from(Buffer.from(inserted, 'base64')));
  expect(slide.charts.map((x) => x.name)).toEqual(['Stacked area chart']);
  expect(slide.charts[0].frame).toEqual({ x: 60, y: 100, w: 600, h: 360 });
});

test('before PowerPointApi 1.8 Insert puts the chart on a new slide (Path A)', async ({ page }) => {
  await openPane(page, { maxSet: '1.5' });
  await page.click('#applyBtn');
  await waitLog(page, 'insert ');
  expect((await log(page)).find((l) => l.startsWith('insert '))).toContain('"path":"A-newslide"');
  expect(await calls(page)).toContain('insertSlidesFromBase64 UseDestinationTheme 256#');
  const inserted = await page.evaluate(() => (window as any).__OC_INSERTED[0]);
  const [slide] = readCharts(Uint8Array.from(Buffer.from(inserted, 'base64')));
  expect(slide.charts.map((x) => x.name)).toEqual(['Stacked column chart']);
  await expect(page.locator('#toast')).toContainText('new slide');
});

test('the task pane reports the PowerPointApi matrix (spike OJ-S1) and uses the slide theme (1.10)', async ({ page }) => {
  await openPane(page, { theme: { Accent1: '#AA0000' } });
  const ready = (await log(page)).find((l) => l.startsWith('ready '))!;
  expect(JSON.parse(ready.slice(6)).PowerPointApi['1.10']).toBe(true);
  await page.click('#applyBtn');
  await waitLog(page, 'insert ');
  const fills = await page.evaluate(() => (window as any).__OC_MOCK.shapes()[0].children.map((s: any) => s.fill.color));
  expect(fills).toContain('#AA0000');
});
