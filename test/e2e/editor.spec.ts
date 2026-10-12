// End-to-end test of every control in the web editor (src/web/editor.html).
import { expect, test, type Page } from '@playwright/test';
import { readFileSync, writeFileSync } from 'node:fs';
import { unzip } from './unzip.ts';
import { CHART_TYPES } from '../../src/core/spec.ts';

const BASE = () => process.env.E2E_BASE_URL!;
const PREVIEW_BUDGET_MS = 200;

let errors: string[] = [];
let expected400 = false; // set by tests that provoke an engine error on purpose

test.beforeEach(async ({ page }) => {
  errors = []; expected400 = false;
  page.on('console', (m) => { if (m.type() === 'error') errors.push(`console: ${m.text()}`); });
  page.on('pageerror', (e) => errors.push(`pageerror: ${e.message}`));
  await page.goto(BASE());
  await page.evaluate(() => localStorage.clear());
  await page.reload();
  await expect(page.locator('#preview svg')).toBeVisible();
  await settled(page);
});

test.afterEach(async ({ page }) => {
  const shown = await page.locator('#issues li.error').allTextContents();
  expect(errors.filter((e) => !(expected400 && e.includes("status of 400"))), "console errors").toEqual([]);
  expect(shown, 'editor error messages').toEqual([]);
});

/** Wait until every edit has been rendered. */
async function settled(page: Page) {
  await expect.poll(() => page.evaluate(() => {
    const p = document.getElementById('preview')!;
    return p.dataset.seq !== undefined && p.dataset.seq === document.body.dataset.seq;
  }), { timeout: 5000 }).toBe(true);
}

/** Run an edit, wait for the preview to re-render, and check it was fast. Returns the new SVG. */
async function edit(page: Page, action: () => Promise<unknown>): Promise<string> {
  const before = await page.evaluate(() => Number(document.body.dataset.seq));
  await action();
  await expect.poll(() => page.evaluate(() => Number(document.body.dataset.seq)), { message: 'edit registered' }).toBeGreaterThan(before);
  await settled(page);
  const ms = await page.evaluate(() => Number(document.getElementById('preview')!.dataset.ms));
  expect(ms, 'edit-to-preview latency (ms)').toBeLessThan(PREVIEW_BUDGET_MS);
  return page.locator('#preview').innerHTML();
}

const texts = (page: Page) => page.locator('#preview svg text').allTextContents();
const rects = (page: Page, fill: string) => page.locator(`#preview svg rect[fill="#${fill}"]`).count();
const cell = (page: Page, r: number, c: number) => page.locator(`#sheet input[data-r="${r}"][data-c="${c}"]`);
const deck = (page: Page) => page.evaluate(() => JSON.parse(localStorage.getItem('open-charts-deck-v1')!));
const NAVY = '1F3A5F';
const navyPaths = (page: Page) => page.locator(`#preview svg path[fill="#${NAVY}"]`).count();
/** Chart paths (areas, slices, wedges), not the arrowhead marker in <defs>. */
const paths = (page: Page) => page.locator('#preview svg path').evaluateAll((ps) => ps.filter((p) => !p.closest('marker')).length);
const visible = (page: Page, ids: string[]) => Promise.all(ids.map((id) => page.locator(`#${id}`).isVisible()));

/** Pick a chart type from the gallery. */
async function setType(page: Page, type: string) {
  await page.click('#typeBtn');
  await page.click(`#gallery .gal-item[data-type="${type}"]`);
}

async function paste(page: Page, r: number, c: number, text: string) {
  await cell(page, r, c).evaluate((el, t) => {
    const dt = new DataTransfer();
    dt.setData('text/plain', t);
    el.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
  }, text);
}

test('sample deck renders with title, totals, CAGR and a passing layout check', async ({ page }) => {
  const t = await texts(page);
  expect(t).toContain('Revenue grew 2.4x in four years, led by Enterprise');
  expect(t).toContain('2025E');
  expect(t.some((s) => /^\+?\d+(\.\d)?%$/.test(s.replace(/\s/g, '')) || /CAGR|%/.test(s))).toBe(true);
  expect(await rects(page, NAVY)).toBeGreaterThanOrEqual(5);
  // The only issues on the sample are the thin "Services" labels the engine hides; never overlaps.
  for (const li of await page.locator('#issues li').allTextContents()) expect(li).toMatch(/^Label hidden \(no room\): Services/);
  await expect(page.locator('#delSlide')).toBeDisabled();
  await expect(page.locator('#tabs .tab')).toHaveCount(1);
});

test('title edits the preview and the slide tab', async ({ page }) => {
  await edit(page, () => page.fill('#title', 'Margins doubled'));
  expect(await texts(page)).toContain('Margins doubled');
  await expect(page.locator('#tabs .tab')).toHaveText(['1. Margins doubled']);
});

test('fast typing: the preview ends on the last keystroke', async ({ page }) => {
  await page.fill('#title', '');
  await edit(page, () => page.locator('#title').pressSequentially('Typed one key at a time', { delay: 5 }));
  expect(await texts(page)).toContain('Typed one key at a time');
});

test('every chart type in the gallery renders a chart', async ({ page }) => {
  const types = await page.locator('#gallery .gal-item').evaluateAll((bs) => bs.map((b) => (b as HTMLElement).dataset.type!));
  expect(types.sort()).toEqual([...CHART_TYPES].sort());
  for (const type of types) {
    const svg = await edit(page, () => setType(page, type));
    expect((await deck(page)).slides[0].chart.type).toBe(type);
    if (type === 'line') expect(await page.locator('#preview svg line').count()).toBeGreaterThan(4);
    else if (['area', 'area100', 'pie', 'doughnut'].includes(type)) expect(await navyPaths(page), `${type} draws navy areas or slices`).toBeGreaterThan(0);
    else if (type === 'harvey') expect(await page.locator('#preview svg ellipse').count(), 'harvey draws balls').toBeGreaterThan(0);
    else if (type === 'scatter' || type === 'bubble') expect(await page.locator(`#preview svg ellipse[fill="#${NAVY}"]`).count(), `${type} draws points`).toBeGreaterThan(2);
    else if (type === 'gantt') expect(await page.locator('#preview svg polygon').count(), 'gantt draws milestone diamonds').toBeGreaterThan(0);
    else if (type === 'table') expect((await texts(page)).length, 'table draws header, row names and cells').toBeGreaterThan(8);
    else expect(await rects(page, NAVY), `${type} draws navy bars`).toBeGreaterThan(0);
    if (type === 'percent') expect(svg).toContain('%');
  }
});

test('unit suffix and decimals change the value labels', async ({ page }) => {
  await edit(page, () => page.fill('#unit', 'bn'));
  expect(await texts(page)).toContain('42bn');
  await edit(page, () => page.fill('#decimals', '1'));
  expect(await texts(page)).toContain('42.0bn');
  await edit(page, () => page.fill('#decimals', ''));
  expect(await texts(page)).toContain('42bn');
});

test('series labels: right, legend, none', async ({ page }) => {
  await edit(page, () => page.selectOption('#seriesLabels', 'none'));
  expect(await texts(page)).not.toContain('Enterprise');
  await edit(page, () => page.selectOption('#seriesLabels', 'right'));
  expect(await texts(page)).toContain('Enterprise');
  await edit(page, () => page.selectOption('#seriesLabels', 'legend'));
  expect(await texts(page)).toContain('Enterprise');
  expect((await deck(page)).slides[0].chart.options.seriesLabels).toBe('legend');
});

test('column gap changes bar width', async ({ page }) => {
  const width = async () => Number(await page.locator(`#preview svg rect[fill="#${NAVY}"]`).first().getAttribute('width'));
  const w0 = await width();
  await edit(page, () => page.fill('#gap', '1.5'));
  const w1 = await width();
  expect(w1).toBeLessThan(w0);
  await edit(page, () => page.fill('#gap', '0.1'));
  expect(await width()).toBeGreaterThan(w0);
});

test('totals and value axis checkboxes', async ({ page }) => {
  const total2021 = '97m';
  expect(await texts(page)).toContain(total2021);
  await edit(page, () => page.uncheck('#totals'));
  expect(await texts(page)).not.toContain(total2021);
  await edit(page, () => page.check('#totals'));
  expect(await texts(page)).toContain(total2021);
  const lines0 = await page.locator('#preview svg line').count();
  await edit(page, () => page.check('#valueAxis'));
  expect(await page.locator('#preview svg line').count()).toBeGreaterThan(lines0);
  await edit(page, () => page.uncheck('#valueAxis'));
  expect(await page.locator('#preview svg line').count()).toBe(lines0);
});

test('CAGR and difference arrows', async ({ page }) => {
  const cagr = async () => (await texts(page)).filter((t) => /%/.test(t));
  expect(await cagr()).toHaveLength(1);
  await edit(page, () => page.selectOption('#cagrFrom', ''));
  expect(await cagr()).toHaveLength(0);
  await edit(page, () => page.selectOption('#cagrFrom', '1'));
  await edit(page, () => page.selectOption('#cagrTo', '3'));
  expect((await deck(page)).slides[0].chart.options.cagr).toEqual({ from: 1, to: 3 });
  expect(await cagr()).toHaveLength(1);
  // Difference: 2021 total 97 → 2024 total 181 = +84.
  await edit(page, () => page.selectOption('#diffFrom', '0'));
  await edit(page, () => page.selectOption('#diffTo', '3'));
  expect((await texts(page)).some((t) => t.replace(/\s/g, '').includes('+84'))).toBe(true);
});

test('datasheet: edit a cell, rename category and series', async ({ page }) => {
  await edit(page, () => cell(page, 0, 0).fill('500'));
  expect(await texts(page)).toContain('500m');
  await edit(page, () => cell(page, -1, 0).fill('FY21'));
  expect(await texts(page)).toContain('FY21');
  await expect(page.locator('#cagrFrom option[value="0"]')).toHaveText('FY21');
  await edit(page, () => cell(page, 0, -1).fill('Large accounts'));
  expect(await texts(page)).toContain('Large accounts');
});

test('datasheet: Excel paste into the corner replaces the table', async ({ page }) => {
  await edit(page, () => paste(page, -1, -1, '\tA\tB\tC\r\nNorth\t1\t2\t3\r\nSouth\t4\t5\t6\r\n'));
  const c = (await deck(page)).slides[0].chart;
  expect(c.categories).toEqual(['A', 'B', 'C']);
  expect(c.series).toEqual([{ name: 'North', values: [1, 2, 3] }, { name: 'South', values: [4, 5, 6] }]);
  await expect(cell(page, 1, 2)).toHaveValue('6');
  const t = await texts(page);
  expect(t).toEqual(expect.arrayContaining(['A', 'B', 'C', 'North', 'South']));
});

test('datasheet: Excel paste into a body cell grows the table', async ({ page }) => {
  await edit(page, () => paste(page, 3, 4, '1,000\t(5)\n7%\t8\n'));
  const c = (await deck(page)).slides[0].chart;
  expect(c.categories).toHaveLength(6);
  expect(c.series).toHaveLength(5);
  expect(c.series[3].values.slice(4)).toEqual([1000, -5]);
  expect(c.series[4].values.slice(4)).toEqual([7, 8]);
  await expect(cell(page, 4, 5)).toHaveValue('8');
});

test('datasheet: + / − category and series', async ({ page }) => {
  const count = () => page.locator('#sheet input').count();
  const n0 = await count(); // (1 + 5) * (1 + 4)
  expect(n0).toBe(30);
  await edit(page, () => page.click('#addCat'));
  expect(await count()).toBe(35);
  expect(await texts(page)).toContain('Cat 6');
  await edit(page, () => page.click('#addSer'));
  expect(await count()).toBe(42);
  await edit(page, () => cell(page, 4, 5).fill('9'));
  expect((await deck(page)).slides[0].chart.series[4].values[5]).toBe(9);
  await edit(page, () => page.click('#delSer'));
  await edit(page, () => page.click('#delCat'));
  expect(await count()).toBe(30);
  expect(await texts(page)).not.toContain('Cat 6');
});

test('removing the category a CAGR arrow points at keeps the chart valid', async ({ page }) => {
  await edit(page, () => page.click('#delCat'));
  const o = (await deck(page)).slides[0].chart.options;
  expect(o.cagr === undefined || o.cagr.to < 4).toBe(true);
  expect(await rects(page, NAVY)).toBe(4);
});

test('slides: add, switch, delete', async ({ page }) => {
  await edit(page, () => page.click('#addSlide'));
  await expect(page.locator('#tabs .tab')).toHaveCount(2);
  await expect(page.locator('#tabs .tab').nth(1)).toHaveAttribute('aria-selected', 'true');
  await edit(page, () => page.fill('#title', 'Second'));
  await edit(page, () => setType(page, 'line'));
  await edit(page, () => page.locator('#tabs .tab').first().click());
  expect(await texts(page)).toContain('Revenue grew 2.4x in four years, led by Enterprise');
  await expect(page.locator('#typeBtn')).toHaveAttribute('data-type', 'stacked');
  await edit(page, () => page.locator('#tabs .tab').nth(1).click());
  expect(await texts(page)).toContain('Second');
  await expect(page.locator('#typeBtn')).toHaveAttribute('data-type', 'line');
  await edit(page, () => page.click('#delSlide'));
  await expect(page.locator('#tabs .tab')).toHaveCount(1);
  expect(await texts(page)).toContain('Revenue grew 2.4x in four years, led by Enterprise');
  expect((await deck(page)).slides).toHaveLength(1);
});

test('the deck survives a reload', async ({ page }) => {
  await edit(page, () => page.fill('#title', 'Persisted'));
  await page.reload();
  await settled(page);
  expect(await texts(page)).toContain('Persisted');
});

test('Export JSON downloads the deck', async ({ page }) => {
  await edit(page, () => page.fill('#title', 'Exported'));
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#exportBtn')]);
  expect(dl.suggestedFilename()).toBe('charts.json');
  const json = JSON.parse(readFileSync(await dl.path(), 'utf8'));
  expect(json.slides[0].title).toBe('Exported');
  expect(json.slides[0].chart.series[0].values).toEqual([42, 58, 75, 96, 118]);
});

test('Import JSON loads a deck or a bare chart through the file chooser', async ({ page }, info) => {
  const file = info.outputPath('import.json');
  writeFileSync(file, readFileSync(new URL('../../examples/mekko.json', import.meta.url)));
  const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.click('#importBtn')]);
  await edit(page, () => chooser.setFiles(file));
  await expect(page.locator('#tabs .tab')).toHaveCount(3);
  await expect(page.locator('#typeBtn')).toHaveAttribute('data-type', 'mekko');
  expect(await texts(page)).toContain('Market map: Enterprise is the largest pool and the one we lead');

  const bare = info.outputPath('bare.json');
  writeFileSync(bare, JSON.stringify({ type: 'line', categories: ['a', 'b'], series: [{ name: 'S', values: [1, 2] }] }));
  const [c2] = await Promise.all([page.waitForEvent('filechooser'), page.click('#importBtn')]);
  await edit(page, () => c2.setFiles(bare));
  await expect(page.locator('#tabs .tab')).toHaveCount(1);
  await expect(page.locator('#typeBtn')).toHaveAttribute('data-type', 'line');
});

test('Import JSON reports a broken file without losing the deck', async ({ page }, info) => {
  const bad = info.outputPath('bad.json');
  writeFileSync(bad, '{ not json');
  const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.click('#importBtn')]);
  await chooser.setFiles(bad);
  await expect(page.locator('#issues li.error')).toContainText("Couldn't read bad.json");
  await page.evaluate(() => document.querySelector('#issues')!.replaceChildren()); // acknowledged
  expect(await texts(page)).toContain('Revenue grew 2.4x in four years, led by Enterprise');
});

test('Download .pptx is a valid deck with one slide per tab', async ({ page }) => {
  await edit(page, () => page.click('#addSlide'));
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#downloadBtn')]);
  expect(dl.suggestedFilename()).toBe('charts.pptx');
  const buf = readFileSync(await dl.path());
  expect(buf.subarray(0, 2).toString()).toBe('PK');
  const files = unzip(buf);
  expect(files.has('[Content_Types].xml')).toBe(true);
  const s1 = files.get('ppt/slides/slide1.xml')!.toString('utf8');
  expect(s1).toContain('Revenue grew 2.4x in four years, led by Enterprise');
  expect(s1).toContain('<p:grpSp>');
  expect(files.has('ppt/slides/slide2.xml')).toBe(true);
});

test('issues panel: hidden labels, a clean pass, and non-numeric cells', async ({ page }) => {
  await expect(page.locator('#issues li.info').first()).toContainText('Label hidden (no room): Services / 2021');
  // Make Services big enough to label everywhere: the check passes.
  for (let j = 0; j < 5; j++) await edit(page, () => cell(page, 3, j).fill('30'));
  await expect(page.locator('#issues li')).toHaveText(['Layout check passed: no overlaps']);
  // A typo is flagged on the cell and in the panel instead of silently becoming empty.
  await edit(page, () => cell(page, 0, 1).fill('5o'));
  await expect(cell(page, 0, 1)).toHaveAttribute('aria-invalid', 'true');
  await expect(page.locator('#issues li').first()).toContainText('"5o" in Enterprise / 2022 is not a number');
  await edit(page, () => cell(page, 0, 1).fill('50'));
  await expect(cell(page, 0, 1)).toHaveAttribute('aria-invalid', 'false');
  await expect(page.locator('#issues li')).toHaveText(['Layout check passed: no overlaps']);
});

test('engine errors are shown, the preview is marked stale, and the editor recovers', async ({ page }, info) => {
  const file = info.outputPath('mismatch.json');
  expected400 = true;
  writeFileSync(file, JSON.stringify({ type: 'stacked', categories: ['a', 'b', 'c'], series: [{ name: 'S', values: [1, 2] }] }));
  const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.click('#importBtn')]);
  await chooser.setFiles(file);
  await expect(page.locator('#issues li.error')).toContainText(/values|categories/);
  await expect(page.locator('#previewPanel')).toHaveClass(/stale/);
  await expect(page.locator('#staleBanner')).toBeVisible();
  await edit(page, () => page.click('#delCat'));
  await expect(page.locator('#previewPanel')).not.toHaveClass(/stale/);
  expect(await texts(page)).toEqual(expect.arrayContaining(['a', 'b']));
});

test('transpose swaps categories and series', async ({ page }) => {
  await edit(page, () => page.click('#transpose'));
  const c = (await deck(page)).slides[0].chart;
  expect(c.categories).toEqual(['Enterprise', 'Mid-market', 'SMB', 'Services']);
  expect(c.series.map((s: { name: string }) => s.name)).toEqual(['2021', '2022', '2023', '2024', '2025E']);
  expect(c.series[4].values).toEqual([118, 60, 29, 9]);
  expect(c.options.cagr).toBeUndefined(); // pointed at the old categories
  await expect(cell(page, -1, 3)).toHaveValue('Services');
  expect(await texts(page)).toEqual(expect.arrayContaining(['Enterprise', '2025E']));
  await edit(page, () => page.click('#transpose'));
  expect((await deck(page)).slides[0].chart.series[0]).toEqual({ name: 'Enterprise', values: [42, 58, 75, 96, 118] });
});

test('reverse series flips the stack order', async ({ page }) => {
  // Value label of the bottom segment of the first column.
  const bottomLabel = async () => {
    const bars = await page.locator(`#preview svg rect[fill="#${NAVY}"]`).evaluateAll((els) => els.map((e) => ({ x: +e.getAttribute('x')!, y: +e.getAttribute('y')!, h: +e.getAttribute('height')! })));
    const b = bars.sort((p, q) => p.x - q.x)[0];
    const labels = await page.locator('#preview svg text tspan').evaluateAll((els) => els.map((e) => ({ y: +e.getAttribute('y')!, t: e.textContent })));
    return labels.filter((l) => l.y > b.y && l.y < b.y + b.h).map((l) => l.t);
  };
  expect(await bottomLabel()).toContain('42m');
  await edit(page, () => page.click('#reverse'));
  expect((await deck(page)).slides[0].chart.series.map((s: { name: string }) => s.name)).toEqual(['Services', 'SMB', 'Mid-market', 'Enterprise']);
  await expect(cell(page, 0, -1)).toHaveValue('Services');
  expect(await bottomLabel()).not.toContain('42m');
});

test('datasheet keyboard navigation: arrows, Enter, Shift+Enter', async ({ page }) => {
  await cell(page, 0, 0).click();
  await cell(page, 0, 0).selectText();
  await page.keyboard.press('ArrowRight');
  await expect(cell(page, 0, 1)).toBeFocused();
  await page.keyboard.press('ArrowDown');
  await expect(cell(page, 1, 1)).toBeFocused();
  await page.keyboard.press('Enter');
  await expect(cell(page, 2, 1)).toBeFocused();
  await page.keyboard.press('Shift+Enter');
  await expect(cell(page, 1, 1)).toBeFocused();
  await page.keyboard.press('ArrowLeft');
  await expect(cell(page, 1, 0)).toBeFocused();
  await page.keyboard.press('ArrowLeft');
  await expect(cell(page, 1, -1)).toBeFocused();
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('ArrowUp');
  await expect(cell(page, -1, -1)).toBeFocused();
  // Typing over a selected cell replaces it; the caret then moves inside the text before leaving the cell.
  await page.keyboard.press('ArrowRight');
  await expect(cell(page, -1, 0)).toBeFocused();
  await edit(page, () => page.keyboard.type('FY21'));
  await page.keyboard.press('ArrowLeft');
  await expect(cell(page, -1, 0)).toBeFocused();
  await page.keyboard.press('Home');
  await page.keyboard.press('ArrowLeft');
  await expect(cell(page, -1, -1)).toBeFocused();
  expect((await deck(page)).slides[0].chart.categories[0]).toBe('FY21');
});

test('− Category / − Series remove the selected cell\'s column / row', async ({ page }) => {
  await cell(page, 1, 2).click();
  await edit(page, () => page.click('#delCat'));
  let c = (await deck(page)).slides[0].chart;
  expect(c.categories).toEqual(['2021', '2022', '2024', '2025E']);
  expect(c.options.cagr).toEqual({ from: 0, to: 3 }); // re-indexed, still 2021 → 2025E
  await cell(page, 1, 0).click();
  await edit(page, () => page.click('#delSer'));
  c = (await deck(page)).slides[0].chart;
  expect(c.series.map((s: { name: string }) => s.name)).toEqual(['Enterprise', 'SMB', 'Services']);
});

test('switching type shows the controls that apply and keeps defaults sensible', async ({ page }) => {
  const visible = async () => ({
    gap: await page.locator('#gap').isVisible(),
    totals: await page.locator('#totals').isVisible(),
    cagr: await page.locator('#cagrFrom').isVisible(),
    mekko: await page.locator('#mekko').isVisible(),
    zero: await page.locator('#zeroBased').isVisible(),
    totalsAt: await page.locator('#totalsAt').isVisible(),
  });
  expect(await visible()).toEqual({ gap: true, totals: true, cagr: true, mekko: false, zero: false, totalsAt: false });
  await edit(page, () => setType(page, 'line'));
  expect(await visible()).toEqual({ gap: false, totals: false, cagr: false, mekko: false, zero: true, totalsAt: false });
  await expect(page.locator('#valueAxis')).toBeChecked(); // line default
  expect((await deck(page)).slides[0].chart.options.totals).toBeUndefined(); // stacked default dropped
  await edit(page, () => setType(page, 'clustered'));
  expect(await visible()).toEqual({ gap: true, totals: false, cagr: false, mekko: false, zero: false, totalsAt: false });
  await expect(page.locator('#valueAxis')).not.toBeChecked();
  await edit(page, () => setType(page, 'stacked'));
  await expect(page.locator('#totals')).toBeChecked();
  expect(await texts(page)).toContain('97m');
});

test('line: axis start auto / zero / fit data', async ({ page }) => {
  await edit(page, () => setType(page, 'line'));
  await edit(page, () => page.fill('#unit', ''));
  await edit(page, () => paste(page, -1, -1, '\t2021\t2022\t2023\nNRR\t112\t118\t121\n'));
  const lowest = async () => Math.min(...(await texts(page)).filter((t) => /^\d+$/.test(t)).map(Number));
  await expect(page.locator('#zeroBased')).toHaveValue('');
  expect(await lowest()).toBeGreaterThan(0); // narrow range: auto fits the data
  await edit(page, () => page.selectOption('#zeroBased', 'true'));
  expect((await deck(page)).slides[0].chart.options.zeroBased).toBe(true);
  expect(await lowest()).toBe(0);
  await edit(page, () => page.selectOption('#zeroBased', 'false'));
  expect(await lowest()).toBeGreaterThan(0);
  await edit(page, () => page.selectOption('#zeroBased', ''));
  expect((await deck(page)).slides[0].chart.options.zeroBased).toBeUndefined();
});

test('mekko: 100% and absolute heights', async ({ page }) => {
  await edit(page, () => setType(page, 'mekko'));
  await expect(page.locator('#mekko')).toHaveValue('percent');
  const heights = async () => (await page.locator(`#preview svg rect[fill="#${NAVY}"]`).evaluateAll((els) => els.map((e) => +e.getAttribute('height')!)));
  const pct = await heights();
  await edit(page, () => page.selectOption('#mekko', 'unit'));
  expect((await deck(page)).slides[0].chart.options.mekko).toBe('unit');
  expect(await heights()).not.toEqual(pct);
});

test('waterfall: first and last are totals by default, and total columns can be toggled', async ({ page }) => {
  await edit(page, () => setType(page, 'waterfall'));
  await edit(page, () => paste(page, -1, -1, '\tStart\tPrice\tCost\tEnd\nEBITDA\t100\t20\t-5\t\n'));
  const boxes = page.locator('#totalsAt input');
  await expect(boxes).toHaveCount(4);
  expect(await boxes.evaluateAll((bs) => bs.map((b) => (b as HTMLInputElement).checked))).toEqual([true, false, false, true]);
  expect(await texts(page)).toEqual(expect.arrayContaining(['100m', '115m']));
  await edit(page, () => boxes.nth(2).check());
  expect((await deck(page)).slides[0].chart.options.totalsAt).toEqual([0, 2, 3]);
  await expect(page.locator('#issues li.info')).toHaveCount(0);
  await edit(page, () => page.click('#addSer'));
  await expect(page.locator('#issues li.info')).toContainText('Waterfall shows the first series (EBITDA) only');
});

test('line: step toggle draws an ortho line and is dropped for other types', async ({ page }) => {
  await edit(page, () => setType(page, 'line'));
  await expect(page.locator('#step')).toBeVisible();
  await expect(page.locator('#step')).not.toBeChecked();
  const diagonal = async () => (await page.locator('#preview svg line[stroke-width="2"]').evaluateAll((ls) =>
    ls.filter((l) => l.getAttribute('x1') !== l.getAttribute('x2') && l.getAttribute('y1') !== l.getAttribute('y2')).length));
  expect(await diagonal()).toBeGreaterThan(0);
  await edit(page, () => page.check('#step'));
  expect((await deck(page)).slides[0].chart.options.step).toBe(true);
  expect(await diagonal(), 'every data segment is horizontal or vertical').toBe(0);
  await edit(page, () => setType(page, 'stacked'));
  await expect(page.locator('#step')).toBeHidden();
  expect((await deck(page)).slides[0].chart.options.step).toBeUndefined();
});

test('area and 100% area: filled bands, area controls only, 100% labels add up', async ({ page }) => {
  await edit(page, () => setType(page, 'area'));
  expect(await visible(page, ['totals', 'valueAxis', 'seriesLabels', 'gap', 'cagrFrom', 'step'])).toEqual([true, true, true, false, false, false]);
  expect(await paths(page)).toBe(4); // one band per series
  await edit(page, () => page.check('#totals'));
  expect(await texts(page)).toContain('216m');
  await edit(page, () => setType(page, 'area100'));
  const pcts = (await texts(page)).filter((t) => /^\d+%$/.test(t)).map((t) => parseInt(t));
  // 4 series x 5 years, rounded so each year adds up to exactly 100.
  expect(pcts.length % 5).toBe(0);
  expect(pcts.reduce((a, b) => a + b, 0)).toBe(100 * (pcts.length / 4));
});

test('pie and doughnut: one slice per category, shares add up to 100, first series only', async ({ page }) => {
  await edit(page, () => setType(page, 'pie'));
  expect(await visible(page, ['seriesLabels', 'valueAxis', 'gap', 'unit', 'totals'])).toEqual([true, false, false, false, false]);
  expect(await paths(page)).toBe(5);
  const shares = (await texts(page)).map((t) => t.match(/(\d+)%$/)?.[1]).filter(Boolean).map(Number);
  expect(shares.reduce((a, b) => a + b, 0)).toBe(100);
  await expect(page.locator('#issues li.info')).toContainText('Pie shows the first series (Enterprise) only');
  await edit(page, () => setType(page, 'doughnut'));
  await expect(page.locator('#totals')).toBeVisible();
  await edit(page, () => page.check('#totals'));
  expect(await texts(page)).toContain('389m'); // Enterprise across the five years, in the hole
});

test('harvey: switching seeds a 0–4 assessment grid of balls', async ({ page }) => {
  await edit(page, () => setType(page, 'harvey'));
  const c = (await deck(page)).slides[0].chart;
  expect(c.categories).toEqual(['Price', 'Features', 'Support', 'Integrations']);
  expect(await page.locator('#preview svg ellipse').count()).toBe(12);
  expect(await visible(page, ['unit', 'valueAxis', 'seriesLabels', 'decimals'])).toEqual([false, false, false, false]);
  // Edit a score: 4 → 1 turns a full ball into a quarter wedge.
  const before = await paths(page);
  await edit(page, () => cell(page, 0, 1).fill('1'));
  expect(await paths(page)).toBe(before + 1);
  await expect(page.locator('#issues li.ok')).toHaveCount(1);
});

// ---- Combination, percent axis, Mekko XY, scatter, bubble, table, Gantt ---------------------------

test('combo: the last series becomes a line; series can be switched between columns and line', async ({ page }) => {
  await edit(page, () => setType(page, 'combo'));
  const c = (await deck(page)).slides[0].chart;
  expect(c.series.at(-1).as).toBe('line');
  await expect(page.locator('#asLine')).toBeVisible();
  const markers = () => page.locator('#preview svg ellipse').count();
  expect(await markers()).toBe(5);
  await edit(page, () => page.locator('#asLine input').nth(0).check());
  expect(await markers()).toBe(10);
  await edit(page, () => page.locator('#asLine input').nth(0).uncheck());
  await edit(page, () => page.selectOption('#combo', 'clustered'));
  expect((await deck(page)).slides[0].chart.options.combo).toBe('clustered');
  expect(await rects(page, NAVY)).toBe(6); // five bars and the legend swatch
  await expect(page.locator('#issues li.error')).toHaveCount(0);
});

test('percent: "Values are %" keeps each column at its own total', async ({ page }) => {
  await edit(page, () => setType(page, 'percent'));
  await expect(page.locator('#percentData')).toBeVisible();
  await edit(page, () => page.check('#valueAxis'));
  const tops = async () => page.locator(`#preview svg rect[fill="#${NAVY}"]`).evaluateAll((els) => els.map((e) => +e.getAttribute('height')!));
  const before = await tops();
  await edit(page, () => page.check('#percentData'));
  expect((await deck(page)).slides[0].chart.options.percentData).toBe(true);
  expect(await texts(page)).toContain('42%'); // the entered value, not its share of the column
  expect(await texts(page)).toContain('100%'); // axis
  expect(await tops()).not.toEqual(before);
  await edit(page, () => setType(page, 'stacked'));
  expect((await deck(page)).slides[0].chart.options.percentData).toBeUndefined();
});

test('mekkoxy: first row is the width, order by height', async ({ page }) => {
  await edit(page, () => setType(page, 'mekkoxy'));
  await expect(page.locator('#sheetHint')).toContainText('first row is the column width');
  const xs = async () => page.locator('#preview svg rect[width]').evaluateAll((els) => els.filter((e) => e.getAttribute('fill') !== 'none').map((e) => +e.getAttribute('x')!));
  const before = await xs();
  await edit(page, () => page.selectOption('#sort', 'descending'));
  expect((await deck(page)).slides[0].chart.options.sort).toBe('descending');
  expect(await xs()).not.toEqual(before);
});

test('scatter and bubble: datasheet cells are points, axis titles', async ({ page }) => {
  await edit(page, () => setType(page, 'scatter'));
  await expect(cell(page, 0, 0)).toHaveValue('42; 31');
  await expect(page.locator('#valueAxis')).toBeHidden();
  await edit(page, () => cell(page, 0, 0).fill('10; 20'));
  expect((await deck(page)).slides[0].chart.series[0].values[0]).toEqual([10, 20]);
  await edit(page, () => cell(page, 0, 1).fill('oops'));
  await expect(page.locator('#issues')).toContainText('is not "x; y"');
  await edit(page, () => cell(page, 0, 1).fill('12; 25'));
  await edit(page, () => page.fill('#xTitle', 'Seats'));
  await edit(page, () => page.fill('#yTitle', 'Price'));
  expect(await texts(page)).toEqual(expect.arrayContaining(['Seats', 'Price', '2021']));
  await edit(page, () => setType(page, 'bubble'));
  await expect(cell(page, 0, 0)).toHaveValue('10; 20; 1');
  await edit(page, () => cell(page, 0, 0).fill('10; 20; 9'));
  const radii = await page.locator(`#preview svg ellipse[fill="#${NAVY}"]`).evaluateAll((els) => els.map((e) => +e.getAttribute('rx')!));
  expect(Math.max(...radii) / Math.min(...radii)).toBeCloseTo(3, 1); // area ∝ size: √9 = 3
  await edit(page, () => setType(page, 'stacked'));
  expect((await deck(page)).slides[0].chart.series.map((s: { name: string }) => s.name)).toEqual(['X', 'Y', 'Size']);
});

test('table: header, row names and formatted numbers; text cells are kept', async ({ page }) => {
  await edit(page, () => setType(page, 'table'));
  await expect(page.locator('#valueAxis')).toBeHidden();
  expect(await texts(page)).toEqual(expect.arrayContaining(['2021', '2025E', 'Enterprise', '42m', '118m']));
  await edit(page, () => cell(page, 0, 0).fill('n/a'));
  expect(await texts(page)).toContain('n/a');
  expect((await deck(page)).slides[0].chart.series[0].values[0]).toBe('n/a');
  await edit(page, () => setType(page, 'stacked'));
  expect((await deck(page)).slides[0].chart.series[0].values[0]).toBeNull();
});

test('gantt: dates in the datasheet drive bars, milestones and the calendar scale', async ({ page }) => {
  await edit(page, () => setType(page, 'gantt'));
  await expect(page.locator('#unit')).toBeHidden();
  await expect(cell(page, 0, 0)).toHaveValue('2026-01-05');
  expect(await page.locator('#preview svg polygon').count()).toBe(4);
  expect(await texts(page)).toEqual(expect.arrayContaining(['2026', 'Jan', 'Feb 2', 'Diagnostic']));
  const width = async () => Number(await page.locator(`#preview svg rect[fill="#${NAVY}"]`).first().getAttribute('width'));
  const w0 = await width();
  await edit(page, () => cell(page, 1, 0).fill('2026-03-02'));
  expect(await width()).toBeGreaterThan(w0 * 1.5);
  await edit(page, () => setType(page, 'stacked'));
  expect((await deck(page)).slides[0].chart.series[0].values[0]).toBe(42);
});
