// The "Edit data" pop-up (src/web/editor.html ?popup=1; src/addin/popup.ts) against the stand-in host
// (host.ts: confirms each apply after 300 ms, one at a time, and logs what reached "PowerPoint").
import { expect, test, type APIRequestContext, type Page } from '@playwright/test';

const HOST = () => process.env.E2E_HOST_URL!;
const MOD = 'ControlOrMeta';
const chart = { type: 'clustered', categories: ['North', 'South', 'East'], series: [{ name: 'Sales', values: [10, 20, 30] }], options: { unit: 'm' } };
const cell = (page: Page, r: number, c: number) => page.locator(`#sheet input[data-r="${r}"][data-c="${c}"]`);
type Entry = { state: string; chart: { series: Array<{ values: number[] }> }; at: number };
const hostLog = async (request: APIRequestContext, id: string) => (await (await request.get(`${HOST()}/test/log/${id}`)).json()) as Entry[];
const applied = async (request: APIRequestContext, id: string) => (await hostLog(request, id)).filter((e) => e.state === 'applied').map((e) => e.chart.series[0].values);

let errors: string[] = [];
test.beforeEach(({ page }) => {
  errors = [];
  page.on('console', (m) => { if (m.type() === 'error') errors.push(m.text()); });
  page.on('pageerror', (e) => errors.push(e.message));
});
test.afterEach(() => { expect(errors).toEqual([]); });

/** The preloaded pop-up page (as the Mac helper keeps it), then a chart opened in it. */
async function openPopup(page: Page, request: APIRequestContext, prefix = 'pop') {
  const id = `${prefix}${Date.now().toString(16)}${Math.floor(Math.random() * 1e4)}`;
  expect((await request.post(`${HOST()}/api/session/${id}`, { data: { chart, title: 'Regional sales' } })).ok()).toBe(true);
  await page.setViewportSize({ width: 620, height: 360 });
  if (!page.url().includes('popup=1')) {
    await page.goto(`${HOST()}/?popup=1`);
    await expect(page.locator('body')).toHaveClass(/popup/);
  }
  const ms = await page.evaluate(async (sid) => {
    const t = performance.now();
    await (window as unknown as { ocPopupOpen(id: string): Promise<void> }).ocPopupOpen(sid);
    return performance.now() - t;
  }, id);
  await expect(cell(page, 0, 0)).toBeFocused();
  return { id, ms };
}

test('pop-up: compact datasheet layout, no preview, Done/Cancel/Options; opens a preloaded chart fast', async ({ page, request }) => {
  const { ms } = await openPopup(page, request);
  expect(ms, 'ocPopupOpen in the preloaded page').toBeLessThan(300);
  await expect(page.locator('#previewPanel')).toBeHidden();
  await expect(page.locator('#side')).toBeHidden();
  await expect(page.locator('#sheet')).toBeVisible();
  await expect(page.locator('#applyBtn')).toContainText('Done');
  await expect(page.locator('#cancelBtn')).toBeVisible();
  await page.click('#optionsBtn');
  await expect(page.locator('#side')).toBeVisible();
  await expect(page.locator('#typeBtn')).toBeVisible();
  await page.click('#optionsBtn');
  await expect(page.locator('#side')).toBeHidden();
  // Fits the window: no page scroll.
  expect(await page.evaluate(() => document.documentElement.scrollHeight <= innerHeight + 1)).toBe(true);
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth + 1)).toBe(true);
});

test('pop-up: an edit reaches the slide on its own ~400 ms after the last keystroke', async ({ page, request }) => {
  const { id } = await openPopup(page, request);
  await page.keyboard.type('15');
  const typed = Date.now();
  await page.waitForTimeout(200);
  expect(await applied(request, id), 'nothing before the debounce').toEqual([]);
  await expect.poll(() => applied(request, id)).toEqual([[15, 20, 30]]);
  const at = (await hostLog(request, id)).find((e) => e.state === 'applied')!.at;
  expect(at - typed).toBeGreaterThanOrEqual(350);
  expect(at - typed).toBeLessThan(1500);
  await expect(page.locator('#statusText')).toHaveText('Updated on slide');
  expect(await page.locator('#toast').isVisible()).toBe(false); // no toast per keystroke
});

test('pop-up: edits while an apply is in flight are coalesced into one more apply with the latest data', async ({ page, request }) => {
  const { id } = await openPopup(page, request);
  await page.keyboard.type('11');
  await page.waitForFunction(() => document.body.dataset.live === 'busy');
  // The host takes 300 ms per apply: these three edits land while the first is being drawn.
  await page.keyboard.press('Tab'); await page.keyboard.type('21');
  await page.waitForTimeout(450);
  await page.keyboard.press('Tab'); await page.keyboard.type('31');
  await page.waitForTimeout(100);
  await page.keyboard.press('Tab');
  await expect.poll(() => applied(request, id), { timeout: 5000 }).toContainEqual([11, 21, 31]);
  await page.waitForTimeout(900);
  const all = await applied(request, id);
  expect(all[0]).toEqual([11, 20, 30]);
  expect(all.at(-1)).toEqual([11, 21, 31]);
  expect(all.length, `applies: ${JSON.stringify(all)}`).toBeLessThanOrEqual(3);
  // Never two in flight: each apply was sent after the previous one was drawn (300 ms apart at least).
  const at = (await hostLog(request, id)).filter((e) => e.state === 'applied').map((e) => e.at);
  for (let i = 1; i < at.length; i++) expect(at[i] - at[i - 1]).toBeGreaterThanOrEqual(290);
});

test('pop-up: Escape puts back the chart it opened with, then cancels', async ({ page, request }) => {
  const { id } = await openPopup(page, request);
  await page.keyboard.type('99');
  await expect.poll(() => applied(request, id)).toEqual([[99, 20, 30]]);
  await page.keyboard.press('Escape');
  await expect.poll(() => page.evaluate(() => document.body.dataset.closed)).toBe('cancel');
  await expect.poll(async () => (await hostLog(request, id)).map((e) => e.state)).toEqual(['open', 'applied', 'applied', 'cancelled']);
  expect((await applied(request, id)).at(-1)).toEqual([10, 20, 30]);
});

test('pop-up: Escape without changes only cancels (the slide is not touched)', async ({ page, request }) => {
  const { id } = await openPopup(page, request);
  await page.keyboard.press('Escape');
  await expect.poll(async () => (await hostLog(request, id)).map((e) => e.state)).toEqual(['open', 'cancelled']);
});

test('pop-up: Cmd+Enter sends the last edit at once and closes, keeping it', async ({ page, request }) => {
  const { id } = await openPopup(page, request);
  await page.keyboard.type('42');
  await page.keyboard.press(`${MOD}+Enter`);
  await expect.poll(() => page.evaluate(() => document.body.dataset.closed)).toBe('done');
  await expect.poll(() => applied(request, id)).toEqual([[42, 20, 30]]);
  const states = (await hostLog(request, id)).map((e) => e.state);
  expect(states).not.toContain('cancelled');
});

test('pop-up: the preloaded page is reused for the next chart with a clean slate', async ({ page, request }) => {
  const first = await openPopup(page, request);
  await page.keyboard.type('77');
  await expect.poll(() => applied(request, first.id)).toEqual([[77, 20, 30]]);
  await page.keyboard.press(`${MOD}+Enter`);
  await expect.poll(() => page.evaluate(() => document.body.dataset.closed)).toBe('done');
  const second = await openPopup(page, request);
  expect(second.ms).toBeLessThan(300);
  await expect(cell(page, 0, 0)).toHaveValue('10');
  await expect(page.locator('#statusText')).toHaveText('Edits update the slide');
  await page.keyboard.press('Escape');
  await expect.poll(async () => (await hostLog(request, second.id)).map((e) => e.state)).toEqual(['open', 'cancelled']);
  expect(await applied(request, first.id)).toEqual([[77, 20, 30]]);
});

test('pop-up: an invalid cell is not sent to the slide; fixing it is', async ({ page, request }) => {
  const id = `pop${Date.now().toString(16)}bad`;
  const bad = { type: 'stacked', categories: ['a', 'b'], series: [{ name: 'S', values: [1] }] };
  await request.post(`${HOST()}/api/session/${id}`, { data: { chart: bad } });
  await page.setViewportSize({ width: 620, height: 360 });
  await page.goto(`${HOST()}/?session=${id}&popup=1`);
  await expect(cell(page, 0, 0)).toBeFocused();
  await page.keyboard.type('5');
  await page.waitForTimeout(800);
  expect(await applied(request, id)).toEqual([]);
  await cell(page, 0, 1).fill('2');
  await expect.poll(() => applied(request, id)).toEqual([[5, 2]]);
  errors = errors.filter((e) => !e.includes('400'));
});

test('element toolbar page: controls from the property table; a click writes the override and applies', async ({ page, request }) => {
  const id = `tb${Date.now().toString(16)}`;
  expect((await request.post(`${HOST()}/api/session/${id}`, { data: { chart, title: 'Sales' } })).ok()).toBe(true);
  await page.setViewportSize({ width: 320, height: 44 });
  await page.goto(`${HOST()}/toolbar`);
  await page.evaluate((sid) => (window as unknown as { ocPopupOpen(a: unknown): Promise<void> }).ocPopupOpen({ id: sid, roles: ['label:value:s0:c1'] }), id);
  await expect(page.getByTitle('Larger text')).toBeVisible();
  await expect(page.getByTitle('Text turned 90° right')).toBeVisible();
  await page.getByTitle('Text turned 90° right').click();
  await expect(page.getByTitle('Text turned 90° right')).toHaveAttribute('aria-pressed', 'true');
  await expect.poll(async () => (await hostLog(request, id)).filter((e) => e.state === 'applied').length).toBe(1);
  const last = (await hostLog(request, id)).at(-1)!.chart as unknown as { overrides: unknown };
  expect(last.overrides).toEqual({ 'label:value:s0:c1': { rotate: 90 } });
  // A bar takes only a fill.
  await page.evaluate((sid) => (window as unknown as { ocPopupOpen(a: unknown): Promise<void> }).ocPopupOpen({ id: sid, roles: ['bar:s0:c1'] }), id);
  await expect(page.getByTitle('Larger text')).toHaveCount(0);
  await page.getByTitle('Fill #A8461A').click();
  await expect.poll(async () => ((await hostLog(request, id)).at(-1)!.chart as unknown as { overrides: Record<string, unknown> }).overrides['bar:s0:c1']).toEqual({ fill: '#A8461A' });
});
