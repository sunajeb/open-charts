// Theme colours and a same-scale group: a CLI deck with two charts in one scale group opens in
// PowerPoint and gets another colour scheme, so the charts recolour by themselves (theme colours).
// Update chart on the waterfall re-lays out its partner; Edit data + Apply on the stacked chart keeps
// the group scale; resizing the stacked chart re-lays out the waterfall at the new scale.
import { readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { PPAM_DIR } from '../../../src/addin/cli.ts';
import { specKey } from '../../../src/addin/engine.ts';
import { buildDeck, type DeckSpec } from '../../../src/deck.ts';
import { createMeasurer } from '../../../src/fonts.ts';
import { readCharts } from '../../../src/ooxml/pptx-read.ts';
import { unzip } from '../../../src/ooxml/unzip.ts';
import { q, sleep } from '../powerpoint.ts';
import type { Ctx } from '../scenarios.ts';

export async function run(ctx: Ctx) {
  const { pp, check } = ctx;
  const pngs = join(ctx.out, 'png');
  const pair: DeckSpec = JSON.parse(readFileSync(join(ctx.root, 'examples', 'two-charts.json'), 'utf8'));
  pair.slides = [pair.slides[0]];
  const themedDeck = join(PPAM_DIR, 'live-themed.pptx'); // a folder PowerPoint's sandbox may read
  writeFileSync(themedDeck, buildDeck(pair, createMeasurer()).pptx);
  try {
    pp.open(themedDeck);
    pp.exportCopy(join(pngs, 'before'), 'png');
    const scheme = '/Applications/Microsoft PowerPoint.app/Contents/Resources/Office Themes/Theme Colors/Red Orange.xml';
    try {
      pp.tell(`apply theme color scheme slide 1 of active presentation file name ${q(scheme)}`);
    } catch (e) {
      if ((e as Error).name === 'LiveTimeout') throw e;
      pp.tell(`apply theme color scheme slide 1 of active presentation file name (POSIX file ${q(scheme)} as text)`);
    }
    pp.goTo(1);
    pp.macro('SelectChart', '2');
    const upd = pp.macro('Run', 'btnUpdate');
    check(/RESULT btnUpdate OK/.test(upd), `Update chart on the waterfall (${upd.split('\n').find((l) => /RESULT/.test(l)) ?? upd})`);
    check(/SYNC revenue 1 of 1 charts/.test(upd), 'Update chart re-lays out the group partner');

    // Edit data + Apply: the stacked chart's new data keeps its scale, the partner follows.
    const stacked = pp.charts(join(ctx.out, 'before-edit.pptx'))[0]?.charts.find((c) => c.spec && JSON.parse(c.spec).type === 'stacked');
    pp.macro('SelectChart', '1');
    const editOut = pp.macro('Run', 'btnEditData');
    if (check(!!stacked?.spec && /RESULT btnEditData OK/.test(editOut), `Edit data on the stacked chart (${editOut.split('\n').find((l) => /RESULT/.test(l)) ?? 'no RESULT line'})`) && stacked?.spec) {
      const spec = JSON.parse(stacked.spec);
      spec.series[0].values[2] = 45;
      const a = await pp.apply(specKey(stacked.spec), spec);
      ctx.time('apply (same-scale pair)', a.ms);
      check(/APPLIED \w+ OK/.test(a.applied) && a.outcome?.state === 'drawn', `Apply updated the chart (${a.applied || 'no APPLIED line'}, session ${a.outcome?.state}${a.outcome?.message ? `: ${a.outcome.message}` : ''})`);
      check(/SYNC revenue 1 of 1 charts/.test(a.log), 'Apply re-lays out the group partner');
    }

    // Resize the stacked chart (its own scale changes): the selection change re-lays it out and then
    // its partner, at the new group scale.
    pp.macro('SelectChart', '1');
    pp.macro('ResizeChart', '1', '0.8');
    const before = pp.logLines().length;
    pp.macro('SelectChart', '0');
    const synced = pp.waitForLog(/SYNC revenue/, before);
    sleep(500);
    check(/SYNC revenue 1 of 1 charts/.test(synced), `resizing a chart re-lays out its group partner (${synced || 'timeout'})`);

    const themedOut = join(ctx.out, 'themed.pptx');
    pp.exportCopy(themedOut, 'pptx');
    pp.exportCopy(join(pngs, 'after'), 'png');
    // The add-in drew theme colours as theme colours, resolved in the slide's new scheme.
    const bytes = new Uint8Array(readFileSync(themedOut));
    const slideXml = unzip(bytes).get('ppt/slides/slide1.xml')?.text() ?? '';
    for (const slot of ['accent1', 'tx1']) check(slideXml.includes(`<a:schemeClr val="${slot}"`), `the add-in's shapes use schemeClr ${slot}`);
    const charts = readCharts(bytes)[0]?.charts ?? [];
    check(charts.length === 2, `2 charts after the updates (${charts.map((c) => `${c.name} frame ${JSON.stringify(c.frame)}`).join('; ')})`);
    ctx.note(`PNGs before/after the colour scheme: ${pngs}`);
  } finally {
    rmSync(themedDeck, { force: true });
  }
}
