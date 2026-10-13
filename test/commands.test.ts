import { test } from 'node:test';
import assert from 'node:assert/strict';
import { newScaleGroup, runCommand, selectedCategories } from '../src/core/commands.ts';
import { samples } from '../src/core/samples.ts';
import type { ChartSpec } from '../src/core/types.ts';
import { handleRequest, parseRequest } from '../src/addin/engine.ts';
import { buildAddin, CHART_TOOLS, CONTEXT_MENUS } from '../src/addin/package.ts';
import { createMeasurer } from '../src/fonts.ts';

/** The stacked sample without its CAGR arrow. */
const stacked = (): ChartSpec => { const s = structuredClone(samples.stacked as ChartSpec); delete s.options?.cagr; return s; };
const ok = <T>(r: { spec: T } | { error: string }): T => { assert.ok('spec' in r, 'error' in r ? r.error : ''); return (r as { spec: T }).spec; };
const err = (r: { error?: string } | object): string => { assert.ok('error' in r); return (r as { error: string }).error; };

test('commands: totals toggle against the type default, never changing the input', () => {
  const s = stacked();
  const before = JSON.stringify(s);
  const off = ok(runCommand(s, 'totals'));
  assert.equal(off.options?.totals, false, 'stacked shows totals by default: the first click hides them');
  assert.equal(ok(runCommand(off, 'totals')).options?.totals, true);
  assert.equal(JSON.stringify(s), before);
  assert.match(err(runCommand(structuredClone(samples.line as ChartSpec), 'totals')), /not available/);
});

test('commands: CAGR / difference arrows span the selected columns, else first to last; again removes', () => {
  const s = stacked();
  const n = s.categories.length;
  const a = ok(runCommand(s, 'cagr'));
  assert.deepEqual(a.options?.cagr, { from: 0, to: n - 1 });
  assert.equal(ok(runCommand(a, 'cagr')).options?.cagr, undefined, 'second click removes it');
  const sel = ok(runCommand(s, 'difference', ['bar:s0:c3', 'label:value:s1:c1', 'bar:s1:c1']));
  assert.deepEqual(sel.options?.difference, { from: 1, to: 3 });
  // Other columns selected than the arrow's: the arrow moves there (keeping its mode).
  const withMode = { ...sel, options: { ...sel.options, difference: { from: 1, to: 3, mode: 'relative' as const } } };
  assert.deepEqual(ok(runCommand(withMode, 'difference', ['bar:s0:c0', 'bar:s0:c2'])).options?.difference, { from: 0, to: 2, mode: 'relative' });
  assert.equal(ok(runCommand(withMode, 'difference', ['bar:s0:c1', 'bar:s0:c3'])).options?.difference, undefined, 'same columns: toggles off');
  assert.match(err(runCommand(structuredClone(samples.waterfall as ChartSpec), 'cagr')), /not available/);
  assert.deepEqual(selectedCategories(['label:total:c4', 'bar:s2:c0', 'legend:s1']), [0, 4]);
});

test('commands: reset label position unpins the selected labels, or every pinned label', () => {
  const s: ChartSpec = { ...stacked(), overrides: { 'label:total:c1': { pin: { x: 0.2, y: 0.1 }, bold: true }, 'label:value:s0:c0': { pin: { x: 0.5, y: 0.5 } } } };
  const one = ok(runCommand(s, 'unpin', ['label:total:c1']));
  assert.deepEqual(one.overrides, { 'label:total:c1': { bold: true }, 'label:value:s0:c0': { pin: { x: 0.5, y: 0.5 } } });
  assert.deepEqual(ok(runCommand(s, 'unpin')).overrides, { 'label:total:c1': { bold: true } });
  assert.match(err(runCommand(stacked(), 'unpin')), /No label/);
  assert.match(err(runCommand(s, 'unpin', ['label:value:s1:c1'])), /already/);
});

test('commands: same-scale group in and out; group names do not collide', () => {
  const g = ok(runCommand(stacked(), 'scaleGroup', [], 'scale1'));
  assert.equal(g.options?.scaleGroup, 'scale1');
  assert.equal(ok(runCommand(g, 'scaleGroup', [], '')).options?.scaleGroup, undefined);
  assert.match(err(runCommand(structuredClone(samples.pie as ChartSpec), 'scaleGroup', [], 'x')), /not available/);
  assert.equal(newScaleGroup(['scale1', 'revenue', 'scale3']), 'scale2');
  assert.match(err(runCommand(stacked(), 'bogus')), /Unknown/);
});

test('engine op "command": roles and group parsed, chart re-laid out in place with the new options', async () => {
  const spec = JSON.stringify(samples.stacked);
  const r = parseRequest(`command\ncagr\n10,20,600,300\n\nroles=bar:s0:c1,bar:s0:c3\ngroup=g1\n${JSON.stringify(stacked())}`);
  assert.deepEqual([r.op, r.arg, r.roles, r.group], ['command', 'cagr', ['bar:s0:c1', 'bar:s0:c3'], 'g1']);
  assert.equal(r.context, '10,20,600,300\n\n', 'roles= / group= are not layout context');
  const out = await handleRequest(r, createMeasurer());
  assert.doesNotMatch(out, /^E\t/m, out.slice(0, 300));
  const tag = /^T\tOC_SPEC\t(.*)$/m.exec(out)?.[1];
  assert.ok(tag);
  const drawn = JSON.parse(tag!.replace(/\\(.)/g, (_, c) => (c === 't' ? '\t' : c === 'n' ? '\n' : c)));
  assert.deepEqual(drawn.options.cagr, { from: 1, to: 3 });
  assert.match(out, /^F\t10,20,600,300$/m, 'same frame');
  const bad = await handleRequest(parseRequest(`command\ntotals\n10,20,600,300\n\n${JSON.stringify(samples.line)}`), createMeasurer());
  assert.match(bad, /^E\tTotals are not available/m);
  const grouped = await handleRequest(parseRequest(`command\nscaleGroup\n10,20,600,300\n\ngroup=scale1\n${spec}`), createMeasurer());
  assert.match(grouped, /^T\tOC_SGROUP\tscale1$/m);
});

test('ppam: Chart tools group and right-click items for every chart command, type lists from OPTIONS', () => {
  const { ribbon, source, prefix } = buildAddin();
  assert.match(ribbon, /<group id="grpChartTools" label="Chart tools"/);
  for (const t of CHART_TOOLS) {
    assert.ok(ribbon.includes(`<button id="${t.id}" `), t.id);
    assert.ok(source.includes(`ElseIf id = "${t.id}" Then`), `dispatcher runs ${t.id}`);
  }
  assert.ok(ribbon.indexOf('<contextMenus>') > ribbon.indexOf('</ribbon>'), 'contextMenus after ribbon (schema order)');
  for (const menu of CONTEXT_MENUS) {
    const block = new RegExp(`<contextMenu idMso="${menu}">(.*?)</contextMenu>`).exec(ribbon)?.[1] ?? '';
    for (const id of ['btnEditData', 'btnUpdate', 'btnSameScale', 'btnResetLabel']) assert.ok(block.includes(`id="ctx${menu.replace('ContextMenu', '')}_${id}"`), `${menu} ${id}`);
    assert.ok([...block.matchAll(/<button [^>]*>/g)].every((b) => b[0].includes(`getVisible="${prefix}_getVisible"`)), `${menu}: every item only on our charts`);
  }
  assert.ok(new Set([...ribbon.matchAll(/ id="([^"]+)"/g)].map((m) => m[1])).size === [...ribbon.matchAll(/ id="([^"]+)"/g)].length, 'control ids unique');
  assert.ok(source.includes(`Sub ${prefix}_getVisible(`));
  assert.match(source, /Private Const TOTALS_TYPES As String = "\|stacked\|percent\|[a-z|]*"/);
  assert.match(source, /Private Const SCALE_TYPES As String = "\|stacked\|clustered\|waterfall\|line\|area\|combo\|"/);
  assert.match(source, /CallByName Application, "StartNewUndoEntry", VbMethod/);
});
