// Label-candidate placer (src/core/labels.ts) and the thin-segment rules of every chart that uses it:
// boxed on the segment when the label overflows a little, else outside with a leader, else hidden.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import {
  checkScene, layoutChart,
  type ChartSpec, type LineShape, type RectShape, type Scene, type TextShape,
} from '../src/core/index.ts';
import { placeLabels, segmentCandidates, type LabelRequest } from '../src/core/labels.ts';
import { overlaps, segmentHitsBox, type Box } from '../src/core/place.ts';
import { createMeasurer } from '../src/fonts.ts';
import { CHART_FRAME } from '../src/deck.ts';
import { allCases, allFrames } from './resize-matrix.ts';

const m = createMeasurer();
const F = CHART_FRAME;
const byRole = (s: Scene, re: RegExp) => s.shapes.filter((_, i) => re.test(s.roles![i]));
const one = <T>(s: Scene, role: string) => s.shapes[s.roles!.indexOf(role)] as T;
const textsOf = (s: Scene, re: RegExp) => byRole(s, re).filter((x): x is TextShape => x.kind === 'text');
const strict = (s: Scene) => checkScene(s, { strict: true, measurer: m });
const deck = (f: string, i: number): ChartSpec => JSON.parse(readFileSync(`examples/${f}.json`, 'utf8')).slides[i].chart;
const bounds: Box = { x: 0, y: 0, w: 200, h: 200 };

/** Every leader starts on its segment's edge, ends at its label and crosses no other text. */
function assertLeaders(scene: Scene) {
  const texts = textsOf(scene, /./);
  scene.roles!.forEach((r, i) => {
    const mm = r.match(/^leader:(s\d+:c\d+)$/);
    if (!mm) return;
    const l = scene.shapes[i] as LineShape;
    const seg = one<RectShape>(scene, `bar:${mm[1]}`);
    const label = textsOf(scene, new RegExp(`^label:(value|total):(${mm[1]}|${mm[1].split(':')[1]})$`)).find((t) => Math.abs(t.y + t.h / 2 - l.y2) < 1 || Math.abs(t.x + t.w / 2 - l.x2) < 1)!;
    assert.ok(label, `${r} has its label`);
    const onEdge = (Math.abs(l.x1 - seg.x) < 0.01 || Math.abs(l.x1 - (seg.x + seg.w)) < 0.01) ? l.y1 >= seg.y - 0.01 && l.y1 <= seg.y + seg.h + 0.01
      : (Math.abs(l.y1 - seg.y) < 0.01 || Math.abs(l.y1 - (seg.y + seg.h)) < 0.01) && l.x1 >= seg.x - 0.01 && l.x1 <= seg.x + seg.w + 0.01;
    assert.ok(onEdge, `${r} starts on the edge of its segment`);
    for (const t of texts) if (t !== label) assert.ok(!segmentHitsBox(l.x1, l.y1, l.x2, l.y2, { x: t.x + 0.5, y: t.y + 0.5, w: t.w - 1, h: t.h - 1 }), `${r} crosses "${t.text}"`);
  });
}

// ---- Placer --------------------------------------------------------------------------------------

test('placer: cheapest valid candidate wins; obstacles and bounds rule candidates out', () => {
  const anchor = { x: 90, y: 100, w: 20, h: 2 };
  const cands = segmentCandidates(anchor, 16, 10, { sides: ['right', 'left'], lead: 3 });
  assert.equal(cands[0].side, 'right');
  assert.deepEqual(placeLabels([{ anchor, candidates: cands }], [anchor], bounds), [0]);
  // Something right of the segment: the label goes left, level with it.
  const wall = { x: 112, y: 80, w: 10, h: 40 };
  const [k] = placeLabels([{ anchor, candidates: cands }], [anchor, wall], bounds);
  assert.equal(cands[k!].side, 'left');
  assert.equal(cands[k!].box.y + 5, 101);
  // Hemmed in on both sides and by the bounds: no position, reported as null.
  const tight = { x: 0, y: 95, w: 200, h: 12 };
  const req: LabelRequest = { anchor, candidates: segmentCandidates(anchor, 16, 10, { sides: ['right', 'left'], lead: 3, stagger: 0 }) };
  assert.deepEqual(placeLabels([req], [anchor, wall, { x: 70, y: 80, w: 17, h: 40 }], tight), [null]);
});

test('placer: two labels too close together go one to the right, the other to the left', () => {
  const a = { x: 90, y: 100, w: 20, h: 2 }, b = { x: 90, y: 102, w: 20, h: 2 };
  const reqs = [a, b].map((anchor) => ({ anchor, candidates: segmentCandidates(anchor, 16, 10, { sides: ['right', 'left'], lead: 3 }) }));
  const choice = placeLabels(reqs, [a, b], bounds);
  const sides = choice.map((k, i) => reqs[i].candidates[k!].side).sort();
  assert.deepEqual(sides, ['left', 'right']);
  const [p, q] = choice.map((k, i) => reqs[i].candidates[k!].box);
  assert.ok(!overlaps(p, q, -1));
});

test('placer: a label overflowing its segment a little is boxed on it; a much larger one goes outside', () => {
  const seg = { x: 50, y: 100, w: 40, h: 8 }, thin = { x: 50, y: 100, w: 40, h: 3 };
  const boxed = segmentCandidates(seg, 16, 12, { sides: ['right'], lead: 3, boxCover: [] });
  assert.equal(boxed[placeLabels([{ anchor: seg, candidates: boxed }], [seg], bounds)[0]!].side, 'boxed');
  const out = segmentCandidates(thin, 16, 12, { sides: ['right'], lead: 3, boxCover: [] });
  assert.ok(!out.some((c) => c.side === 'boxed'), 'over twice the segment: never boxed');
  // A boxed label may cover its own column but nothing else.
  const neighbour = { x: 50, y: 108, w: 40, h: 30 };
  assert.equal(boxed[placeLabels([{ anchor: seg, candidates: boxed }], [seg, neighbour], bounds)[0]!].side, 'right');
  const own = segmentCandidates(seg, 16, 12, { sides: ['right'], lead: 3, boxCover: [neighbour] });
  assert.equal(own[placeLabels([{ anchor: seg, candidates: own }], [seg, neighbour], bounds)[0]!].side, 'boxed');
});

test('placer: a boxed label never pokes out of its column: its text must fit the column width', () => {
  const seg = { x: 50, y: 100, w: 20, h: 40 };
  // Text 21 wide (+2 air) on a 20-wide column: outside, never boxed.
  assert.ok(!segmentCandidates(seg, 23, 12, { sides: ['right'], lead: 3, boxCover: [] }).some((c) => c.side === 'boxed'));
  // Text 19 wide (+2 air): boxed, the box trimmed to the column.
  const fit = segmentCandidates(seg, 21, 12, { sides: ['right'], lead: 3, boxCover: [] }).find((c) => c.side === 'boxed')!;
  assert.deepEqual([fit.box.x, fit.box.w], [seg.x, seg.w]);
  // Every chart at every resize frame: a filled (boxed) value label stays within its column's width.
  for (const c of allCases()) {
    for (const f of allFrames()) {
      const scene = layoutChart(c.spec, f, m);
      scene.shapes.forEach((t, i) => {
        const mm = scene.roles![i].match(/^label:value:(s\d+:c\d+)$/);
        if (t.kind !== 'text' || !t.fill || !mm) return;
        const bar = one<RectShape | undefined>(scene, `bar:${mm[1]}`);
        if (!bar) return;
        assert.ok(t.x >= bar.x - 0.01 && t.x + t.w <= bar.x + bar.w + 0.01, `${c.name} ${f.w}x${f.h}: "${t.text}" box wider than its column`);
      });
    }
  }
});

test('placer: leaders never cross each other or another label; a blocked label evicts one blocker', () => {
  // Three stacked thin segments, only the right side allowed: labels stagger in segment order, so leaders do not cross.
  const segs = [0, 1, 2].map((i) => ({ x: 90, y: 100 + i * 2, w: 20, h: 2 }));
  const reqs = segs.map((anchor) => ({ anchor, candidates: segmentCandidates(anchor, 16, 10, { sides: ['right'], lead: 3 }) }));
  const choice = placeLabels(reqs, segs, bounds);
  assert.ok(choice.every((k) => k != null));
  const picked = choice.map((k, i) => reqs[i].candidates[k!]);
  for (let i = 0; i < 3; i++) {
    for (let j = 0; j < 3; j++) {
      if (i === j) continue;
      const l = picked[i].leader!, b = picked[j].box;
      assert.ok(!segmentHitsBox(l.x1, l.y1, l.x2, l.y2, { x: b.x + 0.5, y: b.y + 0.5, w: b.w - 1, h: b.h - 1 }));
    }
  }
  assert.ok(picked[0].box.y < picked[1].box.y && picked[1].box.y < picked[2].box.y, 'top segment, top label');
  // Eviction: A can sit right or left; B only right. Greedy places A right first (cheaper), then B moves A.
  const a = { x: 90, y: 100, w: 20, h: 2 };
  const ra = { anchor: a, candidates: segmentCandidates(a, 16, 10, { sides: ['right', 'left'], lead: 3, stagger: 0 }) };
  const rb = { anchor: a, candidates: segmentCandidates(a, 16, 10, { sides: ['right'], lead: 3, stagger: 0 }) };
  const [ka, kb] = placeLabels([ra, rb], [a], bounds);
  assert.equal(ra.candidates[ka!].side, 'left');
  assert.equal(rb.candidates[kb!].side, 'right');
});

test('quality gate: a boxed segment label may cover its own column only', () => {
  const t = (x: number, fill?: string): TextShape => ({ kind: 'text', x, y: 10, w: 20, h: 14, text: '5', size: 12, color: '000000', align: 'center', valign: 'middle', ...(fill ? { fill } : {}) });
  const bar = (x: number): RectShape => ({ kind: 'rect', x, y: 0, w: 30, h: 40, fill: '1F3A5F' });
  const scene: Scene = { frame: { x: 0, y: 0, w: 200, h: 100 }, hidden: [], shapes: [bar(0), bar(100), t(5, '7FA3CC'), t(95, '7FA3CC')], roles: ['bar:s1:c0', 'bar:s1:c1', 'label:value:s0:c0', 'label:value:s0:c0'] };
  assert.deepEqual(checkScene(scene).map((i) => i.detail), ['"5" × bar:s1:c1']);
});

// ---- Charts --------------------------------------------------------------------------------------

test('stacked: the sample deck labels every thin Services segment, boxed or outside with a leader', () => {
  const scene = layoutChart(deck('deck', 0), F, m);
  assert.deepEqual(scene.hidden, []);
  assert.deepEqual(strict(scene), []);
  for (const c of [0, 1, 2]) {
    const label = one<TextShape>(scene, `label:value:s3:c${c}`);
    assert.ok(label, `Services / column ${c} is labelled`);
    const leader = scene.roles!.includes(`leader:s3:c${c}`);
    assert.ok(leader || label.fill, 'outside labels have a leader, the others are boxed');
  }
  assertLeaders(scene);
});

test('stacked: every thin segment of the crowded quarterly chart is placed, leaders cross nothing', () => {
  const scene = layoutChart(deck('deck', 3), F, m);
  assert.deepEqual(scene.hidden, []);
  assert.deepEqual(strict(scene), []);
  assertLeaders(scene);
  // No label beside the last column lands among the series names on the right.
  const names = textsOf(scene, /^label:series:/);
  const nameX = Math.min(...names.map((t) => t.x));
  for (const t of textsOf(scene, /^label:value:s\d+:c7$/)) assert.ok(t.x + t.w <= nameX, `"${t.text}" stays left of the series names`);
});

test('stacked: labels too thin for any spot are hidden and reported, never overlapping', () => {
  const big = [90, 95, 100, 105, 110, 115];
  const cats = ['a', 'b', 'c', 'd', 'e', 'f'];
  const scene = layoutChart({ type: 'stacked', categories: cats, series: [{ name: 'Big', values: big }, { name: 'Thin', values: big.map(() => 1) }, { name: 'Top', values: big }], options: { gap: 0.05 } }, F, m);
  assert.deepEqual(scene.hidden, cats.map((c) => `Thin / ${c}: segment too small`));
  assert.ok(!scene.roles!.some((r) => r.startsWith('leader:')));
  assert.deepEqual(checkScene(scene, { measurer: m }), []);
});

test('100% chart: thin shares go outside the column with a leader', () => {
  const scene = layoutChart({ type: 'percent', categories: ['2023', '2024', '2025'], series: [{ name: 'A', values: [60, 62, 64] }, { name: 'B', values: [39, 37, 35.5] }, { name: 'C', values: [1, 1, 0.5] }] }, F, m);
  assert.deepEqual(scene.hidden, []);
  assert.deepEqual(strict(scene), []);
  for (const c of [0, 1]) assert.ok(scene.roles!.includes(`leader:s2:c${c}`), `C / column ${c} has a leader`);
  assertLeaders(scene);
});

test('horizontal bars: thin segments are labelled above or below the bar with a vertical leader', () => {
  const scene = layoutChart({
    type: 'stacked', categories: ['North', 'South', 'East'],
    series: [{ name: 'A', values: [50, 60, 40] }, { name: 'B', values: [30, 20, 35] }, { name: 'C', values: [0.6, 0.8, 0.4] }, { name: 'D', values: [20, 15, 25] }],
    options: { orientation: 'horizontal', decimals: 1 },
  }, F, m);
  assert.deepEqual(scene.hidden, []);
  assert.deepEqual(strict(scene), []);
  for (const c of [0, 1, 2]) {
    const seg = one<RectShape>(scene, `bar:s2:c${c}`), label = one<TextShape>(scene, `label:value:s2:c${c}`), l = one<LineShape>(scene, `leader:s2:c${c}`);
    assert.ok(label.y + label.h <= seg.y || label.y >= seg.y + seg.h, 'outside the bar');
    assert.ok(Math.abs(l.x1 - (seg.x + seg.w / 2)) < 0.01, 'leader from the segment centre');
  }
  assertLeaders(scene);
});

test('Mekko: thin segments of the last column go beside it; the series names move right to make room', () => {
  const scene = layoutChart(deck('mekko', 1), F, m);
  assert.deepEqual(scene.hidden, []);
  assert.deepEqual(strict(scene), []);
  assertLeaders(scene);
  const names = textsOf(scene, /^label:series:/);
  const outside = textsOf(scene, /^label:value:s\d+:c3$/).filter((t) => scene.roles!.includes(`leader:${scene.roles![scene.shapes.indexOf(t)].slice(12)}`));
  assert.ok(outside.length >= 2);
  for (const t of outside) assert.ok(t.x + t.w <= Math.min(...names.map((n) => n.x)), `"${t.text}" left of the series names`);
});

test('Mekko: the narrow Rest-of-world column gets all its shares beside it', () => {
  const scene = layoutChart(deck('mekko', 2), F, m);
  assert.deepEqual(scene.hidden, []);
  assert.deepEqual(strict(scene), []);
  assertLeaders(scene);
  const seg = one<RectShape>(scene, 'bar:s0:c4');
  for (const s of [0, 1, 2, 3]) assert.ok(one<TextShape>(scene, `label:value:s${s}:c4`).x >= seg.x + seg.w);
});

test('waterfall: a thin segment inside a stack goes beside the column with a leader, the outer ones above and below', () => {
  const spec: ChartSpec = {
    type: 'waterfall', categories: ['Start', 'A', 'End'],
    series: [{ name: 'a', values: [0.2, 0.1, null] }, { name: 'b', values: [0, 0.1, null] }, { name: 'c', values: [0, 0.1, null] }, { name: 'd', values: [0, 5, null] }],
    options: { seriesLabels: 'none' },
  };
  const scene = layoutChart(spec, { x: 20, y: 20, w: 400, h: 150 }, m);
  assert.deepEqual(scene.hidden, []);
  assert.deepEqual(strict(scene), []);
  const col = one<RectShape>(scene, 'bar:s3:c1');
  for (const s of [0, 1, 2]) {
    const t = one<TextShape>(scene, `label:value:s${s}:c1`);
    assert.ok(t.x + t.w <= col.x || t.x >= col.x + col.w, `+0.1 of series ${s} sits beside the column`);
    assert.ok(scene.roles!.includes(`leader:s${s}:c1`));
  }
  assertLeaders(scene);
  // Two negative deltas in one column: the upper one is labelled above, the lower one below.
  const wf = layoutChart(deck('waterfall', 2), F, m);
  const top = one<RectShape>(wf, 'bar:s0:c6'), bottom = one<RectShape>(wf, 'bar:s1:c6');
  assert.ok(one<TextShape>(wf, 'label:value:s0:c6').y + 13 <= top.y, 'organic −3.1m above');
  assert.ok(one<TextShape>(wf, 'label:value:s1:c6').y >= bottom.y + bottom.h, 'acquired −1.2m below');
});
