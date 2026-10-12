import { test } from 'node:test';
import assert from 'node:assert/strict';
import { flattenPath, pathBounds } from '../src/core/path.ts';
import { writePptx } from '../src/render/pptx.ts';
import { sceneToSvg } from '../src/render/svg.ts';
import { unzip } from '../src/ooxml/unzip.ts';
import type { PathShape, Scene } from '../src/core/types.ts';

const wedge: PathShape = {
  kind: 'path', fill: '1F3A5F', stroke: 'FFFFFF', strokeWidth: 1,
  commands: [{ op: 'M', x: 100, y: 100 }, { op: 'A', cx: 100, cy: 100, r: 50, start: -90, sweep: 90 }, { op: 'Z' }],
};

test('arcs flatten to points on the circle, starting and ending at the requested angles', () => {
  const [sub] = flattenPath(wedge.commands);
  assert.equal(sub.closed, true);
  assert.deepEqual(sub.points[0], { x: 100, y: 100 });
  const first = sub.points[1], last = sub.points.at(-1)!;
  assert.ok(Math.abs(first.x - 100) < 1e-9 && Math.abs(first.y - 50) < 1e-9, '12 o\'clock');
  assert.ok(Math.abs(last.x - 150) < 1e-9 && Math.abs(last.y - 100) < 1e-9, '3 o\'clock');
  for (const p of sub.points.slice(1)) assert.ok(Math.abs(Math.hypot(p.x - 100, p.y - 100) - 50) < 1e-9);
  const b = pathBounds(wedge);
  assert.ok(Math.abs(b.x - 100) < 1e-9 && Math.abs(b.y - 50) < 1e-9 && Math.abs(b.w - 50) < 1e-9 && Math.abs(b.h - 50) < 1e-9);
});

test('path and preset shapes render to PowerPoint custom geometry and SVG', () => {
  const scene: Scene = {
    frame: { x: 0, y: 0, w: 400, h: 300 }, hidden: [],
    shapes: [wedge, { kind: 'rect', x: 10, y: 10, w: 20, h: 20, fill: 'C0504D', geom: 'diamond' }],
    roles: ['wedge:s0:c0', 'marker:s1:c0'],
  };
  const files = unzip(writePptx([{ scenes: [{ scene, name: 'Chart', altText: 'test' }] }], 'Arial'));
  const slide = files.get('ppt/slides/slide1.xml')!.text();
  assert.match(slide, /<a:custGeom>.*<a:moveTo>.*<a:lnTo>.*<a:close\/>/);
  assert.match(slide, /<a:prstGeom prst="diamond">/);
  const svg = sceneToSvg([scene], { w: 400, h: 300 });
  assert.match(svg, /<path d="M100,100 L100,50 .* Z" fill="#1F3A5F"/);
  assert.match(svg, /<polygon points="20,10 30,20 20,30 10,20"/);
});
