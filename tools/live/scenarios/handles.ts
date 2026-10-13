// The drag handles: a stacked chart with a value axis, selected; what the server and the overlay
// show (GET /api/addin/handles; no screen capture):
//   1. mapping: at three zoom levels and after moving the window, the handles' view (slide -> screen)
//      has the zoom as its factor and moves with the window
//   2. gap: a real mouse drag (CGEvent: the terminal needs Accessibility) on the gap handle -> the
//      chart's OC_SPEC has the new options.gap and the chart was re-laid out (one APPLY)
//   3. label: two clicks select a value label, a drag pins it -> overrides[role].pin
// The mouse moves: nobody should use the Mac meanwhile.
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { tokenHeaders } from '../../../src/addin/token.ts';
import { sleep } from '../powerpoint.ts';
import type { Ctx } from '../scenarios.ts';

// Mouse events (CGEvent, global top-left points): a tiny Swift tool, compiled once per source.
const MOUSE_SRC = `import CoreGraphics
import Foundation
let a = CommandLine.arguments.dropFirst().map { Double($0)! }
func post(_ t: CGEventType, _ x: Double, _ y: Double, _ clicks: Int64 = 1) {
  let e = CGEvent(mouseEventSource: nil, mouseType: t, mouseCursorPosition: CGPoint(x: x, y: y), mouseButton: .left)!
  e.setIntegerValueField(.mouseEventClickState, value: clicks)
  e.post(tap: .cghidEventTap)
}
// mode 0: drag (x0 y0 x1 y1 steps); 1: click (x y); 2: double-click (x y)
let mode = Int(a[0])
if mode == 0 {
  let (x0, y0, x1, y1, n) = (a[1], a[2], a[3], a[4], Int(a[5]))
  post(.mouseMoved, x0 - 3, y0); usleep(120_000)
  post(.mouseMoved, x0, y0); usleep(250_000)
  post(.leftMouseDown, x0, y0); usleep(60_000)
  for i in 1...n { let f = Double(i) / Double(n); post(.leftMouseDragged, x0 + (x1 - x0) * f, y0 + (y1 - y0) * f); usleep(16_000) }
  usleep(150_000)
  post(.leftMouseUp, x1, y1)
} else {
  let (x, y) = (a[1], a[2])
  post(.mouseMoved, x, y); usleep(250_000)
  post(.leftMouseDown, x, y); post(.leftMouseUp, x, y)
  if mode == 2 { usleep(80_000); post(.leftMouseDown, x, y, 2); post(.leftMouseUp, x, y, 2) }
}
`;

type State = {
  id?: string; view?: { k: number; ox: number; oy: number }; options?: Record<string, unknown>; overrides?: Record<string, { pin?: unknown }>; roles?: string[];
  handles?: Array<{ id: string; rect: number[] }>; labels?: Array<{ role: string; rect: number[] }>; overlay?: Record<string, unknown>;
};

export async function run(ctx: Ctx) {
  const { pp, check } = ctx;
  const dir = join(ctx.root, 'out', 'live', '.mouse');
  const exe = join(dir, `mouse-${createHash('sha1').update(MOUSE_SRC).digest('hex').slice(0, 8)}`);
  if (!existsSync(exe)) {
    mkdirSync(dir, { recursive: true });
    writeFileSync(`${exe}.swift`, MOUSE_SRC);
    execFileSync('xcrun', ['swiftc', '-O', `${exe}.swift`, '-o', exe], { stdio: 'ignore', timeout: 90_000 });
  }
  const mouse = (...a: number[]) => execFileSync(exe, a.map(String), { timeout: 10_000 });

  const state = async (): Promise<State> => (await fetch('http://127.0.0.1:4174/api/addin/handles', { headers: tokenHeaders(), signal: AbortSignal.timeout(3000) })).json() as Promise<State>;
  const until = async (ok: (s: State) => boolean, ms = 8000) => {
    let s: State = {};
    for (const end = Math.min(Date.now() + ms, pp.deadline); Date.now() < end; await new Promise((r) => setTimeout(r, 150))) {
      try { s = await state(); if (ok(s)) return s; } catch { /* server starting */ }
    }
    return s;
  };
  const reselect = () => { pp.macro('SelectChart', '0'); sleep(300); pp.macro('SelectChart', '1'); };
  const centre = (r: number[]) => [r[0] + r[2] / 2, r[1] + r[3] / 2];

  pp.newPresentation();
  pp.newSlide();
  pp.run('chart_stacked');
  sleep(1500);
  reselect();
  let s = await until((x) => !!x.handles?.length);
  check(!!s.handles?.length, `handles shown for the selected chart: ${JSON.stringify(s.handles?.map((h) => h.id))}`);
  check(s.overlay?.visible === true && s.overlay?.ignoresMouse === true, `overlay visible and click-through away from handles (${JSON.stringify(s.overlay)})`);

  // 1. Mapping at three zoom levels and with the window moved.
  for (const zoom of [50, 75, 120]) {
    pp.tell(`set zoom of view of document window 1 to ${zoom}`);
    sleep(400);
    reselect();
    s = await until((x) => !!x.view && Math.abs(x.view.k - zoom / 100) < 0.05);
    check(!!s.view && Math.abs(s.view.k - zoom / 100) < 0.02, `zoom ${zoom}%: view factor ${s.view?.k.toFixed(4)} (offset ${s.view?.ox.toFixed(1)}, ${s.view?.oy.toFixed(1)})`);
  }
  pp.tell('set zoom of view of document window 1 to 75');
  reselect();
  const before = await until((x) => !!x.view && Math.abs(x.view.k - 0.75) < 0.05);
  const b0 = pp.tell('get bounds of document window 1').split(',').map((n) => Number(n.trim()));
  pp.tell(`set bounds of document window 1 to {${b0[0] + 40}, ${b0[1] + 30}, ${b0[2] + 40}, ${b0[3] + 30}}`);
  sleep(600);
  reselect();
  const after = await until((x) => !!x.view && !!before.view && Math.abs(x.view.ox - before.view.ox) > 5);
  check(!!after.view && !!before.view && Math.abs(after.view.ox - before.view.ox - 40) < 3 && Math.abs(after.view.oy - before.view.oy - 30) < 3,
    `window moved by (40, 30): view moved by (${((after.view?.ox ?? 0) - (before.view?.ox ?? 0)).toFixed(1)}, ${((after.view?.oy ?? 0) - (before.view?.oy ?? 0)).toFixed(1)})`);
  try { pp.tell(`set bounds of document window 1 to {${b0.join(', ')}}`); } catch (err) { ctx.note(`window not moved back: ${(err as Error).message.trim()}`); }
  sleep(600);
  reselect();
  s = await until((x) => !!x.handles?.some((h) => h.id === 'gap'));

  // 2. Gap drag.
  const gap = s.handles?.find((h) => h.id === 'gap');
  if (!check(!!gap, 'the stacked chart has a gap handle') || !gap) return;
  const gap0 = s.options?.gap;
  const [gx, gy] = centre(gap.rect);
  let from = pp.logLines().length;
  mouse(0, gx, gy, gx - 12, gy, 20);
  const hline = pp.waitForLog(/HANDLES gap ->/, from, 15_000);
  const apply = pp.waitForLog(/APPLY \w+ (done|ran|failed)/, from, 20_000);
  ctx.note(`${hline.slice(20)} / ${apply.slice(20)}`);
  reselect();
  s = await until((x) => x.options?.gap !== gap0);
  check(typeof s.options?.gap === 'number' && (s.options.gap as number) > 0.5, `gap dragged narrower: options.gap ${String(gap0)} -> ${String(s.options?.gap)} (read back from the chart's OC_SPEC)`);
  check(/ done /.test(apply), 'the chart was re-laid out in PowerPoint (one APPLY)');

  // 3. Label drag: two clicks select a value label, then drag it.
  const label = s.labels?.find((l) => l.role.startsWith('label:value:s0:'));
  if (!check(!!label, 'a value label is on the chart') || !label) return;
  const [lx, ly] = centre(label.rect);
  mouse(1, lx, ly); sleep(700); mouse(1, lx, ly);
  s = await until((x) => !!x.handles?.some((h) => h.id === `label:${label.role}`));
  const lh = s.handles?.find((h) => h.id === `label:${label.role}`);
  if (!check(!!lh, `selected ${label.role}: its box is a handle (roles ${JSON.stringify(s.roles)})`)) return;
  from = pp.logLines().length;
  mouse(0, lx, ly, lx + 30, ly - 25, 20);
  const l = pp.waitForLog(/HANDLES label:/, from, 15_000);
  const a2 = pp.waitForLog(/APPLY \w+ (done|ran|failed)/, from, 20_000);
  ctx.note(`${l.slice(20)} / ${a2.slice(20)}`);
  reselect();
  s = await until((x) => !!x.overrides?.[label.role]?.pin);
  check(!!s.overrides?.[label.role]?.pin, `label pinned: ${JSON.stringify(s.overrides?.[label.role])}`);
}
