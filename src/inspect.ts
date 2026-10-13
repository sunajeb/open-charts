// Human-readable summary of the open-charts objects in a .pptx (`open-charts inspect`).
import type { Frame } from './core/types.ts';
import type { SlideInfo } from './ooxml/pptx-read.ts';

const f = (fr?: Frame) => (fr ? [fr.x, fr.y, fr.w, fr.h].map((v) => +v.toFixed(1)).join(',') : '?');
const bytes = (n: number) => (n < 1024 ? `${n} B` : `${(n / 1024).toFixed(1)} KB`);
const moved = (a?: Frame, b?: Frame) => !!a && !!b && [a.x - b.x, a.y - b.y, a.w - b.w, a.h - b.h].some((d) => Math.abs(d) > 0.5);

export function formatInspect(file: string, slides: SlideInfo[]): string {
  const charts = slides.flatMap((s) => s.charts);
  const contracts = [...new Set(charts.map((c) => c.contract))].join(', ') || '-';
  const lines = [`${file}: ${slides.length} slide${slides.length === 1 ? '' : 's'}, ${charts.length} chart${charts.length === 1 ? '' : 's'} (tag contract ${contracts})`];
  for (const c of charts) {
    const kinds = new Map<string, number>();
    for (const r of c.roles) { const k = r.split(':')[0]; kinds.set(k, (kinds.get(k) ?? 0) + 1); }
    const roleSummary = [...kinds].map(([k, n]) => `${k} ${n}`).join(', ') || 'none';
    const type = c.spec ? safeType(c.spec) : '?';
    lines.push(
      `  slide ${c.slide}  "${c.name}"  ${type}  frame ${f(c.frame)}${moved(c.frame, c.bounds) ? `  RESIZED to ${f(c.bounds)}` : ''}`,
      `    engine ${c.engine ?? '?'}  font ${c.font ?? '?'}  spec ${c.spec === undefined ? c.error ?? 'missing' : bytes(Buffer.byteLength(c.spec))}${c.specChunks > 1 ? ` in ${c.specChunks} chunks` : ''}  tags ${bytes(c.tagBytes)}`,
      `    ${c.children} shapes, ${c.roles.length} roles: ${roleSummary}`,
    );
  }
  return lines.join('\n');
}

function safeType(spec: string): string {
  try { return String(JSON.parse(spec).type ?? '?'); } catch { return 'invalid OC_SPEC'; }
}
