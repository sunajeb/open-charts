// SVG preview of a slide. Uses the same scene as the PowerPoint writer, so the two always match.
import type { Scene, Shape } from '../core/types.ts';
import { flattenPath, rectGeom } from '../core/path.ts';
import { colorHex } from '../core/theme.ts';

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
// A theme colour is drawn as the hex it was resolved to ("1F3A5F@accent1" → #1F3A5F).
const c = (v?: string) => (v ? `#${colorHex(v)}` : 'none');
const n = (v: number) => +v.toFixed(2);
const dashAttr = (d?: boolean | 'dot') => (d ? ` stroke-dasharray="${d === 'dot' ? '1 2' : '4 3'}"` : '');
/** An element's font: its own (an override; the theme fonts "+mj-lt" / "+mn-lt" preview as the chart's), else the chart's. */
const faceOf = (f: string | undefined, chart: string) => (f && !f.startsWith('+') ? f : chart);

function shape(s: Shape, font: string): string {
  if (s.kind === 'rect') {
    const stroke = s.stroke ? ` stroke="${c(s.stroke)}" stroke-width="${s.strokeWidth ?? 0.75}"${dashAttr(s.dash)}` : '';
    const g = rectGeom(s);
    if (g === 'diamond' || g === 'triangle') {
      const pts = g === 'diamond'
        ? [[s.x + s.w / 2, s.y], [s.x + s.w, s.y + s.h / 2], [s.x + s.w / 2, s.y + s.h], [s.x, s.y + s.h / 2]]
        : [[s.x + s.w / 2, s.y], [s.x + s.w, s.y + s.h], [s.x, s.y + s.h]];
      return `<polygon points="${pts.map(([x, y]) => `${n(x)},${n(y)}`).join(' ')}" fill="${c(s.fill)}"${stroke}/>`;
    }
    if (g === 'roundRect') return `<rect x="${n(s.x)}" y="${n(s.y)}" width="${n(s.w)}" height="${n(s.h)}" rx="${n(Math.min(s.w, s.h) / 6)}" fill="${c(s.fill)}"${stroke}/>`;
    if (g === 'ellipse') return `<ellipse cx="${n(s.x + s.w / 2)}" cy="${n(s.y + s.h / 2)}" rx="${n(s.w / 2)}" ry="${n(s.h / 2)}" fill="${c(s.fill)}"${stroke}/>`;
    return `<rect x="${n(s.x)}" y="${n(s.y)}" width="${n(s.w)}" height="${n(s.h)}" fill="${c(s.fill)}"${stroke}/>`;
  }
  if (s.kind === 'path') {
    const d = flattenPath(s.commands).map(({ points, closed }) =>
      `M${points.map((p) => `${n(p.x)},${n(p.y)}`).join(' L')}${closed ? ' Z' : ''}`).join(' ');
    const stroke = s.stroke ? ` stroke="${c(s.stroke)}" stroke-width="${s.strokeWidth ?? 0.75}"${dashAttr(s.dash)} stroke-linejoin="round"` : '';
    return `<path d="${d}" fill="${c(s.fill)}"${stroke}/>`;
  }
  if (s.kind === 'line') {
    return `<line x1="${n(s.x1)}" y1="${n(s.y1)}" x2="${n(s.x2)}" y2="${n(s.y2)}" stroke="${c(s.color)}" stroke-width="${s.width}"${dashAttr(s.dash)}${s.arrowStart ? ' marker-start="url(#arrow)"' : ''}${s.arrowEnd ? ' marker-end="url(#arrow)"' : ''}/>`;
  }
  const lines = s.text.split('\n');
  const lh = s.lineHeight ?? s.size * 1.149;
  const block = lines.length * lh;
  // Turned text (an override) is laid out in its unturned box, centred on the footprint, then rotated.
  const b0 = s.rotate ? { x: s.x + s.w / 2 - s.h / 2, y: s.y + s.h / 2 - s.w / 2, w: s.h, h: s.w } : s;
  const top = s.valign === 'top' ? b0.y : s.valign === 'bottom' ? b0.y + b0.h - block : b0.y + (b0.h - block) / 2;
  const x = s.align === 'left' ? b0.x : s.align === 'right' ? b0.x + b0.w : b0.x + b0.w / 2;
  const anchor = { left: 'start', center: 'middle', right: 'end' }[s.align];
  const box = s.fill || s.stroke
    ? `<rect x="${n(s.x)}" y="${n(s.y)}" width="${n(s.w)}" height="${n(s.h)}" rx="${s.rounded ? n(s.h / 2) : 0}" fill="${c(s.fill)}" stroke="${c(s.stroke)}" stroke-width="${s.strokeWidth ?? 0.75}"/>`
    : '';
  const spans = lines.map((l, i) => `<tspan x="${n(x)}" y="${n(top + lh * i + lh / 2)}">${esc(l)}</tspan>`).join('');
  const rot = s.rotate ? ` transform="rotate(${s.rotate} ${n(s.x + s.w / 2)} ${n(s.y + s.h / 2)})"` : '';
  return `${box}<text font-family="${esc(faceOf(s.font, font))}, Helvetica, sans-serif" font-size="${s.size}"${s.bold ? ' font-weight="bold"' : ''} fill="${c(s.color)}" text-anchor="${anchor}" dominant-baseline="central"${rot}>${spans}</text>`;
}

export function sceneToSvg(scenes: Scene[], size: { w: number; h: number }, title?: string, font = 'Arial'): string {
  const head = title
    ? `<text x="36" y="46" font-family="${esc(font)}, Helvetica, sans-serif" font-size="24" font-weight="bold" fill="#1A1A1A">${esc(title)}</text>`
    : '';
  const body = scenes.flatMap((sc) => sc.shapes.map((s) => shape(s, font))).join('\n');
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${size.w} ${size.h}" width="${size.w}" height="${size.h}">
<defs><marker id="arrow" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="#7F7F7F"/></marker></defs>
<rect width="100%" height="100%" fill="#FFFFFF"/>
${head}
${body}
</svg>`;
}
