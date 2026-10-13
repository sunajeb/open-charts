// SVG preview of a slide. Uses the same scene as the PowerPoint writer, so the two always match.
import type { Scene, Shape } from '../core/types.ts';
import { flattenPath, rectGeom } from '../core/path.ts';
import { colorHex } from '../core/theme.ts';

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
// A theme colour is drawn as the hex it was resolved to ("1F3A5F@accent1" → #1F3A5F).
const c = (v?: string) => (v ? `#${colorHex(v)}` : 'none');
const n = (v: number) => +v.toFixed(2);

function shape(s: Shape, font: string): string {
  if (s.kind === 'rect') {
    const stroke = s.stroke ? ` stroke="${c(s.stroke)}" stroke-width="${s.strokeWidth ?? 0.75}"` : '';
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
    const stroke = s.stroke ? ` stroke="${c(s.stroke)}" stroke-width="${s.strokeWidth ?? 0.75}"${s.dash ? ' stroke-dasharray="4 3"' : ''} stroke-linejoin="round"` : '';
    return `<path d="${d}" fill="${c(s.fill)}"${stroke}/>`;
  }
  if (s.kind === 'line') {
    return `<line x1="${n(s.x1)}" y1="${n(s.y1)}" x2="${n(s.x2)}" y2="${n(s.y2)}" stroke="${c(s.color)}" stroke-width="${s.width}"${s.dash ? ' stroke-dasharray="4 3"' : ''}${s.arrowEnd ? ' marker-end="url(#arrow)"' : ''}/>`;
  }
  const lines = s.text.split('\n');
  const lh = s.lineHeight ?? s.size * 1.149;
  const block = lines.length * lh;
  const top = s.valign === 'top' ? s.y : s.valign === 'bottom' ? s.y + s.h - block : s.y + (s.h - block) / 2;
  const x = s.align === 'left' ? s.x : s.align === 'right' ? s.x + s.w : s.x + s.w / 2;
  const anchor = { left: 'start', center: 'middle', right: 'end' }[s.align];
  const box = s.fill || s.stroke
    ? `<rect x="${n(s.x)}" y="${n(s.y)}" width="${n(s.w)}" height="${n(s.h)}" rx="${s.rounded ? n(s.h / 2) : 0}" fill="${c(s.fill)}" stroke="${c(s.stroke)}" stroke-width="0.75"/>`
    : '';
  const b = s.bullet;
  const tx = b && s.align === 'left' ? x + b.at + b.hang : x;
  const spans = lines.map((l, i) => `<tspan x="${n(tx)}" y="${n(top + lh * i + lh / 2)}">${esc(l)}</tspan>`).join('');
  const mark = b && s.align === 'left'
    ? `<text font-family="${esc(font)}, Helvetica, sans-serif" font-size="${s.size}"${s.bold ? ' font-weight="bold"' : ''} fill="${c(s.color)}" x="${n(x + b.at)}" y="${n(top + lh / 2)}" dominant-baseline="central">${esc(b.number !== undefined ? `${b.number}.` : b.char ?? '\u2022')}</text>`
    : '';
  return `${box}<text font-family="${esc(font)}, Helvetica, sans-serif" font-size="${s.size}"${s.bold ? ' font-weight="bold"' : ''} fill="${c(s.color)}" text-anchor="${anchor}" dominant-baseline="central">${spans}</text>${mark}`;
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
