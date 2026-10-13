// Style looks that change appearance only: label colours by role and line styles (colour, width,
// dash) by role. Applied after layout — nothing here may move a shape, so switching a style file
// keeps every frame, role and label decision and only recolours / restrokes (style files, §P2).
import type { LineLook, Scene, Shape, Style } from './types.ts';
import { resolveColor, themeOf } from './theme.ts';

/** Label kinds `style.labels` colours, and the shape roles each one covers. */
export const LABEL_KINDS = ['value', 'total', 'category', 'series', 'axis', 'decoration'] as const;
/** Line kinds `style.lines` styles, and the shape roles each one covers. */
export const LINE_KINDS = ['axis', 'gridline', 'connector', 'leader', 'decoration'] as const;
export type LabelKind = typeof LABEL_KINDS[number];
export type LineKind = typeof LINE_KINDS[number];
/** Limits on a line width (points). */
export const LINE_WIDTH = { min: 0.25, max: 6 } as const;

const DECORATION = /^(arrow:|valueline:|indicator|bracket:|today)/;

function labelKind(role: string): LabelKind | undefined {
  if (role.startsWith('label:value')) return 'value';
  if (role.startsWith('label:total')) return 'total';
  if (role.startsWith('label:category')) return 'category';
  if (role.startsWith('label:series') || role.startsWith('legend')) return 'series';
  if (role === 'axis' || role.startsWith('axis:')) return 'axis';
  if (DECORATION.test(role)) return 'decoration';
  return undefined;
}

function lineKind(role: string): LineKind | undefined {
  if (role === 'axis' || role.startsWith('axis:')) return 'axis';
  if (role.startsWith('gridline')) return 'gridline';
  if (role.startsWith('connector')) return 'connector';
  if (role.startsWith('leader')) return 'leader';
  if (DECORATION.test(role)) return 'decoration';
  return undefined;
}

/**
 * Applies `style.labels` and `style.lines` to a laid-out scene, in place. `style` is the resolved
 * style the scene was laid out with (colours already "hex@slot"). A label is recoloured only when it
 * carries one of the style's neutral text colours (text, muted, axis): labels the engine coloured for
 * contrast on a fill, or in a series colour, keep theirs.
 */
export function applyLooks(scene: Scene, style: Style): void {
  const { labels, lines } = style;
  if (!labels && !lines) return;
  const theme = themeOf(style);
  const r = (c: string) => resolveColor(c, theme);
  const neutral = new Set([style.text, style.muted, style.axis]);
  const roles = scene.roles ?? [];
  scene.shapes.forEach((s: Shape, i) => {
    const role = roles[i] ?? '';
    if (s.kind === 'text') {
      const k = labelKind(role);
      const c = k && labels?.[k];
      if (c && neutral.has(s.color)) s.color = r(c);
      // A decoration bubble's outline follows the decoration line colour.
      const look = lines?.decoration;
      if (s.stroke && look?.color && k === 'decoration') s.stroke = r(look.color);
      return;
    }
    const look: LineLook | undefined = lines?.[lineKind(role) ?? ('' as LineKind)];
    if (!look) return;
    if (s.kind === 'line') {
      if (look.color) s.color = r(look.color);
      if (look.width !== undefined) s.width = look.width;
      if (look.dash !== undefined) s.dash = look.dash;
    } else if (s.kind === 'path') {
      // Arrow heads and value-line markers are filled paths in the line colour.
      if (look.color && s.fill) s.fill = r(look.color);
      if (look.color && s.stroke) s.stroke = r(look.color);
    }
  });
}
