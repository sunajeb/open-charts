import type { Scene, TextShape } from './types.ts';
import { overlaps } from './place.ts';

export interface Issue { kind: 'overlap' | 'out-of-frame'; detail: string }

/** Layout quality gate: no two labels overlap and nothing spills outside the chart frame. */
export function checkScene(scene: Scene, tol = 1): Issue[] {
  const issues: Issue[] = [];
  const f = scene.frame;
  const texts = scene.shapes.filter((s): s is TextShape => s.kind === 'text');
  for (const t of texts) {
    if (t.x < f.x - tol || t.y < f.y - tol || t.x + t.w > f.x + f.w + tol || t.y + t.h > f.y + f.h + tol) {
      issues.push({ kind: 'out-of-frame', detail: `"${t.text}"` });
    }
  }
  for (let i = 0; i < texts.length; i++) {
    for (let j = i + 1; j < texts.length; j++) {
      if (overlaps(texts[i], texts[j], tol)) {
        issues.push({ kind: 'overlap', detail: `"${texts[i].text}" × "${texts[j].text}"` });
      }
    }
  }
  return issues;
}
