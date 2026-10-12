// Machine-readable layout report for `--report` and the `--strict` gate.
import type { BuiltSlide } from './deck.ts';

export interface Issue { kind: string; detail: string }
export interface Report {
  ok: boolean;
  issues: number;
  slides: { slide: number; title?: string; issues: Issue[] }[];
}

/** Issue strings are "kind: detail" (see buildDeck). */
export function parseIssue(s: string): Issue {
  const i = s.indexOf(': ');
  return i < 0 ? { kind: 'issue', detail: s } : { kind: s.slice(0, i), detail: s.slice(i + 2) };
}

export function buildReport(slides: Pick<BuiltSlide, 'title' | 'issues'>[]): Report {
  const out = slides.map((s, i) => ({ slide: i + 1, ...(s.title !== undefined ? { title: s.title } : {}), issues: s.issues.map(parseIssue) }));
  const issues = out.reduce((a, s) => a + s.issues.length, 0);
  return { ok: issues === 0, issues, slides: out };
}
