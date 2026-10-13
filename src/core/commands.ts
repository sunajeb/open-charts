// Chart commands of the PowerPoint add-in's ribbon and right-click menu that change a chart's spec
// without its data: totals, CAGR / difference arrows, same scale, reset label position. Pure: each
// returns the new spec (the input is not changed) or a message saying why the command does not
// apply. The add-in engine (src/addin/engine.ts op "command") lays the result out like any update.
import type { ChartSpec } from './types.ts';
import { appliesTo, optionDefaults } from './spec.ts';
import { kindOf, unpinElement } from './overrides.ts';

export const CHART_COMMANDS = ['totals', 'cagr', 'difference', 'unpin', 'scaleGroup'] as const;
export type ChartCommand = typeof CHART_COMMANDS[number];

export type CommandResult<T> = { spec: T } | { error: string };

const NAMES: Record<string, string> = { totals: 'Totals', cagr: 'A CAGR arrow', difference: 'A difference arrow', scaleGroup: 'Same scale' };

/** Category indices named by the selected elements' roles (`…:c3` → 3), sorted, unique. */
export function selectedCategories(roles: string[]): number[] {
  const out = new Set<number>();
  for (const r of roles) {
    const m = /(?:^|:)c(\d+)(?=:|$)/.exec(r);
    if (m) out.add(Number(m[1]));
  }
  return [...out].sort((a, b) => a - b);
}

/**
 * The arrow's span: the first and last selected categories when two or more are selected (the user
 * clicked the columns to connect), else the first and last category of the chart.
 */
function span(spec: ChartSpec, roles: string[]): { from: number; to: number } | undefined {
  const n = spec.categories?.length ?? 0;
  const sel = selectedCategories(roles).filter((c) => c < n);
  if (sel.length >= 2) return { from: sel[0], to: sel[sel.length - 1] };
  return n >= 2 ? { from: 0, to: n - 1 } : undefined;
}

/** True when the chart draws the option now (set in the spec, or on by default). */
export function hasOption(spec: ChartSpec, key: 'totals' | 'cagr' | 'difference'): boolean {
  const v = spec.options?.[key];
  if (key === 'totals') return (v ?? optionDefaults(spec.type).totals) === true;
  return !!v;
}

/**
 * Runs `cmd` on `spec`. `roles` are the OC_ROLE values of the selected elements ([] when the whole
 * chart is selected). `arg`: the group name for scaleGroup ("" removes the chart from its group).
 *   totals      turns the totals on or off (off when they are drawn now)
 *   cagr        adds a CAGR arrow (between the selected columns, else first to last) or removes it;
 *               with other columns selected than the arrow's, moves it there
 *   difference  the same for the difference arrow
 *   unpin       resets the selected labels' positions (all pinned labels when none is selected)
 *   scaleGroup  puts the chart in same-scale group `arg`, or takes it out
 */
export function runCommand<T extends ChartSpec>(spec: T, cmd: string, roles: string[] = [], arg = ''): CommandResult<T> {
  if (!(CHART_COMMANDS as readonly string[]).includes(cmd)) return { error: `Unknown chart command "${cmd}".` };
  const c = cmd as ChartCommand;
  if (c === 'unpin') {
    const labels = roles.filter((r) => kindOf(r) === 'label');
    const pinned = Object.keys(spec.overrides ?? {}).filter((k) => spec.overrides![k]?.pin);
    const targets = labels.length ? labels.filter((r) => pinned.includes(r)) : pinned;
    if (!targets.length) return { error: labels.length ? 'This label is where open-charts puts it already.' : 'No label of this chart has been moved.' };
    return { spec: targets.reduce((s, r) => unpinElement(s, r), spec) };
  }
  const key = c;
  if (!appliesTo(key, spec.type)) return { error: `${NAMES[key]} ${key === 'totals' ? 'are' : 'is'} not available for this chart type.` };
  const options = { ...(spec.options ?? {}) } as Record<string, unknown>;
  if (key === 'scaleGroup') {
    if (arg) options.scaleGroup = arg; else delete options.scaleGroup;
  } else if (key === 'totals') {
    options.totals = !hasOption(spec, 'totals');
  } else {
    // Selected columns other than the arrow's own move it there; otherwise the command toggles it.
    const now = options[key] as { from: number; to: number } | null | undefined;
    const s = span(spec, roles);
    const moved = now && s && selectedCategories(roles).length >= 2 && (now.from !== s.from || now.to !== s.to);
    if (now && !moved) delete options[key];
    else if (!s) return { error: `${NAMES[key]} needs at least two categories.` };
    else options[key] = { ...(now ?? {}), ...s };
  }
  const out = { ...spec, options } as T;
  if (!Object.keys(options).length) delete (out as { options?: unknown }).options;
  return { spec: out };
}

/** A same-scale group name not used by `taken` ("scale1", "scale2", …). */
export function newScaleGroup(taken: Iterable<string>): string {
  const used = new Set(taken);
  for (let i = 1; ; i++) if (!used.has(`scale${i}`)) return `scale${i}`;
}
