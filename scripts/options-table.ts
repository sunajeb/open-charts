// The options table of docs/DATA-FORMAT.md, generated from the applicability table and the per-type
// defaults in src/core/spec.ts (test/docs.test.ts checks the doc holds exactly this table).
// Run to print it: node --experimental-strip-types --no-warnings scripts/options-table.ts
import { CHART_TYPES, OPTIONS, OPTION_KEYS, optionDefaults, type OptionKey } from '../src/core/spec.ts';

const show = (v: unknown) => (typeof v === 'boolean' ? (v ? 'on' : 'off') : `\`${v}\``);

function typesCell(k: OptionKey): string {
  const types: readonly string[] = OPTIONS[k].types;
  if (types.length === CHART_TYPES.length) return 'all';
  const missing = CHART_TYPES.filter((t) => !types.includes(t));
  return missing.length <= 3 ? `all except ${missing.join(', ')}` : types.join(', ');
}

function defaultCell(k: OptionKey): string {
  const byValue = new Map<string, string[]>();
  for (const t of OPTIONS[k].types) {
    const v = (optionDefaults(t) as Record<string, unknown>)[k];
    if (v === undefined) continue;
    const key = show(v);
    byValue.set(key, [...(byValue.get(key) ?? []), t]);
  }
  if (!byValue.size) return OPTIONS[k].auto ?? 'none';
  if (byValue.size === 1) return [...byValue.keys()][0];
  // The value most types share is "otherwise"; the others are listed by type.
  const groups = [...byValue].sort((a, b) => a[1].length - b[1].length);
  const rest = groups.pop()!;
  return `${groups.map(([v, ts]) => `${v} for ${ts.join(', ')}`).join('; ')}; ${rest[0]} otherwise`;
}

export function optionsTable(): string {
  const rows = OPTION_KEYS.map((k) => `| [\`${k}\`](#${k.toLowerCase()}) | ${OPTIONS[k].value} | ${typesCell(k)} | ${defaultCell(k)} |`);
  return ['| Option | Value | Chart types | Default |', '|---|---|---|---|', ...rows].join('\n');
}

if (import.meta.url === `file://${process.argv[1]}`) console.log(optionsTable());
