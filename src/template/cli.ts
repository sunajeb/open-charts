// `open-charts fill template.pptx data.json -o out.pptx [--dry-run] [--strict] [--style s.json] [--each items.json] [--report r.json]`
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname } from 'node:path';
import type { Measurer } from '../core/text.ts';
import type { StyleFile } from '../core/style-file.ts';
import { fillTemplate, type FillReport } from './fill.ts';
import { checkValues, FILL_LIMITS, type FillData, type FillValues } from './data.ts';
import type { SpecError } from '../core/spec.ts';

export interface FillCliOptions {
  out?: string;
  dryRun: boolean;
  strict: boolean;
  each?: string;
  report?: string;
  style?: StyleFile;
  measurer: Measurer;
}

/** Runs one fill (or one per `--each` item); returns the process exit code. */
export function fillMain(templatePath: string, dataPath: string, o: FillCliOptions): number {
  const template = new Uint8Array(readFileSync(templatePath));
  const data = JSON.parse(readFileSync(dataPath, 'utf8')) as FillData;
  const out = o.out ?? templatePath.replace(/\.pptx$/i, '-filled.pptx');

  // --each: one deck per item, the item's values over the data's; "{name}" in -o takes item values.
  const runs: { values?: FillValues; out: string }[] = [];
  if (o.each) {
    const items = JSON.parse(readFileSync(o.each, 'utf8')) as unknown;
    const errors: SpecError[] = [];
    if (!Array.isArray(items) || !items.length) errors.push({ path: '/', message: 'must be a non-empty array of values, one per deck' });
    else if (items.length > FILL_LIMITS.each) errors.push({ path: '/', message: `at most ${FILL_LIMITS.each} items` });
    else items.forEach((v, i) => checkValues(v, `/${i}`, errors));
    if (errors.length) throw new Error(`Invalid ${o.each}:\n${errors.map((e) => `  ${e.path}: ${e.message}`).join('\n')}`);
    if (!/\{[A-Za-z0-9_.-]+\}/.test(out)) throw new Error('--each writes one deck per item: put a field in the output name, e.g. -o "out/{region}.pptx"');
    for (const v of items as FillValues[]) runs.push({ values: v, out: out.replace(/\{([A-Za-z0-9_.-]+)\}/g, (m, k: string) => (typeof v[k] === 'string' || typeof v[k] === 'number' ? safeName(String(v[k])) : m)) });
    const dup = runs.map((r) => r.out).find((p, i, a) => a.indexOf(p) !== i);
    if (dup) throw new Error(`--each: two items write ${dup}; use a field that differs per item in -o`);
  } else runs.push({ out });

  const reports: (FillReport & { out?: string })[] = [];
  let problems = 0;
  for (const r of runs) {
    const d: FillData = r.values ? { ...data, values: { ...data.values, ...r.values } } : data;
    const { pptx, report } = fillTemplate(template, d, { measurer: o.measurer, style: o.style, dryRun: o.dryRun });
    reports.push({ ...report, ...(o.dryRun ? {} : { out: r.out }) });
    problems += report.problems;
    if (runs.length > 1 || o.dryRun) console.log(`${o.dryRun ? 'Would write' : 'Writing'} ${r.out}`);
    console.log(formatFillReport(report, o.dryRun));
    if (pptx) {
      mkdirSync(dirname(r.out), { recursive: true });
      writeFileSync(r.out, pptx);
      console.log(`Wrote ${r.out} (${report.slides.length} slide${report.slides.length === 1 ? '' : 's'})`);
    }
  }
  if (o.report) {
    mkdirSync(dirname(o.report), { recursive: true });
    writeFileSync(o.report, `${JSON.stringify(reports.length === 1 ? reports[0] : reports, null, 2)}\n`);
  }
  if (o.strict && problems) {
    console.error(`--strict: ${problems} problem${problems === 1 ? '' : 's'} (missing values, chart errors or layout issues)`);
    return 1;
  }
  return 0;
}

const safeName = (s: string) => s.replace(/[\\/:*?"<>|\u0000-\u001f]+/g, '-').trim() || '-';

const f1 = (v: number) => String(Math.round(v * 10) / 10);

/** The report as text: per output slide its template slide, charts (frame, status) and fields. */
export function formatFillReport(r: FillReport, dryRun = false): string {
  const lines: string[] = [];
  lines.push(`Template: ${r.template.slides} slides; theme accent1 ${r.template.theme.accent1 ?? '?'}, font ${r.template.font ?? '?'}`);
  for (const s of r.slides) {
    const parts: string[] = [];
    for (const c of s.charts) {
      const at = c.frame ? ` at ${f1(c.frame.x)},${f1(c.frame.y)} ${f1(c.frame.w)}x${f1(c.frame.h)}` : '';
      parts.push(`    chart:${c.name}${at}: ${c.status === 'filled' ? (dryRun ? 'would fill' : 'filled') : c.status}${c.detail ? ` (${c.detail})` : ''}`);
      for (const i of c.issues) parts.push(`      ${i}`);
    }
    const filled = s.fields.filter((x) => x.status === 'filled').map((x) => x.name);
    const missing = s.fields.filter((x) => x.status === 'missing').map((x) => x.name);
    if (filled.length) parts.push(`    fields: ${[...new Set(filled)].join(', ')}`);
    if (missing.length) parts.push(`    missing: ${[...new Set(missing)].map((x) => `{{${x}}}`).join(', ')}`);
    lines.push(`  slide ${s.index} (template slide ${s.source})${parts.length ? '' : ': no slots, unchanged'}`, ...parts);
  }
  if (r.unused.length) lines.push(`  unused values: ${r.unused.join(', ')}`);
  for (const w of r.warnings) lines.push(`  warning: ${w}`);
  lines.push(r.ok ? '  OK' : `  ${r.problems} problem${r.problems === 1 ? '' : 's'}`);
  return lines.join('\n');
}
