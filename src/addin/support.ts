// Help for the user and for support: the first-run welcome, the sample deck and "Report a problem".
//
// Engine requests (src/addin/engine.ts), answered with M records the add-in shows:
//   welcome  the welcome text (ends with "Open a sample deck now?"); remembered as shown in
//            settings.json ("welcomed": true), so the add-in shows it by itself only once
//   sample   writes a deck with one sample chart per chart type into the add-in's home and opens it
//   report   arg "spec" with the selected chart's OC_SPEC after the params (only when the user agreed),
//            app=<PowerPoint version and OS>: writes the support zip into <home>/reports and shows it
//            in Finder / Explorer
// The zip holds the add-in log, settings, addin.json and versions. Never the token or slide content.
import { spawn } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { arch, release } from 'node:os';
import { join } from 'node:path';
import { isSea } from 'node:sea';
import { zip, type ZipInput } from '../render/zip.ts';
import { buildDeck } from '../deck.ts';
import type { Measurer } from '../core/index.ts';
import { ENGINE_VERSION } from '../version.ts';
import { writeSettings } from '../update.ts';
import { FORMAT_VERSION, record } from './format.ts';
import { addinHome } from './launcher.ts';
import { flatLabel, galleryItems, sampleFor } from './registry.ts';

export const GUIDE_URL = 'https://github.com/sunajeb/open-charts/blob/main/docs/USER-GUIDE.md';
export const ISSUES_URL = 'https://github.com/sunajeb/open-charts/issues';
/** The log's last bytes that go into a report (it can grow large with verbose logging on). */
export const LOG_TAIL = 1024 * 1024;

export const WELCOME = [
  'Welcome to open-charts: native charts you can edit, right in PowerPoint.',
  '',
  'Everything is on the open-charts tab, left of Home.',
  'Insert a chart: pick a type from the Insert chart gallery. Then click Edit data to type your numbers.',
  '',
  `User guide: ${GUIDE_URL}`,
  'To see this again: Welcome, on the right of the open-charts tab.',
].join('\n');

/** Opens a file with its app (PowerPoint for a .pptx) and shows a file in Finder / Explorer. */
export interface Shell { open(path: string): void; reveal(path: string): void }

/** Opens a file with its app, or (`reveal`) selects it in Finder / Explorer. Explorer takes the
 * quotes as written (verbatim arguments): `/select,"<path>"`. */
export function fileCommand(path: string, reveal: boolean, platform = process.platform): [string, string[]] {
  if (platform === 'win32') return ['explorer.exe', [reveal ? `/select,"${path}"` : `"${path}"`]];
  return ['open', reveal ? ['-R', path] : [path]];
}

const run = ([cmd, args]: [string, string[]]) => spawn(cmd, args, { stdio: 'ignore', detached: true, windowsHide: true, windowsVerbatimArguments: true }).unref();
export const systemShell: Shell = { open: (p) => run(fileCommand(p, false)), reveal: (p) => run(fileCommand(p, true)) };

const answer = (...messages: string[]) => [record('OC', FORMAT_VERSION), ...messages.map((m) => record('M', m))].join('\n') + '\n';

/** The welcome text, now remembered as shown. */
export function welcome(home = addinHome()): string {
  try { writeSettings(home, { welcomed: true }); } catch { /* shown again next time: harmless */ }
  return WELCOME;
}

/** Writes the sample deck (one slide per chart type, sample data) into `home`. Returns its path. */
export function writeSampleDeck(m: Measurer, home = addinHome()): string {
  const slides = galleryItems().flatMap((g) => {
    const s = sampleFor(g.id);
    return 'spec' in s ? [{ title: `${flatLabel(g)}: select it, then click Edit data`, chart: s.spec }] : [];
  });
  const path = join(home, 'open-charts-sample.pptx');
  mkdirSync(home, { recursive: true });
  try {
    writeFileSync(path, buildDeck({ slides }, m).pptx);
  } catch (err) {
    if (!existsSync(path)) throw err; // still open in PowerPoint (Windows locks it): open that one
  }
  return path;
}

export interface ReportInput { home?: string; app?: string; spec?: string; now?: Date }

/** What goes into the report: name → contents. No token, no slide content. */
export function reportFiles({ home = addinHome(), app = '', spec, now = new Date() }: ReportInput): ZipInput[] {
  const read = (name: string) => { try { return readFileSync(join(home, name)); } catch { return undefined; } };
  const files: ZipInput[] = [];
  const log = read('addin.log');
  if (log) files.push({ name: 'addin.log', data: log.subarray(Math.max(0, log.length - LOG_TAIL)) });
  for (const name of ['settings.json', 'addin.json']) {
    const data = read(name);
    if (data) files.push({ name, data });
  }
  const versions = {
    openCharts: ENGINE_VERSION,
    engine: isSea() ? 'packaged' : 'node',
    node: process.versions.node,
    os: `${process.platform} ${release()} ${arch()}`,
    powerPoint: app || powerPointVersion() || 'unknown',
    created: now.toISOString(),
  };
  files.push({ name: 'versions.json', data: `${JSON.stringify(versions, null, 2)}\n` });
  if (spec) files.push({ name: 'chart-spec.json', data: spec.endsWith('\n') ? spec : `${spec}\n` });
  return files;
}

/** PowerPoint for Mac's version from its Info.plist (the add-in sends its own; this is for the CLI). */
function powerPointVersion(): string | undefined {
  if (process.platform !== 'darwin') return undefined;
  try {
    const plist = readFileSync('/Applications/Microsoft PowerPoint.app/Contents/Info.plist', 'utf8');
    return /<key>CFBundleShortVersionString<\/key>\s*<string>([^<]+)</.exec(plist)?.[1];
  } catch { return undefined; }
}

/** Writes the report zip into <home>/reports. Returns its path. */
export function writeReport(input: ReportInput = {}): string {
  const home = input.home ?? addinHome(), now = input.now ?? new Date();
  const stamp = now.toISOString().slice(0, 19).replace(/[-:]/g, '').replace('T', '-');
  const dir = join(home, 'reports');
  mkdirSync(dir, { recursive: true });
  const path = join(dir, `open-charts-report-${stamp}.zip`);
  writeFileSync(path, zip(reportFiles({ ...input, home, now })));
  return path;
}

/** What the user is told about a report written to `path`. */
export const reportMessage = (path: string, withSpec: boolean) =>
  `Saved a problem report:\n${path}\n\nIt holds the add-in's log, settings and version numbers${withSpec ? ', and the selected chart\'s data' : ''}. No slide content.\n\nPlease attach it to a new issue at ${ISSUES_URL} and say what you did and what went wrong.`;

/** The welcome, sample and report requests (engine.ts). */
export function handleSupport(op: string, arg: string, params: { app?: string; spec?: string }, m: Measurer, home = addinHome(), shell: Shell = systemShell): string {
  switch (op) {
    case 'welcome':
      return answer(`${welcome(home)}\n\nOpen a sample deck now?`);
    case 'sample':
      shell.open(writeSampleDeck(m, home));
      return answer();
    default: {
      const spec = arg === 'spec' && params.spec ? params.spec : undefined;
      const path = writeReport({ home, app: params.app, spec });
      shell.reveal(path);
      return answer(reportMessage(path, !!spec));
    }
  }
}
