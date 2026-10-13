// "Update available": the client side of open-charts updates (docs/release.md, "Updates").
//
// The installed engine asks GitHub Releases for the latest release, at most once a day, and only
// when the user turned it on (`open-charts addin update on`, i.e. "updateCheck": true in
// <add-in home>/settings.json; off by default: an install never calls out on its own). It only
// tells: the result is logged and served at GET /api/addin/update. It never downloads or runs
// anything; a release is installed by the user from its signed package.
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { ENGINE_VERSION } from './version.ts';

export const RELEASES_URL = 'https://api.github.com/repos/sunajeb/open-charts/releases/latest';
export const CHECK_EVERY_MS = 24 * 60 * 60 * 1000;

export interface Settings { updateCheck?: boolean }
export interface UpdateInfo {
  current: string;
  latest: string;
  available: boolean;
  /** The release page (what the user is sent to). */
  url: string;
  /** This platform's installer in the release, if it has one. */
  asset?: string;
  checkedAt: string;
}

/** a.b.c[-pre] compared numerically; a pre-release is older than its release. 1, 0 or -1. */
export function compareVersions(a: string, b: string): number {
  const parse = (v: string) => {
    const m = /^v?(\d+)\.(\d+)\.(\d+)(?:-([0-9A-Za-z.-]+))?/.exec(v.trim());
    if (!m) throw new Error(`not a version: ${v}`);
    return { nums: [Number(m[1]), Number(m[2]), Number(m[3])], pre: m[4] };
  };
  const x = parse(a), y = parse(b);
  for (let i = 0; i < 3; i++) if (x.nums[i] !== y.nums[i]) return x.nums[i] > y.nums[i] ? 1 : -1;
  if (x.pre === y.pre) return 0;
  if (!x.pre) return 1;
  if (!y.pre) return -1;
  return x.pre > y.pre ? 1 : -1;
}

/** The installer a user of `platform` would take from a release (the per-user one on Windows). */
export function installerAsset(names: string[], platform = process.platform): string | undefined {
  if (platform === 'darwin') return names.find((n) => /^open-charts-.*\.pkg$/.test(n));
  if (platform === 'win32') return names.find((n) => /^open-charts-[^-]+(-[0-9A-Za-z.]+)?\.msi$/.test(n) && !n.endsWith('-machine.msi'));
  return undefined;
}

interface Release { tag_name?: string; html_url?: string; draft?: boolean; prerelease?: boolean; assets?: { name: string; browser_download_url: string }[] }

/** Asks the releases endpoint now. */
export async function checkForUpdate(opts: { current?: string; url?: string; fetch?: typeof fetch; platform?: string; now?: Date } = {}): Promise<UpdateInfo> {
  const current = opts.current ?? ENGINE_VERSION;
  const url = opts.url ?? process.env.OPEN_CHARTS_RELEASES_URL ?? RELEASES_URL;
  const res = await (opts.fetch ?? fetch)(url, {
    headers: { accept: 'application/vnd.github+json', 'user-agent': `open-charts/${current}` },
    signal: AbortSignal.timeout(10_000),
  });
  if (res.status === 404) throw new Error(`release check: no published release at ${url}`);
  if (!res.ok) throw new Error(`release check: HTTP ${res.status} from ${url}`);
  const r = await res.json() as Release;
  if (!r.tag_name || r.draft) throw new Error('release check: no published release');
  const latest = r.tag_name.replace(/^v/, '');
  const asset = installerAsset((r.assets ?? []).map((a) => a.name), opts.platform);
  return {
    current, latest,
    available: !r.prerelease && compareVersions(latest, current) > 0,
    url: r.html_url ?? `https://github.com/sunajeb/open-charts/releases/tag/${r.tag_name}`,
    ...(asset ? { asset } : {}),
    checkedAt: (opts.now ?? new Date()).toISOString(),
  };
}

export const settingsFile = (home: string) => join(home, 'settings.json');
const updateFile = (home: string) => join(home, 'update.json');

export function readSettings(home: string): Settings {
  try { return JSON.parse(readFileSync(settingsFile(home), 'utf8')) as Settings; } catch { return {}; }
}

export function writeSettings(home: string, change: Settings) {
  mkdirSync(home, { recursive: true });
  writeFileSync(settingsFile(home), `${JSON.stringify({ ...readSettings(home), ...change }, null, 2)}\n`);
}

/** The last check's result, if there was one. */
export function lastUpdateCheck(home: string): UpdateInfo | undefined {
  try { return JSON.parse(readFileSync(updateFile(home), 'utf8')) as UpdateInfo; } catch { return undefined; }
}

/**
 * The background check (the add-in server runs it as it starts): nothing unless the setting is on,
 * at most once per CHECK_EVERY_MS, never throws. Returns the (possibly cached) result.
 */
export async function maybeCheckForUpdate(home: string, opts: Parameters<typeof checkForUpdate>[0] = {}): Promise<UpdateInfo | undefined> {
  if (readSettings(home).updateCheck !== true) return undefined;
  const now = opts.now ?? new Date();
  const last = lastUpdateCheck(home);
  if (last && (last.current === (opts.current ?? ENGINE_VERSION)) && now.getTime() - Date.parse(last.checkedAt) < CHECK_EVERY_MS) return last;
  try {
    const info = await checkForUpdate({ ...opts, now });
    mkdirSync(home, { recursive: true });
    writeFileSync(updateFile(home), `${JSON.stringify(info, null, 2)}\n`);
    return info;
  } catch {
    return existsSync(updateFile(home)) ? last : undefined; // offline, rate-limited: try again next start
  }
}
