// "Update available": the client side of open-charts updates (docs/release.md, "Updates").
//
// The installed engine asks GitHub Releases for the latest release, at most once a day, and only
// when the user turned it on (`open-charts addin update on`, i.e. "updateCheck": true in
// <add-in home>/settings.json; off by default: an install never calls out on its own). It only
// tells: the result is logged and served at GET /api/addin/update. Installing (`open-charts addin
// update install`, installUpdate) is the user's own action: it downloads this platform's installer
// and SHA256SUMS.txt, checks the hash and the publisher's signature, and only then opens the
// installer, which the user confirms. An IT policy file can force the check off (policyUpdateOff).
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { ENGINE_VERSION } from './version.ts';

export const RELEASES_URL = 'https://api.github.com/repos/sunajeb/open-charts/releases/latest';
export const CHECK_EVERY_MS = 24 * 60 * 60 * 1000;

/** settings.json keys this file and the welcome (src/addin/support.ts) write. */
export interface Settings { updateCheck?: boolean; welcomed?: boolean }
export interface UpdateInfo {
  current: string;
  latest: string;
  available: boolean;
  /** The release page (what the user is sent to). */
  url: string;
  /** This platform's installer in the release, if it has one, and where to download it and its checksums. */
  asset?: string;
  assetUrl?: string;
  sumsUrl?: string;
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
  const assets = r.assets ?? [];
  const asset = installerAsset(assets.map((a) => a.name), opts.platform);
  const urlOf = (name?: string) => assets.find((a) => a.name === name)?.browser_download_url;
  const assetUrl = urlOf(asset), sumsUrl = urlOf('SHA256SUMS.txt');
  return {
    current, latest,
    available: !r.prerelease && compareVersions(latest, current) > 0,
    url: r.html_url ?? `https://github.com/sunajeb/open-charts/releases/tag/${r.tag_name}`,
    ...(asset ? { asset } : {}), ...(assetUrl ? { assetUrl } : {}), ...(sumsUrl ? { sumsUrl } : {}),
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
export async function maybeCheckForUpdate(home: string, opts: Parameters<typeof checkForUpdate>[0] & { policy?: string } = {}): Promise<UpdateInfo | undefined> {
  if (readSettings(home).updateCheck !== true || policyUpdateOff(opts.policy)) return undefined;
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

// ---- IT policy ---------------------------------------------------------------------------------

/** The machine-wide policy file IT can deploy (docs/IT-DEPLOYMENT.md); users cannot write there. */
export function policyFile(platform = process.platform, env = process.env): string | undefined {
  if (platform === 'darwin') return '/Library/Application Support/open-charts/policy.json';
  if (platform === 'win32') return `${env.ProgramData ?? 'C:\\ProgramData'}\\open-charts\\policy.json`;
  return undefined;
}

/** True when the policy file says {"updateCheck": false}: no check, no update offered. */
export function policyUpdateOff(file = policyFile()): boolean {
  try { return !!file && (JSON.parse(readFileSync(file, 'utf8')) as Settings).updateCheck === false; } catch { return false; }
}

// ---- Installing an update ----------------------------------------------------------------------

/**
 * Who signs open-charts' installers: the Apple Developer ID team ID and the Windows Authenticode
 * subject. Empty until the signing accounts exist (docs/release.md); while empty, no download passes
 * the signature check, so nothing unsigned is ever opened.
 */
export const SIGNER = { appleTeamId: '', windowsSubject: '' };

/** The SHA-256 listed for `name` in a SHA256SUMS.txt ("<hex>  <name>" lines), lower case. */
export function expectedSha256(sums: string, name: string): string | undefined {
  for (const line of sums.split(/\r?\n/)) {
    const m = /^([0-9a-fA-F]{64})\s+\*?(.+)$/.exec(line.trim());
    if (m && m[2] === name) return m[1].toLowerCase();
  }
  return undefined;
}

/** Whether the installer at `file` is signed by SIGNER: '' when it is, else why not. */
export function signatureProblem(file: string, platform = process.platform, signer = SIGNER, run = (cmd: string, args: string[]) => execFileSync(cmd, args, { encoding: 'utf8', timeout: 60_000 })): string {
  if (platform === 'darwin') {
    if (!signer.appleTeamId) return 'this version of open-charts does not know its signing identity yet';
    let out = '';
    try { out = run('/usr/sbin/pkgutil', ['--check-signature', file]); } catch (e) { return `pkgutil: ${(e as Error).message.split('\n')[0]}`; }
    if (!out.includes(`(${signer.appleTeamId})`) || !/Developer ID Installer/.test(out)) return 'the package is not signed by open-charts';
    try { run('/usr/sbin/spctl', ['--assess', '--type', 'install', file]); } catch { return 'macOS Gatekeeper does not accept the package'; }
    return '';
  }
  if (platform === 'win32') {
    if (!signer.windowsSubject) return 'this version of open-charts does not know its signing identity yet';
    let out = '';
    const ps = `$s = Get-AuthenticodeSignature -LiteralPath '${file.replace(/'/g, "''")}'; "$($s.Status)|$($s.SignerCertificate.Subject)"`;
    try { out = run('powershell.exe', ['-NoProfile', '-Command', ps]).trim(); } catch (e) { return `Get-AuthenticodeSignature: ${(e as Error).message.split('\n')[0]}`; }
    const [status, subject = ''] = out.split('|');
    return status === 'Valid' && subject.includes(signer.windowsSubject) ? '' : `the installer's signature is ${status || 'missing'}`;
  }
  return 'updates are installed on macOS and Windows only';
}

export interface InstallDeps {
  fetch?: typeof fetch;
  platform?: string;
  dir?: string;
  verify?: (file: string) => string;
  open?: (file: string) => void;
}

/**
 * Downloads the update's installer and SHA256SUMS.txt, checks the hash and the signature, then
 * opens the installer (the user confirms it there). Anything that fails a check is deleted and
 * never opened. Returns the installer's path.
 */
export async function installUpdate(info: UpdateInfo, deps: InstallDeps = {}): Promise<string> {
  if (!info.available) throw new Error(`open-charts ${info.current} is up to date.`);
  if (!info.asset || !info.assetUrl || !info.sumsUrl) throw new Error(`The release has no installer for this computer with checksums. Download it from ${info.url}`);
  if (policyUpdateOff()) throw new Error('Updates are managed by your IT department.');
  const get = async (url: string) => {
    const res = await (deps.fetch ?? fetch)(url, { headers: { 'user-agent': `open-charts/${info.current}` }, signal: AbortSignal.timeout(300_000) });
    if (!res.ok) throw new Error(`Download failed: HTTP ${res.status} for ${url}`);
    return new Uint8Array(await res.arrayBuffer());
  };
  const want = expectedSha256(new TextDecoder().decode(await get(info.sumsUrl)), info.asset);
  if (!want) throw new Error(`SHA256SUMS.txt does not list ${info.asset}. Download it from ${info.url}`);
  const bytes = await get(info.assetUrl);
  const dir = deps.dir ?? join(tmpdir(), 'open-charts-update');
  mkdirSync(dir, { recursive: true });
  const file = join(dir, info.asset);
  writeFileSync(file, bytes);
  const fail = (why: string) => { rmSync(file, { force: true }); return new Error(`The downloaded installer was not opened: ${why}. Download it from ${info.url}`); };
  if (createHash('sha256').update(bytes).digest('hex') !== want) throw fail('its checksum does not match');
  const platform = deps.platform ?? process.platform;
  const problem = (deps.verify ?? ((f) => signatureProblem(f, platform)))(file);
  if (problem) throw fail(problem);
  (deps.open ?? ((f) => (platform === 'win32' ? execFileSync('msiexec.exe', ['/i', f]) : execFileSync('/usr/bin/open', [f]))))(file);
  return file;
}
