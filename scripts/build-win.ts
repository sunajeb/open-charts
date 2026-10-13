// `npm run build:win`: the Windows package in dist/win/ (docs/windows.md):
//   open-charts.exe   the engine as one Node single-executable application (SEA): no Node on the PC
//   open-charts.ppam  the add-in (the same one the Mac installs: one ribbon, VBA and tag contract)
//   addin.json        its callback prefix, for the server's Apply
//   install.ps1, uninstall.ps1, e2e-powerpoint-win.ps1
//
// How the exe is made (dev-time only; works on macOS, Linux and Windows):
//   1. esbuild bundles src/cli.ts into one CommonJS file. Every `import.meta.url` becomes the URL of
//      the source file it was in, under a resource root: the repo for the plain bundle, and for the
//      exe a folder its embedded assets (editor page, VBA source, ribbon icons) are unpacked into on first start.
//   2. `node --experimental-sea-config` turns it into a SEA blob (no code cache or snapshot, so the
//      blob is the same on every platform).
//   3. The official win-x64 node.exe of the very same Node version (a blob only runs in the Node
//      that made it) is downloaded from nodejs.org, checked against the release's SHASUMS256.txt
//      and cached; postject injects the blob; the now invalid Authenticode signature is removed.
//
//   --host   also builds dist/host/open-charts for this machine (macOS/Linux), to try the exe's
//            code paths (asset unpacking, self-spawn) without a Windows PC.
//   --skip-exe   bundle + ppam only (no download).
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { chmodSync, copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { dirname, join, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildAddin } from '../src/addin/package.ts';

const ROOT = fileURLToPath(new URL('../', import.meta.url));
const OUT = join(ROOT, 'dist', 'win');
const BUNDLE = join(OUT, 'open-charts.cjs');
const args = new Set(process.argv.slice(2));
export const SEA_FUSE = 'NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2';

/** Files the engine reads at run time (by import.meta.url), embedded in the exe as SEA assets. */
export function assetFiles(root = ROOT): string[] {
  const walk = (dir: string): string[] => readdirSync(join(root, dir), { withFileTypes: true })
    .flatMap((e) => (e.isDirectory() ? walk(join(dir, e.name)) : [join(dir, e.name)]));
  return [...walk('src/web'), ...walk('addin/vba'), ...walk('addin/icons')].map((p) => p.split(sep).join('/')).sort();
}

/** The repo-relative path of a source file, '/'-separated: what its import.meta.url stands for. */
const rel = (file: string) => relative(ROOT, file).split(sep).join('/');

/**
 * src/cli.ts uses top-level await, which CommonJS (the only format a SEA runs) has not: its body
 * after the imports runs in an async function instead.
 */
export function wrapTopLevelAwait(source: string): string {
  const lines = source.replace(/^#!.*\n/, '').split('\n');
  let i = 0;
  while (i < lines.length && /^(import\s|\/\/|\s*$)/.test(lines[i])) i++;
  return `${lines.slice(0, i).join('\n')}\n(async () => {\n${lines.slice(i).join('\n')}\n})().catch((err) => { console.error(err); process.exit(1); });\n`;
}

/** Each source file's import.meta.url, as its URL under the resource root. */
export const replaceImportMeta = (source: string, relPath: string) => source.replace(/\bimport\.meta\.url\b/g, `__ocUrl(${JSON.stringify(relPath)})`);

export function banner(assets: string[], assetHash: string): string {
  return `"use strict";
// open-charts engine bundle (scripts/build-win.ts). Resources (by import.meta.url) resolve under
// the repo for the plain bundle, and for the single executable under the folder its assets are
// unpacked into: %LOCALAPPDATA%\\open-charts\\res\\<hash>.
var __ocSea = (() => { try { return require('node:sea').isSea() ? require('node:sea') : null; } catch { return null; } })();
if (!__ocSea) globalThis.__OC_BUNDLE__ = __filename;
var __ocRootDir;
function __ocRoot() {
  if (__ocRootDir) return __ocRootDir;
  const path = require('node:path'), fs = require('node:fs');
  if (!__ocSea) return (__ocRootDir = path.resolve(__dirname, '..', '..'));
  const root = path.join(process.env.LOCALAPPDATA || require('node:os').tmpdir(), 'open-charts', 'res', ${JSON.stringify(assetHash)});
  if (!fs.existsSync(path.join(root, '.complete'))) {
    const tmp = root + '.' + process.pid + '.tmp';
    for (const key of ${JSON.stringify(assets)}) {
      const file = path.join(tmp, ...key.split('/'));
      fs.mkdirSync(path.dirname(file), { recursive: true });
      fs.writeFileSync(file, Buffer.from(__ocSea.getRawAsset(key)));
    }
    fs.writeFileSync(path.join(tmp, '.complete'), '');
    try { fs.renameSync(tmp, root); } catch { fs.rmSync(tmp, { recursive: true, force: true }); } // another process was first
  }
  return (__ocRootDir = root);
}
function __ocUrl(rel) { return require('node:url').pathToFileURL(require('node:path').join(__ocRoot(), ...rel.split('/'))).href; }
`;
}

async function bundle(assets: string[], assetHash: string) {
  const esbuild = await import('esbuild');
  await esbuild.build({
    entryPoints: [join(ROOT, 'src', 'cli.ts')],
    outfile: BUNDLE,
    bundle: true,
    platform: 'node',
    format: 'cjs',
    target: 'node22',
    banner: { js: banner(assets, assetHash) },
    logLevel: 'warning',
    plugins: [{
      name: 'open-charts-resources',
      setup(build) {
        build.onLoad({ filter: /\.(ts|js)$/ }, (a) => {
          const r = rel(a.path);
          if (r.startsWith('..') || r.includes('node_modules/')) return undefined;
          let source = readFileSync(a.path, 'utf8');
          if (r === 'src/cli.ts') source = wrapTopLevelAwait(source);
          return { contents: replaceImportMeta(source, r), loader: a.path.endsWith('.ts') ? 'ts' : 'js' };
        });
      },
    }],
  });
}

/** sha256 of a file, hex. */
const sha256 = (file: string) => createHash('sha256').update(readFileSync(file)).digest('hex');

/** The line for `name` in a SHASUMS256.txt ("<sha256>  win-x64/node.exe"). */
export function shasumFor(shasums: string, name: string): string | undefined {
  for (const line of shasums.split(/\r?\n/)) {
    const m = /^([0-9a-f]{64})\s+\*?(.+)$/.exec(line.trim());
    if (m && m[2] === name) return m[1];
  }
  return undefined;
}

/**
 * A URL's bytes: fetch, else curl (which follows the system's proxy settings where Node's fetch
 * does not, e.g. behind a corporate proxy or a sandbox that only lets curl out).
 */
async function download(url: string): Promise<Buffer> {
  try {
    const r = await fetch(url);
    if (!r.ok) throw new Error(`${url}: HTTP ${r.status}`);
    return Buffer.from(await r.arrayBuffer());
  } catch (err) {
    try {
      return execFileSync('curl', ['-fsSL', '--max-time', '600', url], { maxBuffer: 1 << 30 });
    } catch {
      throw err;
    }
  }
}

/** The official node.exe of this Node version, verified against nodejs.org's SHASUMS256.txt (cached). */
async function officialNodeExe(version = process.version): Promise<string> {
  const base = `https://nodejs.org/dist/${version}`;
  const cacheDir = join(process.env.OPEN_CHARTS_CACHE || join(process.env.XDG_CACHE_HOME || join(homedir(), '.cache'), 'open-charts'), 'node');
  const file = join(cacheDir, `node-${version}-win-x64.exe`);
  const sums = (await download(`${base}/SHASUMS256.txt`)).toString('utf8');
  const expected = shasumFor(sums, 'win-x64/node.exe');
  if (!expected) throw new Error(`no win-x64/node.exe in ${base}/SHASUMS256.txt`);
  if (existsSync(file) && sha256(file) === expected) return file;
  console.log(`Downloading ${base}/win-x64/node.exe`);
  const bytes = await download(`${base}/win-x64/node.exe`);
  const got = createHash('sha256').update(bytes).digest('hex');
  if (got !== expected) throw new Error(`node.exe checksum mismatch: got ${got}, SHASUMS256.txt says ${expected}`);
  mkdirSync(cacheDir, { recursive: true });
  writeFileSync(file, bytes);
  console.log(`node.exe ${version} verified (sha256 ${got})`);
  return file;
}

/**
 * Removes a PE file's Authenticode signature (the certificate table, data directory 4), which no
 * longer matches once the blob is injected: an unsigned exe is the honest state. Returns the bytes.
 */
export function stripSignature(pe: Buffer): Buffer {
  const lfanew = pe.readUInt32LE(0x3c);
  if (pe.readUInt32LE(lfanew) !== 0x00004550) throw new Error('not a PE file');
  const opt = lfanew + 24;
  const magic = pe.readUInt16LE(opt);
  const dirs = opt + (magic === 0x20b ? 112 : 96); // PE32+ : PE32
  const entry = dirs + 4 * 8;
  const offset = pe.readUInt32LE(entry), size = pe.readUInt32LE(entry + 4);
  if (!offset || !size) return pe;
  pe.writeUInt32LE(0, entry);
  pe.writeUInt32LE(0, entry + 4);
  return offset + size >= pe.length - 7 ? pe.subarray(0, offset) : pe; // the table is at the end (8-byte aligned)
}

async function seaBlob(main: string, blob: string, assets: string[]) {
  const config = join(dirname(blob), 'sea-config.json');
  writeFileSync(config, JSON.stringify({
    main, output: blob, disableExperimentalSEAWarning: true, useSnapshot: false, useCodeCache: false,
    assets: Object.fromEntries(assets.map((a) => [a, join(ROOT, a)])),
  }, null, 2));
  execFileSync(process.execPath, ['--experimental-sea-config', config], { stdio: 'inherit' });
}

async function inject(exe: string, blob: string, macho = false) {
  const { inject: postject } = await import('postject');
  await postject(exe, 'NODE_SEA_BLOB', readFileSync(blob), { sentinelFuse: SEA_FUSE, ...(macho ? { machoSegmentName: 'NODE_SEA' } : {}) });
}

async function main() {
  rmSync(OUT, { recursive: true, force: true });
  mkdirSync(OUT, { recursive: true });
  const assets = assetFiles();
  const assetHash = createHash('sha256').update(assets.map((a) => `${a}\0${sha256(join(ROOT, a))}`).join('\n')).digest('hex').slice(0, 16);
  await bundle(assets, assetHash);
  console.log(`Bundled ${rel(BUNDLE)} (${statSync(BUNDLE).size} bytes, ${assets.length} assets)`);

  const built = buildAddin();
  writeFileSync(join(OUT, 'open-charts.ppam'), built.ppam);
  writeFileSync(join(OUT, 'addin.json'), `${JSON.stringify({ prefix: built.prefix }, null, 2)}\n`);
  for (const f of ['addin/windows/install.ps1', 'addin/windows/uninstall.ps1', 'tools/e2e-powerpoint-win.ps1']) {
    copyFileSync(join(ROOT, f), join(OUT, f.split('/').at(-1)!));
  }
  console.log(`Wrote ${rel(OUT)}/open-charts.ppam (callback prefix ${built.prefix}), addin.json, install.ps1, uninstall.ps1, e2e-powerpoint-win.ps1`);
  if (args.has('--skip-exe')) return;

  const blob = join(OUT, 'sea-prep.blob');
  await seaBlob(BUNDLE, blob, assets);
  const exe = join(OUT, 'open-charts.exe');
  // Unsigned before injecting too: postject (LIEF) otherwise warns that the signature is corrupted.
  writeFileSync(exe, stripSignature(readFileSync(await officialNodeExe())));
  await inject(exe, blob);
  writeFileSync(exe, stripSignature(readFileSync(exe)));
  console.log(`Wrote ${rel(exe)} (${statSync(exe).size} bytes, Node ${process.version})`);

  if (args.has('--host') && process.platform !== 'win32') {
    const host = join(ROOT, 'dist', 'host', 'open-charts');
    mkdirSync(dirname(host), { recursive: true });
    copyFileSync(process.execPath, host);
    chmodSync(host, 0o755);
    if (process.platform === 'darwin') execFileSync('codesign', ['--remove-signature', host]);
    await inject(host, blob, process.platform === 'darwin');
    if (process.platform === 'darwin') execFileSync('codesign', ['--sign', '-', host]);
    console.log(`Wrote ${rel(host)} (this machine's SEA, for trying the exe's code paths)`);
  }
  rmSync(blob, { force: true });
  rmSync(join(OUT, 'sea-config.json'), { force: true });
}

if (/build-win\.ts$/.test(process.argv[1] ?? '')) await main();
