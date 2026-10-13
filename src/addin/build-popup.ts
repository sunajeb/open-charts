// `npm run build:popup` (scripts/build-popup.ts): builds the Mac pop-up helper (addin/popup-mac) into dist/OpenChartsPopup.app
// with swiftc from the Command Line Tools (no Xcode project), ad-hoc signed. `addin install` calls
// buildPopup() too and copies the app into the add-in's home.
import { execFileSync } from 'node:child_process';
import { copyFileSync, mkdirSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = fileURLToPath(new URL('../../', import.meta.url));
const SRC = join(ROOT, 'addin', 'popup-mac');

/** True when swiftc can be run (Command Line Tools installed). */
export function canBuildPopup(): boolean {
  try { execFileSync('xcrun', ['--find', 'swiftc'], { stdio: 'ignore' }); return true; } catch { return false; }
}

/**
 * Builds the helper app at `out` (default dist/OpenChartsPopup.app); returns its path. `archs`
 * ['arm64', 'x64'] makes a universal app (the package build); `sign` false leaves it unsigned for
 * the caller to sign.
 */
export function buildPopup(out = join(ROOT, 'dist', 'OpenChartsPopup.app'), archs: string[] = [process.arch], sign = true): string {
  if (process.platform !== 'darwin') throw new Error('The pop-up helper is built on a Mac.');
  rmSync(out, { recursive: true, force: true });
  const macos = join(out, 'Contents', 'MacOS');
  mkdirSync(macos, { recursive: true });
  copyFileSync(join(SRC, 'Info.plist'), join(out, 'Contents', 'Info.plist'));
  const exe = join(macos, 'open-charts-popup');
  const slices = archs.map((arch) => {
    const target = arch === 'arm64' ? 'arm64' : 'x86_64';
    const file = archs.length === 1 ? exe : `${exe}-${target}`;
    execFileSync('xcrun', ['swiftc', '-O', '-swift-version', '5', '-target', `${target}-apple-macos12.0`,
      '-framework', 'Cocoa', '-framework', 'WebKit', join(SRC, 'main.swift'), '-o', file], { stdio: ['ignore', 'inherit', 'inherit'] });
    return file;
  });
  if (slices.length > 1) {
    execFileSync('lipo', ['-create', '-output', exe, ...slices]);
    for (const f of slices) rmSync(f);
  }
  // Ad-hoc signature over the whole bundle (Info.plist sealed): enough to run on this Mac; see
  // docs/USER-GUIDE.md "Editing data" for Gatekeeper.
  if (sign) execFileSync('codesign', ['--force', '--sign', '-', '--timestamp=none', out], { stdio: ['ignore', 'inherit', 'inherit'] });
  return out;
}
