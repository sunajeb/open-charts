// `npm run build:popup`: builds the Mac pop-up helper (addin/popup-mac) into dist/OpenChartsPopup.app
// with swiftc from the Command Line Tools (no Xcode project), ad-hoc signed. `addin install` calls
// buildPopup() too and copies the app into the add-in's home.
import { execFileSync } from 'node:child_process';
import { copyFileSync, mkdirSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = fileURLToPath(new URL('../', import.meta.url));
const SRC = join(ROOT, 'addin', 'popup-mac');

/** True when swiftc can be run (Command Line Tools installed). */
export function canBuildPopup(): boolean {
  try { execFileSync('xcrun', ['--find', 'swiftc'], { stdio: 'ignore' }); return true; } catch { return false; }
}

/** Builds the helper app at `out` (default dist/OpenChartsPopup.app); returns its path. */
export function buildPopup(out = join(ROOT, 'dist', 'OpenChartsPopup.app')): string {
  if (process.platform !== 'darwin') throw new Error('The pop-up helper is built on a Mac.');
  rmSync(out, { recursive: true, force: true });
  const macos = join(out, 'Contents', 'MacOS');
  mkdirSync(macos, { recursive: true });
  copyFileSync(join(SRC, 'Info.plist'), join(out, 'Contents', 'Info.plist'));
  execFileSync('xcrun', ['swiftc', '-O', '-swift-version', '5', '-target', `${process.arch === 'arm64' ? 'arm64' : 'x86_64'}-apple-macos12.0`,
    '-framework', 'Cocoa', '-framework', 'WebKit', join(SRC, 'main.swift'), '-o', join(macos, 'open-charts-popup')], { stdio: ['ignore', 'inherit', 'inherit'] });
  // Ad-hoc signature over the whole bundle (Info.plist sealed): enough to run on this Mac; see
  // docs/USER-GUIDE.md "Editing data" for Gatekeeper.
  execFileSync('codesign', ['--force', '--sign', '-', '--timestamp=none', out], { stdio: ['ignore', 'inherit', 'inherit'] });
  return out;
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const t = Date.now();
  const out = buildPopup(process.argv[2]);
  console.log(`Built ${out} (${Date.now() - t} ms)`);
}
