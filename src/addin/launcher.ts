// The `open-charts` command. `addin install` writes a small launcher into the add-in's home and links
// it into a folder on PATH when it can; every message that tells the user to run a command names one
// that works whether or not that link exists.
import { existsSync, lstatSync, mkdirSync, readlinkSync, rmSync, symlinkSync, accessSync, constants, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { delimiter, join } from 'node:path';

/** Where the add-in keeps its engine copy, settings, log and launcher. */
export function addinHome(): string {
  return process.env.OPEN_CHARTS_HOME ?? join(homedir(), 'Library', 'Application Support', 'open-charts');
}

export const launcherPath = (home = addinHome()) => join(home, 'bin', 'open-charts');

/** A shell word for a path: $HOME instead of the home folder, double-quoted (the path has spaces). */
const shellPath = (p: string) => `"${p.startsWith(homedir() + '/') ? `$HOME${p.slice(homedir().length)}` : p}"`;

/** The command to (re)install the add-in, as the user can paste it into Terminal. */
export const reinstallCommand = (home = addinHome()) => `${shellPath(launcherPath(home))} addin install`;

/**
 * The launcher: the Node that ran the install (or the one on PATH, if that one has gone) running the
 * open-charts folder the install was run from, so `open-charts addin install` picks up updates.
 */
export function launcherScript(node: string, cli: string): string {
  const q = (s: string) => `'${s.replace(/'/g, `'\\''`)}'`;
  return `#!/bin/sh
# open-charts launcher, written by \`open-charts addin install\`.
NODE=${q(node)}
CLI=${q(cli)}
[ -x "$NODE" ] || NODE=$(command -v node)
if [ ! -f "$CLI" ]; then
  echo "open-charts: $CLI is gone. Run 'node --experimental-strip-types --no-warnings src/cli.ts addin install' in the open-charts folder." >&2
  exit 1
fi
exec "$NODE" --experimental-strip-types --no-warnings "$CLI" "$@"
`;
}

/** Folders on PATH where a link to the launcher would make `open-charts` work, most usual first. */
function pathDirs(): string[] {
  const onPath = (process.env.PATH ?? '').split(delimiter).filter(Boolean);
  const preferred = [join(homedir(), '.local', 'bin'), join(homedir(), 'bin'), '/opt/homebrew/bin', '/usr/local/bin'];
  return preferred.filter((d) => onPath.includes(d));
}

const writable = (d: string) => { try { accessSync(d, constants.W_OK); return true; } catch { return false; } };
const ours = (link: string, target: string) => { try { return lstatSync(link).isSymbolicLink() && readlinkSync(link) === target; } catch { return false; } };

/**
 * Writes the launcher and links it into a folder on PATH if one is writable and has no other
 * `open-charts`. Returns the lines to print.
 */
export function installLauncher(node: string, cli: string, home = addinHome()): string[] {
  const bin = launcherPath(home);
  mkdirSync(join(home, 'bin'), { recursive: true });
  writeFileSync(bin, launcherScript(node, cli), { mode: 0o755 });
  const lines = [`Launcher written to ${bin}`];
  for (const dir of pathDirs()) {
    const link = join(dir, 'open-charts');
    if (ours(link, bin)) return [...lines, `The open-charts command is on your PATH (${link}).`];
    if (existsSync(link) || !writable(dir)) continue;
    try { symlinkSync(bin, link); } catch { continue; }
    return [...lines, `Linked ${link}: the open-charts command now works in new Terminal windows.`];
  }
  return [
    ...lines,
    'To type `open-charts` in Terminal, add its folder to your PATH, e.g.:',
    `  echo 'export PATH="${shellPath(join(home, 'bin')).slice(1, -1)}:$PATH"' >> ~/.zshrc`,
    `Until then, run it as: ${shellPath(bin)} …`,
  ];
}

/** Removes the launcher and any link to it on PATH. */
export function removeLauncher(home = addinHome()) {
  const bin = launcherPath(home);
  for (const dir of pathDirs()) {
    const link = join(dir, 'open-charts');
    if (ours(link, bin)) rmSync(link, { force: true });
  }
  rmSync(join(home, 'bin'), { recursive: true, force: true });
}
