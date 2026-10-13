// One PowerPoint session shared by the live scenarios: AppleScript calls with short timeouts (a
// timed-out call reads and dismisses any PowerPoint dialog and throws LiveTimeout), the add-in's
// macros, presentations the runner made (closed after each scenario), and leaving PowerPoint as
// found (quit only if the run started it).
import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { basename, join } from 'node:path';
import { addinHome } from '../../src/addin/server.ts';
import { readCharts, type SlideInfo } from '../../src/ooxml/pptx-read.ts';
import { LiveTimeout } from './core.ts';

export const APP = 'Microsoft PowerPoint';
const SERVER = 'http://127.0.0.1:4174';

/** Runs a program; throws on failure. Error objects carry `code`/`signal` as from execFileSync. */
export type Exec = (file: string, args: string[], timeoutMs: number) => string;
const realExec: Exec = (file, args, timeoutMs) =>
  execFileSync(file, args, { encoding: 'utf8', timeout: timeoutMs, stdio: ['ignore', 'pipe', 'pipe'] }).trim();

export const sleep = (ms: number) => Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
export const q = (s: string) => JSON.stringify(s);

/** An osascript failure that means "no answer in time" (AppleEvent timeout, or the process was killed). */
export function isTimeout(e: unknown): boolean {
  const err = e as { code?: string; signal?: string; message?: string; stderr?: string };
  return err.code === 'ETIMEDOUT' || err.signal === 'SIGTERM' || /-1712|timed out/i.test(`${err.message ?? ''} ${err.stderr ?? ''}`);
}

// Reads every PowerPoint dialog or sheet's text and presses its default button.
export const DIALOG_SCRIPT = `tell application "System Events" to tell process "${APP}"
set out to ""
repeat with w in windows
  try
    if (subrole of w is "AXDialog") or (subrole of w is "AXSystemDialog") or (exists sheet 1 of w) then
      try
        set out to out & (value of static texts of w as text) & " | "
      end try
      try
        set out to out & (value of static texts of sheet 1 of w as text) & " | "
      end try
      keystroke return
    end if
  end try
end repeat
return out
end tell`;

export interface PowerPointOptions {
  exec?: Exec;
  /** Default seconds for one AppleScript call. */
  callSec?: number;
  log?: string;
}

export class PowerPoint {
  readonly exec: Exec;
  readonly callSec: number;
  readonly logFile: string;
  /** PowerPoint was running before the run: never quit it (true until the runner has looked). */
  wasRunning = true;
  /** Presentations the runner made or opened (names), closed after each scenario. */
  owned: string[] = [];
  /** End of the current scenario's budget (ms since epoch); every call is capped by it. */
  deadline = Infinity;
  /** The add-in answered a macro since PowerPoint (re)started. */
  ready = false;
  prefix = '';
  slides = 0;
  /** Pause between macro attempts while the add-in compiles. */
  retryMs = 3000;

  constructor(o: PowerPointOptions = {}) {
    this.exec = o.exec ?? realExec;
    this.callSec = o.callSec ?? 30;
    this.logFile = o.log ?? join(addinHome(), 'addin.log');
  }

  running(): boolean {
    try { this.exec('pgrep', ['-x', APP], 5000); return true; } catch { return false; }
  }

  /** Seconds left for one call: its own limit, capped by the scenario's budget. */
  limit(sec: number): number {
    const left = (this.deadline - Date.now()) / 1000;
    if (left <= 1) throw new LiveTimeout('the scenario used up its time budget');
    return Math.max(1, Math.min(sec, Math.floor(left)));
  }

  /** One AppleScript; a call that does not answer in time reads + dismisses PowerPoint's dialog and throws. */
  osa(script: string, sec = this.callSec, what = script.split('\n').find((l) => l.trim() && !/^tell /.test(l))?.trim().slice(0, 80) ?? 'AppleScript'): string {
    const t = this.limit(sec);
    try {
      return this.exec('osascript', ['-e', `with timeout of ${t} seconds\n${script}\nend timeout`], (t + 5) * 1000);
    } catch (e) {
      if (!isTimeout(e)) throw e;
      throw new LiveTimeout(`no answer from PowerPoint within ${t} s: ${what}`, this.dismissDialog());
    }
  }

  tell(body: string, sec?: number): string {
    return this.osa(`tell application "${APP}"\n${body}\nend tell`, sec, body.split('\n')[0].slice(0, 80));
  }

  /** Text of PowerPoint's dialogs (pressing each one's default button); '' when there is none. */
  dismissDialog(): string {
    try { return this.exec('osascript', ['-e', DIALOG_SCRIPT], 15_000).replace(/( \| )+$/, ''); } catch (e) { return `(could not read dialogs: ${String((e as Error).message).split('\n')[0]})`; }
  }

  logLines(): string[] {
    return existsSync(this.logFile) ? readFileSync(this.logFile, 'utf8').split('\n').filter(Boolean) : [];
  }

  waitForLog(re: RegExp, from: number, ms = 20_000): string {
    for (const end = Math.min(Date.now() + ms, this.deadline); Date.now() < end; sleep(100)) {
      const hit = this.logLines().slice(from).find((l) => re.test(l));
      if (hit) return hit;
    }
    return '';
  }

  /**
   * Runs one of the add-in's macros; returns the add-in log lines it wrote. Right after PowerPoint
   * (re)started the add-in's VBA is still compiling and calls fail (-18): retried for up to 45 s.
   */
  macro(proc: string, ...args: string[]): string {
    if (!this.prefix) throw new Error('the add-in macro prefix is not known');
    const before = this.logLines().length;
    const script = `run VB macro macro name ${q(`${this.prefix}_${proc}`)} list of parameters {${args.map(q).join(', ')}}`;
    const sec = proc === 'ExportCopy' ? 90 : this.ready ? this.callSec : 60;
    for (const end = Date.now() + 45_000; ;) {
      try { this.tell(script, sec); break; } catch (e) {
        if (e instanceof LiveTimeout || this.ready || Date.now() > end) throw e;
        const dialog = this.dismissDialog();
        if (dialog) throw new LiveTimeout(`${proc}: PowerPoint showed a dialog while the add-in loaded`, dialog);
        sleep(this.retryMs);
      }
    }
    this.ready = true;
    const out = this.logLines().slice(before).join('\n');
    if (/still updating a chart/.test(out)) { sleep(500); return this.macro(proc, ...args); }
    return out;
  }

  /** A ribbon command through the add-in's own Dispatch; returns its RESULT line and the log. */
  run(id: string): { result: string; log: string; ms: number } {
    const t = Date.now();
    const log = this.macro('Run', id);
    return { result: log.split('\n').find((l) => l.includes(`RESULT ${id}`)) ?? '', log, ms: Date.now() - t };
  }

  /** The add-in is loaded (checked once per PowerPoint start); a clear failure instead of 45 s of retries. */
  loaded = false;
  ensureLoaded() {
    if (this.loaded) return;
    if (this.tell('get loaded of add in "open-charts"') !== 'true') throw new Error('open-charts is not loaded in PowerPoint: run with --install');
    this.loaded = true;
  }

  /** Starts PowerPoint if needed and makes a new presentation (owned by the runner). */
  newPresentation(): string {
    let name = '';
    try { name = this.tell('activate\nmake new presentation\nget name of active presentation', 60); } catch (e) {
      if (e instanceof LiveTimeout || !/-60[09]/.test(String((e as Error).message))) throw e;
      sleep(3000); // PowerPoint still quitting or starting
      name = this.tell('activate\nmake new presentation\nget name of active presentation', 60);
    }
    this.owned.push(name);
    this.slides = 0;
    return name;
  }

  /** Opens a deck (from a folder PowerPoint's sandbox may read); owned by the runner. */
  open(file: string): string {
    this.tell(`activate\nopen POSIX file ${q(file)}`, 60);
    const name = this.tell('get name of active presentation');
    this.owned.push(name || basename(file));
    return name;
  }

  newSlide(layout = 'blank'): number {
    this.slides++;
    this.tell(`make new slide at end of active presentation with properties {layout:slide layout ${layout}}\ngo to slide view of document window 1 number ${this.slides}`);
    return this.slides;
  }

  goTo(n: number) { this.tell(`go to slide view of document window 1 number ${n}`); }

  /** Closes the presentations this runner made; leaves everything else open. */
  closeOwned() {
    const names = this.owned.splice(0).reverse();
    if (!names.length || !this.running()) return;
    for (const n of names) {
      try { this.tell(`if exists presentation ${q(n)} then close presentation ${q(n)} saving no`, 20); } catch { /* gone, or PowerPoint is stuck: reported by the scenario */ }
    }
  }

  /** Quits PowerPoint, only if this run started it. Waits until the process is gone. */
  quit(): boolean {
    if (this.wasRunning || !this.running()) return false;
    try { this.exec('osascript', ['-e', `tell application "${APP}" to quit saving no`], 30_000); } catch { /* quitting */ }
    for (const end = Date.now() + 20_000; Date.now() < end && this.running(); sleep(250));
    this.ready = false;
    this.loaded = false;
    this.owned = [];
    return true;
  }

  /** Names of open presentations; [] when PowerPoint is not running (never starts it). */
  openDocuments(): string[] {
    if (!this.running()) return [];
    const out = this.exec('osascript', ['-e', `tell application "${APP}" to get name of every presentation`], 20_000);
    return out ? out.split(', ').filter(Boolean) : [];
  }

  /** Keystrokes for PowerPoint only: refused when another app is frontmost (someone may be using the Mac). */
  key(what: string) {
    this.osa(`tell application "${APP}" to activate`, 10);
    sleep(200);
    if (this.osa('tell application "System Events" to get name of first process whose frontmost is true', 10) !== APP) throw new Error('PowerPoint is not the frontmost app');
    this.osa(`tell application "System Events" to ${what}`, 10);
  }

  /** Saves a copy of the active presentation (PowerPoint's sandbox: the add-in's bridge copies it out). */
  exportCopy(dest: string, kind: 'pptx' | 'png'): string { return this.macro('ExportCopy', dest, kind); }

  /** Exports the active presentation to `file` and reads its charts like `open-charts inspect`. */
  charts(file: string): SlideInfo[] {
    this.exportCopy(file, 'pptx');
    if (!existsSync(file)) throw new Error(`the deck was not exported to ${file}`);
    return readCharts(new Uint8Array(readFileSync(file)));
  }

  // ---- the add-in's local server (Edit data sessions) ----

  async session(id: string): Promise<{ state: string; rev: number; message?: string; chart?: any } | undefined> {
    return fetch(`${SERVER}/api/session/${id}`, { signal: AbortSignal.timeout(5000) }).then((r) => (r.ok ? r.json() : undefined), () => undefined);
  }

  async post(id: string, body: unknown): Promise<{ rev: number }> {
    // A pooled keep-alive socket may have been closed by the server meanwhile: retry once.
    const send = () => fetch(`${SERVER}/api/session/${id}`, { method: 'POST', body: JSON.stringify(body), signal: AbortSignal.timeout(5000) });
    return (await (await send().catch(send)).json()) as { rev: number };
  }

  /** Apply in the editor: POSTs the chart as the editor does and waits for the server's outcome. */
  async apply(id: string, chart: unknown, waitMs = 30_000) {
    const before = this.logLines().length, t = Date.now();
    const { rev } = await this.post(id, { chart, state: 'applied' });
    let outcome: Awaited<ReturnType<PowerPoint['session']>> = undefined;
    for (const end = Math.min(Date.now() + waitMs, this.deadline); Date.now() < end; sleep(100)) {
      outcome = await this.session(id);
      if (outcome && outcome.rev > rev && ['drawn', 'error'].includes(outcome.state)) break;
    }
    const ms = Date.now() - t;
    const applied = this.waitForLog(/APPLIED|APPLY .* (failed|skipped)/, before, 1000);
    return { outcome, applied, ms, log: this.logLines().slice(before).join('\n') };
  }
}

/** The installed add-in's macro prefix (from addin.json), or '' when it is not installed. */
export function installedPrefix(home = addinHome()): string {
  try { return JSON.parse(readFileSync(join(home, 'addin.json'), 'utf8')).prefix ?? ''; } catch { return ''; }
}
