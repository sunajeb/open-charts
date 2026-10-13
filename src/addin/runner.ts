// Runs the add-in's VBA macros from the add-in server (Apply) through one long-lived osascript.
// Starting osascript costs ~0.1 s per call, and seconds while the Mac is busy (the editor's browser
// launching as the user opens it); a warm helper only sends the Apple event (~40 ms). The helper is
// JavaScript for Automation reading one JSON request per line on stdin and answering one per line:
//   {"macro": "<name>", "args": ["…"]}  ->  {"out": "<the macro's return value>"} | {"error": "…"}
//   {"ping": true}                      ->  {"out": "pong"}   (no Apple event)
//   {"connect": true}                   ->  {"out": "connected"}: one read-only Apple event (get name),
//                                           so the first Apply does not pay for setting up the
//                                           connection (seconds while the Mac is busy)
// It never starts PowerPoint (a closed PowerPoint has no chart to update) and exits when stdin
// closes, so it never outlives the server.
import { spawn, type ChildProcess } from 'node:child_process';

export const HELPER_JXA = `ObjC.import('Foundation');
function run() {
  const input = $.NSFileHandle.fileHandleWithStandardInput, output = $.NSFileHandle.fileHandleWithStandardOutput;
  const ppt = Application('Microsoft PowerPoint');
  const reply = (r) => output.writeData($.NSString.alloc.initWithUTF8String(JSON.stringify(r) + '\\n').dataUsingEncoding($.NSUTF8StringEncoding));
  let buf = '';
  for (;;) {
    const data = input.availableData;
    if (Number(data.length) === 0) return; // stdin closed (NSData length comes back as a string)
    buf += $.NSString.alloc.initWithDataEncoding(data, $.NSUTF8StringEncoding).js;
    for (let nl = buf.indexOf('\\n'); nl >= 0; nl = buf.indexOf('\\n')) {
      const line = buf.slice(0, nl);
      buf = buf.slice(nl + 1);
      try {
        const req = JSON.parse(line);
        if (req.ping) reply({ out: 'pong' });
        else if (req.connect) reply({ out: ppt.running() ? (ppt.name(), 'connected') : 'not running' });
        else if (!ppt.running()) reply({ error: 'execution error: PowerPoint is not running. (-600)' });
        else reply({ out: String(ppt.runVBMacro({ macroName: req.macro, listOfParameters: req.args })) });
      } catch (e) {
        reply({ error: 'execution error: ' + (e.message || String(e)) + ' (' + e.errorNumber + ')' });
      }
    }
  }
}`;

/** What a macro run gives: its return value, or an error (killed: it timed out and was stopped). */
export type MacroResult = { out: string } | { error: { message: string; killed?: boolean } };

export interface MacroRunner {
  run(macro: string, args: string[], timeoutMs: number): Promise<MacroResult>;
  ping(timeoutMs: number): Promise<boolean>;
  connect(timeoutMs: number): Promise<MacroResult>;
  stop(): void;
}

/**
 * One helper process, started on first use (or `warm()`), requests queued one at a time. A request
 * that times out (e.g. macOS is still asking whether PowerPoint may be controlled) stops the helper;
 * the next one starts a fresh helper.
 */
export function macroRunner(command = 'osascript', args = ['-l', 'JavaScript', '-e', HELPER_JXA]): MacroRunner & { warm(): void } {
  let child: ChildProcess | undefined;
  let buffer = '';
  let waiting: ((line: string | undefined) => void) | undefined;
  let queue: Promise<unknown> = Promise.resolve();

  const start = () => {
    if (child && child.exitCode === null && !child.killed) return child;
    buffer = '';
    const c = spawn(command, args, { stdio: ['pipe', 'pipe', 'ignore'], windowsHide: true });
    c.stdout!.setEncoding('utf8');
    c.stdout!.on('data', (d: string) => {
      buffer += d;
      for (let nl = buffer.indexOf('\n'); nl >= 0 && waiting; nl = buffer.indexOf('\n')) {
        const line = buffer.slice(0, nl);
        buffer = buffer.slice(nl + 1);
        const w = waiting;
        waiting = undefined;
        w(line);
      }
    });
    // Only this helper's own request fails with it: a helper stopped after a timeout may exit later,
    // while its successor is already answering the next request.
    const gone = () => {
      if (child !== c) return;
      child = undefined;
      const w = waiting;
      waiting = undefined;
      w?.(undefined);
    };
    c.on('exit', gone);
    c.on('error', gone);
    c.stdin!.on('error', () => { /* the helper exited: reported by "exit" */ });
    c.unref();
    (c.stdout as unknown as { unref?: () => void }).unref?.();
    child = c;
    return c;
  };

  const send = (request: object, timeoutMs: number) => {
    const next = queue.then(() => new Promise<{ line?: string; killed?: boolean }>((resolve) => {
      const c = start();
      const timer = setTimeout(() => { waiting = undefined; c.kill('SIGKILL'); resolve({ killed: true }); }, timeoutMs);
      waiting = (line) => { clearTimeout(timer); resolve({ line }); };
      c.stdin!.write(`${JSON.stringify(request)}\n`);
    }));
    queue = next.catch(() => undefined);
    return next;
  };

  return {
    warm() { start(); },
    async ping(timeoutMs) {
      const r = await send({ ping: true }, timeoutMs);
      return !!r.line && JSON.parse(r.line).out === 'pong';
    },
    connect(timeoutMs) { return answer(send({ connect: true }, timeoutMs), 'connect', timeoutMs); },
    run(macro, macroArgs, timeoutMs) { return answer(send({ macro, args: macroArgs }, timeoutMs), macro, timeoutMs); },
    stop() { child?.stdin?.end(); child = undefined; },
  };
}

async function answer(request: Promise<{ line?: string; killed?: boolean }>, what: string, timeoutMs: number): Promise<MacroResult> {
  const r = await request;
  if (r.killed) return { error: { message: `run VB macro ${what} timed out after ${timeoutMs} ms`, killed: true } };
  if (r.line === undefined) return { error: { message: 'execution error: the AppleScript helper stopped. (-1)' } };
  try {
    const reply = JSON.parse(r.line) as { out?: string; error?: string };
    return reply.error !== undefined ? { error: { message: reply.error } } : { out: reply.out ?? '' };
  } catch {
    return { error: { message: `execution error: unexpected answer from the AppleScript helper: ${r.line.slice(0, 200)} (-1)` } };
  }
}
