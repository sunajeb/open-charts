// Global setup: start the editor servers on free ports, stop them after the run.
import { spawn } from 'node:child_process';
import { randomBytes } from 'node:crypto';
import { createServer } from 'node:net';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../..', import.meta.url));

function freePort(): Promise<number> {
  return new Promise((resolve, reject) => {
    const s = createServer();
    s.unref();
    s.on('error', reject);
    s.listen(0, '127.0.0.1', () => {
      const { port } = s.address() as { port: number };
      s.close(() => resolve(port));
    });
  });
}

async function start(args: string[]): Promise<{ child: ReturnType<typeof spawn>; base: string }> {
  const port = await freePort();
  const child = spawn(process.execPath, ['--experimental-strip-types', '--no-warnings', ...args, String(port)], {
    cwd: root,
    stdio: ['ignore', 'pipe', 'inherit'],
  });
  await new Promise<void>((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`${args[0]} did not start within 15 s`)), 15_000);
    child.stdout!.on('data', (d: Buffer) => {
      if (d.toString().includes('http://')) { clearTimeout(timer); resolve(); }
    });
    child.on('exit', (code) => { clearTimeout(timer); reject(new Error(`${args[0]} exited with ${code}`)); });
  });
  child.stdout!.resume();
  return { child, base: `http://127.0.0.1:${port}` };
}

export default async function setup() {
  // The host is guarded like the add-in's server: a per-run token (src/addin/token.ts).
  process.env.E2E_TOKEN = randomBytes(32).toString('hex');
  // The plain editor (`open-charts serve`), and a host that confirms Apply like the add-in (host.ts).
  const editor = await start(['src/cli.ts', 'serve', '--port']);
  const host = await start(['test/e2e/host.ts']);
  process.env.E2E_BASE_URL = editor.base;
  process.env.E2E_HOST_URL = host.base;
  return async () => { editor.child.kill(); host.child.kill(); };
}
