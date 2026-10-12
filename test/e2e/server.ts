// Global setup: start the editor server on a free port, stop it after the run.
import { spawn } from 'node:child_process';
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

export default async function setup() {
  const port = await freePort();
  const child = spawn(process.execPath, ['--experimental-strip-types', '--no-warnings', 'src/cli.ts', 'serve', '--port', String(port)], {
    cwd: root,
    stdio: ['ignore', 'pipe', 'inherit'],
  });
  await new Promise<void>((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('editor server did not start within 15 s')), 15_000);
    child.stdout!.on('data', (d: Buffer) => {
      if (d.toString().includes('http://')) { clearTimeout(timer); resolve(); }
    });
    child.on('exit', (code) => { clearTimeout(timer); reject(new Error(`editor server exited with ${code}`)); });
  });
  child.stdout!.resume();
  process.env.E2E_BASE_URL = `http://127.0.0.1:${port}`;
  return async () => { child.kill(); };
}
