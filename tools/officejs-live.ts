// Live check of the Office.js add-in in real PowerPoint for Mac (macOS only, not in CI):
//   npm run officejs:live
// Office.js cannot be driven from outside, so the task pane tests itself: this tool builds the
// add-in, serves it on https://localhost:4175 with the self-test switched on, sideloads the
// manifest, opens a new presentation and waits for the pane (Home ▸ open-charts — the one click a
// person has to make, unless the deck auto-opens it). The pane then inserts a chart (Path B),
// edits it in place, reverts, inserts a freeform chart (Path A) and exports every slide back to
// this tool, which reads the real OOXML with the same reader as `open-charts inspect`.
// Shares PowerPoint with other live tools through the lock /private/tmp/open-charts-ppt.lock.
// Owner prerequisites (once): npm run officejs:cert and trust the certificate (see its output).
import { execFileSync } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, rmdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildOfficeJs } from '../scripts/build-officejs.ts';
import { certDir, devTls, logFile, OFFICEJS_PORT, serveOfficeJs } from '../src/officejs/server.ts';
import { readCharts } from '../src/ooxml/pptx-read.ts';
import { samples } from '../src/core/samples.ts';

const ROOT = fileURLToPath(new URL('../', import.meta.url));
const OUT = join(ROOT, 'out', 'officejs-live');
const LOCK = '/private/tmp/open-charts-ppt.lock';
const WEF = join(process.env.HOME!, 'Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef');
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));
const osa = (s: string, t = 60) => execFileSync('osascript', ['-e', `with timeout of ${t} seconds\n${s}\nend timeout`], { encoding: 'utf8' }).trim();
const ppt = (s: string, t?: number) => osa(`tell application "Microsoft PowerPoint"\n${s}\nend tell`, t);
const running = () => osa('application "Microsoft PowerPoint" is running') === 'true';

function ownerStep(text: string): never {
  console.error(`\nOWNER STEP NEEDED:\n${text}\n`);
  process.exit(2);
}

if (process.platform !== 'darwin') { console.log('officejs:live runs on macOS with PowerPoint; skipped.'); process.exit(0); }

// 1. Certificate: must exist and be trusted, or PowerPoint's web view refuses the pane.
const tls = devTls();
if (!tls) ownerStep('Run: npm run officejs:cert   — then trust the certificate as it prints.');
const crt = join(certDir(), 'localhost.crt');
// The user's trust settings list the certificate by its common name once it is trusted.
let trusted = false;
try { trusted = execFileSync('security', ['dump-trust-settings'], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] }).includes('open-charts localhost'); } catch { /* none at all */ }
if (!trusted && !process.argv.includes('--skip-trust-check')) {
  ownerStep(`The dev certificate is not trusted yet. In Terminal run:\n  security add-trusted-cert -r trustRoot -k ~/Library/Keychains/login.keychain-db "${crt}"`);
}

// 2. Build, serve with the self-test on, sideload.
mkdirSync(OUT, { recursive: true });
const dist = await buildOfficeJs({ quiet: true });
const slides: string[] = [];
const paneLog: string[] = [];
const server = serveOfficeJs({ root: dist, port: OFFICEJS_PORT, tls, log: logFile(), selftest: true, onSlide: (b64) => slides.push(b64), onLog: (l) => paneLog.push(l) });
const lines = async () => paneLog;
// PowerPoint (and its add-in folder), one at a time.
// The owner can run this while holding the lock themselves: OPEN_CHARTS_LOCK=held npm run officejs:live
const ownsLock = process.env.OPEN_CHARTS_LOCK === 'held';
for (let waited = 0; !ownsLock; waited += 30) {
  try { mkdirSync(LOCK); break; } catch { /* held */ }
  if (waited % 300 === 0) console.log(`Waiting for ${LOCK} (another live run or the owner is using PowerPoint)…`);
  await sleep(30_000);
}
try {
  mkdirSync(WEF, { recursive: true });
  copyFileSync(join(dist, 'manifest.xml'), join(WEF, 'open-charts-officejs.manifest.xml'));
} catch (e) {
  server.close();
  if (!ownsLock) try { rmdirSync(LOCK); } catch { /* */ }
  ownerStep(`macOS did not let this process write PowerPoint's add-in folder (${(e as Error).message}).
Copy the manifest yourself:  cp "${join(dist, 'manifest.xml')}" "${WEF}/open-charts-officejs.manifest.xml"
(or run npm run officejs:sideload from Terminal and allow access to "data from other apps").`);
}

const wasRunning = running();
const failures: string[] = [];
try {
  ppt('activate\nmake new presentation');
  console.log('PowerPoint: new presentation open. Click Home ▸ open-charts to open the pane (the self-test then runs by itself).');
  let ready = '';
  for (const end = Date.now() + 5 * 60_000; Date.now() < end && !ready; await sleep(1000)) ready = (await lines()).find((l) => l.includes(' ready ')) ?? '';
  if (!ready) throw new Error('The task pane did not report "ready" within 5 minutes (not opened, or the certificate is not trusted).');
  console.log(ready);
  let done = '';
  for (const end = Date.now() + 3 * 60_000; Date.now() < end && !done; await sleep(1000)) done = (await lines()).find((l) => l.includes(' selftest-done ')) ?? '';
  const all = await lines();
  writeFileSync(join(OUT, 'pane.log'), all.join('\n') + '\n');
  for (const l of all.filter((x) => / (selftest-step|insert|apply|revert)[ -]/.test(x))) console.log(l);
  if (!done) failures.push('self-test did not finish within 3 minutes');
  for (const l of all.filter((x) => x.includes('"ok":false'))) failures.push(l);

  // 4. What is really on the slides (exported by PowerPoint itself), read like `open-charts inspect`.
  const charts = slides.flatMap((b64, i) => {
    const bytes = Uint8Array.from(Buffer.from(b64, 'base64'));
    writeFileSync(join(OUT, `slide${i + 1}.pptx`), bytes);
    return readCharts(bytes).flatMap((s) => s.charts);
  });
  console.log(`Charts on the slides: ${charts.map((c) => `${c.name} (${c.frame ? [c.frame.x, c.frame.y, c.frame.w, c.frame.h].join(',') : 'no frame'})`).join('; ') || 'none'}`);
  const clustered = charts.find((c) => c.name === 'Clustered column chart');
  if (!clustered) failures.push('no clustered column chart on the slides');
  else if (JSON.parse(clustered.spec ?? '{}').series?.[0]?.values?.[0] !== samples.clustered!.series[0].values[0]) failures.push('Revert did not restore the clustered chart\'s data');
  if (!charts.some((c) => c.name === 'Stacked area chart')) failures.push('no area chart (Path A) on the slides');
  // An independent look through AppleScript: the shapes PowerPoint reports on the current slide.
  try { console.log(`AppleScript, shapes on the current slide: ${ppt('get name of every shape of slide 1 of active presentation')}`); } catch { /* informational */ }
} catch (e) {
  failures.push((e as Error).message);
} finally {
  try { ppt('close active presentation saving no'); } catch { /* already closed */ }
  if (!wasRunning) { try { ppt('quit saving no'); } catch { /* gone */ } }
  if (!ownsLock) try { rmdirSync(LOCK); } catch { /* not ours any more */ }
  server.close();
}
writeFileSync(join(OUT, 'result.json'), JSON.stringify({ failures, slides: slides.length }, null, 2));
if (failures.length) { console.error(`\nFAILED:\n- ${failures.join('\n- ')}`); process.exit(1); }
console.log('\nOK: inserted (Path B), edited in place, reverted, inserted a freeform chart (Path A) in real PowerPoint.');
if (existsSync(join(OUT, 'pane.log'))) console.log(`Pane log: ${join(OUT, 'pane.log')}`);
