// Sideloads the Office.js add-in into PowerPoint for Mac (npm run officejs:sideload): copies
// dist/officejs/manifest.xml into PowerPoint's developer add-in folder. Restart PowerPoint, then
// Home ▸ open-charts (or Insert ▸ Add-ins ▸ My Add-ins ▸ Developer Add-ins). --remove takes it out.
import { copyFileSync, existsSync, mkdirSync, rmSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

export const WEF = join(homedir(), 'Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef');
export const SIDELOADED = join(WEF, 'open-charts-officejs.manifest.xml');
const manifest = fileURLToPath(new URL('../dist/officejs/manifest.xml', import.meta.url));

if (process.argv.includes('--remove')) {
  rmSync(SIDELOADED, { force: true });
  console.log(`Removed ${SIDELOADED}. Restart PowerPoint.`);
} else {
  if (!existsSync(manifest)) { console.error('No manifest. Run: npm run build:officejs'); process.exit(1); }
  mkdirSync(WEF, { recursive: true });
  copyFileSync(manifest, SIDELOADED);
  console.log(`Sideloaded: ${SIDELOADED}
Next: npm run officejs:serve (keep it running), restart PowerPoint, then Home ▸ open-charts.`);
}
