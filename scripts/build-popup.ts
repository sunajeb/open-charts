// `npm run build:popup [out.app]`: builds the Mac pop-up helper into dist/OpenChartsPopup.app.
import { buildPopup } from '../src/addin/build-popup.ts';

const t = Date.now();
const out = buildPopup(process.argv[2]);
console.log(`Built ${out} (${Date.now() - t} ms)`);
