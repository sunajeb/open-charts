// Text measurement in the browser: the bundled Liberation Sans (metric-compatible with Arial, SIL OFL,
// vendor/fonts) parsed by the same code the CLI uses, so scenes are byte-identical to Node's with the
// same font files (test/e2e/officejs.spec.ts). Charts in another font are still measured with these
// metrics; the pane says so (docs/plan/03 §3.2 — a canvas fallback is future work).
import { measurerFromFonts, type Measurer } from '../core/text.ts';

export const FONT_FILES = { regular: 'fonts/LiberationSans-Regular.ttf', bold: 'fonts/LiberationSans-Bold.ttf' };

/** Fetches both faces relative to `base` (the task pane's URL) and builds the measurer. */
export async function loadMeasurer(base: string | URL = (globalThis as { location?: Location }).location?.href ?? ''): Promise<Measurer> {
  const get = async (p: string) => {
    const r = await fetch(new URL(p, base));
    if (!r.ok) throw new Error(`Could not load ${p} (HTTP ${r.status})`);
    return r.arrayBuffer();
  };
  const [regular, bold] = await Promise.all([get(FONT_FILES.regular), get(FONT_FILES.bold)]);
  return measurerFromFonts(regular, bold);
}
