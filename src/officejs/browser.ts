// The engine + writer as one browser module (dist/officejs/engine.js, ESM): everything the CLI lays
// out and writes, with no Node APIs. Used by the parity test (test/e2e/officejs.spec.ts) and
// available to other web hosts; the task pane bundles the same sources (taskpane.ts).
export * from '../core/index.ts';
export { samples } from '../core/samples.ts';
export { buildDeck, chartContent, altText, slideCharts, CHART_FRAME } from '../deck.ts';
export type { DeckSpec, DeckSlide, BuiltSlide } from '../deck.ts';
export { writePptx, groupXml } from '../render/pptx.ts';
export { sceneToSvg } from '../render/svg.ts';
export { chartTags, specJson, parseFrame } from '../render/tags.ts';
export { ENGINE_VERSION } from '../version.ts';
export { loadMeasurer, FONT_FILES } from './measure.ts';
export { buildChart, chartDeck, layoutFor, renderDeck, resolveFrame, splitSpec, toBase64, fromBase64 } from './engine.ts';
export { planScene, hostCaps } from './plan.ts';
export type { HostCaps, Op, Plan } from './plan.ts';
export { patchSlide, browserInflate } from './slide-patch.ts';
