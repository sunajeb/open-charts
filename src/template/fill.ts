// Template fill: `open-charts fill template.pptx data.json -o out.pptx`. Edits the user's own
// .pptx in place (docs/plan/04 §A.3):
//   - a shape named `chart:<name>` (Selection Pane), or whose whole text is `{{chart:<name>}}`, is
//     replaced by the chart `values[name]` laid out in that shape's frame (its own a:xfrm, else its
//     layout / master placeholder's), as a tagged open-charts group named `chart:<name>` — so the
//     add-ins can edit it and a second fill finds it again;
//   - `{{name}}` in any text (shapes, tables, notes) becomes `values[name]`, keeping the run's
//     formatting, even when PowerPoint split the field across runs;
//   - `slides` picks, orders and repeats (`each`) template slides;
//   - charts take the template's theme: series in accent slots are written as schemeClr and the
//     theme's colours drive label contrast; the theme's body font is the default chart font;
//   - every part we do not edit is written back byte-identical.
import type { ChartSpec, Frame, Scene, Style } from '../core/types.ts';
import type { Measurer } from '../core/text.ts';
import { checkScene, defaultStyle, layoutGroup, migrate } from '../core/index.ts';
import { THEME_SLOTS, themeOf, type ThemeColors } from '../core/theme.ts';
import { fileStyle, mergeStyles, styleFor, type StyleFile } from '../core/style-file.ts';
import { InvalidStyleFileError } from '../core/style-file.ts';
import { chartContent } from '../deck.ts';
import { groupXml, Ids, tagPartXml, TAGS_CT, TAGS_REL, type AddTags } from '../render/pptx.ts';
import { child, decodeEntities, escXml, findAll, parseXml, path, textOf, type XmlNode } from '../ooxml/xml.ts';
import { cNvPr, isShape } from '../ooxml/pptx-read.ts';
import { Pkg } from './package.ts';
import { isChartValue, validateFillData, type FillData, type FillValues } from './data.ts';

const EMU = 12700;
const PRES = 'ppt/presentation.xml';
const FIELD = /\{\{\s*([A-Za-z0-9_.-]{1,100})\s*\}\}/g;
const CHART_NAME = /^chart:([A-Za-z0-9_.-]{1,100})$/i;
const CHART_TEXT = /^\{\{\s*chart:([A-Za-z0-9_.-]{1,100})\s*\}\}$/i;

export interface FillOptions {
  measurer: Measurer;
  style?: StyleFile;
  /** Report only: plan and lay out everything, write nothing. */
  dryRun?: boolean;
}

export interface ChartSlotReport { name: string; frame?: Frame; status: 'filled' | 'missing' | 'error'; detail?: string; issues: string[] }
export interface SlideReport { index: number; source: number; charts: ChartSlotReport[]; fields: { name: string; status: 'filled' | 'missing' }[] }
export interface FillReport {
  template: { slides: number; theme: Partial<ThemeColors>; font?: string };
  slides: SlideReport[];
  /** Values no slot used. */
  unused: string[];
  warnings: string[];
  /** Missing values + chart errors + layout issues. */
  problems: number;
  ok: boolean;
  /** Parts written with new content, and parts removed (dropped slides and what only they used). */
  changedParts: string[];
  removedParts: string[];
}

export class InvalidFillDataError extends Error {}

export function fillTemplate(template: Uint8Array, data: FillData, opts: FillOptions): { pptx?: Uint8Array; report: FillReport } {
  const errors = validateFillData(data);
  if (errors.length) throw new InvalidFillDataError(`Invalid fill data:\n${errors.map((e) => `  ${e.path}: ${e.message}`).join('\n')}`);
  const pkg = new Pkg(template);
  const before = pkg.reachable();
  const presRels = pkg.rels(PRES);
  const presXml = pkg.text(PRES);
  const pres = parseXml(presXml);
  const sldIds = findAll(pres, 'p:sldId');
  const tplSlides = sldIds.map((s) => ({ id: Number(s.attrs.id), rid: s.attrs['r:id'], part: presRels.get(s.attrs['r:id'])?.target ?? '' }));
  if (!tplSlides.length) throw new Error('The template has no slides');

  const { theme, font: themeFont } = templateTheme(pkg);
  const warnings: string[] = [];
  const used = new Set<string>();
  const globals = data.values ?? {};

  // The output plan: which template slide, with which values, in order.
  const plan: { src: number; values: FillValues }[] = [];
  if (!data.slides) tplSlides.forEach((_, i) => plan.push({ src: i, values: globals }));
  else {
    for (const [k, s] of data.slides.entries()) {
      if (s.slide > tplSlides.length) throw new InvalidFillDataError(`/slides/${k}/slide: the template has ${tplSlides.length} slides, not ${s.slide}`);
      const base = { ...globals, ...s.values };
      for (const item of s.each ?? [{}]) plan.push({ src: s.slide - 1, values: { ...base, ...item } });
    }
  }

  // Slide parts: the first use of a template slide edits it in place, later uses edit a copy.
  const firstUse = new Set<number>();
  let nextId = Math.max(255, ...tplSlides.map((s) => s.id));
  const out = plan.map((p) => {
    const t = tplSlides[p.src];
    if (!firstUse.has(p.src)) { firstUse.add(p.src); return { ...p, part: t.part, id: t.id, rid: t.rid }; }
    const part = pkg.clone(t.part);
    const rid = pkg.addRel(PRES, 'slide', part);
    return { ...p, part, id: ++nextId, rid };
  });

  const slides: SlideReport[] = out.map((o, i) => fillSlide(pkg, o.part, o.values, { index: i + 1, source: o.src + 1 }, opts, theme, themeFont, data.style, used, warnings));

  // presentation.xml: the new slide list (and sections), only when it changed.
  const sameList = out.length === tplSlides.length && out.every((o, i) => o.id === tplSlides[i].id);
  if (!sameList) {
    const lst = findAll(pres, 'p:sldIdLst')[0];
    const body = out.map((o) => `<p:sldId id="${o.id}" r:id="${o.rid}"/>`).join('');
    let xml = presXml.slice(0, lst.innerStart) + body + presXml.slice(lst.innerEnd);
    xml = updateSections(xml, out.map((o) => ({ id: o.id, src: tplSlides[o.src].id })), warnings);
    pkg.set(PRES, xml);
    const keep = new Set(out.map((o) => o.rid));
    const dropped = tplSlides.filter((t) => !keep.has(t.rid)).map((t) => t.rid);
    if (dropped.length) pkg.setRels(PRES, [...pkg.rels(PRES).values()].filter((r) => !dropped.includes(r.id)));
    if (findAll(pres, 'p:custShow').length) warnings.push('the template has custom shows; slides they list may be missing or repeated');
  }
  pkg.collect(before);

  const allValues = new Set([...Object.keys(globals), ...(data.slides ?? []).flatMap((s) => [...Object.keys(s.values ?? {}), ...(s.each ?? []).flatMap(Object.keys)])]);
  const unused = [...allValues].filter((k) => !used.has(k)).sort();
  const problems = slides.reduce((a, s) => a + s.fields.filter((f) => f.status === 'missing').length + s.charts.reduce((b, c) => b + (c.status === 'filled' ? c.issues.length : 1), 0), 0);
  const pptx = opts.dryRun ? undefined : pkg.toBytes();
  return {
    pptx,
    report: {
      template: { slides: tplSlides.length, theme, font: themeFont },
      slides, unused, warnings, problems, ok: problems === 0,
      changedParts: pkg.changedParts().sort(), removedParts: pkg.removedParts().sort(),
    },
  };
}

// ---- One slide ---------------------------------------------------------------------------------

interface Edit { start: number; end: number; text: string }

function fillSlide(pkg: Pkg, part: string, values: FillValues, at: { index: number; source: number }, opts: FillOptions, theme: Partial<ThemeColors>, themeFont: string | undefined, dataStyle: Partial<Style> | undefined, used: Set<string>, warnings: string[]): SlideReport {
  const xml = pkg.text(part);
  const root = parseXml(xml);
  const tree = findAll(root, 'p:spTree')[0];
  const report: SlideReport = { ...at, charts: [], fields: [] };
  const edits: Edit[] = [];

  // 1. Chart slots: top-level shapes named chart:<name> or holding only {{chart:<name>}}.
  const slots: { name: string; shape: XmlNode; frame?: Frame }[] = [];
  const visit = (n: XmlNode, nested: boolean) => {
    for (const s of n.children.filter(isShape)) {
      const name = slotName(s);
      if (name) {
        if (nested) report.charts.push({ name, status: 'error', detail: 'inside a group: ungroup it or move it out of the group', issues: [] });
        else slots.push({ name, shape: s, frame: shapeFrame(pkg, part, s) });
      } else if (s.name === 'p:grpSp') visit(s, true);
    }
  };
  if (tree) visit(tree, false);

  const items: { slot: typeof slots[number]; spec: ChartSpec; frame: Frame }[] = [];
  for (const slot of slots) {
    const v = values[slot.name];
    if (v === undefined) { report.charts.push({ name: slot.name, frame: slot.frame, status: 'missing', detail: `no value "${slot.name}"`, issues: [] }); continue; }
    used.add(slot.name);
    if (!isChartValue(v)) { report.charts.push({ name: slot.name, frame: slot.frame, status: 'error', detail: `"${slot.name}" is text; a chart slot needs a chart spec`, issues: [] }); continue; }
    if (!slot.frame) { report.charts.push({ name: slot.name, status: 'error', detail: 'the shape has no size (and no placeholder to inherit one from)', issues: [] }); continue; }
    const chart = migrate(v);
    // Style: file (for this type) < data style < chart style; the template's theme always wins, so
    // contrast is computed on the colours PowerPoint will draw.
    const style = mergeStyles(themeFont ? { font: themeFont } : undefined, styleFor(opts.style, chart.type), dataStyle, chart.style, { theme });
    items.push({ slot, spec: { ...chart, style }, frame: slot.frame });
  }

  if (items.length) {
    const scenes: Scene[] = layoutGroup(items.map((it) => ({ spec: it.spec, frame: it.frame })), opts.measurer, defaultStyle);
    const ids = new Ids(maxShapeId(xml));
    const addTags: AddTags = (tags) => {
      const tag = pkg.freeName('ppt/tags/tag1.xml');
      pkg.set(tag, tagPartXml(tags), TAGS_CT);
      return pkg.addRel(part, TAGS_REL, tag);
    };
    items.forEach((it, k) => {
      const scene = scenes[k];
      const font = it.spec.style?.font ?? defaultStyle.font;
      const issues = [...checkScene(scene, { measurer: opts.measurer }).map((i) => `${i.kind}: ${i.detail}`), ...scene.hidden.map((h) => `hidden: ${h}`)];
      report.charts.push({ name: it.slot.name, frame: it.frame, status: 'filled', issues });
      if (opts.dryRun) return;
      const content = chartContent(it.spec, scene, font, undefined, `chart:${it.slot.name}`);
      edits.push({ start: it.slot.shape.start, end: it.slot.shape.end, text: groupXml(content, ids, font, addTags) });
    });
  }

  // 2. Text fields everywhere else on the slide.
  const skip = items.map((it) => [it.slot.shape.start, it.slot.shape.end]);
  edits.push(...fieldEdits(xml, root, values, report, used, (n) => skip.some(([a, b]) => n.start >= a && n.end <= b)));
  if (edits.length && !opts.dryRun) pkg.set(part, splice(xml, edits));

  // 3. Speaker notes.
  const notes = [...pkg.rels(part).values()].find((r) => r.type.endsWith('/notesSlide'));
  if (notes && pkg.has(notes.target)) {
    const nx = pkg.text(notes.target);
    const ne = fieldEdits(nx, parseXml(nx), values, report, used, () => false);
    if (ne.length && !opts.dryRun) pkg.set(notes.target, splice(nx, ne));
  }
  if (report.charts.some((c) => c.status === 'filled') && opts.measurer && themeFont && !/^(Arial|Liberation Sans|Arimo|Helvetica)$/i.test(themeFont) && !opts.style?.font && !dataStyle?.font) {
    const w = `charts use the template's font "${themeFont}", measured with Arial metrics unless you pass --font`;
    if (!warnings.includes(w)) warnings.push(w);
  }
  return report;
}

function slotName(s: XmlNode): string | undefined {
  const name = cNvPr(s)?.attrs.name ?? '';
  const m = CHART_NAME.exec(name.trim());
  if (m) return m[1];
  if (s.name !== 'p:sp') return undefined;
  const t = findAll(s, 'a:t').map(textOf).join('').trim();
  return CHART_TEXT.exec(t)?.[1];
}

function maxShapeId(xml: string): number {
  let max = 1;
  for (const m of xml.matchAll(/<p:cNvPr\b[^>]*?\bid="(\d+)"/g)) max = Math.max(max, Number(m[1]));
  return max;
}

/** `{{name}}` replacements in every paragraph, as edits on the source string (runs keep their a:rPr). */
function fieldEdits(xml: string, root: XmlNode, values: FillValues, report: SlideReport, used: Set<string>, skip: (n: XmlNode) => boolean): Edit[] {
  const edits: Edit[] = [];
  for (const p of findAll(root, 'a:p')) {
    if (skip(p)) continue;
    // Runs of consecutive a:r; line breaks and fields end a run group.
    let group: XmlNode[] = [];
    const flush = () => { if (group.length) edits.push(...groupEdits(xml, group, values, report, used)); group = []; };
    for (const c of p.children) {
      if (c.name === 'a:r' && child(c, 'a:t')) group.push(c);
      else flush();
    }
    flush();
  }
  return edits;
}

function groupEdits(xml: string, runs: XmlNode[], values: FillValues, report: SlideReport, used: Set<string>): Edit[] {
  const texts = runs.map((r) => textOf(child(r, 'a:t')!));
  const joined = texts.join('');
  if (!joined.includes('{{')) return [];
  const starts: number[] = [];
  texts.reduce((a, t) => { starts.push(a); return a + t.length; }, 0);
  const runAt = (pos: number) => { let k = 0; while (k + 1 < runs.length && starts[k + 1] <= pos) k++; return k; };
  const next = [...texts];
  let changed = false;
  // Right to left, so earlier offsets stay valid.
  const matches = [...joined.matchAll(FIELD)].reverse();
  for (const m of matches) {
    const name = m[1];
    if (/^chart:/i.test(name)) continue;
    const v = values[name];
    if (v === undefined || isChartValue(v)) { report.fields.push({ name, status: 'missing' }); continue; }
    used.add(name);
    report.fields.push({ name, status: 'filled' });
    const a = m.index!, b = a + m[0].length;
    const ra = runAt(a), rb = runAt(b - 1);
    const value = String(v);
    if (ra === rb) next[ra] = next[ra].slice(0, a - starts[ra]) + value + next[ra].slice(b - starts[ra]);
    else {
      next[ra] = next[ra].slice(0, a - starts[ra]) + value;
      for (let k = ra + 1; k < rb; k++) next[k] = '';
      next[rb] = next[rb].slice(b - starts[rb]);
    }
    changed = true;
  }
  report.fields.reverse();
  if (!changed) return [];
  const edits: Edit[] = [];
  runs.forEach((r, k) => {
    if (next[k] === texts[k]) return;
    const t = child(r, 'a:t')!;
    const rPr = child(r, 'a:rPr');
    const rPrXml = rPr ? xml.slice(rPr.start, rPr.end) : '';
    // A line break in a value becomes a:br with the run's formatting.
    const lines = next[k].split(/\r?\n/);
    const text = lines.map((l) => `<a:t>${escXml(l)}</a:t>`).join(`</a:r><a:br>${rPrXml}</a:br><a:r>${rPrXml}`);
    edits.push({ start: t.start, end: t.end, text });
  });
  return edits;
}

function splice(src: string, edits: Edit[]): string {
  const sorted = [...edits].sort((a, b) => b.start - a.start);
  let out = src;
  for (const e of sorted) out = out.slice(0, e.start) + e.text + out.slice(e.end);
  return out;
}

// ---- Frames and placeholders -------------------------------------------------------------------

function xfrmOf(shape: XmlNode): Frame | undefined {
  const pr = shape.children.find((c) => c.name === 'p:spPr' || c.name === 'p:grpSpPr');
  const x = child(pr, 'a:xfrm') ?? child(shape, 'p:xfrm');
  const off = child(x, 'a:off'), ext = child(x, 'a:ext');
  if (!off || !ext) return undefined;
  const f = { x: +off.attrs.x / EMU, y: +off.attrs.y / EMU, w: +ext.attrs.cx / EMU, h: +ext.attrs.cy / EMU };
  return [f.x, f.y, f.w, f.h].every(Number.isFinite) && f.w > 0 && f.h > 0 ? f : undefined;
}

/** A shape's frame: its own a:xfrm, else its placeholder's on the layout, else on the master. */
function shapeFrame(pkg: Pkg, slidePart: string, shape: XmlNode): Frame | undefined {
  const own = xfrmOf(shape);
  if (own) return own;
  const ph = path(shape.children.find((c) => c.name.startsWith('p:nv')), 'p:nvPr', 'p:ph');
  if (!ph) return undefined;
  const layout = [...pkg.rels(slidePart).values()].find((r) => r.type.endsWith('/slideLayout'))?.target;
  const master = layout && [...pkg.rels(layout).values()].find((r) => r.type.endsWith('/slideMaster'))?.target;
  for (const [p, byIdx] of [[layout, true], [master, false]] as const) {
    if (!p || !pkg.has(p)) continue;
    const match = findPlaceholder(parseXml(pkg.text(p)), ph.attrs.type ?? 'body', ph.attrs.idx, byIdx);
    const f = match && xfrmOf(match);
    if (f) return f;
  }
  return undefined;
}

function findPlaceholder(root: XmlNode, type: string, idx: string | undefined, byIdx: boolean): XmlNode | undefined {
  const shapes = findAll(root, 'p:sp');
  const phOf = (s: XmlNode) => path(s.children.find((c) => c.name.startsWith('p:nv')), 'p:nvPr', 'p:ph');
  const norm = (t: string | undefined) => (t === 'ctrTitle' ? 'title' : t === 'subTitle' || t === 'obj' || !t ? 'body' : t);
  if (byIdx && idx !== undefined) {
    const hit = shapes.find((s) => phOf(s)?.attrs.idx === idx);
    if (hit) return hit;
  }
  return shapes.find((s) => { const p = phOf(s); return p && norm(p.attrs.type) === norm(type); });
}

// ---- Theme -------------------------------------------------------------------------------------

/** Colours and body font of the template's (first master's) theme. */
export function templateTheme(pkg: Pkg): { theme: Partial<ThemeColors>; font?: string } {
  const master = [...pkg.rels(PRES).values()].find((r) => r.type.endsWith('/slideMaster'));
  const t = master && [...pkg.rels(master.target).values()].find((r) => r.type.endsWith('/theme'));
  const theme: Partial<ThemeColors> = {};
  if (!t || !pkg.has(t.target)) return { theme };
  const root = parseXml(pkg.text(t.target));
  const scheme = findAll(root, 'a:clrScheme')[0];
  for (const k of THEME_SLOTS) {
    const slot = scheme && child(scheme, `a:${k}`);
    const c = slot && (child(slot, 'a:srgbClr')?.attrs.val ?? child(slot, 'a:sysClr')?.attrs.lastClr);
    if (c && /^[0-9A-Fa-f]{6}$/.test(c)) theme[k] = c.toUpperCase();
  }
  const minor = findAll(root, 'a:minorFont')[0];
  const font = minor && child(minor, 'a:latin')?.attrs.typeface;
  return { theme, font: font ? decodeEntities(font) : undefined };
}

// ---- Sections ----------------------------------------------------------------------------------

/**
 * Keeps PowerPoint sections (p14:sectionLst) consistent with the new slide list: each output slide
 * joins its template slide's section. Sections that would no longer be contiguous are dropped
 * (with a warning) rather than written wrong.
 */
function updateSections(xml: string, out: { id: number; src: number }[], warnings: string[]): string {
  const root = parseXml(xml);
  const lst = findAll(root, 'p14:sectionLst')[0];
  if (!lst) return xml;
  const sections = findAll(lst, 'p14:section');
  const sectionOf = new Map<number, number>();
  sections.forEach((s, k) => findAll(s, 'p14:sldId').forEach((x) => sectionOf.set(Number(x.attrs.id), k)));
  const seq = out.map((o) => sectionOf.get(o.src) ?? -1);
  const contiguous = seq.every((k, i) => i === 0 || k === seq[i - 1] || !seq.slice(0, i).includes(k));
  if (!contiguous || seq.includes(-1)) {
    warnings.push('sections removed: the new slide order splits them');
    const ext = lst.parent!;
    return xml.slice(0, ext.start) + xml.slice(ext.end);
  }
  const edits: Edit[] = sections.map((s, k) => {
    const ids = findAll(s, 'p14:sldIdLst')[0];
    const body = out.filter((_, i) => seq[i] === k).map((o) => `<p14:sldId id="${o.id}"/>`).join('');
    return ids ? { start: ids.innerStart, end: ids.innerEnd, text: body } : { start: s.innerEnd, end: s.innerEnd, text: `<p14:sldIdLst>${body}</p14:sldIdLst>` };
  });
  return splice(xml, edits);
}

export { InvalidStyleFileError, themeOf, fileStyle };
