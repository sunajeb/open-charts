// Reads open-charts objects back out of any .pptx: finds groups carrying tag contract v1, their
// frames, roles and specs. Powers `open-charts inspect` and `open-charts import`.
import type { ChartSpec, Frame, Style } from '../core/types.ts';
import type { DeckSpec } from '../deck.ts';
import { TAG_CONTRACT_VERSION, parseFrame, specJson, upper } from '../render/tags.ts';
import { TITLE_NAME } from '../render/pptx.ts';
import { unzip, type ZipEntry } from './unzip.ts';
import { child, findAll, kids, parseXml, path, textOf, type XmlNode } from './xml.ts';

export type Entries = Map<string, ZipEntry>;
export interface Rel { id: string; type: string; target: string; external: boolean }

const EMU = 12700;

/** "ppt/slides/slide1.xml" → "ppt/slides/_rels/slide1.xml.rels" */
export function relsPath(part: string): string {
  const i = part.lastIndexOf('/');
  return `${part.slice(0, i + 1)}_rels/${part.slice(i + 1)}.rels`;
}

/** Resolves a relationship target relative to the part that owns it. */
export function resolveTarget(fromPart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const out = fromPart.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') out.pop();
    else if (seg !== '.' && seg !== '') out.push(seg);
  }
  return out.join('/');
}

export function readRels(entries: Entries, part: string): Map<string, Rel> {
  const out = new Map<string, Rel>();
  const e = entries.get(relsPath(part));
  if (!e) return out;
  for (const r of findAll(parseXml(e.text()), 'Relationship')) {
    const external = r.attrs.TargetMode === 'External';
    out.set(r.attrs.Id, { id: r.attrs.Id, type: r.attrs.Type, target: external ? r.attrs.Target : resolveTarget(part, r.attrs.Target), external });
  }
  return out;
}

/** Slide part names in presentation order. */
export function slideParts(entries: Entries): string[] {
  const pres = 'ppt/presentation.xml';
  const e = entries.get(pres);
  if (!e) throw new Error('Not a PowerPoint package (ppt/presentation.xml missing)');
  const rels = readRels(entries, pres);
  return findAll(parseXml(e.text()), 'p:sldId').map((s) => rels.get(s.attrs['r:id'])?.target).filter((t): t is string => !!t);
}

/** Tags attached to a shape via p:nvPr/p:custDataLst, as {NAME: value}. */
export function shapeTags(entries: Entries, rels: Map<string, Rel>, shape: XmlNode): Record<string, string> {
  const nv = shape.children.find((c) => c.name.startsWith('p:nv'));
  const out: Record<string, string> = {};
  for (const t of kids(path(nv, 'p:nvPr', 'p:custDataLst'), 'p:tags')) {
    const part = rels.get(t.attrs['r:id'])?.target;
    const e = part ? entries.get(part) : undefined;
    if (!e) continue;
    for (const tag of findAll(parseXml(e.text()), 'p:tag')) out[tag.attrs.name.toUpperCase()] = tag.attrs.val ?? '';
  }
  return out;
}

export const cNvPr = (shape: XmlNode) => shape.children.find((c) => c.name.startsWith('p:nv'))?.children.find((c) => c.name === 'p:cNvPr');

const SHAPES = new Set(['p:sp', 'p:grpSp', 'p:cxnSp', 'p:pic', 'p:graphicFrame']);
export const isShape = (n: XmlNode) => SHAPES.has(n.name);

/** Frame of a shape from its own a:xfrm, in points (undefined when inherited from a placeholder). */
export function xfrmFrame(shape: XmlNode): Frame | undefined {
  const pr = shape.children.find((c) => c.name === 'p:spPr' || c.name === 'p:grpSpPr');
  const x = child(pr, 'a:xfrm') ?? child(shape, 'p:xfrm');
  const off = child(x, 'a:off'), ext = child(x, 'a:ext');
  if (!off || !ext) return undefined;
  return { x: +off.attrs.x / EMU, y: +off.attrs.y / EMU, w: +ext.attrs.cx / EMU, h: +ext.attrs.cy / EMU };
}

/** Paragraph texts of a shape joined with "\n". */
export function shapeText(shape: XmlNode): string {
  const body = child(shape, 'p:txBody');
  return kids(body, 'a:p').map((p) => findAll(p, 'a:t').map(textOf).join('')).join('\n');
}

export interface ChartInfo {
  slide: number;
  name: string;
  /** OC_FRAME: the frame the layout was planned for. */
  frame?: Frame;
  /** Current group bounds; differs from frame when the user resized or moved the chart. */
  bounds?: Frame;
  contract: string;
  engine?: string;
  font?: string;
  spec?: string;
  specChunks: number;
  /** Bytes of all group-level tag values (UTF-8). */
  tagBytes: number;
  children: number;
  roles: string[];
}

export interface SlideInfo { index: number; part: string; title?: string; charts: ChartInfo[] }

export function readCharts(bytes: Uint8Array): SlideInfo[] {
  const entries = unzip(bytes);
  return slideParts(entries).map((part, i) => {
    const root = parseXml(entries.get(part)?.text() ?? '<p:sld/>');
    const rels = readRels(entries, part);
    const tree = find1(root, 'p:spTree');
    const charts: ChartInfo[] = [];
    let title: string | undefined;
    const visit = (n: XmlNode) => {
      for (const s of n.children.filter(isShape)) {
        const nm = cNvPr(s)?.attrs.name ?? '';
        if (n === tree && s.name === 'p:sp' && title === undefined && (nm === TITLE_NAME || isTitlePlaceholder(s))) title = shapeText(s);
        if (s.name !== 'p:grpSp') continue;
        const tags = shapeTags(entries, rels, s);
        if (tags.OC_V === undefined) { visit(s); continue; }
        const members = s.children.filter(isShape);
        const spec = specJson(tags);
        charts.push({
          slide: i + 1, name: nm, frame: parseFrame(tags.OC_FRAME), bounds: xfrmFrame(s),
          contract: tags.OC_V, engine: tags.OC_ENGINE, font: tags.OC_FONT, spec,
          specChunks: tags.OC_SPEC_N ? Number(tags.OC_SPEC_N) : spec === undefined ? 0 : 1,
          tagBytes: Object.values(tags).reduce((a, v) => a + Buffer.byteLength(v), 0),
          children: members.length,
          roles: members.map((m) => upper(shapeTags(entries, rels, m)).OC_ROLE).filter((r): r is string => !!r),
        });
      }
    };
    if (tree) visit(tree);
    return { index: i + 1, part, title, charts };
  });
}

function find1(n: XmlNode, name: string) { return findAll(n, name)[0]; }

function isTitlePlaceholder(sp: XmlNode): boolean {
  const ph = path(sp, 'p:nvSpPr', 'p:nvPr', 'p:ph');
  return ph?.attrs.type === 'title' || ph?.attrs.type === 'ctrTitle';
}

/** Recovers the DeckSpec from OC_SPEC tags. One chart per slide; extra charts become extra slides. */
export function importDeck(bytes: Uint8Array): { deck: DeckSpec; warnings: string[] } {
  const warnings: string[] = [];
  const slides: DeckSpec['slides'] = [];
  let style: Partial<Style> | undefined;
  for (const s of readCharts(bytes)) {
    s.charts.forEach((c, k) => {
      if (Number(c.contract) > Number(TAG_CONTRACT_VERSION)) warnings.push(`slide ${s.index}: "${c.name}" was made with a newer open-charts (tag contract ${c.contract}); importing as-is`);
      if (c.spec === undefined) { warnings.push(`slide ${s.index}: "${c.name}" has no OC_SPEC`); return; }
      const { style: st, ...chart } = JSON.parse(c.spec) as ChartSpec & { style?: Partial<Style> };
      if (st && !style) style = st;
      if (k > 0) warnings.push(`slide ${s.index}: chart ${k + 1} ("${c.name}") imported as its own slide`);
      slides.push({ ...(k === 0 && s.title !== undefined ? { title: s.title } : {}), chart });
    });
  }
  return { deck: { ...(style ? { style } : {}), slides }, warnings };
}
