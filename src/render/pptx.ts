// Writes native PowerPoint (OOXML) slides. Each chart becomes one grouped object of plain shapes,
// so the deck opens and edits in any PowerPoint, Keynote or LibreOffice without this tool.
// Every group carries tag contract v1 (OC_* tags, see tags.ts) and each child its OC_ROLE, as real
// PowerPoint tags: p:custDataLst → ppt/tags/tagN.xml, visible to VBA Shape.Tags and Office.js.
import type { PathShape, Scene, Shape, TextShape } from '../core/types.ts';
import { flattenPath, pathBounds, rectGeom } from '../core/path.ts';
import { colorHex, colorSlot, DEFAULT_THEME, type ThemeColors } from '../core/theme.ts';
import type { Tags } from './tags.ts';
import { zip } from './zip.ts';

export interface SceneContent {
  scene: Scene;
  name: string;
  altText: string;
  /** Group-level tags, normally chartTags(...). */
  tags?: Tags;
}

export interface SlideContent {
  title?: string;
  scenes: SceneContent[];
}

/** Registers a tag part for one shape and returns its slide-relative relationship id. */
export type AddTags = (tags: Tags) => string;

export const SLIDE = { w: 960, h: 540 }; // points, 16:9
export const EMU = 12700;
const e = (pt: number) => Math.round(pt * EMU);

export const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
export const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
export const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const TAGS_REL = `${REL}/tags`;
export const TAGS_CT = 'application/vnd.openxmlformats-officedocument.presentationml.tags+xml';
/** Name of the slide title text box; `import` reads the title back from it. */
export const TITLE_NAME = 'Title';

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
/** A scene colour as DrawingML: a theme colour (`schemeClr`) when it came from a theme slot, else RGB. */
export const clrXml = (c: string) => {
  const slot = colorSlot(c);
  return slot ? `<a:schemeClr val="${slot}"/>` : `<a:srgbClr val="${colorHex(c)}"/>`;
};
const fill = (c?: string) => (c ? `<a:solidFill>${clrXml(c)}</a:solidFill>` : '<a:noFill/>');
const ln = (c?: string, w = 0.75, extra = '') =>
  c ? `<a:ln w="${e(w)}">${fill(c)}${extra}</a:ln>` : '<a:ln><a:noFill/></a:ln>';
const nvPr = (tagsRel?: string) =>
  (tagsRel ? `<p:nvPr><p:custDataLst><p:tags r:id="${tagsRel}"/></p:custDataLst></p:nvPr>` : '<p:nvPr/>');

export function tagPartXml(tags: Tags): string {
  return `${XML}<p:tagLst ${NS}>${tags.map(([k, v]) => `<p:tag name="${esc(k)}" val="${esc(v)}"/>`).join('')}</p:tagLst>`;
}

/** Shape id allocator; ids must be unique per slide, so start above the slide's highest id. */
export class Ids {
  n: number;
  constructor(start = 1) { this.n = start; }
  next() { return ++this.n; }
}

const ROLE_NAMES: Record<string, string> = {
  bar: 'Bar', marker: 'Marker', line: 'Line', label: 'Label', axis: 'Axis', gridline: 'Gridline', tick: 'Tick',
  arrow: 'Arrow', bubble: 'Bubble', connector: 'Connector', legend: 'Legend',
};

/** Selection-pane name: "Bar s0 c2", "Label value s1 c2"; "<Kind> <id>" when there is no role. */
export function shapeName(s: Shape, id: number, role?: string): string {
  if (role) {
    const [head, ...rest] = role.split(':');
    return [ROLE_NAMES[head] ?? head, ...rest].join(' ');
  }
  if (s.kind === 'rect') return `${rectGeom(s) === 'ellipse' ? 'Marker' : 'Bar'} ${id}`;
  if (s.kind === 'path') return `Shape ${id}`;
  return `${s.kind === 'line' ? 'Line' : 'Label'} ${id}`;
}

function shapeXml(s: Shape, ids: Ids, font: string, role?: string, addTags?: AddTags): string {
  const id = ids.next();
  const name = esc(shapeName(s, id, role));
  const nv = nvPr(role && addTags ? addTags([['OC_ROLE', role]]) : undefined);
  if (s.kind === 'rect') {
    return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/>${nv}</p:nvSpPr><p:spPr><a:xfrm><a:off x="${e(s.x)}" y="${e(s.y)}"/><a:ext cx="${e(s.w)}" cy="${e(s.h)}"/></a:xfrm><a:prstGeom prst="${rectGeom(s)}"><a:avLst/></a:prstGeom>${fill(s.fill)}${ln(s.stroke, s.strokeWidth)}</p:spPr></p:sp>`;
  }
  if (s.kind === 'path') return pathXml(s, id, name, nv);
  if (s.kind === 'line') {
    const x = Math.min(s.x1, s.x2), y = Math.min(s.y1, s.y2);
    const flip = `${s.x2 < s.x1 ? ' flipH="1"' : ''}${s.y2 < s.y1 ? ' flipV="1"' : ''}`;
    const extra = `${s.dash ? '<a:prstDash val="dash"/>' : ''}${s.arrowEnd ? '<a:tailEnd type="triangle" w="med" len="med"/>' : ''}`;
    return `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvCxnSpPr/>${nv}</p:nvCxnSpPr><p:spPr><a:xfrm${flip}><a:off x="${e(x)}" y="${e(y)}"/><a:ext cx="${e(Math.abs(s.x2 - s.x1))}" cy="${e(Math.abs(s.y2 - s.y1))}"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>${ln(s.color, s.width, extra)}</p:spPr></p:cxnSp>`;
  }
  return textXml(s, id, font, name, nv);
}

/** Freeform as custom geometry, from the same flattened outline every renderer uses. */
function pathXml(s: PathShape, id: number, name: string, nv: string): string {
  const b = pathBounds(s);
  const w = Math.max(1, e(b.w)), h = Math.max(1, e(b.h));
  const pt = (p: { x: number; y: number }) => `<a:pt x="${e(p.x - b.x)}" y="${e(p.y - b.y)}"/>`;
  const paths = flattenPath(s.commands).map(({ points, closed }) =>
    `<a:path w="${w}" h="${h}"${s.fill && closed ? '' : ' fill="none"'}><a:moveTo>${pt(points[0])}</a:moveTo>${points.slice(1).map((p) => `<a:lnTo>${pt(p)}</a:lnTo>`).join('')}${closed ? '<a:close/>' : ''}</a:path>`,
  ).join('');
  const dash = s.dash ? '<a:prstDash val="dash"/>' : '';
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/>${nv}</p:nvSpPr><p:spPr><a:xfrm><a:off x="${e(b.x)}" y="${e(b.y)}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>${paths}</a:pathLst></a:custGeom>${fill(s.fill)}${ln(s.stroke, s.strokeWidth, dash)}</p:spPr></p:sp>`;
}

function textXml(t: TextShape, id: number, font: string, name: string, nv = '<p:nvPr/>'): string {
  const anchor = { top: 't', middle: 'ctr', bottom: 'b' }[t.valign];
  const algn = { left: 'l', center: 'ctr', right: 'r' }[t.align];
  const geom = t.rounded ? '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst></a:prstGeom>' : '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
  const rPr = `<a:rPr lang="en-US" sz="${Math.round(t.size * 100)}"${t.bold ? ' b="1"' : ''} kern="400000" dirty="0">${fill(t.color)}<a:latin typeface="${esc(font)}"/><a:cs typeface="${esc(font)}"/></a:rPr>`;
  const paras = t.text.split('\n').map((line) =>
    `<a:p><a:pPr algn="${algn}"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc></a:pPr><a:r>${rPr}<a:t>${esc(line)}</a:t></a:r></a:p>`).join('');
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr txBox="1"/>${nv}</p:nvSpPr><p:spPr><a:xfrm><a:off x="${e(t.x)}" y="${e(t.y)}"/><a:ext cx="${e(t.w)}" cy="${e(t.h)}"/></a:xfrm>${geom}${fill(t.fill)}${ln(t.stroke)}</p:spPr><p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="${anchor}" rtlCol="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>${paras}</p:txBody></p:sp>`;
}

/**
 * One chart as a p:grpSp. Group tags go on the group, OC_ROLE from scene.roles[i] on each child.
 * Used by writePptx and by template fill (which supplies its own ids and tag registration).
 */
export function groupXml(c: SceneContent, ids: Ids, font: string, addTags?: AddTags): string {
  const f = c.scene.frame;
  const roles = (c.scene as Scene & { roles?: string[] }).roles;
  const xfrm = `<a:off x="${e(f.x)}" y="${e(f.y)}"/><a:ext cx="${e(f.w)}" cy="${e(f.h)}"/><a:chOff x="${e(f.x)}" y="${e(f.y)}"/><a:chExt cx="${e(f.w)}" cy="${e(f.h)}"/>`;
  const id = ids.next();
  const nv = nvPr(c.tags?.length && addTags ? addTags(c.tags) : undefined);
  const body = c.scene.shapes.map((s, i) => shapeXml(s, ids, font, roles?.[i] || undefined, addTags)).join('');
  return `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="${esc(c.name)}" descr="${esc(c.altText)}"/><p:cNvGrpSpPr/>${nv}</p:nvGrpSpPr><p:grpSpPr><a:xfrm>${xfrm}</a:xfrm></p:grpSpPr>${body}</p:grpSp>`;
}

function slideXml(slide: SlideContent, font: string, addTags: AddTags): string {
  const ids = new Ids();
  let body = '';
  if (slide.title) {
    body += textXml({ kind: 'text', x: 36, y: 24, w: SLIDE.w - 72, h: 44, text: slide.title, size: 24, color: '1A1A1A', bold: true, align: 'left', valign: 'top' }, ids.next(), font, TITLE_NAME);
  }
  for (const s of slide.scenes) body += groupXml(s, ids, font, addTags);
  return `${XML}<p:sld ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${body}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

const rels = (items: [string, string][]) =>
  `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.map(([type, target], i) => `<Relationship Id="rId${i + 1}" Type="${REL}/${type}" Target="${target}"/>`).join('')}</Relationships>`;

const EMPTY_TREE = '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree>';

/** theme1.xml of a CLI deck: the style's theme colours (default: the open-charts theme) and font. */
function theme(font: string, t: ThemeColors): string {
  const clr = (n: string, v: string) => `<a:${n}><a:srgbClr val="${v}"/></a:${n}>`;
  const solid = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  return `${XML}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="open-charts"><a:themeElements><a:clrScheme name="open-charts">${(['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'] as const).map((k) => clr(k, t[k])).join('')}${clr('hlink', '0563C1')}${clr('folHlink', '954F72')}</a:clrScheme><a:fontScheme name="open-charts"><a:majorFont><a:latin typeface="${esc(font)}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="${esc(font)}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="open-charts"><a:fillStyleLst>${solid}${solid}${solid}</a:fillStyleLst><a:lnStyleLst><a:ln w="6350">${solid}</a:ln><a:ln w="12700">${solid}</a:ln><a:ln w="19050">${solid}</a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst>${solid}${solid}${solid}</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
}

export function writePptx(slides: SlideContent[], font = 'Arial', themeColors: ThemeColors = DEFAULT_THEME): Uint8Array {
  const n = slides.length;
  const tagParts: string[] = [];
  const slideFiles: { name: string; data: string }[] = [];
  slides.forEach((s, i) => {
    const slideRels: [string, string][] = [['slideLayout', '../slideLayouts/slideLayout1.xml']];
    const addTags: AddTags = (tags) => {
      tagParts.push(tagPartXml(tags));
      slideRels.push(['tags', `../tags/tag${tagParts.length}.xml`]);
      return `rId${slideRels.length}`;
    };
    slideFiles.push({ name: `ppt/slides/slide${i + 1}.xml`, data: slideXml(s, font, addTags) });
    slideFiles.push({ name: `ppt/slides/_rels/slide${i + 1}.xml.rels`, data: rels(slideRels) });
  });

  const files: { name: string; data: string }[] = [];
  files.push({
    name: '[Content_Types].xml',
    data: `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>${slides.map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')}${tagParts.map((_, i) => `<Override PartName="/ppt/tags/tag${i + 1}.xml" ContentType="${TAGS_CT}"/>`).join('')}</Types>`,
  });
  files.push({ name: '_rels/.rels', data: rels([['officeDocument', 'ppt/presentation.xml']]) });
  files.push({
    name: 'ppt/presentation.xml',
    data: `${XML}<p:presentation ${NS} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 3}"/>`).join('')}</p:sldIdLst><p:sldSz cx="${e(SLIDE.w)}" cy="${e(SLIDE.h)}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
  });
  files.push({
    name: 'ppt/_rels/presentation.xml.rels',
    data: rels([['slideMaster', 'slideMasters/slideMaster1.xml'], ['theme', 'theme/theme1.xml'], ...Array.from({ length: n }, (_, i) => ['slide', `slides/slide${i + 1}.xml`] as [string, string])]),
  });
  files.push({
    name: 'ppt/slideMasters/slideMaster1.xml',
    data: `${XML}<p:sldMaster ${NS}><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>${EMPTY_TREE}</p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`,
  });
  files.push({ name: 'ppt/slideMasters/_rels/slideMaster1.xml.rels', data: rels([['slideLayout', '../slideLayouts/slideLayout1.xml'], ['theme', '../theme/theme1.xml']]) });
  files.push({
    name: 'ppt/slideLayouts/slideLayout1.xml',
    data: `${XML}<p:sldLayout ${NS} type="blank" preserve="1"><p:cSld name="Blank">${EMPTY_TREE}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
  });
  files.push({ name: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', data: rels([['slideMaster', '../slideMasters/slideMaster1.xml']]) });
  files.push({ name: 'ppt/theme/theme1.xml', data: theme(font, themeColors) });
  files.push(...slideFiles);
  tagParts.forEach((data, i) => files.push({ name: `ppt/tags/tag${i + 1}.xml`, data }));
  return zip(files);
}
