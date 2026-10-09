// Writes native PowerPoint (OOXML) slides. Each chart becomes one grouped object of plain shapes,
// so the deck opens and edits in any PowerPoint, Keynote or LibreOffice without this tool.
import type { Scene, Shape, TextShape } from '../core/types.ts';
import { zip } from './zip.ts';

export interface SlideContent {
  title?: string;
  scenes: { scene: Scene; name: string; altText: string }[];
}

export const SLIDE = { w: 960, h: 540 }; // points, 16:9
const EMU = 12700;
const e = (pt: number) => Math.round(pt * EMU);

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const fill = (c?: string) => (c ? `<a:solidFill><a:srgbClr val="${c.replace('#', '')}"/></a:solidFill>` : '<a:noFill/>');
const ln = (c?: string, w = 0.75, extra = '') =>
  c ? `<a:ln w="${e(w)}">${fill(c)}${extra}</a:ln>` : '<a:ln><a:noFill/></a:ln>';

class Ids { n = 1; next() { return ++this.n; } }

function shapeXml(s: Shape, ids: Ids, font: string): string {
  const id = ids.next();
  if (s.kind === 'rect') {
    return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Bar ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${e(s.x)}" y="${e(s.y)}"/><a:ext cx="${e(s.w)}" cy="${e(s.h)}"/></a:xfrm><a:prstGeom prst="${s.ellipse ? 'ellipse' : 'rect'}"><a:avLst/></a:prstGeom>${fill(s.fill)}${ln(s.stroke, s.strokeWidth)}</p:spPr></p:sp>`;
  }
  if (s.kind === 'line') {
    const x = Math.min(s.x1, s.x2), y = Math.min(s.y1, s.y2);
    const flip = `${s.x2 < s.x1 ? ' flipH="1"' : ''}${s.y2 < s.y1 ? ' flipV="1"' : ''}`;
    const extra = `${s.dash ? '<a:prstDash val="dash"/>' : ''}${s.arrowEnd ? '<a:tailEnd type="triangle" w="med" len="med"/>' : ''}`;
    return `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="Line ${id}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm${flip}><a:off x="${e(x)}" y="${e(y)}"/><a:ext cx="${e(Math.abs(s.x2 - s.x1))}" cy="${e(Math.abs(s.y2 - s.y1))}"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>${ln(s.color, s.width, extra)}</p:spPr></p:cxnSp>`;
  }
  return textXml(s, id, font);
}

function textXml(t: TextShape, id: number, font: string): string {
  const anchor = { top: 't', middle: 'ctr', bottom: 'b' }[t.valign];
  const algn = { left: 'l', center: 'ctr', right: 'r' }[t.align];
  const geom = t.rounded ? '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst></a:prstGeom>' : '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
  const rPr = `<a:rPr lang="en-US" sz="${Math.round(t.size * 100)}"${t.bold ? ' b="1"' : ''} dirty="0">${fill(t.color)}<a:latin typeface="${esc(font)}"/><a:cs typeface="${esc(font)}"/></a:rPr>`;
  const paras = t.text.split('\n').map((line) =>
    `<a:p><a:pPr algn="${algn}"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc></a:pPr><a:r>${rPr}<a:t>${esc(line)}</a:t></a:r></a:p>`).join('');
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Label ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${e(t.x)}" y="${e(t.y)}"/><a:ext cx="${e(t.w)}" cy="${e(t.h)}"/></a:xfrm>${geom}${fill(t.fill)}${ln(t.stroke)}</p:spPr><p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="${anchor}" rtlCol="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>${paras}</p:txBody></p:sp>`;
}

function groupXml(scene: Scene, name: string, alt: string, ids: Ids, font: string): string {
  const f = scene.frame;
  const xfrm = `<a:off x="${e(f.x)}" y="${e(f.y)}"/><a:ext cx="${e(f.w)}" cy="${e(f.h)}"/><a:chOff x="${e(f.x)}" y="${e(f.y)}"/><a:chExt cx="${e(f.w)}" cy="${e(f.h)}"/>`;
  const id = ids.next();
  const body = scene.shapes.map((s) => shapeXml(s, ids, font)).join('');
  return `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="${esc(name)}" descr="${esc(alt)}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm>${xfrm}</a:xfrm></p:grpSpPr>${body}</p:grpSp>`;
}

function slideXml(slide: SlideContent, font: string): string {
  const ids = new Ids();
  let body = '';
  if (slide.title) {
    body += textXml({ kind: 'text', x: 36, y: 24, w: SLIDE.w - 72, h: 44, text: slide.title, size: 24, color: '1A1A1A', bold: true, align: 'left', valign: 'top' }, ids.next(), font);
  }
  for (const s of slide.scenes) body += groupXml(s.scene, s.name, s.altText, ids, font);
  return `${XML}<p:sld ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${body}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

const rels = (items: [string, string][]) =>
  `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.map(([type, target], i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`).join('')}</Relationships>`;

const EMPTY_TREE = '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree>';

function theme(font: string): string {
  const clr = (n: string, v: string) => `<a:${n}><a:srgbClr val="${v}"/></a:${n}>`;
  const solid = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  return `${XML}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="open-charts"><a:themeElements><a:clrScheme name="open-charts">${clr('dk1', '000000')}${clr('lt1', 'FFFFFF')}${clr('dk2', '1F3A5F')}${clr('lt2', 'EEEEEE')}${clr('accent1', '1F3A5F')}${clr('accent2', '3E6A9E')}${clr('accent3', '7FA3CC')}${clr('accent4', 'B9CDE5')}${clr('accent5', 'A6A6A6')}${clr('accent6', 'D9D9D9')}${clr('hlink', '0563C1')}${clr('folHlink', '954F72')}</a:clrScheme><a:fontScheme name="open-charts"><a:majorFont><a:latin typeface="${esc(font)}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="${esc(font)}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="open-charts"><a:fillStyleLst>${solid}${solid}${solid}</a:fillStyleLst><a:lnStyleLst><a:ln w="6350">${solid}</a:ln><a:ln w="12700">${solid}</a:ln><a:ln w="19050">${solid}</a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst>${solid}${solid}${solid}</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
}

export function writePptx(slides: SlideContent[], font = 'Arial'): Uint8Array {
  const n = slides.length;
  const files: { name: string; data: string }[] = [];
  files.push({
    name: '[Content_Types].xml',
    data: `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>${slides.map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')}</Types>`,
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
  files.push({ name: 'ppt/theme/theme1.xml', data: theme(font) });
  slides.forEach((s, i) => {
    files.push({ name: `ppt/slides/slide${i + 1}.xml`, data: slideXml(s, font) });
    files.push({ name: `ppt/slides/_rels/slide${i + 1}.xml.rels`, data: rels([['slideLayout', '../slideLayouts/slideLayout1.xml']]) });
  });
  return zip(files);
}
