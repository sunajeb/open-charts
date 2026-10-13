// Writes examples/template/template.pptx: a realistic 8-slide "client" template for `open-charts fill`
// and its tests. A slide master with title / body placeholders and text styles, three layouts, a
// brand theme (teal / amber accents), a notes master and a notes slide, two sections, a table, a
// picture, and slots of every kind: a named rectangle (`chart:revenue`), content placeholders that
// inherit their frame from the layout (`chart:mix_left`, `chart:mix_right`), a `{{chart:…}}` text
// slot, and `{{fields}}` — one split across runs the way PowerPoint splits them. Parts are deflated
// like PowerPoint's own, so the byte-identical test covers copied compressed entries.
//
//   node --experimental-strip-types scripts/make-template.ts [out.pptx]
import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname } from 'node:path';
import { deflateRawSync } from 'node:zlib';
import { crc32, zip } from '../src/render/zip.ts';
import { encodePng } from '../test/goldens/png.ts';

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument';
const E = (pt: number) => Math.round(pt * 12700);
const W = 960, H = 540;

const rels = (items: [string, string][]) =>
  `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.map(([t, target], i) => `<Relationship Id="rId${i + 1}" Type="${t.startsWith('http') ? t : `${R}/${t}`}" Target="${target}"/>`).join('')}</Relationships>`;
const xfrm = (x: number, y: number, w: number, h: number) => `<a:xfrm><a:off x="${E(x)}" y="${E(y)}"/><a:ext cx="${E(w)}" cy="${E(h)}"/></a:xfrm>`;
const tree = (body: string) => `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${body}</p:spTree>`;
const run = (t: string, rPr = '<a:rPr lang="en-US" dirty="0"/>') => `<a:r>${rPr}<a:t>${t}</a:t></a:r>`;
const para = (...runs: string[]) => `<a:p>${runs.join('')}</a:p>`;
const body = (...paras: string[]) => `<p:txBody><a:bodyPr/><a:lstStyle/>${paras.join('')}</p:txBody>`;
/** A placeholder shape (no xfrm: it inherits its frame from the layout). */
const ph = (id: number, name: string, phXml: string, ...paras: string[]) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${phXml}</p:nvPr></p:nvSpPr><p:spPr/>${paras.length ? body(...paras) : ''}</p:sp>`;
/** A placeholder with its own frame (layouts and masters). */
const phAt = (id: number, name: string, phXml: string, f: [number, number, number, number], txt = '', extra = '') =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${phXml}</p:nvPr></p:nvSpPr><p:spPr>${xfrm(...f)}${extra}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p>${txt ? run(txt) : ''}<a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
const box = (id: number, name: string, f: [number, number, number, number], paras: string[], fill = '', anchor = 't') =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(...f)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fill ? `<a:solidFill>${fill}</a:solidFill>` : '<a:noFill/>'}</p:spPr><p:txBody><a:bodyPr wrap="square" anchor="${anchor}"/><a:lstStyle/>${paras.join('')}</p:txBody></p:sp>`;
const SMALL = '<a:rPr lang="en-US" sz="1000" dirty="0"><a:solidFill><a:schemeClr val="tx2"/></a:solidFill></a:rPr>';
const SLOT_FILL = '<a:schemeClr val="bg2"/>';

// ---- Theme: "Harbor" — teal and amber, Arial ------------------------------------------------------

function theme(name: string): string {
  const c = (k: string, v: string) => `<a:${k}><a:srgbClr val="${v}"/></a:${k}>`;
  const ph = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  return `${XML}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${name}"><a:themeElements><a:clrScheme name="Harbor"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>${c('dk2', '0B3C49')}${c('lt2', 'EEF2F3')}${c('accent1', '0F6E7A')}${c('accent2', 'F2A541')}${c('accent3', '5FA8A0')}${c('accent4', 'A8D5CF')}${c('accent5', '7A7A7A')}${c('accent6', 'C8C8C8')}${c('hlink', '0F6E7A')}${c('folHlink', '7A4F9E')}</a:clrScheme><a:fontScheme name="Harbor"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Harbor"><a:fillStyleLst>${ph}${ph}${ph}</a:fillStyleLst><a:lnStyleLst><a:ln w="6350">${ph}</a:ln><a:ln w="12700">${ph}</a:ln><a:ln w="19050">${ph}</a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst>${ph}${ph}${ph}</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
}

const lvl = (n: number, sz: number, extra = '') => `<a:lvl${n}pPr marL="${n === 1 ? 0 : (n - 1) * 228600}" indent="0"${extra}><a:defRPr sz="${sz}"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl${n}pPr>`;
const CLRMAP = 'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"';

const master = `${XML}<p:sldMaster ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>${tree(
  phAt(2, 'Title Placeholder 1', '<p:ph type="title"/>', [36, 22, 888, 56], 'Click to edit Master title style')
  + phAt(3, 'Text Placeholder 2', '<p:ph type="body" idx="1"/>', [36, 96, 888, 384], 'Click to edit Master text styles')
  + phAt(4, 'Footer Placeholder 3', '<p:ph type="ftr" sz="quarter" idx="3"/>', [36, 504, 600, 20])
  // The brand bar.
  + `<p:sp><p:nvSpPr><p:cNvPr id="5" name="Brand bar"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr><p:spPr>${xfrm(0, 0, W, 8)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>`,
)}</p:cSld><p:clrMap ${CLRMAP}/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId2"/><p:sldLayoutId id="2147483651" r:id="rId3"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="2600" b="1"><a:solidFill><a:schemeClr val="tx2"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle>${lvl(1, 1800)}${lvl(2, 1600)}${lvl(3, 1400)}</p:bodyStyle><p:otherStyle>${lvl(1, 1800)}</p:otherStyle></p:txStyles></p:sldMaster>`;

const layout = (name: string, type: string, shapes: string) =>
  `${XML}<p:sldLayout ${NS} type="${type}" preserve="1"><p:cSld name="${name}">${tree(shapes)}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
const layouts = [
  layout('Title Slide', 'title', phAt(2, 'Title 1', '<p:ph type="ctrTitle"/>', [72, 180, 816, 90], 'Click to edit title') + phAt(3, 'Subtitle 2', '<p:ph type="subTitle" idx="1"/>', [72, 280, 816, 50], 'Click to edit subtitle')),
  layout('Title and Content', 'obj', phAt(2, 'Title 1', '<p:ph type="title"/>', [36, 22, 888, 56]) + phAt(3, 'Content Placeholder 2', '<p:ph idx="1"/>', [36, 96, 888, 384])),
  layout('Two Content', 'twoObj', phAt(2, 'Title 1', '<p:ph type="title"/>', [36, 22, 888, 56]) + phAt(3, 'Content Placeholder 2', '<p:ph sz="half" idx="1"/>', [36, 96, 432, 384]) + phAt(4, 'Content Placeholder 3', '<p:ph sz="half" idx="2"/>', [492, 96, 432, 384])),
];

const notesMaster = `${XML}<p:notesMaster ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>${tree(
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr>${xfrm(54, 54, 432, 243)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr></p:sp>`
  + phAt(3, 'Notes Placeholder 2', '<p:ph type="body" sz="quarter" idx="3"/>', [54, 320, 432, 300]),
)}</p:cSld><p:clrMap ${CLRMAP}/><p:notesStyle>${lvl(1, 1200)}</p:notesStyle></p:notesMaster>`;

const notesSlide = (text: string[]) => `${XML}<p:notes ${NS}><p:cSld>${tree(
  ph(2, 'Slide Image Placeholder 1', '<p:ph type="sldImg"/>') + ph(3, 'Notes Placeholder 2', '<p:ph type="body" idx="1"/>', para(...text)),
)}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;

// ---- Slides ------------------------------------------------------------------------------------

const slide = (shapes: string) => `${XML}<p:sld ${NS}><p:cSld>${tree(shapes)}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;

const cell = (t: string, head = false) => `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="${head ? 'l' : 'r'}"/>${run(t, `<a:rPr lang="en-US" sz="1400"${head ? ' b="1"' : ''} dirty="0"/>`)}</a:p></a:txBody><a:tcPr/></a:tc>`;
const table = (id: number, rows: string[][]) => `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="KPI table"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${E(72)}" y="${E(120)}"/><a:ext cx="${E(816)}" cy="${E(rows.length * 40)}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr><a:tblGrid>${rows[0].map(() => `<a:gridCol w="${E(816 / rows[0].length)}"/>`).join('')}</a:tblGrid>${rows.map((r, i) => `<a:tr h="${E(40)}">${r.map((t, j) => cell(t, i === 0 || j === 0)).join('')}</a:tr>`).join('')}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;

const slides: { xml: string; layout: number; notes?: string[]; image?: boolean }[] = [
  { // 1 Cover: the client field split across three runs, as PowerPoint does after an edit.
    layout: 1,
    xml: slide(ph(2, 'Title 1', '<p:ph type="ctrTitle"/>', para(run('{{'), run('client', '<a:rPr lang="en-GB" dirty="0"/>'), run('}}: {{project}}')))
      + ph(3, 'Subtitle 2', '<p:ph type="subTitle" idx="1"/>', para(run('Board update · {{date}}')))),
  },
  { // 2 Agenda: no slots (untouched).
    layout: 2,
    xml: slide(ph(2, 'Title 1', '<p:ph type="title"/>', para(run('Agenda'))) + ph(3, 'Content Placeholder 2', '<p:ph idx="1"/>', para(run('Where revenue comes from')), para(run('Channel mix')), para(run('Regions')), para(run('Next steps')))),
  },
  { // 3 One chart in a named rectangle; source line; speaker notes with a field.
    layout: 2, notes: [run('Talk track: '), run('{{revenue_note}}')],
    xml: slide(ph(2, 'Title 1', '<p:ph type="title"/>', para(run('{{revenue_title}}')))
      + box(4, 'chart:revenue', [36, 96, 888, 372], [para(run('Revenue chart', '<a:rPr lang="en-US" sz="1400" dirty="0"><a:solidFill><a:schemeClr val="tx2"/></a:solidFill></a:rPr>'))], SLOT_FILL, 'ctr')
      + box(5, 'Source', [36, 476, 888, 20], [para(run('Source: {{source}}', SMALL))])),
  },
  { // 4 Two charts in content placeholders (frames inherited from the "Two Content" layout).
    layout: 3,
    xml: slide(ph(2, 'Title 1', '<p:ph type="title"/>', para(run('{{mix_title}}')))
      + ph(3, 'chart:mix_left', '<p:ph sz="half" idx="1"/>')
      + ph(4, 'chart:mix_right', '<p:ph sz="half" idx="2"/>')),
  },
  { // 5 One slide per region (fill data repeats it with "each"): a {{chart:…}} text slot.
    layout: 2,
    xml: slide(ph(2, 'Title 1', '<p:ph type="title"/>', para(run('{{region}}: '), run('{{headline}}')))
      + box(4, 'Rectangle 3', [36, 96, 600, 384], [para(run('{{chart:region_sales}}'))], SLOT_FILL, 'ctr')
      + box(5, 'Comment', [660, 96, 264, 384], [para(run('{{region_comment}}', '<a:rPr lang="en-US" sz="1400" dirty="0"/>'))], '<a:schemeClr val="lt2"/>')),
  },
  { // 6 Fields in a table.
    layout: 2,
    xml: slide(ph(2, 'Title 1', '<p:ph type="title"/>', para(run('Key figures')))
      + table(3, [['', 'FY2025', 'FY2026'], ['Revenue', '{{rev_25}}', '{{rev_26}}'], ['EBITDA margin', '{{margin_25}}', '{{margin_26}}'], ['Headcount', '{{fte_25}}', '{{fte_26}}']])),
  },
  { // 7 A picture (untouched; the media part must come out byte-identical).
    layout: 2, image: true,
    xml: slide(ph(2, 'Title 1', '<p:ph type="title"/>', para(run('Our footprint')))
      + `<p:pic><p:nvPicPr><p:cNvPr id="3" name="Picture 2" descr="Footprint map"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrm(240, 110, 480, 340)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`),
  },
  { // 8 Closing (untouched).
    layout: 1,
    xml: slide(ph(2, 'Title 1', '<p:ph type="ctrTitle"/>', para(run('Thank you'))) + ph(3, 'Subtitle 2', '<p:ph type="subTitle" idx="1"/>', para(run('Questions and discussion')))),
  },
];

function picture(): Uint8Array {
  // A simple teal "map" placeholder picture (stripes), so the slide has real media.
  const w = 240, h = 170, data = new Uint8Array(w * h * 4);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const i = (y * w + x) * 4, band = ((x + y) >> 4) & 1;
    data.set(band ? [15, 110, 122, 255] : [168, 213, 207, 255], i);
  }
  return encodePng({ width: w, height: h, data });
}

export function makeTemplate(): Uint8Array {
  const parts: { name: string; data: string | Uint8Array }[] = [];
  const add = (name: string, data: string | Uint8Array) => parts.push({ name, data });
  const n = slides.length;
  add('[Content_Types].xml', `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="${CT}.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${CT}.presentationml.slideMaster+xml"/>${slides.map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="${CT}.presentationml.slide+xml"/>`).join('')}<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="${CT}.presentationml.notesMaster+xml"/><Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="${CT}.presentationml.notesSlide+xml"/><Override PartName="/ppt/presProps.xml" ContentType="${CT}.presentationml.presProps+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="${CT}.presentationml.viewProps+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="${CT}.theme+xml"/><Override PartName="/ppt/theme/theme2.xml" ContentType="${CT}.theme+xml"/><Override PartName="/ppt/tableStyles.xml" ContentType="${CT}.presentationml.tableStyles+xml"/>${layouts.map((_, i) => `<Override PartName="/ppt/slideLayouts/slideLayout${i + 1}.xml" ContentType="${CT}.presentationml.slideLayout+xml"/>`).join('')}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="${CT}.extended-properties+xml"/></Types>`);
  add('_rels/.rels', rels([['officeDocument', 'ppt/presentation.xml'], ['http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml'], ['extended-properties', 'docProps/app.xml']]));
  // presentation.xml.rels: master, slides, notes master, props, theme, table styles.
  const presRels: [string, string][] = [['slideMaster', 'slideMasters/slideMaster1.xml'], ...slides.map((_, i) => ['slide', `slides/slide${i + 1}.xml`] as [string, string]), ['notesMaster', 'notesMasters/notesMaster1.xml'], ['presProps', 'presProps.xml'], ['viewProps', 'viewProps.xml'], ['theme', 'theme/theme1.xml'], ['tableStyles', 'tableStyles.xml']];
  const sid = (i: number) => 256 + i;
  const section = (name: string, guid: string, ids: number[]) => `<p14:section name="${name}" id="{${guid}}"><p14:sldIdLst>${ids.map((i) => `<p14:sldId id="${sid(i)}"/>`).join('')}</p14:sldIdLst></p14:section>`;
  add('ppt/presentation.xml', `${XML}<p:presentation ${NS} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:notesMasterIdLst><p:notesMasterId r:id="rId${n + 2}"/></p:notesMasterIdLst><p:sldIdLst>${slides.map((_, i) => `<p:sldId id="${sid(i)}" r:id="rId${i + 2}"/>`).join('')}</p:sldIdLst><p:sldSz cx="${E(W)}" cy="${E(H)}"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle>${lvl(1, 1800)}</p:defaultTextStyle><p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">${section('Introduction', '6C1A7E52-3A0B-4C58-9C0E-1F2B3C4D5E61', [0, 1])}${section('Analysis', '7D2B8F63-4B1C-4D69-8D1F-2A3B4C5D6E72', [2, 3, 4, 5, 6])}${section('Close', '8E3C9074-5C2D-4E7A-9E20-3B4C5D6E7F83', [7])}</p14:sectionLst></p:ext></p:extLst></p:presentation>`);
  add('ppt/_rels/presentation.xml.rels', rels(presRels));
  add('ppt/presProps.xml', `${XML}<p:presentationPr ${NS}/>`);
  add('ppt/viewProps.xml', `${XML}<p:viewPr ${NS}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`);
  add('ppt/tableStyles.xml', `${XML}<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`);
  add('ppt/slideMasters/slideMaster1.xml', master);
  add('ppt/slideMasters/_rels/slideMaster1.xml.rels', rels([...layouts.map((_, i) => ['slideLayout', `../slideLayouts/slideLayout${i + 1}.xml`] as [string, string]), ['theme', '../theme/theme1.xml']]));
  layouts.forEach((l, i) => {
    add(`ppt/slideLayouts/slideLayout${i + 1}.xml`, l);
    add(`ppt/slideLayouts/_rels/slideLayout${i + 1}.xml.rels`, rels([['slideMaster', '../slideMasters/slideMaster1.xml']]));
  });
  add('ppt/theme/theme1.xml', theme('Harbor'));
  add('ppt/theme/theme2.xml', theme('Office Theme'));
  add('ppt/notesMasters/notesMaster1.xml', notesMaster);
  add('ppt/notesMasters/_rels/notesMaster1.xml.rels', rels([['theme', '../theme/theme2.xml']]));
  slides.forEach((s, i) => {
    add(`ppt/slides/slide${i + 1}.xml`, s.xml);
    const r: [string, string][] = [['slideLayout', `../slideLayouts/slideLayout${s.layout}.xml`]];
    if (s.image) r.push(['image', '../media/image1.png']);
    if (s.notes) r.push(['notesSlide', '../notesSlides/notesSlide1.xml']);
    add(`ppt/slides/_rels/slide${i + 1}.xml.rels`, rels(r));
    if (s.notes) {
      add('ppt/notesSlides/notesSlide1.xml', notesSlide(s.notes));
      add('ppt/notesSlides/_rels/notesSlide1.xml.rels', rels([['notesMaster', '../notesMasters/notesMaster1.xml'], ['slide', `../slides/slide${i + 1}.xml`]]));
    }
  });
  add('ppt/media/image1.png', picture());
  add('docProps/core.xml', `${XML}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Harbor client template</dc:title><dc:creator>open-charts</dc:creator></cp:coreProperties>`);
  add('docProps/app.xml', `${XML}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office PowerPoint</Application><Slides>${n}</Slides><Notes>1</Notes></Properties>`);
  // Deflate every part (PNG stays stored, as Office does), like a real saved deck.
  const enc = new TextEncoder();
  return zip(parts.map(({ name, data }) => {
    const bytes = typeof data === 'string' ? enc.encode(data) : data;
    if (name.endsWith('.png')) return { name, data: bytes };
    return { name, raw: { method: 8, crc: crc32(bytes), size: bytes.length, compressed: new Uint8Array(deflateRawSync(bytes)) } };
  }));
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const out = process.argv[2] ?? 'examples/template/template.pptx';
  mkdirSync(dirname(out), { recursive: true });
  writeFileSync(out, makeTemplate());
  console.log(`Wrote ${out}`);
}
