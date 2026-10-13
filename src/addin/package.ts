// Builds open-charts.ppam, the PowerPoint add-in package, the way the reference product ships its
// own: a macro-enabled add-in presentation with no slides, a ribbon (customUI/customUI14.xml) and a
// vbaProject.bin holding one thin VBA module. Every ribbon callback carries the same per-build
// prefix and lands in that module's dispatcher.
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { writePptx, XML } from '../render/pptx.ts';
import { unzip } from '../ooxml/unzip.ts';
import { zip, type ZipInput } from '../render/zip.ts';
import { ENGINE_VERSION } from '../version.ts';
import { flatLabel, galleryItems, iconOf, tileDirections, type ChartTypeEntry } from './registry.ts';
import { iconFile, iconPng, TILE, tilePng } from './icons.ts';
import { DIRECTIONS } from '../core/spec.ts';
import { vbaProjectBin } from './vba-project.ts';
import { reinstallCommand } from './launcher.ts';

export const MODULE_NAME = 'OpenCharts';
// Must differ from every module name, or VBA refuses to open the project.
export const PROJECT_NAME = 'OpenChartsAddIn';
export const BAS_FILE = new URL('../../addin/vba/OpenCharts.bas', import.meta.url);
export const CLS_FILE = new URL('../../addin/vba/AppEvents.cls', import.meta.url);
const PLACEHOLDER = /OCPREFIX/g;

export const ADDIN_CT = 'application/vnd.ms-powerpoint.addin.macroEnabled.main+xml';
export const VBA_CT = 'application/vnd.ms-office.vbaProject';
export const VBA_REL = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject';
export const UI_REL = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility';
export const UI_NS = 'http://schemas.microsoft.com/office/2009/07/customui';
export const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
/** The ribbon's icons (tools/icons.ts writes them from src/addin/icons.ts). */
export const ICON_DIR = new URL('../../addin/icons/', import.meta.url);

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Per-build callback prefix: derived from what is built, so identical inputs give identical bytes. */
export function buildPrefix(...parts: string[]): string {
  return `OC${createHash('sha1').update(parts.join('\0')).digest('hex').slice(0, 24).toUpperCase()}`;
}

/** customUI image id of an icon (a relationship id in customUI/_rels/customUI14.xml.rels). */
const img = (icon: string) => `ic_${icon}`;

/** The direction buttons beside the gallery: ▲ ▼ ◀ ▶. */
export const DIRS = [
  { id: 'dirUp', icon: 'dirUp', label: 'Up', tip: 'columns growing up' },
  { id: 'dirDown', icon: 'dirDown', label: 'Down', tip: 'columns hanging down from the top' },
  { id: 'dirLeft', icon: 'dirLeft', label: 'Left', tip: 'bars growing to the left' },
  { id: 'dirRight', icon: 'dirRight', label: 'Right', tip: 'bars growing to the right (a line chart becomes a profile)' },
] as const;

/** One image the ribbon shows: its customUI id, its PNG file in addin/icons, and how to draw it. */
export interface RibbonImage { id: string; file: string; draw: () => Buffer }

/**
 * The ribbon's images: one tile per gallery item (icon over label, drawn at 2× for Retina), the
 * gallery button's icon, and the four direction arrows. tools/icons.ts writes them to addin/icons.
 */
export function ribbonImages(items: ChartTypeEntry[]): RibbonImage[] {
  return [
    ...items.map((g) => ({ id: tileImage(g), file: `tile-${g.id}.png`, draw: () => tilePng(iconOf(g), g.label) })),
    { id: img('stacked'), file: iconFile('stacked', 64), draw: () => iconPng('stacked', 64) },
    ...DIRS.map((d) => ({ id: img(d.icon), file: iconFile(d.icon, 32), draw: () => iconPng(d.icon, 32) })),
  ];
}
const tileImage = (g: ChartTypeEntry) => `tile_${g.id.replace(/[^A-Za-z0-9]/g, '_')}`;

export function ribbonXml(prefix: string, items: ChartTypeEntry[]): string {
  const cb = (name: string) => `${prefix}_${name}`;
  // The gallery: one tile per chart, its icon over its label (both in the tile image: PowerPoint for
  // Mac would put a label beside the image), 5 per row like the reference's.
  const tiles = items.map((g) => {
    const dirs = tileDirections(g);
    const tip = `${g.description ?? ''}${dirs.length > 1 ? ` Directions: ${dirs.join(', ')} (the arrows beside the gallery).` : ''}`.trim();
    return `<item id="chart_${esc(g.id)}" label="${esc(flatLabel(g))}" image="${tileImage(g)}" screentip="${esc(flatLabel(g))}"${tip ? ` supertip="${esc(tip)}"` : ''}/>`;
  }).join('');
  const dirButton = (d: typeof DIRS[number]) =>
    `<toggleButton id="${d.id}" label="${d.label}" showLabel="false" image="${img(d.icon)}" screentip="Direction: ${d.label.toLowerCase()}" supertip="With a chart selected: redraw it with ${d.tip}. With none selected: the next chart you insert grows this way." onAction="${cb('onDirection')}" getPressed="${cb('getDirPressed')}" getEnabled="${cb('getDirEnabled')}"/>`;
  return `${XML}<customUI xmlns="${UI_NS}" onLoad="${cb('onLoad')}">
<ribbon><tabs>
<tab id="tabOpenCharts" label="open-charts" insertBeforeMso="TabHome">
<group id="grpCharts" label="Charts" image="${img('stacked')}">
<gallery id="galCharts" label="Insert chart" image="${img('stacked')}" size="large" columns="5" itemWidth="${TILE.w}" itemHeight="${TILE.h}" showItemLabel="false" screentip="Insert chart" supertip="Insert a chart on this slide with sample data. It goes into an empty content placeholder if there is one. Pick a direction with the arrows beside the gallery first, or change it afterwards. Then click Edit data to type your numbers." onAction="${cb('onGallery')}">${tiles}</gallery>
<labelControl id="lblDirection" label="Direction"/>
<box id="boxDirUpDown" boxStyle="horizontal">${dirButton(DIRS[0])}${dirButton(DIRS[1])}</box>
<box id="boxDirLeftRight" boxStyle="horizontal">${dirButton(DIRS[2])}${dirButton(DIRS[3])}</box>
<separator id="sepCharts"/>
<button id="btnEditData" label="Edit data" imageMso="ChartEditDataSource" size="large" screentip="Edit data" supertip="Open the selected chart's datasheet. Apply there and the chart on the slide updates." onAction="${cb('onAction')}" getEnabled="${cb('getEnabled')}"/>
<button id="btnUpdate" label="Update chart" imageMso="RefreshAll" size="large" screentip="Update chart" supertip="Lay the selected chart out again at its current size and position, with any applied edits. Resized charts also update by themselves when you click elsewhere." onAction="${cb('onAction')}" getEnabled="${cb('getEnabled')}"/>
</group>
${agendaGroupXml(cb('onAction'))}
</tab>
</tabs></ribbon>
</customUI>
`;
}

/** Agenda controls (docs/USER-GUIDE.md "Agenda"): [id suffix, label, imageMso, large, tip]. */
export const AGENDA_BUTTONS: [string, string, string, boolean, string][] = [
  ['Chapter', 'Chapter', 'SlideNew', true, 'Insert an agenda slide after this slide: it starts a new chapter. Type the chapter name over "Chapter". With a heading selected, the new chapter goes after that chapter.'],
  ['Toc', 'Table of Contents', 'TableOfContentsGallery', true, 'Insert a slide listing the top-level chapters, without highlighting, before the first chapter slide.'],
  ['Update', 'Update agenda', 'RefreshAll', true, 'Update every agenda slide: renamed, restyled, nested, added, moved or deleted chapters. Happens by itself when you click elsewhere after most changes.'],
  ['Up', 'Move up', 'OutlineMoveUp', false, 'Move the selected chapter up; its slides move along.'],
  ['Down', 'Move down', 'OutlineMoveDown', false, 'Move the selected chapter down; its slides move along.'],
  ['Indent', 'Demote', 'OutlineDemote', false, 'Make the selected chapter a subchapter (one text level down).'],
  ['Outdent', 'Promote', 'OutlinePromote', false, 'Move the selected chapter one text level up.'],
];
export const AGENDA_MENU: [string, string, string][] = [
  ['Split', 'Split Agenda', 'Start a new, independent agenda at this slide (e.g. the backup).'],
  ['Join', 'Join Agendas', 'Remove the agenda split on this slide.'],
  ['RemoveHidden', 'Remove Hidden Agenda Slides', 'Delete agenda slides that are hidden because their chapter has no content before its first subchapter. The chapters stay in the agenda.'],
  ['InsertHidden', 'Insert Hidden Agenda Slides', 'Insert the agenda slides removed with Remove Hidden Agenda Slides again.'],
  ['Placeholder', 'Agenda Placeholder', 'In Slide Master view, on a layout named "open-charts agenda": mark where agendas go on agenda slides.'],
];

function agendaGroupXml(onAction: string): string {
  const buttons = AGENDA_BUTTONS.map(([id, label, icon, large, tip]) =>
    `<button id="btnAgenda${id}" label="${esc(label)}" imageMso="${icon}"${large ? ' size="large"' : ''} screentip="${esc(label)}" supertip="${esc(tip)}" onAction="${onAction}"/>`).join('\n');
  const menu = AGENDA_MENU.map(([id, label, tip]) => `<button id="btnAgenda${id}" label="${esc(label)}" screentip="${esc(label)}" supertip="${esc(tip)}" onAction="${onAction}"/>`).join('');
  return `<group id="grpAgenda" label="Agenda" imageMso="TableOfContentsGallery">
${buttons}
<menu id="menuAgenda" label="More" imageMso="SlideHide" screentip="More agenda commands">${menu}</menu>
</group>`;
}

/** customUI/_rels/customUI14.xml.rels: one image relationship per icon. */
export function ribbonRels(images: RibbonImage[]): string {
  return `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${images.map((i) =>
    `<Relationship Id="${i.id}" Type="${IMAGE_REL}" Target="images/${i.file}"/>`).join('')}</Relationships>`;
}

/** The directions each chart type can grow in, for VBA: "stacked:up,down,left,right;line:up,left,right;…". */
export const directionTable = () => Object.entries(DIRECTIONS).map(([t, ds]) => `${t}:${ds.join(',')}`).join(';');

/** The module source: callback prefix filled in, and the reinstall command the error messages show. */
export function vbaSource(prefix: string, bas = readFileSync(BAS_FILE, 'utf8')): string {
  const reinstall = `"${reinstallCommand().replace(/"/g, '""')}"`;
  return bas.replace(PLACEHOLDER, prefix)
    .replace(/^(Private Const REINSTALL As String = ).*$/m, (_, decl) => decl + reinstall)
    .replace(/^(Private Const DIRECTIONS As String = ).*$/m, (_, decl) => `${decl}"${directionTable()}"`);
}

export interface BuiltAddin { ppam: Uint8Array; prefix: string; ribbon: string; source: string }

export function buildAddin(bas = readFileSync(BAS_FILE, 'utf8'), cls = readFileSync(CLS_FILE, 'utf8')): BuiltAddin {
  const items = galleryItems();
  const images = ribbonImages(items);
  const pngs = images.map((i) => ({ ...i, data: readFileSync(new URL(i.file, ICON_DIR)) }));
  const prefix = buildPrefix(ENGINE_VERSION, bas, cls, JSON.stringify(items), directionTable(), ...pngs.map((p) => p.data.toString('base64')));
  const ribbon = ribbonXml(prefix, items);
  const source = vbaSource(prefix, bas);
  const hex = createHash('sha1').update(prefix).digest('hex').toUpperCase();
  const id = `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20, 32)}}`;
  const vba = vbaProjectBin({ name: PROJECT_NAME, id, modules: [{ name: MODULE_NAME, source }, { name: 'AppEvents', source: cls, kind: 'class' }], seed: parseInt(hex.slice(0, 2), 16) });

  // Start from an empty presentation (master, layout, theme, no slides) and make it an add-in.
  const base = unzip(writePptx([]));
  const files: ZipInput[] = [];
  for (const [name, e] of base) {
    let text = e.text();
    if (name === '[Content_Types].xml') {
      text = text
        .replace('application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml', ADDIN_CT)
        .replace('<Default Extension="rels"', `<Default Extension="bin" ContentType="${VBA_CT}"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels"`);
    } else if (name === '_rels/.rels') {
      text = text.replace('</Relationships>', `<Relationship Id="rIdUI" Type="${UI_REL}" Target="customUI/customUI14.xml"/></Relationships>`);
    } else if (name === 'ppt/_rels/presentation.xml.rels') {
      text = text.replace('</Relationships>', `<Relationship Id="rIdVBA" Type="${VBA_REL}" Target="vbaProject.bin"/></Relationships>`);
    } else if (name === 'ppt/presentation.xml') {
      text = text.replace('<p:sldIdLst></p:sldIdLst>', '');
    }
    files.push({ name, data: text });
  }
  files.push({ name: 'ppt/vbaProject.bin', data: vba });
  files.push({ name: 'customUI/customUI14.xml', data: ribbon });
  files.push({ name: 'customUI/_rels/customUI14.xml.rels', data: ribbonRels(images) });
  for (const p of pngs) files.push({ name: `customUI/images/${p.file}`, data: p.data });
  return { ppam: zip(files), prefix, ribbon, source };
}
