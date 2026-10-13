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
import { DIRECTIONS, OPTIONS } from '../core/spec.ts';
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
${chartToolsGroup(cb)}
</tab>
</tabs></ribbon>
${contextMenusXml(cb)}
</customUI>
`;
}

/** A chart command: ribbon button in the Chart tools group, and the same item in right-click menus. */
export interface ChartTool { id: string; label: string; imageMso: string; tip: string; menu: boolean }

/** Chart tools (the reference product's right-click and Same-scale commands; core/commands.ts). */
export const CHART_TOOLS: ChartTool[] = [
  { id: 'btnTotals', label: 'Totals', imageMso: 'AutoSum', tip: 'Show or hide the totals above the columns of the selected chart.', menu: true },
  { id: 'btnCagr', label: 'CAGR arrow', imageMso: 'ChartTrendline', tip: 'Add a compound annual growth arrow to the selected chart, from the first to the last column, or between two columns you select (Shift-click their segments). Click again to remove it.', menu: true },
  { id: 'btnDifference', label: 'Difference arrow', imageMso: 'ShapeStraightConnectorArrow', tip: 'Add an arrow showing the difference between two column totals: the first and last column, or two columns you select. Click again to remove it.', menu: true },
  { id: 'btnSameScale', label: 'Same scale', imageMso: 'SizeToTallest', tip: 'Select two or more charts (Shift-click), then click: equal values are drawn equally tall in all of them, and stay so when you edit or resize one.', menu: true },
  { id: 'btnUnscale', label: 'Remove from same scale', imageMso: 'ObjectsUngroup', tip: 'The selected charts go back to their own scale.', menu: true },
  { id: 'btnPlainText', label: 'Convert to plain text', imageMso: 'TextBoxInsert', tip: 'Keep the selected label\'s text as it is now: data edits no longer change it. It still moves with the chart.', menu: true },
  { id: 'btnResetLabel', label: 'Reset label position', imageMso: 'PictureReset', tip: 'Put the selected label (or every label you moved, when the whole chart is selected) back where open-charts places it.', menu: true },
];

function chartToolsGroup(cb: (name: string) => string): string {
  const buttons = CHART_TOOLS.map((t) => `<button id="${t.id}" label="${esc(t.label)}" imageMso="${t.imageMso}" screentip="${esc(t.label)}" supertip="${esc(t.tip)}" onAction="${cb('onAction')}" getEnabled="${cb('getEnabled')}"/>`).join('\n');
  return `<group id="grpChartTools" label="Chart tools" imageMso="ChartInsert">\n${buttons}\n</group>`;
}

/**
 * PowerPoint's right-click menus that open on a chart (the group), one of its pieces (a shape, a
 * freeform, a line), a label being edited, or several charts. Each item's id is
 * ctx<menu>_<button id>: the dispatcher runs the button's command, and getVisible shows it only on
 * open-charts charts, when the command applies to what is selected.
 */
export const CONTEXT_MENUS = ['ContextMenuObjectsGroup', 'ContextMenuShape', 'ContextMenuShapeFreeform', 'ContextMenuShapeConnector', 'ContextMenuTextEdit', 'ContextMenuMultipleItems'];

export function contextMenusXml(cb: (name: string) => string): string {
  const items: { id: string; label: string; imageMso: string }[] = [
    { id: 'btnEditData', label: 'Edit data', imageMso: 'ChartEditDataSource' },
    { id: 'btnUpdate', label: 'Update chart', imageMso: 'RefreshAll' },
    ...CHART_TOOLS.filter((t) => t.menu),
  ];
  // No separator: one would show on every shape's menu (a separator has no getVisible).
  const menus = CONTEXT_MENUS.map((menu) => {
    const tag = menu.replace('ContextMenu', '');
    const buttons = items.map((t) => `<button id="ctx${tag}_${t.id}" label="${esc(t.label)}" imageMso="${t.imageMso}" onAction="${cb('onAction')}" getVisible="${cb('getVisible')}"/>`).join('');
    return `<contextMenu idMso="${menu}">${buttons}</contextMenu>`;
  }).join('\n');
  return `<contextMenus>\n${menus}\n</contextMenus>`;
}

/** The VBA's "|type|type|" constants for the commands that only some chart types take. */
const TYPE_CONSTS: Record<string, keyof typeof OPTIONS> = { TOTALS_TYPES: 'totals', ARROW_TYPES: 'cagr', SCALE_TYPES: 'scaleGroup' };

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
  let out = bas.replace(PLACEHOLDER, prefix)
    .replace(/^(Private Const REINSTALL As String = ).*$/m, (_, decl) => decl + reinstall)
    .replace(/^(Private Const DIRECTIONS As String = ).*$/m, (_, decl) => `${decl}"${directionTable()}"`);
  for (const [name, key] of Object.entries(TYPE_CONSTS)) {
    out = out.replace(new RegExp(`^(Private Const ${name} As String = ).*$`, 'm'), (_, decl) => `${decl}"|${OPTIONS[key].types.join('|')}|"`);
  }
  return out;
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
