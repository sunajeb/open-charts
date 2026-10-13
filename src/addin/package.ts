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
import { galleryItems, type ChartTypeEntry } from './registry.ts';
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

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Per-build callback prefix: derived from what is built, so identical inputs give identical bytes. */
export function buildPrefix(...parts: string[]): string {
  return `OC${createHash('sha1').update(parts.join('\0')).digest('hex').slice(0, 24).toUpperCase()}`;
}

export function ribbonXml(prefix: string, items: ChartTypeEntry[]): string {
  const cb = (name: string) => `${prefix}_${name}`;
  const families = [...new Set(items.map((g) => g.family ?? 'More'))];
  const menu = families.map((fam, i) => {
    const buttons = items.filter((g) => (g.family ?? 'More') === fam).map((g) =>
      `<button id="chart_${esc(g.id)}" label="${esc(g.label)}" imageMso="${esc(g.imageMso)}"${g.description ? ` description="${esc(g.description)}" supertip="${esc(g.description)}"` : ''} screentip="${esc(g.label)} chart" onAction="${cb('onAction')}"/>`).join('');
    return `<menuSeparator id="sepFamily${i}" title="${esc(fam)}"/>${buttons}`;
  }).join('');
  return `${XML}<customUI xmlns="${UI_NS}" onLoad="${cb('onLoad')}">
<ribbon><tabs>
<tab id="tabOpenCharts" label="open-charts" insertBeforeMso="TabHome">
<group id="grpCharts" label="Charts" imageMso="ChartInsert">
<menu id="menuCharts" label="Insert chart" imageMso="ChartInsert" size="large" itemSize="large" screentip="Insert chart" supertip="Insert a chart on this slide with sample data. It goes into an empty content placeholder if there is one. Then click Edit data to type your numbers.">${menu}</menu>
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

/** The module source: callback prefix filled in, and the reinstall command the error messages show. */
export function vbaSource(prefix: string, bas = readFileSync(BAS_FILE, 'utf8')): string {
  const reinstall = `"${reinstallCommand().replace(/"/g, '""')}"`;
  return bas.replace(PLACEHOLDER, prefix).replace(/^(Private Const REINSTALL As String = ).*$/m, (_, decl) => decl + reinstall);
}

export interface BuiltAddin { ppam: Uint8Array; prefix: string; ribbon: string; source: string }

export function buildAddin(bas = readFileSync(BAS_FILE, 'utf8'), cls = readFileSync(CLS_FILE, 'utf8')): BuiltAddin {
  const items = galleryItems();
  const prefix = buildPrefix(ENGINE_VERSION, bas, cls, JSON.stringify(items));
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
        .replace('<Default Extension="rels"', `<Default Extension="bin" ContentType="${VBA_CT}"/><Default Extension="rels"`);
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
  return { ppam: zip(files), prefix, ribbon, source };
}
