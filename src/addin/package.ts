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

export const MODULE_NAME = 'OpenCharts';
// Must differ from every module name, or VBA refuses to open the project.
export const PROJECT_NAME = 'OpenChartsAddIn';
export const BAS_FILE = new URL('../../addin/vba/OpenCharts.bas', import.meta.url);
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
  const gallery = items.map((g) =>
    `<item id="chart_${esc(g.id)}" label="${esc(g.label)}" imageMso="${esc(g.imageMso)}" screentip="${esc(g.label)}"/>`).join('');
  return `${XML}<customUI xmlns="${UI_NS}" onLoad="${cb('onLoad')}">
<ribbon><tabs>
<tab id="tabOpenCharts" label="open-charts" insertBeforeMso="TabHome">
<group id="grpCharts" label="Charts" imageMso="ChartInsert">
<gallery id="galCharts" label="Insert chart" imageMso="ChartInsert" size="large" columns="4" showItemLabel="true" showItemImage="true" screentip="Insert chart" supertip="Insert an open-charts chart on this slide with sample data. Select it and click Edit data to change the data." onAction="${cb('onGallery')}">${gallery}</gallery>
<button id="btnEditData" label="Edit data" imageMso="ChartEditDataSource" size="large" screentip="Edit data" supertip="Open the selected chart's data in the open-charts editor. Click Update chart afterwards." onAction="${cb('onAction')}"/>
<button id="btnUpdate" label="Update chart" imageMso="RefreshAll" size="large" screentip="Update chart" supertip="Lay the selected chart out again in its current position and size, with any edited data." onAction="${cb('onAction')}"/>
</group>
</tab>
</tabs></ribbon>
</customUI>
`;
}

export function vbaSource(prefix: string, bas = readFileSync(BAS_FILE, 'utf8')): string {
  return bas.replace(PLACEHOLDER, prefix);
}

export interface BuiltAddin { ppam: Uint8Array; prefix: string; ribbon: string; source: string }

export function buildAddin(bas = readFileSync(BAS_FILE, 'utf8')): BuiltAddin {
  const items = galleryItems();
  const prefix = buildPrefix(ENGINE_VERSION, bas, JSON.stringify(items));
  const ribbon = ribbonXml(prefix, items);
  const source = vbaSource(prefix, bas);
  const hex = createHash('sha1').update(prefix).digest('hex').toUpperCase();
  const id = `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20, 32)}}`;
  const vba = vbaProjectBin({ name: PROJECT_NAME, id, modules: [{ name: MODULE_NAME, source }], seed: parseInt(hex.slice(0, 2), 16) });

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
