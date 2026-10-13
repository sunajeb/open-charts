// Path A on the user's own slide (PowerPointApi 1.8): the host exports the slide as a one-slide
// .pptx (Slide.exportAsBase64), this puts the chart group into that slide's XML — in place of the
// old chart group, keeping its z-order, or on top — and the host inserts the result after the
// original and deletes the original. Every part we do not touch is copied byte for byte.
// Pure apart from `inflate`: the browser passes DecompressionStream, Node tests pass node:zlib.
import { readZip, type ZipEntry } from '../ooxml/zipread.ts';
import { findAll, kids, parseXml, path, type XmlNode } from '../ooxml/xml.ts';
import { groupXml, Ids, tagPartXml, TAGS_CT, TAGS_REL, type SceneContent } from '../render/pptx.ts';
import { zip, type ZipInput } from '../render/zip.ts';

export type InflateAsync = (compressed: Uint8Array, size: number) => Promise<Uint8Array>;

/** Raw-deflate decompression with the browser's DecompressionStream. */
export const browserInflate: InflateAsync = async (compressed) => {
  const ds = new (globalThis as unknown as { DecompressionStream: new (f: string) => TransformStream<Uint8Array, Uint8Array> }).DecompressionStream('deflate-raw');
  const buf = await new Response(new Blob([compressed as BlobPart]).stream().pipeThrough(ds)).arrayBuffer();
  return new Uint8Array(buf);
};

export interface PatchOptions {
  /** cNvPr name of the group to replace (the host renames the old chart to this marker before exporting). */
  replaceName?: string;
  font: string;
}

/**
 * Returns the patched package. Throws when the export has no slide, or `replaceName` names no group
 * at the top level of the slide's shape tree.
 */
export async function patchSlide(pkg: Uint8Array, content: SceneContent, opts: PatchOptions, inflate: InflateAsync): Promise<Uint8Array> {
  const inflated = new Map<Uint8Array, Uint8Array>();
  const entries = readZip(pkg, (c) => {
    const out = inflated.get(c);
    if (!out) throw new Error('internal: part read before it was inflated');
    return out;
  });
  const text = async (e: ZipEntry) => {
    if (e.raw.method === 8 && !inflated.has(e.raw.compressed)) inflated.set(e.raw.compressed, await inflate(e.raw.compressed, e.raw.size));
    return e.text();
  };

  const slideName = [...entries.keys()].filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort()[0];
  if (!slideName) throw new Error('The exported slide has no slide part.');
  const relsName = slideName.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels';
  const slide = await text(entries.get(slideName)!);
  const rels = entries.get(relsName) ? await text(entries.get(relsName)!) : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  const types = await text(entries.get('[Content_Types].xml')!);

  const doc = parseXml(slide);
  const tree = path(doc.children.find((c) => c.name === 'p:sld'), 'p:cSld', 'p:spTree');
  if (!tree) throw new Error('The exported slide has no shape tree.');
  // Ids above every shape id on the slide; relationship ids above every rId; tag parts above every tagN.
  const maxId = Math.max(1, ...findAll(tree, 'p:cNvPr').map((n) => Number(n.attrs.id) || 0));
  const relIds = [...rels.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]));
  let nextRel = Math.max(0, ...relIds);
  let nextTag = Math.max(0, ...[...entries.keys()].map((n) => Number(/^ppt\/tags\/tag(\d+)\.xml$/.exec(n)?.[1] ?? 0)));
  const newParts: ZipInput[] = [];
  const newRels: string[] = [];
  const newTypes: string[] = [];
  const addTags = (tags: [string, string][]) => {
    const part = `tag${++nextTag}.xml`;
    newParts.push({ name: `ppt/tags/${part}`, data: tagPartXml(tags) });
    newTypes.push(`<Override PartName="/ppt/tags/${part}" ContentType="${TAGS_CT}"/>`);
    const id = `rId${++nextRel}`;
    newRels.push(`<Relationship Id="${id}" Type="${TAGS_REL}" Target="../tags/${part}"/>`);
    return id;
  };
  const group = groupXml(content, new Ids(maxId), opts.font, addTags);

  let patched: string;
  if (opts.replaceName !== undefined) {
    const old = kids(tree, 'p:grpSp').find((g: XmlNode) => path(g, 'p:nvGrpSpPr', 'p:cNvPr')?.attrs.name === opts.replaceName);
    if (!old) throw new Error('The chart to replace was not found on the exported slide.');
    patched = slide.slice(0, old.start) + group + slide.slice(old.end);
  } else {
    patched = slide.slice(0, tree.innerEnd) + group + slide.slice(tree.innerEnd);
  }
  const at = rels.lastIndexOf('</Relationships>');
  const patchedRels = rels.slice(0, at) + newRels.join('') + rels.slice(at);
  const typesAt = types.lastIndexOf('</Types>');
  const patchedTypes = types.slice(0, typesAt) + newTypes.join('') + types.slice(typesAt);

  const out: ZipInput[] = [];
  const changed = new Map([[slideName, patched], [relsName, patchedRels], ['[Content_Types].xml', patchedTypes]]);
  for (const [name, e] of entries) out.push(changed.has(name) ? { name, data: changed.get(name)! } : { name, raw: e.raw });
  if (!entries.has(relsName)) out.push({ name: relsName, data: patchedRels });
  out.push(...newParts);
  return zip(out);
}
