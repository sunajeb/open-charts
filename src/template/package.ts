// An editable view of an OOXML package for template fill: parts are read from the original zip and
// only the parts we change are re-serialised; every other part is written back as its original
// compressed bytes, so it comes out byte-identical (src/render/zip.ts raw entries).
import { unzip, type ZipEntry } from '../ooxml/unzip.ts';
import { readRels, relsPath, resolveTarget, type Entries, type Rel } from '../ooxml/pptx-read.ts';
import { escXml, findAll, parseXml } from '../ooxml/xml.ts';
import { REL, XML } from '../render/pptx.ts';
import { zip, type ZipInput } from '../render/zip.ts';

/** Relationship types whose target a cloned slide gets its own copy of (others are shared). */
const OWNED = ['/notesSlide', '/tags', '/chart', '/package', '/chartUserShapes', '/oleObject', 'chartStyle', 'chartColorStyle'];
/** Relationship types a cloned slide drops (comments belong to the original slide). */
const DROPPED = ['/comments', 'modernComment'];

const CT_PATH = '[Content_Types].xml';

export class Pkg {
  entries: Entries;
  /** Original part order; new parts are appended. */
  private order: string[];
  private changed = new Map<string, string | Uint8Array>();
  private removed = new Set<string>();
  private ctAdd = new Map<string, string>();
  private ctRemove = new Set<string>();
  private overrides: Map<string, string>;

  constructor(bytes: Uint8Array) {
    this.entries = unzip(bytes);
    this.order = [...this.entries.keys()];
    const ct = this.entries.get(CT_PATH);
    if (!ct) throw new Error('Not an OOXML package ([Content_Types].xml missing)');
    this.overrides = new Map(findAll(parseXml(ct.text()), 'Override').map((o) => [o.attrs.PartName.replace(/^\//, ''), o.attrs.ContentType]));
  }

  has(part: string) { return !this.removed.has(part) && (this.changed.has(part) || this.entries.has(part)); }

  text(part: string): string {
    const c = this.changed.get(part);
    if (c !== undefined) return typeof c === 'string' ? c : new TextDecoder().decode(c);
    const e = this.entries.get(part);
    if (!e || this.removed.has(part)) throw new Error(`Part ${part} is missing`);
    return e.text();
  }

  set(part: string, data: string | Uint8Array, contentType?: string) {
    if (!this.entries.has(part) && !this.changed.has(part)) this.order.push(part);
    this.changed.set(part, data);
    this.removed.delete(part);
    if (contentType && !this.overrides.has(part)) this.ctAdd.set(part, contentType);
  }

  remove(part: string) {
    this.removed.add(part);
    this.changed.delete(part);
    if (this.overrides.has(part)) this.ctRemove.add(part);
    this.ctAdd.delete(part);
  }

  /** Current relationships of a part (as edited). */
  rels(part: string): Map<string, Rel> {
    const p = relsPath(part);
    if (!this.has(p)) return new Map();
    return readRels(new Map([[p, { text: () => this.text(p) } as ZipEntry]]), part);
  }

  /** Writes the relationships of a part (targets relative to it). */
  setRels(part: string, rels: Rel[]) {
    const body = rels.map((r) => `<Relationship Id="${escXml(r.id)}" Type="${escXml(r.type)}" Target="${escXml(r.external ? r.target : relTarget(part, r.target))}"${r.external ? ' TargetMode="External"' : ''}/>`).join('');
    this.set(relsPath(part), `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`);
  }

  /** Adds a relationship and returns its id (rIdN above every existing one). */
  addRel(part: string, type: string, target: string): string {
    const rels = [...this.rels(part).values()];
    const n = Math.max(0, ...rels.map((r) => Number(/^rId(\d+)$/.exec(r.id)?.[1] ?? 0))) + 1;
    const id = `rId${n}`;
    this.setRels(part, [...rels, { id, type: type.startsWith('http') ? type : `${REL}/${type}`, target, external: false }]);
    return id;
  }

  contentType(part: string) { return this.overrides.get(part) ?? this.ctAdd.get(part); }

  /** A free part name like ppt/tags/tag7.xml (number above every existing part with that stem). */
  freeName(like: string): string {
    const m = /^(.*?)(\d*)(\.[^./]+)$/.exec(like)!;
    const re = new RegExp(`^${escRe(m[1])}(\\d+)${escRe(m[3])}$`);
    let n = 0;
    for (const p of [...this.order, ...this.changed.keys()]) { const k = re.exec(p); if (k) n = Math.max(n, Number(k[1])); }
    return `${m[1]}${n + 1}${m[3]}`;
  }

  /**
   * Copies a part for a new owner: same content, own copies of owned targets (notes, tags, charts),
   * shared targets (layout, images, links) and back-references rewired through `map`.
   */
  clone(part: string, map = new Map<string, string>()): string {
    const copy = this.freeName(part);
    map.set(part, copy);
    this.set(copy, this.entries.has(part) && !this.changed.has(part) ? this.entries.get(part)!.bytes() : this.text(part), this.contentType(part));
    const rels: Rel[] = [];
    for (const r of this.rels(part).values()) {
      if (r.external) { rels.push(r); continue; }
      if (DROPPED.some((t) => r.type.endsWith(t) || r.type.includes(t))) continue;
      const mapped = map.get(r.target);
      if (mapped) rels.push({ ...r, target: mapped });
      else if (OWNED.some((t) => r.type.endsWith(t) || r.type.includes(t)) && this.has(r.target)) rels.push({ ...r, target: this.clone(r.target, map) });
      else rels.push(r);
    }
    if (this.has(relsPath(part))) this.setRels(copy, rels);
    return copy;
  }

  /** Every part reachable from the package root through relationships (as edited). */
  reachable(): Set<string> {
    const seen = new Set<string>();
    const visit = (owner: string) => {
      for (const r of this.rels(owner).values()) {
        if (r.external || seen.has(r.target) || !this.has(r.target)) continue;
        seen.add(r.target);
        visit(r.target);
      }
    };
    visit('');
    return seen;
  }

  /** Removes parts that were reachable before the edits and no longer are (and their rels). */
  collect(before: Set<string>) {
    const now = this.reachable();
    for (const p of before) {
      if (now.has(p)) continue;
      this.remove(p);
      if (this.has(relsPath(p))) this.remove(relsPath(p));
    }
  }

  /** The package as a zip: untouched parts as their original compressed bytes. */
  toBytes(): Uint8Array {
    if (this.ctAdd.size || this.ctRemove.size) {
      let ct = this.text(CT_PATH);
      for (const p of this.ctRemove) ct = ct.replace(new RegExp(`<Override[^>]*PartName="/${escRe(escXml(p))}"[^>]*/>`), '');
      const add = [...this.ctAdd].map(([p, t]) => `<Override PartName="/${escXml(p)}" ContentType="${escXml(t)}"/>`).join('');
      ct = ct.replace('</Types>', `${add}</Types>`);
      this.changed.set(CT_PATH, ct);
    }
    const files: ZipInput[] = [];
    for (const name of this.order) {
      if (this.removed.has(name)) continue;
      const c = this.changed.get(name);
      if (c !== undefined) files.push({ name, data: c });
      else files.push({ name, raw: this.entries.get(name)!.raw });
    }
    return zip(files);
  }

  /** Names of the parts written with new content (for reports and tests). */
  changedParts(): string[] { return [...this.changed.keys()].filter((p) => !this.removed.has(p)); }
  removedParts(): string[] { return [...this.removed]; }
}

const escRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** Relationship target of `target` written from `from` (relative, like PowerPoint writes them). */
export function relTarget(from: string, target: string): string {
  const a = from.split('/').slice(0, -1), b = target.split('/');
  let i = 0;
  while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
  return [...Array(a.length - i).fill('..'), ...b.slice(i)].join('/');
}

export { resolveTarget };
