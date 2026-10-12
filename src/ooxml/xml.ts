// Small purpose-built XML parser for OOXML parts: elements, attributes, text — with source offsets,
// so callers can splice edits into the original string and leave everything else byte-identical.
// No namespace resolution (names keep their prefix, e.g. "p:sp"), no DTD, no entity expansion
// beyond the five predefined entities and numeric character references.

export interface XmlNode {
  name: string;
  attrs: Record<string, string>;
  children: XmlNode[];
  /** Decoded text directly inside this element (not descendants). */
  text: string;
  /** Offset of "<" of the start tag. */
  start: number;
  /** Offset just past the end tag (or the self-closing start tag). */
  end: number;
  /** Offset just past the start tag / at the start of the end tag (equal to end for self-closing). */
  innerStart: number;
  innerEnd: number;
  parent?: XmlNode;
}

const ENT: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

export function decodeEntities(s: string): string {
  if (!s.includes('&')) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-z]+);/g, (m, e: string) => {
    if (e[0] === '#') {
      const cp = e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return cp <= 0x10ffff ? String.fromCodePoint(cp) : m;
    }
    return ENT[e] ?? m;
  });
}

export const escXml = (s: string) =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const ATTR = /([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

/** Parses a document and returns a synthetic root whose children are the top-level elements. */
export function parseXml(src: string): XmlNode {
  const root: XmlNode = { name: '#document', attrs: {}, children: [], text: '', start: 0, end: src.length, innerStart: 0, innerEnd: src.length };
  let cur = root;
  let i = 0;
  while (i < src.length) {
    const lt = src.indexOf('<', i);
    if (lt < 0) { cur.text += decodeEntities(src.slice(i)); break; }
    if (lt > i) cur.text += decodeEntities(src.slice(i, lt));
    if (src.startsWith('<?', lt)) { i = need(src, '?>', lt) + 2; continue; }
    if (src.startsWith('<!--', lt)) { i = need(src, '-->', lt) + 3; continue; }
    if (src.startsWith('<![CDATA[', lt)) {
      const e = need(src, ']]>', lt);
      cur.text += src.slice(lt + 9, e);
      i = e + 3;
      continue;
    }
    if (src.startsWith('<!', lt)) { // DOCTYPE: skipped, never expanded
      let depth = 0, j = lt + 2;
      for (; j < src.length; j++) {
        if (src[j] === '[') depth++;
        else if (src[j] === ']') depth--;
        else if (src[j] === '>' && depth <= 0) break;
      }
      i = j + 1;
      continue;
    }
    const gt = tagEnd(src, lt);
    if (src[lt + 1] === '/') {
      const name = src.slice(lt + 2, gt).trim();
      if (cur === root || cur.name !== name) throw new Error(`XML: unexpected </${name}> at ${lt}`);
      cur.innerEnd = lt;
      cur.end = gt + 1;
      cur = cur.parent!;
      i = gt + 1;
      continue;
    }
    const selfClosing = src[gt - 1] === '/';
    const body = src.slice(lt + 1, selfClosing ? gt - 1 : gt);
    const m = /^[^\s/>]+/.exec(body);
    if (!m) throw new Error(`XML: bad tag at ${lt}`);
    const attrs: Record<string, string> = {};
    ATTR.lastIndex = m[0].length;
    for (let a; (a = ATTR.exec(body));) attrs[a[1]] = decodeEntities(a[2] ?? a[3]);
    const node: XmlNode = { name: m[0], attrs, children: [], text: '', start: lt, end: gt + 1, innerStart: gt + 1, innerEnd: gt + 1, parent: cur };
    cur.children.push(node);
    if (!selfClosing) cur = node;
    i = gt + 1;
  }
  if (cur !== root) throw new Error(`XML: unclosed <${cur.name}>`);
  return root;
}

function need(src: string, s: string, from: number): number {
  const k = src.indexOf(s, from);
  if (k < 0) throw new Error(`XML: unterminated construct at ${from}`);
  return k;
}

/** Index of the ">" closing the tag that starts at lt, skipping quoted attribute values. */
function tagEnd(src: string, lt: number): number {
  let q = '';
  for (let j = lt + 1; j < src.length; j++) {
    const ch = src[j];
    if (q) { if (ch === q) q = ''; }
    else if (ch === '"' || ch === "'") q = ch;
    else if (ch === '>') return j;
  }
  throw new Error(`XML: unterminated tag at ${lt}`);
}

export const child = (n: XmlNode | undefined, name: string) => n?.children.find((c) => c.name === name);
export const kids = (n: XmlNode | undefined, name: string) => n?.children.filter((c) => c.name === name) ?? [];

/** Follows a path of direct-child names, e.g. path(sp, 'p:nvSpPr', 'p:cNvPr'). */
export function path(n: XmlNode | undefined, ...names: string[]): XmlNode | undefined {
  for (const name of names) n = child(n, name);
  return n;
}

/** All descendants (depth-first, document order) with the given name. */
export function findAll(n: XmlNode, name: string, out: XmlNode[] = []): XmlNode[] {
  for (const c of n.children) {
    if (c.name === name) out.push(c);
    findAll(c, name, out);
  }
  return out;
}

export const find = (n: XmlNode, name: string) => findAll(n, name)[0];

/** Concatenated text of a node and all its descendants. */
export function textOf(n: XmlNode): string {
  return n.text + n.children.map(textOf).join('');
}
