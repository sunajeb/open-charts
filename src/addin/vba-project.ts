// Writes ppt/vbaProject.bin (MS-OVBA) from plain VBA source — no VBE needed.
// Like EPPlus and the reference add-in, the module streams hold only compressed source
// (MODULEOFFSET 0, no p-code) and _VBA_PROJECT is the 7-byte version-independent stub, so
// Office compiles the project the first time it loads it.
import { writeCfb } from './cfb.ts';
import { compress } from './ovba.ts';

export interface VbaModule {
  name: string;
  source: string;
  /** "class" for a class module (e.g. one holding a WithEvents Application). Default: standard module. */
  kind?: 'standard' | 'class';
}

export interface VbaProjectOptions {
  name: string;
  modules: VbaModule[];
  /** Project GUID, "{XXXXXXXX-...}". */
  id: string;
  /** Seed for the CMG/DPB/GC obfuscation bytes, for reproducible builds. */
  seed?: number;
}

const CODEPAGE = 1252;

/** Code page 1252 encoding; the module source is kept ASCII so this is exact. */
export function encode1252(s: string): Uint8Array {
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c > 0x7e && c !== 0x09) throw new Error(`VBA source must be ASCII (found U+${c.toString(16)} at ${i})`);
    out[i] = c;
  }
  return out;
}

const utf16 = (s: string) => {
  const out = new Uint8Array(s.length * 2);
  for (let i = 0; i < s.length; i++) { out[i * 2] = s.charCodeAt(i) & 0xff; out[i * 2 + 1] = s.charCodeAt(i) >> 8; }
  return out;
};

class Rec {
  bytes: number[] = [];
  u16(v: number) { this.bytes.push(v & 0xff, (v >> 8) & 0xff); return this; }
  u32(v: number) { this.bytes.push(v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff, (v >>> 24) & 0xff); return this; }
  raw(b: Uint8Array) { this.bytes.push(...b); return this; }
  /** id, 4-byte size, payload */
  rec(id: number, payload: Uint8Array) { return this.u16(id).u32(payload.length).raw(payload); }
  u32rec(id: number, v: number) { return this.u16(id).u32(4).u32(v); }
}

/** The decompressed `dir` stream (MS-OVBA 2.3.4.2): project information, no references, modules. */
export function dirStream(name: string, modules: VbaModule[]): Uint8Array {
  const r = new Rec();
  const a = (s: string) => encode1252(s);
  r.u32rec(0x0001, 3); // SysKind: 64-bit (what current Office writes; Office rewrites it on save)
  r.u32rec(0x0002, 0x409); // Lcid
  r.u32rec(0x0014, 0x409); // LcidInvoke
  r.u16(0x0003).u32(2).u16(CODEPAGE);
  r.rec(0x0004, a(name));
  r.rec(0x0005, a('')).rec(0x0040, utf16('')); // DocString + unicode
  r.rec(0x0006, a('')).rec(0x003d, a('')); // HelpFile 1 and 2
  r.u32rec(0x0007, 0); // HelpContext
  r.u32rec(0x0008, 0); // LibFlags
  r.u16(0x0009).u32(4).u32(1).u16(0); // Version: major 1, minor 0
  r.rec(0x000c, a('')).rec(0x003c, utf16('')); // Constants + unicode
  r.u16(0x000f).u32(2).u16(modules.length); // PROJECTMODULES
  r.u16(0x0013).u32(2).u16(0xffff); // PROJECTCOOKIE
  for (const m of modules) {
    r.rec(0x0019, a(m.name)).rec(0x0047, utf16(m.name)); // MODULENAME + unicode
    r.rec(0x001a, a(m.name)).rec(0x0032, utf16(m.name)); // MODULESTREAMNAME + unicode
    r.rec(0x001c, a('')).rec(0x0048, utf16('')); // MODULEDOCSTRING + unicode
    r.u32rec(0x0031, 0); // MODULEOFFSET: source starts at 0, no p-code
    r.u32rec(0x001e, 0); // MODULEHELPCONTEXT
    r.u16(0x002c).u32(2).u16(0xffff); // MODULECOOKIE
    r.u16(m.kind === 'class' ? 0x0022 : 0x0021).u32(0); // MODULETYPE: procedural / document-or-class
    r.u16(0x002b).u32(0); // terminator
  }
  r.u16(0x0010).u32(0); // dir terminator
  return Uint8Array.from(r.bytes);
}

/** MS-OVBA 2.4.3 data "encryption" (obfuscation) used for CMG, DPB and GC in the PROJECT stream. */
export function encryptData(projectId: string, data: Uint8Array, seed: number): string {
  let projKey = 0;
  for (const c of projectId) projKey = (projKey + c.charCodeAt(0)) & 0xff;
  const version = 2;
  seed &= 0xff;
  const out = [seed, seed ^ version, seed ^ projKey];
  let unencrypted1 = projKey, encrypted1 = seed ^ projKey, encrypted2 = seed ^ version;
  const push = (b: number) => {
    const enc = b ^ ((encrypted2 + unencrypted1) & 0xff);
    out.push(enc);
    encrypted2 = encrypted1; encrypted1 = enc; unencrypted1 = b;
  };
  const ignored = (seed & 6) / 2;
  for (let i = 0; i < ignored; i++) push(0x07);
  const len = data.length;
  for (const b of [len & 0xff, (len >> 8) & 0xff, (len >> 16) & 0xff, (len >>> 24) & 0xff]) push(b);
  for (const b of data) push(b);
  return out.map((b) => b.toString(16).toUpperCase().padStart(2, '0')).join('');
}

/** Inverse of encryptData; returns the payload. */
export function decryptData(hex: string): Uint8Array {
  const b = hex.match(/../g)!.map((x) => parseInt(x, 16));
  const seed = b[0], projKey = seed ^ b[2];
  let unencrypted1 = projKey, encrypted1 = b[2], encrypted2 = b[1];
  const out: number[] = [];
  for (let i = 3; i < b.length; i++) {
    const dec = b[i] ^ ((encrypted2 + unencrypted1) & 0xff);
    out.push(dec);
    encrypted2 = encrypted1; encrypted1 = b[i]; unencrypted1 = dec;
  }
  const skip = (seed & 6) / 2;
  const len = out[skip] | (out[skip + 1] << 8) | (out[skip + 2] << 16) | (out[skip + 3] << 24);
  return Uint8Array.from(out.slice(skip + 4, skip + 4 + len));
}

/** The PROJECT stream: unprotected, visible project. */
export function projectStream(opts: VbaProjectOptions): string {
  const seed = opts.seed ?? 0x5a;
  const lines = [
    `ID="${opts.id}"`,
    ...opts.modules.map((m) => (m.kind === 'class' ? `Class=${m.name}` : `Module=${m.name}`)),
    `Name="${opts.name}"`,
    'HelpContextID="0"',
    `CMG="${encryptData(opts.id, new Uint8Array(4), seed)}"`, // protection state: none
    `DPB="${encryptData(opts.id, new Uint8Array(1), seed + 1)}"`, // password: none
    `GC="${encryptData(opts.id, Uint8Array.of(0xff), seed + 2)}"`, // visibility: visible
    '',
    '[Host Extender Info]',
    '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000',
    '',
  ];
  return lines.join('\r\n') + '\r\n';
}

const CLASS_ATTRIBUTES = [
  'Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"',
  'Attribute VB_GlobalNameSpace = False',
  'Attribute VB_Creatable = False',
  'Attribute VB_PredeclaredId = False',
  'Attribute VB_Exposed = False',
  'Attribute VB_TemplateDerived = False',
  'Attribute VB_Customizable = False',
];

/** Source as VBA stores it: CRLF line endings, `Attribute VB_Name` (and a class's attributes) first. */
export function moduleSource(m: VbaModule): string {
  const body = m.source.replace(/\r\n?/g, '\n').replace(/^(Attribute VB_\w+ = .*\n)+/, '');
  const head = [`Attribute VB_Name = "${m.name}"`, ...(m.kind === 'class' ? CLASS_ATTRIBUTES : [])].join('\n');
  return `${head}\n${body}`.replace(/\n/g, '\r\n');
}

export function vbaProjectBin(opts: VbaProjectOptions): Uint8Array {
  // VBA will not open a project whose name is also a module's name ("couldn't open VBA").
  const clash = opts.modules.find((m) => m.name.toLowerCase() === opts.name.toLowerCase());
  if (clash) throw new Error(`VBA project name "${opts.name}" must differ from module "${clash.name}"`);
  const vba: Record<string, Uint8Array> = {
    _VBA_PROJECT: Uint8Array.of(0xcc, 0x61, 0xff, 0xff, 0x00, 0x00, 0x00),
    dir: compress(dirStream(opts.name, opts.modules)),
  };
  for (const m of opts.modules) vba[m.name] = compress(encode1252(moduleSource(m)));
  return writeCfb({ VBA: vba, PROJECT: encode1252(projectStream(opts)) });
}
