import { test } from 'node:test';
import assert from 'node:assert/strict';
import { compress, decompress } from '../src/addin/ovba.ts';
import { compareNames, readCfb, writeCfb, type CfbStorage } from '../src/addin/cfb.ts';
import { decryptData, dirStream, encryptData, moduleSource, vbaProjectBin } from '../src/addin/vba-project.ts';

const enc = (s: string) => new TextEncoder().encode(s);

test('OVBA compression round-trips text, repeats, binary and chunk boundaries', () => {
  const cases = [
    '', 'a', 'abc', 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa',
    'Sub Foo()\r\n  MsgBox "hi"\r\nEnd Sub\r\n'.repeat(500),
    'x'.repeat(4096), 'x'.repeat(4097), 'abcdefgh'.repeat(1024 + 3),
  ].map(enc);
  let seed = 7;
  cases.push(Uint8Array.from({ length: 9000 }, () => (seed = (seed * 1103515245 + 12345) >>> 0) >>> 24));
  for (const c of cases) {
    const z = compress(c);
    assert.equal(z[0], 1);
    assert.deepEqual(decompress(z), c);
  }
  assert.ok(compress(cases[4]).length < cases[4].length / 5, 'repetitive source compresses well');
});

test('OVBA chunk headers: signature 0b011, compressed flag, size - 3', () => {
  const z = compress(enc('x'.repeat(5000)));
  const h1 = z[1] | (z[2] << 8);
  assert.equal((h1 >> 12) & 7, 3);
  assert.equal(h1 >> 15, 1);
  const next = 1 + (h1 & 0xfff) + 3;
  assert.equal(((z[next] | (z[next + 1] << 8)) >> 12) & 7, 3, 'second chunk header follows the first');
});

test('CFB writer round-trips nested storages, mini and regular streams', () => {
  const big = Uint8Array.from({ length: 70_000 }, (_, i) => i % 251);
  const tree: CfbStorage = {
    VBA: { _VBA_PROJECT: Uint8Array.of(1, 2, 3), dir: enc('dir'), Module1: big, Empty: new Uint8Array(0) },
    PROJECT: enc('ID="{x}"'),
    PROJECTwm: Uint8Array.from({ length: 4096 }, (_, i) => i & 0xff),
  };
  const back = readCfb(writeCfb(tree));
  assert.deepEqual(back, tree);
});

test('CFB sibling order is length first, then case-insensitive', () => {
  assert.deepEqual(['dir', 'PROJECT', 'VBA', '_VBA_PROJECT', 'abc'].sort(compareNames), ['abc', 'dir', 'VBA', 'PROJECT', '_VBA_PROJECT']);
});

test('PROJECT obfuscation matches a reference CMG and round-trips', () => {
  // CMG from a real PowerPoint add-in: seed 0x0E, 4-byte protection state for project ID {4A203AE8-...}.
  const id = '{4A203AE8-4D47-46D1-91E8-A13D17872D25}';
  assert.deepEqual(decryptData('0E0CB4C1BCCFC7D3C7D3C3D7C3D7'), Uint8Array.of(4, 0, 0, 0)); // its protection state
  assert.equal(encryptData(id, new Uint8Array(4), 0x0e).slice(0, 6), '0E0CB4');
  for (const seed of [0, 1, 2, 6, 0x5a, 0xff]) assert.deepEqual(decryptData(encryptData(id, Uint8Array.of(0xff), seed)), Uint8Array.of(0xff));
});

test('vbaProject.bin holds PROJECT, VBA/_VBA_PROJECT, VBA/dir and the compressed module source', () => {
  const source = 'Option Explicit\nSub Hello()\n  MsgBox "hi"\nEnd Sub\n';
  assert.throws(() => vbaProjectBin({ name: 'OpenCharts', id: '{x}', modules: [{ name: 'openCharts', source }] }), /must differ/);
  const bin = vbaProjectBin({ name: 'OpenChartsAddIn', id: '{11111111-2222-3333-4444-555555555555}', modules: [{ name: 'OpenCharts', source }] });
  const cfb = readCfb(bin) as { PROJECT: Uint8Array; VBA: Record<string, Uint8Array> };
  assert.match(new TextDecoder().decode(cfb.PROJECT), /^ID="\{11111111-2222-3333-4444-555555555555\}"\r\nModule=OpenCharts\r\nName="OpenChartsAddIn"\r\n/);
  assert.deepEqual(cfb.VBA._VBA_PROJECT, Uint8Array.of(0xcc, 0x61, 0xff, 0xff, 0, 0, 0));
  assert.deepEqual(decompress(cfb.VBA.dir), dirStream('OpenChartsAddIn', [{ name: 'OpenCharts', source }]));
  const text = new TextDecoder().decode(decompress(cfb.VBA.OpenCharts));
  assert.equal(text, moduleSource({ name: 'OpenCharts', source }));
  assert.ok(text.startsWith('Attribute VB_Name = "OpenCharts"\r\nOption Explicit\r\n'));
  assert.throws(() => vbaProjectBin({ name: 'P', id: '{x}', modules: [{ name: 'M', source: 'MsgBox "é"' }] }), /ASCII/);
});
