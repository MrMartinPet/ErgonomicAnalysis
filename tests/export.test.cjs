// Run with: node tests/export.test.cjs (Node 18+, no installation required).
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.resolve(__dirname, '..');
const ctx = { console, setTimeout, clearTimeout, ArrayBuffer, Buffer, Uint8Array, DataView, Blob, TextEncoder, TextDecoder, structuredClone,
  document: { createElement: () => ({ getContext: () => ({ font: '', measureText: text => ({ width: text.length * 12 }) }) }), getElementById: () => ({}) }
};
ctx.window = ctx;
vm.createContext(ctx);
for (const filename of ['vendor/pptxgen.bundle.js', 'ole-package.js', 'export-pptx.js']) vm.runInContext(fs.readFileSync(path.join(root, filename), 'utf8'), ctx);

// Independent MS-CFB reader: verify payload bytes across mini-FAT, FAT and DIFAT.
function olePayload(input) {
  const b = Buffer.from(input);
  assert.equal(b.subarray(0, 8).toString('hex'), 'd0cf11e0a1b11ae1');
  const fatIds = [];
  for (let at = 76; at < 512; at += 4) if (b.readUInt32LE(at) < 0xfffffffa) fatIds.push(b.readUInt32LE(at));
  let dif = b.readUInt32LE(68);
  for (let count = b.readUInt32LE(72); count; count--) {
    const sector = b.subarray((dif + 1) * 512, (dif + 2) * 512);
    for (let at = 0; at < 508; at += 4) if (sector.readUInt32LE(at) < 0xfffffffa) fatIds.push(sector.readUInt32LE(at));
    dif = sector.readUInt32LE(508);
  }
  const fat = fatIds.flatMap(id => Array.from({ length: 128 }, (_, i) => b.readUInt32LE((id + 1) * 512 + i * 4)));
  function chain(start, table, bytes, size, offset) {
    const chunks = [], seen = new Set();
    while (start < 0xfffffffa) {
      assert(!seen.has(start), 'FAT chain must not loop');
      seen.add(start);
      chunks.push(bytes.subarray(offset + start * size, offset + (start + 1) * size));
      start = table[start];
    }
    return Buffer.concat(chunks);
  }
  const directory = chain(b.readUInt32LE(48), fat, b, 512, 512);
  const entries = {};
  for (let at = 0; at < directory.length; at += 128) {
    const length = directory.readUInt16LE(at + 64);
    if (length) entries[directory.subarray(at, at + length - 2).toString('utf16le')] = { start: directory.readUInt32LE(at + 116), size: directory.readUInt32LE(at + 120) };
  }
  const native = entries['\x01Ole10Native'];
  let data;
  if (native.size >= 4096) data = chain(native.start, fat, b, 512, 512).subarray(0, native.size);
  else {
    const rootEntry = entries['Root Entry'];
    const minis = chain(rootEntry.start, fat, b, 512, 512).subarray(0, rootEntry.size);
    const miniFatBytes = chain(b.readUInt32LE(60), fat, b, 512, 512);
    const miniFat = Array.from({ length: miniFatBytes.length / 4 }, (_, i) => miniFatBytes.readUInt32LE(i * 4));
    data = chain(native.start, miniFat, minis, 64, 0).subarray(0, native.size);
  }
  assert.equal(data.readUInt32LE(0), data.length - 4);
  let at = 6;
  const skipString = () => { at = data.indexOf(0, at) + 1; };
  skipString(); skipString(); at += 8; skipString();
  const length = data.readUInt32LE(at); at += 4;
  return data.subarray(at, at + length);
}

(async () => {
  for (const length of [0, 15, 4000, 4096, 20 * 1024 * 1024]) {
    const source = Uint8Array.from({ length }, (_, i) => i % 251);
    const packed = ctx.KimOlePackage('Instruktion_åäö.txt', source);
    assert.deepEqual(olePayload(packed), Buffer.from(source));
    if (length > 10 * 1024 * 1024) assert(Buffer.from(packed).readUInt32LE(72) > 0, 'Large files exercise DIFAT');
  }
  const photo = 'data:image/png;base64,' + fs.readFileSync(path.join(root, 'posture_1.png')).toString('base64');
  const sections = [1, 2, 3, 4, 5, 8].map(n => ({ number: n, title: 'Fråga ' + n, score: n === 1 || n === 8 ? null : 3.5, choice: 'Egen poäng', explanation: 'Eget mellanvärde', comment: `MOTIVERING_${n}\n` + (n === 5 ? 'Lång observation. '.repeat(600) + 'SLUT_MARKOR' : 'Observerat vid arbetsstationen.') }));
  const report = { metadata: ['60', 'Formning', 'Station 1'], sections, scores: ['3.5', '1', '2', '6.5', '6', '39'], range: '25 till 49 poäng', result: 'Kraftigt ökad belastningssituation.', legend: ['0 till 9', '10 till 24', '25 till 49', '50+'] };
  const evidence = { sections: Object.fromEntries([1, 2, 3, 4, 5, 8].map(n => [n, { images: [{ name: 'foto.png', caption: 'BILDTEXT_' + n, width: 500, height: 359, data: photo }] }])), attachments: [{ name: 'Arbetsinstruktion.txt', description: 'Gällande instruktion.', bytes: new TextEncoder().encode('Originaldokumentets innehåll: åäö.') }] };
  const zip = await ctx.JSZip.loadAsync(await ctx.KimPowerPoint.build(report, evidence));
  const slides = Object.keys(zip.files).filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name)).sort((a, b) => Number(a.match(/(\d+)\.xml/)[1]) - Number(b.match(/(\d+)\.xml/)[1]));
  const xml = await Promise.all(slides.map(name => zip.file(name).async('string')));
  const all = xml.join('\n');
  for (const n of [1, 2, 3, 4, 5, 8]) { assert(all.includes('MOTIVERING_' + n)); assert(all.includes('BILDTEXT_' + n)); }
  assert(all.includes('SLUT_MARKOR'), 'Long comments must never be truncated');
  assert(xml.at(-1).includes('Sammanfattning'), 'The summary remains the final slide');
  assert(all.includes('<a:tbl>'), 'Scores must remain native editable tables');
  assert.equal((all.match(/<p:oleObj /g) || []).length, 1);
  assert.equal((all.match(/<p:pic>/g) || []).length, 7, 'Six separate photos and the attachment picture');
  assert.deepEqual(olePayload(await zip.file('ppt/embeddings/oleObject1.bin').async('uint8array')), Buffer.from(evidence.attachments[0].bytes));
  const withoutFiles = await ctx.JSZip.loadAsync(await ctx.KimPowerPoint.build(report, { sections: {}, attachments: [] }));
  assert(!Object.keys(withoutFiles.files).some(name => /^ppt\/embeddings\/.+\.bin$/.test(name)));
  console.log('PASS: editable slides, six photo sections, complete long comments, final summary, optional attachments and byte-perfect OLE files up to 20 MB.');
})().catch(error => { console.error(error); process.exitCode = 1; });
