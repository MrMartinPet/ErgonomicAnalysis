// Run with: node tests/export.test.cjs (Node 18+, no installation required).
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.resolve(__dirname, '..');
const ctx = { console, setTimeout, clearTimeout, ArrayBuffer, Buffer, Uint8Array, DataView, Blob, TextEncoder, TextDecoder, structuredClone,
  document: { createElement: () => ({ getContext: () => ({ font: '', measureText(text) { return { width: text.length * (parseFloat(this.font) || 16) * .52 }; } }) }), getElementById: () => ({}) }
};
ctx.window = ctx;
vm.createContext(ctx);
for (const filename of ['vendor/pptxgen.bundle.js', 'ole-package.js', 'report-model.js', 'pptx-layout.js', 'export-pptx.js', 'export-excel.js']) vm.runInContext(fs.readFileSync(path.join(root, filename), 'utf8'), ctx);

async function validatePowerPointPackage(zip) {
  for (const name of Object.keys(zip.files).filter(n=>/^ppt\/(slides|slideMasters|slideLayouts)\/[^/]+\.xml$/.test(n))) {
    const xml=await zip.file(name).async('string');
    const ids=Array.from(xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g),m=>m[1]);
    assert.equal(new Set(ids).size,ids.length,`${name}: duplicate drawing object IDs make the PowerPoint file non-conformant`);
  }
  const types=await zip.file('[Content_Types].xml').async('string');
  for (const match of types.matchAll(/<Override\b[^>]*\bPartName="([^"]+)"/g)) assert(zip.file(match[1].replace(/^\//,'')),`Missing declared package part: ${match[1]}`);
  for (const name of Object.keys(zip.files).filter(n=>n.endsWith('.rels'))) {
    const base=name==='_rels/.rels'?'':name.slice(0,name.lastIndexOf('/_rels/'));
    const rels=await zip.file(name).async('string');
    for(const match of rels.matchAll(/<Relationship\b[^>]*\/>/g)) {
      if(/TargetMode="External"/.test(match[0]))continue;
      const target=/\bTarget="([^"]+)"/.exec(match[0])[1];
      const resolved=path.posix.normalize(target.startsWith('/')?target:path.posix.join(base,target)).replace(/^\//,'');
      assert(zip.file(resolved),`${name}: missing relationship target ${target}`);
    }
  }
}

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
  const scores = {2:6,3:3.5,4:1,5:2};
  const sections = [1, 2, 3, 4, 5, 8].map(n => ({ number: n, title: 'Fråga ' + n, score: scores[n] ?? null, choice: 'Egen poäng', explanation: 'Eget mellanvärde', comment: `MOTIVERING_${n}\n` + (n === 5 ? 'Lång observation. '.repeat(600) + 'SLUT_MARKOR' : 'Observerat vid arbetsstationen.') }));
  const report = { metadata: ['60', 'Formning', 'Station 1'], sections, scores: ['3.5', '1', '2', '6.5', '6', '39'], range: '25 till 49 poäng', result: 'Kraftigt ökad belastningssituation.', legend: ['0 till 9', '10 till 24', '25 till 49', '50+'] };
  const evidence = { sections: Object.fromEntries([1, 2, 3, 4, 5, 8].map(n => [n, { images: [{ name: 'foto.png', caption: 'BILDTEXT_' + n, width: 500, height: 359, data: photo }] }])), attachments: [{ name: 'Arbetsinstruktion.txt', description: 'Gällande instruktion.', bytes: new TextEncoder().encode('Originaldokumentets innehåll: åäö.') }] };
  const assets={logoData:'data:image/png;base64,'+fs.readFileSync(path.join(root,'Dometic_hori_office_neg_logga_vit.png')).toString('base64')};
  const zip = await ctx.JSZip.loadAsync(await ctx.KimPowerPoint.build(report, evidence, assets));
  await validatePowerPointPackage(zip);
  const slides = Object.keys(zip.files).filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name)).sort((a, b) => Number(a.match(/(\d+)\.xml/)[1]) - Number(b.match(/(\d+)\.xml/)[1]));
  const xml = await Promise.all(slides.map(name => zip.file(name).async('string')));
  const all = xml.join('\n');
  for (const n of [1, 2, 3, 4, 5, 8]) { assert(all.includes('MOTIVERING_' + n)); assert(all.includes('BILDTEXT_' + n)); }
  assert(all.includes('SLUT_MARKOR'), 'Long comments must never be truncated');
  assert(xml[0].includes('Samlad riskklass') && xml[0].includes('Valda poäng och motiveringar'), 'The complete one-pager is the first slide');
  assert(all.includes('typeface="Gibson"') && all.includes('typeface="Gibson Light"'), 'Dometic typefaces are used');
  assert(all.includes('<a:tbl>'), 'Scores must remain native editable tables');
  assert.equal((all.match(/<p:oleObj /g) || []).length, 1);
  assert.equal((all.match(/<p:pic>/g) || []).length,7+slides.length,'Six separate photos, the attachment picture and one logo per slide');
  assert.deepEqual(olePayload(await zip.file('ppt/embeddings/oleObject1.bin').async('uint8array')), Buffer.from(evidence.attachments[0].bytes));
  const withoutFiles = await ctx.JSZip.loadAsync(await ctx.KimPowerPoint.build(report, { sections: {}, attachments: [] }));
  await validatePowerPointPackage(withoutFiles);
  assert(!Object.keys(withoutFiles.files).some(name => /^ppt\/embeddings\/.+\.bin$/.test(name)));
  const shortReport = {...report,sections:sections.map(s=>({...s,comment:'Kort motivering.'}))};
  const compact = await ctx.JSZip.loadAsync(await ctx.KimPowerPoint.build(shortReport, evidence, assets));
  await validatePowerPointPackage(compact);
  assert.equal(Object.keys(compact.files).filter(n=>/^ppt\/slides\/slide\d+\.xml$/.test(n)).length,3,'Six photos and a document fit into three slides including the one-pager');
  const sevenFiles=Array.from({length:7},(_,i)=>({name:`Bilaga ${i+1}.txt`,description:'Dokumentunderlag.',bytes:new Uint8Array([i,42,255])}));
  const several=await ctx.JSZip.loadAsync(await ctx.KimPowerPoint.build(shortReport,{...evidence,attachments:sevenFiles}));
  await validatePowerPointPackage(several);
  for(let i=0;i<sevenFiles.length;i++) assert.deepEqual(olePayload(await several.file(`ppt/embeddings/oleObject${i+1}.bin`).async('uint8array')),Buffer.from(sevenFiles[i].bytes));
  const summaryOnly = await ctx.JSZip.loadAsync(await ctx.KimPowerPoint.build(shortReport,{sections:{},attachments:[]},assets));
  await validatePowerPointPackage(summaryOnly);
  assert.equal(Object.keys(summaryOnly.files).filter(n=>/^ppt\/slides\/slide\d+\.xml$/.test(n)).length,1,'A short analysis without photos needs one slide');
  const xlsx=await ctx.JSZip.loadAsync(await ctx.KimExcel.build(report,evidence));
  const worksheet=await xlsx.file('xl/worksheets/sheet1.xml').async('string');
  assert(worksheet.includes('SUM(E14:E16)*E13'),'Excel risk is a real recalculating formula');
  assert(worksheet.includes('r="E13" s="5"><v>6</v>'),'Time score is numeric and editable');
  const details=await xlsx.file('xl/worksheets/sheet2.xml').async('string');
  assert(details.includes('SLUT_MARKOR'),'Excel keeps the full long comment');
  assert.equal(Object.keys(xlsx.files).filter(n=>/^xl\/media\/image\d+\./.test(n)).length,6,'Excel has native picture objects');
  for(const [value,level] of [[9.99,1],[10,2],[24.99,2],[25,3],[49.99,3],[50,4]]) assert.equal(ctx.KimReport.riskInfo(value).level,level);
  const html=fs.readFileSync(path.join(root,'index.html'),'utf8');
  for(const id of ['pptxBtn','excelBtn','printBtn']) assert(html.includes(`id="${id}"`));
  for(const id of ['packageBtn','pdfBtn','csvBtn']) assert(!html.includes(`id="${id}"`));
  console.log('PASS: one-pager first, Gibson, three-slide normal report, complete long notes, native Excel formulas/images, risk boundaries and unchanged document payloads.');
})().catch(error => { console.error(error); process.exitCode = 1; });
