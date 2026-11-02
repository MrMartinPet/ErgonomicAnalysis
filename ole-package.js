/* Minimal MS-CFB / MS-OLEDS Package writer for document attachments.
 * No external file references or executable code are added to the presentation.
 * Format references: https://learn.microsoft.com/openspecs/windows_protocols/ms-cfb/
 * https://learn.microsoft.com/openspecs/windows_protocols/ms-oleds/
 */
((root) => {
  'use strict';
  const FREE = 0xffffffff, END = 0xfffffffe, FAT = 0xfffffffd, DIF = 0xfffffffc;
  const concat = (...arrays) => {
    const result = new Uint8Array(arrays.reduce((sum, a) => sum + a.length, 0));
    let at = 0;
    for (const array of arrays) { result.set(array, at); at += array.length; }
    return result;
  };
  const word = n => { const b = new Uint8Array(2); new DataView(b.buffer).setUint16(0, n, true); return b; };
  const dword = n => { const b = new Uint8Array(4); new DataView(b.buffer).setUint32(0, n, true); return b; };
  const ansi = s => Uint8Array.from([...s, '\0'], c => c.charCodeAt(0) <= 255 ? c.charCodeAt(0) : 95);
  const clsid = Uint8Array.from([12, 0, 3, 0, 0, 0, 0, 0, 192, 0, 0, 0, 0, 0, 0, 70]);
  function packageStreams(filename, bytes) {
    const name = ansi(filename.replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_').slice(0, 180) || 'bilaga');
    const native = concat(word(2), name, name, word(0), word(3), dword(name.length), name, dword(bytes.length), bytes);
    const header = concat(dword(0xfffe0001), dword(0x00000a03), dword(0xffffffff), clsid);
    const packageName = ansi('Package');
    const compObj = concat(header, dword(packageName.length), packageName, dword(0), dword(0));
    return [
      { name: '\u0001Ole', data: concat(dword(0x02000001), dword(0), dword(0), dword(0), dword(0)) },
      { name: '\u0001CompObj', data: compObj },
      { name: '\u0001Ole10Native', data: concat(dword(native.length), native) }
    ];
  }

  function compoundPackage(filename, bytes) {
    bytes = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    const streams = packageStreams(filename, bytes);
    const sectors = [], chains = [];
    function allocate(data) {
      const first = sectors.length;
      const count = Math.ceil(data.length / 512);
      for (let i = 0; i < count; i++) {
        const sector = new Uint8Array(512);
        sector.set(data.subarray(i * 512, (i + 1) * 512));
        sectors.push(sector);
      }
      if (count) chains.push({ first, count });
      return count ? first : END;
    }
    const miniParts = [], miniFat = [];
    let miniCount = 0;
    for (const stream of streams) {
      if (stream.data.length >= 4096) { stream.start = allocate(stream.data); continue; }
      stream.start = miniCount;
      const count = Math.ceil(stream.data.length / 64);
      const part = new Uint8Array(count * 64);
      part.set(stream.data);
      miniParts.push(part);
      for (let i = 0; i < count; i++) miniFat.push(i === count - 1 ? END : miniCount + i + 1);
      miniCount += count;
    }
    const miniStream = concat(...miniParts);
    const miniStart = allocate(miniStream);
    const miniFatData = new Uint8Array(Math.ceil(miniFat.length * 4 / 512) * 512).fill(255);
    const miniView = new DataView(miniFatData.buffer);
    miniFat.forEach((value, i) => miniView.setUint32(i * 4, value, true));
    const miniFatStart = allocate(miniFatData);
    const directory = new Uint8Array(512);
    const dirStart = allocate(directory);
    let fatCount = 1, difCount = 0;
    while (true) {
      const nextFat = Math.ceil((sectors.length + fatCount + difCount) / 128);
      const nextDif = Math.ceil(Math.max(0, nextFat - 109) / 127);
      if (nextFat === fatCount && nextDif === difCount) break;
      fatCount = nextFat; difCount = nextDif;
    }
    const fatIds = Array.from({ length: fatCount }, (_, i) => sectors.length + i);
    const difIds = Array.from({ length: difCount }, (_, i) => sectors.length + fatCount + i);
    const fatData = new Uint8Array(fatCount * 512).fill(255);
    const fatView = new DataView(fatData.buffer);
    for (const { first, count } of chains) {
      for (let i = 0; i < count; i++) fatView.setUint32((first + i) * 4, i === count - 1 ? END : first + i + 1, true);
    }
    for (const id of fatIds) fatView.setUint32(id * 4, FAT, true);
    for (const id of difIds) fatView.setUint32(id * 4, DIF, true);
    for (let i = 0; i < fatCount; i++) sectors.push(fatData.slice(i * 512, (i + 1) * 512));
    for (let i = 0; i < difCount; i++) {
      const part = new Uint8Array(512).fill(255);
      const view = new DataView(part.buffer);
      for (let j = 0; j < 127; j++) {
        const id = fatIds[109 + i * 127 + j];
        if (id !== undefined) view.setUint32(j * 4, id, true);
      }
      view.setUint32(508, i === difCount - 1 ? END : difIds[i + 1], true);
      sectors.push(part);
    }
    function entry(index, name, type, color, left, right, child, start, length) {
      const offset = index * 128;
      const view = new DataView(sectors[dirStart].buffer);
      for (let i = 0; i < name.length; i++) view.setUint16(offset + i * 2, name.charCodeAt(i), true);
      view.setUint16(offset + 64, (name.length + 1) * 2, true);
      view.setUint8(offset + 66, type);
      view.setUint8(offset + 67, color);
      view.setUint32(offset + 68, left, true);
      view.setUint32(offset + 72, right, true);
      view.setUint32(offset + 76, child, true);
      if (type === 5) sectors[dirStart].set(clsid, offset + 80);
      view.setUint32(offset + 116, start, true);
      view.setUint32(offset + 120, length, true);
    }
    // MS-CFB directory ordering is name length followed by case-insensitive name.
    entry(0, 'Root Entry', 5, 1, FREE, FREE, 2, miniStart, miniStream.length);
    entry(1, streams[0].name, 2, 0, FREE, FREE, FREE, streams[0].start, streams[0].data.length);
    entry(2, streams[1].name, 2, 1, 1, 3, FREE, streams[1].start, streams[1].data.length);
    entry(3, streams[2].name, 2, 0, FREE, FREE, FREE, streams[2].start, streams[2].data.length);
    const header = new Uint8Array(512);
    header.set([208, 207, 17, 224, 161, 177, 26, 225]);
    const h = new DataView(header.buffer);
    h.setUint16(24, 0x003e, true); h.setUint16(26, 3, true);
    h.setUint16(28, 0xfffe, true); h.setUint16(30, 9, true); h.setUint16(32, 6, true);
    h.setUint32(44, fatCount, true); h.setUint32(48, dirStart, true);
    h.setUint32(56, 4096, true); h.setUint32(60, miniFatStart, true);
    h.setUint32(64, miniFatData.length / 512, true);
    h.setUint32(68, difIds[0] ?? END, true); h.setUint32(72, difCount, true);
    for (let i = 0; i < 109; i++) h.setUint32(76 + i * 4, fatIds[i] ?? FREE, true);
    return concat(header, ...sectors);
  }
  root.KimOlePackage = compoundPackage;
})(typeof window !== 'undefined' ? window : globalThis);
