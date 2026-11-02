/* PowerPoint export: editable Dometic report with embedded document packages. */
(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const R = window.KimReport;
  const XML = R.xml;
  function attribute(tag, name) {
    const match = new RegExp('\\b' + name + '="([^"]*)"').exec(tag);
    return match ? match[1].replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>') : null;
  }

  async function finalizePackage(zip) {
    // PptxGenJS can give a table the same ID as an earlier title when a logo is
    // present. Drawing IDs must be unique within each slide/master/layout part.
    const drawingParts = Object.keys(zip.files).filter(name => /^ppt\/(slides|slideMasters|slideLayouts)\/[^/]+\.xml$/.test(name));
    for (const name of drawingParts) {
      const xml = await zip.file(name).async('string');
      for (const match of xml.matchAll(/<a:(?:tcPr|bodyPr)\b[^>]*\banchor="([^"]+)"/g)) {
        if (!['t', 'ctr', 'b', 'just', 'dist'].includes(match[1])) {
          throw new Error('PowerPoint-filen innehåller en ogiltig textformatering. Uppdatera sidan och försök igen.');
        }
      }
      const ids = Array.from(xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g), match => Number(match[1]));
      let next = Math.max(1, ...ids) + 1;
      const used = new Set();
      const normalized = xml.replace(/<p:cNvPr\b[^>]*>/g, tag => {
        const id = attribute(tag, 'id');
        if (!used.has(id)) { used.add(id); return tag; }
        const replacement = String(next++);
        used.add(replacement);
        return tag.replace(/\bid="\d+"/, `id="${replacement}"`);
      });
      zip.file(name, normalized);
    }

    // PptxGenJS places the notes master after the slide list. ECMA-376's
    // CT_Presentation sequence requires it before that list.
    let presentation = await zip.file('ppt/presentation.xml').async('string');
    const notesMaster = presentation.match(/<p:notesMasterIdLst>[^]*?<\/p:notesMasterIdLst>/);
    if (notesMaster && presentation.indexOf(notesMaster[0]) > presentation.indexOf('<p:sldIdLst>')) {
      presentation = presentation.replace(notesMaster[0], '').replace('<p:sldIdLst>', notesMaster[0] + '<p:sldIdLst>');
      zip.file('ppt/presentation.xml', presentation);
    }

    // The writer also declares extra, unused masters for multi-slide reports.
    // Remove declarations for parts it did not create; referenced parts are
    // checked below, so a genuinely missing dependency cannot be hidden.
    const types = await zip.file('[Content_Types].xml').async('string');
    zip.file('[Content_Types].xml', types.replace(/<Override\b[^>]*\/>/g, tag => {
      const part = attribute(tag, 'PartName');
      return part && zip.file(part.replace(/^\//, '')) ? tag : '';
    }));
    for (const name of Object.keys(zip.files).filter(path => path.endsWith('.rels'))) {
      const base = name === '_rels/.rels' ? '' : name.slice(0, name.lastIndexOf('/_rels/'));
      const xml = await zip.file(name).async('string');
      for (const match of xml.matchAll(/<Relationship\b[^>]*\/>/g)) {
        if (attribute(match[0], 'TargetMode') === 'External') continue;
        const target = attribute(match[0], 'Target');
        if (!target) throw new Error('PowerPoint-filens struktur kunde inte kontrolleras.');
        const segments = [];
        const absolute = target.startsWith('/') ? target : base + '/' + target;
        for (const segment of absolute.split('/')) {
          if (segment === '..') segments.pop();
          else if (segment && segment !== '.') segments.push(segment);
        }
        if (!zip.file(segments.join('/'))) throw new Error('PowerPoint-filen saknar en nödvändig del. Försök exportera igen.');
      }
    }
    return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
  }

  async function embedDocuments(buffer, objects) {
    const zip = await JSZip.loadAsync(buffer, { checkCRC32: true });
    let contentTypes = await zip.file('[Content_Types].xml').async('string');
    objects.forEach((object, i) => {
      contentTypes = contentTypes.replace('</Types>', `<Override PartName="/ppt/embeddings/oleObject${i + 1}.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/></Types>`);
    });
    zip.file('[Content_Types].xml', contentTypes);
    for (let i = 0; i < objects.length; i++) {
      const { slideNumber, file, x, y, w, h } = objects[i];
      const path = `ppt/slides/slide${slideNumber}.xml`;
      const relPath = `ppt/slides/_rels/slide${slideNumber}.xml.rels`;
      let xml = await zip.file(path).async('string');
      let rels = await zip.file(relPath).async('string');
      const pictures = xml.match(/<p:pic>[^]*?<\/p:pic>/g) || [];
      const picture = pictures.find(pic => pic.includes(`name="KIM_ATTACHMENT_${i}"`));
      if (!picture) throw new Error('Bilagans plats i presentationen kunde inte skapas.');
      const ids = Array.from(xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g), match => Number(match[1]));
      const shapeId = Math.max(1, ...ids) + 1;
      const relId = `rIdKimAttachment${i + 1}`;
      const emu = n => Math.round(n * 914400);
      const frame = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${shapeId}" name="${XML(file.name)}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj spid="_x0000_s${10000 + i}" name="${XML(file.name)}" showAsIcon="1" r:id="${relId}" imgW="${emu(w)}" imgH="${emu(h)}" progId="Package"><p:embed/>${picture}</p:oleObj></a:graphicData></a:graphic></p:graphicFrame>`;
      xml = xml.replace(picture, frame);
      rels = rels.replace('</Relationships>', `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject${i + 1}.bin"/></Relationships>`);
      zip.file(path, xml);
      zip.file(relPath, rels);
      zip.file(`ppt/embeddings/oleObject${i + 1}.bin`, KimOlePackage(file.name, file.bytes));
    }
    return finalizePackage(zip);
  }


  window.KimEmbedDocuments = embedDocuments;
  async function exportReport() {
    if (window.KimExportBusy) return;
    window.KimExportBusy = true;
    const buttons = ['pptxBtn', 'excelBtn', 'printBtn'].map($);
    buttons.forEach(b => { if (b) b.disabled = true; });
    const original = $('pptxBtn').textContent;
    $('pptxBtn').textContent = 'Skapar PowerPoint…';
    try {
      await window.KimEvidence.flush();
      const report = window.KimAnalysis.report(), evidence = structuredClone(window.KimEvidence.getState());
      const buffer = await window.KimBuildPowerPoint(report, evidence, await R.assets());
      const filename = R.safeName(['KIM1', ...report.metadata.filter(Boolean)].join('_'));
      window.KimEvidence.download(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }), filename + '.pptx');
      window.KimEvidence.message('PowerPoint är klar. Text, tabeller och bilder är redigerbara.');
    } catch (error) {
      console.error('PowerPoint export failed', error);
      window.KimEvidence.message('Exporten misslyckades: ' + (error.message || 'försök igen.'), true);
    } finally {
      window.KimExportBusy = false;
      buttons.forEach(b => { if (b) b.disabled = false; });
      $('pptxBtn').textContent = original;
    }
  }
  $('pptxBtn').onclick = exportReport;
  window.KimPowerPoint = { build: (...args) => window.KimBuildPowerPoint(...args) };
})();
