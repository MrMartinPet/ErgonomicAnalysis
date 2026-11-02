/* PowerPoint export: editable Dometic report with embedded document packages. */
(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const R = window.KimReport;
  const XML = R.xml;
  async function embedDocuments(buffer, objects) {
    if (!objects.length) return buffer;
    const zip = await JSZip.loadAsync(buffer);
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
    return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
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
