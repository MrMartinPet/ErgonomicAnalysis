/* Editable, offline PowerPoint export for KIM 1. */
(() => {
  'use strict';
  const R = window.KimReport;
  const $ = id => document.getElementById(id);
  const WIDTH = 11.6929, HEIGHT = 8.2677;
  const BLUE = '156082', INK = '172333', GREY = '637081';
  const XML = text => String(text).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[c]);
  const safeName = text => String(text || 'analys').replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_').trim().slice(0, 160) || 'analys';
  const BLANK_PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR4nGNgAAIAAAUAAXpeqz8AAAAASUVORK5CYII=';

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

  async function build(report, evidence) {
    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: 'A4_LIGGANDE', width: WIDTH, height: HEIGHT });
    pptx.layout = 'A4_LIGGANDE';
    pptx.author = 'Dometic';
    pptx.subject = 'Ergonomianalys KIM 1 med motiveringar, bilder och bilagor';
    pptx.title = ['Ergonomianalys KIM 1', ...report.metadata.filter(Boolean)].join(' – ');
    pptx.company = 'Dometic';
    pptx.lang = 'sv-SE';
    pptx.theme = { headFontFace: R.HEADING, bodyFontFace: R.FONT, lang: 'sv-SE' };

    const margin = .5, width = WIDTH - 2 * margin, bottom = 7.69;
    const date = new Date().toLocaleDateString('sv-SE', { timeZone: 'Europe/Stockholm' });
    const objects = [];
    let slideCount = 0;

    function text(slide, value, x, y, w, h, style = {}) {
      const options = { fontSize: 11, color: INK, fontFace: R.FONT, margin: 0, valign: 'top', ...style };
      slide.addText(String(value ?? '').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, ''), {
        x, y, w, h, lineSpacing: options.fontSize * 1.18, ...options
      });
    }

    function base(title) {
      const slide = pptx.addSlide();
      slide.background = { color: 'FFFFFF' };
      text(slide, title, margin, .37, width, .42, { fontFace: R.HEADING, fontSize: 24 });
      text(slide, `DOMETIC  /  KIM 1  /  ${date}`, margin, 7.89, 8, .2, { fontSize: 8.5, color: GREY });
      text(slide, ++slideCount, WIDTH - 1, 7.89, .5, .2, { fontSize: 8.5, color: GREY, align: 'right' });
      return slide;
    }

    const summary = base('Ergonomianalys KIM 1');
    const metadataLabels = ['Fabrik', 'Process', 'Underprocess'];
    const metadataWidths = [2, 4.1, 4.1];
    const metadataX = [margin, 2.65, 6.95];
    const longMetadata = [];
    metadataLabels.forEach((label, i) => {
      const value = report.metadata[i];
      const shown = R.preview(value, metadataWidths[i], 12, 2);
      text(summary, label.toUpperCase(), metadataX[i], 1.0, metadataWidths[i], .17, { fontSize: 9, color: GREY, bold: true });
      text(summary, shown.text, metadataX[i], 1.19, metadataWidths[i], .44, { fontSize: 12 });
      if (shown.cut) longMetadata.push(`${label}: ${value}`);
    });

    const risk = R.riskInfo(report.scores[5]);
    text(summary, R.format(report.scores[5]), margin, 1.76, 1.65, 1.22, {
      fontSize: 37, fontFace: R.HEADING, align: 'center', valign: 'mid', fill: { color: risk.fill }
    });
    text(summary, `Riskklass ${risk.level} · ${risk.label}`, 2.39, 1.8, 8.55, .35, { fontSize: 19, fontFace: R.HEADING });
    text(summary, R.wrap(report.result || risk.action, 8.55, 11).join('\n'), 2.39, 2.22, 8.55, .77);
    text(summary, `Riskpoäng = (${report.scores.slice(0, 3).map(R.format).join(' + ')}) × ${R.format(report.scores[4])} = ${R.format(report.scores[5])}`,
      margin, 3.12, width, .25, { fontSize: 11, bold: true });

    text(summary, 'Ergonomiberäkning', margin, 3.67, 4.5, .25, { fontSize: 14, bold: true });
    const scoreRows = ['Belastningspoäng', 'Arbetsförhållanden', 'Arbetsställning', 'Summa belastning', 'Tidspoäng', 'Riskpoäng']
      .map((label, i) => [label, R.format(report.scores[i])]);
    // Editable text rows avoid the duplicate drawing IDs emitted by addTable.
    scoreRows.forEach(([label, value], i) => {
      const y = 4.05 + i * .39, total = i === scoreRows.length - 1;
      const fill = total ? 'EAF5F3' : i % 2 === 0 ? 'F3F3F2' : 'FFFFFF';
      text(summary, ' ', margin, y, 4.5, .37, { fill: { color: fill } });
      text(summary, label, margin + .11, y, 3.3, .37, { bold: total, valign: 'mid' });
      text(summary, value, 3.96, y, .91, .37, {
        fontSize: 11, bold: true, align: 'right', valign: 'mid'
      });
    });

    text(summary, 'Resultatförklaring', 5.32, 3.67, 5.87, .25, { fontSize: 14, bold: true });
    let legendY = 4.05;
    report.legend.forEach((entry, i) => {
      const [range, ...description] = entry.trim().split('\n');
      const lines = R.wrap(description.join(' ').trim(), 5.55, 9.5).filter(Boolean);
      const height = Math.max(.37, .29 + lines.length * .16);
      const fill = i === risk.level - 1 ? risk.fill : 'F3F3F2';
      text(summary, ' ', 5.32, legendY, 5.87, height, { fill: { color: fill } });
      text(summary, range, 5.43, legendY + .045, 5.55, .19, { fontSize: 10, bold: true });
      if (lines.length) text(summary, lines.join('\n'), 5.43, legendY + .24, 5.55, height - .22, { fontSize: 9.5 });
      legendY += height + .06;
    });

    let slide, y = bottom;
    function reserve(height, title) {
      if (y + height > bottom) {
        slide = base(title);
        y = .98;
      }
      const top = y;
      y += height + .13;
      return top;
    }

    for (const section of report.sections) {
      const number = section.number;
      const title = `${number}. ${R.titles[number] || section.title}`;
      const photos = evidence.sections[number]?.images || [];
      const photoWidth = 1.87, photoHeight = 1.14, photoGap = .13, photoX = 5.32;
      const captions = photos.map((photo, i) => {
        const value = photo.caption || photo.name;
        const shown = R.preview(value, photoWidth, 9, 2);
        return shown.cut ? {
          text: R.preview(`Bild ${i + 1} · ${value}`, photoWidth, 9, 2).text,
          full: `Bild ${i + 1}: ${value}`
        } : { text: shown.text, full: '' };
      });
      const description = section.explanation && section.explanation !== section.choice ? section.explanation : '';
      const choice = section.score == null ? '' : `${R.format(section.score)} p · ${section.choice}`;
      const details = [number === 1 ? longMetadata.join('\n') : '', choice, description,
        section.comment.trim(), ...captions.map(caption => caption.full)].filter(Boolean).join('\n');
      const lines = R.wrap(details || (photos.length ? '' : 'Ingen kommentar angiven.'), photos.length ? 4.55 : width - .3, 11);
      let nextLine = 0, nextPhoto = 0, continued = false;
      do {
        const photoRow = photos.slice(nextPhoto, nextPhoto + 3);
        const chunk = lines.slice(nextLine, nextLine + (photoRow.length ? 9 : 29));
        const height = Math.max(.54 + chunk.length * .187, photoRow.length ? 2.03 : 0);
        const top = reserve(height, 'Bedömning och bildunderlag');
        text(slide, title + (continued ? ' · fortsättning' : ''), margin, top, width, .32, {
          fontSize: 13, bold: true, fill: { color: 'EAF5F3' }, margin: [3, 7, 0, 7]
        });
        if (chunk.length) text(slide, chunk.join('\n'), margin + .10, top + .44,
          photos.length ? 4.55 : width - .3, height - .42);
        photoRow.forEach((photo, i) => {
          const x = photoX + (3 - photoRow.length) * (photoWidth + photoGap) / 2 + i * (photoWidth + photoGap);
          const scale = Math.min(photoWidth / photo.width, photoHeight / photo.height);
          const w = photo.width * scale, h = photo.height * scale;
          slide.addImage({ data: photo.data, x: x + (photoWidth - w) / 2, y: top + .44 + (photoHeight - h) / 2,
            w, h, altText: photo.caption || photo.name, objectName: `Fråga ${number} – ${photo.name}` });
          const caption = captions[nextPhoto + i];
          text(slide, caption.text, x, top + 1.63, photoWidth, .38, { fontSize: 9, color: GREY });
        });
        nextLine += chunk.length;
        nextPhoto += photoRow.length;
        continued = true;
      } while (nextLine < lines.length || nextPhoto < photos.length);
    }

    y = bottom;
    evidence.attachments.forEach((file, i) => {
      const body = R.wrap(`${file.name}${file.description ? '\n' + file.description : ''}`, width - .25, 11);
      for (let offset = 0; offset < body.length; offset += 27) {
        const lines = body.slice(offset, offset + 27);
        const top = reserve(.61 + lines.length * .19, 'Dokumentbilagor');
        text(slide, lines.join('\n'), margin, top, width, lines.length * .19 + .1);
        if (offset) continue;
        const x = margin, y = top + lines.length * .19 + .07, w = 4.7, h = .33;
        text(slide, 'Öppna bilagan i PowerPoint för Windows', x, y, w, h, { fontSize: 10, bold: true, color: BLUE });
        slide.addImage({ data: BLANK_PNG, x, y, w, h, objectName: `KIM_ATTACHMENT_${i}`, altText: `Öppna ${file.name}` });
        objects.push({ slideNumber: slideCount, file, x, y, w, h });
      }
    });
    return embedDocuments(await pptx.write({ outputType: 'arraybuffer' }), objects);
  }

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
      const buffer = await build(report, evidence);
      const filename = safeName(['KIM1', ...report.metadata.filter(Boolean)].join('_'));
      window.KimEvidence.download(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }), filename + '.pptx');
      window.KimEvidence.message('PowerPoint är klar. Text, poäng och bilder är redigerbara.');
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
  window.KimPowerPoint = { build };
})();
