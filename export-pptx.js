/* Editable, offline PowerPoint export for KIM 1. */
(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const WIDTH = 11.6929, HEIGHT = 8.2677;
  const BLUE = '156082', INK = '172333', GREY = '637081';
  const XML = text => String(text).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[c]);
  const safeName = text => String(text || 'analys').replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_').trim().slice(0, 160) || 'analys';
  const BLANK_PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR4nGNgAAIAAAUAAXpeqz8AAAAASUVORK5CYII=';

  function wrapText(text, width, fontSize) {
    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d');
    context.font = `${fontSize * 96 / 72}px Arial`;
    const maxWidth = width * 96 * .92;
    const lines = [];
    for (const paragraph of String(text).replace(/\r/g, '').split('\n')) {
      if (!paragraph) { lines.push(''); continue; }
      let line = '';
      for (const word of paragraph.split(/\s+/)) {
        if (line && context.measureText(line + ' ' + word).width > maxWidth) { lines.push(line); line = ''; }
        if (context.measureText(word).width <= maxWidth) { line += (line ? ' ' : '') + word; continue; }
        for (const char of word) {
          if (line && context.measureText(line + char).width > maxWidth) { lines.push(line); line = ''; }
          line += char;
        }
      }
      lines.push(line);
    }
    return lines;
  }

  function addTextPages(pptx, title, body, base) {
    const lines = wrapText(body, WIDTH - 1.2, 18);
    let index = 0, page = 0;
    do {
      const slide = base(title + (page ? ' (fortsättning)' : ''));
      const part = lines.slice(index, index + 18).join('\n');
      slide.addText(part, { x: .6, y: 1.4, w: WIDTH - 1.2, h: 5.95, fontFace: 'Arial', fontSize: 18, color: INK, margin: 0, breakLine: false, valign: 'top', lineSpacingMultiple: 1.13, paraSpaceAfterPt: 3, objectName: 'Redigerbar bedömningstext' });
      index += 18; page++;
    } while (index < lines.length);
  }

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
    pptx.theme = { headFontFace: 'Arial', bodyFontFace: 'Arial', lang: 'sv-SE' };
    let slideCount = 0;
    const date = new Date().toLocaleDateString('sv-SE', { timeZone: 'Europe/Stockholm' });
    function base(title) {
      const slide = pptx.addSlide();
      slide.background = { color: 'FFFFFF' };
      slide.addText(title, { x: .6, y: .35, w: WIDTH - 1.2, h: .85, fontFace: 'Arial', fontSize: title.length > 65 ? 25 : 30, bold: true, color: BLUE, margin: 0, valign: 'mid', objectName: 'Redigerbar rubrik' });
      slide.addText(`Dometic   KIM 1   ${date}`, { x: .6, y: HEIGHT - .46, w: 9.4, h: .2, fontFace: 'Arial', fontSize: 10, color: GREY, margin: 0 });
      slide.addText(String(++slideCount), { x: WIDTH - 1.15, y: HEIGHT - .46, w: .5, h: .2, fontFace: 'Arial', fontSize: 10, color: GREY, align: 'right', margin: 0 });
      return slide;
    }
    const scoreRows = () => ['Belastningspoäng', 'Arbetsförhållanden', 'Arbetsställning', 'Summa belastning', 'Tidspoäng', 'Riskpoäng'].map((label, i) => [label, report.scores[i]]);
    function calculation(title, summary = false) {
      const slide = base(title);
      slide.addTable(scoreRows(), { x: .6, y: 1.5, w: 5, colW: [3.8, 1.2], rowH: .55, fontFace: 'Arial', fontSize: 18, color: INK, border: { color: 'D9E1EA', pt: .7 }, margin: .12, fill: 'FFFFFF', autoPage: false });
      slide.addText(report.range, { x: 6, y: 1.5, w: 4.9, h: .55, fontSize: 23, bold: true, color: BLUE, margin: 0 });
      slide.addText(report.result, { x: 6, y: 2.3, w: 4.9, h: 2.8, fontSize: 18, color: INK, margin: 0, valign: 'top', breakLine: false });
      slide.addText(`(${report.scores[0]} + ${report.scores[1]} + ${report.scores[2]}) × ${report.scores[4]} = ${report.scores[5]}`, { x: .6, y: 5.5, w: 10.4, h: .6, fontSize: 24, bold: true, color: BLUE, margin: 0 });
      if (summary) {
        slide.addText(['Fabrik', 'Process', 'Underprocess'].map((label, i) => `${label}: ${report.metadata[i] || 'Ej angivet'}`).join('\n'), { x: .6, y: 6.45, w: 10.4, h: .95, fontSize: 14, color: GREY, margin: 0, valign: 'top' });
      }
    }
    function sectionPages(section) {
      const n = section.number;
      const title = n === 8 ? '8. Övriga kommentarer och observationer' : `${n}. ${section.title}`;
      const parts = [];
      if (n === 1) parts.push(['Fabrik', 'Process', 'Underprocess'].map((label, i) => `${label}: ${report.metadata[i] || 'Ej angivet'}`).join('\n'));
      if (section.score !== null) parts.push(`Vald nivå: ${section.choice}\nPoäng: ${section.score}\n${section.explanation}`);
      parts.push(`${n === 8 ? section.title : 'Motivering / kommentar'}\n${section.comment.trim() || 'Ej angivet'}`);
      addTextPages(pptx, title, parts.join('\n\n'), base);
      const images = evidence.sections[n]?.images || [];
      for (let at = 0; at < images.length; at += 2) {
        const slide = base(`${n}. Bilder och observationer`);
        images.slice(at, at + 2).forEach((image, column) => {
          const x = .6 + column * 5.45;
          const w = 5.04, h = 4.4, y = 1.4;
          // Separate native picture objects retain their original proportions.
          const ratio = image.width / image.height;
          let iw = w, ih = w / ratio;
          if (ih > h) { ih = h; iw = h * ratio; }
          slide.addImage({ data: image.data, x: x + (w - iw) / 2, y: y + (h - ih) / 2, w: iw, h: ih, altText: image.caption || image.name, objectName: `Fråga ${n} – ${image.name}` });
          const captionLines = wrapText(image.caption || image.name, w, 16);
          slide.addText(captionLines.slice(0, 4).join('\n'), { x, y: 5.98, w, h: 1.12, fontSize: 16, color: INK, margin: 0, valign: 'top', objectName: 'Redigerbar bildtext' });
          // Long captions continue on their own slide instead of being truncated.
        });
        for (const image of images.slice(at, at + 2)) {
          if (wrapText(image.caption || image.name, 5.04, 16).length > 4) addTextPages(pptx, `${n}. Bildtext – ${image.name}`, image.caption || image.name, base);
        }
      }
    }
    report.sections.filter(section => section.number <= 5).forEach(sectionPages);
    calculation('6. Ergonomiberäkning');
    addTextPages(pptx, '7. Resultatförklaring', report.legend.join('\n\n'), base);
    sectionPages(report.sections.find(section => section.number === 8));
    const objects = [];
    evidence.attachments.forEach((file, i) => {
      const slide = base(`Dokumentbilaga ${i + 1}`);
      slide.addText(file.name, { x: .6, y: 1.55, w: WIDTH - 1.2, h: .8, fontSize: 23, bold: true, color: INK, margin: 0, valign: 'top', objectName: 'Bilagans filnamn' });
      const x = .6, y = 2.8, w = 4.7, h = .7;
      slide.addText('Öppna bilagan', { x, y, w, h, fontSize: 24, bold: true, underline: { color: BLUE }, color: BLUE, margin: 0 });
      slide.addImage({ data: BLANK_PNG, x, y, w, h, objectName: `KIM_ATTACHMENT_${i}`, altText: `Öppna ${file.name}` });
      slide.addText('Dubbelklicka på Öppna bilagan i PowerPoint för Windows.\nDokumentet finns inbäddat i denna presentation.', { x: .6, y: 3.85, w: 10.4, h: 1, fontSize: 17, color: GREY, margin: 0, valign: 'top' });
      const lines = wrapText(file.description || '', 10.4, 18);
      slide.addText(lines.slice(0, 5).join('\n'), { x: .6, y: 5.2, w: 10.4, h: 1.6, fontSize: 18, color: INK, margin: 0, valign: 'top' });
      objects.push({ slideNumber: slideCount, file, x, y, w, h });
      if (lines.length > 5) addTextPages(pptx, `Bilaga ${i + 1} – beskrivning`, file.description, base);
    });
    calculation('Sammanfattning – Ergonomianalys KIM 1', true);
    return embedDocuments(await pptx.write({ outputType: 'arraybuffer' }), objects);
  }

  let exporting = false;
  async function exportReport(withOriginals) {
    if (exporting) return;
    exporting = true;
    const buttons = [$('pptxBtn'), $('packageBtn')];
    buttons.forEach(button => button.disabled = true);
    const button = withOriginals ? buttons[1] : buttons[0];
    const oldText = button.textContent;
    button.textContent = 'Skapar export…';
    try {
      await window.KimEvidence.flush();
      window.KimEvidence.message('Skapar redigerbar PowerPoint…');
      const report = window.KimAnalysis.report();
      const evidence = structuredClone(window.KimEvidence.getState());
      const buffer = await build(report, evidence);
      const filename = safeName(['KIM1', ...report.metadata.filter(Boolean)].join('_'));
      if (withOriginals) {
        const zip = new JSZip();
        zip.file(filename + '.pptx', buffer);
        evidence.attachments.forEach((file, i) => zip.file(`Bilagor/${String(i + 1).padStart(2, '0')}_${safeName(file.name)}`, file.bytes));
        zip.file('Las_mig.txt', 'PowerPoint-filen innehåller redigerbar text och separata bilder.\r\nDokumenten är även inbäddade i presentationen och finns i mappen Bilagor.\r\nÖppna inbäddade dokument i PowerPoint för Windows.\r\n');
        window.KimEvidence.download(await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' }), filename + '.zip');
      } else {
        window.KimEvidence.download(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }), filename + '.pptx');
      }
      window.KimEvidence.message('Exporten är klar. Text och bilder kan redigeras i PowerPoint.');
    } catch (error) {
      console.error('PowerPoint export failed', error);
      window.KimEvidence.message('Exporten misslyckades: ' + (error.message || 'försök igen.'), true);
    } finally {
      exporting = false;
      buttons.forEach(button => button.disabled = false);
      button.textContent = oldText;
    }
  }
  $('pptxBtn').onclick = () => exportReport(false);
  $('packageBtn').onclick = () => exportReport(true);
  window.KimPowerPoint = { build };
})();
