/* Dometic A4 landscape report. Every label, score, table and photograph stays editable. */
(() => {
  'use strict';
  const R = window.KimReport, W = 11.6929, H = 8.2677, M = .45, BODY = W - M * 2;
  const BLANK = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR4nGNgAAIAAAUAAXpeqz8AAAAASUVORK5CYII=';
  async function build(report, evidence, assets = {}) {
    const m = R.model(report, evidence), pptx = new PptxGenJS();
    pptx.defineLayout({ name: 'A4_LIGGANDE', width: W, height: H });
    pptx.layout = 'A4_LIGGANDE';
    pptx.author = 'Dometic'; pptx.company = 'Dometic'; pptx.subject = 'Ergonomianalys KIM 1';
    pptx.title = ['Ergonomianalys KIM 1', ...report.metadata.filter(Boolean)].join(' – ');
    pptx.theme = { headFontFace: R.HEADING, bodyFontFace: R.FONT, lang: 'sv-SE' };
    let count = 0;
    const text = (slide, content, x, y, w, h, options = {}) => slide.addText(String(content ?? ''), { x, y, w, h, fontFace: R.FONT, fontSize: 12, color: R.C.black, margin: 0, breakLine: false, valign: 'top', paraSpaceAfterPt: 0, lineSpacingMultiple: 1.05, ...options });
    const rule = (slide, y) => slide.addShape('line', { x: M, y, w: BODY, h: 0, line: { color: R.C.line, width: .55 } });
    function base(title) {
      const slide = pptx.addSlide();
      slide.background = { color: R.C.white };
      text(slide, title.toLocaleUpperCase('sv-SE'), M, .29, BODY, .58, { fontFace: R.HEADING, fontSize: 29, objectName: 'Rubrik' });
      text(slide, `KIM 1   ${m.date}`, M, 7.94, 5.8, .17, { fontSize: 8.5, color: R.C.grey });
      text(slide, ++count, 8.4, 7.94, .45, .17, { fontSize: 8.5, color: R.C.grey, align: 'right' });
      if (assets.logoData) slide.addImage({ data: assets.logoData, x: 9.48, y: 7.88, w: 1.76, h: 1.76 * 269 / 2023, altText: 'Dometic', objectName: 'Dometic logotyp' });
      rule(slide, 7.77);
      return slide;
    }
    const summary = base('Ergonomianalys KIM 1');
    const mx = [M, 2.7, 7.0], mw = [2.0, 4.0, 4.24];
    ['Fabrik', 'Process', 'Underprocess'].forEach((label, i) => {
      text(summary, label.toUpperCase(), mx[i], 1.02, mw[i], .17, { fontSize: 8.7, bold: true, color: R.C.grey });
      text(summary, m.meta[i], mx[i], 1.22, mw[i], .28, { fontSize: 14, objectName: label });
    });
    rule(summary, 1.65);
    text(summary, R.format(m.risk), M, 1.82, 1.75, .95, { fontFace: R.HEADING, fontSize: 60, fit: 'shrink', objectName: 'Riskpoäng' });
    text(summary, 'RISKPOÄNG', M + .03, 2.78, 1.7, .18, { fontSize: 9.5, bold: true, color: R.C.grey });
    text(summary, `Riskklass ${m.riskInfo.level}: ${m.riskInfo.label}`, 2.45, 1.94, 8.45, .43, { fontFace: R.HEADING, fontSize: 23, objectName: 'Samlad riskklass' });
    text(summary, m.riskInfo.action, 2.45, 2.41, 8.2, .42, { fontSize: 13.5, objectName: 'Bedömning och åtgärdsbehov' });
    text(summary, `(${m.scores.slice(0, 3).map(R.format).join(' + ')}) × ${R.format(m.scores[4])} = ${R.format(m.risk)}`, 2.45, 2.88, 8.2, .18, { fontSize: 10.5, color: R.C.grey, objectName: 'Beräkning' });
    const head = value => ({ text: value, options: { bold: true, color: R.C.white, fill: '2C6267', fontSize: 11 } });
    const rows = [['Faktor', 'Vald nivå', 'Poäng', 'Motivering'].map(head), ...m.factors.map(f => [
      { text: f.label, options: { bold: true } }, { text: f.choice, options: { fontSize: 11.5 } },
      { text: R.format(f.score), options: { fontSize: 17, bold: true, align: 'center', valign: 'mid' } },
      { text: f.commentPreview, options: { fontSize: 12 } }
    ])];
    summary.addTable(rows, { x: M, y: 3.25, w: BODY, colW: [1.98, 3.14, .86, BODY - 5.98], rowH: [.34, .70, .70, .70, .70], fontFace: R.FONT, fontSize: 12, color: R.C.black, fill: R.C.white, margin: [5, 7, 4, 7], border: { color: R.C.line, pt: .5 }, valign: 'top', autoPage: false, objectName: 'Valda poäng och motiveringar' });
    [['<10  Låg', 'E9F0D8'], ['10–<25  Ökad', 'FFF2BD'], ['25–<50  Kraftigt ökad', 'F8E3CD'], ['≥50  Hög', 'F6DEDB']].forEach(([label, fill], i) => text(summary, label, M + i * BODY / 4, 6.55, BODY / 4, .27, { fontSize: 10.5, fill: { color: fill }, margin: [2.5, 4, 0, 4] }));
    text(summary, 'PROCESS OCH AVGRÄNSNING', M, 7.00, 5.15, .17, { fontSize: 9, bold: true });
    text(summary, m.intro, M, 7.20, 5.15, .55, { fontSize: 11.5, objectName: 'Processkommentar' });
    text(summary, 'ÖVRIGA OBSERVATIONER', 6.08, 7.00, 5.15, .17, { fontSize: 9, bold: true });
    text(summary, m.observation, 6.08, 7.20, 5.15, .55, { fontSize: 11.5, objectName: 'Övriga observationer' });

    const objects = [];
    let placedAttachments = 0;
    for (let at = 0; at < m.photos.length; at += 4) {
      const photos = m.photos.slice(at, at + 4), slide = base('Bildunderlag'), large = photos.length <= 2;
      const addFiles = large && at + 4 >= m.photos.length && m.attachments.length > 0 && m.attachments.length <= 3;
      photos.forEach((photo, i) => {
        const x = M + i % 2 * 5.58, y = large ? 1.14 : 1.14 + Math.floor(i / 2) * 3.25;
        const w = photos.length === 1 ? BODY : 5.21, h = large ? (addFiles ? 4.2 : 5.5) : 2.2;
        text(slide, `${photo.number}. ${photo.label}${photo.score === null || photo.score === undefined ? '' : '   ' + R.format(photo.score) + ' p'}`, x, y, w, .25, { fontSize: 13.2, bold: true });
        let iw = w, ih = w * photo.height / photo.width;
        if (ih > h) { ih = h; iw = h * photo.width / photo.height; }
        slide.addImage({ data: photo.data, x: x + (w - iw) / 2, y: y + .38 + (h - ih) / 2, w: iw, h: ih, altText: photo.caption || photo.name, objectName: `Fråga ${photo.number} – ${photo.name}` });
        text(slide, photo.captionPreview, x, y + .43 + h, w, .45, { fontSize: 11.5, objectName: 'Bildtext' });
      });
      if (addFiles) {
        text(slide, 'DOKUMENTBILAGOR – DUBBELKLICKA FÖR ATT ÖPPNA I POWERPOINT FÖR WINDOWS', M, 6.54, BODY, .2, { fontSize: 9, bold: true, color: R.C.grey });
        m.attachments.forEach((file, i) => {
          const x = M + i * 3.66, w = 3.42, y = 6.86;
          text(slide, R.preview(file.name, w, 11.5, 2).text, x, y, w, .38, { fontSize: 11.5, bold: true, underline: { color: R.C.black } });
          slide.addImage({ data: BLANK, x, y, w, h: .38, objectName: `KIM_ATTACHMENT_${i}`, altText: `Öppna ${file.name}` });
          const description = R.preview(file.description, w, 9.5, 2);
          text(slide, description.text, x, 7.26, w, .32, { fontSize: 9.5 });
          if (description.cut && !m.notes.some(n => n.title === `Bilaga – ${file.name}`)) m.notes.push({ title: `Bilaga – ${file.name}`, text: file.description });
          objects.push({ slideNumber: count, file, x, y, w, h: .38 });
        });
        placedAttachments = m.attachments.length;
      }
    }

    for (let at = placedAttachments; at < m.attachments.length; at += 6) {
      const slide = base('Dokumentbilagor');
      text(slide, 'Dubbelklicka på dokumentnamnet i PowerPoint för Windows.', M, 1.05, BODY, .28, { fontSize: 11.5, color: R.C.grey });
      m.attachments.slice(at, at + 6).forEach((file, i) => {
        const index = at + i, y = 1.6 + i * .9;
        text(slide, R.preview(file.name, 3.8, 13.5, 2).text, M, y, 3.9, .55, { fontSize: 13.5, bold: true, underline: { color: R.C.black }, objectName: 'Bilagans filnamn' });
        slide.addImage({ data: BLANK, x: M, y, w: 3.9, h: .55, objectName: `KIM_ATTACHMENT_${index}`, altText: `Öppna ${file.name}` });
        text(slide, file.descriptionPreview, 4.6, y, 6.55, .55, { fontSize: 12 });
        rule(slide, y + .73);
        objects.push({ slideNumber: count, file, x: M, y, w: 3.9, h: .55 });
      });
    }

    if (m.notes.length) {
      let slide, column = 2, cursor = 0;
      function nextColumn() {
        if (column >= 1) { slide = base('Fördjupade observationer'); column = 0; } else column++;
        cursor = 1.16;
      }
      for (const note of m.notes) {
        const lines = R.wrap(note.text, 5.15, 12.5);
        let start = 0, continued = false;
        do {
          const heading = R.wrap(note.title + (continued ? ' (forts.)' : ''), 5.15, 13);
          const headingHeight = heading.length * .205;
          if (column > 1 || cursor + headingHeight + .4 > 7.48) nextColumn();
          const available = Math.max(1, Math.floor((7.48 - cursor - headingHeight - .08) / .205));
          const part = lines.slice(start, start + available), x = M + column * 5.58;
          text(slide, heading.join('\n'), x, cursor, 5.15, headingHeight + .02, { fontSize: 13, bold: true });
          text(slide, part.join('\n'), x, cursor + headingHeight + .08, 5.15, part.length * .205 + .04, { fontSize: 12.5, objectName: 'Fullständig motivering' });
          cursor += headingHeight + .08 + part.length * .205 + .22;
          start += part.length; continued = true;
          if (start < lines.length) nextColumn();
        } while (start < lines.length);
      }
    }
    return window.KimEmbedDocuments(await pptx.write({ outputType: 'arraybuffer' }), objects);
  }
  window.KimBuildPowerPoint = build;
})();
