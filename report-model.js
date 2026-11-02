/* Shared presentation/print/Excel model. Scoring remains owned by the app. */
(() => {
  'use strict';
  const FONT = 'Gibson', HEADING = 'Gibson Light';
  const C = { black: '000000', white: 'FFFFFF', grey: '6B6B6B', line: 'C6C6C6', light: 'F3F3F2', aqua: '98E9DF', green: '96B33C', orange: 'E67D2C', red: 'DB867D' };
  const titles = { 1: 'Process och avgränsning', 2: 'Lyft per skift', 3: 'Lastens vikt', 4: 'Arbetsförhållanden', 5: 'Arbetsställning', 8: 'Övriga observationer' };
  const xml = value => String(value ?? '').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[c]);
  const number = value => Number.isFinite(Number(value)) ? Number(value) : 0;
  const format = value => number(value).toLocaleString('sv-SE', { maximumFractionDigits: 4 });
  const safeName = value => String(value || 'KIM1').replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_').trim().slice(0, 160) || 'KIM1';

  function riskInfo(value) {
    const risk = number(value);
    if (risk >= 50) return { level: 4, label: 'Hög belastning', range: '50 eller mer', fill: 'F6DEDB', accent: C.red, action: 'Arbetsplatsens utformning måste ändras.' };
    if (risk >= 25) return { level: 3, label: 'Kraftigt ökad belastning', range: '25 till under 50', fill: 'F8E3CD', accent: C.orange, action: 'Ny utformning av arbetsplatsen rekommenderas.' };
    if (risk >= 10) return { level: 2, label: 'Ökad belastning', range: '10 till under 25', fill: 'FFF2BD', accent: 'D3B64C', action: 'Anpassning av arbetsplatsen är fördelaktig för personer med lägre fysisk kapacitet.' };
    return { level: 1, label: 'Låg belastning', range: 'Under 10', fill: 'E9F0D8', accent: C.green, action: 'Fysisk överbelastning är osannolik.' };
  }

  function wrap(text, width, fontSize = 12) {
    const context = document.createElement('canvas').getContext('2d');
    context.font = `${fontSize * 96 / 72}px "${FONT}", Arial, sans-serif`;
    const max = width * 96 * .92;
    const lines = [];
    for (const paragraph of String(text ?? '').replace(/\r/g, '').split('\n')) {
      if (!paragraph.trim()) { lines.push(''); continue; }
      let line = '';
      for (const word of paragraph.trim().split(/\s+/)) {
        if (line && context.measureText(`${line} ${word}`).width > max) { lines.push(line); line = ''; }
        if (context.measureText(word).width <= max) { line += (line ? ' ' : '') + word; continue; }
        for (const char of word) {
          if (line && context.measureText(line + char).width > max) { lines.push(line); line = ''; }
          line += char;
        }
      }
      lines.push(line);
    }
    return lines;
  }

  function preview(text, width, size, count) {
    if (!String(text || '').trim()) return { text: '–', cut: false };
    const lines = wrap(text, width, size);
    if (lines.length <= count) return { text: lines.join('\n'), cut: false };
    const selected = lines.slice(0, count);
    selected[count - 1] = selected[count - 1].replace(/\s+\S*$/, '').slice(0, -1) + '…';
    return { text: selected.join('\n'), cut: true };
  }

  function model(report, evidence) {
    const sections = Object.fromEntries(report.sections.map(s => [s.number, s]));
    const notes = [];
    const factors = [2, 3, 4, 5].map(n => {
      const section = sections[n];
      const comment = preview(section.comment, 4.15, 12, 3);
      if (comment.cut) notes.push({ title: `${n}. ${titles[n]} – motivering`, text: section.comment });
      return { ...section, label: titles[n], commentPreview: comment.text };
    });
    const intro = preview(sections[1]?.comment, 5.14, 11.5, 3);
    const observation = preview(sections[8]?.comment, 5.14, 11.5, 3);
    if (intro.cut) notes.push({ title: '1. Process och avgränsning', text: sections[1].comment });
    if (observation.cut) notes.push({ title: '8. Övriga observationer', text: sections[8].comment });
    const meta = report.metadata.map((value, i) => {
      const widths = [2.0, 4.15, 4.15];
      const shown = preview(value, widths[i], 14, 1);
      if (shown.cut) notes.push({ title: ['Fabrik', 'Process', 'Underprocess'][i], text: value });
      return shown.text;
    });
    const photos = Object.entries(evidence.sections || {}).flatMap(([n, section]) => (section.images || []).map(image => {
      const caption = preview(image.caption || image.name, 5.1, 11.5, 2);
      if (caption.cut) notes.push({ title: `${n}. ${titles[n]} – bildtext (${image.name})`, text: image.caption || image.name });
      return { ...image, number: Number(n), label: titles[n], captionPreview: caption.text, score: sections[n]?.score };
    }));
    const attachments = (evidence.attachments || []).map(file => {
      const description = preview(file.description, 6.6, 12, 2);
      if (description.cut) notes.push({ title: `Bilaga – ${file.name}`, text: file.description });
      return { ...file, descriptionPreview: description.text };
    });
    return { ...report, sections, factors, notes, photos, attachments, meta, intro: intro.text, observation: observation.text, risk: number(report.scores[5]), riskInfo: riskInfo(report.scores[5]), date: new Date().toLocaleDateString('sv-SE', { timeZone: 'Europe/Stockholm' }) };
  }

  let logoPromise;
  function assets() {
    if (!logoPromise) logoPromise = (async () => {
      const source = document.querySelector('.logo');
      if (!source) return {};
      await source.decode();
      const canvas = document.createElement('canvas');
      canvas.width = source.naturalWidth; canvas.height = source.naturalHeight;
      canvas.getContext('2d').drawImage(source, 0, 0);
      return { logoData: canvas.toDataURL('image/png') };
    })();
    return logoPromise;
  }

  function buildPrint(report, evidence) {
    const m = model(report, evidence);
    const e = xml;
    const header = title => `<header class="reportHeader"><h1>${e(title)}</h1><img src="./Dometic_hori_office_neg_logga_vit.png" alt="Dometic"></header>`;
    const footer = `<footer class="reportFooter">KIM 1 &nbsp; ${e(m.date)}</footer>`;
    const summary = `<article class="compactPrintPage">${header('Ergonomianalys KIM 1')}
      <div class="reportMeta">${['Fabrik', 'Process', 'Underprocess'].map((label, i) => `<div><small>${label}</small><strong>${e(m.meta[i])}</strong></div>`).join('')}</div>
      <div class="reportRisk"><div class="reportRiskValue">${format(m.risk)}<small>Riskpoäng</small></div><div><h2>Riskklass ${m.riskInfo.level}: ${e(m.riskInfo.label)}</h2><p>${e(m.riskInfo.action)}</p><small>(${m.scores.slice(0, 3).map(format).join(' + ')}) × ${format(m.scores[4])} = ${format(m.risk)}</small></div></div>
      <table class="reportTable"><thead><tr><th>Faktor</th><th>Vald nivå</th><th>Poäng</th><th>Motivering</th></tr></thead><tbody>${m.factors.map(f => `<tr><td>${e(f.label)}</td><td>${e(f.choice)}</td><td class="reportPoints">${format(f.score)}</td><td>${e(f.commentPreview)}</td></tr>`).join('')}</tbody></table>
      <div class="reportLegend">${['<10 Låg', '10–<25 Ökad', '25–<50 Kraftigt ökad', '≥50 Hög'].map((s, i) => `<span style="background:#${['E9F0D8', 'FFF2BD', 'F8E3CD', 'F6DEDB'][i]}">${e(s)}</span>`).join('')}</div>
      <div class="reportBottom"><div><h3>Process och avgränsning</h3><p>${e(m.intro)}</p></div><div><h3>Övriga observationer</h3><p>${e(m.observation)}</p></div></div>${footer}</article>`;
    const pages = [summary];
    for (let at = 0; at < m.photos.length; at += 4) {
      pages.push(`<article class="compactPrintPage">${header('Bildunderlag')}<div class="reportPhotoGrid">${m.photos.slice(at, at + 4).map(p => `<figure><h3>${p.number}. ${e(p.label)}</h3><img src="${e(p.data)}" alt="${e(p.caption || p.name)}"><figcaption>${e(p.captionPreview)}</figcaption></figure>`).join('')}</div>${footer}</article>`);
    }
    if (m.notes.length) pages.push(`<article class="reportTextPage">${header('Fördjupade observationer')}${m.notes.map(n => `<section><h3>${e(n.title)}</h3><p>${e(n.text)}</p></section>`).join('')}${footer}</article>`);
    if (m.attachments.length) pages.push(`<article class="reportTextPage">${header('Dokumentbilagor')}<table class="reportTable"><thead><tr><th>Dokument</th><th>Beskrivning</th></tr></thead><tbody>${m.attachments.map(a => `<tr><td>${e(a.name)}</td><td>${e(a.description || '–')}</td></tr>`).join('')}</tbody></table>${footer}</article>`);
    document.getElementById('printAnalysis').innerHTML = pages.join('');
    document.getElementById('printSheet').replaceChildren();
  }
  window.KimReport = { FONT, HEADING, C, titles, xml, number, format, safeName, riskInfo, wrap, preview, model, assets, buildPrint };
})();
