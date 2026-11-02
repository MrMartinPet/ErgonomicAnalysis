/* KIM 1: observations, photographs and document attachments. All data stays in this browser. */
(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const numbers = [1, 2, 3, 4, 5, 8];
  const sectionIds = { 1: 'processSection', 2: 'lifts', 3: 'weight', 4: 'condition', 5: 'posture', 8: 'commentsSection' };
  let state = { sections: Object.fromEntries(numbers.map(n => [n, { comment: '', images: [] }])), attachments: [] };
  let database = null;
  let initialized = false;
  let timer;
  let work = Promise.resolve();
  let saveWork = Promise.resolve();
  let resetting = false;
  let pending = 0;
  let revision = 0;
  let lastDownloadUrl = null;
  const uid = () => crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2);
  const sizeText = size => size < 1048576 ? `${Math.ceil(size / 1024)} kB` : `${(size / 1048576).toFixed(1).replace('.', ',')} MB`;
  const make = (tag, className, text) => {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (text !== undefined) node.textContent = text;
    return node;
  };

  function message(text, error = false) {
    $('appMessage').textContent = text;
    $('appMessage').classList.toggle('error', error);
  }

  function download(blob, filename) {
    if (lastDownloadUrl) URL.revokeObjectURL(lastDownloadUrl);
    const url = URL.createObjectURL(blob);
    lastDownloadUrl = url;
    const link = make('a');
    link.href = url;
    link.download = filename;
    link.className = 'btn primary downloadLink';
    link.textContent = 'Ladda ned ' + filename;
    let links = $('downloadLinks');
    if (!links) {
      links = make('div', 'downloadLinks');
      links.id = 'downloadLinks';
      links.dataset.noPrint = '';
      $('appMessage').after(links);
    }
    links.replaceChildren(link);
    link.click();
  }

  function makeUploader(text, id, accept, handler) {
    const label = make('label', 'btn secondary uploadLabel');
    label.htmlFor = id;
    const input = make('input', 'uploadInput');
    input.type = 'file';
    input.id = id;
    input.multiple = true;
    input.accept = accept;
    input.addEventListener('change', () => {
      const files = Array.from(input.files);
      input.value = '';
      enqueue(() => handler(files));
    });
    label.append(input, make('span', '', text));
    return label;
  }

  for (const n of numbers) {
    const panel = make('div', 'evidence');
    panel.dataset.question = n;
    if (n !== 8) {
      const label = make('label', 'evidenceLabel', n === 1 ? 'Motivering / kommentar' : 'Motivering till vald poäng');
      label.htmlFor = 'comment-' + n;
      const textarea = make('textarea');
      textarea.id = 'comment-' + n;
      textarea.placeholder = n === 1 ? 'Beskriv arbetsmomentet och bakgrunden till bedömningen…' : 'Beskriv varför ni valt denna poäng. Ange gärna mätningar eller konkreta observationer…';
      textarea.addEventListener('input', () => { state.sections[n].comment = textarea.value; changed(); });
      panel.append(label, textarea);
    }
    const tools = make('div', 'evidenceTools');
    tools.dataset.noPrint = '';
    tools.append(makeUploader('Lägg till bilder', 'imageInput-' + n, 'image/jpeg,image/png,image/webp,image/gif,image/heic,image/heif', files => addImages(n, files)));
    tools.append(make('span', 'evidenceHint', 'Flera bilder kan väljas. Lägg gärna till en bildtext.'));
    const grid = make('div', 'imageGrid');
    grid.id = 'images-' + n;
    panel.append(tools, grid);
    $(sectionIds[n]).append(panel);
  }

  const attachments = make('div', 'attachmentArea');
  attachments.append(make('h4', '', 'Dokumentbilagor'));
  const hint = make('p', 'evidenceHint', 'Lägg till exempelvis instruktioner, mätningar eller tidigare analyser. Dokumenten följer med inuti PowerPoint-filen. Högst 20 MB per dokument och 50 MB sammanlagt.');
  hint.dataset.noPrint = '';
  const attachmentTools = make('div', 'evidenceTools');
  attachmentTools.dataset.noPrint = '';
  attachmentTools.append(makeUploader('Bifoga dokument', 'attachmentInput', '.pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.txt,.csv,.rtf,.odt,.ods,.odp,.zip', addAttachments));
  const attachmentList = make('div', 'attachmentList');
  attachmentList.id = 'attachmentList';
  attachments.append(hint, attachmentTools, attachmentList);
  $('commentsSection').append(attachments);

  function renderImages(n) {
    const grid = $('images-' + n);
    grid.replaceChildren();
    state.sections[n].images.forEach(image => {
      const figure = make('figure', 'evidenceImage');
      const img = make('img');
      img.src = image.data;
      img.alt = image.caption || image.name;
      const caption = make('figcaption');
      caption.append(make('span', 'imageName', image.name));
      const captionInput = make('input');
      captionInput.value = image.caption || '';
      captionInput.placeholder = 'Bildtext / observation';
      captionInput.setAttribute('aria-label', 'Bildtext för ' + image.name);
      captionInput.addEventListener('input', () => { image.caption = captionInput.value; img.alt = image.caption || image.name; changed(); });
      const actions = make('div', 'imageActions');
      actions.dataset.noPrint = '';
      const remove = make('button', 'btn secondary', 'Ta bort bild');
      remove.type = 'button';
      remove.setAttribute('aria-label', 'Ta bort ' + image.name);
      remove.onclick = () => { state.sections[n].images = state.sections[n].images.filter(item => item.id !== image.id); renderImages(n); changed(); };
      actions.append(remove);
      caption.append(captionInput, actions);
      figure.append(img, caption);
      grid.append(figure);
    });
  }

  function renderAttachments() {
    attachmentList.replaceChildren();
    state.attachments.forEach(file => {
      const row = make('div', 'attachmentItem');
      const info = make('div');
      info.append(make('div', 'attachmentName', file.name), make('div', 'attachmentSize', sizeText(file.bytes.byteLength)));
      const description = make('input');
      description.value = file.description || '';
      description.placeholder = 'Kort beskrivning av bilagan';
      description.setAttribute('aria-label', 'Beskrivning för ' + file.name);
      description.addEventListener('input', () => { file.description = description.value; changed(); });
      info.append(description);
      const actions = make('div', 'attachmentActions');
      actions.dataset.noPrint = '';
      const save = make('button', 'btn secondary', 'Ladda ned');
      save.type = 'button';
      save.onclick = () => download(new Blob([file.bytes], { type: file.type }), file.name);
      const remove = make('button', 'btn secondary', 'Ta bort');
      remove.type = 'button';
      remove.setAttribute('aria-label', 'Ta bort bilagan ' + file.name);
      remove.onclick = () => { state.attachments = state.attachments.filter(item => item.id !== file.id); renderAttachments(); changed(); };
      actions.append(save, remove);
      row.append(info, actions);
      attachmentList.append(row);
    });
  }

  function enqueue(task) {
    pending++;
    $('draftState').textContent = 'Läser in filer…';
    work = work.then(() => ready).then(task).catch(error => message(error.message || 'Kunde inte läsa filen.', true)).finally(() => {
      pending--;
      if (!pending) changed();
    });
    return work;
  }

  async function readImage(file) {
    if (file.size > 20 * 1048576) throw new Error(`${file.name}: bilden är större än 20 MB.`);
    const url = URL.createObjectURL(file);
    try {
      const img = new Image();
      img.src = url;
      await img.decode();
      const scale = Math.min(1, 2000 / Math.max(img.naturalWidth, img.naturalHeight));
      const canvas = document.createElement('canvas');
      canvas.width = Math.max(1, Math.round(img.naturalWidth * scale));
      canvas.height = Math.max(1, Math.round(img.naturalHeight * scale));
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      return { id: uid(), name: file.name, caption: '', width: canvas.width, height: canvas.height, data: canvas.toDataURL('image/jpeg', .88) };
    } catch {
      throw new Error(`${file.name}: bildformatet kunde inte läsas. Använd JPG, PNG eller WebP.`);
    } finally { URL.revokeObjectURL(url); }
  }

  async function addImages(n, files) {
    const errors = [];
    for (const file of files) {
      if (state.sections[n].images.length >= 20) { errors.push(`Fråga ${n}: högst 20 bilder per fråga.`); break; }
      try { state.sections[n].images.push(await readImage(file)); } catch (error) { errors.push(error.message); }
    }
    renderImages(n);
    message(errors.join(' '), errors.length > 0);
  }

  async function addAttachments(files) {
    const errors = [];
    for (const file of files) {
      if (!/\.(pdf|docx?|xlsx?|pptx?|txt|csv|rtf|odt|ods|odp|zip)$/i.test(file.name)) { errors.push(`${file.name}: välj ett dokument, exempelvis PDF, Word eller Excel.`); continue; }
      const total = state.attachments.reduce((sum, item) => sum + item.bytes.byteLength, 0);
      if (file.size > 20 * 1048576 || total + file.size > 50 * 1048576) { errors.push(`${file.name}: högst 20 MB per dokument och 50 MB sammanlagt.`); continue; }
      if (state.attachments.length >= 30) { errors.push('Högst 30 dokumentbilagor.'); break; }
      state.attachments.push({ id: uid(), name: file.name, type: file.type || 'application/octet-stream', bytes: new Uint8Array(await file.arrayBuffer()), description: '' });
    }
    renderAttachments();
    message(errors.join(' '), errors.length > 0);
  }

  function changed() {
    if (!initialized || resetting) return;
    revision++;
    clearTimeout(timer);
    $('draftState').textContent = pending ? 'Läser in filer…' : 'Sparar analysen i denna webbläsare…';
    timer = setTimeout(save, 450);
  }

  function transaction(mode, operation) {
    return new Promise((resolve, reject) => {
      const tx = database.transaction('drafts', mode);
      const request = operation(tx.objectStore('drafts'));
      let result;
      if (request) request.onsuccess = () => { result = request.result; };
      tx.oncomplete = () => resolve(result);
      tx.onerror = () => reject(tx.error || new Error('Lagringen misslyckades.'));
      tx.onabort = () => reject(tx.error || new Error('Lagringen avbröts.'));
    });
  }

  function save() {
    clearTimeout(timer);
    if (!initialized || resetting || !database) return saveWork;
    const savedRevision = revision;
    const saved = { version: 1, core: window.KimAnalysis.snapshot(), evidence: structuredClone(state) };
    saveWork = saveWork.then(() => transaction('readwrite', store => store.put(saved, 'current'))).then(() => {
      if (!pending && revision === savedRevision) {
        $('draftState').textContent = 'Analysen sparas automatiskt i denna webbläsare.';
        window.KimAnalysis.clean();
      }
    }).catch(() => {
      $('draftState').textContent = 'Analysen kunde inte sparas i webbläsaren. Exportera innan du stänger sidan.';
    });
    return saveWork;
  }

  const ready = (async () => {
    try {
      database = await new Promise((resolve, reject) => {
        const request = indexedDB.open('dometic-kim1-analysis', 1);
        request.onupgradeneeded = () => request.result.createObjectStore('drafts');
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
        request.onblocked = () => reject(new Error('Lagringen är blockerad.'));
      });
      const saved = await transaction('readonly', store => store.get('current'));
      if (saved?.version === 1) {
        window.KimAnalysis.restore(saved.core);
        for (const n of numbers) {
          const section = saved.evidence?.sections?.[n];
          if (section) state.sections[n] = section;
          if (n !== 8) $('comment-' + n).value = state.sections[n].comment || '';
          renderImages(n);
        }
        state.attachments = saved.evidence?.attachments || [];
        renderAttachments();
      }
      $('draftState').textContent = saved ? 'Den senast sparade analysen har öppnats. Ändringar sparas i denna webbläsare.' : 'Analysen sparas automatiskt i denna webbläsare.';
    } catch {
      database = null;
      $('draftState').textContent = 'Automatisk lagring är inte tillgänglig. Exportera innan du stänger sidan.';
    } finally { initialized = true; }
  })();

  document.addEventListener('input', event => {
    if (event.target.matches('input:not([type=file]), textarea') && event.target.closest('.screen > .section')) changed();
  });
  document.addEventListener('click', event => { if (event.target.closest('.option')) changed(); });
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') save(); });
  document.addEventListener('focusout', event => { if (event.target.matches('input, textarea')) save(); });
  window.addEventListener('pagehide', save);

  window.KimEvidence = {
    ready,
    message,
    download,
    getState: () => state,
    async flush() { await ready; await work; await save(); },
    async reset() {
      resetting = true;
      clearTimeout(timer);
      await ready;
      await work;
      await saveWork;
      if (database) await transaction('readwrite', store => store.delete('current'));
    }
  };
})();
