(function () {
  'use strict';

  const { XlsxFile, DocxFile, serialToDate } = window.OOXML;
  const FORMS = window.FORMS;
  const CLIENTS = window.CLIENTS;
  const formById = (id) => FORMS.find((f) => f.id === id);

  const $ = (sel, el = document) => el.querySelector(sel);
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const app = $('#app');

  // ---------- opslag ----------
  const LS_DRAFTS = 'dael.drafts.v1';
  const LS_MAIL = 'dael.mail.v1';

  function lsGet(key, fallback) {
    try {
      const v = localStorage.getItem(key);
      return v ? JSON.parse(v) : fallback;
    } catch (e) {
      return fallback;
    }
  }
  function lsSet(key, value) {
    try {
      localStorage.setItem(key, JSON.stringify(value));
      return true;
    } catch (e) {
      return false;
    }
  }

  // Geïmporteerde bestanden bewaren we in IndexedDB, zodat foto's/schetsen in het bestand behouden blijven
  const idb = (() => {
    let dbp = null;
    const open = () => {
      if (!dbp) {
        dbp = new Promise((res, rej) => {
          try {
            const req = indexedDB.open('dael-rapportage', 1);
            req.onupgradeneeded = () => req.result.createObjectStore('files');
            req.onsuccess = () => res(req.result);
            req.onerror = () => rej(req.error);
          } catch (e) {
            rej(e);
          }
        });
      }
      return dbp;
    };
    const tx = async (mode, fn) => {
      const db = await open();
      return new Promise((res, rej) => {
        const t = db.transaction('files', mode);
        const r = fn(t.objectStore('files'));
        t.oncomplete = () => res(r && r.result);
        t.onerror = () => rej(t.error);
      });
    };
    return {
      put: (k, v) => tx('readwrite', (s) => s.put(v, k)).catch(() => null),
      get: (k) => tx('readonly', (s) => s.get(k)).catch(() => null),
      del: (k) => tx('readwrite', (s) => s.delete(k)).catch(() => null),
    };
  })();

  // ---------- formulieren voorbereiden ----------
  function allFields(form) {
    const out = [];
    for (const sec of form.sections) {
      for (const it of sec.items) {
        if (it.grid) it.rows.forEach((r) => r.fields.forEach((f) => f && out.push(f)));
        else out.push(it);
      }
    }
    return out;
  }

  const templateCache = {};
  function loadTemplate(form) {
    if (!templateCache[form.template]) {
      templateCache[form.template] = fetch(form.template).then((r) => {
        if (!r.ok) throw new Error('Sjabloon kon niet geladen worden (' + r.status + ').');
        return r.arrayBuffer();
      });
    }
    return templateCache[form.template].then((b) => b.slice(0));
  }

  const openDoc = (form, buf) => (form.kind === 'xlsx' ? XlsxFile.load(buf) : DocxFile.load(buf));

  const prepared = {};
  async function prepare(form) {
    if (prepared[form.id]) return prepared[form.id];
    const doc = await openDoc(form, await loadTemplate(form));
    const fields = allFields(form);
    for (const f of fields) {
      if (form.kind === 'xlsx') {
        const ref = f.cell || f.cells[0];
        if (f.cell) f.auto = (await doc.read(form.sheet, f.cell)).formula;
        const key = window.VF_DV[form.sheet] && window.VF_DV[form.sheet][ref];
        if (key && !f.options) f.options = window.VF_LISTS[key];
      } else if (f.p) {
        const cp = doc.cellParas(f.p[0], f.p[1], f.p[2]);
        const para = cp && cp.paras[f.p[3]];
        f.ph = para ? doc.paraText(para) : '';
      }
    }
    const defaults = await readValues(form, doc);
    prepared[form.id] = { fields, defaults };
    return prepared[form.id];
  }

  // ---------- lezen / schrijven ----------
  const PLACEHOLDER = /…|\.\.\.\./;
  const norm = (s) => String(s || '').replace(/\s+/g, ' ').trim();

  function cleanDocxText(f, text) {
    if (f.ph && PLACEHOLDER.test(f.ph) && norm(text) === norm(f.ph)) return '';
    let s = text.trim();
    const pre = (f.prefix || '').trim();
    const suf = (f.suffix || '').trim();
    if (pre && s.startsWith(pre)) s = s.slice(pre.length).trim();
    if (suf && s.endsWith(suf)) s = s.slice(0, -suf.length).trim();
    if (/^[\s.…_°]*$/.test(s)) s = '';
    return s;
  }

  function toIsoDate(v) {
    const m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/.exec(v || '');
    return m ? `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}` : '';
  }
  function fromIsoDate(v) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v || '');
    return m ? `${m[3]}-${m[2]}-${m[1]}` : v;
  }

  async function readValues(form, doc) {
    const v = {};
    for (const f of allFields(form)) {
      if (form.kind === 'xlsx') {
        if (f.cells) {
          const parts = [];
          for (const c of f.cells) parts.push((await doc.read(form.sheet, c)).value);
          while (parts.length && !parts[parts.length - 1].trim()) parts.pop();
          v[f.id] = parts.join('\n');
        } else {
          const r = await doc.read(form.sheet, f.cell);
          let val = r.value;
          if (f.type === 'date' && r.numeric && /^\d+(\.\d+)?$/.test(val)) val = serialToDate(parseFloat(val));
          v[f.id] = val;
        }
      } else if (f.type === 'check') {
        v[f.id] = doc.getCheck(f.cb);
      } else if (f.type === 'choice') {
        v[f.id] = f.opts.filter(([, n]) => doc.getCheck(n)).map(([l]) => l);
      } else if (f.rest) {
        const [t, r, c, from, to] = f.rest;
        const cp = doc.cellParas(t, r, c);
        if (!cp) { v[f.id] = ''; continue; }
        const last = to == null ? cp.paras.length - 1 : Math.min(to, cp.paras.length - 1);
        const lines = [];
        for (let i = from; i <= last; i++) lines.push(doc.paraText(cp.paras[i]));
        v[f.id] = lines.join('\n').replace(/^\s*\n/, '').replace(/\s+$/, '');
      } else if (f.p) {
        const cp = doc.cellParas(f.p[0], f.p[1], f.p[2]);
        const para = cp && cp.paras[f.p[3]];
        v[f.id] = para ? cleanDocxText(f, doc.paraText(para)) : '';
      }
    }
    return v;
  }

  function wrapLines(text, n, width) {
    const lines = [];
    for (const raw of String(text || '').split('\n')) {
      let s = raw;
      while (s.length > width) {
        let cut = s.lastIndexOf(' ', width);
        if (cut < width * 0.5) cut = width;
        lines.push(s.slice(0, cut).trimEnd());
        s = s.slice(cut).trimStart();
      }
      lines.push(s);
    }
    if (lines.length > n) {
      const head = lines.slice(0, n - 1);
      head.push(lines.slice(n - 1).join('\n'));
      return head;
    }
    while (lines.length < n) lines.push('');
    return lines;
  }

  async function writeValues(form, doc, values) {
    for (const f of allFields(form)) {
      const val = values[f.id];
      if (form.kind === 'xlsx') {
        if (f.auto) continue;
        if (f.cells) {
          const parts = wrapLines(val, f.cells.length, f.width || 95);
          for (let i = 0; i < f.cells.length; i++) await doc.write(form.sheet, f.cells[i], parts[i]);
        } else {
          await doc.write(form.sheet, f.cell, val == null ? '' : String(val).trim());
        }
      } else if (f.type === 'check') {
        doc.setCheck(f.cb, !!val);
      } else if (f.type === 'choice') {
        const sel = Array.isArray(val) ? val : [];
        for (const [l, n] of f.opts) doc.setCheck(n, sel.includes(l));
      } else if (f.rest) {
        const [t, r, c, from, to] = f.rest;
        const cp = doc.cellParas(t, r, c);
        if (!cp) continue;
        doc.ensurePara(cp.tc, cp.paras, from);
        const last = to == null ? cp.paras.length - 1 : Math.min(to, cp.paras.length - 1);
        doc.setParaText(cp.paras[from], String(val || '').replace(/\s+$/, ''));
        for (let i = from + 1; i <= last; i++) doc.setParaText(cp.paras[i], '');
      } else if (f.p) {
        const cp = doc.cellParas(f.p[0], f.p[1], f.p[2]);
        if (!cp) continue;
        const para = doc.ensurePara(cp.tc, cp.paras, f.p[3]);
        const s = String(val == null ? '' : val).trim();
        let text;
        if (s) text = (f.prefix || '') + s + (f.suffix || '');
        else text = f.ph && PLACEHOLDER.test(f.ph) ? f.ph : ((f.prefix || '') + (f.suffix || '')).trim();
        doc.setParaText(para, text);
      }
    }
  }

  async function buildFile(draft) {
    const form = formById(draft.formId);
    await prepare(form);
    let base = null;
    if (draft.hasBase) base = await idb.get('base:' + draft.id);
    const buf = base || (await loadTemplate(form));
    const doc = await openDoc(form, buf);
    await writeValues(form, doc, draft.values);
    const blob = await doc.toBlob();
    return { blob, name: sanitizeName(form.fileName(draft.values)) };
  }

  const sanitizeName = (n) => n.replace(/[\\/:*?"<>|]+/g, '-').replace(/\s+/g, ' ').trim();

  // ---------- concepten ----------
  let drafts = lsGet(LS_DRAFTS, []);
  const saveDrafts = () => {
    if (!lsSet(LS_DRAFTS, drafts)) toast('Let op: opslaan op dit apparaat is niet gelukt.', 'warn');
  };
  const newId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);

  function draftTitle(d) {
    const form = formById(d.formId);
    const t = form && d.values[form.titleField];
    return t ? String(t) : 'Zonder sitenummer';
  }

  // ---------- voortgang ----------
  function isFilled(f, v) {
    if (f.auto) return null;
    if (f.type === 'check') return null; // los vinkje: niet aanvinken is ook een antwoord
    if (Array.isArray(v)) return v.length > 0;
    return v != null && String(v).trim() !== '';
  }

  function sectionStats(sec, values) {
    let total = 0;
    let done = 0;
    const fields = [];
    for (const it of sec.items) {
      if (it.grid) it.rows.forEach((r) => r.fields.forEach((f) => f && fields.push(f)));
      else fields.push(it);
    }
    for (const f of fields) {
      const s = isFilled(f, values[f.id]);
      if (s === null) continue;
      total++;
      if (s) done++;
    }
    return { total, done };
  }

  function formStats(form, values) {
    let total = 0;
    let done = 0;
    for (const sec of form.sections) {
      const s = sectionStats(sec, values);
      total += s.total;
      done += s.done;
    }
    return { total, done, pct: total ? Math.round((done / total) * 100) : 0 };
  }

  // ---------- toast ----------
  let toastTimer = null;
  function toast(msg, kind = 'ok') {
    let el = $('#toast');
    if (!el) {
      el = document.createElement('div');
      el.id = 'toast';
      el.setAttribute('role', 'status');
      document.body.appendChild(el);
    }
    el.className = 'toast show ' + kind;
    el.textContent = msg;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => el.classList.remove('show'), 3600);
  }

  // ---------- routering ----------
  const state = { view: 'home', draft: null, section: 0, finishFile: null };

  function go(view, opts = {}) {
    Object.assign(state, { view }, opts);
    render();
    window.scrollTo({ top: 0 });
  }

  function render() {
    if (state.view === 'home') renderHome();
    else if (state.view === 'kies') renderChooseForm();
    else if (state.view === 'form') renderForm();
    else if (state.view === 'finish') renderFinish();
  }

  // ---------- iconen ----------
  const ICON = {
    upload: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 16V4m0 0-5 5m5-5 5 5M4 16v3a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-3" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    arrow: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 12h14m-6-6 6 6-6 6" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    back: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M19 12H5m6 6-6-6 6-6" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    trash: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 7h16M10 11v6m4-6v6M6 7l1 12a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1l1-12M9 7V4h6v3" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    download: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 4v12m0 0 5-5m-5 5-5-5M4 20h16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    mail: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 6h18v12H3zM3 7l9 6 9-6" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/></svg>',
    check: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="m5 12 5 5 9-10" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    xls: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 3h9l5 5v13H6z" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/><path d="m9 11 5 6m0-6-5 6" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>',
    doc: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 3h9l5 5v13H6z" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/><path d="M9 12h7M9 15.5h7M9 19h4" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>',
    home: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 11 12 4l8 7v9h-5v-6H9v6H4z" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/></svg>',
  };

  function header(extra = '') {
    return `<header class="topbar">
      <button class="brand" data-act="home" aria-label="Naar start">
        <span class="wordmark">DAEL</span><span class="brand-sub">Site rapportage</span>
      </button>
      <div class="topbar-extra">${extra}</div>
    </header>`;
  }

  // ---------- startscherm ----------
  function renderHome() {
    const list = drafts.slice().sort((a, b) => b.updated - a.updated);
    app.innerHTML = `${header()}
    <main class="wrap">
      <section class="hero">
        <p class="eyebrow">Nieuw rapport</p>
        <h1>Voor wie vul je het rapport in?</h1>
      </section>
      <div class="client-grid">
        ${CLIENTS.map((c) => `
          <button class="client-card client-${c.id}" data-act="client" data-id="${c.id}">
            <span class="client-logo">${c.id === 'vodafone' ? ICON.xls : ICON.doc}</span>
            <span class="client-text"><span class="client-name">${esc(c.name)}</span><span class="client-desc">${esc(c.desc)}</span></span>
            <span class="client-go">${ICON.arrow}</span>
          </button>`).join('')}
      </div>

      <section class="block">
        <h2>Verder met een bestaand bestand</h2>
        <label class="dropzone" id="dropzone">
          <input type="file" id="importFile" accept=".xlsx,.docx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.openxmlformats-officedocument.wordprocessingml.document" hidden>
          <span class="dz-icon">${ICON.upload}</span>
          <span><strong>Importeer een (half) ingevulde Excel of Word</strong><br><span class="muted">Het Vodafone TSS-rapport of KPN RCT-rapport. Antwoorden worden ingelezen zodat je verder kunt.</span></span>
        </label>
      </section>

      ${list.length ? `
      <section class="block">
        <h2>Concepten op dit apparaat</h2>
        <ul class="draft-list">
          ${list.map((d) => {
            const form = formById(d.formId);
            if (!form) return '';
            const st = formStats(form, d.values);
            return `<li class="draft">
              <button class="draft-main" data-act="open-draft" data-id="${d.id}">
                <span class="draft-badge badge-${form.klant}">${form.klant === 'kpn' ? 'KPN' : 'Vodafone'}</span>
                <span class="draft-info">
                  <span class="draft-title">${esc(draftTitle(d))}</span>
                  <span class="draft-meta">${esc(form.title)} · ${new Date(d.updated).toLocaleString('nl-NL', { dateStyle: 'short', timeStyle: 'short' })}</span>
                  <span class="bar"><span style="width:${st.pct}%"></span></span>
                </span>
                <span class="draft-pct">${st.pct}%</span>
              </button>
              <button class="icon-btn" data-act="del-draft" data-id="${d.id}" aria-label="Concept verwijderen">${ICON.trash}</button>
            </li>`;
          }).join('')}
        </ul>
      </section>` : ''}
      <p class="foot muted">Je antwoorden worden automatisch op dit apparaat bewaard. Er wordt niets naar een server gestuurd.</p>
    </main>`;

    const input = $('#importFile');
    input.addEventListener('change', () => input.files[0] && importFile(input.files[0]));
    const dz = $('#dropzone');
    ['dragenter', 'dragover'].forEach((e) => dz.addEventListener(e, (ev) => { ev.preventDefault(); dz.classList.add('drag'); }));
    ['dragleave', 'drop'].forEach((e) => dz.addEventListener(e, (ev) => { ev.preventDefault(); dz.classList.remove('drag'); }));
    dz.addEventListener('drop', (ev) => ev.dataTransfer.files[0] && importFile(ev.dataTransfer.files[0]));
  }

  function renderChooseForm() {
    const client = CLIENTS.find((c) => c.id === state.client);
    app.innerHTML = `${header()}
    <main class="wrap">
      <button class="link-back" data-act="home">${ICON.back} Terug</button>
      <section class="hero">
        <p class="eyebrow">${esc(client.name)}</p>
        <h1>Welk rapport?</h1>
      </section>
      <div class="client-grid">
        ${client.forms.map((id) => {
          const f = formById(id);
          return `<button class="client-card client-${f.klant}" data-act="new" data-id="${f.id}">
            <span class="client-logo">${f.kind === 'xlsx' ? ICON.xls : ICON.doc}</span>
            <span class="client-text"><span class="client-name">${esc(f.title)}</span><span class="client-desc">${esc(f.subtitle)}</span></span>
            <span class="client-go">${ICON.arrow}</span>
          </button>`;
        }).join('')}
      </div>
    </main>`;
  }

  async function startNew(formId) {
    const form = formById(formId);
    busy(true, 'Formulier laden…');
    try {
      const { fields, defaults } = await prepare(form);
      const values = {};
      for (const f of fields) {
        if (form.startEmpty) values[f.id] = f.type === 'check' ? false : f.type === 'choice' ? [] : f.auto ? defaults[f.id] : '';
        else values[f.id] = JSON.parse(JSON.stringify(defaults[f.id] == null ? '' : defaults[f.id]));
      }
      const d = { id: newId(), formId, values, created: Date.now(), updated: Date.now(), hasBase: false };
      drafts.push(d);
      saveDrafts();
      go('form', { draft: d, section: 0 });
    } catch (e) {
      console.error(e);
      toast(e.message || 'Formulier kon niet geladen worden.', 'err');
    } finally {
      busy(false);
    }
  }

  // ---------- importeren ----------
  async function importFile(file) {
    busy(true, 'Bestand inlezen…');
    try {
      const buf = await file.arrayBuffer();
      const zip = await JSZip.loadAsync(buf.slice(0));
      let form = null;
      let values = null;
      if (zip.file('xl/workbook.xml')) {
        const doc = await XlsxFile.load(buf.slice(0));
        const candidates = FORMS.filter((f) => f.kind === 'xlsx' && doc.hasSheet(f.sheet));
        if (!candidates.length) throw new Error('Dit Excel-bestand lijkt niet op het Vodafone TSS-rapport.');
        // Kies het tabblad waarin het meest is ingevuld ten opzichte van het lege sjabloon
        let best = -1;
        for (const f of candidates) {
          const { defaults } = await prepare(f);
          const vals = await readValues(f, doc);
          const changed = Object.keys(vals).filter((k) => norm(vals[k]) !== norm(defaults[k])).length;
          if (changed > best) { best = changed; form = f; values = vals; }
        }
      } else if (zip.file('word/document.xml')) {
        const doc = await DocxFile.load(buf.slice(0));
        form = FORMS.find((f) => f.kind === 'docx' && f.detect(doc));
        if (!form) throw new Error('Dit Word-bestand lijkt niet op het KPN RCT-rapport.');
        await prepare(form);
        values = await readValues(form, doc);
      } else {
        throw new Error('Kies een .xlsx- of .docx-bestand.');
      }
      const d = { id: newId(), formId: form.id, values, created: Date.now(), updated: Date.now(), hasBase: true, sourceName: file.name };
      await idb.put('base:' + d.id, buf);
      d.hasBase = !!(await idb.get('base:' + d.id));
      drafts.push(d);
      saveDrafts();
      toast(`Ingelezen als ${form.title} (${form.klant === 'kpn' ? 'KPN' : 'Vodafone'}).`);
      go('form', { draft: d, section: 0 });
    } catch (e) {
      console.error(e);
      toast(e.message || 'Bestand kon niet gelezen worden.', 'err');
    } finally {
      busy(false);
    }
  }

  // ---------- formulier ----------
  function optionsFor(f) {
    return Array.isArray(f.options) ? f.options : null;
  }
  const chipMode = (opts) => opts && opts.length <= 6 && opts.every((o) => String(o).length <= 22);

  function fieldControl(f, v, compact) {
    const id = 'f_' + f.id;
    if (f.auto) {
      return `<div class="auto" title="Wordt door Excel berekend">${v ? esc(v) : 'automatisch'}</div>`;
    }
    if (f.type === 'check') {
      return `<label class="switch${compact ? ' compact' : ''}"><input type="checkbox" id="${id}" data-f="${f.id}" ${v ? 'checked' : ''}><span class="track"><span class="thumb"></span></span>${compact ? '' : '<span class="switch-label">Ja</span>'}</label>`;
    }
    if (f.type === 'choice') {
      const sel = Array.isArray(v) ? v : [];
      return `<div class="chips" role="group">${f.opts.map(([l]) => `<button type="button" class="chip${sel.includes(l) ? ' on' : ''}" data-choice="${f.id}" data-val="${esc(l)}" aria-pressed="${sel.includes(l)}">${f.multi ? '<span class="box"></span>' : ''}${esc(l)}</button>`).join('')}</div>${f.multi ? '<div class="hint">Meerdere antwoorden mogelijk</div>' : ''}`;
    }
    if (f.type === 'area') {
      return `<textarea id="${id}" data-f="${f.id}" rows="${f.rows || 3}" placeholder="Typ hier…">${esc(v)}</textarea>`;
    }
    if (f.type === 'date') {
      return `<input type="date" id="${id}" data-f="${f.id}" data-date="1" value="${esc(toIsoDate(v))}">`;
    }
    const opts = optionsFor(f);
    if (opts && chipMode(opts) && !compact) {
      const extra = v && !opts.includes(v) ? [v] : [];
      return `<div class="chips" role="group">${[...opts, ...extra].map((o) => `<button type="button" class="chip${o === v ? ' on' : ''}" data-single="${f.id}" data-val="${esc(o)}" aria-pressed="${o === v}">${esc(o)}</button>`).join('')}</div>`;
    }
    if (opts) {
      const extra = v && !opts.includes(v) ? `<option value="${esc(v)}">${esc(v)}</option>` : '';
      if (opts.length <= 8) {
        return `<select id="${id}" data-f="${f.id}"><option value=""></option>${extra}${opts.map((o) => `<option${o === v ? ' selected' : ''}>${esc(o)}</option>`).join('')}</select>`;
      }
      const listId = 'dl_' + f.id;
      return `<input type="text" id="${id}" data-f="${f.id}" list="${listId}" value="${esc(v)}" autocomplete="off" placeholder="Kies of typ…"><datalist id="${listId}">${opts.map((o) => `<option value="${esc(o)}">`).join('')}</datalist>`;
    }
    const im = f.inputmode ? ` inputmode="${f.inputmode}"` : '';
    const type = f.inputmode === 'email' ? 'email' : f.inputmode === 'tel' ? 'tel' : 'text';
    return `<input type="${type}" id="${id}" data-f="${f.id}" value="${esc(v)}"${im} autocomplete="off">`;
  }

  function fieldBlock(f, values) {
    const v = values[f.id];
    const filled = isFilled(f, v);
    const forAttr = ['text', 'area', 'date'].includes(f.type) && !(optionsFor(f) && chipMode(optionsFor(f))) ? ` for="f_${f.id}"` : '';
    return `<div class="field${filled ? ' filled' : ''}" id="blk_${f.id}">
      <label class="label"${forAttr}>${esc(f.label)}${filled ? `<span class="tick">${ICON.check}</span>` : ''}</label>
      ${f.hint ? `<div class="hint">${esc(f.hint)}</div>` : ''}
      ${fieldControl(f, v, false)}
    </div>`;
  }

  function gridBlock(g, values) {
    return `<div class="gridblock">
      <div class="grid-title">${esc(g.title)}</div>
      ${g.hint ? `<div class="hint">${esc(g.hint)}</div>` : ''}
      <div class="gtable" style="--cols:${g.cols.length}">
        <div class="grow ghead"><div class="gcell glabel"></div>${g.cols.map((c) => `<div class="gcell">${esc(c)}</div>`).join('')}</div>
        ${g.rows.map((r) => `<div class="grow" id="grow_${esc(r.fields.find(Boolean)?.id || '')}">
          <div class="gcell glabel">${esc(r.label)}</div>
          ${r.fields.map((f, i) => `<div class="gcell${f ? '' : ' empty'}">${f ? `<span class="gmini">${esc(g.cols[i])}</span>${fieldControl(f, values[f.id], true)}` : ''}</div>`).join('')}
        </div>`).join('')}
      </div>
    </div>`;
  }

  function renderForm() {
    const d = state.draft;
    const form = formById(d.formId);
    const secs = form.sections;
    const si = Math.max(0, Math.min(state.section, secs.length - 1));
    const sec = secs[si];
    const st = formStats(form, d.values);
    const isLast = si === secs.length - 1;

    app.innerHTML = `<div class="stick">${header(`<button class="ghost-btn" data-act="download" title="Download tussentijds">${ICON.download}<span>Download</span></button>`)}
    <div class="formbar">
      <div class="formbar-inner">
        <div class="formbar-title">
          <span class="draft-badge badge-${form.klant}">${form.klant === 'kpn' ? 'KPN' : 'Vodafone'}</span>
          <span class="ft">${esc(form.title)}</span>
          <span class="fsite">${esc(draftTitle(d))}</span>
        </div>
        <div class="progress"><div class="bar big"><span style="width:${st.pct}%"></span></div><span class="pct">${st.done}/${st.total}</span></div>
      </div>
    </div></div>
    <div class="layout">
      <nav class="secnav" aria-label="Onderdelen">
        <select class="secselect" data-act="jump" aria-label="Kies onderdeel">
          ${secs.map((s, i) => { const ss = sectionStats(s, d.values); return `<option value="${i}"${i === si ? ' selected' : ''}>${i + 1}/${secs.length} · ${esc(s.title)} (${ss.done}/${ss.total})</option>`; }).join('')}
        </select>
        <ol class="seclist">
          ${secs.map((s, i) => {
            const ss = sectionStats(s, d.values);
            const cls = i === si ? 'cur' : ss.total && ss.done === ss.total ? 'done' : ss.done ? 'part' : '';
            return `<li><button class="${cls}" data-act="section" data-i="${i}"><span class="dot"></span><span class="sn">${esc(s.title)}</span><span class="sc">${ss.done}/${ss.total}</span></button></li>`;
          }).join('')}
          <li><button data-act="finish" class="finish-link"><span class="dot"></span><span class="sn">Afronden &amp; versturen</span></button></li>
        </ol>
      </nav>
      <main class="formmain">
        <h1 class="sectitle">${esc(sec.title)}</h1>
        ${sec.hint ? `<p class="sechint">${esc(sec.hint)}</p>` : ''}
        <div class="fields">
          ${sec.items.map((it) => (it.grid ? gridBlock(it, d.values) : fieldBlock(it, d.values))).join('')}
        </div>
        <div class="pager">
          ${si > 0 ? `<button class="btn secondary" data-act="section" data-i="${si - 1}">${ICON.back} Vorige</button>` : '<span></span>'}
          ${isLast ? `<button class="btn primary" data-act="finish">Afronden ${ICON.arrow}</button>` : `<button class="btn primary" data-act="section" data-i="${si + 1}">Volgende ${ICON.arrow}</button>`}
        </div>
      </main>
    </div>`;

    if (state.focus) {
      const el = document.getElementById('blk_' + state.focus) || document.getElementById('grow_' + state.focus);
      const inp = document.getElementById('f_' + state.focus);
      state.focus = null;
      if (el) {
        el.scrollIntoView({ block: 'center' });
        el.classList.add('flash');
        if (inp) inp.focus({ preventScroll: true });
      }
    }
  }

  // ---------- waarde wijzigen ----------
  let saveTimer = null;
  function setValue(id, value) {
    const d = state.draft;
    d.values[id] = value;
    d.updated = Date.now();
    clearTimeout(saveTimer);
    saveTimer = setTimeout(saveDrafts, 400);
    updateFieldState(id);
  }

  function updateFieldState(id) {
    const d = state.draft;
    const form = formById(d.formId);
    const f = prepared[form.id].fields.find((x) => x.id === id);
    const blk = document.getElementById('blk_' + id);
    if (blk && f) {
      const filled = isFilled(f, d.values[id]);
      blk.classList.toggle('filled', !!filled);
      const lab = blk.querySelector('.label');
      const tick = lab.querySelector('.tick');
      if (filled && !tick) lab.insertAdjacentHTML('beforeend', `<span class="tick">${ICON.check}</span>`);
      if (!filled && tick) tick.remove();
    }
    const st = formStats(form, d.values);
    const bar = $('.formbar .bar span');
    if (bar) bar.style.width = st.pct + '%';
    const pct = $('.formbar .pct');
    if (pct) pct.textContent = `${st.done}/${st.total}`;
    const fsite = $('.formbar .fsite');
    if (fsite) fsite.textContent = draftTitle(d);
    form.sections.forEach((s, i) => {
      const ss = sectionStats(s, d.values);
      const btn = document.querySelector(`.seclist button[data-i="${i}"]`);
      if (btn) {
        btn.querySelector('.sc').textContent = `${ss.done}/${ss.total}`;
        if (!btn.classList.contains('cur')) {
          btn.classList.toggle('done', !!ss.total && ss.done === ss.total);
          btn.classList.toggle('part', ss.done > 0 && ss.done < ss.total);
        }
      }
      const opt = document.querySelector(`.secselect option[value="${i}"]`);
      if (opt) opt.textContent = `${i + 1}/${form.sections.length} · ${s.title} (${ss.done}/${ss.total})`;
    });
  }

  app.addEventListener('input', (e) => {
    const el = e.target;
    if (!el.dataset.f || el.type === 'checkbox') return;
    setValue(el.dataset.f, el.dataset.date ? fromIsoDate(el.value) : el.value);
  });
  app.addEventListener('change', (e) => {
    const el = e.target;
    if (el.dataset.f) {
      if (el.type === 'checkbox') setValue(el.dataset.f, el.checked);
      else setValue(el.dataset.f, el.dataset.date ? fromIsoDate(el.value) : el.value);
    } else if (el.dataset.act === 'jump') {
      go('form', { section: parseInt(el.value, 10) });
    } else if (el.id === 'mailTo' || el.id === 'mailMsg') {
      const m = lsGet(LS_MAIL, {});
      m[el.id] = el.value;
      lsSet(LS_MAIL, m);
    }
  });

  app.addEventListener('click', (e) => {
    const chip = e.target.closest('[data-choice],[data-single]');
    if (chip) {
      const d = state.draft;
      const val = chip.dataset.val;
      if (chip.dataset.single) {
        const id = chip.dataset.single;
        const next = d.values[id] === val ? '' : val;
        setValue(id, next);
        chip.parentNode.querySelectorAll('.chip').forEach((c) => {
          const on = c.dataset.val === next;
          c.classList.toggle('on', on);
          c.setAttribute('aria-pressed', on);
        });
      } else {
        const id = chip.dataset.choice;
        const form = formById(d.formId);
        const f = prepared[form.id].fields.find((x) => x.id === id);
        let sel = Array.isArray(d.values[id]) ? d.values[id].slice() : [];
        if (sel.includes(val)) sel = sel.filter((x) => x !== val);
        else sel = f.multi ? [...sel, val] : [val];
        setValue(id, sel);
        chip.parentNode.querySelectorAll('.chip').forEach((c) => {
          const on = sel.includes(c.dataset.val);
          c.classList.toggle('on', on);
          c.setAttribute('aria-pressed', on);
        });
      }
      return;
    }
    const btn = e.target.closest('[data-act]');
    if (!btn || btn.tagName === 'SELECT') return;
    const act = btn.dataset.act;
    if (act === 'home') go('home');
    else if (act === 'client') {
      const c = CLIENTS.find((x) => x.id === btn.dataset.id);
      if (c.forms.length === 1) startNew(c.forms[0]);
      else go('kies', { client: c.id });
    } else if (act === 'new') startNew(btn.dataset.id);
    else if (act === 'open-draft') openDraft(btn.dataset.id);
    else if (act === 'del-draft') deleteDraft(btn.dataset.id);
    else if (act === 'section') go('form', { section: parseInt(btn.dataset.i, 10) });
    else if (act === 'finish') go('finish');
    else if (act === 'back-form') go('form');
    else if (act === 'goto-field') go('form', { section: parseInt(btn.dataset.s, 10), focus: btn.dataset.f });
    else if (act === 'download') downloadCurrent();
    else if (act === 'share') shareCurrent();
    else if (act === 'mailto') mailtoCurrent();
  });

  async function openDraft(id) {
    const d = drafts.find((x) => x.id === id);
    if (!d) return;
    busy(true, 'Concept openen…');
    try {
      await prepare(formById(d.formId));
      go('form', { draft: d, section: 0 });
    } catch (e) {
      toast(e.message, 'err');
    } finally {
      busy(false);
    }
  }

  function deleteDraft(id) {
    const d = drafts.find((x) => x.id === id);
    if (!d) return;
    if (!confirm(`Concept "${draftTitle(d)}" verwijderen? Dit kan niet ongedaan worden gemaakt.`)) return;
    drafts = drafts.filter((x) => x.id !== id);
    saveDrafts();
    idb.del('base:' + id);
    renderHome();
  }

  // ---------- afronden ----------
  function openItems(form, values) {
    const out = [];
    form.sections.forEach((sec, si) => {
      for (const it of sec.items) {
        const fs = it.grid ? it.rows.flatMap((r) => r.fields.filter(Boolean)) : [it];
        for (const f of fs) if (isFilled(f, values[f.id]) === false) out.push({ si, f });
      }
    });
    return out;
  }

  function renderFinish() {
    const d = state.draft;
    const form = formById(d.formId);
    const st = formStats(form, d.values);
    const open = openItems(form, d.values);
    const mail = lsGet(LS_MAIL, {});
    const name = sanitizeName(form.fileName(d.values));
    const bySection = {};
    open.forEach((o) => (bySection[o.si] = bySection[o.si] || []).push(o.f));
    const canShareFiles = !!(navigator.canShare && window.File && navigator.canShare({ files: [new File([''], 'x.txt', { type: 'text/plain' })] }));

    app.innerHTML = `${header()}
    <main class="wrap">
      <button class="link-back" data-act="back-form">${ICON.back} Terug naar formulier</button>
      <section class="hero">
        <p class="eyebrow">Afronden</p>
        <h1>${st.done === st.total ? 'Alles is ingevuld' : `${st.done} van ${st.total} vragen ingevuld`}</h1>
        <div class="bar big"><span style="width:${st.pct}%"></span></div>
      </section>

      ${open.length ? `<details class="block openlist"${open.length <= 12 ? ' open' : ''}>
        <summary><strong>${open.length} ${open.length === 1 ? 'vraag' : 'vragen'} nog leeg</strong> <span class="muted">— niet alles hoeft ingevuld te zijn; je kunt ook zo versturen.</span></summary>
        ${Object.keys(bySection).map((si) => `<div class="open-sec"><div class="open-sec-title">${esc(form.sections[si].title)}</div>
          <div class="open-items">${bySection[si].map((f) => `<button class="open-item" data-act="goto-field" data-s="${si}" data-f="${f.id}">${esc(f.label)}</button>`).join('')}</div></div>`).join('')}
      </details>` : ''}

      <section class="block sendbox">
        <h2>Versturen</h2>
        <div class="filecard">${form.kind === 'xlsx' ? ICON.xls : ICON.doc}<div><div class="fname">${esc(name)}</div><div class="muted" id="fileState">Bestand wordt gemaakt…</div></div></div>
        <div class="field"><label class="label" for="mailTo">E-mailadres ontvanger</label>
          <input type="email" id="mailTo" inputmode="email" autocomplete="email" placeholder="naam@bedrijf.nl" value="${esc(mail.mailTo || '')}"></div>
        <div class="field"><label class="label" for="mailMsg">Bericht</label>
          <textarea id="mailMsg" rows="3">${esc(mail.mailMsg || 'Hierbij het ingevulde rapport.\n\nMet vriendelijke groet,')}</textarea></div>
        <div class="send-actions">
          ${canShareFiles ? `<button class="btn primary" data-act="share" disabled>${ICON.mail} Mail met bijlage</button>` : ''}
          <button class="btn ${canShareFiles ? 'secondary' : 'primary'}" data-act="mailto" disabled>${ICON.mail} ${canShareFiles ? 'Open e-mail' : 'Mail versturen'}</button>
          <button class="btn secondary" data-act="download" disabled>${ICON.download} Download</button>
        </div>
        <p class="hint">${canShareFiles
          ? '“Mail met bijlage” opent het deelmenu: kies Mail of Outlook, het bestand zit er al aan. Het e-mailadres van de ontvanger staat op je klembord om te plakken.'
          : '“Mail versturen” downloadt het bestand en opent je e-mailprogramma met ontvanger en onderwerp ingevuld. Voeg het gedownloade bestand daar als bijlage toe.'}</p>
      </section>
    </main>`;

    state.finishFile = null;
    buildFile(d).then((file) => {
      state.finishFile = file;
      const fs = $('#fileState');
      if (fs) fs.textContent = `${form.kind === 'xlsx' ? 'Excel' : 'Word'} · ${(file.blob.size / 1024).toFixed(0)} KB · klaar`;
      app.querySelectorAll('.send-actions .btn').forEach((b) => (b.disabled = false));
    }).catch((e) => {
      console.error(e);
      const fs = $('#fileState');
      if (fs) fs.textContent = 'Fout bij maken van bestand: ' + e.message;
    });
  }

  function triggerDownload(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }

  async function downloadCurrent() {
    const d = state.draft;
    if (!d) return;
    try {
      const file = state.view === 'finish' && state.finishFile ? state.finishFile : await withBusy('Bestand maken…', () => buildFile(d));
      triggerDownload(file.blob, file.name);
      toast('Bestand gedownload.');
    } catch (e) {
      console.error(e);
      toast('Bestand maken mislukt: ' + e.message, 'err');
    }
  }

  function mailParts() {
    const d = state.draft;
    const form = formById(d.formId);
    const to = ($('#mailTo') && $('#mailTo').value.trim()) || '';
    const msg = ($('#mailMsg') && $('#mailMsg').value) || '';
    const subject = `${form.klant === 'kpn' ? 'KPN RCT rapport' : 'Vodafone TSS report'} – ${draftTitle(d)}`;
    return { to, msg, subject };
  }

  async function shareCurrent() {
    const file = state.finishFile;
    if (!file) return;
    const { to, msg, subject } = mailParts();
    const f = new File([file.blob], file.name, { type: file.blob.type });
    if (to && navigator.clipboard) navigator.clipboard.writeText(to).catch(() => {});
    try {
      await navigator.share({ files: [f], title: subject, text: msg });
    } catch (e) {
      if (e && e.name === 'AbortError') return;
      console.error(e);
      toast('Delen lukte niet; gebruik “Open e-mail” of “Download”.', 'warn');
    }
  }

  function mailtoCurrent() {
    const file = state.finishFile;
    if (!file) return;
    const { to, msg, subject } = mailParts();
    triggerDownload(file.blob, file.name);
    const body = `${msg}\n\n(Bijlage: ${file.name})`;
    setTimeout(() => {
      window.location.href = `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
    }, 400);
    toast('Bestand gedownload – voeg het toe als bijlage in je e-mail.');
  }

  // ---------- bezig-indicator ----------
  function busy(on, msg) {
    let el = $('#busy');
    if (!el) {
      el = document.createElement('div');
      el.id = 'busy';
      el.innerHTML = '<div class="spinner"></div><div class="busy-msg"></div>';
      document.body.appendChild(el);
    }
    el.querySelector('.busy-msg').textContent = msg || '';
    el.classList.toggle('show', on);
  }
  async function withBusy(msg, fn) {
    busy(true, msg);
    try {
      return await fn();
    } finally {
      busy(false);
    }
  }

  window.addEventListener('beforeunload', () => {
    clearTimeout(saveTimer);
    saveDrafts();
  });
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'hidden') saveDrafts();
  });

  // Testhaak voor automatische controles
  window.DAEL = { prepare, readValues, writeValues, buildFile, openDoc, loadTemplate, allFields };

  if ('serviceWorker' in navigator && location.protocol === 'https:') {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }

  render();
})();
