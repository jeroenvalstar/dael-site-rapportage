/*
 * Lezen en schrijven van .xlsx en .docx direct op XML-niveau.
 * Alles buiten de ingevulde velden (opmaak, formules, logo's, beveiliging) blijft ongewijzigd.
 */
(function () {
  'use strict';

  const NS = {
    s: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    pr: 'http://schemas.openxmlformats.org/package/2006/relationships',
    w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
    xml: 'http://www.w3.org/XML/1998/namespace',
  };

  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

  function parseXml(str) {
    const doc = new DOMParser().parseFromString(str, 'application/xml');
    const err = doc.getElementsByTagName('parsererror')[0];
    if (err) throw new Error('XML kon niet gelezen worden: ' + err.textContent.slice(0, 200));
    return doc;
  }

  function serializeXml(doc) {
    let s = new XMLSerializer().serializeToString(doc);
    s = s.replace(/^<\?xml[^>]*\?>\s*/, '');
    return XML_DECL + s;
  }

  function kids(el, ns, name) {
    const out = [];
    for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
      if (c.namespaceURI === ns && c.localName === name) out.push(c);
    }
    return out;
  }

  function kid(el, ns, name) {
    for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
      if (c.namespaceURI === ns && c.localName === name) return c;
    }
    return null;
  }

  function resolvePath(base, target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = base.split('/');
    parts.pop();
    for (const seg of target.split('/')) {
      if (seg === '..') parts.pop();
      else if (seg !== '.') parts.push(seg);
    }
    return parts.join('/');
  }

  // ---------- Excel ----------

  function colToNum(col) {
    let n = 0;
    for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }

  function splitRef(ref) {
    const m = /^([A-Z]+)(\d+)$/.exec(ref);
    if (!m) throw new Error('Ongeldige celverwijzing: ' + ref);
    return { col: m[1], row: parseInt(m[2], 10), colNum: colToNum(m[1]) };
  }

  // Excel-datum (serieel getal) naar dd-mm-jjjj
  function serialToDate(n) {
    const ms = Math.round((n - 25569) * 86400 * 1000);
    const d = new Date(ms);
    if (isNaN(d)) return String(n);
    const p = (x) => String(x).padStart(2, '0');
    return `${p(d.getUTCDate())}-${p(d.getUTCMonth() + 1)}-${d.getUTCFullYear()}`;
  }

  class XlsxFile {
    static async load(data) {
      const zip = await JSZip.loadAsync(data);
      const f = new XlsxFile(zip);
      await f._init();
      return f;
    }

    constructor(zip) {
      this.zip = zip;
      this.sheetPaths = {};
      this.sheetDocs = {};
      this.cellMaps = {};
      this.dirty = new Set();
      this.sst = [];
    }

    async _init() {
      const wbPath = 'xl/workbook.xml';
      const wbFile = this.zip.file(wbPath);
      if (!wbFile) throw new Error('Dit is geen geldig Excel-bestand (.xlsx).');
      this.wbDoc = parseXml(await wbFile.async('string'));
      const rels = parseXml(await this.zip.file('xl/_rels/workbook.xml.rels').async('string'));
      const relMap = {};
      for (const r of rels.getElementsByTagNameNS(NS.pr, 'Relationship')) {
        relMap[r.getAttribute('Id')] = resolvePath(wbPath, r.getAttribute('Target'));
      }
      for (const s of this.wbDoc.getElementsByTagNameNS(NS.s, 'sheet')) {
        this.sheetPaths[s.getAttribute('name')] = relMap[s.getAttributeNS(NS.r, 'id')];
      }
      const sstFile = this.zip.file('xl/sharedStrings.xml');
      if (sstFile) {
        const sstDoc = parseXml(await sstFile.async('string'));
        for (const si of sstDoc.getElementsByTagNameNS(NS.s, 'si')) {
          let txt = '';
          for (const t of si.getElementsByTagNameNS(NS.s, 't')) {
            if (t.parentNode.localName === 'rPh') continue; // fonetische hulptekst overslaan
            txt += t.textContent;
          }
          this.sst.push(txt);
        }
      }
    }

    hasSheet(name) {
      return !!this.sheetPaths[name];
    }

    async _sheet(name) {
      if (this.sheetDocs[name]) return this.sheetDocs[name];
      const path = this.sheetPaths[name];
      if (!path) throw new Error('Tabblad niet gevonden: ' + name);
      const doc = parseXml(await this.zip.file(path).async('string'));
      const map = {};
      for (const c of doc.getElementsByTagNameNS(NS.s, 'c')) map[c.getAttribute('r')] = c;
      this.sheetDocs[name] = doc;
      this.cellMaps[name] = map;
      return doc;
    }

    async read(sheet, ref) {
      await this._sheet(sheet);
      const c = this.cellMaps[sheet][ref];
      if (!c) return { value: '', formula: false, numeric: false };
      const formula = !!kid(c, NS.s, 'f');
      const t = c.getAttribute('t');
      const v = kid(c, NS.s, 'v');
      let value = '';
      if (t === 's') value = v ? this.sst[parseInt(v.textContent, 10)] || '' : '';
      else if (t === 'inlineStr') {
        const is = kid(c, NS.s, 'is');
        if (is) for (const tt of is.getElementsByTagNameNS(NS.s, 't')) value += tt.textContent;
      } else if (v) value = v.textContent;
      const numeric = !t || t === 'n';
      if (numeric && value !== '' && /^-?\d+\.\d{6,}$/.test(value)) {
        value = String(parseFloat(parseFloat(value).toPrecision(12)));
      }
      return { value, formula, numeric };
    }

    async write(sheet, ref, value) {
      const doc = await this._sheet(sheet);
      const map = this.cellMaps[sheet];
      let c = map[ref];
      value = value == null ? '' : String(value);
      if (!c) {
        if (value === '') return;
        c = this._createCell(doc, ref);
        map[ref] = c;
      }
      for (const name of ['f', 'v', 'is']) {
        const e = kid(c, NS.s, name);
        if (e) c.removeChild(e);
      }
      c.removeAttribute('t');
      if (value !== '') {
        if (/^-?(0|[1-9]\d{0,14})([.,]\d+)?$/.test(value)) {
          const v = doc.createElementNS(NS.s, 'v');
          v.textContent = value.replace(',', '.');
          c.appendChild(v);
        } else {
          c.setAttribute('t', 'inlineStr');
          const is = doc.createElementNS(NS.s, 'is');
          const t = doc.createElementNS(NS.s, 't');
          t.setAttributeNS(NS.xml, 'xml:space', 'preserve');
          t.textContent = value;
          is.appendChild(t);
          c.appendChild(is);
        }
      }
      this.dirty.add(sheet);
    }

    _createCell(doc, ref) {
      const { row, colNum } = splitRef(ref);
      const sheetData = doc.getElementsByTagNameNS(NS.s, 'sheetData')[0];
      let rowEl = null;
      let before = null;
      for (const r of kids(sheetData, NS.s, 'row')) {
        const rn = parseInt(r.getAttribute('r'), 10);
        if (rn === row) { rowEl = r; break; }
        if (rn > row) { before = r; break; }
      }
      if (!rowEl) {
        rowEl = doc.createElementNS(NS.s, 'row');
        rowEl.setAttribute('r', String(row));
        sheetData.insertBefore(rowEl, before);
      }
      let cellBefore = null;
      for (const c of kids(rowEl, NS.s, 'c')) {
        if (splitRef(c.getAttribute('r')).colNum > colNum) { cellBefore = c; break; }
      }
      const c = doc.createElementNS(NS.s, 'c');
      c.setAttribute('r', ref);
      const s = rowEl.getAttribute('s');
      if (s && rowEl.getAttribute('customFormat') === '1') c.setAttribute('s', s);
      rowEl.insertBefore(c, cellBefore);
      return c;
    }

    async toBlob() {
      for (const name of this.dirty) {
        this.zip.file(this.sheetPaths[name], serializeXml(this.sheetDocs[name]));
      }
      if (this.dirty.size) {
        // Formules (zoals functie/bedrijf bij een naam) opnieuw laten berekenen bij openen
        let calc = this.wbDoc.getElementsByTagNameNS(NS.s, 'calcPr')[0];
        if (calc) {
          calc.setAttribute('fullCalcOnLoad', '1');
          this.zip.file('xl/workbook.xml', serializeXml(this.wbDoc));
        }
      }
      return this.zip.generateAsync({
        type: 'blob',
        compression: 'DEFLATE',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
    }
  }

  // ---------- Word ----------

  // Zoekt alle velden (fldChar begin/separate/end) in documentvolgorde
  function scanFields(doc) {
    const ff = [];
    const refs = [];
    const stack = [];
    const runs = Array.from(doc.getElementsByTagNameNS(NS.w, 'r'));
    for (const r of runs) {
      let isFieldRun = false;
      for (let c = r.firstElementChild; c; c = c.nextElementSibling) {
        if (c.namespaceURI !== NS.w) continue;
        if (c.localName === 'fldChar') {
          isFieldRun = true;
          const type = c.getAttributeNS(NS.w, 'fldCharType');
          if (type === 'begin') {
            const f = { ffData: kid(c, NS.w, 'ffData'), instr: '', phase: 'instr', beginRun: r, sepRun: null, result: [], nested: false };
            if (stack.length) stack[stack.length - 1].nested = true;
            stack.push(f);
            if (f.ffData) ff.push(f);
          } else if (type === 'separate' && stack.length) {
            const f = stack[stack.length - 1];
            f.phase = 'result';
            f.sepRun = r;
          } else if (type === 'end' && stack.length) {
            const f = stack.pop();
            if (!f.ffData) {
              const m = /^\s*REF\s+(\S+)/.exec(f.instr);
              if (m) { f.name = m[1]; refs.push(f); }
            }
          }
        } else if (c.localName === 'instrText') {
          isFieldRun = true;
          if (stack.length) stack[stack.length - 1].instr += c.textContent;
        }
      }
      if (!isFieldRun && stack.length) {
        const f = stack[stack.length - 1];
        if (f.phase === 'result') f.result.push(r);
      }
    }
    return { ff, refs };
  }

  function closestPara(el) {
    while (el && !(el.namespaceURI === NS.w && el.localName === 'p')) el = el.parentNode;
    return el;
  }

  function runText(r) {
    let out = '';
    for (let c = r.firstElementChild; c; c = c.nextElementSibling) {
      if (c.namespaceURI !== NS.w) continue;
      if (c.localName === 't') out += c.textContent;
      else if (c.localName === 'br' || c.localName === 'cr') out += '\n';
      else if (c.localName === 'tab') out += '\t';
      else if (c.localName === 'noBreakHyphen') out += '-';
    }
    return out;
  }

  function resultGroups(f) {
    const groups = [];
    let last = null;
    for (const r of f.result) {
      if (!r.parentNode) continue;
      const p = closestPara(r);
      if (p !== last) { groups.push([]); last = p; }
      groups[groups.length - 1].push(r);
    }
    return groups;
  }

  function fieldResultText(f) {
    return resultGroups(f).map((g) => g.map(runText).join('')).join('\n');
  }

  function setRunText(doc, r, text) {
    for (let c = r.firstElementChild; c; ) {
      const next = c.nextElementSibling;
      if (!(c.namespaceURI === NS.w && c.localName === 'rPr')) r.removeChild(c);
      c = next;
    }
    String(text).split('\n').forEach((line, i) => {
      if (i > 0) r.appendChild(doc.createElementNS(NS.w, 'w:br'));
      const t = doc.createElementNS(NS.w, 'w:t');
      t.setAttributeNS(NS.xml, 'xml:space', 'preserve');
      t.textContent = line;
      r.appendChild(t);
    });
  }

  // Zet de zichtbare tekst van een veld; regels worden over de bestaande alinea's verdeeld
  function setFieldResult(doc, f, text) {
    let groups = resultGroups(f);
    if (!groups.length) {
      if (!f.sepRun) return;
      const r = doc.createElementNS(NS.w, 'w:r');
      const rPr = kid(f.sepRun, NS.w, 'rPr') || kid(f.beginRun, NS.w, 'rPr');
      if (rPr) r.appendChild(rPr.cloneNode(true));
      f.sepRun.parentNode.insertBefore(r, f.sepRun.nextSibling);
      groups = [[r]];
    }
    const lines = String(text).split('\n');
    groups.forEach((g, k) => {
      const t = k < groups.length - 1 ? lines[k] || '' : lines.slice(k).join('\n');
      setRunText(doc, g[0], t);
      for (const r of g.slice(1)) r.parentNode.removeChild(r);
    });
    f.result = groups.map((g) => g[0]);
  }

  class DocxFile {
    static async load(data) {
      const zip = await JSZip.loadAsync(data);
      const f = new DocxFile(zip);
      await f._init();
      return f;
    }

    constructor(zip) {
      this.zip = zip;
    }

    async _init() {
      const file = this.zip.file('word/document.xml');
      if (!file) throw new Error('Dit is geen geldig Word-bestand (.docx).');
      this.doc = parseXml(await file.async('string'));
      this.body = this.doc.getElementsByTagNameNS(NS.w, 'body')[0];
      this.tables = kids(this.body, NS.w, 'tbl');
      this.checkboxes = Array.from(this.doc.getElementsByTagNameNS(NS.w14, 'checkbox'));
      const scan = scanFields(this.doc);
      this.legacy = scan.ff;
      this.refs = scan.refs;
    }

    // ---------- klassieke formuliervelden (FORMTEXT / FORMCHECKBOX / FORMDROPDOWN) ----------

    legacyName(i) {
      const f = this.legacy[i];
      const n = f && kid(f.ffData, NS.w, 'name');
      return n ? n.getAttributeNS(NS.w, 'val') || '' : '';
    }

    legacyType(i) {
      const f = this.legacy[i];
      if (!f) return null;
      if (kid(f.ffData, NS.w, 'checkBox')) return 'check';
      if (kid(f.ffData, NS.w, 'ddList')) return 'dropdown';
      return 'text';
    }

    legacyText(i) {
      const f = this.legacy[i];
      return f ? fieldResultText(f) : '';
    }

    setLegacyText(i, value) {
      const f = this.legacy[i];
      if (!f || f.nested) return;
      setFieldResult(this.doc, f, value || '\u2002\u2002\u2002\u2002\u2002');
    }

    legacyCheck(i) {
      const f = this.legacy[i];
      const cb = f && kid(f.ffData, NS.w, 'checkBox');
      if (!cb) return false;
      const el = kid(cb, NS.w, 'checked') || kid(cb, NS.w, 'default');
      if (!el) return false;
      const v = el.getAttributeNS(NS.w, 'val');
      return !v || v === '1' || v === 'true' || v === 'on';
    }

    setLegacyCheck(i, on) {
      const f = this.legacy[i];
      const cb = f && kid(f.ffData, NS.w, 'checkBox');
      if (!cb) return;
      let ch = kid(cb, NS.w, 'checked');
      if (!ch) {
        ch = this.doc.createElementNS(NS.w, 'w:checked');
        cb.appendChild(ch);
      }
      ch.setAttributeNS(NS.w, 'w:val', on ? '1' : '0');
    }

    legacyOptions(i) {
      const f = this.legacy[i];
      const dd = f && kid(f.ffData, NS.w, 'ddList');
      return dd ? kids(dd, NS.w, 'listEntry').map((e) => e.getAttributeNS(NS.w, 'val')) : [];
    }

    legacyDropdown(i) {
      const f = this.legacy[i];
      const dd = f && kid(f.ffData, NS.w, 'ddList');
      if (!dd) return '';
      const el = kid(dd, NS.w, 'result') || kid(dd, NS.w, 'default');
      const idx = el ? parseInt(el.getAttributeNS(NS.w, 'val'), 10) || 0 : 0;
      return this.legacyOptions(i)[idx] || '';
    }

    setLegacyDropdown(i, value) {
      const f = this.legacy[i];
      const dd = f && kid(f.ffData, NS.w, 'ddList');
      if (!dd) return;
      const idx = Math.max(0, this.legacyOptions(i).indexOf(value));
      let res = kid(dd, NS.w, 'result');
      if (!res) {
        res = this.doc.createElementNS(NS.w, 'w:result');
        dd.insertBefore(res, dd.firstChild);
      }
      res.setAttributeNS(NS.w, 'w:val', String(idx));
    }

    // Verwijzingen (REF-velden, ook in kop- en voetteksten) gelijktrekken met de formuliervelden
    async updateRefs() {
      const values = {};
      this.legacy.forEach((f, i) => {
        const name = this.legacyName(i);
        if (name && this.legacyType(i) === 'text' && !f.nested) values[name] = this.legacyText(i);
      });
      const apply = (doc, refs) => {
        let changed = false;
        for (const r of refs) {
          if (!(r.name in values) || r.nested) continue;
          const want = values[r.name] || '\u2002\u2002\u2002\u2002\u2002';
          if (fieldResultText(r) !== want) {
            setFieldResult(doc, r, want);
            changed = true;
          }
        }
        return changed;
      };
      apply(this.doc, this.refs);
      this.extraParts = this.extraParts || {};
      for (const path of Object.keys(this.zip.files)) {
        if (!/^word\/(header|footer)\d*\.xml$/.test(path)) continue;
        const xml = this.extraParts[path] ? null : await this.zip.file(path).async('string');
        if (xml !== null && !/REF\s/.test(xml)) continue;
        const doc = this.extraParts[path] || parseXml(xml);
        if (apply(doc, scanFields(doc).refs)) this.extraParts[path] = doc;
      }
    }

    _cells(tr) {
      const out = [];
      const walk = (el) => {
        for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
          if (c.namespaceURI !== NS.w) continue;
          if (c.localName === 'tc') out.push(c);
          else if (c.localName === 'sdt' || c.localName === 'sdtContent' || c.localName === 'customXml') walk(c);
        }
      };
      walk(tr);
      return out;
    }

    cellParas(t, r, c) {
      const tbl = this.tables[t];
      if (!tbl) return null;
      const tr = kids(tbl, NS.w, 'tr')[r];
      if (!tr) return null;
      const tc = this._cells(tr)[c];
      if (!tc) return null;
      return { tc, paras: Array.from(tc.getElementsByTagNameNS(NS.w, 'p')) };
    }

    // Tekst van een alinea, zonder de ☐/☒ tekens van selectievakjes
    paraText(p) {
      let out = '';
      const walk = (el) => {
        for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
          if (c.namespaceURI === NS.w) {
            if (c.localName === 'sdt' && c.getElementsByTagNameNS(NS.w14, 'checkbox').length) continue;
            if (c.localName === 't') { out += c.textContent; continue; }
            if (c.localName === 'br' || c.localName === 'cr') { out += '\n'; continue; }
            if (c.localName === 'tab') { out += '\t'; continue; }
            if (c.localName === 'pPr' || c.localName === 'rPr') continue;
          }
          walk(c);
        }
      };
      walk(p);
      return out;
    }

    _runProps(p) {
      for (const r of p.getElementsByTagNameNS(NS.w, 'r')) {
        if (r.getElementsByTagNameNS(NS.w, 't').length) {
          const rPr = kid(r, NS.w, 'rPr');
          return rPr ? rPr.cloneNode(true) : null;
        }
      }
      const pPr = kid(p, NS.w, 'pPr');
      const markRPr = pPr && kid(pPr, NS.w, 'rPr');
      if (markRPr) {
        const rPr = this.doc.createElementNS(NS.w, 'w:rPr');
        for (let c = markRPr.firstElementChild; c; c = c.nextElementSibling) {
          if (!['ins', 'del', 'moveFrom', 'moveTo'].includes(c.localName)) rPr.appendChild(c.cloneNode(true));
        }
        return rPr;
      }
      return null;
    }

    setParaText(p, text) {
      const rPr = this._runProps(p);
      for (let c = p.firstElementChild; c; ) {
        const next = c.nextElementSibling;
        if (!(c.namespaceURI === NS.w && c.localName === 'pPr')) p.removeChild(c);
        c = next;
      }
      if (!text) return;
      const r = this.doc.createElementNS(NS.w, 'w:r');
      if (rPr) r.appendChild(rPr);
      text.split('\n').forEach((line, i) => {
        if (i > 0) r.appendChild(this.doc.createElementNS(NS.w, 'w:br'));
        const t = this.doc.createElementNS(NS.w, 'w:t');
        t.setAttributeNS(NS.xml, 'xml:space', 'preserve');
        t.textContent = line;
        r.appendChild(t);
      });
      p.appendChild(r);
    }

    // Zorgt dat alinea-index idx bestaat in de cel (voegt lege alinea's toe na de laatste)
    ensurePara(tc, paras, idx) {
      while (paras.length <= idx) {
        const last = paras[paras.length - 1];
        const p = this.doc.createElementNS(NS.w, 'w:p');
        const pPr = last && kid(last, NS.w, 'pPr');
        if (pPr) p.appendChild(pPr.cloneNode(true));
        last.parentNode.insertBefore(p, last.nextSibling);
        paras.push(p);
      }
      return paras[idx];
    }

    getCheck(i) {
      const cb = this.checkboxes[i];
      if (!cb) return false;
      const ch = kid(cb, NS.w14, 'checked');
      const v = ch && ch.getAttributeNS(NS.w14, 'val');
      return v === '1' || v === 'true';
    }

    setCheck(i, on) {
      const cb = this.checkboxes[i];
      if (!cb) return;
      let ch = kid(cb, NS.w14, 'checked');
      if (!ch) {
        ch = this.doc.createElementNS(NS.w14, 'w14:checked');
        cb.insertBefore(ch, cb.firstChild);
      }
      ch.setAttributeNS(NS.w14, 'w14:val', on ? '1' : '0');
      const stateEl = kid(cb, NS.w14, on ? 'checkedState' : 'uncheckedState');
      const code = stateEl && stateEl.getAttributeNS(NS.w14, 'val');
      const glyph = code ? String.fromCharCode(parseInt(code, 16)) : on ? '☒' : '☐';
      const sdt = cb.parentNode.parentNode; // w14:checkbox -> w:sdtPr -> w:sdt
      const content = kid(sdt, NS.w, 'sdtContent');
      if (!content) return;
      const ts = content.getElementsByTagNameNS(NS.w, 't');
      if (ts.length) {
        ts[0].textContent = glyph;
        for (let k = 1; k < ts.length; k++) ts[k].textContent = '';
      }
    }

    async toBlob() {
      this.zip.file('word/document.xml', serializeXml(this.doc));
      for (const [path, doc] of Object.entries(this.extraParts || {})) this.zip.file(path, serializeXml(doc));
      return this.zip.generateAsync({
        type: 'blob',
        compression: 'DEFLATE',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
    }
  }

  window.OOXML = { XlsxFile, DocxFile, serialToDate };
})();
