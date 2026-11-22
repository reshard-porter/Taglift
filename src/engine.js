/* Taglift engine — converts Word run formatting into literal HTML tags inside a .docx.
   Runs in the browser (and in Node for testing, given DOMParser/XMLSerializer). */
const Taglift = (() => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const XML_NS = 'http://www.w3.org/XML/1998/namespace';
  const KEYS = ['b', 'i', 'u', 's', 'sub', 'sup'];
  const TAGSETS = {
    classic:  { b: 'b', i: 'i', u: 'u', s: 's', sub: 'sub', sup: 'sup' },
    semantic: { b: 'strong', i: 'em', u: 'u', s: 'del', sub: 'sub', sup: 'sup' },
    legacy:   { b: 'b', i: 'i', u: 'u', s: 'strike', sub: 'sub', sup: 'sup' },
  };
  // Canonical child order of w:rPr (CT_RPr); Word is strict about it.
  const RPR_ORDER = ['rStyle','rFonts','b','bCs','i','iCs','caps','smallCaps','strike','dstrike','outline','shadow','emboss','imprint','noProof','snapToGrid','vanish','webHidden','color','spacing','w','kern','position','sz','szCs','highlight','u','effect','bdr','shd','fitText','vertAlign','rtl','cs','em','lang','eastAsianLayout','specVanish','oMath','rPrChange'];

  /* ---------------- ZIP ---------------- */
  const CRC_TABLE = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  function crc32(u8) { let c = 0xFFFFFFFF; for (let i = 0; i < u8.length; i++) c = CRC_TABLE[(c ^ u8[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  async function streamThrough(u8, stream) {
    const blob = new Blob([u8]);
    const out = await new Response(blob.stream().pipeThrough(stream)).arrayBuffer();
    return new Uint8Array(out);
  }
  const inflate = (u8) => streamThrough(u8, new DecompressionStream('deflate-raw'));
  const deflate = (u8) => streamThrough(u8, new CompressionStream('deflate-raw'));

  function readZip(buf) {
    const u8 = new Uint8Array(buf), dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    if (u8.length >= 4 && u8[0] === 0xD0 && u8[1] === 0xCF && u8[2] === 0x11 && u8[3] === 0xE0)
      throw new UserError('This file is password-protected or is an old-style .doc file. Remove the password (or re-save it as .docx in Word) and try again.');
    let e = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { e = i; break; }
    if (e < 0) throw new UserError('This doesn’t look like a valid .docx file.');
    const n = dv.getUint16(e + 10, true), cdOff = dv.getUint32(e + 16, true);
    const td = new TextDecoder();
    const entries = [];
    let p = cdOff;
    for (let k = 0; k < n; k++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new UserError('The .docx file appears to be damaged.');
      const flags = dv.getUint16(p + 8, true), method = dv.getUint16(p + 10, true);
      const time = dv.getUint16(p + 12, true), date = dv.getUint16(p + 14, true);
      const crc = dv.getUint32(p + 16, true), csize = dv.getUint32(p + 20, true), usize = dv.getUint32(p + 24, true);
      const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
      const lho = dv.getUint32(p + 42, true);
      const nameBytes = u8.slice(p + 46, p + 46 + nlen);
      const name = td.decode(nameBytes);
      const ds = lho + 30 + dv.getUint16(lho + 26, true) + dv.getUint16(lho + 28, true);
      entries.push({ name, nameBytes, flags, method, time, date, crc, csize, usize, raw: u8.subarray(ds, ds + csize) });
      p += 46 + nlen + xlen + clen;
    }
    return entries;
  }
  async function entryData(en) {
    if (en.method === 0) return en.raw;
    if (en.method === 8) return inflate(en.raw);
    throw new UserError('Unsupported compression in .docx.');
  }
  async function writeZip(entries) {
    const parts = [], central = [];
    let off = 0;
    for (const en of entries) {
      let { method, crc, csize, usize, raw } = en;
      if (en.newData) {
        usize = en.newData.length; crc = crc32(en.newData);
        raw = await deflate(en.newData); method = 8; csize = raw.length;
      }
      const flags = en.flags & 0x0800; // keep UTF-8 flag only; sizes are written up front
      const nb = en.nameBytes;
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, flags, true); lh.setUint16(8, method, true);
      lh.setUint16(10, en.time, true); lh.setUint16(12, en.date, true); lh.setUint32(14, crc, true);
      lh.setUint32(18, csize, true); lh.setUint32(22, usize, true); lh.setUint16(26, nb.length, true); lh.setUint16(28, 0, true);
      parts.push(new Uint8Array(lh.buffer), nb, raw);
      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, flags, true);
      ch.setUint16(10, method, true); ch.setUint16(12, en.time, true); ch.setUint16(14, en.date, true); ch.setUint32(16, crc, true);
      ch.setUint32(20, csize, true); ch.setUint32(24, usize, true); ch.setUint16(28, nb.length, true);
      ch.setUint32(42, off, true);
      central.push(new Uint8Array(ch.buffer), nb);
      off += 30 + nb.length + raw.length;
    }
    let cdSize = 0; for (const c of central) cdSize += c.length;
    const eo = new DataView(new ArrayBuffer(22));
    eo.setUint32(0, 0x06054b50, true); eo.setUint16(8, entries.length, true); eo.setUint16(10, entries.length, true);
    eo.setUint32(12, cdSize, true); eo.setUint32(16, off, true);
    return new Blob([...parts, ...central, new Uint8Array(eo.buffer)], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  }

  class UserError extends Error {}

  /* ---------------- XML helpers ---------------- */
  const isW = (n, name) => n && n.nodeType === 1 && n.namespaceURI === W && n.localName === name;
  const kids = (n) => { const a = []; for (let c = n.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) a.push(c); return a; };
  const child = (n, name) => { for (let c = n && n.firstChild; c; c = c.nextSibling) if (isW(c, name)) return c; return null; };
  const wattr = (n, name) => n.getAttributeNS(W, name) || n.getAttribute('w:' + name) || '';

  function onOff(el) {
    if (!el) return undefined;
    const v = wattr(el, 'val');
    return !(v === '0' || v === 'false' || v === 'off');
  }
  function parseRPr(rPr) {
    const f = {};
    if (!rPr) return f;
    const b = onOff(child(rPr, 'b')); if (b !== undefined) f.b = b;
    const i = onOff(child(rPr, 'i')); if (i !== undefined) f.i = i;
    const u = child(rPr, 'u'); if (u) f.u = wattr(u, 'val') !== 'none';
    const st = onOff(child(rPr, 'strike')), ds = onOff(child(rPr, 'dstrike'));
    if (st !== undefined || ds !== undefined) f.s = !!(st || ds);
    const va = child(rPr, 'vertAlign');
    if (va) { const v = wattr(va, 'val'); f.sub = v === 'subscript'; f.sup = v === 'superscript'; }
    return f;
  }
  const merge = (...fs) => Object.assign({}, ...fs);

  function loadStyles(stylesDoc) {
    const S = { defaults: {}, byId: {}, defaultPara: null, cache: {} };
    if (!stylesDoc) return S;
    const root = stylesDoc.documentElement;
    const dd = child(root, 'docDefaults');
    const rd = dd && child(dd, 'rPrDefault');
    S.defaults = parseRPr(rd && child(rd, 'rPr'));
    for (const st of kids(root)) {
      if (!isW(st, 'style')) continue;
      const id = wattr(st, 'styleId'), type = wattr(st, 'type');
      const based = child(st, 'basedOn');
      S.byId[id] = { type, basedOn: based ? wattr(based, 'val') : null, f: parseRPr(child(st, 'rPr')) };
      if (type === 'paragraph' && /^(1|true|on)$/.test(wattr(st, 'default'))) S.defaultPara = id;
    }
    return S;
  }
  function styleFmt(S, id, depth = 0) {
    if (!id || !S.byId[id] || depth > 20) return {};
    if (S.cache[id]) return S.cache[id];
    const s = S.byId[id];
    return (S.cache[id] = merge(styleFmt(S, s.basedOn, depth + 1), s.f));
  }

  function setRPrChild(rPr, el) {
    const idx = RPR_ORDER.indexOf(el.localName);
    for (const c of kids(rPr)) {
      const ci = RPR_ORDER.indexOf(c.localName);
      if (c.namespaceURI === W && ci > idx) { rPr.insertBefore(el, c); return; }
    }
    rPr.appendChild(el);
  }

  /* ---------------- core ---------------- */
  const CONTENT_SKIP = new Set(['rPr', 'lastRenderedPageBreak', 'commentReference', 'annotationRef', 'separator', 'continuationSeparator', 'delText', 'delInstrText', 'instrText', 'dayShort','monthShort','yearShort','dayLong','monthLong','yearLong']);
  const OPAQUE_WS = { tab: '\t', ptab: '\t', br: '\n', cr: '\n' };
  const OPAQUE_TEXT = new Set(['noBreakHyphen', 'softHyphen', 'sym', 'footnoteReference', 'endnoteReference', 'footnoteRef', 'endnoteRef', 'pgNum']);
  const NEUTRAL = new Set(['drawing', 'pict', 'object']);
  const DELETED = new Set(['del', 'moveFrom']);

  function processPart(doc, S, opts, stats, previews) {
    const tags = TAGSETS[opts.tagStyle] || TAGSETS.classic;
    const enabled = KEYS.filter((k) => opts.formats[k]);
    if (!enabled.length) return false;
    let changed = false;

    // Map each run to its nearest paragraph, flag deleted runs.
    const byPara = new Map();
    const runs = Array.from(doc.getElementsByTagNameNS(W, 'r'));
    for (const r of runs) {
      let p = r.parentNode, del = false;
      while (p && !isW(p, 'p')) { if (p.namespaceURI === W && DELETED.has(p.localName)) del = true; if (isW(p, 'rPrChange')) { p = null; break; } p = p.parentNode; }
      if (!p) continue;
      if (!byPara.has(p)) byPara.set(p, []);
      byPara.get(p).push({ r, del });
    }
    const paras = Array.from(doc.getElementsByTagNameNS(W, 'p')).filter((p) => byPara.has(p));
    const fieldStack = [];

    for (const p of paras) {
      stats.paragraphs++;
      // Paragraph style formatting (applies to all runs unless overridden).
      let pFmt = {};
      if (opts.styles) {
        const pPr = child(p, 'pPr'), ps = pPr && child(pPr, 'pStyle');
        pFmt = styleFmt(S, ps ? wattr(ps, 'val') : S.defaultPara);
      }
      // 1. split multi-content runs and build atoms
      const atoms = [];
      const touchedRuns = [];
      for (const { r, del } of byPara.get(p)) {
        if (!r.parentNode) continue;
        const rPr = child(r, 'rPr');
        const content = kids(r).filter((c) => !isW(c, 'rPr'));
        let pieces = [r];
        if (content.length > 1) {
          pieces = content.map((c) => {
            const nr = r.cloneNode(false);
            if (rPr) nr.appendChild(rPr.cloneNode(true));
            nr.appendChild(c);
            r.parentNode.insertBefore(nr, r);
            return nr;
          });
          r.parentNode.removeChild(r);
        }
        if (del) continue;
        // effective formatting
        let inherited = merge(opts.styles ? S.defaults : {}, pFmt);
        const rs = rPr && child(rPr, 'rStyle');
        if (rs && opts.styles) inherited = merge(inherited, styleFmt(S, wattr(rs, 'val')));
        const eff = merge(inherited, parseRPr(rPr));
        const set = new Set(enabled.filter((k) => eff[k]));
        for (const run of pieces) {
          const c = kids(run).find((x) => !isW(x, 'rPr'));
          if (!c || c.namespaceURI !== W) continue;
          const ln = c.localName;
          if (ln === 'fldChar') {
            const t = wattr(c, 'fldCharType');
            if (t === 'begin') fieldStack.push({ code: true });
            else if (t === 'separate' && fieldStack.length) fieldStack[fieldStack.length - 1].code = false;
            else if (t === 'end') fieldStack.pop();
            continue;
          }
          if (fieldStack.some((f) => f.code)) continue;
          if (CONTENT_SKIP.has(ln)) continue;
          const info = { run, inherited, set };
          if (ln === 't') {
            const txt = c.textContent;
            if (!txt.length) continue;
            touchedRuns.push(info);
            for (const ch of txt) atoms.push({ run, t: c, ch, ws: /\s/.test(ch), set: new Set(set), orig: set, inh: inherited });
          } else if (OPAQUE_WS[ln]) {
            touchedRuns.push(info);
            atoms.push({ run, ch: OPAQUE_WS[ln], ws: true, set: new Set(set), orig: set, opaque: true, inh: inherited });
          } else if (OPAQUE_TEXT.has(ln)) {
            touchedRuns.push(info);
            atoms.push({ run, ch: ln === 'sym' ? '◆' : ln.includes('ootnote') || ln.includes('ndnote') ? '¹' : '-', ws: false, set: new Set(set), orig: set, opaque: true, inh: inherited });
          } else if (NEUTRAL.has(ln)) {
            atoms.push({ run, ch: '▣', ws: false, set: new Set(), orig: new Set(), opaque: true, inh: inherited });
          }
        }
      }
      if (!atoms.length) continue;

      // 2. move edge whitespace outside of tags
      if (!opts.spacesInside) {
        for (const k of enabled) {
          let i = 0;
          while (i < atoms.length) {
            if (!atoms[i].set.has(k)) { i++; continue; }
            let j = i; while (j < atoms.length && atoms[j].set.has(k)) j++;
            let a = i; while (a < j && atoms[a].ws) atoms[a++].set.delete(k);
            let b = j - 1; while (b >= a && atoms[b].ws) atoms[b--].set.delete(k);
            i = j;
          }
        }
      }

      // 3. compute properly nested tag insertions
      const prefix = new Array(atoms.length).fill('');
      let suffix = '';
      const stack = [];
      let any = false;
      const spanLen = (k, from) => { let n = 0; while (from + n < atoms.length && atoms[from + n].set.has(k)) n++; return n; };
      for (let x = 0; x < atoms.length; x++) {
        const S2 = atoms[x].set;
        let j = stack.findIndex((k) => !S2.has(k));
        let s = '';
        if (j >= 0) { for (let y = stack.length - 1; y >= j; y--) s += `</${tags[stack[y]]}>`; stack.length = j; }
        const need = enabled.filter((k) => S2.has(k) && !stack.includes(k));
        need.sort((a, b) => spanLen(b, x) - spanLen(a, x) || KEYS.indexOf(a) - KEYS.indexOf(b));
        for (const k of need) { s += `<${tags[k]}>`; stack.push(k); stats.tags[k]++; any = true; }
        prefix[x] = s;
      }
      for (let y = stack.length - 1; y >= 0; y--) suffix += `</${tags[stack[y]]}>`;

      if (!any && !(opts.escape && atoms.some((a) => !a.opaque && /[<>&]/.test(a.ch)))) {
        if (opts.strip) for (const info of touchedRuns) if (stripRun(doc, info, enabled)) changed = true;
        continue;
      }
      changed = true;
      if (any) stats.changedParagraphs++;

      // preview capture
      if (any && previews.length < opts.previewLimit) {
        const segs = []; let out = '';
        for (let x = 0; x < atoms.length; x++) {
          const a = atoms[x];
          const key = [...a.orig].sort().join(',');
          if (segs.length && segs[segs.length - 1].key === key) segs[segs.length - 1].text += a.ch; else segs.push({ key, fmt: [...a.orig], text: a.ch });
          out += prefix[x] + esc(a.ch, opts.escape);
        }
        previews.push({ before: segs, after: out + suffix, part: opts.partLabel });
      }

      // 4. write text back
      const byT = new Map();
      atoms.forEach((a, x) => {
        if (a.opaque) {
          if (prefix[x]) insertTagRun(doc, a.run, prefix[x], 'before', opts, enabled, a);
          return;
        }
        if (!byT.has(a.t)) byT.set(a.t, '');
        byT.set(a.t, byT.get(a.t) + prefix[x] + esc(a.ch, opts.escape));
      });
      const last = atoms[atoms.length - 1];
      if (suffix) {
        if (last.opaque) insertTagRun(doc, last.run, suffix, 'after', opts, enabled, last);
        else byT.set(last.t, byT.get(last.t) + suffix);
      }
      for (const [t, txt] of byT) { t.textContent = txt; t.setAttributeNS(XML_NS, 'xml:space', 'preserve'); }
      if (opts.strip) for (const info of touchedRuns) stripRun(doc, info, enabled);
    }
    return changed;
  }
  function esc(ch, on) { if (!on) return ch; return ch === '<' ? '&lt;' : ch === '>' ? '&gt;' : ch === '&' ? '&amp;' : ch; }

  function insertTagRun(doc, refRun, text, where, opts, enabled, atom) {
    const nr = refRun.cloneNode(false);
    const rPr = child(refRun, 'rPr');
    if (rPr) nr.appendChild(rPr.cloneNode(true));
    const t = doc.createElementNS(W, 'w:t');
    t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
    t.textContent = text;
    nr.appendChild(t);
    refRun.parentNode.insertBefore(nr, where === 'before' ? refRun : refRun.nextSibling);
    if (opts.strip) stripRun(doc, { run: nr, inherited: atom.inh || {} }, enabled, true);
  }

  function stripRun(doc, info, enabled, quiet) {
    const { run, inherited } = info;
    let rPr = child(run, 'rPr');
    if (info.stripped) return false;
    info.stripped = true;
    let changed = false;
    const remove = (name, pred) => { const el = rPr && child(rPr, name); if (el && (!pred || pred(el))) { rPr.removeChild(el); changed = true; } };
    const addOff = (name, val) => {
      if (!rPr) { rPr = doc.createElementNS(W, 'w:rPr'); run.insertBefore(rPr, run.firstChild); }
      const el = doc.createElementNS(W, 'w:' + name); el.setAttributeNS(W, 'w:val', val); setRPrChild(rPr, el); changed = true;
    };
    const en = new Set(enabled);
    if (en.has('b')) { remove('b'); remove('bCs'); if (inherited.b) { addOff('b', '0'); addOff('bCs', '0'); } }
    if (en.has('i')) { remove('i'); remove('iCs'); if (inherited.i) { addOff('i', '0'); addOff('iCs', '0'); } }
    if (en.has('u')) { remove('u'); if (inherited.u) addOff('u', 'none'); }
    if (en.has('s')) { remove('strike'); remove('dstrike'); if (inherited.s) { addOff('strike', '0'); addOff('dstrike', '0'); } }
    if (en.has('sub') || en.has('sup')) {
      remove('vertAlign', (el) => { const v = wattr(el, 'val'); return (v === 'subscript' && en.has('sub')) || (v === 'superscript' && en.has('sup')); });
      if ((inherited.sub && en.has('sub')) || (inherited.sup && en.has('sup'))) addOff('vertAlign', 'baseline');
    }
    if (rPr && !rPr.firstChild) run.removeChild(rPr);
    return changed;
  }

  /* ---------------- public API ---------------- */
  const PART_RE = /^word\/(document\d*|header\d+|footer\d+|footnotes|endnotes)\.xml$/;
  function partLabel(name) {
    const m = name.match(PART_RE)[1];
    if (m.startsWith('document')) return 'Body';
    if (m.startsWith('header')) return 'Header';
    if (m.startsWith('footer')) return 'Footer';
    return m === 'footnotes' ? 'Footnotes' : 'Endnotes';
  }

  async function convert(buffer, opts, env = {}) {
    const DP = env.DOMParser || DOMParser, XS = env.XMLSerializer || XMLSerializer;
    opts = Object.assign({ formats: { b: true, i: true, u: true, s: true, sub: true, sup: true }, tagStyle: 'classic', strip: true, styles: true, notes: true, spacesInside: false, escape: false, previewLimit: 300 }, opts);
    const entries = readZip(buffer);
    const byName = Object.fromEntries(entries.map((e) => [e.name, e]));
    if (!Object.keys(byName).some((n) => /^word\/document\d*\.xml$/.test(n))) throw new UserError('This file has no Word document inside. Is it really a .docx?');
    const td = new TextDecoder(), te = new TextEncoder();
    const parse = (s) => { const d = new DP().parseFromString(s, 'application/xml'); const err = d.getElementsByTagName('parsererror')[0]; if (err) throw new UserError('The document contains XML that couldn’t be read.'); return d; };
    const S = byName['word/styles.xml'] ? loadStyles(parse(td.decode(await entryData(byName['word/styles.xml'])))) : loadStyles(null);
    const stats = { paragraphs: 0, changedParagraphs: 0, tags: Object.fromEntries(KEYS.map((k) => [k, 0])) };
    const previews = [];
    const order = (n) => (n.startsWith('word/document') ? 0 : 1);
    const targets = entries.filter((e) => PART_RE.test(e.name) && (opts.notes || e.name.startsWith('word/document'))).sort((a, b) => order(a.name) - order(b.name));
    for (const en of targets) {
      let xml = td.decode(await entryData(en));
      const decl = xml.match(/^﻿?<\?xml[^>]*\?>\s*/);
      const doc = parse(xml);
      if (processPart(doc, S, Object.assign({}, opts, { partLabel: partLabel(en.name) }), stats, previews)) {
        let out = new XS().serializeToString(doc);
        if (!/^<\?xml/.test(out)) out = (decl ? decl[0].replace(/^﻿/, '').trimEnd() : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>') + '\r\n' + out;
        en.newData = te.encode(out);
      }
    }
    const blob = await writeZip(entries);
    return { blob, stats, previews };
  }

  return { convert, UserError, KEYS, TAGSETS, _readZip: readZip };
})();
if (typeof module !== 'undefined') module.exports = Taglift;
