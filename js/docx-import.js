/* ==========================================================
   Schreibwerk – Import von Word-Dokumenten (.docx)
   ========================================================== */
(function(){
'use strict';
const SW = window.SW;
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const esc = SW.esc;
const TW_CM = 2.54 / 1440;

const kids = (el, name) => el ? Array.from(el.childNodes).filter(n => n.nodeType === 1 && (!name || n.localName === name)) : [];
const kid = (el, name) => kids(el, name)[0] || null;
const attr = (el, name) => el ? (el.getAttributeNS(W, name) ?? el.getAttribute('w:' + name)) : null;
const onOff = el => { if (!el) return null; const v = attr(el, 'val'); return !(v === '0' || v === 'false' || v === 'off' || v === 'none'); };

const HL = {yellow:'#ffff00', green:'#00ff00', cyan:'#00ffff', magenta:'#ff00ff', blue:'#0000ff', red:'#ff0000', darkBlue:'#000080', darkCyan:'#008080', darkGreen:'#008000', darkMagenta:'#800080', darkRed:'#800000', darkYellow:'#808000', darkGray:'#808080', lightGray:'#c0c0c0', black:'#000000', white:'#ffffff'};

function parseXml(s){ return new DOMParser().parseFromString(s, 'application/xml'); }

async function readText(zip, path){ const f = zip.file(path); return f ? await f.async('string') : null; }

/* ---------- Formatvorlagen ---------- */
function readStyles(xml, theme){
  const fontOf = f => attr(f, 'ascii') || attr(f, 'hAnsi') || ((attr(f, 'asciiTheme') || attr(f, 'hAnsiTheme') || '').startsWith('major') ? theme.major : (attr(f, 'asciiTheme') || attr(f, 'hAnsiTheme')) ? theme.minor : null);
  const map = {};
  const defaults = {font:null, size:null, after:null, line:null};
  if (!xml) return {map, defaults};
  const d = parseXml(xml);
  const dd = d.getElementsByTagNameNS(W, 'docDefaults')[0];
  if (dd){
    const rp = dd.getElementsByTagNameNS(W, 'rPr')[0];
    if (rp){
      const f = kid(rp, 'rFonts'); if (f) defaults.font = fontOf(f);
      const sz = kid(rp, 'sz'); if (sz) defaults.size = +attr(sz, 'val') / 2;
    }
    const pp = dd.getElementsByTagNameNS(W, 'pPr')[0];
    if (pp){
      const sp = kid(pp, 'spacing');
      if (sp){ if (attr(sp, 'after') != null) defaults.after = +attr(sp, 'after') / 20; if (attr(sp, 'line') && (attr(sp, 'lineRule') || 'auto') === 'auto') defaults.line = +attr(sp, 'line') / 240; }
    }
  }
  for (const s of d.getElementsByTagNameNS(W, 'style')){
    const id = attr(s, 'styleId');
    const nameEl = kid(s, 'name');
    const based = kid(s, 'basedOn');
    map[id] = {
      id, type:attr(s, 'type'), name:(nameEl ? attr(nameEl, 'val') : id) || '',
      basedOn:based ? attr(based, 'val') : null,
      pPr:kid(s, 'pPr'), rPr:kid(s, 'rPr'), isDefault:attr(s, 'default') === '1'
    };
  }
  const normal = Object.values(map).find(s => s.type === 'paragraph' && s.isDefault);
  if (normal){
    if (normal.rPr){
      const f = kid(normal.rPr, 'rFonts'); if (f && fontOf(f)) defaults.font = fontOf(f);
      const sz = kid(normal.rPr, 'sz'); if (sz) defaults.size = +attr(sz, 'val') / 2;
    }
    if (normal.pPr){
      const sp = kid(normal.pPr, 'spacing');
      if (sp){ if (attr(sp, 'after') != null) defaults.after = +attr(sp, 'after') / 20; if (attr(sp, 'line') && (attr(sp, 'lineRule') || 'auto') === 'auto') defaults.line = +attr(sp, 'line') / 240; }
    }
  }
  return {map, defaults};
}
/* Formatvorlage → Schreibwerk-Absatztyp */
function classifyStyle(styles, id){
  let s = styles.map[id], guard = 0;
  while (s && guard++ < 10){
    const n = s.name.toLowerCase().replace(/\s+/g, ' ');
    let m;
    if ((m = n.match(/^(heading|überschrift) (\d)$/))) return {tag:'h' + Math.min(3, +m[2])};
    if (n === 'title' || n === 'titel') return {tag:'p', cls:'sw-title'};
    if (n === 'caption' || n === 'beschriftung') return {tag:'p', cls:'sw-caption'};
    if (n === 'subtitle' || n === 'untertitel') return {tag:'p', cls:'sw-subtitle'};
    if (n === 'quote' || n === 'zitat' || n === 'intense quote' || n === 'intensives zitat') return {tag:'p', cls:'sw-quote'};
    if (n === 'no spacing' || n === 'kein leerraum') return {tag:'p', cls:'sw-nospace'};
    if (n === 'toc heading' || n === 'inhaltsverzeichnisüberschrift') return {tag:'h1'};
    if (s.pPr){ const ol = kid(s.pPr, 'outlineLvl'); if (ol && +attr(ol, 'val') < 3) return {tag:'h' + (+attr(ol, 'val') + 1)}; }
    if (s.isDefault) return null;
    s = s.basedOn ? styles.map[s.basedOn] : null;
  }
  return null;
}
/* Formatkette einer benutzerdefinierten Vorlage auflösen (ohne Standard) */
function styleChain(styles, id){
  const chain = [];
  let s = styles.map[id], guard = 0;
  while (s && guard++ < 10 && !s.isDefault){ chain.unshift(s); s = s.basedOn ? styles.map[s.basedOn] : null; }
  return chain;
}

/* ---------- Nummerierung ---------- */
function readNumbering(xml){
  const nums = {};
  if (!xml) return nums;
  const d = parseXml(xml);
  const abs = {};
  for (const a of d.getElementsByTagNameNS(W, 'abstractNum')){
    const lv = {};
    for (const l of kids(a, 'lvl')){
      const nf = kid(l, 'numFmt'), lt = kid(l, 'lvlText'), st = kid(l, 'start');
      lv[attr(l, 'ilvl')] = {fmt:nf ? attr(nf, 'val') : 'decimal', text:lt ? attr(lt, 'val') : '', start:st ? +attr(st, 'val') : 1};
    }
    abs[attr(a, 'abstractNumId')] = lv;
  }
  for (const n of d.getElementsByTagNameNS(W, 'num')){
    const a = kid(n, 'abstractNumId');
    const levels = JSON.parse(JSON.stringify(abs[attr(a, 'val')] || {}));
    for (const o of kids(n, 'lvlOverride')){
      const so = kid(o, 'startOverride');
      if (so && levels[attr(o, 'ilvl')]) levels[attr(o, 'ilvl')].start = +attr(so, 'val');
    }
    Object.defineProperty(levels, '_abs', {value:attr(a, 'val')});
    Object.defineProperty(levels, '_override', {value:kids(n, 'lvlOverride').length > 0});
    nums[attr(n, 'numId')] = levels;
  }
  return nums;
}
function listInfo(lvl){
  if (!lvl) return {tag:'ul', attrs:''};
  if (lvl.fmt === 'bullet' || lvl.fmt === 'none'){
    const t = lvl.text || '';
    let b = 'disc';
    if (t === 'o') b = 'circle';
    else if (t === '' || t === '▪' || t === '■') b = 'square';
    else if (t === '–' || t === '-' || t === '—') b = 'dash';
    else if (t === '' || t === '➢' || t === '') b = 'arrow';
    else if (t === '' || t === '✓') b = 'check';
    const ls = {circle:'list-style-type:circle', square:'list-style-type:square'}[b];
    return {tag:'ul', attrs:` data-bullet="${b}"${ls ? ` style="${ls}"` : ''}`};
  }
  let f = 'decimal';
  const paren = /\)$/.test(lvl.text || '');
  if (lvl.fmt === 'lowerLetter') f = 'lower-alpha';
  else if (lvl.fmt === 'upperLetter') f = 'upper-alpha';
  else if (lvl.fmt === 'lowerRoman') f = 'lower-roman';
  else if (lvl.fmt === 'upperRoman') f = 'upper-roman';
  else if (paren) f = 'decimal-paren';
  const ls = {'upper-alpha':'upper-alpha','lower-roman':'lower-roman','upper-roman':'upper-roman'}[f];
  return {tag:'ol', attrs:` data-fmt="${f}"${ls ? ` style="list-style-type:${ls}"` : ''}${lvl.start > 1 ? ` start="${lvl.start}"` : ''}`};
}

/* ---------- Haupt-Import ---------- */
SW.importDocx = async function(buf){
  const zip = await JSZip.loadAsync(buf);
  const docPath = 'word/document.xml';
  const docXml = await readText(zip, docPath);
  if (!docXml) throw new Error('Keine gültige Word-Datei (word/document.xml fehlt).');

  // 1) Schreibwerk-Quelle vorhanden und unverändert?
  for (const name of Object.keys(zip.files)){
    if (!/^customXml\/item\d+\.xml$/.test(name)) continue;
    const t = await zip.file(name).async('string');
    if (!t.includes('urn:schreibwerk:v1')) continue;
    try {
      const b64 = t.replace(/^[\s\S]*?<schreibwerk[^>]*>/, '').replace(/<\/schreibwerk>[\s\S]*$/, '').trim();
      const bin = atob(b64);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      const src = JSON.parse(new TextDecoder().decode(bytes));
      if (src.hash === SW.hashStr(docXml)){
        const images = [];
        for (const [id, path] of Object.entries(src.images || {})){
          const f = zip.file(path);
          if (f){
            const ext = path.split('.').pop();
            const blob = new Blob([await f.async('uint8array')], {type:'image/' + (ext === 'jpg' ? 'jpeg' : ext)});
            images.push({id, blob});
          }
        }
        return {html:src.html, doc:src.doc, images, exact:true};
      }
    } catch(e){ /* weiter mit normalem Import */ }
  }

  // 2) Allgemeiner Word-Import
  const relsXml = await readText(zip, 'word/_rels/document.xml.rels');
  const rels = {};
  if (relsXml) for (const r of parseXml(relsXml).getElementsByTagName('Relationship'))
    rels[r.getAttribute('Id')] = {target:r.getAttribute('Target'), mode:r.getAttribute('TargetMode'), type:r.getAttribute('Type')};
  const theme = {minor:'Calibri', major:'Calibri Light'};
  const themeXml = await readText(zip, 'word/theme/theme1.xml');
  if (themeXml){
    const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const td = parseXml(themeXml);
    const mi = td.getElementsByTagNameNS(A, 'minorFont')[0], ma = td.getElementsByTagNameNS(A, 'majorFont')[0];
    const lat = el => el && el.getElementsByTagNameNS(A, 'latin')[0] ? el.getElementsByTagNameNS(A, 'latin')[0].getAttribute('typeface') : null;
    theme.minor = lat(mi) || theme.minor; theme.major = lat(ma) || theme.major;
  }
  const styles = readStyles(await readText(zip, 'word/styles.xml'), theme);
  const numbering = readNumbering(await readText(zip, 'word/numbering.xml'));
  const footnotes = {};
  const fnXml = await readText(zip, 'word/footnotes.xml');
  if (fnXml) for (const f of parseXml(fnXml).getElementsByTagNameNS(W, 'footnote')){
    const id = attr(f, 'id');
    footnotes[id] = Array.from(f.getElementsByTagNameNS(W, 't')).map(t => t.textContent).join('').trim();
  }
  const comments = {};
  const cmXml = await readText(zip, 'word/comments.xml');
  if (cmXml) for (const c of parseXml(cmXml).getElementsByTagNameNS(W, 'comment')){
    comments[attr(c, 'id')] = {
      author:attr(c, 'author') || 'Autor', date:attr(c, 'date') || new Date().toISOString(), done:false,
      text:Array.from(c.getElementsByTagNameNS(W, 'p')).map(p => Array.from(p.getElementsByTagNameNS(W, 't')).map(t => t.textContent).join('')).join('\n').trim()
    };
  }
  const usedComments = {};
  const images = [];
  const imgCache = {};
  async function imageFor(rid){
    if (imgCache[rid]) return imgCache[rid];
    const rel = rels[rid]; if (!rel) return null;
    const path = rel.target.startsWith('/') ? rel.target.slice(1) : 'word/' + rel.target.replace(/^\.\//, '');
    const f = zip.file(path); if (!f) return null;
    const ext = path.split('.').pop().toLowerCase();
    if (!['png','jpg','jpeg','gif','bmp','webp'].includes(ext)) return null;
    let blob = new Blob([await f.async('uint8array')], {type:'image/' + (ext === 'jpg' ? 'jpeg' : ext)});
    if (ext === 'bmp' || ext === 'webp') blob = await SW.normalizeImageBlob(blob);
    const id = 'img' + SW.uid();
    images.push({id, blob});
    imgCache[rid] = id;
    return id;
  }

  const d = parseXml(docXml);
  const body = d.getElementsByTagNameNS(W, 'body')[0];
  const allSects = Array.from(body.getElementsByTagNameNS(W, 'sectPr'));
  const sect = kids(body, 'sectPr')[0] || allSects[0];
  const firstSect = allSects.length ? (allSects.find(x => x !== kids(body, 'sectPr')[0]) || allSects[0]) : null;
  const doc = SW.defaultDoc();
  if (styles.defaults.font) doc.base.font = styles.defaults.font;
  if (styles.defaults.size) doc.base.size = styles.defaults.size;
  if (styles.defaults.after != null) doc.base.after = styles.defaults.after;
  if (styles.defaults.line) doc.base.line = Math.round(styles.defaults.line * 100) / 100;
  if (!styles.defaults.font) doc.base.font = 'Times New Roman';
  if (sect){
    const pg = kid(sect, 'pgSz'), mar = kid(sect, 'pgMar');
    if (pg){
      let w = +attr(pg, 'w') * TW_CM, h = +attr(pg, 'h') * TW_CM;
      if (w > h) { doc.orientation = 'landscape'; [w, h] = [h, w]; }
      let best = 'A4', bd = 1e9;
      for (const [k, v] of Object.entries(SW.PAGE_SIZES)){ const dd = Math.abs(v[0] - w) + Math.abs(v[1] - h); if (dd < bd){ bd = dd; best = k; } }
      doc.pageSize = best;
    }
    if (mar){
      const c = n => Math.round(Math.abs(+attr(mar, n) || 0) * TW_CM * 100) / 100;
      doc.margins = {top:c('top'), bottom:c('bottom'), left:c('left'), right:c('right')};
    }
    // Kopf-/Fußzeilen
    const pnt = kid(firstSect || sect, 'pgNumType') || kid(sect, 'pgNumType');
    if (pnt){
      const f = attr(pnt, 'fmt'), st = attr(pnt, 'start');
      doc.pageNum = {start:st != null ? +st : 1, fmt:['lowerRoman','upperRoman','lowerLetter','upperLetter'].includes(f) ? f : 'decimal'};
    }
    if (kid(firstSect || sect, 'titlePg') && onOff(kid(firstSect || sect, 'titlePg'))) doc.titlePage = true;
    const hf = async (refName, key, type = 'default') => {
      let refs = [];
      for (const sx of [sect, ...allSects]){ refs = kids(sx, refName); if (refs.some(r => (attr(r, 'type') || 'default') === type)) break; }
      const ref = refs.find(r => (attr(r, 'type') || 'default') === type) || (type === 'default' ? refs.find(r => attr(r, 'type') !== 'first') : null);
      if (!ref) return;
      const rel = rels[ref.getAttributeNS(R, 'id')];
      if (!rel) return;
      const xml = await readText(zip, 'word/' + rel.target);
      if (!xml) return;
      const hd = parseXml(xml);
      const paras = Array.from(hd.getElementsByTagNameNS(W, 'p'));
      let align = 'left';
      const lines = paras.map(p => {
        const jc = p.getElementsByTagNameNS(W, 'jc')[0];
        if (jc) align = {center:'center', right:'right', end:'right'}[attr(jc, 'val')] || align;
        let s = '', inInstr = '', mode = 0;
        const walk = el => {
          for (const n of kids(el)){
            const ln = n.localName;
            if (ln === 'fldSimple'){ const ins = (attr(n, 'instr') || '').trim().split(/\s+/)[0].toUpperCase(); s += ins === 'PAGE' ? '{SEITE}' : ins === 'NUMPAGES' ? '{SEITEN}' : n.textContent; continue; }
            if (ln === 'fldChar'){ const t = attr(n, 'fldCharType'); if (t === 'begin'){ mode = 1; inInstr = ''; } else if (t === 'separate'){ mode = 2; const ins = inInstr.trim().split(/\s+/)[0].toUpperCase(); if (ins === 'PAGE') s += '{SEITE}'; else if (ins === 'NUMPAGES' || ins === 'SECTIONPAGES') s += '{SEITEN}'; else mode = 3; } else mode = 0; continue; }
            if (ln === 'instrText'){ inInstr += n.textContent; continue; }
            if (ln === 't'){ if (mode === 0 || mode === 3) s += n.textContent; continue; }
            if (ln === 'tab'){ if (mode === 0) s += '   '; continue; }
            if (ln === 'pPr' || ln === 'rPr') continue;
            walk(n);
          }
        };
        walk(p);
        return s;
      }).filter(Boolean);
      doc[key] = {text:lines.join('\n').trim(), align};
    };
    await hf('headerReference', 'header');
    await hf('footerReference', 'footer');
    if (doc.titlePage){ await hf('headerReference', 'firstHeader', 'first'); await hf('footerReference', 'firstFooter', 'first'); }
  }

  /* ---------- Absätze ---------- */
  const fld = {stack:[], skipToc:false, tocEmitted:false, cmt:[]};
  const MNS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
  /* Beschriftungsnummer: vorangehendes „Abbildung “ in das Feld übernehmen */
  function seqHtml(h, type){
    const re = new RegExp(type.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(\\s|&nbsp;)*((</[a-z]+>)*)$');
    const m = h.match(re);
    if (m) h = h.slice(0, m.index) + (m[2] || '');
    return h + `<span class="sw-seq" data-seq="${esc(type)}" contenteditable="false">${esc(type)} 1</span>`;
  }
  /* OMML (Word-Formel) → LaTeX für den Formel-Editor */
  const NARY_TEX = {'∑':'\\sum', '∏':'\\prod', '∐':'\\coprod', '∫':'\\int', '∬':'\\iint', '∭':'\\iiint', '∮':'\\oint', '⋃':'\\bigcup', '⋂':'\\bigcap'};
  const ACC_TEX = {'\u0302':'\\hat', '^':'\\hat', '\u20D7':'\\vec', '→':'\\vec', '\u0307':'\\dot', '˙':'\\dot', '\u0308':'\\ddot', '\u0303':'\\tilde', '~':'\\tilde', '\u0305':'\\overline', '¯':'\\overline', '\u030C':'\\check'};
  const FUNCS = ['sin','cos','tan','cot','ln','log','exp','lim','max','min','sinh','cosh','tanh','arcsin','arccos','arctan','det','dim'];
  function mk(el, name){ return el ? Array.from(el.children).find(c => c.localName === name && c.namespaceURI === MNS) : null; }
  function mval(el, name){ const p = mk(el, name); return p ? (p.getAttributeNS(MNS, 'val') ?? p.getAttribute('m:val')) : null; }
  function texEsc(t){ return t.replace(/\\/g, '\\backslash ').replace(/([{}%#&_$])/g, '\\$1').replace(/\^/g, '\\hat{}'); }
  function grp(t){ t = t.trim(); return t.length === 1 ? t : `{${t}}`; }
  function ommlToTex(el){
    if (!el) return '';
    let out = '';
    for (const c of Array.from(el.children)){
      if (c.namespaceURI !== MNS) continue;
      const k = c.localName;
      const pr = mk(c, k + 'Pr');
      const sub = n => ommlToTex(mk(c, n));
      switch (k){
        case 'r': {
          const t = Array.from(c.getElementsByTagNameNS(MNS, 't')).map(x => x.textContent).join('');
          const plain = mval(mk(c, 'rPr'), 'sty') === 'p' || (mk(c, 'rPr') && mk(mk(c, 'rPr'), 'nor'));
          if (plain && /^[A-Za-zÄÖÜäöüß]{2,}$/.test(t.trim())) out += FUNCS.includes(t.trim()) ? `\\${t.trim()} ` : `\\text{${t.trim()}}`;
          else if (/^[A-Za-zÄÖÜäöüß ]{4,}$/.test(t) && / /.test(t)) out += `\\text{${t}}`;
          else out += texEsc(t);
          break;
        }
        case 'f': out += mval(pr, 'type') === 'noBar' ? `\\binom{${sub('num')}}{${sub('den')}}` : `\\frac{${sub('num')}}{${sub('den')}}`; break;
        case 'sSup': out += `${grp(sub('e'))}^{${sub('sup')}}`; break;
        case 'sSub': out += `${grp(sub('e'))}_{${sub('sub')}}`; break;
        case 'sSubSup': out += `${grp(sub('e'))}_{${sub('sub')}}^{${sub('sup')}}`; break;
        case 'sPre': out += `{}_{${sub('sub')}}^{${sub('sup')}}${grp(sub('e'))}`; break;
        case 'rad': { const deg = sub('deg').trim(); out += deg && mval(pr, 'degHide') !== '1' && mval(pr, 'degHide') !== 'on' ? `\\sqrt[${deg}]{${sub('e')}}` : `\\sqrt{${sub('e')}}`; break; }
        case 'nary': {
          const ch = mval(pr, 'chr') || '∫';
          const s1 = sub('sub').trim(), s2 = sub('sup').trim();
          out += (NARY_TEX[ch] || ch) + (s1 ? `_{${s1}}` : '') + (s2 ? `^{${s2}}` : '') + ' ' + sub('e');
          break;
        }
        case 'd': {
          let b = mval(pr, 'begChr'), e = mval(pr, 'endChr');
          b = b == null ? '(' : b; e = e == null ? ')' : e;
          const f = x => !x ? '.' : x === '{' ? '\\{' : x === '}' ? '\\}' : x === '|' ? '|' : x === '‖' ? '\\|' : x === '⟨' ? '\\langle' : x === '⟩' ? '\\rangle' : x;
          const parts = Array.from(c.children).filter(x => x.localName === 'e').map(x => ommlToTex(x));
          const sep = mval(pr, 'sepChr') || '|';
          out += `\\left${f(b)} ${parts.join(` ${sep} `)} \\right${f(e)}`;
          break;
        }
        case 'acc': { const ch = mval(pr, 'chr') || '\u0302'; out += `${ACC_TEX[ch] || '\\hat'}{${sub('e')}}`; break; }
        case 'bar': out += (mval(pr, 'pos') === 'top' ? '\\overline' : '\\underline') + `{${sub('e')}}`; break;
        case 'func': {
          const name = sub('fName').trim().replace(/^\\text\{(.*)\}$/, '$1').replace(/^\\(\w+)\s*$/, '$1');
          out += (FUNCS.includes(name) ? `\\${name}` : `\\operatorname{${name}}`) + ' ' + sub('e');
          break;
        }
        case 'limLow': { const e = sub('e').trim(); out += /^\\?lim\s*$/.test(e) || e === '\\lim' ? `\\lim_{${sub('lim')}}` : `\\underset{${sub('lim')}}{${e}}`; break; }
        case 'limUpp': out += `\\overset{${sub('lim')}}{${sub('e')}}`; break;
        case 'groupChr': out += (mval(pr, 'pos') === 'top' ? '\\overbrace' : '\\underbrace') + `{${sub('e')}}`; break;
        case 'm': {
          const rows = Array.from(c.children).filter(x => x.localName === 'mr').map(r => Array.from(r.children).filter(x => x.localName === 'e').map(x => ommlToTex(x)).join(' & '));
          out += `\\begin{matrix} ${rows.join(' \\\\ ')} \\end{matrix}`;
          break;
        }
        case 'eqArr': {
          const rows = Array.from(c.children).filter(x => x.localName === 'e').map(x => ommlToTex(x));
          out += `\\begin{aligned} ${rows.join(' \\\\ ')} \\end{aligned}`;
          break;
        }
        case 'box': case 'borderBox': case 'phant': case 'e': case 'oMath': out += sub('e') || ommlToTex(c); break;
        default:
          if (!/Pr$/.test(k)) out += ommlToTex(c);
      }
    }
    return out;
  }

  function runStyle(rp){
    const st = [];
    const wrap = [];
    if (!rp) return {st, wrap};
    const b = kid(rp, 'b'); if (b && onOff(b)) wrap.push('b'); else if (b && !onOff(b)) st.push('font-weight:normal');
    const i = kid(rp, 'i'); if (i && onOff(i)) wrap.push('i');
    const u = kid(rp, 'u'); if (u && attr(u, 'val') !== 'none') wrap.push('u');
    if ((kid(rp, 'strike') && onOff(kid(rp, 'strike'))) || (kid(rp, 'dstrike') && onOff(kid(rp, 'dstrike')))) wrap.push('s');
    const va = kid(rp, 'vertAlign');
    if (va){ const v = attr(va, 'val'); if (v === 'superscript') wrap.push('sup'); else if (v === 'subscript') wrap.push('sub'); }
    const c = kid(rp, 'color'); if (c && attr(c, 'val') && attr(c, 'val') !== 'auto') st.push('color:#' + attr(c, 'val'));
    const sz = kid(rp, 'sz'); if (sz) st.push('font-size:' + (+attr(sz, 'val') / 2) + 'pt');
    const f = kid(rp, 'rFonts'); if (f && (attr(f, 'ascii') || attr(f, 'hAnsi'))) st.push(`font-family:"${attr(f, 'ascii') || attr(f, 'hAnsi')}"`);
    const hl = kid(rp, 'highlight'); if (hl && HL[attr(hl, 'val')]) st.push('background-color:' + HL[attr(hl, 'val')]);
    const sh = kid(rp, 'shd'); if (sh && attr(sh, 'fill') && !/^(auto|FFFFFF)$/i.test(attr(sh, 'fill')) && !hl) st.push('background-color:#' + attr(sh, 'fill'));
    const caps = kid(rp, 'caps'); if (caps && onOff(caps)) st.push('text-transform:uppercase');
    return {st, wrap};
  }
  function mergeStyle(base, extra){
    return {st:[...base.st, ...extra.st.filter(s => !base.st.some(b => b.split(':')[0] === s.split(':')[0]))].concat(), wrap:[...new Set([...base.wrap, ...extra.wrap])]};
  }
  function wrapHtml(inner, rs){
    if (!inner) return '';
    let h = inner;
    rs.wrap.forEach(t => { h = `<${t}>${h}</${t}>`; });
    if (rs.st.length) h = `<span style="${esc(rs.st.join(';'))}">${h}</span>`;
    return h;
  }

  async function runsHtml(parent, baseRs, out){
    // liefert HTML; out.breaks sammelt Seitenumbrüche
    let h = '';
    for (const n of kids(parent)){
      const ln = n.localName;
      if (ln === 'pPr') continue;
      if (ln === 'r'){
        const rp = kid(n, 'rPr');
        let rs = runStyle(rp);
        const rst = rp && kid(rp, 'rStyle');
        if (rst){
          const chain = styleChain(styles, attr(rst, 'val'));
          let cs = {st:[], wrap:[]};
          chain.forEach(s => { cs = mergeStyle(runStyle(s.rPr), cs); });
          if (!/hyperlink/i.test(attr(rst, 'val'))) rs = mergeStyle(rs, cs);
        }
        rs = mergeStyle(rs, baseRs);
        let inner = '';
        for (const c of kids(n)){
          const cl = c.localName;
          if (cl === 'fldChar'){
            const t = attr(c, 'fldCharType');
            if (t === 'begin') fld.stack.push({instr:'', phase:1});
            else if (t === 'separate'){
              const f = fld.stack[fld.stack.length - 1];
              if (f){
                f.phase = 2;
                if (/^\s*TOC\b/i.test(f.instr)){
                  f.toc = true; fld.skipToc = true;
                  const cm = f.instr.match(/\\c\s+"?([^"\\]+?)"?\s*(\\|$)/);
                  if (cm) out.tof = cm[1].trim();
                  else if (!fld.tocEmitted){ fld.tocEmitted = true; out.toc = true; }
                }
                const sq = f.instr.match(/^\s*SEQ\s+(\S+)/i);
                if (sq){ f.seq = sq[1].replace(/"/g, ''); h += wrapHtml(inner, rs); inner = ''; h = seqHtml(h, f.seq); }
              }
            }
            else if (t === 'end'){ const f = fld.stack.pop(); if (f && f.toc) fld.skipToc = false; }
            continue;
          }
          if (cl === 'instrText'){ const f = fld.stack[fld.stack.length - 1]; if (f) f.instr += c.textContent; continue; }
          const inInstr = fld.stack.some(f => f.phase === 1 || (f.phase === 2 && f.seq));
          if (inInstr || fld.skipToc) continue;
          if (cl === 't') inner += esc(c.textContent).replace(/ {2}/g, ' &nbsp;');
          else if (cl === 'tab') inner += '<span class="sw-tab">\t</span>';
          else if (cl === 'br' || cl === 'cr'){
            const ty = attr(c, 'type');
            if (ty === 'page'){ h += wrapHtml(inner, rs); inner = ''; h += '\u0000PB\u0000'; }
            else inner += '<br>';
          }
          else if (cl === 'noBreakHyphen') inner += '‑';
          else if (cl === 'softHyphen') inner += '&shy;';
          else if (cl === 'sym'){ const code = parseInt(attr(c, 'char') || '20', 16); inner += esc(String.fromCharCode(code >= 0xF000 ? code - 0xF000 : code)); }
          else if (cl === 'footnoteReference'){
            h += wrapHtml(inner, rs); inner = '';
            h += `<sup class="sw-fn" contenteditable="false" data-note="${esc(footnotes[attr(c, 'id')] || '')}">*</sup>`;
          }
          else if (cl === 'drawing' || cl === 'pict' || cl === 'object'){
            const blip = c.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'blip')[0];
            let rid = blip ? blip.getAttributeNS(R, 'embed') : null;
            if (!rid){ const vi = c.getElementsByTagNameNS('urn:schemas-microsoft-com:vml', 'imagedata')[0]; if (vi) rid = vi.getAttributeNS(R, 'id'); }
            const ext = c.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', 'extent')[0];
            const wpx = ext ? Math.round(+ext.getAttribute('cx') / 9525) : null;
            const id = rid ? await imageFor(rid) : null;
            const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
            const anc = c.getElementsByTagNameNS(WP, 'anchor')[0];
            let wrapAttr = '', posCss = '';
            if (anc){
              const has = n2 => anc.getElementsByTagNameNS(WP, n2).length > 0;
              const ph = anc.getElementsByTagNameNS(WP, 'positionH')[0], pv = anc.getElementsByTagNameNS(WP, 'positionV')[0];
              const al = ph && ph.getElementsByTagNameNS(WP, 'align')[0];
              const ox = ph && ph.getElementsByTagNameNS(WP, 'posOffset')[0], oy = pv && pv.getElementsByTagNameNS(WP, 'posOffset')[0];
              const xpx = ox ? Math.round(+ox.textContent / 9525) : 0, ypx = oy ? Math.round(+oy.textContent / 9525) : 0;
              const contentPx = ((SW.PAGE_SIZES[doc.pageSize] || [21])[doc.orientation === 'landscape' ? 1 : 0] - doc.margins.left - doc.margins.right) * SW.CM;
              if (has('wrapSquare') || has('wrapTight') || has('wrapThrough')){
                const right = al ? /right|outside/.test(al.textContent) : xpx + (wpx || 0) / 2 > contentPx / 2;
                wrapAttr = ` data-wrap="${right ? 'right' : 'left'}"`;
              } else if (has('wrapTopAndBottom')) wrapAttr = ' data-wrap="topbottom"';
              else {
                wrapAttr = ` data-wrap="${anc.getAttribute('behindDoc') === '1' ? 'behind' : 'front'}"`;
                const left = al ? (/center/.test(al.textContent) ? (contentPx - (wpx || 0)) / 2 : /right/.test(al.textContent) ? contentPx - (wpx || 0) : 0) : xpx;
                posCss = `;left:${Math.round(left)}px;top:${ypx}px`;
              }
            }
            if (id) inner += `<img data-img-id="${id}" alt=""${wrapAttr}${wpx || posCss ? ` style="${wpx ? `width:${wpx}px` : ''}${posCss}"` : ''}>`;
            else {
              // Textfelder in Formen: Text übernehmen
              const tx = Array.from(c.getElementsByTagNameNS(W, 'txbxContent'));
              if (tx.length) inner += esc(tx.map(t => Array.from(t.getElementsByTagNameNS(W, 't')).map(x => x.textContent).join('')).join(' '));
            }
          }
        }
        const piece = wrapHtml(inner, rs);
        h += piece && fld.cmt.length ? `<span class="sw-cmt" data-cid="w${fld.cmt[fld.cmt.length - 1]}">${piece}</span>` : piece;
        continue;
      }
      if (ln === 'hyperlink'){
        const rid = n.getAttributeNS(R, 'id');
        const anchor = attr(n, 'anchor');
        const href = rid && rels[rid] ? rels[rid].target : (anchor ? '#' + anchor : null);
        const inner = await runsHtml(n, baseRs, out);
        if (fld.skipToc) continue;
        h += href ? `<a href="${esc(href)}">${inner}</a>` : inner;
        continue;
      }
      if (ln === 'fldSimple'){
        const instr = attr(n, 'instr') || '';
        const sq = instr.match(/^\s*SEQ\s+(\S+)/i);
        if (sq){ h = seqHtml(h, sq[1].replace(/"/g, '')); continue; }
        h += await runsHtml(n, baseRs, out); continue;
      }
      if (ln === 'commentRangeStart'){ const id = attr(n, 'id'); if (comments[id]){ fld.cmt.push(id); usedComments[id] = true; } continue; }
      if (ln === 'commentRangeEnd'){ const id = attr(n, 'id'); fld.cmt = fld.cmt.filter(x => x !== id); continue; }
      if (ln === 'oMath' || ln === 'oMathPara'){
        const maths = ln === 'oMathPara' ? Array.from(n.getElementsByTagNameNS(MNS, 'oMath')) : [n];
        maths.forEach(m => { const tex = ommlToTex(m).trim(); if (tex) h += `<span class="sw-math" contenteditable="false" data-tex="${esc(tex)}">${esc(tex)}</span>`; });
        if (ln === 'oMathPara') out.mathPara = true;
        continue;
      }
      if (ln === 'smartTag' || ln === 'ins' || ln === 'sdt' || ln === 'customXml' || ln === 'sdtContent' || ln === 'moveTo'){
        const c = ln === 'sdt' ? kid(n, 'sdtContent') : n;
        if (c) h += await runsHtml(c, baseRs, out);
        continue;
      }
      // del, moveFrom, bookmarkStart … ignorieren
    }
    return h;
  }

  async function paraHtml(p){
    const pp = kid(p, 'pPr');
    let tag = 'p', cls = '', css = [], attrs = '';
    let baseRs = {st:[], wrap:[]};
    let num = null, pageBreakBefore = false;
    const applyPPr = (pr) => {
      if (!pr) return;
      const jc = kid(pr, 'jc');
      if (jc){ const a = {center:'center', right:'right', end:'right', both:'justify', distribute:'justify', left:'', start:''}[attr(jc, 'val')]; setCss('text-align', a); }
      const ind = kid(pr, 'ind');
      if (ind){
        const l = attr(ind, 'left') ?? attr(ind, 'start'), r = attr(ind, 'right') ?? attr(ind, 'end');
        const fl = attr(ind, 'firstLine'), hg = attr(ind, 'hanging');
        if (l != null) setCss('margin-left', +l ? (+l * TW_CM).toFixed(2) + 'cm' : '');
        if (r != null) setCss('margin-right', +r ? (+r * TW_CM).toFixed(2) + 'cm' : '');
        if (fl != null && +fl) setCss('text-indent', (+fl * TW_CM).toFixed(2) + 'cm');
        if (hg != null && +hg) setCss('text-indent', (-hg * TW_CM).toFixed(2) + 'cm');
      }
      const sp = kid(pr, 'spacing');
      if (sp){
        if (attr(sp, 'before') != null) setCss('margin-top', (+attr(sp, 'before') / 20) + 'pt');
        if (attr(sp, 'after') != null) setCss('margin-bottom', (+attr(sp, 'after') / 20) + 'pt');
        const line = attr(sp, 'line'), rule = attr(sp, 'lineRule') || 'auto';
        if (line && rule === 'auto'){ const m = Math.round(+line / 240 * 100) / 100; lsVal = m; setCss('line-height', (m * 1.22).toFixed(3)); }
        else if (line && (rule === 'exact' || rule === 'atLeast')) setCss('line-height', (+line / 20) + 'pt');
      }
      const shd = kid(pr, 'shd');
      if (shd && attr(shd, 'fill') && !/^(auto|FFFFFF)$/i.test(attr(shd, 'fill'))) setCss('background-color', '#' + attr(shd, 'fill'));
      const bdr = kid(pr, 'pBdr');
      if (bdr){
        const sides = ['top','bottom','left','right'].filter(sn => { const e = kid(bdr, sn); return e && !['nil','none'].includes(attr(e, 'val')); });
        if (sides.length === 4) cls = (cls + ' sw-box').trim();
        else sides.forEach(sn => { const e = kid(bdr, sn); setCss('border-' + sn, `${Math.max(0.5, (+attr(e, 'sz') || 4) / 8)}pt solid #${attr(e, 'color') && attr(e, 'color') !== 'auto' ? attr(e, 'color') : '000000'}`); if (sn === 'bottom') setCss('padding-bottom', '1pt'); });
      }
      if (kid(pr, 'pageBreakBefore') && onOff(kid(pr, 'pageBreakBefore'))) pageBreakBefore = true;
    };
    const cssMap = {};
    let lsVal = null;
    const setCss = (k, v) => { if (v === '' || v == null) delete cssMap[k]; else cssMap[k] = v; };

    const wasSkipping = fld.skipToc;
    if (pp){
      const ps = kid(pp, 'pStyle');
      const sid = ps ? attr(ps, 'val') : null;
      const sname = sid && styles.map[sid] ? styles.map[sid].name.toLowerCase() : '';
      if (sname === 'toc heading' || sname === 'inhaltsverzeichnisüberschrift'){ await runsHtml(p, baseRs, {}); return {skip:true}; }
      if (sid){
        const cl = classifyStyle(styles, sid);
        if (cl){ tag = cl.tag; if (cl.cls) cls = cl.cls; }
        else {
          // eigene Vorlage: Formatierung übernehmen
          const chain = styleChain(styles, sid);
          chain.forEach(s => { applyPPr(s.pPr); baseRs = mergeStyle(runStyle(s.rPr), baseRs); });
          const ln = (styles.map[sid] && styles.map[sid].name || '').toLowerCase();
          if (/list paragraph|listenabsatz/.test(ln)) { delete cssMap['margin-left']; }
        }
      }
      applyPPr(pp);
      let np = kid(pp, 'numPr');
      if (!np && sid){
        // Nummerierung aus der Formatvorlage (z. B. „Aufzählungszeichen“)
        let st = styles.map[sid], g = 0;
        while (st && g++ < 10 && !np){ np = st.pPr && kid(st.pPr, 'numPr'); st = st.basedOn ? styles.map[st.basedOn] : null; }
      }
      if (np){
        const ni = kid(np, 'numId'), il = kid(np, 'ilvl');
        const id = ni ? attr(ni, 'val') : null;
        if (id && id !== '0' && numbering[id]) num = {id, lvl:il ? +attr(il, 'val') : 0};
      }
      if (num){ delete cssMap['margin-left']; delete cssMap['text-indent']; }
    }
    const out = {};
    let inner = await runsHtml(p, baseRs, out);
    if (lsVal) attrs += ` data-ls="${lsVal}"`;
    const style = Object.entries(cssMap).map(([k, v]) => `${k}:${v}`).join(';');
    if (style) attrs += ` style="${esc(style)}"`;
    if (cls) attrs += ` class="${cls}"`;
    if ((fld.skipToc || wasSkipping) && !out.toc && !inner.replace(/<[^>]+>/g, '').trim()) return {skip:true};
    const parts = inner.split('\u0000PB\u0000');
    const html = parts.map((pt, i) => {
      const content = pt.trim() ? pt : '<br>';
      return (i > 0 ? '<div class="sw-pagebreak" contenteditable="false"></div>' : '') + (i > 0 && !pt.trim() && i === parts.length - 1 ? '' : `<${tag}${attrs}>${content}</${tag}>`);
    }).join('');
    let html2 = html;
    if (!num && inner.includes('sw-math') && (out.mathPara || !inner.replace(/<span class="sw-math"[^>]*>.*?<\/span>/g, '').replace(/<[^>]+>/g, '').replace(/&nbsp;|\s/g, ''))){
      html2 = html2.replace(/^<p( class="([^"]*)")?/, (m0, a, cls) => `<p class="${((cls || '') + ' sw-mathblock').trim()}"`);
    }
    // Abschnittsende mit Spalten?
    const sp = pp && kid(pp, 'sectPr');
    let sectCols = null;
    if (sp){ const co = kid(sp, 'cols'); sectCols = {n:co ? +(attr(co, 'num') || 1) : 1, line:co ? attr(co, 'sep') === '1' || attr(co, 'sep') === 'true' : false}; }
    return {html:(pageBreakBefore ? '<div class="sw-pagebreak" contenteditable="false"></div>' : '') + html2, num, inner, tag, attrs, toc:out.toc, tof:out.tof, sectCols};
  }

  async function cellHtml(tc){ return await blocksHtml(tc); }

  async function tableHtml(tbl){
    const grid = kid(tbl, 'tblGrid');
    const cols = kids(grid, 'gridCol').map(g => +attr(g, 'w') || 0);
    const total = cols.reduce((a, b) => a + b, 0) || 1;
    const tp = kid(tbl, 'tblPr');
    let noBorder = false;
    const ts = tp && kid(tp, 'tblStyle');
    const tb = tp && kid(tp, 'tblBorders');
    if (tb){
      const vals = kids(tb).map(e => attr(e, 'val'));
      noBorder = vals.length && vals.every(v => v === 'nil' || v === 'none');
    } else if (!ts || /^(TableNormal|NormaleTabelle)$/i.test(attr(ts, 'val'))) noBorder = true;
    if (ts && styles.map[attr(ts, 'val')]){
      const st = styles.map[attr(ts, 'val')];
      if (/grid|raster|table/i.test(st.name) && !tb) noBorder = /normal/i.test(st.name);
    }
    const contentCm = (SW.PAGE_SIZES[doc.pageSize][doc.orientation === 'landscape' ? 1 : 0]) - doc.margins.left - doc.margins.right;
    const tableCm = total * TW_CM;
    const widthPct = Math.min(100, Math.round(tableCm / contentCm * 1000) / 10);
    let h = `<table${noBorder ? ' class="sw-noborder"' : ''}${widthPct < 98 ? ` style="width:${widthPct}%"` : ''}><colgroup>`;
    cols.forEach(c => h += `<col style="width:${(c / total * 100).toFixed(2)}%">`);
    h += '</colgroup><tbody>';
    // vMerge auflösen
    const rows = kids(tbl, 'tr');
    const model = [];
    rows.forEach((tr, ri) => {
      const cells = [];
      let ci = 0;
      for (const tc of kids(tr).flatMap(n => n.localName === 'sdt' ? kids(kid(n, 'sdtContent'), 'tc') : (n.localName === 'tc' ? [n] : []))){
        const pr = kid(tc, 'tcPr');
        const gs = pr && kid(pr, 'gridSpan') ? +attr(kid(pr, 'gridSpan'), 'val') : 1;
        const vm = pr && kid(pr, 'vMerge');
        const vmv = vm ? (attr(vm, 'val') || 'continue') : null;
        cells.push({tc, col:ci, span:gs, vm:vmv, rowspan:1});
        ci += gs;
      }
      model.push(cells);
    });
    model.forEach((cells, ri) => cells.forEach(c => {
      if (c.vm === 'continue'){
        for (let r = ri - 1; r >= 0; r--){
          const up = model[r].find(x => x.col === c.col);
          if (up && up.vm !== 'continue'){ up.rowspan++; break; }
          if (!up) break;
        }
      }
    }));
    for (let ri = 0; ri < rows.length; ri++){
      h += '<tr>';
      for (const c of model[ri]){
        if (c.vm === 'continue') continue;
        const pr = kid(c.tc, 'tcPr');
        const css = [];
        const sh = pr && kid(pr, 'shd');
        if (sh && attr(sh, 'fill') && !/^(auto|FFFFFF)$/i.test(attr(sh, 'fill'))) css.push('background-color:#' + attr(sh, 'fill'));
        const va = pr && kid(pr, 'vAlign');
        if (va && attr(va, 'val') === 'center') css.push('vertical-align:middle');
        if (va && attr(va, 'val') === 'bottom') css.push('vertical-align:bottom');
        h += `<td${c.span > 1 ? ` colspan="${c.span}"` : ''}${c.rowspan > 1 ? ` rowspan="${c.rowspan}"` : ''}${css.length ? ` style="${css.join(';')}"` : ''}>${await cellHtml(c.tc) || '<p><br></p>'}</td>`;
      }
      h += '</tr>';
    }
    return h + '</tbody></table>';
  }

  async function blocksHtml(container, isBody){
    let html = '';
    let secStart = 0;
    const endSection = cols => {
      if (!isBody || !cols || cols.n <= 1) { secStart = html.length; return; }
      html = html.slice(0, secStart) + `<div class="sw-cols" data-n="${Math.min(4, cols.n)}"${cols.line ? ' data-line="1"' : ''}>` + html.slice(secStart) + '</div>';
      secStart = html.length;
    };
    // Listen-Stapel
    let stack = [];   // {numId, lvl, tag}
    const closeAll = () => { let s = ''; while (stack.length) s += `</li></${stack.pop().tag}>`; return s; };
    const items = [];
    const collect = el => { for (const n of kids(el)){ if (n.localName === 'sdt'){ const c = kid(n, 'sdtContent'); if (c) collect(c); } else if (n.localName === 'customXml' || n.localName === 'ins') collect(n); else items.push(n); } };
    collect(container);
    for (const n of items){
      const ln = n.localName;
      if (ln === 'p'){
        const r = await paraHtml(n);
        if (r.skip){ if (r.sectCols){ html += closeAll(); endSection(r.sectCols); } continue; }
        if (r.tof){ html += closeAll() + `<div class="sw-tof" data-seq="${esc(r.tof)}" contenteditable="false"></div>`; continue; }
        if (r.toc){ html += closeAll() + '<div class="sw-toc" contenteditable="false"></div>'; continue; }
        if (fld.skipToc) continue;
        if (r.sectCols){
          // Absatz gehört noch zum endenden Abschnitt
          if (r.num){ /* selten – als normaler Absatz behandeln */ }
          html += closeAll() + (r.inner.replace(/<[^>]+>/g, '').trim() || r.inner.includes('<img') ? r.html : '');
          endSection(r.sectCols);
          continue;
        }
        if (r.num){
          const levels = numbering[r.num.id];
          const lv = levels[r.num.lvl] || levels['0'];
          const info = listInfo(lv);
          const lvl = r.num.lvl;
          const la = r.attrs.replace(/ class="[^"]*"/, '');
          while (stack.length && stack[stack.length - 1].lvl > lvl) html += `</li></${stack.pop().tag}>`;
          let top = stack[stack.length - 1];
          if (top && top.lvl === lvl){
            const same = top.numId === r.num.id || (top.tag === info.tag && top.attrs === info.attrs && (info.tag === 'ul' || (top.abs === levels._abs && !levels._override)));
            if (same){ html += `</li><li${la}>${r.inner || '<br>'}`; continue; }
            html += `</li></${stack.pop().tag}>`;
          }
          html += `<${info.tag}${info.attrs}><li${la}>${r.inner || '<br>'}`;
          stack.push({numId:r.num.id, lvl, tag:info.tag, attrs:info.attrs, abs:levels._abs});
          continue;
        }
        html += closeAll() + r.html;
        continue;
      }
      if (ln === 'tbl'){ html += closeAll() + await tableHtml(n); continue; }
    }
    html += closeAll();
    if (isBody && sect){ const co = kid(sect, 'cols'); if (co && +(attr(co, 'num') || 1) > 1) endSection({n:+attr(co, 'num'), line:attr(co, 'sep') === '1'}); }
    return html;
  }

  let html = await blocksHtml(body, true);
  Object.keys(usedComments).forEach(id => { doc.comments['w' + id] = comments[id]; });
  // Fußnoten-Markierung korrekt nummerieren erfolgt im Layout
  if (!html.trim()) html = '<p><br></p>';
  return {html, doc, images, exact:false};
};

/* Text- und HTML-Dateien */
SW.importText = function(text){
  const lines = text.replace(/\r\n?/g, '\n').split('\n');
  return lines.map(l => `<p>${l ? esc(l).replace(/\t/g, '<span class="sw-tab">\t</span>') : '<br>'}</p>`).join('');
};
})();
