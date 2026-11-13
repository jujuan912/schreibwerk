/* ==========================================================
   Schreibwerk – Kern: Zustand, Auswahl, Rückgängig, Befehle
   ========================================================== */
(function(){
'use strict';

const ed = document.getElementById('editor');
const CM = 96/2.54;            // Pixel pro Zentimeter
const PT = 96/72;              // Pixel pro Punkt

const PAGE_SIZES = { A4:[21,29.7], A5:[14.8,21], A3:[29.7,42], Letter:[21.59,27.94] };
const MARGIN_PRESETS = {
  normal:{top:2.5,bottom:2,left:2.5,right:2.5},
  schmal:{top:1.27,bottom:1.27,left:1.27,right:1.27},
  mittel:{top:2.54,bottom:2.54,left:1.91,right:1.91},
  breit:{top:2.54,bottom:2.54,left:5.08,right:5.08},
  din:{top:2.7,bottom:2,left:2.5,right:2}
};
const FONT_SIZES = [8,9,10,10.5,11,12,14,16,18,20,22,24,26,28,36,48,72];
const BLOCK_SEL = 'p,h1,h2,h3,h4,h5,h6,li,blockquote,pre';

function defaultDoc(){
  return {
    title:'Dokument1', fileName:null,
    pageSize:'A4', orientation:'portrait',
    margins:{...MARGIN_PRESETS.normal},
    header:{text:'',align:'left'},
    footer:{text:'',align:'center'},
    base:{font:'Calibri', size:11, after:8, line:1.08},
    titlePage:false,                              // erste Seite anders
    firstHeader:{text:'',align:'left'},
    firstFooter:{text:'',align:'center'},
    pageNum:{start:1, fmt:'decimal'},
    abbr:{},                                      // Abkürzung → Bedeutung
    comments:{},                                  // id → {text, author, date, done}
    author:'', created:new Date().toISOString()
  };
}

const SW = window.SW = {
  ed, CM, PT, PAGE_SIZES, MARGIN_PRESETS, FONT_SIZES, BLOCK_SEL,
  doc: defaultDoc(),
  images: new Map(),          // id -> {blob, url, type}
  zoom: 1,
  dirty: false,
  fileHandle: null,
  lastRange: null,
  defaultDoc
};

/* ---------------- Hilfsfunktionen ---------------- */
SW.$ = (s, r=document) => r.querySelector(s);
SW.$$ = (s, r=document) => Array.from(r.querySelectorAll(s));
SW.esc = s => String(s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
SW.uid = () => Math.random().toString(36).slice(2, 10);
SW.fmtNum = (n, d=2) => Number(n).toLocaleString('de-DE', {maximumFractionDigits:d});
SW.parseNum = s => parseFloat(String(s).replace(',', '.'));

SW.toast = function(msg, action, ms=3200){
  const t = document.getElementById('toast');
  t.innerHTML = '';
  t.append(document.createTextNode(msg));
  if (action){
    const b = document.createElement('button');
    b.textContent = action.label;
    b.style.cssText = 'margin-left:14px;color:#9cc4ff;font-weight:600';
    b.onclick = () => { t.hidden = true; action.run(); };
    t.append(b);
  }
  t.hidden = false;
  clearTimeout(SW._toastT);
  SW._toastT = setTimeout(() => t.hidden = true, action ? Math.max(ms, 9000) : ms);
};

SW.rgbToHex = function(c){
  if (!c) return null;
  c = c.trim();
  if (c[0] === '#'){
    if (c.length === 4) c = '#' + c[1]+c[1]+c[2]+c[2]+c[3]+c[3];
    return c.slice(1,7).toUpperCase();
  }
  const m = c.match(/rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+))?/);
  if (!m) {
    const named = {black:'000000',white:'FFFFFF',red:'FF0000',blue:'0000FF',green:'008000',yellow:'FFFF00',gray:'808080',grey:'808080',navy:'000080',maroon:'800000',purple:'800080',teal:'008080',orange:'FFA500'};
    return named[c.toLowerCase()] || null;
  }
  if (m[4] !== undefined && parseFloat(m[4]) === 0) return null;
  return [m[1],m[2],m[3]].map(x => (+x).toString(16).padStart(2,'0')).join('').toUpperCase();
};

/* ---------------- Auswahl ---------------- */
SW.inEditor = function(node){
  return !!node && (node === ed || ed.contains(node));
};
SW.getRange = function(){
  const s = getSelection();
  if (s.rangeCount && SW.inEditor(s.getRangeAt(0).commonAncestorContainer)) return s.getRangeAt(0);
  return null;
};
SW.rememberRange = function(){
  const r = SW.getRange();
  if (r) SW.lastRange = r.cloneRange();
};
SW.focusEditor = function(){
  if (document.activeElement !== ed){
    ed.focus({preventScroll:true});
    if (SW.lastRange && SW.inEditor(SW.lastRange.startContainer)){
      const s = getSelection(); s.removeAllRanges(); s.addRange(SW.lastRange);
    }
  }
  if (!SW.getRange()){
    const r = document.createRange();
    const first = ed.querySelector(BLOCK_SEL) || ed;
    r.setStart(first, 0); r.collapse(true);
    const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  }
};

function innermostBlock(node){
  let n = node.nodeType === 1 ? node : node.parentNode;
  while (n && n !== ed){
    if (n.nodeType === 1 && (n.matches(BLOCK_SEL) || n.matches('td,th,div.sw-box'))) return n;
    n = n.parentNode;
  }
  return null;
}
SW.innermostBlock = innermostBlock;

function allBlocks(){
  return SW.$$(BLOCK_SEL + ',td,th', ed).filter(b => !b.closest('.sw-toc'));
}

/* Position als (Block-Index, Zeichen-Offset) – robust gegen Neuaufbau des DOM */
function pointToPos(node, off, blocks){
  const b = innermostBlock(node);
  if (!b){
    const r = document.createRange(); r.setStart(ed, 0);
    try { r.setEnd(node, off); } catch(e){ return {b:-1, o:0}; }
    return {b:-1, o:r.toString().length};
  }
  const r = document.createRange();
  r.selectNodeContents(b);
  try { r.setEnd(node, off); } catch(e){ return {b:blocks.indexOf(b), o:0}; }
  return {b:blocks.indexOf(b), o:r.toString().length};
}
function posToPoint(pos, blocks){
  let root = pos.b >= 0 ? blocks[pos.b] : ed;
  if (!root) root = blocks[blocks.length-1] || ed;
  let rest = pos.o;
  const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let t, last = null;
  while ((t = w.nextNode())){
    if (t.parentNode.closest && t.parentNode.closest('.sw-toc,.sw-fn')) { continue; }
    if (rest <= t.length) return [t, rest];
    rest -= t.length; last = t;
  }
  if (last) return [last, last.length];
  return [root, 0];
}
SW.saveSel = function(){
  const r = SW.getRange();
  if (!r) return null;
  const blocks = allBlocks();
  return {
    s: pointToPos(r.startContainer, r.startOffset, blocks),
    e: r.collapsed ? null : pointToPos(r.endContainer, r.endOffset, blocks)
  };
};
SW.restoreSel = function(saved, focus=true){
  if (!saved) return;
  const blocks = allBlocks();
  const r = document.createRange();
  try {
    const [sn, so] = posToPoint(saved.s, blocks);
    r.setStart(sn, so);
    if (saved.e){ const [en, eo] = posToPoint(saved.e, blocks); r.setEnd(en, eo); }
    else r.collapse(true);
  } catch(e){ return; }
  const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  SW.lastRange = r.cloneRange();
  if (focus && document.activeElement !== ed) ed.focus({preventScroll:true});
};

/* Ausgewählte Absätze (innerste Blöcke) */
SW.selectedBlocks = function(readOnly){
  const r = SW.getRange() || SW.lastRange;
  if (!r || !SW.inEditor(r.startContainer)) return [];
  if (r.collapsed){
    let b = innermostBlock(r.startContainer);
    if (!readOnly && (!b || b.matches('td,th'))){
      b = SW.wrapLooseInline(r.startContainer) || b;
    }
    return b ? [b] : [];
  }
  const list = SW.$$(BLOCK_SEL, ed).filter(b => r.intersectsNode(b) && !b.closest('.sw-toc'));
  const inner = list.filter(b => !list.some(o => o !== b && b.contains(o)));
  if (!inner.length){
    const b = innermostBlock(r.startContainer);
    return b ? [b] : [];
  }
  return inner;
};

/* Lose Inline-Inhalte (z.B. direkt im Editor oder in Zellen) in <p> packen */
SW.wrapLooseInline = function(node){
  let container = node.nodeType === 1 ? node : node.parentNode;
  // finde nächsten Container (Editor oder Zelle)
  while (container && container !== ed && !container.matches('td,th')) {
    if (container.matches && container.matches(BLOCK_SEL)) return container;
    container = container.parentNode;
  }
  if (!container) return null;
  // Lauf der Inline-Knoten um "node" finden
  let child = node;
  while (child && child.parentNode !== container) child = child.parentNode;
  if (!child) return null;
  const isInline = n => n.nodeType === 3 || (n.nodeType === 1 && !n.matches(BLOCK_SEL + ',table,ul,ol,hr,div,.sw-pagebreak'));
  if (!isInline(child)) return null;
  let start = child, end = child;
  while (start.previousSibling && isInline(start.previousSibling)) start = start.previousSibling;
  while (end.nextSibling && isInline(end.nextSibling)) end = end.nextSibling;
  const sel = SW.saveSel();
  const p = document.createElement('p');
  container.insertBefore(p, start);
  let n = start;
  while (n){ const next = n === end ? null : n.nextSibling; p.appendChild(n); n = next; }
  if (!p.firstChild) p.innerHTML = '<br>';
  SW.restoreSel(sel);
  return p;
};

/* Bei leerer Auswahl im Wort: Wort markieren (wie Word) */
SW.expandToWord = function(){
  const s = getSelection();
  const r = SW.getRange();
  if (!r || !r.collapsed) return false;
  const n = r.startContainer;
  if (n.nodeType !== 3) return false;
  const t = n.data, o = r.startOffset;
  const isW = c => c && /[\p{L}\p{N}_'’-]/u.test(c);
  if (!isW(t[o-1]) || !isW(t[o])) return false;
  let a = o, b = o;
  while (a > 0 && isW(t[a-1])) a--;
  while (b < t.length && isW(t[b])) b++;
  const nr = document.createRange();
  nr.setStart(n, a); nr.setEnd(n, b);
  s.removeAllRanges(); s.addRange(nr);
  SW._wordExpanded = {node:n, caret:o};
  return true;
};
SW.collapseBackFromWord = function(){
  const w = SW._wordExpanded;
  SW._wordExpanded = null;
  if (!w) return;
  const r = SW.getRange();
  if (!r) return;
  // Caret ans Ende der Auswahl zurücksetzen
  const s = getSelection();
  r.collapse(false);
  s.removeAllRanges(); s.addRange(r);
};

/* ---------------- Rückgängig / Wiederholen ---------------- */
const H = {stack:[], idx:-1, timer:null, max:200, restoring:false};
SW.cleanClone = function(){
  const c = ed.cloneNode(true);
  c.querySelectorAll('.sw-gap').forEach(g => g.remove());
  c.querySelectorAll('tr.sw-gaprow').forEach(g => g.remove());
  c.querySelectorAll('[data-swp]').forEach(el => el.removeAttribute('data-swp'));
  c.querySelectorAll('[data-sw-mt]').forEach(el => el.removeAttribute('data-sw-mt'));
  c.querySelectorAll('[data-swh]').forEach(el => el.removeAttribute('data-swh'));
  c.querySelectorAll('.sw-cols-cont').forEach(cont => {
    const prev = cont.previousElementSibling;
    if (prev && prev.classList.contains('sw-cols')){ while (cont.firstChild) prev.appendChild(cont.firstChild); cont.remove(); }
    else cont.classList.remove('sw-cols-cont');
  });
  c.querySelectorAll('[style=""]').forEach(el => el.removeAttribute('style'));
  c.querySelectorAll('.sw-imgsel').forEach(el => el.classList.remove('sw-imgsel'));
  c.querySelectorAll('.sw-cellsel').forEach(el => el.classList.remove('sw-cellsel'));
  c.querySelectorAll('[class=""]').forEach(el => el.removeAttribute('class'));
  c.normalize();
  return c;
};
SW.cleanHTML = () => SW.cleanClone().innerHTML;

const STATE_KEYS = ['pageSize','orientation','margins','header','footer','base','titlePage','firstHeader','firstFooter','pageNum','abbr','comments'];
function docState(){ const o = {}; STATE_KEYS.forEach(k => o[k] = SW.doc[k]); return JSON.stringify(o); }

SW.commit = function(){
  clearTimeout(H.timer); H.timer = null;
  if (H.restoring) return;
  const snap = {html:SW.cleanHTML(), sel:SW.saveSel(), doc:docState()};
  const cur = H.stack[H.idx];
  if (cur && cur.html === snap.html && cur.doc === snap.doc){ cur.sel = snap.sel || cur.sel; return; }
  H.stack = H.stack.slice(0, H.idx + 1);
  H.stack.push(snap);
  if (H.stack.length > H.max) H.stack.shift();
  H.idx = H.stack.length - 1;
  SW.updateUndoButtons && SW.updateUndoButtons();
};
SW.scheduleCommit = function(){
  clearTimeout(H.timer);
  H.timer = setTimeout(SW.commit, 450);
};
SW.flushCommit = function(){ if (H.timer) SW.commit(); };
SW.resetHistory = function(){ H.stack = []; H.idx = -1; SW.commit(); };
SW.canUndo = () => H.idx > 0 || !!H.timer;
SW.canRedo = () => H.idx < H.stack.length - 1;

function restoreSnap(s){
  H.restoring = true;
  ed.innerHTML = s.html;
  const d = JSON.parse(s.doc);
  Object.assign(SW.doc, d);
  SW.fixImages();
  SW.applyBase && SW.applyBase();
  SW.relayout && SW.relayout(true);
  SW.restoreSel(s.sel);
  H.restoring = false;
  SW.markDirty();
  SW.updateUndoButtons && SW.updateUndoButtons();
  SW.refreshUI && SW.refreshUI();
}
SW.undo = function(){
  SW.flushCommit();
  if (H.idx <= 0) return;
  H.idx--; restoreSnap(H.stack[H.idx]);
};
SW.redo = function(){
  if (H.idx >= H.stack.length - 1) return;
  H.idx++; restoreSnap(H.stack[H.idx]);
};

/* Befehlsklammer: Zustand davor sichern, danach festschreiben */
SW.edit = function(fn){
  SW.flushCommit();
  SW.focusEditor();
  const res = fn();
  SW.normalizeDoc();
  SW.commit();
  SW.markDirty();
  SW.requestLayout && SW.requestLayout();
  SW.refreshUI && SW.refreshUI();
  return res;
};

SW.markDirty = function(){
  if (H.restoring) return;
  SW.dirty = true;
  SW.updateTitle && SW.updateTitle();
  SW.scheduleAutosave && SW.scheduleAutosave();
};

/* ---------------- Grundlegende Formatierung ---------------- */
SW.native = function(cmd, val=null){
  document.execCommand('styleWithCSS', false, true);
  return document.execCommand(cmd, false, val);
};

/* Wendet eine Inline-Formatierung an: bei Caret im Wort auf das Wort */
SW.inline = function(fn){
  return SW.edit(() => {
    const exp = SW.expandToWord();
    fn();
    if (exp) SW.collapseBackFromWord();
  });
};

/* Schriftgrad in pt setzen */
SW.setFontSize = function(pt){
  pt = Math.max(1, Math.min(1638, Math.round(pt * 2) / 2));
  SW.inline(() => SW.applySpanStyle('fontSize', pt + 'pt'));
};

/* Generische Span-Formatierung über Marker */
SW.applySpanStyle = function(prop, value){
  const r = SW.getRange();
  if (!r) return;
  if (r.collapsed){
    // Leerer Span mit Nullbreitenzeichen für die nächste Eingabe
    const span = document.createElement('span');
    span.style[prop] = value;
    span.textContent = '​';
    r.insertNode(span);
    const nr = document.createRange();
    nr.setStart(span.firstChild, 1); nr.collapse(true);
    const s = getSelection(); s.removeAllRanges(); s.addRange(nr);
    return;
  }
  document.execCommand('styleWithCSS', false, false);
  document.execCommand('fontSize', false, '7');
  const marks = SW.$$('font[size="7"]', ed);
  document.execCommand('styleWithCSS', false, true);
  const s = getSelection();
  let first = null, last = null;
  marks.forEach(f => {
    const span = document.createElement('span');
    span.style[prop] = value;
    while (f.firstChild) span.appendChild(f.firstChild);
    f.replaceWith(span);
    // innere Formatierung derselben Eigenschaft entfernen
    span.querySelectorAll('[style]').forEach(x => { x.style[prop] = ''; if (!x.getAttribute('style')) x.removeAttribute('style'); });
    if (!first) first = span;
    last = span;
  });
  if (first){
    const nr = document.createRange();
    nr.setStartBefore(first); nr.setEndAfter(last);
    s.removeAllRanges(); s.addRange(nr);
  }
};

/* Aktuelle Formatierung an der Auswahl ermitteln */
SW.currentFormat = function(){
  const r = SW.getRange() || SW.lastRange;
  let n = r ? r.startContainer : ed;
  if (r && !r.collapsed && n.nodeType === 1 && n.childNodes[r.startOffset]) n = n.childNodes[r.startOffset];
  if (n.nodeType === 3) n = n.parentNode;
  if (!SW.inEditor(n)) n = ed.querySelector(BLOCK_SEL) || ed;
  const cs = getComputedStyle(n);
  const block = innermostBlock(n);
  const fam = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
  return {
    font: fam,
    size: Math.round(parseFloat(cs.fontSize) / PT * 2) / 2,
    align: block ? getComputedStyle(block).textAlign : 'left',
    block
  };
};

/* ---------------- Absatzformatierung ---------------- */
SW.setAlign = function(a){
  SW.edit(() => {
    let blocks = SW.selectedBlocks();
    const img = SW.selectedImage && SW.selectedImage();
    if (img) blocks = [innermostBlock(img) || SW.wrapLooseInline(img)].filter(Boolean);
    const all = blocks.every(b => getComputedStyle(b).textAlign === a || (a==='left' && getComputedStyle(b).textAlign==='start'));
    blocks.forEach(b => { b.style.textAlign = (all && a !== 'left') || a === 'left' ? '' : a; });
    blocks.forEach(b => { if (!b.getAttribute('style')) b.removeAttribute('style'); });
  });
};

SW.setLineSpacing = function(m){
  SW.edit(() => SW.selectedBlocks().forEach(b => {
    b.dataset.ls = m;
    b.style.lineHeight = (m * 1.22).toFixed(3);
  }));
};

SW.setBlockProp = function(fn){
  SW.edit(() => SW.selectedBlocks().forEach(b => { fn(b); if (!b.getAttribute('style')) b.removeAttribute('style'); }));
};

SW.indent = function(dir){
  const blocks = SW.selectedBlocks();
  const inList = blocks.length && blocks.every(b => b.closest('li'));
  if (inList){
    SW.edit(() => {
      SW.native(dir > 0 ? 'indent' : 'outdent');
      SW.fixLists();
    });
    return;
  }
  SW.setBlockProp(b => {
    const cur = parseFloat(b.style.marginLeft) || 0;     // in cm (gespeichert)
    const curCm = b.style.marginLeft.endsWith('cm') ? cur : cur / CM;
    let next = Math.round((curCm + dir * 1.25) / 1.25) * 1.25;
    if (next <= 0) b.style.marginLeft = '';
    else b.style.marginLeft = next.toFixed(2) + 'cm';
  });
};

/* Formatvorlagen */
SW.STYLES = [
  {id:'normal', name:'Standard', tag:'p', cls:''},
  {id:'nospace', name:'Kein Leerraum', tag:'p', cls:'sw-nospace'},
  {id:'h1', name:'Überschrift 1', tag:'h1', cls:''},
  {id:'h2', name:'Überschrift 2', tag:'h2', cls:''},
  {id:'h3', name:'Überschrift 3', tag:'h3', cls:''},
  {id:'title', name:'Titel', tag:'p', cls:'sw-title'},
  {id:'subtitle', name:'Untertitel', tag:'p', cls:'sw-subtitle'},
  {id:'quote', name:'Zitat', tag:'p', cls:'sw-quote'}
];
const STYLE_CLASSES = ['sw-nospace','sw-title','sw-subtitle','sw-quote'];
SW.styleOf = function(b){
  if (!b) return 'normal';
  const t = b.tagName.toLowerCase();
  if (/^h[1-3]$/.test(t)) return t;
  if (/^h[4-6]$/.test(t)) return 'h3';
  for (const s of SW.STYLES) if (s.cls && b.classList.contains(s.cls)) return s.id;
  return 'normal';
};
SW.retag = function(el, tag){
  if (el.tagName.toLowerCase() === tag) return el;
  const n = document.createElement(tag);
  for (const a of el.attributes) n.setAttribute(a.name, a.value);
  while (el.firstChild) n.appendChild(el.firstChild);
  el.replaceWith(n);
  return n;
};
SW.applyStyle = function(id){
  const st = SW.STYLES.find(s => s.id === id);
  if (!st) return;
  SW.edit(() => {
    const sel = SW.saveSel();
    SW.selectedBlocks().forEach(b => {
      if (b.matches('td,th')) b = SW.wrapLooseInline(b.firstChild || b) || b;
      if (b.matches('li')) {
        // Formatvorlage im Listeneintrag: Inhalt in Absatz umwandeln
        if (st.tag === 'p' && !st.cls){ STYLE_CLASSES.forEach(c => b.classList.remove(c)); return; }
        return;
      }
      if (b.matches('td,th')) return;
      const n = SW.retag(b, st.tag);
      STYLE_CLASSES.forEach(c => n.classList.remove(c));
      if (st.cls) n.classList.add(st.cls);
      if (!n.getAttribute('class')) n.removeAttribute('class');
    });
    SW.restoreSel(sel);
  });
};

/* ---------------- Listen ---------------- */
SW.closestList = function(node){
  const n = node && (node.nodeType === 1 ? node : node.parentNode);
  return n && n.closest ? n.closest('ul,ol') : null;
};
SW.listsInSelection = function(){
  const r = SW.getRange();
  if (!r) return [];
  const set = new Set();
  const l = SW.closestList(r.startContainer); if (l) set.add(l);
  SW.$$('ul,ol', ed).forEach(x => { if (r.intersectsNode(x) && !r.collapsed) set.add(x); });
  return [...set];
};
const BULLET_STYLES = {disc:'', circle:'circle', square:'square', dash:'', arrow:'', check:''};
const NUM_STYLES = {'decimal':'', 'decimal-paren':'', 'lower-alpha':'', 'upper-alpha':'upper-alpha', 'lower-roman':'lower-roman', 'upper-roman':'upper-roman'};

SW.toggleList = function(kind, fmt){
  SW.edit(() => {
    const tag = kind === 'ul' ? 'UL' : 'OL';
    const before = SW.closestList(SW.getRange().startContainer);
    if (!(fmt && before && before.tagName === tag)){
      SW.native(kind === 'ul' ? 'insertUnorderedList' : 'insertOrderedList');
    }
    const r = SW.getRange();
    const lists = SW.listsInSelection().filter(l => l.tagName === tag);
    if (r) { const l = SW.closestList(r.startContainer); if (l && l.tagName === tag && !lists.includes(l)) lists.push(l); }
    lists.forEach(l => {
      if (kind === 'ul'){
        const f = fmt || l.dataset.bullet || 'disc';
        l.dataset.bullet = f;
        l.style.listStyleType = BULLET_STYLES[f] || '';
      } else {
        const f = fmt || l.dataset.fmt || 'decimal';
        l.dataset.fmt = f;
        l.style.listStyleType = NUM_STYLES[f] || '';
      }
      if (!l.getAttribute('style')) l.removeAttribute('style');
    });
    SW.fixLists();
  });
};
SW.restartNumbering = function(){
  SW.edit(() => {
    const r = SW.getRange(); if (!r) return;
    const li = (r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentNode).closest('li');
    if (!li) return;
    const ol = li.parentNode;
    if (ol.tagName !== 'OL' || ol.firstElementChild === li) return;
    const n = ol.cloneNode(false);
    let x = li;
    while (x){ const nx = x.nextSibling; n.appendChild(x); x = nx; }
    ol.after(n);
    SW.restoreSel(SW.saveSel());
  });
};
/* Listen bereinigen: leere Listen entfernen, verschachtelte Listen erben Formate */
SW.fixLists = function(){
  SW.$$('ul,ol', ed).forEach(l => {
    if (!l.querySelector('li')) { l.remove(); return; }
    const parentList = l.parentNode.closest && l.parentNode.closest('ul,ol');
    if (parentList && parentList.tagName === l.tagName){
      if (l.tagName === 'UL' && !l.dataset.bullet){
        const lv = SW.$$('ul', ed).filter(x => x.contains(l) && x !== l).length;
        l.dataset.bullet = ['disc','circle','square'][lv % 3];
        l.style.listStyleType = BULLET_STYLES[l.dataset.bullet] || '';
      }
      if (l.tagName === 'OL' && !l.dataset.fmt){
        const lv = SW.$$('ol', ed).filter(x => x.contains(l) && x !== l).length;
        l.dataset.fmt = ['decimal','lower-alpha','lower-roman'][lv % 3];
        l.style.listStyleType = NUM_STYLES[l.dataset.fmt] || '';
      }
      if (!l.getAttribute('style')) l.removeAttribute('style');
    }
  });
};

/* ---------------- Groß-/Kleinschreibung ---------------- */
SW.changeCase = function(mode){
  SW.edit(() => {
    SW.expandToWord();
    const r = SW.getRange();
    if (!r || r.collapsed) return;
    const nodes = SW.textNodesInRange(r);
    const full = nodes.map(x => x.node.data.slice(x.start, x.end)).join('');
    let out;
    if (mode === 'lower') out = full.toLocaleLowerCase('de-DE');
    else if (mode === 'upper') out = full.toLocaleUpperCase('de-DE').replace(/ẞ/g, 'SS');
    else if (mode === 'title') out = full.toLocaleLowerCase('de-DE').replace(/(^|[\s\-–(„"])(\p{L})/gu, (m, a, b) => a + b.toLocaleUpperCase('de-DE'));
    else if (mode === 'toggle') out = [...full].map(c => c === c.toLocaleUpperCase('de-DE') ? c.toLocaleLowerCase('de-DE') : c.toLocaleUpperCase('de-DE')).join('');
    else out = full.toLocaleLowerCase('de-DE').replace(/(^\s*|[.!?]\s+)(\p{L})/gu, (m, a, b) => a + b.toLocaleUpperCase('de-DE'));
    // Längen können sich ändern (ß → SS) – Zeichen verteilen
    let pos = 0;
    nodes.forEach((x, i) => {
      const len = x.end - x.start;
      const part = i === nodes.length - 1 ? out.slice(pos) : out.slice(pos, pos + len);
      pos += len;
      x.node.data = x.node.data.slice(0, x.start) + part + x.node.data.slice(x.end);
      x.end = x.start + part.length;
    });
    const f = nodes[0], l = nodes[nodes.length - 1];
    const nr = document.createRange();
    nr.setStart(f.node, f.start); nr.setEnd(l.node, l.end);
    const s = getSelection(); s.removeAllRanges(); s.addRange(nr);
  });
};
SW.textNodesInRange = function(r){
  const res = [];
  const root = r.commonAncestorContainer.nodeType === 3 ? r.commonAncestorContainer.parentNode : r.commonAncestorContainer;
  const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let t;
  while ((t = w.nextNode())){
    if (!r.intersectsNode(t)) continue;
    if (t.parentNode.closest('.sw-toc,.sw-fn,.sw-gap')) continue;
    const start = t === r.startContainer ? r.startOffset : 0;
    const end = t === r.endContainer ? r.endOffset : t.length;
    if (end > start) res.push({node:t, start, end});
  }
  return res;
};

/* ---------------- Einfügen von Blöcken ---------------- */
/* Fügt ein Block-Element an der Cursorposition ein (teilt den Absatz wie Word) */
SW.insertBlockAtCaret = function(node, opts={}){
  const newP = () => { const p = document.createElement('p'); p.innerHTML = '<br>'; return p; };
  const isEmpty = el => !el.textContent.replace(/\u200B/g, '').trim() && !el.querySelector('img');
  const place = el => {
    let tgt = el;
    if (opts.caretInside && node.matches('table')) tgt = node.querySelector('td,th');
    const p = tgt.matches('table') ? tgt.querySelector('td,th') : tgt;
    const nr = document.createRange(); nr.setStart(p, 0); nr.collapse(true);
    const s = getSelection(); s.removeAllRanges(); s.addRange(nr);
  };
  const r = SW.getRange();
  if (!r){ ed.appendChild(node); const p = newP(); node.after(p); place(p); return node; }
  if (!r.collapsed) r.deleteContents();
  let blk = innermostBlock(r.startContainer);
  if (!blk || blk.matches('td,th')) blk = SW.wrapLooseInline(r.startContainer) || blk;
  let top = blk;
  while (top && top.parentNode !== ed && !(top.parentNode.matches && top.parentNode.matches('td,th'))) top = top.parentNode;
  if (!top){ ed.appendChild(node); const p = newP(); node.after(p); place(p); return node; }
  if (blk === top && blk.matches('p,h1,h2,h3,h4,h5,h6,blockquote,pre')){
    const rr = SW.getRange();
    const after = document.createRange();
    after.setStart(rr.startContainer, rr.startOffset);
    after.setEnd(blk, blk.childNodes.length);
    const frag = after.extractContents();
    const tail = blk.cloneNode(false);
    tail.appendChild(frag);
    if (isEmpty(blk) && !isEmpty(tail)){
      // Cursor am Absatzanfang: Element davor
      while (blk.firstChild) blk.firstChild.remove();
      while (tail.firstChild) blk.appendChild(tail.firstChild);
      blk.before(node); place(blk); return node;
    }
    if (isEmpty(blk)) { blk.innerHTML = '<br>'; blk.before(node); place(blk); return node; }
    if (isEmpty(tail)){
      blk.after(node);
      let nx = node.nextElementSibling;
      if (!nx || opts.newPara !== false){
        const p = newP();
        node.after(p); nx = p;
      }
      place(nx); return node;
    }
    blk.after(node); node.after(tail); place(tail); return node;
  }
  // in Liste/Tabelle o. ä.: hinter dem obersten Element einfügen
  top.after(node);
  const p = newP(); node.after(p); place(p);
  return node;
};

SW.insertHTML = function(html){
  SW.native('insertHTML', html);
};

/* ---------------- Bilder ---------------- */
SW.registerImage = function(blob, id){
  id = id || ('img' + SW.uid());
  const url = URL.createObjectURL(blob);
  SW.images.set(id, {blob, url, type:blob.type || 'image/png'});
  return {id, url};
};
/* Bilder im DOM mit gespeicherten Blobs verbinden */
SW.fixImages = function(root=ed){
  root.querySelectorAll('img[data-img-id]').forEach(img => {
    const rec = SW.images.get(img.dataset.imgId);
    if (rec && img.getAttribute('src') !== rec.url) img.src = rec.url;
  });
};
SW.normalizeImageBlob = async function(file){
  const ok = ['image/png','image/jpeg','image/gif'];
  if (ok.includes(file.type)) return file;
  // In PNG umwandeln (WebP/BMP öffnet Word nicht überall)
  const bmp = await createImageBitmap(file);
  const c = document.createElement('canvas');
  c.width = bmp.width; c.height = bmp.height;
  c.getContext('2d').drawImage(bmp, 0, 0);
  return await new Promise(res => c.toBlob(res, 'image/png'));
};
SW.insertImageFile = async function(file){
  const blob = await SW.normalizeImageBlob(file);
  const {id, url} = SW.registerImage(blob);
  const dims = await new Promise(res => { const i = new Image(); i.onload = () => res([i.naturalWidth, i.naturalHeight]); i.onerror = () => res([300, 200]); i.src = url; });
  const maxW = ed.clientWidth;
  let w = dims[0];
  if (w > maxW) w = maxW;
  SW.edit(() => {
    SW.insertHTML(`<img data-img-id="${id}" src="${url}" style="width:${Math.round(w)}px" alt="">`);
  });
};

/* ---------------- Dokument bereinigen ---------------- */
/* Auswahl über DOM-Umbauten hinweg erhalten (Textknoten bleiben bestehen) */
SW.keepSelection = function(fn){
  const r = SW.getRange();
  const st = r ? {sc:r.startContainer, so:r.startOffset, ec:r.endContainer, eo:r.endOffset} : null;
  fn();
  if (st && st.sc.isConnected && st.ec.isConnected && SW.inEditor(st.sc)){
    try {
      const nr = document.createRange(); nr.setStart(st.sc, st.so); nr.setEnd(st.ec, st.eo);
      const s = getSelection(); s.removeAllRanges(); s.addRange(nr);
    } catch(e){}
  }
};
const BAD_NEST = 'ul,ol,table,p,div:not(.sw-gap),h1,h2,h3,h4,h5,h6,hr,blockquote';
SW.fixNesting = function(){
  const bad = SW.$$('p,h1,h2,h3,h4,h5,h6', ed).filter(p => Array.from(p.children).some(c => c.matches(BAD_NEST)));
  if (!bad.length) return false;
  SW.keepSelection(() => bad.forEach(p => {
    if (!p.isConnected) return;
    const parts = [];
    let cur = null;
    for (const c of Array.from(p.childNodes)){
      if (c.nodeType === 1 && c.matches(BAD_NEST)){ cur = null; parts.push(c); }
      else {
        if (!cur){ cur = p.cloneNode(false); cur.removeAttribute('data-swp'); parts.push(cur); }
        cur.appendChild(c);
      }
    }
    const keep = parts.filter(n => !(n.nodeName === p.nodeName && n !== p && !n.textContent.trim() && !n.querySelector('img,.sw-fn,.sw-tab')));
    p.replaceWith(...keep);
  }));
  return true;
};

SW.normalizeDoc = function(){
  SW.fixNesting();
  // Editor nie ohne Absatz
  if (!ed.firstElementChild || !ed.textContent && !ed.querySelector('img,table,hr,.sw-toc,.sw-pagebreak,li')){
    if (!ed.querySelector('p,h1,h2,h3,li')) ed.innerHTML = '<p><br></p>';
  }
  // lose Textknoten direkt im Editor in Absätze packen
  Array.from(ed.childNodes).forEach(n => {
    if (n.nodeType === 3 && n.data.trim()){
      const p = document.createElement('p'); n.replaceWith(p); p.appendChild(n);
    } else if (n.nodeType === 1 && n.tagName === 'DIV' && !n.className){
      SW.retag(n, 'p');
    } else if (n.nodeType === 1 && n.tagName === 'BR'){
      n.remove();
    }
  });
  // Nach Tabelle/Umbruch/TOC am Ende muss ein Absatz folgen
  // Spaltenabschnitte: leere entfernen, lose Inhalte in Absätze
  ed.querySelectorAll(':scope > .sw-cols').forEach(c => {
    if (!c.firstElementChild && !c.textContent.trim()){ const p = document.createElement('p'); p.innerHTML = '<br>'; c.replaceWith(p); return; }
    Array.from(c.childNodes).forEach(n => { if (n.nodeType === 3 && n.data.trim()){ const p = document.createElement('p'); n.replaceWith(p); p.appendChild(n); } else if (n.nodeType === 1 && n.tagName === 'BR') n.remove(); });
  });
  const last = ed.lastElementChild;
  if (last && last.matches('table,.sw-toc,.sw-tof,.sw-abbrlist,.sw-cols,.sw-pagebreak,hr')){
    const p = document.createElement('p'); p.innerHTML = '<br>'; ed.appendChild(p);
  }
  // Leere Formatierungs-Spans entfernen
  ed.querySelectorAll('span:not([class]):not([style])').forEach(s => s.replaceWith(...s.childNodes));
  ed.querySelectorAll('font').forEach(f => {
    const span = document.createElement('span');
    if (f.color) span.style.color = f.color;
    if (f.face) span.style.fontFamily = f.face;
    if (f.getAttribute('style')) span.style.cssText += f.getAttribute('style');
    while (f.firstChild) span.appendChild(f.firstChild);
    f.replaceWith(span);
  });
  SW.fixLists();
};

})();
