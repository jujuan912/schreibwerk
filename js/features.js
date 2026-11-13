/* ==========================================================
   Schreibwerk – Erweiterungen: Kommentare, Textspalten,
   erste Seite/Seitenzahlen, Bilder (Umbruch, Verschieben,
   Zuschneiden), Beschriftungen & Verzeichnisse, Formeln
   ========================================================== */
(function(){
'use strict';
const SW = window.SW, ed = SW.ed, $ = SW.$, $$ = SW.$$, C = SW.C, esc = SW.esc;
const paper = $('#paper');

/* ============================================================
   Hilfen
   ============================================================ */
function hasHighlights(){ return !!(window.CSS && CSS.highlights && window.Highlight); }
function closestIn(node, sel){
  const n = node && (node.nodeType === 1 ? node : node.parentNode);
  const r = n && n.closest ? n.closest(sel) : null;
  return r && ed.contains(r) ? r : null;
}
function topLevel(node){
  let n = node.nodeType === 1 ? node : node.parentNode;
  while (n && n.parentNode !== ed) n = n.parentNode;
  return n;
}
/* Markierung in Spans mit eigener Klasse verpacken (wie Word-Bereiche) */
function wrapSelection(make){
  const r = SW.getRange();
  if (!r || r.collapsed) return [];
  document.execCommand('styleWithCSS', false, false);
  document.execCommand('fontSize', false, '7');
  document.execCommand('styleWithCSS', false, true);
  const out = [];
  $$('font[size="7"]', ed).forEach(f => {
    const span = make();
    while (f.firstChild) span.appendChild(f.firstChild);
    f.replaceWith(span);
    out.push(span);
  });
  if (out.length){
    const nr = document.createRange();
    nr.setStartBefore(out[0]); nr.setEndAfter(out[out.length - 1]);
    const s = getSelection(); s.removeAllRanges(); s.addRange(nr);
  }
  return out;
}
function placeCaretAfter(node){
  const r = document.createRange();
  r.setStartAfter(node); r.collapse(true);
  const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  SW.lastRange = r.cloneRange();
}
function placeCaretEnd(el){
  const r = document.createRange();
  r.selectNodeContents(el); r.collapse(false);
  const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  SW.lastRange = r.cloneRange();
}
function author(){
  try { return localStorage.getItem('sw-author') || ''; } catch(e){ return ''; }
}
function setAuthor(n){ try { localStorage.setItem('sw-author', n); } catch(e){} }
function askAuthor(){
  return new Promise(res => {
    const cur = author();
    if (cur){ res(cur); return; }
    SW.dialog({
      title:'Ihr Name', width:380,
      body:'<p style="font-size:13px;margin:0 0 6px">Unter welchem Namen sollen Ihre Kommentare erscheinen?</p><div class="dlg-row"><input id="aN" type="text" placeholder="z. B. Julian"></div>',
      buttons:[{label:'Abbrechen', action:() => res(null)}, {label:'OK', primary:true, action:d => { const v = d.querySelector('#aN').value.trim() || 'Autor'; setAuthor(v); res(v); }}]
    });
  });
}
const fmtDate = iso => { try { return new Date(iso).toLocaleString('de-DE', {day:'2-digit', month:'2-digit', year:'numeric', hour:'2-digit', minute:'2-digit'}); } catch(e){ return ''; } };
const initials = n => (n || '?').split(/\s+/).map(w => w[0] || '').join('').slice(0, 2).toUpperCase();

/* ============================================================
   1) KOMMENTARE AM RAND
   ============================================================ */
let showComments = true;
let activeCid = null;
const PANE_W = 270, PANE_GAP = 22;
const pane = document.createElement('div');
pane.id = 'cmtPane';
pane.className = 'cmt-pane';
paper.appendChild(pane);

const cmtSpans = () => $$('.sw-cmt[data-cid]', ed);
function cmtIds(){
  const seen = [];
  cmtSpans().forEach(s => { if (!seen.includes(s.dataset.cid)) seen.push(s.dataset.cid); });
  return seen;
}
SW.sidePaneWidth = () => (showComments && ed.querySelector('.sw-cmt[data-cid]')) ? PANE_W + PANE_GAP + 10 : 0;

function cmtData(cid){
  const all = SW.doc.comments || (SW.doc.comments = {});
  return all[cid] || (all[cid] = {text:'', author:author() || 'Autor', date:new Date().toISOString(), done:false});
}
function renderComments(){
  const ids = showComments ? cmtIds() : [];
  document.body.classList.toggle('cmt-hidden', !showComments);
  $('#btnShowCmt') && $('#btnShowCmt').classList.toggle('on', showComments && cmtIds().length > 0);
  if (!ids.length){ pane.hidden = true; pane.innerHTML = ''; paintActive(); return; }
  pane.hidden = false;
  const g = SW.geo(), z = SW.zoom;
  pane.style.left = (g.W + PANE_GAP) + 'px';
  pane.style.width = PANE_W + 'px';
  const pr = paper.getBoundingClientRect();
  // vorhandene Karten wiederverwenden (Fokus beim Tippen bleibt erhalten)
  const existing = {};
  $$('.cmt-card', pane).forEach(c => existing[c.dataset.cid] = c);
  let y = 0;
  ids.forEach(cid => {
    const d = cmtData(cid);
    let card = existing[cid];
    delete existing[cid];
    if (!card){
      card = document.createElement('div');
      card.className = 'cmt-card';
      card.dataset.cid = cid;
      card.innerHTML = `<div class="cmt-head"><span class="cmt-ava"></span><div class="cmt-who"><b></b><small></small></div><button class="cmt-btn cmt-done" title="Als erledigt markieren">${SW.icon('check')}</button><button class="cmt-btn cmt-del" title="Kommentar löschen">${SW.icon('trash-2')}</button></div><div class="cmt-text" data-ph="Kommentar eingeben …"></div>`;
      const t = card.querySelector('.cmt-text');
      try { t.contentEditable = 'plaintext-only'; } catch(e){ t.contentEditable = 'true'; }
      if (t.contentEditable !== 'plaintext-only') t.contentEditable = 'true';
      t.addEventListener('input', () => { cmtData(cid).text = t.innerText.replace(/\n$/, ''); SW.markDirty(); SW.scheduleCommit(); });
      t.addEventListener('focus', () => { setActive(cid); });
      t.addEventListener('blur', () => { SW.flushCommit(); });
      t.addEventListener('keydown', e => { if (e.key === 'Escape'){ e.preventDefault(); gotoComment(cid, true); } e.stopPropagation(); });
      card.querySelector('.cmt-del').addEventListener('click', () => deleteComment(cid));
      card.querySelector('.cmt-done').addEventListener('click', () => { SW.edit(() => { cmtData(cid).done = !cmtData(cid).done; }); });
      card.addEventListener('mousedown', e => { if (!e.target.closest('.cmt-text,.cmt-btn')){ e.preventDefault(); gotoComment(cid, false); } });
      pane.appendChild(card);
    }
    card.querySelector('.cmt-ava').textContent = initials(d.author);
    card.querySelector('.cmt-who b').textContent = d.author || 'Autor';
    card.querySelector('.cmt-who small').textContent = fmtDate(d.date) + (d.done ? ' · erledigt' : '');
    card.classList.toggle('done', !!d.done);
    card.classList.toggle('active', cid === activeCid);
    const t = card.querySelector('.cmt-text');
    if (document.activeElement !== t && t.innerText.replace(/\n$/, '') !== (d.text || '')) t.textContent = d.text || '';
    // Position: auf Höhe der Textstelle, Überlappungen nach unten schieben
    const first = ed.querySelector(`.sw-cmt[data-cid="${cid}"]`);
    const ay = first ? (first.getBoundingClientRect().top - pr.top) / z : y;
    const top = Math.max(ay - 6, y);
    card.style.top = top + 'px';
    y = top + card.offsetHeight + 8;
  });
  Object.values(existing).forEach(c => c.remove());
  paintActive();
}
SW.renderComments = renderComments;
function paintActive(){
  if (!hasHighlights()) return;
  if (!activeCid || !showComments){ CSS.highlights.delete('sw-cmt-active'); return; }
  const ranges = $$(`.sw-cmt[data-cid="${activeCid}"]`, ed).map(s => { const r = new Range(); r.selectNodeContents(s); return r; });
  if (ranges.length) CSS.highlights.set('sw-cmt-active', new Highlight(...ranges)); else CSS.highlights.delete('sw-cmt-active');
}
function setActive(cid){
  if (activeCid === cid) return;
  activeCid = cid;
  $$('.cmt-card', pane).forEach(c => c.classList.toggle('active', c.dataset.cid === cid));
  paintActive();
}
function gotoComment(cid, focusText){
  const span = ed.querySelector(`.sw-cmt[data-cid="${cid}"]`);
  if (!span) return;
  setActive(cid);
  const r = document.createRange(); r.selectNodeContents(span); r.collapse(true);
  ed.focus({preventScroll:true});
  const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  SW.scrollToEl(span);
  if (focusText){
    const t = pane.querySelector(`.cmt-card[data-cid="${cid}"] .cmt-text`);
    if (t){ t.focus(); placeCaretEnd(t); }
  }
}
function deleteComment(cid){
  if (!cid) return;
  SW.edit(() => {
    $$(`.sw-cmt[data-cid="${cid}"]`, ed).forEach(s => s.replaceWith(...s.childNodes));
    delete (SW.doc.comments || {})[cid];
    ed.normalize();
  });
  if (activeCid === cid) activeCid = null;
  SW.relayout(true);
}
function cidAtCaret(){
  const r = SW.getRange();
  if (!r) return null;
  const s = closestIn(r.startContainer, '.sw-cmt[data-cid]');
  return s ? s.dataset.cid : null;
}
C.commentNew = async () => {
  SW.focusEditor();
  SW.rememberRange();
  let r = SW.getRange();
  if (r && r.collapsed){ SW.expandToWord(); SW._wordExpanded = null; r = SW.getRange(); }
  if (!r || r.collapsed){ SW.toast('Bitte zuerst den Text markieren, zu dem der Kommentar gehören soll.'); return; }
  const saved = r.cloneRange();
  const name = await askAuthor();
  if (!name) return;
  ed.focus({preventScroll:true});
  const s = getSelection(); s.removeAllRanges(); s.addRange(saved);
  const cid = 'c' + SW.uid();
  showComments = true;
  SW.edit(() => {
    wrapSelection(() => { const sp = document.createElement('span'); sp.className = 'sw-cmt'; sp.dataset.cid = cid; return sp; });
    SW.doc.comments = SW.doc.comments || {};
    SW.doc.comments[cid] = {text:'', author:name, date:new Date().toISOString(), done:false};
  });
  SW.relayout(true);
  activeCid = cid;
  renderComments();
  const t = pane.querySelector(`.cmt-card[data-cid="${cid}"] .cmt-text`);
  if (t){ t.focus(); }
};
C.commentDelete = () => {
  const cid = cidAtCaret() || activeCid;
  if (!cid){ SW.toast('Setzen Sie den Cursor in einen kommentierten Text oder klicken Sie auf einen Kommentar.'); return; }
  deleteComment(cid);
};
function stepComment(dir){
  const ids = cmtIds();
  if (!ids.length){ SW.toast('Das Dokument enthält keine Kommentare.'); return; }
  const cur = cidAtCaret() || activeCid;
  let i = ids.indexOf(cur);
  i = i < 0 ? (dir > 0 ? 0 : ids.length - 1) : (i + dir + ids.length) % ids.length;
  showComments = true;
  SW.relayout(true);
  gotoComment(ids[i], false);
}
C.commentNext = () => stepComment(1);
C.commentPrev = () => stepComment(-1);
C.toggleComments = () => { showComments = !showComments; SW.relayout(true); renderComments(); };

/* ============================================================
   2) TEXTSPALTEN
   ============================================================ */
C.columns = n => {
  n = +n;
  SW.edit(() => {
    const sel = SW.saveSel();
    SW.clearArtifacts();               // Fortsetzungen auf Folgeseiten zusammenführen
    SW.restoreSel(sel);
    const r = SW.getRange();
    if (!r) return;
    const inCols = closestIn(r.startContainer, '.sw-cols');
    const endCols = closestIn(r.endContainer, '.sw-cols');
    if (inCols && (r.collapsed || inCols === endCols)){
      if (n <= 1) inCols.replaceWith(...inCols.childNodes);
      else inCols.dataset.n = n;
      SW.restoreSel(sel);
      return;
    }
    // betroffene Blöcke der obersten Ebene
    let blocks;
    if (r.collapsed) blocks = Array.from(ed.children);
    else {
      const a = topLevel(r.startContainer), b = topLevel(r.endContainer);
      blocks = [];
      let on = false;
      for (const c of Array.from(ed.children)){ if (c === a) on = true; if (on) blocks.push(c); if (c === b) break; }
    }
    // vorhandene Spaltenabschnitte auflösen
    const flat = [];
    blocks.forEach(b => { if (b.classList && b.classList.contains('sw-cols')){ const kids = Array.from(b.children); b.replaceWith(...kids); flat.push(...kids); } else flat.push(b); });
    if (n > 1 && flat.length){
      // Seitenumbrüche bleiben außerhalb
      while (flat.length && flat[0].classList.contains('sw-pagebreak')) flat.shift();
      while (flat.length && flat[flat.length - 1].classList.contains('sw-pagebreak')) flat.pop();
      const box = document.createElement('div');
      box.className = 'sw-cols';
      box.dataset.n = n;
      flat[0].before(box);
      flat.forEach(b => box.appendChild(b));
    }
    SW.restoreSel(sel);
  });
};
C.colLine = () => {
  const sel = SW.saveSel(); SW.clearArtifacts(); SW.restoreSel(sel);
  const r = SW.getRange() || SW.lastRange;
  const c = r && closestIn(r.startContainer, '.sw-cols');
  if (!c){ SW.toast('Setzen Sie den Cursor zuerst in einen Spaltenabschnitt.'); return; }
  SW.edit(() => { if (c.dataset.line === '1') delete c.dataset.line; else c.dataset.line = '1'; });
};

/* ============================================================
   3) ERSTE SEITE ANDERS & SEITENZAHLEN
   ============================================================ */
const PN_RE = /Seite \{SEITE\} von \{SEITEN\}|\{SEITE\}/g;
const stripPN = t => (t || '').replace(PN_RE, '').replace(/\n{2,}/g, '\n').trim();
$('#chkTitlePage').addEventListener('change', e => {
  SW.edit(() => { SW.doc.titlePage = e.target.checked; });
  if (e.target.checked) SW.toast('Erste Seite hat jetzt eigene Kopf-/Fußzeile (leer = Deckblatt ohne Seitenzahl). Doppelklick auf die Kopfzeile von Seite 1 zum Bearbeiten.', null, 6000);
});
C.pageNumberDialog = () => {
  const D = SW.doc;
  const pn = D.pageNum || {start:1, fmt:'decimal'};
  const hasIn = t => /\{SEITE\}/.test(t || '');
  let pos = 'f-center';
  if (hasIn(D.footer.text)) pos = 'f-' + D.footer.align;
  else if (hasIn(D.header.text)) pos = 'h-' + D.header.align;
  const ofY = /\{SEITEN\}/.test(D.footer.text + D.header.text);
  SW.dialog({
    title:'Seitenzahlen', width:440,
    body:`<div class="dlg-row"><label>Position</label><select id="pnPos">
        <option value="f-center">Seitenende – zentriert</option><option value="f-right">Seitenende – rechts</option><option value="f-left">Seitenende – links</option>
        <option value="h-right">Seitenanfang – rechts</option><option value="h-center">Seitenanfang – zentriert</option><option value="h-left">Seitenanfang – links</option>
        <option value="none">Keine Seitenzahlen</option></select></div>
      <div class="dlg-row"><label>Zahlenformat</label><select id="pnFmt">
        <option value="decimal">1, 2, 3 …</option><option value="lowerRoman">i, ii, iii …</option><option value="upperRoman">I, II, III …</option>
        <option value="lowerLetter">a, b, c …</option><option value="upperLetter">A, B, C …</option></select></div>
      <div class="dlg-row"><label>Beginnen bei</label><input id="pnStart" type="number" min="0" max="9999" value="${pn.start ?? 1}"></div>
      <div class="dlg-row"><label class="chk" style="min-width:0"><input type="checkbox" id="pnOf"${ofY ? ' checked' : ''}> Als „Seite X von Y“</label></div>
      <div class="dlg-row"><label class="chk" style="min-width:0"><input type="checkbox" id="pnTitle"${D.titlePage ? ' checked' : ''}> Erste Seite ohne Seitenzahl (Deckblatt)</label></div>
      <p class="muted">Tipp: Mit „Beginnen bei 0“ und Deckblatt bekommt die zweite Seite die Nummer 1.</p>`,
    onOpen:d => { d.querySelector('#pnPos').value = pos; d.querySelector('#pnFmt').value = pn.fmt || 'decimal'; },
    buttons:[{label:'Abbrechen'}, {label:'OK', primary:true, action:d => {
      const p = d.querySelector('#pnPos').value, fmt = d.querySelector('#pnFmt').value;
      const start = Math.max(0, parseInt(d.querySelector('#pnStart').value, 10) || 0);
      const tok = d.querySelector('#pnOf').checked ? 'Seite {SEITE} von {SEITEN}' : '{SEITE}';
      const title = d.querySelector('#pnTitle').checked;
      SW.edit(() => {
        D.header.text = stripPN(D.header.text); D.footer.text = stripPN(D.footer.text);
        if (p !== 'none'){
          const [where, align] = p.split('-');
          const tgt = where === 'f' ? D.footer : D.header;
          tgt.text = (tgt.text ? tgt.text + '\n' : '') + tok;
          tgt.align = align;
        }
        D.pageNum = {start, fmt};
        D.titlePage = title;
        if (title){
          D.firstHeader = {text:stripPN((D.firstHeader || {}).text), align:(D.firstHeader || {}).align || 'left'};
          D.firstFooter = {text:stripPN((D.firstFooter || {}).text), align:(D.firstFooter || {}).align || 'center'};
        }
      });
    }}]
  });
};
// Kopf- und Fußzeile mit „Erste Seite anders“
C.headerFooter = which => {
  const D = SW.doc;
  D.firstHeader = D.firstHeader || {text:'', align:'left'};
  D.firstFooter = D.firstFooter || {text:'', align:'center'};
  const sel = a => ['left','center','right'].map(v => `<option value="${v}"${a === v ? ' selected' : ''}>${{left:'Links', center:'Zentriert', right:'Rechts'}[v]}</option>`).join('');
  const body = `
    <div class="dlg-section">Kopfzeile</div>
    <div class="dlg-row"><textarea id="hT" placeholder="Text der Kopfzeile">${esc(D.header.text)}</textarea><select id="hA" style="flex:0 0 110px">${sel(D.header.align)}</select></div>
    <div class="dlg-section">Fußzeile</div>
    <div class="dlg-row"><textarea id="fT" placeholder="Text der Fußzeile">${esc(D.footer.text)}</textarea><select id="fA" style="flex:0 0 110px">${sel(D.footer.align)}</select></div>
    <div class="dlg-row"><label class="chk" style="min-width:0"><input type="checkbox" id="tP"${D.titlePage ? ' checked' : ''}> Erste Seite anders (z. B. Deckblatt)</label></div>
    <div id="firstBox"${D.titlePage ? '' : ' hidden'}>
      <div class="dlg-section">Kopfzeile erste Seite</div>
      <div class="dlg-row"><textarea id="fhT" placeholder="leer lassen = keine Kopfzeile auf Seite 1">${esc(D.firstHeader.text)}</textarea><select id="fhA" style="flex:0 0 110px">${sel(D.firstHeader.align)}</select></div>
      <div class="dlg-section">Fußzeile erste Seite</div>
      <div class="dlg-row"><textarea id="ffT" placeholder="leer lassen = keine Seitenzahl auf Seite 1">${esc(D.firstFooter.text)}</textarea><select id="ffA" style="flex:0 0 110px">${sel(D.firstFooter.align)}</select></div>
    </div>
    <p class="muted">Platzhalter: <b>{SEITE}</b> = Seitenzahl, <b>{SEITEN}</b> = Seitenanzahl. Beispiel: „Seite {SEITE} von {SEITEN}“.</p>`;
  SW.dialog({
    title:'Kopf- und Fußzeile', width:520, body,
    onOpen:d => {
      d.querySelector('#tP').addEventListener('change', e => { d.querySelector('#firstBox').hidden = !e.target.checked; });
      const f = d.querySelector(which === 'footer' ? '#fT' : which === 'fheader' ? '#fhT' : which === 'ffooter' ? '#ffT' : '#hT');
      setTimeout(() => f.focus(), 0);
    },
    buttons:[{label:'Abbrechen'}, {label:'OK', primary:true, action:d => {
      const v = id => d.querySelector(id).value;
      SW.edit(() => {
        D.header = {text:v('#hT'), align:v('#hA')};
        D.footer = {text:v('#fT'), align:v('#fA')};
        D.titlePage = d.querySelector('#tP').checked;
        D.firstHeader = {text:v('#fhT'), align:v('#fhA')};
        D.firstFooter = {text:v('#ffT'), align:v('#ffA')};
      });
    }}]
  });
};
// Doppelklick auf Kopf-/Fußzeile der ersten Seite → passende Felder
paper.addEventListener('dblclick', e => {
  const hf = e.target.closest('.hf');
  if (!hf) return;
  const page = hf.closest('.page');
  if (SW.doc.titlePage && page && page === $('#pages').firstElementChild){
    e.stopImmediatePropagation();
    C.headerFooter(hf.classList.contains('hf-h') ? 'fheader' : 'ffooter');
  }
}, true);

/* ============================================================
   4) BILDER: TEXTUMBRUCH, FREI VERSCHIEBEN, ZUSCHNEIDEN
   ============================================================ */
const blockOf = el => el.closest('p,h1,h2,h3,h4,h5,h6,li,td,th,blockquote') || null;
C.imgWrap = mode => {
  const img = SW.selectedImage();
  if (!img){ SW.toast('Bitte zuerst ein Bild anklicken.'); return; }
  const z = SW.zoom;
  const ir = img.getBoundingClientRect();
  SW.edit(() => {
    let blk = blockOf(img);
    if (!blk){ blk = SW.wrapLooseInline(img); }
    if (mode === 'inline'){ delete img.dataset.wrap; img.style.left = ''; img.style.top = ''; }
    else img.dataset.wrap = mode;
    if (mode === 'front' || mode === 'behind'){
      const br = blk.getBoundingClientRect();
      img.style.left = Math.round((ir.left - br.left) / z) + 'px';
      img.style.top = Math.round((ir.top - br.top) / z) + 'px';
    } else { img.style.left = ''; img.style.top = ''; }
  });
  setTimeout(() => { SW.selectImage(img); SW.positionHandles(); }, 30);
};
/* Bild ziehen */
let ghost = null, dropMark = null;
SW.startImageDrag = function(e, img){
  if (e.button !== 0) return;
  const mode = img.dataset.wrap || 'inline';
  const z = SW.zoom;
  const x0 = e.clientX, y0 = e.clientY;
  const l0 = parseFloat(img.style.left) || 0, t0 = parseFloat(img.style.top) || 0;
  const free = mode === 'front' || mode === 'behind';
  let moved = false;
  SW.flushCommit();
  const mv = ev => {
    const dx = ev.clientX - x0, dy = ev.clientY - y0;
    if (!moved && Math.hypot(dx, dy) < 5) return;
    if (!moved){
      moved = true;
      document.body.classList.add('img-dragging');
      if (!free){
        ghost = img.cloneNode();
        ghost.className = 'img-ghost';
        const r = img.getBoundingClientRect();
        ghost.style.cssText = `position:fixed;left:0;top:0;width:${r.width}px;height:${r.height}px;pointer-events:none;opacity:.55;z-index:90;outline:2px dashed #185abd`;
        document.body.appendChild(ghost);
        dropMark = document.createElement('div');
        dropMark.className = 'drop-mark';
        document.body.appendChild(dropMark);
      }
    }
    if (free){
      img.style.left = Math.round(l0 + dx / z) + 'px';
      img.style.top = Math.round(t0 + dy / z) + 'px';
      SW.positionHandles();
    } else {
      ghost.style.transform = `translate(${ev.clientX + 8}px, ${ev.clientY + 8}px)`;
      const cr = document.caretRangeFromPoint && document.caretRangeFromPoint(ev.clientX, ev.clientY);
      const rect = cr && SW.inEditor(cr.startContainer) ? (cr.getClientRects()[0] || (cr.startContainer.nodeType === 1 ? cr.startContainer.getBoundingClientRect() : null)) : null;
      if (rect){ dropMark.hidden = false; dropMark.style.cssText = `left:${rect.left}px;top:${rect.top}px;height:${Math.max(14, rect.height)}px`; }
      else dropMark.hidden = true;
    }
  };
  const up = ev => {
    removeEventListener('mousemove', mv); removeEventListener('mouseup', up);
    document.body.classList.remove('img-dragging');
    if (ghost){ ghost.remove(); ghost = null; }
    if (dropMark){ dropMark.remove(); dropMark = null; }
    if (!moved) return;
    if (free){
      // An den Absatz unter dem Bild hängen, damit es beim Umbruch mitwandert
      const ir = img.getBoundingClientRect();
      img.style.visibility = 'hidden';
      const under = document.elementFromPoint(ir.left + 2, ir.top + 2);
      img.style.visibility = '';
      const nb = under && SW.inEditor(under) ? blockOf(under) : null;
      const ob = blockOf(img);
      if (nb && nb !== ob && !nb.contains(img)){
        const br = nb.getBoundingClientRect();
        nb.insertBefore(img, nb.firstChild);
        img.style.left = Math.round((ir.left - br.left) / z) + 'px';
        img.style.top = Math.round((ir.top - br.top) / z) + 'px';
      }
      SW.normalizeDoc(); SW.commit(); SW.markDirty(); SW.relayout(true);
      SW.selectImage(img);
      return;
    }
    const cr = document.caretRangeFromPoint && document.caretRangeFromPoint(ev.clientX, ev.clientY);
    if (!cr || !SW.inEditor(cr.startContainer) || img.contains(cr.startContainer)) return;
    const er = ed.getBoundingClientRect();
    SW.edit(() => {
      const oldBlk = blockOf(img);
      cr.insertNode(img);
      if (mode === 'left' || mode === 'right') img.dataset.wrap = ev.clientX > er.left + er.width / 2 ? 'right' : 'left';
      if (oldBlk && oldBlk !== blockOf(img) && !oldBlk.textContent.trim() && !oldBlk.querySelector('img,br') && oldBlk.matches('p')) oldBlk.innerHTML = '<br>';
    });
    SW.selectImage(img);
  };
  addEventListener('mousemove', mv); addEventListener('mouseup', up);
};

/* Zuschneiden */
C.imgCrop = () => {
  const img = SW.selectedImage();
  if (!img){ SW.toast('Bitte zuerst ein Bild anklicken.'); return; }
  const nw = img.naturalWidth, nh = img.naturalHeight;
  if (!nw || !nh) return;
  const s = Math.min(520 / nw, 360 / nh, 1);
  const W = Math.round(nw * s), H = Math.round(nh * s);
  let c = {x:0, y:0, w:W, h:H};
  const hasOrig = img.dataset.orig && SW.images.get(img.dataset.orig);
  const dlg = SW.dialog({
    title:'Bild zuschneiden', width:Math.max(420, W + 76),
    body:`<p class="muted" style="margin:0 0 8px">Ziehen Sie an den Ecken und Kanten des Rahmens. Den Rahmen selbst können Sie verschieben.</p>
      <div class="crop-outer"><div class="crop-wrap" style="width:${W}px;height:${H}px"><img src="${img.src}" alt="" draggable="false"><div class="crop-box">${['nw','n','ne','e','se','s','sw','w'].map(h => `<i data-h="${h}"></i>`).join('')}</div></div></div>
      <div class="dlg-row" style="margin-top:10px"><span id="cropInfo" class="muted"></span></div>`,
    buttons:[
      ...(hasOrig ? [{label:'Originalbild', action:() => {
        const rec = SW.images.get(img.dataset.orig);
        const w0 = img.getBoundingClientRect().width / SW.zoom;
        SW.edit(() => { img.dataset.imgId = img.dataset.orig; delete img.dataset.orig; img.src = rec.url; img.style.width = Math.round(w0) + 'px'; });
        setTimeout(() => { SW.relayout(true); SW.selectImage(img); }, 80);
      }}] : []),
      {label:'Abbrechen'},
      {label:'Zuschneiden', primary:true, action:async () => {
        const sx = Math.round(c.x / s), sy = Math.round(c.y / s), sw = Math.max(1, Math.round(c.w / s)), sh = Math.max(1, Math.round(c.h / s));
        if (sw >= nw - 1 && sh >= nh - 1) return;
        const cv = document.createElement('canvas'); cv.width = sw; cv.height = sh;
        cv.getContext('2d').drawImage(img, sx, sy, sw, sh, 0, 0, sw, sh);
        const oldRec = SW.images.get(img.dataset.imgId);
        const type = oldRec && /jpe?g/.test(oldRec.type) ? 'image/jpeg' : 'image/png';
        const blob = await new Promise(res => cv.toBlob(res, type, 0.92));
        const {id, url} = SW.registerImage(blob);
        const w0 = img.getBoundingClientRect().width / SW.zoom;
        SW.edit(() => {
          if (!img.dataset.orig) img.dataset.orig = img.dataset.imgId;
          img.dataset.imgId = id; img.src = url;
          img.style.width = Math.max(16, Math.round(w0 * sw / nw)) + 'px';
        });
        setTimeout(() => { SW.relayout(true); SW.selectImage(img); }, 80);
      }}
    ]
  });
  const box = dlg.el.querySelector('.crop-box'), info = dlg.el.querySelector('#cropInfo');
  const draw = () => {
    box.style.cssText = `left:${c.x}px;top:${c.y}px;width:${c.w}px;height:${c.h}px`;
    info.textContent = `Ausschnitt: ${Math.round(c.w / s)} × ${Math.round(c.h / s)} Pixel`;
  };
  draw();
  box.addEventListener('mousedown', e => {
    e.preventDefault();
    const h = e.target.dataset.h || 'move';
    const x0 = e.clientX, y0 = e.clientY, c0 = {...c};
    const mv = ev => {
      const dx = ev.clientX - x0, dy = ev.clientY - y0;
      let {x, y, w, h: hh} = c0;
      if (h === 'move'){ x = Math.max(0, Math.min(W - w, x + dx)); y = Math.max(0, Math.min(H - hh, y + dy)); }
      else {
        if (h.includes('w')){ const nx = Math.max(0, Math.min(x + w - 10, x + dx)); w += x - nx; x = nx; }
        if (h.includes('e')) w = Math.max(10, Math.min(W - x, w + dx));
        if (h.includes('n')){ const ny = Math.max(0, Math.min(y + hh - 10, y + dy)); hh += y - ny; y = ny; }
        if (h.includes('s')) hh = Math.max(10, Math.min(H - y, hh + dy));
      }
      c = {x, y, w, h:hh}; draw();
    };
    const up = () => { removeEventListener('mousemove', mv); removeEventListener('mouseup', up); };
    addEventListener('mousemove', mv); addEventListener('mouseup', up);
  });
};

/* ============================================================
   5) BESCHRIFTUNGEN, ABBILDUNGS- UND ABKÜRZUNGSVERZEICHNIS
   ============================================================ */
C.caption = () => {
  SW.rememberRange();
  const img = SW.selectedImage();
  const r = SW.lastRange;
  const cell = r && closestIn(r.startContainer, 'td,th');
  const math = r && closestIn(r.startContainer, '.sw-mathblock');
  const def = img ? 'Abbildung' : cell ? 'Tabelle' : math ? 'Formel' : 'Abbildung';
  const count = t => $$(`.sw-seq[data-seq="${t}"]`, ed).length + 1;
  SW.dialog({
    title:'Beschriftung', width:440,
    body:`<div class="dlg-row"><label style="min-width:90px">Bezeichnung</label><select id="cpT"><option>Abbildung</option><option>Tabelle</option><option>Formel</option></select></div>
      <div class="dlg-row"><label style="min-width:90px">Beschriftung</label><span id="cpL" class="cap-label"></span><input id="cpX" type="text" placeholder="Beschreibung, z. B. Aufbau der Steuerung"></div>
      <div class="dlg-row"><label style="min-width:90px">Position</label><select id="cpP"><option value="below">Unter dem ausgewählten Element</option><option value="above">Über dem ausgewählten Element</option></select></div>
      <p class="muted">Die Nummern zählen automatisch weiter. Beschriftungen erscheinen im Abbildungs- bzw. Tabellenverzeichnis.</p>`,
    onOpen:d => {
      const t = d.querySelector('#cpT'), l = d.querySelector('#cpL'), p = d.querySelector('#cpP');
      t.value = def;
      const upd = () => { l.textContent = `${t.value} ${count(t.value)}:`; p.value = t.value === 'Tabelle' ? 'above' : 'below'; };
      t.addEventListener('change', upd); upd();
      setTimeout(() => d.querySelector('#cpX').focus(), 0);
    },
    buttons:[{label:'Abbrechen'}, {label:'OK', primary:true, action:d => {
      const type = d.querySelector('#cpT').value, txt = d.querySelector('#cpX').value.trim(), pos = d.querySelector('#cpP').value;
      SW.edit(() => {
        const p = document.createElement('p');
        p.className = 'sw-caption';
        p.innerHTML = `<span class="sw-seq" data-seq="${esc(type)}" contenteditable="false">${esc(type)} ${count(type)}</span>${txt ? ': ' + esc(txt) : ':\u00A0'}`;
        let target = null;
        if (img) target = img.closest('table') && type === 'Tabelle' ? img.closest('table') : (blockOf(img) || img);
        else if (cell) target = cell.closest('table');
        else if (SW.lastRange) target = SW.innermostBlock(SW.lastRange.startContainer);
        if (target && target.closest && target.closest('td,th') && type === 'Tabelle') target = target.closest('table');
        if (target && target.matches && target.matches('td,th')) target = target.closest('table');
        if (target && target.closest('li')) target = target.closest('ul,ol');
        // innerhalb von Tabellen/Listen: an die oberste Ebene
        if (target && type !== 'Abbildung'){ while (target.parentNode !== ed && !target.parentNode.classList.contains('sw-cols')) target = target.parentNode; }
        if (!target){ ed.appendChild(p); }
        else if (pos === 'above') target.before(p);
        else target.after(p);
        placeCaretEnd(p);
      });
      SW.relayout(true);
    }}]
  });
};
C.insertTof = type => {
  SW.edit(() => {
    const t = document.createElement('div');
    t.className = 'sw-tof'; t.dataset.seq = type || 'Abbildung'; t.setAttribute('contenteditable', 'false');
    SW.insertBlockAtCaret(t);
  });
  SW.relayout(true);
};
C.insertAbbrList = () => {
  if (ed.querySelector('.sw-abbrlist')){ SW.toast('Das Dokument enthält bereits ein Abkürzungsverzeichnis.'); return; }
  SW.edit(() => {
    const t = document.createElement('div');
    t.className = 'sw-abbrlist'; t.setAttribute('contenteditable', 'false');
    SW.insertBlockAtCaret(t);
  });
  SW.relayout(true);
  if (!Object.keys(SW.doc.abbr || {}).length) SW.toast('Erfassen Sie Abkürzungen über „Abkürzung definieren“ oder „Abkürzungen verwalten“.', null, 5000);
};
C.abbrDefine = () => {
  SW.rememberRange();
  let r = SW.getRange();
  if (r && r.collapsed){ SW.expandToWord(); SW._wordExpanded = null; r = SW.getRange(); }
  const word = r ? r.toString().trim() : '';
  const ab = SW.doc.abbr || {};
  SW.dialog({
    title:'Abkürzung definieren', width:440,
    body:`<div class="dlg-row"><label style="min-width:90px">Abkürzung</label><input id="abK" type="text" value="${esc(word)}" placeholder="z. B. SPS"></div>
      <div class="dlg-row"><label style="min-width:90px">Bedeutung</label><input id="abV" type="text" value="${esc(ab[word] || '')}" placeholder="z. B. Speicherprogrammierbare Steuerung"></div>`,
    onOpen:d => setTimeout(() => (word ? d.querySelector('#abV') : d.querySelector('#abK')).focus(), 0),
    buttons:[{label:'Abbrechen'}, {label:'OK', primary:true, action:d => {
      const k = d.querySelector('#abK').value.trim(), v = d.querySelector('#abV').value.trim();
      if (!k) return false;
      SW.edit(() => { SW.doc.abbr = {...(SW.doc.abbr || {}), [k]:v}; });
      SW.toast(`„${k}“ ins Abkürzungsverzeichnis übernommen.`);
    }}]
  });
};
C.abbrManage = () => {
  const work = Object.entries(SW.doc.abbr || {}).map(([k, v]) => ({k, v}));
  const dlg = SW.dialog({
    title:'Abkürzungen verwalten', width:580,
    body:`<div class="abbr-edit"><table><thead><tr><th>Abkürzung</th><th>Bedeutung</th><th></th></tr></thead><tbody id="abRows"></tbody></table></div>
      <div class="dlg-row" style="gap:8px"><button class="btn" id="abAdd">+ Zeile</button><button class="btn" id="abScan">Im Text suchen</button><span class="muted" id="abInfo"></span></div>`,
    buttons:[{label:'Abbrechen'}, {label:'Übernehmen', primary:true, action:() => {
      const o = {};
      work.forEach(e => { const k = e.k.trim(); if (k) o[k] = e.v.trim(); });
      SW.edit(() => { SW.doc.abbr = o; });
    }}]
  });
  const rows = dlg.el.querySelector('#abRows');
  const render = () => {
    work.sort((a, b) => !a.k ? 1 : !b.k ? -1 : a.k.localeCompare(b.k, 'de', {sensitivity:'base'}));
    rows.innerHTML = work.map((e, i) => `<tr><td><input data-i="${i}" data-f="k" value="${esc(e.k)}"></td><td><input data-i="${i}" data-f="v" value="${esc(e.v)}" placeholder="Bedeutung eingeben"></td><td><button class="dlg-x" data-del="${i}" title="Entfernen">${SW.icon('x')}</button></td></tr>`).join('') ||
      '<tr><td colspan="3" class="muted" style="padding:10px">Noch keine Abkürzungen.</td></tr>';
  };
  rows.addEventListener('input', e => { const t = e.target; if (t.dataset.i) work[+t.dataset.i][t.dataset.f] = t.value; });
  rows.addEventListener('click', e => { const b = e.target.closest('[data-del]'); if (b){ work.splice(+b.dataset.del, 1); render(); } });
  dlg.el.querySelector('#abAdd').onclick = () => { work.push({k:'', v:''}); render(); const ins = rows.querySelectorAll('input[data-f="k"]'); ins[ins.length - 1] && ins[ins.length - 1].focus(); };
  dlg.el.querySelector('#abScan').onclick = () => {
    const text = SW.cleanClone().textContent;
    const found = new Set((text.match(/\b[A-ZÄÖÜ][A-ZÄÖÜ0-9]{1,6}s?\b/g) || []).map(w => w.replace(/s$/, '')).filter(w => w.length >= 2));
    let n = 0;
    found.forEach(w => { if (!work.some(e => e.k === w)){ work.push({k:w, v:''}); n++; } });
    render();
    dlg.el.querySelector('#abInfo').textContent = n ? `${n} mögliche Abkürzungen gefunden – Bedeutung ergänzen oder Zeile entfernen.` : 'Keine neuen Abkürzungen gefunden.';
  };
  render();
};

/* ============================================================
   6) FORMEL-EDITOR (LaTeX → MathML, offline)
   ============================================================ */
SW.renderMath = function(span){
  const tex = span.dataset.tex || '';
  const display = !!span.closest('.sw-mathblock');
  try {
    span.innerHTML = '';
    if (window.temml) temml.render(tex, span, {displayMode:display, throwOnError:false, annotate:false});
    else span.textContent = tex;
  } catch(e){ span.textContent = tex; }
  span.setAttribute('contenteditable', 'false');
};
SW.renderAllMath = function(root = ed){
  root.querySelectorAll('.sw-math').forEach(s => { if (!s.querySelector('math') || s.dataset.r !== s.dataset.tex){ SW.renderMath(s); s.dataset.r = s.dataset.tex; } });
};
const PALETTE = [
  ['Struktur', [
    ['\\frac{a}{b}', 'Bruch', '\\frac{□}{□}'], ['x^{2}', 'Hochgestellt', '^{□}'], ['x_{i}', 'Tiefgestellt', '_{□}'],
    ['\\sqrt{x}', 'Wurzel', '\\sqrt{□}'], ['\\sqrt[n]{x}', 'n-te Wurzel', '\\sqrt[□]{□}'],
    ['\\sum_{i=1}^{n}', 'Summe', '\\sum_{i=1}^{n} '], ['\\int_{a}^{b}', 'Integral', '\\int_{□}^{□} '], ['\\lim_{x\\to\\infty}', 'Grenzwert', '\\lim_{x \\to \\infty} '],
    ['\\left(x\\right)', 'Klammern', '\\left( □ \\right)'], ['\\vec{v}', 'Vektor', '\\vec{□}'], ['\\overline{x}', 'Mittelwert/Strich', '\\overline{□}'], ['\\dot{x}', 'Ableitung (Punkt)', '\\dot{□}'],
    ['\\begin{pmatrix}a&b\\\\c&d\\end{pmatrix}', 'Matrix 2×2', '\\begin{pmatrix} □ & □ \\\\ □ & □ \\end{pmatrix}'], ['\\text{Text}', 'Normaler Text', '\\text{□}']
  ]],
  ['Operatoren', [
    ['\\cdot', 'Mal (Punkt)'], ['\\times', 'Mal (Kreuz)'], ['\\div', 'Geteilt'], ['\\pm', 'Plus-Minus'], ['=', 'Gleich'], ['\\approx', 'Ungefähr'], ['\\neq', 'Ungleich'],
    ['\\leq', 'Kleiner gleich'], ['\\geq', 'Größer gleich'], ['\\sim', 'Proportional'], ['\\to', 'Pfeil'], ['\\Rightarrow', 'Folgt'], ['\\Leftrightarrow', 'Äquivalent'], ['\\infty', 'Unendlich'],
    ['\\partial', 'Partiell'], ['\\nabla', 'Nabla'], ['^{\\circ}', 'Grad'], ['\\%', 'Prozent'], ['\\in', 'Element von'], ['\\forall', 'Für alle']
  ]],
  ['Griechisch', [
    ['\\alpha'], ['\\beta'], ['\\gamma'], ['\\delta'], ['\\varepsilon'], ['\\eta'], ['\\vartheta'], ['\\lambda'], ['\\mu'], ['\\pi'], ['\\rho'], ['\\sigma'], ['\\tau'], ['\\varphi'], ['\\omega'],
    ['\\Delta'], ['\\Phi'], ['\\Sigma'], ['\\Omega']
  ]]
];
const EXAMPLES = [
  ['Ohmsches Gesetz', 'U = R \\cdot I'], ['Leistung', 'P = U \\cdot I = I^{2} \\cdot R'], ['Kraft', 'F = m \\cdot a'],
  ['Mitternachtsformel', 'x_{1,2} = \\frac{-b \\pm \\sqrt{b^{2} - 4ac}}{2a}'], ['Pythagoras', 'c = \\sqrt{a^{2} + b^{2}}'],
  ['Mittelwert', '\\overline{x} = \\frac{1}{n} \\sum_{i=1}^{n} x_{i}'], ['PT1-Glied', 'G(s) = \\frac{K}{1 + T \\cdot s}'],
  ['Kreisfrequenz', '\\omega = 2 \\pi \\cdot f']
];
function texIcon(tex){
  const s = document.createElement('span');
  try { temml.render(tex, s, {throwOnError:false}); } catch(e){ s.textContent = tex; }
  return s.innerHTML;
}
C.equation = (span) => {
  if (!window.temml){ SW.toast('Formel-Bibliothek nicht gefunden (lib/temml).'); return; }
  SW.rememberRange();
  const editing = span && span.classList && span.classList.contains('sw-math') ? span : null;
  const tex0 = editing ? editing.dataset.tex : '';
  const disp0 = editing ? !!editing.closest('.sw-mathblock') : true;
  const pal = PALETTE.map(([name, items]) => `<div class="eq-group"><div class="eq-gname">${name}</div><div class="eq-btns">${items.map(([t, tip, ins]) =>
    `<button class="eq-btn" data-ins="${esc(ins || t)}" title="${esc(tip || t)}">${texIcon(t)}</button>`).join('')}</div></div>`).join('');
  const dlg = SW.dialog({
    title:editing ? 'Formel bearbeiten' : 'Formel einfügen', width:680,
    body:`<div class="eq-pal">${pal}</div>
      <div class="dlg-row" style="align-items:flex-start"><textarea id="eqTex" class="eq-tex" spellcheck="false" placeholder="z. B. U = R \\cdot I   oder   \\frac{a}{b}">${esc(tex0)}</textarea></div>
      <div class="eq-prev" id="eqPrev"></div>
      <div class="dlg-row" style="gap:14px;flex-wrap:wrap"><label class="chk" style="min-width:0"><input type="checkbox" id="eqDisp"${disp0 ? ' checked' : ''}> Als eigene Zeile (zentriert)</label>
        <select id="eqEx" style="flex:0 0 200px"><option value="">Beispiel einfügen …</option>${EXAMPLES.map(([n, t], i) => `<option value="${i}">${n}</option>`).join('')}</select></div>
      <p class="muted" style="margin:4px 0 0">Schreibweise wie LaTeX: <b>^</b> hoch, <b>_</b> tief, <b>{ }</b> gruppiert, <b>\\frac{}{}</b> Bruch, <b>Tab</b> springt zum nächsten □. Die Formel wird in Word als echte, bearbeitbare Formel gespeichert.</p>`,
    buttons:[
      ...(editing ? [{label:'Formel löschen', action:() => { SW.edit(() => { const p = editing.closest('.sw-mathblock'); if (p && p.querySelectorAll('.sw-math').length === 1 && !p.textContent.replace(editing.textContent, '').trim()){ p.innerHTML = '<br>'; p.classList.remove('sw-mathblock'); if (!p.className) p.removeAttribute('class'); } else editing.remove(); }); }}] : []),
      {label:'Abbrechen'},
      {label:editing ? 'Übernehmen' : 'Einfügen', primary:true, action:d => {
        const tex = d.querySelector('#eqTex').value.trim();
        if (!tex) return false;
        insertOrUpdate(editing, tex, d.querySelector('#eqDisp').checked);
      }}
    ]
  });
  const ta = dlg.el.querySelector('#eqTex'), prev = dlg.el.querySelector('#eqPrev');
  const upd = () => {
    prev.innerHTML = '';
    const v = ta.value.trim();
    if (!v){ prev.innerHTML = '<span class="muted">Vorschau</span>'; return; }
    try { temml.render(v, prev, {displayMode:true, throwOnError:true}); prev.classList.remove('err'); }
    catch(e){ prev.textContent = 'Noch unvollständig: ' + String(e.message || e).replace(/^.*?Error:\s*/, ''); prev.classList.add('err'); }
  };
  ta.addEventListener('input', upd);
  ta.addEventListener('keydown', e => {
    if (e.key === 'Enter' && !e.shiftKey){ e.preventDefault(); dlg.el.querySelector('.dlg-foot .btn.primary').click(); }
    if (e.key === 'Tab' && !e.shiftKey){
      // zum nächsten Platzhalter □ springen
      let i = ta.value.indexOf('□', ta.selectionEnd);
      if (i < 0) i = ta.value.indexOf('□');
      if (i >= 0){ e.preventDefault(); ta.setSelectionRange(i, i + 1); }
    }
    e.stopPropagation();
  });
  dlg.el.querySelector('.eq-pal').addEventListener('mousedown', e => { if (e.target.closest('.eq-btn')) e.preventDefault(); });
  dlg.el.querySelector('.eq-pal').addEventListener('click', e => {
    const b = e.target.closest('.eq-btn'); if (!b) return;
    insertSnippet(ta, b.dataset.ins);
    upd();
  });
  dlg.el.querySelector('#eqEx').addEventListener('change', e => {
    const ex = EXAMPLES[+e.target.value]; if (!ex) return;
    ta.value = ex[1]; e.target.value = ''; ta.focus(); upd();
  });
  setTimeout(() => { ta.focus(); ta.setSelectionRange(ta.value.length, ta.value.length); }, 0);
  upd();
};
/* Baustein an Cursor einfügen, erstes □ markieren */
function insertSnippet(ta, ins){
  const a = ta.selectionStart, b = ta.selectionEnd;
  const selText = ta.value.slice(a, b);
  let txt = ins;
  if (selText && txt.includes('□')) txt = txt.replace('□', selText);
  const needSpace = /^\\[a-zA-Z]+$/.test(txt) && /[a-zA-Z]$/.test(ta.value.slice(0, a));
  txt = (needSpace ? ' ' : '') + txt;
  ta.value = ta.value.slice(0, a) + txt + ta.value.slice(b);
  ta.focus();
  const ph = ta.value.indexOf('□', a);
  if (ph >= 0 && ph < a + txt.length){ ta.setSelectionRange(ph, ph + 1); }
  else { const p = a + txt.length; ta.setSelectionRange(p, p); }
  // Platzhalter, die noch übrig sind, beim Rendern als Leerstelle zeigen
}
function insertOrUpdate(editing, tex, display){
  tex = tex.replace(/□/g, '{}');
  SW.edit(() => {
    let span = editing;
    if (!span){
      span = document.createElement('span');
      span.className = 'sw-math';
      span.setAttribute('contenteditable', 'false');
      span.dataset.tex = tex;
      if (display){
        const p = document.createElement('p');
        p.className = 'sw-mathblock';
        p.appendChild(span);
        SW.insertBlockAtCaret(p);
      } else {
        SW.focusEditor();
        const r = SW.getRange();
        r.deleteContents();
        r.insertNode(span);
        const after = document.createTextNode('\u200B');
        span.after(after);
        placeCaretAfter(after);
      }
    } else {
      span.dataset.tex = tex;
      const blk = span.closest('.sw-mathblock');
      if (display && !blk){
        const own = SW.innermostBlock(span);
        if (own && own.matches('p') && !own.textContent.replace(span.textContent, '').replace(/\u200B/g, '').trim()) own.classList.add('sw-mathblock');
        else { const p = document.createElement('p'); p.className = 'sw-mathblock'; (own || span).after(p); p.appendChild(span); }
      } else if (!display && blk){
        blk.classList.remove('sw-mathblock');
        if (!blk.className) blk.removeAttribute('class');
      }
    }
    SW.renderMath(span);
    span.dataset.r = tex;
  });
}
ed.addEventListener('click', e => {
  const m = e.target.closest && e.target.closest('.sw-math');
  if (m && ed.contains(m)){ e.preventDefault(); C.equation(m); return; }
  const row = e.target.closest && e.target.closest('.sw-tof .toc-row[data-cap]');
  if (row){
    const type = row.closest('.sw-tof').dataset.seq || 'Abbildung';
    const ent = SW.tofEntries(type)[+row.dataset.cap];
    if (ent) SW.scrollToEl(ent.h);
  }
  const al = e.target.closest && e.target.closest('.sw-abbrlist');
  if (al) C.abbrManage();
  const cm = e.target.closest && e.target.closest('.sw-cmt[data-cid]');
  if (cm && showComments){ setActive(cm.dataset.cid); }
});

/* ============================================================
   Anbindung: Tastatur, Zustandsanzeige, Layout
   ============================================================ */
document.addEventListener('keydown', e => {
  if (e.target.closest && e.target.closest('.dialog')) return;
  const mod = e.ctrlKey || e.metaKey;
  if (mod && e.altKey && e.key.toLowerCase() === 'k'){ e.preventDefault(); e.stopImmediatePropagation(); C.commentNew(); return; }
  if (e.altKey && e.shiftKey && !mod && (e.key === '=' || e.code === 'Digit0')){ e.preventDefault(); e.stopImmediatePropagation(); C.equation(); return; }
  // Formel mit Entf/Rücktaste als Ganzes löschen
  if ((e.key === 'Backspace' || e.key === 'Delete') && document.activeElement === ed){
    const r = SW.getRange();
    if (r && r.collapsed){
      let n = null;
      const c = r.startContainer, o = r.startOffset;
      if (e.key === 'Backspace'){
        if (c.nodeType === 3 && (o === 0 || (o === 1 && c.data[0] === '\u200B'))) n = c.previousSibling;
        else if (c.nodeType === 1) n = c.childNodes[o - 1];
      } else {
        if (c.nodeType === 3 && o >= c.length) n = c.nextSibling;
        else if (c.nodeType === 1) n = c.childNodes[o];
      }
      if (n && n.nodeType === 1 && n.classList.contains('sw-math')){ e.preventDefault(); e.stopImmediatePropagation(); SW.edit(() => n.remove()); }
    }
  }
}, true);

const prevAfter = SW.afterLayout;
SW.afterLayout = function(){
  prevAfter && prevAfter();
  renderComments();
};
let frT = 0;
function featureRefresh(){
  cancelAnimationFrame(frT);
  frT = requestAnimationFrame(() => {
    $('#chkTitlePage').checked = !!SW.doc.titlePage;
    const cid = cidAtCaret();
    if (cid && cid !== activeCid) setActive(cid);
    const r = SW.getRange();
    const cols = r && closestIn(r.startContainer, '.sw-cols');
    const cb = $('[data-menu="colsMenu"]');
    if (cb) cb.classList.toggle('on', !!cols);
    const img = SW.selectedImage();
    const wb = $('[data-menu="wrapMenu"]');
    $$('#wrapMenu [data-arg]').forEach(b => b.classList.toggle('checked', !!img && (img.dataset.wrap || 'inline') === b.dataset.arg));
    if (wb) wb.classList.toggle('on', !!img && !!img.dataset.wrap);
  });
}
const prevRefresh = SW.refreshUI;
SW.refreshUI = function(){ prevRefresh && prevRefresh(); featureRefresh(); };
document.addEventListener('selectionchange', () => { if (SW.getRange()) featureRefresh(); });
// Ältere Teile der App rufen SW.refreshUI direkt auf – erneut zeichnen
SW.renderAllMath();
SW.relayout();
})();
