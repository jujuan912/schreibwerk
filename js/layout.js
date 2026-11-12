/* ==========================================================
   Schreibwerk – Seitenlayout: Seiten, Umbrüche, Kopf-/Fußzeilen,
   Fußnoten, Inhaltsverzeichnis, Lineal, Zoom
   ========================================================== */
(function(){
'use strict';
const SW = window.SW, ed = SW.ed, CM = SW.CM, PT = SW.PT;
const paper = document.getElementById('paper');
const pagesEl = document.getElementById('pages');
const zoomBox = document.getElementById('zoomBox');
const canvas = document.getElementById('canvas');
let GAP = 24;
Object.defineProperty(SW, '_printGap', {set(v){ GAP = v ? 0 : 24; }, get(){ return GAP === 0; }});

/* ---------- Geometrie ---------- */
SW.geo = function(){
  let [w, h] = SW.PAGE_SIZES[SW.doc.pageSize] || SW.PAGE_SIZES.A4;
  if (SW.doc.orientation === 'landscape') [w, h] = [h, w];
  const m = SW.doc.margins;
  return {
    W: w*CM, H: h*CM, wCm:w, hCm:h,
    mt: m.top*CM, mb: m.bottom*CM, ml: m.left*CM, mr: m.right*CM,
    cw: (w - m.left - m.right)*CM,
    hdr: 1.25*CM, ftr: 1.25*CM
  };
};

SW.applyBase = function(){
  const b = SW.doc.base;
  ed.style.fontFamily = `"${b.font}", Calibri, Carlito, "Segoe UI", Arial, sans-serif`;
  ed.style.fontSize = b.size + 'pt';
  ed.style.lineHeight = (b.line * 1.22).toFixed(3);
  ed.style.setProperty('--sw-after', b.after + 'pt');
  let st = document.getElementById('swBaseStyle');
  if (!st){ st = document.createElement('style'); st.id = 'swBaseStyle'; document.head.appendChild(st); }
  st.textContent = `#editor p:not([class]),#editor p.sw-box{margin-bottom:${b.after}pt}#editor ul,#editor ol,#editor table{margin-bottom:${b.after}pt}#editor :is(ul,ol,li) :is(ul,ol){margin-bottom:0}#editor td > :is(p,ul,ol):last-child,#editor th > :is(p,ul,ol):last-child{margin-bottom:0}`;
};

function applyGeometry(g){
  paper.style.width = g.W + 'px';
  ed.style.left = g.ml + 'px';
  ed.style.top = g.mt + 'px';
  ed.style.width = g.cw + 'px';
}

/* ---------- Layout-Artefakte entfernen ---------- */
/* Verschiebungen werden über ein eigenes Stylesheet gesetzt – nie im Inhalt */
const pushSheet = document.createElement('style');
pushSheet.id = 'swPushStyle';
document.head.appendChild(pushSheet);
let pushRules = {}, pushSeq = 0;
function setPush(el, px){
  let id = el.getAttribute('data-swp');
  if (!id){ id = String(++pushSeq); el.setAttribute('data-swp', id); }
  pushRules[id] = px;
  pushSheet.textContent = Object.entries(pushRules).map(([k, v]) => `#editor [data-swp="${k}"]{margin-top:${v.toFixed(2)}px !important}`).join('');
}
function clearArtifacts(){
  ed.querySelectorAll('.sw-gap').forEach(g => {
    const p = g.parentNode; g.remove();
    p && p.normalize();
  });
  ed.querySelectorAll('tr.sw-gaprow').forEach(g => g.remove());
  ed.querySelectorAll('[data-swp]').forEach(el => el.removeAttribute('data-swp'));
  pushRules = {};
  pushSheet.textContent = '';
}
SW.clearArtifacts = clearArtifacts;

/* ---------- Einheiten (Layout-Bausteine) ---------- */
function units(){
  const out = [];
  const addList = (list) => {
    for (const li of list.children){
      if (li.tagName === 'LI') out.push({el:li, kind:'block', split:isSplittable(li)});
      else if (li.tagName === 'UL' || li.tagName === 'OL') addList(li);
    }
  };
  for (const el of ed.children){
    if (el.classList.contains('sw-pagebreak')) { out.push({el, kind:'break'}); continue; }
    if (el.tagName === 'UL' || el.tagName === 'OL'){ addList(el); continue; }
    if (el.tagName === 'TABLE'){
      const rows = el.rows;
      for (let i = 0; i < rows.length; i++){
        if (rows[i].classList.contains('sw-gaprow')) continue;
        out.push({el:rows[i], kind:'row', table:el, first:i === 0});
      }
      continue;
    }
    out.push({el, kind:'block', split:isSplittable(el), heading:/^H[1-3]$/.test(el.tagName) || el.classList.contains('sw-title')});
  }
  return out;
}
function isSplittable(el){
  if (!el.matches('p,h1,h2,h3,h4,h5,h6,li,blockquote,pre')) return false;
  if (el.querySelector('img,table,ul,ol')) return false;
  return el.textContent.length > 0;
}

/* ---------- Messen ---------- */
let paperTop = 0, Z = 1;
function rectOf(el){
  const r = el.getBoundingClientRect();
  return {top:(r.top - paperTop)/Z, bottom:(r.bottom - paperTop)/Z, height:r.height/Z};
}

/* Vergrößert den oberen Abstand eines Elements, bis es bei target beginnt */
function pushTo(u, target){
  let el = u.el;
  if (u.kind === 'row'){
    if (u.first){ el = u.table; }
    else {
      const top = rectOf(el).top;
      const gr = document.createElement('tr');
      gr.className = 'sw-gaprow';
      gr.setAttribute('contenteditable', 'false');
      const td = document.createElement('td');
      td.colSpan = 99;
      td.style.height = Math.max(0, target - top) + 'px';
      gr.appendChild(td);
      el.before(gr);
      const now = rectOf(el).top;
      if (Math.abs(now - target) > 0.5) td.style.height = Math.max(0, parseFloat(td.style.height) + (target - now)) + 'px';
      return;
    }
  }
  for (let i = 0; i < 4; i++){
    const top = rectOf(el).top;
    const diff = target - top;
    if (Math.abs(diff) < 0.5) break;
    const cur = parseFloat(getComputedStyle(el).marginTop) || 0;
    setPush(el, Math.max(0, cur + diff));
  }
}

/* Absatz zeilengenau über Seitengrenze teilen. Liefert Seitenindex des Endes */
function splitBlock(el, page, P){
  const texts = [];
  const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  let t, total = 0;
  while ((t = w.nextNode())){
    if (t.parentNode.closest('.sw-fn,.sw-gap')) continue;
    if (!t.length) continue;
    texts.push({t, s:total}); total += t.length;
  }
  if (!total) return null;
  const rng = document.createRange();
  const loc = i => { let lo = 0, hi = texts.length - 1; while (lo < hi){ const m = (lo + hi + 1) >> 1; if (texts[m].s <= i) lo = m; else hi = m - 1; } return [texts[lo].t, i - texts[lo].s]; };
  const charRect = i => {
    for (let k = i; k < total; k++){
      const [n, o] = loc(k);
      rng.setStart(n, o); rng.setEnd(n, o + 1);
      const rs = rng.getClientRects();
      if (rs.length){ const r = rs[rs.length - 1]; if (r.height > 0) return {top:(r.top - paperTop)/Z, bottom:(r.bottom - paperTop)/Z, i:k}; }
    }
    return null;
  };
  const lh = parseFloat(getComputedStyle(el).lineHeight) || 16;
  let guard = 0;
  while (guard++ < 200){
    const limit = P.bot(page);
    const elR = rectOf(el);
    if (elR.bottom <= limit + 0.5) return page;
    // erstes Zeichen, dessen Unterkante über die Grenze ragt
    let lo = 0, hi = total - 1, found = -1;
    while (lo <= hi){
      const mid = (lo + hi) >> 1;
      const r = charRect(mid);
      if (!r) { hi = mid - 1; continue; }
      if (r.bottom > limit + 0.5){ found = r.i; hi = mid - 1; } else lo = mid + 1;
    }
    if (found < 0) return page;
    let fr = charRect(found);
    // Zeilenanfang finden (gleiche Oberkante)
    let k = found;
    while (k > 0){ const pr = charRect(k - 1); if (!pr || Math.abs(pr.top - fr.top) > 2) break; k--; }
    // Schusterjungen/Hurenkinder: min. 2 Zeilen pro Seite
    const contentTop = elR.top + (parseFloat(getComputedStyle(el).paddingTop) || 0);
    const linesBefore = Math.round((fr.top - contentTop) / lh);
    const lastR = charRect(total - 1);
    const linesAfter = lastR ? Math.round((lastR.bottom - fr.top) / lh) : 1;
    if (linesBefore < 2){
      if (elR.top <= P.top(page) + 1) return page;    // passt nirgends besser
      pushTo({el, kind:'block'}, P.top(page + 1));
      page++;
      continue;
    }
    if (linesAfter === 1 && linesBefore >= 3){
      // eine Zeile mehr auf die nächste Seite
      let j = k - 1; const prev = charRect(j);
      if (prev){ while (j > 0){ const pr = charRect(j - 1); if (!pr || Math.abs(pr.top - prev.top) > 2) break; j--; } k = j; fr = charRect(k) || fr; }
    }
    // Abstandshalter einfügen
    const [n, o] = loc(k);
    let at = n;
    if (o > 0) at = n.splitText(o);
    const gap = document.createElement('span');
    gap.className = 'sw-gap';
    gap.setAttribute('contenteditable', 'false');
    const half = (lh - (fr.bottom - fr.top)) / 2;
    const lineTop = fr.top - Math.max(0, half);
    gap.style.height = Math.max(0, P.top(page + 1) - lineTop) + 'px';
    at.parentNode.insertBefore(gap, at);
    // Feinkorrektur
    const nr = document.createRange(); nr.setStart(at, 0); nr.setEnd(at, Math.min(1, at.length));
    const rs = nr.getClientRects();
    if (rs.length){
      const nt = (rs[0].top - paperTop)/Z - Math.max(0, half);
      const corr = P.top(page + 1) - nt;
      if (Math.abs(corr) > 0.5) gap.style.height = Math.max(0, parseFloat(gap.style.height) + corr) + 'px';
    }
    // Text-Knoten-Liste neu aufbauen (splitText)
    texts.length = 0; total = 0;
    const w2 = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    while ((t = w2.nextNode())){ if (t.parentNode.closest('.sw-fn,.sw-gap') || !t.length) continue; texts.push({t, s:total}); total += t.length; }
    page++;
  }
  return page;
}

/* ---------- Ein Paginierungsdurchlauf ---------- */
function paginate(g, fnH){
  const step = g.H + GAP;
  const P = {
    top: p => p*step + g.mt,
    bot: p => p*step + g.H - g.mb - (fnH[p] || 0)
  };
  const U = units();
  let page = 0, force = false;
  for (let i = 0; i < U.length; i++){
    const u = U[i];
    if (u.kind === 'break'){ force = true; continue; }
    let r = rectOf(u.el);
    while (r.top >= (page + 1)*step - 0.5) page++;
    if (force){
      force = false;
      if (r.top < P.top(page + 1) - 0.5){ pushTo(u, P.top(page + 1)); page++; r = rectOf(u.el); }
    }
    // Überschrift nicht allein am Seitenende ("Absätze nicht trennen")
    if (u.heading && i + 1 < U.length && U[i+1].kind !== 'break' && r.bottom <= P.bot(page)){
      const nr = rectOf(U[i+1].el);
      const firstLine = Math.min(nr.height, (parseFloat(getComputedStyle(U[i+1].el).lineHeight) || 16) * 2);
      if (nr.top + firstLine > P.bot(page) + 0.5 && r.top > P.top(page) + 1){
        pushTo(u, P.top(page + 1)); page++; r = rectOf(u.el);
      }
    }
    if (r.bottom > P.bot(page) + 0.5){
      if (u.split){
        const endPage = splitBlock(u.el, page, P);
        if (endPage !== null){ page = endPage; continue; }
      }
      if (r.top > P.top(page) + 1){
        pushTo(u, P.top(page + 1)); page++;
        r = rectOf(u.el);
      }
      while (r.bottom > (page + 1)*step) page++;
    }
  }
  // Seitenanzahl
  const last = ed.lastElementChild;
  let n = 1;
  if (last){
    const b = Math.max(rectOf(last).bottom, rectOf(ed).bottom - 1);
    n = Math.max(1, Math.floor((b - 1) / step) + 1);
    if (force) n++;
  }
  return n;
}

/* ---------- Fußnoten ---------- */
let fnMeasure = null;
function measureFootnotes(g){
  if (!fnMeasure){
    fnMeasure = document.createElement('div');
    fnMeasure.className = 'page';
    fnMeasure.style.cssText = 'position:absolute;visibility:hidden;left:-9999px;top:0;';
    document.body.appendChild(fnMeasure);
  }
  const refs = SW.$$('.sw-fn', ed);
  refs.forEach((r, i) => { r.textContent = String(i + 1); });
  if (!refs.length) return [];
  fnMeasure.innerHTML = '<div class="fn-area" style="position:static;width:' + g.cw + 'px"></div>';
  const area = fnMeasure.firstChild;
  return refs.map((r, i) => {
    area.innerHTML = `<p class="fn-line"><sup>${i + 1}</sup>${SW.esc(r.dataset.note || '')}</p>`;
    return {ref:r, h:area.firstChild.getBoundingClientRect().height + 2};
  });
}
function fnPages(fns, g){
  const step = g.H + GAP;
  const per = {};
  fns.forEach(f => {
    const p = Math.max(0, Math.floor(rectOf(f.ref).top / step));
    (per[p] = per[p] || []).push(f);
  });
  return per;
}

/* ---------- Inhaltsverzeichnis ---------- */
function tocEntries(g){
  const step = g.H + GAP;
  return SW.$$('h1,h2,h3', ed).filter(h => !h.closest('.sw-toc') && h.textContent.trim()).map(h => ({
    h, lvl:+h.tagName[1], text:h.textContent.replace(/​/g, '').trim(),
    page: Math.floor(rectOf(h).top / step) + 1
  }));
}
function renderTocs(g){
  const tocs = SW.$$('.sw-toc', ed);
  if (!tocs.length) return false;
  const ents = tocEntries(g);
  const html = '<div class="toc-head">Inhaltsverzeichnis</div>' + (ents.length
    ? ents.map((e, i) => `<div class="toc-row l${e.lvl}" data-i="${i}"><span class="toc-t">${SW.esc(e.text)}</span><span class="toc-dots"></span><span class="toc-p">${e.page}</span></div>`).join('')
    : '<div class="toc-empty">Keine Einträge gefunden – formatieren Sie Überschriften mit „Überschrift 1–3“.</div>');
  let changed = false;
  tocs.forEach(t => { if (t.innerHTML !== html){ t.innerHTML = html; changed = true; } });
  return changed;
}
SW.tocEntries = () => tocEntries(SW.geo());

/* ---------- Seitenhintergründe ---------- */
function hfText(tpl, i, n){
  return (tpl || '').replace(/\{SEITE\}/g, i + 1).replace(/\{SEITEN\}/g, n);
}
function renderPages(g, n, fnPer){
  const step = g.H + GAP;
  while (pagesEl.children.length < n){
    const d = document.createElement('div');
    d.className = 'page';
    d.innerHTML = '<div class="hf hf-h" data-ph="Kopfzeile – doppelklicken zum Bearbeiten"></div><div class="hf hf-f" data-ph="Fußzeile – doppelklicken zum Bearbeiten"></div><div class="fn-area" hidden></div>';
    pagesEl.appendChild(d);
  }
  while (pagesEl.children.length > n) pagesEl.lastChild.remove();
  const D = SW.doc;
  [...pagesEl.children].forEach((d, i) => {
    d.style.top = (i*step) + 'px';
    d.style.height = g.H + 'px';
    const h = d.children[0], f = d.children[1], fa = d.children[2];
    h.style.cssText = `left:${g.ml}px;width:${g.cw}px;top:${g.hdr}px;text-align:${D.header.align}`;
    f.style.cssText = `left:${g.ml}px;width:${g.cw}px;bottom:${g.ftr}px;text-align:${D.footer.align}`;
    h.textContent = hfText(D.header.text, i, n);
    f.textContent = hfText(D.footer.text, i, n);
    h.classList.toggle('empty', !D.header.text);
    f.classList.toggle('empty', !D.footer.text);
    const fns = fnPer[i] || [];
    if (fns.length){
      fa.hidden = false;
      const h2 = fns.reduce((s, x) => s + x.h, 0) + 6;
      fa.style.cssText = `left:${g.ml}px;width:${g.cw}px;top:${g.H - g.mb - h2}px`;
      fa.innerHTML = fns.map(x => `<p class="fn-line" data-fn="${SW.$$('.sw-fn', ed).indexOf(x.ref)}"><sup>${x.ref.textContent}</sup>${SW.esc(x.ref.dataset.note || '')}</p>`).join('');
    } else { fa.hidden = true; fa.innerHTML = ''; }
  });
  paper.style.height = (n*step - GAP) + 'px';
  SW.pageCount = n;
}

/* ---------- Hauptfunktion ---------- */
let busy = false;
SW.relayout = function(keepSel){
  if (busy || SW._composing) return;
  busy = true;
  const g = SW.geo();
  applyGeometry(g);
  const hadFocus = document.activeElement === ed;
  const sel = hadFocus || keepSel ? SW.saveSel() : null;
  const st = canvas.scrollTop;
  Z = SW.zoom;
  paperTop = paper.getBoundingClientRect().top;

  clearArtifacts();
  let fns = measureFootnotes(g);
  let fnH = [];
  let n = paginate(g, fnH);
  if (fns.length){
    const per = fnPages(fns, g);
    fnH = [];
    Object.keys(per).forEach(p => fnH[p] = per[p].reduce((s, x) => s + x.h, 0) + 10);
    clearArtifacts();
    n = paginate(g, fnH);
  }
  if (renderTocs(g)){
    clearArtifacts();
    n = paginate(g, fnH);
    renderTocs(g);
  }
  renderPages(g, n, fns.length ? fnPages(fns, g) : {});
  updateZoomBox(g, n);
  if (sel && hadFocus) SW.restoreSel(sel, false);
  canvas.scrollTop = st;
  busy = false;
  SW.renderRuler();
  SW.afterLayout && SW.afterLayout();
};

let layoutT = null;
SW.requestLayout = function(delay=60){
  clearTimeout(layoutT);
  layoutT = setTimeout(() => SW.relayout(), delay);
};

/* ---------- Zoom ---------- */
function updateZoomBox(g, n){
  const step = g.H + GAP;
  zoomBox.style.width = (g.W * SW.zoom) + 'px';
  zoomBox.style.height = ((n || SW.pageCount || 1) * step - GAP) * SW.zoom + 'px';
  paper.style.transform = `scale(${SW.zoom})`;
}
SW.setZoom = function(z){
  z = Math.max(0.25, Math.min(3, Math.round(z * 100) / 100));
  const before = canvas.scrollTop / Math.max(1, canvas.scrollHeight);
  SW.zoom = z;
  updateZoomBox(SW.geo());
  SW.relayout(true);
  canvas.scrollTop = before * canvas.scrollHeight;
  document.getElementById('zoomSlider').value = Math.round(z * 100);
  document.getElementById('zoomPct').textContent = Math.round(z * 100) + ' %';
  SW.positionHandles && SW.positionHandles();
};
SW.zoomFit = function(mode){
  const g = SW.geo();
  const availW = canvas.clientWidth - 48, availH = canvas.clientHeight - 60;
  SW.setZoom(mode === 'page' ? Math.min(availW / g.W, availH / g.H) : availW / g.W);
};

/* ---------- Seite des Cursors ---------- */
SW.caretPage = function(){
  const r = SW.getRange();
  if (!r) return 1;
  let rect = r.getClientRects()[0];
  if (!rect){
    const n = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentNode;
    rect = n.getBoundingClientRect();
  }
  const g = SW.geo();
  const y = (rect.top - paper.getBoundingClientRect().top) / SW.zoom;
  return Math.max(1, Math.min(SW.pageCount || 1, Math.floor(y / (g.H + GAP)) + 1));
};
SW.scrollToEl = function(el){
  const r = el.getBoundingClientRect(), c = canvas.getBoundingClientRect();
  canvas.scrollTop += r.top - c.top - 80;
};
SW.scrollToPage = function(p){
  const g = SW.geo();
  canvas.scrollTop = ((p - 1) * (g.H + GAP)) * SW.zoom;
};

/* ---------- Lineal ---------- */
const ruler = document.getElementById('ruler');
SW.renderRuler = function(){
  if (document.body.classList.contains('no-ruler')) return;
  const g = SW.geo(), z = SW.zoom;
  ruler.style.width = (g.W * z) + 'px';
  const blk = SW.selectedBlocks(true)[0];
  let indL = 0, indR = 0, first = 0;
  if (blk){
    const cs = getComputedStyle(blk);
    const edR = ed.getBoundingClientRect(), bR = blk.getBoundingClientRect();
    indL = (bR.left - edR.left) / z + parseFloat(cs.paddingLeft || 0);
    indR = (edR.right - bR.right) / z;
    first = parseFloat(cs.textIndent) || 0;
  }
  let h = `<div class="r-white" style="left:${g.ml*z}px;width:${g.cw*z}px"></div>`;
  const zero = g.ml;
  for (let cm = -Math.floor(g.ml / CM * 4) / 4; zero + cm*CM <= g.W + 0.1; cm += 0.25){
    const x = (zero + cm*CM) * z;
    if (x < 0) continue;
    const whole = Math.abs(cm - Math.round(cm)) < 0.01;
    if (whole && Math.round(cm) !== 0) h += `<span class="r-num" style="left:${x}px">${Math.abs(Math.round(cm))}</span>`;
    else if (!whole) h += `<span class="r-tick${Math.abs(cm*2 - Math.round(cm*2)) < 0.01 ? ' half' : ''}" style="left:${x}px"></span>`;
  }
  h += `<span class="r-handle" data-m="left" style="left:${g.ml*z}px" title="Linker Seitenrand"></span>`;
  h += `<span class="r-handle" data-m="right" style="left:${(g.W - g.mr)*z}px" title="Rechter Seitenrand"></span>`;
  h += `<span class="r-ind" data-i="left" style="left:${(g.ml + indL)*z}px" title="Linker Einzug"></span>`;
  h += `<span class="r-ind first" data-i="first" style="left:${(g.ml + indL + first)*z}px" title="Erstzeileneinzug"></span>`;
  h += `<span class="r-ind" data-i="right" style="left:${(g.W - g.mr - indR)*z}px" title="Rechter Einzug"></span>`;
  ruler.innerHTML = h;
};
ruler.addEventListener('mousedown', e => {
  const t = e.target;
  if (!t.dataset.m && !t.dataset.i) return;
  e.preventDefault();
  const g = SW.geo(), z = SW.zoom;
  const rl = ruler.getBoundingClientRect().left;
  const blocks = t.dataset.i ? SW.selectedBlocks() : [];
  if (t.dataset.i) SW.flushCommit();
  const move = ev => {
    let xCm = Math.round(((ev.clientX - rl) / z) / CM * 4) / 4;   // Raster 0,25 cm
    xCm = Math.max(0, Math.min(g.wCm, xCm));
    if (t.dataset.m === 'left') SW.doc.margins.left = Math.min(xCm, g.wCm - SW.doc.margins.right - 2);
    else if (t.dataset.m === 'right') SW.doc.margins.right = Math.max(0, Math.min(g.wCm - xCm, g.wCm - SW.doc.margins.left - 2));
    else {
      const rel = xCm - SW.doc.margins.left;
      blocks.forEach(b => {
        if (t.dataset.i === 'left'){
          const ti = parseFloat(b.style.textIndent) || 0;
          b.style.marginLeft = rel > 0 ? rel.toFixed(2) + 'cm' : '';
          if (ti) b.style.textIndent = ti + 'cm';
        } else if (t.dataset.i === 'first'){
          const ml = b.style.marginLeft.endsWith('cm') ? parseFloat(b.style.marginLeft) : 0;
          const v = rel - ml;
          b.style.textIndent = Math.abs(v) > 0.01 ? v.toFixed(2) + 'cm' : '';
        } else {
          const right = g.wCm - SW.doc.margins.right - xCm;
          b.style.marginRight = right > 0 ? right.toFixed(2) + 'cm' : '';
        }
        if (!b.getAttribute('style')) b.removeAttribute('style');
      });
    }
    SW.relayout(true);
  };
  const up = () => {
    removeEventListener('mousemove', move); removeEventListener('mouseup', up);
    SW.commit(); SW.markDirty(); SW.refreshUI && SW.refreshUI();
  };
  addEventListener('mousemove', move); addEventListener('mouseup', up);
});

/* Doppelklick auf Kopf-/Fußzeile, Klick auf Fußnoten */
paper.addEventListener('dblclick', e => {
  const hf = e.target.closest('.hf');
  if (hf) SW.cmd('headerFooter', hf.classList.contains('hf-h') ? 'header' : 'footer');
});
pagesEl.addEventListener('click', e => {
  const l = e.target.closest('.fn-line');
  if (l) SW.editFootnote(SW.$$('.sw-fn', ed)[+l.dataset.fn]);
});

window.addEventListener('resize', () => SW.requestLayout(120));
if (document.fonts) document.fonts.ready.then(() => SW.requestLayout(10));
})();
