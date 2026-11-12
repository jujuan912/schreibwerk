/* ==========================================================
   Schreibwerk – Oberfläche: Menüband, Dialoge, Dateien, Tastatur
   ========================================================== */
(function(){
'use strict';
const SW = window.SW, ed = SW.ed, $ = SW.$, $$ = SW.$$;
const VERSION = '1.0.0';
const canvas = $('#canvas');

/* ---------------- Icons ---------------- */
SW.icon = function(name, cls='i'){
  const p = (window.SW_ICONS || {})[name];
  return p ? `<svg class="${cls}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${p}</svg>` : '';
};
function paintIcons(root=document){
  root.querySelectorAll('[data-icon]').forEach(el => {
    if (el.querySelector(':scope > svg.i')) return;
    el.insertAdjacentHTML('afterbegin', SW.icon(el.dataset.icon));
  });
}

/* ---------------- Schriften & Größen ---------------- */
const FONTS = ['Calibri','Calibri Light','Aptos','Arial','Arial Black','Bahnschrift','Cambria','Candara','Century Gothic','Comic Sans MS','Consolas','Constantia','Corbel','Courier New','Franklin Gothic Medium','Garamond','Georgia','Impact','Lucida Console','Palatino Linotype','Segoe UI','Tahoma','Times New Roman','Trebuchet MS','Verdana'];
$('#fontList').innerHTML = FONTS.map(f => `<option value="${f}">`).join('');
$('#sizeList').innerHTML = SW.FONT_SIZES.map(s => `<option value="${String(s).replace('.', ',')}">`).join('');

/* ---------------- Farben ---------------- */
const THEME = [
  ['FFFFFF','000000','E7E6E6','44546A','4472C4','ED7D31','A5A5A5','FFC000','5B9BD5','70AD47'],
  ['F2F2F2','7F7F7F','D0CECE','D6DCE4','D9E2F3','FBE5D5','EDEDED','FFF2CC','DEEBF6','E2EFD9'],
  ['D8D8D8','595959','AEABAB','ADB9CA','B4C6E7','F7CBAC','DBDBDB','FEE599','BDD7EE','C5E0B3'],
  ['BFBFBF','3F3F3F','757070','8496B0','8EAADB','F4B183','C9C9C9','FFD966','9DC3E6','A8D08D'],
  ['A5A5A5','262626','3A3838','323F4F','2F5496','C55A11','7B7B7B','BF8F00','2E75B5','538135'],
  ['C00000','FF0000','FFC000','FFFF00','92D050','00B050','00B0F0','0070C0','002060','7030A0']
];
const HIGHLIGHTS = ['FFFF00','00FF00','00FFFF','FF00FF','0000FF','FF0000','000080','008080','008000','800080','800000','808000','808080','C0C0C0','000000'];
const colorState = {foreColor:'#C00000', highlight:'#FFFF00', shading:'#DBE5F1', cellShade:'#D9E2F3'};
function buildColorMenu(menu){
  const tgt = menu.dataset.target;
  let h = '';
  if (tgt === 'highlight'){
    h += '<div class="menu-title">Hervorhebung</div><div class="color-grid">' + HIGHLIGHTS.map(c => `<button data-color="#${c}" style="background:#${c}" title="#${c}"></button>`).join('') + '</div>';
    h += '<div class="menu-sep"></div><button class="menu-btn" data-color="none">Keine Farbe</button>';
  } else {
    h += `<button class="menu-btn" data-color="${tgt === 'foreColor' ? 'auto' : 'none'}">${tgt === 'foreColor' ? 'Automatisch' : 'Keine Farbe'}</button>`;
    h += '<div class="menu-title">Designfarben</div><div class="color-grid">' + THEME.slice(0, 5).flat().map(c => `<button data-color="#${c}" style="background:#${c}" title="#${c}"></button>`).join('') + '</div>';
    h += '<div class="menu-title">Standardfarben</div><div class="color-grid">' + THEME[5].map(c => `<button data-color="#${c}" style="background:#${c}" title="#${c}"></button>`).join('') + '</div>';
    h += '<div class="menu-sep"></div><label class="menu-btn">Weitere Farben … <input type="color" value="#2f5496"></label>';
  }
  menu.innerHTML = h;
  menu.addEventListener('click', e => {
    const b = e.target.closest('[data-color]');
    if (!b) return;
    applyColor(tgt, b.dataset.color);
    closeMenus();
  });
  const inp = menu.querySelector('input[type=color]');
  if (inp) inp.addEventListener('change', () => { applyColor(tgt, inp.value); closeMenus(); });
}
function applyColor(tgt, c){
  const swatch = {foreColor:'#fcSwatch', highlight:'#hlSwatch', shading:'#shSwatch', cellShade:'#cellSwatch'}[tgt];
  if (c !== 'none' && c !== 'auto'){ colorState[tgt] = c; $(swatch).style.background = c; }
  SW.cmd(tgt, c);
}
$$('.color-menu').forEach(buildColorMenu);

/* ---------------- Symbole ---------------- */
const SYMBOLS = '€ £ $ ¥ © ® ™ § ¶ † ° ± × ÷ ≈ ≠ ≤ ≥ ∞ √ ∑ ∏ ∫ ∂ ∆ ∅ ∈ ∩ ∪ ⊥ ∥ ∠ ‰ ½ ¼ ¾ ² ³ ¹ µ Ω α β γ δ ε η θ λ π ρ σ τ φ ω Δ Φ Σ → ← ↑ ↓ ↔ ⇒ ⇔ • · … – — „ “ ” ‚ ‘ ’ « » ✓ ✗ ★ ☎ ✉ ⌀ ∢ Ø'.split(' ');
$('#symbolGrid').innerHTML = SYMBOLS.map(s => `<button data-sym="${s}" title="${s}">${s}</button>`).join('');
$('#symbolGrid').addEventListener('click', e => {
  const b = e.target.closest('[data-sym]');
  if (!b) return;
  closeMenus();
  SW.edit(() => SW.native('insertText', b.dataset.sym));
});

/* ---------------- Tabellen-Raster ---------------- */
const picker = $('#tblPicker');
picker.innerHTML = Array.from({length:80}, (_, i) => `<i data-r="${Math.floor(i/10)+1}" data-c="${i%10+1}"></i>`).join('');
picker.addEventListener('mouseover', e => {
  const t = e.target.closest('i'); if (!t) return;
  const r = +t.dataset.r, c = +t.dataset.c;
  picker.querySelectorAll('i').forEach(x => x.classList.toggle('on', +x.dataset.r <= r && +x.dataset.c <= c));
  $('#tblPickLabel').textContent = `Tabelle ${c} × ${r}`;
});
picker.addEventListener('click', e => {
  const t = e.target.closest('i'); if (!t) return;
  closeMenus();
  SW.insertTable(+t.dataset.r, +t.dataset.c);
});

/* ---------------- Formatvorlagen-Galerie ---------------- */
const STYLE_PREVIEW = {
  normal:'font:11pt Calibri,Carlito,sans-serif', nospace:'font:11pt Calibri,Carlito,sans-serif',
  h1:'font:14pt "Calibri Light",Calibri,sans-serif;color:#2f5496', h2:'font:12pt "Calibri Light",Calibri,sans-serif;color:#2f5496',
  h3:'font:11pt "Calibri Light",Calibri,sans-serif;color:#1f3763', title:'font:17pt "Calibri Light",Calibri,sans-serif',
  subtitle:'font:11pt Calibri,sans-serif;color:#5a5a5a;letter-spacing:.5pt', quote:'font:italic 11pt Calibri,sans-serif;color:#404040'
};
$('#styleGallery').innerHTML = SW.STYLES.map(s => `<button class="style-item" data-style="${s.id}" title="${s.name}"><span class="s-prev" style='${STYLE_PREVIEW[s.id]}'>AaBbCc</span><small>${s.name}</small></button>`).join('');
$('#styleGallery').addEventListener('click', e => {
  const b = e.target.closest('[data-style]');
  if (b) SW.applyStyle(b.dataset.style);
});

/* ---------------- Menüs ---------------- */
let openMenu = null, menuBtn = null;
function closeMenus(){
  $$('.menu.open').forEach(m => m.classList.remove('open'));
  openMenu = null; menuBtn = null;
}
SW.closeMenus = closeMenus;
function showMenu(id, anchor, at){
  const m = document.getElementById(id);
  if (openMenu === m && !at){ closeMenus(); return; }
  closeMenus();
  m.classList.add('open');
  const r = at || anchor.getBoundingClientRect();
  let left = at ? at.x : r.left, top = at ? at.y : r.bottom + 2;
  const mw = m.offsetWidth, mh = m.offsetHeight;
  if (left + mw > innerWidth - 6) left = innerWidth - mw - 6;
  if (top + mh > innerHeight - 6) top = Math.max(6, (at ? at.y : r.top) - mh - 2);
  m.style.left = left + 'px'; m.style.top = top + 'px';
  openMenu = m; menuBtn = anchor;
}
document.addEventListener('mousedown', e => {
  if (openMenu && !openMenu.contains(e.target) && !(menuBtn && menuBtn.contains(e.target))) closeMenus();
});

/* Ribbon & Menüs: Fokus im Editor behalten */
function keepFocus(e){
  const t = e.target;
  if (t.closest('input,select,textarea,label')) return;
  if (t.closest('button,.menu,.color-grid')) e.preventDefault();
}
['#ribbon', '#menus', '#titlebar .tb-left', '#statusbar'].forEach(s => $(s).addEventListener('mousedown', keepFocus));

document.addEventListener('click', e => {
  const b = e.target.closest('[data-cmd],[data-menu]');
  if (!b || b.disabled) return;
  if (b.closest('#dialogHost')) return;
  if (b.dataset.menu){ showMenu(b.dataset.menu, b); return; }
  const inMenu = b.closest('.menu');
  if (inMenu) closeMenus();
  SW.cmd(b.dataset.cmd, b.dataset.arg);
});

/* ---------------- Registerkarten ---------------- */
$('#tabs').addEventListener('click', e => {
  const t = e.target.closest('.tab');
  if (!t) return;
  if (t.dataset.tab === 'datei'){ openBackstage(); return; }
  selectTab(t.dataset.tab);
  if (document.body.classList.contains('ribbon-collapsed')) document.body.classList.remove('ribbon-collapsed');
});
function selectTab(name){
  $$('.tab').forEach(t => t.classList.toggle('active', t.dataset.tab === name));
  $$('.panel').forEach(p => p.classList.toggle('active', p.dataset.panel === name));
}
$('#ribbonToggle').addEventListener('click', () => { document.body.classList.toggle('ribbon-collapsed'); SW.requestLayout(); });

/* ---------------- Dialoge ---------------- */
SW.dialog = function({title, body, buttons=[], modal=true, onOpen, width}){
  const host = $('#dialogHost');
  const d = document.createElement('div');
  d.className = 'dialog' + (modal ? '' : ' nomodal');
  if (width) d.style.width = width + 'px';
  d.innerHTML = `<div class="dlg-head"><span>${title}</span><button class="dlg-x" title="Schließen">${SW.icon('x')}</button></div><div class="dlg-body"></div><div class="dlg-foot"></div>`;
  const bodyEl = d.querySelector('.dlg-body');
  if (typeof body === 'string') bodyEl.innerHTML = body; else bodyEl.appendChild(body);
  const foot = d.querySelector('.dlg-foot');
  const close = () => {
    d.remove();
    if (modal) $('#dlgBackdrop').hidden = !host.querySelector('.dialog:not(.nomodal)');
    SW.focusEditor();
  };
  buttons.forEach(b => {
    const btn = document.createElement('button');
    btn.className = 'btn' + (b.primary ? ' primary' : '');
    btn.textContent = b.label;
    btn.onclick = () => { if (b.action && b.action(d) === false) return; if (b.close !== false) close(); };
    foot.appendChild(btn);
  });
  if (!buttons.length) foot.remove();
  d.querySelector('.dlg-x').onclick = close;
  d.addEventListener('keydown', e => {
    if (e.key === 'Escape'){ e.stopPropagation(); close(); }
    if (e.key === 'Enter' && !e.shiftKey && e.target.tagName !== 'TEXTAREA' && e.target.tagName !== 'BUTTON'){
      const p = buttons.find(b => b.primary);
      if (p){ e.preventDefault(); if (p.action && p.action(d) === false) return; if (p.close !== false) close(); }
    }
  });
  // Verschieben
  const head = d.querySelector('.dlg-head');
  head.addEventListener('mousedown', e => {
    if (e.target.closest('button')) return;
    const r = d.getBoundingClientRect();
    const dx = e.clientX - r.left, dy = e.clientY - r.top;
    d.style.transform = 'none'; d.style.left = r.left + 'px'; d.style.top = r.top + 'px'; d.style.right = 'auto';
    const mv = ev => { d.style.left = (ev.clientX - dx) + 'px'; d.style.top = Math.max(0, ev.clientY - dy) + 'px'; };
    const up = () => { removeEventListener('mousemove', mv); removeEventListener('mouseup', up); };
    addEventListener('mousemove', mv); addEventListener('mouseup', up);
  });
  host.appendChild(d);
  if (modal) $('#dlgBackdrop').hidden = false;
  paintIcons(d);
  const first = d.querySelector('input:not([type=checkbox]),select,textarea');
  setTimeout(() => { if (first){ first.focus(); if (first.select) first.select(); } }, 0);
  onOpen && onOpen(d);
  return {el:d, close};
};
$('#dlgBackdrop').addEventListener('mousedown', e => e.preventDefault());

/* ---------------- Tabellen ---------------- */
SW.insertTable = function(rows, cols){
  const w = (100 / cols).toFixed(3);
  let h = '<colgroup>' + Array(cols).fill(`<col style="width:${w}%">`).join('') + '</colgroup><tbody>';
  for (let r = 0; r < rows; r++) h += '<tr>' + Array(cols).fill('<td><p><br></p></td>').join('') + '</tr>';
  const t = document.createElement('table');
  t.innerHTML = h + '</tbody>';
  SW.edit(() => SW.insertBlockAtCaret(t, {caretInside:true}));
};
function curCell(){
  const r = SW.getRange() || SW.lastRange;
  if (!r) return null;
  const n = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentNode;
  return n.closest ? n.closest('td,th') : null;
}
SW.curCell = curCell;
function tableGrid(tbl){
  const rows = Array.from(tbl.rows).filter(r => !r.classList.contains('sw-gaprow'));
  const grid = [];
  rows.forEach((tr, ri) => {
    grid[ri] = grid[ri] || [];
    let ci = 0;
    for (const cell of tr.cells){
      while (grid[ri][ci]) ci++;
      for (let r = 0; r < (cell.rowSpan || 1); r++)
        for (let c = 0; c < (cell.colSpan || 1); c++){
          grid[ri + r] = grid[ri + r] || [];
          grid[ri + r][ci + c] = {cell, r0:ri, c0:ci};
        }
      ci += cell.colSpan || 1;
    }
  });
  return {rows, grid, ncols:Math.max(...grid.map(g => g.length))};
}
function cellPos(cell){
  const tbl = cell.closest('table');
  const G = tableGrid(tbl);
  for (let r = 0; r < G.grid.length; r++)
    for (let c = 0; c < G.grid[r].length; c++)
      if (G.grid[r][c] && G.grid[r][c].cell === cell) return {tbl, G, r, c};
  return null;
}
function newCell(){ const td = document.createElement('td'); td.innerHTML = '<p><br></p>'; return td; }
function setCols(tbl, n, insertAt, del){
  let cg = tbl.querySelector(':scope > colgroup');
  if (!cg){ cg = document.createElement('colgroup'); tbl.prepend(cg); }
  const cols = Array.from(cg.children);
  let ws = cols.map(c => parseFloat(c.style.width) || 0);
  if (ws.length !== n - (insertAt != null ? 1 : 0) + (del != null ? 1 : 0)) ws = [];
  if (insertAt != null && ws.length){ const nw = ws[Math.max(0, Math.min(ws.length - 1, insertAt - 1))] || 100 / n; ws.splice(insertAt, 0, nw); }
  if (del != null && ws.length) ws.splice(del, 1);
  if (!ws.length) ws = Array(n).fill(100 / n);
  const sum = ws.reduce((a, b) => a + b, 0);
  cg.innerHTML = ws.map(w => `<col style="width:${(w / sum * 100).toFixed(3)}%">`).join('');
}
const T = {
  rowInsert(below){
    const cell = curCell(); if (!cell) return;
    SW.edit(() => {
      const P = cellPos(cell); const {G} = P;
      const r = below ? P.r + ((cell.rowSpan || 1) - 1) : P.r;
      const tr = document.createElement('tr');
      const seen = new Set();
      for (let c = 0; c < G.ncols; c++){
        const s = G.grid[r][c]; if (!s || seen.has(s.cell)) continue;
        seen.add(s.cell);
        const spansBoundary = below ? (s.r0 + (s.cell.rowSpan || 1) - 1 > r) : (s.r0 < r);
        if (spansBoundary) s.cell.rowSpan = (s.cell.rowSpan || 1) + 1;
        else { const td = newCell(); if (s.cell.colSpan > 1) td.colSpan = s.cell.colSpan; tr.appendChild(td); }
      }
      const ref = G.rows[r];
      if (below) ref.after(tr); else ref.before(tr);
      const first = tr.querySelector('td'); if (first) placeIn(first);
    });
  },
  colInsert(right){
    const cell = curCell(); if (!cell) return;
    SW.edit(() => {
      const P = cellPos(cell); const {G, tbl} = P;
      const c = right ? P.c + (cell.colSpan || 1) : P.c;   // neue Spalte beginnt bei c
      const seen = new Set();
      G.rows.forEach((tr, r) => {
        const s = c < G.ncols ? G.grid[r][c] : null;
        const left = c > 0 ? G.grid[r][c - 1] : null;
        if (s && left && s.cell === left.cell){       // Zelle überspannt Grenze
          if (!seen.has(s.cell)){ s.cell.colSpan++; seen.add(s.cell); }
          return;
        }
        if (s && s.r0 !== r) return;                  // vertikal verbunden – oben behandelt
        if (!s && left && left.r0 !== r) return;
        const td = newCell();
        if (s) s.cell.before(td);
        else tr.appendChild(td);
        if (s && s.cell.rowSpan > 1) td.rowSpan = s.cell.rowSpan;
        else if (!s && left && left.cell.rowSpan > 1) td.rowSpan = left.cell.rowSpan;
      });
      setCols(tbl, G.ncols + 1, c);
    });
  },
  rowDelete(){
    const cell = curCell(); if (!cell) return;
    SW.edit(() => {
      const P = cellPos(cell); const {G, tbl, r} = P;
      if (G.rows.length === 1){ removeTable(tbl); return; }
      const seen = new Set();
      for (let c = 0; c < G.ncols; c++){
        const s = G.grid[r][c]; if (!s || seen.has(s.cell)) continue; seen.add(s.cell);
        if ((s.cell.rowSpan || 1) > 1){
          if (s.r0 < r) s.cell.rowSpan--;
          else {
            // Zelle in die nächste Zeile verschieben
            s.cell.rowSpan--;
            const nextRow = G.rows[r + 1];
            const after = Array.from(nextRow.cells).find(x => { const p = G.grid[r + 1].findIndex(g => g && g.cell === x); return p > c; });
            nextRow.insertBefore(s.cell, after || null);
          }
        }
      }
      const tr = G.rows[r];
      const target = G.rows[r + 1] || G.rows[r - 1];
      tr.remove();
      const td = target && target.querySelector('td,th'); if (td) placeIn(td);
    });
  },
  colDelete(){
    const cell = curCell(); if (!cell) return;
    SW.edit(() => {
      const P = cellPos(cell); const {G, tbl, c} = P;
      if (G.ncols === 1){ removeTable(tbl); return; }
      const seen = new Set();
      G.rows.forEach((tr, r) => {
        const s = G.grid[r][c]; if (!s || seen.has(s.cell)) return; seen.add(s.cell);
        if ((s.cell.colSpan || 1) > 1) s.cell.colSpan--; else s.cell.remove();
      });
      setCols(tbl, G.ncols - 1, null, c);
      Array.from(tbl.rows).forEach(r => { if (!r.cells.length) r.remove(); });
      const td = tbl.querySelector('td,th'); if (td) placeIn(td);
    });
  },
  merge(){
    const cell = curCell(); if (!cell) return;
    SW.edit(() => {
      const next = cell.nextElementSibling;
      if (!next || (next.rowSpan || 1) !== (cell.rowSpan || 1)){ SW.toast('Keine passende Zelle rechts zum Verbinden.'); return; }
      const hasText = el => el.textContent.trim() || el.querySelector('img');
      if (hasText(next)){ while (next.firstChild) cell.appendChild(next.firstChild); }
      cell.colSpan = (cell.colSpan || 1) + (next.colSpan || 1);
      next.remove();
    });
  },
  split(){
    const cell = curCell(); if (!cell) return;
    SW.edit(() => {
      if ((cell.colSpan || 1) > 1){ cell.colSpan--; const td = newCell(); if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan; cell.after(td); }
      else if ((cell.rowSpan || 1) > 1){
        const P = cellPos(cell);
        const r = P.r + cell.rowSpan - 1;
        cell.rowSpan--;
        const row = P.G.rows[r];
        const after = Array.from(row.cells).find(x => { const p = P.G.grid[r].findIndex(g => g && g.cell === x); return p > P.c; });
        const td = newCell(); if (cell.colSpan > 1) td.colSpan = cell.colSpan;
        row.insertBefore(td, after || null);
      } else SW.toast('Diese Zelle ist nicht verbunden.');
    });
  },
  equal(){
    const cell = curCell(); if (!cell) return;
    SW.edit(() => { const P = cellPos(cell); P.tbl.querySelector(':scope > colgroup')?.remove(); setCols(P.tbl, P.G.ncols); });
  }
};
function removeTable(tbl){
  const p = document.createElement('p'); p.innerHTML = '<br>';
  tbl.replaceWith(p); placeIn(p);
}
function placeIn(el){
  const tgt = el.querySelector('p,h1,h2,h3,li') || el;
  const r = document.createRange();
  r.selectNodeContents(tgt); r.collapse(true);
  const s = getSelection(); s.removeAllRanges(); s.addRange(r);
}
function moveCell(dir){
  const cell = curCell(); if (!cell) return false;
  const tbl = cell.closest('table');
  const cells = $$('td,th', tbl).filter(c => c.closest('table') === tbl && !c.closest('tr.sw-gaprow'));
  const i = cells.indexOf(cell) + dir;
  if (i >= cells.length){ T.rowInsert(true); return true; }
  if (i < 0) return true;
  const t = cells[i];
  const tgt = t.querySelector('p,h1,h2,h3,li') || t;
  const r = document.createRange(); r.selectNodeContents(tgt);
  const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  return true;
}

/* ---------------- Bilder ---------------- */
let selImg = null;
const handles = document.createElement('div');
handles.id = 'imgHandles';
handles.hidden = true;
handles.innerHTML = '<span class="h tl"></span><span class="h tr"></span><span class="h bl"></span><span class="h br"></span>';
$('#paper').appendChild(handles);
SW.selectedImage = () => selImg && ed.contains(selImg) ? selImg : null;
function selectImage(img){
  if (selImg) selImg.classList.remove('sw-imgsel');
  selImg = img;
  if (img){
    img.classList.add('sw-imgsel');
    const r = document.createRange(); r.selectNode(img);
    const s = getSelection(); s.removeAllRanges(); s.addRange(r);
    SW.lastRange = r.cloneRange();
  }
  SW.positionHandles();
  updateContextTabs();
}
SW.positionHandles = function(){
  if (!SW.selectedImage()){ handles.hidden = true; return; }
  const pr = $('#paper').getBoundingClientRect(), r = selImg.getBoundingClientRect(), z = SW.zoom;
  handles.hidden = false;
  handles.style.left = (r.left - pr.left) / z + 'px';
  handles.style.top = (r.top - pr.top) / z + 'px';
  handles.style.width = r.width / z + 'px';
  handles.style.height = r.height / z + 'px';
  const cm = SW.CM;
  $('#imgWidth').value = (r.width / z / cm).toFixed(2);
  $('#imgHeight').value = (r.height / z / cm).toFixed(2);
};
handles.addEventListener('mousedown', e => {
  const h = e.target.closest('.h'); if (!h || !selImg) return;
  e.preventDefault(); e.stopPropagation();
  SW.flushCommit();
  const startX = e.clientX, w0 = selImg.getBoundingClientRect().width / SW.zoom;
  const left = h.classList.contains('tl') || h.classList.contains('bl');
  const maxW = ed.clientWidth;
  const mv = ev => {
    let w = w0 + (left ? -1 : 1) * (ev.clientX - startX) / SW.zoom;
    w = Math.max(16, Math.min(maxW, w));
    selImg.style.width = Math.round(w) + 'px';
    selImg.style.height = '';
    SW.positionHandles();
  };
  const up = () => { removeEventListener('mousemove', mv); removeEventListener('mouseup', up); SW.commit(); SW.markDirty(); SW.relayout(true); SW.positionHandles(); };
  addEventListener('mousemove', mv); addEventListener('mouseup', up);
});
ed.addEventListener('mousedown', e => {
  if (e.target.tagName === 'IMG'){ e.preventDefault(); ed.focus({preventScroll:true}); selectImage(e.target); return; }
  if (selImg) selectImage(null);
});
function setImgCm(w, h){
  const img = SW.selectedImage(); if (!img) return;
  SW.edit(() => {
    const ratio = img.naturalHeight / img.naturalWidth || (img.offsetHeight / img.offsetWidth);
    if (w) img.style.width = Math.min(ed.clientWidth, w * SW.CM) + 'px';
    else if (h) img.style.width = Math.min(ed.clientWidth, h * SW.CM / ratio) + 'px';
  });
  setTimeout(SW.positionHandles, 80);
}
$('#imgWidth').addEventListener('change', e => setImgCm(SW.parseNum(e.target.value), null));
$('#imgHeight').addEventListener('change', e => setImgCm(null, SW.parseNum(e.target.value)));

/* ---------------- Tabellen-Spaltenbreite ziehen ---------------- */
const colRes = document.createElement('div');
colRes.id = 'colResizer'; colRes.hidden = true;
$('#paper').appendChild(colRes);
let colTarget = null;
ed.addEventListener('mousemove', e => {
  if (colRes.classList.contains('drag')) return;
  const cell = e.target.closest && e.target.closest('td,th');
  if (!cell){ if (e.target !== colRes) colRes.hidden = true; return; }
  const r = cell.getBoundingClientRect();
  if (Math.abs(e.clientX - r.right) < 5 && cell.nextElementSibling){
    const pr = $('#paper').getBoundingClientRect(), tr = cell.closest('table').getBoundingClientRect(), z = SW.zoom;
    colRes.hidden = false;
    colRes.style.left = (r.right - pr.left) / z + 'px';
    colRes.style.top = (tr.top - pr.top) / z + 'px';
    colRes.style.height = tr.height / z + 'px';
    colTarget = cell;
  } else colRes.hidden = true;
});
colRes.addEventListener('mousedown', e => {
  if (!colTarget) return;
  e.preventDefault();
  const P = cellPos(colTarget); if (!P) return;
  const tbl = P.tbl;
  setCols(tbl, P.G.ncols);
  const cols = Array.from(tbl.querySelector(':scope > colgroup').children);
  const ci = P.c + (colTarget.colSpan || 1) - 1;
  if (!cols[ci + 1]) return;
  SW.flushCommit();
  colRes.classList.add('drag');
  const tw = tbl.getBoundingClientRect().width;
  const a0 = parseFloat(cols[ci].style.width), b0 = parseFloat(cols[ci + 1].style.width);
  const x0 = e.clientX, left0 = parseFloat(colRes.style.left);
  const mv = ev => {
    let d = (ev.clientX - x0) / tw * 100;
    d = Math.max(-a0 + 3, Math.min(b0 - 3, d));
    cols[ci].style.width = (a0 + d).toFixed(3) + '%';
    cols[ci + 1].style.width = (b0 - d).toFixed(3) + '%';
    colRes.style.left = (left0 + d / 100 * tw / SW.zoom) + 'px';
  };
  const up = () => { removeEventListener('mousemove', mv); removeEventListener('mouseup', up); colRes.classList.remove('drag'); colRes.hidden = true; SW.commit(); SW.markDirty(); SW.relayout(true); };
  addEventListener('mousemove', mv); addEventListener('mouseup', up);
});

/* ---------------- Format übertragen ---------------- */
let painter = null;
function captureFormat(){
  const r = SW.getRange(); if (!r) return null;
  let n = r.startContainer; if (n.nodeType === 3) n = n.parentNode;
  const cs = getComputedStyle(n);
  return {
    bold: document.queryCommandState('bold'), italic: document.queryCommandState('italic'),
    underline: document.queryCommandState('underline'), strike: document.queryCommandState('strikeThrough'),
    font: cs.fontFamily.split(',')[0].replace(/["']/g, ''), size: parseFloat(cs.fontSize) / SW.PT,
    color: cs.color, bg: (() => { let x = n; while (x && x !== ed){ const b = getComputedStyle(x).backgroundColor; if (SW.rgbToHex(b)) return b; x = x.parentNode; } return null; })()
  };
}
function applyPainter(){
  const f = painter; painter = null;
  $$('[data-cmd="formatPainter"]').forEach(b => b.classList.remove('on'));
  ed.style.cursor = '';
  if (!f) return;
  SW.edit(() => {
    SW.native('removeFormat');
    if (f.bold) SW.native('bold');
    if (f.italic) SW.native('italic');
    if (f.underline) SW.native('underline');
    if (f.strike) SW.native('strikeThrough');
    SW.native('fontName', f.font);
    SW.native('foreColor', f.color);
    if (f.bg) SW.native('hiliteColor', f.bg);
    SW.applySpanStyle('fontSize', (Math.round(f.size * 2) / 2) + 'pt');
  });
}
ed.addEventListener('mouseup', () => {
  if (painter) setTimeout(() => { const r = SW.getRange(); if (r && !r.collapsed) applyPainter(); }, 0);
});

/* ---------------- Fußnoten ---------------- */
SW.editFootnote = function(ref, isNew){
  if (!ref) return;
  const dlg = SW.dialog({
    title:`Fußnote ${ref.textContent}`, width:440,
    body:`<div class="dlg-row"><textarea id="fnText" placeholder="Text der Fußnote">${SW.esc(ref.dataset.note || '')}</textarea></div>`,
    buttons:[
      {label:'Fußnote löschen', action:() => { SW.edit(() => ref.remove()); }},
      {label:'Abbrechen', action:() => { if (isNew && !ref.dataset.note) SW.edit(() => ref.remove()); }},
      {label:'OK', primary:true, action:d => { const v = d.querySelector('#fnText').value.trim(); SW.edit(() => { ref.dataset.note = v; }); }}
    ]
  });
};
ed.addEventListener('click', e => {
  const fn = e.target.closest('.sw-fn');
  if (fn) SW.editFootnote(fn);
  const toc = e.target.closest('.toc-row');
  if (toc){
    const ents = SW.tocEntries();
    const ent = ents[+toc.dataset.i];
    if (ent) SW.scrollToEl(ent.h);
  }
  const a = e.target.closest('a[href]');
  if (a && (e.ctrlKey || e.metaKey)){
    e.preventDefault();
    const href = a.getAttribute('href');
    if (href.startsWith('#')){ const t = document.getElementById(href.slice(1)); if (t) SW.scrollToEl(t); }
    else window.open(href, '_blank', 'noopener');
  }
});

/* ---------------- Befehle ---------------- */
const C = {};
SW.cmd = function(name, arg){
  const f = C[name];
  if (!f){ console.warn('Unbekannter Befehl', name); return; }
  f(arg);
};

// Datei
C.save = () => SW.save();
C.saveAs = () => SW.saveAs();
C.open = () => { closeBackstage(); $('#fileOpen').click(); };
C.exportPdf = () => { closeBackstage(); SW.printDoc(true); };
C.print = () => { closeBackstage(); SW.printDoc(false); };
C.closeBackstage = () => closeBackstage();
C.undo = () => SW.undo();
C.redo = () => SW.redo();

// Zwischenablage
C.cut = () => { SW.focusEditor(); if (!document.execCommand('cut')) SW.toast('Bitte Strg+X verwenden.'); };
C.copy = () => { SW.focusEditor(); if (!document.execCommand('copy')) SW.toast('Bitte Strg+C verwenden.'); };
C.paste = async () => {
  SW.focusEditor();
  try {
    const items = await navigator.clipboard.read();
    for (const it of items){
      const img = it.types.find(t => t.startsWith('image/'));
      if (img){ await SW.insertImageFile(await it.getType(img)); return; }
      if (it.types.includes('text/html')){ const h = await (await it.getType('text/html')).text(); SW.edit(() => SW.insertHTML(SW.sanitize(h))); return; }
      if (it.types.includes('text/plain')){ const t = await (await it.getType('text/plain')).text(); SW.edit(() => SW.native('insertText', t)); return; }
    }
  } catch(e){
    SW.toast('Einfügen über die Schaltfläche wurde vom Browser blockiert – bitte Strg+V verwenden.');
  }
};
C.formatPainter = () => {
  painter = captureFormat();
  $$('[data-cmd="formatPainter"]').forEach(b => b.classList.toggle('on', !!painter));
  ed.style.cursor = painter ? 'copy' : '';
  if (painter) SW.toast('Markieren Sie jetzt den Text, auf den das Format übertragen werden soll.');
};

// Zeichen
['bold','italic','underline','strikeThrough','subscript','superscript'].forEach(c => {
  C[c] = () => SW.inline(() => SW.native(c));
});
C.growFont = () => {
  const s = SW.currentFormat().size;
  SW.setFontSize(SW.FONT_SIZES.find(x => x > s + 0.01) || s + 10);
};
C.shrinkFont = () => {
  const s = SW.currentFormat().size;
  SW.setFontSize([...SW.FONT_SIZES].reverse().find(x => x < s - 0.01) || Math.max(1, s - 1));
};
C.changeCase = m => SW.changeCase(m);
C.clearFormat = () => SW.edit(() => {
  const exp = SW.expandToWord();
  SW.native('removeFormat');
  SW.native('unlink');
  const r = SW.getRange();
  if (r) SW.$$('span[style]', ed).forEach(s => { if (r.intersectsNode(s) && !s.closest('.sw-toc')) s.replaceWith(...s.childNodes); });
  SW.selectedBlocks().forEach(b => {
    if (b.matches('td,th,li')) return;
    const n = SW.retag(b, 'p');
    ['sw-nospace','sw-title','sw-subtitle','sw-quote','sw-box'].forEach(c => n.classList.remove(c));
    n.removeAttribute('style'); n.removeAttribute('data-ls');
    if (!n.getAttribute('class')) n.removeAttribute('class');
  });
  if (exp) SW.collapseBackFromWord();
});
C.foreColor = c => SW.inline(() => SW.native('foreColor', !c || c === 'auto' ? '#000000' : c || colorState.foreColor));
C.highlight = c => SW.inline(() => SW.native('hiliteColor', c === 'none' ? 'transparent' : (c || colorState.highlight)));

// Absatz
C.bullets = f => SW.toggleList('ul', f);
C.numbering = f => SW.toggleList('ol', f);
C.restartNumbering = () => SW.restartNumbering();
C.indent = () => SW.indent(1);
C.outdent = () => SW.indent(-1);
C.alignLeft = () => SW.setAlign('left');
C.alignCenter = () => SW.setAlign('center');
C.alignRight = () => SW.setAlign('right');
C.alignJustify = () => SW.setAlign('justify');
C.lineSpacing = m => SW.setLineSpacing(+m);
C.spaceBefore = () => SW.setBlockProp(b => { b.style.marginTop = parseFloat(getComputedStyle(b).marginTop) > 0 ? '0pt' : '12pt'; });
C.spaceAfter = () => SW.setBlockProp(b => { b.style.marginBottom = parseFloat(getComputedStyle(b).marginBottom) > 0 ? '0pt' : SW.doc.base.after + 'pt'; });
C.shading = c => SW.setBlockProp(b => { b.style.backgroundColor = c === 'none' ? '' : (c || colorState.shading); });
C.paraBorder = side => SW.setBlockProp(b => {
  b.classList.remove('sw-box'); b.style.borderTop = ''; b.style.borderBottom = ''; b.style.paddingBottom = '';
  if (side === 'box') b.classList.add('sw-box');
  if (side === 'bottom'){ b.style.borderBottom = '0.5pt solid #000000'; b.style.paddingBottom = '1pt'; }
  if (side === 'top') b.style.borderTop = '0.5pt solid #000000';
  if (!b.getAttribute('class')) b.removeAttribute('class');
});
C.toggleMarks = () => { document.body.classList.toggle('show-marks'); $('#chkMarks').checked = document.body.classList.contains('show-marks'); $('#btnMarks').classList.toggle('on', $('#chkMarks').checked); SW.relayout(true); };
C.selectAll = () => { SW.focusEditor(); const r = document.createRange(); r.selectNodeContents(ed); const s = getSelection(); s.removeAllRanges(); s.addRange(r); };

C.paragraphDialog = () => {
  const b = SW.selectedBlocks()[0]; if (!b) return;
  const cs = getComputedStyle(b);
  const cmOf = v => v && v.endsWith('cm') ? parseFloat(v) : (parseFloat(v) || 0) / SW.CM;
  const ti = cmOf(b.style.textIndent);
  const ls = b.dataset.ls || (parseFloat(cs.lineHeight) / parseFloat(cs.fontSize) / 1.22).toFixed(2);
  const al = cs.textAlign === 'start' ? 'left' : cs.textAlign;
  SW.dialog({
    title:'Absatz', width:470,
    body:`<div class="dlg-section">Allgemein</div>
      <div class="dlg-row"><label>Ausrichtung</label><select id="pAl"><option value="left">Links</option><option value="center">Zentriert</option><option value="right">Rechts</option><option value="justify">Blocksatz</option></select></div>
      <div class="dlg-section">Einzug</div>
      <div class="dlg-grid"><span>Links (cm)</span><input id="pL" type="number" step="0.25" value="${cmOf(b.style.marginLeft).toFixed(2)}"><span>Sondereinzug</span><select id="pSp"><option value="none">(ohne)</option><option value="first">Erste Zeile</option><option value="hang">Hängend</option></select>
      <span>Rechts (cm)</span><input id="pR" type="number" step="0.25" value="${cmOf(b.style.marginRight).toFixed(2)}"><span>Um (cm)</span><input id="pSpV" type="number" step="0.25" min="0" value="${Math.abs(ti).toFixed(2)}"></div>
      <div class="dlg-section">Abstand</div>
      <div class="dlg-grid"><span>Vor (Pt.)</span><input id="pB" type="number" step="6" min="0" value="${Math.round(parseFloat(cs.marginTop) / SW.PT)}"><span>Zeilenabstand</span><select id="pLs"><option value="1">Einfach</option><option value="1.08">1,08 (Standard)</option><option value="1.15">1,15</option><option value="1.5">1,5 Zeilen</option><option value="2">Doppelt</option><option value="custom">Mehrfach</option></select>
      <span>Nach (Pt.)</span><input id="pA" type="number" step="6" min="0" value="${Math.round(parseFloat(cs.marginBottom) / SW.PT)}"><span>Mehrfach</span><input id="pLv" type="number" step="0.1" min="0.5" max="5" value="${(+ls).toFixed(2)}"></div>`,
    onOpen:d => {
      d.querySelector('#pAl').value = ['left','center','right','justify'].includes(al) ? al : 'left';
      d.querySelector('#pSp').value = ti > 0 ? 'first' : ti < 0 ? 'hang' : 'none';
      const sel = d.querySelector('#pLs'); const opt = [...sel.options].find(o => Math.abs(+o.value - ls) < 0.02);
      sel.value = opt ? opt.value : 'custom';
      sel.onchange = () => { if (sel.value !== 'custom') d.querySelector('#pLv').value = sel.value; };
    },
    buttons:[{label:'Abbrechen'}, {label:'OK', primary:true, action:d => {
      const v = id => SW.parseNum(d.querySelector(id).value) || 0;
      const al2 = d.querySelector('#pAl').value, sp = d.querySelector('#pSp').value;
      SW.setBlockProp(x => {
        x.style.textAlign = al2 === 'left' ? '' : al2;
        x.style.marginLeft = v('#pL') ? v('#pL').toFixed(2) + 'cm' : '';
        x.style.marginRight = v('#pR') ? v('#pR').toFixed(2) + 'cm' : '';
        x.style.textIndent = sp === 'none' || !v('#pSpV') ? '' : (sp === 'hang' ? -v('#pSpV') : v('#pSpV')).toFixed(2) + 'cm';
        x.style.marginTop = v('#pB') + 'pt';
        x.style.marginBottom = v('#pA') + 'pt';
        const m = v('#pLv') || 1.08;
        x.dataset.ls = m; x.style.lineHeight = (m * 1.22).toFixed(3);
      });
    }}]
  });
};

// Bearbeiten
C.find = () => SW.openFind(false);
C.replace = () => SW.openFind(true);

// Einfügen
C.pageBreak = () => SW.edit(() => {
  const d = document.createElement('div');
  d.className = 'sw-pagebreak'; d.setAttribute('contenteditable', 'false');
  SW.insertBlockAtCaret(d);
});
C.tableDialog = () => SW.dialog({
  title:'Tabelle einfügen', width:320,
  body:'<div class="dlg-row"><label>Spaltenanzahl</label><input id="tC" type="number" min="1" max="63" value="3"></div><div class="dlg-row"><label>Zeilenanzahl</label><input id="tR" type="number" min="1" max="500" value="3"></div>',
  buttons:[{label:'Abbrechen'}, {label:'OK', primary:true, action:d => { const c = Math.max(1, Math.min(63, +d.querySelector('#tC').value || 3)), r = Math.max(1, Math.min(500, +d.querySelector('#tR').value || 3)); setTimeout(() => SW.insertTable(r, c), 0); }}]
});
C.insertImage = () => $('#fileImage').click();
C.link = () => {
  SW.rememberRange();
  const r = SW.lastRange;
  const a = r && (r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentNode).closest('a');
  const txt = a ? a.textContent : (r ? r.toString() : '');
  SW.dialog({
    title:'Link einfügen', width:440,
    body:`<div class="dlg-row"><label>Anzuzeigender Text</label><input id="lT" type="text" value="${SW.esc(txt)}"></div><div class="dlg-row"><label>Adresse</label><input id="lU" type="text" placeholder="https://… oder mailto:…" value="${SW.esc(a ? a.getAttribute('href') : '')}"></div>`,
    onOpen:d => setTimeout(() => d.querySelector('#lU').focus(), 0),
    buttons:[
      ...(a ? [{label:'Link entfernen', action:() => SW.edit(() => { a.replaceWith(...a.childNodes); })}] : []),
      {label:'Abbrechen'},
      {label:'OK', primary:true, action:d => {
        let url = d.querySelector('#lU').value.trim();
        const t = d.querySelector('#lT').value || url;
        if (!url) return;
        if (!/^(https?:|mailto:|#|file:|tel:)/i.test(url)) url = (/@/.test(url) && !/\//.test(url) ? 'mailto:' : 'https://') + url;
        SW.edit(() => {
          if (a){ a.setAttribute('href', url); if (t !== a.textContent) a.textContent = t; return; }
          SW.insertHTML(`<a href="${SW.esc(url)}">${SW.esc(t)}</a>`);
        });
      }}
    ]
  });
};
C.headerFooter = which => {
  const D = SW.doc;
  const body = `
    <div class="dlg-section">Kopfzeile</div>
    <div class="dlg-row"><textarea id="hT" placeholder="Text der Kopfzeile">${SW.esc(D.header.text)}</textarea></div>
    <div class="dlg-row"><label>Ausrichtung</label><select id="hA"><option value="left">Links</option><option value="center">Zentriert</option><option value="right">Rechts</option></select></div>
    <div class="dlg-section">Fußzeile</div>
    <div class="dlg-row"><textarea id="fT" placeholder="Text der Fußzeile">${SW.esc(D.footer.text)}</textarea></div>
    <div class="dlg-row"><label>Ausrichtung</label><select id="fA"><option value="left">Links</option><option value="center">Zentriert</option><option value="right">Rechts</option></select></div>
    <p class="muted">Platzhalter: <b>{SEITE}</b> = aktuelle Seitenzahl, <b>{SEITEN}</b> = Seitenanzahl. Beispiel: „Seite {SEITE} von {SEITEN}“.</p>`;
  SW.dialog({
    title:'Kopf- und Fußzeile', width:480, body,
    onOpen:d => {
      d.querySelector('#hA').value = D.header.align; d.querySelector('#fA').value = D.footer.align;
      const f = d.querySelector(which === 'footer' ? '#fT' : '#hT'); setTimeout(() => f.focus(), 0);
    },
    buttons:[{label:'Abbrechen'}, {label:'OK', primary:true, action:d => {
      SW.edit(() => {
        D.header = {text:d.querySelector('#hT').value, align:d.querySelector('#hA').value};
        D.footer = {text:d.querySelector('#fT').value, align:d.querySelector('#fA').value};
      });
    }}]
  });
};
C.pageNumber = pos => SW.edit(() => {
  const D = SW.doc;
  const strip = t => t.replace(/Seite \{SEITE\} von \{SEITEN\}|\{SEITE\}/g, '').trim();
  D.header.text = strip(D.header.text); D.footer.text = strip(D.footer.text);
  if (pos === 'footer-center'){ D.footer = {text:(D.footer.text ? D.footer.text + '\n' : '') + '{SEITE}', align:'center'}; }
  if (pos === 'footer-right'){ D.footer = {text:(D.footer.text ? D.footer.text + '\n' : '') + '{SEITE}', align:'right'}; }
  if (pos === 'footer-of'){ D.footer = {text:(D.footer.text ? D.footer.text + '\n' : '') + 'Seite {SEITE} von {SEITEN}', align:'center'}; }
  if (pos === 'header-right'){ D.header = {text:(D.header.text ? D.header.text + '\n' : '') + '{SEITE}', align:'right'}; }
});
C.insertDate = () => {
  const now = new Date();
  const opts = [
    now.toLocaleDateString('de-DE'),
    now.toLocaleDateString('de-DE', {weekday:'long', day:'numeric', month:'long', year:'numeric'}),
    now.toLocaleDateString('de-DE', {day:'numeric', month:'long', year:'numeric'}),
    now.toLocaleDateString('de-DE', {day:'2-digit', month:'2-digit', year:'2-digit'}),
    now.toISOString().slice(0, 10),
    now.toLocaleDateString('de-DE') + ' ' + now.toLocaleTimeString('de-DE', {hour:'2-digit', minute:'2-digit'}),
    now.toLocaleTimeString('de-DE', {hour:'2-digit', minute:'2-digit'})
  ];
  SW.rememberRange();
  SW.dialog({
    title:'Datum und Uhrzeit', width:340,
    body:`<div class="dlg-row"><select id="dF" size="7" style="height:auto">${opts.map((o, i) => `<option value="${i}"${i === 0 ? ' selected' : ''}>${o}</option>`).join('')}</select></div>`,
    buttons:[{label:'Abbrechen'}, {label:'OK', primary:true, action:d => { const v = opts[+d.querySelector('#dF').value]; setTimeout(() => SW.edit(() => SW.native('insertText', v)), 0); }}]
  });
};
C.insertHr = () => SW.edit(() => SW.insertBlockAtCaret(document.createElement('hr')));
C.textbox = () => SW.setBlockProp(b => b.classList.toggle('sw-box'));
C.footnote = () => {
  let ref;
  SW.edit(() => {
    SW.focusEditor();
    const r = SW.getRange(); if (!r) return;
    r.collapse(false);
    ref = document.createElement('sup');
    ref.className = 'sw-fn'; ref.setAttribute('contenteditable', 'false'); ref.dataset.note = ''; ref.textContent = '*';
    r.insertNode(ref);
    const nr = document.createRange(); nr.setStartAfter(ref); nr.collapse(true);
    const s = getSelection(); s.removeAllRanges(); s.addRange(nr);
  });
  SW.relayout(true);
  if (ref) SW.editFootnote(ref, true);
};
C.insertToc = () => {
  if (ed.querySelector('.sw-toc')){ SW.toast('Das Dokument enthält bereits ein Inhaltsverzeichnis.'); return; }
  SW.edit(() => {
    const t = document.createElement('div');
    t.className = 'sw-toc'; t.setAttribute('contenteditable', 'false');
    SW.insertBlockAtCaret(t);
  });
  SW.relayout(true);
};
C.updateToc = () => { SW.relayout(true); SW.toast(ed.querySelector('.sw-toc') ? 'Inhaltsverzeichnis aktualisiert.' : 'Kein Inhaltsverzeichnis vorhanden.'); };

// Layout
C.margins = k => SW.edit(() => { SW.doc.margins = {...SW.MARGIN_PRESETS[k]}; });
C.orientation = o => SW.edit(() => { SW.doc.orientation = o; });
C.pageSize = s => SW.edit(() => { SW.doc.pageSize = s; });
C.pageSetup = () => {
  const D = SW.doc, m = D.margins;
  SW.dialog({
    title:'Seite einrichten', width:460,
    body:`<div class="dlg-section">Seitenränder (cm)</div>
      <div class="dlg-grid"><span>Oben</span><input id="mT" type="number" step="0.1" min="0" value="${m.top}"><span>Unten</span><input id="mB" type="number" step="0.1" min="0" value="${m.bottom}">
      <span>Links</span><input id="mL" type="number" step="0.1" min="0" value="${m.left}"><span>Rechts</span><input id="mR" type="number" step="0.1" min="0" value="${m.right}"></div>
      <div class="dlg-section">Papier</div>
      <div class="dlg-row"><label>Format</label><select id="pS">${Object.keys(SW.PAGE_SIZES).map(k => `<option${k === D.pageSize ? ' selected' : ''}>${k}</option>`).join('')}</select></div>
      <div class="dlg-row"><label>Ausrichtung</label><select id="pO"><option value="portrait">Hochformat</option><option value="landscape"${D.orientation === 'landscape' ? ' selected' : ''}>Querformat</option></select></div>
      <div class="dlg-section">Standardschrift des Dokuments</div>
      <div class="dlg-grid"><span>Schriftart</span><input id="bF" type="text" list="fontList" value="${SW.esc(D.base.font)}"><span>Größe (Pt.)</span><input id="bS" type="number" step="0.5" min="6" value="${D.base.size}"></div>`,
    buttons:[{label:'Abbrechen'}, {label:'OK', primary:true, action:d => {
      const v = id => Math.max(0, SW.parseNum(d.querySelector(id).value) || 0);
      SW.edit(() => {
        D.margins = {top:v('#mT'), bottom:v('#mB'), left:v('#mL'), right:v('#mR')};
        D.pageSize = d.querySelector('#pS').value; D.orientation = d.querySelector('#pO').value;
        D.base.font = d.querySelector('#bF').value.trim() || 'Calibri';
        D.base.size = v('#bS') || 11;
        SW.applyBase();
      });
    }}]
  });
};

// Überprüfen
C.toggleSpell = () => {
  const on = ed.spellcheck = !ed.spellcheck;
  $('#btnSpell').classList.toggle('on', on);
  // Neu zeichnen erzwingen
  const html = ed.innerHTML; ed.blur(); ed.innerHTML = html; SW.fixImages();
  SW.relayout();
  SW.toast(on ? 'Rechtschreibprüfung eingeschaltet (Deutsch).' : 'Rechtschreibprüfung ausgeschaltet.');
};
SW.countWords = function(){
  const clone = SW.cleanClone();
  clone.querySelectorAll('.sw-toc,.sw-fn').forEach(x => x.remove());
  clone.querySelectorAll('p,h1,h2,h3,li,td,th,div,br').forEach(x => x.append(' '));
  const t = clone.textContent.replace(/​/g, '');
  const words = (t.match(/[^\s ]+/g) || []).length;
  const chars = t.replace(/\s+$/, '').replace(/ (?= |$)/g, '');
  return {
    words,
    chars: ed.innerText.replace(/[\n\r​]/g, '').length,
    charsNoSpace: t.replace(/[\s ]/g, '').length,
    paras: $$('p,h1,h2,h3,li', ed).filter(p => p.textContent.trim() && !p.closest('.sw-toc')).length,
    pages: SW.pageCount || 1
  };
};
C.wordCount = () => {
  const c = SW.countWords();
  SW.dialog({
    title:'Wörter zählen', width:330,
    body:`<table class="count-table"><tr><td>Seiten</td><td>${c.pages}</td></tr><tr><td>Wörter</td><td>${c.words.toLocaleString('de-DE')}</td></tr><tr><td>Zeichen (ohne Leerzeichen)</td><td>${c.charsNoSpace.toLocaleString('de-DE')}</td></tr><tr><td>Zeichen (mit Leerzeichen)</td><td>${c.chars.toLocaleString('de-DE')}</td></tr><tr><td>Absätze</td><td>${c.paras.toLocaleString('de-DE')}</td></tr></table>`,
    buttons:[{label:'Schließen', primary:true}]
  });
};

// Ansicht
C.zoomTo = z => SW.setZoom(+z);
C.zoomFit = m => SW.zoomFit(m);
C.zoomIn = () => SW.setZoom(SW.zoom + 0.1);
C.zoomOut = () => SW.setZoom(SW.zoom - 0.1);
C.closeNav = () => { $('#chkNav').checked = false; toggleNav(false); };
C.toggleDark = () => {
  const on = document.documentElement.classList.toggle('dark');
  try { localStorage.setItem('sw-dark', on ? '1' : ''); } catch(e){}
};
C.shortcuts = () => SW.dialog({
  title:'Tastenkombinationen', width:520,
  body:`<table class="kbd-table">${[
    ['Speichern','Strg+S'],['Speichern unter','F12'],['Öffnen','Strg+O'],['Neues Dokument','Strg+N'],['Drucken / PDF','Strg+P'],
    ['Rückgängig / Wiederholen','Strg+Z / Strg+Y'],['Fett','Strg+B oder Strg+Umschalt+F'],['Kursiv','Strg+I oder Strg+Umschalt+K'],['Unterstrichen','Strg+U'],
    ['Hochgestellt / Tiefgestellt','Strg++ / Strg+#'],['Schrift größer / kleiner','Strg+> / Strg+<'],['Links / Zentriert / Rechts / Blocksatz','Strg+L / E / R / J'],
    ['Überschrift 1 / 2 / 3','Strg+Alt+1 / 2 / 3'],['Standard','Strg+Umschalt+N'],['Aufzählung','Strg+Umschalt+L'],['Seitenumbruch','Strg+Enter'],
    ['Zeilenumbruch im Absatz','Umschalt+Enter'],['Suchen / Ersetzen','Strg+F / Strg+H'],['Link einfügen','Strg+K'],['Fußnote','Strg+Alt+F'],['Alles markieren','Strg+A'],['Formatierungszeichen','Strg+*'],['Rechtschreibung ein/aus','F7']
  ].map(([a, b]) => `<tr><td>${a}</td><td>${b.split(' oder ').map(k => k.split(' / ').map(x => `<kbd>${x}</kbd>`).join(' / ')).join(' oder ')}</td></tr>`).join('')}</table>`,
  buttons:[{label:'Schließen', primary:true}]
});

// Tabellenlayout
C.tblRowAbove = () => T.rowInsert(false);
C.tblRowBelow = () => T.rowInsert(true);
C.tblColLeft = () => T.colInsert(false);
C.tblColRight = () => T.colInsert(true);
C.tblDelRow = () => T.rowDelete();
C.tblDelCol = () => T.colDelete();
C.tblDelTable = () => { const c = curCell(); if (c) SW.edit(() => removeTable(c.closest('table'))); };
C.tblMerge = () => T.merge();
C.tblSplit = () => T.split();
C.tblEqual = () => T.equal();
C.cellShade = c => { const cell = curCell(); if (!cell) return; const r = SW.getRange(); const cells = r && !r.collapsed ? $$('td,th', cell.closest('table')).filter(x => r.intersectsNode(x)) : [cell]; SW.edit(() => cells.forEach(x => { x.style.backgroundColor = c === 'none' ? '' : (c || colorState.cellShade); if (!x.getAttribute('style')) x.removeAttribute('style'); })); };
$('#tblHeader').addEventListener('change', e => { const c = curCell(); if (c) SW.edit(() => c.closest('table').classList.toggle('sw-headrow', e.target.checked)); });
$('#tblBorders').addEventListener('change', e => { const c = curCell(); if (c) SW.edit(() => c.closest('table').classList.toggle('sw-noborder', !e.target.checked)); });

// Bild
C.imgReset = () => { const i = SW.selectedImage(); if (i) SW.edit(() => { i.style.width = Math.min(ed.clientWidth, i.naturalWidth) + 'px'; }); setTimeout(SW.positionHandles, 80); };
C.imgBorder = () => { const i = SW.selectedImage(); if (i) SW.edit(() => i.classList.toggle('sw-imgborder')); };
C.imgDelete = () => { const i = SW.selectedImage(); if (i){ selectImage(null); SW.edit(() => i.remove()); } };

/* ---------------- Schriftart-/Größenfelder ---------------- */
function bindCombo(input, apply){
  input.addEventListener('mousedown', () => SW.rememberRange());
  input.addEventListener('focus', () => { SW.rememberRange(); input.select(); });
  const go = () => { const v = input.value.trim(); if (v) apply(v); };
  input.addEventListener('change', go);
  input.addEventListener('keydown', e => {
    if (e.key === 'Enter'){ e.preventDefault(); go(); SW.focusEditor(); }
    if (e.key === 'Escape'){ SW.focusEditor(); refreshUI(); }
  });
}
bindCombo($('#fontName'), v => { SW.focusEditor(); SW.inline(() => SW.native('fontName', v)); });
bindCombo($('#fontSize'), v => { const n = SW.parseNum(v); if (n > 0){ SW.focusEditor(); SW.setFontSize(n); } });
function bindLayoutField(id, fn){
  const el = $(id);
  el.addEventListener('mousedown', () => SW.rememberRange());
  el.addEventListener('change', () => { const v = SW.parseNum(el.value); if (isNaN(v)) return; SW.focusEditor(); SW.setBlockProp(b => fn(b, Math.max(0, v))); });
}
bindLayoutField('#indLeft', (b, v) => { b.style.marginLeft = v ? v.toFixed(2) + 'cm' : ''; });
bindLayoutField('#indRight', (b, v) => { b.style.marginRight = v ? v.toFixed(2) + 'cm' : ''; });
bindLayoutField('#spBefore', (b, v) => { b.style.marginTop = v + 'pt'; });
bindLayoutField('#spAfter', (b, v) => { b.style.marginBottom = v + 'pt'; });

/* ---------------- Zustandsanzeige ---------------- */
function updateContextTabs(){
  const inTbl = !!curCell() && !SW.selectedImage();
  const img = !!SW.selectedImage();
  const tTab = $('.tab[data-tab="tabelle"]'), iTab = $('.tab[data-tab="bild"]');
  tTab.hidden = !inTbl; iTab.hidden = !img;
  const active = $('.tab.active');
  if ((!inTbl && active === tTab) || (!img && active === iTab)) selectTab('start');
  if (inTbl){
    const t = curCell().closest('table');
    $('#tblHeader').checked = t.classList.contains('sw-headrow');
    $('#tblBorders').checked = !t.classList.contains('sw-noborder');
  }
}
let uiT = null;
function refreshUI(){
  cancelAnimationFrame(uiT);
  uiT = requestAnimationFrame(() => { try { refreshNow(); } catch(err){ console.warn(err); } });
}
function refreshNow(){
  {
    const r = SW.getRange();
    if (r){
      ['bold','italic','underline','strikeThrough','subscript','superscript','insertUnorderedList','insertOrderedList'].forEach(c => {
        let on = false; try { on = document.queryCommandState(c); } catch(e){}
        $$(`[data-state="${c}"]`).forEach(b => b.classList.toggle('on', on));
      });
      const f = SW.currentFormat();
      if (document.activeElement !== $('#fontName')) $('#fontName').value = f.font;
      if (document.activeElement !== $('#fontSize')) $('#fontSize').value = String(f.size).replace('.', ',');
      const al = f.align === 'start' ? 'left' : f.align === 'end' ? 'right' : f.align;
      $$('#ribbon [data-align]').forEach(b => b.classList.toggle('on', b.dataset.align === al));
      const st = SW.styleOf(f.block);
      $$('.style-item').forEach(b => b.classList.toggle('on', b.dataset.style === st));
      if (f.block){
        const cs = getComputedStyle(f.block);
        const cm = v => v && v.endsWith('cm') ? parseFloat(v) : 0;
        if (document.activeElement !== $('#indLeft')) $('#indLeft').value = cm(f.block.style.marginLeft).toFixed(2).replace(/\.?0+$/, '') || '0';
        if (document.activeElement !== $('#indRight')) $('#indRight').value = cm(f.block.style.marginRight).toFixed(2).replace(/\.?0+$/, '') || '0';
        if (document.activeElement !== $('#spBefore')) $('#spBefore').value = Math.round(parseFloat(cs.marginTop) / SW.PT);
        if (document.activeElement !== $('#spAfter')) $('#spAfter').value = Math.round(parseFloat(cs.marginBottom) / SW.PT);
      }
      SW.renderRuler();
    }
    updateContextTabs();
    $('#stPage').textContent = `Seite ${SW.caretPage()} von ${SW.pageCount || 1}`;
    SW.updateUndoButtons();
  }
}
SW.refreshUI = refreshUI;
SW.updateUndoButtons = () => {
  $$('[data-cmd="undo"]').forEach(b => b.disabled = !SW.canUndo());
  $$('[data-cmd="redo"]').forEach(b => b.disabled = !SW.canRedo());
};
let wcT = null;
function updateWordCount(){
  clearTimeout(wcT);
  wcT = setTimeout(() => {
    const n = SW.countWords().words;
    $('#stWords').textContent = `${n.toLocaleString('de-DE')} ${n === 1 ? 'Wort' : 'Wörter'}`;
  }, 400);
}
SW.afterLayout = () => {
  $('#stPage').textContent = `Seite ${SW.caretPage()} von ${SW.pageCount || 1}`;
  updateWordCount();
  if (!$('#navPane').hidden) renderNav();
  SW.positionHandles();
};
$('#stWords').addEventListener('click', () => C.wordCount());
$('#zoomPct').addEventListener('click', () => SW.setZoom(1));
$('#zoomSlider').addEventListener('input', e => SW.setZoom(+e.target.value / 100));


SW.updateTitle = function(){
  $('#docTitle').textContent = SW.doc.title;
  $('#dirtyMark').textContent = SW.dirty ? '•' : '';
  document.title = `${SW.dirty ? '• ' : ''}${SW.doc.title} – Schreibwerk`;
};

/* ---------------- Ansicht-Schalter ---------------- */
$('#chkRuler').addEventListener('change', e => { document.body.classList.toggle('no-ruler', !e.target.checked); SW.renderRuler(); });
$('#chkMarks').addEventListener('change', () => C.toggleMarks());
function toggleNav(on){ $('#navPane').hidden = !on; if (on) renderNav(); SW.requestLayout(10); }
$('#chkNav').addEventListener('change', e => toggleNav(e.target.checked));
function renderNav(){
  const ents = SW.tocEntries();
  $('#navList').innerHTML = ents.length
    ? ents.map((e, i) => `<button class="nav-item l${e.lvl}" data-i="${i}" title="${SW.esc(e.text)}">${SW.esc(e.text)}</button>`).join('')
    : '<div class="nav-empty">Erstellen Sie eine interaktive Gliederung, indem Sie Überschrift-Formatvorlagen anwenden.</div>';
}
$('#navList').addEventListener('click', e => {
  const b = e.target.closest('.nav-item'); if (!b) return;
  const ent = SW.tocEntries()[+b.dataset.i];
  if (ent){ SW.scrollToEl(ent.h); const r = document.createRange(); r.selectNodeContents(ent.h); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); ed.focus({preventScroll:true}); }
});

/* ---------------- Suchen & Ersetzen ---------------- */
const F = {matches:[], cur:-1, q:'', dlg:null};
function textGroups(){
  // Text pro Block sammeln
  const groups = [];
  const blocks = $$(SW.BLOCK_SEL + ',td,th', ed).filter(b => !b.closest('.sw-toc'));
  const inner = blocks.filter(b => !b.querySelector(SW.BLOCK_SEL + ',td,th'));
  const seen = new Set();
  const add = (root) => {
    const nodes = []; let text = '';
    const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    let t;
    while ((t = w.nextNode())){
      if (t.parentNode.closest('.sw-toc,.sw-fn,.sw-gap')) continue;
      if (seen.has(t)) continue;
      seen.add(t);
      nodes.push({t, s:text.length}); text += t.data;
    }
    if (nodes.length) groups.push({nodes, text});
  };
  inner.forEach(add);
  add(ed);
  return groups;
}
SW.findAll = function(q, opt={}){
  F.matches = [];
  if (!q) return [];
  const flags = opt.mc ? 'g' : 'gi';
  const src = q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(opt.ww ? `(?<![\\p{L}\\p{N}_])${src}(?![\\p{L}\\p{N}_])` : src, flags + 'u');
  for (const g of textGroups()){
    let m;
    re.lastIndex = 0;
    while ((m = re.exec(g.text))){
      if (!m[0].length){ re.lastIndex++; continue; }
      const a = m.index, b = m.index + m[0].length;
      const loc = i => { let k = g.nodes.length - 1; while (k > 0 && g.nodes[k].s > i) k--; return [g.nodes[k].t, i - g.nodes[k].s]; };
      const [sn, so] = loc(a); const endLoc = (() => { let k = g.nodes.length - 1; while (k > 0 && g.nodes[k].s >= b) k--; return [g.nodes[k].t, b - g.nodes[k].s]; })();
      const r = document.createRange();
      r.setStart(sn, so); r.setEnd(endLoc[0], endLoc[1]);
      F.matches.push(r);
    }
  }
  return F.matches;
};
function paintFind(){
  if (!window.CSS || !CSS.highlights) return;
  CSS.highlights.set('sw-find', new Highlight(...F.matches));
  if (F.cur >= 0 && F.matches[F.cur]) CSS.highlights.set('sw-find-cur', new Highlight(F.matches[F.cur]));
  else CSS.highlights.delete('sw-find-cur');
}
SW.clearFind = () => { F.matches = []; F.cur = -1; if (window.CSS && CSS.highlights){ CSS.highlights.delete('sw-find'); CSS.highlights.delete('sw-find-cur'); } };
function gotoMatch(i){
  if (!F.matches.length) return;
  F.cur = (i + F.matches.length) % F.matches.length;
  const r = F.matches[F.cur];
  const rect = r.getBoundingClientRect(), c = canvas.getBoundingClientRect();
  if (rect.top < c.top + 40 || rect.bottom > c.bottom - 40) canvas.scrollTop += rect.top - c.top - c.height / 3;
  const s = getSelection(); s.removeAllRanges(); s.addRange(r.cloneRange());
  SW.lastRange = r.cloneRange();
  paintFind();
}
SW.openFind = function(withReplace){
  const pre = (SW.getRange() && SW.getRange().toString().length < 80) ? SW.getRange().toString() : F.q;
  if (F.dlg) F.dlg.close();
  F.dlg = SW.dialog({
    title:withReplace ? 'Suchen und Ersetzen' : 'Suchen', modal:false, width:400,
    body:`<div class="dlg-row"><label style="min-width:80px">Suchen nach</label><input id="fQ" type="text" value="${SW.esc(pre)}"></div>
      ${withReplace ? '<div class="dlg-row"><label style="min-width:80px">Ersetzen durch</label><input id="fR" type="text"></div>' : ''}
      <div class="dlg-row" style="gap:16px"><label class="chk" style="min-width:0"><input type="checkbox" id="fMc"> Groß-/Kleinschreibung</label><label class="chk" style="min-width:0"><input type="checkbox" id="fWw"> Nur ganzes Wort</label></div>
      <div class="find-info" id="fInfo"></div>`,
    buttons:[
      ...(withReplace ? [{label:'Alle ersetzen', close:false, action:d => replaceAll(d)}, {label:'Ersetzen', close:false, action:d => replaceOne(d)}] : []),
      {label:'Vorheriges', close:false, action:() => { gotoMatch(F.cur - 1); info(); }},
      {label:'Weitersuchen', primary:true, close:false, action:d => { if (!F.matches.length) run(d); gotoMatch(F.cur + 1); info(); }}
    ],
    onOpen:d => {
      const go = () => { run(d); F.cur = -1; if (F.matches.length) { const r = SW.lastRange; let i = 0; if (r) { i = F.matches.findIndex(m => m.compareBoundaryPoints(Range.START_TO_START, r) >= 0); if (i < 0) i = 0; } F.cur = i - 1; } info(); paintFind(); };
      d.querySelector('#fQ').addEventListener('input', go);
      d.querySelector('#fMc').addEventListener('change', go);
      d.querySelector('#fWw').addEventListener('change', go);
      go();
    }
  });
  const oldClose = F.dlg.close;
  F.dlg.close = () => { SW.clearFind(); oldClose(); F.dlg = null; };
  F.dlg.el.querySelector('.dlg-x').onclick = F.dlg.close;
  function opts(d){ return {mc:d.querySelector('#fMc').checked, ww:d.querySelector('#fWw').checked}; }
  function run(d){ F.q = d.querySelector('#fQ').value; SW.findAll(F.q, opts(d)); paintFind(); }
  function info(){ const el = F.dlg && F.dlg.el.querySelector('#fInfo'); if (!el) return; el.textContent = !F.q ? '' : F.matches.length ? `${F.cur >= 0 ? F.cur + 1 + ' von ' : ''}${F.matches.length} Treffer` : 'Keine Treffer gefunden.'; }
  function replaceOne(d){
    if (F.cur < 0 || !F.matches[F.cur]){ run(d); gotoMatch(0); info(); return; }
    const rep = d.querySelector('#fR').value;
    SW.edit(() => { const r = F.matches[F.cur]; r.deleteContents(); r.insertNode(document.createTextNode(rep)); });
    const keep = F.cur;
    run(d); gotoMatch(keep); info();
  }
  function replaceAll(d){
    run(d);
    const n = F.matches.length;
    if (!n){ info(); return; }
    const rep = d.querySelector('#fR').value;
    SW.edit(() => { for (let i = F.matches.length - 1; i >= 0; i--){ const r = F.matches[i]; r.deleteContents(); r.insertNode(document.createTextNode(rep)); } ed.normalize(); });
    SW.clearFind();
    d.querySelector('#fInfo').textContent = `${n} Ersetzungen vorgenommen.`;
  }
};
// Schnellsuche (Titelleiste) & Navigationsbereich
$('#quickSearch').addEventListener('keydown', e => {
  if (e.key === 'Enter'){ e.preventDefault(); const q = e.target.value; F.q = q; SW.openFind(false); }
  if (e.key === 'Escape'){ e.target.value = ''; SW.focusEditor(); }
});
$('#navSearch').addEventListener('input', e => {
  const q = e.target.value.trim();
  const box = $('#navResults');
  if (!q){ box.innerHTML = ''; SW.clearFind(); return; }
  SW.findAll(q, {});
  paintFind();
  box.innerHTML = F.matches.length ? F.matches.slice(0, 200).map((r, i) => {
    const t = (r.startContainer.data || '');
    const a = Math.max(0, r.startOffset - 30);
    const before = t.slice(a, r.startOffset), hit = r.toString(), after = t.slice(r.endOffset, r.endOffset + 40);
    return `<button class="nav-res" data-i="${i}">${a > 0 ? '…' : ''}${SW.esc(before)}<mark>${SW.esc(hit)}</mark>${SW.esc(after)}</button>`;
  }).join('') : '<div class="nav-empty">Keine Treffer.</div>';
});
$('#navResults').addEventListener('click', e => { const b = e.target.closest('.nav-res'); if (b) gotoMatch(+b.dataset.i); });

/* ---------------- Einfügen aus Zwischenablage (bereinigen) ---------------- */
SW.sanitize = function(html){
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const allowedTags = new Set(['P','H1','H2','H3','H4','H5','H6','B','STRONG','I','EM','U','S','STRIKE','DEL','SUB','SUP','BR','UL','OL','LI','TABLE','THEAD','TBODY','TFOOT','TR','TD','TH','A','SPAN','IMG','HR','BLOCKQUOTE','DIV','FONT','COLGROUP','COL']);
  const blockStyles = ['text-align','margin-left'];
  const inlineStyles = ['font-weight','font-style','text-decoration','text-decoration-line','color','background-color','font-size','font-family','vertical-align'];
  const clean = el => {
    for (const n of Array.from(el.childNodes)){
      if (n.nodeType === 8){ n.remove(); continue; }
      if (n.nodeType === 3) continue;
      if (n.nodeType !== 1){ n.remove(); continue; }
      const tag = n.tagName;
      if (['SCRIPT','STYLE','META','LINK','TITLE','XML','O:P','HEAD','NOSCRIPT','IFRAME','OBJECT','SVG','BUTTON','INPUT','SELECT','TEXTAREA'].includes(tag) || /:/.test(n.tagName)){
        if (tag === 'O:P' || /:/.test(tag)){ n.replaceWith(...n.childNodes); clean(el); return; }
        n.remove(); continue;
      }
      clean(n);
      if (!allowedTags.has(tag)){ n.replaceWith(...n.childNodes); continue; }
      // Attribute filtern
      const st = n.style;
      const keep = {};
      const list = /^(P|H\d|LI|DIV|BLOCKQUOTE|TD|TH)$/.test(tag) ? [...blockStyles, ...(tag === 'TD' || tag === 'TH' ? ['background-color'] : [])] : inlineStyles;
      if (tag === 'SPAN' || tag === 'FONT' || /^(B|I|U|A|STRONG|EM)$/.test(tag)) list.push(...inlineStyles);
      list.forEach(p => { const v = st.getPropertyValue(p); if (v) keep[p] = v; });
      if (keep['font-size']){ const v = keep['font-size']; if (v.endsWith('px')) keep['font-size'] = (parseFloat(v) * 0.75) + 'pt'; }
      if (keep['font-family']) keep['font-family'] = keep['font-family'].split(',')[0];
      if (keep.color && SW.rgbToHex(keep.color) === '000000') delete keep.color;
      if (keep['background-color'] && !SW.rgbToHex(keep['background-color'])) delete keep['background-color'];
      const attrs = {};
      if (tag === 'A' && /^(https?:|mailto:|#)/i.test(n.getAttribute('href') || '')) attrs.href = n.getAttribute('href');
      if ((tag === 'TD' || tag === 'TH')){ if (n.colSpan > 1) attrs.colspan = n.colSpan; if (n.rowSpan > 1) attrs.rowspan = n.rowSpan; }
      if (tag === 'IMG'){
        const src = n.getAttribute('src') || '';
        if (!src.startsWith('data:image/')){ n.remove(); continue; }
        attrs.src = src;
        if (n.width) keep.width = Math.min(n.width, ed.clientWidth) + 'px';
      }
      if (tag === 'FONT' && n.getAttribute('color')) keep.color = n.getAttribute('color');
      for (const a of Array.from(n.attributes)) n.removeAttribute(a.name);
      Object.entries(attrs).forEach(([k, v]) => n.setAttribute(k, v));
      const css = Object.entries(keep).map(([k, v]) => `${k}:${v}`).join(';');
      if (css) n.setAttribute('style', css);
      if (/^H[4-6]$/.test(tag)) SW.retag(n, 'h3');
      if (tag === 'DIV') SW.retag(n, 'p');
      if (tag === 'BLOCKQUOTE'){ const p = SW.retag(n, 'p'); p.className = 'sw-quote'; }
      if (tag === 'FONT') SW.retag(n, 'span');
      if (tag === 'SPAN' && !css) n.replaceWith(...n.childNodes);
    }
  };
  clean(doc.body);
  doc.body.querySelectorAll('table').forEach(t => {
    t.removeAttribute('style');
    if (!t.querySelector('colgroup')){
      const n = Math.max(...Array.from(t.rows).map(r => Array.from(r.cells).reduce((s, c) => s + (c.colSpan || 1), 0)), 1);
      const cg = document.createElement('colgroup');
      cg.innerHTML = Array(n).fill(`<col style="width:${(100 / n).toFixed(3)}%">`).join('');
      t.prepend(cg);
    } else t.querySelectorAll('col').forEach(c => c.removeAttribute('width'));
  });
  return doc.body.innerHTML.replace(/ /g, ' ');
};
ed.addEventListener('paste', async e => {
  const dt = e.clipboardData;
  if (!dt) return;
  const files = Array.from(dt.files || []).filter(f => f.type.startsWith('image/'));
  if (files.length){ e.preventDefault(); for (const f of files) await SW.insertImageFile(f); return; }
  const html = dt.getData('text/html');
  if (html){
    e.preventDefault();
    SW.edit(() => {
      SW.insertHTML(SW.sanitize(html));
      // eingefügte Data-URL-Bilder registrieren
      $$('img:not([data-img-id])', ed).forEach(img => {
        const b = dataUrlBlob(img.src);
        if (b){ const {id, url} = SW.registerImage(b); img.dataset.imgId = id; img.src = url; }
      });
    });
    return;
  }
  const text = dt.getData('text/plain');
  if (text){ e.preventDefault(); SW.edit(() => SW.native('insertText', text)); }
});
function dataUrlBlob(u){
  if (!u || !u.startsWith('data:')) return null;
  try { const [h, d] = u.split(','); const mime = (h.match(/data:([^;]+)/) || [])[1]; const bin = atob(d); const a = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) a[i] = bin.charCodeAt(i); return new Blob([a], {type:mime}); } catch(e){ return null; }
}

/* ---------------- Editor-Ereignisse ---------------- */
ed.addEventListener('compositionstart', () => SW._composing = true);
ed.addEventListener('compositionend', () => { SW._composing = false; SW.requestLayout(); });
ed.addEventListener('beforeinput', e => {
  if (e.inputType === 'historyUndo'){ e.preventDefault(); SW.undo(); }
  if (e.inputType === 'historyRedo'){ e.preventDefault(); SW.redo(); }
});
ed.addEventListener('input', e => {
  if (e.inputType === 'insertParagraph'){
    const r = SW.getRange();
    const b = r && SW.innermostBlock(r.startContainer);
    if (b && !b.textContent.trim()){
      // Titel/Untertitel/Zitat → nächster Absatz „Standard“ (wie Word)
      ['sw-title','sw-subtitle','sw-quote'].forEach(c => b.classList.remove(c));
      if (!b.getAttribute('class')) b.removeAttribute('class');
      if (/^H\d$/.test(b.tagName) && b.previousElementSibling && b.previousElementSibling.tagName === b.tagName){ const p = SW.retag(b, 'p'); const nr = document.createRange(); nr.setStart(p, 0); nr.collapse(true); getSelection().removeAllRanges(); getSelection().addRange(nr); }
      b.removeAttribute('data-swp');
    }
  }
  SW.fixNesting();
  SW.markDirty();
  SW.scheduleCommit();
  SW.requestLayout(40);
  refreshUI();
});
document.addEventListener('selectionchange', () => {
  if (SW.getRange()){ SW.rememberRange(); refreshUI(); }
});
ed.addEventListener('contextmenu', e => {
  if (e.ctrlKey) return;  // Strg+Rechtsklick: Browsermenü (Rechtschreibvorschläge)
  const sel = getSelection();
  // Bei Rechtschreibfehlern Browsermenü zulassen: Rechtsklick auf Wort ohne Auswahl
  if (sel.isCollapsed && e.target.nodeType === 1 && !e.target.closest('td,th')) return;
  e.preventDefault();
  SW.rememberRange();
  $('#ctxMenu').classList.toggle('in-table', !!curCell());
  showMenu('ctxMenu', null, {x:e.clientX, y:e.clientY});
});
ed.addEventListener('dragover', e => { if ([...e.dataTransfer.types].includes('Files')) e.preventDefault(); });
ed.addEventListener('drop', async e => {
  const files = Array.from(e.dataTransfer.files || []);
  if (!files.length) return;
  e.preventDefault();
  const doc = files.find(f => /\.(docx|txt|html?)$/i.test(f.name));
  if (doc){ SW.openFile(doc); return; }
  const r = document.caretRangeFromPoint ? document.caretRangeFromPoint(e.clientX, e.clientY) : null;
  if (r){ const s = getSelection(); s.removeAllRanges(); s.addRange(r); }
  for (const f of files.filter(f => f.type.startsWith('image/'))) await SW.insertImageFile(f);
});
document.addEventListener('dragover', e => e.preventDefault());
document.addEventListener('drop', e => {
  if (e.defaultPrevented) return;
  e.preventDefault();
  const f = Array.from(e.dataTransfer.files || []).find(f => /\.(docx|txt|html?)$/i.test(f.name));
  if (f) SW.openFile(f);
});
canvas.addEventListener('scroll', () => { closeMenus(); });
canvas.addEventListener('mousedown', e => {
  // Klick auf leeren Bereich unter dem Text → ans Ende
  if (e.target === canvas || e.target.id === 'zoomBox' || e.target.id === 'paper' || e.target.classList.contains('page')){
    const pr = $('#paper').getBoundingClientRect();
    if (e.clientY > ed.getBoundingClientRect().bottom && e.clientX > pr.left && e.clientX < pr.right){
      e.preventDefault();
      ed.focus({preventScroll:true});
      const last = ed.lastElementChild;
      const r = document.createRange(); r.selectNodeContents(last.matches('table') ? last.querySelector('td:last-child') || last : last); r.collapse(false);
      const s = getSelection(); s.removeAllRanges(); s.addRange(r);
    }
  }
});

/* ---------------- Tastatur ---------------- */
document.addEventListener('keydown', e => {
  const mod = e.ctrlKey || e.metaKey;
  const k = e.key.toLowerCase();
  const inEd = document.activeElement === ed;
  const inDialog = e.target.closest && e.target.closest('.dialog');
  if (e.key === 'Escape'){
    if (!$('#backstage').hidden){ closeBackstage(); return; }
    if (openMenu){ closeMenus(); return; }
    if (painter){ painter = null; $$('[data-cmd="formatPainter"]').forEach(b => b.classList.remove('on')); ed.style.cursor = ''; }
  }
  if (inDialog) return;
  if (mod && !e.altKey){
    if (k === 's'){ e.preventDefault(); SW.save(); return; }
    if (k === 'o'){ e.preventDefault(); C.open(); return; }
    if (k === 'n'){ e.preventDefault(); if (!e.shiftKey) SW.newDoc('blank'); else if (inEd) SW.applyStyle('normal'); return; }
    if (k === 'p'){ e.preventDefault(); SW.printDoc(false); return; }
    if (k === 'f' && !e.shiftKey){ e.preventDefault(); SW.openFind(false); return; }
    if (k === 'h'){ e.preventDefault(); SW.openFind(true); return; }
    if (k === 'z' && !e.shiftKey){ e.preventDefault(); SW.undo(); return; }
    if (k === 'y' || (k === 'z' && e.shiftKey)){ e.preventDefault(); SW.redo(); return; }
    if (e.key === '*' || (e.shiftKey && e.code === 'Digit8')){ e.preventDefault(); C.toggleMarks(); return; }
  }
  if (e.key === 'F12'){ e.preventDefault(); SW.saveAs(); return; }
  if (e.key === 'F7'){ e.preventDefault(); C.toggleSpell(); return; }
  if (!inEd) return;

  // Ausgewähltes Bild löschen
  if (SW.selectedImage() && (e.key === 'Delete' || e.key === 'Backspace')){ e.preventDefault(); C.imgDelete(); return; }

  if (mod && e.altKey){
    if (k === '1' || k === '2' || k === '3'){ e.preventDefault(); SW.applyStyle('h' + k); return; }
    if (k === 'f'){ e.preventDefault(); C.footnote(); return; }
  }
  if (mod){
    if (e.shiftKey && k === 'f'){ e.preventDefault(); C.bold(); return; }
    if (e.shiftKey && k === 'k'){ e.preventDefault(); C.italic(); return; }
    if (e.shiftKey && k === 'l'){ e.preventDefault(); C.bullets(); return; }
    if (e.shiftKey && k === 'u'){ e.preventDefault(); C.underline(); return; }
    if (!e.shiftKey && k === 'b'){ e.preventDefault(); C.bold(); return; }
    if (!e.shiftKey && k === 'i'){ e.preventDefault(); C.italic(); return; }
    if (!e.shiftKey && k === 'u'){ e.preventDefault(); C.underline(); return; }
    if (k === 'e'){ e.preventDefault(); C.alignCenter(); return; }
    if (k === 'l'){ e.preventDefault(); C.alignLeft(); return; }
    if (k === 'r'){ e.preventDefault(); C.alignRight(); return; }
    if (k === 'j'){ e.preventDefault(); C.alignJustify(); return; }
    if (k === 'k'){ e.preventDefault(); C.link(); return; }
    if (k === '#'){ e.preventDefault(); C.subscript(); return; }
    if (k === '+' || e.key === '+'){ e.preventDefault(); C.superscript(); return; }
    if (e.key === '>' || (e.shiftKey && e.key === '<')){ e.preventDefault(); C.growFont(); return; }
    if (e.key === '<'){ e.preventDefault(); C.shrinkFont(); return; }
    if (e.key === 'Enter'){ e.preventDefault(); C.pageBreak(); return; }
  }
  if (e.key === 'Tab' && !mod && !e.altKey){
    if (curCell()){ e.preventDefault(); moveCell(e.shiftKey ? -1 : 1); refreshUI(); return; }
    const r = SW.getRange();
    const li = r && (r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentNode).closest('li');
    if (li){
      // Tab am Anfang eines Listeneintrags: Ebene ändern (wie Word)
      const pre = document.createRange(); pre.selectNodeContents(li); pre.setEnd(r.startContainer, r.startOffset);
      if (!pre.toString().trim() || !r.collapsed){ e.preventDefault(); SW.indent(e.shiftKey ? -1 : 1); return; }
    }
    if (e.shiftKey){ e.preventDefault(); return; }
    e.preventDefault();
    SW.edit(() => SW.insertHTML('<span class="sw-tab">\t</span>​'));
    return;
  }
  // Backspace am Anfang eines Absatzes nach einem Seitenumbruch: Umbruch entfernen
  if (e.key === 'Backspace' && !mod){
    const r = SW.getRange();
    if (r && r.collapsed){
      const b = SW.innermostBlock(r.startContainer);
      if (b && b.parentNode === ed){
        const pre = document.createRange(); pre.selectNodeContents(b); pre.setEnd(r.startContainer, r.startOffset);
        const prev = b.previousElementSibling;
        if (!pre.toString() && prev && (prev.classList.contains('sw-pagebreak') || prev.classList.contains('sw-toc') || prev.tagName === 'HR')){
          e.preventDefault(); SW.edit(() => prev.remove()); return;
        }
      }
    }
  }
  if (e.key === 'Delete' && !mod){
    const r = SW.getRange();
    if (r && r.collapsed){
      const b = SW.innermostBlock(r.startContainer);
      if (b && b.parentNode === ed){
        const post = document.createRange(); post.selectNodeContents(b); post.setStart(r.startContainer, r.startOffset);
        const next = b.nextElementSibling;
        if (!post.toString() && next && (next.classList.contains('sw-pagebreak') || next.classList.contains('sw-toc') || next.tagName === 'HR')){
          e.preventDefault(); SW.edit(() => next.remove()); return;
        }
      }
    }
  }
});

/* ---------------- Datei: Neu / Öffnen / Speichern ---------------- */
function openBackstage(){
  SW.rememberRange();
  closeMenus();
  $('#backstage').hidden = false;
  showBs('neu');
  renderInfo();
  checkRestore();
}
function closeBackstage(){ $('#backstage').hidden = true; SW.focusEditor(); }
function showBs(name){
  $$('.bs-item[data-bs]').forEach(b => b.classList.toggle('active', b.dataset.bs === name));
  $$('.bs-page').forEach(p => p.classList.toggle('active', p.dataset.bspage === name));
}
$('#backstage').addEventListener('click', e => {
  const b = e.target.closest('.bs-item[data-bs]');
  if (b){ showBs(b.dataset.bs); if (b.dataset.bs === 'info') renderInfo(); }
  const t = e.target.closest('.tpl');
  if (t) SW.newDoc(t.dataset.tpl);
});
function renderInfo(){
  const c = SW.countWords();
  const g = SW.geo();
  $('#propList').innerHTML = `<dt>Titel</dt><dd>${SW.esc(SW.doc.title)}</dd><dt>Seiten</dt><dd>${c.pages}</dd><dt>Wörter</dt><dd>${c.words.toLocaleString('de-DE')}</dd><dt>Format</dt><dd>${SW.doc.pageSize}, ${SW.doc.orientation === 'landscape' ? 'Quer' : 'Hoch'} (${SW.fmtNum(g.wCm)} × ${SW.fmtNum(g.hCm)} cm)</dd><dt>Erstellt</dt><dd>${new Date(SW.doc.created).toLocaleString('de-DE')}</dd><dt>Speicherort</dt><dd>${SW.fileHandle ? 'Datei auf diesem Computer' : SW.doc.fileName ? 'Downloads-Ordner (zuletzt gespeichert)' : 'Noch nicht gespeichert'}</dd>`;
  $('#verInfo').textContent = `Version ${VERSION}`;
}
const openDrop = $('#openDrop');
openDrop.addEventListener('click', e => { if (!e.target.closest('button')) $('#fileOpen').click(); });
openDrop.addEventListener('dragover', e => { e.preventDefault(); openDrop.classList.add('over'); });
openDrop.addEventListener('dragleave', () => openDrop.classList.remove('over'));
openDrop.addEventListener('drop', e => { e.preventDefault(); e.stopPropagation(); openDrop.classList.remove('over'); const f = e.dataTransfer.files[0]; if (f) SW.openFile(f); });

function confirmDiscard(){
  if (!SW.dirty) return Promise.resolve(true);
  return new Promise(res => {
    SW.dialog({
      title:'Schreibwerk', width:420,
      body:`<p style="font-size:13px">Möchten Sie die Änderungen an „${SW.esc(SW.doc.title)}“ speichern?</p>`,
      buttons:[
        {label:'Abbrechen', action:() => res(false)},
        {label:'Nicht speichern', action:() => res(true)},
        {label:'Speichern', primary:true, action:() => { SW.save().then(ok => res(!!ok)); }}
      ]
    }).el.querySelector('.dlg-x').addEventListener('click', () => res(false));
  });
}

function loadDocument(html, doc, images){
  SW.images.forEach(r => URL.revokeObjectURL(r.url));
  SW.images.clear();
  (images || []).forEach(im => SW.registerImage(im.blob, im.id));
  SW.doc = Object.assign(SW.defaultDoc(), doc || {});
  SW.doc.base = Object.assign(SW.defaultDoc().base, (doc && doc.base) || {});
  ed.innerHTML = html || '<p><br></p>';
  SW.fixImages();
  SW.normalizeDoc();
  // Tabellen ohne Spaltengruppe ergänzen
  $$('table', ed).forEach(t => { if (!t.querySelector(':scope > colgroup')){ const n = Math.max(1, ...Array.from(t.rows).map(r => Array.from(r.cells).reduce((s, c) => s + (c.colSpan || 1), 0))); setCols(t, n); } });
  SW.applyBase();
  SW.dirty = false;
  SW.clearFind();
  selectImage(null);
  SW.relayout();
  SW.resetHistory();
  SW.updateTitle();
  canvas.scrollTop = 0;
  const first = ed.querySelector('p,h1,h2,h3,li,td') || ed;
  const r = document.createRange(); r.selectNodeContents(first); r.collapse(true);
  ed.focus({preventScroll:true});
  const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  refreshUI();
}

SW.newDoc = async function(tpl){
  if (!(await confirmDiscard())) return;
  closeBackstage();
  const t = TEMPLATES[tpl] || TEMPLATES.blank;
  SW.fileHandle = null;
  const doc = Object.assign(SW.defaultDoc(), t.doc ? t.doc() : {});
  doc.title = nextDocName();
  loadDocument(t.html(), doc, []);
};
let docCounter = 1;
function nextDocName(){ return 'Dokument' + (++docCounter); }

SW.openFile = async function(file){
  if (!(await confirmDiscard())) return;
  closeBackstage();
  try {
    const name = file.name.replace(/\.(docx|txt|html?)$/i, '');
    if (/\.docx$/i.test(file.name)){
      SW.toast('Datei wird geöffnet …');
      const res = await SW.importDocx(await file.arrayBuffer());
      res.doc = res.doc || {};
      res.doc.title = name; res.doc.fileName = file.name;
      loadDocument(res.html, res.doc, res.images);
      SW.toast(res.exact ? `„${file.name}“ geöffnet.` : `„${file.name}“ geöffnet. Hinweis: Nicht unterstützte Word-Elemente (z. B. Formen, Kommentare, Spalten) wurden vereinfacht.`);
    } else if (/\.html?$/i.test(file.name)){
      const html = SW.sanitize(await file.text());
      loadDocument(html, {title:name, fileName:name + '.docx'}, []);
      $$('img:not([data-img-id])', ed).forEach(img => { const b = dataUrlBlob(img.src); if (b){ const {id, url} = SW.registerImage(b); img.dataset.imgId = id; img.src = url; } });
    } else {
      loadDocument(SW.importText(await file.text()), {title:name, fileName:name + '.docx'}, []);
    }
    SW.fileHandle = file.handle || null;
    SW.updateTitle();
  } catch(err){
    console.error(err);
    SW.toast('Die Datei konnte nicht geöffnet werden: ' + err.message, null, 6000);
  }
};
$('#fileOpen').addEventListener('change', async e => {
  const f = e.target.files[0]; e.target.value = '';
  if (f) SW.openFile(f);
});
$('#fileImage').addEventListener('change', async e => {
  const files = Array.from(e.target.files); e.target.value = '';
  SW.focusEditor();
  for (const f of files) await SW.insertImageFile(f);
});
// Öffnen per Dateiauswahl des Systems (sofern vom Browser unterstützt)
C.open = async () => {
  closeBackstage();
  if (window.showOpenFilePicker){
    try {
      const [h] = await showOpenFilePicker({types:[{description:'Word-Dokument', accept:{'application/vnd.openxmlformats-officedocument.wordprocessingml.document':['.docx']}}, {description:'Text', accept:{'text/plain':['.txt'], 'text/html':['.html','.htm']}}]});
      const f = await h.getFile();
      f.handle = /\.docx$/i.test(f.name) ? h : null;
      SW.openFile(f);
      return;
    } catch(e){ if (e.name === 'AbortError') return; }
  }
  $('#fileOpen').click();
};

function safeName(n){ return (n || 'Dokument').replace(/[\\/:*?"<>|]+/g, '_').trim() || 'Dokument'; }
async function writeBlob(blob){
  if (SW.fileHandle && SW.fileHandle.createWritable){
    const w = await SW.fileHandle.createWritable();
    await w.write(blob); await w.close();
    return 'handle';
  }
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = SW.doc.fileName;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  return 'download';
}
SW.save = async function(){
  if (!SW.doc.fileName && !SW.fileHandle) return SW.saveAs();
  try {
    const blob = await SW.buildDocx();
    const how = await writeBlob(blob);
    afterSave(how);
    return true;
  } catch(e){
    console.error(e);
    SW.toast('Speichern fehlgeschlagen: ' + e.message, null, 6000);
    return false;
  }
};
SW.saveAs = async function(){
  closeBackstage();
  if (window.showSaveFilePicker){
    try {
      const h = await showSaveFilePicker({suggestedName:safeName(SW.doc.title) + '.docx', types:[{description:'Word-Dokument', accept:{'application/vnd.openxmlformats-officedocument.wordprocessingml.document':['.docx']}}]});
      SW.fileHandle = h;
      SW.doc.fileName = h.name;
      SW.doc.title = h.name.replace(/\.docx$/i, '');
      return await SW.save();
    } catch(e){ if (e.name === 'AbortError') return false; }
  }
  return new Promise(res => {
    SW.dialog({
      title:'Speichern unter', width:420,
      body:`<div class="dlg-row"><label style="min-width:80px">Dateiname</label><input id="sN" type="text" value="${SW.esc(SW.doc.title)}"><span>.docx</span></div><p class="muted">Die Datei wird über den Browser gespeichert (normalerweise im Ordner „Downloads“). Tipp: In den Browser-Einstellungen „Vor dem Download nach Speicherort fragen“ aktivieren, um den Ordner selbst zu wählen.</p>`,
      buttons:[{label:'Abbrechen', action:() => res(false)}, {label:'Speichern', primary:true, action:d => {
        const n = safeName(d.querySelector('#sN').value);
        SW.doc.title = n; SW.doc.fileName = n + '.docx';
        SW.fileHandle = null;
        SW.save().then(res);
      }}]
    });
  });
};
function afterSave(how){
  SW.dirty = false;
  SW.updateTitle();
  const t = new Date().toLocaleTimeString('de-DE', {hour:'2-digit', minute:'2-digit'});
  $('#stSaved').textContent = `Gespeichert ${t}`;
  SW.toast(how === 'handle' ? `„${SW.doc.fileName}“ gespeichert.` : `„${SW.doc.fileName}“ gespeichert (Downloads).`);
  saveDraft();
}
SW.preparePrint = function(){
  SW.flushCommit();
  selectImage(null);
  SW.clearFind();
  const oldZoom = SW.zoom;
  if (oldZoom !== 1) SW.setZoom(1); else SW.relayout();
  const g = SW.geo(), n = SW.pageCount || 1, step = g.H + 24;
  // Jede Seite als eigener, exakt zugeschnittener Block – druckt genau wie die Bildschirmansicht
  let st = document.getElementById('swPrintStyle');
  if (!st){ st = document.createElement('style'); st.id = 'swPrintStyle'; document.head.appendChild(st); }
  st.textContent = `@page{size:${g.wCm}cm ${g.hCm}cm;margin:0}
@media print{body > *:not(#printRoot){display:none !important}#printRoot{display:block !important}}
#printRoot{display:none}
#printRoot .pp{position:relative;overflow:hidden;width:${g.wCm}cm;height:calc(${g.hCm}cm - 1px);break-after:page;background:#fff}
#printRoot .pp:last-child{break-after:auto}
#printRoot .pp > .pin{position:absolute;left:0;width:${g.W}px}
#printRoot #pages .page{box-shadow:none}
#printRoot .hf{border:0 !important}#printRoot .hf.empty::before{content:none !important}
#printRoot .sw-toc:hover{outline:0}`;
  const root = document.createElement('div');
  root.id = 'printRoot';
  const src = document.getElementById('paper');
  for (let i = 0; i < n; i++){
    const pp = document.createElement('div'); pp.className = 'pp';
    const pin = document.createElement('div'); pin.className = 'pin';
    pin.style.top = (-i * step) + 'px';
    const c = src.cloneNode(true);
    c.removeAttribute('id'); c.style.transform = 'none'; c.style.position = 'relative';
    c.querySelectorAll('[id]').forEach(x => { if (x.id !== 'pages') x.removeAttribute('id'); });
    c.querySelectorAll('[contenteditable]').forEach(x => x.removeAttribute('contenteditable'));
    const edc = c.querySelector('[spellcheck]'); if (edc) edc.id = 'editor';
    pin.appendChild(c); pp.appendChild(pin); root.appendChild(pp);
  }
  root.querySelectorAll('.pp > .pin > div > div').forEach(x => { if (x.id !== 'pages' && !x.matches('[spellcheck]')) x.remove(); });
  document.body.appendChild(root);
  const oldTitle = document.title;
  document.title = SW.doc.title;
  return () => {
    root.remove();
    document.title = oldTitle;
    if (oldZoom !== 1) SW.setZoom(oldZoom);
    SW.updateTitle();
  };
};
SW.printDoc = function(pdf){
  const cleanup = SW.preparePrint();
  if (pdf) SW.toast('Wählen Sie im Druckdialog als Ziel „Als PDF speichern“.', null, 5000);
  setTimeout(() => { window.print(); cleanup(); }, 60);
};

/* ---------------- Automatische Sicherung (nur lokal) ---------------- */
let db = null;
function idb(){
  if (db) return Promise.resolve(db);
  return new Promise((res, rej) => {
    try {
      const r = indexedDB.open('schreibwerk', 1);
      r.onupgradeneeded = () => r.result.createObjectStore('drafts');
      r.onsuccess = () => { db = r.result; res(db); };
      r.onerror = () => rej(r.error);
    } catch(e){ rej(e); }
  });
}
async function saveDraft(){
  try {
    const d = await idb();
    const clean = SW.cleanClone();
    const ids = new Set($$('img[data-img-id]', clean).map(i => i.dataset.imgId));
    const images = [...ids].map(id => SW.images.get(id)).filter(Boolean).map((r, i) => ({id:[...ids][i], blob:r.blob}));
    clean.querySelectorAll('img[data-img-id]').forEach(i => i.removeAttribute('src'));
    const rec = {html:clean.innerHTML, doc:JSON.parse(JSON.stringify(SW.doc)), images, savedAt:Date.now(), words:SW.countWords().words, dirty:SW.dirty};
    d.transaction('drafts', 'readwrite').objectStore('drafts').put(rec, 'current');
  } catch(e){ /* Browser-Speicher nicht verfügbar */ }
}
let autoT = null;
SW.scheduleAutosave = () => { clearTimeout(autoT); autoT = setTimeout(saveDraft, 2500); };
async function readDraft(){
  try {
    const d = await idb();
    return await new Promise(res => { const r = d.transaction('drafts').objectStore('drafts').get('current'); r.onsuccess = () => res(r.result); r.onerror = () => res(null); });
  } catch(e){ return null; }
}
async function checkRestore(){
  const rec = await readDraft();
  const box = $('#restoreBox');
  if (!rec || !rec.html){ box.hidden = true; return; }
  box.hidden = false;
  const b = $('#restoreBtn');
  b.textContent = `${rec.doc.title} – ${new Date(rec.savedAt).toLocaleString('de-DE')} (${rec.words} Wörter)`;
  b.onclick = () => restoreDraft(rec);
}
async function restoreDraft(rec){
  if (!(await confirmDiscard())) return;
  closeBackstage();
  loadDocument(rec.html, rec.doc, rec.images);
  SW.dirty = !!rec.dirty; SW.updateTitle();
  SW.toast('Entwurf wiederhergestellt.');
}
window.addEventListener('beforeunload', e => {
  saveDraft();
  if (SW.dirty){ e.preventDefault(); e.returnValue = ''; }
});

/* ---------------- Vorlagen ---------------- */
const today = () => new Date().toLocaleDateString('de-DE', {day:'numeric', month:'long', year:'numeric'});
const TEMPLATES = {
  blank:{html:() => '<p><br></p>'},
  brief:{
    doc:() => ({margins:{...SW.MARGIN_PRESETS.din}}),
    html:() => `<p class="sw-nospace"><span style="font-size:8pt">Vorname Nachname · Straße 1 · 96047 Bamberg</span></p>
<p class="sw-nospace"><br></p><p class="sw-nospace">Firma Beispiel GmbH</p><p class="sw-nospace">Frau Erika Mustermann</p><p class="sw-nospace">Musterstraße 12</p><p class="sw-nospace">12345 Musterstadt</p>
<p class="sw-nospace"><br></p><p class="sw-nospace"><br></p><p style="text-align:right">Bamberg, ${today()}</p>
<p><br></p><p><b>Betreff des Schreibens</b></p><p><br></p><p>Sehr geehrte Frau Mustermann,</p>
<p>hier steht der Text Ihres Briefes. Klicken Sie einfach hinein und beginnen Sie zu schreiben.</p>
<p>Mit freundlichen Grüßen</p><p><br></p><p>Vorname Nachname</p>`
  },
  bericht:{
    doc:() => ({footer:{text:'Seite {SEITE} von {SEITEN}', align:'center'}}),
    html:() => `<p class="sw-title">Titel des Berichts</p><p class="sw-subtitle">Untertitel · ${today()}</p>
<div class="sw-toc" contenteditable="false"></div>
<h1>Einleitung</h1><p>Beschreiben Sie hier Ausgangslage und Ziel.</p>
<h1>Hauptteil</h1><h2>Abschnitt 1</h2><p>Text zum ersten Abschnitt.</p><h2>Abschnitt 2</h2><p>Text zum zweiten Abschnitt.</p>
<h1>Fazit</h1><p>Fassen Sie die Ergebnisse zusammen.</p>`
  },
  protokoll:{
    doc:() => ({footer:{text:'Seite {SEITE}', align:'right'}}),
    html:() => `<p class="sw-title">Protokoll</p><p class="sw-subtitle">Besprechung · ${today()}</p>
<table><colgroup><col style="width:30%"><col style="width:70%"></colgroup><tbody>
<tr><td><p><b>Datum / Uhrzeit</b></p></td><td><p>${new Date().toLocaleDateString('de-DE')}, 10:00–11:00 Uhr</p></td></tr>
<tr><td><p><b>Ort</b></p></td><td><p><br></p></td></tr><tr><td><p><b>Teilnehmende</b></p></td><td><p><br></p></td></tr><tr><td><p><b>Protokoll</b></p></td><td><p><br></p></td></tr></tbody></table>
<h1>Tagesordnung</h1><ol data-fmt="decimal"><li>Begrüßung</li><li>Thema 1</li><li>Thema 2</li><li>Verschiedenes</li></ol>
<h1>Ergebnisse</h1><p>Wichtigste Beschlüsse und Erkenntnisse.</p>
<h1>Aufgaben</h1>
<table class="sw-headrow"><colgroup><col style="width:55%"><col style="width:25%"><col style="width:20%"></colgroup><tbody>
<tr><td><p>Was</p></td><td><p>Wer</p></td><td><p>Bis wann</p></td></tr><tr><td><p><br></p></td><td><p><br></p></td><td><p><br></p></td></tr><tr><td><p><br></p></td><td><p><br></p></td><td><p><br></p></td></tr></tbody></table>`
  }
};

/* ---------------- Start ---------------- */
paintIcons();
try { if (localStorage.getItem('sw-dark')) document.documentElement.classList.add('dark'); } catch(e){}
document.execCommand('defaultParagraphSeparator', false, 'p');
SW.applyBase();
SW.relayout();
SW.resetHistory();
SW.updateTitle();
ed.focus();
refreshUI();
// Entwurf aus letzter Sitzung anbieten
readDraft().then(rec => {
  if (rec && rec.html && rec.dirty && rec.words > 0){
    SW.toast(`Ungespeicherter Entwurf „${rec.doc.title}“ vom ${new Date(rec.savedAt).toLocaleString('de-DE')} gefunden.`, {label:'Wiederherstellen', run:() => restoreDraft(rec)});
  }
});
// Start mit Zoom „Seitenbreite“ bei kleinen Fenstern
if (canvas.clientWidth < SW.geo().W + 60) SW.zoomFit('width');
})();
