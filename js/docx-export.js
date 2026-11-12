/* ==========================================================
   Schreibwerk – Export als Word-Dokument (.docx, Office Open XML)
   ========================================================== */
(function(){
'use strict';
const SW = window.SW, CM = SW.CM;

const TW_PER_CM = 1440 / 2.54;              // Twips pro cm
const pxToTw = px => Math.round(px * 15);   // 1 px = 15 Twips (96 dpi)
const cmToTw = cm => Math.round(cm * TW_PER_CM);
const ptToTw = pt => Math.round(pt * 20);
const pxToEmu = px => Math.round(px * 9525);
const x = s => String(s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DOC_NS = `xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="${NS_R}" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="${NS_W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" mc:Ignorable="w14"`;

/* Word-Standard-Hervorhebungsfarben */
const HIGHLIGHT = {FFFF00:'yellow','00FF00':'green','00FFFF':'cyan',FF00FF:'magenta','0000FF':'blue',FF0000:'red','000080':'darkBlue','008080':'darkCyan','008000':'darkGreen','800080':'darkMagenta','800000':'darkRed','808000':'darkYellow','808080':'darkGray',C0C0C0:'lightGray','000000':'black'};

const STYLE_MAP = {h1:'Heading1', h2:'Heading2', h3:'Heading3', h4:'Heading3', h5:'Heading3', h6:'Heading3', title:'Title', subtitle:'Subtitle', quote:'Quote', nospace:'NoSpacing'};

/* ---------- Kontext ---------- */
function makeCtx(){
  return {
    rels: [],            // {id, type, target, mode}
    media: [],           // {name, blob}
    nums: [],            // {numId, abs}
    abstracts: [],       // {id, kind, fmt}
    footnotes: [],
    docPrId: 1,
    relSeq: 10,
    imgSeq: 1,
    rel(type, target, mode){
      const id = 'rId' + (this.relSeq++);
      this.rels.push({id, type, target, mode});
      return id;
    }
  };
}

/* ---------- Hilfen ---------- */
function cssSize(v){
  // liefert pt
  if (!v) return null;
  const n = parseFloat(v);
  if (isNaN(n)){
    const kw = {'xx-small':7,'x-small':7.5,small:10,medium:12,large:13.5,'x-large':18,'xx-large':24,'xxx-large':36};
    return kw[v] || null;
  }
  if (v.endsWith('pt')) return n;
  if (v.endsWith('px')) return n * 0.75;
  if (v.endsWith('em')) return n * (SW.doc.base.size);
  if (v.endsWith('cm')) return n * 72 / 2.54;
  return n * 0.75;
}
function lenToTw(v){
  if (!v) return 0;
  const n = parseFloat(v);
  if (isNaN(n)) return 0;
  if (v.endsWith('cm')) return cmToTw(n);
  if (v.endsWith('mm')) return cmToTw(n / 10);
  if (v.endsWith('pt')) return ptToTw(n);
  if (v.endsWith('in')) return Math.round(n * 1440);
  return pxToTw(n);
}
function origMarginTop(el){ return el.style.marginTop; }
function cleanText(t){
  return t.replace(/​/g, '').replace(/[\r\n]+/g, ' ').replace(/ {2,}/g, ' ').replace(/ /g, ' ');
}

/* ---------- Absatzeigenschaften ---------- */
function pPr(el, ctx, opts={}){
  let s = '';
  const st = opts.style || (el ? styleId(el) : null);
  if (st) s += `<w:pStyle w:val="${st}"/>`;
  if (opts.keepNext) s += '<w:keepNext/>';
  if (opts.pageBreakBefore) s += '<w:pageBreakBefore/>';
  if (opts.num) s += `<w:numPr><w:ilvl w:val="${opts.num.lvl}"/><w:numId w:val="${opts.num.id}"/></w:numPr>`;
  if (el){
    const css = el.style;
    // Rahmen
    let bdr = '';
    const side = (name, v) => {
      if (!v || v === 'none' || /^0/.test(v)) return '';
      const m = v.match(/([\d.]+)(pt|px)?\s+(solid|dashed|dotted|double)?\s*(.*)/);
      let sz = 4, col = '000000';
      if (m){ sz = Math.max(2, Math.round((m[2] === 'px' ? +m[1]*0.75 : +m[1]) * 8)); col = SW.rgbToHex(m[4]) || '000000'; }
      return `<w:${name} w:val="single" w:sz="${sz}" w:space="1" w:color="${col}"/>`;
    };
    if (el.classList && el.classList.contains('sw-box')) bdr = ['top','left','bottom','right'].map(n => `<w:${n} w:val="single" w:sz="4" w:space="4" w:color="000000"/>`).join('');
    else bdr = side('top', css.borderTop) + side('bottom', css.borderBottom);
    if (bdr) s += `<w:pBdr>${bdr}</w:pBdr>`;
    const bg = SW.rgbToHex(css.backgroundColor);
    if (bg) s += `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`;
    // Abstände
    const sp = [];
    const mt = origMarginTop(el), mb = css.marginBottom;
    if (mt) sp.push(`w:before="${lenToTw(mt)}"`);
    if (mb) sp.push(`w:after="${lenToTw(mb)}"`);
    if (el.dataset && el.dataset.ls) sp.push(`w:line="${Math.round(parseFloat(el.dataset.ls) * 240)}" w:lineRule="auto"`);
    if (sp.length) s += `<w:spacing ${sp.join(' ')}/>`;
    // Einzüge
    const ind = [];
    const ml = lenToTw(css.marginLeft), mr = lenToTw(css.marginRight), ti = lenToTw(css.textIndent);
    if (opts.num){
      // Listen: Einzug kommt aus der Nummerierung
    } else {
      if (ml) ind.push(`w:left="${ml}"`);
      if (mr) ind.push(`w:right="${mr}"`);
      if (ti > 0) ind.push(`w:firstLine="${ti}"`);
      if (ti < 0) ind.push(`w:hanging="${-ti}"`);
    }
    if (ind.length) s += `<w:ind ${ind.join(' ')}/>`;
    // Ausrichtung
    const a = css.textAlign || (opts.inheritAlign || '');
    const jc = {center:'center', right:'right', end:'right', justify:'both', left:'left', start:'left'}[a];
    if (jc && jc !== 'left') s += `<w:jc w:val="${jc}"/>`;
  } else if (opts.inheritAlign){
    const jc = {center:'center', right:'right', justify:'both'}[opts.inheritAlign];
    if (jc) s += `<w:jc w:val="${jc}"/>`;
  }
  return s ? `<w:pPr>${s}</w:pPr>` : '';
}
function styleId(el){
  if (!el || !el.tagName) return null;
  const t = el.tagName.toLowerCase();
  if (STYLE_MAP[t]) return STYLE_MAP[t];
  if (el.classList.contains('sw-title')) return 'Title';
  if (el.classList.contains('sw-subtitle')) return 'Subtitle';
  if (el.classList.contains('sw-quote')) return 'Quote';
  if (el.classList.contains('sw-nospace')) return 'NoSpacing';
  return null;
}

/* ---------- Läufe (Text) ---------- */
function rPr(p){
  let s = '';
  if (p.rStyle) s += `<w:rStyle w:val="${p.rStyle}"/>`;
  if (p.font) s += `<w:rFonts w:ascii="${x(p.font)}" w:hAnsi="${x(p.font)}" w:cs="${x(p.font)}"/>`;
  if (p.b) s += '<w:b/><w:bCs/>';
  if (p.i) s += '<w:i/><w:iCs/>';
  if (p.strike) s += '<w:strike/>';
  if (p.color) s += `<w:color w:val="${p.color}"/>`;
  if (p.size) s += `<w:sz w:val="${Math.round(p.size*2)}"/><w:szCs w:val="${Math.round(p.size*2)}"/>`;
  if (p.bg && HIGHLIGHT[p.bg]) s += `<w:highlight w:val="${HIGHLIGHT[p.bg]}"/>`;
  if (p.u) s += '<w:u w:val="single"/>';
  if (p.bg && !HIGHLIGHT[p.bg]) s += `<w:shd w:val="clear" w:color="auto" w:fill="${p.bg}"/>`;
  if (p.va) s += `<w:vertAlign w:val="${p.va}"/>`;
  return s ? `<w:rPr>${s}</w:rPr>` : '';
}
function textRun(text, p){
  text = cleanText(text);
  if (!text) return '';
  const parts = text.split('\t');
  let out = '';
  parts.forEach((t, i) => {
    if (i > 0) out += `<w:r>${rPr(p)}<w:tab/></w:r>`;
    if (t) out += `<w:r>${rPr(p)}<w:t xml:space="preserve">${x(t)}</w:t></w:r>`;
  });
  return out;
}
function inheritProps(el, p){
  const q = {...p};
  const t = el.tagName;
  if (t === 'B' || t === 'STRONG') q.b = true;
  if (t === 'I' || t === 'EM') q.i = true;
  if (t === 'U' || t === 'INS') q.u = true;
  if (t === 'S' || t === 'STRIKE' || t === 'DEL') q.strike = true;
  if (t === 'SUB') q.va = 'subscript';
  if (t === 'SUP') q.va = 'superscript';
  const cs = el.style;
  if (cs){
    const fw = cs.fontWeight;
    if (fw === 'bold' || +fw >= 600) q.b = true;
    if (fw === 'normal' || (+fw && +fw < 600)) q.b = false;
    if (cs.fontStyle === 'italic') q.i = true;
    if (cs.fontStyle === 'normal') q.i = false;
    const td = cs.textDecorationLine || cs.textDecoration;
    if (td){
      if (td.includes('underline')) q.u = true;
      if (td.includes('line-through')) q.strike = true;
      if (td === 'none') { q.u = false; q.strike = false; }
    }
    if (cs.color){ const c = SW.rgbToHex(cs.color); if (c) q.color = c; }
    if (cs.backgroundColor){ const c = SW.rgbToHex(cs.backgroundColor); q.bg = c || undefined; }
    if (cs.fontSize){ const s = cssSize(cs.fontSize); if (s) q.size = s; }
    if (cs.fontFamily){ q.font = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim(); }
    if (cs.verticalAlign === 'super') q.va = 'superscript';
    if (cs.verticalAlign === 'sub') q.va = 'subscript';
  }
  return q;
}

function inlineXml(node, p, ctx){
  let out = '';
  for (const n of node.childNodes){
    if (n.nodeType === 3){ out += textRun(n.data, p); continue; }
    if (n.nodeType !== 1) continue;
    if (n.classList.contains('sw-gap')) continue;
    const t = n.tagName;
    if (t === 'BR'){ if (n.nextSibling || node.childNodes.length > 1 && n.previousSibling) out += `<w:r>${rPr(p)}<w:br/></w:r>`; continue; }
    if (t === 'IMG'){ out += imageRun(n, ctx); continue; }
    if (n.classList.contains('sw-tab')){ out += `<w:r>${rPr(p)}<w:tab/></w:r>`; continue; }
    if (n.classList.contains('sw-fn')){ out += footnoteRef(n, ctx); continue; }
    if (t === 'A' && n.getAttribute('href')){
      const href = n.getAttribute('href');
      const q = inheritProps(n, {...p, rStyle:'Hyperlink'});
      if (href.startsWith('#')) out += `<w:hyperlink w:anchor="${x(href.slice(1))}">${inlineXml(n, q, ctx)}</w:hyperlink>`;
      else {
        const id = ctx.rel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', href, 'External');
        out += `<w:hyperlink r:id="${id}" w:history="1">${inlineXml(n, q, ctx)}</w:hyperlink>`;
      }
      continue;
    }
    if (/^(P|DIV|H\d|UL|OL|TABLE|LI)$/.test(t)) continue;   // Blöcke werden separat behandelt
    out += inlineXml(n, inheritProps(n, p), ctx);
  }
  return out;
}

/* ---------- Bilder ---------- */
function imageRun(img, ctx){
  const id = img.dataset.imgId;
  let blob = null, ext = 'png';
  const rec = id && SW.images.get(id);
  if (rec){ blob = rec.blob; }
  else if (img.src.startsWith('data:')) blob = dataUrlToBlob(img.src);
  if (!blob) return '';
  const type = blob.type || 'image/png';
  ext = type.includes('jpeg') ? 'jpeg' : type.includes('gif') ? 'gif' : 'png';
  const name = `image${ctx.imgSeq++}.${ext}`;
  ctx.media.push({name, blob, imgId:id});
  const rid = ctx.rel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', 'media/' + name);
  const w = img.offsetWidth || img.naturalWidth || 200;
  const h = img.offsetHeight || img.naturalHeight || 150;
  const cx = pxToEmu(w), cy = pxToEmu(h);
  const pid = ctx.docPrId++;
  return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${pid}" name="Bild ${pid}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${pid}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${img.classList.contains('sw-imgborder') ? '<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>' : ''}</pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}
function dataUrlToBlob(u){
  try {
    const [head, data] = u.split(',');
    const mime = (head.match(/data:([^;]+)/) || [])[1] || 'image/png';
    const bin = atob(data);
    const arr = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    return new Blob([arr], {type:mime});
  } catch(e){ return null; }
}

/* ---------- Fußnoten ---------- */
function footnoteRef(el, ctx){
  const id = ctx.footnotes.length + 1;
  ctx.footnotes.push({id, text:el.dataset.note || ''});
  return `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
}

/* ---------- Listen ---------- */
const BULLET_DEF = {
  disc:{t:'', f:'Symbol'}, circle:{t:'o', f:'Courier New'}, square:{t:'', f:'Wingdings'},
  dash:{t:'–', f:'Calibri'}, arrow:{t:'', f:'Wingdings'}, check:{t:'', f:'Wingdings'}
};
const NUM_DEF = {
  'decimal':{fmt:'decimal', txt:'.'}, 'decimal-paren':{fmt:'decimal', txt:')'},
  'lower-alpha':{fmt:'lowerLetter', txt:')'}, 'upper-alpha':{fmt:'upperLetter', txt:'.'},
  'lower-roman':{fmt:'lowerRoman', txt:'.'}, 'upper-roman':{fmt:'upperRoman', txt:'.'}
};
function listNum(list, ctx){
  const kind = list.tagName === 'UL' ? 'bullet' : 'num';
  const fmt = kind === 'bullet' ? (list.dataset.bullet || ({circle:'circle', square:'square'}[list.style.listStyleType]) || 'disc')
                                : (list.dataset.fmt || ({'upper-alpha':'upper-alpha','lower-alpha':'lower-alpha','lower-roman':'lower-roman','upper-roman':'upper-roman'}[list.style.listStyleType]) || 'decimal');
  let abs = ctx.abstracts.find(a => a.kind === kind && a.fmt === fmt);
  if (!abs){ abs = {id:ctx.abstracts.length, kind, fmt}; ctx.abstracts.push(abs); }
  const numId = ctx.nums.length + 1;
  const start = parseInt(list.getAttribute('start') || '1', 10) || 1;
  ctx.nums.push({numId, abs:abs.id, start});
  return numId;
}
function listXml(list, ctx, lvl){
  const numId = listNum(list, ctx);
  let out = '';
  for (const li of list.children){
    if (li.tagName === 'UL' || li.tagName === 'OL'){ out += listXml(li, ctx, Math.min(8, lvl + 1)); continue; }
    if (li.tagName !== 'LI') continue;
    // Inhalt des Listeneintrags: Inline-Läufe → Absätze, verschachtelte Listen → nächste Ebene
    let first = true;
    let buf = [];
    const para = (nodes, el, style) => {
      out += `<w:p>${pPr(el, ctx, {style, num:first ? {lvl, id:numId} : null})}${inlineXml({childNodes:nodes}, {}, ctx)}</w:p>`;
      first = false;
    };
    const flush = () => {
      if (buf.some(n => n.nodeType === 1 || n.data.trim())) para(buf, li, 'ListParagraph');
      buf = [];
    };
    for (const c of Array.from(li.childNodes)){
      if (c.nodeType === 1 && (c.tagName === 'UL' || c.tagName === 'OL')){
        flush();
        if (first) para([], li, 'ListParagraph');
        out += listXml(c, ctx, Math.min(8, lvl + 1));
        continue;
      }
      if (c.nodeType === 1 && /^(P|H\d|DIV)$/.test(c.tagName)){
        flush();
        para(Array.from(c.childNodes), c, styleId(c) || 'ListParagraph');
        continue;
      }
      buf.push(c);
    }
    flush();
    if (first) para([], li, 'ListParagraph');
  }
  return out;
}
/* Inline-Inhalt eines Elements ohne Block-Kinder */
function inlineXmlFiltered(el, ctx){
  return inlineXml(el, {}, ctx);
}

/* ---------- Tabellen ---------- */
function tableXml(tbl, ctx, maxW){
  const rows = Array.from(tbl.rows).filter(r => !r.classList.contains('sw-gaprow'));
  // Raster aufbauen
  const grid = [];
  rows.forEach((tr, ri) => {
    grid[ri] = grid[ri] || [];
    let ci = 0;
    for (const cell of tr.cells){
      while (grid[ri][ci]) ci++;
      const cs = cell.colSpan || 1, rs = cell.rowSpan || 1;
      for (let r = 0; r < rs; r++){
        grid[ri + r] = grid[ri + r] || [];
        for (let c = 0; c < cs; c++) grid[ri + r][ci + c] = {cell, start:r === 0 && c === 0, row0:ri, col0:ci, cont:r > 0 && c === 0};
      }
      ci += cs;
    }
  });
  const ncols = Math.max(1, ...grid.map(r => r.length));
  // Spaltenbreiten messen
  let widths = [];
  const cols = tbl.querySelectorAll(':scope > colgroup > col');
  if (cols.length === ncols) widths = Array.from(cols).map(c => c.getBoundingClientRect().width / SW.zoom);
  if (!widths.length || widths.some(w => !w)){
    const tw = tbl.getBoundingClientRect().width / SW.zoom || maxW;
    widths = Array(ncols).fill(tw / ncols);
  }
  const twW = widths.map(w => Math.max(100, pxToTw(w)));
  const noBorder = tbl.classList.contains('sw-noborder');
  const headRow = tbl.classList.contains('sw-headrow');
  let out = '<w:tbl><w:tblPr>';
  out += noBorder ? '<w:tblStyle w:val="TableNormal"/>' : '<w:tblStyle w:val="TableGrid"/>';
  out += `<w:tblW w:w="${twW.reduce((a, b) => a + b, 0)}" w:type="dxa"/>`;
  if (noBorder) out += '<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>';
  out += '<w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>';
  out += '<w:tblGrid>' + twW.map(w => `<w:gridCol w:w="${w}"/>`).join('') + '</w:tblGrid>';
  rows.forEach((tr, ri) => {
    out += '<w:tr>';
    if (ri === 0 && headRow) out += '<w:trPr><w:tblHeader/></w:trPr>';
    const g = grid[ri] || [];
    for (let ci = 0; ci < ncols;){
      const slot = g[ci];
      if (!slot){ out += `<w:tc><w:tcPr><w:tcW w:w="${twW[ci]}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`; ci++; continue; }
      const cell = slot.cell;
      const span = cell.colSpan || 1;
      const w = twW.slice(ci, ci + span).reduce((a, b) => a + b, 0);
      let tcPr = `<w:tcW w:w="${w}" w:type="dxa"/>`;
      if (span > 1) tcPr += `<w:gridSpan w:val="${span}"/>`;
      if ((cell.rowSpan || 1) > 1 || slot.cont) tcPr += slot.cont ? '<w:vMerge/>' : '<w:vMerge w:val="restart"/>';
      const bg = SW.rgbToHex(cell.style.backgroundColor) || (ri === 0 && headRow ? 'D9E2F3' : null);
      if (bg) tcPr += `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`;
      if (cell.style.verticalAlign === 'middle') tcPr += '<w:vAlign w:val="center"/>';
      if (cell.style.verticalAlign === 'bottom') tcPr += '<w:vAlign w:val="bottom"/>';
      let body = '';
      if (slot.cont) body = '<w:p/>';
      else {
        body = blocksXml(cell, ctx, {cell:true, bold:ri === 0 && headRow, maxW:w / 15});
        if (!body || !/<w:p[ >\/]/.test(body.slice(-400)) || body.endsWith('</w:tbl>')) body += '<w:p/>';
      }
      out += `<w:tc><w:tcPr>${tcPr}</w:tcPr>${body}</w:tc>`;
      ci += span;
    }
    out += '</w:tr>';
  });
  return out + '</w:tbl>';
}

/* ---------- Inhaltsverzeichnis (Feld) ---------- */
function tocXml(ctx){
  const ents = SW.tocEntries();
  const cw = SW.geo().cw;
  const tab = `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${pxToTw(cw) - 10}"/></w:tabs>`;
  let out = '<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t>Inhaltsverzeichnis</w:t></w:r></w:p>';
  const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>';
  const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
  if (!ents.length){
    return out + `<w:p>${begin}<w:r><w:t>Keine Einträge für das Inhaltsverzeichnis gefunden.</w:t></w:r>${end}</w:p>`;
  }
  ents.forEach((e, i) => {
    out += `<w:p><w:pPr><w:pStyle w:val="TOC${e.lvl}"/>${tab}</w:pPr>${i === 0 ? begin : ''}<w:r><w:t xml:space="preserve">${x(e.text)}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>${e.page}</w:t></w:r>${i === ents.length - 1 ? end : ''}</w:p>`;
  });
  return out;
}

/* ---------- Blöcke ---------- */
function blocksXml(root, ctx, opts={}){
  let out = '';
  let pendingBreak = false;
  const inlineBuf = [];
  const flushInline = () => {
    if (!inlineBuf.length) return;
    const xml = inlineXml({childNodes:inlineBuf.slice()}, opts.bold ? {b:true} : {}, ctx);
    inlineBuf.length = 0;
    if (xml.replace(/<w:r><w:t xml:space="preserve"> <\/w:t><\/w:r>/g, '').trim() || opts.cell)
      out += `<w:p>${pPr(null, ctx, {pageBreakBefore:pendingBreak, inheritAlign:root.style && root.style.textAlign})}${xml}</w:p>`;
    pendingBreak = false;
  };
  for (const n of Array.from(root.childNodes)){
    if (n.nodeType === 3){ if (n.data.trim()) inlineBuf.push(n); continue; }
    if (n.nodeType !== 1) continue;
    const t = n.tagName;
    if (n.classList.contains('sw-gap') || n.classList.contains('sw-gaprow')) continue;
    if (n.classList.contains('sw-pagebreak')){ flushInline(); pendingBreak = true; continue; }
    if (n.classList.contains('sw-toc')){ flushInline(); out += tocXml(ctx); continue; }
    if (t === 'P' || /^H[1-6]$/.test(t) || t === 'BLOCKQUOTE' || t === 'PRE'){
      flushInline();
      const bq = t === 'BLOCKQUOTE' ? 'Quote' : null;
      out += `<w:p>${pPr(n, ctx, {pageBreakBefore:pendingBreak, style:bq || undefined, keepNext:false})}${inlineXml(n, opts.bold ? {b:true} : {}, ctx)}</w:p>`;
      pendingBreak = false;
      continue;
    }
    if (t === 'UL' || t === 'OL'){
      flushInline();
      if (pendingBreak){ out += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'; pendingBreak = false; }
      out += listXml(n, ctx, 0);
      continue;
    }
    if (t === 'TABLE'){
      flushInline();
      if (pendingBreak){ out += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'; pendingBreak = false; }
      out += tableXml(n, ctx, opts.maxW || SW.geo().cw);
      continue;
    }
    if (t === 'HR'){
      flushInline();
      out += '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>';
      continue;
    }
    if (t === 'DIV'){
      flushInline();
      if (n.querySelector('p,h1,h2,h3,ul,ol,table,div')) out += blocksXml(n, ctx, opts);
      else out += `<w:p>${pPr(n, ctx, {pageBreakBefore:pendingBreak})}${inlineXml(n, {}, ctx)}</w:p>`;
      pendingBreak = false;
      continue;
    }
    if (t === 'COLGROUP') continue;
    inlineBuf.push(n);
  }
  flushInline();
  if (pendingBreak) out += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  return out;
}

/* ---------- Kopf-/Fußzeile ---------- */
function hfXml(tag, cfg){
  const jc = {center:'center', right:'right'}[cfg.align];
  let runs = '';
  (cfg.text || '').split(/(\{SEITE\}|\{SEITEN\})/).forEach(part => {
    if (part === '{SEITE}') runs += '<w:fldSimple w:instr=" PAGE   \\* MERGEFORMAT "><w:r><w:t>1</w:t></w:r></w:fldSimple>';
    else if (part === '{SEITEN}') runs += '<w:fldSimple w:instr=" NUMPAGES   \\* MERGEFORMAT "><w:r><w:t>1</w:t></w:r></w:fldSimple>';
    else if (part){
      part.split('\n').forEach((line, i) => {
        if (i > 0) runs += '<w:r><w:br/></w:r>';
        if (line) runs += `<w:r><w:t xml:space="preserve">${x(line)}</w:t></w:r>`;
      });
    }
  });
  const st = tag === 'hdr' ? 'Header' : 'Footer';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:${tag} ${DOC_NS}><w:p><w:pPr><w:pStyle w:val="${st}"/>${jc ? `<w:jc w:val="${jc}"/>` : ''}</w:pPr>${runs}</w:p></w:${tag}>`;
}

/* ---------- Formatvorlagen ---------- */
function stylesXml(){
  const b = SW.doc.base;
  const f = x(b.font);
  const cw = pxToTw(SW.geo().cw);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${NS_W}" xmlns:r="${NS_R}">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${f}" w:eastAsia="${f}" w:hAnsi="${f}" w:cs="Times New Roman"/><w:kern w:val="2"/><w:sz w:val="${Math.round(b.size*2)}"/><w:szCs w:val="${Math.round(b.size*2)}"/><w:lang w:val="de-DE" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="${ptToTw(b.after)}" w:line="${Math.round(b.line*240)}" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="0" w:defUnhideWhenUsed="0" w:defQFormat="0" w:count="376"/>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:color w:val="2F5496"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:color w:val="2F5496"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:color w:val="1F3763"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="11"/><w:qFormat/><w:rPr><w:color w:val="5A5A5A"/><w:spacing w:val="15"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="29"/><w:qFormat/><w:pPr><w:spacing w:before="200" w:after="160"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="404040"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="NoSpacing"><w:name w:val="No Spacing"/><w:uiPriority w:val="1"/><w:qFormat/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:tabs><w:tab w:val="center" w:pos="${Math.round(cw/2)}"/><w:tab w:val="right" w:pos="${cw}"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:tabs><w:tab w:val="center" w:pos="${Math.round(cw/2)}"/><w:tab w:val="right" w:pos="${cw}"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="100"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="TOC2"><w:name w:val="toc 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="100"/><w:ind w:left="220"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="TOC3"><w:name w:val="toc 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="100"/><w:ind w:left="440"/></w:pPr></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>
<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
</w:styles>`;
}

function numberingXml(ctx){
  const lvlXml = (abs, l) => {
    const left = 720 * (l + 1);
    if (abs.kind === 'bullet'){
      const d = BULLET_DEF[abs.fmt] || BULLET_DEF.disc;
      return `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${x(d.t)}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="${d.f}" w:hAnsi="${d.f}" w:hint="default"/></w:rPr></w:lvl>`;
    }
    const base = NUM_DEF[abs.fmt] || NUM_DEF.decimal;
    return `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="${base.fmt}"/><w:lvlText w:val="%${l + 1}${base.txt}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr></w:lvl>`;
  };
  let s = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="${NS_W}">`;
  ctx.abstracts.forEach(a => {
    s += `<w:abstractNum w:abstractNumId="${a.id}"><w:multiLevelType w:val="hybridMultilevel"/>`;
    for (let l = 0; l < 9; l++) s += lvlXml(a, l);
    s += '</w:abstractNum>';
  });
  ctx.nums.forEach(n => {
    s += `<w:num w:numId="${n.numId}"><w:abstractNumId w:val="${n.abs}"/>`;
    for (let l = 0; l < 9; l++) s += `<w:lvlOverride w:ilvl="${l}"><w:startOverride w:val="${n.start}"/></w:lvlOverride>`;
    s += '</w:num>';
  });
  return s + '</w:numbering>';
}

function footnotesXml(ctx){
  let s = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:footnotes ${DOC_NS}>`;
  s += '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>';
  s += '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';
  ctx.footnotes.forEach(f => {
    s += `<w:footnote w:id="${f.id}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> ${x(f.text)}</w:t></w:r></w:p></w:footnote>`;
  });
  return s + '</w:footnotes>';
}

function hash(str){
  let h = 0x811c9dc5;
  for (let i = 0; i < str.length; i++){ h ^= str.charCodeAt(i); h = Math.imul(h, 16777619); }
  return (h >>> 0).toString(16);
}
SW.hashStr = hash;

function b64utf8(s){
  const bytes = new TextEncoder().encode(s);
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}

/* ---------- Hauptfunktion ---------- */
SW.buildDocx = async function(){
  SW.relayout(true);
  const ctx = makeCtx();
  const D = SW.doc, g = SW.geo();
  const body = blocksXml(SW.ed, ctx, {});
  const hasHdr = !!D.header.text, hasFtr = !!D.footer.text;
  let hdrId = '', ftrId = '';
  if (hasHdr) hdrId = ctx.rel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/header', 'header1.xml');
  if (hasFtr) ftrId = ctx.rel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer', 'footer1.xml');
  const land = D.orientation === 'landscape';
  const sect = `<w:sectPr>${hasHdr ? `<w:headerReference w:type="default" r:id="${hdrId}"/>` : ''}${hasFtr ? `<w:footerReference w:type="default" r:id="${ftrId}"/>` : ''}<w:pgSz w:w="${cmToTw(g.wCm)}" w:h="${cmToTw(g.hCm)}"${land ? ' w:orient="landscape"' : ''}/><w:pgMar w:top="${cmToTw(D.margins.top)}" w:right="${cmToTw(D.margins.right)}" w:bottom="${cmToTw(D.margins.bottom)}" w:left="${cmToTw(D.margins.left)}" w:header="709" w:footer="709" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>`;
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${DOC_NS}><w:body>${body || '<w:p/>'}${sect}</w:body></w:document>`;

  const zip = new JSZip();
  const hasNum = ctx.nums.length > 0, hasFn = ctx.footnotes.length > 0;
  // Beziehungen des Dokuments
  let rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
  rels += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
  rels += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>';
  rels += '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>';
  if (hasNum) rels += '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>';
  if (hasFn) rels += '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>';
  rels += '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>';
  ctx.rels.forEach(r => { rels += `<Relationship Id="${r.id}" Type="${r.type}" Target="${x(r.target)}"${r.mode ? ` TargetMode="${r.mode}"` : ''}/>`; });
  rels += '</Relationships>';

  // Quellinhalt für verlustfreies Wiederöffnen in Schreibwerk
  const clean = SW.cleanClone();
  const imgMap = {};
  ctx.media.forEach(m => { if (m.imgId) imgMap[m.imgId] = 'word/media/' + m.name; });
  const src = {v:1, app:'Schreibwerk', hash:hash(documentXml), html:clean.innerHTML, doc:{pageSize:D.pageSize, orientation:D.orientation, margins:D.margins, header:D.header, footer:D.footer, base:D.base, author:D.author, created:D.created}, images:imgMap};

  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const exts = new Set(ctx.media.map(m => m.name.split('.').pop()));
  let ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>';
  exts.forEach(e => { ct += `<Default Extension="${e}" ContentType="image/${e}"/>`; });
  ct += '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>';
  ct += '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>';
  ct += '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>';
  ct += '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>';
  if (hasNum) ct += '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>';
  if (hasFn) ct += '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>';
  if (hasHdr) ct += '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>';
  if (hasFtr) ct += '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>';
  ct += '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>';
  ct += '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>';

  zip.file('[Content_Types].xml', ct);
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${x(D.title)}</dc:title><dc:creator>${x(D.author || '')}</dc:creator><cp:lastModifiedBy>${x(D.author || '')}</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${(D.created || now).replace(/\.\d+Z$/, 'Z')}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
  zip.file('docProps/app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Schreibwerk</Application><Pages>${SW.pageCount || 1}</Pages><Words>${SW.countWords ? SW.countWords().words : 0}</Words></Properties>`);
  zip.file('word/document.xml', documentXml);
  zip.file('word/_rels/document.xml.rels', rels);
  zip.file('word/styles.xml', stylesXml());
  zip.file('word/settings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:zoom w:percent="100"/><w:defaultTabStop w:val="708"/><w:hyphenationZone w:val="425"/><w:characterSpacingControl w:val="doNotCompress"/>${hasFn ? '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>' : ''}<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat><w:themeFontLang w:val="de-DE"/><w:decimalSymbol w:val=","/><w:listSeparator w:val=";"/></w:settings>`);
  zip.file('word/fontTable.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:fonts xmlns:w="${NS_W}"><w:font w:name="${x(D.base.font)}"/><w:font w:name="Calibri Light"/><w:font w:name="Symbol"/><w:font w:name="Wingdings"/><w:font w:name="Courier New"/></w:fonts>`);
  if (hasNum) zip.file('word/numbering.xml', numberingXml(ctx));
  if (hasFn) zip.file('word/footnotes.xml', footnotesXml(ctx));
  if (hasHdr) zip.file('word/header1.xml', hfXml('hdr', D.header));
  if (hasFtr) zip.file('word/footer1.xml', hfXml('ftr', D.footer));
  ctx.media.forEach(m => zip.file('word/media/' + m.name, m.blob));
  zip.file('customXml/item1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<schreibwerk xmlns="urn:schreibwerk:v1">${b64utf8(JSON.stringify(src))}</schreibwerk>`);
  zip.file('customXml/itemProps1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ds:datastoreItem ds:itemID="{8A1F3C52-6B0E-4C1D-9E2A-5D7B3F4E1A90}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="urn:schreibwerk:v1"/></ds:schemaRefs></ds:datastoreItem>');
  zip.file('customXml/_rels/item1.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>');
  return await zip.generateAsync({type:'blob', mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document', compression:'DEFLATE'});
};
})();
