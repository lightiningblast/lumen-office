/* Lumen Write - word processor UI */
(function () {
  'use strict';
  const { $, $$, toast, form, FileIO, esc, cssToHex, baseName } = Lumen;

  const doc = $('#doc'), sheet = $('#sheet'), wrap = $('#pagewrap'), zoomer = $('#zoomer'), stage = $('#stage');
  const hdr = $('#hdr'), ftr = $('#ftr');
  let lastHF = ftr;
  hdr.addEventListener('focus', () => { lastHF = hdr; }); ftr.addEventListener('focus', () => { lastHF = ftr; });
  const state = { handle: null, dirty: false, zoom: 1 };
  let page = { w: 12240, h: 15840, t: 1440, r: 1440, b: 1440, l: 1440 };
  const FILE_TYPES = [{ description: 'Word documents', accept: { [DocxIO.DOCX_MIME]: ['.docx'] } }];

  try { document.execCommand('defaultParagraphSeparator', false, 'p'); } catch (e) { /* ignore */ }
  try { document.execCommand('styleWithCSS', false, true); } catch (e) { /* ignore */ }

  /* ---------------- selection helpers ---------------- */
  let saved = null;
  document.addEventListener('selectionchange', () => {
    const s = getSelection();
    if (s.rangeCount && doc.contains(s.anchorNode)) {
      saved = s.getRangeAt(0).cloneRange();
      syncToolbar();
    }
  });
  function restore() {
    if (!saved) { doc.focus(); return; }
    const s = getSelection();
    if (!(s.rangeCount && doc.contains(s.anchorNode))) {
      doc.focus();
      s.removeAllRanges();
      s.addRange(saved);
    }
  }
  function exec(cmd, val) {
    restore();
    document.execCommand('styleWithCSS', false, true);
    document.execCommand(cmd, false, val);
    markDirty();
    syncToolbar();
  }
  function markDirty() {
    if (!state.dirty) { state.dirty = true; $('#dirty').textContent = '● unsaved'; }
    scheduleStats();
  }
  function markClean() { state.dirty = false; $('#dirty').textContent = ''; }

  /* keep focus in the editor when pressing ribbon buttons */
  $('.ribbon').addEventListener('mousedown', (e) => {
    const t = e.target;
    if (t.closest('select') || t.closest('input')) return;
    e.preventDefault();
  });

  $$('[data-cmd]').forEach(b => b.addEventListener('click', () => {
    const c = b.dataset.cmd;
    if (c === 'cut' || c === 'copy') { restore(); document.execCommand(c); return; }
    exec(c);
  }));

  /* ---------------- font controls ---------------- */
  const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

  function selectionIsCollapsed() { const s = getSelection(); return s.rangeCount && s.getRangeAt(0).collapsed; }

  function setFontSize(pt) {
    restore();
    const sel = getSelection();
    if (!sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    if (range.collapsed) {
      const sp = document.createElement('span');
      sp.style.fontSize = pt + 'pt';
      sp.textContent = '​';
      range.insertNode(sp);
      const r = document.createRange();
      r.setStart(sp.firstChild, 1);
      r.collapse(true);
      sel.removeAllRanges();
      sel.addRange(r);
      markDirty();
      return;
    }
    document.execCommand('styleWithCSS', false, true);
    document.execCommand('fontSize', false, '7');
    const made = [];
    doc.querySelectorAll('font[size="7"], span[style*="xxx-large"]').forEach(el => {
      let sp = el;
      if (el.tagName === 'FONT') {
        sp = document.createElement('span');
        while (el.firstChild) sp.appendChild(el.firstChild);
        el.replaceWith(sp);
      }
      sp.style.fontSize = pt + 'pt';
      sp.querySelectorAll('[style*="font-size"]').forEach(e => e.style.removeProperty('font-size'));
      made.push(sp);
    });
    if (made.length) {
      const r = document.createRange();
      r.setStartBefore(made[0]);
      r.setEndAfter(made[made.length - 1]);
      sel.removeAllRanges();
      sel.addRange(r);
    }
    markDirty();
    syncToolbar();
  }
  function currentSizePt() {
    const s = getSelection();
    if (!s.rangeCount) return 11;
    let n = s.anchorNode;
    if (n && n.nodeType === 3) n = n.parentNode;
    if (!n || !doc.contains(n)) return 11;
    return Math.round(parseFloat(getComputedStyle(n).fontSize) * 0.75 * 2) / 2;
  }
  $('#font-size').addEventListener('change', e => setFontSize(parseFloat(e.target.value)));
  $('#font-name').addEventListener('change', e => { exec('fontName', e.target.value); });
  $('#size-up').addEventListener('click', () => { const c = currentSizePt(); setFontSize(SIZES.find(s => s > c) || c + 2); });
  $('#size-down').addEventListener('click', () => { const c = currentSizePt(); setFontSize([...SIZES].reverse().find(s => s < c) || Math.max(1, c - 1)); });

  function bindColor(inputId, swId, applyId, cmd) {
    const inp = $(inputId), sw = $(swId);
    inp.addEventListener('input', () => { sw.style.background = inp.value; exec(cmd, inp.value); });
    $(applyId).addEventListener('click', () => exec(cmd, inp.value));
    inp.addEventListener('mousedown', () => { restore(); });
  }
  bindColor('#fcolor', '#sw-fcolor', '#apply-fcolor', 'foreColor');
  bindColor('#hcolor', '#sw-hcolor', '#apply-hcolor', 'hiliteColor');

  /* ---------------- paragraph controls ---------------- */
  const BLOCK_SEL = 'p,h1,h2,h3,h4,h5,h6,li,div:not(.pb),blockquote';
  function blocksInSelection() {
    const s = getSelection();
    if (!s.rangeCount) return [];
    const r = s.getRangeAt(0);
    const out = [];
    doc.querySelectorAll(BLOCK_SEL).forEach(el => {
      if (r.intersectsNode(el) && !el.querySelector('p,h1,h2,h3,h4,h5,h6,li,div,table,ul,ol,blockquote')) out.push(el);
    });
    if (!out.length) {
      let n = r.startContainer;
      while (n && n !== doc) { if (n.nodeType === 1 && /^(P|H[1-6]|LI|DIV)$/.test(n.tagName)) { out.push(n); break; } n = n.parentNode; }
    }
    return out;
  }
  const STYLES = { normal: ['p', ''], title: ['p', 'title'], subtitle: ['p', 'subtitle'], h1: ['h1', ''], h2: ['h2', ''], h3: ['h3', ''], h4: ['h4', ''], h5: ['h5', ''], h6: ['h6', ''], quote: ['p', 'quote'] };
  $('#block-style').addEventListener('change', e => {
    restore();
    const [tag, cls] = STYLES[e.target.value];
    document.execCommand('formatBlock', false, '<' + tag + '>');
    blocksInSelection().forEach(b => { b.className = cls; if (cls || tag !== 'p') b.style.removeProperty('margin-left'); });
    markDirty();
    syncToolbar();
  });
  $('#line-spacing').addEventListener('change', e => {
    restore();
    const v = e.target.value;
    blocksInSelection().forEach(b => { if (v) b.style.lineHeight = v; else b.style.removeProperty('line-height'); });
    e.target.value = '';
    markDirty();
  });
  function indent(dir) {
    restore();
    const bl = blocksInSelection();
    if (bl.some(b => b.tagName === 'LI')) { exec(dir > 0 ? 'indent' : 'outdent'); return; }
    bl.forEach(b => { const cur = parseFloat(b.style.marginLeft) || 0; const v = Math.max(0, cur + dir * 48); if (v) b.style.marginLeft = v + 'px'; else b.style.removeProperty('margin-left'); });
    markDirty();
  }
  $('#indent').addEventListener('click', () => indent(1));
  $('#outdent').addEventListener('click', () => indent(-1));

  /* ---------------- toolbar state ---------------- */
  function syncToolbar() {
    const q = (c) => { try { return document.queryCommandState(c); } catch (e) { return false; } };
    $$('[data-cmd]').forEach(b => {
      const c = b.dataset.cmd;
      if (['bold', 'italic', 'underline', 'strikeThrough', 'subscript', 'superscript', 'insertUnorderedList', 'insertOrderedList', 'justifyLeft', 'justifyCenter', 'justifyRight', 'justifyFull'].includes(c)) b.classList.toggle('active', q(c));
    });
    const s = getSelection();
    if (!s.rangeCount || !doc.contains(s.anchorNode)) return;
    let n = s.anchorNode;
    if (n.nodeType === 3) n = n.parentNode;
    const cs = getComputedStyle(n);
    const fam = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
    const fn = $('#font-name');
    if (![...fn.options].some(o => o.value === fam)) { const o = document.createElement('option'); o.textContent = fam; fn.appendChild(o); }
    fn.value = fam;
    const pt = Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2;
    const fs = $('#font-size');
    if (![...fs.options].some(o => parseFloat(o.value) === pt)) { const o = document.createElement('option'); o.textContent = pt; fs.appendChild(o); }
    fs.value = String(pt);
    let b = n;
    while (b && b !== doc && !/^(P|H[1-6]|LI|DIV|TD|TH)$/.test(b.tagName)) b = b.parentNode;
    let v = 'normal';
    if (b && b !== doc) {
      if (/^H[1-6]$/.test(b.tagName)) v = b.tagName.toLowerCase();
      else if (b.classList.contains('title')) v = 'title';
      else if (b.classList.contains('subtitle')) v = 'subtitle';
      else if (b.classList.contains('quote')) v = 'quote';
    }
    $('#block-style').value = v;
    const inTable = !!closestCell();
    Lumen.showTab('table', inTable);
  }

  /* ---------------- stats / pages / zoom ---------------- */
  let statsTimer = 0;
  function scheduleStats() { clearTimeout(statsTimer); statsTimer = setTimeout(updateStats, 250); }
  function updateStats() {
    if (typeof renumberFootnotes === 'function') renumberFootnotes();
    if (typeof renumberEndnotes === 'function' && typeof enlist !== 'undefined') renumberEndnotes();
    const text = doc.innerText.replace(/​/g, '');
    const words = (text.trim().match(/\S+/g) || []).length;
    $('#st-words').textContent = words + (words === 1 ? ' word' : ' words');
    $('#st-chars').textContent = text.replace(/\n/g, '').length + ' characters';
    const ph = page.h / 15;
    const pages = Math.max(1, Math.ceil((sheet.offsetHeight - 2) / ph));
    $('#st-page').textContent = 'Page count: ' + pages;
  }
  const STYLE_SEL = { normal: '#doc', title: '#doc p.title', subtitle: '#doc p.subtitle', h1: '#doc h1', h2: '#doc h2', h3: '#doc h3', h4: '#doc h4', h5: '#doc h5', h6: '#doc h6', quote: '#doc p.quote' };
  const STYLE_NAMES = { normal: 'Normal', title: 'Title', subtitle: 'Subtitle', h1: 'Heading 1', h2: 'Heading 2', h3: 'Heading 3', h4: 'Heading 4', h5: 'Heading 5', h6: 'Heading 6', quote: 'Quote' };
  function applyStyleCss() {
    let el = document.getElementById('style-ov');
    if (!el) { el = document.createElement('style'); el.id = 'style-ov'; document.head.appendChild(el); }
    const ov = page.styles || {};
    el.textContent = Object.keys(ov).filter(k => STYLE_SEL[k] && ov[k]).map(k => {
      const o = ov[k];
      return `${STYLE_SEL[k]} { font-family: '${String(o.ff).replace(/['\\<]/g, '')}', 'Calibri', sans-serif; font-size: ${+o.sz || 11}pt; color: #${/^[0-9A-Fa-f]{6}$/.test(o.color) ? o.color : '000000'}; font-weight: ${o.b ? 'bold' : 'normal'}; font-style: ${o.i ? 'italic' : 'normal'}; }`;
    }).join('\n');
  }
  async function modifyStyleDialog() {
    const D = window.DocxStyleDefaults;
    let key = $('#block-style').value || 'normal';
    for (;;) {
      const cur = Object.assign({}, D[key], (page.styles || {})[key]);
      const v = await form('Modify style', [
        { id: 'key', label: 'Style', type: 'select', value: key, options: Object.keys(STYLE_NAMES).map(k => [k, STYLE_NAMES[k] + ((page.styles || {})[k] ? ' (modified)' : '')]) },
        { type: 'note', label: 'Showing the settings of "' + STYLE_NAMES[key] + '". Pick another style and press Apply to switch to it.' },
        { id: 'ff', label: 'Font', value: cur.ff },
        { id: 'sz', label: 'Size (pt)', type: 'number', value: cur.sz },
        { id: 'color', label: 'Colour (hex, e.g. 2F5496)', value: cur.color },
        { id: 'b', label: 'Bold', type: 'checkbox', value: !!cur.b },
        { id: 'i', label: 'Italic', type: 'checkbox', value: !!cur.i },
        { id: 'reset', label: 'Reset this style to the built-in look', type: 'checkbox', value: false },
      ], 'Apply');
      if (!v) return;
      if (v.key !== key) { key = v.key; continue; }
      setStyleOv(key, v.reset ? null : { ff: String(v.ff || '').trim() || D[key].ff, sz: Math.min(400, Math.max(1, +v.sz || D[key].sz)), color: String(v.color || '').replace('#', '').trim().toUpperCase(), b: !!v.b, i: !!v.i });
      return;
    }
  }
  function setStyleOv(key, o) {
    const D = window.DocxStyleDefaults;
    if (!D[key]) return;
    page.styles = Object.assign({}, page.styles);
    if (o && !/^[0-9A-F]{6}$/.test(o.color || '')) o.color = D[key].color;
    const d = D[key];
    if (!o || (o.ff === d.ff && +o.sz === d.sz && o.color === d.color && !!o.b === d.b && !!o.i === d.i)) delete page.styles[key]; else page.styles[key] = o;
    if (!Object.keys(page.styles).length) delete page.styles;
    applyStyleCss(); markDirty();
    toast('Style "' + STYLE_NAMES[key] + '" ' + (o ? 'updated' : 'reset'));
  }
  function applyPage() {
    applyStyleCss();
    const px = (v) => v / 15;
    sheet.style.width = px(page.w) + 'px';
    sheet.style.minHeight = px(page.h) + 'px';
    sheet.style.padding = `${px(page.t)}px ${px(page.r)}px ${px(page.b)}px ${px(page.l)}px`;
    sheet.style.setProperty('--ph', px(page.h) + 'px');
    doc.style.minHeight = Math.max(100, px(page.h - page.t - page.b)) + 'px';
    doc.style.columnCount = page.cols > 1 ? page.cols : 'auto';
    doc.style.columnGap = '48px';
    $('#lay-cols').value = String(page.cols > 1 ? page.cols : 1);
    for (const h of [hdr, ftr]) { h.style.left = px(page.l) + 'px'; h.style.right = px(page.r) + 'px'; }
    hdr.style.top = Math.max(8, px(page.t) / 2 - 10) + 'px';
    ftr.style.bottom = Math.max(8, px(page.b) / 2 - 10) + 'px';
    $('#print-style').textContent = `@page{size:${page.w / 1440}in ${page.h / 1440}in;margin:${page.t / 1440}in ${page.r / 1440}in ${page.b / 1440}in ${page.l / 1440}in}`;
    applyZoom(state.zoom);
  }
  function applyZoom(z) {
    state.zoom = Math.min(2.5, Math.max(0.4, z));
    wrap.style.transform = `scale(${state.zoom})`;
    fitZoomer();
    $('#zoom').value = Math.round(state.zoom * 100);
    $('#st-zoom').textContent = Math.round(state.zoom * 100) + '%';
    updateStats();
  }
  function fitZoomer() {
    zoomer.style.width = sheet.offsetWidth * state.zoom + 'px';
    zoomer.style.height = sheet.offsetHeight * state.zoom + 'px';
    positionFrame();
  }
  new ResizeObserver(() => { fitZoomer(); scheduleStats(); }).observe(sheet);
  $('#zoom').addEventListener('input', e => applyZoom(e.target.value / 100));
  $('#zoom-in').addEventListener('click', () => applyZoom(state.zoom + 0.1));
  $('#zoom-out').addEventListener('click', () => applyZoom(state.zoom - 0.1));
  $('#zoom-100').addEventListener('click', () => applyZoom(1));
  $('#zoom-fit').addEventListener('click', () => applyZoom((stage.clientWidth - 48) / sheet.offsetWidth));

  /* ---------------- layout ---------------- */
  $('#lay-margins').addEventListener('change', e => {
    const m = { normal: [1440, 1440], narrow: [720, 720], moderate: [1440, 1080], wide: [1440, 2880] }[e.target.value];
    if (m) { page.t = page.b = m[0]; page.l = page.r = m[1]; applyPage(); markDirty(); }
    e.target.value = '';
  });
  $('#lay-size').addEventListener('change', e => {
    const s = { letter: [12240, 15840], a4: [11906, 16838], legal: [12240, 20160], a5: [8391, 11906] }[e.target.value];
    if (s) { const land = page.w > page.h; page.w = land ? s[1] : s[0]; page.h = land ? s[0] : s[1]; applyPage(); markDirty(); }
    e.target.value = '';
  });
  $('#lay-orient').addEventListener('click', () => { [page.w, page.h] = [page.h, page.w]; applyPage(); markDirty(); });

  /* ---------------- tables ---------------- */
  function closestCell() {
    const s = getSelection();
    if (!s.rangeCount) return null;
    let n = s.anchorNode;
    while (n && n !== doc) { if (n.nodeType === 1 && (n.tagName === 'TD' || n.tagName === 'TH')) return n; n = n.parentNode; }
    return null;
  }
  function cellsInSelection() {
    const s = getSelection();
    if (!s.rangeCount) return [];
    const r = s.getRangeAt(0);
    const cells = [...doc.querySelectorAll('td,th')].filter(c => r.intersectsNode(c));
    const c = closestCell();
    return cells.length ? cells : (c ? [c] : []);
  }
  function resetTableWidths(table) {
    const cg = table.querySelector('colgroup');
    if (cg) cg.remove();
    table.querySelectorAll('td,th').forEach(c => c.style.removeProperty('width'));
    table.style.width = '100%';
  }
  function newCellLike(c) {
    const n = document.createElement('td');
    n.innerHTML = '<br>';
    if (c.style.backgroundColor && false) n.style.backgroundColor = c.style.backgroundColor;
    return n;
  }
  function tableOp(op) {
    restore();
    const cell = closestCell();
    if (!cell) return;
    const row = cell.parentElement, table = row.closest('table');
    const tr = [...table.rows];
    const colIdx = cell.cellIndex;
    if (op === 'row-above' || op === 'row-below') {
      const nr = document.createElement('tr');
      for (const c of row.cells) { const n = newCellLike(c); n.colSpan = c.colSpan; nr.appendChild(n); }
      if (op === 'row-above') row.before(nr); else row.after(nr);
    } else if (op === 'col-left' || op === 'col-right') {
      for (const r of tr) {
        const c = r.cells[Math.min(colIdx, r.cells.length - 1)];
        const n = newCellLike(c);
        if (op === 'col-left') c.before(n); else c.after(n);
      }
      resetTableWidths(table);
    } else if (op === 'del-row') {
      if (tr.length === 1) table.remove(); else row.remove();
    } else if (op === 'del-col') {
      if (row.cells.length === 1) table.remove();
      else { for (const r of tr) { const c = r.cells[Math.min(colIdx, r.cells.length - 1)]; if (c) c.remove(); } resetTableWidths(table); }
    } else if (op === 'del-table') table.remove();
    else if (op === 'borders') table.classList.toggle('nb');
    markDirty();
    syncToolbar();
  }
  for (const [id, op] of [['t-row-above', 'row-above'], ['t-row-below', 'row-below'], ['t-col-left', 'col-left'], ['t-col-right', 'col-right'], ['t-del-row', 'del-row'], ['t-del-col', 'del-col'], ['t-del-table', 'del-table'], ['t-borders', 'borders']]) {
    $('#' + id).addEventListener('click', () => tableOp(op));
  }
  function shadeCells(color) { restore(); cellsInSelection().forEach(c => { if (color) c.style.backgroundColor = color; else c.style.removeProperty('background-color'); }); markDirty(); }
  $('#shade').addEventListener('input', e => { $('#sw-shade').style.background = e.target.value; shadeCells(e.target.value); });
  $('#shade').addEventListener('mousedown', () => restore());
  $('#apply-shade').addEventListener('click', () => shadeCells($('#shade').value));
  $('#t-noshade').addEventListener('click', () => shadeCells(null));

  $('#ins-table').addEventListener('click', async () => {
    restore();
    const saveRange = saved;
    const v = await form('Insert table', [{ id: 'rows', label: 'Rows', type: 'number', value: 3, min: 1, max: 100 }, { id: 'cols', label: 'Columns', type: 'number', value: 3, min: 1, max: 20 }], 'Insert');
    if (!v) return;
    saved = saveRange;
    restore();
    const rows = Math.max(1, Math.min(100, v.rows | 0)), cols = Math.max(1, Math.min(20, v.cols | 0));
    const cw = Math.floor((page.w - page.l - page.r) / 15 / cols);
    let h = '<table style="width:' + cw * cols + 'px"><colgroup>' + `<col style="width:${cw}px">`.repeat(cols) + '</colgroup><tbody>';
    for (let r = 0; r < rows; r++) h += '<tr>' + `<td style="width:${cw}px"><br></td>`.repeat(cols) + '</tr>';
    h += '</tbody></table><p><br></p>';
    document.execCommand('insertHTML', false, h);
    markDirty();
    syncToolbar();
  });

  /* ---------------- images ---------------- */
  function loadImage(src) { return new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = src; }); }
  async function insertImageFile(file) {
    let url = await Lumen.readFileAsDataUrl(file);
    const img = await loadImage(url);
    if (!/^data:image\/(png|jpeg|gif|bmp)/.test(url)) {
      const c = document.createElement('canvas');
      c.width = img.naturalWidth || 300; c.height = img.naturalHeight || 200;
      c.getContext('2d').drawImage(img, 0, 0);
      url = c.toDataURL('image/png');
    }
    const maxW = (page.w - page.l - page.r) / 15;
    let w = img.naturalWidth, h = img.naturalHeight;
    if (w > maxW) { h = h * maxW / w; w = maxW; }
    restore();
    document.execCommand('insertHTML', false, `<img src="${url}" style="width:${Math.round(w)}px;height:${Math.round(h)}px">`);
    markDirty();
  }
  $('#ins-image').addEventListener('click', () => {
    restore();
    const saveRange = saved;
    const i = document.createElement('input');
    i.type = 'file';
    i.accept = 'image/*';
    i.onchange = async () => { if (i.files[0]) { saved = saveRange; try { await insertImageFile(i.files[0]); } catch (e) { toast('Could not read that image.'); } } };
    i.click();
  });

  /* image selection + resize */
  let selImg = null;
  const frame = $('#imgframe');
  function positionFrame() {
    if (!selImg || !doc.contains(selImg)) { frame.hidden = true; return; }
    const r = selImg.getBoundingClientRect(), s = stage.getBoundingClientRect();
    frame.hidden = false;
    frame.style.left = (r.left - s.left + stage.scrollLeft) + 'px';
    frame.style.top = (r.top - s.top + stage.scrollTop) + 'px';
    frame.style.width = r.width + 'px';
    frame.style.height = r.height + 'px';
  }
  doc.addEventListener('click', e => {
    if (selImg) selImg.classList.remove('sel');
    if (e.target.tagName === 'IMG') { selImg = e.target; selImg.classList.add('sel'); positionFrame(); } else { selImg = null; frame.hidden = true; }
    if (e.ctrlKey) {
      const a = e.target.closest && e.target.closest('a');
      const href = a && a.getAttribute('href');
      if (href && href[0] === '#') {
        const t = [...doc.querySelectorAll('span.bm')].find(b => b.dataset.n === href.slice(1));
        if (t) (t.parentElement || t).scrollIntoView({ block: 'center', behavior: 'smooth' });
      } else if (href) window.open(a.href, '_blank', 'noopener');
    }
  });
  stage.addEventListener('scroll', positionFrame);
  $('#imgh').addEventListener('mousedown', e => {
    e.preventDefault();
    if (!selImg) return;
    const startX = e.clientX, w0 = selImg.getBoundingClientRect().width / state.zoom, ratio = selImg.getBoundingClientRect().height / selImg.getBoundingClientRect().width;
    const mv = (ev) => {
      const w = Math.max(16, w0 + (ev.clientX - startX) / state.zoom);
      selImg.style.width = Math.round(w) + 'px';
      selImg.style.height = Math.round(w * ratio) + 'px';
      positionFrame();
    };
    const up = () => { document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); markDirty(); };
    document.addEventListener('mousemove', mv);
    document.addEventListener('mouseup', up);
  });

  /* ---------------- misc inserts ---------------- */
  $('#ins-link').addEventListener('click', async () => { await insertLink(); });
  async function insertLink() {
    restore();
    const saveRange = saved;
    const sel = getSelection().toString();
    const v = await form('Insert link', [{ id: 'url', label: 'Address', value: 'https://' }, { id: 'text', label: 'Text to display', value: sel }], 'Insert');
    if (!v || !v.url) return;
    saved = saveRange;
    restore();
    if (sel) document.execCommand('createLink', false, v.url);
    else document.execCommand('insertHTML', false, `<a href="${esc(v.url)}">${esc(v.text || v.url)}</a>`);
    markDirty();
  }
  $('#ins-unlink').addEventListener('click', () => exec('unlink'));

  /* ---------------- footnotes ---------------- */
  const fnlist = $('#fnlist');
  function renumberFootnotes() {
    const ns = [...doc.querySelectorAll('sup.fn')];
    ns.forEach((s, i) => { s.textContent = String(i + 1); });
    fnlist.innerHTML = ns.map((s, i) => `<div><sup>${i + 1}</sup> ${esc(s.dataset.t || '')}</div>`).join('');
    fnlist.style.display = ns.length ? 'block' : 'none';
  }
  async function insertFootnote() {
    restore();
    const saveRange = saved;
    const v = await form('Insert footnote', [{ id: 't', label: 'Footnote text', value: '' }], 'Insert');
    if (!v) return;
    saved = saveRange;
    restore();
    document.execCommand('insertHTML', false, `<sup class="fn" data-t="${esc(v.t)}" contenteditable="false">1</sup>`);
    renumberFootnotes(); markDirty();
  }
  async function editFootnote(s) {
    const v = await form('Edit footnote', [{ id: 't', label: 'Footnote text', value: s.dataset.t || '' }, { id: 'del', label: 'Delete this footnote', type: 'checkbox', value: false }], 'OK');
    if (!v) return;
    if (v.del) s.remove(); else s.dataset.t = v.t;
    renumberFootnotes(); markDirty();
  }
  doc.addEventListener('dblclick', e => { const s = e.target.closest && e.target.closest('sup.fn'); if (s) editFootnote(s); });
  $('#ins-fn').addEventListener('click', insertFootnote);

  /* ---------------- comments ---------------- */
  async function insertComment() {
    restore();
    const sel = getSelection();
    if (!sel.rangeCount || sel.isCollapsed || !doc.contains(sel.anchorNode)) { toast('Select the text you want to comment on first.'); return; }
    const saveRange = saved;
    const frag = sel.getRangeAt(0).cloneContents();
    const holder = document.createElement('div');
    holder.appendChild(frag);
    if (holder.querySelector('p,div,h1,h2,h3,li,table')) { toast('Select text within a single paragraph to comment on it.'); return; }
    const v = await form('New comment', [{ id: 't', label: 'Comment', value: '' }, { id: 'a', label: 'Author', value: localStorage.getItem('lumen-author') || 'Author' }], 'Add');
    if (!v || !v.t.trim()) return;
    try { localStorage.setItem('lumen-author', v.a); } catch (e) { /* ignore */ }
    saved = saveRange;
    restore();
    document.execCommand('insertHTML', false, `<span class="cmt" data-t="${esc(v.t)}" data-a="${esc(v.a)}" title="${esc(v.a + ': ' + v.t)}">${holder.innerHTML}</span>`);
    markDirty();
  }
  async function editComment(s) {
    const v = await form('Comment', [{ id: 't', label: 'Comment', value: s.dataset.t || '' }, { id: 'del', label: 'Delete this comment', type: 'checkbox', value: false }], 'OK');
    if (!v) return;
    if (v.del) s.replaceWith(...s.childNodes);
    else { s.dataset.t = v.t; s.title = (s.dataset.a || '') + ': ' + v.t; }
    markDirty();
  }
  doc.addEventListener('dblclick', e => { const s = e.target.closest && e.target.closest('span.cmt'); if (s && !e.target.closest('sup.fn')) editComment(s); });
  $('#ins-cmt').addEventListener('click', insertComment);

  /* ---------------- table of contents ---------------- */
  function buildToc() {
    const hs = [...doc.querySelectorAll('h1,h2,h3')];
    if (!hs.length) { toast('Add some headings (Heading 1-3) first.'); return null; }
    const ph = page.h / 15, top = sheet.getBoundingClientRect().top;
    return hs.map(h => {
      const lvl = +h.tagName[1];
      const pg = Math.max(1, Math.floor((h.getBoundingClientRect().top - top) / state.zoom / ph) + 1);
      return `<p class="toc toc${lvl}">${esc(h.textContent.trim())}<span class="tocdots" contenteditable="false"></span><span class="tocpg">${pg}</span></p>`;
    }).join('');
  }
  function updateToc() {
    const old = [...doc.querySelectorAll('p.toc')];
    if (!old.length) return false;
    const html = buildToc();
    if (html === null) return true;
    const tpl = document.createElement('div');
    tpl.innerHTML = html;
    old[0].before(...tpl.childNodes);
    old.forEach(p => p.remove());
    markDirty();
    return true;
  }
  $('#ins-toc').addEventListener('click', () => {
    restore();
    const html = buildToc();
    if (html === null) return;
    document.execCommand('insertHTML', false, html + '<p><br></p>');
    updateToc();
    markDirty();
  });
  $('#upd-toc').addEventListener('click', () => { if (!updateToc()) toast('There is no table of contents in this document yet.'); });
  $('#lay-cols').addEventListener('change', e => { page.cols = parseInt(e.target.value, 10) || 1; applyPage(); markDirty(); });
  $('#ins-header').addEventListener('click', () => { hdr.focus(); stage.scrollTop = 0; });
  $('#ins-footer').addEventListener('click', () => { ftr.focus(); stage.scrollTop = stage.scrollHeight; });
  const insField = (kind) => {
    const t = (document.activeElement === hdr || document.activeElement === ftr) ? document.activeElement : lastHF;
    t.focus();
    document.execCommand('insertHTML', false, `<span class="fld" data-f="${kind}" contenteditable="false">1</span>&nbsp;`);
    markDirty();
  };
  $('#ins-pgnum').addEventListener('click', () => insField('PAGE'));
  $('#ins-pgtotal').addEventListener('click', () => insField('NUMPAGES'));
  [hdr, ftr].forEach(h => h.addEventListener('input', markDirty));
  $('#ins-hr').addEventListener('click', () => { restore(); document.execCommand('insertHTML', false, '<hr><p><br></p>'); markDirty(); });
  function pageBreak() { restore(); document.execCommand('insertHTML', false, '<div class="pb" contenteditable="false"></div><p><br></p>'); markDirty(); }
  $('#ins-pb').addEventListener('click', pageBreak);
  $('#ins-date').addEventListener('click', () => { restore(); document.execCommand('insertText', false, new Date().toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' })); markDirty(); });
  $('#ins-symbol').addEventListener('click', async () => {
    restore();
    const saveRange = saved;
    const v = await form('Insert symbol', [{ id: 'sym', label: 'Symbol', type: 'select', value: '©', options: ['©', '®', '™', '°', '±', '×', '÷', '≤', '≥', '≠', '≈', '∞', 'π', 'Ω', 'µ', '→', '←', '↑', '↓', '✓', '★', '•', '€', '£', '¥', '§', '¶'] }], 'Insert');
    if (!v) return;
    saved = saveRange;
    restore();
    document.execCommand('insertText', false, v.sym);
    markDirty();
  });
  $('#btn-selectall').addEventListener('click', () => { restore(); document.execCommand('selectAll'); });

  /* ---------------- generic node insert ---------------- */
  function insertNodeAtCaret(node, caretAfter) {
    restore();
    const s = getSelection();
    if (!s.rangeCount) return;
    const r = s.getRangeAt(0);
    r.deleteContents();
    r.insertNode(node);
    if (caretAfter !== false) {
      const c = document.createRange();
      c.setStartAfter(node); c.collapse(true);
      s.removeAllRanges(); s.addRange(c);
    }
    markDirty();
  }
  const author = () => { try { return localStorage.getItem('lumen-author') || 'Author'; } catch (e) { return 'Author'; } };

  /* ---------------- endnotes ---------------- */
  const enlist = $('#enlist');
  function roman(n) {
    const t = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
    let o = '';
    for (const [v, s] of t) while (n >= v) { o += s; n -= v; }
    return o;
  }
  function renumberEndnotes() {
    const ns = [...doc.querySelectorAll('sup.en')];
    ns.forEach((s, i) => { const r = roman(i + 1); if (s.textContent !== r) s.textContent = r; });
    const html = ns.map((s, i) => `<div><sup>${roman(i + 1)}</sup> ${esc(s.dataset.t || '')}</div>`).join('');
    if (enlist.innerHTML !== html) enlist.innerHTML = html;
    enlist.style.display = ns.length ? 'block' : 'none';
  }
  $('#ins-en').addEventListener('click', async () => {
    restore();
    const saveRange = saved;
    const v = await form('Insert endnote', [{ id: 't', label: 'Endnote text', value: '' }], 'Insert');
    if (!v) return;
    saved = saveRange;
    const s = document.createElement('sup');
    s.className = 'en'; s.dataset.t = v.t; s.setAttribute('contenteditable', 'false'); s.textContent = 'i';
    insertNodeAtCaret(s);
    renumberEndnotes();
  });
  doc.addEventListener('dblclick', async e => {
    const s = e.target.closest && e.target.closest('sup.en');
    if (!s) return;
    const v = await form('Edit endnote', [{ id: 't', label: 'Endnote text', value: s.dataset.t || '' }, { id: 'del', label: 'Delete this endnote', type: 'checkbox', value: false }], 'OK');
    if (!v) return;
    if (v.del) s.remove(); else s.dataset.t = v.t;
    renumberEndnotes(); markDirty();
  });

  /* ---------------- text box / equation / caption ---------------- */
  $('#ins-tbox').addEventListener('click', () => {
    restore();
    document.execCommand('insertHTML', false, '<div class="tbox"><p>Text box</p></div><p><br></p>');
    markDirty();
  });
  const EQ_HELP = 'Type the equation in linear form: a/b for a fraction, x^2, x_1, sqrt(x), brackets to group. Example: x = (-b + sqrt(b^2 - 4ac))/(2a)';
  $('#ins-eq').addEventListener('click', async () => {
    restore();
    const saveRange = saved;
    const v = await form('Insert equation', [{ type: 'note', label: EQ_HELP }, { id: 'src', label: 'Equation', value: 'x = (-b + sqrt(b^2 - 4ac))/(2a)' }], 'Insert');
    if (!v || !v.src.trim()) return;
    saved = saveRange;
    insertNodeAtCaret(DocxIO.Eq.span(v.src.trim()));
  });
  doc.addEventListener('dblclick', async e => {
    const s = e.target.closest && e.target.closest('span.eq');
    if (!s) return;
    const v = await form('Edit equation', [{ type: 'note', label: EQ_HELP }, { id: 'src', label: 'Equation', value: s.dataset.src || '' }, { id: 'del', label: 'Delete this equation', type: 'checkbox', value: false }], 'OK');
    if (!v) return;
    if (v.del || !v.src.trim()) s.remove(); else s.replaceWith(DocxIO.Eq.span(v.src.trim()));
    markDirty();
  });
  $('#ins-caption').addEventListener('click', async () => {
    restore();
    const saveRange = saved;
    const v = await form('Insert caption', [{ id: 'l', label: 'Label', type: 'select', value: 'Figure', options: ['Figure', 'Table', 'Equation'] }, { id: 't', label: 'Caption text', value: '' }], 'Insert');
    if (!v) return;
    saved = saveRange;
    restore();
    const n = [...doc.querySelectorAll('p')].filter(p => new RegExp('^' + v.l + ' \\d+').test(p.textContent)).length + 1;
    document.execCommand('insertHTML', false, `<p><i>${v.l} ${n}${v.t ? ': ' + esc(v.t) : ''}</i></p>`);
    markDirty();
  });

  /* ---------------- bookmarks & cross-references ---------------- */
  function makeBookmark(name) {
    const s = document.createElement('span');
    s.className = 'bm'; s.dataset.n = name; s.setAttribute('contenteditable', 'false'); s.title = 'Bookmark: ' + name;
    return s;
  }
  $('#ins-bm').addEventListener('click', async () => {
    restore();
    const saveRange = saved;
    const existing = [...doc.querySelectorAll('span.bm')].map(b => b.dataset.n).filter(n => !/^_/.test(n));
    const fields = [{ id: 'n', label: 'Bookmark name', value: '' }];
    if (existing.length) fields.push({ id: 'rm', label: 'Or delete bookmark', type: 'select', value: '', options: [''].concat(existing) });
    const v = await form('Bookmark', fields, 'OK');
    if (!v) return;
    if (v.rm) { doc.querySelectorAll('span.bm').forEach(b => { if (b.dataset.n === v.rm) b.remove(); }); markDirty(); return; }
    let name = (v.n || '').trim().replace(/[^\p{L}\p{N}_]/gu, '_');
    if (!name) return;
    if (!/^\p{L}/u.test(name)) name = 'b' + name;
    if (existing.includes(name)) { toast('A bookmark with that name already exists.'); return; }
    saved = saveRange;
    insertNodeAtCaret(makeBookmark(name.slice(0, 40)));
  });
  $('#ins-xref').addEventListener('click', async () => {
    restore();
    const saveRange = saved;
    const targets = [];
    doc.querySelectorAll('span.bm').forEach(b => { if (!/^_/.test(b.dataset.n)) targets.push({ label: 'Bookmark: ' + b.dataset.n, name: b.dataset.n, text: b.dataset.n }); });
    doc.querySelectorAll('h1,h2,h3,h4,h5,h6').forEach(h => { const t = h.textContent.trim(); if (t) targets.push({ label: 'Heading: ' + t.slice(0, 50), h, text: t }); });
    if (!targets.length) { toast('Add a bookmark or a heading first.'); return; }
    const v = await form('Cross-reference', [{ id: 'to', label: 'Refer to', type: 'select', value: targets[0].label, options: targets.map(t => t.label) }, { id: 'text', label: 'Text (blank = automatic)', value: '' }], 'Insert');
    if (!v) return;
    const t = targets.find(x => x.label === v.to);
    if (!t) return;
    let name = t.name;
    if (t.h) {
      let b = t.h.querySelector('span.bm');
      if (!b) {
        let i = 1;
        const used = new Set([...doc.querySelectorAll('span.bm')].map(x => x.dataset.n));
        while (used.has('_Ref' + i)) i++;
        b = makeBookmark('_Ref' + i);
        t.h.prepend(b);
      }
      name = b.dataset.n;
    }
    saved = saveRange;
    const a = document.createElement('a');
    a.setAttribute('href', '#' + name);
    a.textContent = v.text || t.text;
    insertNodeAtCaret(a);
  });

  /* ---------------- track changes ---------------- */
  const myChange = (el) => el && el.dataset.a === author();
  function newChange(tag) {
    const el = document.createElement(tag);
    el.className = 'tc';
    el.dataset.a = author();
    el.dataset.d = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    el.title = (tag === 'ins' ? 'Inserted by ' : 'Deleted by ') + el.dataset.a;
    return el;
  }
  function setCaret(node, offset) {
    const r = document.createRange();
    r.setStart(node, offset); r.collapse(true);
    const s = getSelection();
    s.removeAllRanges(); s.addRange(r);
  }
  function textNodesIn(range) {
    const out = [];
    const root = range.commonAncestorContainer.nodeType === 3 ? range.commonAncestorContainer.parentNode : range.commonAncestorContainer;
    const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    while (w.nextNode()) {
      const n = w.currentNode;
      if (!range.intersectsNode(n) || !n.length) continue;
      const pe = n.parentElement;
      if (pe.closest('[contenteditable="false"]')) continue;
      const s = n === range.startContainer ? range.startOffset : 0;
      const e = n === range.endContainer ? range.endOffset : n.length;
      if (e > s) out.push({ n, s, e });
    }
    return out;
  }
  function markDeleted(range, backward) {
    const parts = textNodesIn(range);
    if (!parts.length) return false;
    let first = null, last = null;
    for (const p of parts) {
      let t = p.n;
      if (p.e < t.length) t.splitText(p.e);
      if (p.s > 0) t = t.splitText(p.s);
      const pe = t.parentElement;
      if (pe.closest('del.tc')) { first = first || t; last = t; continue; }
      const own = pe.closest('ins.tc');
      if (own && myChange(own)) {
        const mark = document.createTextNode('');
        t.replaceWith(mark);
        first = first || mark; last = mark;
        continue;
      }
      const d = newChange('del');
      t.replaceWith(d);
      d.appendChild(t);
      first = first || d; last = d;
    }
    const r = document.createRange();
    if (backward) r.setStartBefore(first); else r.setStartAfter(last);
    r.collapse(true);
    const s = getSelection();
    s.removeAllRanges(); s.addRange(r);
    doc.querySelectorAll('ins.tc').forEach(i => { if (!i.textContent && !i.querySelector('img,br')) i.remove(); });
    doc.normalize();
    markDirty();
    return true;
  }
  function insertTracked(text) {
    const s = getSelection();
    if (!s.rangeCount) return;
    let r = s.getRangeAt(0);
    if (!r.collapsed) { markDeleted(r, false); r = getSelection().getRangeAt(0); }
    let c = r.startContainer, o = r.startOffset;
    const pe = c.nodeType === 3 ? c.parentElement : c;
    const inIns = pe.closest('ins.tc');
    if (c.nodeType === 3 && inIns && myChange(inIns)) {
      c.insertData(o, text);
      setCaret(c, o + text.length);
      markDirty();
      return;
    }
    const inDel = pe.closest('del.tc');
    if (inDel) { r = document.createRange(); r.setStartAfter(inDel); r.collapse(true); c = r.startContainer; o = r.startOffset; }
    // extend an adjacent insertion of mine when the caret sits right after it
    let prev = null;
    if (c.nodeType === 3 && o === 0) prev = c.previousSibling;
    else if (c.nodeType === 1) prev = c.childNodes[o - 1] || null;
    if (prev && prev.nodeType === 1 && prev.matches('ins.tc') && myChange(prev) && prev.lastChild && prev.lastChild.nodeType === 3) {
      prev.lastChild.appendData(text);
      setCaret(prev.lastChild, prev.lastChild.length);
      markDirty();
      return;
    }
    const ins = newChange('ins');
    const tn = document.createTextNode(text);
    ins.appendChild(tn);
    r.insertNode(ins);
    setCaret(tn, tn.length);
    markDirty();
  }
  function targetRange(e) {
    const tr = e.getTargetRanges ? e.getTargetRanges()[0] : null;
    const r = document.createRange();
    if (tr) { r.setStart(tr.startContainer, tr.startOffset); r.setEnd(tr.endContainer, tr.endOffset); return r; }
    const s = getSelection();
    return s.rangeCount ? s.getRangeAt(0).cloneRange() : r;
  }
  doc.addEventListener('beforeinput', e => {
    if (!state.track || e.isComposing) return;
    const t = e.inputType;
    if (t === 'insertText' || t === 'insertReplacementText') {
      const text = e.data || (e.dataTransfer && e.dataTransfer.getData('text/plain')) || '';
      if (!text) return;
      e.preventDefault();
      const r = targetRange(e);
      if (!r.collapsed) markDeleted(r, false);
      insertTracked(text);
    } else if (/^delete/.test(t)) {
      const r = targetRange(e);
      if (r.collapsed) return;
      if (textNodesIn(r).length) { e.preventDefault(); markDeleted(r, /Backward|Cut/.test(t)); }
    }
  });
  function setTrack(on) {
    state.track = !!on;
    $('#rv-track').classList.toggle('on', state.track);
    $('#st-track').textContent = state.track ? 'Track changes: on' : '';
  }
  $('#rv-track').addEventListener('click', () => { setTrack(!state.track); toast(state.track ? 'Track changes is on. Typing and deleting are now recorded.' : 'Track changes is off.'); });
  const changes = () => [...doc.querySelectorAll('ins.tc, del.tc')];
  function resolveChange(el, accept) {
    const keep = (el.tagName === 'INS') === accept;
    if (keep) el.replaceWith(...el.childNodes); else el.remove();
  }
  function currentChanges() {
    const s = getSelection();
    if (!s.rangeCount || !doc.contains(s.anchorNode)) return [];
    const r = s.getRangeAt(0);
    if (!r.collapsed) { const hit = changes().filter(c => r.intersectsNode(c)); if (hit.length) return hit; }
    const n = s.anchorNode.nodeType === 3 ? s.anchorNode.parentElement : s.anchorNode;
    const c = n.closest('ins.tc, del.tc');
    return c ? [c] : [];
  }
  function gotoChange(dir) {
    const all = changes();
    if (!all.length) { toast('There are no tracked changes in this document.'); return null; }
    const s = getSelection();
    let idx = dir > 0 ? 0 : all.length - 1;
    if (s.rangeCount && doc.contains(s.anchorNode)) {
      const r = s.getRangeAt(0);
      const pos = all.map(c => { const cr = document.createRange(); cr.selectNode(c); return dir > 0 ? cr.compareBoundaryPoints(Range.START_TO_START, r) : cr.compareBoundaryPoints(Range.END_TO_END, r); });
      const found = dir > 0 ? pos.findIndex(p => p > 0) : pos.map(p => p < 0).lastIndexOf(true);
      if (found >= 0) idx = found;
    }
    const c = all[idx];
    const r = document.createRange();
    r.selectNodeContents(c);
    doc.focus();
    s.removeAllRanges(); s.addRange(r);
    c.scrollIntoView({ block: 'center' });
    return c;
  }
  function resolveCurrent(accept) {
    restore();
    let cur = currentChanges();
    if (!cur.length) { const c = gotoChange(1); if (!c) return; toast('Moved to the next change. Press again to ' + (accept ? 'accept' : 'reject') + ' it.'); return; }
    cur.forEach(c => resolveChange(c, accept));
    doc.normalize();
    markDirty();
  }
  function resolveAll(accept) {
    const all = changes();
    if (!all.length) { toast('There are no tracked changes in this document.'); return; }
    all.forEach(c => resolveChange(c, accept));
    doc.normalize();
    markDirty();
    toast(all.length + (accept ? ' changes accepted.' : ' changes rejected.'));
  }
  $('#rv-next').addEventListener('click', () => gotoChange(1));
  $('#rv-prev').addEventListener('click', () => gotoChange(-1));
  $('#rv-accept').addEventListener('click', () => resolveCurrent(true));
  $('#rv-reject').addEventListener('click', () => resolveCurrent(false));
  $('#rv-accept-all').addEventListener('click', () => resolveAll(true));
  $('#rv-reject-all').addEventListener('click', () => resolveAll(false));
  $('#rv-cmt').addEventListener('click', insertComment);
  $('#rv-count').addEventListener('click', () => {
    const text = doc.innerText.replace(/​/g, '');
    const words = (text.trim().match(/\S+/g) || []).length;
    const paras = [...doc.querySelectorAll('p,h1,h2,h3,h4,h5,h6,li')].filter(p => p.textContent.trim()).length;
    form('Word count', [{ type: 'note', label: `Words: ${words}` }, { type: 'note', label: `Characters (no spaces): ${text.replace(/\s/g, '').length}` }, { type: 'note', label: `Characters (with spaces): ${text.replace(/\n/g, '').length}` }, { type: 'note', label: `Paragraphs: ${paras}` }], 'Close');
  });

  /* ---------------- mail merge ---------------- */
  const mm = { fields: [], rows: [], preview: false, csv: 'Name,Company,City\nAda Lovelace,Analytical Engines,London\nAlan Turing,Bletchley Park,Manchester' };
  function parseCsv(text) {
    const rows = [];
    let row = [], cur = '', q = false;
    const delim = (text.split('\n')[0].match(/\t/) ? '\t' : (text.split('\n')[0].match(/;/) && !text.split('\n')[0].match(/,/) ? ';' : ','));
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) { if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += ch; }
      else if (ch === '"') q = true;
      else if (ch === delim) { row.push(cur); cur = ''; }
      else if (ch === '\n') { row.push(cur); cur = ''; rows.push(row); row = []; }
      else if (ch !== '\r') cur += ch;
    }
    if (cur || row.length) { row.push(cur); rows.push(row); }
    return rows.filter(r => r.some(c => c.trim()));
  }
  function loadRecipients(text) {
    const rows = parseCsv(text);
    if (rows.length < 2) return false;
    mm.csv = text;
    mm.fields = rows[0].map(f => f.trim().replace(/[^\p{L}\p{N}_]/gu, '_')).filter(Boolean);
    mm.rows = rows.slice(1).map(r => { const o = {}; mm.fields.forEach((f, i) => { o[f] = (r[i] || '').trim(); }); return o; });
    return true;
  }
  async function editRecipients() {
    const v = await form('Mail merge recipients', [{ type: 'note', label: 'One recipient per line. The first line holds the field names. Paste from a spreadsheet or a CSV file.' }, { id: 'csv', label: 'Recipients', type: 'textarea', rows: 9, value: mm.csv }], 'Use this list');
    if (!v) return false;
    if (!loadRecipients(v.csv)) { toast('Enter a header line and at least one recipient.'); return false; }
    toast(mm.rows.length + ' recipients, fields: ' + mm.fields.join(', '));
    return true;
  }
  function showPreview(on) {
    mm.preview = !!on && mm.rows.length > 0;
    $('#mm-preview').classList.toggle('on', mm.preview);
    doc.querySelectorAll('span.mf').forEach(s => { s.textContent = mm.preview ? (mm.rows[0][s.dataset.n] || '') : '«' + s.dataset.n + '»'; });
  }
  $('#mm-list').addEventListener('click', editRecipients);
  $('#mm-field').addEventListener('click', async () => {
    restore();
    const saveRange = saved;
    if (!mm.fields.length && !(await editRecipients())) return;
    const v = await form('Insert merge field', [{ id: 'f', label: 'Field', type: 'select', value: mm.fields[0], options: mm.fields }], 'Insert');
    if (!v) return;
    saved = saveRange;
    const s = document.createElement('span');
    s.className = 'mf'; s.dataset.n = v.f; s.setAttribute('contenteditable', 'false');
    s.textContent = mm.preview ? (mm.rows[0][v.f] || '') : '«' + v.f + '»';
    insertNodeAtCaret(s);
  });
  $('#mm-preview').addEventListener('click', async () => {
    if (!mm.rows.length && !(await editRecipients())) return;
    showPreview(!mm.preview);
  });
  $('#mm-finish').addEventListener('click', async () => {
    if (!doc.querySelector('span.mf')) { toast('Insert at least one merge field first.'); return; }
    if (!mm.rows.length && !(await editRecipients())) return;
    if (!(await Lumen.confirmBox('Finish & merge', `This replaces the document with ${mm.rows.length} merged copies, one per recipient, separated by page breaks. Save the template first if you want to keep it.`, 'Merge'))) return;
    showPreview(false);
    const tpl = doc.innerHTML;
    const parts = mm.rows.map(row => {
      const d = document.createElement('div');
      d.innerHTML = tpl;
      d.querySelectorAll('span.mf').forEach(s => s.replaceWith(document.createTextNode(row[s.dataset.n] || '')));
      return d.innerHTML;
    });
    doc.innerHTML = parts.join('<div class="pb" contenteditable="false"></div>');
    state.handle = null;
    setTitle(titleEl.value + ' - merged');
    markDirty();
    toast('Merged ' + mm.rows.length + ' records.');
  });

  /* ---------------- find & replace ---------------- */
  const fb = $('#findbar');
  function openFind(focusReplace) {
    fb.hidden = false;
    const s = getSelection().toString();
    if (s && !/\n/.test(s)) $('#find-q').value = s;
    (focusReplace ? $('#find-r') : $('#find-q')).focus();
    $('#find-q').select();
  }
  function findNext(back) {
    const q = $('#find-q').value;
    if (!q) return false;
    doc.focus();
    const ok = window.find(q, $('#find-case').checked, !!back, true, false, false, false);
    const s = getSelection();
    const inside = ok && s.rangeCount && doc.contains(s.anchorNode);
    $('#find-msg').textContent = inside ? '' : 'No more matches';
    if (inside) { const n = s.anchorNode.nodeType === 3 ? s.anchorNode.parentNode : s.anchorNode; n.scrollIntoView({ block: 'center' }); }
    return inside;
  }
  $('#btn-find').addEventListener('click', () => openFind(false));
  $('#find-close').addEventListener('click', () => { fb.hidden = true; doc.focus(); });
  $('#find-next').addEventListener('click', () => findNext(false));
  $('#find-prev').addEventListener('click', () => findNext(true));
  $('#find-q').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); findNext(e.shiftKey); } if (e.key === 'Escape') $('#find-close').click(); });
  $('#find-rep').addEventListener('click', () => {
    const q = $('#find-q').value, s = getSelection();
    const cs = $('#find-case').checked;
    const t = s.toString();
    if (t && (cs ? t === q : t.toLowerCase() === q.toLowerCase()) && doc.contains(s.anchorNode)) { document.execCommand('insertText', false, $('#find-r').value); markDirty(); }
    findNext(false);
  });
  $('#find-all').addEventListener('click', () => {
    const q = $('#find-q').value;
    if (!q) return;
    doc.focus();
    const s = getSelection();
    s.collapse(doc, 0);
    let n = 0;
    while (n < 5000 && window.find(q, $('#find-case').checked, false, false, false, false, false)) {
      if (!doc.contains(getSelection().anchorNode)) break;
      document.execCommand('insertText', false, $('#find-r').value);
      n++;
    }
    $('#find-msg').textContent = n + ' replaced';
    if (n) markDirty();
  });

  /* ---------------- keyboard ---------------- */
  doc.addEventListener('keydown', e => {
    if (e.key === 'Tab') {
      e.preventDefault();
      const cell = closestCell();
      const li = (() => { let n = getSelection().anchorNode; while (n && n !== doc) { if (n.nodeName === 'LI') return n; n = n.parentNode; } return null; })();
      if (cell) {
        const cells = [...cell.closest('table').querySelectorAll('td,th')];
        let i = cells.indexOf(cell) + (e.shiftKey ? -1 : 1);
        if (i >= cells.length) { tableOp('row-below'); i = cells.length; }
        const target = [...cell.closest('table').querySelectorAll('td,th')][Math.max(0, i)];
        if (target) { const r = document.createRange(); r.selectNodeContents(target); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); }
      } else if (li) exec(e.shiftKey ? 'outdent' : 'indent');
      else document.execCommand('insertText', false, '\t');
      return;
    }
    if (e.key === 'Enter' && e.ctrlKey) { e.preventDefault(); pageBreak(); }
  });
  document.addEventListener('keydown', e => {
    const k = e.key.toLowerCase();
    if (e.ctrlKey || e.metaKey) {
      if (k === 's') { e.preventDefault(); save(e.shiftKey); }
      else if (k === 'o') { e.preventDefault(); openDialog(); }
      else if (k === 'n' && !e.shiftKey) { e.preventDefault(); newDoc(); }
      else if (k === 'p') { e.preventDefault(); window.print(); }
      else if (k === 'f') { e.preventDefault(); openFind(false); }
      else if (k === 'h') { e.preventDefault(); openFind(true); }
      else if (k === 'k') { e.preventDefault(); insertLink(); }
      else if (!e.shiftKey && (k === 'e' || k === 'l' || k === 'r' || k === 'j') && doc.contains(document.activeElement)) {
        e.preventDefault();
        exec({ e: 'justifyCenter', l: 'justifyLeft', r: 'justifyRight', j: 'justifyFull' }[k]);
      }
    }
  });
  doc.addEventListener('input', markDirty);

  /* ---------------- paste ---------------- */
  const KEEP_STYLE = ['font-weight', 'font-style', 'text-decoration', 'text-decoration-line', 'color', 'background-color', 'font-family', 'font-size', 'text-align', 'margin-left', 'text-indent', 'vertical-align', 'width', 'height'];
  function cleanHtml(html) {
    const d = new DOMParser().parseFromString(html, 'text/html');
    d.querySelectorAll('script,style,meta,link,title,head,xml,object,iframe').forEach(n => n.remove());
    d.querySelectorAll('*').forEach(el => {
      if (el.tagName.includes(':')) { el.replaceWith(...el.childNodes); return; }
      const keep = {};
      const st = el.style;
      for (const p of KEEP_STYLE) { const v = st.getPropertyValue(p); if (v && !/^(inherit|initial)$/.test(v)) keep[p] = v; }
      const attrs = { href: el.getAttribute('href'), src: el.getAttribute('src'), colspan: el.getAttribute('colspan'), rowspan: el.getAttribute('rowspan'), alt: el.getAttribute('alt'), start: el.getAttribute('start') };
      [...el.attributes].forEach(a => el.removeAttribute(a.name));
      for (const k in attrs) if (attrs[k] && !/^javascript:/i.test(attrs[k])) el.setAttribute(k, attrs[k]);
      const css = Object.entries(keep).map(([k, v]) => k + ':' + v).join(';');
      if (css) el.setAttribute('style', css);
    });
    const walker = document.createTreeWalker(d.body, NodeFilter.SHOW_TEXT);
    const texts = [];
    while (walker.nextNode()) texts.push(walker.currentNode);
    texts.forEach(t => { if (!t.parentElement.closest('pre')) t.nodeValue = t.nodeValue.replace(/[\r\n\t ]+/g, ' '); });
    return d.body.innerHTML.replace(/<!--[\s\S]*?-->/g, '');
  }
  doc.addEventListener('paste', async e => {
    const cd = e.clipboardData;
    if (!cd) return;
    const files = [...cd.files].filter(f => f.type.startsWith('image/'));
    if (files.length) { e.preventDefault(); for (const f of files) await insertImageFile(f); return; }
    const html = cd.getData('text/html');
    if (html) { e.preventDefault(); document.execCommand('insertHTML', false, cleanHtml(html)); markDirty(); return; }
    const text = cd.getData('text/plain');
    if (text) {
      e.preventDefault();
      const lines = text.replace(/\r/g, '').split('\n');
      if (lines.length === 1) document.execCommand('insertText', false, text);
      else document.execCommand('insertHTML', false, lines.map(l => '<p>' + (esc(l) || '<br>') + '</p>').join(''));
      markDirty();
    }
  });

  /* ---------------- file operations ---------------- */
  const titleEl = $('#title');
  function setTitle(n) { titleEl.value = n; document.title = n + ' - Lumen Write'; }
  titleEl.addEventListener('input', () => { document.title = titleEl.value + ' - Lumen Write'; markDirty(); });

  async function newDoc() {
    if (state.dirty && !(await Lumen.confirmBox('Discard changes?', 'This document has unsaved changes. Create a new one anyway?', 'Discard'))) return;
    doc.innerHTML = '<p><br></p>';
    hdr.innerHTML = ''; ftr.innerHTML = '';
    setTrack(false);
    page = { w: 12240, h: 15840, t: 1440, r: 1440, b: 1440, l: 1440 };
    state.handle = null;
    setTitle('Document1');
    applyPage();
    markClean();
    doc.focus();
  }
  async function openDialog() {
    if (state.dirty && !(await Lumen.confirmBox('Discard changes?', 'This document has unsaved changes. Open another file anyway?', 'Discard'))) return;
    const r = await FileIO.open([{ description: 'Documents', accept: { [DocxIO.DOCX_MIME]: ['.docx'], 'text/plain': ['.txt'], 'text/html': ['.html', '.htm'] } }], '.docx,.txt,.html,.htm');
    if (r) await loadFile(r.file, r.handle);
  }
  async function loadFile(file, handle) {
    try {
      const ext = (file.name.match(/\.([^.]+)$/) || [, ''])[1].toLowerCase();
      if (ext === 'docx' || ext === 'dotx' || ext === 'docm') {
        const r = await DocxIO.read(await file.arrayBuffer());
        doc.replaceChildren(...r.root.childNodes);
        hdr.innerHTML = r.header || ''; ftr.innerHTML = r.footer || '';
        page = r.page;
        setTrack(r.track);
        state.handle = handle;
        setTitle(baseName(file.name));
      } else if (ext === 'txt') {
        const t = await file.text();
        doc.innerHTML = t.replace(/\r/g, '').split('\n').map(l => '<p>' + (esc(l) || '<br>') + '</p>').join('');
        state.handle = null;
        setTitle(baseName(file.name));
      } else if (ext === 'html' || ext === 'htm') {
        doc.innerHTML = cleanHtml(await file.text()) || '<p><br></p>';
        state.handle = null;
        setTitle(baseName(file.name));
      } else if (ext === 'doc') {
        toast('Old binary .doc files are not supported. Open it once in any word processor and save as .docx.', 6000);
        return;
      } else { toast('Unsupported file type: .' + ext); return; }
      if (!doc.firstChild) doc.innerHTML = '<p><br></p>';
      applyPage();
      markClean();
      doc.scrollTop = 0;
      stage.scrollTop = 0;
      toast('Opened ' + file.name);
    } catch (err) {
      console.error(err);
      toast('Could not open file: ' + err.message, 6000);
    }
  }
  async function save(saveAs) {
    try {
      const blob = await DocxIO.write(doc, page, { title: titleEl.value, header: hdr.innerHTML, footer: ftr.innerHTML, track: state.track });
      const r = await FileIO.save(blob, (titleEl.value || 'Document') + '.docx', FILE_TYPES, state.handle, saveAs);
      if (r) {
        state.handle = r.handle;
        if (r.handle) setTitle(baseName(r.name));
        markClean();
        toast('Saved ' + r.name);
      }
    } catch (err) {
      console.error(err);
      toast('Save failed: ' + err.message, 6000);
    }
  }
  $('#style-mod').addEventListener('click', modifyStyleDialog);

  /* ---------------- citations & bibliography (APA style) ---------------- */
  const saveSel = () => { restore(); return saved; };
  const restoreSel = (r) => { saved = r; };
  const SRC_TYPES = [['book', 'Book'], ['article', 'Journal article'], ['web', 'Web site'], ['report', 'Report']];
  const authorsOf = (s) => String(s.author || '').split(';').map(a => a.trim()).filter(Boolean);
  const lastOf = (a) => a.split(',')[0].trim();
  const initials = (a) => { const f = (a.split(',')[1] || '').trim(); return f ? f.split(/[\s-]+/).filter(Boolean).map(x => x[0].toUpperCase() + '.').join(' ') : ''; };
  function citeText(s) {
    if (!s) return '(Missing source)';
    const au = authorsOf(s).map(lastOf);
    const who = au.length === 0 ? (s.title || s.tag) : au.length === 1 ? au[0] : au.length === 2 ? au[0] + ' & ' + au[1] : au[0] + ' et al.';
    return '(' + who + ', ' + (s.year || 'n.d.') + ')';
  }
  function bibEntry(s) {
    const au = authorsOf(s).map(a => lastOf(a) + (initials(a) ? ', ' + initials(a) : ''));
    const who = au.length ? (au.length > 1 ? au.slice(0, -1).join(', ') + ', & ' + au[au.length - 1] : au[0]) : '';
    const yr = '(' + (s.year || 'n.d.') + ').';
    let out = (who ? esc(who) + ' ' : '') + yr + ' ';
    if (s.type === 'article') out += esc(s.title || '') + '. <i>' + esc(s.journal || '') + '</i>.';
    else if (s.type === 'web') out += '<i>' + esc(s.title || '') + '</i>.' + (s.url ? ' Retrieved from ' + esc(s.url) : '');
    else out += '<i>' + esc(s.title || '') + '</i>.' + (s.city ? ' ' + esc(s.city) + ':' : '') + (s.publisher ? ' ' + esc(s.publisher) + '.' : '');
    return out;
  }
  const sources = () => (page.sources = page.sources || []);
  function makeTag(s) {
    const base = ((authorsOf(s)[0] ? lastOf(authorsOf(s)[0]) : (s.title || 'Src')).replace(/[^A-Za-z]/g, '').slice(0, 3) || 'Src') + String(s.year || '').slice(-2);
    let t = base, i = 1; while (sources().some(x => x.tag === t)) t = base + String.fromCharCode(96 + i++);
    return t;
  }
  async function sourceForm(s, title) {
    s = s || { type: 'book' };
    const v = await form(title, [
      { id: 'type', label: 'Type of source', type: 'select', value: s.type || 'book', options: SRC_TYPES },
      { id: 'author', label: 'Author(s) — Last, First; separate several with ;', value: s.author || '' },
      { id: 'title', label: 'Title', value: s.title || '' },
      { id: 'year', label: 'Year', value: s.year || '' },
      { id: 'journal', label: 'Journal name (articles)', value: s.journal || '' },
      { id: 'publisher', label: 'Publisher', value: s.publisher || '' },
      { id: 'city', label: 'City', value: s.city || '' },
      { id: 'url', label: 'URL (web sites)', value: s.url || '' },
    ], 'OK');
    if (!v) return null;
    if (!String(v.title).trim() && !String(v.author).trim()) { toast('Enter at least an author or a title.'); return null; }
    const o = {}; for (const k of ['type', 'author', 'title', 'year', 'journal', 'publisher', 'city', 'url']) o[k] = String(v[k] || '').trim();
    return o;
  }
  function refreshCitations() {
    doc.querySelectorAll('span.cite').forEach(sp => { sp.textContent = citeText(sources().find(x => x.tag === sp.dataset.tag)); });
  }
  async function insertCitation() {
    const range = saveSel();
    const list = sources();
    const v = await form('Insert citation', [
      { id: 'src', label: 'Source', type: 'select', value: list.length ? list[list.length - 1].tag : '__new', options: [...list.map(s => [s.tag, citeText(s) + '  ' + (s.title || '')]), ['__new', 'Add new source…']] },
    ], 'Insert');
    if (!v) return;
    let tag = v.src;
    if (tag === '__new') {
      const s = await sourceForm(null, 'Add new source');
      if (!s) return;
      s.tag = makeTag(s); list.push(s); tag = s.tag;
    }
    restoreSel(range);
    const sp = document.createElement('span');
    sp.className = 'cite'; sp.dataset.tag = tag; sp.setAttribute('contenteditable', 'false');
    sp.textContent = citeText(list.find(x => x.tag === tag));
    insertNodeAtCaret(sp);
    markDirty();
  }
  async function manageSources() {
    const list = sources();
    const v = await form('Manage sources', [
      { type: 'note', label: list.length ? list.length + ' source(s) in this document.' : 'No sources yet.' },
      { id: 'pick', label: 'Source', type: 'select', value: list.length ? list[0].tag : '__new', options: [...list.map(s => [s.tag, citeText(s) + '  ' + (s.title || '')]), ['__new', 'Add new source…']] },
      { id: 'act', label: 'Action', type: 'select', value: 'edit', options: [['edit', 'Edit / add'], ['delete', 'Delete']] },
    ], 'Go');
    if (!v) return;
    if (v.pick === '__new') { const s = await sourceForm(null, 'Add new source'); if (s) { s.tag = makeTag(s); list.push(s); markDirty(); toast('Source added'); } return; }
    const i = list.findIndex(s => s.tag === v.pick);
    if (i < 0) return;
    if (v.act === 'delete') {
      if (doc.querySelector(`span.cite[data-tag="${CSS.escape(v.pick)}"]`)) { toast('This source is cited in the document. Delete those citations first.', 5000); return; }
      list.splice(i, 1); markDirty(); toast('Source deleted'); return;
    }
    const s = await sourceForm(list[i], 'Edit source');
    if (s) { s.tag = list[i].tag; list[i] = s; refreshCitations(); markDirty(); toast('Source updated'); }
  }
  function bibHeading() { return [...doc.querySelectorAll('h1,h2,p.title')].find(h => /^\s*(references|bibliography|works cited)\s*$/i.test(h.textContent)); }
  function updateBibliography(insertIfMissing) {
    const list = sources().slice().sort((a, b) => (authorsOf(a)[0] || a.title || '').localeCompare(authorsOf(b)[0] || b.title || ''));
    if (!list.length) { toast('Add a source first (Insert citation or Manage sources).'); return; }
    refreshCitations();
    let h = bibHeading();
    if (!h) {
      if (!insertIfMissing) { toast('No "References" heading found. Use Bibliography to insert one.'); return; }
      const range = saveSel();
      h = document.createElement('h1'); h.textContent = 'References';
      let anchor = range && range.startContainer;
      while (anchor && anchor.parentNode !== doc) anchor = anchor.parentNode;
      if (anchor) anchor.after(h); else doc.appendChild(h);
    }
    let n = h.nextElementSibling;
    while (n && !/^H[1-6]$/.test(n.tagName) && !(n.classList && n.classList.contains('pb'))) { const nx = n.nextElementSibling; n.remove(); n = nx; }
    let after = h;
    for (const s of list) { const p = document.createElement('p'); p.style.marginLeft = '48px'; p.style.textIndent = '-48px'; p.innerHTML = bibEntry(s); after.after(p); after = p; }
    markDirty(); toast('Bibliography updated');
  }
  $('#cite-ins').addEventListener('click', insertCitation);
  $('#cite-src').addEventListener('click', manageSources);
  $('#cite-bib').addEventListener('click', () => updateBibliography(true));
  $('#cite-upd').addEventListener('click', () => updateBibliography(false));
  $('#btn-new').addEventListener('click', newDoc);
  $('#btn-open').addEventListener('click', openDialog);
  $('#btn-save').addEventListener('click', () => save(false));
  $('#btn-saveas').addEventListener('click', () => save(true));
  $('#btn-print').addEventListener('click', () => window.print());
  if (FileIO.onLaunch) FileIO.onLaunch((f, h) => loadFile(f, h));
  FileIO.onDrop(async f => { if (f.type.startsWith('image/')) { restore(); await insertImageFile(f); } else await loadFile(f, null); });
  window.addEventListener('beforeunload', e => { if (state.dirty) { e.preventDefault(); e.returnValue = ''; } });

  Lumen.initRibbon();
  applyPage();
  setTitle('Document1');
  doc.focus();
  updateStats();

  // test / automation hook
  window.LumenWrite = { doc, loadFile, save, getPage: () => page, setPage: (p) => { page = p; applyPage(); }, state, setTrack, insertTracked, markDeleted, loadRecipients, mm, setStyleOv, modifyStyleDialog, sources, citeText, bibEntry, updateBibliography, refreshCitations };
})();
