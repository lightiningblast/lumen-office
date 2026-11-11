/* Lumen Cells - spreadsheet UI (canvas grid) */
(function () {
  'use strict';
  const { $, $$, toast, form, FileIO, baseName } = Lumen;
  const { Workbook, Sheet, XErr, MAXR, MAXC, colName, cellName, isErr, rewriteRefs, shiftFormula, splitRef } = Formula;

  let wb = new Workbook();
  let sh = wb.sheets[0];
  const state = { handle: null, dirty: false };
  const FILE_TYPES = [{ description: 'Excel workbooks', accept: { [XlsxIO.XLSX_MIME]: ['.xlsx'] } }];
  const scroller = $('#scroller'), sizer = $('#sizer'), cv = $('#grid'), ctx = cv.getContext('2d'), editor = $('#editor'), finput = $('#finput'), namebox = $('#namebox');
  const HH = 24;
  let HW = 46;
  let dpr = window.devicePixelRatio || 1;

  /* ================= axes ================= */
  function makeAxis(def, map) {
    const keys = [...map.keys()].sort((a, b) => a - b);
    const pre = [];
    let acc = 0;
    keys.forEach(k => { acc += map.get(k) - def; pre.push(acc); });
    const size = (i) => map.has(i) ? map.get(i) : def;
    const offset = (i) => {
      let lo = 0, hi = keys.length;
      while (lo < hi) { const m = (lo + hi) >> 1; if (keys[m] < i) lo = m + 1; else hi = m; }
      return i * def + (lo ? pre[lo - 1] : 0);
    };
    const indexAt = (pos, max) => {
      if (pos <= 0) return 0;
      let lo = 0, hi = max - 1;
      while (lo < hi) { const m = (lo + hi + 1) >> 1; if (offset(m) <= pos) lo = m; else hi = m - 1; }
      return lo;
    };
    return { size, offset, indexAt };
  }
  let colAx, rowAx;
  function rebuildAxes() {
    colAx = makeAxis(sh.defColW, sh.colW);
    rowAx = makeAxis(sh.defRowH, sh.rowH);
    const digits = String(Math.max(100, sh.bounds().maxR + 50)).length;
    HW = Math.max(40, 14 + digits * 8);
    sizer.style.width = (HW + colAx.offset(MAXC)) + 'px';
    sizer.style.height = (HH + rowAx.offset(MAXR)) + 'px';
  }

  /* ================= history ================= */
  const hist = { stack: [], i: 0 };
  function pushHist(e) { hist.stack.length = hist.i; hist.stack.push(e); hist.i++; if (hist.stack.length > 200) { hist.stack.shift(); hist.i--; } }
  function undo() { if (hist.i > 0) { hist.stack[--hist.i].undo(); afterHist(); } }
  function redo() { if (hist.i < hist.stack.length) { hist.stack[hist.i++].redo(); afterHist(); } }
  function afterHist() { wb.touch(); if (!wb.sheets.includes(sh)) sh = wb.sheets[Math.min(wb.active, wb.sheets.length - 1)]; markDirty(); rebuildAxes(); renderTabs(); render(); syncBar(); }
  const cloneCell = (c) => { if (!c) return null; const o = {}; if (c.v !== undefined) o.v = c.v; if (c.f != null) o.f = c.f; if (c.s) o.s = c.s; if (c.cv !== undefined) o.cv = c.cv; return o; };

  function applyEdit(fn) {
    const entry = [];
    const record = (s, r, c, nc) => {
      const k = r * MAXC + c;
      const before = cloneCell(s.cells.get(k));
      const after = nc ? cloneCell(nc) : null;
      entry.push({ s, k, before, after });
      if (after) s.cells.set(k, after); else s.cells.delete(k);
    };
    fn(record);
    if (!entry.length) return;
    pushHist({
      undo() { for (let i = entry.length - 1; i >= 0; i--) { const e = entry[i]; if (e.before) e.s.cells.set(e.k, cloneCell(e.before)); else e.s.cells.delete(e.k); } },
      redo() { for (const e of entry) { if (e.after) e.s.cells.set(e.k, cloneCell(e.after)); else e.s.cells.delete(e.k); } },
    });
    wb.touch(); markDirty(); rebuildAxes(); render(); syncBar();
  }
  function snapAll() {
    return {
      active: wb.active,
      sheets: wb.sheets.map(s => ({ s, name: s.name, cells: new Map([...s.cells].map(([k, c]) => [k, cloneCell(c)])), colW: new Map(s.colW), rowH: new Map(s.rowH), merges: s.merges.map(m => ({ ...m })), freeze: { ...s.freeze }, defColW: s.defColW, defRowH: s.defRowH, showGrid: s.showGrid, charts: JSON.parse(JSON.stringify(s.charts || [])), cf: JSON.parse(JSON.stringify(s.cf || [])), dv: JSON.parse(JSON.stringify(s.dv || [])), comments: JSON.parse(JSON.stringify(s.comments || [])), tables: JSON.parse(JSON.stringify(s.tables || [])), sparks: JSON.parse(JSON.stringify(s.sparks || [])), filter: s.filter ? JSON.parse(JSON.stringify(s.filter)) : null })), names: { ...(wb.names || {}) },
    };
  }
  function restoreAll(snap) {
    wb.sheets = snap.sheets.map(x => {
      const s = x.s;
      s.name = x.name; s.cells = new Map([...x.cells].map(([k, c]) => [k, cloneCell(c)])); s.colW = new Map(x.colW); s.rowH = new Map(x.rowH); s.merges = x.merges.map(m => ({ ...m })); s.freeze = { ...x.freeze }; s.defColW = x.defColW; s.defRowH = x.defRowH; s.showGrid = x.showGrid; s.charts = JSON.parse(JSON.stringify(x.charts || [])); s.cf = JSON.parse(JSON.stringify(x.cf || [])); s.dv = JSON.parse(JSON.stringify(x.dv || [])); s.comments = JSON.parse(JSON.stringify(x.comments || [])); s.tables = JSON.parse(JSON.stringify(x.tables || [])); s.sparks = JSON.parse(JSON.stringify(x.sparks || [])); s.filter = x.filter ? JSON.parse(JSON.stringify(x.filter)) : null;
      return s;
    });
    wb.active = Math.min(snap.active, wb.sheets.length - 1);
    wb.names = { ...(snap.names || {}) };
  }
  function structural(fn) {
    const before = snapAll();
    fn();
    const after = snapAll();
    pushHist({ undo() { restoreAll(before); }, redo() { restoreAll(after); } });
    wb.touch(); markDirty(); rebuildAxes(); renderTabs(); render(); syncBar();
  }

  /* ================= selection ================= */
  const S = () => sh.sel;
  function rng() { const s = S(); return { r1: Math.min(s.r, s.er), r2: Math.max(s.r, s.er), c1: Math.min(s.c, s.ec), c2: Math.max(s.c, s.ec) }; }
  function rngExpanded() {
    let g = rng(), ch = true;
    while (ch) { ch = false; for (const m of sh.merges) if (m.r1 <= g.r2 && m.r2 >= g.r1 && m.c1 <= g.c2 && m.c2 >= g.c1) { const n = { r1: Math.min(g.r1, m.r1), r2: Math.max(g.r2, m.r2), c1: Math.min(g.c1, m.c1), c2: Math.max(g.c2, m.c2) }; if (n.r1 !== g.r1 || n.r2 !== g.r2 || n.c1 !== g.c1 || n.c2 !== g.c2) { g = n; ch = true; } } }
    return g;
  }
  function select(r, c, er, ec, keepScroll) {
    const s = S();
    s.r = Math.max(0, Math.min(MAXR - 1, r)); s.c = Math.max(0, Math.min(MAXC - 1, c));
    s.er = er === undefined ? s.r : Math.max(0, Math.min(MAXR - 1, er)); s.ec = ec === undefined ? s.c : Math.max(0, Math.min(MAXC - 1, ec));
    if (!keepScroll) ensureVisible(er === undefined ? s.r : s.er, ec === undefined ? s.c : s.ec);
    syncBar(); render();
  }
  function getCell(r, c) { return sh.get(r, c); }
  function value(r, c, s2) { const cell = (s2 || sh).get(r, c); if (!cell) return null; if (cell.f != null) return Formula.cellValue(wb, s2 || sh, r, c); return cell.v === undefined ? null : cell.v; }
  function rawText(cell) {
    if (!cell) return '';
    if (cell.f != null) return '=' + cell.f;
    const v = cell.v;
    if (v === undefined || v === null) return '';
    if (v instanceof XErr) return v.code;
    if (typeof v === 'number') {
      const nf = wb.styles[cell.s || 0].nf;
      if (nf && NumFmt.isDateFormat(nf)) return NumFmt.format(v, Math.floor(v) === v ? 'yyyy-mm-dd' : 'yyyy-mm-dd hh:mm:ss');
      if (nf && /%/.test(nf) && !NumFmt.isDateFormat(nf)) return NumFmt.format(v, '0.##########%').replace(/\.%$/, '%');
      return String(+v.toPrecision(15));
    }
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return String(v);
  }
  function display(cell, r, c) {
    if (!cell) return '';
    const v = cell.f != null ? Formula.cellValue(wb, sh, r, c) : cell.v;
    if (v === undefined || v === null) return '';
    if (isErr(v)) return v.code;
    const nf = wb.styles[cell.s || 0].nf;
    if (typeof v === 'number') return NumFmt.format(v, nf);
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (nf && nf.includes('@')) return NumFmt.format(String(v), nf);
    return String(v);
  }
  function a1(rg) { const s = cellName(rg.r1, rg.c1); return rg.r1 === rg.r2 && rg.c1 === rg.c2 ? s : s + ':' + cellName(rg.r2, rg.c2); }

  function ensureVisible(r, c) {
    const W = scroller.clientWidth, H = scroller.clientHeight;
    const fw = colAx.offset(sh.freeze.c), fh = rowAx.offset(sh.freeze.r);
    if (c >= sh.freeze.c) {
      const x0 = colAx.offset(c), x1 = x0 + colAx.size(c);
      const sx = scroller.scrollLeft;
      if (x0 < sx + fw) scroller.scrollLeft = Math.max(0, x0 - fw);
      else if (x1 > sx + W - HW) scroller.scrollLeft = x1 - (W - HW);
    }
    if (r >= sh.freeze.r) {
      const y0 = rowAx.offset(r), y1 = y0 + rowAx.size(r);
      const sy = scroller.scrollTop;
      if (y0 < sy + fh) scroller.scrollTop = Math.max(0, y0 - fh);
      else if (y1 > sy + H - HH) scroller.scrollTop = y1 - (H - HH);
    }
  }

  function syncBar() {
    const g = rng();
    namebox.value = a1(g);
    if (!editing) finput.value = rawText(getCell(S().r, S().c));
    // sync toolbar
    const st = wb.styles[(getCell(S().r, S().c) || {}).s || 0];
    $$('[data-fmt]').forEach(b => b.classList.toggle('active', !!st[b.dataset.fmt]));
    $$('[data-h]').forEach(b => b.classList.toggle('active', st.h === b.dataset.h));
    $$('[data-v]').forEach(b => b.classList.toggle('active', st.v === b.dataset.v));
    $('#b-wrap').classList.toggle('active', !!st.wrap);
    $('#f-name').value = st.font || 'Calibri';
    const fs = $('#f-size');
    const sz = String(st.size || 11);
    if (![...fs.options].some(o => o.value === sz)) { const o = document.createElement('option'); o.textContent = sz; fs.appendChild(o); }
    fs.value = sz;
    const nfSel = $('#nf-select');
    nfSel.value = [...nfSel.options].some(o => o.value === (st.nf || 'General')) ? (st.nf || 'General') : 'General';
    // status
    let sum = 0, n = 0, cnt = 0;
    const cap = 200000;
    let seen = 0;
    outer: for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
      if (++seen > cap) break outer;
      const cell = sh.get(r, c);
      if (!cell) continue;
      const v = value(r, c);
      if (v === null || v === undefined || v === '') continue;
      cnt++;
      if (typeof v === 'number') { sum += v; n++; }
    }
    $('#st-sel').textContent = cnt > 1 ? `Average: ${n ? NumFmt.general(sum / n) : '-'}    Count: ${cnt}    Sum: ${NumFmt.general(sum)}` : '';
  }

  /* ================= rendering ================= */
  let raf = 0;
  function render() { if (!raf) raf = requestAnimationFrame(draw); }
  function resize() {
    const w = scroller.clientWidth, h = scroller.clientHeight;
    dpr = window.devicePixelRatio || 1;
    cv.width = Math.max(1, w * dpr); cv.height = Math.max(1, h * dpr);
    cv.style.width = w + 'px'; cv.style.height = h + 'px';
    render();
  }
  new ResizeObserver(resize).observe(scroller);
  scroller.addEventListener('scroll', () => { if (editing) positionEditor(); render(); });

  function fontOf(st) { return `${st.i ? 'italic ' : ''}${st.b ? 'bold ' : ''}${((st.size || 11) * 4 / 3).toFixed(2)}px ${st.font ? '"' + st.font + '"' : 'Calibri'}, "Segoe UI", sans-serif`; }
  const hex = (c) => '#' + c;

  function visibleList(ax, count, frozenN, scroll, extent, headerPx) {
    const fixed = [], moving = [];
    for (let i = 0; i < frozenN; i++) fixed.push({ i, p: headerPx + ax.offset(i), s: ax.size(i) });
    const fOff = ax.offset(frozenN);
    let i0 = Math.max(frozenN, ax.indexAt(scroll + fOff, count));
    for (let i = i0; i < count; i++) {
      const p = headerPx + ax.offset(i) - scroll;
      if (p > extent) break;
      const s = ax.size(i);
      if (s > 0 || true) moving.push({ i, p, s });
    }
    return { fixed, moving };
  }

  function draw() {
    raf = 0;
    if (!colAx) rebuildAxes();
    const W = scroller.clientWidth, H = scroller.clientHeight;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, W, H);
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, W, H);
    const sx = scroller.scrollLeft, sy = scroller.scrollTop;
    const fc = sh.freeze.c, fr = sh.freeze.r;
    const fw = colAx.offset(fc), fh = rowAx.offset(fr);
    const cols = visibleList(colAx, MAXC, fc, sx, W, HW);
    const rows = visibleList(rowAx, MAXR, fr, sy, H, HH);
    const groups = [
      [rows.moving, cols.moving, HW + fw, HH + fh, W, H, false, false],
      [rows.moving, cols.fixed, HW, HH + fh, HW + fw, H, true, false],
      [rows.fixed, cols.moving, HW + fw, HH, W, HH + fh, false, true],
      [rows.fixed, cols.fixed, HW, HH, HW + fw, HH + fh, true, true],
    ];
    for (const [rl, cl, x0, y0, x1, y1, fx, fy] of groups) {
      if (!rl.length || !cl.length || x1 <= x0 || y1 <= y0) continue;
      ctx.save();
      ctx.beginPath(); ctx.rect(x0, y0, x1 - x0, y1 - y0); ctx.clip();
      drawGroup(rl, cl, fx, fy, sx, sy, x0, y0, x1, y1);
      ctx.restore();
    }
    // selection
    drawSelection(sx, sy, fw, fh, W, H);
    // freeze lines
    ctx.strokeStyle = '#8a8a8a'; ctx.lineWidth = 1;
    if (fc) { ctx.beginPath(); ctx.moveTo(HW + fw + .5, HH); ctx.lineTo(HW + fw + .5, H); ctx.stroke(); }
    if (fr) { ctx.beginPath(); ctx.moveTo(HW, HH + fh + .5); ctx.lineTo(W, HH + fh + .5); ctx.stroke(); }
    drawHeaders(cols, rows, W, H, fw, fh);
    drawFilterArrows(W, H);
    drawCommentMarks(W, H);
    syncCharts();
    syncDv();
  }

  /* ================= charts ================= */
  const chartLayer = $('#charts');
  const chartEls = new Map();
  let selChart = null;
  function syncCharts() {
    chartLayer.style.left = HW + 'px'; chartLayer.style.top = HH + 'px';
    chartLayer.style.width = Math.max(0, scroller.clientWidth - HW) + 'px'; chartLayer.style.height = Math.max(0, scroller.clientHeight - HH) + 'px';
    const sx = scroller.scrollLeft, sy = scroller.scrollTop;
    const live = new Set();
    for (const c of sh.charts) {
      live.add(c.id);
      let el = chartEls.get(c.id);
      if (!el) { el = document.createElement('div'); el.className = 'chart'; el.dataset.id = c.id; el.innerHTML = '<canvas></canvas><i class="ch-h"></i>'; chartLayer.appendChild(el); chartEls.set(c.id, el); }
      el.style.left = (c.x - sx) + 'px'; el.style.top = (c.y - sy) + 'px'; el.style.width = c.w + 'px'; el.style.height = c.h + 'px';
      el.classList.toggle('sel', selChart === c.id);
      const key = wb.ver + '|' + dpr + '|' + JSON.stringify(c);
      if (el._key !== key) {
        el._key = key;
        const cvs = el.firstChild;
        cvs.width = c.w * dpr; cvs.height = c.h * dpr; cvs.style.width = c.w + 'px'; cvs.style.height = c.h + 'px';
        const cx = cvs.getContext('2d');
        cx.setTransform(dpr, 0, 0, dpr, 0, 0);
        Charts.draw(cx, c.w, c.h, c, Charts.chartData(wb, c, sh));
      }
    }
    for (const [id, el] of chartEls) if (!live.has(id)) { el.remove(); chartEls.delete(id); }
  }
  const refStr = (s, r1, c1, r2, c2) => Formula.quoteSheet(s.name) + '!$' + colName(c1) + '$' + (r1 + 1) + (r1 === r2 && c1 === c2 ? '' : ':$' + colName(c2) + '$' + (r2 + 1));
  function regionAround(g) {
    if (g.r1 !== g.r2 || g.c1 !== g.c2) return g;
    const b = sh.bounds();
    const ne = (r, c) => { const v = value(r, c); return v !== null && v !== undefined && v !== ''; };
    let { r1, r2, c1, c2 } = g, ch = true;
    while (ch) {
      ch = false;
      const row = (rr) => { for (let c = Math.max(0, c1 - 1); c <= c2 + 1; c++) if (ne(rr, c)) return true; return false; };
      const col = (cc) => { for (let r = Math.max(0, r1 - 1); r <= r2 + 1; r++) if (ne(r, cc)) return true; return false; };
      if (r1 > 0 && row(r1 - 1)) { r1--; ch = true; }
      if (r2 < b.maxR && row(r2 + 1)) { r2++; ch = true; }
      if (c1 > 0 && col(c1 - 1)) { c1--; ch = true; }
      if (c2 < b.maxC && col(c2 + 1)) { c2++; ch = true; }
    }
    return { r1, r2, c1, c2 };
  }
  function makeChart(type) {
    const g = regionAround(rng());
    const isNum = (r, c) => typeof value(r, c) === 'number';
    const isStr = (r, c) => typeof value(r, c) === 'string' && value(r, c) !== '';
    let hasHeader = false;
    for (let c = g.c1; c <= g.c2; c++) if (isStr(g.r1, c) && g.r2 > g.r1 && isNum(g.r1 + 1, c)) hasHeader = true;
    if (!hasHeader && g.r2 > g.r1) { let strs = 0; for (let c = g.c1; c <= g.c2; c++) if (isStr(g.r1, c)) strs++; if (strs === g.c2 - g.c1 + 1 && g.c2 > g.c1) hasHeader = true; }
    const d0 = g.r1 + (hasHeader ? 1 : 0);
    let hasCat = false;
    if (g.c2 > g.c1 || !isNum(d0, g.c1)) { let strs = 0, n = 0; for (let r = d0; r <= g.r2; r++) { n++; if (!isNum(r, g.c1)) strs++; } hasCat = n > 0 && strs >= n / 2; }
    const sc0 = g.c1 + (hasCat ? 1 : 0);
    const series = [];
    for (let c = sc0; c <= g.c2; c++) {
      let any = false; for (let r = d0; r <= g.r2; r++) if (isNum(r, c)) any = true;
      if (!any) continue;
      series.push({ tx: hasHeader ? refStr(sh, g.r1, c, g.r1, c) : '', cat: hasCat ? refStr(sh, d0, g.c1, g.r2, g.c1) : '', val: refStr(sh, d0, c, g.r2, c), color: Charts.PALETTE[series.length % Charts.PALETTE.length] });
    }
    if (!series.length) { toast('Select a range of cells with numbers first.'); return; }
    const ch = { id: 'c' + Math.random().toString(36).slice(2, 8), type, title: series.length === 1 && series[0].tx ? (value(g.r1, sc0) || '') + '' : 'Chart Title', legend: type === 'pie' || series.length > 1, x: Math.round(colAx.offset(g.c2 + 1) + 24), y: Math.round(rowAx.offset(g.r1)), w: 480, h: 300, series: type === 'pie' ? series.slice(0, 1) : series };
    structural(() => { sh.charts.push(ch); });
    selChart = ch.id; render();
  }
  function chartHist(before) { const after = snapAll(); pushHist({ undo() { restoreAll(before); }, redo() { restoreAll(after); } }); markDirty(); }
  let chartDrag = null;
  chartLayer.addEventListener('mousedown', e => {
    const el = e.target.closest('.chart');
    if (!el) return;
    e.preventDefault();
    if (editing) commitEdit(0, 0);
    const c = sh.charts.find(x => x.id === el.dataset.id);
    selChart = c.id;
    chartDrag = { c, mode: e.target.classList.contains('ch-h') ? 'size' : 'move', x0: e.clientX, y0: e.clientY, ox: c.x, oy: c.y, ow: c.w, oh: c.h, before: snapAll(), moved: false };
    render();
  });
  window.addEventListener('mousemove', e => {
    if (!chartDrag) return;
    const d = chartDrag, dx = e.clientX - d.x0, dy = e.clientY - d.y0;
    if (Math.abs(dx) + Math.abs(dy) > 2) d.moved = true;
    if (d.mode === 'move') { d.c.x = Math.max(0, d.ox + dx); d.c.y = Math.max(0, d.oy + dy); } else { d.c.w = Math.max(120, d.ow + dx); d.c.h = Math.max(90, d.oh + dy); }
    render();
  });
  window.addEventListener('mouseup', () => { if (chartDrag && chartDrag.moved) chartHist(chartDrag.before); chartDrag = null; });
  chartLayer.addEventListener('dblclick', async e => {
    const el = e.target.closest('.chart');
    if (!el) return;
    const c = sh.charts.find(x => x.id === el.dataset.id);
    const v = await form('Edit chart', [{ id: 'title', label: 'Title', value: c.title }, { id: 'type', label: 'Type', type: 'select', value: c.type, options: [['col', 'Column'], ['bar', 'Bar'], ['line', 'Line'], ['area', 'Area'], ['pie', 'Pie']] }, { id: 'legend', label: 'Show legend', type: 'checkbox', value: c.legend !== false }, { id: 'range', label: 'New data range (optional)', value: '' }], 'OK');
    if (!v) return;
    structural(() => {
      c.title = v.title; c.type = v.type; c.legend = v.legend;
      const rg = v.range && Formula.parseRange(v.range.trim());
      if (rg) {
        const first = c.series[0] || {};
        const hdr = !!first.tx;
        const hasCat = !!first.cat;
        const d0 = rg.r1 + (hdr ? 1 : 0), sc0 = rg.c1 + (hasCat ? 1 : 0), ser = [];
        for (let k = sc0; k <= rg.c2; k++) ser.push({ tx: hdr ? refStr(sh, rg.r1, k, rg.r1, k) : '', cat: hasCat ? refStr(sh, d0, rg.c1, rg.r2, rg.c1) : '', val: refStr(sh, d0, k, rg.r2, k), color: Charts.PALETTE[ser.length % Charts.PALETTE.length] });
        if (ser.length) c.series = c.type === 'pie' ? ser.slice(0, 1) : ser;
      }
    });
  });
  function deleteChart() { if (!selChart) return false; const id = selChart; selChart = null; structural(() => { sh.charts = sh.charts.filter(x => x.id !== id); }); return true; }
  const retargetCharts = (fn) => { for (const s of wb.sheets) for (const c of (s.charts || [])) for (const se of c.series) for (const k of ['tx', 'cat', 'val']) if (se[k]) se[k] = rewriteRefs(se[k], ref => fn(ref, s)); };

  function drawHeaders(cols, rows, W, H, fw, fh) {
    const g = rng();
    const selAllC = g.r1 === 0 && g.r2 === MAXR - 1, selAllR = g.c1 === 0 && g.c2 === MAXC - 1;
    ctx.font = '12px "Segoe UI", sans-serif';
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.save(); ctx.beginPath(); ctx.rect(HW, 0, W - HW, HH); ctx.clip();
    const drawCol = (it, frozen) => {
      const inSel = it.i >= g.c1 && it.i <= g.c2;
      if (!frozen && it.p + it.s < HW + fw) return;
      ctx.fillStyle = inSel ? (selAllC ? '#a8d1b9' : '#d3e5da') : '#f3f3f3';
      ctx.fillRect(it.p, 0, it.s, HH);
      ctx.strokeStyle = '#d0d0d0'; ctx.lineWidth = 1;
      ctx.beginPath(); ctx.moveTo(it.p + it.s - .5, 0); ctx.lineTo(it.p + it.s - .5, HH); ctx.moveTo(it.p, HH - .5); ctx.lineTo(it.p + it.s, HH - .5); ctx.stroke();
      if (inSel) { ctx.fillStyle = '#217346'; ctx.fillRect(it.p, HH - 2, it.s, 2); }
      ctx.fillStyle = inSel ? '#185c37' : '#444';
      if (it.s > 0) ctx.fillText(colName(it.i), it.p + it.s / 2, HH / 2);
    };
    cols.moving.forEach(it => drawCol(it, false));
    cols.fixed.forEach(it => drawCol(it, true));
    ctx.restore();
    ctx.save(); ctx.beginPath(); ctx.rect(0, HH, HW, H - HH); ctx.clip();
    const drawRow = (it, frozen) => {
      const inSel = it.i >= g.r1 && it.i <= g.r2;
      if (!frozen && it.p + it.s < HH + fh) return;
      ctx.fillStyle = inSel ? (selAllR ? '#a8d1b9' : '#d3e5da') : '#f3f3f3';
      ctx.fillRect(0, it.p, HW, it.s);
      ctx.strokeStyle = '#d0d0d0';
      ctx.beginPath(); ctx.moveTo(0, it.p + it.s - .5); ctx.lineTo(HW, it.p + it.s - .5); ctx.moveTo(HW - .5, it.p); ctx.lineTo(HW - .5, it.p + it.s); ctx.stroke();
      if (inSel) { ctx.fillStyle = '#217346'; ctx.fillRect(HW - 2, it.p, 2, it.s); }
      ctx.fillStyle = inSel ? '#185c37' : '#444';
      if (it.s > 8) ctx.fillText(String(it.i + 1), HW / 2, it.p + it.s / 2);
    };
    rows.moving.forEach(it => drawRow(it, false));
    rows.fixed.forEach(it => drawRow(it, true));
    ctx.restore();
    ctx.fillStyle = '#f3f3f3'; ctx.fillRect(0, 0, HW, HH);
    ctx.strokeStyle = '#d0d0d0'; ctx.strokeRect(.5, .5, HW - 1, HH - 1);
    ctx.fillStyle = '#b5b5b5'; ctx.beginPath(); ctx.moveTo(HW - 4, HH - 14); ctx.lineTo(HW - 4, HH - 4); ctx.lineTo(HW - 14, HH - 4); ctx.closePath(); ctx.fill();
  }

  function wrapLines(text, maxW, font) {
    ctx.font = font;
    const out = [];
    for (const para of String(text).split('\n')) {
      const words = para.split(/(\s+)/);
      let line = '';
      for (const w of words) {
        if (ctx.measureText(line + w).width > maxW && line.trim()) { out.push(line.trimEnd()); line = w.trimStart(); } else line += w;
      }
      out.push(line);
    }
    return out;
  }

  function drawGroup(rows, cols, fx, fy, sx, sy, gx0, gy0, gx1, gy1) {
    const xOf = (c) => HW + colAx.offset(c) - (fx ? 0 : sx);
    const yOf = (r) => HH + rowAx.offset(r) - (fy ? 0 : sy);
    const rMin = rows[0].i, rMax = rows[rows.length - 1].i, cMin = cols[0].i, cMax = cols[cols.length - 1].i;
    // gridlines
    if (sh.showGrid) {
      ctx.strokeStyle = '#e1e1e1'; ctx.lineWidth = 1;
      ctx.beginPath();
      for (const c of cols) { const x = c.p + c.s - .5; ctx.moveTo(x, gy0); ctx.lineTo(x, gy1); }
      for (const r of rows) { const y = r.p + r.s - .5; ctx.moveTo(gx0, y); ctx.lineTo(gx1, y); }
      ctx.stroke();
    }
    // table header fill and banded rows
    const tbls = (sh.tables || []).filter(t => t.r1 <= rMax && t.r2 >= rMin && t.c1 <= cMax && t.c2 >= cMin);
    for (const t of tbls) {
      const tc = tableColors(t.style), h0 = t.hdr !== false ? 1 : 0;
      for (const r of rows) {
        if (r.i < t.r1 || r.i > t.r2) continue;
        const isHdr = h0 && r.i === t.r1, isTot = t.totals && r.i === t.r2;
        if (!isHdr && (isTot || (r.i - t.r1 - h0) % 2 !== 0)) continue;
        ctx.fillStyle = hex(isHdr ? tc.h : tc.b);
        for (const c of cols) if (c.i >= t.c1 && c.i <= t.c2) ctx.fillRect(c.p, r.p, c.s, r.s);
      }
    }
    const merges = sh.merges.filter(m => m.r1 <= rMax && m.r2 >= rMin && m.c1 <= cMax && m.c2 >= cMin);
    const covered = (r, c) => { for (const m of merges) if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return true; return false; };
    const items = [];
    // cells present in view
    const visCount = rows.length * cols.length;
    if (visCount < sh.cells.size) {
      for (const r of rows) for (const c of cols) { const cell = sh.cells.get(r.i * MAXC + c.i); if (cell && !covered(r.i, c.i)) items.push({ cell, r: r.i, c: c.i, x: c.p, y: r.p, w: c.s, h: r.s }); }
    } else {
      const rSet = new Map(rows.map(r => [r.i, r])), cSet = new Map(cols.map(c => [c.i, c]));
      for (const [k, cell] of sh.cells) { const r = Math.floor(k / MAXC), c = k - r * MAXC; const ri = rSet.get(r), ci = cSet.get(c); if (ri && ci && !covered(r, c)) items.push({ cell, r, c, x: ci.p, y: ri.p, w: ci.s, h: ri.s }); }
    }
    for (const sp of Formula.spills(wb, sh)) {
      for (const ri of rows) {
        const i = ri.i - sp.r;
        if (i < 0 || i >= sp.rows.length) continue;
        for (const ci of cols) {
          const j = ci.i - sp.c;
          if (j < 0 || j >= sp.rows[0].length || (i === 0 && j === 0) || sh.get(ri.i, ci.i) || covered(ri.i, ci.i)) continue;
          items.push({ cell: { v: sp.rows[i][j] === undefined ? null : sp.rows[i][j], spill: true }, r: ri.i, c: ci.i, x: ci.p, y: ri.p, w: ci.s, h: ri.s });
        }
      }
    }
    for (const m of merges) {
      const cell = sh.get(m.r1, m.c1) || {};
      const x = xOf(m.c1), y = yOf(m.r1);
      items.push({ cell, r: m.r1, c: m.c1, x, y, w: xOf(m.c2 + 1) - x, h: yOf(m.r2 + 1) - y, merged: true });
    }
    // fills
    const hasCf = sh.cf && sh.cf.length;
    if (hasCf) for (const it of items) it.ov = CondFmt.evaluate(wb, sh, it.r, it.c, it.cell);
    for (const it of items) {
      const st = wb.styles[it.cell.s || 0];
      const bg = (it.ov && it.ov.style.bg) || st.bg;
      if (bg) { ctx.fillStyle = hex(bg); ctx.fillRect(it.x, it.y, it.w, it.h); }
      if (it.ov && it.ov.bar) { const bw = Math.max(1, (it.w - 6) * it.ov.bar.pct); ctx.fillStyle = hex(it.ov.bar.color); ctx.globalAlpha = 0.75; ctx.fillRect(it.x + 2, it.y + 3, bw, it.h - 6); ctx.globalAlpha = 1; }
    }
    // text
    for (const it of items) {
      const cell = it.cell;
      let st = it.ov && Object.keys(it.ov.style).length ? Object.assign({}, wb.styles[cell.s || 0], it.ov.style) : wb.styles[cell.s || 0];
      if (tbls.length) {
        const t = tbls.find(t => it.r >= t.r1 && it.r <= t.r2 && it.c >= t.c1 && it.c <= t.c2);
        if (t && t.hdr !== false && it.r === t.r1) st = Object.assign({}, st, { b: true, color: st.color || 'FFFFFF' });
        else if (t && t.totals && it.r === t.r2) st = Object.assign({}, st, { b: true });
      }
      const text = display(cell, it.r, it.c);
      if (text !== '') drawText(cell, st, text, it, gx0, gx1, merges.length ? covered : null);
    }
    // sparklines
    for (const sp of (sh.sparks || [])) {
      const ri = rows.find(r => r.i === sp.r), ci = cols.find(c => c.i === sp.c);
      if (ri && ci) drawSpark(sp, ci.p, ri.p, ci.s, ri.s);
    }
    // borders
    for (const it of items) {
      const st = wb.styles[it.cell.s || 0];
      for (const [k, side] of [['bl', 'l'], ['br', 'r'], ['bt', 't'], ['bb', 'b']]) {
        if (!st[k]) continue;
        const [style, color] = st[k].split(':');
        ctx.strokeStyle = hex(color || '000000');
        ctx.lineWidth = style === 'medium' ? 2 : style === 'thick' ? 3 : 1;
        ctx.setLineDash(style === 'dashed' || style === 'mediumDashed' ? [4, 3] : style === 'dotted' || style === 'hair' ? [1.5, 2] : []);
        ctx.beginPath();
        const x0 = it.x, y0 = it.y, x1 = it.x + it.w, y1 = it.y + it.h;
        if (side === 'l') { ctx.moveTo(x0 + .5, y0); ctx.lineTo(x0 + .5, y1); }
        if (side === 'r') { ctx.moveTo(x1 - .5, y0); ctx.lineTo(x1 - .5, y1); }
        if (side === 't') { ctx.moveTo(x0, y0 + .5); ctx.lineTo(x1, y0 + .5); }
        if (side === 'b') { ctx.moveTo(x0, y1 - .5); ctx.lineTo(x1, y1 - .5); }
        ctx.stroke();
        ctx.setLineDash([]);
      }
    }
  }

  function drawText(cell, st, text, it, gx0, gx1, covered) {
    const v = cell.f != null ? Formula.cellValue(wb, sh, it.r, it.c) : cell.v;
    const isNum = typeof v === 'number';
    let h = st.h || (isNum ? 'right' : (typeof v === 'boolean' || isErr(v)) ? 'center' : 'left');
    const font = fontOf(st);
    ctx.font = font;
    ctx.fillStyle = st.color ? hex(st.color) : '#000';
    const pad = 4 + (st.indent ? st.indent * 9 : 0);
    const fpx = (st.size || 11) * 4 / 3;
    const { x, y, w, h: hgt } = it;
    ctx.textAlign = h === 'center' ? 'center' : h === 'right' ? 'right' : 'left';
    const px = h === 'center' ? x + w / 2 : h === 'right' ? x + w - pad : x + pad;
    let lines;
    if (st.wrap) lines = wrapLines(text, w - 8, font); else lines = [text];
    const lh = fpx * 1.25;
    let tw = ctx.measureText(lines[0]).width;
    if (!st.wrap && isNum && tw > w - 4 && (!st.nf || st.nf === 'General') && Math.abs(v) < 1e11) {
      for (let p = 9; p >= 1; p--) { const t = Math.abs(v) >= 1 || p > 1 ? String(+v.toPrecision(p)) : String(Math.round(v)); if (t.indexOf('e') < 0) { lines = [t]; tw = ctx.measureText(t).width; if (tw <= w - 4) break; } }
    }
    if (!st.wrap && isNum && tw > w - 4) { lines = ['#'.repeat(Math.max(1, Math.floor((w - 4) / (fpx * 0.55))))]; tw = 0; }
    ctx.save();
    ctx.beginPath();
    let cx0 = x, cx1 = x + w;
    if (!st.wrap && !it.merged && tw > w - 2 * pad && !isNum) {
      // overflow into empty neighbours
      if (h === 'left') { let c = it.c + 1, extra = 0; while (cx1 + extra < gx1 && c < MAXC && !(sh.cells.get(it.r * MAXC + c) && display(sh.cells.get(it.r * MAXC + c), it.r, c) !== '')) { extra += colAx.size(c); c++; if (x + w + extra - x > tw + pad * 2) break; } cx1 += extra; }
      else if (h === 'right') { let c = it.c - 1, extra = 0; while (cx0 - extra > gx0 && c >= 0 && !(sh.cells.get(it.r * MAXC + c) && display(sh.cells.get(it.r * MAXC + c), it.r, c) !== '')) { extra += colAx.size(c); c--; if (w + extra > tw + pad * 2) break; } cx0 -= extra; }
    }
    ctx.rect(cx0, y, cx1 - cx0, hgt); ctx.clip();
    const total = lines.length * lh;
    const va = st.v || 'bottom';
    let ty = va === 'top' ? y + 3 : va === 'center' ? y + (hgt - total) / 2 : y + hgt - total - 2;
    ctx.textBaseline = 'top';
    lines.forEach((ln, i) => {
      const yy = ty + i * lh + (lh - fpx) / 2 - 1;
      ctx.fillText(ln, px, yy);
      if (st.u || st.s) {
        const lw = ctx.measureText(ln).width;
        const lx = h === 'center' ? px - lw / 2 : h === 'right' ? px - lw : px;
        ctx.fillRect(lx, st.u ? yy + fpx + 1 : yy + fpx * 0.55, lw, 1);
      }
    });
    ctx.restore();
  }

  function cellRectView(r, c, r2, c2) {
    const fc = sh.freeze.c, fr = sh.freeze.r, sx = scroller.scrollLeft, sy = scroller.scrollTop;
    const x = HW + colAx.offset(c) - (c >= fc ? sx : 0), y = HH + rowAx.offset(r) - (r >= fr ? sy : 0);
    const x2 = HW + colAx.offset(c2 + 1) - (c2 + 1 > fc ? sx : 0), y2 = HH + rowAx.offset(r2 + 1) - (r2 + 1 > fr ? sy : 0);
    return { x, y, w: x2 - x, h: y2 - y };
  }

  function drawSelection(sx, sy, fw, fh, W, H) {
    const g = rngExpanded();
    const a = rng();
    ctx.save();
    ctx.beginPath(); ctx.rect(HW, HH, W - HW, H - HH); ctx.clip();
    const R = cellRectView(g.r1, g.c1, Math.min(g.r2, MAXR - 1), Math.min(g.c2, MAXC - 1));
    const multi = a.r1 !== a.r2 || a.c1 !== a.c2;
    if (multi) {
      ctx.fillStyle = 'rgba(33,115,70,.12)';
      ctx.fillRect(R.x, R.y, R.w, R.h);
      // clear tint on active cell
      const A = cellRectView(S().r, S().c, S().r, S().c);
      const m = sh.mergeAt(S().r, S().c);
      const AA = m ? cellRectView(m.r1, m.c1, m.r2, m.c2) : A;
      ctx.fillStyle = '#fff';
      ctx.fillRect(AA.x, AA.y, AA.w, AA.h);
      const cell = getCell(S().r, S().c);
      if (cell) { const st = wb.styles[cell.s || 0]; if (st.bg) { ctx.fillStyle = hex(st.bg); ctx.fillRect(AA.x, AA.y, AA.w, AA.h); } const t = display(cell, S().r, S().c); if (t) { ctx.save(); ctx.beginPath(); ctx.rect(AA.x, AA.y, AA.w, AA.h); ctx.clip(); drawText(cell, st, t, { r: S().r, c: S().c, x: AA.x, y: AA.y, w: AA.w, h: AA.h, merged: true }, HW, W, null); ctx.restore(); } }
    }
    ctx.strokeStyle = '#217346'; ctx.lineWidth = 2;
    ctx.strokeRect(R.x + 1, R.y + 1, R.w - 2, R.h - 2);
    ctx.fillStyle = '#217346';
    ctx.fillRect(R.x + R.w - 4, R.y + R.h - 4, 7, 7);
    ctx.strokeStyle = '#fff'; ctx.lineWidth = 1; ctx.strokeRect(R.x + R.w - 4.5, R.y + R.h - 4.5, 8, 8);
    ctx.restore();
    // copy marquee
    if (clip && clip.sheet === sh) {
      const C = cellRectView(clip.r1, clip.c1, clip.r2, clip.c2);
      ctx.save(); ctx.beginPath(); ctx.rect(HW, HH, W - HW, H - HH); ctx.clip();
      ctx.setLineDash([5, 4]); ctx.strokeStyle = '#217346'; ctx.lineWidth = 1.5; ctx.strokeRect(C.x + .5, C.y + .5, C.w - 1, C.h - 1); ctx.restore();
    }
  }

  /* ================= hit testing & mouse ================= */
  function hit(clientX, clientY) {
    const rect = scroller.getBoundingClientRect();
    const x = clientX - rect.left, y = clientY - rect.top;
    const W = scroller.clientWidth, H = scroller.clientHeight;
    if (x >= W || y >= H || x < 0 || y < 0) return { zone: 'bar' };
    const sx = scroller.scrollLeft, sy = scroller.scrollTop;
    const fc = sh.freeze.c, fr = sh.freeze.r, fw = colAx.offset(fc), fh = rowAx.offset(fr);
    const colAtX = (px) => { const d = px - HW; return d < fw ? colAx.indexAt(d, MAXC) : colAx.indexAt(d + sx, MAXC); };
    const rowAtY = (py) => { const d = py - HH; return d < fh ? rowAx.indexAt(d, MAXR) : rowAx.indexAt(d + sy, MAXR); };
    const c = x >= HW ? colAtX(x) : -1, r = y >= HH ? rowAtY(y) : -1;
    const o = { zone: 'cell', r, c, x, y };
    if (x < HW && y < HH) { o.zone = 'corner'; return o; }
    if (y < HH) {
      o.zone = 'colh';
      const edgeX = (i) => HW + colAx.offset(i + 1) - (i >= fc ? sx : 0);
      if (Math.abs(edgeX(c) - x) <= 4) o.edge = c;
      else if (c > 0 && Math.abs(edgeX(c - 1) - x) <= 4) o.edge = c - 1;
    } else if (x < HW) {
      o.zone = 'rowh';
      const edgeY = (i) => HH + rowAx.offset(i + 1) - (i >= fr ? sy : 0);
      if (Math.abs(edgeY(r) - y) <= 4) o.edge = r;
      else if (r > 0 && Math.abs(edgeY(r - 1) - y) <= 4) o.edge = r - 1;
    } else {
      // fill handle
      const g = rngExpanded();
      const R = cellRectView(g.r1, g.c1, g.r2, g.c2);
      if (Math.abs(x - (R.x + R.w)) <= 5 && Math.abs(y - (R.y + R.h)) <= 5) o.zone = 'fill';
    }
    return o;
  }

  let drag = null;
  scroller.addEventListener('mousedown', e => {
    const h = hit(e.clientX, e.clientY);
    if (h.zone === 'bar') return;
    if (selChart) { selChart = null; render(); }
    if (sh.filter) {
      const rc = scroller.getBoundingClientRect();
      const fc = filterArrowHit(e.clientX - rc.left, e.clientY - rc.top);
      if (fc >= 0) { e.preventDefault(); openFilterMenu(fc); return; }
    }
    // editing a formula: insert reference instead of committing
    if (editing && h.zone === 'cell' && refMode()) { e.preventDefault(); insertRefAt(h.r, h.c); return; }
    if (editing) commitEdit(0, 0);
    scroller.focus();
    if (h.zone === 'colh' && h.edge !== undefined) { e.preventDefault(); drag = { t: 'resizeCol', i: h.edge, x0: e.clientX, w0: colAx.size(h.edge) }; return; }
    if (h.zone === 'rowh' && h.edge !== undefined) { e.preventDefault(); drag = { t: 'resizeRow', i: h.edge, y0: e.clientY, h0: rowAx.size(h.edge) }; return; }
    e.preventDefault();
    if (h.zone === 'corner') { select(0, 0, MAXR - 1, MAXC - 1, true); return; }
    if (h.zone === 'colh') {
      if (e.shiftKey) select(S().r, S().c, MAXR - 1, h.c, true); else select(0, h.c, MAXR - 1, h.c, true);
      drag = { t: 'cols', c0: e.shiftKey ? S().c : h.c }; return;
    }
    if (h.zone === 'rowh') {
      if (e.shiftKey) select(S().r, S().c, h.r, MAXC - 1, true); else select(h.r, 0, h.r, MAXC - 1, true);
      drag = { t: 'rows', r0: e.shiftKey ? S().r : h.r }; return;
    }
    if (h.zone === 'fill') { const g = rngExpanded(); drag = { t: 'fill', src: g }; return; }
    if (e.shiftKey) select(S().r, S().c, h.r, h.c, true);
    else select(h.r, h.c);
    drag = { t: 'cells' };
  });
  window.addEventListener('mousemove', e => {
    if (!drag) {
      if (e.target === scroller) {
        const h = hit(e.clientX, e.clientY);
        scroller.style.cursor = (h.zone === 'colh' && h.edge !== undefined) ? 'col-resize' : (h.zone === 'rowh' && h.edge !== undefined) ? 'row-resize' : h.zone === 'fill' ? 'crosshair' : (h.zone === 'colh' || h.zone === 'rowh') ? 'pointer' : 'cell';
      }
      return;
    }
    if (drag.t === 'resizeCol') { sh.colW.set(drag.i, Math.max(0, drag.w0 + e.clientX - drag.x0)); rebuildAxes(); render(); if (editing) positionEditor(); return; }
    if (drag.t === 'resizeRow') { sh.rowH.set(drag.i, Math.max(0, drag.h0 + e.clientY - drag.y0)); rebuildAxes(); render(); return; }
    const h = hit(e.clientX, e.clientY);
    const rect = scroller.getBoundingClientRect();
    // auto-scroll
    if (e.clientX > rect.right - 24 - 17) scroller.scrollLeft += 20; else if (e.clientX < rect.left + HW + 6) scroller.scrollLeft -= 20;
    if (e.clientY > rect.bottom - 24 - 17) scroller.scrollTop += 20; else if (e.clientY < rect.top + HH + 6) scroller.scrollTop -= 20;
    if (h.zone === 'bar') return;
    const x = Math.max(HW, Math.min(h.x, scroller.clientWidth - 1)), y = Math.max(HH, Math.min(h.y, scroller.clientHeight - 1));
    const hh = hit(rect.left + x, rect.top + y);
    if (drag.t === 'cells' && hh.r >= 0 && hh.c >= 0) select(S().r, S().c, hh.r, hh.c, true);
    else if (drag.t === 'cols' && hh.c >= 0) select(0, drag.c0, MAXR - 1, hh.c, true);
    else if (drag.t === 'rows' && hh.r >= 0) select(drag.r0, 0, hh.r, MAXC - 1, true);
    else if (drag.t === 'fill' && hh.r >= 0 && hh.c >= 0) { drag.to = { r: hh.r, c: hh.c }; select(drag.src.r1, drag.src.c1, Math.max(drag.src.r2, hh.r), Math.max(drag.src.c2, hh.c), true); }
  });
  window.addEventListener('mouseup', () => {
    if (drag && (drag.t === 'resizeCol' || drag.t === 'resizeRow')) { markDirty(); }
    if (drag && drag.t === 'fill' && drag.to) {
      const s = drag.src, to = drag.to;
      const down = to.r > s.r2, right = to.c > s.c2;
      if (down || right) { fillRange(s, down && (!right || to.r - s.r2 >= to.c - s.c2) ? 'down' : 'right', down ? to.r : s.r2, right ? to.c : s.c2); }
    }
    drag = null;
  });
  scroller.addEventListener('dblclick', e => {
    const h = hit(e.clientX, e.clientY);
    if (h.zone === 'colh' && h.edge !== undefined) { autoFitCol(h.edge); return; }
    if (h.zone === 'cell' && h.r >= 0) startEdit(null);
  });

  function autoFitCol(c) {
    const b = sh.bounds();
    let mx = 0;
    ctx.save();
    for (let r = 0; r <= b.maxR; r++) {
      const cell = sh.get(r, c);
      if (!cell) continue;
      const t = display(cell, r, c);
      if (!t) continue;
      ctx.font = fontOf(wb.styles[cell.s || 0]);
      mx = Math.max(mx, ctx.measureText(t.split('\n')[0]).width);
    }
    ctx.restore();
    sh.colW.set(c, Math.max(24, Math.min(600, Math.ceil(mx + 12))));
    rebuildAxes(); render(); markDirty();
  }

  /* ================= editing ================= */
  let editing = null;
  function positionEditor() {
    if (!editing) return;
    const m = sh.mergeAt(editing.r, editing.c);
    const R = m ? cellRectView(m.r1, m.c1, m.r2, m.c2) : cellRectView(editing.r, editing.c, editing.r, editing.c);
    const st = wb.styles[(getCell(editing.r, editing.c) || {}).s || 0];
    editor.style.left = (R.x - 1) + 'px'; editor.style.top = (R.y - 1) + 'px';
    editor.style.minWidth = (R.w + 1) + 'px'; editor.style.height = (R.h + 1) + 'px';
    editor.style.font = fontOf(st);
    editor.style.color = st.color ? hex(st.color) : '#000';
    const needW = Math.max(R.w + 1, measure(editor.value, editor.style.font) + 14);
    editor.style.width = Math.min(needW, scroller.clientWidth - R.x) + 'px';
  }
  function measure(t, font) { ctx.save(); ctx.font = font; const w = Math.max(...String(t).split('\n').map(l => ctx.measureText(l).width)); ctx.restore(); return w; }
  function startEdit(initial) {
    if (editing) return;
    const r = S().r, c = S().c;
    const m = sh.mergeAt(r, c);
    const ar = m ? m.r1 : r, ac = m ? m.c1 : c;
    editing = { r: ar, c: ac, orig: rawText(getCell(ar, ac)), typed: initial !== null && initial !== undefined };
    editor.value = initial !== null && initial !== undefined ? initial : editing.orig;
    editor.style.display = 'block';
    positionEditor();
    editor.focus();
    const L = editor.value.length;
    editor.setSelectionRange(L, L);
    finput.value = editor.value;
  }
  function cancelEdit() { if (!editing) return; editing = null; editor.style.display = 'none'; scroller.focus(); syncBar(); }
  function commitEdit(dr, dc) {
    if (!editing) return true;
    const { r, c, orig } = editing;
    const text = editor.value;
    if (text !== orig) {
      failReason = '';
      if (!setFromText(r, c, text)) { if (failReason !== 'dv') toast('That formula has a syntax error. Fix it or press Esc to cancel.', 3500); return false; }
    }
    editing = null; editor.style.display = 'none';
    scroller.focus();
    if (dr || dc) moveSel(dr, dc, false);
    syncBar(); render();
    return true;
  }
  function setFromText(r, c, text) {
    let cell;
    const old = getCell(r, c);
    const st = wb.styles[(old && old.s) || 0];
    if (text.trim() === '') cell = old && old.s ? { s: old.s } : null;
    else if (text[0] === '=' && text.length > 1 && st.nf !== '@') {
      try { Formula.parse(text.slice(1)); } catch (e) { return false; }
      cell = { f: text.slice(1).trim() };
      if (old && old.s) cell.s = old.s;
    } else if (text[0] === "'") { cell = { v: text.slice(1) }; if (old && old.s) cell.s = old.s; }
    else if (st.nf === '@') { cell = { v: text }; if (old && old.s) cell.s = old.s; }
    else {
      const p = NumFmt.parseInput(text);
      cell = { v: p.v };
      let s = old && old.s ? old.s : 0;
      if (p.nf && (!st.nf || st.nf === 'General')) s = wb.styleId({ ...wb.styles[s], nf: p.nf });
      if (s) cell.s = s;
    }
    if (!(text[0] === '=' && text.length > 1 && st.nf !== '@') && cell && cell.v !== undefined) {
      const bad = Validate.check(wb, sh, r, c, cell.v, text);
      if (bad) {
        failReason = 'dv';
        const msg = (bad.title ? bad.title + ': ' : '') + Validate.message(bad);
        if ((bad.style || 'stop') === 'stop') { toast(msg, 4500); return false; }
        toast(msg + ' (accepted)', 4500);
      }
    } else if (text.trim() === '' && sh.dv && sh.dv.length) {
      const bad = Validate.check(wb, sh, r, c, null, '');
      if (bad && (bad.style || 'stop') === 'stop') { failReason = 'dv'; toast(Validate.message(bad), 4500); return false; }
    }
    applyEdit(rec => rec(sh, r, c, cell));
    return true;
  }
  let failReason = '';
  const OPS = '=+-*/(,^&<>:;';
  function refMode() { if (!editing) return false; const v = editor.value; if (v[0] !== '=') return false; const pos = editor.selectionStart; const p = v[pos - 1]; return p !== undefined && OPS.includes(p) || (editor.selectionStart === editor.selectionEnd && /[A-Za-z]\d*$/.test(v.slice(0, pos)) && false); }
  function insertRefAt(r, c) {
    const pos = editor.selectionStart;
    const t = cellName(r, c);
    editor.value = editor.value.slice(0, pos) + t + editor.value.slice(editor.selectionEnd);
    editor.setSelectionRange(pos + t.length, pos + t.length);
    finput.value = editor.value;
    positionEditor();
  }
  editor.addEventListener('input', () => { finput.value = editor.value; positionEditor(); });
  editor.addEventListener('keydown', e => {
    if (e.key === 'Enter' && !e.altKey) { e.preventDefault(); commitEdit(e.shiftKey ? -1 : 1, 0); }
    else if (e.key === 'Enter' && e.altKey) { const p = editor.selectionStart; editor.value = editor.value.slice(0, p) + '\n' + editor.value.slice(editor.selectionEnd); editor.setSelectionRange(p + 1, p + 1); e.preventDefault(); positionEditor(); }
    else if (e.key === 'Tab') { e.preventDefault(); commitEdit(0, e.shiftKey ? -1 : 1); }
    else if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
    else if (e.key === 'F2') { e.preventDefault(); editing.typed = false; }
    else if (editing.typed && !e.shiftKey && !e.ctrlKey && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key) && !refMode() === true) {
      e.preventDefault();
      const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
      commitEdit(d[0], d[1]);
    }
  });
  finput.addEventListener('focus', () => { if (!editing) { startEdit(null); finput.focus(); } });
  finput.addEventListener('input', () => { if (editing) { editor.value = finput.value; positionEditor(); } });
  finput.addEventListener('keydown', e => {
    if (e.key === 'Enter') { e.preventDefault(); editor.value = finput.value; commitEdit(1, 0); }
    else if (e.key === 'Escape') { cancelEdit(); finput.blur(); }
  });
  namebox.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
      let txt = namebox.value.trim();
      if (wb.names && wb.names[txt.toUpperCase()]) txt = wb.names[txt.toUpperCase()].replace(/^(?:'[^']+'|[^'!]+)!/, '');
      const rg = Formula.parseRange(txt);
      if (rg) { select(rg.r1, rg.c1, rg.r2, rg.c2); scroller.focus(); } else { toast('Not a valid cell reference'); syncBar(); }
    }
  });

  /* ================= navigation & keys ================= */
  function moveSel(dr, dc, extend) {
    const s = S();
    if (extend) { select(s.r, s.c, s.er + dr, s.ec + dc); return; }
    // skip merged
    let r = s.r + dr, c = s.c + dc;
    const m = sh.mergeAt(s.r, s.c);
    if (m) { if (dr > 0) r = m.r2 + 1; if (dc > 0) c = m.c2 + 1; if (dr < 0) r = m.r1 - 1; if (dc < 0) c = m.c1 - 1; }
    const m2 = sh.mergeAt(Math.max(0, r), Math.max(0, c));
    if (m2) { r = m2.r1; c = m2.c1; }
    select(r, c);
  }
  function jump(dr, dc, extend) {
    const s = S();
    let r = extend ? s.er : s.r, c = extend ? s.ec : s.c;
    const b = sh.bounds();
    const has = (rr, cc) => { const v = value(rr, cc); return v !== null && v !== undefined && v !== ''; };
    const lim = (rr, cc) => rr < 0 || cc < 0 || rr >= MAXR || cc >= MAXC;
    let nr = r + dr, nc = c + dc;
    if (lim(nr, nc)) return;
    if (has(r, c) && has(nr, nc)) { while (!lim(nr + dr, nc + dc) && has(nr + dr, nc + dc)) { nr += dr; nc += dc; } }
    else { while (!lim(nr, nc) && !has(nr, nc) && nr <= Math.max(b.maxR, 0) + (dr > 0 ? MAXR : 0) && nc <= (dc > 0 ? MAXC : b.maxC + 1) ) { if (dr < 0 && nr === 0) break; if (dc < 0 && nc === 0) break; nr += dr; nc += dc; if (dr > 0 && nr > b.maxR) { nr = MAXR - 1; break; } if (dc > 0 && nc > b.maxC) { nc = MAXC - 1; break; } } }
    nr = Math.max(0, Math.min(MAXR - 1, nr)); nc = Math.max(0, Math.min(MAXC - 1, nc));
    if (extend) select(s.r, s.c, nr, nc); else select(nr, nc);
  }

  scroller.addEventListener('keydown', e => {
    if (editing) return;
    const ctrl = e.ctrlKey || e.metaKey;
    const k = e.key;
    const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[k]) { e.preventDefault(); const [dr, dc] = arrows[k]; if (ctrl) jump(dr, dc, e.shiftKey); else moveSel(dr, dc, e.shiftKey); return; }
    if (k === 'Tab') { e.preventDefault(); moveSel(0, e.shiftKey ? -1 : 1, false); return; }
    if (k === 'Enter') { e.preventDefault(); const g = rng(); if (g.r1 !== g.r2 || g.c1 !== g.c2) { const s = S(); let r = s.r + (e.shiftKey ? -1 : 1), c = s.c; if (r > g.r2) { r = g.r1; c++; } if (r < g.r1) { r = g.r2; c--; } if (c > g.c2) c = g.c1; const sv = { er: s.er, ec: s.ec }; s.r = r; s.c = c; ensureVisible(r, c); syncBar(); render(); } else moveSel(e.shiftKey ? -1 : 1, 0, false); return; }
    if (k === 'F2') { e.preventDefault(); startEdit(null); return; }
    if (k === 'Home') { e.preventDefault(); if (ctrl) select(0, 0); else select(S().r, 0); return; }
    if (k === 'End' && ctrl) { e.preventDefault(); const b = sh.bounds(); select(Math.max(0, b.maxR), Math.max(0, b.maxC)); return; }
    if (k === 'PageDown' || k === 'PageUp') { e.preventDefault(); const n = Math.max(1, Math.floor((scroller.clientHeight - HH) / sh.defRowH) - 1) * (k === 'PageUp' ? -1 : 1); moveSel(n, 0, e.shiftKey); return; }
    if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); if (!deleteChart()) clearContents(); return; }
    if (ctrl) {
      const lk = k.toLowerCase();
      if (lk === 'z') { e.preventDefault(); e.shiftKey ? redo() : undo(); return; }
      if (lk === 'y') { e.preventDefault(); redo(); return; }
      if (lk === 'b') { e.preventDefault(); toggleFmt('b'); return; }
      if (lk === 'i') { e.preventDefault(); toggleFmt('i'); return; }
      if (lk === 'u') { e.preventDefault(); toggleFmt('u'); return; }
      if (lk === 'a') { e.preventDefault(); select(0, 0, MAXR - 1, MAXC - 1, true); return; }
      if (lk === 'd') { e.preventDefault(); fillSel('down'); return; }
      if (lk === 'r') { e.preventDefault(); fillSel('right'); return; }
      if (lk === 'f') { e.preventDefault(); doFind(); return; }
      if (lk === 'h') { e.preventDefault(); doReplace(); return; }
      if (lk === 's') { e.preventDefault(); save(e.shiftKey); return; }
      if (lk === 'o') { e.preventDefault(); openDialog(); return; }
      if (lk === 'c' || lk === 'v' || lk === 'x') return; // handled by clipboard events
      return;
    }
    if (k.length === 1 && !e.altKey) { e.preventDefault(); startEdit(k); }
  });
  // global shortcuts when focus is elsewhere (e.g. ribbon)
  document.addEventListener('keydown', e => {
    if ((e.ctrlKey || e.metaKey) && !editing && e.target !== scroller && !/^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName)) {
      const lk = e.key.toLowerCase();
      if (lk === 's') { e.preventDefault(); save(e.shiftKey); } else if (lk === 'o') { e.preventDefault(); openDialog(); } else if (lk === 'z') { e.preventDefault(); undo(); } else if (lk === 'y') { e.preventDefault(); redo(); }
    }
  });

  /* ================= clipboard ================= */
  let clip = null;
  function tsvOf(g) {
    const lines = [];
    for (let r = g.r1; r <= g.r2; r++) {
      const row = [];
      for (let c = g.c1; c <= g.c2; c++) { let t = display(getCell(r, c), r, c); if (/[\t\n"]/.test(t)) t = '"' + t.replace(/"/g, '""') + '"'; row.push(t); }
      lines.push(row.join('\t'));
    }
    return lines.join('\n');
  }
  function htmlOf(g) {
    let h = '<table>';
    for (let r = g.r1; r <= g.r2; r++) { h += '<tr>'; for (let c = g.c1; c <= g.c2; c++) h += '<td>' + Lumen.esc(display(getCell(r, c), r, c)) + '</td>'; h += '</tr>'; }
    return h + '</table>';
  }
  function doCopy(e, cut) {
    if (editing || /^(INPUT|TEXTAREA)$/.test((e.target || {}).tagName || '')) return;
    const g = rngExpanded();
    const gg = { r1: g.r1, c1: g.c1, r2: Math.min(g.r2, sh.bounds().maxR < 0 ? 0 : Math.max(sh.bounds().maxR, g.r1)), c2: Math.min(g.c2, Math.max(sh.bounds().maxC, g.c1)) };
    const cells = new Map();
    for (let r = gg.r1; r <= gg.r2; r++) for (let c = gg.c1; c <= gg.c2; c++) { const cell = getCell(r, c); if (cell) cells.set((r - gg.r1) * MAXC + (c - gg.c1), cloneCell(cell)); }
    const tsv = tsvOf(gg);
    clip = { ...gg, cells, cut: !!cut, tsv, sheet: sh };
    e.clipboardData.setData('text/plain', tsv);
    e.clipboardData.setData('text/html', htmlOf(gg));
    e.preventDefault();
    render();
  }
  document.addEventListener('copy', e => doCopy(e, false));
  document.addEventListener('cut', e => doCopy(e, true));
  document.addEventListener('paste', e => {
    if (editing || /^(INPUT|TEXTAREA|SELECT)$/.test((e.target || {}).tagName || '')) return;
    e.preventDefault();
    const text = (e.clipboardData || window.clipboardData).getData('text/plain');
    const g = rng();
    if (clip && text === clip.tsv) {
      const dr = g.r1 - clip.r1, dc = g.c1 - clip.c1;
      const h = clip.r2 - clip.r1, w = clip.c2 - clip.c1;
      applyEdit(rec => {
        if (clip.cut) for (let r = clip.r1; r <= clip.r2; r++) for (let c = clip.c1; c <= clip.c2; c++) rec(clip.sheet, r, c, null);
        for (let r = 0; r <= h; r++) for (let c = 0; c <= w; c++) {
          const src = clip.cells.get(r * MAXC + c);
          const tr = g.r1 + r, tc = g.c1 + c;
          if (tr >= MAXR || tc >= MAXC) continue;
          if (!src) { if (!clip.cut) rec(sh, tr, tc, null); continue; }
          const n = cloneCell(src);
          if (n.f != null && !clip.cut) n.f = shiftFormula(n.f, dr, dc);
          delete n.cv;
          rec(sh, tr, tc, n);
        }
      });
      if (clip.cut) clip = null;
      select(g.r1, g.c1, g.r1 + h, g.c1 + w);
    } else if (text) {
      const rows = parseTSV(text);
      applyEdit(rec => {
        rows.forEach((row, i) => row.forEach((t, j) => {
          const r = g.r1 + i, c = g.c1 + j;
          if (r >= MAXR || c >= MAXC) return;
          const old = getCell(r, c);
          let cell;
          if (t === '') cell = old && old.s ? { s: old.s } : null;
          else if (t[0] === '=' && t.length > 1) cell = { f: t.slice(1) };
          else { const p = NumFmt.parseInput(t); cell = { v: p.v }; let s = old && old.s || 0; if (p.nf && !wb.styles[s].nf) s = wb.styleId({ ...wb.styles[s], nf: p.nf }); if (s) cell.s = s; }
          if (cell && old && old.s && !cell.s) cell.s = old.s;
          rec(sh, r, c, cell);
        }));
      });
      select(g.r1, g.c1, g.r1 + rows.length - 1, g.c1 + Math.max(...rows.map(r => r.length)) - 1);
    }
  });
  function parseTSV(text, delim) {
    delim = delim || '\t';
    const rows = [];
    let row = [], cur = '', q = false;
    text = text.replace(/\r\n?/g, '\n');
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) { if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += ch; }
      else if (ch === '"' && cur === '') q = true;
      else if (ch === delim) { row.push(cur); cur = ''; }
      else if (ch === '\n') { row.push(cur); rows.push(row); row = []; cur = ''; }
      else cur += ch;
    }
    if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
    return rows;
  }

  /* ================= cell operations ================= */
  function clampRows(g) { const b = sh.bounds(); return { r1: g.r1, c1: g.c1, r2: Math.min(g.r2, g.r2 === MAXR - 1 ? Math.max(b.maxR, 200) : g.r2), c2: Math.min(g.c2, g.c2 === MAXC - 1 ? Math.max(b.maxC, 26) : g.c2) }; }
  function clearContents() {
    const g = clampRows(rngExpanded());
    applyEdit(rec => { for (const [k, cell] of [...sh.cells]) { const r = Math.floor(k / MAXC), c = k - r * MAXC; if (r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2) { if (cell.s) rec(sh, r, c, { s: cell.s }); else rec(sh, r, c, null); } } });
  }
  function formatSel(patch) {
    const g = clampRows(rngExpanded());
    applyEdit(rec => {
      for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
        const old = getCell(r, c);
        const st = wb.styles[(old && old.s) || 0];
        const ns = wb.styleId({ ...st, ...patch });
        if (!old && ns === 0) continue;
        if ((old && (old.s || 0)) === ns) continue;
        const n = old ? cloneCell(old) : {};
        n.s = ns; if (!ns) delete n.s;
        if (!old && ns === 0) continue;
        rec(sh, r, c, n);
      }
    });
  }
  function toggleFmt(k) { const st = wb.styles[(getCell(S().r, S().c) || {}).s || 0]; formatSel({ [k]: !st[k] }); }
  function setBorders(kind) {
    const g = rngExpanded();
    const b = 'thin:000000';
    const cl = clampRows(g);
    applyEdit(rec => {
      for (let r = cl.r1; r <= cl.r2; r++) for (let c = cl.c1; c <= cl.c2; c++) {
        const old = getCell(r, c);
        const st = { ...wb.styles[(old && old.s) || 0] };
        const edge = { l: c === cl.c1, r: c === cl.c2, t: r === cl.r1, b: r === cl.r2 };
        if (kind === 'none') { delete st.bl; delete st.br; delete st.bt; delete st.bb; }
        else if (kind === 'all') { st.bl = st.br = st.bt = st.bb = b; }
        else if (kind === 'outer') { if (edge.l) st.bl = b; if (edge.r) st.br = b; if (edge.t) st.bt = b; if (edge.b) st.bb = b; }
        else if (kind === 'bottom' && edge.b) st.bb = b;
        else if (kind === 'top' && edge.t) st.bt = b;
        else if (kind === 'left' && edge.l) st.bl = b;
        else if (kind === 'right' && edge.r) st.br = b;
        const ns = wb.styleId(st);
        if ((old && old.s || 0) === ns) continue;
        const n = old ? cloneCell(old) : {};
        n.s = ns; if (!ns) delete n.s;
        rec(sh, r, c, n);
      }
    });
  }
  function changeDecimals(delta) {
    const st = wb.styles[(getCell(S().r, S().c) || {}).s || 0];
    let nf = st.nf || 'General';
    if (nf === 'General') nf = '0';
    const m = /0(?:\.(0+))?/.exec(nf.replace(/\[[^\]]*\]|"[^"]*"/g, m2 => ' '.repeat(m2.length)));
    if (!m) return;
    const cur = m[1] ? m[1].length : 0;
    const nn = Math.max(0, cur + delta);
    const idx = m.index;
    nf = nf.slice(0, idx) + '0' + (nn ? '.' + '0'.repeat(nn) : '') + nf.slice(idx + m[0].length);
    formatSel({ nf });
  }
  function fillRange(src, dir, toR, toC) {
    applyEdit(rec => {
      if (dir === 'down') {
        const h = src.r2 - src.r1 + 1;
        for (let r = src.r2 + 1; r <= toR; r++) for (let c = src.c1; c <= src.c2; c++) {
          const sr = src.r1 + ((r - src.r1) % h);
          const sc = getCell(sr, c);
          if (!sc) { rec(sh, r, c, null); continue; }
          const n = cloneCell(sc); delete n.cv;
          if (n.f != null) n.f = shiftFormula(n.f, r - sr, 0);
          else if (typeof n.v === 'number' && h === 1 && src.r1 === src.r2 && false) n.v += r - sr;
          rec(sh, r, c, n);
        }
      } else {
        const w = src.c2 - src.c1 + 1;
        for (let c = src.c2 + 1; c <= toC; c++) for (let r = src.r1; r <= src.r2; r++) {
          const scol = src.c1 + ((c - src.c1) % w);
          const sc = getCell(r, scol);
          if (!sc) { rec(sh, r, c, null); continue; }
          const n = cloneCell(sc); delete n.cv;
          if (n.f != null) n.f = shiftFormula(n.f, 0, c - scol);
          rec(sh, r, c, n);
        }
      }
    });
    select(src.r1, src.c1, dir === 'down' ? toR : src.r2, dir === 'right' ? toC : src.c2);
  }
  function fillSel(dir) { const g = rngExpanded(); if (dir === 'down') { if (g.r2 > g.r1) fillRange({ ...g, r2: g.r1 }, 'down', g.r2, g.c2); } else if (g.c2 > g.c1) fillRange({ ...g, c2: g.c1 }, 'right', g.r2, g.c2); }

  function autoSum() {
    const g = rng();
    const s = S();
    let target, ref;
    if (g.r1 === g.r2 && g.c1 === g.c2) {
      // look above
      let r = s.r - 1;
      while (r >= 0 && typeof value(r, s.c) === 'number') r--;
      if (r < s.r - 1) { ref = a1({ r1: r + 1, c1: s.c, r2: s.r - 1, c2: s.c }); }
      else { let c = s.c - 1; while (c >= 0 && typeof value(s.r, c) === 'number') c--; if (c < s.c - 1) ref = a1({ r1: s.r, c1: c + 1, r2: s.r, c2: s.c - 1 }); }
      if (!ref) { startEdit('=SUM('); return; }
      applyEdit(rec => rec(sh, s.r, s.c, { f: 'SUM(' + ref + ')', s: (getCell(s.r, s.c) || {}).s }));
    } else {
      applyEdit(rec => { for (let c = g.c1; c <= g.c2; c++) rec(sh, g.r2 + 1, c, { f: 'SUM(' + a1({ r1: g.r1, c1: c, r2: g.r2, c2: c }) + ')' }); });
    }
  }

  /* ---------- structure ---------- */
  function mapRefAxis(ref, ownerSheet, axis, at, n, del) {
    // returns undefined (unchanged), null (#REF!) or new ref
    const target = ref.sheet ? wb.sheetByName(ref.sheet) : ownerSheet;
    if (target !== sh) return undefined;
    const key = axis === 'r' ? 'r' : 'c';
    const isOnlyOther = (p) => (axis === 'r' ? p.col : p.row);
    const a = { ...ref.a }, b = ref.b ? { ...ref.b } : null;
    if (!ref.b) {
      if (isOnlyOther(a)) return undefined;
      const v = a[key];
      if (del) { if (v >= at && v < at + n) return null; if (v >= at + n) a[key] = v - n; }
      else if (v >= at) a[key] = v + n;
      return { a, b: null };
    }
    if (isOnlyOther(a) && isOnlyOther(b)) return undefined;
    let s = a[key], e = b[key];
    if (s === undefined || e === undefined) return undefined;
    if (s > e) [s, e] = [e, s];
    if (del) {
      const ns = s < at ? s : (s >= at + n ? s - n : at);
      const ne = e < at ? e : (e >= at + n ? e - n : at - 1);
      if (ne < ns) return null;
      a[key] = ns; b[key] = ne;
    } else { a[key] = s >= at ? s + n : s; b[key] = e >= at ? e + n : e; }
    return { a, b };
  }
  function shiftStructure(axis, at, n, del) {
    structural(() => {
      for (const s of wb.sheets) {
        for (const cell of s.cells.values()) if (cell.f != null) { cell.f = rewriteRefs(cell.f, ref => mapRefAxis(ref, s, axis, at, n, del)); delete cell._ast; }
      }
      retargetCharts((ref, s) => mapRefAxis(ref, s, axis, at, n, del));
      const moved = new Map();
      for (const [k, cell] of sh.cells) {
        let r = Math.floor(k / MAXC), c = k - r * MAXC;
        let v = axis === 'r' ? r : c;
        if (del) { if (v >= at && v < at + n) continue; if (v >= at + n) v -= n; }
        else if (v >= at) v += n;
        if (axis === 'r') r = v; else c = v;
        if (r < MAXR && c < MAXC) moved.set(r * MAXC + c, cell);
      }
      sh.cells = moved;
      const map = axis === 'r' ? sh.rowH : sh.colW;
      const nm = new Map();
      for (const [i, w] of map) { let v = i; if (del) { if (v >= at && v < at + n) continue; if (v >= at + n) v -= n; } else if (v >= at) v += n; nm.set(v, w); }
      if (axis === 'r') sh.rowH = nm; else sh.colW = nm;
      sh.merges = sh.merges.map(m => {
        const k1 = axis === 'r' ? 'r1' : 'c1', k2 = axis === 'r' ? 'r2' : 'c2';
        const o = { ...m };
        if (del) {
          const s0 = m[k1], e0 = m[k2];
          o[k1] = s0 < at ? s0 : (s0 >= at + n ? s0 - n : at);
          o[k2] = e0 < at ? e0 : (e0 >= at + n ? e0 - n : at - 1);
          if (o[k2] < o[k1]) return null;
        } else { if (m[k1] >= at) o[k1] += n; if (m[k2] >= at) o[k2] += n; }
        return o;
      }).filter(m => m && (m.r1 !== m.r2 || m.c1 !== m.c2));
      sh.sparks = (sh.sparks || []).map(m => {
        const k = axis === 'r' ? 'r' : 'c';
        const o = { ...m };
        if (del) { if (m[k] >= at && m[k] < at + n) return null; if (m[k] >= at + n) o[k] -= n; } else if (m[k] >= at) o[k] += n;
        return o;
      }).filter(Boolean);
      for (const key of ['cf', 'dv', 'tables']) {
        sh[key] = (sh[key] || []).map(m => {
          const k1 = axis === 'r' ? 'r1' : 'c1', k2 = axis === 'r' ? 'r2' : 'c2';
          const o = { ...m };
          if (del) {
            const s0 = m[k1], e0 = m[k2];
            o[k1] = s0 < at ? s0 : (s0 >= at + n ? s0 - n : at);
            o[k2] = e0 < at ? e0 : (e0 >= at + n ? e0 - n : at - 1);
            if (o[k2] < o[k1]) return null;
          } else { if (m[k1] >= at) o[k1] += n; if (m[k2] >= at) o[k2] += n; }
          return o;
        }).filter(Boolean);
      }
      sh.comments = (sh.comments || []).map(m => {
        const k = axis === 'r' ? 'r' : 'c';
        const o = { ...m };
        if (del) { if (m[k] >= at && m[k] < at + n) return null; if (m[k] >= at + n) o[k] -= n; } else if (m[k] >= at) o[k] += n;
        return o;
      }).filter(Boolean);
      if (sh.filter) {
        const f = sh.filter, k1 = axis === 'r' ? 'r1' : 'c1', k2 = axis === 'r' ? 'r2' : 'c2';
        if (!del) { if (f[k1] >= at) f[k1] += n; if (f[k2] >= at) f[k2] += n; }
        else if (f[k1] >= at && f[k1] < at + n) sh.filter = null;
        else { if (f[k1] >= at + n) f[k1] -= n; if (f[k2] >= at + n) f[k2] -= n; else if (f[k2] >= at) f[k2] = at - 1; }
      }
      if (axis === 'r' && sh.freeze.r > at) sh.freeze.r = Math.max(0, sh.freeze.r + (del ? -Math.min(n, sh.freeze.r - at) : n));
    });
  }

  function mergeSel() {
    const g = rngExpanded();
    if (g.r1 === g.r2 && g.c1 === g.c2) return;
    structural(() => {
      sh.merges = sh.merges.filter(m => !(m.r1 <= g.r2 && m.r2 >= g.r1 && m.c1 <= g.c2 && m.c2 >= g.c1));
      for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) if (r !== g.r1 || c !== g.c1) { const cell = sh.get(r, c); if (cell) { if (cell.s) sh.cells.set(r * MAXC + c, { s: cell.s }); else sh.cells.delete(r * MAXC + c); } }
      sh.merges.push({ ...g });
    });
    select(g.r1, g.c1, g.r2, g.c2);
  }
  function unmergeSel() { const g = rngExpanded(); structural(() => { sh.merges = sh.merges.filter(m => !(m.r1 <= g.r2 && m.r2 >= g.r1 && m.c1 <= g.c2 && m.c2 >= g.c1)); }); }

  /* ---------- sort / find ---------- */
  function sortSel(asc, opts) {
    let g = rng();
    const b = sh.bounds();
    if (g.r1 === g.r2 && g.c1 === g.c2) {
      // expand to contiguous region
      const nonEmpty = (r, c) => { const v = value(r, c); return v !== null && v !== undefined && v !== ''; };
      let { r1, r2, c1, c2 } = { r1: g.r1, r2: g.r2, c1: g.c1, c2: g.c2 };
      let changed = true;
      while (changed) {
        changed = false;
        const any = (rr, cc1, cc2) => { for (let c = Math.max(0, cc1); c <= cc2; c++) if (nonEmpty(rr, c)) return true; return false; };
        const anyC = (cc, rr1, rr2) => { for (let r = Math.max(0, rr1); r <= rr2; r++) if (nonEmpty(r, cc)) return true; return false; };
        if (r1 > 0 && any(r1 - 1, c1 - 1, c2 + 1)) { r1--; changed = true; }
        if (r2 < b.maxR && any(r2 + 1, c1 - 1, c2 + 1)) { r2++; changed = true; }
        if (c1 > 0 && anyC(c1 - 1, r1 - 1, r2 + 1)) { c1--; changed = true; }
        if (c2 < b.maxC && anyC(c2 + 1, r1 - 1, r2 + 1)) { c2++; changed = true; }
      }
      g = { r1, r2, c1, c2 };
    }
    const keyCol = Math.min(Math.max(S().c, g.c1), g.c2);
    const first = value(g.r1, keyCol), second = value(Math.min(g.r1 + 1, g.r2), keyCol);
    const header = opts && opts.header !== undefined ? opts.header : (g.r2 > g.r1 && typeof first === 'string' && typeof second !== 'string');
    const r0 = header ? g.r1 + 1 : g.r1;
    if (g.r2 <= r0 - (header ? 0 : 0) && g.r2 === r0) return;
    const rowsData = [];
    for (let r = r0; r <= g.r2; r++) { const cells = []; for (let c = g.c1; c <= g.c2; c++) cells.push(cloneCell(getCell(r, c))); rowsData.push({ r, cells, key: value(r, keyCol) }); }
    const rank = (v) => v === null || v === undefined || v === '' ? 3 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2;
    const cmp = (x, y) => { const rx = rank(x.key), ry = rank(y.key); if (rx === 3 || ry === 3) return rx === ry ? 0 : (rx === 3 ? 1 : -1); if (rx !== ry) return (rx - ry) * (asc ? 1 : -1); let c = rx === 1 ? String(x.key).toLowerCase().localeCompare(String(y.key).toLowerCase()) : (x.key < y.key ? -1 : x.key > y.key ? 1 : 0); return asc ? c : -c; };
    const sorted = rowsData.map((x, i) => ({ ...x, i })).sort((x, y) => cmp(x, y) || x.i - y.i);
    applyEdit(rec => {
      sorted.forEach((row, idx) => {
        const tr = r0 + idx;
        row.cells.forEach((cell, j) => {
          const c = g.c1 + j;
          if (cell && cell.f != null) { cell.f = shiftFormula(cell.f, tr - row.r, 0); delete cell.cv; }
          rec(sh, tr, c, cell);
        });
      });
    });
    select(g.r1, g.c1, g.r2, g.c2);
  }
  async function doFind() {
    const v = await form('Find', [{ id: 'q', label: 'Find what', value: lastFind }], 'Find next');
    if (!v || !v.q) return;
    lastFind = v.q;
    findNext(v.q);
  }
  let lastFind = '';
  function findNext(q) {
    const b = sh.bounds();
    const keys = [...sh.cells.keys()].sort((a, c) => a - c);
    const cur = S().r * MAXC + S().c;
    const needle = q.toLowerCase();
    const matches = (k) => { const r = Math.floor(k / MAXC), c = k - r * MAXC; return display(sh.cells.get(k), r, c).toLowerCase().includes(needle); };
    let hit = keys.find(k => k > cur && matches(k));
    if (hit === undefined) hit = keys.find(k => matches(k));
    if (hit === undefined) { toast('Nothing found for "' + q + '"'); return; }
    const r = Math.floor(hit / MAXC);
    select(r, hit - r * MAXC);
  }
  async function doReplace() {
    const v = await form('Replace', [{ id: 'q', label: 'Find what', value: lastFind }, { id: 'r', label: 'Replace with', value: '' }], 'Replace all');
    if (!v || !v.q) return;
    lastFind = v.q;
    let n = 0;
    const re = new RegExp(v.q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
    applyEdit(rec => {
      for (const [k, cell] of [...sh.cells]) {
        if (cell.f != null || typeof cell.v !== 'string' || !re.test(cell.v)) { re.lastIndex = 0; continue; }
        re.lastIndex = 0;
        const r = Math.floor(k / MAXC), c = k - r * MAXC;
        const nc = cloneCell(cell); nc.v = cell.v.replace(re, () => { n++; return v.r; });
        rec(sh, r, c, nc);
      }
    });
    toast(n + ' replacement' + (n === 1 ? '' : 's'));
  }

  /* ================= sheets ================= */
  function renderTabs() {
    const list = $('#tablist');
    list.innerHTML = '';
    wb.sheets.forEach((s, i) => {
      const d = document.createElement('div');
      d.className = 'stab' + (s === sh ? ' active' : '');
      d.textContent = s.name;
      d.addEventListener('click', () => switchSheet(i));
      d.addEventListener('dblclick', renameSheet);
      list.appendChild(d);
    });
  }
  function switchSheet(i) {
    if (editing) commitEdit(0, 0);
    sh.scroll = { x: scroller.scrollLeft, y: scroller.scrollTop };
    selChart = null;
    wb.active = i; sh = wb.sheets[i];
    if (sh.pivot) refreshPivot(sh, true);
    rebuildAxes(); renderTabs(); syncBar();
    scroller.scrollLeft = sh.scroll.x; scroller.scrollTop = sh.scroll.y;
    $('#v-grid').checked = sh.showGrid;
    render();
  }
  function uniqueName(base) { let n = base, i = 2; while (wb.sheetByName(n)) n = base + ' (' + i++ + ')'; return n; }
  function addSheet() {
    structural(() => { let i = wb.sheets.length + 1; let n; do { n = 'Sheet' + i++; } while (wb.sheetByName(n)); wb.sheets.push(new Sheet(n)); wb.active = wb.sheets.length - 1; sh = wb.sheets[wb.active]; });
  }
  function dupSheet() {
    structural(() => {
      const n = new Sheet(uniqueName(sh.name));
      for (const [k, c] of sh.cells) { const o = cloneCell(c); n.cells.set(k, o); }
      n.charts = JSON.parse(JSON.stringify(sh.charts || [])).map(c => { c.id = 'c' + Math.random().toString(36).slice(2, 8); for (const se of c.series) for (const k of ['tx', 'cat', 'val']) if (se[k]) se[k] = rewriteRefs(se[k], ref => (ref.sheet && ref.sheet === sh.name) ? { a: ref.a, b: ref.b, sheet: n.name } : undefined); return c; });
      n.sparks = JSON.parse(JSON.stringify(sh.sparks || []));
      n.tables = JSON.parse(JSON.stringify(sh.tables || [])).map(t => { t.name = uniqueTableName(); return t; });
      n.dv = JSON.parse(JSON.stringify(sh.dv || [])); n.comments = JSON.parse(JSON.stringify(sh.comments || []));
      n.cf = JSON.parse(JSON.stringify(sh.cf || [])).map(r => ({ ...r, id: 'f' + Math.random().toString(36).slice(2, 8) }));
      n.colW = new Map(sh.colW); n.rowH = new Map(sh.rowH); n.merges = sh.merges.map(m => ({ ...m })); n.freeze = { ...sh.freeze }; n.showGrid = sh.showGrid; n.defColW = sh.defColW; n.defRowH = sh.defRowH;
      wb.sheets.splice(wb.sheets.indexOf(sh) + 1, 0, n); wb.active = wb.sheets.indexOf(n); sh = n;
    });
  }
  async function renameSheet() {
    const v = await form('Rename sheet', [{ id: 'n', label: 'Name', value: sh.name }], 'Rename');
    if (!v) return;
    const nn = v.n.trim().replace(/[\\/?*[\]:]/g, '');
    if (!nn) return;
    if (wb.sheets.some(s => s !== sh && s.name.toLowerCase() === nn.toLowerCase())) { toast('A sheet with that name already exists.'); return; }
    const old = sh.name;
    structural(() => {
      for (const s of wb.sheets) for (const cell of s.cells.values()) if (cell.f != null) { cell.f = rewriteRefs(cell.f, ref => (ref.sheet && ref.sheet.toLowerCase() === old.toLowerCase()) ? { a: ref.a, b: ref.b, sheet: nn } : undefined); delete cell._ast; }
      retargetCharts(ref => (ref.sheet && ref.sheet.toLowerCase() === old.toLowerCase()) ? { a: ref.a, b: ref.b, sheet: nn } : undefined);
      sh.name = nn;
    });
  }
  async function deleteSheet() {
    if (wb.sheets.length < 2) { toast('A workbook needs at least one sheet.'); return; }
    if (!(await Lumen.confirmBox('Delete sheet', 'Delete "' + sh.name + '"? You can undo this with Ctrl+Z.', 'Delete'))) return;
    const old = sh.name;
    structural(() => {
      const i = wb.sheets.indexOf(sh);
      wb.sheets.splice(i, 1);
      retargetCharts(ref => (ref.sheet && ref.sheet.toLowerCase() === old.toLowerCase()) ? null : undefined);
      for (const s of wb.sheets) for (const cell of s.cells.values()) if (cell.f != null) { cell.f = rewriteRefs(cell.f, ref => (ref.sheet && ref.sheet.toLowerCase() === old.toLowerCase()) ? null : undefined); delete cell._ast; }
      wb.active = Math.min(i, wb.sheets.length - 1); sh = wb.sheets[wb.active];
    });
  }

  /* ================= ribbon wiring ================= */
  Lumen.initRibbon();
  $('.ribbon').addEventListener('mousedown', e => { if (!e.target.closest('select,input')) e.preventDefault(); });
  const doIt = (fn) => (...a) => { if (editing) commitEdit(0, 0); fn(...a); scroller.focus(); };
  $('#b-undo').onclick = undo; $('#b-redo').onclick = redo;
  $('#b-copy').onclick = () => document.execCommand('copy');
  $('#b-cut').onclick = () => document.execCommand('cut');
  $$('[data-fmt]').forEach(b => b.onclick = () => toggleFmt(b.dataset.fmt));
  $$('[data-h]').forEach(b => b.onclick = () => { const st = wb.styles[(getCell(S().r, S().c) || {}).s || 0]; formatSel({ h: st.h === b.dataset.h ? '' : b.dataset.h }); });
  $$('[data-v]').forEach(b => b.onclick = () => { const st = wb.styles[(getCell(S().r, S().c) || {}).s || 0]; formatSel({ v: st.v === b.dataset.v ? '' : b.dataset.v }); });
  $('#b-wrap').onclick = () => toggleFmt('wrap');
  $('#f-name').onchange = e => { formatSel({ font: e.target.value === 'Calibri' ? '' : e.target.value }); scroller.focus(); };
  $('#f-size').onchange = e => { formatSel({ size: parseFloat(e.target.value) === 11 ? '' : parseFloat(e.target.value) }); scroller.focus(); };
  $('#f-border').onchange = e => { if (e.target.value) setBorders(e.target.value); e.target.value = ''; scroller.focus(); };
  const colorBind = (inp, sw, apply, key) => { $(inp).addEventListener('input', e => { $(sw).style.background = e.target.value; formatSel({ [key]: e.target.value.slice(1).toUpperCase() }); }); $(apply).onclick = () => formatSel({ [key]: $(inp).value.slice(1).toUpperCase() }); };
  colorBind('#fill', '#sw-fill', '#apply-fill', 'bg');
  colorBind('#color', '#sw-color', '#apply-color', 'color');
  $('#b-nofill').onclick = () => formatSel({ bg: '' });
  $('#nf-select').onchange = e => { formatSel({ nf: e.target.value === 'General' ? '' : e.target.value }); scroller.focus(); };
  $('#nf-cur').onclick = () => formatSel({ nf: '$#,##0.00' });
  $('#nf-pct').onclick = () => formatSel({ nf: '0%' });
  $('#nf-com').onclick = () => formatSel({ nf: '#,##0.00' });
  $('#nf-inc').onclick = () => changeDecimals(1);
  $('#nf-dec').onclick = () => changeDecimals(-1);
  $('#b-merge').onclick = () => { mergeSel(); formatSel({ h: 'center' }); };
  $('#b-unmerge').onclick = unmergeSel;
  $('#b-autosum').onclick = doIt(autoSum);
  $('#b-filld').onclick = () => fillSel('down'); $('#b-fillr').onclick = () => fillSel('right');
  $('#b-clear').onclick = clearContents;
  $('#b-clearfmt').onclick = () => { const g = clampRows(rngExpanded()); applyEdit(rec => { for (const [k, cell] of [...sh.cells]) { const r = Math.floor(k / MAXC), c = k - r * MAXC; if (r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2 && cell.s) { const n = cloneCell(cell); delete n.s; if (n.v === undefined && n.f == null) rec(sh, r, c, null); else rec(sh, r, c, n); } } }); };
  $('#b-find').onclick = doFind; $('#d-replace').onclick = doReplace;
  $('#i-row').onclick = () => { const g = rngExpanded(); shiftStructure('r', g.r1, g.r2 - g.r1 + 1, false); };
  $('#i-col').onclick = () => { const g = rngExpanded(); shiftStructure('c', g.c1, g.c2 - g.c1 + 1, false); };
  $('#d-row').onclick = () => { const g = rngExpanded(); shiftStructure('r', g.r1, g.r2 - g.r1 + 1, true); };
  $('#d-col').onclick = () => { const g = rngExpanded(); shiftStructure('c', g.c1, g.c2 - g.c1 + 1, true); };
  /* ---------- data validation ---------- */
  const DV_OPS = [['between', 'between'], ['notBetween', 'not between'], ['equal', 'equal to'], ['notEqual', 'not equal to'], ['greaterThan', 'greater than'], ['lessThan', 'less than'], ['greaterThanOrEqual', 'greater than or equal to'], ['lessThanOrEqual', 'less than or equal to']];
  async function validationDialog() {
    if (editing) commitEdit(0, 0);
    const g = clampRows(rngExpanded());
    const ex = Validate.ruleAt(sh, S().r, S().c) || {};
    const v = await form('Data validation', [
      { id: 'type', label: 'Allow', type: 'select', value: ex.type || 'list', options: [['list', 'List'], ['whole', 'Whole number'], ['decimal', 'Decimal'], ['date', 'Date'], ['textLength', 'Text length'], ['custom', 'Custom formula']] },
      { id: 'op', label: 'Data', type: 'select', value: ex.op || 'between', options: DV_OPS },
      { id: 'f1', label: 'Value / minimum / source', value: ex.f1 || '' },
      { id: 'f2', label: 'Maximum (for "between")', value: ex.f2 || '' },
      { type: 'note', label: 'List source: items separated by commas (Yes,No,Maybe) or a range such as =$A$1:$A$5. Custom: a formula such as =A1>0.' },
      { id: 'blank', type: 'checkbox', label: 'Ignore blank', value: ex.blank !== false },
      { id: 'drop', type: 'checkbox', label: 'In-cell dropdown (lists)', value: ex.drop !== false },
      { id: 'style', label: 'Error alert', type: 'select', value: ex.style || 'stop', options: [['stop', 'Stop'], ['warning', 'Warning'], ['information', 'Information']] },
      { id: 'title', label: 'Error title', value: ex.title || '' },
      { id: 'msg', label: 'Error message', value: ex.msg || '' },
    ], 'OK');
    if (!v) return;
    if (v.f1 === '' && v.type !== 'custom') { toast('Enter a value or source for the rule.'); return; }
    structural(() => {
      sh.dv = sh.dv.filter(d => !(d.r1 >= g.r1 && d.r2 <= g.r2 && d.c1 >= g.c1 && d.c2 <= g.c2));
      sh.dv.unshift({ id: 'v' + Math.random().toString(36).slice(2, 8), ...g, type: v.type, op: v.op, f1: v.f1, f2: v.f2, blank: v.blank, drop: v.drop, style: v.style, title: v.title, msg: v.msg });
    });
  }
  function clearValidation() { const g = rngExpanded(); structural(() => { sh.dv = sh.dv.filter(d => !(d.r1 <= g.r2 && d.r2 >= g.r1 && d.c1 <= g.c2 && d.c2 >= g.c1)); }); }
  const dvBtn = document.createElement('div'), dvMenu = document.createElement('div');
  dvBtn.id = 'dvbtn'; dvBtn.textContent = '▾'; dvMenu.id = 'dvmenu';
  $('#gridwrap').append(dvBtn, dvMenu);
  function syncDv() {
    const d = Validate.ruleAt(sh, S().r, S().c);
    if (!d || d.type !== 'list' || d.drop === false || editing) { dvBtn.style.display = 'none'; dvMenu.style.display = 'none'; return; }
    const R = cellRectView(S().r, S().c, S().r, S().c);
    if (R.x < HW || R.y < HH || R.x + R.w > scroller.clientWidth || R.y + R.h > scroller.clientHeight) { dvBtn.style.display = 'none'; return; }
    dvBtn.style.display = 'block';
    dvBtn.style.left = (R.x + R.w - 18) + 'px'; dvBtn.style.top = (R.y + 1) + 'px'; dvBtn.style.height = (R.h - 2) + 'px'; dvBtn.style.lineHeight = (R.h - 2) + 'px';
  }
  dvBtn.addEventListener('mousedown', e => {
    e.preventDefault(); e.stopPropagation();
    if (dvMenu.style.display === 'block') { dvMenu.style.display = 'none'; return; }
    const d = Validate.ruleAt(sh, S().r, S().c);
    if (!d) return;
    const r = S().r, c = S().c;
    dvMenu.innerHTML = '';
    for (const it of Validate.listItems(wb, sh, d).slice(0, 300)) {
      const o = document.createElement('div');
      o.textContent = it;
      o.onmousedown = ev => { ev.preventDefault(); dvMenu.style.display = 'none'; failReason = ''; setFromText(r, c, it); scroller.focus(); };
      dvMenu.appendChild(o);
    }
    const R = cellRectView(r, c, r, c);
    dvMenu.style.left = R.x + 'px'; dvMenu.style.top = (R.y + R.h) + 'px'; dvMenu.style.minWidth = R.w + 'px'; dvMenu.style.display = 'block';
  });
  scroller.addEventListener('scroll', () => { dvMenu.style.display = 'none'; });

  /* ---------- autofilter ---------- */
  function toggleFilter() {
    if (sh.filter) { clearFilterAll(); return; }
    const g = regionAround(rng());
    if (g.r2 <= g.r1 && g.c2 <= g.c1) { toast('Select a cell inside your data (with a header row) first.'); return; }
    structural(() => { sh.filter = { r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2, cols: {}, orig: {} }; });
  }
  function clearFilterAll() {
    structural(() => { const f = sh.filter; if (f) for (let r = f.r1 + 1; r <= f.r2; r++) if (sh.rowH.get(r) === 0) sh.rowH.delete(r); sh.filter = null; });
  }
  function applyFilter() {
    const f = sh.filter;
    if (!f) return;
    structural(() => {
      f.orig = f.orig || {};
      for (let r = f.r1 + 1; r <= f.r2; r++) {
        let show = true;
        for (const ck in f.cols) {
          const col = +ck;
          const set = new Set(f.cols[ck].vals);
          if (!set.has(display(getCell(r, col), r, col))) { show = false; break; }
        }
        if (!show) { if (!(r in f.orig)) f.orig[r] = sh.rowH.has(r) ? sh.rowH.get(r) : null; sh.rowH.set(r, 0); }
        else if (r in f.orig || sh.rowH.get(r) === 0) { const o = f.orig[r]; if (o === null || o === undefined) sh.rowH.delete(r); else sh.rowH.set(r, o); delete f.orig[r]; }
      }
    });
  }
  function drawFilterArrows(W, H) {
    const f = sh.filter;
    if (!f) return;
    ctx.save(); ctx.beginPath(); ctx.rect(HW, HH, W - HW, H - HH); ctx.clip();
    for (let c = f.c1; c <= f.c2; c++) {
      const R = cellRectView(f.r1, c, f.r1, c);
      if (R.x + R.w < HW || R.x > W || R.y + R.h < HH || R.y > H) continue;
      const bx = R.x + R.w - 17, by = R.y + R.h - 17;
      const on = !!f.cols[c];
      ctx.fillStyle = on ? '#d7ecdf' : '#fff'; ctx.fillRect(bx, by, 15, 15);
      ctx.strokeStyle = on ? '#217346' : '#9a9a9a'; ctx.lineWidth = 1; ctx.strokeRect(bx + .5, by + .5, 14, 14);
      ctx.fillStyle = on ? '#217346' : '#555';
      ctx.beginPath(); ctx.moveTo(bx + 4, by + 5.5); ctx.lineTo(bx + 11, by + 5.5); ctx.lineTo(bx + 7.5, by + 10.5); ctx.closePath(); ctx.fill();
    }
    ctx.restore();
  }
  function filterArrowHit(x, y) {
    const f = sh.filter;
    if (!f || x < HW || y < HH) return -1;
    for (let c = f.c1; c <= f.c2; c++) {
      const R = cellRectView(f.r1, c, f.r1, c);
      if (x >= R.x + R.w - 18 && x <= R.x + R.w && y >= R.y + R.h - 18 && y <= R.y + R.h) return c;
    }
    return -1;
  }
  const fMenu = document.createElement('div');
  fMenu.id = 'fmenu';
  $('#gridwrap').appendChild(fMenu);
  function openFilterMenu(c) {
    const f = sh.filter;
    const seen = new Map();
    for (let r = f.r1 + 1; r <= f.r2; r++) { const t = display(getCell(r, c), r, c); seen.set(t, (seen.get(t) || 0) + 1); }
    const vals = [...seen.keys()].sort((a, b) => { const na = Number(a), nb = Number(b); return (!isNaN(na) && !isNaN(nb) && a !== '' && b !== '') ? na - nb : a.toLowerCase().localeCompare(b.toLowerCase()); }).slice(0, 500);
    const cur = f.cols[c] ? new Set(f.cols[c].vals) : null;
    fMenu.innerHTML = '<div class="fm-act" data-a="az">Sort A to Z</div><div class="fm-act" data-a="za">Sort Z to A</div><div class="fm-act" data-a="cond">Filter by condition&hellip;</div><div class="fm-act" data-a="clear">Clear filter</div><hr><label class="fm-all"><input type="checkbox" checked> (Select all)</label><div class="fm-list"></div><div class="fm-btns"><button data-a="ok">OK</button><button data-a="cancel">Cancel</button></div>';
    const list = fMenu.querySelector('.fm-list');
    vals.forEach(v => { const l = document.createElement('label'); l.innerHTML = '<input type="checkbox"> <span></span>'; l.querySelector('input').checked = !cur || cur.has(v); l.querySelector('input').dataset.v = v; l.querySelector('span').textContent = v === '' ? '(Blanks)' : v; list.appendChild(l); });
    const all = fMenu.querySelector('.fm-all input');
    const sync = () => { all.checked = [...list.querySelectorAll('input')].every(i => i.checked); };
    sync();
    all.onchange = () => list.querySelectorAll('input').forEach(i => { i.checked = all.checked; });
    list.onchange = sync;
    const R = cellRectView(f.r1, c, f.r1, c);
    fMenu.style.left = Math.max(HW, Math.min(R.x + R.w - 200, scroller.clientWidth - 220)) + 'px'; fMenu.style.top = (R.y + R.h) + 'px'; fMenu.style.display = 'block';
    fMenu.onclick = e => {
      const a = e.target.dataset && e.target.dataset.a;
      if (!a) return;
      fMenu.style.display = 'none';
      if (a === 'cancel') return;
      if (a === 'az' || a === 'za') { select(f.r1, c, f.r2, f.c2, true); sortSel(a === 'az', { header: true }); return; }
      if (a === 'clear') { structural(() => { delete f.cols[c]; }); applyFilter(); return; }
      if (a === 'cond') {
        (async () => {
          const v = await form('Filter by condition', [{ id: 'op', label: 'Show rows where value', type: 'select', value: 'contains', options: [['contains', 'contains'], ['notcontains', 'does not contain'], ['begins', 'begins with'], ['ends', 'ends with'], ['=', 'equals'], ['<>', 'does not equal'], ['>', 'is greater than'], ['<', 'is less than'], ['>=', 'is greater than or equal to'], ['<=', 'is less than or equal to'], ['between', 'is between']] }, { id: 'a', label: 'Value', value: '' }, { id: 'b', label: 'And (for "between")', value: '' }], 'Filter');
          if (!v || v.a === '') return;
          const num = (s) => (s !== '' && !isNaN(Number(s))) ? Number(s) : null;
          const ok = (t) => {
            const x = t.toLowerCase(), y = v.a.toLowerCase();
            const nx = num(t), ny = num(v.a), nb = num(v.b);
            switch (v.op) {
              case 'contains': return x.includes(y); case 'notcontains': return !x.includes(y); case 'begins': return x.startsWith(y); case 'ends': return x.endsWith(y);
              case '=': return nx !== null && ny !== null ? nx === ny : x === y; case '<>': return nx !== null && ny !== null ? nx !== ny : x !== y;
              case 'between': return nx !== null && ny !== null && nb !== null && nx >= Math.min(ny, nb) && nx <= Math.max(ny, nb);
              default: if (nx === null || ny === null) return false; return v.op === '>' ? nx > ny : v.op === '<' ? nx < ny : v.op === '>=' ? nx >= ny : nx <= ny;
            }
          };
          const chosen = vals.filter(ok);
          structural(() => { f.cols[c] = { vals: chosen }; });
          applyFilter();
        })();
        return;
      }
      const chosen = [...list.querySelectorAll('input')].filter(i => i.checked).map(i => i.dataset.v);
      structural(() => { if (chosen.length === vals.length) delete f.cols[c]; else f.cols[c] = { vals: chosen }; });
      applyFilter();
    };
  }
  document.addEventListener('mousedown', e => { if (!fMenu.contains(e.target)) fMenu.style.display = 'none'; if (!dvMenu.contains(e.target) && e.target !== dvBtn) dvMenu.style.display = 'none'; });

  /* ---------- named ranges ---------- */
  async function nameDialog() {
    wb.names = wb.names || {};
    const existing = Object.keys(wb.names);
    const g = rng();
    const def = Formula.quoteSheet(sh.name) + '!$' + colName(g.c1) + '$' + (g.r1 + 1) + ((g.r1 === g.r2 && g.c1 === g.c2) ? '' : ':$' + colName(g.c2) + '$' + (g.r2 + 1));
    const v = await form('Defined names', [
      { id: 'existing', label: 'Existing name', type: 'select', value: '', options: [['', '(new name)'], ...existing.map(n => [n, n + '  =  ' + wb.names[n]])] },
      { id: 'name', label: 'Name', value: '' },
      { id: 'ref', label: 'Refers to', value: def },
      { id: 'del', label: 'Delete the existing name', type: 'checkbox', value: false },
    ], 'OK');
    if (!v) return;
    if (v.existing && v.del) { structural(() => { delete wb.names[v.existing]; }); return; }
    const name = (v.name || v.existing || '').trim();
    if (!name) return;
    if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) || /^[A-Za-z]{1,3}\d+$/i.test(name) || /^(true|false)$/i.test(name)) { toast('That is not a valid name (use letters, digits, _ and ., and do not look like a cell address).', 5000); return; }
    structural(() => { wb.names[name.toUpperCase()] = v.ref.replace(/^=/, ''); wb.names[name.toUpperCase()] = wb.names[name.toUpperCase()]; });
    toast('Defined name ' + name);
  }

  /* ---------- cell comments ---------- */
  const cmTip = document.createElement('div');
  cmTip.id = 'cmtip';
  $('#gridwrap').appendChild(cmTip);
  const cmAt = (r, c) => (sh.comments || []).find(x => x.r === r && x.c === c);
  async function commentDialog() {
    const r = S().r, c = S().c;
    const ex = cmAt(r, c);
    const v = await form(ex ? 'Edit comment' : 'New comment', [{ id: 'text', label: 'Comment for ' + cellName(r, c), value: ex ? ex.text : '' }], 'OK');
    if (!v) return;
    structural(() => {
      sh.comments = sh.comments.filter(x => !(x.r === r && x.c === c));
      if (v.text.trim()) sh.comments.push({ r, c, text: v.text, author: 'Author' });
    });
  }
  function deleteComment() { const r = S().r, c = S().c; if (!cmAt(r, c)) return; structural(() => { sh.comments = sh.comments.filter(x => !(x.r === r && x.c === c)); }); }
  function drawCommentMarks(W, H) {
    if (!sh.comments || !sh.comments.length) return;
    ctx.save(); ctx.beginPath(); ctx.rect(HW, HH, W - HW, H - HH); ctx.clip();
    ctx.fillStyle = '#e03a3a';
    for (const m of sh.comments) {
      const R = cellRectView(m.r, m.c, m.r, m.c);
      if (R.x + R.w < HW || R.x > W || R.y + R.h < HH || R.y > H) continue;
      ctx.beginPath(); ctx.moveTo(R.x + R.w - 7, R.y + 1); ctx.lineTo(R.x + R.w - 1, R.y + 1); ctx.lineTo(R.x + R.w - 1, R.y + 7); ctx.closePath(); ctx.fill();
    }
    ctx.restore();
  }
  scroller.addEventListener('mousemove', e => {
    if (drag || !sh.comments || !sh.comments.length) { cmTip.style.display = 'none'; return; }
    const h = hit(e.clientX, e.clientY);
    const m = h.zone === 'cell' && h.r >= 0 ? cmAt(h.r, h.c) : null;
    if (!m) { cmTip.style.display = 'none'; return; }
    cmTip.textContent = (m.author ? m.author + ':\n' : '') + m.text;
    cmTip.style.left = Math.min(h.x + 14, scroller.clientWidth - 230) + 'px'; cmTip.style.top = (h.y + 14) + 'px'; cmTip.style.display = 'block';
  });
  $('#dv-edit').onclick = validationDialog; $('#dv-clear').onclick = clearValidation;
  $('#flt-toggle').onclick = toggleFilter; $('#flt-reapply').onclick = applyFilter;
  $('#nm-edit').onclick = nameDialog;
  $('#cm-edit').onclick = commentDialog; $('#cm-del').onclick = deleteComment;

  /* ---------- conditional formatting ---------- */
  const STYLE_OPTS = [['red', 'Light red fill, dark red text'], ['yellow', 'Yellow fill, dark yellow text'], ['green', 'Green fill, dark green text'], ['redfill', 'Light red fill'], ['redtext', 'Red text'], ['bold', 'Bold text']];
  async function addCf(kind) {
    if (editing) commitEdit(0, 0);
    const g = clampRows(rngExpanded());
    const base = { id: 'f' + Math.random().toString(36).slice(2, 8), r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 };
    const styleField = { id: 'style', label: 'Format with', type: 'select', value: 'red', options: STYLE_OPTS };
    let rule = null;
    if (['>', '<', '=', 'between'].includes(kind)) {
      const v = await form('Highlight cells ' + (kind === 'between' ? 'between' : kind === '=' ? 'equal to' : kind === '>' ? 'greater than' : 'less than'), [{ id: 'v1', label: kind === 'between' ? 'Minimum' : 'Value', value: '' }, ...(kind === 'between' ? [{ id: 'v2', label: 'Maximum', value: '' }] : []), styleField], 'OK');
      if (!v || v.v1 === '') return;
      rule = { ...base, type: 'cell', op: kind, v1: v.v1, v2: v.v2, style: CondFmt.PRESETS[v.style] };
    } else if (kind === 'text') {
      const v = await form('Text that contains', [{ id: 'text', label: 'Text', value: '' }, { id: 'op', label: 'Rule', type: 'select', value: 'contains', options: [['contains', 'contains'], ['begins', 'begins with'], ['ends', 'ends with'], ['notcontains', 'does not contain']] }, styleField], 'OK');
      if (!v || !v.text) return;
      rule = { ...base, type: 'text', op: v.op, text: v.text, style: CondFmt.PRESETS[v.style] };
    } else if (kind === 'dup') {
      const v = await form('Duplicate values', [styleField], 'OK');
      if (!v) return;
      rule = { ...base, type: 'dup', style: CondFmt.PRESETS[v.style] };
    } else if (kind === 'top' || kind === 'bottom') {
      const v = await form(kind === 'top' ? 'Top items' : 'Bottom items', [{ id: 'n', label: 'Number of items', type: 'number', value: 10, min: 1, max: 1000 }, styleField], 'OK');
      if (!v) return;
      rule = { ...base, type: 'top', bottom: kind === 'bottom', n: Math.max(1, v.n | 0), style: CondFmt.PRESETS[v.style] };
    } else if (kind.startsWith('scale:')) {
      const sc = CondFmt.SCALES[kind.slice(6)];
      rule = { ...base, type: 'scale', cfvo: sc.cfvo.map(x => ({ ...x })), colors: sc.colors.slice() };
    } else if (kind === 'bar') {
      rule = { ...base, type: 'bar', color: '638EC6' };
    } else if (kind === 'expr') {
      const v = await form('Custom formula', [{ type: 'note', label: 'Write a formula for the top-left cell of the selection, for example =$B2>100. It applies to every cell in the selection.' }, { id: 'f', label: 'Formula', value: '=' }, styleField], 'OK');
      if (!v || v.f.trim().length < 2) return;
      const f = v.f.trim().replace(/^=/, '');
      try { Formula.parse(f); } catch (e) { toast('That formula has a syntax error.'); return; }
      rule = { ...base, type: 'expr', f, style: CondFmt.PRESETS[v.style] };
    } else if (kind === 'clear-sel') {
      const g2 = rngExpanded();
      structural(() => { sh.cf = sh.cf.filter(r => !(r.r1 >= g2.r1 && r.r2 <= g2.r2 && r.c1 >= g2.c1 && r.c2 <= g2.c2)); });
      return;
    } else if (kind === 'clear-all') { structural(() => { sh.cf = []; }); return; }
    else if (kind === 'manage') {
      if (!sh.cf.length) { toast('This sheet has no conditional formatting rules.'); return; }
      const v = await form('Manage rules', [{ id: 'i', label: 'Rule', type: 'select', value: '0', options: sh.cf.map((r, i) => [String(i), CondFmt.describe(r)]) }, { id: 'del', label: 'Delete this rule', type: 'checkbox', value: false }], 'Apply');
      if (v && v.del) structural(() => { sh.cf.splice(+v.i, 1); });
      return;
    }
    if (!rule) return;
    structural(() => { rule.p = sh.cf.length + 1; sh.cf.push(rule); });
  }
  $('#cf-add').onchange = e => { const v = e.target.value; e.target.value = ''; if (v) addCf(v); };
  $('#i-chart').onchange = e => { const v = e.target.value; e.target.value = ''; if (v) makeChart(v); };
  $('#c-del').onclick = deleteChart;
  $('#s-add').onclick = addSheet; $('#addsheet').onclick = addSheet; $('#s-dup').onclick = dupSheet; $('#s-ren').onclick = renameSheet; $('#s-del').onclick = deleteSheet;
  $('#sort-az').onclick = () => sortSel(true); $('#sort-za').onclick = () => sortSel(false);
  const fsel = $('#i-func');
  Formula.functions().forEach(n => { const o = document.createElement('option'); o.textContent = n; fsel.appendChild(o); });
  fsel.onchange = () => { if (fsel.value) { startEdit(editing ? editor.value + fsel.value + '(' : '=' + fsel.value + '('); } fsel.value = ''; };
  $('#v-freeze').onclick = () => structural(() => { sh.freeze = { r: S().r, c: S().c }; });
  $('#v-freeze-row').onclick = () => structural(() => { sh.freeze = { r: 1, c: sh.freeze.c }; });
  $('#v-freeze-col').onclick = () => structural(() => { sh.freeze = { r: sh.freeze.r, c: 1 }; });
  $('#v-unfreeze').onclick = () => structural(() => { sh.freeze = { r: 0, c: 0 }; });
  $('#v-grid').onchange = e => { sh.showGrid = e.target.checked; markDirty(); render(); };

  /* ================= tables ================= */
  const ACCENTS = [['000000', 'D9D9D9'], ['4472C4', 'D9E1F2'], ['ED7D31', 'FCE4D6'], ['A5A5A5', 'EDEDED'], ['FFC000', 'FFF2CC'], ['5B9BD5', 'DDEBF7'], ['70AD47', 'E2EFDA']];
  function tableColors(style) {
    const m = /(\d+)$/.exec(style || '');
    const a = ACCENTS[m ? (parseInt(m[1], 10) - 1) % 7 : 1];
    return { h: a[0], b: a[1] };
  }
  function uniqueTableName() {
    const used = new Set();
    for (const s of wb.sheets) for (const t of (s.tables || [])) used.add(t.name.toLowerCase());
    let i = 1; while (used.has('table' + i)) i++;
    used.add('table' + i);
    return 'Table' + i;
  }
  const tableAt = (r, c) => (sh.tables || []).find(t => r >= t.r1 && r <= t.r2 && c >= t.c1 && c <= t.c2);
  const hits = (a, b) => !(a.r2 < b.r1 || a.r1 > b.r2 || a.c2 < b.c1 || a.c1 > b.c2);
  function makeTable(style) {
    if (editing) commitEdit(0, 0);
    const here = tableAt(S().r, S().c);
    if (here) { structural(() => { here.style = style; }); return; }
    const g = clampRows(regionAround(rng()));
    if (g.r2 <= g.r1) { toast('A table needs a header row and at least one row of data. Select your data first.'); return; }
    if ((sh.tables || []).some(t => hits(t, g))) { toast('That range overlaps another table.'); return; }
    if (sh.merges.some(m => hits(m, g))) { toast('Unmerge the cells in that range before making it a table.'); return; }
    structural(() => {
      // header cells must be unique, non-empty text
      const seen = new Set();
      for (let c = g.c1; c <= g.c2; c++) {
        const cell = getCell(g.r1, c);
        let n = (cell ? display(cell, g.r1, c) : '').trim() || 'Column' + (c - g.c1 + 1);
        let u = n, i = 2; while (seen.has(u.toLowerCase())) u = n + i++;
        seen.add(u.toLowerCase());
        if (!cell || cell.f != null || cell.v !== u) { const o = { v: u }; if (cell && cell.s) o.s = cell.s; sh.cells.set(g.r1 * MAXC + c, o); }
      }
      if (!sh.tables) sh.tables = [];
      sh.tables.push({ name: uniqueTableName(), r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2, style });
      if (!sh.filter) sh.filter = { r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2, cols: {}, orig: {} };
    });
    toast('Table created. Use Table > Convert to range to remove it.');
  }
  function tableToRange() {
    const t = tableAt(S().r, S().c);
    if (!t) { toast('Select a cell inside a table first.'); return; }
    structural(() => {
      sh.tables = sh.tables.filter(x => x !== t);
      if (sh.filter && hits(sh.filter, t) && !Object.keys(sh.filter.cols || {}).length) sh.filter = null;
    });
  }
  function tableTotals() {
    const t = tableAt(S().r, S().c);
    if (!t) { toast('Select a cell inside a table first.'); return; }
    if (t.totals) {
      structural(() => { for (let c = t.c1; c <= t.c2; c++) sh.cells.delete(t.r2 * MAXC + c); t.totals = false; delete t.tcols; t.r2--; });
      return;
    }
    const r = t.r2 + 1;
    for (let c = t.c1; c <= t.c2; c++) if (getCell(r, c) && display(getCell(r, c), r, c) !== '') { toast('The row below the table must be empty to add a total row.'); return; }
    structural(() => {
      const d1 = t.r1 + (t.hdr !== false ? 1 : 0);
      t.tcols = [];
      for (let c = t.c1; c <= t.c2; c++) {
        let numeric = false;
        for (let rr = d1; rr <= t.r2; rr++) if (typeof value(rr, c) === 'number') { numeric = true; break; }
        if (c === t.c1 && !numeric) { sh.cells.set(r * MAXC + c, { v: 'Total' }); t.tcols.push({ label: 'Total', fn: '' }); }
        else if (numeric) { sh.cells.set(r * MAXC + c, { f: `SUBTOTAL(109,${cellName(d1, c)}:${cellName(t.r2, c)})` }); t.tcols.push({ label: '', fn: 'sum' }); }
        else t.tcols.push({ label: '', fn: '' });
      }
      t.totals = true; t.r2 = r;
    });
  }
  $('#tb-add').onchange = e => {
    const v = e.target.value; e.target.value = '';
    if (!v) return;
    if (v === 'range') tableToRange(); else if (v === 'totals') tableTotals(); else makeTable(v);
  };

  /* ================= sparklines ================= */
  function sparkRange(sp) {
    let ref = sp.ref, s = sh;
    const i = ref.lastIndexOf('!');
    if (i >= 0) { s = wb.sheetByName(ref.slice(0, i).replace(/^'|'$/g, '').replace(/''/g, "'")); ref = ref.slice(i + 1); }
    const g = Formula.parseRange(ref);
    return s && g ? { s, g } : null;
  }
  function drawSpark(sp, x, y, w, h) {
    const q = sparkRange(sp);
    if (!q) return;
    const vals = [];
    for (let r = q.g.r1; r <= q.g.r2; r++) for (let c = q.g.c1; c <= q.g.c2; c++) { const v = Formula.cellValue(wb, q.s, r, c); vals.push(typeof v === 'number' ? v : null); }
    const nums = vals.filter(v => v !== null);
    const px = 3, py = 3, iw = w - 2 * px, ih = h - 2 * py;
    if (!nums.length || iw < 4 || ih < 4) return;
    let mn = Math.min(...nums), mx = Math.max(...nums);
    ctx.save(); ctx.beginPath(); ctx.rect(x, y, w, h); ctx.clip();
    const col = hex(sp.color || '376092');
    if (sp.type === 'column' || sp.type === 'stacked') {
      const lo = Math.min(0, mn), hi = Math.max(0, mx), span = (hi - lo) || 1, bw = iw / vals.length;
      vals.forEach((v, i) => {
        if (v === null) return;
        let top, bot;
        if (sp.type === 'stacked') { if (v === 0) return; top = v > 0 ? 0 : 0.5; bot = top + 0.5; }
        else { top = (hi - Math.max(v, 0)) / span; bot = (hi - Math.min(v, 0)) / span; }
        ctx.fillStyle = v < 0 ? '#d00000' : col;
        ctx.fillRect(x + px + bw * i + bw * .12, y + py + ih * top, bw * .76, Math.max(1, ih * (bot - top)));
      });
    } else {
      if (mx === mn) { mx += 1; mn -= 1; }
      ctx.strokeStyle = col; ctx.lineWidth = 1.5; ctx.lineJoin = 'round'; ctx.beginPath();
      let pen = false;
      vals.forEach((v, i) => {
        if (v === null) { pen = false; return; }
        const X = x + px + (vals.length === 1 ? iw / 2 : iw * i / (vals.length - 1)), Y = y + py + ih * (1 - (v - mn) / (mx - mn));
        if (pen) ctx.lineTo(X, Y); else ctx.moveTo(X, Y);
        pen = true;
      });
      ctx.stroke();
    }
    ctx.restore();
  }
  async function addSpark(type) {
    if (editing) commitEdit(0, 0);
    const g = rng();
    if (type === 'clear') { structural(() => { sh.sparks = (sh.sparks || []).filter(s => !(s.r >= g.r1 && s.r <= g.r2 && s.c >= g.c1 && s.c <= g.c2)); }); return; }
    const multi = g.r1 !== g.r2 || g.c1 !== g.c2;
    const v = await form('Insert sparklines', [
      { type: 'note', label: 'A sparkline is a tiny chart inside one cell. Give the data range, and the cell (or column of cells) where the sparklines go. One sparkline is made per row of data.' },
      { id: 'data', label: 'Data range', value: multi ? cellName(g.r1, g.c1) + ':' + cellName(g.r2, g.c2) : '' },
      { id: 'at', label: 'Place at (first cell)', value: multi ? cellName(g.r1, g.c2 + 1) : cellName(g.r1, g.c1) },
      { id: 'color', label: 'Colour', type: 'select', value: '376092', options: [['376092', 'Blue'], ['C0504D', 'Red'], ['4F9D3A', 'Green'], ['ED7D31', 'Orange'], ['7030A0', 'Purple'], ['404040', 'Dark grey']] },
    ], 'Insert');
    if (!v) return;
    const d = Formula.parseRange(String(v.data).trim()), at = Formula.parseCell(String(v.at).trim().replace(/\$/g, ''));
    if (!d || !at) { toast('Enter a data range like B2:M2 and a cell like N2.'); return; }
    structural(() => {
      if (!sh.sparks) sh.sparks = [];
      const byCol = d.c1 === d.c2 && d.r2 > d.r1;   // a single column of data makes one sparkline
      const n = byCol ? 1 : d.r2 - d.r1 + 1;
      for (let i = 0; i < n; i++) {
        const r = at.r + i, c = at.c;
        sh.sparks = sh.sparks.filter(s => !(s.r === r && s.c === c));
        sh.sparks.push({ r, c, ref: byCol ? cellName(d.r1, d.c1) + ':' + cellName(d.r2, d.c2) : cellName(d.r1 + i, d.c1) + ':' + cellName(d.r1 + i, d.c2), type, color: v.color });
      }
    });
  }
  $('#sp-add').onchange = e => { const v = e.target.value; e.target.value = ''; if (v) addSpark(v); };

  /* ================= pivot summary ================= */
  const AGG = {
    sum: a => a.reduce((x, y) => x + y, 0), count: a => a.length,
    avg: a => a.length ? a.reduce((x, y) => x + y, 0) / a.length : 0,
    max: a => a.length ? Math.max(...a) : 0, min: a => a.length ? Math.min(...a) : 0,
  };
  function buildPivot(g, rowF, colF, valF, agg) {
    const src = sh;
    const keyOf = (r, c) => { const cell = src.get(r, c); return cell ? display(cell, r, c) : ''; };
    const rowKeys = [], colKeys = [], bucket = new Map();
    const seenR = new Set(), seenC = new Set();
    for (let r = g.r1 + 1; r <= g.r2; r++) {
      const rk = keyOf(r, rowF), ck = colF < 0 ? '' : keyOf(r, colF);
      if (!seenR.has(rk)) { seenR.add(rk); rowKeys.push(rk); }
      if (!seenC.has(ck)) { seenC.add(ck); colKeys.push(ck); }
      const v = Formula.cellValue(wb, src, r, valF);
      const id = rk + '\u0001' + ck;
      if (!bucket.has(id)) bucket.set(id, []);
      if (agg === 'count') { if (v !== null && v !== undefined && v !== '') bucket.get(id).push(1); }
      else if (typeof v === 'number') bucket.get(id).push(v);
    }
    const cmp = (a, b) => { const x = Number(a), y = Number(b); return (a !== '' && b !== '' && !isNaN(x) && !isNaN(y)) ? x - y : String(a).localeCompare(String(b)); };
    rowKeys.sort(cmp); colKeys.sort(cmp);
    const fn = AGG[agg];
    const get = (rks, cks) => { let all = []; for (const rk of rks) for (const ck of cks) { const b = bucket.get(rk + '\u0001' + ck); if (b) all = all.concat(b); } return all; };
    const out = [];
    const label = { sum: 'Sum', count: 'Count', avg: 'Average', max: 'Max', min: 'Min' }[agg] + ' of ' + keyOf(g.r1, valF);
    if (colF < 0) {
      out.push([keyOf(g.r1, rowF), label]);
      for (const rk of rowKeys) out.push([rk === '' ? '(blank)' : rk, fn(get([rk], colKeys))]);
      out.push(['Grand Total', fn(get(rowKeys, colKeys))]);
    } else {
      out.push([label, ...colKeys.map(k => k === '' ? '(blank)' : k), 'Grand Total']);
      for (const rk of rowKeys) out.push([rk === '' ? '(blank)' : rk, ...colKeys.map(ck => { const a = get([rk], [ck]); return a.length || agg === 'count' ? fn(a) : ''; }), fn(get([rk], colKeys))]);
      out.push(['Grand Total', ...colKeys.map(ck => fn(get(rowKeys, [ck]))), fn(get(rowKeys, colKeys))]);
    }
    return out;
  }
  // writes a pivot result into sheet n (clearing what the previous refresh left there)
  function fillPivot(n, out) {
    const bold = wb.styleId({ b: true, bg: 'D9E1F2' }), boldTot = wb.styleId({ b: true, bt: 'thin:4472C4' });
    if (n.pivot && n.pivot.area) { const a = n.pivot.area; for (let r = 0; r <= a.r; r++) for (let c = 0; c <= a.c; c++) n.cells.delete(r * MAXC + c); }
    out.forEach((row, r) => row.forEach((v, c) => {
      if (v === '') return;
      const cell = {};
      if (typeof v === 'string') { const p = c === 0 || r === 0 ? { v } : NumFmt.parseInput(v); cell.v = p.v; } else cell.v = Math.round(v * 1e10) / 1e10;
      if (r === 0) cell.s = bold; else if (r === out.length - 1) cell.s = boldTot;
      n.cells.set(r * MAXC + c, cell);
    }));
    if (n.pivot) n.pivot.area = { r: out.length - 1, c: Math.max(...out.map(r => r.length)) - 1 };
  }
  function pivotFromDef(def) {
    const src = wb.sheetByName(def.src);
    if (!src) return null;
    const keep = sh;
    sh = src;
    try {
      // the source may have grown: extend the range down to the last filled row of the first column
      const g = Object.assign({}, def.g);
      while (src.get(g.r2 + 1, g.c1) || src.get(g.r2 + 1, def.row)) g.r2++;
      def.g = g;
      return buildPivot(g, def.row, def.col, def.val, def.agg);
    } finally { sh = keep; }
  }
  function refreshPivot(n, quiet) {
    n = n || sh;
    if (!n.pivot) { if (!quiet) toast('This sheet is not a PivotTable. Click a PivotTable sheet first.'); return false; }
    const out = pivotFromDef(n.pivot);
    if (!out) { if (!quiet) toast('The source sheet "' + n.pivot.src + '" no longer exists.'); return false; }
    if (quiet) { fillPivot(n, out); wb.touch(); return true; }
    structural(() => fillPivot(n, out));
    toast('PivotTable refreshed');
    return true;
  }
  function refreshAllPivots() { let k = 0; structural(() => { for (const s of wb.sheets) if (s.pivot) { const o = pivotFromDef(s.pivot); if (o) { fillPivot(s, o); k++; } } }); toast(k ? 'Refreshed ' + k + ' PivotTable' + (k > 1 ? 's' : '') : 'No PivotTables in this workbook'); }
  function placePivot(out, srcStyleOfVal, def) {
    structural(() => {
      const n = new Sheet(uniqueName('Pivot'));
      if (def) { n.pivot = def; fillPivot(n, out); n.colW.set(0, 140); n.sel = { r: 0, c: 0, er: 0, ec: 0 }; wb.sheets.splice(wb.sheets.indexOf(sh) + 1, 0, n); wb.active = wb.sheets.indexOf(n); sh = n; return; }
      const bold = wb.styleId({ b: true, bg: 'D9E1F2' }), boldTot = wb.styleId({ b: true, bt: 'thin:4472C4' });
      out.forEach((row, r) => row.forEach((v, c) => {
        if (v === '') return;
        const cell = {};
        if (typeof v === 'string') { const p = c === 0 || r === 0 ? { v } : NumFmt.parseInput(v); cell.v = p.v; } else cell.v = Math.round(v * 1e10) / 1e10;
        if (r === 0) cell.s = bold; else if (r === out.length - 1) cell.s = boldTot;
        n.cells.set(r * MAXC + c, cell);
      }));
      n.colW.set(0, 140);
      n.sel = { r: 0, c: 0, er: 0, ec: 0 };
      wb.sheets.splice(wb.sheets.indexOf(sh) + 1, 0, n); wb.active = wb.sheets.indexOf(n); sh = n;
    });
  }
  async function pivotDialog() {
    if (editing) commitEdit(0, 0);
    const t = tableAt(S().r, S().c);
    const g = t ? { r1: t.r1, c1: t.c1, r2: t.totals ? t.r2 - 1 : t.r2, c2: t.c2 } : clampRows(regionAround(rng()));
    if (g.r2 <= g.r1 || g.c2 <= g.c1) { toast('Select a cell inside data that has a header row and at least two columns.'); return; }
    const heads = [];
    for (let c = g.c1; c <= g.c2; c++) { const cell = getCell(g.r1, c); heads.push([String(c), (cell ? display(cell, g.r1, c) : '') || 'Column ' + colName(c)]); }
    let numCol = g.c2;
    for (let c = g.c1; c <= g.c2; c++) if (typeof value(g.r1 + 1, c) === 'number') numCol = c;
    const v = await form('PivotTable summary', [
      { type: 'note', label: 'Summarises ' + cellName(g.r1, g.c1) + ':' + cellName(g.r2, g.c2) + ' on a new sheet. It updates when you open that sheet or press Refresh (Data tab, or Alt+F5). New rows added under the data are picked up.' },
      { id: 'row', label: 'Rows', type: 'select', value: String(g.c1), options: heads },
      { id: 'col', label: 'Columns', type: 'select', value: '-1', options: [['-1', '(none)'], ...heads] },
      { id: 'val', label: 'Values', type: 'select', value: String(numCol), options: heads },
      { id: 'agg', label: 'Summarise by', type: 'select', value: 'sum', options: [['sum', 'Sum'], ['count', 'Count'], ['avg', 'Average'], ['max', 'Max'], ['min', 'Min']] },
    ], 'Create');
    if (!v) return;
    placePivot(buildPivot(g, +v.row, +v.col, +v.val, v.agg), null, { src: sh.name, g: Object.assign({}, g), row: +v.row, col: +v.col, val: +v.val, agg: v.agg });
  }
  $('#pv-make').onclick = pivotDialog;
  $('#pv-refresh').onclick = () => refreshPivot();
  $('#pv-refresh-all').onclick = refreshAllPivots;
  window.addEventListener('keydown', e => { if (e.altKey && e.key === 'F5') { e.preventDefault(); refreshPivot(); } else if (e.ctrlKey && e.altKey && e.key === 'F5') { e.preventDefault(); refreshAllPivots(); } });

  /* ================= goal seek ================= */
  function goalSeek(target, goal, changing) {
    // returns { ok, x, got }; leaves the sheet untouched
    const k = changing.r * MAXC + changing.c;
    const orig = cloneCell(sh.cells.get(k));
    const sid = orig && orig.s;
    const f = (x) => { const o = { v: x }; if (sid) o.s = sid; sh.cells.set(k, o); wb.touch(); const v = Formula.cellValue(wb, sh, target.r, target.c); return typeof v === 'number' ? v - goal : NaN; };
    let x0 = orig && typeof orig.v === 'number' ? orig.v : 0;
    let best = { x: x0, e: Infinity };
    const note = (x, e) => { if (!isNaN(e) && Math.abs(e) < Math.abs(best.e)) best = { x, e }; };
    try {
      let f0 = f(x0); note(x0, f0);
      let x1 = x0 === 0 ? 1 : x0 * 1.1 + 0.01, f1 = f(x1); note(x1, f1);
      // secant
      for (let i = 0; i < 80 && Math.abs(best.e) > 1e-9; i++) {
        if (isNaN(f0) || isNaN(f1) || f1 === f0) break;
        let x2 = x1 - f1 * (x1 - x0) / (f1 - f0);
        if (!isFinite(x2)) break;
        x0 = x1; f0 = f1; x1 = x2; f1 = f(x1); note(x1, f1);
      }
      // bracket + bisection fallback
      if (Math.abs(best.e) > 1e-7) {
        let lo = null, hi = null, flo, fhi;
        const c0 = best.x;
        for (let step = 1; step < 1e12 && lo === null; step *= 2) {
          const a = c0 - step, b = c0 + step, fa = f(a), fb = f(b); note(a, fa); note(b, fb);
          const fc = f(c0);
          if (!isNaN(fa) && !isNaN(fc) && fa * fc <= 0) { lo = a; hi = c0; flo = fa; fhi = fc; }
          else if (!isNaN(fb) && !isNaN(fc) && fb * fc <= 0) { lo = c0; hi = b; flo = fc; fhi = fb; }
          else if (!isNaN(fa) && !isNaN(fb) && fa * fb <= 0) { lo = a; hi = b; flo = fa; fhi = fb; }
        }
        if (lo !== null) for (let i = 0; i < 200 && Math.abs(best.e) > 1e-10; i++) {
          const m = (lo + hi) / 2, fm = f(m); note(m, fm);
          if (isNaN(fm)) break;
          if (flo * fm <= 0) { hi = m; fhi = fm; } else { lo = m; flo = fm; }
        }
      }
    } finally {
      if (orig) sh.cells.set(k, orig); else sh.cells.delete(k);
      wb.touch();
    }
    const scale = Math.max(1, Math.abs(goal));
    return { ok: Math.abs(best.e) <= 1e-6 * scale, x: Math.abs(best.x - Math.round(best.x * 1e9) / 1e9) < 1e-12 ? best.x : +best.x.toPrecision(12), got: goal + best.e };
  }
  async function goalSeekDialog() {
    if (editing) commitEdit(0, 0);
    const v = await form('Goal Seek', [
      { type: 'note', label: 'Finds the input value that makes a formula reach the result you want.' },
      { id: 'set', label: 'Set cell (a formula)', value: cellName(S().r, S().c) },
      { id: 'to', label: 'To value', type: 'number', value: 0 },
      { id: 'by', label: 'By changing cell', value: '' },
    ], 'Solve');
    if (!v) return;
    const target = Formula.parseCell(String(v.set).trim().replace(/\$/g, '')), changing = Formula.parseCell(String(v.by).trim().replace(/\$/g, ''));
    const goal = Number(v.to);
    if (!target || !changing || isNaN(goal)) { toast('Enter two cell references like B5 and B2, and a number.'); return; }
    const tc = getCell(target.r, target.c);
    if (!tc || tc.f == null) { toast('The "Set cell" must contain a formula.'); return; }
    const cc = getCell(changing.r, changing.c);
    if (cc && cc.f != null) { toast('The "By changing cell" must hold a plain number, not a formula.'); return; }
    const res = goalSeek(target, goal, changing);
    if (!res.ok) { toast('Goal Seek could not find a solution. Closest result: ' + (+res.got.toPrecision(8)) + ' with ' + cellName(changing.r, changing.c) + ' = ' + res.x, 7000); return; }
    applyEdit(rec => { const o = { v: res.x }; if (cc && cc.s) o.s = cc.s; rec(sh, changing.r, changing.c, o); });
    toast('Goal Seek found a solution: ' + cellName(changing.r, changing.c) + ' = ' + res.x);
  }
  $('#gs-run').onclick = goalSeekDialog;

  /* ================= Solver ================= */
  // Nelder-Mead with penalties for the constraints, several restarts, then integer rounding + a local polish.
  function parseCons(text) {
    const out = [];
    for (const raw of String(text || '').replace(/\r/g, '').split(/\n|;/)) {
      const line = raw.trim().replace(/\$/g, '');
      if (!line) continue;
      let m = /^([A-Za-z]{1,3}\d+(?::[A-Za-z]{1,3}\d+)?)\s*(<=|>=|=|\bint\b|\bbin\b)\s*(.*)$/i.exec(line);
      if (!m) return { error: 'Cannot read constraint "' + line + '". Use forms like  B2 <= 10,  B2:B5 >= 0,  B3 = C3,  B2 int' };
      const g = Formula.parseRange(m[1]);
      const op = m[2].toLowerCase();
      let rhs = m[3].trim();
      if (op === 'int' || op === 'bin') { out.push({ g, op }); continue; }
      if (rhs === '') return { error: 'Constraint "' + line + '" needs a value on the right.' };
      const num = Number(rhs);
      out.push({ g, op, rhs: isNaN(num) ? (Formula.parseCell(rhs) || null) : num, rhsTxt: rhs });
      if (isNaN(num) && !Formula.parseCell(rhs)) return { error: 'The right side of "' + line + '" must be a number or a single cell.' };
    }
    return { list: out };
  }
  function solve(objCell, mode, target, vars, cons, opts) {
    opts = opts || {};
    const keys = vars.map(v => v.r * MAXC + v.c);
    const orig = keys.map(k => cloneCell(sh.cells.get(k)));
    const set = (x) => { keys.forEach((k, i) => { const o = { v: x[i] }; if (orig[i] && orig[i].s) o.s = orig[i].s; sh.cells.set(k, o); }); wb.touch(); };
    const num = (r, c) => { const v = Formula.cellValue(wb, sh, r, c); return typeof v === 'number' ? v : (typeof v === 'boolean' ? +v : NaN); };
    const ints = new Set(), bins = new Set();
    cons.forEach(c => { if (c.op === 'int' || c.op === 'bin') for (let r = c.g.r1; r <= c.g.r2; r++) for (let cc = c.g.c1; cc <= c.g.c2; cc++) { const i = keys.indexOf(r * MAXC + cc); if (i >= 0) (c.op === 'int' ? ints : bins).add(i); } });
    const violation = () => {
      let v = 0;
      for (const c of cons) {
        if (c.op === 'int' || c.op === 'bin') continue;
        for (let r = c.g.r1; r <= c.g.r2; r++) for (let cc = c.g.c1; cc <= c.g.c2; cc++) {
          const a = num(r, cc), b = typeof c.rhs === 'number' ? c.rhs : num(c.rhs.r, c.rhs.c);
          if (isNaN(a) || isNaN(b)) { v += 1e6; continue; }
          if (c.op === '<=') v += Math.max(0, a - b); else if (c.op === '>=') v += Math.max(0, b - a); else v += Math.abs(a - b);
        }
      }
      bins.forEach(i => { /* handled by rounding */ });
      return v;
    };
    const objective = () => { const o = num(objCell.r, objCell.c); if (isNaN(o)) return Infinity; return mode === 'max' ? -o : mode === 'min' ? o : Math.abs(o - target); };
    let mu = 1e3;
    const F = (x) => { set(x); const o = objective(); const vi = violation(); return o + mu * vi * (1 + Math.abs(o)) + mu * vi; };
    const nm = (x0, scale, iters) => {
      const n = x0.length, a = 1, g = 2, rho = 0.5, sg = 0.5;
      let P = [x0.slice()];
      for (let i = 0; i < n; i++) { const p = x0.slice(); p[i] += scale[i]; P.push(p); }
      let Fv = P.map(F);
      for (let it = 0; it < iters; it++) {
        const ord = Fv.map((f, i) => i).sort((i, j) => Fv[i] - Fv[j]);
        P = ord.map(i => P[i]); Fv = ord.map(i => Fv[i]);
        if (Math.abs(Fv[n] - Fv[0]) < 1e-12 * (1 + Math.abs(Fv[0])) && it > 50) break;
        const c = new Array(n).fill(0);
        for (let i = 0; i < n; i++) for (let j = 0; j < n; j++) c[j] += P[i][j] / n;
        const xr = c.map((v, j) => v + a * (v - P[n][j])), fr = F(xr);
        if (fr < Fv[0]) { const xe = c.map((v, j) => v + g * (xr[j] - v)), fe = F(xe); if (fe < fr) { P[n] = xe; Fv[n] = fe; } else { P[n] = xr; Fv[n] = fr; } }
        else if (fr < Fv[n - 1]) { P[n] = xr; Fv[n] = fr; }
        else {
          const xc = c.map((v, j) => v + rho * (P[n][j] - v)), fc = F(xc);
          if (fc < Fv[n]) { P[n] = xc; Fv[n] = fc; }
          else for (let i = 1; i <= n; i++) { P[i] = P[0].map((v, j) => v + sg * (P[i][j] - v)); Fv[i] = F(P[i]); }
        }
      }
      let bi = 0; for (let i = 1; i <= n; i++) if (Fv[i] < Fv[bi]) bi = i;
      return { x: P[bi], f: Fv[bi] };
    };
    let best = null;
    try {
      const x0 = orig.map(c => (c && typeof c.v === 'number') ? c.v : 0);
      const starts = [x0];
      for (let k = 0; k < (opts.restarts || 6); k++) starts.push(x0.map(v => v + (Math.random() - 0.5) * 2 * Math.max(1, Math.abs(v)) * (k + 1)));
      for (const s of starts) {
        let cur = { x: s, f: F(s) };
        for (const sc of [1, 0.1, 0.01, 0.001]) {
          mu = 1e3;
          const r = nm(cur.x, cur.x.map(v => Math.max(1e-3, Math.abs(v) * 0.5 + 1) * sc), 1500);
          if (r.f <= cur.f) cur = r;
        }
        if (!best || cur.f < best.f) best = cur;
      }
      // integers / binaries: round, then re-optimise the free variables around the rounded ones
      if (ints.size || bins.size) {
        const x = best.x.slice();
        ints.forEach(i => { x[i] = Math.round(x[i]); }); bins.forEach(i => { x[i] = x[i] >= 0.5 ? 1 : 0; });
        let fb = F(x);
        const iv = [...ints, ...bins];
        if (iv.length <= 6) {
          // small integer problems: try every integer combination in a box around the relaxed answer (at most ~20000)
          const r = Math.max(1, Math.floor((Math.pow(20000, 1 / iv.length) - 2) / 2));
          const lo = iv.map(i => bins.has(i) ? 0 : Math.floor(best.x[i]) - r), hi = iv.map(i => bins.has(i) ? 1 : Math.ceil(best.x[i]) + r);
          const y = x.slice();
          const rec = (k) => {
            if (k === iv.length) { const fy = F(y); if (fy < fb - 1e-12) { fb = fy; for (const i of iv) x[i] = y[i]; } return; }
            for (let v = lo[k]; v <= hi[k]; v++) { y[iv[k]] = v; rec(k + 1); }
          };
          rec(0);
        } else {
          // larger ones: single and paired steps until nothing improves
          for (let pass = 0, better = true; pass < 20 && better; pass++) {
            better = false;
            for (const i of iv) for (const d of [-1, 1]) {
              const y = x.slice(); y[i] += d; if (bins.has(i) && (y[i] < 0 || y[i] > 1)) continue;
              const fy = F(y); if (fy < fb - 1e-12) { x[i] = y[i]; fb = fy; better = true; }
            }
            for (const i of iv) for (const j of iv) if (i < j) for (const di of [-1, 1]) for (const dj of [-1, 1]) {
              const y = x.slice(); y[i] += di; y[j] += dj; if ((bins.has(i) && (y[i] < 0 || y[i] > 1)) || (bins.has(j) && (y[j] < 0 || y[j] > 1))) continue;
              const fy = F(y); if (fy < fb - 1e-12) { x[i] = y[i]; x[j] = y[j]; fb = fy; better = true; }
            }
          }
        }
        best = { x, f: fb };
      }
      set(best.x);
      const vi = violation();
      const ov = num(objCell.r, objCell.c);
      best.ok = vi < 1e-6 && !isNaN(ov) && (mode !== 'value' || Math.abs(ov - target) < 1e-6 * Math.max(1, Math.abs(target)));
      best.value = ov; best.viol = vi;
      best.x = best.x.map((v, i) => (ints.has(i) || bins.has(i)) ? Math.round(v) : (Math.abs(v - Math.round(v)) < 1e-7 ? Math.round(v) : +v.toPrecision(10)));
    } finally {
      keys.forEach((k, i) => { if (orig[i]) sh.cells.set(k, orig[i]); else sh.cells.delete(k); });
      wb.touch();
    }
    return best;
  }
  async function solverDialog() {
    if (editing) commitEdit(0, 0);
    const last = sh.solver || {};
    const v = await form('Solver', [
      { type: 'note', label: 'Finds values for the changing cells that make the objective as large or small as possible, or equal to a value, while keeping every constraint.' },
      { id: 'obj', label: 'Set objective (a formula cell)', value: last.obj || cellName(S().r, S().c) },
      { id: 'mode', label: 'To', type: 'select', value: last.mode || 'max', options: [['max', 'Max'], ['min', 'Min'], ['value', 'Value of']] },
      { id: 'target', label: 'Value (for "Value of")', type: 'number', value: last.target != null ? last.target : 0 },
      { id: 'vars', label: 'By changing cells (e.g. B2:B4 or B2,C5)', value: last.vars || '' },
      { id: 'cons', label: 'Subject to the constraints (one per line: B2 <= 10, B2:B4 >= 0, C7 = 100, B3 int)', type: 'textarea', rows: 5, value: last.cons || '' },
    ], 'Solve');
    if (!v) return;
    const obj = Formula.parseCell(String(v.obj).trim().replace(/\$/g, ''));
    if (!obj) { toast('Enter the objective as a cell reference like D10.'); return; }
    const oc = getCell(obj.r, obj.c);
    if (!oc || oc.f == null) { toast('The objective cell must contain a formula.'); return; }
    const vars = [];
    for (const part of String(v.vars).split(',')) {
      const g = Formula.parseRange(part.trim());
      if (!g) { toast('Enter the changing cells like B2:B4 or B2,C5.'); return; }
      for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
        const cc = getCell(r, c);
        if (cc && cc.f != null) { toast(cellName(r, c) + ' holds a formula. Changing cells must be plain numbers.'); return; }
        vars.push({ r, c });
      }
    }
    if (!vars.length || vars.length > 50) { toast(vars.length ? 'Solver handles up to 50 changing cells.' : 'Enter at least one changing cell.'); return; }
    const pc = parseCons(v.cons);
    if (pc.error) { toast(pc.error, 7000); return; }
    sh.solver = { obj: v.obj, mode: v.mode, target: Number(v.target) || 0, vars: v.vars, cons: v.cons };
    toast('Solving…');
    await new Promise(r => setTimeout(r, 30));
    const res = solve(obj, v.mode, Number(v.target) || 0, vars, pc.list);
    const keep = await form(res.ok ? 'Solver found a solution' : 'Solver could not find a feasible solution', [
      { type: 'note', label: (res.ok ? 'All constraints are satisfied. ' : 'The best values found still break a constraint (total violation ' + (+res.viol.toPrecision(4)) + '). ') + 'Objective ' + cellName(obj.r, obj.c) + ' = ' + (+Number(res.value).toPrecision(10)) + '.' },
      { type: 'note', label: vars.map((p, i) => cellName(p.r, p.c) + ' = ' + res.x[i]).join(',  ') },
    ], res.ok ? 'Keep Solver solution' : 'Keep these values anyway');
    if (!keep) { toast('Original values restored'); return; }
    applyEdit(rec => vars.forEach((p, i) => { const cc = getCell(p.r, p.c); const o = { v: res.x[i] }; if (cc && cc.s) o.s = cc.s; rec(sh, p.r, p.c, o); }));
  }
  $('#solver-run').onclick = solverDialog;

  /* ================= remove duplicates / text to columns ================= */
  async function removeDuplicates() {
    if (editing) commitEdit(0, 0);
    const g = clampRows(regionAround(rng()));
    if (g.r2 <= g.r1) { toast('Select the data to check for duplicates first.'); return; }
    const v = await form('Remove duplicates', [
      { type: 'note', label: 'Checks ' + cellName(g.r1, g.c1) + ':' + cellName(g.r2, g.c2) + '. A row is a duplicate when every cell in it matches an earlier row.' },
      { id: 'hdr', label: 'My data has headers', type: 'checkbox', value: true },
    ], 'Remove');
    if (!v) return;
    const start = g.r1 + (v.hdr ? 1 : 0), seen = new Set(), keep = [];
    for (let r = start; r <= g.r2; r++) {
      const key = [];
      for (let c = g.c1; c <= g.c2; c++) { const cell = getCell(r, c); key.push(cell ? display(cell, r, c) : ''); }
      const k = key.join('\u0001');
      if (seen.has(k)) continue;
      seen.add(k); keep.push(r);
    }
    const removed = (g.r2 - start + 1) - keep.length;
    if (!removed) { toast('No duplicate rows found.'); return; }
    structural(() => {
      const rowsData = keep.map(r => { const row = []; for (let c = g.c1; c <= g.c2; c++) { const cell = sh.cells.get(r * MAXC + c); row.push(cell ? cloneCell(cell) : null); } return { r, row }; });
      for (let r = start; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) sh.cells.delete(r * MAXC + c);
      rowsData.forEach((d, i) => d.row.forEach((cell, j) => {
        if (!cell) return;
        const nr = start + i;
        if (cell.f != null && nr !== d.r) { try { cell.f = Formula.shiftFormula(cell.f, nr - d.r, 0); } catch (e) { /* keep as is */ } }
        sh.cells.set(nr * MAXC + g.c1 + j, cell);
      }));
    });
    toast(removed + ' duplicate row' + (removed === 1 ? '' : 's') + ' removed; ' + keep.length + ' unique row' + (keep.length === 1 ? '' : 's') + ' remain.');
  }
  async function textToColumns() {
    if (editing) commitEdit(0, 0);
    const g = clampRows(rng());
    if (g.c1 !== g.c2) { toast('Select cells in a single column first.'); return; }
    const v = await form('Text to columns', [
      { id: 'd', label: 'Split at', type: 'select', value: ',', options: [[',', 'Comma'], [';', 'Semicolon'], ['\t', 'Tab'], [' ', 'Space'], ['custom', 'Other (type below)']] },
      { id: 'custom', label: 'Other separator', value: '' },
    ], 'Split');
    if (!v) return;
    const sep = v.d === 'custom' ? v.custom : v.d;
    if (!sep) { toast('Type the separator to split at.'); return; }
    let n = 0;
    applyEdit(rec => {
      for (let r = g.r1; r <= g.r2; r++) {
        const cell = getCell(r, g.c1);
        if (!cell || cell.f != null || typeof cell.v !== 'string' || cell.v.indexOf(sep) < 0) continue;
        let parts = cell.v.split(sep);
        if (sep === ' ') parts = parts.filter(p => p !== '');
        parts.forEach((t, i) => {
          t = t.trim();
          if (t === '') { if (i === 0) rec(sh, r, g.c1, cell.s ? { s: cell.s } : null); return; }
          const p = NumFmt.parseInput(t), o = { v: p.v };
          if (p.nf) o.s = wb.styleId({ nf: p.nf }); else if (i === 0 && cell.s) o.s = cell.s;
          rec(sh, r, g.c1 + i, o);
        });
        n++;
      }
    });
    toast(n ? n + ' cell' + (n === 1 ? '' : 's') + ' split.' : 'Nothing to split: no selected cell contains that separator.');
  }
  $('#dd-rm').onclick = removeDuplicates; $('#t2c').onclick = textToColumns;

  /* ================= files ================= */
  const titleEl = $('#title');
  function setTitle(n) { titleEl.value = n; document.title = n + ' - Cells'; }
  titleEl.addEventListener('input', () => { document.title = titleEl.value + ' - Cells'; markDirty(); });
  function markDirty() { if (!state.dirty) { state.dirty = true; $('#dirty').textContent = '● unsaved'; } }
  function markClean() { state.dirty = false; $('#dirty').textContent = ''; }
  function initSheetState() { for (const s of wb.sheets) { s.sel = { r: 0, c: 0, er: 0, ec: 0 }; } }
  function setWorkbook(w) {
    wb = w; initSheetState(); selChart = null; sh = wb.sheets[wb.active] || wb.sheets[0];
    hist.stack = []; hist.i = 0; clip = null;
    wb.touch(); rebuildAxes(); renderTabs(); syncBar();
    scroller.scrollLeft = 0; scroller.scrollTop = 0;
    $('#v-grid').checked = sh.showGrid;
    render();
  }
  async function newBook() {
    if (state.dirty && !(await Lumen.confirmBox('Discard changes?', 'This workbook has unsaved changes. Create a new one anyway?', 'Discard'))) return;
    setWorkbook(new Workbook()); state.handle = null; setTitle('Book1'); markClean();
  }
  async function openDialog() {
    if (state.dirty && !(await Lumen.confirmBox('Discard changes?', 'This workbook has unsaved changes. Open another file anyway?', 'Discard'))) return;
    const r = await FileIO.open([{ description: 'Spreadsheets', accept: { [XlsxIO.XLSX_MIME]: ['.xlsx'], 'text/csv': ['.csv'], 'text/tab-separated-values': ['.tsv'] } }], '.xlsx,.csv,.tsv');
    if (r) await loadFile(r.file, r.handle);
  }
  async function loadFile(file, handle) {
    try {
      const ext = (file.name.match(/\.([^.]+)$/) || [, ''])[1].toLowerCase();
      if (ext === 'xlsx' || ext === 'xlsm' || ext === 'xltx') { setWorkbook(await XlsxIO.read(await file.arrayBuffer())); state.handle = handle; }
      else if (ext === 'csv' || ext === 'tsv' || ext === 'txt') {
        const w = new Workbook();
        const rows = parseTSV(await file.text(), ext === 'tsv' ? '\t' : (await file.slice(0, 2000).text()).split('\t').length > (await file.slice(0, 2000).text()).split(',').length ? '\t' : ',');
        const s = w.sheets[0];
        rows.forEach((row, r) => row.forEach((t, c) => { if (t === '') return; const p = NumFmt.parseInput(t); const cell = { v: p.v }; if (p.nf) cell.s = w.styleId({ nf: p.nf }); s.cells.set(r * MAXC + c, cell); }));
        setWorkbook(w); state.handle = null;
      } else if (ext === 'xls') { toast('Old binary .xls files are not supported. Save as .xlsx first.', 6000); return; }
      else { toast('Unsupported file type: .' + ext); return; }
      setTitle(baseName(file.name)); markClean();
      toast('Opened ' + file.name);
    } catch (err) { console.error(err); toast('Could not open file: ' + err.message, 6000); }
  }
  async function save(saveAs) {
    try {
      if (editing) commitEdit(0, 0);
      sh.scroll = { x: scroller.scrollLeft, y: scroller.scrollTop };
      const blob = await XlsxIO.write(wb);
      const r = await FileIO.save(blob, (titleEl.value || 'Book1') + '.xlsx', FILE_TYPES, state.handle, saveAs);
      if (r) { state.handle = r.handle; if (r.handle) setTitle(baseName(r.name)); markClean(); toast('Saved ' + r.name); }
    } catch (err) { console.error(err); toast('Save failed: ' + err.message, 6000); }
  }
  function exportCsv() {
    const b = sh.bounds();
    const rows = [];
    for (let r = 0; r <= b.maxR; r++) { const row = []; for (let c = 0; c <= b.maxC; c++) { let t = display(getCell(r, c), r, c); if (/[",\n]/.test(t)) t = '"' + t.replace(/"/g, '""') + '"'; row.push(t); } rows.push(row.join(',')); }
    FileIO.save(new Blob([rows.join('\r\n')], { type: 'text/csv' }), sh.name + '.csv', [{ description: 'CSV', accept: { 'text/csv': ['.csv'] } }], null, true);
  }
  $('#btn-new').onclick = newBook; $('#btn-open').onclick = openDialog; $('#btn-save').onclick = () => save(false); $('#btn-saveas').onclick = () => save(true); $('#btn-csv').onclick = exportCsv;
  if (FileIO.onLaunch) FileIO.onLaunch((f, h) => loadFile(f, h));
  FileIO.onDrop(f => loadFile(f, null));
  window.addEventListener('beforeunload', e => { if (state.dirty) { e.preventDefault(); e.returnValue = ''; } });

  /* ================= start ================= */
  initSheetState();
  rebuildAxes(); renderTabs(); syncBar(); resize();
  setTitle('Book1');
  scroller.focus();
  window.LumenCells = { get wb() { return wb; }, get sh() { return sh; }, setWorkbook, loadFile, save, select, startEdit, commitEdit, setFromText, display, undo, redo, formatSel, shiftStructure, sortSel, mergeSel, rebuildAxes, draw, state, makeTable, tableToRange, tableTotals, buildPivot, placePivot, goalSeek, structural, applyEdit, removeDuplicates, textToColumns, refreshPivot, refreshAllPivots, switchSheet, solve, parseCons };
})();
