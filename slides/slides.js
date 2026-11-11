/* Lumen Slides - presentation editor UI */
(function () {
  'use strict';
  const { $, $$, toast, form, FileIO, esc, baseName } = Lumen;
  const stage = $('#stage'), stagebox = $('#stagebox'), view = $('#slideview'), selbox = $('#selbox'), thumbs = $('#thumbs');
  const FILE_TYPES = [{ description: 'PowerPoint presentations', accept: { [PptxIO.PPTX_MIME]: ['.pptx'] } }];
  try { document.execCommand('defaultParagraphSeparator', false, 'p'); document.execCommand('styleWithCSS', false, true); } catch (e) { /* ignore */ }

  let pres = { w: 1280, h: 720, media: {}, slides: [] };
  let cur = 0, selId = null, zoom = 1, editingId = null, nextId = 1, mediaN = 0;
  const state = { handle: null, dirty: false };
  let clipboard = null;
  let selSet = [];
  const THEMES = {
    office: { name: 'Office', bg: 'FFFFFF', fg: '000000', sub: '595959', dk2: '44546A', lt2: 'E7E6E6', acc: ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'], major: 'Calibri Light', minor: 'Calibri' },
    slate: { name: 'Slate', bg: 'F4F5F7', fg: '1F2933', sub: '52606D', dk2: '323F4B', lt2: 'E4E7EB', acc: ['3E4C59', '2CB1BC', 'F0B429', 'E12D39', '7B8794', '27AB83'], major: 'Segoe UI', minor: 'Segoe UI' },
    midnight: { name: 'Midnight', bg: '0F172A', fg: 'F8FAFC', sub: 'CBD5E1', dk2: '1E293B', lt2: 'E2E8F0', acc: ['38BDF8', 'F472B6', 'A3E635', 'FBBF24', 'A78BFA', '34D399'], major: 'Segoe UI', minor: 'Segoe UI' },
    forest: { name: 'Forest', bg: 'F3F7F0', fg: '1B3A2B', sub: '4F6F52', dk2: '2D4A3A', lt2: 'DDE7D5', acc: ['2F7D4F', 'C9A227', '8C5A2B', '5E8C61', 'B5651D', '3B6E8F'], major: 'Georgia', minor: 'Verdana' },
    sunset: { name: 'Sunset', bg: 'FFF7ED', fg: '431407', sub: '9A3412', dk2: '7C2D12', lt2: 'FFEDD5', acc: ['EA580C', 'DB2777', 'F59E0B', '7C3AED', 'DC2626', '0D9488'], major: 'Georgia', minor: 'Arial' },
    ocean: { name: 'Ocean', bg: 'F0F9FF', fg: '0C4A6E', sub: '0369A1', dk2: '075985', lt2: 'E0F2FE', acc: ['0284C7', '0D9488', '6366F1', 'F59E0B', '0EA5E9', '14B8A6'], major: 'Arial', minor: 'Arial' },
    charcoal: { name: 'Charcoal', bg: '262626', fg: 'FAFAFA', sub: 'D4D4D4', dk2: '404040', lt2: 'E5E5E5', acc: ['F97316', 'FACC15', '22D3EE', 'A3A3A3', 'FB7185', '4ADE80'], major: 'Impact', minor: 'Arial' },
    paper: { name: 'Paper', bg: 'FBF8F1', fg: '2B2B2B', sub: '6B6B6B', dk2: '4A4A4A', lt2: 'EFE9DB', acc: ['8B1E3F', '3C6E71', 'D9A441', '284B63', '9A8C98', '6A994E'], major: 'Times New Roman', minor: 'Georgia' },
  };
  const theme = () => pres.theme || THEMES.office;
  const darken = (hex, k) => [0, 2, 4].map(i => ('0' + Math.round(parseInt(hex.slice(i, i + 2), 16) * (k || 0.7)).toString(16)).slice(-2)).join('').toUpperCase();

  /* ---------------- geometry ---------------- */
  function geomPath(t, w, h) {
    const P = (pts) => 'M' + pts.map(p => p[0].toFixed(1) + ' ' + p[1].toFixed(1)).join('L') + 'Z';
    switch (t) {
      case 'roundRect': { const r = Math.min(w, h) * 0.1667; return `M${r} 0H${w - r}Q${w} 0 ${w} ${r}V${h - r}Q${w} ${h} ${w - r} ${h}H${r}Q0 ${h} 0 ${h - r}V${r}Q0 0 ${r} 0Z`; }
      case 'ellipse': return `M0 ${h / 2}A${w / 2} ${h / 2} 0 1 0 ${w} ${h / 2}A${w / 2} ${h / 2} 0 1 0 0 ${h / 2}Z`;
      case 'triangle': return P([[w / 2, 0], [w, h], [0, h]]);
      case 'rtTriangle': return P([[0, 0], [w, h], [0, h]]);
      case 'diamond': return P([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
      case 'pentagon': return P([[w / 2, 0], [w, h * 0.38], [w * 0.81, h], [w * 0.19, h], [0, h * 0.38]]);
      case 'hexagon': return P([[w * 0.25, 0], [w * 0.75, 0], [w, h / 2], [w * 0.75, h], [w * 0.25, h], [0, h / 2]]);
      case 'parallelogram': return P([[w * 0.25, 0], [w, 0], [w * 0.75, h], [0, h]]);
      case 'trapezoid': return P([[w * 0.2, 0], [w * 0.8, 0], [w, h], [0, h]]);
      case 'rightArrow': { const hx = w - Math.min(h * 0.5, w * 0.5); return P([[0, h * .25], [hx, h * .25], [hx, 0], [w, h / 2], [hx, h], [hx, h * .75], [0, h * .75]]); }
      case 'leftArrow': { const hx = Math.min(h * 0.5, w * 0.5); return P([[w, h * .25], [hx, h * .25], [hx, 0], [0, h / 2], [hx, h], [hx, h * .75], [w, h * .75]]); }
      case 'upArrow': { const hy = Math.min(w * 0.5, h * 0.5); return P([[w * .25, h], [w * .25, hy], [0, hy], [w / 2, 0], [w, hy], [w * .75, hy], [w * .75, h]]); }
      case 'downArrow': { const hy = h - Math.min(w * 0.5, h * 0.5); return P([[w * .25, 0], [w * .25, hy], [0, hy], [w / 2, h], [w, hy], [w * .75, hy], [w * .75, 0]]); }
      case 'star5': { const pts = []; for (let i = 0; i < 10; i++) { const a = -Math.PI / 2 + i * Math.PI / 5, r = i % 2 ? 0.382 : 1; pts.push([w / 2 + Math.cos(a) * r * w / 2, h / 2 + Math.sin(a) * r * h / 2]); } return P(pts); }
      default: return `M0 0H${w}V${h}H0Z`;
    }
  }
  const hexc = (c) => c ? '#' + c : 'none';

  /* ---------------- rendering ---------------- */
  const plain = (html) => { const d = document.createElement('div'); d.innerHTML = html || ''; return d.textContent.replace(/​/g, '').trim(); };
  const PROMPT = { title: 'Click to add title', ctrTitle: 'Click to add title', subTitle: 'Click to add subtitle', body: 'Click to add text' };

  function buildShape(s, edit, live) {
    const el = document.createElement('div');
    el.className = 'shape';
    el.dataset.id = s.id;
    el.style.cssText = `left:${s.x}px;top:${s.y}px;width:${s.w}px;height:${s.h}px;` + (s.rot ? `transform:rotate(${s.rot}deg);` : '');
    if (s.type === 'line') {
      const x1 = s.flipH ? s.w : 0, x2 = s.flipH ? 0 : s.w, y1 = s.flipV ? s.h : 0, y2 = s.flipV ? 0 : s.h;
      const lw = s.line ? s.line.w : 2, col = s.line ? s.line.color : '000000';
      const ang = Math.atan2(y2 - y1, x2 - x1), hl = 6 + lw * 2;
      let head = '';
      if (s.tail) head = `<polygon points="${x2},${y2} ${x2 - hl * Math.cos(ang - 0.4)},${y2 - hl * Math.sin(ang - 0.4)} ${x2 - hl * Math.cos(ang + 0.4)},${y2 - hl * Math.sin(ang + 0.4)}" fill="#${col}"/>`;
      el.innerHTML = `<svg><line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="transparent" stroke-width="14"/><line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="#${col}" stroke-width="${lw}"${s.line && s.line.dash ? ' stroke-dasharray="6 4"' : ''}/>${head}</svg>`;
      return el;
    }
    if (s.type === 'image') {
      el.classList.add('img');
      const c = s.crop || { l: 0, t: 0, r: 0, b: 0 };
      const fw = 1 - c.l - c.r, fh = 1 - c.t - c.b;
      const im = document.createElement('img');
      im.src = pres.media[s.img] || '';
      im.draggable = false;
      im.style.cssText = `width:${s.w / fw}px;height:${s.h / fh}px;left:${-c.l / fw * s.w}px;top:${-c.t / fh * s.h}px;`;
      el.appendChild(im);
      return el;
    }
    if (s.type === 'media') {
      el.classList.add('img', 'media');
      const src = pres.media[s.src] || '';
      if (live) {
        const m = document.createElement(s.kind === 'audio' ? 'audio' : 'video');
        m.controls = true; m.src = src; m.preload = 'metadata';
        if (s.kind !== 'audio' && pres.media[s.img]) m.poster = pres.media[s.img];
        m.style.cssText = `position:absolute;left:0;top:${s.kind === 'audio' ? Math.max(0, s.h / 2 - 27) : 0}px;width:${s.w}px;height:${s.kind === 'audio' ? Math.min(54, s.h) : s.h}px;background:${s.kind === 'audio' ? 'transparent' : '#000'}`;
        el.appendChild(m);
      } else {
        const im = document.createElement('img');
        im.src = pres.media[s.img] || ''; im.draggable = false;
        im.style.cssText = `width:${s.w}px;height:${s.h}px;left:0;top:0;`;
        el.appendChild(im);
        const badge = document.createElement('div');
        badge.className = 'mbadge';
        badge.textContent = s.kind === 'audio' ? '♫' : '▶';
        el.appendChild(badge);
      }
      return el;
    }
    if (s.type === 'chart') {
      const d = window.devicePixelRatio || 1;
      const cvs = document.createElement('canvas');
      cvs.width = Math.max(1, s.w * d); cvs.height = Math.max(1, s.h * d);
      cvs.style.cssText = `width:${s.w}px;height:${s.h}px;display:block;background:#fff`;
      const cx = cvs.getContext('2d');
      cx.setTransform(d, 0, 0, d, 0, 0);
      Charts.draw(cx, s.w, s.h, { type: s.chart.type, title: s.chart.title, legend: s.chart.legend }, { cats: s.chart.cats, series: s.chart.series });
      el.appendChild(cvs);
      return el;
    }
    if (s.type === 'table') {
      const t = document.createElement('table');
      t.style.cssText = `width:${s.colW.reduce((a, b) => a + b, 0)}px;font-size:${s.fs || 14}pt;font-family:'${s.ff || 'Calibri'}';color:#${s.fc || '000'};line-height:1.2`;
      t.innerHTML = '<colgroup>' + s.colW.map(w => `<col style="width:${w}px">`).join('') + '</colgroup>';
      const tb = document.createElement('tbody');
      s.rows.forEach((row, ri) => {
        const tr = document.createElement('tr');
        tr.style.height = s.rowH[ri] + 'px';
        row.forEach((cell, ci) => {
          const td = document.createElement('td');
          td.innerHTML = cell.html;
          if (cell.fill) td.style.background = '#' + cell.fill;
          td.dataset.r = ri; td.dataset.c = ci;
          if (edit) td.contentEditable = 'true';
          tr.appendChild(td);
        });
        tb.appendChild(tr);
      });
      t.appendChild(tb);
      el.appendChild(t);
      return el;
    }
    if (s.type !== 'text') {
      const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
      svg.setAttribute('viewBox', `0 0 ${s.w} ${s.h}`);
      svg.setAttribute('preserveAspectRatio', 'none');
      const p = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      p.setAttribute('d', geomPath(s.type, s.w, s.h));
      p.setAttribute('fill', hexc(s.fill));
      if (s.line) { p.setAttribute('stroke', '#' + s.line.color); p.setAttribute('stroke-width', s.line.w); if (s.line.dash) p.setAttribute('stroke-dasharray', '6 4'); } else p.setAttribute('stroke', 'none');
      svg.appendChild(p);
      el.appendChild(svg);
    } else {
      if (s.fill) el.style.background = '#' + s.fill;
      if (s.line) el.style.border = `${s.line.w}px solid #${s.line.color}`;
    }
    const txt = document.createElement('div');
    txt.className = 'txt';
    txt.style.justifyContent = { t: 'flex-start', ctr: 'center', b: 'flex-end' }[s.valign || 't'];
    const tc = document.createElement('div');
    tc.className = 'tc';
    tc.style.cssText = `font-size:${s.fs || 18}pt;color:#${s.fc || '000000'};font-family:'${s.ff || 'Calibri'}'`;
    tc.innerHTML = s.html || '<p><br></p>';
    txt.appendChild(tc);
    el.appendChild(txt);
    if (edit && s.ph && !plain(s.html)) {
      const pr = document.createElement('div');
      pr.className = 'prompt';
      pr.textContent = PROMPT[s.ph] || 'Click to add text';
      pr.style.fontSize = Math.min(28, (s.fs || 24)) + 'pt';
      el.appendChild(pr);
    }
    if (edit && s.id === editingId) { el.classList.add('editing'); tc.contentEditable = 'true'; }
    return el;
  }
  function renderSlideInto(host, slide, edit, live) {
    const sno = pres.slides.indexOf(slide) + 1;
    for (const s of slide.shapes) if (s.ph === 'sldNum' && sno > 0) s.html = '<p style="text-align:right">' + sno + '</p>';
    host.innerHTML = '';
    host.style.background = '#' + (slide.bg || 'FFFFFF');
    if (!slide.isMaster && !slide.noMaster && pres.master && pres.master.shapes) for (const s of pres.master.shapes) {
      const el = buildShape(s, false, live); el.classList.add('mshape'); el.removeAttribute('data-id'); host.appendChild(el);
    }
    for (const s of slide.shapes) host.appendChild(buildShape(s, edit, live));
  }
  let masterMode = false;
  const masterView = {
    isMaster: true, layout: 'blank', notes: '',
    get bg() { const s = pres.slides[cur]; return (s && s.bg) || theme().bg; }, set bg(v) { },
    get shapes() { return pres.master.shapes; }, set shapes(v) { pres.master.shapes = v; },
  };
  const slide = () => masterMode ? masterView : pres.slides[cur];
  const shapeOf = (id) => slide() && slide().shapes.find(s => s.id === id);

  function fitZoom() {
    zoom = Math.max(0.2, Math.min((stage.clientWidth - 48) / pres.w, (stage.clientHeight - 48) / pres.h));
    layoutStage();
  }
  function layoutStage() {
    stagebox.style.width = pres.w * zoom + 'px'; stagebox.style.height = pres.h * zoom + 'px';
    view.style.width = pres.w + 'px'; view.style.height = pres.h + 'px';
    view.style.transform = `scale(${zoom})`;
    $('#st-zoom').textContent = Math.round(zoom * 100) + '%';
    positionSel();
  }
  function renderStage() {
    if (!slide()) return;
    renderSlideInto(view, slide(), true);
    layoutStage();
    positionSel();
    $('#st-slide').textContent = masterMode ? 'Slide Master' : `Slide ${cur + 1} of ${pres.slides.length}`;
    const nt = $('#notes');
    if (document.activeElement !== nt) nt.value = slide().notes || '';
    if (typeof syncSlideTools === 'function') syncSlideTools();
  }
  $('#notes').addEventListener('input', e => { if (slide()) { slide().notes = e.target.value; markDirty(); } });
  $('#notes').addEventListener('change', () => commit());
  let thumbTimer = 0;
  function renderThumbs() { clearTimeout(thumbTimer); thumbTimer = setTimeout(doThumbs, 60); }
  function doThumbs() {
    thumbs.innerHTML = '';
    const sc = 178 / pres.w;
    if (masterMode) {
      const t = document.createElement('div');
      t.className = 'thumb active';
      t.innerHTML = '<span class="num">M</span><div class="tv"><div class="slidehost"></div></div>';
      const tv = $('.tv', t), host = $('.slidehost', t);
      tv.style.height = pres.h * sc + 'px';
      host.style.width = pres.w + 'px'; host.style.height = pres.h + 'px'; host.style.transform = `scale(${sc})`;
      renderSlideInto(host, masterView, false);
      thumbs.appendChild(t);
      return;
    }
    pres.slides.forEach((sl, i) => {
      const t = document.createElement('div');
      t.className = 'thumb' + (i === cur ? ' active' : '') + (sl.hidden ? ' hiddenslide' : '');
      t.draggable = true;
      t.innerHTML = `<span class="num">${i + 1}</span><div class="tv"><div class="slidehost"></div></div>`;
      const tv = $('.tv', t), host = $('.slidehost', t);
      tv.style.height = pres.h * sc + 'px';
      host.style.width = pres.w + 'px'; host.style.height = pres.h + 'px'; host.style.transform = `scale(${sc})`;
      renderSlideInto(host, sl, false);
      t.addEventListener('click', () => gotoSlide(i));
      t.addEventListener('dragstart', e => { e.dataTransfer.setData('text/x-slide', String(i)); });
      t.addEventListener('dragover', e => { e.preventDefault(); t.classList.add('drop'); });
      t.addEventListener('dragleave', () => t.classList.remove('drop'));
      t.addEventListener('drop', e => {
        e.preventDefault(); t.classList.remove('drop');
        const from = parseInt(e.dataTransfer.getData('text/x-slide'), 10);
        if (isNaN(from) || from === i) return;
        const [m] = pres.slides.splice(from, 1);
        pres.slides.splice(i, 0, m);
        cur = i; commit();
      });
      thumbs.appendChild(t);
    });
  }
  function gotoSlide(i) { if (masterMode) return; finishEdit(); cur = Math.max(0, Math.min(pres.slides.length - 1, i)); selId = null; selSet = []; renderStage(); renderThumbs(); }

  /* ---------------- history ---------------- */
  const hist = { stack: [], i: -1 };
  const snap = () => JSON.stringify({ w: pres.w, h: pres.h, theme: pres.theme || null, master: pres.master || null, slides: pres.slides });
  function commit() {
    const s = snap();
    if (hist.i >= 0 && hist.stack[hist.i] === s) { renderThumbs(); return; }
    hist.stack.length = hist.i + 1; hist.stack.push(s); hist.i++;
    if (hist.stack.length > 100) { hist.stack.shift(); hist.i--; }
    markDirty(); renderStage(); renderThumbs();
  }
  function restoreHist(i) {
    const o = JSON.parse(hist.stack[i]);
    pres.w = o.w; pres.h = o.h; pres.slides = o.slides; pres.theme = o.theme || undefined; pres.master = o.master || (masterMode ? { shapes: [] } : undefined); selSet = [];
    cur = Math.min(cur, pres.slides.length - 1); if (!shapeOf(selId)) selId = null;
    editingId = null; markDirty(); fitZoom(); renderStage(); renderThumbs();
  }
  function undo() { finishEdit(); if (hist.i > 0) restoreHist(--hist.i); }
  function redo() { if (hist.i < hist.stack.length - 1) restoreHist(++hist.i); }
  function markDirty() { if (!state.dirty) { state.dirty = true; $('#dirty').textContent = '● unsaved'; } }
  function markClean() { state.dirty = false; $('#dirty').textContent = ''; }

  /* ---------------- selection ---------------- */
  const dimW = (s) => s.type === 'table' ? s.colW.reduce((a, b) => a + b, 0) : s.w;
  const dimH = (s) => s.type === 'table' ? s.rowH.reduce((a, b) => a + b, 0) : s.h;
  function groupIds(id) { const s = shapeOf(id); if (!s) return []; return (s.g !== undefined && s.g !== null) ? slide().shapes.filter(x => x.g === s.g).map(x => x.id) : [id]; }
  function selShapes() {
    if (selId === null || !shapeOf(selId)) return [];
    if (!selSet.includes(selId)) return [shapeOf(selId)];
    return selSet.map(shapeOf).filter(Boolean);
  }
  function bbox(list) {
    let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
    for (const s of list) { x1 = Math.min(x1, s.x); y1 = Math.min(y1, s.y); x2 = Math.max(x2, s.x + dimW(s)); y2 = Math.max(y2, s.y + dimH(s)); }
    return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
  }
  function select(id, add) {
    if (id !== selId) finishEdit();
    if (id === null || id === undefined) { selId = null; selSet = []; }
    else {
      const ids = groupIds(id);
      const curSet = selShapes().map(s => s.id);
      if (add) {
        if (curSet.includes(id)) { selSet = curSet.filter(x => !ids.includes(x)); selId = selSet.length ? selSet[selSet.length - 1] : null; }
        else { selSet = [...new Set([...curSet, ...ids])]; selId = id; }
      } else if (curSet.length > 1 && curSet.includes(id)) { selSet = curSet; selId = id; }
      else { selSet = ids; selId = id; }
    }
    positionSel(); syncTools();
  }
  function positionSel() {
    const many = selShapes();
    const stSel = $('#st-sel');
    if (many.length > 1) {
      const b = bbox(many);
      selbox.hidden = false; selbox.classList.add('multi');
      selbox.style.left = b.x * zoom + 'px'; selbox.style.top = b.y * zoom + 'px';
      selbox.style.width = b.w * zoom + 'px'; selbox.style.height = b.h * zoom + 'px';
      selbox.style.transform = '';
      if (stSel) stSel.textContent = many.length + ' shapes selected' + (many.every(s => s.g !== undefined && s.g !== null && s.g === many[0].g) ? ' (group)' : '');
      return;
    }
    selbox.classList.remove('multi');
    if (stSel) stSel.textContent = '';
    const s = shapeOf(selId);
    if (!s) { selbox.hidden = true; return; }
    selbox.hidden = false;
    selbox.style.left = s.x * zoom + 'px'; selbox.style.top = s.y * zoom + 'px';
    const w = s.type === 'table' ? s.colW.reduce((a, b) => a + b, 0) : s.w, h = s.type === 'table' ? s.rowH.reduce((a, b) => a + b, 0) : s.h;
    selbox.style.width = w * zoom + 'px'; selbox.style.height = h * zoom + 'px';
    selbox.style.transform = s.rot ? `rotate(${s.rot}deg)` : '';
    selbox.style.display = editingId === selId ? 'none' : 'block';
  }
  function syncTools() {
    const s = shapeOf(selId);
    if (!s) return;
    if (s.fill) { $('#fill').value = '#' + s.fill; $('#sw-fill').style.background = '#' + s.fill; }
    if (s.line) { $('#line').value = '#' + s.line.color; $('#sw-line').style.background = '#' + s.line.color; $('#line-w').value = String([0, 1, 2, 4, 6].includes(Math.round(s.line.w)) ? Math.round(s.line.w) : 2); }
    $('#f-size').value = String(s.fs || 18);
    if (typeof syncSlideTools === 'function') syncSlideTools();
  }

  /* ---------------- mouse ---------------- */
  let drag = null;
  const pt = (e) => ({ x: e.clientX / zoom, y: e.clientY / zoom });
  stagebox.addEventListener('mousedown', e => {
    if (e.target.dataset && e.target.dataset.h) return; // handled below
    const el = e.target.closest && e.target.closest('.shape');
    if (!el) { if (e.target === view || e.target === stagebox) { finishEdit(); select(null); } return; }
    const id = parseInt(el.dataset.id, 10);
    if (editingId === id && e.target.closest('.tc')) return; // editing text
    if (e.target.closest('td') && selId === id) return; // editing table cell
    if (e.target.closest('td')) { select(id); return; }
    e.preventDefault();
    if (editingId !== null) finishEdit();
    select(id, e.shiftKey);
    const items = selShapes();
    if (!items.length || !items.some(s => s.id === id)) return;
    drag = { t: 'move', id, p0: pt(e), items: items.map(s => ({ s, x0: s.x, y0: s.y })), moved: false };
  });
  selbox.addEventListener('mousedown', e => {
    const h = e.target.dataset && e.target.dataset.h;
    if (!h) return;
    e.preventDefault(); e.stopPropagation();
    const s = shapeOf(selId);
    const many = selShapes();
    if (many.length > 1) {
      if (h === 'rot') return;
      const box = bbox(many);
      drag = { t: 'gsize', h, id: selId, p0: pt(e), box, items: many.map(q => ({ s: q, x: q.x, y: q.y, w: dimW(q), h: dimH(q), cw: q.colW && q.colW.slice(), rh: q.rowH && q.rowH.slice() })) };
      return;
    }
    if (h === 'rot') { const r = selbox.getBoundingClientRect(); drag = { t: 'rot', id: selId, cx: r.left + r.width / 2, cy: r.top + r.height / 2 }; return; }
    drag = { t: 'size', h, id: selId, p0: pt(e), o: { x: s.x, y: s.y, w: s.type === 'table' ? s.colW.reduce((a, b) => a + b, 0) : s.w, h: s.type === 'table' ? s.rowH.reduce((a, b) => a + b, 0) : s.h }, cw: s.colW && s.colW.slice(), rh: s.rowH && s.rowH.slice(), shift: e.shiftKey };
  });
  window.addEventListener('mousemove', e => {
    if (!drag) return;
    const s = shapeOf(drag.id);
    if (!s) return;
    const el = view.querySelector(`.shape[data-id="${s.id}"]`);
    if (drag.t === 'move') {
      const p = pt(e);
      const dx = p.x - drag.p0.x, dy = p.y - drag.p0.y;
      if (!drag.moved && Math.abs(dx) + Math.abs(dy) < 3 / zoom) return;
      drag.moved = true;
      for (const it of drag.items) {
        it.s.x = Math.round(it.x0 + dx); it.s.y = Math.round(it.y0 + dy);
        const e2 = view.querySelector(`.shape[data-id="${it.s.id}"]`);
        if (e2) { e2.style.left = it.s.x + 'px'; e2.style.top = it.s.y + 'px'; }
      }
    } else if (drag.t === 'gsize') {
      const p = pt(e), o = drag.box, h = drag.h;
      const dx = p.x - drag.p0.x, dy = p.y - drag.p0.y;
      let x = o.x, y = o.y, w = o.w, hh = o.h;
      if (h.includes('e')) w = o.w + dx; if (h.includes('s')) hh = o.h + dy;
      if (h.includes('w')) { w = o.w - dx; x = o.x + dx; } if (h.includes('n')) { hh = o.h - dy; y = o.y + dy; }
      if (h.length === 2 && !e.altKey) { const r = o.w / o.h; if (Math.abs(w - o.w) > Math.abs(hh - o.h) * r) hh = w / r; else w = hh * r; if (h.includes('w')) x = o.x + o.w - w; if (h.includes('n')) y = o.y + o.h - hh; }
      w = Math.max(12, w); hh = Math.max(12, hh);
      const kx = w / Math.max(1, o.w), ky = hh / Math.max(1, o.h);
      for (const it of drag.items) {
        const q = it.s;
        q.x = Math.round(x + (it.x - o.x) * kx); q.y = Math.round(y + (it.y - o.y) * ky);
        if (q.type === 'table') { q.colW = it.cw.map(c => c * kx); q.rowH = it.rh.map(c => c * ky); q.w = it.w * kx; q.h = it.h * ky; }
        else { q.w = Math.max(1, Math.round(it.w * kx)); q.h = Math.max(q.type === 'line' ? 0 : 1, Math.round(it.h * ky)); }
      }
      renderStage(); drag.moved = true; positionSel(); return;
    } else if (drag.t === 'rot') {
      let a = Math.atan2(e.clientY - drag.cy, e.clientX - drag.cx) * 180 / Math.PI + 90;
      if (e.shiftKey) a = Math.round(a / 15) * 15;
      s.rot = Math.round(((a % 360) + 360) % 360);
      el.style.transform = `rotate(${s.rot}deg)`;
    } else if (drag.t === 'size') {
      const p = pt(e), o = drag.o, h = drag.h;
      let dx = p.x - drag.p0.x, dy = p.y - drag.p0.y;
      let x = o.x, y = o.y, w = o.w, hh = o.h;
      if (h.includes('e')) w = o.w + dx; if (h.includes('s')) hh = o.h + dy;
      if (h.includes('w')) { w = o.w - dx; x = o.x + dx; } if (h.includes('n')) { hh = o.h - dy; y = o.y + dy; }
      if ((e.shiftKey || s.type === 'image') && h.length === 2) { const r = o.w / o.h; if (Math.abs(w - o.w) > Math.abs(hh - o.h) * r) hh = w / r; else w = hh * r; if (h.includes('w')) x = o.x + o.w - w; if (h.includes('n')) y = o.y + o.h - hh; }
      w = Math.max(12, w); hh = Math.max(12, hh);
      s.x = Math.round(x); s.y = Math.round(y);
      if (s.type === 'table') { const kx = w / o.w, ky = hh / o.h; s.colW = drag.cw.map(c => c * kx); s.rowH = drag.rh.map(r => r * ky); s.w = w; s.h = hh; renderStage(); drag.moved = true; return; }
      s.w = Math.round(w); s.h = Math.round(hh);
      const fresh = buildShape(s, true);
      el.replaceWith(fresh);
    }
    drag.moved = true;
    positionSel();
  });
  window.addEventListener('mouseup', () => { if (drag && drag.moved) commit(); drag = null; });
  stagebox.addEventListener('dblclick', e => {
    const el = e.target.closest && e.target.closest('.shape');
    if (!el) return;
    const s = shapeOf(parseInt(el.dataset.id, 10));
    if (s && s.type === 'chart') { chartDialog(s); return; }
    if (s && ['image', 'line', 'table', 'media', 'chart'].indexOf(s.type) < 0) { selSet = [s.id]; selId = s.id; startEdit(s.id); }
  });

  /* ---------------- charts ---------------- */
  const chartToCsv = (ch) => [['Category', ...ch.series.map(s => s.name)], ...ch.cats.map((c, i) => [c, ...ch.series.map(s => s.vals[i] === null || s.vals[i] === undefined ? '' : s.vals[i])])].map(r => r.join(',')).join('\n');
  function csvToChart(txt, base) {
    const rows = txt.replace(/\r/g, '').split('\n').map(l => l.split(',').map(x => x.trim())).filter(r => r.some(x => x !== ''));
    if (rows.length < 2 || rows[0].length < 2) return null;
    const head = rows[0];
    const series = [];
    for (let c = 1; c < head.length; c++) series.push({ name: head[c] || 'Series' + c, color: (base && base.series[c - 1] && base.series[c - 1].color) || Charts.PALETTE[(c - 1) % Charts.PALETTE.length], vals: rows.slice(1).map(r => { const n = parseFloat(r[c]); return isNaN(n) ? null : n; }) });
    return { cats: rows.slice(1).map(r => r[0] || ''), series };
  }
  async function chartDialog(s) {
    const isNew = !s;
    const ch = s ? s.chart : { type: 'col', title: 'Chart Title', legend: true, cats: ['Q1', 'Q2', 'Q3', 'Q4'], series: [{ name: 'Sales', color: Charts.PALETTE[0], vals: [12, 19, 8, 15] }, { name: 'Costs', color: Charts.PALETTE[1], vals: [8, 11, 6, 9] }] };
    const v = await form(isNew ? 'Insert chart' : 'Edit chart', [
      { id: 'type', label: 'Type', type: 'select', value: ch.type, options: [['col', 'Column'], ['bar', 'Bar'], ['line', 'Line'], ['area', 'Area'], ['pie', 'Pie']] },
      { id: 'title', label: 'Title', value: ch.title },
      { id: 'legend', label: 'Show legend', type: 'checkbox', value: ch.legend !== false },
      { type: 'note', label: 'Data: first row = series names, first column = categories, comma separated.' },
      { id: 'data', label: 'Data', type: 'textarea', rows: 7, value: chartToCsv(ch) },
    ], 'OK');
    if (!v) return;
    const parsed = csvToChart(v.data, ch);
    if (!parsed) { toast('Enter at least a header row and one data row.'); return; }
    const nc = { type: v.type, title: v.title, legend: v.legend, cats: parsed.cats, series: v.type === 'pie' ? parsed.series.slice(0, 1) : parsed.series };
    finishEdit();
    if (isNew) addShape({ type: 'chart', x: Math.round(pres.w / 2 - 300), y: Math.round(pres.h / 2 - 190), w: 600, h: 380, rot: 0, chart: nc, fill: null, line: null });
    else s.chart = nc;
    commit();
  }
  $('#i-chart').onclick = () => chartDialog(null);

  /* ---------------- transitions, animations, align ---------------- */
  const setTrans = (fn) => { const sl = slide(); fn(sl); commit(); };
  $('#tr-type').onchange = e => setTrans(sl => { sl.trans = e.target.value === 'none' ? null : { type: e.target.value, speed: $('#tr-speed').value }; });
  $('#tr-speed').onchange = e => setTrans(sl => { if (sl.trans) sl.trans.speed = e.target.value; });
  $('#tr-adv').onchange = e => setTrans(sl => { const n = parseFloat(e.target.value); sl.advAfter = n > 0 ? n : 0; });
  $('#tr-all').onclick = () => { const t = slide().trans ? { ...slide().trans } : null; const a = slide().advAfter || 0; pres.slides.forEach(sl => { sl.trans = t ? { ...t } : null; sl.advAfter = a; }); commit(); toast('Transition applied to all slides'); };
  $('#an-type').onchange = e => { const ss = selShapes(); if (!ss.length) { toast('Select a shape first.'); e.target.value = 'none'; return; } const av = e.target.value; ss.forEach(s => { s.anim = av === 'none' ? undefined : av; }); commit(); };
  function syncSlideTools() {
    const sl = slide();
    if (!sl) return;
    $('#tr-type').value = sl.trans ? sl.trans.type : 'none';
    $('#tr-speed').value = (sl.trans && sl.trans.speed) || 'med';
    $('#tr-adv').value = sl.advAfter || '';
    const s = shapeOf(selId);
    $('#an-type').value = (s && s.anim) || 'none';
    const hb = $('#s-hide'); if (hb) hb.classList.toggle('on', !!sl.hidden);
    const nm = $('#bg-nomaster'); if (nm) { nm.checked = !!sl.noMaster; nm.disabled = masterMode; }
  }
  $('#b-align').onchange = e => {
    const v = e.target.value; e.target.value = '';
    if (!v) return;
    const many = selShapes();
    if (!many.length) return;
    finishEdit();
    if (v === 'disth' || v === 'distv') {
      if (many.length < 3) { toast('Select at least three shapes to distribute.'); return; }
      const hz = v === 'disth';
      const arr = many.slice().sort((a, b) => hz ? a.x - b.x : a.y - b.y);
      const size = (s) => hz ? dimW(s) : dimH(s);
      const first = arr[0], last = arr[arr.length - 1];
      const span = (hz ? last.x : last.y) + size(last) - (hz ? first.x : first.y);
      const gap = (span - arr.reduce((a, s) => a + size(s), 0)) / (arr.length - 1);
      let pos = hz ? first.x : first.y;
      for (const s of arr) { if (hz) s.x = Math.round(pos); else s.y = Math.round(pos); pos += size(s) + gap; }
      commit(); return;
    }
    const box = many.length > 1 ? bbox(many) : { x: 0, y: 0, w: pres.w, h: pres.h };
    for (const s of many) {
      const w = dimW(s), h = dimH(s);
      if (v === 'left') s.x = box.x; else if (v === 'center') s.x = Math.round(box.x + (box.w - w) / 2); else if (v === 'right') s.x = Math.round(box.x + box.w - w);
      else if (v === 'top') s.y = box.y; else if (v === 'middle') s.y = Math.round(box.y + (box.h - h) / 2); else if (v === 'bottom') s.y = Math.round(box.y + box.h - h);
    }
    commit();
  };

  /* ---------------- text editing ---------------- */
  const tcOf = (id) => view.querySelector(`.shape[data-id="${id}"] .tc`);
  function cleanTc(tc) {
    const d = tc.cloneNode(true);
    d.querySelectorAll('[contenteditable]').forEach(n => n.removeAttribute('contenteditable'));
    for (const c of [...d.children]) if (c.tagName === 'DIV') { const p = document.createElement('p'); p.innerHTML = c.innerHTML; c.replaceWith(p); }
    let h = d.innerHTML.replace(/​/g, '');
    if (!d.textContent.replace(/​/g, '') && !d.querySelector('img')) h = h.includes('<p') ? h : '<p><br></p>';
    return h || '<p><br></p>';
  }
  function startEdit(id, selectAllText) {
    finishEdit();
    editingId = id;
    renderStage();
    const tc = tcOf(id);
    if (!tc) return;
    tc.focus();
    const r = document.createRange();
    r.selectNodeContents(tc);
    if (!selectAllText) r.collapse(false);
    const sel = getSelection(); sel.removeAllRanges(); sel.addRange(r);
    positionSel();
  }
  function finishEdit() {
    if (editingId === null) return;
    const id = editingId, tc = tcOf(id), s = shapeOf(id);
    editingId = null;
    if (tc && s) { const h = cleanTc(tc); if (h !== s.html) { s.html = h; commit(); return; } }
    renderStage();
  }
  stagebox.addEventListener('focusout', e => {
    const td = e.target.closest && e.target.closest('td');
    if (td && selId !== null) {
      const s = shapeOf(selId);
      if (s && s.type === 'table') { const cell = s.rows[+td.dataset.r][+td.dataset.c]; const h = cleanTc(td); if (cell.html !== h) { cell.html = h; markDirty(); hist.stack.length = hist.i + 1; hist.stack.push(snap()); hist.i++; renderThumbs(); } }
    }
  });
  stagebox.addEventListener('input', () => markDirty());
  stagebox.addEventListener('keydown', e => {
    if (e.key === 'Escape' && editingId !== null) { e.preventDefault(); finishEdit(); return; }
    if (e.key === 'Tab' && e.target.closest && e.target.closest('.tc')) {
      e.preventDefault();
      const sel = getSelection();
      let n = sel.anchorNode; while (n && !(n.nodeType === 1 && n.parentNode && n.parentNode.classList && n.parentNode.classList.contains('tc'))) n = n.parentNode;
      if (n && n.dataset && n.dataset.bu) {
        const l = Math.max(0, Math.min(8, (parseInt(n.dataset.l, 10) || 0) + (e.shiftKey ? -1 : 1)));
        n.dataset.l = l; n.style.marginLeft = (24 + l * 36) + 'px';
      }
    }
  });

  /* formatting helpers */
  let savedRange = null;
  document.addEventListener('selectionchange', () => { const s = getSelection(); if (s.rangeCount && view.contains(s.anchorNode)) savedRange = s.getRangeAt(0).cloneRange(); });
  function restoreRange() { if (!savedRange) return; const s = getSelection(); if (!(s.rangeCount && view.contains(s.anchorNode))) { s.removeAllRanges(); s.addRange(savedRange); } }
  function withText(fn) {
    if (editingId !== null || (document.activeElement && document.activeElement.closest && document.activeElement.closest('td'))) {
      restoreRange(); fn(); markDirty();
      if (editingId !== null) { const tc = tcOf(editingId), s = shapeOf(editingId); if (tc && s) s.html = cleanTc(tc); }
      return;
    }
    const s = shapeOf(selId);
    if (!s || ['image', 'line', 'table', 'media', 'chart'].includes(s.type)) return;
    const tc = tcOf(selId);
    if (!tc) return;
    tc.contentEditable = 'true'; tc.focus();
    const r = document.createRange(); r.selectNodeContents(tc);
    const sel = getSelection(); sel.removeAllRanges(); sel.addRange(r);
    fn();
    s.html = cleanTc(tc);
    tc.contentEditable = 'false'; tc.blur(); sel.removeAllRanges();
    commit();
  }
  const exec = (c, v) => withText(() => document.execCommand(c, false, v));
  $('.ribbon').addEventListener('mousedown', e => { if (!e.target.closest('select,input')) e.preventDefault(); });
  $$('[data-cmd]').forEach(b => b.onclick = () => exec(b.dataset.cmd));
  $('#tc').addEventListener('input', e => { $('#sw-tc').style.background = e.target.value; exec('foreColor', e.target.value); });
  $('#tc').addEventListener('mousedown', restoreRange);
  $('#apply-tc').onclick = () => exec('foreColor', $('#tc').value);
  $('#f-name').onchange = e => exec('fontName', e.target.value);
  $('#f-size').onchange = e => {
    const pt = parseFloat(e.target.value);
    withText(() => {
      const sel = getSelection();
      if (!sel.rangeCount) return;
      const range = sel.getRangeAt(0);
      if (range.collapsed) { const sp = document.createElement('span'); sp.style.fontSize = pt + 'pt'; sp.textContent = '​'; range.insertNode(sp); const r = document.createRange(); r.setStart(sp.firstChild, 1); r.collapse(true); sel.removeAllRanges(); sel.addRange(r); return; }
      document.execCommand('fontSize', false, '7');
      view.querySelectorAll('font[size="7"], span[style*="xxx-large"]').forEach(el => {
        let sp = el;
        if (el.tagName === 'FONT') { sp = document.createElement('span'); while (el.firstChild) sp.appendChild(el.firstChild); el.replaceWith(sp); }
        sp.style.fontSize = pt + 'pt';
        sp.querySelectorAll('[style*="font-size"]').forEach(x => x.style.removeProperty('font-size'));
      });
    });
  };
  function blocksNow() {
    const tc = editingId !== null ? tcOf(editingId) : tcOf(selId);
    if (!tc) return [];
    const sel = getSelection();
    if (editingId === null || !sel.rangeCount) return [...tc.children];
    const r = sel.getRangeAt(0);
    const out = [...tc.children].filter(c => r.intersectsNode(c));
    return out.length ? out : [...tc.children];
  }
  function toggleBullet(kind) {
    withText(() => {
      const bl = blocksNow();
      const on = bl.every(b => b.dataset.bu === kind);
      bl.forEach(b => {
        if (on) { delete b.dataset.bu; delete b.dataset.l; b.style.marginLeft = ''; b.style.textIndent = ''; b.style.removeProperty('--ind'); }
        else { b.dataset.bu = kind; b.style.marginLeft = '28px'; b.style.textIndent = '-28px'; b.style.setProperty('--ind', '28px'); }
      });
    });
  }
  $('#b-bullets').onclick = () => toggleBullet('dot');
  $('#b-numbers').onclick = () => toggleBullet('num');
  $$('[data-va]').forEach(b => b.onclick = () => { const s = shapeOf(selId); if (s) { s.valign = b.dataset.va; commit(); } });

  /* shape properties */
  const sprop = (fn) => { const ss = selShapes(); if (!ss.length) return; finishEdit(); ss.forEach(s => fn(s)); commit(); };
  $('#fill').addEventListener('input', e => { $('#sw-fill').style.background = e.target.value; sprop(s => { s.fill = e.target.value.slice(1).toUpperCase(); if (s.type === 'text' && !s.ph) s.type = 'rect'; }); });
  $('#apply-fill').onclick = () => sprop(s => { s.fill = $('#fill').value.slice(1).toUpperCase(); });
  $('#b-nofill').onclick = () => sprop(s => { s.fill = null; });
  $('#line').addEventListener('input', e => { $('#sw-line').style.background = e.target.value; sprop(s => { s.line = { color: e.target.value.slice(1).toUpperCase(), w: (s.line && s.line.w) || 2 }; }); });
  $('#apply-line').onclick = () => sprop(s => { s.line = { color: $('#line').value.slice(1).toUpperCase(), w: (s.line && s.line.w) || 2 }; });
  $('#line-w').onchange = e => sprop(s => { const w = parseFloat(e.target.value); s.line = w ? { color: (s.line && s.line.color) || '2F528F', w } : null; });
  $('#b-front').onclick = () => sprop(s => { const a = slide().shapes; a.splice(a.indexOf(s), 1); a.push(s); });
  $('#b-back').onclick = () => sprop(s => { const a = slide().shapes; a.splice(a.indexOf(s), 1); a.unshift(s); });
  $('#b-delshape').onclick = delShape; $('#b-dupshape').onclick = dupShape;
  function delShape() { const ss = selShapes(); if (!ss.length) return; editingId = null; slide().shapes = slide().shapes.filter(s => !ss.includes(s)); selId = null; selSet = []; commit(); }
  function pasteShapes(list, off) {
    const gmap = {}, ids = [];
    for (const s of list) {
      const c = JSON.parse(JSON.stringify(s));
      c.id = nextId++; c.x += off; c.y += off; delete c.xid;
      if (c.g !== undefined && c.g !== null) { if (!gmap[c.g]) gmap[c.g] = 'g' + (nextId++); c.g = gmap[c.g]; }
      slide().shapes.push(c); ids.push(c.id);
    }
    selSet = ids; selId = ids.length ? ids[ids.length - 1] : null;
    commit();
  }
  function dupShape() { const ss = selShapes(); if (ss.length) pasteShapes(ss, 24); }
  function groupSel() {
    const ss = selShapes();
    if (ss.length < 2) { toast('Select two or more shapes (Shift+click) to group.'); return; }
    if (ss.some(s => s.ph)) { toast('Placeholders cannot be grouped.'); return; }
    finishEdit();
    const gk = 'g' + (nextId++);
    ss.forEach(s => { s.g = gk; });
    selSet = ss.map(s => s.id);
    commit(); toast('Grouped ' + ss.length + ' shapes');
  }
  function ungroupSel() {
    const ss = selShapes().filter(s => s.g !== undefined && s.g !== null);
    if (!ss.length) { toast('Select a group first.'); return; }
    ss.forEach(s => { delete s.g; });
    commit(); toast('Ungrouped');
  }
  function selectAll() { finishEdit(); selSet = slide().shapes.map(s => s.id); selId = selSet.length ? selSet[selSet.length - 1] : null; positionSel(); syncTools(); }
  $('#b-group').onclick = groupSel; $('#b-ungroup').onclick = ungroupSel; $('#b-selall').onclick = selectAll;

  /* ---------------- slides ---------------- */
  function layoutShapes(layout) {
    const k = pres.w / 1280, sy = pres.h / 720;
    const T = (ph, x, y, w, h, fs, html, extra) => Object.assign({ id: nextId++, type: 'text', ph, x: Math.round(x * k), y: Math.round(y * sy), w: Math.round(w * k), h: Math.round(h * sy), rot: 0, fill: null, line: null, html, valign: 't', fs, fc: theme().fg, ff: theme().minor }, extra || {});
    const empty = '<p><br></p>';
    const bullets = '<p data-bu="dot" style="margin-left:28px;text-indent:-28px;--ind:28px;margin-bottom:10pt"><br></p>';
    if (layout === 'title') return [T('ctrTitle', 160, 118, 960, 250, 60, '<p style="text-align:center"><br></p>', { valign: 'b', ff: theme().major }), T('subTitle', 160, 378, 960, 174, 28, '<p style="text-align:center"><br></p>', { fc: theme().sub })];
    if (layout === 'titleContent') return [T('title', 88, 38, 1104, 139, 44, empty, { valign: 'ctr', ff: theme().major }), T('body', 88, 192, 1104, 456, 28, bullets)];
    if (layout === 'titleOnly') return [T('title', 88, 38, 1104, 139, 44, empty, { valign: 'ctr', ff: theme().major })];
    return [];
  }
  function hfShapes(layout) {
    const hf = pres.hf;
    if (!hf || (hf.notOnTitle && layout === 'title')) return [];
    const out = [], th = theme(), k = pres.w / 1280, ky = pres.h / 720;
    const base = { rot: 0, fill: null, line: null, valign: 'ctr', fs: 12, fc: th.sub || '7F7F7F', ff: th.minor };
    if (hf.footer) out.push(Object.assign({ id: nextId++, type: 'text', ph: 'ftr', x: Math.round(424 * k), y: Math.round(668 * ky), w: Math.round(432 * k), h: Math.round(38 * ky), html: '<p style="text-align:center">' + esc(hf.footer) + '</p>' }, base));
    if (hf.num) out.push(Object.assign({ id: nextId++, type: 'text', ph: 'sldNum', x: Math.round(904 * k), y: Math.round(668 * ky), w: Math.round(288 * k), h: Math.round(38 * ky), html: '<p style="text-align:right">1</p>' }, base));
    return out;
  }
  function applyHeaderFooter(hf) {
    finishEdit();
    pres.hf = hf && (hf.num || hf.footer) ? hf : undefined;
    pres.slides.forEach(sl => {
      sl.shapes = sl.shapes.filter(s => s.ph !== 'ftr' && s.ph !== 'sldNum');
      sl.shapes.push(...hfShapes(sl.layout));
    });
    selId = null; selSet = [];
    commit();
  }
  async function headerFooterDialog() {
    const hf = pres.hf || { num: false, footer: '', notOnTitle: true };
    const v = await Lumen.form('Header & Footer', [
      { id: 'num', label: 'Slide number', type: 'checkbox', value: !!hf.num },
      { id: 'footer', label: 'Footer text (leave empty for none)', value: hf.footer || '' },
      { id: 'notOnTitle', label: "Don't show on title slides", type: 'checkbox', value: hf.notOnTitle !== false },
    ], 'Apply to all');
    if (v) applyHeaderFooter({ num: !!v.num, footer: String(v.footer || '').trim(), notOnTitle: !!v.notOnTitle });
  }
  $('#i-hf').onclick = headerFooterDialog;
  function setMasterMode(on) {
    finishEdit(); selId = null; selSet = [];
    if (on) pres.master = pres.master || { shapes: [] };
    masterMode = !!on;
    document.body.classList.toggle('mastermode', masterMode);
    $('#masterbar').hidden = !masterMode;
    $('#v-master').classList.toggle('on', masterMode);
    if (!masterMode && pres.master && !pres.master.shapes.length) delete pres.master;
    renderStage(); renderThumbs();
  }
  $('#v-master').onclick = () => setMasterMode(!masterMode);
  $('#m-close').onclick = () => setMasterMode(false);
  $('#bg-nomaster').onchange = e => { if (masterMode) return; slide().noMaster = e.target.checked; commit(); };
  function makeSlide(layout) { return { bg: theme().bg, layout, shapes: layoutShapes(layout).concat(hfShapes(layout)) }; }
  $('#new-layout').onchange = e => {
    if (!e.target.value) return;
    finishEdit();
    pres.slides.splice(cur + 1, 0, makeSlide(e.target.value)); cur++; selId = null; e.target.value = '';
    commit();
  };
  $('#s-dup').onclick = () => { finishEdit(); const c = JSON.parse(JSON.stringify(slide())); c.shapes.forEach(s => { s.id = nextId++; }); pres.slides.splice(cur + 1, 0, c); cur++; selId = null; commit(); };
  $('#s-del').onclick = () => { finishEdit(); if (pres.slides.length < 2) { pres.slides[0] = makeSlide('blank'); } else { pres.slides.splice(cur, 1); cur = Math.min(cur, pres.slides.length - 1); } selId = null; commit(); };

  /* ---------------- insert ---------------- */
  function addShape(s) { s.id = nextId++; slide().shapes.push(s); selId = s.id; selSet = [s.id]; return s; }
  $('#i-text').onclick = () => { finishEdit(); const s = addShape({ type: 'text', x: Math.round(pres.w / 2 - 200), y: Math.round(pres.h / 2 - 30), w: 400, h: 60, rot: 0, fill: null, line: null, html: '<p>Text</p>', valign: 't', fs: 24, fc: theme().fg, ff: theme().minor }); commit(); startEdit(s.id, true); };
  $('#i-shape').onchange = e => {
    const v = e.target.value; e.target.value = '';
    if (!v) return;
    finishEdit();
    const cx = pres.w / 2, cy = pres.h / 2;
    if (v === 'line' || v === 'arrow') addShape({ type: 'line', x: cx - 150, y: cy, w: 300, h: 0, rot: 0, line: { color: '000000', w: 3 }, tail: v === 'arrow' });
    else addShape({ type: v, x: cx - 150, y: cy - 100, w: 300, h: 200, rot: 0, fill: theme().acc[0], line: { color: darken(theme().acc[0]), w: 2 }, html: '<p style="text-align:center"><br></p>', valign: 'ctr', fs: 20, fc: 'FFFFFF', ff: theme().minor });
    commit();
  };
  function normImage(url) {
    return new Promise((res, rej) => {
      const im = new Image();
      im.onload = () => {
        let out = url;
        if (!/^data:image\/(png|jpeg|gif|bmp)/.test(url)) { const c = document.createElement('canvas'); c.width = im.naturalWidth || 300; c.height = im.naturalHeight || 200; c.getContext('2d').drawImage(im, 0, 0); out = c.toDataURL('image/png'); }
        res({ url: out, w: im.naturalWidth, h: im.naturalHeight });
      };
      im.onerror = rej; im.src = url;
    });
  }
  async function insertImage(file) {
    const r = await normImage(await Lumen.readFileAsDataUrl(file));
    const id = 'u' + Date.now().toString(36) + (++mediaN);
    pres.media[id] = r.url;
    let w = r.w, h = r.h;
    const mx = pres.w * 0.6, my = pres.h * 0.7;
    const k = Math.min(1, mx / w, my / h);
    w *= k; h *= k;
    finishEdit();
    addShape({ type: 'image', x: Math.round((pres.w - w) / 2), y: Math.round((pres.h - h) / 2), w: Math.round(w), h: Math.round(h), rot: 0, img: id, crop: null, fill: null, line: null, alt: file.name });
    commit();
  }
  $('#i-image').onclick = () => { const i = document.createElement('input'); i.type = 'file'; i.accept = 'image/*'; i.onchange = () => { if (i.files[0]) insertImage(i.files[0]).catch(() => toast('Could not read that image.')); }; i.click(); };
  $('#i-table').onclick = async () => {
    const v = await form('Insert table', [{ id: 'rows', label: 'Rows', type: 'number', value: 3, min: 1, max: 20 }, { id: 'cols', label: 'Columns', type: 'number', value: 3, min: 1, max: 10 }], 'Insert');
    if (!v) return;
    const R = Math.max(1, v.rows | 0), C = Math.max(1, v.cols | 0);
    const W = Math.min(pres.w - 160, C * 220), cw = W / C;
    const rows = []; for (let r = 0; r < R; r++) rows.push(Array.from({ length: C }, () => ({ html: '<p><br></p>', fill: r === 0 ? 'D9E2F3' : null })));
    finishEdit();
    addShape({ type: 'table', x: Math.round((pres.w - W) / 2), y: 200, w: W, h: R * 48, rot: 0, rows, colW: Array(C).fill(cw), rowH: Array(R).fill(48), fs: 16, fc: '000000', ff: 'Calibri' });
    commit();
  };
  window.addEventListener('paste', async e => {
    if (e.target.closest && (e.target.closest('.tc') || e.target.closest('td') || /INPUT|TEXTAREA/.test(e.target.tagName))) {
      // plain-text paste inside text
      const t = e.clipboardData && e.clipboardData.getData('text/plain');
      if (t && !(e.clipboardData.files && e.clipboardData.files.length)) { e.preventDefault(); document.execCommand('insertText', false, t); }
      return;
    }
    const f = e.clipboardData && [...e.clipboardData.files].find(x => x.type.startsWith('image/'));
    if (f) { e.preventDefault(); insertImage(f); return; }
    if (clipboard) { e.preventDefault(); pasteShapes(Array.isArray(clipboard) ? clipboard : [clipboard], 20); }
  });

  /* ---------------- audio / video ---------------- */
  function posterFor(kind, url) {
    const draw = (frame, w, h) => {
      const c = document.createElement('canvas'); c.width = w; c.height = h;
      const x = c.getContext('2d');
      if (frame) { try { x.drawImage(frame, 0, 0, w, h); } catch (e) { frame = null; } }
      if (!frame) {
        x.fillStyle = kind === 'audio' ? '#E7E6E6' : '#1F1F1F'; x.fillRect(0, 0, w, h);
        x.fillStyle = kind === 'audio' ? '#44546A' : '#FFFFFF';
        if (kind === 'audio') { x.beginPath(); x.moveTo(w * .28, h * .4); x.lineTo(w * .42, h * .4); x.lineTo(w * .58, h * .24); x.lineTo(w * .58, h * .76); x.lineTo(w * .42, h * .6); x.lineTo(w * .28, h * .6); x.closePath(); x.fill(); x.lineWidth = Math.max(3, w * .03); x.strokeStyle = '#44546A'; x.beginPath(); x.arc(w * .58, h * .5, w * .14, -0.8, 0.8); x.stroke(); }
        else { x.beginPath(); x.moveTo(w * .42, h * .34); x.lineTo(w * .62, h * .5); x.lineTo(w * .42, h * .66); x.closePath(); x.fill(); }
      }
      return c.toDataURL('image/png');
    };
    if (kind === 'audio') return Promise.resolve({ url: draw(null, 160, 160), w: 160, h: 160 });
    return new Promise(res => {
      const v = document.createElement('video');
      let done = false;
      const fin = (ok) => { if (done) return; done = true; const w = (ok && v.videoWidth) || 640, h = (ok && v.videoHeight) || 360; res({ url: draw(ok ? v : null, w, h), w, h }); };
      v.muted = true; v.preload = 'auto';
      v.onloadeddata = () => { try { v.currentTime = Math.min(0.1, (v.duration || 1) / 2); } catch (e) { fin(true); } };
      v.onseeked = () => fin(true);
      v.onerror = () => fin(false);
      setTimeout(() => fin(v.readyState >= 2), 2500);
      v.src = url;
    });
  }
  const MEDIA_OK = /^(video\/(mp4|quicktime|webm|x-ms-wmv|x-msvideo)|audio\/(mpeg|mp3|wav|x-wav|wave|mp4|x-m4a|ogg|x-ms-wma))$/;
  async function insertMedia(file, url0) {
    const url = url0 || await Lumen.readFileAsDataUrl(file);
    const mime = (/^data:([^;,]*)/.exec(url) || [0, ''])[1];
    if (!MEDIA_OK.test(mime)) { toast('Unsupported media type' + (mime ? ': ' + mime : '') + '. Use MP4, MOV, MP3, WAV or M4A.', 5000); return null; }
    const kind = mime.startsWith('audio/') ? 'audio' : 'video';
    const po = await posterFor(kind, url);
    const mid = 'v' + Date.now().toString(36) + (++mediaN), pid = 'u' + Date.now().toString(36) + (++mediaN);
    pres.media[mid] = url; pres.media[pid] = po.url;
    let w = po.w, h = po.h;
    const k = Math.min(1, pres.w * 0.6 / w, pres.h * 0.7 / h);
    w = Math.round(w * k); h = Math.round(h * k);
    finishEdit();
    const s = addShape({ type: 'media', kind, x: Math.round((pres.w - w) / 2), y: Math.round((pres.h - h) / 2), w, h, rot: 0, src: mid, img: pid, fill: null, line: null, alt: (file && file.name) || (kind === 'audio' ? 'Audio' : 'Video') });
    commit();
    return s;
  }
  const pickMedia = (accept) => { const i = document.createElement('input'); i.type = 'file'; i.accept = accept; i.onchange = () => { if (i.files[0]) insertMedia(i.files[0]).catch(() => toast('Could not read that file.')); }; i.click(); };
  $('#i-video').onclick = () => pickMedia('video/mp4,video/quicktime,video/webm,.mp4,.mov,.m4v,.wmv,.avi');
  $('#i-audio').onclick = () => pickMedia('audio/*,.mp3,.wav,.m4a');

  /* ---------------- diagrams (built from grouped native shapes) ---------------- */
  const escT = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  function buildDiagram(kind, items) {
    const th = theme(), n = items.length, out = [];
    const gk = 'g' + (nextId++);
    const box = (x, y, w, h, text, ci, type) => out.push({ type: type || 'roundRect', x: Math.round(x), y: Math.round(y), w: Math.round(w), h: Math.round(h), rot: 0, fill: th.acc[ci % th.acc.length], line: null, html: `<p style="text-align:center">${escT(text)}</p>`, valign: 'ctr', fs: 18, fc: 'FFFFFF', ff: th.minor, g: gk });
    const link = (x1, y1, x2, y2, arrow) => out.push({ type: 'line', x: Math.round(Math.min(x1, x2)), y: Math.round(Math.min(y1, y2)), w: Math.round(Math.abs(x2 - x1)), h: Math.round(Math.abs(y2 - y1)), rot: 0, flipH: x2 < x1, flipV: y2 < y1, line: { color: th.sub || '595959', w: 2 }, tail: !!arrow, g: gk });
    const W = pres.w, H = pres.h, mx = W * 0.08;
    if (kind === 'process') {
      const gap = 40, w = (W - 2 * mx - gap * (n - 1)) / n, h = Math.min(160, H * 0.24), y = H / 2 - h / 2 + 30;
      items.forEach((t, i) => { box(mx + i * (w + gap), y, w, h, t, i); if (i < n - 1) link(mx + i * (w + gap) + w + 4, y + h / 2, mx + (i + 1) * (w + gap) - 4, y + h / 2, true); });
    } else if (kind === 'chevron') {
      const w = (W - 2 * mx) / n, h = Math.min(150, H * 0.22), y = H / 2 - h / 2 + 30;
      items.forEach((t, i) => box(mx + i * w, y, w - 8, h, t, i, 'rightArrow'));
    } else if (kind === 'cycle') {
      const cx = W / 2, cy = H / 2 + 40, R = Math.min(W, H) * 0.3, bw = 190, bh = 80;
      const P = items.map((t, i) => { const a = -Math.PI / 2 + i * 2 * Math.PI / n; return [cx + Math.cos(a) * R * 1.25, cy + Math.sin(a) * R * 0.85]; });
      P.forEach((p, i) => { const q = P[(i + 1) % n]; if (n > 1) link(p[0] + (q[0] - p[0]) * 0.28, p[1] + (q[1] - p[1]) * 0.28, p[0] + (q[0] - p[0]) * 0.72, p[1] + (q[1] - p[1]) * 0.72, true); });
      items.forEach((t, i) => box(P[i][0] - bw / 2, P[i][1] - bh / 2, bw, bh, t, i, 'ellipse'));
    } else if (kind === 'hierarchy') {
      const kidsN = Math.max(1, n - 1), bw = Math.min(240, (W - 2 * mx - 30 * (kidsN - 1)) / kidsN), bh = 84, y1 = H * 0.3, y2 = H * 0.62;
      const tw = (bw + 30) * kidsN - 30, x0 = (W - tw) / 2;
      for (let i = 1; i < n; i++) { const cxk = x0 + (i - 1) * (bw + 30) + bw / 2; link(W / 2, y1 + bh, cxk, y2, false); }
      box(W / 2 - 130, y1, 260, bh, items[0], 0);
      for (let i = 1; i < n; i++) box(x0 + (i - 1) * (bw + 30), y2, bw, bh, items[i], 1 + (i - 1) % 5);
    } else if (kind === 'pyramid') {
      const top = H * 0.26, hh = (H * 0.64) / n, maxW = W * 0.6;
      items.forEach((t, i) => { const w = maxW * (i + 1) / n; box(W / 2 - w / 2, top + i * hh, w, hh - 8, t, i, 'rect'); });
    } else {
      const top = H * 0.27, hh = Math.min(96, (H * 0.66) / n);
      items.forEach((t, i) => { box(mx, top + i * hh, 76, hh - 12, String(i + 1), i, 'ellipse'); box(mx + 96, top + i * hh, W - 2 * mx - 96, hh - 12, t, i); });
    }
    return out;
  }
  function insertDiagram(kind, items) {
    finishEdit();
    const list = buildDiagram(kind, items);
    list.forEach(s => { s.id = nextId++; slide().shapes.push(s); });
    selSet = list.map(s => s.id); selId = selSet[selSet.length - 1];
    commit();
  }
  async function diagramDialog(kind) {
    const v = await Lumen.form('Insert diagram', [
      { id: 'kind', label: 'Layout', type: 'select', value: kind, options: [['process', 'Process (boxes and arrows)'], ['chevron', 'Chevron process'], ['cycle', 'Cycle'], ['hierarchy', 'Hierarchy (first line is the top box)'], ['pyramid', 'Pyramid'], ['list', 'Numbered list']] },
      { id: 'items', label: 'Items (one per line)', type: 'textarea', rows: 6, value: 'Plan\nBuild\nTest\nShip' },
    ], 'Insert');
    if (!v) return;
    const items = String(v.items).replace(/\r/g, '').split('\n').map(x => x.trim()).filter(Boolean).slice(0, 8);
    if (!items.length) { toast('Enter at least one item.'); return; }
    insertDiagram(v.kind, items);
  }
  $('#i-diagram').onchange = e => { const v = e.target.value; e.target.value = ''; if (v) diagramDialog(v); };

  /* ---------------- themes ---------------- */
  function applyTheme(key) {
    const nt = THEMES[key];
    if (!nt) return;
    finishEdit();
    const old = theme();
    const up = (c) => String(c || '').toUpperCase();
    pres.slides.forEach(sl => {
      if (!sl.bg || up(sl.bg) === up(old.bg)) sl.bg = nt.bg;
      sl.shapes.forEach(s => {
        const fi = s.fill ? old.acc.findIndex(a => up(a) === up(s.fill)) : -1;
        if (fi >= 0) { s.fill = nt.acc[fi]; if (s.line && s.line.color && (up(s.line.color) === darken(old.acc[fi]) || up(s.line.color) === '2F528F')) s.line.color = darken(nt.acc[fi]); }
        if (s.type === 'line' && s.line) { if (up(s.line.color) === up(old.fg)) s.line.color = nt.fg; else if (up(s.line.color) === up(old.sub)) s.line.color = nt.sub; }
        if (s.ph) {
          const isTitle = s.ph === 'title' || s.ph === 'ctrTitle';
          s.ff = isTitle ? nt.major : nt.minor;
          s.fc = s.ph === 'subTitle' ? nt.sub : nt.fg;
        } else if (s.html !== undefined || s.type === 'table') {
          if (up(s.fc) === up(old.fg) && !s.fill) s.fc = nt.fg;
          if (!s.ff || s.ff === old.minor) s.ff = nt.minor; else if (s.ff === old.major) s.ff = nt.major;
        }
        if (s.type === 'chart' && s.chart && s.chart.series) s.chart.series.forEach(se => { const ci = old.acc.findIndex(a => up(a) === up(se.color)); if (ci >= 0) se.color = nt.acc[ci]; });
      });
    });
    pres.theme = JSON.parse(JSON.stringify(nt));
    commit(); toast('Theme: ' + nt.name);
  }
  $('#th-pick').onchange = e => { const v = e.target.value; e.target.value = ''; if (v) applyTheme(v); };
  $('#s-hide').onclick = () => { slide().hidden = !slide().hidden; commit(); toast(slide().hidden ? 'Slide hidden in slide show' : 'Slide shown in slide show'); };

  /* ---------------- design ---------------- */
  $('#bg').addEventListener('input', e => { $('#sw-bg').style.background = e.target.value; slide().bg = e.target.value.slice(1).toUpperCase(); commit(); });
  $('#apply-bg').onclick = () => { slide().bg = $('#bg').value.slice(1).toUpperCase(); commit(); };
  $('#bg-all').onclick = () => { const b = $('#bg').value.slice(1).toUpperCase(); pres.slides.forEach(s => { s.bg = b; }); commit(); };
  $('#slide-size').onchange = e => {
    const v = e.target.value; e.target.value = '';
    if (!v) return;
    const [w, h] = v.split('x').map(Number);
    const kx = w / pres.w, ky = h / pres.h;
    pres.slides.forEach(sl => sl.shapes.forEach(s => { s.x *= kx; s.y *= ky; s.w *= kx; s.h *= ky; if (s.colW) { s.colW = s.colW.map(c => c * kx); s.rowH = s.rowH.map(c => c * ky); } }));
    pres.w = w; pres.h = h; fitZoom(); commit();
  };

  /* ---------------- slide show ---------------- */
  let showIdx = 0;
  let pvMode = false, aud = null, pvStart = 0, pvTick = 0, showBoxes = [];
  function paintInto(box, W, H, slide, live) {
    const k = Math.max(0.01, Math.min(W / pres.w, H / pres.h));
    box.style.width = pres.w * k + 'px'; box.style.height = pres.h * k + 'px';
    box.innerHTML = '';
    const h = box.ownerDocument.createElement('div');
    h.className = 'slidehost';
    h.style.cssText = `box-shadow:none;width:${pres.w}px;height:${pres.h}px;transform:scale(${k});transform-origin:0 0;position:absolute;left:0;top:0;overflow:hidden`;
    box.appendChild(h);
    renderSlideInto(h, slide, false, live);
    return h;
  }
  const audOpen = () => aud && !aud.closed && aud.document.getElementById('aud');
  function nextVisible(i) { let n = i + 1; while (n < pres.slides.length && pres.slides[n].hidden) n++; return n < pres.slides.length ? n : -1; }
  function showRender(playTrans) {
    const sl = pres.slides[showIdx];
    const targets = [];
    if (!pvMode) targets.push([$('#showslide'), window.innerWidth, window.innerHeight]);
    else {
      const c = $('#pv-main');
      targets.push([$('#pv-cur'), c.clientWidth - 24, c.clientHeight - 24]);
      if (audOpen()) targets.push([aud.document.getElementById('aud'), aud.innerWidth, aud.innerHeight]);
    }
    showBoxes = targets.map(([box, W, H]) => paintInto(box, W, H, sl, true));
    if (pvMode) {
      const nx = nextVisible(showIdx), nb = $('#pv-next');
      if (nx >= 0) paintInto(nb, $('#pv-side').clientWidth - 24, 220, pres.slides[nx], false); else { nb.innerHTML = '<div class="pv-end">End of slide show</div>'; nb.style.width = ''; nb.style.height = ''; }
      $('#pv-notes').textContent = sl.notes || '(no notes for this slide)';
      $('#pv-count').textContent = 'Slide ' + (showIdx + 1) + ' of ' + pres.slides.length;
    }
    showAnims = sl.shapes.filter(s => s.anim);
    showRevealed = 0;
    for (const h of showBoxes) for (const s of showAnims) { if (/^(exit|path|emph)-/.test(String(s.anim))) continue; const el = h.querySelector(`.shape[data-id="${s.id}"]`); if (el) el.style.visibility = 'hidden'; }
    const host = showBoxes[0].parentNode;
    clearTimeout(showTimer);
    if (sl.advAfter > 0) showTimer = setTimeout(() => showStep(1), sl.advAfter * 1000 + (playTrans && sl.trans ? 0 : 0));
    const sn = $('#shownotes');
    sn.textContent = sl.notes || '(no notes)';
    if (playTrans && sl.trans && sl.trans.type !== 'none') {
      const dur = { slow: 1200, med: 700, fast: 350 }[sl.trans.speed] || 700;
      const kf = {
        fade: [{ opacity: 0 }, { opacity: 1 }],
        push: [{ transform: 'translateX(100%)' }, { transform: 'translateX(0)' }],
        cover: [{ transform: 'translateX(100%)' }, { transform: 'translateX(0)' }],
        pull: [{ transform: 'translateX(-100%)' }, { transform: 'translateX(0)' }],
        wipe: [{ clipPath: 'inset(0 100% 0 0)' }, { clipPath: 'inset(0 0 0 0)' }],
        split: [{ clipPath: 'inset(0 50% 0 50%)' }, { clipPath: 'inset(0 0 0 0)' }],
        zoom: [{ transform: 'scale(.3)', opacity: 0 }, { transform: 'scale(1)', opacity: 1 }],
      }[sl.trans.type];
      if (kf) for (const h of showBoxes) { const bx = h.parentNode; if (bx.animate) bx.animate(kf, { duration: dur, easing: 'ease-out' }); }
    }
  }
  let showAnims = [], showRevealed = 0, showTimer = 0;
  function revealNext() {
    const s = showAnims[showRevealed++];
    for (const h of showBoxes) { const el = h.querySelector(`.shape[data-id="${s.id}"]`); if (el) revealEl(s, el); }
  }
  function revealEl(s, el) {
    if (/^(path|emph)-/.test(String(s.anim))) {
      const base = el.style.transform || '';
      let kfp;
      if (s.anim === 'emph-spin') kfp = [{ transform: base + ' rotate(0deg)' }, { transform: base + ' rotate(360deg)' }];
      else if (s.anim === 'emph-grow') kfp = [{ transform: base + ' scale(1)' }, { transform: base + ' scale(1.5)' }];
      else { const d = { 'path-right': [0.25 * pres.w, 0], 'path-left': [-0.25 * pres.w, 0], 'path-up': [0, -0.25 * pres.h], 'path-down': [0, 0.25 * pres.h] }[s.anim] || [0, 0]; kfp = [{ transform: `translate(0px,0px) ${base}` }, { transform: `translate(${d[0]}px,${d[1]}px) ${base}` }]; }
      if (el.animate) el.animate(kfp, { duration: 2000, easing: 'ease-in-out', fill: 'forwards' });
      return;
    }
    if (String(s.anim).startsWith('exit-')) {
      const t = s.anim.slice(5);
      const kfo = { fade: [{ opacity: 1 }, { opacity: 0 }], wipe: [{ clipPath: 'inset(0)' }, { clipPath: 'inset(0 0 0 100%)' }], fly: [{ transform: el.style.transform || 'none', opacity: 1 }, { transform: (el.style.transform || '') + ' translateY(70%)', opacity: 0 }] }[t];
      if (kfo && el.animate) { const an = el.animate(kfo, { duration: 500, easing: 'ease-in', fill: 'forwards' }); an.onfinish = () => { el.style.visibility = 'hidden'; }; } else el.style.visibility = 'hidden';
      return;
    }
    el.style.visibility = 'visible';
    const kf = { fade: [{ opacity: 0 }, { opacity: 1 }], wipe: [{ clipPath: 'inset(0 100% 0 0)' }, { clipPath: 'inset(0)' }], fly: [{ transform: el.style.transform + ' translateY(70%)', opacity: 0 }, { transform: el.style.transform || 'none', opacity: 1 }] }[s.anim];
    if (kf && el.animate) el.animate(kf, { duration: 500, easing: 'ease-out' });
  }
  function startShow(from, presenter) {
    finishEdit();
    while (from < pres.slides.length - 1 && pres.slides[from].hidden) from++;
    pvMode = !!presenter;
    $('#show').classList.toggle('pv', pvMode);
    $('#pv').hidden = !pvMode; $('#showslide').hidden = pvMode;
    showIdx = from; $('#show').hidden = false;
    if (pvMode) {
      pvStart = Date.now(); clearInterval(pvTick);
      const tick = () => { const s = Math.floor((Date.now() - pvStart) / 1000); $('#pv-timer').textContent = Math.floor(s / 60) + ':' + String(s % 60).padStart(2, '0'); $('#pv-clock').textContent = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }); };
      tick(); pvTick = setInterval(tick, 1000);
      showRender(false);
      return;
    }
    showRender(false);
    const el = $('#show');
    if (el.requestFullscreen) el.requestFullscreen().catch(() => { });
  }
  function endShow() {
    clearTimeout(showTimer); clearInterval(pvTick);
    $('#show').hidden = true;
    if (audOpen()) { try { aud.close(); } catch (e) { } }
    aud = null;
    if (document.fullscreenElement) document.exitFullscreen().catch(() => { });
    gotoSlide(showIdx);
  }
  // Audience window: drag it to the projector / second screen (or it is placed there automatically when Windows reports one), press F11 for full screen
  async function openAudience() {
    let feat = 'popup=yes,width=960,height=540';
    try {
      if ('getScreenDetails' in window) {
        const sd = await window.getScreenDetails();
        const other = sd.screens.find(sc => sc !== sd.currentScreen);
        if (other) feat = `popup=yes,left=${other.availLeft},top=${other.availTop},width=${other.availWidth},height=${other.availHeight}`;
      }
    } catch (e) { }
    aud = window.open('', 'lumen-audience', feat);
    if (!aud) { toast('The browser blocked the audience window. Allow pop-ups for this page and try again.', 6000); return; }
    const css = ['../shared/common.css', 'slides.css'].map(u => new URL(u, location.href).href);
    aud.document.open();
    aud.document.write('<!doctype html><html><head><meta charset="utf-8"><title>' + esc(($('#title').value || 'Presentation') + ' - Slide Show') + '</title>' + css.map(h => '<link rel="stylesheet" href="' + h + '">').join('') +
      '<style>html,body{margin:0;height:100%;background:#000;overflow:hidden;cursor:none}body{display:grid;place-items:center}#aud{position:relative;overflow:hidden}</style></head><body><div id="aud"></div></body></html>');
    aud.document.close();
    aud.document.addEventListener('click', e => { if (e.target.closest && e.target.closest('video,audio')) return; showStep(1); });
    aud.document.addEventListener('keydown', showKey);
    aud.addEventListener('resize', () => showRender(false));
    aud.addEventListener('load', () => showRender(false));
    setTimeout(() => showRender(false), 150);
  }
  $('#pv-aud').onclick = e => { e.stopPropagation(); openAudience(); };
  $('#pv-prev').onclick = e => { e.stopPropagation(); showStep(-1); };
  $('#pv-nextb').onclick = e => { e.stopPropagation(); showStep(1); };
  $('#pv-end').onclick = e => { e.stopPropagation(); endShow(); };
  $('#pv-reset').onclick = e => { e.stopPropagation(); pvStart = Date.now(); };
  $('#v-presenter').onclick = () => startShow(cur, true);
  const showStep = (d) => {
    if (d > 0 && showRevealed < showAnims.length) { revealNext(); return; }
    let n = showIdx + d;
    while (n >= 0 && n < pres.slides.length && pres.slides[n].hidden) n += d;
    if (n < 0) return;
    if (n >= pres.slides.length) { endShow(); return; }
    showIdx = n;
    showRender(d > 0);
  };
  $('#show').addEventListener('click', e => {
    if (e.target.closest && e.target.closest('video,audio')) return;
    if (pvMode && !e.target.closest('#pv-main')) return;
    showStep(1);
  });
  window.addEventListener('resize', () => { if (!$('#show').hidden) showRender(); else fitZoom(); });
  window.addEventListener('beforeunload', () => { if (audOpen()) aud.close(); });
  document.addEventListener('fullscreenchange', () => { if (!document.fullscreenElement && !$('#show').hidden) endShow(); });
  $('#btn-show').onclick = () => startShow(0); $('#v-show').onclick = () => startShow(0); $('#v-showcur').onclick = () => startShow(cur);
  $('#v-fit').onclick = fitZoom;
  $('#tr-preview').onclick = () => startShow(cur);

  /* ---------------- keyboard ---------------- */
  document.addEventListener('keydown', e => {
    if (!$('#show').hidden) { showKey(e); return; }
    mainKey(e);
  });
  function showKey(e) {
    {
      if (e.key === 'Escape') endShow();
      else if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter', 'n'].includes(e.key)) { e.preventDefault(); showStep(1); }
      else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p'].includes(e.key)) { e.preventDefault(); showStep(-1); }
      else if (e.key === 'm' || e.key === 'M') { const sn = $('#shownotes'); sn.hidden = !sn.hidden; }
      else if (e.key === 'Home') { showIdx = 0; showRender(); } else if (e.key === 'End') { showIdx = pres.slides.length - 1; showRender(); }
      else if (e.key === 'b' || e.key === 'B' || e.key === '.') { for (const h of showBoxes) h.style.visibility = h.style.visibility === 'hidden' ? '' : 'hidden'; }
      return;
    }
  }
  function mainKey(e) {
    const ctrl = e.ctrlKey || e.metaKey, k = e.key.toLowerCase();
    if (ctrl && k === 's') { e.preventDefault(); save(e.shiftKey); return; }
    if (ctrl && k === 'o') { e.preventDefault(); openDialog(); return; }
    const inText = e.target.closest && (e.target.closest('.tc') || e.target.closest('td') || /INPUT|TEXTAREA|SELECT/.test(e.target.tagName));
    if (inText) { if (ctrl && k === 'z' && !e.target.closest('.tc') && !e.target.closest('td')) return; return; }
    if (ctrl && k === 'z') { e.preventDefault(); e.shiftKey ? redo() : undo(); return; }
    if (ctrl && k === 'y') { e.preventDefault(); redo(); return; }
    if (ctrl && k === 'd') { e.preventDefault(); dupShape(); return; }
    if (ctrl && k === 'c') { const ss = selShapes(); if (ss.length) clipboard = JSON.parse(JSON.stringify(ss)); return; }
    if (ctrl && k === 'x') { const ss = selShapes(); if (ss.length) { clipboard = JSON.parse(JSON.stringify(ss)); delShape(); } return; }
    if (ctrl && k === 'g') { e.preventDefault(); if (e.shiftKey) ungroupSel(); else groupSel(); return; }
    if (ctrl && k === 'a') { e.preventDefault(); selectAll(); return; }
    if (e.key === 'F5') { e.preventDefault(); startShow(e.shiftKey || e.altKey ? cur : 0, e.altKey); return; }
    const s = shapeOf(selId);
    if (s) {
      if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); delShape(); return; }
      if (e.key === 'Enter' || e.key === 'F2') { if (['image', 'line', 'table', 'media', 'chart'].indexOf(s.type) < 0) { e.preventDefault(); startEdit(s.id); } return; }
      const d = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[e.key];
      if (d) { e.preventDefault(); const st = e.shiftKey ? 10 : 1; selShapes().forEach(q => { q.x += d[0] * st; q.y += d[1] * st; }); commit(); return; }
      if (e.key === 'Escape') { select(null); return; }
    } else {
      if (e.key === 'PageDown' || e.key === 'ArrowDown' || e.key === 'ArrowRight') { gotoSlide(cur + 1); e.preventDefault(); }
      if (e.key === 'PageUp' || e.key === 'ArrowUp' || e.key === 'ArrowLeft') { gotoSlide(cur - 1); e.preventDefault(); }
      if ((e.key === 'Delete') && pres.slides.length && !masterMode) { $('#s-del').click(); }
    }
  }

  /* ---------------- files ---------------- */
  const titleEl = $('#title');
  const setTitle = (n) => { titleEl.value = n; document.title = n + ' - Lumen Slides'; };
  titleEl.addEventListener('input', () => { document.title = titleEl.value + ' - Lumen Slides'; markDirty(); });
  function setPres(p) {
    pres = p; if (!pres.slides.length) pres.slides.push(makeSlide('title'));
    let mx = 0; pres.slides.forEach(sl => sl.shapes.forEach(s => { mx = Math.max(mx, s.id || 0); })); nextId = mx + 1;
    if (pres.master) pres.master.shapes.forEach(s => { s.id = nextId++; });
    masterMode = false; document.body.classList.remove('mastermode');
    cur = 0; selId = null; selSet = []; editingId = null;
    hist.stack = []; hist.i = -1;
    fitZoom(); renderStage();
    hist.stack.push(snap()); hist.i = 0;
    renderThumbs();
  }
  async function newPres() {
    if (state.dirty && !(await Lumen.confirmBox('Discard changes?', 'This presentation has unsaved changes. Create a new one anyway?', 'Discard'))) return;
    setPres({ w: 1280, h: 720, media: {}, slides: [makeSlide('title')] }); state.handle = null; setTitle('Presentation1'); markClean();
  }
  async function openDialog() {
    if (state.dirty && !(await Lumen.confirmBox('Discard changes?', 'This presentation has unsaved changes. Open another file anyway?', 'Discard'))) return;
    const r = await FileIO.open([{ description: 'Presentations', accept: { [PptxIO.PPTX_MIME]: ['.pptx'] } }], '.pptx');
    if (r) await loadFile(r.file, r.handle);
  }
  async function loadFile(file, handle) {
    try {
      const ext = (file.name.match(/\.([^.]+)$/) || [, ''])[1].toLowerCase();
      if (ext === 'ppt') { toast('Old binary .ppt files are not supported. Save as .pptx first.', 6000); return; }
      if (!['pptx', 'pptm', 'potx', 'ppsx'].includes(ext)) { toast('Unsupported file type: .' + ext); return; }
      const p = await PptxIO.read(await file.arrayBuffer());
      setPres(p); state.handle = handle; setTitle(baseName(file.name)); markClean();
      toast(p.lost ? `Opened ${file.name}. ${p.lost} chart/diagram object(s) are shown as placeholders and will not be preserved on save.` : 'Opened ' + file.name, p.lost ? 7000 : 2600);
    } catch (err) { console.error(err); toast('Could not open file: ' + err.message, 6000); }
  }
  async function save(saveAs) {
    try {
      finishEdit();
      const blob = await PptxIO.write(pres);
      const r = await FileIO.save(blob, (titleEl.value || 'Presentation1') + '.pptx', FILE_TYPES, state.handle, saveAs);
      if (r) { state.handle = r.handle; if (r.handle) setTitle(baseName(r.name)); markClean(); toast('Saved ' + r.name); }
    } catch (err) { console.error(err); toast('Save failed: ' + err.message, 6000); }
  }
  $('#btn-new').onclick = newPres; $('#btn-open').onclick = openDialog; $('#btn-save').onclick = () => save(false); $('#btn-saveas').onclick = () => save(true);
  $('#b-undo').onclick = undo; $('#b-redo').onclick = redo;
  if (FileIO.onLaunch) FileIO.onLaunch((f, h) => loadFile(f, h));
  FileIO.onDrop(f => { if (f.type.startsWith('image/')) insertImage(f); else if (/^(video|audio)\//.test(f.type)) insertMedia(f); else loadFile(f, null); });
  window.addEventListener('beforeunload', e => { if (state.dirty) { e.preventDefault(); e.returnValue = ''; } });
  new ResizeObserver(() => { if ($('#show').hidden) fitZoom(); }).observe(stage);

  Lumen.initRibbon();
  setPres({ w: 1280, h: 720, media: {}, slides: [makeSlide('title')] });
  setTitle('Presentation1'); markClean();
  window.LumenSlides = { get pres() { return pres; }, setPres, loadFile, save, commit, makeSlide, renderSlideInto, get state() { return state; }, applyHeaderFooter, startShow, endShow, showStep, setMasterMode, get masterMode() { return masterMode; }, get showIdx() { return showIdx; }, select, selShapes, groupSel, ungroupSel, selectAll, addShape, applyTheme, insertMedia, insertDiagram, buildDiagram, dupShape, delShape, gotoSlide, THEMES, get cur() { return cur; }, get selId() { return selId; } };
})();
