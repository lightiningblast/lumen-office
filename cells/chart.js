/* Charts for Lumen Cells: data lookup, canvas drawing, and DrawingML chart XML (read + write).
   chart = { id, type: 'col'|'bar'|'line'|'area'|'pie', title, legend, x, y, w, h,
             series: [ { tx: 'Sheet1!$B$1', cat: 'Sheet1!$A$2:$A$5', val: 'Sheet1!$B$2:$B$5', color } ] } */
(function (g) {
  'use strict';
  const { xesc } = Lumen;
  const PALETTE = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '264478', '9E480E'];
  const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

  /* ---------- references ---------- */
  function parseRef(wb, str, defSheet) {
    if (!str) return null;
    const m = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(str.trim());
    let sh = defSheet, body = str;
    if (m) { sh = wb.sheetByName((m[1] || m[2]).replace(/''/g, "'")); body = m[3]; }
    const rg = Formula.parseRange(body);
    if (!sh || !rg) return null;
    return { sh, ...rg };
  }
  function refValues(wb, str, defSheet) {
    const r = parseRef(wb, str, defSheet);
    if (!r) return [];
    const out = [];
    for (let rr = r.r1; rr <= r.r2; rr++) for (let c = r.c1; c <= r.c2; c++) {
      const cell = r.sh.get(rr, c);
      let v = cell ? (cell.f != null ? Formula.cellValue(wb, r.sh, rr, c) : cell.v) : null;
      out.push({ v, nf: wb.styles[(cell && cell.s) || 0].nf });
    }
    return out;
  }
  function chartData(wb, ch, defSheet) {
    const series = ch.series.map((s, i) => {
      const vals = refValues(wb, s.val, defSheet).map(x => (typeof x.v === 'number' ? x.v : (typeof x.v === 'string' && x.v.trim() !== '' && !isNaN(Number(x.v)) ? Number(x.v) : null)));
      let name = '';
      if (s.tx) { const t = refValues(wb, s.tx, defSheet)[0]; name = t && t.v != null ? String(t.v) : ''; }
      return { name: name || 'Series' + (i + 1), color: s.color || PALETTE[i % PALETTE.length], vals };
    });
    let cats = [];
    const first = ch.series.find(s => s.cat);
    if (first) cats = refValues(wb, first.cat, defSheet).map(x => x.v === null ? '' : (typeof x.v === 'number' ? NumFmt.format(x.v, x.nf) : String(x.v)));
    const n = Math.max(0, ...series.map(s => s.vals.length), cats.length);
    for (let i = cats.length; i < n; i++) cats.push(String(i + 1));
    return { cats, series };
  }

  /* ---------- drawing ---------- */
  function niceScale(min, max) {
    if (min === max) { if (max === 0) max = 1; else { min = Math.min(0, min); max = Math.max(0, max); } }
    const range = max - min;
    let step = Math.pow(10, Math.floor(Math.log10(range / 5)));
    const err = range / 5 / step;
    step *= err >= 7.5 ? 10 : err >= 3.5 ? 5 : err >= 1.5 ? 2 : 1;
    return { min: Math.floor(min / step) * step, max: Math.ceil(max / step) * step, step };
  }
  const fmtNum = (v) => { const a = Math.abs(v); if (a >= 1e9) return +(v / 1e9).toPrecision(3) + 'B'; if (a >= 1e6) return +(v / 1e6).toPrecision(3) + 'M'; return String(+v.toPrecision(6)); };

  function draw(ctx, W, H, ch, data) {
    ctx.save();
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, H);
    ctx.strokeStyle = '#d9d9d9'; ctx.lineWidth = 1; ctx.strokeRect(.5, .5, W - 1, H - 1);
    ctx.font = '12px "Segoe UI", Calibri, sans-serif';
    ctx.textBaseline = 'middle';
    let top = 10, left = 10, right = W - 10, bottom = H - 10;
    if (ch.title) { ctx.font = '16px "Segoe UI", Calibri, sans-serif'; ctx.fillStyle = '#404040'; ctx.textAlign = 'center'; ctx.fillText(ch.title, W / 2, 22, W - 20); top = 40; ctx.font = '12px "Segoe UI", Calibri, sans-serif'; }
    const isPie = ch.type === 'pie';
    const legendItems = isPie ? data.cats.map((c, i) => ({ name: c, color: PALETTE[i % PALETTE.length] })) : data.series.map(s => ({ name: s.name, color: s.color }));
    if (ch.legend !== false && legendItems.length && !(data.series.length === 1 && !isPie && false)) {
      let lw = 0;
      for (const l of legendItems) lw = Math.max(lw, ctx.measureText(l.name).width);
      lw = Math.min(lw + 28, W * 0.3);
      right = W - lw - 10;
      let y = Math.max(top + 10, H / 2 - legendItems.length * 9);
      ctx.textAlign = 'left'; ctx.fillStyle = '#595959';
      for (const l of legendItems.slice(0, 30)) { ctx.fillStyle = '#' + l.color; ctx.fillRect(right + 12, y - 5, 10, 10); ctx.fillStyle = '#595959'; ctx.fillText(l.name, right + 27, y, lw - 22); y += 18; }
    }
    if (!data.series.length || !data.cats.length) { ctx.fillStyle = '#999'; ctx.textAlign = 'center'; ctx.fillText('No data', W / 2, H / 2); ctx.restore(); return; }
    if (isPie) {
      const vals = (data.series[0].vals).map(v => Math.max(0, v || 0));
      const total = vals.reduce((a, b) => a + b, 0) || 1;
      const cx = (left + right) / 2, cy = (top + bottom) / 2, r = Math.max(10, Math.min(right - left, bottom - top) / 2 - 6);
      let a0 = -Math.PI / 2;
      vals.forEach((v, i) => {
        const a1 = a0 + v / total * Math.PI * 2;
        ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, r, a0, a1); ctx.closePath();
        ctx.fillStyle = '#' + PALETTE[i % PALETTE.length]; ctx.fill(); ctx.strokeStyle = '#fff'; ctx.lineWidth = 1.5; ctx.stroke();
        if (v / total > 0.05) { const am = (a0 + a1) / 2; ctx.fillStyle = '#fff'; ctx.textAlign = 'center'; ctx.fillText(Math.round(v / total * 100) + '%', cx + Math.cos(am) * r * 0.65, cy + Math.sin(am) * r * 0.65); }
        a0 = a1;
      });
      ctx.restore(); return;
    }
    const all = data.series.flatMap(s => s.vals).filter(v => v !== null);
    let mn = Math.min(0, ...all), mx = Math.max(0, ...all);
    if (!all.length) { mn = 0; mx = 1; }
    const sc = niceScale(mn, mx);
    const horiz = ch.type === 'bar';
    const labelW = Math.max(...[sc.min, sc.max].map(v => ctx.measureText(fmtNum(v)).width)) + 8;
    const catH = 22;
    let px0, px1, py0, py1;
    if (horiz) {
      const catW = Math.min(110, Math.max(...data.cats.map(c => ctx.measureText(c).width)) + 10);
      px0 = left + catW; px1 = right - 8; py0 = top; py1 = bottom - 20;
    } else { px0 = left + labelW; px1 = right - 8; py0 = top + 6; py1 = bottom - catH; }
    const nc = data.cats.length;
    const toV = (v) => horiz ? px0 + (v - sc.min) / (sc.max - sc.min) * (px1 - px0) : py1 - (v - sc.min) / (sc.max - sc.min) * (py1 - py0);
    // grid + value axis labels
    ctx.lineWidth = 1;
    for (let v = sc.min; v <= sc.max + sc.step / 2; v += sc.step) {
      const p = toV(v);
      ctx.strokeStyle = '#e6e6e6'; ctx.beginPath();
      if (horiz) { ctx.moveTo(p, py0); ctx.lineTo(p, py1); } else { ctx.moveTo(px0, p); ctx.lineTo(px1, p); }
      ctx.stroke();
      ctx.fillStyle = '#595959';
      if (horiz) { ctx.textAlign = 'center'; ctx.fillText(fmtNum(v), p, py1 + 12); } else { ctx.textAlign = 'right'; ctx.fillText(fmtNum(v), px0 - 6, p); }
    }
    ctx.strokeStyle = '#bfbfbf'; ctx.beginPath();
    const zero = toV(0);
    if (horiz) { ctx.moveTo(zero, py0); ctx.lineTo(zero, py1); } else { ctx.moveTo(px0, zero); ctx.lineTo(px1, zero); }
    ctx.stroke();
    const slot = (horiz ? (py1 - py0) : (px1 - px0)) / nc;
    // category labels
    ctx.fillStyle = '#595959';
    const every = Math.max(1, Math.ceil(ctx.measureText('MMMMMM').width / Math.max(slot, 1)));
    data.cats.forEach((c, i) => {
      if (i % every) return;
      if (horiz) { ctx.textAlign = 'right'; ctx.fillText(c, px0 - 6, py0 + slot * (i + 0.5), 104); }
      else { ctx.textAlign = 'center'; ctx.fillText(c, px0 + slot * (i + 0.5), py1 + 14, Math.max(30, slot * every - 4)); }
    });
    const ns = data.series.length;
    if (ch.type === 'col' || ch.type === 'bar') {
      const gw = slot * 0.8, bw = gw / ns;
      data.series.forEach((s, si) => {
        ctx.fillStyle = '#' + s.color;
        s.vals.forEach((v, i) => {
          if (v === null) return;
          const off = slot * 0.1 + si * bw + slot * i;
          if (horiz) { const a = zero, b = toV(v); ctx.fillRect(Math.min(a, b), py0 + off, Math.abs(b - a), bw - 1); }
          else { const a = zero, b = toV(v); ctx.fillRect(px0 + off, Math.min(a, b), bw - 1, Math.abs(b - a)); }
        });
      });
    } else {
      data.series.forEach(s => {
        const pts = s.vals.map((v, i) => v === null ? null : [px0 + slot * (i + 0.5), toV(v)]);
        if (ch.type === 'area') {
          ctx.beginPath(); let started = false, lastX = 0;
          pts.forEach(p => { if (!p) return; if (!started) { ctx.moveTo(p[0], zero); ctx.lineTo(p[0], p[1]); started = true; } else ctx.lineTo(p[0], p[1]); lastX = p[0]; });
          if (started) { ctx.lineTo(lastX, zero); ctx.closePath(); ctx.globalAlpha = 0.6; ctx.fillStyle = '#' + s.color; ctx.fill(); ctx.globalAlpha = 1; }
        }
        ctx.strokeStyle = '#' + s.color; ctx.lineWidth = 2.25; ctx.beginPath(); let st = false;
        pts.forEach(p => { if (!p) { st = false; return; } if (!st) { ctx.moveTo(p[0], p[1]); st = true; } else ctx.lineTo(p[0], p[1]); });
        ctx.stroke();
        if (ch.type === 'line') { ctx.fillStyle = '#' + s.color; pts.forEach(p => { if (p) { ctx.beginPath(); ctx.arc(p[0], p[1], 3.5, 0, Math.PI * 2); ctx.fill(); } }); }
      });
    }
    ctx.restore();
  }

  /* ---------- chart XML (write) ---------- */
  const q = (n) => Formula.quoteSheet(n);
  function strCache(vals) { return `<c:strCache><c:ptCount val="${vals.length}"/>` + vals.map((v, i) => `<c:pt idx="${i}"><c:v>${xesc(v)}</c:v></c:pt>`).join('') + '</c:strCache>'; }
  function numCache(vals) { return `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>` + vals.map((v, i) => v === null ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('') + '</c:numCache>'; }
  function chartXml(ch, data) {
    const isPie = ch.type === 'pie', isLine = ch.type === 'line', isArea = ch.type === 'area';
    const sers = ch.series.map((s, i) => {
      const d = data.series[i];
      const color = d.color;
      const sp = isLine ? `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>` : (isPie ? '' : `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>${isArea ? '' : '<c:invertIfNegative val="0"/>'}`);
      let dpt = '';
      if (isPie) dpt = data.cats.map((_, k) => `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${PALETTE[k % PALETTE.length]}"/></a:solidFill></c:spPr></c:dPt>`).join('');
      if (ch.lit) {
        const strLit = `<c:strLit><c:ptCount val="${data.cats.length}"/>` + data.cats.map((v, k) => `<c:pt idx="${k}"><c:v>${xesc(v)}</c:v></c:pt>`).join('') + '</c:strLit>';
        const numLit = `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${d.vals.length}"/>` + d.vals.map((v, k) => v === null ? '' : `<c:pt idx="${k}"><c:v>${v}</c:v></c:pt>`).join('') + '</c:numLit>';
        return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:v>${xesc(d.name)}</c:v></c:tx>${sp}${dpt}<c:cat>${strLit}</c:cat><c:val>${numLit}</c:val>${isLine ? '<c:smooth val="0"/>' : ''}</c:ser>`;
      }
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>` +
        (s.tx ? `<c:tx><c:strRef><c:f>${xesc(s.tx)}</c:f>${strCache([d.name])}</c:strRef></c:tx>` : `<c:tx><c:v>${xesc(d.name)}</c:v></c:tx>`) + sp + dpt +
        (s.cat ? `<c:cat><c:strRef><c:f>${xesc(s.cat)}</c:f>${strCache(data.cats)}</c:strRef></c:cat>` : '') +
        `<c:val><c:numRef><c:f>${xesc(s.val)}</c:f>${numCache(d.vals)}</c:numRef></c:val>${isLine ? '<c:smooth val="0"/>' : ''}</c:ser>`;
    }).join('');
    let plot;
    const axIds = '<c:axId val="111111111"/><c:axId val="222222222"/>';
    if (isPie) plot = `<c:pieChart><c:varyColors val="1"/>${sers}<c:firstSliceAng val="0"/></c:pieChart>`;
    else if (isLine) plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}<c:marker val="1"/>${axIds}</c:lineChart>`;
    else if (isArea) plot = `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}${axIds}</c:areaChart>`;
    else plot = `<c:barChart><c:barDir val="${ch.type === 'bar' ? 'bar' : 'col'}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${sers}<c:gapWidth val="100"/>${axIds}</c:barChart>`;
    let axes = '';
    if (!isPie) {
      const h = ch.type === 'bar';
      axes = `<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="${h ? 'maxMin' : 'minMax'}"/></c:scaling><c:delete val="0"/><c:axPos val="${h ? 'l' : 'b'}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="222222222"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
        `<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${h ? 'b' : 'l'}"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="111111111"/><c:crosses val="${h ? 'max' : 'autoZero'}"/><c:crossBetween val="between"/></c:valAx>`;
    }
    const title = ch.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0"/></a:pPr><a:r><a:rPr lang="en-US" sz="1400" b="0"/><a:t>${xesc(ch.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
    return XML + `<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${plot}${axes}</c:plotArea>${ch.legend === false ? '' : '<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>'}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
  }

  /* ---------- chart XML (read) ---------- */
  const kid = (el, ln) => { for (let c = el && el.firstElementChild; c; c = c.nextElementSibling) if (c.localName === ln) return c; return null; };
  const kids = (el, ln) => { const r = []; for (let c = el && el.firstElementChild; c; c = c.nextElementSibling) if (c.localName === ln) r.push(c); return r; };
  const fOf = (el) => { const f = el && el.getElementsByTagNameNS('*', 'f')[0]; return f ? f.textContent : ''; };
  function parseChart(root) {
    const chart = kid(root, 'chart');
    const plot = chart && kid(chart, 'plotArea');
    if (!plot) return null;
    let type = null, node = null;
    for (const c of plot.children) {
      const n = c.localName;
      if (n === 'barChart') { type = (kid(c, 'barDir') && kid(c, 'barDir').getAttribute('val') === 'bar') ? 'bar' : 'col'; node = c; break; }
      if (n === 'lineChart' || n === 'line3DChart') { type = 'line'; node = c; break; }
      if (n === 'areaChart') { type = 'area'; node = c; break; }
      if (n === 'pieChart' || n === 'doughnutChart' || n === 'pie3DChart') { type = 'pie'; node = c; break; }
      if (n === 'scatterChart') { type = 'line'; node = c; break; }
    }
    if (!node) return null;
    const series = [];
    for (const s of kids(node, 'ser')) {
      const val = fOf(kid(s, 'val')) || fOf(kid(s, 'yVal'));
      if (!val) continue;
      const sp = kid(s, 'spPr');
      let color = null;
      const sf = sp && (kid(sp, 'solidFill') || (kid(sp, 'ln') && kid(kid(sp, 'ln'), 'solidFill')));
      const clr = sf && kid(sf, 'srgbClr');
      if (clr) color = clr.getAttribute('val').toUpperCase();
      series.push({ tx: fOf(kid(s, 'tx')), cat: fOf(kid(s, 'cat')) || fOf(kid(s, 'xVal')), val, color });
    }
    let title = '';
    const t = kid(chart, 'title');
    if (t) title = [...t.getElementsByTagNameNS('*', 't')].map(x => x.textContent).join('');
    return { type, title, legend: !!kid(chart, 'legend'), series };
  }

  // Read a chart part with cached/literal data (used by PowerPoint): { type, title, legend, cats[], series:[{name, vals[], color}] }
  function parseChartLit(root) {
    const base = parseChart(root);
    const chart = kid(root, 'chart');
    const plot = chart && kid(chart, 'plotArea');
    if (!plot) return null;
    let node = null, type = null;
    for (const c of plot.children) {
      const n = c.localName;
      if (/^(bar|line|area|pie|doughnut|scatter|bar3D|pie3D|line3D)Chart$/.test(n)) {
        node = c;
        type = n === 'barChart' || n === 'bar3DChart' ? ((kid(c, 'barDir') && kid(c, 'barDir').getAttribute('val') === 'bar') ? 'bar' : 'col') : /^line/.test(n) || n === 'scatterChart' ? 'line' : n === 'areaChart' ? 'area' : 'pie';
        break;
      }
    }
    if (!node) return null;
    const ptsOf = (el) => { if (!el) return []; const pc = el.getElementsByTagNameNS('*', 'ptCount')[0]; const n = pc ? parseInt(pc.getAttribute('val'), 10) : 0; const out = new Array(n).fill(null); for (const p of el.getElementsByTagNameNS('*', 'pt')) { const v = p.getElementsByTagNameNS('*', 'v')[0]; out[parseInt(p.getAttribute('idx'), 10)] = v ? v.textContent : null; } return out; };
    const series = []; let cats = [];
    for (const s of kids(node, 'ser')) {
      const val = kid(s, 'val') || kid(s, 'yVal');
      if (!val) continue;
      const sp = kid(s, 'spPr');
      const sf = sp && (kid(sp, 'solidFill') || (kid(sp, 'ln') && kid(kid(sp, 'ln'), 'solidFill')));
      const clr = sf && kid(sf, 'srgbClr');
      const c = ptsOf(kid(s, 'cat') || kid(s, 'xVal'));
      if (c.length > cats.length) cats = c.map(x => x === null ? '' : x);
      const tx = kid(s, 'tx');
      const nm = tx ? ([...tx.getElementsByTagNameNS('*', 'v')].map(v => v.textContent).join('') || '') : '';
      series.push({ name: nm, vals: ptsOf(val).map(v => v === null || v === '' || isNaN(Number(v)) ? null : Number(v)), color: clr ? clr.getAttribute('val').toUpperCase() : PALETTE[series.length % PALETTE.length] });
    }
    if (!series.length) return null;
    return { type, title: base ? base.title : '', legend: base ? base.legend : true, cats, series };
  }

  g.Charts = { parseChartLit, PALETTE, chartData, draw, chartXml, parseChart, parseRef, refValues, NS_C };
})(window);
