/* XLSX (SpreadsheetML) reader / writer for Lumen Cells. */
(function (g) {
  'use strict';
  const { xesc } = Lumen;
  const { Workbook, Sheet, XErr, ERR_BY_CODE, MAXC, colName, colIndex, cellName, shiftFormula } = Formula;
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const NS_M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const REL = NS_R;
  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

  const kids = (el, ln) => { const r = []; for (let c = el && el.firstElementChild; c; c = c.nextElementSibling) if (c.localName === ln) r.push(c); return r; };
  const kid = (el, ln) => { for (let c = el && el.firstElementChild; c; c = c.nextElementSibling) if (c.localName === ln) return c; return null; };
  function parseXml(s) {
    const d = new DOMParser().parseFromString(s, 'application/xml');
    if (d.getElementsByTagName('parsererror').length) throw new Error('Could not parse XML in the workbook.');
    return d.documentElement;
  }
  function dirname(p) { const i = p.lastIndexOf('/'); return i < 0 ? '' : p.slice(0, i); }
  function resolvePath(base, t) {
    if (t.startsWith('/')) return t.slice(1);
    const parts = base ? base.split('/') : [];
    for (const s of t.split('/')) { if (s === '..') parts.pop(); else if (s !== '.') parts.push(s); }
    return parts.join('/');
  }
  function parseRels(xml, base) {
    const out = {};
    if (!xml) return out;
    for (const r of parseXml(xml).children) out[r.getAttribute('Id')] = { type: r.getAttribute('Type'), path: resolvePath(base, r.getAttribute('Target')) };
    return out;
  }

  /* ---------- colour helpers ---------- */
  const INDEXED = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];
  function applyTint(hex, tint) {
    if (!tint) return hex;
    let r = parseInt(hex.slice(0, 2), 16) / 255, g2 = parseInt(hex.slice(2, 4), 16) / 255, b = parseInt(hex.slice(4, 6), 16) / 255;
    const mx = Math.max(r, g2, b), mn = Math.min(r, g2, b);
    let h = 0, s = 0, l = (mx + mn) / 2;
    if (mx !== mn) {
      const d = mx - mn;
      s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
      h = mx === r ? (g2 - b) / d + (g2 < b ? 6 : 0) : mx === g2 ? (b - r) / d + 2 : (r - g2) / d + 4;
      h /= 6;
    }
    l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
    const hue = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; };
    let rr, gg, bb;
    if (s === 0) rr = gg = bb = l;
    else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; rr = hue(p, q, h + 1 / 3); gg = hue(p, q, h); bb = hue(p, q, h - 1 / 3); }
    return [rr, gg, bb].map(x => ('0' + Math.round(x * 255).toString(16)).slice(-2)).join('').toUpperCase();
  }

  /* ====================================================================
     READ
     ==================================================================== */
  async function read(buf) {
    const z = await Zip.read(buf);
    let wbPath = 'xl/workbook.xml';
    if (z.has('_rels/.rels')) {
      const r = parseRels(await z.text('_rels/.rels'), '');
      for (const k in r) if (/officeDocument$/.test(r[k].type)) wbPath = r[k].path;
    }
    if (!z.has(wbPath)) throw new Error('This file does not contain a workbook (xl/workbook.xml not found).');
    const base = dirname(wbPath);
    const rels = parseRels(await z.textOrNull(base + '/_rels/' + wbPath.split('/').pop() + '.rels'), base);
    const wbRoot = parseXml(await z.text(wbPath));

    // theme colours
    const themeColors = [];
    for (const k in rels) if (/\/theme$/.test(rels[k].type) && z.has(rels[k].path)) {
      const t = parseXml(await z.text(rels[k].path));
      const cs = t.getElementsByTagNameNS('*', 'clrScheme')[0];
      const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
      if (cs) for (const nm of order) {
        const el = kid(cs, nm);
        const c = el && el.firstElementChild;
        themeColors.push(c ? (c.getAttribute('val') && c.localName === 'srgbClr' ? c.getAttribute('val') : (c.getAttribute('lastClr') || '000000')) : '000000');
      }
    }
    const color = (el) => {
      if (!el) return undefined;
      const rgb = el.getAttribute('rgb');
      if (rgb) return rgb.slice(-6).toUpperCase();
      const th = el.getAttribute('theme');
      if (th !== null && themeColors.length) return applyTint(themeColors[+th] || '000000', parseFloat(el.getAttribute('tint') || '0'));
      const ix = el.getAttribute('indexed');
      if (ix !== null) return INDEXED[+ix] || undefined;
      return undefined;
    };

    // shared strings
    const sst = [];
    for (const k in rels) if (/sharedStrings$/.test(rels[k].type) && z.has(rels[k].path)) {
      const root = parseXml(await z.text(rels[k].path));
      for (const si of kids(root, 'si')) {
        let t = '';
        const walk = (n) => { for (const c of n.children) { if (c.localName === 't') t += c.textContent; else if (c.localName === 'r') walk(c); } };
        walk(si);
        sst.push(t);
      }
    }

    const wb = new Workbook();
    wb.sheets = [];
    const dnEl = kid(wbRoot, 'definedNames');
    if (dnEl) for (const dn of kids(dnEl, 'definedName')) {
      const nm = dn.getAttribute('name');
      if (!nm || nm.startsWith('_xlnm') || dn.getAttribute('localSheetId') !== null || dn.getAttribute('hidden') === '1') continue;
      wb.names[nm.toUpperCase()] = dn.textContent.trim().replace(/^=/, '');
    }

    // styles
    const xfStyle = [0];
    const dxfs = [];
    for (const k in rels) if (/\/styles$/.test(rels[k].type) && z.has(rels[k].path)) {
      const st = parseXml(await z.text(rels[k].path));
      const numFmts = {};
      const nf = kid(st, 'numFmts');
      if (nf) for (const n of kids(nf, 'numFmt')) numFmts[n.getAttribute('numFmtId')] = n.getAttribute('formatCode');
      const fonts = [];
      const fe = kid(st, 'fonts');
      if (fe) for (const f of kids(fe, 'font')) {
        const o = {};
        const flag = (n) => { const e = kid(f, n); return e ? e.getAttribute('val') !== '0' && e.getAttribute('val') !== 'false' : false; };
        if (flag('b')) o.b = true; if (flag('i')) o.i = true; if (flag('strike')) o.s = true;
        const u = kid(f, 'u'); if (u && u.getAttribute('val') !== 'none') o.u = true;
        const sz = kid(f, 'sz'); if (sz) o.size = parseFloat(sz.getAttribute('val'));
        const nm = kid(f, 'name'); if (nm) o.font = nm.getAttribute('val');
        const c = color(kid(f, 'color')); if (c && c !== '000000') o.color = c;
        fonts.push(o);
      }
      const fills = [];
      const fl = kid(st, 'fills');
      if (fl) for (const f of kids(fl, 'fill')) {
        const pf = kid(f, 'patternFill');
        let bg;
        if (pf && pf.getAttribute('patternType') === 'solid') bg = color(kid(pf, 'fgColor')) || color(kid(pf, 'bgColor'));
        else if (kid(f, 'gradientFill')) { const gs = kid(kid(f, 'gradientFill'), 'stop'); bg = gs && color(kid(gs, 'color')); }
        fills.push(bg);
      }
      const borders = [];
      const bd = kid(st, 'borders');
      if (bd) for (const b of kids(bd, 'border')) {
        const o = {};
        for (const [side, key] of [['left', 'bl'], ['right', 'br'], ['top', 'bt'], ['bottom', 'bb']]) {
          const e = kid(b, side);
          if (e && e.getAttribute('style')) o[key] = e.getAttribute('style') + ':' + (color(kid(e, 'color')) || '000000');
        }
        borders.push(o);
      }
      const dx = kid(st, 'dxfs');
      if (dx) for (const d of kids(dx, 'dxf')) {
        const o = {};
        const f = kid(d, 'font');
        if (f) {
          const fl = (n) => { const e = kid(f, n); return e ? e.getAttribute('val') !== '0' && e.getAttribute('val') !== 'false' : false; };
          if (fl('b')) o.b = true; if (fl('i')) o.i = true;
          const c = color(kid(f, 'color')); if (c) o.color = c;
        }
        const fi = kid(d, 'fill');
        const pf = fi && kid(fi, 'patternFill');
        if (pf) { const bgc = color(kid(pf, 'bgColor')) || color(kid(pf, 'fgColor')); if (bgc) o.bg = bgc; }
        dxfs.push(o);
      }
      const cx = kid(st, 'cellXfs');
      if (cx) {
        xfStyle.length = 0;
        for (const xf of kids(cx, 'xf')) {
          const o = {};
          const font = fonts[+xf.getAttribute('fontId') || 0] || {};
          Object.assign(o, font);
          if (o.size === 11) delete o.size;
          if (o.font === 'Calibri') delete o.font;
          const fill = fills[+xf.getAttribute('fillId') || 0];
          if (fill) o.bg = fill;
          Object.assign(o, borders[+xf.getAttribute('borderId') || 0] || {});
          const id = xf.getAttribute('numFmtId') || '0';
          const code = numFmts[id] || NumFmt.BUILTIN[id];
          if (code && code !== 'General') o.nf = code;
          const al = kid(xf, 'alignment');
          if (al) {
            const h = al.getAttribute('horizontal'); if (h && h !== 'general') o.h = h === 'centerContinuous' ? 'center' : h;
            const v = al.getAttribute('vertical'); if (v && v !== 'bottom') o.v = v;
            if (al.getAttribute('wrapText') === '1' || al.getAttribute('wrapText') === 'true') o.wrap = true;
            const ind = parseInt(al.getAttribute('indent'), 10); if (ind) o.indent = ind;
          }
          xfStyle.push(wb.styleId(o));
        }
      }
    }

    // sheets
    const sheetsEl = kid(wbRoot, 'sheets');
    const bv = kid(kid(wbRoot, 'bookViews'), 'workbookView');
    for (const s of kids(sheetsEl, 'sheet')) {
      const rid = s.getAttributeNS(NS_R, 'id') || s.getAttribute('r:id');
      const rel = rels[rid];
      const sh = new Sheet(s.getAttribute('name'));
      wb.sheets.push(sh);
      if (!rel || !z.has(rel.path)) continue;
      const root = parseXml(await z.text(rel.path));
      const fp = kid(root, 'sheetFormatPr');
      if (fp) {
        const dc = parseFloat(fp.getAttribute('defaultColWidth')), dr = parseFloat(fp.getAttribute('defaultRowHeight'));
        if (dc > 0) sh.defColW = Math.round(dc * 7 + 5);
        if (dr > 0) sh.defRowH = Math.round(dr / 0.75);
      }
      const colsEl = kid(root, 'cols');
      if (colsEl) for (const c of kids(colsEl, 'col')) {
        const mn = +c.getAttribute('min') - 1, mx = +c.getAttribute('max') - 1;
        const hidden = c.getAttribute('hidden') === '1' || c.getAttribute('hidden') === 'true';
        const w = hidden ? 0 : Math.round(parseFloat(c.getAttribute('width') || '8.43') * 7 + 5);
        if (mx - mn > 300) { if (!hidden) sh.defColW = w; continue; }
        for (let i = mn; i <= mx && i < MAXC; i++) sh.colW.set(i, w);
      }
      const shared = {};
      const arrRefs = [];
      const sd = kid(root, 'sheetData');
      if (sd) for (const row of kids(sd, 'row')) {
        const rr = +row.getAttribute('r') - 1;
        const ht = row.getAttribute('ht');
        if (row.getAttribute('hidden') === '1') sh.rowH.set(rr, 0);
        else if (ht && (row.getAttribute('customHeight') === '1' || Math.abs(parseFloat(ht) / 0.75 - sh.defRowH) > 1)) sh.rowH.set(rr, Math.round(parseFloat(ht) / 0.75));
        let nextC = 0;
        for (const c of kids(row, 'c')) {
          const ref = c.getAttribute('r');
          let col, r;
          if (ref) { const m = /^([A-Z]+)(\d+)$/.exec(ref); col = colIndex(m[1]); r = +m[2] - 1; } else { col = nextC; r = rr; }
          nextC = col + 1;
          const cell = {};
          const sIdx = c.getAttribute('s');
          if (sIdx && xfStyle[+sIdx]) cell.s = xfStyle[+sIdx];
          const t = c.getAttribute('t');
          const ve = kid(c, 'v'), fe = kid(c, 'f');
          let val;
          if (t === 'inlineStr') { const is = kid(c, 'is'); val = is ? [...is.getElementsByTagName('t')].map(x => x.textContent).join('') : ''; }
          else if (ve) {
            const txt = ve.textContent;
            if (t === 's') val = sst[+txt] !== undefined ? sst[+txt] : '';
            else if (t === 'str') val = txt;
            else if (t === 'b') val = txt === '1';
            else if (t === 'e') val = ERR_BY_CODE[txt] || new XErr(txt);
            else { val = parseFloat(txt); if (isNaN(val)) val = txt; }
          }
          if (fe) {
            let f = fe.textContent;
            const ft = fe.getAttribute('t'), si = fe.getAttribute('si');
            if (ft === 'array' && fe.getAttribute('ref') && f) arrRefs.push({ o: { r, c: col }, rg: Formula.parseRange(fe.getAttribute('ref')) });
            if (ft === 'shared') {
              if (f) shared[si] = { f, r, c: col };
              else if (shared[si]) f = shiftFormula(shared[si].f, r - shared[si].r, col - shared[si].c);
            }
            if (f) { cell.f = f; cell.cv = val; }
            else if (val !== undefined) cell.v = val;
          } else if (val !== undefined) cell.v = val;
          if (cell.v !== undefined || cell.f !== undefined || cell.s) sh.cells.set(r * MAXC + col, cell);
        }
      }
      for (const { o, rg } of arrRefs) {
        if (!rg || (rg.r1 === rg.r2 && rg.c1 === rg.c2)) continue;
        for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
          if (r === o.r && c === o.c) continue;
          const cell = sh.cells.get(r * MAXC + c);
          if (cell && cell.f == null) { if (cell.s) sh.cells.set(r * MAXC + c, { s: cell.s }); else sh.cells.delete(r * MAXC + c); }
        }
      }
      const mc = kid(root, 'mergeCells');
      if (mc) for (const m of kids(mc, 'mergeCell')) { const rg = Formula.parseRange(m.getAttribute('ref')); if (rg) sh.merges.push(rg); }
      // conditional formatting
      for (const cfEl of kids(root, 'conditionalFormatting')) {
        for (const sq of (cfEl.getAttribute('sqref') || '').split(/\s+/).filter(Boolean)) {
          const rg = Formula.parseRange(sq);
          if (!rg) continue;
          for (const cr of kids(cfEl, 'cfRule')) {
            const type = cr.getAttribute('type');
            const dxId = cr.getAttribute('dxfId');
            const style = dxId !== null && dxfs[+dxId] ? { ...dxfs[+dxId] } : null;
            const fs = [...cr.getElementsByTagNameNS('*', 'formula')].map(f => f.textContent);
            const lit = (s) => { if (s === undefined) return ''; if (/^"(.*)"$/s.test(s)) return s.slice(1, -1).replace(/""/g, '"'); if (s !== '' && !isNaN(Number(s))) return s; return '=' + s; };
            const rule = { id: 'f' + Math.random().toString(36).slice(2, 8), ...rg, p: parseInt(cr.getAttribute('priority'), 10) || 9999 };
            if (type === 'cellIs') {
              const op = { greaterThan: '>', lessThan: '<', greaterThanOrEqual: '>=', lessThanOrEqual: '<=', equal: '=', notEqual: '<>', between: 'between' }[cr.getAttribute('operator')];
              if (!op || !style) continue;
              Object.assign(rule, { type: 'cell', op, v1: lit(fs[0]), v2: lit(fs[1]), style });
            } else if (['containsText', 'beginsWith', 'endsWith', 'notContainsText'].includes(type)) {
              if (!style) continue;
              Object.assign(rule, { type: 'text', op: { containsText: 'contains', beginsWith: 'begins', endsWith: 'ends', notContainsText: 'notcontains' }[type], text: cr.getAttribute('text') || '', style });
            } else if (type === 'duplicateValues') { if (!style) continue; Object.assign(rule, { type: 'dup', style }); }
            else if (type === 'top10') { if (!style || cr.getAttribute('percent') === '1') continue; Object.assign(rule, { type: 'top', n: parseInt(cr.getAttribute('rank'), 10) || 10, bottom: cr.getAttribute('bottom') === '1', style }); }
            else if (type === 'colorScale') {
              const cs = kid(cr, 'colorScale');
              const cfvo = kids(cs, 'cfvo').map(v => ({ t: v.getAttribute('type'), v: v.getAttribute('val') }));
              const colors = kids(cs, 'color').map(c => color(c) || 'FFFFFF');
              if (cfvo.length < 2 || cfvo.length !== colors.length) continue;
              Object.assign(rule, { type: 'scale', cfvo, colors });
            } else if (type === 'dataBar') {
              const db = kid(cr, 'dataBar');
              Object.assign(rule, { type: 'bar', color: (db && color(kid(db, 'color'))) || '638EC6' });
            } else if (type === 'expression') { if (!style || !fs[0]) continue; Object.assign(rule, { type: 'expr', f: fs[0], style }); }
            else continue;
            sh.cf.push(rule);
          }
        }
      }
      sh.cf.sort((a, b) => a.p - b.p);
      // data validation
      const dvs = kid(root, 'dataValidations');
      if (dvs) for (const dv of kids(dvs, 'dataValidation')) {
        const type = dv.getAttribute('type');
        if (!['list', 'whole', 'decimal', 'date', 'textLength', 'custom'].includes(type)) continue;
        const f1 = kid(dv, 'formula1'), f2 = kid(dv, 'formula2');
        const conv = (el) => {
          if (!el) return '';
          const t = el.textContent;
          if (type === 'list') return /^".*"$/s.test(t) ? t.slice(1, -1).replace(/""/g, '"') : '=' + t;
          if (type === 'custom') return '=' + t;
          return t !== '' && !isNaN(Number(t)) ? t : '=' + t;
        };
        for (const sq of (dv.getAttribute('sqref') || '').split(/\s+/).filter(Boolean)) {
          const rg = Formula.parseRange(sq);
          if (!rg) continue;
          sh.dv.push({ id: 'v' + Math.random().toString(36).slice(2, 8), ...rg, type, op: dv.getAttribute('operator') || 'between', f1: conv(f1), f2: conv(f2), blank: dv.getAttribute('allowBlank') === '1' || dv.getAttribute('allowBlank') === 'true', drop: !(dv.getAttribute('showDropDown') === '1' || dv.getAttribute('showDropDown') === 'true'), style: dv.getAttribute('errorStyle') || 'stop', title: dv.getAttribute('errorTitle') || '', msg: dv.getAttribute('error') || '' });
        }
      }
      // autofilter
      const af = kid(root, 'autoFilter');
      if (af && af.getAttribute('ref')) {
        const rg = Formula.parseRange(af.getAttribute('ref'));
        if (rg) {
          const f = { ...rg, cols: {}, orig: {} };
          for (const fc of kids(af, 'filterColumn')) {
            const fl = kid(fc, 'filters');
            if (!fl) continue;
            const vals = kids(fl, 'filter').map(x => x.getAttribute('val'));
            if (fl.getAttribute('blank') === '1') vals.push('');
            f.cols[rg.c1 + (parseInt(fc.getAttribute('colId'), 10) || 0)] = { vals };
          }
          sh.filter = f;
        }
      }
      // comments (notes)
      try {
        const sdir0 = dirname(rel.path);
        const srels0 = parseRels(await z.textOrNull(sdir0 + '/_rels/' + rel.path.split('/').pop() + '.rels'), sdir0);
        for (const k in srels0) {
          if (/\/comments$/.test(srels0[k].type) && z.has(srels0[k].path)) {
            const croot = parseXml(await z.text(srels0[k].path));
            const authors = kids(kid(croot, 'authors'), 'author').map(a => a.textContent);
            for (const cm of kids(kid(croot, 'commentList'), 'comment')) {
              const p = Formula.parseCell(cm.getAttribute('ref') || '');
              if (!p) continue;
              sh.comments.push({ r: p.r, c: p.c, text: [...cm.getElementsByTagNameNS('*', 't')].map(t => t.textContent).join(''), author: authors[parseInt(cm.getAttribute('authorId'), 10) || 0] || '' });
            }
          }
        }
      } catch (e) { console.warn('comments', e); }
      // tables
      try {
        const sdir1 = dirname(rel.path);
        const srels1 = parseRels(await z.textOrNull(sdir1 + '/_rels/' + rel.path.split('/').pop() + '.rels'), sdir1);
        sh.tables = [];
        for (const k in srels1) {
          if (!/\/table$/.test(srels1[k].type) || !z.has(srels1[k].path)) continue;
          const troot = parseXml(await z.text(srels1[k].path));
          const rg = Formula.parseRange(troot.getAttribute('ref') || '');
          if (!rg) continue;
          const si = kid(troot, 'tableStyleInfo');
          const t = { name: troot.getAttribute('displayName') || troot.getAttribute('name') || 'Table' + (sh.tables.length + 1), ...rg, style: (si && si.getAttribute('name')) || 'TableStyleMedium2' };
          if (troot.getAttribute('headerRowCount') === '0') t.hdr = false;
          if ((parseInt(troot.getAttribute('totalsRowCount'), 10) || 0) > 0) { t.totals = true; t.tcols = kids(kid(troot, 'tableColumns'), 'tableColumn').map(tc => ({ label: tc.getAttribute('totalsRowLabel') || '', fn: tc.getAttribute('totalsRowFunction') || '' })); }
          sh.tables.push(t);
          const taf = kid(troot, 'autoFilter');
          if (taf && !sh.filter) {
            const frg = Formula.parseRange(taf.getAttribute('ref') || '');
            if (frg) {
              const f = { ...frg, cols: {}, orig: {} };
              for (const fc of kids(taf, 'filterColumn')) {
                const fl = kid(fc, 'filters');
                if (!fl) continue;
                const vals = kids(fl, 'filter').map(x => x.getAttribute('val'));
                if (fl.getAttribute('blank') === '1') vals.push('');
                f.cols[frg.c1 + (parseInt(fc.getAttribute('colId'), 10) || 0)] = { vals };
              }
              sh.filter = f;
            }
          }
        }
      } catch (e) { console.warn('tables', e); }
      // sparklines
      try {
        sh.sparks = [];
        for (const g of root.getElementsByTagNameNS('*', 'sparklineGroup')) {
          const type = g.getAttribute('type') === 'column' ? 'column' : g.getAttribute('type') === 'stacked' ? 'stacked' : 'line';
          const cs = [...g.children].find(e => e.localName === 'colorSeries');
          const color = cs && cs.getAttribute('rgb') ? cs.getAttribute('rgb').slice(-6) : '376092';
          for (const s1 of g.getElementsByTagNameNS('*', 'sparkline')) {
            const fEl = [...s1.children].find(e => e.localName === 'f'), qEl = [...s1.children].find(e => e.localName === 'sqref');
            if (!fEl || !qEl) continue;
            let ref = fEl.textContent.trim();
            const bang = ref.lastIndexOf('!');
            if (bang >= 0 && ref.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'") === sh.name) ref = ref.slice(bang + 1);
            const p = Formula.parseCell(qEl.textContent.trim().split(/[ :]/)[0]);
            if (p) sh.sparks.push({ r: p.r, c: p.c, ref: ref.replace(/\$/g, ''), type, color });
          }
        }
      } catch (e) { console.warn('sparklines', e); }
      // charts
      const drEl = kid(root, 'drawing');
      if (drEl) {
        try {
          const rid = drEl.getAttributeNS(NS_R, 'id') || drEl.getAttribute('r:id');
          const sdir = dirname(rel.path);
          const srels = parseRels(await z.textOrNull(sdir + '/_rels/' + rel.path.split('/').pop() + '.rels'), sdir);
          const drel = srels[rid];
          if (drel && z.has(drel.path)) {
            const droot = parseXml(await z.text(drel.path));
            const ddir = dirname(drel.path);
            const drels = parseRels(await z.textOrNull(ddir + '/_rels/' + drel.path.split('/').pop() + '.rels'), ddir);
            const offs = (map, def, i) => { let s = 0; for (let k = 0; k < i; k++) s += map.has(k) ? map.get(k) : def; return s; };
            const pos = (m) => { const gv = (n) => { const e = kid(m, n); return e ? parseFloat(e.textContent) || 0 : 0; }; return { x: offs(sh.colW, sh.defColW, gv('col')) + gv('colOff') / 9525, y: offs(sh.rowH, sh.defRowH, gv('row')) + gv('rowOff') / 9525 }; };
            for (const an of droot.children) {
              if (an.localName !== 'twoCellAnchor' && an.localName !== 'oneCellAnchor' && an.localName !== 'absoluteAnchor') continue;
              const cel = an.getElementsByTagNameNS('*', 'chart')[0];
              if (!cel) continue;
              const crid = cel.getAttributeNS(NS_R, 'id') || cel.getAttribute('r:id');
              const crel = drels[crid];
              if (!crel || !z.has(crel.path)) continue;
              const ch = Charts.parseChart(parseXml(await z.text(crel.path)));
              if (!ch) continue;
              let x = 0, y = 0, w = 480, h = 300;
              if (an.localName === 'twoCellAnchor') { const a = pos(kid(an, 'from')), b = pos(kid(an, 'to')); x = a.x; y = a.y; w = Math.max(60, b.x - a.x); h = Math.max(60, b.y - a.y); }
              else if (an.localName === 'oneCellAnchor') { const a = pos(kid(an, 'from')), ex = kid(an, 'ext'); x = a.x; y = a.y; if (ex) { w = parseFloat(ex.getAttribute('cx')) / 9525; h = parseFloat(ex.getAttribute('cy')) / 9525; } }
              sh.charts.push({ id: 'c' + Math.random().toString(36).slice(2, 8), x: Math.round(x), y: Math.round(y), w: Math.round(w), h: Math.round(h), ...ch });
            }
          }
        } catch (e) { console.warn('Could not read charts', e); }
      }
      const sv = kid(kid(root, 'sheetViews'), 'sheetView');
      if (sv) {
        if (sv.getAttribute('showGridLines') === '0' || sv.getAttribute('showGridLines') === 'false') sh.showGrid = false;
        const pane = kid(sv, 'pane');
        if (pane && /frozen/.test(pane.getAttribute('state') || '')) sh.freeze = { r: parseInt(pane.getAttribute('ySplit')) || 0, c: parseInt(pane.getAttribute('xSplit')) || 0 };
      }
    }
    if (!wb.sheets.length) wb.sheets.push(new Sheet('Sheet1'));
    if (bv) wb.active = Math.min(wb.sheets.length - 1, parseInt(bv.getAttribute('activeTab')) || 0);
    // refreshable PivotTables saved by Lumen Cells
    try {
      const cx = z.has('docProps/custom.xml') ? parseXml(await z.text('docProps/custom.xml')) : null;
      if (cx) for (const p of cx.getElementsByTagNameNS('*', 'property')) {
        if (p.getAttribute('name') !== 'LumenPivots') continue;
        for (const it of JSON.parse(p.textContent || '[]')) { const s = wb.sheetByName(it.sheet); if (s && it.def && wb.sheetByName(it.def.src)) s.pivot = it.def; }
      }
    } catch (e) { console.warn('pivot settings ignored', e); }
    return wb;
  }

  /* ====================================================================
     WRITE
     ==================================================================== */
  // Table column names come from the header cells; they must be unique, non-empty text.
  function tableColNames(wb, sh, t) {
    const seen = new Set(), out = [];
    for (let c = t.c1; c <= t.c2; c++) {
      let n = '';
      if (t.hdr !== false) { const v = Formula.cellValue(wb, sh, t.r1, c); n = (v === null || v === undefined || typeof v === 'object') ? '' : String(v); }
      n = n.replace(/[\r\n]+/g, ' ').trim() || 'Column' + (c - t.c1 + 1);
      let u = n, i = 2; while (seen.has(u.toLowerCase())) u = n + i++;
      seen.add(u.toLowerCase()); out.push(u);
    }
    return out;
  }
  function normTables(wb) {
    let ch = false;
    for (const sh of wb.sheets) for (const t of (sh.tables || [])) {
      if (t.hdr === false) continue;
      tableColNames(wb, sh, t).forEach((n, i) => {
        const k = t.r1 * MAXC + t.c1 + i, cell = sh.cells.get(k);
        if (!cell || cell.f != null || cell.v !== n) { const o = { v: n }; if (cell && cell.s) o.s = cell.s; sh.cells.set(k, o); ch = true; }
      });
    }
    if (ch) wb.touch();
  }

  function build(wb) {
    normTables(wb);
    // styles
    const fonts = [{ size: 11, font: 'Calibri' }], fontKeys = new Map([[JSON.stringify(fonts[0]), 0]]);
    const fills = [null, 'gray'], fillKeys = new Map();
    const borders = [{}], borderKeys = new Map([['{}', 0]]);
    const numFmts = [], nfKeys = new Map();
    const xfs = [];
    const intern = (arr, map, o) => { const k = JSON.stringify(o); let i = map.get(k); if (i === undefined) { i = arr.length; arr.push(o); map.set(k, i); } return i; };
    wb.styles.forEach((s) => {
      const f = { size: s.size || 11, font: s.font || 'Calibri' };
      if (s.b) f.b = 1; if (s.i) f.i = 1; if (s.u) f.u = 1; if (s.s) f.s = 1; if (s.color) f.color = s.color;
      const fontId = intern(fonts, fontKeys, f);
      let fillId = 0;
      if (s.bg) fillId = intern(fills, fillKeys, s.bg);
      const b = {};
      for (const k of ['bl', 'br', 'bt', 'bb']) if (s[k]) b[k] = s[k];
      const borderId = intern(borders, borderKeys, b);
      let nfId = 0;
      if (s.nf && s.nf !== 'General') {
        if (NumFmt.BUILTIN_REV[s.nf] !== undefined) nfId = NumFmt.BUILTIN_REV[s.nf];
        else { if (!nfKeys.has(s.nf)) { nfKeys.set(s.nf, 164 + numFmts.length); numFmts.push(s.nf); } nfId = nfKeys.get(s.nf); }
      }
      xfs.push({ nfId, fontId, fillId, borderId, h: s.h, v: s.v, wrap: s.wrap, indent: s.indent });
    });
    const dxfs = [], dxfKeys = new Map();
    const dxfOf = (st) => { const k = JSON.stringify(st); if (!dxfKeys.has(k)) { dxfKeys.set(k, dxfs.length); dxfs.push(st); } return dxfKeys.get(k); };
    wb.sheets.forEach(s => (s.cf || []).forEach(r => { if (r.style) dxfOf(r.style); }));
    const clr = (c) => `<color rgb="FF${c}"/>`;
    let stylesXml = XML + `<styleSheet xmlns="${NS_M}">`;
    if (numFmts.length) stylesXml += `<numFmts count="${numFmts.length}">` + numFmts.map((c, i) => `<numFmt numFmtId="${164 + i}" formatCode="${xesc(c)}"/>`).join('') + '</numFmts>';
    stylesXml += `<fonts count="${fonts.length}">` + fonts.map(f => '<font>' + (f.b ? '<b/>' : '') + (f.i ? '<i/>' : '') + (f.s ? '<strike/>' : '') + (f.u ? '<u/>' : '') + `<sz val="${f.size}"/>` + (f.color ? clr(f.color) : '<color theme="1"/>') + `<name val="${xesc(f.font)}"/><family val="2"/></font>`).join('') + '</fonts>';
    stylesXml += `<fills count="${fills.length}">` + fills.map((f, i) => i === 0 ? '<fill><patternFill patternType="none"/></fill>' : i === 1 ? '<fill><patternFill patternType="gray125"/></fill>' : `<fill><patternFill patternType="solid"><fgColor rgb="FF${f}"/><bgColor indexed="64"/></patternFill></fill>`).join('') + '</fills>';
    stylesXml += `<borders count="${borders.length}">` + borders.map(b => {
      const side = (tag, v) => { if (!v) return `<${tag}/>`; const [st, c] = v.split(':'); return `<${tag} style="${st}">${clr(c || '000000')}</${tag}>`; };
      return '<border>' + side('left', b.bl) + side('right', b.br) + side('top', b.bt) + side('bottom', b.bb) + '<diagonal/></border>';
    }).join('') + '</borders>';
    stylesXml += '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>';
    stylesXml += `<cellXfs count="${xfs.length}">` + xfs.map(x => {
      const al = (x.h || x.v || x.wrap || x.indent) ? `<alignment${x.h ? ` horizontal="${x.h}"` : ''}${x.v ? ` vertical="${x.v}"` : ''}${x.wrap ? ' wrapText="1"' : ''}${x.indent ? ` indent="${x.indent}"` : ''}/>` : '';
      return `<xf numFmtId="${x.nfId}" fontId="${x.fontId}" fillId="${x.fillId}" borderId="${x.borderId}" xfId="0"${x.nfId ? ' applyNumberFormat="1"' : ''}${x.fontId ? ' applyFont="1"' : ''}${x.fillId ? ' applyFill="1"' : ''}${x.borderId ? ' applyBorder="1"' : ''}${al ? ' applyAlignment="1">' + al + '</xf>' : '/>'}`;
    }).join('') + '</cellXfs>';
    stylesXml += '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>';
    if (dxfs.length) stylesXml += `<dxfs count="${dxfs.length}">` + dxfs.map(d => '<dxf>' + ((d.b || d.i || d.color) ? '<font>' + (d.b ? '<b/>' : '') + (d.i ? '<i/>' : '') + (d.color ? clr(d.color) : '') + '</font>' : '') + (d.bg ? `<fill><patternFill><bgColor rgb="FF${d.bg}"/></patternFill></fill>` : '') + '</dxf>').join('') + '</dxfs>';
    stylesXml += '</styleSheet>';

    // shared strings + sheets
    const sst = [], sstMap = new Map();
    const sstIdx = (s) => { let i = sstMap.get(s); if (i === undefined) { i = sst.length; sst.push(s); sstMap.set(s, i); } return i; };
    let nStr = 0;
    let drawN = 0, chartN = 0, cmtN = 0, tableN = 0, hasVml = false;
    const extras = [], cts = [];
    const sheetXmls = wb.sheets.map((sh, si) => {
      const rows = new Map();
      let maxR = 0, maxC = 0;
      for (const [k, cell] of sh.cells) {
        const r = Math.floor(k / MAXC), c = k - r * MAXC;
        if (!rows.has(r)) rows.set(r, []);
        rows.get(r).push([c, cell]);
        if (r > maxR) maxR = r; if (c > maxC) maxC = c;
      }
      for (const sp of Formula.spills(wb, sh)) {
        for (let i = 0; i < sp.rows.length; i++) for (let j = 0; j < sp.rows[0].length; j++) {
          if (i === 0 && j === 0) continue;
          const r = sp.r + i, c = sp.c + j;
          if (sh.get(r, c)) continue;
          if (!rows.has(r)) rows.set(r, []);
          rows.get(r).push([c, { v: sp.rows[i][j] === undefined ? null : sp.rows[i][j] }]);
          if (r > maxR) maxR = r; if (c > maxC) maxC = c;
        }
      }
      for (const r of sh.rowH.keys()) if (!rows.has(r)) rows.set(r, []);
      const rowKeys = [...rows.keys()].sort((a, b) => a - b);
      let sd = '';
      for (const r of rowKeys) {
        const cells = rows.get(r).sort((a, b) => a[0] - b[0]);
        const h = sh.rowH.get(r);
        sd += `<row r="${r + 1}"` + (h === 0 ? ' hidden="1"' : (h ? ` ht="${h * 0.75}" customHeight="1"` : '')) + '>';
        for (const [c, cell] of cells) {
          const ref = cellName(r, c);
          const s = cell.s ? ` s="${cell.s}"` : '';
          if (cell.f != null) {
            let v = Formula.cellValue(wb, sh, r, c);
            let t = '', vx = '';
            if (typeof v === 'number') vx = isFinite(v) ? `<v>${v}</v>` : (t = ' t="e"', '<v>#NUM!</v>');
            else if (typeof v === 'boolean') { t = ' t="b"'; vx = `<v>${v ? 1 : 0}</v>`; }
            else if (v instanceof XErr) { t = ' t="e"'; vx = `<v>${xesc(v.code)}</v>`; }
            else if (typeof v === 'string') { t = ' t="str"'; vx = `<v>${xesc(v)}</v>`; }
            const arrRef = cell._arr ? ` t="array" ref="${ref}:${cellName(r + cell._arr.length - 1, c + cell._arr[0].length - 1)}"` : '';
            sd += `<c r="${ref}"${s}${t}><f${arrRef}>${xesc(cell.f)}</f>${vx}</c>`;
          } else if (cell.v === undefined || cell.v === null) sd += `<c r="${ref}"${s}/>`;
          else if (typeof cell.v === 'number') sd += isFinite(cell.v) ? `<c r="${ref}"${s}><v>${cell.v}</v></c>` : `<c r="${ref}"${s} t="e"><v>#NUM!</v></c>`;
          else if (typeof cell.v === 'boolean') sd += `<c r="${ref}"${s} t="b"><v>${cell.v ? 1 : 0}</v></c>`;
          else if (cell.v instanceof XErr) sd += `<c r="${ref}"${s} t="e"><v>${xesc(cell.v.code)}</v></c>`;
          else { nStr++; sd += `<c r="${ref}"${s} t="s"><v>${sstIdx(String(cell.v))}</v></c>`; }
        }
        sd += '</row>';
      }
      let x = XML + `<worksheet xmlns="${NS_M}" xmlns:r="${NS_R}">`;
      x += `<dimension ref="A1${rowKeys.length ? ':' + cellName(maxR, maxC) : ''}"/>`;
      x += `<sheetViews><sheetView${sh.showGrid ? '' : ' showGridLines="0"'}${si === wb.active ? ' tabSelected="1"' : ''} workbookViewId="0">`;
      if (sh.freeze.r || sh.freeze.c) {
        const tl = cellName(sh.freeze.r, sh.freeze.c);
        const pane = sh.freeze.r && sh.freeze.c ? 'bottomRight' : (sh.freeze.r ? 'bottomLeft' : 'topRight');
        x += `<pane${sh.freeze.c ? ` xSplit="${sh.freeze.c}"` : ''}${sh.freeze.r ? ` ySplit="${sh.freeze.r}"` : ''} topLeftCell="${tl}" activePane="${pane}" state="frozen"/><selection pane="${pane}"/>`;
      }
      x += '</sheetView></sheetViews>';
      x += `<sheetFormatPr defaultRowHeight="${sh.defRowH * 0.75}"${sh.defColW !== 64 ? ` defaultColWidth="${(sh.defColW - 5) / 7}"` : ''}/>`;
      if (sh.colW.size) {
        x += '<cols>' + [...sh.colW.keys()].sort((a, b) => a - b).map(c => { const w = sh.colW.get(c); return `<col min="${c + 1}" max="${c + 1}" width="${w === 0 ? 8.43 : Math.max(0, (w - 5) / 7)}"${w === 0 ? ' hidden="1"' : ''} customWidth="1"/>`; }).join('') + '</cols>';
      }
      x += `<sheetData>${sd}</sheetData>`;
      const filterCols = (f) => Object.keys(f.cols || {}).map(ck => {
        const vals = f.cols[ck].vals || [];
        const blank = vals.includes('');
        return `<filterColumn colId="${+ck - f.c1}"><filters${blank ? ' blank="1"' : ''}>` + vals.filter(v => v !== '').map(v => `<filter val="${xesc(v)}"/>`).join('') + '</filters></filterColumn>';
      }).join('');
      const tables = sh.tables || [];
      // a sheet-level filter may not overlap a table: the table carries its own
      const filterInTable = sh.filter && tables.find(t => !(sh.filter.r2 < t.r1 || sh.filter.r1 > t.r2 || sh.filter.c2 < t.c1 || sh.filter.c1 > t.c2));
      if (sh.filter && !filterInTable) {
        const f = sh.filter;
        x += `<autoFilter ref="${cellName(f.r1, f.c1)}:${cellName(f.r2, f.c2)}">` + filterCols(f) + '</autoFilter>';
      }
      if (sh.merges.length) x += `<mergeCells count="${sh.merges.length}">` + sh.merges.map(m => `<mergeCell ref="${cellName(m.r1, m.c1)}:${cellName(m.r2, m.c2)}"/>`).join('') + '</mergeCells>';
      (sh.cf || []).forEach((r, ri) => {
        const sq = cellName(r.r1, r.c1) + ((r.r1 === r.r2 && r.c1 === r.c2) ? '' : ':' + cellName(r.r2, r.c2));
        const tl = cellName(r.r1, r.c1);
        const pr = ri + 1;
        const dx = r.style ? ` dxfId="${dxfOf(r.style)}"` : '';
        const fv = (s) => { s = String(s === undefined || s === null ? '' : s); if (s[0] === '=') return xesc(s.slice(1)); if (s !== '' && !isNaN(Number(s))) return s; return '&quot;' + xesc(s.replace(/"/g, '""')).replace(/&amp;quot;/g, '&quot;') + '&quot;'; };
        let body = '';
        if (r.type === 'cell') {
          const op = { '>': 'greaterThan', '<': 'lessThan', '>=': 'greaterThanOrEqual', '<=': 'lessThanOrEqual', '=': 'equal', '<>': 'notEqual', between: 'between' }[r.op];
          body = `<cfRule type="cellIs"${dx} priority="${pr}" operator="${op}"><formula>${fv(r.v1)}</formula>${r.op === 'between' ? `<formula>${fv(r.v2)}</formula>` : ''}</cfRule>`;
        } else if (r.type === 'text') {
          const t = String(r.text || '');
          const q2 = '&quot;' + xesc(t.replace(/"/g, '""')) + '&quot;';
          const m = { contains: ['containsText', `NOT(ISERROR(SEARCH(${q2},${tl})))`], begins: ['beginsWith', `LEFT(${tl},${t.length})=${q2}`], ends: ['endsWith', `RIGHT(${tl},${t.length})=${q2}`], notcontains: ['notContainsText', `ISERROR(SEARCH(${q2},${tl}))`] }[r.op || 'contains'];
          body = `<cfRule type="${m[0]}"${dx} priority="${pr}" operator="${m[0]}" text="${xesc(t)}"><formula>${m[1]}</formula></cfRule>`;
        } else if (r.type === 'dup') body = `<cfRule type="duplicateValues"${dx} priority="${pr}"/>`;
        else if (r.type === 'top') body = `<cfRule type="top10"${dx} priority="${pr}" rank="${r.n || 10}"${r.bottom ? ' bottom="1"' : ''}/>`;
        else if (r.type === 'scale') body = `<cfRule type="colorScale" priority="${pr}"><colorScale>` + r.cfvo.map(v => `<cfvo type="${v.t}"${v.v !== undefined && v.v !== null && v.t !== 'min' && v.t !== 'max' ? ` val="${xesc(v.v)}"` : ''}/>`).join('') + r.colors.map(c => `<color rgb="FF${c}"/>`).join('') + '</colorScale></cfRule>';
        else if (r.type === 'bar') body = `<cfRule type="dataBar" priority="${pr}"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF${r.color || '638EC6'}"/></dataBar></cfRule>`;
        else if (r.type === 'expr') body = `<cfRule type="expression"${dx} priority="${pr}"><formula>${xesc(r.f)}</formula></cfRule>`;
        if (body) x += `<conditionalFormatting sqref="${sq}">${body}</conditionalFormatting>`;
      });
      if (sh.dv && sh.dv.length) {
        x += `<dataValidations count="${sh.dv.length}">` + sh.dv.map(d => {
          const sq = cellName(d.r1, d.c1) + ((d.r1 === d.r2 && d.c1 === d.c2) ? '' : ':' + cellName(d.r2, d.c2));
          const fx = (s) => { s = String(s === undefined || s === null ? '' : s); if (d.type === 'list') return s[0] === '=' ? xesc(s.slice(1)) : '&quot;' + xesc(s.replace(/"/g, '""')).replace(/&amp;quot;/g, '&quot;') + '&quot;'; return s[0] === '=' ? xesc(s.slice(1)) : xesc(s); };
          const withOp = !['list', 'custom'].includes(d.type);
          return `<dataValidation type="${d.type}"${d.style && d.style !== 'stop' ? ` errorStyle="${d.style}"` : ''}${withOp && d.op && d.op !== 'between' ? ` operator="${d.op}"` : ''}${d.blank !== false ? ' allowBlank="1"' : ''}${d.drop === false ? ' showDropDown="1"' : ''} showInputMessage="1" showErrorMessage="1"${d.title ? ` errorTitle="${xesc(d.title)}"` : ''}${d.msg ? ` error="${xesc(d.msg)}"` : ''} sqref="${sq}"><formula1>${fx(d.f1)}</formula1>${withOp && (d.op || 'between').toLowerCase().includes('between') ? `<formula2>${fx(d.f2)}</formula2>` : ''}</dataValidation>`;
        }).join('') + '</dataValidations>';
      }
      x += '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';
      const rels = [];
      if (sh.charts && sh.charts.length) {
        const dn = ++drawN;
        const cellAt = (map, def, px) => { let i = 0, acc = 0; while (i < 20000) { const w = map.has(i) ? map.get(i) : def; if (acc + w > px) break; acc += w; i++; } return { i, off: Math.max(0, px - acc) }; };
        let anchors = '', drRels = '';
        sh.charts.forEach((ch, k) => {
          const cn = ++chartN;
          extras.push({ path: `xl/charts/chart${cn}.xml`, data: Charts.chartXml(ch, Charts.chartData(wb, ch, sh)) });
          cts.push(`<Override PartName="/xl/charts/chart${cn}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);
          drRels += `<Relationship Id="rId${k + 1}" Type="${REL}/chart" Target="../charts/chart${cn}.xml"/>`;
          const a = cellAt(sh.colW, sh.defColW, ch.x), b = cellAt(sh.colW, sh.defColW, ch.x + ch.w);
          const c = cellAt(sh.rowH, sh.defRowH, ch.y), d = cellAt(sh.rowH, sh.defRowH, ch.y + ch.h);
          const E = (v) => Math.round(v * 9525);
          anchors += `<xdr:twoCellAnchor><xdr:from><xdr:col>${a.i}</xdr:col><xdr:colOff>${E(a.off)}</xdr:colOff><xdr:row>${c.i}</xdr:row><xdr:rowOff>${E(c.off)}</xdr:rowOff></xdr:from><xdr:to><xdr:col>${b.i}</xdr:col><xdr:colOff>${E(b.off)}</xdr:colOff><xdr:row>${d.i}</xdr:row><xdr:rowOff>${E(d.off)}</xdr:rowOff></xdr:to>` +
            `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${k + 2}" name="Chart ${k + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${Charts.NS_C}"><c:chart xmlns:c="${Charts.NS_C}" r:id="rId${k + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
        });
        extras.push({ path: `xl/drawings/drawing${dn}.xml`, data: XML + `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${NS_R}">${anchors}</xdr:wsDr>` });
        extras.push({ path: `xl/drawings/_rels/drawing${dn}.xml.rels`, data: XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${drRels}</Relationships>` });
        rels.push(`<Relationship Id="rId1" Type="${REL}/drawing" Target="../drawings/drawing${dn}.xml"/>`);
        cts.push(`<Override PartName="/xl/drawings/drawing${dn}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
        x += '<drawing r:id="rId1"/>';
      }
      if (sh.comments && sh.comments.length) {
        const cn = ++cmtN;
        const authors = [...new Set(sh.comments.map(c => c.author || 'Author'))];
        extras.push({ path: `xl/comments${cn}.xml`, data: XML + `<comments xmlns="${NS_M}"><authors>${authors.map(a => `<author>${xesc(a)}</author>`).join('')}</authors><commentList>` + sh.comments.map(c => `<comment ref="${cellName(c.r, c.c)}" authorId="${authors.indexOf(c.author || 'Author')}"><text><r><t xml:space="preserve">${xesc(c.text)}</t></r></text></comment>`).join('') + '</commentList></comments>' });
        cts.push(`<Override PartName="/xl/comments${cn}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`);
        const vml = '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="' + cn + '"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' +
          sh.comments.map((c, k) => `<v:shape id="_x0000_s${cn * 1024 + k + 1}" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:2pt;width:140pt;height:60pt;z-index:${k + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${c.c + 1}, 15, ${Math.max(0, c.r - 1)}, 2, ${c.c + 3}, 15, ${c.r + 3}, 16</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${c.r}</x:Row><x:Column>${c.c}</x:Column></x:ClientData></v:shape>`).join('') + '</xml>';
        extras.push({ path: `xl/drawings/vmlDrawing${cn}.vml`, data: vml });
        rels.push(`<Relationship Id="rId2" Type="${REL}/comments" Target="../comments${cn}.xml"/>`, `<Relationship Id="rId3" Type="${REL}/vmlDrawing" Target="../drawings/vmlDrawing${cn}.vml"/>`);
        x += '<legacyDrawing r:id="rId3"/>';
        hasVml = true;
      }
      if (tables.length) {
        x += `<tableParts count="${tables.length}">`;
        tables.forEach((t, k) => {
          const tn = ++tableN;
          const hdr = t.hdr !== false;
          const names = tableColNames(wb, sh, t);
          const ref = `${cellName(t.r1, t.c1)}:${cellName(t.r2, t.c2)}`;
          const afRef = `${cellName(t.r1, t.c1)}:${cellName(t.totals ? t.r2 - 1 : t.r2, t.c2)}`;
          const fc = (filterInTable === t && sh.filter.r1 === t.r1 && sh.filter.c1 === t.c1 && sh.filter.c2 === t.c2) ? filterCols(sh.filter) : '';
          const nm = xesc(t.name || ('Table' + tn));
          extras.push({ path: `xl/tables/table${tn}.xml`, data: XML + `<table xmlns="${NS_M}" id="${tn}" name="${nm}" displayName="${nm}" ref="${ref}"${hdr ? '' : ' headerRowCount="0"'}${t.totals ? ' totalsRowCount="1"' : ' totalsRowShown="0"'}>` +
            (hdr ? `<autoFilter ref="${afRef}">${fc}</autoFilter>` : '') +
            `<tableColumns count="${names.length}">` + names.map((n, i) => { const tc = (t.tcols || [])[i] || {}; return `<tableColumn id="${i + 1}" name="${xesc(n)}"${t.totals && tc.label ? ` totalsRowLabel="${xesc(tc.label)}"` : ''}${t.totals && tc.fn ? ` totalsRowFunction="${xesc(tc.fn)}"` : ''}/>`; }).join('') + '</tableColumns>' +
            `<tableStyleInfo name="${xesc(t.style || 'TableStyleMedium2')}" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>` });
          cts.push(`<Override PartName="/xl/tables/table${tn}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`);
          rels.push(`<Relationship Id="rId${10 + k}" Type="${REL}/table" Target="../tables/table${tn}.xml"/>`);
          x += `<tablePart r:id="rId${10 + k}"/>`;
        });
        x += '</tableParts>';
      }
      if (sh.sparks && sh.sparks.length) {
        const c8 = (c) => 'FF' + String(c || '376092').toUpperCase();
        x += '<extLst><ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' +
          sh.sparks.map(sp => {
            const f = sp.ref.indexOf('!') >= 0 ? sp.ref : Formula.quoteSheet(sh.name) + '!' + sp.ref;
            return `<x14:sparklineGroup displayEmptyCellsAs="gap"${sp.type === 'column' ? ' type="column"' : sp.type === 'stacked' ? ' type="stacked"' : ''}${sp.type === 'stacked' ? ' negative="1"' : ''}><x14:colorSeries rgb="${c8(sp.color)}"/><x14:colorNegative rgb="FFD00000"/><x14:colorAxis rgb="FF000000"/><x14:colorMarkers rgb="FFD00000"/><x14:colorFirst rgb="FFD00000"/><x14:colorLast rgb="FFD00000"/><x14:colorHigh rgb="FFD00000"/><x14:colorLow rgb="FFD00000"/><x14:sparklines><x14:sparkline><xm:f>${xesc(f)}</xm:f><xm:sqref>${cellName(sp.r, sp.c)}</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup>`;
          }).join('') + '</x14:sparklineGroups></ext></extLst>';
      }
      if (rels.length) extras.push({ path: `xl/worksheets/_rels/sheet${si + 1}.xml.rels`, data: XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>` });
      return x + '</worksheet>';
    });
    if (hasVml) cts.unshift('<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>');
    return { stylesXml, sheetXmls, sst, nStr, extras, cts };
  }

  async function write(wb) {
    const { stylesXml, sheetXmls, sst, nStr, extras, cts } = build(wb);
    const zw = Zip.writer();
    // refreshable PivotTable set-ups travel in a custom document property (Excel keeps it, other apps ignore it)
    const pivots = wb.sheets.filter(s => s.pivot).map(s => ({ sheet: s.name, def: { src: s.pivot.src, g: s.pivot.g, row: s.pivot.row, col: s.pivot.col, val: s.pivot.val, agg: s.pivot.agg, area: s.pivot.area } }));
    if (pivots.length) {
      extras.push({ path: 'docProps/custom.xml', data: XML + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="LumenPivots"><vt:lpwstr>' + xesc(JSON.stringify(pivots)) + '</vt:lpwstr></property></Properties>' });
      cts.push('<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>');
    }
    zw.add('[Content_Types].xml', XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      sheetXmls.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
      cts.join('') +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>');
    for (const e of extras) zw.add(e.path, e.data);
    zw.add('_rels/.rels', XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/>${pivots.length ? `<Relationship Id="rId4" Type="${REL}/custom-properties" Target="docProps/custom.xml"/>` : ''}</Relationships>`);
    zw.add('xl/workbook.xml', XML + `<workbook xmlns="${NS_M}" xmlns:r="${NS_R}"><workbookPr/><bookViews><workbookView activeTab="${wb.active}"/></bookViews><sheets>` +
      wb.sheets.map((s, i) => `<sheet name="${xesc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') + '</sheets>' + (Object.keys(wb.names || {}).length ? '<definedNames>' + Object.keys(wb.names).sort().map(k => `<definedName name="${xesc(k)}">${xesc(wb.names[k])}</definedName>`).join('') + '</definedNames>' : '') + '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>');
    const n = wb.sheets.length;
    zw.add('xl/_rels/workbook.xml.rels', XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      wb.sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') +
      `<Relationship Id="rId${n + 1}" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rId${n + 2}" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`);
    zw.add('xl/styles.xml', stylesXml);
    zw.add('xl/sharedStrings.xml', XML + `<sst xmlns="${NS_M}" count="${nStr}" uniqueCount="${sst.length}">` + sst.map(s => `<si><t xml:space="preserve">${xesc(s)}</t></si>`).join('') + '</sst>');
    sheetXmls.forEach((x, i) => zw.add(`xl/worksheets/sheet${i + 1}.xml`, x));
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    zw.add('docProps/core.xml', XML + `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Cells</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
    zw.add('docProps/app.xml', XML + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Cells</Application></Properties>');
    return zw.toBlob(XLSX_MIME);
  }

  g.XlsxIO = { read, write, XLSX_MIME };
})(window);
