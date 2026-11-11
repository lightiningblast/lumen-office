/* Conditional formatting engine for Lumen Cells.
   rule = { id, r1,c1,r2,c2, p (priority),
            type: 'cell' | 'text' | 'dup' | 'top' | 'scale' | 'bar' | 'expr',
            op, v1, v2        (cell: > < >= <= = <> between; values are literals or '=formula')
            text              (text: op = contains|begins|ends|notcontains)
            n, bottom         (top: top/bottom N)
            f                 (expr: formula relative to the top-left cell of the range)
            cfvo: [{t:'min'|'max'|'num'|'percent'|'percentile', v}], colors: ['RRGGBB',...]   (scale)
            color             (bar)
            style: { bg, color, b, i }                                                        } */
(function (g) {
  'use strict';
  const MAXC = Formula.MAXC;
  const PRESETS = {
    red: { bg: 'FFC7CE', color: '9C0006' }, yellow: { bg: 'FFEB9C', color: '9C5700' }, green: { bg: 'C6EFCE', color: '006100' },
    redfill: { bg: 'FFC7CE' }, redtext: { color: '9C0006' }, bold: { b: true },
  };
  const SCALES = {
    gyr: { cfvo: [{ t: 'min' }, { t: 'percentile', v: 50 }, { t: 'max' }], colors: ['63BE7B', 'FFEB84', 'F8696B'] },
    ryg: { cfvo: [{ t: 'min' }, { t: 'percentile', v: 50 }, { t: 'max' }], colors: ['F8696B', 'FFEB84', '63BE7B'] },
    wg: { cfvo: [{ t: 'min' }, { t: 'max' }], colors: ['FFFFFF', '63BE7B'] },
    wr: { cfvo: [{ t: 'min' }, { t: 'max' }], colors: ['FFFFFF', 'F8696B'] },
    wb: { cfvo: [{ t: 'min' }, { t: 'max' }], colors: ['FFFFFF', '5A8AC6'] },
  };
  const cache = { ver: -1, stats: new Map(), asts: new Map() };
  const num = (s) => { if (typeof s === 'number') return s; const t = String(s).trim(); const n = Number(t); return t !== '' && !isNaN(n) ? n : null; };
  const hex2 = (h) => [0, 2, 4].map(i => parseInt(h.slice(i, i + 2), 16));
  const toHex = (a) => a.map(x => ('0' + Math.round(Math.max(0, Math.min(255, x))).toString(16)).slice(-2)).join('').toUpperCase();

  function stat(wb, sh, rule) {
    if (cache.ver !== wb.ver) { cache.ver = wb.ver; cache.stats.clear(); cache.asts.clear(); }
    let s = cache.stats.get(rule);
    if (s) return s;
    const nums = [], cnt = new Map();
    for (const [k, cell] of sh.cells) {
      const r = Math.floor(k / MAXC), c = k - r * MAXC;
      if (r < rule.r1 || r > rule.r2 || c < rule.c1 || c > rule.c2) continue;
      const v = cell.f != null ? Formula.cellValue(wb, sh, r, c) : cell.v;
      if (typeof v === 'number') nums.push(v);
      if (v !== null && v !== undefined && v !== '' && !(v instanceof Formula.XErr)) { const key = String(v).toLowerCase(); cnt.set(key, (cnt.get(key) || 0) + 1); }
    }
    nums.sort((a, b) => a - b);
    s = { nums, min: nums[0], max: nums[nums.length - 1], cnt };
    cache.stats.set(rule, s);
    return s;
  }
  function evalFormula(wb, sh, text, rule, r, c) {
    const f = Formula.shiftFormula(text, r - rule.r1, c - rule.c1);
    let ast = cache.asts.get(f);
    if (ast === undefined) { try { ast = Formula.parse(f); } catch (e) { ast = null; } cache.asts.set(f, ast); }
    if (!ast) return null;
    let v = Formula.evalNode(ast, { wb, sh, r, c });
    if (v && v.isRef) v = Formula.cellValue(wb, v.sh, v.r1, v.c1);
    if (v && v.isArr) v = v.rows[0][0];
    return v;
  }
  function litVal(wb, sh, raw, rule, r, c) {
    if (raw === undefined || raw === null) return null;
    const s = String(raw);
    if (s[0] === '=') return evalFormula(wb, sh, s.slice(1), rule, r, c);
    const n = num(s);
    return n !== null ? n : s;
  }
  function pointValue(cf, st) {
    switch (cf.t) {
      case 'min': return st.min;
      case 'max': return st.max;
      case 'num': return Number(cf.v);
      case 'percent': return st.min + (st.max - st.min) * Number(cf.v) / 100;
      case 'percentile': { const a = st.nums; if (!a.length) return 0; const pos = Number(cf.v) / 100 * (a.length - 1); const lo = Math.floor(pos), hi = Math.ceil(pos); return a[lo] + (a[hi] - a[lo]) * (pos - lo); }
      default: return st.min;
    }
  }

  function match(wb, sh, rule, r, c, v) {
    switch (rule.type) {
      case 'cell': {
        if (v === null || v === undefined || v === '' || v instanceof Formula.XErr) return null;
        const a = litVal(wb, sh, rule.v1, rule, r, c), b = rule.op === 'between' ? litVal(wb, sh, rule.v2, rule, r, c) : null;
        let ok = false;
        if (typeof v === 'number' && typeof a === 'number') {
          switch (rule.op) {
            case '>': ok = v > a; break; case '<': ok = v < a; break; case '>=': ok = v >= a; break; case '<=': ok = v <= a; break;
            case '=': ok = v === a; break; case '<>': ok = v !== a; break;
            case 'between': ok = typeof b === 'number' && v >= Math.min(a, b) && v <= Math.max(a, b); break;
          }
        } else {
          const x = String(v).toLowerCase(), y = String(a === null ? '' : a).toLowerCase();
          switch (rule.op) { case '=': ok = x === y; break; case '<>': ok = x !== y; break; case '>': ok = x > y; break; case '<': ok = x < y; break; case '>=': ok = x >= y; break; case '<=': ok = x <= y; break; }
        }
        return ok ? { style: rule.style } : null;
      }
      case 'text': {
        if (v === null || v === undefined || v === '') return rule.op === 'notcontains' && false ? { style: rule.style } : null;
        const s = String(v).toLowerCase(), t = String(rule.text || '').toLowerCase();
        const ok = rule.op === 'begins' ? s.startsWith(t) : rule.op === 'ends' ? s.endsWith(t) : rule.op === 'notcontains' ? !s.includes(t) : s.includes(t);
        return ok ? { style: rule.style } : null;
      }
      case 'dup': {
        if (v === null || v === undefined || v === '') return null;
        return (stat(wb, sh, rule).cnt.get(String(v).toLowerCase()) || 0) > 1 ? { style: rule.style } : null;
      }
      case 'top': {
        if (typeof v !== 'number') return null;
        const a = stat(wb, sh, rule).nums;
        if (!a.length) return null;
        const n = Math.min(rule.n || 10, a.length);
        const thr = rule.bottom ? a[n - 1] : a[a.length - n];
        return (rule.bottom ? v <= thr : v >= thr) ? { style: rule.style } : null;
      }
      case 'scale': {
        if (typeof v !== 'number') return null;
        const st = stat(wb, sh, rule);
        if (!st.nums.length) return null;
        const pts = rule.cfvo.map(p => pointValue(p, st));
        const cols = rule.colors.map(hex2);
        let col;
        if (v <= pts[0]) col = cols[0];
        else if (v >= pts[pts.length - 1]) col = cols[cols.length - 1];
        else {
          let i = 0; while (i < pts.length - 2 && v > pts[i + 1]) i++;
          const t = pts[i + 1] === pts[i] ? 0 : (v - pts[i]) / (pts[i + 1] - pts[i]);
          col = cols[i].map((x, k) => x + (cols[i + 1][k] - x) * t);
        }
        return { style: { bg: toHex(col) } };
      }
      case 'bar': {
        if (typeof v !== 'number') return null;
        const st = stat(wb, sh, rule);
        if (!st.nums.length) return null;
        const lo = Math.min(0, st.min), hi = Math.max(st.max, lo + 1e-12);
        return { bar: { pct: Math.max(0, Math.min(1, (v - lo) / (hi - lo))), color: rule.color || '638EC6' } };
      }
      case 'expr': {
        const res = evalFormula(wb, sh, rule.f, rule, r, c);
        const t = res === true || (typeof res === 'number' && res !== 0) || (typeof res === 'string' && /^true$/i.test(res));
        return t ? { style: rule.style } : null;
      }
    }
    return null;
  }

  function evaluate(wb, sh, r, c, cell) {
    if (!sh.cf || !sh.cf.length) return null;
    let v;
    let ov = null;
    for (const rule of sh.cf) {
      if (r < rule.r1 || r > rule.r2 || c < rule.c1 || c > rule.c2) continue;
      if (v === undefined) v = cell ? (cell.f != null ? Formula.cellValue(wb, sh, r, c) : (cell.v === undefined ? null : cell.v)) : null;
      const res = match(wb, sh, rule, r, c, v);
      if (!res) continue;
      ov = ov || { style: {}, bar: null };
      if (res.style) for (const k in res.style) if (!(k in ov.style)) ov.style[k] = res.style[k];
      if (res.bar && !ov.bar) ov.bar = res.bar;
    }
    return ov;
  }

  const OPN = { '>': 'greater than', '<': 'less than', '>=': '>=', '<=': '<=', '=': 'equal to', '<>': 'not equal to', between: 'between' };
  function describe(rule) {
    const A = cellRange(rule);
    switch (rule.type) {
      case 'cell': return `${A}: cell value ${OPN[rule.op]} ${rule.v1}${rule.op === 'between' ? ' and ' + rule.v2 : ''}`;
      case 'text': return `${A}: text ${rule.op === 'notcontains' ? 'does not contain' : rule.op === 'begins' ? 'begins with' : rule.op === 'ends' ? 'ends with' : 'contains'} "${rule.text}"`;
      case 'dup': return `${A}: duplicate values`;
      case 'top': return `${A}: ${rule.bottom ? 'bottom' : 'top'} ${rule.n}`;
      case 'scale': return `${A}: ${rule.colors.length}-color scale`;
      case 'bar': return `${A}: data bars`;
      case 'expr': return `${A}: formula =${rule.f}`;
    }
    return A;
  }
  function cellRange(r) { const a = Formula.cellName(r.r1, r.c1); return r.r1 === r.r2 && r.c1 === r.c2 ? a : a + ':' + Formula.cellName(r.r2, r.c2); }

  g.CondFmt = { evaluate, describe, PRESETS, SCALES, cellRange };
})(window);
