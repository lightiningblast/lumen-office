/* Number / date format engine (Excel-style format codes) and input parsing. */
(function (g) {
  'use strict';
  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

  const BUILTIN = {
    0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
    14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yyyy h:mm',
    37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
    45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
  };
  const BUILTIN_REV = {};
  for (const k in BUILTIN) if (!(BUILTIN[k] in BUILTIN_REV)) BUILTIN_REV[BUILTIN[k]] = +k;

  /* ---------- serial dates ---------- */
  function serialToParts(serial) {
    let days = Math.floor(serial);
    let secs = Math.round((serial - days) * 86400);
    if (secs >= 86400) { days += 1; secs -= 86400; }
    let y, m, d, dow;
    if (days === 0) { y = 1900; m = 1; d = 0; dow = 6; }
    else if (days === 60) { y = 1900; m = 2; d = 29; dow = 4; }
    else {
      const base = days < 60 ? Date.UTC(1899, 11, 31) : Date.UTC(1899, 11, 30);
      const dt = new Date(base + days * 864e5);
      y = dt.getUTCFullYear(); m = dt.getUTCMonth() + 1; d = dt.getUTCDate(); dow = dt.getUTCDay();
    }
    return { y, m, d, dow, hh: Math.floor(secs / 3600), mm: Math.floor((secs % 3600) / 60), ss: secs % 60, days, secs };
  }
  function dateToSerial(y, m, d) {
    const t = Date.UTC(y, m - 1, d);
    let s = Math.round((t - Date.UTC(1899, 11, 30)) / 864e5);
    if (s <= 60) s -= 1;
    return s;
  }

  /* ---------- rounding ---------- */
  function shift(x, d) {
    const [m, e] = x.toExponential().split('e');
    return Number(m + 'e' + (+e + d));
  }
  function rnd(x, d) { // round half away from zero, decimal-faithful
    if (!isFinite(x)) return x;
    const neg = x < 0;
    x = Math.abs(x);
    let r = Math.round(shift(x, d));
    r = shift(r, -d);
    return neg ? -r : r;
  }
  function fixed(x, d) { // x >= 0
    const r = rnd(x, d);
    return r.toFixed(d);
  }

  function general(v) {
    if (v === 0) return '0';
    const a = Math.abs(v);
    if (a >= 1e11 || a < 1e-9) {
      let s = v.toExponential(5).replace(/\.?0+e/, 'e');
      return s.replace('e', 'E').replace(/E([+-])(\d)$/, 'E$10$2');
    }
    let s = String(+v.toPrecision(15));
    if (s.length > 11) {
      s = String(+v.toPrecision(10));
      if (s.indexOf('e') >= 0) s = String(v.toFixed(0));
      // keep at most 11 characters, Excel style
      if (s.length > 11 && s.indexOf('.') >= 0) s = s.slice(0, 11).replace(/\.$/, '');
    }
    return s;
  }

  /* ---------- code parsing ---------- */
  const cache = new Map();
  function parseCode(code) {
    if (cache.has(code)) return cache.get(code);
    // split into sections
    const secs = [];
    let cur = '', q = false, br = false;
    for (let i = 0; i < code.length; i++) {
      const ch = code[i];
      if (ch === '\\') { cur += ch + (code[i + 1] || ''); i++; continue; }
      if (ch === '"') q = !q;
      else if (!q && ch === '[') br = true;
      else if (!q && ch === ']') br = false;
      if (ch === ';' && !q && !br) { secs.push(cur); cur = ''; } else cur += ch;
    }
    secs.push(cur);
    const parsed = secs.map(parseSection);
    cache.set(code, parsed);
    return parsed;
  }
  function parseSection(s) {
    const tokens = [];
    let color = null, cond = null;
    const n = s.length;
    for (let i = 0; i < n; i++) {
      const ch = s[i];
      if (ch === '"') { let j = s.indexOf('"', i + 1); if (j < 0) j = n; tokens.push({ t: 'lit', v: s.slice(i + 1, j) }); i = j; continue; }
      if (ch === '\\') { tokens.push({ t: 'lit', v: s[i + 1] || '' }); i++; continue; }
      if (ch === '_') { tokens.push({ t: 'lit', v: ' ' }); i++; continue; }
      if (ch === '*') { i++; continue; }
      if (ch === '[') {
        const j = s.indexOf(']', i);
        const b = s.slice(i + 1, j < 0 ? n : j);
        i = j < 0 ? n : j;
        if (/^(red|blue|green|black|white|cyan|magenta|yellow)$/i.test(b)) color = b.toLowerCase();
        else if (/^(h+|m+|s+)$/i.test(b)) tokens.push({ t: 'el', v: b[0].toLowerCase(), n: b.length });
        else if (b[0] === '$') { const m = /^\$([^-]*)/.exec(b); if (m && m[1]) tokens.push({ t: 'lit', v: m[1] }); }
        else if (/^[<>=]/.test(b)) cond = b;
        continue;
      }
      const rest = s.slice(i);
      let m;
      if ((m = /^General/i.exec(rest))) { tokens.push({ t: 'gen' }); i += 6; continue; }
      if ((m = /^(AM\/PM|A\/P)/i.exec(rest))) { tokens.push({ t: 'ampm', v: m[1] }); i += m[1].length - 1; continue; }
      if ((m = /^E[+-]/i.exec(rest))) { tokens.push({ t: 'sci', v: m[0] }); i += 1; continue; }
      if (/[ymdhs]/i.test(ch)) {
        let j = i; while (j < n && s[j].toLowerCase() === ch.toLowerCase()) j++;
        const lt = ch.toLowerCase();
        tokens.push({ t: lt === 'd' ? 'day' : lt, n: j - i });
        i = j - 1; continue;
      }
      if (ch === '0' || ch === '#' || ch === '?') { tokens.push({ t: 'd', v: ch }); continue; }
      if (ch === '.') { tokens.push({ t: 'dot' }); continue; }
      if (ch === ',') { tokens.push({ t: 'comma' }); continue; }
      if (ch === '%') { tokens.push({ t: 'pct' }); continue; }
      if (ch === '@') { tokens.push({ t: 'at' }); continue; }
      tokens.push({ t: 'lit', v: ch });
    }
    const isDate = tokens.some(t => t.t === 'y' || t.t === 'day' || t.t === 'm' || t.t === 'h' || t.t === 's' || t.t === 'ampm' || t.t === 'el');
    return { tokens, color, cond, isDate, hasNum: tokens.some(t => t.t === 'd'), isText: tokens.some(t => t.t === 'at') };
  }

  function isDateFormat(code) {
    if (!code || code === 'General') return false;
    const secs = parseCode(code);
    return secs[0].isDate;
  }

  /* ---------- rendering ---------- */
  function renderDate(sec, serial) {
    const p = serialToParts(serial);
    const toks = sec.tokens;
    const ampm = toks.some(t => t.t === 'ampm');
    let out = '';
    for (let i = 0; i < toks.length; i++) {
      const t = toks[i];
      switch (t.t) {
        case 'lit': out += t.v; break;
        case 'comma': out += ','; break;
        case 'dot': out += '.'; break;
        case 'd': out += t.v === '0' ? '0' : ''; break;
        case 'y': out += t.n <= 2 ? String(p.y % 100).padStart(2, '0') : String(p.y).padStart(4, '0'); break;
        case 'm': {
          let prev = null, next = null;
          for (let k = i - 1; k >= 0; k--) if (toks[k].t !== 'lit') { prev = toks[k]; break; }
          for (let k = i + 1; k < toks.length; k++) if (toks[k].t !== 'lit') { next = toks[k]; break; }
          const minute = (prev && (prev.t === 'h' || (prev.t === 'el' && prev.v === 'h'))) || (next && (next.t === 's' || (next.t === 'el' && next.v === 's')));
          if (minute) out += t.n >= 2 ? String(p.mm).padStart(2, '0') : String(p.mm);
          else if (t.n === 1) out += p.m;
          else if (t.n === 2) out += String(p.m).padStart(2, '0');
          else if (t.n === 3) out += MONTHS[p.m - 1].slice(0, 3);
          else if (t.n === 4) out += MONTHS[p.m - 1];
          else out += MONTHS[p.m - 1][0];
          break;
        }
        case 'day':
          if (t.n === 1) out += p.d;
          else if (t.n === 2) out += String(p.d).padStart(2, '0');
          else if (t.n === 3) out += DAYS[p.dow].slice(0, 3);
          else out += DAYS[p.dow];
          break;
        case 'h': {
          let h = p.hh;
          if (ampm) { h = h % 12; if (h === 0) h = 12; }
          out += t.n >= 2 ? String(h).padStart(2, '0') : String(h);
          break;
        }
        case 's': out += t.n >= 2 ? String(p.ss).padStart(2, '0') : String(p.ss); break;
        case 'ampm': out += p.hh >= 12 ? (t.v.length === 1 ? 'P' : 'PM') : (t.v.length === 1 ? 'A' : 'AM'); break;
        case 'el': {
          const totalSecs = Math.round(serial * 86400);
          const v = t.v === 'h' ? Math.floor(totalSecs / 3600) : t.v === 'm' ? Math.floor(totalSecs / 60) : totalSecs;
          out += String(v).padStart(t.n, '0');
          break;
        }
        default: break;
      }
    }
    return out;
  }
  function renderNumber(sec, v, negPrefix) {
    const toks = sec.tokens;
    const hasSci = toks.some(t => t.t === 'sci');
    let pct = 0;
    toks.forEach(t => { if (t.t === 'pct') pct++; });
    // locate numeric tokens
    const dotIdx = toks.findIndex(t => t.t === 'dot');
    let firstD = -1, lastD = -1;
    toks.forEach((t, i) => { if (t.t === 'd') { if (firstD < 0) firstD = i; lastD = i; } });
    let scale = 0;
    if (lastD >= 0) { for (let i = lastD + 1; i < toks.length && toks[i].t === 'comma'; i++) scale++; }
    let grouping = false;
    if (firstD >= 0) for (let i = firstD; i < lastD; i++) if (toks[i].t === 'comma') grouping = true;
    const fracPl = [];
    const intPl = [];
    toks.forEach((t, i) => { if (t.t === 'd') (dotIdx >= 0 && i > dotIdx ? fracPl : intPl).push(i); });
    let x = Math.abs(v) * Math.pow(100, pct) / Math.pow(1000, scale);
    let intStr, fracStr, expStr = '';
    if (hasSci) {
      const sciIdx = toks.findIndex(t => t.t === 'sci');
      const mantFrac = fracPl.filter(i => i < sciIdx).length;
      const expDigits = toks.slice(sciIdx + 1).filter(t => t.t === 'd').length || 2;
      let e = x === 0 ? 0 : Math.floor(Math.log10(x));
      let m = x / Math.pow(10, e);
      if (rnd(m, mantFrac) >= 10) { m /= 10; e++; }
      const f = fixed(m, mantFrac).split('.');
      intStr = f[0]; fracStr = f[1] || '';
      expStr = (e < 0 ? '-' : (toks[sciIdx].v[1] === '+' ? '+' : '')) + String(Math.abs(e)).padStart(expDigits, '0');
    } else {
      const f = fixed(x, fracPl.length).split('.');
      intStr = f[0]; fracStr = f[1] || '';
    }
    const minInt = intPl.filter(i => toks[i].v === '0').length;
    if (intStr === '0' && minInt === 0 && intPl.length) intStr = '';
    if (intStr.length < minInt) intStr = intStr.padStart(minInt, '0');
    // '?' placeholders pad with spaces
    const qInt = intPl.filter(i => toks[i].v === '?').length;
    if (qInt && intStr.length < intPl.length) intStr = intStr.padStart(intPl.length, ' ');
    let intOut = intStr;
    if (grouping) intOut = intStr.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    // frac: strip trailing zeros for '#', pad '?' with spaces
    let fracOut = fracStr;
    for (let k = fracPl.length - 1; k >= 0; k--) {
      const ph = toks[fracPl[k]].v;
      if (ph === '0') break;
      if (fracOut[k] === '0') fracOut = fracOut.slice(0, k) + (ph === '?' ? ' ' : '') + fracOut.slice(k + 1);
      else break;
    }
    // contiguous check for int placeholders
    let contiguous = true;
    if (intPl.length > 1) for (let i = intPl[0]; i <= intPl[intPl.length - 1]; i++) if (toks[i].t !== 'd' && toks[i].t !== 'comma') contiguous = false;
    let out = '';
    let intDone = false, fracDone = false;
    for (let i = 0; i < toks.length; i++) {
      const t = toks[i];
      if (t.t === 'lit') out += t.v;
      else if (t.t === 'd') {
        if (dotIdx >= 0 && i > dotIdx) { if (!fracDone) { out += fracOut; fracDone = true; } }
        else if (contiguous) { if (!intDone) { out += intOut; intDone = true; } }
        else {
          const idx = intPl.indexOf(i);
          const fromRight = intPl.length - 1 - idx;
          const digit = intStr[intStr.length - 1 - fromRight];
          const extra = idx === 0 ? intStr.slice(0, Math.max(0, intStr.length - intPl.length)) : '';
          out += extra + (digit !== undefined ? digit : (t.v === '0' ? '0' : ''));
        }
      }
      else if (t.t === 'dot') { if (fracPl.length) out += '.'; }
      else if (t.t === 'pct') out += '%';
      else if (t.t === 'sci') out += (t.v[0]) + expStr;
      else if (t.t === 'comma') { /* grouping / scaling handled */ }
    }
    return negPrefix + out;
  }

  function format(v, code) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (!code || code === 'General') {
      return typeof v === 'number' ? general(v) : String(v);
    }
    const secs = parseCode(code);
    if (typeof v === 'string') {
      const sec = secs.find(s => s.isText) || null;
      if (!sec) return v;
      return sec.tokens.map(t => t.t === 'at' ? v : (t.t === 'lit' ? t.v : '')).join('');
    }
    if (!isFinite(v)) return String(v);
    let idx = 0;
    const evalCond = (c, val) => { const m = /^([<>=]+)\s*(-?[\d.]+)/.exec(c); if (!m) return false; const n = parseFloat(m[2]); switch (m[1]) { case '>': return val > n; case '<': return val < n; case '>=': return val >= n; case '<=': return val <= n; case '=': return val === n; case '<>': return val !== n; } return false; };
    let neg = v < 0;
    let usedNegSection = false;
    if (secs.length > 1 && (secs[0].cond || secs[1].cond)) {
      if (secs[0].cond && evalCond(secs[0].cond, v)) idx = 0;
      else if (secs[1].cond && evalCond(secs[1].cond, v)) { idx = 1; usedNegSection = neg; }
      else idx = secs.length > 2 ? 2 : (neg ? 1 : 0);
    } else if (secs.length === 1) idx = 0;
    else if (v > 0) idx = 0;
    else if (v < 0) { idx = 1; usedNegSection = true; }
    else idx = secs.length > 2 ? 2 : 0;
    const sec = secs[idx];
    if (!sec) return general(v);
    if (sec.tokens.length === 1 && sec.tokens[0].t === 'gen') return general(v);
    if (sec.isDate) return renderDate(sec, v < 0 ? 0 : v);
    if (sec.isText && !sec.hasNum) return sec.tokens.map(t => t.t === 'at' ? general(v) : (t.t === 'lit' ? t.v : '')).join('');
    if (!sec.hasNum) return sec.tokens.map(t => t.t === 'lit' ? t.v : '').join('');
    const prefix = neg && !usedNegSection ? '-' : '';
    return renderNumber(sec, v, prefix);
  }

  /* ---------- input parsing ---------- */
  function parseInput(str) {
    // returns { v, nf } for a typed literal (not a formula)
    const s = str.trim();
    if (s === '') return { v: undefined };
    if (/^(true|false)$/i.test(s)) return { v: s.toLowerCase() === 'true' };
    let m;
    if (/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(s)) return { v: parseFloat(s) };
    if (/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) { const dec = (s.split('.')[1] || '').length; return { v: parseFloat(s.replace(/,/g, '')), nf: dec ? '#,##0.' + '0'.repeat(dec) : '#,##0' }; }
    if ((m = /^([+-]?(?:\d+\.?\d*|\.\d+))%$/.exec(s))) { const dec = (m[1].split('.')[1] || '').length; return { v: parseFloat(m[1]) / 100, nf: dec ? '0.' + '0'.repeat(dec) + '%' : '0%' }; }
    if ((m = /^(-?)\$\s?([\d,]+(?:\.\d+)?)$/.exec(s))) { const num = parseFloat(m[2].replace(/,/g, '')); const dec = (m[2].split('.')[1] || '').length; return { v: m[1] ? -num : num, nf: '$#,##0.' + '0'.repeat(dec || 2) }; }
    if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s))) {
      const y = +m[1], mo = +m[2], d = +m[3];
      if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) {
        let v = dateToSerial(y, mo, d);
        if (m[4]) { v += (+m[4] * 3600 + +m[5] * 60 + +(m[6] || 0)) / 86400; return { v, nf: 'yyyy-mm-dd hh:mm' + (m[6] ? ':ss' : '') }; }
        return { v, nf: 'yyyy-mm-dd' };
      }
    }
    if ((m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(s))) {
      let y = +m[3]; if (y < 100) y += y < 30 ? 2000 : 1900;
      const mo = +m[1], d = +m[2];
      if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) return { v: dateToSerial(y, mo, d), nf: 'm/d/yyyy' };
    }
    if ((m = /^(\d{1,2})[- ]([A-Za-z]{3,9})[- ](\d{2,4})$/.exec(s))) {
      const mi = MONTHS.findIndex(n => n.toLowerCase().startsWith(m[2].toLowerCase().slice(0, 3)));
      let y = +m[3]; if (y < 100) y += y < 30 ? 2000 : 1900;
      if (mi >= 0) return { v: dateToSerial(y, mi + 1, +m[1]), nf: 'd-mmm-yy' };
    }
    if ((m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i.exec(s))) {
      let h = +m[1];
      if (m[4]) { h = h % 12 + (m[4].toUpperCase() === 'PM' ? 12 : 0); }
      if (h < 24 && +m[2] < 60) return { v: (h * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400, nf: m[4] ? (m[3] ? 'h:mm:ss AM/PM' : 'h:mm AM/PM') : (m[3] ? 'h:mm:ss' : 'h:mm') };
    }
    return { v: str };
  }

  g.NumFmt = {
    format,
    isDateFormat, parseInput, serialToParts, dateToSerial, BUILTIN, BUILTIN_REV, general, rnd, MONTHS, DAYS,
  };
})(window);
