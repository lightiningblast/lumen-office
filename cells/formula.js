/* Workbook model + formula engine for Lumen Cells. */
(function (g) {
  'use strict';
  const MAXR = 1048576, MAXC = 16384;

  class XErr { constructor(code) { this.code = code; } toString() { return this.code; } }
  const E = { DIV0: new XErr('#DIV/0!'), VALUE: new XErr('#VALUE!'), REF: new XErr('#REF!'), NAME: new XErr('#NAME?'), NA: new XErr('#N/A'), NUM: new XErr('#NUM!'), NULL: new XErr('#NULL!') };
  const ERR_BY_CODE = {};
  for (const k in E) ERR_BY_CODE[E[k].code] = E[k];

  /* ---------- addressing ---------- */
  function colName(c) { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }
  function colIndex(s) { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; }
  function cellName(r, c) { return colName(c) + (r + 1); }
  function parseCell(s) { const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(s); return m ? { r: +m[2] - 1, c: colIndex(m[1]) } : null; }
  function parseRange(s) {
    s = s.replace(/\$/g, '');
    const p = s.split(':');
    const a = parseCell(p[0]);
    if (!a) return null;
    const b = p[1] ? parseCell(p[1]) : a;
    if (!b) return null;
    return { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) };
  }

  /* ---------- model ---------- */
  class Sheet {
    constructor(name) {
      this.name = name; this.cells = new Map(); this.colW = new Map(); this.rowH = new Map();
      this.defColW = 64; this.defRowH = 20; this.merges = []; this.freeze = { r: 0, c: 0 }; this.showGrid = true;
      this.charts = []; this.cf = []; this.dv = []; this.comments = []; this.filter = null; this.tables = []; this.sparks = [];
      this.sel = { r1: 0, c1: 0, r2: 0, c2: 0, ar: 0, ac: 0 }; this.scroll = { x: 0, y: 0 };
    }
    key(r, c) { return r * MAXC + c; }
    get(r, c) { return this.cells.get(r * MAXC + c); }
    bounds() {
      let mr = -1, mc = -1;
      for (const k of this.cells.keys()) { const r = Math.floor(k / MAXC), c = k - r * MAXC; if (r > mr) mr = r; if (c > mc) mc = c; }
      return { maxR: mr, maxC: mc };
    }
    mergeAt(r, c) { for (const m of this.merges) if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return m; return null; }
  }

  class Workbook {
    constructor() {
      this.sheets = [new Sheet('Sheet1')]; this.styles = [{}]; this._sk = new Map([['{}', 0]]);
      this.names = {}; this._nm = new Map();
      this.ver = 1; this.active = 0; this._bver = -1; this._bcache = new Map();
    }
    styleId(o) {
      const clean = {};
      Object.keys(o).sort().forEach(k => { const v = o[k]; if (v !== undefined && v !== null && v !== false && v !== '') clean[k] = v; });
      const k = JSON.stringify(clean);
      let id = this._sk.get(k);
      if (id === undefined) { id = this.styles.length; this.styles.push(clean); this._sk.set(k, id); }
      return id;
    }
    sheetByName(n) { n = String(n).toLowerCase(); return this.sheets.find(s => s.name.toLowerCase() === n); }
    touch() { this.ver++; }
    bounds(sh) {
      if (this._bver !== this.ver) { this._bcache.clear(); this._bver = this.ver; }
      let b = this._bcache.get(sh);
      if (!b) { b = sh.bounds(); this._bcache.set(sh, b); }
      return b;
    }
  }

  /* ---------- reference rewriting ---------- */
  const REF_SRC = "(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!)?(\\$?[A-Za-z]{1,3}\\$?\\d+(?::\\$?[A-Za-z]{1,3}\\$?\\d+)?|\\$?[A-Za-z]{1,3}:\\$?[A-Za-z]{1,3}(?![A-Za-z0-9_(])|\\$?\\d+:\\$?\\d+)(?![A-Za-z0-9_(])";
  function splitRef(txt) {
    // txt like $A$1 or A1:B2 or A:A or 1:3
    const parts = txt.split(':');
    const one = (p) => {
      let m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(p);
      if (m) return { ac: !!m[1], c: colIndex(m[2]), ar: !!m[3], r: +m[4] - 1 };
      m = /^(\$?)([A-Za-z]{1,3})$/.exec(p);
      if (m) return { ac: !!m[1], c: colIndex(m[2]), col: true };
      m = /^(\$?)(\d+)$/.exec(p);
      if (m) return { ar: !!m[1], r: +m[2] - 1, row: true };
      return null;
    };
    const a = one(parts[0]);
    const b = parts[1] ? one(parts[1]) : null;
    return { a, b };
  }
  function fmtPart(p) {
    if (p.col) return (p.ac ? '$' : '') + colName(p.c);
    if (p.row) return (p.ar ? '$' : '') + (p.r + 1);
    return (p.ac ? '$' : '') + colName(p.c) + (p.ar ? '$' : '') + (p.r + 1);
  }
  function quoteSheet(n) {
    if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n)) return n;
    return "'" + n.replace(/'/g, "''") + "'";
  }
  // fn(ref) -> new ref object {a,b} / null (=> #REF!) / undefined (keep)
  function rewriteRefs(formula, fn) {
    // skip string literals
    const out = [];
    const re = /"(?:[^"]|"")*"/g;
    let last = 0, m;
    const process = (seg) => seg.replace(new RegExp('(?<![A-Za-z0-9_.$!])' + REF_SRC, 'g'), (all, q, u, body) => {
      const sheet = q ? q.replace(/''/g, "'") : (u || null);
      const sp = splitRef(body);
      if (!sp.a) return all;
      const res = fn({ sheet, a: sp.a, b: sp.b });
      if (res === undefined) return all;
      if (res === null) return '#REF!';
      let prefix = all.slice(0, all.length - body.length);
      if (res.sheet !== undefined) prefix = res.sheet ? quoteSheet(res.sheet) + '!' : '';
      return prefix + fmtPart(res.a) + (res.b ? ':' + fmtPart(res.b) : '');
    });
    while ((m = re.exec(formula))) { out.push(process(formula.slice(last, m.index)), m[0]); last = m.index + m[0].length; }
    out.push(process(formula.slice(last)));
    return out.join('');
  }
  function shiftFormula(f, dr, dc) {
    return rewriteRefs(f, ref => {
      const mv = (p) => {
        const q = Object.assign({}, p);
        if (!p.row && !p.ac) q.c = p.c + dc;
        if (!p.col && !p.ar) q.r = p.r + dr;
        if ((q.c !== undefined && (q.c < 0 || q.c >= MAXC)) || (q.r !== undefined && (q.r < 0 || q.r >= MAXR))) return null;
        return q;
      };
      const a = mv(ref.a), b = ref.b ? mv(ref.b) : null;
      if (!a || (ref.b && !b)) return null;
      return { a, b };
    });
  }

  /* ---------- tokenizer / parser ---------- */
  const RE_FN = /([A-Za-z_][A-Za-z0-9_.]*)\(/y;
  const RE_REF = new RegExp(REF_SRC + '', 'y');
  const RE_NUM = /(\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/y;
  const RE_BOOL = /(TRUE|FALSE)(?![A-Za-z0-9_(])/iy;
  const RE_ERR = /#(?:DIV\/0!|VALUE!|REF!|NAME\?|N\/A|NUM!|NULL!)/y;
  const RE_NAME = /[A-Za-z_][A-Za-z0-9_.]*/y;

  function tokenize(src) {
    const t = [];
    let i = 0;
    const n = src.length;
    const tryRe = (re) => { re.lastIndex = i; return re.exec(src); };
    while (i < n) {
      const ch = src[i];
      if (/\s/.test(ch)) { i++; continue; }
      if (ch === '"') {
        let j = i + 1, s = '';
        while (j < n) { if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; } s += src[j++]; }
        t.push({ k: 'str', v: s }); i = j + 1; continue;
      }
      let m;
      if ((m = tryRe(RE_FN))) { t.push({ k: 'fn', v: m[1].toUpperCase() }); i += m[0].length; t.push({ k: '(' }); continue; }
      if ((m = tryRe(RE_REF))) {
        const sp = splitRef(m[3]);
        if (sp.a) { t.push({ k: 'ref', sheet: m[1] ? m[1].replace(/''/g, "'") : (m[2] || null), sp }); i += m[0].length; continue; }
      }
      if ((m = tryRe(RE_BOOL))) { t.push({ k: 'bool', v: m[1].toUpperCase() === 'TRUE' }); i += m[0].length; continue; }
      if ((m = tryRe(RE_NUM))) { t.push({ k: 'num', v: parseFloat(m[0]) }); i += m[0].length; continue; }
      if ((m = tryRe(RE_ERR))) { t.push({ k: 'err', v: m[0] }); i += m[0].length; continue; }
      if ((m = tryRe(RE_NAME))) { t.push({ k: 'name', v: m[0] }); i += m[0].length; continue; }
      const two = src.substr(i, 2);
      if (two === '<>' || two === '<=' || two === '>=') { t.push({ k: 'op', v: two }); i += 2; continue; }
      if ('+-*/^&=<>%(),:;'.includes(ch)) { t.push({ k: ch === '(' || ch === ')' || ch === ',' || ch === ';' ? (ch === ';' ? ',' : ch) : 'op', v: ch }); i++; continue; }
      throw new Error('Unexpected character ' + ch);
    }
    return t;
  }

  function parse(src) {
    const toks = tokenize(src);
    let p = 0;
    const peek = () => toks[p];
    const isOp = (v) => toks[p] && toks[p].k === 'op' && toks[p].v === v;
    function cmp() {
      let l = concat();
      while (peek() && peek().k === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(peek().v)) { const op = toks[p++].v; l = { t: 'bin', op, l, r: concat() }; }
      return l;
    }
    function concat() { let l = add(); while (isOp('&')) { p++; l = { t: 'bin', op: '&', l, r: add() }; } return l; }
    function add() { let l = mul(); while (isOp('+') || isOp('-')) { const op = toks[p++].v; l = { t: 'bin', op, l, r: mul() }; } return l; }
    function mul() { let l = pow(); while (isOp('*') || isOp('/')) { const op = toks[p++].v; l = { t: 'bin', op, l, r: pow() }; } return l; }
    function pow() { let l = unary(); while (isOp('^')) { p++; l = { t: 'bin', op: '^', l, r: unary() }; } return l; }
    function unary() {
      if (isOp('-')) { p++; return { t: 'un', op: '-', e: unary() }; }
      if (isOp('+')) { p++; return unary(); }
      return postfix();
    }
    function postfix() { let e = primary(); while (isOp('%')) { p++; e = { t: 'pct', e }; } return e; }
    function primary() {
      const tk = toks[p++];
      if (!tk) throw new Error('Unexpected end');
      switch (tk.k) {
        case 'num': return { t: 'num', v: tk.v };
        case 'str': return { t: 'str', v: tk.v };
        case 'bool': return { t: 'bool', v: tk.v };
        case 'err': return { t: 'err', v: ERR_BY_CODE[tk.v] || E.VALUE };
        case 'ref': {
          let a = tk.sp.a, b = tk.sp.b || tk.sp.a;
          // range via ':' between two refs
          if (isOp(':') && toks[p + 1] && toks[p + 1].k === 'ref' && !tk.sp.b) { p++; const t2 = toks[p++]; b = t2.sp.a; }
          let r1 = a.row || a.r !== undefined ? a.r : 0, r2 = b.r !== undefined ? b.r : MAXR - 1;
          let c1 = a.c !== undefined ? a.c : 0, c2 = b.c !== undefined ? b.c : MAXC - 1;
          if (a.col) { r1 = 0; } if (b.col) { r2 = MAXR - 1; }
          if (a.row) { c1 = 0; } if (b.row) { c2 = MAXC - 1; }
          return { t: 'ref', sheet: tk.sheet, r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: Math.min(c1, c2), c2: Math.max(c1, c2) };
        }
        case 'name': return { t: 'name', v: tk.v };
        case 'fn': {
          p++; // (
          const args = [];
          if (toks[p] && toks[p].k === ')') p++;
          else {
            for (;;) {
              if (toks[p] && (toks[p].k === ',' || toks[p].k === ')')) args.push({ t: 'empty' });
              else args.push(cmp());
              if (toks[p] && toks[p].k === ',') { p++; continue; }
              if (toks[p] && toks[p].k === ')') { p++; break; }
              throw new Error('Expected , or )');
            }
          }
          return { t: 'fn', name: tk.v, args };
        }
        case '(': { const e = cmp(); if (!toks[p] || toks[p].k !== ')') throw new Error('Missing )'); p++; return e; }
        default: throw new Error('Unexpected token');
      }
    }
    const ast = cmp();
    if (p < toks.length) throw new Error('Unexpected token');
    return ast;
  }

  /* ---------- evaluation helpers ---------- */
  const isErr = (v) => v instanceof XErr;
  const isRef = (v) => v && v.isRef;
  const isArr = (v) => v && v.isArr;

  const SPILL_RE = /^\s*(UNIQUE|SORT|SORTBY|FILTER|SEQUENCE|RANDARRAY|TRANSPOSE|VSTACK|HSTACK|TEXTSPLIT|TAKE|DROP|MMULT|TOCOL|TOROW)\s*\(|^\s*[$A-Za-z]+\$?\d+:\$?[A-Za-z]+\$?\d+\s*([-+*\/^&][^()]*)?$/i;
  // Formula cells whose result is a multi-cell array ("spill"): [{ r, c, rows }]
  function spills(wb, sh) {
    if (sh._spv === wb.ver && sh._arrs) return sh._arrs;
    if (sh._spb) return [];
    sh._spb = true;
    const out = [];
    try {
      for (const [k, cell] of sh.cells) {
        if (cell.f == null || !SPILL_RE.test(cell.f)) continue;
        const r = Math.floor(k / MAXC), c = k - r * MAXC;
        cellValue(wb, sh, r, c);
        if (cell._arr) out.push({ r, c, rows: cell._arr });
      }
    } finally { sh._spb = false; }
    sh._spv = wb.ver;
    sh._arrs = out;
    return out;
  }
  function spillAt(wb, sh, r, c) {
    for (const s of spills(wb, sh)) {
      if (r < s.r || c < s.c || (r === s.r && c === s.c)) continue;
      if (r - s.r < s.rows.length && c - s.c < s.rows[0].length) { const v = s.rows[r - s.r][c - s.c]; return v === undefined ? null : v; }
    }
    return null;
  }
  function cellValue(wb, sh, r, c, stack) {
    const cell = sh.get(r, c);
    if (!cell) return sh._spb ? null : spillAt(wb, sh, r, c);
    if (cell.f == null) return cell.v === undefined ? null : cell.v;
    if (cell._ver === wb.ver) return cell._val;
    if (cell._busy) return E.REF;
    cell._busy = true;
    let val;
    try {
      if (!cell._ast || cell._astF !== cell.f) { cell._ast = parse(cell.f); cell._astF = cell.f; }
      val = evalNode(cell._ast, { wb, sh, r, c });
      if (isRef(val)) { if (val.r1 === val.r2 && val.c1 === val.c2) val = refScalar(val, { wb }); else val = toArr(val, { wb }); }
      cell._arr = null;
      if (isArr(val)) { if (val.rows.length > 1 || val.rows[0].length > 1) cell._arr = val.rows; val = val.rows[0][0]; }
      if (val === undefined) val = null;
      if (isErr(val) && val === E.NAME && cell.cv !== undefined && cell.cv !== null) val = cell.cv;
    } catch (e) { val = E.NAME; if (cell.cv !== undefined && cell.cv !== null) val = cell.cv; }
    cell._busy = false;
    cell._ver = wb.ver;
    cell._val = val;
    return val;
  }
  function refScalar(ref, ctx) { return cellValue(ctx.wb, ref.sh, ref.r1, ref.c1); }
  function toArr(x, ctx) {
    if (isArr(x)) return x;
    if (isRef(x)) {
      const b = ctx.wb.bounds(x.sh);
      const r2 = Math.min(x.r2, Math.max(x.r1, b.maxR)), c2 = Math.min(x.c2, Math.max(x.c1, b.maxC));
      const rows = [];
      for (let r = x.r1; r <= r2; r++) { const row = []; for (let c = x.c1; c <= c2; c++) row.push(cellValue(ctx.wb, x.sh, r, c)); rows.push(row); }
      return { isArr: true, rows };
    }
    return { isArr: true, rows: [[x]] };
  }
  function scalar(x, ctx) {
    if (isRef(x)) return refScalar(x, ctx);
    if (isArr(x)) return x.rows[0][0];
    return x;
  }
  function numFrom(v) { // strict coercion for arithmetic
    if (typeof v === 'number') return v;
    if (v === null || v === undefined) return 0;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (isErr(v)) return v;
    const s = String(v).trim();
    if (s === '') return E.VALUE;
    const pi = NumFmt.parseInput(s);
    if (typeof pi.v === 'number') return pi.v;
    return E.VALUE;
  }
  function str(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') return NumFmt.general(v).replace(/^(-?\d+\.\d*?)0+$/, '$1');
    return String(v);
  }
  function lift(a, b, f, ctx) {
    if (isRef(a)) a = toArr(a, ctx);
    if (isRef(b)) b = toArr(b, ctx);
    if (isArr(a) || isArr(b)) {
      const A = isArr(a) ? a.rows : null, B = isArr(b) ? b.rows : null;
      const R = Math.max(A ? A.length : 1, B ? B.length : 1), C = Math.max(A ? A[0].length : 1, B ? B[0].length : 1);
      const rows = [];
      for (let r = 0; r < R; r++) {
        const row = [];
        for (let c = 0; c < C; c++) {
          const x = A ? (A[r] && A[r][c] !== undefined ? A[r][c] : (A.length === 1 || A[0].length === 1 ? A[Math.min(r, A.length - 1)][Math.min(c, A[0].length - 1)] : E.NA)) : a;
          const y = B ? (B[r] && B[r][c] !== undefined ? B[r][c] : (B.length === 1 || B[0].length === 1 ? B[Math.min(r, B.length - 1)][Math.min(c, B[0].length - 1)] : E.NA)) : b;
          row.push(f(x, y));
        }
        rows.push(row);
      }
      return { isArr: true, rows };
    }
    return f(a, b);
  }
  function compare(x, y) {
    const rank = (v) => v === null || v === undefined ? 0 : typeof v === 'number' ? 1 : typeof v === 'string' ? 2 : 3;
    if (x === null || x === undefined) x = typeof y === 'string' ? '' : (typeof y === 'boolean' ? false : 0);
    if (y === null || y === undefined) y = typeof x === 'string' ? '' : (typeof x === 'boolean' ? false : 0);
    const rx = rank(x), ry = rank(y);
    if (rx !== ry) return rx < ry ? -1 : 1;
    if (rx === 2) { const a = x.toLowerCase(), b = y.toLowerCase(); return a < b ? -1 : a > b ? 1 : 0; }
    return x < y ? -1 : x > y ? 1 : 0;
  }
  function binop(op, x, y) {
    if (isErr(x)) return x;
    if (isErr(y)) return y;
    if (op === '&') return str(x) + str(y);
    if (['=', '<>', '<', '>', '<=', '>='].includes(op)) {
      const c = compare(x, y);
      switch (op) { case '=': return c === 0; case '<>': return c !== 0; case '<': return c < 0; case '>': return c > 0; case '<=': return c <= 0; default: return c >= 0; }
    }
    const a = numFrom(x), b = numFrom(y);
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    switch (op) {
      case '+': return a + b;
      case '-': return a - b;
      case '*': return a * b;
      case '/': return b === 0 ? E.DIV0 : a / b;
      case '^': { const r = Math.pow(a, b); return isNaN(r) ? E.NUM : r; }
    }
    return E.VALUE;
  }

  function evalNode(n, ctx) {
    switch (n.t) {
      case 'num': case 'str': case 'bool': return n.v;
      case 'err': return n.v;
      case 'empty': return null;
      case 'name': {
        const txt = ctx.wb.names && ctx.wb.names[n.v.toUpperCase()];
        if (!txt) return E.NAME;
        let ast = ctx.wb._nm.get(txt);
        if (!ast) { try { ast = parse(txt); } catch (e) { return E.NAME; } ctx.wb._nm.set(txt, ast); }
        return evalNode(ast, ctx);
      }
      case 'ref': {
        const sh = n.sheet ? ctx.wb.sheetByName(n.sheet) : ctx.sh;
        if (!sh) return E.REF;
        return { isRef: true, sh, r1: n.r1, c1: n.c1, r2: n.r2, c2: n.c2 };
      }
      case 'un': { const v = evalNode(n.e, ctx); return lift(v, 0, (a) => { if (isErr(a)) return a; const x = numFrom(a); return isErr(x) ? x : -x; }, ctx); }
      case 'pct': { const v = evalNode(n.e, ctx); return lift(v, 0, (a) => { const x = numFrom(a); return isErr(x) ? x : x / 100; }, ctx); }
      case 'bin': {
        const a = evalNode(n.l, ctx), b = evalNode(n.r, ctx);
        const sa = (isRef(a) && a.r1 === a.r2 && a.c1 === a.c2) ? refScalar(a, ctx) : a;
        const sb = (isRef(b) && b.r1 === b.r2 && b.c1 === b.c2) ? refScalar(b, ctx) : b;
        return lift(sa, sb, (x, y) => binop(n.op, x, y), ctx);
      }
      case 'fn': return callFn(n, ctx);
    }
    return E.VALUE;
  }

  /* ---------- function library ---------- */
  const F = {};   // eager: (args(raw), ctx) => value
  const L = {};   // lazy:  (argNodes, ctx) => value

  function callFn(n, ctx) {
    if (L[n.name]) return L[n.name](n.args, ctx);
    const f = F[n.name];
    if (!f) return E.NAME;
    const args = n.args.map(a => evalNode(a, ctx));
    try { return f(args, ctx); } catch (e) { return E.VALUE; }
  }
  // flatten args to numbers (ranges ignore text/bool/blank; direct args coerce)
  function nums(args, ctx, opts) {
    opts = opts || {};
    const out = [];
    for (const a of args) {
      if (isRef(a) || isArr(a)) {
        for (const row of toArr(a, ctx).rows) for (const v of row) {
          if (isErr(v)) return v;
          if (typeof v === 'number') out.push(v);
          else if (opts.countText && v !== null && v !== undefined) out.push(typeof v === 'boolean' ? +v : 0);
        }
      } else {
        if (isErr(a)) return a;
        if (a === null && opts.skipEmptyDirect) continue;
        const x = (typeof a === 'string' && a.trim() === '') ? E.VALUE : numFrom(a);
        if (isErr(x)) return x;
        out.push(x);
      }
    }
    return out;
  }
  const all = (args, ctx) => { const out = []; for (const a of args) { if (isRef(a) || isArr(a)) toArr(a, ctx).rows.forEach(r => r.forEach(v => out.push(v))); else out.push(a); } return out; };
  const num1 = (v, ctx) => { v = scalar(v, ctx); return numFrom(v); };
  const mathFn = (f) => (args, ctx) => { const x = num1(args[0], ctx); if (isErr(x)) return x; const r = f(x, args, ctx); return (typeof r === 'number' && !isFinite(r)) ? E.NUM : r; };
  const reduceNums = (f, empty) => (args, ctx) => { const v = nums(args, ctx); if (isErr(v)) return v; return v.length ? f(v) : empty; };

  F.SUM = reduceNums(v => v.reduce((a, b) => a + b, 0), 0);
  F.AVERAGE = reduceNums(v => v.reduce((a, b) => a + b, 0) / v.length, E.DIV0);
  F.MIN = reduceNums(v => Math.min(...v), 0);
  F.MAX = reduceNums(v => Math.max(...v), 0);
  F.PRODUCT = reduceNums(v => v.reduce((a, b) => a * b, 1), 0);
  F.SUMSQ = reduceNums(v => v.reduce((a, b) => a + b * b, 0), 0);
  F.MEDIAN = reduceNums(v => { v = v.slice().sort((a, b) => a - b); const m = v.length >> 1; return v.length % 2 ? v[m] : (v[m - 1] + v[m]) / 2; }, E.NUM);
  F.MODE = reduceNums(v => { const c = new Map(); let best = null, bc = 1; v.forEach(x => { const k = (c.get(x) || 0) + 1; c.set(x, k); if (k > bc) { bc = k; best = x; } }); return best === null ? E.NA : best; }, E.NA);
  F['MODE.SNGL'] = F.MODE;
  const varp = (v) => { const m = v.reduce((a, b) => a + b, 0) / v.length; return v.reduce((a, b) => a + (b - m) * (b - m), 0); };
  F.VAR = F['VAR.S'] = reduceNums(v => v.length > 1 ? varp(v) / (v.length - 1) : E.DIV0, E.DIV0);
  F['VAR.P'] = F.VARP = reduceNums(v => varp(v) / v.length, E.DIV0);
  F.STDEV = F['STDEV.S'] = reduceNums(v => v.length > 1 ? Math.sqrt(varp(v) / (v.length - 1)) : E.DIV0, E.DIV0);
  F['STDEV.P'] = F.STDEVP = reduceNums(v => Math.sqrt(varp(v) / v.length), E.DIV0);
  F.COUNT = (args, ctx) => { let n = 0; for (const a of args) { if (isRef(a) || isArr(a)) toArr(a, ctx).rows.forEach(r => r.forEach(v => { if (typeof v === 'number') n++; })); else if (typeof a === 'number' || (typeof a === 'string' && !isErr(numFrom(a)))) n++; } return n; };
  F.COUNTA = (args, ctx) => all(args, ctx).filter(v => v !== null && v !== undefined).length;
  F.COUNTBLANK = (args, ctx) => all(args, ctx).filter(v => v === null || v === undefined || v === '').length;
  F.LARGE = (args, ctx) => { const v = nums([args[0]], ctx); if (isErr(v)) return v; const k = num1(args[1], ctx); v.sort((a, b) => b - a); return v[k - 1] === undefined ? E.NUM : v[k - 1]; };
  F.SMALL = (args, ctx) => { const v = nums([args[0]], ctx); if (isErr(v)) return v; const k = num1(args[1], ctx); v.sort((a, b) => a - b); return v[k - 1] === undefined ? E.NUM : v[k - 1]; };
  F.RANK = F['RANK.EQ'] = (args, ctx) => { const x = num1(args[0], ctx); const v = nums([args[1]], ctx); if (isErr(v)) return v; const asc = args[2] !== undefined && num1(args[2], ctx) !== 0; if (!v.includes(x)) return E.NA; return 1 + v.filter(y => asc ? y < x : y > x).length; };
  F.PERCENTILE = (args, ctx) => { const v = nums([args[0]], ctx).sort((a, b) => a - b); const k = num1(args[1], ctx); const pos = k * (v.length - 1); const lo = Math.floor(pos), hi = Math.ceil(pos); return v[lo] + (v[hi] - v[lo]) * (pos - lo); };
  F.ABS = mathFn(Math.abs);
  F.INT = mathFn(Math.floor);
  F.SQRT = mathFn(x => x < 0 ? E.NUM : Math.sqrt(x));
  F.EXP = mathFn(Math.exp);
  F.LN = mathFn(x => x <= 0 ? E.NUM : Math.log(x));
  F.LOG10 = mathFn(x => x <= 0 ? E.NUM : Math.log10(x));
  F.LOG = mathFn((x, a, ctx) => { if (x <= 0) return E.NUM; const b = a[1] !== undefined ? num1(a[1], ctx) : 10; return Math.log(x) / Math.log(b); });
  F.SIN = mathFn(Math.sin); F.COS = mathFn(Math.cos); F.TAN = mathFn(Math.tan); F.ASIN = mathFn(Math.asin); F.ACOS = mathFn(Math.acos); F.ATAN = mathFn(Math.atan);
  F.DEGREES = mathFn(x => x * 180 / Math.PI); F.RADIANS = mathFn(x => x * Math.PI / 180);
  F.SIGN = mathFn(Math.sign);
  F.FACT = mathFn(x => { if (x < 0) return E.NUM; let r = 1; for (let i = 2; i <= Math.floor(x); i++) r *= i; return r; });
  F.ATAN2 = (a, ctx) => Math.atan2(num1(a[1], ctx), num1(a[0], ctx));
  F.PI = () => Math.PI;
  F.RAND = () => Math.random();
  F.RANDBETWEEN = (a, ctx) => { const lo = num1(a[0], ctx), hi = num1(a[1], ctx); return lo + Math.floor(Math.random() * (hi - lo + 1)); };
  const roundTo = (mode) => (a, ctx) => {
    const x = num1(a[0], ctx), d = a[1] !== undefined ? Math.trunc(num1(a[1], ctx)) : 0;
    if (isErr(x)) return x; if (isErr(d)) return d;
    const sh = (v, k) => { const [m, e] = v.toExponential().split('e'); return Number(m + 'e' + (+e + k)); };
    const ax = Math.abs(x); let r;
    if (mode === 'round') r = Math.round(sh(ax, d)); else if (mode === 'up') r = Math.ceil(sh(ax, d) - 1e-12); else r = Math.floor(sh(ax, d) + 1e-12);
    r = sh(r, -d);
    return x < 0 ? -r : r;
  };
  F.ROUND = roundTo('round'); F.ROUNDUP = roundTo('up'); F.ROUNDDOWN = roundTo('down'); F.TRUNC = roundTo('down');
  F.MOD = (a, ctx) => { const x = num1(a[0], ctx), y = num1(a[1], ctx); if (isErr(x)) return x; if (isErr(y)) return y; if (y === 0) return E.DIV0; return x - y * Math.floor(x / y); };
  F.POWER = (a, ctx) => binop('^', scalar(a[0], ctx), scalar(a[1], ctx));
  F.CEILING = (a, ctx) => { const x = num1(a[0], ctx), s = a[1] !== undefined ? num1(a[1], ctx) : 1; return s === 0 ? 0 : Math.ceil(x / s) * s; };
  F.FLOOR = (a, ctx) => { const x = num1(a[0], ctx), s = a[1] !== undefined ? num1(a[1], ctx) : 1; return s === 0 ? 0 : Math.floor(x / s) * s; };
  F.GCD = (a, ctx) => { const v = nums(a, ctx); const g2 = (x, y) => y ? g2(y, x % y) : x; return v.reduce(g2); };
  F.LCM = (a, ctx) => { const v = nums(a, ctx); const g2 = (x, y) => y ? g2(y, x % y) : x; return v.reduce((x, y) => x / g2(x, y) * y); };
  F.SUMPRODUCT = (args, ctx) => {
    const arrs = args.map(a => toArr(a, ctx));
    const R = arrs[0].rows.length, C = arrs[0].rows[0].length;
    let s = 0;
    for (let r = 0; r < R; r++) for (let c = 0; c < C; c++) {
      let p = 1;
      for (const A of arrs) { if (!A.rows[r] || A.rows[r][c] === undefined) return E.VALUE; const v = A.rows[r][c]; if (isErr(v)) return v; p *= (typeof v === 'number' ? v : (typeof v === 'boolean' ? +v : 0)); }
      s += p;
    }
    return s;
  };
  F.PMT = (a, ctx) => { const r = num1(a[0], ctx), n = num1(a[1], ctx), pv = num1(a[2], ctx), fv = a[3] ? num1(a[3], ctx) : 0, t = a[4] ? num1(a[4], ctx) : 0; if (r === 0) return -(pv + fv) / n; const k = Math.pow(1 + r, n); return -(r * (pv * k + fv)) / ((1 + r * t) * (k - 1)); };
  F.FV = (a, ctx) => { const r = num1(a[0], ctx), n = num1(a[1], ctx), pmt = num1(a[2], ctx), pv = a[3] ? num1(a[3], ctx) : 0, t = a[4] ? num1(a[4], ctx) : 0; if (r === 0) return -(pv + pmt * n); const k = Math.pow(1 + r, n); return -(pv * k + pmt * (1 + r * t) * (k - 1) / r); };
  F.PV = (a, ctx) => { const r = num1(a[0], ctx), n = num1(a[1], ctx), pmt = num1(a[2], ctx), fv = a[3] ? num1(a[3], ctx) : 0, t = a[4] ? num1(a[4], ctx) : 0; if (r === 0) return -(fv + pmt * n); const k = Math.pow(1 + r, n); return -(fv + pmt * (1 + r * t) * (k - 1) / r) / k; };
  F.NPV = (a, ctx) => { const r = num1(a[0], ctx); const v = nums(a.slice(1), ctx); return v.reduce((s, x, i) => s + x / Math.pow(1 + r, i + 1), 0); };

  /* criteria functions */
  function matcher(crit) {
    if (typeof crit === 'number' || typeof crit === 'boolean') return (v) => typeof v === typeof crit && v === crit || (typeof crit === 'number' && typeof v === 'string' && Number(v) === crit);
    crit = str(crit);
    let m = /^(<=|>=|<>|<|>|=)(.*)$/.exec(crit);
    let op = '=', rhs = crit;
    if (m) { op = m[1]; rhs = m[2]; }
    const rn = rhs.trim() !== '' && !isNaN(Number(rhs)) ? Number(rhs) : null;
    const wild = /[*?]/.test(rhs) ? new RegExp('^' + rhs.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~?\*/g, '.*').replace(/~?\?/g, '.') + '$', 'i') : null;
    return (v) => {
      if (isErr(v)) return false;
      if (op === '=' || op === '<>') {
        let eq;
        if (rn !== null) eq = (typeof v === 'number' && v === rn);
        else if (wild) eq = typeof v === 'string' && wild.test(v);
        else if (rhs === '') eq = v === null || v === '';
        else eq = (typeof v === 'string' && v.toLowerCase() === rhs.toLowerCase()) || (typeof v === 'boolean' && String(v).toLowerCase() === rhs.toLowerCase());
        return op === '=' ? eq : !eq;
      }
      if (rn !== null) { if (typeof v !== 'number') return false; return op === '<' ? v < rn : op === '>' ? v > rn : op === '<=' ? v <= rn : v >= rn; }
      if (typeof v !== 'string') return false;
      const c = v.toLowerCase().localeCompare(rhs.toLowerCase());
      return op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
    };
  }
  function condRows(pairs, ctx) { // pairs: [[range, criteria],...] -> boolean array
    let flags = null;
    for (const [rg, cr] of pairs) {
      const arr = toArr(rg, ctx).rows.flat();
      const m = matcher(scalar(cr, ctx));
      if (!flags) flags = arr.map(v => m(v)); else { if (arr.length !== flags.length) return E.VALUE; arr.forEach((v, i) => { flags[i] = flags[i] && m(v); }); }
    }
    return flags;
  }
  F.COUNTIF = (a, ctx) => { const f = condRows([[a[0], a[1]]], ctx); return isErr(f) ? f : f.filter(Boolean).length; };
  F.COUNTIFS = (a, ctx) => { const p = []; for (let i = 0; i < a.length; i += 2) p.push([a[i], a[i + 1]]); const f = condRows(p, ctx); return isErr(f) ? f : f.filter(Boolean).length; };
  const sumLike = (red, empty) => (a, ctx, hasFirstSum) => { };
  F.SUMIF = (a, ctx) => { const f = condRows([[a[0], a[1]]], ctx); if (isErr(f)) return f; const sr = toArr(a[2] !== undefined ? a[2] : a[0], ctx).rows.flat(); let s = 0; f.forEach((x, i) => { if (x && typeof sr[i] === 'number') s += sr[i]; }); return s; };
  F.SUMIFS = (a, ctx) => { const p = []; for (let i = 1; i < a.length; i += 2) p.push([a[i], a[i + 1]]); const f = condRows(p, ctx); if (isErr(f)) return f; const sr = toArr(a[0], ctx).rows.flat(); let s = 0; f.forEach((x, i) => { if (x && typeof sr[i] === 'number') s += sr[i]; }); return s; };
  F.AVERAGEIF = (a, ctx) => { const f = condRows([[a[0], a[1]]], ctx); if (isErr(f)) return f; const sr = toArr(a[2] !== undefined ? a[2] : a[0], ctx).rows.flat(); let s = 0, n = 0; f.forEach((x, i) => { if (x && typeof sr[i] === 'number') { s += sr[i]; n++; } }); return n ? s / n : E.DIV0; };
  F.AVERAGEIFS = (a, ctx) => { const p = []; for (let i = 1; i < a.length; i += 2) p.push([a[i], a[i + 1]]); const f = condRows(p, ctx); if (isErr(f)) return f; const sr = toArr(a[0], ctx).rows.flat(); let s = 0, n = 0; f.forEach((x, i) => { if (x && typeof sr[i] === 'number') { s += sr[i]; n++; } }); return n ? s / n : E.DIV0; };
  const mxif = (fn) => (a, ctx) => { const p = []; for (let i = 1; i < a.length; i += 2) p.push([a[i], a[i + 1]]); const f = condRows(p, ctx); if (isErr(f)) return f; const sr = toArr(a[0], ctx).rows.flat(); const v = []; f.forEach((x, i) => { if (x && typeof sr[i] === 'number') v.push(sr[i]); }); return v.length ? fn(...v) : 0; };
  F.MAXIFS = mxif(Math.max); F.MINIFS = mxif(Math.min);

  /* logical */
  const truthy = (v) => { if (isErr(v)) return v; if (typeof v === 'boolean') return v; if (typeof v === 'number') return v !== 0; if (v === null || v === undefined) return false; if (typeof v === 'string') { if (/^true$/i.test(v)) return true; if (/^false$/i.test(v)) return false; return E.VALUE; } return false; };
  L.IF = (n, ctx) => {
    const c = truthy(scalar(evalNode(n[0], ctx), ctx));
    if (isErr(c)) return c;
    if (c) return n[1] ? evalNode(n[1], ctx) : true;
    return n[2] ? evalNode(n[2], ctx) : false;
  };
  L.IFS = (n, ctx) => { for (let i = 0; i + 1 < n.length; i += 2) { const c = truthy(scalar(evalNode(n[i], ctx), ctx)); if (isErr(c)) return c; if (c) return evalNode(n[i + 1], ctx); } return E.NA; };
  L.IFERROR = (n, ctx) => { const v = scalar(evalNode(n[0], ctx), ctx); return isErr(v) ? evalNode(n[1], ctx) : v; };
  L.IFNA = (n, ctx) => { const v = scalar(evalNode(n[0], ctx), ctx); return v === E.NA ? evalNode(n[1], ctx) : v; };
  L.CHOOSE = (n, ctx) => { const i = num1(evalNode(n[0], ctx), ctx); if (isErr(i)) return i; const k = Math.floor(i); return n[k] ? evalNode(n[k], ctx) : E.VALUE; };
  L.SWITCH = (n, ctx) => { const v = scalar(evalNode(n[0], ctx), ctx); for (let i = 1; i + 1 < n.length; i += 2) if (compare(v, scalar(evalNode(n[i], ctx), ctx)) === 0) return evalNode(n[i + 1], ctx); return n.length % 2 === 0 ? evalNode(n[n.length - 1], ctx) : E.NA; };
  const boolReduce = (isAnd) => (a, ctx) => { const v = all(a, ctx).filter(x => x !== null && typeof x !== 'string'); if (!v.length) return E.VALUE; let r = isAnd; for (const x of v) { const t = truthy(x); if (isErr(t)) return t; r = isAnd ? (r && t) : (r || t); } return r; };
  F.AND = boolReduce(true); F.OR = boolReduce(false);
  F.NOT = (a, ctx) => { const t = truthy(scalar(a[0], ctx)); return isErr(t) ? t : !t; };
  F.XOR = (a, ctx) => all(a, ctx).filter(x => truthy(x) === true).length % 2 === 1;
  F.TRUE = () => true; F.FALSE = () => false;
  F.ISBLANK = (a, ctx) => { const v = scalar(a[0], ctx); return v === null || v === undefined; };
  F.ISNUMBER = (a, ctx) => typeof scalar(a[0], ctx) === 'number';
  F.ISTEXT = (a, ctx) => typeof scalar(a[0], ctx) === 'string';
  F.ISLOGICAL = (a, ctx) => typeof scalar(a[0], ctx) === 'boolean';
  F.ISERROR = (a, ctx) => isErr(scalar(a[0], ctx));
  F.ISERR = (a, ctx) => { const v = scalar(a[0], ctx); return isErr(v) && v !== E.NA; };
  F.ISNA = (a, ctx) => scalar(a[0], ctx) === E.NA;
  F.ISEVEN = (a, ctx) => num1(a[0], ctx) % 2 === 0;
  F.ISODD = (a, ctx) => Math.abs(num1(a[0], ctx) % 2) === 1;
  F.NA = () => E.NA;
  F.N = (a, ctx) => { const v = scalar(a[0], ctx); return typeof v === 'number' ? v : (typeof v === 'boolean' ? +v : 0); };
  F.T = (a, ctx) => { const v = scalar(a[0], ctx); return typeof v === 'string' ? v : ''; };

  /* text */
  const s1 = (v, ctx) => str(scalar(v, ctx));
  F.CONCAT = F.CONCATENATE = (a, ctx) => { let s = ''; for (const v of all(a, ctx)) { if (isErr(v)) return v; s += str(v); } return s; };
  F.TEXTJOIN = (a, ctx) => { const d = s1(a[0], ctx), skip = truthy(scalar(a[1], ctx)); const v = all(a.slice(2), ctx); for (const x of v) if (isErr(x)) return x; return v.filter(x => !(skip && (x === null || x === ''))).map(str).join(d); };
  F.LEFT = (a, ctx) => s1(a[0], ctx).slice(0, a[1] !== undefined ? num1(a[1], ctx) : 1);
  F.RIGHT = (a, ctx) => { const s = s1(a[0], ctx), n = a[1] !== undefined ? num1(a[1], ctx) : 1; return n <= 0 ? '' : s.slice(-n); };
  F.MID = (a, ctx) => { const s = s1(a[0], ctx), st = num1(a[1], ctx), n = num1(a[2], ctx); if (st < 1 || n < 0) return E.VALUE; return s.substr(st - 1, n); };
  F.LEN = (a, ctx) => s1(a[0], ctx).length;
  F.UPPER = (a, ctx) => s1(a[0], ctx).toUpperCase();
  F.LOWER = (a, ctx) => s1(a[0], ctx).toLowerCase();
  F.PROPER = (a, ctx) => s1(a[0], ctx).toLowerCase().replace(/(^|[^a-z])([a-z])/g, (m, p, c) => p + c.toUpperCase());
  F.TRIM = (a, ctx) => s1(a[0], ctx).trim().replace(/ +/g, ' ');
  F.CLEAN = (a, ctx) => s1(a[0], ctx).replace(/[\x00-\x1f]/g, '');
  F.EXACT = (a, ctx) => s1(a[0], ctx) === s1(a[1], ctx);
  F.REPT = (a, ctx) => s1(a[0], ctx).repeat(Math.max(0, num1(a[1], ctx)));
  F.CHAR = (a, ctx) => String.fromCharCode(num1(a[0], ctx));
  F.CODE = (a, ctx) => s1(a[0], ctx).charCodeAt(0) || E.VALUE;
  F.SUBSTITUTE = (a, ctx) => { const s = s1(a[0], ctx), o = s1(a[1], ctx), n = s1(a[2], ctx); if (!o) return s; if (a[3] !== undefined) { let k = num1(a[3], ctx), idx = -1; while (k-- > 0) { idx = s.indexOf(o, idx + 1); if (idx < 0) return s; } return s.slice(0, idx) + n + s.slice(idx + o.length); } return s.split(o).join(n); };
  F.REPLACE = (a, ctx) => { const s = s1(a[0], ctx), st = num1(a[1], ctx), n = num1(a[2], ctx); return s.slice(0, st - 1) + s1(a[3], ctx) + s.slice(st - 1 + n); };
  F.FIND = (a, ctx) => { const i = s1(a[1], ctx).indexOf(s1(a[0], ctx), a[2] !== undefined ? num1(a[2], ctx) - 1 : 0); return i < 0 ? E.VALUE : i + 1; };
  F.SEARCH = (a, ctx) => { const pat = s1(a[0], ctx).toLowerCase(); const re = new RegExp(pat.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.')); const hay = s1(a[1], ctx).toLowerCase(); const st = a[2] !== undefined ? num1(a[2], ctx) - 1 : 0; const m = re.exec(hay.slice(st)); return m ? m.index + st + 1 : E.VALUE; };
  F.TEXT = (a, ctx) => { const v = scalar(a[0], ctx); if (isErr(v)) return v; return NumFmt.format(typeof v === 'string' && !isErr(numFrom(v)) ? numFrom(v) : v, s1(a[1], ctx)); };
  F.VALUE = (a, ctx) => { const v = scalar(a[0], ctx); const x = numFrom(v); return x; };
  F.NUMBERVALUE = F.VALUE;

  /* dates */
  const nowSerial = () => { const d = new Date(); return NumFmt.dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()) + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400; };
  F.TODAY = () => Math.floor(nowSerial());
  F.NOW = () => nowSerial();
  F.DATE = (a, ctx) => { let y = num1(a[0], ctx), m = num1(a[1], ctx), d = num1(a[2], ctx); if (y < 1900) y += 1900; const dt = new Date(Date.UTC(y, m - 1, d)); return NumFmt.dateToSerial(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate()); };
  const dp = (f) => (a, ctx) => { const v = num1(a[0], ctx); if (isErr(v)) return v; return f(NumFmt.serialToParts(v)); };
  F.YEAR = dp(p => p.y); F.MONTH = dp(p => p.m); F.DAY = dp(p => p.d); F.HOUR = dp(p => p.hh); F.MINUTE = dp(p => p.mm); F.SECOND = dp(p => p.ss);
  F.WEEKDAY = (a, ctx) => { const p = NumFmt.serialToParts(num1(a[0], ctx)); const t = a[1] !== undefined ? num1(a[1], ctx) : 1; return t === 2 ? (p.dow === 0 ? 7 : p.dow) : (t === 3 ? (p.dow + 6) % 7 : p.dow + 1); };
  F.TIME = (a, ctx) => (num1(a[0], ctx) * 3600 + num1(a[1], ctx) * 60 + num1(a[2], ctx)) / 86400;
  F.DATEVALUE = (a, ctx) => { const r = NumFmt.parseInput(s1(a[0], ctx)); return typeof r.v === 'number' ? Math.floor(r.v) : E.VALUE; };
  F.DAYS = (a, ctx) => num1(a[0], ctx) - num1(a[1], ctx);
  F.EDATE = (a, ctx) => { const p = NumFmt.serialToParts(num1(a[0], ctx)); const dt = new Date(Date.UTC(p.y, p.m - 1 + num1(a[1], ctx), p.d)); return NumFmt.dateToSerial(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate()); };
  F.EOMONTH = (a, ctx) => { const p = NumFmt.serialToParts(num1(a[0], ctx)); const dt = new Date(Date.UTC(p.y, p.m + num1(a[1], ctx), 0)); return NumFmt.dateToSerial(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate()); };
  F.DATEDIF = (a, ctx) => { const s = NumFmt.serialToParts(num1(a[0], ctx)), e = NumFmt.serialToParts(num1(a[1], ctx)); const u = s1(a[2], ctx).toUpperCase(); let months = (e.y - s.y) * 12 + e.m - s.m - (e.d < s.d ? 1 : 0); if (u === 'D') return Math.floor(num1(a[1], ctx)) - Math.floor(num1(a[0], ctx)); if (u === 'M') return months; if (u === 'Y') return Math.floor(months / 12); return E.NUM; };

  /* lookup */
  F.ROW = (a, ctx) => a.length && isRef(a[0]) ? a[0].r1 + 1 : ctx.r + 1;
  F.COLUMN = (a, ctx) => a.length && isRef(a[0]) ? a[0].c1 + 1 : ctx.c + 1;
  F.ROWS = (a, ctx) => toArr(a[0], ctx).rows.length;
  F.COLUMNS = (a, ctx) => toArr(a[0], ctx).rows[0].length;
  F.INDEX = (a, ctx) => {
    const A = toArr(a[0], ctx).rows;
    let r = a[1] !== undefined ? num1(a[1], ctx) : 1, c = a[2] !== undefined ? num1(a[2], ctx) : 1;
    if (A.length === 1 && a[2] === undefined) { c = r; r = 1; }
    if (r === 0 && c > 0) return { isArr: true, rows: A.map(x => [x[c - 1]]) };
    if (!A[r - 1] || A[r - 1][c - 1] === undefined) return E.REF;
    return A[r - 1][c - 1];
  };
  F.MATCH = (a, ctx) => {
    const v = scalar(a[0], ctx), arr = toArr(a[1], ctx).rows.flat(), t = a[2] !== undefined ? num1(a[2], ctx) : 1;
    if (t === 0) { const m = typeof v === 'string' && /[*?]/.test(v) ? matcher(v) : (x) => compare(x, v) === 0 && typeof x === typeof v; const i = arr.findIndex(m); return i < 0 ? E.NA : i + 1; }
    let best = -1;
    for (let i = 0; i < arr.length; i++) { if (arr[i] === null) continue; const c = compare(arr[i], v); if (t > 0 ? c <= 0 : c >= 0) best = i; else break; }
    return best < 0 ? E.NA : best + 1;
  };
  F.VLOOKUP = (a, ctx) => {
    const v = scalar(a[0], ctx), T = toArr(a[1], ctx).rows, ci = num1(a[2], ctx), approx = a[3] === undefined || truthy(scalar(a[3], ctx));
    if (ci < 1 || ci > T[0].length) return E.REF;
    let row = -1;
    if (!approx) { row = T.findIndex(r => (typeof v === 'string' && /[*?]/.test(v)) ? matcher(v)(r[0]) : (compare(r[0], v) === 0 && typeof r[0] === typeof v)); }
    else for (let i = 0; i < T.length; i++) { if (T[i][0] !== null && compare(T[i][0], v) <= 0 && typeof T[i][0] === typeof v) row = i; else if (typeof T[i][0] === typeof v) break; }
    return row < 0 ? E.NA : T[row][ci - 1];
  };
  F.HLOOKUP = (a, ctx) => {
    const v = scalar(a[0], ctx), T = toArr(a[1], ctx).rows, ri = num1(a[2], ctx), approx = a[3] === undefined || truthy(scalar(a[3], ctx));
    if (ri < 1 || ri > T.length) return E.REF;
    let col = -1;
    if (!approx) col = T[0].findIndex(x => compare(x, v) === 0 && typeof x === typeof v);
    else for (let i = 0; i < T[0].length; i++) { if (T[0][i] !== null && compare(T[0][i], v) <= 0) col = i; else break; }
    return col < 0 ? E.NA : T[ri - 1][col];
  };
  F.XLOOKUP = (a, ctx) => {
    const v = scalar(a[0], ctx), look = toArr(a[1], ctx).rows.flat(), ret = toArr(a[2], ctx).rows;
    const i = look.findIndex(x => compare(x, v) === 0 && typeof x === typeof v);
    if (i < 0) return a[3] !== undefined ? scalar(a[3], ctx) : E.NA;
    return ret.length === 1 ? ret[0][i] : (ret[i] ? ret[i][0] : E.NA);
  };
  F.LOOKUP = (a, ctx) => { const v = scalar(a[0], ctx), look = toArr(a[1], ctx).rows.flat(), ret = a[2] !== undefined ? toArr(a[2], ctx).rows.flat() : look; let best = -1; for (let i = 0; i < look.length; i++) { if (look[i] !== null && compare(look[i], v) <= 0) best = i; else break; } return best < 0 ? E.NA : ret[best]; };
  F.INDIRECT = (a, ctx) => { const s = s1(a[0], ctx); const m = /^(?:'?([^'!]+)'?!)?(.+)$/.exec(s); const rg = m && parseRange(m[2]); if (!rg) return E.REF; const sh = m[1] ? ctx.wb.sheetByName(m[1]) : ctx.sh; if (!sh) return E.REF; return { isRef: true, sh, ...rg }; };
  F.TRANSPOSE = (a, ctx) => { const A = toArr(a[0], ctx).rows; return { isArr: true, rows: A[0].map((_, c) => A.map(r => r[c])) }; };

  /* ---------- additional functions ---------- */
  const pairs = (a, ctx) => { const x = toArr(a[0], ctx).rows.flat(), y = toArr(a[1], ctx).rows.flat(); const px = [], py = []; for (let i = 0; i < Math.min(x.length, y.length); i++) if (typeof x[i] === 'number' && typeof y[i] === 'number') { px.push(x[i]); py.push(y[i]); } return [px, py]; };
  const mean = (v) => v.reduce((a, b) => a + b, 0) / v.length;
  const regress = (px, py) => { const mx = mean(px), my = mean(py); let sxy = 0, sxx = 0, syy = 0; px.forEach((x, i) => { sxy += (x - mx) * (py[i] - my); sxx += (x - mx) * (x - mx); syy += (py[i] - my) * (py[i] - my); }); return { mx, my, sxy, sxx, syy }; };
  F.CORREL = F.PEARSON = (a, ctx) => { const [x, y] = pairs(a, ctx); if (x.length < 2) return E.DIV0; const r = regress(x, y); return r.sxy / Math.sqrt(r.sxx * r.syy); };
  F.SLOPE = (a, ctx) => { const [y, x] = pairs(a, ctx); if (x.length < 2) return E.DIV0; const r = regress(x, y); return r.sxy / r.sxx; };
  F.INTERCEPT = (a, ctx) => { const [y, x] = pairs(a, ctx); if (x.length < 2) return E.DIV0; const r = regress(x, y); return r.my - r.sxy / r.sxx * r.mx; };
  F.RSQ = (a, ctx) => { const [x, y] = pairs(a, ctx); const r = regress(x, y); return r.sxy * r.sxy / (r.sxx * r.syy); };
  F.FORECAST = F['FORECAST.LINEAR'] = (a, ctx) => { const x0 = num1(a[0], ctx); const [y, x] = pairs(a.slice(1), ctx); const r = regress(x, y); return r.my + r.sxy / r.sxx * (x0 - r.mx); };
  F.COVAR = F['COVARIANCE.P'] = (a, ctx) => { const [x, y] = pairs(a, ctx); const r = regress(x, y); return r.sxy / x.length; };
  F.AVERAGEA = (a, ctx) => { const v = all(a, ctx).filter(x => x !== null && x !== undefined); if (!v.length) return E.DIV0; return v.reduce((s, x) => s + (typeof x === 'number' ? x : (typeof x === 'boolean' ? +x : 0)), 0) / v.length; };
  F.GEOMEAN = reduceNums(v => Math.exp(v.reduce((s, x) => s + Math.log(x), 0) / v.length), E.NUM);
  F.HARMEAN = reduceNums(v => v.length / v.reduce((s, x) => s + 1 / x, 0), E.NUM);
  F.QUARTILE = F['QUARTILE.INC'] = (a, ctx) => F.PERCENTILE([a[0], num1(a[1], ctx) / 4], ctx);
  F['PERCENTILE.INC'] = F.PERCENTILE;
  F.PERCENTRANK = (a, ctx) => { const v = nums([a[0]], ctx).sort((p, q) => p - q); const x = num1(a[1], ctx); const i = v.findIndex(y => y >= x); if (i < 0 || v.length < 2) return E.NA; return +(i / (v.length - 1)).toFixed(3); };
  F.EVEN = mathFn(x => { const r = Math.ceil(Math.abs(x) / 2) * 2; return x < 0 ? -r : r; });
  F.ODD = mathFn(x => { let r = Math.ceil(Math.abs(x)); if (r % 2 === 0) r++; return x < 0 ? -r : r; });
  F.MROUND = (a, ctx) => { const x = num1(a[0], ctx), m = num1(a[1], ctx); return m === 0 ? 0 : Math.round(x / m) * m; };
  F.QUOTIENT = (a, ctx) => { const y = num1(a[1], ctx); return y === 0 ? E.DIV0 : Math.trunc(num1(a[0], ctx) / y); };
  F.COMBIN = (a, ctx) => { const n = num1(a[0], ctx), k = num1(a[1], ctx); if (k > n || k < 0) return E.NUM; let r = 1; for (let i = 1; i <= k; i++) r = r * (n - k + i) / i; return Math.round(r); };
  F.PERMUT = (a, ctx) => { const n = num1(a[0], ctx), k = num1(a[1], ctx); if (k > n || k < 0) return E.NUM; let r = 1; for (let i = 0; i < k; i++) r *= (n - i); return r; };
  F.SINH = mathFn(Math.sinh); F.COSH = mathFn(Math.cosh); F.TANH = mathFn(Math.tanh);
  F.ROMAN = mathFn(x => { const m = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']]; let s = ''; x = Math.floor(x); for (const [v, r] of m) while (x >= v) { s += r; x -= v; } return s; });
  F.SUBTOTAL = (a, ctx) => {
    const code = num1(a[0], ctx);
    const ign = code > 100;
    const base = ign ? code - 100 : code;
    const vals = [];
    for (const arg of a.slice(1)) {
      if (!(isRef(arg))) { const x = numFrom(arg); if (typeof x === 'number') vals.push(x); continue; }
      const b = ctx.wb.bounds(arg.sh);
      for (let r = arg.r1; r <= Math.min(arg.r2, Math.max(arg.r1, b.maxR)); r++) {
        if (arg.sh.rowH.get(r) === 0) continue;
        for (let c = arg.c1; c <= Math.min(arg.c2, Math.max(arg.c1, b.maxC)); c++) { const v = cellValue(ctx.wb, arg.sh, r, c); if (typeof v === 'number') vals.push(v); }
      }
    }
    switch (base) {
      case 1: return vals.length ? mean(vals) : E.DIV0; case 2: case 3: return vals.length; case 4: return vals.length ? Math.max(...vals) : 0; case 5: return vals.length ? Math.min(...vals) : 0;
      case 6: return vals.reduce((p, x) => p * x, 1); case 7: return vals.length > 1 ? Math.sqrt(varp(vals) / (vals.length - 1)) : E.DIV0; case 8: return vals.length ? Math.sqrt(varp(vals) / vals.length) : E.DIV0;
      case 9: return vals.reduce((s, x) => s + x, 0); case 10: return vals.length > 1 ? varp(vals) / (vals.length - 1) : E.DIV0; case 11: return vals.length ? varp(vals) / vals.length : E.DIV0;
    }
    return E.VALUE;
  };
  F.OFFSET = (a, ctx) => { const b = a[0]; if (!isRef(b)) return E.VALUE; const dr = num1(a[1], ctx), dc = num1(a[2], ctx); const h = a[3] !== undefined ? num1(a[3], ctx) : b.r2 - b.r1 + 1, w = a[4] !== undefined ? num1(a[4], ctx) : b.c2 - b.c1 + 1; const r1 = b.r1 + dr, c1 = b.c1 + dc; if (r1 < 0 || c1 < 0) return E.REF; return { isRef: true, sh: b.sh, r1, c1, r2: r1 + h - 1, c2: c1 + w - 1 }; };
  F.HYPERLINK = (a, ctx) => a[1] !== undefined ? s1(a[1], ctx) : s1(a[0], ctx);
  F.DOLLAR = (a, ctx) => NumFmt.format(num1(a[0], ctx), '$#,##0' + (a[1] === undefined || num1(a[1], ctx) > 0 ? '.' + '0'.repeat(a[1] === undefined ? 2 : num1(a[1], ctx)) : '') + ';-$#,##0' + (a[1] === undefined || num1(a[1], ctx) > 0 ? '.' + '0'.repeat(a[1] === undefined ? 2 : num1(a[1], ctx)) : ''));
  F.FIXED = (a, ctx) => NumFmt.format(num1(a[0], ctx), (a[2] !== undefined && truthy(scalar(a[2], ctx)) ? '0' : '#,##0') + (a[1] === undefined || num1(a[1], ctx) > 0 ? '.' + '0'.repeat(a[1] === undefined ? 2 : num1(a[1], ctx)) : ''));
  F.TEXTBEFORE = (a, ctx) => { const s = s1(a[0], ctx), d = s1(a[1], ctx); const i = s.indexOf(d); return i < 0 ? E.NA : s.slice(0, i); };
  F.TEXTAFTER = (a, ctx) => { const s = s1(a[0], ctx), d = s1(a[1], ctx); const i = s.indexOf(d); return i < 0 ? E.NA : s.slice(i + d.length); };
  F.WEEKNUM = (a, ctx) => { const p = NumFmt.serialToParts(num1(a[0], ctx)); const d = new Date(Date.UTC(p.y, p.m - 1, p.d)); const start = new Date(Date.UTC(p.y, 0, 1)); return Math.floor(((d - start) / 864e5 + start.getUTCDay()) / 7) + 1; };
  F.YEARFRAC = (a, ctx) => Math.abs(num1(a[1], ctx) - num1(a[0], ctx)) / 365.25;
  const isWork = (s) => { const d = NumFmt.serialToParts(s).dow; return d !== 0 && d !== 6; };
  F.NETWORKDAYS = (a, ctx) => { let s = Math.floor(num1(a[0], ctx)), e = Math.floor(num1(a[1], ctx)); const sg = e < s ? -1 : 1; if (sg < 0) [s, e] = [e, s]; let n = 0; for (let d = s; d <= e; d++) if (isWork(d)) n++; return n * sg; };
  F.WORKDAY = (a, ctx) => { let d = Math.floor(num1(a[0], ctx)), n = num1(a[1], ctx); const st = n < 0 ? -1 : 1; while (n !== 0) { d += st; if (isWork(d)) n -= st; } return d; };
  F.NPER = (a, ctx) => { const r = num1(a[0], ctx), p = num1(a[1], ctx), pv = num1(a[2], ctx), fv = a[3] ? num1(a[3], ctx) : 0; if (r === 0) return -(pv + fv) / p; return Math.log((p - fv * r) / (p + pv * r)) / Math.log(1 + r); };
  F.RATE = (a, ctx) => { const n = num1(a[0], ctx), p = num1(a[1], ctx), pv = num1(a[2], ctx), fv = a[3] ? num1(a[3], ctx) : 0; let r = 0.1; for (let i = 0; i < 100; i++) { const k = Math.pow(1 + r, n); const f = pv * k + p * (k - 1) / r + fv; const d = (pv * n * Math.pow(1 + r, n - 1) + p * (n * Math.pow(1 + r, n - 1) * r - (k - 1)) / (r * r)); const nr = r - f / d; if (!isFinite(nr)) return E.NUM; if (Math.abs(nr - r) < 1e-10) return nr; r = nr; } return E.NUM; };
  F.IRR = (a, ctx) => { const v = nums([a[0]], ctx); let r = 0.1; for (let i = 0; i < 200; i++) { let f = 0, d = 0; v.forEach((x, t) => { f += x / Math.pow(1 + r, t); d -= t * x / Math.pow(1 + r, t + 1); }); const nr = r - f / d; if (!isFinite(nr)) return E.NUM; if (Math.abs(nr - r) < 1e-10) return nr; r = nr; } return E.NUM; };
  F.SLN = (a, ctx) => (num1(a[0], ctx) - num1(a[1], ctx)) / num1(a[2], ctx);
  F.IPMT = (a, ctx) => { const r = num1(a[0], ctx), per = num1(a[1], ctx), n = num1(a[2], ctx), pv = num1(a[3], ctx); const pmt = F.PMT([r, n, pv], ctx); const bal = pv * Math.pow(1 + r, per - 1) + pmt * (Math.pow(1 + r, per - 1) - 1) / r; return -(bal * r); };
  F.PPMT = (a, ctx) => F.PMT([a[0], a[2], a[3]], ctx) - F.IPMT(a, ctx);

  /* ---------- dynamic array functions ---------- */
  const key2 = (row) => row.map(x => (x === null ? '' : typeof x) + ':' + String(x).toLowerCase()).join('\u0001');
  const asRows = (A, byCol) => byCol ? A[0].map((_, c) => A.map(r => r[c])) : A.map(r => r.slice());
  const fromRows = (rows, byCol) => ({ isArr: true, rows: byCol ? rows[0].map((_, c) => rows.map(r => r[c])) : rows });
  const flag = (v, ctx) => v !== undefined && truthy(scalar(v, ctx)) === true;
  F.UNIQUE = (a, ctx) => {
    const byCol = flag(a[1], ctx), once = flag(a[2], ctx);
    const rows = asRows(toArr(a[0], ctx).rows, byCol);
    const cnt = new Map(), order = [];
    for (const r of rows) { const k = key2(r); if (!cnt.has(k)) { cnt.set(k, 0); order.push([k, r]); } cnt.set(k, cnt.get(k) + 1); }
    const out = order.filter(([k]) => !once || cnt.get(k) === 1).map(([, r]) => r);
    return out.length ? fromRows(out, byCol) : E.NA;
  };
  const cmpv = (x, y) => compare(x, y);
  F.SORT = (a, ctx) => {
    const byCol = flag(a[3], ctx);
    const idx = a[1] !== undefined ? num1(a[1], ctx) - 1 : 0;
    const desc = a[2] !== undefined && num1(a[2], ctx) === -1;
    const rows = asRows(toArr(a[0], ctx).rows, byCol);
    rows.sort((p, q) => (desc ? -1 : 1) * cmpv(p[idx], q[idx]));
    return fromRows(rows, byCol);
  };
  F.SORTBY = (a, ctx) => {
    const rows = toArr(a[0], ctx).rows.map(r => r.slice());
    const keys = [];
    for (let i = 1; i < a.length; i += 2) keys.push({ v: toArr(a[i], ctx).rows.flat(), desc: a[i + 1] !== undefined && num1(a[i + 1], ctx) === -1 });
    const idx = rows.map((_, i) => i);
    idx.sort((p, q) => { for (const k of keys) { const c = cmpv(k.v[p], k.v[q]); if (c) return k.desc ? -c : c; } return p - q; });
    return { isArr: true, rows: idx.map(i => rows[i]) };
  };
  F.FILTER = (a, ctx) => {
    const A = toArr(a[0], ctx).rows, inc = toArr(a[1], ctx).rows;
    const byRow = inc.length === A.length && inc[0].length === 1 || (inc.length === A.length && A[0].length !== inc[0].length);
    let out;
    if (inc[0].length === 1 || inc.length === A.length && inc[0].length === 1) out = A.filter((r, i) => truthy(inc[i][0]) === true);
    else out = A.map(r => r.filter((x, j) => truthy(inc[0][j]) === true));
    if (!out.length || !out[0].length) return a[2] !== undefined ? scalar(a[2], ctx) : E.NA;
    return { isArr: true, rows: out };
  };
  F.SEQUENCE = (a, ctx) => {
    const R = num1(a[0], ctx), C = a[1] !== undefined ? num1(a[1], ctx) : 1, s = a[2] !== undefined ? num1(a[2], ctx) : 1, st = a[3] !== undefined ? num1(a[3], ctx) : 1;
    if (!(R >= 1) || !(C >= 1) || R * C > 1e6) return E.VALUE;
    const rows = []; let v = s;
    for (let r = 0; r < R; r++) { const row = []; for (let c = 0; c < C; c++) { row.push(v); v += st; } rows.push(row); }
    return { isArr: true, rows };
  };
  F.RANDARRAY = (a, ctx) => {
    const R = a[0] !== undefined ? num1(a[0], ctx) : 1, C = a[1] !== undefined ? num1(a[1], ctx) : 1, lo = a[2] !== undefined ? num1(a[2], ctx) : 0, hi = a[3] !== undefined ? num1(a[3], ctx) : 1, whole = flag(a[4], ctx);
    const rows = [];
    for (let r = 0; r < R; r++) { const row = []; for (let c = 0; c < C; c++) row.push(whole ? lo + Math.floor(Math.random() * (hi - lo + 1)) : lo + Math.random() * (hi - lo)); rows.push(row); }
    return { isArr: true, rows };
  };
  F.VSTACK = (a, ctx) => { const parts = a.map(x => toArr(x, ctx).rows); const W = Math.max(...parts.map(p => p[0].length)); return { isArr: true, rows: parts.flatMap(p => p.map(r => Array.from({ length: W }, (_, i) => r[i] === undefined ? E.NA : r[i]))) }; };
  F.HSTACK = (a, ctx) => { const parts = a.map(x => toArr(x, ctx).rows); const H = Math.max(...parts.map(p => p.length)); return { isArr: true, rows: Array.from({ length: H }, (_, r) => parts.flatMap(p => p[r] ? p[r] : Array(p[0].length).fill(E.NA))) }; };
  F.TAKE = (a, ctx) => { const A = toArr(a[0], ctx).rows; const n = num1(a[1], ctx); const m = a[2] !== undefined ? num1(a[2], ctx) : null; let rows = n >= 0 ? A.slice(0, n) : A.slice(n); if (m !== null) rows = rows.map(r => m >= 0 ? r.slice(0, m) : r.slice(m)); return rows.length && rows[0].length ? { isArr: true, rows } : E.VALUE; };
  F.DROP = (a, ctx) => { const A = toArr(a[0], ctx).rows; const n = num1(a[1], ctx); const m = a[2] !== undefined ? num1(a[2], ctx) : 0; let rows = n >= 0 ? A.slice(n) : A.slice(0, n); if (m) rows = rows.map(r => m >= 0 ? r.slice(m) : r.slice(0, m)); return rows.length && rows[0].length ? { isArr: true, rows } : E.VALUE; };
  F.TOCOL = (a, ctx) => ({ isArr: true, rows: toArr(a[0], ctx).rows.flat().map(x => [x]) });
  F.TOROW = (a, ctx) => ({ isArr: true, rows: [toArr(a[0], ctx).rows.flat()] });
  F.MMULT = (a, ctx) => { const A = toArr(a[0], ctx).rows, B = toArr(a[1], ctx).rows; if (A[0].length !== B.length) return E.VALUE; return { isArr: true, rows: A.map(r => B[0].map((_, j) => r.reduce((s, x, k) => s + (numFrom(x) || 0) * (numFrom(B[k][j]) || 0), 0))) }; };
  F.TEXTSPLIT = (a, ctx) => { const s = s1(a[0], ctx), cd = s1(a[1], ctx); const rd = a[2] !== undefined ? s1(a[2], ctx) : null; const rows = (rd ? s.split(rd) : [s]).map(x => x.split(cd)); const W = Math.max(...rows.map(r => r.length)); return { isArr: true, rows: rows.map(r => Array.from({ length: W }, (_, i) => r[i] === undefined ? E.NA : r[i])) }; };
  F.XMATCH = (a, ctx) => { const v = scalar(a[0], ctx), arr = toArr(a[1], ctx).rows.flat(); const i = arr.findIndex(x => compare(x, v) === 0 && typeof x === typeof v); return i < 0 ? E.NA : i + 1; };
  F.SUMXMY2 = (a, ctx) => { const x = toArr(a[0], ctx).rows.flat(), y = toArr(a[1], ctx).rows.flat(); return x.reduce((s, v, i) => s + Math.pow((numFrom(v) || 0) - (numFrom(y[i]) || 0), 2), 0); };
  F.TRIMMEAN = (a, ctx) => { const v = nums([a[0]], ctx).sort((p, q) => p - q); const k = Math.floor(v.length * num1(a[1], ctx) / 2); const t = v.slice(k, v.length - k); return t.length ? mean(t) : E.NUM; };
  F.MAXA = (a, ctx) => Math.max(...all(a, ctx).filter(x => x !== null).map(x => typeof x === 'number' ? x : (typeof x === 'boolean' ? +x : 0)));
  F.MINA = (a, ctx) => Math.min(...all(a, ctx).filter(x => x !== null).map(x => typeof x === 'number' ? x : (typeof x === 'boolean' ? +x : 0)));
  F.ISFORMULA = (a, ctx) => { const b = a[0]; if (!isRef(b)) return false; const cell = b.sh.get(b.r1, b.c1); return !!(cell && cell.f != null); };
  F.ISREF = (a) => isRef(a[0]);
  F.TYPE = (a, ctx) => { const v = scalar(a[0], ctx); return typeof v === 'number' ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 4 : isErr(v) ? 16 : 1; };
  F.ADDRESS = (a, ctx) => { const r = num1(a[0], ctx), c = num1(a[1], ctx), abs = a[2] !== undefined ? num1(a[2], ctx) : 1; const col = colName(c - 1); const s = abs === 1 ? `$${col}$${r}` : abs === 2 ? `${col}$${r}` : abs === 3 ? `$${col}${r}` : `${col}${r}`; return a[4] !== undefined ? s1(a[4], ctx) + '!' + s : s; };
  F.AREAS = () => 1;
  F.UNICODE = (a, ctx) => s1(a[0], ctx).codePointAt(0) || E.VALUE;
  F.UNICHAR = (a, ctx) => String.fromCodePoint(num1(a[0], ctx));
  F.DAYS360 = (a, ctx) => { const s = NumFmt.serialToParts(num1(a[0], ctx)), e = NumFmt.serialToParts(num1(a[1], ctx)); let d1 = Math.min(s.d, 30), d2 = e.d === 31 && d1 === 30 ? 30 : e.d; return (e.y - s.y) * 360 + (e.m - s.m) * 30 + (d2 - d1); };
  F.TIMEVALUE = (a, ctx) => { const r = NumFmt.parseInput(s1(a[0], ctx)); return typeof r.v === 'number' ? r.v - Math.floor(r.v) : E.VALUE; };
  F.DEC2BIN = (a, ctx) => Math.trunc(num1(a[0], ctx)).toString(2).padStart(a[1] !== undefined ? num1(a[1], ctx) : 0, '0');
  F.DEC2HEX = (a, ctx) => Math.trunc(num1(a[0], ctx)).toString(16).toUpperCase().padStart(a[1] !== undefined ? num1(a[1], ctx) : 0, '0');
  F.DEC2OCT = (a, ctx) => Math.trunc(num1(a[0], ctx)).toString(8);
  F.BIN2DEC = (a, ctx) => parseInt(s1(a[0], ctx), 2);
  F.HEX2DEC = (a, ctx) => parseInt(s1(a[0], ctx), 16);
  F.OCT2DEC = (a, ctx) => parseInt(s1(a[0], ctx), 8);
  F.BASE = (a, ctx) => Math.trunc(num1(a[0], ctx)).toString(num1(a[1], ctx)).toUpperCase().padStart(a[2] !== undefined ? num1(a[2], ctx) : 0, '0');
  F.DECIMAL = (a, ctx) => parseInt(s1(a[0], ctx), num1(a[1], ctx));
  F.FACTDOUBLE = mathFn(x => { let r = 1; for (let i = Math.floor(x); i > 1; i -= 2) r *= i; return r; });
  F.SQRTPI = mathFn(x => Math.sqrt(x * Math.PI));
  F.ASINH = mathFn(Math.asinh); F.ACOSH = mathFn(Math.acosh); F.ATANH = mathFn(Math.atanh);
  F.CEILING_MATH = F['CEILING.MATH'] = F.CEILING; F['FLOOR.MATH'] = F.FLOOR;
  F.ISO_CEILING = F.CEILING;
  F.NORM_DIST = F['NORM.DIST'] = (a, ctx) => { const x = num1(a[0], ctx), m = num1(a[1], ctx), s = num1(a[2], ctx), cum = flag(a[3], ctx); const z = (x - m) / s; if (!cum) return Math.exp(-z * z / 2) / (s * Math.sqrt(2 * Math.PI)); const t = 1 / (1 + 0.2316419 * Math.abs(z)); const d = 0.3989423 * Math.exp(-z * z / 2); let p = d * t * (0.3193815 + t * (-0.3565638 + t * (1.781478 + t * (-1.821256 + t * 1.330274)))); return z > 0 ? 1 - p : p; };
  F.CONFIDENCE = F['CONFIDENCE.NORM'] = (a, ctx) => 1.959964 * num1(a[1], ctx) / Math.sqrt(num1(a[2], ctx));
  F.STANDARDIZE = (a, ctx) => (num1(a[0], ctx) - num1(a[1], ctx)) / num1(a[2], ctx);
  F.DEVSQ = reduceNums(v => varp(v), 0);
  F.AVEDEV = reduceNums(v => { const m = mean(v); return mean(v.map(x => Math.abs(x - m))); }, E.NUM);
  F.SKEW = reduceNums(v => { const n = v.length, m = mean(v), s = Math.sqrt(varp(v) / (n - 1)); return n > 2 ? n / ((n - 1) * (n - 2)) * v.reduce((t, x) => t + Math.pow((x - m) / s, 3), 0) : E.DIV0; }, E.DIV0);
  F.COUNTUNIQUE = (a, ctx) => new Set(all(a, ctx).filter(x => x !== null).map(x => String(x).toLowerCase())).size;
  F.COLUMNS = F.COLUMNS;
  F.SHEET = (a, ctx) => ctx.wb.sheets.indexOf(ctx.sh) + 1;
  F.SHEETS = (a, ctx) => ctx.wb.sheets.length;
  F.INFO = () => 'Cells';
  F.CELL = (a, ctx) => { const t = s1(a[0], ctx).toLowerCase(); const b = a[1]; if (!isRef(b)) return E.VALUE; if (t === 'address') return '$' + colName(b.c1) + '$' + (b.r1 + 1); if (t === 'row') return b.r1 + 1; if (t === 'col') return b.c1 + 1; if (t === 'contents') return cellValue(ctx.wb, b.sh, b.r1, b.c1); if (t === 'type') { const v = cellValue(ctx.wb, b.sh, b.r1, b.c1); return v === null ? 'b' : typeof v === 'string' ? 'l' : 'v'; } return E.VALUE; };

  g.Formula = { spills, Workbook, Sheet, XErr, E, ERR_BY_CODE, MAXR, MAXC, colName, colIndex, cellName, parseCell, parseRange, parse, evalNode, cellValue, rewriteRefs, shiftFormula, splitRef, quoteSheet, fmtPart, isErr, str, functions: () => Object.keys(F).concat(Object.keys(L)).sort(), toArr };
})(window);
