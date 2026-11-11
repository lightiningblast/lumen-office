/* Data validation for Lumen Cells.
   rule = { id, r1,c1,r2,c2, type: 'list'|'whole'|'decimal'|'date'|'textLength'|'custom',
            op: 'between'|'notBetween'|'equal'|'notEqual'|'greaterThan'|'lessThan'|'greaterThanOrEqual'|'lessThanOrEqual',
            f1, f2   (literals or '=formula'; list: comma separated text or '=range'),
            blank (allow blank), drop (in-cell dropdown), style: 'stop'|'warning'|'information', title, msg } */
(function (g) {
  'use strict';
  const { E, isErr } = Formula;

  function ruleAt(sh, r, c) {
    for (const d of (sh.dv || [])) if (r >= d.r1 && r <= d.r2 && c >= d.c1 && c <= d.c2) return d;
    return null;
  }
  function evalExpr(wb, sh, text, r, c) {
    try {
      const ast = Formula.parse(text);
      let v = Formula.evalNode(ast, { wb, sh, r, c });
      if (v && v.isRef) v = Formula.cellValue(wb, v.sh, v.r1, v.c1);
      if (v && v.isArr) v = v.rows[0][0];
      return v;
    } catch (e) { return null; }
  }
  function operand(wb, sh, raw, r, c, type) {
    if (raw === undefined || raw === null || raw === '') return null;
    const s = String(raw);
    if (s[0] === '=') return evalExpr(wb, sh, s.slice(1), r, c);
    if (!isNaN(Number(s))) return Number(s);
    if (type === 'date') { const p = NumFmt.parseInput(s); if (typeof p.v === 'number') return p.v; }
    return s;
  }
  // Items for a list rule (array of strings)
  function listItems(wb, sh, d) {
    const src = String(d.f1 || '');
    if (src[0] === '=') {
      const body = src.slice(1);
      const m = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(body);
      let target = sh, rng = body;
      if (m) { target = wb.sheetByName((m[1] || m[2]).replace(/''/g, "'")); rng = m[3]; }
      const rg = target && Formula.parseRange(rng);
      if (!rg) {
        const v = evalExpr(wb, sh, body, d.r1, d.c1);
        return v === null || v === undefined ? [] : [String(v)];
      }
      const out = [];
      for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
        const cell = target.get(r, c);
        if (!cell) continue;
        const v = cell.f != null ? Formula.cellValue(wb, target, r, c) : cell.v;
        if (v !== null && v !== undefined && v !== '') out.push(isErr(v) ? v.code : (typeof v === 'number' ? NumFmt.format(v, wb.styles[cell.s || 0].nf) : String(v)));
      }
      return out;
    }
    return src.split(',').map(x => x.trim()).filter(x => x !== '');
  }
  // returns null when valid, or the rule when invalid. 'v' is the new typed value (parsed), 'text' the typed text.
  function check(wb, sh, r, c, v, text) {
    const d = ruleAt(sh, r, c);
    if (!d) return null;
    if (v === undefined || v === null || v === '') return d.blank === false ? d : null;
    if (d.type === 'list') {
      const items = listItems(wb, sh, d).map(x => x.toLowerCase());
      return items.includes(String(text !== undefined ? text : v).toLowerCase()) ? null : d;
    }
    if (d.type === 'custom') {
      const res = evalExpr(wb, sh, String(d.f1 || '').replace(/^=/, ''), r, c);
      return (res === true || (typeof res === 'number' && res !== 0)) ? null : d;
    }
    let x;
    if (d.type === 'textLength') x = String(typeof v === 'number' ? text : v).length;
    else {
      if (typeof v !== 'number') return d;
      if (d.type === 'whole' && !Number.isInteger(v)) return d;
      x = v;
    }
    const a = operand(wb, sh, d.f1, r, c, d.type), b = operand(wb, sh, d.f2, r, c, d.type);
    if (typeof a !== 'number') return null;
    let ok;
    switch (d.op || 'between') {
      case 'between': ok = typeof b === 'number' && x >= Math.min(a, b) && x <= Math.max(a, b); break;
      case 'notBetween': ok = typeof b === 'number' && (x < Math.min(a, b) || x > Math.max(a, b)); break;
      case 'equal': ok = x === a; break; case 'notEqual': ok = x !== a; break;
      case 'greaterThan': ok = x > a; break; case 'lessThan': ok = x < a; break;
      case 'greaterThanOrEqual': ok = x >= a; break; case 'lessThanOrEqual': ok = x <= a; break;
      default: ok = true;
    }
    return ok ? null : d;
  }
  function message(d) {
    if (d.msg) return d.msg;
    switch (d.type) {
      case 'list': return 'The value must be one of the items in the list.';
      case 'whole': return 'The value must be a whole number' + opText(d) + '.';
      case 'decimal': return 'The value must be a number' + opText(d) + '.';
      case 'date': return 'The value must be a date' + opText(d) + '.';
      case 'textLength': return 'The text length must be valid' + opText(d) + '.';
      default: return 'The value does not meet the validation rule for this cell.';
    }
  }
  function opText(d) {
    const t = { between: ' between ' + d.f1 + ' and ' + d.f2, notBetween: ' not between ' + d.f1 + ' and ' + d.f2, equal: ' equal to ' + d.f1, notEqual: ' not equal to ' + d.f1, greaterThan: ' greater than ' + d.f1, lessThan: ' less than ' + d.f1, greaterThanOrEqual: ' of at least ' + d.f1, lessThanOrEqual: ' of at most ' + d.f1 }[d.op || 'between'];
    return t || '';
  }
  g.Validate = { ruleAt, check, listItems, message };
})(window);
