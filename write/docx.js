/* DOCX (WordprocessingML) reader and writer for Write.
   write(editorRoot, page)  -> Blob
   read(arrayBuffer)        -> { root: HTMLElement, page }          */
(function (global) {
  'use strict';
  const { xesc, cssToHex, dataUrlToBytes, bytesToDataUrl, MIME_BY_EXT, EXT_BY_MIME } = Lumen;

  const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

  const HL = {
    yellow: 'FFFF00', green: '00FF00', cyan: '00FFFF', magenta: 'FF00FF', blue: '0000FF', red: 'FF0000',
    darkBlue: '000080', darkCyan: '008080', darkGreen: '008000', darkMagenta: '800080', darkRed: '800000',
    darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0', black: '000000', white: 'FFFFFF',
  };
  const HL_REV = {};
  for (const k in HL) HL_REV[HL[k]] = k;

  function lenPx(v) {
    if (!v) return null;
    const m = /^(-?[\d.]+)(px|pt|in|cm|mm|pc|em|rem)?$/.exec(String(v).trim());
    if (!m) return null;
    const n = parseFloat(m[1]);
    switch (m[2] || 'px') {
      case 'pt': return n * 4 / 3;
      case 'in': return n * 96;
      case 'cm': return n * 96 / 2.54;
      case 'mm': return n * 96 / 25.4;
      case 'pc': return n * 16;
      case 'em': case 'rem': return n * 16;
      default: return n;
    }
  }
  const NAMED_PX = { 'xx-small': 9, 'x-small': 10, small: 13, medium: 16, large: 18, 'x-large': 24, 'xx-large': 32, 'xxx-large': 48, '-webkit-xxx-large': 48 };
  const FONT_ATTR_PT = [0, 7.5, 10, 12, 13.5, 18, 24, 36];

  /* ====================================================================
     WRITER
     ==================================================================== */
  const BLOCK_TAGS = new Set(['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'UL', 'OL', 'LI', 'TABLE', 'BLOCKQUOTE', 'PRE', 'HR', 'CENTER', 'SECTION', 'ARTICLE', 'HEADER', 'FOOTER', 'FIGURE', 'FIGCAPTION', 'ADDRESS', 'DL', 'DT', 'DD', 'FORM', 'MAIN', 'NAV', 'ASIDE']);
  const isBlock = (n) => n.nodeType === 1 && BLOCK_TAGS.has(n.tagName);

  function runCtx(el, c) {
    const n = Object.assign({}, c);
    const t = el.tagName;
    if (t === 'B' || t === 'STRONG' || t === 'TH') n.b = true;
    else if (t === 'I' || t === 'EM' || t === 'CITE' || t === 'VAR' || t === 'DFN') n.i = true;
    else if (t === 'U' || t === 'INS') n.u = true;
    else if (t === 'S' || t === 'STRIKE' || t === 'DEL') n.s = true;
    else if (t === 'SUB') { n.sub = true; n.sup = false; }
    else if (t === 'SUP') { n.sup = true; n.sub = false; }
    else if (t === 'CODE' || t === 'TT' || t === 'KBD' || t === 'SAMP' || t === 'PRE') n.font = 'Consolas';
    else if (t === 'MARK') n.hl = 'FFFF00';
    else if (t === 'FONT') {
      const col = cssToHex(el.getAttribute('color'));
      if (col) n.color = col;
      const face = el.getAttribute('face');
      if (face) n.font = face.split(',')[0].replace(/["']/g, '').trim();
      const sz = parseInt(el.getAttribute('size'), 10);
      if (sz >= 1 && sz <= 7) n.size = FONT_ATTR_PT[sz];
    }
    const st = el.style;
    if (st && st.length) {
      const fw = st.fontWeight;
      if (fw) { if (fw === 'bold' || fw === 'bolder' || parseInt(fw, 10) >= 600) n.b = true; else n.b = false; }
      const fi = st.fontStyle;
      if (fi) n.i = (fi === 'italic' || fi === 'oblique');
      const td = (st.textDecorationLine || st.textDecoration || '');
      if (td.includes('underline')) n.u = true;
      if (td.includes('line-through')) n.s = true;
      if (td === 'none' && t !== 'A') { n.u = false; n.s = false; }
      const col = cssToHex(st.color);
      if (col) n.color = col;
      const bg = cssToHex(st.backgroundColor);
      if (bg) n.hl = bg;
      if (st.fontFamily) n.font = st.fontFamily.split(',')[0].replace(/["']/g, '').trim();
      if (st.fontSize) {
        let px = NAMED_PX[st.fontSize] || lenPx(st.fontSize);
        if (px) n.size = Math.round(px * 0.75 * 2) / 2;
      }
      const va = st.verticalAlign;
      if (va === 'super') { n.sup = true; n.sub = false; } else if (va === 'sub') { n.sub = true; n.sup = false; }
      if (st.textTransform === 'uppercase') n.caps = true;
      if (st.fontVariant === 'small-caps' || st.fontVariantCaps === 'small-caps') n.smallCaps = true;
    }
    return n;
  }

  function rPrXml(c) {
    let x = '';
    if (c.link) x += '<w:rStyle w:val="Hyperlink"/>';
    if (c.font) { const f = xesc(c.font); x += `<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}"/>`; }
    if (c.b) x += '<w:b/><w:bCs/>';
    if (c.i) x += '<w:i/><w:iCs/>';
    if (c.caps) x += '<w:caps/>';
    if (c.smallCaps) x += '<w:smallCaps/>';
    if (c.s) x += '<w:strike/>';
    if (c.color) x += `<w:color w:val="${c.color}"/>`;
    if (c.size) { const v = Math.max(2, Math.round(c.size * 2)); x += `<w:sz w:val="${v}"/><w:szCs w:val="${v}"/>`; }
    let shd = '';
    if (c.hl) {
      if (HL_REV[c.hl]) x += `<w:highlight w:val="${HL_REV[c.hl]}"/>`;
      else shd = `<w:shd w:val="clear" w:color="auto" w:fill="${c.hl}"/>`;
    }
    if (c.u) x += '<w:u w:val="single"/>';
    x += shd;
    if (c.sup) x += '<w:vertAlign w:val="superscript"/>';
    else if (c.sub) x += '<w:vertAlign w:val="subscript"/>';
    return x ? `<w:rPr>${x}</w:rPr>` : '';
  }

  function textRuns(text, ctx) {
    // Split on tabs so they become real <w:tab/> elements.
    const rpr = rPrXml(ctx);
    const parts = text.split('\t');
    let out = '';
    parts.forEach((p, i) => {
      if (i > 0) out += `<w:r>${rpr}<w:tab/></w:r>`;
      if (p !== '') out += `<w:r>${rpr}<w:t xml:space="preserve">${xesc(p)}</w:t></w:r>`;
    });
    return out;
  }

  function newCol() { return { xml: '', real: false, br: false }; }

  function inline(n, ctx, col, S) {
    if (n.nodeType === 3) {
      let t = n.nodeValue.replace(/[​﻿]/g, '').replace(/ /g, ' ');
      if (!ctx.pre) t = t.replace(/[\r\n]+/g, ' ');
      if (t === '') return;
      if (!/\S/.test(t) && col.xml === '' && !ctx.keepWs) return; // leading blank
      col.xml += textRuns(t, ctx);
      if (/\S/.test(t) || /\t/.test(t)) col.real = true;
      return;
    }
    if (n.nodeType !== 1) return;
    const tag = n.tagName;
    if (tag === 'SCRIPT' || tag === 'STYLE' || tag === 'INPUT' || tag === 'BUTTON' || tag === 'SELECT' || tag === 'TEXTAREA') return;
    if (tag === 'BR') {
      const p = n.parentNode;
      const last = !n.nextSibling || (n.nextSibling.nodeType === 3 && !n.nextSibling.nodeValue.trim() && !n.nextSibling.nextSibling);
      if (last && p && (isBlock(p) || p.id === 'doc' || p.tagName === 'TD' || p.tagName === 'TH')) { col.br = col.br || col.xml === ''; return; }
      col.xml += `<w:r>${rPrXml(ctx)}<w:br/></w:r>`;
      col.real = true;
      return;
    }
    if (tag === 'IMG') {
      const x = imageXml(n, S);
      if (x) { col.xml += `<w:r>${x}</w:r>`; col.real = true; }
      else if (n.alt) { col.xml += textRuns(n.alt, ctx); col.real = true; }
      return;
    }
    if (tag === 'A') {
      const href = n.getAttribute('href') || '';
      const c2 = runCtx(n, ctx);
      if (href.charAt(0) === '#' && href.length > 1) {
        const inner = newCol();
        c2.link = true;
        for (const ch of n.childNodes) inline(ch, c2, inner, S);
        if (inner.xml) { col.xml += `<w:hyperlink w:anchor="${xesc(href.slice(1))}" w:history="1">${inner.xml}</w:hyperlink>`; col.real = true; }
        return;
      }
      if (/^(https?:|mailto:|ftp:|tel:)/i.test(href)) {
        const rid = 'rId' + (S.rid++);
        S.rels.push(`<Relationship Id="${rid}" Type="${REL}/hyperlink" Target="${xesc(href)}" TargetMode="External"/>`);
        const inner = newCol();
        c2.link = true;
        for (const ch of n.childNodes) inline(ch, c2, inner, S);
        if (inner.xml) { col.xml += `<w:hyperlink r:id="${rid}" w:history="1">${inner.xml}</w:hyperlink>`; col.real = true; }
        return;
      }
      for (const ch of n.childNodes) inline(ch, c2, col, S);
      return;
    }
    if (n.classList && n.classList.contains('fn') && S.fn) {
      const id = S.fn.length + 1;
      S.fn.push({ id, text: n.dataset.t || '' });
      col.xml += `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
      col.real = true;
      return;
    }
    if (n.classList && n.classList.contains('en') && S.en) {
      const id = S.en.length + 1;
      S.en.push({ id, text: n.dataset.t || '' });
      col.xml += `<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="${id}"/></w:r>`;
      col.real = true;
      return;
    }
    if ((tag === 'INS' || tag === 'DEL') && n.classList.contains('tc') && S.main) {
      const inner = newCol();
      for (const ch of n.childNodes) inline(ch, Object.assign({}, ctx, { keepWs: true }), inner, S);
      if (inner.xml) {
        const at = `w:id="${S.rev++}" w:author="${xesc(n.dataset.a || 'Author')}" w:date="${xesc(n.dataset.d || S.now)}"`;
        if (tag === 'INS') col.xml += `<w:ins ${at}>${inner.xml}</w:ins>`;
        else col.xml += `<w:del ${at}>${inner.xml.replace(/<w:t( xml:space="preserve")?>/g, '<w:delText xml:space="preserve">').replace(/<\/w:t>/g, '</w:delText>')}</w:del>`;
        col.real = true;
      }
      return;
    }
    if (n.classList && n.classList.contains('bm')) {
      if (S.main && n.dataset.n) { const id = S.rev++; col.xml += `<w:bookmarkStart w:id="${id}" w:name="${xesc(n.dataset.n)}"/><w:bookmarkEnd w:id="${id}"/>`; }
      return;
    }
    if (n.classList && n.classList.contains('cite')) {
      const tag = (n.dataset.tag || 'Source').replace(/[^\w-]/g, '_');
      col.xml += `<w:sdt><w:sdtPr><w:id w:val="${100000 + (S.rev++)}"/><w:citation/></w:sdtPr><w:sdtContent><w:fldSimple w:instr=" CITATION ${tag} \\l 1033 "><w:r>${rPrXml(ctx)}<w:t xml:space="preserve">${xesc(n.textContent || '')}</w:t></w:r></w:fldSimple></w:sdtContent></w:sdt>`;
      col.real = true;
      return;
    }
    if (n.classList && n.classList.contains('mf')) {
      const nm = (n.dataset.n || 'Field').replace(/[^\w]/g, '_');
      col.xml += `<w:fldSimple w:instr=" MERGEFIELD ${nm} \\* MERGEFORMAT "><w:r>${rPrXml(ctx)}<w:t>«${xesc(nm)}»</w:t></w:r></w:fldSimple>`;
      col.real = true;
      return;
    }
    if (n.classList && n.classList.contains('eq')) {
      if (S.main) { col.xml += '<m:oMath>' + Eq.toOmml(Eq.parse(n.dataset.src || '')) + '</m:oMath>'; col.real = true; }
      else { col.xml += textRuns(n.dataset.src || '', ctx); col.real = true; }
      return;
    }
    if (n.classList && n.classList.contains('cmt') && S.cm) {
      const id = S.cm.length;
      S.cm.push({ id, text: n.dataset.t || '', author: n.dataset.a || 'Author' });
      const inner = newCol();
      const c2 = runCtx(n, ctx);
      for (const ch of n.childNodes) inline(ch, c2, inner, S);
      if (inner.xml) {
        col.xml += `<w:commentRangeStart w:id="${id}"/>${inner.xml}<w:commentRangeEnd w:id="${id}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;
        col.real = true;
      }
      return;
    }
    if (n.classList && n.classList.contains('tocdots')) { col.xml += '<w:r><w:tab/></w:r>'; return; }
    if (n.classList && n.classList.contains('fld')) {
      const f = n.dataset.f === 'NUMPAGES' ? 'NUMPAGES' : 'PAGE';
      col.xml += `<w:fldSimple w:instr=" ${f} "><w:r>${rPrXml(ctx)}<w:t>1</w:t></w:r></w:fldSimple>`;
      col.real = true;
      return;
    }
    const c2 = runCtx(n, ctx);
    if (n.classList && n.classList.contains('pb')) return;
    for (const ch of n.childNodes) inline(ch, c2, col, S);
  }

  function imageXml(img, S) {
    const d = dataUrlToBytes(img.getAttribute('src') || '');
    if (!d) return '';
    const ext = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp' }[d.mime];
    if (!ext) return '';
    const nw = img.naturalWidth || 0, nh = img.naturalHeight || 0;
    let w = lenPx(img.style.width) || parseFloat(img.getAttribute('width')) || nw || 200;
    let h = lenPx(img.style.height) || parseFloat(img.getAttribute('height'));
    if (!h) h = (nw && nh) ? w * nh / nw : w * 0.75;
    const cx = Math.round(w * 9525), cy = Math.round(h * 9525);
    const name = 'image' + (++S.imgN) + '.' + ext;
    S.media.push({ path: 'word/media/' + name, bytes: d.bytes });
    S.exts.add(ext);
    const rid = 'rId' + (S.rid++);
    S.rels.push(`<Relationship Id="${rid}" Type="${REL}/image" Target="media/${name}"/>`);
    const id = S.docPr++;
    const alt = xesc(img.getAttribute('alt') || name);
    return `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="Picture ${id}" descr="${alt}"/>` +
      `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
  }

  /* ---------- equations: linear text <-> tree <-> OMML / HTML ---------- */
  const Eq = (function () {
    const WORD = /[\p{L}\p{N}.]/u;
    function parse(src) {
      let i = 0;
      const strip = (x) => x.t === 'd' ? x.e : [x];
      const atom = () => {
        if (i >= src.length) return { t: 'r', s: '' };
        if (src[i] === '(') { i++; const e = seq(); if (src[i] === ')') i++; return { t: 'd', e }; }
        if (src.startsWith('sqrt(', i)) { i += 5; const e = seq(); if (src[i] === ')') i++; return { t: 'rad', e }; }
        let s = src[i++];
        if (WORD.test(s)) while (i < src.length && WORD.test(src[i]) && !src.startsWith('sqrt(', i)) s += src[i++];
        return { t: 'r', s };
      };
      const script = () => {
        let a = atom();
        while (i < src.length && (src[i] === '^' || src[i] === '_')) { const op = src[i++]; const b = atom(); a = { t: op === '^' ? 'sup' : 'sub', b: [a], e: strip(b) }; }
        return a;
      };
      const frac = () => {
        let a = script();
        while (i < src.length && src[i] === '/') { i++; const b = script(); a = { t: 'f', n: strip(a), d: strip(b) }; }
        return a;
      };
      function seq() { const out = []; while (i < src.length && src[i] !== ')') out.push(frac()); return out; }
      const all = [];
      while (i < src.length) { if (src[i] === ')') { all.push({ t: 'r', s: ')' }); i++; } else all.push(...seq()); }
      return all;
    }
    function toOmml(list) {
      return list.map(x => {
        switch (x.t) {
          case 'r': return x.s === '' ? '' : `<m:r><m:t xml:space="preserve">${xesc(x.s)}</m:t></m:r>`;
          case 'f': return `<m:f><m:num>${toOmml(x.n)}</m:num><m:den>${toOmml(x.d)}</m:den></m:f>`;
          case 'sup': return `<m:sSup><m:e>${toOmml(x.b)}</m:e><m:sup>${toOmml(x.e)}</m:sup></m:sSup>`;
          case 'sub': return `<m:sSub><m:e>${toOmml(x.b)}</m:e><m:sub>${toOmml(x.e)}</m:sub></m:sSub>`;
          case 'rad': return `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${toOmml(x.e)}</m:e></m:rad>`;
          case 'd': return `<m:d><m:e>${toOmml(x.e)}</m:e></m:d>`;
        }
        return '';
      }).join('');
    }
    function toHtml(list) {
      const h = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      return list.map(x => {
        switch (x.t) {
          case 'r': return h(x.s);
          case 'f': return `<span class="eqf"><span>${toHtml(x.n)}</span><span>${toHtml(x.d)}</span></span>`;
          case 'sup': return `${toHtml(x.b)}<sup>${toHtml(x.e)}</sup>`;
          case 'sub': return `${toHtml(x.b)}<sub>${toHtml(x.e)}</sub>`;
          case 'rad': return `&radic;<span class="eqr">${toHtml(x.e)}</span>`;
          case 'd': return `(${toHtml(x.e)})`;
        }
        return '';
      }).join('');
    }
    function fromOmml(el) {
      const part = (p, ln) => { for (const c of p.children) if (c.localName === ln) return fromOmml(c); return ''; };
      const arg = (s) => /^[\p{L}\p{N}.]+$/u.test(s) ? s : '(' + s + ')';
      let out = '';
      for (const c of el.children) {
        switch (c.localName) {
          case 'r': for (const t of c.children) if (t.localName === 't') out += t.textContent; break;
          case 'f': out += arg(part(c, 'num')) + '/' + arg(part(c, 'den')); break;
          case 'sSup': out += arg(part(c, 'e')) + '^' + arg(part(c, 'sup')); break;
          case 'sSub': out += arg(part(c, 'e')) + '_' + arg(part(c, 'sub')); break;
          case 'sSubSup': out += arg(part(c, 'e')) + '_' + arg(part(c, 'sub')) + '^' + arg(part(c, 'sup')); break;
          case 'rad': out += 'sqrt(' + part(c, 'e') + ')'; break;
          case 'd': out += '(' + [...c.children].filter(k => k.localName === 'e').map(fromOmml).join(',') + ')'; break;
          case 'nary': { const ch = c.getElementsByTagNameNS('*', 'chr')[0]; out += ((ch && (ch.getAttribute('m:val') || ch.getAttributeNS('*', 'val'))) || '∫') + '_' + arg(part(c, 'sub')) + '^' + arg(part(c, 'sup')) + ' ' + part(c, 'e'); break; }
          case 'rPr': case 'ctrlPr': case 'fPr': case 'radPr': case 'dPr': case 'sSupPr': case 'sSubPr': case 'naryPr': case 'deg': break;
          default: out += fromOmml(c);
        }
      }
      return out;
    }
    function span(src) {
      const s = document.createElement('span');
      s.className = 'eq'; s.dataset.src = src; s.setAttribute('contenteditable', 'false');
      s.innerHTML = toHtml(parse(src));
      return s;
    }
    return { parse, toOmml, toHtml, fromOmml, span };
  })();

  function textboxXml(n, S) {
    const w = n.offsetWidth || lenPx(n.style.width) || 260, h = n.offsetHeight || 80;
    const cx = Math.round(w * 9525), cy = Math.round(h * 9525);
    const id = S.docPr++;
    let inner = blocks(n, { run: {}, textW: Math.round(w * 15) }, S, false);
    if (!inner || inner.endsWith('</w:tbl>')) inner += '<w:p/>';
    const fill = cssToHex(n.style.backgroundColor) || 'FFFFFF';
    return `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="Text Box ${id}"/><wp:cNvGraphicFramePr/>` +
      `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
      `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>` +
      `<wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
  }

  function paraCtx(el, pc) {
    const n = Object.assign({}, pc);
    const t = el.tagName;
    n.run = runCtx(el, pc.run || {});
    n.run.hl = (pc.run || {}).hl; // block backgrounds are paragraph shading, not run highlights
    if (/^H[1-6]$/.test(t)) n.style = 'Heading' + t[1];
    else if (el.classList.contains('title')) n.style = 'Title';
    else if (el.classList.contains('subtitle')) n.style = 'Subtitle';
    else if (el.classList.contains('quote')) n.style = 'Quote';
    else if (el.classList.contains('toc')) { n.style = el.classList.contains('toc3') ? 'TOC3' : el.classList.contains('toc2') ? 'TOC2' : 'TOC1'; n.tocTab = true; }
    if (t === 'PRE') n.run.pre = true;
    const st = el.style;
    const ta = (st && st.textAlign) || el.getAttribute('align');
    if (ta) {
      const m = { left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', justify: 'both' }[ta];
      if (m) n.jc = m;
    }
    if (t === 'CENTER') n.jc = 'center';
    if (t !== 'LI' && st) {
      const ml = lenPx(st.marginLeft);
      if (ml) n.left = (pc.left || 0) + Math.round(ml * 15);
      const ti = lenPx(st.textIndent);
      if (ti) n.first = Math.round(ti * 15);
    }
    if (t === 'BLOCKQUOTE' && !el.classList.contains('quote')) {
      if (!(st && lenPx(st.marginLeft))) n.left = (pc.left || 0) + 720;
    }
    if (st) {
      const mt = lenPx(st.marginTop), mb = lenPx(st.marginBottom);
      if (mt != null && st.marginTop) n.before = Math.max(0, Math.round(mt * 15));
      if (mb != null && st.marginBottom) n.after = Math.max(0, Math.round(mb * 15));
      const lh = st.lineHeight;
      if (lh) {
        if (/^[\d.]+$/.test(lh)) n.line = Math.round(parseFloat(lh) * 240);
        else if (/%$/.test(lh)) n.line = Math.round(parseFloat(lh) * 2.4);
      }
      const shd = cssToHex(st.backgroundColor);
      if (shd && t !== 'LI') n.shd = shd;
    }
    return n;
  }

  function pPrXml(pc, numPr) {
    let x = '';
    let style = pc.style;
    if (numPr && !style) style = 'ListParagraph';
    if (style) x += `<w:pStyle w:val="${style}"/>`;
    if (numPr) x += numPr;
    if (pc.shd) x += `<w:shd w:val="clear" w:color="auto" w:fill="${pc.shd}"/>`;
    if (pc.tocTab) x += `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${pc.textW || 9350}"/></w:tabs>`;
    if (pc.before != null || pc.after != null || pc.line != null) {
      x += '<w:spacing' + (pc.before != null ? ` w:before="${pc.before}"` : '') + (pc.after != null ? ` w:after="${pc.after}"` : '') +
        (pc.line != null ? ` w:line="${pc.line}" w:lineRule="auto"` : '') + '/>';
    }
    if (!numPr && (pc.left || pc.first)) {
      x += '<w:ind' + (pc.left ? ` w:left="${pc.left}"` : '') + (pc.first ? (pc.first < 0 ? ` w:hanging="${-pc.first}"` : ` w:firstLine="${pc.first}"`) : '') + '/>';
    }
    if (pc.jc) x += `<w:jc w:val="${pc.jc}"/>`;
    return x ? `<w:pPr>${x}</w:pPr>` : '';
  }

  function para(runsXml, pc, S) {
    let numPr = '';
    if (pc.list && !pc.list.used) {
      pc.list.used = true;
      numPr = `<w:numPr><w:ilvl w:val="${pc.list.ilvl}"/><w:numId w:val="${pc.list.numId}"/></w:numPr>`;
    }
    return `<w:p>${pPrXml(pc, numPr)}${runsXml}</w:p>`;
  }

  function hasBlockChild(el) {
    for (const c of el.children) if (isBlock(c) || c.classList.contains('pb')) return true;
    return false;
  }

  function blocks(parent, pc, S, ex) {
    let out = '';
    let col = newCol();
    const flush = () => {
      if ((col.xml || col.br) && (col.real || col.br || ex)) out += para(col.xml, pc, S);
      col = newCol();
    };
    for (const n of parent.childNodes) {
      if (n.nodeType === 1 && (isBlock(n) || n.classList.contains('pb'))) {
        flush();
        const x = block(n, pc, S);
        if (x.startsWith('<w:tbl>') && out.endsWith('</w:tbl>')) out += '<w:p/>';
        out += x;
      } else inline(n, pc.run || {}, col, S);
    }
    flush();
    return out;
  }

  function block(n, pc, S) {
    const t = n.tagName;
    if (n.classList.contains('pb')) return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    if (n.classList.contains('tbox')) return S.main ? textboxXml(n, S) : blocks(n, pc, S, false);
    if (t === 'TABLE') return tableXml(n, S);
    if (t === 'UL' || t === 'OL') return listXml(n, pc, S);
    if (t === 'HR') return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>';
    const pc2 = paraCtx(n, pc);
    const isPara = /^(P|H[1-6]|LI|PRE|DD|DT)$/.test(t) || (t === 'DIV' && !hasBlockChild(n));
    let x = blocks(n, pc2, S, isPara);
    if (!x && isPara) x = para('', pc2, S);
    return x;
  }

  function listXml(el, pc, S) {
    const ordered = el.tagName === 'OL';
    const parent = pc.list;
    const ilvl = parent ? Math.min(8, parent.ilvl + 1) : 0;
    let numId;
    if (!ordered) numId = 1;
    else if (parent && parent.ordered) numId = parent.numId;
    else {
      numId = S.nextNum++;
      S.nums.push({ numId, start: parseInt(el.getAttribute('start'), 10) || 1 });
    }
    let out = '';
    for (const li of el.children) {
      if (li.tagName !== 'LI') {
        if (li.tagName === 'UL' || li.tagName === 'OL') out += listXml(li, Object.assign({}, pc, { list: { numId, ilvl, ordered, used: true } }), S);
        continue;
      }
      const pc2 = Object.assign({}, pc, { list: { numId, ilvl, ordered, used: false }, left: undefined, first: undefined });
      out += block(li, pc2, S);
    }
    return out;
  }

  function tableXml(tbl, S) {
    const rows = [...tbl.rows];
    if (!rows.length) return '';
    const carry = [];
    const plan = [];
    let ncols = 0;
    for (let r = 0; r < rows.length; r++) {
      const rowPlan = [];
      let c = 0;
      const skip = () => {
        while (carry[c] && carry[c].left > 0) {
          rowPlan.push({ cont: true, col: c, span: carry[c].span, cell: carry[c].cell });
          carry[c].left--;
          c += carry[c].span;
        }
      };
      for (const cell of rows[r].cells) {
        skip();
        const cs = Math.max(1, cell.colSpan || 1);
        const rs = Math.min(Math.max(1, cell.rowSpan || 1), rows.length - r);
        rowPlan.push({ cell, col: c, span: cs, restart: rs > 1 });
        if (rs > 1) carry[c] = { left: rs - 1, span: cs, cell };
        c += cs;
      }
      skip();
      ncols = Math.max(ncols, c);
      plan.push(rowPlan);
    }
    // column widths (twips)
    let grid = null;
    for (const rp of plan) {
      if (rp.length === ncols && rp.every(e => !e.cont && e.span === 1)) {
        const ws = rp.map(e => e.cell.offsetWidth * 15);
        if (ws.every(w => w > 0)) { grid = ws.map(Math.round); break; }
      }
    }
    if (!grid) { const w = Math.floor(S.textW / Math.max(1, ncols)); grid = Array(ncols).fill(w); }
    const total = grid.reduce((a, b) => a + b, 0);
    const noBorder = tbl.classList.contains('nb');
    const bdr = noBorder ? '' : '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(s => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>';
    let x = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="${total}" w:type="dxa"/>${noBorder ? '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(s => `<w:${s} w:val="nil"/>`).join('') + '</w:tblBorders>' : bdr}<w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>`;
    x += '<w:tblGrid>' + grid.map(w => `<w:gridCol w:w="${w}"/>`).join('') + '</w:tblGrid>';
    for (const rp of plan) {
      x += '<w:tr>';
      for (const e of rp) {
        let w = 0;
        for (let k = e.col; k < e.col + e.span; k++) w += grid[k] || 0;
        let pr = `<w:tcW w:w="${w}" w:type="dxa"/>`;
        if (e.span > 1) pr += `<w:gridSpan w:val="${e.span}"/>`;
        if (e.cont) pr += '<w:vMerge/>';
        else if (e.restart) pr += '<w:vMerge w:val="restart"/>';
        const bg = cssToHex(e.cell.style.backgroundColor) || cssToHex(e.cell.getAttribute('bgcolor'));
        if (bg) pr += `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`;
        let content = '';
        if (!e.cont) {
          const base = { run: runCtx(e.cell, {}) };
          base.run.hl = undefined;
          const ta = e.cell.style.textAlign;
          if (ta) base.jc = { left: 'left', center: 'center', right: 'right', justify: 'both' }[ta];
          content = blocks(e.cell, base, S, false);
        }
        if (!content || content.endsWith('</w:tbl>')) content += '<w:p/>';
        x += `<w:tc><w:tcPr>${pr}</w:tcPr>${content}</w:tc>`;
      }
      x += '</w:tr>';
    }
    return x + '</w:tbl>';
  }

  /* ---------- static parts ---------- */
  const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

  const STYLE_IDS = { normal: 'Normal', title: 'Title', subtitle: 'Subtitle', h1: 'Heading1', h2: 'Heading2', h3: 'Heading3', h4: 'Heading4', h5: 'Heading5', h6: 'Heading6', quote: 'Quote' };
  const STYLE_DEFAULTS = {
    normal: { ff: 'Calibri', sz: 11, color: '000000', b: false, i: false },
    title: { ff: 'Calibri Light', sz: 28, color: '000000', b: false, i: false },
    subtitle: { ff: 'Calibri', sz: 11, color: '595959', b: false, i: false },
    h1: { ff: 'Calibri Light', sz: 16, color: '2F5496', b: false, i: false },
    h2: { ff: 'Calibri Light', sz: 13, color: '2F5496', b: false, i: false },
    h3: { ff: 'Calibri Light', sz: 12, color: '1F3763', b: false, i: false },
    h4: { ff: 'Calibri Light', sz: 11, color: '2F5496', b: false, i: true },
    h5: { ff: 'Calibri Light', sz: 11, color: '2F5496', b: false, i: false },
    h6: { ff: 'Calibri Light', sz: 11, color: '1F3763', b: false, i: false },
    quote: { ff: 'Calibri', sz: 11, color: '404040', b: false, i: true },
  };
  window.DocxStyleDefaults = STYLE_DEFAULTS;
  // user-modified styles (page.styles) are written into styles.xml so Word shows them as real styles
  function stylesXml(ov) {
    let x = stylesXmlBase();
    if (!ov) return x;
    const xe = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
    for (const key of Object.keys(ov)) {
      const id = STYLE_IDS[key];
      if (!id || !ov[key]) continue;
      const o = Object.assign({}, STYLE_DEFAULTS[key], ov[key]);
      const re = new RegExp('(<w:style [^>]*w:styleId="' + id + '">)([\\s\\S]*?)(</w:style>)');
      x = x.replace(re, (m, a, inner, z) => {
        const old = (/<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(inner) || [0, ''])[1];
        const keep = (tag) => (new RegExp('<w:' + tag + ' [^>]*/>').exec(old) || [''])[0];
        const f = xe(o.ff || 'Calibri'), hp = Math.max(2, Math.round((+o.sz || 11) * 2));
        const col = /^[0-9A-Fa-f]{6}$/.test(o.color || '') ? o.color.toUpperCase() : '000000';
        const rpr = `<w:rPr><w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}"/>${o.b ? '<w:b/><w:bCs/>' : ''}${o.i ? '<w:i/><w:iCs/>' : ''}<w:color w:val="${col}"/>${keep('spacing')}${keep('kern')}<w:sz w:val="${hp}"/><w:szCs w:val="${hp}"/></w:rPr>`;
        return a + inner.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/, '') + rpr + z;
      });
    }
    return x;
  }
  function stylesXmlBase() {
    const hd = (id, name, next, size, color, extra, font, rpr) =>
      `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>` +
      `<w:pPr><w:keepNext/><w:keepLines/>${extra}</w:pPr><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"/>${rpr || ''}<w:color w:val="${color}"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:style>`;
    const LIGHT = 'Calibri Light';
    return XML + `<w:styles xmlns:w="${NS_W}">` +
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
      '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' +
      '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
      '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' +
      hd('Heading1', 'heading 1', 'Normal', 32, '2F5496', '<w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="0"/>', LIGHT) +
      hd('Heading2', 'heading 2', 'Normal', 26, '2F5496', '<w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="1"/>', LIGHT) +
      hd('Heading3', 'heading 3', 'Normal', 24, '1F3763', '<w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="2"/>', LIGHT) +
      hd('Heading4', 'heading 4', 'Normal', 22, '2F5496', '<w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="3"/>', 'Calibri Light', '<w:i/><w:iCs/>') +
      hd('Heading5', 'heading 5', 'Normal', 22, '2F5496', '<w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="4"/>', LIGHT) +
      hd('Heading6', 'heading 6', 'Normal', 22, '1F3763', '<w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="5"/>', LIGHT) +
      '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:eastAsia="Calibri Light" w:cs="Calibri Light"/><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="11"/><w:qFormat/><w:rPr><w:color w:val="595959"/><w:spacing w:val="15"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="29"/><w:qFormat/><w:pPr><w:spacing w:before="200" w:after="160"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="404040"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="CommentText"><w:name w:val="annotation text"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="EndnoteText"><w:name w:val="endnote text"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="endnote reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>' +
      [1, 2, 3].map(l => `<w:style w:type="paragraph" w:styleId="TOC${l}"><w:name w:val="toc ${l}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="100"/><w:ind w:left="${(l - 1) * 220}"/></w:pPr></w:style>`).join('') +
      '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:spacing w:after="0"/><w:ind w:left="720"/></w:pPr></w:style>' +
      '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>' +
      '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>' +
      '</w:styles>';
  }

  function numberingXml(nums) {
    const bulletChars = ['•', 'o', '▪'];
    const bulletFonts = ['Arial', 'Courier New', 'Arial'];
    const numFmts = ['decimal', 'lowerLetter', 'lowerRoman'];
    let a0 = '', a1 = '';
    for (let l = 0; l < 9; l++) {
      const ind = `<w:pPr><w:ind w:left="${720 * (l + 1)}" w:hanging="360"/></w:pPr>`;
      a0 += `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${bulletChars[l % 3]}"/><w:lvlJc w:val="left"/>${ind}<w:rPr><w:rFonts w:ascii="${bulletFonts[l % 3]}" w:hAnsi="${bulletFonts[l % 3]}" w:cs="${bulletFonts[l % 3]}" w:hint="default"/></w:rPr></w:lvl>`;
      a1 += `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="${numFmts[l % 3]}"/><w:lvlText w:val="%${l + 1}."/><w:lvlJc w:val="left"/>${ind}</w:lvl>`;
    }
    let x = XML + `<w:numbering xmlns:w="${NS_W}">` +
      `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${a0}</w:abstractNum>` +
      `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${a1}</w:abstractNum>` +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
    for (const n of nums) x += `<w:num w:numId="${n.numId}"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="${n.start}"/></w:lvlOverride></w:num>`;
    return x + '</w:numbering>';
  }

  async function write(root, page, meta) {
    page = page || { w: 12240, h: 15840, t: 1440, r: 1440, b: 1440, l: 1440 };
    meta = meta || {};
    const S = { rels: [], media: [], rid: 10, imgN: 0, nums: [], nextNum: 2, docPr: 1, exts: new Set(['png', 'jpeg']), textW: page.w - page.l - page.r, fn: [], cm: [], en: [], main: true, rev: 1, now: new Date().toISOString().replace(/\.\d+Z$/, 'Z') };
    let body = blocks(root, { run: {}, textW: page.w - page.l - page.r }, S);
    if (!body || body.endsWith('</w:tbl>')) body += '<w:p/>';
    const hfXml = (html, tag) => {
      const d = document.createElement('div');
      d.innerHTML = html;
      d.querySelectorAll('img').forEach(i => i.remove());
      d.querySelectorAll('a').forEach(a => a.replaceWith(...a.childNodes));
      const S2 = { rels: [], media: [], rid: 10, imgN: 0, nums: [], nextNum: 2, docPr: 1, exts: new Set(), textW: S.textW };
      let b = blocks(d, { run: {} }, S2);
      if (!b || b.endsWith('</w:tbl>')) b += '<w:p/>';
      return XML + `<w:${tag} xmlns:w="${NS_W}" xmlns:r="${NS_R}">${b}</w:${tag}>`;
    };
    const hasText = (h) => { const d = document.createElement('div'); d.innerHTML = h || ''; return !!d.textContent.trim() || !!d.querySelector('.fld'); };
    const hdrOn = hasText(meta.header), ftrOn = hasText(meta.footer);
    if (hdrOn) S.rels.push(`<Relationship Id="rIdHdr1" Type="${REL}/header" Target="header1.xml"/>`);
    if (ftrOn) S.rels.push(`<Relationship Id="rIdFtr1" Type="${REL}/footer" Target="footer1.xml"/>`);
    const sect = `<w:sectPr>${hdrOn ? '<w:headerReference w:type="default" r:id="rIdHdr1"/>' : ''}${ftrOn ? '<w:footerReference w:type="default" r:id="rIdFtr1"/>' : ''}<w:pgSz w:w="${page.w}" w:h="${page.h}"${page.w > page.h ? ' w:orient="landscape"' : ''}/><w:pgMar w:top="${page.t}" w:right="${page.r}" w:bottom="${page.b}" w:left="${page.l}" w:header="708" w:footer="708" w:gutter="0"/>${page.cols > 1 ? `<w:cols w:num="${page.cols}" w:space="720"/>` : ''}</w:sectPr>`;
    if (S.fn.length) S.rels.push(`<Relationship Id="rIdFn1" Type="${REL}/footnotes" Target="footnotes.xml"/>`);
    if (S.cm.length) S.rels.push(`<Relationship Id="rIdCm1" Type="${REL}/comments" Target="comments.xml"/>`);
    if (S.en.length) S.rels.push(`<Relationship Id="rIdEn1" Type="${REL}/endnotes" Target="endnotes.xml"/>`);
    const srcs = (page && page.sources) || [];
    if (srcs.length) S.rels.push(`<Relationship Id="rIdBib1" Type="${REL}/customXml" Target="../customXml/item1.xml"/>`);
    const docXml = XML + `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${body}${sect}</w:body></w:document>`;

    const zw = Zip.writer();
    const ct = XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      [...S.exts].map(e => `<Default Extension="${e}" ContentType="image/${e}"/>`).join('') +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
      (S.cm.length ? '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' : '') +
      (S.fn.length ? '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' : '') +
      (S.en.length ? '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>' : '') +
      (hdrOn ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : '') +
      (ftrOn ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' : '') +
      (srcs.length ? '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>' : '') +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>';
    if (S.cm.length) {
      const dt = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
      zw.add('word/comments.xml', XML + `<w:comments xmlns:w="${NS_W}" xmlns:r="${NS_R}">` + S.cm.map(c => `<w:comment w:id="${c.id}" w:author="${xesc(c.author)}" w:date="${dt}" w:initials="${xesc((c.author || 'A').slice(0, 2))}"><w:p><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r><w:r><w:t xml:space="preserve">${xesc(c.text)}</w:t></w:r></w:p></w:comment>`).join('') + '</w:comments>');
    }
    if (S.fn.length) {
      const sep = (type, id, inner) => `<w:footnote w:type="${type}" w:id="${id}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r>${inner}</w:r></w:p></w:footnote>`;
      zw.add('word/footnotes.xml', XML + `<w:footnotes xmlns:w="${NS_W}" xmlns:r="${NS_R}">${sep('separator', -1, '<w:separator/>')}${sep('continuationSeparator', 0, '<w:continuationSeparator/>')}` +
        S.fn.map(f => `<w:footnote w:id="${f.id}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> ${xesc(f.text)}</w:t></w:r></w:p></w:footnote>`).join('') + '</w:footnotes>');
    }
    if (S.en.length) {
      const sep = (type, id, inner) => `<w:endnote w:type="${type}" w:id="${id}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r>${inner}</w:r></w:p></w:endnote>`;
      zw.add('word/endnotes.xml', XML + `<w:endnotes xmlns:w="${NS_W}" xmlns:r="${NS_R}">${sep('separator', -1, '<w:separator/>')}${sep('continuationSeparator', 0, '<w:continuationSeparator/>')}` +
        S.en.map(f => `<w:endnote w:id="${f.id}"><w:p><w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> ${xesc(f.text)}</w:t></w:r></w:p></w:endnote>`).join('') + '</w:endnotes>');
    }
    if (srcs.length) {
      const B = 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography';
      const person = (a) => { const p = String(a).split(','); const last = p[0].trim(), first = (p[1] || '').trim(); return `<b:Person><b:Last>${xesc(last)}</b:Last>${first ? `<b:First>${xesc(first)}</b:First>` : ''}</b:Person>`; };
      const TYPES = { book: 'Book', article: 'JournalArticle', web: 'InternetSite', report: 'Report' };
      zw.add('customXml/item1.xml', XML + `<b:Sources SelectedStyle="\\APASixthEditionOfficeOnline.xsl" StyleName="APA" Version="6" xmlns:b="${B}" xmlns="${B}">` + srcs.map((s, i) => {
        const authors = String(s.author || '').split(';').map(a => a.trim()).filter(Boolean);
        return `<b:Source><b:Tag>${xesc(s.tag)}</b:Tag><b:SourceType>${TYPES[s.type] || 'Book'}</b:SourceType><b:Guid>{${('00000000' + (i + 1)).slice(-8)}-1A2B-4C3D-8E9F-0123456789AB}</b:Guid>` +
          (authors.length ? `<b:Author><b:Author><b:NameList>${authors.map(person).join('')}</b:NameList></b:Author></b:Author>` : '') +
          `<b:Title>${xesc(s.title || '')}</b:Title>${s.year ? `<b:Year>${xesc(s.year)}</b:Year>` : ''}${s.publisher ? `<b:Publisher>${xesc(s.publisher)}</b:Publisher>` : ''}${s.city ? `<b:City>${xesc(s.city)}</b:City>` : ''}` +
          (s.journal ? `<b:JournalName>${xesc(s.journal)}</b:JournalName>` : '') + (s.url ? `<b:URL>${xesc(s.url)}</b:URL>` : '') + '</b:Source>';
      }).join('') + '</b:Sources>');
      zw.add('customXml/itemProps1.xml', XML + `<ds:datastoreItem ds:itemID="{6A1B2C3D-4E5F-4A6B-9C8D-7E6F5A4B3C2D}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="${B}"/></ds:schemaRefs></ds:datastoreItem>`);
      zw.add('customXml/_rels/item1.xml.rels', XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/customXmlProps" Target="itemProps1.xml"/></Relationships>`);
    }
    if (hdrOn) zw.add('word/header1.xml', hfXml(meta.header, 'hdr'));
    if (ftrOn) zw.add('word/footer1.xml', hfXml(meta.footer, 'ftr'));
    zw.add('[Content_Types].xml', ct);
    zw.add('_rels/.rels', XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/></Relationships>`);
    zw.add('word/document.xml', docXml);
    zw.add('word/_rels/document.xml.rels', XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${REL}/numbering" Target="numbering.xml"/><Relationship Id="rId3" Type="${REL}/settings" Target="settings.xml"/>${S.rels.join('')}</Relationships>`);
    zw.add('word/styles.xml', stylesXml(page && page.styles));
    zw.add('word/numbering.xml', numberingXml(S.nums));
    zw.add('word/settings.xml', XML + `<w:settings xmlns:w="${NS_W}"><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/>${meta.track ? '<w:trackRevisions/>' : ''}${S.fn.length ? '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>' : ''}${S.en.length ? '<w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>' : ''}<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`);
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    zw.add('docProps/core.xml', XML + `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xesc(meta.title || '')}</dc:title><dc:creator>Write</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
    zw.add('docProps/app.xml', XML + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Write</Application></Properties>');
    for (const m of S.media) zw.add(m.path, m.bytes);
    return zw.toBlob(DOCX_MIME);
  }

  /* ====================================================================
     READER
     ==================================================================== */
  const kids = (el, ln) => { const r = []; for (let c = el && el.firstElementChild; c; c = c.nextElementSibling) if (c.localName === ln) r.push(c); return r; };
  const kid = (el, ln) => { for (let c = el && el.firstElementChild; c; c = c.nextElementSibling) if (c.localName === ln) return c; return null; };
  const wa = (el, n) => { if (!el) return null; const v = el.getAttributeNS(NS_W, n); return v != null ? v : el.getAttribute('w:' + n); };
  const on = (e) => { if (!e) return undefined; const v = wa(e, 'val'); return !(v === '0' || v === 'false' || v === 'off'); };
  const int = (v) => { const n = parseInt(v, 10); return isNaN(n) ? undefined : n; };
  const defined = (o) => { const r = {}; for (const k in o) if (o[k] !== undefined) r[k] = o[k]; return r; };

  function parseXml(str) {
    const d = new DOMParser().parseFromString(str, 'application/xml');
    if (d.getElementsByTagName('parsererror').length) throw new Error('Could not parse XML in the document.');
    return d.documentElement;
  }
  function dirname(p) { const i = p.lastIndexOf('/'); return i < 0 ? '' : p.slice(0, i); }
  function resolvePath(base, target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = (base ? base.split('/') : []);
    for (const s of target.split('/')) { if (s === '..') parts.pop(); else if (s !== '.') parts.push(s); }
    return parts.join('/');
  }
  function parseRels(xml, base) {
    const out = {};
    if (!xml) return out;
    const root = parseXml(xml);
    for (const r of root.children) {
      const mode = r.getAttribute('TargetMode');
      out[r.getAttribute('Id')] = { type: r.getAttribute('Type'), target: r.getAttribute('Target'), path: mode === 'External' ? null : resolvePath(base, r.getAttribute('Target')), external: mode === 'External' };
    }
    return out;
  }

  function parseRPr(rpr, theme) {
    const p = {};
    if (!rpr) return p;
    let e;
    if ((e = kid(rpr, 'b'))) p.b = on(e);
    if ((e = kid(rpr, 'i'))) p.i = on(e);
    if ((e = kid(rpr, 'strike'))) p.s = on(e);
    else if ((e = kid(rpr, 'dstrike'))) p.s = on(e);
    if ((e = kid(rpr, 'caps'))) p.caps = on(e);
    if ((e = kid(rpr, 'smallCaps'))) p.smallCaps = on(e);
    if ((e = kid(rpr, 'vanish'))) p.vanish = on(e);
    if ((e = kid(rpr, 'u'))) { const v = wa(e, 'val'); p.u = !!v && v !== 'none'; }
    if ((e = kid(rpr, 'color'))) { const v = wa(e, 'val'); if (v) p.color = v === 'auto' ? 'auto' : v.toUpperCase(); }
    if ((e = kid(rpr, 'sz'))) { const v = int(wa(e, 'val')); if (v) p.size = v / 2; }
    if ((e = kid(rpr, 'highlight'))) { const v = wa(e, 'val'); p.hl = (v && v !== 'none' && HL[v]) || null; }
    if ((e = kid(rpr, 'shd'))) { const f = wa(e, 'fill'); if (f && f !== 'auto') p.shd = f.toUpperCase(); }
    if ((e = kid(rpr, 'vertAlign'))) { const v = wa(e, 'val'); p.sup = v === 'superscript'; p.sub = v === 'subscript'; }
    if ((e = kid(rpr, 'rFonts'))) {
      p.font = wa(e, 'ascii') || wa(e, 'hAnsi');
      if (!p.font) {
        const th = wa(e, 'asciiTheme') || wa(e, 'hAnsiTheme');
        if (th && theme) p.font = /^major/.test(th) ? theme.major : theme.minor;
      }
    }
    if ((e = kid(rpr, 'rStyle'))) p.rstyle = wa(e, 'val');
    return p;
  }
  function parsePPr(ppr) {
    const p = {};
    if (!ppr) return p;
    let e;
    if ((e = kid(ppr, 'pStyle'))) p.pstyle = wa(e, 'val');
    if ((e = kid(ppr, 'jc'))) { const v = wa(e, 'val'); p.jc = { left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' }[v]; }
    if ((e = kid(ppr, 'ind'))) {
      p.left = int(wa(e, 'left') != null ? wa(e, 'left') : wa(e, 'start'));
      p.right = int(wa(e, 'right') != null ? wa(e, 'right') : wa(e, 'end'));
      p.first = int(wa(e, 'firstLine'));
      p.hanging = int(wa(e, 'hanging'));
    }
    if ((e = kid(ppr, 'spacing'))) {
      p.before = int(wa(e, 'before'));
      p.after = int(wa(e, 'after'));
      p.line = int(wa(e, 'line'));
      p.lineRule = wa(e, 'lineRule') || undefined;
    }
    if ((e = kid(ppr, 'numPr'))) {
      const n = kid(e, 'numId'), l = kid(e, 'ilvl');
      if (n) p.numId = int(wa(n, 'val'));
      if (l) p.ilvl = int(wa(l, 'val'));
    }
    if ((e = kid(ppr, 'shd'))) { const f = wa(e, 'fill'); if (f && f !== 'auto') p.shd = f.toUpperCase(); }
    if ((e = kid(ppr, 'pBdr'))) { const b = kid(e, 'bottom'); if (b && wa(b, 'val') !== 'nil' && wa(b, 'val') !== 'none') p.borderBottom = true; }
    if ((e = kid(ppr, 'pageBreakBefore'))) p.pageBreakBefore = on(e);
    return p;
  }

  const BASE = {
    p: { size: 11, font: 'Calibri', before: 0, after: 160 },
    h1: { size: 16, font: 'Calibri Light', color: '2F5496', before: 240, after: 0 },
    h2: { size: 13, font: 'Calibri Light', color: '2F5496', before: 40, after: 0 },
    h3: { size: 12, font: 'Calibri Light', color: '1F3763', before: 40, after: 0 },
    h4: { size: 11, font: 'Calibri Light', color: '2F5496', i: true, before: 40, after: 0 },
    h5: { size: 11, font: 'Calibri Light', color: '2F5496', before: 40, after: 0 },
    h6: { size: 11, font: 'Calibri Light', color: '1F3763', before: 40, after: 0 },
    title: { size: 28, font: 'Calibri Light', before: 0, after: 0 },
    subtitle: { size: 11, font: 'Calibri', color: '595959', before: 0, after: 160 },
    quote: { size: 11, font: 'Calibri', color: '404040', i: true, before: 200, after: 160 },
  };

  async function read(arrayBuffer) {
    const z = await Zip.read(arrayBuffer);
    let mainPath = 'word/document.xml';
    if (z.has('_rels/.rels')) {
      const r = parseRels(await z.text('_rels/.rels'), '');
      for (const k in r) if (/officeDocument$/.test(r[k].type) && r[k].path) mainPath = r[k].path;
    }
    if (!z.has(mainPath)) throw new Error('This file does not contain a Word document (word/document.xml not found).');
    const base = dirname(mainPath);
    const relsPath = base + (base ? '/' : '') + '_rels/' + mainPath.split('/').pop() + '.rels';
    const rels = parseRels(await z.textOrNull(relsPath), base);
    const relOf = (suffix) => { for (const k in rels) if (rels[k].type.endsWith('/' + suffix)) return rels[k]; return null; };

    // theme
    const theme = { major: 'Calibri Light', minor: 'Calibri' };
    const th = relOf('theme');
    if (th && z.has(th.path)) {
      const t = parseXml(await z.text(th.path));
      const mj = t.getElementsByTagNameNS('*', 'majorFont')[0], mn = t.getElementsByTagNameNS('*', 'minorFont')[0];
      const lat = (f) => { const l = f && f.getElementsByTagNameNS('*', 'latin')[0]; return l && l.getAttribute('typeface'); };
      theme.major = lat(mj) || theme.major;
      theme.minor = lat(mn) || theme.minor;
    }
    // styles
    const styles = {};
    const defaults = { rPr: {}, pPr: {} };
    let defaultPara = null;
    const sr = relOf('styles');
    if (sr && z.has(sr.path)) {
      const root = parseXml(await z.text(sr.path));
      const dd = kid(root, 'docDefaults');
      if (dd) {
        defaults.rPr = parseRPr(kid(kid(dd, 'rPrDefault'), 'rPr'), theme);
        defaults.pPr = parsePPr(kid(kid(dd, 'pPrDefault'), 'pPr'));
      }
      for (const st of kids(root, 'style')) {
        const id = wa(st, 'styleId');
        const nm = kid(st, 'name'), bo = kid(st, 'basedOn');
        styles[id] = { id, type: wa(st, 'type'), name: nm ? (wa(nm, 'val') || '') : '', basedOn: bo ? wa(bo, 'val') : null, pPr: parsePPr(kid(st, 'pPr')), rPr: parseRPr(kid(st, 'rPr'), theme), tblPr: kid(st, 'tblPr') };
        if (wa(st, 'default') === '1' && wa(st, 'type') === 'paragraph') defaultPara = id;
      }
    }
    const chainOf = (id) => {
      const list = [];
      const seen = new Set();
      while (id && styles[id] && !seen.has(id)) { seen.add(id); list.unshift(styles[id]); id = styles[id].basedOn; }
      return list;
    };
    // numbering
    const numbering = { abs: {}, nums: {} };
    const nr = relOf('numbering');
    if (nr && z.has(nr.path)) {
      const root = parseXml(await z.text(nr.path));
      for (const a of kids(root, 'abstractNum')) {
        const lv = {};
        for (const l of kids(a, 'lvl')) {
          const f = kid(l, 'numFmt'), s = kid(l, 'start');
          lv[int(wa(l, 'ilvl'))] = { fmt: f ? wa(f, 'val') : 'decimal', start: s ? int(wa(s, 'val')) : 1 };
        }
        numbering.abs[wa(a, 'abstractNumId')] = lv;
      }
      for (const n of kids(root, 'num')) {
        const an = kid(n, 'abstractNumId');
        const ov = {};
        for (const o of kids(n, 'lvlOverride')) { const so = kid(o, 'startOverride'); if (so) ov[int(wa(o, 'ilvl'))] = int(wa(so, 'val')); }
        numbering.nums[wa(n, 'numId')] = { abs: an ? wa(an, 'val') : null, ov };
      }
    }
    // images
    const images = {};
    for (const id in rels) {
      const r = rels[id];
      if (!r.external && /\/image$/.test(r.type) && r.path && z.has(r.path)) {
        const ext = r.path.split('.').pop().toLowerCase();
        const mime = MIME_BY_EXT[ext];
        if (mime) images[id] = bytesToDataUrl(await z.bytes(r.path), mime);
      }
    }
    const footnotes = {};
    const fr = relOf('footnotes');
    if (fr && z.has(fr.path)) {
      try {
        const froot = parseXml(await z.text(fr.path));
        for (const fnEl of kids(froot, 'footnote')) {
          const id = wa(fnEl, 'id');
          if (wa(fnEl, 'type') && wa(fnEl, 'type') !== 'normal') continue;
          footnotes[id] = [...fnEl.getElementsByTagNameNS(NS_W, 't')].map(t => t.textContent).join('').trim();
        }
      } catch (e) { /* ignore */ }
    }
    const endnotes = {};
    const er0 = relOf('endnotes');
    if (er0 && z.has(er0.path)) {
      try {
        const eroot = parseXml(await z.text(er0.path));
        for (const enEl of kids(eroot, 'endnote')) {
          if (wa(enEl, 'type') && wa(enEl, 'type') !== 'normal') continue;
          endnotes[wa(enEl, 'id')] = [...enEl.getElementsByTagNameNS(NS_W, 't')].map(t => t.textContent).join('').trim();
        }
      } catch (e) { /* ignore */ }
    }
    let track = false;
    const setr = relOf('settings');
    if (setr && z.has(setr.path)) { try { track = /<w:trackRevisions\b(?![^>]*w:val="(0|false)")/.test(await z.text(setr.path)); } catch (e) { /* ignore */ } }
    const comments = {};
    const cr = relOf('comments');
    if (cr && z.has(cr.path)) {
      try {
        const croot = parseXml(await z.text(cr.path));
        for (const cEl of kids(croot, 'comment')) comments[wa(cEl, 'id')] = { author: wa(cEl, 'author') || 'Author', text: [...cEl.getElementsByTagNameNS(NS_W, 't')].map(t => t.textContent).join('') };
      } catch (e) { /* ignore */ }
    }
    const ctx = { styles, chainOf, defaults, numbering, theme, images, rels, defaultPara, footnotes, comments, endnotes };

    const docRoot = parseXml(await z.text(mainPath));
    const body = kid(docRoot, 'body');
    const root = document.createElement('div');
    convertBlocks(flatten(body), root, ctx);
    if (!root.firstChild) root.innerHTML = '<p><br></p>';

    const page = { w: 12240, h: 15840, t: 1440, r: 1440, b: 1440, l: 1440 };
    const sect = kid(body, 'sectPr');
    if (sect) {
      const sz = kid(sect, 'pgSz'), mg = kid(sect, 'pgMar'), cl = kid(sect, 'cols');
      if (cl && int(wa(cl, 'num')) > 1) page.cols = int(wa(cl, 'num'));
      if (sz) { page.w = int(wa(sz, 'w')) || page.w; page.h = int(wa(sz, 'h')) || page.h; }
      if (mg) { page.t = int(wa(mg, 'top')) ?? page.t; page.r = int(wa(mg, 'right')) ?? page.r; page.b = int(wa(mg, 'bottom')) ?? page.b; page.l = int(wa(mg, 'left')) ?? page.l; }
    }
    const hf = {};
    if (sect) {
      for (const [tag, key] of [['headerReference', 'header'], ['footerReference', 'footer']]) {
        const refs = kids(sect, tag);
        const ref = refs.find(r => wa(r, 'type') === 'default') || refs[0];
        if (!ref) continue;
        const rel = rels[ref.getAttributeNS(NS_R, 'id') || ref.getAttribute('r:id')];
        if (!rel || !z.has(rel.path)) continue;
        try {
          const hroot = parseXml(await z.text(rel.path));
          const box = document.createElement('div');
          ctx.fs = null;
          convertBlocks(flatten(hroot), box, ctx);
          hf[key] = box.textContent.trim() || box.querySelector('.fld') ? box.innerHTML : '';
        } catch (e) { /* ignore header errors */ }
      }
    }
    // bibliography sources (customXml part with the b:Sources root)
    try {
      for (const nm of z.names()) {
        if (!/^customXml\/item\d+\.xml$/i.test(nm)) continue;
        const txt = await z.text(nm);
        if (!/officeDocument\/2006\/bibliography/.test(txt)) continue;
        const sx = parseXml(txt);
        const T = (el, n) => { const e = el.getElementsByTagNameNS('*', n)[0]; return e ? e.textContent.trim() : ''; };
        const TYPES = { Book: 'book', JournalArticle: 'article', InternetSite: 'web', DocumentFromInternetSite: 'web', Report: 'report' };
        const list = [];
        for (const s of sx.getElementsByTagNameNS('*', 'Source')) {
          const people = [...s.getElementsByTagNameNS('*', 'Person')].filter(p => p.parentNode && p.parentNode.parentNode && p.parentNode.parentNode.localName === 'Author').map(p => T(p, 'Last') + (T(p, 'First') ? ', ' + T(p, 'First') : ''));
          const corp = T(s, 'Corporate');
          list.push({ tag: T(s, 'Tag'), type: TYPES[T(s, 'SourceType')] || 'book', author: people.length ? people.join('; ') : corp, title: T(s, 'Title'), year: T(s, 'Year'), publisher: T(s, 'Publisher'), city: T(s, 'City'), journal: T(s, 'JournalName'), url: T(s, 'URL') });
        }
        if (list.length) page.sources = list;
        break;
      }
    } catch (e) { console.warn('bibliography ignored', e); }
    // styles that differ from the built-in look become editable style overrides
    {
      const ov = {};
      for (const key of Object.keys(STYLE_IDS)) {
        const ch = chainOf(STYLE_IDS[key]);
        if (!ch.length) continue;
        const eff = Object.assign({}, defaults.rPr);
        if (key !== 'normal' && defaultPara && styles[defaultPara] && !ch.includes(styles[defaultPara])) Object.assign(eff, styles[defaultPara].rPr);
        for (const st of ch) Object.assign(eff, st.rPr);
        const d = STYLE_DEFAULTS[key];
        const o = { ff: eff.font || d.ff, sz: eff.size || d.sz, color: eff.color && eff.color !== 'auto' ? eff.color : (eff.color === 'auto' ? '000000' : d.color), b: !!eff.b, i: !!eff.i };
        if (o.ff !== d.ff || o.sz !== d.sz || o.color !== d.color || o.b !== d.b || o.i !== d.i) ov[key] = o;
      }
      if (Object.keys(ov).length) page.styles = ov;
    }
    return { root, page, header: hf.header || '', footer: hf.footer || '', track };
  }

  function flatten(parent) {
    const out = [];
    for (const c of parent.children) {
      const ln = c.localName;
      if (ln === 'p' || ln === 'tbl') out.push(c);
      else if (ln === 'sdt') { const sc = kid(c, 'sdtContent'); if (sc) out.push(...flatten(sc)); }
      else if (ln === 'customXml' || ln === 'ins' || ln === 'moveTo' || ln === 'smartTag') out.push(...flatten(c));
    }
    return out;
  }

  function convertBlocks(els, container, ctx) {
    let stack = [];
    for (const el of els) {
      if (el.localName === 'tbl') {
        stack = [];
        container.appendChild(convertTable(el, ctx));
        continue;
      }
      const r = convertPara(el, ctx);
      if (r.list) {
        placeListItem(r.nodes[0], r.list, container, stack);
      } else {
        stack = [];
        for (const n of r.nodes) container.appendChild(n);
      }
    }
  }

  function placeListItem(li, info, container, stack) {
    const L = info.ilvl;
    while (stack.length && (stack[stack.length - 1].lvl > L || (stack[stack.length - 1].lvl === L && stack[stack.length - 1].numId !== info.numId))) stack.pop();
    let top = stack[stack.length - 1];
    if (!top || top.lvl < L) {
      const bullet = info.fmt === 'bullet';
      const listEl = document.createElement(bullet ? 'ul' : 'ol');
      if (!bullet) {
        const type = { lowerLetter: 'lower-alpha', upperLetter: 'upper-alpha', lowerRoman: 'lower-roman', upperRoman: 'upper-roman' }[info.fmt];
        if (type) listEl.style.listStyleType = type;
        if (info.start && info.start !== 1) listEl.setAttribute('start', info.start);
      }
      if (top) (top.el.lastElementChild || top.el).appendChild(listEl);
      else container.appendChild(listEl);
      top = { el: listEl, lvl: L, numId: info.numId };
      stack.push(top);
    }
    top.el.appendChild(li);
  }

  function mergeProps(a, b) { return Object.assign(a, defined(b)); }

  function runCss(e, base) {
    let css = '';
    const font = e.font || base.font;
    if (font && font.toLowerCase() !== (base.font || '').toLowerCase()) css += `font-family:'${font}';`;
    if (e.size && e.size !== base.size) css += `font-size:${e.size}pt;`;
    const col = (e.color && e.color !== 'auto') ? e.color : '000000';
    if (col !== (base.color || '000000')) css += `color:#${col};`;
    if (!!e.b !== !!base.b) css += `font-weight:${e.b ? 'bold' : 'normal'};`;
    if (!!e.i !== !!base.i) css += `font-style:${e.i ? 'italic' : 'normal'};`;
    let td = '';
    if (e.u) td += ' underline';
    if (e.s) td += ' line-through';
    if (td) css += `text-decoration:${td.trim()};`;
    const bg = e.hl || e.shd;
    if (bg) css += `background-color:#${bg};`;
    if (e.sup) css += 'vertical-align:super;font-size:' + ((e.size || base.size || 11) * 0.65) + 'pt;';
    if (e.sub) css += 'vertical-align:sub;font-size:' + ((e.size || base.size || 11) * 0.65) + 'pt;';
    if (e.caps) css += 'text-transform:uppercase;';
    if (e.smallCaps) css += 'font-variant:small-caps;';
    return css;
  }

  function kindOf(sid, ctx) {
    const chain = ctx.chainOf(sid);
    for (let i = chain.length - 1; i >= 0; i--) {
      const nm = (chain[i].name || '').toLowerCase();
      let m = /^heading (\d)$/.exec(nm);
      if (m) return 'h' + Math.min(6, +m[1]);
      if (nm === 'title') return 'title';
      if (nm === 'subtitle') return 'subtitle';
      if (nm === 'quote' || nm === 'intense quote') return 'quote';
      m = /^toc (\d)$/.exec(nm);
      if (m) return 'toc' + Math.min(3, +m[1]);
    }
    return 'p';
  }

  function convertPara(p, ctx) {
    ctx.fs = null;
    const ppr = parsePPr(kid(p, 'pPr'));
    const sid = ppr.pstyle || ctx.defaultPara;
    const chain = ctx.chainOf(sid);
    const ep = Object.assign({}, ctx.defaults.pPr);
    const er = Object.assign({}, ctx.defaults.rPr);
    for (const s of chain) { mergeProps(ep, s.pPr); mergeProps(er, s.rPr); }
    mergeProps(ep, ppr);
    const kind = kindOf(sid, ctx);
    const base = BASE[kind] || BASE.p;

    const nodes = [];
    const hasInline = (container) => !!container.firstChild;
    let cur = null;
    const startPara = () => {
      const tag = /^h\d$/.test(kind) ? kind : 'p';
      cur = document.createElement(tag);
      if (kind === 'title' || kind === 'subtitle' || kind === 'quote') cur.className = kind;
      else if (/^toc\d$/.test(kind)) cur.className = 'toc ' + kind;
    };
    startPara();

    const out = []; // inline nodes or 'PB'
    collectInline(p, out, ctx, er, base, false);
    let pbBefore = ep.pageBreakBefore;
    let openCmt = null;
    for (const n of out) {
      if (n === 'PB') {
        openCmt = null;
        if (hasInline(cur)) nodes.push(cur);
        const d = document.createElement('div');
        d.className = 'pb';
        d.setAttribute('contenteditable', 'false');
        nodes.push(d);
        startPara();
      } else if (n && n.tbox) {
        openCmt = null;
        if (hasInline(cur)) nodes.push(cur);
        nodes.push(n.tbox);
        startPara();
      } else if (n && n.cs !== undefined) {
        const c0 = ctx.comments && ctx.comments[n.cs];
        if (c0) { openCmt = document.createElement('span'); openCmt.className = 'cmt'; openCmt.dataset.t = c0.text; openCmt.dataset.a = c0.author; openCmt.title = c0.author + ': ' + c0.text; cur.appendChild(openCmt); }
      } else if (n && n.ce !== undefined) openCmt = null;
      else if (n instanceof Node) (openCmt || cur).appendChild(n);
    }
    if (/^toc\d$/.test(kind) && cur.textContent.includes('\t')) {
      const tx = cur.textContent, ix = tx.lastIndexOf('\t');
      cur.textContent = '';
      cur.appendChild(document.createTextNode(tx.slice(0, ix).replace(/\t/g, ' ')));
      const dots = document.createElement('span'); dots.className = 'tocdots'; dots.setAttribute('contenteditable', 'false'); cur.appendChild(dots);
      const pg = document.createElement('span'); pg.className = 'tocpg'; pg.textContent = tx.slice(ix + 1); cur.appendChild(pg);
    }
    let tail = true;
    if (!hasInline(cur)) {
      tail = nodes.length === 0;
      if (tail) cur.appendChild(document.createElement('br'));
    }

    // list?
    let numId = ep.numId, ilvl = ep.ilvl || 0;
    if (!/^h\d$/.test(kind) && numId && numId !== 0 && tail && nodes.length === 0) {
      const num = ctx.numbering.nums[numId];
      const lv = num && ctx.numbering.abs[num.abs] && ctx.numbering.abs[num.abs][ilvl];
      const li = document.createElement('li');
      while (cur.firstChild) li.appendChild(cur.firstChild);
      if (!li.firstChild || (li.childNodes.length === 1 && li.firstChild.nodeName === 'BR')) { /* keep empty item */ }
      let css = '';
      if (ep.jc && ep.jc !== 'left') css += 'text-align:' + ep.jc + ';';
      if (css) li.setAttribute('style', css);
      const start = (num && num.ov && num.ov[ilvl]) || (lv && lv.start) || 1;
      return { nodes: [li], list: { numId, ilvl, fmt: lv ? lv.fmt : 'bullet', start } };
    }

    // paragraph level CSS
    const lastPara = tail ? cur : null;
    if (lastPara || nodes.length) {
      const targets = nodes.filter(n => n.nodeName !== 'DIV');
      if (tail) targets.push(cur);
      let css = '';
      if (ep.jc && ep.jc !== 'left') css += `text-align:${ep.jc};`;
      if (ep.left) css += `margin-left:${Math.round(ep.left / 15)}px;`;
      if (ep.right) css += `margin-right:${Math.round(ep.right / 15)}px;`;
      const fl = ep.hanging ? -ep.hanging : ep.first;
      if (fl) css += `text-indent:${Math.round(fl / 15)}px;`;
      if ((ep.before || 0) !== base.before) css += `margin-top:${(ep.before || 0) / 20}pt;`;
      if ((ep.after || 0) !== base.after) css += `margin-bottom:${(ep.after || 0) / 20}pt;`;
      if (ep.line && (!ep.lineRule || ep.lineRule === 'auto')) {
        const ratio = ep.line / 240;
        if (Math.abs(ratio - 259 / 240) > 0.01) css += `line-height:${+ratio.toFixed(3)};`;
      } else if (ep.line && ep.lineRule === 'exact') css += `line-height:${ep.line / 20}pt;`;
      else if (!ep.line) css += 'line-height:1;';
      if (ep.shd) css += `background-color:#${ep.shd};`;
      for (const t of targets) if (css) t.setAttribute('style', css);
    }
    // horizontal rule
    if (ep.borderBottom && nodes.length === 0 && cur.textContent === '') {
      return { nodes: [document.createElement('hr')], list: null };
    }
    if (tail) nodes.push(cur);
    if (pbBefore) {
      const d = document.createElement('div');
      d.className = 'pb';
      d.setAttribute('contenteditable', 'false');
      nodes.unshift(d);
    }
    return { nodes, list: null };
  }

  function fieldSpan(kind) {
    const s = document.createElement('span');
    s.className = 'fld'; s.dataset.f = kind; s.setAttribute('contenteditable', 'false');
    s.textContent = '1';
    return s;
  }
  function citeSpan(tag, text) {
    const s = document.createElement('span');
    s.className = 'cite'; s.dataset.tag = tag; s.setAttribute('contenteditable', 'false');
    s.textContent = (text || '').trim() || '(' + tag + ')';
    return s;
  }
  function mergeSpan(name) {
    const s = document.createElement('span');
    s.className = 'mf'; s.dataset.n = name; s.setAttribute('contenteditable', 'false');
    s.textContent = '«' + name + '»';
    return s;
  }
  function textboxDiv(holder, ctx, wPx) {
    const d = document.createElement('div');
    d.className = 'tbox';
    if (wPx) d.style.width = Math.round(wPx) + 'px';
    const keep = ctx.fs;
    convertBlocks(flatten(holder), d, ctx);
    ctx.fs = keep;
    if (!d.firstChild) d.innerHTML = '<p><br></p>';
    return d;
  }
  function collectInline(parent, out, ctx, er, base, inLink) {
    for (const ch of parent.children) {
      switch (ch.localName) {
        case 'r': convertRun(ch, out, ctx, er, base); break;
        case 'commentRangeStart': out.push({ cs: wa(ch, 'id') }); break;
        case 'commentRangeEnd': out.push({ ce: wa(ch, 'id') }); break;
        case 'hyperlink': {
          const rid = ch.getAttributeNS(NS_R, 'id') || ch.getAttribute('r:id');
          const rel = rid && ctx.rels[rid];
          const inner = [];
          collectInline(ch, inner, ctx, er, base, true);
          const anc = wa(ch, 'anchor');
          if (anc && !rid) {
            const a = document.createElement('a');
            a.setAttribute('href', '#' + anc);
            for (const n of inner) if (n instanceof Node) a.appendChild(n);
            if (a.firstChild) out.push(a);
          } else if (rel && rel.external && rel.target) {
            const a = document.createElement('a');
            a.setAttribute('href', rel.target);
            for (const n of inner) if (n !== 'PB' && n instanceof Node) a.appendChild(n);
            if (a.firstChild) out.push(a);
          } else out.push(...inner);
          break;
        }
        case 'fldSimple': {
          const ins = wa(ch, 'instr') || '';
          const mfm = /\bMERGEFIELD\s+"?([^\s"\\]+)/.exec(ins);
          const cim = /\bCITATION\s+"?([^\s"\\]+)/.exec(ins);
          if (cim) out.push(citeSpan(cim[1], ch.textContent));
          else if (mfm) out.push(mergeSpan(mfm[1]));
          else if (/\b(NUMPAGES|PAGE)\b/.test(ins)) out.push(fieldSpan(/NUMPAGES/.test(ins) ? 'NUMPAGES' : 'PAGE'));
          else collectInline(ch, out, ctx, er, base, inLink);
          break;
        }
        case 'ins': case 'moveTo': case 'del': case 'moveFrom': {
          const inner = [];
          collectInline(ch, inner, ctx, er, base, inLink);
          const isIns = ch.localName === 'ins' || ch.localName === 'moveTo';
          let w = null;
          for (const n of inner) {
            if (!(n instanceof Node)) { w = null; out.push(n); continue; }
            if (!w) {
              w = document.createElement(isIns ? 'ins' : 'del');
              w.className = 'tc';
              w.dataset.a = wa(ch, 'author') || 'Author';
              if (wa(ch, 'date')) w.dataset.d = wa(ch, 'date');
              w.title = (isIns ? 'Inserted by ' : 'Deleted by ') + w.dataset.a;
              out.push(w);
            }
            w.appendChild(n);
          }
          break;
        }
        case 'bookmarkStart': {
          const nm = wa(ch, 'name');
          if (nm && (nm[0] !== '_' || /^_Ref/.test(nm))) { const s = document.createElement('span'); s.className = 'bm'; s.dataset.n = nm; s.setAttribute('contenteditable', 'false'); s.title = 'Bookmark: ' + nm; out.push(s); }
          break;
        }
        case 'oMath': out.push(Eq.span(Eq.fromOmml(ch))); break;
        case 'oMathPara': for (const m of ch.children) if (m.localName === 'oMath') out.push(Eq.span(Eq.fromOmml(m))); break;
        case 'smartTag': case 'customXml':
          collectInline(ch, out, ctx, er, base, inLink); break;
        case 'sdt': { const sc = kid(ch, 'sdtContent'); if (sc) collectInline(sc, out, ctx, er, base, inLink); break; }
        default: break;
      }
    }
  }

  function convertRun(r, out, ctx, er, base) {
    const rp = parseRPr(kid(r, 'rPr'), ctx.theme);
    const eff = Object.assign({}, er);
    if (rp.rstyle) for (const s of ctx.chainOf(rp.rstyle)) mergeProps(eff, s.rPr);
    mergeProps(eff, rp);
    if (eff.vanish) return;
    const css = runCss(eff, base);
    let buf = '';
    const flush = () => {
      if (!buf) return;
      if (css) { const s = document.createElement('span'); s.setAttribute('style', css); s.textContent = buf; out.push(s); }
      else out.push(document.createTextNode(buf));
      buf = '';
    };
    for (const c of r.children) {
      switch (c.localName) {
        case 't': if (!(ctx.fs && ctx.fs.skip)) buf += c.textContent; break;
        case 'instrText': if (ctx.fs && ctx.fs.inInstr) ctx.fs.instr += c.textContent; break;
        case 'fldChar': {
          const ty = wa(c, 'fldCharType');
          const fs = ctx.fs || (ctx.fs = { instr: '', inInstr: false, skip: false });
          if (ty === 'begin') { fs.instr = ''; fs.inInstr = true; fs.skip = false; }
          else if (ty === 'separate') {
            fs.inInstr = false;
            const mfm = /\bMERGEFIELD\s+"?([^\s"\\]+)/.exec(fs.instr);
            const cim = /\bCITATION\s+"?([^\s"\\]+)/.exec(fs.instr);
            if (cim) { flush(); out.push(citeSpan(cim[1], '')); fs.skip = true; }
            else if (mfm) { flush(); out.push(mergeSpan(mfm[1])); fs.skip = true; }
            else if (/\b(NUMPAGES|PAGE)\b/.test(fs.instr)) { flush(); out.push(fieldSpan(/NUMPAGES/.test(fs.instr) ? 'NUMPAGES' : 'PAGE')); fs.skip = true; }
          } else if (ty === 'end') { fs.skip = false; fs.inInstr = false; }
          break;
        }
        case 'tab': buf += '\t'; break;
        case 'noBreakHyphen': buf += '‑'; break;
        case 'sym': { const ch = wa(c, 'char'); if (ch) { let n = parseInt(ch, 16); if (n >= 0xF000) n -= 0xF000; buf += String.fromCharCode(n); } break; }
        case 'br': {
          const t = wa(c, 'type');
          flush();
          if (t === 'page') out.push('PB');
          else out.push(document.createElement('br'));
          break;
        }
        case 'cr': flush(); out.push(document.createElement('br')); break;
        case 'footnoteReference': {
          flush();
          const sp = document.createElement('sup');
          sp.className = 'fn'; sp.setAttribute('contenteditable', 'false');
          sp.dataset.t = (ctx.footnotes && ctx.footnotes[wa(c, 'id')]) || '';
          sp.textContent = '1';
          out.push(sp);
          break;
        }
        case 'delText': buf += c.textContent; break;
        case 'endnoteReference': {
          flush();
          const sp = document.createElement('sup');
          sp.className = 'en'; sp.setAttribute('contenteditable', 'false');
          sp.dataset.t = (ctx.endnotes && ctx.endnotes[wa(c, 'id')]) || '';
          sp.textContent = 'i';
          out.push(sp);
          break;
        }
        case 'AlternateContent': {
          flush();
          let dr = null;
          for (const alt of c.children) { dr = alt.getElementsByTagNameNS('*', 'drawing')[0] || alt.getElementsByTagNameNS('*', 'pict')[0]; if (dr) break; }
          if (dr) pushDrawing(dr, out, ctx);
          break;
        }
        case 'drawing': case 'pict': flush(); pushDrawing(c, out, ctx); break;
        default: break;
      }
    }
    flush();
  }

  function pushDrawing(c, out, ctx) {
    const tb = c.getElementsByTagNameNS(NS_W, 'txbxContent')[0];
    if (tb) {
      let w = 0;
      const ext = c.getElementsByTagNameNS('*', 'extent')[0];
      if (ext) w = parseInt(ext.getAttribute('cx'), 10) / 9525;
      else { const sh = c.getElementsByTagNameNS('*', 'shape')[0]; const m = sh && /width:([\d.]+)pt/.exec(sh.getAttribute('style') || ''); if (m) w = m[1] * 4 / 3; }
      out.push({ tbox: textboxDiv(tb, ctx, w) });
      return;
    }
    const im = c.localName === 'pict' ? convertPict(c, ctx) : convertDrawing(c, ctx);
    if (im) out.push(im);
  }
  function convertDrawing(dr, ctx) {
    const blip = dr.getElementsByTagNameNS('*', 'blip')[0];
    if (!blip) return null;
    const rid = blip.getAttributeNS(NS_R, 'embed') || blip.getAttribute('r:embed');
    const src = ctx.images[rid];
    if (!src) return null;
    const ext = dr.getElementsByTagNameNS('*', 'extent')[0];
    const img = document.createElement('img');
    img.setAttribute('src', src);
    if (ext) {
      const cx = parseInt(ext.getAttribute('cx'), 10), cy = parseInt(ext.getAttribute('cy'), 10);
      if (cx > 0 && cy > 0) img.setAttribute('style', `width:${Math.round(cx / 9525)}px;height:${Math.round(cy / 9525)}px;`);
    }
    const dp = dr.getElementsByTagNameNS('*', 'docPr')[0];
    if (dp && dp.getAttribute('descr')) img.setAttribute('alt', dp.getAttribute('descr'));
    return img;
  }
  function convertPict(pict, ctx) {
    const idata = pict.getElementsByTagNameNS('*', 'imagedata')[0];
    if (!idata) return null;
    const rid = idata.getAttributeNS(NS_R, 'id') || idata.getAttribute('r:id');
    const src = ctx.images[rid];
    if (!src) return null;
    const img = document.createElement('img');
    img.setAttribute('src', src);
    const shape = pict.getElementsByTagNameNS('*', 'shape')[0];
    const st = shape && shape.getAttribute('style');
    if (st) {
      const w = /width:([\d.]+)pt/.exec(st), h = /height:([\d.]+)pt/.exec(st);
      if (w && h) img.setAttribute('style', `width:${Math.round(w[1] * 4 / 3)}px;height:${Math.round(h[1] * 4 / 3)}px;`);
    }
    return img;
  }

  function tblHasBorders(tblPr) {
    const b = kid(tblPr, 'tblBorders');
    if (!b) return false;
    for (const s of b.children) { const v = wa(s, 'val'); if (v && v !== 'nil' && v !== 'none') return true; }
    return false;
  }

  function convertTable(tbl, ctx) {
    const tblPr = kid(tbl, 'tblPr');
    let bordered = tblHasBorders(tblPr);
    const ts = tblPr && kid(tblPr, 'tblStyle');
    if (!bordered && ts) for (const s of ctx.chainOf(wa(ts, 'val'))) if (s.tblPr && tblHasBorders(s.tblPr)) bordered = true;
    const gridEl = kid(tbl, 'tblGrid');
    const grid = gridEl ? kids(gridEl, 'gridCol').map(g => int(wa(g, 'w')) || 0) : [];
    const table = document.createElement('table');
    if (!bordered) table.className = 'nb';
    const rows = kids(tbl, 'tr');
    // detect borders on cells too
    const spanTargets = []; // col -> cell element awaiting vMerge continuation
    const tbody = document.createElement('tbody');
    let maxCols = grid.length;
    for (const tr of rows) {
      const trEl = document.createElement('tr');
      let col = 0;
      for (const tc of kids(tr, 'tc')) {
        const pr = kid(tc, 'tcPr');
        const gs = pr && kid(pr, 'gridSpan');
        const span = gs ? (int(wa(gs, 'val')) || 1) : 1;
        const vm = pr && kid(pr, 'vMerge');
        if (vm && wa(vm, 'val') !== 'restart') {
          const top = spanTargets[col];
          if (top) top.rowSpan = (top.rowSpan || 1) + 1;
          col += span;
          continue;
        }
        const td = document.createElement('td');
        if (span > 1) td.colSpan = span;
        if (vm) spanTargets[col] = td; else spanTargets[col] = null;
        let css = '';
        const shd = pr && kid(pr, 'shd');
        if (shd) { const f = wa(shd, 'fill'); if (f && f !== 'auto') css += `background-color:#${f};`; }
        const va = pr && kid(pr, 'vAlign');
        if (va && wa(va, 'val') === 'center') css += 'vertical-align:middle;';
        else if (va && wa(va, 'val') === 'bottom') css += 'vertical-align:bottom;';
        let w = 0;
        for (let k = col; k < col + span; k++) w += grid[k] || 0;
        if (!w && pr) { const tw = kid(pr, 'tcW'); if (tw && wa(tw, 'type') === 'dxa') w = int(wa(tw, 'w')) || 0; }
        if (w) css += `width:${Math.round(w / 15)}px;`;
        if (css) td.setAttribute('style', css);
        convertBlocks(flatten(tc), td, ctx);
        if (!td.firstChild) td.innerHTML = '<p><br></p>';
        trEl.appendChild(td);
        col += span;
      }
      maxCols = Math.max(maxCols, col);
      tbody.appendChild(trEl);
    }
    const total = grid.reduce((a, b) => a + b, 0);
    if (total) {
      table.style.width = Math.round(total / 15) + 'px';
      const cg = document.createElement('colgroup');
      for (const g of grid) { const c = document.createElement('col'); c.style.width = Math.round(g / 15) + 'px'; cg.appendChild(c); }
      table.appendChild(cg);
    }
    table.appendChild(tbody);
    return table;
  }

  global.DocxIO = { write, read, DOCX_MIME, Eq };
})(window);
