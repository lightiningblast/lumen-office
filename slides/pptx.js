/* PPTX (PresentationML) reader / writer for Slides.
   Model: pres = { w, h (px, 1px = 9525 EMU), media: {id: dataUrl}, slides: [ { bg, layout, shapes: [...] } ] }
   shape  = { id, type, x, y, w, h, rot, flipH, flipV, fill, line:{color,w}|null, html, valign, ph, fs, fc, ff, img, crop, tail,
              rows:[[{html,fill}]], colW:[], rowH:[] }                                                                  */
(function (g) {
  'use strict';
  const { xesc, cssToHex, dataUrlToBytes, bytesToDataUrl, MIME_BY_EXT } = Lumen;
  const NS = {
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  };
  const REL = NS.r;
  const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
  const EMU = 9525;
  const NSDECL = `xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"`;
  const emu = (px) => Math.round(px * EMU);

  function lenPx(v) {
    if (!v) return null;
    const m = /^(-?[\d.]+)(px|pt|in|cm|mm|em)?$/.exec(String(v).trim());
    if (!m) return null;
    const n = parseFloat(m[1]);
    switch (m[2] || 'px') { case 'pt': return n * 4 / 3; case 'in': return n * 96; case 'cm': return n * 96 / 2.54; case 'mm': return n * 96 / 25.4; case 'em': return n * 16; default: return n; }
  }

  /* ====================================================================
     WRITE
     ==================================================================== */
  const LAYOUTS = [
    { key: 'title', name: 'Title Slide', type: 'title' },
    { key: 'titleContent', name: 'Title and Content', type: 'obj' },
    { key: 'titleOnly', name: 'Title Only', type: 'titleOnly' },
    { key: 'blank', name: 'Blank', type: 'blank' },
  ];

  function runCtxOf(el, c) {
    const n = Object.assign({}, c);
    const t = el.tagName;
    if (t === 'B' || t === 'STRONG') n.b = true;
    else if (t === 'I' || t === 'EM') n.i = true;
    else if (t === 'U') n.u = true;
    else if (t === 'S' || t === 'STRIKE' || t === 'DEL') n.s = true;
    else if (t === 'SUB') n.sub = true;
    else if (t === 'SUP') n.sup = true;
    else if (t === 'FONT') {
      const col = cssToHex(el.getAttribute('color')); if (col) n.color = col;
      const face = el.getAttribute('face'); if (face) n.font = face.split(',')[0].replace(/["']/g, '').trim();
      const sz = parseInt(el.getAttribute('size'), 10); if (sz) n.size = [0, 7.5, 10, 12, 13.5, 18, 24, 36][sz] || n.size;
    }
    const st = el.style;
    if (st && st.length) {
      const fw = st.fontWeight; if (fw) n.b = (fw === 'bold' || fw === 'bolder' || parseInt(fw, 10) >= 600);
      const fi = st.fontStyle; if (fi) n.i = (fi === 'italic' || fi === 'oblique');
      const td = st.textDecorationLine || st.textDecoration || '';
      if (td.includes('underline')) n.u = true;
      if (td.includes('line-through')) n.s = true;
      const col = cssToHex(st.color); if (col) n.color = col;
      if (st.fontFamily) n.font = st.fontFamily.split(',')[0].replace(/["']/g, '').trim();
      if (st.fontSize) { const px = lenPx(st.fontSize); if (px) n.size = Math.round(px * 0.75 * 2) / 2; }
      if (st.verticalAlign === 'super') n.sup = true; else if (st.verticalAlign === 'sub') n.sub = true;
    }
    return n;
  }
  function rPrXml(c, base, S, tag) {
    const size = c.size || base.fs || 18;
    let attrs = ` lang="en-US" sz="${Math.round(size * 100)}"`;
    if (c.b) attrs += ' b="1"'; if (c.i) attrs += ' i="1"'; if (c.u) attrs += ' u="sng"'; if (c.s) attrs += ' strike="sngStrike"';
    if (c.sup) attrs += ' baseline="30000"'; else if (c.sub) attrs += ' baseline="-25000"';
    attrs += ' dirty="0"';
    const col = c.color || base.fc || '000000';
    const font = xesc(c.font || base.ff || 'Calibri');
    let inner = `<a:solidFill><a:srgbClr val="${col}"/></a:solidFill><a:latin typeface="${font}"/><a:cs typeface="${font}"/>`;
    if (c.href && S) {
      const rid = 'rId' + (S.rid++);
      S.rels.push(`<Relationship Id="${rid}" Type="${REL}/hyperlink" Target="${xesc(c.href)}" TargetMode="External"/>`);
      inner += `<a:hlinkClick r:id="${rid}"/>`;
    }
    return `<a:${tag || 'rPr'}${attrs}>${inner}</a:${tag || 'rPr'}>`;
  }
  function runsOf(node, c, base, S, out) {
    if (node.nodeType === 3) {
      const t = node.nodeValue.replace(/[​﻿]/g, '').replace(/ /g, ' ').replace(/[\r\n]+/g, ' ');
      if (t) out.push(`<a:r>${rPrXml(c, base, S)}<a:t>${xesc(t)}</a:t></a:r>`);
      return;
    }
    if (node.nodeType !== 1) return;
    if (node.tagName === 'BR') { if (node.nextSibling) out.push(`<a:br>${rPrXml(c, base, S)}</a:br>`); return; }
    let c2 = runCtxOf(node, c);
    if (node.tagName === 'A' && /^(https?:|mailto:)/i.test(node.getAttribute('href') || '')) c2 = Object.assign(c2, { href: node.getAttribute('href'), u: true });
    for (const ch of node.childNodes) runsOf(ch, c2, base, S, out);
  }
  const ALGN = { left: 'l', start: 'l', center: 'ctr', right: 'r', end: 'r', justify: 'just' };
  function paraXml(el, base, S, forceBullet) {
    const st = el.style || {};
    let pattrs = '';
    const ml = lenPx(st.marginLeft), ti = lenPx(st.textIndent);
    const bu = forceBullet || (el.dataset && el.dataset.bu) || '';
    const lvl = (el.dataset && parseInt(el.dataset.l, 10)) || 0;
    if (ml) pattrs += ` marL="${emu(ml)}"`;
    if (ti) pattrs += ` indent="${emu(ti)}"`;
    else if (bu && ml) pattrs += ` indent="${-emu(Math.min(ml, 36))}"`;
    const al = ALGN[st.textAlign || el.getAttribute('align') || ''];
    if (al) pattrs += ` algn="${al}"`;
    if (lvl) pattrs += ` lvl="${Math.min(8, lvl)}"`;
    let pchildren = '';
    const lh = st.lineHeight;
    if (lh && /^[\d.]+$/.test(lh)) pchildren += `<a:lnSpc><a:spcPct val="${Math.round(parseFloat(lh) * 100000)}"/></a:lnSpc>`;
    const mt = lenPx(st.marginTop), mb = lenPx(st.marginBottom);
    if (mt) pchildren += `<a:spcBef><a:spcPts val="${Math.round(mt * 0.75 * 100)}"/></a:spcBef>`;
    if (mb) pchildren += `<a:spcAft><a:spcPts val="${Math.round(mb * 0.75 * 100)}"/></a:spcAft>`;
    if (bu === 'num') pchildren += '<a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/>';
    else if (bu) pchildren += `<a:buFont typeface="Arial"/><a:buChar char="${xesc((el.dataset && el.dataset.bc) || '•')}"/>`;
    else pchildren += '<a:buNone/>';
    const runs = [];
    const c0 = runCtxOf(el, {});
    for (const ch of el.childNodes) runsOf(ch, c0, base, S, runs);
    return `<a:p><a:pPr${pattrs}>${pchildren}</a:pPr>${runs.join('')}${rPrXml(c0, base, null, 'endParaRPr')}</a:p>`;
  }
  const BLK = new Set(['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'BLOCKQUOTE']);
  function textBodyParas(html, base, S) {
    const host = document.createElement('div');
    host.innerHTML = html || '';
    let out = '';
    let pending = null;
    const flush = () => { if (pending) { out += paraXml(pending, base, S); pending = null; } };
    const walk = (parent, inList) => {
      for (const n of parent.childNodes) {
        if (n.nodeType === 1 && (n.tagName === 'UL' || n.tagName === 'OL')) {
          flush();
          for (const li of n.children) { if (li.tagName === 'LI') out += paraXml(li, base, S, n.tagName === 'OL' ? 'num' : 'dot'); else if (li.tagName === 'UL' || li.tagName === 'OL') walk(li.parentNode === n ? { childNodes: [li] } : li, true); }
        } else if (n.nodeType === 1 && BLK.has(n.tagName)) {
          flush();
          const hasBlock = [...n.children].some(c => BLK.has(c.tagName) || c.tagName === 'UL' || c.tagName === 'OL');
          if (hasBlock) walk(n, inList); else out += paraXml(n, base, S);
        } else {
          if (!pending) { pending = document.createElement('p'); }
          pending.appendChild(n.cloneNode(true));
        }
      }
    };
    walk(host, false);
    flush();
    if (!out) out = `<a:p><a:endParaRPr lang="en-US" sz="${Math.round((base.fs || 18) * 100)}" dirty="0"/></a:p>`;
    return out;
  }

  const fillXml = (c) => c ? `<a:solidFill><a:srgbClr val="${c}"/></a:solidFill>` : '<a:noFill/>';
  function lnXml(line, tail) {
    if (!line) return '<a:ln><a:noFill/></a:ln>';
    return `<a:ln w="${emu(line.w)}">${fillXml(line.color)}${line.dash ? `<a:prstDash val="${line.dash}"/>` : ''}${tail ? '<a:tailEnd type="triangle"/>' : ''}</a:ln>`;
  }
  const PRST = { text: 'rect', rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', rtTriangle: 'rtTriangle', diamond: 'diamond', pentagon: 'pentagon', hexagon: 'hexagon', parallelogram: 'parallelogram', trapezoid: 'trapezoid', rightArrow: 'rightArrow', leftArrow: 'leftArrow', upArrow: 'upArrow', downArrow: 'downArrow', star5: 'star5', image: 'rect' };
  function xfrm(s, tag) {
    const ang = s.rot ? Math.round(((s.rot % 360) + 360) % 360 * 60000) : 0;
    return `<a:xfrm${ang ? ` rot="${ang}"` : ''}${s.flipH ? ' flipH="1"' : ''}${s.flipV ? ' flipV="1"' : ''}><a:off x="${emu(s.x)}" y="${emu(s.y)}"/><a:ext cx="${emu(Math.max(0, s.w))}" cy="${emu(Math.max(0, s.h))}"/></a:xfrm>`;
  }
  function bodyPr(s) {
    const anchor = { t: 't', ctr: 'ctr', b: 'b' }[s.valign || 't'] || 't';
    return `<a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="${anchor}"><a:noAutofit/></a:bodyPr><a:lstStyle/>`;
  }
  const MEDIA_EXT = { 'video/mp4': 'mp4', 'video/quicktime': 'mov', 'video/x-ms-wmv': 'wmv', 'video/x-msvideo': 'avi', 'video/webm': 'webm', 'audio/mpeg': 'mp3', 'audio/mp3': 'mp3', 'audio/wav': 'wav', 'audio/x-wav': 'wav', 'audio/wave': 'wav', 'audio/mp4': 'm4a', 'audio/x-m4a': 'm4a', 'audio/ogg': 'ogg', 'audio/x-ms-wma': 'wma' };
  const MEDIA_MIME = { mp4: 'video/mp4', m4v: 'video/mp4', mov: 'video/quicktime', wmv: 'video/x-ms-wmv', avi: 'video/x-msvideo', webm: 'video/webm', mp3: 'audio/mpeg', wav: 'audio/wav', m4a: 'audio/mp4', ogg: 'audio/ogg', wma: 'audio/x-ms-wma' };
  const BLANK_PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
  const PH_NAME = { title: 'Title', ctrTitle: 'Title', subTitle: 'Subtitle', body: 'Content Placeholder', ftr: 'Footer Placeholder', sldNum: 'Slide Number Placeholder' };

  function shapeXml(s, id, S, pres) {
    const base = { fs: s.fs, fc: s.fc, ff: s.ff };
    if (s.type === 'line') {
      return `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="Line ${id}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>${xfrm(s)}<a:prstGeom prst="line"><a:avLst/></a:prstGeom>${lnXml(s.line || { color: '000000', w: 2 }, s.tail)}</p:spPr></p:cxnSp>`;
    }
    if (s.type === 'image') {
      const d = dataUrlToBytes(pres.media[s.img] || '');
      if (!d) return '';
      const ext = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp' }[d.mime];
      if (!ext) return '';
      const name = 'image' + (++S.imgN) + '.' + ext;
      S.media.push({ path: 'ppt/media/' + name, bytes: d.bytes });
      S.exts.add(ext);
      const rid = 'rId' + (S.rid++);
      S.rels.push(`<Relationship Id="${rid}" Type="${REL}/image" Target="../media/${name}"/>`);
      const c = s.crop;
      const src = c && (c.l || c.t || c.r || c.b) ? `<a:srcRect l="${Math.round(c.l * 100000)}" t="${Math.round(c.t * 100000)}" r="${Math.round(c.r * 100000)}" b="${Math.round(c.b * 100000)}"/>` : '';
      return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}" descr="${xesc(s.alt || '')}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/>${src}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrm(s)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
    }
    if (s.type === 'media') {
      const d = dataUrlToBytes(pres.media[s.src] || '');
      const ext = d && MEDIA_EXT[d.mime];
      const po = dataUrlToBytes(pres.media[s.img] || '') || dataUrlToBytes(BLANK_PNG);
      const pext = { 'image/png': 'png', 'image/jpeg': 'jpeg' }[po.mime];
      if (!ext || !pext) return '';
      const mname = 'media' + (++S.G.mediaN) + '.' + ext;
      S.media.push({ path: 'ppt/media/' + mname, bytes: d.bytes });
      S.G.mediaExts[ext] = d.mime;
      const pname = 'image' + (++S.imgN) + '.' + pext;
      S.media.push({ path: 'ppt/media/' + pname, bytes: po.bytes });
      S.exts.add(pext);
      const r1 = 'rId' + (S.rid++), r2 = 'rId' + (S.rid++), r3 = 'rId' + (S.rid++);
      const vid = s.kind !== 'audio';
      S.rels.push(`<Relationship Id="${r1}" Type="${REL}/${vid ? 'video' : 'audio'}" Target="../media/${mname}"/><Relationship Id="${r2}" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="../media/${mname}"/><Relationship Id="${r3}" Type="${REL}/image" Target="../media/${pname}"/>`);
      return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${xesc(s.alt || (vid ? 'Video' : 'Audio'))}"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:${vid ? 'videoFile' : 'audioFile'} r:link="${r1}"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="${r2}"/></p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="${r3}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrm(s)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
    }
    if (s.type === 'chart') {
      const ch = s.chart;
      const data = { cats: ch.cats, series: ch.series.map(x => ({ name: x.name, color: x.color, vals: x.vals })) };
      const cn = ++S.G.chartN;
      S.G.charts.push({ n: cn, xml: Charts.chartXml({ type: ch.type, title: ch.title, legend: ch.legend, lit: true, series: ch.series.map(() => ({})) }, data) });
      const rid = 'rId' + (S.rid++);
      S.rels.push(`<Relationship Id="${rid}" Type="${REL}/chart" Target="../charts/chart${cn}.xml"/>`);
      return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Chart ${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${emu(s.x)}" y="${emu(s.y)}"/><a:ext cx="${emu(s.w)}" cy="${emu(s.h)}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rid}"/></a:graphicData></a:graphic></p:graphicFrame>`;
    }
    if (s.type === 'table') {
      const colW = s.colW, rowH = s.rowH;
      let x = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${emu(s.x)}" y="${emu(s.y)}"/><a:ext cx="${emu(colW.reduce((a, b) => a + b, 0))}" cy="${emu(rowH.reduce((a, b) => a + b, 0))}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>` +
        colW.map(w => `<a:gridCol w="${emu(w)}"/>`).join('') + '</a:tblGrid>';
      s.rows.forEach((row, ri) => {
        x += `<a:tr h="${emu(rowH[ri])}">`;
        row.forEach(cell => {
          const ln = ['L', 'R', 'T', 'B'].map(k => `<a:ln${k} w="12700"><a:solidFill><a:srgbClr val="A6A6A6"/></a:solidFill></a:ln${k}>`).join('');
          x += `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${textBodyParas(cell.html, { fs: s.fs || 14, fc: s.fc || '000000', ff: s.ff || 'Calibri' }, S)}</a:txBody><a:tcPr>${ln}${cell.fill ? fillXml(cell.fill) : '<a:noFill/>'}</a:tcPr></a:tc>`;
        });
        x += '</a:tr>';
      });
      return x + '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
    }
    // text boxes + preset shapes
    const prst = PRST[s.type] || 'rect';
    const isBox = s.type === 'text';
    let nv;
    if (s.ph) nv = `<p:nvSpPr><p:cNvPr id="${id}" name="${PH_NAME[s.ph] || 'Placeholder'} ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${s.ph}"${({ body: ' idx="1"', subTitle: ' idx="1"', ftr: ' sz="quarter" idx="11"', sldNum: ' sz="quarter" idx="12"' })[s.ph] || ''}/></p:nvPr></p:nvSpPr>`;
    else nv = `<p:nvSpPr><p:cNvPr id="${id}" name="${isBox ? 'TextBox' : 'Shape'} ${id}"/><p:cNvSpPr${isBox ? ' txBox="1"' : ''}/><p:nvPr/></p:nvSpPr>`;
    const ln = lnXml(s.line);
    const paras = s.ph === 'sldNum'
      ? `<a:p><a:pPr algn="r"/><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US" sz="${Math.round((s.fs || 12) * 100)}"><a:solidFill><a:srgbClr val="${s.fc || '7F7F7F'}"/></a:solidFill></a:rPr><a:t>${S.slideNo || 1}</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p>`
      : textBodyParas(s.html, base, S);
    return `<p:sp>${nv}<p:spPr>${xfrm(s)}<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${fillXml(s.fill)}${ln}</p:spPr><p:txBody>${bodyPr(s)}${paras}</p:txBody></p:sp>`;
  }

  const GRP = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

  function themeXml(th) {
    th = th || {};
    const ac = th.acc || ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
    const mj = xesc(th.major || 'Calibri Light'), mn = xesc(th.minor || 'Calibri');
    const clr = (n, v) => `<a:${n}><a:srgbClr val="${v}"/></a:${n}>`;
    const sol = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
    const lnS = (w) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr">${sol}<a:prstDash val="solid"/></a:ln>`;
    return XML + `<a:theme xmlns:a="${NS.a}" name="${xesc(th.name || 'Office')} Theme"><a:themeElements><a:clrScheme name="${xesc(th.name || 'Office')}"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>${clr('dk2', th.dk2 || '44546A')}${clr('lt2', th.lt2 || 'E7E6E6')}${ac.map((v, i) => clr('accent' + (i + 1), v)).join('')}${clr('hlink', '0563C1')}${clr('folHlink', '954F72')}</a:clrScheme>` +
      `<a:fontScheme name="Office"><a:majorFont><a:latin typeface="${mj}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="${mn}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
      `<a:fmtScheme name="Office"><a:fillStyleLst>${sol}${sol}${sol}</a:fillStyleLst><a:lnStyleLst>${lnS(6350)}${lnS(12700)}${lnS(19050)}</a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst>${sol}${sol}${sol}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
  }
  function lvlStyle(tag, sz, font, bullet) {
    let x = `<p:${tag}>`;
    for (let l = 1; l <= 9; l++) {
      const marL = bullet ? 228600 + (l - 1) * 457200 : (l - 1) * 457200;
      x += `<a:lvl${l}pPr marL="${marL}" indent="${bullet ? -228600 : 0}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="${bullet ? 1000 : 0}"/></a:spcBef>${bullet ? '<a:buFont typeface="Arial"/><a:buChar char="&#8226;"/>' : '<a:buNone/>'}<a:defRPr sz="${sz}" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="${font}"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${l}pPr>`;
    }
    return x + `</p:${tag}>`;
  }
  const phSp = (id, name, type, idx, x, y, w, h, extra) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${type}"${idx ? ` idx="${idx}"` : ''}/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr${extra || ''}/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Click to edit</a:t></a:r></a:p></p:txBody></p:sp>`;
  function masterXml(pres, MS) {
    let extra = '';
    if (pres && pres.master && pres.master.shapes && MS) { let id = 10; for (const s of pres.master.shapes) { if (s.ph) continue; const x = shapeXml(s, id, MS, pres); if (x) { extra += x; id++; } } }
    return XML + `<p:sldMaster ${NSDECL}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${GRP}${extra}` +
      phSp(2, 'Title Placeholder 1', 'title', 0, 838200, 365125, 10515600, 1325563, ' vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"') +
      phSp(3, 'Text Placeholder 2', 'body', 1, 838200, 1825625, 10515600, 4351338, ' vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"') +
      '</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
      '<p:sldLayoutIdLst>' + LAYOUTS.map((l, i) => `<p:sldLayoutId id="${2147483649 + i}" r:id="rId${i + 1}"/>`).join('') + '</p:sldLayoutIdLst>' +
      `<p:txStyles>${lvlStyle('titleStyle', 4400, '+mj-lt', false)}${lvlStyle('bodyStyle', 2800, '+mn-lt', true)}${lvlStyle('otherStyle', 1800, '+mn-lt', false)}</p:txStyles></p:sldMaster>`;
  }
  function layoutXml(l) {
    let sps = '';
    if (l.type === 'title') sps = phSp(2, 'Title 1', 'ctrTitle', 0, 1524000, 1122363, 9144000, 2387600, ' anchor="b"') + phSp(3, 'Subtitle 2', 'subTitle', 1, 1524000, 3602038, 9144000, 1655762);
    else if (l.type === 'obj') sps = phSp(2, 'Title 1', 'title', 0, 838200, 365125, 10515600, 1325563) + phSp(3, 'Content Placeholder 2', 'body', 1, 838200, 1825625, 10515600, 4351338);
    else if (l.type === 'titleOnly') sps = phSp(2, 'Title 1', 'title', 0, 838200, 365125, 10515600, 1325563);
    return XML + `<p:sldLayout ${NSDECL} type="${l.type}" preserve="1"><p:cSld name="${l.name}"><p:spTree>${GRP}${sps}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
  }

  function transXml(sl) {
    const t = sl.trans;
    const adv = sl.advAfter > 0 ? ` advTm="${Math.round(sl.advAfter * 1000)}"` : '';
    if ((!t || t.type === 'none') && !adv) return '';
    const spd = ['slow', 'med', 'fast'].includes(t && t.speed) ? t.speed : 'med';
    const inner = { fade: '<p:fade/>', push: '<p:push dir="l"/>', wipe: '<p:wipe dir="l"/>', split: '<p:split orient="vert" dir="out"/>', cover: '<p:cover dir="l"/>', pull: '<p:pull dir="l"/>', zoom: '<p:zoom dir="in"/>' }[t && t.type] || '';
    return `<p:transition spd="${spd}"${adv}>${inner}</p:transition>`;
  }
  function timingXml(sl, ids) {
    const list = sl.shapes.map((s, i) => ({ s, id: ids[i] })).filter(o => o.s.anim && o.id);
    const medias = sl.shapes.map((s, i) => ({ s, id: ids[i] })).filter(o => o.s.type === 'media' && o.id);
    if (!list.length && !medias.length) return '';
    let n = 2;
    const nid = () => ++n;
    let clicks = '', bld = '';
    for (const { s, id } of list) {
      const isSp0 = !['image', 'table', 'chart', 'line', 'media'].includes(s.type);
      const tg0 = `<p:tgtEl><p:spTgt spid="${id}"/></p:tgtEl>`;
      const special = /^(path|emph)-/.test(String(s.anim));
      if (special) {
        let eff0 = '', pid = 0, pcls = 'emph';
        if (s.anim === 'emph-spin') { pid = 8; eff0 = `<p:animRot by="21600000"><p:cBhvr><p:cTn id="${nid()}" dur="2000" fill="hold"/>${tg0}<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>`; }
        else if (s.anim === 'emph-grow') { pid = 6; eff0 = `<p:animScale><p:cBhvr><p:cTn id="${nid()}" dur="2000" fill="hold"/>${tg0}</p:cBhvr><p:by x="150000" y="150000"/></p:animScale>`; }
        else {
          pcls = 'path';
          const d = s.anim.slice(5);
          pid = { down: 42, left: 35, right: 63, up: 64 }[d] || 63;
          const path = { down: 'M 0 0 L 0 0.25 E', left: 'M 0 0 L -0.25 0 E', right: 'M 0 0 L 0.25 0 E', up: 'M 0 0 L 0 -0.25 E' }[d] || 'M 0 0 L 0.25 0 E';
          eff0 = `<p:animMotion origin="layout" path="${path}" pathEditMode="relative" ptsTypes=""><p:cBhvr><p:cTn id="${nid()}" dur="2000" fill="hold"/>${tg0}<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion>`;
        }
        const k1 = nid(), k2 = nid(), k3 = nid();
        clicks += `<p:par><p:cTn id="${k1}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${k2}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${k3}" presetID="${pid}" presetClass="${pcls}" presetSubtype="0" fill="hold"${isSp0 ? ' grpId="0"' : ''} nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${eff0}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`;
        if (isSp0) bld += `<p:bldP spid="${id}" grpId="0"${(s.fill || s.line) ? ' animBg="1"' : ''}/>`;
        continue;
      }
      const exit = String(s.anim).startsWith('exit-');
      const a = exit ? s.anim.slice(5) : s.anim;
      const preset = { appear: [1, 0], fade: [10, 0], wipe: [22, 8], fly: [2, 4] }[a] || [10, 0];
      const isSp = !['image', 'table', 'chart', 'line', 'media'].includes(s.type);
      const tg = `<p:tgtEl><p:spTgt spid="${id}"/></p:tgtEl>`;
      const dir = exit ? 'out' : 'in';
      const setv = `<p:set><p:cBhvr><p:cTn id="${nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${exit && a !== 'appear' ? 499 : 0}"/></p:stCondLst></p:cTn>${tg}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${exit ? 'hidden' : 'visible'}"/></p:to></p:set>`;
      let eff = '';
      if (a === 'fade') eff = `<p:animEffect transition="${dir}" filter="fade"><p:cBhvr><p:cTn id="${nid()}" dur="500"/>${tg}</p:cBhvr></p:animEffect>`;
      else if (a === 'wipe') eff = `<p:animEffect transition="${dir}" filter="wipe(left)"><p:cBhvr><p:cTn id="${nid()}" dur="500"/>${tg}</p:cBhvr></p:animEffect>`;
      else if (a === 'fly') {
        const mk = (attr, v0, v1) => `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="${nid()}" dur="500" fill="hold"/>${tg}<p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="${v0}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="${v1}"/></p:val></p:tav></p:tavLst></p:anim>`;
        eff = mk('ppt_x', '#ppt_x', '#ppt_x') + (exit ? mk('ppt_y', '#ppt_y', '1+#ppt_h/2') : mk('ppt_y', '1+#ppt_h/2', '#ppt_y'));
      }
      const c1 = nid(), c2 = nid(), c3 = nid();
      clicks += `<p:par><p:cTn id="${c1}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${c2}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${c3}" presetID="${preset[0]}" presetClass="${exit ? 'exit' : 'entr'}" presetSubtype="${preset[1]}" fill="hold"${isSp ? ' grpId="0"' : ''} nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${exit ? eff + setv : setv + eff}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`;
      if (isSp) bld += `<p:bldP spid="${id}" grpId="0"${(s.fill || s.line) ? ' animBg="1"' : ''}/>`;
    }
    const seq = clicks ? `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${clicks}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>` : '';
    const mnodes = medias.map(({ s, id }) => { const tag = s.kind === 'audio' ? 'audio' : 'video'; return `<p:${tag}><p:cMediaNode vol="80000"><p:cTn id="${nid()}" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${id}"/></p:tgtEl></p:cMediaNode></p:${tag}>`; }).join('');
    return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>${seq}${mnodes}</p:childTnLst></p:cTn></p:par></p:tnLst>${bld ? `<p:bldLst>${bld}</p:bldLst>` : ''}</p:timing>`;
  }

  async function write(pres) {
    const G = { chartN: 0, charts: [], mediaN: 0, mediaExts: {} };
    const zw = Zip.writer();
    const allExts = new Set(['png', 'jpeg']);
    const slideParts = [];
    const notesList = [];
    pres.slides.forEach((sl, i) => {
      const S = { rels: [], media: [], rid: 2, imgN: pres._imgBase || 0, exts: allExts, G, slideNo: i + 1 };
      pres._imgBase = (pres._imgBase || 0);
      let id = 2;
      const layoutIdx = Math.max(0, LAYOUTS.findIndex(l => l.key === sl.layout));
      let shapes = '';
      const ids = new Array(sl.shapes.length).fill(0);
      const done = new Set();
      const dims = (t) => t.type === 'table' ? [t.colW.reduce((a, b) => a + b, 0), t.rowH.reduce((a, b) => a + b, 0)] : [t.w, t.h];
      sl.shapes.forEach((s, idx) => {
        if (done.has(idx)) return;
        const mem = [];
        if (s.g !== undefined && s.g !== null) sl.shapes.forEach((t, j) => { if (t.g === s.g) mem.push(j); });
        if (mem.length > 1) {
          const gid = id++;
          let inner = '', x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
          for (const j of mem) {
            done.add(j);
            const t = sl.shapes[j];
            const x = shapeXml(t, id, S, pres);
            if (!x) continue;
            ids[j] = id++; inner += x;
            const [w, h] = dims(t);
            x1 = Math.min(x1, t.x); y1 = Math.min(y1, t.y); x2 = Math.max(x2, t.x + w); y2 = Math.max(y2, t.y + h);
          }
          if (!inner) return;
          const bx = `<a:off x="${emu(x1)}" y="${emu(y1)}"/><a:ext cx="${emu(Math.max(1, x2 - x1))}" cy="${emu(Math.max(1, y2 - y1))}"/><a:chOff x="${emu(x1)}" y="${emu(y1)}"/><a:chExt cx="${emu(Math.max(1, x2 - x1))}" cy="${emu(Math.max(1, y2 - y1))}"/>`;
          shapes += `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${gid}" name="Group ${gid}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm>${bx}</a:xfrm></p:grpSpPr>${inner}</p:grpSp>`;
          return;
        }
        const x = shapeXml(s, id, S, pres);
        if (x) { ids[idx] = id++; shapes += x; }
      });
      const bg = sl.bg ? `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${sl.bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>` : '';
      const xml = XML + `<p:sld ${NSDECL}${sl.hidden ? ' show="0"' : ''}${sl.noMaster && pres.master && pres.master.shapes && pres.master.shapes.length ? ' showMasterSp="0"' : ''}><p:cSld>${bg}<p:spTree>${GRP}${shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${transXml(sl)}${timingXml(sl, ids)}</p:sld>`;
      if (sl.notes && sl.notes.trim()) {
        const k = notesList.length + 1;
        notesList.push(i + 1);
        S.rels.push(`<Relationship Id="rIdNotes" Type="${REL}/notesSlide" Target="../notesSlides/notesSlide${k}.xml"/>`);
        const paras = sl.notes.replace(/\r/g, '').split('\n').map(t => t ? `<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${xesc(t)}</a:t></a:r></a:p>` : '<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>').join('');
        zw.add(`ppt/notesSlides/notesSlide${k}.xml`, XML + `<p:notes ${NSDECL}><p:cSld><p:spTree>${GRP}<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paras}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`);
        zw.add(`ppt/notesSlides/_rels/notesSlide${k}.xml.rels`, XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/notesMaster" Target="../notesMasters/notesMaster1.xml"/><Relationship Id="rId2" Type="${REL}/slide" Target="../slides/slide${i + 1}.xml"/></Relationships>`);
      }
      const rels = XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout${layoutIdx + 1}.xml"/>${S.rels.join('')}</Relationships>`;
      zw.add(`ppt/slides/slide${i + 1}.xml`, xml);
      zw.add(`ppt/slides/_rels/slide${i + 1}.xml.rels`, rels);
      for (const m of S.media) zw.add(m.path, m.bytes);
      pres._imgBase = S.imgN;
      slideParts.push(i + 1);
    });
    // slide master graphics (written before the content types so their image types are declared)
    const MS = { rels: [], media: [], rid: LAYOUTS.length + 2, imgN: pres._imgBase || 0, exts: allExts, G, slideNo: 0 };
    const mXml = masterXml(pres, MS);
    for (const m of MS.media) zw.add(m.path, m.bytes);
    pres._imgBase = MS.imgN;
    const n = pres.slides.length;
    for (const c of G.charts) zw.add(`ppt/charts/chart${c.n}.xml`, c.xml);
    zw.add('[Content_Types].xml', XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      [...allExts].map(e => `<Default Extension="${e}" ContentType="image/${e}"/>`).join('') +
      Object.keys(G.mediaExts).map(e => `<Default Extension="${e}" ContentType="${G.mediaExts[e]}"/>`).join('') +
      '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
      '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>' +
      LAYOUTS.map((_, i) => `<Override PartName="/ppt/slideLayouts/slideLayout${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`).join('') +
      slideParts.map(i => `<Override PartName="/ppt/slides/slide${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('') +
      '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
      G.charts.map(c => `<Override PartName="/ppt/charts/chart${c.n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`).join('') +
      (notesList.length ? '<Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>' +
        notesList.map((_, k) => `<Override PartName="/ppt/notesSlides/notesSlide${k + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`).join('') : '') +
      '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>' +
      '<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>' +
      '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>');
    zw.add('_rels/.rels', XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/></Relationships>`);
    const nmId = n + 6;
    zw.add('ppt/presentation.xml', XML + `<p:presentation ${NSDECL} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>${notesList.length ? `<p:notesMasterIdLst><p:notesMasterId r:id="rId${nmId}"/></p:notesMasterIdLst>` : ''}<p:sldIdLst>` +
      slideParts.map((i, k) => `<p:sldId id="${256 + k}" r:id="rId${i + 1}"/>`).join('') +
      `</p:sldIdLst><p:sldSz cx="${emu(pres.w)}" cy="${emu(pres.h)}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
    zw.add('ppt/_rels/presentation.xml.rels', XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL}/slideMaster" Target="slideMasters/slideMaster1.xml"/>` +
      slideParts.map(i => `<Relationship Id="rId${i + 1}" Type="${REL}/slide" Target="slides/slide${i}.xml"/>`).join('') +
      `<Relationship Id="rId${n + 2}" Type="${REL}/presProps" Target="presProps.xml"/><Relationship Id="rId${n + 3}" Type="${REL}/viewProps" Target="viewProps.xml"/><Relationship Id="rId${n + 4}" Type="${REL}/theme" Target="theme/theme1.xml"/><Relationship Id="rId${n + 5}" Type="${REL}/tableStyles" Target="tableStyles.xml"/>` +
      (notesList.length ? `<Relationship Id="rId${nmId}" Type="${REL}/notesMaster" Target="notesMasters/notesMaster1.xml"/>` : '') + '</Relationships>');
    if (notesList.length) {
      zw.add('ppt/theme/theme2.xml', themeXml());
      zw.add('ppt/notesMasters/notesMaster1.xml', XML + `<p:notesMaster ${NSDECL}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${GRP}` +
        '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="1143000"/><a:ext cx="5486400" cy="3086100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr></p:sp>' +
        '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="4400550"/><a:ext cx="5486400" cy="3600450"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r></a:p></p:txBody></p:sp>' +
        '</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:notesMaster>');
      zw.add('ppt/notesMasters/_rels/notesMaster1.xml.rels', XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/theme" Target="../theme/theme2.xml"/></Relationships>`);
    }
    zw.add('ppt/presProps.xml', XML + `<p:presentationPr ${NSDECL}/>`);
    zw.add('ppt/viewProps.xml', XML + `<p:viewPr ${NSDECL}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`);
    zw.add('ppt/tableStyles.xml', XML + `<a:tblStyleLst xmlns:a="${NS.a}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`);
    zw.add('ppt/theme/theme1.xml', themeXml(pres.theme));
    zw.add('ppt/slideMasters/slideMaster1.xml', mXml);
    zw.add('ppt/slideMasters/_rels/slideMaster1.xml.rels', XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      LAYOUTS.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout${i + 1}.xml"/>`).join('') + `<Relationship Id="rId${LAYOUTS.length + 1}" Type="${REL}/theme" Target="../theme/theme1.xml"/>${MS.rels.join('')}</Relationships>`);
    LAYOUTS.forEach((l, i) => {
      zw.add(`ppt/slideLayouts/slideLayout${i + 1}.xml`, layoutXml(l));
      zw.add(`ppt/slideLayouts/_rels/slideLayout${i + 1}.xml.rels`, XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);
    });
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    zw.add('docProps/core.xml', XML + `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Slides</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
    zw.add('docProps/app.xml', XML + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Slides</Application></Properties>');
    pres._imgBase = 0;
    return zw.toBlob(PPTX_MIME);
  }

  /* ====================================================================
     READ
     ==================================================================== */
  const kids = (el, ln) => { const r = []; for (let c = el && el.firstElementChild; c; c = c.nextElementSibling) if (c.localName === ln) r.push(c); return r; };
  const kid = (el, ln) => { for (let c = el && el.firstElementChild; c; c = c.nextElementSibling) if (c.localName === ln) return c; return null; };
  const num = (v, d) => { const n = parseFloat(v); return isNaN(n) ? d : n; };
  function parseXml(s) { const d = new DOMParser().parseFromString(s, 'application/xml'); if (d.getElementsByTagName('parsererror').length) throw new Error('Could not parse XML in the presentation.'); return d.documentElement; }
  function dirname(p) { const i = p.lastIndexOf('/'); return i < 0 ? '' : p.slice(0, i); }
  function resolvePath(base, t) { if (t.startsWith('/')) return t.slice(1); const parts = base ? base.split('/') : []; for (const s of t.split('/')) { if (s === '..') parts.pop(); else if (s !== '.') parts.push(s); } return parts.join('/'); }
  function parseRels(xml, base) { const out = {}; if (!xml) return out; for (const r of parseXml(xml).children) { const ext = r.getAttribute('TargetMode') === 'External'; out[r.getAttribute('Id')] = { type: r.getAttribute('Type'), target: r.getAttribute('Target'), path: ext ? null : resolvePath(base, r.getAttribute('Target')) }; } return out; }
  async function relsFor(z, part) { const dir = dirname(part); return parseRels(await z.textOrNull(dir + '/_rels/' + part.split('/').pop() + '.rels'), dir); }

  function rgbToHsl(r, g2, b) { r /= 255; g2 /= 255; b /= 255; const mx = Math.max(r, g2, b), mn = Math.min(r, g2, b); let h = 0, s = 0; const l = (mx + mn) / 2; if (mx !== mn) { const d = mx - mn; s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn); h = mx === r ? (g2 - b) / d + (g2 < b ? 6 : 0) : mx === g2 ? (b - r) / d + 2 : (r - g2) / d + 4; h /= 6; } return [h, s, l]; }
  function hslToRgb(h, s, l) { const f = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; }; if (s === 0) return [l, l, l].map(x => Math.round(x * 255)); const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; return [f(p, q, h + 1 / 3), f(p, q, h), f(p, q, h - 1 / 3)].map(x => Math.round(x * 255)); }
  const toHex = (rgb) => rgb.map(x => ('0' + Math.max(0, Math.min(255, x)).toString(16)).slice(-2)).join('').toUpperCase();
  const PRST_CLR = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080', orange: 'FFA500', purple: '800080', cyan: '00FFFF', magenta: 'FF00FF', ltGray: 'D3D3D3', dkGray: 'A9A9A9' };

  function resolveColor(el, T, phClr) {
    // el: the colour element (srgbClr/schemeClr/...) or a container such as solidFill
    if (!el) return null;
    if (el.localName === 'solidFill') el = el.firstElementChild;
    if (!el) return null;
    let hex = null;
    const v = el.getAttribute('val');
    switch (el.localName) {
      case 'srgbClr': hex = (v || '000000').toUpperCase(); break;
      case 'sysClr': hex = (el.getAttribute('lastClr') || '000000').toUpperCase(); break;
      case 'prstClr': hex = PRST_CLR[v] || '000000'; break;
      case 'schemeClr': {
        let key = v;
        if (key === 'phClr') hex = phClr || '000000';
        else { key = T.clrMap[key] || key; hex = T.colors[key] || '000000'; }
        break;
      }
      default: return null;
    }
    for (const m of el.children) {
      const val = num(m.getAttribute('val'), 0) / 100000;
      if (m.localName === 'lumMod' || m.localName === 'lumOff' || m.localName === 'satMod') {
        // collect in order of appearance
      }
    }
    let rgb = [0, 2, 4].map(i => parseInt(hex.slice(i, i + 2), 16));
    let hsl = null;
    for (const m of el.children) {
      const val = num(m.getAttribute('val'), 0) / 100000;
      switch (m.localName) {
        case 'lumMod': hsl = hsl || rgbToHsl(...rgb); hsl[2] = Math.min(1, hsl[2] * val); rgb = hslToRgb(...hsl); break;
        case 'lumOff': hsl = hsl || rgbToHsl(...rgb); hsl[2] = Math.max(0, Math.min(1, hsl[2] + val)); rgb = hslToRgb(...hsl); break;
        case 'satMod': hsl = hsl || rgbToHsl(...rgb); hsl[1] = Math.min(1, hsl[1] * val); rgb = hslToRgb(...hsl); break;
        case 'tint': rgb = rgb.map(c => Math.round(c * val + 255 * (1 - val))); hsl = null; break;
        case 'shade': rgb = rgb.map(c => Math.round(c * val)); hsl = null; break;
        default: break;
      }
      if (['lumMod', 'lumOff', 'satMod', 'tint', 'shade'].includes(m.localName) && m.localName !== 'lumMod' && m.localName !== 'lumOff' && m.localName !== 'satMod') hsl = null;
    }
    return toHex(rgb);
  }
  function fillOf(spPr, T, phClr) {
    // returns undefined (not specified), null (noFill) or hex
    if (!spPr) return undefined;
    if (kid(spPr, 'noFill')) return null;
    const sf = kid(spPr, 'solidFill');
    if (sf) return resolveColor(sf, T, phClr);
    const gf = kid(spPr, 'gradFill');
    if (gf) { const gs = kid(kid(gf, 'gsLst'), 'gs'); return gs ? resolveColor(gs, T, phClr) : undefined; }
    return undefined;
  }

  async function loadMaster(z, path, cache, T0) {
    if (cache[path]) return cache[path];
    const root = parseXml(await z.text(path));
    const rels = await relsFor(z, path);
    const info = { path, root, rels, ph: {}, deco: [], T: T0 };
    cache[path] = info;
    return info;
  }

  async function read(buffer) {
    const z = await Zip.read(buffer);
    let presPath = 'ppt/presentation.xml';
    if (z.has('_rels/.rels')) { const r = parseRels(await z.text('_rels/.rels'), ''); for (const k in r) if (/officeDocument$/.test(r[k].type) && r[k].path) presPath = r[k].path; }
    if (!z.has(presPath)) throw new Error('This file does not contain a presentation (ppt/presentation.xml not found).');
    const presRoot = parseXml(await z.text(presPath));
    const presRels = await relsFor(z, presPath);
    const sz = kid(presRoot, 'sldSz');
    const pres = { w: Math.round(num(sz && sz.getAttribute('cx'), 12192000) / EMU), h: Math.round(num(sz && sz.getAttribute('cy'), 6858000) / EMU), media: {}, slides: [], notes: [] };
    // default text style
    const dts = kid(presRoot, 'defaultTextStyle');
    let mediaN = 0;
    const mediaCache = {};
    const getMedia = async (path) => {
      if (mediaCache[path]) return mediaCache[path];
      const ext = path.split('.').pop().toLowerCase();
      const mime = MIME_BY_EXT[ext] || MEDIA_MIME[ext];
      if (!mime || !z.has(path)) return null;
      const id = 'm' + (++mediaN);
      pres.media[id] = bytesToDataUrl(await z.bytes(path), mime);
      mediaCache[path] = id;
      return id;
    };
    const themeCache = {}, masterCache = {}, layoutCache = {};
    const lost = { n: 0 };
    const sldIds = kids(kid(presRoot, 'sldIdLst'), 'sldId');
    for (const sid of sldIds) {
      const rid = sid.getAttributeNS(NS.r, 'id') || sid.getAttribute('r:id');
      const rel = presRels[rid];
      if (!rel || !z.has(rel.path)) continue;
      const slideRoot = parseXml(await z.text(rel.path));
      const slideRels = await relsFor(z, rel.path);
      let layoutRel = null;
      for (const k in slideRels) if (/\/slideLayout$/.test(slideRels[k].type)) layoutRel = slideRels[k];
      let layout = null, master = null, T = { colors: {}, clrMap: {}, major: 'Calibri Light', minor: 'Calibri' };
      if (layoutRel && z.has(layoutRel.path)) {
        if (!layoutCache[layoutRel.path]) layoutCache[layoutRel.path] = { root: parseXml(await z.text(layoutRel.path)), rels: await relsFor(z, layoutRel.path) };
        layout = layoutCache[layoutRel.path];
        let mrel = null;
        for (const k in layout.rels) if (/\/slideMaster$/.test(layout.rels[k].type)) mrel = layout.rels[k];
        if (mrel && z.has(mrel.path)) {
          if (!masterCache[mrel.path]) {
            const mroot = parseXml(await z.text(mrel.path));
            const mrels = await relsFor(z, mrel.path);
            let trel = null; for (const k in mrels) if (/\/theme$/.test(mrels[k].type)) trel = mrels[k];
            const theme = { colors: {}, major: 'Calibri Light', minor: 'Calibri' };
            if (trel && z.has(trel.path)) {
              const tr = parseXml(await z.text(trel.path));
              const cs = tr.getElementsByTagNameNS('*', 'clrScheme')[0];
              if (cs) for (const c of cs.children) { const e = c.firstElementChild; if (e) theme.colors[c.localName] = (e.localName === 'sysClr' ? e.getAttribute('lastClr') : e.getAttribute('val')) || '000000'; }
              const mj = tr.getElementsByTagNameNS('*', 'majorFont')[0], mn = tr.getElementsByTagNameNS('*', 'minorFont')[0];
              const lat = (f) => { const l = f && f.getElementsByTagNameNS('*', 'latin')[0]; return l && l.getAttribute('typeface'); };
              theme.major = lat(mj) || theme.major; theme.minor = lat(mn) || theme.minor;
            }
            const cm = kid(mroot, 'clrMap');
            const clrMap = {};
            if (cm) for (const a of cm.attributes) clrMap[a.name] = a.value;
            if (!pres.theme) pres.theme = { name: 'Imported', bg: 'FFFFFF', fg: theme.colors.dk1 || '000000', sub: '595959', dk2: theme.colors.dk2 || '44546A', lt2: theme.colors.lt2 || 'E7E6E6', acc: [1, 2, 3, 4, 5, 6].map(i => theme.colors['accent' + i] || '4472C4'), major: theme.major, minor: theme.minor };
            masterCache[mrel.path] = { root: mroot, rels: mrels, T: { colors: theme.colors, clrMap, major: theme.major, minor: theme.minor } };
          }
          master = masterCache[mrel.path];
          T = master.T;
        }
      }
      const ctxs = { z, pres, T, getMedia, lost, dts, layout, master, slideRels, grp: { n: 0 } };
      const slide = { bg: null, layout: 'blank', shapes: [], notes: '' };
      if (slideRoot.getAttribute('show') === '0') slide.hidden = true;
      pres.slides.push(slide);
      for (const k in slideRels) {
        if (/\/notesSlide$/.test(slideRels[k].type) && slideRels[k].path && z.has(slideRels[k].path)) {
          try {
            const nroot = parseXml(await z.text(slideRels[k].path));
            const tree = kid(kid(nroot, 'cSld'), 'spTree');
            for (const sp of kids(tree, 'sp')) {
              const ph = kid(kid(kid(sp, 'nvSpPr'), 'nvPr'), 'ph');
              if (ph && ph.getAttribute('type') === 'body') slide.notes = kids(kid(sp, 'txBody'), 'p').map(p => [...p.getElementsByTagNameNS('*', 't')].map(t => t.textContent).join('')).join('\n').replace(/\n+$/, '');
            }
          } catch (e) { /* ignore notes errors */ }
        }
      }
      slide.bg = findBg(slideRoot, T) || (layout && findBg(layout.root, T)) || (master && findBg(master.root, T)) || 'FFFFFF';
      // decorative (non placeholder) shapes from master / layout
      const showMaster = slideRoot.getAttribute('showMasterSp') !== '0';
      const layoutShows = !layout || layout.root.getAttribute('showMasterSp') !== '0';
      // graphics on the first slide master become the editable Slide Master; other masters are copied onto their slides
      const firstM = !!master && (pres._mroot = pres._mroot || master.root) === master.root;
      if (showMaster && layoutShows && master) {
        if (firstM) {
          if (!pres.master) { pres.master = { shapes: [] }; await collect(kid(kid(master.root, 'cSld'), 'spTree'), pres.master.shapes, { ...ctxs, rels: master.rels, deco: true, ident: IDENT }); pres.master.shapes.forEach((s, i) => { s.id = 900000 + i; delete s.xid; delete s.g; }); }
        } else await collect(kid(kid(master.root, 'cSld'), 'spTree'), slide.shapes, { ...ctxs, rels: master.rels, deco: true, ident: IDENT });
      }
      if (!(showMaster && layoutShows && firstM)) slide.noMaster = true;
      if (showMaster && layout) await collect(kid(kid(layout.root, 'cSld'), 'spTree'), slide.shapes, { ...ctxs, rels: layout.rels, deco: true, ident: IDENT });
      await collect(kid(kid(slideRoot, 'cSld'), 'spTree'), slide.shapes, { ...ctxs, rels: slideRels, deco: false, ident: IDENT });
      // transition
      const trEl = [...slideRoot.getElementsByTagNameNS('*', 'transition')][0];
      if (trEl) {
        const ch0 = trEl.firstElementChild;
        const nm = ch0 ? ch0.localName : 'none';
        const t = ({ fade: 'fade', push: 'push', wipe: 'wipe', split: 'split', cover: 'cover', pull: 'pull', zoom: 'zoom', none: 'none' })[nm] || 'fade';
        if (t !== 'none') slide.trans = { type: t, speed: trEl.getAttribute('spd') || 'med' };
        const adv = parseInt(trEl.getAttribute('advTm'), 10);
        if (adv > 0) slide.advAfter = adv / 1000;
      }
      // entrance animations
      const tmEl = slideRoot.getElementsByTagNameNS('*', 'timing')[0];
      if (tmEl) {
        for (const ctn of tmEl.getElementsByTagNameNS('*', 'cTn')) {
          const pcls = ctn.getAttribute('presetClass');
          if (pcls !== 'entr' && pcls !== 'exit' && pcls !== 'path' && pcls !== 'emph') continue;
          const tgt = ctn.getElementsByTagNameNS('*', 'spTgt')[0];
          if (!tgt) continue;
          const shp = slide.shapes.find(s => s.xid !== undefined && String(s.xid) === tgt.getAttribute('spid'));
          const pidv = ctn.getAttribute('presetID');
          if (shp && !shp.anim && pcls === 'path') { shp.anim = ({ 42: 'path-down', 35: 'path-left', 63: 'path-right', 64: 'path-up' })[pidv] || 'path-right'; continue; }
          if (shp && !shp.anim && pcls === 'emph') { shp.anim = pidv === '6' ? 'emph-grow' : 'emph-spin'; continue; }
          if (shp && !shp.anim) shp.anim = (pcls === 'exit' ? 'exit-' : '') + (({ 1: 'appear', 10: 'fade', 22: 'wipe', 2: 'fly' })[ctn.getAttribute('presetID')] || 'fade');
        }
      }
      const lt = layout && layout.root.getAttribute('type');
      slide.layout = ({ title: 'title', obj: 'titleContent', titleOnly: 'titleOnly', blank: 'blank' })[lt] || (slide.shapes.some(s => s.ph) ? 'titleContent' : 'blank');
      slide.shapes.forEach((s, i) => { s.id = i + 1; });
    }
    pres.lost = lost.n;
    delete pres._mroot;
    if (pres.master && !pres.master.shapes.length) delete pres.master;
    if (!pres.master) pres.slides.forEach(sl => { delete sl.noMaster; });
    return pres;
  }
  const IDENT = { sx: 1, sy: 1, tx: 0, ty: 0 };

  function findBg(root, T) {
    const bg = kid(kid(root, 'cSld'), 'bg');
    if (!bg) return null;
    const bp = kid(bg, 'bgPr');
    if (bp) { const f = fillOf(bp, T); return f || null; }
    const br = kid(bg, 'bgRef');
    if (br) { const c = br.firstElementChild; return c ? resolveColor(c, T) : null; }
    return null;
  }

  function xfrmOf(spPr, tr) {
    const x = kid(spPr, 'xfrm');
    if (!x) return null;
    const off = kid(x, 'off'), ext = kid(x, 'ext');
    if (!off || !ext) return null;
    const X = num(off.getAttribute('x'), 0) / EMU, Y = num(off.getAttribute('y'), 0) / EMU, W = num(ext.getAttribute('cx'), 0) / EMU, H = num(ext.getAttribute('cy'), 0) / EMU;
    return { x: X * tr.sx + tr.tx, y: Y * tr.sy + tr.ty, w: W * tr.sx, h: H * tr.sy, rot: num(x.getAttribute('rot'), 0) / 60000, flipH: x.getAttribute('flipH') === '1', flipV: x.getAttribute('flipV') === '1' };
  }

  // find placeholder shape in layout/master matching type/idx
  function findPh(root, ph) {
    if (!root || !ph) return null;
    const tree = kid(kid(root, 'cSld'), 'spTree');
    if (!tree) return null;
    const type = ph.getAttribute('type'), idx = ph.getAttribute('idx');
    const norm = (t) => (t === 'ctrTitle' ? 'title' : (t === 'subTitle' ? 'body' : (t || 'body')));
    let best = null;
    for (const sp of tree.children) {
      if (sp.localName !== 'sp') continue;
      const p = kid(kid(kid(sp, 'nvSpPr'), 'nvPr'), 'ph');
      if (!p) continue;
      const t2 = p.getAttribute('type'), i2 = p.getAttribute('idx');
      if (idx && i2 === idx && (norm(t2) === norm(type) || !type)) return sp;
      if (!idx && norm(t2) === norm(type)) best = best || sp;
      if (idx && !best && norm(t2) === norm(type) && !i2) best = sp;
    }
    return best;
  }

  async function collect(tree, out, c) {
    if (!tree) return;
    for (const el of tree.children) {
      const ln = el.localName;
      if (ln === 'sp' || ln === 'cxnSp') await readSp(el, out, c);
      else if (ln === 'pic') await readPic(el, out, c);
      else if (ln === 'graphicFrame') await readFrame(el, out, c);
      else if (ln === 'grpSp') {
        const gp = kid(el, 'grpSpPr'), x = kid(gp, 'xfrm');
        let tr = c.ident;
        if (x) {
          const off = kid(x, 'off'), ext = kid(x, 'ext'), co = kid(x, 'chOff'), ce = kid(x, 'chExt');
          const cw = num(ce.getAttribute('cx'), 1) || 1, ch = num(ce.getAttribute('cy'), 1) || 1;
          const sx = num(ext.getAttribute('cx'), cw) / cw, sy = num(ext.getAttribute('cy'), ch) / ch;
          const t2 = { sx: sx, sy: sy, tx: num(off.getAttribute('x'), 0) / EMU - num(co.getAttribute('x'), 0) / EMU * sx, ty: num(off.getAttribute('y'), 0) / EMU - num(co.getAttribute('y'), 0) / EMU * sy };
          tr = { sx: c.ident.sx * t2.sx, sy: c.ident.sy * t2.sy, tx: c.ident.tx + t2.tx * c.ident.sx, ty: c.ident.ty + t2.ty * c.ident.sy };
          // child coordinates are in EMU/EMU already scaled; apply later via ident (offsets expressed in px)
          tr = { sx: t2.sx * c.ident.sx, sy: t2.sy * c.ident.sy, tx: c.ident.tx + t2.tx * c.ident.sx, ty: c.ident.ty + t2.ty * c.ident.sy };
        }
        const n0 = out.length;
        await collect(el, out, { ...c, ident: tr });
        if (!c.deco && c.grp && out.length - n0 > 1) { const gk = 'r' + (++c.grp.n); for (let k = n0; k < out.length; k++) out[k].g = gk; }
      }
    }
  }

  async function readPic(el, out, c) {
    const spPr = kid(el, 'spPr');
    const xf = xfrmOf(spPr, c.ident);
    if (!xf) return;
    const bf = kid(el, 'blipFill'), blip = bf && kid(bf, 'blip');
    const rid = blip && (blip.getAttributeNS(NS.r, 'embed') || blip.getAttribute('r:embed'));
    const rel = rid && c.rels[rid];
    if (!rel || !rel.path) return;
    const id = await c.getMedia(rel.path);
    if (!id) { c.lost.n++; return; }
    const nvPr = kid(kid(el, 'nvPicPr'), 'nvPr');
    const mf = nvPr && (kid(nvPr, 'videoFile') || kid(nvPr, 'audioFile'));
    if (mf) {
      const m14 = nvPr.getElementsByTagNameNS('*', 'media')[0];
      const mrid = (m14 && (m14.getAttributeNS(NS.r, 'embed') || m14.getAttribute('r:embed'))) || mf.getAttributeNS(NS.r, 'link') || mf.getAttribute('r:link');
      const mrel = mrid && c.rels[mrid];
      const mid = mrel && mrel.path ? await c.getMedia(mrel.path) : null;
      if (mid) {
        const cp = kid(kid(el, 'nvPicPr'), 'cNvPr');
        out.push({ type: 'media', kind: mf.localName === 'audioFile' ? 'audio' : 'video', ...xf, src: mid, img: id, alt: cp ? cp.getAttribute('name') || '' : '', fill: null, line: null, xid: cp && !c.deco ? cp.getAttribute('id') : undefined });
        return;
      }
    }
    const sr = kid(bf, 'srcRect');
    const crop = sr ? { l: num(sr.getAttribute('l'), 0) / 100000, t: num(sr.getAttribute('t'), 0) / 100000, r: num(sr.getAttribute('r'), 0) / 100000, b: num(sr.getAttribute('b'), 0) / 100000 } : null;
    const nv = kid(kid(el, 'nvPicPr'), 'cNvPr');
    out.push({ type: 'image', ...xf, img: id, crop, alt: nv ? nv.getAttribute('descr') || '' : '', fill: null, line: null, xid: nv && !c.deco ? nv.getAttribute('id') : undefined });
  }

  async function readFrame(el, out, c) {
    const x = kid(el, 'xfrm');
    const off = kid(x, 'off'), ext = kid(x, 'ext');
    if (!off || !ext) return;
    const tr = c.ident;
    const X = num(off.getAttribute('x'), 0) / EMU * tr.sx + tr.tx, Y = num(off.getAttribute('y'), 0) / EMU * tr.sy + tr.ty, W = num(ext.getAttribute('cx'), 0) / EMU * tr.sx, H = num(ext.getAttribute('cy'), 0) / EMU * tr.sy;
    const gd = kid(kid(el, 'graphic'), 'graphicData');
    const tbl = gd && kid(gd, 'tbl');
    const chEl = gd && kid(gd, 'chart');
    const xidOf = () => { const nv = kid(el, 'nvGraphicFramePr'); const cp = nv && kid(nv, 'cNvPr'); return cp && !c.deco ? cp.getAttribute('id') : undefined; };
    if (chEl) {
      const rid = chEl.getAttributeNS(NS.r, 'id') || chEl.getAttribute('r:id');
      const rel = c.rels[rid];
      let parsed = null;
      try { if (rel && rel.path && c.z.has(rel.path)) parsed = Charts.parseChartLit(parseXml(await c.z.text(rel.path))); } catch (e) { parsed = null; }
      if (parsed) { out.push({ type: 'chart', x: X, y: Y, w: W, h: H, rot: 0, chart: parsed, fill: null, line: null, xid: xidOf() }); return; }
    }
    if (!tbl) {
      c.lost.n++;
      out.push({ type: 'rect', x: X, y: Y, w: W, h: H, rot: 0, fill: 'F2F2F2', line: { color: 'A6A6A6', w: 1, dash: 'dash' }, html: '<p style="text-align:center">(Unsupported object: chart, diagram or embedded content)</p>', valign: 'ctr', fs: 14, fc: '7F7F7F', ff: 'Calibri' });
      return;
    }
    const colW = kids(kid(tbl, 'tblGrid'), 'gridCol').map(g2 => num(g2.getAttribute('w'), 0) / EMU * tr.sx);
    const rowH = [], rows = [];
    for (const tr2 of kids(tbl, 'tr')) {
      rowH.push(num(tr2.getAttribute('h'), 370840) / EMU * tr.sy);
      const row = [];
      for (const tc of kids(tr2, 'tc')) {
        const tx = kid(tc, 'txBody');
        const tcPr = kid(tc, 'tcPr');
        const f = fillOf(tcPr, c.T);
        const html = tx ? paragraphsHtml(tx, c, { defSize: 18, tr: null, phType: 'other', chain: [], bodyColor: '000000', scale: 1 }) : '<p><br></p>';
        row.push({ html, fill: f || null });
        const gs = parseInt(tc.getAttribute('gridSpan') || '1', 10);
        for (let k = 1; k < gs; k++) row.push({ html: '<p><br></p>', fill: null });
      }
      rows.push(row);
    }
    if (!rows.length) return;
    out.push({ type: 'table', x: X, y: Y, w: W, h: H, rot: 0, rows, colW, rowH, fs: 14, fc: '000000', ff: c.T.minor, xid: xidOf() });
  }

  async function readSp(el, out, c) {
    const isCxn = el.localName === 'cxnSp';
    const nvPr = kid(kid(el, isCxn ? 'nvCxnSpPr' : 'nvSpPr'), 'nvPr');
    const ph = nvPr && kid(nvPr, 'ph');
    if (c.deco && ph) return; // decorative pass: skip placeholders
    const spPr = kid(el, 'spPr');
    let xf = xfrmOf(spPr, c.ident);
    const style = kid(el, 'style');
    const txBody = kid(el, 'txBody');
    // inherit geometry from layout/master placeholders
    let layoutSp = null, masterSp = null;
    if (ph && c.layout) layoutSp = findPh(c.layout.root, ph);
    if (ph && c.master) masterSp = findPh(c.master.root, ph);
    if (!xf && ph) {
      xf = (layoutSp && xfrmOf(kid(layoutSp, 'spPr'), IDENT)) || (masterSp && xfrmOf(kid(masterSp, 'spPr'), IDENT));
    }
    if (!xf) return;
    const text = txBody ? txBody.textContent.replace(/\s+/g, '') : '';
    let fill = fillOf(spPr, c.T);
    let line;
    const ln = spPr && kid(spPr, 'ln');
    const styleClr = (tag) => { const r = style && kid(style, tag); const idx = r ? parseInt(r.getAttribute('idx'), 10) : 0; return { idx, clr: r && r.firstElementChild ? resolveColor(r.firstElementChild, c.T) : null }; };
    if (fill === undefined) { const fr = styleClr('fillRef'); fill = fr.idx > 0 ? fr.clr : null; }
    if (ln) {
      if (kid(ln, 'noFill')) line = null;
      else {
        const lc = kid(ln, 'solidFill') ? resolveColor(kid(ln, 'solidFill'), c.T) : null;
        const lr = styleClr('lnRef');
        const col = lc || (lr.idx > 0 ? lr.clr : null);
        const w = ln.getAttribute('w') ? num(ln.getAttribute('w'), 12700) / EMU : (lr.idx > 0 ? lr.idx * 1.33 : 1);
        const pd = kid(ln, 'prstDash');
        line = col ? { color: col, w: Math.max(0.5, w), dash: pd && pd.getAttribute('val') !== 'solid' ? 'dash' : undefined } : null;
      }
    } else { const lr = styleClr('lnRef'); line = lr.idx > 0 && lr.clr ? { color: lr.clr, w: Math.max(1, lr.idx * 1.33) } : null; }
    const geom = spPr && kid(spPr, 'prstGeom');
    let prst = geom ? geom.getAttribute('prst') : 'rect';
    if (!geom && spPr && kid(spPr, 'custGeom')) prst = 'rect';
    const KNOWN = ['rect', 'roundRect', 'ellipse', 'triangle', 'rtTriangle', 'diamond', 'pentagon', 'hexagon', 'parallelogram', 'trapezoid', 'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'star5'];
    let type = 'rect';
    if (isCxn || prst === 'line' || prst === 'straightConnector1') type = 'line';
    else if (KNOWN.includes(prst)) type = prst;
    else if (prst === 'flowChartProcess' || prst === 'snip1Rect' || prst === 'round2SameRect') type = 'rect';
    else if (prst === 'oval') type = 'ellipse';
    const s = { type, ...xf, fill: type === 'line' ? null : fill, line };
    { const nvp = kid(el, isCxn ? 'nvCxnSpPr' : 'nvSpPr'); const cp = nvp && kid(nvp, 'cNvPr'); if (cp && !c.deco) s.xid = cp.getAttribute('id'); }
    if (type === 'line') {
      s.line = line || { color: '000000', w: 1 };
      const tail = ln && kid(ln, 'tailEnd');
      s.tail = !!(tail && tail.getAttribute('type') && tail.getAttribute('type') !== 'none');
      out.push(s);
      return;
    }
    // text
    const hasText = text.length > 0;
    // empty title/body placeholders stay (they are the "click to add" boxes); empty date/footer/number/picture ones go
    if (ph && !hasText && !fill && !line && !/^(title|ctrTitle|subTitle|body)$/.test(ph.getAttribute('type') || 'body')) return;
    const phType = ph ? (ph.getAttribute('type') || 'body') : null;
    const kind = !ph ? 'other' : (/title/i.test(phType) ? 'title' : (phType === 'body' || phType === 'subTitle' || phType === 'obj' || !phType ? 'body' : 'other'));
    const chain = [];
    const mTx = c.master && kid(c.master.root, 'txStyles');
    if (c.dts) chain.push(c.dts);
    if (mTx) { const st = kid(mTx, kind === 'title' ? 'titleStyle' : (kind === 'body' ? 'bodyStyle' : 'otherStyle')); if (st) chain.push(st); }
    for (const sp of [masterSp, layoutSp]) { const ls = sp && kid(kid(sp, 'txBody'), 'lstStyle'); if (ls) chain.push(ls); }
    const ownLst = txBody && kid(txBody, 'lstStyle');
    if (ownLst) chain.push(ownLst);
    const bp = txBody && kid(txBody, 'bodyPr');
    const lbp = layoutSp && kid(kid(layoutSp, 'txBody'), 'bodyPr');
    const anchor = (bp && bp.getAttribute('anchor')) || (lbp && lbp.getAttribute('anchor')) || (masterSp && kid(kid(masterSp, 'txBody'), 'bodyPr') && kid(kid(masterSp, 'txBody'), 'bodyPr').getAttribute('anchor')) || 't';
    const naf = bp && kid(bp, 'normAutofit');
    const scale = naf && naf.getAttribute('fontScale') ? num(naf.getAttribute('fontScale'), 100000) / 100000 : 1;
    const fr = style && kid(style, 'fontRef');
    const bodyColor = fr && fr.firstElementChild ? resolveColor(fr.firstElementChild, c.T) : (c.T.colors[c.T.clrMap.tx1 || 'dk1'] || '000000');
    s.valign = anchor === 'ctr' ? 'ctr' : (anchor === 'b' ? 'b' : 't');
    if (txBody) s.html = paragraphsHtml(txBody, c, { chain, kind, scale, bodyColor, ph: !!ph }); else s.html = '<p><br></p>';
    s.fs = 18; s.fc = bodyColor; s.ff = c.T.minor;
    if (ph) s.ph = ({ ctrTitle: 'ctrTitle', title: 'title', subTitle: 'subTitle', body: 'body', ftr: 'ftr', sldNum: 'sldNum' })[phType] || 'body';
    if (!fill && !line && type === 'rect') s.type = 'text';
    if (!ph && s.type === 'rect' && !hasText && !fill && !line) return;
    if (c.deco && !hasText && !fill && !line) return;
    out.push(s);
  }

  function mergeLvl(chain, lvl) {
    const out = { attrs: {}, def: {}, bu: undefined, spcBef: undefined, spcAft: undefined, lnSpc: undefined };
    for (const lst of chain) {
      const p = kid(lst, 'lvl' + (lvl + 1) + 'pPr');
      if (!p) continue;
      for (const a of p.attributes) out.attrs[a.name] = a.value;
      const d = kid(p, 'defRPr');
      if (d) { for (const a of d.attributes) out.def[a.name] = a.value; const sf = kid(d, 'solidFill'); if (sf) out.def._fill = sf; const lat = kid(d, 'latin'); if (lat) out.def._font = lat.getAttribute('typeface'); }
      for (const ch of p.children) {
        if (ch.localName === 'buNone') out.bu = { none: true };
        else if (ch.localName === 'buChar') out.bu = { char: ch.getAttribute('char') };
        else if (ch.localName === 'buAutoNum') out.bu = { num: true };
        else if (ch.localName === 'spcBef') out.spcBef = ch; else if (ch.localName === 'spcAft') out.spcAft = ch; else if (ch.localName === 'lnSpc') out.lnSpc = ch;
      }
    }
    return out;
  }
  const escH = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  function paragraphsHtml(txBody, c, o) {
    const chain = o.chain || [];
    let html = '';
    const scale = o.scale || 1;
    for (const p of kids(txBody, 'p')) {
      const pPr = kid(p, 'pPr');
      const lvl = pPr ? parseInt(pPr.getAttribute('lvl') || '0', 10) : 0;
      const eff = mergeLvl(chain, lvl);
      const pa = {};
      if (pPr) for (const a of pPr.attributes) pa[a.name] = a.value;
      const attrs = Object.assign({}, eff.attrs, pa);
      let bu = eff.bu;
      let spcBef = eff.spcBef, spcAft = eff.spcAft, lnSpc = eff.lnSpc;
      if (pPr) for (const ch of pPr.children) {
        if (ch.localName === 'buNone') bu = { none: true };
        else if (ch.localName === 'buChar') bu = { char: ch.getAttribute('char') };
        else if (ch.localName === 'buAutoNum') bu = { num: true };
        else if (ch.localName === 'spcBef') spcBef = ch; else if (ch.localName === 'spcAft') spcAft = ch; else if (ch.localName === 'lnSpc') lnSpc = ch;
      }
      let css = '';
      const al = { l: 'left', ctr: 'center', r: 'right', just: 'justify' }[attrs.algn];
      if (al && al !== 'left') css += `text-align:${al};`;
      const marL = num(attrs.marL, 0) / EMU, ind = num(attrs.indent, 0) / EMU;
      let pAttr = '';
      if (bu && !bu.none) {
        pAttr = bu.num ? ' data-bu="num"' : ` data-bu="dot" data-bc="${escH(bu.char || '•')}"`;
        pAttr += lvl ? ` data-l="${lvl}"` : '';
        const hang = ind < 0 ? -ind : 24;
        css += `margin-left:${Math.round(marL || hang)}px;text-indent:${-Math.round(hang)}px;--ind:${Math.round(hang)}px;`;
      } else {
        if (marL) css += `margin-left:${Math.round(marL)}px;`;
        if (ind) css += `text-indent:${Math.round(ind)}px;`;
      }
      const ptsOf = (e) => { if (!e) return null; const pts = kid(e, 'spcPts'); return pts ? num(pts.getAttribute('val'), 0) / 100 : null; };
      const bef = ptsOf(spcBef), aft = ptsOf(spcAft);
      if (bef) css += `margin-top:${bef}pt;`;
      css += `margin-bottom:${aft || 0}pt;`;
      if (lnSpc) { const pct = kid(lnSpc, 'spcPct'); if (pct) css += `line-height:${num(pct.getAttribute('val'), 100000) / 100000};`; }
      let inner = '';
      const defSz = num(eff.def.sz, 1800) / 100 * scale;
      for (const r of p.children) {
        if (r.localName === 'br') { inner += '<br>'; continue; }
        if (r.localName !== 'r' && r.localName !== 'fld') continue;
        const t = kid(r, 't');
        const txt = t ? t.textContent : '';
        const rPr = kid(r, 'rPr');
        const ra = {};
        if (rPr) for (const a of rPr.attributes) ra[a.name] = a.value;
        let rcss = '';
        const sz = ra.sz ? num(ra.sz, 1800) / 100 * scale : defSz;
        rcss += `font-size:${Math.round(sz * 2) / 2}pt;`;
        const b = ra.b !== undefined ? ra.b === '1' || ra.b === 'true' : (eff.def.b === '1');
        const i = ra.i !== undefined ? ra.i === '1' || ra.i === 'true' : (eff.def.i === '1');
        if (b) rcss += 'font-weight:bold;';
        if (i) rcss += 'font-style:italic;';
        const u = ra.u !== undefined ? ra.u !== 'none' : (eff.def.u && eff.def.u !== 'none');
        const st = ra.strike !== undefined ? ra.strike !== 'noStrike' : false;
        if (u || st) rcss += `text-decoration:${u ? 'underline' : ''} ${st ? 'line-through' : ''};`;
        const bl = num(ra.baseline, 0);
        if (bl > 0) rcss += 'vertical-align:super;'; else if (bl < 0) rcss += 'vertical-align:sub;';
        let col = null;
        const sf = rPr && kid(rPr, 'solidFill');
        if (sf) col = resolveColor(sf, c.T);
        else if (eff.def._fill) col = resolveColor(eff.def._fill, c.T);
        rcss += `color:#${col || o.bodyColor || '000000'};`;
        let font = null;
        const lat = rPr && kid(rPr, 'latin');
        font = lat ? lat.getAttribute('typeface') : eff.def._font;
        if (font) font = font === '+mj-lt' ? c.T.major : (font === '+mn-lt' ? c.T.minor : font);
        else font = (o.kind === 'title') ? c.T.major : c.T.minor;
        rcss += `font-family:'${font}';`;
        const href = rPr && kid(rPr, 'hlinkClick');
        let span = `<span style="${rcss}">${escH(txt)}</span>`;
        if (href) { const rid = href.getAttributeNS(NS.r, 'id') || href.getAttribute('r:id'); const rel = rid && c.slideRels && c.slideRels[rid]; if (rel && rel.target && /^https?:/.test(rel.target)) span = `<a href="${escH(rel.target)}">${span}</a>`; }
        inner += span;
      }
      if (!inner) {
        const e = kid(p, 'endParaRPr');
        const sz = e && e.getAttribute('sz') ? num(e.getAttribute('sz'), 1800) / 100 * scale : defSz;
        inner = `<br>`;
        css += `font-size:${sz}pt;`;
      }
      html += `<p${pAttr} style="${css}">${inner}</p>`;
    }
    return html || '<p><br></p>';
  }

  g.PptxIO = { write, read, PPTX_MIME, LAYOUTS };
})(window);
