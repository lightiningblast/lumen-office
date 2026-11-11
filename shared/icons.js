/* Lumen Office icon set: thin, rounded line icons in the spirit of Apple's SF Symbols.
   Replaces the emoji / symbol-character icons of the toolbars when the page loads. */
(function () {
  'use strict';
  const P = {
    undo: '<path d="M9 14 4 9l5-5"/><path d="M4 9h10.5a5.5 5.5 0 0 1 0 11H11"/>',
    redo: '<path d="m15 14 5-5-5-5"/><path d="M20 9H9.5a5.5 5.5 0 0 0 0 11H13"/>',
    cut: '<circle cx="6" cy="18" r="2.5"/><circle cx="18" cy="18" r="2.5"/><path d="M7.8 16.2 18 4M16.2 16.2 6 4"/>',
    copy: '<rect x="8" y="8" width="12" height="12" rx="2.5"/><path d="M16 8V6.5A2.5 2.5 0 0 0 13.5 4h-7A2.5 2.5 0 0 0 4 6.5v7A2.5 2.5 0 0 0 6.5 16H8"/>',
    eraser: '<path d="M8 20h12"/><path d="m4.6 15.4 9.8-9.8a2 2 0 0 1 2.8 0l2.2 2.2a2 2 0 0 1 0 2.8L11 19H7.4z"/><path d="m9 11 5 5"/>',
    search: '<circle cx="11" cy="11" r="6.5"/><path d="m20 20-4.4-4.4"/>',
    selectall: '<rect x="4" y="4" width="16" height="16" rx="3" stroke-dasharray="3 2.6"/><path d="M8 9.5h8M8 12.5h8M8 15.5h5"/>',
    table: '<rect x="3.5" y="4.5" width="17" height="15" rx="2.5"/><path d="M3.5 9.5h17M3.5 14.5h17M9.5 9.5v10M15 9.5v10"/>',
    photo: '<rect x="3.5" y="4.5" width="17" height="15" rx="2.5"/><circle cx="9" cy="10" r="1.6"/><path d="m4 17.5 5-4.5 4 3.5 3-2.5 4 3.5"/>',
    link: '<path d="M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1"/><path d="M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1"/>',
    unlink: '<path d="M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1"/><path d="M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1"/><path d="M4 4l16 16"/>',
    list: '<path d="M9 6h11M9 12h11M9 18h11"/><circle cx="4.5" cy="6" r=".9"/><circle cx="4.5" cy="12" r=".9"/><circle cx="4.5" cy="18" r=".9"/>',
    refresh: '<path d="M20 12a8 8 0 1 1-2.3-5.6"/><path d="M20 4v5h-5"/>',
    footnote: '<path d="M4 18h10M4 13h10M4 8h7"/><path d="M17 4.5h1.5v5.5M17 10h3"/>',
    comment: '<path d="M20 12a7.5 7.5 0 0 1-11 6.6L4 20l1.4-4.4A7.5 7.5 0 1 1 20 12z"/>',
    header: '<rect x="4" y="3.5" width="16" height="17" rx="2.5"/><path d="M4 8.5h16M8 6h4"/>',
    footer: '<rect x="4" y="3.5" width="16" height="17" rx="2.5"/><path d="M4 15.5h16M8 18h4"/>',
    number: '<path d="M10 4 8 20M16 4l-2 16M5 9h15M4 15h15"/>',
    sigma: '<path d="M17 5H7l6 7-6 7h10"/>',
    line: '<path d="M4 12h16"/>',
    pagebreak: '<path d="M6 3v5h12V3M6 21v-5h12v5"/><path d="M3 12h2.5M8.5 12h2M13.5 12h2M18.5 12H21"/>',
    calendar: '<rect x="4" y="5" width="16" height="15" rx="2.5"/><path d="M4 10h16M8 3v4M16 3v4"/>',
    omega: '<path d="M5 19h4v-2a7 7 0 1 1 6 0v2h4"/>',
    textbox: '<rect x="3.5" y="5" width="17" height="14" rx="2.5"/><path d="M8 9h8M12 9v7"/>',
    equation: '<path d="M14 5h-1a3 3 0 0 0-3 3v9a3 3 0 0 1-3 3H6M7 11h6.5"/><path d="m15 13 5 5M20 13l-5 5"/>',
    endnote: '<path d="M4 6h16M4 10h16M4 14h9"/><path d="M16 14.5h1.5V20M16 20h3"/>',
    quote: '<path d="M5 17v-5a5 5 0 0 1 4-4.9M5 13h4v4H5zM14 17v-5a5 5 0 0 1 4-4.9M14 13h4v4h-4z"/>',
    book: '<path d="M12 6.5C10 5 7 4.5 4 5v13c3-.5 6 0 8 1.5 2-1.5 5-2 8-1.5V5c-3-.5-6 0-8 1.5zM12 6.5V19"/>',
    bookmark: '<path d="M7 4h10a1 1 0 0 1 1 1v15l-6-4-6 4V5a1 1 0 0 1 1-1z"/>',
    arrowupright: '<path d="M7 17 17 7M9 7h8v8"/>',
    caption: '<rect x="4" y="4" width="16" height="11" rx="2.5"/><path d="M7 19h10"/>',
    people: '<circle cx="9" cy="8.5" r="3.2"/><path d="M3.5 19a5.5 5.5 0 0 1 11 0"/><circle cx="17" cy="9" r="2.5"/><path d="M16 14.2a4.5 4.5 0 0 1 5 4.8"/>',
    field: '<path d="m10 7-5 5 5 5M19 7l-5 5 5 5"/>',
    eye: '<path d="M2.5 12S6 5.5 12 5.5 21.5 12 21.5 12 18 18.5 12 18.5 2.5 12 2.5 12z"/><circle cx="12" cy="12" r="3"/>',
    check: '<path d="m5 12.5 4.5 4.5L19 7.5"/>',
    checkall: '<path d="m2.5 12.5 4 4L14.5 8.5M10 16l1 1 10-10"/>',
    xmark: '<path d="M6.5 6.5l11 11M17.5 6.5l-11 11"/>',
    xall: '<path d="M3.5 7.5l8 8M11.5 7.5l-8 8M13.5 7.5l7 7M20.5 7.5l-7 7"/>',
    pencil: '<path d="M4 20h4L19 9a2.8 2.8 0 0 0-4-4L4 16z"/><path d="m13.5 6.5 4 4"/>',
    chevleft: '<path d="m15 5-7 7 7 7"/>',
    chevright: '<path d="m9 5 7 7-7 7"/>',
    textsize: '<path d="M4 7V5h16v2M12 5v14M9 19h6"/>',
    orientation: '<rect x="8" y="3.5" width="10" height="13" rx="2"/><path d="M4 12.5a8 8 0 0 0 8 7.5M4 12.5V16M4 12.5h3.5"/>',
    minus: '<path d="M5 12h14"/>',
    plus: '<path d="M12 5v14M5 12h14"/>',
    actual: '<circle cx="11" cy="11" r="6.5"/><path d="m20 20-4.4-4.4M8.5 9.8h5M8.5 12.2h5"/>',
    fitwidth: '<path d="M4 6v12M20 6v12M7 12h10M9.5 9.5 7 12l2.5 2.5M14.5 9.5 17 12l-2.5 2.5"/>',
    pivot: '<rect x="3.5" y="3.5" width="17" height="17" rx="2.5"/><path d="M3.5 9h17M9 3.5v17M13 13.5h4M15 11.5v4"/>',
    sortaz: '<path d="M7 4v16M4 17l3 3 3-3"/><path d="M14 10l2.5-6 2.5 6M14.8 8h3.4M14 14h5l-5 6h5"/>',
    sortza: '<path d="M7 4v16M4 17l3 3 3-3"/><path d="M14 4h5l-5 6h5M14 20l2.5-6 2.5 6M14.8 18h3.4"/>',
    filter: '<path d="M4 5h16l-6 7.5V19l-4-2v-4.5z"/>',
    validate: '<rect x="4" y="4" width="16" height="16" rx="3.5"/><path d="m8 12.5 3 3 5-6"/>',
    unvalidate: '<rect x="4" y="4" width="16" height="16" rx="3.5"/><path d="m9 9 6 6M15 9l-6 6"/>',
    dedupe: '<rect x="8" y="8" width="12" height="12" rx="2.5"/><path d="M16 8V6.5A2.5 2.5 0 0 0 13.5 4h-7A2.5 2.5 0 0 0 4 6.5v7A2.5 2.5 0 0 0 6.5 16H8"/><path d="M11.5 14h5"/>',
    columns: '<rect x="3.5" y="5" width="17" height="14" rx="2.5"/><path d="M12 5v14M6.5 9.5h3M6.5 12.5h3M14.5 9.5h3M14.5 12.5h3"/>',
    target: '<circle cx="12" cy="12" r="8"/><circle cx="12" cy="12" r="4.5"/><circle cx="12" cy="12" r=".8"/>',
    sliders: '<path d="M4 7h9M17 7h3M4 17h3M11 17h9"/><circle cx="15" cy="7" r="2"/><circle cx="9" cy="17" r="2"/>',
    tag: '<path d="M4 4h7.5l8.5 8.5-7.5 7.5L4 11.5z"/><circle cx="8.5" cy="8.5" r="1.3"/>',
    trash: '<path d="M5 7h14M10 7V5h4v2M7 7l1 12.5h8L17 7"/>',
    swap: '<path d="M4 8h13l-3-3M20 16H7l3 3"/>',
    chart: '<path d="M4 20h16"/><rect x="5.5" y="11" width="3" height="7" rx="1"/><rect x="10.5" y="6" width="3" height="12" rx="1"/><rect x="15.5" y="13" width="3" height="5" rx="1"/>',
    video: '<rect x="3.5" y="6" width="12" height="12" rx="2.5"/><path d="m15.5 10 5-3v10l-5-3z"/>',
    audio: '<path d="M4 9.5v5h3.5L12 18V6L7.5 9.5z"/><path d="M15.5 9a4 4 0 0 1 0 6M18 6.5a7.5 7.5 0 0 1 0 11"/>',
    play: '<path d="M8 5.5v13l10.5-6.5z"/>',
    playfrom: '<path d="M4.5 5v14"/><path d="M9 5.5v13l10.5-6.5z"/>',
    presenter: '<rect x="3.5" y="4" width="17" height="11.5" rx="2"/><path d="M12 15.5V20M8 20h8M7 8h6M7 11h4"/>',
    master: '<rect x="3.5" y="4.5" width="17" height="15" rx="2.5"/><rect x="7" y="8" width="10" height="8" rx="1.5"/>',
    alignbottom: '<path d="M4 20h16M12 4v12M8 12l4 4 4-4"/>',
    aligncenter: '<path d="M4 12h16M12 3v6M9 6l3 3 3-3M12 21v-6M9 18l3-3 3 3"/>',
    aligntop: '<path d="M4 4h16M12 20V8M8 12l4-4 4 4"/>',
    nofill: '<circle cx="12" cy="12" r="8"/><path d="M6.5 17.5 17.5 6.5"/>',
    newdoc: '<path d="M14 3.5H7.5A2.5 2.5 0 0 0 5 6v12a2.5 2.5 0 0 0 2.5 2.5h9A2.5 2.5 0 0 0 19 18V8.5z"/><path d="M14 3.5v5h5M12 11.5v6M9 14.5h6"/>',
    folder: '<path d="M3.5 7.5a2 2 0 0 1 2-2H9l2 2h7.5a2 2 0 0 1 2 2V18a2 2 0 0 1-2 2h-13a2 2 0 0 1-2-2z"/>',
    save: '<path d="M12 4v11M7.5 10.5 12 15l4.5-4.5"/><path d="M5 15v3.5A1.5 1.5 0 0 0 6.5 20h11a1.5 1.5 0 0 0 1.5-1.5V15"/>',
    saveas: '<path d="M11 4v9M7 9.5l4 4 4-4"/><path d="M5 15v3.5A1.5 1.5 0 0 0 6.5 20H12"/><path d="m15.5 21 .5-2.5 3.5-3.5 2 2-3.5 3.5z"/>',
    print: '<path d="M7 9V4h10v5"/><rect x="3.5" y="9" width="17" height="8" rx="2"/><path d="M7 14h10v6H7z"/>',
    grid: '<rect x="4" y="4" width="7" height="7" rx="2"/><rect x="13" y="4" width="7" height="7" rx="2"/><rect x="4" y="13" width="7" height="7" rx="2"/><rect x="13" y="13" width="7" height="7" rx="2"/>',
    export: '<path d="M12 15V4M7.5 8.5 12 4l4.5 4.5"/><path d="M5 13v5.5A1.5 1.5 0 0 0 6.5 20h11a1.5 1.5 0 0 0 1.5-1.5V13"/>',
    numbers: '<path d="M10 6h10M10 12h10M10 18h10"/><path d="M4 4.5h1.5V9M4 9h3M4 14.5h2.5a1 1 0 0 1 0 2L4 19.5h3"/>',
  };
  const svg = (n) => `<svg class="sf" viewBox="0 0 24 24" aria-hidden="true">${P[n]}</svg>`;

  // which icon each control gets (by id)
  const BY_ID = {
    'btn-find': 'search', 'btn-selectall': 'selectall', 'ins-table': 'table', 'ins-image': 'photo', 'ins-link': 'link', 'ins-unlink': 'unlink',
    'ins-toc': 'list', 'upd-toc': 'refresh', 'ins-fn': 'footnote', 'ins-cmt': 'comment', 'ins-header': 'header', 'ins-footer': 'footer',
    'ins-pgnum': 'number', 'ins-pgtotal': 'sigma', 'ins-hr': 'line', 'ins-pb': 'pagebreak', 'ins-date': 'calendar', 'ins-symbol': 'omega',
    'ins-tbox': 'textbox', 'ins-eq': 'equation', 'ins-en': 'endnote', 'cite-ins': 'quote', 'cite-src': 'book', 'cite-bib': 'list',
    'ins-bm': 'bookmark', 'ins-xref': 'arrowupright', 'ins-caption': 'caption', 'mm-list': 'people', 'mm-field': 'field', 'mm-preview': 'eye',
    'mm-finish': 'check', 'rv-track': 'pencil', 'rv-prev': 'chevleft', 'rv-next': 'chevright', 'rv-accept': 'check', 'rv-reject': 'xmark',
    'rv-accept-all': 'checkall', 'rv-reject-all': 'xall', 'rv-cmt': 'comment', 'rv-count': 'textsize', 'lay-orient': 'orientation',
    'zoom-out': 'minus', 'zoom-100': 'actual', 'zoom-in': 'plus', 'zoom-fit': 'fitwidth',
    'pv-make': 'pivot', 'pv-refresh': 'refresh', 'sort-az': 'sortaz', 'sort-za': 'sortza', 'flt-toggle': 'filter', 'flt-reapply': 'refresh',
    'dv-edit': 'validate', 'dv-clear': 'unvalidate', 'dd-rm': 'dedupe', 't2c': 'columns', 'gs-run': 'target', 'solver-run': 'sliders',
    'nm-edit': 'tag', 'cm-edit': 'comment', 'cm-del': 'trash', 'd-replace': 'swap', 'b-copy': 'copy', 'b-cut': 'cut', 'b-nofill': 'nofill',
    'b-undo': 'undo', 'b-redo': 'redo', 'b-numbers': 'numbers',
    'i-text': 'textbox', 'i-image': 'photo', 'i-table': 'table', 'i-chart': 'chart', 'i-video': 'video', 'i-audio': 'audio', 'i-hf': 'number',
    'tr-preview': 'play', 'v-show': 'play', 'v-showcur': 'playfrom', 'v-presenter': 'presenter', 'v-fit': 'fitwidth', 'v-master': 'master',
    'btn-new': 'newdoc', 'btn-open': 'folder', 'btn-save': 'save', 'btn-saveas': 'saveas', 'btn-print': 'print', 'btn-show': 'play', 'btn-csv': 'export',
  };
  const BY_ATTR = [
    ['[data-cmd="undo"]', 'undo'], ['[data-cmd="redo"]', 'redo'], ['[data-cmd="cut"]', 'cut'], ['[data-cmd="copy"]', 'copy'],
    ['[data-cmd="removeFormat"]', 'eraser'], ['[data-v="bottom"]', 'alignbottom'], ['[data-v="center"]', 'aligncenter'], ['[data-v="top"]', 'aligntop'],
    ['[data-va="b"]', 'alignbottom'], ['[data-va="ctr"]', 'aligncenter'], ['[data-va="t"]', 'aligntop'],
  ];

  function apply(el, name) {
    if (!el || !P[name] || el.dataset.sf) return;
    el.dataset.sf = name;
    const ic = el.querySelector('.ic');
    if (ic) {
      // the visible label is hidden in the compact toolbar, so it becomes the tooltip
      const label = [...el.childNodes].filter(n => n !== ic).map(n => n.textContent).join('').trim();
      if (label) { if (!el.title) el.title = label; el.setAttribute('aria-label', label); }
      ic.innerHTML = svg(name);
      return;
    }
    if (el.classList.contains('btn-text')) {
      // toolbar buttons: icon only, the text becomes the tooltip
      const label = el.textContent.replace(/^[\s▶]+/, '').trim();
      if (!el.title) el.title = label;
      el.setAttribute('aria-label', label);
      el.innerHTML = svg(name);
      el.classList.add('icon-only');
      return;
    }
    if (!el.getAttribute('aria-label') && el.title) el.setAttribute('aria-label', el.title);
    el.innerHTML = svg(name);
  }

  function run() {
    for (const id in BY_ID) apply(document.getElementById(id), BY_ID[id]);
    for (const [sel, name] of BY_ATTR) document.querySelectorAll('button' + sel).forEach(b => apply(b, name));
    const home = document.querySelector('.appbar a.home');
    if (home && !home.dataset.sf) { home.dataset.sf = 'grid'; home.title = 'Office Apps'; home.setAttribute('aria-label', 'Office Apps'); home.innerHTML = svg('grid'); home.classList.add('icon-only'); }
    compact();
    document.documentElement.classList.add('sf-ready');
  }

  // compact layout: tabs move into the title row; clicking the active tab hides / shows the tools
  function compact() {
    const bar = document.querySelector('header.appbar'), tabs = document.querySelector('body > nav.tabs');
    if (bar && tabs) {
      const sp = bar.querySelector('.spacer');
      const sp2 = document.createElement('div'); sp2.className = 'spacer2';
      if (sp) { sp.after(tabs); tabs.after(sp2); sp.className = 'spacer2'; } else bar.appendChild(tabs);
    }
    let hidden = false;
    try { hidden = localStorage.getItem('lumen-tools-hidden') === '1'; } catch (e) { }
    document.body.classList.toggle('tools-hidden', hidden);
    document.querySelectorAll('.tab').forEach(t => t.addEventListener('mousedown', () => {
      const wasActive = t.classList.contains('active');
      // clicking the shown tab hides the tools; clicking any tab while hidden shows them again
      const hide = wasActive && !document.body.classList.contains('tools-hidden');
      document.body.classList.toggle('tools-hidden', hide);
      try { localStorage.setItem('lumen-tools-hidden', hide ? '1' : '0'); } catch (e) { }
    }));
  }
  window.LumenIcons = { svg, names: () => Object.keys(P), apply };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', run); else run();
})();
