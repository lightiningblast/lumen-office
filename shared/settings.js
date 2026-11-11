/* Lumen Office settings: one panel for Write, Cells and Slides (gear button in the title row, or Ctrl+,).
   Loaded in <head> so the theme is applied before the page is drawn. Stored per user in this browser profile. */
(function () {
  'use strict';
  const KEY = 'lumen-settings';
  const DEFAULTS = {
    theme: 'system',        // system | light | dark
    accent: 'blue',
    uiSize: 'medium',       // small | medium | large
    toolbar: 'show',        // show | hide
    labels: false,          // names under toolbar icons
    statusbar: true,
    motion: true,
    pageView: 'pages',      // Write: pages | continuous
    spellcheck: true,
    author: '',
    autoUpdate: true,
  };
  const ACCENTS = {
    blue: ['Blue', '#0a72e8'], purple: ['Purple', '#8e4ec6'], pink: ['Pink', '#e2457f'], red: ['Red', '#e5484d'],
    orange: ['Orange', '#f2811d'], yellow: ['Yellow', '#d9a400'], green: ['Green', '#2f9e55'], graphite: ['Graphite', '#7c7c82'],
  };

  let S = Object.assign({}, DEFAULTS);
  try { Object.assign(S, JSON.parse(localStorage.getItem(KEY) || '{}')); } catch (e) { }
  try { const a = localStorage.getItem('lumen-author'); if (a && !S.author) S.author = a; } catch (e) { }
  try { if (localStorage.getItem('lumen-tools-hidden') === '1') S.toolbar = 'hide'; } catch (e) { }

  const root = document.documentElement;
  const host = window.chrome && window.chrome.webview;

  function apply() {
    root.dataset.theme = S.theme;
    const ac = ACCENTS[S.accent] || ACCENTS.blue;
    root.style.setProperty('--sys-blue', ac[1]);
    root.style.setProperty('--sys-blue-soft', ac[1] + (S.theme === 'dark' ? '33' : '1f'));
    root.dataset.ui = S.uiSize;
    root.classList.toggle('no-motion', !S.motion);
    const body = document.body;
    if (body) {
      body.classList.toggle('tools-hidden', S.toolbar === 'hide');
      body.classList.toggle('show-labels', !!S.labels);
      body.classList.toggle('no-statusbar', !S.statusbar);
      document.querySelectorAll('#doc, #hdr, #ftr, #notes, .tc').forEach(el => el.setAttribute('spellcheck', S.spellcheck ? 'true' : 'false'));
    }
    try { localStorage.setItem('lumen-tools-hidden', S.toolbar === 'hide' ? '1' : '0'); } catch (e) { }
    try { localStorage.setItem('lumen-author', S.author || ''); } catch (e) { }
    // the Lumen program windows follow the theme too (title bar and every open window)
    try { if (host) host.postMessage(JSON.stringify({ type: 'theme', value: S.theme })); } catch (e) { }
    window.dispatchEvent(new CustomEvent('lumen-settings', { detail: Object.assign({}, S) }));
  }
  function save() { try { localStorage.setItem(KEY, JSON.stringify(S)); } catch (e) { } }
  function set(k, v) { S[k] = v; save(); apply(); render(); }
  function get(k) { return S[k]; }

  // other Lumen windows change settings too: follow them
  window.addEventListener('storage', e => {
    if (e.key !== KEY) return;
    try { S = Object.assign({}, DEFAULTS, JSON.parse(e.newValue || '{}')); } catch (x) { }
    apply(); render();
  });

  /* ---------------- the panel ---------------- */
  const ICON = '<svg class="sf" viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="3"/><path d="M12 2.8v2.4M12 18.8v2.4M21.2 12h-2.4M5.2 12H2.8M18.5 5.5l-1.7 1.7M7.2 16.8l-1.7 1.7M18.5 18.5l-1.7-1.7M7.2 7.2 5.5 5.5"/><circle cx="12" cy="12" r="6.2"/></svg>';
  let panel = null, btn = null;
  const app = location.pathname.includes('/write/') ? 'write' : location.pathname.includes('/cells/') ? 'cells' : location.pathname.includes('/slides/') ? 'slides' : 'office';

  function seg(key, opts) {
    return '<div class="ls-seg" role="radiogroup">' + opts.map(([v, l]) => '<button type="button" role="radio" data-k="' + key + '" data-v="' + v + '" aria-checked="' + (String(S[key]) === v) + '"' + (String(S[key]) === v ? ' class="on"' : '') + '>' + l + '</button>').join('') + '</div>';
  }
  function sw(key) { const v = key === 'toolbarOn' ? S.toolbar !== 'hide' : !!S[key]; return '<button type="button" class="ls-switch' + (v ? ' on' : '') + '" role="switch" aria-checked="' + v + '" data-k="' + key + '" data-toggle="1"><i></i></button>'; }
  function row(label, ctl, hint) { return '<div class="ls-row"><div class="ls-l">' + label + (hint ? '<small>' + hint + '</small>' : '') + '</div>' + ctl + '</div>'; }
  const esc = (s) => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);

  function render() {
    if (!panel || panel.hidden) return;
    const ver = panel.dataset.ver || '';
    let h = '<div class="ls-head"><b>Settings</b><button type="button" class="ls-x" data-close="1" aria-label="Close">&times;</button></div><div class="ls-body">';
    h += '<div class="ls-sec">Appearance</div>';
    h += row('Theme', seg('theme', [['system', 'Auto'], ['light', 'Light'], ['dark', 'Dark']]), 'Auto follows Windows');
    h += row('Accent colour', '<div class="ls-dots">' + Object.keys(ACCENTS).map(k => '<button type="button" class="ls-dot' + (S.accent === k ? ' on' : '') + '" title="' + ACCENTS[k][0] + '" aria-label="' + ACCENTS[k][0] + '" data-k="accent" data-v="' + k + '" style="--c:' + ACCENTS[k][1] + '"></button>').join('') + '</div>');
    h += row('Interface size', seg('uiSize', [['small', 'Small'], ['medium', 'Medium'], ['large', 'Large']]));
    h += row('Animations', sw('motion'));
    h += '<div class="ls-sec">Toolbar</div>';
    h += row('Show tools', sw('toolbarOn'), 'Or click the current tab');
    h += row('Names under icons', sw('labels'));
    h += row('Status bar', sw('statusbar'));
    if (app === 'write') {
      h += '<div class="ls-sec">Documents</div>';
      h += row('Page view', seg('pageView', [['pages', 'Pages'], ['continuous', 'Continuous']]), 'Pages shows each page separately');
    }
    if (app !== 'office') h += row('Check spelling', sw('spellcheck'));
    h += row('Your name', '<input class="ls-in" data-k="author" value="' + esc(S.author) + '" placeholder="For comments and changes" spellcheck="false">');
    h += '<div class="ls-sec">Updates</div>';
    h += row('Check automatically', sw('autoUpdate'));
    h += '<div class="ls-row"><div class="ls-l">Version ' + esc(ver) + '</div><button type="button" class="ls-btn" data-act="update">Check now</button></div>';
    h += '<div class="ls-foot"><button type="button" class="ls-link" data-act="reset">Reset all settings</button></div></div>';
    panel.innerHTML = h;
  }

  function open() {
    if (!panel) build();
    panel.hidden = false;
    btn && btn.classList.add('on');
    render();
    const f = panel.querySelector('.ls-seg .on, .ls-switch'); if (f) f.focus();
  }
  function close() { if (panel) panel.hidden = true; btn && btn.classList.remove('on'); }
  function toggle() { if (panel && !panel.hidden) close(); else open(); }

  function build() {
    panel = document.createElement('div');
    panel.id = 'lumen-settings'; panel.hidden = true; panel.setAttribute('role', 'dialog'); panel.setAttribute('aria-label', 'Settings');
    document.body.appendChild(panel);
    fetch((app === 'office' ? '' : '../') + 'version.json', { cache: 'no-store' }).then(r => r.json()).then(v => { panel.dataset.ver = v.version; render(); }).catch(() => { });
    panel.addEventListener('click', e => {
      const t = e.target.closest('button'); if (!t) return;
      if (t.dataset.close) return close();
      if (t.dataset.act === 'update') { if (window.Lumen && Lumen.checkUpdates) Lumen.checkUpdates(true); return; }
      if (t.dataset.act === 'reset') { S = Object.assign({}, DEFAULTS, { author: S.author }); save(); apply(); render(); return; }
      const k = t.dataset.k; if (!k) return;
      if (k === 'toolbarOn') return set('toolbar', S.toolbar === 'hide' ? 'show' : 'hide');
      if (t.dataset.toggle) return set(k, !S[k]);
      set(k, t.dataset.v);
    });
    panel.addEventListener('change', e => { if (e.target.dataset.k === 'author') set('author', e.target.value.trim()); });
    document.addEventListener('mousedown', e => { if (panel && !panel.hidden && !panel.contains(e.target) && !(btn && btn.contains(e.target))) close(); });
    document.addEventListener('keydown', e => { if (e.key === 'Escape' && panel && !panel.hidden) { e.stopPropagation(); close(); } }, true);
  }

  function init() {
    apply();
    const bar = document.querySelector('header.appbar');
    if (bar) {
      btn = document.createElement('button');
      btn.type = 'button'; btn.className = 'btn-text icon-only ls-gear'; btn.title = 'Settings (Ctrl+,)'; btn.setAttribute('aria-label', 'Settings');
      btn.innerHTML = ICON;
      btn.addEventListener('click', toggle);
      bar.appendChild(btn);
    }
    document.addEventListener('keydown', e => { if ((e.ctrlKey || e.metaKey) && e.key === ',') { e.preventDefault(); toggle(); } });
    // the tab toggle (click the current tab) writes the same setting
    window.addEventListener('click', () => {
      const hidden = document.body.classList.contains('tools-hidden') ? 'hide' : 'show';
      if (hidden !== S.toolbar) { S.toolbar = hidden; save(); render(); }
    });
  }

  apply();  // theme before first paint
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
  window.LumenSettings = { get, set, open, close, defaults: DEFAULTS };
})();
