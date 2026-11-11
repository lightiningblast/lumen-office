/* Shared helpers for the Lumen Office apps. */
(function (global) {
  'use strict';
  const $ = (s, r) => (r || document).querySelector(s);
  const $$ = (s, r) => [...(r || document).querySelectorAll(s)];

  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  // Escape for XML text/attributes and drop characters XML 1.0 forbids.
  function xesc(s) {
    return String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  const _cv = document.createElement('canvas').getContext('2d');
  // Any CSS colour -> 'RRGGBB' (no #) or null when transparent / invalid.
  function cssToHex(str) {
    if (!str) return null;
    str = String(str).trim().toLowerCase();
    if (!str || str === 'transparent' || str === 'inherit' || str === 'initial' || str === 'currentcolor') return null;
    let m = /^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+%?))?\s*\)$/.exec(str);
    if (m) {
      if (m[4] !== undefined && parseFloat(m[4]) === 0) return null;
      return [m[1], m[2], m[3]].map(n => ('0' + Math.min(255, +n).toString(16)).slice(-2)).join('').toUpperCase();
    }
    m = /^#([0-9a-f]{6})$/.exec(str);
    if (m) return m[1].toUpperCase();
    m = /^#([0-9a-f]{3})$/.exec(str);
    if (m) return m[1].split('').map(c => c + c).join('').toUpperCase();
    _cv.fillStyle = '#010203';
    _cv.fillStyle = str;
    const v = _cv.fillStyle;
    if (v === '#010203' && str !== '#010203') return null;
    m = /^#([0-9a-f]{6})$/.exec(v);
    return m ? m[1].toUpperCase() : null;
  }

  function dataUrlToBytes(url) {
    const m = /^data:([^;,]*)(;base64)?,(.*)$/s.exec(url);
    if (!m) return null;
    let bin;
    if (m[2]) bin = atob(m[3]); else bin = decodeURIComponent(m[3]);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return { mime: m[1] || 'application/octet-stream', bytes: out };
  }
  function bytesToDataUrl(bytes, mime) {
    let s = '';
    const CH = 0x8000;
    for (let i = 0; i < bytes.length; i += CH) s += String.fromCharCode.apply(null, bytes.subarray(i, i + CH));
    return 'data:' + mime + ';base64,' + btoa(s);
  }
  const MIME_BY_EXT = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', tif: 'image/tiff', tiff: 'image/tiff' };
  const EXT_BY_MIME = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/svg+xml': 'svg', 'image/webp': 'webp' };

  function readFileAsDataUrl(file) {
    return new Promise((res, rej) => {
      const r = new FileReader();
      r.onload = () => res(r.result);
      r.onerror = () => rej(r.error);
      r.readAsDataURL(file);
    });
  }

  /* ---------- toast ---------- */
  let toastTimer;
  function toast(msg, ms) {
    let t = $('#toast');
    if (!t) { t = document.createElement('div'); t.id = 'toast'; document.body.appendChild(t); }
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove('show'), ms || 2600);
  }

  /* ---------- modal form ---------- */
  function form(title, fields, okLabel) {
    return new Promise(resolve => {
      const back = document.createElement('div');
      back.className = 'modal-back';
      const box = document.createElement('div');
      box.className = 'modal';
      let html = '<h3>' + esc(title) + '</h3>';
      for (const f of fields) {
        if (f.type === 'note') { html += '<p class="note">' + esc(f.label) + '</p>'; continue; }
        html += '<label class="fld"><span>' + esc(f.label) + '</span>';
        if (f.type === 'select') {
          html += '<select data-id="' + f.id + '">' + f.options.map(o => {
            const v = Array.isArray(o) ? o[0] : o, l = Array.isArray(o) ? o[1] : o;
            return '<option value="' + esc(v) + '"' + (String(v) === String(f.value) ? ' selected' : '') + '>' + esc(l) + '</option>';
          }).join('') + '</select>';
        } else if (f.type === 'textarea') {
          html += '<textarea data-id="' + f.id + '" rows="' + (f.rows || 6) + '" style="width:260px;font:12px Consolas,monospace;border:1px solid #bbb;border-radius:4px;padding:4px">' + esc(f.value == null ? '' : f.value) + '</textarea>';
        } else if (f.type === 'checkbox') {
          html += '<input type="checkbox" data-id="' + f.id + '"' + (f.value ? ' checked' : '') + '>';
        } else {
          html += '<input data-id="' + f.id + '" type="' + (f.type || 'text') + '" value="' + esc(f.value == null ? '' : f.value) + '"' + (f.min != null ? ' min="' + f.min + '"' : '') + (f.max != null ? ' max="' + f.max + '"' : '') + '>';
        }
        html += '</label>';
      }
      html += '<div class="actions"><button class="btn-text" data-act="cancel">Cancel</button><button class="btn-text primary" data-act="ok">' + esc(okLabel || 'OK') + '</button></div>';
      box.innerHTML = html;
      back.appendChild(box);
      document.body.appendChild(back);
      box.style.maxWidth = '520px';
      const first = box.querySelector('input,select,textarea');
      if (first) { first.focus(); if (first.select) first.select(); }
      const done = (ok) => {
        let out = null;
        if (ok) {
          out = {};
          $$('[data-id]', box).forEach(el => {
            out[el.dataset.id] = el.type === 'checkbox' ? el.checked : (el.type === 'number' ? parseFloat(el.value) : el.value);
          });
        }
        back.remove();
        resolve(out);
      };
      box.addEventListener('click', e => {
        const a = e.target.dataset && e.target.dataset.act;
        if (a === 'ok') done(true); else if (a === 'cancel') done(false);
      });
      box.addEventListener('keydown', e => {
        if (e.key === 'Enter' && e.target.tagName !== 'BUTTON' && e.target.tagName !== 'TEXTAREA') { e.preventDefault(); done(true); }
        if (e.key === 'Escape') { e.preventDefault(); done(false); }
      });
      back.addEventListener('mousedown', e => { if (e.target === back) done(false); });
    });
  }
  function confirmBox(title, message, okLabel) {
    return form(title, [{ type: 'note', label: message }], okLabel || 'OK').then(r => !!r);
  }

  /* ---------- file open / save ---------- */
  const FileIO = {
    async open(types, legacyAccept) {
      if (window.showOpenFilePicker) {
        try {
          const [h] = await window.showOpenFilePicker({ types, multiple: false });
          return { file: await h.getFile(), handle: h };
        } catch (e) {
          if (e && e.name === 'AbortError') return null;
        }
      }
      return new Promise(res => {
        const i = document.createElement('input');
        i.type = 'file';
        i.accept = legacyAccept || '';
        i.onchange = () => res(i.files[0] ? { file: i.files[0], handle: null } : null);
        i.click();
      });
    },
    async save(blob, name, types, handle, saveAs) {
      // a file opened from Windows Explorer: Ctrl+S writes it back through the local Lumen server
      if (handle && handle.lumenToken && !saveAs) {
        const r = await fetch('/_file/' + handle.lumenToken, { method: 'PUT', headers: { 'X-Lumen': '1' }, body: blob });
        if (!r.ok) { let m = 'Save failed'; try { m = (await r.json()).error || m; } catch (e) { } throw new Error(m); }
        return { handle, name: handle.name };
      }
      if (handle && handle.lumenToken) handle = null;
      if (window.showSaveFilePicker) {
        try {
          if (!handle || saveAs) handle = await window.showSaveFilePicker({ suggestedName: name, types });
          const w = await handle.createWritable();
          await w.write(blob);
          await w.close();
          return { handle, name: handle.name };
        } catch (e) {
          if (e && e.name === 'AbortError') return null;
          console.warn('Falling back to download:', e);
        }
      }
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = name;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
      return { handle: null, name };
    },
    // Windows: files opened from Explorer ("Open with" an installed Lumen app) arrive through the launch queue
    onLaunch(cb) {
      // opened by the installed "Open with Lumen ..." entry: the local server hands the file over by token
      const tok = new URLSearchParams(location.search).get('open');
      if (tok && /^[\w-]{10,}$/.test(tok)) {
        (async () => {
          try {
            const r = await fetch('/_file/' + tok);
            if (!r.ok) throw new Error('The file could not be read (' + r.status + ').');
            const name = decodeURIComponent(r.headers.get('X-Lumen-Name') || 'Document');
            const file = new File([await r.blob()], name);
            history.replaceState(null, '', location.pathname);
            cb(file, { lumenToken: tok, name });
          } catch (e) { toast(String(e.message || e), 6000); }
        })();
      }
      if (!('launchQueue' in window)) return;
      window.launchQueue.setConsumer(async (p) => {
        if (!p || !p.files || !p.files.length) return;
        try { const h = p.files[0]; cb(await h.getFile(), h); } catch (e) { console.warn(e); }
      });
    },
    onDrop(cb) {
      window.addEventListener('dragover', e => { if (e.dataTransfer && [...e.dataTransfer.types].includes('Files')) e.preventDefault(); });
      window.addEventListener('drop', e => {
        if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length && !e.defaultPrevented) {
          e.preventDefault();
          cb(e.dataTransfer.files[0]);
        }
      });
    },
  };

  /* ---------- ribbon tab switching ---------- */
  function initRibbon() {
    $$('.tab').forEach(t => t.addEventListener('click', () => {
      $$('.tab').forEach(x => x.classList.toggle('active', x === t));
      $$('.panel').forEach(p => p.classList.toggle('active', p.dataset.panel === t.dataset.tab));
    }));
  }
  function showTab(name, show) {
    const t = $('.tab[data-tab="' + name + '"]');
    if (t) t.hidden = !show;
  }
  function activateTab(name) {
    const t = $('.tab[data-tab="' + name + '"]');
    if (t) t.click();
  }

  /* ---------- Windows desktop integration ---------- */
  /* ---------- updates from GitHub (through the local Lumen server) ---------- */
  async function checkUpdates(manual) {
    let st;
    try {
      const r = await fetch('/_update/status' + (manual ? '?refresh=1' : ''), { cache: 'no-store' });
      if (!r.ok || !(r.headers.get('Content-Type') || '').includes('json')) { if (manual) toast('Updates work in the installed apps.', 6000); return null; }
      st = await r.json();
    } catch (e) { if (manual) toast('Could not reach the update service.'); return null; }
    if (manual && !st.available) { toast(st.error || ('Version ' + st.current + ' is the newest version.'), 6000); return st; }
    if (!st.available) return st;
    let snooze = null;
    try { snooze = localStorage.getItem('lumen-update-later'); } catch (e) { }
    if (!manual && snooze === st.latest) return st;
    if (document.getElementById('lumen-update')) return st;
    const bar = document.createElement('div');
    bar.id = 'lumen-update';
    bar.style.cssText = 'position:fixed;left:50%;bottom:36px;transform:translateX(-50%);z-index:6000;background:#1f1f1f;color:#fff;padding:10px 14px;border-radius:8px;box-shadow:0 6px 24px rgba(0,0,0,.35);font:14px "Segoe UI",system-ui,sans-serif;display:flex;gap:10px;align-items:center;max-width:92vw';
    bar.innerHTML = '<span></span><button data-u="now">Update now</button><button data-u="notes">What\'s new</button><button data-u="later">Later</button>';
    bar.firstChild.textContent = 'Version ' + st.latest + ' is available (you have ' + st.current + ').';
    bar.querySelectorAll('button').forEach(b => { b.style.cssText = 'background:#3d3d3d;color:#fff;border:1px solid #666;border-radius:4px;padding:4px 10px;font:inherit;cursor:pointer'; });
    bar.querySelector('[data-u="now"]').style.background = '#2b579a';
    bar.addEventListener('click', async e => {
      const u = e.target.dataset && e.target.dataset.u;
      if (u === 'later') { try { localStorage.setItem('lumen-update-later', st.latest); } catch (x) { } bar.remove(); }
      if (u === 'notes') { await form('What\'s new in ' + st.latest, [{ type: 'note', label: st.notes || 'No release notes.' }], 'Close'); }
      if (u === 'now') {
        if (document.getElementById('dirty') && document.getElementById('dirty').textContent.trim() && !(await confirmBox('Unsaved changes', 'Save your work first: the window reloads after the update. Update anyway?', 'Update'))) return;
        bar.firstChild.textContent = 'Downloading and installing ' + st.latest + '…';
        bar.querySelectorAll('button').forEach(b => { b.disabled = true; });
        try {
          const r = await fetch('/_update/apply', { method: 'POST', headers: { 'X-Lumen': '1' } });
          const j = await r.json();
          if (!r.ok) throw new Error(j.error || 'update failed');
          bar.firstChild.textContent = 'Updated to ' + j.version + '. Restarting…';
          const t0 = Date.now();
          const wait = async () => { try { const p = await fetch('/_ping', { cache: 'no-store' }); if (p.ok && Date.now() - t0 > 1500) { location.reload(); return; } } catch (x) { } setTimeout(wait, 500); };
          setTimeout(wait, 1000);
        } catch (err) { bar.firstChild.textContent = 'Update failed: ' + (err.message || err); bar.querySelectorAll('button').forEach(b => { b.disabled = false; }); }
      }
    });
    document.body.appendChild(bar);
    return st;
  }

  function initDesktop() {
    setTimeout(() => { if (!window.LumenSettings || LumenSettings.get('autoUpdate') !== false) checkUpdates(false); }, 3000);
    // offline cache + installable app (Edge/Chrome: menu > Apps > Install)
    if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol) && !/[?&]nosw\b/.test(location.search)) {
      navigator.serviceWorker.register(new URL('../sw.js', location.href).href, { scope: new URL('../', location.href).href }).catch(() => { });
    }
    // Office-style keys: F12 Save As, Ctrl+P print, Shift+F12 / Ctrl+S save
    window.addEventListener('keydown', e => {
      if (e.key === 'F12' && !e.ctrlKey && !e.altKey) {
        const b = document.getElementById(e.shiftKey ? 'btn-save' : 'btn-saveas');
        if (b) { e.preventDefault(); b.click(); }
      } else if ((e.ctrlKey || e.metaKey) && !e.shiftKey && !e.altKey && e.key.toLowerCase() === 'p') {
        const b = document.getElementById('btn-print');
        if (b) { e.preventDefault(); b.click(); }
      }
    });
    // Windows high-contrast / dark title bar colour follows the app colour
    const brand = document.querySelector('.brand');
    if (brand && !document.querySelector('meta[name="theme-color"]')) {
      const m = document.createElement('meta'); m.name = 'theme-color'; m.content = getComputedStyle(brand).backgroundColor; document.head.appendChild(m);
    }
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', initDesktop); else initDesktop();

  function baseName(n) { return String(n || '').replace(/\.[^.]+$/, ''); }
  function nextTick() { return new Promise(r => setTimeout(r, 0)); }

  global.Lumen = {
    $, $$, esc, xesc, cssToHex, dataUrlToBytes, bytesToDataUrl, MIME_BY_EXT, EXT_BY_MIME,
    readFileAsDataUrl, toast, form, confirmBox, FileIO, initRibbon, showTab, activateTab, baseName, nextTick, checkUpdates,
  };
})(window);
