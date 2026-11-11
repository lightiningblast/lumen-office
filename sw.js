// Offline cache for Lumen Office. Network first (so updates show up), cached copy when the local server is not running.
const CACHE = 'lumen-office-1.2.0';
const ASSETS = [
  "./",
  "index.html",
  "shared/common.css",
  "shared/common.js",
  "shared/icons.js",
  "shared/theme.css",
  "shared/zip.js",
  "write/docx.js",
  "write/index.html",
  "write/manifest.webmanifest",
  "write/write.css",
  "write/write.js",
  "cells/cells.css",
  "cells/chart.js",
  "cells/condfmt.js",
  "cells/formula.js",
  "cells/index.html",
  "cells/manifest.webmanifest",
  "cells/numfmt.js",
  "cells/sheet.js",
  "cells/validate.js",
  "cells/xlsx.js",
  "slides/index.html",
  "slides/manifest.webmanifest",
  "slides/pptx.js",
  "slides/slides.css",
  "slides/slides.js",
  "icons/cells-192.png",
  "icons/cells-512.png",
  "icons/office-192.png",
  "icons/office-512.png",
  "icons/slides-192.png",
  "icons/slides-512.png",
  "icons/write-192.png",
  "icons/write-512.png"
];
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith('lumen-office-') && k !== CACHE).map(k => caches.delete(k))))); });
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin || u.pathname.startsWith('/_')) return;
  e.respondWith(fetch(e.request).then(r => {
    if (r && r.ok) { const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); }
    return r;
  }).catch(() => caches.match(e.request, { ignoreSearch: true }).then(r => r || Response.error())));
});
