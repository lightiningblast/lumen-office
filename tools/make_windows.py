"""Generates the Windows integration files: icons (.png/.ico), web app manifests, the offline cache list."""
import io, json, os
from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.makedirs(os.path.join(ROOT, 'icons'), exist_ok=True)

APPS = {
    'write': ('W', '#2b579a', 'Lumen Write', 'Word processor for .docx documents',
              {'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'], 'text/plain': ['.txt']}),
    'cells': ('X', '#217346', 'Lumen Cells', 'Spreadsheets for .xlsx workbooks',
              {'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'], 'text/csv': ['.csv']}),
    'slides': ('P', '#b7472a', 'Lumen Slides', 'Presentations for .pptx decks',
               {'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx']}),
    'office': ('L', '#444a57', 'Lumen Office', 'Word processor, spreadsheet and presentation editor', None),
}


def font(size):
    for f in ('segoeuib.ttf', 'arialbd.ttf'):
        try:
            return ImageFont.truetype(os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts', f), size)
        except OSError:
            pass
    return ImageFont.load_default()


def icon(letter, color, size):
    s = size * 4
    im = Image.new('RGBA', (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.rounded_rectangle([s * 0.04, s * 0.04, s * 0.96, s * 0.96], radius=s * 0.2, fill=color)
    f = font(int(s * 0.62))
    box = d.textbbox((0, 0), letter, font=f)
    d.text(((s - (box[2] - box[0])) / 2 - box[0], (s - (box[3] - box[1])) / 2 - box[1]), letter, font=f, fill='white')
    return im.resize((size, size), Image.LANCZOS)


for key, (letter, color, name, desc, accept) in APPS.items():
    for n in (192, 512):
        icon(letter, color, n).save(os.path.join(ROOT, 'icons', f'{key}-{n}.png'))
    icon(letter, color, 256).save(os.path.join(ROOT, 'icons', f'{key}.ico'),
                                  sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    up = '' if key == 'office' else '../'
    man = {
        'name': name, 'short_name': name, 'description': desc,
        'start_url': './index.html', 'scope': './', 'id': './index.html',
        'display': 'standalone',
        'background_color': '#f3f3f3', 'theme_color': color,
        'icons': [{'src': f'{up}icons/{key}-{n}.png', 'sizes': f'{n}x{n}', 'type': 'image/png', 'purpose': 'any'} for n in (192, 512)],
    }
    if accept:
        man['file_handlers'] = [{'action': './index.html', 'accept': accept}]
        man['launch_handler'] = {'client_mode': 'navigate-new'}
    target = os.path.join(ROOT, 'manifest.webmanifest' if key == 'office' else f'{key}/manifest.webmanifest')
    io.open(target, 'w', encoding='utf8').write(json.dumps(man, indent=2))

# offline cache list for the service worker
assets = ['./', 'index.html']
for folder in ('shared', 'write', 'cells', 'slides', 'icons'):
    for fn in sorted(os.listdir(os.path.join(ROOT, folder))):
        if fn.endswith(('.js', '.css', '.html', '.png', '.webmanifest')):
            assets.append(f'{folder}/{fn}')
sw = '''// Offline cache for Lumen Office. Network first (so updates show up), cached copy when the local server is not running.
const CACHE = 'lumen-office-%s';
const ASSETS = %s;
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
''' % (json.load(io.open(os.path.join(ROOT, 'version.json'), encoding='utf-8-sig'))['version'], json.dumps(assets, indent=2))
io.open(os.path.join(ROOT, 'sw.js'), 'w', encoding='utf8').write(sw)
print('icons, manifests and sw.js written;', len(assets), 'cached files')
