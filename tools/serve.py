"""Lumen Office local server for Windows.

Serves the apps on http://127.0.0.1:8765 (only this computer can reach it) and opens the requested app in its own
Edge/Chrome window. If the server is already running it hands the request to it and just opens the window.

  pythonw tools\\serve.py                        start page
  pythonw tools\\serve.py write                  Lumen Write   (also: cells, slides)
  pythonw tools\\serve.py write "C:\\a\\b.docx"   open a file (what Explorer runs for "Open with Lumen Write")
  pythonw tools\\serve.py --no-open              only start the server

Safety: requests must come to 127.0.0.1/localhost (blocks DNS-rebinding); requests that change something need the
X-Lumen header from a page of this server (browsers will not send it cross-site) or the per-run secret that only
local programs can read.
"""
import http.server, json, os, secrets, shutil, socket, subprocess, sys, threading, time, urllib.parse, urllib.request, webbrowser

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import updater  # noqa: E402

PORT = int(os.environ.get('LUMEN_PORT', '8765'))
DATA = os.path.join(os.environ.get('LOCALAPPDATA') or os.path.expanduser('~'), 'Lumen Office')
SECRET_FILE = os.path.join(DATA, f'server-{PORT}.secret')
APPS = {'write': 'write/index.html', 'cells': 'cells/index.html', 'slides': 'slides/index.html'}
EXT_APP = {'.docx': 'write', '.docm': 'write', '.dotx': 'write', '.txt': 'write', '.xlsx': 'cells', '.xlsm': 'cells', '.csv': 'cells', '.pptx': 'slides'}
OPEN_TYPES = set(EXT_APP)
SECRET = secrets.token_urlsafe(24)
OPENED = {}  # token -> absolute path of a file the user opened from Explorer
LOCK = threading.Lock()


def register(path):
    path = os.path.abspath(path)
    if os.path.splitext(path)[1].lower() not in OPEN_TYPES or not os.path.isfile(path):
        return None
    with LOCK:
        for t, p in OPENED.items():
            if p == path:
                return t
        t = secrets.token_urlsafe(18)
        OPENED[t] = path
        return t


class Handler(http.server.SimpleHTTPRequestHandler):
    extensions_map = {**http.server.SimpleHTTPRequestHandler.extensions_map,
                      '.js': 'text/javascript', '.webmanifest': 'application/manifest+json', '.ico': 'image/x-icon', '.json': 'application/json'}

    def __init__(self, *a, **k):
        super().__init__(*a, directory=ROOT, **k)

    def log_message(self, *a):
        pass

    def end_headers(self):
        self.send_header('Cache-Control', 'no-cache')
        self.send_header('X-Content-Type-Options', 'nosniff')
        super().end_headers()

    # ---- checks ----
    def host_ok(self):
        h = (self.headers.get('Host') or '').lower()
        return h in (f'127.0.0.1:{PORT}', f'localhost:{PORT}')

    def from_page(self):
        o = self.headers.get('Origin')
        same = o in (None, f'http://127.0.0.1:{PORT}', f'http://localhost:{PORT}')
        return same and self.headers.get('X-Lumen') == '1'

    def from_local_program(self):
        return secrets.compare_digest(self.headers.get('X-Lumen-Secret', ''), SECRET)

    def reply(self, code, obj=None, body=None, ctype='application/json'):
        data = body if body is not None else json.dumps(obj if obj is not None else {}).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def body(self, limit=300 * 1024 * 1024):
        n = int(self.headers.get('Content-Length') or 0)
        if n > limit:
            raise ValueError('too large')
        return self.rfile.read(n)

    # ---- routes ----
    def do_GET(self):
        if not self.host_ok():
            return self.reply(403, {'error': 'forbidden host'})
        u = urllib.parse.urlparse(self.path)
        if u.path == '/_ping':
            return self.reply(200, {'app': 'lumen-office', 'version': updater.current_version()})
        if u.path.startswith('/_file/'):
            p = OPENED.get(u.path[7:])
            if not p or not os.path.isfile(p):
                return self.reply(404, {'error': 'file not found'})
            with open(p, 'rb') as f:
                data = f.read()
            self.send_response(200)
            self.send_header('Content-Type', 'application/octet-stream')
            self.send_header('Content-Length', str(len(data)))
            self.send_header('X-Lumen-Name', urllib.parse.quote(os.path.basename(p)))
            self.end_headers()
            self.wfile.write(data)
            return
        if u.path == '/_update/status':
            force = 'refresh=1' in (u.query or '')
            return self.reply(200, updater.check(force=force))
        if u.path.startswith('/_') or '/tools/' in u.path or u.path.startswith('/_testout'):
            return self.reply(404, {'error': 'not found'})
        return super().do_GET()

    def do_HEAD(self):
        if not self.host_ok():
            return self.reply(403, {'error': 'forbidden host'})
        return super().do_HEAD()

    def do_PUT(self):
        if not self.host_ok() or not self.from_page():
            return self.reply(403, {'error': 'forbidden'})
        u = urllib.parse.urlparse(self.path)
        if u.path.startswith('/_file/'):
            p = OPENED.get(u.path[7:])
            if not p:
                return self.reply(404, {'error': 'unknown file'})
            try:
                data = self.body()
                tmp = p + '.lumen-tmp'
                with open(tmp, 'wb') as f:
                    f.write(data)
                os.replace(tmp, p)
            except PermissionError:
                return self.reply(423, {'error': 'The file is open in another program or read-only.'})
            except Exception as e:
                return self.reply(500, {'error': str(e)})
            return self.reply(200, {'saved': os.path.basename(p)})
        return self.reply(404, {'error': 'not found'})

    def do_POST(self):
        if not self.host_ok():
            return self.reply(403, {'error': 'forbidden host'})
        u = urllib.parse.urlparse(self.path)
        if u.path == '/_open' and self.from_local_program():
            try:
                path = json.loads(self.body(10000) or b'{}').get('path', '')
            except ValueError:
                return self.reply(400, {'error': 'bad request'})
            t = register(path)
            return self.reply(200 if t else 404, {'token': t} if t else {'error': 'cannot open that file'})
        if u.path == '/_update/apply' and self.from_page():
            try:
                ver = updater.apply(log=lambda *a: None)
            except Exception as e:
                return self.reply(409, {'error': str(e)})
            self.reply(200, {'version': ver, 'restarting': True})
            threading.Thread(target=restart, daemon=True).start()
            return
        return self.reply(403, {'error': 'forbidden'})


SERVER = None


def restart():
    """After an update: start a fresh server process (it waits for the port), then stop this one."""
    time.sleep(0.5)
    exe = sys.executable
    if os.path.basename(exe).lower() == 'python.exe' and os.path.exists(exe[:-10] + 'pythonw.exe'):
        exe = exe[:-10] + 'pythonw.exe'
    subprocess.Popen([exe, os.path.abspath(__file__), '--no-open', '--wait-port'], close_fds=True, cwd=ROOT)
    if SERVER:
        SERVER.shutdown()


def ping():
    try:
        req = urllib.request.Request(f'http://127.0.0.1:{PORT}/_ping', headers={'Host': f'127.0.0.1:{PORT}'})
        with urllib.request.urlopen(req, timeout=1) as r:
            return json.loads(r.read()).get('app') == 'lumen-office'
    except Exception:
        return False


def port_busy():
    try:
        with socket.create_connection(('127.0.0.1', PORT), timeout=0.3):
            return True
    except OSError:
        return False


def browser():
    cands = [
        os.path.expandvars(r'%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe'),
        os.path.expandvars(r'%ProgramFiles%\Microsoft\Edge\Application\msedge.exe'),
        os.path.expandvars(r'%ProgramFiles%\Google\Chrome\Application\chrome.exe'),
        os.path.expandvars(r'%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe'),
        os.path.expandvars(r'%LocalAppData%\Google\Chrome\Application\chrome.exe'),
    ]
    for c in cands:
        if os.path.isfile(c):
            return c
    return shutil.which('msedge') or shutil.which('chrome')


HOST_EXE = {'write': 'Write.exe', 'cells': 'Cells.exe', 'slides': 'Slides.exe', '': 'Office Apps.exe'}


def open_window(app, token=None):
    # prefer the Lumen program windows; fall back to an Edge/Chrome app window
    host = os.path.join(ROOT, HOST_EXE.get(app, 'Lumen Office.exe'))
    if os.path.isfile(host):
        subprocess.Popen([host] + (['--token', token] if token else []), close_fds=True, cwd=ROOT)
        return
    url = f'http://127.0.0.1:{PORT}/' + APPS.get(app, 'index.html') + (f'?open={token}' if token else '')
    exe = browser()
    if exe:
        subprocess.Popen([exe, f'--app={url}'], close_fds=True)  # a window without tabs or address bar
    else:
        webbrowser.open(url)


def hand_over(path):
    """Ask the already running server to register a file; returns its token."""
    try:
        with open(SECRET_FILE, encoding='ascii') as f:
            sec = f.read().strip()
        req = urllib.request.Request(f'http://127.0.0.1:{PORT}/_open', data=json.dumps({'path': os.path.abspath(path)}).encode(),
                                     headers={'Content-Type': 'application/json', 'X-Lumen-Secret': sec}, method='POST')
        with urllib.request.urlopen(req, timeout=3) as r:
            return json.loads(r.read()).get('token')
    except Exception:
        return None


def main():
    global SERVER
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    no_open = '--no-open' in sys.argv
    app = args[0].lower() if args and args[0].lower() in APPS else ''
    path = next((a for a in args if a.lower() not in APPS), None)
    if path and not app:
        app = EXT_APP.get(os.path.splitext(path)[1].lower(), 'write')

    if '--wait-port' in sys.argv:
        for _ in range(50):
            if not port_busy():
                break
            time.sleep(0.1)

    if ping():  # already running: let it serve, just open a window
        if not no_open:
            token = hand_over(path) if path else None
            open_window(app, token)
        return
    if port_busy():
        msg = f'Port {PORT} is used by another program, so Lumen Office cannot start.'
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(None, msg, 'Lumen Office', 0x10)
        except Exception:
            print(msg)
        return

    os.makedirs(DATA, exist_ok=True)
    # leftovers from an update that replaced programs while they were running
    for dp, _, fns in os.walk(ROOT):
        for fn in fns:
            if '.lumen-old-' in fn:
                try:
                    os.remove(os.path.join(dp, fn))
                except OSError:
                    pass
    with open(SECRET_FILE, 'w', encoding='ascii') as f:
        f.write(SECRET)
    SERVER = http.server.ThreadingHTTPServer(('127.0.0.1', PORT), Handler)
    token = register(path) if path else None
    if not no_open:
        threading.Thread(target=lambda: (time.sleep(0.3), open_window(app, token)), daemon=True).start()
    # quiet background check so the apps can show "update available"
    threading.Thread(target=lambda: (time.sleep(5), updater.check()), daemon=True).start()
    try:
        SERVER.serve_forever()
    finally:
        try:
            os.remove(SECRET_FILE)
        except OSError:
            pass


if __name__ == '__main__':
    main()
