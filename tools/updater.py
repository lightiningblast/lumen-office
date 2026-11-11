"""Lumen Office updater: installs the newest GitHub Release of the repository named in update.json.

A release must carry an asset named LumenOffice-<version>.zip (made by tools/make_release.py) and, next to it,
LumenOffice-<version>.zip.sha256. The checksum is required: an update whose download does not match it is refused.

  python tools/updater.py            check, ask, install (for the "Check for updates" shortcut)
  python tools/updater.py --check    print the status as JSON
  python tools/updater.py --apply    install without asking (used by the app's "Update now" button)

Environment (for testing only): LUMEN_UPDATE_API replaces https://api.github.com.
"""
import hashlib, io, json, os, re, shutil, sys, tempfile, time, urllib.request, zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(os.environ.get('LOCALAPPDATA') or os.path.expanduser('~'), 'Lumen Office')
CACHE = os.path.join(DATA, 'update-status.json')
API = os.environ.get('LUMEN_UPDATE_API', 'https://api.github.com').rstrip('/')
UA = 'LumenOffice-Updater'
# files the user owns; an update never overwrites them
KEEP = {'update.json', 'python.path'}


def _json(path, default=None):
    try:
        with open(path, encoding='utf-8-sig') as f:  # -sig: Windows PowerShell 5.1 writes UTF-8 with a BOM
            return json.load(f)
    except (OSError, ValueError):
        return default


def current_version(root=ROOT):
    return (_json(os.path.join(root, 'version.json'), {}) or {}).get('version', '0.0.0')


def settings(root=ROOT):
    s = _json(os.path.join(root, 'update.json'), {}) or {}
    s['repo'] = (s.get('repo') or '').strip().strip('/')
    return s


def vtuple(v):
    nums = [int(x) for x in re.findall(r'\d+', str(v))[:3]]
    return tuple(nums + [0] * (3 - len(nums)))


def _get(url, accept='application/vnd.github+json', limit=200 * 1024 * 1024):
    req = urllib.request.Request(url, headers={'User-Agent': UA, 'Accept': accept})
    with urllib.request.urlopen(req, timeout=30) as r:
        data = r.read(limit + 1)
    if len(data) > limit:
        raise RuntimeError('download is larger than expected')
    return data


def check(root=ROOT, force=False):
    """Returns a status dict; caches it for check_every_hours unless force."""
    s = settings(root)
    cur = current_version(root)
    status = {'current': cur, 'repo': s['repo'], 'available': False, 'checked': int(time.time())}
    if not s['repo']:
        status['error'] = 'No GitHub repository is set. Put it in update.json as "owner/name".'
        return status
    if not re.fullmatch(r'[\w.-]+/[\w.-]+', s['repo']):
        status['error'] = 'update.json "repo" must look like owner/name.'
        return status
    cached = _json(CACHE)
    hours = float(s.get('check_every_hours') or 24)
    if not force and cached and cached.get('repo') == s['repo'] and cached.get('current') == cur and time.time() - cached.get('checked', 0) < hours * 3600:
        return cached
    try:
        if s.get('allow_prerelease'):
            rels = json.loads(_get(f'{API}/repos/{s["repo"]}/releases?per_page=10'))
            rel = next((r for r in rels if not r.get('draft')), None)
        else:
            rel = json.loads(_get(f'{API}/repos/{s["repo"]}/releases/latest'))
        if not rel:
            raise RuntimeError('the repository has no releases yet')
        tag = rel.get('tag_name') or ''
        assets = {a['name']: a['browser_download_url'] for a in rel.get('assets', [])}
        zipname = next((n for n in assets if re.fullmatch(r'LumenOffice-[\w.-]+\.zip', n)), None)
        status.update(latest=tag.lstrip('vV'), notes=(rel.get('body') or '')[:2000], page=rel.get('html_url', ''),
                      zip=assets.get(zipname, ''), sha256=assets.get((zipname or '') + '.sha256', ''))
        status['available'] = vtuple(tag) > vtuple(cur) and bool(status['zip'])
        if vtuple(tag) > vtuple(cur) and not status['zip']:
            status['error'] = f'Release {tag} has no LumenOffice-*.zip file attached.'
    except Exception as e:  # network down, repo private or missing, rate limit...
        status['error'] = f'Could not check GitHub: {e}'
    try:
        os.makedirs(DATA, exist_ok=True)
        with open(CACHE, 'w', encoding='utf-8') as f:
            json.dump(status, f)
    except OSError:
        pass
    return status


def apply(root=ROOT, status=None, log=print):
    st = status or check(root, force=True)
    if not st.get('available'):
        raise RuntimeError(st.get('error') or 'Lumen Office is up to date.')
    if not st.get('sha256'):
        raise RuntimeError('The release has no .sha256 checksum file, so the download cannot be verified. Update refused.')
    log(f'Downloading Lumen Office {st["latest"]}...')
    blob = _get(st['zip'], accept='application/octet-stream')
    want = _get(st['sha256'], accept='application/octet-stream').decode('ascii', 'replace').split()[0].lower()
    got = hashlib.sha256(blob).hexdigest()
    if got != want:
        raise RuntimeError('The download does not match its checksum. Update refused.')
    stage = tempfile.mkdtemp(prefix='lumen-update-')
    try:
        with zipfile.ZipFile(io.BytesIO(blob)) as z:
            for n in z.namelist():
                p = os.path.normpath(n)
                if p.startswith('..') or os.path.isabs(p) or ':' in p:
                    raise RuntimeError('The update contains an unsafe file path. Update refused.')
            z.extractall(stage)
        # the zip may hold the files at its top level or inside one folder
        src = stage
        entries = os.listdir(stage)
        if len(entries) == 1 and os.path.isdir(os.path.join(stage, entries[0])):
            src = os.path.join(stage, entries[0])
        if not os.path.isfile(os.path.join(src, 'version.json')) or not os.path.isdir(os.path.join(src, 'write')):
            raise RuntimeError('The download is not a Lumen Office release. Update refused.')
        log('Installing...')
        backup = os.path.join(DATA, 'previous-version')
        shutil.rmtree(backup, ignore_errors=True)
        for dirpath, dirs, files in os.walk(src):
            rel = os.path.relpath(dirpath, src)
            os.makedirs(os.path.join(root, rel), exist_ok=True)
            for fn in files:
                relf = os.path.normpath(os.path.join(rel, fn))
                if relf in KEEP and os.path.exists(os.path.join(root, relf)):
                    continue
                dst = os.path.join(root, relf)
                if os.path.exists(dst):
                    os.makedirs(os.path.dirname(os.path.join(backup, relf)), exist_ok=True)
                    shutil.copy2(dst, os.path.join(backup, relf))
                tmp = dst + '.new'
                shutil.copy2(os.path.join(dirpath, fn), tmp)
                try:
                    os.replace(tmp, dst)
                except PermissionError:
                    # a running program (Lumen Write.exe, a WebView2 DLL...) cannot be overwritten, but Windows
                    # lets it be renamed: move it aside (deleted on the next start) and put the new file in place
                    os.replace(dst, dst + '.lumen-old-' + str(int(time.time())))
                    os.replace(tmp, dst)
    finally:
        shutil.rmtree(stage, ignore_errors=True)
    try:
        os.remove(CACHE)
    except OSError:
        pass
    log(f'Lumen Office is now version {current_version(root)}. A copy of the old files is in {backup}.')
    return current_version(root)


def main():
    if '--check' in sys.argv:
        print(json.dumps(check(force='--force' in sys.argv), indent=2))
        return
    if '--apply' in sys.argv:
        try:
            apply()
        except Exception as e:
            print('Update failed:', e)
            sys.exit(1)
        return
    st = check(force=True)
    print(f'Lumen Office {st["current"]}')
    if st.get('error'):
        print(st['error'])
    elif not st['available']:
        print(f'You have the newest version (latest on GitHub: {st.get("latest", "?")}).')
    else:
        print(f'Version {st["latest"]} is available.')
        if st.get('notes'):
            print('\n' + st['notes'] + '\n')
        if input('Install it now? [Y/n] ').strip().lower() in ('', 'y', 'yes'):
            try:
                apply(status=st)
                print('Done. Close and reopen any Lumen windows to use the new version.')
            except Exception as e:
                print('Update failed:', e)
    input('\nPress Enter to close.')


if __name__ == '__main__':
    main()
