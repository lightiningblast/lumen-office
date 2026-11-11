"""Builds dist/LumenOffice-<version>.zip and its .sha256 file: the package that Setup installs and the updater downloads.

  python tools/make_release.py                 uses version.json
  python tools/make_release.py --expect v1.2.0 fails if version.json does not say 1.2.0 (used by the GitHub workflow)
"""
import hashlib, json, os, subprocess, sys, zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INCLUDE_DIRS = ['shared', 'write', 'cells', 'slides', 'icons']
INCLUDE_FILES = ['index.html', 'sw.js', 'manifest.webmanifest', 'version.json', 'update.json',
                 'Setup.cmd', 'Setup.ps1', 'Uninstall.ps1', 'README.txt', 'Lumen Office.cmd',
                 'tools/serve.py', 'tools/updater.py',
                 'Lumen Write.exe', 'Lumen Cells.exe', 'Lumen Slides.exe', 'Lumen Office.exe',
                 'Microsoft.Web.WebView2.Core.dll', 'Microsoft.Web.WebView2.WinForms.dll', 'WebView2Loader.dll',
                 'host/lib/WebView2-LICENSE.txt']


def main():
    ver = json.load(open(os.path.join(ROOT, 'version.json'), encoding='utf-8-sig'))['version']
    if '--expect' in sys.argv:
        want = sys.argv[sys.argv.index('--expect') + 1].lstrip('vV')
        if want != ver:
            sys.exit(f'version.json says {ver} but the tag says {want}. Update version.json first.')
    # refresh icons, manifests and the offline cache list (its cache name carries the version)
    subprocess.check_call([sys.executable, os.path.join(ROOT, 'tools', 'make_windows.py')])
    # the Lumen program windows (Lumen Write.exe ...)
    subprocess.check_call([sys.executable, os.path.join(ROOT, 'tools', 'build_host.py')])
    files = list(INCLUDE_FILES)
    for d in INCLUDE_DIRS:
        for dp, _, fns in os.walk(os.path.join(ROOT, d)):
            for fn in fns:
                files.append(os.path.relpath(os.path.join(dp, fn), ROOT).replace('\\', '/'))
    missing = [f for f in files if not os.path.isfile(os.path.join(ROOT, f))]
    if missing:
        sys.exit('missing files: ' + ', '.join(missing))
    os.makedirs(os.path.join(ROOT, 'dist'), exist_ok=True)
    out = os.path.join(ROOT, 'dist', f'LumenOffice-{ver}.zip')
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in sorted(set(files)):
            src = os.path.join(ROOT, f)
            if f == 'update.json':
                # ship the repository setting but never a user's local edits beyond it
                cfg = json.load(open(src, encoding='utf-8-sig'))
                z.writestr(f, json.dumps(cfg, indent=2) + '\n')
            else:
                z.write(src, f)
    digest = hashlib.sha256(open(out, 'rb').read()).hexdigest()
    with open(out + '.sha256', 'w', encoding='ascii') as f:
        f.write(f'{digest}  {os.path.basename(out)}\n')
    print(f'{out}\n{digest}\n{len(set(files))} files')


if __name__ == '__main__':
    main()
