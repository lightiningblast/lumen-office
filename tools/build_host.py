"""Compiles host/LumenHost.cs into Lumen Write.exe, Lumen Cells.exe, Lumen Slides.exe and Lumen Office.exe (project root)
with the C# compiler that ships with Windows (.NET Framework 4.x), and copies the WebView2 DLLs next to them."""
import glob, os, shutil, subprocess, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WIN = os.environ.get('WINDIR', r'C:\Windows')
EXES = {'Write': 'write', 'Cells': 'cells', 'Slides': 'slides', 'Office Apps': 'office'}
# copies under the old names: an existing install that updates itself starts one of these once, which moves it to the new names
LEGACY = {'Lumen Write': 'Write', 'Lumen Cells': 'Cells', 'Lumen Slides': 'Slides', 'Lumen Office': 'Office Apps'}
DLLS = ['Microsoft.Web.WebView2.Core.dll', 'Microsoft.Web.WebView2.WinForms.dll', 'WebView2Loader.dll']


def csc():
    c = sorted(glob.glob(os.path.join(WIN, 'Microsoft.NET', 'Framework64', 'v4*', 'csc.exe')))
    if not c:
        sys.exit('csc.exe (.NET Framework 4) not found')
    return c[-1]


def main():
    lib = os.path.join(ROOT, 'host', 'lib')
    fw = os.path.dirname(csc())
    refs = [os.path.join(fw, 'System.Windows.Forms.dll'), os.path.join(fw, 'System.Drawing.dll'),
            os.path.join(fw, 'System.Web.Extensions.dll'), os.path.join(fw, 'System.IO.Compression.dll'),
            os.path.join(lib, 'Microsoft.Web.WebView2.Core.dll'), os.path.join(lib, 'Microsoft.Web.WebView2.WinForms.dll')]
    for name, icon in EXES.items():
        out = os.path.join(ROOT, name + '.exe')
        cmd = [csc(), '/nologo', '/target:winexe', '/platform:x64', '/optimize+', '/out:' + out,
               '/win32icon:' + os.path.join(ROOT, 'icons', icon + '.ico'),
               '/win32manifest:' + os.path.join(ROOT, 'host', 'app.manifest'),
               '/define:APP_' + icon.upper()] + ['/reference:' + r for r in refs] + [os.path.join(ROOT, 'host', 'LumenHost.cs')]
        r = subprocess.run(cmd, capture_output=True, text=True)
        if r.returncode:
            print(r.stdout, r.stderr)
            sys.exit('compile failed: ' + name)
        print('built', os.path.basename(out), os.path.getsize(out), 'bytes')
    for d in DLLS:
        shutil.copy2(os.path.join(lib, d), os.path.join(ROOT, d))
    for old, new in LEGACY.items():
        shutil.copy2(os.path.join(ROOT, new + '.exe'), os.path.join(ROOT, old + '.exe'))


if __name__ == '__main__':
    main()
