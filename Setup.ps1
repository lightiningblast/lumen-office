# Lumen Office installer (per user, no administrator rights).
#   - copies the apps to %LOCALAPPDATA%\Programs\Lumen Office
#   - Start menu folder: Lumen Write, Lumen Cells, Lumen Slides, Lumen Office, Check for updates, Uninstall
#   - optional desktop shortcuts
#   - "Open with" entries for .docx .xlsx .csv .pptx (your default apps are NOT changed)
#   - an entry in Settings > Apps > Installed apps, with an uninstaller
# Running it again over an existing install upgrades it and keeps your update.json setting.
param(
  [string]$InstallDir = (Join-Path $env:LOCALAPPDATA 'Programs\Lumen Office'),
  [switch]$NoDesktop,
  [switch]$Quiet,
  [switch]$FilesOnly   # copy files only: no shortcuts, no registry (used for testing)
)
$ErrorActionPreference = 'Stop'
$src = Split-Path -Parent $MyInvocation.MyCommand.Path
function Say($t) { if (-not $Quiet) { Write-Host $t } }
function Ask($q, $def) { if ($Quiet) { return $def }; $a = Read-Host "$q [$(if ($def) {'Y/n'} else {'y/N'})]"; if ($a -eq '') { return $def }; return $a -match '^(y|yes)$' }

$ver = (Get-Content -Raw (Join-Path $src 'version.json') | ConvertFrom-Json).version
Say "Lumen Office $ver setup"
Say "Install folder: $InstallDir"

# ---- Python (runs the small local server) ----
$py = $null
foreach ($n in 'pythonw.exe', 'python.exe', 'py.exe') { $c = Get-Command $n -ErrorAction SilentlyContinue; if ($c) { $py = $c.Source; break } }
if ($py) {
  $okVer = & ($py -replace 'pythonw\.exe$', 'python.exe') -c "import sys; print(int(sys.version_info >= (3, 8)))" 2>$null
  if ($okVer -ne '1') { $py = $null }
}
# the Lumen programs (Lumen Write.exe ...) need no Python; only the older browser-based launcher does
$needPy = -not (Test-Path (Join-Path $src 'Lumen Write.exe'))
if (-not $py -and -not $FilesOnly -and $needPy) {
  Say 'Python 3.8 or newer is needed and was not found.'
  if ((Get-Command winget -ErrorAction SilentlyContinue) -and (Ask 'Install Python 3.12 now with winget (Microsoft package manager)?' $true)) {
    winget install --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements
    $env:Path = [Environment]::GetEnvironmentVariable('Path', 'User') + ';' + [Environment]::GetEnvironmentVariable('Path', 'Machine')
    $c = Get-Command pythonw.exe -ErrorAction SilentlyContinue; if ($c) { $py = $c.Source }
  }
  if (-not $py) { Say 'Install Python from https://www.python.org/downloads/ and run Setup again.'; if (-not $Quiet) { Read-Host 'Press Enter to close' }; exit 1 }
}
$pyw = if ($py) { $py -replace 'python\.exe$', 'pythonw.exe' -replace 'py\.exe$', 'pyw.exe' } else { 'pythonw.exe' }

# ---- close running Lumen windows and the server so their files can be replaced ----
$running = Get-Process -Name 'Lumen Write', 'Lumen Cells', 'Lumen Slides', 'Lumen Office' -ErrorAction SilentlyContinue |
  Where-Object { $_.Path -and $_.Path -like "$InstallDir*" }
if ($running) {
  if (-not $Quiet) { Read-Host 'Lumen windows are open. Save your work in them, then press Enter to close them and continue' | Out-Null }
  $running | ForEach-Object { $_.CloseMainWindow() | Out-Null }
  Start-Sleep -Seconds 2
  $running | Where-Object { -not $_.HasExited } | Stop-Process -Force -ErrorAction SilentlyContinue
}
Get-CimInstance Win32_Process -Filter "Name like 'python%'" -ErrorAction SilentlyContinue |
  Where-Object { $_.CommandLine -and $_.CommandLine -like '*serve.py*' -and $_.CommandLine -like "*$InstallDir*" } |
  ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }

# ---- copy files ----
$keep = $null
$cfg = Join-Path $InstallDir 'update.json'
if (Test-Path $cfg) { $keep = Get-Content -Raw $cfg }
New-Item -ItemType Directory -Force $InstallDir | Out-Null
$items = 'shared', 'write', 'cells', 'slides', 'icons', 'host', 'index.html', 'sw.js', 'manifest.webmanifest', 'version.json', 'update.json', 'Setup.cmd', 'Setup.ps1', 'Uninstall.ps1', 'README.txt', 'Lumen Office.cmd',
  'Lumen Write.exe', 'Lumen Cells.exe', 'Lumen Slides.exe', 'Lumen Office.exe', 'Microsoft.Web.WebView2.Core.dll', 'Microsoft.Web.WebView2.WinForms.dll', 'WebView2Loader.dll'
if ((Resolve-Path $src).Path.TrimEnd('\') -ne [IO.Path]::GetFullPath($InstallDir).TrimEnd('\')) {
  foreach ($i in $items) { $p = Join-Path $src $i; if (Test-Path $p) { Copy-Item $p -Destination $InstallDir -Recurse -Force } }
  New-Item -ItemType Directory -Force (Join-Path $InstallDir 'tools') | Out-Null
  foreach ($t in 'serve.py', 'updater.py') { Copy-Item (Join-Path $src "tools\$t") -Destination (Join-Path $InstallDir 'tools') -Force }
}
if ($keep) {
  $old = $keep | ConvertFrom-Json
  if ($old.repo) {
    $new = Get-Content -Raw $cfg | ConvertFrom-Json; $new.repo = $old.repo
    [IO.File]::WriteAllText($cfg, ($new | ConvertTo-Json), (New-Object Text.UTF8Encoding $false))   # UTF-8 without BOM
  }
}
if ($py) { [IO.File]::WriteAllText((Join-Path $InstallDir 'python.path'), ($py -replace 'python\.exe$', 'pythonw.exe' -replace 'py\.exe$', 'pyw.exe'), (New-Object Text.UTF8Encoding $false)) }
Say 'Files copied.'
if ($FilesOnly) { Say 'Files-only mode: done.'; exit 0 }
# the Lumen program windows (own taskbar/dock entry); older packages without them fall back to the Python launcher
$hasExe = Test-Path (Join-Path $InstallDir 'Lumen Write.exe')
function ExeOf($app) { $n = @{ write = 'Lumen Write'; cells = 'Lumen Cells'; slides = 'Lumen Slides'; '' = 'Lumen Office' }[$app]; Join-Path $InstallDir "$n.exe" }

# ---- shortcuts ----
$shell = New-Object -ComObject WScript.Shell
$menu = Join-Path ([Environment]::GetFolderPath('Programs')) 'Lumen Office'
New-Item -ItemType Directory -Force $menu | Out-Null
$serve = Join-Path $InstallDir 'tools\serve.py'
function Link($dir, $name, $target, $arguments, $icon) {
  $l = $shell.CreateShortcut((Join-Path $dir "$name.lnk"))
  $l.TargetPath = $target; $l.Arguments = $arguments; $l.WorkingDirectory = $InstallDir
  $l.IconLocation = "$icon,0"; $l.Description = $name; $l.Save()
}
$apps = @(@{ N = 'Lumen Write'; A = 'write' }, @{ N = 'Lumen Cells'; A = 'cells' }, @{ N = 'Lumen Slides'; A = 'slides' }, @{ N = 'Lumen Office'; A = '' })
foreach ($a in $apps) {
  $ico = Join-Path $InstallDir ("icons\" + $(if ($a.A) { $a.A } else { 'office' }) + '.ico')
  if ($hasExe) { Link $menu $a.N (ExeOf $a.A) '' (ExeOf $a.A) }
  else { Link $menu $a.N $pyw ("`"$serve`" " + $a.A).Trim() $ico }
}
if ($hasExe) { Link $menu 'Check for updates' (ExeOf '') '--updates' (ExeOf '') }
else {
  $pyc = $pyw -replace 'pythonw\.exe$', 'python.exe' -replace 'pyw\.exe$', 'py.exe'
  Link $menu 'Check for updates' $pyc ("`"" + (Join-Path $InstallDir 'tools\updater.py') + "`"") (Join-Path $InstallDir 'icons\office.ico')
}
Link $menu 'Uninstall Lumen Office' 'powershell.exe' ("-NoProfile -ExecutionPolicy Bypass -File `"" + (Join-Path $InstallDir 'Uninstall.ps1') + "`"") (Join-Path $InstallDir 'icons\office.ico')
if (-not $NoDesktop -and (Ask 'Add Lumen Write, Cells and Slides to the desktop?' $true)) {
  $desk = [Environment]::GetFolderPath('Desktop')
  foreach ($a in $apps[0..2]) {
    if ($hasExe) { Link $desk $a.N (ExeOf $a.A) '' (ExeOf $a.A) }
    else { Link $desk $a.N $pyw "`"$serve`" $($a.A)" (Join-Path $InstallDir "icons\$($a.A).ico") }
  }
}

# ---- "Open with" for Office files (adds a choice, does not change your default app) ----
$classes = 'HKCU:\Software\Classes'
$types = @(
  @{ Id = 'LumenOffice.Write';  Name = 'Lumen Write document';     App = 'write';  Ext = '.docx' },
  @{ Id = 'LumenOffice.Cells';  Name = 'Lumen Cells workbook';     App = 'cells';  Ext = '.xlsx', '.csv' },
  @{ Id = 'LumenOffice.Slides'; Name = 'Lumen Slides presentation'; App = 'slides'; Ext = '.pptx' }
)
foreach ($t in $types) {
  $k = "$classes\$($t.Id)"
  New-Item -Force $k | Out-Null; Set-Item $k $t.Name
  New-Item -Force "$k\DefaultIcon" | Out-Null; Set-Item "$k\DefaultIcon" ((Join-Path $InstallDir "icons\$($t.App).ico") + ',0')
  New-Item -Force "$k\shell\open\command" | Out-Null
  if ($hasExe) { Set-Item "$k\shell\open\command" "`"$(ExeOf $t.App)`" `"%1`"" }
  else { Set-Item "$k\shell\open\command" "`"$pyw`" `"$serve`" $($t.App) `"%1`"" }
  # lets Windows show the right name and icon in "Open with"
  $ap = "$classes\Applications\$(Split-Path -Leaf (ExeOf $t.App))"
  if ($hasExe) { New-Item -Force "$ap\shell\open\command" | Out-Null; Set-Item "$ap\shell\open\command" "`"$(ExeOf $t.App)`" `"%1`""; New-ItemProperty -Force -Path $ap -Name FriendlyAppName -PropertyType String -Value ($t.Name -replace ' (document|workbook|presentation)$', '') | Out-Null }
  foreach ($e in $t.Ext) {
    New-Item -Force "$classes\$e\OpenWithProgids" | Out-Null
    New-ItemProperty -Force -Path "$classes\$e\OpenWithProgids" -Name $t.Id -PropertyType String -Value '' | Out-Null
  }
}

# ---- Installed apps entry ----
$un = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\LumenOffice'
New-Item -Force $un | Out-Null
$size = [int]((Get-ChildItem $InstallDir -Recurse -File | Measure-Object Length -Sum).Sum / 1KB)
$vals = @{
  DisplayName = 'Lumen Office'; DisplayVersion = $ver; Publisher = 'Lumen Office'; InstallLocation = $InstallDir
  DisplayIcon = (Join-Path $InstallDir 'icons\office.ico')
  UninstallString = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$(Join-Path $InstallDir 'Uninstall.ps1')`""
  InstallDate = (Get-Date -Format 'yyyyMMdd')
}
foreach ($k in $vals.Keys) { New-ItemProperty -Force -Path $un -Name $k -PropertyType String -Value $vals[$k] | Out-Null }
foreach ($k in 'NoModify', 'NoRepair') { New-ItemProperty -Force -Path $un -Name $k -PropertyType DWord -Value 1 | Out-Null }
New-ItemProperty -Force -Path $un -Name EstimatedSize -PropertyType DWord -Value $size | Out-Null

Say ''
Say "Lumen Office $ver is installed."
Say 'Start it from the Start menu (Lumen Office folder). To open Office files with it: right-click a file > Open with > Lumen Write / Cells / Slides.'
$repo = (Get-Content -Raw $cfg | ConvertFrom-Json).repo
if (-not $repo) { Say "Updates: no GitHub repository is set yet. Put it in $cfg as ""owner/name""." }
Say 'Open Lumen Write, Cells or Slides from the Start menu or the desktop.'
