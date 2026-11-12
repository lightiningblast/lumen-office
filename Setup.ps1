# Office Apps installer: Write, Cells and Slides (per user, no administrator rights).
#   - copies the apps to %LOCALAPPDATA%\Programs\Office Apps
#   - Start menu: Write, Cells, Slides
#   - optional desktop shortcuts
#   - "Open with" entries for .docx .xlsx .csv .pptx (your default apps are not changed)
#   - an entry in Settings > Apps > Installed apps, with an uninstaller
# Running it over an older install (also the old "Lumen" one) upgrades it in place of the old one.
param(
  [string]$InstallDir = (Join-Path $env:LOCALAPPDATA 'Programs\Office Apps'),
  [switch]$NoDesktop,
  [switch]$Quiet,
  [switch]$FilesOnly   # copy files only: no shortcuts, no registry (used for testing)
)
$ErrorActionPreference = 'Stop'
$src = Split-Path -Parent $MyInvocation.MyCommand.Path
function Say($t) { if (-not $Quiet) { Write-Host $t } }
function Ask($q, $def) { if ($Quiet) { return $def }; $a = Read-Host "$q [$(if ($def) {'Y/n'} else {'y/N'})]"; if ($a -eq '') { return $def }; return $a -match '^(y|yes)$' }

$ver = (Get-Content -Raw (Join-Path $src 'version.json') | ConvertFrom-Json).version
$un = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\LumenOffice'
$oldDir = (Get-ItemProperty $un -ErrorAction SilentlyContinue).InstallLocation
Say "Office Apps $ver setup"
Say "Install folder: $InstallDir"

$APPS = @(
  @{ A = 'write';  N = 'Write';  Old = 'Lumen Write';  Id = 'LumenOffice.Write';  Type = 'Write document';      Ext = @('.docx') },
  @{ A = 'cells';  N = 'Cells';  Old = 'Lumen Cells';  Id = 'LumenOffice.Cells';  Type = 'Cells workbook';      Ext = @('.xlsx', '.csv') },
  @{ A = 'slides'; N = 'Slides'; Old = 'Lumen Slides'; Id = 'LumenOffice.Slides'; Type = 'Slides presentation'; Ext = @('.pptx') }
)
$ALLNAMES = 'Write', 'Cells', 'Slides', 'Office Apps', 'Lumen Write', 'Lumen Cells', 'Lumen Slides', 'Lumen Office'

# ---- close running windows of these apps so their files can be replaced ----
$dirs = @($InstallDir, $oldDir) | Where-Object { $_ }
$running = Get-Process -Name $ALLNAMES -ErrorAction SilentlyContinue | Where-Object { $p = $_.Path; $p -and ($dirs | Where-Object { $p -like "$_*" }) }
if ($running) {
  if (-not $Quiet) { Read-Host 'Write, Cells or Slides is open. Save your work, then press Enter to close them and continue' | Out-Null }
  $running | ForEach-Object { $_.CloseMainWindow() | Out-Null }
  Start-Sleep -Seconds 2
  $running | Where-Object { -not $_.HasExited } | Stop-Process -Force -ErrorAction SilentlyContinue
}
Get-CimInstance Win32_Process -Filter "Name like 'python%'" -ErrorAction SilentlyContinue |
  Where-Object { $c = $_.CommandLine; $c -and $c -like '*serve.py*' -and ($dirs | Where-Object { $c -like "*$_*" }) } |
  ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }

# ---- copy files ----
$keep = $null
foreach ($d in $dirs) { $f = Join-Path $d 'update.json'; if (-not $keep -and (Test-Path $f)) { $keep = Get-Content -Raw $f } }
New-Item -ItemType Directory -Force $InstallDir | Out-Null
$items = 'shared', 'write', 'cells', 'slides', 'icons', 'host', 'index.html', 'sw.js', 'manifest.webmanifest', 'version.json', 'update.json', 'Setup.cmd', 'Setup.ps1', 'Uninstall.ps1', 'README.txt',
  'Write.exe', 'Cells.exe', 'Slides.exe', 'Office Apps.exe', 'Microsoft.Web.WebView2.Core.dll', 'Microsoft.Web.WebView2.WinForms.dll', 'WebView2Loader.dll'
if ((Resolve-Path $src).Path.TrimEnd('\') -ne [IO.Path]::GetFullPath($InstallDir).TrimEnd('\')) {
  foreach ($i in $items) { $p = Join-Path $src $i; if (Test-Path $p) { Copy-Item $p -Destination $InstallDir -Recurse -Force } }
  New-Item -ItemType Directory -Force (Join-Path $InstallDir 'tools') | Out-Null
  foreach ($t in 'serve.py', 'updater.py', 'fix_openwith.ps1') { $p = Join-Path $src "tools\$t"; if (Test-Path $p) { Copy-Item $p -Destination (Join-Path $InstallDir 'tools') -Force } }
}
# the old "Lumen ..." program files are not needed any more
foreach ($o in 'Lumen Write.exe', 'Lumen Cells.exe', 'Lumen Slides.exe', 'Lumen Office.exe', 'Lumen Office.cmd') { Remove-Item -Force (Join-Path $InstallDir $o) -ErrorAction SilentlyContinue }
$cfg = Join-Path $InstallDir 'update.json'
if ($keep) {
  $old = $keep | ConvertFrom-Json
  if ($old.repo) {
    $new = Get-Content -Raw $cfg | ConvertFrom-Json; $new.repo = $old.repo
    [IO.File]::WriteAllText($cfg, ($new | ConvertTo-Json), (New-Object Text.UTF8Encoding $false))   # UTF-8 without BOM
  }
}
Say 'Files copied.'
if ($FilesOnly) { Say 'Files-only mode: done.'; exit 0 }
function ExeOf($n) { Join-Path $InstallDir "$n.exe" }

# ---- shortcuts (and the old "Lumen" ones removed) ----
$shell = New-Object -ComObject WScript.Shell
$programs = [Environment]::GetFolderPath('Programs')
$desk = [Environment]::GetFolderPath('Desktop')
function Link($path, $target, $arguments) {
  $l = $shell.CreateShortcut($path)
  $l.TargetPath = $target; $l.Arguments = $arguments; $l.WorkingDirectory = $InstallDir
  $l.IconLocation = "$target,0"; $l.Description = [IO.Path]::GetFileNameWithoutExtension($path); $l.Save()
}
$hadDesktop = $false
foreach ($a in $APPS) { if (Test-Path (Join-Path $desk "$($a.Old).lnk")) { $hadDesktop = $true } }
Remove-Item -Recurse -Force (Join-Path $programs 'Lumen Office') -ErrorAction SilentlyContinue
foreach ($a in $APPS) { Remove-Item -Force (Join-Path $desk "$($a.Old).lnk") -ErrorAction SilentlyContinue }
Remove-Item -Force (Join-Path $desk 'Lumen Office.lnk') -ErrorAction SilentlyContinue
foreach ($a in $APPS) { Link (Join-Path $programs "$($a.N).lnk") (ExeOf $a.N) '' }
if ($hadDesktop -or (-not $NoDesktop -and (Ask 'Add Write, Cells and Slides to the desktop?' $true))) {
  foreach ($a in $APPS) { Link (Join-Path $desk "$($a.N).lnk") (ExeOf $a.N) '' }
}

# ---- "Open with" for Office files (adds a choice; default apps you picked keep working) ----
$classes = 'HKCU:\Software\Classes'
foreach ($a in $APPS) {
  $exe = ExeOf $a.N; $cmd = "`"$exe`" `"%1`""
  $k = "$classes\$($a.Id)"
  New-Item -Force $k | Out-Null; Set-Item $k $a.Type
  New-Item -Force "$k\DefaultIcon" | Out-Null; Set-Item "$k\DefaultIcon" ((Join-Path $InstallDir "icons\$($a.A).ico") + ',0')
  New-Item -Force "$k\shell\open\command" | Out-Null; Set-Item "$k\shell\open\command" $cmd
  # Windows remembers a default app by program name, so the old "Lumen ..." entry stays and now opens the new program
  foreach ($exeName in "$($a.N).exe", "$($a.Old).exe") {
    $ap = "$classes\Applications\$exeName"
    if ($exeName -eq "$($a.N).exe" -or (Test-Path $ap)) {
      New-Item -Force "$ap\shell\open\command" | Out-Null; Set-Item "$ap\shell\open\command" $cmd
      New-ItemProperty -Force -Path $ap -Name FriendlyAppName -PropertyType String -Value $a.N | Out-Null
      New-Item -Force "$ap\DefaultIcon" | Out-Null; Set-Item "$ap\DefaultIcon" ((Join-Path $InstallDir "icons\$($a.A).ico") + ',0')
    }
  }
  foreach ($e in $a.Ext) {
    New-Item -Force "$classes\$e\OpenWithProgids" | Out-Null
    New-ItemProperty -Force -Path "$classes\$e\OpenWithProgids" -Name $a.Id -PropertyType String -Value '' | Out-Null
  }
}

# one tidy "Open with" entry per app, no stale "Lumen" names or icons
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $InstallDir 'tools\fix_openwith.ps1') -InstallDir $InstallDir

# ---- Installed apps entry ----
New-Item -Force $un | Out-Null
$size = [int]((Get-ChildItem $InstallDir -Recurse -File | Measure-Object Length -Sum).Sum / 1KB)
$vals = @{
  DisplayName = 'Office Apps (Write, Cells, Slides)'; DisplayVersion = $ver; Publisher = 'Office Apps'; InstallLocation = $InstallDir
  DisplayIcon = (Join-Path $InstallDir 'icons\office.ico')
  UninstallString = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$(Join-Path $InstallDir 'Uninstall.ps1')`""
  InstallDate = (Get-Date -Format 'yyyyMMdd')
}
foreach ($k in $vals.Keys) { New-ItemProperty -Force -Path $un -Name $k -PropertyType String -Value $vals[$k] | Out-Null }
foreach ($k in 'NoModify', 'NoRepair') { New-ItemProperty -Force -Path $un -Name $k -PropertyType DWord -Value 1 | Out-Null }
New-ItemProperty -Force -Path $un -Name EstimatedSize -PropertyType DWord -Value $size | Out-Null

# ---- remove an old install in another folder (the former "Lumen Office" folder) ----
if ($oldDir -and ([IO.Path]::GetFullPath($oldDir).TrimEnd('\') -ne [IO.Path]::GetFullPath($InstallDir).TrimEnd('\')) -and (Test-Path (Join-Path $oldDir 'version.json'))) {
  Remove-Item -Recurse -Force $oldDir -ErrorAction SilentlyContinue
  if (Test-Path $oldDir) { Say "Some old files in $oldDir are in use; they can be deleted later." }
}

Say ''
Say "Write, Cells and Slides $ver are installed. Open them from the Start menu or the desktop."
$repo = (Get-Content -Raw $cfg | ConvertFrom-Json).repo
if (-not $repo) { Say "Updates: no GitHub repository is set yet. Put it in $cfg as ""owner/name""." }
