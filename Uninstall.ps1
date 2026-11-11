# Removes Write, Cells and Slides for the current user: program files, shortcuts, "Open with" entries and the
# Installed apps entry. Your documents are not touched.
param([switch]$Quiet)
$ErrorActionPreference = 'Continue'
$un = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\LumenOffice'
$dir = (Get-ItemProperty $un -ErrorAction SilentlyContinue).InstallLocation
if (-not $dir) { $dir = Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $Quiet) {
  $a = Read-Host "Remove Write, Cells and Slides from $dir ? Your documents are not touched. [y/N]"
  if ($a -notmatch '^(y|yes)$') { exit 0 }
}
$names = 'Write', 'Cells', 'Slides', 'Office Apps', 'Lumen Write', 'Lumen Cells', 'Lumen Slides', 'Lumen Office'
# close the apps, then stop an old local server if one runs
Get-Process -Name $names -ErrorAction SilentlyContinue | Where-Object { $_.Path -and $_.Path -like "$dir*" } | Stop-Process -Force -ErrorAction SilentlyContinue
Get-CimInstance Win32_Process -Filter "Name like 'python%'" -ErrorAction SilentlyContinue |
  Where-Object { $_.CommandLine -and $_.CommandLine -like '*serve.py*' -and $_.CommandLine -like "*$dir*" } |
  ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
# shortcuts (current and old names)
$programs = [Environment]::GetFolderPath('Programs'); $desk = [Environment]::GetFolderPath('Desktop')
Remove-Item -Recurse -Force (Join-Path $programs 'Lumen Office') -ErrorAction SilentlyContinue
foreach ($n in $names) {
  foreach ($d in $programs, $desk) {
    $l = Join-Path $d "$n.lnk"
    # only our own shortcuts: they point into the install folder
    if (Test-Path $l) { $t = (New-Object -ComObject WScript.Shell).CreateShortcut($l).TargetPath; if ($t -like "$dir*") { Remove-Item -Force $l -ErrorAction SilentlyContinue } }
  }
}
# "Open with" entries
$classes = 'HKCU:\Software\Classes'
foreach ($id in 'LumenOffice.Write', 'LumenOffice.Cells', 'LumenOffice.Slides') { Remove-Item -Recurse -Force "$classes\$id" -ErrorAction SilentlyContinue }
foreach ($e in '.docx', '.xlsx', '.csv', '.pptx') {
  foreach ($id in 'LumenOffice.Write', 'LumenOffice.Cells', 'LumenOffice.Slides') { Remove-ItemProperty -Path "$classes\$e\OpenWithProgids" -Name $id -ErrorAction SilentlyContinue }
}
foreach ($n in $names) {
  $ap = "$classes\Applications\$n.exe"
  $c = (Get-ItemProperty "$ap\shell\open\command" -ErrorAction SilentlyContinue).'(default)'
  if ($c -and $c -like "*$dir*") { Remove-Item -Recurse -Force $ap -ErrorAction SilentlyContinue }
}
Remove-Item -Recurse -Force $un -ErrorAction SilentlyContinue
# app data (settings, window positions, update cache, previous-version backup)
Remove-Item -Recurse -Force (Join-Path $env:LOCALAPPDATA 'Lumen Office') -ErrorAction SilentlyContinue
# program files: only delete a folder that really is this app's install
if ((Test-Path (Join-Path $dir 'version.json')) -and ((Test-Path (Join-Path $dir 'Write.exe')) -or (Test-Path (Join-Path $dir 'Lumen Write.exe')))) {
  Set-Location $env:TEMP
  Remove-Item -Recurse -Force $dir -ErrorAction SilentlyContinue
}
if (-not $Quiet) { Write-Host 'Write, Cells and Slides have been removed.'; Start-Sleep -Seconds 2 }
