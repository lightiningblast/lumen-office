# Removes Lumen Office for the current user: program files, shortcuts, "Open with" entries and the Installed apps entry.
# Your documents are not touched.
param([switch]$Quiet)
$ErrorActionPreference = 'Continue'
$un = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\LumenOffice'
$dir = (Get-ItemProperty $un -ErrorAction SilentlyContinue).InstallLocation
if (-not $dir) { $dir = Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $Quiet) {
  $a = Read-Host "Remove Lumen Office from $dir ? Your documents are not touched. [y/N]"
  if ($a -notmatch '^(y|yes)$') { exit 0 }
}
# stop the local server
Get-CimInstance Win32_Process -Filter "Name like 'python%'" -ErrorAction SilentlyContinue |
  Where-Object { $_.CommandLine -and $_.CommandLine -like '*serve.py*' -and $_.CommandLine -like "*$dir*" } |
  ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
# shortcuts
Remove-Item -Recurse -Force (Join-Path ([Environment]::GetFolderPath('Programs')) 'Lumen Office') -ErrorAction SilentlyContinue
$desk = [Environment]::GetFolderPath('Desktop')
foreach ($n in 'Lumen Write', 'Lumen Cells', 'Lumen Slides', 'Lumen Office') { Remove-Item -Force (Join-Path $desk "$n.lnk") -ErrorAction SilentlyContinue }
# "Open with" entries
$classes = 'HKCU:\Software\Classes'
foreach ($id in 'LumenOffice.Write', 'LumenOffice.Cells', 'LumenOffice.Slides') { Remove-Item -Recurse -Force "$classes\$id" -ErrorAction SilentlyContinue }
foreach ($e in '.docx', '.xlsx', '.csv', '.pptx') {
  foreach ($id in 'LumenOffice.Write', 'LumenOffice.Cells', 'LumenOffice.Slides') { Remove-ItemProperty -Path "$classes\$e\OpenWithProgids" -Name $id -ErrorAction SilentlyContinue }
}
Remove-Item -Recurse -Force $un -ErrorAction SilentlyContinue
# app data (server secret, update cache, previous-version backup)
Remove-Item -Recurse -Force (Join-Path $env:LOCALAPPDATA 'Lumen Office') -ErrorAction SilentlyContinue
# program files: only delete a folder that really is a Lumen Office install
if ((Test-Path (Join-Path $dir 'version.json')) -and (Test-Path (Join-Path $dir 'tools\serve.py')) -and ($dir -like "*Lumen Office*")) {
  Set-Location $env:TEMP
  Remove-Item -Recurse -Force $dir -ErrorAction SilentlyContinue
}
if (-not $Quiet) { Write-Host 'Lumen Office has been removed.'; Start-Sleep -Seconds 2 }
