# Tidies the Windows "Open with" list for Write, Cells and Slides:
#   - exactly one entry per app, offered only for its own file types
#   - the entry your chosen default app is attached to is the one that stays (defaults keep working)
#   - stale names/icons Windows cached for the old "Lumen ..." programs are removed (only those of this app)
# Run by Setup; safe to run again.
param([Parameter(Mandatory = $true)][string]$InstallDir)
$ErrorActionPreference = 'Continue'
$cls = 'HKCU:\Software\Classes'
$APPS = @(
  @{ A = 'write';  N = 'Write';  Old = 'Lumen Write';  Id = 'LumenOffice.Write';  Ext = @('.docx') },
  @{ A = 'cells';  N = 'Cells';  Old = 'Lumen Cells';  Id = 'LumenOffice.Cells';  Ext = @('.xlsx', '.csv') },
  @{ A = 'slides'; N = 'Slides'; Old = 'Lumen Slides'; Id = 'LumenOffice.Slides'; Ext = @('.pptx') }
)
function Show($key, $types) {
  Remove-ItemProperty -Path $key -Name NoOpenWith -ErrorAction SilentlyContinue
  Remove-Item -Recurse -Force "$key\SupportedTypes" -ErrorAction SilentlyContinue
  New-Item -Force "$key\SupportedTypes" | Out-Null
  foreach ($e in $types) { New-ItemProperty -Force -Path "$key\SupportedTypes" -Name $e -PropertyType String -Value '' | Out-Null }
}
function Hide($key) { if (Test-Path $key) { New-ItemProperty -Force -Path $key -Name NoOpenWith -PropertyType String -Value '' | Out-Null; Remove-Item -Recurse -Force "$key\SupportedTypes" -ErrorAction SilentlyContinue } }

foreach ($a in $APPS) {
  $exe = Join-Path $InstallDir "$($a.N).exe"
  $legacy = "$cls\Applications\$($a.Old).exe"
  $current = "$cls\Applications\$($a.N).exe"
  if (Test-Path $legacy) {
    # a default app was chosen under the old program name: that entry is the one shown (it opens the new program)
    Show $legacy $a.Ext
    Hide $current
    foreach ($e in $a.Ext) { Remove-ItemProperty -Path "$cls\$e\OpenWithProgids" -Name $a.Id -ErrorAction SilentlyContinue }
  } else {
    # fresh install: the file-type entry is the one shown
    Hide $current
    foreach ($e in $a.Ext) {
      New-Item -Force "$cls\$e\OpenWithProgids" | Out-Null
      New-ItemProperty -Force -Path "$cls\$e\OpenWithProgids" -Name $a.Id -PropertyType String -Value '' | Out-Null
    }
  }
}

# names Windows cached for this app's old program files (paths that no longer exist); other programs are untouched
$mc = 'HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache'
$k = Get-Item $mc -ErrorAction SilentlyContinue
if ($k) {
  foreach ($v in $k.GetValueNames()) {
    $path = $v -replace '\.(FriendlyAppName|ApplicationCompany)$', ''
    $leaf = Split-Path -Leaf $path
    $ours = $leaf -in @('Lumen Write.exe', 'Lumen Cells.exe', 'Lumen Slides.exe', 'Lumen Office.exe', 'Write.exe', 'Cells.exe', 'Slides.exe', 'Office Apps.exe') -and
            ($k.GetValue($v) -match '^(Lumen (Office|Write|Cells|Slides)|Office Apps|Write|Cells|Slides)$')
    if ($ours -and -not (Test-Path -LiteralPath $path)) { Remove-ItemProperty -Path $mc -Name $v -ErrorAction SilentlyContinue }
  }
}

# tell Windows that file associations changed, and refresh its icon cache
Add-Type -Namespace LumenFix -Name Shell -MemberDefinition '[DllImport("shell32.dll")] public static extern void SHChangeNotify(int e, uint f, IntPtr a, IntPtr b);' -ErrorAction SilentlyContinue
try { [LumenFix.Shell]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero) } catch { }
$ie4 = Join-Path $env:WINDIR 'System32\ie4uinit.exe'
if (Test-Path $ie4) { Start-Process $ie4 -ArgumentList '-show' -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue }
