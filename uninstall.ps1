param(
  [switch]$KeepSettings,
  [string]$HomeDir = ""
)

$ErrorActionPreference = "Stop"
$RepoDir = Split-Path -Parent $MyInvocation.MyCommand.Path

function Resolve-NodeBin {
  $candidates = @()
  $cmd = Get-Command node -ErrorAction SilentlyContinue
  if ($cmd) { $candidates += $cmd.Source }
  $agyNode = Join-Path $env:LOCALAPPDATA "Programs\antigravity\resources\app\bin\agy-node.exe"
  if (Test-Path $agyNode) { $candidates += $agyNode }
  foreach ($bin in $candidates) {
    try {
      $ver = (& $bin -p "process.versions.node").Trim()
      $parts = $ver.Split('.') | ForEach-Object { [int]$_ }
      if ($parts[0] -gt 22 -or ($parts[0] -eq 22 -and $parts[1] -ge 13)) {
        return $bin
      }
    } catch {}
  }
  throw "未找到 Node.js >= 22.13.0。"
}

$node = Resolve-NodeBin
$argsList = @(Join-Path $RepoDir "scripts\uninstall.mjs")
if ($KeepSettings) { $argsList += "--keep-settings" }
if ($HomeDir) { $argsList += "--home"; $argsList += $HomeDir }

& $node @argsList
exit $LASTEXITCODE
