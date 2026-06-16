# FetchConvertTools.ps1
# Best-effort downloader for the THIRD-PARTY utilities under .\Convert, driven
# by the Tools.inix manifest. For each tool it resolves the LATEST version and
# (re)installs it only when the installed version differs or the tool is
# missing, recording versions in .\Convert\Tools.lock. Called from
# BuildEdSharp.cmd. NEVER throws: a failed tool is reported and skipped so the
# build always continues.
#
# PANDOC: when Pandoc is installed or upgraded, ModernizePandocConfig.ps1 is run
# to bring the EdSharp.ini / EdSharp.inix conversion flags up to Pandoc 3.x
# (drop -S, markdown_github -> gfm, --reference-docx -> --reference-doc), so the
# newer Pandoc does not break the Convert menu.

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$root     = Split-Path -Parent $MyInvocation.MyCommand.Path
$convert  = Join-Path $root "Convert"
$manifest = Join-Path $root "Tools.inix"
$lockFile = Join-Path $convert "Tools.lock"

if (-not (Test-Path $manifest)) { Write-Host "[tools] Tools.inix not found; nothing to do."; return }

function Read-Manifest($path) {
  $list = @(); $cur = $null
  foreach ($raw in Get-Content -LiteralPath $path) {
    $line = $raw.Trim()
    if ($line -eq "" -or $line.StartsWith(";")) { continue }
    if ($line -match '^\[(.+)\]$') {
      if ($cur) { $list += $cur }
      $cur = @{ name = $Matches[1] }
      continue
    }
    if ($cur -and $line -match '^([^=]+)=(.*)$') { $cur[$Matches[1].Trim()] = $Matches[2].Trim() }
  }
  if ($cur) { $list += $cur }
  return $list
}

function Read-Lock($path) {
  $d = @{}
  if (Test-Path $path) {
    foreach ($raw in Get-Content -LiteralPath $path) {
      if ($raw -match '^([^=]+)=(.*)$') { $d[$Matches[1].Trim()] = $Matches[2].Trim() }
    }
  }
  return $d
}

function Write-Lock($path, $dict) {
  $lines = @(); foreach ($k in ($dict.Keys | Sort-Object)) { $lines += ("{0}={1}" -f $k, $dict[$k]) }
  Set-Content -LiteralPath $path -Value $lines -Encoding ASCII
}

function Version-FromName($name) {
  if ($name -match '([0-9]+(?:\.[0-9]+)+)') { return $Matches[1] }
  return $name
}

# Resolve a tool to @{ url; version } using sfrss (if any), then src.
function Resolve-Latest($t) {
  if ($t.page) {
    try {
      $parts = $t.page -split "\|", 3
      $resp = Invoke-WebRequest -Uri $parts[0] -UseBasicParsing -UserAgent "EdSharp-Build"
      $text = [string]$resp.Content
      if ($text -match $parts[1]) {
        $v = $Matches[1]
        return @{ url = ($parts[2] -replace "\{v\}", $v); version = $v }
      }
      Write-Host ("[tools] {0}: version pattern not found on page; using src." -f $t.name)
    } catch { Write-Host ("[tools] {0}: page lookup failed ({1}); using src." -f $t.name, $_.Exception.Message) }
  }
  if ($t.src -like "github:*") {
    $spec = $t.src.Substring(7); $parts = $spec -split "\|", 2
    $repo = $parts[0]; $pat = $parts[1]
    $rel = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/releases/latest" -Headers @{ "User-Agent" = "EdSharp-Build" } -UseBasicParsing
    $asset = $rel.assets | Where-Object { $_.name -match $pat } | Select-Object -First 1
    if (-not $asset) { throw "no asset matching /$pat/ in latest $repo release" }
    return @{ url = $asset.browser_download_url; version = $rel.tag_name }
  }
  return @{ url = $t.src; version = (Version-FromName ([IO.Path]::GetFileName(($t.src -split "\?")[0]))) }
}

function Find-Exe($base, $exe) {
  if (-not (Test-Path $base)) { return $null }
  return Get-ChildItem -Path $base -Recurse -Filter $exe -ErrorAction SilentlyContinue | Select-Object -First 1
}

function Install-FromUrl($t, $url, $target) {
  $tmp = Join-Path $env:TEMP ("edsharp_" + [guid]::NewGuid().ToString("N"))
  New-Item -ItemType Directory -Path $tmp | Out-Null
  try {
    $zip = Join-Path $tmp ([IO.Path]::GetFileName(($url -split "\?")[0]))
    Invoke-WebRequest -Uri $url -OutFile $zip -UseBasicParsing -UserAgent "EdSharp-Build"
    $fs = [IO.File]::OpenRead($zip); $sig = New-Object byte[] 2; [void]$fs.Read($sig, 0, 2); $fs.Close()
    if (-not ($sig[0] -eq 0x50 -and $sig[1] -eq 0x4B)) { throw "server returned non-ZIP content (likely an HTML interstitial); try a different mirror/URL" }
    $ex = Join-Path $tmp "x"; Expand-Archive -Path $zip -DestinationPath $ex -Force
    $top = @(Get-ChildItem -Path $ex); $srcRoot = $ex
    if ($top.Count -eq 1 -and $top[0].PSIsContainer) { $srcRoot = $top[0].FullName }
    if (Test-Path $target) { Remove-Item (Join-Path $target "*") -Recurse -Force -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Force -Path $target | Out-Null
    if ($t.flatten -eq "true") {
      $found = Get-ChildItem -Path $srcRoot -Recurse -Filter $t.file | Select-Object -First 1
      if (-not $found) { throw "archive did not contain $($t.file)" }
      Copy-Item -Path (Join-Path $found.Directory.FullName "*") -Destination $target -Recurse -Force
    } else {
      Copy-Item -Path (Join-Path $srcRoot "*") -Destination $target -Recurse -Force
    }
  } finally { Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

# Fallback: install via winget or choco (dev machine only), then copy the exe
# found on PATH into Convert\<dir>. Used only when the direct download fails.
function Install-FromPackageManager($t, $target) {
  $mgrs = @()
  if ($t.winget -and (Get-Command winget -ErrorAction SilentlyContinue)) { $mgrs += @{ tool="winget"; args=@("install","--id",$t.winget,"--silent","--accept-package-agreements","--accept-source-agreements") } }
  if ($t.choco  -and (Get-Command choco  -ErrorAction SilentlyContinue)) { $mgrs += @{ tool="choco";  args=@("install",$t.choco,"-y") } }
  foreach ($m in $mgrs) {
    try {
      Write-Host ("[tools] {0}: trying {1} fallback..." -f $t.name, $m.tool)
      & $m.tool @($m.args) | Out-Null
      $cmd = Get-Command $t.file -ErrorAction SilentlyContinue
      if ($cmd) {
        New-Item -ItemType Directory -Force -Path $target | Out-Null
        Copy-Item -Path (Join-Path (Split-Path -Parent $cmd.Source) "*") -Destination $target -Recurse -Force -ErrorAction SilentlyContinue
        if (Find-Exe $target $t.file) { return $true }
      }
    } catch { Write-Host ("[tools] {0}: {1} fallback failed ({2})." -f $t.name, $m.tool, $_.Exception.Message) }
  }
  return $false
}

$tools = Read-Manifest $manifest
$lock  = Read-Lock $lockFile
New-Item -ItemType Directory -Force -Path $convert | Out-Null

foreach ($t in $tools) {
  $target = Join-Path $convert $t.dir
  $have   = Find-Exe $target $t.file
  $want   = $null
  try { $want = Resolve-Latest $t } catch { Write-Host ("[tools] {0}: could not resolve latest ({1})." -f $t.name, $_.Exception.Message) }

  if ($have -and $want -and $lock[$t.name] -eq $want.version) {
    Write-Host ("[tools] {0} up to date ({1})." -f $t.name, $want.version); continue
  }
  if (-not $want) {
    if ($have) { Write-Host ("[tools] {0} present; version check unavailable, keeping existing." -f $t.name) }
    else { Write-Host ("[tools] {0} missing and latest could not be resolved; skipped." -f $t.name) }
    continue
  }

  $action = if ($have) { "upgrading to" } else { "installing" }
  Write-Host ("[tools] {0} {1} {2} (best-effort)..." -f $t.name, $action, $want.version)
  $ok = $false
  try { Install-FromUrl $t $want.url $target; $ok = [bool](Find-Exe $target $t.file) }
  catch { Write-Host ("[tools] {0}: download failed ({1})." -f $t.name, $_.Exception.Message) }
  if (-not $ok) { $ok = Install-FromPackageManager $t $target }

  if ($ok) {
    $lock[$t.name] = $want.version; Write-Lock $lockFile $lock
    Write-Host ("[tools] {0} ready ({1})." -f $t.name, $want.version)
    if ($t.name -eq "Pandoc") {
      $mod = Join-Path $root "ModernizePandocConfig.ps1"
      if (Test-Path $mod) {
        Write-Host "[tools] Pandoc changed; modernizing conversion flags..."
        try { & powershell -NoProfile -ExecutionPolicy Bypass -File $mod -Root $root } catch { Write-Host ("[tools] flag modernizer error: {0}" -f $_.Exception.Message) }
      }
    }
  } else {
    Write-Host ("[tools] {0} could not be installed; install it manually into Convert\{1} or set a winget/choco id in Tools.inix." -f $t.name, $t.dir)
  }
}
