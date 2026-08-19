# installPandoc.ps1 -- fetch pandoc.exe into EdSharp's Convert folder.
#
# pandoc is not packaged inside EdSharp_Setup.exe. It is roughly 200 megabytes,
# GitHub warns about it on every push, and not every EdSharp user converts
# documents. So the installer offers this script on its Finish page instead,
# and the script also works by hand at any later time:
#   installPandoc.cmd   (from the EdSharp installation folder)
#
# What it does, in order:
#   1. If Convert\pandoc.exe already exists beside this script, report and stop.
#   2. If a pandoc.exe is already on this machine (on PATH, or in the usual
#      pandoc install folders), copy that one rather than downloading.
#   3. Otherwise download the newest Windows pandoc from the official GitHub
#      releases and place pandoc.exe into the Convert folder.
#
# The Convert folder lives under Program Files, so writing to it needs an
# elevated process. The installer runs this step elevated. When run by hand,
# run it from an administrator command prompt; the script checks first and
# says so plainly if it cannot write.
#
# Logging: EdSharp is installed under Program Files, so the log cannot sit
# beside the script; it goes to %LOCALAPPDATA%\EdSharp\logs, one timestamped
# file per run, recording the environment, every setting, every action with
# its result, and any error in full.
#
# Exit codes: 0 pandoc is in place (fetched now, copied, or already present);
# 1 something prevented that, and the log says what.

$c_sApiUrl = "https://api.github.com/repos/jgm/pandoc/releases/latest"
$c_sAppName = "EdSharp"
$c_sAssetSuffix = "windows-x86_64.zip"

$sConvertDir = ""
$sLogDir = ""
$sLogFile = ""
$sScriptDir = ""
$sTargetFile = ""

$sScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$sConvertDir = Join-Path $sScriptDir "Convert"
$sTargetFile = Join-Path $sConvertDir "pandoc.exe"
$sLogDir = Join-Path $env:LOCALAPPDATA "$c_sAppName\logs"
$sLogFile = Join-Path $sLogDir ("installPandoc_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".log")

function writeLog($sText) {
  $sStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  $sLine = "$sStamp  $sText"
  Add-Content -LiteralPath $sLogFile -Value $sLine -Encoding UTF8
  Write-Host $sText
}

function findExistingPandoc() {
  # The first hit wins. PATH is checked before the usual folders because a
  # copy on PATH is the one the user has chosen to live with.
  $lCandidates = @()
  $oCommand = Get-Command "pandoc.exe" -ErrorAction SilentlyContinue
  if ($null -ne $oCommand) { $lCandidates += $oCommand.Source }
  $lCandidates += (Join-Path $env:LOCALAPPDATA "Pandoc\pandoc.exe")
  $lCandidates += (Join-Path $env:ProgramFiles "Pandoc\pandoc.exe")
  if ($null -ne ${env:ProgramFiles(x86)}) { $lCandidates += (Join-Path ${env:ProgramFiles(x86)} "Pandoc\pandoc.exe") }
  foreach ($sCandidate in $lCandidates) {
    if (Test-Path -LiteralPath $sCandidate) { return $sCandidate }
  }
  return ""
}

function reportVersion($sExeFile) {
  # First line of pandoc --version, e.g. "pandoc 3.6.3", so the log names the
  # build that ended up installed.
  try {
    $sVersion = (& $sExeFile --version 2>$null | Select-Object -First 1)
    if ($null -ne $sVersion) { writeLog "Installed: $sVersion" }
  } catch {
    writeLog "Version check failed (pandoc is in place; this is informational only): $($_.Exception.Message)"
  }
}

New-Item -ItemType Directory -Path $sLogDir -Force | Out-Null

try {
  # -- Environment, for debugging. ------------------------------------------
  writeLog "installPandoc starting."
  writeLog "Script: $($MyInvocation.MyCommand.Path)"
  writeLog "Command line: $($MyInvocation.Line)"
  writeLog "PowerShell: $($PSVersionTable.PSVersion), platform: $([System.Environment]::OSVersion.VersionString)"
  writeLog "Working directory: $(Get-Location)"
  writeLog "Settings: target=$sTargetFile, api=$c_sApiUrl, asset suffix=$c_sAssetSuffix, log=$sLogFile"

  # -- 1. Already in place? -------------------------------------------------
  if (Test-Path -LiteralPath $sTargetFile) {
    writeLog "pandoc.exe is already present at $sTargetFile. Nothing to do."
    reportVersion $sTargetFile
    exit 0
  }

  # -- Can this process write to the Convert folder? ------------------------
  New-Item -ItemType Directory -Path $sConvertDir -Force -ErrorAction Stop | Out-Null
  $sProbeFile = Join-Path $sConvertDir "installPandoc.probe"
  try {
    Set-Content -LiteralPath $sProbeFile -Value "probe" -ErrorAction Stop
    Remove-Item -LiteralPath $sProbeFile -ErrorAction SilentlyContinue
  } catch {
    writeLog "Cannot write to $sConvertDir. This folder is under Program Files, so the script must run elevated. Run installPandoc.cmd from an administrator command prompt, or rerun the EdSharp installer and check the pandoc box."
    exit 1
  }

  # -- 2. A copy already on this machine? -----------------------------------
  $sExistingFile = findExistingPandoc
  if ($sExistingFile -ne "") {
    writeLog "Found an existing pandoc at $sExistingFile. Copying it instead of downloading."
    Copy-Item -LiteralPath $sExistingFile -Destination $sTargetFile -Force -ErrorAction Stop
    writeLog "Copied to $sTargetFile."
    reportVersion $sTargetFile
    exit 0
  }

  # -- 3. Download the newest release. --------------------------------------
  writeLog "No existing pandoc found. Asking GitHub for the newest release."
  [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
  $oRelease = Invoke-RestMethod -Uri $c_sApiUrl -Headers @{ "User-Agent" = "EdSharp-installPandoc" } -ErrorAction Stop
  writeLog "Newest release: $($oRelease.tag_name)"
  $oAsset = $oRelease.assets | Where-Object { $_.name -like "*$c_sAssetSuffix" } | Select-Object -First 1
  if ($null -eq $oAsset) { writeLog "No asset ending in $c_sAssetSuffix was found in release $($oRelease.tag_name). The release layout may have changed; report this."; exit 1 }
  $nSizeMb = [Math]::Round($oAsset.size / 1MB, 1)
  writeLog "Downloading $($oAsset.name) ($nSizeMb MB). This may take a few minutes."

  $sZipFile = Join-Path $env:TEMP $oAsset.name
  $sUnpackDir = Join-Path $env:TEMP "installPandoc_unpack"
  Invoke-WebRequest -Uri $oAsset.browser_download_url -OutFile $sZipFile -ErrorAction Stop
  writeLog "Download finished: $sZipFile ($([Math]::Round((Get-Item -LiteralPath $sZipFile).Length / 1MB, 1)) MB on disk)."

  if (Test-Path -LiteralPath $sUnpackDir) { Remove-Item -LiteralPath $sUnpackDir -Recurse -Force }
  Expand-Archive -LiteralPath $sZipFile -DestinationPath $sUnpackDir -Force -ErrorAction Stop
  writeLog "Archive unpacked to $sUnpackDir."

  $fPandoc = Get-ChildItem -LiteralPath $sUnpackDir -Recurse -Filter "pandoc.exe" | Select-Object -First 1
  if ($null -eq $fPandoc) { writeLog "pandoc.exe was not inside the downloaded archive. The release layout may have changed; report this."; exit 1 }
  Copy-Item -LiteralPath $fPandoc.FullName -Destination $sTargetFile -Force -ErrorAction Stop
  writeLog "pandoc.exe placed at $sTargetFile."

  Remove-Item -LiteralPath $sZipFile -Force -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath $sUnpackDir -Recurse -Force -ErrorAction SilentlyContinue
  writeLog "Temporary download files removed."

  reportVersion $sTargetFile
  writeLog "installPandoc finished successfully."
  exit 0
} catch {
  writeLog "FAILED: $($_.Exception.Message)"
  writeLog "Details: $($_ | Out-String)"
  writeLog "The log above is at: $sLogFile"
  exit 1
}
