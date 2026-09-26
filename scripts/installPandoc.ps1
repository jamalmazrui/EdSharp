# installPandoc.ps1 -- make sure Pandoc is on this machine, once.
#
# Run by the EdSharp installer's finish page, and by hand from an
# administrator prompt. installPandoc.cmd is the wrapper.
#
# WHAT CHANGED ON 26 SEPTEMBER 2026, AND WHY. This script used to keep a
# private pandoc.exe inside EdSharp's own folder, and when it found one
# already on the machine it COPIED it in. A beta tester who had Pandoc
# through Chocolatey therefore got a second copy, and said so. The Homer
# rule is the opposite: a shared component -- Pandoc, Python, Node, Ollama
# -- installs machine-wide to its own default folder, so every Homer
# program shares one copy and an app upgrade, which replaces the app's
# folder, cannot destroy it. EdSharp's conversions now ask for "pandoc" by
# name and take whatever the PATH gives.
#
# So the whole job is:
#   1. If pandoc answers on the PATH, or sits in a place Windows will put on
#      the PATH, report which one and stop. Nothing is copied anywhere.
#   2. Otherwise install it machine-wide with winget, saying "Downloading"
#      first, because the download takes a while and silence looks like a
#      hang.
#
# Logging: EdSharp is installed under Program Files, so the log goes to
# %LOCALAPPDATA%\EdSharp\logs, one stamped file per run, with the
# environment, every check, every command and its exit code.
#
# Exit codes: 0 Pandoc is available; 1 it is not and could not be installed.

param([switch]$bQuiet)

$c_sWingetId = "JohnMacFarlane.Pandoc"

$sLogDir = Join-Path $env:LOCALAPPDATA "EdSharp\logs"
if (-not (Test-Path -LiteralPath $sLogDir)) { New-Item -ItemType Directory -Path $sLogDir -Force | Out-Null }
$sLogFile = Join-Path $sLogDir ("EdSharp-installPandoc-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".log")

function writeLog($sText) {
  $sLine = (Get-Date -Format "yyyy-MM-dd HH:mm:ss") + "  " + $sText
  Add-Content -LiteralPath $sLogFile -Value $sLine -Encoding UTF8
  if (-not $bQuiet) { Write-Host $sText }
}

function findPandoc() {
  # On the PATH first, which is where a person's own copy answers -- winget,
  # Chocolatey, the official installer and Scoop all put it there.
  $oCommand = Get-Command "pandoc.exe" -ErrorAction SilentlyContinue
  if ($oCommand) { return $oCommand.Source }
  # Then the places an installer puts it that this process's PATH may not
  # yet show, because the PATH was changed after this process started.
  $lPlaces = @(
    (Join-Path $env:ProgramFiles "Pandoc\pandoc.exe"),
    (Join-Path $env:LOCALAPPDATA "Pandoc\pandoc.exe"),
    (Join-Path $env:ProgramData "chocolatey\bin\pandoc.exe"),
    (Join-Path $env:LOCALAPPDATA "Microsoft\WinGet\Links\pandoc.exe")
  )
  foreach ($sPlace in $lPlaces) {
    if (Test-Path -LiteralPath $sPlace) { return $sPlace }
  }
  return ""
}

writeLog "installPandoc started"
writeLog "Script: $($MyInvocation.MyCommand.Path)"
writeLog "PowerShell: $($PSVersionTable.PSVersion), user: $env:USERNAME, quiet: $bQuiet"

$sFound = findPandoc
if ($sFound -ne "") {
  writeLog "Pandoc is already on this machine: $sFound. Nothing to do."
  $sVersion = (& $sFound --version 2>&1 | Select-Object -First 1)
  writeLog "Version: $sVersion"
  exit 0
}

writeLog "Pandoc is not on this machine. Downloading and installing it machine-wide with winget."
Write-Host "Downloading Pandoc. This can take a minute."
$sWinget = Get-Command "winget.exe" -ErrorAction SilentlyContinue
if (-not $sWinget) {
  writeLog "ERROR: winget is not available, so Pandoc could not be installed. Install it from pandoc.org and run EdSharp again."
  exit 1
}
$lOutput = & winget install --id $c_sWingetId --scope machine --silent --accept-source-agreements --accept-package-agreements 2>&1
foreach ($sLine in $lOutput) { writeLog "  winget | $sLine" }
writeLog "winget exit code: $LASTEXITCODE"

$sFound = findPandoc
if ($sFound -ne "") {
  writeLog "Pandoc installed: $sFound"
  exit 0
}
writeLog "ERROR: winget reported $LASTEXITCODE and no pandoc.exe can be found afterwards."
exit 1
