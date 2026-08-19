# installJawsScripts.ps1 -- install (or remove) EdSharp's JAWS settings family.
#
# Modeled on HomerView's installer script, which does this job for a more
# complex script set without a single message box: everything goes to the log,
# the installer's closing Results box summarizes, and this script talks to
# that box through a small result file. EdSharp.exe has no part in this;
# installing screen reader scripts is the installer's job, not the editor's.
#
# What it does, per installed JAWS version found under the user's roaming
# application data (Freedom Scientific\JAWS\<version> with a Settings folder):
#   - Copies each SUBFOLDER of {app}\Scripts into Settings\<same subfolder>
#     (files sitting at the root of Scripts go to Settings\enu).
#   - Compiles every .jss placed in Settings\enu with that JAWS version's
#     scompile.exe, so the .jsb is built where JAWS loads it.
#   - With -bUninstall, removes the files it would have copied, plus the
#     .jsb compiled from each .jss.
#
# The installer runs this AS THE ORIGINAL USER (runasoriginaluser), because
# JAWS keeps its settings in the user's own profile and the installer itself
# is elevated. That also means the installer cannot read this user's log
# folder afterward -- so the last act here is writing a two-line result file
# to C:\temp\EdSharp_jaws.result: the exit code, then the log folder path.
# The Results box reads it to report the outcome and to place the setup log
# beside this one.
#
# Arguments (none are required):
#   -bQuiet        say nothing on the console; the log gets everything anyway.
#   -bUninstall    remove the scripts instead of installing them.
#   -pathLogFile   write the log here instead of the EdSharp logs folder;
#                  the uninstaller passes a temporary-folder path, because
#                  the EdSharp logs folder does not survive an uninstall.

param([switch]$bQuiet, [switch]$bUninstall, [string]$pathLogFile = "")

$c_sResultFile = "C:\temp\EdSharp_jaws.result"

$sScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$sLogDir = Join-Path $env:LOCALAPPDATA "EdSharp\logs"
if ($pathLogFile -eq "") {
  New-Item -ItemType Directory -Force -Path $sLogDir | Out-Null
  # The consolidated log: pandoc, JAWS, and Inno Setup all append to this
  # one file under dated banners. -pathLogFile still redirects it, which the
  # uninstaller uses because this folder does not survive an uninstall.
  $pathLogFile = Join-Path $sLogDir "EdSharp_setup.log"
} else {
  $sLogDir = Split-Path -Parent $pathLogFile
}
Add-Content -LiteralPath $pathLogFile -Value "" -Encoding UTF8
Add-Content -LiteralPath $pathLogFile -Value ("==== installJawsScripts  " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss") + " ====") -Encoding UTF8

function writeLog($sMessage) {
  $sLine = "{0:yyyy-MM-dd HH:mm:ss}  {1}" -f (Get-Date), $sMessage
  Add-Content -LiteralPath $pathLogFile -Value $sLine -Encoding UTF8
  if (-not $bQuiet) { Write-Host $sMessage }
}

function writeResult($iCode) {
  # The two-line handshake the Results box reads: exit code, then log folder.
  try {
    New-Item -ItemType Directory -Force -Path "C:\temp" | Out-Null
    Set-Content -LiteralPath $c_sResultFile -Value @("$iCode", "$sLogDir") -Encoding ASCII
  } catch {
    writeLog "WARNING: the result file could not be written: $($_.Exception.Message)"
  }
}

$iExit = 1
try {
  writeLog "EdSharp JAWS scripts $(if ($bUninstall) { 'removal' } else { 'installation' }) starting."
  writeLog "Script: $($MyInvocation.MyCommand.Path)"
  writeLog "PowerShell: $($PSVersionTable.PSVersion), user: $env:USERNAME"
  writeLog "Arguments: bQuiet=$bQuiet bUninstall=$bUninstall pathLogFile=$pathLogFile"
  $sIssFile = Join-Path $sScriptDir "EdSharp_Setup.iss"
  if (Test-Path -LiteralPath $sIssFile) {
    $matchVersion = [regex]::Match([System.IO.File]::ReadAllText($sIssFile), "(?m)^AppVersion=(.+)$")
    if ($matchVersion.Success) { writeLog "EdSharp version: $($matchVersion.Groups[1].Value.Trim())" }
  }

  # The source layout drives everything: subfolders of Scripts map onto
  # Settings subfolders by name, and root files belong to enu.
  $sScriptsDir = Join-Path $sScriptDir "Scripts"
  if (-not (Test-Path -LiteralPath $sScriptsDir)) { throw "The Scripts folder was not found beside this script: $sScriptsDir" }
  $dBuckets = @{}
  $lRootFiles = @(Get-ChildItem -LiteralPath $sScriptsDir -File)
  if ($lRootFiles.Count -gt 0) { $dBuckets["enu"] = $lRootFiles }
  foreach ($folderSub in @(Get-ChildItem -LiteralPath $sScriptsDir -Directory)) {
    $lSubFiles = @(Get-ChildItem -LiteralPath $folderSub.FullName -File)
    if ($lSubFiles.Count -gt 0) {
      if ($dBuckets.ContainsKey($folderSub.Name)) { $dBuckets[$folderSub.Name] = @($dBuckets[$folderSub.Name]) + $lSubFiles }
      else { $dBuckets[$folderSub.Name] = $lSubFiles }
    }
  }
  foreach ($sBucket in ($dBuckets.Keys | Sort-Object)) {
    writeLog "Source bucket $sBucket`: $($dBuckets[$sBucket].Count) files"
  }
  if ($dBuckets.Count -eq 0) { throw "The Scripts folder is empty; there is nothing to install." }

  $sJawsRoot = Join-Path $env:APPDATA "Freedom Scientific\JAWS"
  $lVersions = @()
  if (Test-Path -LiteralPath $sJawsRoot) {
    $lVersions = @(Get-ChildItem -LiteralPath $sJawsRoot -Directory | Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName "Settings") })
  }
  writeLog "JAWS versions with settings for $env:USERNAME`: $($lVersions.Count)"
  if ($lVersions.Count -eq 0) { throw "No JAWS version with a Settings folder was found under $sJawsRoot." }

  $iCopied = 0
  $iCompiled = 0
  $iRemoved = 0
  $iFailed = 0
  foreach ($folderVersion in $lVersions) {
    $sVersion = $folderVersion.Name
    $sSettingsDir = Join-Path $folderVersion.FullName "Settings"
    foreach ($sBucket in ($dBuckets.Keys | Sort-Object)) {
      $sDestDir = Join-Path $sSettingsDir $sBucket
      if ($bUninstall) {
        $iBucketRemoved = 0
        foreach ($fileSource in $dBuckets[$sBucket]) {
          foreach ($sName in @($fileSource.Name) + $(if ($fileSource.Extension -ieq ".jss") { @([System.IO.Path]::ChangeExtension($fileSource.Name, ".jsb")) } else { @() })) {
            $sTarget = Join-Path $sDestDir $sName
            if (Test-Path -LiteralPath $sTarget) {
              Remove-Item -LiteralPath $sTarget -Force
              $iRemoved = $iRemoved + 1
              $iBucketRemoved = $iBucketRemoved + 1
              writeLog "  removed $sTarget"
            }
          }
        }
        writeLog "JAWS $sVersion / $sBucket`: removed $iBucketRemoved"
      } else {
        New-Item -ItemType Directory -Force -Path $sDestDir | Out-Null
        foreach ($fileSource in $dBuckets[$sBucket]) {
          Copy-Item -LiteralPath $fileSource.FullName -Destination (Join-Path $sDestDir $fileSource.Name) -Force
          $iCopied = $iCopied + 1
          writeLog "  copied $($fileSource.Name) -> $sDestDir"
        }
        writeLog "JAWS $sVersion / $sBucket`: done"
      }
    }
    if (-not $bUninstall) {
      # Compile with THIS version's scompile, so the .jsb format matches the
      # JAWS that will load it. The program folder is machine-wide, so the
      # version folder name there matches the settings folder name.
      $sCompile = ""
      foreach ($sProgramRoot in @($env:ProgramFiles, ${env:ProgramFiles(x86)})) {
        if ($sProgramRoot) {
          $sCandidate = Join-Path $sProgramRoot "Freedom Scientific\JAWS\$sVersion\scompile.exe"
          if (Test-Path -LiteralPath $sCandidate) { $sCompile = $sCandidate; break }
        }
      }
      if ($sCompile -eq "") {
        writeLog "JAWS $sVersion`: scompile.exe was not found in the program folder; the .jss files are copied but not compiled. JAWS compiles a script itself when it is next loaded from the script manager."
      } else {
        writeLog "JAWS $sVersion`: compiler $sCompile"
        $sEnuDir = Join-Path $sSettingsDir "enu"
        foreach ($fileScript in @(Get-ChildItem -LiteralPath $sEnuDir -File -Filter "*.jss" | Where-Object { $sName = $_.Name; ($dBuckets.Values | ForEach-Object { $_ } | Where-Object { $_.Name -ieq $sName }).Count -gt 0 })) {
          & $sCompile $fileScript.FullName | Out-Null
          if ($LASTEXITCODE -eq 0) {
            $iCompiled = $iCompiled + 1
            writeLog "  compiled $($fileScript.Name)"
          } else {
            $iFailed = $iFailed + 1
            writeLog "  FAILED to compile $($fileScript.Name) (exit $LASTEXITCODE)"
          }
        }
      }
    }
  }

  if ($bUninstall) {
    writeLog "EdSharp JAWS scripts: $iRemoved removed."
    $iExit = 0
  } else {
    writeLog "EdSharp JAWS scripts: $iCopied copied, $iCompiled compiled$(if ($iFailed -gt 0) { ", $iFailed FAILED" })."
    $iExit = $(if ($iFailed -gt 0) { 2 } else { 0 })
  }
} catch {
  writeLog "FAILED: $($_.Exception.Message)"
  writeLog "At: $($_.InvocationInfo.PositionMessage)"
  $iExit = 1
}
if (-not $bUninstall) { writeResult $iExit }
writeLog "Done. Exit code $iExit. The log is at $pathLogFile"
exit $iExit
