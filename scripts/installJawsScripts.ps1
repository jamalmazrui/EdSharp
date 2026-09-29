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
# folder afterward -- so the last act here is writing a two-line result file:
# the exit code, then the log folder path. The Results box reads it to report
# the outcome and to place the setup log beside this one.
#
# That file belongs with the logs, not loose in C:\temp where an earlier
# version left it. It is written to the EdSharp logs folder, and only if
# that cannot be reached -- the rare case of the installer and this script
# running as different Windows accounts -- to a shared folder under the
# shared program data folder, which the installer also checks. Any leftover from the
# old C:\temp location is deleted on the way past.
#
# Arguments (none are required):
#   -bQuiet        say nothing on the console; the log gets everything anyway.
#   -bUninstall    remove the scripts instead of installing them.
#   -pathLogFile   write the log here instead of the EdSharp logs folder;
#                  the uninstaller passes a temporary-folder path, because
#                  the EdSharp logs folder does not survive an uninstall.

param([switch]$bQuiet, [switch]$bUninstall, [string]$pathLogFile = "", [string]$pathResultFile = "",
      [string]$sState = "", [string]$pathStateFile = "")

$sScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# ---- the scripts' fingerprint, and what is installed (29 September 2026) ----
# So the finish page can say Install, Update or Reinstall for the JAWS scripts
# and the NVDA add-on, as it does for every component. The fingerprint is of
# the script sources' names and contents -- never a compiled .jsb -- so it
# changes only when a script does. A successful install writes it into each
# JAWS version's Settings\enu as EdSharp.scripts.fingerprint.
$sMarkerName = "EdSharp.scripts.fingerprint"
function sourcesFingerprint() {
  $oBuffer = New-Object IO.MemoryStream
  $sSourceDir = Join-Path $sScriptDir "jaws"
  foreach ($oFile in @(Get-ChildItem -LiteralPath $sSourceDir -File | Where-Object { @(".jbs", ".jcf", ".jdf", ".jgf", ".jkm", ".jsd", ".jsh", ".jsm", ".jss", ".qs", ".qsm", ".sbl") -contains $_.Extension.ToLowerInvariant() } | Sort-Object { $_.Name.ToLowerInvariant() })) {
    $aName = [Text.Encoding]::UTF8.GetBytes($oFile.Name.ToLowerInvariant())
    $oBuffer.Write($aName, 0, $aName.Length)
    $aBytes = [IO.File]::ReadAllBytes($oFile.FullName)
    $oBuffer.Write($aBytes, 0, $aBytes.Length)
  }
  $oHash = [Security.Cryptography.SHA256]::Create()
  return (($oHash.ComputeHash($oBuffer.ToArray()) | ForEach-Object { $_.ToString("x2") }) -join "")
}
function readManifest([string] $sText) {
  $dValues = @{}
  foreach ($sLine in ($sText -split "`r?`n")) {
    if ($sLine -match '^\s*(\w+)\s*=\s*"?(.*?)"?\s*$') { $dValues[$Matches[1]] = $Matches[2] }
  }
  return $dValues
}
if ($sState -ne "") {
  # One word for the installer: none, install, update or reinstall.
  $sAnswer = "none"
  try {
    if ($sState -eq "jaws") {
      $sRoot = Join-Path $env:APPDATA "Freedom Scientific\JAWS"
      $lsEnu = @()
      if (Test-Path -LiteralPath $sRoot) {
        $lsEnu = @(Get-ChildItem -LiteralPath $sRoot -Directory | ForEach-Object { Join-Path $_.FullName "Settings\enu" } | Where-Object { Test-Path -LiteralPath $_ })
      }
      if ($lsEnu.Count -gt 0) {
        $sPrint = sourcesFingerprint
        $iOurs = 0; $iSame = 0
        foreach ($sEnu in $lsEnu) {
          if (-not (Test-Path -LiteralPath (Join-Path $sEnu "EdSharp.jsb"))) { continue }
          $iOurs += 1
          $sMarker = Join-Path $sEnu $sMarkerName
          if ((Test-Path -LiteralPath $sMarker) -and (([IO.File]::ReadAllText($sMarker)).Trim() -eq $sPrint)) { $iSame += 1 }
        }
        if ($iOurs -eq 0) { $sAnswer = "install" }
        elseif ($iSame -eq $lsEnu.Count) { $sAnswer = "reinstall" }
        else { $sAnswer = "update" }
      }
    } elseif ($sState -eq "nvda") {
      $sAddon = Join-Path (Split-Path -Parent $sScriptDir) "EdSharp.nvda-addon"
      $bNvda = (Test-Path -LiteralPath (Join-Path $env:APPDATA "nvda")) -or
               (Test-Path -LiteralPath (Join-Path ${env:ProgramFiles(x86)} "NVDA\nvda.exe")) -or
               (Test-Path -LiteralPath (Join-Path $env:ProgramFiles "NVDA\nvda.exe"))
      if ($bNvda -and (Test-Path -LiteralPath $sAddon)) {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $oArchive = [IO.Compression.ZipFile]::OpenRead($sAddon)
        try {
          $oReader = New-Object IO.StreamReader($oArchive.GetEntry("manifest.ini").Open())
          $dShipped = readManifest $oReader.ReadToEnd(); $oReader.Dispose()
        } finally { $oArchive.Dispose() }
        $sInstalled = Join-Path $env:APPDATA ("nvda\addons\" + $dShipped["name"] + "\manifest.ini")
        if (-not (Test-Path -LiteralPath $sInstalled)) { $sAnswer = "install" }
        elseif ((readManifest ([IO.File]::ReadAllText($sInstalled)))["version"] -eq $dShipped["version"]) { $sAnswer = "reinstall" }
        else { $sAnswer = "update" }
      }
    }
  } catch {
    $sAnswer = "install"
  }
  Set-Content -LiteralPath $pathStateFile -Value $sAnswer -Encoding ASCII
  exit 0
}
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

$bScrubBlocked = $false

function writeLog($sMessage) {
  $sLine = "{0:yyyy-MM-dd HH:mm:ss}  {1}" -f (Get-Date), $sMessage
  Add-Content -LiteralPath $pathLogFile -Value $sLine -Encoding UTF8
  if (-not $bQuiet) { Write-Host $sMessage }
}

function writeResult($iCode) {
  # The two-line handshake the Results box reads: exit code, then log folder.
  # First choice is the path the installer asked for, then this user's own
  # logs folder, then a shared folder under the public profile for the rare
  # case of two different accounts. Whichever succeeds first wins.
  $lCandidates = @()
  if ($pathResultFile -ne "") { $lCandidates += $pathResultFile }
  $lCandidates += (Join-Path $sLogDir "EdSharp_jaws.result")
  $lCandidates += (Join-Path $env:ProgramData "EdSharp\logs\EdSharp_jaws.result")
  $bWritten = $false
  foreach ($sCandidate in $lCandidates) {
    try {
      New-Item -ItemType Directory -Force -Path (Split-Path -Parent $sCandidate) | Out-Null
      Set-Content -LiteralPath $sCandidate -Value @("$iCode", "$sLogDir") -Encoding ASCII
      writeLog "Result file: $sCandidate"
      $bWritten = $true
      break
    } catch {
      continue
    }
  }
  if (-not $bWritten) { writeLog "WARNING: no result file could be written; the Results box will report the step as not run." }
  # Tidy away the loose file older versions left in C:\temp.
  try {
    if (Test-Path -LiteralPath "C:\temp\EdSharp_jaws.result") { Remove-Item -LiteralPath "C:\temp\EdSharp_jaws.result" -Force -ErrorAction SilentlyContinue }
  } catch { }
}

$iExit = 1
try {
  writeLog "EdSharp JAWS scripts $(if ($bUninstall) { 'removal' } else { 'installation' }) starting."
  writeLog "Script: $($MyInvocation.MyCommand.Path)"
  writeLog "PowerShell: $($PSVersionTable.PSVersion), user: $env:USERNAME"
  writeLog "Arguments: bQuiet=$bQuiet bUninstall=$bUninstall pathLogFile=$pathLogFile"
  # Installed in scripts: the installer script is one level up.
  $sIssFile = Join-Path (Split-Path -Parent $sScriptDir) "EdSharp_setup.iss"
  if (Test-Path -LiteralPath $sIssFile) {
    $matchVersion = [regex]::Match([System.IO.File]::ReadAllText($sIssFile), "(?m)^AppVersion=(.+)$")
    if ($matchVersion.Success) { writeLog "EdSharp version: $($matchVersion.Groups[1].Value.Trim())" }
  }

  # The source layout drives everything: subfolders of Scripts map onto
  # Settings subfolders by name, and root files belong to enu.
  # The JAWS files are in scripts\jaws, beside this script (27 September
  # 2026; the layout before kept them in a Scripts folder beside it).
  $sScriptsDir = Join-Path $sScriptDir "jaws"
  if (-not (Test-Path -LiteralPath $sScriptsDir)) { throw "The jaws folder was not found beside this script: $sScriptsDir" }
  $dBuckets = @{}
  # JAWS FILES ONLY (29 September 2026). An installer of 26 September copied
  # the whole scripts folder into scripts\jaws, and nothing ever removed those
  # files, so every install since copied installPython.cmd, gitPush.cmd and
  # the rest into each JAWS version's settings folder. Only a JAWS file type
  # is copied now, and the strays are removed below.
  $lsJawsTypes = @(".jbs", ".jcf", ".jdf", ".jgf", ".jkm", ".jsd", ".jsh", ".jsm", ".jss", ".qs", ".qsm", ".sbl")
  $lRootFiles = @(Get-ChildItem -LiteralPath $sScriptsDir -File | Where-Object { $lsJawsTypes -contains $_.Extension.ToLowerInvariant() })
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
      if ($sBucket -eq "enu" -and (Test-Path -LiteralPath $sDestDir)) {
        # The strays from before: any file in this settings folder named like
        # one of the program's own scripts, which no JAWS file is.
        foreach ($fileOwn in @(Get-ChildItem -LiteralPath $sScriptDir -File)) {
          $sStray = Join-Path $sDestDir $fileOwn.Name
          if (Test-Path -LiteralPath $sStray) {
            try { Remove-Item -LiteralPath $sStray -Force; writeLog "  removed stray $($fileOwn.Name), copied there by an earlier install" }
            catch { writeLog "  WARNING: could not remove stray $($fileOwn.Name): $($_.Exception.Message)" }
          }
        }
      }
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
        # Scrub the retired Process LaTeX feature from the copies just
        # placed, whatever the source still carries: its F12-family key
        # bindings from every .jkm, and its Script blocks from every
        # .jss. The compile below then rebuilds each .jsb clean, so no
        # install can resurrect the feature and reclaim F12 from the
        # Chat with AI command.
        foreach ($fileMap in @(Get-ChildItem -LiteralPath $sDestDir -File -Filter "*.jkm")) {
          $lKept = @(); $iDropped = 0
          foreach ($sLine in @(Get-Content -LiteralPath $fileMap.FullName)) {
            # Two removals. Any binding whose script name mentions LaTeX in
            # any spelling goes, and so does EVERY F12-family binding
            # whatever it is called: the retired Process LaTeX feature also
            # bound F12 keys under other names, such as a "metrix" command,
            # and those bindings swallow EdSharp's own F12 and Control+F12
            # (Chat with AI and Copy Log).
            if (($sLine -match "=.*late[xc]") -or ($sLine -match "^\s*[^;=]*\bf12\b[^=]*=")) { $iDropped = $iDropped + 1 } else { $lKept += $sLine }
          }
          if ($iDropped -gt 0) {
            # JAWS keeps its key map open while it is running, so this
            # write can fail with a sharing error. That must never fail
            # the installation: the scripts are already copied and
            # working, and the only casualty is a retired key binding.
            # It is retried after a moment, then reported as advice.
            $bWritten = $false
            foreach ($iTry in 1..3) {
              try {
                Set-Content -LiteralPath $fileMap.FullName -Value $lKept -Force -ErrorAction Stop
                $bWritten = $true
                break
              } catch {
                Start-Sleep -Milliseconds 400
              }
            }
            if ($bWritten) { writeLog "  scrubbed $iDropped retired binding(s) from $($fileMap.Name)" }
            else {
              writeLog "  WARNING: $($fileMap.Name) is open in JAWS, so $iDropped retired binding(s) could not be removed."
              writeLog "  Close JAWS, or restart it, and run installJawsScripts.cmd from the EdSharp folder."
              $script:bScrubBlocked = $true
            }
          }
        }
        foreach ($fileScript in @(Get-ChildItem -LiteralPath $sDestDir -File -Filter "*.jss")) {
          $sBody = [System.IO.File]::ReadAllText($fileScript.FullName)
          $sClean = [regex]::Replace($sBody, "(?ims)^[ \t]*Script[ \t]+\w*late[xc]\w*[ \t]*\(.*?^[ \t]*EndScript[ \t]*\r?\n?", "")
          if ($sClean -ne $sBody) {
            try {
              [System.IO.File]::WriteAllText($fileScript.FullName, $sClean)
              writeLog "  scrubbed retired script block(s) from $($fileScript.Name)"
            } catch {
              writeLog "  WARNING: $($fileScript.Name) is open in JAWS, so retired scripts could not be removed: $($_.Exception.Message)"
              $script:bScrubBlocked = $true
            }
          }
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
        # NOT COMPILED IS NOT INSTALLED (29 September 2026): scripts left
        # uncompiled are removed below with the rest.
        writeLog "JAWS $sVersion`: ERROR scompile.exe was not found in the program folder, so the scripts cannot be compiled for this version."
        $iFailed = $iFailed + 1
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
    if ($iFailed -eq 0) {
      # What makes the next installer say Reinstall rather than Update.
      $sPrint = sourcesFingerprint
      foreach ($oVersion in @(Get-ChildItem -LiteralPath (Join-Path $env:APPDATA "Freedom Scientific\JAWS") -Directory -ErrorAction SilentlyContinue)) {
        $sEnu = Join-Path $oVersion.FullName "Settings\enu"
        if (Test-Path -LiteralPath (Join-Path $sEnu "EdSharp.jsb")) {
          Set-Content -LiteralPath (Join-Path $sEnu $sMarkerName) -Value $sPrint -Encoding ASCII
        }
      }
      writeLog "Fingerprint written for the next installer: $sPrint"
    }
    # IF ANY SCRIPT DID NOT COMPILE, NOTHING IS LEFT INSTALLED (29 September
    # 2026). This script's own removal -- the one the uninstaller runs -- takes
    # out every file it copies and every .jsb compiled from them, in every JAWS
    # version, so no version is left with scripts that half load. Exit code 2
    # tells the Results box.
    if ($iFailed -gt 0) {
      writeLog "ERROR: not every script compiled, so none are left installed. Removing what this run placed."
      & powershell -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath -bUninstall -bQuiet -pathLogFile $pathLogFile 2>&1 | Out-Null
      writeLog "run exit=$LASTEXITCODE cmd=""installJawsScripts -bUninstall"""
    }
  }
} catch {
  writeLog "FAILED: $($_.Exception.Message)"
  writeLog "At: $($_.InvocationInfo.PositionMessage)"
  $iExit = 1
}
if (-not $bUninstall) { writeResult $iExit }
writeLog "Done. Exit code $iExit. The log is at $pathLogFile"
exit $iExit
