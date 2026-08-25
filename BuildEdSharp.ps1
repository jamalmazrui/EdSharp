# BuildEdSharp.ps1 -- build EdSharp 5.0 for .NET Framework 4.8, AnyCPU.
#
# Replaces the earlier BuildEdSharp.cmd batch logic, which died silently inside
# its ReverseMarkdown fetch (the log ended with no ERROR line -- the classic
# signature of a cmd parse casualty in a caret-continued PowerShell block).
# All logic now lives here in PowerShell, where every step runs inside
# try/catch and CANNOT fail without the reason landing in the log.
#
# What it does, in order:
#   1. Starts a fresh, detailed log: BuildEdSharp.log beside this script.
#   2. Records the environment and every effective setting.
#   3. Fetches NuGet dependencies that are not already present:
#      ReverseMarkdown 4.7.1 (and HtmlAgilityPack, which ReverseMarkdown 4.x
#      needs at run time). Each download is verified to really be a zip archive
#      before unpacking, so an error page saved as a file is caught and named.
#   4. Fetches nvdaControllerClient.dll if missing (optional; a warning, not
#      a failure, since the iss ships it with skipifsourcedoesntexist).
#   5. Finds a C# compiler: Roslyn csc from Visual Studio / Build Tools first
#      (via vswhere), then the .NET Framework 4 csc as a logged fallback
#      (which supports only C# 5 -- the log says which one it chose).
#   6. Compiles EdSharp.dll from the support sources, then EdSharp.exe from
#      EdSharp.cs, logging the full compiler command line and all output.
#   7. Compiles EdSharp_Setup.exe with Inno Setup's ISCC when Inno Setup 6
#      is installed, following the buildHomerScribe pattern: a machine
#      without Inno Setup still finishes successfully, with a note saying
#      how to compile the installer by hand.
#   8. Ends with a report of every artifact and its size, and SUCCESS or
#      FAILED as the last word of the log.
#
# ASSUMPTIONS, marked so they are easy to correct:
#   A1. EdSharp.dll is built from: Lbc.cs, Say.cs, Inix.cs, KeyMap.cs, Web.cs
#       (whichever exist; missing ones are logged and skipped).
#   A2. EdSharp.exe is built from EdSharp.cs, referencing EdSharp.dll,
#       Tektosyne.dll, Ude.dll, ReverseMarkdown.dll, HtmlAgilityPack.dll.
#   A3. Target: .NET Framework 4.8, AnyCPU, /target:winexe, with
#       EdSharp.manifest and EdSharp.ico applied when present.
# If any assumption is wrong, the log will show exactly which compile step
# failed and with what message; report that back and the fix is small.
#
# Exit codes: 0 build succeeded; 1 build failed (see BuildEdSharp.log).

param([string]$sMode = "")

# "buildEdSharp console" makes a TEMPORARY DEBUGGING BUILD: the exe becomes a
# console program, so the class of failure that kills a windowed EdSharp in
# silence -- a type initializer or assembly that fails before the first line
# of Main -- prints itself to the command window instead. Console mode also
# implies nobump (debug builds must not burn release numbers) and skips the
# installer step (a console EdSharp must never ship). Run plain buildEdSharp
# afterward to restore the real, windowed build.
$bConsole = ($sMode -ieq "console")

# ---- constants --------------------------------------------------------------
$c_sHtmlAgilityPackVersion = "1.11.72"
# Markdig is pinned at the latest release (decision of 24 August 2026):
# the 1.x line still targets netstandard2.0, which .NET Framework 4.8
# loads through the netstandard facade, so the newest Markdig runs in the
# EdSharp binary. The core API used here (Markdown.ToHtml, ToPlainText,
# MarkdownPipelineBuilder.UseAdvancedExtensions) is unchanged from 0.x.
$c_sMarkdigVersion = "1.3.2"
$c_sNvdaClientUrl = "https://download.nvaccess.org/releases/stable/nvda_2026.1_controllerClient.zip"
# ReverseMarkdown is pinned at the last line that supports .NET Framework:
# the 4.x series ships netstandard2.0, which net48 consumes; 5.x and 6.x ship
# only net8.0 and later, which net48 cannot reference at all (proved by the
# 19 Aug 2026 build log). EdSharp.cs only ever compiled against a 4.x dll, so
# this is the API it is written to. Do not raise this to 5.x or 6.x while
# EdSharp targets .NET Framework 4.8.
$c_sReverseMarkdownVersion = "4.7.1"

# The support sources; since the name-collision lesson they compile INTO
# EdSharp.exe rather than into a separate EdSharp.dll.
$c_lDllSources = @("Lbc.cs", "Say.cs", "Inix.cs", "KeyMap.cs", "Web.cs")
$c_lExeReferences = @("System.dll", "System.Core.dll", "System.Data.dll", "System.Drawing.dll", "System.IO.Compression.dll", "System.Web.dll", "System.Windows.Forms.dll", "System.Xml.dll", "Microsoft.VisualBasic.dll")
# UI Automation, which the support sources use (System.Windows.Automation and
# the provider interfaces such as IRawElementProviderSimple). On a machine
# without the .NET Framework Developer Pack these assemblies live in the WPF
# subfolder of the Framework runtime directory, so the compile steps add a
# /lib search path covering both the runtime directory and its WPF subfolder.
$c_lUiaReferences = @("UIAutomationClient.dll", "UIAutomationProvider.dll", "UIAutomationTypes.dll", "WindowsBase.dll")
$c_lLibTargetPreference = @("net48", "net472", "net462", "net461", "net46", "net45", "netstandard2.0")

# ---- paths and log ----------------------------------------------------------
$sScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$sLogFile = Join-Path $sScriptDir "BuildEdSharp.log"
Set-Content -LiteralPath $sLogFile -Value "" -Encoding UTF8

$bFailed = $false

function writeLog($sText) {
  $sStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  Add-Content -LiteralPath $sLogFile -Value "$sStamp  $sText" -Encoding UTF8
  Write-Host $sText
}

function runTool($sExeFile, $lArguments, $sPurpose) {
  # Runs one external command, logging the verbatim command line, every line
  # of its output, and its exit code. Returns the exit code.
  writeLog "RUN ($sPurpose): `"$sExeFile`" $($lArguments -join " ")"
  $sOutput = & $sExeFile @lArguments 2>&1 | Out-String
  $iExit = $LASTEXITCODE
  if ($sOutput.Trim() -ne "") { Add-Content -LiteralPath $sLogFile -Value $sOutput -Encoding UTF8; Write-Host $sOutput }
  writeLog "EXIT ($sPurpose): $iExit"
  return $iExit
}

function fetchNugetPackage($sPackageId, $sVersion) {
  # Downloads a nupkg from the official NuGet flat-container url, VERIFIES the
  # file is really a zip (an error body saved to disk is the failure that cost
  # an evening on the GetAudibleInfo build), unpacks it, and copies the best
  # lib DLLs for .NET Framework 4.8 into the script folder. Throws on failure;
  # the caller's catch puts the details in the log.
  $sIdLower = $sPackageId.ToLower()
  $sUrl = "https://api.nuget.org/v3-flatcontainer/$sIdLower/$sVersion/$sIdLower.$sVersion.nupkg"
  $sZipFile = Join-Path $env:TEMP "$sIdLower.$sVersion.zip"
  $sUnpackDir = Join-Path $env:TEMP "$sIdLower.$sVersion.unpack"

  writeLog "Fetching $sPackageId $sVersion from $sUrl"
  Invoke-WebRequest -Uri $sUrl -OutFile $sZipFile -UseBasicParsing -ErrorAction Stop
  $fZip = Get-Item -LiteralPath $sZipFile
  writeLog "Downloaded $($fZip.Length) bytes to $sZipFile"

  # A real zip begins with the bytes PK. Anything else is an error body.
  $binHeader = [System.IO.File]::ReadAllBytes($sZipFile)[0..1]
  if (-not ($binHeader[0] -eq 0x50 -and $binHeader[1] -eq 0x4B)) { throw "$sPackageId download is not a zip archive (first bytes are not PK). The url may be wrong or the server returned an error page: $sUrl" }
  writeLog "Zip signature verified for $sPackageId."

  if (Test-Path -LiteralPath $sUnpackDir) { Remove-Item -LiteralPath $sUnpackDir -Recurse -Force }
  Expand-Archive -LiteralPath $sZipFile -DestinationPath $sUnpackDir -Force -ErrorAction Stop
  writeLog "Unpacked to $sUnpackDir"

  $sChosenDir = ""
  foreach ($sTarget in $c_lLibTargetPreference) {
    $sCandidateDir = Join-Path $sUnpackDir "lib\$sTarget"
    if (Test-Path -LiteralPath $sCandidateDir) { $sChosenDir = $sCandidateDir; break }
  }
  if ($sChosenDir -eq "") {
    $lFound = Get-ChildItem -LiteralPath (Join-Path $sUnpackDir "lib") -Directory -ErrorAction SilentlyContinue | ForEach-Object { $_.Name }
    throw "$sPackageId $sVersion has no lib target usable for .NET Framework 4.8. Targets present: $($lFound -join ", ")"
  }
  writeLog "Chose lib target: $sChosenDir"

  foreach ($fDll in Get-ChildItem -LiteralPath $sChosenDir -Filter "*.dll") {
    # A running EdSharp loads these libraries from this folder, and a locked
    # dll fails the copy with an unhelpful IOException. Name the culprit and
    # the remedy instead.
    $lLockers = @(Get-Process -Name EdSharp, ijs -ErrorAction SilentlyContinue | Where-Object { try { (Split-Path $_.Path -Parent) -ieq $sScriptDir } catch { $false } })
    if ($lLockers.Count -gt 0) {
      $sWho = ($lLockers | ForEach-Object { "$($_.ProcessName) (pid $($_.Id))" }) -join ", "
      throw "Cannot replace $($fDll.Name): $sWho is running from $sScriptDir and holds it. Close it and run the build again."
    }
    Copy-Item -LiteralPath $fDll.FullName -Destination (Join-Path $sScriptDir $fDll.Name) -Force
    writeLog "Copied $($fDll.Name) ($($fDll.Length) bytes) to the build folder."
  }
  Remove-Item -LiteralPath $sZipFile -Force -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath $sUnpackDir -Recurse -Force -ErrorAction SilentlyContinue
}

function findCsc() {
  # Roslyn csc from Visual Studio or Build Tools first (modern C#), located
  # through vswhere; the .NET Framework csc (C# 5 only) as a logged fallback.
  $sVswhereFile = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
  if (Test-Path -LiteralPath $sVswhereFile) {
    $sInstallDir = & $sVswhereFile -latest -products * -requires Microsoft.Component.MSBuild -property installationPath 2>$null | Select-Object -First 1
    if ($null -ne $sInstallDir -and $sInstallDir -ne "") {
      $sRoslynFile = Join-Path $sInstallDir "MSBuild\Current\Bin\Roslyn\csc.exe"
      if (Test-Path -LiteralPath $sRoslynFile) { writeLog "Compiler: Roslyn csc via vswhere: $sRoslynFile"; return $sRoslynFile }
    }
  }
  $sFrameworkFile = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
  if (Test-Path -LiteralPath $sFrameworkFile) { writeLog "Compiler: .NET Framework csc (C# 5 ONLY -- if the sources use newer syntax, install Visual Studio Build Tools): $sFrameworkFile"; return $sFrameworkFile }
  throw "No C# compiler found. Install Visual Studio Build Tools 2022 (https://visualstudio.microsoft.com/downloads/) or repair the .NET Framework."
}

function compareVersions($sLeft, $sRight) {
  # -1, 0, or 1, comparing dotted numbers part by part with missing parts as
  # zero. A part that is not a number makes its version compare as older,
  # so an odd tag can never become the base to count from.
  $lLeft = $sLeft.Split(".")
  $lRight = $sRight.Split(".")
  $iCount = [Math]::Max($lLeft.Count, $lRight.Count)
  for ($iPart = 0; $iPart -lt $iCount; $iPart++) {
    $iLeft = 0
    $iRight = 0
    if ($iPart -lt $lLeft.Count) { if (-not [int]::TryParse($lLeft[$iPart], [ref]$iLeft)) { return -1 } }
    if ($iPart -lt $lRight.Count) { if (-not [int]::TryParse($lRight[$iPart], [ref]$iRight)) { return 1 } }
    if ($iLeft -lt $iRight) { return -1 }
    if ($iLeft -gt $iRight) { return 1 }
  }
  return 0
}

function nextVersion($sCurrent) {
  # The last dotted part goes up by one; fewer than three parts gain one,
  # so 5.0 becomes 5.0.1 and a bare 7 becomes 7.0.1 (HomerScribe's rule).
  $lParts = $sCurrent.Split(".")
  if ($lParts.Count -ge 3) {
    $lParts[$lParts.Count - 1] = [string]([int]$lParts[$lParts.Count - 1] + 1)
    return ($lParts -join ".")
  }
  if ($lParts.Count -eq 2) { return "$sCurrent.1" }
  return "$sCurrent.0.1"
}

function findJsc() {
  # The JScript .NET compiler ships inside the .NET Framework itself, 64-bit
  # first. It builds the runtime evaluator, so a machine without it needs a
  # Framework repair, not a workaround.
  $sJscFile = Join-Path $env:SystemRoot "Microsoft.NET\Framework64\v4.0.30319\jsc.exe"
  if (Test-Path -LiteralPath $sJscFile) { return $sJscFile }
  $sJscFile = Join-Path $env:SystemRoot "Microsoft.NET\Framework\v4.0.30319\jsc.exe"
  if (Test-Path -LiteralPath $sJscFile) { return $sJscFile }
  return ""
}

function findIscc() {
  # Inno Setup 6's command-line compiler, in the two standard locations,
  # exactly as buildHomerScribe.cmd looks for it. Returns "" when absent,
  # which is a note rather than a failure.
  $sIsccFile = Join-Path ${env:ProgramFiles(x86)} "Inno Setup 6\ISCC.exe"
  if (Test-Path -LiteralPath $sIsccFile) { return $sIsccFile }
  $sIsccFile = Join-Path $env:ProgramFiles "Inno Setup 6\ISCC.exe"
  if (Test-Path -LiteralPath $sIsccFile) { return $sIsccFile }
  return ""
}

function findNetstandardFacade() {
  # ReverseMarkdown 5.x ships as netstandard2.0; compiling a net48 program
  # against it needs the netstandard facade reference. Run time needs nothing
  # extra on 4.8. Returns "" when not found, and the compile step warns.
  $sFacadeFile = Join-Path ${env:ProgramFiles(x86)} "Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8\Facades\netstandard.dll"
  if (Test-Path -LiteralPath $sFacadeFile) { return $sFacadeFile }
  $sFacadeFile = Join-Path ${env:ProgramFiles(x86)} "Reference Assemblies\Microsoft\Framework\.NETFramework\v4.7.2\Facades\netstandard.dll"
  if (Test-Path -LiteralPath $sFacadeFile) { return $sFacadeFile }
  # The Framework's own copies, present on every 4.7.1+ machine even without
  # any Developer Pack: the GAC entry first, then the runtime Facades folder.
  $sFacadeFile = Join-Path $env:WINDIR "Microsoft.NET\assembly\GAC_MSIL\netstandard\v4.0_2.0.0.0__cc7b13ffcd2ddd51\netstandard.dll"
  if (Test-Path -LiteralPath $sFacadeFile) { return $sFacadeFile }
  $sFacadeFile = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\Facades\netstandard.dll"
  if (Test-Path -LiteralPath $sFacadeFile) { return $sFacadeFile }
  return ""
}

# ---- main -------------------------------------------------------------------
try {
  Set-Location -LiteralPath $sScriptDir

  writeLog "EdSharp build starting."
  writeLog "Script: $($MyInvocation.MyCommand.Path)"
  writeLog "Command line: $($MyInvocation.Line)"
  writeLog "PowerShell: $($PSVersionTable.PSVersion), platform: $([System.Environment]::OSVersion.VersionString), 64-bit process: $([System.Environment]::Is64BitProcess)"
  writeLog "Working directory: $(Get-Location)"
  writeLog "Settings: ReverseMarkdown=$c_sReverseMarkdownVersion, HtmlAgilityPack=$c_sHtmlAgilityPackVersion, Markdig=$c_sMarkdigVersion, log=$sLogFile"

  # ---- version: bump the iss AppVersion past every released tag ----------
  # EdSharp's version lives in EdSharp_Setup.iss (the plain AppVersion=
  # directive; that file is the single source, playing the role version.txt
  # plays for the other Homer Tools). v5.0 through v5.0.10 are already
  # released, so a build that leaves the number alone hands tagRelease a
  # version it must refuse. This is HomerScribe's takeNextVersion pattern:
  # increment the last dotted part, step over any number that already
  # carries a v-tag on origin, and rewrite the version lines in the iss.
  # Run "buildEdSharp nobump" to keep the current number.
  $sIssFile = Join-Path $sScriptDir "EdSharp_Setup.iss"
  if (-not (Test-Path -LiteralPath $sIssFile)) { throw "EdSharp_Setup.iss was not found beside the build script." }
  $sIss = [System.IO.File]::ReadAllText($sIssFile)
  $matchVersion = [regex]::Match($sIss, "(?m)^AppVersion=(.+)$")
  if (-not $matchVersion.Success) { throw "No AppVersion= line was found in EdSharp_Setup.iss." }
  $sOldVersion = $matchVersion.Groups[1].Value.Trim()
  if ($sMode -ieq "nobump" -or $bConsole) {
    writeLog "Version: $sOldVersion (nobump: keeping the current number)"
  } else {
    $dReleased = @{}
    $sRemoteTags = (& git ls-remote --tags origin "v*" 2>&1 | Out-String)
    if ($LASTEXITCODE -eq 0) {
      foreach ($sLine in $sRemoteTags -split "`n") {
        if ($sLine -match "refs/tags/v([^\^\s]+)\s*$") { $dReleased[$Matches[1]] = $true }
      }
      writeLog "Released versions on origin: $($dReleased.Count)"
    } else {
      writeLog "WARNING: the released tags could not be read, so the next number is taken blindly. tagRelease remains the final check."
    }
    # Start from the HIGHEST released version or the iss version, whichever
    # is greater, then increment. Plain step-over would fill gaps: the real
    # tag list skips v5.0.3, and minting 5.0.3 after v5.0.10 exists would
    # put a "new" release below an old one.
    $sBase = $sOldVersion
    foreach ($sReleased in $dReleased.Keys) {
      if ((compareVersions $sReleased $sBase) -gt 0) { $sBase = $sReleased }
    }
    if ($sBase -ne $sOldVersion) { writeLog "Highest released version is v$sBase, above the iss's $sOldVersion; counting from there." }
    $sNewVersion = nextVersion $sBase
    $iGuard = 0
    while ($dReleased.ContainsKey($sNewVersion) -and $iGuard -lt 200) {
      writeLog "Version v$sNewVersion is already released; stepping over it."
      $sNewVersion = nextVersion $sNewVersion
      $iGuard = $iGuard + 1
    }
    $sIss = [regex]::Replace($sIss, "(?m)^AppVersion=.*$", "AppVersion=$sNewVersion")
    $sIss = [regex]::Replace($sIss, "(?m)^VersionInfoVersion=.*$", "VersionInfoVersion=$sNewVersion")
    $sIss = [regex]::Replace($sIss, "(?m)^(AppVerName=.*?)" + [regex]::Escape($sOldVersion) + "(.*)$", "`${1}$sNewVersion`${2}")
    [System.IO.File]::WriteAllText($sIssFile, $sIss, (New-Object System.Text.UTF8Encoding($true)))
    writeLog "Version: $sOldVersion -> $sNewVersion (written to EdSharp_Setup.iss; tagRelease will tag v$sNewVersion)"
  }

  [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
  writeLog "TLS 1.2 enabled for downloads."

  # ---- 3. NuGet dependencies ----
  # PRESENT is not the same as RIGHT. A stray wrong-version Markdig once sat
  # where the pinned 0.42 belonged; the old exists-check trusted it, and every
  # EdSharp built that day died at startup wanting System.Memory -- silently,
  # because even the error dialog needed the broken library. So each pinned
  # dll is now verified by the version stamped inside it: the wrong version
  # is deleted and the pinned one fetched.
  foreach ($lPin in @(
      @("ReverseMarkdown", $c_sReverseMarkdownVersion),
      @("HtmlAgilityPack", $c_sHtmlAgilityPackVersion),
      @("Markdig", $c_sMarkdigVersion))) {
    $sPackageId = $lPin[0]
    $sPinVersion = $lPin[1]
    $sDllFile = Join-Path $sScriptDir "$sPackageId.dll"
    $bFetch = $true
    if (Test-Path -LiteralPath $sDllFile) {
      try {
        $sFound = [System.Reflection.AssemblyName]::GetAssemblyName($sDllFile).Version.ToString()
      } catch {
        $sFound = "unreadable ($($_.Exception.Message))"
      }
      # The pin "0.42.0" matches assembly version "0.42.0.0": compare on the
      # pin's own dotted parts.
      $lPinParts = $sPinVersion.Split(".")
      $lFoundParts = $sFound.Split(".")
      $bMatch = ($lFoundParts.Count -ge $lPinParts.Count)
      if ($bMatch) {
        for ($iPart = 0; $iPart -lt $lPinParts.Count; $iPart++) {
          if ($lFoundParts[$iPart] -ne $lPinParts[$iPart]) { $bMatch = $false }
        }
      }
      if ($bMatch) {
        writeLog "$sPackageId.dll present with the pinned version $sFound; fetch skipped."
        $bFetch = $false
      } else {
        writeLog "$sPackageId.dll present but WRONG VERSION: found $sFound, pinned $sPinVersion. Deleting and refetching."
        Remove-Item -LiteralPath $sDllFile -Force
      }
    }
    if ($bFetch) {
      fetchNugetPackage $sPackageId $sPinVersion
      $sAfter = [System.Reflection.AssemblyName]::GetAssemblyName($sDllFile).Version.ToString()
      writeLog "$sPackageId.dll fetched; version on disk is now $sAfter."
    }
  }

  # ---- 4. NVDA controller client (optional) ----
  $sNvdaDllFile = Join-Path $sScriptDir "nvdaControllerClient.dll"
  if (Test-Path -LiteralPath $sNvdaDllFile) { writeLog "nvdaControllerClient.dll already present; fetch skipped." }
  else {
    try {
      writeLog "Fetching NVDA controller client from $c_sNvdaClientUrl"
      $sNvdaZipFile = Join-Path $env:TEMP "nvdaControllerClient.zip"
      $sNvdaUnpackDir = Join-Path $env:TEMP "nvdaControllerClient.unpack"
      Invoke-WebRequest -Uri $c_sNvdaClientUrl -OutFile $sNvdaZipFile -UseBasicParsing -ErrorAction Stop
      if (Test-Path -LiteralPath $sNvdaUnpackDir) { Remove-Item -LiteralPath $sNvdaUnpackDir -Recurse -Force }
      Expand-Archive -LiteralPath $sNvdaZipFile -DestinationPath $sNvdaUnpackDir -Force -ErrorAction Stop
      $fNvdaDll = Get-ChildItem -LiteralPath $sNvdaUnpackDir -Recurse -Filter "nvdaControllerClient64.dll" | Select-Object -First 1
      if ($null -eq $fNvdaDll) { $fNvdaDll = Get-ChildItem -LiteralPath $sNvdaUnpackDir -Recurse -Filter "nvdaControllerClient.dll" | Select-Object -First 1 }
      if ($null -eq $fNvdaDll) { writeLog "WARNING: no controller client dll found inside the NVDA archive; continuing without it." }
      else { Copy-Item -LiteralPath $fNvdaDll.FullName -Destination $sNvdaDllFile -Force; writeLog "Copied $($fNvdaDll.Name) as nvdaControllerClient.dll." }
      Remove-Item -LiteralPath $sNvdaZipFile -Force -ErrorAction SilentlyContinue
      Remove-Item -LiteralPath $sNvdaUnpackDir -Recurse -Force -ErrorAction SilentlyContinue
    } catch {
      writeLog "WARNING: NVDA controller client fetch failed ($($_.Exception.Message)); continuing, since the installer treats it as optional."
    }
  }

  # ---- 5. compiler ----
  $sCscFile = findCsc
  # Reference search path for assemblies that are not beside the compiler:
  # the Framework runtime directory and its WPF subfolder, which is where the
  # UI Automation and WindowsBase assemblies live when no Developer Pack is
  # installed.
  $sRuntimeDir = [System.Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory().TrimEnd("\")
  $sWpfDir = Join-Path $sRuntimeDir "WPF"
  $sLibSearch = "/lib:$sRuntimeDir,$sWpfDir"
  writeLog "Reference search path: $sRuntimeDir and $sWpfDir"
  $sFacadeFile = findNetstandardFacade
  if ($sFacadeFile -ne "") { writeLog "netstandard facade (required for Markdig): $sFacadeFile" }
  else { writeLog "WARNING: netstandard.dll facade not found anywhere; compiling against Markdig (netstandard2.0) will fail. Installing the .NET Framework 4.8 Developer Pack fixes this." }

  # ---- 6a. EdSharp.dll (assumption A1) ----
  $lDllExisting = @()
  # WHAT EdSharp.dll REALLY IS -- two lessons in one place.
  #
  # The name-collision lesson: the C# support sources must NOT be compiled
  # into a library called EdSharp.dll. Both assemblies would carry the simple
  # name "EdSharp", .NET binds weak-named assemblies by simple name, and
  # every runtime request for the library's types would be answered with the
  # exe itself -- a TypeLoadException the console build finally printed after
  # a day of silent deaths. The C# is therefore ONE assembly: every .cs file
  # compiles into EdSharp.exe.
  #
  # The evaluator lesson: EdSharp.dll nevertheless EXISTS, and always did --
  # it is the JScript .NET assembly compiled from EdSharp.js by jsc.exe,
  # giving EdSharp its runtime evaluation of expressions (the FileDir / DbDo
  # model). The C# never references it at compile time; it is loaded by
  # reflection from its path at run time, and in that LoadFrom context the
  # shared simple name is tolerated, as years of working EdSharp proved.
  # An earlier fix deleted this file as "stale" -- that would have cost the
  # evaluator; it is REBUILT here instead.
  # A running EdSharp holds EdSharp.exe and the loaded evaluator EdSharp.dll,
  # and both compilers then fail with a sharing violation (jsc's JS2008 was
  # the first sighting -- itself proof that the fixed EdSharp launched and
  # lived). The graceful window-close is used, never a kill: an unsaved-
  # changes prompt inside EdSharp still protects the work, and if it stays
  # open the build stops with a plain sentence instead of a compiler error.
  $lRunning = @(Get-Process -Name "EdSharp" -ErrorAction SilentlyContinue)
  if ($lRunning.Count -gt 0) {
    writeLog "EdSharp is running (process $(@($lRunning | ForEach-Object { $_.Id }) -join ', ')); asking it to close so the build can write its files."
    foreach ($processEdSharp in $lRunning) { [void]$processEdSharp.CloseMainWindow() }
    $iWaited = 0
    while ($iWaited -lt 10 -and @(Get-Process -Name "EdSharp" -ErrorAction SilentlyContinue).Count -gt 0) {
      Start-Sleep -Seconds 1
      $iWaited = $iWaited + 1
    }
    if (@(Get-Process -Name "EdSharp" -ErrorAction SilentlyContinue).Count -gt 0) {
      throw "EdSharp is still running and holds the build outputs. Save your work, close EdSharp, and run buildEdSharp again."
    }
    writeLog "EdSharp closed; the build continues."
  }

  $sJsFile = Join-Path $sScriptDir "EdSharp.js"
  if (-not (Test-Path -LiteralPath $sJsFile)) { throw "EdSharp.js was not found beside the build script; the runtime evaluator cannot be built." }
  $sJscFile = findJsc
  if ($sJscFile -eq "") { throw "jsc.exe was not found in the .NET Framework folders; repair the .NET Framework to build the runtime evaluator." }
  writeLog "JScript compiler: $sJscFile"
  $sEvaluatorFile = Join-Path $sScriptDir "EdSharp.dll"
  if (Test-Path -LiteralPath $sEvaluatorFile) { Remove-Item -LiteralPath $sEvaluatorFile -Force }
  if ((runTool $sJscFile @("/nologo", "/target:library", "/out:EdSharp.dll", "EdSharp.js") "compile EdSharp.dll from EdSharp.js") -ne 0) { throw "The evaluator build failed; the jsc output above names the reason." }

  # ---- 6b. EdSharp.exe (assumptions A2, A3) ----
  if (-not (Test-Path -LiteralPath (Join-Path $sScriptDir "EdSharp.cs"))) { throw "EdSharp.cs was not found beside the build script; nothing to compile." }

  # EdSharp.cs reads BuildVersion.Version, the standard Homer Tools scheme
  # (HomerDescribe, DbDo, bookFido): the build generates Version.cs declaring
  # a static class BuildVersion, gitignored and never committed, so the
  # version literal lives in one source. For EdSharp that source is the
  # AppVersion line of EdSharp_Setup.iss. The class sits in the GLOBAL
  # namespace, which every namespace can see, so this works no matter how
  # EdSharp.cs is organized. The cleanup that removed the fetched dlls
  # evidently removed the generated Version.cs too, which is why the compile
  # suddenly could not find a name that had "always been there".
  $sVersion = "5.0"
  $sIssFile = Join-Path $sScriptDir "EdSharp_Setup.iss"
  if (Test-Path -LiteralPath $sIssFile) {
    foreach ($sLine in Get-Content -LiteralPath $sIssFile) {
      if ($sLine -match "^AppVersion=(.+)$") { $sVersion = $Matches[1].Trim(); break }
    }
    writeLog "Version read from EdSharp_Setup.iss: $sVersion"
  } else { writeLog "EdSharp_Setup.iss not found; BuildVersion falls back to $sVersion." }
  $sStaleFile = Join-Path $sScriptDir "BuildVersion.cs"
  if (Test-Path -LiteralPath $sStaleFile) { Remove-Item -LiteralPath $sStaleFile -Force; writeLog "Removed stale BuildVersion.cs from an earlier build revision." }
  $sBreak = [char]13 + [char]10
  $sVersionSource = "// Version.cs -- generated by BuildEdSharp.ps1 on every run; do not edit or commit." + $sBreak + "public static class BuildVersion" + $sBreak + "{" + $sBreak + "  public const string Version = `"$sVersion`";" + $sBreak + "}" + $sBreak
  Set-Content -LiteralPath (Join-Path $sScriptDir "Version.cs") -Value $sVersionSource -Encoding UTF8 -NoNewline
  writeLog "Generated Version.cs with BuildVersion.Version = $sVersion"
  # Keep the generated file out of the repository, the same way the other
  # Homer Tools do, so a stale committed copy can never rewind the number.
  $sGitignoreFile = Join-Path $sScriptDir ".gitignore"
  if (Test-Path -LiteralPath $sGitignoreFile) {
    $lIgnoreLines = Get-Content -LiteralPath $sGitignoreFile
    if ($lIgnoreLines -notcontains "Version.cs") { Add-Content -LiteralPath $sGitignoreFile -Value "Version.cs"; writeLog "Added Version.cs to .gitignore." }
    else { writeLog ".gitignore already lists Version.cs." }
  } else { writeLog "No .gitignore found; skipped the ignore entry." }
  $sTarget = "/target:winexe"
  if ($bConsole) {
    $sTarget = "/target:exe"
    writeLog "CONSOLE MODE: compiling EdSharp.exe as a console program so startup errors print. NOT for release."
  }
  $lArguments = @("/nologo", $sTarget, "/out:EdSharp.exe", "/platform:anycpu", "/optimize+", $sLibSearch)
  if (Test-Path -LiteralPath (Join-Path $sScriptDir "EdSharp.ico")) { $lArguments += "/win32icon:EdSharp.ico" }
  if (Test-Path -LiteralPath (Join-Path $sScriptDir "EdSharp.manifest")) { $lArguments += "/win32manifest:EdSharp.manifest" }
  foreach ($sReference in $c_lExeReferences) { $lArguments += "/r:$sReference" }
  foreach ($sReference in $c_lUiaReferences) { $lArguments += "/r:$sReference" }
  if ($sFacadeFile -ne "") { $lArguments += "/r:$sFacadeFile" }
  if (Test-Path -LiteralPath (Join-Path $sScriptDir "ReverseMarkdown.dll")) { $lArguments += "/r:ReverseMarkdown.dll" }
  if (Test-Path -LiteralPath (Join-Path $sScriptDir "HtmlAgilityPack.dll")) { $lArguments += "/r:HtmlAgilityPack.dll" }
  if (Test-Path -LiteralPath (Join-Path $sScriptDir "Markdig.dll")) { $lArguments += "/r:Markdig.dll" }
  if (Test-Path -LiteralPath (Join-Path $sScriptDir "Tektosyne.dll")) { $lArguments += "/r:Tektosyne.dll" }
  if (Test-Path -LiteralPath (Join-Path $sScriptDir "Ude.dll")) { $lArguments += "/r:Ude.dll" }
  # One assembly: the main source, the support sources, and the generated
  # version, all inside EdSharp.exe.
  $lArguments += @("EdSharp.cs")
  foreach ($sSourceFile in $c_lDllSources) {
    if (Test-Path -LiteralPath (Join-Path $sScriptDir $sSourceFile)) { $lArguments += $sSourceFile }
    else { writeLog "Support source $sSourceFile not found; skipped." }
  }
  $lArguments += @("Version.cs")
  if ((runTool $sCscFile $lArguments "compile EdSharp.exe") -ne 0) { throw "EdSharp.exe compilation failed; the compiler output above names the lines." }

  # ---- 5b. sqlean.dll: keep the SQLite extension bundle current ----
  # sqlean ships as TWO files in {app} (iss decision of 24 August 2026):
  # sqlean.exe, the sqlean SHELL from the nalgeon/sqlite builds project,
  # which is FROZEN upstream -- there is nothing newer to fetch, so the
  # copy beside this script ships as it is; and sqlean.dll, the
  # single-file extension bundle from nalgeon/sqlean, which is still
  # actively released -- the win-x64 zip asset of each release carries it.
  # This step asks the sqlean releases for the newest tag and refreshes
  # the DLL when it differs from the stamp in sqlean.version. The frozen
  # shell can .load the newer DLL, so refreshing the DLL is what keeps the
  # extension set current. Every failure -- offline, API limit, missing
  # asset -- is logged and leaves the existing copy in place: freshness is
  # wanted, but the build never breaks over it.
  $sSqleanDllFile = Join-Path $sScriptDir "sqlean.dll"
  $sStampFile = Join-Path $sScriptDir "sqlean.version"
  try {
    $oRelease = Invoke-RestMethod -Uri "https://api.github.com/repos/nalgeon/sqlean/releases/latest" -TimeoutSec 30
    $sLatest = "$($oRelease.tag_name)"
    $sHave = if (Test-Path -LiteralPath $sStampFile) { (Get-Content -LiteralPath $sStampFile -Raw).Trim() } else { "" }
    if ($sLatest -and ($sLatest -ne $sHave -or -not (Test-Path -LiteralPath $sSqleanDllFile))) {
      $oAsset = $oRelease.assets | Where-Object { $_.name -match "win.*x64.*\.zip$" } | Select-Object -First 1
      if ($null -eq $oAsset) {
        writeLog "sqlean $sLatest has no Windows x64 zip asset; the current sqlean.dll stays."
      } else {
        $sZipFile = Join-Path $env:TEMP "sqlean_fetch.zip"
        Invoke-WebRequest -Uri $oAsset.browser_download_url -OutFile $sZipFile -TimeoutSec 300
        $sUnpackDir = Join-Path $env:TEMP "sqlean_fetch"
        if (Test-Path -LiteralPath $sUnpackDir) { Remove-Item -LiteralPath $sUnpackDir -Recurse -Force }
        Expand-Archive -LiteralPath $sZipFile -DestinationPath $sUnpackDir -Force
        $oFound = Get-ChildItem -LiteralPath $sUnpackDir -Recurse -Filter "sqlean.dll" | Select-Object -First 1
        if ($null -eq $oFound) {
          writeLog "The sqlean $sLatest archive holds no sqlean.dll; the current copy stays."
        } else {
          Copy-Item -LiteralPath $oFound.FullName -Destination $sSqleanDllFile -Force
          Set-Content -LiteralPath $sStampFile -Value $sLatest
          writeLog "sqlean.dll refreshed to $sLatest. (sqlean.exe, the shell, is frozen upstream and ships as-is; it can .load the newer dll.)"
        }
      }
    } else {
      writeLog "sqlean.dll is current ($sHave)."
    }
  } catch {
    writeLog "sqlean freshness check skipped: $($_.Exception.Message). The current copies stay."
  }

  # ---- 6b2. Convert\inixVert.exe: the Inix table converter ----
  # Compiled from the thin wrapper plus the shared Inix.cs, so the Import
  # and Export tables can move data between .inix, .csv, .tsv, .md, and
  # .xlsx with no Office and no ACE provider (added 24 August 2026).
  if (Test-Path -LiteralPath (Join-Path $sScriptDir "inixVert.cs")) {
    $lArguments = @("/nologo", "/target:exe", "/out:Convert\inixVert.exe", "/platform:anycpu", "/optimize+", $sLibSearch)
    foreach ($sReference in @("System.dll", "System.Core.dll", "System.IO.Compression.dll", "System.Xml.dll")) { $lArguments += "/r:$sReference" }
    $lArguments += @("inixVert.cs", "Inix.cs")
    if ((runTool $sCscFile $lArguments "compile Convert\inixVert.exe") -ne 0) { throw "The inixVert build failed; the compiler output above names the lines." }
  } else {
    writeLog "inixVert.cs is not present, so Convert\inixVert.exe is left as it is."
  }

  # ---- 6c. paired documentation: every tracked .md keeps a fresh .htm ----
  # Judgment call, 23 August 2026 (audit follow-up): the documentation set
  # pairs each Markdown file with an .htm made by pandoc, and a manual
  # conversion is a step someone forgets -- the shipped EdSharp.htm had
  # drifted behind EdSharp.md. So the build regenerates the .htm for every
  # GIT-TRACKED root .md whose .htm is missing or older. Tracked only, so
  # personal notes that live in the folder are never touched; a newly
  # created .md joins the rule at the first build after it is committed.
  $sPandocFile = Join-Path $sScriptDir "Convert\Pandoc\pandoc.exe"
  if (-not (Test-Path -LiteralPath $sPandocFile)) {
    writeLog "pandoc is not at Convert\Pandoc, so the .htm documentation pairs are left as they are (run installPandoc to fetch it)."
  } else {
    $lTrackedMd = @()
    try { $lTrackedMd = @(& git -c core.quotepath=false ls-files "*.md" 2>$null) } catch {}
    $lTrackedMd = @($lTrackedMd | Where-Object { $_ -and ($_ -notmatch "[/\\]") })
    $iFresh = 0
    foreach ($sMdName in $lTrackedMd) {
      $sMdFile = Join-Path $sScriptDir $sMdName.Trim()
      if (-not (Test-Path -LiteralPath $sMdFile)) { continue }
      $sHtmFile = [System.IO.Path]::ChangeExtension($sMdFile, ".htm")
      $bMake = -not (Test-Path -LiteralPath $sHtmFile)
      if (-not $bMake) { $bMake = ((Get-Item -LiteralPath $sMdFile).LastWriteTime -gt (Get-Item -LiteralPath $sHtmFile).LastWriteTime) }
      if (-not $bMake) { continue }
      if ((runTool $sPandocFile @($sMdFile, "-f", "gfm", "-t", "html", "-s", "-o", $sHtmFile) "regenerate $([System.IO.Path]::GetFileName($sHtmFile))") -ne 0) {
        writeLog "WARNING: the .htm for $sMdName was not regenerated; pandoc's output above says why."
      } else { $iFresh++ }
    }
    $sPairs = if ($iFresh -eq 1) { "pair" } else { "pairs" }
    writeLog "Documentation: $iFresh .htm $sPairs regenerated."
  }

  # ---- 7. installer, if Inno Setup is present (buildHomerScribe pattern) ----
  if ($bConsole) {
    writeLog "CONSOLE MODE: the installer step is skipped; run plain buildEdSharp for the release build."
    $sIsccFile = ""
  } else {
  $sIsccFile = findIscc
  if ($sIsccFile -eq "") {
    writeLog "Inno Setup was not found, so no installer was built. To produce EdSharp_Setup.exe, install Inno Setup 6 or open EdSharp_Setup.iss in Inno Setup and choose Compile."
  } else {
    writeLog "Inno Setup: $sIsccFile"
    if (-not (Test-Path -LiteralPath (Join-Path $sScriptDir "EdSharp_Setup.iss"))) { throw "EdSharp_Setup.iss was not found beside the build script, so the installer cannot be compiled." }
    if ((runTool $sIsccFile @("EdSharp_Setup.iss") "compile EdSharp_Setup.exe") -ne 0) { throw "The installer build failed; the Inno Setup output above names the reason." }
    writeLog "Built EdSharp_Setup.exe."
  }

    }

  # ---- 8. report ----
  writeLog "Artifacts:"
  foreach ($sArtifact in @("EdSharp.exe", "EdSharp.dll", "ReverseMarkdown.dll", "HtmlAgilityPack.dll", "Markdig.dll", "nvdaControllerClient.dll", "EdSharp_Setup.exe")) {
    $sArtifactFile = Join-Path $sScriptDir $sArtifact
    if (Test-Path -LiteralPath $sArtifactFile) { writeLog "  $sArtifact  $((Get-Item -LiteralPath $sArtifactFile).Length) bytes" }
    else { writeLog "  $sArtifact  ABSENT" }
  }
  writeLog "Build SUCCESS."
  exit 0
} catch {
  $bFailed = $true
  writeLog "FAILED: $($_.Exception.Message)"
  writeLog "At: $($_.InvocationInfo.PositionMessage)"
  writeLog "Details: $($_ | Out-String)"
  writeLog "Build FAILED. Send this whole file: $sLogFile"
  exit 1
}
