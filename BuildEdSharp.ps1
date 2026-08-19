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

# ---- constants --------------------------------------------------------------
$c_sHtmlAgilityPackVersion = "1.11.72"
# Markdig is pinned at the last 0.x release. The 1.x line (current as of
# mid-2026) is a major-version bump, and EdSharp.cs was written against the
# 0.x API; raise this only when ready to test. Markdig ships netstandard2.0
# for Framework use, so the netstandard facade reference below is required.
$c_sMarkdigVersion = "0.42.0"
$c_sNvdaClientUrl = "https://download.nvaccess.org/releases/stable/nvda_2026.1_controllerClient.zip"
# ReverseMarkdown is pinned at the last line that supports .NET Framework:
# the 4.x series ships netstandard2.0, which net48 consumes; 5.x and 6.x ship
# only net8.0 and later, which net48 cannot reference at all (proved by the
# 19 Aug 2026 build log). EdSharp.cs only ever compiled against a 4.x dll, so
# this is the API it is written to. Do not raise this to 5.x or 6.x while
# EdSharp targets .NET Framework 4.8.
$c_sReverseMarkdownVersion = "4.7.1"

$c_lDllSources = @("Lbc.cs", "Say.cs", "Inix.cs", "KeyMap.cs", "Web.cs")
$c_lExeReferences = @("System.dll", "System.Core.dll", "System.Data.dll", "System.Drawing.dll", "System.Web.dll", "System.Windows.Forms.dll", "System.Xml.dll", "Microsoft.VisualBasic.dll")
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
  if ($sMode -ieq "nobump") {
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
  if (Test-Path -LiteralPath (Join-Path $sScriptDir "ReverseMarkdown.dll")) { writeLog "ReverseMarkdown.dll already present; fetch skipped." }
  else { fetchNugetPackage "ReverseMarkdown" $c_sReverseMarkdownVersion }

  if (Test-Path -LiteralPath (Join-Path $sScriptDir "HtmlAgilityPack.dll")) { writeLog "HtmlAgilityPack.dll already present; fetch skipped." }
  else { fetchNugetPackage "HtmlAgilityPack" $c_sHtmlAgilityPackVersion }

  if (Test-Path -LiteralPath (Join-Path $sScriptDir "Markdig.dll")) { writeLog "Markdig.dll already present; fetch skipped." }
  else { fetchNugetPackage "Markdig" $c_sMarkdigVersion }

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
  foreach ($sSourceFile in $c_lDllSources) {
    if (Test-Path -LiteralPath (Join-Path $sScriptDir $sSourceFile)) { $lDllExisting += $sSourceFile }
    else { writeLog "Support source $sSourceFile not found; skipped." }
  }
  if ($lDllExisting.Count -eq 0) { writeLog "No support sources found; EdSharp.dll step skipped entirely." }
  else {
    $lArguments = @("/nologo", "/target:library", "/out:EdSharp.dll", "/platform:anycpu", "/optimize+", $sLibSearch)
    foreach ($sReference in $c_lExeReferences) { $lArguments += "/r:$sReference" }
    foreach ($sReference in $c_lUiaReferences) { $lArguments += "/r:$sReference" }
    if ($sFacadeFile -ne "") { $lArguments += "/r:$sFacadeFile" }
    if (Test-Path -LiteralPath (Join-Path $sScriptDir "ReverseMarkdown.dll")) { $lArguments += "/r:ReverseMarkdown.dll" }
    if (Test-Path -LiteralPath (Join-Path $sScriptDir "HtmlAgilityPack.dll")) { $lArguments += "/r:HtmlAgilityPack.dll" }
    if (Test-Path -LiteralPath (Join-Path $sScriptDir "Markdig.dll")) { $lArguments += "/r:Markdig.dll" }
    if (Test-Path -LiteralPath (Join-Path $sScriptDir "Tektosyne.dll")) { $lArguments += "/r:Tektosyne.dll" }
    if (Test-Path -LiteralPath (Join-Path $sScriptDir "Ude.dll")) { $lArguments += "/r:Ude.dll" }
    $lArguments += $lDllExisting
    if ((runTool $sCscFile $lArguments "compile EdSharp.dll") -ne 0) { throw "EdSharp.dll compilation failed; the compiler output above names the lines." }
  }

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
  $lArguments = @("/nologo", "/target:winexe", "/out:EdSharp.exe", "/platform:anycpu", "/optimize+", $sLibSearch)
  if (Test-Path -LiteralPath (Join-Path $sScriptDir "EdSharp.ico")) { $lArguments += "/win32icon:EdSharp.ico" }
  if (Test-Path -LiteralPath (Join-Path $sScriptDir "EdSharp.manifest")) { $lArguments += "/win32manifest:EdSharp.manifest" }
  foreach ($sReference in $c_lExeReferences) { $lArguments += "/r:$sReference" }
  foreach ($sReference in $c_lUiaReferences) { $lArguments += "/r:$sReference" }
  if ($sFacadeFile -ne "") { $lArguments += "/r:$sFacadeFile" }
  if (Test-Path -LiteralPath (Join-Path $sScriptDir "EdSharp.dll")) { $lArguments += "/r:EdSharp.dll" }
  if (Test-Path -LiteralPath (Join-Path $sScriptDir "ReverseMarkdown.dll")) { $lArguments += "/r:ReverseMarkdown.dll" }
  if (Test-Path -LiteralPath (Join-Path $sScriptDir "HtmlAgilityPack.dll")) { $lArguments += "/r:HtmlAgilityPack.dll" }
  if (Test-Path -LiteralPath (Join-Path $sScriptDir "Markdig.dll")) { $lArguments += "/r:Markdig.dll" }
  if (Test-Path -LiteralPath (Join-Path $sScriptDir "Tektosyne.dll")) { $lArguments += "/r:Tektosyne.dll" }
  if (Test-Path -LiteralPath (Join-Path $sScriptDir "Ude.dll")) { $lArguments += "/r:Ude.dll" }
  $lArguments += @("EdSharp.cs", "Version.cs")
  if ((runTool $sCscFile $lArguments "compile EdSharp.exe") -ne 0) { throw "EdSharp.exe compilation failed; the compiler output above names the lines." }

  # ---- 7. installer, if Inno Setup is present (buildHomerScribe pattern) ----
  $sIsccFile = findIscc
  if ($sIsccFile -eq "") {
    writeLog "Inno Setup was not found, so no installer was built. To produce EdSharp_Setup.exe, install Inno Setup 6 or open EdSharp_Setup.iss in Inno Setup and choose Compile."
  } else {
    writeLog "Inno Setup: $sIsccFile"
    if (-not (Test-Path -LiteralPath (Join-Path $sScriptDir "EdSharp_Setup.iss"))) { throw "EdSharp_Setup.iss was not found beside the build script, so the installer cannot be compiled." }
    if ((runTool $sIsccFile @("EdSharp_Setup.iss") "compile EdSharp_Setup.exe") -ne 0) { throw "The installer build failed; the Inno Setup output above names the reason." }
    writeLog "Built EdSharp_Setup.exe."
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
