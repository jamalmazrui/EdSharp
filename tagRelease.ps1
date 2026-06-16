# tagRelease.ps1
# Tag, push, and publish a GitHub Release for the program in the
# current repo, uploading its installer as the release asset.
#
# Generic: this script works in any of your repos without editing.
# Everything that used to be hardcoded is now discovered at runtime:
#
#   * Installer  -- the single "*_setup.exe" file in the repo root.
#                   Exactly one must be present (the script bails if
#                   there are zero or more than one).
#   * Program    -- the installer's root name, e.g. "EdSharp_setup.exe"
#                   yields program "EdSharp".
#   * Setup script -- the matching "<root>.iss", e.g.
#                   "EdSharp_setup.iss".
#   * Version    -- the program version, read from the matching .iss.
#                   Both Inno styles are supported: a [Setup]
#                   "AppVersion=X.Y" directive (EdSharp, FileDir) and a
#                   "#define AppVersion ""X.Y.Z""" preprocessor line
#                   (DbDuo), including the directive-references-define
#                   idiom. This is the single source of truth, so
#                   nothing version-related is hardcoded here.
#   * Owner/repo -- parsed from the "origin" git remote, so the public
#                   download URL is always correct even when the GitHub
#                   repo name differs in case from the local folder
#                   (e.g. local C:\EdSharp, remote .../Edsharp).
#
# Run from the repo root:
#   cd C:\EdSharp
#   .\tagRelease.cmd            (or .ps1) -- publishes; warns on a dirty tree
#   .\tagRelease.cmd -StrictTree           -- bails if the tree is not clean
#
# This script is for the maintainer's release workflow. It is NOT
# distributed to end users; list it (and tagRelease.cmd and
# tagRelease.log) in .gitignore so they stay out of the source browser
# and out of release zips.
#
# Workflow:
#   1. Discover the lone "*_setup.exe" installer in the repo root and
#      derive the program name and the matching ".iss" setup script.
#   2. Read the version from the .iss (a [Setup] AppVersion= directive
#      or a #define AppVersion line) and validate it looks like a
#      dotted numeric version.
#   3. Parse the GitHub owner and repo from the origin remote.
#   4. Confirm the installer exists; report its size and mtime.
#   5. Report the working-tree state. Uncommitted or untracked changes
#      (other than the build artifacts: installer, program .exe, and
#      this log) produce a WARNING and the run continues. Pass
#      -StrictTree to bail on a dirty tree instead.
#   6. Create the tag (or skip if it already exists), then push it.
#   7. Create the GitHub Release with --generate-notes and --latest,
#      attaching the installer; OR if the release already exists,
#      replace the asset with --clobber so re-runs work as updates.
#   8. Verify the public latest-download URL and report the HTTP status.
#
# Logging:
#   All output is captured to .\tagRelease.log via Start-Transcript in
#   the current directory. A FRESH log is written on every run (the old
#   one is overwritten), so the log always reflects the most recent
#   attempt only. Timestamps use the local machine clock. The log is
#   always closed cleanly, even on failure.
#
# Requirements:
#   - git in PATH and authenticated for push.
#   - gh in PATH and authenticated (run: gh auth login).
#   - PowerShell 5.1+ (ships with Windows 10/11).

[CmdletBinding()]
param(
    # By default a messy working tree only produces a warning and the
    # release still publishes. Pass -StrictTree to bail instead (e.g. to
    # guarantee the tag matches exactly what is committed).
    [switch]$StrictTree,
    # Deprecated and accepted only for backward compatibility: proceeding
    # on a dirty tree is now the default, so this switch is a no-op.
    [switch]$AllowDirty
)

# ============================================================
# Setup
# ============================================================

$ErrorActionPreference = 'Stop'

# Constants (named like variables, declared separately, no magic numbers).
$DefaultMaxRedirects  = 5
$DefaultPreviewLines  = 25
$DefaultVersionRegex  = '^\d+(\.\d+){0,3}$'
$DefaultLogName       = 'tagRelease.log'

$sRepoPath = $PWD.Path

# Fresh, fixed-name log in the current directory (overwritten each run).
$sLogPath = Join-Path $sRepoPath $DefaultLogName

# Start transcript logging.
try {
    Start-Transcript -LiteralPath $sLogPath -Force | Out-Null
} catch {
    Write-Host "ERROR: Could not start transcript at $sLogPath" -ForegroundColor Red
    Write-Host $_
    exit 1
}

# ============================================================
# Helpers
# ============================================================

function discoverInstaller {
    # Find the single "*_setup.exe" in the repo root. Returns the
    # FileInfo object. Throws clearly if there are zero or more than
    # one, so the caller never guesses which program to release.
    param([Parameter(Mandatory)] [string] $sPath)
    $aExe = @(Get-ChildItem -LiteralPath $sPath -Filter '*_setup.exe' -File -ErrorAction SilentlyContinue)
    if ($aExe.Count -eq 0) { throw "No *_setup.exe found in $sPath. Build the installer first, then re-run." }
    if ($aExe.Count -gt 1) {
        $sList = ($aExe | ForEach-Object { $_.Name }) -join ', '
        throw "More than one *_setup.exe found in $sPath ($sList). tagRelease expects exactly one installer per repo."
    }
    return $aExe[0]
}

function getVersionFromIss {
    # Read the app version from an Inno Setup .iss, supporting every
    # style used across these repos:
    #   * [Setup] directive:  AppVersion=4.0          (EdSharp, FileDir)
    #   * Preprocessor:       #define AppVersion "1.0.58"  (DbDuo)
    #   * Directive + token:  AppVersion={#AppVersion}  resolved against
    #                         the matching #define.
    # If AppVersion is absent entirely, falls back to a trailing version
    # on AppVerName (e.g. "EdSharp 4.0" -> "4.0"). The resolved value is
    # validated against a dotted numeric pattern so a malformed line can
    # never produce a bogus tag.
    param(
        [Parameter(Mandatory)] [string] $sPath,
        [Parameter(Mandatory)] [string] $sPattern
    )
    if (-not (Test-Path -LiteralPath $sPath -PathType Leaf)) { throw "Setup script not found: $sPath" }
    $aLines = Get-Content -LiteralPath $sPath -Encoding UTF8

    # Pass 1: collect #define symbols (quoted or bare) for token resolution.
    $dDefines = @{}
    foreach ($sLine in $aLines) {
        if ($sLine -match '^\s*#define\s+(\w+)\s+"([^"]*)"') { $dDefines[$Matches[1]] = $Matches[2].Trim(); continue }
        if ($sLine -match '^\s*#define\s+(\w+)\s+([^\s;]+)')  { $dDefines[$Matches[1]] = $Matches[2].Trim() }
    }

    # Pass 2: find AppVersion= INSIDE the [Setup] section, skipping
    # comment lines (';'). Tracking the section avoids a stray match
    # elsewhere (e.g. in [Code]).
    $sRaw = $null
    $bInSetup = $false
    foreach ($sLine in $aLines) {
        $sTrim = $sLine.Trim()
        if ($sTrim -match '^\[(.+)\]\s*$') { $bInSetup = ($Matches[1] -ieq 'Setup'); continue }
        if (-not $bInSetup) { continue }
        if ($sTrim.StartsWith(';')) { continue }
        if ($sTrim -match '^AppVersion\s*=\s*(.+?)\s*$') { $sRaw = $Matches[1].Trim(); break }
    }

    # Choose the source: directive first, then a #define AppVersion,
    # then an intelligent fallback to AppVerName's trailing version.
    $sFound = $null
    if ($sRaw) {
        $sFound = $sRaw
    } elseif ($dDefines.ContainsKey('AppVersion')) {
        $sFound = $dDefines['AppVersion']
    } else {
        foreach ($sLine in $aLines) {
            $sTrim = $sLine.Trim()
            if ($sTrim.StartsWith(';')) { continue }
            if ($sTrim -match '^AppVerName\s*=.*?(\d+(?:\.\d+){0,3})\s*$') { $sFound = $Matches[1]; break }
        }
    }
    if (-not $sFound) { throw "No version in $sPath. Looked for a [Setup] AppVersion= directive, a #define AppVersion, and a trailing version on AppVerName." }

    # Resolve any {#Token} references against the collected #define symbols.
    foreach ($oMatch in [regex]::Matches($sFound, '\{#\s*(\w+)\s*\}')) {
        $sSym = $oMatch.Groups[1].Value
        if ($dDefines.ContainsKey($sSym)) { $sFound = $sFound.Replace($oMatch.Value, $dDefines[$sSym]) }
    }

    # Strip surrounding quotes and whitespace, then validate.
    $sFound = $sFound.Trim().Trim('"').Trim()
    if ($sFound -match '\{#') { throw "AppVersion in $sPath has an unresolved preprocessor token ($sFound). Add a matching #define or hardcode the version." }
    if ($sFound -notmatch $sPattern) { throw "Version ""$sFound"" in $sPath is not a dotted numeric version (e.g. 4.0 or 1.2.3)." }
    return $sFound
}

function parseRemote {
    # Parse the GitHub owner and repo from an origin remote URL.
    # Handles https (.../owner/repo[.git]) and ssh (git@host:owner/repo[.git]).
    # Returns a dictionary with Owner and Repo keys.
    param([Parameter(Mandatory)] [string] $sUrl)
    if ($sUrl -match 'github\.com[/:]([^/]+)/(.+?)(?:\.git)?/?$') {
        return @{ Owner = $Matches[1]; Repo = $Matches[2] }
    }
    throw "Could not parse GitHub owner/repo from origin remote: $sUrl"
}

function invokeChecked {
    # Run an external program; show its output; throw on non-zero exit.
    # Local 'Continue' lets us inspect $LASTEXITCODE before the script
    # level 'Stop' turns native stderr into a terminating exception.
    param(
        [Parameter(Mandatory)] [string]   $sExe,
        [Parameter(Mandatory)] [string[]] $aArgs,
        [string]                          $sLabel = $null
    )
    if (-not $sLabel) { $sLabel = "$sExe $($aArgs -join ' ')" }
    Write-Host "  > $sLabel" -ForegroundColor DarkGray
    $ErrorActionPreference = 'Continue'
    & $sExe @aArgs
    $iCode = $LASTEXITCODE
    if ($iCode -ne 0) { throw "Command failed (exit $iCode): $sLabel" }
}

function tryInvoke {
    # Run an external program; return the exit code, never throw.
    # Stderr is swallowed since callers use this to ASK whether
    # something exists; a non-zero exit is the expected "no" answer.
    param(
        [Parameter(Mandatory)] [string]   $sExe,
        [Parameter(Mandatory)] [string[]] $aArgs
    )
    $ErrorActionPreference = 'Continue'
    & $sExe @aArgs 2>$null | Out-Null
    return $LASTEXITCODE
}

# ============================================================
# Main
# ============================================================

$iExitCode = 0

try {
    Write-Host "=== tagRelease.ps1 (generic) ==="
    Write-Host "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Host "Log:     $sLogPath  (fresh each run)"
    Write-Host "Repo:    $sRepoPath"
    if ($StrictTree) { Write-Host "Mode:    -StrictTree (will bail on uncommitted changes)" }
    Write-Host ""

    # Diagnostics block to make the log useful for debugging.
    Write-Host "--- Environment ---"
    Write-Host "PowerShell: $($PSVersionTable.PSVersion)"
    Write-Host "OS:         $([System.Environment]::OSVersion.VersionString)"
    Write-Host "User:       $([System.Environment]::UserName)"
    Write-Host "Machine:    $([System.Environment]::MachineName)"

    # Confirm we are in a git working tree.
    $iCode = tryInvoke -sExe 'git' -aArgs @('rev-parse', '--is-inside-work-tree')
    if ($iCode -ne 0) { throw "$sRepoPath is not a git working tree. Run this script from the repo root." }

    # Tool checks.
    Write-Host ""
    Write-Host "--- Tool checks ---"
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) { throw "git is not in PATH." }
    Write-Host "git: $(& git --version | Select-Object -First 1)"
    if (-not (Get-Command gh -ErrorAction SilentlyContinue)) { throw "gh is not in PATH. Install GitHub CLI from https://cli.github.com/ and run: gh auth login" }
    $sGhVer = (& gh --version 2>&1 | Out-String).Trim()
    $sGhFirstLine = ($sGhVer -split "`n")[0].Trim()
    Write-Host "gh:  $sGhFirstLine"
    Write-Host ""
    Write-Host "Checking gh authentication..."
    $iAuthCode = tryInvoke -sExe 'gh' -aArgs @('auth', 'status')
    if ($iAuthCode -ne 0) {
        & gh auth status
        throw "gh is not authenticated. Run: gh auth login --web --git-protocol https"
    }
    Write-Host "gh authentication: OK"

    # 1. Discover the installer, program name, and matching .iss file.
    Write-Host ""
    Write-Host "--- Discovery ---"
    $oExe       = discoverInstaller -sPath $sRepoPath
    $sSetupExe  = $oExe.Name
    $sName      = $sSetupExe -replace '(?i)_setup\.exe$', ''
    $sIssFile   = $sSetupExe -replace '(?i)\.exe$', '.iss'
    $sIssPath   = Join-Path $sRepoPath $sIssFile
    Write-Host "Program:    $sName"
    Write-Host "Installer:  $sSetupExe"
    Write-Host "Setup .iss: $sIssFile"

    # 2. Read and validate version.
    $sVersion = getVersionFromIss -sPath $sIssPath -sPattern $DefaultVersionRegex
    $sTag = "v$sVersion"
    Write-Host "Version:    $sVersion  (from $sIssFile)"
    Write-Host "Tag:        $sTag"

    # 3. Parse owner/repo from the origin remote.
    $sRemoteUrl = (& git remote get-url origin 2>$null)
    if (-not $sRemoteUrl) { throw "Could not read the 'origin' remote URL. Is a remote configured?" }
    $dRemote   = parseRemote -sUrl $sRemoteUrl.Trim()
    $sOwner    = $dRemote.Owner
    $sRepoName = $dRemote.Repo
    Write-Host "Remote:     $($sRemoteUrl.Trim())"
    Write-Host "Owner/repo: $sOwner/$sRepoName"

    # 4. Confirm the installer is present (it is, from discovery) and report it.
    Write-Host ""
    Write-Host "--- Asset check ---"
    Write-Host ("Asset:   {0}" -f $oExe.Name)
    Write-Host ("  size:  {0:N0} bytes" -f $oExe.Length)
    Write-Host ("  mtime: {0}" -f $oExe.LastWriteTime)

    # 5. Report working-tree state (non-fatal by default).
    #
    # The release asset is the installer file on disk and the tag simply
    # marks the current commit, so the tree does NOT need to be clean to
    # publish. By default we WARN about uncommitted/untracked changes and
    # proceed; pass -StrictTree to bail instead. Expected build artifacts
    # (this script's log, the installer, and the program's own .exe) are
    # filtered out of the noise entirely.
    Write-Host ""
    Write-Host "--- Working tree check ---"
    $sStatus = & git status --porcelain 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) { throw "git status failed." }
    $sLogPattern = [regex]::Escape($DefaultLogName)
    $sExePattern = [regex]::Escape($sSetupExe)
    $sProgPattern = [regex]::Escape("$sName.exe")
    $aChanges = @()
    foreach ($sLine in ($sStatus -split "`r?`n")) {
        $sTrim = $sLine.Trim()
        if ($sTrim.Length -eq 0) { continue }
        if ($sTrim -match "$sLogPattern\s*$")  { continue }
        if ($sTrim -match "$sExePattern\s*$")  { continue }
        if ($sTrim -match "$sProgPattern\s*$") { continue }
        $aChanges += $sLine
    }
    if ($aChanges.Count -eq 0) {
        Write-Host "Working tree is clean (build artifacts ignored)."
    } else {
        # Porcelain '??' means untracked; anything else is a tracked change.
        $aUntracked = @($aChanges | Where-Object { $_ -match '^\s*\?\?' })
        $aTracked   = @($aChanges | Where-Object { $_ -notmatch '^\s*\?\?' })
        $sSummary = "Working tree has $($aChanges.Count) uncommitted change(s): $($aTracked.Count) tracked, $($aUntracked.Count) untracked."
        if ($StrictTree) {
            Write-Host $sSummary -ForegroundColor Yellow
            foreach ($sLine in $aChanges) { Write-Host "  $sLine" -ForegroundColor Yellow }
            throw "$sSummary -StrictTree was specified, so stopping. Commit the changes or drop -StrictTree to publish anyway."
        }
        Write-Host "WARN: $sSummary" -ForegroundColor Yellow
        Write-Host "WARN: Publishing anyway (warning only; pass -StrictTree to bail)." -ForegroundColor Yellow
        # Print a capped preview so a junk-drawer repo does not flood the log.
        $iPreview = [Math]::Min($aChanges.Count, $DefaultPreviewLines)
        foreach ($sLine in ($aChanges | Select-Object -First $iPreview)) { Write-Host "  $sLine" -ForegroundColor Yellow }
        if ($aChanges.Count -gt $iPreview) {
            Write-Host "  ... and $($aChanges.Count - $iPreview) more (run 'git status' for the full list)." -ForegroundColor Yellow
        }
    }

    # 6. Create or confirm tag, then push.
    Write-Host ""
    Write-Host "--- Tag ---"
    $iCode = tryInvoke -sExe 'git' -aArgs @('rev-parse', $sTag)
    if ($iCode -ne 0) {
        Write-Host "Creating tag $sTag ..."
        invokeChecked -sExe 'git' -aArgs @('tag', '-a', $sTag, '-m', "$sName $sVersion")
        Write-Host "Pushing tag $sTag to origin ..."
        invokeChecked -sExe 'git' -aArgs @('push', 'origin', $sTag)
    } else {
        Write-Host "Tag $sTag already exists locally. Ensuring it is pushed to origin ..."
        try {
            $ErrorActionPreference = 'Continue'
            & git push origin $sTag 2>$null | Out-Null
        } catch {
            # Already-up-to-date is acceptable.
        }
    }

    # 7. Create or update the GitHub Release.
    Write-Host ""
    Write-Host "--- Release ---"
    $iCode = tryInvoke -sExe 'gh' -aArgs @('release', 'view', $sTag)
    if ($iCode -ne 0) {
        Write-Host "Creating release $sTag with asset $sSetupExe ..."
        invokeChecked -sExe 'gh' -aArgs @(
            'release', 'create', $sTag, $sSetupExe,
            '--title', "$sName $sVersion",
            '--generate-notes',
            '--latest'
        )
    } else {
        Write-Host "Release $sTag already exists. Replacing asset $sSetupExe ..."
        invokeChecked -sExe 'gh' -aArgs @(
            'release', 'upload', $sTag, $sSetupExe,
            '--clobber'
        )
    }

    # 8. Verify the public URL.
    Write-Host ""
    Write-Host "--- URL verification ---"
    $sUrl = "https://github.com/$sOwner/$sRepoName/releases/latest/download/$sSetupExe"
    Write-Host "Public URL: $sUrl"
    try {
        $oResponse = Invoke-WebRequest -Uri $sUrl -Method Head -MaximumRedirection $DefaultMaxRedirects -UseBasicParsing -ErrorAction Stop
        Write-Host ("URL check: HTTP {0}" -f $oResponse.StatusCode) -ForegroundColor Green
    } catch {
        # Some GitHub asset endpoints reject HEAD even though GET works.
        try {
            $oResponse = Invoke-WebRequest -Uri $sUrl -Method Get -MaximumRedirection $DefaultMaxRedirects -UseBasicParsing -ErrorAction Stop
            Write-Host ("URL check: HTTP {0} (via GET; HEAD was rejected)" -f $oResponse.StatusCode) -ForegroundColor Green
        } catch {
            Write-Host "URL check failed: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "(The release may still be valid; GitHub's CDN can take a few seconds to propagate.)" -ForegroundColor Yellow
        }
    }

    Write-Host ""
    Write-Host "=== Release published. ===" -ForegroundColor Green
    Write-Host ""
    Write-Host "Stable public download URL (does not change between versions):"
    Write-Host ""
    Write-Host "  $sUrl"
    Write-Host ""

} catch {
    $iExitCode = 1
    Write-Host ""
    Write-Host "=== FAILED ===" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    if ($_.ScriptStackTrace) {
        Write-Host ""
        Write-Host "Stack trace:" -ForegroundColor DarkGray
        Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    }
} finally {
    Write-Host ""
    Write-Host "--- Log saved at: $sLogPath ---"
    if ($iExitCode -ne 0) { Write-Host "--- Exit code: $iExitCode ---" }
    try { Stop-Transcript | Out-Null } catch { }
}

exit $iExitCode
