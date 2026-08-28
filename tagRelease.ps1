# tagRelease.ps1
# Tag, push, and publish a GitHub Release for the program in the current
# directory. Generic: works for EdSharp, FileDir, DbDo, or any sibling app
# without editing this script.
#
# Discovery (nothing is hardcoded):
#   * App      -- the name of the current directory. C:\FileDir yields "FileDir".
#   * Setup    -- the matching "<App>_setup.iss" (case-insensitive).
#   * Installer -- "<App>_setup.exe", named from the .iss OutputBaseFilename so
#                  the asset name always matches what the .iss actually produces
#                  (this is the name the F11 Elevate Version command downloads,
#                  and GitHub asset URLs are case-sensitive).
#   * Version  -- read from the .iss. Both Inno styles are supported:
#                  a [Setup] directive   AppVersion=5.0            (EdSharp, FileDir)
#                  a preprocessor define #define AppVersion "1.0.126"  (DbDo)
#   * Owner/repo -- parsed from the git "origin" remote, so the public download
#                  URL is right even if the GitHub repo name differs in case
#                  from the local folder.
#
# VERSION HANDLING (this is what makes F11 / Elevate Version work)
#
#   The F11 Elevate Version command compares the version baked into the running
#   .exe (the "public const string VersionString = ..." line in <App>.cs) with
#   the tag of the latest GitHub release (which this script derives from the
#   .iss AppVersion). Those are two DIFFERENT sources, so they drift:
#   DbDo.cs said 1.0.111 while DbDo_setup.iss said 1.0.126, and EdSharp.cs said
#   5.0.0 while EdSharp_setup.iss said 5.0. When they drift or stay equal, F11
#   cannot see a new build as newer and reports "up to date".
#
#   This script removes that whole class of bug:
#     1. It reads AppVersion from the .iss.
#     2. If that version has ALREADY been released, it bumps the last number by
#        one, so every release gets a genuinely higher version than the one
#        before it -- that is what lets an older install detect this release as
#        newer. If the version has NOT been released yet (because a previous run
#        already bumped it and you have since rebuilt), it is published as-is.
#        This means you never have to remember a flag, and re-running the script
#        never invalidates the installer you just built.
#     3. It writes the bumped version BACK to the .iss (AppVersion, and, when
#        present, AppVerName / VersionInfoVersion) AND into <App>.cs's
#        VersionString constant, so the .exe and the release tag always agree.
#     4. Only then does it tag, release, and upload.
#   Use -Version X.Y.Z to set an explicit version, or -NoBump to release the
#   version already in the .iss unchanged.
#
#   IMPORTANT: because the version is written into <App>.cs, the app must be
#   REBUILT and the installer recompiled after a bump, before it is published.
#   The script compares the installer's timestamp against both version-bearing
#   files and, if the installer is older, stops cleanly (nothing is published),
#   prints the steps, and commits the version files. Just run it again after
#   rebuilding -- no flags:
#
#     .\tagRelease.cmd      bumps if needed, says "rebuild first"
#     Build<App>.cmd        rebuild with the new version
#     (compile the .iss in Inno Setup)
#     .\tagRelease.cmd      publishes that same version (no second bump)
#
# Working tree: uncommitted changes never block a release. The script commits
# the version files it changed (unless -NoCommit) and warns about anything else
# still outstanding, then proceeds.
#
# Run from the repo root:
#   cd C:\FileDir
#   .\tagRelease.cmd                     the normal command; no flags needed
#   .\tagRelease.cmd -Version 5.1        set an explicit version
#   .\tagRelease.cmd -NoBump             never bump, even if already released
#   .\tagRelease.cmd -PrepareOnly        set/sync the version files only, no release
#
# Requirements: git and gh in PATH, gh authenticated (gh auth login),
# PowerShell 5.1+.
#
# This is a maintainer script. Keep tagRelease.ps1, tagRelease.cmd, and
# tagRelease.log in .gitignore so they stay out of the source browser.

# Note: the parameters that receive the .iss text are marked AllowEmptyString /
# AllowEmptyCollection. A PowerShell [Parameter(Mandatory)] on a [string[]] is
# implicitly ValidateNotNullOrEmpty on EVERY element, and an .iss naturally has
# blank lines, so without those attributes binding fails with "Cannot bind
# argument to parameter 'aLines' because it is an empty string."

[CmdletBinding()]
param(
    [string] $Version,
    [switch] $NoBump,
    [switch] $NoCommit,
    [switch] $PrepareOnly,
    [switch] $SkipStaleCheck
)

$ErrorActionPreference = 'Stop'

$sRepoPath = $PWD.Path
$sLogPath  = Join-Path $sRepoPath 'tagRelease.log'

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

function getIssPath {
    # Find "<App>_setup.iss" for the app named by the current directory.
    # Case-insensitive, so C:\EdSharp matches edsharp_setup.iss.
    param(
        [Parameter(Mandatory)] [string] $sPath,
        [Parameter(Mandatory)] [string] $sApp
    )
    $sPattern = "$($sApp)_setup.iss"
    $aFound = @(Get-ChildItem -LiteralPath $sPath -Filter $sPattern -File -ErrorAction SilentlyContinue)
    if ($aFound.Count -eq 1) { return $aFound[0].FullName }
    if ($aFound.Count -gt 1) {
        $sList = ($aFound | ForEach-Object { $_.Name }) -join ', '
        throw "More than one setup script matches $sPattern ($sList)."
    }
    throw "Could not find $sPattern in $sPath. tagRelease assumes the directory name is the app name (e.g. C:\FileDir -> FileDir_setup.iss)."
}

function getIssDirective {
    # Return the value of an Inno [Setup] directive, e.g. OutputBaseFilename.
    # Resolves the {#Token} idiom against a matching #define when present.
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [AllowEmptyString()] [string[]] $aLines,
        [Parameter(Mandatory)] [string]   $sName
    )
    $sValue = ''
    foreach ($sLine in $aLines) {
        if ($sLine -match "^\s*$([regex]::Escape($sName))\s*=\s*(.+?)\s*$") {
            $sValue = $Matches[1]
            break
        }
    }
    if ($sValue -eq '') { return '' }
    # Resolve {#SomeDefine} references against the #define lines.
    while ($sValue -match '\{#(\w+)\}') {
        $sToken = $Matches[1]
        $sResolved = ''
        foreach ($sLine in $aLines) {
            if ($sLine -match "^\s*#define\s+$([regex]::Escape($sToken))\s+`"([^`"]*)`"") {
                $sResolved = $Matches[1]
                break
            }
        }
        if ($sResolved -eq '') { break }
        $sValue = $sValue -replace "\{#$([regex]::Escape($sToken))\}", $sResolved
    }
    return $sValue
}

function getVersionFromIss {
    # Read AppVersion from the .iss, supporting both Inno styles:
    #   #define AppVersion "1.0.126"     (DbDo)
    #   AppVersion=5.0                   (EdSharp, FileDir)
    # The #define is checked first because when both are present the directive
    # is normally just AppVersion={#AppVersion}.
    param([Parameter(Mandatory)] [AllowEmptyCollection()] [AllowEmptyString()] [string[]] $aLines)
    foreach ($sLine in $aLines) {
        if ($sLine -match '^\s*#define\s+AppVersion\s+"([^"]+)"') { return $Matches[1] }
    }
    foreach ($sLine in $aLines) {
        if ($sLine -match '^\s*AppVersion\s*=\s*([^\s{]+)\s*$') { return $Matches[1] }
    }
    throw 'Could not find an AppVersion in the .iss (neither a #define AppVersion line nor an AppVersion= directive).'
}

function bumpVersion {
    # Increment the last dotted-numeric part: 5.0 -> 5.0.1, 1.0.126 -> 1.0.127.
    # A two-part version gains a third part, so an existing 5.0 install sees
    # 5.0.1 as newer (5.0 and 5.0.0 compare EQUAL, which is exactly why simply
    # re-posting a build at the same number never registered as an update).
    param([Parameter(Mandatory)] [string] $sVersion)
    $aParts = $sVersion.Split('.')
    foreach ($sPart in $aParts) {
        if ($sPart -notmatch '^\d+$') { throw "Version '$sVersion' is not dotted-numeric, so it cannot be bumped automatically. Pass -Version X.Y.Z or -NoBump." }
    }
    if ($aParts.Count -lt 3) {
        $aParts = @($aParts) + '1'
    } else {
        $aParts[$aParts.Count - 1] = [string]([int]$aParts[$aParts.Count - 1] + 1)
    }
    return ($aParts -join '.')
}

function setIssVersion {
    # Write the version into the .iss, updating every version-bearing line that
    # is actually present: the #define, the AppVersion directive, and the
    # VersionInfoVersion / AppVerName directives when they carry a literal.
    # Lines that use the {#AppVersion} token need no change. Returns $true if
    # the file was modified. Preserves the file's existing encoding style by
    # writing UTF-8 and CRLF, which is what these .iss files use.
    param(
        [Parameter(Mandatory)] [string]   $sPath,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [AllowEmptyString()] [string[]] $aLines,
        [Parameter(Mandatory)] [string]   $sVersion
    )
    $bChanged = $false
    $aOut = @()
    foreach ($sLine in $aLines) {
        $sNew = $sLine
        if ($sLine -match '^\s*#define\s+AppVersion\s+"[^"]+"') {
            $sNew = $sLine -replace '("(?:[^"]+)")', "`"$sVersion`""
        }
        elseif ($sLine -match '^\s*AppVersion\s*=\s*[^\s{]+\s*$') {
            $sNew = "AppVersion=$sVersion"
        }
        elseif ($sLine -match '^\s*VersionInfoVersion\s*=\s*[\d\.]+\s*$') {
            # VersionInfoVersion must be plain numeric (no "beta" suffixes).
            $sNew = "VersionInfoVersion=$sVersion"
        }
        elseif ($sLine -match '^\s*AppVerName\s*=\s*(.+)$') {
            # Replace only a literal dotted-numeric version inside AppVerName,
            # leaving any surrounding words (e.g. "FileDir 5.0 beta") intact.
            $sRest = $Matches[1]
            if ($sRest -match '\d+(\.\d+)+') {
                $sNew = "AppVerName=" + ($sRest -replace '\d+(\.\d+)+', $sVersion)
            }
        }
        if ($sNew -ne $sLine) { $bChanged = $true }
        $aOut += $sNew
    }
    if ($bChanged) {
        $sText = ($aOut -join "`r`n") + "`r`n"
        [System.IO.File]::WriteAllText($sPath, $sText, (New-Object System.Text.UTF8Encoding($false)))
    }
    return $bChanged
}

function setSourceVersion {
    # Write the version into the app's C# source, so the running .exe reports the
    # same version that this release is tagged with. This is the half that F11
    # actually reads (App.VersionString), and the half that used to drift.
    # Looks for:  public const string VersionString = "X.Y.Z";
    # Returns the path updated, or '' if the app has no such constant.
    param(
        [Parameter(Mandatory)] [string] $sPath,
        [Parameter(Mandatory)] [string] $sApp,
        [Parameter(Mandatory)] [string] $sVersion
    )
    $sCsPath = Join-Path $sPath "$sApp.cs"
    if (-not (Test-Path -LiteralPath $sCsPath -PathType Leaf)) { return '' }
    $sText = [System.IO.File]::ReadAllText($sCsPath)
    $sPattern = '(const\s+string\s+VersionString\s*=\s*")([^"]*)(")'
    if ($sText -notmatch $sPattern) { return '' }
    $sCurrent = ([regex]::Match($sText, $sPattern)).Groups[2].Value
    if ($sCurrent -eq $sVersion) { return $sCsPath }   # already correct
    $sNewText = [regex]::Replace($sText, $sPattern, "`${1}$sVersion`${3}", 1)
    [System.IO.File]::WriteAllText($sCsPath, $sNewText)
    return $sCsPath
}

function getOwnerRepo {
    # Parse "owner/repo" from the origin remote, covering both HTTPS and SSH
    # forms. The remote is authoritative: the GitHub repo may differ in case
    # from the local folder, and asset URLs are case-sensitive.
    $ErrorActionPreference = 'Continue'
    $sUrl = (& git config --get remote.origin.url 2>$null | Out-String).Trim()
    if (-not $sUrl) { throw "No 'origin' remote is configured, so the GitHub owner/repo cannot be determined." }
    $sTrimmed = $sUrl -replace '\.git$', ''
    if ($sTrimmed -match '[:/]([^/:]+)/([^/]+)$') {
        return "$($Matches[1])/$($Matches[2])"
    }
    throw "Could not parse owner/repo from the origin remote: $sUrl"
}

function isReleased {
    # Has this tag already been published?  Asks GitHub first (authoritative,
    # and covers a release made from another machine), then falls back to a
    # local tag.  This is what lets the script decide by itself whether the
    # version in the .iss still needs releasing or has been used already.
    param([Parameter(Mandatory)] [string] $sTag)
    $ErrorActionPreference = 'Continue'
    if (Get-Command gh -ErrorAction SilentlyContinue) {
        & gh release view $sTag 2>$null | Out-Null
        if ($LASTEXITCODE -eq 0) { return $true }
    }
    & git rev-parse --verify --quiet "refs/tags/$sTag" 2>$null | Out-Null
    if ($LASTEXITCODE -eq 0) { return $true }
    return $false
}

function invokeChecked {
    # Run an external program; show output; throw on a non-zero exit.
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
    # Run an external program; return the exit code, never throw. Used to ASK
    # whether something exists, where a non-zero exit is a valid "no".
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
    $sApp = Split-Path -Leaf $sRepoPath

    Write-Host "=== tagRelease.ps1 ==="
    Write-Host "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Host "Log:     $sLogPath"
    Write-Host "Repo:    $sRepoPath"
    Write-Host "App:     $sApp  (from the directory name)"
    Write-Host ""

    # --- Tool checks ---
    Write-Host "--- Tool checks ---"
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) { throw "git is not in PATH." }
    Write-Host "git: $(& git --version | Select-Object -First 1)"

    $iCode = tryInvoke -sExe 'git' -aArgs @('rev-parse', '--is-inside-work-tree')
    if ($iCode -ne 0) { throw "$sRepoPath is not a git working tree." }

    if (-not $PrepareOnly) {
        if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
            throw "gh is not in PATH. Install GitHub CLI from https://cli.github.com/ and run: gh auth login"
        }
        $sGhFirstLine = ((& gh --version 2>&1 | Out-String) -split "`n")[0].Trim()
        Write-Host "gh:  $sGhFirstLine"
        $iAuthCode = tryInvoke -sExe 'gh' -aArgs @('auth', 'status')
        if ($iAuthCode -ne 0) {
            & gh auth status
            throw "gh is not authenticated. Run: gh auth login --web --git-protocol https"
        }
        Write-Host "gh authentication: OK"
    }

    # --- Discovery ---
    Write-Host ""
    Write-Host "--- Discovery ---"
    $sIssPath = getIssPath -sPath $sRepoPath -sApp $sApp
    $sIssName = Split-Path -Leaf $sIssPath
    $aIssLines = [System.IO.File]::ReadAllText($sIssPath) -split "`r?`n"
    Write-Host "Setup script: $sIssName"

    # Asset name comes from OutputBaseFilename so it always matches what Inno
    # actually emits. F11 downloads this exact name and GitHub is case-sensitive.
    $sBase = getIssDirective -aLines $aIssLines -sName 'OutputBaseFilename'
    if (-not $sBase) { $sBase = "$($sApp)_setup" }
    $sSetupExe = "$sBase.exe"
    Write-Host "Installer:    $sSetupExe  (from OutputBaseFilename)"

    $sOwnerRepo = getOwnerRepo
    Write-Host "GitHub repo:  $sOwnerRepo  (from the origin remote)"

    # --- Version ---
    Write-Host ""
    Write-Host "--- Version ---"
    $sOldVersion = getVersionFromIss -aLines $aIssLines
    Write-Host "Current version in $($sIssName): $sOldVersion"

    # Decide the version WITHOUT needing a flag.  The rule: a version is bumped
    # only if it has already been released.  So if the .iss version is not yet
    # published (typically because a previous run, or -PrepareOnly, just bumped it
    # and you have since rebuilt), it is released as-is; re-running the script
    # never invalidates the installer you just built.  If the .iss version IS
    # already published, a new number is needed, so it is bumped.
    if ($Version) {
        $sVersion = $Version.Trim()
        Write-Host "Using the explicit version requested: $sVersion"
    } elseif ($NoBump) {
        $sVersion = $sOldVersion
        Write-Host "-NoBump: releasing the current version unchanged."
    } elseif (isReleased -sTag "v$sOldVersion") {
        $sVersion = bumpVersion -sVersion $sOldVersion
        Write-Host "v$sOldVersion is already released, so a new number is needed."
        Write-Host "Bumped to: $sVersion  (every release gets a higher number, so F11 can detect it)"
    } else {
        $sVersion = $sOldVersion
        Write-Host "v$sOldVersion has not been released yet, so it is published as-is (no bump)."
    }
    $sTag = "v$sVersion"
    Write-Host "Tag:       $sTag"

    # Sync the version into the .iss and the app source, so the .exe's
    # VersionString and the release tag can never disagree.
    $aChangedFiles = @()
    if (setIssVersion -sPath $sIssPath -aLines $aIssLines -sVersion $sVersion) {
        Write-Host "Updated $sIssName to version $sVersion"
        $aChangedFiles += $sIssName
    }
    $sCsUpdated = setSourceVersion -sPath $sRepoPath -sApp $sApp -sVersion $sVersion
    if ($sCsUpdated) {
        $sCsName = Split-Path -Leaf $sCsUpdated
        Write-Host "Synced VersionString in $sCsName to $sVersion"
        if ((& git status --porcelain -- $sCsName 2>$null | Out-String).Trim()) {
            $aChangedFiles += $sCsName
        }
    } else {
        Write-Host "NOTE: no 'const string VersionString' found in $sApp.cs, so the .exe version was not synced." -ForegroundColor Yellow
        Write-Host "      If this app implements Elevate Version (F11), add such a constant so the running" -ForegroundColor Yellow
        Write-Host "      version and the release tag stay in step." -ForegroundColor Yellow
    }

    if ($PrepareOnly) {
        Write-Host ""
        Write-Host "=== -PrepareOnly: version files updated; nothing was tagged or published. ===" -ForegroundColor Green
        Write-Host "Next: rebuild the app, recompile $sIssName in Inno Setup, then run tagRelease again."
        Stop-Transcript | Out-Null
        exit 0
    }

    # --- Asset check ---
    Write-Host ""
    Write-Host "--- Asset check ---"
    if (-not (Test-Path -LiteralPath $sSetupExe -PathType Leaf)) {
        throw "$sSetupExe not found in $sRepoPath. Build the app, compile $sIssName with Inno Setup, then re-run."
    }
    $oExe = Get-Item -LiteralPath $sSetupExe
    Write-Host ("Asset:   {0}" -f $oExe.Name)
    Write-Host ("  size:  {0:N0} bytes" -f $oExe.Length)
    Write-Host ("  mtime: {0}" -f $oExe.LastWriteTime)

    # The installer must be NEWER than both version-bearing files, or it still
    # contains an older version number than the one being tagged -- exactly the
    # drift that breaks F11.  This is checked every run (not only when this run
    # changed the files), so an installer left over from an earlier version is
    # caught too.  When it is stale the script stops cleanly and tells you what to
    # do; it is not an error, just work still to do, so re-running plain
    # tagRelease afterwards finishes the job (the version is not bumped again,
    # because v$sVersion has not been released yet).
    if (-not $SkipStaleCheck) {
        $dtSource = (Get-Item -LiteralPath $sIssPath).LastWriteTime
        $sCsPath = Join-Path $sRepoPath "$sApp.cs"
        if (Test-Path -LiteralPath $sCsPath -PathType Leaf) {
            $dtCs = (Get-Item -LiteralPath $sCsPath).LastWriteTime
            if ($dtCs -gt $dtSource) { $dtSource = $dtCs }
        }
        if ($oExe.LastWriteTime -lt $dtSource) {
            Write-Host ""
            Write-Host "=== REBUILD NEEDED -- nothing was published. ===" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "$sSetupExe is older than the version files, so it does not yet contain" -ForegroundColor Yellow
            Write-Host "version $sVersion.  Publishing it would tag a release whose program reports the" -ForegroundColor Yellow
            Write-Host "wrong version, and F11 would misbehave for your users." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "The version files are now set to $sVersion.  Do this:" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  1. Build$sApp.cmd" -ForegroundColor Yellow
            Write-Host "  2. Compile $sIssName in Inno Setup" -ForegroundColor Yellow
            Write-Host "  3. .\tagRelease.cmd            (no flags needed)" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "Step 3 will publish $sVersion as-is: it does not bump again, because" -ForegroundColor Yellow
            Write-Host "v$sVersion has not been released yet." -ForegroundColor Yellow
            Write-Host ""
            if ($aChangedFiles.Count -gt 0 -and -not $NoCommit) {
                Write-Host "Committing the version files so they are not lost: $($aChangedFiles -join ', ')"
                invokeChecked -sExe 'git' -aArgs (@('add') + $aChangedFiles)
                invokeChecked -sExe 'git' -aArgs @('commit', '-m', "$sApp $sVersion")
            }
            Stop-Transcript | Out-Null
            exit 0
        }
    }

    # --- Commit the version files ---
    # Uncommitted changes never block the release. The version files that this
    # script edited are committed so the tag points at a commit that actually
    # contains this version; anything else outstanding is reported and left alone.
    Write-Host ""
    Write-Host "--- Working tree ---"
    if ($aChangedFiles.Count -gt 0 -and -not $NoCommit) {
        Write-Host "Committing version files: $($aChangedFiles -join ', ')"
        invokeChecked -sExe 'git' -aArgs (@('add') + $aChangedFiles)
        invokeChecked -sExe 'git' -aArgs @('commit', '-m', "$sApp $sVersion")
        $iCode = tryInvoke -sExe 'git' -aArgs @('push')
        if ($iCode -ne 0) { Write-Host "WARN: could not push the version commit; the tag will still be pushed." -ForegroundColor Yellow }
    }
    $sStatus = (& git status --porcelain 2>&1 | Out-String).TrimEnd()
    if ($sStatus) {
        Write-Host "NOTE: other uncommitted changes are present. Releasing anyway:" -ForegroundColor Yellow
        Write-Host $sStatus -ForegroundColor Yellow
    } else {
        Write-Host "Working tree is clean."
    }

    # --- Tag ---
    Write-Host ""
    Write-Host "--- Tag ---"
    $iCode = tryInvoke -sExe 'git' -aArgs @('rev-parse', $sTag)
    if ($iCode -ne 0) {
        Write-Host "Creating tag $sTag ..."
        invokeChecked -sExe 'git' -aArgs @('tag', '-a', $sTag, '-m', "$sApp $sVersion")
        Write-Host "Pushing tag $sTag to origin ..."
        invokeChecked -sExe 'git' -aArgs @('push', 'origin', $sTag)
    } else {
        Write-Host "Tag $sTag already exists locally. Ensuring it is pushed ..."
        $ErrorActionPreference = 'Continue'
        & git push origin $sTag 2>$null | Out-Null
    }

    # --- Release ---
    Write-Host ""
    Write-Host "--- Release ---"
    $iCode = tryInvoke -sExe 'gh' -aArgs @('release', 'view', $sTag)
    if ($iCode -ne 0) {
        Write-Host "Creating release $sTag with asset $sSetupExe ..."
        invokeChecked -sExe 'gh' -aArgs @(
            'release', 'create', $sTag, $sSetupExe,
            '--title', "$sApp $sVersion",
            '--generate-notes',
            '--latest'
        )
    } else {
        Write-Host "Release $sTag already exists. Replacing asset and marking it latest ..."
        invokeChecked -sExe 'gh' -aArgs @('release', 'upload', $sTag, $sSetupExe, '--clobber')
        $ErrorActionPreference = 'Continue'
        & gh release edit $sTag --latest 2>$null | Out-Null
    }

    # --- Verify the public URL ---
    Write-Host ""
    Write-Host "--- URL verification ---"
    $sUrl = "https://github.com/$sOwnerRepo/releases/latest/download/$sSetupExe"
    Write-Host "Public URL: $sUrl"
    try {
        $oResponse = Invoke-WebRequest -Uri $sUrl -Method Head -MaximumRedirection 5 -UseBasicParsing -ErrorAction Stop
        Write-Host ("URL check: HTTP {0}" -f $oResponse.StatusCode) -ForegroundColor Green
    } catch {
        try {
            $oResponse = Invoke-WebRequest -Uri $sUrl -Method Get -MaximumRedirection 5 -UseBasicParsing -ErrorAction Stop
            Write-Host ("URL check: HTTP {0} (via GET; HEAD was rejected)" -f $oResponse.StatusCode) -ForegroundColor Green
        } catch {
            Write-Host "URL check failed: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "(The release may still be valid; GitHub's CDN can take a few seconds to propagate.)" -ForegroundColor Yellow
        }
    }

    Write-Host ""
    Write-Host "=== $sApp $sVersion published. ===" -ForegroundColor Green
    Write-Host ""
    Write-Host "Stable public download URL (does not change between versions):"
    Write-Host ""
    Write-Host "  $sUrl"
    Write-Host ""
    Write-Host "Users on an older version will now see this release from Elevate Version (F11)."
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
