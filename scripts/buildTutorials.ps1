# buildTutorials.ps1 -- write the tutorials and speak them, in one command.
#
#   buildTutorials                    every Tutorial*.inix: documents, then audio
#   buildTutorials Tutorial_Tagging   just that one
#   buildTutorials -docs              documents and feed only, no speaking
#   buildTutorials -sapi              use Windows voices; fetch nothing
#   buildTutorials -live              perform it now through the screen reader, write no file
#   buildTutorials -fetch             allowed to download engines and voices: buildHomerDev
#                                     passes this; an app's build does not
#   buildTutorials -build             called from a build script; means nothing else
#
# WHAT IT DOES, IN ORDER
#
#   1. Makes sure there are two voices worth listening to (see below).
#   2. Runs makeTutorials.py: the sections of Tutorials.md and the feed.
#   3. Speaks each script into its own .mp3 in help\tutorials, and writes
#      Tutorials.m3u beside them.
#   4. Runs makeTutorials.py again, so the feed picks up the audio just made.
#
# ONE .mp3 PER TUTORIAL, IN help\tutorials, AND NO Tutorials.mkv (25 Sep 2026).
# A folder of audio files is found by anybody who looks in help, each file is
# recognised as audio by its extension, and a person plays the one they want.
# The single chaptered recording it used to make was one file that most
# players treated as one track.
#
# THE TWO VOICES, AND WHY THESE
#
# A walkthrough has two speakers and they must never be confused, because the
# whole point is knowing which words came from the program. So one voice should
# sound like a person and the other should sound like a screen reader.
#
# The narrator is PIPER, with the en_US-lessac-medium voice: a small neural
# engine from the Rhasspy project, free, MIT-licensed, entirely offline once
# fetched, and the best free English narration available for Windows. NVDA
# add-ons and other accessibility software use the same engine, so it is a
# known quantity rather than a novelty.
#
# The screen reader's stand-in is ESPEAK NG: free, tiny, and unmistakably
# synthetic -- the voice a listener recognises as a machine within two words.
# That is not a shortcoming here; it is the requirement.
#
# Neither is installed unless it is missing, both are fetched from their own
# projects, and -sapi skips all of it and uses the Windows voices instead. The
# log records every URL, every size and every exit code.
#
# THE SETTINGS, AND WHY THEY ARE NEARLY ALL DEFAULTS
#
# Narration is left at the voice's own rate and pitch. A narrator who sounds
# hurried is harder to follow than one who sounds ordinary, and a tutorial is
# listened to once by somebody who does not yet know the words.
#
# The screen reader's stand-in is the one thing set deliberately: FASTER than
# the narrator, because that is what a screen reader sounds like to the people
# this is for, and because the contrast does the work that a label would
# otherwise have to do. eSpeak runs at 260 words a minute here against a
# narrator near 150; with the Windows voices the same gap is rate 6 against
# rate 4.
#
# Volume stays equal. Making the reader quieter would suggest it matters less,
# and in a walkthrough it is the half that carries the answers. Pitch is left
# alone in both: two voices that already differ in engine and speed do not need
# a third difference, and pitch-shifted speech is tiring over three minutes.

$ErrorActionPreference = "Stop"
$sTool = Split-Path -Parent $MyInvocation.MyCommand.Path

# THE SCRIPTS LIVE IN help, THE TOOLING IN scripts.
#
# FileDir keeps both in one folder; the Homer layout does not, and this script
# was copied without noticing. $sHere is therefore the folder holding the
# tutorial scripts -- help -- and the audio, the playlist and Tutorials.mkv are
# written beside them, where a person looking for a tutorial looks. The tools,
# the voices and the log stay in scripts.
$sHere = Join-Path (Split-Path -Parent $sTool) "help"
if (-not (Test-Path -LiteralPath $sHere)) { $sHere = $sTool }
$sAudioDir = Join-Path $sHere "tutorials"
if (-not (Test-Path -LiteralPath $sAudioDir)) { New-Item -ItemType Directory -Path $sAudioDir | Out-Null }
$sLogDir = Join-Path (Split-Path -Parent $sTool) "logs"
if (-not (Test-Path -LiteralPath $sLogDir)) { New-Item -ItemType Directory -Path $sLogDir | Out-Null }
$sLog = Join-Path $sLogDir ((Split-Path -Leaf (Split-Path -Parent $sTool)) + "-tutorials-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".log")
# WHERE THE VOICES LIVE: ONE COPY, IN THE KIT'S exec FOLDER (25 Sep 2026).
#
# The rule for a shared component is one copy that every Homer app finds --
# never a copy inside an app's own tree, fetched again for the next app. For
# Whisper or Pandoc that copy is where their installer puts it; piper and
# sherpa-onnx have no installer, so there is no default location to look in.
# Every Homer app already relies on C:\HomerDev for its shared classes, so
# the voices live there too, in exec: the Homer folder for binaries that are
# not in git, which is what a fetched engine and the model files it needs
# are -- the same shape as exiftool.exe with its runtime folder beside it.
# C:\HomerDev\exec, fetched once, found by every app's build; LocalFiles.txt
# names it as never pushed, and the kit's own build skips it.
#
# The order, first found wins: HOMER_VOICES, for a machine that keeps them
# elsewhere; a Piper or sherpa-onnx folder somebody put under Program Files by
# hand; the kit's exec folder; and, only when no kit can be found, the old
# place beside this script. A script's [global] VoiceFolder still overrides
# all of it for that series.
function findKit() {
  if ($env:HomerDev -and (Test-Path -LiteralPath (Join-Path $env:HomerDev "CSharp\Lbc.cs"))) { return $env:HomerDev }
  if (Test-Path -LiteralPath "C:\HomerDev\CSharp\Lbc.cs") { return "C:\HomerDev" }
  $sUp = Split-Path -Parent $sTool
  if (Test-Path -LiteralPath (Join-Path $sUp "CSharp\Lbc.cs")) { return $sUp }
  return ""
}
$sTools = ""
if ($env:HOMER_VOICES -and (Test-Path -LiteralPath $env:HOMER_VOICES)) { $sTools = $env:HOMER_VOICES }
if ($sTools -eq "") {
  foreach ($sTry in @((Join-Path $env:ProgramFiles "Piper"), (Join-Path $env:ProgramFiles "sherpa-onnx"))) {
    if (Test-Path -LiteralPath $sTry) { $sTools = $sTry; break }
  }
}
if ($sTools -eq "") {
  $sKit = findKit
  if ($sKit -ne "") { $sTools = Join-Path $sKit "exec" }
}
if ($sTools -eq "") { $sTools = Join-Path $sTool "voices" }

trap {
  $sWhere = ""
  try { $sWhere = " at line " + $_.InvocationInfo.ScriptLineNumber } catch { }
  try {
    Add-Content -LiteralPath $sLog -Value ((Get-Date).ToString("yyyy-MM-dd HH:mm:ss") + "  UNEXPECTED: " + $_.Exception.Message + $sWhere)
    Add-Content -LiteralPath $sLog -Value ($_.ScriptStackTrace)
  } catch { }
  Write-Host "Something unexpected stopped the script. The log has it."
  exit 1
}

function note([string] $sText) {
  Add-Content -LiteralPath $sLog -Value ((Get-Date).ToString("yyyy-MM-dd HH:mm:ss") + "  " + $sText)
}

function say([string] $sText) {
  Write-Host $sText
  note $sText
}

if (Test-Path -LiteralPath $sLog) { Remove-Item -LiteralPath $sLog -Force }
note "buildTutorials starting"
note ("script: " + $MyInvocation.MyCommand.Path)
note ("PowerShell: " + $PSVersionTable.PSVersion.ToString())
note ("platform: " + [Environment]::OSVersion.VersionString)
note ("working directory: " + (Get-Location).Path)
note ("command line: " + [Environment]::CommandLine)
note ("voices folder: " + $sTools)

# ---- what was asked for ----

$bDocsOnly = $false
$bSapi = $false
$bLive = $false
# ONLY THE KIT'S BUILD FETCHES (25 Sep 2026). The voices are shared by every
# app and live in the kit's exec folder, so the kit's build is the one thing
# that downloads them: buildHomerDev passes -fetch. An app's build finds them
# there; when they are missing it says to run buildHomerDev, and speaks nothing.
$bFetch = $false
$sOnly = ""
foreach ($sArg in $args) {
  $sTrimmed = ("" + $sArg).Trim()
  if ($sTrimmed.Length -eq 0) { continue }
  if ($sTrimmed -eq "-docs") { $bDocsOnly = $true; continue }
  if ($sTrimmed -eq "-sapi") { $bSapi = $true; continue }
  if ($sTrimmed -eq "-live") { $bLive = $true; continue }
  if ($sTrimmed -eq "-fetch") { $bFetch = $true; continue }
  # -build says "called from a build script" and means nothing else. It exists
  # because cmd's %* is NOT reset by a bare "call": buildHomerScribe was run
  # as "buildHomerScribe nobump", called this tool with no arguments, and
  # "nobump" arrived here as a script name. A build always passes -build, so
  # %* is its own again. Any other dash-argument is noted and ignored.
  if ($sTrimmed -eq "-build") { continue }
  if ($sTrimmed.StartsWith("-")) { Write-Host ("Ignoring an argument this tool does not know: " + $sTrimmed); continue }
  if ($sTrimmed.StartsWith("-")) { note ("ignoring unknown switch " + $sTrimmed); continue }
  $sOnly = [System.IO.Path]::GetFileNameWithoutExtension($sTrimmed)
}
note ("docs only: " + $bDocsOnly + ", Windows voices: " + $bSapi + ", live: " + $bLive + ", fetch: " + $bFetch + ", only: " + $sOnly)

# ---- the scripts to build ----

$lsScripts = @()
if ($sOnly -ne "") {
  $sOne = Join-Path $sHere ($sOnly + ".inix")
  if (-not (Test-Path -LiteralPath $sOne)) {
    $lsHave = @(Get-ChildItem -LiteralPath $sHere -Filter "Tutorial*.inix" -ErrorAction SilentlyContinue | ForEach-Object { $_.BaseName })
    say ($sOnly + ".inix is not here. The scripts in " + $sHere + " are: " + $(if ($lsHave.Count -gt 0) { $lsHave -join ", " } else { "none" }))
    exit 1
  }
  $lsScripts = @($sOne)
}
else {
  $lsScripts = @(Get-ChildItem -LiteralPath $sHere -Filter "Tutorial*.inix" | Sort-Object Name | ForEach-Object { $_.FullName })
}
if ($lsScripts.Count -eq 0) { say "0 tutorial scripts here."; exit 1 }
note ("scripts: " + (($lsScripts | ForEach-Object { [System.IO.Path]::GetFileName($_) }) -join ", "))

# ---- running a voice, without its chatter ending the run ----
#
# THIS IS WHAT KILLED THE FIRST REAL RUN:
#
#   UNEXPECTED: [piper] [info] Loaded voice in 0.29 second(s)
#
# Not an error -- piper says it every time, on standard error, because that is
# where command-line tools put progress. But $ErrorActionPreference is "Stop",
# and PowerShell 5 turns a native program's standard error into a terminating
# error when it is piped. So the narrator announcing that it had loaded
# successfully ended the build.
#
# runVoice takes the whole thing back to what it should be: standard error goes
# to the log as text, the exit code decides whether anything went wrong, and
# nothing a tool says can stop the script.
function runVoice([string] $sExe, [string[]] $lsArgs, [object] $oInput, [string] $sLabel) {
  # NOTHING TO SAY MEANS NOTHING TO RUN.
  #
  # This is what hung the build partway through the eighth tutorial. Steps that
  # only explain something carry an empty Hear line, and piper reads its text
  # from standard input: handed nothing, it waits for input that never comes,
  # with no output and no exit.
  #
  # THE PARAMETER IS [object], NOT [string], AND THAT MATTERS. PowerShell
  # converts $null to "" when a parameter is typed [string], so the callers that
  # pass $null -- python, ffmpeg, everything that reads no standard input --
  # arrived here indistinguishable from an empty line of speech. The guard then
  # skipped them all, and the very next build wrote no documents at all:
  #
  #     makeTutorial skipped: nothing to say
  #     The documents could not be written.
  #
  # With [object] a null stays a null, and only a genuinely empty string is
  # treated as nothing to say.
  $sInput = $null
  if ($oInput -is [string]) { $sInput = [string] $oInput }
  if ($null -ne $sInput -and $sInput.Trim() -eq "") {
    note ("  " + $sLabel + " skipped: nothing to say")
    return 0
  }
  $sOut = ""
  $sPrevious = $ErrorActionPreference
  $ErrorActionPreference = "Continue"
  try {
    if ($null -eq $sInput) { $sOut = (& $sExe @lsArgs 2>&1 | Out-String) }
    else { $sOut = ($sInput | & $sExe @lsArgs 2>&1 | Out-String) }
  }
  catch {
    note (("  " + $sLabel + " threw: " + $_.Exception.Message))
  }
  finally { $ErrorActionPreference = $sPrevious }
  # sherpa-onnx writes its whole configuration, a progress line per chunk and
  # its timings to standard error; PowerShell wraps the first such line as a
  # "NativeCommandError". None of it is an error. Of a Kokoro run only the
  # timing lines and anything that looks wrong are kept in the log.
  # PowerShell wraps the first line a native program writes to standard error
  # in a NativeCommandError record: five lines of "At C:\...", "+ ...",
  # "CategoryInfo" and "FullyQualifiedErrorId" per piece, none of them from
  # the program. piper's "[info]" lines are its progress. A seven-walk run
  # logged 186 KB of this on 25 Sep 2026. Dropped; a real message survives.
  $bKokoroRun = $sLabel.StartsWith("kokoro")
  foreach ($sLine in ($sOut -split "`r?`n")) {
    $sTrim = $sLine.Trim()
    if ($sTrim -eq "") { continue }
    if ($sTrim -match "^(At line:|At [A-Z]:\\|\+ |CategoryInfo|FullyQualifiedErrorId|~+$)") { continue }
    if ($sTrim -match "\[info\]") { continue }
    if ($bKokoroRun -and -not ($sTrim -match "Elapsed seconds|RTF|error|fail|not found|cannot|unable")) { continue }
    note ("  " + $sLabel + " | " + $sTrim)
  }
  return $LASTEXITCODE
}

# ---- step 1 of 4: the documents ----

function runMake() {
  $sPython = ""
  foreach ($sTry in @("python.exe", "py.exe")) {
    $oFound = Get-Command $sTry -ErrorAction SilentlyContinue
    if ($oFound -and $sPython -eq "") { $sPython = $oFound.Source }
  }
  if ($sPython -eq "") { say "Python was not found, so the documents cannot be written."; return $false }
  note ("Python: " + $sPython)
  # Take the code runVoice RETURNS. $LASTEXITCODE is whatever the last native
  # command in this session set, which after a skipped or wrapped call is
  # nothing at all -- the log showed "exit code: " with a blank after it.
  $iExit = runVoice $sPython @((Join-Path $sTool "makeTutorials.py")) $null "makeTutorial"
  note ("makeTutorials.py exit code: " + $iExit)
  return ($iExit -eq 0)
}

say "Writing the tutorials into Tutorials.md ..."
if (-not (runMake)) { say "The documents could not be written. The log has why."; exit 1 }
if ($bDocsOnly) { say "Documents only, as asked. Nothing was spoken."; exit 0 }

# ---- step 2 of 4: the voices ----

$sFfmpeg = Join-Path $sTool "ffmpeg.exe"
if (-not (Test-Path -LiteralPath $sFfmpeg)) {
  $oFound = Get-Command ffmpeg.exe -ErrorAction SilentlyContinue
  if ($oFound) { $sFfmpeg = $oFound.Source }
}
note ("ffmpeg: " + $sFfmpeg)
if (-not $bLive -and -not (Test-Path -LiteralPath $sFfmpeg)) {
  say "ffmpeg was not found, and it is what joins the pieces into one file."
  say "Run installMediaTools.cmd in this folder."
  exit 1
}

function fetchTo([string] $sUrl, [string] $sPath) {
  # One download, with the whole story in the log: where from, where to, and
  # how big it turned out to be.
  note ("fetching " + $sUrl)
  note ("      to " + $sPath)
  try {
    $oOld = $ProgressPreference
    $ProgressPreference = "SilentlyContinue"
    Invoke-WebRequest -Uri $sUrl -OutFile $sPath -UseBasicParsing
    $ProgressPreference = $oOld
    note ("fetched " + (Get-Item -LiteralPath $sPath).Length + " bytes")
    return $true
  }
  catch {
    note ("fetch failed: " + $_.Exception.Message)
    return $false
  }
}

$sPiper = ""
$sPiperVoice = ""
$sEspeak = ""

if (-not $bSapi -and -not $bLive) {
  New-Item -ItemType Directory -Path $sTools -Force | Out-Null

  # PIPER: the narrator. A release zip and one voice, both fetched once.
  $sPiper = Join-Path $sTools "piper\piper.exe"
  if (-not (Test-Path -LiteralPath $sPiper) -and $bFetch) {
    say "Fetching the narrator voice. This happens once."
    $sZip = Join-Path $sTools "piper.zip"
    if (fetchTo "https://github.com/rhasspy/piper/releases/latest/download/piper_windows_amd64.zip" $sZip) {
      try {
        Expand-Archive -LiteralPath $sZip -DestinationPath $sTools -Force
        # THE DOWNLOAD GOES ONCE IT IS UNPACKED. 22 MB of it, sitting beside the
        # 63 MB voice model and the unpacked copy of itself, in a folder nobody
        # thinks to look in. A tool that fetches something should leave behind
        # only the thing it needed.
        try { Remove-Item -LiteralPath $sZip -Force; note ("removed " + $sZip) } catch { }
      }
      catch { note ("could not unpack piper: " + $_.Exception.Message) }
    }
  }
  if (-not (Test-Path -LiteralPath $sPiper)) {
    $oFound = Get-ChildItem -LiteralPath $sTools -Filter "piper.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($oFound) { $sPiper = $oFound.FullName }
  }
  note ("piper: " + $sPiper + ", present: " + (Test-Path -LiteralPath $sPiper))

  # KOKORO, THROUGH SHERPA-ONNX: the better voice, when it can be fetched.
  #
  # Re-investigated 25 September 2026. Kokoro-82M is an open-weight neural
  # voice model under Apache 2.0, trained on public-domain audio, audio under
  # permissive licences, and synthetic audio -- no share-alike clause and no
  # non-commercial clause anywhere in it, so audio made with it can be
  # published under MIT beside the program. It is markedly more natural than
  # piper's medium voices. sherpa-onnx (also Apache 2.0) runs it on Windows
  # as one executable, sherpa-onnx-offline-tts.exe, with the espeak data it
  # needs inside the model bundle: no Python, nothing to install.
  #
  # Both are fetched once into the voices folder: the newest sherpa-onnx
  # Windows x64 build (found through the GitHub API, preferring the static
  # one), and the int8 English Kokoro bundle. If either cannot be fetched,
  # piper does the job as before -- the log says which.
  $sSherpa = ""
  $sKokoroDir = Join-Path $sTools "kokoro-int8-en-v0_19"
  $oFound = Get-ChildItem -LiteralPath $sTools -Filter "sherpa-onnx-offline-tts.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($oFound) { $sSherpa = $oFound.FullName }
  # A WRONG PACKAGE, FETCHED ON 25 SEP 2026 AND CLEARED HERE: the "static-MD-
  # Debug-lib" package is libraries for linking, hundreds of megabytes and no
  # executable. The package with bin\sherpa-onnx-offline-tts.exe is the SHARED
  # one -- win-x64-shared.tar.bz2 -- so that is the only one taken now, and a
  # folder left by the wrong one is removed.
  foreach ($oOld in (Get-ChildItem -LiteralPath $sTools -Directory -Filter "sherpa-onnx-*" -ErrorAction SilentlyContinue)) {
    if ($oOld.Name -match "lib|Debug|static" -and -not (Get-ChildItem -LiteralPath $oOld.FullName -Filter "sherpa-onnx-offline-tts.exe" -Recurse -ErrorAction SilentlyContinue)) {
      try { Remove-Item -LiteralPath $oOld.FullName -Recurse -Force; note ("removed a sherpa-onnx package with no executable: " + $oOld.FullName) } catch { }
    }
  }
  if ($sSherpa -eq "" -and $bFetch) {
    say "Fetching the Kokoro voice engine. This happens once."
    $sUrl = ""
    try {
      $oRel = Invoke-RestMethod -Uri "https://api.github.com/repos/k2-fsa/sherpa-onnx/releases/latest" -Headers @{ "User-Agent" = "HomerDev-buildTutorials" } -UseBasicParsing
      $lsWin = @($oRel.assets | Where-Object { $_.name -match "win-x64" -and $_.name -match "\.tar\.bz2$" -and $_.name -notmatch "cuda|directml|arm|lib|Debug|no-tts" })
      $oPick = $lsWin | Where-Object { $_.name -match "win-x64-shared\.tar\.bz2$" } | Select-Object -First 1
      if (-not $oPick) { $oPick = $lsWin | Where-Object { $_.name -match "shared" } | Select-Object -First 1 }
      if ($oPick) { $sUrl = $oPick.browser_download_url; note ("sherpa-onnx asset: " + $oPick.name) }
      else { note ("no sherpa-onnx package with executables among: " + (($oRel.assets | ForEach-Object { $_.name }) -join ", ")) }
    }
    catch { note ("GitHub API could not be read for sherpa-onnx: " + $_.Exception.Message) }
    if ($sUrl -ne "") {
      $sTar = Join-Path $sTools "sherpa-onnx.tar.bz2"
      if (fetchTo $sUrl $sTar) {
        # Windows carries tar (libarchive) since 2018, and it reads bzip2.
        $iExit = runVoice "tar.exe" @("-xf", $sTar, "-C", $sTools) $null "tar-sherpa"
        note ("tar exit code: " + $iExit)
        try { Remove-Item -LiteralPath $sTar -Force } catch { }
        $oFound = Get-ChildItem -LiteralPath $sTools -Filter "sherpa-onnx-offline-tts.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($oFound) { $sSherpa = $oFound.FullName }
      }
    }
  }
  # The int8 bundle names its model model.int8.onnx; the fp32 one, model.onnx.
  # On 25 Sep 2026 the bundle was fetched and unpacked and then not recognised,
  # because only the second name was looked for.
  function kokoroModel() {
    foreach ($sName in @("model.int8.onnx", "model.onnx")) {
      $sTry = Join-Path $sKokoroDir $sName
      if (Test-Path -LiteralPath $sTry) { return $sTry }
    }
    return ""
  }
  if ($sSherpa -ne "" -and $bFetch -and (kokoroModel) -eq "") {
    say "Fetching the Kokoro voices. This happens once."
    $sTar = Join-Path $sTools "kokoro.tar.bz2"
    if (fetchTo "https://github.com/k2-fsa/sherpa-onnx/releases/download/tts-models/kokoro-int8-en-v0_19.tar.bz2" $sTar) {
      $iExit = runVoice "tar.exe" @("-xf", $sTar, "-C", $sTools) $null "tar-kokoro"
      note ("tar exit code: " + $iExit)
      try { Remove-Item -LiteralPath $sTar -Force } catch { }
    }
  }
  $sKokoroModel = kokoroModel
  $bKokoro = ($sSherpa -ne "") -and ($sKokoroModel -ne "")
  note ("kokoro: " + $bKokoro + ", engine " + $sSherpa + ", model folder " + $sKokoroDir)
  if (-not $bKokoro -and $bFetch) { say "Kokoro is not available, so piper speaks these." }
  # ONE LINE SAYING WHICH VOICES, AND FROM WHERE. On 25 Sep 2026 a build was
  # stopped by hand because its screen said nothing for three minutes and the
  # person took the silence for a second download of the voices.
  if ($bKokoro) { say ("Voices: Kokoro for the narrator, piper for the reader, from " + $sTools + ". Nothing is downloaded.") }
  elseif ($sPiper -ne "" -and $sPiperVoice -ne "") { say ("Voices: piper, from " + $sTools + ". Nothing is downloaded.") }

  # TWO NEURAL VOICES, ONE ENGINE.
  #
  # fetchVoice takes a piper voice name -- speaker and quality -- and brings back
  # the model and its settings file, once.
  function fetchVoice([string] $sSpeaker, [string] $sQuality) {
    $sName = "en_US-" + $sSpeaker + "-" + $sQuality
    # Beside piper.exe, so exec holds one folder per engine.
    $sPiperDir = Join-Path $sTools "piper"
    if (-not (Test-Path -LiteralPath $sPiperDir)) { New-Item -ItemType Directory -Path $sPiperDir -Force | Out-Null }
    $sModel = Join-Path $sPiperDir ($sName + ".onnx")
    $sJson = $sModel + ".json"
    $sBase = "https://huggingface.co/rhasspy/piper-voices/resolve/main/en/en_US/" + $sSpeaker + "/" + $sQuality + "/"
    if (-not (Test-Path -LiteralPath $sModel)) {
      if (-not $bFetch) { return "" }
      if (-not (fetchTo ($sBase + $sName + ".onnx") $sModel)) { return "" }
    }
    if (-not (Test-Path -LiteralPath $sJson)) {
      if (-not (fetchTo ($sBase + $sName + ".onnx.json") $sJson)) { return "" }
    }
    return $sModel
  }

  if (Test-Path -LiteralPath $sPiper) {
    # THE VOICES ARE CHOSEN BY LICENCE FIRST, THEN BY SOUND.
    #
    # The audio these produce is published in DbDo's repository under MIT, so
    # the training data behind each voice matters as much as the voice does.
    # Most of piper's best-known English voices cannot be used that way:
    #
    #   lessac     Blizzard 2013 corpus -- research purposes only, and the
    #              licence explicitly excludes commercial use of anything
    #              derived from it.
    #   ryan       RyanSpeech, CC BY-NC-SA 4.0: non-commercial AND share-alike,
    #              which fights an MIT repository. It is also fine-tuned FROM
    #              lessac, so it carries that restriction too.
    #   hfc_male,  CC BY-NC-SA 4.0.
    #   hfc_female
    #   libritts_r CC BY 4.0 on the data, but fine-tuned from lessac.
    #
    # Bryce Beattie trained a set of voices for exactly this reason -- his
    # stated goal was voices without restrictive licences -- and those are the
    # ones used here. Both are public domain at the source, trained from scratch
    # or from each other, with no restricted ancestor anywhere in the lineage.
    #
    # THE NARRATOR: kristin, a US English female voice trained from scratch on
    # the LJ Speech dataset, which is public domain.
    $sPiperVoice = fetchVoice "kristin" "medium"

    # SR: john, a US English male voice built from LibriVox recordings, public
    # domain, fine-tuned from kristin. Same family, different sex, which is part
    # of why the two sit together well.
    $sReaderVoice = fetchVoice "john" "medium"
    if ($sReaderVoice -eq "") { $sReaderVoice = $sPiperVoice }
  }
  else { $sPiper = ""; $sReaderVoice = "" }

  # ESPEAK NG: the screen reader's stand-in. Already on the machine more often
  # than not, since other accessibility software carries it.
  $oFound = Get-Command espeak-ng.exe -ErrorAction SilentlyContinue
  if ($oFound) { $sEspeak = $oFound.Source }
  if ($sEspeak -eq "") {
    foreach ($sTry in @("$env:ProgramFiles\eSpeak NG\espeak-ng.exe",
                        "${env:ProgramFiles(x86)}\eSpeak NG\espeak-ng.exe")) {
      if ($sEspeak -eq "" -and (Test-Path -LiteralPath $sTry)) { $sEspeak = $sTry }
    }
  }
  if ($sEspeak -eq "" -and (Get-Command winget.exe -ErrorAction SilentlyContinue)) {
    # Already here? Then nothing is fetched. winget knows about packages it
    # installed; an espeak on the PATH or in its usual folder counts too.
    foreach ($sTry in @("$env:ProgramFiles\eSpeak NG\espeak-ng.exe",
                        "${env:ProgramFiles(x86)}\eSpeak NG\espeak-ng.exe")) {
      if ($sEspeak -eq "" -and (Test-Path -LiteralPath $sTry)) { $sEspeak = $sTry }
    }
    $oOnPath = Get-Command espeak-ng.exe -ErrorAction SilentlyContinue
    if ($sEspeak -eq "" -and $oOnPath) { $sEspeak = $oOnPath.Source }
    if ($sEspeak -ne "") { note ("espeak-ng already installed: " + $sEspeak) }
    if ($sEspeak -eq "") {
    say "Fetching the screen reader voice. This happens once, and takes a minute or two."
    # A FETCH THAT CANNOT FINISH MUST NOT BECOME A HANG.
    #
    # eSpeak NG installs from an MSI, and msiexec can sit waiting for something
    # nobody can see -- an elevation prompt on another desktop, a repair, another
    # installer holding the lock. On Jamal's machine it reached "Starting package
    # install..." and stopped there, with no way to tell a slow install from a
    # stuck one.
    #
    # So it runs as a job with a time limit. If the limit passes, the job is
    # stopped, the log says so, and the build carries on with the Windows voices
    # rather than waiting for something that is not coming. --disable-interactivity
    # tells winget not to ask anything it would otherwise ask.
    # WINDOWS MAY BE ASKING FOR PERMISSION IN A WINDOW YOU CANNOT SEE.
    #
    # This is what actually happened: the install reached "Starting package
    # install..." and stopped, and the reason was a User Account Control prompt
    # that had opened behind everything and taken no focus. Nothing was stuck --
    # it was waiting for an answer nobody knew it had asked for.
    #
    # A script that shells out to an installer has to say so. This one says it
    # before the wait begins, and again as soon as the UAC process appears, and
    # names the keystroke that finds the window.
    say "Windows may ask for permission in a window behind this one. If nothing happens, press Alt+Tab and look for User Account Control."
    $iWaitSeconds = 300
    $oJob = Start-Job -ScriptBlock {
      & winget.exe install --id eSpeak-NG.eSpeak-NG --accept-package-agreements `
        --accept-source-agreements --disable-interactivity --silent 2>&1
    }
    $bToldAgain = $false
    $iWaited = 0
    while ($oJob.State -eq "Running" -and $iWaited -lt $iWaitSeconds) {
      Start-Sleep -Seconds 2
      $iWaited = $iWaited + 2
      # consent.exe IS the User Account Control prompt. If it is running, the
      # answer is one Alt+Tab away rather than minutes away.
      if (-not $bToldAgain -and (Get-Process -Name "consent" -ErrorAction SilentlyContinue)) {
        say "Windows is asking for permission now. Press Alt+Tab to reach the User Account Control window and answer Yes."
        note "consent.exe seen: a UAC prompt is open"
        $bToldAgain = $true
      }
    }
    if ($oJob.State -eq "Running") {
      Stop-Job $oJob
      say ("The screen reader voice did not install within " + [int]($iWaitSeconds / 60) + " minutes, so the Windows voices will be used.")
      note ("winget timed out after " + $iWaitSeconds + " seconds; continuing without espeak")
    }
    else {
      Receive-Job $oJob | ForEach-Object { note ("  | " + $_) }
      note ("winget finished")
    }
    Remove-Job $oJob -Force -ErrorAction SilentlyContinue
    foreach ($sTry in @("$env:ProgramFiles\eSpeak NG\espeak-ng.exe",
                        "${env:ProgramFiles(x86)}\eSpeak NG\espeak-ng.exe")) {
      if ($sEspeak -eq "" -and (Test-Path -LiteralPath $sTry)) { $sEspeak = $sTry }
    }
    }
  }
  note ("espeak-ng: " + $sEspeak)
}

$bWindowsVoices = ($sPiper -eq "" -or $sPiperVoice -eq "")
if ($bWindowsVoices -and -not $bLive) {
  if (-not $bSapi) { say "Falling back to the Windows voices; the log says which fetch fell short." }
  note ("using Windows voices: piper=" + $sPiper + " voice=" + $sPiperVoice + " espeak=" + $sEspeak)
}

# Windows voices are needed for the fallback and for the live narration.
Add-Type -AssemblyName System.Speech
$oSpeaker = New-Object System.Speech.Synthesis.SpeechSynthesizer
$lsVoices = @($oSpeaker.GetInstalledVoices() | Where-Object { $_.Enabled } | ForEach-Object { $_.VoiceInfo.Name })
note ("Windows voices: " + ($lsVoices -join ", "))
$sSapiNarrator = if ($lsVoices.Count -gt 0) { $lsVoices[0] } else { "" }
$sSapiReader = ""
foreach ($sWanted in @("eloquence", "eti-eloquence", "ibmtts")) {
  foreach ($sVoice in $lsVoices) { if ($sSapiReader -eq "" -and $sVoice.ToLower().Contains($sWanted)) { $sSapiReader = $sVoice } }
}
if ($sSapiReader -eq "" -and $lsVoices.Count -gt 1) { $sSapiReader = $lsVoices[1] }
if ($sSapiReader -eq "") { $sSapiReader = $sSapiNarrator }
note ("Windows narrator: " + $sSapiNarrator + ", Windows reader: " + $sSapiReader)

$oJaws = $null
if ($bLive) {
  try { $oJaws = New-Object -ComObject FreedomSci.JawsApi; note "JAWS COM server attached" }
  catch {
    note ("JAWS COM server not available: " + $_.Exception.Message)
    say "JAWS is not running, or its COM server is not available."
    exit 1
  }
}

# SPEED. A screen reader user listens faster than a first-time listener. Piper
# takes a length scale, where less is quicker; eSpeak takes words a minute.
# length_scale is duration: less is quicker. The narrator sits just inside
# natural; SR runs noticeably faster, the way a screen reader does for somebody
# who listens all day. The gap is what tells them apart, along with the speaker.
# TWO VOICES, TOLD APART THREE WAYS: who is speaking, how fast, and how flat.
#
# The narrator is female, high quality, and keeps her natural variation -- the
# register is a colleague explaining something she likes, at a clip an
# experienced listener is comfortable with.
#
# SR is male, high quality, and deliberately even: same clarity, none of the
# expression, the way a screen reader sounds when it is doing its job.
#
# Faster than natural, both of them, because an experienced screen reader user
# listens above the rate a narrator would choose and a tutorial that dawdles is
# one nobody finishes.
# THE NARRATOR, TUNED FOR LISTENERS OF EVERY AGE.
#
# Hearing changes with age, and blind listeners are no exception. A beta tester
# found the narrator hard to follow at 0.72. Two things are known to matter:
#
#   Rate. Older listeners' hearing recovers more slowly between sounds, so
#   speech that is fast for a young ear runs sounds together for an older one.
#   0.80 is one step back from 0.72 -- still brisker than natural (1.0), since
#   slow speech tries the patience of the many listeners who hear well.
#
#   Pitch. Age-related hearing loss takes the high frequencies first, and a
#   woman's voice sits about an octave above a man's -- roughly 224 Hz against
#   132. So the narrator is lowered two semitones, to about 200 Hz: easier on an
#   older ear, still clearly a woman's voice, and still well apart from the
#   screen reader's. Keeping the two voices apart in pitch is itself an aid:
#   a difference in pitch is how any listener tells two voices apart.
#
# Both are [global] settings, NarratorScale and NarratorPitch, in semitones;
# 0 leaves the pitch alone.
$dNarratorScale = 0.80
$dNarratorPitch = -2
# READER SLOWER THAN IT WAS (25 Sep 2026). A beta tester asked for "a bit
# slower"; 0.56 was brisk even for a practised listener. 0.64 keeps the reader
# faster than the narrator, which is what tells them apart, and can be set per
# series with ReaderScale in [global].
$dReaderScale = 0.64
$dReaderNoise = 0.333
# KOKORO SPEAKERS, by index in the English v0.19 voice pack: 0 af, 1 af_bella,
# 2 af_nicole, 3 af_sarah, 4 af_sky, 5 am_adam, 6 am_michael, 7 bf_emma,
# 8 bf_isabella, 9 bm_george, 10 bm_lewis. The narrator is af_sarah, a clear
# American woman; the reader am_michael, an even American man -- the same
# pairing as piper's kristin and john, so the two are told apart the same way.
# KokoroNarrator and KokoroReader in [global] change them.
$iKokoroNarrator = 3
$iKokoroReader = 6
$bKokoroForReader = $false
$dReaderGain = 1.0
if ($null -eq $bKokoro) { $bKokoro = $false }
if ($null -eq $sSherpa) { $sSherpa = "" }
if ($null -eq $sKokoroDir) { $sKokoroDir = "" }
if ($null -eq $sKokoroModel) { $sKokoroModel = "" }
# Silence before the first word of a script, and after each passage. The lead-in
# exists because the screen reader is usually still announcing that a program
# opened when the audio starts.
$dLeadIn = 2.0
$dGap = 0.3
$iEspeakRate = 260
$iSapiNarratorRate = 4
$iSapiReaderRate = 6

# ---- reading a SPEAK script ----

function readScript([string] $sPath) {
  $lsSections = New-Object System.Collections.Generic.List[hashtable]
  $dNow = $null
  foreach ($sRaw in (Get-Content -LiteralPath $sPath -Encoding UTF8)) {
    $sLine = $sRaw.Trim()
    if ($sLine.Length -eq 0 -or $sLine.StartsWith(";") -or $sLine.StartsWith("#")) { continue }
    if ($sLine.StartsWith("[") -and $sLine.EndsWith("]")) {
      $dNow = @{ "_name" = $sLine.Substring(1, $sLine.Length - 2).Trim().ToLower() }
      $lsSections.Add($dNow)
      continue
    }
    if ($null -eq $dNow) { continue }
    $iAt = $sLine.IndexOf("=")
    if ($iAt -lt 1) { continue }
    $sField = $sLine.Substring(0, $iAt).Trim()
    $sValue = $sLine.Substring($iAt + 1).Trim()
    if ($sValue.Length -eq 0) { continue }
    if (-not $dNow.ContainsKey($sField)) { $dNow[$sField] = New-Object System.Collections.Generic.List[string] }
    $dNow[$sField].Add($sValue)
  }
  return $lsSections
}

# ---- speaking one tutorial ----

# ---- [global]: the settings a whole demo script runs under ----
#
# A DEMO SCRIPT is an .inix whose sections are speech passages. Its FIRST
# section is [global], and what it holds applies to every passage after it: the
# voices, how fast each speaks, how flat the screen reader is, the silence
# before the first word and between passages, and where the engines live when
# they are somewhere unusual.
#
# Script 00 sets the series defaults. A later script may carry its own [global]
# to override any of them for itself.
#
# Every key is optional, so a demo script with no [global] behaves as before.
function resolveVoice([string] $sName) {
  # "kristin medium", "john medium" or "en_US-john-medium" all resolve.
  $lsParts = $sName -split "[\s\-]+" | Where-Object { $_ -ne "" -and $_ -ne "en" -and $_ -ne "US" -and $_ -ne "en_US" }
  if ($lsParts.Count -ge 2) {
    $sPath = fetchVoice $lsParts[$lsParts.Count - 2] $lsParts[$lsParts.Count - 1]
    if ($sPath -ne "") { return $sPath }
  }
  note ("[global] named a voice that could not be fetched: " + $sName)
  return ""
}

function applyGlobal($dGlobal) {
  if ($null -eq $dGlobal) { return }
  foreach ($sKey in @("NarratorVoice", "ReaderVoice", "NarratorPitch", "NarratorScale", "ReaderScale",
                      "ReaderFlatness", "LeadIn", "Gap", "VoiceFolder", "PiperPath", "FfmpegPath",
                      "KokoroNarrator", "KokoroReader", "Engine", "ReaderGain", "ReaderOnKokoro")) {
    if (-not $dGlobal.ContainsKey($sKey)) { continue }
    $sValue = ([string] $dGlobal[$sKey][0]).Trim()
    if ($sValue -eq "") { continue }
    switch ($sKey) {
      "NarratorVoice"  { $sFound = resolveVoice $sValue; if ($sFound -ne "") { $script:sPiperVoice = $sFound } }
      "ReaderVoice"    { $sFound = resolveVoice $sValue; if ($sFound -ne "") { $script:sReaderVoice = $sFound } }
      "NarratorScale"  { $script:dNarratorScale = [double] $sValue }
      "NarratorPitch"  { $script:dNarratorPitch = [double] $sValue }
      "ReaderScale"    { $script:dReaderScale = [double] $sValue }
      "ReaderFlatness" { $script:dReaderNoise = [double] $sValue }
      "LeadIn"         { $script:dLeadIn = [double] $sValue }
      "Gap"            { $script:dGap = [double] $sValue }
      "VoiceFolder"    { if (Test-Path -LiteralPath $sValue) { $script:sTools = $sValue } }
      "PiperPath"      { if (Test-Path -LiteralPath $sValue) { $script:sPiper = $sValue } }
      "FfmpegPath"     { if (Test-Path -LiteralPath $sValue) { $script:sFfmpeg = $sValue } }
      "KokoroNarrator" { $script:iKokoroNarrator = [int] $sValue }
      "KokoroReader"   { $script:iKokoroReader = [int] $sValue }
      "Engine"         { if ($sValue -eq "piper") { $script:bKokoro = $false } }
      "ReaderGain"     { $script:dReaderGain = [double] $sValue }
      "ReaderOnKokoro" { $script:bKokoroForReader = ($sValue -eq "1" -or $sValue.ToLower() -eq "yes") }
    }
    note ("[global] " + $sKey + " = " + $sValue)
  }
}

function buildOne([string] $sScript) {
  $sStem = [System.IO.Path]::GetFileNameWithoutExtension($sScript)
  $sOut = Join-Path $sAudioDir ($sStem + ".mp3")
  $sWork = Join-Path $env:TEMP ("buildTutorial_" + [Guid]::NewGuid().ToString("N"))
  $script:iPiece = 0
  $script:lsPieces = New-Object System.Collections.Generic.List[string]
  if (-not $bLive) { New-Item -ItemType Directory -Path $sWork -Force | Out-Null }
  note ("building " + $sStem + ", work folder " + $sWork)
  $dtStarted = Get-Date

  function pieceFile() {
    $script:iPiece = $script:iPiece + 1
    return (Join-Path $sWork ("piece_{0:D4}.wav" -f $script:iPiece))
  }

  function kokoroChunks([string] $sText) {
    # LONG UNBROKEN TEXT IS WHAT MAKES KOKORO SLOW. On 25 Sep 2026 a web
    # address spelled out as words -- eighty characters with no full stop --
    # took 71 seconds for 8 seconds of speech, nine times real time, while an
    # ordinary sentence ran at two. The model's cost climbs with the length of
    # what it is handed at once. So the text is cut at sentence ends and, when
    # a sentence is still long, at commas, into pieces of about 120 characters,
    # spoken one after another and joined.
    $lsOut = New-Object System.Collections.Generic.List[string]
    $lsSentences = [regex]::Split($sText, "(?<=[.!?])\s+")
    foreach ($sSentence in $lsSentences) {
      $sSentence = $sSentence.Trim()
      if ($sSentence -eq "") { continue }
      if ($sSentence.Length -le 120) { $lsOut.Add($sSentence); continue }
      $sPending = ""
      foreach ($sPart in [regex]::Split($sSentence, "(?<=,)\s+")) {
        if ($sPending -ne "" -and ($sPending.Length + $sPart.Length) -gt 120) { $lsOut.Add($sPending.Trim()); $sPending = "" }
        $sPending = ($sPending + " " + $sPart).Trim()
      }
      if ($sPending -ne "") { $lsOut.Add($sPending) }
    }
    return $lsOut
  }

  function speakKokoro([string] $sText, [int] $iSpeaker, [double] $dScale, [string] $sFile, [string] $sLabel) {
    # sherpa-onnx writes a 24 kHz wave; every other piece is 22050 Hz mono, and
    # the join expects one format, so the piece is resampled in place. Each
    # chunk is spoken on its own and the chunks are joined into $sFile.
    #
    # THREADS: two. Measured on 25 Sep 2026: two threads gave 1.6 times real
    # time, every core gave 3 -- the model does not parallelise, and the extra
    # threads only fight each other.
    $sInv = [System.Globalization.CultureInfo]::InvariantCulture
    $lsParts = New-Object System.Collections.Generic.List[string]
    $iChunk = 0
    foreach ($sChunk in (kokoroChunks $sText)) {
      $iChunk = $iChunk + 1
      $sRaw = $sFile + ".k" + $iChunk + ".wav"
      runVoice $sSherpa @(("--kokoro-model=" + $sKokoroModel),
                          ("--kokoro-voices=" + (Join-Path $sKokoroDir "voices.bin")),
                          ("--kokoro-tokens=" + (Join-Path $sKokoroDir "tokens.txt")),
                          ("--kokoro-data-dir=" + (Join-Path $sKokoroDir "espeak-ng-data")),
                          ("--kokoro-length-scale=" + [string]::Format($sInv, "{0:0.00}", $dScale)),
                          "--num-threads=2", ("--sid=" + $iSpeaker), ("--output-filename=" + $sRaw), $sChunk) $null $sLabel | Out-Null
      if (Test-Path -LiteralPath $sRaw) { $lsParts.Add($sRaw) }
    }
    if ($lsParts.Count -eq 0) { return }
    if ($lsParts.Count -eq 1) {
      runVoice $sFfmpeg @("-y", "-loglevel", "error", "-i", $lsParts[0], "-ar", "22050", "-ac", "1", $sFile) $null "ffmpeg-resample" | Out-Null
    }
    else {
      $sList = $sFile + ".chunks.txt"
      Set-Content -LiteralPath $sList -Value ($lsParts | ForEach-Object { "file '" + $_.Replace("'", "'\''") + "'" }) -Encoding ASCII
      runVoice $sFfmpeg @("-y", "-loglevel", "error", "-f", "concat", "-safe", "0", "-i", $sList, "-ar", "22050", "-ac", "1", $sFile) $null "ffmpeg-chunks" | Out-Null
      try { Remove-Item -LiteralPath $sList -Force } catch { }
    }
    foreach ($sPart in $lsParts) { try { Remove-Item -LiteralPath $sPart -Force } catch { } }
  }

  function evenOut([string] $sFile, [double] $dGain) {
    # EVERY PIECE AT THE SAME LOUDNESS. Two engines, two speakers, two
    # settings: the narrator came out noticeably louder than the reader. Each
    # piece is brought to the same measured loudness before the join, and the
    # reader's pieces are then scaled by ReaderGain, 1.0 unless a series sets
    # it in [global].
    if (-not (Test-Path -LiteralPath $sFile)) { return }
    $sInv = [System.Globalization.CultureInfo]::InvariantCulture
    $sEven = $sFile + ".even.wav"
    $sFilter = "loudnorm=I=-16:TP=-1.5:LRA=11"
    if ($dGain -ne 1.0) { $sFilter = $sFilter + ",volume=" + [string]::Format($sInv, "{0:0.00}", $dGain) }
    $iEven = runVoice $sFfmpeg @("-y", "-loglevel", "error", "-i", $sFile, "-af", $sFilter, "-ar", "22050", "-ac", "1", $sEven) $null "ffmpeg-loudness"
    if ($iEven -ne 0 -or -not (Test-Path -LiteralPath $sEven)) { note ("  loudness NOT evened for " + [System.IO.Path]::GetFileName($sFile) + ", exit " + $iEven + "; the piece is used as spoken"); return }
    Move-Item -LiteralPath $sEven -Destination $sFile -Force
  }

  function speakNarrator([string] $sText) {
    if ($bLive) {
      $oSpeaker.SelectVoice($sSapiNarrator)
      $oSpeaker.Rate = $iSapiNarratorRate
      $oSpeaker.SetOutputToDefaultAudioDevice()
      $oSpeaker.Speak($sText)
      note ("live narrator: " + $sText)
      return
    }
    $sFile = pieceFile
    if ($bKokoro -and -not $bWindowsVoices) {
      speakKokoro $sText $iKokoroNarrator $dNarratorScale $sFile "kokoro"
    }
    elseif (-not $bWindowsVoices) {
      # Piper reads its text from standard input and writes one wave file.
      # The narrator keeps piper's own variation, which is what makes a voice
      # sound like somebody rather than something, and a half-second between
      # sentences so a point lands before the next one starts.
      runVoice $sPiper @("-m", $sPiperVoice, "--length_scale", "$dNarratorScale",
                         "--sentence_silence", "0.5", "-f", $sFile) $sText "piper" | Out-Null
      # Lower the pitch without changing the length: play it slower by the pitch
      # ratio, then speed the tempo back up by the same ratio. The voice models
      # speak at 22050 Hz.
      if ($dNarratorPitch -ne 0 -and (Test-Path -LiteralPath $sFile)) {
        $dRatio = [Math]::Pow(2.0, $dNarratorPitch / 12.0)
        $sInv = [System.Globalization.CultureInfo]::InvariantCulture
        $sRate = [string]::Format($sInv, "{0:0}", 22050 * $dRatio)
        $sTempo = [string]::Format($sInv, "{0:0.0000}", 1.0 / $dRatio)
        $sLowered = $sFile + ".low.wav"
        runVoice $sFfmpeg @("-y", "-loglevel", "error", "-i", $sFile,
                            "-af", ("asetrate=" + $sRate + ",aresample=22050,atempo=" + $sTempo),
                            $sLowered) $null "ffmpeg-pitch" | Out-Null
        if (Test-Path -LiteralPath $sLowered) { Move-Item -LiteralPath $sLowered -Destination $sFile -Force }
      }
    }
    else {
      $oSpeaker.SelectVoice($sSapiNarrator)
      $oSpeaker.Rate = $iSapiNarratorRate
      $oSpeaker.SetOutputToWaveFile($sFile)
      $oSpeaker.Speak($sText)
      $oSpeaker.SetOutputToNull()
    }
    evenOut $sFile 1.0
    if (Test-Path -LiteralPath $sFile) { $script:lsPieces.Add($sFile) }
    else { note ("no audio made for: " + $sText) }
    note ("narrator: " + $sText)
  }

  function speakReader([string] $sText) {
    if ($bLive) {
      $oJaws.SayString($sText, $true) | Out-Null
      Start-Sleep -Milliseconds ([Math]::Max(400, $sText.Length * 38))
      note ("live JAWS: " + $sText)
      return
    }
    $sFile = pieceFile
    # THE READER SPEAKS THROUGH PIPER EVEN WHEN KOKORO IS HERE (25 Sep 2026).
    # Half the lines in a walk are the reader's, and piper speaks a line in a
    # second where Kokoro takes twenty; piper's john, flattened, is also what a
    # screen reader sounds like. Kokoro's naturalness goes where it is heard,
    # the narrator. KokoroReader=1 in [global] puts the reader on Kokoro too.
    if ($bKokoro -and $bKokoroForReader -and -not $bWindowsVoices) {
      speakKokoro $sText $iKokoroReader $dReaderScale $sFile "kokoro-sr"
    }
    elseif (-not $bWindowsVoices) {
      # SR SPEAKS THROUGH THE SAME NEURAL ENGINE, faster and in another voice.
      #
      # eSpeak was the wrong choice and the research says why: NVDA's default on
      # Windows 10 and 11 is Windows OneCore -- "responsive, natural-sounding" --
      # and eSpeak is the default only on Windows 8.1 and earlier. Meanwhile the
      # Sonata add-on gives NVDA piper voices, so a piper voice IS a contemporary
      # screen reader voice rather than a stand-in for one.
      #
      # It stays distinguishable by speaker and speed, not by sounding broken.
      if ($sReaderVoice -ne "") {
        # NEUTRAL BY SETTING, NOT BY QUALITY.
        #
        # noise_scale is how much the voice varies in tone, noise_w how much it
        # varies in timing. Piper's defaults -- 0.667 and 0.8 -- are what make a
        # narrator sound alive. Lowering both flattens the delivery into the
        # even, unhurried sameness of a screen reader, using the same
        # high-quality model, so SR is clear AND neutral rather than clear OR
        # neutral.
        #
        # sentence_silence is cut to a fifth of a second: a screen reader does
        # not pause to let a sentence land.
        runVoice $sPiper @("-m", $sReaderVoice, "--length_scale", "$dReaderScale",
                           "--noise_scale", "$dReaderNoise", "--noise_w", "$dReaderNoise",
                           "--sentence_silence", "0.2", "-f", $sFile) $sText "piper-sr" | Out-Null
      }
      else {
        runVoice $sEspeak @("-v", "en-us", "-s", "$iEspeakRate", "-w", $sFile, $sText) $null "espeak" | Out-Null
      }
    }
    else {
      $oSpeaker.SelectVoice($sSapiReader)
      $oSpeaker.Rate = $iSapiReaderRate
      $oSpeaker.SetOutputToWaveFile($sFile)
      $oSpeaker.Speak($sText)
      $oSpeaker.SetOutputToNull()
    }
    evenOut $sFile $dReaderGain
    if (Test-Path -LiteralPath $sFile) { $script:lsPieces.Add($sFile) }
    else { note ("no audio made for: " + $sText) }
    note ("reader: " + $sText)
  }

  function gap([double] $dSeconds) {
    if ($bLive) { Start-Sleep -Milliseconds ([int]($dSeconds * 1000)); return }
    $sFile = pieceFile
    runVoice $sFfmpeg @("-y", "-loglevel", "error", "-f", "lavfi", "-i", "anullsrc=r=22050:cl=mono",
                        "-t", "$dSeconds", $sFile) $null "ffmpeg" | Out-Null
    if (Test-Path -LiteralPath $sFile) { $script:lsPieces.Add($sFile) }
  }

  $lsSections = readScript $sScript
  $dAbout = $lsSections | Where-Object { $_["_name"] -eq "about" } | Select-Object -First 1
  # [global] first, so everything spoken below runs under it.
  applyGlobal ($lsSections | Where-Object { $_["_name"] -eq "global" } | Select-Object -First 1)
  $lsSteps = @($lsSections | Where-Object { $_["_name"] -eq "step" })
  if ($lsSteps.Count -eq 0) { say ($sStem + " holds 0 steps."); return $false }
  note ($sStem + ": steps " + $lsSteps.Count)
  say ("Creating " + $sStem + ".mp3, " + $lsSteps.Count + " steps. A few minutes.")
  note ("loudness: every piece to loudnorm I=-16 TP=-1.5 LRA=11; reader gain " + $dReaderGain.ToString([System.Globalization.CultureInfo]::InvariantCulture))

  # A BEAT BEFORE ANYTHING IS SAID.
  #
  # The screen reader is usually still announcing that a program has opened when
  # the audio starts, so the first sentence lands underneath it and is lost.
  # Two seconds of nothing costs nothing and saves the opening line.
  gap $dLeadIn
  # The opening: what this is, in one sentence, with no names and no apology.
  # The voices are named once because they are somebody's work and public
  # domain -- and never again.
  # NO FIXED OPENING. It used to say which voices these are at the top of every
  # walkthrough -- fourteen times the same sentence. The introductions belong to
  # the first script only, where they are written as its first step.
  if ($dAbout -and $dAbout.ContainsKey("Title")) { speakNarrator $dAbout["Title"][0]; gap 0.35 }
  # SETUP IS WRITTEN, NOT SPOKEN. It used to be read aloud after the title, which
  # put three narrator sentences -- title, setup, first step -- before the screen
  # reader said anything. The listener hears the reader within one sentence now,
  # and the starting state is in the transcript for anybody who wants it.

  $iStepAt = 0
  foreach ($dStep in $lsSteps) {
    $iStepAt = $iStepAt + 1
    if ($iStepAt -gt 1 -and (($iStepAt - 1) % 4) -eq 0) { say ("  step " + $iStepAt + " of " + $lsSteps.Count) }
    # Pause= is the one piece of timing a script can set for itself: seconds of
    # silence before the step is spoken. SSML calls this <break time="2s"/>; our
    # key is the same idea with the angle brackets left off. Everything else --
    # which voice, how fast, how flat -- stays in this script, because it is the
    # same for every passage and belongs in one place rather than in nine files.
    if ($dStep.ContainsKey("Pause")) {
      $dWait = 0.0
      if ([double]::TryParse($dStep["Pause"][0], [ref] $dWait)) { gap $dWait }
    }
    if ($dStep.ContainsKey("Say")) { speakNarrator $dStep["Say"][0]; gap 0.3 }
    if ($dStep.ContainsKey("Key")) { speakNarrator ("Press " + $dStep["Key"][0]); gap 0.35 }
    if ($dStep.ContainsKey("Hear")) {
      foreach ($sHeard in $dStep["Hear"]) { speakReader $sHeard }
      gap 0.4
    }
  }
  speakNarrator "End of the walk."
  if ($dAbout -and $dAbout.ContainsKey("Homework")) {
    gap 0.35
    speakNarrator ("Something to try. " + $dAbout["Homework"][0])
  }

  if ($bLive) { say ($sStem + " was spoken live. No file was written."); return $true }

  note ("pieces: " + $script:lsPieces.Count)
  $sList = Join-Path $sWork "pieces.txt"
  $lsLines = $script:lsPieces | ForEach-Object { "file '" + $_.Replace("'", "'\''") + "'" }
  Set-Content -LiteralPath $sList -Value $lsLines -Encoding ASCII
  runVoice $sFfmpeg @("-y", "-loglevel", "error", "-f", "concat", "-safe", "0", "-i", $sList,
                      "-ar", "22050", "-ac", "1", "-codec:a", "libmp3lame", "-q:a", "4", $sOut) $null "ffmpeg" | Out-Null
  $iExit = $LASTEXITCODE
  note ("ffmpeg exit code: " + $iExit)
  try { Remove-Item -LiteralPath $sWork -Recurse -Force } catch { note ("could not clear " + $sWork) }
  if ($iExit -ne 0 -or -not (Test-Path -LiteralPath $sOut)) { say ($sStem + " could not be joined."); return $false }
  $dTook = ((Get-Date) - $dtStarted).TotalSeconds
  $sTook = $(if ($dTook -lt 90) { ([int]$dTook).ToString() + " seconds" } else { ([int][Math]::Round($dTook / 60.0)).ToString() + " minutes" })
  say ("Created " + [System.IO.Path]::GetFileName($sOut) + " in " + $sTook + ".")
  return $true
}

# THE SCRIPTS ARE CHECKED BEFORE ANYTHING IS SPOKEN (25 Sep 2026). checkTutorial
# reads every script against the format and the reader's grammar -- a key with
# no reader line and no named silence, "Alt+T" where the reader says the words,
# a screen reader named, a file name written rather than said. Speaking a
# script with a problem in it wastes the minutes and publishes the mistake, so
# nothing is spoken while it reports one. -docs and -live skip this.
if (-not $bDocsOnly -and -not $bLive) {
  $sCheck = Join-Path $sTool "checkTutorial.py"
  $sPyForCheck = ""
  foreach ($sTry in @("python.exe", "py.exe")) {
    $oFound = Get-Command $sTry -ErrorAction SilentlyContinue
    if ($oFound -and $sPyForCheck -eq "") { $sPyForCheck = $oFound.Source }
  }
  if ((Test-Path -LiteralPath $sCheck) -and $sPyForCheck -ne "") {
    $iCheck = runVoice $sPyForCheck @($sCheck) $null "checkTutorial"
    note ("checkTutorial exit code: " + $iCheck)
    if ($iCheck -ne 0) {
      say "The tutorial scripts have problems; see above and the tutorials-check log. Nothing was spoken."
      exit 1
    }
  }
}

# NO VOICES AND NOT ALLOWED TO FETCH THEM: say where they come from, and stop.
if (-not $bSapi -and -not $bLive -and -not $bDocsOnly) {
  $bHaveVoice = ($bKokoro) -or ($sPiper -ne "" -and $sPiperVoice -ne "")
  if (-not $bHaveVoice) {
    say "No voices in $sTools. Run buildHomerDev: the kit's build fetches them, once, for every app."
    note "stopping: no voices and -fetch not given"
    exit 1
  }
}

# ---- step 3 of 4: speak them ----

# A TUTORIAL WHOSE .mp3 IS ALREADY HERE IS NOT SPOKEN AGAIN. Speaking is the slow
# part, and the only reliable signal that one needs redoing is that its audio is
# gone. Delete a Tutorial_NN_*.mp3 to have it spoken again; -live and a single
# script named on the command line always speak.
$iDone = 0
$iKept = 0
foreach ($sScript in $lsScripts) {
  $sHave = Join-Path $sAudioDir ([System.IO.Path]::GetFileNameWithoutExtension($sScript) + ".mp3")
  if (-not $bLive -and $sOnly -eq "" -and (Test-Path -LiteralPath $sHave)) {
    note ("kept " + $sHave + ", already spoken")
    $iKept = $iKept + 1
    $iDone = $iDone + 1
    continue
  }
  if (buildOne $sScript) { $iDone = $iDone + 1 }
}
if ($iKept -gt 0) { say ("Kept " + $iKept + " tutorial" + $(if ($iKept -eq 1) { "" } else { "s" }) + " already spoken.") }
$oSpeaker.Dispose()

# ---- step 4 of 5: the playlist ----
#
# Tutorials.m3u, beside the .mp3 files in help\tutorials: a plain playlist in
# tutorial order, for any player that reads one -- the Homer Player in FileDir
# opens it as one track per tutorial, named for the tutorial. No Tutorials.mkv
# any more: see the note at the top.

function writePlaylist() {
  $lsM3u = @("#EXTM3U")
  $iListed = 0
  foreach ($sScript in $lsScripts) {
    $sStem = [System.IO.Path]::GetFileNameWithoutExtension($sScript)
    $sMp3 = Join-Path $sAudioDir ($sStem + ".mp3")
    if (-not (Test-Path -LiteralPath $sMp3)) { continue }
    $sTitle = $sStem
    foreach ($sLine in (Get-Content -LiteralPath $sScript)) {
      if ($sLine -match "^\s*Title\s*=\s*(.+?)\s*$") { $sTitle = $matches[1]; break }
    }
    $lsM3u += ("#EXTINF:-1," + $sTitle)
    $lsM3u += ($sStem + ".mp3")
    $iListed = $iListed + 1
  }
  if ($iListed -eq 0) { say "No audio to list."; return $false }
  $sM3u = Join-Path $sAudioDir "Tutorials.m3u"
  [System.IO.File]::WriteAllLines($sM3u, $lsM3u, (New-Object System.Text.UTF8Encoding($false)))
  say ("Wrote Tutorials.m3u naming " + $iListed + " tutorial" + $(if ($iListed -eq 1) { "" } else { "s" }) + ".")
  return $true
}

if (-not $bLive -and $iDone -gt 0) { writePlaylist | Out-Null }

# ---- step 5 of 5: the feed, now that the audio exists ----

if (-not $bLive -and $iDone -gt 0) {
  say "Writing the feed ..."
  runMake | Out-Null
}

say ($iDone.ToString() + " of " + $lsScripts.Count + " tutorials built.")
if ($iDone -lt $lsScripts.Count) { exit 1 }
exit 0
