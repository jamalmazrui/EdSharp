# summarizeSetup.ps1 -- the single Results box, shown after everything.
#
# WHY POWERSHELL RATHER THAN A BATCH FILE
#
# The batch version of this summary died with exit code 255 partway
# through its first run, taking the Results box with it and leaving the
# installation looking as though it had vanished: it runs hidden, so
# there was no window to find. Batch has no way to bound a command that
# hangs, no reliable quoting for the text these lines carry, and a
# missing label ends the whole script without a word. This does the same
# job with none of that: every probe is given a time limit, every line is
# written the moment it is produced, and any failure is reported inside
# the box rather than ending it.
#
# Run by summarizeSetup.cmd, which supplies the PowerShell parameters.

param([switch]$bQuiet)

$sLogDir = Join-Path $env:LOCALAPPDATA "EdSharp\logs"
$sLogFile = Join-Path $sLogDir "EdSharp_setup.log"
$sSummaryFile = Join-Path $sLogDir "EdSharp_setup_summary.txt"
$sResultsFile = Join-Path $sLogDir "EdSharp_setup_results.txt"
$lLines = @()

function say($sText) {
  # Straight to disk as each line is produced: a summary still in memory
  # when something goes wrong is a summary nobody can read.
  $script:lLines += $sText
  try { Add-Content -LiteralPath $sSummaryFile -Value $sText -Encoding UTF8 } catch { }
  try { if ($sText.Trim() -ne "") { Add-Content -LiteralPath $sLogFile -Value "[summary] $sText" -Encoding UTF8 } } catch { }
}

function runBounded($sExe, $lArguments, $iSeconds) {
  $lQuoted = @()
  foreach ($sArgument in $lArguments) {
    if ($sArgument -match "\s") { $lQuoted += ('"' + $sArgument + '"') } else { $lQuoted += $sArgument }
  }
  $lArguments = $lQuoted
  # Run a command, wait no longer than $iSeconds, and return its first
  # line of output. A tool that hangs -- some report versions over the
  # network -- must never hold up the installation.
  $sOutFile = Join-Path $sLogDir ("EdSharp_probe_" + [guid]::NewGuid().ToString("N") + ".tmp")
  try {
    $oProcess = Start-Process -FilePath $sExe -ArgumentList $lArguments -RedirectStandardOutput $sOutFile -RedirectStandardError ($sOutFile + ".err") -WindowStyle Hidden -PassThru
    if (-not $oProcess.WaitForExit($iSeconds * 1000)) {
      try { $oProcess.Kill() } catch { }
      return ""
    }
    $sText = ""
    if (Test-Path -LiteralPath $sOutFile) { $sText = (Get-Content -LiteralPath $sOutFile -ErrorAction SilentlyContinue) -join "`n" }
    if ($sText.Trim() -eq "" -and (Test-Path -LiteralPath ($sOutFile + ".err"))) { $sText = (Get-Content -LiteralPath ($sOutFile + ".err") -ErrorAction SilentlyContinue) -join "`n" }
    foreach ($sLine in $sText -split "`n") { if ($sLine.Trim() -ne "") { return $sLine.Trim() } }
    return ""
  } catch {
    return ""
  } finally {
    foreach ($sLeftover in @($sOutFile, ($sOutFile + ".err"))) {
      try { if (Test-Path -LiteralPath $sLeftover) { Remove-Item -LiteralPath $sLeftover -Force } } catch { }
    }
  }
}

function runBoundedAll($sExe, $lArguments, $iSeconds) {
  $lQuoted = @()
  foreach ($sArgument in $lArguments) {
    if ($sArgument -match "\s") { $lQuoted += ('"' + $sArgument + '"') } else { $lQuoted += $sArgument }
  }
  $lArguments = $lQuoted
  # The same, but returning everything the command printed rather than its
  # first line -- for answers that are lists.
  $sOutFile = Join-Path $sLogDir ("EdSharp_probe_" + [guid]::NewGuid().ToString("N") + ".tmp")
  try {
    $oProcess = Start-Process -FilePath $sExe -ArgumentList $lArguments -RedirectStandardOutput $sOutFile -RedirectStandardError ($sOutFile + ".err") -WindowStyle Hidden -PassThru
    if (-not $oProcess.WaitForExit($iSeconds * 1000)) {
      try { $oProcess.Kill() } catch { }
      return ""
    }
    if (Test-Path -LiteralPath $sOutFile) { return ((Get-Content -LiteralPath $sOutFile -ErrorAction SilentlyContinue) -join "`n") }
    return ""
  } catch {
    return ""
  } finally {
    foreach ($sLeftover in @($sOutFile, ($sOutFile + ".err"))) {
      try { if (Test-Path -LiteralPath $sLeftover) { Remove-Item -LiteralPath $sLeftover -Force } } catch { }
    }
  }
}

function runExit($sExe, $lArguments, $iSeconds) {
  # Quote any argument containing a space. Start-Process joins the list
  # with spaces and quotes nothing, so python -c "import pymupdf4llm"
  # arrived as -c import pymupdf4llm: Python read "import" as the whole
  # program, failed, and a package that was installed was reported
  # missing. This one line was the whole fault.
  $lQuoted = @()
  foreach ($sArgument in $lArguments) {
    if ($sArgument -match "\s") { $lQuoted += ('"' + $sArgument + '"') } else { $lQuoted += $sArgument }
  }
  $lArguments = $lQuoted
  # The exit code of a bounded run: 0 for success, anything else for
  # failure, and -1 when the command could not be run or outstayed its
  # welcome. Output is ignored on purpose -- a library that prints a
  # warning while importing perfectly well must not be called broken.
  try {
    $oProcess = Start-Process -FilePath $sExe -ArgumentList $lArguments -WindowStyle Hidden -PassThru
    if (-not $oProcess.WaitForExit($iSeconds * 1000)) {
      try { $oProcess.Kill() } catch { }
      return -1
    }
    return $oProcess.ExitCode
  } catch {
    return -1
  }
}

function findExe($sName) {
  # The first real program of this name on the path. Paths under
  # WindowsApps are skipped: those are the Microsoft Store stubs, which
  # answer when asked and then advertise the Store instead of running.
  foreach ($oCommand in @(Get-Command $sName -All -ErrorAction SilentlyContinue)) {
    if ($oCommand.Source -and ($oCommand.Source -notmatch "\\WindowsApps\\")) { return $oCommand.Source }
  }
  return ""
}

function allPythons() {
  # Every real Python on this computer, newest first. More than one is
  # common: an installer put 3.14 in C:\Python314 while the Store or an
  # older setup left 3.13 elsewhere, and a package installed into one is
  # invisible to the other.
  $lFound = @()
  foreach ($oCommand in @(Get-Command python -All -ErrorAction SilentlyContinue)) {
    if ($oCommand.Source -and ($oCommand.Source -notmatch "\\WindowsApps\\")) { $lFound += $oCommand.Source }
  }
  foreach ($sRoot in @($env:ProgramFiles, (Join-Path $env:LOCALAPPDATA "Programs\Python"), "C:\")) {
    if (-not $sRoot -or -not (Test-Path -LiteralPath $sRoot)) { continue }
    foreach ($oDir in @(Get-ChildItem -LiteralPath $sRoot -Directory -Filter "Python3*" -ErrorAction SilentlyContinue)) {
      $sTry = Join-Path $oDir.FullName "python.exe"
      if ((Test-Path -LiteralPath $sTry) -and ($lFound -notcontains $sTry)) { $lFound += $sTry }
    }
  }
  return $lFound
}

function recordedPython() {
  # The interpreter installPdfTools.cmd actually used, which it writes
  # down for exactly this reason: a computer can carry several Pythons,
  # and asking a different one whether a package is installed gets an
  # answer that is true and useless.
  $sFile = Join-Path $sLogDir "EdSharp_python.txt"
  if (Test-Path -LiteralPath $sFile) {
    $sPath = (Get-Content -LiteralPath $sFile -TotalCount 1 -ErrorAction SilentlyContinue)
    if ($sPath) {
      $sPath = $sPath.Trim()
      if ($sPath -and (Test-Path -LiteralPath $sPath)) { return $sPath }
    }
  }
  return ""
}

function findPython() {
  $sRecorded = recordedPython
  if ($sRecorded -ne "") { return $sRecorded }
  # The official python.org build, never the Store stub, and found by
  # location as well as by path -- a Python installed minutes ago is not
  # yet on the path this process inherited.
  $sExe = findExe "python"
  if ($sExe -ne "") { return $sExe }
  $lRoots = @($env:ProgramFiles, (Join-Path $env:LOCALAPPDATA "Programs\Python"), "C:\")
  $lFound = @()
  foreach ($sRoot in $lRoots) {
    if (-not $sRoot -or -not (Test-Path -LiteralPath $sRoot)) { continue }
    foreach ($oDir in @(Get-ChildItem -LiteralPath $sRoot -Directory -Filter "Python3*" -ErrorAction SilentlyContinue)) {
      $sTry = Join-Path $oDir.FullName "python.exe"
      if (Test-Path -LiteralPath $sTry) { $lFound += $sTry }
    }
  }
  if ($lFound.Count -gt 0) { return ($lFound | Sort-Object -Descending)[0] }
  return ""
}

function reportTool($sName, $sExe, $sLater) {
  $sPath = findExe $sExe
  if ($sPath -eq "") {
    say "$sName`: not installed. To add it later, $sLater."
    return
  }
  $sVersion = runBounded $sPath @("--version") 20
  if ($sVersion -eq "") { say "$sName`: installed at $sPath" }
  else { say "$sName`: installed, $sVersion" }
}

function reportPython() {
  $sPython = findPython
  if ($sPython -eq "") {
    if ((Get-Command python -ErrorAction SilentlyContinue) -ne $null) {
      say "Python: not installed. Windows has only the Microsoft Store stub, which is not Python; run installPython.cmd in the EdSharp folder for the official python.org build."
    } else {
      say "Python: not installed. To add it later, run installPython.cmd in the EdSharp folder."
    }
    return $null
  }
  $sVersion = runBounded $sPython @("--version") 20
  say "Python: installed, $sVersion at $sPython"
  return $sPython
}

function reportModule($sName, $sPython, $sModule, $sLater) {
  if (-not $sPython) { say "$sName`: not installed, because Python is missing."; return }
  if ((runExit $sPython @("-c", "import $sModule") 60) -eq 0) { say "$sName`: installed." ; return }
  # A machine can carry more than one Python -- this one has 3.13 and 3.14
  # -- and the packages went into whichever the installer script chose. If
  # another Python has them, say so rather than declaring them missing.
  foreach ($sOther in (allPythons)) {
    if ($sOther -eq $sPython) { continue }
    if ((runExit $sOther @("-c", "import $sModule") 60) -eq 0) { say "$sName`: installed for $sOther" ; return }
  }
  say "$sName`: not installed. To add it later, $sLater."
}

function reportWordNet($sPython) {
  if (-not $sPython) { say "Thesaurus database: not installed, because Python is missing."; return }
  foreach ($sTry in (@($sPython) + (allPythons))) {
    if ((runExit $sTry @("-c", "from nltk.corpus import wordnet; wordnet.synsets('test')") 60) -eq 0) {
      say "Thesaurus database: installed. Press Shift+F7 on a word."
      return
    }
  }
  say "Thesaurus database: not installed. To add it later, run installPdfTools.cmd in the EdSharp folder."
}

function reportModel() {
  $sOllama = findExe "ollama"
  if ($sOllama -eq "") { return }
  # One probe, read whole: the listing is short, and asking twice would
  # double the wait on a cold service.
  $sList = runBoundedAll $sOllama @("list") 20
  if ($sList -eq "") { say "Chat model llama3.2: could not be checked just now. Press F12 in EdSharp; it offers to fetch the model if it is missing." ; return }
  if ($sList -match "llama3\.2") { say "Chat model llama3.2: installed. Press F12 to chat." }
  else { say "Chat model llama3.2: not downloaded. Press F12 in EdSharp and answer Yes when it offers to fetch it." }
}

function reportTranslationModel() {
  $sOllama = findExe "ollama"
  if ($sOllama -eq "") { return }
  $sList = runBoundedAll $sOllama @("list") 20
  if ($sList -match "qwen2\.5:7b") { say "Translation model qwen2.5:7b: installed. Alt+Shift+F7 will use it." }
  else { say "Translation model qwen2.5:7b: not installed. Alt+Shift+F7 uses llama3.2, quicker but less accurate." }
}

function reportCodeModel() {
  $sOllama = findExe "ollama"
  if ($sOllama -eq "") { return }
  $sList = runBoundedAll $sOllama @("list") 20
  if ($sList -match "qwen2\.5-coder") { say "Coding model qwen2.5-coder:7b: installed. F12 will use it for source files." }
  else { say "Coding model qwen2.5-coder:7b: not installed. F12 uses llama3.2 for source files too." }
}

function main() {
  try { New-Item -ItemType Directory -Force -Path $sLogDir | Out-Null } catch { }
  try { if (Test-Path -LiteralPath $sSummaryFile) { Remove-Item -LiteralPath $sSummaryFile -Force } } catch { }

  say ("EdSharp setup results  " + (Get-Date).ToString("yyyy-MM-dd HH:mm"))
  say ""

  # What the installer knew before the checkboxes ran, handed over in a
  # file so that ONE box tells the whole story instead of two telling
  # halves.
  if (Test-Path -LiteralPath $sResultsFile) {
    # Read with the system's own encoding, which is what the installer
    # wrote; a wrong guess here would turn the lines into nonsense.
    foreach ($sLine in @(Get-Content -LiteralPath $sResultsFile -Encoding Default -ErrorAction SilentlyContinue)) { say $sLine }
    try { Remove-Item -LiteralPath $sResultsFile -Force } catch { }
    say ""
  }

  # Recommended first -- the tools EdSharp itself uses -- then the ones
  # that serve some people and not others, in the same order as the
  # checkboxes, so the box reads as the page did.
  say "Components"
  $sPython = reportPython
  reportModule "PDF reader" $sPython "pymupdf4llm" "run installPdfTools.cmd in the EdSharp folder"
  reportWordNet $sPython
  reportTool "Git" "git" "run installGitHub.cmd in the EdSharp folder"
  reportTool "GitHub command line" "gh" "run installGitHub.cmd in the EdSharp folder"
  reportTool "Node.js" "node" "run installNode.cmd in the EdSharp folder"
  reportTool "Ollama" "ollama" "run installOllama.cmd in the EdSharp folder"
  reportModel
  reportTranslationModel
  reportCodeModel
  # Pandoc is installed automatically and is reported in the lines above,
  # so it is not repeated here.

  say ""
  say "Saved as $sSummaryFile"
  say "Full log: $sLogFile"
  say ""
  say "To start EdSharp, press Alt+Control+E."

  if (-not $bQuiet) {
    try {
      Add-Type -AssemblyName System.Windows.Forms
      [void][System.Windows.Forms.MessageBox]::Show(($lLines -join "`r`n"), "EdSharp Setup Results")
    } catch {
      # The box is a courtesy; the file is the record.
      try { Add-Content -LiteralPath $sLogFile -Value "[summary] the results box could not be shown: $($_.Exception.Message)" } catch { }
    }
  }
}

try {
  main
  exit 0
} catch {
  try {
    Add-Content -LiteralPath $sLogFile -Value "[summary] FAILED: $($_.Exception.Message)"
    Add-Type -AssemblyName System.Windows.Forms
    [void][System.Windows.Forms.MessageBox]::Show("The setup summary could not be completed:`r`n" + $_.Exception.Message + "`r`n`r`nThe log is:`r`n" + $sLogFile, "EdSharp Setup Results")
  } catch { }
  exit 0
}
