#!/usr/bin/env python3
r"""tidy.py -- tidy a Homer Tools project: the folder and the repository,
in one pass.

This replaces the two scripts that used to do it. tidy swept the working
folder; tidy swept the git repository. They asked the same question --
"does this file belong to the project?" -- of two places, and keeping them apart
meant two surveys, two plans, and two chances to disagree.

    tidy                   tidy everything: the folder and the repository.
    tidy --no-push         tidy, committing locally; do not push.
    tidy --folder-only     tidy the folder, leaving git alone.
    tidy --repo-only       tidy the repository only.
    tidy --gitignore       write the whitelist .gitignore and stop.

IT JUST DOES IT (1.43.9). tidy used to print a plan and change nothing until
it was run again with --do-it. Nothing it does is lost: a stray file is moved
into notes, which is on this disk and never in git, so a file tidy took is
found there; the only deletions are zero-byte files and things the build
fetches again. So the plan and the doing are one run, and the log records every
move. --do-it is still accepted, and changes nothing, so an old habit or an old
script does no harm.

WHAT BELONGS. Nothing is listed inside this script. A file belongs to the
project when the installer script names it, when RepoFiles.txt names it, when
LocalFiles.txt names it, or when it is one of the standing project files every
Homer project has.

LocalFiles.txt is the project's own customization: files that belong on THIS
disk but never in the repository -- a tutorial's working scripts, fetched voice
models, generated audio. Same form as RepoFiles.txt, one path or pattern a line,
and every line is also written into .gitignore as never pushed.

THE LAYOUT. A Homer development folder mirrors the installed tree: sources and
build files at the top, then configs, data, exec, help, logs, scripts and
templates. So the folder step also PUTS THINGS IN PLACE: a file at the top that
the installer takes from exec\ or help\ is moved there, a stray .log anywhere
goes to logs\, and exec\ and logs\ are never surveyed, being made rather than
authored. Those
three sources are read from the project folder itself, so this script never
needs editing and never needs to know which program it is tidying.

WHAT GETS PUSHED IS A WHITELIST, and RepoFiles.txt alone decides it. A
.gitignore that lists what to leave out can only ever be as complete as the last
time somebody remembered to add a line, and the file that gets pushed by
accident is always the one nobody thought of. So this script writes a .gitignore
that ignores EVERYTHING and then re-includes exactly what RepoFiles.txt names:

    /*
    !/.gitignore
    !/ReadMe.md
    !/CSharp/

Anything dropped into the folder afterwards -- a draft, a download, a log, a
note to yourself -- is invisible to git until somebody names it. Turning that
around, adding a file to the repository means adding one line to RepoFiles.txt,
which is also the line that documents why it is there.

NEVER PUSHED, whatever any file says: self.md and self.htm (the project's own
internal notes), release and create<App>Repo (maintainer scripts), every
.log, the notes folder, and the build products.

WHAT IT DOES TO THE FOLDER

  1. Deletes empty files. A zero-byte file is worse than a missing one: its
     absence would have told you something.
  2. Finds duplicates by content and keeps the one the project names, or the
     oldest when the project names neither.
  3. Puts misplaced files where the layout says: into exec\ or help\ when the
     installer takes them from there, and every stray .log into logs\.
  4. Moves everything else into notes\, under a subfolder named for what it is:
     notes\drafts, notes\mail, notes\archives, notes\other. Nothing is deleted
     in this step, because a file nobody named may still be wanted.

WHAT IT DOES TO THE REPOSITORY

  4. Untracks files that are tracked but do not belong, so they stop being
     pushed. They stay on disk; only git forgets them.
  5. Reports anything large in the history, whether still reachable or not,
     because that is what makes a clone slow and what GitHub complains about.
     Rewriting history is never done automatically: the report tells you the
     command and leaves the decision to you.
  6. Adds what it untracked to .gitignore, grouped under a dated comment, so
     the same files do not come back on the next commit.
  7. Commits, and pushes unless told not to.

A file that is large and fetched at run time -- a model, a converter, an
installer payload -- belongs in neither place: the app's own install script
should fetch it. This script reports such a file rather than guessing.

THE LOG. <App>-tidy-<date>-<time>.log is written in the project's logs folder, and holds the
environment, every setting, every command with its exit code, the whole survey,
and any traceback. The console gets the short version.
"""

import argparse
import datetime
import fnmatch
import hashlib
import os
import platform
import re
import shutil
import subprocess
import sys
import time
import traceback

def homerProjectRoot(sScriptDir):
    """The project folder: the script's own, or its parent when the script
    sits in scripts\\, tools\\ or exec\\ as the Homer layout puts it."""
    if os.path.basename(sScriptDir).lower() in ("scripts", "tools", "exec"):
        return os.path.dirname(sScriptDir)
    return sScriptDir


def homerLogPath(sScriptDir):
    """The Homer log path: <project>\\logs\\<App>-tidy-yyyyMMdd-HHmmss.log.

    One file per run, so an alphabetical sort is a chronological one and the
    whole folder can be zipped and sent. A single fixed name beside the script
    overwrites the evidence of the run before, which is exactly what you want
    to read when something has gone wrong twice.

    The project is the folder the script sits in, or its parent when the script
    is in scripts\\ -- the Homer layout puts tools there and logs one level up.
    """
    sRoot = homerProjectRoot(sScriptDir)
    sApp = os.path.basename(sRoot) or "Homer"
    sLogs = os.path.join(sRoot, "logs")
    os.makedirs(sLogs, exist_ok=True)
    sWhen = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    return os.path.join(sLogs, "%s-tidy-%s.log" % (sApp, sWhen))




c_iLargeBytes = 10 * 1024 * 1024        # what counts as large in the history

# Never pushed, whatever RepoFiles.txt says. Alphabetical, as every list in
# Homer code is. Each is either private to the developer, regenerated on every
# build, or a product rather than a source.
c_lsNeverPushed = [
    # *.jsb (1.43.38): a compiled JAWS script belongs to the JAWS version
    # that built it; the installer compiles on the user's machine.
    "*.exe", "*.jsb", "*.log", "*.obj", "*.pdb", "Version.cs", "__pycache__/",
    ".venv/", "build/", "create*Repo.cmd", "create*Repo.ps1", "dist/", "notes/",
    # /version.py, AT THE TOP ONLY (1.43.16): it is the file a Python app's
    # build generates beside its source. HomerView's NVDA add-on carries its
    # own homer\version.py, which the add-on imports, and a bare version.py
    # here untracked it.
    "exec/", "logs/", "self.htm", "self.md", "release.cmd", "release.ps1", "/version.py",
]

# Folders a build makes. The survey does not walk into them, because what is
# inside belongs to PyInstaller or the compiler rather than to the project.
c_lsSkipFolders = [".git", ".venv", "__pycache__", "build", "dist", "exec", "logs", "notes", "venv"]
# THE LOG GOES IN logs\, WITH A TIMESTAMP IN ITS NAME. Homer convention:
# <App>-<task>-yyyyMMdd-HHmmss.log, one per run, so an alphabetical sort is a
# chronological one and the whole folder can be zipped and sent. A single
# fixed name beside the script overwrites the evidence of the run before.
# THE LOG GOES IN logs\, WITH A TIMESTAMP IN ITS NAME. Homer convention:
# <App>-<task>-yyyyMMdd-HHmmss.log, one per run, so an alphabetical sort is a
# chronological one and the whole folder can be zipped and sent. A single
# fixed name beside the script overwrites the evidence of the run before.
c_lsInstallerExcludes = [".git", ".venv", "__pycache__", "*.pyc", "venv"]   # build leftovers no installer may ship
c_sLogName = "tidy.log"
c_sNotes = "notes"

# The standing files of a Homer project, whether or not the installer names
# them. Patterns are matched against the name, case-insensitively.
c_lsStandingNames = [
    r"^readme\.(md|htm)$", r"^history\.(md|htm)$", r"^license\.(md|htm)$",
    r"^developer\.(md|htm)$", r"^hotkeys\.(md|htm)$", r"^announce\.(md|htm)$",
    r"^version\.txt$", r"^repofiles\.txt$", r"^\.gitignore$", r"^\.gitattributes$",
    r"^build[a-z0-9_]*\.(cmd|ps1|py)$", r"^create[a-z0-9_]*repo\.(cmd|ps1)$",
    r"^tagrelease\.(cmd|ps1)$", r"^homertidy\.(cmd|py)$",
    r"^install[a-z0-9_]*\.(cmd|ps1)$", r"^get[a-z0-9_]*\.(cmd|ps1)$",
    # A project's own script logs are rewritten on every run and are already
    # ignored by git, so they stay where the script that writes them expects.
    r"^localfiles\.txt$", r"^keepencoding\.txt$", r"^self\.(md|htm)$",
    r"^(homertidy|tagrelease|summarizesetup)\.log$",
    r"^(build|create|new|clean|tidy)[a-z0-9_]*\.log$",
    r"^[a-z0-9_+-]+\.(cs|py|js|iss|ico|inix|manifest|config|lua)$",
    # The fingerprint beside each tutorial's audio (1.60.3), 04_Open_and_Move.sha256: buildTutorials keeps it to know the
    # audio is current, and the repository carries it with the audio so a build elsewhere speaks nothing it need not.
    # One digit since the pattern of ten (1.62.0), 4_Open_and_Move.sha256; two-digit names are still kept until retired.
    r"^\d{1,2}_[a-z0-9_]+\.sha256$",
]

# Where an unnamed file goes when it is moved out of the way.
c_ldFolders = [
    ("archives", [".zip", ".7z", ".rar", ".gz", ".tar", ".cab", ".msi", ".exe"]),
    ("mail", [".eml", ".msg"]),
    ("drafts", [".md", ".htm", ".html", ".txt", ".docx", ".doc", ".rtf", ".pdf"]),
]

sScriptDir = homerProjectRoot(os.path.dirname(os.path.abspath(__file__)))
# THE PROJECT, NOT WHEREVER THE PROMPT HAPPENED TO BE. Running this from
# scripts\ surveyed scripts\ and called the project "scripts". The project is
# the folder the script belongs to, whatever directory it was launched from.
sRoot = homerProjectRoot(sScriptDir)
sLogPath = homerLogPath(sScriptDir)
oLog = None


# --- saying things ----------------------------------------------------------

def logLine(sText):
    """One event in the Homer log format (1.43.21), as log.py and Log.cs write:
    an ISO 8601 time with milliseconds and UTC offset, a five-character level,
    then the text; a line continuing the one above starts "| ", and no line is
    blank or unstamped. The level is ERROR or WARN when the text says so."""
    if oLog is None: return True
    import datetime as _datetime
    sText = (sText or "").replace("\r\n", "\n").rstrip("\n")
    if not sText.strip(): return True
    import re as _re
    sLevel = ("ERROR" if _re.search(r"\b(ERROR|FAIL|FAILED)\b", sText)
              else "WARN" if _re.search(r"\bWARN(ING)?\b", sText) else "INFO")
    # A LEADING LEVEL WORD IS THE LEVEL (1.43.33): "WARN: x" is written
    # "WARN  x", not "WARN  WARN: x".
    oLead = _re.match(r"(ERROR|WARN|WARNING)\b:?\s*", sText)
    if oLead: sText = sText[oLead.end():] or sText
    sPrefix = "%s %-5s " % (_datetime.datetime.now().astimezone().isoformat(timespec="milliseconds"), sLevel)
    lsOut = []
    for iAt, sOne in enumerate(sText.split("\n")):
        if iAt and not sOne.strip(): continue
        lsOut.append(sPrefix + ("| " if iAt else "") + sOne.rstrip())
    oLog.write("\n".join(lsOut) + "\n")
    oLog.flush()
    return True

def logValue(sValue):
    """A value as the Homer log format writes it: bare when it can be, quoted
    when it holds a space, a quote or an equals sign."""
    import re as _re
    s = "" if sValue is None else str(sValue)
    if s and not _re.search(r'[\s"=]', s): return s
    s = s.replace('"', '\\"')
    if s.endswith("\\"): s += "\\"
    return '"' + s + '"'


def logFact(sKey, sValue):
    """One environment fact: env key=value."""
    return logLine("env %s=%s" % (sKey, logValue(sValue)))

def logWindows():
    """The Windows actually running, worded as Log.cs and log.py word it:
    "Windows 11 25H2 (10.0.26200.9550)"."""
    try:
        import winreg as _winreg
        with _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion") as oKey:
            def read(sName):
                try: return str(_winreg.QueryValueEx(oKey, sName)[0])
                except OSError: return ""
            sBuild, sUbr, sDisplay = read("CurrentBuild"), read("UBR"), read("DisplayVersion")
        sName = "Windows 11" if sBuild.isdigit() and int(sBuild) >= 22000 else "Windows 10"
        return ("%s %s" % (sName, sDisplay)).strip() + " (10.0.%s%s)" % (sBuild, "." + sUbr if sUbr else "")
    except Exception:
        import platform as _platform
        return _platform.platform()




def sayLine(sText=""):
    print(sText)
    logLine("CONSOLE: " + sText)
    return True


def countNoun(iCount, sSingular, sPlural=None):
    """"1 file", "0 files", "2 files" -- the noun always matches the count."""
    if sPlural is None: sPlural = sSingular + "s"
    return "%d %s" % (iCount, sSingular if iCount == 1 else sPlural)


# --- running git ------------------------------------------------------------

def runGit(lsArgs, bQuiet=False):
    """Run git and return (iCode, sOutput). Never raises."""
    lsFull = ["git"] + lsArgs
    sCmd = " ".join(lsFull)
    nStarted = time.time()
    logLine("run start cmd=" + logValue(sCmd))
    try:
        oResult = subprocess.run(lsFull, capture_output=True, text=True, cwd=sRoot)
    except Exception as oError:
        logLine("ERROR run failed message=%s cmd=%s" % (logValue(str(oError)), logValue(sCmd)))
        return (1, "")
    logLine("run exit=%d ms=%d cmd=%s" % (oResult.returncode, (time.time() - nStarted) * 1000, logValue(sCmd)))
    if oResult.stdout and not bQuiet: logLine("STDOUT:\n" + oResult.stdout.rstrip())
    if oResult.stderr: logLine("STDERR:\n" + oResult.stderr.rstrip())
    return (oResult.returncode, oResult.stdout or "")


def isGitRepo():
    iCode, sOut = runGit(["rev-parse", "--is-inside-work-tree"], True)
    return iCode == 0


# --- what belongs -----------------------------------------------------------

def appName():
    """The app name, taken from the folder. C:\\EdSharp yields EdSharp."""
    return os.path.basename(os.path.normpath(sRoot)) or "this project"


def namedByInstaller():
    """Every file name the <App>_setup.iss mentions on a Source: line."""
    lsNames = []
    for sName in os.listdir(sRoot):
        if not sName.lower().endswith("_setup.iss"): continue
        try:
            sText = open(os.path.join(sRoot, sName), "rb").read().decode("utf-8-sig",
                                                                        errors="replace")
        except Exception as oError:
            logLine("Could not read %s: %s" % (sName, oError))
            continue
        for oMatch in re.finditer(r'Source:\s*"([^"]+)"', sText, re.IGNORECASE):
            lsNames.append(oMatch.group(1))
        lsNames.append(sName)
    return lsNames


def installerExcludes():
    """Adds Excludes for build leftovers to every Source: line of the <App>_setup.iss that takes in a folder with
    recursesubdirs: a Python virtual environment, compiled Python and Git folders. Returns the number of lines changed.

    WHY (1.60.2, 7 October 2026): DbDo's installer took in templates\\samples\\.venv, a virtual environment left by a
    build, and shipped it: 4,650 of its 4,805 files, 2,290 of them compiled Python. Tidy runs before release builds the
    installer, so a line repaired here is clean in that same build. Inno Setup matches an Excludes pattern against the end
    of each path, so ".venv" leaves out a folder of that name at any depth, with everything in it."""
    iChanged = 0
    for sName in sorted(os.listdir(sRoot)):
        if not sName.lower().endswith("_setup.iss"): continue
        sPath = os.path.join(sRoot, sName)
        bRaw = open(sPath, "rb").read()
        bBom = bRaw[:3] == b"\xef\xbb\xbf"
        sText = bRaw.decode("utf-8-sig", errors="replace")
        lsLines, lsOut = sText.split("\n"), []
        for sLine in lsLines:
            oSource = re.match(r'(?i)\s*Source:\s*"[^"]+"', sLine)
            if oSource and re.search(r"(?i)\brecursesubdirs\b", sLine):
                oExcludes = re.search(r'(?i)Excludes:\s*"([^"]*)"', sLine)
                lsHave = [x.strip() for x in oExcludes.group(1).split(",") if x.strip()] if oExcludes else []
                lsAdd = [x for x in c_lsInstallerExcludes if x.lower() not in [h.lower() for h in lsHave]]
                if lsAdd:
                    sAll = ",".join(lsHave + lsAdd)
                    sEnd = "\r" if sLine.endswith("\r") else ""
                    sBody = sLine[:-1] if sEnd else sLine
                    sNew = re.sub(r'(?i)Excludes:\s*"[^"]*"', 'Excludes: "' + sAll + '"', sBody) if oExcludes else sBody.rstrip() + '; Excludes: "' + sAll + '"'
                    logLine("Installer %s: added Excludes %s to: %s" % (sName, ", ".join(lsAdd), sBody.strip()))
                    sLine = sNew + sEnd
                    iChanged += 1
            lsOut.append(sLine)
        if iChanged:
            with open(sPath, "wb") as oFile: oFile.write((b"\xef\xbb\xbf" if bBom else b"") + "\n".join(lsOut).encode("utf-8"))
    return iChanged


def namedByRepoFiles():
    """Every line of RepoFiles.txt that is not blank or a comment."""
    sPath = os.path.join(sRoot, "RepoFiles.txt")
    if not os.path.exists(sPath): return []
    lsNames = []
    for sLine in open(sPath, "rb").read().decode("utf-8-sig", errors="replace").splitlines():
        sLine = sLine.strip()
        if sLine == "" or sLine.startswith("#") or sLine.startswith(";"): continue
        lsNames.append(sLine)
    return lsNames


def namedByLocalFiles():
    """Every line of LocalFiles.txt: on this disk, never in the repository."""
    sPath = os.path.join(sRoot, "LocalFiles.txt")
    if not os.path.exists(sPath): return []
    lsNames = []
    for sLine in open(sPath, "rb").read().decode("utf-8-sig", errors="replace").splitlines():
        sLine = sLine.strip()
        if sLine == "" or sLine.startswith("#") or sLine.startswith(";"): continue
        lsNames.append(sLine)
    return lsNames


def placementFor(sRelative, lsNamed):
    """Where a misplaced file belongs, or "" when it is not misplaced.

    A file at the top whose name the installer knows only inside the exec or
    help folder is in the wrong place: that is where the layout keeps it. A
    .log anywhere outside the logs folder belongs in logs.
    """
    sNorm = sRelative.replace("\\", "/")
    if sNorm.lower().endswith(".log") and not sNorm.lower().startswith("logs/"):
        return "logs"
    if "/" in sNorm: return ""
    # WHERE THE PROJECT SAYS A FILE LIVES, IT STAYS (1.43.11). EdSharp keeps
    # Tektosyne.dll and nvdaControllerClient.dll at the top, named there in
    # RepoFiles.txt, and the libraries its build fetches there too, under a
    # *.dll line in LocalFiles.txt; its installer ships them from exec, where
    # the build copies them. Seeing the installer's exec\ names, tidy "put
    # them in place" -- found exec already held the build's copy -- and moved
    # every one into notes, and git recorded the two carried libraries as
    # deleted. A file RepoFiles.txt or LocalFiles.txt names where it is is in
    # its place.
    if matchesAny(sNorm, exactNames(namedByRepoFiles())) or matchesAny(sNorm, namedByLocalFiles()):
        return ""
    for sNamed in lsNamed:
        sNamedNorm = sNamed.replace("\\", "/")
        if "/" not in sNamedNorm or "*" in sNamedNorm: continue
        sFolder, sName = sNamedNorm.rsplit("/", 1)
        if sName.lower() == sNorm.lower() and sFolder.lower() in ("exec", "help"):
            return sFolder
    return ""


def belongsTo(sRelative, lsNamed):
    """Does this path belong to the project?"""
    sLower = sRelative.replace("\\", "/").lower()
    sName = os.path.basename(sLower)

    if sLower.startswith(".git/") or sLower.startswith(c_sNotes + "/"): return True
    for sPattern in c_lsStandingNames:
        if re.match(sPattern, sName): return True
    for sNamed in lsNamed:
        sNamedLower = sNamed.replace("\\", "/").lower()
        if sNamedLower == sLower or sNamedLower == sName: return True
        # A wildcard in the installer, such as Samples\* or *.dll.
        if "*" in sNamedLower:
            sRegex = "^" + re.escape(sNamedLower).replace(r"\*", ".*") + "$"
            if re.match(sRegex, sLower) or re.match(sRegex, sName): return True
        # A folder named wholesale.
        if sNamedLower.endswith("/") and sLower.startswith(sNamedLower): return True
    return False


# FETCHED THINGS ARE DELETED, NEVER ARCHIVED (25 Sep 2026). On this date a
# tidy found the tutorial voices left in an app's scripts\voices -- piper,
# its models, a sherpa-onnx package -- called all 486 files strays, moved
# them into notes\other, and committed them. A thing the build fetches from
# the web belongs in neither the folder nor the history: it is deleted, and
# the log says so. The patterns are the engines and models the kit fetches.
c_lsRefetchable = ["*.lib", "*.onnx", "*.onnx.json", "*.ort", "*.tar.bz2", "espeak-ng-data",
                   "espeak-ng.dll", "kokoro-*", "onnxruntime*", "piper", "piper.exe",
                   "piper_phonemize.dll", "sherpa-onnx*", "voices"]


def isRefetchable(sRelative):
    """True when any part of the path names something the build fetches."""
    lsParts = re.split(r"[\\/]", sRelative)
    for sPart in lsParts:
        for sPattern in c_lsRefetchable:
            if fnmatch.fnmatch(sPart.lower(), sPattern.lower()): return True
    return False


def deleteRefetchable(sRelative):
    sPath = os.path.join(sRoot, sRelative)
    try:
        if os.path.isdir(sPath): shutil.rmtree(sPath)
        elif os.path.exists(sPath): os.remove(sPath)
        logLine("DELETED FETCHED: %s (the build fetches it again when needed)" % sRelative)
        return True
    except Exception as oError:
        logLine("COULD NOT DELETE %s: %s" % (sRelative, oError))
        return False


def noteFolderFor(sName):
    """Which notes subfolder an unnamed file goes to."""
    sExt = os.path.splitext(sName)[1].lower()
    for sFolder, lsExts in c_ldFolders:
        if sExt in lsExts: return sFolder
    return "other"


def hashOf(sPath):
    oHash = hashlib.sha256()
    with open(sPath, "rb") as oFile:
        while True:
            binChunk = oFile.read(1 << 20)
            if not binChunk: break
            oHash.update(binChunk)
    return oHash.hexdigest()


# --- the survey -------------------------------------------------------------

def surveyFolder(lsNamed):
    """Return (lsEmpty, ldDuplicates, lsStrays, ltPlace)."""
    lsEmpty = []
    lsStrays = []
    ltPlace = []
    dByHash = {}

    # AN EMPTY LOG IS DELETED TOO (1.43.22). logs is not surveyed -- every file
    # there is a run's record -- but a zero-byte one records nothing, and the
    # kit's release refused to publish over three of them in C:\\HomerDev\\logs.
    sLogs = os.path.join(sRoot, "logs")
    if os.path.isdir(sLogs):
        for sName in sorted(os.listdir(sLogs)):
            sFull = os.path.join(sLogs, sName)
            if os.path.isfile(sFull) and os.path.getsize(sFull) == 0:
                lsEmpty.append(os.path.relpath(sFull, sRoot))

    for sDirPath, lsDirs, lsFiles in os.walk(sRoot):
        lsDirs[:] = [s for s in lsDirs if s.lower() not in c_lsSkipFolders]
        for sName in sorted(lsFiles):
            sFull = os.path.join(sDirPath, sName)
            sRelative = os.path.relpath(sFull, sRoot)
            if os.path.getsize(sFull) == 0:
                lsEmpty.append(sRelative)
                continue
            sPlace = placementFor(sRelative, lsNamed)
            if sPlace:
                ltPlace.append((sRelative, sPlace))
                continue
            if not belongsTo(sRelative, lsNamed):
                lsStrays.append(sRelative)
            try:
                dByHash.setdefault(hashOf(sFull), []).append(sRelative)
            except Exception as oError:
                logLine("Could not hash %s: %s" % (sRelative, oError))

    ldDuplicates = []
    for sHash, lsPaths in sorted(dByHash.items()):
        if len(lsPaths) < 2: continue
        # A copy the project names is never a duplicate to remove, however many
        # there are. Sample folders carry their own scripts, and two samples
        # may well carry the same one -- each is where its database looks.
        lsKeepers = [s for s in lsPaths if belongsTo(s, lsNamed)]
        if lsKeepers:
            lsGone = [s for s in lsPaths if s not in lsKeepers]
            if lsGone: ldDuplicates.append((lsKeepers[0], lsGone))
            continue
        sKeep = sorted(lsPaths, key=lambda s: os.path.getmtime(os.path.join(sRoot, s)))[0]
        ldDuplicates.append((sKeep, [s for s in lsPaths if s != sKeep]))
    return (lsEmpty, ldDuplicates, lsStrays, ltPlace)


def matchesAny(sRelative, lsNames):
    """Does a path match any line of a name list: a path, a bare name, a
    pattern with *, or a folder ending in /?"""
    sLower = sRelative.replace("\\", "/").lower()
    sName = os.path.basename(sLower)
    for sNamed in lsNames:
        # A leading / anchors a name to the top of the project, as in .gitignore.
        bAnchored = sNamed.replace("\\", "/").startswith("/")
        sNamedLower = sNamed.replace("\\", "/").lower().lstrip("/")
        if sNamedLower == sLower or (not bAnchored and sNamedLower == sName): return True
        if "*" in sNamedLower:
            sRegex = "^" + re.escape(sNamedLower).replace(r"\*", ".*") + "$"
            if re.match(sRegex, sLower) or re.match(sRegex, sName): return True
        if sNamedLower.endswith("/") and sLower.startswith(sNamedLower): return True
    return False


# THE REPOSITORY HOLDS WHAT RepoFiles.txt NAMES, AND NOTHING ELSE (1.43.9).
# The folder test -- does a file belong to the project at all? -- also accepts
# what the installer and LocalFiles.txt name and the standing names, such as
# version.txt and any .cs. Used for the repository too, it called version.txt,
# Version.cs and the release scripts "named by the project", so DbDo's tidy on
# 26 September found "0 files tracked that the project does not name" while
# all four were tracked and RepoFiles.txt names none of them. A tracked file
# now stays only when RepoFiles.txt names it and LocalFiles.txt does not.
c_lsAlwaysTracked = [".gitattributes", ".gitignore", "KeepEncoding.txt", "LocalFiles.txt", "RepoFiles.txt"]


def exactNames(lsNames):
    """The lines of a name list that name one file: no wildcard, no folder."""
    return [s for s in lsNames if "*" not in s and not s.replace("\\", "/").endswith("/")]


def carveOuts(lsRepo, lsLocal):
    """RepoFiles.txt entries inside a folder kept off the repository.

    exec is on no machine's repository -- it is what a build makes -- but the
    kit's own libraries live there, exec\\CSharp and exec\\homer, because exec
    is where the code that runs belongs (1.43.22). A RepoFiles.txt line that
    lies inside a folder LocalFiles.txt or the never-pushed list keeps off is
    a deliberate exception to that folder, and wins."""
    lsFolders = [s.replace("\\", "/").strip("/").lower() + "/" for s in list(lsLocal) + c_lsNeverPushed
                 if s.replace("\\", "/").endswith("/") and "*" not in s]
    lsCarve = []
    for sEntry in lsRepo:
        sClean = sEntry.replace("\\", "/").lstrip("/")
        for sFolder in lsFolders:
            if sClean.lower().startswith(sFolder) and sClean.lower() != sFolder:
                lsCarve.append(sClean)
                break
    return lsCarve


def staysTracked(sRelative, lsRepo, lsLocal):
    """Does a tracked file belong in the repository?

    A NAME IN RepoFiles.txt OUTRANKS A PATTERN IN LocalFiles.txt (1.43.10).
    EdSharp's LocalFiles.txt says *.dll, because the build fetches its
    libraries, and its RepoFiles.txt names Tektosyne.dll, which cannot be
    fetched and so must be carried. Named exactly, it stays. Anything else
    LocalFiles.txt matches, or the kit never pushes (c_lsNeverPushed: the
    release scripts, Version.cs, any .exe), goes; so does anything
    RepoFiles.txt does not name at all.
    """
    if matchesAny(sRelative, exactNames(lsRepo)): return True
    if matchesAny(sRelative, carveOuts(lsRepo, lsLocal)): return True
    if matchesAny(sRelative, c_lsNeverPushed): return False
    if matchesAny(sRelative, lsLocal): return False
    return matchesAny(sRelative, lsRepo)


def surveyRepo(lsNamed):
    """Return (lsTrackedStrays, lsLargeInHistory, sStatus)."""
    iCode, sOut = runGit(["ls-files"], True)
    lsTracked = [s.strip() for s in sOut.splitlines() if s.strip()]
    lsRepo = namedByRepoFiles() + c_lsAlwaysTracked
    lsLocal = namedByLocalFiles()
    lsTrackedStrays = [s for s in lsTracked if not staysTracked(s, lsRepo, lsLocal)]

    lsLarge = []
    iCode, sObjects = runGit(
        ["rev-list", "--objects", "--all"], True)
    if iCode == 0 and sObjects:
        iCode, sSizes = runGit(
            ["cat-file", "--batch-all-objects", "--batch-check=%(objectname) %(objecttype) %(objectsize)"],
            True)
        dBig = {}
        for sLine in sSizes.splitlines():
            lsParts = sLine.split()
            if len(lsParts) != 3 or lsParts[1] != "blob": continue
            if int(lsParts[2]) >= c_iLargeBytes: dBig[lsParts[0]] = int(lsParts[2])
        if dBig:
            for sLine in sObjects.splitlines():
                lsParts = sLine.split(" ", 1)
                if len(lsParts) == 2 and lsParts[0] in dBig:
                    lsLarge.append((lsParts[1], dBig[lsParts[0]]))
            lsLarge = sorted(set(lsLarge), key=lambda t: -t[1])

    iCode, sStatus = runGit(["status", "--porcelain", "--untracked-files=no"], True)
    return (lsTrackedStrays, lsLarge, sStatus.strip())


# --- the fixes --------------------------------------------------------------

def moveToNotes(sRelative):
    sFrom = os.path.join(sRoot, sRelative)
    sFolder = os.path.join(sRoot, c_sNotes, noteFolderFor(sRelative))
    os.makedirs(sFolder, exist_ok=True)
    sTo = os.path.join(sFolder, os.path.basename(sRelative))
    iSuffix = 2
    while os.path.exists(sTo):
        sStem, sExt = os.path.splitext(os.path.basename(sRelative))
        sTo = os.path.join(sFolder, "%s_%d%s" % (sStem, iSuffix, sExt))
        iSuffix += 1
    shutil.move(sFrom, sTo)
    logLine("MOVED: %s -> %s" % (sRelative, os.path.relpath(sTo, sRoot)))
    return True


def writeWhitelistGitignore():
    """Write a .gitignore that ignores everything except what RepoFiles.txt names.

    This is the whole publishing policy in one file: a whitelist, generated, so
    a file nobody named cannot be pushed by accident. Anything already ignored
    by name in c_lsNeverPushed stays ignored even if RepoFiles.txt names it,
    because those are private or generated and naming one is a mistake.
    """
    lsNamed = namedByRepoFiles()
    if not lsNamed:
        sayLine("There is no RepoFiles.txt here, so the whitelist cannot be written.")
        logLine("No RepoFiles.txt; .gitignore left alone.")
        return False

    lsLines = [
        "# .gitignore for %s, generated by tidy on %s." % (appName(), datetime.date.today().isoformat()),
        "#",
        "# THIS IS A WHITELIST. Everything is ignored, and then exactly what",
        "# RepoFiles.txt names is put back. A file dropped into this folder is",
        "# invisible to git until somebody names it there, which is the point: the",
        "# file that gets pushed by accident is always the one nobody thought of.",
        "#",
        "# To add a file to the repository, add one line to RepoFiles.txt and run",
        "# tidy again. Do not edit this file by hand; it is overwritten.",
        "",
        "# Ignore everything.",
        "/*",
        "",
        "# Put back what the project names.",
    ]
    # A FILE IN A SUBFOLDER NEEDS ITS FOLDER PUT BACK FIRST (1.43.0). "/*"
    # ignores the folder help itself, and git never looks inside an ignored
    # folder, so "!/help/Announce.md" alone put back nothing: every file named
    # one by one in help\ or scripts\ -- as the rules for RepoFiles.txt ask --
    # was silently left out of the repository. So each such folder is put back
    # and its contents ignored again ("!/help/" then "/help/*"), and only then
    # are the named files put back. A folder named whole ("help/") is simply
    # put back, contents and all.
    lsClean = sorted(set(s.replace("\\", "/") for s in lsNamed), key=lambda s: s.lower())
    setWhole = set(s.strip("/").lower() for s in lsClean if s.endswith("/"))
    lsFolders = []
    for sName in lsClean:
        lsParts = sName.strip("/").split("/")
        for iDepth in range(1, len(lsParts)):
            sFolder = "/".join(lsParts[:iDepth])
            if sFolder not in lsFolders: lsFolders.append(sFolder)
    for sFolder in sorted(lsFolders, key=lambda s: (s.count("/"), s.lower())):
        bInsideWhole = any(sFolder.lower() == s or sFolder.lower().startswith(s + "/") for s in setWhole)
        if bInsideWhole: continue
        lsLines.append("!/" + sFolder + "/")
        lsLines.append("/" + sFolder + "/*")
    for sName in lsClean:
        sClean = sName.strip("/")
        if not sClean: continue
        lsLines.append("!/" + sClean + ("/" if sName.endswith("/") else ""))
    lsLines.append("!/.gitignore")
    # .gitattributes, which keeps the Homer CRLF line endings, is put back
    # whether or not RepoFiles.txt names it, as .gitignore is.
    lsLines.append("!/.gitattributes")
    lsLines.append("!/KeepEncoding.txt")
    lsLines.append("")
    lsLocal = namedByLocalFiles()
    if lsLocal:
        lsLines.append("# On this disk only, from LocalFiles.txt.")
        lsLines.extend(s.replace("\\", "/") for s in lsLocal)
        lsLines.append("")
        # A name in RepoFiles.txt outranks a pattern in LocalFiles.txt: a file
        # named exactly there is put back after the LocalFiles.txt patterns,
        # so *.dll there cannot take the one library the repository carries.
        lsBack = [s for s in exactNames(lsNamed) if matchesAny(s, lsLocal)]
        if lsBack:
            lsLines.append("# Named in RepoFiles.txt, so put back after the patterns above.")
            lsLines.extend("!/" + s.replace("\\", "/").strip("/") for s in lsBack)
            lsLines.append("")
    lsLines.append("# Never pushed, whatever RepoFiles.txt says: private notes, maintainer")
    lsLines.append("# scripts, logs, and anything a build makes.")
    for sPattern in c_lsNeverPushed:
        lsLines.append(sPattern)
    lsLines.append("")

    # A FOLDER KEPT OFF THE REPOSITORY WITH NAMED EXCEPTIONS INSIDE (1.43.22).
    # git never looks inside an ignored folder, so "exec/" would hide the kit's
    # exec\\CSharp and exec\\homer whatever came after it. Such a folder is
    # ignored by its contents instead ("/exec/*"), and the exceptions RepoFiles.txt
    # names inside it are put back last, where the last match wins.
    lsCarve = carveOuts(lsNamed, namedByLocalFiles())
    if lsCarve:
        setOuter = set()
        for sCarve in lsCarve:
            setOuter.add(sCarve.split("/")[0].lower())
        for iAt, sLine in enumerate(lsLines):
            sBare = sLine.strip().strip("/").lower()
            if sLine.strip().endswith("/") and sBare in setOuter:
                lsLines[iAt] = "/" + sLine.strip().strip("/") + "/*"
        lsLines.append("# Named in RepoFiles.txt inside a folder kept off the repository.")
        for sCarve in lsCarve:
            lsLines.append("!/" + sCarve)
        lsLines.append("")

    sPath = os.path.join(sRoot, ".gitignore")
    sText = "\r\n".join(lsLines)
    open(sPath, "wb").write(("\ufeff" + sText).encode("utf-8"))
    sayLine("Wrote a whitelist .gitignore naming %s." % countNoun(len(lsNamed), "entry", "entries"))
    logLine("WROTE: %s" % sPath)
    return True


def addToGitignore(lsPaths):
    """Add paths under one dated comment, without disturbing what is there."""
    if not lsPaths: return True
    sPath = os.path.join(sRoot, ".gitignore")
    sText = ""
    if os.path.exists(sPath):
        sText = open(sPath, "rb").read().decode("utf-8-sig", errors="replace")
    lsExisting = set(s.strip() for s in sText.splitlines())
    lsNew = [s for s in lsPaths if s.replace("\\", "/") not in lsExisting]
    if not lsNew: return True
    if sText and not sText.endswith("\n"): sText += "\n"
    sText += "\n# Untracked by tidy on %s: these do not belong in the repository.\n" % \
             datetime.date.today().isoformat()
    sText += "notes/\n" if "notes/" not in lsExisting else ""
    for sLine in lsNew:
        sText += sLine.replace("\\", "/") + "\n"
    sText = sText.replace("\r\n", "\n").replace("\n", "\r\n")
    open(sPath, "wb").write(("\ufeff" + sText).encode("utf-8"))
    logLine("GITIGNORE: added %s" % countNoun(len(lsNew), "line"))
    return True


def findKitFolder():
    """Where the kit is: the HomerDev variable, a HomerDev folder above or
    beside the project, or one at the top of a fixed drive."""
    lsDirs = [os.environ.get("HomerDev", "")]
    for sStart in (sRoot, os.getcwd(), os.path.dirname(os.path.abspath(__file__))):
        sDir = os.path.abspath(sStart)
        while True:
            lsDirs.append(os.path.join(sDir, "HomerDev"))
            sUp = os.path.dirname(sDir)
            if sUp == sDir: break
            sDir = sUp
    if os.name == "nt":
        import ctypes
        lsDirs += [sLetter + ":\\HomerDev" for sLetter in "CDEFGHIJKLMNOPQRSTUVWXYZ"
                   if ctypes.windll.kernel32.GetDriveTypeW(sLetter + ":\\") == 3]
    for sDir in lsDirs:
        if sDir and os.path.isfile(os.path.join(sDir, "version.txt")) and os.path.isdir(os.path.join(sDir, "Templates")):
            return os.path.abspath(sDir)
    return ""


# Files an app may legitimately hold byte for byte as the kit holds them: the
# shared scripts the build refreshes from the kit, the licence, the policy
# files that are the same everywhere. Never taken for strays.
c_lsNeverStray = ("scripts/", "License.md", "License.htm", ".gitattributes", ".gitignore", "KeepEncoding.txt", "version.txt")

# One-off repair scripts delivered during 5 and 6 October 2026, each of which
# did its work once; the build retires them so the release check, which reads
# scripts too, stops meeting their text.
c_lsRetiredScripts = ("scripts/repairFileDir.cmd", "scripts/restoreFileDir.cmd", "scripts/retireRepair.cmd", "scripts/removeKitFiles.cmd")


def removeStrayKitFiles(bRepo):
    """A HomerDev.zip unarchived into an app's folder leaves the kit's files
    beside the app's (6 October 2026: twice, into C:\\DbDo). A stray is a file
    at the same relative path as one of the kit's, byte for byte the kit's,
    and not one the app may share. Each is deleted, taken out of git's index,
    and logged; the folders left empty go too. The kit itself is never tidied
    this way."""
    sKit = findKitFolder()
    if not sKit or os.path.abspath(sRoot).lower() == sKit.lower():
        logLine("stray kit files: no kit found or this is the kit; skipped")
        return 0
    lsKitFiles = []
    for sLine in open(os.path.join(sKit, "RepoFiles.txt"), "r", encoding="utf-8-sig"):
        sName = sLine.strip()
        if not sName or sName[0] in "#;": continue
        sPath = os.path.join(sKit, sName.replace("/", os.sep))
        if sName.endswith("/") and os.path.isdir(sPath):
            for sDir, lsSub, lsNames in os.walk(sPath):
                for sFile in lsNames:
                    lsKitFiles.append(os.path.relpath(os.path.join(sDir, sFile), sKit).replace(os.sep, "/"))
        elif os.path.isfile(sPath):
            lsKitFiles.append(sName)
    iRemoved = 0
    lsDirs = set()
    for sRel in sorted(lsKitFiles):
        if any(sRel == s or (s.endswith("/") and sRel.startswith(s)) for s in c_lsNeverStray): continue
        sMine = os.path.join(sRoot, sRel.replace("/", os.sep))
        if not os.path.isfile(sMine): continue
        try:
            if hashOf(sMine) != hashOf(os.path.join(sKit, sRel.replace("/", os.sep))): continue
        except OSError:
            continue
        os.remove(sMine)
        iRemoved += 1
        lsDirs.add(os.path.dirname(sMine))
        logLine("stray kit file removed: " + sRel)
        if bRepo: runGit(["rm", "-q", "--cached", sRel], bQuiet=True)
    for sRel in c_lsRetiredScripts:
        sMine = os.path.join(sRoot, sRel.replace("/", os.sep))
        if os.path.isfile(sMine):
            os.remove(sMine); iRemoved += 1
            logLine("retired one-off script removed: " + sRel)
            if bRepo: runGit(["rm", "-q", "--cached", sRel], bQuiet=True)
    for sDir in sorted(lsDirs, key=len, reverse=True):
        while sDir and os.path.abspath(sDir).lower() != os.path.abspath(sRoot).lower():
            try:
                if not os.listdir(sDir): os.rmdir(sDir); logLine("empty folder removed: " + os.path.relpath(sDir, sRoot))
                else: break
            except OSError: break
            sDir = os.path.dirname(sDir)
    if iRemoved: sayLine("%s of the kit's, or of a one-off repair, removed from the folder." % countNoun(iRemoved, "file", "files"))
    return iRemoved


def removeKitSampleLeftovers(bRepo):
    """The kit's samples, unarchived into an app, leave what their build made beside the files removeStrayKitFiles
    takes away: the sample programs, PyInstaller's build and dist folders, their logs, the generated version files.
    None is in the kit's RepoFiles, so none is a stray by that test, and DbDo's installer shipped them -- 30 files,
    among them six sample programs (8 October 2026). An app's templates\\samples folder holding nothing but the kit's
    sample leftovers is removed whole, and logged; one holding anything else is left alone, with the reason logged.
    The kit itself is never tidied this way."""
    sKit = findKitFolder()
    if not sKit or os.path.abspath(sRoot).lower() == sKit.lower(): return 0
    sSamples = os.path.join(sRoot, "templates", "samples")
    if not os.path.isdir(sSamples): return 0
    c_lsLeftovers = ["FruitBasket*", "build", "dist", "logs", ".venv", "venv", "__pycache__", "Version.cs", "version.py",
                     "version.txt", "accept.inix", "*.spec", "*.pyc", "evidence-*.md"]
    lsOther = [s for s in os.listdir(sSamples) if not any(fnmatch.fnmatch(s.lower(), sPattern.lower()) for sPattern in c_lsLeftovers)]
    if lsOther:
        logLine("kit sample leftovers: templates\\samples kept, since it holds " + ", ".join(sorted(lsOther)[:5]) + ", not the kit's")
        return 0
    iFiles = sum(len(lsNames) for _, _, lsNames in os.walk(sSamples))
    shutil.rmtree(sSamples, ignore_errors=True)
    logLine("kit sample leftovers: templates\\samples removed, %d files the kit's samples had built" % iFiles)
    if bRepo: runGit(["rm", "-r", "-q", "--cached", "--ignore-unmatch", "templates/samples"], bQuiet=True)
    sTemplates = os.path.join(sRoot, "templates")
    if os.path.isdir(sTemplates) and not os.listdir(sTemplates):
        os.rmdir(sTemplates)
        logLine("empty folder removed: templates")
    sayLine("The kit's sample programs, left in templates\\samples by an earlier unarchive, removed: %s." % countNoun(iFiles, "file", "files"))
    return iFiles


# --- the plan ---------------------------------------------------------------


def loadKind():
    """The kit's kind.py (1.45.0): beside this script, or in the kit's scripts
    folder wherever the kit is (1.46.0): the HomerDev variable, then a folder
    named HomerDev above or beside where this runs, at any depth, then one at
    the top of any fixed drive. Without it the project is taken to be an app,
    as every script assumed before there were four kinds."""
    lsDirs = [os.path.dirname(os.path.abspath(__file__)), os.path.join(os.environ.get("HomerDev", ""), "scripts")]
    for sStart in (os.getcwd(), os.path.dirname(os.path.abspath(__file__))):
        sDir = os.path.abspath(sStart)
        while True:
            lsDirs.append(os.path.join(sDir, "scripts"))
            lsDirs.append(os.path.join(sDir, "HomerDev", "scripts"))
            sUp = os.path.dirname(sDir)
            if sUp == sDir: break
            sDir = sUp
    if os.name == "nt":
        import ctypes
        lsDirs += [sLetter + ":\\HomerDev\\scripts" for sLetter in "CDEFGHIJKLMNOPQRSTUVWXYZ"
                   if ctypes.windll.kernel32.GetDriveTypeW(sLetter + ":\\") == 3]
    for sDir in lsDirs:
        if sDir and os.path.isfile(os.path.join(sDir, "kind.py")):
            if sDir not in sys.path: sys.path.insert(0, sDir)
            import kind
            return kind.projectKind
    return lambda sFolder: ("app", "kind.py was not found, so taken to be an app")

def main():
    global oLog, sRoot
    oParser = argparse.ArgumentParser(
        description="Tidy a Homer Tools project folder and its repository.")
    oParser.add_argument("--do-it", action="store_true",
                         help="accepted for old habits; tidy always carries its plan out")
    oParser.add_argument("--no-push", action="store_true",
                         help="commit locally but do not push")
    oParser.add_argument("--folder-only", action="store_true",
                         help="leave git alone")
    oParser.add_argument("--repo-only", action="store_true",
                         help="leave the folder alone")
    oParser.add_argument("--gitignore", action="store_true",
                         help="write the whitelist .gitignore from RepoFiles.txt and stop")
    oParser.add_argument("--path", default="",
                         help="the project folder; the current directory by default")
    dArguments = oParser.parse_args()

    if dArguments.path: sRoot = os.path.abspath(dArguments.path)

    # THE LOG GOES IN THE PROJECT'S logs FOLDER, one file per run, named as the
    # program names its own: <App>-tidy-yyyyMMdd-HHmmss.log. An alphabetical
    # sort is a chronological one, and zipping logs gathers every session.
    global sLogPath
    sLogDir = os.path.join(sRoot, "logs")
    os.makedirs(sLogDir, exist_ok=True)
    sLogPath = os.path.join(sLogDir, "%s-tidy-%s.log" % (os.path.basename(sRoot.rstrip("\\/")),
                            datetime.datetime.now().strftime("%Y%m%d-%H%M%S")))
    oLog = open(sLogPath, "w", encoding="utf-8")
    logLine("tidy start pid=%d" % os.getpid())
    logFact("script", os.path.abspath(__file__))
    logFact("python", platform.python_version())
    logFact("windows", logWindows())
    logFact("project", sRoot)
    logFact("arguments", " ".join(sys.argv[1:]))
    logLine("settings no-push=%s folder-only=%s repo-only=%s" %
            (dArguments.no_push, dArguments.folder_only,
             dArguments.repo_only))

    bDoIt = True
    # FOUR KINDS OF HOMER RESOURCE (1.45.0). A page project that is not a git
    # repository has nothing for tidy to tidy: post stages and publishes it.
    sKind, sWhy = loadKind()(sRoot)
    logLine("setting kind=%s reason=%s" % (sKind, logValue(sWhy)))
    if sKind == "page" and not isGitRepo():
        sayLine("%s is a page project, not a repository: post publishes it, and tidy has nothing to do." % appName())
        logLine("tidy end")
        return 0
    sayLine("%s in %s" % (appName(), sRoot))
    if not dArguments.repo_only:
        # The step itself declines when this folder is the kit's own; the kind
        # is not trusted here, since the strays are what mislead it.
        try: removeStrayKitFiles(isGitRepo() and not dArguments.folder_only)
        except Exception as oError: logLine("stray kit files: skipped, " + str(oError))
        try: removeKitSampleLeftovers(isGitRepo() and not dArguments.folder_only)
        except Exception as oError: logLine("kit sample leftovers: skipped, " + str(oError))

    # THE BUILD SCRIPT TAKES ITS SHORT NAME HERE (1.43.56). Since 1.43.55 an
    # app's build script is build.cmd, and check fails build<App>.cmd. Tidy runs
    # before every push, so when an old name is still here it runs the kit's
    # renameBuild on this folder -- git mv, every reference updated, its own log
    # -- and the push that follows records the rename. Nothing is asked of you.
    sApp = os.path.basename(sRoot.rstrip("\\/"))
    lsOldBuild = [s for s in os.listdir(sRoot) if s.lower() in (("build" + sApp + ".cmd").lower(), ("build" + sApp + ".ps1").lower())]
    if sKind != "app": lsOldBuild = []  # the kit renames its own, in build.py (1.43.58); a page or collection has no build
    # EVERY HOMER APP IS MIT (1.50.2): the kit's relicense runs before the
    # push, as an old build name is renamed. With --if-needed it does nothing,
    # and writes no log, when License.md and everything else already say MIT.
    if sKind == "app" and not dArguments.repo_only:
        sRelicense = ""
        lsKitDirs = [os.environ.get("HomerDev", "")]
        try:
            import kind
            lsKitDirs.append(kind.findKit([sRoot]))
        except Exception:
            pass
        lsKitDirs.append(os.path.join(os.path.dirname(sRoot.rstrip("\\/")), "HomerDev"))
        for sKit in lsKitDirs:
            if sKit and os.path.isfile(os.path.join(sKit, "scripts", "relicense.py")):
                sRelicense = os.path.join(sKit, "scripts", "relicense.py")
                break
        logLine("relicense: %s" % (sRelicense or "not found"))
        if sRelicense:
            oResult = subprocess.run([sys.executable, sRelicense, sRoot, "--if-needed"], capture_output=True, text=True)
            for sLine in (oResult.stdout + oResult.stderr).splitlines(): logLine("relicense: " + sLine)
            logLine("relicense exit %d" % oResult.returncode)
            if oResult.stdout.strip(): sayLine("Relicensed under the MIT License, as every Homer app is; its log says what changed.")
    if lsOldBuild and not dArguments.repo_only:
        # The kit wherever it is (1.46.0): kind.py's findKit, which assumes
        # only Windows and the folder name, never a drive or a depth.
        lsKit = [os.environ.get("HomerDev", "")]
        try:
            import kind
            lsKit.append(kind.findKit([sRoot]))
        except Exception:
            pass
        lsKit.append(os.path.join(os.path.dirname(sRoot.rstrip("\\/")), "HomerDev"))
        sRename = ""
        for sKit in lsKit:
            if sKit and os.path.isfile(os.path.join(sKit, "scripts", "renameBuild.py")):
                sRename = os.path.join(sKit, "scripts", "renameBuild.py")
                break
        logLine("old build script names: %s; renameBuild: %s" % (", ".join(lsOldBuild), sRename or "not found"))
        if sRename:
            sayLine("Renaming %s to the short name, build.%s." % (" and ".join(lsOldBuild), "cmd and build.ps1" if len(lsOldBuild) > 1 else "cmd"))
            oResult = subprocess.run([sys.executable, sRename, sRoot], capture_output=True, text=True)
            for sLine in (oResult.stdout + oResult.stderr).splitlines():
                logLine("renameBuild: " + sLine)
            logLine("renameBuild exit %d" % oResult.returncode)
            if oResult.returncode != 0: sayLine("renameBuild could not finish; its log in logs says why.")
        else:
            sayLine("The build script still has its old name, and the kit's renameBuild was not found to fix it.")

    if dArguments.gitignore:
        writeWhitelistGitignore()
        logLine("tidy end")
        return 0

    sayLine()

    lsNamed = namedByInstaller() + namedByRepoFiles() + namedByLocalFiles()
    logLine("Named by the project: %s" % countNoun(len(lsNamed), "entry", "entries"))
    if not lsNamed:
        if sKind == "app":
            sayLine("Nothing names the project's files: there is no <App>_setup.iss and")
            sayLine("no RepoFiles.txt here. Refusing to guess. Add one and run again.")
        else:
            sayLine("Nothing names this %s's files: there is no RepoFiles.txt here." % sKind)
            sayLine("Refusing to guess. Add one and run again.")
        return 1

    iChanges = 0

    # ---- the installer: no build leftovers in it ----
    iInstaller = installerExcludes()
    if iInstaller:
        sayLine("Installer: %s now leave%s out build leftovers (.venv, __pycache__, compiled Python, .git)." % (countNoun(iInstaller, "folder line"), "s" if iInstaller == 1 else ""))
        iChanges += iInstaller

    # ---- the folder ----
    if not dArguments.repo_only:
        lsEmpty, ldDuplicates, lsStrays, ltPlace = surveyFolder(lsNamed)
        sayLine("Folder")
        sayLine("  %s to put in place" % countNoun(len(ltPlace), "file"))
        for sPath, sFolder in ltPlace: logLine("PLACE: %s -> %s/" % (sPath, sFolder))
        sayLine("  %s to delete" % countNoun(len(lsEmpty), "empty file"))
        sayLine("  %s to remove" % countNoun(sum(len(l) for _, l in ldDuplicates), "duplicate"))
        sayLine("  %s to move into notes" % countNoun(len(lsStrays), "file"))
        for sPath in lsEmpty: logLine("EMPTY: " + sPath)
        for sKeep, lsGone in ldDuplicates:
            logLine("DUPLICATE: keeping %s, removing %s" % (sKeep, ", ".join(lsGone)))
        for sPath in lsStrays:
            if isRefetchable(sPath): logLine("STRAY, FETCHED: %s -> deleted" % sPath)
            else: logLine("STRAY: %s -> notes/%s" % (sPath, noteFolderFor(sPath)))

        if bDoIt:
            # In place first. When the proper folder already holds a copy, that
            # copy is the one the build made or the installer ships, and the
            # loose one goes to notes rather than over it.
            for sPath, sFolder in ltPlace:
                sFrom = os.path.join(sRoot, sPath)
                sTo = os.path.join(sRoot, sFolder, os.path.basename(sPath))
                if not os.path.exists(sFrom): continue
                os.makedirs(os.path.join(sRoot, sFolder), exist_ok=True)
                if os.path.exists(sTo):
                    moveToNotes(sPath)
                else:
                    shutil.move(sFrom, sTo)
                    logLine("PLACED: %s -> %s" % (sPath, os.path.relpath(sTo, sRoot)))
                iChanges += 1
            for sPath in lsEmpty:
                os.remove(os.path.join(sRoot, sPath))
                logLine("DELETED EMPTY: " + sPath)
                iChanges += 1
            for sKeep, lsGone in ldDuplicates:
                for sPath in lsGone:
                    if os.path.exists(os.path.join(sRoot, sPath)):
                        moveToNotes(sPath)
                        iChanges += 1
            for sPath in lsStrays:
                if not os.path.exists(os.path.join(sRoot, sPath)): continue
                if isRefetchable(sPath):
                    if deleteRefetchable(sPath): iChanges += 1
                else:
                    moveToNotes(sPath)
                    iChanges += 1
            # And anything fetched that an earlier tidy archived into notes.
            sNotes = os.path.join(sRoot, "notes")
            if os.path.isdir(sNotes):
                for sDir, lsDirs, lsFiles in os.walk(sNotes):
                    for sName in lsFiles:
                        sRel = os.path.relpath(os.path.join(sDir, sName), sRoot)
                        if isRefetchable(sRel) and deleteRefetchable(sRel): iChanges += 1
        sayLine()

    # ---- the repository ----
    if not dArguments.folder_only:
        if not isGitRepo():
            sayLine("This folder is not a git repository, so there is nothing to tidy there.")
            logLine("Not a git repository.")
        else:
            lsTrackedStrays, lsLarge, sStatus = surveyRepo(lsNamed)
            sayLine("Repository")
            sayLine("  %s tracked that the project does not name" %
                    countNoun(len(lsTrackedStrays), "file"))
            sayLine("  %s larger than 10 MB in the history" %
                    countNoun(len(lsLarge), "object"))
            for sPath in lsTrackedStrays: logLine("TRACKED STRAY, to be untracked (the file stays on disk): " + sPath)
            for sPath, iSize in lsLarge:
                logLine("LARGE IN HISTORY: %s, %.1f MB" % (sPath, iSize / 1048576.0))

            if lsLarge:
                sayLine()
                sayLine("  A large object stays in the history until the history is rewritten,")
                sayLine("  which this script will not do for you. When you want that:")
                sayLine("    git filter-repo --path <file> --invert-paths")
                sayLine("  Something large that the program fetches at run time belongs in")
                sayLine("  neither the folder nor the history; let the install script get it.")

            # NOTHING IS STAGED WITHOUT A WHITELIST (25 Sep 2026). With no
            # RepoFiles.txt, "git add -A" swept 480 fetched files into a commit.
            # The whitelist .gitignore is what makes add -A safe, and it can be
            # written only from RepoFiles.txt; without one the repository is
            # left exactly as it was.
            bWhitelist = os.path.isfile(os.path.join(sRoot, "RepoFiles.txt"))
            if bDoIt and not bWhitelist:
                sayLine("  No RepoFiles.txt here, so nothing was staged, committed or pushed.")
                sayLine("  RepoFiles.txt names what the repository carries; add it and run again.")
                logLine("git phase skipped: no RepoFiles.txt")
            if bDoIt and bWhitelist:
                # The whitelist is rewritten on every pass, so RepoFiles.txt and
                # .gitignore cannot drift apart.
                writeWhitelistGitignore()
            if bDoIt and bWhitelist and lsTrackedStrays:
                for sPath in lsTrackedStrays:
                    runGit(["rm", "--cached", "--quiet", sPath])
                    iChanges += 1
                runGit(["add", "-A"])
                iCode, sOut = runGit(["diff", "--cached", "--quiet"])
                if iCode != 0:
                    runGit(["commit", "-m", "%s: tidy the repository" % appName()])
                    if not dArguments.no_push:
                        iCode, sOut = runGit(["push"])
                        if iCode != 0:
                            sayLine("  The push failed. Nothing local was lost; see the log.")
            sayLine()

    sayLine("%s made. Anything moved is in notes, which git never takes; the log names each move." %
            countNoun(iChanges, "change"))
    logLine("tidy end")
    return 0


if __name__ == "__main__":
    iCode = 1
    try:
        iCode = main()
    except Exception:
        try:
            logLine("TRACEBACK:\n" + traceback.format_exc())
        except Exception:
            pass
        print("Something went wrong. The details are in %s." % sLogPath)
    finally:
        if oLog is not None: oLog.close()
    sys.exit(iCode)
