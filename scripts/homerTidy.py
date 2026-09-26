#!/usr/bin/env python3
r"""homerTidy.py -- tidy a Homer Tools project: the folder and the repository,
in one pass.

This replaces the two scripts that used to do it. homerTidy swept the working
folder; homerTidy swept the git repository. They asked the same question --
"does this file belong to the project?" -- of two places, and keeping them apart
meant two surveys, two plans, and two chances to disagree.

    homerTidy                   survey everything and print the plan. Changes
                                nothing.
    homerTidy --do-it           carry the plan out.
    homerTidy --do-it --no-push  carry it out locally; do not push.
    homerTidy --folder-only     survey or fix the folder, leaving git alone.
    homerTidy --repo-only       survey or fix the repository only.
    homerTidy --gitignore       write the whitelist .gitignore and stop.

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
internal notes), tagRelease and create<App>Repo (maintainer scripts), every
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
    "*.exe", "*.log", "*.obj", "*.pdb", "Version.cs", "__pycache__/",
    ".venv/", "build/", "create*Repo.cmd", "create*Repo.ps1", "dist/", "notes/",
    "exec/", "logs/", "self.htm", "self.md", "tagRelease.cmd", "tagRelease.ps1", "version.py",
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
c_sLogName = "homerTidy.log"
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
    r"^localfiles\.txt$", r"^self\.(md|htm)$",
    r"^(homertidy|tagrelease|summarizesetup)\.log$",
    r"^(build|create|new|clean|tidy)[a-z0-9_]*\.log$",
    r"^[a-z0-9_+-]+\.(cs|py|js|iss|ico|inix|manifest|config|lua)$",
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
    if oLog is None: return True
    oLog.write(sText + "\n")
    oLog.flush()
    return True


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
    logLine("RUN: " + " ".join(lsFull))
    try:
        oResult = subprocess.run(lsFull, capture_output=True, text=True, cwd=sRoot)
    except Exception as oError:
        logLine("RUN FAILED: %s" % oError)
        return (1, "")
    logLine("EXIT: %d" % oResult.returncode)
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


def surveyRepo(lsNamed):
    """Return (lsTrackedStrays, lsLargeInHistory, sStatus)."""
    iCode, sOut = runGit(["ls-files"], True)
    lsTracked = [s.strip() for s in sOut.splitlines() if s.strip()]
    lsTrackedStrays = [s for s in lsTracked if not belongsTo(s, lsNamed)]

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
        "# .gitignore for %s, generated by homerTidy on %s." % (appName(), datetime.date.today().isoformat()),
        "#",
        "# THIS IS A WHITELIST. Everything is ignored, and then exactly what",
        "# RepoFiles.txt names is put back. A file dropped into this folder is",
        "# invisible to git until somebody names it there, which is the point: the",
        "# file that gets pushed by accident is always the one nobody thought of.",
        "#",
        "# To add a file to the repository, add one line to RepoFiles.txt and run",
        "# homerTidy again. Do not edit this file by hand; it is overwritten.",
        "",
        "# Ignore everything.",
        "/*",
        "",
        "# Put back what the project names.",
    ]
    for sName in sorted(set(s.replace("\\", "/") for s in lsNamed), key=lambda s: s.lower()):
        sClean = sName.strip("/")
        if not sClean: continue
        lsLines.append("!/" + sClean + ("/" if sName.endswith("/") else ""))
    lsLines.append("!/.gitignore")
    lsLines.append("")
    lsLocal = namedByLocalFiles()
    if lsLocal:
        lsLines.append("# On this disk only, from LocalFiles.txt.")
        lsLines.extend(s.replace("\\", "/") for s in lsLocal)
        lsLines.append("")
    lsLines.append("# Never pushed, whatever RepoFiles.txt says: private notes, maintainer")
    lsLines.append("# scripts, logs, and anything a build makes.")
    for sPattern in c_lsNeverPushed:
        lsLines.append(sPattern)
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
    sText += "\n# Untracked by homerTidy on %s: these do not belong in the repository.\n" % \
             datetime.date.today().isoformat()
    sText += "notes/\n" if "notes/" not in lsExisting else ""
    for sLine in lsNew:
        sText += sLine.replace("\\", "/") + "\n"
    sText = sText.replace("\r\n", "\n").replace("\n", "\r\n")
    open(sPath, "wb").write(("\ufeff" + sText).encode("utf-8"))
    logLine("GITIGNORE: added %s" % countNoun(len(lsNew), "line"))
    return True


# --- the plan ---------------------------------------------------------------

def main():
    global oLog, sRoot
    oParser = argparse.ArgumentParser(
        description="Tidy a Homer Tools project folder and its repository.")
    oParser.add_argument("--do-it", action="store_true",
                         help="make the changes, rather than only describing them")
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
    logLine("homerTidy started %s" % datetime.datetime.now().isoformat(" ", "seconds"))
    logLine("Script: %s" % os.path.abspath(__file__))
    logLine("Python: %s" % sys.version.replace("\n", " "))
    logLine("Platform: %s" % platform.platform())
    logLine("Project: %s" % sRoot)
    logLine("Command line: %s" % " ".join(sys.argv))
    logLine("Settings: do-it=%s no-push=%s folder-only=%s repo-only=%s" %
            (dArguments.do_it, dArguments.no_push, dArguments.folder_only,
             dArguments.repo_only))

    bDoIt = dArguments.do_it
    sayLine("%s in %s" % (appName(), sRoot))

    if dArguments.gitignore:
        writeWhitelistGitignore()
        logLine("Finished %s" % datetime.datetime.now().isoformat(" ", "seconds"))
        return 0

    if not bDoIt: sayLine("This is the plan only. Nothing will be changed.")
    sayLine()

    lsNamed = namedByInstaller() + namedByRepoFiles() + namedByLocalFiles()
    logLine("Named by the project: %s" % countNoun(len(lsNamed), "entry", "entries"))
    if not lsNamed:
        sayLine("Nothing names the project's files: there is no <App>_setup.iss and")
        sayLine("no RepoFiles.txt here. Refusing to guess. Add one and run again.")
        return 1

    iChanges = 0

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
            for sPath in lsTrackedStrays: logLine("TRACKED STRAY: " + sPath)
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

    if not bDoIt:
        sayLine("Run it again with --do-it to carry this out.")
    else:
        sayLine("%s made." % countNoun(iChanges, "change"))
    logLine("Finished %s" % datetime.datetime.now().isoformat(" ", "seconds"))
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
