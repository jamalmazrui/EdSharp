#!/usr/bin/env python3
r"""checkHomerApp.py -- gather evidence about a Homer app, without looking at it.

WHAT THIS IS FOR
================

A sighted developer checks by looking: skim the diff, glance at the window,
notice that a control is in the wrong place. A blind developer checks by
instrumenting. In AI-assisted development that is not an accommodation, it is
sound engineering, because nobody can eyeball generated code fast enough to
trust it.

This script is the instrument. It runs every check that can be made to answer
yes or no, records the command and the exit code for each, and writes an
evidence report saying three things:

    what was verified        a check ran and passed, and here is the command
    what was not checked     a check was skipped, and here is why
    what remains uncertain   the things no script can settle

The last list is the honest part. A report that claims everything is fine is
worth nothing; a report that says what it did not look at is worth a great deal.

USAGE

    checkHomerApp                    check the app in this folder
    checkHomerApp --path C:\JobDo    check another one
    checkHomerApp --build            build it first, and count that as evidence
    checkHomerApp --quiet            the report only, no console summary

WHAT IT CHECKS

  documents   the standard set, each with its .htm
  encoding    UTF-8 with a BOM and CRLF, except .cmd and .bat
  empty       no zero-byte file, because absence is more honest than emptiness
  version     version.txt present and agreeing with the installer script
  publish     RepoFiles.txt present and .gitignore generated from it
  logging     the program opens a log at startup
  naming      no accessible name set to a caption a screen reader already reads
  keys        no Alt+Control binding, no letter claimed twice in one dialog
  build       the build script runs and returns zero            (--build)
  smoke       the program starts, answers --help, and exits zero (--build)
  accept      every check in accept.inix, run with its expected exit code

ACCEPTANCE CRITERIA belong to the app, not to this script. Put them in
accept.inix beside the source, one section each:

    [check]
    Name   = the help text names every switch
    Run    = JobDo.exe --help
    Expect = 0
    Wants  = --source

Run is a command, Expect its exit code, and Wants an optional string its output
must contain. That is the whole language, on purpose: a criterion nobody can
read is a criterion nobody writes.

The report goes in the project's logs folder as <App>-evidence-<stamp>.md,
and the detailed log beside it as <App>-check-<stamp>.log. Run it in the
project folder or in its scripts folder; both mean the project.
"""

import argparse
import datetime
import glob
import os
import platform
import re
import subprocess
import sys
import traceback

# The standard document set. ReadMe and License sit at the top of the project;
# everything else lives in help, which is where the Homer layout puts documents.
c_lsDocumentsTop = ["License", "ReadMe"]
# THE DEFAULT DOCUMENT SET (22 Aug 2026): ReadMe and License at the top; the
# app's own guide, Developer, History and Hotkeys in help. Tutorials only when
# the app has tutorial scripts. Announce and FAQ are welcome, never required:
# a check that demanded FAQ refused a release on 25 Sep 2026.
c_lsDocumentsHelp = ["Developer", "History", "Hotkeys"]
c_lsSkipFolders = [".git", ".venv", "__pycache__", "build", "dist", "notes", "venv"]
c_lsGeneratedFiles = ["version.py", "version.cs"]   # written by every build
c_lsTextExt = (".cs", ".py", ".ps1", ".md", ".htm", ".inix", ".txt", ".iss", ".cmd", ".bat")

sScriptDir = os.path.dirname(os.path.abspath(__file__))


def projectRoot(sStart):
    """The project is the current folder, or its parent when the current
    folder is the project's scripts or exec folder. One rule for every tool;
    on 25 Sep 2026 this check, run from scripts, judged the scripts folder."""
    if os.path.basename(sStart).lower() in ("scripts", "exec", "tools"):
        sParent = os.path.dirname(sStart)
        if (os.path.isfile(os.path.join(sParent, "version.txt")) or os.path.isfile(os.path.join(sParent, "RepoFiles.txt"))
                or glob.glob(os.path.join(sParent, "*_setup.iss"))):
            return sParent
    return sStart


sRoot = projectRoot(os.getcwd())
sLogDir = os.path.join(sRoot, "logs")
os.makedirs(sLogDir, exist_ok=True)
sStamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
sLogPath = os.path.join(sLogDir, "%s-check-%s.log" % (os.path.basename(sRoot), sStamp))
oLog = None
lsFindings = []          # (sName, sVerdict, sEvidence)


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
    if sPlural is None: sPlural = sSingular + "s"
    return "%d %s" % (iCount, sSingular if iCount == 1 else sPlural)


def finding(sName, sVerdict, sEvidence):
    """One check, its verdict, and what the verdict rests on.

    Three verdicts and no others. "pass" means a check ran and succeeded;
    "fail" means it ran and did not; "skip" means it could not run, and the
    evidence says why. Nothing is ever recorded as passing because it was not
    looked at.
    """
    lsFindings.append((sName, sVerdict, sEvidence))
    logLine("%-10s %-5s %s" % (sName, sVerdict.upper(), sEvidence))
    return True


def runCommand(lsArgs, sShell=""):
    """Run a command, log it with its exit code, return (iCode, sOutput)."""
    logLine("RUN: " + (sShell or " ".join(lsArgs)))
    try:
        if sShell:
            oResult = subprocess.run(sShell, shell=True, cwd=sRoot,
                                     capture_output=True, text=True, timeout=900)
        else:
            oResult = subprocess.run(lsArgs, cwd=sRoot,
                                     capture_output=True, text=True, timeout=900)
    except Exception as oError:
        logLine("RUN FAILED: %s" % oError)
        return (1, str(oError))
    logLine("EXIT: %d" % oResult.returncode)
    sOut = (oResult.stdout or "") + (oResult.stderr or "")
    if sOut: logLine("OUTPUT:\n" + sOut[-4000:])
    return (oResult.returncode, sOut)


# --- the app ----------------------------------------------------------------

def appName():
    return os.path.basename(os.path.normpath(sRoot)) or "this project"


def sourceFiles():
    lsFiles = []
    for sDirPath, lsDirs, lsNames in os.walk(sRoot):
        lsDirs[:] = [s for s in lsDirs if s.lower() not in c_lsSkipFolders]
        for sName in sorted(lsNames):
            if sName.lower().endswith((".cs", ".py")):
                lsFiles.append(os.path.join(sDirPath, sName))
    return lsFiles


def isLibrary(sPath, sText):
    """Is this one of the kit's own modules rather than the app's code?

    A shared class sets accessible names and names key combinations on purpose;
    flagging those would make the checker cry wolf, and a checker that cries
    wolf is worse than none, because people learn to ignore it.
    """
    sShown = os.path.relpath(sPath, sRoot).replace(os.sep, "/").lower()
    if sShown.startswith("csharp/") or sShown.startswith("homer/"): return True
    return "namespace Homer" in sText or "part of the shared Homer toolkit" in sText


def codeLines(sText):
    """The text with comment-only lines removed.

    A comment that explains a rule is not a breach of it. KeyName.cs says
    "Alt+Control+key is reserved" in a comment, and that sentence is the rule
    rather than a violation of it.
    """
    lsKept = []
    for sLine in sText.splitlines():
        sTrimmed = sLine.strip()
        if sTrimmed.startswith(("//", "#", ";", "rem ", "'")): continue
        lsKept.append(sLine)
    return "\n".join(lsKept)


def readText(sPath):
    try:
        return open(sPath, "rb").read().decode("utf-8-sig", errors="replace")
    except Exception:
        return ""


# --- the checks -------------------------------------------------------------

def checkDocuments():
    sApp = appName()
    lsHelp = sorted(c_lsDocumentsHelp + [sApp])
    if glob.glob(os.path.join(sRoot, "help", "Tutorial_*.inix")): lsHelp.append("Tutorials")
    lsWanted = ([(sRoot, s) for s in c_lsDocumentsTop] +
                [(os.path.join(sRoot, "help"), s) for s in lsHelp])
    lsMissing = [s for sFolder, s in lsWanted
                 if not os.path.exists(os.path.join(sFolder, s + ".md"))]
    lsNoHtm = [s for sFolder, s in lsWanted
               if os.path.exists(os.path.join(sFolder, s + ".md"))
               and not os.path.exists(os.path.join(sFolder, s + ".htm"))]
    if lsMissing or lsNoHtm:
        return finding("documents", "fail",
                       "missing: %s; no .htm for: %s" %
                       (", ".join(lsMissing) or "none", ", ".join(lsNoHtm) or "none"))
    return finding("documents", "pass", "%s present, each with its .htm" %
                   countNoun(len(lsWanted), "document"))


def namedByProject():
    """RepoFiles.txt entries, or None when there is no such file."""
    sPath = os.path.join(sRoot, "RepoFiles.txt")
    if not os.path.isfile(sPath): return None
    lsNamed = []
    for sLine in readText(sPath).splitlines():
        sLine = sLine.strip()
        if sLine and not sLine.startswith("#"): lsNamed.append(sLine.replace("\\", "/").lower())
    return lsNamed


def isNamed(sRelative, lsNamed):
    sLower = sRelative.replace("\\", "/").lower()
    for sNamed in lsNamed:
        if sNamed == sLower: return True
        if sNamed.endswith("/") and sLower.startswith(sNamed): return True
        if "*" in sNamed and re.match("^" + re.escape(sNamed).replace(r"\*", ".*") + "$", sLower): return True
    return False


def checkEncoding():
    # ONLY THE PROJECT'S OWN FILES, the ones RepoFiles.txt names. A stray in
    # the folder is homerTidy's business; 54 of the 72 faults that refused a
    # release on 25 Sep 2026 were strays the project never named.
    lsNamed = namedByProject()
    lsWrong = []
    iChecked = 0
    for sDirPath, lsDirs, lsNames in os.walk(sRoot):
        lsDirs[:] = [s for s in lsDirs if s.lower() not in c_lsSkipFolders]
        for sName in sorted(lsNames):
            if sName.lower() in c_lsGeneratedFiles: continue
            if not sName.lower().endswith(c_lsTextExt): continue
            sPath = os.path.join(sDirPath, sName)
            if lsNamed is not None and not isNamed(os.path.relpath(sPath, sRoot), lsNamed): continue
            sShown = os.path.relpath(sPath, sRoot).replace(os.sep, "/")
            try:
                binData = open(sPath, "rb").read()
            except Exception:
                continue
            if not binData: continue
            iChecked += 1
            bBom = binData.startswith(b"\xef\xbb\xbf")
            bWantBom = not sName.lower().endswith((".cmd", ".bat", "version.txt"))
            if bBom != bWantBom:
                lsWrong.append("%s: %s a byte order mark" %
                               (sShown, "should not have" if bBom else "needs"))
            if binData.count(b"\n") != binData.count(b"\r\n"):
                lsWrong.append("%s: not every line ending is CRLF" % sShown)
    for sLine in lsWrong: logLine("ENCODING: " + sLine)
    if lsWrong:
        return finding("encoding", "fail", "%s wrong, all listed in the log" %
                       countNoun(len(lsWrong), "file"))
    return finding("encoding", "pass", "%s in UTF-8 with a BOM and CRLF" %
                   countNoun(iChecked, "file"))


def checkEmpty():
    lsEmpty = []
    for sDirPath, lsDirs, lsNames in os.walk(sRoot):
        lsDirs[:] = [s for s in lsDirs if s.lower() not in c_lsSkipFolders]
        for sName in sorted(lsNames):
            sPath = os.path.join(sDirPath, sName)
            try:
                if os.path.getsize(sPath) == 0:
                    lsEmpty.append(os.path.relpath(sPath, sRoot))
            except OSError:
                pass
    if lsEmpty:
        return finding("empty", "fail", "zero bytes: " + ", ".join(lsEmpty[:10]))
    return finding("empty", "pass", "0 zero-byte files")


def checkVersion():
    sVersionPath = os.path.join(sRoot, "version.txt")
    if not os.path.exists(sVersionPath):
        return finding("version", "skip", "no version.txt; the build makes one")
    sVersion = readText(sVersionPath).strip()
    lsIss = glob.glob(os.path.join(sRoot, "*_setup.iss"))
    if not lsIss:
        return finding("version", "pass", "version.txt holds %s; no installer script here" % sVersion)
    sIss = readText(lsIss[0])
    if "version.txt" in sIss:
        return finding("version", "pass",
                       "version.txt holds %s, and %s reads that file rather than a literal"
                       % (sVersion, os.path.basename(lsIss[0])))
    return finding("version", "fail",
                   "%s carries its own version literal instead of reading version.txt"
                   % os.path.basename(lsIss[0]))


def checkPublish():
    bRepoFiles = os.path.exists(os.path.join(sRoot, "RepoFiles.txt"))
    sGitignore = readText(os.path.join(sRoot, ".gitignore"))
    if not bRepoFiles:
        return finding("publish", "fail", "no RepoFiles.txt, so nothing names what may be pushed")
    if "THIS IS A WHITELIST" not in sGitignore:
        return finding("publish", "fail",
                       ".gitignore is not the generated whitelist; run homerTidy --gitignore")
    return finding("publish", "pass", "RepoFiles.txt names the whitelist and .gitignore was generated from it")


def checkLogging():
    # Evidence of a log: the kit's Log.start, or a program that opens a file
    # under a logs folder itself -- HomerScribe writes its own, and was refused
    # a release on 25 Sep 2026 for not calling the kit's.
    lsWith = [s for s in sourceFiles()
              if re.search(r"\bLog\.start\s*\(|\blog\.start\s*\(", readText(s))
              or (re.search(r"[\"'][^\"']*logs[\"'\\/]", readText(s)) and re.search(r"\.log\b", readText(s)))]
    if not lsWith:
        return finding("logging", "fail",
                       "no source file calls Log.start or log.start, so a failure would leave no record")
    return finding("logging", "pass", "%s opens a session log" %
                   ", ".join(os.path.basename(s) for s in lsWith))


def checkNaming():
    """An accessible name equal to a caption is read twice by a screen reader."""
    lsBad = []
    for sPath in sourceFiles():
        sText = readText(sPath)
        if isLibrary(sPath, sText): continue
        sText = codeLines(sText)
        lsCaptions = set(re.findall(r'add\w*\(\s*"([^"]+)"', sText))
        lsCaptions |= set(re.findall(r'\.Text\s*=\s*"([^"]+)"', sText))
        lsPlain = set(s.replace("&", "").strip().rstrip(":") for s in lsCaptions)
        for sName in re.findall(r'AccessibleName\s*=\s*"([^"]+)"', sText):
            if sName.replace("&", "").strip().rstrip(":") in lsPlain:
                lsBad.append("%s: accessible name %r repeats a caption" %
                             (os.path.basename(sPath), sName))
    for sLine in lsBad: logLine("NAMING: " + sLine)
    if lsBad:
        return finding("naming", "fail", "%s would be announced twice" %
                       countNoun(len(lsBad), "control"))
    return finding("naming", "pass", "0 accessible names repeat a caption")


def checkKeys():
    """Alt+Control is reserved, and one dialog must not claim a letter twice."""
    lsBad = []
    for sPath in sourceFiles() + glob.glob(os.path.join(sRoot, "*.inix")):
        sText = readText(sPath)
        if isLibrary(sPath, sText): continue
        sText = codeLines(sText)
        sBase = os.path.basename(sPath)
        # ALT+CONTROL IS FOR DESKTOP SHORTCUTS -- with one family excepted (25 Sep
        # 2026): the navigation keys. Alt+Control with an arrow, Home, End, Page
        # Up or Page Down moves a cursor inside a window and takes nothing from
        # the desktop, which uses letters. A desktop shortcut's own letter --
        # Alt+Control+D opens DbDo -- is the sanctioned use and is not in source.
        c_lsNavigation = ("arrow", "arrows", "up", "down", "left", "right", "home", "end",
                          "pageup", "pagedown", "uparrow", "downarrow", "leftarrow", "rightarrow")
        for sKey in re.findall(r"\b(?:Alt\+Control|Control\+Alt)\+\w+", sText):
            if sKey.rsplit("+", 1)[1].lower() in c_lsNavigation: continue
            lsBad.append("%s: %s is reserved for Windows desktop shortcuts" % (sBase, sKey))
        # ACCESS LETTERS COMPETE ONLY WHERE THEY ARE PRESSED.
        #
        # A letter belongs to one menu or one dialog: File's O and Edit's O are
        # two different keys in two different places, and both are right.
        # Counting every "&X" in a source file works for an app with a single
        # dialog and fails every app with a menu bar -- DbDo, 175 items in eight
        # menus, produced 26 "problems", none of them real.
        #
        # So each label is counted against where it appears: the named container
        # when the call has one -- addItem(miFile, "&Open...") -- and otherwise
        # the method building that window. Labels in two different methods never
        # compete, because a person never sees them at once.
        dByOwner = {}
        sMethod = "(file)"
        for sLine in sText.splitlines():
            oDef = re.match(r"\s{0,8}(?:public |private |internal |protected |static |override |virtual |async )+[\w<>\[\],\s\.]+?\s(\w+)\s*\(", sLine)
            if oDef: sMethod = oDef.group(1)
            # A CAPTION IS SHORT. An ampersand inside a sentence of help text is
            # prose that happens to hold the character; a control's caption is
            # a few words. Only strings of forty characters or fewer are read
            # as captions, so prose stops being counted as trigger letters.
            for oHit in re.finditer(r'(?:\b\w+\(\s*(\w+)\s*,\s*)?"(?=[^"]{0,40}")[^"&]*&([A-Za-z])', sLine):
                # The first argument names a menu only when it looks like one --
                # miFile, menuMain. A title or a prompt passed first, such as
                # promptText(sTitle, "&Question"), is not a container, and two
                # dialogs that both take sTitle are two dialogs.
                sArg = oHit.group(1) or ""
                sOwner = sArg if re.match(r"(mi|menu|m)[A-Z]", sArg) or sArg.lower().startswith("menu") else sMethod
                dByOwner.setdefault(sOwner, {})
                dByOwner[sOwner][oHit.group(2).lower()] = dByOwner[sOwner].get(oHit.group(2).lower(), 0) + 1
        for sOwner in sorted(dByOwner):
            for sLetter, iCount in sorted(dByOwner[sOwner].items()):
                if iCount > 1:
                    lsBad.append("%s: the access key %s is claimed %d times in %s"
                                 % (sBase, sLetter, iCount, sOwner))
    for sLine in lsBad: logLine("KEYS: " + sLine)
    if lsBad:
        return finding("keys", "fail", "%s; every one is in the log" %
                       countNoun(len(lsBad), "key problem"))
    return finding("keys", "pass", "0 reserved combinations, 0 access keys claimed twice")


def checkBuild(bBuild):
    if not bBuild:
        return finding("build", "skip", "not asked for; run with --build to make the build itself evidence")
    # THE APP'S OWN BUILD SCRIPT: build<App>.cmd, named after the folder. Any
    # other build*.cmd -- buildTutorials.cmd, say -- is a tool, not the build;
    # on 25 Sep 2026 this ran buildTutorials.cmd with "nobump" as a script name.
    sOwn = os.path.join(sRoot, "build" + os.path.basename(sRoot) + ".cmd")
    lsBuild = [sOwn] if os.path.isfile(sOwn) else [s for s in glob.glob(os.path.join(sRoot, "build*.cmd"))
                                                     if os.path.basename(s).lower() not in ("buildtutorials.cmd",)]
    if not lsBuild:
        return finding("build", "skip", "no build script here")
    sScript = os.path.basename(lsBuild[0])
    iCode, sOut = runCommand([], sShell='"%s" nobump' % sScript)
    if iCode == 0:
        return finding("build", "pass", "%s returned 0" % sScript)
    return finding("build", "fail", "%s returned %d; its own log has the compiler output" % (sScript, iCode))


def checkSmoke(bBuild):
    if not bBuild:
        return finding("smoke", "skip", "not asked for; comes with --build")
    lsExe = [s for s in glob.glob(os.path.join(sRoot, "*.exe"))
             if not s.lower().endswith("_setup.exe")]
    if not lsExe:
        return finding("smoke", "skip", "no executable here to start")
    sExe = os.path.basename(lsExe[0])
    iCode, sOut = runCommand([], sShell='"%s" --help' % sExe)
    if iCode == 0 and sOut.strip():
        return finding("smoke", "pass", "%s --help returned 0 and wrote %s" %
                       (sExe, countNoun(len(sOut.splitlines()), "line")))
    if iCode == 0:
        return finding("smoke", "fail", "%s --help returned 0 but said nothing" % sExe)
    return finding("smoke", "fail", "%s --help returned %d" % (sExe, iCode))


def checkAccept():
    """The app's own acceptance criteria, which only the app can state."""
    sPath = os.path.join(sRoot, "accept.inix")
    if not os.path.exists(sPath):
        return finding("accept", "skip",
                       "no accept.inix, so nothing here states what done means")
    lsChecks = []
    dNow = None
    for sLine in readText(sPath).splitlines():
        sLine = sLine.strip()
        if not sLine or sLine.startswith(";") or sLine.startswith("#"): continue
        if sLine.startswith("[") and sLine.endswith("]"):
            dNow = {}
            if sLine[1:-1].strip().lower() == "check": lsChecks.append(dNow)
            continue
        if dNow is None or "=" not in sLine: continue
        sField, sValue = sLine.split("=", 1)
        dNow[sField.strip().lower()] = sValue.strip()

    if not lsChecks:
        return finding("accept", "skip", "accept.inix holds 0 checks")
    iPassed = 0
    for dCheck in lsChecks:
        sName = dCheck.get("name", dCheck.get("run", "unnamed"))
        sRun = dCheck.get("run", "")
        iExpect = int(dCheck.get("expect", "0") or 0)
        sWants = dCheck.get("wants", "")
        if not sRun:
            logLine("ACCEPT: %s has no Run line" % sName)
            continue
        iCode, sOut = runCommand([], sShell=sRun)
        bOk = (iCode == iExpect) and (sWants == "" or sWants in sOut)
        logLine("ACCEPT: %s -> exit %d (wanted %d)%s: %s" %
                (sName, iCode, iExpect,
                 ", output wanted %r" % sWants if sWants else "",
                 "pass" if bOk else "fail"))
        if bOk: iPassed += 1
    if iPassed == len(lsChecks):
        return finding("accept", "pass", "%s passed" % countNoun(iPassed, "acceptance check"))
    return finding("accept", "fail", "%d of %s passed; the log names each one" %
                   (iPassed, countNoun(len(lsChecks), "acceptance check")))


# --- the report -------------------------------------------------------------

def writeReport():
    """The evidence report: what was verified, what was not, what is uncertain."""
    sStamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    sPath = os.path.join(sLogDir, "%s-evidence-%s.md" % (os.path.basename(sRoot), sStamp))
    lsPassed = [t for t in lsFindings if t[1] == "pass"]
    lsFailed = [t for t in lsFindings if t[1] == "fail"]
    lsSkipped = [t for t in lsFindings if t[1] == "skip"]

    lsLines = [
        "---",
        'title: "Evidence report: %s"' % appName(),
        'date: "%s"' % datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "---",
        "",
        "# Evidence report: %s" % appName(),
        "",
        "Written by checkHomerApp. Every line below rests on a command that ran,",
        "and the command and its exit code are in the check log beside this report.",
        "",
        "## What was verified",
        "",
    ]
    lsLines += ["- **%s** -- %s" % (s, e) for s, v, e in lsPassed] or ["- nothing"]
    lsLines += ["", "## What failed", ""]
    lsLines += ["- **%s** -- %s" % (s, e) for s, v, e in lsFailed] or ["- nothing"]
    lsLines += ["", "## What was not checked", ""]
    lsLines += ["- **%s** -- %s" % (s, e) for s, v, e in lsSkipped] or ["- nothing"]
    lsLines += [
        "",
        "## What remains uncertain",
        "",
        "No script settles these, and a report that did not say so would be worse",
        "than no report:",
        "",
        "- **Whether the program does the right thing.** Every check above asks",
        "  whether it behaves as built, not whether what was built was wanted.",
        "- **What a screen reader actually says.** A passing naming check means no",
        "  caption is repeated in the source, not that the speech makes sense.",
        "  Somebody has to listen.",
        "- **Whether the keyboard flow is usable.** No conflict is not the same as",
        "  a good order.",
        "- **Whether generated code is secure or private.** Nothing here reads the",
        "  code for intent, credentials, or what it sends where.",
        "- **Whether it still works next month.** These checks ran today, against",
        "  today's dependencies.",
        "",
        "## How to read this",
        "",
        "A check that passed is evidence. A check that was skipped is not a pass,",
        "and it is listed separately for that reason. The honest summary of any",
        "run is the three lists together.",
        "",
    ]
    open(sPath, "wb").write(("\ufeff" + "\r\n".join(lsLines)).encode("utf-8"))
    return sPath


def main():
    global oLog, sRoot
    oParser = argparse.ArgumentParser(description="Gather evidence about a Homer app.")
    oParser.add_argument("--build", action="store_true",
                         help="build the app first, and count the build as evidence")
    oParser.add_argument("--path", default="", help="the app folder; this one by default")
    oParser.add_argument("--quiet", action="store_true", help="write the report, say little")
    dArguments = oParser.parse_args()
    if dArguments.path: sRoot = os.path.abspath(dArguments.path)

    oLog = open(sLogPath, "w", encoding="utf-8")
    logLine("checkHomerApp started %s" % datetime.datetime.now().isoformat(" ", "seconds"))
    logLine("Script: %s" % os.path.abspath(__file__))
    logLine("Python: %s" % sys.version.replace("\n", " "))
    logLine("Platform: %s" % platform.platform())
    logLine("App folder: %s" % sRoot)
    logLine("Command line: %s" % " ".join(sys.argv))
    logLine("Settings: build=%s quiet=%s" % (dArguments.build, dArguments.quiet))

    if not dArguments.quiet: sayLine("Checking %s in %s" % (appName(), sRoot))

    checkDocuments()
    checkEncoding()
    checkEmpty()
    checkVersion()
    checkPublish()
    checkLogging()
    checkNaming()
    checkKeys()
    checkBuild(dArguments.build)
    checkSmoke(dArguments.build)
    checkAccept()

    sReport = writeReport()
    iPassed = len([t for t in lsFindings if t[1] == "pass"])
    iFailed = len([t for t in lsFindings if t[1] == "fail"])
    iSkipped = len([t for t in lsFindings if t[1] == "skip"])
    if not dArguments.quiet:
        sayLine()
        sayLine("%s passed, %s failed, %s not checked." %
                (countNoun(iPassed, "check"), countNoun(iFailed, "check"),
                 countNoun(iSkipped, "check")))
        for sName, sVerdict, sEvidence in lsFindings:
            if sVerdict == "fail": sayLine("  failed: %s -- %s" % (sName, sEvidence))
        sayLine("The report is %s." % os.path.basename(sReport))
    logLine("Finished %s" % datetime.datetime.now().isoformat(" ", "seconds"))
    return 1 if iFailed else 0


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
