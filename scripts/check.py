#!/usr/bin/env python3
r"""check.py -- gather evidence about a Homer app, without looking at it.

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

    check                    check the app in this folder
    check --path C:\JobDo    check another one
    check --build            build it first, and count that as evidence
    check --quiet            the report only, no console summary

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
import time
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


def sayConsoleRunning(sCommand):
    """One short console line for a command about to run."""
    try:
        print("  running: " + sCommand[:100], flush=True)
    except Exception:
        pass
    return True


def runCommand(lsArgs, sShell=""):
    """Run a command, log it with its exit code, return (iCode, sOutput).

    NOTHING IS WAITED FOR FROM THE KEYBOARD (1.43.17). A command's input is
    empty, so a "pause" or a prompt in it returns at once rather than waiting
    unseen: HomerView's checkHomerViewQuality.cmd ends with "Press any key",
    its output was captured, and its release sat silent until it was stopped
    by hand. The console names each command as it starts, so a long one is
    seen to be running.
    """
    sCmd = sShell or " ".join(lsArgs)
    nStarted = time.time()
    logLine("run start cmd=" + logValue(sCmd))
    sayConsoleRunning(sShell or " ".join(lsArgs))
    try:
        if sShell:
            # A LINE THAT BEGINS "cmd /c" IS NOT WRAPPED IN A SECOND cmd /c.
            #
            # shell=True already runs the line through cmd /c. A line that
            # itself begins with "cmd /c if exist X (exit 0) else (exit 1)" was
            # therefore parsed twice, and the parentheses and the else came
            # apart on the second pass: the answer was 1 whatever was true.
            # EdSharp's seven acceptance checks and DbDo's nine all failed that
            # way on 25 September, on lines that were correct.
            #
            # So such a line is handed to cmd once, with /s so that the outer
            # quotes are stripped and nothing inside is re-parsed, and with
            # shell=False so Python does not add a wrapper of its own.
            import re as _re
            oCmd = _re.match(r"\s*cmd(?:\.exe)?\s+/c\s+(.*)$", sShell, _re.I | _re.S)
            if oCmd and os.name == "nt":
                oResult = subprocess.run('cmd /s /c "' + oCmd.group(1).strip() + '"', shell=False, cwd=sRoot,
                                         capture_output=True, text=True, timeout=900,
                                         stdin=subprocess.DEVNULL)
            else:
                oResult = subprocess.run(sShell, shell=True, cwd=sRoot,
                                         capture_output=True, text=True, timeout=900,
                                         stdin=subprocess.DEVNULL)
        else:
            oResult = subprocess.run(lsArgs, cwd=sRoot,
                                     capture_output=True, text=True, timeout=900,
                                         stdin=subprocess.DEVNULL)
    except Exception as oError:
        logLine("ERROR run failed message=%s cmd=%s" % (logValue(str(oError)), logValue(sCmd)))
        return (1, str(oError))
    logLine("run exit=%d ms=%d cmd=%s" % (oResult.returncode, (time.time() - nStarted) * 1000, logValue(sCmd)))
    sOut = (oResult.stdout or "") + (oResult.stderr or "")
    if sOut: logLine("OUTPUT:\n" + sOut[-4000:])
    return (oResult.returncode, sOut)


# --- the app ----------------------------------------------------------------

def appName():
    return os.path.basename(os.path.normpath(sRoot)) or "this project"


def sourceFiles():
    """The project's own .cs and .py files: those RepoFiles.txt names, when it
    exists. A leftover in the folder -- EdSharp's pre-kit Samples folder, its
    fruitBasket.cs still on disk though no longer tracked -- is tidy's
    business, as it is for the encoding check (1.43.11)."""
    lsNamed = namedByProject()
    lsFiles = []
    for sDirPath, lsDirs, lsNames in os.walk(sRoot):
        lsDirs[:] = [s for s in lsDirs if s.lower() not in c_lsSkipFolders]
        for sName in sorted(lsNames):
            if sName.lower().endswith((".cs", ".py")):
                sPath = os.path.join(sDirPath, sName)
                if lsNamed is not None and not isNamed(os.path.relpath(sPath, sRoot), lsNamed): continue
                lsFiles.append(sPath)
    return lsFiles


def isLibrary(sPath, sText):
    """Is this one of the kit's own modules rather than the app's code?

    A shared class sets accessible names and names key combinations on purpose;
    flagging those would make the checker cry wolf, and a checker that cries
    wolf is worse than none, because people learn to ignore it.
    """
    sShown = os.path.relpath(sPath, sRoot).replace(os.sep, "/").lower()
    if sShown.startswith(("exec/csharp/", "exec/python/", "csharp/", "homer/")): return True
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
    # the folder is tidy's business; 54 of the 72 faults that refused a
    # release on 25 Sep 2026 were strays the project never named.
    lsNamed = namedByProject()
    # KeepEncoding.txt names other people's files, which keep the encoding
    # they came with; fixEncoding leaves them alone, so the check does too.
    lsKeep = []
    sKeepPath = os.path.join(sRoot, "KeepEncoding.txt")
    if os.path.isfile(sKeepPath):
        for sLine in readText(sKeepPath).splitlines():
            sLine = sLine.strip()
            if sLine and not sLine.startswith(("#", ";")): lsKeep.append(sLine.replace("\\", "/").lower().lstrip("/"))
    lsWrong = []
    iChecked = 0
    for sDirPath, lsDirs, lsNames in os.walk(sRoot):
        lsDirs[:] = [s for s in lsDirs if s.lower() not in c_lsSkipFolders]
        for sName in sorted(lsNames):
            if sName.lower() in c_lsGeneratedFiles: continue
            if not sName.lower().endswith(c_lsTextExt): continue
            sPath = os.path.join(sDirPath, sName)
            if lsNamed is not None and not isNamed(os.path.relpath(sPath, sRoot), lsNamed): continue
            if lsKeep and isNamed(os.path.relpath(sPath, sRoot), lsKeep): continue
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
    """Zero-byte files among the project's own. logs is passed over (1.43.22):
    a run's record is not part of the project, and the kit's release was
    refused over three empty logs; tidy deletes those."""
    lsEmpty = []
    for sDirPath, lsDirs, lsNames in os.walk(sRoot):
        lsDirs[:] = [s for s in lsDirs if s.lower() not in c_lsSkipFolders and s.lower() != "logs"]
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
                       ".gitignore is not the generated whitelist; run tidy --gitignore")
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


def normalizeKey(sKey):
    """A key combination as a comparable string: modifiers sorted, lower case,
    Ctrl and Control alike -- "Ctrl+Alt+Shift+H" and "alt+control+shift+h"
    both become "alt+control+shift+h"."""
    lsParts = [s.strip().lower() for s in sKey.replace(" ", "").split("+") if s.strip()]
    if not lsParts: return ""
    lsParts = ["control" if s == "ctrl" else s for s in lsParts]
    return "+".join(sorted(lsParts[:-1]) + [lsParts[-1]])


def desktopShortcut():
    """The app's own desktop shortcut from its installer -- a #define HotKey
    or a HotKey: on an [Icons] line -- normalized; "" when there is none."""
    for sIss in glob.glob(os.path.join(sRoot, "*_setup.iss")):
        oMatch = re.search(r"(?im)(?:#define\s+s?HotKey\s+\"|\bHotKey\s*[:=]\s*\"?)((?:alt|ctrl|control|shift)(?:\+(?:alt|ctrl|control|shift))*\+\w+)", readText(sIss))
        if oMatch: return normalizeKey(oMatch.group(1))
    return ""


# Words that can stand where a declaration's type does but begin a statement.
c_lsNotTypes = ("await", "case", "catch", "else", "for", "foreach", "if", "lock", "new", "return",
                "sizeof", "switch", "throw", "typeof", "using", "while", "yield")


def checkLocal():
    r"""ONLY THE LOCAL TREE (1.43.49). A Homer app keeps its settings, data,
    scripts and logs under %LOCALAPPDATA%\<App>, never under %APPDATA% (the
    Roaming tree). Any use of the Roaming tree in the project's own code,
    installer or scripts fails -- except where it names JAWS's or NVDA's own
    folders, which live there and are theirs, not the app's. Comment-only lines
    do not count."""
    reRoaming = re.compile(r"SpecialFolder\.ApplicationData\b|\{userappdata\}|\$env:APPDATA\b|%APPDATA%|"
                           r"environ(?:\.get)?\s*[\(\[]\s*[\"']APPDATA[\"']|getenv\s*\(\s*[\"']APPDATA[\"']", re.I)
    reTheirs = re.compile(r"Freedom Scientific|\bJAWS\b|\bnvda\b", re.I)
    lsNamed = namedByProject()
    lsBad = []
    for sDirPath, lsDirs, lsNames in os.walk(sRoot):
        lsDirs[:] = [s for s in lsDirs if s.lower() not in c_lsSkipFolders]
        for sName in sorted(lsNames):
            if not sName.lower().endswith((".cs", ".py", ".iss", ".ps1", ".cmd", ".js")): continue
            sPath = os.path.join(sDirPath, sName)
            if lsNamed is not None and not isNamed(os.path.relpath(sPath, sRoot), lsNamed): continue
            sText = readText(sPath)
            if isLibrary(sPath, sText): continue
            lsLines = sText.splitlines()
            for iAt, sLine in enumerate(lsLines, 1):
                if sLine.strip().startswith(("//", "#", ";", "rem ", "REM ", "'", "*", "(*", "{")): continue
                # Exempt: JAWS's or NVDA's own folders, named on this line or the
                # next (a path is often split across two), and the one-time move
                # of an earlier version's files out of Roaming, whose code names
                # the Roaming tree as what it is.
                sNext = lsLines[iAt] if iAt < len(lsLines) else ""
                if re.search(r"roaming", sLine, re.I) and "%APPDATA%" not in sLine: continue
                if reRoaming.search(sLine) and not reTheirs.search(sLine) and not reTheirs.search(sNext):
                    lsBad.append("%s line %d: %s" % (os.path.relpath(sPath, sRoot), iAt, sLine.strip()[:120]))
    for sLine in lsBad: logLine("ROAMING: " + sLine)
    if lsBad:
        return finding("local", "fail", "%s use the Roaming tree (%%APPDATA%%); a Homer app keeps its files under %%LOCALAPPDATA%%" %
                       countNoun(len(lsBad), "line"))
    return finding("local", "pass", "0 lines use the Roaming tree")


def finishEntries(sText):
    """The finish page's entries in an installer script: each [Run] entry with
    postinstall, its continued lines joined."""
    oRun = re.search(r"(?mi)^\[Run\]\s*$", sText)
    if not oRun: return []
    oEnd = re.search(r"(?m)^\[[A-Za-z]+\]\s*$", sText[oRun.end():])
    sRun = sText[oRun.end(): oRun.end() + oEnd.start()] if oEnd else sText[oRun.end():]
    lsEntries = []
    sCurrent = ""
    for sLine in sRun.splitlines():
        sStrip = sLine.strip()
        if not sStrip or sStrip.startswith(";"):
            if sCurrent and not sCurrent.rstrip().endswith("\\"): lsEntries.append(sCurrent); sCurrent = ""
            continue
        sCurrent += " " + sStrip.rstrip("\\")
        if not sStrip.endswith("\\"): lsEntries.append(sCurrent); sCurrent = ""
    if sCurrent: lsEntries.append(sCurrent)
    return [s.strip() for s in lsEntries if "postinstall" in s.lower()]


def checkFinishPage():
    r"""THE FINISH PAGE'S RULES (1.43.51; help\FinishPage.md). In every installer
    script: an Install or Update box is ticked, a Reinstall box never is, and
    Launch is ticked; every component box says Install, Update or Reinstall;
    nothing starts NVDA (opening an .nvda-addon or running nvda.exe), and no
    label says NVDA must be running. EdSharp's and DbDo's installers broke the
    first rule on 30 September after being called migrated."""
    lsBad = []
    # The installer the build compiles: <App>_setup.iss at the top of the project.
    for sDirPath, lsNames in [(sRoot, os.listdir(sRoot))]:
        for sName in sorted(lsNames):
            if not sName.lower().endswith("_setup.iss"): continue
            sPath = os.path.join(sDirPath, sName)
            sRel = os.path.relpath(sPath, sRoot)
            for sEntry in finishEntries(readText(sPath)):
                def field(sName):
                    oMatch = re.search(sName + r':\s*("(?:[^"]|"")*"|[^;]+)', sEntry, re.I)
                    return oMatch.group(1).strip().strip('"') if oMatch else ""
                sCheck = field("Check"); sFlags = field("Flags").lower(); sDesc = field("Description"); sFile = field("Filename")
                bUnticked = "unchecked" in sFlags
                oVerb = re.search(r"is(Install|Update|Reinstall)|(?:Needs?)(Install|Update)|(IsCurrent|AreCurrent)|isModel(Install|Reinstall)", sCheck)
                sVerb = ""
                if oVerb:
                    lsGroups = oVerb.groups()
                    sVerb = lsGroups[0] or lsGroups[1] or ("Reinstall" if lsGroups[2] else "") or lsGroups[3] or ""
                elif re.match(r"^(Install|Update|Reinstall)\b", sDesc):
                    sVerb = sDesc.split()[0]
                sWhat = sDesc or sFile
                if sVerb in ("Install", "Update") and bUnticked: lsBad.append("%s: %s box not ticked: %s" % (sRel, sVerb, sWhat))
                if sVerb == "Reinstall" and not bUnticked: lsBad.append("%s: Reinstall box ticked: %s" % (sRel, sWhat))
                if re.search(r"launch", sDesc, re.I) and bUnticked: lsBad.append("%s: Launch box not ticked" % sRel)
                if not sVerb and not re.search(r"launch|open|guide|read|view", sDesc, re.I):
                    lsBad.append("%s: box without Install, Update or Reinstall: %s" % (sRel, sWhat))
                if re.search(r"nvda-addon|nvda\.exe|GetNvdaPath", sFile, re.I): lsBad.append("%s: box starts NVDA: %s" % (sRel, sFile))
                if re.search(r"must be running", sDesc, re.I): lsBad.append("%s: label says NVDA must be running" % sRel)
    for sLine in lsBad: logLine("FINISH PAGE: " + sLine)
    if lsBad:
        return finding("finish", "fail", "%s on the finish page" % countNoun(len(lsBad), "rule broken", "rules broken"))
    return finding("finish", "pass", "every finish-page box follows the rules")


def checkKeys():
    """Alt+Control is reserved, and one dialog must not claim a letter twice."""
    lsBad = []
    # The project's own .inix files only, as for the sources: FileDir's
    # pre-kit Hotkeys.inix at the top, superseded by configs\Hotkeys.inix,
    # still named keys the program no longer has (1.43.15).
    lsNamedInix = namedByProject()
    lsInix = [s for s in glob.glob(os.path.join(sRoot, "*.inix"))
              if lsNamedInix is None or isNamed(os.path.relpath(s, sRoot), lsNamedInix)]
    for sPath in sourceFiles() + lsInix:
        sText = readText(sPath)
        if isLibrary(sPath, sText): continue
        sText = codeLines(sText)
        sBase = os.path.basename(sPath)
        # An ampersand in Python is not an access key unless the file builds
        # WinForms controls (through pythonnet). An NVDA add-on's "&" is prose
        # or an operator, and HomerView's eight source files were being read
        # for trigger letters they cannot have.
        bCaptions = sPath.lower().endswith(".cs") or "System.Windows.Forms" in sText
        # ALT+CONTROL IS FOR DESKTOP SHORTCUTS -- with one family excepted (25 Sep
        # 2026): the navigation keys. Alt+Control with an arrow, Home, End, Page
        # Up or Page Down moves a cursor inside a window and takes nothing from
        # the desktop, which uses letters. A desktop shortcut's own letter --
        # Alt+Control+D opens DbDo -- is the sanctioned use and is not in source.
        c_lsNavigation = ("arrow", "arrows", "up", "down", "left", "right", "home", "end",
                          "pageup", "pagedown", "uparrow", "downarrow", "leftarrow", "rightarrow")
        # THE WHOLE COMBINATION IS READ (1.43.18): Alt+Control+Shift+H, not
        # "Alt+Control+Shift" with Shift taken for the key. And in Python the
        # keys are the gestures NVDA binds -- "kb:alt+control+shift+h" --
        # rather than every mention in a message or a docstring: HomerView
        # tells its user "Press Alt+Control+Shift+H" in eight places, and a
        # docstring still named a key the add-on no longer binds.
        if sPath.lower().endswith(".py"):
            lsCombos = ["+".join(p.capitalize() for p in s.split("+"))
                        for s in re.findall(r"(?i)\bkb:((?:alt\+control|control\+alt)(?:\+shift)?\+[^\"'\s,\]]+)", sText)]
        else:
            lsCombos = re.findall(r"\b(?:Alt\+Control|Control\+Alt)(?:\+Shift)?\+\w+", sText)
        for sKey in lsCombos:
            if sKey.rsplit("+", 1)[1].lower() in c_lsNavigation: continue
            # The app's own desktop shortcut, as its installer declares it,
            # is the sanctioned use: FileDir's Hotkeys.inix lists Alt+Control+F
            # because that is how FileDir is opened, and HomerView's NVDA
            # add-on binds Alt+Control+Shift+H, its own shortcut's key.
            if normalizeKey(sKey) == desktopShortcut(): continue
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
        if not bCaptions: sText = ""
        for sLine in sText.splitlines():
            oDef = re.match(r"\s{0,8}(?:public |private |internal |protected |static |override |virtual |async )+[\w<>\[\],\s\.]+?\s(\w+)\s*\(", sLine)
            # A PYTHON FUNCTION IS A WINDOW'S BUILDER TOO (1.43.0). Only a def at
            # the left margin starts a new owner: a nested def is a handler
            # inside the dialog being built, and its captions belong with it.
            # Without this, all of urlCheck.py's 7,000 lines were one owner.
            if not oDef: oDef = re.match(r"(?:async )?def (\w+)\s*\(", sLine)
            # A METHOD WITH NO MODIFIER STARTS AN OWNER TOO (1.43.14). FileDir
            # declares its handlers as "void menuEditRename_Click(object sender,
            # EventArgs e) {" at the left margin, with no public or private, and
            # none of them was seen: every caption after Delete_Recycle was
            # counted as Delete_Recycle's, 39 false "claimed twice" in one file.
            # A type and a name before "(", with no "=" ahead of it and not a
            # statement keyword, is a declaration.
            if not oDef:
                oDecl = re.match(r"\s{0,8}([\w<>\[\],\.]+)\s+(\w+)\s*\(", sLine)
                if oDecl and oDecl.group(1) not in c_lsNotTypes and "=" not in sLine[:oDecl.end()]:
                    oDef = re.match(r"\s{0,8}[\w<>\[\],\.]+\s+(\w+)\s*\(", sLine)
            if oDef: sMethod = oDef.group(1)
            # A CAPTION IS SHORT. An ampersand inside a sentence of help text is
            # prose that happens to hold the character; a control's caption is
            # a few words. Only strings of forty characters or fewer are read
            # as captions, so prose stops being counted as trigger letters.
            # An HTML entity -- &amp; &lt; &nbsp; -- is not a trigger letter
            # (1.43.0): a program that writes HTML reports is full of them.
            # A MENU ITEM'S VARIABLE NAMES ITS MENU (1.43.14). FileDir builds its
            # whole menu bar in one constructor with menuEditTagAll =
            # menu_Helper("Tag &All", ...) and menuHelpAbout =
            # menu_Helper("&About", ...): Tag All and About are in different
            # menus, and counting them in one owner made 16 false clashes. An
            # assignment to menuEdit... or miEdit... puts the caption in the
            # Edit menu; to menuEdit itself, on the menu bar.
            sLineOwner = ""
            oAssign = re.match(r"\s*(?:[\w<>\[\]]+\s+)?(?:this\.)?(menu|mi)([A-Z][a-z0-9]+)(\w*)\s*=", sLine)
            if oAssign:
                sLineOwner = "menu bar" if oAssign.group(3) == "" else "menu " + oAssign.group(2)
            # THE SAME CAPTION TWICE IS ONE CLAIM (1.43.14). Code that shows
            # ButtonDialog(..., {"&No", "&Yes"}) and then tests case "&No" or
            # sChoice == "&User" names one button several times; only two
            # DIFFERENT captions with one letter compete. Each letter keeps the
            # set of captions that claim it.
            for oHit in re.finditer(r'(?:\b\w+\(\s*(\w+)\s*,\s*)?"(?=([^"]{0,40})")[^"&]*&(?![A-Za-z]+;|#\d+;)([A-Za-z])', sLine):
                # The first argument names a menu only when it looks like one --
                # miFile, menuMain. A title or a prompt passed first, such as
                # promptText(sTitle, "&Question"), is not a container, and two
                # dialogs that both take sTitle are two dialogs.
                sArg = oHit.group(1) or ""
                sOwner = sArg if re.match(r"(mi|menu|m)[A-Z]", sArg) or sArg.lower().startswith("menu") else (sLineOwner or sMethod)
                sLetter = oHit.group(3).lower()
                sCaption = oHit.group(2).replace("&", "").strip().lower()
                dByOwner.setdefault(sOwner, {}).setdefault(sLetter, set()).add(sCaption)
        for sOwner in sorted(dByOwner):
            for sLetter, setCaptions in sorted(dByOwner[sOwner].items()):
                iCount = len(setCaptions)
                if iCount > 1:
                    lsBad.append("%s: the access key %s is claimed %d times in %s"
                                 % (sBase, sLetter, iCount, sOwner))
    for sLine in lsBad: logLine("KEYS: " + sLine)
    if lsBad:
        return finding("keys", "fail", "%s; every one is in the log" %
                       countNoun(len(lsBad), "key problem"))
    return finding("keys", "pass", "0 reserved combinations, 0 access keys claimed twice")


def checkBuildName():
    """THE BUILD SCRIPT IS build.cmd (1.43.55). The folder already names the
    app, so build<App>.cmd said it twice. An old name fails the check and names
    the kit's renameBuild, which renames it and every reference to it."""
    sApp = os.path.basename(sRoot)
    # The kit follows the same rule since 1.43.58: its build script is
    # build.cmd too, and its build removes buildHomerDev.cmd and .py.
    lsOld = [s for s in os.listdir(sRoot) if s.lower() in (("build" + sApp + ".cmd").lower(), ("build" + sApp + ".ps1").lower())]
    if lsOld:
        for sName in lsOld: logLine("BUILD NAME: %s should be %s" % (sName, "build" + os.path.splitext(sName)[1].lower()))
        return finding("buildname", "fail", "%s still named after the app; run the kit's scripts\\renameBuild here" % " and ".join(lsOld))
    if os.path.isfile(os.path.join(sRoot, "build.cmd")):
        return finding("buildname", "pass", "the build script is build.cmd")
    return finding("buildname", "skip", "no build script here")


def checkBuild(bBuild):
    if not bBuild:
        return finding("build", "skip", "not asked for; run with --build to make the build itself evidence")
    # THE APP'S OWN BUILD SCRIPT: build.cmd (1.43.55; before, build<App>.cmd,
    # still run if not yet renamed). Any other build*.cmd -- buildTutorials.cmd,
    # say -- is a tool, not the build; on 25 Sep 2026 this ran
    # buildTutorials.cmd with "nobump" as a script name.
    sOwn = os.path.join(sRoot, "build.cmd")
    if not os.path.isfile(sOwn): sOwn = os.path.join(sRoot, "build" + os.path.basename(sRoot) + ".cmd")
    lsBuild = [sOwn] if os.path.isfile(sOwn) else [s for s in glob.glob(os.path.join(sRoot, "build*.cmd"))
                                                     if os.path.basename(s).lower() not in ("buildtutorials.cmd",)]
    if not lsBuild:
        return finding("build", "skip", "no build script here")
    sScript = os.path.basename(lsBuild[0])
    iCode, sOut = runCommand([], sShell='"%s" nobump' % sScript)
    if iCode == 0:
        return finding("build", "pass", "%s returned 0" % sScript)
    return finding("build", "fail", "%s returned %d; its own log has the compiler output" % (sScript, iCode))


def isWindowed(sPath):
    """True when the executable's PE header names the Windows GUI subsystem."""
    try:
        binHead = open(sPath, "rb").read(4096)
        iPe = int.from_bytes(binHead[0x3C:0x40], "little")
        if binHead[iPe:iPe + 4] != b"PE\0\0": return False
        # The optional header follows the 20-byte file header; Subsystem sits
        # at the same offset, 68, in both the 32-bit and 64-bit forms.
        iSubsystem = int.from_bytes(binHead[iPe + 24 + 68:iPe + 24 + 70], "little")
        return iSubsystem == 2
    except Exception:
        return False


def checkSmoke(bBuild):
    if not bBuild:
        return finding("smoke", "skip", "not asked for; comes with --build")
    # THE PROGRAM LIVES IN exec (the Homer layout, 21 Sep 2026); an app not yet
    # moved keeps it at the top. exec is looked in first, so a stale top-level
    # copy left from before the move is never the one tested.
    lsExe = [s for s in glob.glob(os.path.join(sRoot, "exec", "*.exe")) + glob.glob(os.path.join(sRoot, "*.exe"))
             if not s.lower().endswith("_setup.exe")]
    if not lsExe:
        return finding("smoke", "skip", "no executable here to start")
    sExe = os.path.relpath(lsExe[0], sRoot)
    # A WINDOWED PROGRAM IS NOT STARTED (1.43.4). It has no console to answer
    # --help on; started here it opens its window and waits for a person, and
    # the check sat for its full fifteen-minute timeout before calling that a
    # failure. The PE header says which kind a program is.
    if isWindowed(lsExe[0]):
        return finding("smoke", "skip", "%s is a windowed program, so it is started by hand, not here" % sExe)
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
        "Written by check. Every line below rests on a command that ran,",
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
    logLine("check start pid=%d" % os.getpid())
    logFact("script", os.path.abspath(__file__))
    logFact("python", platform.python_version())
    logFact("windows", logWindows())
    logFact("project", sRoot)
    logFact("arguments", " ".join(sys.argv[1:]))
    logLine("settings build=%s quiet=%s" % (dArguments.build, dArguments.quiet))

    if not dArguments.quiet: sayLine("Checking %s in %s" % (appName(), sRoot))

    checkDocuments()
    checkEncoding()
    checkEmpty()
    checkVersion()
    checkPublish()
    checkLogging()
    checkNaming()
    checkLocal()
    checkFinishPage()
    checkBuildName()
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
    logLine("check end")
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
