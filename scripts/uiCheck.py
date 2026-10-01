#!/usr/bin/env python3
r"""uiCheck.py -- run a program, press the keys, and check what actually appeared.

WHY THIS EXISTS

Everything else in this kit reads files or runs a compiler. Neither answers the
only question that matters: does the program behave, once it is running, the way
its documentation says?

Up to now the answer was "somebody has to open it and press the keys". That is
the wrong answer. A check a person has to remember is a check that stops
happening in the week it matters, and the point of this whole kit is that
evidence is gathered rather than recalled.

So this drives a real program through Windows UI Automation -- the same
interface a screen reader uses to find out what is on the screen. It starts the
program, sends keystrokes, and asks the accessibility tree what is there
afterwards. If a control has no accessible name, this finds nothing, which is
exactly the failure worth catching.

WHAT IT CANNOT DO, and this is the honest boundary: it does not hear speech. UI
Automation tells you a control exists and what it is called; it does not tell you
what JAWS said. Capturing that is possible -- NVDA can log speech at debug level
-- and is not done here. Every report says so.

THE TEST FILE. Tests live beside the program in uiTest.inix, so an app carries
its own:

    [test]
    Name = the basket takes a fruit and says so
    Run = FruitBasketCs.exe

    [step]
    Keys = apple{ENTER}
    Wants = apple
    Note = the fruit is in the list after Enter

    [step]
    Keys = %r
    Title = Fruit basket report
    Escape = yes

  Run     the program to start; one per [test]
  Keys    what to send, in pywinauto syntax: {ENTER} {ESC} {F4} %x is Alt+x,
          ^x is Control+x, +x is Shift+x
  Wants   a string that must appear somewhere in the window's accessibility tree
  Title   a window with this title must exist (a * at the end matches a prefix)
  Escape  yes to press Escape afterwards, for a step that opened a dialog
  Wait    seconds to wait before looking, when something is slow
  Note    for the report; never sent

USAGE

    uiCheck                         every uiTest.inix beside this script
    uiCheck --path C:\JobDo         an app's own tests
    uiCheck --keep                  leave the program running at the end

Writes evidence-ui-<yyyymmdd-hhmmss>.md and uiCheck.log beside this script.
Exit code 0 when nothing failed, 1 when something did.
"""

import argparse
import datetime
import glob
import os
import platform
import subprocess
import sys
import time
import traceback

c_fStartSeconds = 8.0        # how long a program may take to show its window
c_fStepSeconds = 0.6         # settle time after a keystroke

sScriptDir = os.path.dirname(os.path.abspath(__file__))
sLogPath = os.path.join(sScriptDir, "uiCheck.log")
sRoot = sScriptDir
oLog = None
lsFindings = []              # (sTest, sStep, sVerdict, sEvidence)


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


def finding(sTest, sStep, sVerdict, sEvidence):
    lsFindings.append((sTest, sStep, sVerdict, sEvidence))
    logLine("%-28s %-34s %-5s %s" % (sTest[:28], sStep[:34], sVerdict.upper(), sEvidence))
    return True


# --- the driver -------------------------------------------------------------

def loadDriver():
    """pywinauto, installed if it is not here. The build never asks a person."""
    try:
        from pywinauto.application import Application
        return Application
    except ImportError:
        pass
    sayLine("Installing pywinauto, which drives the window...")
    iCode = subprocess.call([sys.executable, "-m", "pip", "install", "--quiet",
                             "pywinauto"])
    logLine("pip install pywinauto exit %d" % iCode)
    try:
        from pywinauto.application import Application
        return Application
    except ImportError:
        return None


def treeText(window):
    """Every name and value in the window, as the accessibility tree reports it.

    This is the same tree a screen reader reads. A control missing from it is a
    control a screen reader cannot announce, so an empty answer here is a real
    finding rather than a limitation of the check.
    """
    lsText = []
    try:
        lsText.append(window.window_text())
    except Exception:
        pass
    try:
        for control in window.descendants():
            try:
                sText = control.window_text()
                if sText: lsText.append(sText)
            except Exception:
                continue
            try:
                # A list reports its items separately from its own name.
                for sItem in control.texts():
                    if sItem and sItem not in lsText: lsText.append(sItem)
            except Exception:
                continue
    except Exception as oError:
        logLine("TREE FAILED: %s" % oError)
    return lsText


def titleExists(Application, sTitle):
    """Is a window with this title open? A trailing * matches a prefix."""
    from pywinauto import Desktop
    try:
        for window in Desktop(backend="uia").windows():
            try:
                sText = window.window_text()
            except Exception:
                continue
            if sTitle.endswith("*"):
                if sText.startswith(sTitle[:-1]): return True
            elif sText == sTitle:
                return True
    except Exception as oError:
        logLine("DESKTOP SCAN FAILED: %s" % oError)
    return False


# --- reading the tests ------------------------------------------------------

def readTests(sPath):
    """The .inix as a list of tests, each with its steps."""
    lTests = []
    dNow = None
    for sLine in open(sPath, "rb").read().decode("utf-8-sig", "replace").splitlines():
        sLine = sLine.strip()
        if not sLine or sLine.startswith(";") or sLine.startswith("#"): continue
        if sLine.startswith("[") and sLine.endswith("]"):
            sName = sLine[1:-1].strip().lower()
            if sName == "test":
                dNow = {"_steps": []}
                lTests.append(dNow)
            elif sName == "step" and lTests:
                dNow = {}
                lTests[-1]["_steps"].append(dNow)
            continue
        if dNow is None or "=" not in sLine: continue
        sField, sValue = sLine.split("=", 1)
        dNow[sField.strip().lower()] = sValue.strip()
    return lTests


# --- running one test -------------------------------------------------------

def runTest(Application, dTest, bKeep):
    sName = dTest.get("name", "unnamed test")
    sExe = dTest.get("run", "")
    sExePath = os.path.join(sRoot, sExe)
    if not sExe:
        return finding(sName, "start", "skip", "the test names no program to run")
    if not os.path.isfile(sExePath):
        return finding(sName, "start", "skip",
                       "%s is not here; build it first" % sExe)

    sayLine("  %s" % sName)
    app = None
    try:
        app = Application(backend="uia").start('"%s"' % sExePath,
                                               work_dir=sRoot, timeout=c_fStartSeconds)
        time.sleep(1.2)
        window = app.top_window()
        window.wait("visible ready", timeout=c_fStartSeconds)
        finding(sName, "start", "pass", "%s opened a window titled %r" %
                (sExe, window.window_text()))
    except Exception as oError:
        logLine("START FAILED: %s" % traceback.format_exc())
        finding(sName, "start", "fail", "%s did not open a window: %s" % (sExe, oError))
        try:
            if app is not None: app.kill()
        except Exception:
            pass
        return False

    iStep = 0
    for dStep in dTest["_steps"]:
        iStep += 1
        sStep = dStep.get("note", "step %d" % iStep)
        sKeys = dStep.get("keys", "")
        try:
            if sKeys:
                window.set_focus()
                window.type_keys(sKeys, with_spaces=True, set_foreground=True)
            time.sleep(float(dStep.get("wait", c_fStepSeconds)))

            sTitle = dStep.get("title", "")
            if sTitle:
                if titleExists(Application, sTitle):
                    finding(sName, sStep, "pass", "a window titled %r appeared" % sTitle)
                else:
                    finding(sName, sStep, "fail", "no window titled %r" % sTitle)

            sWants = dStep.get("wants", "")
            if sWants:
                lsText = treeText(app.top_window())
                logLine("TREE: " + " | ".join(lsText[:60]))
                if any(sWants.lower() in s.lower() for s in lsText):
                    finding(sName, sStep, "pass",
                            "%r is in the window, as %s reports it" %
                            (sWants, "UI Automation"))
                else:
                    finding(sName, sStep, "fail",
                            "%r is not in the window; %s found there" %
                            (sWants, countNoun(len(lsText), "name")))

            if not sTitle and not sWants:
                finding(sName, sStep, "pass", "keys sent: %s" % sKeys)

            if dStep.get("escape", "").lower() in ("yes", "y", "1", "true"):
                window.type_keys("{ESC}", set_foreground=True)
                time.sleep(c_fStepSeconds)
        except Exception as oError:
            logLine("STEP FAILED: %s" % traceback.format_exc())
            finding(sName, sStep, "fail", str(oError))

    try:
        if not bKeep: app.kill()
    except Exception:
        pass
    return True


# --- the report -------------------------------------------------------------

def writeReport():
    sStamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    sPath = os.path.join(sScriptDir, "evidence-ui-%s.md" % sStamp)
    lsPassed = [t for t in lsFindings if t[2] == "pass"]
    lsFailed = [t for t in lsFindings if t[2] == "fail"]
    lsSkipped = [t for t in lsFindings if t[2] == "skip"]

    lsLines = [
        "---",
        'title: "Evidence report: what the programs did"',
        'date: "%s"' % datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "---",
        "",
        "# Evidence report: what the programs did",
        "",
        "Written by uiCheck. Each line below rests on a program that was started,",
        "keys that were sent to it, and the Windows UI Automation tree read back",
        "afterwards -- the same interface a screen reader uses to find out what is",
        "on the screen. uiCheck.log holds every keystroke and every name found.",
        "",
        "## What was verified",
        "",
    ]
    lsLines += ["- **%s** -- %s: %s" % (t, s, e) for t, s, v, e in lsPassed] or ["- nothing"]
    lsLines += ["", "## What failed", ""]
    lsLines += ["- **%s** -- %s: %s" % (t, s, e) for t, s, v, e in lsFailed] or ["- nothing"]
    lsLines += ["", "## What was not checked", ""]
    lsLines += ["- **%s** -- %s: %s" % (t, s, e) for t, s, v, e in lsSkipped] or ["- nothing"]
    lsLines += [
        "",
        "## What remains uncertain",
        "",
        "- **What a screen reader actually says.** UI Automation reports that a",
        "  control exists and what it is called. It does not report speech. NVDA",
        "  can log what it speaks, and capturing that is the next step; it is not",
        "  done here.",
        "- **Whether the order is sensible.** The tree says a control is present,",
        "  not that tabbing through the window is pleasant.",
        "- **Whether the program did the right thing.** These checks ask whether",
        "  it behaves as documented, not whether the documentation was right.",
        "- **Anything not in a test.** A program is only as checked as its",
        "  uiTest.inix says.",
        "",
    ]
    open(sPath, "wb").write(("\ufeff" + "\r\n".join(lsLines)).encode("utf-8"))
    return sPath


def main():
    global oLog, sRoot
    oParser = argparse.ArgumentParser(description="Run a program and check what appeared.")
    oParser.add_argument("--keep", action="store_true", help="leave the program running")
    oParser.add_argument("--path", default="", help="the folder holding uiTest.inix")
    dArguments = oParser.parse_args()
    if dArguments.path: sRoot = os.path.abspath(dArguments.path)

    oLog = open(sLogPath, "w", encoding="utf-8")
    logLine("uiCheck start pid=%d" % os.getpid())
    logFact("script", os.path.abspath(__file__))
    logFact("python", platform.python_version())
    logFact("windows", logWindows())
    logLine("Folder: %s" % sRoot)
    logFact("arguments", " ".join(sys.argv[1:]))

    if not sys.platform.startswith("win"):
        sayLine("uiCheck drives Windows programs, so it does nothing on %s." % sys.platform)
        logLine("NOT WINDOWS; nothing run")
        return 0

    Application = loadDriver()
    if Application is None:
        sayLine("pywinauto could not be installed, so no program was driven.")
        logLine("NO DRIVER")
        return 1

    lsFiles = sorted(glob.glob(os.path.join(sRoot, "uiTest*.inix")))
    if not lsFiles:
        sayLine("There is no uiTest.inix in %s, so there is nothing to check." % sRoot)
        return 1

    for sFile in lsFiles:
        sayLine("Tests from %s" % os.path.basename(sFile))
        for dTest in readTests(sFile):
            runTest(Application, dTest, dArguments.keep)

    sReport = writeReport()
    iPassed = len([t for t in lsFindings if t[2] == "pass"])
    iFailed = len([t for t in lsFindings if t[2] == "fail"])
    iSkipped = len([t for t in lsFindings if t[2] == "skip"])
    sayLine()
    sayLine("%s passed, %s failed, %s not checked." %
            (countNoun(iPassed, "check"), countNoun(iFailed, "check"),
             countNoun(iSkipped, "check")))
    for sTest, sStep, sVerdict, sEvidence in lsFindings:
        if sVerdict == "fail": sayLine("  failed: %s -- %s" % (sTest, sEvidence))
    sayLine("The report is %s." % os.path.basename(sReport))
    logLine("uiCheck end")
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
