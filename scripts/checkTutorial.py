#!/usr/bin/env python3
"""checkTutorial.py -- check the tutorial scripts before their audio is made.

    checkTutorial                 every help\\Tutorial_*.inix in the project
    checkTutorial Tutorial_01     one of them, by stem or file name

WHAT IT CHECKS, script by script, step by step, and reports with the step
number so the line is easy to find:

  - the file parses: every line inside a [section] is Key=Value;
  - every [step] has a Say, and a step with a Key has either a Hear line or a
    Say or Note that names the silence ("nothing", "silent", "silence");
  - no Hear line writes an access key as "Alt+T": the reader says the words,
    "Alt plus T";
  - no script names a screen reader: they are written for everyone's reader;
  - a Hear line about a check box keeps the reader's order: label, "check
    box", state, then the key;
  - a Hear line does not carry a file name the way it is written --
    Introduction_to_Windows.mp3 -- but the way it is said: Introduction to
    Windows dot m p 3;
  - the first script (00) teaches the repeat key, Insert plus Up Arrow;
  - a Key line uses Homer key names: "Control", not "Ctrl", and modifiers
    in alphabetical order.

Zero problems is the answer the build wants; buildTutorials runs this first
and speaks nothing while there are problems. The log is
logs\\<App>-tutorials-check-yyyyMMdd-HHmmss.log in the project.
"""

import datetime, glob, io, os, platform, re, sys

c_lsReaderNames = ["JAWS", "NVDA", "Narrator", "VoiceOver", "Fusion", "ZoomText"]
c_lsSilenceWords = ["nothing", "silent", "silence", "did not say", "says nothing"]
c_lsModifierOrder = ["Alt", "Control", "Shift", "Windows"]

lsProblems = []
oLog = None


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




def say(sText):
    print(sText)
    logLine("CONSOLE: " + sText)
    return True


lsNotices = []

def notice(sText):
    lsNotices.append(sText)
    logLine("NOTICE: " + sText)

def problem(sScript, iStep, sText):
    sWhere = os.path.basename(sScript) + (", step %d" % iStep if iStep else "")
    lsProblems.append("%s: %s" % (sWhere, sText))
    logLine("PROBLEM: %s: %s" % (sWhere, sText))
    return True


def readScript(sPath):
    """Sections as a list of (name, dict of key -> list of values)."""
    lsSections = []
    dNow = None
    sName = ""
    iLine = 0
    with io.open(sPath, encoding="utf-8-sig") as oFile:
        for sRaw in oFile:
            iLine += 1
            sLine = sRaw.strip()
            if not sLine or sLine.startswith((";", "#")): continue
            if sLine.startswith("[") and sLine.endswith("]"):
                sName = sLine[1:-1].strip().lower()
                dNow = {}
                lsSections.append((sName, dNow))
                continue
            if dNow is None: continue  # before the first section, as the tool ignores it
            if "=" not in sLine:
                problem(sPath, 0, "line %d is not Key=Value: %s" % (iLine, sLine[:60]))
                continue
            sKey, sValue = sLine.split("=", 1)
            dNow.setdefault(sKey.strip(), []).append(sValue.strip())
    return lsSections


def firstOf(dSection, sKey):
    lsValues = dSection.get(sKey, [])
    return lsValues[0] if lsValues else ""


def checkKeyName(sScript, iStep, sKey):
    if not sKey: return True
    if re.search(r"\bCtrl\b", sKey): problem(sScript, iStep, "Key writes Ctrl; Homer spells out Control: %s" % sKey)
    lsParts = [s.strip() for s in sKey.split("+")]
    lsMods = [s for s in lsParts[:-1] if s in c_lsModifierOrder]
    if lsMods != sorted(lsMods, key=c_lsModifierOrder.index):
        problem(sScript, iStep, "Key modifiers are not in alphabetical order (Alt, Control, Shift, Windows): %s" % sKey)
    return True


def checkHear(sScript, iStep, sHear):
    if re.search(r"\bAlt\+[A-Za-z0-9]", sHear):
        problem(sScript, iStep, "Hear writes an access key with a plus sign; the reader says the words: %s" % sHear[:70])
    if re.search(r"[A-Za-z0-9]_[A-Za-z0-9]", sHear) or re.search(r"\.(mp3|mp4|pdf|md|htm|txt|docx|wav|mkv)\b", sHear, re.I):
        problem(sScript, iStep, "Hear carries a file name as written, not as said (spaces, then dot m p 3): %s" % sHear[:70])
    if "check box" in sHear.lower():
        # label, "check box", state, key -- so "check box" precedes the state and the key
        sLower = sHear.lower()
        iBox = sLower.find("check box")
        sAfter = sLower[iBox:]
        if not re.search(r"check box,? (not checked|checked|unchecked)", sAfter):
            problem(sScript, iStep, "Hear for a check box should read label, check box, state, key: %s" % sHear[:70])
    return True


def checkOne(sScript, bFirst):
    lsSections = readScript(sScript)
    lsSteps = [d for sName, d in lsSections if sName == "step"]
    dAbout = next((d for sName, d in lsSections if sName == "about"), {})
    if not dAbout: problem(sScript, 0, "no [about] section")
    for sKey in ("Title", "Intro", "Setup", "Homework"):
        if not firstOf(dAbout, sKey): problem(sScript, 0, "[about] lacks %s" % sKey)
    if not lsSteps: problem(sScript, 0, "no [step] sections")
    sWhole = io.open(sScript, encoding="utf-8-sig").read()
    for sName in c_lsReaderNames:
        if re.search(r"\b%s\b" % re.escape(sName), sWhole):
            problem(sScript, 0, "names a screen reader (%s); write \"the screen reader\"" % sName)
    for iAt, dStep in enumerate(lsSteps, 1):
        sSay = firstOf(dStep, "Say")
        sKey = firstOf(dStep, "Key")
        lsHear = [s for s in dStep.get("Hear", []) if s]
        sNote = firstOf(dStep, "Note")
        if not sSay: problem(sScript, iAt, "no Say line")
        checkKeyName(sScript, iAt, sKey)
        if sKey and not lsHear:
            sNext = firstOf(lsSteps[iAt], "Say") if iAt < len(lsSteps) else ""
            sText = (sNote + " " + sNext).lower()
            if not any(w in sText for w in c_lsSilenceWords):
                problem(sScript, iAt, "Key %s has no Hear line and nothing names the silence" % sKey)
        for sHear in lsHear: checkHear(sScript, iAt, sHear)
    if bFirst and not re.search(r"Insert (plus )?Up Arrow", sWhole):
        problem(sScript, 0, "the first script does not teach the repeat key, Insert plus Up Arrow")
    # The orientation key goes with the repeat key: the trainers teach both in
    # their first module, and a listener who can repeat a line but cannot ask
    # "where am I" is half equipped.
    if bFirst and not re.search(r"Insert (plus )?Tab", sWhole):
        # A notice, as the pattern is: a set written before the orientation key
        # joined the first walk (5 October 2026) is still spoken, and told.
        notice(os.path.basename(sScript) + ": the first walk should teach the orientation key, Insert plus Tab")
    logLine("%s: %d steps, %d Hear lines" % (os.path.basename(sScript), len(lsSteps), sum(len([h for h in d.get("Hear", []) if h]) for d in lsSteps)))
    return True


def projectRoot(sScriptDir):
    if os.path.basename(sScriptDir).lower() in ("scripts", "tools", "exec"):
        return os.path.dirname(sScriptDir)
    return sScriptDir



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
    global oLog
    sRoot = projectRoot(os.path.dirname(os.path.abspath(__file__)))
    sHelp = os.path.join(sRoot, "help")
    if not os.path.isdir(sHelp): sHelp = sRoot
    sLogDir = os.path.join(sRoot, "logs")
    os.makedirs(sLogDir, exist_ok=True)
    sLogPath = os.path.join(sLogDir, "%s-tutorials-check-%s.log" % (os.path.basename(sRoot), datetime.datetime.now().strftime("%Y%m%d-%H%M%S")))
    oLog = io.open(sLogPath, "w", encoding="utf-8")
    logLine("checkTutorial start pid=%d" % os.getpid())
    sKind, sWhy = loadKind()(sRoot)
    logLine("setting kind=%s reason=%s" % (sKind, sWhy))
    if sKind in ("collection", "page"):
        print("A %s has no program to walk through, so there is no tutorial to check." % sKind)
        return 0
    logFact("script", os.path.abspath(__file__))
    logFact("python", platform.python_version())
    logFact("windows", logWindows())
    logFact("project", sRoot)
    logFact("arguments", " ".join(sys.argv[1:]))
    if len(sys.argv) > 1:
        sStem = os.path.splitext(os.path.basename(sys.argv[1]))[0]
        lsScripts = [os.path.join(sHelp, sStem + ".inix")]
    else:
        lsScripts = sorted(glob.glob(os.path.join(sHelp, "Tutorial_*.inix")))
    if not lsScripts or not os.path.isfile(lsScripts[0]):
        say("0 tutorial scripts found in %s." % sHelp)
        return 0
    for iAt, sScript in enumerate(lsScripts):
        checkOne(sScript, bFirst=(iAt == 0 and len(sys.argv) == 1))
    # FIVE MINUTES IS THE CEILING (5 October 2026): at the voices' pace about
    # twenty-five steps. Twenty-eight is the line here, so a walk a little over
    # passes and one plainly over is sent back to be cut or split.
    for sScript in lsScripts:
        iSteps = readText(sScript).count("[step]") if "readText" in globals() else open(sScript, "rb").read().decode("utf-8-sig").count("[step]")
        if iSteps > 28:
            problem(os.path.basename(sScript), 0, "%d steps is more than five minutes; cut what an earlier walk taught, or split it" % iSteps)
    # THE TWELVE-WALK PATTERN (5 October 2026) is reported as NOTICES, not
    # problems: a program whose set is not yet the pattern still has its clean
    # walks spoken and still releases, and hears on every build what the set
    # lacks. A problem is something wrong in a walk; an incomplete set is work
    # not yet done, and the tool should not silence a program for that.
    if len(sys.argv) == 1 and sKind != "kit":
        c_dFixed = {"00": "Overview_and_Table_of_Contents", "01": "Install_and_Launch", "02": "User_Interface_Concepts",
                    "03": "Key_Patterns", "09": "Glossary", "10": "Conclusion", "11": "More_Information"}
        dHave = {}
        for sScript in lsScripts:
            m = re.match(r"Tutorial_(\d\d)_(.+)\.inix$", os.path.basename(sScript))
            if m: dHave[m.group(1)] = m.group(2)
        for sNum, sName in sorted(c_dFixed.items()):
            if dHave.get(sNum) != sName:
                notice("the pattern wants Tutorial_%s_%s.inix%s" % (sNum, sName, (", not " + dHave[sNum]) if sNum in dHave else ""))
        # Tasks are 04 to 08: at least one, at most five, filled from 04 up with
        # no gap, so the numbers are the order and the order is the numbers.
        lsTasks = [s for s in sorted(dHave) if s in ("04", "05", "06", "07", "08")]
        if not lsTasks: notice("the pattern wants at least one task walk, numbered from 04")
        for iAt, sNum in enumerate(lsTasks):
            if int(sNum) != 4 + iAt: notice("the task walks must run from 04 without a gap; found %s" % ", ".join(lsTasks)); break
        for sNum in sorted(dHave):
            if sNum not in c_dFixed and sNum not in ("04", "05", "06", "07", "08"):
                notice("Tutorial_%s is outside the pattern of 00 to 11" % sNum)
    for sText in lsProblems: say("  " + sText)
    if lsNotices:
        say("  The set is not yet the Homer pattern of twelve walks:")
        for sText in lsNotices: say("    " + sText)
    say("%d script%s checked, %d problem%s." % (len(lsScripts), "" if len(lsScripts) == 1 else "s",
        len(lsProblems), "" if len(lsProblems) == 1 else "s"))
    logLine("checkTutorial end")
    oLog.close()
    return 0 if not lsProblems else 1


if __name__ == "__main__":
    sys.exit(main())
