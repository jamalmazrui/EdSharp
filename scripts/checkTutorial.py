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
# SHOW, THEN TELL (1.58.0). In a Homer walk the host names a concept and the
# screen reader then shows it: the control that takes focus, read as name,
# role, value and state, with the hint after; or the window that comes
# forward, read by its title. The author heard the user-interface walk run on
# with the host alone and found it far less instructive than an exchange. So in
# that walk, 1_User_Interface (it was 02 and 03 before the pattern of ten), no
# more than c_iMaxHostRun steps in a row may pass without the reader, and at
# least c_nMinReaderShare of the steps must carry the reader. Other walks get a
# notice for a longer run. 0 is prose and a table of contents by design, so it
# is not measured.
c_dConceptWalks = {"1": "User Interface"}
c_iLongSay = 60
c_iMaxHostRun = 2
c_iNoticeHostRun = 3
c_lsUnmeasuredWalks = ["0"]
c_nMinReaderShare = 0.5

lsProblems = []
oLog = None


# HOW LONG A WALK RUNS (1.64.4), fitted to twenty walks whose audio the build
# measured on 8 October 2026 (DbDo's and HomerView's): the host speaks at about
# 225 words a minute, the reader voice at about 575, and each step adds about
# 3.6 seconds of pauses and turn-taking. Average error 1%, worst 13%; the kit's
# own walk 8, not used in the fit, measured 182 seconds and is predicted 188.
# Words alone at 186 a minute had run 21% short on average, and up to 33%, on
# walks of many short two-voice steps.
c_nHostSecondsPerWord = 0.267
c_nReaderSecondsPerWord = 0.104
c_nSecondsPerStep = 3.63

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
    # THE READER SPEAKS CONCRETE WORDS (10 October 2026). The host may speak in general terms; the reader says what a
    # screen reader would say at that moment -- a real control's name, role, value and state, a real title, the words
    # typed -- never a description of it, such as "(the current control, with its state and position)". No walk uses
    # one screen reader's own term for a voice ("message voice"), and none ships the template's CHANGE ME lines.
    if re.search(r"(?i)\bmessage voice\b", sWhole):
        problem(sScript, 0, "says \"message voice\", one screen reader's term; describe what is heard instead")
    for iLine, sLine in enumerate(sWhole.replace("\r", "").split("\n"), 1):
        if re.match(r"Hear=\s*\(", sLine):
            problem(sScript, 0, "line %d: a reader line describes speech instead of speaking it: %s; write the words the reader says" % (iLine, sLine[5:80]))
        if sLine.startswith("Hear=") and "CHANGE ME" in sLine:
            problem(sScript, 0, "line %d: a reader line still holds the template's CHANGE ME; write what the reader says" % iLine)
    # THE OVERVIEW'S GENTLE REMINDER (10 October 2026): a listener already knows the reader's own keys, so the
    # Overview reminds, briefly, of three worth pressing when lost -- Insert plus Up Arrow, Insert plus Tab and Insert
    # plus T, each answered by a concrete reader line -- and that Caps Lock serves in place of Insert where set up.
    if re.match(r"Tutorial_0_", os.path.basename(sScript)):
        for sKeyWant, sWhat in ((r"Insert\+Up ?Arrow", "Insert plus Up Arrow"), (r"Insert\+Tab", "Insert plus Tab"), (r"Insert\+T\b", "Insert plus T")):
            mStep = re.search(r"(?ms)^Key=%s\s*$\n(?:Hear=(.*?)$)?" % sKeyWant, sWhole.replace("\r", ""))
            if not mStep: problem(sScript, 0, "the Overview does not remind of %s, one of the three reader keys" % sWhat)
            elif not (mStep.group(1) or "").strip(): problem(sScript, 0, "the Overview's %s step has no reader line" % sWhat)
        if not re.search(r"(?i)caps ?lock", sWhole):
            problem(sScript, 0, "the Overview does not say that Caps Lock can serve in place of Insert")
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
    oNumber = re.match(r"Tutorial_(\d)_", os.path.basename(sScript))
    sNumber = oNumber.group(1) if oNumber else ""
    if sNumber and sNumber not in c_lsUnmeasuredWalks and lsSteps:
        iRun, iLongest, iAtLongest, iWithReader = 0, 0, 0, 0
        for iAt, dStep in enumerate(lsSteps, 1):
            if [s for s in dStep.get("Hear", []) if s]:
                iWithReader += 1
                iRun = 0
                continue
            iRun += 1
            if iRun > iLongest: iLongest, iAtLongest = iRun, iAt
        nShare = float(iWithReader) / len(lsSteps)
        logLine("%s: the reader speaks in %d of %d steps; the longest stretch of the host alone is %d step%s" % (os.path.basename(sScript), iWithReader, len(lsSteps), iLongest, "" if iLongest == 1 else "s"))
        if sNumber in c_dConceptWalks:
            if iLongest > c_iMaxHostRun: problem(sScript, iAtLongest, "%d steps in a row with the host alone; in the %s walk each concept is shown by the reader, its name, role, value, state and hint, right after it is named" % (iLongest, c_dConceptWalks[sNumber]))
            if nShare < c_nMinReaderShare: problem(sScript, 0, "the reader speaks in only %d of %d steps; in the %s walk at least half the steps show a concept in the reader's voice" % (iWithReader, len(lsSteps), c_dConceptWalks[sNumber]))
        elif iLongest > c_iNoticeHostRun:
            notice("%s: %d steps in a row with the host alone, ending at step %d; let the reader show what the host has named" % (os.path.basename(sScript), iLongest, iAtLongest))
    for iAt, dStep in enumerate(lsSteps, 1):
        iWords = len(firstOf(dStep, "Say").split())
        if iWords > c_iLongSay: notice("%s step %d: a Say line of %d words; say less, then let the reader show it" % (os.path.basename(sScript), iAt, iWords))
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
    # BETWEEN THREE AND FIVE MINUTES for parts 01 to 10 (6 October 2026): under
    # three is usually too thin to be worth a listener's start; 00 and 11 may
    # be short. A glossary's steps are two short lines each, so its count is
    # not a measure of its length and it is exempt from the ceiling; the tool
    # measures the audio itself and says what runs under or over.
    for sScript in lsScripts:
        sBase = os.path.basename(sScript)
        iSteps = readText(sScript).count("[step]") if "readText" in globals() else open(sScript, "rb").read().decode("utf-8-sig").count("[step]")
        mNum = re.match(r"Tutorial_(\d)_", sBase)
        sNum = mNum.group(1) if mNum else ""
        # Step counts are a guess at length; the tool's measurement of the
        # audio is the fact. A high count is a notice, so a walk of many short
        # two-voice exchanges is not refused for being brisk. The conclusion,
        # 9, holds the glossary, whose steps are two short lines each, so its
        # count is no measure of its length.
        # WORDS, NOT STEPS (kit 1.63.1): DbDo's walks had 12 to 14 steps, passed
        # the step rule, and ran 1:26 to 1:55, because their lines were short.
        # The kit's own walks are spoken at about 186 words a minute (155 to 204,
        # measured 8 October 2026), so the spoken words -- every Say and Hear
        # line -- give the estimate: about 560 for three minutes, 930 for five.
        sWalk = open(sScript, "rb").read().decode("utf-8-sig")
        iHostWords = sum(len(sLine.split()) for sLine in re.findall(r"(?mi)^\s*Say\s*=\s*(.*)$", sWalk))
        iReaderWords = sum(len(sLine.split()) for sLine in re.findall(r"(?mi)^\s*Hear\s*=\s*(.*)$", sWalk))
        iWalkSteps = len(re.findall(r"(?mi)^\s*\[step\]\s*$", sWalk))
        iWords = iHostWords + iReaderWords
        dMinutes = (iHostWords * c_nHostSecondsPerWord + iReaderWords * c_nReaderSecondsPerWord + iWalkSteps * c_nSecondsPerStep) / 60
        # A notice outside 2.75 to 5.25 predicted minutes: the prediction errs by up to about 13%, so a walk
        # inside that band may land either side of three or five, and the tool's measurement after speaking settles it.
        if sNum and sNum not in ("0", "9") and dMinutes < 2.75:
            notice(sBase + ": predicted %.1f minutes (%d host words, %d reader words, %d steps), likely under three; the guideline wants three to five for parts 1 to 8" % (dMinutes, iHostWords, iReaderWords, iWalkSteps))
        elif sNum and sNum != "9" and dMinutes > 5.25:
            notice(sBase + ": predicted %.1f minutes; the tool will measure it -- cut what an earlier walk taught, or split it, if it is over five" % dMinutes)
    # THE PATTERN OF TEN (7 October 2026, replacing the twelve of 5 October):
    # one digit sorts the set. 0_Overview, 1_User_Interface (the concepts and
    # the key patterns together), seven task walks numbered 2 to 8 (usually
    # 2_Install_and_Launch first), and 9_Conclusion (the conclusion and
    # summary, the glossary, and more information, help built in among it).
    # Every set is the full ten. Reported as NOTICES, not problems: a program whose set is not
    # yet the pattern still has its clean walks spoken and still releases, and
    # hears on every build what the set lacks -- including a set still named
    # with two digits, which is told what each old walk becomes.
    if len(sys.argv) == 1:
        c_dFixed = {"0": "Overview", "1": "User_Interface", "9": "Conclusion"}
        c_dOld = {"00": "0_Overview", "01": "2_Install_and_Launch, the first task", "02": "1_User_Interface, with 03", "03": "1_User_Interface, with 02",
                  "09": "9_Conclusion, with 10 and 11", "10": "9_Conclusion, with 09 and 11", "11": "9_Conclusion, with 09 and 10"}
        dHave, lsOld = {}, []
        for sScript in lsScripts:
            sName = os.path.basename(sScript)
            m = re.match(r"Tutorial_(\d)_(.+)\.inix$", sName)
            if m: dHave[m.group(1)] = m.group(2)
            m2 = re.match(r"Tutorial_(\d\d)_(.+)\.inix$", sName)
            if m2: lsOld.append((m2.group(1), sName))
        for sNum, sName in lsOld:
            notice("%s uses the old two-digit numbering; it becomes %s" % (sName, c_dOld.get(sNum, ("task walk %d, after 2_Install_and_Launch" % (int(sNum) - 1)) if 4 <= int(sNum) <= 8 else "part of the pattern of ten")))
        for sNum, sName in sorted(c_dFixed.items()):
            if dHave.get(sNum) != sName:
                notice("the pattern wants Tutorial_%s_%s.inix%s" % (sNum, sName, (", not " + dHave[sNum]) if sNum in dHave else ""))
        # Tasks are 2 to 8: seven of them, every number used, so every set is
        # the full ten and walk 9 is always the conclusion.
        for sNum in ("2", "3", "4", "5", "6", "7", "8"):
            if sNum not in dHave: notice("the pattern wants task walk %s; the seven tasks are numbered 2 to 8" % sNum)
    for sText in lsProblems: say("  " + sText)
    if lsNotices:
        say("  Notices -- work to do, never a reason to stop the build:")
        for sText in lsNotices: say("    " + sText)
    say("%d script%s checked, %d problem%s." % (len(lsScripts), "" if len(lsScripts) == 1 else "s",
        len(lsProblems), "" if len(lsProblems) == 1 else "s"))
    logLine("checkTutorial end")
    oLog.close()
    return 0 if not lsProblems else 1


if __name__ == "__main__":
    sys.exit(main())
