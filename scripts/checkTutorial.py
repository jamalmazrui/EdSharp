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
    if oLog is not None:
        oLog.write(sText + "\n")
        oLog.flush()
    return True


def say(sText):
    print(sText)
    logLine("CONSOLE: " + sText)
    return True


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
    logLine("%s: %d steps, %d Hear lines" % (os.path.basename(sScript), len(lsSteps), sum(len([h for h in d.get("Hear", []) if h]) for d in lsSteps)))
    return True


def projectRoot(sScriptDir):
    if os.path.basename(sScriptDir).lower() in ("scripts", "tools", "exec"):
        return os.path.dirname(sScriptDir)
    return sScriptDir


def main():
    global oLog
    sRoot = projectRoot(os.path.dirname(os.path.abspath(__file__)))
    sHelp = os.path.join(sRoot, "help")
    if not os.path.isdir(sHelp): sHelp = sRoot
    sLogDir = os.path.join(sRoot, "logs")
    os.makedirs(sLogDir, exist_ok=True)
    sLogPath = os.path.join(sLogDir, "%s-tutorials-check-%s.log" % (os.path.basename(sRoot), datetime.datetime.now().strftime("%Y%m%d-%H%M%S")))
    oLog = io.open(sLogPath, "w", encoding="utf-8")
    logLine("checkTutorial started %s" % datetime.datetime.now().isoformat(" ", "seconds"))
    logLine("Script: %s" % os.path.abspath(__file__))
    logLine("Python: %s" % sys.version.replace("\n", " "))
    logLine("Platform: %s" % platform.platform())
    logLine("Project: %s" % sRoot)
    logLine("Command line: %s" % " ".join(sys.argv))
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
    for sText in lsProblems: say("  " + sText)
    say("%d script%s checked, %d problem%s." % (len(lsScripts), "" if len(lsScripts) == 1 else "s",
        len(lsProblems), "" if len(lsProblems) == 1 else "s"))
    logLine("Finished %s" % datetime.datetime.now().isoformat(" ", "seconds"))
    oLog.close()
    return 0 if not lsProblems else 1


if __name__ == "__main__":
    sys.exit(main())
