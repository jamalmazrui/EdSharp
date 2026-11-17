r"""testSay.py -- every Say command of a Homer app, pressed once and twice, then described, with what the reader said.

WHY THIS EXISTS (9 November 2026 in DbDo; in the kit for every app since 17 November 2026)

A Homer app's Say commands each speak one fact, and a second press may spell a short value or show a longer one in a
window. Jamal hears them through JAWS, and what JAWS says can differ from what the app meant to say. This script
writes a macro script of every Say command in the app's key file and plays it with the kit's runMacro, which presses
each key in the real program, records the computer's sound while the reader speaks, writes it down with a local
speech recognizer, and sets it beside what the app's own log says it said:

    1. Each Say command run once, then twice in quick succession, then Escape, which closes any window a second press
       opened.
    2. Control+F1 turns on the Key Describer, each Say command's key is pressed once, so it is described rather than
       run, and Control+F1 turns the Key Describer off.

DbDo's testSay pressed Shift+A to Shift+Z; this one takes the Say commands from the key file, so FileDir's Shift+8
and Alt+Y are heard too, and a command moved to another key is followed.

BEFORE RUNNING. Open the app on something worth hearing about -- DbDo on a table with records, FileDir on a folder
with files -- with JAWS running as usual; an app that is not running is started from its own build. Then keep your
hands off the keyboard. To stop, press Control+C in the console, or switch to any other window.

NO ARGUMENT IS NEEDED (17 November 2026). A Homer script with no arguments does the most likely and most complete
thing. Run from HomerDev, which has no Say commands of its own, testSay used to stop with "No Say command with a key
matched"; now it tests every Homer app it finds beside the kit, one after another, and says the time it expects first.

USAGE (from an app's folder or its scripts folder, or from HomerDev's)

    testSay                     every Say command, about a minute for every three; from HomerDev, every app's in turn
    testSay --apps DbDo,FileDir from HomerDev, only those apps
    testSay --only F,O,8        only the Say commands on those keys (a single letter or digit is a key)
    testSay --only position     only the Say commands with that word in their names
    testSay --letters ABC       only the Say commands on Shift+A, Shift+B and Shift+C (DbDo's old option)
    testSay --no-describer      leave out the Key Describer pass
    testSay --wait 8            listen up to eight seconds after each key (default 6)
    testSay --check             write the macro and check every line is understood, without pressing anything

WHAT IT WRITES

    logs\<App>-testSay-<stamp>.log      this script's own log: environment, settings, every command and exit code
    logs\testSay-<stamp>\               the folder to upload:
        testSay.inix                    the macro script it played, plain lines anyone can read and play again
        <App>-recordWalk-<stamp>.md     the report: key by key, heard and logged speech side by side
        <App>-recordWalk-<stamp>.log    recordWalk's log, and runMacro's
        recordings\                     one recording per key
        appLogs\                        the app's own session logs from the run, with every say: line
        testSay.md                      what is in the folder

Exit code: runMacro's (1 is usual, since the macro expects nothing in particular), or 2 when it could not start.
"""

import argparse, datetime, glob, os, platform, re, shutil, subprocess, sys, time, traceback

sStamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
sScriptDir = os.path.dirname(os.path.abspath(__file__))


def projectFolder():
    """The app's folder: the working folder when it holds configs, else the folder above this script's."""
    for sDir in (os.getcwd(), os.path.dirname(os.getcwd()), os.path.dirname(sScriptDir)):
        if os.path.isdir(os.path.join(sDir, "configs")) and os.path.isdir(os.path.join(sDir, "scripts")): return sDir
    return os.path.dirname(sScriptDir)


sProject = projectFolder()
sApp = os.path.basename(os.path.normpath(sProject))
sLogDir = os.path.join(sProject, "logs")
sLogPath = os.path.join(sLogDir, sApp + "-testSay-" + sStamp + ".log")
sOutDir = os.path.join(sLogDir, "testSay-" + sStamp)


def logLine(sText):
    os.makedirs(sLogDir, exist_ok=True)
    with open(sLogPath, "a", encoding="utf-8-sig", newline="\r\n") as fLog:
        fLog.write(datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S") + " " + sText + "\n")
    return True


def say(sText):
    print(sText, flush=True)
    logLine("CONSOLE: " + sText)
    return True


def findRunMacro():
    """runMacro.py beside this script, in the app's scripts, or in the kit."""
    lsCandidates = [os.path.join(sScriptDir, "runMacro.py"), os.path.join(sProject, "scripts", "runMacro.py"),
                    os.path.join(os.path.dirname(sProject), "HomerDev", "scripts", "runMacro.py"),
                    os.path.join(os.path.splitdrive(sProject)[0] + os.sep, "HomerDev", "scripts", "runMacro.py")]
    for sCandidate in lsCandidates:
        logLine("runMacro candidate " + sCandidate + (" found" if os.path.isfile(sCandidate) else " not found"))
        if os.path.isfile(sCandidate): return sCandidate
    return ""


def sayCommands(oRunMacro, sOnly, sLetters):
    """[(name, key)] of the app's Say commands that have a key, by name, narrowed as asked."""
    dCommands = oRunMacro.appCommands(oRunMacro.keyFileOf(sProject, sApp))
    lsChosen = sorted(((sName, sKey) for sName, sKey in dCommands.values() if sName.lower().startswith("say ") and sKey),
                      key=lambda t: t[0].lower())
    if sLetters:
        setKeys = set("shift+" + s.lower() for s in sLetters)
        lsChosen = [t for t in lsChosen if t[1].lower() in setKeys]
    if sOnly:
        lsWanted = [s.strip().lower() for s in sOnly.split(",") if s.strip()]
        def wanted(sName, sKey):
            sBare = sKey.lower().split("+")[-1]
            # A single letter or digit names a key (F is Shift+F); anything longer is a key or words of a name.
            return any(s == sBare or s == sKey.lower() or (len(s) > 1 and s in sName.lower()) for s in lsWanted)
        lsChosen = [t for t in lsChosen if wanted(*t)]
    return lsChosen


def writeMacro(sPath, lsChosen, iWait, bDescriber):
    """The macro script: each Say command once, twice and Escape; then each key described."""
    lsLines = ["FileTask = macroScript", "App = " + sApp, "Name = Every Say command of " + sApp + ", heard",
               "Listen = " + str(iWait), "Note = written by testSay on " + sStamp, "", "[steps]"]
    for sName, sKey in lsChosen:
        lsLines += [sName, sName + " twice", "Press Escape"]
    if bDescriber:
        lsLines.append("Press Control+F1")
        lsLines += ["Press " + sKey for sName, sKey in lsChosen]
        lsLines.append("Press Control+F1")
    with open(sPath, "w", encoding="utf-8-sig", newline="\r\n") as fOut:
        fOut.write("\n".join(lsLines) + "\n")
    logLine("macro written: %s, %d lines under [steps]" % (sPath, len(lsLines) - 7))
    return sPath


def newerThan(sPattern, nSince):
    return sorted(s for s in glob.glob(sPattern) if os.path.getmtime(s) >= nSince)


def copyInto(lsFiles, sFolder):
    os.makedirs(sFolder, exist_ok=True)
    for sFile in lsFiles:
        shutil.copy2(sFile, sFolder)
        logLine("copied " + sFile + " into " + sFolder)
    return len(lsFiles)


def suiteApps():
    """The Homer apps beside the kit (or at the drive's root) that have a key file and a scripts folder."""
    lsFound = []
    for sParent in (os.path.dirname(sProject), os.path.splitdrive(sProject)[0] + os.sep):
        for sName in ("DbDo", "EdSharp", "FileDir", "HomerView"):
            sDir = os.path.join(sParent, sName)
            if sName not in [os.path.basename(s) for s in lsFound] and os.path.isdir(os.path.join(sDir, "configs")) and os.path.isdir(os.path.join(sDir, "scripts")): lsFound.append(sDir)
    return lsFound


def runSuite(oArgs):
    """From HomerDev: this script once for each app, in the app's folder, with the same options."""
    lsApps = suiteApps()
    if oArgs.apps: lsApps = [s for s in lsApps if os.path.basename(s).lower() in [t.strip().lower() for t in oArgs.apps.split(",")]]
    logLine("suite run over: " + "; ".join(lsApps))
    if not lsApps:
        say("No Homer app was found beside HomerDev, so nothing was pressed.")
        return 2
    lsPass = [s for s in sys.argv[1:] if not s.lower().startswith("--apps")]
    if "--apps" in [s.lower() for s in sys.argv[1:]]:
        iAt = [s.lower() for s in sys.argv[1:]].index("--apps"); lsPass = sys.argv[1:iAt + 1] + sys.argv[iAt + 3:]
    iMinutes = 0
    sRunMacro = findRunMacro()
    if sRunMacro:
        sys.path.insert(0, os.path.dirname(sRunMacro))
        import runMacro as oRunMacro
        for sDir in lsApps:
            sName = os.path.basename(sDir)
            iCount = sum(1 for sCommand, sKey in oRunMacro.appCommands(oRunMacro.keyFileOf(sDir, sName)).values() if sCommand.lower().startswith("say ") and sKey)
            iMinutes += (iCount * 3 + (0 if oArgs.no_describer else iCount + 2)) * (oArgs.wait + 3) // 60
    say("testSay from HomerDev: %d app%s, %s, about %d minutes in all. Control+C stops it." % (
        len(lsApps), "" if len(lsApps) == 1 else "s", ", ".join(os.path.basename(s) for s in lsApps), max(1, iMinutes)))
    iWorst = 0
    for sDir in lsApps:
        lCommand = [sys.executable, os.path.abspath(__file__)] + lsPass
        logLine("run: " + subprocess.list2cmdline(lCommand) + " in " + sDir)
        iCode = subprocess.call(lCommand, cwd=sDir)
        logLine("%s exit code %d" % (os.path.basename(sDir), iCode))
        iWorst = max(iWorst, iCode)
    say("testSay from HomerDev finished; each app's results are in its own logs folder.")
    return iWorst


def main():
    oParser = argparse.ArgumentParser(description="Run every Say command of a Homer app, record what the reader says, and gather it under logs.")
    oParser.add_argument("--only", default="")
    oParser.add_argument("--letters", default="")
    oParser.add_argument("--no-describer", action="store_true")
    oParser.add_argument("--wait", type=int, default=6)
    oParser.add_argument("--check", action="store_true")
    oParser.add_argument("--apps", default="")
    oArgs = oParser.parse_args([s.lower() if s.startswith("--") else s for s in sys.argv[1:]])
    sLetters = "".join(s for s in oArgs.letters.upper() if "A" <= s <= "Z")
    logLine("testSay started; script=%s; Python %s; %s; cwd=%s; command line=%s" % (
        os.path.abspath(__file__), sys.version.split()[0], platform.platform(), os.getcwd(), " ".join(sys.argv)))
    logLine("settings: app=%s; project=%s; only=%s; letters=%s; wait=%d; describer=%s; check=%s; folder=%s" % (
        sApp, sProject, oArgs.only or "all", sLetters or "all", oArgs.wait, "no" if oArgs.no_describer else "yes", oArgs.check, sOutDir))
    if sApp.lower() == "homerdev" or oArgs.apps: return runSuite(oArgs)
    sRunMacro = findRunMacro()
    if not sRunMacro:
        say("runMacro.py was not found in this app's scripts or the kit's, so nothing was pressed.")
        return 2
    sys.path.insert(0, os.path.dirname(sRunMacro))
    import runMacro as oRunMacro
    lsChosen = sayCommands(oRunMacro, oArgs.only, sLetters)
    logLine("Say commands chosen: " + "; ".join("%s (%s)" % t for t in lsChosen))
    if not lsChosen:
        say("No Say command with a key matched, so nothing was pressed.")
        return 2
    os.makedirs(sOutDir, exist_ok=True)
    sMacro = writeMacro(os.path.join(sOutDir, "testSay.inix"), lsChosen, oArgs.wait, not oArgs.no_describer)
    lCommand = [sys.executable, sRunMacro, sMacro, "--keep"] + (["--check"] if oArgs.check else [])
    if not oArgs.check:
        iSteps = len(lsChosen) * 3 + (0 if oArgs.no_describer else len(lsChosen) + 2)
        iMinutes = max(1, iSteps * (oArgs.wait + 3) // 60)
        if iMinutes > 10: say("This takes about %d minutes. Control+C stops it." % iMinutes)
        say("testSay: %d Say command%s in %s, %d key%s, about %d minute%s." % (len(lsChosen), "" if len(lsChosen) == 1 else "s", sApp, iSteps, "" if iSteps == 1 else "s", iMinutes, "" if iMinutes == 1 else "s"))
    nStart = time.time()
    logLine("run: " + subprocess.list2cmdline(lCommand) + " in " + sProject)
    try:
        iCode = subprocess.call(lCommand, cwd=sProject)
    except Exception:
        logLine("runMacro did not start: " + traceback.format_exc())
        say("runMacro could not be started; the log says why.")
        return 2
    logLine("runMacro exit code %d after %.0f seconds" % (iCode, time.time() - nStart))
    if oArgs.check: return iCode
    # GATHERED IN ONE FOLDER, to upload with logs: the reports and logs, the recordings, the app's own logs.
    iReports = copyInto(newerThan(os.path.join(sLogDir, sApp + "-recordWalk-*.*"), nStart) + newerThan(os.path.join(sLogDir, sApp + "-runMacro-*.log"), nStart), sOutDir)
    iWaves = 0
    for sFolder in newerThan(os.path.join(sProject, "results", "recordWalk-*"), nStart):
        if os.path.isdir(sFolder): iWaves += copyInto(sorted(glob.glob(os.path.join(sFolder, "*.wav"))), os.path.join(sOutDir, "recordings"))
    sAppLogs = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), sApp, "logs")
    iAppLogs = copyInto(newerThan(os.path.join(sAppLogs, sApp + "-*.log"), nStart - 60), os.path.join(sOutDir, "appLogs"))
    with open(os.path.join(sOutDir, "testSay.md"), "w", encoding="utf-8-sig", newline="\r\n") as fOut:
        fOut.write("\n".join([
            "# testSay, " + sApp + ", " + sStamp, "",
            "- Say commands: %d; each run once, twice, then Escape%s." % (len(lsChosen), "" if oArgs.no_describer else "; then each key described with the Key Describer on"),
            "- The macro played: testSay.inix, which runMacro can play again.",
            "- runMacro exit code: %d (0 means every expectation held; this macro expects nothing, so 1 is usual)." % iCode,
            "- Reports and logs: %d file%s." % (iReports, "" if iReports == 1 else "s"),
            "- Recordings: %d, in recordings." % iWaves,
            "- %s's own logs: %d, in appLogs." % (sApp, iAppLogs),
            "- This script's log: " + os.path.basename(sLogPath) + ", in logs.", ""]))
    say("testSay finished: %d recording%s and %d report file%s in %s. Upload the logs folder." % (
        iWaves, "" if iWaves == 1 else "s", iReports, "" if iReports == 1 else "s", sOutDir))
    return iCode


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception:
        logLine("ERROR " + traceback.format_exc())
        print("testSay stopped on an error; the log says why: " + sLogPath)
        sys.exit(2)
