r"""runMacro.py -- play a Homer macro script: plain lines of commands, run in the real app, heard through the reader.

WHY THIS EXISTS (17 November 2026)

Jamal asked whether the Homer command set and the screen readers' command set, both named verb then noun, could be
the words of a macro a person reads and writes as easily as a to-do list, with no macro editor: any text editor
opens it, EdSharp included. This is that. A macro script is an .inix file whose header says FileTask = macroScript,
and whose [steps] section holds one plain line per step:

    FileTask = macroScript
    App = FileDir
    Name = Hear the filter and the position

    [steps]
    ; a line starting with a semicolon is a note
    Say Filter and Order
    Expect speech filter
    Say Position twice
    Press Escape
    Say Line
    Expect speech filter

Each line is one of these, matched in this order:

    a command of the app        its name as the menus give it, such as Say Filter and Order; run by its key, or
                                through the Alternate Menu when it has none. "twice" at the end presses its key twice.
    a screen reader command     its script name, with or without spaces, such as Say Line or SayWindowTitle; run by
                                the key JAWS gives it in its own key map (NVDA's key when NVDA is the reader)
    Expect focus <text>         the control with focus after the step before holds this text
    Expect speech <text>        the reader was heard saying this after the step before (most of its words)
    Expect window <text>        the window in front after the step before has this in its title
    JAWS <script>               a JAWS script by name, when the app has a command of the same name
    Press <key>                 a key by its Homer name, such as Control+Shift+F or Escape; "twice" presses it twice
    Run <command>               an app command by name, when a reader script has the same name
    Start <app>                 brings the app forward, started when it is not running (every macro does this first)
    Type <text>                 types the text, letter by letter; quotation marks around it are dropped
    Wait <seconds>              waits before the next step
    Wait for window <text>      the same as Expect window

WHAT IS HIDDEN FROM THE PERSON. Nothing in a macro says how long anything takes. Before each key the runner waits for
the reader to fall quiet; after it, it listens until the reader has finished, up to fifteen seconds; a window is given
five more seconds to appear. The kit's recordWalk does the pressing, the listening, the recording and the writing
down, so a macro and a tutorial walk are played by one engine, and the macro is first translated into a walk.

HOW A COMMAND BECOMES KEYS. The app's key file (configs\Hotkeys.inix, or DbDo's configs\DbDo.inix) names every
command and its keys; the first key without a reader key is pressed. A command with no key is run from the Alternate
Menu: Alt+F10, Control+F, the command's name, Enter to filter, Enter to run it. A reader command is looked up in JAWS's
own default.jkm, the user's copy first, so a reassigned key is followed.

USAGE

    runMacro                             every macro in the app's scripts\tests folder; from HomerDev, every app's
    runMacro path\to\macro.inix          play it, with the app in front; the report says what was heard and seen
    runMacro --check path\to\macro.inix  translate only: say which lines are not understood (exit 1 when any)
    runMacro --no-hear path\macro.inix   play it without writing the speech down (Expect speech is then not checked)
    runMacro --keep path\macro.inix      leave the app open at the end, even when the macro started it

Writes logs\<App>-runMacro-<stamp>.log and the walk it plays, logs\<App>-macro-<stamp>.inix, then recordWalk's report
and log. Exit code: 0 when every Expect line held, 1 when one did not, 2 when the macro could not be played.

TIDY RUNS THE IMPORTANT ONES. Macros in an app's scripts\tests folder are the app's live tests. Tidy calls
liveTests() here after every build: each test is played when its fingerprint -- the macro, the program, the key file
and these scripts -- has changed since its last pass, the app is closed (so the build just made is the one tested, and
nothing Jamal has open is touched), and a screen reader is running. A failed live test is reported, never a reason for
tidy to fail; it runs again on the next tidy, because only a pass is recorded.
"""

import argparse, datetime, glob, os, platform, re, subprocess, sys, time, traceback

sStamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
sScriptDir = os.path.dirname(os.path.abspath(__file__))
sFileTask = "macroScript"
iDefaultListen = 6
sLogPath = ""
fnOuterLog = None

# FileDir names a few keys by the character they type; each is the key that types it on a US keyboard.
dCharacterKeys = {"accent": "GraveAccent", "exclamation": "Shift+1", "leftbrace": "Shift+LeftBracket",
                  "leftparen": "Shift+9", "question": "Shift+Slash", "rightbrace": "Shift+RightBracket",
                  "rightparen": "Shift+0", "tilde": "Shift+GraveAccent"}
# Names people say for a reader command, by the name JAWS gives its script.
dReaderAliases = {"saycontrol": "saywindowpromptandtext", "sayfocus": "saywindowpromptandtext",
                  "saytitle": "saywindowtitle"}
# When JAWS's key map cannot be read, or NVDA is the reader: the few every reader user knows.
dJawsFallback = {"sayall": "Insert+DownArrow", "sayline": "Insert+UpArrow", "saywindowpromptandtext": "Insert+Tab",
                 "saywindowtitle": "Insert+T"}
dNvdaKeys = {"sayall": "NVDA+DownArrow", "sayline": "NVDA+UpArrow", "saywindowpromptandtext": "NVDA+Tab",
             "saywindowtitle": "NVDA+T"}


def logLine(sText):
    """One line to the detailed log, and to tidy's when tidy is the caller."""
    if fnOuterLog: fnOuterLog("runMacro: " + sText)
    if not sLogPath: return False
    with open(sLogPath, "a", encoding="utf-8-sig") as fLog:
        fLog.write(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "  " + sText + "\n")
    return True


def sayLine(sText):
    print(sText, flush=True)
    logLine("CONSOLE: " + sText)
    return True


def countNoun(iCount, sNoun):
    return "%d %s%s" % (iCount, sNoun, "" if iCount == 1 else "s")


def projectOf(sMacro):
    """The app's folder: the one holding a scripts folder, above the macro or the working folder."""
    for sStart in (os.path.dirname(os.path.abspath(sMacro)), os.getcwd()):
        sDir = sStart
        while sDir and os.path.dirname(sDir) != sDir:
            if os.path.isdir(os.path.join(sDir, "scripts")) and os.path.isdir(os.path.join(sDir, "configs")): return sDir
            sDir = os.path.dirname(sDir)
    return os.getcwd()


# ---- reading a macro ----

def readMacro(sPath):
    """(header dictionary, [(line number, step text)]). The header is the lines before the first section, read as
    .inix; the [steps] section is read line by line, as written, so a step may hold an equals sign."""
    with open(sPath, "r", encoding="utf-8-sig", errors="replace") as fIn:
        lsLines = fIn.read().splitlines()
    dHeader, lSteps, sSection = {}, [], ""
    for iLine, sRaw in enumerate(lsLines, 1):
        sLine = sRaw.strip()
        if not sLine or sLine.startswith(";") or sLine.startswith("#"): continue
        if sLine.startswith("[") and sLine.endswith("]"):
            sSection = sLine[1:-1].strip().lower()
            continue
        if sSection == "":
            if "=" in sLine:
                sName, sValue = [s.strip() for s in sLine.split("=", 1)]
                dHeader[sName.lower()] = sValue.strip('"')
        elif sSection == "steps":
            lSteps.append((iLine, sLine))
    return dHeader, lSteps


# ---- the app's commands ----

def keyFileOf(sProject, sApp):
    for sName in (sApp + ".inix", "Hotkeys.inix"):
        sPath = os.path.join(sProject, "configs", sName)
        if os.path.isfile(sPath) and "[hotkeys]" in open(sPath, encoding="utf-8-sig", errors="replace").read().lower(): return sPath
    return ""


def homerKey(sKey):
    """A key from a key file in the walk's spelling: FileDir's character names turned into keys, JAWSKey into JAWS."""
    lParts = [s.strip() for s in re.split(r"\+(?!$)", sKey.strip()) if s.strip()]
    if not lParts: return ""
    sLast = lParts[-1]
    if sLast.lower() in dCharacterKeys:
        lMore = dCharacterKeys[sLast.lower()].split("+")
        lParts = lParts[:-1] + [s for s in lMore[:-1] if s not in lParts[:-1]] + lMore[-1:]
    return "+".join("JAWS" if s.lower() == "jawskey" else s for s in lParts)


def appCommands(sKeyFile):
    """{lower-case command name: (name, first key or "")} from a key file's [Hotkeys] section. The first key that needs
    no reader key is the one pressed; a backtick-fenced value is skipped, as it is not a command."""
    dCommands = {}
    if not sKeyFile: return dCommands
    sSection, bFence = "", False
    for sRaw in open(sKeyFile, encoding="utf-8-sig", errors="replace").read().splitlines():
        sLine = sRaw.strip()
        if bFence:
            if sLine == "`": bFence = False
            continue
        if not sLine or sLine.startswith(";"): continue
        if sLine.startswith("["): sSection = sLine.lower(); continue
        if sSection != "[hotkeys]" or "=" not in sLine: continue
        if sLine.endswith("=`"): bFence = True; continue
        sName, sRest = [s.strip() for s in sLine.split("=", 1)]
        oMatch = re.match(r"^((?:[^,]|,(?=\s*or\s))*)", sRest)
        sKeys = oMatch.group(1).strip() if oMatch else ""
        lKeys = [] if sKeys.lower() in ("", "none", "no key") else [s.strip() for s in re.split(r"\s*,?\s+or\s+|\s*,\s*", sKeys) if s.strip()]
        lPlain = [homerKey(s) for s in lKeys if not re.search(r"(^|\+)(jaws|jawskey|nvda|insert|windows)\+", s, re.I)]
        dCommands[sName.lower()] = (sName, lPlain[0] if lPlain else "")
    return dCommands


# ---- the reader's commands ----

def processNames():
    """The lower-case program names of every running process, read from Windows directly.

    NOT TASKLIST (17 November 2026): with five build_release cycles running at once, tasklist took more than the 30
    seconds allowed, so FileDir's tidy read "no screen reader" with JAWS running. A process snapshot takes
    milliseconds."""
    import ctypes
    from ctypes import wintypes
    class ProcessEntry(ctypes.Structure):
        _fields_ = [("dwSize", wintypes.DWORD), ("cntUsage", wintypes.DWORD), ("th32ProcessID", wintypes.DWORD),
                    ("th32DefaultHeapID", ctypes.c_size_t), ("th32ModuleID", wintypes.DWORD), ("cntThreads", wintypes.DWORD),
                    ("th32ParentProcessID", wintypes.DWORD), ("pcPriClassBase", ctypes.c_long), ("dwFlags", wintypes.DWORD),
                    ("szExeFile", ctypes.c_wchar * 260)]
    oKernel = ctypes.windll.kernel32
    oKernel.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
    hSnap = oKernel.CreateToolhelp32Snapshot(2, 0)  # TH32CS_SNAPPROCESS
    setNames = set()
    if not hSnap or hSnap == wintypes.HANDLE(-1).value: return setNames
    try:
        oEntry = ProcessEntry()
        oEntry.dwSize = ctypes.sizeof(ProcessEntry)
        bMore = oKernel.Process32FirstW(hSnap, ctypes.byref(oEntry))
        while bMore:
            setNames.add(oEntry.szExeFile.lower())
            bMore = oKernel.Process32NextW(hSnap, ctypes.byref(oEntry))
    finally:
        oKernel.CloseHandle(hSnap)
    return setNames


def readerRunning():
    """"jaws", "nvda" or "" -- from the running programs."""
    try:
        setNames = processNames()
    except Exception as oError:
        logLine("the running programs could not be read: %s" % oError)
        return ""
    if "jfw.exe" in setNames: return "jaws"
    if "nvda.exe" in setNames: return "nvda"
    return ""


def jawsKeyMaps():
    """JAWS's default.jkm files, the newest version's, the user's copy first so a reassigned key is followed."""
    lFound = []
    # JAWS's own folders: the user's settings live in the Roaming tree, which is theirs, not a Homer app's.
    for sJawsFolder in (os.path.join(os.environ.get("APPDATA", "") or "?", "Freedom Scientific", "JAWS"), os.path.join(os.environ.get("ProgramData", "") or "?", "Freedom Scientific", "JAWS")):
        lVersions = sorted(glob.glob(os.path.join(sJawsFolder, "*", "Settings", "enu", "default.jkm")),
                           key=lambda s: [int(x) if x.isdigit() else 0 for x in re.split(r"[.\\/]", s)], reverse=True)
        if lVersions: lFound.append(lVersions[0])
    return lFound


def readerCommands(sReader):
    """{script name in lower case, no spaces: key} for the running reader."""
    if sReader == "nvda": return dict(dNvdaKeys)
    dScripts = {}
    for sPath in jawsKeyMaps():
        sSection = ""
        try:
            for sRaw in open(sPath, encoding="utf-8-sig", errors="replace").read().splitlines():
                sLine = sRaw.strip()
                if sLine.startswith("["): sSection = sLine.lower(); continue
                if sSection not in ("[common keys]", "[desktop keys]") or "=" not in sLine or sLine.startswith(";"): continue
                sKey, sScript = [s.strip() for s in sLine.split("=", 1)]
                dScripts.setdefault(sScript.lower(), sKey)
            logLine("JAWS key map read: " + sPath)
        except OSError as oError:
            logLine("JAWS key map unreadable: %s: %s" % (sPath, oError))
    for sScript, sKey in dJawsFallback.items(): dScripts.setdefault(sScript, sKey)
    return dScripts


# ---- translating ----

def translate(lSteps, dCommands, dReader, sApp):
    """([walk step dictionaries], [problems]). Each step: sKey, sType, iPress, sSay, lHear, lWindow, lFocus, nPause."""
    lWalk, lsProblems = [], []
    nPause = 0.0
    def addStep(sSay, sKey="", sType="", iPress=1):
        nonlocal nPause
        lWalk.append({"sKey": sKey, "sType": sType, "iPress": iPress, "sSay": sSay, "lHear": [], "lWindow": [], "lFocus": [], "nPause": nPause})
        nPause = 0.0
    def expect(iLine, sKind, sText):
        if not lWalk:
            lsProblems.append("line %d: Expect %s comes before any step to expect it of" % (iLine, sKind))
            return
        lWalk[-1]["l" + sKind.capitalize()].append(sText)
    def runCommand(iLine, sName, sText, iPress):
        sShown, sKey = dCommands[sName.lower()]
        if sKey:
            addStep(sText, sKey=sKey, iPress=iPress)
            return
        # NO KEY: the Alternate Menu, filtered to the command's own line, which begins "<name>, ".
        addStep(sText + ": the Alternate Menu", sKey="Alt+F10")
        lWalk[-1]["lWindow"].append("Alternate Menu")
        addStep(sText + ": filter the list", sKey="Control+F")
        addStep(sText + ": its name", sType=sShown + ",", sKey="Enter")
        addStep(sText + ": run it", sKey="Enter", iPress=iPress)
    def runReader(iLine, sName, sText, iPress):
        sScript = re.sub(r"\s+", "", sName).lower()
        sScript = dReaderAliases.get(sScript, sScript)
        if sScript not in dReader:
            lsProblems.append("line %d: the screen reader has no command called %s" % (iLine, sName))
            return
        addStep(sText, sKey=dReader[sScript], iPress=iPress)
    for iLine, sText in lSteps:
        sLine = re.sub(r"\s+", " ", sText).strip()
        iPress = 1
        oTwice = re.match(r"^(.*\S)\s+twice$", sLine, re.I)
        sBody = oTwice.group(1) if oTwice else sLine
        if oTwice: iPress = 2
        sLower = sBody.lower()
        oVerb = re.match(r"^(\w+)\s*(.*)$", sBody)
        sVerb, sRest = (oVerb.group(1).lower(), oVerb.group(2).strip()) if oVerb else ("", "")
        sArg = sRest[1:-1] if len(sRest) >= 2 and sRest[0] == sRest[-1] == '"' else sRest
        if sLower in dCommands: runCommand(iLine, sBody, sText, iPress)
        elif re.sub(r"\s+", "", sLower) in dReader or dReaderAliases.get(re.sub(r"\s+", "", sLower), "") in dReader:
            runReader(iLine, sBody, sText, iPress)
        elif sVerb == "expect":
            oKind = re.match(r"^(focus|speech|window)\b\s*(.*)$", sRest, re.I)
            if not oKind:
                lsProblems.append("line %d: Expect is followed by focus, speech or window: %s" % (iLine, sText)); continue
            sKind, sWhat = oKind.group(1).lower(), oKind.group(2).strip().strip('"')
            if not sWhat: lsProblems.append("line %d: Expect %s needs the text to look for" % (iLine, sKind)); continue
            expect(iLine, "hear" if sKind == "speech" else sKind, sWhat)
        elif sVerb == "jaws": runReader(iLine, sRest, sText, iPress)
        elif sVerb == "press":
            if not sArg: lsProblems.append("line %d: Press needs a key" % iLine); continue
            addStep(sText, sKey=homerKey(sArg), iPress=iPress)
        elif sVerb == "run":
            if sArg.lower() in dCommands: runCommand(iLine, sArg, sText, iPress)
            else: lsProblems.append("line %d: %s has no command called %s" % (iLine, sApp, sArg))
        elif sVerb == "start":
            if sArg and sArg.lower() != sApp.lower(): lsProblems.append("line %d: this macro plays in %s, not %s" % (iLine, sApp, sArg))
        elif sVerb == "type":
            if not sArg: lsProblems.append("line %d: Type needs text" % iLine); continue
            addStep(sText, sType=sArg)
        elif sVerb == "wait":
            oFor = re.match(r"^for\s+window\s+(.+)$", sRest, re.I)
            if oFor: expect(iLine, "window", oFor.group(1).strip().strip('"')); continue
            try: nPause += float(sArg.split()[0].rstrip("s"))
            except (ValueError, IndexError): lsProblems.append("line %d: Wait needs a number of seconds, or for window: %s" % (iLine, sText))
        else:
            lsProblems.append("line %d: not a command of %s, a screen reader command, or a step this runner knows: %s" % (iLine, sApp, sText))
    return lWalk, lsProblems


def writeWalk(sPath, lWalk, sMacro, iListen):
    lsLines = ["; written by runMacro from " + sMacro + ", " + sStamp, ""]
    for dStep in lWalk:
        lsLines += ["[step]", "Say=" + dStep["sSay"]]
        if dStep["sType"]: lsLines.append("Type=" + dStep["sType"])
        if dStep["sKey"]: lsLines.append("Key=" + dStep["sKey"])
        if dStep["iPress"] > 1: lsLines.append("Press=" + str(dStep["iPress"]))
        if dStep["nPause"]: lsLines.append("Pause=" + str(dStep["nPause"]))
        lsLines.append("Wait=" + str(iListen + (2 if dStep["iPress"] > 1 else 0)))
        lsLines += ["Hear=" + s for s in dStep["lHear"]]
        lsLines += ["Window=" + s for s in dStep["lWindow"]] + ["Focus=" + s for s in dStep["lFocus"]] + [""]
    with open(sPath, "w", encoding="utf-8-sig", newline="\r\n") as fOut:
        fOut.write("\n".join(lsLines))
    return sPath


# ---- the app's window ----

def appIsOpen(sApp):
    """True when <App>.exe is running, whichever copy."""
    try:
        return (sApp + ".exe").lower() in processNames()
    except Exception as oError:
        logLine("the running programs could not be read: %s" % oError)
        return False


def findRecordWalk(sProject):
    lsCandidates = [os.path.join(sScriptDir, "recordWalk.py"), os.path.join(sProject, "scripts", "recordWalk.py")]
    try:
        import kind as oKind
        sKit = oKind.findKit([sProject])
        if sKit: lsCandidates.append(os.path.join(sKit, "scripts", "recordWalk.py"))
    except Exception:
        pass
    return next((s for s in lsCandidates if os.path.isfile(s)), "")


# ---- playing ----

def play(sMacro, bCheck=False, bHear=True, bKeep=False, sProject="", bStartProject=False):
    """Plays one macro; returns the exit code."""
    global sLogPath
    sMacro = os.path.abspath(sMacro)
    sProject = sProject or projectOf(sMacro)
    dHeader, lSteps = readMacro(sMacro)
    sApp = dHeader.get("app", "") or os.path.basename(sProject)
    os.makedirs(os.path.join(sProject, "logs"), exist_ok=True)
    if not sLogPath:
        sLogPath = os.path.join(sProject, "logs", sApp + "-runMacro-" + sStamp + ".log")
    logLine("runMacro: %s; Python %s; %s; cwd %s; command line %s" % (os.path.abspath(__file__), sys.version.split()[0], platform.platform(), os.getcwd(), " ".join(sys.argv)))
    logLine("settings: macro %s; project %s; app %s; check %s; hear %s; keep %s; start the project's build %s" % (sMacro, sProject, sApp, bCheck, bHear, bKeep, bStartProject))
    if dHeader.get("filetask", "").lower() != sFileTask.lower():
        sayLine("%s is not a macro script: its header needs FileTask = %s." % (os.path.basename(sMacro), sFileTask))
        return 2
    if not lSteps:
        sayLine("%s has no steps: they go under [steps], one to a line." % os.path.basename(sMacro))
        return 2
    sKeyFile = keyFileOf(sProject, sApp)
    dCommands = appCommands(sKeyFile)
    logLine("key file %s: %d commands" % (sKeyFile or "(none found)", len(dCommands)))
    sReader = readerRunning() if os.name == "nt" else ""
    dReader = readerCommands(sReader or "jaws")
    logLine("reader running: %s; %d reader commands known" % (sReader or "none", len(dReader)))
    lWalk, lsProblems = translate(lSteps, dCommands, dReader, sApp)
    for sProblem in lsProblems: logLine("PROBLEM " + sProblem)
    try: iListen = max(1, min(30, int(dHeader.get("listen", "") or iDefaultListen)))
    except ValueError: iListen = iDefaultListen
    sWalk = writeWalk(os.path.join(sProject, "logs", sApp + "-macro-" + os.path.splitext(os.path.basename(sMacro))[0] + "-" + sStamp + ".inix"), lWalk, sMacro, iListen)
    logLine("walk written: %s, %s" % (sWalk, countNoun(len(lWalk), "step")))
    if lsProblems:
        sayLine("%s: %d line%s not understood, so it was not played:" % (os.path.basename(sMacro), len(lsProblems), "" if len(lsProblems) == 1 else "s"))
        for sProblem in lsProblems: sayLine("  " + sProblem)
        return 1 if bCheck else 2
    if bCheck:
        sayLine("%s: %s, all understood." % (os.path.basename(sMacro), countNoun(len(lWalk), "step")))
        return 0
    if os.name != "nt":
        sayLine("A macro plays only on Windows.")
        return 2
    sRecordWalk = findRecordWalk(sProject)
    if not sRecordWalk:
        sayLine("recordWalk.py was not found beside runMacro, in the app's scripts or in the kit, so nothing was played.")
        return 2
    bWasOpen = appIsOpen(sApp)
    lCommand = [sys.executable, sRecordWalk, "--walk", sWalk, "--app", sApp, "--no-axe"]
    if bKeep or bWasOpen: lCommand.append("--keep")       # what was open before stays open
    if not bHear: lCommand.append("--no-transcribe")
    sExe = os.path.join(sProject, "exec", sApp + ".exe")
    if bStartProject and os.path.isfile(sExe): lCommand += ["--exe", sExe]
    sayLine("Playing %s: %s in %s. Keep your hands off the keyboard; switch windows to stop." % (dHeader.get("name", os.path.basename(sMacro)), countNoun(len(lWalk), "step"), sApp))
    logLine("run: " + subprocess.list2cmdline(lCommand) + " in " + sProject)
    try:
        iCode = subprocess.call(lCommand, cwd=sProject)
    except Exception:
        logLine("recordWalk did not start: " + traceback.format_exc())
        sayLine("recordWalk could not be started; the log says why.")
        return 2
    logLine("recordWalk exit code %d" % iCode)
    return iCode


# ---- tidy's live tests ----

def liveTests(sProject, fnLog=None, fnSay=None):
    """Plays the app's scripts\\tests\\*.inix macros whose fingerprint changed since their last pass. Returns a list of
    one-line results; never raises, so tidy goes on whatever happens here."""
    global fnOuterLog, sLogPath
    fnOuterLog = fnLog
    fnSay = fnSay or sayLine
    lsDone = []
    try:
        sApp = os.path.basename(os.path.normpath(sProject))
        lsTests = sorted(glob.glob(os.path.join(sProject, "scripts", "tests", "*.inix")), key=str.lower)
        if not lsTests: return lsDone
        sLogPath = os.path.join(sProject, "logs", sApp + "-runMacro-" + sStamp + ".log")
        # Every test is first read and translated: a line nobody understands is a finding even when nothing is played.
        for sTest in lsTests:
            if play(sTest, bCheck=True, sProject=sProject) != 0: lsDone.append("%s has lines runMacro does not understand; the log names them" % os.path.basename(sTest))
        if lsDone: return lsDone
        if os.name != "nt": return lsDone
        if os.environ.get("HomerLive", "") == "0":
            logLine("HomerLive=0, so the live tests are not played"); return lsDone
        sExe = os.path.join(sProject, "exec", sApp + ".exe")
        if not os.path.isfile(sExe):
            logLine("no build at %s, so the live tests are not played" % sExe); return lsDone
        sKit = sScriptDir
        sys.path.insert(0, sKit)
        import fingerprint
        oCache = fingerprint.Cache(sProject, fnLog=logLine)
        lsInputs = [sExe, keyFileOf(sProject, sApp), os.path.join(sKit, "runMacro.py"), os.path.join(sKit, "recordWalk.py")]
        lsDue = []
        for sTest in lsTests:
            sPrint = oCache.fingerprint([s for s in lsInputs + [sTest] if s and os.path.isfile(s)])
            sKey = "live test " + os.path.basename(sTest)
            if oCache.isCurrent(sKey, sPrint, []): logLine("%s: passed with these files before; not played again" % sKey)
            else: lsDue.append((sTest, sKey, sPrint))
        if not lsDue: return lsDone
        if appIsOpen(sApp):
            fnSay("%s is open, so %s for a tidy when it is closed." % (sApp, "its live test waits" if len(lsDue) == 1 else "its %d live tests wait" % len(lsDue)))
            return lsDone
        if not readerRunning():
            fnSay("No screen reader is running, so %s's live tests wait for a tidy with JAWS or NVDA running." % sApp)
            return lsDone
        for sTest, sKey, sPrint in lsDue:
            fnSay("Live test %s: %s will open, and keys will be pressed in it for about a minute. Keep your hands off the keyboard." % (os.path.splitext(os.path.basename(sTest))[0], sApp))
            iCode = play(sTest, sProject=sProject, bStartProject=True)
            if iCode == 0:
                oCache.record(sKey, sPrint, [])
                lsDone.append("live test %s passed" % os.path.basename(sTest))
            else:
                lsDone.append("live test %s did not pass (exit %d); recordWalk's report in logs says what was heard" % (os.path.basename(sTest), iCode))
        for s in lsDone: fnSay(s[:1].upper() + s[1:] + ".")
    except Exception:
        logLine("live tests stopped on an error: " + traceback.format_exc())
        lsDone.append("the live tests stopped on an error; the log says why")
    return lsDone


def defaultMacros():
    """NO ARGUMENT IS NEEDED (17 November 2026): with no macro named, the most likely wish is the app's own tests --
    every .inix in scripts\\tests of the app this is run from, or, from HomerDev, of every Homer app beside it."""
    sHere = os.getcwd()
    for sDir in (sHere, os.path.dirname(sHere)):
        if os.path.isdir(os.path.join(sDir, "scripts")) and os.path.isdir(os.path.join(sDir, "configs")): sHere = sDir; break
    lsFolders = [sHere]
    if os.path.basename(os.path.normpath(sHere)).lower() == "homerdev":
        lsFolders = [os.path.join(os.path.dirname(sHere), s) for s in ("DbDo", "EdSharp", "FileDir", "HomerView")]
    lsMacros = []
    for sFolder in lsFolders: lsMacros += sorted(glob.glob(os.path.join(sFolder, "scripts", "tests", "*.inix")))
    return lsMacros


def main():
    oParser = argparse.ArgumentParser(description="Play a Homer macro script in its app, heard through the screen reader.")
    oParser.add_argument("macro", nargs="*")
    oParser.add_argument("--check", action="store_true")
    oParser.add_argument("--no-hear", action="store_true")
    oParser.add_argument("--keep", action="store_true")
    oArgs = oParser.parse_args([s.lower() if s.startswith("--") else s for s in sys.argv[1:]])
    iWorst = 0
    if not oArgs.macro:
        oArgs.macro = defaultMacros()
        sayLine("No macro was named, so %s found in scripts\\tests %s taken%s." % (countNoun(len(oArgs.macro), "macro"), "is" if len(oArgs.macro) == 1 else "are", "" if oArgs.macro else "; there is nothing to play"))
        if not oArgs.macro: return 2
    for sMacro in oArgs.macro:
        if not os.path.isfile(sMacro):
            sayLine("No macro at %s." % sMacro); iWorst = max(iWorst, 2); continue
        iWorst = max(iWorst, play(sMacro, bCheck=oArgs.check, bHear=not oArgs.no_hear, bKeep=oArgs.keep))
    return iWorst


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception:
        logLine("ERROR " + traceback.format_exc())
        print("runMacro stopped on an error; the log says why: " + sLogPath)
        sys.exit(2)
