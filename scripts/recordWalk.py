r"""recordWalk.py -- drive a Homer app through a walk, record what the screen reader says, and compare.

WHY THIS EXISTS (3 November 2026)

uiCheck drives a program through Windows UI Automation and checks what appears, and says plainly that it
does not hear speech. The tutorials' Hear lines are what a screen reader should say, written by hand and
marked "confirm with a live run". This script is that live run: it presses each step's keys in the real
program while JAWS (or NVDA) runs as usual, records the computer's own sound, writes it down with a local
speech recognizer, and sets beside it, step by step, what the walk expected, what UI Automation says has
focus, and what the app itself said in its speech log.

WHAT IT NEEDS. Python 3 and the internet the first time, to install pywinauto, comtypes, soundcard, numpy
and faster-whisper with pip, for this user only, and to download the speech model (about 500 MB for small.en,
once). No administrator rights. The app must be running, or installed where Windows can start it.

WHILE IT RUNS, keep your hands off the keyboard: it presses keys in the app. To stop it, switch to any
other window -- before every step it checks that the app is still in front, and stops if not.

THE WALK. The same format as a tutorial walk: [step] sections with Key=, Hear= and Say= lines, plus two of
its own, both optional:

    Type=seahawks       text typed before the Key, letter by letter
    Wait=6              seconds to listen after the key, instead of the default
    Press=2             the key pressed twice in quick succession, a double press

recordWalk menus walks every menu item and the Alternate Menu without a walk file; --menus and --items set how far.

USAGE

    recordWalk                          WITH NO ARGUMENTS, THE SENSIBLE WHOLE: the pilot walk, each window scanned once with
                                        Axe.Windows, and then Tutorial 1 performed live -- the menus through the real screen
                                        reader -- when its recording is missing or older than the tutorial or the program
    recordWalk --walk C:\path\walk.inix another walk
    recordWalk --steps 3                only the first three steps
    recordWalk --wait 5                 listen five seconds after each key (default 4)
    recordWalk --model base.en          a smaller, faster, less accurate recognizer (default small.en)
    recordWalk --no-transcribe          record and report, and write nothing down
    recordWalk --no-axe                 leave out the scans
    recordWalk 1                        the same as recordWalk --tutorial --walk 1: Tutorial 1 performed live
    recordWalk --tutorial --walk 1       (or the tutorial's path; a number finds help\Tutorial_1_*.inix)
                                        perform a tutorial live: the host's lines in one Windows voice, the real screen
                                        reader recorded at each key, joined into one recording, <walk>-live.wav and .mp3
    recordWalk --keep                   leave the app open at the end; by default it is closed, so its closing says the run is over
    recordWalk --speaker Headphones     record the sound output whose name holds Headphones (the log names them all)
    recordWalk --app DbDo               the app, when the folder does not say

Writes logs\<App>-recordWalk-<stamp>.log, logs\<App>-recordWalk-<stamp>.md (the report), and one .wav per
step under results\recordWalk-<stamp>. Exit code 0 when every step was heard as expected, 1 otherwise.
"""

import argparse, ctypes, datetime, difflib, glob, json, os, platform, re, shutil, subprocess, sys, time, traceback, urllib.request, wave, zipfile

iDefaultWait, iSampleRate, nMatchNeeded = 4, 16000, 0.6
# Listening (3 November 2026): chunks of a quarter second; speech is a chunk louder than 0.01 (JAWS peaked 0.13 to 0.22);
# half a second of quiet before a key, 0.8 after speech ends, at most ten seconds to settle and fifteen to listen.
nChunkSeconds, nMaxListen, nMaxSettle, nQuietAfter, nQuietBefore, nSpeechLevel = 0.25, 15, 10, 0.8, 0.5, 0.01
sDefaultModel = "small.en"

sStamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
sScriptDir = os.path.dirname(os.path.abspath(__file__))
# THE APP'S FOLDER (5 November 2026): recordWalk lives in the kit's scripts, and is run from the app's folder; that
# folder, when it has a scripts folder of its own, is the app's. Run from the app's own scripts, the folder above it.
sProject = os.getcwd() if (os.path.isdir(os.path.join(os.getcwd(), "scripts"))
                           and os.path.normcase(os.path.abspath(os.getcwd())) != os.path.normcase(os.path.dirname(sScriptDir))) \
    else os.path.dirname(sScriptDir)

sLogPath = ""
sStartExe = ""
lsReport = []


def logLine(sText):
    """Writes one line to the detailed log."""
    if not sLogPath: return False
    with open(sLogPath, "a", encoding="utf-8") as fLog:
        fLog.write(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3] + "  " + sText + "\n")
    return True


def sayLine(sText=""):
    """Writes a line to the console, which JAWS reads, and to the log."""
    print(sText, flush=True)
    logLine("CONSOLE: " + sText)
    return True


# ---- the menus, read and checked (10 November 2026) ----
#
# WHY. recordWalk menus pressed Down Arrow 25 times in every menu, so a short menu wrapped and repeated, a submenu was
# never opened, and nothing was checked: the report held only what the recognizer made of JAWS's speech. Now the
# menus are read first through UI Automation -- each menu's items, a submenu's items, each item's access letter and
# key -- and the walk is made from what is there: Down Arrow once per item, Right Arrow into a submenu and Left Arrow
# out, Escape and Right Arrow to the next menu. Homer's menu rules are checked on what was read, and the findings
# lead the report.
lsMenuAudit = []

def uiaKeyOf(oElement, sProperty):
    """An element's AccessKey or AcceleratorKey, as UI Automation reports it; "" when it has none."""
    try: return (getattr(oElement.element_info.element, "Current" + sProperty) or "").strip()
    except Exception: return ""

def letterOf(sKey):
    """The letter of a key such as "Alt+F" or "Control+O"; "" when the key is not a letter."""
    sLast = re.split(r"[+ ]", sKey.strip())[-1] if sKey.strip() else ""
    return sLast.upper() if len(sLast) == 1 and sLast.isalpha() else ""

def hasSubmenu(oItem):
    try: return oItem.iface_expand_collapse.CurrentExpandCollapseState != 3   # 3 is a leaf, with nothing to open
    except Exception: return False

def menuItemsOf(oItem, iDepth):
    """The items of the menu oItem opens, a submenu's items one level down; oItem is closed again after."""
    lItems = []
    try:
        oItem.expand(); time.sleep(0.5)
    except Exception as exceptionError:
        logLine("menu audit: " + oItem.window_text() + " did not open: " + str(exceptionError)); return lItems
    lKids = oItem.children(control_type="MenuItem")
    if not lKids:
        # Windows Forms may put an open menu under the desktop, as its own Menu window.
        try:
            from pywinauto import Desktop
            for oMenu in Desktop(backend="uia").windows(control_type="Menu"):
                lKids = oMenu.children(control_type="MenuItem")
                if lKids: break
        except Exception as exceptionError: logLine("menu audit: no open menu found: " + str(exceptionError))
    for oKid in lKids:
        dItem = {"name": oKid.window_text().strip(), "access": letterOf(uiaKeyOf(oKid, "AccessKey")), "key": uiaKeyOf(oKid, "AcceleratorKey"),
                 "enabled": oKid.is_enabled(), "items": []}
        if iDepth < 1 and hasSubmenu(oKid): dItem["items"] = menuItemsOf(oKid, iDepth + 1)
        lItems.append(dItem)
    try: oItem.collapse()
    except Exception: pass
    return lItems

def menuTreeFromCode(sRoot):
    """The app's menus as its code builds them (10 November 2026): each menu's items in order, a submenu's items one
    level down, hidden items left out, and each item's access letter and key. UI Automation found Windows Forms
    menus but none of their items, twice, so the walk fell back to pressing Down Arrow 25 times, and short menus
    wrapped and repeated; the code is the same source the release gate checks. [] when no menu is found."""
    sCode = ""
    for sPath in sorted(glob.glob(os.path.join(sRoot, "*.cs"))):
        try:
            with open(sPath, encoding="utf-8-sig", errors="replace") as fIn: sText = fIn.read()
        except Exception: continue
        sCode += "\n".join(s for s in sText.splitlines() if not s.lstrip().startswith("//")) + "\n"
    def letterOfCaption(sCaption):
        oMatch = re.search(r"&([^&])", sCaption)
        return oMatch.group(1).upper() if oMatch else ""
    def plain(sCaption): return sCaption.replace("&", "").replace("...", "").strip()
    setHidden = set(re.findall(r"\b(\w+)\.Visible\s*=\s*false\s*;", sCode))
    lMenus = []
    lTops = re.findall(r'(mi[A-Za-z]+)\s*=\s*addMenu\("([^"]*)"\)', sCode)
    if lTops:
        # DbDo: addMenu for each menu, addItem(parent, ...) for each item, parent.DropDownItems.Add(sub) for a submenu,
        # all in the order the code runs them.
        dSub = dict(re.findall(r'ToolStripMenuItem\s+(mi[A-Za-z]+)\s*=\s*new\s+\w*MenuItem\("([^"]*)"\)', sCode))
        dChildren = {}
        oScan = re.compile(r'(?:(\w+)\s*=\s*)?addItem\((mi[A-Za-z]+),\s*"([^"]*)",\s*"[^"]*",\s*([^,]+?),|(mi[A-Za-z]+)\.DropDownItems\.Add\((mi[A-Za-z]+)\)')
        for oMatch in oScan.finditer(sCode):
            if oMatch.group(2):
                sVar, sParent, sCaption, sKeys = oMatch.group(1), oMatch.group(2), oMatch.group(3), oMatch.group(4)
                if sVar and sVar in setHidden: continue
                dChildren.setdefault(sParent, []).append({"name": plain(sCaption), "access": letterOfCaption(sCaption),
                    "key": sKeys.replace("Keys.", "").replace(" | ", "+").replace("|", "+").strip(), "enabled": True, "items": [], "var": sVar or ""})
            elif oMatch.group(5) and oMatch.group(6) in dSub:
                sSub = oMatch.group(6)
                dChildren.setdefault(oMatch.group(5), []).append({"name": plain(dSub[sSub]), "access": letterOfCaption(dSub[sSub]),
                    "key": "", "enabled": True, "items": [], "var": sSub})
        for dList in dChildren.values():
            for dItem in dList:
                if dItem["var"] in dChildren and not dItem["items"]: dItem["items"] = dChildren[dItem["var"]]
        for sVar, sCaption in lTops:
            lMenus.append({"name": plain(sCaption), "access": letterOfCaption(sCaption), "items": dChildren.get(sVar, [])})
        return lMenus
    # EdSharp and FileDir: each item assigned to a variable, then each menu's items listed in order with AddRange.
    dItems = {}
    for sVar, sCaption, sRest in re.findall(r'(\w+)\s*=\s*(?:CreateMenuItem|menu_Helper|CreateMenu)\("([^"]*)"(.*?)\)\s*;', sCode):
        oKey = re.match(r',\s*"([^"]*)"', sRest)
        dItems[sVar] = {"name": plain(sCaption), "access": letterOfCaption(sCaption), "key": oKey.group(1) if oKey else "", "enabled": True, "items": []}
    for sVar, sCaption in re.findall(r'(\w+)\s*=\s*new\s+\w*MenuItem\("([^"]*)"', sCode):
        dItems.setdefault(sVar, {"name": plain(sCaption), "access": letterOfCaption(sCaption), "key": "", "enabled": True, "items": []})
    dLists = {}
    for sParent, sList in re.findall(r'(\w+)\.DropDownItems\.AddRange\(new ToolStrip(?:Menu)?Item\[\]\s*\{([^}]*)\}', sCode, re.S):
        dLists[sParent] = [s for s in re.findall(r"\w+", sList) if s in dItems and s not in setHidden]
    def build(sVar, iDepth):
        dItem = dict(dItems[sVar])
        if iDepth < 2 and sVar in dLists: dItem["items"] = [build(s, iDepth + 1) for s in dLists[sVar]]
        return dItem
    oTops = re.search(r'\w+\.Items\.AddRange\(new ToolStrip(?:Menu)?Item\[\]\s*\{([^}]*)\}', sCode, re.S)
    if not oTops: return []
    for sTop in re.findall(r"\w+", oTops.group(1)):
        if sTop in dItems:
            dMenu = build(sTop, 0)
            lMenus.append({"name": dMenu["name"], "access": dMenu["access"], "items": dMenu["items"]})
    return lMenus


def openMenuItems(sPid, setSeen):
    """The items of the menu that has just opened: a visible window of the app's process not in setSeen. Each item is
    {name, access, key, enabled, sub}, sub True when it opens a submenu."""
    from pywinauto import Desktop
    for iTry in range(6):
        time.sleep(0.4)
        for oWin in Desktop(backend="uia").windows(process=sPid):
            if oWin.handle in setSeen: continue
            try: lKids = oWin.descendants(control_type="MenuItem")
            except Exception: lKids = []
            if lKids:
                setSeen.add(oWin.handle)
                return [{"name": o.window_text().strip(), "access": letterOf(uiaKeyOf(o, "AccessKey")), "key": uiaKeyOf(o, "AcceleratorKey"),
                         "enabled": o.is_enabled(), "sub": hasSubmenu(o), "items": []} for o in lKids]
    return []

def readMenus(sApp):
    """The app's menus, read by opening each with the keyboard (10 November 2026): Windows Forms menus did not open
    for UI Automation's expand, so each menu is opened as a person opens it -- F10, Right Arrow to it, Down Arrow --
    and read from the window that appears; a submenu is opened with Right Arrow and read the same way."""
    try:
        from pywinauto.application import Application
        from pywinauto.keyboard import send_keys
        lWindows = windowsOf(sApp + ".exe")
        hWnd = lWindows[0][0]
        oApp = Application(backend="uia").connect(handle=hWnd)
        sPid = oApp.process
        oWin = oApp.window(handle=hWnd)
        lBars = []
        for oChild in oWin.children():
            try:
                if oChild.element_info.control_type in ("MenuBar", "ToolBar") and oChild.children(control_type="MenuItem"): lBars.append(oChild)
            except Exception: pass
        if not lBars: logLine("menu reader: no menu bar among the window's children"); return []
        oBar = max(lBars, key=lambda o: len(o.children(control_type="MenuItem")))
        lMenus = [{"name": o.window_text().strip(), "access": letterOf(uiaKeyOf(o, "AccessKey")), "items": []} for o in oBar.children(control_type="MenuItem")]
        for iMenu, dMenu in enumerate(lMenus):
            oWin.set_focus(); time.sleep(0.3)
            setSeen = set(o.handle for o in __import__("pywinauto").Desktop(backend="uia").windows(process=sPid))
            send_keys("{F10}" + "{RIGHT}" * iMenu + "{DOWN}")
            dMenu["items"] = openMenuItems(sPid, setSeen)
            iAt = 0
            for iItem, dItem in enumerate(dMenu["items"]):
                if not dItem["sub"]: continue
                send_keys("{DOWN}" * (iItem - iAt)); iAt = iItem
                send_keys("{RIGHT}")
                dItem["items"] = [dict(d, sub=False) for d in openMenuItems(sPid, setSeen)]
                send_keys("{ESC}"); time.sleep(0.2)
                if not dItem["items"]: dItem["sub"] = False
            send_keys("{ESC}{ESC}{ESC}"); time.sleep(0.3)
            logLine("menu reader: %s, %d items" % (dMenu["name"], len(dMenu["items"])))
        return lMenus
    except Exception:
        logLine("menu reader: the menus could not be read: " + traceback.format_exc()); return []

def checkMenus(lMenus):
    """Homer's menu rules, checked on what was read: each line a finding; [] when every rule holds. The same rules
    check.py's release gate applies (10 November 2026): no letter shared in one menu; a letter starts a word of the
    name and is the key's letter, Shift or not; no letter on a function, arrow or symbol key; a key's letter starts a
    word. An item with no letter is not a finding: a function key's item, or one whose key letter another item has,
    takes none."""
    lFindings = []
    def words(sName): return re.findall(r"[A-Za-z][A-Za-z0-9']*", sName)
    def mainKey(sKey):
        lParts = [s for s in re.split(r"[+ ]", sKey.split(" or ")[0].strip()) if s]
        lMain = [s for s in lParts if s not in ("Alt", "Control", "Shift", "Ctrl")]
        return lMain[-1] if lMain else ""
    def exempt(sLetter, sName):
        return (sLetter == "X" and re.search(r"ex", sName, re.I)) or (sLetter == "V" and re.search(r"voke", sName, re.I)) \
            or (sLetter == "Z" and re.search(r"toggle|status", sName, re.I)) \
            or (sLetter in "XCVZY" and re.search(r"cut|copy|paste|undo|redo|repeat", sName, re.I))
    def checkLevel(sWhere, lItems):
        dByLetter = {}
        for d in lItems:
            sName = d["name"]; sLetter = d.get("access", ""); sMain = mainKey(d.get("key", ""))
            bLetterKey = len(sMain) == 1 and sMain.isalpha()
            bSymbolKey = sMain not in ("", "None") and not (len(sMain) == 1 and sMain.isalnum()) and not re.match(r"^D\d$", sMain)
            lInitials = [w[0].upper() for w in words(sName)]
            if sLetter:
                dByLetter.setdefault(sLetter, []).append(sName)
                if sLetter not in lInitials and not exempt(sLetter, sName):
                    lFindings.append("%s: %s has access letter %s, which starts none of its words" % (sWhere, sName, sLetter))
                if bLetterKey and sMain.upper() != sLetter and not exempt(sMain.upper(), sName):
                    lFindings.append("%s: %s has access letter %s but key %s" % (sWhere, sName, sLetter, d["key"]))
                if bSymbolKey:
                    lFindings.append("%s: %s has access letter %s, but its key %s is not a letter" % (sWhere, sName, sLetter, d["key"]))
            if d.get("items"): checkLevel(sWhere + " > " + sName, d["items"])
        for sLetter, lNames in sorted(dByLetter.items()):
            if len(lNames) > 1: lFindings.append("%s: access letter %s is shared by %s" % (sWhere, sLetter, "; ".join(lNames)))
    checkLevel("menu bar", [{"name": d["name"], "access": d["access"], "key": "", "items": []} for d in lMenus])
    for dMenu in lMenus: checkLevel(dMenu["name"], dMenu["items"])
    return lFindings

def menuWalkSteps(lMenus, iAlternate):
    """The walk's lines for what was read: each item heard by name, submenus entered and left."""
    lLines = []
    def step(sKey, sSay, sHear="", iWait=2):
        lLines.extend(["[step]", "Say=" + sSay, "Key=" + sKey, "Wait=" + str(iWait)] + (["Hear=" + sHear] if sHear else []) + [""])
    step("F10", "F10, the menu bar.", lMenus[0]["name"] if lMenus else "", 3)
    for iMenu, dMenu in enumerate(lMenus):
        if iMenu > 0: step("Right", "The " + dMenu["name"] + " menu.", dMenu["name"])
        for iItem, dItem in enumerate(dMenu["items"]):
            step("Down", "%s menu, %s." % (dMenu["name"], dItem["name"]), dItem["name"], 3 if iItem == 0 else 2)
            if dItem["items"]:
                for iSub, dSub in enumerate(dItem["items"]):
                    step("Right" if iSub == 0 else "Down", "%s submenu, %s." % (dItem["name"], dSub["name"]), dSub["name"])
                step("Left", "Back to %s." % dItem["name"], dItem["name"])
        step("Escape", "The " + dMenu["name"] + " menu closed.", dMenu["name"])
    step("Escape", "Leaving the menu bar.")
    step("Alt+F10", "Alt+F10, the Alternate Menu.", "Alternate Menu", 4)
    for iItem in range(2, max(2, iAlternate) + 1): step("Down", "Alternate Menu, command %d." % iItem)
    step("End", "Alternate Menu, the last command.")
    step("Home", "Alternate Menu, the first command.")
    step("Escape", "The Alternate Menu closed.")
    return lLines


def ensurePackages():
    """Installs, for this user only, the packages the run needs, when they are missing. Returns True when all import."""
    # SOUNDCARD FIRST (3 November 2026). pywinauto starts COM in multithreaded mode as it loads; soundcard then
    # starts COM itself, is told "already started" (S_FALSE, 1) and treats that success as an error -- "Error
    # 0x100000001" on all three of the first runs. Loaded first, soundcard starts COM, and pywinauto and comtypes
    # accept the apartment they find.
    dNeeded = {"soundcard": "soundcard", "numpy": "numpy", "pywinauto": "pywinauto", "comtypes": "comtypes"}
    lMissing = []
    for sModule, sPackage in dNeeded.items():
        try:
            __import__(sModule)
        except ImportError:
            lMissing.append(sPackage)
    if lMissing:
        sayLine("Installing " + ", ".join(lMissing) + " for this user. This happens once.")
        iCode = subprocess.call([sys.executable, "-m", "pip", "install", "--user", "--quiet"] + lMissing)
        logLine("pip install " + " ".join(lMissing) + " exit " + str(iCode))
    try:
        import soundcard, numpy, pywinauto, comtypes
        return True
    except ImportError as exceptionError:
        sayLine("A package could not be installed: " + str(exceptionError) + ". The log has pip's answer.")
        return False


def ensureWhisper():
    """faster-whisper, installed when missing; None when it cannot be."""
    try:
        from faster_whisper import WhisperModel
        return WhisperModel
    except ImportError:
        pass
    sayLine("Installing faster-whisper, the local speech recognizer. This happens once.")
    iCode = subprocess.call([sys.executable, "-m", "pip", "install", "--user", "--quiet", "faster-whisper"])
    logLine("pip install faster-whisper exit " + str(iCode))
    try:
        from faster_whisper import WhisperModel
        return WhisperModel
    except ImportError as exceptionError:
        logLine("faster-whisper still does not import: " + str(exceptionError))
        return None


def findWhisperCpp():
    """HomerScribe's whisper.cpp, when it is installed: (program, model) or ("", ""). The fallback when
    faster-whisper cannot be installed, which happens on a Python too new for its engine."""
    lRoots = [os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), "HomerScribe"),
              os.path.join(os.environ.get("LOCALAPPDATA", ""), "HomerScribe")]
    sExe, sModel = "", ""
    for sRoot in lRoots:
        for sDir, _lDirs, lFiles in os.walk(sRoot):
            for sFile in lFiles:
                sLower = sFile.lower()
                if not sExe and sLower in ("whisper-cli.exe", "main.exe", "whisper.exe"): sExe = os.path.join(sDir, sFile)
                if sLower.startswith("ggml-") and sLower.endswith(".bin"):
                    if not sModel or ".en" in sLower: sModel = os.path.join(sDir, sFile)
    logLine("whisper.cpp: program " + (sExe or "not found") + ", model " + (sModel or "not found"))
    return (sExe, sModel) if sExe and sModel else ("", "")


def transcribeCpp(sExe, sModel, sWave):
    """Writes a recording down with whisper.cpp."""
    sOut = os.path.splitext(sWave)[0]
    lCommand = [sExe, "-m", sModel, "-f", sWave, "-l", "en", "-nt", "-otxt", "-of", sOut]
    oRun = subprocess.run(lCommand, capture_output=True, text=True, timeout=300)
    logLine("whisper.cpp exit " + str(oRun.returncode) + " for " + os.path.basename(sWave))
    if os.path.isfile(sOut + ".txt"):
        with open(sOut + ".txt", "r", encoding="utf-8", errors="replace") as fIn: return " ".join(fIn.read().split())
    return ""


# ---- the walk ----

def readWalk(sPath, bAll=False):
    """The walk's steps: a list of dictionaries with sKey, sType, iWait, sSay and lHear. A walk keeps only the steps
    that press or type something; a tutorial (bAll) keeps its narration too."""
    lSteps, dStep = [], None
    with open(sPath, "r", encoding="utf-8-sig") as fWalk:
        for sRaw in fWalk:
            sLine = sRaw.strip()
            if sLine.startswith(";") or not sLine: continue
            if sLine.lower() == "[step]":
                dStep = {"sKey": "", "sType": "", "iWait": 0, "iPress": 1, "sSay": "", "lHear": [], "lWindow": [], "lFocus": [], "nPause": 0.0}
                lSteps.append(dStep)
                continue
            if sLine.startswith("["):
                dStep = None
                continue
            if dStep is None or "=" not in sLine: continue
            sName, sValue = [s.strip() for s in sLine.split("=", 1)]
            sName = sName.lower()
            if sName == "key": dStep["sKey"] = sValue
            elif sName == "type": dStep["sType"] = sValue
            elif sName == "say": dStep["sSay"] = sValue
            elif sName == "hear" and sValue: dStep["lHear"].append(sValue)
            # WHAT WINDOWS SHOWS, BESIDE WHAT WAS HEARD (17 November 2026, for runMacro): Window= is text the title of
            # the window in front must hold after the step, Focus= text the focused control's name, role or value
            # must hold, and Pause= seconds to wait before the key, for a macro's Wait line.
            elif sName == "window" and sValue: dStep["lWindow"].append(sValue)
            elif sName == "focus" and sValue: dStep["lFocus"].append(sValue)
            elif sName == "pause":
                try: dStep["nPause"] = max(0.0, min(600.0, float(sValue)))
                except ValueError: logLine("Pause is not a number of seconds: " + sValue)
            elif sName == "press":
                # PRESS=2 (9 November 2026): the key pressed twice in quick succession, as a person double-presses a
                # Homer key to have its second meaning -- spell a value, or show a summary in a window.
                try: dStep["iPress"] = max(1, min(2, int(sValue)))
                except ValueError: logLine("Press is not 1 or 2: " + sValue)
            elif sName == "wait":
                try: dStep["iWait"] = int(sValue)
                except ValueError: logLine("Wait is not a whole number: " + sValue)
    return lSteps if bAll else [d for d in lSteps if d["sKey"] or d["sType"]]


# ---- A TUTORIAL PERFORMED LIVE (5 November 2026) ----------------------------------------------------------------------
# --tutorial reads a tutorial's [step] sections, narration and all. Before each key, the host's Say line is spoken by a
# Windows voice into a file -- never aloud, so it cannot spill into the recording of the screen reader -- then the key
# is pressed and the real reader recorded, as in any walk. A step with no key but a Hear line, where the tutorial has a
# second voice play the screen reader reading, gets a second Windows voice. Every piece is 16 kHz, 16-bit mono, the
# recording's own format, so they join end to end: <walk>-live.wav, and an .mp3 beside it when ffmpeg can be found.
def speechFrames(sText, iVoice):
    """sText spoken by the iVoice'th enabled Windows voice (the first when there are fewer), as 16-bit samples."""
    import numpy, tempfile
    sPath = os.path.join(tempfile.gettempdir(), "recordWalk-say.wav")
    sScript = ("Add-Type -AssemblyName System.Speech; $o = New-Object System.Speech.Synthesis.SpeechSynthesizer; "
               "$l = @($o.GetInstalledVoices() | Where-Object { $_.Enabled }); if ($l.Count -gt %d) { $o.SelectVoice($l[%d].VoiceInfo.Name) }; "
               "$f = New-Object System.Speech.AudioFormat.SpeechAudioFormatInfo(%d, [System.Speech.AudioFormat.AudioBitsPerSample]::Sixteen, "
               "[System.Speech.AudioFormat.AudioChannel]::Mono); $o.SetOutputToWaveFile($env:RW_WAV, $f); $o.Speak($env:RW_TEXT); $o.Dispose()") % (iVoice, iVoice, iSampleRate)
    oRun = subprocess.run(["powershell", "-NoProfile", "-Command", sScript], env=dict(os.environ, RW_WAV=sPath, RW_TEXT=sText),
                          capture_output=True, text=True, timeout=120)
    if oRun.returncode != 0: logLine("voice exit " + str(oRun.returncode) + ": " + (oRun.stderr or "").strip()[:500])
    with wave.open(sPath, "rb") as oWave:
        return numpy.frombuffer(oWave.readframes(oWave.getnframes()), dtype=numpy.int16)


def launchKey(sProject, sApp):
    """The app's own launch hotkey, from #define HotKey in its installer script, as a set of lower-case key names."""
    sIss = os.path.join(sProject, sApp + "_setup.iss")
    try:
        with open(sIss, encoding="utf-8-sig", errors="replace") as fIn:
            oMatch = re.search(r'^#define\s+HotKey\s+"([^"]+)"', fIn.read(), re.M)
        return keySet(oMatch.group(1)) if oMatch else set()
    except Exception:
        return set()


def keySet(sKey):
    dSame = {"ctrl": "control", "ctl": "control"}
    return set(dSame.get(s.strip().lower(), s.strip().lower()) for s in (sKey or "").split("+") if s.strip())


def framesToInt16(aFrames):
    import numpy
    aMono = aFrames.mean(axis=1) if aFrames.ndim > 1 else aFrames
    return (numpy.clip(aMono, -1.0, 1.0) * 32767).astype(numpy.int16)


def writeLive(sWaves, sWalk, lPieces):
    """The pieces joined, with a short pause between, as <walk>-live.wav; an .mp3 beside it when ffmpeg is found."""
    import numpy
    aGap = numpy.zeros(int(iSampleRate * 0.4), dtype=numpy.int16)
    lJoined = []
    for aPiece in lPieces: lJoined += [aPiece, aGap]
    sBase = os.path.join(sWaves, os.path.splitext(os.path.basename(sWalk))[0] + "-live")
    with wave.open(sBase + ".wav", "wb") as oWave:
        oWave.setnchannels(1); oWave.setsampwidth(2); oWave.setframerate(iSampleRate)
        oWave.writeframes(numpy.concatenate(lJoined).tobytes() if lJoined else b"")
    sFfmpeg = shutil.which("ffmpeg") or ""
    for sKit in (os.environ.get("homerDev", ""), r"C:\HomerDev"):
        if not sFfmpeg and sKit and os.path.isfile(os.path.join(sKit, "exec", "ffmpeg.exe")): sFfmpeg = os.path.join(sKit, "exec", "ffmpeg.exe")
    if sFfmpeg:
        oRun = subprocess.run([sFfmpeg, "-y", "-loglevel", "error", "-i", sBase + ".wav", "-codec:a", "libmp3lame", "-q:a", "4", sBase + ".mp3"], capture_output=True, text=True, timeout=600)
        logLine("ffmpeg exit " + str(oRun.returncode) + " " + (oRun.stderr or "").strip()[:500])
        if oRun.returncode == 0: return sBase + ".mp3"
    return sBase + ".wav"


dNamedKeys = {"enter": "{ENTER}", "escape": "{ESC}", "esc": "{ESC}", "tab": "{TAB}", "space": "{SPACE}",
              "backspace": "{BACKSPACE}", "delete": "{DELETE}", "insert": "{INSERT}", "home": "{HOME}", "end": "{END}",
              "page up": "{PGUP}", "page down": "{PGDN}", "up arrow": "{UP}", "down arrow": "{DOWN}",
              "left arrow": "{LEFT}", "right arrow": "{RIGHT}", "up": "{UP}", "down": "{DOWN}", "left": "{LEFT}",
              "right": "{RIGHT}", "apps": "{APPS}", "applications": "{APPS}",
              # HOMER'S OWN KEY NAMES (17 November 2026): a walk or macro written in the suite's names -- UpArrow,
              # PageDown, NumPadStar -- was typed as text, since only "up arrow" and "page down" were known.
              "uparrow": "{UP}", "downarrow": "{DOWN}", "leftarrow": "{LEFT}", "rightarrow": "{RIGHT}",
              "pageup": "{PGUP}", "pagedown": "{PGDN}", "numpadstar": "{VK_MULTIPLY}", "numpadslash": "{VK_DIVIDE}",
              "numpadplus": "{VK_ADD}", "numpaddash": "{VK_SUBTRACT}", "numpaddelete": "{VK_DECIMAL}",
              "numpad0": "{VK_NUMPAD0}", "numpad1": "{VK_NUMPAD1}", "numpad2": "{VK_NUMPAD2}", "numpad3": "{VK_NUMPAD3}",
              "numpad4": "{VK_NUMPAD4}", "numpad5": "{VK_NUMPAD5}", "numpad6": "{VK_NUMPAD6}", "numpad7": "{VK_NUMPAD7}",
              "numpad8": "{VK_NUMPAD8}", "numpad9": "{VK_NUMPAD9}"}
# The punctuation keys by Homer's names (help\KeyNames.md), as the character the key types unshifted.
dCharKeys = {"apostrophe": "'", "backslash": "\\", "comma": ",", "dash": "-", "equals": "=", "graveaccent": "`",
             "leftbracket": "[", "period": ".", "rightbracket": "]", "semicolon": ";", "slash": "/"}
dModifiers = {"control": "^", "ctrl": "^", "alt": "%", "shift": "+"}


def keysFor(sKey):
    """A walk's key name, such as Alt+Shift+F1 or Down Arrow, in pywinauto's send_keys syntax.
    Insert, the JAWS key, is held down around the rest, as JAWS needs it."""
    lParts = [s.strip() for s in sKey.split("+") if s.strip()]
    if not lParts: return ""
    sPrefix, bInsert = "", False
    for sPart in lParts[:-1]:
        sLower = sPart.lower()
        if sLower in dModifiers: sPrefix += dModifiers[sLower]
        elif sLower in ("insert", "jaws", "jawskey", "nvda"): bInsert = True
        else: logLine("unknown modifier in key " + repr(sKey) + ": " + sPart)
    sLast = lParts[-1]
    sLower = sLast.lower()
    if sLower in dNamedKeys: sMain = dNamedKeys[sLower]
    elif sLower in dCharKeys: sMain = escapeText(dCharKeys[sLower])
    elif re.match(r"^f([1-9]|1[0-2])$", sLower): sMain = "{" + sLower.upper() + "}"
    elif len(sLast) == 1: sMain = escapeText(sLast.lower())
    else: sMain = escapeText(sLast)
    sKeys = sPrefix + sMain if len(sMain) == 1 or sMain.startswith("{") else sPrefix + "(" + sMain + ")"
    if bInsert: sKeys = "{INSERT down}" + sKeys + "{INSERT up}"
    return sKeys


def escapeText(sText):
    """Text typed letter by letter: send_keys' own characters wrapped in braces."""
    return "".join("{" + c + "}" if c in "+^%~(){}[]" else c for c in sText)


# ---- the app and its window ----

def exeNameOf(iPid):
    """The file name of the program a process runs, or "" when it cannot be read (from uiCheck)."""
    from ctypes import wintypes
    hProcess = ctypes.windll.kernel32.OpenProcess(0x1000, False, iPid)
    if not hProcess: return ""
    try:
        oBuffer = ctypes.create_unicode_buffer(1024)
        iSize = wintypes.DWORD(1024)
        if ctypes.windll.kernel32.QueryFullProcessImageNameW(hProcess, 0, oBuffer, ctypes.byref(iSize)):
            return os.path.basename(oBuffer.value)
        return ""
    finally:
        ctypes.windll.kernel32.CloseHandle(hProcess)


def exePathOf(iPid):
    """The full path of the program a process runs, or "" when it cannot be read."""
    from ctypes import wintypes
    hProcess = ctypes.windll.kernel32.OpenProcess(0x1000, False, iPid)
    if not hProcess: return ""
    try:
        oBuffer = ctypes.create_unicode_buffer(1024)
        iSize = wintypes.DWORD(1024)
        if ctypes.windll.kernel32.QueryFullProcessImageNameW(hProcess, 0, oBuffer, ctypes.byref(iSize)):
            return oBuffer.value
        return ""
    finally:
        ctypes.windll.kernel32.CloseHandle(hProcess)


def appIdentity(sExe):
    """WHICH BUILD WAS WALKED (5 November 2026): twice a walk's misses could only be explained by guessing that an older
    build had run, from the wording heard. The report now says, under its title, the program's path, its file version
    and when that file was built, so a run is matched to its build without guessing."""
    try:
        lWindows = windowsOf(sExe)
        if not lWindows: return "The app's program could not be identified: no window of " + sExe + " was found."
        sPath = exePathOf(lWindows[0][1])
        if not sPath: return "The app's program could not be identified: its path could not be read."
        sVersion = ""
        try:
            oRun = subprocess.run(["powershell", "-NoProfile", "-Command", "(Get-Item -LiteralPath '" + sPath.replace("'", "''") + "').VersionInfo.FileVersion"],
                                  capture_output=True, text=True, timeout=30)
            sVersion = (oRun.stdout or "").strip()
        except Exception: pass
        sBuilt = datetime.datetime.fromtimestamp(os.path.getmtime(sPath)).strftime("%d %B %Y at %H:%M")
        return "The app: " + sPath + (", version " + sVersion if sVersion else "") + ", built " + sBuilt + "."
    except Exception as oError:
        return "The app's program could not be identified: " + str(oError)


def foregroundExe():
    """The program whose window is in front, and that window's title."""
    from ctypes import wintypes
    hWnd = ctypes.windll.user32.GetForegroundWindow()
    iPid = wintypes.DWORD()
    ctypes.windll.user32.GetWindowThreadProcessId(hWnd, ctypes.byref(iPid))
    iLength = ctypes.windll.user32.GetWindowTextLengthW(hWnd)
    oTitle = ctypes.create_unicode_buffer(iLength + 1)
    ctypes.windll.user32.GetWindowTextW(hWnd, oTitle, iLength + 1)
    return exeNameOf(iPid.value), oTitle.value


def windowsOf(sExe):
    """Visible top-level windows with a title, of any process running sExe: [(handle, pid, title)] (from uiCheck)."""
    from ctypes import wintypes
    lFound = []
    fnProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    def onWindow(hWnd, lParam):
        if not ctypes.windll.user32.IsWindowVisible(hWnd): return True
        iLength = ctypes.windll.user32.GetWindowTextLengthW(hWnd)
        if iLength == 0: return True
        oTitle = ctypes.create_unicode_buffer(iLength + 1)
        ctypes.windll.user32.GetWindowTextW(hWnd, oTitle, iLength + 1)
        iPid = wintypes.DWORD()
        ctypes.windll.user32.GetWindowThreadProcessId(hWnd, ctypes.byref(iPid))
        if exeNameOf(iPid.value).lower() == sExe.lower(): lFound.append((hWnd, iPid.value, oTitle.value))
        return True
    ctypes.windll.user32.EnumWindows(fnProc(onWindow), 0)
    return lFound


def bringForward(sApp):
    """The app's main window in front: started first when it is not running. True when it is in front."""
    sExe = sApp + ".exe"
    if not windowsOf(sExe):
        # --exe FIRST (17 November 2026): a live test in tidy tests the build just made, not the installed copy.
        lPlaces = ([sStartExe] if sStartExe else []) + [os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), sApp, "exec", sExe),
                   os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), sApp, sExe),
                   os.path.join(sProject, "exec", sExe)]
        sFound = next((s for s in lPlaces if os.path.isfile(s)), "")
        if not sFound:
            sayLine(sApp + " is not running, and " + sExe + " was not found to start it.")
            return False
        sayLine("Starting " + sApp + ".")
        logLine("starting " + sFound)
        subprocess.Popen([sFound], cwd=os.path.dirname(sFound))
        nUntil = time.time() + 30
        while not windowsOf(sExe) and time.time() < nUntil: time.sleep(0.5)
    lWindows = windowsOf(sExe)
    if not lWindows:
        sayLine(sApp + " has no window to bring forward.")
        return False
    hWnd = lWindows[0][0]
    from pywinauto.application import Application
    try:
        Application(backend="uia").connect(handle=hWnd).window(handle=hWnd).set_focus()
    except Exception as exceptionError:
        logLine("set_focus failed: " + str(exceptionError))
    time.sleep(1.0)
    sExeNow, sTitle = foregroundExe()
    logLine("in front: " + sExeNow + " " + repr(sTitle))
    return sExeNow.lower() == sExe.lower()


def closeApp(sApp):
    """Closes the app as Alt+F4 would, so its closing says the run is over (3 November 2026). Each window is asked
    to close; nothing is forced, and an app still open after ten seconds is left so and logged."""
    sExe = sApp + ".exe"
    for hWnd, iPid, sTitle in windowsOf(sExe):
        ctypes.windll.user32.PostMessageW(hWnd, 0x0010, 0, 0)  # WM_CLOSE
        logLine("asked to close: " + repr(sTitle))
    nUntil = time.time() + 10
    while windowsOf(sExe) and time.time() < nUntil: time.sleep(0.5)
    bClosed = not windowsOf(sExe)
    logLine(sApp + (" closed" if bClosed else " still open after ten seconds; left so"))
    return bClosed


# ---- AXE.WINDOWS, THE ENGINE OF ACCESSIBILITY INSIGHTS FOR WINDOWS ------------------------------------------------
# --axe scans the app after each step. Since 5 November 2026 the fetching and scanning live in the kit's axeScan.py,
# shared with the check every push runs, so the rules and the advice are one; it sits beside this script in the kit
# and in every app whose build copies the kit's scripts.
def getAxeCli():
    import axeScan
    axeScan.useLog(logLine)
    return axeScan.getAxeCli()


def axeScan(sCli, iPid, sFolder, sScanId):
    import axeScan as oAxe
    oAxe.useLog(logLine)
    return oAxe.axeScan(sCli, iPid, sFolder, sScanId)


def focusNow():
    """What UI Automation says has keyboard focus: name, kind, help text and automation id."""
    try:
        from pywinauto.uia_defines import IUIA
        oElement = IUIA().iuia.GetFocusedElement()
        lParts = []
        for sLabel, sProperty in (("name", "CurrentName"), ("kind", "CurrentLocalizedControlType"),
                                  ("help", "CurrentHelpText"), ("id", "CurrentAutomationId")):
            try:
                sValue = getattr(oElement, sProperty) or ""
            except Exception:
                sValue = ""
            if sValue: lParts.append(sLabel + ": " + str(sValue).strip())
        return "; ".join(lParts) or "(nothing has focus)"
    except Exception as exceptionError:
        return "(UI Automation did not answer: " + str(exceptionError) + ")"


# ---- the app's own speech ----

def speechLogFiles(sApp):
    """The app's speech log files, where Say writes what the app said, newest last.

    THE SESSION LOG TOO (9 November 2026). A Homer app writes what it says as "say:" lines in its session log,
    <App>-<date>-<time>.log, kept in %LOCALAPPDATA%\\<App>\\logs and in the project's own logs folder; only a file named
    for speech was looked for, so every step of DbDo's walks reported "nothing in its speech log" though DbDo had
    spoken. Session logs of the last two days are read now, for their say: lines only."""
    lFiles = []
    nRecent = time.time() - 2 * 86400
    for sFolder in (os.path.join(os.environ.get("LOCALAPPDATA", ""), sApp, "logs"), os.path.join(sProject, "logs")):
        lFiles += glob.glob(os.path.join(sFolder, "*peech*"))
        lFiles += [s for s in glob.glob(os.path.join(sFolder, sApp + "-2*.log"))
                   if re.match(re.escape(sApp) + r"-\d{8}-\d{6}\.log$", os.path.basename(s), re.I) and os.path.getmtime(s) >= nRecent]
    return sorted(set(lFiles), key=os.path.getmtime)


def speechSizes(sApp):
    return {s: os.path.getsize(s) for s in speechLogFiles(sApp)}


def speechSince(sApp, dBefore):
    """What the app's speech logs gained since dBefore."""
    lNew = []
    for sPath in speechLogFiles(sApp):
        iFrom = dBefore.get(sPath, 0)
        try:
            with open(sPath, "rb") as fIn:
                fIn.seek(iFrom)
                sText = fIn.read().decode("utf-8", "replace")
            if "peech" not in os.path.basename(sPath):
                # a session log: only what the app said, from its say: lines
                sText = " | ".join(oMatch.group(1).strip() for oMatch in re.finditer(r"\bsay: (.*)", sText, re.I))
            lNew.append(sText.strip())
        except OSError as exceptionError:
            logLine("speech log unreadable: " + sPath + ": " + str(exceptionError))
    return " | ".join(s for s in lNew if s)


# ---- sound ----

def loopbackMicrophone(sWanted):
    """A speaker as a recording source, so the recording is what the computer plays: the screen reader. The default
    speakers, or the first whose name holds sWanted (--speaker). Every output device is named in the log, since
    JAWS may speak through one that is not the default."""
    import soundcard
    lSpeakers = soundcard.all_speakers()
    oDefault = soundcard.default_speaker()
    for oOne in lSpeakers: logLine("sound output: " + oOne.name + (" (the default)" if oOne.name == oDefault.name else ""))
    oSpeaker = oDefault
    if sWanted:
        lMatch = [o for o in lSpeakers if sWanted.lower() in o.name.lower()]
        if lMatch: oSpeaker = lMatch[0]
        else: sayLine("No sound output's name holds " + sWanted + "; recording the default, " + oDefault.name + ".")
    sayLine("Recording what plays on " + oSpeaker.name + ".")
    return soundcard.get_microphone(id=str(oSpeaker.name), include_loopback=True)


def writeWave(sPath, aFrames):
    """Writes recorded frames as 16-bit mono at the recording rate."""
    import numpy
    aMono = aFrames.mean(axis=1) if aFrames.ndim > 1 else aFrames
    aInt = (numpy.clip(aMono, -1.0, 1.0) * 32767).astype(numpy.int16)
    with wave.open(sPath, "wb") as oWave:
        oWave.setnchannels(1); oWave.setsampwidth(2); oWave.setframerate(iSampleRate)
        oWave.writeframes(aInt.tobytes())
    return sPath


def wordsOf(sText):
    return re.findall(r"[a-z0-9]+", (sText or "").lower())


def hearMatch(sExpected, sHeard):
    """How much of an expected line was heard: the share of its words that were heard, each matched to a heard word
    spelled nearly alike. THE RECOGNIZER MISHEARS FAST SPEECH (3 November 2026): at Jamal's JAWS rate it wrote
    "show now" for "shown now", "looks" for "look's" and "rhymes" for "prime's", so exact, in-order matching scored
    3 and 4 lines of 10 on runs where JAWS said every line. A word now counts when a heard word is at least 75%
    alike, and order no longer matters within a step."""
    lWant, lGot = wordsOf(sExpected), wordsOf(sHeard)
    if not lWant: return 1.0
    iFound = 0
    for sWant in lWant:
        if sWant in lGot or any(difflib.SequenceMatcher(None, sWant, sGot).ratio() >= 0.75 for sGot in lGot): iFound += 1
    return iFound / float(len(lWant))


# ---- the run ----

def readMenusToFile(sApp, sPath):
    """--read-menus: the menus of sApp's window, written to sPath as JSON, in this process alone, with its own log."""
    global sLogPath
    sLogPath = os.path.splitext(sPath)[0] + ".log"
    logLine("menu reader started for %s; Python %s; %s" % (sApp, sys.version.split()[0], platform.platform()))
    lMenus = readMenus(sApp)
    with open(sPath, "w", encoding="utf-8") as fOut: json.dump(lMenus, fOut)
    print("%d menus" % len(lMenus))
    return 0 if lMenus else 1


def main():
    if len(sys.argv) == 4 and sys.argv[1] == "--read-menus":
        return readMenusToFile(sys.argv[2], sys.argv[3])
    global sLogPath
    oParser = argparse.ArgumentParser(description="Drive a Homer app through a walk and record what is said.")
    oParser.add_argument("which", nargs="?", default="", help="a tutorial's number, performed live, or a walk's path")
    oParser.add_argument("--walk", default=os.path.join(sProject, "scripts", "walks", "pilot.inix"))  # the app's own walks
    oParser.add_argument("--app", default=os.path.basename(sProject))
    oParser.add_argument("--steps", type=int, default=0)
    oParser.add_argument("--wait", type=int, default=iDefaultWait)
    oParser.add_argument("--model", default=sDefaultModel)
    oParser.add_argument("--no-transcribe", action="store_true")
    oParser.add_argument("--speaker", default="")
    oParser.add_argument("--keep", action="store_true")
    oParser.add_argument("--exe", default="")              # the program to start when the app is not running (runMacro: the project's own build)
    oParser.add_argument("--axe", action="store_true")      # kept for old command lines; scans are on by default
    oParser.add_argument("--no-axe", action="store_true")
    oParser.add_argument("--tutorial", action="store_true")
    oParser.add_argument("--menus", type=int, default=7)     # with "menus": how many top-level menus to walk
    oParser.add_argument("--items", type=int, default=25)    # with "menus": how many items to walk in each menu and the Alternate Menu
    oArgs = oParser.parse_args([s.lower() if s.startswith("--") else s for s in sys.argv[1:]])

    global sStartExe
    sStartExe = oArgs.exe
    sLogs = os.path.join(sProject, "logs"); os.makedirs(sLogs, exist_ok=True)
    sLogPath = os.path.join(sLogs, oArgs.app + "-recordWalk-" + sStamp + ".log")
    sReportPath = os.path.join(sLogs, oArgs.app + "-recordWalk-" + sStamp + ".md")
    sWaves = os.path.join(sProject, "results", "recordWalk-" + sStamp); os.makedirs(sWaves, exist_ok=True)
    logLine("recordWalk started: " + os.path.abspath(__file__))
    logLine("Python " + sys.version.replace("\n", " ") + " at " + sys.executable + "; " + platform.platform())
    logLine("working folder " + os.getcwd() + "; command line " + " ".join(sys.argv))
    logLine("settings: app " + oArgs.app + ", walk " + oArgs.walk + ", steps " + str(oArgs.steps) + ", wait " + str(oArgs.wait)
            + ", model " + oArgs.model + ", transcribe " + str(not oArgs.no_transcribe))

    # recordWalk 1 PERFORMS TUTORIAL 1 LIVE (5 November 2026, evening): "recordWalk" alone walks the pilot, and the
    # menu tour needed --tutorial --walk 1, which no one remembers. A number given first is a tutorial, performed live.
    # recordWalk menus WALKS THE MENUS AND THE ALTERNATE MENU (9 November 2026). No walk is written by hand: one is made
    # here, in logs. F10 opens the menu bar and Down Arrow the first menu; Down Arrow then visits each item, and Right
    # Arrow the next menu, for --menus menus of --items items each (a menu shorter than that wraps round, which the
    # report shows as repeats). Escape twice leaves the menus. Alt+F10 opens the Alternate Menu, Down Arrow visits its
    # first --items commands, Home and End reach its ends, and Escape closes it. Each item is one step, heard and logged.
    if oArgs.which.strip().lower() == "menus":
        sMenuWalk = os.path.join(sLogs, oArgs.app + "-menus-" + sStamp + ".inix")
        lLines = ["; written by recordWalk menus, " + sStamp, ""]
        # FROM THE CODE FIRST (10 November 2026): the app's own source, beside this script, says each menu's items in
        # order -- the same reading the release gate checks. UI Automation is tried only when the code gives none.
        lMenus = menuTreeFromCode(sProject)
        if lMenus and any(d["items"] for d in lMenus):
            logLine("menus read from the code in " + sProject + ": " + ", ".join("%s (%d)" % (d["name"], len(d["items"])) for d in lMenus))
            sayLine("Read " + oArgs.app + "'s menus from its code.")
        else:
            lMenus = []
        # IN A PROCESS OF ITS OWN (10 November 2026). Reading the menus through UI Automation, in this process, came
        # before the recording, and the walk of 10 November stopped at "Starting in 5 seconds": pywinauto and soundcard
        # share COM badly, as ensurePackages notes. A separate process reads the menus into a file, with 90 seconds to
        # do it, and this one never touches UI Automation before it records.
        # THE PACKAGES FIRST (10 November 2026): bringForward loads pywinauto, and soundcard must start COM before it,
        # or the recording fails with "Error 0x100000001", as the walk of 10 November did at 8:32.
        if not lMenus and ensurePackages() and bringForward(oArgs.app):
            sayLine("Reading " + oArgs.app + "'s menus.")
            sMenuFile = os.path.join(sLogs, oArgs.app + "-menuRead-" + sStamp + ".json")
            lReader = [sys.executable, os.path.abspath(__file__), "--read-menus", oArgs.app, sMenuFile]
            try:
                oDone = subprocess.run(lReader, timeout=90, capture_output=True, text=True)
                logLine("menu reader exit %d: %s" % (oDone.returncode, (oDone.stdout + oDone.stderr).strip()[:2000]))
                if os.path.isfile(sMenuFile):
                    with open(sMenuFile, encoding="utf-8") as fIn: lMenus = json.load(fIn)
            except subprocess.TimeoutExpired:
                logLine("menu reader: still reading after 90 seconds, so stopped")
                sayLine("Reading the menus took too long.")
            bringForward(oArgs.app)
        if lMenus and not any(d["items"] for d in lMenus):
            logLine("menu audit: %d menus but no items read, so the fixed walk is used" % len(lMenus))
            lMenus = []
        if lMenus:
            lFindings = checkMenus(lMenus)
            iItems = sum(len(d["items"]) + sum(len(s["items"]) for s in d["items"]) for d in lMenus)
            lsMenuAudit.extend(["## Menu audit", "",
                "%d menus, %d items, submenus included; %d finding%s against Homer's menu rules." % (len(lMenus), iItems, len(lFindings), "" if len(lFindings) == 1 else "s"), ""]
                + ["- " + s for s in lFindings] + [""] + ["### The menus as read", ""])
            for dMenu in lMenus:
                lsMenuAudit.append("- **%s** (Alt+%s)" % (dMenu["name"], dMenu["access"] or "?"))
                for dItem in dMenu["items"]:
                    lsMenuAudit.append("    - %s, letter %s, key %s%s" % (dItem["name"], dItem["access"] or "none", dItem["key"] or "none", "" if dItem["enabled"] else ", unavailable now"))
                    for dSub in dItem["items"]:
                        lsMenuAudit.append("        - %s, letter %s, key %s" % (dSub["name"], dSub["access"] or "none", dSub["key"] or "none"))
            lsMenuAudit.append("")
            sayLine("%d menus read, %d items; %d finding%s against the menu rules." % (len(lMenus), iItems, len(lFindings), "" if len(lFindings) == 1 else "s"))
            lLines += menuWalkSteps(lMenus, oArgs.items)
            with open(sMenuWalk, "w", encoding="utf-8-sig", newline="\r\n") as fWalk:
                fWalk.write("\n".join(lLines))
            logLine("menu walk written from the menus as read: %s" % sMenuWalk)
            oArgs.which = sMenuWalk
        else:
            sayLine("The menus could not be read, so the walk presses the keys a fixed number of times instead.")
            def menuStep(sKey, sSay, iWait=2):
                lLines.extend(["[step]", "Say=" + sSay, "Key=" + sKey, "Wait=" + str(iWait), ""])
            menuStep("F10", "F10, the menu bar.", 3)
            for iMenu in range(1, max(1, oArgs.menus) + 1):
                menuStep("Down", "Menu %d, opened." % iMenu, 3)
                for iItem in range(2, max(2, oArgs.items) + 1):
                    menuStep("Down", "Menu %d, item %d." % (iMenu, iItem))
                if iMenu < oArgs.menus:
                    menuStep("Escape", "Menu %d, closed." % iMenu)
                    menuStep("Right", "The next menu's name.")
            menuStep("Escape", "Leaving the menu.")
            menuStep("Escape", "Leaving the menu bar.")
            menuStep("Alt+F10", "Alt+F10, the Alternate Menu.", 4)
            for iItem in range(2, max(2, oArgs.items) + 1):
                menuStep("Down", "Alternate Menu, command %d." % iItem)
            menuStep("End", "Alternate Menu, the last command.")
            menuStep("Home", "Alternate Menu, the first command.")
            menuStep("Escape", "The Alternate Menu closed.")
            with open(sMenuWalk, "w", encoding="utf-8-sig", newline="\r\n") as fWalk:
                fWalk.write("\n".join(lLines))
            logLine("menu walk written: %s, %d menus of %d items" % (sMenuWalk, oArgs.menus, oArgs.items))
            oArgs.which = sMenuWalk
    if oArgs.which:
        oArgs.walk = oArgs.which
        if oArgs.which.strip().isdigit(): oArgs.tutorial = True
    # A NUMBER NAMES A TUTORIAL, AND A RELATIVE PATH IS THE PROJECT'S (5 November 2026): recordWalk is usually
    # started from the scripts folder, so help\... and 1 are looked for from the project folder.
    if re.match(r"^\d+$", oArgs.walk.strip()):
        lFound = glob.glob(os.path.join(sProject, "help", "Tutorial_" + oArgs.walk.strip() + "_*.inix"))
        if lFound: oArgs.walk = lFound[0]
    if not os.path.isfile(oArgs.walk) and os.path.isfile(os.path.join(sProject, oArgs.walk)):
        oArgs.walk = os.path.join(sProject, oArgs.walk)
    if not os.path.isfile(oArgs.walk):
        sayLine("No walk at " + oArgs.walk + ".")
        return 1
    lSteps = readWalk(oArgs.walk, oArgs.tutorial)
    if oArgs.steps > 0: lSteps = lSteps[:oArgs.steps]
    sayLine("recordWalk: " + str(len(lSteps)) + (" step" if len(lSteps) == 1 else " steps") + " from " + os.path.basename(oArgs.walk) + ".")
    if not ensurePackages(): return 1
    WhisperModel = None if oArgs.no_transcribe else ensureWhisper()
    oModel = None
    if WhisperModel is not None:
        # SAID ONLY WHEN TRUE (6 November 2026): every run said the model is downloaded "the first time", though after
        # the first it loads from its cache in seconds -- and the faster-whisper component fetches it at install.
        sCache = os.path.join(os.path.expanduser("~"), ".cache", "huggingface", "hub", "models--Systran--faster-whisper-" + oArgs.model)
        if os.path.isdir(sCache): sayLine("Loading the recognizer, " + oArgs.model + ".")
        else: sayLine("Getting the recognizer, " + oArgs.model + ", once; this takes a few minutes, and later walks load it in seconds.")
        try:
            oModel = WhisperModel(oArgs.model, device="cpu", compute_type="int8")
        except Exception as exceptionError:
            sayLine("The recognizer did not load, so the steps are recorded but not written down. The log says why.")
            logLine("WhisperModel failed: " + traceback.format_exc())
    sCppExe, sCppModel = ("", "")
    if oModel is None and not oArgs.no_transcribe:
        sCppExe, sCppModel = findWhisperCpp()
        if sCppExe: sayLine("Using HomerScribe's whisper.cpp to write the speech down.")
        else: sayLine("No recognizer could be loaded, so the steps are recorded but not written down.")
    try:
        oMicrophone = loopbackMicrophone(oArgs.speaker)
    except Exception:
        sayLine("The computer's sound could not be recorded. The log says why.")
        logLine("loopback failed: " + traceback.format_exc())
        return 1

    sayLine("Starting in 5 seconds. Keep your hands off the keyboard. To stop, switch to another window.")
    time.sleep(5)
    if not bringForward(oArgs.app):
        sayLine(oArgs.app + " could not be brought to the front, so nothing was pressed.")
        return 1

    from pywinauto.keyboard import send_keys
    iMatched, iChecked, iNoisySteps = 0, 0, 0
    lsReport.extend(["# recordWalk: " + oArgs.app + ", " + os.path.basename(oArgs.walk), "",
                     "Recorded " + datetime.datetime.now().strftime("%d %B %Y at %H:%M") + ". Each step: the key pressed, what the walk expected, "
                     "what was heard, what UI Automation says has focus, and what " + oArgs.app + " itself said.", ""])
    sIdentity = appIdentity(oArgs.app + ".exe")
    lsReport.extend([sIdentity, ""])
    logLine(sIdentity)
    sAxeCli, iAxeErrors = "", 0
    # EACH WINDOW SCANNED ONCE, BY DEFAULT, AND SAID (5 November 2026, night): the kit's axeWindow scans one window by
    # its handle and speaks every 10 seconds, so a scan never sounds like a hang; --no-axe leaves the scans out.
    if not oArgs.no_axe:
        try:
            sAxeCli = getAxeCli()
            sayLine("Each step is also scanned with Axe.Windows, the engine of Accessibility Insights.")
        except Exception:
            logLine("AxeWindowsCLI not available: " + traceback.format_exc())
            sayLine("AxeWindowsCLI could not be got, so no scans; the log says why.")
    lLive = []
    bAxeStopped = False
    setScanned, lAxeFindings = set(), []
    setLaunch = launchKey(sProject, oArgs.app)
    logLine("launch key: " + ("+".join(sorted(setLaunch)) if setLaunch else "none found"))
    for iIndex, dStep in enumerate(lSteps, 1):
        sExeNow, sTitleNow = foregroundExe()
        # ITS OWN CONSOLE IS NOT A REASON TO STOP (10 November 2026). testSay's run of 9 November stopped at step 97 of
        # 106 because the command prompt running testSay had come to the front. When the window in front is a console
        # -- the one running this walk -- the app is brought back once and the walk goes on; any other window stops it,
        # as before, so keys never go where they were not meant to.
        if sExeNow.lower() != (oArgs.app + ".exe").lower() and sExeNow.lower() in ("cmd.exe", "conhost.exe", "windowsterminal.exe", "openconsole.exe", "python.exe", "py.exe", "powershell.exe", "pwsh.exe"):
            logLine("before step %d the console %s %r was in front; bringing %s back" % (iIndex, sExeNow, sTitleNow, oArgs.app))
            if bringForward(oArgs.app):
                time.sleep(0.5)
                sExeNow, sTitleNow = foregroundExe()
        if sExeNow.lower() != (oArgs.app + ".exe").lower():
            sayLine("Stopped before step " + str(iIndex) + ": " + (sTitleNow or sExeNow or "another window") + " is in front.")
            logLine("stopped: foreground " + sExeNow + " " + repr(sTitleNow))
            break
        # THE APP'S OWN LAUNCH KEY IS NOT PRESSED (5 November 2026, night). Tutorial 1 opens DbDo with Alt+Control+D, its
        # desktop hotkey; recordWalk has already started DbDo, so the key went to Windows Explorer, which took the front,
        # and the walk stopped at step 3. The narration is kept; the key, already done, is not pressed.
        if oArgs.tutorial and dStep["sKey"] and setLaunch and keySet(dStep["sKey"]) == setLaunch:
            logLine("step " + str(iIndex) + ": " + dStep["sKey"] + " is " + oArgs.app + "'s launch key, and it is open already; not pressed")
            try:
                if dStep["sSay"]: lLive.append(speechFrames(dStep["sSay"], 0))
                if dStep["lHear"]: lLive.append(speechFrames(" ".join(dStep["lHear"]), 1))
            except Exception:
                logLine("tutorial voice failed at step " + str(iIndex) + ": " + traceback.format_exc())
            continue
        if oArgs.tutorial:
            try:
                if dStep["sSay"]: lLive.append(speechFrames(dStep["sSay"], 0))
                if not (dStep["sKey"] or dStep["sType"]):
                    if dStep["lHear"]: lLive.append(speechFrames(" ".join(dStep["lHear"]), 1))
                    logLine("step " + str(iIndex) + ": narration only")
                    continue
            except Exception:
                logLine("tutorial voice failed at step " + str(iIndex) + ": " + traceback.format_exc())
                if not (dStep["sKey"] or dStep["sType"]): continue
        iWait = dStep["iWait"] or oArgs.wait
        if dStep.get("nPause"):
            logLine("step " + str(iIndex) + ": pause %.1f s first, as the walk asks" % dStep["nPause"])
            time.sleep(dStep["nPause"])
        dBefore = speechSizes(oArgs.app)
        sKeys = keysFor(dStep["sKey"]) if dStep["sKey"] else ""
        logLine("step " + str(iIndex) + ": type " + repr(dStep["sType"]) + ", key " + repr(dStep["sKey"]) + " as " + repr(sKeys) + ", listen " + str(iWait) + " s")
        # LISTEN UNTIL JAWS STOPS, NOT FOR A FIXED TIME (3 November 2026). The recordings of nine walks showed JAWS
        # still speaking when the four-second window closed -- Filter Records, every Escape back to the grid, Order
        # Records -- and the rest spilling into the next step's recording, where it was taken for that step's words.
        # So: before the key, wait for half a second of quiet (up to ten seconds); after it, record until JAWS has been
        # quiet for 0.8 seconds, at most fifteen seconds, and stop at Wait= only if nothing was said at all.
        import numpy
        iChunk = int(iSampleRate * nChunkSeconds)
        def levelOf(aChunk):
            aOne = aChunk.mean(axis=1) if aChunk.ndim > 1 else aChunk
            return float(numpy.sqrt(numpy.mean(numpy.square(aOne)))) if aOne.size else 0.0
        with oMicrophone.recorder(samplerate=iSampleRate) as oRecorder:
            nQuiet, nWaited = 0.0, 0.0
            while nQuiet < nQuietBefore and nWaited < nMaxSettle:
                nLevel = levelOf(oRecorder.record(numframes=iChunk))
                nQuiet = nQuiet + nChunkSeconds if nLevel < nSpeechLevel else 0.0
                nWaited += nChunkSeconds
            bNoisy = nWaited >= nMaxSettle
            if bNoisy:
                logLine("  still sound after %d s of waiting for quiet; pressed anyway" % nMaxSettle)
                iNoisySteps += 1
                # OTHER SOUND IS SAID ONCE (17 November 2026): DbDo's first live test recorded a talk playing elsewhere,
                # and its words were written down as what JAWS said. The console says so, once, while it still matters.
                if iNoisySteps == 1: sayLine("Other sound is playing, so what was heard may not be the screen reader. Pause it for a cleaner run.")
            if dStep["sType"]: send_keys(escapeText(dStep["sType"]), with_spaces=True, pause=0.05)
            # A KEY IS A KEY, NOT A CHARACTER (17 November 2026): send_keys sends a plain character with no modifier as a
            # Unicode packet, not as its key. FileDir's testSay pressed Apostrophe for Say Item Name, JAWS's pass-through
            # script retyped the packet as Semicolon, and FileDir marked the item instead. Keys go as virtual keys;
            # only Type= text goes as characters.
            if sKeys: send_keys(sKeys, pause=0.05, vk_packet=False)
            if sKeys and dStep.get("iPress", 1) >= 2:
                time.sleep(0.25)   # well inside the 1.2 seconds Homer counts as one double press
                send_keys(sKeys, pause=0.05, vk_packet=False)
            lChunks, nQuiet, nSpoken, nElapsed, bSpoke = [], 0.0, 0.0, 0.0, False
            while True:
                aChunk = oRecorder.record(numframes=iChunk)
                lChunks.append(aChunk)
                nElapsed += nChunkSeconds
                if levelOf(aChunk) >= nSpeechLevel: bSpoke, nQuiet, nSpoken = True, 0.0, nSpoken + nChunkSeconds
                else: nQuiet += nChunkSeconds
                if bSpoke and nQuiet >= nQuietAfter: break
                if not bSpoke and nElapsed >= iWait: break
                if nElapsed >= nMaxListen: break
            aFrames = numpy.concatenate(lChunks)
            bCutOff = bSpoke and nQuiet < nQuietAfter
        logLine("  listened %.2f s; speech %.2f s%s" % (nElapsed, nSpoken, "; STILL SPEAKING at the limit" if bCutOff else ""))
        sWave = writeWave(os.path.join(sWaves, "step%02d.wav" % iIndex), aFrames)
        if oArgs.tutorial: lLive.append(framesToInt16(aFrames))
        # HOW LOUD THE STEP WAS (3 November 2026): a recording of the wrong device is silence, which looks like
        # a reader that said nothing. The peak says which: below about 0.01 nothing was played there.
        aMono = (aFrames.mean(axis=1) if aFrames.ndim > 1 else aFrames).astype(numpy.float32)
        nPeak = float(numpy.abs(aMono).max()) if aMono.size else 0.0
        logLine("  loudness peak %.3f" % nPeak)
        sFocus = focusNow()
        sApp = speechSince(oArgs.app, dBefore)
        sAxe = ""
        # ONE WINDOW AT A TIME, EACH ONCE (5 November 2026, evening). The kit's axeWindow scans the window in front by its
        # handle, through the Axe.Windows library -- a dialog without the record list behind it -- and speaks as it
        # goes: at the start, every 10 seconds, and at the end, between this step's recording and the next. Each window
        # is scanned the first time it is in front; the main window, which may hold a long list, takes longest, and
        # says so. When the library cannot be loaded, the walk says so once and goes on unscanned.
        if sAxeCli and not bAxeStopped:
            try:
                lWindows = windowsOf(oArgs.app + ".exe")
                if lWindows:
                    import axeWindow
                    axeWindow.useLog(logLine)
                    iPid = lWindows[0][1]
                    hWnd = axeWindow.frontWindowOf(iPid)
                    sTitle = axeWindow.windowTitle(hWnd) if hWnd else ""
                    # DIALOGS ONLY (6 November 2026). The 8:56 walk spent seven minutes on DbDo's main window, record
                    # list and all; every push already scans an app's main window, so a walk scans the windows it
                    # opens -- the ones another window owns -- and passes the main window by.
                    if hWnd and not ctypes.windll.user32.GetWindow(hWnd, 4):
                        sAxe = (sTitle or "the main window") + ": the main window, scanned by every push, so not here"
                        hWnd = 0
                    if hWnd and sTitle in setScanned:
                        sAxe = sTitle + " was scanned at an earlier step"
                    elif hWnd:
                        setScanned.add(sTitle)
                        dScan = axeWindow.scanWindow(iPid, hWnd, os.path.join(sWaves, "axe"), "step%02d" % iIndex)
                        if dScan["sProblem"]:
                            sAxe = "no scan of " + (sTitle or "the window") + ": " + dScan["sProblem"]
                            if "could not be loaded" in dScan["sProblem"]:
                                bAxeStopped = True
                                sayLine("Axe.Windows could not be loaded, so the rest of the walk is not scanned; the log says why.")
                        else:
                            iAxeErrors += dScan["iErrors"]
                            lAxeFindings.extend(dScan["lFindings"])
                            sAxe = (str(dScan["iErrors"]) + (" error" if dScan["iErrors"] == 1 else " errors") + " in " + (sTitle or "the window")
                                    + (": " + "; ".join(dScan["lFindings"][:3]) + (" and more" if len(dScan["lFindings"]) > 3 else "") if dScan["lFindings"] else ""))
            except Exception:
                logLine("axe scan failed for step " + str(iIndex) + ": " + traceback.format_exc())
                sAxe = "the scan could not run; the log says why"
        sHeard = ""
        if oModel is not None:
            try:
                # THE SOUND ITSELF, NOT THE FILE (3 November 2026): given a file, faster-whisper opens it with PyAV,
                # and the PyAV pip installs is newer than it expects ("unexpected keyword argument
                # 'metadata_errors'"). Given the samples, 16 kHz mono floats, it needs no decoder at all.
                lSegments, _oInfo = oModel.transcribe(aMono, language="en", vad_filter=True)
                sHeard = " ".join(o.text.strip() for o in lSegments).strip()
            except Exception:
                logLine("transcription failed for step " + str(iIndex) + ": " + traceback.format_exc())
                sHeard = "(not written down: the recognizer failed; the log says why)"
        elif sCppExe:
            try:
                sHeard = transcribeCpp(sCppExe, sCppModel, sWave)
            except Exception:
                logLine("whisper.cpp failed for step " + str(iIndex) + ": " + traceback.format_exc())
        logLine("  heard: " + sHeard)
        logLine("  focus: " + sFocus)
        logLine("  app said: " + sApp)
        lsReport.append("## Step " + str(iIndex) + ": " + (("type " + dStep["sType"] + ", then ") if dStep["sType"] else "") + (dStep["sKey"] or ""))
        lsReport.append("")
        if dStep["sSay"]: lsReport.append("- **Walk says:** " + dStep["sSay"])
        for sWant in dStep["lHear"]:
            nShare = hearMatch(sWant, sHeard + " " + sApp)
            bHeard = nShare >= nMatchNeeded
            iChecked += 1; iMatched += 1 if bHeard else 0
            lsReport.append("- **Expected:** " + sWant + " -- " + ("heard" if bHeard else "NOT HEARD") + " (" + str(int(nShare * 100)) + "% of its words)")
        # THE WINDOW AND THE FOCUS (17 November 2026): checked after the listening, when the step has settled; a window
        # that is slow to open is given up to five more seconds.
        for sWant in dStep.get("lWindow", []):
            nUntil = time.time() + 5
            sTitleAfter = foregroundExe()[1]
            while sWant.lower() not in sTitleAfter.lower() and time.time() < nUntil:
                time.sleep(0.25); sTitleAfter = foregroundExe()[1]
            bShown = sWant.lower() in sTitleAfter.lower()
            iChecked += 1; iMatched += 1 if bShown else 0
            logLine("  window wanted %r, in front %r: %s" % (sWant, sTitleAfter, "yes" if bShown else "NO"))
            lsReport.append("- **Window expected:** " + sWant + " -- " + ("shown" if bShown else "NOT SHOWN; in front: " + (sTitleAfter or "no title")))
        for sWant in dStep.get("lFocus", []):
            bFocused = sWant.lower() in sFocus.lower()
            iChecked += 1; iMatched += 1 if bFocused else 0
            logLine("  focus wanted %r: %s" % (sWant, "yes" if bFocused else "NO"))
            lsReport.append("- **Focus expected:** " + sWant + " -- " + ("there" if bFocused else "NOT THERE"))
        lsReport.append("- **Heard:** " + (sHeard or "(nothing written down)") + (" -- OTHER SOUND was playing before the key, so this may not be the reader" if bNoisy else ""))
        lsReport.append("- **Focus:** " + sFocus)
        lsReport.append("- **" + oArgs.app + " said:** " + (sApp or "(nothing in its speech log)"))
        if sAxeCli: lsReport.append("- **Axe.Windows:** " + (sAxe or "(no scan)"))
        lsReport.append("- **Recording:** " + os.path.relpath(sWave, sProject) + ", %.1f s of speech in %.1f s, loudness peak %.3f%s%s" % (nSpoken, nElapsed, nPeak,
                        " -- SILENT: nothing played on the recorded output; try --speaker" if nPeak < 0.01 else "",
                        " -- STILL SPEAKING at %d s, so the end is missing" % nMaxListen if bCutOff else ""))
        lsReport.append("")
        logLine("step " + str(iIndex) + " recorded")
    lsReport.insert(3, "**" + str(iMatched) + " of " + str(iChecked) + " expected lines heard.**" + (" Other sound was playing at %d of the steps, so what was heard there may not be the reader; the app's own log is the surer witness." % iNoisySteps if iNoisySteps else ""))
    lsReport.insert(4, "")
    if sAxeCli: lsReport.insert(5, "**Axe.Windows: " + str(iAxeErrors) + (" error" if iAxeErrors == 1 else " errors") + " across the steps; each .a11ytest file opens in Accessibility Insights for Windows.**\n")
    if lsMenuAudit: lsReport[5:5] = lsMenuAudit
    with open(sReportPath, "w", encoding="utf-8") as fOut: fOut.write("\n".join(lsReport) + "\n")
    if sAxeCli: sayLine("Axe.Windows: " + str(iAxeErrors) + (" error" if iAxeErrors == 1 else " errors") + " across the steps; the report lists them by step.")
    if oArgs.tutorial and lLive:
        try:
            sLive = writeLive(sWaves, oArgs.walk, lLive)
            sayLine("The tutorial, performed live, is " + os.path.relpath(sLive, sProject) + ".")
        except Exception:
            logLine("live recording not written: " + traceback.format_exc())
            sayLine("The live recording could not be written; the log says why.")
    if not oArgs.keep:
        bClosed = closeApp(oArgs.app)
        if not bClosed: sayLine(oArgs.app + " did not close by itself, so it is still open.")
    sayLine("recordWalk is done. " + str(iMatched) + " of " + str(iChecked) + " expected lines heard. The report is " + sReportPath + ".")
    sayLine("Upload the report, the log and the results\\recordWalk-" + sStamp + " folder.")
    iCode = 0 if iMatched == iChecked else 1
    # TUTORIAL 1 LIVE WHEN IT IS DUE (5 November 2026, night): a plain recordWalk also performs Tutorial 1 -- the menus
    # through the real screen reader -- when its live recording is missing, or older than the tutorial or the program,
    # so no argument has to be remembered; it runs as a second pass, with the app started afresh.
    if not oArgs.which and not oArgs.tutorial and os.path.basename(oArgs.walk).lower() == "pilot.inix":
        lTutorial = glob.glob(os.path.join(sProject, "help", "Tutorial_1_*.inix"))
        if lTutorial:
            lLive = glob.glob(os.path.join(sProject, "results", "recordWalk-*", os.path.splitext(os.path.basename(lTutorial[0]))[0] + "-live.*"))
            nLive = max([os.path.getmtime(s) for s in lLive] or [0])
            sExe = os.path.join(sProject, "exec", oArgs.app + ".exe")
            nNewest = max(os.path.getmtime(lTutorial[0]), os.path.getmtime(sExe) if os.path.isfile(sExe) else 0)
            if nLive < nNewest:
                sayLine("Now Tutorial 1, performed live: the menus through the screen reader. Keep your hands off the keyboard.")
                lCommand = [sys.executable, os.path.abspath(__file__), "1"] + (["--no-axe"] if oArgs.no_axe else [])
                logLine("tutorial pass: " + " ".join(lCommand))
                iCode = max(iCode, subprocess.call(lCommand, cwd=os.getcwd()))
            else:
                logLine("Tutorial 1's live recording is current: " + max(lLive, key=os.path.getmtime))
    return iCode


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        sayLine("Stopped.")
        sys.exit(1)
    except Exception:
        sayLine("recordWalk stopped on an error; the log has it.")
        logLine("ERROR: " + traceback.format_exc())
        sys.exit(1)
