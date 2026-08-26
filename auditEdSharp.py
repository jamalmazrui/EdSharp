r"""auditEdSharp.py -- checks on the EdSharp sources that a compiler cannot make.

    python auditEdSharp.py            (run from C:\EdSharp)
    python auditEdSharp.py -pathRoot C:\EdSharp

WHAT THIS IS FOR

The compiler proves the code is valid C#. It cannot prove that two
commands do not claim the same shortcut, that a regular expression in
the compiler table is even a legal pattern, that every button in a
dialog has its own access key, or that the conversion table points at
scripts that exist. Every one of those has broken EdSharp at least once,
and each break cost a test-and-report cycle with a user. This script
checks them in a second, before the build.

Each check prints PASS or FAIL with a plain sentence. The exit code is 0
when everything passes and 1 when anything fails, so a build script can
stop on it. A detailed log is written beside this script, whatever
happens.
"""

import datetime
import os
import re
import sys

c_sLogName = "auditEdSharp.log"

pathRoot = os.path.dirname(os.path.abspath(__file__))
if "-pathRoot" in sys.argv:
    pathRoot = sys.argv[sys.argv.index("-pathRoot") + 1]
pathLog = os.path.join(os.path.dirname(os.path.abspath(__file__)), c_sLogName)
fileLog = None
lFailures = []


def say(sMessage=""):
    print(sMessage)
    if fileLog:
        try:
            fileLog.write(sMessage + "\n")
            fileLog.flush()
        except Exception:
            pass


def report(sName, bPassed, sDetail=""):
    say(("PASS  " if bPassed else "FAIL  ") + sName + ((": " + sDetail) if sDetail else ""))
    if not bPassed:
        lFailures.append(sName)


def readFile(sName):
    pathFile = os.path.join(pathRoot, sName)
    if not os.path.isfile(pathFile):
        return None
    with open(pathFile, "r", encoding="utf-8-sig", errors="replace") as fileIn:
        return fileIn.read()


def plural(iCount, sNoun):
    return str(iCount) + " " + sNoun + ("" if iCount == 1 else "s")


def checkShortcutsUnique(sCode):
    """No two commands may claim the same key."""
    dKeys = {}
    lClashes = []
    for oMatch in re.finditer(r'CreateMenuItem\(\s*"([^"]*)"\s*,\s*"([^"]*)"', sCode):
        sName, sKey = oMatch.group(1), oMatch.group(2).strip()
        if not sKey or sKey == "~":
            continue
        sNormal = "+".join(sorted(p.strip().lower() for p in sKey.split("+")))
        if sNormal in dKeys and dKeys[sNormal] != sName:
            lClashes.append(sKey + " claimed by " + dKeys[sNormal] + " and " + sName)
        else:
            dKeys[sNormal] = sName
    report("Command shortcuts are unique", not lClashes,
           "; ".join(lClashes) if lClashes else plural(len(dKeys), "shortcut") + " checked")


def checkAccessKeysUnique(sCode):
    """Every dialog's buttons must carry distinct access keys."""
    lProblems = []
    iChecked = 0
    iSkipped = 0
    lCalls = re.findall(r'runWithButtons\(new string\[\]\s*\{([^}]*)\}', sCode)
    # PickAndChoose and Choose take their buttons the same way, and add
    # Cancel themselves when the caller has not.
    lCalls += re.findall(r'(?:PickAndChoose|Choose)\([^;]*?new string\[\]\s*\{([^}]*)\}', sCode)
    for sInside in lCalls:
        oMatch = type("m", (), {"group": staticmethod(lambda i, s=sInside: s)})
        lLabels = re.findall(r'"([^"]*)"', oMatch.group(1))
        if not lLabels:
            continue
        iChecked += 1
        # A label built at run time -- "&" + a variable -- cannot be checked
        # from the source, so a set containing one is counted and skipped.
        if any(sLabel.strip() in ("&", "") for sLabel in lLabels):
            iSkipped += 1
            continue
        dSeen = {}
        # Help is added automatically by the dialog and claims H.
        lAll = list(lLabels) + ["&Help"]
        if not any(l.replace("&", "").lower() == "cancel" for l in lLabels):
            lAll.append("&Cancel")
        for sLabel in lAll:
            iAmp = sLabel.find("&")
            if iAmp < 0 or iAmp + 1 >= len(sLabel):
                lProblems.append(sLabel + " has no access key")
                continue
            sKey = sLabel[iAmp + 1].upper()
            if sKey in dSeen:
                lProblems.append(sKey + " claimed by " + dSeen[sKey] + " and " + sLabel)
            dSeen[sKey] = sLabel
    report("Dialog buttons have distinct access keys", not lProblems,
           "; ".join(lProblems) if lProblems else plural(iChecked, "button set") + " checked, "
           + str(iSkipped) + " with run-time labels skipped")


def checkRegexesCompile(sInix):
    """Every pattern in the compiler table must be a legal expression."""
    lProblems = []
    iChecked = 0
    for sKey in ("JumpPosition", "AbbreviateOutput", "NavigatePart"):
        for oMatch in re.finditer(r'^' + sKey + r'="(.*)"\s*$', sInix, re.M):
            sPattern = oMatch.group(1)
            iChecked += 1
            try:
                re.compile(sPattern)
            except re.error as oError:
                lProblems.append(sKey + " " + sPattern + " (" + str(oError) + ")")
    report("Compiler table patterns compile", not lProblems,
           "; ".join(lProblems) if lProblems else plural(iChecked, "pattern") + " checked")


def checkConversionScriptsExist(sInix):
    """Every conversion command must name a script that is present.

    Only the scripts EdSharp itself ships -- the .cmd and .py files kept
    in the repository -- are treated as required. The .exe tools are
    fetched into Convert by the build and are absent from a fresh clone,
    so a missing one is reported for information rather than as a
    failure."""
    lMissingScripts = []
    lMissingTools = []
    lScripts, lTools = set(), set()
    for oMatch in re.finditer(r'Convert\\+(?:[A-Za-z0-9_]+\\+)*([A-Za-z0-9_]+\.(?:cmd|py|exe))', sInix):
        sName = oMatch.group(1)
        (lTools if sName.lower().endswith(".exe") else lScripts).add(sName)
    for sName in sorted(lScripts):
        if not findUnder(os.path.join(pathRoot, "Convert"), sName):
            lMissingScripts.append(sName)
    for sName in sorted(lTools):
        if not findUnder(os.path.join(pathRoot, "Convert"), sName):
            lMissingTools.append(sName)
    report("Conversion scripts named in the table exist", not lMissingScripts,
           "; ".join(lMissingScripts) if lMissingScripts else plural(len(lScripts), "script") + " checked")
    if lMissingTools:
        say("NOTE  fetched tools not present in this folder (the build or installer supplies them): "
            + ", ".join(lMissingTools))


def findUnder(pathDir, sName):
    """True when a file of this name exists anywhere under a folder."""
    if not os.path.isdir(pathDir):
        return False
    for pathWalk, lDirs, lFiles in os.walk(pathDir):
        for sFile in lFiles:
            if sFile.lower() == sName.lower():
                return True
    return False


def checkBracesBalance(sCode, sName):
    """Braces must balance outside comments, strings and character literals."""
    i, n, iDepth = 0, len(sCode), 0
    while i < n:
        c = sCode[i]
        if c == "/" and i + 1 < n and sCode[i + 1] == "/":
            j = sCode.find("\n", i)
            i = n if j < 0 else j
            continue
        if c == "/" and i + 1 < n and sCode[i + 1] == "*":
            j = sCode.find("*/", i + 2)
            i = n if j < 0 else j + 2
            continue
        if c == "@" and i + 1 < n and sCode[i + 1] == '"':
            i += 2
            while i < n:
                if sCode[i] == '"':
                    if i + 1 < n and sCode[i + 1] == '"':
                        i += 2
                        continue
                    i += 1
                    break
                i += 1
            continue
        if c == '"':
            i += 1
            while i < n:
                if sCode[i] == "\\":
                    i += 2
                    continue
                if sCode[i] == '"':
                    i += 1
                    break
                i += 1
            continue
        if c == "'":
            i += 1
            while i < n:
                if sCode[i] == "\\":
                    i += 2
                    continue
                if sCode[i] == "'":
                    i += 1
                    break
                i += 1
            continue
        if c == "{":
            iDepth += 1
        elif c == "}":
            iDepth -= 1
        i += 1
    report(sName + " braces balance", iDepth == 0, "depth at end of file is " + str(iDepth))


def checkPowerShellBalance(sScript, sName):
    """Braces, parentheses and brackets must balance in a PowerShell file.

    PowerShell reports a mismatched bracket at the line where the parser
    finally gives up, which is often far from the edit that caused it,
    and a script that will not parse writes no log at all -- so the
    build fails with nothing to read. Counting here, outside comments
    and strings, names the depth at the end of the file before the
    build ever runs."""
    iBrace, iParen, iBracket = 0, 0, 0
    i, n = 0, len(sScript)
    while i < n:
        c = sScript[i]
        if c == "#" and (i == 0 or sScript[i - 1] != "$"):
            j = sScript.find("\n", i)
            i = n if j < 0 else j
            continue
        if c == "<" and i + 1 < n and sScript[i + 1] == "#":
            j = sScript.find("#>", i + 2)
            i = n if j < 0 else j + 2
            continue
        if c == "'":
            i += 1
            while i < n:
                if sScript[i] == "'":
                    if i + 1 < n and sScript[i + 1] == "'":
                        i += 2
                        continue
                    i += 1
                    break
                i += 1
            continue
        if c == '"':
            i += 1
            while i < n:
                if sScript[i] == "`":
                    i += 2
                    continue
                if sScript[i] == '"':
                    i += 1
                    break
                i += 1
            continue
        if c == "{": iBrace += 1
        elif c == "}": iBrace -= 1
        elif c == "(": iParen += 1
        elif c == ")": iParen -= 1
        elif c == "[": iBracket += 1
        elif c == "]": iBracket -= 1
        i += 1
    lProblems = []
    if iBrace != 0: lProblems.append("braces " + str(iBrace))
    if iParen != 0: lProblems.append("parentheses " + str(iParen))
    if iBracket != 0: lProblems.append("brackets " + str(iBracket))
    report(sName + " brackets balance", not lProblems,
           "; ".join(lProblems) if lProblems else "braces, parentheses and brackets all balance")


def checkSpellCheckInterfaces(sCode):
    """The spell checker's COM declarations must match Windows' own.

    A COM interface is a table of function pointers in a fixed order. A
    gap of the wrong size calls the wrong function, and nothing says so
    until the feature fails at run time -- which is exactly how the first
    version of this code shipped. The layouts below are from Microsoft's
    spellcheck.h; this compares them with what EdSharp declares."""
    dLayouts = {
        "ISpellCheckerFactory": ["get_SupportedLanguages", "IsSupported", "CreateSpellChecker"],
        "ISpellChecker": ["get_LanguageTag", "Check", "Suggest", "Add", "Ignore", "AutoCorrect",
                          "GetOptionValue", "get_OptionIds", "get_Id", "get_LocalizedName",
                          "add_SpellCheckerChanged", "remove_SpellCheckerChanged",
                          "GetOptionDescription", "ComprehensiveCheck"],
        "ISpellingError": ["get_StartIndex", "get_Length", "get_CorrectiveAction", "get_Replacement"],
    }
    dGuids = {
        "ISpellCheckerFactory": "8E018A9D-2415-4677-BF08-794EA61F94BB",
        "ISpellChecker": "B6FD0B71-E2BC-4653-8D05-F197E412770A",
        "IEnumSpellingError": "803E3BD4-2828-4410-8290-418D1D73C762",
        "ISpellingError": "B7C82D61-FBE8-4B47-9B27-6C0D2E0DE0A3",
    }
    lProblems = []
    for sName, sGuid in dGuids.items():
        if not re.search(r'Guid\("' + sGuid + r'"\)[^\n]*\]\s*\ninterface ' + sName + r'\b', sCode, re.I):
            lProblems.append(sName + " does not carry its documented interface identifier")
    for sName, lMethods in dLayouts.items():
        oBody = re.search(r"interface " + sName + r" \{(.*?)\n\} // " + sName, sCode, re.S)
        if not oBody:
            lProblems.append(sName + " was not found")
            continue
        iSlot = 0
        for sLine in oBody.group(1).split("\n"):
            sLine = sLine.split("//")[0].strip()
            if not sLine:
                continue
            oGap = re.match(r"void _VtblGap\d+_(\d+)\(\);", sLine)
            if oGap:
                iSlot += int(oGap.group(1))
                continue
            sLine = re.sub(r"^\[[^\]]*\]\s*", "", sLine)
            sLine = re.sub(r"\[[^\]]*\]\s*", "", sLine)
            oProperty = re.match(r"[\w\.<>\[\]]+ (\w+) \{ get; \}", sLine)
            oCall = re.search(r"\b(\w+)\s*\(", sLine)
            sHere = ("get_" + oProperty.group(1)) if oProperty else (oCall.group(1) if oCall else None)
            if sHere is None:
                continue
            if iSlot >= len(lMethods) or lMethods[iSlot] != sHere:
                lProblems.append(sName + " slot " + str(iSlot) + " declares " + sHere
                                 + " where Windows has " + (lMethods[iSlot] if iSlot < len(lMethods) else "nothing"))
            iSlot += 1
    report("Spell check interfaces match Windows' own layout", not lProblems,
           "; ".join(lProblems) if lProblems else plural(len(dLayouts), "interface") + " checked")


def normalizeKeyName(sKey):
    """One spelling for a key, so that Alt+D7 and Alt+7 compare equal.

    Hotkeys.ini spells keys the way a person says them, because its text
    is what the Key Describer reads aloud; the code spells them the way
    the framework names them. Neither is wrong, so both are reduced to a
    common form before comparing."""
    dSynonyms = {"back": "backspace", "oemquestion": "slash", "oemcomma": "comma",
                 "oemperiod": "period", "oemminus": "dash", "oemplus": "equals",
                 "oemsemicolon": "semicolon", "semi-colon": "semicolon",
                 "oemtilde": "backquote", "oemquotes": "quote", "oempipe": "backslash",
                 "oemopenbrackets": "leftbracket", "oem6": "rightbracket",
                 "oemclosebrackets": "rightbracket", "]": "rightbracket", "[": "leftbracket",
                 "oem5": "backslash", "\\": "backslash", "apostrophe": "quote",
                 "oem1": "semicolon", "oem2": "slash", "oem3": "backquote",
                 "oem4": "leftbracket", "oem7": "quote", "oem8": "backquote",
                 "rightarrow": "right", "leftarrow": "left", "uparrow": "up", "downarrow": "down"}
    lParts = []
    for sPart in sKey.split("+"):
        sPart = sPart.strip().lower().replace("&", "")
        if not sPart:
            continue
        if len(sPart) == 2 and sPart[0] == "d" and sPart[1].isdigit():
            sPart = sPart[1]
        lParts.append(dSynonyms.get(sPart, sPart))
    return "+".join(lParts)


def checkCommandsDescribed(sCode, sHotkeys):
    """Every command must have a description, and the right key with it.

    Key Describer mode reads these lines aloud. A command missing from
    Hotkeys.ini answers "No description available", which is exactly the
    moment a person wanted to be told something -- and a description
    naming the wrong key is worse than none."""
    # The built-in table in EdSharp.cs is the source of truth, so it is
    # checked first; Hotkeys.ini is a supplement anyone may edit.
    dDescribed = {}
    oTable = re.search(r"c_aCommandSummaries = new string\[\] \{\n(.*?)\n\}; // c_aCommandSummaries", sCode, re.S)
    if oTable:
        for sRow in re.findall(r'^"((?:[^"\\]|\\.)*)",?$', oTable.group(1), re.M):
            sRow = sRow.replace('\\"', '"').replace("\\\\", "\\")
            if "\\t" not in sRow:
                continue
            sName, sValue = sRow.split("\\t", 1)
            dDescribed[sName] = sValue
    else:
        lFailures.append("the built-in description table was not found in EdSharp.cs")
    for sLine in sHotkeys.replace("\r\n", "\n").split("\n"):
        sLine = sLine.strip()
        if "=" not in sLine or sLine.startswith(";") or sLine.startswith("["):
            continue
        sName, sValue = sLine.split("=", 1)
        if sName.strip() not in dDescribed:
            dDescribed[sName.strip()] = sValue.strip()
    lMissing = []
    lWrongKey = []
    lCommands = re.findall(r'CreateMenuItem\(\s*"([^"]*)"\s*,\s*"([^"]*)"', sCode)
    for sText, sKey in lCommands:
        sName = sText.replace("&", "").replace("...", "").strip()
        sValue = dDescribed.get(sName, dDescribed.get("Say " + sName, ""))
        if sValue == "":
            lMissing.append(sName)
            continue
        sDescribedKey = sValue.split(",")[0].strip()
        if sKey and sDescribedKey and normalizeKeyName(sDescribedKey) != normalizeKeyName(sKey):
            lWrongKey.append(sName + " is bound to " + sKey + " but described as " + sDescribedKey)
    report("Every command has a description in the program", not lMissing,
           "; ".join(lMissing) if lMissing else plural(len(lCommands), "command") + " checked")
    report("Descriptions name the right key", not lWrongKey,
           "; ".join(lWrongKey) if lWrongKey else "keys agree with the bindings")


def checkInterfaceAccessibility(sCode):
    """A public method may not return or accept a private type.

    The COM interfaces are declared inside the frame class with no
    modifier, which makes them private; a public method handing one back
    fails to compile with "inconsistent accessibility". That cost a build
    on 26 August 2026, so it is checked here where it costs a second."""
    lPrivateTypes = re.findall(r"^interface (\w+) \{", sCode, re.M)
    lProblems = []
    for oMatch in re.finditer(r"^public\s+([\w<>\[\], .]+?)\s+(\w+)\s*\(([^)]*)\)\s*\{", sCode, re.M):
        sReturn, sName, sArguments = oMatch.groups()
        lTypes = set(re.findall(r"\b(\w+)\b", sReturn + " " + sArguments))
        lExposed = sorted(lTypes.intersection(lPrivateTypes))
        if lExposed:
            lProblems.append(sName + " exposes " + ", ".join(lExposed))
    report("Public methods do not expose private interfaces", not lProblems,
           "; ".join(lProblems) if lProblems else plural(len(lPrivateTypes), "interface") + " checked")


def checkOfficeDependencies(sCode):
    """Report which features still reach for Microsoft Office, by name."""
    lUses = []
    for oMatch in re.finditer(r'COM\.WordAccess', sCode):
        iLine = sCode.count("\n", 0, oMatch.start()) + 1
        sBefore = sCode[:oMatch.start()]
        lMethods = re.findall(r'public\s+(?:static\s+)?[\w<>\[\]]+\s+(\w+)\s*\(', sBefore)
        lUses.append((lMethods[-1] if lMethods else "?") + " (line " + str(iLine) + ")")
    # Word is expected only in the named fallbacks and legacy readers.
    lAllowed = ("SpellCheckWord", "ThesaurusWord", "MailWord", "WordFile2String", "WordSource2TargetFormat")
    lUnexpected = [s for s in lUses if not s.split(" ")[0] in lAllowed]
    report("Microsoft Office is used only by named fallbacks", not lUnexpected,
           "; ".join(lUnexpected) if lUnexpected else "; ".join(lUses) if lUses else "no uses found")


def checkOptionsDocumented(sIni, sDoc):
    """Every option EdSharp ships should be mentioned in the guide."""
    lUndocumented = []
    lChecked = []
    for oMatch in re.finditer(r'^([A-Za-z][A-Za-z0-9]*)="', sIni, re.M):
        sOption = oMatch.group(1)
        if sOption in lChecked:
            continue
        lChecked.append(sOption)
        if sOption not in sDoc:
            lUndocumented.append(sOption)
    # Documentation completeness is a goal rather than an invariant: an
    # undocumented option is worth knowing about, not worth stopping a
    # build for. It is reported as a note.
    if lUndocumented:
        say("NOTE  options not mentioned in the guide: " + ", ".join(lUndocumented))
    report("Shipped options were checked against the guide", True,
           plural(len(lChecked), "option") + " checked, " + str(len(lUndocumented)) + " not yet in the guide")


def main():
    global fileLog
    fileLog = open(pathLog, "w", encoding="utf-8")
    say("EdSharp audit " + datetime.datetime.now().isoformat(" ", "seconds"))
    say("Folder: " + pathRoot)
    say("Python: " + sys.version.split()[0])
    say()

    sCode = readFile("EdSharp.cs")
    sInix = readFile("EdSharp.inix")
    sIni = readFile("EdSharp.ini")
    sDoc = readFile("EdSharp.md")
    sInix2 = readFile("Inix.cs")

    if sCode is None:
        report("EdSharp.cs is present", False, "not found in " + pathRoot)
    else:
        checkBracesBalance(sCode, "EdSharp.cs")
        checkShortcutsUnique(sCode)
        checkAccessKeysUnique(sCode)
        checkOfficeDependencies(sCode)
        checkSpellCheckInterfaces(sCode)
        checkInterfaceAccessibility(sCode)
        sHotkeys = readFile("Hotkeys.ini")
        if sHotkeys is None:
            report("Hotkeys.ini is present", False, "not found in " + pathRoot)
        else:
            checkCommandsDescribed(sCode, sHotkeys)
    if sInix2 is not None:
        checkBracesBalance(sInix2, "Inix.cs")
    for sScriptName in ("BuildEdSharp.ps1", "summarizeSetup.ps1", "installJawsScripts.ps1"):
        sScript = readFile(sScriptName)
        if sScript is not None:
            checkPowerShellBalance(sScript, sScriptName)
    if sInix is not None:
        checkRegexesCompile(sInix)
        checkConversionScriptsExist(sInix)
    else:
        report("EdSharp.inix is present", False, "not found")
    if sIni is not None and sDoc is not None:
        checkOptionsDocumented(sIni, sDoc)

    say()
    if lFailures:
        say(plural(len(lFailures), "check") + " failed: " + ", ".join(lFailures))
    else:
        say("All checks passed.")
    say("Log: " + pathLog)
    fileLog.close()
    return 1 if lFailures else 0


if __name__ == "__main__":
    sys.exit(main())
