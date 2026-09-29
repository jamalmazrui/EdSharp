#!/usr/bin/env python3
"""fixEncoding.py -- put the project's own text files into the Homer encoding.

    fixEncoding            every file RepoFiles.txt names, in the project
    fixEncoding --check    report only, change nothing (exit 1 if any is wrong)

THE HOMER ENCODING is UTF-8 with a byte order mark and CRLF line endings,
except that .cmd and .bat files take CRLF and no mark. Pandoc writes UTF-8
with no mark and bare newlines, so every .htm a build makes -- and every one
delivered from elsewhere -- is wrong until something fixes it. On 25 Sep 2026
a release was refused because eighteen delivered .htm files had never been
put right. Every Homer build runs this over its own files, so the check that
follows has nothing to find.

OTHER PEOPLE'S FILES KEEP THEIR ENCODING (1.43.11). A project may carry
third-party tools, dictionaries, or files that demonstrate another encoding;
a byte order mark on a Lua filter or a tool's config file is read as part of
its first line. KeepEncoding.txt, beside RepoFiles.txt and in the same form --
a folder ending in /, a file, or a pattern with * -- names what this tool and
check leave exactly as it is.

WHICH FILES: the ones RepoFiles.txt names (a folder named there means its
text files), because those are the project's own. A stray file in the folder
is tidy's business, not this tool's, and is left as it is. Without a
RepoFiles.txt the tool does nothing and says so.

Run it in the project folder or its scripts folder; both mean the project.
The log is logs\\<App>-encoding-yyyyMMdd-HHmmss.log; the console says how
many files were fixed.
"""

import datetime, glob, io, os, platform, re, sys

c_lsTextExt = (".cmd", ".bat", ".cs", ".htm", ".html", ".inix", ".iss", ".json", ".lua", ".md",
               ".ps1", ".py", ".spec", ".txt", ".xml", ".m3u")
c_lsNoBom = (".cmd", ".bat")
# exec is walked, not pruned (1.43.22): the kit keeps its libraries in
# exec\\CSharp and exec\\homer, and only what RepoFiles.txt names is touched, so
# a build's binaries there are left alone either way.
c_lsSkipFolders = ("logs", "notes", ".git", "packages", "work", "__pycache__")

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




def projectRoot(sStart):
    if os.path.basename(sStart).lower() in ("scripts", "exec", "tools"):
        sParent = os.path.dirname(sStart)
        if (os.path.isfile(os.path.join(sParent, "version.txt")) or os.path.isfile(os.path.join(sParent, "RepoFiles.txt"))
                or glob.glob(os.path.join(sParent, "*_setup.iss"))):
            return sParent
    return sStart


def readNamed(sRoot):
    """The entries of RepoFiles.txt: files, folders (ending in /), patterns."""
    sPath = os.path.join(sRoot, "RepoFiles.txt")
    if not os.path.isfile(sPath): return None
    lsNamed = []
    with io.open(sPath, encoding="utf-8-sig") as oFile:
        for sLine in oFile:
            sLine = sLine.strip()
            if not sLine or sLine.startswith("#"): continue
            lsNamed.append(sLine.replace("\\", "/"))
    return lsNamed


def readKeep(sRoot):
    """The entries of KeepEncoding.txt, lower case with forward slashes."""
    sPath = os.path.join(sRoot, "KeepEncoding.txt")
    if not os.path.isfile(sPath): return []
    lsKeep = []
    with io.open(sPath, encoding="utf-8-sig") as oFile:
        for sLine in oFile:
            sLine = sLine.strip()
            if not sLine or sLine.startswith("#") or sLine.startswith(";"): continue
            lsKeep.append(sLine.replace("\\", "/").lower().lstrip("/"))
    return lsKeep


def isKept(sRoot, sPath, lsKeep):
    """Does KeepEncoding.txt name this file, its folder, or a pattern it fits?"""
    sRel = os.path.relpath(sPath, sRoot).replace(os.sep, "/").lower()
    for sKeep in lsKeep:
        if sKeep == sRel: return True
        if sKeep.endswith("/") and sRel.startswith(sKeep): return True
        if "*" in sKeep and re.match("^" + re.escape(sKeep).replace(r"\*", ".*") + "$", sRel): return True
    return False


def ownFiles(sRoot, lsNamed):
    """Every text file the whitelist covers, as absolute paths, less what
    KeepEncoding.txt keeps as it is."""
    lsKeep = readKeep(sRoot)
    lsFiles = []
    for sNamed in lsNamed:
        if sNamed.endswith("/"):
            sFolder = os.path.join(sRoot, sNamed.rstrip("/").replace("/", os.sep))
            for sDirPath, lsDirs, lsNames in os.walk(sFolder):
                lsDirs[:] = [s for s in lsDirs if s.lower() not in c_lsSkipFolders]
                for sName in lsNames:
                    if sName.lower().endswith(c_lsTextExt): lsFiles.append(os.path.join(sDirPath, sName))
        elif "*" in sNamed:
            for sPath in glob.glob(os.path.join(sRoot, sNamed.replace("/", os.sep))):
                if os.path.isfile(sPath) and sPath.lower().endswith(c_lsTextExt): lsFiles.append(sPath)
        else:
            sPath = os.path.join(sRoot, sNamed.replace("/", os.sep))
            if os.path.isfile(sPath) and sPath.lower().endswith(c_lsTextExt): lsFiles.append(sPath)
    lsFiles = [sPath for sPath in lsFiles if not isKept(sRoot, sPath, lsKeep)]
    lsSeen = []
    for sPath in lsFiles:
        sKey = os.path.normcase(os.path.abspath(sPath))
        if sKey not in lsSeen: lsSeen.append(sKey)
    return sorted(lsSeen)


def wantedBytes(sPath, binData):
    """What the file should hold; None when it is not text we can decode."""
    try:
        sText = binData.decode("utf-8-sig")
    except UnicodeDecodeError:
        return None
    sText = sText.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    bBom = not sPath.lower().endswith(c_lsNoBom) and os.path.basename(sPath).lower() != "version.txt"
    return (("\ufeff" + sText) if bBom else sText).encode("utf-8")


def main():
    global oLog
    bCheck = "--check" in sys.argv
    sRoot = projectRoot(os.getcwd())
    sLogDir = os.path.join(sRoot, "logs")
    os.makedirs(sLogDir, exist_ok=True)
    sLogPath = os.path.join(sLogDir, "%s-encoding-%s.log" % (os.path.basename(sRoot), datetime.datetime.now().strftime("%Y%m%d-%H%M%S")))
    oLog = io.open(sLogPath, "w", encoding="utf-8")
    logLine("fixEncoding start pid=%d" % os.getpid())
    logFact("script", os.path.abspath(__file__))
    logFact("python", platform.python_version())
    logFact("windows", logWindows())
    logFact("project", sRoot)
    logFact("arguments", " ".join(sys.argv[1:]))
    logFact("mode", "check" if bCheck else "fix")
    lsNamed = readNamed(sRoot)
    if lsNamed is None:
        print("No RepoFiles.txt here, so nothing names the project's own files. Nothing was changed.")
        logLine("NO RepoFiles.txt")
        return 0
    lsFiles = ownFiles(sRoot, lsNamed)
    logLine("Named entries: %d; text files covered: %d" % (len(lsNamed), len(lsFiles)))
    iWrong = 0
    iFixed = 0
    for sPath in lsFiles:
        try:
            binData = open(sPath, "rb").read()
        except Exception as oError:
            logLine("COULD NOT READ %s: %s" % (sPath, oError))
            continue
        if not binData: continue
        binWanted = wantedBytes(sPath, binData)
        if binWanted is None:
            logLine("NOT UTF-8, left alone: %s" % os.path.relpath(sPath, sRoot))
            continue
        if binWanted == binData: continue
        iWrong += 1
        sShown = os.path.relpath(sPath, sRoot)
        if bCheck:
            logLine("WRONG: %s" % sShown)
            continue
        try:
            open(sPath, "wb").write(binWanted)
            iFixed += 1
            logLine("FIXED: %s" % sShown)
        except Exception as oError:
            logLine("COULD NOT WRITE %s: %s" % (sPath, oError))
    if bCheck:
        print("%d file%s checked, %d wrong." % (len(lsFiles), "" if len(lsFiles) == 1 else "s", iWrong))
        return 1 if iWrong else 0
    print("%d file%s checked, %d fixed." % (len(lsFiles), "" if len(lsFiles) == 1 else "s", iFixed))
    logLine("fixEncoding end")
    return 0


if __name__ == "__main__":
    sys.exit(main())
