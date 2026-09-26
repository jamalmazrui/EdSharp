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

WHICH FILES: the ones RepoFiles.txt names (a folder named there means its
text files), because those are the project's own. A stray file in the folder
is homerTidy's business, not this tool's, and is left as it is. Without a
RepoFiles.txt the tool does nothing and says so.

Run it in the project folder or its scripts folder; both mean the project.
The log is logs\\<App>-encoding-yyyyMMdd-HHmmss.log; the console says how
many files were fixed.
"""

import datetime, glob, io, os, platform, re, sys

c_lsTextExt = (".cmd", ".bat", ".cs", ".htm", ".html", ".inix", ".iss", ".json", ".lua", ".md",
               ".ps1", ".py", ".spec", ".txt", ".xml", ".m3u")
c_lsNoBom = (".cmd", ".bat")
c_lsSkipFolders = ("logs", "notes", "exec", ".git", "packages", "work", "__pycache__")

oLog = None


def logLine(sText):
    if oLog is not None:
        oLog.write(sText + "\n")
        oLog.flush()
    return True


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


def ownFiles(sRoot, lsNamed):
    """Every text file the whitelist covers, as absolute paths."""
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
    logLine("fixEncoding started %s" % datetime.datetime.now().isoformat(" ", "seconds"))
    logLine("Script: %s" % os.path.abspath(__file__))
    logLine("Python: %s" % sys.version.replace("\n", " "))
    logLine("Platform: %s" % platform.platform())
    logLine("Project: %s" % sRoot)
    logLine("Command line: %s" % " ".join(sys.argv))
    logLine("Mode: %s" % ("check only" if bCheck else "fix"))
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
    logLine("Finished %s" % datetime.datetime.now().isoformat(" ", "seconds"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
