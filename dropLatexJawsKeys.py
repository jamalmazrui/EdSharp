r"""dropLatexJawsKeys.py -- remove the obsolete Process LaTeX feature from
the installed JAWS scripts for EdSharp, freeing F12 for Chat with AI.

Run it from C:\EdSharp with no arguments, or use dropLatexJawsKeys.cmd:

    python dropLatexJawsKeys.py

WHAT IT DOES

EdSharp once had a Process LaTeX feature, and the JAWS scripts bound it
to F12 and possibly other F12 combinations. Those bindings now swallow
the F12 that EdSharp's Chat with AI command uses. Cleaning only the
installed copies is not enough: the installer copies the Scripts folder
into every JAWS Settings folder on each setup run, so a dirty source
reinstalls the problem. This script therefore cleans, in order: the
repository source Scripts folder beside this script, the installed
program's Scripts folder under Program Files when it is writable, and
then every JAWS version's user settings folder, where it also:

  1. EdSharp.jkm: removes every key binding whose script name mentions
     LaTeX (any spelling), and any F12-family binding pointing at such
     a script.
  2. EdSharp.jss: removes every Script block whose name mentions LaTeX,
     from its Script line through its EndScript line.
  3. Recompiles EdSharp.jss with the newest scompile.exe found, so the
     change takes effect the next time JAWS loads the scripts.

Backups with a .bak extension sit beside every changed file, so one
copy undoes everything. The script is safe to re-run: with nothing
left to remove, it says so and changes nothing. A detailed log is
written beside this script, whatever happens.
"""

import datetime
import glob
import os
import re
import subprocess
import sys

c_sLogName = "dropLatexJawsKeys.log"

pathRoot = os.path.dirname(os.path.abspath(__file__))
pathLog = os.path.join(pathRoot, c_sLogName)
fileLog = None


def say(sMessage=""):
    """Print and log. Every line the user sees is also in the log."""
    print(sMessage)
    if fileLog:
        try:
            fileLog.write(sMessage + "\n")
            fileLog.flush()
        except Exception:
            pass


def isLatexName(sName):
    return "latex" in sName.lower() or "latec" in sName.lower()


def isF12Key(sKey):
    return "f12" in sKey.lower()


def cleanJkm(pathFile):
    """Remove LaTeX bindings from a JAWS key map. Returns lines removed."""
    with open(pathFile, "r", encoding="utf-8", errors="replace") as f:
        lLines = f.read().splitlines()
    lKept = []
    iRemoved = 0
    for sLine in lLines:
        sTrim = sLine.strip()
        if "=" in sTrim and not sTrim.startswith(";") and not sTrim.startswith("["):
            sKey, sScript = sTrim.split("=", 1)
            if isLatexName(sScript) or (isF12Key(sKey) and isLatexName(sScript)):
                iRemoved += 1
                say("  removed binding: " + sTrim)
                continue
        lKept.append(sLine)
    if iRemoved:
        with open(pathFile + ".bak", "w", encoding="utf-8") as f:
            f.write("\n".join(lLines) + "\n")
        with open(pathFile, "w", encoding="utf-8") as f:
            f.write("\n".join(lKept) + "\n")
    return iRemoved


def cleanJss(pathFile):
    """Remove LaTeX Script blocks from a JAWS script source. Returns blocks removed."""
    with open(pathFile, "r", encoding="utf-8", errors="replace") as f:
        lLines = f.read().splitlines()
    lKept = []
    iRemoved = 0
    bSkipping = False
    for sLine in lLines:
        sTrim = sLine.strip()
        if not bSkipping:
            match = re.match(r"(?i)Script\s+(\w+)", sTrim)
            if match and isLatexName(match.group(1)):
                bSkipping = True
                iRemoved += 1
                say("  removed script block: " + match.group(1))
                continue
            lKept.append(sLine)
        else:
            if re.match(r"(?i)EndScript\b", sTrim):
                bSkipping = False
            continue
    if iRemoved:
        with open(pathFile + ".bak", "w", encoding="utf-8") as f:
            f.write("\n".join(lLines) + "\n")
        with open(pathFile, "w", encoding="utf-8") as f:
            f.write("\n".join(lKept) + "\n")
    return iRemoved


def findScompile():
    """The newest scompile.exe among installed JAWS versions, or empty."""
    lCandidates = []
    for sRoot in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                  os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
        lCandidates += glob.glob(os.path.join(sRoot, "Freedom Scientific", "JAWS", "*", "scompile.exe"))
    if not lCandidates:
        return ""
    def versionKey(sPath):
        match = re.search(r"JAWS[\\/](\d+(?:\.\d+)?)", sPath)
        return float(match.group(1)) if match else 0.0
    return max(lCandidates, key=versionKey)


def plural(iCount, sNoun):
    return str(iCount) + " " + sNoun + ("" if iCount == 1 else "s")


def main():
    global fileLog
    fileLog = open(pathLog, "w", encoding="utf-8")
    say("dropLatexJawsKeys started " + datetime.datetime.now().isoformat(" ", "seconds"))
    say("Script: " + os.path.abspath(__file__))
    say("Python: " + sys.version.split()[0] + ", platform: " + sys.platform)
    say("Working directory: " + os.getcwd())
    say()

    iBindings = 0
    iBlocks = 0
    lRecompile = []

    # Phase one: the source Scripts folders, so no future install can
    # bring the feature back. The repository copy beside this script is
    # the master; the Program Files copy is what installJawsScripts
    # reads on the next setup run.
    lSourceDirs = [os.path.join(pathRoot, "Scripts"),
                   os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), "EdSharp", "Scripts")]
    for pathSource in lSourceDirs:
        if not os.path.isdir(pathSource):
            say("No Scripts folder at " + pathSource)
            continue
        say(pathSource)
        for sName in sorted(os.listdir(pathSource)):
            pathFile = os.path.join(pathSource, sName)
            if not os.path.isfile(pathFile):
                continue
            try:
                if sName.lower().endswith(".jkm"):
                    iBindings += cleanJkm(pathFile)
                elif sName.lower().endswith(".jss"):
                    iBlocks += cleanJss(pathFile)
            except PermissionError:
                say("  NO PERMISSION to change " + pathFile)
                say("  Run this script from an administrator command window to clean it,")
                say("  or rebuild and reinstall after committing the cleaned repository copy.")

    # Phase two: the installed user settings, where JAWS actually loads.
    pathAppData = os.environ.get("APPDATA", "")
    lFolders = sorted(glob.glob(os.path.join(pathAppData, "Freedom Scientific", "JAWS", "*", "Settings", "enu")))
    say()
    say("JAWS settings folders found: " + str(len(lFolders)))
    for pathFolder in lFolders:
        say(pathFolder)
        pathJkm = os.path.join(pathFolder, "EdSharp.jkm")
        pathJss = os.path.join(pathFolder, "EdSharp.jss")
        if os.path.isfile(pathJkm):
            iBindings += cleanJkm(pathJkm)
        else:
            say("  no EdSharp.jkm here")
        if os.path.isfile(pathJss):
            iRemoved = cleanJss(pathJss)
            iBlocks += iRemoved
            if iRemoved:
                lRecompile.append(pathJss)
        else:
            say("  no EdSharp.jss here")
    say()
    say(plural(iBindings, "binding") + " removed; " + plural(iBlocks, "script block") + " removed.")

    if lRecompile:
        pathScompile = findScompile()
        if pathScompile:
            say("Compiler: " + pathScompile)
            for pathJss in lRecompile:
                say("Compiling " + pathJss)
                oResult = subprocess.run([pathScompile, pathJss], capture_output=True, text=True)
                say("  exit code " + str(oResult.returncode))
                if oResult.stdout.strip():
                    say("  " + oResult.stdout.strip())
                if oResult.stderr.strip():
                    say("  " + oResult.stderr.strip())
        else:
            say("No scompile.exe was found under Program Files; open each changed")
            say("EdSharp.jss in the JAWS Script Manager and press Control+S to compile.")
    elif iBindings:
        say("Only key map changes were made; no compile is needed for those.")
    else:
        say("Nothing mentioned LaTeX; the scripts were already clean.")
    say()
    say("Restart JAWS, or switch away from EdSharp and back, to load the change.")
    say("Log: " + pathLog)
    fileLog.close()


if __name__ == "__main__":
    try:
        main()
    except Exception as oError:
        say("FAILED: " + repr(oError))
        if fileLog:
            fileLog.close()
        sys.exit(1)
