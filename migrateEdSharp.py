r"""migrateEdSharp.py -- carry C:\EdSharp from its old shape to the Homer one.

    migrateEdSharp.cmd                move and retire
    migrateEdSharp.cmd --survey       print the plan and change nothing

WHY A SCRIPT AND NOT JUST THE ZIP

Unzipping the new EdSharp.zip puts every new and changed file in place, but
an archive cannot move a folder you already have, and cannot delete what is
no longer wanted. That is this script's whole job, and it is meant to be run
once.

WHAT IT DOES

Moves the content folders to where the program now looks for them:

    Convert       -> configs\convert
    Dictionaries  -> data\dictionaries
    Samples       -> templates\samples
    Snippets      -> templates\snippets

Moves the JAWS scripts out of scripts\ into scripts\jaws, leaving scripts\
for the kit's tools and the install scripts, which is what every Homer app
uses it for.

Moves the built program, its libraries and the installer into exec\, which
is where the build writes them now and the only place they belong.

Retires EdSharp's own earlier editions of what the kit now does --
tidyRepo, repoPolicy, moveNotes, restoreMissing, prepareAuditFixes,
applyConvertPolicy, dropLatexJawsKeys, auditEdSharp, summarizeSetup,
ModernizePandocConfig, the root tagRelease, and the old BuildEdSharp pair --
and the copies of the shared Homer classes, which had all drifted from the
kit's. Retired files go to notes\ rather than being deleted, so anything
still wanted can be fetched back; the class copies go for good, because
keeping one is how a build quietly compiles the wrong Lbc.

A detailed log is written in logs\, whatever happens.
"""

import argparse
import datetime
import os
import shutil
import sys

c_sApp = "EdSharp"

# Folder moves: from, to. The destination's parent is made as needed.
c_lFolderMoves = [
    ("Convert",      os.path.join("configs", "convert")),
    ("Dictionaries", os.path.join("data", "dictionaries")),
    ("Samples",      os.path.join("templates", "samples")),
    ("Snippets",     os.path.join("templates", "snippets")),
]

# Documents that move into help, with their .htm partners.
c_lHelpStems = ["Announce", "EdSharp", "FAQ", "History", "Hotkeys", "Tutorials"]
c_lHelpFiles = ["history.txt", "lgpl.txt"]

# Settings that move into configs.
c_lConfigFiles = ["EdSharp.ini", "EdSharp.inix", "Hotkeys.ini", "Tools.inix"]

# Build products, which belong in exec and nowhere else.
c_lExecFiles = [
    "EdSharp.dll", "EdSharp.exe", "EdSharp_Setup.exe", "EdSharp_setup.exe",
    "HtmlAgilityPack.dll", "Markdig.dll", "ReverseMarkdown.dll",
    "Tektosyne.dll", "Ude.dll", "WeCantSpell.Hunspell.dll",
    "nvdaControllerClient.dll", "sqlean.dll", "sqlean.exe",
    "EdSharp.nvda-addon",
]

# The shared classes. Deleted outright: a copy left here is how a build
# silently compiles a different Lbc from the one every other app has.
c_lClassCopies = ["Inix.cs", "KeyMap.cs", "Lbc.cs", "Say.cs", "Web.cs", "inixVert.cs"]

# Earlier editions of what the kit now does, and the odds and ends the new
# layout leaves behind. To notes, not deleted.
c_lRetired = [
    "BuildEdSharp.cmd", "BuildEdSharp.ps1",
    "CamelType_CSharp.htm", "CamelType_CSharp.md",
    "CamelType_JAWSScript.htm", "CamelType_JAWSScript.md",
    "Development.htm", "Development.md",
    "FetchConvertTools.ps1", "FetchUde.ps1",
    "ModernizePandocConfig.ps1",
    "Transform_Example.inix",
    "Tutorial.htm", "Tutorial.md",
    "applyConvertPolicy.cmd", "applyConvertPolicy.py",
    "auditEdSharp.cmd", "auditEdSharp.log", "auditEdSharp.py",
    "dropLatexJawsKeys.cmd", "dropLatexJawsKeys.py",
    "moveNotes.cmd", "moveNotes.log", "moveNotes.py",
    "prepareAuditFixes.cmd",
    "repoPolicy.py",
    "restoreMissing.cmd", "restoreMissing.log", "restoreMissing.py",
    "sqlean.version",
    "summarizeSetup.cmd", "summarizeSetup.ps1",
    "tagRelease.cmd", "tagRelease.ps1",
    "tidyRepo.cmd", "tidyRepo.log", "tidyRepo.py",
]

# Install scripts, which the kit keeps in scripts and the installer ships
# from there.
c_lInstallScripts = [
    "installCodeModel.cmd", "installGitHub.cmd", "installJawsScripts.cmd",
    "installJawsScripts.ps1", "installNode.cmd", "installOllama.cmd",
    "installPandoc.cmd", "installPandoc.ps1", "installPdfTools.cmd",
    "installPython.cmd", "installTranslateModel.cmd",
]

pathRoot = os.path.dirname(os.path.abspath(__file__))
fileLog = None
pathLog = ""


def say(sMessage=""):
    print(sMessage)
    if fileLog:
        try:
            fileLog.write(sMessage + "\n")
            fileLog.flush()
        except Exception:
            pass


def startLog():
    global fileLog, pathLog
    pathLogs = os.path.join(pathRoot, "logs")
    try:
        os.makedirs(pathLogs, exist_ok=True)
        pathLog = os.path.join(pathLogs, c_sApp + "-migrate-"
                               + datetime.datetime.now().strftime("%Y%m%d-%H%M%S") + ".log")
        fileLog = open(pathLog, "w", encoding="utf-8")
    except Exception as oError:
        print("Could not open the log: " + str(oError))
        return
    say(c_sApp + " migrate  " + datetime.datetime.now().isoformat(" ", "seconds"))
    say("  script:            " + os.path.abspath(__file__))
    say("  Python:            " + sys.version.split()[0])
    say("  platform:          " + sys.platform)
    say("  working directory: " + os.getcwd())
    say("  command line:      " + " ".join([os.path.basename(sys.argv[0])] + sys.argv[1:]))
    say("  project folder:    " + pathRoot)
    say()


def plural(iCount, sNoun):
    return f"{iCount} {sNoun}" + ("" if iCount == 1 else "s")


def freeName(pathFolder, sName):
    """A name in a folder that nothing occupies yet."""
    if not os.path.exists(os.path.join(pathFolder, sName)):
        return sName
    sStem, sExtension = os.path.splitext(sName)
    for i in range(1, 100):
        sTry = f"{sStem}-{i:02d}{sExtension}"
        if not os.path.exists(os.path.join(pathFolder, sTry)):
            return sTry
    return None


def moveFile(sName, sIntoFolder, bDo):
    """Move one file into a folder under the project, making it as needed."""
    pathFrom = os.path.join(pathRoot, sName)
    if not os.path.isfile(pathFrom):
        return False
    pathInto = os.path.join(pathRoot, sIntoFolder)
    if not bDo:
        say(f"  {sName} -> {sIntoFolder}\\")
        return True
    os.makedirs(pathInto, exist_ok=True)
    sTarget = freeName(pathInto, os.path.basename(sName))
    if sTarget is None:
        say(f"  COULD NOT PLACE {sName} in {sIntoFolder}")
        return False
    try:
        shutil.move(pathFrom, os.path.join(pathInto, sTarget))
    except Exception as oError:
        say(f"  COULD NOT MOVE {sName}: {oError}")
        return False
    say(f"  moved {sName} -> {sIntoFolder}\\{sTarget}")
    return True


def moveFolder(sFrom, sTo, bDo):
    """Move one folder to its new place, merging into an existing one."""
    pathFrom = os.path.join(pathRoot, sFrom)
    pathTo = os.path.join(pathRoot, sTo)
    if not os.path.isdir(pathFrom):
        return False
    if os.path.abspath(pathFrom) == os.path.abspath(pathTo):
        return False
    if not bDo:
        iHeld = sum(len(f) for _, _, f in os.walk(pathFrom))
        say(f"  {sFrom}\\ -> {sTo}\\  ({plural(iHeld, 'file')})")
        return True
    try:
        os.makedirs(os.path.dirname(pathTo), exist_ok=True)
        if os.path.isdir(pathTo):
            # Merge rather than refuse: a second run, or a folder the zip
            # already created, should not stop the move.
            for sName in os.listdir(pathFrom):
                shutil.move(os.path.join(pathFrom, sName), os.path.join(pathTo, sName))
            os.rmdir(pathFrom)
        else:
            shutil.move(pathFrom, pathTo)
    except Exception as oError:
        say(f"  COULD NOT MOVE the folder {sFrom}: {oError}")
        return False
    say(f"  moved folder {sFrom}\\ -> {sTo}\\")
    return True


def deleteFile(sName, bDo):
    pathFile = os.path.join(pathRoot, sName)
    if not os.path.isfile(pathFile):
        return False
    if not bDo:
        say(f"  {sName} (deleted)")
        return True
    try:
        os.chmod(pathFile, 0o666)
    except OSError:
        pass
    try:
        os.remove(pathFile)
    except OSError as oError:
        say(f"  COULD NOT DELETE {sName}: {oError}")
        return False
    say(f"  deleted {sName}")
    return True


def jawsScriptNames():
    """The screen reader scripts sitting directly in scripts."""
    pathScripts = os.path.join(pathRoot, "scripts")
    if not os.path.isdir(pathScripts):
        return []
    lNames = []
    for sName in sorted(os.listdir(pathScripts), key=str.lower):
        if not os.path.isfile(os.path.join(pathScripts, sName)):
            continue
        if os.path.splitext(sName)[1].lower() in (
                ".jss", ".jsd", ".jsh", ".jkm", ".jcf", ".jsb", ".jbs"):
            lNames.append(sName)
        elif sName.lower().endswith("_scripts_setup.iss"):
            lNames.append(sName)
    return lNames


def main():
    oParser = argparse.ArgumentParser(
        description="Move EdSharp to the Homer folder layout.")
    oParser.add_argument("--survey", action="store_true",
                         help="print the plan and change nothing")
    oArguments = oParser.parse_args()
    bDo = not oArguments.survey

    startLog()
    if not os.path.isfile(os.path.join(pathRoot, "EdSharp.cs")):
        say("This does not look like the EdSharp folder: there is no EdSharp.cs here.")
        say("Run it from C:\\EdSharp.")
        return 1

    say("=" * 68)
    say("PLAN" if not bDo else "DOING IT")
    say("=" * 68)
    say()

    say("Content folders:")
    iFolders = sum(1 for sFrom, sTo in c_lFolderMoves if moveFolder(sFrom, sTo, bDo))
    if iFolders == 0:
        say("  already moved")
    say()

    say("Documents into help:")
    iHelp = 0
    for sStem in c_lHelpStems:
        for sExtension in (".md", ".htm"):
            iHelp += 1 if moveFile(sStem + sExtension, "help", bDo) else 0
    for sName in c_lHelpFiles:
        iHelp += 1 if moveFile(sName, "help", bDo) else 0
    if iHelp == 0:
        say("  already moved")
    say()

    say("Settings into configs:")
    iConfigs = sum(1 for sName in c_lConfigFiles if moveFile(sName, "configs", bDo))
    if iConfigs == 0:
        say("  already moved")
    say()

    say("Build products into exec:")
    iExec = sum(1 for sName in c_lExecFiles if moveFile(sName, "exec", bDo))
    if iExec == 0:
        say("  already moved or not built yet")
    say()

    say("Install scripts into scripts:")
    iInstall = sum(1 for sName in c_lInstallScripts if moveFile(sName, "scripts", bDo))
    if iInstall == 0:
        say("  already moved")
    say()

    say("Screen reader scripts into scripts\\jaws:")
    iJaws = 0
    for sName in jawsScriptNames():
        iJaws += 1 if moveFile(os.path.join("scripts", sName), os.path.join("scripts", "jaws"), bDo) else 0
    if iJaws == 0:
        say("  already moved")
    say()

    say("Earlier editions of what the kit now does, into notes:")
    iRetired = sum(1 for sName in c_lRetired if moveFile(sName, "notes", bDo))
    if iRetired == 0:
        say("  already retired")
    say()

    say("Copies of the shared Homer classes, deleted:")
    iClasses = sum(1 for sName in c_lClassCopies if deleteFile(sName, bDo))
    if iClasses == 0:
        say("  none left")
    say()

    say("=" * 68)
    if not bDo:
        say("This was a description only (--survey). Nothing has been changed.")
        say()
        say("Run it again without --survey to carry the plan out:")
        say()
        say("    migrateEdSharp.cmd")
    else:
        say(f"Moved {iFolders} folders, {iHelp} documents, {iConfigs} settings files,")
        say(f"{iExec} build products, {iInstall} install scripts and {iJaws} reader scripts.")
        say(f"Retired {iRetired} files into notes and deleted {iClasses} class copies.")
        say()
        say("Next:")
        say("  1. buildEdSharp")
        say("  2. exec\\EdSharp.exe            (the quick test)")
        say("  3. scripts\\checkHomerApp --build")
        say("  4. scripts\\gitPush \"Move to the Homer Development Kit.\"")
        say("  5. scripts\\tagRelease")
    say()
    say("The log is at " + pathLog)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception:
        import traceback
        say("")
        say("migrateEdSharp stopped on an unexpected error:")
        for sLine in traceback.format_exc().splitlines():
            say("  " + sLine)
        say("The log is at " + pathLog + ". Nothing further was attempted.")
        sys.exit(1)
    finally:
        if fileLog:
            fileLog.close()
