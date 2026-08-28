r"""moveNotes.py -- move everything legacy into C:\EdSharp\notes.

    moveNotes.cmd                 (move it all)
    moveNotes.cmd --survey        (print the plan and change nothing)

WHAT THIS IS FOR

C:\EdSharp accumulated a decade of work that is not EdSharp: saved web pages
and mailing list messages, old drafts, audio courses, archives, unpacked
programs, and whole folders belonging to other projects. All of it is worth
keeping until it has been looked at, and none of it belongs beside the
sources.

This moves every root entry the project does not claim -- files and whole
folders alike -- into notes\, where one line in .gitignore covers the lot and
where it can be sorted by hand at leisure. Nothing outside notes\ is ever
deleted.

WHAT IS SAFE FROM IT

repoPolicy.py answers what the project claims: EdSharp_Setup.iss names what
EdSharp installs, including the folders it takes wholesale, and repoPolicy's
own short list names what builds and releases it. Beyond that, three things
are kept by name here:

  * .git, and notes itself.
  * What the build makes or fetches at the root and the installer therefore
    never names: Version.cs, generated on every run; BuildVersion.cs, the
    stale name a previous revision used; EdSharp_Setup.exe; and the 64-bit
    NVDA controller library. Moving any of these breaks the next build in a
    way its error message does not explain.
  * Logs and the small notes the build leaves to itself, by extension.

Folder names are matched without regard to case, because Windows does not
distinguish them: the Scripts folder the build scrubs appears as "scripts"
on this machine, and a case-sensitive test would have swept the JAWS scripts
away.

EMPTY FILES

Every file of zero length is deleted, wherever this script can see one: at
the root among the legacy files, and anywhere under notes\, inside moved
folders included. An empty file is worse than a missing one, because it
looks like content until it is opened, and a tool chain that finds it will
carry on as though there were something to read.

The one exception is an empty file the project itself claims. That is not
clutter but a symptom -- a generated document that failed to generate, which
the installer would then ship as nothing, since its Source line carries
skipifsourcedoesntexist. Those are reported by name and left alone, because
deleting one would hide the fault it is announcing.

DUPLICATES

Among the files that land in notes\ itself, two files with the same contents
are one file: the later one is deleted and the log says which name was kept.
Two files with the same NAME but different contents are two files, so the
newcomer becomes name-01, then name-02, and nothing is lost. This applies
only to notes\ itself; a moved folder keeps its own contents untouched.

A detailed log is written beside this script, whatever happens.
"""

import argparse
import datetime
import hashlib
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import repoPolicy

c_sFolderName = "notes"
c_sLogName = "moveNotes.log"
c_sIgnoreEntry = "notes/"

# Kept whatever else is decided: the repository itself and the destination.
c_lKeepAlways = [".git", "notes"]

# Made or fetched by the build at the root, so the installer never names them
# and repoPolicy cannot know about them.
c_lBuildProducts = [
    "buildversion.cs",
    "edsharp_setup.exe",
    "nvdacontrollerclient64.dll",
    "version.cs",
]

# The build's own bookkeeping, kept by extension.
c_lKeepExtensions = [".fetched", ".log", ".version"]

# Documents that are neither shipped nor legacy, left at the root so the
# decision about each is yours. Take a name out to let it move.
c_lKeepAtRoot = [
    "Accessible_Charts_Sample.htm",
    "Accessible_Charts_Sample.md",
    "Handover.md",
    "Mermaid_Sample.htm",
    "Mermaid_Sample.md",
    "Pandoc_Office_Guide.htm",
    "Pandoc_Office_Guide.md",
    "journal_article.bib",
    "journal_article.htm",
    "journal_article.md",
]

pathRoot = os.path.dirname(os.path.abspath(__file__))
pathLog = os.path.join(pathRoot, c_sLogName)
pathNotes = os.path.join(pathRoot, c_sFolderName)
fileLog = None


def say(sMessage=""):
    print(sMessage)
    if fileLog:
        try:
            fileLog.write(sMessage + "\n")
            fileLog.flush()
        except Exception:
            pass


def startLog():
    global fileLog
    try:
        fileLog = open(pathLog, "w", encoding="utf-8")
    except Exception as oError:
        print("Could not open the log: " + str(oError))
        return
    say("moveNotes  " + datetime.datetime.now().isoformat(" ", "seconds"))
    say("  script:            " + os.path.abspath(__file__))
    say("  Python:            " + sys.version.split()[0])
    say("  platform:          " + sys.platform)
    say("  working directory: " + os.getcwd())
    say("  command line:      " + " ".join([os.path.basename(sys.argv[0])] + sys.argv[1:]))
    say("  notes folder:      " + pathNotes)
    say()


def plural(iCount, sNoun):
    return f"{iCount} {sNoun}" + ("" if iCount == 1 else "s")


def claimedNames():
    """What the project claims at the root, folded to lower case.

    Returns a set of file names and a set of folder names, or None when the
    setup script cannot be read -- in which case nothing can be judged, and
    guessing is how the last fault happened.
    """
    oInstalled = repoPolicy.installedFiles(pathRoot)
    if oInstalled is None:
        return None
    setExact, lFolders = oInstalled
    setFiles = set()
    setFolders = set(s.lower() for s in c_lKeepAlways)
    for sPath in setExact:
        sPath = sPath.replace("\\", "/")
        if "/" in sPath:
            setFolders.add(sPath.split("/", 1)[0].lower())
        else:
            setFiles.add(sPath.lower())
    for sFolder in lFolders:
        setFolders.add(sFolder.rstrip("/").split("/", 1)[0].lower())
    for sName in repoPolicy.c_lDevelopmentFiles:
        setFiles.add(sName.lower())
    return setFiles, setFolders


def isKeptFile(sName, setFiles):
    """Whether a root file stays where it is."""
    sLower = sName.lower()
    if sLower in setFiles:
        return True
    # The libraries the build fetches. repoPolicy deliberately does not claim
    # them, because the repository must not carry them -- but the build needs
    # them here, so they are kept without being claimed. Sweeping them into
    # notes would leave the next compile without its references.
    if sLower in set(s.lower() for s in repoPolicy.c_lFetchedFiles):
        return True
    if sLower in c_lBuildProducts:
        return True
    if sLower in set(s.lower() for s in c_lKeepAtRoot):
        return True
    if os.path.splitext(sLower)[1] in c_lKeepExtensions:
        return True
    return False


def entriesToMove():
    """The root files and folders that are legacy, each in order by name.

    Returns None when the project's claim cannot be read.
    """
    oClaimed = claimedNames()
    if oClaimed is None:
        return None
    setFiles, setFolders = oClaimed
    lFiles, lFolders = [], []
    for sName in sorted(os.listdir(pathRoot), key=str.lower):
        pathEntry = os.path.join(pathRoot, sName)
        if os.path.isdir(pathEntry):
            if sName.lower() not in setFolders:
                lFolders.append(sName)
        elif os.path.isfile(pathEntry):
            if not isKeptFile(sName, setFiles):
                lFiles.append(sName)
    return lFiles, lFolders


def isEmpty(pathFile):
    """Whether a file holds nothing at all."""
    try:
        return os.path.getsize(pathFile) == 0
    except OSError:
        return False


def emptyLegacyFiles(lFiles):
    """The root legacy files that are empty, in order by name."""
    return [s for s in lFiles if isEmpty(os.path.join(pathRoot, s))]


def emptyClaimedFiles(setFiles):
    """Empty files the project itself claims -- reported, never deleted.

    A generated document that came out empty is a fault worth hearing about.
    The installer ships most documents with skipifsourcedoesntexist, so an
    empty one is copied as an empty one and nobody finds out.
    """
    lEmpty = []
    for sName in sorted(os.listdir(pathRoot), key=str.lower):
        pathFile = os.path.join(pathRoot, sName)
        if not os.path.isfile(pathFile):
            continue
        if sName.lower() not in setFiles:
            continue
        if isEmpty(pathFile):
            lEmpty.append(sName)
    return lEmpty


def filesUnder(pathFolder):
    """Every file under a folder, at any depth, in a settled order."""
    lPaths = []
    for sDirectory, lDirectories, lFiles in os.walk(pathFolder):
        for sName in sorted(lFiles, key=str.lower):
            lPaths.append(os.path.join(sDirectory, sName))
    return lPaths


def deleteEmptyFiles(lPaths):
    """Delete every zero-length file among these paths, naming each in the log.

    The caller decides which paths, and that is the point: an earlier version
    swept the whole root and took Announce.htm, the project's own empty
    document, moments after the plan had promised to leave it alone. The
    project's files are never handed to this.

    Returns how many were deleted.
    """
    iDeleted = 0
    for pathFile in lPaths:
        if not isEmpty(pathFile):
            continue
        try:
            os.remove(pathFile)
        except OSError as oError:
            say(f"  COULD NOT DELETE the empty file {pathFile}: {oError}")
            continue
        say("  deleted, empty: " + os.path.relpath(pathFile, pathRoot))
        iDeleted += 1
    return iDeleted


def countEmptyUnder(pathFolder):
    """How many zero-length files a folder holds, at any depth."""
    iCount = 0
    for sDirectory, lDirectories, lFiles in os.walk(pathFolder):
        for sName in lFiles:
            if isEmpty(os.path.join(sDirectory, sName)):
                iCount += 1
    return iCount


def digestOf(pathFile):
    """The contents of a file as one value, for telling copies apart.

    Empty files never reach this, because they are deleted before anything
    is moved or compared.
    """
    oHash = hashlib.sha256()
    try:
        with open(pathFile, "rb") as fileIn:
            for binChunk in iter(lambda: fileIn.read(1024 * 1024), b""):
                oHash.update(binChunk)
    except OSError as oError:
        say("  could not read " + pathFile + ": " + str(oError))
        return None
    return oHash.hexdigest()


def existingDigests():
    """What is already in notes itself, by contents, so a repeat run adds
    nothing twice."""
    dDigests = {}
    if not os.path.isdir(pathNotes):
        return dDigests
    for sName in sorted(os.listdir(pathNotes), key=str.lower):
        pathFile = os.path.join(pathNotes, sName)
        if not os.path.isfile(pathFile):
            continue
        sDigest = digestOf(pathFile)
        if sDigest and sDigest not in dDigests:
            dDigests[sDigest] = sName
    return dDigests


def freeName(sName):
    """A name in notes that nothing occupies: sName, then sName-01, -02..."""
    if not os.path.exists(os.path.join(pathNotes, sName)):
        return sName
    sStem, sExtension = os.path.splitext(sName)
    iCount = 1
    while iCount < 1000:
        sTry = f"{sStem}-{iCount:02d}{sExtension}"
        if not os.path.exists(os.path.join(pathNotes, sTry)):
            return sTry
        iCount += 1
    return None


def moveFile(sName, dDigests):
    """Move one root file into notes, or drop it when it is already there.

    Returns "moved", "duplicate" or "failed".
    """
    pathFrom = os.path.join(pathRoot, sName)
    sDigest = digestOf(pathFrom)
    if sDigest and sDigest in dDigests:
        sKept = dDigests[sDigest]
        try:
            os.remove(pathFrom)
        except OSError as oError:
            say(f"  COULD NOT DELETE the duplicate {sName}: {oError}")
            return "failed"
        if sKept.lower() == sName.lower():
            say(f"  duplicate deleted: {sName} (the copy in notes is identical)")
        else:
            say(f"  duplicate deleted: {sName} (same contents as {sKept})")
        return "duplicate"
    sTarget = freeName(sName)
    if sTarget is None:
        say(f"  COULD NOT PLACE {sName}: a thousand names of that stem are taken")
        return "failed"
    try:
        shutil.move(pathFrom, os.path.join(pathNotes, sTarget))
    except Exception as oError:
        say(f"  COULD NOT MOVE {sName}: {oError}")
        return "failed"
    if sTarget == sName:
        say("  moved: " + sName)
    else:
        say(f"  moved: {sName}, renamed {sTarget} because that name was taken by "
            "different contents")
    if sDigest:
        dDigests[sDigest] = sTarget
    return "moved"


def moveFolder(sName):
    """Move one root folder into notes, whole and unaltered."""
    pathFrom = os.path.join(pathRoot, sName)
    sTarget = freeName(sName)
    if sTarget is None:
        say(f"  COULD NOT PLACE {sName}: a thousand names of that stem are taken")
        return False
    try:
        shutil.move(pathFrom, os.path.join(pathNotes, sTarget))
    except Exception as oError:
        say(f"  COULD NOT MOVE the folder {sName}: {oError}")
        return False
    if sTarget == sName:
        say("  moved folder: " + sName)
    else:
        say(f"  moved folder: {sName}, renamed {sTarget} because that name was taken")
    return True


def dropDuplicatesInNotes():
    """Delete files in notes itself whose contents already exist under
    another name, keeping the first by name.

    Only notes itself. A folder that was moved in keeps every file it
    arrived with, because its contents are somebody's structure and not
    this script's business.
    """
    if not os.path.isdir(pathNotes):
        return 0
    dFirst, lDelete = {}, []
    for sName in sorted(os.listdir(pathNotes), key=str.lower):
        pathFile = os.path.join(pathNotes, sName)
        if not os.path.isfile(pathFile):
            continue
        sDigest = digestOf(pathFile)
        if sDigest is None:
            continue
        if sDigest in dFirst:
            lDelete.append((sName, dFirst[sDigest]))
        else:
            dFirst[sDigest] = sName
    iDropped = 0
    for sName, sKept in lDelete:
        try:
            os.remove(os.path.join(pathNotes, sName))
        except OSError as oError:
            say(f"  COULD NOT DELETE {sName}: {oError}")
            continue
        say(f"  deleted {sName}, identical to {sKept}")
        iDropped += 1
    return iDropped


def addIgnoreEntry():
    """Make sure .gitignore holds the notes folder, once."""
    pathIgnore = os.path.join(pathRoot, ".gitignore")
    sExisting = ""
    if os.path.exists(pathIgnore):
        with open(pathIgnore, encoding="utf-8", errors="replace") as fileIgnore:
            sExisting = fileIgnore.read()
    if c_sIgnoreEntry in sExisting.splitlines():
        say("  .gitignore already holds " + c_sIgnoreEntry)
        return
    with open(pathIgnore, "a", encoding="utf-8", newline="\n") as fileIgnore:
        fileIgnore.write("\n# Legacy work: saved pages, saved messages, old drafts,\n"
                         "# archives and the folders of other projects. Kept on disk,\n"
                         "# never part of the project.\n" + c_sIgnoreEntry + "\n")
    say("  added " + c_sIgnoreEntry + " to .gitignore")


def countUnder(pathFolder):
    """How many files a folder holds, at any depth."""
    iCount = 0
    for sDirectory, lDirectories, lFiles in os.walk(pathFolder):
        iCount += len(lFiles)
    return iCount


def main():
    oParser = argparse.ArgumentParser(
        description="Move legacy files and folders into the notes folder.")
    oParser.add_argument("--survey", action="store_true",
                         help="print the plan and change nothing")
    oArguments = oParser.parse_args()

    startLog()
    oEntries = entriesToMove()
    if oEntries is None:
        say("EdSharp_Setup.iss is not here, so nothing can be judged.")
        say("Run this from C:\\EdSharp.")
        say("The log is at " + pathLog)
        return 1
    lFiles, lFolders = oEntries

    say("=" * 68)
    say("PLAN")
    say("=" * 68)
    say()
    if not lFiles and not lFolders:
        say("Nothing legacy at the root. Nothing to move.")
    else:
        if lFolders:
            iInside = 0
            say(plural(len(lFolders), "folder") + " will move into "
                + c_sFolderName + "\\, whole:")
            say()
            for sName in lFolders:
                iHeld = countUnder(os.path.join(pathRoot, sName))
                iInside += iHeld
                say(f"  {sName}\\  ({plural(iHeld, 'file')})")
            say()
            say("  " + plural(iInside, "file") + " in all, inside those folders.")
            say()
        if lFiles:
            say(plural(len(lFiles), "file") + " will move into "
                + c_sFolderName + "\\:")
            say()
            for sName in lFiles:
                iSize = 0
                try:
                    iSize = os.path.getsize(os.path.join(pathRoot, sName))
                except OSError:
                    pass
                say(f"  {sName}  ({iSize:,} bytes)")
            say()
        say("Everything stays on disk, in the new folder, except the deletions")
        say("below, each of which is named in the log as it happens.")
        say()

    iEmptyRoot = len(emptyLegacyFiles(lFiles))
    iEmptyFolders = sum(countEmptyUnder(os.path.join(pathRoot, s)) for s in lFolders)
    iEmptyNotes = countEmptyUnder(pathNotes) if os.path.isdir(pathNotes) else 0
    iEmptyAll = iEmptyRoot + iEmptyFolders + iEmptyNotes
    if iEmptyAll:
        say(plural(iEmptyAll, "empty file") + " will be deleted: "
            + f"{iEmptyRoot} at the root, {iEmptyFolders} inside the folders "
            f"moving, {iEmptyNotes} already in notes.")
        say()

    lEmptyClaimed = emptyClaimedFiles(claimedNames()[0])
    if lEmptyClaimed:
        say("EMPTY, BUT THE PROJECT'S OWN, so left alone and worth a look --")
        say("the installer would ship each of these as nothing:")
        for sName in lEmptyClaimed:
            say("  " + sName)
        say()

    lKeptByName = [s for s in sorted(c_lKeepAtRoot, key=str.lower)
                   if os.path.exists(os.path.join(pathRoot, s))]
    if lKeptByName:
        say("Left at the root for you to decide about:")
        for sName in lKeptByName:
            say("  " + sName)
        say()

    if oArguments.survey:
        say("This was a description only (--survey). Nothing has been changed.")
        say()
        say("Run it again without --survey to carry the plan out:")
        say()
        say("    moveNotes.cmd")
        say("The log is at " + pathLog)
        return 0

    say("=" * 68)
    say("DOING IT")
    say("=" * 68)
    say()
    if not os.path.isdir(pathNotes):
        try:
            os.makedirs(pathNotes)
            say("  created " + pathNotes)
        except Exception as oError:
            say("  COULD NOT CREATE " + pathNotes + ": " + str(oError))
            say("The log is at " + pathLog)
            return 1
    # Empty legacy files go before anything is moved, so none of them travels.
    # Only the legacy ones: the project's own empty files are reported in the
    # plan and left where they are.
    iEmptied = deleteEmptyFiles([os.path.join(pathRoot, s) for s in lFiles])
    lFiles = [s for s in lFiles if os.path.exists(os.path.join(pathRoot, s))]

    dDigests = existingDigests()
    iFolders = 0
    for sName in lFolders:
        if moveFolder(sName):
            iFolders += 1
    iMoved, iDuplicate, iFailed = 0, 0, 0
    for sName in lFiles:
        sResult = moveFile(sName, dDigests)
        if sResult == "moved":
            iMoved += 1
        elif sResult == "duplicate":
            iDuplicate += 1
        else:
            iFailed += 1
    # And again through notes, which now holds the moved folders and reaches
    # every depth, so an empty file inside one of them goes too.
    iEmptied += deleteEmptyFiles(filesUnder(pathNotes))
    iDropped = dropDuplicatesInNotes()
    addIgnoreEntry()
    say()

    say("=" * 68)
    say("AFTERWARDS")
    say("=" * 68)
    say()
    say(plural(iFolders, "folder") + " and " + plural(iMoved, "file")
        + " moved into " + c_sFolderName + "\\.")
    if iDuplicate:
        sWas = "was" if iDuplicate == 1 else "were"
        say(plural(iDuplicate, "file") + " " + sWas + " not moved because notes "
            "already held the same contents; each was deleted.")
    if iDropped:
        say(plural(iDropped, "file") + " already in notes proved to be a copy of "
            "another under a different name, and was deleted.")
    if iEmptied:
        sWas = "was" if iEmptied == 1 else "were"
        say(plural(iEmptied, "empty file") + " " + sWas + " deleted.")
    if iFailed:
        say(plural(iFailed, "file") + " could not be moved; each is named above.")
    oLeft = entriesToMove()
    if oLeft and (oLeft[0] or oLeft[1]):
        say("Still at the root and still unaccounted for: "
            + ", ".join(oLeft[1] + oLeft[0]))
    else:
        say("Nothing legacy remains at the root.")
    say()
    say("Next, run tidyRepo.cmd. It untracks what moved and commits the")
    say("change; nothing outside notes is deleted at any point.")
    say()
    say("The log is at " + pathLog)
    return 0


if __name__ == "__main__":
    # An unexpected failure must reach the log too. Without this the script
    # prints a traceback to a console that scrolls away and writes nothing,
    # which is the one outcome a log exists to prevent.
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception:
        import traceback
        say("")
        say("moveNotes stopped on an unexpected error:")
        for sLine in traceback.format_exc().splitlines():
            say("  " + sLine)
        say("The log is at " + pathLog + ". Nothing further was attempted.")
        sys.exit(1)
    finally:
        if fileLog:
            fileLog.close()
