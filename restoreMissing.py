r"""restoreMissing.py -- put back files the project needs that have vanished.

    restoreMissing.cmd                (restore whatever is missing)
    restoreMissing.cmd --survey       (print what is missing and change nothing)

WHAT THIS IS FOR

On 28 August fifteen installer scripts and three fetch scripts disappeared
from C:\EdSharp, and the build stopped on its own audit saying so. They were
gone from the folder AND from the current commit, which is why tidyRepo's
restore step could not help: that step asks git which tracked files are
missing, and a file removed from the commit is not tracked any more.

Git had not forgotten them. Every one was still in the history, one commit
back. This walks the list of files the project claims, finds the ones that
are not on disk, and brings each back from the newest commit that still
holds it.

WHERE IT LOOKS, IN ORDER

  1. The working tree. Present means nothing to do.
  2. The current commit. A tracked file that is merely deleted from disk
     comes straight back.
  3. The history, newest first, including the remote's branch. For each
     commit that touched the file, the file is read from that commit and, if
     that commit is the one that deleted it, from its parent.

Nothing is overwritten: a file that is already on disk is left exactly as it
is, whatever the history says. This only ever creates what is absent.

A detailed log is written beside this script, whatever happens.
"""

import argparse
import datetime
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import repoPolicy

c_sLogName = "restoreMissing.log"
c_iCommitsToTry = 40

pathRoot = os.path.dirname(os.path.abspath(__file__))
pathLog = os.path.join(pathRoot, c_sLogName)
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
    say("restoreMissing  " + datetime.datetime.now().isoformat(" ", "seconds"))
    say("  script:            " + os.path.abspath(__file__))
    say("  Python:            " + sys.version.split()[0])
    say("  platform:          " + sys.platform)
    say("  working directory: " + os.getcwd())
    say("  command line:      " + " ".join([os.path.basename(sys.argv[0])] + sys.argv[1:]))
    say()


def gitText(lArguments):
    """Run a git command for its text output, recording it in the log."""
    say("  > git " + " ".join(lArguments))
    try:
        oResult = subprocess.run(["git", "-c", "core.quotepath=false"] + lArguments,
                                 cwd=pathRoot, capture_output=True, text=True,
                                 encoding="utf-8", errors="replace")
    except Exception as oError:
        say("    could not run it: " + str(oError))
        return None
    if oResult.returncode != 0:
        say("    exit code " + str(oResult.returncode))
        return None
    return oResult.stdout or ""


def gitBytes(lArguments):
    """Run a git command for its raw output, so a library survives intact."""
    try:
        oResult = subprocess.run(["git"] + lArguments, cwd=pathRoot,
                                 capture_output=True)
    except Exception:
        return None
    if oResult.returncode != 0:
        return None
    return oResult.stdout


def plural(iCount, sNoun):
    return f"{iCount} {sNoun}" + ("" if iCount == 1 else "s")


def claimedFiles():
    """Every file the project claims by name, folders excluded.

    Folder-wide Source lines are left out on purpose: this restores named
    files, and guessing at the contents a folder ought to have is how a
    tidy-up turns into an invention.
    """
    oInstalled = repoPolicy.installedFiles(pathRoot)
    if oInstalled is None:
        return None
    setExact, lFolders = oInstalled
    lNames = sorted(set(setExact) | set(repoPolicy.c_lDevelopmentFiles), key=str.lower)
    return [s for s in lNames if "/" not in s]


def missingFiles(lClaimed):
    """The claimed files that are not on disk."""
    return [s for s in lClaimed
            if not os.path.exists(os.path.join(pathRoot, s.replace("/", os.sep)))]


def revisionsFor(sPath):
    """Places worth looking for a file, newest first."""
    lRevisions = ["HEAD", "origin/master"]
    sLog = gitText(["rev-list", "-n", str(c_iCommitsToTry), "--all", "--", sPath])
    for sCommit in [s.strip() for s in (sLog or "").splitlines() if s.strip()]:
        # The commit that touched it, and its parent: when the touch was the
        # deletion, the file lives in the parent and not in the commit itself.
        lRevisions.append(sCommit)
        lRevisions.append(sCommit + "^")
    return lRevisions


def restoreOne(sPath):
    """Bring one file back from the newest place that still holds it.

    git checkout does the work, not a hand-written copy of the stored bytes.
    The difference matters here: .gitattributes says every text file is CRLF
    on checkout, and git applies that on the way out. Writing the stored form
    straight to disk produced .cmd and .ps1 files with Unix line endings.
    """
    for sRevision in revisionsFor(sPath):
        if gitBytes(["cat-file", "-e", f"{sRevision}:{sPath}"]) is None:
            continue
        if gitText(["checkout", sRevision, "--", sPath]) is None:
            say(f"  COULD NOT CHECK OUT {sPath} from {sRevision}")
            continue
        pathFile = os.path.join(pathRoot, sPath.replace("/", os.sep))
        iSize = os.path.getsize(pathFile) if os.path.exists(pathFile) else 0
        say(f"  restored {sPath} from {sRevision} ({iSize:,} bytes)")
        return sRevision
    say(f"  NOT IN THE HISTORY AT ALL: {sPath}")
    return None


def main():
    oParser = argparse.ArgumentParser(
        description="Restore files the project needs that are missing from disk.")
    oParser.add_argument("--survey", action="store_true",
                         help="print what is missing and change nothing")
    oArguments = oParser.parse_args()

    startLog()
    if not os.path.isdir(os.path.join(pathRoot, ".git")):
        say("This is not a git working folder, so there is nothing to restore from.")
        say("Run it from C:\\EdSharp.")
        say("The log is at " + pathLog)
        return 1
    lClaimed = claimedFiles()
    if lClaimed is None:
        say("EdSharp_Setup.iss is not here, so nothing can be judged.")
        say("Run it from C:\\EdSharp.")
        say("The log is at " + pathLog)
        return 1
    gitText(["fetch", "origin", "--quiet"])
    lMissing = missingFiles(lClaimed)

    say("=" * 68)
    say("PLAN")
    say("=" * 68)
    say()
    say(plural(len(lClaimed), "file") + " is claimed by name."
        if len(lClaimed) == 1 else
        plural(len(lClaimed), "file") + " are claimed by name.")
    if not lMissing:
        say("None of them is missing. Nothing to do.")
        say()
        say("The log is at " + pathLog)
        return 0
    sIs = "is" if len(lMissing) == 1 else "are"
    say(f"{len(lMissing)} of them {sIs} not on disk:")
    say()
    for sPath in lMissing:
        say("  " + sPath)
    say()

    if oArguments.survey:
        say("This was a description only (--survey). Nothing has been changed.")
        say()
        say("Run it again without --survey to bring them back:")
        say()
        say("    restoreMissing.cmd")
        say("The log is at " + pathLog)
        return 0

    say("=" * 68)
    say("DOING IT")
    say("=" * 68)
    say()
    lDone, lFailed = [], []
    for sPath in lMissing:
        if restoreOne(sPath):
            lDone.append(sPath)
        else:
            lFailed.append(sPath)
    say()

    say("=" * 68)
    say("AFTERWARDS")
    say("=" * 68)
    say()
    say(plural(len(lDone), "file") + " restored.")
    if lFailed:
        say(plural(len(lFailed), "file") + " could not be found anywhere in the")
        say("history, so each must come from a backup or be written again:")
        for sPath in lFailed:
            say("  " + sPath)
    lStill = missingFiles(lClaimed)
    if lStill:
        say("Still missing: " + ", ".join(lStill))
    else:
        say("Nothing the project claims is missing now.")
    say()
    say("Next, run auditEdSharp.cmd to confirm, then build.")
    say()
    say("The log is at " + pathLog)
    return 1 if lFailed else 0


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
        say("restoreMissing stopped on an unexpected error:")
        for sLine in traceback.format_exc().splitlines():
            say("  " + sLine)
        say("The log is at " + pathLog + ". Nothing further was attempted.")
        sys.exit(1)
    finally:
        if fileLog:
            fileLog.close()
