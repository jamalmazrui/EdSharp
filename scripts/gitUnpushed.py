#!/usr/bin/env python3
"""gitUnpushed.py -- undo the commits that have not been pushed, keeping every file.

WHEN IT IS FOR. A tidy or a hand commit has swept something into the
repository that should never go there -- fetched voices, a model, a build's
output -- and the push was stopped, or has not happened yet. The commit is
only local. This puts the branch back to what the remote has, leaves every
file on disk exactly as it is, and unstages everything, so the next commit
can be made properly (homerTidy, with RepoFiles.txt in place, makes it).

WHAT IT REFUSES. If every local commit is already on the remote there is
nothing to undo, and it says so. It never rewrites what has been pushed.

Run it in the project folder. It logs to logs\\<App>-unpushed-yyyyMMdd-HHmmss.log.
"""

import datetime, os, platform, subprocess, sys

oLog = None


def projectRoot(sStart):
    """The current folder, or its parent when the current folder is the
    project's scripts or exec folder -- the one rule every Homer tool follows."""
    if os.path.basename(sStart).lower() in ("scripts", "exec", "tools"):
        sParent = os.path.dirname(sStart)
        if os.path.isfile(os.path.join(sParent, "version.txt")) or os.path.isfile(os.path.join(sParent, "RepoFiles.txt")):
            return sParent
    return sStart


sRoot = projectRoot(os.getcwd())


def logLine(sText):
    if oLog is not None:
        oLog.write(sText + "\n")
        oLog.flush()
    return True


def say(sText):
    print(sText)
    logLine("CONSOLE: " + sText)
    return True


def runGit(lsArgs):
    logLine("RUN: git " + " ".join(lsArgs))
    oResult = subprocess.run(["git"] + lsArgs, cwd=sRoot, capture_output=True, text=True)
    logLine("EXIT: %d" % oResult.returncode)
    if oResult.stdout.strip(): logLine("STDOUT:\n" + oResult.stdout.rstrip())
    if oResult.stderr.strip(): logLine("STDERR:\n" + oResult.stderr.rstrip())
    return oResult.returncode, oResult.stdout.strip()


def main():
    global oLog
    sLogDir = os.path.join(sRoot, "logs")
    os.makedirs(sLogDir, exist_ok=True)
    sLogPath = os.path.join(sLogDir, "%s-unpushed-%s.log" % (os.path.basename(sRoot), datetime.datetime.now().strftime("%Y%m%d-%H%M%S")))
    oLog = open(sLogPath, "w", encoding="utf-8")
    logLine("gitUnpushed started %s" % datetime.datetime.now().isoformat(" ", "seconds"))
    logLine("Script: %s" % os.path.abspath(__file__))
    logLine("Python: %s" % sys.version.replace("\n", " "))
    logLine("Platform: %s" % platform.platform())
    logLine("Working directory: %s" % sRoot)
    logLine("Command line: %s" % " ".join(sys.argv))
    iCode, sOut = runGit(["rev-parse", "--is-inside-work-tree"])
    if iCode != 0:
        say("This folder is not a git repository.")
        return 1
    iCode, sBranch = runGit(["rev-parse", "--abbrev-ref", "HEAD"])
    iCode, sUpstream = runGit(["rev-parse", "--abbrev-ref", "--symbolic-full-name", "@{u}"])
    if iCode != 0:
        say("The branch %s has no remote branch to compare with, so nothing can be told apart as unpushed." % sBranch)
        return 1
    iCode, sList = runGit(["log", "--oneline", sUpstream + "..HEAD"])
    lsCommits = [l for l in sList.split("\n") if l.strip()]
    if not lsCommits:
        say("Every commit on %s is already on %s. Nothing to undo." % (sBranch, sUpstream))
        return 0
    say("%d commit%s on %s not yet on %s:" % (len(lsCommits), "" if len(lsCommits) == 1 else "s", sBranch, sUpstream))
    for sLine in lsCommits: say("  " + sLine)
    # Mixed reset: the branch goes back to the remote's commit, the index is
    # cleared, and the working files are untouched.
    iCode, sOut = runGit(["reset", "--mixed", sUpstream])
    if iCode != 0:
        say("The reset failed. Nothing was changed; the log has the message.")
        return 1
    say("Undone. Every file is as it was, nothing is staged, and %s is back at %s." % (sBranch, sUpstream))
    say("Run homerTidy --do-it to make the commit properly; it needs RepoFiles.txt.")
    say("Log: " + sLogPath)
    logLine("Finished %s" % datetime.datetime.now().isoformat(" ", "seconds"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
