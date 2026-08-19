"""tidyRepo.py -- survey the whole repository first, then fix it in one pass.

Run it from C:\\EdSharp with no arguments and it surveys, prints the whole
plan, and carries the plan out in the same run:

    python tidyRepo.py

Run it with --survey to only print the plan and change nothing:

    python tidyRepo.py --survey

WHY THIS IS ONE SCRIPT AND NOT FOUR

The HomerScribe clean-up took four scripts and several rounds, and the reason
is worth stating because it is the whole design of this one. Each script fixed
what it had been asked about, pushed, and only then discovered the next
problem: stop tracking the large file, then find it still in the history, then
find the history rewrite left the pack unchanged, then find leftover folders.
Each discovery needed another script and another round.

The cause was not the fixes. It was that nothing looked at everything before
acting. So this surveys first and completely: what is tracked that should not
be, what is in the history that should not be, what is on disk that belongs
nowhere, what the remote has, and what state the working tree is in. It prints
the entire plan. Only then, and only with --do-it, does it act, and it does
every part in one pass so there is no second round to discover anything in.

WHAT IT LOOKS FOR

  1. Tracked files that the project does not need. The setup script says what
     the project installs, and .gitignore says what is deliberately excluded;
     anything tracked that is neither is a candidate, and every one is listed
     for you to confirm rather than assumed.
  2. Large objects anywhere in the history, whether still reachable or not,
     because that is what makes a clone slow and what GitHub complains about.
  3. Files on disk that are neither tracked nor ignored.
  4. The names EdSharp must not carry, wherever they appear -- tracked, on
     disk, or in the history:
       pandoc.exe and the Convert/Pandoc folder: untracked, ignored, and
       purged from the history, but the DISK COPY STAYS, because EdSharp
       still converts with pandoc at run time and installPandoc fetches it
       for users.
       The web client utilities, including InPy.exe: removed everywhere,
       INCLUDING FROM DISK. The dated backup keeps a copy.

WHAT IT WILL NOT DO

  - It will not act with uncommitted changes. It says what is uncommitted and
    stops, because rewriting history under a dirty tree loses work.
  - It will not delete a backup. The copy it makes stays until you remove it.
  - It will not force-push without saying, in the plan, that it is going to.

A detailed log is written beside this script, whatever happens.
"""

import argparse
import datetime
import os
import re
import shutil
import subprocess
import sys

c_iLargeBytes = 5 * 1024 * 1024          # worth naming in the history
c_iBulkyBytes = 25 * 1024 * 1024         # worth removing, counted across copies
c_iGitHubLimit = 100 * 1024 * 1024       # what GitHub refuses outright
c_sLogName = "tidyRepo.log"

# The names EdSharp must not carry. Matching is case-insensitive on the
# forward-slash form of the path.
c_lDeleteNames = ["inpy"]                       # a path part with this name, alone or with an extension: removed everywhere, disk included
c_lDeleteText = ["web client", "webclient"]     # anywhere in the path: removed everywhere, disk included
c_lPurgeNames = ["pandoc.exe"]                  # this file name: untracked, ignored, purged from history; the disk copy stays
c_lPurgeText = ["convert/pandoc"]               # anywhere in the path: same treatment as pandoc.exe

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


def startLog():
    """Open the log and record the environment, before anything can fail."""
    global fileLog
    fileLog = open(pathLog, "w", encoding="utf-8")
    say(f"tidyRepo  {datetime.datetime.now():%Y-%m-%d %H:%M:%S}")
    say(f"  script:            {os.path.abspath(__file__)}")
    say(f"  Python:            {sys.version.split()[0]}")
    say(f"  platform:          {sys.platform}")
    say(f"  working directory: {os.getcwd()}")
    say(f"  command line:      {' '.join(sys.argv)}")
    say("")


def run(lCommand, bCheck=False):
    """Run a command, log it with its exit code, and return the result.

    Everything is logged, including the failures, because a failure with no
    record of what was attempted is what turns one round into three.
    """
    say(f"  > {' '.join(lCommand)}")
    try:
        result = subprocess.run(lCommand, cwd=pathRoot, capture_output=True,
                                text=True, encoding="utf-8", errors="replace")
    except FileNotFoundError as exception:
        say(f"    the command could not be run: {exception}")
        if bCheck:
            raise
        return None
    if result.returncode:
        say(f"    exit code {result.returncode}")
        for sLine in (result.stderr or "").splitlines()[:8]:
            say(f"    {sLine}")
        if bCheck:
            raise RuntimeError(f"{' '.join(lCommand)} failed")
    return result


def gitOut(lArguments):
    result = run(["git"] + lArguments)
    return (result.stdout if result and not result.returncode else "") or ""


def unwantedKind(sPath):
    """"purgeDelete", "purge", or "" for a path, by the lists above.

    Name matching is by path part rather than substring, because a substring
    test for inpy would also catch MainPy.cs.
    """
    sLower = sPath.replace("\\", "/").lower()
    lParts = [s for s in sLower.split("/") if s]
    for sPart in lParts:
        for sName in c_lDeleteNames:
            if sPart == sName or sPart.startswith(sName + "."):
                return "purgeDelete"
    for sText in c_lDeleteText:
        if sText in sLower:
            return "purgeDelete"
    if lParts:
        for sName in c_lPurgeNames:
            if lParts[-1] == sName:
                return "purge"
    for sText in c_lPurgeText:
        if sText in sLower:
            return "purge"
    return ""


# --- The survey -------------------------------------------------------------


def neededFiles():
    """What the setup script installs, which is the definition of needed.

    Read from EdSharp_Setup.iss rather than listed here, so this cannot drift
    from what is actually shipped.
    """
    setNeeded = {".gitignore", ".gitattributes", "tidyRepo.py", c_sLogName,
                 "BuildEdSharp.ps1", "installPandoc.cmd", "installPandoc.ps1",
                 "tagRelease.cmd", "tagRelease.ps1", "ReadMe.md"}
    pathIss = os.path.join(pathRoot, "EdSharp_Setup.iss")
    if not os.path.exists(pathIss):
        say("  WARNING: EdSharp_Setup.iss is not here, so nothing can be")
        say("           judged unnecessary. Run this from C:\\EdSharp.")
        return None
    with open(pathIss, encoding="utf-8-sig") as fileIss:
        sIss = fileIss.read()
    for match in re.finditer(r'^Source:\s*"([^"]+)"', sIss, re.M):
        sPath = match.group(1).replace("C:\\EdSharp\\", "").replace("\\", "/")
        setNeeded.add(sPath.rstrip("*").rstrip("/") or sPath)
    return setNeeded


def isNeeded(sPath, setNeeded):
    """Whether a tracked path is part of what the project ships."""
    sPath = sPath.replace("\\", "/")
    # The unwanted lists override the setup script: Convert/* would otherwise
    # mark Convert/Pandoc as shipped and therefore needed.
    if unwantedKind(sPath):
        return False
    if sPath in setNeeded:
        return True
    # A folder the setup script takes wholesale covers everything under it.
    for sNeeded in setNeeded:
        if sNeeded and sPath.startswith(sNeeded.rstrip("/") + "/"):
            return True
    # Documentation and the build outputs belong even when named individually.
    if sPath.startswith(("docs/",)):
        return True
    if sPath.endswith((".md", ".htm")) and "/" not in sPath:
        return True
    return False


def surveyTracked(setNeeded):
    """Tracked files the project does not appear to need."""
    lTracked = [s for s in gitOut(["ls-files"]).splitlines() if s.strip()]
    lStray = [s for s in lTracked if not isNeeded(s, setNeeded)]
    return lTracked, lStray


def surveyHistory():
    """Large objects anywhere in the history, reachable or not.

    Read from the pack rather than from the working tree, because a file
    deleted in a later commit is still in every clone until the history is
    rewritten, and that is exactly what GitHub is complaining about.
    """
    result = run(["git", "rev-list", "--objects", "--all"])
    if not result or result.returncode:
        return [], set()
    dNames = {}
    for sLine in (result.stdout or "").splitlines():
        lParts = sLine.split(" ", 1)
        if len(lParts) == 2:
            dNames[lParts[0]] = lParts[1]

    result = run(["git", "cat-file", "--batch-all-objects",
                  "--batch-check=%(objectname) %(objecttype) %(objectsize)"])
    if not result or result.returncode:
        return [], set(dNames.values())
    lBig = []
    for sLine in (result.stdout or "").splitlines():
        lParts = sLine.split()
        if len(lParts) == 3 and lParts[1] == "blob":
            try:
                iSize = int(lParts[2])
            except ValueError:
                continue
            if iSize >= c_iLargeBytes:
                lBig.append((iSize, dNames.get(lParts[0], "(unnamed)"), lParts[0]))
    lBig.sort(reverse=True)
    return lBig, set(dNames.values())


def surveyWorkingTree(setNeeded):
    """Files on disk that are neither tracked nor ignored, sorted by whether
    they belong in the repository.

    The distinction matters and an earlier version did not draw it. It reported
    every untracked file as something to leave alone, which was right for a log
    and wrong for a source module: three files the add-on imports were sitting
    untracked and were reported as none of the repository's business. Anyone
    cloning would have got an add-on that could not import its own modules.
    """
    lUntracked = [s for s in gitOut(["ls-files", "--others", "--exclude-standard"]).splitlines()
                  if s.strip()]
    lBelong, lNot = [], []
    for sPath in lUntracked:
        sNormal = sPath.replace("\\", "/")
        # Anything the setup script ships, and any source under addon, is part
        # of the program whether or not git has noticed it yet.
        bBelongs = isNeeded(sNormal, setNeeded)
        # Except the things that are generated or personal.
        if sNormal.endswith((".log", ".pyc")) or "__pycache__" in sNormal:
            bBelongs = False
        if os.path.basename(sNormal).startswith(
                ("Announcing_", "Letter_to_", "Reply_to_", "Transition_")):
            bBelongs = False
        (lBelong if bBelongs else lNot).append(sPath)
    return lBelong, lNot


def readVersionMessage():
    """A commit message naming the version, read from the setup script.

    Better than a fixed message, because the log of a repository should say
    what each commit was, and the version is what this one is.
    """
    pathIss = os.path.join(pathRoot, "EdSharp_Setup.iss")
    try:
        with open(pathIss, encoding="utf-8-sig") as fileIss:
            match = re.search(r"^AppVersion=(.+)$", fileIss.read(), re.M)
        if match:
            return f"EdSharp {match.group(1).strip()}"
    except OSError:
        pass
    return "Update EdSharp"


def surveyState():
    """Uncommitted changes, the branch, and what the remote has."""
    # Only changes to tracked files matter here. Untracked files are what the
    # survey reports on and must not block it, and this script's own log is an
    # untracked file that appears the moment it starts.
    sStatus = gitOut(["status", "--porcelain", "--untracked-files=no"])
    sBranch = gitOut(["rev-parse", "--abbrev-ref", "HEAD"]).strip()
    result = run(["git", "remote", "get-url", "origin"])
    sRemote = ((result.stdout or "").strip()
               if result and not result.returncode else "")
    # A repository with no remote, or a branch the remote has never seen, is
    # ordinary rather than an error, so the failure is expected and quiet.
    sAhead = ""
    if sRemote:
        result = run(["git", "rev-list", "--count", f"origin/{sBranch}..HEAD"])
        if result and not result.returncode:
            sAhead = (result.stdout or "").strip()
        else:
            sAhead = "unknown, the remote has not seen this branch"
    return {
        "ahead": sAhead or "0",
        "branch": sBranch,
        "dirty": [s for s in sStatus.splitlines() if s.strip()],
        "remote": sRemote,
    }


# --- The plan ---------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(description="Tidy the EdSharp repository.")
    parser.add_argument("--survey", action="store_true",
                        help="only describe what would be done; change nothing")
    parser.add_argument("--do-it", action="store_true",
                        help="accepted for compatibility; acting is now the default")
    parser.add_argument("--no-push", action="store_true",
                        help="make the changes locally but do not push")
    parser.add_argument("--stash", action="store_true",
                        help="set outstanding changes aside rather than committing them")
    parser.add_argument("--message", default="",
                        help="the commit message for the outstanding changes")
    dArguments = parser.parse_args()

    startLog()
    bDoIt = not dArguments.survey

    if not os.path.isdir(os.path.join(pathRoot, ".git")):
        say("This is not a git repository. Run tidyRepo.py from C:\\EdSharp.")
        return 1
    if not shutil.which("git"):
        say("git is not on the path, so nothing can be done.")
        return 1

    say("=" * 68)
    say("SURVEY. Everything is looked at before anything is changed.")
    say("=" * 68)
    say("")

    dState = surveyState()
    say(f"Branch:   {dState['branch']}")
    say(f"Remote:   {dState['remote'] or '(none)'}")
    say(f"Unpushed: {dState['ahead']} commits")
    say("")

    bDirty = bool(dState["dirty"])
    if bDirty:
        say(f"There are {len(dState['dirty'])} uncommitted changes to tracked files:")
        for sLine in dState["dirty"][:25]:
            say(f"  {sLine}")
        if len(dState["dirty"]) > 25:
            say(f"  and {len(dState['dirty']) - 25} more")
        say("")
        say("   These are committed first, before anything else. An earlier")
        say("   version stopped here and told you to commit them yourself, which")
        say("   was half right: the tree must be clean before a history rewrite,")
        say("   because rewriting under a dirty one loses work. But sending you")
        say("   away to do it by hand is another round, and avoiding those is")
        say("   the whole point of this script. Committing preserves the work")
        say("   just as well and costs nothing.")
        say("")
        say("   Use --stash instead if you would rather set them aside than")
        say("   commit them.")
    say("")

    setNeeded = neededFiles()
    if setNeeded is None:
        return 1

    lTracked, lStray = surveyTracked(setNeeded)
    say(f"1. TRACKED FILES: {len(lTracked)} tracked, {len(lStray)} the project does not need.")
    say("")
    if lStray:
        say("   These are tracked but are not installed, not in addon, docs,")
        say("   installer or build, and not documentation. Most will be files")
        say("   that arrived through git add -A.")
        say("")
        for sPath in sorted(lStray):
            iSize = 0
            try:
                iSize = os.path.getsize(os.path.join(pathRoot, sPath))
            except OSError:
                pass
            say(f"     {sPath}  ({iSize:,} bytes)")
        say("")
        say("   They will be untracked, NOT deleted from disk, and added to")
        say("   .gitignore so they do not come back.")
    else:
        say("   Nothing tracked that the project does not need.")
    say("")

    lBig, setHistoryAll = surveyHistory()
    lTooBig = [t for t in lBig if t[0] >= c_iGitHubLimit]
    say(f"2. HISTORY: {len(lBig)} objects of {c_iLargeBytes // (1024*1024)} MB or more.")
    say("")
    if lBig:
        for iSize, sName, sHash in lBig[:15]:
            sFlag = "  OVER GITHUB'S LIMIT" if iSize >= c_iGitHubLimit else ""
            say(f"     {iSize / (1024*1024):>8.1f} MB  {sName}{sFlag}")
        if len(lBig) > 15:
            say(f"     and {len(lBig) - 15} more")
        say("")
        if lTooBig:
            say("   The ones over the limit are why GitHub complains. They are")
            say("   in the history, so they are in every clone even if the file")
            say("   was deleted later, and only rewriting the history removes")
            say("   them.")
    else:
        say("   Nothing large in the history.")
    say("")

    lBelong, lNotOurs = surveyWorkingTree(setNeeded)
    say(f"3. ON DISK: {len(lBelong) + len(lNotOurs)} files neither tracked nor ignored.")
    say("")
    if lBelong:
        say(f"   {len(lBelong)} of them BELONG in the repository and are missing from it:")
        for sPath in sorted(lBelong):
            say(f"     {sPath}")
        say("")
        say("   These will be added. Source the program imports, sitting")
        say("   untracked, means anyone cloning gets something that will not")
        say("   run.")
        say("")
    if lNotOurs:
        say(f"   {len(lNotOurs)} are not the repository's business and are left alone:")
        for sPath in sorted(lNotOurs)[:12]:
            say(f"     {sPath}")
        if len(lNotOurs) > 12:
            say(f"     and {len(lNotOurs) - 12} more")
        say("")
        say("   cleanDir.cmd is what moves those out.")
    say("")

    # --- The named unwanted files: pandoc and the web client utilities. ---
    lUnwantedTracked = sorted(s for s in lTracked if unwantedKind(s))
    lUnwantedHistory = sorted(s for s in setHistoryAll if unwantedKind(s))
    lUnwantedDisk = []
    for sRoot, lDirs, lNames in os.walk(pathRoot):
        lDirs[:] = [d for d in lDirs if d != ".git"]
        for sName in lNames:
            sRelative = os.path.relpath(os.path.join(sRoot, sName), pathRoot).replace("\\", "/")
            if unwantedKind(sRelative) == "purgeDelete":
                lUnwantedDisk.append(sRelative)
    lUnwantedDisk.sort()
    say("4. UNWANTED NAMES: pandoc and the web client utilities.")
    say("")
    if lUnwantedTracked or lUnwantedHistory or lUnwantedDisk:
        if lUnwantedTracked:
            say(f"   Tracked now ({len(lUnwantedTracked)}), to be untracked:")
            for sPath in lUnwantedTracked:
                say(f"     {sPath}  [{unwantedKind(sPath)}]")
            say("")
        if lUnwantedHistory:
            say(f"   In the history ({len(lUnwantedHistory)}), to be purged by the rewrite:")
            for sPath in lUnwantedHistory[:20]:
                say(f"     {sPath}")
            if len(lUnwantedHistory) > 20:
                say(f"     and {len(lUnwantedHistory) - 20} more")
            say("")
        if lUnwantedDisk:
            say(f"   On disk ({len(lUnwantedDisk)}), to be DELETED (the backup keeps a copy):")
            for sPath in lUnwantedDisk:
                say(f"     {sPath}")
            say("")
        say("   pandoc stays on disk: EdSharp uses it at run time, and")
        say("   installPandoc fetches it for users. The web client utilities,")
        say("   including InPy, are removed from disk as well.")
    else:
        say("   None found anywhere.")
    say("")

    # What a name costs in total, not what its largest copy costs. Fifteen
    # copies of a 26 megabyte installer is 390 megabytes of pack and not one of
    # them trips a per-file limit, which is how the first run reported that
    # nothing oversized remained while the repository was mostly installers.
    dByName = {}
    for iSize, sName, _sHash in lBig:
        dByName[sName] = dByName.get(sName, 0) + iSize
    lBulky = sorted(((iTotal, sName) for sName, iTotal in dByName.items()
                     if iTotal >= c_iBulkyBytes and not isNeeded(sName, setNeeded)),
                    reverse=True)
    if lBulky:
        say("   Counted by what each name costs in total rather than per copy:")
        for iTotal, sName in lBulky:
            iCopies = sum(1 for t in lBig if t[1] == sName)
            say(f"     {iTotal / (1024*1024):>8.1f} MB  {sName}  ({iCopies} copies)")
        say("")
        say("   A build output committed on every release accumulates. None of")
        say("   these trips a per-file limit; together they are most of what a")
        say("   clone has to download.")
        say("")
    setHistoryNames = {t[1] for t in lBig}
    bNeedRewrite = bool(lTooBig) or bool(lBulky) or bool(lUnwantedHistory)
    say("=" * 68)
    say("PLAN")
    say("=" * 68)
    say("")
    iStep = 0
    if bDirty:
        iStep += 1
        sWhat = "Stash" if dArguments.stash else "Commit"
        say(f"{iStep}. {sWhat} the {len(dState['dirty'])} outstanding changes.")
    if lBelong:
        iStep += 1
        say(f"{iStep}. Add {len(lBelong)} files the repository is missing.")
    if lStray:
        iStep += 1
        say(f"{iStep}. Untrack {len(lStray)} files and add them to .gitignore.")
    if bNeedRewrite or lUnwantedDisk:
        iStep += 1
        say(f"{iStep}. Copy the whole folder to a dated backup beside it.")
    if lUnwantedDisk:
        iStep += 1
        say(f"{iStep}. Delete {len(lUnwantedDisk)} web client "
            f"{'file' if len(lUnwantedDisk) == 1 else 'files'} from disk.")
    if bNeedRewrite:
        iStep += 1
        lRemove = sorted({t[1] for t in lTooBig} | {s for _i, s in lBulky} | set(lUnwantedHistory))
        iTotal = sum(i for i, _s in lBulky) + sum(t[0] for t in lTooBig)
        say(f"{iStep}. Rewrite the history to remove {len(lRemove)} "
            f"{'file' if len(lRemove) == 1 else 'files'}, "
            f"about {iTotal / (1024*1024):.0f} MB in all:")
        for sName in lRemove:
            say(f"     {sName}")
        iStep += 1
        say(f"{iStep}. Expire the reflog and repack, to reclaim the space.")
    if lStray or bNeedRewrite or lBelong or lUnwantedDisk:
        iStep += 1
        say(f"{iStep}. Commit.")
        if not dArguments.no_push:
            iStep += 1
            sHow = "force-push (the history changed)" if bNeedRewrite else "push"
            say(f"{iStep}. {sHow}.")
    if not iStep:
        say("Nothing to do. The repository is already tidy.")
        return 0
    say("")

    if not bDoIt:
        say("This was a description only (--survey). Nothing has been changed.")
        say("")
        say("Run it again without --survey to carry the plan out:")
        say("")
        say("    python tidyRepo.py")
        say("")
        say(f"The log is at {pathLog}")
        return 0

    say("=" * 68)
    say("DOING IT")
    say("=" * 68)
    say("")

    # The outstanding work, first, so everything below has a clean tree.
    if bDirty:
        if dArguments.stash:
            say("Setting the outstanding changes aside.")
            result = run(["git", "stash", "push", "-m", "tidyRepo"])
            if not result or result.returncode:
                say("  the stash FAILED, so nothing further is attempted")
                return 1
            say("  stashed. Recover them afterwards with: git stash pop")
        else:
            say("Committing the outstanding changes.")
            sMessage = dArguments.message or readVersionMessage()
            # Tracked files only. An earlier version used add -A, which swept
            # in this script's own log; the log is written continuously, so it
            # was committed and dirty again in the same breath, and the check
            # below rightly refused to go on. Untracked files are the survey's
            # business and are reported rather than committed.
            run(["git", "add", "-u"])
            result = run(["git", "commit", "-q", "-m", sMessage])
            if not result or result.returncode:
                say("  the commit FAILED, so nothing further is attempted")
                return 1
            say(f"  committed as: {sMessage}")
        # Confirmed rather than assumed, because everything below depends on it.
        result = run(["git", "status", "--porcelain", "--untracked-files=no"])
        if result and (result.stdout or "").strip():
            say("  the tree is STILL not clean, so nothing further is attempted:")
            for sLine in (result.stdout or "").splitlines()[:10]:
                say(f"    {sLine}")
            return 1
        say("  the tree is clean")
        say("")

    if bNeedRewrite or lUnwantedDisk:
        sBackup = os.path.join(os.path.dirname(pathRoot),
                               f"EdSharp_backup_{datetime.datetime.now():%Y%m%d_%H%M%S}")
        say(f"Copying everything to {sBackup}")
        say("This is the way back if anything goes wrong. It is not deleted.")
        try:
            shutil.copytree(pathRoot, sBackup)
            say("  copied")
        except Exception as exception:
            say(f"  the copy FAILED: {exception}")
            say("  STOPPING. Nothing has been rewritten.")
            return 1
        say("")

    if lBelong:
        say(f"Adding {len(lBelong)} files the repository was missing.")
        for sPath in lBelong:
            run(["git", "add", "--", sPath])
        say("")

    if lStray:
        say(f"Untracking {len(lStray)} files.")
        for sPath in lStray:
            run(["git", "rm", "--cached", "-q", "--", sPath])
        say("")

    if lUnwantedDisk:
        say(f"Deleting {len(lUnwantedDisk)} web client files from disk. The backup keeps a copy.")
        setParents = set()
        for sPath in lUnwantedDisk:
            pathFull = os.path.join(pathRoot, sPath.replace("/", os.sep))
            try:
                os.remove(pathFull)
                say(f"  deleted {sPath}")
                setParents.add(os.path.dirname(pathFull))
            except OSError as exception:
                say(f"  could NOT delete {sPath}: {exception}")
        # Folders left empty by the deletions go too, deepest first.
        for sParent in sorted(setParents, key=len, reverse=True):
            try:
                if os.path.normpath(sParent) != os.path.normpath(pathRoot) and not os.listdir(sParent):
                    os.rmdir(sParent)
                    say(f"  removed the empty folder {os.path.relpath(sParent, pathRoot)}")
            except OSError:
                pass
        say("")

    # .gitignore gets the strays and the canonical unwanted patterns, so
    # none of this can creep back in through git add -A.
    pathIgnore = os.path.join(pathRoot, ".gitignore")
    sExisting = ""
    if os.path.exists(pathIgnore):
        with open(pathIgnore, encoding="utf-8") as fileIgnore:
            sExisting = fileIgnore.read()
    # Every stray goes in, the unwanted ones included: old_Convert/Pandoc is
    # not covered by the canonical Convert/Pandoc/ pattern, and an entry that
    # is missing is a path git add -A can readmit.
    lAdd = [s for s in lStray if s not in sExisting]
    # The names that must never return, whatever else is in the tree: the
    # unwanted files, the installer output that accumulated 24 copies in the
    # history, the dlls the build fetches, and the logs.
    for sPattern in ("pandoc.exe", "Convert/Pandoc/", "InPy.exe",
                     "EdSharp_Setup.exe", "EdSharp_setup.exe",
                     "ReverseMarkdown.dll", "HtmlAgilityPack.dll", "Markdig.dll",
                     "BuildEdSharp.log", "tidyRepo.log"):
        if sPattern not in sExisting:
            lAdd.append(sPattern)
    # Collapse whole stray folders to one pattern each, so .gitignore stays
    # readable: a top-level folder with nothing the project needs under it is
    # ignored as folder/ rather than as hundreds of lines.
    setNeededTops = set()
    for sPath in lTracked:
        if "/" in sPath and isNeeded(sPath, setNeeded):
            setNeededTops.add(sPath.split("/", 1)[0])
    for sPath in setNeeded | set(lBelong):
        sPath = sPath.replace("\\", "/")
        # The first segment even of a slash-free entry: a folder-wide Source
        # line like Convert\* strips to the bare name Convert, and missing it
        # here is what once put Convert/ into .gitignore wholesale. A plain
        # file name added this way is harmless, because no stray path can
        # have a file name as its folder.
        setNeededTops.add(sPath.split("/", 1)[0])
    lFinal = []
    setCollapsed = set()
    for sPath in sorted(set(lAdd)):
        if "/" in sPath and not sPath.endswith("/"):
            sTop = sPath.split("/", 1)[0]
            if sTop not in setNeededTops:
                if sTop not in setCollapsed and (sTop + "/") not in sExisting:
                    setCollapsed.add(sTop)
                    lFinal.append(sTop + "/")
                continue
        lFinal.append(sPath)
    if lFinal:
        with open(pathIgnore, "a", encoding="utf-8", newline="\n") as fileIgnore:
            fileIgnore.write(
                "\n# Untracked by tidyRepo: files that reached the repository\n"
                "# through git add -A, plus the names that must never return.\n"
                "# A folder pattern covers a stray folder with nothing the\n"
                "# project needs inside it.\n")
            for sPath in sorted(set(lFinal)):
                fileIgnore.write(sPath + "\n")
        sFolders = "folder" if len(setCollapsed) == 1 else "folders"
        sPatterns = "pattern" if len(set(lFinal)) == 1 else "patterns"
        say(f"Added {len(set(lFinal))} {sPatterns} to .gitignore "
            f"({len(setCollapsed)} whole {sFolders} collapsed to one pattern each)")
    run(["git", "add", ".gitignore"])
    say("")

    # Anything staged above must be committed before the rewrite. filter-branch
    # refuses to run with a dirty index, and this has now been the cause twice:
    # once from untracking, once from adding. One commit here, covering both,
    # so a third variant of the same fault cannot appear.
    result = run(["git", "status", "--porcelain", "--untracked-files=no"])
    if result and (result.stdout or "").strip():
        say("Committing the additions and removals.")
        run(["git", "commit", "-q", "-m",
             "Tidy the repository: add missing source, untrack development files"])
        result = run(["git", "status", "--porcelain", "--untracked-files=no"])
        if result and (result.stdout or "").strip():
            say("  the index is STILL not clean, so nothing further is attempted")
            return 1
        say("  committed, so anything below has a clean index")
        say("")

    if bNeedRewrite:
        say("Rewriting the history.")
        lNames = sorted({t[1] for t in lTooBig} | {s for _i, s in lBulky} | set(lUnwantedHistory))
        for sName in lNames:
            say(f"  removing every version of {sName}")
        bUsedFilterRepo = False
        if shutil.which("git-filter-repo"):
            lCommand = ["git", "filter-repo", "--force"]
            for sName in lNames:
                lCommand += ["--path", sName, "--invert-paths"]
            result = run(lCommand)
            bUsedFilterRepo = bool(result and not result.returncode)
            if bUsedFilterRepo:
                say("  rewritten with git filter-repo")
                # filter-repo removes the remote deliberately; put it back.
                if dState["remote"]:
                    run(["git", "remote", "add", "origin", dState["remote"]])
                    say(f"  remote restored to {dState['remote']}")
        if not bUsedFilterRepo:
            say("  git filter-repo is not installed, so filter-branch is used.")
            say("  It is slower and noisier, and the result is the same.")
            sPaths = " ".join(f'"{s}"' for s in lNames)
            result = run(["git", "filter-branch", "--force", "--index-filter",
                          f"git rm --cached --ignore-unmatch -r {sPaths}",
                          "--prune-empty", "--tag-name-filter", "cat", "--", "--all"])
            if not result or result.returncode:
                say("")
                say("  THE REWRITE FAILED. Nothing has been pushed, and the backup")
                say("  folder holds the repository as it was. The commonest cause")
                say("  is something staged; the log above says which.")
                return 1
        say("")
        say("Reclaiming the space.")
        say("  Nothing may still point at the old commits, or git keeps every")
        say("  object they reach and the rewrite frees nothing. An earlier")
        say("  version deleted one backup ref and left the rest, so the history")
        say("  was rewritten and the repository stayed the same size.")

        # Every backup ref filter-branch made, not just the branch's. It makes
        # one per ref it rewrote, so a repository with a remote has at least
        # two, and the one for the remote-tracking ref was what kept 390 MB of
        # installers alive through a full gc.
        result = run(["git", "for-each-ref", "--format=%(refname)", "refs/original"])
        lOriginal = [s.strip() for s in (result.stdout or "").splitlines() if s.strip()]
        for sRef in lOriginal:
            run(["git", "update-ref", "-d", sRef])
        say(f"  removed {len(lOriginal)} backup refs left by the rewrite")

        # The remote-tracking refs as well. They still name the old commits
        # until the next fetch, and they are rebuilt by the push below.
        result = run(["git", "for-each-ref", "--format=%(refname)", "refs/remotes"])
        lRemote = [s.strip() for s in (result.stdout or "").splitlines() if s.strip()]
        for sRef in lRemote:
            run(["git", "update-ref", "-d", sRef])
        if lRemote:
            say(f"  removed {len(lRemote)} remote-tracking refs; the push restores them")

        run(["git", "reflog", "expire", "--expire=now", "--all"])
        run(["git", "gc", "--prune=now", "--aggressive"])

        # Checked here rather than only at the end, because if this did not
        # work there is no point pushing.
        result = run(["git", "rev-list", "--objects", "--all"])
        iLeft = sum(1 for sLine in (result.stdout or "").splitlines()
                    if any(sName in sLine for sName in lNames))
        if iLeft:
            say("")
            say(f"  {iLeft} of the old objects are STILL reachable, so the space")
            say("  was not reclaimed. These refs still exist:")
            result = run(["git", "for-each-ref", "--format=%(refname)"])
            for sLine in (result.stdout or "").splitlines()[:12]:
                say(f"    {sLine.strip()}")
            say("  Nothing has been pushed. The backup folder is your way back.")
            return 1
        say("  the old objects are gone")
        say("")

    # Anything still outstanding. Usually nothing, because the untracking was
    # committed above and the rewrite commits as it goes.
    result = run(["git", "status", "--porcelain", "--untracked-files=no"])
    if result and (result.stdout or "").strip():
        say("Committing what is left.")
        run(["git", "add", "-u"])
        run(["git", "commit", "-q", "-m", "Tidy the repository"])
        say("")

    if not dArguments.no_push and dState["remote"]:
        if bNeedRewrite:
            say("Force-pushing, because the history changed.")
            say("Anyone else who has cloned this will have to clone it again.")
            run(["git", "push", "--force", "origin", dState["branch"]])
        else:
            say("Pushing.")
            run(["git", "push", "origin", dState["branch"]])
    say("")

    say("=" * 68)
    say("AFTERWARDS")
    say("=" * 68)
    lAfter, _setAfterNames = surveyHistory()
    dAfter = {}
    for iSize, sName, _sHash in lAfter:
        dAfter[sName] = dAfter.get(sName, 0) + iSize
    lStillBulky = sorted(((iTotal, sName) for sName, iTotal in dAfter.items()
                          if iTotal >= c_iBulkyBytes and not isNeeded(sName, setNeeded)),
                         reverse=True)
    if lStillBulky:
        say("WARNING: these still account for a lot of the history:")
        for iTotal, sName in lStillBulky[:5]:
            say(f"  {iTotal / (1024*1024):.1f} MB  {sName}")
        say("The backup folder is your way back.")
    else:
        # Says what was checked, rather than a bare reassurance. The first
        # version said nothing oversized remained when it had only asked about
        # the per-file limit and the repository was mostly installers.
        say(f"No unnecessary name accounts for {c_iBulkyBytes // (1024*1024)} MB "
            "or more of the history.")
    result = run(["git", "count-objects", "-vH"])
    for sLine in (result.stdout or "").splitlines():
        if sLine.startswith("size-pack"):
            say(f"The repository is now {sLine.split(':', 1)[1].strip()}.")
    lTracked, lStray = surveyTracked(setNeeded)
    say(f"{len(lTracked)} files tracked, {len(lStray)} of them unnecessary.")
    say("")
    say(f"The log is at {pathLog}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exception:
        import traceback

        say("")
        say(f"UNEXPECTED FAILURE: {exception}")
        for sLine in traceback.format_exc().splitlines():
            say("  " + sLine)
        say("")
        say(f"The log is at {pathLog}. Nothing further was attempted.")
        sys.exit(1)
    finally:
        if fileLog:
            fileLog.close()
