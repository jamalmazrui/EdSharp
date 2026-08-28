"""repoPolicy.py -- what belongs in the public EdSharp repository.

Two programs need this answer and they must not disagree about it:
tidyRepo.py, which untracks whatever does not belong, and auditEdSharp.py,
which fails the build when something that does not belong has come back. So
the answer lives here, in one list, and both of them import it.

WHY THIS FILE EXISTS

For weeks the repository carried saved web pages, saved emails and old
drafts, and tidyRepo reported nothing wrong. The cause was a single line in
its isNeeded test:

    if sPath.endswith((".md", ".htm")) and "/" not in sPath:
        return True

It was meant to spare the documentation set. What it actually said was that
any file at the top of the folder whose name ends in .md or .htm is part of
the project. Every Stack Overflow capture, every saved mailing list message,
every old draft ends in .htm or .md and sits at the top of the folder. So
the survey declared 38 of them needed, found nothing to untrack, and printed
a clean report. .gitignore could not help either, because .gitignore has no
effect on a file that is already tracked; adding a name to it after the fact
changes nothing at all.

THE RULE THAT REPLACES IT

A file belongs in the public repository only if it is named. There is no
pattern that admits a file by the look of its name. There are two ways to be
named:

  1. The installer names it. EdSharp_Setup.iss is the list of what EdSharp
     ships, and a file on that list is part of the project by definition.
     A folder-wide Source line, such as Convert\\*, covers everything under
     that folder.
  2. This file names it, in c_lDevelopmentFiles below. These are the files
     that build, audit and release EdSharp. They are not installed on a
     user's machine, but a clone without them cannot build EdSharp, so they
     are tracked.

Anything else is a development aid. It stays on disk, where it is useful,
and out of the repository, where it is not. Handover.md is the clearest
example: it is written for the next conversation, not for a user, so it is
deliberately absent from both lists.

TO CHANGE WHAT THE REPOSITORY CARRIES

Add a Source line to EdSharp_Setup.iss if EdSharp should install the file.
Add the name to c_lDevelopmentFiles if the build needs it. Doing neither is
how a file is kept out. There is no third place to look.
"""

import os
import re

# Files the repository tracks that the installer does not ship: everything
# needed to build, check and release EdSharp from a fresh clone. In order by
# name, so a reader can find one and so an addition has an obvious home.
c_lDevelopmentFiles = [
    ".gitattributes",
    ".gitignore",
    "BuildEdSharp.ps1",
    "applyConvertPolicy.cmd",
    "applyConvertPolicy.py",
    "auditEdSharp.cmd",
    "auditEdSharp.py",
    "inixVert.cs",
    "prepareAuditFixes.cmd",
    "repoPolicy.py",
    "sqlean.version",
    "tagRelease.cmd",
    "tagRelease.ps1",
    "tidyRepo.cmd",
    "tidyRepo.py",
]


def installedFiles(pathRoot):
    """What EdSharp_Setup.iss says the project ships.

    Returns a set of exact names and a list of folder prefixes, or None when
    the setup script cannot be read -- in which case nothing can be judged
    and the caller should say so rather than guess.
    """
    pathIss = os.path.join(pathRoot, "EdSharp_Setup.iss")
    if not os.path.isfile(pathIss):
        return None
    with open(pathIss, encoding="utf-8-sig", errors="replace") as fileIss:
        sIss = fileIss.read()
    setExact, lFolders = set(), []
    for oMatch in re.finditer(r'^Source:\s*"([^"]+)"', sIss, re.M):
        sPath = oMatch.group(1).replace("C:\\EdSharp\\", "").replace("\\", "/")
        if sPath.endswith("/*") or sPath.endswith("*"):
            # A folder-wide line such as Convert/* covers the whole tree.
            sFolder = sPath.rstrip("*").rstrip("/")
            if sFolder:
                lFolders.append(sFolder + "/")
            continue
        setExact.add(sPath)
    return setExact, lFolders


def isNamedByInstaller(sPath, setExact, lFolders):
    """Whether the setup script ships this exact path, by name or by folder."""
    sPath = sPath.replace("\\", "/")
    if sPath in setExact:
        return True
    for sFolder in lFolders:
        if sPath.startswith(sFolder):
            return True
    return False


def isDevelopmentFile(sPath):
    """Whether this path is one of the named build and release files."""
    return sPath.replace("\\", "/") in c_lDevelopmentFiles


def isRepoFile(sPath, setExact, lFolders):
    """Whether a path belongs in the public repository at all.

    The whole rule, in one place: named by the installer, or named as a
    development file. Nothing is admitted by the look of its name.
    """
    return isNamedByInstaller(sPath, setExact, lFolders) or isDevelopmentFile(sPath)


def strayFiles(lPaths, setExact, lFolders):
    """The paths in a list that do not belong, in order by name."""
    return sorted(s for s in lPaths
                  if not isRepoFile(s.replace("\\", "/"), setExact, lFolders))
