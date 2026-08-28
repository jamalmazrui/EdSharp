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
     that folder. The exception is c_lFetchedFiles below: libraries the
     installer ships but the build downloads, which a clone fetches rather
     than carries.
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
    "moveNotes.cmd",
    "moveNotes.py",
    "prepareAuditFixes.cmd",
    "repoPolicy.py",
    "restoreMissing.cmd",
    "restoreMissing.py",
    "sqlean.version",
    "tagRelease.cmd",
    "tagRelease.ps1",
    "tidyRepo.cmd",
    "tidyRepo.py",
]


# Named by the installer, but fetched by the build rather than kept in the
# repository. The distinction matters: the installer must ship them, so they
# have Source lines, yet a clone should download them rather than carry them.
#
# This list exists because all four were committed on 28 August. .gitignore
# held their names, which kept them out; .gitignore was lost, and with the
# only thing standing between them and the repository gone, tidyRepo read
# their Source lines, saw them untracked, and added them. Naming them here
# puts the answer where the policy is, so no missing file can change it.
c_lFetchedFiles = [
    "HtmlAgilityPack.dll",
    "Markdig.dll",
    "ReverseMarkdown.dll",
    "sqlean.dll",
]


def isFetched(sPath):
    """Whether the build downloads this rather than the repository holding it."""
    return sPath.replace("\\", "/") in c_lFetchedFiles


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
        # A fetched library is shipped but never tracked, so it is not part
        # of what the repository claims.
        if not isFetched(sPath):
            setExact.add(sPath)
    return setExact, lFolders


def wholesaleFolders(pathRoot):
    """Folders the installer ships entirely, with nothing held back.

    A Source line with an Excludes clause does not qualify: Convert\\* ships
    the tree but leaves out sources, project files and archives, so a file
    found under it is not automatically wanted. Scripts, Snippets, Samples
    and Dictionaries have no Excludes, so everything in them is shipped and
    everything in them belongs in the repository.
    """
    pathIss = os.path.join(pathRoot, "EdSharp_Setup.iss")
    if not os.path.isfile(pathIss):
        return set()
    with open(pathIss, encoding="utf-8-sig", errors="replace") as fileIss:
        sIss = fileIss.read()
    setFolders = set()
    for oMatch in re.finditer(r'^Source:\s*"([^"]+)"([^\n]*)$', sIss, re.M):
        sPath, sRest = oMatch.group(1), oMatch.group(2)
        if "*" not in sPath:
            continue
        if "excludes:" in sRest.lower():
            continue
        sFolder = sPath.replace("\\", "/").rstrip("*").rstrip("/")
        if sFolder:
            setFolders.add(sFolder.lower())
    return setFolders


def isUnderWholesaleFolder(sPath, setWholesale):
    """Whether a path sits inside a folder the installer ships entirely."""
    sLower = sPath.replace("\\", "/").lower()
    return any(sLower.startswith(s + "/") for s in setWholesale)


def isNamedByInstaller(sPath, setExact, lFolders):
    """Whether the setup script ships this exact path, by name or by folder.

    Folder matching ignores case, because Windows does. The setup script says
    Scripts\\*; the folder on disk is called scripts; and a case-sensitive test
    called every JAWS script a stray, which is why they were ignored and have
    been absent from the repository. File names are still matched exactly, so
    that git's own view of a name is never contradicted.
    """
    sPath = sPath.replace("\\", "/")
    if sPath in setExact:
        return True
    sLower = sPath.lower()
    for sFolder in lFolders:
        if sLower.startswith(sFolder.lower()):
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
