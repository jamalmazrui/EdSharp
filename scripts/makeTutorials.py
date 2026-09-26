"""makeTutorials.py -- write the walkthrough section of Tutorials.md.

WHAT IT DOES

Reads Tutorial.inix, which holds one step to a section: the narration, the key
to press, what a screen reader says in answer, and a note for the reader. Writes
that out as Markdown, into Tutorials.md, between two markers.

WHY A GENERATED SECTION RATHER THAN A NEW DOCUMENT

Tutorials.md already ships, is already converted to HTML by the build, and is
already where somebody looks for a tutorial. A new document would need a line in
the installer script, a line in the repository list, and a place in every list of
documents. The walkthrough is a section of the tutorials, so it lives in the
tutorials.

Everything between the markers is replaced each time. Everything outside them is
left exactly as it was, so the hand-written tutorials and this one share a file
without either disturbing the other.

RUN IT WITH NO ARGUMENTS. The log is written beside this script.
"""

import datetime
import os
import subprocess
import platform
import re
import sys
import traceback

c_sScript = os.path.abspath(__file__)
c_sTool = os.path.dirname(c_sScript)
# THE SCRIPTS LIVE IN help, THE TOOLING IN scripts. FileDir keeps both in one
# folder; the Homer layout does not, and this file was copied from there. So
# c_sHere is where the tutorials and their audio are, and the log stays with the
# tool that writes it.
c_sHere = os.path.join(os.path.dirname(c_sTool), "help")
if not os.path.isdir(c_sHere):
    c_sHere = c_sTool
import datetime as _dt
c_sLogDir = os.path.join(os.path.dirname(c_sTool), "logs")
os.makedirs(c_sLogDir, exist_ok=True)
c_sLog = os.path.join(c_sLogDir, os.path.basename(os.path.dirname(c_sTool)) + "-tutorials-text-%s.log" % _dt.datetime.now().strftime("%Y%m%d-%H%M%S"))
import glob

# EVERY SCRIPT IN THE FOLDER, in name order. Tutorial.inix is the first walk;
# Tutorial_Tagging.inix, Tutorial_Zipping.inix and the rest are one task each.
# Adding a tutorial means adding a file, and nothing else.
c_sPattern = os.path.join(c_sHere, "Tutorial*.inix")
c_sTarget = os.path.join(c_sHere, "Tutorials.md")

c_sFeed = os.path.join(c_sHere, "TutorialFeed.xml")

# The least a document can be and still take a spliced walkthrough: a title, a
# contents list, and a section 1 to splice in front of.
c_sSkeleton = """# Tutorials

## Contents

- [1. Where to Go Next](#1-where-to-go-next)

## 1. Where to Go Next

Press F1 for the guide, Alt+Shift+H for the hotkey list, and Alt+F10 for every
command in one window.
"""
c_sStartMark = "<!-- walkthrough: written by makeTutorials.py, do not edit between the markers -->"
c_sEndMark = "<!-- walkthrough ends -->"


def note(sText):
    """One line to the log, and nothing to the screen unless it matters."""
    try:
        with open(c_sLog, "a", encoding="utf-8") as oFile:
            oFile.write(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S  ") + sText + "\n")
    except Exception:
        pass


def say(sText):
    """A short plain sentence on the screen, and the same line in the log."""
    print(sText)
    note(sText)


def startLog():
    try:
        if os.path.isfile(c_sLog):
            os.remove(c_sLog)
    except Exception:
        pass
    note("makeTutorial " + datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    note("script: " + c_sScript)
    note("Python: " + platform.python_version())
    note("platform: " + sys.platform)
    note("working directory: " + os.getcwd())
    note("command line: " + " ".join(sys.argv))
    note("scripts: " + c_sPattern)
    note("target: " + c_sTarget)


def readFile(sPath):
    """Text of a file, whatever byte order mark it carries."""
    with open(sPath, "r", encoding="utf-8-sig") as oFile:
        return oFile.read().replace("\r\n", "\n")


def writeFile(sPath, sText):
    """Homer standard: UTF-8 with a byte order mark, and CRLF line endings."""
    with open(sPath, "wb") as oFile:
        oFile.write(b"\xef\xbb\xbf" + sText.replace("\n", "\r\n").encode("utf-8"))


def readSteps(sText):
    """The sections of Tutorial.inix, in order.

    A tolerant reader rather than a strict one: a key may repeat, and repeats
    are collected in order, which is how a step holds several lines of speech.
    """
    lsSections = []
    dNow = None
    for sRaw in sText.split("\n"):
        sLine = sRaw.strip()
        if len(sLine) == 0 or sLine.startswith(";") or sLine.startswith("#"):
            continue
        if sLine.startswith("[") and sLine.endswith("]"):
            dNow = {"_name": sLine[1:-1].strip().lower()}
            lsSections.append(dNow)
            continue
        if dNow is None or "=" not in sLine:
            continue
        sField, sValue = sLine.split("=", 1)
        sField = sField.strip()
        sValue = sValue.strip()
        if len(sValue) == 0:
            continue
        dNow.setdefault(sField, []).append(sValue)
    return lsSections


def firstOf(dSection, sField):
    lsValues = dSection.get(sField, [])
    return lsValues[0] if lsValues else ""


def buildMarkdown(lsSections, sNumber):
    """One tutorial, as Markdown, numbered as the caller says."""
    dAbout = {}
    lsSteps = []
    for dSection in lsSections:
        if dSection["_name"] == "about":
            dAbout = dSection
        elif dSection["_name"] == "step":
            lsSteps.append(dSection)

    lsOut = []
    sTitle = firstOf(dAbout, "Title") or "A Walk Through This Program"
    # A TITLE THAT ALREADY CARRIES ITS NUMBER KEEPS IT ALONE. The generated
    # counter -- 0, 0a, 0b -- existed because titles had no numbers of their own;
    # with "01 - Installing DbDo" it produced "0a. 01 - Installing DbDo", which
    # reads as two numbering schemes arguing.
    if re.match(r"^\d", sTitle):
        lsOut.append("## " + sTitle)
    else:
        lsOut.append("## " + sNumber + ". " + sTitle)
    lsOut.append("")
    sIntro = firstOf(dAbout, "Intro")
    if sIntro:
        lsOut.append(sIntro)
        lsOut.append("")
    sSetup = firstOf(dAbout, "Setup")
    if sSetup:
        lsOut.append("**Before you start:** " + sSetup)
        lsOut.append("")

    iNumber = 0
    for dStep in lsSteps:
        iNumber += 1
        sSay = firstOf(dStep, "Say")
        sKey = firstOf(dStep, "Key")
        lsHear = dStep.get("Hear", [])
        sNote = firstOf(dStep, "Note")

        lsOut.append("### Step " + str(iNumber) + (": " + sKey if sKey else ""))
        lsOut.append("")
        if sSay:
            lsOut.append(sSay)
            lsOut.append("")
        if lsHear:
            # THE SPEAKER IS LABELLED, NOT NAMED.
            #
            # Podcast tutorials in this field label their speakers -- Jonathan
            # Mosen's transcripts read "Jonathan:" then "VoiceOver:" or "JAWS:"
            # -- so a reader always knows which voice said what. A person
            # demonstrating needs no name here, and the reader is whichever one
            # you use, so the label says what it is rather than who.
            lsOut.append("Screen reader:")
            lsOut.append("")
            for sHeard in lsHear:
                lsOut.append("- " + sHeard)
            lsOut.append("")
        if sNote:
            lsOut.append(sNote)
            lsOut.append("")

    sHomework = firstOf(dAbout, "Homework")
    if sHomework:
        lsOut.append("**Something to try:** " + sHomework)
        lsOut.append("")

    return "\n".join(lsOut).rstrip()


def spliceIntoTutorials(sDocument, sSection):
    """The walkthrough put back into Tutorials.md.

    Between the markers if they are there; after the contents list if they are
    not, which is where a first tutorial belongs.
    """
    iStart = sDocument.find(c_sStartMark)
    iEnd = sDocument.find(c_sEndMark)
    if iStart >= 0 and iEnd > iStart:
        return sDocument[:iStart] + sSection + sDocument[iEnd + len(c_sEndMark):]

    sAnchor = "\n## 1. "
    iAnchor = sDocument.find(sAnchor)
    if iAnchor < 0:
        note("no place found for the walkthrough; appending it")
        return sDocument.rstrip() + "\n\n" + sSection + "\n"
    return sDocument[:iAnchor] + "\n" + sSection + "\n" + sDocument[iAnchor:]


def addToContents(sDocument, sSection):
    """Contents lines for the generated tutorials, in front of the first one.

    Taken from the headings just written, so the list and the sections cannot
    disagree, and a tutorial added as a file needs nothing else done to it.
    """
    lsEntries = []
    for sLine in sSection.split("\n"):
        if not sLine.startswith("## "):
            continue
        sTitle = sLine[3:].strip()
        sAnchor = "#" + "".join(
            (ch.lower() if ch.isalnum() else ("-" if ch in " ." else ""))
            for ch in sTitle).strip("-")
        while "--" in sAnchor:
            sAnchor = sAnchor.replace("--", "-")
        lsEntries.append("- [" + sTitle + "](" + sAnchor + ")")
    if not lsEntries:
        return sDocument

    # Old generated lines go, so nothing accumulates across runs.
    lsKept = [sLine for sLine in sDocument.split("\n")
              if not (sLine.startswith("- [0") and "](#0" in sLine)]
    sDocument = "\n".join(lsKept)

    # THE FIRST HAND-WRITTEN CONTENTS LINE, whatever it is called. This was
    # FileDir's own section title, which meant the contents list silently did
    # not update in any other app: the sections appeared and nothing pointed at
    # them. Any line starting "- [1. " does the job.
    iAt = -1
    for sLine in sDocument.split("\n"):
        if sLine.startswith("- [1. "):
            iAt = sDocument.find(sLine)
            break
    if iAt < 0:
        note("no contents list found; leaving it alone")
        return sDocument
    return sDocument[:iAt] + "\n".join(lsEntries) + "\n" + sDocument[iAt:]


def secondsOf(sPath):
    """How long an audio file runs, asked of ffprobe.

    A duration of zero in a feed is worse than no duration at all: a player
    shows it, and it is a lie. So the length is measured rather than guessed,
    and left out when it cannot be.
    """
    sProbe = os.path.join(c_sTool, "ffprobe.exe")
    if not os.path.isfile(sProbe):
        sProbe = "ffprobe"
    try:
        oResult = subprocess.run(
            [sProbe, "-v", "error", "-show_entries", "format=duration",
             "-of", "default=noprint_wrappers=1:nokey=1", sPath],
            capture_output=True, text=True, timeout=30)
        dSeconds = float(oResult.stdout.strip())
        return int(round(dSeconds))
    except Exception as oError:
        note("could not measure " + sPath + ": " + str(oError))
        return 0


def clockOf(iSeconds):
    return "%02d:%02d:%02d" % (iSeconds // 3600, (iSeconds % 3600) // 60, iSeconds % 60)


def anchorOf(sTitle):
    """The anchor Pandoc gives a heading, so a link lands on the right one."""
    sAnchor = "".join((ch.lower() if ch.isalnum() else ("-" if ch in " ." else "")) for ch in sTitle)
    while "--" in sAnchor:
        sAnchor = sAnchor.replace("--", "-")
    return sAnchor.strip("-")


def writeFeed(lsEpisodes, dFeed):
    """A podcast feed for the walkthroughs, so they can be subscribed to.

    Only the tutorials that have been spoken appear: an item without audio is
    not an episode. Each one links to its transcript, which is the section of
    Tutorials.htm the same run wrote.
    """
    if not lsEpisodes:
        note("no audio files found, so no feed is written")
        return 0
    sBase = dFeed.get("Base", "")
    if sBase and not sBase.endswith("/"):
        sBase += "/"
    sNow = datetime.datetime.now(datetime.timezone.utc).strftime("%a, %d %b %Y %H:%M:%S +0000")

    lsOut = []
    lsOut.append('<?xml version="1.0" encoding="UTF-8"?>')
    lsOut.append('<rss version="2.0" xmlns:itunes="http://www.itunes.com/dtds/podcast-1.0.dtd">')
    lsOut.append("  <channel>")
    lsOut.append("    <title>" + escape(dFeed.get("Title", "Homer Tools Walkthroughs")) + "</title>")
    lsOut.append("    <link>" + escape(sBase or "Tutorials.htm") + "</link>")
    lsOut.append("    <description>" + escape(dFeed.get("Description", "")) + "</description>")
    lsOut.append("    <language>en-us</language>")
    lsOut.append("    <lastBuildDate>" + sNow + "</lastBuildDate>")
    lsOut.append("    <itunes:author>" + escape(dFeed.get("Author", "")) + "</itunes:author>")
    lsOut.append("    <itunes:explicit>false</itunes:explicit>")
    lsOut.append('    <itunes:category text="Technology"/>')
    if dFeed.get("Email", ""):
        lsOut.append("    <itunes:owner><itunes:name>" + escape(dFeed.get("Author", ""))
                     + "</itunes:name><itunes:email>" + escape(dFeed["Email"]) + "</itunes:email></itunes:owner>")
    for iAt, dEpisode in enumerate(lsEpisodes, 1):
        sTranscript = (sBase or "") + "Tutorials.htm#" + anchorOf(dEpisode["number"] + ". " + dEpisode["title"])
        lsOut.append("    <item>")
        lsOut.append("      <title>" + escape(dEpisode["title"]) + "</title>")
        lsOut.append("      <description>" + escape(dEpisode["intro"] + " Full transcript: " + sTranscript)
                     + "</description>")
        lsOut.append("      <link>" + escape(sTranscript) + "</link>")
        lsOut.append('      <guid isPermaLink="false">homer-' + escape(dEpisode["stem"].lower()) + "</guid>")
        lsOut.append("      <enclosure url=\"" + escape((sBase or "") + dEpisode["audio"])
                     + "\" length=\"" + str(dEpisode["bytes"]) + "\" type=\"audio/mpeg\"/>")
        if dEpisode["seconds"] > 0:
            lsOut.append("      <itunes:duration>" + clockOf(dEpisode["seconds"]) + "</itunes:duration>")
        lsOut.append("      <itunes:explicit>false</itunes:explicit>")
        lsOut.append("    </item>")
    lsOut.append("  </channel>")
    lsOut.append("</rss>")
    writeFile(c_sFeed, "\n".join(lsOut) + "\n")
    return len(lsEpisodes)


def escape(sText):
    return (sText or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def main():
    startLog()
    lsSources = sorted(glob.glob(c_sPattern))
    note("scripts found: " + str(len(lsSources)) + " -- " + ", ".join(os.path.basename(s) for s in lsSources))
    if len(lsSources) == 0:
        say("0 tutorial scripts here, so there is nothing to write.")
        return 1
    if not os.path.isfile(c_sTarget):
        # A PROJECT WITH NO DOCUMENT YET GETS ONE, rather than an error.
        # The generated walkthroughs are spliced in ahead of section 1, so a
        # document has to exist and has to have a section 1 to aim at. Writing
        # a skeleton here means a new app needs no manual step before its first
        # tutorial builds -- which is the rule every Homer build script follows.
        note("no target; writing a skeleton: " + c_sTarget)
        say("Tutorials.md was not here, so a starting one was written.")
        try:
            with open(c_sTarget, "wb") as oFile:
                oFile.write(("\ufeff" + c_sSkeleton).replace("\n", "\r\n").encode("utf-8"))
        except Exception as oError:
            say("Tutorials.md could not be written.")
            note("write failed: " + str(oError))
            return 1

    lsBlocks = []
    lsEpisodes = []
    dFeed = {}
    iStepsAll = 0
    iLetter = 0
    for sSource in lsSources:
        try:
            lsSections = readSteps(readFile(sSource))
        except Exception as oError:
            say(os.path.basename(sSource) + " could not be read.")
            note("read failed: " + str(oError))
            note(traceback.format_exc())
            return 1
        iSteps = len([d for d in lsSections if d["_name"] == "step"])
        note(os.path.basename(sSource) + ": sections " + str(len(lsSections)) + ", steps " + str(iSteps))
        if iSteps == 0:
            say(os.path.basename(sSource) + " holds 0 steps.")
            return 1
        iStepsAll += iSteps
        # The first walk is section 0; the task tutorials are 0a, 0b and so on,
        # which keeps them together at the front without renumbering anything
        # that was already written.
        sNumber = "0" if len(lsBlocks) == 0 else "0" + chr(ord("a") + iLetter)
        if len(lsBlocks) > 0:
            iLetter += 1
        lsBlocks.append(buildMarkdown(lsSections, sNumber))

        dAbout = {}
        for dSection in lsSections:
            if dSection["_name"] == "about":
                dAbout = dSection
            elif dSection["_name"] == "feed" and not dFeed:
                for sKey in dSection:
                    if sKey != "_name":
                        dFeed[sKey] = dSection[sKey][0]
        sStem = os.path.splitext(os.path.basename(sSource))[0]
        # The audio lives in help\tutorials, one .mp3 per script (25 Sep 2026).
        sAudio = os.path.join(c_sHere, "tutorials", sStem + ".mp3")
        if os.path.isfile(sAudio):
            lsEpisodes.append({
                "stem": sStem,
                "number": sNumber,
                "title": firstOf(dAbout, "Title") or sStem,
                "intro": firstOf(dAbout, "Setup") or firstOf(dAbout, "Intro"),
                "audio": "tutorials/" + sStem + ".mp3",
                "bytes": os.path.getsize(sAudio),
                "seconds": secondsOf(sAudio)})
        else:
            note("no audio yet for " + sStem + "; it will join the feed once buildTutorials has run")

    try:
        sSection = c_sStartMark + "\n\n" + "\n\n".join(lsBlocks) + "\n\n" + c_sEndMark
        sDocument = readFile(c_sTarget)
        sDocument = addToContents(spliceIntoTutorials(sDocument, sSection), sSection)
        writeFile(c_sTarget, sDocument)
    except Exception as oError:
        say("Tutorials.md could not be written.")
        note("write failed: " + str(oError))
        note(traceback.format_exc())
        return 1

    say("Wrote " + str(len(lsSources)) + " tutorials into Tutorials.md: " + str(iStepsAll) + " steps.")
    try:
        iFeed = writeFeed(lsEpisodes, dFeed)
        if iFeed > 0:
            say("Wrote TutorialFeed.xml: " + str(iFeed) + " with audio.")
    except Exception as oError:
        note("feed failed: " + str(oError))
        note(traceback.format_exc())
        say("The feed could not be written. The log has why.")
    note("finished")
    return 0


if __name__ == "__main__":
    sys.exit(main())
