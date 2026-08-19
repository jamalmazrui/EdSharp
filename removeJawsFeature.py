r"""removeJawsFeature.py -- take the JAWS script installer OUT of EdSharp.exe.

Run it from C:\EdSharp with no arguments:

    python removeJawsFeature.py

WHY

Installing screen reader scripts is the installer's job, done since today by
installJawsScripts.ps1 on the HomerView model. The old route -- EdSharp.exe
--install-jaws-settings and the JawsScripts class behind it -- is now dead
code, and dead code that shows message boxes is worth removing rather than
leaving to surprise someone later.

WHAT IT REMOVES FROM EdSharp.cs

  1. In Main: the argument loop that handles --install-jaws-settings,
     together with the comment lines directly above it -- and, if the earlier
     quietJaws patch was applied, its --quiet lines as well.
  2. The whole JawsScripts class, wherever it sits, found by matching the
     braces of its body with a scanner that understands C# strings, verbatim
     strings, character literals, and comments, so a brace inside a string
     cannot derail it.

The edit is applied once and is safe to re-run: when the argument is already
gone, the script reports so and changes nothing. A backup EdSharp.cs.bak is
written before anything changes, so one file copy undoes everything. After
running this, run buildEdSharp; the compile is the proof the surgery was
clean.

A detailed log is written beside this script, whatever happens.
"""

import datetime
import os
import re
import sys

c_sBackupName = "EdSharp.cs.bak"
c_sLogName = "removeJawsFeature.log"
c_sSourceName = "EdSharp.cs"

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


def stripCodeOnly(sLine, dState):
    """Return the line with string/char/comment contents blanked, so braces
    can be counted on what the compiler treats as code.

    dState carries one flag between lines: whether a /* comment is open.
    Verbatim strings (@"...") spanning lines are not handled, because the
    EdSharp source does not use them across lines; a mismatch surfaces as an
    unbalanced-brace failure, which stops the script rather than corrupting
    the file.
    """
    lOut = []
    bBlock = dState["bBlockComment"]
    i = 0
    iLength = len(sLine)
    while i < iLength:
        sChar = sLine[i]
        if bBlock:
            if sChar == "*" and i + 1 < iLength and sLine[i + 1] == "/":
                bBlock = False
                i += 2
            else:
                i += 1
            continue
        if sChar == "/" and i + 1 < iLength and sLine[i + 1] == "/":
            break
        if sChar == "/" and i + 1 < iLength and sLine[i + 1] == "*":
            bBlock = True
            i += 2
            continue
        if sChar == "@" and i + 1 < iLength and sLine[i + 1] == '"':
            i += 2
            while i < iLength:
                if sLine[i] == '"':
                    if i + 1 < iLength and sLine[i + 1] == '"':
                        i += 2
                        continue
                    i += 1
                    break
                i += 1
            continue
        if sChar == '"':
            i += 1
            while i < iLength:
                if sLine[i] == "\\":
                    i += 2
                    continue
                if sLine[i] == '"':
                    i += 1
                    break
                i += 1
            continue
        if sChar == "'":
            i += 1
            while i < iLength:
                if sLine[i] == "\\":
                    i += 2
                    continue
                if sLine[i] == "'":
                    i += 1
                    break
                i += 1
            continue
        lOut.append(sChar)
        i += 1
    dState["bBlockComment"] = bBlock
    return "".join(lOut)


def findBlockEnd(lLines, iStart, dState):
    """The index of the line where braces opened from iStart return to zero.

    Counting starts at the first opening brace at or after iStart. The
    comment/string scanner keeps braces inside literals from being counted.
    """
    iDepth = 0
    bStarted = False
    for i in range(iStart, len(lLines)):
        sCode = stripCodeOnly(lLines[i], dState)
        for sChar in sCode:
            if sChar == "{":
                iDepth += 1
                bStarted = True
            elif sChar == "}":
                iDepth -= 1
        if bStarted and iDepth <= 0:
            return i
    return -1


def commentStart(lLines, iLine):
    """The first line of the contiguous // comment block directly above."""
    i = iLine
    while i > 0 and lLines[i - 1].lstrip().startswith("//"):
        i -= 1
    return i


def main():
    global fileLog
    fileLog = open(pathLog, "w", encoding="utf-8")
    say(f"removeJawsFeature  {datetime.datetime.now():%Y-%m-%d %H:%M:%S}")
    say(f"  script:            {os.path.abspath(__file__)}")
    say(f"  Python:            {sys.version.split()[0]}")
    say(f"  platform:          {sys.platform}")
    say(f"  working directory: {os.getcwd()}")
    say(f"  command line:      {' '.join(sys.argv)}")
    say("")

    pathSource = os.path.join(pathRoot, c_sSourceName)
    if not os.path.isfile(pathSource):
        say(f"FAILED: {pathSource} was not found. Run this from C:\\EdSharp.")
        return 1
    with open(pathSource, "rb") as fileSource:
        binData = fileSource.read()
    bBom = binData.startswith(b"\xef\xbb\xbf")
    sText = binData.decode("utf-8-sig")
    sBreak = "\r\n" if "\r\n" in sText else "\n"
    lLines = sText.replace("\r\n", "\n").split("\n")
    say(f"Read {c_sSourceName}: {len(binData)} bytes, {len(lLines)} lines.")

    if not any("--install-jaws-settings" in s for s in lLines):
        say("Already applied: the file no longer mentions --install-jaws-settings.")
        say("If EdSharp.exe still handles it, run buildEdSharp so the change is compiled in.")
        return 0

    lDrop = set()

    # Part one: the Main handler. The argument loop is the foreach nearest
    # above the line that names the argument.
    lArgHits = [i for i, s in enumerate(lLines) if '"--install-jaws-settings"' in s]
    say(f"Argument lines: {len(lArgHits)} at {[i + 1 for i in lArgHits]}")
    if len(lArgHits) != 1:
        say("FAILED: expected exactly one handler; the file differs from what was reviewed. Send this log.")
        return 1
    lLoopHits = [i for i, s in enumerate(lLines)
                 if i < lArgHits[0] and "foreach" in s and "cmdLineArgs" in s and "sArg" in s and "sQuietArg" not in s]
    if not lLoopHits:
        say("FAILED: no argument loop found above the handler. Send this log.")
        return 1
    iLoop = lLoopHits[-1]
    iLoopEnd = findBlockEnd(lLines, iLoop, {"bBlockComment": False})
    if iLoopEnd < 0:
        say("FAILED: the handler loop's braces did not balance. Send this log.")
        return 1
    iFrom = commentStart(lLines, iLoop)
    say(f"Removing the Main handler: lines {iFrom + 1}-{iLoopEnd + 1} (comments included).")
    lDrop.update(range(iFrom, iLoopEnd + 1))

    # The quietJaws lines, when that patch was applied: the bQuiet flag, the
    # sQuietArg loop, and their comments.
    lQuietHits = [i for i, s in enumerate(lLines) if "bool bQuiet" in s]
    if lQuietHits:
        iQuiet = lQuietHits[0]
        lQuietLoop = [i for i, s in enumerate(lLines) if i >= iQuiet and "sQuietArg" in s and "foreach" in s]
        if lQuietLoop:
            iQuietEnd = findBlockEnd(lLines, lQuietLoop[0], {"bBlockComment": False})
            iQuietFrom = commentStart(lLines, iQuiet)
            if iQuietEnd > 0:
                say(f"Removing the quietJaws lines: {iQuietFrom + 1}-{iQuietEnd + 1}.")
                lDrop.update(range(iQuietFrom, iQuietEnd + 1))

    # Part two: the JawsScripts class, wherever it lives.
    lClassHits = [i for i, s in enumerate(lLines) if re.search(r"\bclass\s+JawsScripts\b", s)]
    say(f"JawsScripts class: {len(lClassHits)} at {[i + 1 for i in lClassHits]}")
    if len(lClassHits) != 1:
        say("FAILED: expected exactly one JawsScripts class. Send this log.")
        return 1
    iClass = lClassHits[0]
    iClassEnd = findBlockEnd(lLines, iClass, {"bBlockComment": False})
    if iClassEnd < 0:
        say("FAILED: the class braces did not balance. Send this log.")
        return 1
    iClassFrom = commentStart(lLines, iClass)
    say(f"Removing the class: lines {iClassFrom + 1}-{iClassEnd + 1} "
        f"({iClassEnd - iClassFrom + 1} lines, comments included).")
    lDrop.update(range(iClassFrom, iClassEnd + 1))

    lNew = [s for i, s in enumerate(lLines) if i not in lDrop]

    # The whole file's braces must balance exactly as before minus nothing:
    # every removed region balanced itself, so the total must still balance.
    dState = {"bBlockComment": False}
    iBalance = 0
    for sLine in lNew:
        for sChar in stripCodeOnly(sLine, dState):
            if sChar == "{":
                iBalance += 1
            elif sChar == "}":
                iBalance -= 1
    say(f"Brace balance after removal: {iBalance} (0 is correct).")
    if iBalance != 0:
        say("FAILED: the result does not balance; nothing was written. Send this log.")
        return 1
    if any("JawsScripts" in s or "--install-jaws-settings" in s for s in lNew):
        say("FAILED: a reference survived the removal; nothing was written. Send this log.")
        return 1

    pathBackup = os.path.join(pathRoot, c_sBackupName)
    with open(pathBackup, "wb") as fileBackup:
        fileBackup.write(binData)
    say(f"Backup written: {pathBackup}")
    sNew = sBreak.join(lNew)
    with open(pathSource, "wb") as fileSource:
        if bBom:
            fileSource.write(b"\xef\xbb\xbf")
        fileSource.write(sNew.encode("utf-8"))
    say(f"Wrote {c_sSourceName}: {os.path.getsize(pathSource)} bytes, "
        f"{len(lLines) - len(lNew)} lines removed.")
    say("")
    say("Done. Run buildEdSharp; a clean compile is the proof the surgery was clean.")
    return 0


if __name__ == "__main__":
    iExit = 1
    try:
        iExit = main()
    except Exception:
        import traceback
        say("FAILED with an unexpected error:")
        say(traceback.format_exc())
    finally:
        if fileLog:
            fileLog.close()
    sys.exit(iExit)
