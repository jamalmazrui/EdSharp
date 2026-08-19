r"""quietJaws.py -- teach EdSharp.exe a --quiet argument for the JAWS step.

Run it from C:\EdSharp with no arguments:

    python quietJaws.py

WHAT AND WHY

The installer's Finish page runs "EdSharp.exe --install-jaws-settings", and
that routine ends by showing its own report box ("EdSharp JAWS scripts: 72
copied, 18 compiled").  That box was the only feedback when the command was
run by hand; now that the installer's closing Results box reports the JAWS
outcome, hearing two boxes in a row is one box too many.

This script edits EdSharp.cs so that:

  1. The routine writes its full report to the EdSharp logs folder
     (%LOCALAPPDATA%\EdSharp\logs\installJawsScripts.log) EVERY time --
     quiet or not -- so the details are always on disk with the other logs.
  2. A new --quiet argument (also /quiet) skips the report box.  The
     installer passes it; a person running the command by hand does not,
     and still hears the box.

The edit is anchored on unique lines rather than exact whitespace, applied
once, and is safe to re-run: a second run reports "already applied" and
changes nothing.  A backup copy EdSharp.cs.bak is written first.  After
this, run buildEdSharp to compile the change into EdSharp.exe.

A detailed log is written beside this script, whatever happens.
"""

import datetime
import os
import re
import sys

c_sBackupName = "EdSharp.cs.bak"
c_sLogName = "quietJaws.log"
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


def main():
    global fileLog
    fileLog = open(pathLog, "w", encoding="utf-8")
    say(f"quietJaws  {datetime.datetime.now():%Y-%m-%d %H:%M:%S}")
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
    say(f"Read {c_sSourceName}: {len(binData)} bytes, {len(lLines)} lines, "
        f"{'with' if bBom else 'without'} BOM, {'CRLF' if sBreak == chr(13) + chr(10) else 'LF'} line breaks.")

    if any("bQuiet" in sLine for sLine in lLines):
        say("Already applied: the file mentions bQuiet. Nothing to do.")
        say("If the box still appears, run buildEdSharp so the change is compiled in.")
        return 0

    # Anchor one: the report box line, unique by its title text.
    lBoxHits = [i for i, s in enumerate(lLines)
                if "MessageBox.Show(sReport" in s and "EdSharp JAWS scripts" in s]
    say(f"Report-box line: {len(lBoxHits)} match{'es' if len(lBoxHits) != 1 else ''} "
        f"{[i + 1 for i in lBoxHits]}")
    if len(lBoxHits) != 1:
        say("FAILED: expected exactly one report-box line. The file differs from what was reviewed; send this log.")
        return 1
    iBox = lBoxHits[0]

    # Anchor two: the enclosing argument loop, the nearest one above the box.
    lLoopHits = [i for i, s in enumerate(lLines)
                 if i < iBox and "foreach" in s and "cmdLineArgs" in s and "string sArg" in s]
    say(f"Argument-loop line above it: {len(lLoopHits)} match{'es' if len(lLoopHits) != 1 else ''} "
        f"{[i + 1 for i in lLoopHits]}")
    if not lLoopHits:
        say("FAILED: no argument loop found above the report box; send this log.")
        return 1
    iLoop = lLoopHits[-1]

    sBoxIndent = re.match(r"\s*", lLines[iBox]).group(0)
    sLoopIndent = re.match(r"\s*", lLines[iLoop]).group(0)

    lQuiet = [
        sLoopIndent + "// --quiet (or /quiet) skips the report box below. The installer passes",
        sLoopIndent + "// it, because its own closing Results box reports the JAWS outcome and two",
        sLoopIndent + "// boxes in a row is one too many; a person running the command by hand",
        sLoopIndent + "// still hears the box.",
        sLoopIndent + "bool bQuiet = false;",
        sLoopIndent + "foreach (string sQuietArg in cmdLineArgs) {",
        sLoopIndent + "if (sQuietArg.Equals(\"--quiet\", StringComparison.OrdinalIgnoreCase)",
        sLoopIndent + "|| sQuietArg.Equals(\"/quiet\", StringComparison.OrdinalIgnoreCase)) bQuiet = true;",
        sLoopIndent + "}",
    ]
    lReport = [
        sBoxIndent + "string sTitle = \"EdSharp JAWS scripts: \" + iCopied + \" copied, \" + iCompiled + \" compiled\";",
        sBoxIndent + "// The full report always lands in the EdSharp logs folder, beside the",
        sBoxIndent + "// installer's and installPandoc's logs, so the details survive whether or",
        sBoxIndent + "// not anyone saw a box.",
        sBoxIndent + "try {",
        sBoxIndent + "string sLogDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), \"EdSharp\");",
        sBoxIndent + "sLogDir = Path.Combine(sLogDir, \"logs\");",
        sBoxIndent + "Directory.CreateDirectory(sLogDir);",
        sBoxIndent + "System.IO.File.WriteAllText(Path.Combine(sLogDir, \"installJawsScripts.log\"), sTitle + \"\\r\\n\\r\\n\" + sReport);",
        sBoxIndent + "}",
        sBoxIndent + "catch {}",
        sBoxIndent + "if (!bQuiet) MessageBox.Show(sReport, sTitle);",
    ]

    lNew = lLines[:iLoop] + lQuiet + lLines[iLoop:iBox] + lReport + lLines[iBox + 1:]
    say("")
    say(f"Inserting the quiet-argument check before line {iLoop + 1} and replacing line {iBox + 1} with the logged, conditional report.")

    pathBackup = os.path.join(pathRoot, c_sBackupName)
    with open(pathBackup, "wb") as fileBackup:
        fileBackup.write(binData)
    say(f"Backup written: {pathBackup}")

    sNew = sBreak.join(lNew)
    with open(pathSource, "wb") as fileSource:
        if bBom:
            fileSource.write(b"\xef\xbb\xbf")
        fileSource.write(sNew.encode("utf-8"))
    say(f"Wrote {c_sSourceName}: {os.path.getsize(pathSource)} bytes.")
    say("")
    say("Done. Run buildEdSharp to compile the change into EdSharp.exe;")
    say("the build will also rebuild the installer with the --quiet parameter.")
    return 0


if __name__ == "__main__":
    iExit = 1
    try:
        iExit = main()
    except Exception as exception:
        import traceback
        say("FAILED with an unexpected error:")
        say(traceback.format_exc())
    finally:
        if fileLog:
            fileLog.close()
    sys.exit(iExit)
