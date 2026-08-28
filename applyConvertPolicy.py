r"""applyConvertPolicy.py -- finish the conversion-policy change on disk.

Run it from C:\EdSharp with no arguments:

    python applyConvertPolicy.py

The revised EdSharp.inix and the new Convert .cmd files arrive by
unarchiving EdSharp.zip; this script does the two things an unarchive
cannot:

  1. DELETES the retired batch files. The policy is .cmd, never .bat, and
     doc2txt is orphaned entirely (nothing references it; the Import table
     uses any2txt.cmd). Removed if present:
       Convert\doc2txt.bat    Convert\doc2txt.cmd
       Convert\pbw.bat        Convert\Mingw.bat
     (pbw.cmd and MinGW.cmd stay; the Compilers overrides in EdSharp.inix
     now call them.)

  2. STRIPS the [Import] and [Export] sections from the SHIPPED
     EdSharp.ini beside this script, so EdSharp.inix alone carries the
     conversion tables, as intended. A backup EdSharp.ini.convertpolicy.bak
     is written first. Your personal EdSharp.ini in the data folder is not
     touched; the tombstones in EdSharp.inix hide its stale entries.

Safe to run twice: missing files and already-stripped sections are
reported, not errors. Everything is logged beside this script.
"""

import datetime
import os
import re
import sys

c_sLogName = "applyConvertPolicy.log"

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


def main():
    global fileLog
    fileLog = open(pathLog, "w", encoding="utf-8")
    say(f"applyConvertPolicy  {datetime.datetime.now():%Y-%m-%d %H:%M:%S}")
    say(f"  script:            {os.path.abspath(__file__)}")
    say(f"  Python:            {sys.version.split()[0]}")
    say(f"  working directory: {os.getcwd()}")
    say("")

    iRemoved = 0
    for sName in ("doc2txt.bat", "doc2txt.cmd", "pbw.bat", "Mingw.bat"):
        sPath = os.path.join(pathRoot, "Convert", sName)
        if os.path.isfile(sPath):
            os.remove(sPath)
            say(f"Deleted Convert\\{sName}")
            iRemoved += 1
        else:
            say(f"Convert\\{sName} is already gone.")
    sFiles = "file" if iRemoved == 1 else "files"
    say(f"{iRemoved} retired {sFiles} deleted.")
    say("")

    pathIni = os.path.join(pathRoot, "EdSharp.ini")
    if not os.path.isfile(pathIni):
        say("EdSharp.ini is not beside this script; nothing to strip.")
        say(f"Done. The log is at {pathLog}")
        return 0
    with open(pathIni, "rb") as fileIni:
        binData = fileIni.read()
    bBom = binData.startswith(b"\xef\xbb\xbf")
    sText = binData.decode("utf-8-sig", errors="replace")
    sBreak = "\r\n" if "\r\n" in sText else "\n"
    lLines = sText.replace("\r\n", "\n").split("\n")
    lOut = []
    sSection = ""
    iDropped = 0
    for sLine in lLines:
        matchHeader = re.match(r"^\[(.+)\]\s*$", sLine)
        if matchHeader:
            sSection = matchHeader.group(1).strip().lower()
        if sSection in ("import", "export"):
            iDropped += 1
            continue
        lOut.append(sLine)
    if iDropped == 0:
        say("EdSharp.ini already has no Import or Export section; nothing to strip.")
    else:
        pathBackup = pathIni + ".convertpolicy.bak"
        with open(pathBackup, "wb") as fileBackup:
            fileBackup.write(binData)
        say(f"Backup written: {pathBackup}")
        with open(pathIni, "wb") as fileIni:
            if bBom:
                fileIni.write(b"\xef\xbb\xbf")
            fileIni.write(sBreak.join(lOut).encode("utf-8"))
        sLines = "line" if iDropped == 1 else "lines"
        say(f"Stripped the Import and Export sections from EdSharp.ini ({iDropped} {sLines} removed).")
        say("EdSharp.inix alone now carries the conversion tables.")
    say("")
    say(f"Done. The log is at {pathLog}")
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
