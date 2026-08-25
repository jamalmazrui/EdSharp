r"""pdfRich.py -- convert a PDF to RICH Markdown for EdSharp's Import.

    python pdfRich.py "source.pdf" "target.md"

WHY THIS AND NOT A TEXT EXTRACTOR

Plain text from a PDF loses everything a screen reader user navigates
by: headings, lists, tables, emphasis, and reading order. This helper
uses PyMuPDF4LLM, which reads a PDF's own structure -- font sizes become
heading levels, bullet runs become lists, ruled areas become Markdown
tables -- and writes Markdown. EdSharp then turns that Markdown into
HTML or a Word document in the binary or with Pandoc, so one rich
conversion serves every target.

Nothing here needs Microsoft Word. Install the requirement once with
installPdfTools.cmd in the EdSharp folder.

A detailed log is written beside the target as <target>.log whatever
happens, so EdSharp's error dialog can quote the reason.
"""

import datetime
import os
import sys

c_iExitBadArguments = 2
c_iExitNoLibrary = 3
c_iExitFailed = 4
c_iExitNoOutput = 5

lLogLines = []


def say(sMessage=""):
    """Collect a line for the log beside the target."""
    lLogLines.append(sMessage)
    print(sMessage)


def writeLog(pathTarget):
    try:
        with open(pathTarget + ".log", "w", encoding="utf-8") as fileLog:
            fileLog.write("\n".join(lLogLines) + "\n")
    except Exception:
        pass


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        return c_iExitBadArguments
    pathSource = os.path.abspath(sys.argv[1])
    pathTarget = os.path.abspath(sys.argv[2])

    say("pdfRich started " + datetime.datetime.now().isoformat(" ", "seconds"))
    say("Script: " + os.path.abspath(__file__))
    say("Python: " + sys.version.split()[0] + " (" + ("64-bit" if sys.maxsize > 2**32 else "32-bit") + ")")
    say("Source: " + pathSource)
    say("Target: " + pathTarget)

    if not os.path.isfile(pathSource):
        say("FAILED: the source file was not found.")
        writeLog(pathTarget)
        return c_iExitBadArguments

    try:
        import pymupdf4llm
    except ImportError as oError:
        say("FAILED: the PDF reader is not installed (" + str(oError) + ").")
        say("Run installPdfTools.cmd in the EdSharp program folder to install it.")
        say("It needs Python, which the EdSharp installer offers as a checkbox.")
        writeLog(pathTarget)
        return c_iExitNoLibrary

    try:
        say("Reading the PDF and building Markdown ...")
        sMarkdown = pymupdf4llm.to_markdown(pathSource)
    except Exception as oError:
        say("FAILED: " + repr(oError))
        writeLog(pathTarget)
        return c_iExitFailed

    if not sMarkdown or not sMarkdown.strip():
        say("FAILED: the PDF produced no text. It may be a scan of images,")
        say("which needs optical character recognition rather than conversion.")
        writeLog(pathTarget)
        return c_iExitNoOutput

    try:
        if os.path.exists(pathTarget):
            os.remove(pathTarget)
        with open(pathTarget, "w", encoding="utf-8") as fileTarget:
            fileTarget.write(sMarkdown)
    except Exception as oError:
        say("FAILED writing the target: " + repr(oError))
        writeLog(pathTarget)
        return c_iExitFailed

    iHeadings = sum(1 for sLine in sMarkdown.split("\n") if sLine.startswith("#"))
    iTables = sMarkdown.count("\n|")
    say("Done. " + str(len(sMarkdown)) + " characters, "
        + str(iHeadings) + (" heading" if iHeadings == 1 else " headings") + ", "
        + str(iTables) + (" table row" if iTables == 1 else " table rows") + ".")
    # On success the log is not kept: EdSharp reads it only to explain a
    # failure, and a stale log beside a good conversion is noise.
    try:
        if os.path.exists(pathTarget + ".log"):
            os.remove(pathTarget + ".log")
    except Exception:
        pass
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as oError:
        say("FAILED: " + repr(oError))
        try:
            writeLog(os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else "pdfRich")
        except Exception:
            pass
        sys.exit(c_iExitFailed)
