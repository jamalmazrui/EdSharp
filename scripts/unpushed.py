#!/usr/bin/env python3
"""unpushed.py -- undo the commits that have not been pushed, keeping every file.

WHEN IT IS FOR. A tidy or a hand commit has swept something into the
repository that should never go there -- fetched voices, a model, a build's
output -- and the push was stopped, or has not happened yet. The commit is
only local. This puts the branch back to what the remote has, leaves every
file on disk exactly as it is, and unstages everything, so the next commit
can be made properly (tidy, with RepoFiles.txt in place, makes it).

WHAT IT REFUSES. If every local commit is already on the remote there is
nothing to undo, and it says so. It never rewrites what has been pushed.

Run it in the project folder. It logs to logs\\<App>-unpushed-yyyyMMdd-HHmmss.log.
"""

import datetime, os, platform, subprocess, sys, time

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
    """One event in the Homer log format (1.43.21), as log.py and Log.cs write:
    an ISO 8601 time with milliseconds and UTC offset, a five-character level,
    then the text; a line continuing the one above starts "| ", and no line is
    blank or unstamped. The level is ERROR or WARN when the text says so."""
    if oLog is None: return True
    import datetime as _datetime
    sText = (sText or "").replace("\r\n", "\n").rstrip("\n")
    if not sText.strip(): return True
    import re as _re
    sLevel = ("ERROR" if _re.search(r"\b(ERROR|FAIL|FAILED)\b", sText)
              else "WARN" if _re.search(r"\bWARN(ING)?\b", sText) else "INFO")
    # A LEADING LEVEL WORD IS THE LEVEL (1.43.33): "WARN: x" is written
    # "WARN  x", not "WARN  WARN: x".
    oLead = _re.match(r"(ERROR|WARN|WARNING)\b:?\s*", sText)
    if oLead: sText = sText[oLead.end():] or sText
    sPrefix = "%s %-5s " % (_datetime.datetime.now().astimezone().isoformat(timespec="milliseconds"), sLevel)
    lsOut = []
    for iAt, sOne in enumerate(sText.split("\n")):
        if iAt and not sOne.strip(): continue
        lsOut.append(sPrefix + ("| " if iAt else "") + sOne.rstrip())
    oLog.write("\n".join(lsOut) + "\n")
    oLog.flush()
    return True

def logValue(sValue):
    """A value as the Homer log format writes it: bare when it can be, quoted
    when it holds a space, a quote or an equals sign."""
    import re as _re
    s = "" if sValue is None else str(sValue)
    if s and not _re.search(r'[\s"=]', s): return s
    s = s.replace('"', '\\"')
    if s.endswith("\\"): s += "\\"
    return '"' + s + '"'


def logFact(sKey, sValue):
    """One environment fact: env key=value."""
    return logLine("env %s=%s" % (sKey, logValue(sValue)))

def logWindows():
    """The Windows actually running, worded as Log.cs and log.py word it:
    "Windows 11 25H2 (10.0.26200.9550)"."""
    try:
        import winreg as _winreg
        with _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion") as oKey:
            def read(sName):
                try: return str(_winreg.QueryValueEx(oKey, sName)[0])
                except OSError: return ""
            sBuild, sUbr, sDisplay = read("CurrentBuild"), read("UBR"), read("DisplayVersion")
        sName = "Windows 11" if sBuild.isdigit() and int(sBuild) >= 22000 else "Windows 10"
        return ("%s %s" % (sName, sDisplay)).strip() + " (10.0.%s%s)" % (sBuild, "." + sUbr if sUbr else "")
    except Exception:
        import platform as _platform
        return _platform.platform()




def say(sText):
    print(sText)
    logLine("CONSOLE: " + sText)
    return True


def runGit(lsArgs):
    sCmd = "git " + " ".join(lsArgs)
    nStarted = time.time()
    logLine("run start cmd=" + logValue(sCmd))
    oResult = subprocess.run(["git"] + lsArgs, cwd=sRoot, capture_output=True, text=True)
    logLine("run exit=%d ms=%d cmd=%s" % (oResult.returncode, (time.time() - nStarted) * 1000, logValue(sCmd)))
    if oResult.stdout.strip(): logLine("STDOUT:\n" + oResult.stdout.rstrip())
    if oResult.stderr.strip(): logLine("STDERR:\n" + oResult.stderr.rstrip())
    return oResult.returncode, oResult.stdout.strip()


def main():
    global oLog
    sLogDir = os.path.join(sRoot, "logs")
    os.makedirs(sLogDir, exist_ok=True)
    sLogPath = os.path.join(sLogDir, "%s-unpushed-%s.log" % (os.path.basename(sRoot), datetime.datetime.now().strftime("%Y%m%d-%H%M%S")))
    oLog = open(sLogPath, "w", encoding="utf-8")
    logLine("unpushed start pid=%d" % os.getpid())
    logFact("script", os.path.abspath(__file__))
    logFact("python", platform.python_version())
    logFact("windows", logWindows())
    logLine("Working directory: %s" % sRoot)
    logFact("arguments", " ".join(sys.argv[1:]))
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
    say("Run tidy to make the commit properly; it needs RepoFiles.txt.")
    say("Log: " + sLogPath)
    logLine("unpushed end")
    return 0


if __name__ == "__main__":
    sys.exit(main())
