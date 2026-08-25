@echo off
rem auditEdSharp.cmd -- runs the audit checks; see auditEdSharp.py.
python "%~dp0auditEdSharp.py" %*
if not %errorlevel%==0 pause
