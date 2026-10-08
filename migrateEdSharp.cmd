@echo off
rem migrateEdSharp.cmd -- runs migrateEdSharp.py with the system Python.
rem Arguments are passed straight through; --survey prints the plan and
rem changes nothing. Run this once, after unzipping the new EdSharp.zip.
python "%~dp0migrateEdSharp.py" %*
exit /b %errorlevel%
