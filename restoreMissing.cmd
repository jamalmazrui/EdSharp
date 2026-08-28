@echo off
rem restoreMissing.cmd -- runs restoreMissing.py with the system Python.
rem Arguments are passed straight through; none are needed for a restore.
python "%~dp0restoreMissing.py" %*
exit /b %errorlevel%
