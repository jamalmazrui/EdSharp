@echo off
rem applyConvertPolicy.cmd -- runs applyConvertPolicy.py with the system Python.
python "%~dp0applyConvertPolicy.py" %*
exit /b %errorlevel%
