@echo off
rem removeJawsFeature.cmd -- runs removeJawsFeature.py with the system Python.
python "%~dp0removeJawsFeature.py" %*
exit /b %errorlevel%
