@echo off
rem dropLatexJawsKeys.cmd -- runs the Python script of the same name.
python "%~dp0dropLatexJawsKeys.py" %*
if not %errorlevel% == 0 pause
