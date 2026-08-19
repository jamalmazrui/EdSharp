@echo off
rem quietJaws.cmd -- runs quietJaws.py with the system Python. No arguments needed.
python "%~dp0quietJaws.py" %*
exit /b %errorlevel%
