@echo off
rem moveNotes.cmd -- runs moveNotes.py with the system Python. Arguments are
rem passed straight through; none are needed for the normal move.
python "%~dp0moveNotes.py" %*
exit /b %errorlevel%
