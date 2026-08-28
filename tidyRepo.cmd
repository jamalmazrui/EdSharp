@echo off
rem tidyRepo.cmd -- runs tidyRepo.py with the system Python. Arguments are
rem passed straight through; none are needed for the normal tidy run.
python "%~dp0tidyRepo.py" %*
exit /b %errorlevel%
