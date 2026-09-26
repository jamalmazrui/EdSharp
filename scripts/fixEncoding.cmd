@echo off
rem fixEncoding.cmd -- put the project's own text files into the Homer encoding. Run in the project or its scripts folder.
python "%~dp0fixEncoding.py" %*
