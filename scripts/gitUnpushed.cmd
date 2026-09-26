@echo off
rem gitUnpushed.cmd -- undo the commits not yet pushed, keeping every file. Run in the project folder.
python "%~dp0gitUnpushed.py" %*
