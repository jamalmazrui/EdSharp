@echo off
rem brl2txt.bat - back-translate contracted braille to text through
rem liblouis. Called by the EdSharp [Import] entries brl2txt and
rem brf2txt, replacing the retired NFBTrans back-translator.
rem Arguments:
rem   %1  the Convert\liblouis folder (short path, no quotes)
rem   %2  the braille source file (quoted long path)
rem   %3  the text target file (short path)
rem The display table en-us-brf.dis maps North American ASCII braille
rem to dot patterns; en-us-g2.ctb then back-translates contracted
rem (grade 2) English. Change the tables here for another language;
rem the full set is in share\liblouis\tables.
setlocal
set "sDir=%~1"
"%sDir%\bin\lou_translate.exe" --backward "%sDir%\share\liblouis\tables\en-us-brf.dis,%sDir%\share\liblouis\tables\en-us-g2.ctb" < %2 > %3
endlocal
