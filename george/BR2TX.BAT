@echo off
echo NFB Trans status file > stat.txt
echo --- >> stat.txt
echo Grade Two Braille To Text >> stat.txt
if not "%1" == "" goto chkipt
echo no input file specified >> stat.txt
winout no input file specified
goto end
:chkipt
if exist %1 goto chkopt
echo the input file does not exist >> stat.txt
winout the input file does not exist
goto end
:chkopt
if not "%2" == "" goto allok
echo no output file specified >> stat.txt
winout no output file specified
goto end
:allok
rem handle long file names
md \btTemp
copy /y %1 \btTemp\temp.brf
nfbtrans.exe tm=13 \btTemp\temp.brf >\btTemp\temp.txt
move /y \btTemp\temp.txt %2
del \btTemp\temp.brf
rd \btTemp
winout %1 successfully translated to %2
:end