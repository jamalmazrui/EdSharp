@echo off
cls
echo NFBTRANS Installation Program
if not "%1" == "" goto got_param
echo Usage: INSTALL directory
echo Example: INSTALL c:\nfbtrans
goto exit
:got_param
echo This batch file will copy nfbtrans files to %1 and create the file
echo nfbtrans.txt from nfbtrans.fmt with 80 columns per line and 56 lines per
echo page.  nfbtrans.cnf will not be overwritten if it already exists.
pause
if exist %1\nfbtrans.exe goto got_dir
echo creating %1
md %1
:got_dir
rem copy required files
echo copying readme.txt
copy readme.txt %1 >nul
echo copying nfbtrans.exe
copy nfbtrans.exe %1 >nul
echo copying braille.tab
copy braille.tab %1 >nul
echo copying back.tab
copy back.tab %1 >nul
echo copying makedoc.bat
copy makedoc.bat %1 >nul
echo copying nfbtrans.fmt
copy nfbtrans.fmt %1 >nul
echo copying english.dic
copy english.dic %1 >nul
echo copying spell.dat
copy spell.dat %1 >nul
rem skip nfbtrans.cnf if it already exists.
if exist %1\nfbtrans.cnf goto got_cnf
echo copying nfbtrans.cnf
copy nfbtrans.cnf %1 >nul
:got_cnf
echo Creating %1\nfbtrans.txt Please wait...
if exist nfbtrans.fmt goto got_file
echo nfbtrans.fmt not found...
goto exit
:got_file

rem Create nfbprint.fmt identical to nfbtrans.fmt except for the line of
rem formatting commands at the beginning.
echo ~0 ~-pw=80 ~-ls=99 ~-pl=56 ~-if=12 ~-fc=. ~o5 ~p- ~-ip=0 ~-oc=1 >%1\temp.$$$
copy %1\temp.$$$ + nfbtrans.fmt %1\nfbprint.fmt >nul
del %1\temp.$$$ >nul
nfbtrans qm=2 ob=oc=15 %1\nfbprint.fmt >%1\nfbtrans.txt
del %1\nfbprint.fmt >nul
cd %1 >nul

:exit
