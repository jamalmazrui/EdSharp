@echo off
rem PruneConvert.cmd - remove Convert-tree tools that the modernized
rem conversion tables no longer reference. Run it from the EdSharp
rem project folder (the folder holding this script and Convert).
rem Safe to run more than once; anything already gone is skipped.
rem What goes, and why:
rem   GetText        retired text extractor; hung on modern Windows
rem   NFBTrans       replaced by liblouis backward translation
rem   HTM2TXT        replaced by Pandoc, 2htm, and in-program code
rem   EasyEncode     encoding is handled natively (YieldEncoding)
rem   WdVert, PpVert Office COM text extraction replaced by 2htm;
rem                  XlVert stays as the only CSV producer
rem   Xpdf extras    only pdftotext.exe is referenced by the tables
rem   AStyle debris  sources, build files, docs; astyle.exe, its
rem                  option files, and LICENSE.md stay
rem   doc2txt.bat/.cmd  legacy wrappers; any2txt.cmd is current
setlocal
cd /d "%~dp0"
if not exist "Convert" echo Convert folder not found beside this script.& exit /b 1
rmdir /s /q "Convert\GetText" 2>nul
rmdir /s /q "Convert\NFBTrans" 2>nul
rmdir /s /q "Convert\HTM2TXT" 2>nul
rmdir /s /q "Convert\EasyEncode" 2>nul
del /q "Convert\OfficeConvert\WdVert.exe" 2>nul
del /q "Convert\OfficeConvert\PpVert.exe" 2>nul
for %%f in (pdfdetach pdffonts pdfimages pdfinfo pdftohtml pdftopng pdftoppm pdftops) do del /q "Convert\Xpdf\%%f.exe" 2>nul
rmdir /s /q "Convert\AStyle\build" 2>nul
rmdir /s /q "Convert\AStyle\src" 2>nul
rmdir /s /q "Convert\AStyle\doc" 2>nul
rmdir /s /q "Convert\AStyle\man" 2>nul
rmdir /s /q "Convert\AStyle\sh-completion" 2>nul
del /q "Convert\AStyle\CMakeLists.txt" 2>nul
del /q "Convert\AStyle\README.md" 2>nul
del /q "Convert\doc2txt.bat" 2>nul
del /q "Convert\doc2txt.cmd" 2>nul
echo Convert tree pruned.
endlocal
