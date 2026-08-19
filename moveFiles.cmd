@echo off
rem moveFiles.cmd - move obsolete and superseded files of the EdSharp
rem and FileDir upgrade to the C:\DbDuo archive, replacing anything
rem already archived under the same name. Part of the cleanup and
rem pruning stage that follows the 64-bit milestone. Safe to run more
rem than once: anything already moved or absent is skipped.
rem
rem What moves, and why it is obsolete:
rem   Jaws.cs (all three apps)   merged into the shared Say.cs
rem   PruneConvert.cmd           superseded by this archiving script
rem   Convert\GetText            retired extractor; hung on modern Windows
rem   Convert\NFBTrans           replaced by liblouis back-translation
rem   Convert\HTM2TXT            replaced by Pandoc, 2htm, and in-app code
rem   Convert\EasyEncode         encoding is handled natively
rem   WdVert.exe, PpVert.exe     Office COM extraction replaced by 2htm;
rem                              XlVert.exe stays as the CSV producer
rem   Xpdf extras                only pdftotext.exe is referenced
rem   AStyle sources and builds  astyle.exe, its option files, and the
rem                              license stay
rem   doc2txt.bat, doc2txt.cmd   legacy wrappers; any2txt.cmd is current

setlocal
set "sArchive=C:\DbDuo"
set "iMoved=0"

rem ---- EdSharp ----
set "sApp=C:\EdSharp"
set "sDest=%sArchive%\EdSharp"
if not exist "%sApp%" (echo Skipping EdSharp: %sApp% not found.& goto fileDir)
md "%sDest%" 2>nul
md "%sDest%\Convert" 2>nul
md "%sDest%\Convert\OfficeConvert" 2>nul
md "%sDest%\Convert\Xpdf" 2>nul
md "%sDest%\Convert\AStyle" 2>nul
if exist "%sApp%\Jaws.cs" (if exist "%sDest%\Jaws.cs" del /q "%sDest%\Jaws.cs") & if exist "%sApp%\Jaws.cs" move /y "%sApp%\Jaws.cs" "%sDest%\" >nul && (echo moved EdSharp\Jaws.cs& set /a iMoved+=1)
if exist "%sApp%\PruneConvert.cmd" move /y "%sApp%\PruneConvert.cmd" "%sDest%\" >nul && (echo moved EdSharp\PruneConvert.cmd& set /a iMoved+=1)
for %%d in (GetText NFBTrans HTM2TXT EasyEncode) do (
  if exist "%sApp%\Convert\%%d" (
    if exist "%sDest%\Convert\%%d" rmdir /s /q "%sDest%\Convert\%%d"
    move /y "%sApp%\Convert\%%d" "%sDest%\Convert\%%d" >nul && (echo moved EdSharp\Convert\%%d& set /a iMoved+=1)
  )
)
for %%f in (WdVert.exe PpVert.exe) do (
  if exist "%sApp%\Convert\OfficeConvert\%%f" move /y "%sApp%\Convert\OfficeConvert\%%f" "%sDest%\Convert\OfficeConvert\" >nul && (echo moved EdSharp\Convert\OfficeConvert\%%f& set /a iMoved+=1)
)
for %%f in (pdfdetach.exe pdffonts.exe pdfimages.exe pdfinfo.exe pdftohtml.exe pdftopng.exe pdftoppm.exe pdftops.exe) do (
  if exist "%sApp%\Convert\Xpdf\%%f" move /y "%sApp%\Convert\Xpdf\%%f" "%sDest%\Convert\Xpdf\" >nul && (echo moved EdSharp\Convert\Xpdf\%%f& set /a iMoved+=1)
)
for %%d in (build src doc man sh-completion) do (
  if exist "%sApp%\Convert\AStyle\%%d" (
    if exist "%sDest%\Convert\AStyle\%%d" rmdir /s /q "%sDest%\Convert\AStyle\%%d"
    move /y "%sApp%\Convert\AStyle\%%d" "%sDest%\Convert\AStyle\%%d" >nul && (echo moved EdSharp\Convert\AStyle\%%d& set /a iMoved+=1)
  )
)
for %%f in (CMakeLists.txt README.md) do (
  if exist "%sApp%\Convert\AStyle\%%f" move /y "%sApp%\Convert\AStyle\%%f" "%sDest%\Convert\AStyle\" >nul && (echo moved EdSharp\Convert\AStyle\%%f& set /a iMoved+=1)
)
for %%f in (doc2txt.bat doc2txt.cmd) do (
  if exist "%sApp%\Convert\%%f" move /y "%sApp%\Convert\%%f" "%sDest%\Convert\" >nul && (echo moved EdSharp\Convert\%%f& set /a iMoved+=1)
)

:fileDir
rem ---- FileDir ----
set "sApp=C:\FileDir"
set "sDest=%sArchive%\FileDir"
if not exist "%sApp%" (echo Skipping FileDir: %sApp% not found.& goto dbDo)
md "%sDest%" 2>nul
if exist "%sApp%\Jaws.cs" move /y "%sApp%\Jaws.cs" "%sDest%\" >nul && (echo moved FileDir\Jaws.cs& set /a iMoved+=1)

:dbDo
rem ---- DbDo (same Jaws.cs merge applies) ----
set "sApp=C:\DbDo"
set "sDest=%sArchive%\DbDo"
if not exist "%sApp%" (echo Skipping DbDo: %sApp% not found.& goto done)
md "%sDest%" 2>nul
if exist "%sApp%\Jaws.cs" move /y "%sApp%\Jaws.cs" "%sDest%\" >nul && (echo moved DbDo\Jaws.cs& set /a iMoved+=1)

:done
echo.
echo Archived %iMoved% item(s) to %sArchive%.
endlocal
