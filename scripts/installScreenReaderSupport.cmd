@echo off
rem installScreenReaderSupport.cmd -- put this app's JAWS scripts and NVDA
rem add-on where each screen reader looks for them.
rem
rem SHIPPED WITH EVERY HOMER APP THAT HAS SCRIPTS, and its installer checkbox is
rem CHECKED BY DEFAULT. A blind user installing a Homer tool wants the scripts;
rem making them notice and tick a box is the friction this suite exists to
rem remove.
rem
rem What it looks for, beside this script:
rem   <App>_JAWS.zip    the JAWS script files, unpacked into the user's own
rem                     JAWS settings folder for every JAWS version found
rem   <App>.nvda-addon  the NVDA add-on, handed to NVDA to install
rem
rem Either may be absent; what is here is installed and what is not is skipped
rem without complaint.
rem
rem AI NOTE FOR CUSTOMIZING: nothing here is app-specific. The app name comes
rem from the folder. If an app's scripts need compiling rather than copying, add
rem that step after the unpack and log its exit code.
rem
rem NOTHING PAUSES. LOG: %LOCALAPPDATA%\<App>\logs\<App>_setup.log.
setlocal EnableExtensions EnableDelayedExpansion

rem The app name comes from the folder this script is installed into, climbing
rem one level from exec or scripts. INSTALLERS PUT THIS SCRIPT IN SCRIPTS
rem (1.43.20): with only exec climbed, the app was named "scripts", so FileDir's
rem finish page looked for scripts_JAWS.zip, found nothing, logged to
rem %LOCALAPPDATA%\scripts\logs, and installed no JAWS scripts at all.
for %%d in ("%~dp0.") do set "sApp=%%~nxd"
if /i "%sApp%"=="exec" for %%d in ("%~dp0..") do set "sApp=%%~nxd"
if /i "%sApp%"=="scripts" for %%d in ("%~dp0..") do set "sApp=%%~nxd"
set "sLogDir=%LOCALAPPDATA%\%sApp%\logs"
set "sLog=%sLogDir%\%sApp%_setup.log"
if not exist "%sLogDir%" mkdir "%sLogDir%" >nul 2>&1

call :logLine "installScreenReaderSupport started %DATE% %TIME%"
call :logLine "Script: %~f0"
call :logLine "App: %sApp%"

set "iInstalled=0"

rem ONE READER AT A TIME, WHEN ASKED (1.43.20). An installer gives the JAWS
rem scripts and the NVDA add-on a box each, so each box runs this with "jaws"
rem or "nvda"; with neither, both are installed, as before.
set "bJaws=1"
set "bNvda=1"
for %%a in (%*) do (
  if /i "%%~a"=="jaws" set "bNvda=0"
  if /i "%%~a"=="nvda" set "bJaws=0"
)
call :logLine "JAWS: %bJaws%  NVDA: %bNvda%"
if "%bJaws%"=="0" goto :nvda

rem ---- JAWS ----------------------------------------------------------------
if not exist "%~dp0%sApp%_JAWS.zip" goto :nvda
call :logLine "Found %sApp%_JAWS.zip"
set "sJawsRoot=%APPDATA%\Freedom Scientific\JAWS"
if not exist "%sJawsRoot%" (
  echo JAWS was not found on this computer, so its scripts were skipped.
  call :logLine "No JAWS settings folder at %sJawsRoot%"
  goto :nvda
)
for /d %%v in ("%sJawsRoot%\*") do (
  if exist "%%v\Settings\enu" (
    echo Installing the JAWS scripts for %%~nxv.
    call :logLine "Unpacking into %%v\Settings\enu"
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
      "Expand-Archive -LiteralPath '%~dp0%sApp%_JAWS.zip' -DestinationPath '%%v\Settings\enu' -Force" >> "%sLog%" 2>&1
    call :logLine "Expand-Archive exit code: !ERRORLEVEL!"
    set /a iInstalled=iInstalled+1
  )
)

:nvda
rem ---- NVDA ----------------------------------------------------------------
if "%bNvda%"=="0" goto :done
if not exist "%~dp0%sApp%.nvda-addon" goto :done
call :logLine "Found %sApp%.nvda-addon"
where nvda >nul 2>&1
if errorlevel 1 (
  if exist "%ProgramFiles(x86)%\NVDA\nvda.exe" (
    set "sNvda=%ProgramFiles(x86)%\NVDA\nvda.exe"
  ) else if exist "%ProgramFiles%\NVDA\nvda.exe" (
    set "sNvda=%ProgramFiles%\NVDA\nvda.exe"
  )
) else (
  set "sNvda=nvda"
)
if not defined sNvda (
  echo NVDA was not found on this computer, so its add-on was skipped.
  call :logLine "NVDA not found."
  goto :done
)
echo Installing the NVDA add-on. NVDA will ask you to confirm.
call :logLine "Handing the add-on to NVDA: !sNvda!"
start "" "!sNvda!" --install-add-on="%~dp0%sApp%.nvda-addon"
set /a iInstalled=iInstalled+1

:done
if "%iInstalled%"=="0" (
  echo No screen reader support was installed.
) else (
  echo Screen reader support installed.
)
call :logLine "Screen reader support installed: %iInstalled% item(s)"
call :logLine "installScreenReaderSupport finished %DATE% %TIME%"
endlocal
exit /b 0

:logLine
echo %~1>> "%sLog%"
goto :eof
