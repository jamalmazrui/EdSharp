@echo off
rem installPython.cmd -- part of EdSharp setup, Homer Tools pattern: probe first,
rem update when present, install when absent, pause on failure so the
rem reason is logged. NOTHING PAUSES: a console waiting for a keypress
rem interrupts the installation, and the summary shown at the very end --
rem after every checkbox has run -- is where the outcome is reported. The console says only what is happening in a few plain
rem words; every detail goes to the consolidated log.
rem 64-bit by rule: every winget call asks for the x64 build, and where a
rem package offers a machine-wide install it is taken, so components land in
rem the default Windows places -- Program Files, and the PATH every program
rem inherits -- rather than in a per-user corner EdSharp would have to hunt
rem for. The python.org installer defaults to a per-user install; --scope
rem machine asks for the all-users install under Program Files instead, so
rem python.exe sits on the machine PATH where Compile finds it.
rem
rem THE OFFICIAL PYTHON, FROM PYTHON.ORG, NOT THE MICROSOFT STORE. The
rem winget package asked for below, Python.Python.3.13, IS the python.org
rem installer; the Store edition is a different package entirely and is
rem never requested here. The Store edition installs into a sandboxed
rem folder, keeps its own copy of site-packages, and refuses some ordinary
rem operations, which is exactly the sort of surprise a screen reader user
rem should not have to debug. The check below also REJECTS the Microsoft
rem "app execution alias" -- the stub Windows puts on the path at
rem %LOCALAPPDATA%\Microsoft\WindowsApps\python.exe, which answers `where
rem python` and then advertises the Store instead of running anything.
setlocal
set "logFile=%LOCALAPPDATA%\EdSharp\logs\EdSharp_setup.log"
if not exist "%LOCALAPPDATA%\EdSharp\logs" mkdir "%LOCALAPPDATA%\EdSharp\logs" >nul 2>&1
echo [installPython.cmd] started %date% %time% >> "%logFile%"
echo If Windows asks permission during an install or update, a User Account
echo Control prompt appears on a separate screen; press Alt+Y to allow it.
echo A large download can also run quietly for several minutes.
echo.

call :findPython
if defined pythonExe goto upgrade_python
where python >nul 2>&1
if not errorlevel 1 (
  echo The python on this computer is only a Microsoft Store stub.
  echo Installing the official Python ...
  echo [installPython.cmd] Store alias found on PATH; installing python.org build >> "%logFile%"
)

echo Installing Python ...
echo [installPython.cmd] winget install Python.Python.3.13 >> "%logFile%"
winget install --id Python.Python.3.13 -e --architecture x64 --scope machine --silent --disable-interactivity --accept-package-agreements --accept-source-agreements >> "%logFile%" 2>&1
echo [installPython.cmd] winget install exit %errorlevel% >> "%logFile%"
call :findPython
if not defined pythonExe goto fail_python
echo Installed at %pythonExe%
echo [installPython.cmd] installed at %pythonExe% >> "%logFile%"
goto done_python

:upgrade_python
echo Updating Python ...
echo [installPython.cmd] winget upgrade Python.Python.3.13 >> "%logFile%"
winget upgrade --id Python.Python.3.13 -e --architecture x64 --scope machine --silent --disable-interactivity --accept-package-agreements --accept-source-agreements >> "%logFile%" 2>&1
echo [installPython.cmd] winget upgrade exit %errorlevel% >> "%logFile%"
if errorlevel 1 (echo Already current.) else (echo Updated.)
goto done_python

:fail_python
echo The official Python could not be installed.
echo.
echo If the Microsoft Store version is installed, it is best removed: it
echo behaves differently from the python.org build in ways that break
echo tools. Open Settings, Apps, Installed apps, remove Python from the
echo Microsoft Store, then run this script again. You may also turn off the
echo stub that pretends to be Python: Settings, Apps, Advanced app
echo settings, App execution aliases, then switch off both Python entries.
echo.
echo The log is:
echo %logFile%
echo [installPython.cmd] FAILED: no real Python after install >> "%logFile%"
exit /b 3

:done_python

echo Done.
echo [installPython.cmd] done >> "%logFile%"
exit /b 0

:findPython
rem Sets pythonExe to a REAL Python, or leaves it empty.
rem Windows ships an "app execution alias" at
rem %LOCALAPPDATA%\Microsoft\WindowsApps\python.exe which answers `where
rem python` and then prints an advertisement for the Microsoft Store. It is
rem not Python, and treating it as Python is what made this script report
rem success while installing nothing. So the alias path is rejected, and
rem whatever remains must actually answer --version.
set "pythonExe="
for /f "delims=" %%p in ('where python 2^>nul') do (
  echo %%p | find /i "\WindowsApps\" >nul
  if errorlevel 1 (
    if not defined pythonExe set "pythonExe=%%p"
  )
)
if not defined pythonExe (
  for %%d in ("%LOCALAPPDATA%\Programs\Python" "%ProgramFiles%" "%SystemDrive%\") do (
    if exist "%%~d\python.exe" if not defined pythonExe set "pythonExe=%%~d\python.exe"
    for /d %%s in ("%%~d\Python3*") do (
      if exist "%%~s\python.exe" if not defined pythonExe set "pythonExe=%%~s\python.exe"
    )
  )
)
if defined pythonExe (
  "%pythonExe%" --version >nul 2>&1
  if errorlevel 1 set "pythonExe="
)
exit /b 0
