@echo off
rem installOllama.cmd -- smart install of Ollama plus a chat model for
rem EdSharp's Chat with AI command. Probe first: an existing Ollama is
rem updated in place, never duplicated. Tool output stays in this
rem window so progress is readable; the log records milestones.
rem 64-bit by rule: the winget calls ask for the x64 build. Ollama installs
rem per user by its own design, into %LOCALAPPDATA%\Programs\Ollama with its
rem models under the profile -- that IS its default Windows location, and one
rem installation serves every program on the machine through its local
rem service, so it is left exactly there.
rem NOTHING PAUSES: a console waiting for a keypress interrupts the
rem installation. Failures are logged, and the summary shown at the very
rem end reports the outcome of every checkbox.
setlocal
set "logFile=%LOCALAPPDATA%\EdSharp\logs\EdSharp_setup.log"
if not exist "%LOCALAPPDATA%\EdSharp\logs" mkdir "%LOCALAPPDATA%\EdSharp\logs" >nul 2>&1
set "modelName=llama3.2"
echo [installOllama] started %date% %time% >> "%logFile%"
echo.

rem A just-installed Ollama is not on this console's PATH yet, and Scott's
rem test showed the gap: where missed it, so the script reinstalled a
rem program that was already there. Probe the install location too.
if exist "%LOCALAPPDATA%\Programs\Ollama\ollama.exe" set "PATH=%LOCALAPPDATA%\Programs\Ollama;%PATH%"
where ollama >nul 2>&1
if errorlevel 1 goto install_ollama
echo Updating Ollama
echo [installOllama] winget upgrade Ollama.Ollama >> "%logFile%"
winget upgrade --id Ollama.Ollama -e --architecture x64 --silent --disable-interactivity --accept-package-agreements --accept-source-agreements >> "%logFile%" 2>&1
echo [installOllama] winget upgrade exit %errorlevel% >> "%logFile%"
if errorlevel 1 (echo Already current.) else (echo Updated.)
goto pull_model
:install_ollama
echo Installing Ollama
echo [installOllama] winget install Ollama.Ollama >> "%logFile%"
winget install --id Ollama.Ollama -e --architecture x64 --silent --disable-interactivity --accept-package-agreements --accept-source-agreements >> "%logFile%" 2>&1
echo [installOllama] winget install exit %errorlevel% >> "%logFile%"
set "PATH=%LOCALAPPDATA%\Programs\Ollama;%PATH%"
where ollama >nul 2>&1
if errorlevel 1 goto fail_ollama

:pull_model
call :ollamaModels
echo %modelList% | find /i "%modelName%" >nul 2>&1
if not errorlevel 1 (
  echo The %modelName% model is already installed.
  echo [installOllama] model already present >> "%logFile%"
  goto done_model
)
echo Fetching the %modelName% model, about 2 GB
echo [installOllama] ollama pull %modelName% >> "%logFile%"
call :ollamaPullHidden %modelName%
echo [installOllama] ollama pull exit %errorlevel% >> "%logFile%"
if errorlevel 1 goto fail_model

:done_model
echo Done.
echo [installOllama] done >> "%logFile%"
exit /b 0

:fail_ollama
echo Ollama was not found after the install step. The log is:
echo %logFile%
echo [installOllama] FAILED: ollama not found after install >> "%logFile%"
exit /b 3

:fail_model
echo The model download did not finish. Run this script again later, or
echo run this at a command prompt: ollama pull %modelName%
echo [installOllama] model pull failed >> "%logFile%"
exit /b 4

rem ---- Talking to Ollama without opening a window ----------------------
rem The ollama command starts its server in a console of its own when the
rem server is not already running, and that window stays on screen looking
rem like something has gone wrong. Ollama also answers over a local web
rem interface, which opens nothing, so presence and model lists are asked
rem that way; only a download needs the command, and that runs hidden.

:ollamaModels
rem Sets modelList to the names Ollama reports, or leaves it empty.
set "modelList="
for /f "delims=" %%m in ('powershell -NoProfile -Command "try { (Invoke-RestMethod -Uri http://localhost:11434/api/tags -TimeoutSec 10).models.name -join \" \" } catch { \"\" }" 2^>nul') do set "modelList=%%m"
exit /b 0

:ollamaPullHidden
rem Downloads %1 with no window of any kind. The command writes its
rem progress to the log rather than to a console nobody should have to
rem look at.
powershell -NoProfile -Command "$p = Start-Process -FilePath 'ollama' -ArgumentList 'pull','%~1' -WindowStyle Hidden -PassThru -Wait; exit $p.ExitCode" >> "%logFile%" 2>&1
exit /b %errorlevel%
