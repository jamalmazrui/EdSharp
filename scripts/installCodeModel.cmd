@echo off
setlocal enabledelayedexpansion
set "sScript=%~n0"
set "sCallerDir=%~dp0"
if not exist "%~dp0installCommon.cmd" (
  echo(
  echo installCommon.cmd is missing from %~dp0
  echo That file is part of this program. Reinstall, or copy it from the
  echo program's zip into this folder, and run this again.
  echo(
  if not defined noPause pause
  exit /b 1
)
call "%~dp0installCommon.cmd" setup "%~f0" %*
rem installCodeModel.cmd -- install the AI model EdSharp uses for coding
rem questions about the file in the window. About 5 gigabytes.
rem Probe first, log milestones, never pause; the Results box reports the
rem outcome. The console says only what is happening.

if exist "%LOCALAPPDATA%\Programs\Ollama" set "PATH=%LOCALAPPDATA%\Programs\Ollama;%PATH%"
where ollama >nul 2>&1
if errorlevel 1 goto no_ollama

call :ollamaModels
echo %modelList% | find /i "%modelName%" >nul 2>&1
if not errorlevel 1 (
  echo The %modelName% model is already installed.
  echo [installCodeModel] already present >> "%log%"
  exit /b 0
)

echo Fetching the %modelName% model, about 5 GB
echo [installCodeModel] ollama pull %modelName% >> "%log%"
call :ollamaPullHidden %modelName%
echo [installCodeModel] pull exit %errorlevel% >> "%log%"
if errorlevel 1 goto failed
echo Done.
echo [installCodeModel] done >> "%log%"
exit /b 0

:no_ollama
echo Ollama is not installed, so the coding model cannot be fetched.
echo Tick the Ollama box as well, or run installOllama.cmd first.
echo [installCodeModel] FAILED: no ollama >> "%log%"
exit /b 7

:failed
echo The model did not download. The log is:
echo %log%
echo [installCodeModel] FAILED >> "%log%"
exit /b 3
