@echo off
setlocal enabledelayedexpansion
set "sScript=%~n0"
set "sCallerDir=%~dp0"
if not exist "%~dp0homerInstall.cmd" (
  echo(
  echo homerInstall.cmd is missing from %~dp0
  echo That file is part of this program. Reinstall, or copy it from the
  echo program's zip into this folder, and run this again.
  echo(
  if not defined noPause pause
  exit /b 1
)
call "%~dp0homerInstall.cmd" setup "%~f0" %*
rem installTranslateModel.cmd -- install the larger AI model EdSharp uses for
rem translation when it is present. The small chat model translates
rem passably; this one translates well, at about 5 gigabytes.
rem Probe first, log milestones, never pause; the Results box reports the
rem outcome. The console says only what is happening.

if exist "%LOCALAPPDATA%\Programs\Ollama" set "PATH=%LOCALAPPDATA%\Programs\Ollama;%PATH%"
where ollama >nul 2>&1
if errorlevel 1 goto no_ollama

call :ollamaModels
echo %modelList% | find /i "%modelName%" >nul 2>&1
if not errorlevel 1 (
  echo The %modelName% model is already installed.
  echo [installTranslateModel] already present >> "%log%"
  exit /b 0
)

echo Fetching the %modelName% model, about 5 GB
echo [installTranslateModel] ollama pull %modelName% >> "%log%"
call :ollamaPullHidden %modelName%
echo [installTranslateModel] pull exit %errorlevel% >> "%log%"
if errorlevel 1 goto failed
echo Done.
echo [installTranslateModel] done >> "%log%"
exit /b 0

:no_ollama
echo Ollama is not installed, so the translation model cannot be fetched.
echo Tick the Ollama box as well, or run installOllama.cmd first.
echo [installTranslateModel] FAILED: no ollama >> "%log%"
exit /b 7

:failed
echo The model did not download. The log is:
echo %log%
echo [installTranslateModel] FAILED >> "%log%"
exit /b 3
