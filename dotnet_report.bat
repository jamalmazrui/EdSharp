@echo off
cls
if exist "%1" del "%1"
start /wait dotnet.exe "%1"
rem start "%1"
notepad.exe "%1"
rem start iexplore "%1"
