@echo off
cls
if exist Unicode.exe del Unicode.exe
csc.exe -nologo Unicode.cs
if exist Unicode.exe Unicode.exe
