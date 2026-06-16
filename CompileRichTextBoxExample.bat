@echo off
cls
echo Compileing
if exist RichTextBoxExample.exe del RichTextBoxExample.exe
jsc.exe -nologo -t:winexe -fast- RichTextBoxExample.js
if exist RichTextBoxExample.exe echo Done
