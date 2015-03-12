@echo off

cd /d %~dp0
SET POWERSHELL=%SystemRoot%\SYSTEM32\WindowsPowerShell\v1.0\powershell.exe

%POWERSHELL% ./SetSubcaseOriginator.ps1 %*
