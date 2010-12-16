@echo off

REM Directory where the powershell script is located:
SET WORKING_DIR=C:\repo\powershell

cd %WORKING_DIR%

SET POWERSHELL=%SystemRoot%\SYSTEM32\WindowsPowerShell\v1.0\powershell.exe

%POWERSHELL% "& ./ChangeCaseStatus.ps1 %*"
