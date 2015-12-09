@echo off

cd /d %~dp0

if "%1" == "" goto usage 
if "%2" == "" goto usage 
IF "%~3" == "" GOTO usage

SET POWERSHELL=%SystemRoot%\SYSTEM32\WindowsPowerShell\v1.0\powershell.exe

%POWERSHELL% ./SolutionMessage.ps1 %1 %2 ""%3""

goto done

:usage
echo.
echo Missing arguments!
echo Usage: SolutionMessage solutionIdNumber slackChannel(without the pound sign) message
echo Example: SolutionMessage 12345 clarify "Solution 12345 Created by someone"
echo.

:done
