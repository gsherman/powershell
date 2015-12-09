@echo off

cd /d %~dp0

if "%1" == "" goto usage 
if "%2" == "" goto usage 
IF "%~3" == "" GOTO usage

SET POWERSHELL=%SystemRoot%\SYSTEM32\WindowsPowerShell\v1.0\powershell.exe

%POWERSHELL% ./CaseMessage.ps1 %1 %2 ""%3""

goto done

:usage
echo.
echo Missing arguments!
echo Usage: CaseMessage caseIdNumber slackChannel(without the pound sign) message
echo Example: CaseMessage 12345 dovetail "Case Dispatch Notification"
echo.

:done



