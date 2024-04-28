@echo off

chcp 65001 > nul

start /wait /b "%~n0" powershell.exe -ExecutionPolicy RemoteSigned -File "%~dp0%~n0.ps1" %* -Parameters "%param%"

if %errorlevel% neq 0 pause