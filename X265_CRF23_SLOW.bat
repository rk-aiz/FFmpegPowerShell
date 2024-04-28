@echo off

chcp 65001 > nul

PowerShell -ExecutionPolicy Bypass -Command "Get-ChildItem -LiteralPath '%~dp0' -Recurse | Unblock-File"

set param=^
 -hide_banner^
 -loglevel info^
 -ignore_unknown^
 -i [INPUT]^
 -analyzeduration 20M -probesize 20M^
 -c:v libx265^
 -maxrate 20M^
 -bufsize 20M^
 -preset slow^
 -crf 23^
 -pix_fmt yuv420p^
 -bf 2^
 -movflags +faststart^
 -c:a aac -ab 256000 -af "channelmap=channel_layout=stereo,aresample=48000:resampler=soxr"^
 [OUTPUT(_x265_c23slow.mp4)]

start /wait /b "%~n0" powershell.exe -ExecutionPolicy RemoteSigned -File "Script\FFmpeg_PowerShell_GUI.ps1" %* -Parameters "%param%"

if %errorlevel% neq 0 pause