@echo off

chcp 932 > nul

PowerShell -ExecutionPolicy Bypass -Command "Get-ChildItem -LiteralPath '%~dp0' -Recurse | Unblock-File"

set param=^
    -hide_banner^
    -loglevel info^
    -ignore_unknown^
    -i [INPUT]^
    -analyzeduration 50M -probesize 50M^
    -c:v libx264^
    -maxrate 50M^
    -bufsize 50M^
    -preset fast^
    -crf 5^
    -pix_fmt yuv420p^
    -bf 2^
    -movflags +faststart^
    -c:a aac -ab 256000 -af "channelmap=channel_layout=stereo,aresample=48000:resampler=soxr"^
    [OUTPUT(_x264_c5fast.mp4)]

start /wait /b "%~n0" powershell.exe -ExecutionPolicy RemoteSigned -File "%~dp0\Script\FFmpeg_PowerShell_GUI.ps1" %* -Parameters "%param%"

if %errorlevel% neq 0 pause