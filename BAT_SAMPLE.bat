@echo off

chcp 932 > nul

set param=^
 -hide_banner^
 -loglevel info^
 -ignore_unknown^
 -i [INPUT]^
 -analyzeduration 20M -probesize 20M^
 -c:v libx264^
 -maxrate 20M^
 -bufsize 20M^
 -preset medium^
 -crf 18^
 -pix_fmt yuv420p^
 -bf 2^
 -movflags +faststart^
 -c:a aac -ab 256000 -af "channelmap=channel_layout=stereo,aresample=48000:resampler=soxr"^
 [OUTPUT(_x264_c18medium.mp4)]

start /wait /b "%~n0" powershell.exe -ExecutionPolicy RemoteSigned -File "Script\FFmpeg_PowerShell_GUI.ps1" %* -Parameters "%param%"

if %errorlevel% neq 0 pause