@echo off
title yt-dlp downloader
setlocal
cd /d "%~dp0Backend process"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Backend process\download_playlist.ps1" %*
endlocal
