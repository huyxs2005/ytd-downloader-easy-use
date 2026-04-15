@echo off
title yt-dlp downloader
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0download_playlist.ps1" %*
endlocal
