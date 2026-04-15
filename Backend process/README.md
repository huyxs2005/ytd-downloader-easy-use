# ytd downloader easy use

Simple Windows wrapper around `yt-dlp` for downloading YouTube and YouTube Music playlists with a PowerShell menu.

## What It Does

- Downloads a single video or a full playlist
- Supports `Audio` and `Video` modes
- Uses `ffmpeg` for post-processing
- Writes numbered files for playlist order
- Opens worker PowerShell windows so you can see live download progress
- Stores metadata with each downloaded item

## Included Files

- `download_playlist.ps1` - main interactive downloader script
- `yt-dlp.exe` - downloader binary
- `bin/` - `ffmpeg`, `ffprobe`, and related tools
- `cookies.txt` - cookie file used for authenticated downloads

## Requirements

- Windows PowerShell
- `yt-dlp.exe`
- `ffmpeg` binaries in `bin/`
- Valid `cookies.txt` if the target content requires login

## How To Use

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\download_playlist.ps1
```

Then choose:

1. Create new or update existing
2. Audio only or video only
3. Paste a YouTube or YouTube Music URL

## Notes

- This repo is meant for personal local use on Windows.
- Do not commit real cookies or personal downloaded media to a public repository.
