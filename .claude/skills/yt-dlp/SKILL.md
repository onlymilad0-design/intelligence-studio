---
name: yt-dlp
description: Download video or audio from YouTube and 1000+ other sites (TikTok, X/Twitter, Vimeo, SoundCloud...) with yt-dlp. Use when the user asks to download a video, audio/MP3, playlist, channel, or subtitles from a URL.
---

# yt-dlp

Source: https://github.com/yt-dlp/yt-dlp — needs `ffmpeg` for merging and audio conversion.

## Setup

```bash
pip install -U yt-dlp   # ffmpeg: apt-get install -y ffmpeg
```

## Common usage

| Task | Command |
|---|---|
| Best quality video | `yt-dlp URL` |
| MP4 up to 1080p | `yt-dlp -f "bv*[height<=1080][ext=mp4]+ba[ext=m4a]/b[ext=mp4]" URL` |
| Audio only as MP3 | `yt-dlp -x --audio-format mp3 URL` |
| List available formats | `yt-dlp -F URL` |
| Subtitles (embedded) | `yt-dlp --write-subs --sub-langs "ar,en" --embed-subs URL` |
| Playlist, items 1-10 | `yt-dlp --playlist-items 1-10 URL` |
| Output folder/name | `yt-dlp -o "downloads/%(title)s.%(ext)s" URL` |
| Skip already downloaded | `yt-dlp --download-archive archive.txt URL` |
| Metadata only (JSON) | `yt-dlp -j --skip-download URL` |

## Python API

```python
import yt_dlp
with yt_dlp.YoutubeDL({"format": "bestaudio", "outtmpl": "%(title)s.%(ext)s"}) as ydl:
    ydl.download(["URL"])
```

## Notes

- If a site breaks, update first: `pip install -U yt-dlp`.
- Age/login-restricted videos need cookies (`--cookies cookies.txt`); never commit cookie files.
- Only download content you have the right to use.
