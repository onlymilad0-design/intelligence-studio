---
name: instaloader
description: Download public Instagram content (posts, profile pictures, stories, highlights, reels, captions, comments, metadata) with the instaloader CLI or Python library. Use when the user asks to download, archive, or scrape Instagram profiles, hashtags, or posts.
---

# Instaloader

Source: https://github.com/instaloader/instaloader — docs: https://instaloader.github.io

## Setup

```bash
pip install -r .claude/skills/instaloader/requirements.txt
instaloader --version
```

## Common CLI usage

| Task | Command |
|---|---|
| Download a public profile | `instaloader PROFILE` |
| Only the profile picture | `instaloader --no-posts PROFILE` |
| A single post by shortcode | `instaloader -- -SHORTCODE` (shortcode from `instagram.com/p/SHORTCODE/`) |
| Hashtag | `instaloader "#hashtag" --count 50` |
| Skip videos / thumbnails | `--no-videos --no-video-thumbnails` |
| Include comments / geotags | `--comments --geotags` |
| Update an existing archive | `instaloader --fast-update PROFILE` |
| Stories / highlights (login required) | `instaloader --login USER --stories --highlights PROFILE` |
| Custom output folder | `--dirname-pattern "downloads/{profile}"` |

Login: `instaloader --login USERNAME` prompts for the password interactively and saves a session file. Never put passwords in commands, files, or commits.

## Python API

```python
import instaloader

L = instaloader.Instaloader(download_videos=True, save_metadata=False)
profile = instaloader.Profile.from_username(L.context, "PROFILE")
for post in profile.get_posts():
    L.download_post(post, target=profile.username)

post = instaloader.Post.from_shortcode(L.context, "SHORTCODE")
L.download_post(post, target="single_post")
```

## Notes

- Instagram rate-limits heavily; instaloader backs off automatically. For large jobs use `--count` and `--fast-update`.
- Anonymous access often hits "401 / Please wait a few minutes" — logging in with a session usually fixes it.
- Only download content you have the right to use; respect Instagram's terms and people's privacy.
