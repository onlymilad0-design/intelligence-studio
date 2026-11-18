#!/bin/bash
# Installs instaloader + yt-dlp and their Claude skills for every project
pip install -q -U instaloader yt-dlp || true
command -v ffmpeg >/dev/null || (apt-get update -qq && apt-get install -y -qq ffmpeg) || true
D="$HOME/.claude/skills"
mkdir -p "$D/instaloader" "$D/yt-dlp"
cat > "$D/instaloader/SKILL.md" <<'EOF'
---
name: instaloader
description: Download public Instagram posts, profile pics, stories, reels and metadata with instaloader. Use for downloading or archiving Instagram content.
---

# Instaloader (https://github.com/instaloader/instaloader)

Install if missing: `pip install -U instaloader`

| Task | Command |
|---|---|
| Public profile | `instaloader PROFILE` |
| Profile picture only | `instaloader --no-posts PROFILE` |
| Single post | `instaloader -- -SHORTCODE` (from instagram.com/p/SHORTCODE/) |
| Hashtag | `instaloader "#tag" --count 50` |
| Skip videos | `--no-videos --no-video-thumbnails` |
| Comments / geotags | `--comments --geotags` |
| Update archive | `instaloader --fast-update PROFILE` |
| Stories / highlights (login) | `instaloader --login USER --stories --highlights PROFILE` |
| Output folder | `--dirname-pattern "downloads/{profile}"` |

Python: `L = instaloader.Instaloader(); post = instaloader.Post.from_shortcode(L.context, "SC"); L.download_post(post, target="out")`

Notes: `--login USER` prompts for the password; never put passwords in commands or files.
Instagram rate-limits; on 401/"wait a few minutes" log in or slow down. Respect privacy and terms.
EOF
cat > "$D/yt-dlp/SKILL.md" <<'EOF'
---
name: yt-dlp
description: Download video/audio from YouTube, TikTok, X, Vimeo, SoundCloud and 1000+ sites with yt-dlp. Use for downloading videos, MP3s, playlists or subtitles.
---

# yt-dlp (https://github.com/yt-dlp/yt-dlp)

Install if missing: `pip install -U yt-dlp` (needs ffmpeg for merging/audio).

| Task | Command |
|---|---|
| Best quality | `yt-dlp URL` |
| MP4 up to 1080p | `yt-dlp -f "bv*[height<=1080][ext=mp4]+ba[ext=m4a]/b[ext=mp4]" URL` |
| Audio as MP3 | `yt-dlp -x --audio-format mp3 URL` |
| List formats | `yt-dlp -F URL` |
| Subtitles | `yt-dlp --write-subs --sub-langs "ar,en" --embed-subs URL` |
| Playlist items | `yt-dlp --playlist-items 1-10 URL` |
| Output name | `yt-dlp -o "downloads/%(title)s.%(ext)s" URL` |
| Skip downloaded | `yt-dlp --download-archive archive.txt URL` |
| Metadata JSON | `yt-dlp -j --skip-download URL` |

Python: `with yt_dlp.YoutubeDL({"format": "bestaudio"}) as y: y.download(["URL"])`

Notes: if a site breaks, run `pip install -U yt-dlp` first. Restricted videos need `--cookies cookies.txt`; never commit cookies.
Only download content you have the right to use.
EOF
# Claude Code tools for every session: Graphify, Headroom, OmniRoute, codebase-memory-mcp
pip install -q --ignore-installed -U pyyaml certifi || true
pip install -q -U graphifyy "headroom-ai[code]" codebase-memory-mcp || true
npm install -g omniroute >/dev/null 2>&1 || true
T=$(mktemp -d)
gc() { git clone -q --depth 1 "https://github.com/$1" "$2"; }
gc rebelytics/one-skill-to-rule-them-all "$T/to" && rm -rf "$D/task-observer" && cp -r "$T/to/skills/task-observer" "$D/"
gc msitarzewski/agency-agents "$T/aa" && (cd "$T/aa" && ./scripts/install.sh --tool claude-code </dev/null >/dev/null 2>&1)
rm -rf "$T"
# OpenMontage video studio, cloned to ~/OpenMontage (run `make setup` there on first use)
[ -d "$HOME/OpenMontage/.git" ] || gc calesthio/OpenMontage "$HOME/OpenMontage" || true
mkdir -p "$D/openmontage"
cat > "$D/openmontage/SKILL.md" <<'EOF'
---
name: openmontage
description: Produce videos (explainers, ads, shorts, animations) with the OpenMontage agentic video studio at ~/OpenMontage. Use whenever the user asks to make, edit or render a video.
---
OpenMontage (https://github.com/calesthio/OpenMontage) lives at `~/OpenMontage`.
1. `cd ~/OpenMontage`; if `.venv` is missing run `make setup` (needs ffmpeg, Node, Python 3.10+).
2. Read `~/OpenMontage/CLAUDE.md` and `AGENT_GUIDE.md`, then follow its pipeline: pick from `pipeline_defs/`, read the stage skills in `skills/`.
3. Put outputs where the user wants them (default: the current project).
EOF
command -v graphify >/dev/null && graphify install --platform claude >/dev/null 2>&1 || true
command -v codebase-memory-mcp >/dev/null && codebase-memory-mcp install -y </dev/null >/dev/null 2>&1 || true
if command -v claude >/dev/null; then
  for m in anthropics/claude-plugins-official thedotmack/claude-mem; do claude plugin marketplace add "$m" >/dev/null 2>&1 || true; done
  for p in claude-code-setup@claude-plugins-official claude-mem@thedotmack; do claude plugin install "$p" -s user >/dev/null 2>&1 || true; done
fi
# Always-on guidance for every project
C="$HOME/.claude/CLAUDE.md"; touch "$C"
grep -q "## Always-on tools" "$C" || cat >> "$C" <<'EOF'

## Always-on tools
- Run the task-observer skill at the start of every session (before the first tool call).
- Codebase questions: use the codebase-memory-mcp tools / graphify graph before grepping files.
- Hand specialist work to a matching agency-agents subagent in ~/.claude/agents (e.g. marketing, sales, engineering).
- Video requests: use the openmontage skill.
EOF
exit 0
