# Video → Skill

A Claude Code command that turns any video link into a reusable Claude skill.

```
/video-to-skill <video-url> [skill-name]
```

What it does:
1. **Downloads** the video (YouTube, Instagram, etc. — anything `yt-dlp` supports)
2. **Grabs key frames** (up to 12, evenly spaced) → `.claude/skills/<name>/frames/`
3. **Transcribes the audio** with Whisper (falls back to the video's subtitles) → `reference/transcript.md`
4. Claude studies the frames and transcript and **writes `SKILL.md`**
5. **Deletes the downloaded video** (always, even if a step fails)

## Setup

```bash
pip install -r requirements.txt     # yt-dlp + faster-whisper (needs Python 3.10–3.13; on Windows: py -3.12 -m pip install -r requirements.txt)
# plus ffmpeg:  sudo apt install ffmpeg   |   brew install ffmpeg   |   winget install ffmpeg
```

## Examples

```
/video-to-skill https://www.youtube.com/watch?v=XXXX surya-namaskar
/video-to-skill https://www.instagram.com/reel/XXXX/
```

Run the script by itself (without Claude writing the SKILL.md):

```bash
python3 scripts/video_to_skill.py <url> [name] --max-frames 20 --model medium --language hi
```

- `--model`: `tiny` / `base` / `small` (default) / `medium` / `large-v3`. Larger is more accurate but slower.
- `--language`: force a language such as `en` or `hi`. By default it is detected automatically.

Download only videos you have the rights to use.
