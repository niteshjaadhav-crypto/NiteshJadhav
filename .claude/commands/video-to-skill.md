---
description: Download a video from a link, grab frames, transcribe the audio, save it as a Claude skill, and delete the video
argument-hint: <video-url> [skill-name]
---

# Video → Skill

Turn the video at this link into a reusable Claude skill: **$ARGUMENTS**

The first argument is the video URL (YouTube, Instagram, etc. — anything yt-dlp supports).
The optional second argument is the skill name. If it is missing, the script names the skill after the video title.

## Step 1 — Find the right Python and make sure the tools are installed

faster-whisper needs Python 3.10–3.13 (newer versions such as 3.15 cannot install it yet).
Pick the first command below that prints `OK`, and use it as `PY` for the rest of this command:

- Windows: `py -3.12 -c "import faster_whisper; print('OK')"` (also try `py -3.13`, `py -3.11`)
- Mac/Linux: `python3 -c "import faster_whisper; print('OK')"`

If none prints `OK`, run `<PY> -m pip install yt-dlp faster-whisper` with a Python 3.10–3.13. If no such Python exists,
tell the user to install Python 3.12 and stop.

`ffmpeg` must also be available. If it is missing, tell the user and stop.

## Step 2 — Download, grab frames, transcribe, delete the video

```bash
<PY> scripts/video_to_skill.py $ARGUMENTS
```

The script:
1. downloads the video (max 720p) into a temporary folder
2. saves up to 12 evenly spaced frames to `.claude/skills/<name>/frames/`
3. transcribes the audio with Whisper (it falls back to the video's subtitles if Whisper fails) and saves it to `.claude/skills/<name>/reference/transcript.md`
4. saves the video's title, channel, duration and description to `reference/meta.json`
5. **always deletes the downloaded video and audio**, even if a step fails

The last line of its output is a JSON summary. If `"ok": false`, show the user the error and stop.
Useful options: `--max-frames 20`, `--model medium` (more accurate but slower), `--language hi`.

## Step 3 — Study the material

- Read `reference/transcript.md` in full.
- Look at every frame in `frames/` with the Read tool. Note what is shown on screen: poses, steps, text overlays, slides, code, diagrams, and the layout or style.
- Read `reference/meta.json` for context.

## Step 4 — Write the skill

Create `.claude/skills/<name>/SKILL.md` in this format:

```markdown
---
name: <name>
description: <One or two sentences: what this skill does AND when Claude should use it. Be specific so it triggers correctly.>
---

# <Clear title>

Source: <video title> by <channel> — <url>

## Overview
<What the video teaches, in 2–4 sentences>

## Key steps / method
<Numbered, actionable steps taken from the video, in order>

## Rules, tips and warnings
<Bullet list of the important do's, don'ts, numbers and details>

## Visual reference
<For each useful frame: `frames/<file>` (mm:ss) — what it shows>

## References
- Full transcript: `reference/transcript.md`
- Video metadata: `reference/meta.json`
```

Guidelines:
- Write instructions Claude can follow, not a summary of the video. Turn what the speaker says into steps and rules.
- Keep only the information that is useful. Leave out intros, "like and subscribe" and sponsor segments.
- Keep the frames that show something useful, and delete the blurry, duplicate or blank ones. Remove a deleted frame from the "Visual reference" list too.

## Step 5 — Report back

Tell the user:
- the skill name and folder path
- a 3–5 line summary of what the skill teaches
- how many frames were kept and how the transcript was made (Whisper or subtitles)
- confirmation that the downloaded video was deleted (`video_deleted: true` in the summary)
