#!/usr/bin/env python3
"""
video_to_skill.py — download a video, grab key frames, transcribe the audio,
then DELETE the video so only the lightweight material is left.

Usage:
    python3 scripts/video_to_skill.py <video-url> <skill-name> [options]

Output (inside .claude/skills/<skill-name>/):
    reference/transcript.md   timestamped transcript
    reference/transcript.txt  plain transcript text
    reference/meta.json       title, uploader, duration, url, ...
    frames/frame_XXX_mm-ss.jpg  key frames (scaled to 720px wide)

The video and audio files are always deleted at the end, even if a step fails.
A JSON summary is printed on the last line of stdout so Claude can read it.

Requirements: ffmpeg, yt-dlp, faster-whisper
    pip install yt-dlp faster-whisper
"""

import argparse
import glob
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile


def log(msg):
    print(f"[video-to-skill] {msg}", file=sys.stderr, flush=True)


def run(cmd):
    log("$ " + " ".join(cmd))
    subprocess.run(cmd, check=True)


def slugify(name):
    s = re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")
    return s[:60] or "video-skill"


def fmt_ts(seconds):
    seconds = int(seconds)
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h:02d}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"


def require(tool):
    if shutil.which(tool) is None:
        sys.exit(f"Missing '{tool}'. Install it first (see script header).")


# ---------------------------------------------------------------- download
def download(url, workdir):
    out_tmpl = os.path.join(workdir, "video.%(ext)s")
    run([
        "yt-dlp", "--no-playlist",
        "-f", "bv*[height<=720]+ba/b[height<=720]/b",
        "--merge-output-format", "mp4",
        "--write-info-json",
        # grab subtitles too: used as a fallback if transcription fails
        "--write-subs", "--write-auto-subs", "--sub-langs", "en.*,hi.*",
        "--sub-format", "vtt", "--ignore-errors",
        "-o", out_tmpl, url,
    ])
    videos = [p for p in glob.glob(os.path.join(workdir, "video.*"))
              if p.endswith((".mp4", ".mkv", ".webm", ".mov"))]
    if not videos:
        raise RuntimeError("Download failed: no video file produced.")
    info = {}
    info_path = os.path.join(workdir, "video.info.json")
    if os.path.exists(info_path):
        with open(info_path, encoding="utf-8") as f:
            raw = json.load(f)
        info = {k: raw.get(k) for k in (
            "title", "uploader", "channel", "duration", "upload_date",
            "description", "webpage_url", "tags")}
    return videos[0], info


def probe_duration(video):
    out = subprocess.run(
        ["ffprobe", "-v", "error", "-show_entries", "format=duration",
         "-of", "default=nw=1:nk=1", video],
        capture_output=True, text=True).stdout.strip()
    try:
        return float(out)
    except ValueError:
        return 0.0


# ------------------------------------------------------------------ frames
def grab_frames(video, frames_dir, duration, max_frames):
    """Evenly spaced frames across the video (skip the first/last 2%)."""
    os.makedirs(frames_dir, exist_ok=True)
    if duration <= 0:
        duration = 60.0
    n = max(1, min(max_frames, int(duration // 2) or 1))
    start, end = duration * 0.02, duration * 0.98
    step = (end - start) / n
    saved = []
    for i in range(n):
        t = start + step * (i + 0.5)
        name = f"frame_{i + 1:03d}_{fmt_ts(t).replace(':', '-')}.jpg"
        path = os.path.join(frames_dir, name)
        subprocess.run(
            ["ffmpeg", "-loglevel", "error", "-y", "-ss", f"{t:.2f}",
             "-i", video, "-frames:v", "1", "-vf", "scale=720:-2",
             "-q:v", "4", path],
            check=False)
        if os.path.exists(path):
            saved.append({"file": name, "time": fmt_ts(t)})
    log(f"saved {len(saved)} frames")
    return saved


# -------------------------------------------------------------- transcript
def extract_audio(video, workdir):
    audio = os.path.join(workdir, "audio.wav")
    run(["ffmpeg", "-loglevel", "error", "-y", "-i", video,
         "-vn", "-ac", "1", "-ar", "16000", audio])
    return audio


def transcribe_whisper(audio, model_size, language):
    from faster_whisper import WhisperModel
    log(f"transcribing with faster-whisper ({model_size}) ...")
    model = WhisperModel(model_size, device="cpu", compute_type="int8")
    segments, info = model.transcribe(
        audio, language=language or None, vad_filter=True)
    segs = [{"start": s.start, "end": s.end, "text": s.text.strip()}
            for s in segments if s.text.strip()]
    return segs, info.language


def transcribe_from_vtt(workdir):
    vtts = sorted(glob.glob(os.path.join(workdir, "*.vtt")))
    if not vtts:
        return []
    log(f"using subtitles fallback: {os.path.basename(vtts[0])}")
    segs, last = [], None
    ts_re = re.compile(r"(\d+):(\d+):(\d+)\.\d+ -->")
    cur = None
    with open(vtts[0], encoding="utf-8", errors="ignore") as f:
        for line in f:
            line = line.strip()
            m = ts_re.match(line)
            if m:
                h, mi, s = map(int, m.groups())
                cur = h * 3600 + mi * 60 + s
                continue
            if not line or cur is None or line.startswith(("WEBVTT", "Kind:", "Language:")):
                continue
            text = re.sub(r"<[^>]+>", "", line).strip()
            if text and text != last:
                segs.append({"start": cur, "end": cur, "text": text})
                last = text
    return segs


def write_transcript(segs, ref_dir, title):
    md = [f"# Transcript — {title}\n"]
    md += [f"**[{fmt_ts(s['start'])}]** {s['text']}  " for s in segs]
    with open(os.path.join(ref_dir, "transcript.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(md) + "\n")
    plain = " ".join(s["text"] for s in segs)
    with open(os.path.join(ref_dir, "transcript.txt"), "w", encoding="utf-8") as f:
        f.write(plain + "\n")
    return len(plain.split())


# -------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("url")
    ap.add_argument("skill_name", nargs="?", default="")
    ap.add_argument("--skills-dir", default=".claude/skills")
    ap.add_argument("--max-frames", type=int, default=12)
    ap.add_argument("--model", default="small",
                    help="whisper size: tiny/base/small/medium/large-v3")
    ap.add_argument("--language", default="",
                    help="force language code, e.g. en or hi (default: auto)")
    args = ap.parse_args()

    for tool in ("ffmpeg", "ffprobe", "yt-dlp"):
        require(tool)

    workdir = tempfile.mkdtemp(prefix="video2skill_")
    summary = {"ok": False}
    try:
        video, info = download(args.url, workdir)
        title = info.get("title") or "video"
        name = slugify(args.skill_name or title)
        skill_dir = os.path.abspath(os.path.join(args.skills_dir, name))
        ref_dir = os.path.join(skill_dir, "reference")
        frames_dir = os.path.join(skill_dir, "frames")
        os.makedirs(ref_dir, exist_ok=True)

        duration = info.get("duration") or probe_duration(video)
        info["duration"] = duration
        info["source_url"] = args.url
        with open(os.path.join(ref_dir, "meta.json"), "w", encoding="utf-8") as f:
            json.dump(info, f, indent=2, ensure_ascii=False)

        frames = grab_frames(video, frames_dir, duration, args.max_frames)

        segs, lang, method = [], None, "none"
        try:
            audio = extract_audio(video, workdir)
            segs, lang = transcribe_whisper(audio, args.model, args.language)
            method = "whisper"
        except Exception as e:  # noqa: BLE001
            log(f"whisper failed ({e}); trying subtitles")
        if not segs:
            segs = transcribe_from_vtt(workdir)
            method = "subtitles" if segs else "none"
        words = write_transcript(segs, ref_dir, title)

        summary = {
            "ok": True,
            "skill_name": name,
            "skill_dir": skill_dir,
            "title": title,
            "duration": fmt_ts(duration),
            "language": lang,
            "transcript_method": method,
            "transcript_words": words,
            "transcript_file": os.path.join(ref_dir, "transcript.md"),
            "frames_dir": frames_dir,
            "frames": frames,
        }
    except Exception as e:  # noqa: BLE001
        summary = {"ok": False, "error": str(e)}
    finally:
        # always remove the downloaded video / audio
        shutil.rmtree(workdir, ignore_errors=True)
        summary["video_deleted"] = not os.path.exists(workdir)
        log("downloaded video deleted" if summary["video_deleted"]
            else "WARNING: could not delete temp dir " + workdir)

    print(json.dumps(summary, ensure_ascii=False))
    sys.exit(0 if summary.get("ok") else 1)


if __name__ == "__main__":
    main()
