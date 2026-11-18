"""Prepare the long screen recording for the guide without changing its source."""
from __future__ import annotations

import argparse
from pathlib import Path
import subprocess

import imageio_ffmpeg


ROOT = Path(__file__).resolve().parents[1]


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", nargs="?", type=Path, default=ROOT / "transfer_news_guide.mp4")
    parser.add_argument("--output", type=Path, default=ROOT / "static" / "walkthrough-detailed.mp4")
    parser.add_argument("--poster", type=Path, default=ROOT / "static" / "walkthrough-detailed-poster.jpg")
    args = parser.parse_args()
    source = args.source.resolve()
    if not source.is_file():
        parser.error(f"Recording not found: {source}")
    if source in (args.output.resolve(), args.poster.resolve()):
        parser.error("Output paths must differ from the original recording.")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.poster.parent.mkdir(parents=True, exist_ok=True)
    ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
    # Screen text keeps its original pixel dimensions. The frame-rate and
    # rate ceiling keep the 12-minute recording within ordinary Git file limits.
    subprocess.run([ffmpeg, "-hide_banner", "-y", "-i", str(source),
        "-map", "0:v:0", "-map", "0:a:0?", "-map_metadata", "-1",
        "-vf", "fps=20", "-c:v", "libx264", "-preset", "medium", "-crf", "26",
        "-maxrate", "700k", "-bufsize", "1400k", "-pix_fmt", "yuv420p",
        "-c:a", "aac", "-b:a", "64k", "-ac", "1", "-movflags", "+faststart",
        str(args.output)], check=True)
    # An empty upload form makes a useful poster without displaying roster rows.
    subprocess.run([ffmpeg, "-hide_banner", "-loglevel", "error", "-y", "-ss", "00:03:00",
        "-i", str(source), "-frames:v", "1", "-vf", "scale=1280:-2", "-q:v", "3", str(args.poster)], check=True)
    original, compressed = source.stat().st_size, args.output.stat().st_size
    print(f"Original: {original / 1_000_000:.1f} MB; web copy: {compressed / 1_000_000:.1f} MB "
          f"({100 * (1 - compressed / original):.1f}% smaller)")


if __name__ == "__main__":
    main()
