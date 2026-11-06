#!/usr/bin/env python3
"""Re-encode the videos that pptx_canvas.py extracted from a .pptx into small web MP4s (H.264, at most 720p, AAC)
and point the deck at them, so that they can be committed and hosted with the slides.

  FFMPEG=/path/to/ffmpeg python web_videos.py deck.json [deck2.json ...]

Needs ffmpeg: on the PATH, or named in FFMPEG (a virtual environment with `pip install imageio-ffmpeg` has one:
`python -c 'import imageio_ffmpeg; print(imageio_ffmpeg.get_ffmpeg_exe())'`). The re-encoded file is media/<hash>.web.mp4; the original is deleted. Videos that are
already <name>.web.mp4 are left alone.
"""
import json, shutil, subprocess, sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import pptx_canvas as pc


def ffmpeg():
    import os
    if os.environ.get("FFMPEG"):
        return os.environ["FFMPEG"]
    try:
        import imageio_ffmpeg
        return imageio_ffmpeg.get_ffmpeg_exe()
    except ImportError:
        exe = shutil.which("ffmpeg")
        if not exe:
            sys.exit("ffmpeg not found: use a virtual environment with imageio-ffmpeg or install ffmpeg")
        return exe


def main():
    exe = ffmpeg()
    for deck_path in map(Path, sys.argv[1:]):
        deck = json.loads(deck_path.read_text(encoding="utf-8"))
        changed = False
        for s in deck["slides"]:
            for it in s.get("items", []) if s["layout"] == "canvas" else []:
                v = it.get("video")
                if not v or v.endswith(".web.mp4"):
                    continue
                src = deck_path.parent / v
                if not src.exists():
                    print(f"  missing {src}")
                    continue
                dst = src.with_suffix("").with_suffix(".web.mp4") if False else src.parent / (src.stem + ".web.mp4")
                if not dst.exists():
                    cmd = [exe, "-y", "-loglevel", "error", "-i", str(src), "-vf", "scale='min(1280,iw)':-2", "-c:v", "libx264",
                           "-preset", "slow", "-crf", "28", "-pix_fmt", "yuv420p", "-c:a", "aac", "-b:a", "96k", "-movflags", "+faststart", str(dst)]
                    subprocess.run(cmd, check=True)
                    print(f"  {src.name}: {src.stat().st_size / 1e6:.1f} MB -> {dst.name}: {dst.stat().st_size / 1e6:.1f} MB")
                it["video"] = f"{Path(v).parent}/{dst.name}"
                changed = True
        if changed:
            deck_path.write_text(pc.dumps(deck), encoding="utf-8")
            for f in deck_path.parent.glob("media/*"):
                if not f.name.endswith(".web.mp4") and not any(f.name == Path(i.get("video", "")).name for s in deck["slides"] for i in s.get("items", [])):
                    f.unlink()


if __name__ == "__main__":
    main()
