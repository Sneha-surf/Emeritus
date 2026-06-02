#!/usr/bin/env python3
"""
POC 2 — Stage 1: Input Ingestion

Extracts from the provided PPTX and green screen video:
  - Speaker notes per slide  →  notes.json
  - Slide images             →  slides/ directory (reuses pipeline.pptx_to_images)
  - Audio track              →  faculty_audio.wav  (used for voice cloning in Stage 3)

Usage:
  python poc2_stage1_ingest.py <pptx_path> <faculty_video> <output_dir>

Example:
  python poc2_stage1_ingest.py slides.pptx faculty.mp4 ./poc2_output
"""

import sys
import json
import subprocess
from pathlib import Path

# Reuse pptx_to_images and natural_sort_key from the existing POC 1 pipeline
sys.path.insert(0, str(Path(__file__).parent))
from pipeline import pptx_to_images


# ── Speaker Notes Extraction ───────────────────────────────────────────────────

def extract_speaker_notes(pptx_path: str) -> dict:
    """
    Parse all slides and return {slide_index (int): notes_text (str)}.
    Slides with no notes return an empty string.
    """
    try:
        from pptx import Presentation
    except ImportError:
        print("[ERROR] python-pptx not installed. Run: pip install python-pptx")
        sys.exit(1)

    prs = Presentation(pptx_path)
    notes = {}
    for i, slide in enumerate(prs.slides):
        try:
            notes_slide = slide.notes_slide
            text = notes_slide.notes_text_frame.text.strip() if notes_slide else ""
        except Exception:
            text = ""
        notes[i] = text
    return notes


# ── Audio Extraction ───────────────────────────────────────────────────────────

def extract_audio(video_path: str, output_wav: str) -> bool:
    """Extract audio from video to a 44.1kHz stereo WAV using FFmpeg."""
    result = subprocess.run([
        'ffmpeg', '-y', '-i', video_path,
        '-vn',
        '-acodec', 'pcm_s16le',
        '-ar', '44100',
        '-ac', '2',
        output_wav
    ], capture_output=True, text=True)

    if result.returncode != 0:
        print(f"[WARN] Audio extraction failed:\n       {result.stderr.strip()}")
        return False
    return True


# ── Stage 1 Runner ─────────────────────────────────────────────────────────────

def run_stage1(pptx_path: str, video_path: str, output_dir: str) -> dict:
    """
    Run Stage 1 ingestion. Returns a manifest dict with paths to all outputs.

    manifest keys:
      notes        — path to notes.json
      slides_dir   — directory containing slide PNGs
      slide_count  — number of slides
      audio        — path to faculty_audio.wav, or None if extraction failed
    """
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    slides_dir = str(out / "slides")

    # ── Speaker notes ──────────────────────────────────────────────────────────
    print("[Stage 1/3] Extracting speaker notes …")
    notes = extract_speaker_notes(pptx_path)
    notes_path = out / "notes.json"
    with open(notes_path, 'w', encoding='utf-8') as f:
        json.dump(notes, f, indent=2, ensure_ascii=False)
    filled = sum(1 for v in notes.values() if v)
    empty  = len(notes) - filled
    print(f"            {len(notes)} slides total — {filled} with notes, {empty} empty")
    print(f"            Notes → {notes_path}")

    # ── Slide images ───────────────────────────────────────────────────────────
    print("[Stage 1/3] Exporting slide images …")
    slide_paths = pptx_to_images(pptx_path, slides_dir)
    print(f"            {len(slide_paths)} slide images → {slides_dir}/")

    # ── Audio ──────────────────────────────────────────────────────────────────
    print("[Stage 1/3] Extracting audio from faculty video …")
    audio_path = str(out / "faculty_audio.wav")
    audio_ok = extract_audio(video_path, audio_path)
    if audio_ok:
        size_mb = Path(audio_path).stat().st_size / 1024 / 1024
        print(f"            Audio ({size_mb:.1f} MB) → {audio_path}")
    else:
        audio_path = None
        print("            [WARN] Audio extraction failed — voice cloning will be skipped")

    manifest = {
        "notes":       str(notes_path),
        "slides_dir":  slides_dir,
        "slide_count": len(slide_paths),
        "audio":       audio_path,
    }

    manifest_path = out / "stage1_manifest.json"
    with open(manifest_path, 'w') as f:
        json.dump(manifest, f, indent=2)
    print(f"\n[Stage 1 Complete] Manifest → {manifest_path}")

    return manifest


# ── CLI ────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print(__doc__)
        sys.exit(1)

    result = run_stage1(
        pptx_path   = sys.argv[1],
        video_path  = sys.argv[2],
        output_dir  = sys.argv[3],
    )
    print()
    for k, v in result.items():
        print(f"  {k}: {v}")
