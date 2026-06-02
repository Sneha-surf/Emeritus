#!/usr/bin/env python3
"""
POC 2 — Stage 2: Reference Frame Extraction

Scans the green screen faculty video to find the single best reference frame:
  - Clear, frontal face detected
  - Sharp (high Laplacian variance — no motion blur)
  - Well-lit (brightness close to mid-range)
  - Face centred and not clipped at edges

Then removes the green screen background from that frame to produce a clean
RGBA reference image that avatar generation tools (Champ, HeyGen, etc.) can use.

Two background-removal strategies (tried in order):
  1. rembg      — ML-based, most accurate for complex edges (hair, hands)
  2. Chroma key — falls back to pipeline.apply_chromakey if rembg not installed

Outputs (written to <output_dir>/):
  reference_frame.jpg   — raw best frame (BGR)
  reference_clean.png   — background-removed RGBA reference image

Usage:
  python poc2_stage2_reference.py <faculty_video> <output_dir> [sample_interval]

Arguments:
  faculty_video    Green screen faculty .mp4 / .mov
  output_dir       Where to write outputs (created if missing)
  sample_interval  Sample every N frames (default: 30 ≈ 1s at 30fps)

Example:
  python poc2_stage2_reference.py faculty.mp4 ./poc2_output 15
"""

import sys
import cv2
import numpy as np
from pathlib import Path


# ── Frame Scoring ──────────────────────────────────────────────────────────────

def score_frame(frame: np.ndarray, face_cascade) -> float:
    """
    Score a candidate frame for reference suitability.
    Returns 0.0 if no face is detected; higher = better.

    Scoring components:
      sharpness        — Laplacian variance (motion-blur rejection)
      brightness_score — penalises too-dark or blown-out frames
      face_size_score  — larger face area relative to frame = better identity detail
      edge_penalty     — penalises faces partially outside the frame boundary
    """
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

    # Sharpness: Laplacian variance (blurry frames have low variance)
    sharpness = cv2.Laplacian(gray, cv2.CV_64F).var()

    # Brightness: prefer frames where mean pixel is close to mid-grey (128)
    mean_brightness = float(gray.mean())
    brightness_score = 1.0 - abs(mean_brightness - 128.0) / 128.0

    # Face detection
    faces = face_cascade.detectMultiScale(
        gray, scaleFactor=1.1, minNeighbors=5, minSize=(80, 80))
    if len(faces) == 0:
        return 0.0

    fx, fy, fw, fh = max(faces, key=lambda f: f[2] * f[3])
    h, w = frame.shape[:2]

    # Face area fraction
    face_size_score = (fw * fh) / float(w * h)

    # Edge penalty: face should not be clipped at frame boundary
    margin = 0.05
    edge_penalty = 1.0
    if fx < w * margin or (fx + fw) > w * (1.0 - margin):
        edge_penalty *= 0.5
    if fy < h * margin or (fy + fh) > h * (1.0 - margin):
        edge_penalty *= 0.5

    return (sharpness / 1000.0) * brightness_score * face_size_score * edge_penalty * 100.0


# ── Best Frame Extraction ──────────────────────────────────────────────────────

def extract_reference_frame(video_path: str, output_dir: str,
                             sample_interval: int = 30) -> str:
    """
    Scan the video and save the highest-scoring frame as reference_frame.jpg.
    Returns the path to the saved frame.
    """
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        raise RuntimeError(f"Cannot open video: {video_path}")

    fps   = cap.get(cv2.CAP_PROP_FPS) or 30.0
    total = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    print(f"[Stage 2] Scanning {total} frames @ {fps:.1f} fps "
          f"(every {sample_interval} frames = ~{sample_interval/fps:.1f}s intervals) …")

    face_cascade = cv2.CascadeClassifier(
        cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')

    best_score     = -1.0
    best_frame     = None
    best_frame_idx = 0
    frames_checked = 0

    fi = 0
    while True:
        ret, frame = cap.read()
        if not ret:
            break
        if fi % sample_interval == 0:
            score = score_frame(frame, face_cascade)
            frames_checked += 1
            if score > best_score:
                best_score     = score
                best_frame     = frame.copy()
                best_frame_idx = fi
            if frames_checked % 20 == 0:
                print(f"          [{fi+1}/{total}] frames scanned, "
                      f"best score: {best_score:.2f}  ", end='\r')
        fi += 1

    cap.release()
    print(f"\n          Checked {frames_checked} candidate frames")
    print(f"          Best frame: #{best_frame_idx}  (score={best_score:.2f})")

    if best_frame is None:
        raise RuntimeError(
            "No face detected in any sampled frame. "
            "Try a smaller sample_interval or check the video input.")

    raw_path = str(out / "reference_frame.jpg")
    cv2.imwrite(raw_path, best_frame, [cv2.IMWRITE_JPEG_QUALITY, 95])
    print(f"          Raw reference frame → {raw_path}")
    return raw_path


# ── Background Removal ─────────────────────────────────────────────────────────

def remove_background_rembg(image_path: str, out_path: str) -> bool:
    """Try rembg ML-based background removal. Returns True on success."""
    try:
        from rembg import remove
    except ImportError:
        return False

    print("[Stage 2] Removing background with rembg (ML) …")
    with open(image_path, 'rb') as f:
        result = remove(f.read())
    with open(out_path, 'wb') as f:
        f.write(result)
    return True


def remove_background_chromakey(image_path: str, out_path: str):
    """Fallback: remove green screen using pipeline.apply_chromakey."""
    sys.path.insert(0, str(Path(__file__).parent))
    from pipeline import apply_chromakey, hex_to_bgr

    print("[Stage 2] Removing background with chroma key (fallback) …")
    frame   = cv2.imread(image_path)
    key_bgr = hex_to_bgr("#74a44b")      # matches DEFAULT_KEY_HEX in pipeline.py
    bgra    = apply_chromakey(frame, key_bgr, sim=60, smooth=12, spill=0.40)
    cv2.imwrite(out_path, bgra)


def remove_background(image_path: str, output_dir: str) -> str:
    """
    Remove background from the reference frame.
    Tries rembg first (better for hair/hand edges); falls back to chroma key.
    Returns path to the clean RGBA PNG.
    """
    out_path = str(Path(output_dir) / "reference_clean.png")

    if not remove_background_rembg(image_path, out_path):
        print("[Stage 2] rembg not available — falling back to chroma key")
        print("          Install rembg for cleaner edges: pip install rembg")
        remove_background_chromakey(image_path, out_path)

    size_kb = Path(out_path).stat().st_size / 1024
    print(f"          Clean reference ({size_kb:.0f} KB) → {out_path}")
    return out_path


# ── Stage 2 Runner ─────────────────────────────────────────────────────────────

def run_stage2(video_path: str, output_dir: str,
               sample_interval: int = 30) -> dict:
    """
    Run Stage 2 reference extraction.
    Returns manifest dict:
      raw_frame        — path to reference_frame.jpg
      clean_reference  — path to reference_clean.png (RGBA, no background)
    """
    raw_frame_path  = extract_reference_frame(video_path, output_dir, sample_interval)
    clean_ref_path  = remove_background(raw_frame_path, output_dir)

    manifest = {
        "raw_frame":       raw_frame_path,
        "clean_reference": clean_ref_path,
    }

    import json
    manifest_path = Path(output_dir) / "stage2_manifest.json"
    with open(manifest_path, 'w') as f:
        json.dump(manifest, f, indent=2)
    print(f"\n[Stage 2 Complete] Manifest → {manifest_path}")

    return manifest


# ── CLI ────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    interval = int(sys.argv[3]) if len(sys.argv) > 3 else 30
    result   = run_stage2(sys.argv[1], sys.argv[2], interval)
    print()
    for k, v in result.items():
        print(f"  {k}: {v}")
