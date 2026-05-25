#!/usr/bin/env python3
"""
PPT + Green Screen Compositor Pipeline

Supports three slide background modes:
  1. .pptx file     → converted to images via LibreOffice + pdf2image
  2. .mp4 slideshow → used as background video, content detected per frame
  3. images dir     → pre-exported PNG/JPG slide images

Usage:
  python pipeline.py <faculty_video> <slides_source> [options]

Examples:
  python pipeline.py faculty.mp4 slides.pptx -o output.mp4
  python pipeline.py faculty.mp4 slideshow.mp4 -o output.mp4
  python pipeline.py faculty.mp4 ./slide_images/ -o output.mp4
  python pipeline.py faculty.mp4 slides.pptx --timings timings.json -o output.mp4
"""

import cv2
import numpy as np
import argparse
import json
import subprocess
import sys
import re
import os
from pathlib import Path

# ── Defaults ───────────────────────────────────────────────────────────────────
DEFAULT_KEY_HEX   = "#74a44b"   # sampled from faculty green screen
DEFAULT_SIM       = 60          # covers uneven screen lighting
DEFAULT_SMOOTH    = 12
DEFAULT_SPILL     = 0.40
DEFAULT_SENS      = 0.04     # edge density threshold for content detection
CENTER_SCALE      = 0.80     # faculty height as fraction of canvas height (centered mode)
RIGHT_SCALE       = 0.45     # faculty height (right-side mode)
RIGHT_MARGIN      = -0.05    # offset from right edge (negative = professor extends slightly off-canvas right)
LERP_SPEED        = 0.06     # position interpolation speed per frame (~0.3s at 30fps)
DETECT_INTERVAL   = 15       # for video-bg mode: re-analyze every N frames


# ── Helpers ────────────────────────────────────────────────────────────────────
def hex_to_bgr(h: str):
    h = h.lstrip('#')
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return (b, g, r)

def lerp(a, b, t):
    return a + (b - a) * t

def natural_sort_key(s):
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', str(s))]

def find_libreoffice() -> str:
    """Return the LibreOffice binary path, checking macOS app bundle if needed."""
    import shutil
    # Standard PATH lookup (Linux / brew-linked)
    lo = shutil.which('libreoffice') or shutil.which('soffice')
    if lo:
        return lo
    # macOS .app bundle (installed via brew cask or direct download)
    mac_path = '/Applications/LibreOffice.app/Contents/MacOS/soffice'
    if Path(mac_path).exists():
        return mac_path
    return None


# ── Input Type Detection ───────────────────────────────────────────────────────
def detect_input_type(path: str) -> str:
    p = Path(path)
    if p.suffix.lower() == '.pptx':
        return 'pptx'
    if p.suffix.lower() in {'.mp4', '.mov', '.avi', '.mkv', '.webm'}:
        return 'video'
    if p.is_dir():
        return 'images'
    raise ValueError(f"Unrecognised slides source: {path}")


# ── PPTX → Images ─────────────────────────────────────────────────────────────
def pptx_to_images(pptx_path: str, tmp_dir: str) -> list:
    """
    Convert PPTX → PNG images via LibreOffice (PPTX→PDF) then PyMuPDF (PDF→PNG).
    Two-step is necessary because LibreOffice's direct --convert-to png only
    outputs the first slide for multi-slide presentations on some platforms.
    """
    Path(tmp_dir).mkdir(parents=True, exist_ok=True)

    # Skip conversion if slides are already cached from a previous run
    cached = sorted(Path(tmp_dir).glob('slide_*.png'), key=natural_sort_key)
    if cached:
        print(f"[1/4] Using {len(cached)} cached slide images from {tmp_dir}")
        print(f"[2/4] {len(cached)} slide images ready (cached)")
        return [str(p) for p in cached]

    lo_bin = find_libreoffice()
    if not lo_bin:
        print("[ERROR] LibreOffice not found.")
        print("        Install: brew install --cask libreoffice")
        sys.exit(1)

    # Step 1: PPTX → PDF (reliably captures all slides)
    print(f"[1/4] Converting PPTX → PDF via LibreOffice …")
    print(f"      Using: {lo_bin}")
    result = subprocess.run(
        [lo_bin, '--headless', '--convert-to', 'pdf',
         '--outdir', tmp_dir, pptx_path],
        capture_output=True, text=True, timeout=120
    )
    if result.returncode != 0:
        print(f"[ERROR] LibreOffice failed:\n{result.stderr}")
        sys.exit(1)

    pdfs = list(Path(tmp_dir).glob('*.pdf'))
    if not pdfs:
        print(f"[ERROR] No PDF produced. stdout: {result.stdout}")
        sys.exit(1)

    # Step 2: PDF → one PNG per page via PyMuPDF (pure Python, no system deps)
    try:
        import fitz  # pymupdf
    except ImportError:
        print("[ERROR] pymupdf not installed. Run: pip install pymupdf")
        sys.exit(1)

    # Suppress MuPDF's C-level warnings (PDF structure/tag issues that don't affect rendering)
    _dn = os.open(os.devnull, os.O_WRONLY); _se = os.dup(2); os.dup2(_dn, 2)
    try:
        pdf = fitz.open(str(pdfs[0]))
    finally:
        os.dup2(_se, 2); os.close(_dn); os.close(_se)

    mat  = fitz.Matrix(2.0, 2.0)   # 2× scale ≈ 144 DPI
    pngs = []
    for i, page in enumerate(pdf):
        pix      = page.get_pixmap(matrix=mat)
        out_path = Path(tmp_dir) / f"slide_{i+1:04d}.png"
        pix.save(str(out_path))
        pngs.append(str(out_path))
    pdf.close()

    if not pngs:
        print("[ERROR] No slides extracted from PDF.")
        sys.exit(1)

    print(f"[2/4] {len(pngs)} slide images ready")
    return pngs


# ── Load Image Directory ───────────────────────────────────────────────────────
def load_images_from_dir(directory: str) -> list:
    exts = {'.png', '.jpg', '.jpeg', '.bmp', '.tiff'}
    files = [p for p in Path(directory).iterdir() if p.suffix.lower() in exts]
    files.sort(key=natural_sort_key)
    imgs = []
    for f in files:
        img = cv2.imread(str(f))
        if img is not None:
            imgs.append(img)
    print(f"      Loaded {len(imgs)} images from {directory}")
    return imgs


# ── Content Detection ──────────────────────────────────────────────────────────
def detect_content(img: np.ndarray, threshold: float = DEFAULT_SENS) -> str:
    """
    Returns 'blank' | 'left' | 'full' | 'photo' based on frame content.
    'photo'  → photographic/room background (low white-pixel fraction) — not a slide
    'blank'  → slide with no text/diagram content → professor centered
    'left'   → content on left half only → professor on right
    'full'   → content fills right side too → professor hidden
    """
    H, W = img.shape[:2]
    crop = img[int(H*0.10):int(H*0.90), int(W*0.10):int(W*0.90)]
    gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)
    _, mid = gray.shape
    mid = mid // 2

    def density(region):
        edges = cv2.Canny(region, 100, 200)
        return np.count_nonzero(edges) / max(edges.size, 1)

    left_d  = density(gray[:, :mid])
    right_d = density(gray[:, mid:])

    # Photo-background detection: edge-sparse AND symmetric (no text/diagram asymmetry).
    # Threshold 0.55 catches both dark campus photos (low whiteFraction) and bright
    # corridor photos (white ceiling/walls push fraction up but still < 0.55).
    # Pure slide backgrounds are ≥ 0.55 white and fall through to blank/left/full.
    white_fraction = np.all(crop > 200, axis=2).mean()
    if white_fraction < 0.55:
        avg_density = (left_d + right_d) / 2
        asymmetry   = abs(left_d - right_d)
        if avg_density < 0.04 and asymmetry < 0.02:
            return 'photo'
        # else: text/graphic creates asymmetry or high density → fall through

    has_left  = left_d  > threshold
    has_right = right_d > threshold

    if not has_left and not has_right:
        return 'blank'
    if has_left and not has_right:
        return 'left'
    return 'full'


# ── Person Detection ───────────────────────────────────────────────────────────
_face_cascade = None

def _get_face_cascade():
    global _face_cascade
    if _face_cascade is None:
        _face_cascade = cv2.CascadeClassifier(
            cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
    return _face_cascade

def detect_person_box(img: np.ndarray) -> dict:
    """
    Returns normalised {x,y,w,h} bounding box for the largest person found in img,
    or None if no person is detected or the detected region is too small to be
    meaningful (< 15 % of slide height — avoids keying on tiny portrait photos).
    Strategy: face cascade → expand to body; fall back to HOG full-person detector.
    """
    h, w = img.shape[:2]
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    faces = _get_face_cascade().detectMultiScale(
        gray, scaleFactor=1.1, minNeighbors=4, minSize=(50, 50))

    if len(faces) > 0:
        fx, fy, fw, fh = max(faces, key=lambda f: f[2] * f[3])
        bw = min(int(fw * 2.5), w)
        bh = min(int(fh * 5.0), h)
        bx = max(0, fx - int((bw - fw) / 2))
        by = fy
        bw = min(bw, w - bx)
        bh = min(bh, h - by)
        # Require the person region to be at least 25% of slide height —
        # stricter than 15% to reject small stock-photo faces in slide templates.
        if bh / h >= 0.25:
            return {'x': bx / w, 'y': by / h, 'w': bw / w, 'h': bh / h}

    hog = cv2.HOGDescriptor()
    hog.setSVMDetector(cv2.HOGDescriptor_getDefaultPeopleDetector())
    regions, _ = hog.detectMultiScale(img, winStride=(8, 8), padding=(4, 4), scale=1.05)
    if len(regions) > 0:
        px, py, pw, ph = max(regions, key=lambda r: r[2] * r[3])
        if ph / h >= 0.25:
            return {'x': px / w, 'y': py / h, 'w': pw / w, 'h': ph / h}

    return None


# ── Chroma Key (vectorised NumPy) ─────────────────────────────────────────────
def apply_chromakey(frame: np.ndarray, key_bgr,
                    sim: int, smooth: int, spill: float) -> np.ndarray:
    """Returns a BGRA frame with the key colour replaced by transparency."""
    f = frame.astype(np.float32)
    kb, kg, kr = float(key_bgr[0]), float(key_bgr[1]), float(key_bgr[2])

    dist = np.sqrt(
        (f[:,:,0] - kb)**2 +
        (f[:,:,1] - kg)**2 +
        (f[:,:,2] - kr)**2
    )

    alpha = np.where(dist < sim, 0.0,
            np.where(dist < sim + smooth, (dist - sim) / max(smooth, 1), 1.0))

    # Extended soft zone: key-hue pixels just outside the smooth threshold
    # (handles floor shadows / dark green-screen corners that slip past the main key)
    sim_soft  = sim * 1.8
    ext_range = max(sim_soft - sim - smooth, 1.0)
    if kg >= kr and kg >= kb:
        key_hue = (f[:,:,1] >= f[:,:,0]) & (f[:,:,1] >= f[:,:,2])
    elif kb >= kr and kb >= kg:
        key_hue = (f[:,:,2] >= f[:,:,0]) & (f[:,:,2] >= f[:,:,1])
    else:
        key_hue = (f[:,:,0] >= f[:,:,1]) & (f[:,:,0] >= f[:,:,2])
    in_ext = (dist >= sim + smooth) & (dist < sim_soft) & key_hue
    alpha   = np.where(in_ext, np.clip((dist - (sim + smooth)) / ext_range, 0.0, 1.0), alpha)

    if spill > 0:
        b_ch, g_ch, r_ch = f[:,:,0], f[:,:,1], f[:,:,2]
        sf = (1.0 - alpha) * spill
        if kg >= kr and kg >= kb:
            f[:,:,1] = np.clip(g_ch - (g_ch - np.maximum(r_ch, b_ch)) * sf, 0, 255)
        elif kb >= kr and kb >= kg:
            f[:,:,0] = np.clip(b_ch - (b_ch - np.maximum(r_ch, g_ch)) * sf, 0, 255)
        else:
            f[:,:,2] = np.clip(r_ch - (r_ch - np.maximum(g_ch, b_ch)) * sf, 0, 255)

    bgra = cv2.cvtColor(f.astype(np.uint8), cv2.COLOR_BGR2BGRA)
    bgra[:,:,3] = (alpha * 255).astype(np.uint8)
    return bgra


# ── Alpha Composite ────────────────────────────────────────────────────────────
def composite(bg: np.ndarray, fg_bgra: np.ndarray,
              x: float, y: float, w: float, h: float) -> np.ndarray:
    """Alpha-composite fg_bgra over bg at floating-point position (x,y,w,h)."""
    out = bg.copy()
    xi, yi, wi, hi = int(round(x)), int(round(y)), max(1, int(round(w))), max(1, int(round(h)))
    BH, BW = bg.shape[:2]

    x0 = max(0, xi);     y0 = max(0, yi)
    x1 = min(BW, xi+wi); y1 = min(BH, yi+hi)
    if x1 <= x0 or y1 <= y0:
        return out

    fg_r = cv2.resize(fg_bgra, (wi, hi))
    fg_crop = fg_r[y0-yi : y1-yi, x0-xi : x1-xi]

    a = fg_crop[:,:,3:4].astype(np.float32) / 255.0
    fg_rgb = fg_crop[:,:,:3].astype(np.float32)
    bg_rgn = out[y0:y1, x0:x1].astype(np.float32)
    out[y0:y1, x0:x1] = (fg_rgb * a + bg_rgn * (1.0 - a)).astype(np.uint8)
    return out


# ── Target Position ────────────────────────────────────────────────────────────
def get_target(content_state: str, OW: int, OH: int, fac_aspect: float,
               center_scale=CENTER_SCALE, right_scale=RIGHT_SCALE,
               margin=RIGHT_MARGIN, person_box=None):
    """Returns (x, y, w, h, visible).  visible=False → skip compositing."""
    if content_state == 'person' and person_box:
        # Replace the detected person with the speaker
        pbh = person_box['h'] * OH
        pbw = person_box['w'] * OW
        pbx = person_box['x'] * OW
        pby = person_box['y'] * OH
        h = pbh
        w = h * fac_aspect
        x = pbx + (pbw - w) / 2.0  # centre within the person box
        y = pby
        return x, y, w, h, True
    elif content_state == 'full':
        # Full-screen content: hide the professor but keep same size/position as center
        # so the opacity lerp doesn't also jerk the position.
        h = OH * center_scale
        w = h * fac_aspect
        x = (OW - w) / 2.0
        y = (OH - h) / 2.0
        return x, y, w, h, False
    elif content_state == 'left':
        # Content on left half → professor on right, same size as centered (no shrink).
        h = OH * center_scale
        w = h * fac_aspect
        x = OW - w - OW * margin
        y = (OH - h) / 2.0   # vertically centered, same as center mode
        return x, y, w, h, True
    else:  # 'blank' or 'photo'
        h = OH * center_scale
        w = h * fac_aspect
        x = (OW - w) / 2.0
        y = (OH - h) / 2.0
        return x, y, w, h, True


# ── Timings ────────────────────────────────────────────────────────────────────
def build_equal_timings(num_slides: int, total_secs: float) -> list:
    d = total_secs / num_slides
    return [{'slide': i, 'start': i*d, 'end': (i+1)*d} for i in range(num_slides)]

def slide_at(t: float, timings: list) -> int:
    for entry in reversed(timings):
        if t >= entry['start']:
            return int(entry['slide'])
    return 0


# ── Merge Audio ────────────────────────────────────────────────────────────────
def merge_audio(video_no_audio: str, audio_source: str, output: str):
    """Use FFmpeg to attach audio from faculty video onto the composited video."""
    tmp = video_no_audio + "_tmp.mp4"
    os.rename(video_no_audio, tmp)
    result = subprocess.run([
        'ffmpeg', '-y',
        '-i', tmp,
        '-i', audio_source,
        '-c:v', 'copy',
        '-c:a', 'aac', '-b:a', '192k',
        '-map', '0:v:0', '-map', '1:a:0',
        '-shortest', output
    ], capture_output=True, text=True)
    if result.returncode == 0:
        os.remove(tmp)
    else:
        os.rename(tmp, output)   # save video without audio rather than losing it
        print(f"[WARN] Audio merge failed (FFmpeg): {result.stderr.strip()}")
        print(f"       Video saved without audio at: {output}")


# ── Pre-roll Renderer ─────────────────────────────────────────────────────────
def generate_preroll(slide_images: list, timings: list, end_time: float,
                     fps: float, OW: int, OH: int, output_path: str) -> bool:
    """Render a silent slideshow for slides that appear before end_time.
    Used as the title/intro segment before the faculty composited video begins."""
    pre_timings = [e for e in timings if e['start'] < end_time]
    if not pre_timings:
        return False

    pre_total = max(1, int(end_time * fps))
    last_slide = max(int(e['slide']) + 1 for e in pre_timings)
    print(f"      Pre-roll: slides 1–{last_slide}, {end_time:.1f}s, {pre_total} frames")

    ffcmd = [
        'ffmpeg', '-y',
        '-f', 'rawvideo', '-vcodec', 'rawvideo',
        '-s', f'{OW}x{OH}', '-pix_fmt', 'bgr24', '-r', str(fps),
        '-i', 'pipe:0',
        '-f', 'lavfi', '-i', 'anullsrc=r=44100:cl=stereo',
        '-c:v', 'libx264', '-preset', 'fast', '-crf', '23', '-pix_fmt', 'yuv420p',
        '-c:a', 'aac', '-b:a', '192k',
        '-map', '0:v:0', '-map', '1:a:0', '-shortest',
        '-movflags', '+faststart',
        output_path
    ]
    ffproc = subprocess.Popen(ffcmd, stdin=subprocess.PIPE, stderr=subprocess.DEVNULL)

    XFADE  = int(fps * 0.5)
    n      = len(slide_images)
    prev_s = -1
    prev_b = None
    xrem   = 0

    for fi in range(pre_total):
        t    = fi / fps
        sidx = min(slide_at(t, pre_timings), n - 1)
        curr = cv2.resize(slide_images[sidx], (OW, OH))

        if sidx != prev_s:
            if prev_s >= 0:
                prev_b = cv2.resize(slide_images[prev_s], (OW, OH))
                xrem   = XFADE
            prev_s = sidx

        if prev_b is not None and xrem > 0:
            alpha  = 1.0 - xrem / XFADE
            frame  = cv2.addWeighted(prev_b, 1.0 - alpha, curr, alpha, 0)
            xrem  -= 1
            if xrem == 0:
                prev_b = None
        else:
            frame = curr

        ffproc.stdin.write(frame.tobytes())
        if (fi + 1) % 150 == 0 or fi == pre_total - 1:
            print(f"      Pre-roll {fi+1}/{pre_total} ({(fi+1)/pre_total*100:.0f}%)  ", end='\r')

    print()
    ffproc.stdin.close()
    return ffproc.wait() == 0


# ── Slides Pipeline (PPTX or images dir) ──────────────────────────────────────
def run_slides_pipeline(
        faculty_video: str, slide_images: list, output: str,
        timings_file: str, key_bgr, sim, smooth, spill,
        sensitivity, lerp_speed, center_scale, right_scale,
        start_slide: int = 0):

    # Content analysis
    print("[3/4] Analysing slide content …")
    content_map  = [detect_content(s, sensitivity) for s in slide_images]
    person_boxes = [None] * len(slide_images)   # REPLACE mode disabled for PPTX slides
    left_count   = sum(1 for s in content_map if s == 'left')
    full_count   = sum(1 for s in content_map if s == 'full')
    photo_count  = sum(1 for s in content_map if s == 'photo')
    blank_count  = len(content_map) - left_count - full_count - photo_count
    print(f"      {left_count} left-content, {full_count} full-screen, "
          f"{photo_count} photo-bg, {blank_count} blank")

    # Open faculty video
    cap = cv2.VideoCapture(faculty_video)
    if not cap.isOpened():
        print(f"[ERROR] Cannot open faculty video: {faculty_video}")
        sys.exit(1)

    fps    = cap.get(cv2.CAP_PROP_FPS) or 30.0
    total  = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    dur    = total / fps
    fac_w  = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
    fac_h  = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
    aspect = fac_w / fac_h
    OW, OH = fac_w, fac_h
    print(f"      Faculty: {fac_w}×{fac_h} @ {fps:.1f}fps, {dur:.1f}s, {total} frames")

    # Timings
    if timings_file:
        with open(timings_file) as f:
            timings = json.load(f)
        print(f"      Using timings from {timings_file}")
    else:
        timings = build_equal_timings(len(slide_images), dur)
        print(f"      Auto-timing: {dur/len(slide_images):.1f}s per slide")
        if start_slide > 1:
            print(f"      [WARN] No timings file loaded — trim will use equal-time estimates.")
            print(f"             Run ⟳ Sync PPTX first for accurate slide timing.")

    print(f"      start_slide={start_slide}  (1 = full video, >1 = trim + pre-roll)")

    # Trim: seek faculty video to the start of the requested slide.
    # When start_slide > 1, slides before it become a silent pre-roll segment.
    original_timings = list(timings)   # keep for pre-roll generation
    trim_start  = 0.0
    need_preroll = False
    if start_slide > 1:
        slide_idx_0 = start_slide - 1  # convert to 0-indexed
        # Look for the first timing entry at or after the requested slide —
        # exact match first, then nearest >= so a missing slide doesn't block the trim.
        sorted_t = sorted(timings, key=lambda e: e['start'])
        match = next((e['start'] for e in sorted_t if int(e['slide']) == slide_idx_0), None)
        if match is None:
            match = next((e['start'] for e in sorted_t if int(e['slide']) >= slide_idx_0), None)
        if match is not None and match > 0:
            trim_start   = match
            need_preroll = True
            frame_offset = int(trim_start * fps)
            print(f"      Trimming faculty to slide {start_slide}: t={trim_start:.2f}s "
                  f"(frame {frame_offset} of {total})")
            print(f"      Pre-roll will cover slides 1–{start_slide - 1}")
            timings = [
                {'slide': e['slide'],
                 'start': round(e['start'] - trim_start, 3),
                 'end':   round(e['end']   - trim_start, 3)}
                for e in timings if e['end'] > trim_start
            ]
            # Seek using frame index — more reliable than MSEC for most codecs
            ok = cap.set(cv2.CAP_PROP_POS_FRAMES, frame_offset)
            actual_frame = int(cap.get(cv2.CAP_PROP_POS_FRAMES))
            if abs(actual_frame - frame_offset) > fps * 2:
                print(f"      [WARN] Frame seek inaccurate (wanted {frame_offset}, at {actual_frame})")
                print(f"             Trying MSEC seek as fallback …")
                cap.set(cv2.CAP_PROP_POS_MSEC, trim_start * 1000)
            total = max(1, int((dur - trim_start) * fps))
            print(f"      Main segment: {total} frames (~{total/fps:.1f}s)")
        else:
            print(f"      [WARN] Could not find start time for slide {start_slide} in timings "
                  f"— no trim applied. Run ⟳ Sync PPTX to generate timings first.")

    # Always write to a temp file so an interrupted encode never corrupts the
    # final output.  Rename to the real path only on clean ffmpeg exit.
    main_target = str(Path(output).with_suffix('')) + '_main_tmp.mp4'
    audio_seek  = ['-ss', str(round(trim_start, 3))] if trim_start > 0 else []
    ffcmd = [
        'ffmpeg', '-y',
        '-f', 'rawvideo', '-vcodec', 'rawvideo',
        '-s', f'{OW}x{OH}', '-pix_fmt', 'bgr24', '-r', str(fps),
        '-i', 'pipe:0',
        *audio_seek, '-i', faculty_video,
        '-c:v', 'libx264', '-preset', 'fast', '-crf', '23', '-pix_fmt', 'yuv420p',
        '-c:a', 'aac', '-b:a', '192k',
        '-map', '0:v:0', '-map', '1:a:0', '-shortest',
        '-movflags', '+faststart',
        main_target
    ]
    ffproc = subprocess.Popen(ffcmd, stdin=subprocess.PIPE, stderr=subprocess.DEVNULL)

    XFADE_FRAMES = int(fps * 0.5)   # 0.5-second slide image crossfade
    ALPHA_SMOOTH = 0.4              # temporal alpha smoothing: fraction of previous frame alpha to retain

    # Scale lerp_speed from the browser's 60fps animation rate to the video fps so
    # position and opacity transitions take the same wall-clock time as the preview.
    per_frame_lerp = 1.0 - (1.0 - lerp_speed) ** (60.0 / fps)

    cur_x = cur_y = cur_w = cur_h = 0.0
    cur_opacity = 1.0
    first = True
    prev_sidx   = -1
    prev_bg     = None   # previous slide image (resized) — used for crossfade
    xfade_rem   = 0      # frames remaining in the current slide image crossfade
    prev_alpha  = None   # previous frame's alpha channel for temporal smoothing

    print(f"[4/4] Compositing {total} frames …")
    for fi in range(total):
        ret, fac_frame = cap.read()
        if not ret:
            break

        t = fi / fps
        sidx = min(slide_at(t, timings), len(slide_images) - 1)
        tgt_x, tgt_y, tgt_w, tgt_h, visible = get_target(
            content_map[sidx], OW, OH, aspect, center_scale, right_scale,
            person_box=person_boxes[sidx])
        tgt_opacity = 1.0 if visible else 0.0

        # On slide change: start the slide image crossfade.
        if first or sidx != prev_sidx:
            if not first:
                prev_bg   = cv2.resize(slide_images[prev_sidx], (OW, OH))
                xfade_rem = XFADE_FRAMES
            else:
                # First frame: snap position and opacity immediately (nothing to lerp from).
                cur_x, cur_y, cur_w, cur_h = tgt_x, tgt_y, tgt_w, tgt_h
                cur_opacity = tgt_opacity
            first     = False
            prev_sidx = sidx

        # Lerp position and opacity toward target every frame — matches browser behaviour.
        # Uses per_frame_lerp (fps-adjusted) so transitions take the same time as preview.
        if fi > 0:
            cur_x = lerp(cur_x, tgt_x, per_frame_lerp)
            cur_y = lerp(cur_y, tgt_y, per_frame_lerp)
            cur_w = lerp(cur_w, tgt_w, per_frame_lerp)
            cur_h = lerp(cur_h, tgt_h, per_frame_lerp)
            cur_opacity = lerp(cur_opacity, tgt_opacity, per_frame_lerp)

        # Build current background
        curr_bg = cv2.resize(slide_images[sidx], (OW, OH))

        # Crossfade: blend previous slide out and current slide in
        if prev_bg is not None and xfade_rem > 0:
            alpha = 1.0 - xfade_rem / XFADE_FRAMES
            bg = cv2.addWeighted(prev_bg, 1.0 - alpha, curr_bg, alpha, 0)
            xfade_rem -= 1
            if xfade_rem == 0:
                prev_bg = None
        else:
            bg = curr_bg

        fg_bgra = apply_chromakey(fac_frame, key_bgr, sim, smooth, spill)

        # Temporal alpha smoothing: blend with previous frame to suppress flicker
        # caused by per-frame compression noise near the green screen boundary.
        new_alpha = fg_bgra[:, :, 3].astype(np.float32)
        if prev_alpha is not None:
            smoothed = prev_alpha * ALPHA_SMOOTH + new_alpha * (1.0 - ALPHA_SMOOTH)
            fg_bgra[:, :, 3] = smoothed.astype(np.uint8)
            prev_alpha = smoothed
        else:
            prev_alpha = new_alpha

        if cur_opacity > 0.01:
            if cur_opacity < 0.99:
                fg_bgra = fg_bgra.copy()
                fg_bgra[:, :, 3] = (fg_bgra[:, :, 3].astype(np.float32) * cur_opacity).astype(np.uint8)
            out_frame = composite(bg, fg_bgra, cur_x, cur_y, cur_w, cur_h)
        else:
            out_frame = bg

        ffproc.stdin.write(out_frame.tobytes())

        if (fi + 1) % 100 == 0 or fi == total - 1:
            mode = content_map[sidx].upper()
            print(f"      {fi+1}/{total} ({(fi+1)/total*100:.0f}%) — Slide {sidx+1} [{mode}]  ", end='\r')

    print()
    cap.release()
    ffproc.stdin.close()
    if ffproc.wait() != 0:
        print(f"[WARN] FFmpeg encoding failed — output may be incomplete")
        try: os.remove(main_target)
        except Exception: pass
    else:
        print(f"[DONE] Main segment → {main_target}")

    # ── Pre-roll concat ────────────────────────────────────────────────────────
    if not need_preroll:
        if Path(main_target).exists():
            os.replace(main_target, output)
            print(f"[DONE] {output}")
        return

    if need_preroll and Path(main_target).exists():
        preroll_path = str(Path(output).with_suffix('')) + '_preroll_tmp.mp4'
        print(f"[4/4] Rendering pre-roll (slides 1–{start_slide - 1}) …")
        ok = generate_preroll(slide_images, original_timings, trim_start,
                              fps, OW, OH, preroll_path)
        if ok and Path(preroll_path).exists():
            print(f"      Concatenating pre-roll + main …")
            # Use filter_complex concat instead of demuxer -c copy:
            # filter_complex re-encodes and handles PTS discontinuities correctly,
            # preventing the output from being truncated when timestamps don't align.
            result = subprocess.run([
                'ffmpeg', '-y',
                '-i', preroll_path,
                '-i', main_target,
                '-filter_complex',
                '[0:v:0][0:a:0][1:v:0][1:a:0]concat=n=2:v=1:a=1[v][a]',
                '-map', '[v]', '-map', '[a]',
                '-c:v', 'libx264', '-preset', 'fast', '-crf', '23', '-pix_fmt', 'yuv420p',
                '-c:a', 'aac', '-b:a', '192k',
                '-movflags', '+faststart',
                output
            ], capture_output=True, text=True)
            if result.returncode == 0:
                print(f"[DONE] {output}  (pre-roll + main concatenated)")
                for p in (preroll_path, main_target):
                    try:
                        os.remove(p)
                    except Exception:
                        pass
            else:
                print(f"[WARN] Concat failed: {result.stderr.strip()}")
                print(f"       Falling back to main segment only")
                try:
                    os.remove(preroll_path)
                except Exception:
                    pass
                if Path(main_target).exists():
                    os.rename(main_target, output)
        else:
            print(f"[WARN] Pre-roll generation failed — keeping main segment only")
            if Path(main_target).exists():
                os.rename(main_target, output)


# ── Video-BG Pipeline (MP4 slideshow as background) ───────────────────────────
def run_video_pipeline(
        faculty_video: str, bg_video: str, output: str,
        key_bgr, sim, smooth, spill,
        sensitivity, lerp_speed, center_scale, right_scale,
        sync_timings_file: str = None):

    fac_cap = cv2.VideoCapture(faculty_video)
    bg_cap  = cv2.VideoCapture(bg_video)

    if not fac_cap.isOpened():
        print(f"[ERROR] Cannot open faculty video: {faculty_video}")
        sys.exit(1)
    if not bg_cap.isOpened():
        print(f"[ERROR] Cannot open background video: {bg_video}")
        sys.exit(1)

    fps    = fac_cap.get(cv2.CAP_PROP_FPS) or 30.0
    total  = int(fac_cap.get(cv2.CAP_PROP_FRAME_COUNT))
    fac_w  = int(fac_cap.get(cv2.CAP_PROP_FRAME_WIDTH))
    fac_h  = int(fac_cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
    aspect = fac_w / fac_h
    OW, OH = fac_w, fac_h

    # Derive true bg fps — cv2 may report a timebase (e.g. 600/1) instead of playback fps
    _bg_fps_raw = bg_cap.get(cv2.CAP_PROP_FPS) or 30.0
    _bg_total   = int(bg_cap.get(cv2.CAP_PROP_FRAME_COUNT))
    if _bg_fps_raw > 120 or _bg_fps_raw < 1:
        _dur_r = subprocess.run(
            ['ffprobe', '-v', 'error', '-select_streams', 'v:0',
             '-show_entries', 'stream=duration',
             '-of', 'default=noprint_wrappers=1:nokey=1', bg_video],
            capture_output=True, text=True)
        try:
            bg_fps = _bg_total / float(_dur_r.stdout.strip())
        except Exception:
            bg_fps = 30.0
    else:
        bg_fps = _bg_fps_raw

    print(f"      Faculty : {fac_w}×{fac_h} @ {fps:.1f}fps, {total} frames")
    print(f"      BG video: {int(bg_cap.get(cv2.CAP_PROP_FRAME_WIDTH))}×{int(bg_cap.get(cv2.CAP_PROP_FRAME_HEIGHT))} @ {bg_fps:.1f}fps")

    ffcmd = [
        'ffmpeg', '-y',
        '-f', 'rawvideo', '-vcodec', 'rawvideo',
        '-s', f'{OW}x{OH}', '-pix_fmt', 'bgr24', '-r', str(fps),
        '-i', 'pipe:0',
        '-i', faculty_video,
        '-c:v', 'libx264', '-preset', 'fast', '-crf', '23', '-pix_fmt', 'yuv420p',
        '-c:a', 'aac', '-b:a', '192k',
        '-map', '0:v:0', '-map', '1:a:0', '-shortest',
        '-movflags', '+faststart',
        output
    ]
    ffproc = subprocess.Popen(ffcmd, stdin=subprocess.PIPE, stderr=subprocess.DEVNULL)

    # Load sync timings if provided
    sync_entries = []
    if sync_timings_file and Path(sync_timings_file).exists():
        with open(sync_timings_file) as _sf:
            sync_entries = json.load(_sf)
        print(f"      Sync timings: {len(sync_entries)} segments loaded")

    cur_x = cur_y = cur_w = cur_h = 0.0
    cur_opacity = 1.0
    first = True
    cached_state = 'blank'
    cached_person_box = None
    last_vsb_start = None   # tracks which sync segment is active
    bg_accum = 0.0          # fractional bg frames to advance per output frame
    bg_frame = None         # last successfully read bg frame
    prev_alpha  = None      # previous frame alpha for temporal smoothing
    ALPHA_SMOOTH = 0.4

    print(f"[4/4] Compositing {total} frames (video-bg mode) …")
    for fi in range(total):
        fac_t = fi / fps

        # Seek bg video when entering a new sync segment
        if sync_entries:
            cur_vsb_start = None
            for e in reversed(sync_entries):
                if fac_t >= e['fac_start']:
                    cur_vsb_start = e['vsb_start']
                    break
            if cur_vsb_start is not None and cur_vsb_start != last_vsb_start:
                bg_cap.set(cv2.CAP_PROP_POS_MSEC, cur_vsb_start * 1000)
                last_vsb_start = cur_vsb_start
                bg_accum = 0.0  # reset after seek so no phantom frames are skipped

        ret_f, fac_frame = fac_cap.read()
        if not ret_f:
            break

        # Advance bg by bg_fps/fps frames per output frame to handle fps mismatch
        bg_accum += bg_fps / fps
        while bg_accum >= 1.0:
            ret_b, tmp = bg_cap.read()
            if ret_b:
                bg_frame = tmp
            else:
                if not sync_entries:
                    bg_cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
                    ret_b, tmp = bg_cap.read()
                    if ret_b:
                        bg_frame = tmp
                break
            bg_accum -= 1.0

        if bg_frame is None:
            bg_frame = np.zeros((OH, OW, 3), dtype=np.uint8)

        if fi % DETECT_INTERVAL == 0:
            state = detect_content(bg_frame, sensitivity)
            if state == 'photo':
                box = detect_person_box(bg_frame)
                if box:
                    cached_state = 'person'
                    cached_person_box = box
                else:
                    cached_state = 'blank'
                    cached_person_box = None
            else:
                cached_state = state
                cached_person_box = None

        tgt_x, tgt_y, tgt_w, tgt_h, visible = get_target(
            cached_state, OW, OH, aspect, center_scale, right_scale,
            person_box=cached_person_box)
        tgt_opacity = 1.0 if visible else 0.0

        if first:
            cur_x, cur_y, cur_w, cur_h = tgt_x, tgt_y, tgt_w, tgt_h
            cur_opacity = tgt_opacity
            first = False
        else:
            cur_x = lerp(cur_x, tgt_x, lerp_speed)
            cur_y = lerp(cur_y, tgt_y, lerp_speed)
            cur_w = lerp(cur_w, tgt_w, lerp_speed)
            cur_h = lerp(cur_h, tgt_h, lerp_speed)
            cur_opacity = lerp(cur_opacity, tgt_opacity, lerp_speed)

        bg      = cv2.resize(bg_frame, (OW, OH))
        fg_bgra = apply_chromakey(fac_frame, key_bgr, sim, smooth, spill)

        new_alpha = fg_bgra[:, :, 3].astype(np.float32)
        if prev_alpha is not None:
            smoothed = prev_alpha * ALPHA_SMOOTH + new_alpha * (1.0 - ALPHA_SMOOTH)
            fg_bgra[:, :, 3] = smoothed.astype(np.uint8)
            prev_alpha = smoothed
        else:
            prev_alpha = new_alpha

        if cached_state == 'person' and cached_person_box:
            pb = cached_person_box
            x0 = max(0, int(pb['x'] * OW));  y0 = max(0, int(pb['y'] * OH))
            x1 = min(OW, int((pb['x'] + pb['w']) * OW))
            y1 = min(OH, int((pb['y'] + pb['h']) * OH))
            bg[y0:y1, x0:x1] = (bg[y0:y1, x0:x1].astype(np.float32) * 0.35).astype(np.uint8)

        if cur_opacity > 0.01:
            if cur_opacity < 0.99:
                fg_bgra = fg_bgra.copy()
                fg_bgra[:, :, 3] = (fg_bgra[:, :, 3].astype(np.float32) * cur_opacity).astype(np.uint8)
            out_frame = composite(bg, fg_bgra, cur_x, cur_y, cur_w, cur_h)
        else:
            out_frame = bg

        ffproc.stdin.write(out_frame.tobytes())

        if (fi + 1) % 100 == 0 or fi == total - 1:
            print(f"      {fi+1}/{total} ({(fi+1)/total*100:.0f}%) [{cached_state.upper()}]  ", end='\r')

    print()
    fac_cap.release()
    bg_cap.release()
    ffproc.stdin.close()
    if ffproc.wait() != 0:
        print(f"[WARN] FFmpeg encoding failed — output may be incomplete")
    else:
        print(f"[DONE] {output}")


# ── Entry Point ────────────────────────────────────────────────────────────────
def main():
    global DETECT_INTERVAL
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument('faculty_video',
        help='Green-screen faculty video (.mp4 / .mov)')
    parser.add_argument('slides_source',
        help='.pptx file  |  slideshow .mp4  |  directory of slide images')
    parser.add_argument('-o', '--output',     default='output.mp4')
    parser.add_argument('-t', '--timings',    default=None,
        help='JSON timing file for PPTX/images mode (optional — auto if omitted)')
    parser.add_argument('--sync',             default=None,
        help='Sync timings JSON from speech_sync.py (video-bg mode only)')
    parser.add_argument('--key',              default=DEFAULT_KEY_HEX,
        help='Chroma key hex colour  (default: #00b140 — green)')
    parser.add_argument('--sim',    type=int,   default=DEFAULT_SIM,
        help='Key similarity radius  (default: 35)')
    parser.add_argument('--smooth', type=int,   default=DEFAULT_SMOOTH,
        help='Key edge smoothness    (default: 8)')
    parser.add_argument('--spill',  type=float, default=DEFAULT_SPILL,
        help='Spill suppression 0-1  (default: 0.30)')
    parser.add_argument('--sensitivity', type=float, default=DEFAULT_SENS,
        help='Content-detection edge-density threshold  (default: 0.05)')
    parser.add_argument('--lerp-speed',  type=float, default=LERP_SPEED,
        help='Position transition speed per frame  (default: 0.06)')
    parser.add_argument('--center-scale',type=float, default=CENTER_SCALE,
        help='Faculty height fraction when centered  (default: 0.80)')
    parser.add_argument('--right-scale', type=float, default=RIGHT_SCALE,
        help='Faculty height fraction on right side  (default: 0.45)')
    parser.add_argument('--detect-interval', type=int, default=DETECT_INTERVAL,
        help='Re-analyse background every N frames in video-bg mode  (default: 15)')
    parser.add_argument('--start-slide', type=int, default=0,
        help='Trim output to start from this slide number (1-indexed, 0 = no trim)')

    args = parser.parse_args()

    # Validate
    if not Path(args.faculty_video).exists():
        print(f"[ERROR] Faculty video not found: {args.faculty_video}")
        sys.exit(1)
    if not Path(args.slides_source).exists():
        print(f"[ERROR] Slides source not found: {args.slides_source}")
        sys.exit(1)

    key_bgr    = hex_to_bgr(args.key)
    input_type = detect_input_type(args.slides_source)
    print(f"Slides source type: {input_type.upper()}")

    DETECT_INTERVAL = args.detect_interval

    if input_type == 'video':
        # ── MP4 slideshow mode ────────────────────────────────────────────────
        print("[3/4] Using background video — content detected per frame")
        run_video_pipeline(
            faculty_video      = args.faculty_video,
            bg_video           = args.slides_source,
            output             = args.output,
            key_bgr            = key_bgr,
            sim                = args.sim,
            smooth             = args.smooth,
            spill              = args.spill,
            sensitivity        = args.sensitivity,
            lerp_speed         = args.lerp_speed,
            center_scale       = args.center_scale,
            right_scale        = args.right_scale,
            sync_timings_file  = args.sync,
        )

    else:
        # ── PPTX or images dir mode ───────────────────────────────────────────
        if input_type == 'pptx':
            tmp_dir = str(Path(args.output).parent / "_slides_tmp")
            slide_paths = pptx_to_images(args.slides_source, tmp_dir)
            slide_images = [cv2.imread(p) for p in slide_paths]
        else:  # images dir
            print("[3/4] Loading slide images …")
            slide_images = load_images_from_dir(args.slides_source)

        if not slide_images or any(s is None for s in slide_images):
            print("[ERROR] Failed to load one or more slide images.")
            sys.exit(1)

        run_slides_pipeline(
            faculty_video  = args.faculty_video,
            slide_images   = slide_images,
            output         = args.output,
            timings_file   = args.timings,
            key_bgr        = key_bgr,
            sim            = args.sim,
            smooth         = args.smooth,
            spill          = args.spill,
            sensitivity    = args.sensitivity,
            lerp_speed     = args.lerp_speed,
            center_scale   = args.center_scale,
            right_scale    = args.right_scale,
            start_slide    = args.start_slide,
        )


if __name__ == '__main__':
    main()
