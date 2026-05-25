#!/usr/bin/env python3
"""
Speech-to-slide sync: matches faculty speech to VSB_Video slides.

Algorithm
---------
1. Detect scene changes in VSB_Video (OpenCV frame-diff, no extra deps)
2. Transcribe faculty audio:
     - faster-whisper  (if installed: pip install faster-whisper)
     - openai-whisper  (if installed: pip install openai-whisper)
     - ffmpeg silencedetect  (always available fallback)
3. OCR each VSB scene frame:
     - pytesseract  (if installed: brew install tesseract && pip install pytesseract)
     - skipped otherwise → falls back to sequential mapping
4. Match speech segments to slides:
     - TF-IDF cosine similarity  (if sklearn installed: pip install scikit-learn)
     - sequential mapping fallback
5. Output sync_timings.json:
     [{"fac_start": 0.0, "fac_end": 8.3, "vsb_start": 0.0}, ...]

Usage
-----
  python speech_sync.py faculty.mp4 vsb_video.mp4 [-o sync_timings.json]
"""
import sys, json, subprocess, os, tempfile, argparse, re
from pathlib import Path
import cv2
import numpy as np


# ── Scene Detection ───────────────────────────────────────────────────────────

def detect_vsb_scenes(video_path: str, diff_threshold: float = 25.0,
                       min_scene_secs: float = 2.0) -> list:
    """Frame-diff scene detection. Returns [(time_secs, frame_idx), ...]."""
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        print(f"[ERROR] Cannot open VSB video: {video_path}")
        return [(0.0, 0)]

    fps   = cap.get(cv2.CAP_PROP_FPS) or 30.0
    total = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    min_gap = int(min_scene_secs * fps)

    print(f"[SYNC] Scanning {Path(video_path).name} for scene changes "
          f"({total} frames @ {fps:.1f}fps)…")

    scenes = [(0.0, 0)]
    prev_gray  = None
    last_fi    = 0

    for fi in range(total):
        ret, frame = cap.read()
        if not ret:
            break
        small = cv2.resize(frame, (160, 90))
        gray  = cv2.cvtColor(small, cv2.COLOR_BGR2GRAY).astype(np.float32)

        if prev_gray is not None and (fi - last_fi) >= min_gap:
            diff = np.abs(gray - prev_gray).mean()
            if diff > diff_threshold:
                t = fi / fps
                scenes.append((t, fi))
                last_fi = fi
                print(f"[SYNC] Scene at {t:.1f}s  (Δ={diff:.1f})")

        prev_gray = gray

    cap.release()
    print(f"[SYNC] Found {len(scenes)} scenes in VSB_Video")
    return scenes


def extract_scene_frame(video_path: str, time_secs: float) -> np.ndarray:
    """Read one frame at time_secs + 1s (skip transition). Returns frame or None."""
    cap = cv2.VideoCapture(video_path)
    cap.set(cv2.CAP_PROP_POS_MSEC, (time_secs + 1.0) * 1000)
    ret, frame = cap.read()
    cap.release()
    return frame if ret else None


# ── Audio / Transcription ─────────────────────────────────────────────────────

def extract_audio_wav(video_path: str, out_wav: str):
    result = subprocess.run(
        ['ffmpeg', '-y', '-i', video_path, '-ar', '16000', '-ac', '1', '-vn', out_wav],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        raise RuntimeError(f"Audio extract failed: {result.stderr[:300]}")


def transcribe_whisper(wav_path: str) -> list:
    """Try faster-whisper then openai-whisper. Returns [{start,end,text}] or None."""
    # faster-whisper (lighter, no torch)
    try:
        import faster_whisper
        print("[SYNC] Transcribing with faster-whisper (base model)…")
        model = faster_whisper.WhisperModel('base', device='cpu', compute_type='int8')
        segs, _ = model.transcribe(wav_path, beam_size=1)
        result = [{'start': s.start, 'end': s.end, 'text': s.text.strip()} for s in segs]
        print(f"[SYNC] Transcribed {len(result)} segments")
        return result
    except ImportError:
        pass

    # openai-whisper
    try:
        import whisper
        print("[SYNC] Transcribing with openai-whisper (base model)…")
        model = whisper.load_model('base')
        result = model.transcribe(wav_path)
        segs = [{'start': s['start'], 'end': s['end'], 'text': s['text'].strip()}
                for s in result['segments']]
        print(f"[SYNC] Transcribed {len(segs)} segments")
        return segs
    except ImportError:
        return None


def speech_segments_from_silence(video_path: str,
                                   silence_db: float = -35.0,
                                   min_silence_secs: float = 1.5) -> list:
    """Fallback: ffmpeg silencedetect → speech segments."""
    print("[SYNC] Whisper not found — using ffmpeg silence detection as fallback")
    result = subprocess.run(
        ['ffmpeg', '-i', video_path,
         '-af', f'silencedetect=noise={silence_db}dB:d={min_silence_secs}',
         '-f', 'null', '-'],
        capture_output=True, text=True
    )
    silences = []
    for line in result.stderr.splitlines():
        m = re.search(r'silence_start: ([\d.]+)', line)
        if m:
            silences.append({'s': float(m.group(1)), 'e': None})
        m = re.search(r'silence_end: ([\d.]+)', line)
        if m and silences and silences[-1]['e'] is None:
            silences[-1]['e'] = float(m.group(1))

    dur_r = subprocess.run(
        ['ffprobe', '-v', 'error', '-show_entries', 'format=duration',
         '-of', 'default=noprint_wrappers=1:nokey=1', video_path],
        capture_output=True, text=True
    )
    try:
        duration = float(dur_r.stdout.strip())
    except Exception:
        duration = 3600.0

    speech, prev_end = [], 0.0
    for sil in silences:
        if sil['e'] is None:
            sil['e'] = sil['s'] + 2.0
        if sil['s'] > prev_end + 0.5:
            speech.append({'start': prev_end, 'end': sil['s'], 'text': ''})
        prev_end = sil['e']
    if prev_end < duration - 0.5:
        speech.append({'start': prev_end, 'end': duration, 'text': ''})

    print(f"[SYNC] Silence detection found {len(speech)} speech segments")
    return speech


# ── OCR ───────────────────────────────────────────────────────────────────────

def ocr_frame(frame: np.ndarray) -> str:
    """pytesseract OCR. Returns '' if unavailable."""
    try:
        import pytesseract
        from PIL import Image
        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        return pytesseract.image_to_string(Image.fromarray(rgb))
    except Exception:
        return ''


# ── Matching ──────────────────────────────────────────────────────────────────

def tfidf_match(speech_segs: list, slide_texts: list):
    """TF-IDF cosine similarity. Returns assignment list or None."""
    try:
        from sklearn.feature_extraction.text import TfidfVectorizer
        from sklearn.metrics.pairwise import cosine_similarity

        n = len(speech_segs)
        all_texts = [s['text'] for s in speech_segs] + slide_texts
        if not any(t.strip() for t in all_texts):
            return None

        vec = TfidfVectorizer(stop_words='english', min_df=1)
        tfidf = vec.fit_transform(all_texts)
        sims  = cosine_similarity(tfidf[:n], tfidf[n:])
        return sims.argmax(axis=1).tolist()
    except (ImportError, ValueError):
        return None


def sequential_match(n_speech: int, n_slides: int) -> list:
    if n_slides == 0:
        return [0] * n_speech
    return [int(i * n_slides / max(n_speech, 1)) % n_slides for i in range(n_speech)]


# ── Main ──────────────────────────────────────────────────────────────────────

def run_sync(faculty_video: str, vsb_video: str, output_file: str,
             scene_threshold: float = 25.0, min_scene_secs: float = 2.0):

    print(f"[SYNC] Faculty : {Path(faculty_video).name}")
    print(f"[SYNC] VSB     : {Path(vsb_video).name}")
    print(f"[SYNC] Output  : {Path(output_file).name}")

    # 1 – scene detection
    scene_list  = detect_vsb_scenes(vsb_video, scene_threshold, min_scene_secs)
    scene_times = [t for t, _ in scene_list]
    n_scenes    = len(scene_times)

    # 2 – transcription
    tmp_wav = tempfile.NamedTemporaryFile(suffix='.wav', delete=False).name
    speech_segs = None
    try:
        print("[SYNC] Extracting faculty audio…")
        extract_audio_wav(faculty_video, tmp_wav)
        speech_segs = transcribe_whisper(tmp_wav)
    except Exception as exc:
        print(f"[WARN] Audio extraction error: {exc}")
    finally:
        try:
            os.unlink(tmp_wav)
        except Exception:
            pass

    if speech_segs is None:
        speech_segs = speech_segments_from_silence(faculty_video)

    if not speech_segs:
        print("[WARN] No speech detected — splitting faculty time evenly across scenes")
        dur_r = subprocess.run(
            ['ffprobe', '-v', 'error', '-show_entries', 'format=duration',
             '-of', 'default=noprint_wrappers=1:nokey=1', faculty_video],
            capture_output=True, text=True
        )
        try:
            duration = float(dur_r.stdout.strip())
        except Exception:
            duration = 60.0
        seg_dur = duration / max(n_scenes, 1)
        speech_segs = [{'start': i*seg_dur, 'end': (i+1)*seg_dur, 'text': ''}
                       for i in range(n_scenes)]

    n_speech = len(speech_segs)
    print(f"[SYNC] {n_speech} speech segments · {n_scenes} VSB scenes")

    # 3 – OCR slides
    print("[SYNC] Running OCR on VSB scene frames…")
    slide_texts = []
    for t in scene_times:
        frame = extract_scene_frame(vsb_video, t)
        text  = ocr_frame(frame) if frame is not None else ''
        slide_texts.append(text.strip())
        preview = text.strip()[:60].replace('\n', ' ')
        print(f"[SYNC]   @{t:.1f}s → {repr(preview) if preview else '(no text)'}")

    # 4 – matching
    has_speech_text = any(s['text'].strip() for s in speech_segs)
    has_slide_text  = any(slide_texts)

    assignments = None
    if has_speech_text and has_slide_text:
        print("[SYNC] Running TF-IDF semantic match…")
        assignments = tfidf_match(speech_segs, slide_texts)
        if assignments:
            print("[SYNC] Semantic matching complete")

    if assignments is None:
        mode = "sequential" if not (has_speech_text and has_slide_text) else "sequential (sklearn not available)"
        print(f"[SYNC] Using {mode} mapping")
        assignments = sequential_match(n_speech, n_scenes)

    # 5 – build and save timings
    raw = []
    for i, seg in enumerate(speech_segs):
        idx = min(int(assignments[i]), n_scenes - 1)
        raw.append({
            'fac_start': round(float(seg['start']), 3),
            'fac_end':   round(float(seg['end']),   3),
            'vsb_start': round(float(scene_times[idx]), 3),
        })

    # Merge adjacent entries with same vsb_start (reduces unnecessary seeks)
    merged = []
    for e in raw:
        if merged and merged[-1]['vsb_start'] == e['vsb_start']:
            merged[-1]['fac_end'] = e['fac_end']
        else:
            merged.append(dict(e))

    with open(output_file, 'w') as f:
        json.dump(merged, f, indent=2)

    print(f"[DONE] Sync timings → {output_file}  ({len(merged)} segments)")
    for e in merged:
        print(f"       fac [{e['fac_start']:.1f}–{e['fac_end']:.1f}s]"
              f"  →  vsb @{e['vsb_start']:.1f}s")


if __name__ == '__main__':
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument('faculty_video')
    ap.add_argument('vsb_video')
    ap.add_argument('-o', '--output',        default='sync_timings.json')
    ap.add_argument('--scene-threshold',     type=float, default=25.0,
                    help='Mean pixel diff to trigger a scene change (default 25)')
    ap.add_argument('--min-scene-secs',      type=float, default=2.0,
                    help='Minimum seconds between scene changes (default 2)')
    args = ap.parse_args()

    for p in (args.faculty_video, args.vsb_video):
        if not Path(p).exists():
            print(f"[ERROR] File not found: {p}")
            sys.exit(1)

    run_sync(args.faculty_video, args.vsb_video, args.output,
             args.scene_threshold, args.min_scene_secs)
