#!/usr/bin/env python3
"""PPTX-to-speech sync: reads slide text from PPTX, transcribes faculty audio with Whisper,
matches each slide to the time range when its content is spoken, writes timings.json.

Output format (consumed by pipeline.py --timings / -t flag):
  [{"slide": 0, "start": 0.0, "end": 45.2}, {"slide": 1, "start": 45.2, "end": 92.0}, ...]

Usage:
  python pptx_sync.py faculty.mp4 slides.pptx [-o timings.json]
"""
import sys, json, subprocess, os, tempfile, argparse, re
import numpy as np
from pathlib import Path


# ── Slide Text Extraction ─────────────────────────────────────────────────────

def extract_slide_texts(pptx_path: str) -> list:
    """Return one text string per slide: speaker notes + visible slide text combined.
    Speaker notes are the primary match source — they contain what the professor
    actually says.  Visible slide text is appended as secondary keyword signal.
    Requires python-pptx."""
    try:
        from pptx import Presentation
    except ImportError:
        print("[SYNC] python-pptx not installed — run: pip install python-pptx")
        sys.exit(1)

    prs = Presentation(pptx_path)
    texts = []
    notes_count = 0
    for slide in prs.slides:
        # ── Speaker notes (primary — what the presenter actually speaks) ──────
        notes_text = ''
        try:
            if slide.has_notes_slide:
                raw = slide.notes_slide.notes_text_frame.text.strip()
                # Skip the default placeholder that PowerPoint inserts in empty notes
                if raw and 'click to edit' not in raw.lower():
                    notes_text = raw
                    notes_count += 1
        except Exception:
            pass

        # ── Visible slide text (secondary — titles, bullets, labels) ─────────
        parts = []
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                t = ' '.join(run.text for run in para.runs).strip()
                if t:
                    parts.append(t)
        slide_text = ' '.join(parts)

        # Combine: notes first (stronger TF-IDF signal), then slide keywords.
        combined = (notes_text + ' ' + slide_text).strip() if notes_text else slide_text
        texts.append(combined)

    print(f"[SYNC] Extracted text from {len(texts)} slides "
          f"({notes_count} with speaker notes)")
    for i, t in enumerate(texts):
        preview = t[:80].replace('\n', ' ')
        print(f"[SYNC]   Slide {i+1}: {repr(preview) if preview else '(no text)'}")
    return texts


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


def speech_segments_from_silence(video_path: str) -> list:
    """Fallback when Whisper is unavailable: ffmpeg silencedetect → speech segments."""
    print("[SYNC] Whisper not found — using ffmpeg silence detection as fallback")
    result = subprocess.run(
        ['ffmpeg', '-i', video_path,
         '-af', 'silencedetect=noise=-35dB:d=1.5', '-f', 'null', '-'],
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


# ── Slide Matching ────────────────────────────────────────────────────────────

def tfidf_match(speech_segs: list, slide_texts: list):
    """Monotone TF-IDF match: cosine similarity with a non-decreasing slide constraint.
    Uses dynamic programming so slides advance in sequence — never jump backward.
    dp[i][j] = best cumulative similarity for first i segments with i-th assigned to slide j."""
    try:
        from sklearn.feature_extraction.text import TfidfVectorizer
        from sklearn.metrics.pairwise import cosine_similarity

        n = len(speech_segs)
        m = len(slide_texts)
        all_texts = [s['text'] for s in speech_segs] + slide_texts
        if not any(t.strip() for t in all_texts):
            return None

        vec   = TfidfVectorizer(stop_words='english', min_df=1)
        tfidf = vec.fit_transform(all_texts)
        sims  = cosine_similarity(tfidf[:n], tfidf[n:])  # n × m

        # DP: ensure slide index never decreases across segments
        dp      = np.full((n, m), -np.inf)
        dp[0]   = sims[0]
        for i in range(1, n):
            # cummax[j] = max(dp[i-1, 0..j]) — best previous state ending at slide ≤ j
            cummax = np.maximum.accumulate(dp[i - 1])
            dp[i]  = cummax + sims[i]

        # Backtrack
        assignments = [0] * n
        assignments[n - 1] = int(np.argmax(dp[n - 1]))
        for i in range(n - 2, -1, -1):
            assignments[i] = int(np.argmax(dp[i, : assignments[i + 1] + 1]))

        distinct = len(set(assignments))
        print(f"[SYNC] Monotone TF-IDF: {n} segments → {distinct} distinct slides used")
        return assignments
    except (ImportError, ValueError):
        return None


def sequential_match(n_speech: int, n_slides: int) -> list:
    """Evenly distribute speech segments across slides in order."""
    if n_slides == 0:
        return [0] * n_speech
    return [int(i * n_slides / max(n_speech, 1)) % n_slides for i in range(n_speech)]


# ── Main ──────────────────────────────────────────────────────────────────────

def run_pptx_sync(faculty_video: str, pptx_path: str, output_file: str):
    print(f"[SYNC] Faculty : {Path(faculty_video).name}")
    print(f"[SYNC] PPTX    : {Path(pptx_path).name}")
    print(f"[SYNC] Output  : {Path(output_file).name}")

    # 1 — extract exact text from every slide (no OCR, no guessing)
    slide_texts = extract_slide_texts(pptx_path)
    n_slides = len(slide_texts)
    if n_slides == 0:
        print("[ERROR] No slides found in PPTX")
        sys.exit(1)

    # 2 — transcribe faculty audio
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

    # Get total duration so last timing entry covers the full video
    dur_r = subprocess.run(
        ['ffprobe', '-v', 'error', '-show_entries', 'format=duration',
         '-of', 'default=noprint_wrappers=1:nokey=1', faculty_video],
        capture_output=True, text=True
    )
    try:
        duration = float(dur_r.stdout.strip())
    except Exception:
        duration = 0.0

    if not speech_segs:
        print("[WARN] No speech detected — splitting faculty duration evenly across slides")
        seg_dur = duration / n_slides
        speech_segs = [{'start': i * seg_dur, 'end': (i + 1) * seg_dur, 'text': ''}
                       for i in range(n_slides)]

    n_speech = len(speech_segs)
    print(f"[SYNC] {n_speech} speech segments · {n_slides} slides")

    # 3 — match each speech segment to its best-fit slide
    has_speech_text = any(s['text'].strip() for s in speech_segs)
    has_slide_text  = any(slide_texts)
    assignments = None

    if has_speech_text and has_slide_text:
        print("[SYNC] Running monotone TF-IDF match (sequential, no backward jumps)…")
        assignments = tfidf_match(speech_segs, slide_texts)
        if assignments:
            print("[SYNC] Sequential matching complete")

    if assignments is None:
        mode = ("sequential" if not (has_speech_text and has_slide_text)
                else "sequential (sklearn not available — pip install scikit-learn)")
        print(f"[SYNC] Using {mode} mapping")
        assignments = sequential_match(n_speech, n_slides)

    # 4 — build timings: merge consecutive segments assigned to the same slide
    raw = []
    for i, seg in enumerate(speech_segs):
        idx = min(int(assignments[i]), n_slides - 1)
        raw.append({'slide': idx,
                    'start': round(float(seg['start']), 3),
                    'end':   round(float(seg['end']),   3)})

    merged = []
    for e in raw:
        if merged and merged[-1]['slide'] == e['slide']:
            merged[-1]['end'] = e['end']   # extend same-slide block
        else:
            merged.append(dict(e))

    # Extend the last entry to cover the full video duration
    if merged and duration > 0:
        merged[-1]['end'] = max(merged[-1]['end'], round(duration, 3))

    # Fill any silent gaps (no speech segment) by holding the previous slide
    filled = []
    for e in merged:
        if filled and filled[-1]['end'] < e['start'] - 0.01:
            filled.append({'slide': filled[-1]['slide'],
                           'start': filled[-1]['end'],
                           'end':   e['start']})
        filled.append(e)

    # Auto-insert placeholder slides that precede the first speech-matched slide.
    # Slides with < 10 meaningful words in their notes (logos, title cards, etc.)
    # are shown briefly at the start without disturbing the TF-IDF-matched timings.
    # All matched timings are shifted forward by 1 s per inserted placeholder so
    # the audio-slide sync remains intact from the first real content slide.
    PLACEHOLDER_SECS = 1.0   # seconds to show each placeholder slide

    if filled and filled[0]['slide'] != 0:
        first_matched = filled[0]['slide']

        def _word_count(text):
            return len(re.findall(r'\b[a-zA-Z]{3,}\b', text))

        placeholders = [i for i in range(first_matched)
                        if _word_count(slide_texts[i]) < 10]
        skipped_with_notes = [i for i in range(first_matched)
                              if _word_count(slide_texts[i]) >= 10]

        if skipped_with_notes:
            print(f"[WARN] Slides {[i+1 for i in skipped_with_notes]} have notes "
                  f"but were not matched — check sync quality or re-run sync.")

        if placeholders:
            shift = len(placeholders) * PLACEHOLDER_SECS
            filled = [{'slide': e['slide'],
                       'start': round(e['start'] + shift, 3),
                       'end':   round(e['end']   + shift, 3)}
                      for e in filled]
            head = [{'slide': idx,
                     'start': round(k * PLACEHOLDER_SECS, 3),
                     'end':   round((k + 1) * PLACEHOLDER_SECS, 3)}
                    for k, idx in enumerate(placeholders)]
            filled = head + filled
            print(f"[SYNC] Inserted {len(placeholders)} placeholder slide(s) at start "
                  f"(+{shift:.1f}s shift on matched timings)")

    with open(output_file, 'w') as f:
        json.dump(filled, f, indent=2)

    print(f"[DONE] Slide timings → {output_file}  ({len(filled)} entries)")
    for e in filled:
        print(f"       Slide {e['slide']+1:2d}  [{e['start']:7.1f}s – {e['end']:7.1f}s]")


if __name__ == '__main__':
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument('faculty_video')
    ap.add_argument('pptx_file')
    ap.add_argument('-o', '--output', default='timings.json')
    args = ap.parse_args()

    for p in (args.faculty_video, args.pptx_file):
        if not Path(p).exists():
            print(f"[ERROR] File not found: {p}")
            sys.exit(1)

    run_pptx_sync(args.faculty_video, args.pptx_file, args.output)
