import cv2
import numpy as np
import subprocess
import os
import gc
import fitz
from PIL import Image

# optional: rembg — U2Net-based background removal (best quality, smooth hair edges).
# Install with:  pip install rembg[gpu]  or  pip install rembg
# If not installed the code falls back to MediaPipe / TFLite / chroma-key.
_REMBG_AVAILABLE = False
_rembg_session   = None
try:
    import rembg as _rembg_module
    from rembg import remove as _rembg_remove, new_session as _rembg_new_session
    from PIL import Image as _PIL_Image
    _REMBG_AVAILABLE = True
except Exception:
    pass

def _get_rembg_session():
    global _rembg_session
    if not _REMBG_AVAILABLE:
        return None
    if _rembg_session is None:
        try:
            _rembg_session = _rembg_new_session("u2net")
            print("[INFO] Segmentation: rembg U2Net (Tier 0 — best quality).")
        except Exception as e:
            print(f"[WARN] rembg session init failed: {e}")
    return _rembg_session

# optional: mediapipe for high-quality person segmentation
# Catch any exception (not just ImportError) because mediapipe 0.10.x can
# raise ValueError due to numpy/sklearn binary incompatibility on import.
try:
    import mediapipe as mp
    _MP_AVAILABLE = True
except Exception:
    _MP_AVAILABLE = False

# optional: TFLite via tensorflow (lite only — does NOT trigger keras/sklearn)
_TF_AVAILABLE = False
try:
    # Import only the lite submodule to avoid triggering keras/sklearn chain
    import tensorflow as _tf_module
    # Verify lite interpreter is actually accessible before committing
    _ = _tf_module.lite.Interpreter
    import tensorflow as tf
    _TF_AVAILABLE = True
except Exception:
    try:
        import tflite_runtime.interpreter as _tflite_rt
        _TF_AVAILABLE = "tflite_runtime"
    except Exception:
        pass

BASE_DIR = "/Users/admin/Desktop/Emeritus"

PPTX_PATH     = os.path.join(BASE_DIR, "CS_W8_S4_VSB_V1_Reviewed.pptx")
FACULTY_VIDEO = os.path.join(BASE_DIR, "Green-screenFacultyRecording.mp4")

SOFFICE_PATH = "/usr/local/bin/soffice"
FFMPEG_PATH  = "ffmpeg"

OUT_DIR     = os.path.join(BASE_DIR, "slides_output")
TEMP_VIDEO  = os.path.join(BASE_DIR, "temp_video_no_audio.mp4")
FINAL_VIDEO = os.path.join(BASE_DIR, "final_synced_output_final.mp4")

WIDTH  = 1920
HEIGHT = 1080
OUTPUT_FPS      = 30
AUDIO_START_TIME = 7

SLIDE_TIMELINE = [
    (0,    2,  0, "hide"),    # Slide 1: hidden               (0 – 2 s)
    (2,    7,  1, "hide"),    # Slide 2: hidden               (2 – 7 s)
    # Cut 1: source 7 – 13 s is removed by add_audio
    (7,   13,  2, "center"),  # Cut-1 region (frames discarded in add_audio)
    (13,  31,  2, "center"),  # Slide 3: center               (13 – 31 s → 18 s output)
    (31,  68,  3, "right"),   # Slide 4: right                (31 – 68 s → 37 s output)
    (68,  96,  4, "center"),  # Slide 5: center               (68 – 96 s → 28 s output)
    # Cut 2: source 96 – 103 s is removed by add_audio
    (96,  103, 4, "center"),  # Cut-2 region (frames discarded in add_audio)
    (103, 133, 5, "right"),   # Slide 6: right                (103 – 133 s → 30 s output)
]
# Total output after both cuts: 7 + 18 + 37 + 28 + 30 = 120 s (2:00)

FACULTY_HEIGHT = 950
RIGHT_SHIFT    = -210   # negative → faculty pushed further right; right edge clips off-screen gracefully

REMBG_INTERVAL  = 3     # compute new segmentation every N frames (3 = 10 Hz at 30fps)
                        # Less frequent updates → more temporal coherence between model runs
ALPHA_SMOOTH    = 0.85  # EMA weight when alpha is INCREASING (foreground growing) — keeps edges smooth
ALPHA_DECAY     = 0.60  # EMA weight when alpha is DECREASING (foreground shrinking) — clears ghosts
                        # 0.60^6 ≈ 4 % → ghost pixel near-invisible in ~6 frames (200 ms); prevents edge flicker
BBOX_SMOOTH     = 0.95  # per-frame EMA weight for bounding box (x1,x2,y1)
TABLE_POOL_DECAY = 0.997 # per-frame retention for pooled table region
                         # 0.3% decay/frame → fades after ~333 frames if unused

_rembg_cache = {
    "bgr_cut": None,
    "alpha":   None,
    "at_frame": -9999,
}

_bbox_smooth     = {"x1": None, "x2": None, "y1": None, "y2": None}
# Persistent pooled table alpha — max-accumulated across frames with slow decay.
# Keeps the table/podium region stable regardless of person sway.
_table_only_pool = None

# Two-buffer temporal smoothing:
#   _alpha_target  — the raw mask from the last model run (updated every REMBG_INTERVAL frames)
#   _alpha_display — what actually gets composited (steps toward target EVERY frame)
# This eliminates the 10 Hz periodic "pop" that occurred when EMA was only
# applied at model-update frames and the alpha sat frozen in between.
_alpha_target  = {"alpha": None}
_alpha_display = {"alpha": None}


def _step_display_alpha():
    """
    Advance the displayed alpha one step toward the current target using
    ASYMMETRIC EMA — different rates for growing vs shrinking alpha.

    Why asymmetric?
    ───────────────
    Symmetric EMA (same weight in both directions) creates a fundamental
    tradeoff: high smoothing prevents flicker but makes ghost pixels at
    previous head/body positions linger for 60+ frames (2+ seconds).

    Asymmetric EMA breaks that tradeoff:

    • alpha INCREASING  (target > display — person entering a region)
      → use ALPHA_SMOOTH = 0.85  (slow ease-in → no flicker on new edges)

    • alpha DECREASING  (target < display — person left this region)
      → use ALPHA_DECAY  = 0.60  (gradual fade-out → ghost clears in ~6 frames)
        0.60^3 ≈ 22 %,  0.60^5 ≈ 8 %,  0.60^6 ≈ 5 %  → effectively gone

    Result: edges appear smoothly, motion ghosts/shadows vanish quickly.
    """
    global _alpha_display, _alpha_target
    t = _alpha_target["alpha"]
    if t is None:
        return
    d = _alpha_display["alpha"]
    if d is None or d.shape != t.shape:
        _alpha_display["alpha"] = t.copy()
        return

    d_f = d.astype(np.float32)
    t_f = t.astype(np.float32)

    # Per-pixel direction: is alpha growing or shrinking?
    result = np.where(
        t_f >= d_f,
        # Growing → slow ease-in (anti-flicker)
        ALPHA_SMOOTH * d_f + (1.0 - ALPHA_SMOOTH) * t_f,
        # Shrinking → fast fade-out (anti-ghost / anti-shadow)
        ALPHA_DECAY  * d_f + (1.0 - ALPHA_DECAY)  * t_f,
    )
    _alpha_display["alpha"] = np.clip(result, 0, 255).astype(np.uint8)

# ---------------------------------------------------------------------------
# Segmentation backends  (priority: MediaPipe > TFLite > pure chroma key)
# ---------------------------------------------------------------------------

# --- Tier 1: MediaPipe ---
_mp_seg = None   # None = uninitialised | False = unavailable

def _get_mp_seg():
    global _mp_seg
    if _mp_seg is None:
        if not _MP_AVAILABLE:
            _mp_seg = False
        else:
            try:
                _mp_seg = mp.solutions.selfie_segmentation.SelfieSegmentation(
                    model_selection=1
                )
                print("[INFO] Segmentation: MediaPipe (Tier 1).")
            except Exception as e:
                print(f"[WARN] MediaPipe init failed: {e}")
                _mp_seg = False
    return _mp_seg if _mp_seg is not False else None


# --- Tier 2: Direct TFLite with selfie_segmenter.tflite ---
_tflite_ctx = None   # None = uninitialised | False = unavailable

def _get_tflite_ctx():
    global _tflite_ctx
    if _tflite_ctx is not None:
        return _tflite_ctx if _tflite_ctx is not False else None

    model_path = os.path.join(BASE_DIR, "selfie_segmenter.tflite")
    if not os.path.exists(model_path) or not _TF_AVAILABLE:
        _tflite_ctx = False
        return None

    try:
        if _TF_AVAILABLE == "tflite_runtime":
            interp = _tflite_rt.Interpreter(model_path=model_path)
        else:
            interp = tf.lite.Interpreter(model_path=model_path)

        interp.allocate_tensors()
        in_det  = interp.get_input_details()
        out_det = interp.get_output_details()
        shape   = in_det[0]["shape"]         # [1, H, W, 3]
        ih, iw  = int(shape[1]), int(shape[2])

        in_dtype = in_det[0]["dtype"]   # np.float32 or np.uint8

        _tflite_ctx = {
            "interp":    interp,
            "in_idx":    in_det[0]["index"],
            "out_idx":   out_det[0]["index"],
            "out_shape": out_det[0]["shape"],
            "ih": ih, "iw": iw,
            "in_dtype": in_dtype,
        }
        print(f"[INFO] Segmentation: TFLite {ih}x{iw} dtype={in_dtype.__name__} (Tier 2).")
    except Exception as e:
        print(f"[WARN] TFLite init failed: {e}")
        _tflite_ctx = False

    return _tflite_ctx if _tflite_ctx is not False else None


def _person_mask(frame_bgr):
    """
    Return float32 person-confidence mask [0,1] at the same size as frame_bgr.
    Returns None when no ML backend is available (chroma-key fallback used).
    """
    h, w = frame_bgr.shape[:2]

    # Tier 1: MediaPipe
    seg = _get_mp_seg()
    if seg is not None:
        small = cv2.resize(frame_bgr, (w // 2, h // 2), interpolation=cv2.INTER_AREA)
        rgb   = cv2.cvtColor(small, cv2.COLOR_BGR2RGB)
        try:
            result = seg.process(rgb)
            if result.segmentation_mask is not None:
                mask = result.segmentation_mask.astype(np.float32)
                return cv2.resize(mask, (w, h), interpolation=cv2.INTER_LINEAR)
        except Exception:
            pass

    # Tier 2: TFLite
    ctx = _get_tflite_ctx()
    if ctx is not None:
        ih, iw   = ctx["ih"], ctx["iw"]
        in_dtype = ctx["in_dtype"]
        small    = cv2.resize(frame_bgr, (iw, ih), interpolation=cv2.INTER_AREA)
        rgb      = cv2.cvtColor(small, cv2.COLOR_BGR2RGB)

        # Handle both float32 models (normalised [0,1]) and uint8 models
        if in_dtype == np.float32:
            inp = (rgb.astype(np.float32) / 255.0)[np.newaxis, ...]
        else:
            inp = rgb.astype(in_dtype)[np.newaxis, ...]

        try:
            interp = ctx["interp"]
            interp.set_tensor(ctx["in_idx"], inp)
            interp.invoke()
            out = interp.get_tensor(ctx["out_idx"])
            if out.ndim == 4:
                mask = out[0, 0] if out.shape[1] == 1 else out[0, :, :, 0]
            else:
                mask = out[0]
            # Normalise uint8 output to [0,1] if needed
            if mask.dtype == np.uint8:
                mask = mask.astype(np.float32) / 255.0
            mask = np.clip(mask, 0.0, 1.0).astype(np.float32)
            return cv2.resize(mask, (w, h), interpolation=cv2.INTER_LINEAR)
        except Exception as e:
            print(f"[WARN] TFLite inference failed: {e}")

    # Tier 3: no ML backend
    return None


# ---------------------------------------------------------------------------
# Guided filter  (edge-preserving alpha refinement, no Gaussian blur)
# He et al., "Guided Image Filtering", TPAMI 2013
# ---------------------------------------------------------------------------
def guided_filter(guide_gray, alpha, r=8, eps=0.001):
    """
    guide_gray and alpha are float32 in [0, 1].
    Sharpens alpha at real edges (hair, shoulders) while smoothing flat areas.
    """
    box = (2 * r + 1, 2 * r + 1)

    mean_I  = cv2.boxFilter(guide_gray,               cv2.CV_32F, box)
    mean_p  = cv2.boxFilter(alpha,                    cv2.CV_32F, box)
    mean_Ip = cv2.boxFilter(guide_gray * alpha,       cv2.CV_32F, box)
    mean_II = cv2.boxFilter(guide_gray * guide_gray,  cv2.CV_32F, box)

    cov_Ip = mean_Ip - mean_I * mean_p
    var_I  = mean_II - mean_I * mean_I

    a = cov_Ip / (var_I + eps)
    b = mean_p - a * mean_I

    mean_a = cv2.boxFilter(a, cv2.CV_32F, box)
    mean_b = cv2.boxFilter(b, cv2.CV_32F, box)

    return np.clip(mean_a * guide_gray + mean_b, 0.0, 1.0)


# ---------------------------------------------------------------------------
def _add_table_to_alpha(frame_bgr, alpha_u8):
    """
    Extend the rembg person alpha downward column-by-column to include the
    desk/table in front of the presenter.

    rembg's U2Net is trained on person segmentation and may not include a
    desk that appears below the torso.  This function finds the lowest opaque
    person pixel in each column and extends the mask downward as long as the
    pixel is NOT a green-screen pixel — so the real physical desk is captured
    but the green background behind it is excluded.
    """
    h, w = alpha_u8.shape

    # Green-screen detection — use a LOOSE threshold so only definite
    # chroma-key green stops the downward extension.
    # Tighter thresholds would mis-classify dark desk pixels as non-green
    # when the green light spills onto the desk surface.
    hsv = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2HSV)
    H   = hsv[:, :, 0].astype(np.float32)
    S   = hsv[:, :, 1].astype(np.float32)
    b_  = frame_bgr[:, :, 0].astype(np.float32)
    g_  = frame_bgr[:, :, 1].astype(np.float32)
    r_  = frame_bgr[:, :, 2].astype(np.float32)
    # Require strong green dominance + high saturation to avoid
    # mis-classifying dark table edges as green-screen.
    is_green = (
        (H >= 38) & (H <= 85) & (S >= 60) &
        (g_ > r_ * 1.15) & (g_ > b_ * 1.15)
    )

    result = alpha_u8.copy()
    # Lower threshold: even faint edge pixels (alpha > 8) count as "person bottom"
    # so the downward extension starts from the true lowest visible boundary.
    person = alpha_u8 > 8

    if not person.any():
        return result

    # Find last (bottom-most) person row per column
    has_person = person.any(axis=0)
    last_row   = np.where(
        has_person,
        h - 1 - np.argmax(person[::-1, :], axis=0),
        -1
    )

    # Extend downward column-by-column while pixels are non-green.
    # Also fill any isolated green pixels surrounded by non-green
    # (avoids stopping at a single spill-contaminated desk pixel).
    for col in range(w):
        lr = int(last_row[col])
        if lr < 0 or lr >= h - 1:
            continue
        consecutive_green = 0
        for row in range(lr + 1, h):
            if is_green[row, col]:
                consecutive_green += 1
                # Stop only after 25 consecutive green rows — allows bridging
                # the green gap that appears between the faculty body and the
                # top of the podium/table.
                if consecutive_green >= 25:
                    break
            else:
                consecutive_green = 0
                result[row, col] = 255

    return result


# ---------------------------------------------------------------------------
def _stable_table_merge(person_alpha_u8, frame_bgr):
    """
    Return person_alpha merged with a TEMPORALLY STABLE table/podium region.

    Problem: _add_table_to_alpha() starts its downward extension from the
    lowest person pixel each frame.  When the faculty breathes or sways, that
    boundary shifts a few pixels and the whole table region shifts with it —
    causing visible table shake in the composited output.

    Solution:
      1. Compute the current table extension (pixels added BELOW the person).
      2. Max-pool that extension into a persistent float32 accumulator that
         decays at TABLE_POOL_DECAY per frame.  The table becomes "sticky":
         once detected it stays solid and only fades if never re-detected
         (e.g., faculty walks out of shot entirely).
      3. Merge: final = max(dynamic_person_alpha, pooled_stable_table).
         The person region is still fully dynamic (natural movement kept);
         only the table pixels are locked via the pool.
    """
    global _table_only_pool

    # Full alpha including current table extension
    full_alpha = _add_table_to_alpha(frame_bgr, person_alpha_u8)

    # Table-only pixels: those added by the extension (not in person mask)
    table_only_f = np.where(
        full_alpha.astype(np.int16) > person_alpha_u8.astype(np.int16),
        full_alpha.astype(np.float32),
        0.0
    )

    # Max-pool with slow decay — accumulates table over time
    if _table_only_pool is None or _table_only_pool.shape != table_only_f.shape:
        _table_only_pool = table_only_f.copy()
    else:
        # Decay existing pool then blend with current detection (take max)
        _table_only_pool = np.maximum(
            _table_only_pool * TABLE_POOL_DECAY,
            table_only_f
        )

    stable_table = np.clip(_table_only_pool, 0, 255).astype(np.uint8)
    return np.maximum(person_alpha_u8, stable_table)


# ---------------------------------------------------------------------------
def get_media_duration(path):
    result = subprocess.run(
        [
            "ffprobe", "-v", "error",
            "-show_entries", "format=duration",
            "-of", "default=noprint_wrappers=1:nokey=1",
            path,
        ],
        capture_output=True, text=True, check=True,
    )
    return float(result.stdout.strip())


_slide_max     = max(end for _, end, _, _ in SLIDE_TIMELINE)
TOTAL_DURATION = _slide_max  # Always end when the last SLIDE_TIMELINE entry ends (slide 6 @ 133 s)


def run_cmd(cmd):
    subprocess.run(cmd, check=True, stdin=subprocess.DEVNULL)


def convert_ppt_to_images():
    os.makedirs(OUT_DIR, exist_ok=True)
    run_cmd([
        SOFFICE_PATH, "--headless",
        "--convert-to", "pdf",
        "--outdir", OUT_DIR,
        PPTX_PATH,
    ])

    pdf_name = os.path.splitext(os.path.basename(PPTX_PATH))[0] + ".pdf"
    pdf_path = os.path.join(OUT_DIR, pdf_name)
    doc      = fitz.open(pdf_path)
    slide_paths = []

    # PPTX slides 1,4,6 are blank (0 shapes). Use content slides instead:
    # PPTX page: 0=blank  1=metadata  2=SQL Injections  4=Learning Obj  6=SQL Queries  7=SQL Injection Attacks
    SLIDE_PAGES = [0, 1, 2, 4, 6, 7]
    for slot, page_idx in enumerate(SLIDE_PAGES):
        page     = doc[page_idx]
        pix      = page.get_pixmap(matrix=fitz.Matrix(2.8, 2.8))
        img_path = os.path.join(OUT_DIR, f"slide_{slot + 1}.png")
        pix.save(img_path)
        slide_paths.append(img_path)

    doc.close()
    return slide_paths


def get_slide_info(t):
    for start, end, slide_index, position in SLIDE_TIMELINE:
        if start <= t < end:
            return slide_index, position
    return 5, "right"


def ease_out(p):
    p = max(0, min(1, p))
    return 1 - (1 - p) * (1 - p)


def animate_slide_2(slide, t):
    local_t  = t - 2
    progress = ease_out((local_t - 0.2) / 1.6)

    if progress >= 1.0:
        return slide.copy()

    original = slide.copy()
    animated = slide.copy()

    x1, y1, x2, y2 = 70, 150, 800, 305
    text_region = original[y1:y2, x1:x2].copy()

    b_ch, g_ch, r_ch = cv2.split(text_region)
    text_mask = np.where(
        (r_ch > 180) & (g_ch > 180) & (b_ch > 180),
        np.uint8(255), np.uint8(0)
    )
    text_mask     = cv2.dilate(text_mask, np.ones((3, 3), np.uint8), iterations=1)
    text_mask_bool = text_mask > 0
    text_mask_3d   = text_mask_bool[:, :, np.newaxis]

    bottom_strip = original[y2 - 15:y2, x1:x2]
    box_bg_col   = np.mean(bottom_strip, axis=0).astype(np.uint8)
    box_bg_fill  = np.broadcast_to(
        box_bg_col[np.newaxis, :, :], text_region.shape
    ).copy()

    animated[y1:y2, x1:x2] = np.where(
        text_mask_3d, box_bg_fill, animated[y1:y2, x1:x2]
    )

    if progress <= 0:
        return animated

    start_x   = -text_region.shape[1]
    current_x = int(start_x + (x1 - start_x) * progress)

    paste_x1 = max(0, current_x)
    paste_x2 = min(WIDTH, current_x + text_region.shape[1])
    src_x1   = paste_x1 - current_x
    src_x2   = src_x1 + (paste_x2 - paste_x1)

    if paste_x2 > paste_x1 and src_x2 > src_x1:
        dest_slice = animated[y1:y2, paste_x1:paste_x2]
        src_slice  = text_region[:, src_x1:src_x2]
        mask_slice = text_mask_bool[:, src_x1:src_x2, np.newaxis]
        animated[y1:y2, paste_x1:paste_x2] = np.where(
            mask_slice, src_slice, dest_slice
        )

    return animated


def animate_slide_4_points(slide, t):
    local_t  = t - 31
    original = slide.copy()
    animated = slide.copy()

    hx1, hy1, hx2, hy2 = 20, 15, 720, 125
    heading_img = original[hy1:hy2, hx1:hx2].copy()

    cv2.rectangle(animated, (hx1, hy1), (hx2, hy2), (255, 255, 255), -1)

    h_progress = ease_out((local_t - 0.15) / 1.3)
    if h_progress > 0:
        h_cur_x = int((hx1 - WIDTH) + WIDTH * h_progress)
        hp1 = max(0, h_cur_x)
        hp2 = min(WIDTH, h_cur_x + heading_img.shape[1])
        hs1 = hp1 - h_cur_x
        hs2 = hs1 + (hp2 - hp1)
        if hp2 > hp1 and hs2 > hs1:
            animated[hy1:hy2, hp1:hp2] = heading_img[:, hs1:hs2]

    # Limit white clear-area to content columns only (0-1380).
    # The faculty sits on the RIGHT side (~x=1540+), so extending to 1920
    # was painting white BEHIND the faculty head and below the body.
    cv2.rectangle(animated, (0, 130), (1380, 960), (255, 255, 255), -1)

    rows = [
        ((20, 140, 1300, 395),  2.0),
        ((20, 400, 1300, 555),  9.0),
        ((20, 560, 1300, 715), 18.0),
        ((20, 720, 1300, 880), 26.0),
    ]

    for box, start_time in rows:
        x1, y1, x2, y2 = box
        progress = ease_out((local_t - start_time) / 1.1)
        if progress <= 0:
            continue

        row_img  = original[y1:y2, x1:x2].copy()
        start_x  = x1 - 260
        current_x = int(start_x + (x1 - start_x) * progress)
        paste_x1  = max(0, current_x)
        paste_x2  = min(WIDTH, current_x + row_img.shape[1])
        src_x1    = paste_x1 - current_x
        src_x2    = src_x1 + (paste_x2 - paste_x1)
        if paste_x2 > paste_x1 and src_x2 > src_x1:
            animated[y1:y2, paste_x1:paste_x2] = row_img[:, src_x1:src_x2]

    return animated


def animate_slide_6(slide, t):
    local_t  = t - 103
    original = slide.copy()
    animated = slide.copy()

    hx1, hy1, hx2, hy2 = 40, 30, 1500, 170
    cx1, cy1, cx2, cy2 = 55, 250, 980, 760
    heading_img = original[hy1:hy2, hx1:hx2].copy()

    cv2.rectangle(animated, (hx1, hy1), (hx2, hy2), (255, 255, 255), -1)
    cv2.rectangle(animated, (cx1, cy1), (cx2, cy2), (255, 255, 255), -1)

    h_progress = ease_out((local_t - 0.15) / 1.3)
    if h_progress > 0:
        h_cur_x = int((hx1 - WIDTH) + WIDTH * h_progress)
        hp1 = max(0, h_cur_x)
        hp2 = min(WIDTH, h_cur_x + heading_img.shape[1])
        hs1 = hp1 - h_cur_x
        hs2 = hs1 + (hp2 - hp1)
        if hp2 > hp1 and hs2 > hs1:
            animated[hy1:hy2, hp1:hp2] = heading_img[:, hs1:hs2]

    if local_t >= 2.0:
        animated[cy1:cy2, cx1:cx2] = original[cy1:cy2, cx1:cx2]

    return animated


# ---------------------------------------------------------------------------
# Green-spill removal
# ---------------------------------------------------------------------------
def remove_green_tint(img, alpha):
    """
    Three-pass green-spill suppression.

    Pass 1: global subtract — reduces green channel proportional to its
            excess over max(R, B).
    Pass 2: hard edge clamp — in the semi-transparent fringe zone the green
            channel is hard-clamped to max(R, B) so no green hue survives
            on hair or shoulder outlines.
    Pass 3: bilateral smooth on the fringe for clean alpha blending.
    """
    img_f = img.astype(np.float32)
    b, g, r = cv2.split(img_f)

    # Pass 1 — global spill reduction (boosted multiplier for denser shoulder spill)
    spill = np.clip(g - np.maximum(r, b), 0.0, 200.0)
    g = g - spill * 4.0    # was 2.8; higher = more aggressive green removal
    r = r + spill * 0.30
    b = b + spill * 0.08
    out = np.clip(cv2.merge([b, g, r]), 0, 255)

    # Pass 2 — edge zone: hard clamp and full fringe neutralisation
    edge_zone = (alpha > 2) & (alpha < 253)
    if edge_zone.any():
        b2, g2, r2 = cv2.split(out)
        avg_rb = (r2.astype(np.float32) + b2.astype(np.float32)) * 0.5
        max_rb = np.maximum(r2, b2)
        # Clamp to max(R,B) first
        g2[edge_zone] = np.minimum(g2[edge_zone], max_rb[edge_zone])
        # Then neutralise any remaining excess vs avg(R,B)
        excess = np.clip(g2[edge_zone].astype(np.float32) - avg_rb[edge_zone], 0.0, None)
        g2[edge_zone] = np.clip(g2[edge_zone].astype(np.float32) - excess, 0, 255).astype(np.uint8)
        out = cv2.merge([b2, g2, r2])

    out = np.clip(out, 0, 255).astype(np.uint8)

    # Pass 3
    edge_mask = ((alpha > 4) & (alpha < 248)).astype(np.uint8) * 255
    edge_mask = cv2.dilate(edge_mask, np.ones((7, 7), np.uint8), iterations=1)
    if edge_mask.any():
        smooth = cv2.bilateralFilter(out, 9, 45, 45)
        out[edge_mask > 0] = smooth[edge_mask > 0]

    return out


def _correct_green_spill(frame_bgr, alpha_u8):
    """
    Two-pass green-spill correction for the faculty matte.

    Pass 1 — broad sweep (ALL foreground pixels):
        Detect any pixel where green exceeds either R or B by ≥1 %.
        Clamp green to max(R, B).  Low threshold catches the subtle halo
        that the old 1.05× threshold missed on shoulder / hair outlines.

    Pass 2 — fringe zone (semi-transparent 1–239 alpha only):
        Replace all excess green with avg(R, B) — stronger than Pass 1.
        These are the edge pixels that create the visible green shadow;
        blending them onto the slide without full despill causes the halo.
        A small fraction of the excess is transferred to R and B to keep
        the skin/fabric tone warm rather than going cold/magenta.
    """
    out   = frame_bgr.astype(np.float32)
    b_ch  = out[:, :, 0]
    g_ch  = out[:, :, 1]
    r_ch  = out[:, :, 2]
    max_rb = np.maximum(r_ch, b_ch)
    avg_rb = (r_ch + b_ch) * 0.5

    # Pass 1: any foreground pixel where green is even slightly dominant
    spill = (alpha_u8 > 0) & (
        (g_ch > r_ch * 1.01) | (g_ch > b_ch * 1.01)
    )
    if spill.any():
        g_ch[spill] = np.minimum(g_ch[spill], max_rb[spill])

    # Pass 2: fringe-zone full neutralisation — replaces green with avg(R,B)
    fringe = (alpha_u8 > 0) & (alpha_u8 < 240)
    if fringe.any():
        excess = np.clip(g_ch[fringe] - avg_rb[fringe], 0.0, None)
        g_ch[fringe] -= excess
        r_ch[fringe] += excess * 0.20   # warm compensation
        b_ch[fringe] += excess * 0.08

    out[:, :, 0] = np.clip(b_ch, 0, 255)
    out[:, :, 1] = np.clip(g_ch, 0, 255)
    out[:, :, 2] = np.clip(r_ch, 0, 255)
    return out.astype(np.uint8)


# ---------------------------------------------------------------------------
# Main faculty extraction
# ---------------------------------------------------------------------------
def ai_cut_faculty(frame, frame_no=0):
    """
    High-quality green-screen extraction: MediaPipe/TFLite + chroma-key hybrid.

    Pipeline:
      A  ML mask  (MediaPipe or TFLite) -> accurate full-body person mask.
         Falls back to pure chroma key if no ML backend is available.
      B  Hard green override  -> definitively-green pixels forced to 0,
         regardless of what the ML mask says.
      C  Soft spill suppression -> graduated alpha in the hair/fringe zone.
      D  Largest-blob filter  -> drop any stray disconnected regions.
      E  Guided filter  -> edge-preserving alpha refinement (NO Gaussian blur
         -> no haze, no blurred outlines).
      F  Hard clamp  -> push near-0 / near-1 to exact 0 / 255 to kill halos.
    """
    global _rembg_cache

    if (
        _rembg_cache["alpha"] is not None
        and (frame_no - _rembg_cache["at_frame"]) < REMBG_INTERVAL
    ):
        # Per-frame EMA: smoothly advance displayed alpha toward target even on
        # cache-hit frames — eliminates the periodic pop.
        _step_display_alpha()

        # ── HARD-ZERO: rembg background pixels + green pixels ───────────────
        # Root cause of motion shadow on body/head movement:
        # The cached target is STALE between rembg updates, so pixels at the
        # OLD head/body position still have target=255 and EMA steps TOWARD
        # them instead of clearing them.
        #
        # Fix 1 — rembg target zero: any pixel the last rembg run marked as
        #   background (target==0) is immediately forced to 0 in display,
        #   regardless of EMA state.  This works for ALL background colours
        #   (white slide, coloured slide) not just green screen.
        #
        # Fix 2 — green hard-zero: also force 0 wherever the current frame
        #   shows definite green screen, in case the person has moved into a
        #   region rembg previously called "person".
        if _alpha_display["alpha"] is not None:
            # Fix 1: rembg target background → immediate zero
            cached_t = _alpha_target["alpha"]
            if cached_t is not None:
                _alpha_display["alpha"][cached_t == 0] = 0

            # Fix 2: current-frame green screen → immediate zero
            hsv_f = cv2.cvtColor(frame, cv2.COLOR_BGR2HSV)
            H_ch  = hsv_f[:, :, 0]
            S_ch  = hsv_f[:, :, 1]
            g_ch  = frame[:, :, 1].astype(np.float32)
            r_ch  = frame[:, :, 2].astype(np.float32)
            b_ch  = frame[:, :, 0].astype(np.float32)
            green_bg = (
                (H_ch >= 38) & (H_ch <= 87) &
                (S_ch >= 60) &
                (g_ch > r_ch * 1.15) & (g_ch > b_ch * 1.15)
            )
            if green_bg.any():
                _alpha_display["alpha"][green_bg] = 0
        # ────────────────────────────────────────────────────────────────────

        disp = _alpha_display["alpha"] if _alpha_display["alpha"] is not None \
               else _rembg_cache["alpha"]
        bgr_cut = _correct_green_spill(frame, disp)
        return bgr_cut, disp

    h, w = frame.shape[:2]

    # Tier 0: rembg U2Net — highest quality, smooth hair/shoulder edges.
    # Uses deep-learning matting; no green-screen assumption needed.
    if _REMBG_AVAILABLE:
        session = _get_rembg_session()
        if session is not None:
            try:
                pil_in  = _PIL_Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
                pil_out = _rembg_remove(pil_in, session=session)
                rgba    = np.array(pil_out)          # H x W x RGBA
                alpha_u8 = rgba[:, :, 3]
                # Merge stable (pooled) table with dynamic person alpha.
                # _stable_table_merge() internally calls _add_table_to_alpha()
                # but max-pools the result so the table never shakes with the
                # person — see function docstring for full explanation.
                alpha_u8 = _stable_table_merge(alpha_u8, frame)
                # NO matte contraction for rembg path — U2Net already produces
                # clean sub-pixel hair edges. Eroding shrinks those fine hair
                # strands to zero which creates the dark shadow outline on hair.
                # Update target; advance display one EMA step this frame.
                _alpha_target["alpha"] = alpha_u8.copy()
                _step_display_alpha()
                disp_alpha = _alpha_display["alpha"]
                # IMPORTANT: use original frame for BGR colors — NOT rembg's
                # processed RGBA output. This makes rembg-update frames and
                # cache-hit frames use the same color source, eliminating the
                # periodic color pop that appeared every REMBG_INTERVAL frames.
                bgr_out  = _correct_green_spill(frame, disp_alpha)
                _rembg_cache["bgr_cut"]  = bgr_out
                _rembg_cache["alpha"]    = alpha_u8   # cache RAW target for next interval
                _rembg_cache["at_frame"] = frame_no
                gc.collect()
                return bgr_out, disp_alpha
            except Exception as e:
                print(f"[WARN] rembg inference failed: {e}")

    hsv = cv2.cvtColor(frame, cv2.COLOR_BGR2HSV).astype(np.float32)
    H, S, V = hsv[:, :, 0], hsv[:, :, 1], hsv[:, :, 2]

    b = frame[:, :, 0].astype(np.float32)
    g = frame[:, :, 1].astype(np.float32)
    r = frame[:, :, 2].astype(np.float32)

    # A. Primary mask
    ml_mask = _person_mask(frame)

    if ml_mask is not None:
        alpha_f = ml_mask.copy()
    else:
        # Chroma-key fallback
        hard_green_fb = (
            (H >= 38) & (H <= 87) & (S >= 55) & (V >= 40) &
            (g > r * 1.12) & (g > b * 1.12)
        )
        alpha_f = np.ones((h, w), dtype=np.float32)
        alpha_f[hard_green_fb] = 0.0

        max_rb_fb    = np.maximum(r, b)
        g_excess_fb  = np.clip(g - max_rb_fb, 0.0, None)
        fringe = (
            (H >= 33) & (H <= 93) & (S >= 22) &
            (g > r * 1.03) & (g > b * 1.03)
        ) & ~hard_green_fb
        if fringe.any():
            alpha_f[fringe] = np.minimum(
                alpha_f[fringe],
                np.clip(1.0 - g_excess_fb[fringe] / 28.0, 0.0, 1.0)
            )

    # B. Hard green override (always remove definite green-screen pixels)
    hard_green = (
        (H >= 38) & (H <= 87) &
        (S >= 55) & (V >= 40) &
        (g > r * 1.10) & (g > b * 1.10)
    )
    alpha_f[hard_green] = 0.0

    # C. Soft spill suppression (green light on hair / fringe edges)
    max_rb   = np.maximum(r, b)
    g_excess = np.clip(g - max_rb, 0.0, None)
    soft_spill = (
        (H >= 34) & (H <= 90) & (S >= 20) &
        (g > r * 1.04) & (g > b * 1.04)
    ) & ~hard_green
    if soft_spill.any():
        suppress = np.clip(1.0 - g_excess / 26.0, 0.0, 1.0)
        alpha_f[soft_spill] = np.minimum(
            alpha_f[soft_spill], suppress[soft_spill]
        )

    # D. Keep only the largest connected foreground blob.
    # Use a larger MORPH_CLOSE kernel (25×25) to bridge any gap between the
    # upper-body blob and the podium/table blob before connected-component
    # analysis — ensures they are treated as one region.
    core_u8 = (alpha_f > 0.30).astype(np.uint8) * 255
    core_u8 = cv2.morphologyEx(
        core_u8, cv2.MORPH_CLOSE,
        np.ones((25, 25), np.uint8), iterations=2
    )
    num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(
        core_u8, connectivity=8
    )
    if num_labels > 1:
        largest   = 1 + np.argmax(stats[1:, cv2.CC_STAT_AREA])
        person_u8 = (labels == largest).astype(np.uint8) * 255
        # Larger dilation (50x50) so the podium area, even if slightly
        # disconnected from the body, is pulled into the valid region.
        person_u8 = cv2.dilate(
            person_u8, np.ones((50, 50), np.uint8), iterations=1
        )
        alpha_f[person_u8 == 0] = 0.0

    # E. Guided filter: edge-preserving refinement
    gray    = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY).astype(np.float32) / 255.0
    alpha_f = guided_filter(gray, alpha_f, r=8, eps=0.001)

    # F. Hard clamp
    alpha_f[alpha_f < 0.10] = 0.0
    alpha_f[alpha_f > 0.80] = 1.0

    # Professional matte contraction: erode x2 then dilate x1.
    _ero_k     = np.ones((3, 3), np.uint8)
    _alpha_tmp = (alpha_f * 255).astype(np.uint8)
    _alpha_tmp = cv2.erode (_alpha_tmp, _ero_k, iterations=2)
    _alpha_tmp = cv2.dilate(_alpha_tmp, _ero_k, iterations=1)
    alpha_f    = _alpha_tmp.astype(np.float32) / 255.0

    alpha_u8 = (alpha_f * 255).astype(np.uint8)
    # Merge stable (pooled) table with dynamic person alpha.
    alpha_u8 = _stable_table_merge(alpha_u8, frame)
    # Update target; advance display one EMA step this frame.
    _alpha_target["alpha"] = alpha_u8.copy()
    _step_display_alpha()
    disp_alpha = _alpha_display["alpha"]
    bgr_cut  = remove_green_tint(frame, disp_alpha)

    # Cache RAW target (not display) so the cache interval logic stays correct.
    _rembg_cache = {"bgr_cut": None, "alpha": alpha_u8, "at_frame": frame_no}
    gc.collect()
    return bgr_cut, disp_alpha


def prepare_faculty(frame, frame_no=0):
    global _bbox_smooth
    faculty, alpha = ai_cut_faculty(frame, frame_no)

    # Crop to the bounding box of the person
    ys, xs = np.where(alpha > 20)
    if len(xs) > 0:
        raw_x1 = float(max(0, int(xs.min()) - 30))
        raw_x2 = float(min(alpha.shape[1], int(xs.max()) + 30))
        raw_y1 = float(max(0, int(ys.min()) - 25))
        raw_y2 = float(min(alpha.shape[0], int(ys.max()) + 60))

        # Asymmetric y2 EMA: expand fast, shrink very slowly
        if _bbox_smooth["x1"] is None:
            _bbox_smooth["x1"] = raw_x1
            _bbox_smooth["x2"] = raw_x2
            _bbox_smooth["y1"] = raw_y1
            _bbox_smooth["y2"] = raw_y2
        else:
            _bbox_smooth["x1"] = BBOX_SMOOTH * _bbox_smooth["x1"] + (1 - BBOX_SMOOTH) * raw_x1
            _bbox_smooth["x2"] = BBOX_SMOOTH * _bbox_smooth["x2"] + (1 - BBOX_SMOOTH) * raw_x2
            _bbox_smooth["y1"] = BBOX_SMOOTH * _bbox_smooth["y1"] + (1 - BBOX_SMOOTH) * raw_y1
            if raw_y2 >= _bbox_smooth["y2"]:
                _bbox_smooth["y2"] = 0.70 * _bbox_smooth["y2"] + 0.30 * raw_y2
            else:
                _bbox_smooth["y2"] = 0.995 * _bbox_smooth["y2"] + 0.005 * raw_y2

        x1 = int(_bbox_smooth["x1"])
        x2 = int(_bbox_smooth["x2"])
        y1 = int(_bbox_smooth["y1"])
        y2 = int(_bbox_smooth["y2"])

        faculty = faculty[y1:y2, x1:x2]
        alpha   = alpha[y1:y2, x1:x2]

    # Remove truly empty rows at the bottom.
    row_max      = alpha.max(axis=1)
    content_rows = np.where(row_max > 2)[0]
    if len(content_rows) > 0:
        faculty = faculty[:content_rows[-1] + 1]
        alpha   = alpha[:content_rows[-1] + 1]

    if faculty.shape[0] == 0:
        blank = np.zeros((FACULTY_HEIGHT, 1, 3), dtype=np.uint8)
        return blank, np.zeros((FACULTY_HEIGHT, 1, 3), dtype=np.float32)
    scale         = FACULTY_HEIGHT / faculty.shape[0]
    faculty_width = int(faculty.shape[1] * scale)
    faculty = cv2.resize(faculty, (faculty_width, FACULTY_HEIGHT), interpolation=cv2.INTER_LANCZOS4)
    alpha   = cv2.resize(alpha,   (faculty_width, FACULTY_HEIGHT), interpolation=cv2.INTER_LANCZOS4)

    alpha_f = np.clip(alpha, 0, 255).astype(np.float32) / 255.0
    # Ramp: 0.20→0, 0.70→1.0. Wide enough to preserve rembg's semi-transparent
    # hair strands (alpha ~0.3–0.5) while still cutting hard background noise.
    alpha_f = np.clip((alpha_f - 0.20) / 0.50, 0.0, 1.0)

    # Feather the edge zone for a smooth, shadow-free outline.
    # A 7×7 Gaussian blur is applied only to fringe pixels (0.01–0.99)
    # so the fully opaque body and fully transparent background are untouched.
    alpha_blur = cv2.GaussianBlur(alpha_f, (7, 7), 1.5)
    edge_zone  = (alpha_f > 0.01) & (alpha_f < 0.99)
    alpha_f    = np.where(edge_zone, alpha_blur, alpha_f)

    alpha_3 = cv2.merge([alpha_f, alpha_f, alpha_f])

    return faculty, alpha_3


def place_faculty(slide, frame, position, frame_no=0):
    faculty, alpha_3 = prepare_faculty(frame, frame_no)

    faculty_h, faculty_w = faculty.shape[:2]

    if position == "center":
        x = (WIDTH - faculty_w) // 2
    else:
        x = WIDTH - faculty_w - RIGHT_SHIFT

    y = HEIGHT - faculty_h

    dst_x1 = max(0, x)
    dst_y1 = max(0, y)
    dst_x2 = min(WIDTH,  x + faculty_w)
    dst_y2 = min(HEIGHT, y + faculty_h)

    src_x1 = max(0, -x)
    src_y1 = max(0, -y)
    src_x2 = src_x1 + (dst_x2 - dst_x1)
    src_y2 = src_y1 + (dst_y2 - dst_y1)

    if dst_x2 <= dst_x1 or dst_y2 <= dst_y1:
        return slide

    roi          = slide[dst_y1:dst_y2, dst_x1:dst_x2]
    faculty_crop = faculty[src_y1:src_y2, src_x1:src_x2]
    alpha_crop   = alpha_3[src_y1:src_y2, src_x1:src_x2]

    blended = (
        faculty_crop.astype(np.float32) * alpha_crop
        + roi.astype(np.float32) * (1.0 - alpha_crop)
    )
    slide[dst_y1:dst_y2, dst_x1:dst_x2] = np.clip(blended, 0, 255).astype(np.uint8)

    return slide


# ---------------------------------------------------------------------------
def create_video(slide_paths):
    slides = []
    for path in slide_paths:
        img = cv2.imread(path)
        if img is None:
            raise FileNotFoundError(f"Slide not found: {path}")
        slides.append(cv2.resize(img, (WIDTH, HEIGHT), interpolation=cv2.INTER_AREA))

    cap = cv2.VideoCapture(FACULTY_VIDEO)
    if not cap.isOpened():
        raise FileNotFoundError("Cannot open faculty video")

    faculty_fps = cap.get(cv2.CAP_PROP_FPS)
    if faculty_fps <= 0:
        faculty_fps = 30.0

    out = cv2.VideoWriter(
        TEMP_VIDEO,
        cv2.VideoWriter_fourcc(*"mp4v"),
        OUTPUT_FPS,
        (WIDTH, HEIGHT),
    )

    total_frames = int(TOTAL_DURATION * OUTPUT_FPS)

    _fac_last_idx = -1
    _fac_frame    = None

    for frame_no in range(total_frames):
        t = frame_no / OUTPUT_FPS

        slide_index, position = get_slide_info(t)
        slide = slides[slide_index].copy()

        if slide_index == 1:
            slide = animate_slide_2(slide, t)

        if position != "hide":
            faculty_time = t - AUDIO_START_TIME
            if faculty_time >= 0:
                target = int(round(faculty_time * faculty_fps))
                while _fac_last_idx < target:
                    _ret, _fframe = cap.read()
                    if not _ret:
                        break
                    _fac_frame    = _fframe
                    _fac_last_idx += 1
                if _fac_frame is not None:
                    slide = place_faculty(slide, _fac_frame, position, frame_no)

        out.write(slide)

        if frame_no % 30 == 0:
            gc.collect()
        if frame_no % 100 == 0:
            print(f"Processing frame {frame_no}/{total_frames}")

    cap.release()
    out.release()


def add_audio():
    """
    Combine temp_video + faculty audio with TWO hard cuts:
      Cut 1: t = 7-13 s, Cut 2: t = 96-103 s
    Output duration: 7 + 83 + 30 = 120 s (exactly 2:00)
    """
    delay_ms = int(AUDIO_START_TIME * 1000)
    C1S, C1E = 7.0,  13.0
    C2S, C2E = 96.0, 103.0

    filter_complex = (
        "[0:v]split=3[vin1][vin2][vin3];"
        f"[vin1]trim=0:{C1S},setpts=PTS-STARTPTS[v1];"
        f"[vin2]trim={C1E}:{C2S},setpts=PTS-STARTPTS[v2];"
        f"[vin3]trim={C2E},setpts=PTS-STARTPTS[v3];"
        f"[1:a]adelay={delay_ms}|{delay_ms},asplit=3[a_del1][a_del2][a_del3];"
        f"[a_del1]atrim=0:{C1S},asetpts=PTS-STARTPTS[a1];"
        f"[a_del2]atrim={C1E}:{C2S},asetpts=PTS-STARTPTS[a2];"
        f"[a_del3]atrim={C2E},asetpts=PTS-STARTPTS[a3];"
        "[v1][v2][v3]concat=n=3:v=1:a=0[vout];"
        "[a1][a2][a3]concat=n=3:v=0:a=1[aout]"
    )

    run_cmd([
        FFMPEG_PATH, "-y",
        "-i", TEMP_VIDEO,
        "-i", FACULTY_VIDEO,
        "-filter_complex", filter_complex,
        "-map", "[vout]",
        "-map", "[aout]",
        "-c:v", "libx264",
        "-preset", "ultrafast",
        "-crf", "20",
        "-pix_fmt", "yuv420p",
        "-r", str(OUTPUT_FPS),
        "-c:a", "aac",
        FINAL_VIDEO,
    ])


if __name__ == "__main__":
    print("Converting PPT to images...")
    slide_paths = convert_ppt_to_images()

    print("Creating synced video...")
    create_video(slide_paths)

    print("Adding audio...")
    add_audio()

    print("Completed successfully!")
    print("Output:", FINAL_VIDEO)
