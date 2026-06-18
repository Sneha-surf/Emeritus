#!/usr/bin/env python3
"""
POC 3 — Stage 4: Narrated Video Generation

Converts the storyboard into a narrated MP4:
  1. PPTX → PNG frames  (LibreOffice + pymupdf preferred; Pillow fallback)
  2. Speaker notes → MP3 (edge-tts primary; gTTS fallback)
  3. frame + audio → video clip  (moviepy)
  4. clips concatenated → storyboard_video.mp4

Usage:
  python poc3_stage4_video.py <storyboard_json> <pptx_path> <output_dir>

Environment:
  TTS_VOICE — Microsoft Edge neural voice (default: en-US-AriaNeural)
"""

import sys
import json
import os
import subprocess
import textwrap
from pathlib import Path

# ── Brand colours (RGB for Pillow) ────────────────────────────────────────────
_DARK  = (26,  26,  46)
_BLUE  = (22,  33,  62)
_GOLD  = (233, 79,  55)
_WHITE = (255, 255, 255)
_LGREY = (230, 230, 230)
_MGREY = (140, 140, 140)

W, H = 1920, 1080   # output resolution

COVER_DUR         = 4.0   # seconds when no audio
DIVIDER_DUR       = 2.5
PAD               = 0.40  # silence padding after each regular audio clip
TRANSITION_DUR    = 0.40  # crossfade between slides — equals PAD so audio never collides
BULLET_PAD        = 0.30  # pause after each bullet narration chunk (was 0.05 — too abrupt)
BULLET_TRANS      = 0.0   # hard cut between bullet reveal frames — 0.4s FadeOut caused a "blink"
BULLET_SILENT_DUR = 0.5   # hold for silent inter-bullet frames (no narration chunk)
                       # because the next frame appears at full brightness while the outgoing
                       # frame is still fading, creating a visible flash rather than a blend.
TTS_VOICE      = "en-US-AriaNeural"  # Microsoft neural voice via edge-tts (free)

# Animation intro: Header (Wipe L→R) → Image (Fade In) → OSTs (Wipe L→R)
_WIPE_DUR    = 0.25   # wipe duration for intro header; bullet wipes use ffmpeg blend (audio continuous)
_FADE_DUR    = 0.40   # image fade-in duration (PIL, silent intro)
_TRANS_FPS   = 24     # frames per second for PIL transition sequences
_HEADER_HOLD = 0.20   # static hold after header wipe finishes
_IMAGE_HOLD  = 0.20   # static hold after image fade finishes
_BULLET_WIPE = 0.30   # wipe duration inside each bullet/row ffmpeg-blend segment
_INTRO_HOLD  = _HEADER_HOLD + _IMAGE_HOLD


# ── Font loader ────────────────────────────────────────────────────────────────
# Candidate paths ordered: Calibri (template font) → Helvetica Neue → Arial →
# Liberation Sans (Linux equivalent of Arial) → DejaVu (last-resort).
# Calibri matches the POC_3.pptx template typeface for visual consistency.

_FONT_CACHE: dict = {}

def _font(size, bold=False):
    from PIL import ImageFont
    key = (size, bold)
    if key in _FONT_CACHE:
        return _FONT_CACHE[key]

    # Priority: Avenir Next (closest macOS match to Montserrat — the POC_3.pptx body font)
    # → Calibri → Helvetica Neue → Arial → Liberation Sans → DejaVu (last-resort)
    bold_candidates = [
        ("/System/Library/Fonts/Avenir Next.ttc",    {"index": 7}),  # DemiBold
        ("/Library/Fonts/Montserrat-Bold.ttf",       {}),
        ("/Library/Fonts/Calibri Bold.ttf",          {}),
        ("/System/Library/Fonts/HelveticaNeue.ttc",  {"index": 3}),
        ("/System/Library/Fonts/Helvetica.ttc",      {"index": 1}),
        ("/Library/Fonts/Arial Bold.ttf",            {}),
        ("/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf", {}),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",         {}),
    ]
    reg_candidates = [
        ("/System/Library/Fonts/Avenir Next.ttc",    {"index": 0}),  # Regular
        ("/Library/Fonts/Montserrat-Regular.ttf",    {}),
        ("/Library/Fonts/Calibri.ttf",               {}),
        ("/System/Library/Fonts/HelveticaNeue.ttc",  {"index": 0}),
        ("/System/Library/Fonts/Helvetica.ttc",      {"index": 0}),
        ("/Library/Fonts/Arial.ttf",                 {}),
        ("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", {}),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",                 {}),
    ]
    fnt = ImageFont.load_default()
    for path, kw in (bold_candidates if bold else reg_candidates):
        if Path(path).exists():
            try:
                fnt = ImageFont.truetype(path, size, **kw)
                break
            except Exception:
                pass
    _FONT_CACHE[key] = fnt
    return fnt


def _draw_wrapped(draw, text, x, y, font, color, max_w, spacing=1.35):
    try:
        avg_w = max(1, font.getlength("n"))
    except Exception:
        avg_w = 12
    cpw = max(1, int(max_w / avg_w))
    lines = []
    for para in text.split("\n"):
        lines += textwrap.wrap(para, cpw) if para.strip() else [""]
    try:
        lh = font.size * spacing
    except Exception:
        lh = 18 * spacing
    cy = y
    for ln in lines:
        draw.text((x, cy), ln, font=font, fill=color)
        cy += lh
    return cy


# ── Pillow slide renderers ─────────────────────────────────────────────────────

def _bg_image(img_path, darken=0.55):
    from PIL import Image
    if img_path and Path(img_path).exists():
        bg = Image.open(img_path).convert("RGB").resize((W, H), Image.LANCZOS)
        dark = Image.new("RGB", (W, H), (0, 0, 0))
        return Image.blend(bg, dark, darken)
    return Image.new("RGB", (W, H), _DARK)


def _render_cover(title, img_path, out_path):
    from PIL import Image, ImageDraw
    bg = _bg_image(img_path, darken=0.50)
    draw = ImageDraw.Draw(bg)
    draw.rectangle([0, H - 130, W, H], fill=_DARK)
    draw.rectangle([0, H - 137, W // 4, H - 131], fill=_GOLD)
    _draw_wrapped(draw, title, 80, int(H * 0.28), _font(80, bold=True), _WHITE, W - 160)
    draw.text((80, H - 90), "AI Learning Module", font=_font(28), fill=_LGREY)
    bg.save(str(out_path))
    return str(out_path)


def _render_divider(sec_num, heading, out_path):
    from PIL import Image, ImageDraw
    bg = Image.new("RGB", (W, H), _DARK)
    draw = ImageDraw.Draw(bg)
    draw.rectangle([0, 0, 14, H], fill=_GOLD)
    draw.text((80, int(H * 0.38)), f"SECTION {sec_num}",
              font=_font(30, bold=True), fill=_GOLD)
    _draw_wrapped(draw, heading, 80, int(H * 0.46),
                  _font(68, bold=True), _WHITE, W - 200)
    bg.save(str(out_path))
    return str(out_path)


def _render_content(scene_type, title, bullets, img_path, out_path):
    from PIL import Image, ImageDraw
    bg = _bg_image(img_path, darken=0.52)

    # Semi-transparent text band at the bottom 40%
    overlay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    ov = ImageDraw.Draw(overlay)
    band_top = int(H * 0.58)
    ov.rectangle([0, band_top, W, H], fill=(18, 18, 36, 220))
    bg = bg.convert("RGBA")
    bg.alpha_composite(overlay)
    bg = bg.convert("RGB")
    draw = ImageDraw.Draw(bg)

    # Gold accent line
    draw.rectangle([0, band_top, W, band_top + 5], fill=_GOLD)

    # Title
    title_y = band_top + 26
    _draw_wrapped(draw, title, 80, title_y, _font(50, bold=True), _WHITE, W - 160)

    # Bullets / quote
    if scene_type == "quote" and bullets:
        _draw_wrapped(draw, f"“{bullets[0]}”",
                      100, band_top + 120, _font(34), _WHITE, W - 220, spacing=1.4)
    elif bullets:
        by = band_top + 115
        for b in bullets[:5]:
            by = _draw_wrapped(draw, f"•  {b}", 100, by,
                               _font(30), _LGREY, W - 220)
            by += 18

    bg.save(str(out_path))
    return str(out_path)


def _split_vo_for_bullets(notes, n):
    """Split narration text into n sentence-chunks for per-bullet TTS."""
    import re, math
    if not notes or not notes.strip() or n <= 1:
        return [notes or ""] * max(n, 1)
    sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', notes.strip()) if s.strip()]
    if not sentences:
        return [notes] * n
    if len(sentences) == n:
        return sentences
    if len(sentences) > n:
        per = len(sentences) / n
        chunks, start = [], 0
        for i in range(n):
            end = math.ceil((i + 1) * per) if i < n - 1 else len(sentences)
            chunks.append(" ".join(sentences[start:end]))
            start = end
        return chunks
    # Fewer sentences than bullets: distribute text uniformly so every bullet gets narration
    text = notes.strip()
    L    = len(text)
    size = max(1, L // n)
    return [text[k * size:(k + 1) * size if k < n - 1 else L].strip() for k in range(n)]


def _bullet_slice_times(items, word_boundaries, total_dur):
    """
    Return (start_sec, dur_sec) per item using word-level boundaries for precision.
    Falls back to character proportion when a keyword isn't found.

    items          — list of bullet/row heading strings
    word_boundaries — list of {"text", "start", "end"} from _tts_edge_with_boundaries
    total_dur      — total audio duration in seconds
    """
    import re as _re
    n = len(items)
    if not n:
        return []

    def _norm(s):
        return _re.sub(r"[^a-z0-9]", "", s.lower())

    starts   = []
    last_wi  = 0          # only search forward so order is preserved

    for item in items:
        # Split bullet text (treat hyphens as spaces so "one-size-fits-all" → 4 words
        # and the short first word "no" is accessible as raw_words[0])
        raw_words = _re.sub(r"[^a-z0-9\s]", "", item.lower()).split()
        if not raw_words:
            raw_words = ["x"]

        found = None
        # Try decreasing phrase lengths (3→2→1 words); skip multi-word phrases
        # composed entirely of 2-char words (too ambiguous) unless it's the last try
        for nw in range(min(3, len(raw_words)), 0, -1):
            phrase = [_norm(w) for w in raw_words[:nw]]
            if nw > 1 and all(len(p) <= 2 for p in phrase):
                continue
            for wi in range(last_wi, len(word_boundaries)):
                if wi + nw > len(word_boundaries):
                    break
                span = [_norm(word_boundaries[wi + j]["text"]) for j in range(nw)]
                if span == phrase:
                    found   = word_boundaries[wi]["start"]
                    last_wi = wi + nw
                    break
            if found is not None:
                break

        if found is None:
            # Proportional fallback
            prev   = starts[-1] if starts else 0.0
            remain = max(0.1, total_dur - prev)
            found  = prev + remain / max(1, n - len(starts)) * 0.15

        starts.append(found)

    # Guarantee strictly increasing
    for i in range(1, n):
        if starts[i] <= starts[i - 1]:
            starts[i] = starts[i - 1] + 0.05

    durations = [max(0.3, starts[i + 1] - starts[i]) for i in range(n - 1)]
    durations.append(max(0.5, total_dur - starts[-1]))

    return list(zip(starts, durations))


def _bullet_keyword_times(narration_text, bullets, total_duration, offset=0.0):
    """
    Return (start_sec, clip_duration_sec) for each bullet based on where its
    keyword first appears in narration_text (character-position proxy for time).

    offset — seconds already consumed by intro animation (header wipe + image fade).
    Bullet 0 starts at `offset` so intro audio plays during the header animation
    rather than silently.  Each subsequent bullet is proportional within the
    remaining [offset, total_duration] window.
    """
    if not narration_text or not bullets:
        available = max(0.1, total_duration - offset)
        d = available / max(len(bullets), 1)
        return [(offset + i * d, d) for i in range(len(bullets))]

    text_lower   = narration_text.lower()
    total_chars  = max(len(narration_text), 1)
    n            = len(bullets)
    search_from  = 0
    positions    = []

    for bullet in bullets:
        words = bullet.lower().split()
        pos   = -1
        # Try decreasing phrase lengths to find the keyword in the narration
        for nw in range(min(3, len(words)), 0, -1):
            phrase = " ".join(words[:nw])
            if len(phrase) < 3:
                continue
            found = text_lower.find(phrase, search_from)
            if found >= 0:
                pos = found
                search_from = found + len(phrase)
                break
        if pos < 0:
            pos = search_from
            search_from += 1
        positions.append(pos)

    # Guarantee strictly increasing (preserves appearance order)
    for i in range(1, n):
        if positions[i] <= positions[i - 1]:
            positions[i] = positions[i - 1] + 1

    # Bullet 0 → offset; others → offset + proportional within remaining window
    available = max(0.1, total_duration - offset)
    starts = [offset] + [offset + (positions[k] / total_chars) * available
                         for k in range(1, n)]

    # Each frame shows until the next bullet's keyword, last frame until end
    durations = [max(0.3, starts[i + 1] - starts[i]) for i in range(n - 1)]
    durations.append(max(0.5, total_duration - starts[-1]))

    return list(zip(starts, durations))


def _get_animated_bullets(slide):
    """
    If the slide has per-paragraph Appear animation (as built by poc3_build_from_table),
    return a dict with bullets list and the animated shape's bounding box in EMU.
    Returns None if no animation is found.
    """
    from pptx.oxml.ns import qn
    timing = slide._element.find(qn("p:timing"))
    if timing is None:
        return None
    bldLst = timing.find(qn("p:bldLst"))
    if bldLst is None:
        return None
    bldP = bldLst.find(qn("p:bldP"))
    if bldP is None or bldP.get("build") != "p":
        return None
    spid = bldP.get("spid")
    if not spid:
        return None

    spTree = slide._element.find(".//" + qn("p:spTree"))
    if spTree is None:
        return None

    for sp in spTree.iter(qn("p:sp")):
        cNvPr = sp.find(".//" + qn("p:cNvPr"))
        if cNvPr is None or cNvPr.get("id") != spid:
            continue
        xfrm = sp.find(".//" + qn("a:xfrm"))
        if xfrm is None:
            return None
        off = xfrm.find(qn("a:off"))
        ext = xfrm.find(qn("a:ext"))
        if off is None or ext is None:
            return None
        left   = int(off.get("x", 0))
        top    = int(off.get("y", 0))
        width  = int(ext.get("cx", 0))
        height = int(ext.get("cy", 0))
        txBody = sp.find(qn("p:txBody"))
        if txBody is None:
            return None
        bullets = []
        for ap in txBody.findall(qn("a:p")):
            text = "".join(r.text or "" for r in ap.findall(".//" + qn("a:t")))
            if text.strip():
                bullets.append(text.strip())
        if not bullets:
            return None
        return {"bullets": bullets,
                "left": left, "top": top, "width": width, "height": height}
    return None


def _get_bullet_content(slide):
    """
    Content-based bullet detector for slides using the '60:40 Icon and text box
    with bullets' layout (Stage 3 never writes PPTX animation XML, so
    _get_animated_bullets always returns None for those slides).
    Reads bullet text from PH_BODY (idx 41) and returns the same dict format.
    """
    from pptx.oxml.ns import qn

    layout = slide.slide_layout
    if "60:40" not in layout.name:
        return None

    PH_BODY = 41
    for ph in slide.placeholders:
        if ph.placeholder_format.idx != PH_BODY:
            continue
        bullets = [p.text.strip() for p in ph.text_frame.paragraphs if p.text.strip()]
        if not bullets:
            return None
        xfrm = ph._element.find(".//" + qn("a:xfrm"))
        if xfrm is None:
            for lph in layout.placeholders:
                if lph.placeholder_format.idx == PH_BODY:
                    xfrm = lph._element.find(".//" + qn("a:xfrm"))
                    break
        if xfrm is None:
            return None
        off = xfrm.find(qn("a:off"))
        ext = xfrm.find(qn("a:ext"))
        if off is None or ext is None:
            return None
        return {
            "bullets": bullets,
            "left":   int(off.get("x", 0)),
            "top":    int(off.get("y", 0)),
            "width":  int(ext.get("cx", 0)),
            "height": int(ext.get("cy", 0)),
        }
    return None


def _get_icon3_rows_info(slide):
    """
    If the slide uses the '60:40 Icon and text' layout (icon3), return a list of
    {heading, body, h_bbox, b_bbox} dicts (one per detected row).
    h_bbox covers the combined bounding box of the icon circle + both text shapes
    for that row, so the progressive-reveal whiteout erases icon and text together.
    Returns None if the layout doesn't match or exactly 3 rows can't be detected.
    Bounding boxes are (left, top, width, height) in EMU.
    """
    from pptx.oxml.ns import qn

    layout = slide.slide_layout
    name   = layout.name
    if "60:40 Icon and text" not in name:
        return None
    if "bullets" in name or "headings" in name:
        return None

    # Collect all non-placeholder, non-picture shapes with valid geometry
    row_shapes = []
    for sh in slide.shapes:
        try:
            if sh.placeholder_format is not None:
                continue
        except Exception:
            pass
        if sh.shape_type == 13:  # PICTURE
            continue
        xfrm = sh._element.find(".//" + qn("a:xfrm"))
        if xfrm is None:
            continue
        off = xfrm.find(qn("a:off"))
        ext = xfrm.find(qn("a:ext"))
        if off is None or ext is None:
            continue
        top    = int(off.get("y", 0))
        left   = int(off.get("x", 0))
        width  = int(ext.get("cx", 0))
        height = int(ext.get("cy", 0))
        text   = sh.text_frame.text.strip() if sh.shape_type == 17 else ""
        row_shapes.append({
            "top": top, "left": left, "width": width, "height": height,
            "bottom": top + height, "right": left + width,
            "text": text, "is_text": sh.shape_type == 17,
        })

    if len(row_shapes) < 3:
        return None

    row_shapes.sort(key=lambda x: x["top"])

    # Group shapes into rows: shapes whose top falls within current_bottom + tolerance
    # belong to the same row (icon circles and text boxes overlap in Y within a row).
    _TOL = 91440  # 0.1" EMU
    groups, current, current_bottom = [], [row_shapes[0]], row_shapes[0]["bottom"]
    for sh in row_shapes[1:]:
        if sh["top"] <= current_bottom + _TOL:
            current.append(sh)
            current_bottom = max(current_bottom, sh["bottom"])
        else:
            groups.append(current)
            current, current_bottom = [sh], sh["bottom"]
    groups.append(current)

    if len(groups) != 3:
        return None

    result = []
    for grp in groups:
        texts = [sh["text"] for sh in grp if sh["is_text"] and sh["text"]]
        combined = (
            min(sh["left"]   for sh in grp),
            min(sh["top"]    for sh in grp),
            max(sh["right"]  for sh in grp) - min(sh["left"]   for sh in grp),
            max(sh["bottom"] for sh in grp) - min(sh["top"]    for sh in grp),
        )
        result.append({
            "heading": texts[0] if texts else "",
            "body":    texts[1] if len(texts) > 1 else "",
            "h_bbox":  combined,
            "b_bbox":  None,
        })

    return result


def _get_comparison_rows_info(slide):
    """
    If the slide uses the '60:40 Text boxes with headings' layout, return a list of
    {heading, body, h_bbox, b_bbox} dicts (one per row). Else return None.
    Bounding boxes are (left, top, width, height) in EMU from the layout placeholders.
    """
    from pptx.oxml.ns import qn

    layout = slide.slide_layout
    if "Text boxes with headings" not in layout.name:
        return None

    _CMP_ROWS = [(35, 36), (37, 38), (39, 40)]

    def layout_bbox(idx):
        # Prefer slide-level xfrm override (set by _build_comparison_slide for alignment)
        for ph in slide.placeholders:
            if ph.placeholder_format.idx != idx:
                continue
            xfrm = ph._element.find(".//" + qn("a:xfrm"))
            if xfrm is not None:
                off = xfrm.find(qn("a:off"))
                ext = xfrm.find(qn("a:ext"))
                if off is not None and ext is not None:
                    return (int(off.get("x", 0)), int(off.get("y", 0)),
                            int(ext.get("cx", 0)), int(ext.get("cy", 0)))
            break
        # Fall back to layout placeholder geometry
        for ph in layout.placeholders:
            if ph.placeholder_format.idx != idx:
                continue
            xfrm = ph._element.find(".//" + qn("a:xfrm"))
            if xfrm is None:
                return None
            off = xfrm.find(qn("a:off"))
            ext = xfrm.find(qn("a:ext"))
            if off is None or ext is None:
                return None
            return (int(off.get("x", 0)), int(off.get("y", 0)),
                    int(ext.get("cx", 0)), int(ext.get("cy", 0)))
        return None

    def slide_text(idx):
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == idx:
                return ph.text_frame.text.strip()
        return ""

    # Collect rounded-rect box shapes sorted top-to-bottom (one per row)
    box_shapes = sorted(
        [sh for sh in slide.shapes
         if sh.shape_type == 1 and "Rounded Rectangle" in sh.name],
        key=lambda s: s.top,
    )

    rows = []
    for i, (h_idx, b_idx) in enumerate(_CMP_ROWS):
        h_bbox = layout_bbox(h_idx)
        b_bbox = layout_bbox(b_idx)
        if h_bbox is None:
            continue
        bs = box_shapes[i] if i < len(box_shapes) else None
        rows.append({
            "heading":  slide_text(h_idx),
            "body":     slide_text(b_idx),
            "h_bbox":   h_bbox,
            "b_bbox":   b_bbox,
            "box_bbox": (bs.left, bs.top, bs.width, bs.height) if bs else None,
        })

    return rows if rows else None


def _get_reflection_info(slide):
    """
    Detect a reflection-question slide: '60:40 blank' layout that contains a
    single Rounded Rectangle (the coloured box) plus a TextBox mentioning
    REFLECTION.  Returns {"box_bbox": (l,t,w,h)} in EMU, or None.
    """
    layout = slide.slide_layout
    if "blank" not in layout.name.lower():
        return None
    has_refl = any(
        hasattr(sh, "text") and "REFLECTION" in sh.text.upper()
        for sh in slide.shapes
    )
    if not has_refl:
        return None
    for sh in slide.shapes:
        if sh.shape_type == 1 and "Rounded Rectangle" in sh.name:
            return {"box_bbox": (sh.left, sh.top, sh.width, sh.height)}
    return None


def _get_picture_bbox(slide):
    """Return (left, top, cx, cy) in EMU of the first picture in the slide, or None."""
    for shape in slide.shapes:
        if shape.shape_type == 13:  # MSO_SHAPE_TYPE.PICTURE
            return (shape.left, shape.top, shape.width, shape.height)
    return None


def _make_intro_frames(lo_frame, img_bbox, content_bboxes, frames_dir, stem, slide_emu):
    """
    Create two intro frames for the Header → Image → OST animation sequence:
      <stem>_title.png  — title only  (image + all content bboxes whited out)
      <stem>_img.png    — title+image (only content bboxes whited out)
    Returns (title_frame_path, image_frame_path).
    """
    from PIL import Image, ImageDraw

    sw, sh = slide_emu

    def to_px(bbox, pad=24):
        left, top, width, height = bbox
        return (max(0, int(left  * W / sw) - pad),
                max(0, int(top   * H / sh) - pad),
                min(W, int((left + width)  * W / sw) + pad),
                min(H, int((top  + height) * H / sh) + pad))

    title_path     = str(frames_dir / f"{stem}_title.png")
    title_img_path = str(frames_dir / f"{stem}_img.png")

    base = Image.open(str(lo_frame)).convert("RGB").resize((W, H), Image.LANCZOS)

    # Title-only: white-out image + all OST content
    t_img = base.copy()
    draw  = ImageDraw.Draw(t_img)
    if img_bbox:
        draw.rectangle(to_px(img_bbox), fill=(255, 255, 255))
    for bbox in content_bboxes:
        if bbox:
            draw.rectangle(to_px(bbox), fill=(255, 255, 255))
    t_img.save(title_path)

    # Title+image: white-out OST content only
    ti_img = base.copy()
    draw2  = ImageDraw.Draw(ti_img)
    for bbox in content_bboxes:
        if bbox:
            draw2.rectangle(to_px(bbox), fill=(255, 255, 255))
    ti_img.save(title_img_path)

    return title_path, title_img_path


def _make_wipe_frames(img_a, img_b, n_frames, frames_dir, prefix):
    """
    Generate n_frames wipe-left-to-right transition PNGs from img_a to img_b.
    Returns list of saved file paths.
    """
    from PIL import Image
    paths = []
    for k in range(n_frames):
        progress = (k + 1) / n_frames
        clip_x   = max(1, int(progress * W))
        result   = img_a.copy()
        result.paste(img_b.crop((0, 0, clip_x, H)), (0, 0))
        p = str(frames_dir / f"{prefix}_wipe{k:02d}.png")
        result.save(p)
        paths.append(p)
    return paths


def _make_fade_frames(img_a, img_b, n_frames, frames_dir, prefix):
    """
    Generate n_frames crossfade transition PNGs from img_a to img_b.
    Returns list of saved file paths.
    """
    from PIL import Image
    paths = []
    for k in range(n_frames):
        alpha  = (k + 1) / n_frames
        result = Image.blend(img_a, img_b, alpha)
        p = str(frames_dir / f"{prefix}_fade{k:02d}.png")
        result.save(p)
        paths.append(p)
    return paths


def _append_wipe(pairs, img_a, img_b, frames_dir, prefix):
    """Add wipe L→R transition frames then a static hold to pairs."""
    from PIL import Image
    n   = max(1, int(_WIPE_DUR * _TRANS_FPS))
    dt  = 1.0 / _TRANS_FPS
    a   = img_a.convert("RGB").resize((W, H), Image.LANCZOS)
    b   = img_b.convert("RGB").resize((W, H), Image.LANCZOS)
    for p in _make_wipe_frames(a, b, n, frames_dir, prefix):
        pairs.append((p, None, dt, 0.0, 0))


def _append_fade(pairs, img_a, img_b, frames_dir, prefix):
    """Add fade transition frames to pairs."""
    from PIL import Image
    n  = max(1, int(_FADE_DUR * _TRANS_FPS))
    dt = 1.0 / _TRANS_FPS
    a  = img_a.convert("RGB").resize((W, H), Image.LANCZOS)
    b  = img_b.convert("RGB").resize((W, H), Image.LANCZOS)
    for p in _make_fade_frames(a, b, n, frames_dir, prefix):
        pairs.append((p, None, dt, 0.0, 0))


def _make_wipe_segment(prev_frame_path, curr_frame_path, audio_path, b_dur,
                       wipe_dur, output_path):
    """
    Build a single MP4 segment where the L→R wipe animation plays simultaneously
    with the audio — no silent gap between animation and narration.

    The wipe completes in wipe_dur seconds; audio plays from T=0 for b_dur seconds.
    Segment duration = max(b_dur, wipe_dur) so the wipe always finishes.
    """
    import subprocess, shutil
    ffmpeg  = shutil.which("ffmpeg") or "/usr/local/bin/ffmpeg"
    dur     = max(b_dur, wipe_dur + 0.05)
    has_aud = audio_path and Path(audio_path).exists()

    # Wipe L→R: at time T, pixels with X ≤ W·min(T,wipe_dur)/wipe_dur show curr (B)
    wd = f"{wipe_dur:.4f}"
    filter_expr = f"if(lte(X,W*min(T,{wd})/{wd}),B,A)"

    cmd = [ffmpeg, "-y", "-loglevel", "error",
           "-loop", "1", "-framerate", "24", "-i", str(prev_frame_path),
           "-loop", "1", "-framerate", "24", "-i", str(curr_frame_path)]

    if has_aud:
        cmd += ["-i", str(audio_path)]
    else:
        cmd += ["-f", "lavfi", "-i", "anullsrc=r=44100:cl=stereo"]

    cmd += [
        "-filter_complex", f"[0][1]blend=all_expr='{filter_expr}'[v]",
        "-map", "[v]", "-map", "2:a",
        "-c:v", "libx264", "-preset", "veryfast", "-pix_fmt", "yuv420p", "-r", "24",
        "-c:a", "aac", "-b:a", "128k", "-ar", "44100", "-ac", "2",
        "-t", f"{dur:.4f}",
        str(output_path),
    ]

    r = subprocess.run(cmd, capture_output=True)
    if r.returncode != 0:
        return None
    return str(output_path)


def _make_fade_segment(prev_frame_path, curr_frame_path, audio_path, b_dur,
                       fade_dur, output_path):
    """
    Build a single MP4 segment where a crossfade animation plays simultaneously
    with the audio.  Fade completes in fade_dur seconds; audio plays from T=0.
    Segment duration = max(b_dur, fade_dur + 0.05).
    """
    import subprocess, shutil
    ffmpeg  = shutil.which("ffmpeg") or "/usr/local/bin/ffmpeg"
    dur     = max(b_dur, fade_dur + 0.05)
    has_aud = audio_path and Path(audio_path).exists()

    # Crossfade: output = A + (B-A) * min(T, fade_dur)/fade_dur
    wd = f"{fade_dur:.4f}"
    filter_expr = f"A+(B-A)*(min(T,{wd})/{wd})"

    cmd = [ffmpeg, "-y", "-loglevel", "error",
           "-loop", "1", "-framerate", "24", "-i", str(prev_frame_path),
           "-loop", "1", "-framerate", "24", "-i", str(curr_frame_path)]

    if has_aud:
        cmd += ["-i", str(audio_path)]
    else:
        cmd += ["-f", "lavfi", "-i", "anullsrc=r=44100:cl=stereo"]

    cmd += [
        "-filter_complex", f"[0][1]blend=all_expr='{filter_expr}'[v]",
        "-map", "[v]", "-map", "2:a",
        "-c:v", "libx264", "-preset", "veryfast", "-pix_fmt", "yuv420p", "-r", "24",
        "-c:a", "aac", "-b:a", "128k", "-ar", "44100", "-ac", "2",
        "-t", f"{dur:.4f}",
        str(output_path),
    ]

    r = subprocess.run(cmd, capture_output=True)
    if r.returncode != 0:
        return None
    return str(output_path)


def _progressive_comparison_frames(lo_frame_path, rows, slide_emu, frames_dir, stem):
    """
    Create N progressive frames for a comparison slide.
    Frame k shows rows 0..k and whites-out rows k+1..n-1.
    Returns list of frame paths in order.
    """
    from PIL import Image, ImageDraw

    sw, sh = slide_emu
    n       = len(rows)
    result  = []

    def to_px(bbox, pad=4):
        left, top, width, height = bbox
        return (max(0, int(left  * W / sw) - pad),
                max(0, int(top   * H / sh) - pad),
                min(W, int((left + width)  * W / sw) + pad),
                min(H, int((top  + height) * H / sh) + pad))

    for k in range(n):
        prog = frames_dir / f"{stem}_r{k:02d}.png"
        img  = Image.open(str(lo_frame_path)).convert("RGB").resize((W, H), Image.LANCZOS)
        draw = ImageDraw.Draw(img)

        # White-out all rows that haven't been revealed yet (text + box outline)
        for j in range(k + 1, n):
            for bbox in (rows[j].get("box_bbox"), rows[j]["h_bbox"], rows[j]["b_bbox"]):
                if bbox:
                    draw.rectangle(to_px(bbox, pad=6), fill=(255, 255, 255))

        img.save(str(prog))
        result.append(str(prog))

    return result


def _progressive_bullet_frames_lo(pptx_path, slide_idx, bullets, frames_dir, stem):
    """
    Create N LO-rendered progressive frames for a bullet slide so the output
    exactly matches the POC_3.pptx template (fonts, colors, panel layout).

    For frame k, bullets 0..k are visible in PH_BODY (idx 41); bullets k+1..n-1
    have their text blanked so they're invisible in the rendered image.

    Returns list of N PNG paths (cached — existing files are reused), or None
    if LibreOffice is unavailable or rendering fails (caller falls back to Pillow).
    """
    import shutil
    from pptx import Presentation as _PrsLo
    from pptx.oxml.ns import qn as _qn

    n = len(bullets)
    soffice = _find_soffice()
    if not soffice:
        return None

    marker   = frames_dir / f"{stem}_b.lo"
    all_pngs = [frames_dir / f"{stem}_b{k:02d}.png" for k in range(n)]

    # Cache valid only when marker is newer than the PPTX (stale on every rebuild)
    pptx_mtime = Path(pptx_path).stat().st_mtime
    cache_ok   = (marker.exists()
                  and marker.stat().st_mtime >= pptx_mtime
                  and all(p.exists() for p in all_pngs))
    if cache_ok:
        return [str(p) for p in all_pngs]

    # Bust stale cache
    for p in all_pngs:
        if p.exists():
            p.unlink()
    if marker.exists():
        marker.unlink()

    result = []
    for k in range(n):
        out_png = all_pngs[k]

        print(f"\n  [LO frame] slide {slide_idx+1} bullet {k+1}/{n} …",
              end=" ", flush=True)

        prs2 = _PrsLo(pptx_path)
        sl = prs2.slides[slide_idx]

        from lxml import etree as _et

        # ── Find and blank future bullets ─────────────────────────────────────
        # _fill_bullets_ph leaves PH[41] text-free (for the purple border only)
        # and puts all bullet text in a free textbox overlaid on top.
        # Find that free textbox by matching paragraph count and first-bullet text,
        # then blank paragraphs k+1..n-1 so only bullets 0..k are visible.
        spTree = sl._element.find(".//" + _qn("p:spTree"))
        bullet_sp = None
        if spTree is not None:
            for sp in spTree.iter(_qn("p:sp")):
                nvPr = sp.find(".//" + _qn("p:nvPr"))
                if nvPr is not None and nvPr.find(_qn("p:ph")) is not None:
                    continue  # skip placeholders
                txb = sp.find(_qn("p:txBody"))
                if txb is None:
                    continue
                non_empty = [ap for ap in txb.findall(_qn("a:p"))
                             if "".join(t.text or "" for t in ap.findall(".//" + _qn("a:t"))).strip()]
                if len(non_empty) == len(bullets) and bullets:
                    first = "".join(t.text or "" for t in non_empty[0].findall(".//" + _qn("a:t"))).strip()
                    if first == bullets[0]:
                        bullet_sp = sp
                        break

        if bullet_sp is not None:
            # Free-textbox path: blank future paragraphs directly.
            # No color/indent fix needed — _fill_bullets_ph already styles them.
            txb = bullet_sp.find(_qn("p:txBody"))
            visible = 0
            for ap in txb.findall(_qn("a:p")):
                ap_text = "".join(t.text or "" for t in ap.findall(".//" + _qn("a:t"))).strip()
                if not ap_text:
                    continue
                if visible > k:
                    for t_el in ap.findall(".//" + _qn("a:t")):
                        t_el.text = ""
                    pPr = ap.find(_qn("a:pPr"))
                    if pPr is not None:
                        for tag in (_qn("a:buChar"), _qn("a:buFont")):
                            for el in list(pPr.findall(tag)):
                                pPr.remove(el)
                        if pPr.find(_qn("a:buNone")) is None:
                            _et.SubElement(pPr, _qn("a:buNone"))
                visible += 1
        else:
            # Fallback: bullets are in PH[41] directly (old pattern).
            for ph in sl.placeholders:
                if ph.placeholder_format.idx != 41:
                    continue
                visible = 0
                for para in ph.text_frame.paragraphs:
                    if para._p.text.strip():
                        if visible > k:
                            for t_el in para._p.findall(".//" + _qn("a:t")):
                                t_el.text = ""
                            bpPr = para._p.find(_qn("a:pPr"))
                            if bpPr is not None:
                                for tag in (_qn("a:buChar"), _qn("a:buFont")):
                                    for el in list(bpPr.findall(tag)):
                                        bpPr.remove(el)
                                if bpPr.find(_qn("a:buNone")) is None:
                                    _et.SubElement(bpPr, _qn("a:buNone"))
                        else:
                            for r_el in para._p.findall(".//" + _qn("a:r")):
                                rPr = r_el.find(_qn("a:rPr"))
                                if rPr is None:
                                    rPr = _et.SubElement(r_el, _qn("a:rPr"))
                                for sf in list(rPr.findall(_qn("a:solidFill"))):
                                    rPr.remove(sf)
                                sf = _et.SubElement(rPr, _qn("a:solidFill"))
                                _et.SubElement(sf, _qn("a:srgbClr")).set("val", "1A1A2E")
                            pPr = para._p.find(_qn("a:pPr"))
                            if pPr is None:
                                pPr = _et.Element(_qn("a:pPr"))
                                para._p.insert(0, pPr)
                            pPr.set("marL", "625475")
                            pPr.set("indent", "-260350")
                        visible += 1
                break

        # ── Fix layout xfrm for PH[41] border rendering ───────────────────────
        # LibreOffice uses the LAYOUT's xfrm for <p:ph> shapes and ignores the
        # slide-level <p:spPr> override. Stamp the aligned position onto the
        # layout so the purple rounded-rect border renders at the correct size.
        lx, ly, lw, lh = 1319951, 2080941, 4420899, 3335098
        title_x = 407823
        for lph in sl.slide_layout.placeholders:
            li    = lph.placeholder_format.idx
            lxfrm = lph._element.find(".//" + _qn("a:xfrm"))
            if lxfrm is None:
                continue
            lo = lxfrm.find(_qn("a:off"))
            le = lxfrm.find(_qn("a:ext"))
            if lo is None or le is None:
                continue
            if li == 41:
                lx = int(lo.get("x", lx)); ly = int(lo.get("y", ly))
                lw = int(le.get("cx", lw)); lh = int(le.get("cy", lh))
            elif li == 0:
                title_x = int(lo.get("x", title_x))

        new_x = title_x
        new_w = lw + lx - new_x

        for lph in sl.slide_layout.placeholders:
            if lph.placeholder_format.idx != 41:
                continue
            lspPr = lph._element.find(_qn("p:spPr"))
            if lspPr is None:
                lspPr = _et.SubElement(lph._element, _qn("p:spPr"))
            for old_xfrm in list(lspPr.findall(_qn("a:xfrm"))):
                lspPr.remove(old_xfrm)
            new_xfrm = _et.Element(_qn("a:xfrm"))
            new_off  = _et.SubElement(new_xfrm, _qn("a:off"))
            new_off.set("x", str(new_x)); new_off.set("y", str(ly))
            new_ext  = _et.SubElement(new_xfrm, _qn("a:ext"))
            new_ext.set("cx", str(new_w)); new_ext.set("cy", str(lh))
            lspPr.insert(0, new_xfrm)
            break

        tmp_dir = frames_dir / f"_lo_{stem}_{k}"
        tmp_dir.mkdir(exist_ok=True)
        tmp_pptx = tmp_dir / "render.pptx"
        prs2.save(str(tmp_pptx))

        # --norestore prevents LibreOffice from locking its user profile between
        # sequential headless calls (a common cause of silent failures on macOS).
        r = subprocess.run(
            [soffice, "--headless", "--norestore", "--convert-to", "pdf",
             "--outdir", str(tmp_dir), str(tmp_pptx)],
            capture_output=True, text=True, timeout=120,
        )
        if r.returncode != 0:
            print(f"WARN ({r.stderr[:60]})")
            shutil.rmtree(str(tmp_dir), ignore_errors=True)
            return None

        pdf_path = tmp_dir / "render.pdf"
        if not pdf_path.exists():
            print("WARN (no PDF)")
            shutil.rmtree(str(tmp_dir), ignore_errors=True)
            return None

        try:
            import fitz
            doc = fitz.open(str(pdf_path))
            if slide_idx < len(doc):
                page = doc[slide_idx]
                zoom = W / page.rect.width
                pix  = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
                pix.save(str(out_png))
                print("done")
            else:
                print(f"WARN (slide_idx {slide_idx} out of range {len(doc)})")
            doc.close()
        except Exception as e:
            print(f"WARN ({e})")
            shutil.rmtree(str(tmp_dir), ignore_errors=True)
            return None

        shutil.rmtree(str(tmp_dir), ignore_errors=True)
        result.append(str(out_png))

    if len(result) == n:
        marker.touch()   # mark all frames as LO-rendered for future cache hits
        return result
    return None


def _overlay_bullets_on_frame(base_path, bullets_visible, all_bullets,
                               bbox_emu, slide_emu, out_path):
    """
    Overlay progressively revealed bullets on a LibreOffice-rendered base frame.
    - Blanks the bullet shape area entirely (removes any LibreOffice ghost-text).
    - Font size is pre-calculated from ALL bullets so it stays constant across frames.
    - Each bullet is word-wrapped to fit within the box width.

    bullets_visible — bullets to render on this frame (1..k of all_bullets)
    all_bullets     — complete bullet list (used for consistent font sizing)
    bbox_emu        — (left, top, width, height) of the bullet shape in EMU
    slide_emu       — (slide_width, slide_height) in EMU
    """
    import textwrap as _tw
    from PIL import Image, ImageDraw

    sw, sh = slide_emu
    left, top, width, height = bbox_emu
    x1 = int(left  * W / sw)
    y1 = int(top   * H / sh)
    x2 = int((left + width)  * W / sw)
    y2 = int((top  + height) * H / sh)

    img  = Image.open(str(base_path)).convert("RGB").resize((W, H), Image.LANCZOS)
    draw = ImageDraw.Draw(img)

    # Fill bullet area with the template's dark-navy panel color; no visible border.
    # This matches the POC_3.pptx "60:40 Icon and text box with bullets" left-panel
    # background so the overlay blends seamlessly with the LibreOffice-rendered slide.
    draw.rectangle([x1, y1, x2, y2], fill=_BLUE)

    pad_x   = max(24, (x2 - x1) // 20)
    pad_y   = 24
    avail_w = (x2 - x1) - 2 * pad_x
    avail_h = (y2 - y1) - 2 * pad_y

    # Find the largest font where ALL bullets fit — font size stays fixed across frames
    fsize = 16
    for try_sz in range(26, 9, -1):
        fnt_t = _font(try_sz)
        try:
            cw = max(1, fnt_t.getlength("n"))
        except Exception:
            cw = try_sz * 0.56
        cpw    = max(5, int(avail_w / cw))
        lh     = int(try_sz * 1.4)
        needed = sum(
            max(1, len(_tw.wrap(f"•  {b}", cpw))) * lh + 6
            for b in all_bullets
        )
        if needed <= avail_h:
            fsize = try_sz
            break

    fnt = _font(fsize)
    try:
        cw = max(1, fnt.getlength("n"))
    except Exception:
        cw = fsize * 0.56
    cpw = max(5, int(avail_w / cw))
    lh  = int(fsize * 1.4)

    cy = y1 + pad_y
    for b in bullets_visible:
        lines = _tw.wrap(f"•  {b}", cpw) or [f"•  {b}"]
        for line in lines:
            if cy + lh > y2 - pad_y:
                break
            draw.text((x1 + pad_x, cy), line, font=fnt, fill=_WHITE)
            cy += lh
        cy += 6  # gap between bullets

    img.save(str(out_path))


# ── PPTX → frames (LibreOffice + pymupdf) ─────────────────────────────────────

def _find_soffice():
    for cmd in ("soffice", "libreoffice",
                "/Applications/LibreOffice.app/Contents/MacOS/soffice",
                "/usr/local/bin/soffice"):
        try:
            r = subprocess.run([cmd, "--version"], capture_output=True, timeout=10)
            if r.returncode == 0:
                return cmd
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass
    return None


def _pptx_to_frames_lo(pptx_path, frames_dir):
    soffice = _find_soffice()
    if not soffice:
        return None

    pdf_dir = frames_dir / "_pdf"
    pdf_dir.mkdir(exist_ok=True)
    r = subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf",
         "--outdir", str(pdf_dir), str(pptx_path)],
        capture_output=True, text=True, timeout=120,
    )
    if r.returncode != 0:
        print(f"\n          [WARN] LibreOffice: {r.stderr[:150]}")
        return None

    pdf_path = pdf_dir / (Path(pptx_path).stem + ".pdf")
    if not pdf_path.exists():
        return None

    try:
        import fitz
    except ImportError:
        print("\n          [WARN] pymupdf not installed — falling back to Pillow renderer")
        return None

    doc = fitz.open(str(pdf_path))
    frames = []
    for i, page in enumerate(doc):
        zoom = W / page.rect.width
        pix  = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        fp   = frames_dir / f"frame_{i + 1:03d}.png"
        pix.save(str(fp))
        frames.append(str(fp))
    doc.close()
    return frames


# ── Slide sequence ─────────────────────────────────────────────────────────────

def _slide_sequence(storyboard, images_dir):
    """Flat list of slide descriptors in presentation order (matches PPTX slide order)."""
    images_dir = Path(images_dir)
    slides = []

    cover_img = images_dir / "cover.png"
    slides.append({
        "type":  "cover",
        "title": storyboard["title"],
        "image": str(cover_img) if cover_img.exists() else None,
        "notes": f"Welcome to this learning module: {storyboard['title']}.",
    })

    for sec in storyboard["sections"]:
        slides.append({
            "type":    "divider",
            "sec_num": sec["section_id"],
            "heading": sec["section_heading"],
            "image":   None,
            "notes":   f"Section {sec['section_id']}: {sec['section_heading']}.",
        })
        for sd in sec["slides"]:
            img = images_dir / f"slide_{sd['slide_number']:03d}.png"
            slides.append({
                "type":    sd["scene_type"],
                "title":   sd["title"],
                "bullets": sd.get("bullets", []),
                "notes":   sd.get("speaker_notes", ""),
                "image":   str(img) if img.exists() else None,
            })
    return slides


def _render_pillow(slides, frames_dir):
    frames = []
    for i, s in enumerate(slides, 1):
        fp = frames_dir / f"frame_{i:03d}.png"
        t  = s["type"]
        if t == "cover":
            _render_cover(s["title"], s["image"], fp)
        elif t == "divider":
            _render_divider(s["sec_num"], s["heading"], fp)
        else:
            _render_content(t, s["title"], s.get("bullets", []), s["image"], fp)
        frames.append(str(fp))
        print(f"          [frame {i:3d}] {t}")
    return frames


def _get_frames(pptx_path, storyboard, images_dir, frames_dir):
    images_dir  = Path(images_dir)
    base_slides = _slide_sequence(storyboard, images_dir)

    print("          Trying LibreOffice rendering …", end=" ", flush=True)
    lo_frames = _pptx_to_frames_lo(pptx_path, frames_dir)
    lo_ok     = bool(lo_frames and len(lo_frames) == len(base_slides))

    if lo_ok:
        print(f"OK ({len(lo_frames)} frames)")
    elif lo_frames:
        print(f"frame count mismatch ({len(lo_frames)} vs {len(base_slides)}) — using Pillow")
    else:
        print("not available — using Pillow renderer")

    animated_slides  = []
    animated_frames  = []

    for i, slide in enumerate(base_slides):
        stype   = slide["type"]
        bullets = slide.get("bullets", [])

        if stype == "bullets" and bullets:
            # Progressive reveal: one Pillow frame per bullet.
            # Full VO audio is generated once then sliced in run_stage4.
            img_path  = slide.get("image")
            n         = len(bullets)
            group_key = f"frame_{i+1:03d}"
            for k in range(n):
                fp = frames_dir / f"frame_{i+1:03d}_b{k:02d}.png"
                _render_content(stype, slide["title"], bullets[:k+1], img_path, fp)
                animated_frames.append(str(fp))
                animated_slides.append({
                    **slide,
                    "notes":   slide.get("notes", ""),  # full VO — sliced later
                    "_bgroup": group_key,
                    "_bidx":   k,
                    "_btotal": n,
                })
        else:
            # Non-bullet slide: prefer LO frame, fall back to Pillow
            if lo_ok:
                fp = lo_frames[i]
            else:
                fp = frames_dir / f"frame_{i+1:03d}.png"
                t = stype
                if t == "cover":
                    _render_cover(slide["title"], slide.get("image"), fp)
                elif t == "divider":
                    _render_divider(slide["sec_num"], slide["heading"], fp)
                else:
                    _render_content(t, slide["title"], bullets, slide.get("image"), fp)
                fp = str(fp)
            animated_frames.append(fp if isinstance(fp, str) else str(fp))
            animated_slides.append(slide)

    return animated_slides, animated_frames


# ── TTS ────────────────────────────────────────────────────────────────────────

def _tts_edge(text, out_path, voice=TTS_VOICE):
    """Microsoft Edge neural TTS — free, no API key required. pip install edge-tts"""
    try:
        import edge_tts, asyncio
        async def _synth():
            communicate = edge_tts.Communicate(text, voice)
            await communicate.save(str(out_path))
        asyncio.run(_synth())
        return Path(out_path).exists()
    except ImportError:
        return False
    except Exception as e:
        print(f"\n          [WARN] edge-tts: {e}")
        return False


def _tts_edge_with_boundaries(text, out_path, voice=TTS_VOICE):
    """Generate TTS via edge-tts streaming and capture word-level timestamps.
    Saves audio to out_path; returns list of {"text", "start", "end"} in seconds."""
    try:
        import edge_tts, asyncio
        async def _go():
            comm = edge_tts.Communicate(text, voice, boundary="WordBoundary")
            audio, words = b"", []
            async for chunk in comm.stream():
                if chunk["type"] == "audio":
                    audio += chunk["data"]
                elif chunk["type"] == "WordBoundary":
                    s = chunk["offset"] / 10_000_000   # 100-ns ticks → seconds
                    d = chunk["duration"] / 10_000_000
                    words.append({"text": chunk["text"], "start": s, "end": s + d})
            return audio, words
        audio_bytes, words = asyncio.run(_go())
        if audio_bytes:
            Path(out_path).write_bytes(audio_bytes)
        return words
    except ImportError:
        return []
    except Exception as e:
        print(f"\n          [WARN] edge-tts stream: {e}")
        return []


def _audio_duration(audio_path):
    """Return audio duration in seconds via ffprobe."""
    cmd = [
        "ffprobe", "-v", "error", "-show_entries", "format=duration",
        "-of", "default=noprint_wrappers=1:nokey=1", str(audio_path),
    ]
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=10)
    try:
        return float(r.stdout.strip())
    except Exception:
        return 3.0


def _slice_audio(audio_path, start, duration, out_path):
    """Extract a timed slice from an audio file via ffmpeg."""
    cmd = [
        "ffmpeg", "-y", "-loglevel", "error",
        "-i", str(audio_path),
        "-ss", f"{start:.3f}", "-t", f"{duration:.3f}",
        "-acodec", "libmp3lame", "-q:a", "2",
        str(out_path),
    ]
    r = subprocess.run(cmd, capture_output=True, timeout=30)
    return r.returncode == 0 and Path(out_path).exists()


def _tts_openai(text, out_path, api_key, voice="nova"):
    try:
        from openai import AzureOpenAI
        endpoint   = os.environ.get("AZURE_OPENAI_ENDPOINT", "")
        api_ver    = os.environ.get("AZURE_OPENAI_API_VERSION", "2024-02-01")
        deployment = os.environ.get("AZURE_OPENAI_TTS_DEPLOYMENT", "tts")
        client = AzureOpenAI(api_key=api_key, azure_endpoint=endpoint, api_version=api_ver)
        resp = client.audio.speech.create(model=deployment, voice=voice, input=text)
        resp.stream_to_file(str(out_path))
        return True
    except Exception as e:
        print(f"\n          [WARN] Azure OpenAI TTS: {e}")
        return False


def _tts_gtts(text, out_path):
    try:
        from gtts import gTTS
        gTTS(text=text, lang="en", slow=False).save(str(out_path))
        return True
    except Exception as e:
        print(f"\n          [WARN] gTTS: {e}")
        return False


def _audio_for(notes, out_path, api_key=None):
    if not notes or not notes.strip():
        return None
    if _tts_edge(notes, out_path):          # Edge TTS — primary engine
        return str(out_path)
    if _tts_gtts(notes, out_path):          # last-resort fallback
        return str(out_path)
    return None


# ── Video assembly (pure ffmpeg — fast, no Python-level frame loop) ────────────

def _assemble(pairs, output_path):
    """
    Encode each (frame, audio, dur) pair as an individual MP4 segment using
    ffmpeg's native image-loop encoder, then stream-copy-concatenate them.
    This avoids moviepy's Python-level frame compositing which lags on 25+ clips.
    """
    import subprocess, shutil, tempfile

    ffmpeg = shutil.which("ffmpeg") or "/usr/local/bin/ffmpeg"
    tmpdir = Path(tempfile.mkdtemp(prefix="poc3_asmbl_"))

    try:
        segments = []
        for idx, (frame_path, audio_path, default_dur, trans_dur, pad) in enumerate(pairs):
            # Pre-built segment (e.g. ffmpeg-blend wipe) — use directly, skip re-encoding
            # Must be absolute path so the concat list.txt resolves it from the tmpdir
            if str(frame_path).endswith(".mp4") and Path(str(frame_path)).exists():
                segments.append(str(Path(str(frame_path)).resolve()))
                continue

            seg      = str(tmpdir / f"s{idx:04d}.mp4")
            has_aud  = audio_path and Path(audio_path).exists()
            dur      = max((_audio_duration(audio_path) + pad) if has_aud else default_dur, 0.15)

            # Optional video fadeout at clip end (only for slide-to-slide transitions)
            vf_args = []
            if trans_dur > 0:
                fade_st = max(0.0, dur - trans_dur)
                vf_args = ["-vf",
                           f"fade=type=out:start_time={fade_st:.4f}:duration={trans_dur:.4f}"]

            base_v = [ffmpeg, "-y", "-loglevel", "error",
                      "-loop", "1", "-framerate", "24", "-i", str(frame_path)]

            # Identical audio params on every segment — required for stream copy concat
            audio_enc = ["-c:a", "aac", "-b:a", "128k", "-ar", "44100", "-ac", "2"]

            if has_aud:
                cmd = base_v + ["-i", str(audio_path)] + vf_args + [
                    "-c:v", "libx264", "-preset", "veryfast", "-pix_fmt", "yuv420p", "-r", "24",
                ] + audio_enc + ["-t", f"{dur:.4f}", seg]
            else:
                cmd = base_v + ["-f", "lavfi", "-i", "anullsrc=r=44100:cl=stereo"] + vf_args + [
                    "-c:v", "libx264", "-preset", "veryfast", "-pix_fmt", "yuv420p", "-r", "24",
                ] + audio_enc + ["-t", f"{dur:.4f}", "-shortest", seg]

            r = subprocess.run(cmd, capture_output=True)
            if r.returncode != 0:
                print(f"  [WARN] seg {idx}: {r.stderr.decode()[-200:]}")
            else:
                segments.append(seg)

        if not segments:
            return None

        list_f = str(tmpdir / "list.txt")
        with open(list_f, "w") as f:
            for s in segments:
                f.write(f"file '{s}'\n")

        # Re-encode video to eliminate 1-frame codec discontinuities at segment seams;
        # normalize audio for consistent loudness across all slides.
        concat_cmd = [ffmpeg, "-y", "-loglevel", "error",
                      "-f", "concat", "-safe", "0", "-i", list_f,
                      "-c:v", "libx264", "-preset", "veryfast",
                      "-pix_fmt", "yuv420p", "-r", "24",
                      "-af", "loudnorm=I=-14:TP=-1.5:LRA=11:linear=true",
                      "-c:a", "aac", "-b:a", "128k", "-ar", "44100",
                      "-movflags", "+faststart",
                      str(output_path)]
        r = subprocess.run(concat_cmd, capture_output=True)
        if r.returncode != 0:
            print(f"  [WARN] concat: {r.stderr.decode()[-400:]}")

        return str(output_path) if Path(output_path).exists() else None

    finally:
        shutil.rmtree(str(tmpdir), ignore_errors=True)



# ── Stage 4 runner ─────────────────────────────────────────────────────────────

def run_stage4(storyboard_path, pptx_path, output_dir, openai_api_key=None):
    with open(storyboard_path, encoding="utf-8") as f:
        storyboard = json.load(f)

    out        = Path(output_dir)
    frames_dir = out / "video_frames"
    audio_dir  = out / "video_audio"
    frames_dir.mkdir(parents=True, exist_ok=True)
    audio_dir.mkdir(parents=True, exist_ok=True)

    images_dir = out / "images"

    # ── Frames ────────────────────────────────────────────────────────────────
    print(f"[Stage 4/4] Building slide frames …")
    slides, frames = _get_frames(pptx_path, storyboard, images_dir, frames_dir)

    # ── Audio ─────────────────────────────────────────────────────────────────
    print(f"[Stage 4/4] Generating audio for {len(slides)} slides …")
    try:
        import edge_tts as _et  # noqa
        _edge_available = True
    except ImportError:
        _edge_available = False
    tts_label = (f"edge-tts ({TTS_VOICE})" if _edge_available
                 else "gTTS (free fallback)")
    print(f"            Engine : {tts_label}")

    pairs = []
    bullet_groups = {}   # group_key → (full_audio_path, total_dur)

    for i, (slide, frame_path) in enumerate(zip(slides, frames), 1):
        notes  = slide.get("notes", "")
        stype  = slide["type"]
        bgroup = slide.get("_bgroup")
        bidx   = slide.get("_bidx", 0)
        btotal = slide.get("_btotal", 1)

        default_dur = COVER_DUR   if stype == "cover"   else \
                      DIVIDER_DUR if stype == "divider" else COVER_DUR

        if bgroup is not None:
            # ── Bullet sub-frame: per-bullet TTS via sentence-split chunks ───
            if bgroup not in bullet_groups:
                note_chunks = _split_vo_for_bullets(notes, btotal)
                grp_audios = []
                for b in range(btotal):
                    ap_b = audio_dir / f"audio_{bgroup}_b{b:02d}.mp3"
                    if ap_b.exists():
                        grp_audios.append(str(ap_b))
                    elif note_chunks[b].strip():
                        print(f"          [{i:2d}/{len(slides)}] TTS bullet {b+1}/{btotal} …",
                              end=" ", flush=True)
                        grp_audio = _audio_for(note_chunks[b], ap_b)
                        print("done" if grp_audio else "skipped")
                        grp_audios.append(grp_audio)
                    else:
                        grp_audios.append(None)
                bullet_groups[bgroup] = grp_audios

            audio_path = bullet_groups[bgroup][bidx] \
                         if bidx < len(bullet_groups[bgroup]) else None
            pairs.append((frame_path, audio_path, default_dur, BULLET_TRANS, BULLET_PAD))
        else:
            # ── Regular slide ─────────────────────────────────────────────────
            frame_stem = Path(frame_path).stem
            ap = audio_dir / f"audio_{frame_stem}.mp3"
            if ap.exists():
                audio_path = str(ap)
            else:
                print(f"          [{i:2d}/{len(slides)}] TTS …", end=" ", flush=True)
                audio_path = _audio_for(notes, ap, openai_api_key)
                print("done" if audio_path else "skipped (no notes)")

            pairs.append((frame_path, audio_path, default_dur, TRANSITION_DUR, PAD))

    # ── Video ─────────────────────────────────────────────────────────────────
    print(f"[Stage 4/4] Assembling video …")
    video_path = out / "storyboard_video.mp4"
    result = _assemble(pairs, video_path)

    if result:
        size_mb = Path(result).stat().st_size / (1024 * 1024)
        print(f"\n[Stage 4 Complete] → {result}  ({size_mb:.1f} MB)")
    else:
        print("[Stage 4] Video assembly failed.")

    manifest = {"video": result, "slide_count": len(slides)}
    with open(out / "stage4_manifest.json", "w") as f:
        json.dump(manifest, f, indent=2)
    return manifest


# ── Direct PPTX → Video (no storyboard JSON required) ─────────────────────────

def run_from_pptx(pptx_path: str, output_dir: str, openai_api_key=None):
    """
    Generate a narrated video directly from a PPTX file.

    - Slides with no speaker notes show silently for COVER_DUR seconds.
    - Slides with per-bullet Appear animations produce one progressive
      overlay frame per bullet; narration is split proportionally.
    - Stale audio cache (files older than the PPTX) is cleared automatically
      to prevent mismatched audio from previous builds.
    """
    from pptx import Presentation as _Prs

    out        = Path(output_dir)
    frames_dir = out / "video_frames"
    audio_dir  = out / "video_audio"
    frames_dir.mkdir(parents=True, exist_ok=True)
    audio_dir.mkdir(parents=True, exist_ok=True)

    # ── Clear stale audio cache when slides are removed (frame numbers shift) ──
    prs = _Prs(pptx_path)
    sw  = prs.slide_width
    sh  = prs.slide_height
    try:
        _max_cached_frame = max(
            int(f.name.split("_")[2]) for f in audio_dir.glob("audio_frame_*.mp3"))
    except ValueError:
        _max_cached_frame = 0
    if _max_cached_frame > len(prs.slides):
        stale = list(audio_dir.glob("audio_frame_*.mp3"))
        print(f"[Video] Clearing {len(stale)} stale audio file(s) (slide count changed) …")
        for f in stale:
            f.unlink()

    # ── Read slide metadata from PPTX ─────────────────────────────────────────

    slides_data = []
    for slide in prs.slides:
        notes = slide.notes_slide.notes_text_frame.text.strip() \
                if slide.has_notes_slide else ""
        anim = _get_animated_bullets(slide) or _get_bullet_content(slide)
        slides_data.append({
            "notes":    notes,
            "anim":     anim,
            "comp":     _get_comparison_rows_info(slide) or _get_icon3_rows_info(slide),
            "refl":     _get_reflection_info(slide),
            "img_bbox": _get_picture_bbox(slide),
        })

    n_slides = len(slides_data)
    print(f"[Video] {n_slides} slide(s) — {Path(pptx_path).name}")

    # ── PPTX → one base PNG per slide via LibreOffice ─────────────────────────
    print("[Video] Converting slides to frames …", end=" ", flush=True)
    lo_frames = _pptx_to_frames_lo(pptx_path, frames_dir)

    if lo_frames and len(lo_frames) == n_slides:
        print(f"OK ({len(lo_frames)} frames)")
    elif lo_frames:
        print(f"count mismatch ({len(lo_frames)} vs {n_slides}) — trimming")
        slides_data = slides_data[:len(lo_frames)]
    else:
        print("LibreOffice unavailable — cannot render PPTX frames")
        return {"video": None, "slide_count": 0}

    # ── Build (frame, audio, dur, trans, pad) pairs ───────────────────────────
    print(f"[Video] Generating audio …")
    pairs = []

    for i, (sd, lo_frame) in enumerate(zip(slides_data, lo_frames), 1):
        notes = sd["notes"]
        anim  = sd["anim"]
        comp  = sd["comp"]
        refl  = sd.get("refl")
        stem  = f"frame_{i:03d}"

        if comp:
            # ── Comparison slide: Header → Image → rows one at a time ─────────
            img_bbox = sd["img_bbox"]
            n        = len(comp)

            # Progressive row reveal frames first — needed as render base for intro
            prog_paths = _progressive_comparison_frames(
                lo_frame, comp, (sw, sh), frames_dir, stem)

            # Intro frames: use first progressive frame as base so the rendering
            # context is identical (eliminates flash at title→row-0 transition)
            all_row_bboxes = [bb for r in comp
                              for bb in (r.get("box_bbox"), r["h_bbox"], r["b_bbox"]) if bb]
            intro_base = prog_paths[0] if prog_paths else lo_frame
            title_frame, img_frame = _make_intro_frames(
                intro_base, img_bbox, all_row_bboxes, frames_dir, stem, (sw, sh))

            # One full TTS → per-row slices timed to word-level boundaries
            full_ap       = audio_dir / f"audio_{stem}_full.mp3"
            bound_ap      = audio_dir / f"audio_{stem}_full.json"
            r_files       = [audio_dir / f"audio_{stem}_r{k:02d}.mp3" for k in range(n)]
            intro_wipe_ap = audio_dir / f"audio_{stem}_intro_wipe.mp3"
            intro_fade_ap = audio_dir / f"audio_{stem}_intro_fade.mp3"
            pre_r_ap      = audio_dir / f"audio_{stem}_pre_row.mp3"

            # Migrate: old full audio without _full suffix → rename
            _old_full = audio_dir / f"audio_{stem}.mp3"
            if _old_full.exists() and not full_ap.exists():
                _old_full.rename(full_ap)
                for _f in r_files + [intro_wipe_ap, intro_fade_ap, pre_r_ap]:
                    if _f.exists(): _f.unlink()

            # Migrate: per-row TTS files without full audio → purge, regenerate
            if not full_ap.exists() and any(f.exists() for f in r_files):
                for _f in r_files + [intro_wipe_ap, intro_fade_ap, pre_r_ap]:
                    if _f.exists(): _f.unlink()
                for _k in range(n):
                    _seg = frames_dir / f"{stem}_r{_k:02d}_anim.mp4"
                    if _seg.exists(): _seg.unlink()

            if not full_ap.exists():
                print(f"  [{i:2d}/{len(slides_data)}] TTS …", end=" ", flush=True)
                _bounds = _tts_edge_with_boundaries(notes, full_ap)
                if not full_ap.exists():
                    _audio_for(notes, full_ap, openai_api_key)
                    _bounds = []
                if _bounds and full_ap.exists():
                    import json as _jmod
                    bound_ap.write_text(_jmod.dumps(_bounds))
                print("done" if full_ap.exists() else "skipped")
            else:
                print(f"  [{i:2d}/{len(slides_data)}] reusing audio")

            # Compute slice times (always — needed for intro slice boundaries)
            import json as _jmod
            _bounds    = _jmod.loads(bound_ap.read_text()) if bound_ap.exists() else []
            _total_dur = _audio_duration(str(full_ap)) if full_ap.exists() else 0.0
            _row_keys  = [r["heading"] for r in comp]
            _stimes    = _bullet_slice_times(_row_keys, _bounds, _total_dur)

            # Invalidate stale r-files when slice boundaries have changed
            if _stimes and r_files[0].exists() and \
                    abs(_audio_duration(str(r_files[0])) - _stimes[0][1]) > 0.15:
                for _f in r_files:
                    if _f.exists(): _f.unlink()
                for _k in range(n):
                    (frames_dir / f"{stem}_r{_k:02d}_anim.mp4").unlink(missing_ok=True)
                (frames_dir / f"{stem}_header_wipe.mp4").unlink(missing_ok=True)

            # Per-row slices
            if full_ap.exists() and not all(f.exists() for f in r_files):
                for _k, (_st, _dur) in enumerate(_stimes):
                    if not r_files[_k].exists():
                        _slice_audio(str(full_ap), _st, _dur, str(r_files[_k]))

            # Intro audio slices: wipe (0→0.45s), fade (0.45→1.05s), pre-row hold
            _iw_dur = _WIPE_DUR + _HEADER_HOLD   # 0.45s
            _if_dur = _FADE_DUR  + _IMAGE_HOLD   # 0.60s
            _it_dur = _iw_dur + _if_dur           # 1.05s
            if full_ap.exists():
                if not intro_wipe_ap.exists():
                    _slice_audio(str(full_ap), 0.0, _iw_dur, str(intro_wipe_ap))
                if not intro_fade_ap.exists():
                    _slice_audio(str(full_ap), _iw_dur, _if_dur, str(intro_fade_ap))
                _first_t = _stimes[0][0] if _stimes else _it_dur
                if _first_t > _it_dur and not pre_r_ap.exists():
                    _slice_audio(str(full_ap), _it_dur, _first_t - _it_dur, str(pre_r_ap))

            r_audios         = [str(f) if f.exists() else None for f in r_files]
            intro_wipe_audio = str(intro_wipe_ap) if intro_wipe_ap.exists() else None
            intro_fade_audio = str(intro_fade_ap) if intro_fade_ap.exists() else None
            pre_r_audio      = str(pre_r_ap)      if pre_r_ap.exists()      else None

            # ── Animation: Header wipe (with VO) → Image fade (with VO) → rows
            wipe_hold_dur = _WIPE_DUR + _HEADER_HOLD   # 0.45 s
            fade_hold_dur = _FADE_DUR  + _IMAGE_HOLD   # 0.60 s

            # Create blank white frame: intro wipe starts from blank (not full slide)
            blank_frame = str(frames_dir / f"{stem}_blank.png")
            if not Path(blank_frame).exists():
                from PIL import Image as _PILbk
                _PILbk.new("RGB", (W, H), (255, 255, 255)).save(blank_frame)
                (frames_dir / f"{stem}_header_wipe.mp4").unlink(missing_ok=True)

            # 1. Header wipe L→R (with narration): blank → title wipes in from left
            seg_hdr = str(frames_dir / f"{stem}_header_wipe.mp4")
            built_hdr = _make_wipe_segment(
                blank_frame, title_frame, intro_wipe_audio, wipe_hold_dur, _WIPE_DUR, seg_hdr)
            if built_hdr:
                pairs.append((built_hdr, None, wipe_hold_dur, 0.0, 0))
            else:
                from PIL import Image as _PILFb
                _bg_fb    = _PILFb.new("RGB", (W, H), (255, 255, 255))
                _title_fb = _PILFb.open(title_frame).convert("RGB")
                _append_wipe(pairs, _bg_fb, _title_fb, frames_dir, f"{stem}_title")
                if _HEADER_HOLD > 0:
                    pairs.append((title_frame, intro_wipe_audio, _HEADER_HOLD, 0.0, 0))

            # 2. Image fade in (with narration)
            seg_ifade = str(frames_dir / f"{stem}_img_fade.mp4")
            built_ifade = _make_fade_segment(
                title_frame, img_frame, intro_fade_audio, fade_hold_dur, _FADE_DUR, seg_ifade)
            if built_ifade:
                pairs.append((built_ifade, None, fade_hold_dur, 0.0, 0))
            else:
                from PIL import Image as _PILFb2
                _title_fb2 = _PILFb2.open(title_frame).convert("RGB")
                _img_fb2   = _PILFb2.open(img_frame).convert("RGB")
                _append_fade(pairs, _title_fb2, _img_fb2, frames_dir, f"{stem}_img")
                if _IMAGE_HOLD > 0:
                    pairs.append((img_frame, intro_fade_audio, _IMAGE_HOLD, 0.0, 0))

            # 2.5 Pre-row hold: bridge from intro end to first row keyword
            if pre_r_audio:
                pairs.append((img_frame, pre_r_audio, _audio_duration(pre_r_audio), 0.0, 0))

            # 3. Rows: each wipes in at its word-boundary timestamp
            _prev_frame_path_r = img_frame

            for k, prog in enumerate(prog_paths):
                audio_path_k = r_audios[k] if k < len(r_audios) else None
                b_dur = _audio_duration(audio_path_k) if audio_path_k else BULLET_SILENT_DUR

                seg_path = str(frames_dir / f"{stem}_r{k:02d}_anim.mp4")
                built = _make_wipe_segment(_prev_frame_path_r, prog,
                                           audio_path_k, b_dur, _BULLET_WIPE, seg_path)
                pairs.append((built if built else prog,
                               None if built else audio_path_k,
                               b_dur, 0.0, 0))
                _prev_frame_path_r = prog

        elif anim:
            # ── Bullet slide: Header → Image → bullets one at a time ──────────
            img_bbox = sd["img_bbox"]
            bullets  = anim["bullets"]
            n        = len(bullets)
            bbox_emu = (anim["left"], anim["top"], anim["width"], anim["height"])

            # Some layouts place an empty styled placeholder (e.g. the purple
            # rounded-rect body container, ph idx=41) whose theme-inherited border
            # is visible even when the placeholder has no text.  Collect those so
            # _make_intro_frames whites them out before any bullets appear.
            _extra_bboxes = []
            for _sh in prs.slides[i - 1].shapes:
                try:
                    _ph = _sh.placeholder_format
                    if _ph is None or _ph.idx == 0:
                        continue
                    if _sh.shape_type == 13:   # picture placeholder
                        continue
                    if _sh.has_text_frame and _sh.text_frame.text.strip() == "":
                        _extra_bboxes.append(
                            (_sh.left, _sh.top, _sh.width, _sh.height))
                except Exception:
                    pass

            # LO progressive frames first — used as render base for intro frames
            lo_progs = _progressive_bullet_frames_lo(
                pptx_path, i - 1, bullets, frames_dir, stem)

            # Intro frames: use first LO progressive frame as base — identical
            # rendering context eliminates the flash at title+image → bullet-0 cut
            intro_base = lo_progs[0] if lo_progs else lo_frame
            title_frame, img_frame = _make_intro_frames(
                intro_base, img_bbox, [bbox_emu] + _extra_bboxes, frames_dir, stem, (sw, sh))

            # One full TTS → per-bullet slices timed to word-level boundaries
            full_ap  = audio_dir / f"audio_{stem}_full.mp3"
            bound_ap = audio_dir / f"audio_{stem}_full.json"
            b_files  = [audio_dir / f"audio_{stem}_b{k:02d}.mp3" for k in range(n)]

            # Migrate: per-bullet TTS files without full audio → purge, regenerate
            if not full_ap.exists() and any(f.exists() for f in b_files):
                for _f in b_files:
                    if _f.exists(): _f.unlink()
                for _k in range(n):
                    _seg = frames_dir / f"{stem}_b{_k:02d}_anim.mp4"
                    if _seg.exists(): _seg.unlink()
                for _nm in ("_intro_wipe", "_intro_fade", "_pre_bullet"):
                    _f = audio_dir / f"audio_{stem}{_nm}.mp3"
                    if _f.exists(): _f.unlink()

            if not full_ap.exists():
                print(f"  [{i:2d}/{len(slides_data)}] TTS …", end=" ", flush=True)
                _bounds = _tts_edge_with_boundaries(notes, full_ap)
                if not full_ap.exists():
                    _audio_for(notes, full_ap, openai_api_key)
                    _bounds = []
                if _bounds and full_ap.exists():
                    import json as _jmod
                    bound_ap.write_text(_jmod.dumps(_bounds))
                print("done" if full_ap.exists() else "skipped")
            else:
                print(f"  [{i:2d}/{len(slides_data)}] reusing audio")

            # Compute slice times (always — needed for intro slice boundaries)
            import json as _jmod
            _bounds    = _jmod.loads(bound_ap.read_text()) if bound_ap.exists() else []
            _total_dur = _audio_duration(str(full_ap)) if full_ap.exists() else 0.0
            _stimes    = _bullet_slice_times(bullets, _bounds, _total_dur)

            # Invalidate stale b-files when slice boundaries have changed
            if _stimes and b_files[0].exists() and \
                    abs(_audio_duration(str(b_files[0])) - _stimes[0][1]) > 0.15:
                for _f in b_files:
                    if _f.exists(): _f.unlink()
                for _k in range(n):
                    (frames_dir / f"{stem}_b{_k:02d}_anim.mp4").unlink(missing_ok=True)
                (frames_dir / f"{stem}_header_wipe.mp4").unlink(missing_ok=True)

            # Per-bullet slices
            if full_ap.exists() and not all(f.exists() for f in b_files):
                for _k, (_st, _dur) in enumerate(_stimes):
                    if not b_files[_k].exists():
                        _slice_audio(str(full_ap), _st, _dur, str(b_files[_k]))

            # Intro audio slices: wipe (0→0.45s), fade (0.45→1.05s), pre-bullet hold
            intro_wipe_ap = audio_dir / f"audio_{stem}_intro_wipe.mp3"
            intro_fade_ap = audio_dir / f"audio_{stem}_intro_fade.mp3"
            pre_b_ap      = audio_dir / f"audio_{stem}_pre_bullet.mp3"
            _iw_dur = _WIPE_DUR + _HEADER_HOLD   # 0.45s
            _if_dur = _FADE_DUR  + _IMAGE_HOLD   # 0.60s
            _it_dur = _iw_dur + _if_dur           # 1.05s
            if full_ap.exists():
                if not intro_wipe_ap.exists():
                    _slice_audio(str(full_ap), 0.0, _iw_dur, str(intro_wipe_ap))
                if not intro_fade_ap.exists():
                    _slice_audio(str(full_ap), _iw_dur, _if_dur, str(intro_fade_ap))
                _first_t = _stimes[0][0] if _stimes else _it_dur
                if _first_t > _it_dur and not pre_b_ap.exists():
                    _slice_audio(str(full_ap), _it_dur, _first_t - _it_dur, str(pre_b_ap))

            b_audios         = [str(f) if f.exists() else None for f in b_files]
            intro_wipe_audio = str(intro_wipe_ap) if intro_wipe_ap.exists() else None
            intro_fade_audio = str(intro_fade_ap) if intro_fade_ap.exists() else None
            pre_b_audio      = str(pre_b_ap)      if pre_b_ap.exists()      else None

            # ── Animation: Header wipe (with VO) → Image fade (with VO) → bullets
            wipe_hold_dur = _WIPE_DUR + _HEADER_HOLD   # 0.45 s
            fade_hold_dur = _FADE_DUR  + _IMAGE_HOLD   # 0.60 s

            # Create blank white frame: intro wipe starts from blank (not full slide)
            blank_frame = str(frames_dir / f"{stem}_blank.png")
            if not Path(blank_frame).exists():
                from PIL import Image as _PILbk
                _PILbk.new("RGB", (W, H), (255, 255, 255)).save(blank_frame)
                (frames_dir / f"{stem}_header_wipe.mp4").unlink(missing_ok=True)

            # 1. Header wipe L→R (with narration): blank → title wipes in from left
            seg_hdr = str(frames_dir / f"{stem}_header_wipe.mp4")
            built_hdr = _make_wipe_segment(
                blank_frame, title_frame, intro_wipe_audio, wipe_hold_dur, _WIPE_DUR, seg_hdr)
            if built_hdr:
                pairs.append((built_hdr, None, wipe_hold_dur, 0.0, 0))
            else:
                from PIL import Image as _PILFb
                _bg_fb    = _PILFb.new("RGB", (W, H), (255, 255, 255))
                _title_fb = _PILFb.open(title_frame).convert("RGB")
                _append_wipe(pairs, _bg_fb, _title_fb, frames_dir, f"{stem}_title")
                if _HEADER_HOLD > 0:
                    pairs.append((title_frame, intro_wipe_audio, _HEADER_HOLD, 0.0, 0))

            # 2. Image fade in (with narration)
            seg_ifade = str(frames_dir / f"{stem}_img_fade.mp4")
            built_ifade = _make_fade_segment(
                title_frame, img_frame, intro_fade_audio, fade_hold_dur, _FADE_DUR, seg_ifade)
            if built_ifade:
                pairs.append((built_ifade, None, fade_hold_dur, 0.0, 0))
            else:
                from PIL import Image as _PILFb2
                _title_fb2 = _PILFb2.open(title_frame).convert("RGB")
                _img_fb2   = _PILFb2.open(img_frame).convert("RGB")
                _append_fade(pairs, _title_fb2, _img_fb2, frames_dir, f"{stem}_img")
                if _IMAGE_HOLD > 0:
                    pairs.append((img_frame, intro_fade_audio, _IMAGE_HOLD, 0.0, 0))

            # 2.5 Pre-bullet hold: bridge from intro end to first bullet keyword
            if pre_b_audio:
                pairs.append((img_frame, pre_b_audio, _audio_duration(pre_b_audio), 0.0, 0))

            # 3. Bullets: each wipes in at its word-boundary timestamp
            _prev_frame_path = img_frame

            for k in range(n):
                if lo_progs:
                    prog = lo_progs[k]
                else:
                    prog = str(frames_dir / f"{stem}_b{k:02d}.png")
                    _overlay_bullets_on_frame(lo_frame, bullets[:k + 1], bullets,
                                              bbox_emu, (sw, sh), prog)

                audio_path_k = b_audios[k] if k < len(b_audios) else None
                b_dur = _audio_duration(audio_path_k) if audio_path_k else BULLET_SILENT_DUR

                seg_path = str(frames_dir / f"{stem}_b{k:02d}_anim.mp4")
                built = _make_wipe_segment(_prev_frame_path, prog,
                                           audio_path_k, b_dur, _BULLET_WIPE, seg_path)
                pairs.append((built if built else prog,
                               None if built else audio_path_k,
                               b_dur, 0.0, 0))
                _prev_frame_path = prog

        elif refl:
            # ── Reflection slide: title static → wipe in question box ─────────
            from PIL import Image as _PILImg, ImageDraw as _PILDraw
            box_bbox   = refl["box_bbox"]

            before_p = str(frames_dir / f"{stem}_refl_before.png")
            if not Path(before_p).exists():
                _img = _PILImg.open(lo_frame).convert("RGB").resize((W, H), _PILImg.LANCZOS)
                _drw = _PILDraw.Draw(_img)
                def _rpx(bbox, pad=6):
                    l, t, ww, hh = bbox
                    return (max(0, int(l*W/sw) - pad), max(0, int(t*H/sh) - pad),
                            min(W, int((l+ww)*W/sw) + pad), min(H, int((t+hh)*H/sh) + pad))
                _drw.rectangle(_rpx(box_bbox), fill=(255, 255, 255))
                _img.save(before_p)

            # Audio (TTS if notes present)
            ap = audio_dir / f"audio_{stem}.mp3"
            if not notes.strip():
                audio_path = None
                print(f"  [{i:2d}/{len(slides_data)}] silent (no notes)")
            elif ap.exists():
                audio_path = str(ap)
                print(f"  [{i:2d}/{len(slides_data)}] reusing audio")
            else:
                print(f"  [{i:2d}/{len(slides_data)}] TTS …", end=" ", flush=True)
                audio_path = _audio_for(notes, ap, openai_api_key)
                print("done" if audio_path else "skipped")

            total_dur  = _audio_duration(audio_path) if audio_path else COVER_DUR
            wipe_dur   = 0.5
            hold_before = min(0.5, total_dur * 0.1)
            hold_after  = max(0.0, total_dur - hold_before - wipe_dur)

            # Static hold (before wipe)
            pairs.append((before_p, None, hold_before, 0.0, 0))

            # Wipe segment (before → after, audio plays during wipe+hold)
            seg_wipe = str(frames_dir / f"{stem}_refl_wipe.mp4")
            # Invalidate cache if audio is newer than the segment (was built silent before)
            _seg_p = Path(seg_wipe)
            if _seg_p.exists() and audio_path:
                if Path(audio_path).stat().st_mtime > _seg_p.stat().st_mtime:
                    _seg_p.unlink()
            if not Path(seg_wipe).exists():
                wipe_hold_dur = wipe_dur + hold_after
                _make_wipe_segment(before_p, lo_frame, audio_path,
                                   wipe_hold_dur, wipe_dur, seg_wipe)
            pairs.append((seg_wipe, None, wipe_dur + hold_after, 0.0, 0))

        else:
            # ── Regular slide: one frame, one clip ────────────────────────────
            ap = audio_dir / f"audio_{stem}.mp3"
            if not notes.strip():
                audio_path = None
                print(f"  [{i:2d}/{len(slides_data)}] silent (no notes)")
            elif ap.exists():
                audio_path = str(ap)
                print(f"  [{i:2d}/{len(slides_data)}] reusing audio")
            else:
                print(f"  [{i:2d}/{len(slides_data)}] TTS …", end=" ", flush=True)
                audio_path = _audio_for(notes, ap, openai_api_key)
                print("done" if audio_path else "skipped")

            dur = _audio_duration(audio_path) if audio_path else COVER_DUR
            pairs.append((lo_frame, audio_path, dur, TRANSITION_DUR, PAD))

    # ── Assemble final video ──────────────────────────────────────────────────
    print("[Video] Assembling …")
    video_path = out / "module9_video.mp4"
    result     = _assemble(pairs, video_path)

    if result:
        size_mb = Path(result).stat().st_size / (1024 * 1024)
        print(f"\n[Video] Done → {result}  ({size_mb:.1f} MB)")
    else:
        print("[Video] Assembly failed.")

    manifest = {"video": result, "slide_count": len(lo_frames)}
    with open(out / "stage4_direct_manifest.json", "w") as f:
        json.dump(manifest, f, indent=2)
    return manifest


# ── CLI ────────────────────────────────────────────────────────────────────────

def _load_env():
    env = Path(__file__).parent.parent / ".env"
    if env.exists():
        with open(env) as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, _, v = line.partition("=")
                    os.environ.setdefault(k.strip(), v.strip())


if __name__ == "__main__":
    _load_env()
    api_key = os.environ.get("AZURE_OPENAI_API_KEY")

    if "--from-pptx" in sys.argv:
        # Direct mode — no storyboard JSON needed
        # Usage: poc3_stage4_video.py --from-pptx <pptx_path> <output_dir>
        idx = sys.argv.index("--from-pptx")
        if len(sys.argv) < idx + 3:
            print("Usage: poc3_stage4_video.py --from-pptx <pptx_path> <output_dir>")
            sys.exit(1)
        result = run_from_pptx(sys.argv[idx + 1], sys.argv[idx + 2],
                               openai_api_key=api_key)
    else:
        # Original storyboard-driven mode
        if len(sys.argv) < 4:
            print(__doc__)
            sys.exit(1)
        result = run_stage4(sys.argv[1], sys.argv[2], sys.argv[3],
                            openai_api_key=api_key)

    print()
    for k, v in result.items():
        print(f"  {k}: {v}")
