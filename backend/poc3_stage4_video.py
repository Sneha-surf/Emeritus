#!/usr/bin/env python3
"""
POC 3 — Stage 4: Narrated Video Generation

Converts the storyboard into a narrated MP4:
  1. PPTX → PNG frames  (LibreOffice + pymupdf preferred; Pillow fallback)
  2. Speaker notes → MP3 (OpenAI TTS preferred; gTTS fallback)
  3. frame + audio → video clip  (moviepy)
  4. clips concatenated → storyboard_video.mp4

Usage:
  python poc3_stage4_video.py <storyboard_json> <pptx_path> <output_dir>

Environment:
  OPENAI_API_KEY  — standard OpenAI key for TTS (falls back to gTTS if unset)
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

COVER_DUR      = 4.0   # seconds when no audio
DIVIDER_DUR    = 2.5
PAD            = 0.40  # silence padding after each regular audio clip
TRANSITION_DUR = 0.40  # crossfade between slides — equals PAD so audio never collides
BULLET_PAD     = 0.05  # tiny pause after each bullet narration chunk
BULLET_TRANS   = 0.25  # cross-fade between bullet reveal frames
TTS_VOICE      = "en-US-AriaNeural"  # Microsoft neural voice via edge-tts (free)


# ── Font loader ────────────────────────────────────────────────────────────────

def _font(size, bold=False):
    from PIL import ImageFont
    bold_candidates = [
        ("/System/Library/Fonts/Helvetica.ttc",      {"index": 1}),
        ("/Library/Fonts/Arial Bold.ttf",            {}),
        ("/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf", {}),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",         {}),
    ]
    reg_candidates = [
        ("/System/Library/Fonts/Helvetica.ttc",      {"index": 0}),
        ("/Library/Fonts/Arial.ttf",                 {}),
        ("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", {}),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",                 {}),
    ]
    for path, kw in (bold_candidates if bold else reg_candidates):
        if Path(path).exists():
            try:
                return ImageFont.truetype(path, size, **kw)
            except Exception:
                pass
    return ImageFont.load_default()


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
            by += 8

    bg.save(str(out_path))
    return str(out_path)


def _split_vo_for_bullets(notes, n):
    """Split narration text into n sentence-chunks for per-bullet TTS."""
    import re
    if not notes or not notes.strip() or n <= 1:
        return [notes or ""] * max(n, 1)
    sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', notes.strip()) if s.strip()]
    if not sentences:
        return [notes] * n
    if len(sentences) <= n:
        return sentences + [""] * (n - len(sentences))
    per = len(sentences) / n
    chunks = []
    for i in range(n):
        start = int(round(i * per))
        end   = int(round((i + 1) * per)) if i < n - 1 else len(sentences)
        chunks.append(" ".join(sentences[start:end]))
    return chunks


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

    rows = []
    for h_idx, b_idx in _CMP_ROWS:
        h_bbox = layout_bbox(h_idx)
        b_bbox = layout_bbox(b_idx)
        if h_bbox is None:
            continue
        rows.append({
            "heading": slide_text(h_idx),
            "body":    slide_text(b_idx),
            "h_bbox":  h_bbox,
            "b_bbox":  b_bbox,
        })

    return rows if rows else None


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

        # White-out all rows that haven't been revealed yet
        for j in range(k + 1, n):
            for bbox in (rows[j]["h_bbox"], rows[j]["b_bbox"]):
                if bbox:
                    draw.rectangle(to_px(bbox), fill=(255, 255, 255))

        img.save(str(prog))
        result.append(str(prog))

    return result


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

    # Rounded rect — matches template (#4F3982 border, near-white lavender fill)
    radius = max(6, int((y2 - y1) * 0.04))
    try:
        draw.rounded_rectangle([x1, y1, x2, y2], radius=radius,
                                fill=(245, 242, 252), outline=(79, 57, 130), width=2)
    except AttributeError:
        draw.rectangle([x1, y1, x2, y2],
                       fill=(245, 242, 252), outline=(79, 57, 130), width=2)

    pad_x   = max(16, (x2 - x1) // 25)
    pad_y   = 16
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
            draw.text((x1 + pad_x, cy), line, font=fnt, fill=(26, 26, 46))
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
        from openai import OpenAI
        client = OpenAI(api_key=api_key)
        resp = client.audio.speech.create(model="tts-1", voice=voice, input=text)
        resp.stream_to_file(str(out_path))
        return True
    except Exception as e:
        print(f"\n          [WARN] OpenAI TTS: {e}")
        return False


def _tts_gtts(text, out_path):
    try:
        from gtts import gTTS
        gTTS(text=text, lang="en", slow=False).save(str(out_path))
        return True
    except Exception as e:
        print(f"\n          [WARN] gTTS: {e}")
        return False


def _audio_for(notes, out_path, api_key):
    if not notes or not notes.strip():
        return None
    if api_key and _tts_openai(notes, out_path, api_key):
        return str(out_path)
    if _tts_edge(notes, out_path):          # Microsoft Aria neural (free)
        return str(out_path)
    if _tts_gtts(notes, out_path):          # last-resort fallback
        return str(out_path)
    return None


# ── Video assembly (moviepy 2.x) ───────────────────────────────────────────────

def _assemble(pairs, output_path):
    # pairs: (frame_path, audio_path, default_dur, trans_dur, pad)
    # trans_dur — crossfade overlap with NEXT clip (0 = hard cut)
    # pad       — silence appended after audio before clip ends
    try:
        from moviepy import ImageClip, AudioFileClip, CompositeVideoClip
        import moviepy.video.fx as vfx
    except ImportError:
        print("[ERROR] moviepy not installed. Run: pip install moviepy")
        return None

    clips_td = []   # [(clip, trans_dur), ...]
    for frame_path, audio_path, default_dur, trans_dur, pad in pairs:
        if audio_path and Path(audio_path).exists():
            try:
                aud  = AudioFileClip(str(audio_path))
                dur  = aud.duration + pad
                clip = (ImageClip(str(frame_path))
                        .with_duration(dur)
                        .with_audio(aud))
            except Exception as e:
                print(f"          [WARN] audio error: {e}")
                clip = ImageClip(str(frame_path)).with_duration(default_dur)
        else:
            clip = ImageClip(str(frame_path)).with_duration(default_dur)
        clips_td.append((clip, trans_dur))

    if not clips_td:
        return None

    timed = []
    t = 0.0
    n = len(clips_td)
    for i, (clip, tr) in enumerate(clips_td):
        orig_dur   = clip.duration
        prev_tr    = clips_td[i - 1][1] if i > 0 else 0.0
        effects    = []
        if i > 0 and prev_tr > 0:
            effects.append(vfx.FadeIn(prev_tr))
        if i < n - 1 and tr > 0:
            effects.append(vfx.FadeOut(tr))
        faded = clip.with_effects(effects) if effects else clip
        timed.append(faded.with_start(t))
        if i < n - 1:
            t += orig_dur - tr   # overlap = tr (0 → hard cut, no overlap)

    total_dur = t + clips_td[-1][0].duration
    final = CompositeVideoClip(timed, size=(W, H)).with_duration(total_dur)
    final.write_videofile(
        str(output_path), fps=24,
        codec="libx264", audio_codec="aac",
        logger=None,
    )
    return str(output_path)



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
    tts_label = ("OpenAI TTS" if openai_api_key
                 else f"edge-tts ({TTS_VOICE})" if _edge_available
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
            # ── Bullet sub-frame: generate full VO once, slice per bullet ──────
            full_ap = audio_dir / f"audio_{bgroup}_full.mp3"

            if bgroup not in bullet_groups:
                if full_ap.exists():
                    full_audio = str(full_ap)
                else:
                    print(f"          [{i:2d}/{len(slides)}] TTS (full VO) …",
                          end=" ", flush=True)
                    full_audio = _audio_for(notes, full_ap, openai_api_key)
                    print("done" if full_audio else "skipped")
                D = _audio_duration(full_audio) if full_audio else default_dur * btotal
                bullet_groups[bgroup] = (full_audio, D)

            full_audio, D = bullet_groups[bgroup]
            chunk   = D / btotal
            start_t = bidx * chunk

            ap = audio_dir / f"audio_{bgroup}_b{bidx:02d}.mp3"
            if ap.exists():
                audio_path = str(ap)
            elif full_audio:
                audio_path = str(ap) if _slice_audio(full_audio, start_t, chunk, ap) else None
            else:
                audio_path = None

            pairs.append((frame_path, audio_path, chunk, BULLET_TRANS, BULLET_PAD))
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

    # ── Clear stale audio cache (frame numbering shifts when slides are added) ─
    pptx_mtime = Path(pptx_path).stat().st_mtime
    stale = [f for f in audio_dir.glob("audio_frame_*.mp3")
             if f.stat().st_mtime < pptx_mtime]
    if stale:
        print(f"[Video] Clearing {len(stale)} stale audio file(s) …")
        for f in stale:
            f.unlink()

    # ── Read slide metadata from PPTX ─────────────────────────────────────────
    prs = _Prs(pptx_path)
    sw  = prs.slide_width
    sh  = prs.slide_height

    slides_data = []
    for slide in prs.slides:
        notes = slide.notes_slide.notes_text_frame.text.strip() \
                if slide.has_notes_slide else ""
        slides_data.append({
            "notes": notes,
            "anim": _get_animated_bullets(slide),
            "comp": _get_comparison_rows_info(slide),
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
    pairs             = []
    bullet_full_cache = {}   # stem → (full_audio_path, total_dur)

    for i, (sd, lo_frame) in enumerate(zip(slides_data, lo_frames), 1):
        notes = sd["notes"]
        anim  = sd["anim"]
        comp  = sd["comp"]
        stem  = f"frame_{i:03d}"

        if comp:
            # ── Comparison slide: reveal one row at a time ────────────────────
            n       = len(comp)
            prog_paths = _progressive_comparison_frames(
                lo_frame, comp, (sw, sh), frames_dir, stem)

            full_ap = audio_dir / f"audio_{stem}_full.mp3"
            if full_ap.exists():
                full_audio = str(full_ap)
            elif notes.strip():
                print(f"  [{i:2d}/{len(slides_data)}] TTS ({n} rows) …",
                      end=" ", flush=True)
                full_audio = _audio_for(notes, full_ap, openai_api_key)
                print("done" if full_audio else "skipped")
            else:
                full_audio = None

            D     = _audio_duration(full_audio) if full_audio else COVER_DUR * n
            chunk = D / n

            for k, prog in enumerate(prog_paths):
                ap = audio_dir / f"audio_{stem}_r{k:02d}.mp3"
                if ap.exists():
                    audio_path = str(ap)
                elif full_audio:
                    audio_path = str(ap) \
                        if _slice_audio(full_audio, k * chunk, chunk, ap) else None
                else:
                    audio_path = None
                pairs.append((prog, audio_path, chunk, BULLET_TRANS, BULLET_PAD))

        elif anim:
            # ── Animated bullet slide: one progressive frame per bullet ───────
            bullets = anim["bullets"]
            n       = len(bullets)
            bbox    = (anim["left"], anim["top"], anim["width"], anim["height"])

            # Generate full narration once, then slice per bullet
            full_ap = audio_dir / f"audio_{stem}_full.mp3"
            if stem not in bullet_full_cache:
                if full_ap.exists():
                    full_audio = str(full_ap)
                elif notes.strip():
                    print(f"  [{i:2d}/{len(slides_data)}] TTS ({n} bullets) …",
                          end=" ", flush=True)
                    full_audio = _audio_for(notes, full_ap, openai_api_key)
                    print("done" if full_audio else "skipped")
                else:
                    full_audio = None
                D = _audio_duration(full_audio) if full_audio else COVER_DUR * n
                bullet_full_cache[stem] = (full_audio, D)

            full_audio, D = bullet_full_cache[stem]
            chunk = D / n

            for k in range(n):
                # Progressive overlay: blank bullet area, draw bullets 1..k.
                # Always regenerated — these are fast Pillow ops and must stay
                # consistent with the current all_bullets sizing calculation.
                prog = frames_dir / f"{stem}_b{k:02d}.png"
                _overlay_bullets_on_frame(lo_frame, bullets[:k + 1], bullets,
                                          bbox, (sw, sh), prog)

                # Slice audio for this bullet
                ap = audio_dir / f"audio_{stem}_b{k:02d}.mp3"
                if ap.exists():
                    audio_path = str(ap)
                elif full_audio:
                    audio_path = str(ap) \
                        if _slice_audio(full_audio, k * chunk, chunk, ap) else None
                else:
                    audio_path = None

                pairs.append((str(prog), audio_path, chunk, BULLET_TRANS, BULLET_PAD))

        else:
            # ── Regular slide: one frame, one clip ────────────────────────────
            ap = audio_dir / f"audio_{stem}.mp3"
            if not notes.strip():
                # Intro/outro or any slide with no speaker notes → silent hold
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
    api_key = os.environ.get("OPENAI_API_KEY")

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
