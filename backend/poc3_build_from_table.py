#!/usr/bin/env python3
"""
POC 3 — Direct Table-to-PPTX Builder

Reads the storyboard table from the Word document and generates a PPTX
using the POC_3.pptx template.

Rules applied:
  • Where visual cues say "faculty", the OST text is shown on-slide instead.
  • VO (narration) goes into speaker notes ONLY — verbatim, unchanged.
  • OST text is placed on slides verbatim — unchanged.
  • Bullet slides get per-paragraph Appear animations (on click).

Usage:
  python poc3_build_from_table.py <docx_path> <output_pptx>
  python poc3_build_from_table.py ../Final_txt_AIM_LAI_M9-V6.docx ../poc3_output/module9.pptx
"""

import sys
import re
from copy import deepcopy
from pathlib import Path

from docx import Document
from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.oxml.ns import qn

# ── Namespace for r:embed ──────────────────────────────────────────────────────
_NS_R      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_EMBED_ATTR = f"{{{_NS_R}}}embed"

# ── Layout names ───────────────────────────────────────────────────────────────
L_WIDE     = "1D_Wide Screen"
L_BULLETS  = "60:40 Icon and text box with bullets"
L_COMPARE  = "60:40 Text boxes with headings"
L_BOOKEND  = "1B_Intro-Outro"

# ── Placeholder indices — bullets layout ──────────────────────────────────────
PH_BULLETS_BANNER = 34   # module title banner (top)
PH_BULLETS_TITLE  = 0    # slide title
PH_BULLETS_BODY   = 41   # bullet body text

# ── Placeholder indices — comparison layout ───────────────────────────────────
PH_CMP_BANNER = 34
PH_CMP_TITLE  = 0
# pairs (heading_idx, body_idx) for each of the 3 rows
_CMP_ROWS = [(35, 36), (37, 38), (39, 40)]

# ── Right-panel image area (the right 40% of a 13.33" wide slide) ─────────────
_RP_LEFT  = Inches(8.2)
_RP_W     = Inches(4.9)
_RP_IMG_H = Emu(int(_RP_W * 9 / 16))
_RP_TOP   = Emu(int((Inches(7.5) - _RP_IMG_H) / 2))


# ── Helpers ────────────────────────────────────────────────────────────────────

def _find_layout(prs: Presentation, name: str):
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            if layout.name == name:
                return layout
    raise ValueError(f"Layout not found: {name!r}")


def _ph(slide, idx):
    for p in slide.placeholders:
        if p.placeholder_format.idx == idx:
            return p
    return None


def _fill_ph(slide, idx, text, *, size=None, bold=None, italic=None,
             color=None, align=None):
    ph = _ph(slide, idx)
    if ph is None:
        return
    tf = ph.text_frame
    tf.clear()
    tf.word_wrap = True
    para = tf.paragraphs[0]
    if align is not None:
        para.alignment = align
    run = para.add_run()
    run.text = text
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color:
        run.font.color.rgb = color


def _fill_bullets_ph(slide, idx, lines: list[str]):
    """
    Populate a placeholder with one <a:p> per line.

    To permanently suppress LibreOffice's "Click to edit text." overlay in
    exported video frames we must:
      1. Call tf.clear() so python-pptx initialises the txBody at the slide level.
      2. Copy spPr (rounded-rect geometry + border) and lstStyle (bullet char +
         font) from the matching layout placeholder so those styles survive after
         the next step.
      3. Strip <p:ph> from the slide's nvPr — once the shape has no placeholder
         marker LibreOffice treats it as a plain shape and never overlays hint text.
    """
    ph = _ph(slide, idx)
    if ph is None:
        return None

    tf = ph.text_frame
    tf.word_wrap = True
    tf.clear()

    # ── Borrow spPr + lstStyle from the layout placeholder ────────────────────
    for lph in slide.slide_layout.placeholders:
        if lph.placeholder_format.idx != idx:
            continue
        lph_el = lph._element

        # spPr — shape geometry (rounded rect, border)
        l_spPr = lph_el.find(qn("p:spPr"))
        s_spPr = ph._element.find(qn("p:spPr"))
        if l_spPr is not None and len(l_spPr) > 0 and s_spPr is not None:
            pos = list(ph._element).index(s_spPr)
            ph._element.remove(s_spPr)
            ph._element.insert(pos, deepcopy(l_spPr))

        # lstStyle + bodyPr from layout txBody
        l_tx = lph_el.find(qn("p:txBody"))
        s_tx = ph._element.find(qn("p:txBody"))
        if l_tx is not None and s_tx is not None:
            for child_tag in (qn("a:bodyPr"), qn("a:lstStyle")):
                l_c = l_tx.find(child_tag)
                s_c = s_tx.find(child_tag)
                if l_c is not None and s_c is not None:
                    p2 = list(s_tx).index(s_c)
                    s_tx.remove(s_c)
                    s_tx.insert(p2, deepcopy(l_c))
        break

    txBody = ph._element.find(qn("p:txBody"))
    if txBody is None:
        return None

    for ap in list(txBody.findall(qn("a:p"))):
        txBody.remove(ap)

    # Ensure text auto-fits inside the (now fixed-size) shape
    bodyPr = txBody.find(qn("a:bodyPr"))
    if bodyPr is not None:
        for tag in (qn("a:normAutofit"), qn("a:spAutoFit"), qn("a:noAutofit")):
            el = bodyPr.find(tag)
            if el is not None:
                bodyPr.remove(el)
        etree.SubElement(bodyPr, qn("a:normAutofit"))

    n   = max(1, len(lines))
    spc = {1: 2400, 2: 1800, 3: 1200, 4: 900, 5: 500}.get(n, 300)

    for line in lines:
        ap  = etree.SubElement(txBody, qn("a:p"))
        pPr = etree.SubElement(ap, qn("a:pPr"))
        pPr.set("lvl", "0")
        spcBef = etree.SubElement(pPr, qn("a:spcBef"))
        spcPts = etree.SubElement(spcBef, qn("a:spcPts"))
        spcPts.set("val", str(spc))
        r   = etree.SubElement(ap, qn("a:r"))
        rPr = etree.SubElement(r, qn("a:rPr"))
        rPr.set("lang", "en-US")
        rPr.set("dirty", "0")
        t   = etree.SubElement(r, qn("a:t"))
        t.text = line

    # ── Strip <p:ph> — converts placeholder to a plain shape so LibreOffice
    # never overlays layout hint text in rendered frames ───────────────────────
    nvPr = ph._element.find(".//" + qn("p:nvPr"))
    if nvPr is not None:
        ph_el = nvPr.find(qn("p:ph"))
        if ph_el is not None:
            nvPr.remove(ph_el)

    return ph


def _write_notes(slide, text: str):
    slide.notes_slide.notes_text_frame.text = text or ""


def _remove_ph(slide, idx):
    ph = _ph(slide, idx)
    if ph is not None:
        ph._element.getparent().remove(ph._element)


def _borrow_bg_picture(slide, layout):
    """Copy full-bleed background image from layout into slide (for LibreOffice)."""
    layout_spTree = layout._element.find(".//" + qn("p:spTree"))
    if layout_spTree is None:
        return
    slide_spTree = slide._element.find(".//" + qn("p:spTree"))
    insert_idx = 2
    for elem in layout_spTree:
        if elem.tag.split("}")[-1] != "pic":
            continue
        xfrm = elem.find(".//" + qn("a:xfrm"))
        ext  = xfrm.find(qn("a:ext")) if xfrm is not None else None
        if ext is None or int(ext.get("cx", "0")) < Inches(10):
            continue
        pic_copy = deepcopy(elem)
        for el in pic_copy.iter():
            old_rId = el.get(_EMBED_ATTR)
            if old_rId:
                try:
                    rel = layout.part.rels[old_rId]
                    new_rId = slide.part.relate_to(rel.target_part, rel.reltype)
                    el.set(_EMBED_ATTR, new_rId)
                except (KeyError, AttributeError):
                    pass
        slide_spTree.insert(insert_idx, pic_copy)
        break


# ── Animation helpers ──────────────────────────────────────────────────────────

def _build_timing_root(slide_el):
    """
    Create (or replace) the <p:timing> element on the slide and return
    (timing, childTnLst_main, seq, bldLst, id_counter).
    The caller appends click-effect entries to childTnLst_main and
    bldP entries to bldLst.
    """
    existing = slide_el.find(qn("p:timing"))
    if existing is not None:
        slide_el.remove(existing)

    _id = [1]
    def next_id():
        v = _id[0]; _id[0] += 1; return str(v)

    timing = etree.SubElement(slide_el, qn("p:timing"))
    tnLst  = etree.SubElement(timing, qn("p:tnLst"))

    root_par = etree.SubElement(tnLst, qn("p:par"))
    root_cTn = etree.SubElement(root_par, qn("p:cTn"))
    root_cTn.set("id", next_id())
    root_cTn.set("dur", "indefin")
    root_cTn.set("restart", "whenNotActive")
    root_cTn.set("nodeType", "tmRoot")
    childTnLst_root = etree.SubElement(root_cTn, qn("p:childTnLst"))

    seq = etree.SubElement(childTnLst_root, qn("p:seq"))
    seq.set("concurrent", "1")
    seq.set("nextAc", "seek")

    main_cTn = etree.SubElement(seq, qn("p:cTn"))
    main_cTn.set("id", next_id())
    main_cTn.set("dur", "indefin")
    main_cTn.set("nodeType", "mainSeq")
    childTnLst_main = etree.SubElement(main_cTn, qn("p:childTnLst"))

    prevCL = etree.SubElement(seq, qn("p:prevCondLst"))
    prev_c = etree.SubElement(prevCL, qn("p:cond"))
    prev_c.set("evt", "onPrevClick")
    prev_c.set("delay", "0")
    etree.SubElement(prev_c, qn("p:tn"))

    bldLst = etree.SubElement(timing, qn("p:bldLst"))
    return childTnLst_main, bldLst, next_id


def _append_click_effect(childTnLst_main, next_id, grp_idx, spid,
                          para_idx=None):
    """
    Append one click-triggered Appear effect to the main sequence.
    If para_idx is not None → targets a specific paragraph via <p:txEl>.
    Otherwise → targets the whole shape.
    """
    outer_par = etree.SubElement(childTnLst_main, qn("p:par"))
    outer_cTn = etree.SubElement(outer_par, qn("p:cTn"))
    outer_cTn.set("id", next_id())
    outer_cTn.set("fill", "hold")
    stCondLst = etree.SubElement(outer_cTn, qn("p:stCondLst"))
    cond = etree.SubElement(stCondLst, qn("p:cond"))
    cond.set("delay", "indefin")
    inner_list = etree.SubElement(outer_cTn, qn("p:childTnLst"))

    inner_par = etree.SubElement(inner_list, qn("p:par"))
    inner_cTn = etree.SubElement(inner_par, qn("p:cTn"))
    inner_cTn.set("id", next_id())
    inner_cTn.set("presetID", "1")      # 1 = Appear
    inner_cTn.set("presetClass", "entr")
    inner_cTn.set("presetSubtype", "0")
    inner_cTn.set("fill", "hold")
    inner_cTn.set("grpId", str(grp_idx))
    inner_cTn.set("nodeType", "clickEffect")
    st2 = etree.SubElement(inner_cTn, qn("p:stCondLst"))
    c2  = etree.SubElement(st2, qn("p:cond"))
    c2.set("delay", "0")
    eff_list = etree.SubElement(inner_cTn, qn("p:childTnLst"))

    set_el = etree.SubElement(eff_list, qn("p:set"))
    cBhvr  = etree.SubElement(set_el, qn("p:cBhvr"))
    cTn_b  = etree.SubElement(cBhvr, qn("p:cTn"))
    cTn_b.set("id", next_id())
    cTn_b.set("dur", "1")
    cTn_b.set("fill", "hold")
    tgtEl  = etree.SubElement(cBhvr, qn("p:tgtEl"))
    spTgt  = etree.SubElement(tgtEl, qn("p:spTgt"))
    spTgt.set("spid", str(spid))

    if para_idx is not None:
        txEl = etree.SubElement(spTgt, qn("p:txEl"))
        pRng = etree.SubElement(txEl, qn("p:pRng"))
        pRng.set("st", str(para_idx))
        pRng.set("end", str(para_idx))

    attrNms = etree.SubElement(cBhvr, qn("p:attrNameLst"))
    attrNm  = etree.SubElement(attrNms, qn("p:attrName"))
    attrNm.text = "style.visibility"
    to_el  = etree.SubElement(set_el, qn("p:to"))
    strVal = etree.SubElement(to_el, qn("p:strVal"))
    strVal.set("val", "visible")


def _add_appear_per_bullet(slide, shape, n_bullets: int):
    """Each paragraph in 'shape' appears on a separate click (Appear effect)."""
    if n_bullets < 1:
        return
    spid = shape.shape_id
    main_list, bldLst, next_id = _build_timing_root(slide._element)
    for i in range(n_bullets):
        _append_click_effect(main_list, next_id, i, spid, para_idx=i)
    bldP = etree.SubElement(bldLst, qn("p:bldP"))
    bldP.set("spid", str(spid))
    bldP.set("grpId", "0")
    bldP.set("animBg", "1")
    bldP.set("build", "p")


def _add_appear_shapes(slide, shapes):
    """
    Each shape in 'shapes' appears on a separate click (whole-shape Appear).
    Builds a single timing block across all shapes.
    """
    if not shapes:
        return
    main_list, _bldLst, next_id = _build_timing_root(slide._element)
    for i, shape in enumerate(shapes):
        _append_click_effect(main_list, next_id, i, shape.shape_id)


# ── Table extraction from .docx ────────────────────────────────────────────────

def _parse_storyboard_table(docx_path: str) -> list[dict]:
    """
    Find the storyboard table (Frame | Type | Narration Chunk | OST | Visual Cues)
    and return a list of row dicts.
    """
    doc = Document(docx_path)
    target = None
    for table in doc.tables:
        if len(table.columns) >= 4:
            header_cells = [c.text.strip().lower() for c in table.rows[0].cells]
            if any("narration" in h for h in header_cells):
                target = table
                break

    if target is None:
        raise ValueError("Storyboard table not found in document")

    rows = []
    for row in target.rows[1:]:
        cells = [c.text.strip() for c in row.cells]
        while len(cells) < 5:
            cells.append("")
        rows.append({
            "frame":      cells[0],
            "type":       cells[1],
            "narration":  cells[2],
            "ost":        cells[3],
            "visual_cue": cells[4],
        })
    return rows


def _is_faculty_cue(visual_cue: str) -> bool:
    """Return True when the visual cue means 'show faculty video'."""
    return bool(re.search(r"faculty", visual_cue, re.IGNORECASE))


def _has_ost(ost: str) -> bool:
    return ost.strip() not in ("", "—", "-", "–")


def _parse_ost(ost: str) -> tuple[str, list[str]]:
    """
    Split OST into (title, [bullet lines]).
    Preserves the exact text of each line — including arrows, dashes, etc.
    Blank separator lines are dropped; trailing whitespace is stripped.
    """
    if not ost or ost.strip() in ("—", "-", "–", ""):
        return "", []
    lines = [ln.rstrip() for ln in ost.splitlines()]   # rstrip only — keep leading indent
    lines = [ln for ln in lines if ln.strip()]           # drop blank separator lines
    lines = [ln for ln in lines if ln.strip() not in ("—", "-", "–")]
    if not lines:
        return "", []
    return lines[0].strip(), [ln.strip() for ln in lines[1:]]


_ABBR_FIX = re.compile(r'\b(ai|ml|roi|kpi|ost|vo|api)\b', re.IGNORECASE)

def _title_case(text: str) -> str:
    """Sentence case with common abbreviations restored to upper."""
    s = text[0].upper() + text[1:] if text else text
    return _ABBR_FIX.sub(lambda m: m.group().upper(), s)


def _derive_topic_from_narration(narration: str) -> tuple[str, str]:
    """
    Return (label, title) for the wide-slide banner.
    - 'label' is a short tag like 'PART 3' (may be empty).
    - 'title' is the main heading text.
    """
    # "Part one/two/three..., heading. Body..."
    m = re.match(
        r'^(part\s+(?:one|two|three|four|five|six|seven|eight|nine|ten|\d+))'
        r'\s*[,\.]\s*([^.]+)',
        narration, re.IGNORECASE,
    )
    if m:
        part_raw = m.group(1).strip()
        heading  = m.group(2).strip().rstrip(".")
        # Convert word to digit for the label: Part Three → PART 3
        _words = {"one":1,"two":2,"three":3,"four":4,"five":5,
                  "six":6,"seven":7,"eight":8,"nine":9,"ten":10}
        last = part_raw.split()[-1].lower()
        num  = _words.get(last, last)
        label = f"PART {num}"
        return label, _title_case(heading)

    # "we explore/cover/examine how to X"
    m2 = re.search(
        r'(?:explore|discuss|cover|examine|look at)\s+how\s+to\s+([^.]+)',
        narration, re.IGNORECASE,
    )
    if m2:
        return "", _title_case(m2.group(1).strip().rstrip("."))

    # "In this video, we X"
    m3 = re.search(r'in this video[,\s]+we\s+([^.]+)', narration, re.IGNORECASE)
    if m3:
        return "", _title_case(m3.group(1).strip().rstrip("."))


    # Fallback: first sentence, capped at 12 words
    dot   = narration.find(".")
    first = narration[:dot].strip() if dot > 0 else narration[:80].strip()
    words = first.split()
    return "", (" ".join(words[:12]) + ("…" if len(words) > 12 else ""))


def _textbox(slide, text, left, top, width, height,
             size=18, bold=False, color=RGBColor(0xFF, 0xFF, 0xFF),
             align=PP_ALIGN.LEFT):
    tb  = slide.shapes.add_textbox(left, top, width, height)
    tf  = tb.text_frame
    tf.word_wrap = True
    para = tf.paragraphs[0]
    para.alignment = align
    run = para.add_run()
    run.text       = text
    run.font.size  = Pt(size)
    run.font.bold  = bold
    run.font.color.rgb = color
    return tb


def _extract_module_title(docx_path: str) -> str:
    """Best-effort: extract a short module label from table metadata."""
    doc = Document(docx_path)
    for table in doc.tables:
        if len(table.columns) == 1:
            for row in table.rows:
                text = row.cells[0].text.strip()
                m = re.search(r"([A-Z]+[_ ][A-Z]+[_ ][A-Z]+[_ ][A-Z]\d+)", text)
                if m:
                    return m.group(1).replace("_", " ")
    return "Module 9"


# ── Slide builders ─────────────────────────────────────────────────────────────

def _clear_layout_ph_prompt(layout, idx: int):
    """
    Erase the prompt text (e.g. 'Click to edit text.') from a layout placeholder.
    LibreOffice headless renders layout prompt text OVER slide content even when
    the slide placeholder already has real text; clearing it at the layout level
    is the only reliable fix.
    """
    for sp in layout.placeholders:
        if sp.placeholder_format.idx != idx:
            continue
        txBody = sp._element.find(qn("p:txBody"))
        if txBody is None:
            break
        for ap in list(txBody.findall(qn("a:p"))):
            txBody.remove(ap)
        ep = etree.SubElement(txBody, qn("a:p"))
        etree.SubElement(ep, qn("a:endParaRPr")).set("lang", "en-US")
        break


def _build_wide_slide(prs, row, banner):
    """
    1D_Wide Screen with a centred dark-navy band and topic title.
    The label ('PART N') is shown in gold above the title when present.
    """
    layout = _find_layout(prs, L_WIDE)
    slide  = prs.slides.add_slide(layout)
    _borrow_bg_picture(slide, layout)

    label, title = _derive_topic_from_narration(row["narration"])

    # Dark navy band across the full width (vertically centred)
    band = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(0), Inches(2.2), Inches(13.33), Inches(2.8),
    )
    band.fill.solid()
    band.fill.fore_color.rgb = RGBColor(0x0D, 0x1B, 0x2A)
    band.line.fill.background()

    if label:
        _textbox(slide, label,
                 Inches(0.8), Inches(2.35), Inches(11.0), Inches(0.5),
                 size=14, bold=True, color=RGBColor(0xFF, 0xC0, 0x00))

    title_top = Inches(2.9) if label else Inches(2.65)
    _textbox(slide, title,
             Inches(0.8), title_top, Inches(11.0), Inches(1.8),
             size=36, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

    _write_notes(slide, row["narration"])
    return slide


def _build_content_slide(prs, row, banner, title, bullets, img_path=None):
    """
    60:40 bullets layout: title + bullets (verbatim OST) on the left,
    Imagen image on the right.  Each bullet gets an Appear-on-click animation.
    Ghost-text suppression is handled inside _fill_bullets_ph (strips <p:ph>).
    """
    layout = _find_layout(prs, L_BULLETS)
    slide  = prs.slides.add_slide(layout)

    _fill_ph(slide, PH_BULLETS_BANNER, banner, size=11)
    _fill_ph(slide, PH_BULLETS_TITLE, title,
             bold=True, italic=False, color=RGBColor(0x1A, 0x1A, 0x2E))
    ph = _fill_bullets_ph(slide, PH_BULLETS_BODY, bullets)

    # Remove every placeholder we did not fill so LibreOffice doesn't
    # render stray empty boxes (icon placeholder, picture placeholder, etc.)
    _keep = {PH_BULLETS_BANNER, PH_BULLETS_TITLE, PH_BULLETS_BODY}
    for p in list(slide.placeholders):
        if p.placeholder_format.idx not in _keep:
            p._element.getparent().remove(p._element)

    if img_path and Path(img_path).exists():
        slide.shapes.add_picture(str(img_path), _RP_LEFT, _RP_TOP, _RP_W, _RP_IMG_H)

    _write_notes(slide, row["narration"])

    if ph and bullets:
        _add_appear_per_bullet(slide, ph, len(bullets))

    return slide


def _build_comparison_slide(prs, row, banner, title, bullets, img_path=None):
    """
    60:40 Text boxes with headings layout: 3 rows of heading + body text.
    Each 'X → Y' bullet is split at '→' to produce a heading/body pair.
    """
    layout = _find_layout(prs, L_COMPARE)

    # Clear layout placeholder prompts for all text boxes in this layout
    for idx in [PH_CMP_TITLE] + [i for pair in _CMP_ROWS for i in pair]:
        _clear_layout_ph_prompt(layout, idx)

    slide = prs.slides.add_slide(layout)

    _fill_ph(slide, PH_CMP_BANNER, banner, size=11)
    _fill_ph(slide, PH_CMP_TITLE, title,
             bold=True, italic=False, color=RGBColor(0x1A, 0x1A, 0x2E))

    for i, (h_idx, b_idx) in enumerate(_CMP_ROWS):
        if i >= len(bullets):
            break
        parts = bullets[i].split("→", 1)
        heading = parts[0].strip()
        body    = parts[1].strip() if len(parts) > 1 else ""
        _fill_ph(slide, h_idx, heading, bold=True, color=RGBColor(0x1A, 0x1A, 0x2E))
        if body:
            _fill_ph(slide, b_idx, body, color=RGBColor(0x44, 0x44, 0x55))

    # Remove every unfilled placeholder
    _keep = {PH_CMP_BANNER, PH_CMP_TITLE} | {i for pair in _CMP_ROWS for i in pair}
    for p in list(slide.placeholders):
        if p.placeholder_format.idx not in _keep:
            p._element.getparent().remove(p._element)

    if img_path and Path(img_path).exists():
        slide.shapes.add_picture(str(img_path), _RP_LEFT, _RP_TOP, _RP_W, _RP_IMG_H)

    _write_notes(slide, row["narration"])
    return slide


def _borrow_layout_shapes(slide, layout):
    """Copy all non-placeholder shapes from layout into the slide's spTree (LibreOffice compat)."""
    layout_spTree = layout._element.find(".//" + qn("p:spTree"))
    slide_spTree  = slide._element.find(".//" + qn("p:spTree"))
    if layout_spTree is None or slide_spTree is None:
        return
    for elem in layout_spTree:
        tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
        if tag not in ("sp", "pic", "grpSp", "graphicFrame"):
            continue
        if tag == "sp":
            nvPr = elem.find(".//" + qn("p:nvPr"))
            if nvPr is not None and nvPr.find(qn("p:ph")) is not None:
                continue
        elem_copy = deepcopy(elem)
        for el in elem_copy.iter():
            old_rId = el.get(_EMBED_ATTR)
            if old_rId:
                try:
                    rel = layout.part.rels[old_rId]
                    el.set(_EMBED_ATTR, slide.part.relate_to(rel.target_part, rel.reltype))
                except (KeyError, AttributeError):
                    pass
        slide_spTree.append(elem_copy)


def _build_bookend_slide(prs):
    """Create a blank intro/outro slide from the 1B_Intro-Outro layout."""
    layout = _find_layout(prs, L_BOOKEND)
    slide  = prs.slides.add_slide(layout)
    _borrow_layout_shapes(slide, layout)
    return slide


# ── Image prompts per frame ────────────────────────────────────────────────────

_IMAGE_PROMPTS = {
    "2": (
        "Professional educational illustration for an AI project management slide "
        "about types of project lifecycle: predictive, iterative, incremental, agile "
        "and hybrid approaches shown as abstract diagram shapes. Modern flat design, "
        "deep blue and gold colour palette, clean background, absolutely no text or letters."
    ),
    "3": (
        "Professional educational illustration comparing AI project lifecycle types: "
        "research and ideation projects with agile loops, deployment and scaling projects "
        "with structured milestones, and hybrid combined flows. Abstract diagram concept, "
        "modern flat design, blue tones, no text or letters."
    ),
    "5": (
        "Professional educational illustration for a reflective thinking moment about "
        "choosing the right AI delivery approach — abstract concept of a decision fork "
        "or thoughtful pause. Minimal modern design, muted blue and teal tones, no text."
    ),
}


def _generate_images(rows, images_dir: Path, gemini_api_key: str) -> dict:
    """Generate one image per W+Img frame. Returns {frame_str: path_or_None}."""
    try:
        from poc3_image_gen import generate_slide_image
    except ImportError:
        print("[WARN] poc3_image_gen not importable — skipping image generation")
        return {}

    paths = {}
    for row in rows:
        frame = row["frame"]
        prompt = _IMAGE_PROMPTS.get(frame)
        if not prompt:
            continue
        out = str(images_dir / f"m9_frame_{frame}.png")
        print(f"  [frame {frame}] Generating image …", end=" ", flush=True)
        p = generate_slide_image(prompt, out, gemini_api_key)
        paths[frame] = p
        print("done" if p else "FAILED")
    return paths


# ── Frame router ───────────────────────────────────────────────────────────────

def _build_slide_for_row(prs, row, banner, img_path=None):
    frame_type = row["type"].strip().lower()
    ost        = row["ost"]
    visual_cue = row["visual_cue"]
    cue_lower  = visual_cue.lower()

    # Wide screen — faculty/title-only slides
    if "wide" in frame_type and "img" not in frame_type:
        if _is_faculty_cue(visual_cue) and _has_ost(ost):
            title, bullets = _parse_ost(ost)
            if not bullets:
                bullets, title = [title], ""
            return _build_content_slide(prs, row, banner, title, bullets, img_path)
        return _build_wide_slide(prs, row, banner)

    # Parse OST for all W+Img slides
    title, bullets = _parse_ost(ost)
    if not bullets:
        bullets, title = [title], ""

    # Comparison / multi-column visual cue → 3-row heading+body layout
    if ("comparison" in cue_lower or "3-column" in cue_lower
            or ("column" in cue_lower and "boxes" in cue_lower)):
        return _build_comparison_slide(prs, row, banner, title, bullets, img_path)

    return _build_content_slide(prs, row, banner, title, bullets, img_path)


# ── Clear template slides ──────────────────────────────────────────────────────

def _clear_slides(prs):
    sldIdLst = prs.slides._sldIdLst
    rids = [el.get(qn("r:id")) for el in list(sldIdLst)]
    for rId in rids:
        if rId:
            try:
                prs.part.drop_rel(rId)
            except Exception:
                pass
    for child in list(sldIdLst):
        sldIdLst.remove(child)


# ── Main ───────────────────────────────────────────────────────────────────────

def build(docx_path: str, output_pptx: str,
          template_path: str | None = None,
          gemini_api_key: str | None = None) -> str:

    script_dir    = Path(__file__).parent
    template_path = template_path or str(script_dir.parent / "POC_3.pptx")

    print(f"[build] Parsing table from {Path(docx_path).name} …")
    rows   = _parse_storyboard_table(docx_path)
    banner = _extract_module_title(docx_path)
    print(f"[build] Module label: {banner!r}  |  {len(rows)} frame(s) found")

    # ── Image generation ──────────────────────────────────────────────────────
    out_path   = Path(output_pptx)
    images_dir = out_path.parent / "images"
    images_dir.mkdir(parents=True, exist_ok=True)

    img_paths = {}
    if gemini_api_key:
        print(f"[build] Generating slide images with Gemini Imagen …")
        img_paths = _generate_images(rows, images_dir, gemini_api_key)
    else:
        # Reuse any already-generated images for this module
        for frame in _IMAGE_PROMPTS:
            p = images_dir / f"m9_frame_{frame}.png"
            if p.exists():
                img_paths[frame] = str(p)
        if img_paths:
            print(f"[build] Reusing {len(img_paths)} existing image(s)")
        else:
            print("[build] No Gemini key — slides will have no right-panel images")

    # ── Build slides ──────────────────────────────────────────────────────────
    print(f"[build] Loading template {Path(template_path).name} …")
    prs = Presentation(template_path)
    _clear_slides(prs)

    # Intro bookend (same 1B_Intro-Outro design as template slide 0)
    _build_bookend_slide(prs)
    print("  [intro]  1B_Intro-Outro")

    for row in rows:
        frame = row["frame"]
        img   = img_paths.get(frame)
        print(f"  Frame {frame}  [{row['type']:15s}]  img={'yes' if img else 'no ':3s}  "
              f"cue={row['visual_cue'][:35]!r}")
        _build_slide_for_row(prs, row, banner, img_path=img)

    # Outro bookend (same design, mirrors intro)
    _build_bookend_slide(prs)
    print("  [outro]  1B_Intro-Outro")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    size_kb = out_path.stat().st_size / 1024
    print(f"\n[build] Saved {len(prs.slides)} slides → {out_path}  ({size_kb:.0f} KB)")
    return str(out_path)


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
    import os
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    _load_env()
    result = build(
        sys.argv[1], sys.argv[2],
        template_path=sys.argv[3] if len(sys.argv) > 3 else None,
        gemini_api_key=os.environ.get("GEMINI_API_KEY"),
    )
    print(f"\nOutput: {result}")
