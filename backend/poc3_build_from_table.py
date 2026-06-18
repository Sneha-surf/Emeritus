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
import math
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
L_ICON3    = "60:40 Icon and text "   # trailing space is in template
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

# ── Placeholder indices — icon-3-row layout ("60:40 Icon and text") ───────────
# PH[34] = slide title (top banner position, y=0.192")
# PH[26/29/30] = circular icon picture-placeholders for rows 1/2/3
# PH[37/39/41] = row headings;  PH[38/40/42] = row body text
_ICON3_TITLE = 34
_ICON3_ROWS  = [(37, 38), (39, 40), (41, 42)]
_ICON3_ICONS = [26, 29, 30]

# ── Right-panel image sizing ──────────────────────────────────────────────────
# Images are 1408×768 (h/w ratio = 768/1408).
# _add_rp_image() sizes each image to match its left content area.
_RP_IMG_W_PX, _RP_IMG_H_PX = 1408, 768   # actual source dimensions


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
    Render bullet lines inside the template's rounded-rect container.

    Strategy: keep the placeholder shape (for the purple rounded-rect border
    inherited from the layout) but leave it text-free.  Bullet content goes
    into a free textbox placed exactly over the placeholder.  LibreOffice
    respects anchor="ctr" on free shapes but ignores it on placeholder shapes,
    so this split is necessary for vertical centering in the video renderer.
    """
    ph = _ph(slide, idx)
    if ph is None:
        return None

    # ── read layout geometry for PH[idx] and PH[0] (title) ──────────────────
    lx41, ly41, lw41, lh41 = 1319951, 2080941, 4420899, 3335098  # EMU fallbacks
    title_x = 407823   # PH[0] x fallback = 0.446"
    for lph in slide.slide_layout.placeholders:
        li    = lph.placeholder_format.idx
        lxfrm = lph._element.find(".//" + qn("a:xfrm"))
        if lxfrm is None:
            continue
        lo = lxfrm.find(qn("a:off")); le = lxfrm.find(qn("a:ext"))
        if lo is None or le is None:
            continue
        if li == idx:
            lx41 = int(lo.get("x", lx41)); ly41 = int(lo.get("y", ly41))
            lw41 = int(le.get("cx", lw41)); lh41 = int(le.get("cy", lh41))
        elif li == 0:
            title_x = int(lo.get("x", title_x))

    new_x = title_x
    new_w = lw41 + lx41 - new_x

    # ── placeholder: clear text, keep border styling, set xfrm ──────────────
    txBody = ph._element.find(qn("p:txBody"))
    if txBody is not None:
        for ap in list(txBody.findall(qn("a:p"))):
            txBody.remove(ap)
        etree.SubElement(txBody, qn("a:p"))   # one empty para — LO renders the border

    spPr = ph._element.find(qn("p:spPr"))
    if spPr is None:
        spPr = etree.SubElement(ph._element, qn("p:spPr"))
    for el in list(spPr.findall(qn("a:xfrm"))):
        spPr.remove(el)
    xfrm = etree.SubElement(spPr, qn("a:xfrm"))
    off  = etree.SubElement(xfrm, qn("a:off"))
    off.set("x", str(new_x)); off.set("y", str(ly41))
    ext  = etree.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", str(new_w)); ext.set("cy", str(lh41))

    # ── free textbox: bullet content, position-based vertical centering ──────
    # LibreOffice ignores anchor="ctr" and tIns on free textboxes in practice.
    # Instead we position the textbox so its TOP starts at the visual centre.
    # The placeholder shape (full-height) still provides the purple border.
    # Visual text height = (spcBef × n) + (total_wrapped_lines × font_size).
    n   = max(1, len(lines))
    spc = {1: 3200, 2: 2400, 3: 1800, 4: 1200, 5: 800}.get(n, 400)

    text_w_pt      = max(10, (new_w - 585216) / 914400 * 72)   # usable text width in pts
    chars_per_line = max(10, int(text_w_pt / (18 * 0.5)))       # ≈ chars that fit per line
    total_lines    = sum(max(1, math.ceil(max(1, len(l)) / chars_per_line)) for l in lines)
    spc_pt         = spc / 100                                   # spcBef in points

    # content_depth = first-to-last visible text pixel.  Empirically: +1pt/line vs formula
    # due to LibreOffice's default line leading.  Also subtract 2× the fixed internal
    # offset (tIns + LO default leading-above-first-line ≈ 0.111") so that the space
    # above the first text pixel equals the space below the last text pixel.
    content_h_emu  = int((spc_pt * n + total_lines * 19) * 12700)   # +1pt/line correction
    _LO_INTERNAL   = 203000                                          # 2 × 0.111" × 914400
    v_offset_emu   = max(0, (lh41 - content_h_emu - _LO_INTERNAL) // 2)

    # Textbox y = ly41 + offset (centred), height = lh41 - offset (fills to box bottom)
    tb_y = ly41 + v_offset_emu
    tb_h = lh41 - v_offset_emu

    txBox = slide.shapes.add_textbox(Emu(new_x), Emu(tb_y), Emu(new_w), Emu(tb_h))
    txBox.text_frame.word_wrap = True

    txBody2 = txBox._element.find(qn("p:txBody"))
    bodyPr2 = txBody2.find(qn("a:bodyPr"))
    if bodyPr2 is not None:
        bodyPr2.set("tIns", "45720")
        bodyPr2.set("bIns", "45720")
        for tag in (qn("a:normAutofit"), qn("a:spAutoFit"), qn("a:noAutofit")):
            el = bodyPr2.find(tag)
            if el is not None:
                bodyPr2.remove(el)
        etree.SubElement(bodyPr2, qn("a:noAutofit"))

    for ap in list(txBody2.findall(qn("a:p"))):
        txBody2.remove(ap)

    for line in lines:
        ap  = etree.SubElement(txBody2, qn("a:p"))
        pPr = etree.SubElement(ap, qn("a:pPr"))
        pPr.set("lvl", "0")
        pPr.set("algn", "l")
        pPr.set("marL", "585216")    # 0.64" left indent
        pPr.set("indent", "-260350")
        spcBef = etree.SubElement(pPr, qn("a:spcBef"))
        etree.SubElement(spcBef, qn("a:spcPts")).set("val", str(spc))
        buFont = etree.SubElement(pPr, qn("a:buFont"))
        buFont.set("typeface", "Arial")
        etree.SubElement(pPr, qn("a:buChar")).set("char", "•")
        r   = etree.SubElement(ap, qn("a:r"))
        rPr = etree.SubElement(r, qn("a:rPr"))
        rPr.set("lang", "en-US"); rPr.set("dirty", "0"); rPr.set("sz", "1800")
        sf  = etree.SubElement(rPr, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr")).set("val", "1A1A2E")
        t   = etree.SubElement(r, qn("a:t"))
        t.text = line

    return txBox   # animations attach to the free textbox


def _write_notes(slide, text: str):
    slide.notes_slide.notes_text_frame.text = text or ""


def _remove_ph(slide, idx):
    ph = _ph(slide, idx)
    if ph is not None:
        ph._element.getparent().remove(ph._element)


def _clean_layout_prompts(prs):
    """
    Wipe stale text from PH[34] (banner) across ALL layouts and PH[41] (body)
    from the bullets layout only.

    LibreOffice merges the LAYOUT's <a:p> nodes with the slide's during rendering,
    so 'Click to edit text' or the template's stale 'AIM PGC LAI M9' course-code
    label bleeds through as ghost text even when the slide placeholder has content.
    PH[34] is cleared from every layout because both the bullets layout and the
    comparison layout carry the stale module-title text there.
    """
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            is_bullets = layout.name == L_BULLETS
            for ph in layout.placeholders:
                idx = ph.placeholder_format.idx
                # Clean PH[34] everywhere; PH[41] only on the bullets layout
                if idx == PH_BULLETS_BANNER:
                    pass  # always clean
                elif idx == PH_BULLETS_BODY and is_bullets:
                    pass  # clean body only on bullets layout
                else:
                    continue
                txBody = ph._element.find(qn("p:txBody"))
                if txBody is None:
                    continue
                for ap in list(txBody.findall(qn("a:p"))):
                    txBody.remove(ap)
                ap = etree.SubElement(txBody, qn("a:p"))
                etree.SubElement(ap, qn("a:endParaRPr")).set("lang", "en-US")


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


def _borrow_wide_screen_pics(slide, layout):
    """Copy ALL full-bleed pictures from layout into slide (corridor BG + faculty PNG)."""
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
        insert_idx += 1   # copy ALL wide pictures, not just the first


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
                          para_idx=None, anim_type="appear"):
    """
    Append one click-triggered entrance animation to the main sequence.
    anim_type: 'appear' | 'wipe' | 'fade'
    If para_idx is not None → targets a specific paragraph via <p:txEl>.
    Otherwise → targets the whole shape.
    """
    _PRESET_ID  = {"appear": "1",  "wipe": "21", "fade": "10"}
    _PRESET_SUB = {"appear": "0",  "wipe": "4",  "fade": "0"}
    _PRESET_DUR = {"appear": "1",  "wipe": "500","fade": "500"}
    _FILTER     = {"wipe": "wipe(left)", "fade": "fade"}

    pid = _PRESET_ID.get(anim_type, "1")
    pst = _PRESET_SUB.get(anim_type, "0")
    dur = _PRESET_DUR.get(anim_type, "1")

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
    inner_cTn.set("presetID", pid)
    inner_cTn.set("presetClass", "entr")
    inner_cTn.set("presetSubtype", pst)
    inner_cTn.set("fill", "hold")
    inner_cTn.set("grpId", str(grp_idx))
    inner_cTn.set("nodeType", "clickEffect")
    inner_cTn.set("dur", dur)
    st2 = etree.SubElement(inner_cTn, qn("p:stCondLst"))
    c2  = etree.SubElement(st2, qn("p:cond"))
    c2.set("delay", "0")
    eff_list = etree.SubElement(inner_cTn, qn("p:childTnLst"))

    flt = _FILTER.get(anim_type)
    if flt:
        # Wipe / Fade: use p:animEffect for proper visual animation in PowerPoint
        anim_el = etree.SubElement(eff_list, qn("p:animEffect"))
        anim_el.set("transition", "in")
        anim_el.set("filter", flt)
        cBhvr2 = etree.SubElement(anim_el, qn("p:cBhvr"))
        cTn_b2 = etree.SubElement(cBhvr2, qn("p:cTn"))
        cTn_b2.set("id", next_id())
        cTn_b2.set("dur", dur)
        tgt2   = etree.SubElement(cBhvr2, qn("p:tgtEl"))
        spTgt2 = etree.SubElement(tgt2, qn("p:spTgt"))
        spTgt2.set("spid", str(spid))
        if para_idx is not None:
            txEl2 = etree.SubElement(spTgt2, qn("p:txEl"))
            pRng2 = etree.SubElement(txEl2, qn("p:pRng"))
            pRng2.set("st", str(para_idx))
            pRng2.set("end", str(para_idx))
    else:
        # Appear: simple visibility set (instant)
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


def _add_slide_entry_animations(slide, title_shape, img_shape, bullet_shape, n_bullets):
    """
    Build the animation sequence: Header (Wipe L→R) → Image (Fade In) → OST bullets (Wipe L→R).
    All on separate clicks so each element can be synced to narration in presentation mode.
    For the video pipeline, LO renders all shapes visible so frame-by-frame reveal
    is handled separately via XML bullet-blanking and PIL white-out.
    """
    main_list, bldLst, next_id = _build_timing_root(slide._element)
    grp = 0

    if title_shape is not None:
        _append_click_effect(main_list, next_id, grp, title_shape.shape_id,
                             anim_type="wipe")
        grp += 1

    if img_shape is not None:
        _append_click_effect(main_list, next_id, grp, img_shape.shape_id,
                             anim_type="fade")
        grp += 1

    if bullet_shape is not None and n_bullets > 0:
        spid = bullet_shape.shape_id
        for i in range(n_bullets):
            _append_click_effect(main_list, next_id, grp, spid,
                                 para_idx=i, anim_type="wipe")
            grp += 1
        bldP = etree.SubElement(bldLst, qn("p:bldP"))
        bldP.set("spid", str(spid))
        bldP.set("grpId", "0")
        bldP.set("animBg", "1")
        bldP.set("build", "p")


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


def _strip_part_heading(text: str) -> str:
    """
    Remove 'Part one/two/..., Heading sentence. ' prefix from narration so
    TTS does not read a section heading aloud as if it were prose.

    Matches patterns like:
      'Part one, type of project lifecycle. Rest of narration…'
      'Part three, managing the cone of uncertainty. Rest…'
    """
    m = re.match(
        r'^part\s+(?:one|two|three|four|five|six|seven|eight|nine|ten|\d+)'
        r'\s*[,\.]\s*[^.]+\.\s*',
        text, re.IGNORECASE,
    )
    return text[m.end():] if m else text


def _narration_to_bullets(narration: str, n: int = 4) -> list[str]:
    """
    Extract up to `n` sentence-form bullets from the narration text.
    Strips 'Part N, heading.' prefix first, then splits on sentence boundaries.
    Only keeps sentences >= 30 chars to filter out very short fragments.
    This produces educational, readable bullets like those in storyboard.pptx
    rather than the raw OST keywords (Predictive / Iterative / Incremental…).
    """
    text = _strip_part_heading(narration)
    sentences = re.split(r'(?<=[.!?])\s+', text.strip())
    return [s.strip() for s in sentences if len(s.strip()) >= 30][:n]


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


def _set_fill_alpha(shape, alpha_pct: int):
    """Set shape fill opacity. alpha_pct: 0=transparent, 100=fully opaque."""
    spPr = shape._element.find(qn("p:spPr"))
    if spPr is None:
        return
    solid = spPr.find(qn("a:solidFill"))
    if solid is None:
        return
    srgb = solid.find(qn("a:srgbClr"))
    if srgb is None:
        return
    for old in list(srgb.findall(qn("a:alpha"))):
        srgb.remove(old)
    a_el = etree.SubElement(srgb, qn("a:alpha"))
    a_el.set("val", str(alpha_pct * 1000))  # 100% = 100000


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


def _add_rp_image(slide, img_path, content_top_emu, content_h_emu):
    """
    Add a right-panel image sized to match the left content area.

    - Top edge aligns with content_top_emu.
    - Height matches content_h_emu using the actual 1408:768 aspect ratio.
      If that requires more width than the right panel offers, we clamp to the
      available width (the height then falls slightly short of the content block,
      but the top still aligns and the proportion stays correct).
    - Left edge starts 0.35" after the left content's right edge (~6.28" → 6.63").
    Returns the added picture shape.
    """
    _LEFT  = int(Inches(7.0))                                # starts after widest text box (6.667")
    _MAX_W = int(Inches(13.33)) - _LEFT - int(Inches(0.05))  # ≈ 6.28" to right edge
    w = min(int(content_h_emu * _RP_IMG_W_PX / _RP_IMG_H_PX), _MAX_W)
    h = int(w * _RP_IMG_H_PX / _RP_IMG_W_PX)
    return slide.shapes.add_picture(str(img_path), _LEFT, int(content_top_emu), w, h)


def _build_wide_slide(prs, row, banner):
    """
    1D_Wide Screen: faculty image as full-bleed background.
    Uses faculty_lady.png (the lady in the corridor) if available; falls back to
    copying corridor BG + faculty PNG from the layout (which shows the professor).
    For faculty-only slides (no OST, visual_cue contains 'faculty'), no text is shown.
    For other wide slides, a centred dark-navy band with topic title is added.
    """
    layout = _find_layout(prs, L_WIDE)
    slide  = prs.slides.add_slide(layout)

    # Prefer the lady's image (already contains the corridor background)
    _faculty_img = Path(__file__).parent.parent / "poc3_output" / "images" / "faculty_lady.png"
    if _faculty_img.exists():
        # Add as full-bleed background — add_picture appends, so call before other shapes
        slide.shapes.add_picture(str(_faculty_img), 0, 0, prs.slide_width, prs.slide_height)
    else:
        # Fallback: copy corridor + professor PNG from layout (for LO compatibility)
        _borrow_wide_screen_pics(slide, layout)

    visual_cue = row.get("visual_cue", "")
    ost        = row.get("ost", "")

    # Faculty-only slide: no text overlay — faculty speaks narration over the corridor
    if _is_faculty_cue(visual_cue) and not _has_ost(ost):
        _write_notes(slide, _strip_part_heading(row["narration"]))
        return slide

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

    _write_notes(slide, _strip_part_heading(row["narration"]))
    return slide


def _build_content_slide(prs, row, banner, title, bullets, img_path=None):
    """
    60:40 bullets layout: title + bullets on the left, Imagen image on the right.
    Each bullet gets an Appear-on-click animation for progressive reveal in the PPTX;
    the video pipeline renders progressive LO frames independently.

    PH[34] (module banner) is intentionally left empty — the template contains a
    stale course-code label that _clean_layout_prompts wipes at the layout level.
    Filling it here would override that clearing and bring back the ghost label.
    """
    layout = _find_layout(prs, L_BULLETS)
    slide  = prs.slides.add_slide(layout)

    # Title — fill PH[0] then slide it up to the banner position (y=0.192")
    # The banner (PH[34]) was at y=0.192" and the title (PH[0]) is at y=0.924"
    # by default.  Moving the title up fills the header area so it reads as the
    # primary slide heading rather than a sub-heading below an empty banner gap.
    _fill_ph(slide, PH_BULLETS_TITLE, title,
             bold=True, italic=False, color=RGBColor(0x1A, 0x1A, 0x2E))
    ph0 = _ph(slide, PH_BULLETS_TITLE)
    if ph0 is not None:
        spPr0 = ph0._element.find(qn("p:spPr"))
        if spPr0 is None:
            spPr0 = etree.SubElement(ph0._element, qn("p:spPr"))
        for el in list(spPr0.findall(qn("a:xfrm"))):
            spPr0.remove(el)
        for lph in layout.placeholders:
            if lph.placeholder_format.idx != PH_BULLETS_TITLE:
                continue
            lxfrm = lph._element.find(".//" + qn("a:xfrm"))
            if lxfrm is not None:
                lo = lxfrm.find(qn("a:off"))
                le = lxfrm.find(qn("a:ext"))
                if lo is not None and le is not None:
                    xfrm0 = etree.SubElement(spPr0, qn("a:xfrm"))
                    off0  = etree.SubElement(xfrm0, qn("a:off"))
                    # Move to banner y (0.192") — same x and width as layout
                    off0.set("x", lo.get("x", "0"))
                    off0.set("y", "175565")   # 0.192" × 914400 = 175565 EMU
                    ext0  = etree.SubElement(xfrm0, qn("a:ext"))
                    ext0.set("cx", le.get("cx", "0"))
                    ext0.set("cy", le.get("cy", "0"))
            break
    ph = _fill_bullets_ph(slide, PH_BULLETS_BODY, bullets)

    # Remove every placeholder we did not fill so LibreOffice doesn't
    # render stray empty boxes (icon placeholder, picture placeholder, etc.)
    _keep = {PH_BULLETS_TITLE, PH_BULLETS_BODY}
    for p in list(slide.placeholders):
        if p.placeholder_format.idx not in _keep:
            p._element.getparent().remove(p._element)

    img_shape = None
    if img_path and Path(img_path).exists():
        # Align image with the bullet box: top=2.276" = 2080941 EMU
        img_shape = _add_rp_image(slide, img_path,
                                  content_top_emu=2080941, content_h_emu=4500000)

    _write_notes(slide, _strip_part_heading(row["narration"]))

    # Animation: Header (Wipe L→R) → Image (Fade) → each bullet (Wipe L→R)
    title_ph = _ph(slide, PH_BULLETS_TITLE)
    if ph and bullets:
        _add_slide_entry_animations(slide, title_ph, img_shape, ph, len(bullets))

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

    # PH_CMP_BANNER (PH[34]) intentionally NOT filled — _clean_layout_prompts wiped
    # the stale 'AIM PGC LAI M9' label from the layout, so leaving it empty keeps
    # the banner area clean on both PPTX and video renderings.
    _fill_ph(slide, PH_CMP_TITLE, title,
             bold=True, italic=False, color=RGBColor(0x1A, 0x1A, 0x2E))

    # Move title to banner y (0.192") — same fix as _build_content_slide.
    # L_COMPARE layout also has PH[0] at y=0.924" (sub-heading area below the
    # empty banner gap). Repositioning it to y=175565 puts it in the header zone.
    ph0 = _ph(slide, PH_CMP_TITLE)
    if ph0 is not None:
        spPr0 = ph0._element.find(qn("p:spPr"))
        if spPr0 is None:
            spPr0 = etree.SubElement(ph0._element, qn("p:spPr"))
        for el in list(spPr0.findall(qn("a:xfrm"))):
            spPr0.remove(el)
        for lph in layout.placeholders:
            if lph.placeholder_format.idx != PH_CMP_TITLE:
                continue
            lxfrm = lph._element.find(".//" + qn("a:xfrm"))
            if lxfrm is not None:
                lo = lxfrm.find(qn("a:off"))
                le = lxfrm.find(qn("a:ext"))
                if lo is not None and le is not None:
                    xfrm0 = etree.SubElement(spPr0, qn("a:xfrm"))
                    off0  = etree.SubElement(xfrm0, qn("a:off"))
                    off0.set("x", lo.get("x", "0"))
                    off0.set("y", "175565")   # 0.192" × 914400 = 175565 EMU
                    ext0  = etree.SubElement(xfrm0, qn("a:ext"))
                    ext0.set("cx", le.get("cx", "0"))
                    ext0.set("cy", le.get("cy", "0"))
            break

    for i, (h_idx, b_idx) in enumerate(_CMP_ROWS):
        if i >= len(bullets):
            break
        parts = bullets[i].split("→", 1)
        heading = parts[0].strip()
        body    = parts[1].strip() if len(parts) > 1 else ""
        _fill_ph(slide, h_idx, heading, bold=True,
                 color=RGBColor(0x1A, 0x1A, 0x2E), align=PP_ALIGN.LEFT)
        if body:
            _fill_ph(slide, b_idx, body,
                     color=RGBColor(0x44, 0x44, 0x55), align=PP_ALIGN.LEFT)

    # Align row placeholders x to match title x (layout has rows indented ~1" vs title)
    title_x = 407824  # 0.446" × 914400 fallback
    for lph in layout.placeholders:
        if lph.placeholder_format.idx != PH_CMP_TITLE:
            continue
        lxfrm = lph._element.find(".//" + qn("a:xfrm"))
        if lxfrm is not None:
            lo = lxfrm.find(qn("a:off"))
            if lo is not None:
                title_x = int(lo.get("x", title_x))
        break

    for h_idx, b_idx in _CMP_ROWS:
        for ph_idx in (h_idx, b_idx):
            for lph in layout.placeholders:
                if lph.placeholder_format.idx != ph_idx:
                    continue
                lxfrm = lph._element.find(".//" + qn("a:xfrm"))
                if lxfrm is None:
                    break
                lo = lxfrm.find(qn("a:off"))
                le = lxfrm.find(qn("a:ext"))
                if lo is None or le is None:
                    break
                old_x = int(lo.get("x", 0))
                old_w = int(le.get("cx", 0))
                new_x = title_x
                new_w = (old_x + old_w) - title_x  # preserve right edge
                ph_sl = _ph(slide, ph_idx)
                if ph_sl is None:
                    break
                spPr = ph_sl._element.find(qn("p:spPr"))
                if spPr is None:
                    spPr = etree.SubElement(ph_sl._element, qn("p:spPr"))
                for el in list(spPr.findall(qn("a:xfrm"))):
                    spPr.remove(el)
                xfrm_new = etree.SubElement(spPr, qn("a:xfrm"))
                off_new  = etree.SubElement(xfrm_new, qn("a:off"))
                off_new.set("x", str(new_x))
                off_new.set("y", lo.get("y", "0"))
                ext_new  = etree.SubElement(xfrm_new, qn("a:ext"))
                ext_new.set("cx", str(new_w))
                ext_new.set("cy", le.get("cy", "0"))
                break

    # Remove every unfilled placeholder (PH_CMP_BANNER excluded — let _clean wipe it)
    _keep = {PH_CMP_TITLE} | {i for pair in _CMP_ROWS for i in pair}
    for p in list(slide.placeholders):
        if p.placeholder_format.idx not in _keep:
            p._element.getparent().remove(p._element)

    # Add a purple rounded-rect box outline around each row group (heading + body).
    # Body placeholder height is equalised to heading height and both text bodies
    # are center-anchored so the text appears visually centered in each box.
    _BOX_PAD = 76200   # ~0.083" padding on each side
    _BOX_ADJ = 6763    # corner-roundness matching template
    _spTree  = slide.shapes._spTree
    _insert_idx = 2    # behind all existing shapes (after nvGrpSpPr + grpSpPr)
    for h_idx, b_idx in _CMP_ROWS:
        h_ph = _ph(slide, h_idx)
        b_ph = _ph(slide, b_idx)
        if h_ph is None:
            continue

        # Read positions from slide-level xfrm override (set by the alignment loop above)
        def _xv(p):
            _sp = p._element.find(qn("p:spPr"))
            if _sp is None: return None
            _xf = _sp.find(qn("a:xfrm"))
            if _xf is None: return None
            _o = _xf.find(qn("a:off")); _e = _xf.find(qn("a:ext"))
            if _o is None or _e is None: return None
            return {"x": int(_o.get("x", 0)), "y": int(_o.get("y", 0)),
                    "cx": int(_e.get("cx", 0)), "cy": int(_e.get("cy", 0))}
        hv = _xv(h_ph)
        bv = _xv(b_ph) if b_ph else None
        if hv is None or bv is None:
            continue

        # Equalise body height to heading height for symmetric vertical centering
        new_body_cy = hv["cy"]
        _b_spPr = b_ph._element.find(qn("p:spPr"))
        if _b_spPr is None:
            _b_spPr = etree.SubElement(b_ph._element, qn("p:spPr"))
        for _el in list(_b_spPr.findall(qn("a:xfrm"))):
            _b_spPr.remove(_el)
        _bxfrm = etree.SubElement(_b_spPr, qn("a:xfrm"))
        _boff  = etree.SubElement(_bxfrm, qn("a:off"))
        _boff.set("x", str(bv["x"])); _boff.set("y", str(bv["y"]))
        _bext  = etree.SubElement(_bxfrm, qn("a:ext"))
        _bext.set("cx", str(bv["cx"])); _bext.set("cy", str(new_body_cy))

        # Center text vertically in both placeholders
        for _tph in (h_ph, b_ph):
            _txBody = _tph._element.find(qn("p:txBody"))
            if _txBody is not None:
                _bPr = _txBody.find(qn("a:bodyPr"))
                if _bPr is None:
                    _bPr = etree.SubElement(_txBody, qn("a:bodyPr"))
                _bPr.set("anchor", "ctr")

        gap        = bv["y"] - (hv["y"] + hv["cy"])
        content_cy = hv["cy"] + gap + new_body_cy
        bx  = hv["x"]  - _BOX_PAD
        by  = hv["y"]  - _BOX_PAD
        bcx = hv["cx"] + 2 * _BOX_PAD
        bcy = content_cy + 2 * _BOX_PAD
        box = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
                                     bx, by, bcx, bcy)
        box.fill.background()
        box.line.color.rgb = RGBColor(0x4F, 0x39, 0x82)
        box.line.width = Pt(3)
        avLst = box._element.find(".//" + qn("a:avLst"))
        if avLst is not None:
            for gd in list(avLst.findall(qn("a:gd"))):
                avLst.remove(gd)
            gd_el = etree.SubElement(avLst, qn("a:gd"))
            gd_el.set("name", "adj")
            gd_el.set("fmla", f"val {_BOX_ADJ}")
        # Move box element to the back of the z-stack (behind text placeholders)
        _spTree.remove(box._element)
        _spTree.insert(_insert_idx, box._element)
        _insert_idx += 1

    img_shape = None
    if img_path and Path(img_path).exists():
        img_shape = _add_rp_image(slide, img_path,
                                  content_top_emu=1600000, content_h_emu=4500000)

    _write_notes(slide, _strip_part_heading(row["narration"]))

    # Animation: Title (Wipe) → Image (Fade) → each row heading (Wipe)
    main_list, bldLst, next_id = _build_timing_root(slide._element)
    grp = 0
    title_ph = _ph(slide, PH_CMP_TITLE)
    if title_ph is not None:
        _append_click_effect(main_list, next_id, grp, title_ph.shape_id, anim_type="wipe")
        grp += 1
    if img_shape is not None:
        _append_click_effect(main_list, next_id, grp, img_shape.shape_id, anim_type="fade")
        grp += 1
    for h_idx, _ in _CMP_ROWS:
        h_ph = _ph(slide, h_idx)
        if h_ph is not None:
            _append_click_effect(main_list, next_id, grp, h_ph.shape_id, anim_type="wipe")
            grp += 1

    return slide


def _stamp_layout_xfrm(slide, layout, ph_idx):
    """
    Copy a layout placeholder's xfrm into the slide placeholder's spPr.
    Required for LibreOffice: without an explicit slide-level xfrm, LO may
    fall back to a different layout's position for shared idx values (e.g.
    PH[41] exists in both L_BULLETS and L_ICON3 at different coordinates).
    """
    lx = ly = lcx = lcy = None
    for lph in layout.placeholders:
        if lph.placeholder_format.idx != ph_idx:
            continue
        lxfrm = lph._element.find(".//" + qn("a:xfrm"))
        if lxfrm is None:
            break
        off = lxfrm.find(qn("a:off"))
        ext = lxfrm.find(qn("a:ext"))
        if off is None or ext is None:
            break
        lx  = int(off.get("x", 0));  ly  = int(off.get("y", 0))
        lcx = int(ext.get("cx", 0)); lcy = int(ext.get("cy", 0))
        break
    if lx is None:
        return
    ph = _ph(slide, ph_idx)
    if ph is None:
        return
    spPr = ph._element.find(qn("p:spPr"))
    if spPr is None:
        spPr = etree.SubElement(ph._element, qn("p:spPr"))
    for el in list(spPr.findall(qn("a:xfrm"))):
        spPr.remove(el)
    xfrm = etree.SubElement(spPr, qn("a:xfrm"))
    off  = etree.SubElement(xfrm, qn("a:off"))
    off.set("x", str(lx));  off.set("y", str(ly))
    ext  = etree.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", str(lcx)); ext.set("cy", str(lcy))


def _build_icon3_slide(prs, row, banner, title, bullets, img_path=None):
    """
    60:40 Icon and text layout: slide title at top (PH[34]), then 3 rows each
    with a coloured icon placeholder and a heading+body text pair.
    bullets may be [h1, b1, h2, b2, h3, b3]  OR  ["h1 → b1", "h2 → b2", "h3 → b3"].
    Row text is placed as free textboxes (not placeholders) so LibreOffice cannot
    inherit misaligned indent/margin from PH[41] which is shared with L_BULLETS.
    """
    layout = _find_layout(prs, L_ICON3)
    slide  = prs.slides.add_slide(layout)

    # Title goes into PH[34] — sits at y=0.192" (top of slide) in this layout
    _fill_ph(slide, _ICON3_TITLE, title,
             bold=True, italic=False, color=RGBColor(0x1A, 0x1A, 0x2E))

    # Expand "A → B" bullets into flat [h1, b1, h2, b2, ...] pairs
    if any("→" in b for b in bullets):
        expanded = []
        for b in bullets:
            parts = b.split("→", 1)
            expanded.append(parts[0].strip())
            expanded.append(parts[1].strip() if len(parts) > 1 else "")
        bullets = expanded

    # Row geometry from L_ICON3 layout inspection (in inches)
    _H_GEOM = [  # (left, top, width, height) for each heading
        (Inches(1.583), Inches(1.681), Inches(4.666), Inches(0.50)),
        (Inches(1.563), Inches(3.233), Inches(4.666), Inches(0.50)),
        (Inches(1.563), Inches(4.837), Inches(4.666), Inches(0.50)),
    ]
    _B_GEOM = [  # body text — top offset below heading
        (Inches(1.583), Inches(2.20), Inches(4.666), Inches(0.85)),
        (Inches(1.563), Inches(3.75), Inches(4.666), Inches(0.85)),
        (Inches(1.563), Inches(5.35), Inches(4.666), Inches(0.85)),
    ]
    # Icon circles: same position/size as layout icon placeholders, drawn as OVALs
    _ICON_GEOM = [
        (Inches(0.533), Inches(1.676), Inches(0.984), Inches(0.984)),
        (Inches(0.533), Inches(3.233), Inches(0.984), Inches(0.984)),
        (Inches(0.533), Inches(4.824), Inches(0.984), Inches(0.984)),
    ]
    _ICON_COLORS = [
        RGBColor(0x44, 0x72, 0xC4),   # blue
        RGBColor(0x70, 0x30, 0xA0),   # purple
        RGBColor(0xC5, 0x5A, 0x11),   # orange-red
    ]

    row_heading_shapes = []
    for i in range(3):
        h_text = bullets[i * 2]     if i * 2     < len(bullets) else ""
        b_text = bullets[i * 2 + 1] if i * 2 + 1 < len(bullets) else ""
        # Draw circle icon
        l, t, w, h = _ICON_GEOM[i]
        dot = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, l, t, w, h)
        dot.fill.solid()
        dot.fill.fore_color.rgb = _ICON_COLORS[i]
        dot.line.fill.background()
        h_shape = None
        if h_text:
            h_shape = _textbox(slide, h_text, *_H_GEOM[i],
                               size=16, bold=True, color=RGBColor(0x1A, 0x1A, 0x2E))
        if b_text:
            _textbox(slide, b_text, *_B_GEOM[i],
                     size=14, bold=False, color=RGBColor(0x55, 0x55, 0x66))
        row_heading_shapes.append(h_shape)

    # Keep only the title placeholder; remove icon placeholders (replaced by OVALs
    # above) and all row text placeholders (replaced by textboxes above).
    _keep = {_ICON3_TITLE}
    for p in list(slide.placeholders):
        if p.placeholder_format.idx not in _keep:
            p._element.getparent().remove(p._element)

    img_shape = None
    if img_path and Path(img_path).exists():
        # Icon3 content spans from first icon top (1.676") to last body bottom (6.2")
        _icon3_top = int(_ICON_GEOM[0][1])
        _icon3_h   = int(_B_GEOM[2][1] + _B_GEOM[2][3]) - _icon3_top
        img_shape = _add_rp_image(slide, img_path,
                                  content_top_emu=_icon3_top, content_h_emu=_icon3_h)

    _write_notes(slide, _strip_part_heading(row["narration"]))

    # Animation: Header (Wipe L→R) → Image (Fade) → each row heading (Wipe L→R)
    title_ph = _ph(slide, _ICON3_TITLE)
    main_list, bldLst, next_id = _build_timing_root(slide._element)
    grp = 0
    if title_ph is not None:
        _append_click_effect(main_list, next_id, grp, title_ph.shape_id, anim_type="wipe")
        grp += 1
    if img_shape is not None:
        _append_click_effect(main_list, next_id, grp, img_shape.shape_id, anim_type="fade")
        grp += 1
    for h_sh in row_heading_shapes:
        if h_sh is not None:
            _append_click_effect(main_list, next_id, grp, h_sh.shape_id, anim_type="wipe")
            grp += 1

    return slide


def _build_reflection_slide(prs, row, question_text, img_path=None):
    """
    Reflection Question slide: Emeritus '2C_Image Dark' style — dark full-bleed
    photo as background, semi-transparent overlay on the lower 42%, gold header
    label, question text in white.
    """
    layout = _find_layout(prs, L_WIDE)
    slide  = prs.slides.add_slide(layout)

    W_sl = prs.slide_width
    H_sl = prs.slide_height

    # Default background: Emeritus '2C_Image Dark' dark office photo
    _dark_bg = Path(__file__).parent.parent / "poc3_output" / "images" / "dark_reflection_bg.jpg"
    _bg_path = _dark_bg if _dark_bg.exists() else None

    # Caller-supplied contextual image overrides the default dark background
    if img_path and Path(img_path).exists():
        _bg_path = Path(img_path)

    if _bg_path:
        try:
            from PIL import Image as _PILImg
            with _PILImg.open(_bg_path) as im:
                iw, ih = im.size
            scale    = max(W_sl / iw, H_sl / ih)
            scaled_w = int(iw * scale)
            scaled_h = int(ih * scale)
            x_off    = (W_sl - scaled_w) // 2
            y_off    = (H_sl - scaled_h) // 2
            slide.shapes.add_picture(str(_bg_path), x_off, y_off, scaled_w, scaled_h)
        except Exception:
            slide.shapes.add_picture(str(_bg_path), 0, 0, W_sl, H_sl)
    else:
        _borrow_wide_screen_pics(slide, layout)

    # Compact floating box — 82% of slide width, centered, lower-third of slide.
    # Leaves ~0.9" of image visible on each side so the image remains dominant.
    box_w    = int(W_sl * 0.82)
    box_h    = int(H_sl * 0.34)
    box_left = (W_sl - box_w) // 2
    box_top  = H_sl - box_h - Inches(0.28)  # small bottom margin

    box = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        box_left, box_top, box_w, box_h
    )
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0x0D, 0x1B, 0x2A)
    _set_fill_alpha(box, 88)
    box.line.fill.background()
    # Set a moderate corner radius
    avLst = box._element.find(".//" + qn("a:avLst"))
    if avLst is not None:
        for gd in list(avLst.findall(qn("a:gd"))):
            avLst.remove(gd)
        gd_el = etree.SubElement(avLst, qn("a:gd"))
        gd_el.set("name", "adj")
        gd_el.set("fmla", "val 10000")  # 10% corner radius

    # Gold "REFLECTION QUESTION" label
    _textbox(slide, "REFLECTION QUESTION",
             box_left + Inches(0.35), box_top + Inches(0.15),
             box_w - Inches(0.7), Inches(0.40),
             size=13, bold=True, color=RGBColor(0xFF, 0xC0, 0x00))

    # Prominent question text in white
    _textbox(slide, question_text,
             box_left + Inches(0.35), box_top + Inches(0.60),
             box_w - Inches(0.7), box_h - Inches(0.70),
             size=28, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

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
        if Path(out).exists():
            # Never regenerate images that already exist — reuse cached version
            paths[frame] = out
            print(f"  [frame {frame}] Using existing image")
            continue
        print(f"  [frame {frame}] Generating image …", end=" ", flush=True)
        p = generate_slide_image(prompt, out, gemini_api_key)
        paths[frame] = p or None
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

    # Reflection question — full-bleed image + dark overlay + prominent text
    if "reflection" in cue_lower or title.lower().startswith("reflection"):
        question = " ".join(bullets) if bullets else title
        return _build_reflection_slide(prs, row, question, img_path)

    # 3-column boxes comparison — uses "60:40 Text boxes with headings" template
    if ("comparison" in cue_lower or "3-column" in cue_lower
            or ("column" in cue_lower and "boxes" in cue_lower)):
        return _build_comparison_slide(prs, row, banner, title, bullets, img_path)

    # Use the exact OST bullets as specified in the storyboard document
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
    template_path = template_path or str(script_dir.parent / "POC3.Template_(1).pptx")

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
    _clean_layout_prompts(prs)   # wipe ghost text from layout PH[41] and PH[34]
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
