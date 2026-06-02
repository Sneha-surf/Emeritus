#!/usr/bin/env python3
"""
POC 3 — Stage 1: Word Document Ingestion

Parses a .docx file and extracts its content section-wise based on heading
hierarchy. Each Heading 1 creates a top-level section; Heading 2 creates
subsections within it. Body paragraphs, bullet lists, and tables are captured
under whichever heading they follow.

Outputs  →  <output_dir>/doc_structure.json

Usage:
  python poc3_stage1_docparse.py <docx_path> <output_dir>

Example:
  python poc3_stage1_docparse.py course_material.docx ./poc3_output
"""

import sys
import json
import re
from pathlib import Path


# ── Implicit section detection (for script-style documents without heading styles) ──

_PART_RE = re.compile(
    r'^(part\s+(?:one|two|three|four|five|six|seven|eight|nine|ten|\d+))\s*[,\.]\s*',
    re.IGNORECASE,
)

_SKIP_RE = re.compile(
    r'^(?:end\s+of\s+video|sb:\s*$|video\s+(?:one|two|three|four|five|six|seven|eight|nine|ten|\d+)\s*[,:])',
    re.IGNORECASE,
)

_VIDEO_TITLE_RE = re.compile(
    r'^video\s+(?:one|two|three|four|five|six|seven|eight|nine|ten|\d+)\s*[,:]\s*(.+)',
    re.IGNORECASE,
)


def _detect_implicit_heading(text: str):
    """
    Detect two patterns used in narration scripts:
      1. "Part N, heading title. Body content..."
      2. "Short heading (≤7 words). Longer body content..."
    Returns (heading_str, remainder_text) or (None, text).
    """
    # Pattern 1 — explicit "Part N," marker
    m = _PART_RE.match(text)
    if m:
        part_label = m.group(1).title()          # e.g. "Part One"
        after = text[m.end():]
        dot = after.find(".")
        if dot >= 0:
            title_part = after[:dot].strip()
            remainder  = after[dot + 1:].strip()
        else:
            title_part = after.strip()
            remainder  = ""
        heading = f"{part_label}: {title_part}" if title_part else part_label
        return heading, remainder

    # Pattern 2 — "Short sentence (≤7 words). Longer content (≥15 words)."
    dot = text.find(".")
    if dot >= 0:
        first = text[:dot].strip()
        rest  = text[dot + 1:].strip()
        if 2 <= len(first.split()) <= 7 and len(rest.split()) >= 15:
            return first, rest

    return None, text


def _is_duplicate(text: str, existing: list) -> bool:
    """True if text's first 60 chars match any existing paragraph (catches duplicate UK/US variants)."""
    key = text[:60].lower()
    return any(p[:60].lower() == key for p in existing)


# ── Helpers ────────────────────────────────────────────────────────────────────

def _heading_level(para) -> int:
    """Return 1-6 if the paragraph is a heading, else 0."""
    style_name = para.style.name or ""
    if style_name.startswith("Heading"):
        try:
            return int(style_name.split()[-1])
        except (ValueError, IndexError):
            pass
    return 0


def _para_text(para) -> str:
    return para.text.strip()


def _is_list_item(para) -> bool:
    """True if the paragraph uses a list style (numbered or bullet)."""
    style_name = para.style.name or ""
    return "List" in style_name


def _table_to_text(table) -> str:
    """Flatten a table into a readable multi-line string."""
    rows = []
    for row in table.rows:
        cells = [cell.text.strip() for cell in row.cells]
        rows.append(" | ".join(cells))
    return "\n".join(rows)


# ── Section Builder ────────────────────────────────────────────────────────────

def _new_section(heading: str, level: int) -> dict:
    return {
        "heading": heading,
        "level": level,
        "paragraphs": [],
        "lists": [],
        "tables": [],
        "subsections": [],
    }


def _append_content(section: dict, para, tables_iter) -> bool:
    """
    Add content from a paragraph (or a table placeholder) to the section.
    Returns True if content was added.
    """
    level = _heading_level(para)
    if level > 0:
        return False  # heading — caller handles this

    text = _para_text(para)
    if not text:
        return True  # blank line — skip silently

    if _is_list_item(para):
        section["lists"].append(text)
    else:
        section["paragraphs"].append(text)
    return True


# ── Main Parser ────────────────────────────────────────────────────────────────

def parse_document(docx_path: str) -> dict:
    """
    Parse the Word document and return a structured dict:

      {
        "title": str,
        "sections": [
          {
            "heading": str,
            "level": int,            # always 1 for top-level sections
            "paragraphs": [str, ...],
            "lists": [str, ...],
            "tables": [str, ...],
            "subsections": [
              { same shape, level=2, subsections=[] }
            ]
          },
          ...
        ]
      }

    Content that appears before the first heading is stored in a synthetic
    "Introduction / Preamble" section so nothing is lost.
    """
    try:
        from docx import Document
    except ImportError:
        print("[ERROR] python-docx not installed. Run: pip install python-docx")
        sys.exit(1)

    doc = Document(docx_path)

    # Document title: prefer a "Video N: Title" line; fall back to core properties / filename
    title = ""
    try:
        title = doc.core_properties.title or ""
    except Exception:
        pass

    extracted_title = ""   # from "Video N: ..." lines in body
    sections = []
    current_h1 = None
    current_h2 = None

    # We iterate over body elements to handle both paragraphs and tables in order
    from docx.oxml.ns import qn
    body = doc.element.body

    for child in body:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "tbl":
            # Table element
            from docx.table import Table as DocxTable
            table_obj = DocxTable(child, doc)
            table_text = _table_to_text(table_obj)
            target = current_h2 or current_h1
            if target is not None:
                target["tables"].append(table_text)
            continue

        if tag != "p":
            continue

        # Paragraph element
        from docx.text.paragraph import Paragraph as DocxPara
        para = DocxPara(child, doc)

        level = _heading_level(para)
        text  = _para_text(para)

        if level == 1:
            if not title:
                title = text  # use first H1 as document title if none found
            current_h1 = _new_section(text, 1)
            current_h2 = None
            sections.append(current_h1)

        elif level == 2:
            if current_h1 is None:
                # H2 with no parent H1 — create a synthetic H1 wrapper
                current_h1 = _new_section("Preamble", 1)
                sections.append(current_h1)
            current_h2 = _new_section(text, 2)
            current_h1["subsections"].append(current_h2)

        elif level >= 3:
            # Treat H3+ as bold paragraphs within the current section
            target = current_h2 or current_h1
            if target is None:
                current_h1 = _new_section("Preamble", 1)
                sections.append(current_h1)
                target = current_h1
            if text:
                target["paragraphs"].append(f"[{text}]")  # mark as sub-heading

        else:
            # Body content
            if not text:
                continue

            # Skip metadata / stage-direction lines; extract title from "Video N: Title"
            if _SKIP_RE.match(text):
                if not extracted_title:
                    m = _VIDEO_TITLE_RE.match(text)
                    if m:
                        extracted_title = m.group(1).strip().rstrip(".")
                continue

            # Check for implicit section heading ("Part N," or short heading pattern)
            heading, remainder = _detect_implicit_heading(text)
            if heading:
                current_h1 = _new_section(heading, 1)
                current_h2 = None
                sections.append(current_h1)
                if remainder:
                    current_h1["paragraphs"].append(remainder)
                continue

            target = current_h2 or current_h1
            if target is None:
                # Content before any heading
                current_h1 = _new_section("Introduction", 1)
                sections.append(current_h1)
                target = current_h1

            if _is_list_item(para):
                target["lists"].append(text)
            elif not _is_duplicate(text, target["paragraphs"]):
                target["paragraphs"].append(text)

    # Prefer extracted "Video N: Title" over a filename-like core-properties title
    if extracted_title and (not title or "_" in title or " - V" in title or " - v" in title):
        title = extracted_title
    if not title:
        title = Path(docx_path).stem

    return {"title": title, "sections": sections}


# ── Stage 1 Runner ─────────────────────────────────────────────────────────────

def run_stage1(docx_path: str, output_dir: str) -> dict:
    """
    Run Stage 1 ingestion. Returns a manifest with paths to outputs.

    manifest keys:
      doc_structure  — path to doc_structure.json
      title          — document title
      section_count  — number of top-level sections
    """
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    print("[Stage 1/3] Parsing Word document …")
    doc = parse_document(docx_path)

    # Stats
    total_subs = sum(len(s["subsections"]) for s in doc["sections"])
    total_paras = sum(
        len(s["paragraphs"]) + sum(len(sub["paragraphs"]) for sub in s["subsections"])
        for s in doc["sections"]
    )
    total_lists = sum(
        len(s["lists"]) + sum(len(sub["lists"]) for sub in s["subsections"])
        for s in doc["sections"]
    )
    print(f'            Title   : "{doc["title"]}"')
    print(f"            Sections: {len(doc['sections'])} top-level, {total_subs} subsections")
    print(f"            Content : {total_paras} paragraphs, {total_lists} list items")

    structure_path = out / "doc_structure.json"
    with open(structure_path, "w", encoding="utf-8") as f:
        json.dump(doc, f, indent=2, ensure_ascii=False)
    print(f"            Doc structure → {structure_path}")

    manifest = {
        "doc_structure": str(structure_path),
        "title": doc["title"],
        "section_count": len(doc["sections"]),
    }

    manifest_path = out / "stage1_manifest.json"
    with open(manifest_path, "w") as f:
        json.dump(manifest, f, indent=2)
    print(f"\n[Stage 1 Complete] Manifest → {manifest_path}")

    return manifest


# ── CLI ────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    result = run_stage1(sys.argv[1], sys.argv[2])
    print()
    for k, v in result.items():
        print(f"  {k}: {v}")
