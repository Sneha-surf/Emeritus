#!/usr/bin/env python3
"""
POC 3 — Word Document–Driven Storyboard Workflow

Converts a descriptive Word document (.docx) into a storyboarded PPTX
presentation using Claude AI.

Pipeline:
  Stage 1  Parse Word doc → doc_structure.json       (section-wise content)
  Stage 2  Claude storyboards each section → storyboard.json  (scene flow)
  Stage 3  Build PPTX skeleton → storyboard.pptx     (ready for review/edit)

Usage:
  python poc3_run.py <docx_path> [output_dir]

Arguments:
  docx_path    Path to the source Word document (.docx)
  output_dir   Where to write all outputs (default: ./poc3_output)

Example:
  python poc3_run.py course_material.docx ./poc3_output

Requirements:
  pip install python-docx anthropic python-pptx
  ANTHROPIC_API_KEY must be set (in .env or environment)
"""

import sys
import json
import time
import os
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from poc3_stage1_docparse   import run_stage1
from poc3_stage2_storyboard import run_stage2
from poc3_stage3_pptx       import run_stage3
from poc3_stage4_video      import run_stage4


def _load_env():
    env = Path(__file__).parent.parent / ".env"
    if env.exists():
        with open(env) as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, _, v = line.partition("=")
                    os.environ.setdefault(k.strip(), v.strip())


def run_poc3(docx_path: str, output_dir: str) -> dict:
    """
    Orchestrate all three POC-3 stages. Returns a combined manifest.
    """
    docx_path  = str(Path(docx_path).resolve())
    output_dir = str(Path(output_dir).resolve())

    _load_env()

    if not Path(docx_path).exists():
        print(f"[ERROR] Document not found: {docx_path}")
        sys.exit(1)

    print("=" * 60)
    print("  POC 3 — Word Document–Driven Storyboard Workflow")
    print("=" * 60)
    print(f"  Input  : {docx_path}")
    print(f"  Output : {output_dir}")
    print()

    t0 = time.time()

    # ── Stage 1 ───────────────────────────────────────────────────────────────
    m1 = run_stage1(docx_path, output_dir)
    print()

    # ── Stage 2 ───────────────────────────────────────────────────────────────
    m2 = run_stage2(m1["doc_structure"], output_dir)
    print()

    # ── Stage 3 ───────────────────────────────────────────────────────────────
    gemini_key = os.environ.get("GEMINI_API_KEY")
    if not gemini_key:
        print("[WARN] GEMINI_API_KEY not set — slides will be built without images")
    m3 = run_stage3(m2["storyboard"], output_dir, gemini_api_key=gemini_key)
    print()

    # ── Stage 4 ───────────────────────────────────────────────────────────────
    openai_key = os.environ.get("AZURE_OPENAI_API_KEY")
    if not openai_key:
        print("[INFO] AZURE_OPENAI_API_KEY not set — TTS will use edge-tts/gTTS (free fallback)")
    m4 = run_stage4(m2["storyboard"], m3["pptx"], output_dir,
                    openai_api_key=openai_key)

    elapsed = time.time() - t0

    # ── Combined manifest ─────────────────────────────────────────────────────
    manifest = {
        "input_docx":      docx_path,
        "output_dir":      output_dir,
        "doc_structure":   m1["doc_structure"],
        "storyboard_json": m2["storyboard"],
        "storyboard_pptx":  m3["pptx"],
        "storyboard_video": m4.get("video"),
        "doc_title":        m1["title"],
        "section_count":    m1["section_count"],
        "total_slides":     m3["total_slides"],
        "elapsed_seconds":  round(elapsed, 1),
    }

    manifest_path = Path(output_dir) / "poc3_manifest.json"
    with open(manifest_path, "w") as f:
        json.dump(manifest, f, indent=2)

    print()
    print("=" * 60)
    print("  POC 3 Complete")
    print("=" * 60)
    print(f'  Document  : "{manifest["doc_title"]}"')
    print(f"  Sections  : {manifest['section_count']}")
    print(f"  Slides    : {manifest['total_slides']}")
    print(f"  Time      : {elapsed:.1f}s")
    print()
    print(f"  Outputs saved to: {output_dir}/")
    print(f"    doc_structure.json  — section-wise document content")
    print(f"    storyboard.json     — AI-generated scene storyboard")
    print(f"    storyboard.pptx     — PPTX ready for review")
    if manifest.get("storyboard_video"):
        print(f"    storyboard_video.mp4 — narrated video")
    print("=" * 60)

    return manifest


# ── CLI ────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    docx   = sys.argv[1]
    outdir = sys.argv[2] if len(sys.argv) > 2 else "./poc3_output"

    run_poc3(docx, outdir)
