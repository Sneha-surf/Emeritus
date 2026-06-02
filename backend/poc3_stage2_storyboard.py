#!/usr/bin/env python3
"""
POC 3 — Stage 2: AI-Driven Storyboarding

Uses GPT to read the section-wise doc_structure.json produced by Stage 1
and generate a full storyboard: one or more slides per section, each with a
scene type, title, bullet points, visual suggestion, and speaker notes.

Outputs  →  <output_dir>/storyboard.json

Usage:
  python poc3_stage2_storyboard.py <doc_structure_json> <output_dir>

Example:
  python poc3_stage2_storyboard.py ./poc3_output/doc_structure.json ./poc3_output

Requires:
  AZURE_OPENAI_API_KEY      — your Azure OpenAI key
  AZURE_OPENAI_ENDPOINT     — e.g. https://<resource>.openai.azure.com
  AZURE_OPENAI_DEPLOYMENT   — your deployment name (e.g. gpt-4o)
  AZURE_OPENAI_API_VERSION  — API version (default: 2024-02-01)
  All can be set in a .env file at the project root.
"""

import sys
import json
import os
from pathlib import Path


# ── Scene types the AI can assign ─────────────────────────────────────────────
# "title" is intentionally excluded — section divider slides are auto-generated
# by Stage 3, so AI title slides would be duplicates.
SCENE_TYPES = ("bullets", "visual", "quote", "table", "summary")

SYSTEM_PROMPT = """\
You are an expert instructional designer who converts academic course content
into storyboards for video-based learning modules.

Given a document section (heading + body text), your job is to plan the slides
that will cover that section. Return ONLY a valid JSON array — no markdown fences,
no extra text.

⚠ CONTENT FIDELITY — CRITICAL:
Every bullet point MUST be a direct quote or minimal rewrite of actual text from
the provided section. Do NOT invent facts, add examples, or include anything not
explicitly stated in the source material. If the section has only 2 clear points,
make 2 bullets — do not pad. Bullets should preserve the document's exact wording
as closely as possible.

Each element is a slide object with these exact keys:
  scene_type        : one of "bullets" | "visual" | "quote" | "table" | "summary"
                      — do NOT use "title"; section headers are auto-generated separately
  title             : the section heading or a short version of it (≤ 10 words);
                      use the document's own headings, not invented ones
  bullets           : if the section contains a storyboard table with an "OST" column,
                      split that column's text on newlines and use each non-empty,
                      non-header line as a separate bullet string (verbatim). Otherwise
                      use 2–6 strings taken VERBATIM from the document text;
                      empty array for visual/quote slides
  visual_description: 1–2 sentences describing the best visual/diagram for this slide
  image_prompt      : detailed AI image-generation prompt — subject, style (e.g.
                      "flat design illustration", "3D isometric diagram"), mood, and
                      colour palette in 2–3 sentences; no readable text in the image
  speaker_notes     : if the section contains a table with a "Narration Chunk" or "VO"
                      column, copy that column's text VERBATIM as the speaker notes for
                      the matching slide — do NOT rephrase or expand it. Otherwise write
                      a 60–120 word script that stays strictly faithful to the document
                      text; do not add examples, invented facts, or filler sentences.

Rules:
- NEVER output a "title" scene_type — section divider slides are added automatically.
- "bullets" slide: 2–6 points taken verbatim/near-verbatim from the source text.
- "visual" slide: use when the document describes a diagram, model, process, or comparison.
- "quote" slide: use for direct quotes or reflection questions from the document;
                 put the quote/question as the sole bullet.
- "summary" slide: closes a section; bullets are the key facts listed in the document.
- Aim for 2–4 slides per section. Do NOT pad with invented content.
- Slide numbering is NOT included — the runner assigns global slide numbers.
"""


# ── Load .env if present ────────────────────────────────────────────────────────

def _load_env():
    env_path = Path(__file__).parent.parent / ".env"
    if env_path.exists():
        with open(env_path) as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, _, v = line.partition("=")
                    os.environ.setdefault(k.strip(), v.strip())


# ── Claude call ────────────────────────────────────────────────────────────────

def _call_gpt(client, deployment: str, section_text: str, doc_title: str) -> list:
    """
    Send one section to Azure OpenAI and return a list of slide dicts.
    """
    response = client.chat.completions.create(
        model=deployment,
        max_tokens=4096,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": (
                    f'Document title: "{doc_title}"\n\n'
                    f"Section content:\n{section_text}"
                ),
            },
        ],
    )

    raw = response.choices[0].message.content.strip()

    # Strip accidental markdown fences
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
        raw = raw.strip()

    try:
        slides = json.loads(raw)
    except json.JSONDecodeError as e:
        print(f"          [WARN] Claude returned invalid JSON: {e}")
        print(f"          Raw output (first 300 chars): {raw[:300]}")
        slides = []

    # Validate and normalise; drop any "title" slides (section dividers handle those)
    valid = []
    for slide in slides:
        if not isinstance(slide, dict):
            continue
        scene = slide.get("scene_type", "bullets")
        if scene == "title":
            continue
        valid.append({
            "scene_type":         scene,
            "title":              slide.get("title", ""),
            "bullets":            slide.get("bullets", []),
            "visual_description": slide.get("visual_description", ""),
            "image_prompt":       slide.get("image_prompt", ""),
            "speaker_notes":      slide.get("speaker_notes", ""),
        })
    return valid


# ── Section → text ─────────────────────────────────────────────────────────────

def _section_to_text(section: dict) -> str:
    """Flatten a section dict into a single prompt-friendly string."""
    parts = [f"# {section['heading']}"]

    if section["paragraphs"]:
        parts.append("\n".join(section["paragraphs"]))

    if section["lists"]:
        parts.append("List items:\n" + "\n".join(f"- {li}" for li in section["lists"]))

    if section["tables"]:
        for i, tbl in enumerate(section["tables"], 1):
            parts.append(f"Table {i}:\n{tbl}")

    for sub in section.get("subsections", []):
        parts.append(f"\n## {sub['heading']}")
        if sub["paragraphs"]:
            parts.append("\n".join(sub["paragraphs"]))
        if sub["lists"]:
            parts.append("\n".join(f"- {li}" for li in sub["lists"]))
        if sub["tables"]:
            for tbl in sub["tables"]:
                parts.append(tbl)

    return "\n\n".join(parts)


# ── Stage 2 Runner ─────────────────────────────────────────────────────────────

def run_stage2(doc_structure_path: str, output_dir: str) -> dict:
    """
    Run Stage 2 storyboarding. Returns a manifest with paths to outputs.

    manifest keys:
      storyboard       — path to storyboard.json
      total_slides     — total number of slides generated
      total_sections   — number of document sections processed
    """
    _load_env()

    try:
        from openai import AzureOpenAI
    except ImportError:
        print("[ERROR] openai SDK not installed. Run: pip install openai")
        sys.exit(1)

    api_key    = os.environ.get("AZURE_OPENAI_API_KEY", "")
    endpoint   = os.environ.get("AZURE_OPENAI_ENDPOINT", "")
    deployment = os.environ.get("AZURE_OPENAI_DEPLOYMENT", "gpt-4o")
    api_version = os.environ.get("AZURE_OPENAI_API_VERSION", "2024-02-01")

    missing = [k for k, v in [
        ("AZURE_OPENAI_API_KEY", api_key),
        ("AZURE_OPENAI_ENDPOINT", endpoint),
    ] if not v]
    if missing:
        print(f"[ERROR] Missing env vars: {', '.join(missing)}. Add them to your .env file.")
        sys.exit(1)

    with open(doc_structure_path, encoding="utf-8") as f:
        doc = json.load(f)

    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    client = AzureOpenAI(
        api_key=api_key,
        azure_endpoint=endpoint,
        api_version=api_version,
    )
    doc_title = doc.get("title", "Untitled")
    sections  = doc.get("sections", [])

    print(f"[Stage 2/3] Generating storyboard for \"{doc_title}\" …")
    print(f"            {len(sections)} sections to process via Claude")

    storyboard_sections = []
    slide_counter = 1

    for i, section in enumerate(sections, 1):
        heading = section.get("heading", f"Section {i}")
        print(f"            [{i}/{len(sections)}] \"{heading}\" …", end=" ", flush=True)

        section_text = _section_to_text(section)
        slides = _call_gpt(client, deployment, section_text, doc_title)

        # Assign sequential slide numbers
        for slide in slides:
            slide["slide_number"] = slide_counter
            slide_counter += 1

        storyboard_sections.append({
            "section_id":      i,
            "section_heading": heading,
            "slides":          slides,
        })
        print(f"{len(slides)} slides")

    total_slides = slide_counter - 1
    print(f"            Total: {total_slides} slides across {len(sections)} sections")

    storyboard = {
        "title":          doc_title,
        "total_slides":   total_slides,
        "sections":       storyboard_sections,
    }

    storyboard_path = out / "storyboard.json"
    with open(storyboard_path, "w", encoding="utf-8") as f:
        json.dump(storyboard, f, indent=2, ensure_ascii=False)
    print(f"            Storyboard → {storyboard_path}")

    manifest = {
        "storyboard":     str(storyboard_path),
        "total_slides":   total_slides,
        "total_sections": len(sections),
    }

    manifest_path = out / "stage2_manifest.json"
    with open(manifest_path, "w") as f:
        json.dump(manifest, f, indent=2)
    print(f"\n[Stage 2 Complete] Manifest → {manifest_path}")

    return manifest


# ── CLI ────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    result = run_stage2(sys.argv[1], sys.argv[2])
    print()
    for k, v in result.items():
        print(f"  {k}: {v}")
