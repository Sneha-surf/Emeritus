#!/usr/bin/env python3
"""Minimal PPTX→PNG helper called by server.py for browser-side previews.
Usage: python3 _pptx_to_png.py <pptx_path> <output_dir>
Prints a JSON object {"slides": [...abs paths...], "person_boxes": [...]} to stdout;
logs go to stderr.
"""
import sys, json
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from pipeline import pptx_to_images

if len(sys.argv) < 3:
    print("Usage: _pptx_to_png.py <pptx_path> <output_dir>", file=sys.stderr)
    sys.exit(1)

pptx_path = sys.argv[1]
out_dir   = sys.argv[2]

# Redirect stdout → stderr so pipeline log lines don't corrupt the JSON output
_real_stdout = sys.stdout
sys.stdout = sys.stderr

try:
    images = pptx_to_images(pptx_path, out_dir)
finally:
    sys.stdout = _real_stdout

# Person detection is disabled for PPTX slides: corridor/campus photo backgrounds
# produce false positives (HOG/face detection fires on decorative photos), and the
# REPLACE positioning mode is only meaningful in video-background (VSB) mode.
person_boxes = [None] * len(images)

print(json.dumps({'slides': images, 'person_boxes': person_boxes}))
