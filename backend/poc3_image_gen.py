#!/usr/bin/env python3
"""
POC 3 — Image Generation via Gemini Imagen

Generates 16:9 slide images using Gemini Imagen 3 (Google AI Studio).
"""

import sys
from pathlib import Path


def generate_slide_image(prompt: str, output_path: str, api_key: str) -> str:
    """
    Generate a 16:9 image from prompt using Gemini Imagen 3.
    Returns output_path on success, None on failure.
    """
    try:
        from google import genai
        from google.genai import types
    except ImportError:
        print("[WARN] google-genai not installed. Run: pip install google-genai")
        return None

    try:
        client = genai.Client(api_key=api_key)
        response = client.models.generate_images(
            model="imagen-4.0-generate-001",
            prompt=prompt,
            config=types.GenerateImagesConfig(
                number_of_images=1,
                aspect_ratio="16:9",
            ),
        )
        img_bytes = response.generated_images[0].image.image_bytes
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, "wb") as f:
            f.write(img_bytes)
        return output_path

    except Exception as e:
        print(f"\n          [WARN] Imagen failed: {e}")
        return None


def darken_image(src_path: str, dst_path: str, factor: float = 0.55) -> str:
    """
    Create a darkened copy of an image for use as a slide background.
    factor: 0.0 = no change, 1.0 = pure black.
    Returns dst_path.
    """
    try:
        from PIL import Image
    except ImportError:
        print("[WARN] Pillow not installed — background image won't be darkened.")
        return src_path

    img = Image.open(src_path).convert("RGB")
    dark = Image.new("RGB", img.size, (0, 0, 0))
    blended = Image.blend(img, dark, factor)
    blended.save(dst_path)
    return dst_path
