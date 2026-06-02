import cv2
import numpy as np
import subprocess
import os
import fitz
from rembg import remove, new_session
from PIL import Image

BASE_DIR = r"C:\Users\Linchana S Y\POC2_Project"
PPTX_PATH = os.path.join(BASE_DIR, "CS_W8_S4_VSB_V1_Reviewed4321.pptx")
AVATAR_VIDEO = os.path.join(BASE_DIR, "avatar_videos", "avatar_hello.mp4")
SOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"
FFMPEG_PATH = "ffmpeg"
OUT_DIR = os.path.join(BASE_DIR, "slides_output")
TEMP_VIDEO = os.path.join(BASE_DIR, "temp_video_no_audio.mp4")
FINAL_VIDEO = os.path.join(BASE_DIR, "final_output", "poc2_final.mp4")

WIDTH = 1920
HEIGHT = 1080
OUTPUT_FPS = 30
TOTAL_DURATION = 120  # 4 slides x 30s

SLIDE_TIMELINE = [
    (0,   30,  0, "hide"),    # Slide 1 — no avatar
    (30,  60,  1, "hide"),    # Slide 2 — no avatar
    (60,  90,  2, "center"),  # Slide 3 — avatar center
    (90,  120, 3, "right"),   # Slide 4 — avatar right
]

AVATAR_HEIGHT = 500
RIGHT_MARGIN = 40

session = new_session("u2net_human_seg")
STABLE_BOX = None


def export_slides_to_images():
    os.makedirs(OUT_DIR, exist_ok=True)
    pdf_path = os.path.join(BASE_DIR, "slides_temp.pdf")
    subprocess.run([
        SOFFICE_PATH, "--headless", "--convert-to", "pdf",
        "--outdir", BASE_DIR, PPTX_PATH
    ], check=True)

    # Rename the generated PDF
    pptx_name = os.path.splitext(os.path.basename(PPTX_PATH))[0]
    generated_pdf = os.path.join(BASE_DIR, pptx_name + ".pdf")
    if os.path.exists(generated_pdf) and generated_pdf != pdf_path:
        os.rename(generated_pdf, pdf_path)

    doc = fitz.open(pdf_path)
    slide_images = []
    for i in range(min(4, len(doc))):  # Only first 4 slides
        page = doc[i]
        mat = fitz.Matrix(1920 / page.rect.width, 1080 / page.rect.height)
        pix = page.get_pixmap(matrix=mat)
        img_path = os.path.join(OUT_DIR, f"slide_{i:02d}.png")
        pix.save(img_path)
        slide_images.append(img_path)
    doc.close()
    print(f"Exported {len(slide_images)} slide images.")
    return slide_images


def ai_cut_avatar(frame_bgr):
    """Remove white background from avatar frame using rembg."""
    global STABLE_BOX
    frame_rgb = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2RGB)
    pil_img = Image.fromarray(frame_rgb)
    result = remove(pil_img, session=session)
    result_np = np.array(result)  # RGBA
    return result_np  # returns RGBA


def prepare_avatar(rgba_frame, target_height):
    """Scale avatar to target height, keeping aspect ratio."""
    h, w = rgba_frame.shape[:2]
    scale = target_height / h
    new_w = int(w * scale)
    new_h = target_height
    resized = cv2.resize(rgba_frame, (new_w, new_h), interpolation=cv2.INTER_AREA)
    return resized


def place_avatar(slide_bgr, avatar_rgba, position="right"):
    """Composite avatar RGBA onto slide BGR at given position."""
    slide_h, slide_w = slide_bgr.shape[:2]
    av_h, av_w = avatar_rgba.shape[:2]

    if position == "right":
        x = slide_w - av_w - RIGHT_MARGIN
        y = slide_h - av_h
    elif position == "center":
        x = (slide_w - av_w) // 2
        y = slide_h - av_h
    else:
        return slide_bgr  # "hide" — no avatar

    x = max(0, min(x, slide_w - av_w))
    y = max(0, min(y, slide_h - av_h))

    result = slide_bgr.copy()
    roi = result[y:y+av_h, x:x+av_w]

    avatar_bgr = avatar_rgba[:, :, :3]
    alpha = avatar_rgba[:, :, 3:4].astype(np.float32) / 255.0

    blended = (avatar_bgr.astype(np.float32) * alpha +
               roi.astype(np.float32) * (1 - alpha))
    result[y:y+av_h, x:x+av_w] = blended.astype(np.uint8)
    return result


def build_video(slide_images):
    os.makedirs(os.path.dirname(TEMP_VIDEO), exist_ok=True)
    os.makedirs(os.path.dirname(FINAL_VIDEO), exist_ok=True)

    cap = cv2.VideoCapture(AVATAR_VIDEO)
    avatar_fps = cap.get(cv2.CAP_PROP_FPS)
    avatar_total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    print(f"Avatar video: {avatar_fps:.1f} fps, {avatar_total_frames} frames")

    fourcc = cv2.VideoWriter_fourcc(*'mp4v')
    out = cv2.VideoWriter(TEMP_VIDEO, fourcc, OUTPUT_FPS, (WIDTH, HEIGHT))

    # Preload slide images
    slides = []
    for path in slide_images:
        img = cv2.imread(path)
        img = cv2.resize(img, (WIDTH, HEIGHT))
        slides.append(img)

    total_frames = TOTAL_DURATION * OUTPUT_FPS
    avatar_frame_idx = 0

    for frame_num in range(total_frames):
        current_time = frame_num / OUTPUT_FPS

        # Find current slide and position
        slide_idx = 0
        position = "hide"
        for start, end, s_idx, pos in SLIDE_TIMELINE:
            if start <= current_time < end:
                slide_idx = s_idx
                position = pos
                break

        slide_bg = slides[min(slide_idx, len(slides)-1)].copy()

        if position != "hide":
            # Read avatar frame (loop if needed)
            ret, av_frame = cap.read()
            if not ret:
                cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
                ret, av_frame = cap.read()
            if ret:
                avatar_rgba = ai_cut_avatar(av_frame)
                avatar_rgba = prepare_avatar(avatar_rgba, AVATAR_HEIGHT)
                slide_bg = place_avatar(slide_bg, avatar_rgba, position)
            avatar_frame_idx += 1
        else:
            # Still advance avatar frame counter to stay in sync (optional)
            pass

        out.write(slide_bg)

        if frame_num % 30 == 0:
            print(f"  Frame {frame_num}/{total_frames} ({current_time:.1f}s) — slide {slide_idx+1}, pos={position}")

    cap.release()
    out.release()
    print(f"Silent video written: {TEMP_VIDEO}")

    # Mux audio from avatar_hello.mp4
    subprocess.run([
        FFMPEG_PATH, "-y",
        "-i", TEMP_VIDEO,
        "-i", AVATAR_VIDEO,
        "-c:v", "copy",
        "-c:a", "aac",
        "-map", "0:v:0",
        "-map", "1:a:0",
        "-shortest",
        FINAL_VIDEO
    ], check=True)
    print(f"\n✅ Final video: {FINAL_VIDEO}")


def main():
    print("Step 1: Exporting slides...")
    slide_images = export_slides_to_images()

    print("\nStep 2: Building composite video...")
    build_video(slide_images)

    print("\nDone! Check:", FINAL_VIDEO)


if __name__ == "__main__":
    main()