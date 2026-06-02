import cv2
import numpy as np
import subprocess
import os
import fitz
from rembg import remove, new_session
from PIL import Image

BASE_DIR = r"C:\ppoocc11"

PPTX_PATH = os.path.join(BASE_DIR, "CS_W8_S4_VSB_V1_Reviewed4321.pptx")
FACULTY_VIDEO = os.path.join(BASE_DIR, "Green-screen Faculty Recording.mp4")

SOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"
FFMPEG_PATH = "ffmpeg"

OUT_DIR = os.path.join(BASE_DIR, "slides_output")
TEMP_VIDEO = os.path.join(BASE_DIR, "temp_video_no_audio.mp4")
FINAL_VIDEO = os.path.join(BASE_DIR, "final_synced_output_final.mp4")

WIDTH = 1920
HEIGHT = 1080
OUTPUT_FPS = 30

AUDIO_START_TIME = 7
TOTAL_DURATION = 120

SLIDE_TIMELINE = [
    (0, 2, 0, "hide"),
    (2, 7, 1, "hide"),
    (7, 31, 2, "center"),
    (31, 68, 3, "right"),
    (68, 95, 4, "center"),
    (95, 120, 5, "right")
]

FACULTY_HEIGHT = 1010
RIGHT_MARGIN = 40

session = new_session("u2net_human_seg")
STABLE_BOX = None


def run_cmd(cmd):
    subprocess.run(cmd, check=True)


def convert_ppt_to_images():
    os.makedirs(OUT_DIR, exist_ok=True)

    run_cmd([
        SOFFICE_PATH,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", OUT_DIR,
        PPTX_PATH
    ])

    pdf_name = os.path.splitext(os.path.basename(PPTX_PATH))[0] + ".pdf"
    pdf_path = os.path.join(OUT_DIR, pdf_name)

    doc = fitz.open(pdf_path)
    slide_paths = []

    for i in range(6):
        page = doc[i]
        pix = page.get_pixmap(matrix=fitz.Matrix(2.8, 2.8))
        img_path = os.path.join(OUT_DIR, f"slide_{i + 1}.png")
        pix.save(img_path)
        slide_paths.append(img_path)

    doc.close()
    return slide_paths


def get_slide_info(t):
    for start, end, slide_index, position in SLIDE_TIMELINE:
        if start <= t < end:
            return slide_index, position
    return 5, "right"


def ease_out(p):
    p = max(0, min(1, p))
    return 1 - (1 - p) * (1 - p)


def animate_slide_2(slide, t):
    local_t = t - 2
    original = slide.copy()
    animated = slide.copy()

    x1, y1, x2, y2 = 145, 165, 770, 290
    word_img = original[y1:y2, x1:x2].copy()

    bg_patch = original[y2 + 35:y2 + 95, x1:x2]
    if bg_patch.size == 0:
        bg_color = (35, 145, 195)
    else:
        bg_color = tuple(map(int, np.mean(bg_patch.reshape(-1, 3), axis=0)))

    cv2.rectangle(
        animated,
        (x1 - 30, y1 - 30),
        (x2 + 40, y2 + 30),
        bg_color,
        -1
    )

    progress = ease_out((local_t - 0.35) / 1.8)

    if progress <= 0:
        return animated

    start_x = x1 - 430
    end_x = x1
    current_x = int(start_x + (end_x - start_x) * progress)

    paste_x1 = max(0, current_x)
    paste_x2 = min(WIDTH, current_x + word_img.shape[1])

    src_x1 = paste_x1 - current_x
    src_x2 = src_x1 + (paste_x2 - paste_x1)

    if paste_x2 > paste_x1:
        animated[y1:y2, paste_x1:paste_x2] = word_img[:, src_x1:src_x2]

    return animated


def animate_slide_4_points(slide, t):
    local_t = t - 31
    original = slide.copy()
    animated = slide.copy()

    cv2.rectangle(animated, (20, 215), (1320, 940), (255, 255, 255), -1)

    rows = [
        ((20, 230, 1300, 370), 3.0),
        ((20, 380, 1300, 520), 12.2),   # changed: appears earlier
        ((20, 530, 1300, 670), 21.0),
        ((20, 680, 1300, 840), 29.0)
    ]

    for box, start_time in rows:
        x1, y1, x2, y2 = box

        progress = ease_out((local_t - start_time) / 1.1)

        if progress <= 0:
            continue

        row_img = original[y1:y2, x1:x2].copy()

        start_x = x1 - 260
        current_x = int(start_x + (x1 - start_x) * progress)

        paste_x1 = max(0, current_x)
        paste_x2 = min(WIDTH, current_x + row_img.shape[1])

        src_x1 = paste_x1 - current_x
        src_x2 = src_x1 + (paste_x2 - paste_x1)

        if paste_x2 > paste_x1:
            animated[y1:y2, paste_x1:paste_x2] = row_img[:, src_x1:src_x2]

    return animated


def animate_slide_6(slide, t):
    local_t = t - 95
    original = slide.copy()
    animated = slide.copy()

    heading_box = (40, 30, 1500, 170)
    content_box = (55, 250, 980, 760)

    hx1, hy1, hx2, hy2 = heading_box
    cx1, cy1, cx2, cy2 = content_box

    if local_t < 2:
        cv2.rectangle(animated, (hx1, hy1), (hx2, hy2), (255, 255, 255), -1)
        cv2.rectangle(animated, (cx1, cy1), (cx2, cy2), (255, 255, 255), -1)
        return animated

    if local_t < 6:
        cv2.rectangle(animated, (cx1, cy1), (cx2, cy2), (255, 255, 255), -1)
        return animated

    cv2.rectangle(animated, (cx1, cy1), (cx2, cy2), (255, 255, 255), -1)

    progress = ease_out((local_t - 6) / 1.0)
    block = original[cy1:cy2, cx1:cx2].copy()

    scale = 0.90 + 0.10 * progress
    new_w = int(block.shape[1] * scale)
    new_h = int(block.shape[0] * scale)

    resized = cv2.resize(block, (new_w, new_h), interpolation=cv2.INTER_LINEAR)

    px = cx1 + (block.shape[1] - new_w) // 2
    py = cy1 + (block.shape[0] - new_h) // 2

    roi = animated[py:py + new_h, px:px + new_w]
    blended = cv2.addWeighted(roi, 1 - progress, resized, progress, 0)

    animated[py:py + new_h, px:px + new_w] = blended

    return animated


def remove_green_tint(img, alpha):
    img = img.astype(np.float32)
    b, g, r = cv2.split(img)

    green_spill = g - np.maximum(r, b)
    green_spill = np.clip(green_spill, 0, 180)

    g = g - green_spill * 1.7
    r = r + green_spill * 0.20
    b = b - green_spill * 0.10

    img = cv2.merge([b, g, r])
    img = np.clip(img, 0, 255).astype(np.uint8)

    edge = ((alpha > 5) & (alpha < 250)).astype(np.uint8) * 255
    edge = cv2.dilate(edge, np.ones((7, 7), np.uint8), iterations=1)
    edge = cv2.GaussianBlur(edge, (9, 9), 0)

    smooth = cv2.bilateralFilter(img, 7, 50, 50)
    img[edge > 0] = smooth[edge > 0]

    return img


def ai_cut_faculty(frame):
    rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    pil_img = Image.fromarray(rgb)

    cut = remove(pil_img, session=session, alpha_matting=False)
    cut_np = np.array(cut)

    rgb_cut = cut_np[:, :, :3]
    alpha = cut_np[:, :, 3]

    bgr_cut = cv2.cvtColor(rgb_cut, cv2.COLOR_RGB2BGR)

    alpha = cv2.medianBlur(alpha, 5)
    alpha = cv2.GaussianBlur(alpha, (5, 5), 0)

    bgr_cut = remove_green_tint(bgr_cut, alpha)

    return bgr_cut, alpha


def get_stable_crop_box(alpha):
    global STABLE_BOX

    ys, xs = np.where(alpha > 20)

    if len(xs) == 0 or len(ys) == 0:
        return STABLE_BOX

    x1, x2 = xs.min(), xs.max()
    y1, y2 = ys.min(), ys.max()

    x1 = max(0, x1 - 25)
    x2 = min(alpha.shape[1], x2 + 25)
    y1 = max(0, y1 - 15)
    y2 = min(alpha.shape[0], y2 + 5)

    new_box = np.array([x1, y1, x2, y2], dtype=np.float32)

    if STABLE_BOX is None:
        STABLE_BOX = new_box
    else:
        STABLE_BOX = STABLE_BOX * 0.96 + new_box * 0.04

    return STABLE_BOX.astype(int)


def prepare_faculty(frame):
    faculty, alpha = ai_cut_faculty(frame)

    box = get_stable_crop_box(alpha)

    if box is not None:
        x1, y1, x2, y2 = box
        faculty = faculty[y1:y2, x1:x2]
        alpha = alpha[y1:y2, x1:x2]

    scale = FACULTY_HEIGHT / faculty.shape[0]
    faculty_width = int(faculty.shape[1] * scale)

    faculty = cv2.resize(
        faculty,
        (faculty_width, FACULTY_HEIGHT),
        interpolation=cv2.INTER_LANCZOS4
    )

    alpha = cv2.resize(
        alpha,
        (faculty_width, FACULTY_HEIGHT),
        interpolation=cv2.INTER_LANCZOS4
    )

    alpha = cv2.GaussianBlur(alpha, (7, 7), 0)
    alpha = alpha.astype(np.float32) / 255.0
    alpha_3 = cv2.merge([alpha, alpha, alpha])

    return faculty, alpha_3


def place_faculty(slide, frame, position):
    faculty, alpha_3 = prepare_faculty(frame)

    faculty_h, faculty_w = faculty.shape[:2]

    if position == "center":
        x = (WIDTH - faculty_w) // 2
    else:
        x = WIDTH - faculty_w - RIGHT_MARGIN

    y = HEIGHT - faculty_h

    x = max(0, x)
    y = max(0, y)

    roi = slide[y:y + faculty_h, x:x + faculty_w]

    min_h = min(roi.shape[0], faculty_h)
    min_w = min(roi.shape[1], faculty_w)

    roi = roi[:min_h, :min_w]
    faculty = faculty[:min_h, :min_w]
    alpha_3 = alpha_3[:min_h, :min_w]

    blended = (
        faculty.astype(np.float32) * alpha_3 +
        roi.astype(np.float32) * (1 - alpha_3)
    )

    slide[y:y + min_h, x:x + min_w] = blended.astype(np.uint8)

    return slide


def create_video(slide_paths):
    slides = []

    for path in slide_paths:
        img = cv2.imread(path)

        if img is None:
            raise FileNotFoundError(f"Slide not found: {path}")

        img = cv2.resize(img, (WIDTH, HEIGHT), interpolation=cv2.INTER_AREA)
        slides.append(img)

    cap = cv2.VideoCapture(FACULTY_VIDEO)

    if not cap.isOpened():
        raise FileNotFoundError("Cannot open faculty video")

    faculty_fps = cap.get(cv2.CAP_PROP_FPS)

    if faculty_fps <= 1:
        faculty_fps = 30

    out = cv2.VideoWriter(
        TEMP_VIDEO,
        cv2.VideoWriter_fourcc(*"mp4v"),
        OUTPUT_FPS,
        (WIDTH, HEIGHT)
    )

    total_frames = int(TOTAL_DURATION * OUTPUT_FPS)

    for frame_no in range(total_frames):
        t = frame_no / OUTPUT_FPS

        slide_index, position = get_slide_info(t)
        slide = slides[slide_index].copy()

        if slide_index == 1:
            slide = animate_slide_2(slide, t)

        if slide_index == 3:
            slide = animate_slide_4_points(slide, t)

        if slide_index == 5:
            slide = animate_slide_6(slide, t)

        if position != "hide":
            faculty_time = t - AUDIO_START_TIME

            if faculty_time >= 0:
                faculty_frame_no = int(faculty_time * faculty_fps)
                cap.set(cv2.CAP_PROP_POS_FRAMES, faculty_frame_no)

                ret, frame = cap.read()

                if ret:
                    slide = place_faculty(slide, frame, position)

        out.write(slide)

        if frame_no % 50 == 0:
            print(f"Processing frame {frame_no}/{total_frames}")

    cap.release()
    out.release()


def add_audio():
    delay_ms = int(AUDIO_START_TIME * 1000)

    run_cmd([
        FFMPEG_PATH,
        "-y",
        "-i", TEMP_VIDEO,
        "-i", FACULTY_VIDEO,
        "-filter_complex", f"[1:a]adelay={delay_ms}|{delay_ms}[a]",
        "-map", "0:v:0",
        "-map", "[a]",
        "-c:v", "libx264",
        "-preset", "ultrafast",
        "-crf", "20",
        "-pix_fmt", "yuv420p",
        "-r", str(OUTPUT_FPS),
        "-c:a", "aac",
        "-shortest",
        FINAL_VIDEO
    ])


if __name__ == "__main__":
    print("Converting PPT to images...")
    slide_paths = convert_ppt_to_images()

    print("Creating synced video...")
    create_video(slide_paths)

    print("Adding audio...")
    add_audio()

    print("Completed successfully!")
    print("Output:", FINAL_VIDEO)