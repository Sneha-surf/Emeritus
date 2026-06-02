import cv2
import numpy as np
import os

input_path = "Green-screen Faculty Recording.mp4"
output_path = "faculty_clean.mp4"

cap = cv2.VideoCapture(input_path)
fourcc = cv2.VideoWriter_fourcc(*'mp4v')
fps    = cap.get(cv2.CAP_PROP_FPS)
w      = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
h      = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
total  = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))

out = cv2.VideoWriter(output_path, fourcc, fps, (w, h))
frame_count = 0

print(f"Processing {total} frames at {fps} fps...")

while cap.isOpened():
    ret, frame = cap.read()
    if not ret:
        break

    hsv = cv2.cvtColor(frame, cv2.COLOR_BGR2HSV)

    # Green screen range — adjust if your green is bright/dark
    lower_green = np.array([35, 40, 40])
    upper_green = np.array([85, 255, 255])

    mask     = cv2.inRange(hsv, lower_green, upper_green)

    # Smooth mask edges to reduce fringe
    mask     = cv2.GaussianBlur(mask, (7, 7), 0)
    _, mask  = cv2.threshold(mask, 128, 255, cv2.THRESH_BINARY)
    mask_inv = cv2.bitwise_not(mask)

    result = cv2.bitwise_and(frame, frame, mask=mask_inv)
    out.write(result)

    frame_count += 1
    if frame_count % 100 == 0:
        print(f"  {frame_count}/{total} frames done...")

cap.release()
out.release()
print(f"\nDone! Saved to: {output_path}")