"""
Voucher Number Scanner - Step 1
Video capture + draggable box overlay + OCR reading

FILES:
   voucher_scanner.py  <- this file: VoucherScanner class, camera/box/zoom/UI loop
   ocr_base.py         <- abstract interface all OCR engines implement
   ocr_tesseract.py    <- active OCR engine (Tesseract + preprocessing)
   ocr_paddle.py        <- stub for PaddleOCR, not implemented yet

To switch OCR engines later: at the bottom of this file, swap which class
you instantiate — `PaddleEngine()` or `TesseractEngine()`. Both implement
the same OCREngine interface, so nothing else needs to change.

SETUP REQUIRED (one-time, outside pip):
1. Install Tesseract OCR engine (this is NOT the python package, it's the actual engine):
   - Windows: https://github.com/UB-Mannheim/tesseract/wiki (installer .exe)
     After install, note the path, usually: C:\\Program Files\\Tesseract-OCR\\tesseract.exe
   - Mac: brew install tesseract
   - Linux: sudo apt install tesseract-ocr

2. Install DroidCam (or similar) on your phone + PC client, connect via USB,
   it will show up as a regular webcam.

PIP INSTALL:
   pip install opencv-python pytesseract

NEXT STEPS (not yet implemented, coming later):
   - auto-clear "Last inserted" display after some time
   - reset OCR history when the box is redrawn (avoid stale streaks)
"""

import cv2
import numpy as np
import pyperclip
import keyboard
from collections import deque
from ocr_tesseract import TesseractEngine
from ocr_paddle import PaddleEngine


class VoucherScanner:
    def __init__(self, ocr_engine, camera_index=0,
                 box_width=400, box_height=80,
                 zoom_step=0.1, zoom_min=1.0, zoom_max=4.0,
                 ocr_every_n_frames=2,
                 capture_group_size=4, capture_delimiter="-",
                 info_bar_height=60,
                 confirm_streak=3, min_lock_digits=6,
                 insert_hotkey="\\"):
        self.ocr_engine = ocr_engine
        self.camera_index = camera_index
        self.box_width = box_width
        self.box_height = box_height
        self.zoom_step = zoom_step
        self.zoom_min = zoom_min
        self.zoom_max = zoom_max
        self.ocr_every_n_frames = ocr_every_n_frames
        self.capture_group_size = capture_group_size
        self.capture_delimiter = capture_delimiter
        self.info_bar_height = info_bar_height
        self.confirm_streak = confirm_streak      # consecutive matching reads to lock
        self.min_lock_digits = min_lock_digits     # ignore short/partial reads
        self.insert_hotkey = insert_hotkey         # global hotkey to insert + advance

        # Mutable runtime state (previously module-level globals)
        self.box = None            # [x1, y1, x2, y2], set on first frame or drag
        self.dragging = False
        self.drag_start = None
        self.rotate_180 = False
        self.zoom = 1.0
        self.frame_count = 0
        self.detected = ""         # raw digits only — this is what later gets
                                    # written to Excel, never delimiter-formatted

        # Auto-lock state
        self.history = deque(maxlen=confirm_streak)
        self.locked = False
        self.locked_value = ""
        self._last_copied = None   # avoids re-copying the same value every frame
        self._last_inserted = None

        self.window_name = "Voucher Scanner"
        self.cap = None

    def _format_display(self, digits):
        """Groups raw digits with the configured delimiter for on-screen display
        only. The underlying self.detected stays pure digits — this is purely
        cosmetic and never affects what gets used elsewhere (e.g. Excel insert)."""
        if not digits or self.capture_group_size <= 0:
            return digits
        groups = [digits[i:i + self.capture_group_size]
                  for i in range(0, len(digits), self.capture_group_size)]
        return self.capture_delimiter.join(groups)

    # ---- mouse handling ----------------------------------------------
    def _mouse_callback(self, event, x, y, flags, param):
        if event == cv2.EVENT_LBUTTONDOWN:
            self.dragging = True
            self.drag_start = (x, y)
            self.box = [x, y, x, y]

        elif event == cv2.EVENT_MOUSEMOVE and self.dragging:
            self.box = [self.drag_start[0], self.drag_start[1], x, y]

        elif event == cv2.EVENT_LBUTTONUP:
            self.dragging = False
            if self.box is not None:
                x1, y1, x2, y2 = self.box
                # normalize so x1<x2 and y1<y2 regardless of drag direction
                self.box = [min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2)]

    # ---- frame processing ----------------------------------------------
    def _apply_zoom(self, frame):
        """Crops toward the center of the frame and scales back up,
        simulating a digital zoom."""
        if self.zoom <= 1.0:
            return frame

        h, w = frame.shape[:2]
        new_w, new_h = int(w / self.zoom), int(h / self.zoom)
        x1 = (w - new_w) // 2
        y1 = (h - new_h) // 2
        cropped = frame[y1:y1 + new_h, x1:x1 + new_w]
        return cv2.resize(cropped, (w, h), interpolation=cv2.INTER_LINEAR)

    def _insert_and_advance(self):
        """Called from the `keyboard` library's global hook — runs regardless
        of which window has focus (Excel, browser, etc). Types the raw
        digits into whatever field is currently focused, presses Enter to
        advance to the next cell, then resets the lock so the same voucher
        can't be double-inserted by an accidental second press."""
        if not self.locked_value:
            return
        keyboard.write(self.locked_value)
        keyboard.send('enter')
        self._last_inserted = self.locked_value

        self.locked = False
        self.history.clear()

    def _ensure_default_box(self, w, h):
        if self.box is None:
            x1 = (w - self.box_width) // 2
            y1 = int(h * 0.35)
            self.box = [x1, y1, x1 + self.box_width, y1 + self.box_height]

    def _update_lock(self, reading):
        """Feeds a new OCR reading into the rolling history and locks once
        `confirm_streak` consecutive reads agree. Keeps running even while
        already locked, so swapping to a new voucher auto-relocks onto the
        new number without any manual unlock step."""
        if len(reading) < self.min_lock_digits:
            self.history.clear()
            self.locked = False
            return

        self.history.append(reading)

        if len(self.history) == self.confirm_streak and len(set(self.history)) == 1:
            if reading != self.locked_value:
                self.locked_value = reading
            self.locked = True

            if self.locked_value != self._last_copied:
                pyperclip.copy(self.locked_value)   # raw digits only, no delimiter
                self._last_copied = self.locked_value
        elif reading != self.locked_value:
            # A different reading is coming in but hasn't been confirmed yet —
            # unlock so the UI honestly reflects "still checking", not stale data.
            self.locked = False

    def _compose_display(self, frame, width):
        """Builds a separate black info bar below the video frame (not
        overlaid on top of the pixels) and stacks them with vconcat."""
        bar = np.zeros((self.info_bar_height, width, 3), dtype=np.uint8)

        if self.locked:
            copied_flag = " (copied)" if self.locked_value == self._last_copied else ""
            status_text = f"LOCKED: {self._format_display(self.locked_value)}{copied_flag}"
            status_color = (0, 200, 0)
        else:
            status_text = f"Reading: {self._format_display(self.detected)}"
            status_color = (0, 255, 255)

        cv2.putText(bar, status_text, (15, 25),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.7, status_color, 2)

        last_inserted_text = f"Last inserted: {self._format_display(self._last_inserted)}" \
            if self._last_inserted else "Last inserted: -"
        cv2.putText(bar, last_inserted_text, (15, 50),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.5, (200, 200, 200), 1)
        cv2.putText(bar, f"Zoom: {self.zoom:.1f}x", (width - 130, 25),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 0), 1)

        return cv2.vconcat([frame, bar])

    def _handle_key(self, key):
        """Returns False if the app should quit."""
        if key == ord('q'):
            return False
        elif key == ord('r'):
            self.rotate_180 = not self.rotate_180
        elif key in (ord('+'), ord('=')):
            self.zoom = min(self.zoom_max, round(self.zoom + self.zoom_step, 2))
        elif key == ord('-'):
            self.zoom = max(self.zoom_min, round(self.zoom - self.zoom_step, 2))
        elif key == ord('c'):
            if self.locked_value:
                pyperclip.copy(self.locked_value)
                self._last_copied = self.locked_value
        return True

    # ---- main loop ----------------------------------------------
    def run(self):
        self.cap = cv2.VideoCapture(self.camera_index)
        if not self.cap.isOpened():
            print(f"Could not open camera index {self.camera_index}. Try a different index.")
            return

        cv2.namedWindow(self.window_name)
        cv2.setMouseCallback(self.window_name, self._mouse_callback)

        # suppress=True stops the raw backslash keypress itself from also
        # landing in whatever field is focused, before our inserted digits do
        keyboard.add_hotkey(self.insert_hotkey, self._insert_and_advance, suppress=True)

        print("Controls: [r] = toggle rotate 180  |  [+/-] = zoom in/out  |  [c] = re-copy locked value  |  [q] = quit")
        print(f"Global hotkey [{self.insert_hotkey}] inserts the locked value + Enter, works even outside this window.")
        print("Left-click and drag on the video to set the capture box.")

        try:
            while True:
                ok, frame = self.cap.read()
                if not ok:
                    print("Failed to read frame from camera.")
                    break

                if self.rotate_180:
                    frame = cv2.rotate(frame, cv2.ROTATE_180)

                frame = self._apply_zoom(frame)

                h, w = frame.shape[:2]
                self._ensure_default_box(w, h)
                x1, y1, x2, y2 = self.box

                self.frame_count += 1
                if self.frame_count % self.ocr_every_n_frames == 0 and x2 > x1 and y2 > y1:
                    crop = frame[y1:y2, x1:x2]
                    if crop.size > 0:
                        self.detected, processed = self.ocr_engine.run(crop)
                        self._update_lock(self.detected)
                        cv2.imshow("OCR Input (what the engine sees)", processed)
                    else:
                        self.detected = ""

                # Alignment box: green when locked, yellow while still checking
                box_color = (0, 200, 0) if self.locked else (0, 255, 255)
                cv2.rectangle(frame, (x1, y1), (x2, y2), box_color, 2)

                display_frame = self._compose_display(frame, w)
                cv2.imshow(self.window_name, display_frame)

                key = cv2.waitKey(1) & 0xFF
                if not self._handle_key(key):
                    break
        finally:
            keyboard.remove_hotkey(self.insert_hotkey)
            self.cap.release()
            cv2.destroyAllWindows()


if __name__ == "__main__":
    engine = PaddleEngine()
    # engine = TesseractEngine()  # swap back to this if you want to compare

    scanner = VoucherScanner(ocr_engine=engine, camera_index=0)
    scanner.run()