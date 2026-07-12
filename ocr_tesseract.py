"""
Tesseract OCR engine — implements OCREngine (see ocr_base.py).
"""

import cv2
import pytesseract
from ocr_base import OCREngine

# If Tesseract isn't in your system PATH, uncomment and set the path:
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


class TesseractEngine(OCREngine):
    def __init__(self, upscale=3, denoise_strength=15, close_iterations=2):
        self.upscale = upscale
        self.denoise_strength = denoise_strength
        self.close_iterations = close_iterations

    def _preprocess(self, crop):
        """grayscale -> upscale -> denoise -> binarize (Otsu) -> morphological closing."""
        gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)

        gray = cv2.resize(gray, None, fx=self.upscale, fy=self.upscale,
                           interpolation=cv2.INTER_CUBIC)

        gray = cv2.fastNlMeansDenoising(gray, h=self.denoise_strength)

        _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

        # Morphological closing: fills small gaps in strokes (helps with
        # hatched/textured/dot-matrix style fonts where strokes aren't solid).
        kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (3, 3))
        closed = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel,
                                   iterations=self.close_iterations)

        return closed

    def run(self, crop):
        processed = self._preprocess(crop)
        config = "--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789 "
        text = pytesseract.image_to_string(processed, config=config)
        return text.strip().replace(" ", ""), processed
