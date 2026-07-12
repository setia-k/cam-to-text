"""
PaddleOCR engine — recognition-only (no text detection).
Implements OCREngine (see ocr_base.py).

We use TextRecognition instead of the full PaddleOCR pipeline because we've
already cropped tightly to just the number line with the drag-box — running
full detection on top of that would be redundant and slower.

SETUP:
    pip install paddlepaddle paddleocr
    (first run downloads the recognition model automatically, ~tens of MB,
    needs internet once; cached locally after that)
"""

import cv2
from paddleocr import TextRecognition
from ocr_base import OCREngine


class PaddleEngine(OCREngine):
    def __init__(self, denoise_strength=10):
        # Loaded once here, not per-frame — model loading is slow (~seconds),
        # inference on each frame is fast.
        self._model = TextRecognition()
        self.denoise_strength = denoise_strength

    def _preprocess(self, crop):
        """Light cleaning only. PaddleOCR is a deep learning model trained on
        real-world text — heavy binarization/thresholding (like we do for
        Tesseract) tends to hurt it rather than help, since it strips away
        texture/gradient information the model was trained to use."""
        gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)
        gray = cv2.fastNlMeansDenoising(gray, h=self.denoise_strength)
        # Paddle expects a 3-channel image
        return cv2.cvtColor(gray, cv2.COLOR_GRAY2BGR)

    def run(self, crop):
        processed = self._preprocess(crop)
        result = self._model.predict(processed)

        if not result:
            return "", processed

        text = result[0]["rec_text"]
        digits = "".join(ch for ch in text if ch.isdigit())
        return digits, processed
