"""
PaddleOCR engine — recognition-only (no text detection).
Implements OCREngine (see ocr_base.py).

We use TextRecognition instead of the full PaddleOCR pipeline because we've
already cropped tightly to just the number line with the drag-box — running
full detection on top of that would be redundant and slower.

SETUP:
    pip install -r requirements.txt   (needs paddleocr >= 3.7 for PP-OCRv6)
    (first run downloads the recognition model automatically, needs internet
    once; cached locally after that)

MODELS (recognition): PP-OCRv6_tiny_rec (4 MB, fastest), PP-OCRv6_small_rec
(20 MB, default), PP-OCRv6_medium_rec (73 MB, most accurate). The old
PP-OCRv5_mobile_rec still works too — pass it as model_name.
"""

import cv2
from paddleocr import TextRecognition
from ocr_base import OCREngine


class PaddleEngine(OCREngine):
    def __init__(self, model_name="PP-OCRv6_medium_rec", denoise_strength=10):
        # Loaded once here, not per-frame — model loading is slow (~seconds),
        # inference on each frame is fast.
        self._model = TextRecognition(model_name=model_name)
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
            self.last_confidence = 0.0
            return "", processed

        text = result[0]["rec_text"]
        self.last_confidence = float(result[0].get("rec_score", 0.0))
        digits = "".join(ch for ch in text if ch.isdigit())
        return digits, processed
