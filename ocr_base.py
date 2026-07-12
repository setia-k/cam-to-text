"""
Base interface all OCR engines must implement.
"""

from abc import ABC, abstractmethod


class OCREngine(ABC):
    @abstractmethod
    def run(self, crop):
        """
        Args:
            crop: BGR image (numpy array) — the cropped box region.
        Returns:
            (text: str, debug_image: np.ndarray) — debug_image is what
            the engine actually processed/saw, shown in the OCR preview window.
        """
        raise NotImplementedError
