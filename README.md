# Voucher Number Scanner

Reads voucher/PIN numbers off a phone camera feed and locks onto a stable
reading automatically — built to cut down on repetitive numpad typing during
manual data entry, not to fully replace verification.

## Files
- `voucher_scanner.py` — main app: `VoucherScanner` class (camera, draggable
  box, zoom, rotate, auto-lock, clipboard, global hotkey)
- `ocr_base.py` — abstract `OCREngine` interface
- `ocr_tesseract.py` — Tesseract engine implementation
- `ocr_paddle.py` — PaddleOCR engine implementation (currently active, more
  accurate on stylized/textured voucher fonts)

## One-time setup

### 1. Phone → PC video feed (Windows)
No app needed on the phone. Uses scrcpy's camera mode + OBS Virtual Camera:

1. Install [scrcpy](https://github.com/Genymobile/scrcpy/releases) (unzip anywhere)
2. On phone: enable USB debugging (Settings → About Phone → tap Build Number
   x7 → Developer Options → USB Debugging)
3. Plug phone in via USB, accept the debugging prompt on the phone
4. In the scrcpy folder, run:
   ```
   scrcpy --video-source=camera --camera-facing=back
   ```
5. Install [OBS Studio](https://obsproject.com)
6. In OBS: Sources → `+` → Window Capture → select the scrcpy window
7. Click **Start Virtual Camera** (bottom right of OBS)

### 2. Tesseract OCR engine (system install, not pip)
Only needed if using the Tesseract engine instead of PaddleOCR:
- Windows installer: https://github.com/UB-Mannheim/tesseract/wiki
- If not on PATH, set the path manually at the top of `ocr_tesseract.py`

### 3. Python dependencies
```
pip install -r requirements.txt
```
First run of the PaddleOCR engine downloads its recognition model
automatically (needs internet once, cached after that).

## Running
```
python voucher_scanner.py
```
`CAMERA_INDEX` inside `VoucherScanner.__init__` may need adjusting — OBS
Virtual Camera usually isn't index 0 if you have a laptop webcam too.

## Controls
| Key / Action | Effect |
|---|---|
| Left-click + drag on video | Redraw the capture box |
| `r` | Toggle 180° rotation (for upside-down phone mounting) |
| `+` / `-` | Zoom in / out |
| `c` | Re-copy the currently locked value to clipboard |
| `q` | Quit |
| `\` (global hotkey) | **Not currently reliable** — intended to type the locked value into the focused field + Enter. Didn't work reliably in testing (likely needs Admin / hook permissions); clipboard + manual paste is the current workflow instead. |

## Workflow
1. Align a voucher number inside the yellow box
2. Box turns **green** once the reading is stable (3 consecutive matching
   reads) — the value auto-copies to clipboard at that moment
3. `Ctrl+V` into the target Excel cell
4. Swap to the next voucher — it auto-unlocks and relocks on the new number,
   no manual reset needed

## Switching OCR engines
At the bottom of `voucher_scanner.py`:
```python
engine = PaddleEngine()      # currently active — better on textured fonts
# engine = TesseractEngine()  # lighter weight, no model download
```
Both implement the same `OCREngine` interface, so nothing else changes.

## Known issues / possible next steps
- Global hotkey (`\`) doesn't reliably fire when a non-Python window (e.g.
  Excel) has focus — needs debugging (try running as Administrator first)
- OCR lock history isn't reset when the box is redrawn mid-session — could
  cause a stale lock briefly after moving the box
- Box position isn't saved between sessions — redrawing is manual each time
  you switch operator/voucher format
