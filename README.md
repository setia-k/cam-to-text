# Voucher Number Scanner

Reads voucher/PIN numbers off a phone camera feed and locks onto a stable
reading automatically — built to cut down on repetitive numpad typing during
manual data entry, not to fully replace verification.

## Files
- `voucher_scanner.py` — main app: `VoucherScanner` class (camera, draggable
  box, zoom, rotate, auto-lock, clipboard, global hotkey)
- `ocr_base.py` — abstract `OCREngine` interface
- `ocr_paddle.py` — PaddleOCR recognition-only engine (PP-OCRv6, CPU)
- `formats.json` — saved display formats (auto-created, editable)

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

### 2. Python dependencies
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
| `f` | Cycle display format |
| `t` | Type a format: saved name (`cw`) or pattern (`3,3,6`). Shows what you type + a live preview; new patterns are saved to `formats.json` |
| `p` | Pause/resume OCR |
| `d` | Toggle insert direction (Enter / Shift+Enter) |
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

## Choosing the OCR model
`PaddleEngine(model_name=...)` in `voucher_scanner.py`: `PP-OCRv6_tiny_rec`
(fastest), `PP-OCRv6_small_rec` (default), `PP-OCRv6_medium_rec` (most accurate).

## v2 (experimental): ordered list instead of typing into Excel
`python voucher_scanner_v2.py` (`--resume` to continue the last session).
`voucher_scanner.py` is untouched in behavior and remains the fallback.

| Key | Effect |
|---|---|
| `\` | Add locked value to the list (blocked if already in this pass) |
| `u` | Undo last entry (then it can be rescanned) |
| `n` | Close a stack; warns if the count isn't 25 |
| `d` | Switch pass 1 / pass 2 |
| `e` | Copy the pass to the clipboard, one per line (pass 2 is reversed) — paste at A1 |

Autosaved to `session.json`; a previous session is renamed, never overwritten.
Format the Excel column as Text before pasting (12-digit numbers otherwise
become scientific notation).

## Known issues / possible next steps
- Global hotkey (`\`) doesn't reliably fire when a non-Python window (e.g.
  Excel) has focus — needs debugging (try running as Administrator first)
- OCR lock history isn't reset when the box is redrawn mid-session — could
  cause a stale lock briefly after moving the box
- Box position isn't saved between sessions — redrawing is manual each time
  you switch operator/voucher format

## scrpcpy commands
`scrcpy --video-source=camera --camera-id=0 --camera-size=720x720 --camera-fps=30 -b 10M --no-audio`