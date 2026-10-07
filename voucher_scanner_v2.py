"""
Voucher Scanner v2 — ordered list instead of typing into Excel.

Builds on voucher_scanner.py (camera, box, zoom, formats, length check) and
changes only the workflow:

  - `\\` appends the locked value to an ordered list (no keystrokes into Excel,
    so no focus problems). The list is autosaved to session.json on every change.
  - The same card can't be added twice: a value already in this pass is blocked,
    and the screen says which entry it duplicates. Fix a mistake with `u` (undo
    last entry) and rescan.
  - Stack counter (25 per stack). `n` closes a stack and warns if the count
    isn't 25, so a skipped card is found per stack, not after 100 cards.
  - Two passes: pass 1 (forward), pass 2 (reverse). `d` switches pass. Export
    of pass 2 is flipped so it pastes in the same row order as pass 1.
  - `e` copies the current pass to the clipboard, one value per line —
    paste at A1 of the Excel column (format that column as Text, 12-digit
    numbers otherwise turn into scientific notation).

Old session.json is renamed to session_<timestamp>.json on start (never
overwritten); run with --resume to continue it instead.
"""

import json
import os
import sys
import time
import cv2
import pyperclip
from voucher_scanner import VoucherScanner
from ocr_paddle import PaddleEngine

SESSION_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "session.json")


class PassState:
    def __init__(self):
        self.entries = []      # ordered values, exactly as inserted
        self.stack_ends = []   # len(entries) at each closed stack


class VoucherScannerV2(VoucherScanner):
    def __init__(self, *args, stack_size=25, resume=False, **kwargs):
        kwargs.setdefault("info_bar_height", 120)
        super().__init__(*args, **kwargs)
        self.stack_size = stack_size
        self.passes = {1: PassState(), 2: PassState()}
        self.pass_no = 1
        self.message = ""
        self.message_until = 0.0
        self._start_session(resume)

    # ---- session persistence ----------------------------------------
    @property
    def cur(self):
        return self.passes[self.pass_no]

    def _start_session(self, resume):
        if not os.path.exists(SESSION_FILE):
            return
        if resume:
            try:
                with open(SESSION_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                for n in (1, 2):
                    p = data.get(f"pass{n}", {})
                    self.passes[n].entries = list(p.get("entries", []))
                    self.passes[n].stack_ends = list(p.get("stack_ends", []))
                self.pass_no = int(data.get("pass_no", 1))
                print(f"Resumed session: pass1={len(self.passes[1].entries)}, "
                      f"pass2={len(self.passes[2].entries)}")
                return
            except (OSError, ValueError):
                print("Could not read session.json, starting fresh.")
        backup = SESSION_FILE.replace(".json", time.strftime("_%Y%m%d_%H%M%S.json"))
        try:
            os.replace(SESSION_FILE, backup)
            print(f"Previous session kept as {os.path.basename(backup)}")
        except OSError as e:
            print(f"Could not archive old session: {e}")

    def _save(self):
        data = {"pass_no": self.pass_no}
        for n, p in self.passes.items():
            data[f"pass{n}"] = {"entries": p.entries, "stack_ends": p.stack_ends}
        try:
            tmp = SESSION_FILE + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=1)
            os.replace(tmp, SESSION_FILE)
        except OSError as e:
            print(f"Could not save session: {e}")

    def _say(self, text, seconds=5):
        self.message = text
        self.message_until = time.time() + seconds

    # ---- lock / insert ----------------------------------------------
    def _dup_index(self, value):
        """Index of `value` in the current pass, or None."""
        try:
            return self.cur.entries.index(value)
        except ValueError:
            return None

    def _update_lock(self, reading):
        """Same streak logic as v1 (length must match the format), minus the
        clipboard copy — in v2 the clipboard is only used by export."""
        if len(reading) != self._expected_digits():
            self.history.clear()
            self.locked = False
            return

        self.history.append(reading)
        if len(self.history) == self.confirm_streak and len(set(self.history)) == 1:
            self.locked_value = reading
            self.locked = True
        elif reading != self.locked_value:
            self.locked = False

    def _insert_and_advance(self):
        """Called from the keyboard hook — only touches a list, so it's fast
        enough to run inline (no thread needed like v1's typing)."""
        if not self.locked or not self.locked_value:
            return
        idx = self._dup_index(self.locked_value)
        if idx is not None:
            if idx == len(self.cur.entries) - 1:
                self._say(f"Already inserted as #{idx + 1} - show the next card")
            else:
                self._say(f"BLOCKED: duplicate of #{idx + 1} (press u to remove the last entry if needed)")
            return
        self.cur.entries.append(self.locked_value)
        self._save()
        n = len(self.cur.entries)
        self._say(f"Inserted #{n}", 2)

    def _undo(self):
        p = self.cur
        if not p.entries:
            self._say("Nothing to undo")
            return
        removed = p.entries.pop()
        while p.stack_ends and p.stack_ends[-1] > len(p.entries):
            p.stack_ends.pop()   # undo crossed back over a closed stack
        self._save()
        self._say(f"Removed #{len(p.entries) + 1}: {self._format_display(removed)}")

    def _close_stack(self):
        p = self.cur
        start = p.stack_ends[-1] if p.stack_ends else 0
        count = len(p.entries) - start
        if count == 0:
            self._say("Stack is empty")
            return
        p.stack_ends.append(len(p.entries))
        self._save()
        stack_no = len(p.stack_ends)
        if count == self.stack_size:
            self._say(f"Stack {stack_no} OK ({count}/{self.stack_size})", 6)
        else:
            self._say(f"Stack {stack_no} SHORT/OVER: {count}/{self.stack_size} - recount this stack", 10)

    def _export(self):
        values = list(self.cur.entries)
        if self.pass_no == 2:
            values.reverse()
        pyperclip.copy("\n".join(values))
        self._say(f"Copied {len(values)} values from pass {self.pass_no}"
                  f"{' (reversed)' if self.pass_no == 2 else ''} - paste at A1", 6)

    def _switch_pass(self):
        self.pass_no = 2 if self.pass_no == 1 else 1
        self.history.clear()
        self.locked = False
        self.locked_value = ""
        self._save()
        self._say(f"Pass {self.pass_no} ({len(self.cur.entries)} entries)")

    # ---- keys -------------------------------------------------------
    def _handle_key(self, key):
        if not self.typing_format and key not in (255, -1):
            if key == ord('u'):
                self._undo()
                return True
            if key == ord('e'):
                self._export()
                return True
            if key == ord('n'):
                self._close_stack()
                return True
            if key == ord('d'):
                self._switch_pass()
                return True
        return super()._handle_key(key)

    # ---- display ----------------------------------------------------
    def _compose_display(self, frame, width):
        if self.typing_format:
            return super()._compose_display(frame, width)

        import numpy as np
        bar = np.zeros((self.info_bar_height, width, 3), dtype=np.uint8)
        p = self.cur
        n = len(p.entries)
        start = p.stack_ends[-1] if p.stack_ends else 0
        in_stack = n - start

        # line 1: status
        if self.paused:
            text, color = "PAUSED (p to resume)", (0, 165, 255)
        elif self.locked:
            idx = self._dup_index(self.locked_value)
            shown = self._format_display(self.locked_value)
            if idx is None:
                text, color = f"LOCKED: {shown}", (0, 200, 0)
            elif idx == n - 1:
                text, color = f"#{idx + 1} inserted - next card", (200, 200, 0)
            else:
                text, color = f"DUPLICATE of #{idx + 1}: {shown}", (0, 0, 255)
        else:
            text = (f"Reading: {self._format_display(self.detected)} "
                    f"({len(self.detected)}/{self._expected_digits()}{self._conf_text()})")
            color = (0, 255, 255)
        cv2.putText(bar, text, (15, 25), cv2.FONT_HERSHEY_SIMPLEX, 0.7, color, 2)

        # line 2: pass + counters
        count_color = (0, 200, 0) if in_stack == self.stack_size else \
                      ((0, 0, 255) if in_stack > self.stack_size else (255, 255, 255))
        cv2.putText(bar, f"Pass {self.pass_no}   Total {n}   Stack {len(p.stack_ends) + 1}: "
                         f"{in_stack}/{self.stack_size}",
                    (15, 55), cv2.FONT_HERSHEY_SIMPLEX, 0.65, count_color, 2)

        # line 3: last 3 entries
        tail = "  ".join(f"#{n - i}:{self._format_display(p.entries[-1 - i])}"
                         for i in range(min(3, n)))
        cv2.putText(bar, tail or "(no entries yet)", (15, 80),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.45, (200, 200, 200), 1)

        # line 4: message / key help
        if time.time() < self.message_until:
            cv2.putText(bar, self.message, (15, 105),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.55, (0, 255, 255), 1)
        else:
            cv2.putText(bar, "\\ add  u undo  n stack done  e export  d pass  f/t format",
                        (15, 105), cv2.FONT_HERSHEY_SIMPLEX, 0.45, (150, 150, 150), 1)

        cv2.putText(bar, f"Zoom: {self.zoom:.1f}x", (width - 130, 25),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 0), 1)
        cv2.putText(bar, f"Format: {self.format_names[self.format_index]}  Len: {self._expected_digits()}",
                    (width - 230, 50), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 0), 1)
        cv2.putText(bar, self._tuning_text(), (width - 230, 75),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.45, (200, 200, 200), 1)
        return cv2.vconcat([frame, bar])

    def _print_controls(self):
        print("v2 controls: [\\] add to list (global)  |  [u] undo last  |  [n] stack done  |  "
              "[e] export to clipboard  |  [d] switch pass 1/2  |  [f] cycle format  |  "
              "[t] type format  |  [p] pause OCR  |  [ / ] min confidence  |  [i] idle skip  |  "
              "[r] rotate  |  [+/-] zoom  |  [q] quit")
        print("Left-click and drag on the video to set the capture box.")


if __name__ == "__main__":
    engine = PaddleEngine()
    scanner = VoucherScannerV2(ocr_engine=engine, camera_index=0,
                               resume="--resume" in sys.argv)
    scanner.run()
