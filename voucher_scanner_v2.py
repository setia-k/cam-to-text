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
  - Two passes: pass 1 goes down from row 1, pass 2 goes up from row 100
    (`d` switches). Pass 2 is kept in row order: each new card is placed at
    its row (100, 99, 98, ...) and the screen/file read like the Excel column.
  - In pass 2 every lock is compared with pass 1 at the same row: MATCH, or
    MISMATCH with the expected number and where the card actually is in
    pass 1. It only warns; `\\` still inserts.
  - `e` copies the current pass to the clipboard, one value per line —
    paste at A1 (pass 1) or B1 (pass 2). Pass 2 is padded with blank lines on
    top if unfinished so rows stay aligned. Format the column as Text.

The last session is resumed automatically on start. To start a new one, press
N twice in the app, or launch with --new. The old session.json is always
renamed to session_<timestamp>.json, never overwritten or deleted.
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


SESSION_VERSION = 2   # 2: pass 2 is stored in row order (see PassState)


class PassState:
    def __init__(self):
        # Pass 1: scan order == row order (entries[0] is row 1).
        # Pass 2: row order too, newest scan first: each new card goes in at
        # the front and pushes the others down, so entries[-1] is always the
        # last row (e.g. 100). The list reads like the Excel column.
        self.entries = []
        self.stack_ends = []   # len(entries) at each closed stack


class VoucherScannerV2(VoucherScanner):
    def __init__(self, *args, stack_size=25, total_rows=100, fresh=False, **kwargs):
        kwargs.setdefault("info_bar_height", 120)
        super().__init__(*args, **kwargs)
        self.stack_size = stack_size
        self.total_rows = total_rows   # pass 2 starts at this row and counts down
        self.passes = {1: PassState(), 2: PassState()}
        self.pass_no = 1
        self.message = ""
        self.message_until = 0.0
        self._new_confirm_until = 0.0
        self._start_session(fresh)

    # ---- session persistence ----------------------------------------
    @property
    def cur(self):
        return self.passes[self.pass_no]

    def _archive_session(self):
        """Rename session.json to session_<timestamp>.json (never deleted)."""
        stamp = time.strftime("_%Y%m%d_%H%M%S")
        backup = SESSION_FILE.replace(".json", f"{stamp}.json")
        n = 1
        while os.path.exists(backup):   # never overwrite an earlier archive
            n += 1
            backup = SESSION_FILE.replace(".json", f"{stamp}_{n}.json")
        try:
            os.replace(SESSION_FILE, backup)
            print(f"Previous session kept as {os.path.basename(backup)}")
            return True
        except OSError as e:
            print(f"Could not archive old session: {e}")
            return False

    def _new_session(self):
        """In-app fresh start (key N, pressed twice): archive and clear."""
        if os.path.exists(SESSION_FILE) and not self._archive_session():
            self._say("Could not archive the old session - not cleared", 6)
            return
        self.passes = {1: PassState(), 2: PassState()}
        self.pass_no = 1
        self.history.clear()
        self.locked = False
        self.locked_value = ""
        self._save()
        self._say("New session started (old one archived)", 6)

    def _start_session(self, fresh):
        if not os.path.exists(SESSION_FILE):
            return
        if not fresh:
            try:
                with open(SESSION_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                for n in (1, 2):
                    p = data.get(f"pass{n}", {})
                    self.passes[n].entries = list(p.get("entries", []))
                    self.passes[n].stack_ends = list(p.get("stack_ends", []))
                if data.get("version", 1) < 2:
                    self.passes[2].entries.reverse()   # old files stored pass 2 in scan order
                self.pass_no = int(data.get("pass_no", 1))
                print(f"Resumed session: pass1={len(self.passes[1].entries)}, "
                      f"pass2={len(self.passes[2].entries)}")
                return
            except (OSError, ValueError):
                print("Could not read session.json, starting fresh.")
        self._archive_session()

    def _save(self):
        data = {"version": SESSION_VERSION, "pass_no": self.pass_no}
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

    # ---- rows -------------------------------------------------------
    def _row_of(self, idx):
        """Excel row of entries[idx] in the current pass."""
        if self.pass_no == 1:
            return idx + 1
        return self.total_rows - len(self.cur.entries) + 1 + idx

    def _next_row(self):
        """Row the next insert in the current pass would take."""
        if self.pass_no == 1:
            return len(self.cur.entries) + 1
        return self.total_rows - len(self.cur.entries)

    def _newest_idx(self):
        return len(self.cur.entries) - 1 if self.pass_no == 1 else 0

    def _check_vs_pass1(self, value):
        """Pass 2 only. Returns None if there's nothing to compare, else
        (matches, expected_value, row_in_pass1_where_value_is_or_None)."""
        if self.pass_no != 2:
            return None
        p1 = self.passes[1].entries
        row = self._next_row()
        if not 1 <= row <= len(p1):
            return None
        where = p1.index(value) + 1 if value in p1 else None
        return p1[row - 1] == value, p1[row - 1], where

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
            row = self._row_of(idx)
            if idx == self._newest_idx():
                self._say(f"Already inserted as #{row} - show the next card")
            else:
                self._say(f"BLOCKED: duplicate of #{row} (press u to remove the last entry if needed)")
            return
        if self.pass_no == 2 and len(self.cur.entries) >= self.total_rows:
            self._say(f"Pass 2 is full ({self.total_rows} rows)")
            return

        row = self._next_row()
        check = self._check_vs_pass1(self.locked_value)   # before inserting: row changes
        if self.pass_no == 1:
            self.cur.entries.append(self.locked_value)
        else:
            self.cur.entries.insert(0, self.locked_value)
        self._save()
        if check is not None and not check[0]:
            self._say(f"Inserted #{row} but it does NOT match pass 1 (expected {self._format_display(check[1])})", 8)
        else:
            self._say(f"Inserted #{row}", 2)

    def _undo(self):
        p = self.cur
        if not p.entries:
            self._say("Nothing to undo")
            return
        row = self._row_of(self._newest_idx())
        removed = p.entries.pop() if self.pass_no == 1 else p.entries.pop(0)
        while p.stack_ends and p.stack_ends[-1] > len(p.entries):
            p.stack_ends.pop()   # undo crossed back over a closed stack
        self._save()
        self._say(f"Removed #{row}: {self._format_display(removed)}")

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
        pad = 0
        if self.pass_no == 2:
            # entries are already in row order; blank lines up top keep each
            # value on its own row when pasted at the top of the column, even
            # if pass 2 isn't finished yet
            pad = max(0, self.total_rows - len(values))
            values = [""] * pad + values
        pyperclip.copy("\n".join(values))
        target = "paste at B1" if self.pass_no == 2 else "paste at A1"
        extra = f", {pad} blank rows on top" if pad else ""
        self._say(f"Copied {len(self.cur.entries)} values from pass {self.pass_no}{extra} - {target}", 6)

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
            if key == ord('N'):
                if time.time() < self._new_confirm_until:
                    self._new_confirm_until = 0.0
                    self._new_session()
                else:
                    self._new_confirm_until = time.time() + 4
                    self._say("Press N again to archive this session and start a new one", 4)
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
        hint = None   # extra explanation for line 4 (pass 2 mismatch)
        if self.paused:
            text, color = "PAUSED (p to resume)", (0, 165, 255)
        elif self.locked:
            idx = self._dup_index(self.locked_value)
            shown = self._format_display(self.locked_value)
            if idx is None:
                row = self._next_row()
                check = self._check_vs_pass1(self.locked_value)
                if check is None:
                    text, color = f"LOCKED #{row}: {shown}", (0, 200, 0)
                elif check[0]:
                    text, color = f"LOCKED #{row}: {shown}  MATCH", (0, 200, 0)
                else:
                    text, color = f"MISMATCH #{row}: {shown}", (0, 128, 255)
                    where = (f"this card is pass 1 #{check[2]}" if check[2]
                             else "this card is not in pass 1")
                    hint = f"pass 1 #{row} is {self._format_display(check[1])} - {where}"
            elif idx == self._newest_idx():
                text, color = f"#{self._row_of(idx)} inserted - next card", (200, 200, 0)
            else:
                text, color = f"DUPLICATE of #{self._row_of(idx)}: {shown}", (0, 0, 255)
        else:
            text = (f"Reading: {self._format_display(self.detected)} "
                    f"({len(self.detected)}/{self._expected_digits()}{self._conf_text()})")
            color = (0, 255, 255)
        cv2.putText(bar, text, (15, 25), cv2.FONT_HERSHEY_SIMPLEX, 0.7, color, 2)

        # line 2: pass + counters
        count_color = (0, 200, 0) if in_stack == self.stack_size else \
                      ((0, 0, 255) if in_stack > self.stack_size else (255, 255, 255))
        cv2.putText(bar, f"Pass {self.pass_no}  Total {n}  Next #{self._next_row()}  "
                         f"Stack {len(p.stack_ends) + 1}: {in_stack}/{self.stack_size}",
                    (15, 55), cv2.FONT_HERSHEY_SIMPLEX, 0.65, count_color, 2)

        # line 3: last 3 entries, newest first, with their rows
        newest = self._newest_idx()
        step = -1 if self.pass_no == 1 else 1
        tail = "  ".join(f"#{self._row_of(newest + step * i)}:{self._format_display(p.entries[newest + step * i])}"
                         for i in range(min(3, n)))
        cv2.putText(bar, tail or "(no entries yet)", (15, 80),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.45, (200, 200, 200), 1)

        # line 4: message / mismatch hint / key help
        if time.time() < self.message_until:
            cv2.putText(bar, self.message, (15, 105),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.55, (0, 255, 255), 1)
        elif hint:
            cv2.putText(bar, hint, (15, 105),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.55, (0, 128, 255), 1)
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

    def _show_debug(self, processed):
        """'What the engine sees' window, enlarged, with the number below it:
        the locked value (green) or the current read (yellow), formatted."""
        import numpy as np
        h, w = processed.shape[:2]
        target_w = max(w, 640)
        img = cv2.resize(processed, (target_w, int(h * target_w / w)),
                         interpolation=cv2.INTER_CUBIC) if target_w != w else processed

        strip = np.zeros((70, img.shape[1], 3), dtype=np.uint8)
        if self.locked:
            value, color, state = self.locked_value, (0, 200, 0), "LOCKED"
        else:
            value, color, state = self.detected, (0, 255, 255), "reading"
        shown = self._format_display(value) if value else "-"
        cv2.putText(strip, shown, (15, 38), cv2.FONT_HERSHEY_SIMPLEX, 1.0, color, 2)
        cv2.putText(strip, f"{state}  {len(value)}/{self._expected_digits()}{self._conf_text()}",
                    (15, 60), cv2.FONT_HERSHEY_SIMPLEX, 0.45, (200, 200, 200), 1)
        cv2.imshow("OCR Input (what the engine sees)", cv2.vconcat([img, strip]))

    def _print_controls(self):
        print("v2 controls: [\\] add to list (global)  |  [u] undo last  |  [n] stack done  |  [N,N] new session  |  "
              "[e] export to clipboard  |  [d] switch pass 1/2  |  [f] cycle format  |  "
              "[t] type format  |  [p] pause OCR  |  [ / ] min confidence  |  [i] idle skip  |  "
              "[r] rotate  |  [+/-] zoom  |  [q] quit")
        print("Left-click and drag on the video to set the capture box.")


if __name__ == "__main__":
    engine = PaddleEngine()
    scanner = VoucherScannerV2(ocr_engine=engine, camera_index=0,
                               fresh="--new" in sys.argv)
    scanner.run()
