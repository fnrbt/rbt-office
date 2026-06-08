#!/usr/bin/env python3
"""Drive the native spreadsheet-editor via its ControlServer FIFO.

Launches the editor, sends keyboard/mouse commands, captures framebuffer
screenshots, then exits. Keystrokes are X11 keycodes translated the same way
as widgets/x11-viewer/KeyMap.fs, so typing maps char -> (keycode, shift).

This is the Batch-1 e2e harness: type formulas into real cells and screenshot
the rendered results for close inspection.
"""
import glob
import os
import struct
import subprocess
import sys
import time

CMD_SCREENSHOT = 1
CMD_EXIT = 2
CMD_CLICK_AT = 11
CMD_SEND_KEY = 14

SHIFT = 50
ENTER = 36
ESC = 9
TAB = 23
BACKSPACE = 22
UP, DOWN, LEFT, RIGHT = 111, 116, 113, 114

# char -> (keycode, shift) mirroring KeyMap.translate
_BASE = {
    'a': 38, 'b': 56, 'c': 54, 'd': 40, 'e': 26, 'f': 41, 'g': 42, 'h': 43,
    'i': 31, 'j': 44, 'k': 45, 'l': 46, 'm': 58, 'n': 57, 'o': 32, 'p': 33,
    'q': 24, 'r': 27, 's': 39, 't': 28, 'u': 30, 'v': 55, 'w': 25, 'x': 53,
    'y': 29, 'z': 52,
    '1': 10, '2': 11, '3': 12, '4': 13, '5': 14, '6': 15, '7': 16, '8': 17,
    '9': 18, '0': 19,
    '-': 20, '=': 21, '[': 34, ']': 35, ';': 47, "'": 48, '`': 49, '\\': 51,
    ',': 59, '.': 60, '/': 61, ' ': 65,
}
_SHIFTED = {
    '!': 10, '@': 11, '#': 12, '$': 13, '%': 14, '^': 15, '&': 16, '*': 17,
    '(': 18, ')': 19, '_': 20, '+': 21, '{': 34, '}': 35, ':': 47, '"': 48,
    '~': 49, '|': 51, '<': 59, '>': 60, '?': 61,
}


def char_to_key(ch):
    if ch in _BASE:
        return _BASE[ch], False
    if ch in _SHIFTED:
        return _SHIFTED[ch], True
    if ch.isalpha() and ch.lower() in _BASE:
        return _BASE[ch.lower()], True  # uppercase letter
    raise ValueError(f"no keycode for {ch!r}")


def find_fifo_pid(timeout=12.0):
    """Find the pid of a running editor by its control FIFO (newest)."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        cands = []
        for p in glob.glob("/tmp/x11viewer-ctl-*"):
            if p.endswith("-resp"):
                continue
            try:
                pid = int(p.rsplit("-", 1)[1])
            except ValueError:
                continue
            if os.path.exists(f"/tmp/x11viewer-ctl-{pid}-resp"):
                cands.append((os.path.getmtime(p), pid))
        if cands:
            return max(cands)[1]
        time.sleep(0.1)
    raise RuntimeError("no editor control FIFO found")


class Editor:
    def __init__(self, exe=None, doc=None, wait=2.5, attach_pid=None):
        if attach_pid is not None:
            self.proc = None
            self.pid = find_fifo_pid() if attach_pid == "auto" else int(attach_pid)
        else:
            args = [exe] + ([doc] if doc else [])
            self.log = open("/tmp/spreadsheet-editor-drive.log", "wb")
            self.proc = subprocess.Popen(
                args, stdout=self.log, stderr=self.log, stdin=subprocess.DEVNULL,
                start_new_session=True, close_fds=True)
            self.pid = self.proc.pid
        self.in_path = f"/tmp/x11viewer-ctl-{self.pid}"
        self.resp_path = f"/tmp/x11viewer-ctl-{self.pid}-resp"
        for _ in range(240):
            if os.path.exists(self.in_path) and os.path.exists(self.resp_path):
                break
            time.sleep(0.05)
        else:
            raise RuntimeError("control FIFOs never appeared")
        time.sleep(wait)  # let the first frame paint
        self._in = open(self.in_path, "wb")  # keep input FIFO open for streaming

    def _cmd(self, cmd_id, payload=b""):
        self._in.write(struct.pack(">II", cmd_id, len(payload)) + payload)
        self._in.flush()
        with open(self.resp_path, "rb") as r:
            rh = r.read(12)
            if len(rh) < 12:
                raise RuntimeError("short response header")
            _, status, n = struct.unpack(">III", rh)
            body = b""
            while len(body) < n:
                chunk = r.read(n - len(body))
                if not chunk:
                    break
                body += chunk
            if status != 0:
                raise RuntimeError(f"cmd {cmd_id} failed: {body.decode('utf-8','replace')}")
            return body

    def key(self, keycode):
        self._cmd(CMD_SEND_KEY, f"{keycode},1".encode())
        self._cmd(CMD_SEND_KEY, f"{keycode},0".encode())

    def press_char(self, ch):
        keycode, shift = char_to_key(ch)
        if shift:
            self._cmd(CMD_SEND_KEY, f"{SHIFT},1".encode())
        self._cmd(CMD_SEND_KEY, f"{keycode},1".encode())
        self._cmd(CMD_SEND_KEY, f"{keycode},0".encode())
        if shift:
            self._cmd(CMD_SEND_KEY, f"{SHIFT},0".encode())

    def type(self, text):
        for ch in text:
            self.press_char(ch)

    def repeat(self, keycode, n):
        for _ in range(n):
            self.key(keycode)

    def goto_a1(self):
        # Clamp to the top-left cell.
        self.key(ESC)
        self.repeat(UP, 40)
        self.repeat(LEFT, 30)

    def goto(self, row, col):
        self.goto_a1()
        self.repeat(DOWN, row)
        self.repeat(RIGHT, col)

    def enter_value(self, text):
        """Type into the current cell and commit with Enter (moves down)."""
        self.type(text)
        self.key(ENTER)

    def click(self, x, y):
        self._cmd(CMD_CLICK_AT, f"{x},{y}".encode())

    def snapshot(self, out):
        body = self._cmd(CMD_SCREENSHOT)
        with open(out, "wb") as f:
            f.write(body)
        return len(body)

    def close(self):
        try:
            self._cmd(CMD_EXIT)
        except Exception:
            pass
        if self.proc is not None:
            try:
                self.proc.wait(timeout=5)
            except Exception:
                self.proc.kill()
            self.log.close()
