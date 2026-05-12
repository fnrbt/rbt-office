#!/usr/bin/env python3
"""Talk to the word-editor ControlServer via its FIFO pair.

Usage:
    control.py <pid> click <x> <y>
    control.py <pid> key <keycode> <press|release>
    control.py <pid> snapshot <out.png>
    control.py <pid> exit
"""
import os
import struct
import sys
import time


CMD_SCREENSHOT = 1
CMD_EXIT = 2
CMD_CLICK_AT = 11
CMD_SEND_KEY = 14


def send(pid: int, cmd_id: int, payload: bytes = b"") -> bytes:
    in_path = f"/tmp/x11viewer-ctl-{pid}"
    resp_path = f"/tmp/x11viewer-ctl-{pid}-resp"

    # Wait briefly for FIFOs to exist if the server just started.
    for _ in range(50):
        if os.path.exists(in_path) and os.path.exists(resp_path):
            break
        time.sleep(0.05)

    header = struct.pack(">II", cmd_id, len(payload))

    # Open response FIFO read-only first (non-blocking) so the server's write doesn't stall.
    # Actually we need to open it AFTER sending so the server can be ready. Use blocking open.
    with open(in_path, "wb") as f:
        f.write(header + payload)
        f.flush()

    with open(resp_path, "rb") as r:
        rh = r.read(12)
        if len(rh) < 12:
            raise RuntimeError(f"short response header: {len(rh)} bytes")
        resp_cmd, status, payload_len = struct.unpack(">III", rh)
        body = b""
        while len(body) < payload_len:
            chunk = r.read(payload_len - len(body))
            if not chunk:
                break
            body += chunk
        if status != 0:
            raise RuntimeError(f"command {cmd_id} failed: {body.decode('utf-8', 'replace')}")
        return body


def main() -> int:
    args = sys.argv[1:]
    if len(args) < 2:
        print(__doc__, file=sys.stderr)
        return 2

    pid = int(args[0])
    op = args[1]

    if op == "click":
        x, y = float(args[2]), float(args[3])
        send(pid, CMD_CLICK_AT, f"{x},{y}".encode())
    elif op == "key":
        keycode = int(args[2])
        pressed = 1 if args[3] == "press" else 0
        send(pid, CMD_SEND_KEY, f"{keycode},{pressed}".encode())
    elif op == "snapshot":
        out = args[2]
        body = send(pid, CMD_SCREENSHOT)
        with open(out, "wb") as f:
            f.write(body)
        print(f"wrote {out} ({len(body)} bytes)")
    elif op == "exit":
        try:
            send(pid, CMD_EXIT)
        except Exception:
            pass
    else:
        print(f"unknown op: {op}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
