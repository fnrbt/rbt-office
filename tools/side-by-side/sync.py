# Per frame: the vertical shift between Word's page column and ours (row-ink profiles), in screen px.
import sys, subprocess, numpy as np
# usage: python sync.py take.mkv (numpy)
f = sys.argv[1]; W, H = 1920, 1080
raw = subprocess.run(['ffmpeg', '-v', 'error', '-i', f, '-vf', 'fps=10', '-f', 'rawvideo', '-pix_fmt', 'gray', '-'], capture_output=True).stdout
n = len(raw) // (W * H); v = np.frombuffer(raw[:n * W * H], np.uint8).reshape(n, H, W)
y0, y1 = 250, 940
left = 255 - v[:, y0 - 120:y1 + 120, 223:850].astype(float)       # Word's text column region
right = 255 - v[:, y0:y1, 1560:2190 - 1920 + 1560 + 630 - 630].astype(float) if False else 255 - v[:, y0:y1, 1121:1748].astype(float)
for i in range(n):
    a = left[i].mean(1); b = right[i].mean(1)
    if b.std() < 1: print(f'{i/10:5.1f}s  (blank)'); continue
    best = min(range(0, 241), key=lambda s: np.sum((a[s:s + len(b)] - b) ** 2))
    print(f'{i/10:5.1f}s  ours shown {best - 120:+4d}px vs Word')
