# Side by side: Word and rbt-office scrolling one document

Recorded on the Windows VM (`winvm`) in its interactive session, at 1920x1080: Word on the left half with its ribbon,
rbt-office's Word shell in an Edge app window on the right; Word's scroll bar thumb is dragged (an eased schedule)
and the page glides between the same two positions over the same time, so the line at a probe height is the same
line on both sides.

    DIST=/path/to/a/dist/snapshot node tools/side-by-side/server.mjs &        # port 16990
    ssh -N -R 16990:127.0.0.1:16990 winvm &                                   # the VM's localhost:16990
    scp tools/side-by-side/drive.ps1 tests/test-documents-showcase/<doc>.docx winvm:C:/sbs/
    tools/side-by-side/vmrun.sh 'powershell -NoProfile -ExecutionPolicy Bypass -File C:\sbs\drive.ps1 -Doc <doc>.docx -Take -TakeMs 12000 -HoldMs 2500 -LagMs -60 -Out take.mkv' 300
    scp winvm:C:/sbs/take.mkv . ; python tools/side-by-side/sync.py take.mkv     # px between the two sides, per frame

Needs `C:\sbs\ffmpeg.exe` on the VM (a BtbN win64 build) and the VM's screen at 1920x1080. The parity tools kill
every WINWORD.EXE between documents: record when no parity run is using the VM.

- Word does not answer automation calls while its thumb is dragged, so positions are read before and after
  (`RangeFromPoint` in the left margin, the picture's range when it hits one) and the drag is rehearsed once
  unrecorded to learn where it ends (it ends in the same place every time).
- Word's page pitch at 100% is 850.96pt (`WORD_PITCH` in page.html; a gap of 9pt), ours 841.9 + 21pt: a position
  maps page by page, a gap onto the gap - so a page boundary passing shows the gaps' difference (about 15px).
- `CAL` (8px) and `-LagMs -60` were measured with sync.py: at rest 0px, moving mostly within 7px.
- The VM manages about 20 frames a second of capture at 1080p (two cores).
