#!/usr/bin/env python3
"""Self-check for capture.py's letterbox-bar detection. Run directly:
python3 test_capture.py
Builds synthetic images in memory -- no screenshots or PowerPoint needed.
"""
import tempfile
from pathlib import Path

from PIL import Image

from capture import detect_crop_box


def _save(im):
    f = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    im.save(f.name)
    return Path(f.name)


def test_asymmetric_letterbox():
    # mirrors what was actually measured on a real capture: top bar 176px,
    # bottom bar 114px, on a 3456x2234 frame -- not centered.
    w, h = 3456, 2234
    im = Image.new("RGB", (w, h), (0, 0, 0))
    content = Image.new("RGB", (w, 1944), (200, 150, 50))
    im.paste(content, (0, 176))
    path = _save(im)
    box = detect_crop_box(path)
    assert box is not None
    left, top, right, bottom = box
    assert top == 176, box
    assert bottom == 176 + 1944, box
    assert left == 0 and right == w, box


def test_pillarbox():
    w, h = 2000, 1200
    im = Image.new("RGB", (w, h), (0, 0, 0))
    content = Image.new("RGB", (1500, h), (10, 200, 10))
    im.paste(content, (250, 0))
    path = _save(im)
    box = detect_crop_box(path)
    assert box == (250, 0, 1750, h), box


def test_no_bars_returns_none():
    w, h = 800, 600
    im = Image.new("RGB", (w, h), (123, 45, 200))
    path = _save(im)
    assert detect_crop_box(path) is None


def test_corner_overlay_does_not_fool_detection():
    # a small bright dot in a corner (e.g. screen-recording indicator)
    # shouldn't be mistaken for the start of real content
    w, h = 1000, 800
    im = Image.new("RGB", (w, h), (0, 0, 0))
    content = Image.new("RGB", (w, 600), (200, 200, 200))
    im.paste(content, (0, 100))
    for x in range(w - 5, w):
        for y in range(5):
            im.putpixel((x, y), (255, 0, 0))  # corner artifact, inside the bar
    path = _save(im)
    box = detect_crop_box(path)
    assert box[1] == 100, box


if __name__ == "__main__":
    tests = [v for k, v in list(globals().items()) if k.startswith("test_")]
    for t in tests:
        t()
        print(f"ok  {t.__name__}")
    print(f"{len(tests)} tests passed")
