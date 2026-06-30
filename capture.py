#!/usr/bin/env python3
"""Drives a PowerPoint slideshow click-by-click and screenshots each step as
a JPG. Pure deterministic automation -- AppleScript/System Events for
control, screencapture for frames, scan.py's manifest for timing decisions.
No AI involved at runtime.

Requires (one-time, manual): Screen Recording + Accessibility permission
granted to whatever app runs this script (System Settings > Privacy &
Security). PowerPoint must be installed. Pillow (pip install Pillow) is
used to detect and crop letterbox/pillarbox bars left by screencapture.

Detects embedded video, extracts it to out_dir/videos/ if present, then
captures every click. Usage:
  python3 capture.py "deck.pptx" out_dir/
  python3 capture.py "deck.pptx" out_dir/ --load-wait 10   # large decks
  python3 capture.py "deck.pptx" out_dir/ --max-clicks 40  # loop_until_stopped decks
  python3 capture.py "deck.pptx" out_dir/ --no-video-extract
"""
import argparse
import hashlib
import subprocess
import sys
import time
from pathlib import Path

from PIL import Image, ImageDraw

from scan import analyze_deck, extract_media

POLL_INTERVAL = 0.25
STABLE_READS = 2
SETTLE_TIMEOUT = 15  # seconds; safety cap per click if nothing stabilizes
END_CARD_MAX_BYTES = 400_000  # PowerPoint's black "End of slide show" card
VIDEO_GRACE_SECONDS = 1.5  # fixed wait on autoplay/looping video slides
RIGHT_ARROW = 124

# ponytail: small fixed-pixel chrome zones, not a percentage of frame height.
# A percentage margin (the original fix) ate into real slide content -- a
# bullet's fly-in-from-below animation passes through the bottom of the
# frame, so a broad bottom strip hid its motion from the stability check and
# the frame got accepted mid-flight. These mask only the menu bar (a thin
# full-width strip) and the popup toolbar (a small bottom-left cluster).
# Bump these if a future machine's chrome turns out taller than this.
MENU_BAR_MARGIN_PX = 50
TOOLBAR_CORNER_W_PX = 220
TOOLBAR_CORNER_H_PX = 60


def osa(script):
    subprocess.run(["osascript", "-e", script], capture_output=True, text=True)


def osa_out(script):
    r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    return r.stdout.strip()


def screenshot(path):
    subprocess.run(["screencapture", "-x", str(path)], check=True)


def content_signature(path):
    """Hashes the frame with two small UI-chrome zones blanked out: the
    system menu bar (a thin strip across the top) and PowerPoint's popup
    toolbar (a small cluster in the bottom-left corner). Both can flicker
    (a ticking clock, a fading toolbar, a recording indicator) independent
    of whether the slide itself has finished animating -- without masking
    them, settle-detection never finds two identical frames and burns the
    full SETTLE_TIMEOUT on every single click even on a static slide.

    Masks (paints black) rather than crops, and keeps the zones small and
    corner/edge-specific, so real slide content -- including animations
    that legitimately pass through the bottom of the frame, like a
    fly-in-from-below bullet -- stays visible to the comparison."""
    im = Image.open(path).convert("RGB")
    w, h = im.size
    masked = im.copy()
    draw = ImageDraw.Draw(masked)
    draw.rectangle([0, 0, w, min(MENU_BAR_MARGIN_PX, h)], fill=(0, 0, 0))
    draw.rectangle(
        [0, max(0, h - TOOLBAR_CORNER_H_PX), min(TOOLBAR_CORNER_W_PX, w), h],
        fill=(0, 0, 0),
    )
    return hashlib.sha256(masked.tobytes()).hexdigest()


def detect_crop_box(image_path, black_thresh=10):
    """Finds the real content rectangle inside a screenshot, stripping
    letterbox/pillarbox bars. screencapture grabs the whole display, but
    PowerPoint's slideshow window doesn't necessarily fit it exactly (its
    aspect ratio differs, and on this Mac the content wasn't even centered
    -- a camera-housing safe area pushed it down) so the bars can be
    asymmetric. Detected once empirically per run rather than assumed from
    geometry, then reused for every frame. Returns a PIL crop box, or None
    if no bars are present."""
    im = Image.open(image_path)
    w, h = im.size
    px = im.load()

    # Sample only the central band of each edge so a corner overlay (e.g. a
    # screen-recording indicator dot) can't fool the scan into thinking a
    # bar has ended early.
    xs = range(int(w * 0.3), int(w * 0.7), max(1, int(w * 0.01)))
    ys = range(int(h * 0.3), int(h * 0.7), max(1, int(h * 0.01)))

    def is_black_row(y):
        return all(max(px[x, y][:3]) <= black_thresh for x in xs)

    def is_black_col(x):
        return all(max(px[x, y][:3]) <= black_thresh for y in ys)

    top = 0
    while top < h and is_black_row(top):
        top += 1
    bottom = h - 1
    while bottom > top and is_black_row(bottom):
        bottom -= 1
    left = 0
    while left < w and is_black_col(left):
        left += 1
    right = w - 1
    while right > left and is_black_col(right):
        right -= 1

    box = (left, top, right + 1, bottom + 1)
    if box == (0, 0, w, h):
        return None
    return box


def crop_in_place(image_path, box):
    im = Image.open(image_path).convert("RGB")
    im.crop(box).save(image_path, quality=95)


def press_right_arrow():
    osa(f"""
    tell application "System Events"
        tell process "Microsoft PowerPoint"
            key code {RIGHT_ARROW}
        end tell
    end tell
    """)


def open_and_start_slideshow(deck_path, load_wait=4):
    # Closes any already-open window for this deck first: a stale open
    # instance can carry forward leftover slide show range settings (e.g.
    # from prior manual testing) that silently override a fresh run.
    osa(f"""
    tell application "Microsoft PowerPoint"
        try
            close (every document whose name is "{deck_path.name}") saving no
        end try
    end tell
    """)
    subprocess.run(["open", "-a", "Microsoft PowerPoint", str(deck_path)])
    time.sleep(load_wait)
    osa("""
    tell application "Microsoft PowerPoint"
        activate
        try
            exit slide show (slide show view of slide show window 1)
        end try
        set p to active presentation
        set sss to slide show settings of p
        set range type of sss to slide show range show all
        run slide show sss
    end tell
    """)
    # PowerPoint's popup toolbar (if enabled in Preferences > Slide Show) can
    # be visible right when the show starts and takes a few seconds of no
    # mouse movement to fade -- this script never moves the mouse, so just
    # wait it out rather than risk baking it into the first frame.
    time.sleep(4)


def exit_slideshow():
    osa("""
    tell application "Microsoft PowerPoint"
        try
            exit slide show (slide show view of slide show window 1)
        end try
    end tell
    """)


def wait_for_frame(tmp_path, skip_settle):
    """Captures into tmp_path. If skip_settle, just waits a fixed grace
    period (pixels on an autoplay/looping video slide will never stabilize).
    Otherwise polls until two consecutive captures match on content_signature
    (which ignores UI chrome) -- comparing the full frame would let a
    flickering menu-bar clock or fading toolbar block stability forever."""
    if skip_settle:
        time.sleep(VIDEO_GRACE_SECONDS)
        screenshot(tmp_path)
        return

    prev_hash, stable, elapsed = None, 0, 0.0
    while True:
        screenshot(tmp_path)
        h = content_signature(tmp_path)
        if h == prev_hash:
            stable += 1
            if stable >= STABLE_READS:
                return
        else:
            stable = 0
        prev_hash = h
        time.sleep(POLL_INTERVAL)
        elapsed += POLL_INTERVAL
        if elapsed > SETTLE_TIMEOUT:
            return


def run_capture(deck_path, out_dir, max_clicks=None, load_wait=4, manifest=None):
    deck_path = Path(deck_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    tmp_path = out_dir / "_poll.jpg"

    if manifest is None:
        manifest = analyze_deck(deck_path)
    steps = manifest["steps"]
    expected_clicks = manifest["total_clicks_to_end"]
    print(
        f"manifest: {manifest['slide_count']} slides, "
        f"{expected_clicks} clicks to end, "
        f"loop_until_stopped={manifest['loop_until_stopped']}"
    )

    open_and_start_slideshow(deck_path, load_wait=load_wait)

    crop_box = None
    saved = []
    for step in steps:
        if max_clicks is not None and step["click"] > max_clicks:
            break
        if step["click"] > 0:
            press_right_arrow()
            time.sleep(0.15)

        wait_for_frame(tmp_path, step["skip_settle"])
        size = tmp_path.stat().st_size
        is_end_card = size < END_CARD_MAX_BYTES

        if is_end_card:
            print(f"click {step['click']}: hit end-of-show card, stopping")
            break

        fname = out_dir / f"slide{step['slide']:03d}_click{step['click']:03d}.jpg"
        tmp_path.replace(fname)

        if crop_box is None:
            crop_box = detect_crop_box(fname) or False
            if crop_box:
                print(f"detected letterbox bars, cropping to {crop_box}")
        if crop_box:
            crop_in_place(fname, crop_box)

        saved.append(fname)
        print(f"click {step['click']} (slide {step['slide']}) -> {fname.name} ({size} bytes)")

    exit_slideshow()
    tmp_path.unlink(missing_ok=True)

    if len(saved) != expected_clicks + 1 and not manifest["loop_until_stopped"]:
        print(
            f"WARNING: captured {len(saved)} frames, manifest expected "
            f"{expected_clicks + 1} (clicks 0..{expected_clicks}). "
            "Deck structure may not match the manifest's assumptions.",
            file=sys.stderr,
        )

    print(f"done: {len(saved)} frames in {out_dir}")
    return saved


def process_deck(deck_path, out_dir, video_dir=None, extract_video=True, max_clicks=None, load_wait=4):
    """Full pipeline: scan the deck once, pull out any embedded video files
    it has, then capture every click. Scanning first (rather than just
    trying to extract blindly) means we only touch the filesystem for video
    when there's actually video to find, and the same manifest drives both
    the extraction decision and the capture run -- no redundant re-scan."""
    deck_path = Path(deck_path)
    out_dir = Path(out_dir)

    manifest = analyze_deck(deck_path)
    has_video = any(s["autoplay_video"] or s["looping_video"] for s in manifest["slides"])

    extracted = []
    if has_video and extract_video:
        video_dir = Path(video_dir) if video_dir else out_dir / "videos"
        extracted = extract_media(deck_path, video_dir)
        print(f"found video on {sum(1 for s in manifest['slides'] if s['autoplay_video'] or s['looping_video'])} "
              f"slide(s), extracted {len(extracted)} file(s) to {video_dir}")
    elif has_video:
        print("video detected but --no-video-extract was set, skipping extraction")
    else:
        print("no video detected")

    saved = run_capture(deck_path, out_dir, max_clicks=max_clicks, load_wait=load_wait, manifest=manifest)
    return {"extracted_media": extracted, "frames": saved}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck", help="path to .pptx")
    ap.add_argument("out_dir", help="directory to write JPGs into")
    ap.add_argument("--max-clicks", type=int, default=None, help="safety cap, mainly for loop_until_stopped decks")
    ap.add_argument("--load-wait", type=int, default=4, help="seconds to wait for PowerPoint to open/render large decks")
    ap.add_argument("--video-dir", default=None, help="where to extract embedded video (default: out_dir/videos)")
    ap.add_argument("--no-video-extract", action="store_true", help="skip extracting embedded video even if present")
    args = ap.parse_args()
    process_deck(
        args.deck,
        args.out_dir,
        video_dir=args.video_dir,
        extract_video=not args.no_video_extract,
        max_clicks=args.max_clicks,
        load_wait=args.load_wait,
    )


if __name__ == "__main__":
    main()
