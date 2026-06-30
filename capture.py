#!/usr/bin/env python3
"""Drives a PowerPoint slideshow click-by-click and screenshots each step as
a JPG. Pure deterministic automation -- AppleScript/System Events for
control, screencapture for frames, scan.py's manifest for timing decisions.
No AI involved at runtime.

Requires (one-time, manual): Screen Recording + Accessibility permission
granted to whatever app runs this script (System Settings > Privacy &
Security). PowerPoint must be installed.

Usage: python3 capture.py "deck.pptx" out_dir/
"""
import argparse
import hashlib
import subprocess
import sys
import time
from pathlib import Path

from scan import analyze_deck

POLL_INTERVAL = 0.25
STABLE_READS = 2
SETTLE_TIMEOUT = 15  # seconds; safety cap per click if nothing stabilizes
END_CARD_MAX_BYTES = 400_000  # PowerPoint's black "End of slide show" card
VIDEO_GRACE_SECONDS = 1.5  # fixed wait on autoplay/looping video slides
RIGHT_ARROW = 124


def osa(script):
    subprocess.run(["osascript", "-e", script], capture_output=True, text=True)


def osa_out(script):
    r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    return r.stdout.strip()


def screenshot(path):
    subprocess.run(["screencapture", "-x", str(path)], check=True)


def file_hash(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


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
    time.sleep(2)


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
    Otherwise polls until two consecutive captures hash identically."""
    if skip_settle:
        time.sleep(VIDEO_GRACE_SECONDS)
        screenshot(tmp_path)
        return

    prev_hash, stable, elapsed = None, 0, 0.0
    while True:
        screenshot(tmp_path)
        h = file_hash(tmp_path)
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


def run_capture(deck_path, out_dir, max_clicks=None, load_wait=4):
    deck_path = Path(deck_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    tmp_path = out_dir / "_poll.jpg"

    manifest = analyze_deck(deck_path)
    steps = manifest["steps"]
    expected_clicks = manifest["total_clicks_to_end"]
    print(
        f"manifest: {manifest['slide_count']} slides, "
        f"{expected_clicks} clicks to end, "
        f"loop_until_stopped={manifest['loop_until_stopped']}"
    )

    open_and_start_slideshow(deck_path, load_wait=load_wait)

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


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("deck", help="path to .pptx")
    ap.add_argument("out_dir", help="directory to write JPGs into")
    ap.add_argument("--max-clicks", type=int, default=None, help="safety cap, mainly for loop_until_stopped decks")
    ap.add_argument("--load-wait", type=int, default=4, help="seconds to wait for PowerPoint to open/render large decks")
    args = ap.parse_args()
    run_capture(args.deck, args.out_dir, args.max_clicks, args.load_wait)


if __name__ == "__main__":
    main()
