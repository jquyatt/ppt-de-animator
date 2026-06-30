# ppt-de-animator

Steps a PowerPoint slideshow through every real click — slide transitions
*and* animation builds — and screenshots each one as a JPG. Fully
deterministic: AppleScript/System Events drive PowerPoint, `screencapture`
grabs frames, a static XML scan decides timing. No AI involved at runtime.

## Setup

1. macOS with PowerPoint installed.
2. Grant the app that will run this script (Terminal, etc.) two permissions
   in System Settings → Privacy & Security:
   - **Screen Recording** — without it, `screencapture` silently fails.
   - **Accessibility** — needed to send right-arrow keystrokes via System
     Events.
3. `pip install -r requirements.txt` (just Pillow, used for cropping).

## Usage

```bash
# Inspect a deck without touching PowerPoint: click count, video slides, etc.
python3 scan.py "deck.pptx"

# Pull embedded video files straight out of the zip
python3 scan.py "deck.pptx" --extract-media videos/

# Run the actual capture
python3 capture.py "deck.pptx" out/
```

`capture.py` writes `out/slide{N}_click{M}.jpg` for every step, in order.
Large decks may need more time to open:

```bash
python3 capture.py "deck.pptx" out/ --load-wait 10
```

## Tests

No fixtures, no PowerPoint required — both use synthetic data:

```bash
python3 test_scan.py
python3 test_capture.py
```

## How it works

**`scan.py`** parses the deck's animation timing XML (a `.pptx` is just a
zip of XML files) to build a manifest before PowerPoint ever opens:

- Counts real click-triggered builds per slide, scoped to the slide's main
  click sequence (a video's own click-to-pause handler lives in a separate
  `interactiveSeq` and must not be counted as an advance step).
- Flags slides with **autoplay video** (`playFrom()` under a timing node
  triggered by slide entry, not a click) and **looping video**
  (`repeatCount="indefinite"`).
- Computes the exact number of clicks needed to reach the end.

**`capture.py`** drives the slideshow with real right-arrow keypresses (not
PowerPoint's `go to next slide` AppleScript command — that one's `current
show position` property doesn't reliably track one click per build, verified
against a deck where it undercounted by more than half). For each step:

- Normally, it polls screenshots until two consecutive captures hash
  identically (the animation has visibly settled) — works regardless of
  effect duration or ordering, no per-effect tuning.
- On a slide flagged autoplay/looping, it skips that wait entirely and just
  grabs a frame after a short fixed delay — those pixels will never stop
  changing, so waiting for "settled" would hang for the full video duration
  (or forever, if looping).
- Detects PowerPoint's end-of-show black card by file size and stops there.

**Letterbox/pillarbox cropping**: `screencapture` grabs the whole display,
but the slideshow content doesn't necessarily fill it — and the bars aren't
guaranteed to be centered (on one test machine a display safe-area pushed
content down: 176px top vs 114px bottom). The crop box is measured once per
run from the first frame and reused for the rest, rather than assumed from
geometry.

## Known gotchas handled

- Stale slideshow range settings from a previously-open PowerPoint window
  can silently override a fresh run — `capture.py` force-closes and resets
  range to "show all" every time.
- `screencapture` refuses to write to dotfile paths (e.g. `.poll.jpg`).
- A deck's true slide order comes from `presentation.xml`'s `sldIdLst`, not
  from `slideN.xml` filename numbers — worth checking if a deck has been
  reordered (not yet handled by `scan.py`; it currently assumes they match).

## Not yet handled

- Auto-advance timers (slide moves on its own after N seconds) could race
  manual click-stepping.
- `loop_until_stopped` decks rely on the manifest's click count rather than
  end-card detection, since the show never reaches one.
- Multi-monitor Presenter View could cause `screencapture` to grab the wrong
  display.
