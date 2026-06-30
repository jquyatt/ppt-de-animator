#!/usr/bin/env python3
"""Static analysis of a .pptx deck: builds a click-by-click capture manifest
and can extract embedded video files. Pure stdlib (zipfile + ElementTree),
no PowerPoint and no AI involved -- a .pptx is just a zip of XML.

Usage:
  python3 scan.py deck.pptx                     # print manifest as JSON
  python3 scan.py deck.pptx --extract-media DIR  # also dump embedded videos
"""
import argparse
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
VIDEO_EXT_RE = re.compile(r"\.(mp4|mov|m4v|wmv|avi)$", re.I)


def _slide_files(z):
    names = [n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)]
    return sorted(names, key=lambda n: int(re.search(r"\d+", n).group()))


def _collect(el, ancestors, cmds, media_nodes):
    """Recursively walk a timing tree, tracking the cTn ancestor chain so we
    can tell, for each p:cmd / p:cMediaNode found, what triggers it."""
    tag = el.tag.split("}")[-1]
    if tag == "cTn":
        ancestors = ancestors + [el]
    if tag == "cmd":
        cmds.append((el, ancestors))
    elif tag == "cMediaNode":
        media_nodes.append((el, ancestors))
    for child in el:
        _collect(child, ancestors, cmds, media_nodes)


def analyze_slide(xml_bytes):
    """Returns click_effects (builds triggered by the user advancing), plus
    autoplay_video / looping_video flags derived from the timing XML.

    autoplay_video: a video's playFrom() call sits under a nodeType=
      "afterEffect" node (i.e. it starts on slide entry, not on a click).
    looping_video: that video's media node has repeatCount="indefinite".
    """
    root = ET.fromstring(xml_bytes)
    timing = root.find(".//p:timing", NS)
    if timing is None:
        return {"click_effects": 0, "autoplay_video": False, "looping_video": False}

    main_seq = timing.find('.//p:cTn[@nodeType="mainSeq"]', NS)
    click_effects = 0
    if main_seq is not None:
        click_effects = len(main_seq.findall('.//p:cTn[@nodeType="clickEffect"]', NS))

    cmds, media_nodes = [], []
    _collect(timing, [], cmds, media_nodes)

    autoplay = False
    for el, ancestors in cmds:
        if not el.get("cmd", "").startswith("playFrom"):
            continue
        for a in reversed(ancestors):
            node_type = a.get("nodeType")
            if node_type:
                if node_type == "afterEffect":
                    autoplay = True
                break

    looping = False
    for el, _ in media_nodes:
        cTn = el.find("p:cTn", NS)
        if cTn is not None and cTn.get("repeatCount") == "indefinite":
            looping = True

    return {"click_effects": click_effects, "autoplay_video": autoplay, "looping_video": looping}


def _build_steps(slides):
    """Flattens per-slide info into one entry per real click, click 0 being
    the initial state before any keypress. skip_settle is set whenever the
    current slide has an autoplay/looping video, since on-screen pixels will
    keep changing for as long as we're on that slide."""
    steps = [{"click": 0, "slide": 1, "skip_settle": slides[0]["autoplay_video"] or slides[0]["looping_video"]}]
    click = 0
    for i, s in enumerate(slides, start=1):
        skip = s["autoplay_video"] or s["looping_video"]
        n_steps = (1 if i > 1 else 0) + s["click_effects"]
        for _ in range(n_steps):
            click += 1
            steps.append({"click": click, "slide": i, "skip_settle": skip})
    return steps


def analyze_deck(path):
    z = zipfile.ZipFile(path)
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    show_pr = pres.find("p:showPr", NS)
    loop_until_stopped = show_pr is not None and show_pr.get("loop") == "1"

    slides = [analyze_slide(z.read(sf)) for sf in _slide_files(z)]
    total_clicks = (len(slides) - 1) + sum(s["click_effects"] for s in slides)

    return {
        "slide_count": len(slides),
        "loop_until_stopped": loop_until_stopped,
        "total_clicks_to_end": total_clicks,
        "slides": slides,
        "steps": _build_steps(slides),
    }


def extract_media(path, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    z = zipfile.ZipFile(path)
    extracted = []
    for n in z.namelist():
        if n.startswith("ppt/media/") and VIDEO_EXT_RE.search(n):
            dest = os.path.join(out_dir, os.path.basename(n))
            with open(dest, "wb") as f:
                f.write(z.read(n))
            extracted.append(dest)
    return extracted


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("deck", help="path to .pptx")
    ap.add_argument("--extract-media", metavar="DIR", help="dump embedded video files here")
    args = ap.parse_args()

    manifest = analyze_deck(args.deck)
    if args.extract_media:
        manifest["extracted_media"] = extract_media(args.deck, args.extract_media)
    print(json.dumps(manifest, indent=2))


if __name__ == "__main__":
    main()
