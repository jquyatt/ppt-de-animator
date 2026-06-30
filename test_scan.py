#!/usr/bin/env python3
"""Self-check for scan.py's timing-XML analysis. Run directly: python3 test_scan.py
Uses synthetic XML matching the real structures verified against actual decks,
so it doesn't depend on any local .pptx file being present.
"""
from scan import analyze_slide

P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'


def slide(timing_inner):
    return f'<p:sld {P}><p:timing>{timing_inner}</p:timing></p:sld>'.encode()


def test_plain_builds():
    xml = slide("""
    <p:tnLst><p:par><p:cTn nodeType="tmRoot"><p:childTnLst><p:seq>
      <p:cTn nodeType="mainSeq"><p:childTnLst>
        <p:par><p:cTn nodeType="clickEffect"/></p:par>
        <p:par><p:cTn nodeType="clickEffect"/></p:par>
        <p:par><p:cTn nodeType="clickEffect"/></p:par>
      </p:childTnLst></p:cTn>
    </p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>
    """)
    r = analyze_slide(xml)
    assert r == {"click_effects": 3, "autoplay_video": False, "looping_video": False}, r


def test_autoplay_video_no_loop():
    xml = slide("""
    <p:tnLst><p:par><p:cTn nodeType="tmRoot"><p:childTnLst><p:seq>
      <p:cTn nodeType="mainSeq"><p:childTnLst/></p:cTn>
    </p:seq>
    <p:cTn nodeType="afterEffect"><p:childTnLst>
      <p:cmd cmd="playFrom(0.0)"><p:cBhvr><p:cTn dur="8512"/></p:cBhvr></p:cmd>
    </p:childTnLst></p:cTn>
    <p:video><p:cMediaNode><p:cTn id="7"/></p:cMediaNode></p:video>
    </p:childTnLst></p:cTn></p:par></p:tnLst>
    """)
    r = analyze_slide(xml)
    assert r == {"click_effects": 0, "autoplay_video": True, "looping_video": False}, r


def test_autoplay_video_looping():
    xml = slide("""
    <p:tnLst><p:par><p:cTn nodeType="tmRoot"><p:childTnLst><p:seq>
      <p:cTn nodeType="mainSeq"><p:childTnLst/></p:cTn>
    </p:seq>
    <p:cTn nodeType="afterEffect"><p:childTnLst>
      <p:cmd cmd="playFrom(0.0)"><p:cBhvr><p:cTn dur="46671"/></p:cBhvr></p:cmd>
    </p:childTnLst></p:cTn>
    <p:video><p:cMediaNode><p:cTn id="7" repeatCount="indefinite"/></p:cMediaNode></p:video>
    </p:childTnLst></p:cTn></p:par></p:tnLst>
    """)
    r = analyze_slide(xml)
    assert r == {"click_effects": 0, "autoplay_video": True, "looping_video": True}, r


def test_interactive_click_not_counted_as_build():
    # A click-to-pause handler on the video shape itself lives in an
    # interactiveSeq, not mainSeq -- it must NOT inflate click_effects.
    xml = slide("""
    <p:tnLst><p:par><p:cTn nodeType="tmRoot"><p:childTnLst>
      <p:seq><p:cTn nodeType="mainSeq"><p:childTnLst>
        <p:par><p:cTn nodeType="clickEffect"/></p:par>
      </p:childTnLst></p:cTn></p:seq>
      <p:seq><p:cTn nodeType="interactiveSeq"><p:childTnLst>
        <p:par><p:cTn nodeType="clickEffect">
          <p:childTnLst><p:cmd cmd="togglePause"/></p:childTnLst>
        </p:cTn></p:par>
      </p:childTnLst></p:cTn></p:seq>
    </p:childTnLst></p:cTn></p:par></p:tnLst>
    """)
    r = analyze_slide(xml)
    assert r["click_effects"] == 1, r


if __name__ == "__main__":
    tests = [v for k, v in list(globals().items()) if k.startswith("test_")]
    for t in tests:
        t()
        print(f"ok  {t.__name__}")
    print(f"{len(tests)} tests passed")
