#!/usr/bin/env python3
"""Screenshot the built HTML and put each slide next to the original render.

Usage:
    python compare.py OUT.html WORK_DIR [--callouts]

Writes WORK_DIR/compare/slide-N.png: left = original (LibreOffice render),
right = HTML in its final state (?static: animations done, no motion).
With --callouts every callout is forced visible, to check where they sit.
Also writes WORK_DIR/compare/hover-N.png for slides with callouts: the first
hover target is hovered for real, proving the callout appears on mouseover.
Prints a pixel-difference score per slide (lower is closer; > ~12 deserves a look).
"""
import glob
import os
import sys

from PIL import Image, ImageChops, ImageStat
from playwright.sync_api import sync_playwright


def main(html_path, work, show_callouts=False):
    out = os.path.join(work, "compare")
    os.makedirs(out, exist_ok=True)
    refs = sorted(glob.glob(os.path.join(work, "ref", "slide-*.png")),
                  key=lambda p: int(p.rsplit("-", 1)[1].split(".")[0]))
    url = "file://" + os.path.abspath(html_path) + "?static" + ("&callouts" if show_callouts else "")
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(viewport={"width": 1600, "height": 960}, device_scale_factor=1)
        pg.goto(url)
        pg.wait_for_timeout(1200)
        stages = pg.query_selector_all(".stage")
        ids = pg.eval_on_selector_all(".slide", "els => els.map(e => +e.dataset.index)")
        for idx, st in zip(ids, stages):
            st.scroll_into_view_if_needed()
            pg.wait_for_timeout(150)
            shot = os.path.join(out, "html-%d.png" % idx)
            st.screenshot(path=shot)
            ref = os.path.join(work, "ref", "slide-%d.png" % idx)
            if not os.path.exists(ref):
                continue
            a = Image.open(ref).convert("RGB")
            h = Image.open(shot).convert("RGB").resize(a.size)
            diff = ImageStat.Stat(ImageChops.difference(a.resize((480, int(480 * a.height / a.width))),
                                                        h.resize((480, int(480 * a.height / a.width))))).mean
            score = sum(diff) / 3
            gap = 24
            canvas = Image.new("RGB", (a.width * 2 + gap, a.height), (40, 40, 40))
            canvas.paste(a, (0, 0))
            canvas.paste(h, (a.width + gap, 0))
            canvas.thumbnail((2400, 2400))
            canvas.save(os.path.join(out, "slide-%d.png" % idx))
            print("slide %d: diff %.1f" % (idx, score))
        # real hover check on the motion version
        pg2 = b.new_page(viewport={"width": 1600, "height": 960})
        pg2.goto("file://" + os.path.abspath(html_path))
        pg2.wait_for_timeout(800)
        for idx in ids:
            hot = pg2.query_selector("#slide-%d .hot" % idx)
            if not hot:
                continue
            pg2.query_selector("#slide-%d" % idx).scroll_into_view_if_needed()
            pg2.wait_for_timeout(5800)  # let the entrance timeline finish
            hot.hover()
            pg2.wait_for_timeout(450)
            shown = pg2.eval_on_selector_all("#slide-%d .el.callout.show" % idx, "e => e.length")
            pg2.query_selector("#slide-%d .stage" % idx).screenshot(path=os.path.join(out, "hover-%d.png" % idx))
            print("slide %d: hover shows %d callout(s)" % (idx, shown))
        b.close()
    print("comparison images in", out)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    main(sys.argv[1], sys.argv[2], "--callouts" in sys.argv)
