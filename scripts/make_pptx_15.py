"""Generate marketing-genius-15.pptx (15 slides)
Run: python3 scripts/make_pptx_15.py
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

from pptx import Presentation
from pptx_helpers import W, H
from slides_part1 import build_slides_1_8
from slides_part2_15 import build_slides_9_15

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

build_slides_1_8(prs, total='15')
build_slides_9_15(prs, total='15')

out = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'marketing-genius-15.pptx')
prs.save(out)
print(f'Saved: {out}  ({len(prs.slides)} slides)')
