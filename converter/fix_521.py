# -*- coding: utf-8 -*-
"""Ручная доводка 5.21 после convert.py: длинный заголовок титула — в три строки, как в исходнике."""
import sys
from pptx import Presentation
from engine import *
import convert as C

p = Presentation(sys.argv[1])
s = p.slides[0]
t = C.get_ph(s, 'ctrTitle')
lines = ['Религия и культура', 'Древней Индии', 'и Древнего Китая']
fs = 36
for i, l in enumerate(lines):  # контроль: строки не заходят на лиса с брелоком
    y0 = bbox(t)[1] + i * fs * 1.25 * 12700
    left = C.title_obstacle_left(s, y0, y0 + fs * 1.25 * 12700)
    w = text_width_pt(l, fs, True) * 1.08 * 12700
    print(l, round(w / EMU, 2), 'край', round((bbox(t)[0] + w) / EMU, 2), 'препятствие', round(left / EMU, 2) if left else None)
Builder.set_title(s, '\n'.join(lines))
for rp in t.iter(A + 'rPr', A + 'endParaRPr'): rp.set('sz', str(fs * 100))
p.save(sys.argv[2])
