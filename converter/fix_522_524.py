# -*- coding: utf-8 -*-
"""Ручная доводка 5.22–5.24 после convert.py (то, что конвертер сам не решает)."""
import sys, copy
from pptx import Presentation
from engine import *

IN = EMU
CT, CB = int(1.0 * IN), int(6.8 * IN)


def by_id(s, i):
    for e in s._element.iter():
        c = e.find('.//p:cNvPr', NS) if etree.QName(e).localname in ('sp', 'pic', 'cxnSp', 'grpSp', 'graphicFrame') else None
        if c is not None and c.get('id') == str(i) and e.getparent() is s.shapes._spTree: return e
    raise KeyError(i)


def by_text(s, start):
    for e in list(s.shapes._spTree)[2:]:
        if text_of(e).strip().startswith(start): return e
    raise KeyError(start)


def set_sz(el, sz):
    for rp in el.iter(A + 'rPr', A + 'endParaRPr'): rp.set('sz', str(sz))


def need(el, w=None):
    return estimate_height(el.find('p:txBody', NS), w or bbox(el)[2])


def center_group(els, top=CT, bot=CB):
    bs = [bbox(e) for e in els]
    y0 = min(b[1] for b in bs); y1 = max(b[1] + b[3] for b in bs)
    d = top + (bot - top - (y1 - y0)) / 2 - y0
    for e, b in zip(els, bs):
        b[1] += d; set_bbox(e, b)


def plate(el, fill):
    spPr = el.find('p:spPr', NS)
    for c in list(spPr):
        if etree.QName(c).localname in ('noFill', 'solidFill', 'gradFill'): spPr.remove(c)
    geom = spPr.find('a:prstGeom', NS)
    f = solid(fill)
    geom.addnext(f)
    bp = el.find('p:txBody/a:bodyPr', NS)
    for k in ('lIns', 'tIns', 'rIns', 'bIns'): bp.set(k, '90000')


def fix_22(p):
    # сл. 16: текст справа от фрески вплотную к картинке — зазор 0,15″
    s = p.slides[15]
    t = by_text(s, 'На Крите существовала'); b = bbox(t)
    pic = s._element.find('.//p:pic', NS); pb = bbox(pic)
    nx = pb[0] + pb[2] + int(0.15 * IN); b[2] = b[0] + b[2] - nx; b[0] = nx; b[3] = need(t, b[2]); set_bbox(t, b)
    # сл. 28: две колонки с цветными шапками — текст 16 pt, плашки по высоте текста, шапка сверху
    s = p.slides[27]
    cols = [by_text(s, 'Общество делилось'), by_text(s, 'Хорошо было развито')]
    heads = [e for e in list(s.shapes._spTree)[2:] if etree.QName(e).localname == 'sp' and not text_of(e).strip()
             and e.find('p:spPr/a:solidFill', NS) is not None]
    hb = bbox(heads[0])
    hh = 0
    for c in cols:
        set_sz(c, 1600)
        c.find('p:txBody/a:bodyPr', NS).set('tIns', str(hb[3] + 90000))
        hh = max(hh, need(c))
    for c, h in zip(cols, heads):
        b = bbox(c); b[1] = bbox(h)[1]; b[3] = hh; set_bbox(c, b)
    center_group(cols + heads)


def fix_24(p):
    # сл. 1: перенос «Поэмы Гомера / «Илиада» и «Одиссея»» — по смыслу; кегль 36 оставляет строки левее брелока
    s = p.slides[0]
    t = [e for e in s.shapes._spTree.iter(P + 'sp') if (e.find('.//p:ph', NS) is not None and e.find('.//p:ph', NS).get('type') == 'ctrTitle')][0]
    Builder.set_title(s, 'Поэмы Гомера\n«Илиада» и «Одиссея»')
    set_sz(t, 3600)
    # сл. 6: три статуи одного размера, низ по одной линии; реплики и стрелки едут вместе со статуями
    s = p.slides[5]
    pics = sorted(s._element.findall('.//p:pic', NS), key=lambda e: bbox(e)[0])
    caps = sorted([by_text(s, 'Афина'), by_text(s, 'Афродита'), by_text(s, 'Гера')], key=lambda e: bbox(e)[0])
    talks = sorted([by_text(s, 'Обещаю воинские'), by_text(s, 'Я найду'), by_text(s, 'Обещаю власть')], key=lambda e: bbox(e)[0])
    arrows = sorted(s._element.findall('.//p:cxnSp', NS), key=lambda e: bbox(e)[0])
    H = int(2.85 * IN); bottom = int(5.85 * IN)
    for pic, cap, tk, ar in zip(pics, caps, talks, arrows):
        b = bbox(pic); cx = b[0] + b[2] / 2; ny = bottom - H
        set_bbox(pic, [cx - H / 2, ny, H, H])
        cb = bbox(cap); set_bbox(cap, [cx - H / 2, bottom + int(0.04 * IN), H, cb[3]])
        cap.find('p:txBody/a:bodyPr', NS).set('anchor', 't')
        # стрелка (повёрнута на 90°, видимая высота = cx) — сразу над головой статуи, реплика — над стрелкой
        ab = bbox(ar); vis_h = ab[2]
        d = (ny - int(0.05 * IN) - vis_h / 2) - (ab[1] + ab[3] / 2)
        for e in (tk, ar):
            eb = bbox(e); eb[1] += d; set_bbox(e, eb)
    # сл. 10: кавычки-декор старого шаблона убраны конвертером — цитату ставим на лавандовую плашку,
    # картинку укрупняем справа от последнего абзаца
    s = p.slides[9]
    t1 = by_text(s, 'Троянцы напали'); q = by_text(s, '«Дай позволение'); t3 = by_text(s, 'Ахилл согласился')
    pic = s._element.find('.//p:pic', NS); cap = by_text(s, '«Ахиллес в доспехах»')
    xf = xfrm_of(q); xf.attrib.pop('rot', None)
    plate(q, 'EEE5FF')
    for e in (t1, t3):
        bp = e.find('p:txBody/a:bodyPr', NS); bp.set('lIns', '0'); bp.set('rIns', '0')
    W = RM - LM; g = int(0.2 * IN)
    y = CT
    set_bbox(t1, [LM, y, W, need(t1, W)]); y += bbox(t1)[3] + g
    set_bbox(q, [LM, y, W, need(q, W)]); y += bbox(q)[3] + g
    pb = bbox(pic); ar = pb[2] / pb[3]
    capw = int(2.6 * IN); ph = int(capw / ar)
    cap_h = need(cap, capw)
    avail = CB - y - cap_h - int(0.04 * IN)
    if ph > avail: ph = avail; capw = int(ph * ar)
    px = RM - capw
    set_bbox(pic, [px, y, capw, ph])
    set_bbox(cap, [px, y + ph + int(0.04 * IN), capw, need(cap, capw)])
    tw = px - LM - int(0.25 * IN)
    set_bbox(t3, [LM, y, tw, need(t3, tw)])
    center_group([t1, q, t3, pic, cap])
    # сл. 24: последнее утверждение рефлексии не помещается в строку при 16 pt — все утверждения 14 pt,
    # последняя строка чуть выше
    s = p.slides[23]
    rows = [e for e in list(s.shapes._spTree)[2:] if etree.QName(e).localname == 'sp' and text_of(e).strip().startswith('Я ')]
    for r in rows: set_sz(r, 1400)
    last = max(rows, key=lambda e: bbox(e)[1]); b = bbox(last)
    nh = max(b[3], need(last)); cy = b[1] + b[3] / 2
    set_bbox(last, [b[0], cy - nh / 2, b[2], nh])


if __name__ == '__main__':
    which, src, dst = sys.argv[1:4]
    p = Presentation(src)
    {'22': fix_22, '24': fix_24}[which](p)
    p.save(dst)
