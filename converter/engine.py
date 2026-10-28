# -*- coding: utf-8 -*-
"""Перенос презентаций старого шаблона Фоксфорда (7.875x5.625in) в новый (10x7.5in, база — пример 5.8)."""
import copy, hashlib, io, re, os, math
from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from PIL import ImageFont
from common import NS, A, P, R, scheme_map, color_of

EMU = 914400
SRC_W, SRC_H = 7560000, 5400000
TW, TH = 9144000, 6858000
K = TW / SRC_W
LM, RM = 360000, TW - 360000            # поля 0.39in
CT, CB = int(0.95 * EMU), int(6.80 * EMU)  # верх/низ контентной зоны
import os as _os
WORK = _os.environ.get('CONV_DIR', '/home/claude/w')
BASE = _os.environ.get('CONV_BASE', WORK + '/ex58.pptx')
ASSETS = WORK + '/assets'

DECOR_SKIP = {'faa3488e67', '05467b9d2c', '34231935b5', '0b85a8f9f9', 'e082501d22',
              # 6 класс: лиса, уголки, полоски и зигзаги старого шаблона
              'ef881efe20', 'ffb86e73c5', '06e1d28562', 'c499013524', '503a35b075', '07e6dab9b5', '8323e22b85',
              # 5 класс: градиентные кавычки «« старого шаблона (5.24 сл. 10)
              '822f63fbc1'}
DECOR_VAZHNO = {'4438fd9d35'}
DECOR_Q = {'97b5376175'}

# ---------- цвета ----------
FILL_MAP = {'AC6CFF': '641FF1', 'B26CF7': '641FF1', 'B963FF': '641FF1', 'FF9900': 'FF9900', 'FE9400': 'FF9900',
            'FFFFFF': 'FFFFFF', 'FBE7FF': 'EEE5FF', 'FFE7FD': 'EEE5FF', 'E0F7FF': 'CCEAFF', 'AAE0F6': 'CCEAFF',
            'EAFF6B': 'FFFED6', 'FEEAC7': 'FFEED7'}


def hsv(h):
    r, g, b = [int(h[i:i + 2], 16) / 255 for i in (0, 2, 4)]
    mx, mn = max(r, g, b), min(r, g, b)
    return mx, mx - mn, (r, g, b)


def map_fill(h):
    if h in FILL_MAP: return FILL_MAP[h]
    v, c, (r, g, b) = hsv(h)
    if c < 0.08: return h  # серые/белые/чёрные
    if v > 0.85 and c < 0.35:  # пастель
        if b >= r and b >= g: return 'CCEAFF' if g > r else 'EEE5FF'
        if r > b and g > b: return 'FFFED6' if g > 0.9 else 'FFEED7'
        return 'EEE5FF'
    if r > 0.8 and g > 0.4 and b < 0.3: return 'FF9900'
    return '641FF1'


def map_line(h):
    return map_fill(h)


def map_text(h):
    v, c, (r, g, b) = hsv(h)
    if v > 0.93 and c < 0.1: return 'FFFFFF'
    if c > 0.25: return '641FF1'
    return '0F1115'


def is_dark(h):
    r, g, b = [int(h[i:i + 2], 16) for i in (0, 2, 4)]
    return (0.299 * r + 0.587 * g + 0.114 * b) < 150


def solid(hexv):
    e = etree.SubElement(etree.Element('x'), A + 'solidFill')
    c = etree.SubElement(e, A + 'srgbClr'); c.set('val', hexv)
    return e


# ---------- шрифты / оценка высоты текста ----------
FONT_R = '/usr/share/fonts/truetype/open-sans/OpenSans-Regular.ttf'
FONT_B = '/usr/share/fonts/truetype/open-sans/OpenSans-Bold.ttf'
_fc = {}


def font(sz, b):
    k = (sz, b)
    if k not in _fc: _fc[k] = ImageFont.truetype(FONT_B if b else FONT_R, int(sz * 20))
    return _fc[k]


def text_width_pt(txt, sz, b):
    return font(sz, b).getlength(txt) / 20.0


def estimate_height(txBody, width_emu):
    """Высота текста (EMU) с учётом переносов."""
    bp = txBody.find('a:bodyPr', NS)
    li = int(bp.get('lIns', 91440)); ri = int(bp.get('rIns', 91440))
    ti = int(bp.get('tIns', 45720)); bi = int(bp.get('bIns', 45720))
    total = 0.0
    for p in txBody.findall('a:p', NS):
        pPr = p.find('a:pPr', NS)
        marL = int(pPr.get('marL', 0)) if pPr is not None else 0
        avail = (width_emu - li - ri - marL) / 12700.0  # pt
        ln = 1.0
        sb = sa = 0
        if pPr is not None:
            e = pPr.find('a:lnSpc/a:spcPct', NS)
            if e is not None: ln = int(e.get('val')) / 100000
            e = pPr.find('a:spcBef/a:spcPts', NS)
            if e is not None: sb = int(e.get('val')) / 100
            e = pPr.find('a:spcAft/a:spcPts', NS)
            if e is not None: sa = int(e.get('val')) / 100
        runs = []
        maxsz = 0
        for r in p:
            t = etree.QName(r).localname
            if t in ('r', 'fld'):
                rp = r.find('a:rPr', NS)
                sz = int(rp.get('sz', 1400)) / 100 if rp is not None else 14
                b = rp is not None and rp.get('b') == '1'
                runs.append((r.findtext('a:t', '', NS), sz, b))
                maxsz = max(maxsz, sz)
            elif t == 'br':
                runs.append(('\n', 0, False))
        if maxsz == 0:
            e = p.find('a:endParaRPr', NS)
            maxsz = int(e.get('sz', 1400)) / 100 if e is not None and e.get('sz') else 14
        # переносы по словам
        lines = 1; x = 0.0
        for txt, sz, b in runs:
            if txt == '\n': lines += 1; x = 0; continue
            for w in re.split(r'(\s+)', txt):
                if not w: continue
                wl = text_width_pt(w, sz or 14, b)
                if x + wl > avail and x > 0 and not w.isspace():
                    lines += 1; x = wl
                else:
                    x += wl
        total += lines * maxsz * 1.362 * ln + sb + sa
    return int(total * 12700 + ti + bi)


# ---------- наследование стилей плейсхолдеров ----------
PPR_GROUPS = [('lnSpc',), ('spcBef',), ('spcAft',), ('buClrTx', 'buClr'), ('buSzTx', 'buSzPct', 'buSzPts'),
              ('buFontTx', 'buFont'), ('buNone', 'buAutoNum', 'buChar', 'buBlip'), ('tabLst',), ('defRPr',)]
RPR_GROUPS = [('ln',), ('noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'), ('highlight',),
              ('latin',), ('ea',), ('cs',), ('sym',), ('hlinkClick',)]


def merge_el(chain, groups, tag):
    """chain: элементы от младшего приоритета к старшему. Возвращает новый элемент."""
    out = etree.Element(tag)
    attrs = {}
    kids = {}
    for e in chain:
        if e is None: continue
        attrs.update(e.attrib)
        for c in e:
            ln = etree.QName(c).localname
            for gi, g in enumerate(groups):
                if ln in g: kids[gi] = c
    for k, v in attrs.items(): out.set(k, v)
    for gi in sorted(kids): out.append(copy.deepcopy(kids[gi]))
    return out


def lvl_of(lst, lvl):
    if lst is None: return None
    return lst.find('a:lvl%dpPr' % (lvl + 1), NS)


def find_layout_ph(slide, ph):
    typ = ph.get('type', 'body'); idx = ph.get('idx')
    lay = slide.slide_layout
    cands = []
    for src in (lay, lay.slide_master):
        found = None
        for sp in src._element.iter(P + 'sp'):
            p2 = sp.find('p:nvSpPr/p:nvPr/p:ph', NS)
            if p2 is None: continue
            t2 = p2.get('type', 'body')
            if src is lay and idx is not None and p2.get('idx') == idx: found = sp; break
        if found is None:
            for sp in src._element.iter(P + 'sp'):
                p2 = sp.find('p:nvSpPr/p:nvPr/p:ph', NS)
                if p2 is None: continue
                t2 = p2.get('type', 'body')
                norm = lambda t: {'ctrTitle': 'title', 'subTitle': 'body'}.get(t, t)
                if norm(t2) == norm(typ): found = sp; break
        cands.append(found)
    return cands  # [layout_sp, master_sp]


def style_chain(slide, sp):
    """Список (lstStyle-источников) от младшего к старшему для заданной фигуры + bodyPr-цепочка + spPr-цепочка."""
    ph = sp.find('p:nvSpPr/p:nvPr/p:ph', NS)
    master = slide.slide_layout.slide_master._element
    lsts, bps, sps = [], [], []
    if ph is not None:
        typ = ph.get('type', 'body')
        tx = master.find('p:txStyles', NS)
        if typ in ('title', 'ctrTitle'): lsts.append(tx.find('p:titleStyle', NS))
        elif typ in ('body', 'subTitle', 'obj'): lsts.append(tx.find('p:bodyStyle', NS))
        else: lsts.append(tx.find('p:otherStyle', NS))
        lay_sp, mas_sp = find_layout_ph(slide, ph)
        for s in (mas_sp, lay_sp):
            if s is not None:
                lsts.append(s.find('p:txBody/a:lstStyle', NS))
                bps.append(s.find('p:txBody/a:bodyPr', NS))
                sps.append(s.find('p:spPr', NS))
    else:
        dts = slide.part.package.presentation_part._element.find('p:defaultTextStyle', NS)
        lsts.append(dts)
    lsts.append(sp.find('p:txBody/a:lstStyle', NS))
    return lsts, bps, sps


# ---------- утилиты XML ----------
def xfrm_of(el):
    tag = etree.QName(el).localname
    if tag == 'grpSp': return el.find('p:grpSpPr/a:xfrm', NS)
    if tag == 'graphicFrame': return el.find('p:xfrm', NS)
    return el.find('p:spPr/a:xfrm', NS)


def bbox(el):
    x = xfrm_of(el)
    if x is None: return None
    o = x.find('a:off', NS); e = x.find('a:ext', NS)
    if o is None or e is None: return None
    return [int(o.get('x')), int(o.get('y')), int(e.get('cx')), int(e.get('cy'))]


def set_bbox(el, b):
    x = xfrm_of(el)
    o = x.find('a:off', NS); e = x.find('a:ext', NS)
    o.set('x', str(int(b[0]))); o.set('y', str(int(b[1]))); e.set('cx', str(max(0, int(b[2])))); e.set('cy', str(max(0, int(b[3]))))


def text_of(el):
    return ''.join(t.text or '' for t in el.iter(A + 't'))


def img_hash(slide, pic):
    b = pic.find('.//a:blip', NS)
    if b is None: return None
    rid = b.get(R + 'embed')
    if not rid: return None
    return hashlib.sha1(slide.part.related_part(rid).blob).hexdigest()[:10]


_flat = {}


def is_flat(slide, pic):
    """Картинка одного цвета (белые/цветные заглушки-маски старого шаблона) — не контент."""
    b = pic.find('.//a:blip', NS)
    if b is None or not b.get(R + 'embed'): return False
    blob = slide.part.related_part(b.get(R + 'embed')).blob
    h = hashlib.sha1(blob).hexdigest()[:10]
    if h not in _flat:
        from PIL import Image
        try:
            im = Image.open(io.BytesIO(blob)).convert('RGBA')
            ex = im.getextrema()
            _flat[h] = all(hi - lo <= 6 for lo, hi in ex[:3]) and ex[3][0] > 200
        except Exception:
            _flat[h] = False
    return _flat[h]


def new_rpr(sz, b=False, color='0F1115', i=False, u=None, font='Open Sans'):
    r = etree.Element(A + 'rPr'); r.set('lang', 'ru'); r.set('sz', str(sz))
    r.set('b', '1' if b else '0')
    if i: r.set('i', '1')
    if u: r.set('u', u)
    r.append(solid(color))
    for t in ('latin', 'ea', 'cs'):
        e = etree.SubElement(r, A + t); e.set('typeface', font)
    return r


# ======================================================================
class Builder:
    def __init__(self, base=None):
        self.prs = Presentation(base or BASE)
        self.orig = list(self.prs.slides._sldIdLst)
        self.donors = list(self.prs.slides)
        self.log = []

    # ---- клонирование слайда-донора ----
    def clone(self, n, keep_content=True):
        d = self.donors[n - 1]
        s = self.prs.slides.add_slide(d.slide_layout)
        tree = s.shapes._spTree
        for el in list(tree)[2:]: tree.remove(el)
        for el in list(d.shapes._spTree)[2:]:
            if not keep_content:
                ph = el.find('.//p:nvPr/p:ph', NS)
                if ph is None or ph.get('type') not in ('title', 'sldNum', 'ctrTitle'): continue
            ne = copy.deepcopy(el)
            self._remap_rels(ne, d.part, s.part)
            tree.append(ne)
        # фон слайда
        bg = d._element.find('p:cSld/p:bg', NS)
        if bg is not None:
            s._element.find('p:cSld', NS).insert(0, copy.deepcopy(bg))
        return s

    def _remap_rels(self, el, src_part, dst_part):
        for e in el.iter():
            for att in (R + 'embed', R + 'link', R + 'id'):
                rid = e.get(att)
                if not rid: continue
                rel = src_part.rels.get(rid) if hasattr(src_part.rels, 'get') else src_part.rels[rid]
                if rel is None: continue
                if rel.is_external:
                    nid = dst_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
                elif rel.reltype == RT.IMAGE:
                    _, nid = dst_part.get_or_add_image_part(io.BytesIO(rel.target_part.blob))
                else:
                    nid = dst_part.relate_to(rel.target_part, rel.reltype)
                e.set(att, nid)

    def add_image(self, s, path_or_blob, x, y, w, h, rot=0, crop=None):
        blob = open(path_or_blob, 'rb').read() if isinstance(path_or_blob, str) else path_or_blob
        _, rid = s.part.get_or_add_image_part(io.BytesIO(blob))
        nid = self.next_id(s)
        xml = ('<p:pic xmlns:p="%s" xmlns:a="%s" xmlns:r="%s"><p:nvPicPr><p:cNvPr id="%d" name="Picture %d"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>'
               '<p:blipFill><a:blip r:embed="%s"/>%s<a:stretch><a:fillRect/></a:stretch></p:blipFill>'
               '<p:spPr><a:xfrm%s><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>') % (
            NS['p'], NS['a'], NS['r'], nid, nid, rid,
            ('<a:srcRect l="%d" t="%d" r="%d" b="%d"/>' % crop) if crop else '',
            (' rot="%d"' % rot) if rot else '', x, y, w, h)
        el = etree.fromstring(xml)
        s.shapes._spTree.append(el)
        return el

    def next_id(self, s):
        m = 0
        for e in s._element.iter(P + 'cNvPr'):
            try: m = max(m, int(e.get('id')))
            except: pass
        return m + 1

    def finish(self, path):
        lst = self.prs.slides._sldIdLst
        for sid in self.orig:
            rid = sid.get(R + 'id')
            lst.remove(sid)
            self.prs.part.drop_rel(rid)
        self.prs.save(path)

    # ---- общие правки текста ----
    @staticmethod
    def set_text_runs(txBody, paragraphs, keep_first_ppr=True):
        """paragraphs: список абзацев; абзац = список (text, rPr-element|None)."""
        ps = txBody.findall('a:p', NS)
        tmpl = ps[0]
        tmpl_r = tmpl.find('a:r', NS)
        base_rpr = tmpl_r.find('a:rPr', NS) if tmpl_r is not None else None
        for p in ps: txBody.remove(p)
        for para in paragraphs:
            p = etree.SubElement(txBody, A + 'p')
            pPr = tmpl.find('a:pPr', NS)
            if pPr is not None and keep_first_ppr: p.append(copy.deepcopy(pPr))
            for txt, rpr in para:
                if txt == '\n':
                    br = etree.SubElement(p, A + 'br')
                    if base_rpr is not None: br.append(copy.deepcopy(base_rpr))
                    continue
                r = etree.SubElement(p, A + 'r')
                rp = rpr if rpr is not None else (copy.deepcopy(base_rpr) if base_rpr is not None else etree.Element(A + 'rPr'))
                r.append(rp)
                t = etree.SubElement(r, A + 't'); t.text = txt
        return txBody

    @staticmethod
    def set_title(s, text):
        for sp in s.shapes._spTree.iter(P + 'sp'):
            ph = sp.find('p:nvSpPr/p:nvPr/p:ph', NS)
            if ph is not None and ph.get('type') in ('title', 'ctrTitle'):
                tb = sp.find('p:txBody', NS)
                lines = text.split('\n')
                para = []
                for i, l in enumerate(lines):
                    if i: para.append(('\n', None))
                    para.append((l, None))
                Builder.set_text_runs(tb, [para])
                return sp
