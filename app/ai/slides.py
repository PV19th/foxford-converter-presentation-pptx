# -*- coding: utf-8 -*-
"""Слайд глазами ИИ: список фигур с координатами (describe) и применение правок раскладки (apply).

Правки трогают только положение и размер, кегль, выравнивание, заливку и порядок слоёв. Текст не меняется
никогда, удалить можно только фигуру без текста."""
from lxml import etree

EMU = 914400
NS = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
A, P = '{%s}' % NS['a'], '{%s}' % NS['p']
KINDS = {'sp': 'фигура', 'pic': 'картинка', 'grpSp': 'группа', 'graphicFrame': 'таблица', 'cxnSp': 'линия'}
PALETTE = ['641FF1', 'EEE5FF', 'FF9900', 'FFEED7', '71D2FF', 'CCEAFF', 'FFFA7D', 'FFFED6', 'FFFFFF']
ALIGN = {'left': 'l', 'center': 'ctr', 'right': 'r'}


def _xfrm(el):
    return next((x for x in (el.find(q, NS) for q in ('p:spPr/a:xfrm', 'p:grpSpPr/a:xfrm', 'p:xfrm')) if x is not None), None)


def _box(el):
    x = _xfrm(el)
    off, ext = (x.find('a:off', NS), x.find('a:ext', NS)) if x is not None else (None, None)
    return (off, ext) if off is not None and ext is not None else (None, None)


def shapes(slide):
    """Фигуры верхнего уровня, у которых есть координаты: {id: элемент}."""
    out = {}
    for el in slide.shapes._spTree.iterchildren():
        if etree.QName(el).localname in KINDS and _box(el)[0] is not None:
            out[int(el.find('.//p:cNvPr', NS).get('id'))] = el
    return out


def text(el):
    return ''.join(t.text or '' for t in el.iter(A + 't'))


def slide_text(slide):
    """Тексты всех фигур слайда — для проверки, что правки их не задели (порядок слоёв не важен)."""
    return sorted(text(el) for el in slide.shapes._spTree.iterchildren())


LOCKED = ('title', 'ctrTitle', 'sldNum', 'ftr', 'dt')  # заголовок, номер слайда, колонтитулы ИИ не трогает


def locked(el):
    ph = el.find('.//p:nvPr/p:ph', NS)
    return ph is not None and ph.get('type') in LOCKED


def describe(slide):
    out = []
    for sid, el in shapes(slide).items():
        off, ext = _box(el)
        d = {'id': sid, 'kind': KINDS[etree.QName(el).localname],
             'x': round(int(off.get('x')) / EMU, 2), 'y': round(int(off.get('y')) / EMU, 2),
             'w': round(int(ext.get('cx')) / EMU, 2), 'h': round(int(ext.get('cy')) / EMU, 2)}
        t = ' '.join(text(el).split())
        if t: d['text'] = t[:200] + ('…' if len(t) > 200 else '')
        sizes = sorted({int(r.get('sz')) / 100 for r in el.iter(A + 'rPr', A + 'endParaRPr') if r.get('sz')})
        if sizes: d['font_pt'] = sizes
        fill = el.find('p:spPr/a:solidFill/a:srgbClr', NS)
        if fill is not None: d['fill'] = fill.get('val')
        if locked(el): d['locked'] = True
        out.append(d)
    return out


def _set_fill(el, color):
    spPr = el.find('p:spPr', NS)
    if spPr is None: return False
    for f in [c for c in spPr if etree.QName(c).localname in ('solidFill', 'gradFill', 'noFill', 'pattFill', 'blipFill')]:
        spPr.remove(f)
    new = etree.Element(A + ('noFill' if color == 'none' else 'solidFill'))
    if color != 'none': etree.SubElement(new, A + 'srgbClr').set('val', color)
    geom = next((c for c in spPr if etree.QName(c).localname in ('prstGeom', 'custGeom')), None)
    spPr.insert(spPr.index(geom) + 1 if geom is not None else len(spPr), new)  # заливка — сразу после формы
    return True


def apply(slide, edits, size):
    """Применяет правки ИИ к слайду. size — (ширина, высота) слайда в EMU. Возвращает число применённых правок."""
    W, H = size
    by_id, done = shapes(slide), 0
    for e in edits:
        el = by_id.get(e.get('id'))
        if el is None or locked(el): continue
        if e.get('delete'):
            if not text(el).strip(): el.getparent().remove(el); done += 1
            continue
        off, ext = _box(el)
        cur = [int(off.get('x')), int(off.get('y')), int(ext.get('cx')), int(ext.get('cy'))]
        new = [cur[i] if e.get(k) is None else round(e[k] * EMU) for i, k in enumerate('xywh')]
        if new != cur:
            kind = etree.QName(el).localname
            if kind in ('pic', 'grpSp') and cur[2] and cur[3]:  # пропорции картинки и группы не искажаем
                if e.get('w') is not None: new[3] = round(new[2] * cur[3] / cur[2])
                elif e.get('h') is not None: new[2] = round(new[3] * cur[2] / cur[3])
            new[2], new[3] = max(new[2], EMU // 20), max(new[3], EMU // 20)
            new[2], new[3] = min(new[2], W), min(new[3], H)
            new[0], new[1] = min(max(new[0], 0), W - new[2]), min(max(new[1], 0), H - new[3])
            if kind == 'graphicFrame' and new[2] != cur[2]:  # ширина таблицы — сумма колонок
                for g in el.iter(A + 'gridCol'): g.set('w', str(round(int(g.get('w')) * new[2] / cur[2])))
            off.set('x', str(new[0])); off.set('y', str(new[1])); ext.set('cx', str(new[2])); ext.set('cy', str(new[3]))
            done += 1
        if e.get('font_pt'):  # по чек-листу кегль только 16, 14 и 12 — берём ближайший
            sz = str(min((16, 14, 12), key=lambda v: abs(v - e['font_pt'])) * 100)
            for r in el.iter(A + 'rPr', A + 'endParaRPr'): r.set('sz', sz)
            done += 1
        if e.get('align') in ALIGN:
            for p in el.iter(A + 'p'):
                pPr = p.find('a:pPr', NS)
                if pPr is None: pPr = etree.Element(A + 'pPr'); p.insert(0, pPr)
                pPr.set('algn', ALIGN[e['align']])
            done += 1
        if e.get('fill') in PALETTE + ['none'] and _set_fill(el, e['fill']): done += 1
        if e.get('z') in ('front', 'back'):
            tree = el.getparent(); tree.remove(el)
            if e['z'] == 'front': tree.append(el)
            else: tree.insert(2, el)  # сразу после служебных nvGrpSpPr и grpSpPr
            done += 1
    return done
