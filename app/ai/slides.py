# -*- coding: utf-8 -*-
"""Слайд глазами ИИ: список фигур с координатами (describe), применение правок раскладки (apply)
и автоматическая проверка раскладки (problems) — она подсказывает ИИ, что не так, и не даёт сделать хуже.

Правки трогают только положение и размер, кегль, выравнивание, заливку и порядок слоёв. Текст не меняется
никогда, удалить можно только фигуру без текста."""
import copy
from lxml import etree
from PIL import ImageFont

EMU = 914400
NS = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
A, P = '{%s}' % NS['a'], '{%s}' % NS['p']
KINDS = {'sp': 'фигура', 'pic': 'картинка', 'grpSp': 'группа', 'graphicFrame': 'таблица', 'cxnSp': 'линия'}
PALETTE = ['641FF1', 'EEE5FF', 'FF9900', 'FFEED7', '71D2FF', 'CCEAFF', 'FFFA7D', 'FFFED6', 'FFFFFF']
ALIGN = {'left': 'l', 'center': 'ctr', 'right': 'r'}
FONT = '/usr/share/fonts/truetype/open-sans/OpenSans-%s.ttf'  # шрифт шаблона, как в converter/engine.py
_fonts = {}


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
    """Тексты всех фигур слайда — для проверки, что правки их не задели (порядок слоёв и удалённые
    пустые фигуры не важны)."""
    return sorted(t for t in (text(el) for el in slide.shapes._spTree.iterchildren()) if t.strip())


LOCKED = ('title', 'ctrTitle', 'sldNum', 'ftr', 'dt')  # заголовок, номер слайда, колонтитулы ИИ не трогает


def locked(el):
    ph = el.find('.//p:nvPr/p:ph', NS)
    return ph is not None and ph.get('type') in LOCKED


def describe(slide):
    out = []
    for sid, el in shapes(slide).items():
        x, y, w, h = _rect(el)
        d = {'id': sid, 'kind': KINDS[etree.QName(el).localname], 'x': round(x, 2), 'y': round(y, 2), 'w': round(w, 2), 'h': round(h, 2)}
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
        cols = list(el.iter(A + 'gridCol'))
        cur = [int(off.get('x')), int(off.get('y')), max([int(ext.get('cx')), sum(int(g.get('w')) for g in cols)]),
               int(ext.get('cy'))]
        new = [cur[i] if e.get(k) is None else round(e[k] * EMU) for i, k in enumerate('xywh')]
        if new != cur:
            kind = etree.QName(el).localname
            if kind in ('pic', 'grpSp') and cur[2] and cur[3]:  # пропорции картинки и группы не искажаем
                if e.get('w') is not None: new[3] = round(new[2] * cur[3] / cur[2])
                elif e.get('h') is not None: new[2] = round(new[3] * cur[2] / cur[3])
            new[2], new[3] = max(new[2], EMU // 20), max(new[3], EMU // 20)
            new[2], new[3] = min(new[2], W), min(new[3], H)
            new[0], new[1] = min(max(new[0], 0), W - new[2]), min(max(new[1], 0), H - new[3])
            if cols and new[2] != cur[2]:  # ширина таблицы — сумма колонок
                total = sum(int(g.get('w')) for g in cols)
                for g in cols: g.set('w', str(round(int(g.get('w')) * new[2] / total)))
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


def restore(slide, original):
    """Вернуть слайду раскладку версии без ИИ (связи с картинками у правок не меняются)."""
    tree = slide.shapes._spTree
    tree.getparent().replace(tree, copy.deepcopy(original.shapes._spTree))


def _rect(el):
    off, ext = _box(el)
    w = int(ext.get('cx'))
    cols = [int(g.get('w')) for g in el.iter(A + 'gridCol')]
    if cols: w = max(w, sum(cols))  # таблица рисуется по ширине колонок, даже если рамка уже
    return [int(off.get('x')) / EMU, int(off.get('y')) / EMU, w / EMU, int(ext.get('cy')) / EMU]


def _overlap(a, b):
    return max(0, min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0])) * max(0, min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1]))


def _width(s, pt, bold):
    """Ширина строки в дюймах."""
    if (pt, bold) not in _fonts: _fonts[pt, bold] = ImageFont.truetype(FONT % ('Bold' if bold else 'Regular'), int(pt * 20))
    return _fonts[pt, bold].getlength(s) / 20 / 72


def _paragraphs(el):
    """[(текст абзаца, кегль, жирный)] — кегль и жирность по самому крупному и жирному куску."""
    out = []
    for p in el.iter(A + 'p'):
        rprs = [r.find('a:rPr', NS) for r in p.iter(A + 'r')]
        sizes = [int(r.get('sz')) / 100 for r in rprs if r is not None and r.get('sz')]
        out.append((''.join(t.text or '' for t in p.iter(A + 't')), max(sizes or [16]),
                    any(r is not None and r.get('b') == '1' for r in rprs)))
    return out


def _short(s, n=40):
    s = ' '.join(s.split())
    return s if len(s) <= n else s[:n] + '…'


def problems(slide):
    """Что видно по геометрии: слово шире плашки, текст не влезает по высоте, текст на картинке, плашки
    с текстом наложились. → [(для ИИ, с id фигур), (для методиста)]."""
    items = [(sid, el, etree.QName(el).localname, _rect(el)) for sid, el in shapes(slide).items() if not locked(el)]
    texts = [(sid, el, r) for sid, el, kind, r in items if kind == 'sp' and text(el).strip()]
    blocks = texts + [(sid, el, r) for sid, el, kind, r in items if kind == 'graphicFrame' and text(el).strip()]  # + таблицы
    pics = [(sid, r) for sid, el, kind, r in items if kind == 'pic']
    out = []
    for sid, el, r in texts:
        body = el.find('.//a:bodyPr', NS)
        ins = lambda k: int(body.get(k, 91440)) / EMU if body is not None else 0.1
        inner, need, name = r[2] - ins('lIns') - ins('rIns'), ins('tIns') + ins('bIns'), _short(text(el))
        for para, pt, bold in _paragraphs(el):
            word = max(para.split() or [''], key=lambda w: _width(w, pt, bold))
            if word and _width(word, pt, bold) > inner + 0.02:
                out.append(('id %d: слово «%s» шире фигуры (нужно %.2f″, есть %.2f″)' % (sid, word, _width(word, pt, bold), inner),
                            'слово «%s» не помещается в плашку — рвётся на части' % word))
                break
            lines, cur = 1, 0.0  # перенос по словам, как в PowerPoint
            for w in para.split():
                ww = _width(w + ' ', pt, bold)
                if cur and cur + ww > inner: lines, cur = lines + 1, ww
                else: cur += ww
            need += lines * pt * 1.2 / 72
        if need > r[3] + 0.15:
            out.append(('id %d «%s»: текст выше фигуры (нужно %.2f″, есть %.2f″)' % (sid, name, need, r[3]),
                        'текст «%s» не помещается в плашку по высоте' % name))
    for sid, el, r in blocks:
        for pid, pr in pics:
            if _overlap(r, pr) > 0.05:
                out.append(('id %d «%s» заходит на картинку id %d' % (sid, _short(text(el)), pid),
                            'текст «%s» заходит на картинку' % _short(text(el))))
    for i, (a, el_a, ra) in enumerate(blocks):
        for b, el_b, rb in blocks[i + 1:]:
            if _overlap(ra, rb) > 0.05:
                out.append(('фигуры id %d и id %d наложились' % (a, b),
                            'плашки «%s» и «%s» наложились' % (_short(text(el_a), 25), _short(text(el_b), 25))))
    return out
