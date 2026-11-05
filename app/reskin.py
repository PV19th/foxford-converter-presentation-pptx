# -*- coding: utf-8 -*-
"""Перенос «пересадкой» — движок для индивидуальных шаблонов (профили с "engine": "reskin").

Старый и новый шаблон автора устроены одинаково: те же макеты, те же места заголовка и текста —
поменялись фон, декор и палитра. Поэтому слайды не пересобираются (как в converter/), а переносятся как есть:
- обычный слайд садится на макет нового шаблона с тем же именем, а если такого нет — с тем же
  расположением заголовка и полей; цвета перекрашиваются по таблице "recolor" профиля
  (составлена по правкам методиста); свой фон слайда убирается — фон даёт макет;
- титул и разделители берутся из шаблона, текст в них переносится со старых слайдов;
- служебные слайды (домашка, настроение, оценка урока, мотивашка, контакты, прощание) ставятся
  из шаблона целиком; вариант выбирает пользователь на сайте (раздел "choices" профиля).
Картинки и стикеры на слайдах остаются старыми — их меняет методист (памятка «Проверьте вручную»).
"""
import copy, io, re
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree

NS = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
A, P, R = ('{%s}' % NS[k] for k in 'apr')
IN = 914400
# цвета, на которых держится вёрстка (текст, фон), — не перекрашиваем
KEEP_SLOTS = {'tx1', 'bg1', 'dk1', 'lt1', 'phClr'}
TEXT_PARENTS = {'rPr', 'defRPr', 'endParaRPr', 'buClr', 'fontRef'}
SKIP = {'sldNum', 'ftr', 'dt'}  # служебные плейсхолдеры — не влияют на выбор макета


# ---------------------------------------------------------------- тексты
def paragraphs(sp):
    """Абзацы фигуры: [[строка, строка…], …] (строки — по разрывам <a:br/>), без пустых абзацев."""
    out = []
    for p in sp.iter(A + 'p'):
        lines, cur = [], ''
        for el in p:
            if el.tag in (A + 'r', A + 'fld'): cur += ''.join(t.text or '' for t in el.iter(A + 't'))
            elif el.tag == A + 'br': lines.append(cur); cur = ''
        lines.append(cur)
        if ''.join(lines).strip(): out.append([l.strip() for l in lines])
    return out


def flat(sp):
    return ' '.join(' '.join(l for l in p) for p in paragraphs(sp)).strip()


def put_text(sp, paras):
    """Текст фигуры шаблона заменяется на paras; оформление — от первого абзаца и первого прогона шаблона."""
    body = sp.find('.//' + P + 'txBody')
    old = body.findall(A + 'p')
    ppr = old[0].find(A + 'pPr') if old else None
    rpr = next(sp.iter(A + 'rPr'), None)
    if rpr is None: rpr = next(sp.iter(A + 'endParaRPr'), None)
    for p in old: body.remove(p)

    def run_pr(tag):
        if rpr is None: return None
        e = copy.deepcopy(rpr); e.tag = A + tag
        for k in ('dirty', 'err'): e.attrib.pop(k, None)
        return e
    for lines in paras:
        p = etree.SubElement(body, A + 'p')
        if ppr is not None: p.append(copy.deepcopy(ppr))
        for i, line in enumerate(lines):
            if i:
                br = etree.SubElement(p, A + 'br')
                if rpr is not None: br.append(run_pr('rPr'))
            r = etree.SubElement(p, A + 'r')
            if rpr is not None: r.append(run_pr('rPr'))
            etree.SubElement(r, A + 't').text = line


def ph_of(el):
    return el.find('./*/' + P + 'nvPr/' + P + 'ph')


def ph_type(ph):
    t = ph.get('type', 'body')
    return {'ctrTitle': 'title'}.get(t, t)


# ---------------------------------------------------------------- геометрия плейсхолдеров
def _xfrm(el):
    x = el.find('./' + P + 'spPr/' + A + 'xfrm')
    if x is None: x = el.find('./' + P + 'xfrm')
    if x is None: x = el.find('./' + P + 'grpSpPr/' + A + 'xfrm')
    if x is None: return None
    off, ext = x.find(A + 'off'), x.find(A + 'ext')
    if off is None or ext is None: return None
    return tuple(int(v) for v in (off.get('x'), off.get('y'), ext.get('cx'), ext.get('cy')))


def layout_phs(layout):
    """{(тип, idx): (x, y, w, h)} плейсхолдеров макета; положение без своего — из мастера."""
    master = {ph_type(ph_of(e)): _xfrm(e) for e in layout.slide_master.shapes._spTree if ph_of(e) is not None}
    out = {}
    for e in layout.shapes._spTree:
        ph = ph_of(e)
        if ph is None or ph.get('type') in SKIP: continue
        g = _xfrm(e) or master.get(ph_type(ph))
        if g: out[(ph_type(ph), ph.get('idx', '0'))] = g
    return out


def inherited(slide, ph):
    """Положение плейсхолдера слайда без своего xfrm — с макета (по idx, потом по типу; там же и мастер)."""
    phs = layout_phs(slide.slide_layout)
    return phs.get((ph_type(ph), ph.get('idx', '0'))) or next((v for (t, _), v in phs.items() if t == ph_type(ph)), None)


def used_phs(slide):
    out = {}
    for e in slide.shapes._spTree:
        ph = ph_of(e)
        if ph is None or ph.get('type') in SKIP: continue
        g = _xfrm(e) or inherited(slide, ph)
        if g: out[(ph_type(ph), ph.get('idx', '0'))] = g
    return out


def decor_boxes(layout):
    """Крупные картинки макета — декор (полоса цветов, фото): (x, y, w, h)."""
    return [g for e in layout.shapes._spTree if e.tag == P + 'pic' and ph_of(e) is None
            and (g := _xfrm(e)) and g[2] * g[3] > IN * IN]


def content_boxes(slide):
    """Где на слайде текст, таблицы и картинки (кроме заголовка)."""
    out = []
    for e in slide.shapes._spTree:
        if e.tag not in (P + 'sp', P + 'pic', P + 'graphicFrame', P + 'grpSp'): continue
        ph = ph_of(e)
        if ph is not None and ph.get('type') in {'title', 'ctrTitle'} | SKIP: continue
        if e.tag == P + 'sp' and not flat(e): continue
        g = _xfrm(e) or (inherited(slide, ph) if ph is not None else None)
        if g: out.append(g)
    return out


def overlap(a, b):
    w = min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0])
    h = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
    return max(0, w) * max(0, h) / IN / IN


def pick_layout(slide, candidates):
    """Макет нового шаблона: с тем же именем, иначе — у которого заголовок и поля стоят там же, где на слайде.
    Макет, чей декор (полоса цветов, фото) лёг бы на текст или картинки слайда, не подходит."""
    name, used, boxes = slide.slide_layout.name, used_phs(slide), content_boxes(slide)

    def dist(a, b): return sum(abs(u - v) for u, v in zip(a, b)) / IN

    def score(lay, phs, decor, bg_picture):
        s = (100 if lay.name == name else 0) - (2 if bg_picture else 0)  # при равенстве — макет с белым фоном
        for (t, _), g in used.items():
            s += max((10 + max(0, 10 - 10 * dist(g, g2)) for (t2, _), g2 in phs.items() if t2 == t), default=0)
        s -= 60 * sum(1 for b in boxes if any(overlap(b, d) > 0.05 for d in decor))
        return s - 0.1 * len(phs)
    return max(candidates, key=lambda c: score(*c))[0]


def pin_positions(slide, el_list):
    """Плейсхолдеры без своего положения получают его со старого макета: на новом макете поле может стоять иначе."""
    for e in el_list:
        ph = ph_of(e)
        if ph is None or _xfrm(e) is not None: continue
        g = inherited(slide, ph)
        sppr = e.find(P + 'spPr')
        if g is None or sppr is None: continue
        x = etree.Element(A + 'xfrm')
        etree.SubElement(x, A + 'off', x=str(g[0]), y=str(g[1]))
        etree.SubElement(x, A + 'ext', cx=str(g[2]), cy=str(g[3]))
        sppr.insert(0, x)


# ---------------------------------------------------------------- цвета
def theme_colors(master):
    """{имя цвета: HEX} темы мастера. Синонимы bg1/tx1/bg2/tx2 — по карте цветов мастера: в шаблонах из
    Google Slides она переставлена (bg2 — тёмный dk2, tx2 — светлый lt2)."""
    t = etree.fromstring(master.part.part_related_by(RT.THEME).blob)
    out = {etree.QName(c).localname: (c[0].get('val') or c[0].get('lastClr') or '').upper()
           for c in t.find('.//' + A + 'clrScheme')}
    cmap = master._element.find(P + 'clrMap')
    for alias, default in (('bg1', 'lt1'), ('tx1', 'dk1'), ('bg2', 'lt2'), ('tx2', 'dk2')):
        out[alias] = out.get(cmap.get(alias, default) if cmap is not None else default)
    return out


def recolor(root, theme, rules):
    """Цвета старой палитры → цвета нового шаблона (rules["fill"] / rules["text"]: HEX → цвет темы
    по прямому имени dk1/lt1/dk2/lt2/accent1–6 — синонимы bg2/tx2 в шаблонах переставлены).
    Осветлённые цвета палитры (lumOff) — в светлые цвета шаблона. Остальные цвета не меняются:
    цвет темы, которого нет в таблице, записывается как был (HEX), чтобы новая тема его не подменила."""
    fill, text = rules.get('fill', {}), rules.get('text', {})
    for c in list(root.iter(A + 'srgbClr', A + 'schemeClr')):
        anc = {etree.QName(x).localname for x in c.iterancestors()}
        if 'effectLst' in anc or 'highlight' in anc: continue
        scheme = c.tag == A + 'schemeClr'
        if scheme and c.get('val') in KEEP_SLOTS: continue
        hexv = theme.get(c.get('val')) if scheme else (c.get('val') or '').upper()
        if not hexv: continue
        role = text if anc & TEXT_PARENTS else fill
        tint = next((int(m.get('val')) for m in c if m.tag == A + 'lumOff'), 0)
        if tint and role is fill and hexv in fill:
            slot, keep_mods = rules['tint_light'] if tint >= 70000 else rules['tint'], False
        else:
            slot, keep_mods = role.get(hexv), False
        if slot is None:
            if not scheme: continue
            slot, keep_mods = None, True  # цвет темы вне таблицы — оставить как был
        new = etree.Element(A + ('schemeClr' if slot else 'srgbClr'), val=slot or hexv)
        for m in c:
            if keep_mods or m.tag == A + 'alpha': new.append(copy.deepcopy(m))
        c.getparent().replace(c, new)


def _hex(c, theme):
    """HEX цвета-элемента (srgbClr / schemeClr без оттенков) или None."""
    if c is None or len(c) and any(m.tag != A + 'alpha' for m in c): return None
    if c.tag == A + 'srgbClr': return (c.get('val') or '').upper() or None
    if c.tag == A + 'schemeClr': return theme.get(c.get('val'))
    return None


def _light(h):
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def fix_contrast(root, theme, rules):
    """Плашка и её текст после перекраски не должны слиться: тёмная заливка под тёмным текстом становится
    светлой (rules["tint"]), светлая под белым текстом — тёмной ("dk2"). Так делал методист: один и тот же
    розовый становился бордовым под белым текстом и оставался розовым под чёрным."""
    for el in root.iter(P + 'sp', A + 'tc'):
        pr = el.find(P + 'spPr') if el.tag == P + 'sp' else el.find(A + 'tcPr')
        sf = pr.find(A + 'solidFill') if pr is not None else None
        fill = _hex(sf[0] if sf is not None and len(sf) else None, theme)
        if not fill: continue
        default = _hex(next(iter(el.findall('./' + P + 'style/' + A + 'fontRef/*')), None), theme) or theme.get('dk1', '000000')
        runs = [r for r in el.iter(A + 'r') if (r.findtext(A + 't') or '').strip()]
        if not runs: continue
        colors = []
        for r in runs:
            c = r.find(A + 'rPr/' + A + 'solidFill')
            colors.append(_hex(c[0] if c is not None and len(c) else None, theme) if c is not None else default)
        if None in colors: continue
        f, t = _light(fill), [_light(c) for c in colors]
        if f < 110 and max(t) < 110: new = rules['tint']
        elif f > 170 and min(t) > 200: new = 'dk2'
        else: continue
        sf[0].getparent().replace(sf[0], etree.Element(A + 'schemeClr', val=new))


# ---------------------------------------------------------------- сборка
class Reskin:
    def __init__(self, base):
        self.prs = Presentation(base)
        self.orig = list(self.prs.slides._sldIdLst)
        self.tpl = list(self.prs.slides)
        self.slide_links = []  # (элемент, атрибут, старый слайд) — ссылки на слайды переставляются в конце
        self.moved = {}        # часть старого слайда → новый слайд (ссылки «перейти к слайду»)

    def _remap(self, el, src, dst):
        for e in el.iter():
            for att in (R + 'embed', R + 'link', R + 'id', R + 'pict'):
                rid = e.get(att)
                if not rid: continue
                rel = src.rels.get(rid) if hasattr(src.rels, 'get') else src.rels[rid]
                if rel is None: continue
                if rel.is_external:
                    e.set(att, dst.relate_to(rel.target_ref, rel.reltype, is_external=True))
                elif rel.reltype == RT.IMAGE:
                    e.set(att, dst.get_or_add_image_part(io.BytesIO(rel.target_part.blob))[1])
                elif rel.reltype == RT.SLIDE:
                    self.slide_links.append((e, att, rel.target_part, dst))
                else:
                    e.set(att, dst.relate_to(rel.target_part, rel.reltype))

    def copy_slide(self, src, layout, keep_bg=False, notes=True):
        """Слайд src (из шаблона или старой презентации) на макете layout: фигуры, анимация, заметки."""
        s = self.prs.slides.add_slide(layout)
        tree = s.shapes._spTree
        for el in list(tree)[2:]: tree.remove(el)
        els = [copy.deepcopy(el) for el in list(src.shapes._spTree)[2:]]
        for el in els:
            self._remap(el, src.part, s.part)
            tree.append(el)
        bg = src._element.find(P + 'cSld/' + P + 'bg')
        if keep_bg and bg is not None: s._element.find(P + 'cSld').insert(0, copy.deepcopy(bg))
        for tag in ('transition', 'timing'):
            t = src._element.find(P + tag)
            if t is not None: s._element.append(copy.deepcopy(t))
        if src._element.get('show') == '0': s._element.set('show', '0')
        if notes and src.has_notes_slide and src.notes_slide.notes_text_frame is not None:
            note = src.notes_slide.notes_text_frame.text
            if note.strip(): s.notes_slide.notes_text_frame.text = note
        self.moved[src.part] = s
        return s, els

    def from_template(self, n):
        d = self.tpl[n - 1]
        return self.copy_slide(d, d.slide_layout, keep_bg=True, notes=False)[0]

    def table_styles(self, old):
        """Стили таблиц старой презентации, которых нет в новой, — переносятся (иначе таблица теряет оформление)."""
        try:
            dst, src = (p.part.part_related_by(RT.TABLE_STYLES) for p in (self.prs, old))
        except KeyError:
            return
        used = {e.text for s in self.prs.slides for e in s._element.iter(A + 'tableStyleId')}
        d = dst._element if hasattr(dst, '_element') else etree.fromstring(dst.blob)
        have = {e.get('styleId') for e in d}
        new = [e for e in etree.fromstring(src.blob) if e.get('styleId') in used - have]
        if not new: return
        d.extend(copy.deepcopy(e) for e in new)
        if not hasattr(dst, '_element'): dst._blob = etree.tostring(d, xml_declaration=True, encoding='UTF-8', standalone=True)

    def finish(self, path):
        for e, att, old, dst in self.slide_links:  # ссылка на слайд урока → на его копию, иначе убрать
            new = self.moved.get(old)
            if new is not None and new.part is not dst: e.set(att, dst.relate_to(new.part, RT.SLIDE))
            elif e.tag in (A + 'hlinkClick', A + 'hlinkHover'): e.getparent().remove(e)
            else: e.attrib.pop(att, None)
        lst = self.prs.slides._sldIdLst
        for sid in self.orig:
            lst.remove(sid)
            self.prs.part.drop_rel(sid.get(R + 'id'))
        self.prs.save(path)


def classify(n, slide, profile):
    """Чем старый слайд является: title, section, служебный слайд из "roles" профиля или content."""
    if n == 1: return 'title'
    text = ' '.join(flat(sp) for sp in slide.shapes._spTree.iter(P + 'sp'))
    for kind, pattern in profile['roles'].items():
        if re.search(pattern, text, re.I):
            if kind == 'bye' and len(text.split()) > 12: continue
            return kind
    if re.search(profile['sections']['layouts'], slide.slide_layout.name) and len(text.split()) <= 20:
        return 'section'
    return 'content'


def split_title(slide):
    """Заголовок и остальной текст слайда (титул, разделитель): (заголовок, подзаголовок, теги)."""
    title, sub, tags, rest = [], [], [], []
    for sp in slide.shapes._spTree.iter(P + 'sp'):
        paras = paragraphs(sp)
        if not paras: continue
        ph = ph_of(sp)
        t = ph.get('type') if ph is not None else None
        if t in ('ctrTitle', 'title') and not title: title = paras
        elif t == 'subTitle' and ph.get('idx') == '2' or any('|' in ' '.join(p) for p in paras): tags += paras
        elif t == 'subTitle': sub += paras
        else: rest += paras
    return title, sub or rest, tags


def fill(slide, wanted):
    """Плейсхолдеры слайда шаблона: {тип: абзацы}; пустое — фигура удаляется (иначе останется заглушка)."""
    for sp in list(slide.shapes._spTree.iter(P + 'sp')):
        ph = ph_of(sp)
        if ph is None: continue
        key = 'tags' if ph.get('type') == 'subTitle' and ph.get('idx') == '2' else ph_type(ph)
        if key not in wanted: continue
        if wanted[key]: put_text(sp, wanted[key])
        else: sp.getparent().remove(sp)


def convert(src, out, profile, choices=None):
    """Старая презентация src → out. Возвращает строки отчёта, как у convert.py:
    [{'src': номер старого слайда или None, 'kind', 'notes', 'replaced': слайд целиком из шаблона}]."""
    from pathlib import Path
    root = Path(__file__).resolve().parent.parent
    choices = dict(choices or {})
    old = Presentation(src)
    b = Reskin(root / profile['base'])
    content = [(lay, layout_phs(lay), decor_boxes(lay), lay._element.find(P + 'cSld/' + P + 'bg//' + A + 'blip') is not None)
               for lay in (b.tpl[n - 1].slide_layout for n in profile['content_layouts'])]
    # особые макеты — только по имени старого макета (рамка «скриншот» на обычном слайде мешает)
    named = [(re.compile(pat), b.tpl[n - 1].slide_layout) for pat, n in profile.get('layout_names', {}).items()]
    sections = profile['sections']
    used = {'dividers': 0, 'questions': 0}
    title_tags = ''
    rows = []
    for n, s in enumerate(old.slides, 1):
        kind = classify(n, s, profile)
        row = {'src': n, 'kind': kind, 'notes': [], 'replaced': False}
        if kind == 'title':
            title, sub, tags = split_title(s)
            title_tags = ' '.join(' '.join(p) for p in tags)
            new = b.from_template(choices.get('title') or profile['choices']['title']['default'])
            fill(new, {'title': title, 'subTitle': sub, 'tags': tags})
            b.moved[s.part] = new
            if not title: row['notes'].append('на старом титуле не нашлось названия урока')
        elif kind == 'section':
            title, sub, _ = split_title(s)
            title = title or sub[:1]; sub = sub[1:] if title == sub[:1] else sub
            # строки старого заголовка разбиты под старый макет — в новом разделителе переносит сам шаблон
            title = [[' '.join(l for l in p if l)] for p in title]
            question = title and title[-1][-1].endswith('?') and not sub
            pool = sections['questions' if question else 'dividers']
            k = 'questions' if question else 'dividers'
            new = b.from_template(pool[used[k] % len(pool)]); used[k] += 1
            fill(new, {'title': title, 'subTitle': sub})
            b.moved[s.part] = new
        elif kind in profile['choices']:
            b.moved[s.part] = b.from_template(choices.get(kind) or auto_choice(profile['choices'][kind], title_tags))
            row['replaced'] = True
        else:
            layout = next((lay for pat, lay in named if pat.search(s.slide_layout.name)), None) or pick_layout(s, content)
            new, els = b.copy_slide(s, layout)
            pin_positions(s, els)
            th, new_th = theme_colors(s.slide_layout.slide_master), theme_colors(layout.slide_master)
            for el in els:
                recolor(el, th, profile['recolor'])
                fix_contrast(el, new_th, profile['recolor'])
        rows.append(row)
    # конец урока — контакты и прощание (так у методиста); чего нет в старой — добавляется
    for kind, note in (('contacts', 'добавлен слайд контактов'), ('bye', 'добавлен прощальный слайд')):
        if kind in profile['choices'] and not any(r['kind'] == kind for r in rows):
            b.from_template(choices.get(kind) or auto_choice(profile['choices'][kind], title_tags))
            row = {'src': None, 'kind': kind, 'notes': [note], 'replaced': True}
            if kind == 'contacts' and rows and rows[-1]['kind'] == 'bye':  # контакты — перед прощанием
                lst = b.prs.slides._sldIdLst
                lst.insert(len(lst) - 2, lst[-1])
                rows.insert(len(rows) - 1, row)
            else:
                rows.append(row)
    b.table_styles(old)
    b.finish(out)
    for i, r in enumerate(rows, 1): r['slide'] = i
    return rows


def auto_choice(c, tags):
    """Вариант по умолчанию; "auto" — по тегам титула (контакты ОГЭ / ЕГЭ: свои QR-коды)."""
    for word, n in (c.get('auto') or {}).items():
        if re.search(word, tags): return n
    return c['default']


if __name__ == '__main__':
    import json, sys
    from pathlib import Path
    prof = json.loads((Path(__file__).resolve().parent.parent / 'templates' / sys.argv[3] / 'profile.json').read_text('utf8'))
    for r in convert(sys.argv[1], sys.argv[2], prof): print(r)
