# -*- coding: utf-8 -*-
"""Собирает базу профиля из фирменного шаблона: python3 tools/build_base.py [папка_профиля ...]

Рецепт — в templates/<профиль>/profile.json, раздел "base_recipe":
- "template": фирменный шаблон (его слайды остаются под своими номерами);
- "append": [[файл, [номера слайдов]], ...] — слайды из готовых уроков (план, «Что нас ждёт» и др.,
  которых нет в шаблоне). Они встают в конец базы по порядку и получают макет шаблона с тем же именем.
  Если урок 4:3, а шаблон 16:9, слайд ужимается по вертикали (squash);
- "titles", "dividers": слайды шаблона, с которых убираются заглушки (подзаголовок «Описание темы»,
  тег «Модуль»), а тег предмета приводится к виду, который понимает конвертер;
- "output": куда положить готовую базу (base/generated/ — в git не попадает).

Конвертер не меняется: он получает обычную базу, как ex58/ex64.
"""
import copy, json, sys
from pathlib import Path
from lxml import etree
from pptx import Presentation

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / 'converter'))
from engine import Builder, NS, A, P, bbox, set_bbox, text_of  # noqa: E402


def layouts_by_name(prs):
    out = {}
    for m in prs.slide_masters:
        for l in m.slide_layouts: out.setdefault(l.name, l)
    return out


def squash(el, vs):
    """Слайд урока 4:3 → база 16:9: по вертикали всё ×vs (у шаблона 16:9 так и есть: 0,39″ → 0,30″,
    6,79″ → 5,09″). Картинки и группы уменьшаются целиком, на месте прежнего центра — без искажения."""
    xfrm = next((x for x in (el.find(q, NS) for q in ('p:spPr/a:xfrm', 'p:grpSpPr/a:xfrm', 'p:xfrm')) if x is not None), None)
    off, ext = (xfrm.find('a:off', NS), xfrm.find('a:ext', NS)) if xfrm is not None else (None, None)
    if off is None or ext is None: return
    x, w = int(off.get('x')), int(ext.get('cx'))
    off.set('y', str(round(int(off.get('y')) * vs))); ext.set('cy', str(round(int(ext.get('cy')) * vs)))
    if etree.QName(el).localname in ('pic', 'grpSp') or el.find('p:spPr/a:blipFill', NS) is not None:
        ext.set('cx', str(round(w * vs))); off.set('x', str(round(x + w * (1 - vs) / 2)))
    for tr in el.iter(A + 'tr'): tr.set('h', str(round(int(tr.get('h')) * vs)))


def append_slide(dst, layouts, src, vs=1):
    """Копия слайда src в конец dst на макете шаблона с тем же именем (vs ≠ 1 — ужать по вертикали)."""
    name = src.slide_layout.name
    if name not in layouts: raise SystemExit('В шаблоне нет макета %r — слайд не перенести' % name)
    s = dst.slides.add_slide(layouts[name])
    tree = s.shapes._spTree
    for el in list(tree)[2:]: tree.remove(el)
    for el in list(src.shapes._spTree)[2:]:
        ne = copy.deepcopy(el)
        Builder._remap_rels(None, ne, src.part, s.part)
        if vs != 1: squash(ne, vs)
        tree.append(ne)
    bg = src._element.find('p:cSld/p:bg', NS)
    if bg is not None: s._element.find('p:cSld', NS).insert(0, copy.deepcopy(bg))


def ph_type(el):
    ph = el.find('.//p:nvPr/p:ph', NS)
    return ph.get('type') if ph is not None else None


def clean_title(s, th):
    tree = s.shapes._spTree
    for el in list(tree.iter(P + 'sp')):
        if ph_type(el) == 'subTitle': el.getparent().remove(el)
    tags = sorted([el for el in tree.iter(P + 'sp') if ph_type(el) is None and bbox(el) and bbox(el)[1] > th - 1.5 * 914400
                   and text_of(el).strip()], key=lambda e: bbox(e)[0])
    # «Предмет · N класс · Модуль 1 · Урок 16» → «Предмет · N класс · Урок 16»
    for i, el in enumerate(tags):
        if text_of(el).strip().startswith('Модуль'):
            shift = (bbox(tags[i + 1])[0] - bbox(el)[0]) if i + 1 < len(tags) else 0
            for nxt in tags[i + 1:]:
                b = bbox(nxt); b[0] -= shift; set_bbox(nxt, b)
            el.getparent().remove(el)
            break
    # предмет: конвертер узнаёт тег по тексту «Русский язык»/«История» и ставит предмет урока
    if tags:
        first = tags[0].find('.//a:r/a:t', NS)
        first.text = 'Русский язык'
        for r in tags[0].findall('.//a:r', NS)[1:]: r.getparent().remove(r)


def clean_divider(s):
    """Описание-заглушку «Добавь описание…» оставляем: конвертер подставит описание из исходника или уберёт её."""


def build(profile_dir):
    profile = json.loads((profile_dir / 'profile.json').read_text('utf8'))
    r = profile.get('base_recipe')
    if not r: return None
    out = ROOT / r['output']
    inputs = [ROOT / r['template'], profile_dir / 'profile.json', Path(__file__)] + [ROOT / f for f, _ in r['append']]
    if out.exists() and out.stat().st_mtime >= max(p.stat().st_mtime for p in inputs):
        return out  # база свежая
    prs = Presentation(ROOT / r['template'])
    slides = list(prs.slides)
    for n in r.get('titles', []): clean_title(slides[n - 1], prs.slide_height)
    for n in r.get('dividers', []): clean_divider(slides[n - 1])
    layouts = layouts_by_name(prs)
    n = len(slides)
    for f, nums in r['append']:
        lesson = Presentation(ROOT / f)
        src, vs = list(lesson.slides), prs.slide_height / lesson.slide_height
        for k in nums:
            append_slide(prs, layouts, src[k - 1], vs); n += 1
            print('  слайд %d ← %s, слайд %d' % (n, f, k))
    out.parent.mkdir(parents=True, exist_ok=True)
    tmp = out.with_suffix('.tmp.pptx')
    prs.save(tmp); tmp.replace(out)
    print('база собрана:', out.relative_to(ROOT), '| слайдов:', n)
    return out


if __name__ == '__main__':
    dirs = [Path(a) for a in sys.argv[1:]] or sorted(p.parent for p in (ROOT / 'templates').glob('*/profile.json'))
    for d in dirs: build(d.resolve())
