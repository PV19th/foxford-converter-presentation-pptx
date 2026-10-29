# -*- coding: utf-8 -*-
import sys, copy, re, json, math, os
from lxml import etree
from pptx import Presentation
from engine import *

IN = int(EMU)
GRADE = '5'
DONOR = dict(title=1, plan=2, wait=4, summary=10, reflection=31, links=33, bye=32, dividers=[3, 11, 18],
             content=[6, 22, 26, 8])
if os.environ.get('CONV_DONOR'):
    DONOR.update(json.loads(os.environ['CONV_DONOR']))

# Старый шаблон 10×7,5″ (уроки 6–11 классов разных предметов): слайд того же размера, что новый, кнопки
# «Вернуться к навигации», ярлыки «Теория» / «Групповая работа» внизу слайда, служебные слайды до титула.
# Правила для него включаются флагом WIDE; старый шаблон 5–6 классов (8,27×5,9″) переносится как раньше.
WIDE = False
TITLE_N = 1  # номер титульного слайда в исходнике


# ---------------------------------------------------------------- стикеры
def sticker(name):
    """Путь к стикеру из assets/ или None, если файла нет.
    Без файла стикер не вставляется: место под него остаётся, а в отчёт по слайду
    пишется «добавить вручную»."""
    path = ASSETS + '/' + name
    return path if os.path.exists(path) else None


# ---------------------------------------------------------------- классификация
def slide_title_info(s):
    """Возвращает (текст заголовка, [элементы-заголовки])."""
    cands = []
    for el in s.shapes._spTree:
        if etree.QName(el).localname != 'sp': continue
        b = bbox(el)
        ph = el.find('p:nvSpPr/p:nvPr/p:ph', NS)
        is_ph_title = ph is not None and ph.get('type') in ('title',) and ph.get('idx') in ((None, '0', '4294967295') if WIDE else (None, '0'))
        if b is None:
            if is_ph_title: cands.append((0, el))
            continue
        if b[0] >= SRC_W: continue
        if b[1] < int(0.55 * IN) and b[3] < int((1.4 if WIDE else 1.0) * IN):
            szs = [int(r.get('sz')) for r in el.iter(A + 'rPr') if r.get('sz')]
            if is_ph_title or (szs and max(szs) >= 2000) or (not szs and not text_of(el).strip()):
                if el.find('p:spPr/a:solidFill', NS) is None:
                    cands.append((b[1], el))
    txt = ''
    for _, el in sorted(cands, key=lambda x: x[0]):
        t = para_text(el).strip()
        if t: txt = t; break
    return txt, [e for _, e in cands]


def para_text(el):
    out = []
    for p in el.iter(A + 'p'):
        s = ''
        for r in p:
            ln = etree.QName(r).localname
            if ln in ('r', 'fld'): s += r.findtext('a:t', '', NS)
            elif ln == 'br': s += '\n'
        out.append(s)
    while out and not out[-1].strip(): out.pop()
    return '\n'.join(out)


def all_text(s):
    return ' '.join(text_of(el) for el in s.shapes._spTree)


def classify(s, i):
    title, _ = slide_title_info(s)
    hs = set()
    for pic in s._element.iter(P + 'pic'):
        h = img_hash(s, pic)
        if h: hs.add(h)
    t = all_text(s)
    if i == TITLE_N: return 'title'
    k = classify_service(s, title, t)
    if k: return k
    if WIDE:
        k = classify_wide(s, title, t)
        if k: return k
    if 'УВИДИМСЯ' in t: return 'bye'
    if 'faa3488e67' in hs or '05467b9d2c' in hs: return 'divider'
    tl = title.lower()
    if tl.startswith('план урока'): return 'plan'
    if tl.startswith('что нас ждёт') or tl.startswith('что нас ждет'): return 'wait'
    if tl.startswith('ссылки'): return 'links'
    circles = [el for el in s.shapes._spTree if etree.QName(el).localname == 'sp' and text_of(el).strip() == '5']
    if len(circles) >= 2: return 'reflection'
    if tl.startswith('подведём итоги') or tl.startswith('подведем итоги'): return 'summary'
    if DONOR.get('why'):
        top = src_items_by_y(s, lambda ne, el: len(text_of(ne).strip()) > 3 and 'навигации' not in text_of(ne))
        if top and re.match(r'(почему стоит посмотреть|зачем смотреть)', text_of(top[0][2]).strip().lower()): return 'why'
    if DONOR.get('think') and (tl.startswith('подумаем') or tl.startswith('работа с картой')):
        pics = [p for p in s._element.iter(P + 'pic') if img_hash(s, p) not in DECOR_SKIP | DECOR_Q | DECOR_VAZHNO
                and not is_offslide(bbox(p))]
        ar = sorted(bbox(p)[2] * bbox(p)[3] for p in pics)
        if len(pics) == 1 or (len(pics) == 2 and ar[0] > 0.6 * ar[1]): return 'think'
    return 'content'


# ---------------------------------------------------------------- старый шаблон 10×7,5″
WIDE_PLAN = ('план урока', 'что обсудим', 'о чем поговорим', 'о чём поговорим', 'сегодня на уроке',
             'by the end of this lesson')


def find_title_slide(slides):
    """Титул — первый слайд с тегом «N класс»; до него бывают служебные слайды (орг. моменты, «Как дела с ДЗ»).
    У уроков ЕГЭ и интенсивов тега нет — тогда первый слайд на титульном макете («TITLE_1…», «Титульный слайд»)."""
    for i, s in enumerate(slides[:8], 1):
        if any(re.fullmatch(r'\d{1,2}\s*класс', text_of(el).strip()) for el in s.shapes._spTree.iter(P + 'sp')):
            return i
    for i, s in enumerate(slides[:8], 1):
        name = s.slide_layout.name
        if name.startswith('TITLE') and not name.startswith('TITLE_ONLY') or 'Титульн' in name: return i
    for i, s in enumerate(slides[:8], 1):
        if get_ph(s, 'ctrTitle') is not None: return i
    return 1


def is_footer(b):
    """Низ слайда старого шаблона 10×7,5″ (и его версии 16:9): ярлык типа задания («Теория», «Групповая работа»)
    с иконкой, номер слайда. В новом шаблоне таких ярлыков нет — не переносим."""
    return WIDE and b is not None and b[1] >= int(SRC_H * 6.5 / 7.5)


# Служебные слайды, которые пользователь может заменить слайдом шаблона (выбор на сайте: ключи CONV_DONOR).
# Не выбрано — слайд переносится из старой презентации как обычный.
SERVICE = [('lesson_end', r'как прош[её]л урок|как прошло занятие|что было самым сложным|настроение после урока|how was the lesson'),
           ('homework', r'как дела с домашн|how is your homework'),
           ('mood', r'настроение (сегодня|перед)|^как (твоё |твое )?настроение|how are you'),
           ('rating', r'как вы оцениваете урок|как вам урок|как тебе урок|help us improve')]
SERVICE_KINDS = [k for k, _ in SERVICE] + ['motivation']


def classify_service(s, title, t):
    for kind, rx in SERVICE:
        # у «Как вы оцениваете урок?» крупная надпись не в шапке слайда — ищем по всему тексту
        where = re.sub(r'\s+', ' ', t if kind == 'rating' else title).strip().lower()
        if DONOR.get(kind) and re.search(rx, where): return kind
    if WIDE and DONOR.get('motivation') and is_motivation(s, title): return 'motivation'
    return None


def is_motivation(s, title):
    """Мотивашка старого шаблона 10×7,5″: лозунг без заголовка вверху («Practice makes perfect!», «Ошибки — это часть
    процесса…»): одна-две надписи, крупная (или любая, если крупной нет) кончается «!»; не перерыв и не прощание."""
    if title: return False
    items = src_items_by_y(s, lambda ne, el: text_of(ne).strip() and 'навигаци' not in text_of(ne).lower())
    texts = [re.sub(r'\s+', ' ', text_of(x[2])).strip() for x in items]
    if not texts or len(texts) > 2 or sum(map(len, texts)) > 200: return False
    if any(x[3].find('.//a:hlinkClick', NS) is not None for x in items): return False
    head = get_ph(s, 'ctrTitle')
    heads = [re.sub(r'\s+', ' ', text_of(head)).strip()] if head is not None and text_of(head).strip() else texts
    return any(h.endswith('!') for h in heads) and not re.search(r'break|перемен|разминк|увидимся|до встречи|bye', ' '.join(texts).lower())


def classify_wide(s, title, t):
    tu = t.upper()
    if re.search(r'УВИДИМСЯ|ДО ВСТРЕЧИ|ДО СВИДАНИЯ|BYE FOR NOW', tu): return 'bye'
    if any(text_of(el).strip().lower().startswith(('использованные источники', 'использованные ресурсы'))
           for el in s.shapes._spTree.iter(P + 'sp')): return 'links'
    tl = title.lower()
    if tl.startswith(WIDE_PLAN): return 'plan' if len(plan_rows(s)) >= 2 else 'content'
    if tl.startswith(('обобщ', 'итоги урока')): return 'summary'
    if not title and wide_divider(s): return 'divider'
    return None


def title_text(src):
    """Тема урока на титуле: заголовок титула, а если его нет (уроки ЕГЭ — тема обычной надписью) —
    самый крупный текст слайда."""
    t = ''
    for el in src.shapes._spTree.iter(P + 'sp'):
        ph = el.find('p:nvSpPr/p:nvPr/p:ph', NS)
        if ph is not None and ph.get('type') == 'ctrTitle': t = para_text(el).strip()
    if not t:
        big = [(max(szs), para_text(el).strip()) for el in src.shapes._spTree.iter(P + 'sp')
               if (szs := [int(r.get('sz')) for r in el.iter(A + 'rPr') if r.get('sz')]) and para_text(el).strip()]
        if big: t = max(big)[1]
    return t


def wide_notes(src, i, kind):
    """Что методисту проверить на слайде старого шаблона 10×7,5″ (строки отчёта)."""
    notes = []
    if i < TITLE_N: notes.append('в старой презентации слайд стоял до титула — проверьте, нужен ли он здесь')
    if kind == 'title':
        sps = [(el, bbox(materialize_sp(src, el))) for el in src.shapes._spTree.iter(P + 'sp')]
        grade = next((b for el, b in sps if b and re.fullmatch(r'\d{1,2}\s*класс', text_of(el).strip())), None)
        rest, theme = [], title_text(src)
        for el, b in sps:
            t = re.sub(r'\s+', ' ', para_text(el).replace('\n', ' ')).strip()
            if not t or el is get_ph(src, 'ctrTitle') or para_text(el).strip() == theme or 'навигации' in t or is_offslide(b): continue
            if grade and b and abs(b[1] - grade[1]) < 0.15 * IN: continue  # теги «предмет · класс · урок»
            rest.append(t[:80])
        if rest: notes.append('текст со старого титула не перенесён: «%s»' % '»; «'.join(rest))
    return notes


def max_sz(el):
    return max([int(r.get('sz')) for r in el.iter(A + 'rPr') if r.get('sz')] + [0])


def wide_divider(s):
    """Отбивка раздела: крупная надпись без заголовка вверху и почти без текста («Vocabulary», «Обобщение»,
    «ПЕРЕМЕНА»). → (название, описание) или None."""
    items = src_items_by_y(s, lambda ne, el: text_of(ne).strip() and 'навигации' not in text_of(ne))
    if not items or len(items) > 2: return None
    head = get_ph(s, 'ctrTitle')
    if head is None or not text_of(head).strip():
        big = [x for x in items if max_sz(x[3]) >= 2800]
        if not big: return None
        head = big[0][3]
    rest = [x for x in items if x[3] is not head]
    if rest and (len(text_of(rest[0][2])) > 200 or rest[0][3].find('.//a:hlinkClick', NS) is not None): return None
    one = lambda el: re.sub(r'\s+', ' ', para_text(el).replace('\n', ' ')).strip()
    return one(head), (one(rest[0][2]) if rest else '')


def plan_rows(src):
    """Пункты плана: вторая колонка таблицы; в старом шаблоне 10×7,5″ — ещё и текст рядом с номерами «01», «1»."""
    rows = []
    for gf in src._element.iter(P + 'graphicFrame'):
        for tr in gf.iter(A + 'tr'):
            cells = tr.findall('a:tc', NS)
            texts = [text_of(c).strip() for c in cells]
            if len(texts) >= 2 and texts[1]:
                rows.append(para_text(cells[1]).replace('\n', ' ').strip())
    if rows or not WIDE: return rows
    _, tels = slide_title_info(src)
    its = src_items_by_y(src, lambda ne, el: el not in tels and text_of(ne).strip() and 'навигации' not in text_of(ne))
    nums = [x for x in its if re.fullmatch(r'\d{1,2}', text_of(x[2]).strip())]
    if len(nums) < 2: return []
    return [re.sub(r'\s+', ' ', para_text(x[2]).replace('\n', ' ')).strip() for x in its if x not in nums]


# ---------------------------------------------------------------- материализация текста
def resolve_text(src, sp_src, ne, smap, role='body', summary=False, keep_br=False):
    """Переписывает txBody у копии ne: явные размеры/цвета/шрифты по правилам нового шаблона."""
    tb = ne.find('p:txBody', NS) if etree.QName(ne).localname != 'tc' else ne.find('a:txBody', NS)
    if tb is None: return
    lsts, bps, _ = style_chain(src, sp_src) if sp_src is not None else ([], [], [])
    # bodyPr
    bp = tb.find('a:bodyPr', NS)
    if bps:
        m = merge_el([b for b in bps if b is not None] + [bp], [], A + 'bodyPr')
        for k, v in m.attrib.items():
            if bp.get(k) is None: bp.set(k, v)
    for c in list(bp):
        if etree.QName(c).localname in ('normAutofit', 'spAutoFit', 'noAutofit'): bp.remove(c)
    fill_hex = None
    for p in tb.findall('a:p', NS):
        pPr = p.find('a:pPr', NS)
        lvl = int(pPr.get('lvl', 0)) if pPr is not None else 0
        chain = [lvl_of(l, lvl) for l in lsts] + [pPr]
        eff = merge_el([c for c in chain if c is not None], PPR_GROUPS, A + 'pPr')
        defs = [c.find('a:defRPr', NS) for c in chain[:-1] if c is not None]
        defs = [d for d in defs if d is not None]
        runs = []
        for r in list(p):
            ln = etree.QName(r).localname
            if ln in ('r', 'fld', 'br') or ln == 'endParaRPr':
                rp = r.find('a:rPr', NS) if ln != 'endParaRPr' else r
                effr = merge_el(defs + ([rp] if rp is not None else []), RPR_GROUPS, A + 'rPr')
                runs.append((r, ln, effr))
        # итоговые свойства
        texts = []
        for r, ln, effr in runs:
            sz = int(effr.get('sz', '1400')) / 100
            b = effr.get('b') == '1'
            col = color_of(effr.find('a:solidFill', NS), smap)
            texts.append([r, ln, effr, sz, b, col])
        # правило «жирный — только термины»: длинные жирные отрезки снимаем
        if not summary:
            seg = []
            def flush(seg):
                words = len(' '.join(x[0].findtext('a:t', '', NS) for x in seg).split())
                if words > 5:
                    for x in seg: x[4] = False
            for item in texts:
                r, ln, effr, sz, b, col = item
                if ln not in ('r',): continue
                t = r.findtext('a:t', '', NS)
                if b: seg.append(item)
                elif t.strip():
                    flush(seg); seg = []
            flush(seg)
        new_ppr = etree.Element(A + 'pPr')
        for k in ('lvl', 'algn', 'marL', 'indent', 'rtl'):
            if eff.get(k) is not None: new_ppr.set(k, eff.get(k))
        for g in ('lnSpc', 'spcBef', 'spcAft'):
            e = eff.find('a:' + g, NS)
            if e is not None: new_ppr.append(copy.deepcopy(e))
        bu = None
        for g in ('buChar', 'buAutoNum', 'buNone', 'buBlip'):
            e = eff.find('a:' + g, NS)
            if e is not None: bu = e
        first_col = None
        first_sz = None
        for item in texts:
            if item[1] == 'r' and item[0].findtext('a:t', '', NS).strip():
                first_col = item[5]; first_sz = item[3]; break
        tcol = map_text(first_col) if first_col else '0F1115'
        has_text = any(it[1] == 'r' and it[0].findtext('a:t', '', NS).strip() for it in texts)
        if bu is not None and etree.QName(bu).localname in ('buChar', 'buBlip') and has_text:
            c = etree.SubElement(new_ppr, A + 'buClr'); c.append(solid(tcol if tcol == 'FFFFFF' else '0F1115')[0])
            etree.SubElement(new_ppr, A + 'buSzPct').set('val', '100000')
            etree.SubElement(new_ppr, A + 'buFont').set('typeface', 'Arial')
            etree.SubElement(new_ppr, A + 'buChar').set('char', '•')
            new_ppr.set('marL', str(342900 + 342900 * lvl)); new_ppr.set('indent', '-228600')
        elif bu is not None and etree.QName(bu).localname == 'buAutoNum' and has_text:
            c = etree.SubElement(new_ppr, A + 'buClr'); c.append(solid('0F1115')[0])
            etree.SubElement(new_ppr, A + 'buSzPct').set('val', '100000')
            etree.SubElement(new_ppr, A + 'buFont').set('typeface', '+mj-lt')
            new_ppr.append(copy.deepcopy(bu))
            new_ppr.set('marL', str(342900 + 342900 * lvl)); new_ppr.set('indent', '-342900')
        else:
            etree.SubElement(new_ppr, A + 'buNone')
            if new_ppr.get('indent') and int(new_ppr.get('indent')) < 0:
                new_ppr.set('indent', '0')
            if new_ppr.get('marL') and int(new_ppr.get('marL')) > 0 and bu is None:
                pass
        if pPr is not None: p.remove(pPr)
        p.insert(0, new_ppr)
        for item in texts:
            r, ln, effr, sz, b, col = item
            nsz = map_size(sz, role)
            ncol = map_text(col) if col else '0F1115'
            nr = new_rpr(int(nsz * 100), b=b, color=ncol, i=effr.get('i') == '1',
                         u=effr.get('u') if effr.get('u') not in (None, 'none') else None)
            if effr.get('baseline'): nr.set('baseline', effr.get('baseline'))
            hl = effr.find('a:hlinkClick', NS)
            if hl is not None:
                nr.append(copy.deepcopy(hl))
            if ln == 'endParaRPr':
                nr.tag = A + 'endParaRPr'
                p.replace(r, nr)
            else:
                old = r.find('a:rPr', NS)
                if old is not None: r.replace(old, nr)
                else: r.insert(0, nr)
        if new_ppr.get('algn') in ('just', 'dist', 'justLow', 'thaiDist'): new_ppr.set('algn', 'l')
        if not has_text:
            for rp in p.iter(A + 'rPr', A + 'endParaRPr'):
                orig = [it[3] for it in texts]
                v = min(min(orig) if orig else 14, 14)
                rp.set('sz', '1400' if v >= 13 else '1200')  # пустые абзацы тоже без «дробных» кеглей
        if not keep_br:
            for br in p.findall('a:br', NS):
                prev = br.getprevious()
                while prev is not None and etree.QName(prev).localname != 'r': prev = prev.getprevious()
                nxt = br.getnext()
                if prev is not None:
                    t = prev.find('a:t', NS)
                    nt = nxt.findtext('a:t', '', NS) if nxt is not None and etree.QName(nxt).localname == 'r' else ''
                    if not (t.text or '').endswith(' ') and not nt.startswith(' '): t.text = (t.text or '') + ' '
                p.remove(br)
    ps = tb.findall('a:p', NS)
    nonempty = [p for p in ps if ''.join(t.text or '' for t in p.iter(A + 't')).strip()]
    if nonempty:
        i0, i1 = ps.index(nonempty[0]), ps.index(nonempty[-1])
        for p in ps[:i0] + ps[i1 + 1:]: tb.remove(p)
    return tb


PLATE_INS = 90000  # поля плашки с текстом: 0,25 см со всех сторон (правило чек-листа, п. 14)


def pad_plate(sp):
    """Плашка с текстом: внутренние поля 0,25 см слева, справа, сверху и снизу."""
    bp = sp.find('p:txBody/a:bodyPr', NS)
    if bp is None or not text_of(sp).strip(): return
    for k in ('lIns', 'tIns', 'rIns', 'bIns'): bp.set(k, str(PLATE_INS))


def map_size(sz, role):
    if role == 'caption': return 12
    if role in ('number', 'card'): return 14
    return 16


# ---------------------------------------------------------------- перенос контентного слайда
def is_offslide(b):
    return b is None or b[0] >= SRC_W - 10000 or b[0] + b[2] <= 0 or b[1] >= SRC_H


def materialize_sp(src, sp):
    """Копия фигуры с унаследованными от макета xfrm/geom/fill/ln; без ph."""
    ne = copy.deepcopy(sp)
    ph = ne.find('p:nvSpPr/p:nvPr/p:ph', NS)
    if ph is not None:
        _, _, sps = style_chain(src, sp)
        spPr = ne.find('p:spPr', NS)
        for s in sps:  # master, layout
            if s is None: continue
            for tag in ('a:xfrm', 'a:prstGeom', 'a:custGeom'):
                if spPr.find(tag, NS) is None and s.find(tag, NS) is not None:
                    if tag == 'a:xfrm': spPr.insert(0, copy.deepcopy(s.find(tag, NS)))
                    elif spPr.find('a:prstGeom', NS) is None and spPr.find('a:custGeom', NS) is None:
                        spPr.append(copy.deepcopy(s.find(tag, NS)))
            fills = [c for c in spPr if etree.QName(c).localname in ('noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill')]
            sfills = [c for c in s if etree.QName(c).localname in ('noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill')]
            if not fills and sfills: insert_spPr_child(spPr, copy.deepcopy(sfills[0]))
            if spPr.find('a:ln', NS) is None and s.find('a:ln', NS) is not None:
                insert_spPr_child(spPr, copy.deepcopy(s.find('a:ln', NS)))
        ph.getparent().remove(ph)
    return ne


SPPR_ORDER = ['xfrm', 'custGeom', 'prstGeom', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill', 'ln',
              'effectLst', 'effectDag', 'scene3d', 'sp3d', 'extLst']


def insert_spPr_child(spPr, child):
    ln = etree.QName(child).localname
    oi = SPPR_ORDER.index(ln)
    for i, c in enumerate(spPr):
        cl = etree.QName(c).localname
        if cl in SPPR_ORDER and SPPR_ORDER.index(cl) > oi:
            spPr.insert(i, child); return
    spPr.append(child)


def restyle_spPr(spPr, smap, caption=False):
    if spPr is None: return None
    fill_hex = None
    for c in list(spPr):
        ln = etree.QName(c).localname
        if ln == 'solidFill':
            h = color_of(c, smap)
            if caption: spPr.replace(c, etree.Element(A + 'noFill')); continue
            fill_hex = map_fill(h); spPr.replace(c, solid(fill_hex))
        elif ln == 'gradFill':
            st = c.find('a:gsLst/a:gs', NS)
            h = color_of(st, smap) if st is not None else 'FFFFFF'
            fill_hex = map_fill(h); spPr.replace(c, solid(fill_hex) if not caption else etree.Element(A + 'noFill'))
        elif ln == 'ln':
            f = c.find('a:solidFill', NS)
            if f is not None:
                h = color_of(f, smap)
                if caption: c.replace(f, etree.Element(A + 'noFill'))
                else: c.replace(f, solid(map_line(h)))
        elif ln in ('effectLst',):
            spPr.remove(c)
    return fill_hex


def renumber(el, off=5000):
    for e in el.iter(P + 'cNvPr'):
        e.set('id', str(int(e.get('id')) + off))
    for e in el.iter(A + 'stCxn', A + 'endCxn'):
        e.set('id', str(int(e.get('id')) + off))


class Item:
    def __init__(self, el, kind, role=''):
        self.el, self.kind, self.role = el, kind, role
        self.caption = None   # подпись, привязанная к картинке
        self.host = None

    @property
    def b(self): return bbox(self.el)


def transform_box(b, sy0, k=None, clamp=True):
    k = K if k is None else k
    x, y, w, h = b
    nx, ny, nw, nh = TW / 2 + (x - SRC_W / 2) * k, CT + (y - sy0) * k, w * k, h * k
    if not clamp: return [nx, ny, nw, nh]
    if x <= int(0.12 * IN) and x + w >= SRC_W - int(0.12 * IN):  # во всю ширину → в поля
        nx, nw = LM, RM - LM
    if nw > RM - LM: nx, nw = LM, RM - LM
    if nx < LM - 20000: nx = LM
    if nx + nw > RM + 20000: nx = RM - nw
    return [nx, ny, nw, nh]


def intersects(a, b, m=0):
    return not (a[0] + a[2] <= b[0] - m or b[0] + b[2] <= a[0] - m or a[1] + a[3] <= b[1] - m or b[1] + b[3] <= a[1] - m)


def area_int(a, b):
    w = min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0]); h = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
    return max(0, w) * max(0, h)


def scale_group(el, sx0_unused, sy0):
    x = el.find('p:grpSpPr/a:xfrm', NS)
    b = bbox(el)
    set_bbox(el, transform_box(b, sy0))


def transfer_content(bld, src, dst, report):
    global CT
    smap = scheme_map(src)
    title, title_els = slide_title_info(src)
    nl = sum(max(1, math.ceil(text_width_pt(l, 24, True) * 1.2 / 655)) for l in (title or ' ').split('\n'))
    CT = int(0.39 * IN * VS + nl * 24 * 1.25 * 12700 + 0.2 * IN * VS)
    tree = dst.shapes._spTree
    items = []
    badges, qmarks = [], []
    src_els = [el for el in src.shapes._spTree][2:]
    tops = []
    for el in src_els:
        if el in title_els: continue
        b = bbox(el)
        if b is None and etree.QName(el).localname == 'sp' and el.find('p:nvSpPr/p:nvPr/p:ph', NS) is not None:
            b = bbox(materialize_sp(src, el))
        if is_offslide(b) or is_footer(b): continue
        tops.append(b[1])
    sy0 = min([int(0.75 * IN)] + tops) if tops else int(0.75 * IN)
    pics_src = []
    for el in src_els:
        if el in title_els: continue
        tag = etree.QName(el).localname
        if tag == 'sp':
            php = el.find('p:nvSpPr/p:nvPr/p:ph', NS)
            if php is not None and php.get('type') in ('sldNum', 'dt', 'ftr'): continue
            ne = materialize_sp(src, el)
            b = bbox(ne)
            if is_offslide(b) or is_footer(b): continue
            txt = text_of(ne).strip()
            if 'ЭТО ВАЖНО' in txt.upper() and len(txt) < 15:
                badges.append(list(b)); continue
            spPr = ne.find('p:spPr', NS)
            has_fill = spPr.find('a:solidFill', NS) is not None or spPr.find('a:gradFill', NS) is not None
            ln = spPr.find('a:ln', NS)
            has_ln = ln is not None and ln.find('a:noFill', NS) is None and (ln.find('a:solidFill', NS) is not None)
            geom = spPr.find('a:prstGeom', NS)
            if not txt and not has_fill and not has_ln and spPr.find('a:custGeom', NS) is None: continue
            if not txt and b[1] < 0.7 * IN: continue
            if not txt and has_fill and not has_ln:
                fh = color_of(spPr.find('a:solidFill', NS), smap) if spPr.find('a:solidFill', NS) is not None else None
                if fh == 'FFFFFF': continue
            if re.fullmatch(r'\d{1,3}', txt) and b[1] > 4.9 * IN and not has_fill and not has_ln: continue
            items.append(Item(ne, 'sp')); items[-1].src = el
        elif tag == 'pic':
            h = img_hash(src, el)
            b = bbox(el)
            if is_offslide(b) or is_footer(b): continue
            if h in DECOR_SKIP: continue
            if h in DECOR_VAZHNO: badges.append(list(b)); continue
            if h in DECOR_Q: qmarks.append(list(b)); continue
            if is_flat(src, el): continue  # одноцветная заглушка-маска
            if WIDE and b[1] <= 0.1 * IN and b[2] >= 0.95 * SRC_W and b[3] <= 2 * IN: continue  # полоса-шапка старого шаблона (9.23 сл. 1)
            ne = copy.deepcopy(el)
            items.append(Item(ne, 'pic')); items[-1].blank_bottom = blank_bottom_frac(src, el)
        elif tag == 'cxnSp':
            ne = copy.deepcopy(el)
            if is_offslide(bbox(ne)) or is_footer(bbox(ne)): continue
            items.append(Item(ne, 'cxn'))
        elif tag == 'grpSp':
            ne = copy.deepcopy(el)
            if is_offslide(bbox(ne)) or is_footer(bbox(ne)): continue
            kids = [c for c in ne.iter(P + 'sp', P + 'pic', P + 'cxnSp')]
            if len(kids) >= 5 and all(etree.QName(c).localname == 'sp' and not text_of(c).strip() for c in kids): continue
            items.append(Item(ne, 'grp')); items[-1].src = el
        elif tag == 'graphicFrame':
            ne = copy.deepcopy(el)
            if is_offslide(bbox(ne)): continue
            items.append(Item(ne, 'tbl'))
    # подписи: мелкий текст рядом с картинкой
    pics = [it for it in items if it.kind == 'pic']
    for it in items:
        if it.kind != 'sp': continue
        szs = [int(r.get('sz')) for r in it.el.iter(A + 'rPr') if r.get('sz')]
        txt = text_of(it.el).strip()
        if not txt or len(txt) > 260: continue
        b = it.b
        small = szs and max(szs) <= 1200
        best = None
        for pi in pics:
            pb = pi.b
            ov = min(b[0] + b[2], pb[0] + pb[2]) - max(b[0], pb[0])
            if ov < 0.4 * min(b[2], pb[2]): continue
            gap_below = b[1] - (pb[1] + pb[3])
            gap_above = pb[1] - (b[1] + b[3])
            inside_bottom = b[1] >= pb[1] + pb[3] * 0.6 and b[1] + b[3] <= pb[1] + pb[3] + 0.35 * IN
            if -0.4 * IN <= gap_below <= 0.35 * IN or inside_bottom:
                d = 0 if inside_bottom else abs(gap_below)
                if best is None or d < best[0]: best = (d, pi, 'below')
            elif 0 <= gap_above <= 0.25 * IN and small:
                if best is None or gap_above < best[0]: best = (gap_above, pi, 'above')
        has_fill = it.el.find('p:spPr/a:solidFill', NS) is not None
        if not best and small and len(txt) < 120:
            for pi in pics:
                pb = pi.b
                if 0 <= b[0] - (pb[0] + pb[2]) <= 0.3 * IN and pb[1] <= b[1] <= pb[1] + pb[3]:
                    it.role = 'caption'
        # подпись слева от нижнего края картинки (текст прижат вправо к картинке) — переносим под картинку
        if not best and small and len(txt) < 140 and not it.role and not has_fill:
            for pi in pics:
                pb = pi.b
                ov = min(b[0] + b[2], pb[0] + pb[2]) - max(b[0], pb[0])
                if -0.2 * IN <= pb[0] - (b[0] + b[2]) <= 0.3 * IN and ov < 0.4 * min(b[2], pb[2]) \
                        and pb[1] + 0.5 * pb[3] <= b[1] <= pb[1] + pb[3] + 0.1 * IN:
                    best = (0, pi, 'below'); it.cap_left = True; break
        # надпись-плашка, лежащая внутри картинки, — это подпись-метка на изображении, а не подпись под ним
        if best and has_fill and b[2] < 0.6 * best[1].b[2] and b[1] + b[3] <= best[1].b[1] + best[1].b[3] + 0.05 * IN \
                and b[0] >= best[1].b[0] - 0.05 * IN and b[0] + b[2] <= best[1].b[0] + best[1].b[2] + 0.05 * IN:
            best = None
        # узкая плашка, начинающаяся внутри картинки (даже если вылезает за её край), — тоже метка на изображении
        if best and has_fill and b[2] < 0.6 * best[1].b[2] and b[1] < best[1].b[1] + best[1].b[3] - 0.1 * IN:
            best = None
        if best and (small or (len(txt) < 70 and szs and max(szs) <= 1400 and not (has_fill and b[2] > 0.8 * SRC_W))
                     or (len(txt) < 140 and b[2] <= best[1].b[2] * 1.3 and (not szs or max(szs) <= 1400) and best[2] == 'below')):
            it.role = 'caption'; it.host = best[1]
            prev = best[1].caption
            if prev is None or best[0] < getattr(prev, 'cap_d', 1e18):
                best[1].caption = it
            it.cap_d = best[0]
    # у картинки остаётся одна подпись — ближайшая; остальные кандидаты — обычный текст
    for it in items:
        if it.role == 'caption' and it.host is not None and it.host.caption is not it:
            it.role = ''; it.host = None
    # стиль + геометрия
    for it in items:
        el = it.el
        it.sb = list(it.b)
        if it.kind == 'sp':
            fill = restyle_spPr(el.find('p:spPr', NS), smap, caption=(it.role == 'caption'))
            txtlen = len(text_of(el).strip())
            b = it.b
            spPr0 = el.find('p:spPr', NS)
            ln0 = spPr0.find('a:ln', NS)
            outlined = ln0 is not None and ln0.find('a:solidFill', NS) is not None
            filled = spPr0.find('a:solidFill', NS) is not None
            role = it.role or ('number' if txtlen <= 2 and b[2] < 0.8 * IN and b[3] < 0.8 * IN else
                               'card' if (outlined or filled) and b[2] < 0.45 * SRC_W else 'body')
            it.role = it.role or role
            resolve_text(src, it.src, el, smap, role=role)
            if getattr(it, 'cap_left', False) and it.role == 'caption':
                for pp in el.iter(A + 'pPr'): pp.set('algn', 'ctr')
            if it.role == 'caption' and it.host is not None:
                # подпись прижимается к картинке: под картинкой — текст сверху блока, над картинкой — снизу (5.21 сл. 6)
                bp_ = el.find('p:txBody/a:bodyPr', NS)
                if bp_ is not None: bp_.set('anchor', 't' if it.sb[1] >= it.host.b[1] + it.host.b[3] * 0.5 else 'b')
            if fill and it.role != 'caption': pad_plate(el)
            if WIDE and it.role == 'caption':  # подпись теряет плашку — белый текст стал бы невидим
                for f in el.iter(A + 'srgbClr'):
                    if f.get('val') == 'FFFFFF' and f.getparent().getparent().tag in (A + 'rPr', A + 'endParaRPr'): f.set('val', '0F1115')
            if fill and is_dark(fill):
                for rp in el.iter(A + 'rPr', A + 'endParaRPr'):
                    f = rp.find('a:solidFill/a:srgbClr', NS)
                    if f is not None and f.get('val') == '0F1115': f.set('val', 'FFFFFF')
        elif it.kind == 'pic':
            restyle_spPr(el.find('p:spPr', NS), smap)
        elif it.kind == 'cxn':
            restyle_spPr(el.find('p:spPr', NS), smap)
        elif it.kind == 'grp':
            for sp in el.iter(P + 'sp'):
                restyle_spPr(sp.find('p:spPr', NS), smap)
                if sp.find('p:txBody', NS) is not None:
                    resolve_text(src, None, sp, smap, role='card')
                    if sp.find('p:spPr/a:solidFill', NS) is not None: pad_plate(sp)
            for c in el.iter(P + 'cxnSp', P + 'pic'):
                restyle_spPr(c.find('p:spPr', NS), smap)
        elif it.kind == 'tbl':
            for tc in el.iter(A + 'tc'):
                resolve_text(src, None, tc, smap, role='card')
                tcPr = tc.find('a:tcPr', NS)
                if tcPr is not None:
                    for f in tcPr.iter(A + 'solidFill'):
                        h = color_of(f, smap); f.getparent().replace(f, solid(map_fill(h)))
            it.grid = [int(g.get('w')) for g in el.iter(A + 'gridCol')]
            it.rows = [int(tr.get('h')) for tr in el.iter(A + 'tr')]
        renumber(el)
    if WIDE: unhide_white_text(items); lift_tabs(items)
    # верх контента считаем только по перенесённым элементам (пустые белые плашки и декор не в счёт)
    tops2 = [it.sb[1] for it in items] + [bb[1] for bb in badges + qmarks]
    if tops2: sy0 = min([int(0.75 * IN)] + tops2)
    # ---- раскладка с подбором масштаба
    zone = CB - CT
    k = K
    plan = [(16, 1.0), (16, 0.96), (16, 0.92), (16, 0.88)] + [(14, K_) for K_ in (1.0, 0.96, 0.92, 0.88, 0.84, 0.80, 0.76, 0.72)]
    mode16 = True
    for fsz, kk in plan:
        k = K * kk
        if fsz == 14 and mode16:
            mode16 = False
            for it in items:
                if it.kind == 'sp' and it.role == 'body':
                    for rp in it.el.iter(A + 'rPr', A + 'endParaRPr'):
                        if rp.get('sz') == '1600': rp.set('sz', '1400')
                    for bs_ in it.el.iter(A + 'buSzPts'): pass
        for it in items:
            if WIDE: set_bbox(it.el, wide_box(it, sy0, k))
            elif it.kind == 'cxn': set_bbox(it.el, transform_box(it.sb, sy0, k, clamp=False))
            else: set_bbox(it.el, transform_box(it.sb, sy0, k))
            if it.kind == 'tbl':
                for g, w in zip(it.el.iter(A + 'gridCol'), it.grid): g.set('w', str(int(w * (K if WIDE else k))))
                for tr, h in zip(it.el.iter(A + 'tr'), it.rows): tr.set('h', str(int(h * k)))
        snap_captions(items)
        declip(items)
        split_columns(items)
        fit_text(items, report)
        if unstack(items): fit_text(items, report)
        shrink_text(items)
        fix_bottom(items)
        bs = [it.b for it in items if it.b]
        H = (max(b[1] + b[3] for b in bs) - min(b[1] for b in bs)) if bs else 0
        if H <= zone: break
    if k < K * 0.999 or not mode16: report.append('масштаб %.2f, кегль %d' % (k / K, 16 if mode16 else 14))
    for it in items:
        bld._remap_rels(it.el, src.part, dst.part)
        tree.append(it.el)
    badges = [transform_box(bb, sy0, k) for bb in badges]
    qmarks = [transform_box(bb, sy0, k) for bb in qmarks]
    # ---- 2) значки ВАЖНО и «?»
    for bb in badges:
        host = None
        for it in items:
            if it.kind == 'sp' and it.role != 'caption' and it.el.find('p:spPr/a:solidFill', NS) is not None:
                if intersects(it.b, bb, int(0.1 * IN)):
                    if host is None or area_int(it.b, bb) > area_int(host.b, bb): host = it
        if host is not None:
            f = host.el.find('p:spPr/a:solidFill/a:srgbClr', NS)
            if f is not None and f.get('val') in ('CCEAFF', 'EEE5FF'): f.set('val', 'FFFED6')
            hb = host.b
            sz = int(min(0.95 * IN, max(0.72 * IN, hb[3] + 0.2 * IN)))
            hb[2] = min(hb[2], RM - sz - int(0.1 * IN) - hb[0]); set_bbox(host.el, hb)
            fit_text([host] + [o for o in items if o is not host], report)
            hb = host.b
            x = hb[0] + hb[2] + int(0.06 * IN); y = hb[1] + hb[3] / 2 - sz / 2
        else:
            sz = int(0.9 * IN)
            cx, cy = bb[0] + bb[2] / 2, bb[1] + bb[3] / 2; x, y = cx - sz / 2, cy - sz / 2
        x = min(x, TW - sz - 0.1 * IN)
        y = min(y, 6.72 * IN * VS - sz)
        if sticker('vazhno.png'):
            pe = bld.add_image(dst, sticker('vazhno.png'), int(x), int(y), sz, sz, rot=470867)
            it = Item(pe, 'badge'); it.host = host; items.append(it)
        else:
            report.append('СТИКЕР ВРУЧНУЮ: «ВАЖНО» справа от плашки (место оставлено)')
    for qb in qmarks:
        h = int(1.55 * IN * VS); w = int(h * 376 / 518)
        cx, cy = qb[0] + qb[2] / 2, qb[1] + qb[3] / 2
        x = min(max(LM, cx - w / 2), RM - w); y = min(cy - h / 2, CB - h)
        qbox = [x, y, w, h]
        for it in items:
            if it.kind == 'sp' and text_of(it.el).strip() and intersects(it.b, qbox):
                hb = it.b
                if hb[0] < x - 1.5 * IN:
                    hb[2] = int(x - hb[0] - 0.1 * IN); set_bbox(it.el, hb)
        fit_text(items, report)
        if sticker('qmark.png'):
            pe = bld.add_image(dst, sticker('qmark.png'), int(x), int(y), w, h)
            items.append(Item(pe, 'badge'))
        else:
            report.append('СТИКЕР ВРУЧНУЮ: объёмный «?» (место оставлено)')
    # ---- 3) укрупнение картинок
    grow_pictures(items)
    # ---- 4) центрирование по вертикали
    center_vertically(items, report)
    # плашки/текстовые блоки, уходящие за низ, обрезаем по нужной высоте текста
    for it in items:
        if it.kind != 'sp': continue
        b = it.b
        if b[1] + b[3] > CB:
            tb = it.el.find('p:txBody', NS)
            need = estimate_height(tb, b[2]) if tb is not None and text_of(it.el).strip() else 0
            nh = max(need, CB - b[1])
            if nh < b[3]:
                b[3] = nh; set_bbox(it.el, b)
            if b[1] + b[3] > CB and b[1] + b[3] - (CB - b[1]) > 0 and it.role == 'caption':
                pass
    # контроль выхода за низ
    bot = max([it.b[1] + it.b[3] for it in items if it.b] + [0])
    if bot > TH - 0.3 * IN:
        report.append('низ контента %.2f″' % (bot / IN))
    return title


def unhide_white_text(items):
    """Старый шаблон 10×7,5″ и 16:9: белый текст лежал на тёмном фоне слайда («Познакомимся», контакты
    преподавателя). Фон не переносится — текст без цветной подложки (своей заливки, плашки, картинки или группы
    под ним) становится тёмным. На оранжевом или фиолетовом белый текст остаётся."""
    def back(it):
        if it.kind in ('pic', 'grp'): return True
        spPr = it.el.find('p:spPr', NS) if it.kind == 'sp' else None
        if spPr is None: return False
        if spPr.find('a:gradFill', NS) is not None or spPr.find('a:blipFill', NS) is not None: return True
        sch = spPr.find('a:solidFill/a:schemeClr', NS)
        if sch is not None: return sch.get('val') not in ('bg1', 'lt1')  # цвет темы, кроме белого
        f = spPr.find('a:solidFill/a:srgbClr', NS)
        if f is None: return False
        r, g, b = (int(f.get('val')[i:i + 2], 16) for i in (0, 2, 4))
        return 0.299 * r + 0.587 * g + 0.114 * b < 200  # не белая и не пастельная
    for it in items:
        # белые «_» и пробелы в исходнике — это пустое место, не текст
        if it.kind != 'sp' or back(it) or not re.search(r'[A-Za-zА-Яа-яЁё0-9]', text_of(it.el)): continue
        cx, cy = it.sb[0] + it.sb[2] / 2, it.sb[1] + it.sb[3] / 2
        if any(o is not it and back(o) and o.sb[0] <= cx <= o.sb[0] + o.sb[2] and o.sb[1] <= cy <= o.sb[1] + o.sb[3]
               for o in items):
            continue
        for f in it.el.iter(A + 'srgbClr'):
            if f.get('val') == 'FFFFFF' and f.getparent().getparent().tag in (A + 'rPr', A + 'endParaRPr'): f.set('val', '0F1115')


def wide_box(it, sy0, k):
    """Старый шаблон 10×7,5″ (масштаб 1:1): уменьшение k — только по высоте; текст, плашки, линии и таблицы
    сохраняют ширину и левое поле (правки методиста, 9.23 сл. 5, 27), картинки и группы уменьшаются целиком
    вокруг своего центра. Слайд шире 10″ (13,33×7,5″) сначала сжимается по ширине в K раз."""
    x, y, w, h = it.sb
    if K != 1: x, w = x * K, w * K
    ny, nh = CT + (y - sy0) * k, h * k
    if it.kind in ('pic', 'grp'):
        nw = w * k / K; x += (w - nw) / 2; w = nw
    if it.kind != 'cxn':
        if w > RM - LM: x, w = LM, RM - LM
        x = min(max(x, LM), RM - w)
    return [x, ny, w, nh]


def lift_tabs(items):
    """Старый шаблон 10×7,5″: ярлык-вкладка («вы узнаете») наезжает на верх большой плашки — ставим его
    на плашку сверху, вплотную (правка методиста, 9.23 сл. 7)."""
    plates = [it for it in items if it.kind == 'sp' and it.el.find('p:spPr/a:solidFill', NS) is not None]
    for t in plates:
        for p in plates:
            tb, pb = t.sb, p.sb
            if t is not p and tb[3] < 0.6 * IN and tb[2] < 0.7 * pb[2] and pb[3] > 2 * tb[3] \
                    and tb[1] < pb[1] < tb[1] + tb[3] and pb[0] - 0.1 * IN <= tb[0] <= pb[0] + pb[2]:
                tb[1] = pb[1] - tb[3]


def blank_bottom_frac(src, pic):
    """Доля белой/прозрачной полосы внизу картинки (с учётом обрезки), не больше 0,3:
    подпись ставится под видимый низ изображения (5.21 сл. 13 — шахматы на белом поле)."""
    from PIL import Image
    try:
        blip = pic.find('.//a:blip', NS)
        im = Image.open(io.BytesIO(src.part.related_part(blip.get(R + 'embed')).blob)).convert('RGBA')
    except Exception:
        return 0.0
    sr = pic.find('.//a:srcRect', NS)
    l = t = r_ = b_ = 0
    if sr is not None: l, t, r_, b_ = [int(sr.get(k, 0)) / 100000 for k in ('l', 't', 'r', 'b')]
    W, H = im.size
    box = (int(W * max(0, l)), int(H * max(0, t)), int(W * (1 - max(0, r_))), int(H * (1 - max(0, b_))))
    if box[2] - box[0] < 4 or box[3] - box[1] < 4: return 0.0
    im = im.crop(box)
    im.thumbnail((200, 200))
    px = im.load(); w, h = im.size
    n = 0
    for yy in range(h - 1, -1, -1):
        row = [px[xx, yy] for xx in range(w)]
        if all(a < 20 or (r >= 240 and g >= 240 and b >= 240) for r, g, b, a in row): n += 1
        else: break
    return min(0.3, n / h)


def fix_bottom(items):
    """Картинка (с подписью), вылезающая за низ зоны, уменьшается сохраняя верх и центр."""
    for it in items:
        if it.kind != 'pic': continue
        pb = it.b
        capb = it.caption.b if it.caption else None
        bottom = capb[1] + capb[3] if capb and capb[1] > pb[1] else pb[1] + pb[3]
        over = bottom - CB
        if over <= 0: continue
        nh = pb[3] - over
        if nh < pb[3] * 0.55: continue
        s_ = nh / pb[3]; nw = pb[2] * s_
        nb = [pb[0] + (pb[2] - nw) / 2, pb[1], nw, nh]
        set_bbox(it.el, nb)
        if capb and capb[1] > pb[1]:
            capb[1] -= over
            if capb[2] <= pb[2] * 1.01: capb[0], capb[2] = nb[0], nb[2]
            set_bbox(it.caption.el, capb)


def snap_captions(items):
    for it in items:
        if it.kind != 'pic' or not it.caption: continue
        pb, cb = it.b, it.caption.b
        if cb[1] >= pb[1] + pb[3] * 0.5:
            old_bottom = cb[1] + cb[3]
            cb[1] = pb[1] + pb[3] * (1 - getattr(it, 'blank_bottom', 0)) + int(0.04 * IN)
            if cb[2] < pb[2] * 1.3 or getattr(it.caption, 'cap_left', False): cb[0], cb[2] = pb[0], pb[2]
            else:
                # широкая подпись: ширина по тексту (не уже картинки), по центру картинки — не мешает соседям (5.21 сл. 13)
                tb_ = it.caption.el.find('p:txBody', NS); bp_ = tb_.find('a:bodyPr', NS)
                wt = max([text_width_pt(''.join(t.text or '' for t in p_.iter(A + 't')), 12, False)
                          for p_ in tb_.findall('a:p', NS)] + [0]) * 12700 * 1.05
                wt += int(bp_.get('lIns', 91440)) + int(bp_.get('rIns', 91440))
                nw = max(pb[2], min(cb[2], wt))
                cb[0] = min(max(pb[0] + pb[2] / 2 - nw / 2, LM), RM - nw); cb[2] = nw
            set_bbox(it.caption.el, cb)
            # подпись опустилась под картинку — текст, который под ней, сдвигаем ниже
            d = cb[1] + cb[3] + int(0.06 * IN)
            for o in items:
                if o is it or o is it.caption or o.b is None or o.kind == 'cxn': continue
                ob = o.b
                if ob[1] >= pb[1] + pb[3] - 0.05 * IN and ob[1] < d and not (ob[0] >= cb[0] + cb[2] or ob[0] + ob[2] <= cb[0]):
                    sh = d - ob[1]
                    for o2 in items:
                        if o2.b is None or o2 is it or o2 is it.caption: continue
                        b2 = o2.b
                        if b2[1] >= ob[1] - 1 and not (b2[0] >= ob[0] + ob[2] or b2[0] + b2[2] <= ob[0]) and o2 is not o:
                            pass
                    ob[1] += sh; set_bbox(o.el, ob)


def split_columns(items):
    """Соседние текстовые колонки, заходящие друг на друга, разводим с зазором."""
    ts = sorted([it for it in items if it.kind == 'sp' and it.role != 'caption' and text_of(it.el).strip()
                 and it.el.find('p:spPr/a:solidFill', NS) is None], key=lambda it: it.b[0])
    for a, b_ in zip(ts, ts[1:]):
        A_, B_ = a.b, b_.b
        vov = min(A_[1] + A_[3], B_[1] + B_[3]) - max(A_[1], B_[1])
        hov = A_[0] + A_[2] - B_[0]
        if vov > 0.5 * min(A_[3], B_[3]) and 0 < hov < 0.4 * min(A_[2], B_[2]):
            mid = B_[0] + hov / 2; g = int(0.06 * IN)
            A_[2] = mid - g - A_[0]; set_bbox(a.el, A_)
            B_[2] = B_[0] + B_[2] - (mid + g); B_[0] = mid + g; set_bbox(b_.el, B_)


def declip(items):
    for it in items:
        if it.kind == 'sp' and it.el.find('p:spPr/a:solidFill', NS) is not None and it.role != 'caption':
            tb = it.b; bp = it.el.find('p:txBody/a:bodyPr', NS)
            if bp is None: continue
            if not hasattr(it, 'ins0'): it.ins0 = (int(bp.get('lIns', 91440)), int(bp.get('rIns', 91440)))
            li, ri = it.ins0
            my_i = items.index(it)
            for pi in items:
                if pi.kind != 'pic' or items.index(pi) < my_i: continue
                pb = pi.b
                if area_int(tb, pb) <= 0: continue
                vov = min(tb[1] + tb[3], pb[1] + pb[3]) - max(tb[1], pb[1])
                if vov < 0.5 * tb[3]: continue
                if pb[0] > tb[0] + 0.55 * tb[2]: ri = max(ri, tb[0] + tb[2] - pb[0] + int(0.1 * IN))
                elif pb[0] + pb[2] < tb[0] + 0.45 * tb[2]: li = max(li, pb[0] + pb[2] - tb[0] + int(0.1 * IN))
            bp.set('lIns', str(int(li))); bp.set('rIns', str(int(ri)))
            # шапка (таблица/плашка), лежащая поверх верха плашки: текст начинаем под ней + 0,25 см
            if not hasattr(it, 'tins0'): it.tins0 = int(bp.get('tIns', 45720))
            ti = it.tins0
            for o in items[my_i + 1:]:
                if o.kind not in ('tbl', 'sp') or o.b is None or o.role == 'caption': continue
                if o.kind == 'sp' and o.el.find('p:spPr/a:solidFill', NS) is None: continue
                ob = list(o.b)
                if o.kind == 'tbl':  # у таблиц из Google Slides ext условный — берём реальные ширину и высоту
                    ob[2] = sum(int(g.get('w')) for g in o.el.iter(A + 'gridCol'))
                    ob[3] = sum(int(tr.get('h')) for tr in o.el.iter(A + 'tr'))
                hov = min(tb[0] + tb[2], ob[0] + ob[2]) - max(tb[0], ob[0])
                # шапка начинается на уровне верха плашки или чуть выше и заходит на неё (5.22 сл. 28)
                if hov < 0.5 * tb[2] or not (tb[1] - 0.35 * IN <= ob[1] <= tb[1] + 0.1 * IN) \
                        or ob[1] + ob[3] <= tb[1] or ob[1] + ob[3] > tb[1] + 0.5 * tb[3]: continue
                ti = max(ti, ob[1] + ob[3] - tb[1] + PLATE_INS)
            bp.set('tIns', str(int(ti)))
            continue
        if it.kind != 'sp' or it.role == 'caption': continue
        tb = it.b
        for pi in items:
            if pi.kind != 'pic': continue
            pb = pi.b
            if area_int(tb, pb) <= 0: continue
            vov = min(tb[1] + tb[3], pb[1] + pb[3]) - max(tb[1], pb[1])
            if pb[1] > tb[1] + 0.5 * tb[3] or vov < 0.35 * tb[3]:
                # картинка заходит снизу — укорачиваем блок (fit_text вернёт нужную высоту)
                if pb[1] > tb[1]: tb[3] = pb[1] - tb[1] - int(0.08 * IN); set_bbox(it.el, tb)
                continue
            if pb[0] > tb[0] + 0.5 * IN:
                nw = pb[0] - tb[0] - int(0.12 * IN)
                if nw >= 0.4 * tb[2]: tb[2] = nw; set_bbox(it.el, tb)
            elif pb[0] + pb[2] < tb[0] + tb[2] - 0.5 * IN and pb[0] < tb[0] + 0.3 * IN:
                nx = pb[0] + pb[2] + int(0.12 * IN); nw = tb[0] + tb[2] - nx
                if nw >= 0.4 * tb[2]: tb[2] = nw; tb[0] = nx; set_bbox(it.el, tb)


def unstack(items):
    """Текстовый блок, верх которого залез под плашку над ним, опускаем под плашку."""
    moved = False
    plates = [it for it in items if it.kind == 'sp' and it.role != 'caption' and it.el.find('p:spPr/a:solidFill', NS) is not None
              and text_of(it.el).strip()]
    for it in items:
        if it.kind != 'sp' or it in plates or not text_of(it.el).strip(): continue
        b = it.b
        for pl in plates:
            pb = pl.b
            ov = min(b[0] + b[2], pb[0] + pb[2]) - max(b[0], pb[0])
            if ov < 0.5 * min(b[2], pb[2]): continue
            if WIDE and b[1] + b[3] <= pb[1] + pb[3] + 0.05 * IN: continue  # надпись лежит на плашке целиком
            if pb[1] < b[1] < pb[1] + pb[3] - 0.03 * IN:
                d = pb[1] + pb[3] + int(0.08 * IN) - b[1]
                for o in items:
                    if o is it or o.b is None or o is pl: continue
                    ob = o.b
                    if ob[1] >= b[1] + b[3] - 0.05 * IN and not (ob[0] >= b[0] + b[2] or ob[0] + ob[2] <= b[0]):
                        ob[1] += d; set_bbox(o.el, ob)
                b[1] += d; set_bbox(it.el, b); moved = True
    return moved


def shrink_text(items):
    """Текстовый блок без заливки, прижатый к верху, с большим пустым хвостом — укорачиваем по тексту,
    чтобы освободить место картинкам (5.22 сл. 26)."""
    for it in items:
        if it.kind != 'sp' or it.role == 'caption': continue
        if it.el.find('p:spPr/a:solidFill', NS) is not None: continue
        tb = it.el.find('p:txBody', NS)
        if tb is None or not text_of(it.el).strip(): continue
        bp = tb.find('a:bodyPr', NS)
        if bp is not None and bp.get('anchor', 't') != 't': continue
        b = it.b
        need = estimate_height(tb, b[2])
        if b[3] > need + 0.25 * IN:
            b[3] = need + int(0.05 * IN); set_bbox(it.el, b)


def fit_text(items, report):
    order = sorted([it for it in items if it.b], key=lambda it: it.b[1])
    for it in order:
        if it.kind != 'sp': continue
        tb = it.el.find('p:txBody', NS)
        if tb is None or not text_of(it.el).strip(): continue
        b = it.b
        need = estimate_height(tb, b[2])
        if need > b[3] + 0.04 * IN:
            d = need - b[3]
            old_bottom = b[1] + b[3]
            b[3] = need; set_bbox(it.el, b)
            # сдвиг всего, что ниже и пересекается по горизонтали
            for o in items:
                if o is it or o.b is None: continue
                ob = o.b
                if ob[1] >= old_bottom - 0.05 * IN and not (ob[0] >= b[0] + b[2] or ob[0] + ob[2] <= b[0]):
                    ob[1] += d; set_bbox(o.el, ob)


def grow_pictures(items):
    bounds = [LM, CT, RM - LM, CB - CT]
    for it in items:
        if it.kind != 'pic': continue
        if it.el.find('p:spPr/a:xfrm', NS).get('rot'): continue
        pb = it.b
        cap = it.caption
        others = [o for o in items if o is not it and o is not cap and o.b is not None and o.kind != 'cxn']
        gap = int(0.12 * IN)

        def obox(o):  # вокруг текста держим зазор 0,12″, чтобы картинка не прилипала к тексту (5.22 сл. 16, 21)
            b_ = o.b
            if o.kind == 'sp' and text_of(o.el).strip():
                return [b_[0] - gap, b_[1] - gap, b_[2] + 2 * gap, b_[3] + 2 * gap]
            return b_
        init_ov = {id(o): ov_rect(pb, obox(o)) for o in others}
        capb = cap.b if cap else None

        def cap_box(nx, ny, nw, nh):
            # подпись едет вместе с картинкой; широкая подпись остаётся своей ширины и центрируется по картинке
            if capb[1] >= pb[1] + pb[3] * 0.5:
                off = capb[1] - pb[1] - pb[3]
                y = ny + nh + (off * nh / pb[3] if off < 0 else off)
            else: y = ny - capb[3] - (pb[1] - capb[1] - capb[3])
            if capb[2] <= max(pb[2], nw) * 1.05: return [nx, y, nw, capb[3]]
            x_ = min(max(nx + (nw - capb[2]) / 2, LM), RM - capb[2])  # широкая подпись не выходит за поля (5.21 сл. 13)
            return [x_, y, capb[2], capb[3]]
        best = (1.0, None)
        for anc in ('c', 'tl', 'tr', 'bl', 'br', 't', 'b', 'l', 'r'):
            s = 1.0
            while s < 1.8:
                s2 = s + 0.02
                nw, nh = pb[2] * s2, pb[3] * s2
                ax = {'l': 0, 'c': 0.5, 'r': 1}[anc[-1] if anc[-1] in 'lr' else 'c']
                ay = {'t': 0, 'c': 0.5, 'b': 1}[anc[0] if anc[0] in 'tb' else 'c']
                nx = pb[0] + (pb[2] - nw) * ax; ny = pb[1] + (pb[3] - nh) * ay
                nb = [nx, ny, nw, nh]
                cb = cap_box(nx, ny, nw, nh) if capb else None
                ok = within(nb, bounds) and (cb is None or within(cb, bounds))
                if ok:
                    for o in others:
                        ob = obox(o)
                        for box in [nb] + ([cb] if cb else []):
                            r0 = ov_rect(box, ob)
                            if r0 is None: continue
                            lim = init_ov.get(id(o)) if box is nb else None
                            if lim is None or not (r0[0] >= lim[0] - 1 and r0[1] >= lim[1] - 1 and r0[2] <= lim[2] + 1 and r0[3] <= lim[3] + 1):
                                ok = False; break
                        if not ok: break
                if not ok: break
                s = s2
            if s > best[0] + 1e-6: best = (s, (anc, s))
        if best[1]:
            anc, s = best[1]
            s = 1 + (s - 1) * 0.97
            nw, nh = pb[2] * s, pb[3] * s
            ax = {'l': 0, 'c': 0.5, 'r': 1}[anc[-1] if anc[-1] in 'lr' else 'c']
            ay = {'t': 0, 'c': 0.5, 'b': 1}[anc[0] if anc[0] in 'tb' else 'c']
            nx = pb[0] + (pb[2] - nw) * ax; ny = pb[1] + (pb[3] - nh) * ay
            set_bbox(it.el, [nx, ny, nw, nh])
            # линии-разделители поверх картинки (например, граница на карте) едут вместе с ней
            for c in items:
                if c.kind != 'cxn' or c.b is None: continue
                cb_ = c.b
                cy = cb_[1] + cb_[3] / 2
                if pb[1] < cy < pb[1] + pb[3] and min(cb_[0] + cb_[2], pb[0] + pb[2]) - max(cb_[0], pb[0]) > 0.5 * pb[2]:
                    k_ = nh / pb[3]
                    cb_[1] = ny + (cb_[1] - pb[1]) * k_; cb_[3] = cb_[3] * k_
                    set_bbox(c.el, cb_)
            if cap:
                set_bbox(cap.el, cap_box(nx, ny, nw, nh))
            for bi in items:
                if bi.kind == 'badge' and bi.host is it: pass


def ov_rect(a, b):
    x0 = max(a[0], b[0]); y0 = max(a[1], b[1]); x1 = min(a[0] + a[2], b[0] + b[2]); y1 = min(a[1] + a[3], b[1] + b[3])
    if x1 <= x0 or y1 <= y0: return None
    return (x0, y0, x1, y1)


def within(b, bd):
    return b[0] >= bd[0] - 1 and b[1] >= bd[1] - 1 and b[0] + b[2] <= bd[0] + bd[2] + 1 and b[1] + b[3] <= bd[1] + bd[3] + 1


def center_vertically(items, report):
    bs = [it.b for it in items if it.b]
    if not bs: return
    top = min(b[1] for b in bs); bot = max(b[1] + b[3] for b in bs)
    H = bot - top
    if H > CB - CT:
        report.append('контент выше зоны: %.2f″' % (H / IN))
        shift = CT - top
    else:
        shift = CT + (CB - CT - H) / 2 - top
    if abs(shift) < 0.12 * IN: return
    for it in items:
        b = it.b
        if b: b[1] += shift; set_bbox(it.el, b)


# ---------------------------------------------------------------- спец-слайды
def src_items_by_y(src, pred):
    res = []
    for el in src.shapes._spTree:
        if etree.QName(el).localname != 'sp': continue
        ne = materialize_sp(src, el)
        b = bbox(ne)
        if b is None or is_offslide(b) or is_footer(b): continue
        if pred(ne, el): res.append((b[1], b[0], ne, el))
    res.sort(key=lambda x: (x[0], x[1]))
    return res


def src_paragraphs(src, sp_src, ne, smap, sz=1600, keep_bold=True, keep_color=True):
    """Абзацы источника → [(text, rPr)] с новыми rPr."""
    ne = copy.deepcopy(ne)
    resolve_text(src, sp_src, ne, smap, role='body', summary=True)
    out = []
    for p in ne.find('p:txBody', NS).findall('a:p', NS):
        para = []
        for r in p:
            ln = etree.QName(r).localname
            if ln == 'r':
                rp = r.find('a:rPr', NS)
                col = rp.find('a:solidFill/a:srgbClr', NS).get('val')
                nr = new_rpr(sz, b=keep_bold and rp.get('b') == '1', color=col if keep_color else '0F1115')
                hl = rp.find('a:hlinkClick', NS)
                if hl is not None: nr.append(copy.deepcopy(hl))
                para.append((r.findtext('a:t', '', NS), nr))
            elif ln == 'br':
                para.append(('\n', None))
        out.append(para)
    while out and not ''.join(t for t, _ in out[-1]).strip(): out.pop()
    while out and not ''.join(t for t, _ in out[0]).strip(): out.pop(0)
    return out


def get_ph(s, typ):
    for sp in s.shapes._spTree.iter(P + 'sp'):
        ph = sp.find('p:nvSpPr/p:nvPr/p:ph', NS)
        if ph is not None and ph.get('type') == typ: return sp
    return None


def build_title(bld, src, lesson_no, title):
    s = bld.clone(DONOR['title'])
    tree = s.shapes._spTree
    # реплики-выноски донора относятся к его уроку — убираем (фигуры с поворотом и кривые соединители)
    if DONOR.get('title_strip_callouts'):
        for el in list(tree)[2:]:
            ln = etree.QName(el).localname
            x = xfrm_of(el)
            if ln == 'cxnSp' or (ln == 'sp' and x is not None and x.get('rot') and el.find('.//p:nvPr/p:ph', NS) is None):
                tree.remove(el)
    # заголовок: подбираем кегль и переносы, чтобы влезло в 2 строки над картинкой
    one = re.sub(r'\s+', ' ', title.replace('\n', ' ')).strip()
    tsp = get_ph(s, 'ctrTitle')
    width_pt = bbox(tsp)[2] / 12700 if tsp is not None and bbox(tsp) else 600
    # кегль не больше, чем у заголовка донора (5.8 — 40 pt из макета, 6.4 — 48 pt)
    max_fs = 48
    if tsp is not None:
        szs = [int(r.get('sz')) for r in tsp.iter(A + 'rPr') if r.get('sz')]
        if szs: max_fs = max(szs) / 100
        else:
            lsts, _, _ = style_chain(s, tsp)
            for l in lsts:
                d = l.find('a:lvl1pPr/a:defRPr', NS) if l is not None else None
                if d is not None and d.get('sz'): max_fs = int(d.get('sz')) / 100
    # ширина строки: картинка донора (лис с брелоком на 5.8) не должна перекрываться текстом
    def line_w(i, fs):
        if DONOR.get('title_pic') or tsp is None: return width_pt
        y0 = bbox(tsp)[1] + i * fs * 1.25 * 12700; y1 = y0 + fs * 1.25 * 12700
        left = title_obstacle_left(s, y0, y1)
        return width_pt if left is None else min(width_pt, (left - bbox(tsp)[0] - 0.15 * IN) / 12700)
    def fits(lines, fs):
        return all(text_width_pt(l, fs, True) * 1.08 <= line_w(i, fs) for i, l in enumerate(lines))
    chosen = None
    for fs in [f for f in (48, 44, 40, 36, 32) if f <= max_fs] or [32]:
        lines = wrap_words(one, fs, width_pt / 1.08)
        if len(lines) <= 2 and not fits(lines, fs):
            lines = wrap_words(one, fs, line_w(1, fs) / 1.08)
        if len(lines) <= 2 and fits(lines, fs): chosen = (fs, lines); break
    if chosen is None: chosen = (32, wrap_words(one, 32, line_w(1, 32) / 1.08))
    fs, lines = chosen
    # если исходный перенос ложится в те же 2 строки — сохраняем авторский
    src_lines = [l.strip() for l in title.split('\n') if l.strip()]
    if len(src_lines) == 2 and fits(src_lines, fs): lines = src_lines
    bld.set_title(s, '\n'.join(lines))
    if tsp is not None:
        for rp in tsp.iter(A + 'rPr', A + 'endParaRPr'): rp.set('sz', str(fs * 100))
    for sp in tree.iter(P + 'sp'):
        t = text_of(sp).strip()
        if t in ('Русский язык', 'История'): set_plain(sp, 'История')
        elif t.startswith('Урок'): set_plain(sp, 'Урок %s' % lesson_no)
        elif t.endswith('класс'): set_plain(sp, '%s класс' % GRADE)
    # иллюстрация титула: картинка из исходного титула → в рамку донора
    if DONOR.get('title_pic'):
        spic = None
        for pic in src._element.iter(P + 'pic'):
            if img_hash(src, pic) in DECOR_SKIP: continue
            b = bbox(pic)
            if b and not is_offslide(b) and (spic is None or b[2] * b[3] > spic[1]): spic = (pic, b[2] * b[3])
        dpic = s._element.find('.//p:pic', NS)
        if spic and dpic is not None:
            from PIL import Image
            blob = src.part.related_part(spic[0].find('.//a:blip', NS).get(R + 'embed')).blob
            im = Image.open(io.BytesIO(blob))
            alpha = im.mode in ('RGBA', 'LA', 'P') and im.convert('RGBA').getchannel('A').getextrema()[0] < 250
            if alpha:  # вырезанный объект — без обрезки, по высоте рамки
                b = bbox(dpic); iw, ih = im.size
                nh = b[3] * 1.1; nw = nh * iw / ih
                cx = b[0] + b[2] / 2
                replace_pic_cover(bld, s, dpic, src, spic[0])
                set_bbox(dpic, [cx - nw / 2, b[1] + b[3] - nh, nw, nh])
                x = xfrm_of(dpic); x.attrib.pop('rot', None)
                for sr in dpic.iter(A + 'srcRect'): sr.getparent().remove(sr)
            else:
                replace_pic_cover(bld, s, dpic, src, spic[0])
        # вторая строка заголовка заходит на картинку — опускаем картинку
        if dpic is not None and len(lines) >= 2 and tsp is not None:
            pb = bbox(dpic)
            if LM + text_width_pt(lines[-1], fs, True) * 12700 * 1.05 > pb[0]:
                need_top = bbox(tsp)[1] + int(len(lines) * fs * 1.35 * 12700) + int(0.12 * IN)
                if pb[1] < need_top:
                    pb[1] = min(need_top, int(6.5 * IN * VS) - pb[3]); set_bbox(dpic, pb)
    return s


def title_obstacle_left(s, y0, y1):
    """Самый левый непрозрачный пиксель картинок титула в полосе y0..y1 (EMU), с учётом групп."""
    from PIL import Image
    best = None
    for pic in s._element.iter(P + 'pic'):
        b = bbox(pic)
        if b is None: continue
        x, y, w, h = map(float, b)
        g = pic.getparent()
        while g is not None and etree.QName(g).localname == 'grpSp':
            gx = g.find('p:grpSpPr/a:xfrm', NS)
            o, e, co, ce = [gx.find('a:' + t, NS) for t in ('off', 'ext', 'chOff', 'chExt')]
            sx = int(e.get('cx')) / max(1, int(ce.get('cx'))); sy = int(e.get('cy')) / max(1, int(ce.get('cy')))
            x = int(o.get('x')) + (x - int(co.get('x'))) * sx; y = int(o.get('y')) + (y - int(co.get('y'))) * sy
            w *= sx; h *= sy
            g = g.getparent()
        if y + h <= y0 or y >= y1: continue
        blob = s.part.related_part(pic.find('.//a:blip', NS).get(R + 'embed')).blob
        im = Image.open(io.BytesIO(blob)).convert('RGBA')
        a = im.getchannel('A')
        r0 = max(0, int((y0 - y) / h * im.height)); r1 = min(im.height, int((y1 - y) / h * im.height) + 1)
        if r1 <= r0: continue
        bb = a.crop((0, r0, im.width, r1)).point(lambda v: 255 if v > 40 else 0).getbbox()
        if bb is None: continue
        lx = x + bb[0] / im.width * w
        best = lx if best is None else min(best, lx)
    return best


def wrap_words(text, fs, width_pt):
    lines, cur = [], ''
    for w in text.split(' '):
        t = (cur + ' ' + w).strip()
        if cur and text_width_pt(t, fs, True) > width_pt:
            lines.append(cur); cur = w
        else: cur = t
    if cur: lines.append(cur)
    # висячие предлоги/союзы переносим на следующую строку
    for i in range(len(lines) - 1):
        ws = lines[i].split(' ')
        if len(ws) > 1 and len(ws[-1]) <= 2:
            lines[i] = ' '.join(ws[:-1]); lines[i + 1] = ws[-1] + ' ' + lines[i + 1]
    return lines


def set_plain(sp, text):
    tb = sp.find('p:txBody', NS)
    ps = tb.findall('a:p', NS)
    rs = ps[0].findall('a:r', NS)
    rs[0].find('a:t', NS).text = text
    for r in rs[1:]: ps[0].remove(r)
    for p in ps[1:]:
        # класс набран в две строки: «5 / класс» — сохраняем структуру донора
        pass


def build_plan(bld, src):
    s = bld.clone(DONOR['plan'])
    smap = scheme_map(src)
    rows = plan_rows(src)
    tbl = s._element.find('.//a:tbl', NS)
    trs = tbl.findall('a:tr', NS)
    tmpl = trs[0]
    for tr in trs: tbl.remove(tr)
    for i, t in enumerate(rows):
        tr = copy.deepcopy(tmpl)
        tcs = tr.findall('a:tc', NS)
        tcs[0].find('.//a:t', NS).text = '%02d' % (i + 1)
        r = tcs[1].find('.//a:r', NS)
        r.find('a:t', NS).text = t
        tbl.append(tr)
    gf = s._element.find('.//p:graphicFrame', NS)
    off = gf.find('p:xfrm/a:off', NS)
    lines = sum(max(1, math.ceil(text_width_pt(t, 16, False) / 225)) for t in rows)
    h = len(rows) * 0.22 * IN + lines * 0.30 * IN
    off.set('y', str(int(max(1.3 * IN * VS, 6.25 * IN * VS - h))))
    # картинка источника → в рамку донора (cover-crop)
    spic = None
    for pic in src._element.iter(P + 'pic'):
        h = img_hash(src, pic)
        if h in DECOR_SKIP: continue
        b = bbox(pic)
        if b and not is_offslide(b) and not is_footer(b) and (spic is None or b[2] * b[3] > spic[1]): spic = (pic, b[2] * b[3])
    dpic = s._element.find('.//p:pic', NS)
    if spic:
        replace_pic_fit(bld, s, dpic, src, spic[0])
    elif WIDE and dpic is not None:
        dpic.getparent().remove(dpic)  # картинка донора — из чужого урока
    if WIDE:
        title, _ = slide_title_info(src)
        if title: bld.set_title(s, title)
    return s


def replace_pic_fit(bld, s, dpic, src, spic, anchor='r'):
    """cover-crop, но если обрезать пришлось бы > 20 % — вписываем картинку целиком (рамка меняет пропорции)."""
    from PIL import Image
    b = bbox(dpic)
    part = src.part.related_part(spic.find('.//a:blip', NS).get(R + 'embed'))
    iw, ih = Image.open(io.BytesIO(part.blob)).size
    sr = spic.find('.//a:srcRect', NS)
    l = t = r_ = b_ = 0
    if sr is not None: l, t, r_, b_ = [int(sr.get(k, 0)) for k in ('l', 't', 'r', 'b')]
    ia = iw * (1 - (l + r_) / 100000) / (ih * (1 - (t + b_) / 100000)); fa = b[2] / b[3]
    if 0.8 <= ia / fa <= 1.25:
        replace_pic_cover(bld, s, dpic, src, spic); return
    replace_pic_cover(bld, s, dpic, src, spic)
    bf = dpic.find('p:blipFill', NS)
    nsr = bf.find('a:srcRect', NS)
    for k, v in zip('ltrb', (l, t, r_, b_)): nsr.set(k, str(int(v)))
    if ia > fa: nw, nh = b[2], b[2] / ia
    else: nw, nh = b[3] * ia, b[3]
    nx = b[0] + b[2] - nw if anchor == 'r' else b[0] + (b[2] - nw) / 2
    set_bbox(dpic, [nx, b[1] + (b[3] - nh) / 2, nw, nh])


def replace_pic_cover(bld, s, dpic, src, spic):
    b = bbox(dpic)
    blip = spic.find('.//a:blip', NS)
    part = src.part.related_part(blip.get(R + 'embed'))
    from PIL import Image
    im = Image.open(io.BytesIO(part.blob))
    iw, ih = im.size
    # учитываем исходную обрезку
    sr = spic.find('.//a:srcRect', NS)
    l = t = r_ = b_ = 0
    if sr is not None:
        l, t, r_, b_ = [int(sr.get(k, 0)) for k in ('l', 't', 'r', 'b')]
    vw = iw * (1 - (l + r_) / 100000); vh = ih * (1 - (t + b_) / 100000)
    fa = b[2] / b[3]; ia = vw / vh
    if ia > fa:  # шире — режем бока
        cut = (1 - fa / ia) * (100000 - l - r_) / 2
        l += cut; r_ += cut
    else:
        cut = (1 - ia / fa) * (100000 - t - b_) / 2
        t += cut; b_ += cut
    _, rid = s.part.get_or_add_image_part(io.BytesIO(part.blob))
    bf = dpic.find('p:blipFill', NS)
    bl = bf.find('a:blip', NS); bl.set(R + 'embed', rid)
    for c in list(bl): bl.remove(c)
    old = bf.find('a:srcRect', NS)
    if old is not None: bf.remove(old)
    nsr = etree.Element(A + 'srcRect')
    for k, v in zip('ltrb', (l, t, r_, b_)): nsr.set(k, str(int(v)))
    bf.insert(1, nsr)


def _src_block(src, el, smap, role='body'):
    ne = materialize_sp(src, el)
    restyle_spPr(ne.find('p:spPr', NS), smap)
    resolve_text(src, el, ne, smap, role=role)
    renumber(ne)
    return ne


def _strip_fill(sp):
    spPr = sp.find('p:spPr', NS)
    for c in list(spPr):
        if etree.QName(c).localname in ('solidFill', 'gradFill'): spPr.replace(c, etree.Element(A + 'noFill'))
    ln = spPr.find('a:ln', NS)
    if ln is not None: spPr.remove(ln)


def _pic_copy(bld, src, dst, pic, box):
    ne = copy.deepcopy(pic)
    restyle_spPr(ne.find('p:spPr', NS), scheme_map(src))
    renumber(ne)
    set_bbox(ne, box)
    bld._remap_rels(ne, src.part, dst.part)
    dst.shapes._spTree.append(ne)
    return ne


def _fit(b, box, ax=0.5, ay=0.5):
    k = min(box[2] / b[2], box[3] / b[3])
    w, h = b[2] * k, b[3] * k
    return [box[0] + (box[2] - w) * ax, box[1] + (box[3] - h) * ay, w, h], k


def build_why(bld, src):
    """«Почему стоит посмотреть этот урок» → макет 6.4: текст узкой колонкой слева внизу, картинка крупно справа."""
    s = bld.clone(DONOR['why'])
    smap = scheme_map(src)
    tree = s.shapes._spTree
    texts = src_items_by_y(src, lambda ne, el: len(text_of(ne).strip()) > 3 and 'навигации' not in text_of(ne))
    head = texts[0]
    bld.set_title(s, para_text(head[2]).replace('\n', ' ').strip())
    # основной текст — самый длинный блок, остальное (подписи-мемы) едет вместе с картинками
    body_src = max(texts[1:], key=lambda x: len(text_of(x[2])))
    body = get_ph(s, 'body')
    paras = src_paragraphs(src, body_src[3], body_src[2], smap, sz=1600)
    pad_plate(body)
    Builder.set_text_runs(body.find('p:txBody', NS), paras)
    dpic = s._element.find('.//p:pic', NS)
    tree.remove(dpic)
    # текст — плашкой под заголовком на всю ширину (как в 6.2), картинка — на всё оставшееся поле
    tb = body.find('p:txBody', NS)
    need = estimate_height(tb, RM - LM)
    set_bbox(body, [LM, int(0.95 * IN * VS), RM - LM, need])
    top = int(0.95 * IN * VS) + need + int(0.2 * IN)
    box = [LM, top, RM - LM, int(6.75 * IN * VS) - top]
    parts = []
    for el in list(src.shapes._spTree)[2:]:
        ln = etree.QName(el).localname
        if ln == 'pic':
            if img_hash(src, el) in DECOR_SKIP or is_offslide(bbox(el)): continue
            parts.append(('pic', el, bbox(el)))
        elif ln == 'sp' and el is not head[3] and el is not body_src[3] and text_of(el).strip() and 'навигации' not in text_of(el):
            ne = materialize_sp(src, el)
            if not is_offslide(bbox(ne)): parts.append(('sp', el, bbox(ne)))
    if parts:
        x0 = min(p[2][0] for p in parts); y0 = min(p[2][1] for p in parts)
        x1 = max(p[2][0] + p[2][2] for p in parts); y1 = max(p[2][1] + p[2][3] for p in parts)
        (bx, by, bw, bh), k = _fit([x0, y0, x1 - x0, y1 - y0], box, ax=0.5, ay=0.5)
        for kind, el, b in parts:
            nb = [bx + (b[0] - x0) * k, by + (b[1] - y0) * k, b[2] * k, b[3] * k]
            if kind == 'pic': _pic_copy(bld, src, s, el, nb)
            else:
                ne = _src_block(src, el, smap, role='card')
                szs = [int(r.get('sz')) for r in el.iter(A + 'rPr') if r.get('sz')]
                if szs:
                    for rp in ne.iter(A + 'rPr', A + 'endParaRPr'): rp.set('sz', str(max(1400, min(2400, int(max(szs) * k * K / 100) * 100))))
                set_bbox(ne, nb); tree.append(ne)
    return s


def build_think(bld, src, report):
    """«Подумаем!» / «Работа с картой» → паттерн 6.4: вопрос справа от заголовка, иллюстрация на всё поле."""
    s = bld.clone(DONOR['think'], keep_content=False)
    smap = scheme_map(src)
    tree = s.shapes._spTree
    title, tels = slide_title_info(src)
    tsp = bld.set_title(s, title)
    tw = int((text_width_pt(title, 24, True) * 1.1 + 10) * 12700)
    tx = tsp.find('p:spPr/a:xfrm/a:ext', NS); tx.set('cx', str(tw))
    pics, qm = [], False
    for el in src._element.iter(P + 'pic'):
        h = img_hash(src, el)
        if h in DECOR_Q: qm = True; continue
        if h in DECOR_SKIP or h in DECOR_VAZHNO or is_offslide(bbox(el)): continue
        pics.append(el)
    pics.sort(key=lambda e: bbox(e)[0])
    qs = src_items_by_y(src, lambda ne, el: el not in tels and len(text_of(ne).strip()) > 2 and 'навигации' not in text_of(ne))
    # вопрос
    qx = max(int(2.9 * IN), LM + tw + int(0.3 * IN))
    q_bottom = int(1.1 * IN * VS)
    for n, (_, _, ne, el) in enumerate(qs):
        q = _src_block(src, el, smap)
        _strip_fill(q)
        tb = q.find('p:txBody', NS)
        bp = tb.find('a:bodyPr', NS); bp.set('lIns', '0'); bp.set('rIns', '0'); bp.set('anchor', 't')
        w = RM - qx
        need = estimate_height(tb, w)
        if need > 1.2 * IN * VS:
            for rp in q.iter(A + 'rPr', A + 'endParaRPr'): rp.set('sz', '1400')
            need = estimate_height(tb, w)
        y = int(0.37 * IN * VS) if n == 0 else q_bottom
        if need > 1.25 * IN * VS or n > 0 and y + need > 1.6 * IN * VS:  # не влезает рядом с заголовком — под заголовок
            qx2 = LM; w = RM - LM; need = estimate_height(tb, w)
            y = max(q_bottom, int(1.0 * IN * VS)); set_bbox(q, [qx2, y, w, need])
        else:
            set_bbox(q, [qx, y, w, need])
        q_bottom = max(q_bottom, y + need)
        tree.append(q)
    top = max(int(1.5 * IN * VS), q_bottom + int(0.12 * IN)); bot = int(6.75 * IN * VS)
    x0 = LM + (int(1.45 * IN) if qm else 0)
    region = [x0, top, RM - x0, bot - top]
    gap = int(0.15 * IN)
    if len(pics) == 1:
        nb, _ = _fit(bbox(pics[0]), region, ax=1.0)
        boxes = [nb]
    else:
        wsum = sum(bbox(p)[2] / bbox(p)[3] for p in pics)
        h = min(region[3], (region[2] - gap * (len(pics) - 1)) / wsum)
        tot = h * wsum + gap * (len(pics) - 1)
        x = region[0] + (region[2] - tot); boxes = []
        for p in pics:
            w = h * bbox(p)[2] / bbox(p)[3]
            boxes.append([x, region[1] + (region[3] - h) / 2, w, h]); x += w + gap
    for p, b in zip(pics, boxes): _pic_copy(bld, src, s, p, b)
    if qm:
        h = int(1.55 * IN * VS); w = int(h * 376 / 518)
        pb = min(b[1] + b[3] for b in boxes)
        if sticker('qmark.png'):
            bld.add_image(s, sticker('qmark.png'), LM + int(0.1 * IN), int(min(bot - h, pb - h)), w, h)
        else:
            report.append('СТИКЕР ВРУЧНУЮ: объёмный «?» слева внизу (место оставлено)')
    return s


def build_divider(bld, src, k):
    s = bld.clone(DONOR['dividers'][k % len(DONOR['dividers'])])
    # описание-заглушка фирменного шаблона: текст из исходника (10×7,5″) или убираем
    desc = (wide_divider(src) or ('', ''))[1] if WIDE else ''
    for sp in list(s.shapes._spTree.iter(P + 'sp')):
        if not text_of(sp).startswith('Добавь описание'): continue
        if desc: Builder.set_text_runs(sp.find('p:txBody', NS), [[(desc, None)]]); desc = ''
        else: sp.getparent().remove(sp)
    if WIDE:
        bld.set_title(s, (wide_divider(src) or ('', ''))[0])
        return s
    t = ''
    for el in src.shapes._spTree.iter(P + 'sp'):
        ph = el.find('p:nvSpPr/p:nvPr/p:ph', NS)
        if ph is not None and ph.get('type') in ('ctrTitle', 'title'):
            t = para_text(el).strip()
    if not t:
        t = max((para_text(el).strip() for el in src.shapes._spTree.iter(P + 'sp') if 'навигации' not in text_of(el)), key=len)
    bld.set_title(s, t)
    return s


def wait_items(src):
    smap = scheme_map(src)
    title, tels = slide_title_info(src)
    its = src_items_by_y(src, lambda ne, el: el not in tels and len(text_of(ne).strip()) > 3 and 'навигации' not in text_of(ne))
    return [(el, ne) for _, _, ne, el in its], smap


def build_wait(bld, src):
    s = bld.clone(DONOR['wait'])
    its, smap = wait_items(src)
    gfs = s._element.findall('.//p:graphicFrame', NS)
    tmpl = copy.deepcopy(gfs[0])
    tree = s.shapes._spTree
    for g in gfs: tree.remove(g)
    n = len(its)
    step = DONOR.get('wait_step', 1.53) * IN * VS
    y0 = DONOR.get('wait_center', 3.74) * IN * VS - step * (n - 1) / 2
    for i, (el, ne) in enumerate(its):
        g = copy.deepcopy(tmpl)
        g.find('p:nvGraphicFramePr/p:cNvPr', NS).set('id', str(900 + i))
        tcs = g.findall('.//a:tc', NS)
        tcs[0].find('.//a:t', NS).text = '%02d' % (i + 1)
        paras = src_paragraphs(src, el, ne, smap, sz=1600, keep_color=False)
        tb = tcs[1].find('a:txBody', NS)
        Builder.set_text_runs(tb, paras)
        g.find('p:xfrm/a:off', NS).set('y', str(int(y0 + i * step)))
        tree.append(g)
    return s


def build_summary(bld, src):
    s = bld.clone(DONOR['summary'])
    smap = scheme_map(src)
    title, tels = slide_title_info(src)
    its = src_items_by_y(src, lambda ne, el: el not in tels and len(text_of(ne).strip()) > 3 and 'навигации' not in text_of(ne))
    # номера-стрелки «1», «2», «3» исходника: строк столько, сколько номеров (незаполненные строки остаются пустыми)
    marks = src_items_by_y(src, lambda ne, el: re.fullmatch(r'\d{1,2}', text_of(ne).strip()) is not None
                           and bbox(ne)[3] > 0.5 * IN and ne.find('p:spPr/a:solidFill', NS) is not None)
    rows = [[x] for x in its]
    if len(marks) >= 2 and len(marks) >= len(its):
        rows = []
        for my, _, mne, _ in marks:
            mb = bbox(mne)
            rows.append([x for x in its if mb[1] - 0.05 * IN <= x[0] + bbox(x[2])[3] / 2 <= mb[1] + mb[3] + 0.05 * IN])
        if sum(len(r) for r in rows) != len(its): rows = [[x] for x in its]
    tbl = s._element.find('.//a:tbl', NS)
    trs = tbl.findall('a:tr', NS)
    tmpl = trs[0]
    for tr in trs: tbl.remove(tr)
    for i, row in enumerate(rows):
        tr = copy.deepcopy(tmpl)
        tcs = tr.findall('a:tc', NS)
        tcs[0].find('.//a:t', NS).text = '%02d' % (i + 1)
        paras = []
        for _, _, ne, el in row: paras += src_paragraphs(src, el, ne, smap, sz=1600)
        if not paras:
            paras = [[('', new_rpr(1600))]]
        Builder.set_text_runs(tcs[1].find('a:txBody', NS), paras)
        tbl.append(tr)
    gf = s._element.find('.//p:graphicFrame', NS)
    n = len(rows)
    y = max(1.9, 2.77 - (n - 3) * 0.45) * IN
    if WIDE:  # старый шаблон 10×7,5″ (правка методиста, 9.23 сл. 31): таблица под картинкой донора, строки сжаты по месту
        y = max(y, 2.1 * IN * VS)
        trs = tbl.findall('a:tr', NS)
        w = [int(g.get('w')) for g in tbl.iter(A + 'gridCol')][1]
        need = lambda tr: estimate_height(tr.findall('a:tc', NS)[1].find('a:txBody', NS), w) + int(0.15 * IN)
        if sum(map(need, trs)) > 6.75 * IN * VS - y:  # не влезает — кегль 14
            for tr in trs:
                for rp in tr.findall('a:tc', NS)[1].iter(A + 'rPr', A + 'endParaRPr'):
                    if rp.get('sz') == '1600': rp.set('sz', '1400')
        hs = [need(tr) for tr in trs]
        y = min(y, max(1.3 * IN * VS, 6.75 * IN * VS - sum(hs)))
        extra = max(0, 6.75 * IN * VS - y - sum(hs)) / len(trs)
        for tr, h in zip(trs, hs): tr.set('h', str(int(max(h, min(int(tr.get('h')), h + extra)))))
        ext = gf.find('p:xfrm/a:ext', NS)
        ext.set('cx', str(sum(int(g.get('w')) for g in tbl.iter(A + 'gridCol')))); ext.set('cy', str(sum(int(tr.get('h')) for tr in trs)))
    gf.find('p:xfrm/a:off', NS).set('y', str(int(y)))
    if title: bld.set_title(s, title)
    return s


def build_reflection(bld, src):
    s = bld.clone(DONOR['reflection'])
    smap = scheme_map(src)
    title, tels = slide_title_info(src)
    # источник: заголовок-вопрос и утверждения
    stm = src_items_by_y(src, lambda ne, el: el not in tels and len(text_of(ne).strip()) > 12 and 'навигации' not in text_of(ne)
                         and 'паузу' not in text_of(ne))
    header = [x for x in stm if 'Выберите' in text_of(x[2]) or 'справедливы' in text_of(x[2])]
    stmts = [x for x in stm if x not in header]
    tree = s.shapes._spTree
    body = get_ph(s, 'body')
    if header:
        paras = src_paragraphs(src, header[0][3], header[0][2], smap, sz=1600)
        # сохранить rPr донора
        tb = body.find('p:txBody', NS)
        d_r = tb.find('.//a:r/a:rPr', NS)
        paras = [[(t, copy.deepcopy(d_r) if t != '\n' else None) for t, _ in para] for para in paras]
        Builder.set_text_runs(tb, paras)
    drows = []
    circles = []
    for sp in tree.iter(P + 'sp'):
        if sp.find('p:nvSpPr/p:nvPr/p:ph', NS) is not None: continue
        t = text_of(sp).strip()
        if t in ('1', '2', '3', '4', '5'): circles.append(sp)
        elif t: drows.append(sp)
    drows.sort(key=lambda e: bbox(e)[1])
    n = len(stmts)
    if n > len(drows):
        bld.log.append('рефлексия: утверждений %d > 4' % n)
    for i, row in enumerate(drows):
        rb = bbox(row)
        if i < n:
            el, ne = stmts[i][3], stmts[i][2]
            txt = para_text(ne).replace('\n', ' ').strip()
            txt = re.sub(r'\s+', ' ', txt)
            tb = row.find('p:txBody', NS)
            d_r = tb.find('.//a:r/a:rPr', NS)
            Builder.set_text_runs(tb, [[(txt, copy.deepcopy(d_r))]])
        else:
            for c in circles:
                cb = bbox(c)
                if abs((cb[1] + cb[3] / 2) - (rb[1] + rb[3] / 2)) < 0.35 * IN: tree.remove(c)
            tree.remove(row)
    if title: bld.set_title(s, title)
    return s


def build_links(bld, srcs):
    s = bld.clone(DONOR['links'])
    paras = []
    for src in srcs:
        smap = scheme_map(src)
        title, tels = slide_title_info(src)
        its = src_items_by_y(src, lambda ne, el: el not in tels and len(text_of(ne).strip()) > 3 and 'навигации' not in text_of(ne)
                             and not (WIDE and text_of(ne).strip().lower().startswith(('использованные', 'ссылки на использованные'))))
        its.sort(key=lambda x: (x[1] // int(1.0 * IN), x[0]))  # по колонкам
        for _, _, ne, el in its:
            for para in src_paragraphs(src, el, ne, smap, sz=1200, keep_bold=False):
                if ''.join(t for t, _ in para).strip(): paras.append((para, src))
    # ссылки: без нумерации, гиперссылки сохраняем
    srcmap = [src for _, src in paras]
    paras = [p for p, _ in paras]
    for para in paras:
        for t, rp in para:
            if rp is not None and rp.find('a:hlinkClick', NS) is not None:
                f = rp.find('a:solidFill', NS); rp.remove(f); rp.set('u', 'sng')
    tsp = get_ph(s, 'title')
    head = 'Ссылки на иллюстрации'
    if WIDE:
        head = next((re.sub(r'\s+', ' ', text_of(el)).strip() for src in srcs for el in src.shapes._spTree.iter(P + 'sp')
                     if text_of(el).strip().lower().startswith('использованные')), head)
    bld.set_title(s, head)
    tb_x = tsp.find('p:spPr/a:xfrm', NS)
    tb_x.find('a:ext', NS).set('cx', str(RM - LM)); tb_x.find('a:ext', NS).set('cy', str(int(0.45 * IN)))
    body = get_ph(s, 'body')
    tree = s.shapes._spTree
    ncol = 2 if len(paras) > 9 else 1
    cols = [paras] if ncol == 1 else [paras[:(len(paras) + 1) // 2], paras[(len(paras) + 1) // 2:]]
    x0 = int(3.0 * IN); wcol = (RM - x0 - (ncol - 1) * int(0.25 * IN)) // ncol
    bodies = [body] + [copy.deepcopy(body) for _ in range(ncol - 1)]
    for ci, (bd, cp) in enumerate(zip(bodies, cols)):
        if ci: tree.append(bd); bd.find('p:nvSpPr/p:cNvPr', NS).set('id', str(950 + ci))
        tb = bd.find('p:txBody', NS)
        p0 = tb.find('a:p', NS)
        pPr = copy.deepcopy(p0.find('a:pPr', NS))
        for c in list(pPr):
            if etree.QName(c).localname.startswith('bu'): pPr.remove(c)
        etree.SubElement(pPr, A + 'buNone')
        pPr.set('marL', '0'); pPr.set('indent', '0')
        for p in tb.findall('a:p', NS): tb.remove(p)
        base = sum(len(c) for c in cols[:ci])
        for pi_, para in enumerate(cp):
            p = etree.SubElement(tb, A + 'p'); p.append(copy.deepcopy(pPr))
            for t, rp in para:
                if t == '\n': etree.SubElement(p, A + 'br'); continue
                r = etree.SubElement(p, A + 'r'); r.append(rp); etree.SubElement(r, A + 't').text = t
            bld._remap_rels(p, srcmap[base + pi_].part, s.part)
        x = bd.find('p:spPr/a:xfrm', NS)
        x.find('a:off', NS).set('x', str(x0 + ci * (wcol + int(0.25 * IN)))); x.find('a:off', NS).set('y', str(int(1.25 * IN * VS)))
        x.find('a:ext', NS).set('cx', str(wcol)); x.find('a:ext', NS).set('cy', str(int(5.3 * IN * VS)))
    return s


# ---------------------------------------------------------------- сборка
def convert(src_path, out_path, lesson_no, overrides=None):
    overrides = overrides or {}
    global SRC_W, SRC_H, K, WIDE, TITLE_N
    sp = Presentation(src_path)
    slides = list(sp.slides)
    if sp.slide_width > 8.5 * IN:  # старый шаблон 10×7,5″: геометрия из файла, титул ищем
        SRC_W, SRC_H, WIDE = sp.slide_width, sp.slide_height, True
        K = TW / SRC_W
        TITLE_N = find_title_slide(slides)
    bld = Builder()
    report = []
    div_k = 0
    section = 0
    kinds = []
    allk = [overrides.get(i, {}).get('kind') or classify(src, i) for i, src in enumerate(slides, 1)]
    for i, src in enumerate(slides, 1):
        kind = allk[i - 1]
        if kind == 'links' and i > 1 and allk[i - 2] == 'links':
            report.append((i, 'links-merged', [])); continue
        kinds.append(kind)
        rep = []
        if kind == 'title':
            title, _ = slide_title_info(src)
            s = build_title(bld, src, lesson_no, title_text(src))
        elif kind == 'plan': s = build_plan(bld, src)
        elif kind == 'divider':
            s = build_divider(bld, src, div_k); div_k += 1; section += 1
        elif kind == 'wait': s = build_wait(bld, src)
        elif kind == 'summary': s = build_summary(bld, src)
        elif kind == 'reflection': s = build_reflection(bld, src)
        elif kind == 'links':
            grp = [src]; j = i
            while j < len(slides) and allk[j] == 'links': grp.append(slides[j]); j += 1
            s = build_links(bld, grp)
        elif kind == 'bye': s = bld.clone(DONOR['bye'])
        elif kind in SERVICE_KINDS: s = bld.clone(DONOR[kind])  # выбранный на сайте слайд шаблона
        elif kind == 'why': s = build_why(bld, src)
        elif kind == 'think': s = build_think(bld, src, rep)
        else:
            donor = DONOR['content'][section % len(DONOR['content'])]
            s = bld.clone(donor, keep_content=False)
            tx = get_ph(s, 'title').find('p:spPr/a:xfrm/a:ext', NS)
            tx.set('cx', str(RM - LM))
            title = transfer_content(bld, src, s, rep)
            if title: bld.set_title(s, title)
            else:
                rep.append('нет заголовка')
                t_sp = get_ph(s, 'title')
                t_sp.getparent().remove(t_sp)
        if WIDE: rep += wide_notes(src, i, kind)
        if src._element.get('show') == '0' and kind != 'links': s._element.set('show', '0')
        s._kind = kind
        # заметки
        if src.has_notes_slide:
            nt = src.notes_slide.notes_text_frame.text if src.notes_slide.notes_text_frame is not None else ''
            if nt.strip(): s.notes_slide.notes_text_frame.text = nt
        report.append((i, kind, rep))
    # прощальный слайд, если в исходнике его не было
    if 'bye' not in kinds:
        bld.clone(DONOR['bye']); kinds.append('bye'); report.append(('+', 'bye', ['добавлен прощальный слайд']))
    lst = bld.prs.slides._sldIdLst
    new = [e for e in lst if e not in bld.orig]
    kinds_new = kinds
    if 'bye' in kinds and 'links' in kinds and kinds.index('links') > kinds.index('bye'):
        bye_el = new[kinds.index('bye')]
        lst.remove(bye_el); lst.append(bye_el)
    # номера слайдов: поле slidenum заполняем текущим номером (иначе после экспорта оно пустое)
    order = [e for e in lst if e not in bld.orig]
    for n, sid in enumerate(order, 1):
        sl = bld.prs.part.related_part(sid.get(R + 'id'))._element
        for f in sl.iter(A + 'fld'):
            if f.get('type') == 'slidenum':
                t = f.find('a:t', NS)
                if t is None: t = etree.SubElement(f, A + 't')
                t.text = str(n)
    bld.finish(out_path)
    return report, kinds


if __name__ == '__main__':
    # номер урока: «5.15» → 5 класс, урок 15 (или просто «15» — тогда 5 класс)
    src, out, no = sys.argv[1], sys.argv[2], sys.argv[3]
    if '.' in no: GRADE, no = no.split('.', 1)
    rep, kinds = convert(src, out, no)
    for r in rep: print(r)
