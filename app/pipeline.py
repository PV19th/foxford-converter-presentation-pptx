# -*- coding: utf-8 -*-
"""Обработка одной презентации: convert.py → shrink.py → verify.py → превью → памятка.

Конвертер не трогаем: запускаем его скрипты как из командной строки.
Запуск без сайта:
    python3 -m app.pipeline старый.pptx папка_результата 5.21 [--template segment-5-11]
"""
import argparse, ast, json, logging, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path
from PIL import Image, ImageFont

ROOT = Path(__file__).resolve().parent.parent
CONVERTER = ROOT / 'converter'
TEMPLATES = ROOT / 'templates'
LESSON_RE = re.compile(r'^(5|6|7|8|9|10|11)\.\d{1,2}$')
# shrink.py сжимает картинки, только если файл больше этого. 18,5 МБ было ограничением старой среды
# (конвертер запускали в чате Claude, крупнее файл не сохранялся); сайту оно не нужно — решение автора
SIZE_LIMIT_MB = '50'
DEFAULT_TEMPLATE = 'segment-5-11'
TAG_FONT = '/usr/share/fonts/truetype/open-sans/OpenSans-%s.ttf'  # шрифт тегов титула, как в engine.py

log = logging.getLogger('pipeline')


class ConversionError(Exception):
    """Ошибка, которую можно показать пользователю как есть."""


# ---------------------------------------------------------------- шаблоны
def load_templates():
    """{id: профиль} из templates/<id>/profile.json."""
    out = {}
    for f in sorted(TEMPLATES.glob('*/profile.json')):
        out[f.parent.name] = json.loads(f.read_text('utf8'))
    return out


def grade_settings(profile, grade):
    """Настройки класса; классы из "grade_aliases" берут настройки другого класса (7–11 → как 6)."""
    grade = profile.get('grade_aliases', {}).get(grade, grade)
    return profile['grades'].get(grade)


def profile_grades(profile):
    return sorted(set(profile['grades']) | set(profile.get('grade_aliases', {})), key=int)


def converter_env(profile, grade, choices=None):
    g = grade_settings(profile, grade)
    env = dict(os.environ, CONV_DIR=str(ROOT), CONV_BASE=str(ROOT / g['base']))
    env.pop('CONV_DONOR', None)
    donor = dict(g.get('donor', {}))
    donor.update(choices or {})  # выбор пользователя: {'title': 9, 'bye': 128}
    if donor: env['CONV_DONOR'] = json.dumps(donor)
    return env


def ensure_base(template_id):
    """Профили на фирменном шаблоне собирают свою базу скриптом (быстро, если база свежая)."""
    if 'base_recipe' in load_templates()[template_id]:
        run([sys.executable, str(ROOT / 'tools/build_base.py'), str(TEMPLATES / template_id)], 'сборка базы шаблона')


P_ = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
A_ = '{http://schemas.openxmlformats.org/drawingml/2006/main}'


def _sp_text(sp):
    return ' '.join(''.join(t.text or '' for t in sp.iter(A_ + 't')).split())


def detect_subject(pptx, subjects):
    """Предмет из тега на титуле старой презентации или None.
    Титул бывает не первым слайдом (старый шаблон 10×7,5″): ищем в первых 8 слайдах тег из списка,
    а если его нет — тег слева от «N класс» в той же строке («English», «Алгебра»)."""
    from pptx import Presentation
    for s in list(Presentation(pptx).slides)[:8]:
        sps = [sp for sp in s._element.iter(P_ + 'sp')]
        hit = next((_sp_text(sp) for sp in sps if _sp_text(sp) in subjects), None)
        if hit: return hit
        grade = next((sp for sp in sps if re.fullmatch(r'\d{1,2} ?класс', _sp_text(sp))), None)
        if grade is None: continue
        box = lambda sp: next(((int(o.get('x')), int(o.get('y'))) for o in sp.iter(A_ + 'off')), None)
        g = box(grade)
        left = [(g[0] - b[0], _sp_text(sp)) for sp in sps if sp is not grade and _sp_text(sp)
                and (b := box(sp)) and abs(b[1] - g[1]) < 150000 and b[0] < g[0]] if g else []
        if left: return min(left)[1]
    return None


def detect_lesson(pptx):
    """Номер урока «класс.урок» по тегам титула («9 класс» и «Урок 14» / «Lesson 14») или None."""
    from pptx import Presentation
    for s in list(Presentation(pptx).slides)[:8]:
        texts = [_sp_text(sp) for sp in s._element.iter(P_ + 'sp')]
        g = next((m.group(1) for t in texts if (m := re.fullmatch(r'(\d{1,2}) ?класс', t))), None)
        n = next((m.group(1) for t in texts if (m := re.fullmatch(r'(?:Урок|Занятие|Lesson) (\d{1,2})', t))), None)
        if g and n: return '%s.%s' % (g, n)
    return None


def set_subject(pptx, subject, slide_no=1):
    """Конвертер ставит на титул тег «История»; для других предметов меняем его текст."""
    from pptx import Presentation
    p = Presentation(pptx)
    for sp in p.slides[slide_no - 1]._element.iter(P_ + 'sp'):
        ts = list(sp.iter(A_ + 't'))
        if ''.join(t.text or '' for t in ts).strip() == 'История':
            ts[0].text = subject
            for t in ts[1:]: t.text = ''
            widen_tag(sp, subject)
    p.save(pptx)


def widen_tag(sp, text):
    """Длинный предмет («Обществознание») не влезает в тег и переносится на две строки:
    тег расширяется по тексту, теги справа в том же ряду («9 класс», «Урок 23») сдвигаются. Уже — не делаем."""
    xfrm, body = sp.find(P_ + 'spPr/' + A_ + 'xfrm'), sp.find('.//' + A_ + 'bodyPr')
    if xfrm is None or body is None: return
    off, ext = xfrm.find(A_ + 'off'), xfrm.find(A_ + 'ext')
    x, y, w, h = int(off.get('x')), int(off.get('y')), int(ext.get('cx')), int(ext.get('cy'))
    sz = next((int(r.get('sz')) / 100 for r in sp.iter(A_ + 'rPr', A_ + 'endParaRPr') if r.get('sz')), 14)
    bold = any(r.get('b') == '1' for r in sp.iter(A_ + 'rPr'))
    font = ImageFont.truetype(TAG_FONT % ('Bold' if bold else 'Regular'), int(sz * 20))
    # у скруглённой плашки поле для текста уже на 0,29 радиуса скругления с каждой стороны
    geom, corner = sp.find(P_ + 'spPr/' + A_ + 'prstGeom[@prst="roundRect"]'), 0
    if geom is not None:
        gd = geom.find('.//' + A_ + 'gd')
        corner = 0.29289 * min(w, h) * (int(gd.get('fmla').split()[-1]) if gd is not None else 16667) / 100000
    need = int(font.getlength(text) / 20 * 12700 + 2 * corner) + int(body.get('lIns', 91440)) + int(body.get('rIns', 91440)) + 45720
    if need <= w: return
    ext.set('cx', str(need))
    for other in sp.getparent().iterchildren(P_ + 'sp'):  # соседние теги: тот же ряд, правее
        o = other.find(P_ + 'spPr/' + A_ + 'xfrm/' + A_ + 'off')
        if o is not None and other is not sp and abs(int(o.get('y')) - y) < 45720 and int(o.get('x')) > x:
            o.set('x', str(int(o.get('x')) + need - w))


def slide_words(pptx, numbers):
    """Слова слайдов с этими номерами — так же, как их считает verify.py."""
    from collections import Counter
    from pptx import Presentation
    c = Counter()
    for i, s in enumerate(Presentation(pptx).slides, 1):
        if i in numbers:
            t = ' '.join(x.text or '' for x in s._element.iter(A_ + 't'))
            c.update(w for w in (w.lower() for w in re.findall(r'[A-Za-zА-Яа-яЁё0-9]+', t)) if not re.fullmatch(r'\d{1,2}', w))
    return c


def find_stubs(pptx, phrases):
    """Тексты-заглушки шаблона («Тема урока», «Текст»…), попавшие в результат: [{'slide', 'text'}]."""
    from pptx import Presentation
    found = []
    for n, s in enumerate(Presentation(pptx).slides, 1):
        for p in s._element.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}p'):
            t = ''.join(x.text or '' for x in p.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t')).strip()
            hit = next((ph for ph in phrases if t == ph or (len(ph) >= 12 and ph in t)), None)
            if hit: found.append({'slide': n, 'text': 'остался текст-заглушка шаблона «%s»' % hit})
    return found


# ---------------------------------------------------------------- запуск шагов
def run(cmd, step, cwd=None, env=None, timeout=600):
    log.info('%s: %s', step, ' '.join(map(str, cmd)))
    try:
        r = subprocess.run(cmd, cwd=cwd, env=env, capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        log.error('%s: превышено время (%d с)', step, timeout)
        raise ConversionError('Обработка заняла слишком много времени (шаг «%s»).' % step)
    if r.returncode != 0:
        log.error('%s: код %d\n%s\n%s', step, r.returncode, r.stdout, r.stderr)
        raise ConversionError('Не удалось выполнить шаг «%s».' % step)
    if r.stderr.strip(): log.warning('%s stderr:\n%s', step, r.stderr)
    return r.stdout


# ---------------------------------------------------------------- разбор отчётов
def parse_report(stdout):
    """Строки convert.py вида (номер_в_исходнике, тип, [замечания]) → список слайдов нового файла.

    Номера в новом файле считаются так же, как их расставляет convert.py:
    склеенные слайды ссылок выпадают, прощальный слайд уходит в конец, если стоял перед ссылками."""
    rows = []
    for line in stdout.splitlines():
        line = line.strip()
        if not line.startswith('('): continue
        try: src, kind, notes = ast.literal_eval(line)
        except (ValueError, SyntaxError): log.warning('не разобрана строка отчёта: %s', line); continue
        if kind == 'links-merged': continue
        rows.append({'src': src if isinstance(src, int) else None, 'kind': kind, 'notes': list(notes)})
    kinds = [r['kind'] for r in rows]
    if 'bye' in kinds and 'links' in kinds and kinds.index('links') > kinds.index('bye'):
        rows.append(rows.pop(kinds.index('bye')))
    for n, r in enumerate(rows, 1): r['slide'] = n
    return rows


def parse_verify(stdout):
    out = {'MISSING': {}, 'EXTRA': {}}
    for line in stdout.splitlines():
        key, _, rest = line.partition(' ')
        if key in out: out[key] = ast.literal_eval(rest.strip())
    return out['MISSING'], out['EXTRA']


def manual_notes(slides):
    """Замечания конвертера → «Добавьте стикеры» и «Проверьте слайды»."""
    stickers, checks = [], []
    for s in slides:
        for note in s['notes']:
            if note.startswith('СТИКЕР ВРУЧНУЮ:'):
                what = note.split(':', 1)[1].replace('(место оставлено)', '').strip()
                stickers.append({'slide': s['slide'], 'text': what})
                continue
            m = re.match(r'масштаб ([\d.]+), кегль (\d+)', note)
            if m:
                if m.group(2) == '14': text = 'кегль уменьшен до 14 — проверьте, не тесно ли'
                else: text = 'не влезало по высоте — сжато до %d %%, проверьте, не мелко ли' % round(float(m.group(1)) * 100)
            elif note.startswith('контент выше зоны'): text = 'содержимое не помещается по высоте'
            elif note.startswith('низ контента'): text = 'содержимое у самого нижнего края'
            elif note == 'нет заголовка': text = 'у слайда нет заголовка'
            elif note == 'добавлен прощальный слайд': continue
            else: text = note
            checks.append({'slide': s['slide'], 'text': text})
    return stickers, checks


# ---------------------------------------------------------------- превью
def render(pptx, outdir, width=1280, thumb_width=360):
    """pptx → PDF (LibreOffice) → JPEG по слайдам + миниатюры. Возвращает имена файлов слайдов.
    soffice нельзя запускать параллельно — вызывающий код обязан идти по одной задаче."""
    outdir = Path(outdir)
    slides_dir, thumbs_dir = outdir / 'slides', outdir / 'thumbs'
    for d in (slides_dir, thumbs_dir):
        shutil.rmtree(d, ignore_errors=True); d.mkdir(parents=True)
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        run(['soffice', '-env:UserInstallation=file://%s/lo' % tmp, '--headless',
             '--convert-to', 'pdf', '--outdir', str(tmp), str(pptx)], 'превью (PDF)', timeout=600)
        pdf = tmp / (Path(pptx).stem + '.pdf')
        if not pdf.exists(): raise ConversionError('Не удалось построить превью слайдов.')
        run(['pdftoppm', '-jpeg', '-jpegopt', 'quality=85', '-scale-to-x', str(width), '-scale-to-y', '-1',
             str(pdf), str(tmp / 's')], 'превью (картинки)', timeout=600)
        pages = sorted(tmp.glob('s-*.jpg'), key=lambda p: int(p.stem.split('-')[1]))
        names = []
        for n, p in enumerate(pages, 1):
            name = '%03d.jpg' % n
            shutil.move(str(p), slides_dir / name)
            im = Image.open(slides_dir / name)
            im.thumbnail((thumb_width, thumb_width)); im.save(thumbs_dir / name, quality=85)
            names.append(name)
    contact_sheet([thumbs_dir / n for n in names], outdir / 'contact.jpg')
    return names


def contact_sheet(paths, out, cols=6, gap=8):
    if not paths: return
    ims = [Image.open(p) for p in paths]
    w, h = ims[0].size
    rows = (len(ims) + cols - 1) // cols
    sheet = Image.new('RGB', (cols * w + (cols + 1) * gap, rows * h + (rows + 1) * gap), 'white')
    for i, im in enumerate(ims):
        sheet.paste(im, (gap + (i % cols) * (w + gap), gap + (i // cols) * (h + gap)))
    sheet.save(out, quality=85)


# ---------------------------------------------------------------- вся обработка
def check_input(path):
    if not zipfile.is_zipfile(path):
        raise ConversionError('Файл не похож на презентацию PowerPoint (.pptx).')
    with zipfile.ZipFile(path) as z:
        if 'ppt/presentation.xml' not in z.namelist():
            raise ConversionError('Файл не похож на презентацию PowerPoint (.pptx).')
        themes = b' '.join(z.read(n) for n in z.namelist() if n.startswith('ppt/theme/theme'))
    # тема нового шаблона называется «Шаблон 5–11 (2026) — Фоксфорд»
    if 'Шаблон 5–11 (2026)'.encode('utf8') in themes:
        raise ConversionError('Похоже, эта презентация уже в новом шаблоне. Загрузите старую версию урока.')


def check_choices(profile, choices):
    """Выбор слайдов пользователем ({'title': 9}) должен быть из вариантов профиля."""
    allowed = profile.get('choices', {})
    for key, n in (choices or {}).items():
        if key not in allowed or n not in allowed[key]['slides']:
            raise ConversionError('У шаблона «%s» нет такого варианта слайда.' % profile['name'])


def process(src, outdir, lesson, template=DEFAULT_TEMPLATE, progress=lambda step: None, choices=None, subject=None):
    """Обрабатывает одну презентацию. В outdir появятся result.pptx, result.json, slides/, thumbs/, contact.jpg."""
    # абсолютные пути: шаги конвертера работают из папки converter/
    src, outdir = Path(src).resolve(), Path(outdir).resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    if not LESSON_RE.match(lesson):
        raise ConversionError('Номер урока нужно писать так: класс.урок, например 5.21 или 6.18.')
    lesson = '%s.%d' % (lesson.split('.')[0], int(lesson.split('.')[1]))  # «9.07» → «Урок 7»
    templates = load_templates()
    if template not in templates: raise ConversionError('Неизвестный шаблон: %s.' % template)
    profile = templates[template]
    grade = lesson.split('.')[0]
    if grade_settings(profile, grade) is None:
        raise ConversionError('Шаблон «%s» пока не поддерживает %s класс.' % (profile['name'], grade))
    check_choices(profile, choices)
    check_input(src)
    source = outdir / 'source.pptx'
    if Path(src).resolve() != source.resolve(): shutil.copyfile(src, source)
    raw, result = outdir / 'converted.pptx', outdir / 'result.pptx'

    progress('convert')
    ensure_base(template)
    out = run([sys.executable, 'convert.py', str(source), str(raw), lesson], 'конвертация',
              cwd=CONVERTER, env=converter_env(profile, grade, choices))
    slides = parse_report(out)
    from . import texts
    subject = subject or detect_subject(source, texts.SUBJECTS) or 'История'
    title_no = next((r['slide'] for r in slides if r['kind'] == 'title'), 1)
    if subject != 'История': set_subject(raw, subject, title_no)
    progress('shrink')
    run([sys.executable, 'shrink.py', str(raw), str(result), SIZE_LIMIT_MB], 'сжатие', cwd=CONVERTER)
    progress('verify')
    missing, extra = parse_verify(run([sys.executable, 'verify.py', str(source), str(result)], 'сверка текста',
                                      cwd=CONVERTER))
    # слайды, заменённые по выбору пользователя слайдами шаблона, — не расхождения: убираем их слова из сверки
    replaced = [r for r in slides if r['kind'] in (choices or {}) and profile.get('choices', {}).get(r['kind'], {}).get('optional')]
    if replaced:
        from collections import Counter
        missing = dict(Counter(missing) - slide_words(source, {r['src'] for r in replaced}))
        extra = dict(Counter(extra) - slide_words(result, {r['slide'] for r in replaced}))
    progress('render')
    previews = render(result, outdir)
    stickers, checks = manual_notes(slides)
    # необязательный выбор (мотивашка, рефлексия, оценка урока) заменяет слайд старой презентации — если он там есть
    kinds = {r['kind'] for r in slides}
    for key in choices or {}:
        c = profile.get('choices', {}).get(key, {})
        if c.get('optional') and key not in kinds:
            checks.append({'slide': None, 'text': 'выбранный слайд «%s» не поставлен: в старой презентации нет такого слайда'
                                                  % c['label']})
    checks = sorted(checks + find_stubs(result, profile.get('stub_phrases', [])), key=lambda c: c['slide'] or 0)
    data = {'lesson': lesson, 'template': template, 'template_name': profile['name'], 'choices': choices or {}, 'subject': subject,
            'slides': slides, 'previews': previews, 'missing': missing, 'extra': extra,
            'stickers': stickers, 'checks': checks,
            'size_mb': round(result.stat().st_size / 1e6, 1)}
    (outdir / 'result.json').write_text(json.dumps(data, ensure_ascii=False, indent=1), 'utf8')
    return data


def parse_choices(items):
    """['title=9', 'bye=128'] → {'title': 9, 'bye': 128}"""
    return {k: int(v) for k, v in (i.split('=', 1) for i in items)}


def main():
    ap = argparse.ArgumentParser(description='Перенос старой презентации в новый шаблон')
    ap.add_argument('src'); ap.add_argument('outdir'); ap.add_argument('lesson', help='класс.урок, например 5.21')
    ap.add_argument('--template', default=DEFAULT_TEMPLATE)
    ap.add_argument('--subject', help='предмет для тега на титуле (по умолчанию — из старой презентации)')
    ap.add_argument('--choice', action='append', default=[], metavar='КЛЮЧ=НОМЕР',
                    help='вариант слайда шаблона, например title=9 или bye=128')
    a = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
    try:
        d = process(a.src, a.outdir, a.lesson, a.template, progress=lambda s: log.info('шаг: %s', s),
                    choices=parse_choices(a.choice), subject=a.subject)
    except ConversionError as e:
        print('ОШИБКА:', e, file=sys.stderr); sys.exit(1)
    print('слайдов:', len(d['previews']), '| размер МБ:', d['size_mb'])
    print('MISSING:', d['missing']); print('EXTRA:', d['extra'])
    for s in d['stickers']: print('стикер: слайд %d — %s' % (s['slide'], s['text']))
    for c in d['checks']: print('проверить: слайд %s — %s' % (c['slide'] or '—', c['text']))


if __name__ == '__main__':
    main()
