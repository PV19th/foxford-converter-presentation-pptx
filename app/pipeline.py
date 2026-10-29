# -*- coding: utf-8 -*-
"""Обработка одной презентации: convert.py → shrink.py → verify.py → превью → памятка.

Конвертер не трогаем: запускаем его скрипты как из командной строки.
Запуск без сайта:
    python3 -m app.pipeline старый.pptx папка_результата 5.21 [--template segment-5-11]
"""
import argparse, ast, json, logging, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path
from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
CONVERTER = ROOT / 'converter'
TEMPLATES = ROOT / 'templates'
LESSON_RE = re.compile(r'^(5|6)\.\d{1,2}$')
SIZE_LIMIT_MB = '18.5'
DEFAULT_TEMPLATE = 'segment-5-11'

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


def converter_env(profile, grade, choices=None):
    g = profile['grades'][grade]
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
                else: text = 'содержимое уменьшено до %d %%' % round(float(m.group(1)) * 100)
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
        tag = re.search(rb'<p:sldSz[^>]*>', z.read('ppt/presentation.xml'))
    size = dict(re.findall(rb'(cx|cy)="(\d+)"', tag.group(0))) if tag else {}
    # 10×7,5″ — размер нового шаблона: такой файл уже перенесён, повторный перенос его испортит
    if size == {b'cx': b'9144000', b'cy': b'6858000'}:
        raise ConversionError('Похоже, эта презентация уже в новом шаблоне. Загрузите старую версию урока.')


def check_choices(profile, choices):
    """Выбор слайдов пользователем ({'title': 9}) должен быть из вариантов профиля."""
    allowed = profile.get('choices', {})
    for key, n in (choices or {}).items():
        if key not in allowed or n not in allowed[key]['slides']:
            raise ConversionError('У шаблона «%s» нет такого варианта слайда.' % profile['name'])


def process(src, outdir, lesson, template=DEFAULT_TEMPLATE, progress=lambda step: None, choices=None):
    """Обрабатывает одну презентацию. В outdir появятся result.pptx, result.json, slides/, thumbs/, contact.jpg."""
    # абсолютные пути: шаги конвертера работают из папки converter/
    src, outdir = Path(src).resolve(), Path(outdir).resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    if not LESSON_RE.match(lesson):
        raise ConversionError('Номер урока нужно писать так: класс.урок, например 5.21 или 6.18.')
    templates = load_templates()
    if template not in templates: raise ConversionError('Неизвестный шаблон: %s.' % template)
    profile = templates[template]
    grade = lesson.split('.')[0]
    if grade not in profile['grades']:
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
    progress('shrink')
    run([sys.executable, 'shrink.py', str(raw), str(result), SIZE_LIMIT_MB], 'сжатие', cwd=CONVERTER)
    progress('verify')
    missing, extra = parse_verify(run([sys.executable, 'verify.py', str(source), str(result)], 'сверка текста',
                                      cwd=CONVERTER))
    progress('render')
    previews = render(result, outdir)
    stickers, checks = manual_notes(slides)
    checks = sorted(checks + find_stubs(result, profile.get('stub_phrases', [])), key=lambda c: c['slide'])
    data = {'lesson': lesson, 'template': template, 'template_name': profile['name'], 'choices': choices or {},
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
    ap.add_argument('--choice', action='append', default=[], metavar='КЛЮЧ=НОМЕР',
                    help='вариант слайда шаблона, например title=9 или bye=128')
    a = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
    try:
        d = process(a.src, a.outdir, a.lesson, a.template, progress=lambda s: log.info('шаг: %s', s),
                    choices=parse_choices(a.choice))
    except ConversionError as e:
        print('ОШИБКА:', e, file=sys.stderr); sys.exit(1)
    print('слайдов:', len(d['previews']), '| размер МБ:', d['size_mb'])
    print('MISSING:', d['missing']); print('EXTRA:', d['extra'])
    for s in d['stickers']: print('стикер: слайд %d — %s' % (s['slide'], s['text']))
    for c in d['checks']: print('проверить: слайд %d — %s' % (c['slide'], c['text']))


if __name__ == '__main__':
    main()
