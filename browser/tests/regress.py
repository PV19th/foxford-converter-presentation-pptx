# -*- coding: utf-8 -*-
"""Проверка версии «в браузере»: те же тестовые презентации, тот же «отпечаток», что у tests/run_tests.py.
Выполняется внутри Pyodide (его запускает regress.mjs). Эталоны — общие с сайтом, tests/baseline/:
«совпадает» значит, что в браузере получается тот же файл, что на сайте."""
import json, logging, re, sys, time

sys.path[:0] = ['/srv/browser', '/srv/tests']
import runner            # noqa: E402 — подменяет запуск шагов в app/pipeline (см. runner.py)
import run_tests as T    # noqa: E402 — отпечаток и номер урока считаем тем же кодом, что тесты сайта

P = runner.P
logging.basicConfig(level=logging.WARNING, format='%(levelname)s %(message)s')


def cases(template):
    """Тестовые презентации профиля — как их перебирает tests/run_tests.py."""
    profile = P.load_templates()[template]
    out = []
    for src in sorted((P.ROOT / profile.get('test_inputs', 'test_inputs/old')).glob('*.pptx')):
        lesson = T.lesson_of(src.name) or P.detect_lesson(src)
        key = lesson or re.sub(r'\W+', '_', src.stem).strip('_')[:40]
        out.append({'key': key, 'lesson': lesson or '11.1', 'src': str(src), 'mb': round(src.stat().st_size / 1e6)})
    return json.dumps(out, ensure_ascii=False)


def run(template, src, key, lesson):
    out = P.ROOT / 'outputs/browser' / template / key
    start = time.time()
    try:
        d = P.process(src, out, lesson, template)
    except P.ConversionError as e:
        return json.dumps({'status': 'ОШИБКА: %s' % e}, ensure_ascii=False)
    bf = (T.BASELINE if template == P.DEFAULT_TEMPLATE else T.BASELINE / template) / (key + '.json')
    status = T.compare_fp(json.loads(bf.read_text('utf8')), T.fingerprint(out / 'converted.pptx')) if bf.exists() else 'нет эталона'
    if any('заглушка' in c['text'] for c in d['checks']): status += ' + ЗАГЛУШКА'
    return json.dumps({'status': status, 'seconds': round(time.time() - start), 'slides': len(d['slides']),
                       'missing': d['missing'], 'size_mb': d['size_mb']}, ensure_ascii=False)
