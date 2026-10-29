# -*- coding: utf-8 -*-
"""Прогон всех test_inputs/old/*.pptx через pipeline (запускается внутри Docker из run_all.sh).

Для каждой презентации:
- результат, превью и памятка → outputs/tests/<урок>/;
- превью эталона методиста (test_inputs/reference) → outputs/tests/<урок>/reference/;
- «отпечаток» результата конвертера сверяется с tests/baseline/<урок>.json
  (--save-baseline — записать текущие отпечатки как эталон);
- сводная страница для сравнения глазами → outputs/tests/index.html.
"""
import argparse, hashlib, html, json, logging, re, subprocess, sys, zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from app import pipeline as P

ROOT = P.ROOT
OLD, REF = ROOT / 'test_inputs/old', ROOT / 'test_inputs/reference'
OUT, BASELINE = ROOT / 'outputs/tests', ROOT / 'tests/baseline'


def lesson_of(name):
    m = re.search(r'(?:^|\D)(\d{1,2})\.(\d{1,2})(?!\d)', name)
    return '%s.%s' % (m.group(1), int(m.group(2))) if m else None


def fingerprint(pptx):
    """sha1 каждой части файла, кроме docProps (там дата сохранения)."""
    with zipfile.ZipFile(pptx) as z:
        parts = {n: hashlib.sha1(z.read(n)).hexdigest()[:12] for n in sorted(z.namelist()) if not n.startswith('docProps/')}
    total = hashlib.sha1(json.dumps(parts, sort_keys=True).encode()).hexdigest()[:12]
    return {'total': total, 'parts': parts}


def compare_fp(old, new):
    if old['total'] == new['total']: return 'совпадает'
    a, b = old['parts'], new['parts']
    diff = sorted(n for n in set(a) | set(b) if a.get(n) != b.get(n))
    return 'ОТЛИЧАЕТСЯ: ' + ', '.join(diff[:15]) + (' …' if len(diff) > 15 else '')


def lesson_page(lesson, d, ref_previews, ref_missing, fp_status):
    rows = []
    for i in range(max(len(d['previews']), len(ref_previews))):
        ours = '<img src="slides/%s">' % d['previews'][i] if i < len(d['previews']) else ''
        ref = '<img src="reference/slides/%s">' % ref_previews[i] if i < len(ref_previews) else ''
        rows.append('<tr><td class="n">%d</td><td>%s</td><td>%s</td></tr>' % (i + 1, ours, ref))
    li = lambda xs: ''.join('<li>слайд %s — %s</li>' % (x['slide'] or '—', html.escape(x['text'])) for x in xs) or '<li>нет</li>'
    return '''<!doctype html><meta charset="utf-8"><title>%(l)s</title>
<style>body{font:14px/1.5 -apple-system,sans-serif;margin:24px;color:#000}h1{font-size:20px;text-transform:uppercase}
table{border-collapse:collapse;width:100%%}td,th{border-top:1px solid #000;padding:8px;vertical-align:top;text-align:left}
td img{width:100%%;border:1px solid #ddd}td.n{width:32px;font-family:monospace}code{font-size:12px}</style>
<p><a href="index.html">← все уроки</a></p><h1>Урок %(l)s</h1>
<p>Отпечаток конвертера: <b>%(fp)s</b></p>
<p>MISSING (есть в старой, нет в новой): <code>%(m)s</code><br>EXTRA: <code>%(e)s</code><br>
Для сравнения — MISSING у эталона методиста: <code>%(rm)s</code></p>
<h2>Добавьте стикеры</h2><ul>%(st)s</ul><h2>Проверьте слайды</h2><ul>%(ch)s</ul>
<table><tr><th>№</th><th>Конвертер</th><th>Эталон методиста</th></tr>%(rows)s</table>''' % dict(
        l=lesson, fp=html.escape(fp_status), m=html.escape(str(d['missing'])), e=html.escape(str(d['extra'])),
        rm=html.escape(str(ref_missing)), st=li(d['stickers']), ch=li(d['checks']), rows=''.join(rows))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--save-baseline', action='store_true')
    ap.add_argument('--only', help='только этот урок, например 5.21')
    ap.add_argument('--template', default=P.DEFAULT_TEMPLATE, help='профиль шаблона из templates/')
    ap.add_argument('--choice', action='append', default=[], metavar='КЛЮЧ=НОМЕР', help='вариант слайда, например bye=128')
    a = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
    global OUT, BASELINE
    if a.template != P.DEFAULT_TEMPLATE:  # у каждого профиля свои результаты и эталоны
        OUT, BASELINE = OUT / a.template, BASELINE / a.template
    OUT.mkdir(parents=True, exist_ok=True); BASELINE.mkdir(parents=True, exist_ok=True)
    summary, failed = [], False
    old = ROOT / P.load_templates()[a.template].get('test_inputs', 'test_inputs/old')  # у 16:9 — своя папка
    for src in sorted(old.glob('*.pptx')):
        lesson = lesson_of(src.name) or P.detect_lesson(src)  # номер из имени файла или с титула
        # уроки ЕГЭ и интенсивы без номера: результаты — по имени файла, на титул — «11 класс · Урок 1»
        key = lesson or re.sub(r'\W+', '_', src.stem).strip('_')[:40]
        lesson = lesson or '11.1'
        if a.only and a.only not in (key, lesson): continue
        print('\n=== %s — %s' % (key, src.name), flush=True)
        out = OUT / key
        try:
            d = P.process(src, out, lesson, a.template, choices=P.parse_choices(a.choice))
        except P.ConversionError as e:
            print('ОШИБКА:', e); failed = True; summary.append((key, 'ОШИБКА: %s' % e, 0)); continue
        fp = fingerprint(out / 'converted.pptx')
        bf = BASELINE / (key + '.json')
        if a.save_baseline or not bf.exists():
            bf.write_text(json.dumps(fp, indent=1), 'utf8'); fp_status = 'записан как эталон'
        else:
            fp_status = compare_fp(json.loads(bf.read_text('utf8')), fp)
            if fp_status != 'совпадает': failed = True
        refs = [r for r in REF.glob('*.pptx') if lesson_of(r.name) == lesson]
        ref_previews, ref_missing = [], {}
        if refs:
            ref_previews = P.render(refs[0], out / 'reference')
            ref_missing, _ = P.parse_verify(subprocess.run(
                [sys.executable, 'verify.py', str(src), str(refs[0])], cwd=P.CONVERTER,
                capture_output=True, text=True).stdout)
        (out / 'index.html').write_text(lesson_page(key, d, ref_previews, ref_missing, fp_status), 'utf8')
        print('слайдов: %d (эталон: %d) | отпечаток: %s' % (len(d['previews']), len(ref_previews), fp_status))
        print('MISSING:', d['missing']); print('EXTRA:', d['extra'])
        if any('заглушка' in c['text'] for c in d['checks']): failed = True
        for s in d['stickers']: print('  стикер: слайд %d — %s' % (s['slide'], s['text']))
        for c in d['checks']: print('  проверить: слайд %s — %s' % (c['slide'] or '—', c['text']))
        summary.append((key, fp_status, len(d['missing'])))
    links = ''.join('<li><a href="%s/index.html">%s</a> — отпечаток: %s; MISSING слов: %d</li>' % (l, l, html.escape(s), m)
                    for l, s, m in summary)
    (OUT / 'index.html').write_text('<!doctype html><meta charset="utf-8"><title>Тесты</title>'
                                    '<style>body{font:14px/1.6 -apple-system,sans-serif;margin:24px}</style>'
                                    '<h1>ТЕСТОВЫЕ ПРОГОНЫ</h1><ul>%s</ul>' % links, 'utf8')
    print('\nСводка:'); [print(' ', l, '—', s) for l, s, _ in summary]
    sys.exit(1 if failed else 0)


if __name__ == '__main__':
    main()
