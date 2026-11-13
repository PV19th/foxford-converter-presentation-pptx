# -*- coding: utf-8 -*-
"""Собирает версию «в браузере» в dist/browser/ (запускается в Docker из browser/build.sh).

Получается папка обычных файлов — сайт без серверной части. В ней:
- страница (browser/web/ + оформление сайта app/web/static/style.css);
- Pyodide — Python для браузера, и библиотеки конвертера (скачиваются один раз в dist/browser-cache/);
- шрифт Open Sans из Docker-образа — тот же файл, по которому конвертер считает переносы на сайте;
- code.zip — converter/, app/pipeline.py, app/reskin.py, app/texts.py, профили шаблонов, browser/runner.py;
- базы шаблонов и миниатюры вариантов слайдов;
- meta.json — то, что сайт отдаёт страницей /api/meta;
- robots.txt и .nojekyll — для публикации (сайт не показывается в поиске; GitHub Pages отдаёт файлы как есть).
Скрытые профили ("hidden") в сборку не попадают.
"""
import hashlib, json, os, re, shutil, sys, urllib.request, zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from app import pipeline as P, texts  # noqa: E402

PYODIDE = '314.0.7'  # при смене версии — прогнать browser/tests/run.sh
CDN = 'https://cdn.jsdelivr.net/pyodide/v%s/full/' % PYODIDE
CORE = ['pyodide.mjs', 'pyodide.asm.mjs', 'pyodide.asm.wasm', 'python_stdlib.zip', 'pyodide-lock.json']
PACKAGES = ['lxml', 'pillow', 'typing-extensions']  # библиотеки из набора Pyodide (собраны для браузера)
PYPI = ['python-pptx', 'xlsxwriter']                # чистый Python с pypi.org; версии — из constraints.txt
FONTS = Path('/usr/share/fonts/truetype/open-sans')
CODE = ['app/__init__.py', 'app/pipeline.py', 'app/reskin.py', 'app/texts.py', 'browser/runner.py',
        'converter/convert.py', 'converter/engine.py', 'converter/common.py', 'converter/shrink.py', 'converter/verify.py']
OUT, CACHE = ROOT / 'dist/browser', ROOT / 'dist/browser-cache'


def fetch(url, name, sha256=None):
    """Файл из интернета → dist/browser-cache/ (второй раз не скачивается). Контрольная сумма проверяется."""
    f = CACHE / name
    if not f.exists():
        print('  скачиваю', name, flush=True)
        req = urllib.request.Request(url, headers={'User-Agent': 'build'})
        with urllib.request.urlopen(req, timeout=120) as r: data = r.read()
        f.write_bytes(data)
    if sha256 and hashlib.sha256(f.read_bytes()).hexdigest() != sha256:
        f.unlink()
        raise SystemExit('Файл %s скачался с ошибкой (не сошлась контрольная сумма). Запустите сборку ещё раз.' % name)
    return f


def pinned(package):
    """Версия библиотеки из constraints.txt — та же, что на сайте."""
    name = re.escape(package).replace(r'\-', '[-_]')  # typing-extensions в constraints.txt — typing_extensions
    m = re.search(r'^%s==(\S+)' % name, (ROOT / 'constraints.txt').read_text('utf8'), re.M | re.I)
    if not m: raise SystemExit('В constraints.txt нет версии %s' % package)
    return m.group(1)


def digest(paths):
    h = hashlib.sha1()
    for p in paths: h.update(p.read_bytes())
    return h.hexdigest()[:8]


def main():
    # папку не удаляем, а очищаем: иначе открытый ./browser/preview.sh перестаёт видеть файлы
    for p in (OUT.iterdir() if OUT.exists() else []):
        shutil.rmtree(p) if p.is_dir() else p.unlink()
    for d in (OUT / 'pyodide', OUT / 'wheels', OUT / 'fonts', CACHE): d.mkdir(parents=True, exist_ok=True)

    print('Pyodide', PYODIDE)
    for name in CORE: shutil.copy(fetch(CDN + name, '%s-%s' % (PYODIDE, name)), OUT / 'pyodide' / name)
    lock = json.loads((OUT / 'pyodide/pyodide-lock.json').read_text('utf8'))['packages']
    for p in PACKAGES:
        shutil.copy(fetch(CDN + lock[p]['file_name'], lock[p]['file_name'], lock[p]['sha256']), OUT / 'pyodide')
        print('  %s %s (на сайте — %s)' % (p, lock[p]['version'], pinned(p)))
    wheels = []
    for p in PYPI:
        with urllib.request.urlopen('https://pypi.org/pypi/%s/%s/json' % (p, pinned(p)), timeout=60) as r:
            whl = next(u for u in json.load(r)['urls'] if u['filename'].endswith('-none-any.whl'))
        shutil.copy(fetch(whl['url'], whl['filename'], whl['digests']['sha256']), OUT / 'wheels')
        wheels.append('wheels/' + whl['filename'])
        print('  %s %s' % (p, pinned(p)))

    for f in ('OpenSans-Regular.ttf', 'OpenSans-Bold.ttf'): shutil.copy(FONTS / f, OUT / 'fonts' / f)

    # шаблоны: профили, базы, миниатюры
    profiles = {k: v for k, v in P.load_templates().items() if not v.get('hidden')}
    code = [ROOT / f for f in CODE] + [P.TEMPLATES / k / 'profile.json' for k in profiles]
    with zipfile.ZipFile(OUT / 'code.zip', 'w', zipfile.ZIP_DEFLATED) as z:
        for f in code: z.write(f, f.relative_to(ROOT).as_posix())
    tpls = []
    for k, v in sorted(profiles.items(), key=lambda kv: kv[1].get('order', 99)):
        if 'base_recipe' in v:
            P.run([sys.executable, str(ROOT / 'tools/build_base.py'), str(P.TEMPLATES / k)], 'сборка базы шаблона')
        grades = P.profile_grades(v)
        bases = {g: P.grade_settings(v, g)['base'] for g in grades} if P.uses_lesson(v) else {'': v['base']}
        for b in set(bases.values()):
            (OUT / b).parent.mkdir(parents=True, exist_ok=True); shutil.copy(ROOT / b, OUT / b)
        thumbs = sorted((P.TEMPLATES / k / 'thumbs').glob('*.jpg'))
        (OUT / 'thumbs' / k).mkdir(parents=True)
        for t in thumbs: shutil.copy(t, OUT / 'thumbs' / k / t.name)
        # то же, что /api/meta на сайте (app/main.py), плюс "bases": какой файл базы нужен классу
        tpls.append({'id': k, 'name': v['name'], 'grades': grades, 'group': v.get('group'), 'lesson': P.uses_lesson(v),
                     'choices': [dict(c, key=key) for key, c in v.get('choices', {}).items()],
                     'ratio': v.get('ratio', '4 / 3'), 'thumbs_v': digest(thumbs), 'bases': bases})
        print('шаблон «%s»: %s' % (v['name'], ', '.join(sorted(set(bases.values())))))

    web = Path(__file__).resolve().parent / 'web'
    shutil.copy(ROOT / 'app/web/static/style.css', OUT / 'style.css')
    for f in web.iterdir():
        if f.name != 'index.html' and not f.name.startswith('.'): shutil.copy(f, OUT / f.name)
    # версия сборки стоит в адресах файлов (?v=…): после обновления браузер не возьмёт старую копию из кеша
    version = digest(sorted(p for p in OUT.rglob('*') if p.is_file()))
    meta = {'templates': tpls, 'steps': [s for s in texts.STEPS if s[0] != 'render'], 'manual_checks': texts.MANUAL_CHECKS,
            'subjects': texts.SUBJECTS, 'contact': os.environ.get('SUPPORT_CONTACT', ''),
            'max_upload_mb': int(os.environ.get('MAX_UPLOAD_MB') or 100),
            'packages': PACKAGES, 'wheels': wheels, 'version': version}
    (OUT / 'meta.json').write_text(json.dumps(meta, ensure_ascii=False), 'utf8')
    (OUT / 'index.html').write_text((web / 'index.html').read_text('utf8').replace('?v=0', '?v=' + version), 'utf8')
    # сайт публичный, но это рабочий инструмент: поисковикам его показывать незачем (и <meta name="robots"> на странице)
    (OUT / 'robots.txt').write_text('User-agent: *\nDisallow: /\n', 'utf8')
    (OUT / '.nojekyll').write_text('', 'utf8')  # GitHub Pages: отдавать файлы как есть, без своей обработки
    size = sum(p.stat().st_size for p in OUT.rglob('*') if p.is_file()) / 1e6
    print('Готово: dist/browser/ — %.0f МБ, версия %s' % (size, version))


if __name__ == '__main__':
    main()
