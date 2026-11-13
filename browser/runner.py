# -*- coding: utf-8 -*-
"""Версия «в браузере» (альтернатива сайту, см. browser/README.md): обработка одной презентации
прямо в браузере методиста, без сервера. Python здесь — Pyodide (Python, собранный для браузера).

Логика та же, что на сайте: вызывается app/pipeline.process. Отличий три, и все они — в этом файле:
- шаги конвертера (convert.py, shrink.py, verify.py) идут не отдельными процессами (в браузере их нет),
  а в этом же Python — run_inline;
- превью не строится: LibreOffice в браузере нет;
- базы шаблонов собраны заранее (browser/build.py), tools/build_base.py не запускается.
converter/ и app/ не меняются.
"""
import contextlib, io, json, logging, os, re, runpy, sys, traceback, unicodedata
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent  # в браузере проект разложен как в Docker: /srv/app, /srv/converter…
if str(ROOT) not in sys.path: sys.path.insert(0, str(ROOT))
from app import pipeline as P, texts  # noqa: E402

log = logging.getLogger('browser')


def run_inline(cmd, step, cwd=None, env=None, timeout=600):
    """Замена pipeline.run: скрипт конвертера выполняется как из командной строки, но в этом же Python.
    cmd — [python, скрипт, параметры…], как у pipeline.run; возвращает то, что скрипт напечатал."""
    script, args = str(cmd[1]), [str(a) for a in cmd[2:]]
    saved = sys.argv, os.getcwd(), dict(os.environ), list(sys.path)
    out = io.StringIO()
    try:
        if env is not None:
            os.environ.clear(); os.environ.update(env)
        if cwd:
            os.chdir(cwd); sys.path.insert(0, str(cwd))
        sys.argv = [script] + args
        # engine.py читает базу и доноров из окружения при импорте — каждый запуск импортирует его заново
        for m in ('convert', 'engine', 'common'): sys.modules.pop(m, None)
        with contextlib.redirect_stdout(out):
            runpy.run_path(script, run_name='__main__')
    except SystemExit as e:
        if e.code not in (None, 0):
            log.error('%s: код %s\n%s', step, e.code, out.getvalue())
            raise P.ConversionError('Не удалось выполнить шаг «%s».' % step)
    except Exception:
        log.error('%s:\n%s\n%s', step, out.getvalue(), traceback.format_exc())
        raise P.ConversionError('Не удалось выполнить шаг «%s».' % step)
    finally:
        sys.argv = saved[0]; os.chdir(saved[1])
        os.environ.clear(); os.environ.update(saved[2])
        sys.path[:] = saved[3]
    return out.getvalue()


# те самые три отличия от сайта
P.run = run_inline
P.render = lambda pptx, outdir, *a, **k: []
P.ensure_base = lambda template_id: None


def convert(src, outdir, filename, template, lesson='', choices=None, subject='', progress=None):
    """Проверки те же, что у сайта при загрузке файла (app/main.py, create_job), затем pipeline.process.
    Возвращает JSON: {'result': …} или {'error': текст для пользователя, 'trace': подробности для автора}."""
    try:
        name = unicodedata.normalize('NFC', filename or 'presentation.pptx')
        if not name.lower().endswith('.pptx'):
            raise P.ConversionError('Нужен файл PowerPoint с расширением .pptx. Старый формат .ppt сначала пересохраните в .pptx.')
        tpl = P.load_templates().get(template)
        if tpl is None: raise P.ConversionError('Такого шаблона нет.')
        if isinstance(choices, str): choices = json.loads(choices)  # из браузера выбор слайдов приходит строкой JSON
        choices = {str(k): int(v) for k, v in (choices or {}).items()}
        P.check_input(src, tpl)
        if P.uses_lesson(tpl):
            subject = re.sub(r'\s+', ' ', subject or '').strip()[:40] or P.detect_subject(src, texts.SUBJECTS)
            if not subject:
                raise P.ConversionError('Не нашли предмет на титуле старой презентации — впишите его в поле «Предмет».')
        else:
            lesson, subject = '', None
        step = lambda s: progress(s) if progress and s != 'render' else None  # шага «превью» здесь нет
        data = P.process(src, outdir, lesson, template, progress=step, choices=choices, subject=subject)
        return json.dumps({'result': data}, ensure_ascii=False)
    except P.ConversionError as e:
        return json.dumps({'error': str(e)}, ensure_ascii=False)
    except MemoryError:
        return json.dumps({'error': 'Браузеру не хватило памяти на эту презентацию. Закройте лишние вкладки и попробуйте '
                                    'ещё раз или воспользуйтесь сайтом.'}, ensure_ascii=False)
    except Exception:
        return json.dumps({'error': texts.UNEXPECTED.split('.')[0] + '.', 'trace': traceback.format_exc()}, ensure_ascii=False)
