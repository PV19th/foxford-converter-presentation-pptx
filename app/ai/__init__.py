# -*- coding: utf-8 -*-
"""ИИ-доработка результата конвертера (этап 7).

Конвертер по правилам — основа. ИИ правит только раскладку слайдов, которые конвертер пометил
(тесно, сжато, не влезло), и никогда не меняет текст. Два круга: правка, затем проверка своих правок
по новой картинке. Правила для ИИ — в prompt.md, их можно править без программиста."""
import datetime, json, logging, shutil
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from pptx import Presentation
from .. import config, pipeline
from . import claude, slides as S
from .claude import AIError

log = logging.getLogger('ai')
RULES = (Path(__file__).parent / 'prompt.md').read_text('utf8')
FLAGS = ('кегль уменьшен', 'сжато до', 'не помещается', 'нижнего края')  # замечания конвертера, которые отдаём ИИ
ROUNDS = 2
CHECK = ('Это слайд после твоих правок. Проверь по картинке: всё ли влезло, нет ли наложений, соблюдены ли правила. '
         'Если всё хорошо — верни пустой список правок, иначе поправь.')
USAGE = config.DATA / 'ai_usage.json'


def available(): return config.AI_ENABLED


def flagged(result):
    """Слайды, которые пометил конвертер и ещё не доработал ИИ."""
    if result.get('ai'): return []
    return sorted({c['slide'] for c in result['checks'] if c['slide'] and any(f in c['text'] for f in FLAGS)})


def spent_today():  # ₽ за сегодня
    try: return json.loads(USAGE.read_text('utf8')).get(str(datetime.date.today()), 0.0)
    except (OSError, ValueError): return 0.0


def _add_spent(usd):
    try: data = json.loads(USAGE.read_text('utf8'))
    except (OSError, ValueError): data = {}
    day = str(datetime.date.today())
    data[day] = round(data.get(day, 0.0) + usd, 4)
    USAGE.parent.mkdir(parents=True, exist_ok=True)
    USAGE.write_text(json.dumps(data, indent=1), 'utf8')


def improve(job_dir, progress=lambda step: None):
    """Дорабатывает result.pptx в папке задачи; версия без ИИ остаётся в result_rules.pptx."""
    d = Path(job_dir)
    result = json.loads((d / 'result.json').read_text('utf8'))
    todo = flagged(result)
    if not todo: raise AIError('Конвертер не пометил слайдов, которые нужно доработать.')
    if spent_today() >= config.AI_DAILY_LIMIT_RUB:
        raise AIError('Дневной лимит на ИИ исчерпан. Попробуйте завтра или напишите автору сайта.')

    progress('ai_prepare')
    pipeline.render(d / 'source.pptx', d / 'ai_old')  # картинки старых слайдов — ИИ видит, как было
    rules_pptx, work = d / 'result_rules.pptx', d / 'ai.pptx'
    if not rules_pptx.exists(): shutil.copyfile(d / 'result.pptx', rules_pptx)
    shutil.copyfile(rules_pptx, work)
    src_of = {s['slide']: s['src'] for s in result['slides']}
    notes = {n: [c['text'] for c in result['checks'] if c['slide'] == n] for n in todo}
    images = {n: d / 'slides' / ('%03d.jpg' % n) for n in todo}
    comments, touched, cost, rounds = {}, set(), 0.0, 0

    for rounds in range(1, ROUNDS + 1):
        progress('ai_fix' if rounds == 1 else 'ai_check')
        prs = Presentation(work)
        W, H = prs.slide_width, prs.slide_height
        zone = [round(0.95 * H / 6858000, 2), round(6.80 * H / 6858000, 2)]  # рабочая зона, как у конвертера

        def ask(n):
            info = {'slide_size_in': [round(W / S.EMU, 3), round(H / S.EMU, 3)], 'content_zone_y_in': zone,
                    'converter_notes': notes[n], 'shapes': S.describe(prs.slides[n - 1])}
            old = d / 'ai_old' / 'slides' / ('%03d.jpg' % src_of[n]) if src_of.get(n) else None
            try:
                return n, claude.ask(RULES, images[n], info, old if old and old.exists() else None,
                                     CHECK if rounds > 1 else None)
            except AIError: raise
            except Exception:
                log.exception('ИИ: слайд %d не обработан', n)
                return n, ({'comment': '', 'edits': []}, 0.0)

        with ThreadPoolExecutor(4) as ex: answers = list(ex.map(ask, todo))
        changed = []
        for n, (answer, usd) in answers:
            cost += usd
            if S.apply(prs.slides[n - 1], answer['edits'], (W, H)):
                changed.append(n); touched.add(n)
                if answer['comment'] and n not in comments: comments[n] = answer['comment']
                log.info('ИИ: слайд %d, круг %d — %s', n, rounds, answer['comment'] or 'без комментария')
        prs.save(work)
        log.info('ИИ: круг %d, слайдов %d, изменено %d, %.0f ₽', rounds, len(todo), len(changed), cost)
        if not changed or rounds == ROUNDS: break
        pipeline.render(work, d / 'ai_new')  # новые картинки — для проверки своих правок
        todo, images = changed, {n: d / 'ai_new' / 'slides' / ('%03d.jpg' % n) for n in changed}

    _add_spent(cost)
    # текст правками не задевается — но проверяем, иначе отдаём версию без ИИ
    if [S.slide_text(s) for s in Presentation(work).slides] != [S.slide_text(s) for s in Presentation(rules_pptx).slides]:
        raise AIError('ИИ задел текст слайдов — его правки отменены, файл остался как после конвертера.')
    progress('ai_render')
    shutil.copyfile(work, d / 'result.pptx')
    result['previews'] = pipeline.render(d / 'result.pptx', d)
    done = sorted(touched)
    result['checks'] = [c for c in result['checks']
                        if not (c['slide'] in done and any(f in c['text'] for f in FLAGS))]
    result['ai'] = {'model': claude.MODEL, 'rounds': rounds, 'cost_rub': round(cost),
                    'slides': [{'slide': n, 'comment': comments.get(n, 'Раскладка поправлена')} for n in done]}
    result['size_mb'] = round((d / 'result.pptx').stat().st_size / 1e6, 1)
    (d / 'result.json').write_text(json.dumps(result, ensure_ascii=False, indent=1), 'utf8')
    for name in ('ai_old', 'ai_new'): shutil.rmtree(d / name, ignore_errors=True)
    work.unlink(missing_ok=True)
    return result
