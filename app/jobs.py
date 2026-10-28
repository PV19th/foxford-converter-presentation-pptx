# -*- coding: utf-8 -*-
"""Очередь задач. Обработка строго по одной: soffice нельзя запускать параллельно.

Каждая задача — папка data/jobs/<id>/ с job.json; результаты переживают перезапуск сайта.
Старые папки удаляются через JOB_TTL_HOURS часов."""
import json, logging, queue, shutil, threading, time, uuid
from . import config, pipeline, texts

log = logging.getLogger('jobs')
_jobs = {}        # id → состояние задачи (то же, что в job.json)
_pending = []     # id задач в очереди, по порядку
_lock = threading.Lock()
_queue = queue.Queue()


def _dir(jid): return config.JOBS / jid


def _save(job):
    (_dir(job['id']) / 'job.json').write_text(json.dumps(job, ensure_ascii=False), 'utf8')


def create(filename, template, lesson):
    """Новая задача; файл презентации вызывающий кладёт в source_path(id) и затем вызывает submit."""
    jid = uuid.uuid4().hex
    _dir(jid).mkdir(parents=True)
    job = {'id': jid, 'filename': filename, 'template': template, 'lesson': lesson,
           'status': 'queued', 'step': None, 'error': None, 'created': time.time()}
    _save(job)
    return job


def source_path(jid): return _dir(jid) / 'source.pptx'


def result_path(jid): return _dir(jid) / 'result.pptx'


def submit(job):
    with _lock:
        _jobs[job['id']] = job
        _pending.append(job['id'])
    _queue.put(job['id'])
    log.info('задача %s: в очереди (%s, %s, %s)', job['id'], job['filename'], job['template'], job['lesson'])


def discard(jid): shutil.rmtree(_dir(jid), ignore_errors=True)


def get(jid):
    """Состояние задачи для страницы или None."""
    job = _jobs.get(jid)
    if job is None: return None
    out = dict(job)
    if job['status'] == 'queued':
        with _lock:
            busy = any(j['status'] == 'running' for j in _jobs.values())
            out['ahead'] = (_pending.index(jid) if jid in _pending else 0) + (1 if busy else 0)
    if job['status'] == 'done':
        out['result'] = json.loads((_dir(jid) / 'result.json').read_text('utf8'))
    return out


def _worker():
    while True:
        jid = _queue.get()
        with _lock:
            if jid in _pending: _pending.remove(jid)
        job = _jobs.get(jid)
        if job is None: continue
        job.update(status='running', step=texts.STEPS[0][0]); _save(job)
        log.info('задача %s: начата', jid)

        def progress(step):
            job['step'] = step; _save(job)
        try:
            pipeline.process(source_path(jid), _dir(jid), job['lesson'], job['template'], progress)
            job.update(status='done', step=None)
            for name in ('source.pptx', 'converted.pptx'): (_dir(jid) / name).unlink(missing_ok=True)
        except pipeline.ConversionError as e:
            job.update(status='error', error=str(e))
        except Exception:
            log.exception('задача %s: непредвиденная ошибка', jid)
            job.update(status='error', error=texts.UNEXPECTED)
        job['finished'] = time.time()
        _save(job)
        log.info('задача %s: %s', jid, job['status'])


def _cleanup():
    while True:
        limit = time.time() - config.JOB_TTL_HOURS * 3600
        for d in config.JOBS.iterdir():
            if d.stat().st_mtime < limit and (_jobs.get(d.name, {}).get('status') not in ('queued', 'running')):
                _jobs.pop(d.name, None)
                shutil.rmtree(d, ignore_errors=True)
                log.info('задача %s: удалена по сроку хранения', d.name)
        time.sleep(15 * 60)


def start():
    """Восстановить задачи с диска и запустить исполнителя и уборку."""
    config.JOBS.mkdir(parents=True, exist_ok=True)
    for f in config.JOBS.glob('*/job.json'):
        try: job = json.loads(f.read_text('utf8'))
        except ValueError: continue
        if job['status'] in ('queued', 'running'):
            job.update(status='error', error=texts.INTERRUPTED); _save(job)
        _jobs[job['id']] = job
    threading.Thread(target=_worker, daemon=True, name='worker').start()
    threading.Thread(target=_cleanup, daemon=True, name='cleanup').start()
