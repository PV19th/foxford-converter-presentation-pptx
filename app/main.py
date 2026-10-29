# -*- coding: utf-8 -*-
"""Сайт: вход по общему паролю, загрузка презентации, очередь, результат.

Запуск: uvicorn app.main:app --host 0.0.0.0 --port 8000 --workers 1
(только один процесс — очередь живёт в памяти)."""
import hashlib, hmac, json, logging, re, time, unicodedata
from contextlib import asynccontextmanager
from logging.handlers import RotatingFileHandler
from pathlib import Path
from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from . import config, jobs, pipeline, texts

WEB = Path(__file__).resolve().parent / 'web'
COOKIE = 'fx_session'
SLIDE_RE = re.compile(r'^\d{3}\.jpg$')

config.LOGS.mkdir(parents=True, exist_ok=True)
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(name)s: %(message)s',
                    handlers=[logging.StreamHandler(),
                              RotatingFileHandler(config.LOGS / 'app.log', maxBytes=5_000_000, backupCount=5)])
log = logging.getLogger('site')



@asynccontextmanager
async def lifespan(app):
    if not config.APP_PASSWORD: log.error('APP_PASSWORD не задан — войти на сайт будет нельзя')
    jobs.start()
    yield


app = FastAPI(lifespan=lifespan, docs_url=None, redoc_url=None, openapi_url=None)
app.mount('/static', StaticFiles(directory=WEB / 'static'), name='static')


# ---------------------------------------------------------------- вход
def _token():
    return hmac.new(config.SECRET_KEY.encode(), config.APP_PASSWORD.encode(), hashlib.sha256).hexdigest()


def _authed(request):
    return bool(config.APP_PASSWORD) and hmac.compare_digest(request.cookies.get(COOKIE, ''), _token())


_fails = {}  # ip → время неудачных попыток входа


@app.middleware('http')
async def _auth(request, call_next):
    path = request.url.path
    if path == '/login' or path.startswith('/static/') or _authed(request):
        return await call_next(request)
    if path.startswith('/api/'): return JSONResponse({'error': 'Нужно войти заново.'}, 401)
    return RedirectResponse('/login', 303)


@app.get('/login')
def login_page(): return FileResponse(WEB / 'login.html')


@app.post('/login')
def login(request: Request, password: str = Form('')):
    ip = request.client.host if request.client else '?'
    recent = [t for t in _fails.get(ip, []) if t > time.time() - 600]
    if len(recent) >= 10:
        return RedirectResponse('/login?e=wait', 303)
    if not config.APP_PASSWORD or not hmac.compare_digest(password, config.APP_PASSWORD):
        _fails[ip] = recent + [time.time()]
        log.warning('неверный пароль с %s', ip)
        return RedirectResponse('/login?e=1', 303)
    _fails.pop(ip, None)
    r = RedirectResponse('/', 303)
    r.set_cookie(COOKIE, _token(), max_age=30 * 24 * 3600, httponly=True, samesite='lax',
                 secure=request.url.scheme == 'https')
    return r


@app.post('/logout')
def logout():
    r = RedirectResponse('/login', 303); r.delete_cookie(COOKIE); return r


# ---------------------------------------------------------------- страница и данные
@app.get('/')
def index(): return FileResponse(WEB / 'index.html')


@app.get('/api/meta')
def meta():
    tpls = [{'id': k, 'name': v['name'], 'grades': pipeline.profile_grades(v),
             'choices': [dict(c, key=key) for key, c in v.get('choices', {}).items()]}
            for k, v in sorted(pipeline.load_templates().items(), key=lambda kv: kv[1].get('order', 99))]
    return {'templates': tpls, 'steps': texts.STEPS, 'manual_checks': texts.MANUAL_CHECKS, 'subjects': texts.SUBJECTS,
            'contact': config.SUPPORT_CONTACT, 'max_upload_mb': config.MAX_UPLOAD_MB,
            'ttl_hours': config.JOB_TTL_HOURS}


def _bad(msg): return JSONResponse({'error': msg}, 400)


@app.post('/api/jobs')
async def create_job(file: UploadFile = File(...), template: str = Form(...), lesson: str = Form(...),
                     choices: str = Form('{}'), subject: str = Form('')):
    # macOS присылает «й» двумя символами — приводим, чтобы имя скачанного файла было обычным
    name = unicodedata.normalize('NFC', Path(file.filename or 'presentation.pptx').name)
    if not name.lower().endswith('.pptx'):
        return _bad('Нужен файл PowerPoint с расширением .pptx. Старый формат .ppt сначала пересохраните в .pptx.')
    if not pipeline.LESSON_RE.match(lesson):
        return _bad('Номер урока нужно писать так: класс.урок, например 5.21.')
    tpl = pipeline.load_templates().get(template)
    if tpl is None: return _bad('Такого шаблона нет.')
    if pipeline.grade_settings(tpl, lesson.split('.')[0]) is None:
        return _bad('Шаблон «%s» пока не поддерживает этот класс.' % tpl['name'])
    try:
        choices = {str(k): int(v) for k, v in json.loads(choices).items()}
        pipeline.check_choices(tpl, choices)
    except (ValueError, AttributeError, pipeline.ConversionError):
        return _bad('Такого варианта слайда нет. Обновите страницу.')
    subject = re.sub(r'\s+', ' ', subject).strip()[:40]
    job = jobs.create(name, template, lesson, choices)
    limit, size = config.MAX_UPLOAD_MB * 1024 * 1024, 0
    with open(jobs.source_path(job['id']), 'wb') as f:
        while chunk := await file.read(1024 * 1024):
            size += len(chunk)
            if size > limit:
                f.close(); jobs.discard(job['id'])
                return _bad('Файл больше %d МБ. Сожмите картинки или напишите автору.' % config.MAX_UPLOAD_MB)
            f.write(chunk)
    try:
        pipeline.check_input(jobs.source_path(job['id']))
    except pipeline.ConversionError as e:
        jobs.discard(job['id']); return _bad(str(e))
    job['subject'] = subject or pipeline.detect_subject(jobs.source_path(job['id']), texts.SUBJECTS)
    if not job['subject']:
        jobs.discard(job['id'])
        return _bad('Не нашли предмет на титуле старой презентации — впишите его в поле «Предмет».')
    jobs.submit(job)
    return {'id': job['id']}


@app.get('/api/templates/{tid}/thumbs/{n}.jpg')
def template_thumb(tid: str, n: int):
    p = pipeline.TEMPLATES / tid / 'thumbs' / ('%d.jpg' % n)
    if tid not in pipeline.load_templates() or not p.exists(): return JSONResponse({'error': 'Нет такой картинки.'}, 404)
    return FileResponse(p, media_type='image/jpeg', headers={'Cache-Control': 'private, max-age=86400'})


def _find(jid):
    return jobs.get(jid) if re.fullmatch(r'[0-9a-f]{32}', jid) else None


@app.get('/api/jobs/{jid}')
def job_status(jid: str):
    job = _find(jid)
    if job is None: return JSONResponse({'error': 'Задача не найдена — возможно, её уже удалили по сроку хранения.'}, 404)
    return job


@app.get('/api/jobs/{jid}/download')
def download(jid: str):
    job = _find(jid)
    if not job or job['status'] != 'done': return JSONResponse({'error': 'Файл не готов.'}, 404)
    return FileResponse(jobs.result_path(jid), filename=Path(job['filename']).stem + '_новый.pptx',
                        media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation')


@app.get('/api/jobs/{jid}/{kind}/{name}')
def slide_image(jid: str, kind: str, name: str):
    job = _find(jid)
    if not job or kind not in ('slides', 'thumbs') or not SLIDE_RE.match(name):
        return JSONResponse({'error': 'Нет такого слайда.'}, 404)
    p = config.JOBS / jid / kind / name
    if not p.exists(): return JSONResponse({'error': 'Нет такого слайда.'}, 404)
    return FileResponse(p, media_type='image/jpeg', headers={'Cache-Control': 'private, max-age=86400'})
