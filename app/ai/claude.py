# -*- coding: utf-8 -*-
"""Провайдер ИИ — Anthropic (Claude). Ключ — ANTHROPIC_API_KEY из .env.

Один запрос — один слайд: картинки старого и нового слайда, список фигур и замечания конвертера.
Ответ — строго по схеме SCHEMA: список правок раскладки и короткий комментарий для методиста."""
import base64, json
import anthropic

MODEL = 'claude-opus-5-5'
EFFORT = 'high'  # раскладка — задача на сообразительность; medium дешевле, но хуже
PRICE_IN, PRICE_OUT, PRICE_CACHE_READ = 4.0, 20.0, 0.2  # $ за 1 млн токенов (Claude Opus 5.5)

_num = {'anyOf': [{'type': 'number'}, {'type': 'null'}]}
_enum = lambda *v: {'anyOf': [{'type': 'string', 'enum': list(v)}, {'type': 'null'}]}
SCHEMA = {
    'type': 'object', 'additionalProperties': False, 'required': ['comment', 'edits'],
    'properties': {
        'comment': {'type': 'string'},
        'edits': {'type': 'array', 'items': {
            'type': 'object', 'additionalProperties': False,
            'required': ['id', 'x', 'y', 'w', 'h', 'font_pt', 'align', 'fill', 'z', 'delete'],
            'properties': {
                'id': {'type': 'integer'}, 'x': _num, 'y': _num, 'w': _num, 'h': _num, 'font_pt': _num,
                'align': _enum('left', 'center', 'right'),
                'fill': _enum('641FF1', 'EEE5FF', 'FF9900', 'FFEED7', '71D2FF', 'CCEAFF', 'FFFA7D', 'FFFED6', 'FFFFFF', 'none'),
                'z': _enum('front', 'back'),
                'delete': {'type': 'boolean'}}}}}}


class AIError(Exception):
    """Ошибка, которую можно показать пользователю как есть."""


def _image(path):
    return {'type': 'image', 'source': {'type': 'base64', 'media_type': 'image/jpeg',
                                        'data': base64.standard_b64encode(path.read_bytes()).decode()}}


def ask(rules, new_image, info, old_image=None, note=None):
    """→ ({'comment', 'edits'}, стоимость в $). info — словарь с фигурами, рабочей зоной и замечаниями."""
    content = []
    if old_image: content += [{'type': 'text', 'text': 'Старый слайд (как было у автора урока):'}, _image(old_image)]
    content += [{'type': 'text', 'text': 'Новый слайд сейчас:'}, _image(new_image),
                {'type': 'text', 'text': (note + '\n\n' if note else '') + json.dumps(info, ensure_ascii=False)}]
    try:
        r = anthropic.Anthropic().beta.messages.create(
            model=MODEL, max_tokens=16000,
            betas=['server-side-fallback-2026-07-01'], fallbacks='default',  # отказ модели — повтор на другой
            system=[{'type': 'text', 'text': rules, 'cache_control': {'type': 'ephemeral'}}],
            output_config={'effort': EFFORT, 'format': {'type': 'json_schema', 'schema': SCHEMA}},
            messages=[{'role': 'user', 'content': content}])
    except anthropic.AuthenticationError:
        raise AIError('Ключ ИИ не подошёл. Проверьте ANTHROPIC_API_KEY в настройках сайта.')
    except anthropic.PermissionDeniedError:
        raise AIError('Нет доступа к ИИ: возможно, API Anthropic недоступен из этой страны.')
    except anthropic.RateLimitError:
        raise AIError('ИИ сейчас перегружен. Попробуйте через пару минут.')
    except anthropic.APIConnectionError:
        raise AIError('Нет связи с ИИ. Проверьте интернет или VPN.')
    u = r.usage
    cost = (u.input_tokens * PRICE_IN + (u.cache_creation_input_tokens or 0) * PRICE_IN * 1.25
            + (u.cache_read_input_tokens or 0) * PRICE_CACHE_READ + u.output_tokens * PRICE_OUT) / 1e6
    if r.stop_reason != 'end_turn':  # отказ или не хватило места на ответ — слайд оставляем как есть
        return {'comment': '', 'edits': []}, cost
    return json.loads(next(b.text for b in r.content if b.type == 'text')), cost
