# -*- coding: utf-8 -*-
"""ИИ-доработка из командной строки (в Docker): python3 -m app.ai outputs/tests/<профиль>/<урок>"""
import logging, sys
from . import improve, AIError

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
try:
    r = improve(sys.argv[1], progress=lambda s: logging.info('шаг: %s', s))
except AIError as e:
    print('ОШИБКА:', e, file=sys.stderr); sys.exit(1)
print('ИИ поправил слайдов: %d | кругов: %d | стоимость: %d ₽' % (len(r['ai']['slides']), r['ai']['rounds'], r['ai']['cost_rub']))
for s in r['ai']['slides']: print('  слайд %d — %s' % (s['slide'], s['comment']))
