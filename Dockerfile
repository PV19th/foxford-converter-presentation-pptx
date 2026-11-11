FROM ubuntu:24.04

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    LANG=C.UTF-8

# LibreOffice — превью слайдов; Open Sans — расчёт переносов в engine.py; poppler — PDF → картинки.
# Pipeline-Depth=0: пакеты без конвейера HTTP — иначе в некоторых сетях сборка падает с «Hash Sum mismatch»
RUN printf 'Acquire::http::Pipeline-Depth "0";\nAcquire::Retries "3";\n' > /etc/apt/apt.conf.d/90network \
 && apt-get update && apt-get install -y --no-install-recommends \
        python3 python3-venv \
        libreoffice-impress \
        fonts-open-sans fonts-liberation2 \
        poppler-utils \
    && rm -rf /var/lib/apt/lists/*

# engine.py берёт шрифт по жёсткому пути — без него сборка должна упасть сразу
RUN test -f /usr/share/fonts/truetype/open-sans/OpenSans-Regular.ttf \
 && test -f /usr/share/fonts/truetype/open-sans/OpenSans-Bold.ttf

RUN python3 -m venv /venv
ENV PATH=/venv/bin:$PATH

COPY converter/requirements.txt /tmp/converter-requirements.txt
COPY app/requirements.txt /tmp/app-requirements.txt
COPY constraints.txt /tmp/constraints.txt
RUN pip install --no-cache-dir -r /tmp/converter-requirements.txt -r /tmp/app-requirements.txt -c /tmp/constraints.txt

WORKDIR /srv
COPY base/ base/
COPY templates/ templates/
COPY converter/ converter/
COPY tools/ tools/
COPY app/ app/

# базы профилей на фирменном шаблоне собираются скриптом (base/generated/)
RUN python3 tools/build_base.py

# конвертер ищет assets/ со стикерами в CONV_DIR
ENV CONV_DIR=/srv

EXPOSE 8000
# один процесс: очередь задач живёт в памяти, soffice — строго по одной задаче
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "1"]
