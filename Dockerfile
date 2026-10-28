FROM ubuntu:24.04

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    LANG=C.UTF-8

# LibreOffice — превью слайдов; Open Sans — расчёт переносов в engine.py; poppler — PDF → картинки
RUN apt-get update && apt-get install -y --no-install-recommends \
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

COPY converter/requirements.txt /tmp/requirements.txt
COPY constraints.txt /tmp/constraints.txt
RUN pip install --no-cache-dir -r /tmp/requirements.txt -c /tmp/constraints.txt

WORKDIR /srv
COPY base/ base/
COPY templates/ templates/
COPY converter/ converter/
COPY app/ app/

# конвертер ищет assets/ со стикерами в CONV_DIR
ENV CONV_DIR=/srv
