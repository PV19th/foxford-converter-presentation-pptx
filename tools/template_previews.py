# -*- coding: utf-8 -*-
"""Миниатюры фирменного шаблона: листы для docs/template/ и картинки выбора слайдов для сайта.

Запускать после замены шаблона в base/ (нужен LibreOffice, поэтому в Docker):
  docker compose run --rm -v "$PWD:/srv" app python3 tools/template_previews.py [папка_профиля]

- docs/template/миниатюры_001-050.jpg … — все слайды шаблона с номерами (5 в ряд);
- templates/<профиль>/thumbs/<N>.jpg — слайды из раздела "choices" профиля.
Старые картинки удаляются.
"""
import json, shutil, subprocess, sys, tempfile
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
FONT = '/usr/share/fonts/truetype/open-sans/OpenSans-Regular.ttf'


def render(pptx, tmp):
    """Слайды → JPEG 400×300 (40 точек на дюйм для слайда 10×7,5″)."""
    shutil.copy(pptx, tmp / 't.pptx')
    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', str(tmp), str(tmp / 't.pptx')],
                   check=True, capture_output=True)
    subprocess.run(['pdftoppm', '-r', '40', '-jpeg', str(tmp / 't.pdf'), str(tmp / 'p')], check=True)
    return sorted(tmp.glob('p-*.jpg'))


def sheets(pages, out):
    for f in out.glob('миниатюры_*.jpg'): f.unlink()
    font = ImageFont.truetype(FONT, 11)
    for start in range(0, len(pages), 50):
        chunk = pages[start:start + 50]
        sheet = Image.new('RGB', (2000, 316 * ((len(chunk) + 4) // 5)), 'white')
        d = ImageDraw.Draw(sheet)
        for i, p in enumerate(chunk):
            x, y = (i % 5) * 400, (i // 5) * 316
            d.text((x + 2, y + 1), '#%d' % (start + i + 1), fill='black', font=font)
            sheet.paste(Image.open(p).resize((400, 300)), (x, y + 16))
        sheet.save(out / ('миниатюры_%03d-%03d.jpg' % (start + 1, start + len(chunk))), quality=85)


def thumbs(pages, profile_dir):
    profile = json.loads((profile_dir / 'profile.json').read_text('utf8'))
    out = profile_dir / 'thumbs'
    out.mkdir(exist_ok=True)
    for f in out.glob('*.jpg'): f.unlink()
    nums = sorted({n for c in profile.get('choices', {}).values() for n in c['slides']})
    for n in nums:
        Image.open(pages[n - 1]).resize((320, 240), Image.LANCZOS).save(out / ('%d.jpg' % n), quality=88)
    return nums


if __name__ == '__main__':
    profile_dir = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else ROOT / 'templates' / 'foxford_5-11_v01_26'
    template = ROOT / json.loads((profile_dir / 'profile.json').read_text('utf8'))['base_recipe']['template']
    with tempfile.TemporaryDirectory() as tmp:
        pages = render(template, Path(tmp))
        sheets(pages, ROOT / 'docs' / 'template')
        nums = thumbs(pages, profile_dir)
    print('слайдов в шаблоне:', len(pages), '| миниатюр выбора:', len(nums))
