"""Сжатие шаблона: непрозрачные PNG -> JPEG (q93, без цветовой субдискретизации),
прозрачные PNG -> пережатие без потерь, если выходит меньше. Ссылки в .rels обновляются."""
import sys, zipfile, io, math
from PIL import Image
src, dst = sys.argv[1], sys.argv[2]
z = zipfile.ZipFile(src)
files = {n: z.read(n) for n in z.namelist()}
order = z.namelist()
ren, psnrs, saved = {}, [], 0
for n in [n for n in order if n.startswith('ppt/media/') and n.lower().endswith('.png')]:
    raw = files[n]; im = Image.open(io.BytesIO(raw)); im.load()
    rgba = im.convert('RGBA'); opaque = rgba.getchannel('A').getextrema()[0] >= 250
    if opaque:
        buf = io.BytesIO(); rgb = rgba.convert('RGB')
        rgb.save(buf, 'JPEG', quality=93, subsampling=0, optimize=True)
        if buf.tell() < len(raw) * 0.8:
            back = Image.open(io.BytesIO(buf.getvalue())).convert('RGB')
            # контроль качества: PSNR по уменьшенной копии
            a = rgb.resize((400, int(400 * rgb.height / rgb.width))); b = back.resize(a.size)
            mse = sum((x - y) ** 2 for pa, pb in zip(zip(*[iter(a.tobytes())]*3), zip(*[iter(b.tobytes())]*3)) for x, y in zip(pa, pb)) / (a.width * a.height * 3)
            psnrs.append(99 if mse == 0 else 10 * math.log10(255 ** 2 / mse))
            nn = n[:-4] + '_c.jpg'; saved += len(raw) - buf.tell()
            del files[n]; files[nn] = buf.getvalue(); ren[n.split('/')[-1]] = nn.split('/')[-1]
            order[order.index(n)] = nn
            continue
    buf = io.BytesIO(); im.save(buf, 'PNG', optimize=True)
    if buf.tell() < len(raw): saved += len(raw) - buf.tell(); files[n] = buf.getvalue()
for n in files:
    if n.endswith('.rels'):
        t = files[n].decode('utf8')
        for a, b in ren.items(): t = t.replace('media/' + a + '"', 'media/' + b + '"')
        files[n] = t.encode('utf8')
ct = files['[Content_Types].xml'].decode('utf8')
if 'Extension="jpg"' not in ct:
    ct = ct.replace('<Default ', '<Default Extension="jpg" ContentType="image/jpeg"/><Default ', 1)
files['[Content_Types].xml'] = ct.encode('utf8')
with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
    for n in order: o.writestr(n, files[n])
print('PNG->JPEG:', len(ren), '| сэкономлено МБ:', round(saved / 1e6, 1),
      '| PSNR мин/сред:', round(min(psnrs), 1), round(sum(psnrs) / len(psnrs), 1))
