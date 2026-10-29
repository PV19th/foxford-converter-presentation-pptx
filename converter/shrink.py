import sys, zipfile, io, re
from PIL import Image
src, dst, limit = sys.argv[1], sys.argv[2], float(sys.argv[3])*1e6
z = zipfile.ZipFile(src)
files = {n: z.read(n) for n in z.namelist()}
total = sum(len(v) for v in files.values())
media = sorted([n for n in files if n.startswith('ppt/media/') and n.endswith('.png')], key=lambda n: -len(files[n]))
ren = {}
for n in media:
    if total < limit: break
    im = Image.open(io.BytesIO(files[n]))
    if im.mode in ('RGBA', 'LA', 'P'):
        a = im.convert('RGBA').getchannel('A')
        if a.getextrema()[0] < 250: continue
    buf = io.BytesIO(); im.convert('RGB').save(buf, 'JPEG', quality=92, optimize=True)
    if buf.tell() > len(files[n]) * 0.8: continue
    nn = n[:-4] + '_c.jpg'
    total += buf.tell() - len(files[n])
    del files[n]; files[nn] = buf.getvalue(); ren[n.split('/')[-1]] = nn.split('/')[-1]
# всё ещё тяжело (старый шаблон 10×7,5″: сотни фото и стикеров) — уменьшаем картинки до размера,
# в котором их видно на слайде: сначала 200 точек на дюйм, потом 150, потом 110
if total >= limit:
    import posixpath
    from lxml import etree
    A_, P_ = '{http://schemas.openxmlformats.org/drawingml/2006/main}', '{http://schemas.openxmlformats.org/presentationml/2006/main}'
    R_ = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    shown = {}  # картинка → наибольшая ширина и высота на слайде, дюймы (с учётом обрезки)
    for n in [n for n in files if re.fullmatch(r'ppt/(slides|slideLayouts|slideMasters)/\w+\.xml', n)]:
        d, f = n.rsplit('/', 1)
        rels = etree.fromstring(files.get(d + '/_rels/' + f + '.rels', b'<r/>'))
        target = {r.get('Id'): posixpath.normpath(posixpath.join(d, r.get('Target', ''))) for r in rels}
        for pic in etree.fromstring(files[n]).iter(P_ + 'pic'):
            blip, ext = pic.find('.//' + A_ + 'blip'), pic.find('.//' + A_ + 'xfrm/' + A_ + 'ext')
            if blip is None or ext is None or target.get(blip.get(R_ + 'embed')) not in files: continue
            sr = pic.find('.//' + A_ + 'srcRect')
            cut = lambda a, b: 1 - (int(sr.get(a, 0)) + int(sr.get(b, 0))) / 100000 if sr is not None else 1
            w, h = int(ext.get('cx')) / 914400 / max(cut('l', 'r'), 0.05), int(ext.get('cy')) / 914400 / max(cut('t', 'b'), 0.05)
            m = target[blip.get(R_ + 'embed')]
            shown[m] = (max(shown.get(m, (0, 0))[0], w), max(shown.get(m, (0, 0))[1], h))
    for dpi in (200, 150, 110):
        for n in sorted(shown, key=lambda n: -len(files[n])):
            if total < limit: break
            if not n.lower().endswith(('.jpg', '.jpeg', '.png')): continue
            im = Image.open(io.BytesIO(files[n]))
            k = max(shown[n][0] * dpi / im.size[0], shown[n][1] * dpi / im.size[1])
            if k >= 0.9: continue  # и так не крупнее нужного
            im = im.resize((max(1, round(im.size[0] * k)), max(1, round(im.size[1] * k))), Image.LANCZOS)
            buf = io.BytesIO()
            if n.lower().endswith('.png'): im.save(buf, 'PNG', optimize=True)
            else: im.convert('RGB').save(buf, 'JPEG', quality=88, optimize=True)
            if buf.tell() >= len(files[n]): continue
            total += buf.tell() - len(files[n]); files[n] = buf.getvalue()
    # не хватило — пережимаем JPEG (качество 82) и уменьшаем самые крупные картинки до 1280 px
    for n in sorted([n for n in files if n.startswith('ppt/media/') and n.lower().endswith(('.jpg', '.jpeg', '.png'))],
                    key=lambda n: -len(files[n])):
        if total < limit or len(files[n]) < 100000: break
        im = Image.open(io.BytesIO(files[n]))
        jpg = not n.lower().endswith('.png')
        if max(im.size) > 1280: im.thumbnail((1280, 1280), Image.LANCZOS)
        elif not jpg: continue
        buf = io.BytesIO()
        if jpg: im.convert('RGB').save(buf, 'JPEG', quality=82, optimize=True)
        else: im.save(buf, 'PNG', optimize=True)
        if buf.tell() >= len(files[n]) * 0.95: continue
        total += buf.tell() - len(files[n]); files[n] = buf.getvalue()
for n in list(files):
    if n.endswith('.rels'):
        t = files[n].decode('utf8')
        for a, b in ren.items(): t = t.replace('media/' + a + '"', 'media/' + b + '"')
        files[n] = t.encode('utf8')
ct = files['[Content_Types].xml'].decode('utf8')
if 'Extension="jpg"' not in ct:
    ct = ct.replace('<Default ', '<Default Extension="jpg" ContentType="image/jpeg"/><Default ', 1)
files['[Content_Types].xml'] = ct.encode('utf8')
with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
    o.writestr('[Content_Types].xml', files.pop('[Content_Types].xml'))
    for n, v in files.items(): o.writestr(n, v)
print(len(ren), 'converted; size', round(total/1e6,1))
