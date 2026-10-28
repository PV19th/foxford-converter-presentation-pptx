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
