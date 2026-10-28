import sys, re, collections
from pptx import Presentation
from lxml import etree
A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
SKIP_PHR=['Вернуться к навигации','Поставь на паузу','ЭТО ВАЖНО','УВИДИМСЯ','НА СЛЕДУЮЩЕМ УРОКЕ']
def words(path):
    p=Presentation(path); c=collections.Counter(); per={}
    for i,s in enumerate(p.slides,1):
        txt=[]
        for t in s._element.iter(A+'t'):
            txt.append(t.text or '')
        T=' '.join(txt)
        for ph in SKIP_PHR: T=T.replace(ph,' ')
        ws=[w.lower() for w in re.findall(r'[A-Za-zА-Яа-яЁё0-9]+',T)]
        ws=[w for w in ws if not re.fullmatch(r'\d{1,2}',w)]
        per[i]=collections.Counter(ws); c.update(ws)
    return c, per
a,pa=words(sys.argv[1]); b,pb=words(sys.argv[2])
miss=a-b; extra=b-a
print('source words',sum(a.values()),'output',sum(b.values()))
print('MISSING',dict(miss))
print('EXTRA',dict(extra))
