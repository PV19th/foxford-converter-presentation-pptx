import re, copy
from lxml import etree
from pptx import Presentation
NS={'a':'http://schemas.openxmlformats.org/drawingml/2006/main','p':'http://schemas.openxmlformats.org/presentationml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
A='{%s}'%NS['a']; P='{%s}'%NS['p']; R='{%s}'%NS['r']
def theme_of(slide):
    m=slide.slide_layout.slide_master
    for rel in m.part.rels.values():
        if rel.reltype.endswith('/theme'):
            return etree.fromstring(rel.target_part.blob)
def scheme_map(slide):
    th=theme_of(slide); d={}
    cs=th.find('.//a:clrScheme',NS)
    for c in cs:
        tag=etree.QName(c).localname
        v=c.find('a:srgbClr',NS)
        if v is None: v=c.find('a:sysClr',NS); val=v.get('lastClr') if v is not None else '000000'
        else: val=v.get('val')
        d[tag]=val
    m=slide.slide_layout.slide_master._element.find('p:clrMap',NS)
    for k,v in m.attrib.items(): d[k]=d.get(v,'000000')
    return d
def apply_mods(hexv, el):
    r,g,b=[int(hexv[i:i+2],16) for i in (0,2,4)]
    for m in el:
        t=etree.QName(m).localname; v=int(m.get('val','100000'))/100000
        if t=='lumMod': r,g,b=r*v,g*v,b*v
        elif t=='lumOff': r,g,b=r+255*v,g+255*v,b+255*v
        elif t=='tint': r,g,b=[c+(255-c)*(1-v) for c in (r,g,b)]
        elif t=='shade': r,g,b=r*v,g*v,b*v
    cl=lambda x:max(0,min(255,int(round(x))))
    return '%02X%02X%02X'%(cl(r),cl(g),cl(b))
def color_of(el, smap):
    """el: element containing srgbClr/schemeClr child (e.g. solidFill). returns hex or None"""
    if el is None: return None
    for c in el:
        t=etree.QName(c).localname
        if t=='srgbClr': return apply_mods(c.get('val'),c)
        if t=='schemeClr': return apply_mods(smap.get(c.get('val'),'000000'),c)
        if t=='prstClr': return {'black':'000000','white':'FFFFFF'}.get(c.get('val'),'000000')
    return None
