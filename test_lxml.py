# -*- coding: utf-8 -*-
from lxml import etree
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def qn(tag):
    return '{%s}%s' % (W, tag)
print(repr(qn('w:r')))
r = etree.Element(qn('w:r'))
print('OK:', etree.QName(r).localname)
r2 = etree.Element(etree.QName(W, 'r'))
print('OK2:', etree.QName(r2).localname)
