"""Repair recognized heading styles, preserving fields, bookmarks and run contents."""
import argparse,re
from docx import Document
from docx.text.paragraph import Paragraph
from engine import configure_styles,numbering,fonts,HEADINGS,ROLES,qn,el,canonicalize
def level_for(name,mapping):
 if name in mapping:return int(mapping[name])
 aliases={'Title':None,'标题 1':1,'标题 2':2,'标题 3':3,'一级标题':2,'二级标题_改':3,'章标题':1,'节标题':2,'小节标题':3}
 if name in HEADINGS:return HEADINGS.index(name)+1
 return aliases.get(name)
def strip_prefix(p):
 nodes=p._p.xpath('./w:r/w:t');text=''.join(n.text or '' for n in nodes)
 m=re.match(r'^\s*\d+(?:\.\d+){0,2}(?:[.、]?\s+|[、])',text)
 if not m:return
 count=m.end()
 for n in nodes:
  t=n.text or '';n.text=t[min(count,len(t)):];count=max(0,count-len(t))
  if not count:break
def repair(doc,mapping=None,strip=False):
 mapping=mapping or {};configure_styles(doc);nid=numbering(doc);changed=[]
 for i,node in enumerate(doc.element.body.xpath('.//w:p')):
  p=Paragraph(node,doc);level=level_for(p.style.name,mapping)
  if not level:continue
  p.style=HEADINGS[level-1]
  pp=p._p.get_or_add_pPr()
  # Remove conflicting direct paragraph formatting but keep section breaks and bookmarks.
  for c in list(pp):
   if c.tag in [qn('w:numPr'),qn('w:jc'),qn('w:ind'),qn('w:spacing'),qn('w:outlineLvl')]:pp.remove(c)
  np=el('w:numPr');np.extend([el('w:ilvl',val=level-1),el('w:numId',val=nid)]);pp.append(np)
  for r in p.runs:fonts(r._r.get_or_add_rPr(),*ROLES[p.style.name])
  if strip:strip_prefix(p)
  changed.append({'paragraph':i+1,'level':level})
 canonicalize(doc);return changed
def main():
 import json
 p=argparse.ArgumentParser();p.add_argument('input');p.add_argument('output');p.add_argument('--style-map',help='JSON style name to level mapping');p.add_argument('--strip-manual-numbers',action='store_true');a=p.parse_args()
 if __import__('pathlib').Path(a.input).resolve()==__import__('pathlib').Path(a.output).resolve():p.error('Output must differ from input')
 d=Document(a.input);changed=repair(d,json.loads(a.style_map or '{}'),a.strip_manual_numbers);d.save(a.output);print(json.dumps(changed,ensure_ascii=False))
if __name__=='__main__':main()
