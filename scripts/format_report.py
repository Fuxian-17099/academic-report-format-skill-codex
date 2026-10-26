"""Conservative normalization: retain package objects and text; report ambiguous roles."""
import argparse,json
from pathlib import Path
from docx import Document
from docx.text.paragraph import Paragraph
from engine import BODY,ROLES,HEADINGS,fonts,qn,update_flag,canonicalize
from repair_heading_numbering import repair
def format_report(source,output,mapping=None,strip=False):
 if Path(source).resolve()==Path(output).resolve():raise ValueError('Never overwrite the source report')
 d=Document(source);changes=repair(d,mapping,strip);warnings=[]
 for i,node in enumerate(d.element.body.xpath('.//w:p')):
  p=Paragraph(node,d);name=p.style.name
  # TOC cached result, equations, OLE and drawings retain role-specific structures.
  if node.xpath('.//w:object | .//m:oMath | .//w:drawing | .//w:pPr/w:sectPr') or name.lower().startswith('toc'):continue
  if p.text.strip()=='参考文献':p.style='参考文献标题'
  elif name in ['Normal','正文','Body Text',BODY] and p.text.strip():
   if p._p.xpath('./w:pPr/w:numPr'):warnings.append({'paragraph':i+1,'message':'Unclassified list preserved'});continue
   p.style=BODY
  elif name=='参考文献':p.style='参考文献条目'
  elif name not in ROLES and p.text.strip():warnings.append({'paragraph':i+1,'message':'Unknown role preserved: '+name});continue
  if p.style.name in ROLES:
   for c in list(node.get_or_add_pPr()):
    if c.tag in [qn('w:spacing'),qn('w:ind'),qn('w:jc')]:node.get_or_add_pPr().remove(c)
   for r in p.runs:
    east,size,bold=ROLES[p.style.name];fonts(r._r.get_or_add_rPr(),east,size,bold if p.style.name in HEADINGS+['参考文献标题','参考文献条目'] else None)
 update_flag(d);canonicalize(d);d.save(output)
 return {'headings':changes,'warnings':warnings,'note':'Existing static citations/captions are not guessed. Supply explicit IDs and reconstruct via JSON if validator reports missing fields.'}
def main():
 p=argparse.ArgumentParser();p.add_argument('input');p.add_argument('output');p.add_argument('--style-map');p.add_argument('--strip-manual-numbers',action='store_true');a=p.parse_args()
 print(json.dumps(format_report(a.input,a.output,json.loads(a.style_map or '{}'),a.strip_manual_numbers),ensure_ascii=False,indent=2))
if __name__=='__main__':main()
