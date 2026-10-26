"""Read all relevant OOXML parts without executing document content."""
import argparse,json,hashlib
from zipfile import ZipFile
from pathlib import Path
from lxml import etree as E
from engine import NS
def inspect(path):
 result={'sha256':hashlib.sha256(Path(path).read_bytes()).hexdigest(),'parts':{}}
 with ZipFile(path) as z:
  for name in z.namelist():
   data=z.read(name);entry={'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest()}
   if name.endswith(('.xml','.rels')):
    r=E.fromstring(data,E.XMLParser(resolve_entities=False,no_network=True))
    if name=='word/styles.xml':entry['styles']=[{'id':s.get('{'+NS['w']+'}styleId'),'definition':E.tostring(s,encoding='unicode')} for s in r.findall('w:style',NS)]
    elif name=='word/document.xml':
     entry.update(sections=[E.tostring(s,encoding='unicode') for s in r.findall('.//w:sectPr',NS)],paragraphs=len(r.findall('.//w:p',NS)),tables=len(r.findall('.//w:tbl',NS)),math=len(r.findall('.//m:oMath',NS)),fields=r.xpath('//w:instrText/text() | //w:fldSimple/@w:instr',namespaces=NS))
    elif name=='word/numbering.xml' or name=='word/settings.xml' or name.startswith(('word/header','word/footer')) or name.endswith('.rels'):entry['xml']=E.tostring(r,encoding='unicode')
   result['parts'][name]=entry
 return result
def main():
 p=argparse.ArgumentParser();p.add_argument('input');p.add_argument('--output',required=True);a=p.parse_args()
 Path(a.output).write_text(json.dumps(inspect(a.input),ensure_ascii=False,indent=2),encoding='utf-8');print('PASS OOXML inspected')
if __name__=='__main__':main()
