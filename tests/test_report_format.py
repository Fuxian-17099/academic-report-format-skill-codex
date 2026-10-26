"""Executable positive, destructive negative and semantic refresh tests."""
import unittest,json,sys,re,uuid,shutil
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from copy import deepcopy
from lxml import etree as E
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT/'scripts'))
from engine import generate,NS,qn,equation,math_node
from validate_report import validate
from format_report import format_report
from repair_heading_numbering import repair
from docx import Document

class ReportTests(unittest.TestCase):
 def setUp(self):
  scratch=Path.cwd()/'work';scratch.mkdir(exist_ok=True)
  self.dir=scratch/('test-'+uuid.uuid4().hex);self.dir.mkdir();self.addCleanup(shutil.rmtree,self.dir)
  self.data=json.loads((ROOT/'examples/example-input.json').read_text(encoding='utf-8'))
  self.path=self.dir/'report.docx';generate(self.data,self.path,ROOT/'examples')
 def patch(self,part,fn):
  with ZipFile(self.path) as z:parts={n:z.read(n) for n in z.namelist()}
  root=E.fromstring(parts[part]);fn(root);parts[part]=E.tostring(root)
  with ZipFile(self.path,'w',ZIP_DEFLATED) as z:
   for n,b in parts.items():z.writestr(n,b)
 def fail_has(self,fragment):
  result=validate(self.path);self.assertEqual(result['status'],'FAIL');self.assertTrue(any(fragment in e['message'] for e in result['errors']),result)
 def test_generation(self):self.assertEqual(validate(self.path)['status'],'PASS')
 def test_bad_number_font(self):
  self.patch('word/numbering.xml',lambda r:r.xpath('//w:abstractNum[w:lvl/w:pStyle]/w:lvl/w:rPr/w:rFonts',namespaces=NS)[0].set(qn('w:ascii'),'宋体'));self.fail_has('NUMBER')
 def test_bad_number_bold(self):
  self.patch('word/numbering.xml',lambda r:r.xpath('//w:abstractNum[w:lvl/w:pStyle]/w:lvl/w:rPr/w:b',namespaces=NS)[0].set(qn('w:val'),'0'));self.fail_has('NUMBER')
 def test_bad_body_font(self):
  self.patch('word/styles.xml',lambda r:r.xpath('//w:style[w:name/@w:val="正文_改"]/w:rPr/w:rFonts',namespaces=NS)[0].set(qn('w:hAnsi'),'Calibri'));self.fail_has('hAnsi')
 def test_body_normal(self):
  self.patch('word/document.xml',lambda r:r.xpath('//w:p[w:pPr/w:pStyle/@w:val="正文_改"]/w:pPr/w:pStyle',namespaces=NS)[0].set(qn('w:val'),'Normal'));self.fail_has('ordinary body')
 def test_manual_heading(self):
  self.patch('word/document.xml',lambda r:setattr(r.xpath('//w:p[w:pPr/w:pStyle/@w:val="Heading1"]/w:r/w:t',namespaces=NS)[0],'text','1 伪手工标题'));self.fail_has('manual duplicate')
 def test_reference_heading_font(self):
  self.patch('word/styles.xml',lambda r:r.xpath('//w:style[w:name/@w:val="参考文献标题"]/w:rPr/w:rFonts',namespaces=NS)[0].set(qn('w:eastAsia'),'黑体'));self.fail_has('Chinese font')
 def test_missing_bookmark(self):
  self.patch('word/document.xml',lambda r:r.remove(r[0]) if False else r.xpath('//w:bookmarkStart',namespaces=NS)[0].getparent().remove(r.xpath('//w:bookmarkStart',namespaces=NS)[0]));self.fail_has('reference target missing')
 def test_duplicate_bookmark(self):
  self.patch('word/document.xml',lambda r:r.xpath('//w:bookmarkStart',namespaces=NS)[1].set(qn('w:name'),r.xpath('//w:bookmarkStart',namespaces=NS)[0].get(qn('w:name'))));self.fail_has('duplicate bookmark')
 def test_missing_seq(self):
  def change(r):
   n=next(n for n in r.xpath('//w:instrText',namespaces=NS) if 'SEQ Figure' in n.text);n.text=' PAGE '
  self.patch('word/document.xml',change);self.fail_has('SEQ Figure')
 def test_bad_image_width(self):
  self.patch('word/document.xml',lambda r:r.xpath('//wp:extent',namespaces=NS)[0].set('cx','99999999'));self.fail_has('image exceeds')
 def test_bad_table_width(self):
  self.patch('word/document.xml',lambda r:r.xpath('//w:tblGrid/w:gridCol',namespaces=NS)[0].set(qn('w:w'),'20000'));self.fail_has('table grid')
 def test_bad_field_structure(self):
  self.patch('word/document.xml',lambda r:r.xpath('//w:fldChar[@w:fldCharType="end"]',namespaces=NS)[-1].getparent().remove(r.xpath('//w:fldChar[@w:fldCharType="end"]',namespaces=NS)[-1]));self.fail_has('unclosed field')
 def test_no_update_flag(self):
  self.patch('word/settings.xml',lambda r:r.find('w:updateFields',NS).set(qn('w:val'),'false'));self.fail_has('updateFields')
 def test_missing_reference_rejected(self):
  self.data['blocks'].append({'type':'paragraph','text':'[[ref:absent]]'})
  with self.assertRaisesRegex(ValueError,'Unknown reference'):generate(self.data,self.dir/'bad.docx',ROOT/'examples')
 def test_duplicate_id_rejected(self):
  self.data['references'].append(self.data['references'][0])
  with self.assertRaisesRegex(ValueError,'Duplicate'):generate(self.data,self.dir/'bad.docx',ROOT/'examples')
 def test_heading_jump_rejected(self):
  self.data['blocks'][0]['level']=3
  with self.assertRaisesRegex(ValueError,'jump'):generate(self.data,self.dir/'bad.docx',ROOT/'examples')
 def test_mathtype_slot(self):
  eq=next(b for b in self.data['blocks'] if b['type']=='equation');eq['mode']='mathtype';generate(self.data,self.path,ROOT/'examples');self.assertEqual(validate(self.path)['status'],'PASS')
 def test_all_math_structures(self):
  cases=[{'type':'sqrt','body':'α'},{'type':'root','degree':'3','body':'x'},{'type':'integral','lower':'0','upper':'∞','body':'x dx'},{'type':'sum','lower':'i=0','upper':'n','body':'i'},{'type':'matrix','rows':[['a','b'],['c','d']]},{'type':'subsup','base':'x','sub':'i','sup':'2'}]
  for c in cases:
   node=equation({'math':c});self.assertEqual(E.QName(node).localname,'oMath');self.assertGreater(len(node),0)
  with self.assertRaises(ValueError):math_node({'type':'matrix','rows':[['1'],['1','2']]})
 def test_formatter_preserves_text_objects(self):
  out=self.dir/'formatted.docx';format_report(self.path,out)
  with ZipFile(self.path) as a,ZipFile(out) as b:
   ar=E.fromstring(a.read('word/document.xml'));br=E.fromstring(b.read('word/document.xml'))
   self.assertEqual(ar.xpath('//w:t/text()',namespaces=NS),br.xpath('//w:t/text()',namespaces=NS))
   for expr in ['//m:oMath','//w:bookmarkStart','//wp:extent','//w:instrText']:self.assertEqual(len(ar.xpath(expr,namespaces=NS)),len(br.xpath(expr,namespaces=NS)))
   for n in a.namelist():
    if n.startswith('word/media/'):self.assertEqual(a.read(n),b.read(n))
  self.assertEqual(validate(out)['status'],'PASS')
 def test_heading_repair(self):
  self.patch('word/numbering.xml',lambda r:r.xpath('//w:abstractNum[w:lvl/w:pStyle]/w:lvl/w:rPr/w:rFonts',namespaces=NS)[0].set(qn('w:ascii'),'宋体'))
  d=Document(self.path);repair(d);d.save(self.path);self.assertEqual(validate(self.path)['status'],'PASS')
 def test_template_privacy(self):
  with ZipFile(ROOT/'assets/academic-report-template.docx') as z:
   self.assertFalse(any(x.startswith(('word/media/','word/embeddings/','customXml/','word/comments')) for x in z.namelist()))
   r=E.fromstring(z.read('docProps/core.xml'))
   self.assertFalse(any((n.text or '').strip() for n in r if E.QName(n).localname in ['creator','lastModifiedBy','title','subject','description']))
   r=E.fromstring(z.read('word/document.xml'));self.assertEqual(r.xpath('//w:t/text()',namespaces=NS),[])
 def test_markdown_interface(self):
  data={'title':'Markdown 输入验证','markdown':'# 第一章\n\n正文 FPGA 100 MHz。\n\n## 节\n\n### 小节\n\n内容。'}
  generate(data,self.path);self.assertEqual(validate(self.path)['status'],'PASS')
if __name__=='__main__':unittest.main()
