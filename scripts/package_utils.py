"""Settings/metadata-only patch; leaves document XML and field caches byte-identical."""
from zipfile import ZipFile,ZIP_DEFLATED
from pathlib import Path
from lxml import etree as E
from engine import NS,qn
def finalize(path):
 path=Path(path)
 with ZipFile(path) as z:parts={n:z.read(n) for n in z.namelist()}
 r=E.fromstring(parts['word/settings.xml']);flag=r.find('w:updateFields',NS)
 if flag is None:flag=E.SubElement(r,qn('w:updateFields'))
 flag.set(qn('w:val'),'true');parts['word/settings.xml']=E.tostring(r,encoding='UTF-8',xml_declaration=True,standalone=True)
 if 'docProps/core.xml' in parts:
  r=E.fromstring(parts['docProps/core.xml'])
  for node in r:
   if E.QName(node).localname in ['creator','lastModifiedBy','title','subject','description','keywords','category','identifier']:node.text=None
  parts['docProps/core.xml']=E.tostring(r,encoding='UTF-8',xml_declaration=True,standalone=True)
 # Word may add its current user's name to person.xml after editing; generated final copy is scrubbed.
 with ZipFile(path,'w',ZIP_DEFLATED) as z:
  for name,data in parts.items():z.writestr(name,data)
