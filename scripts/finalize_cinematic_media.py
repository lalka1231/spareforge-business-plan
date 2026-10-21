"""Convert native movie play commands from click to automatic With Previous.
Run after native_cinematic.applescript; native open/save preserves this form.
"""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
ROOT=Path(__file__).resolve().parents[1];PATH=ROOT/'presentations/AGROVECTOR_CINEMATIC.pptx';P='http://schemas.openxmlformats.org/presentationml/2006/main';NS={'p':P}
def finalize():
 with ZipFile(PATH) as z:parts={n:z.read(n) for n in z.namelist()}
 for i in range(1,23):
  name=f'ppt/slides/slide{i}.xml';tree=E.fromstring(parts[name]);main=tree.find('.//p:cTn[@nodeType="mainSeq"]',NS)
  assert main is not None,'Run native_cinematic.applescript first'
  assert main.find('.//p:cmd[@cmd="playFrom(0.0)"]',NS) is not None
  for cond in main.findall('.//p:cond',NS):
   if cond.get('delay')=='indefinite':cond.set('delay','0')
  for ct in main.findall('.//p:cTn',NS):
   if ct.get('nodeType')=='clickEffect':ct.set('nodeType','withEffect')
  parts[name]=E.tostring(tree,xml_declaration=True,encoding='UTF-8',standalone=True)
 with ZipFile(PATH,'w',ZIP_DEFLATED) as z:
  for n,data in parts.items():z.writestr(n,data)
 print('22 native play commands now With Previous with zero wrapper delays; loop retained.')
if __name__=='__main__':finalize()
