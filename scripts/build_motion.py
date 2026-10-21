"""Non-destructive motion edition of the existing 22-slide deck."""
from pathlib import Path
import json, zipfile, hashlib
from lxml import etree as E
from pptx import Presentation
from pptx.util import Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
ROOT=Path(__file__).resolve().parents[1]
SRC=ROOT/'presentations/AGROVECTOR_22_SLIDES_DESIGN.pptx'
OUT=ROOT/'presentations/AGROVECTOR_22_SLIDES_MOTION.pptx'
P='http://schemas.openxmlformats.org/presentationml/2006/main'
MC='http://schemas.openxmlformats.org/markup-compatibility/2006'
M='http://schemas.microsoft.com/office/powerpoint/2015/09/main'
NS={'p':P}
def el(name,**attrs):return E.Element('{'+P+'}'+name,{k:str(v) for k,v in attrs.items()})
def sub(parent,name,**attrs):x=el(name,**attrs);parent.append(x);return x

def timing(groups):
 timing=el('timing'); root=sub(sub(timing,'tnLst'),'par'); ct=sub(root,'cTn',id=1,dur='indefinite',restart='never',nodeType='tmRoot'); seq=sub(sub(ct,'childTnLst'),'seq',concurrent=1,nextAc='seek'); main=sub(seq,'cTn',id=2,dur='indefinite',nodeType='mainSeq'); children=sub(main,'childTnLst'); counter=3
 for group in groups:
  wrapper=sub(children,'par'); c=sub(wrapper,'cTn',id=counter,fill='hold');counter+=1
  sub(sub(c,'stCondLst'),'cond',delay='indefinite'); inner=sub(sub(c,'childTnLst'),'par'); ic=sub(inner,'cTn',id=counter,fill='hold');counter+=1;sub(sub(ic,'stCondLst'),'cond',delay=0); effects=sub(ic,'childTnLst')
  for j,spid in enumerate(group):
   par=sub(effects,'par'); effect=sub(par,'cTn',id=counter,presetID=10,presetClass='entr',presetSubtype=0,fill='hold',nodeType='clickEffect' if j==0 else 'withEffect');counter+=1
   sub(sub(effect,'stCondLst'),'cond',delay=0); cl=sub(effect,'childTnLst'); fade=sub(cl,'animEffect',transition='in',filter='fade'); bh=sub(fade,'cBhvr');sub(bh,'cTn',id=counter,dur=420);counter+=1;sub(sub(bh,'tgtEl'),'spTgt',spid=spid)
 prev=sub(seq,'prevCondLst');cond=sub(prev,'cond',evt='onPrev',delay=0);sub(cond,'tgtEl').append(el('sldTgt'))
 nxt=sub(seq,'nextCondLst');cond=sub(nxt,'cond',evt='onNext',delay=0);sub(cond,'tgtEl').append(el('sldTgt'))
 return timing

def build():
 prs=Presentation(SRC); plan=[]
 for i,s in enumerate(prs.slides,1):
  # Stable !! names explicitly pair just the moving progress geometry in Morph.
  rail=s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(.55), Inches(.08), Inches(12.15), Inches(.035));rail.name='!!AV_progress_rail';rail.fill.solid();rail.fill.fore_color.rgb=RGBColor.from_string('37584A');rail.line.fill.background()
  bar=s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(.55), Inches(.08), Inches(12.15*i/22), Inches(.035));bar.name='!!AV_progress_fill';bar.fill.solid();bar.fill.fore_color.rgb=RGBColor.from_string('C9EF70');bar.line.fill.background()
  candidates=[sh for sh in s.shapes if 1.75<sh.top/914400<6.1 and ((sh.has_text_frame and sh.text.strip()) or sh.has_chart or sh.has_table)]
  # Card layouts reveal one complete text column per click, keeping claims and qualifiers together.
  if i in {1,3,4,5,7,10,11,13,19,20,21}:
   buckets={}
   for sh in candidates:buckets.setdefault(round(sh.left/914400),[]).append(sh.shape_id)
   groups=list(buckets.values())
   if len(groups)>4:groups=[sum(groups[j:j+3],[]) for j in range(0,len(groups),3)]
  elif i in {17,18}:groups=[[sh.shape_id for sh in candidates]]
  else:groups=[] # Dense evidence tables and schedules remain instantly legible.
  plan.append({'slide':i,'transition':'morph' if i in {3,4,5,11,13,14,15,17,18,19,20,21,22} else 'fade','click_groups':groups})
 prs.save(OUT)
 with zipfile.ZipFile(OUT) as z:parts={n:z.read(n) for n in z.namelist()}
 for item in plan:
  name=f'ppt/slides/slide{item["slide"]}.xml';tree=E.fromstring(parts[name])
  for tag in ['transition','timing']:
   for old in tree.findall('p:'+tag,NS):tree.remove(old)
  trans=el('transition',spd='med',advClick=1)
  if item['transition']=='morph':
   ac=E.SubElement(trans,'{'+MC+'}AlternateContent',nsmap={'mc':MC}); choice=E.SubElement(ac,'{'+MC+'}Choice',Requires='p159',nsmap={'p159':M}); E.SubElement(choice,'{'+M+'}morph',option='byObject'); fallback=E.SubElement(ac,'{'+MC+'}Fallback');fallback.append(el('fade'))
  else:trans.append(el('fade'))
  index=next((j for j,x in enumerate(tree) if E.QName(x).localname in ('timing','extLst')),len(tree));tree.insert(index,trans)
  if item['click_groups']:tree.insert(index+1,timing(item['click_groups']))
  parts[name]=E.tostring(tree,xml_declaration=True,encoding='UTF-8',standalone=True)
 with zipfile.ZipFile(OUT,'w',zipfile.ZIP_DEFLATED) as z:
  for n,b in parts.items():z.writestr(n,b)
 # Content and notes must be identical; progress shapes contain no text.
 after=Presentation(OUT); before=Presentation(SRC)
 assert len(after.slides)==22
 for a,b in zip(before.slides,after.slides):
  assert [s.text for s in a.shapes if s.has_text_frame and s.text]==[s.text for s in b.shapes if s.has_text_frame and s.text]
  assert a.notes_slide.notes_text_frame.text==b.notes_slide.notes_text_frame.text
 with zipfile.ZipFile(OUT) as z:
  assert z.testzip() is None
  for p in plan:
   t=E.fromstring(z.read(f'ppt/slides/slide{p["slide"]}.xml')); ids={x.get('id') for x in t.findall('.//p:cNvPr',NS)}
   assert all(str(x) in ids for group in p['click_groups'] for x in group)
 report={'output':str(OUT),'sha256':hashlib.sha256(OUT.read_bytes()).hexdigest(),'slides':22,'morph_transitions':sum(x['transition']=='morph' for x in plan),'animated_slides':sum(bool(x['click_groups']) for x in plan),'plan':plan,'verified':'ZIP, shape targets, unchanged text/notes; native playback not yet verified','compatibility':'Morph requires recent Microsoft PowerPoint; fade fallback. No automatic slide advance; next click reveals a group.'}
 (ROOT/'presentations/MOTION_VALIDATION.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
 print(json.dumps({k:v for k,v in report.items() if k!='plan'},ensure_ascii=False,indent=2))
if __name__=='__main__':build()
