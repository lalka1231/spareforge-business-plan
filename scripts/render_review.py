"""Render actual Office artifacts with LibreOffice; rasterize PDFs for review."""
from pathlib import Path
import subprocess,shutil,json,hashlib
import pymupdf as fitz
from PIL import Image,ImageDraw
ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'review/2026-10-20';OUT.mkdir(parents=True,exist_ok=True)
exe=shutil.which('soffice')
if not exe:raise SystemExit('LibreOffice is required for actual rendering')
files=[ROOT/'documents/AGROVECTOR_FULL_BUSINESS_PLAN.docx',ROOT/'presentations/AGROVECTOR_22_SLIDES.pptx']
for f in files:
 p=subprocess.run([exe,'-env:UserInstallation=file:///private/tmp/agrovector-lo-review-20261020','--headless','--convert-to','pdf','--outdir',str(OUT),str(f)],capture_output=True,text=True,timeout=120)
 print(p.stdout,p.stderr)
 if p.returncode:raise RuntimeError(p.returncode)
report={'renderer':'LibreOffice','date':'2026-10-20','files':[],'visual_inspection':'pending; rendering itself is not visual acceptance'}
for f,prefix,cols,batch in [(files[0],'doc',3,9),(files[1],'deck',3,9)]:
 pdfpath=OUT/(f.stem+'.pdf');pdf=fitz.open(pdfpath)
 if prefix=='deck':assert len(pdf)==22
 paths=[];bounds=[]
 for i,page in enumerate(pdf):
  path=OUT/f'{prefix}-{i+1:03d}.png';page.get_pixmap(matrix=fitz.Matrix(1.25,1.25),alpha=False).save(path);paths.append(path)
  for block in page.get_text('dict')['blocks']:
   for line in block.get('lines',[]):
    for sp in line['spans']:
     b=fitz.Rect(sp['bbox'])
     if b.x0 < -1 or b.y0 < -1 or b.x1>page.rect.width+1 or b.y1>page.rect.height+1:bounds.append({'page':i+1,'text':sp['text'],'bbox':list(b)})
 contacts=[]
 for start in range(0,len(paths),batch):
  cellw=400;cellh=590 if prefix=='doc' else 255
  count=min(batch,len(paths)-start);sheet=Image.new('RGB',(cols*cellw,((count+cols-1)//cols)*cellh),'#cccccc');draw=ImageDraw.Draw(sheet)
  for j,path in enumerate(paths[start:start+batch]):
   im=Image.open(path);im.thumbnail((390,cellh-28));x=(j%cols)*cellw;y=(j//cols)*cellh;sheet.paste(im,(x,y+24));draw.text((x+5,y+5),path.stem,fill='black')
  cp=OUT/f'{prefix}-contact-{start//batch+1:02d}.jpg';sheet.save(cp,quality=90);contacts.append(str(cp.relative_to(ROOT)))
 report['files'].append({'source':str(f.relative_to(ROOT)),'source_sha256':hashlib.sha256(f.read_bytes()).hexdigest(),'pdf':str(pdfpath.relative_to(ROOT)),'pages':len(pdf),'text_outside_page':bounds,'contacts':contacts})
(ROOT/'data/render_validation.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8');print(json.dumps(report,ensure_ascii=False,indent=2))
