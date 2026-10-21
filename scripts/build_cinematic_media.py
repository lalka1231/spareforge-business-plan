"""Prepare licensed stock + original 3D drone into moving slide scenes."""
from pathlib import Path
import subprocess, json, shutil
from PIL import Image,ImageDraw,ImageFilter
import imageio_ffmpeg
ROOT=Path(__file__).resolve().parents[1]; OUT=ROOT/'presentations/cinematic_assets';OUT.mkdir(exist_ok=True)
FF=imageio_ffmpeg.get_ffmpeg_exe()
def run(args):
 p=subprocess.run([FF,'-hide_banner','-loglevel','error','-y']+args,stdout=subprocess.PIPE,stderr=subprocess.PIPE)
 if p.returncode:raise RuntimeError(p.stderr.decode())
def ensure_sources():
 import urllib.request
 urls={'/tmp/agrovector-field-horizon.mp4':'https://www.pexels.com/download/video/9130276/','/tmp/agrovector-37280502.mp4':'https://videos.pexels.com/video-files/37280502/15793273_3840_2160_25fps.mp4','/tmp/agrovector-9289796.mp4':'https://videos.pexels.com/video-files/9289796/9289796-uhd_3840_2160_30fps.mp4','/tmp/agrovector-5308472.mp4':'https://videos.pexels.com/video-files/5308472/5308472-hd_1920_1080_30fps.mp4'}
 for target,url in urls.items():
  if Path(target).exists():continue
  with urllib.request.urlopen(urllib.request.Request(url,headers={'User-Agent':'Mozilla/5.0'}),timeout=120) as response,open(target,'wb') as dest:shutil.copyfileobj(response,dest)
  assert b'ftyp' in Path(target).read_bytes()[:64], 'Invalid downloaded MP4'
def build():
 ensure_sources()
 drone=Image.open(OUT/'drone_render.png').convert('RGBA');assert drone.getextrema()[3][0]==0
 # Spray is illustrative particle haze, not a dose/coverage simulation.
 canvas=Image.new('RGBA',(1200,1000)); mist=Image.new('RGBA',canvas.size);d=ImageDraw.Draw(mist)
 import random
 rng=random.Random(6)
 for k in range(1800):
  t=rng.random();x=600+rng.gauss(0,90+160*t);y=480+440*t
  d.ellipse((x,y,x+2.3,y+5),fill=(225,243,245,int(35*(1-t)+8)))
 mist=mist.filter(ImageFilter.GaussianBlur(1));canvas.alpha_composite(mist);canvas.alpha_composite(drone,(0,0));canvas.save(OUT/'drone_spray.png')
 # Soft editorial scrim keeps real typography readable without obscuring the product.
 shade=Image.new('RGBA',(1280,720));px=shade.load()
 for x in range(1280):
  for y in range(720):px[x,y]=(5,16,17,int(210*(1-x/1280)**1.4+30))
 shade.save(OUT/'scrim.png')
 sources={'hover':'/tmp/agrovector-field-horizon.mp4','flight':'/tmp/agrovector-field-horizon.mp4','detail':'/tmp/agrovector-37280502.mp4','survey':'/tmp/agrovector-9289796.mp4','wheat':'/tmp/agrovector-5308472.mp4'}
 for key,src in sources.items():
  args=['-ss','0','-i',src]
  if key in ('hover','flight'):
   args+=['-loop','1','-i',str(OUT/('drone_render.png' if key=='hover' else 'drone_spray.png'))]
   size='600:-1' if key=='hover' else '460:-1'
   x='640+15*sin(2*PI*t/8)' if key=='hover' else '640+180*sin(2*PI*t/8)'
   y='105+12*sin(2*PI*t/4)' if key=='hover' else '115+14*sin(2*PI*t/4)'
   if key=='flight':args+=['-i','/tmp/agrovector-5308472.mp4']
   bgfilter='scale=1600:900:force_original_aspect_ratio=increase,crop=1280:720:160:180' if key=='flight' else 'scale=1280:720:force_original_aspect_ratio=increase,crop=1280:720'
   if key=='flight':
    graph=f"[0:v]{bgfilter},setsar=1[bg];[2:v]scale=1280:720,format=rgba,geq=r='r(X,Y)':g='g(X,Y)':b='b(X,Y)':a='clip((Y-410)/310,0,1)*230'[fg];[bg][fg]overlay=0:0:shortest=1[field];[1:v]scale={size},format=rgba[drone];[field][drone]overlay=x='{x}':y='{y}':shortest=1[v]"
   else:graph=f'[0:v]{bgfilter},setsar=1[bg];[1:v]scale={size},format=rgba[drone];[bg][drone]overlay=x=\'{x}\':y=\'{y}\':shortest=1[v]'
   args+=['-filter_complex',graph,'-map','[v]']
  else:args+=['-vf','scale=1280:720:force_original_aspect_ratio=increase,crop=1280:720,setsar=1']
  args+=['-t','8' if key not in ('wheat',) else '6','-r','25','-an','-c:v','libx264','-preset','fast','-crf','23','-pix_fmt','yuv420p','-movflags','+faststart',str(OUT/(key+'.mp4'))]
  run(args)
  run(['-i',str(OUT/(key+'.mp4')),'-frames:v','1',str(OUT/(key+'_poster.jpg'))])
  print('Rendered',key,flush=True)
 ledger={'license':'https://www.pexels.com/license/','classification':'Third-party illustrative stock plus original synthetic drone; not AgroVector operations','sources':[{'role':'golden field horizon','url':'https://www.pexels.com/video/aerial-footage-of-wheat-field-9130276/'},{'role':'spraying drone','url':'https://www.pexels.com/video/aerial-drone-spraying-pesticides-in-farmland-37280502/'},{'role':'survey texture','url':'https://www.pexels.com/video/aerial-footage-of-a-wheat-field-9289796/'},{'role':'close-up grain','url':'https://www.pexels.com/video/close-up-view-of-a-wheat-5308472/'}],'limitations':['Stock shots are different locations, not one documented field','Composite flight and spray are concept visualization, not physical/agronomic simulation','Playback loop can visibly restart the stock clip; continuous single-camera take is not asserted'],'codec':'H.264 yuv420p 1280x720 25fps, no audio'}
 (OUT/'asset_ledger.json').write_text(json.dumps(ledger,ensure_ascii=False,indent=2)+'\n')
if __name__=='__main__':build()
