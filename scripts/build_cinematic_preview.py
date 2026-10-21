"""Playable intro preview (not a recording of native PowerPoint playback)."""
from pathlib import Path
from PIL import Image,ImageDraw,ImageFont
import subprocess,imageio_ffmpeg
ROOT=Path(__file__).resolve().parents[1]; A=ROOT/'presentations/cinematic_assets';V=ROOT/'review/cinematic';V.mkdir(exist_ok=True)
FF=imageio_ffmpeg.get_ffmpeg_exe();FONT='/System/Library/Fonts/Supplemental/Arial.ttf';BOLD='/System/Library/Fonts/Supplemental/Arial Bold.ttf'
def run(args):
 p=subprocess.run([FF,'-y','-hide_banner','-loglevel','error']+args,stdout=subprocess.PIPE,stderr=subprocess.PIPE)
 if p.returncode:raise RuntimeError(p.stderr.decode())
for i in (1,2):
 im=Image.open(A/'scrim.png').convert('RGBA');d=ImageDraw.Draw(im)
 def t(x,y,text,size=30,bold=False,c='white'):
  d.multiline_text((x,y),text,font=ImageFont.truetype(BOLD if bold else FONT,size),fill=c,spacing=14)
 t(62,45,'АГРОВЕКТОР   /   '+('01  /  РЕЗЮМЕ' if i==1 else '02  /  SMART-ЦЕЛИ'),16,True,'#D6ED89')
 if i==1:
  t(62,130,'АгроВектор',72,True);t(65,252,'Мониторинг посевов.\nТочечное внесение.\nПроверяемый отчет.',30)
  t(65,430,'Татарстан / План 2027–2031',22,False,'#D6ED89');t(65,505,'9,5 млн ₽ — стартовая потребность\nЗаказы и разрешения не подтверждены.',23)
 else:
  t(62,125,'Сначала доказательства —\nзатем сезон',45,True);t(65,350,'Правовая карта → интервью → пилот',24,True,'#D6ED89');t(65,412,'Пилот 100 га — до 30.04.2027\nВ пределах планового объема сезона.\nДаты — цели, не выполненные работы.',24)
 d.rectangle((0,673,1280,720),fill=(5,15,15,220));t(62,688,'Визуализация концепции + сток Pexels. Не съемка AgroVector.',13,False,'#CCDACF')
 im.save(A/f'intro-overlay-{i}.png')
 run(['-i',str(A/('hover.mp4' if i==1 else 'flight.mp4')),'-loop','1','-i',str(A/f'intro-overlay-{i}.png'),'-filter_complex','[0:v][1:v]overlay=0:0:shortest=1[v]','-map','[v]','-t','6','-an','-c:v','libx264','-preset','fast','-crf','21','-pix_fmt','yuv420p',str(V/f'intro-scene-{i}.mp4')])
run(['-i',str(V/'intro-scene-1.mp4'),'-i',str(V/'intro-scene-2.mp4'),'-filter_complex','[0:v]settb=AVTB,setpts=PTS-STARTPTS,fps=25[a];[1:v]settb=AVTB,setpts=PTS-STARTPTS,fps=25[b];[a][b]xfade=transition=slideup:duration=1.1:offset=4.9[v]','-map','[v]','-an','-c:v','libx264','-preset','fast','-crf','21','-pix_fmt','yuv420p','-movflags','+faststart',str(V/'AGROVECTOR_INTRO_PREVIEW.mp4')])
print(V/'AGROVECTOR_INTRO_PREVIEW.mp4')
