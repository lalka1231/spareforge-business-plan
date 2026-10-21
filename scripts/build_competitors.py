"""One canonical competitive matrix, shared by report and slides."""
from pathlib import Path
import csv,json
ROOT=Path(__file__).resolve().parents[1]
FULL_HEADERS=['Исполнитель / статус','Услуги, территория и цена','Комплект, маршрут и отчетность','Опыт, доля и доказательства']
SLIDE_HEADERS=['Исполнитель','Услуги / территория','Цена и база, руб./га','Доказательства / ограничения']
def records():
 with (ROOT/'data/competitors.csv').open(encoding='utf-8',newline='') as f:r=list(csv.DictReader(f))
 assert len(r)==6 and len({x['id'] for x in r})==6
 return r

def full_rows():
 out=[]
 for r in records():
  refs=''.join('['+s+']' for s in r['source_ids'].split(';') if s)
  out.append([r['name']+'; '+r['classification']+'; '+r['segment'],
   'Услуги: '+r['services']+'. Территория: '+r['geography']+'. Цена: '+(r['public_price_rub_per_ha']+' руб./га; ' if r['public_price_rub_per_ha'] and r['id']!='agrovector' else '')+r['price_basis']+'. НДС: '+r['vat_status']+'.',
   'Комплект: '+r['service_package']+'. Минимум/маршрут: '+r['minimum_route']+'. Отчет: '+r['reporting']+'.',
   'Опыт/объем: '+r['experience_volumes']+'. Доля: '+r['market_share_status']+'. '+refs+' Проверено: '+r['checked_at']+'. Публикация: '+r['publication_date']+'. '+r['verification_status']+'.'])
 # Every published numeric competitor price must survive the full projection.
 for r,row in zip(records(),out):
  if r['public_price_rub_per_ha']:
   assert r['public_price_rub_per_ha'] in row[1], 'Missing public price: '+r['id']
 return out

def slide_rows():
 return [[r['short_name']+(' (план)' if r['id']=='agrovector' else ''),r['slide_offer'],r['slide_price'],r['slide_evidence']] for r in records()]

def markdown():
 rows=[FULL_HEADERS,['---']*4]+full_rows()
 return '\n'.join('| '+' | '.join(c.replace('|','/') for c in row)+' |' for row in rows)+'\n'

if __name__=='__main__':
 data={'checked_at':'2026-10-20','full_headers':FULL_HEADERS,'full_rows':full_rows(),'slide_headers':SLIDE_HEADERS,'slide_rows':slide_rows()}
 (ROOT/'data/competitive_matrix.json').write_text(json.dumps(data,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
 print('Canonical matrix:',len(data['full_rows']),'participants; one full table, one slide projection')
