"""Register dated excerpts retrieved on 20.10.2026; preserve historical IDs."""
from pathlib import Path
import json,subprocess,sys
ROOT=Path(__file__).resolve().parents[1]
TOOL=Path.home()/'.hermes/skills/research/grounded-citations/scripts/sources.py'
def cli(*args):
 return subprocess.check_output([sys.executable,str(TOOL),'--ledger',str(ROOT/'data/source_ledger.json'),*args],text=True)
NEW=[
 ('https://xag-ap.ru/p100pro.php','Агропрофи: XAG P100 Pro, карточка модели','2 990 000 ₽','Прямое HTTPS-чтение; извлекатель вернул иную версию 2,9 млн / 50 л, не использованную как текущая оферта.'),
 ('https://www.xa.com/en/p100pro/p100prospecs','XAG P100 Pro: официальная спецификация','Rated Capacity 50 L','Производитель; номинальная емкость бака. Не максимальная взлетная масса.'),
 ('https://16.fsvps.gov.ru/news/s-1-marta-2026-goda-vstupili-v-zakonnuyu-silu-izmeneniya-v-zakonodatelstvo-rf-kasayushhiesya-primeneniya-pesticzidov-i-agrohimikatov/','Россельхознадзор: уведомления и ФГИС Сатурн с 01.03.2026','Уведомления должны размещаться в средствах массовой информации, а также в Федеральной государственной информационной системе прослеживаемости пестицидов и агрохимикатов (ФГИС «Сатурн»).','Официальное разъяснение; дата публикации в полученном тексте не указана; применимость проверяется по конкретному заданию.'),
 ('http://publication.pravo.gov.ru/Document/View/1600202211300002','Указ Республики Татарстан от 25.11.2022 № УП-813','Об использовании беспилотных воздушных судов на территории Республики Татарстан','Публикация 30.11.2022; сейчас проверена карточка, не полный действующий текст исключений.')]
for url,title,quote,lim in NEW:
 print(cli('add',url,'--title',title).strip())
 ledger=json.loads((ROOT/'data/source_ledger.json').read_text());s=next(x for x in ledger['sources'] if x['url'].rstrip('/')==url.rstrip('/'));sid=s['id']
 if sid==30:
  quote='Rated Capacity\n50 L'
 body=title+'\n'+url+'\nПроверено: 2026-10-20\nСтатус: '+lim+'\n\n'+quote+'\n'
 if sid==29:body+='\nБак для жидкости - 60 л\n'
 path=ROOT/f'docs/evidence/{sid:02d}.txt';path.write_text(body,encoding='utf-8')
 print(cli('quote',str(sid),'--text',quote,'--from',str(path)).strip())
 if sid==29:print(cli('quote',str(sid),'--text','Бак для жидкости - 60 л','--from',str(path)).strip())
# Add current short excerpts to older records while retaining their historical evidence.
updates={
 2:('Эффективность обработки до 18 гектаров в час, ширина распыления до 7 метров.','Страница услуги прочитана; главная недоступна; реклама не фактическая сменная выработка.'),
 3:('50 000+ гектаров обработано','Накопленный счетчик без года/территории; прямой HTML содержит цену P100 Pro 2 990 000 ₽ и бак 60 л.'),
 4:('от 200 до 700 га/смена','Заявленная возможность, не подтвержденная смена одной бригады.'),
 5:('в 2025 году — 8 дронов и 2 прицепа в республике, разрешения оформлены, обработано >40 000 га по нашим проектам','Самозаявление о проектах Татарстана; не подтвержден собственный оплаченный сервисный объем.'),
 6:('Разметка и последующая обработка картофельного поля площадью 500 Га','Кейс без даты и актов; не годовая выручка или доля.'),
 7:('Полный комплект документов — акты, сертификаты, закрывающие документы.','Заявленная отчетность; минимальный заказ и налоговый состав тарифа не раскрыты.')}
ledger=json.loads((ROOT/'data/source_ledger.json').read_text())
for sid,(quote,lim) in updates.items():
 s=next(x for x in ledger['sources'] if x['id']==sid);path=ROOT/f'docs/evidence/{sid:02d}.txt'
 old=path.read_text(); marker='\n\nПовторное чтение 2026-10-20\n'
 old=old.split(marker)[0];path.write_text(old+marker+lim+'\n'+quote+'\n',encoding='utf-8')
 print(cli('quote',str(sid),'--text',quote,'--from',str(path)).strip())
ledger=json.loads((ROOT/'data/source_ledger.json').read_text())
checked={2,3,4,5,6,7,8,9,17,29,30,31,32}
for s in ledger['sources']:
 s['last_attempted']='2026-10-20' if s['id'] in checked|{1,11,12,16,18,19,20,21,22,23,24,25,26,27,28} else None
 if s['id'] in checked:
  s['accessed']='2026-10-20';s['verification_status']='Содержательное новое чтение; ограничения см. исследовательские отчеты A/C и реестр.'
 else:s['verification_status']='Архив 05.09.2026; полная повторная проверка не выполнена или источник недоступен. Не считать свежим подтверждением.'
 if s['id']==32:s['verification_status']='Проверена только официальная карточка; действующие исключения и полный текст не подтверждены.'
(ROOT/'data/source_ledger.json').write_text(json.dumps(ledger,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
old=json.loads((ROOT/'data/source_map.json').read_text());by={s['id']:s for s in old}
for s in ledger['sources']:
 rec=by.setdefault(s['id'],{'id':s['id'],'url':s['url'],'title':s['title'],'quote':s['quotes'][-1]['text']})
 rec.update(checked_at=s['accessed'],last_attempted=s.get('last_attempted'),verification_status=s['verification_status'])
 rec['quotes']=[q['text'] for q in s['quotes']]
(ROOT/'data/source_map.json').write_text(json.dumps(list(by.values()),ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
# Generate source list and a transparent per-source verification register.
header='# Реестр источников AgroVector — 20.10.2026\n\nДаты проверок различаются. Снимок/карточка не приравниваются к актуальному полному нормативному тексту. Новые исследования: research_2026_10_20_A.md и research_2026_10_20_C.md.\n\n'
text=header
for s in ledger['sources']:
 text+=f"## [{s['id']}] {s['title']}\n\n{s['url']}\n\nПоследнее содержательное чтение: {s['accessed']}; попытка обновления: {s.get('last_attempted') or 'не выполнялась в этой редакции'}. {s['verification_status']}\n\n"
 for q in s['quotes'][-2:]:text+='> '+q['text'].replace('\n','\n> ')+'\n\n'
(ROOT/'docs/09_SOURCES.md').write_text('\n'.join(line.rstrip() for line in text.splitlines()).rstrip()+'\n',encoding='utf-8')
print('Sources registered:',len(ledger['sources']))
