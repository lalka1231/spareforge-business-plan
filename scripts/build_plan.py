"""Generate one consistent WBS, CPM, Gantt and RACI dataset (stdlib)."""
from pathlib import Path
import csv,json
from datetime import date,timedelta
ROOT=Path(__file__).resolve().parents[1]
# all durations and assignments are planning assumptions; finish-to-start, calendar weeks
TASKS=[
('A','1.1','Проверка региона, полетов и схемы АХР',4,[], 'Юрист','Директор','Агроном;Операционный руководитель','Инвестор','Письменная карта разрешений и стоп-факторов'),
('B','1.2','Интервью и предварительные заявки хозяйств',5,[], 'Продажи','Директор','Агроном','Инвестор','20 интервью; заявки на 6000 га-проходов'),
('C','1.3','ООО, договоры, учет и страховой проект',3,['A'],'Юрист','Директор','Бухгалтер','Инвестор','ООО; договоры; страховые предложения'),
('D','2.1','Инвестиционное решение и финансирование',1,['B','C'],'Директор','Инвестор','Бухгалтер;Юрист','Команда','9.5 млн руб.; пройден юридический gate'),
('E','2.2','Закупка и приемка оборудования',5,['D'],'Операционный руководитель','Директор','Агроном','Бухгалтер','Комплектность; серийные номера; акты'),
('F','2.3','Подбор и гражданская переподготовка',6,['D'],'Операционный руководитель','Директор','Учебный партнер;Агроном','Инвестор','Два оператора; экзамен и допуск'),
('G','2.4','Разрешительная готовность и учет БВС',6,['E'],'Юрист','Директор','Операционный руководитель','Клиенты','Документы на конкретные БВС/работы; либо no-go'),
('H','3.1','SOP, CRM, карты и контроль качества',3,['E','F'],'Агроном','Операционный руководитель','Юрист','Продажи','Чек-листы; шаблон отчета; аварийный план'),
('I','3.2','Стендовые и контролируемые приемочные испытания',2,['G','H'],'Оператор','Операционный руководитель','Агроном;Клиент','Директор','Акт приемки; полевой пилот 100 га включен в апрельский сезон'),
('J','3.3','Сезонные договоры и маршрутизация',2,['I'],'Продажи','Директор','Агроном;Клиент','Бухгалтер','10 клиентов; согласованный календарь'),
('M','3.4','Предсезонная готовность и ожидание погодного окна',8,['J'],'Операционный руководитель','Директор','Агроном;Клиент','Инвестор','Повторная проверка разрешений и погодного окна'),
('K','4.1','Первый коммерческий сезон',22,['M'],'Оператор','Операционный руководитель','Агроном;Клиент','Директор','6000 га-проходов; 4000 га-обследований'),
('L','4.2','Аудит сезона и решение о продолжении',2,['K'],'Бухгалтер','Директор','Агроном;Операционный руководитель','Инвестор','Факт P&L/ДДС; претензии; продления')]
start=date(2026,9,7)
rows=[]; by={}
for id,wbs,name,dur,pred,r,a,c,i,deliverable in TASKS:
    es=max((by[p]['ef'] for p in pred),default=0); ef=es+dur
    row=dict(id=id,wbs=wbs,task=name,duration_weeks=dur,predecessors=';'.join(pred),es=es,ef=ef,start_date=str(start+timedelta(weeks=es)),finish_date=str(start+timedelta(weeks=ef)-timedelta(days=1)),R=r,A=a,C=c,I=i,deliverable=deliverable,type='[ДОПУЩЕНИЕ]')
    rows.append(row);by[id]=row
end=max(r['ef'] for r in rows)
for row in reversed(rows):
    successors=[r for r in rows if row['id'] in r['predecessors'].split(';')]
    row['lf']=min((r['ls'] for r in successors),default=end);row['ls']=row['lf']-row['duration_weeks'];row['slack_weeks']=row['ls']-row['es'];row['critical']=row['slack_weeks']==0
with (ROOT/'data/project_plan.csv').open('w',encoding='utf-8',newline='') as f:
    w=csv.DictWriter(f,fieldnames=rows[0].keys(),lineterminator='\n');w.writeheader();w.writerows(rows)
(ROOT/'data/schedule_summary.json').write_text(json.dumps({'duration_weeks':end,'critical_tasks':[r['id'] for r in rows if r['critical']],'tasks':rows},ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'weeks':end,'critical':[r['id'] for r in rows if r['critical']],'commercial_start':by['K']['start_date']},ensure_ascii=False))
