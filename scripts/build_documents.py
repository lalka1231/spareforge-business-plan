"""Assemble the investor dossier with generated schedule tables."""
from pathlib import Path
import json,csv
ROOT=Path(__file__).resolve().parents[1]
TITLES=['Резюме','SMART-цели','Описание продукта и стратегическая канва','Business Model Canvas','Маркетинговый план','Анализ конкурентов','Развернутый SWOT','Пять сил Портера','Организационно-правовая форма и регуляторный контур','Стейкхолдеры: власть и интерес','Производственный план','Организационный план','WBS / иерархическая структура работ','Диаграмма Ганта','Сетевой график','Матрица ответственности RACI','Финансовая модель','NPV и окупаемость','Потребность в инвестициях','Структура источников инвестиций','Основные технико-экономические показатели','Анализ рисков']
def table(headers,rows):
    return '| '+' | '.join(headers)+' |\n| '+' | '.join(['---']*len(headers))+' |\n'+'\n'.join('| '+' | '.join(str(x) for x in row)+' |' for row in rows)+'\n'
def build():
    sections=json.loads((ROOT/'data/sections.json').read_text());tasks=json.loads((ROOT/'data/schedule_summary.json').read_text())
    rows=tasks['tasks']
    sections['13']+='\n\n'+table(['WBS','ID','Работа','Результат'],[[r['wbs'],r['id'],r['task'],r['deliverable']] for r in rows])
    sections['14']+='\n\n[РАСЧЕТ] Календарь в календарных неделях; связи finish-to-start. Начало подготовки 07.09.2026. Даты — план, не выданные разрешения.\n\n'+table(['ID','Начало','Окончание','Недели','Резерв, нед.'],[[r['id'],r['start_date'],r['finish_date'],r['duration_weeks'],r['slack_weeks']] for r in rows])
    sections['14']+='\n```mermaid\ngantt\n    title АгроВектор — подготовка и первый сезон\n    dateFormat YYYY-MM-DD\n'
    for r in rows:
        sections['14']+=f"    {r['id']} {r['task']} :{'crit, ' if r['critical'] else ''}{r['id']}, {r['start_date']}, {r['duration_weeks']*7}d\n"
    sections['14']+='```\n'
    sections['15']+='\n\n[РАСЧЕТ] Критический путь: **'+' → '.join(tasks['critical_tasks'])+f"**. Продолжительность {tasks['duration_weeks']} недель, включая погодный буфер M и первый сезон K. Это плановый CPM, не прогноз срока выдачи документов.\n\n"
    sections['15']+=table(['ID','Предшественники','ES','EF','LS','LF','Резерв'],[[r['id'],r['predecessors'] or '—',r['es'],r['ef'],r['ls'],r['lf'],r['slack_weeks']] for r in rows])
    sections['15']+='\n```mermaid\nflowchart LR\n'
    for r in rows:
        sections['15']+=f"    {r['id']}[\"{r['id']}: {r['duration_weeks']} нед.\"]\n"
        for p in r['predecessors'].split(';'):
            if p:sections['15']+=f"    {p} --> {r['id']}\n"
    sections['15']+='```\n'
    sections['16']+='\n\n'+table(['ID','R — выполняет','A — принимает','C — консультирует','I — информируется'],[[r['id'],r['R'],r['A'],r['C'],r['I']] for r in rows])
    header='# АгроВектор\n\n## Инвестиционный бизнес-план ранней стадии\n\nСервис применения сельскохозяйственных дронов для мониторинга посевов и точечного внесения удобрений, средств защиты растений и биопрепаратов.\n\nВерсия 05.09.2026 · Плановый горизонт 2027–2031 · Валюта: рубли РФ.\n\n**Статус:** инвестиционная гипотеза до пилота; продажи, разрешения и финансирование не подтверждены. **Рекомендация:** условное поэтапное финансирование, а не безусловная закупка флота.\n\n**Маркировка:** [ФАКТ] — проверенный внешний источник; [РАСЧЕТ] — результат модели; [ДОПУЩЕНИЕ] — проектное решение или еще не подтвержденная гипотеза. Маркер в начале абзаца или перед таблицей распространяется на весь абзац/таблицу, если строка не помечена иначе.\n\n'
    body=header+'\n\n'.join(f'## {i}. {title}\n\n{sections[str(i)]}' for i,title in enumerate(TITLES,1))
    import re
    body=re.sub(r'(?<=[а-яА-ЯёЁ])(?=\d)|(?<=\d)(?=[а-яА-ЯёЁ])',' ',body)
    if (ROOT/'data/references_block.md').exists():
        body+='\n\n'+(ROOT/'data/references_block.md').read_text(encoding='utf-8')
    (ROOT/'docs/FINAL_BUSINESS_PLAN.md').write_text(body+'\n',encoding='utf-8')
    (ROOT/'data/section_titles.json').write_text(json.dumps(TITLES,ensure_ascii=False,indent=2),encoding='utf-8')
    print(f'Assembled {len(TITLES)} sections; {len(body)} characters')
if __name__=='__main__':build()
