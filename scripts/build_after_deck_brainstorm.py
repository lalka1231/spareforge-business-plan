#!/usr/bin/env python3
"""Build the defense deck while preserving the source Morph and animation XML."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree
import json, copy, re

ROOT=Path(__file__).resolve().parents[1]
SRC=ROOT/'presentations/AGROVECTOR_PHOTO_MORPH.pptx'
OUT=ROOT/'presentations/AGROVECTOR_AFTER_brainstorm.pptx'
NS={'a':'http://schemas.openxmlformats.org/drawingml/2006/main','p':'http://schemas.openxmlformats.org/presentationml/2006/main'}
REVISION_DATE='18.11.2026'

# Existing shape IDs preserve animation targets.
REPL={
 (2,8):'Правовая проверка к 01.11: выполнение не подтверждено.',
 (2,10):'20 интервью к 29.11 · 10 клиентов к 21.03',
 (2,12):'Пилот 100 га к 30.04, в объеме сезона.',
 (3,8):'Один паспорт поля → операция → карта и акт',
 (3,10):'Цель: отчет до 48 ч; внесение от 100 га-проходов.',
 (4,52):'6 000 га-проходов;\n4 000 га-обследований;\n10 клиентов; повторы;\nотчет ≤48 ч; вклад/касса.',
 (4,55):'Почему мы: один ответственный за карту, операцию и акт.\nМеньше согласований двух подрядчиков; без покупки своего БВС.\nПроверяем на платном пилоте.',
 (4,58):'Потенциал: карты полей и повторные договоры. Проверим после сезона.',
 (4,67):'CAPEX 7,0 млн: борта 4,0; транспорт 1,4; монитор 0,6; безопасность 0,4; IT 0,2; узлы 0,4.\nOPEX 2027: команда 4,687; прочий фикс. 1,2; перем. 2,64 млн ₽ (400/60 ₽ за га).',
 (5,20):'Вся Россия → доступный рынок → план первого сезона',
 (5,24):'TAM / все посевы России',
 (5,25):'164,543 млрд ₽/год\n\n79,298 млн га × 2 075 ₽/га\n\nВсе категории хозяйств РФ; расчетная корзина двух услуг.',
 (5,27):'SAM / Россия, пригодные поля',
 (5,28):'24,682 млрд ₽/год\n\n79,298 млн га × 15% × 2 075 ₽/га\n\n15% — сценарий пригодности, не число действующих клиентов.',
 (5,31):'8,8 млн ₽ / 2027\n\n6 000 га-проходов × 1 300 ₽ + 4 000 га-обследований × 250 ₽\n\nЦель одной бригады: 10 клиентов.',
 (5,32):'Локальный кластер Татарстана — 124,5 млн ₽ (сценарий 200 × 2 000 га × 15%). SOM/SAM России — 0,036%; к кластеру — 7,07%. Площадь: ЕЭК, 2025; 15% и кратность — сценарий.',
 (6,16):'АГРОВЕКТОР / 06 / АНАЛИЗ КОНКУРЕНТОВ',
 (6,17):'Примерные цены аналогичных компаний',
 (6,21):'Выручка БФО — все направления юрлица, не доля рынка и не оценка компании. Публичные ориентиры внесения: от 900–1 000 ₽/га; состав, логистика и НДС различаются. Источники — в примечаниях.',
 (7,20):'SWOT-анализ и четыре стратегии',
 (7,24):'Кадровый план: приглашать ветеранов СВО, включая Героев России, с опытом БВС; оценивать навыки и психологическую устойчивость, обучать гражданским правилам. Кадровый канал: [46].',
 (9,8):'ООО · ОКВЭД 01.61 · капитал 10 тыс. ₽.',
 (9,10):'Капитал входит во взнос учредителя 2,5 млн ₽.',
 (9,12):'Перед выездом: БВС, поле, препарат и персонал.',
 (10,25):'Регулятор и страховщик: заранее сверить применимые документы, уведомления и покрытие; повторить при изменениях.\nДиректор/юрист — до допуска и при инциденте.',
 (10,28):'Инвестор и ключевой клиент: бюджет, условия транша, паспорт поля, календарь, акт и платеж.\nДиректор — еженедельно и перед решением.',
 (10,31):'Хозяйства вне маршрута: отслеживать спрос без обещаний и персональных данных.\nДиректор — ежемесячно; пересмотр при заявке.',
 (10,34):'Операторы и агроном: задание, ограничения и стоп-сигналы. Соседи и пчеловоды: применимые уведомления.\nОпер. руководитель — перед выездом.',
 (11,8):'Пилот 100 га в апреле включен в 6 000.',
 (11,10):'6 000/130 = 46,15 дня; 4 000/500 = 8 дней.',
 (11,12):'Итого 54,15 эффективного дня одной бригады.',
 (12,20):'Годовая стоимость команды: 4,69 млн ₽',
 (12,26):'98 479 ₽/мес. в Татарстане за январь–июнь 2026.\nПлан на 2027: по 100 000 ₽ в месяц до НДФЛ.',
 (12,28):'СТРАХОВЫЕ ВЗНОСЫ',
 (12,29):'В расчете: общие взносы — 30%.\nТравматизм — 0,2% до определения класса риска.',
 (12,30):'Источник средней по региону: республиканская комиссия, выпуск № 8(240).',
 (13,69):'20 дочерних работ внутри 13 пакетов project_plan.csv. Дерево показывает состав; сроки и зависимости — на слайдах 14–15.',
 (14,29):'A  Правовая проверка',
 (14,55):'Недели от старта · выполнение этапа A к 1 ноября не подтверждено.',
 (15,62):'Исходная последовательность работ. Сроки уточняются по готовности.',
 (16,17):'Матрица ответственности по 13 работам',
 (16,21):'R — выполняет    A — утверждает    C — консультирует    I — информируется',
 (16,22):'Пустая ячейка: роль в работе не участвует. Опер. рук. — первый оператор.',
 (17,50):'Выручка и денежный поток в 2027–2031 годах',
 (17,55):'Выручка: 8,800\nПеременные: 2,640\nФикс. OPEX: 5,887\nEBITDA: 0,273\nFCF: −0,255\nФиксировано: 490 600 ₽/мес.',
 (17,57):'Резерв 2 млн ₽ без выручки: касса уходит в минус в 5-м месяце.',
 (18,23):'Базовый сценарий: NPV −1,07 млн ₽',
 (18,28):'NPV базы: −1,068 млн ₽\nОкупаемость: 3,68 года;\nдисконтир.: не достигнута\nза 5 лет.\nЦена внесения: 1 450 ₽/га-проход\n20 дней: NPV +1,007 млн ₽\n16 дней: NPV −1,516 млн ₽\nЭто сценарии, не новая база.',
 (19,8):'Бюджет: оснащение 7,0 + запуск 0,5 + резерв 2,0 млн ₽.',
 (19,12):'Дефицит при отсрочке: 1 мес. — 0,42 млн ₽, 2 мес. — 1,57 млн ₽.',
 (20,8):'План вкладов: учредитель 2,5 + инвестор 7,0 млн ₽.',
 (20,12):'Аванс 30% и отсрочка на месяц: остаток от 21 760 ₽.',
 (22,6):'Сначала проверки,\nзатем решение о закупке',
 (22,8):'Проверить допуск, спрос и цены техники.',
 (22,10):'Сверить команду, загрузку и кассу.',
 (22,12):'По итогам пилота пересчитать экономику.',
}
NOTES={
 1:'Проект готовит ограниченные проверки спроса, экономики и допуска. Решение о закупке принимается после подтверждения условий запуска и результатов пилота.',
 11:'Статус: GO только на кабинетную проверку мощности; NO-GO на закупку до подтверждения выполнимости. Одна основная/резервная техника не образует две бригады. 20 дней — ресурсный экран, 16 дней — стресс существующего расчета; эффективные дни и нормы 130/500 га в день не подтверждены натурным замером.',
 17:'Статус: GO только на проверку фиксированных затрат и кассовой модели; NO-GO на закупку до launch gates. Расчетный фиксированный расход 490 600 ₽/мес. Первый отрицательный остаток в 5-м месяце — отдельный стресс без выручки при резерве 2 млн ₽, а не базовая касса и не утверждение о реально доступном резерве. Сохранены base тарифы 1 300/250 ₽, gross зарплаты и базовый NPV; базовый FCF/EBITDA не пересчитан.',
 18:'Статус: GO только на проверку цены/ресурса; NO-GO на закупку до подтверждения gates. Существующий расчет: совместный сценарий 1 450 ₽ за внесение + 20 дней даёт NPV +1 006 843 ₽; 16 дней при той же цене — −1 516 180 ₽. Сценарии не подтверждают спрос или фактическую производительность. Базовый NPV не менялся.',
 19:'Статус: GO только на сбор сопоставимых полных КП и платежного календаря; NO-GO на закупку до подтверждения бюджета и ликвидности. При задержке оплаты 1 месяц без аванса минимум −417 800 ₽; 2 месяца −1 568 400 ₽ (сценарии существующего расчета).',
 20:'Статус: GO только на согласование условий финансирования/оплаты; NO-GO на закупку до закрытия gates. При лаге 1 месяц аванс 30% даёт лишь 21 760 ₽ месячного минимума — это не предсезонный аванс и не безопасный буфер. Источники средств не подтверждены.',
 22:'Статус решения: GO только на обратимые проверки; NO-GO на безотзывную закупку и коммерческие обещания до закрытия применимых gates. Проверить: конкретный правовой/безопасностный маршрут, оплачиваемый спрос и условия, сопоставимую комплектацию/сервис и людей, помесячную мощность и кассу/буфер. Внешние риски остаются незакрытыми; это не заявление, что реальные риски устранены. После проверок — отдельное решение по пилоту/закупке.',
}

def slide_xml(z,n): return etree.fromstring(z.read(f'ppt/slides/slide{n}.xml'))
def texts(root): return [''.join(x.itertext()) for x in root.xpath('//a:p',namespaces=NS)]
def set_shape(root,sid,value):
    replace_text(find_shape(root,sid).find('p:txBody',NS),value)


def find_shape(root,sid):
    shapes=root.xpath(f'//p:sp[p:nvSpPr/p:cNvPr[@id="{sid}"]]',namespaces=NS)
    assert len(shapes)==1,(sid,len(shapes))
    return shapes[0]


def replace_text(body,value):
    paras=body.findall('a:p',NS)
    # Preserve first run properties and paragraph properties; only replace textual runs.
    lines=value.split('\n')
    if len(paras)<len(lines):
        for _ in range(len(lines)-len(paras)):
            p=etree.Element('{%s}p'%NS['a']); etree.SubElement(p,'{%s}endParaRPr'%NS['a']); body.append(p)
        paras=body.findall('a:p',NS)
    style=body.xpath('./a:p/a:pPr',namespaces=NS)
    style=copy.deepcopy(style[0]) if style else None
    template=body.xpath('./a:p/a:r/a:rPr|./a:p/a:pPr/a:defRPr',namespaces=NS)
    template=copy.deepcopy(template[0]) if template else etree.Element('{%s}rPr'%NS['a'])
    template.tag='{%s}rPr'%NS['a']
    for idx,p in enumerate(paras):
        if p.find('{%s}pPr'%NS['a']) is None and style is not None:p.insert(0,copy.deepcopy(style))
        for r in list(p.xpath('./a:r|./a:fld',namespaces=NS)): p.remove(r)
        if idx<len(lines) and lines[idx]:
            r=etree.Element('{%s}r'%NS['a']); rp=copy.deepcopy(template)
            r.append(rp)
            tx=etree.SubElement(r,'{%s}t'%NS['a']); tx.text=lines[idx]
            end=p.find('{%s}endParaRPr'%NS['a'])
            p.insert(list(p).index(end) if end is not None else len(p),r)


def set_shape_font(root,sid,points):
    shape=find_shape(root,sid)
    for props in shape.xpath('.//a:rPr|.//a:defRPr',namespaces=NS):
        props.set('sz',str(round(points*100)))


def set_box(shape,x,y,w,h):
    transform=shape.find('p:spPr/a:xfrm',NS)
    off=transform.find('a:off',NS); ext=transform.find('a:ext',NS)
    off.set('x',str(round(x*914400)));off.set('y',str(round(y*914400)))
    ext.set('cx',str(round(w*914400)));ext.set('cy',str(round(h*914400)))


def update_swot_table(root):
    table=root.find('.//a:tbl',NS)
    cells=[
        ['Внутренние / внешние','O — возможности\nСложные поля; платные пилоты; кадры с опытом БВС','T — угрозы\nОграничения; погода и отказы; давление цен и кассы'],
        ['S — сильные стороны\nЕдиный заказ и акт; прозрачные тарифы; резервный борт; план привлечения ветеранов СВО, включая Героев России',
         'SO: платный пилот с картой и приемкой; привлечь подготовленных операторов БВС.\nДиректор + агроном; критерий — договор и оплата.',
         'ST: перед выездом сверить допуск, погоду и технику; отработать решения под нагрузкой.\nОпер. руководитель; критерий — чек-лист.'],
        ['W — слабые стороны\nНет оплаченного сезона; одна бригада; техника и кандидаты на проверке',
         'WO: собрать КП и проверить кандидатов; провести пилот.\nДиректор; критерий — КП, экзамен, акт и платеж.',
         'WT: лимитировать месячную загрузку и кассовый риск.\nДиректор + оператор; критерий — нет двойного бронирования.']
    ]
    for row,values in zip(table.findall('a:tr',NS),cells):
        for cell,value in zip(row.findall('a:tc',NS),values):
            body=cell.find('a:txBody',NS);replace_text(body,value)
            for props in body.xpath('.//a:rPr|.//a:defRPr',namespaces=NS):props.set('sz','1250')


def update_wbs(root):
    data=json.loads((ROOT/'data/defense_wbs.json').read_text())
    # The 13 original cards keep their IDs and animation targets. Seven extra
    # visual cards expand the hierarchy without altering the project network.
    column_pairs=[[(31,32),(33,34),(35,36)],[(41,42),(43,44),(45,46),(47,48)],
                  [(53,54),(55,56),(57,58),(59,60)],[(65,66),(67,68)]]
    xs=[0.78,3.85,6.93,10.01]
    tree=root.find('.//p:spTree',NS)
    next_id=max(int(x) for x in root.xpath('//p:cNvPr/@id',namespaces=NS))+1
    for col,phase in enumerate(data['phases']):
        template_bg=find_shape(root,column_pairs[col][0][0])
        template_text=find_shape(root,column_pairs[col][0][1])
        for row,(ident,label,parent) in enumerate(phase['tasks']):
            if row<len(column_pairs[col]):
                bg=find_shape(root,column_pairs[col][row][0])
                tx=find_shape(root,column_pairs[col][row][1])
            else:
                bg=copy.deepcopy(template_bg);tx=copy.deepcopy(template_text)
                for shape in (bg,tx):
                    c=shape.find('p:nvSpPr/p:cNvPr',NS)
                    c.set('id',str(next_id));c.set('name',f'WBS {ident} {next_id}')
                    next_id+=1
                    tree.append(shape)
            y=3.52+row*.47
            set_box(bg,xs[col],y,2.63,.40)
            set_box(tx,xs[col]+.10,y+.06,2.43,.28)
            replace_text(tx.find('p:txBody',NS),f'{ident}  {label}')
            for props in tx.xpath('.//a:rPr|.//a:defRPr',namespaces=NS):props.set('sz','1320')


def add_bvs_source(root,entries):
    with ZipFile(SRC) as z: donor=slide_xml(z,7)
    shape=copy.deepcopy(find_shape(donor,24))
    shape.find('p:nvSpPr/p:cNvPr',NS).set('id','21')
    shape.find('p:nvSpPr/p:cNvPr',NS).set('name','BVS official source')
    set_box(shape,.66,6.64,11.8,.30)
    replace_text(shape.find('p:txBody',NS),
                 'БВС — беспилотное воздушное судно. Росавиация: favt.gov.ru/dejatelnost-gosudarstvennaa-registracia-vs-i-uchet-bvs-svs/')
    for props in shape.xpath('.//a:rPr|.//a:defRPr',namespaces=NS):props.set('sz','1050')
    rel_path='ppt/slides/_rels/slide8.xml.rels'
    rels=etree.fromstring(entries[rel_path])
    ids={r.get('Id') for r in rels};number=1
    while f'rId{number}' in ids:number+=1
    rid=f'rId{number}'
    relns='http://schemas.openxmlformats.org/package/2006/relationships'
    etree.SubElement(rels,f'{{{relns}}}Relationship',Id=rid,
                     Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
                     Target='https://www.favt.gov.ru/dejatelnost-gosudarstvennaa-registracia-vs-i-uchet-bvs-svs/',
                     TargetMode='External')
    rns='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    for props in shape.xpath('.//a:rPr',namespaces=NS):
        props.set('u','sng')
        for old in props.findall(f'{{{NS["a"]}}}solidFill'):props.remove(old)
        fill=etree.Element(f'{{{NS["a"]}}}solidFill')
        etree.SubElement(fill,f'{{{NS["a"]}}}srgbClr',val='D9ED85')
        props.insert(0,fill)
    etree.SubElement(shape.find('p:nvSpPr/p:cNvPr',NS),
                     f'{{{NS["a"]}}}hlinkClick',attrib={f'{{{rns}}}id':rid})
    root.find('.//p:spTree',NS).append(shape)
    entries[rel_path]=etree.tostring(rels,xml_declaration=True,encoding='UTF-8',standalone=True)


def update_competitor_table(root):
    comparison=json.loads((ROOT/'data/defense_competitors.json').read_text())
    table=root.find('.//a:tbl',NS)
    grid=table.find('a:tblGrid',NS)
    for column in list(grid):
        grid.remove(column)
    for inches in (2.01,1.80,2.13,2.47,2.22,1.47):
        etree.SubElement(grid,f"{{{NS['a']}}}gridCol",w=str(round(inches*914400)))
    templates=table.findall('a:tr',NS)
    for row in templates:
        table.remove(row)
    for index,values in enumerate([comparison['headers'],*comparison['rows']]):
        template=templates[0 if index==0 else 1+(index-1)%2]
        row=copy.deepcopy(template)
        cell_template=row.find('a:tc',NS)
        for cell in row.findall('a:tc',NS):
            row.remove(cell)
        row.set('h',str(round((.43 if index==0 else .49)*914400)))
        for value in values:
            cell=copy.deepcopy(cell_template)
            body=cell.find('a:txBody',NS)
            replace_text(body,value)
            for props in body.xpath('.//a:rPr|.//a:defRPr',namespaces=NS):
                props.set('sz','980' if index else '1100')
            cell.find('a:tcPr',NS).set('marT','18000')
            row.append(cell)
        table.append(row)
    frame=table.xpath('ancestor::p:graphicFrame',namespaces=NS)[0]
    frame.find('p:xfrm/a:ext',NS).set('cy',str(round((.43+len(comparison['rows'])*.49)*914400)))


def style_teaching_tables(number,root):
    if number==8:
        for row in root.findall('.//a:tbl/a:tr',NS)[1:]:
            cell=row.findall('a:tc',NS)[1]
            level=''.join(cell.xpath('.//a:t/text()',namespaces=NS))
            fill=cell.find('a:tcPr/a:solidFill/a:srgbClr',NS)
            fill.set('val',{'Высокий':'793D3D','Средний–высокий':'785530','Средний':'665625'}[level])
    if number==16:
        for row in root.findall('.//a:tbl/a:tr',NS)[1:]:
            for cell in row.findall('a:tc',NS)[1:]:
                body=cell.find('a:txBody',NS)
                for props in body.findall('a:p/a:pPr',NS):
                    props.set('algn','ctr')
                if ''.join(cell.xpath('.//a:t/text()',namespaces=NS))=='—':
                    replace_text(body,'')


def adjust_label_positions(number,root):
    if number==14:
        find_shape(root,55).find('p:spPr/a:xfrm/a:off',NS).set('y',str(round(2.28*914400)))
    if number in (2,3,9,11,19,20,22):
        for sid in (8,10,12):
            position=find_shape(root,sid).find('p:spPr/a:xfrm/a:off',NS)
            position.set('y',str(int(position.get('y'))-45720))


def update_footer_and_markers(root):
    for node in root.findall('.//a:t',NS):
        value=node.text or ''
        value=re.sub(r'\[(?:ДОПУЩЕНИЕ(?:\s*/\s*РАСЧЕТ)?|РАСЧЕТ)\]\s*','',value)
        if value.startswith('Проектные гипотезы и расчеты;'):
            value=f'АгроВектор · Материалы к защите · {REVISION_DATE}'
        else:
            value=value.replace('Редакция 21.10.2026',f'Редакция {REVISION_DATE}')
        if value!=node.text:
            node.text=value

def notes_part(z,n):
    rel=etree.fromstring(z.read(f'ppt/slides/_rels/slide{n}.xml.rels'))
    for r in rel:
        if r.get('Type','').endswith('/notesSlide'):
            target=r.get('Target'); return 'ppt/'+target.replace('../','') if target.startswith('../') else 'ppt/slides/'+target
    raise AssertionError(f'no notes for {n}')

def source_fingerprints(z):
    out=[]
    for n in range(1,23):
        r=slide_xml(z,n)
        out.append((etree.tostring(r.xpath('//p:transition',namespaces=NS)[0]) if r.xpath('//p:transition',namespaces=NS) else b'',
                    etree.tostring(r.xpath('//p:timing',namespaces=NS)[0]) if r.xpath('//p:timing',namespaces=NS) else b''))
    return out

def build():
    assert SRC.exists()
    with ZipFile(SRC) as zin:
        original=source_fingerprints(zin)
        entries={name:zin.read(name) for name in zin.namelist()}
    for n in range(1,23):
        part=f'ppt/slides/slide{n}.xml'
        root=etree.fromstring(entries[part])
        for (slide,sid),value in REPL.items():
            if slide==n:
                set_shape(root,sid,value)
        if n==4:
            for sid,points in {52:11.2,55:13.0,67:12.4}.items():set_shape_font(root,sid,points)
            set_box(find_shape(root,67),.75,5.88,5.7,1.00)
        if n==5:
            for sid in (25,28,31):set_shape_font(root,sid,18.0)
            set_shape_font(root,32,13.3)
        if n==6:
            update_competitor_table(root)
            set_shape_font(root,21,12.0)
        if n==7:
            update_swot_table(root)
            set_shape_font(root,24,11.0)
        if n==8:add_bvs_source(root,entries)
        if n==9:
            for sid in (8,10,12):set_shape_font(root,sid,19.0)
        if n==10:
            for sid in (25,28,31,34):set_shape_font(root,sid,15.5)
        if n==11:
            for sid in (8,10,12):set_shape_font(root,sid,19.5)
        if n==13:update_wbs(root)
        style_teaching_tables(n,root)
        adjust_label_positions(n,root)
        update_footer_and_markers(root)
        entries[part]=etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
    full=(ROOT/'docs/FINAL_BUSINESS_PLAN_brainstorm.md').read_text()
    sections=re.findall(r'^## (\d+)\. ([\s\S]*?)(?=^## \d+\.|^## Sources|\Z)',full,re.M)
    assert len(sections)==22
    for number,section in sections:
        n=int(number);text=section+'\n\n'+NOTES.get(n,'')
        part=notes_part(ZipFile(SRC),n); root=etree.fromstring(entries[part])
        # append a separately formatted status paragraph to existing notes body
        body=root.xpath('//p:sp[p:nvSpPr/p:nvPr/p:ph[@type="body"]]/p:txBody',namespaces=NS)
        if not body: body=root.xpath('//p:sp/p:txBody',namespaces=NS)
        assert body
        for oldp in list(body[0].findall('{%s}p'%NS['a'])): body[0].remove(oldp)
        p=etree.SubElement(body[0],'{%s}p'%NS['a']); r=etree.SubElement(p,'{%s}r'%NS['a']); t=etree.SubElement(r,'{%s}t'%NS['a']); t.text='\n\n'+text
        entries[part]=etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
    # Repair inherited dark chart titles on the analytical slides only.
    cns='http://schemas.openxmlformats.org/drawingml/2006/chart'
    for n in (11,17,18,19,20,21):
        rel=etree.fromstring(entries[f'ppt/slides/_rels/slide{n}.xml.rels'])
        for relationship in rel:
            if not relationship.get('Type','').endswith('/chart'):continue
            path='ppt/'+relationship.get('Target').replace('../','')
            root=etree.fromstring(entries[path])
            chart=root.find('{%s}chart'%cns)
            if chart is not None and chart.find('{%s}title'%cns) is None:
                deleted=chart.find('{%s}autoTitleDeleted'%cns)
                if deleted is not None and deleted.get('val')=='0':chart.insert(0,etree.Element('{%s}title'%cns))
            if n==17 and chart.find(f'{{{cns}}}title') is None:
                chart.insert(0,etree.Element(f'{{{cns}}}title'))
            for title in root.findall('.//{%s}title'%cns):
                # LibreOffice ignores txPr for automatic chart titles; make the inherited label explicit.
                label='Финансовые показатели, млн ₽' if n==17 else (
                    root.xpath('string((//c:ser/c:tx//c:v)[1])',namespaces={**NS,'c':cns}) or 'NPV, млн ₽'
                )
                oldtx=title.find('{%s}tx'%cns)
                if oldtx is not None:title.remove(oldtx)
                tx=etree.Element('{%s}tx'%cns);rich=etree.SubElement(tx,'{%s}rich'%cns)
                etree.SubElement(rich,'{%s}bodyPr'%NS['a']);etree.SubElement(rich,'{%s}lstStyle'%NS['a'])
                p=etree.SubElement(rich,'{%s}p'%NS['a']);r=etree.SubElement(p,'{%s}r'%NS['a'])
                rp=etree.SubElement(r,'{%s}rPr'%NS['a'],sz='1400')
                fill=etree.SubElement(rp,'{%s}solidFill'%NS['a']);etree.SubElement(fill,'{%s}srgbClr'%NS['a'],val='FFFFFF')
                etree.SubElement(r,'{%s}t'%NS['a']).text=label
                title.insert(0,tx)
                props=title.find('{%s}txPr'%cns)
                if props is not None:title.remove(props)
                props=etree.SubElement(title,'{%s}txPr'%cns)
                etree.SubElement(props,'{%s}bodyPr'%NS['a']);etree.SubElement(props,'{%s}lstStyle'%NS['a'])
                p=etree.SubElement(props,'{%s}p'%NS['a']);pp=etree.SubElement(p,'{%s}pPr'%NS['a'])
                rp=etree.SubElement(pp,'{%s}defRPr'%NS['a'],sz='1400')
                fill=etree.SubElement(rp,'{%s}solidFill'%NS['a']);etree.SubElement(fill,'{%s}srgbClr'%NS['a'],val='FFFFFF')
                for rp in title.findall('.//{%s}rPr'%NS['a'])+title.findall('.//{%s}defRPr'%NS['a']):
                    for fill in list(rp.findall('{%s}solidFill'%NS['a'])):rp.remove(fill)
                    fill=etree.SubElement(rp,'{%s}solidFill'%NS['a']);etree.SubElement(fill,'{%s}srgbClr'%NS['a'],val='FFFFFF')
            if n==17:
                legend=etree.Element(f'{{{cns}}}legend')
                etree.SubElement(legend,f'{{{cns}}}legendPos',val='b')
                etree.SubElement(legend,f'{{{cns}}}overlay',val='0')
                legend.append(copy.deepcopy(chart.find(f'{{{cns}}}title/{{{cns}}}txPr')))
                chart.insert(chart.index(chart.find(f'{{{cns}}}plotArea'))+1,legend)
            if n==18:
                root.find(f'.//{{{cns}}}catAx/{{{cns}}}tickLblPos').set('val','low')
            entries[path]=etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
    OUT.parent.mkdir(parents=True,exist_ok=True)
    with ZipFile(OUT,'w',ZIP_DEFLATED) as zout:
        for name,data in entries.items(): zout.writestr(name,data)
    # Readback invariants
    with ZipFile(OUT) as z:
        assert len([x for x in z.namelist() if x.startswith('ppt/slides/slide') and x.endswith('.xml') and '/_rels/' not in x])==22
        assert source_fingerprints(z)==original, 'transition/timing changed'
        for (n,sid),value in REPL.items():
            root=slide_xml(z,n); alltext='\n'.join(texts(root)); assert value.replace('\n','')[:12] in alltext.replace('\n','') or value.split('\n')[0] in alltext
        for n,text in NOTES.items(): assert text in ''.join(etree.fromstring(z.read(notes_part(z,n))).itertext())
    print(json.dumps({'output':str(OUT),'slides':22,'content_asserts':len(REPL),'notes_statuses':len(NOTES),'timing_transition_preserved_all':True},ensure_ascii=False))

if __name__=='__main__': build()
