#!/usr/bin/env python3
"""Create a content-only brainstorm sibling deck; preserve all source slide XML effects."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree
import json, shutil, tempfile, copy, re

ROOT=Path(__file__).resolve().parents[1]
SRC=ROOT/'presentations/AGROVECTOR_PHOTO_MORPH.pptx'
OUT=ROOT/'presentations/AGROVECTOR_AFTER_brainstorm.pptx'
NS={'a':'http://schemas.openxmlformats.org/drawingml/2006/main','p':'http://schemas.openxmlformats.org/presentationml/2006/main'}

# Existing shape IDs chosen from the source deck; no geometry/formatting changes.
REPL={
 (11,10):'[РАСЧЕТ] 2027: 54,15 полевых дня · одна бригада',
 (11,12):'20 эффективных дней — ресурсный экран; 16 — жесткий стресс. Не замер.',
 (17,50):'База: прежние тарифы; новые стресс-сценарии не меняют base NPV',
 (17,55):'Выручка: 8,800\nПеременные: 2,640\nФикс. OPEX: 5,887\nEBITDA: 0,273\nFCF: -0,255\nФикс. расход: 490 600 ₽/мес.\nБез выручки: минус с 5-го мес.¹',
 (17,57):'Выручка / EBITDA / FCF; ¹ стресс без выручки, резерв 2 млн ₽.',
 (18,23):'Базовый NPV сохранен; устойчивость зависит от ресурса одной бригады',
 (18,28):'NPV базы: −1,068 млн ₽\nОкупаемость: 3,68 года;\nдисконтир.: не достигнута\nза 5 лет.\nЦена внесения: 1 450 ₽/га-проход\n20 дней: NPV +1,007 млн ₽\n16 дней: NPV −1,516 млн ₽\nЭто сценарии, не новая база.',
 (19,12):'Кассовый стресс: лаг 1 мес. −417 800 ₽; 2 мес. −1 568 400 ₽.',
 (20,12):'Лаг 1 мес. + аванс 30%: минимум лишь 21 760 ₽; буфера нет.',
 (22,6):'Проверки — GO.\nЗакупка — после gates.',
 (22,8):'Проверить: допуск, спрос, КП и команду, мощность и кассу.',
 (22,10):'Внешние проверки открыты. До их закрытия — NO-GO на закупку.',
}
NOTES={
 1:'Статус: GO только на обратимые проверки; NO-GO на безотзывную закупку до закрытия launch gates. Исходный визуальный слайд оставлен без изменений; статус — только в заметках. Это логика анализа, не утверждение о фактическом снятии рисков.',
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
    shape=root.xpath(f'//p:sp[p:nvSpPr/p:cNvPr[@id="{sid}"]]',namespaces=NS)
    assert len(shape)==1,(sid,len(shape))
    paras=shape[0].xpath('./p:txBody/a:p',namespaces=NS)
    # Preserve first run properties and paragraph properties; only replace textual runs.
    lines=value.split('\n')
    if len(paras)<len(lines):
        body=shape[0].xpath('./p:txBody',namespaces=NS)[0]
        for _ in range(len(lines)-len(paras)):
            p=etree.Element('{%s}p'%NS['a']); etree.SubElement(p,'{%s}endParaRPr'%NS['a']); body.append(p)
        paras=shape[0].xpath('./p:txBody/a:p',namespaces=NS)
    style=shape[0].xpath('./p:txBody/a:p/a:pPr',namespaces=NS)
    style=copy.deepcopy(style[0]) if style else None
    template=shape[0].xpath('./p:txBody/a:p/a:r/a:rPr|./p:txBody/a:p/a:pPr/a:defRPr',namespaces=NS)
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
    # Validate visual invariance of title slides by preserving slide XML exactly.
    before=[slide_xml(ZipFile(SRC),n) for n in (1,2)]
    for (n,sid),value in REPL.items():
        root=etree.fromstring(entries[f'ppt/slides/slide{n}.xml']); set_shape(root,sid,value)
        entries[f'ppt/slides/slide{n}.xml']=etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
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
    for n in (17,18,19,20,21):
        rel=etree.fromstring(entries[f'ppt/slides/_rels/slide{n}.xml.rels'])
        for relationship in rel:
            if not relationship.get('Type','').endswith('/chart'):continue
            path='ppt/'+relationship.get('Target').replace('../','')
            root=etree.fromstring(entries[path])
            chart=root.find('{%s}chart'%cns)
            if chart is not None and chart.find('{%s}title'%cns) is None:
                deleted=chart.find('{%s}autoTitleDeleted'%cns)
                if deleted is not None and deleted.get('val')=='0':chart.insert(0,etree.Element('{%s}title'%cns))
            for title in root.findall('.//{%s}title'%cns):
                # LibreOffice ignores txPr for automatic chart titles; make the inherited label explicit.
                label=root.xpath('string((//c:ser/c:tx//c:v)[1])',namespaces={**NS,'c':cns}) or 'NPV, млн ₽'
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
            entries[path]=etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
    OUT.parent.mkdir(parents=True,exist_ok=True)
    with ZipFile(OUT,'w',ZIP_DEFLATED) as zout:
        for name,data in entries.items(): zout.writestr(name,data)
    # Readback invariants
    with ZipFile(OUT) as z:
        assert len([x for x in z.namelist() if x.startswith('ppt/slides/slide') and x.endswith('.xml') and '/_rels/' not in x])==22
        assert source_fingerprints(z)==original, 'transition/timing changed'
        for n in (1,2): assert z.read(f'ppt/slides/slide{n}.xml')==ZipFile(SRC).read(f'ppt/slides/slide{n}.xml')
        for (n,sid),value in REPL.items():
            root=slide_xml(z,n); alltext='\n'.join(texts(root)); assert value.replace('\n','')[:12] in alltext.replace('\n','') or value.split('\n')[0] in alltext
        for n,text in NOTES.items(): assert text in ''.join(etree.fromstring(z.read(notes_part(z,n))).itertext())
    print(json.dumps({'output':str(OUT),'slides':22,'content_asserts':len(REPL),'notes_statuses':len(NOTES),'timing_transition_preserved_all':True,'first_two_slide_xml_identical':True},ensure_ascii=False))

if __name__=='__main__': build()
