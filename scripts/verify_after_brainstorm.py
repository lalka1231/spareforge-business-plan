"""Read-only acceptance of the after-brainstorm deliverables."""
from pathlib import Path
from zipfile import ZipFile
import hashlib,json,re,subprocess
from lxml import etree as E
from docx import Document
from pptx import Presentation
import pymupdf
from build_finance import model,metrics
ROOT=Path(__file__).resolve().parents[1]
PARENT='428ed6b1ea89876a6a910299f843efaa8507062e'
NS={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}

def verify():
    changed=subprocess.check_output(['git','diff',PARENT,'--name-only'],cwd=ROOT,text=True).splitlines()
    prior=set(subprocess.check_output(['git','ls-tree','-r','--name-only',PARENT],cwd=ROOT,text=True).splitlines())
    changed_snapshots=prior.intersection(changed)-{'README.md'}
    assert not changed_snapshots,'Historical files changed: '+str(changed_snapshots)
    assert 'README_AFTER_brainstorm.md' in (ROOT/'README.md').read_text().split('##')[1]
    current=json.loads((ROOT/'data/finance_corrected_brainstorm.json').read_text())
    baseline=json.loads((ROOT/'data/finance_summary.json').read_text())
    assert current['base']==metrics(model())==baseline['base']
    assert current['yearly_scenarios']['base']==baseline['yearly_scenarios']['base']
    cash=current['cash_runway'];reserve=model()[0]['reserve_funding_rub'];burn=model()[1]['fixed_opex_rub']/12
    assert burn==cash['no_revenue_fixed_burn_per_month_rub']==490600
    first=next(i for i in range(1,120) if reserve-i*burn<0)
    assert first==cash['no_revenue_first_negative_month']==5
    assert reserve-4*burn>=0 and reserve-5*burn<0
    md=(ROOT/'docs/FINAL_BUSINESS_PLAN_brainstorm.md').read_text()
    sections=re.findall(r'^## (\d+)\. ([\s\S]*?)(?=^## \d+\.|^## Sources|\Z)',md,re.M)
    assert [int(x[0]) for x in sections]==list(range(1,23))
    assert '{fmt(' not in md and 'Не достигнута года' not in md
    assert '79 298 000 га' in md and '164 543 350 000' in md and '24 681 502 500' in md
    assert 'SAM России' in md and 'локальный кластер Татарстана' in md
    assert '6 000 га-проходов / 130' in md and '4 000 га-обследований / 500' in md
    market=json.loads((ROOT/'data/defense_market.json').read_text())
    unit=market['treatment_passes_per_ha_scenario']*market['treatment_price_rub_2027']+market['monitoring_surveys_per_ha_scenario']*market['monitoring_price_rub_2027']
    assert unit==2075
    assert market['tam_rub']==market['area_russia_ha_2025_rounded']*unit
    assert market['sam_rub']==market['tam_rub']*market['serviceable_share_scenario']
    assert market['local_tatarstan_cluster_rub']==200*2000*market['serviceable_share_scenario']*unit
    assert market['som_rub_2027']==6000*1300+4000*250==8800000
    assert abs((6000/130+4000/500)-54.15384615384615)<1e-8
    wbs=json.loads((ROOT/'data/defense_wbs.json').read_text())
    assert sum(len(p['tasks']) for p in wbs['phases'])==20
    assert 'на 18.11.2026 выполнение не подтверждено' in md
    doc=Document(str(ROOT/'documents/AGROVECTOR_FULL_BUSINESS_PLAN_brainstorm.docx'))
    dt='\n'.join(p.text for p in doc.paragraphs)+'\n'+'\n'.join(c.text for t in doc.tables for r in t.rows for c in r.cells)
    numbered=[p.text for p in doc.paragraphs if p.style and p.style.name=='Heading 1' and re.match(r'^\d+\.',p.text)]
    assert len(numbered)==22 and len(doc.inline_shapes)==2 and '{fmt(' not in dt
    comparison=json.loads((ROOT/'data/defense_competitors.json').read_text())
    matrices=[t for t in doc.tables if [c.text for c in t.rows[0].cells]==comparison['headers']]
    assert len(matrices)==1
    normalize=lambda line:' '.join(line.replace(';',' ').split())
    assert [[normalize(c.text) for c in r.cells] for r in matrices[0].rows[1:]]==[
        [normalize(c) for c in r] for r in comparison['rows']
    ]
    ppt=Presentation(str(ROOT/'presentations/AGROVECTOR_AFTER_brainstorm.pptx'));assert len(ppt.slides)==22
    market_slide=' '.join(s.text for s in ppt.slides[4].shapes if s.has_text_frame)
    assert all(value in market_slide for value in ('164,543','24,682','8,8 млн','124,5 млн'))
    assert 'не подписанные заказы' not in market_slide
    for (number,section),slide in zip(sections,ppt.slides):
        note=slide.notes_slide.notes_text_frame.text
        assert section in note,(number,'Notes are not synchronized to revised report')
        assert '{fmt(' not in note
        assert not any(re.search(r'допущени[ея]',s.text,re.I) for s in slide.shapes if s.has_text_frame)
    comp_slide=[s.table for s in ppt.slides[5].shapes if s.has_table]
    assert len(comp_slide)==1
    assert [[c.text for c in row.cells] for row in comp_slide[0].rows]==[
        comparison['headers'],*comparison['rows']
    ]
    wbs_text=' '.join(s.text for s in ppt.slides[12].shapes if s.has_text_frame)
    assert all(task[0] in wbs_text for phase in wbs['phases'] for task in phase['tasks'])
    raci=next(s.table for s in ppt.slides[15].shapes if s.has_table)
    for row in list(raci.rows)[1:]:
        roles=[c.text.split('/') for c in list(row.cells)[1:]]
        assert sum('A' in r for r in roles)==1 and any('R' in r for r in roles)
    assert any(not c.text for row in raci.rows for c in list(row.cells)[1:])
    for marker in ('490 600','21 760','AgroVector_After_brainstorm'):
        assert marker in md and marker in dt,marker
    annual=current['yearly_scenarios']['base'][1:]
    finance_charts=[s.chart for s in ppt.slides[16].shapes if s.has_chart]
    assert len(finance_charts)==1 and finance_charts[0].has_legend
    for ser,key in zip(finance_charts[0].series,('revenue_rub','ebitda_rub','free_cash_flow_rub')):
        assert all(abs(a-b[key]/1e6)<1e-8 for a,b in zip(ser.values,annual))
    stress_charts=[s.chart for s in ppt.slides[17].shapes if s.has_chart]
    assert len(stress_charts)==1
    assert all(abs(a-b['npv_rub']/1e6)<1e-8 for a,b in zip(stress_charts[0].series[0].values,baseline['sensitivity']))
    chart_count=0
    with ZipFile(ROOT/'presentations/AGROVECTOR_PHOTO_MORPH.pptx') as old,ZipFile(ROOT/'presentations/AGROVECTOR_AFTER_brainstorm.pptx') as new:
        assert new.testzip() is None
        for i in range(1,23):
            path=f'ppt/slides/slide{i}.xml';a=E.fromstring(old.read(path));b=E.fromstring(new.read(path))
            if i<=2:
                assert [E.tostring(x) for x in a.findall('.//p:pic',NS)]==[E.tostring(x) for x in b.findall('.//p:pic',NS)]
            for tag in ('transition','timing'):
                assert [E.tostring(x) for x in a.xpath('//p:'+tag,namespaces=NS)]==[E.tostring(x) for x in b.xpath('//p:'+tag,namespaces=NS)],(i,tag)
        for path in old.namelist():
            if path.startswith('ppt/media/'):assert old.read(path)==new.read(path),path
            if re.fullmatch(r'ppt/charts/chart\d+\.xml',path):
                a=E.fromstring(old.read(path));b=E.fromstring(new.read(path))
                assert [E.tostring(x,method='c14n') for x in a.xpath('//c:ser',namespaces=NS)]==[E.tostring(x,method='c14n') for x in b.xpath('//c:ser',namespaces=NS)],path
                chart_count+=1
    pdf_counts={}
    for relative,count in [('presentations/AGROVECTOR_AFTER_brainstorm.pdf',22),('documents/AGROVECTOR_FULL_BUSINESS_PLAN_brainstorm.pdf',None)]:
        pdf=pymupdf.open(ROOT/relative);assert len(pdf)>0
        if count:assert len(pdf)==count
        text='\n'.join(p.get_text() for p in pdf);assert '{fmt(' not in text
        assert '490600' in re.sub(r'\s+','',text)
        pdf_counts[relative]=len(pdf)
    slide_pdf=pymupdf.open(ROOT/'presentations/AGROVECTOR_AFTER_brainstorm.pdf')
    assert any('favt.gov.ru' in x.get('uri','') for x in slide_pdf[7].get_links())
    assets=['docs/FINAL_BUSINESS_PLAN_brainstorm.md','data/finance_corrected_brainstorm.json','documents/AGROVECTOR_FULL_BUSINESS_PLAN_brainstorm.docx','presentations/AGROVECTOR_AFTER_brainstorm.pptx',*pdf_counts]
    for name in assets:assert (ROOT/name).stat().st_size<100*1024**2
    return {'verified':True,'parent_commit':PARENT,'historical_snapshots_unchanged':True,'sections':22,'slides':22,'notes_synchronized':22,'timing_transition_preserved':22,'first_two_slide_images_preserved':True,'chart_series_unchanged':chart_count,'word_tables':len(doc.tables),'word_diagrams':len(doc.inline_shapes),'competitor_rows_synchronized':len(comparison['rows']),'pdf_pages':pdf_counts,'npv_rub':current['base']['npv_rub'],'fixed_monthly_rub':burn,'first_negative_no_revenue_month':first,'hashes':{name:hashlib.sha256((ROOT/name).read_bytes()).hexdigest() for name in assets},'limits':['Static LibreOffice render; no new native PowerPoint slideshow playback test','External legal, tax, market, staffing, supplier, operational and financing gates remain open']}

if __name__=='__main__':print(json.dumps(verify(),ensure_ascii=False,indent=2))
