#!/usr/bin/env python3
"""Read-only acceptance of brainstorm artifacts and unchanged approved baseline."""
import csv
import hashlib
import importlib.util
import json
import math
from pathlib import Path
import re
import subprocess
import sys

ROOT=Path(__file__).resolve().parents[1]
BASE='7c05732803c7bedff0ce15f35647195e61de0790'


def load(path):
    return json.loads((ROOT/path).read_text())


def csv_rows(path):
    with (ROOT/path).open(encoding='utf-8-sig',newline='') as stream:
        return list(csv.DictReader(stream))


def verify():
    baseline_paths=subprocess.check_output(['git','ls-tree','-r','--name-only',BASE],cwd=ROOT).decode().splitlines()
    # This compares content with the approved commit, not merely a later snapshot.
    modified=subprocess.check_output(['git','diff','--name-only',BASE,'--'],cwd=ROOT).decode().splitlines()
    assert all(Path(p).stem.endswith('_brainstorm') for p in modified),modified
    assert all((ROOT/p).exists() for p in baseline_paths),'Deleted baseline file'
    results=load('data/resilience_brainstorm.json')
    hashes=results['verification']['source_sha256']
    assert all(hashlib.sha256((ROOT/p).read_bytes()).hexdigest()==sha for p,sha in hashes.items())
    registers=['finance','market','operations','readiness']
    domain_risks=[]
    for domain in registers:
        domain_risks.extend(load(f'data/{domain}_risks_brainstorm.json'))
    risks=load('data/risk_register_brainstorm.json')
    ids={r['id'] for r in risks}
    assert len(ids)==len(risks)==len(domain_risks)
    assert ids=={r['id'] for r in domain_risks}
    assert len(csv_rows('data/risk_register_brainstorm.csv'))==len(risks)
    for r in risks:
        assert len(r['alternatives'])>=2 and r['experiment'] and r['stop_rule'] and r['owner'],r['id']
        assert all((ROOT/p).exists() for p in r['evidence']),r['id']
    coverage=load('data/coverage_brainstorm.json')
    assert len(coverage)==22 and {c['section'] for c in coverage}==set(range(1,23))
    assert all(c['risk_ids'] and set(c['risk_ids'])<=ids and not c['baseline_changed'] for c in coverage)
    assert len(results['scenarios'])==len(csv_rows('data/scenario_comparison_brainstorm.csv'))==23
    annual=csv_rows('data/scenario_annual_brainstorm.csv')
    assert len(annual)==23*6
    for s in results['scenarios']:
        rows=[r for r in annual if r['scenario']==s['scenario']]
        assert len(rows)==6
        independent_npv=sum(float(r['free_cash_flow_rub'])/1.24**int(r['period']) for r in rows)
        assert math.isclose(independent_npv,s['npv_rub'],abs_tol=.01),s['scenario']
        for r in rows[1:]:
            cf=float(r['revenue_rub'])*.94-float(r['variable_cost_rub'])-float(r['fixed_opex_rub'])-float(r['additional_capex_rub'])+float(r['reserve_return_rub'])
            assert math.isclose(cf,float(r['free_cash_flow_rub']),abs_tol=.01)
    by_id={s['scenario']:s for s in results['scenarios']}
    canonical=load('data/finance_summary.json')
    assert math.isclose(by_id['base']['npv_rub'],canonical['base']['npv_rub'],abs_tol=.01)
    assert by_id['capacity_12']['npv_rub'] < by_id['capacity_16']['npv_rub'] < by_id['capacity_20']['npv_rub'] < by_id['base']['npv_rub']
    assert by_id['capacity20_price1450']['npv_rub']>0
    assert all(by_id[k]['npv_rub']<0 for k in ('capacity16_price1450','capacity20_price1450_down10','capacity20_price1450_costplus10'))
    # Closed-form thresholds independently from exported cash flows and volumes.
    threshold_prices=results['thresholds']['npv_zero_treatment_price_2027_rub']
    for case,key in [('base','unconstrained'),('capacity_20','capacity_20'),('capacity_16','capacity_16'),('capacity_12','capacity_12')]:
        rows=[r for r in annual if r['scenario']==case and int(r['period'])>0]
        pv_one_ruble=sum(float(r['treatment_ha_passes'])*1.04**(int(r['period'])-1)*.94/1.24**int(r['period']) for r in rows)
        closed_form=1300-by_id[case]['npv_rub']/pv_one_ruble
        assert math.isclose(closed_form,threshold_prices[key],abs_tol=.01),key
    saving_factor=sum(1.05**(p-1)/1.24**p for p in range(1,6))
    assert math.isclose(-by_id['base']['npv_rub']/saving_factor,results['thresholds']['npv_zero_net_fixed_saving_year2027_rub'],abs_tol=.01)
    assert by_id['capex_5_5m']['fcf_2027_rub']==by_id['base']['fcf_2027_rub']
    cash=csv_rows('data/liquidity_brainstorm.csv')
    months=csv_rows('data/liquidity_monthly_brainstorm.csv')
    assert len(cash)==12 and len(months)==144
    for c in cash:
        selected=[m for m in months if m['lag_months']==c['lag_months'] and m['advance_fraction']==c['advance_fraction']]
        assert len(selected)==12
        minimum=min([2000000]+[float(m['closing_cash_rub']) for m in selected])
        assert math.isclose(minimum,float(c['minimum_cash_rub']),abs_tol=.01)
        assert math.isclose(max(0,-minimum),float(c['additional_liquidity_rub']),abs_tol=.01)
    cash0={int(c['lag_months']):c for c in cash if float(c['advance_fraction'])==0}
    assert math.isclose(float(cash0[1]['minimum_cash_rub']),-417800,abs_tol=.01)
    assert math.isclose(float(cash0[2]['minimum_cash_rub']),-1568400,abs_tol=.01)
    paths=subprocess.check_output(['git','ls-files','--cached','--others','--exclude-standard','-z'],cwd=ROOT).decode().split('\0')
    brainstorm=[p for p in paths if p and p not in baseline_paths]
    assert all(Path(p).stem.endswith('_brainstorm') for p in brainstorm),brainstorm
    # Validate published local Markdown targets (ignore URLs, anchors, and code literals).
    for name in ['README_brainstorm.md','docs/AGROVECTOR_brainstorm.md']:
        p=ROOT/name
        for target in re.findall(r'\]\(([^)]+)\)',p.read_text()):
            if not target.startswith(('http://','https://','#')):
                assert (p.parent/target.split('#')[0]).exists(),(name,target)
    report=(ROOT/'docs/AGROVECTOR_brainstorm.md').read_text()
    assert str(len(risks)) in report and 'не редактировались' in report
    runway=canonical['cash_runway']
    first_negative=next(m for m in range(1,13) if 2000000-m*runway['no_revenue_fixed_burn_per_month_rub']<0)
    warnings=[]
    if runway['no_revenue_first_negative_month']!=first_negative:
        warnings.append({'file':'data/finance_summary.json','field':'cash_runway.no_revenue_first_negative_month','declared':runway['no_revenue_first_negative_month'],'calculated':first_negative,'scope':'baseline metadata; no change made'})
    if '350 000 RUB/месяц' in (ROOT/'docs/FINANCE_METHOD.md').read_text():
        warnings.append({'file':'docs/FINANCE_METHOD.md','problem':'stale monthly fixed-OPEX prose','correct_rub':runway['no_revenue_fixed_burn_per_month_rub'],'scope':'baseline prose; no change made'})
    return {'verified':True,'baseline_commit':BASE,'baseline_unchanged':True,
            'risk_count':len(risks),'sections_covered':len(coverage),'scenarios':23,'liquidity_cases':12,
            'annual_rows':len(annual),'liquidity_month_rows':len(months),
            'brainstorm_files':len(brainstorm),'known_baseline_warnings':warnings,
            'limitations':['No new legal/market/customer evidence','No new deck generated or motion certified']}

if __name__=='__main__':
    print(json.dumps(verify(),ensure_ascii=False,indent=2))
