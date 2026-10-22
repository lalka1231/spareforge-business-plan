#!/usr/bin/env python3
"""Isolated scenario analysis; imports the current model without running its build.
All outputs are _brainstorm files. This is not an approved replacement baseline.
"""
import csv
import hashlib
import json
import math
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
import build_finance as f

BASELINE = '7c05732803c7bedff0ce15f35647195e61de0790'


def rows_for(*, field_days=None, fixed_saving=0, keep_reserve_return=True,
             customer_absorbs_vat=True, **kwargs):
    rows = f.model(**kwargs)
    previous_revenue = 0
    for r in rows[1:]:
        age = r['period'] - 1
        if field_days is not None:
            required = r['treatment_ha_passes']/130 + r['monitoring_ha_surveys']/500
            retained = sum(s * min(1, field_days/(required*s)) if required*s else 0
                           for s in f.SEASON)
            for key in ('treatment_ha_passes', 'monitoring_ha_surveys',
                        'treatment_revenue_rub', 'monitoring_revenue_rub',
                        'revenue_rub', 'variable_cost_rub'):
                r[key] *= retained
        # A saving is a hypothetical net saving, after extra outsourcing costs.
        r['fixed_opex_rub'] -= fixed_saving * 1.05**age
        assert r['fixed_opex_rub'] >= 0
        active = previous_revenue > r['vat_threshold_assumed_rub']
        cumulative = vat = 0
        # Keep actual monthly delivery weights when a capacity screen is active.
        raw = f.model(**kwargs)[r['period']]
        required = raw['treatment_ha_passes']/130 + raw['monitoring_ha_surveys']/500
        for share in f.SEASON:
            scale = min(1, field_days/(required*share)) if field_days is not None and required*share else 1
            gross = raw['revenue_rub'] * share * scale
            net = gross if customer_absorbs_vat or not active else gross/1.05
            if active:
                vat += net*.05
            cumulative += net
            if cumulative > r['vat_threshold_assumed_rub']:
                active = True
        if not customer_absorbs_vat:
            # Gross customer spending is capped at the old net-price budget.
            r['revenue_rub'] -= vat
            ratio = r['revenue_rub']/(r['treatment_revenue_rub']+r['monitoring_revenue_rub']) if raw['revenue_rub'] else 1
            r['treatment_revenue_rub'] *= ratio
            r['monitoring_revenue_rub'] *= ratio
            r['treatment_price_ex_vat_rub'] *= ratio
            r['monitoring_price_ex_vat_rub'] *= ratio
        previous_revenue = r['revenue_rub']
        r['vat_collected_rub'] = r['vat_remitted_rub'] = vat
        r['customer_billings_including_vat_rub'] = r['revenue_rub']+vat
        r['contribution_rub'] = r['revenue_rub']-r['variable_cost_rub']
        r['ebitda_rub'] = r['contribution_rub']-r['fixed_opex_rub']
        r['ebit_rub'] = r['ebitda_rub']-r['depreciation_rub']
        r['usn_tax_rub'] = r['revenue_rub']*.06
        r['accounting_profit_rub'] = r['ebit_rub']-r['usn_tax_rub']
        if not keep_reserve_return:
            r['reserve_return_rub'] = 0
        r['free_cash_flow_rub'] = r['ebitda_rub']-r['usn_tax_rub']-r['additional_capex_rub']+r['reserve_return_rub']
        r['present_value_rub'] = r['free_cash_flow_rub']/(1+f.RATE)**r['period']
    cumulative = discounted = 0
    for r in rows:
        cumulative += r['free_cash_flow_rub']
        discounted += r['present_value_rub']
        r['cumulative_cf_rub'] = cumulative
        r['cumulative_discounted_cf_rub'] = discounted
    return rows


def liquidity(rows, lag: int = 0, prepayment: float = 0.0) -> dict:
    """Advance at start of service month, balance after lag months; no pre-season advance.
    USN cash provision on receipt; no bad debt, VAT or daily financing model.
    """
    first = rows[1]
    cash = 2_000_000
    result = []
    for i, share in enumerate(f.SEASON):
        current = first['revenue_rub']*share
        old = first['revenue_rub']*f.SEASON[i-lag] if i >= lag else 0
        received = prepayment*current+(1-prepayment)*old
        net = received*.94-first['variable_cost_rub']*share-first['fixed_opex_rub']/12
        cash += net
        result.append({'month': i+1, 'lag_months': lag, 'advance_fraction': prepayment,
                       'receipts_rub': received, 'net_cash_rub': net, 'closing_cash_rub': cash})
    minimum = min([2_000_000]+[x['closing_cash_rub'] for x in result])
    return {'minimum_cash_rub': minimum, 'additional_liquidity_rub': max(0, -minimum),
            'year_end_cash_rub': cash, 'months': result}


def threshold(fn, lo, hi):
    if fn(lo) >= 0:
        return lo
    if fn(hi) < 0:
        return None
    for _ in range(70):
        mid = (lo+hi)/2
        if fn(mid) < 0:
            lo = mid
        else:
            hi = mid
    return (lo+hi)/2


def write_csv(name, rows):
    with (ROOT/'data'/name).open('w', encoding='utf-8', newline='') as out:
        writer = csv.DictWriter(out, fieldnames=list(rows[0]), lineterminator='\n')
        writer.writeheader()
        writer.writerows(rows)


def main():
    tracked = subprocess.check_output(['git', 'ls-files', '-z'], cwd=ROOT).decode().split('\0')
    protected = {p: hashlib.sha256((ROOT/p).read_bytes()).hexdigest()
                 for p in tracked if p and not Path(p).stem.endswith('_brainstorm')}
    specs = [
        ('base', 'Текущая база, не рекомендация', {}),
        ('treatment_1400', 'Внесение 1400, спрос не подтвержден', {'treatment_price_factor':1400/1300}),
        ('treatment_1450', 'Внесение 1450, спрос не подтвержден', {'treatment_price_factor':1450/1300}),
        ('monitoring_350', 'Мониторинг 350, ценность отчета не подтверждена', {'price_factor':350/250,'treatment_price_factor':250/350}),
        ('net_fixed_saving_300k', 'Чистая экономия 300 тыс. в 2027 с индексацией; доступность команды не доказана', {'fixed_saving':300000}),
        ('net_fixed_saving_600k', 'Чистая экономия 600 тыс.; не разрешение урезать зарплаты', {'fixed_saving':600000}),
        ('capex_5_5m', 'CAPEX 5.5 млн: комплектность и сервис не подтверждены', {'initial_capex':5500000}),
        ('procurement_8_98m', 'CAPEX 8.98 млн: неполный ценовой стресс', {'initial_capex':8980000}),
        ('capacity_20', '20 эффективных дней в месяц: ресурсный экран, не погода', {'field_days':20}),
        ('capacity_16', '16 эффективных дней: гипотетический более жесткий экран', {'field_days':16}),
        ('capacity_12', '12 эффективных дней: гипотетический стресс', {'field_days':12}),
        ('capacity20_price1400', '20 дней и внесение1400', {'field_days':20,'treatment_price_factor':1400/1300}),
        ('capacity20_price1450', '20 дней и внесение1450', {'field_days':20,'treatment_price_factor':1450/1300}),
        ('capacity16_price1450', '16 дней и внесение1450', {'field_days':16,'treatment_price_factor':1450/1300}),
        ('capacity20_price1450_down10', '20 дней, внесение1450, объемы минус10%', {'field_days':20,'treatment_price_factor':1450/1300,'volume_factor':.9}),
        ('capacity20_price1450_costplus10', '20 дней, внесение1450, переменные плюс10%', {'field_days':20,'treatment_price_factor':1450/1300,'variable_factor':1.1}),
        ('price_minus10', 'Оба тарифа минус10%', {'price_factor':.9}),
        ('volume_minus25', 'Оба объема минус25%', {'volume_factor':.75}),
        ('variable_plus20', 'Переменные плюс20%', {'variable_factor':1.2}),
        ('one_year_delay', 'Задержка с условным OPEX 900тыс. до найма, не с полным штатом', {'delay':True}),
        ('no_reserve_release', 'Продолжение без возврата резерва, без terminal value', {'keep_reserve_return':False}),
        ('vat_inside_customer_budget', 'НДС внутри прежнего бюджета покупателя, а не сверх', {'customer_absorbs_vat':False}),
        ('combined_downside', 'Объемы−25%, тарифы−10%, переменные+20%, CAPEX8.98млн', {'volume_factor':.75,'price_factor':.9,'variable_factor':1.2,'initial_capex':8980000}),
    ]
    scenarios, annual, by_id = [], [], {}
    for key, label, params in specs:
        rows = rows_for(**params)
        by_id[key] = rows
        values = f.metrics(rows)
        scenarios.append({'scenario':key, 'description':label, **values,
                          'revenue_2027_rub':rows[1]['revenue_rub'],
                          'ebitda_2027_rub':rows[1]['ebitda_rub'],
                          'fcf_2027_rub':rows[1]['free_cash_flow_rub'],
                          'parameters_json':json.dumps(params, ensure_ascii=False),
                          'status':'[РАСЧЕТ] на непроверенных допущениях'})
        annual.extend({'scenario':key, **r} for r in rows)
    base = by_id['base']
    cash: list[dict] = [dict(lag_months=lag, advance_fraction=advance, **liquidity(base,lag,advance))
            for lag in (0,1,2,3) for advance in (0,.3,.5)]
    monthly = [dict(m) for x in cash for m in x['months']]
    for x in cash:
        del x['months']
    prices = {}
    for days in (None,20,16,12):
        key = 'unconstrained' if days is None else f'capacity_{days}'
        prices[key] = threshold(lambda p:f.metrics(rows_for(field_days=days,treatment_price_factor=p/1300))['npv_rub'],1300,4000)
    thresholds = {
        'npv_zero_treatment_price_2027_rub':prices,
        'npv_zero_net_fixed_saving_year2027_rub':threshold(lambda s:f.metrics(rows_for(fixed_saving=s))['npv_rub'],0,2000000),
        'year1_operating_cash_zero_treatment_price_rub':(base[1]['fixed_opex_rub']+base[1]['variable_cost_rub']-.94*base[1]['monitoring_revenue_rub'])/(.94*base[1]['treatment_ha_passes']),
        'year1_cash_zero_proportional_volume_factor':base[1]['fixed_opex_rub']/(base[1]['contribution_rub']-base[1]['usn_tax_rub']),
        'extra_initial_reserve_no_revenue_year1_rub':base[1]['fixed_opex_rub']-2000000,
        'year1_monthly_fixed_rub':base[1]['fixed_opex_rub']/12,
        'npv_zero_capex_reduction_rub':-f.metrics(base)['npv_rub'],
    }
    # Independent regression evidence, not only self-comparison between generated files.
    canonical = json.loads((ROOT/'data/finance_summary.json').read_text())
    assert math.isclose(f.metrics(base)['npv_rub'], canonical['base']['npv_rub'],abs_tol=.01)
    assert math.isclose(sum(r['free_cash_flow_rub']/1.24**r['period'] for r in base),canonical['base']['npv_rub'],abs_tol=.01)
    cap = json.loads((ROOT/'data/capacity_scenario_summary.json').read_text())
    assert math.isclose(f.metrics(by_id['capacity_20'])['npv_rub'],cap['npv_rub'],abs_tol=.01)
    assert math.isclose(by_id['base'][1]['ebitda_rub'],8800000-2640000-5887200,abs_tol=.01)
    assert math.isclose(liquidity(base)['minimum_cash_rub'],528200,abs_tol=.01)
    old_cash = json.loads((ROOT/'data/supplemental_summary.json').read_text())['collection_delay']
    for c in old_cash:
        assert math.isclose(liquidity(base,c['delay_months'])['minimum_cash_rub'],c['min_cash_rub'],abs_tol=.01)
    for name, rows in by_id.items():
        for r in rows[1:]:
            assert math.isclose(r['revenue_rub'],r['treatment_revenue_rub']+r['monitoring_revenue_rub'],abs_tol=.01),name
            assert math.isclose(r['free_cash_flow_rub'],r['revenue_rub']*.94-r['variable_cost_rub']-r['fixed_opex_rub']-r['additional_capex_rub']+r['reserve_return_rub'],abs_tol=.01),name
    for key,p in prices.items():
        days = None if key=='unconstrained' else int(key.split('_')[-1])
        assert abs(f.metrics(rows_for(field_days=days,treatment_price_factor=p/1300))['npv_rub']) < .01
    assert all(hashlib.sha256((ROOT/p).read_bytes()).hexdigest()==sha for p,sha in protected.items())
    limitations = [
        'Не вероятностный прогноз: не задавались распределения и вероятности погоды/спроса.',
        'Цена/экономия/CAPEX — тесты чувствительности, не полученные КП и не рекомендации внедрять их без проверки.',
        'Одна бригада; равномерная потеря обеих услуг в перегруженном месяце; переносов и приоритизации нет.',
        'Эффективные дни не подтверждены наблюдениями; рабочие нормативы 130/500 остаются гипотезами.',
        'Налоговые параметры унаследованы, не подтверждены повторно; УСН6% без уменьшения на взносы.',
        'VAT-inside-budget: налоговая активация приближенно по месячной сезонности; конечный тариф проверяет бухгалтер.',
        'Дебиторка: аванс в месяц услуги, не до сезона; остаток с лагом; налоги резервируются при получении; нет плохих долгов.',
        'Остатки на конец месяца не исключают внутримесячный разрыв; нет календаря ежедневных платежей.',
        'Без возврата резерва — только проверка одного притока, не полноценная оценка продолжающегося бизнеса.',
        'Сценарий задержки с OPEX900тыс. предполагает, что штат еще не нанят; полный штат стоит5887200 в год.',
        'NPV проекта не равен доходности конкретного инвестора: доли, дивиденды, долг и стоимость выхода не заданы.',
    ]
    result={'baseline_commit':BASELINE,'classification':'[РАСЧЕТ]; все новые параметры [ДОПУЩЕНИЕ]',
            'scenarios':scenarios,'thresholds':thresholds,'liquidity':cash,'limitations':limitations,
            'verification':{'baseline_reconciled':True,'capacity20_reconciled':True,
                            'delay_cases_reconciled':True,'protected_baseline_files_unchanged':True,
                            'source_sha256':protected}}
    (ROOT/'data/resilience_brainstorm.json').write_text(json.dumps(result,ensure_ascii=False,indent=2,allow_nan=False)+'\n')
    write_csv('scenario_comparison_brainstorm.csv',scenarios)
    write_csv('scenario_annual_brainstorm.csv',annual)
    write_csv('liquidity_brainstorm.csv',cash)
    write_csv('liquidity_monthly_brainstorm.csv',monthly)
    print(json.dumps({'scenario_count':len(scenarios),'liquidity_cases':len(cash),'thresholds':thresholds,'baseline_unchanged':True},ensure_ascii=False,indent=2))

if __name__=='__main__':
    main()
