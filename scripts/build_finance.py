#!/usr/bin/env python3
"""Reproducible AgroVector model. Standard library only; RUB, nominal prices."""
import csv
import json
import math
from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[1]


TREATMENT = [6000, 9000, 11000, 12000, 13000]
MONITORING = [4000, 6000, 8000, 9000, 10000]
REPLACEMENTS = [0, 300000, 600000, 900000, 600000]
FIXED = {'payroll_including_contributions': 3000000, 'rent': 360000,
         'insurance_compliance': 300000, 'marketing': 240000,
         'it': 120000, 'administration': 180000}
EQUIPMENT = [('agrodrones_main_and_backup_with_batteries_charging', 2, 4000000),
             ('vehicle_trailer', 1, 1400000), ('monitoring_drone', 1, 600000),
             ('mixing_safety', 1, 400000), ('it', 1, 200000),
             ('capital_spares', 1, 400000)]
RATE = .24


def model(price_factor=1., volume_factor=1., variable_factor=1., delay=False, treatment_price_factor=1.):
    rows = []
    replacements = ([0] + REPLACEMENTS[:4]) if delay else REPLACEMENTS
    previous_revenue = 0
    for i in range(5):
        age = i - int(delay)
        treat = TREATMENT[age] * volume_factor if age >= 0 else 0
        monitor = MONITORING[age] * volume_factor if age >= 0 else 0
        pt, pm = 1300 * 1.04**i * price_factor * treatment_price_factor, 250 * 1.04**i * price_factor
        vt, vm = 400 * 1.05**i * variable_factor, 60 * 1.05**i * variable_factor
        revenue_t, revenue_m = treat * pt, monitor * pm
        revenue = revenue_t + revenue_m
        variable = treat * vt + monitor * vm
        fixed = 900000 if age < 0 else sum(FIXED.values()) * 1.05**i
        threshold = [20000000,20000000,20000000,15000000,10000000][i]
        # Start-year test and next-month activation, per published July 2026 amendment.
        vat_active = previous_revenue > threshold
        cumulative_revenue = 0
        vat = 0
        for share in SEASON:
            monthly_revenue = revenue * share
            if vat_active: vat += monthly_revenue * .05
            cumulative_revenue += monthly_revenue
            if cumulative_revenue > threshold: vat_active = True
        previous_revenue = revenue
        depreciation = 7000000 / 5 + sum(replacements[j] / 3 for j in range(i+1) if i-j < 3)
        ebitda = revenue - variable - fixed
        tax = revenue * .06
        recovery = 2000000 if i == 4 else 0
        cf = ebitda - tax - replacements[i] + recovery
        rows.append(dict(year=2027+i, period=i+1, treatment_ha_passes=treat,
                         monitoring_ha_surveys=monitor, treatment_price_ex_vat_rub=pt,
                         monitoring_price_ex_vat_rub=pm, treatment_variable_unit_rub=vt,
                         monitoring_variable_unit_rub=vm, treatment_revenue_rub=revenue_t,
                         monitoring_revenue_rub=revenue_m, revenue_rub=revenue,
                         variable_cost_rub=variable, contribution_rub=revenue-variable,
                         fixed_opex_rub=fixed, ebitda_rub=ebitda, depreciation_rub=depreciation,
                         ebit_rub=ebitda-depreciation, usn_tax_rub=tax,
                         accounting_profit_rub=ebitda-depreciation-tax,
                         vat_threshold_assumed_rub=threshold, vat_collected_rub=vat,
                         vat_remitted_rub=vat, customer_billings_including_vat_rub=revenue+vat,
                         additional_capex_rub=replacements[i], initial_capex_rub=0,
                         launch_cost_rub=0, reserve_funding_rub=0,
                         reserve_return_rub=recovery, free_cash_flow_rub=cf,
                         discount_factor=1/(1+RATE)**(i+1), present_value_rub=cf/(1+RATE)**(i+1)))
    initial = dict.fromkeys(rows[0], 0)
    initial.update(year=2026, period=0, initial_capex_rub=7000000, launch_cost_rub=500000,
                   reserve_funding_rub=2000000, free_cash_flow_rub=-9500000,
                   discount_factor=1, present_value_rub=-9500000)
    rows.insert(0, initial)
    cumulative = discounted = 0
    for row in rows:
        cumulative += row['free_cash_flow_rub']
        discounted += row['present_value_rub']
        row['cumulative_cf_rub'] = cumulative
        row['cumulative_discounted_cf_rub'] = discounted
    return rows


# April–September; matches the 22-week commercial season in project_plan.csv.
SEASON = [0, 0, 0, .10, .20, .25, .20, .15, .10, 0, 0, 0]


def payback(flows):
    cumulative = flows[0]
    for i, cf in enumerate(flows[1:], 1):
        if cumulative < 0 <= cumulative + cf:
            return i-1 + (-cumulative)/cf
        cumulative += cf
    return None


def metrics(rows) -> dict:
    flows = [r['free_cash_flow_rub'] for r in rows]
    discounted = [r['present_value_rub'] for r in rows]
    # Interpolation must not spread the end-Y5 reserve return across Y5.
    def adjusted_payback(values, discounted_mode=False):
        before_final = sum(values[:-1])
        reserve = rows[-1]['reserve_return_rub']
        if discounted_mode:
            reserve *= rows[-1]['discount_factor']
        operating_final = values[-1] - reserve
        if before_final < 0 and before_final + operating_final < 0 <= before_final + values[-1]:
            return 5.0
        modified = values[:-1] + [operating_final]
        return payback(modified)
    return dict(initial_investment_rub=-flows[0], discount_rate=RATE,
                npv_rub=sum(discounted), cumulative_net_cf_rub=sum(flows),
                simple_payback_years=adjusted_payback(flows),
                discounted_payback_years=adjusted_payback(discounted, True),
                terminal_equipment_value_rub=0,
                terminal_book_value_not_cash_rub=7000000+sum(r['additional_capex_rub'] for r in rows)-sum(r['depreciation_rub'] for r in rows))


def seasonality(year1):
    cash = 2000000
    monthly = []
    for month, share in enumerate(SEASON, 1):
        revenue = year1['revenue_rub'] * share
        variable = year1['variable_cost_rub'] * share
        fixed = year1['fixed_opex_rub'] / 12
        tax = revenue * .06
        net = revenue-variable-fixed-tax
        monthly.append(dict(month=f'2027-{month:02}', share=share,
                            treatment_ha_passes=year1['treatment_ha_passes']*share,
                            monitoring_ha_surveys=year1['monitoring_ha_surveys']*share,
                            opening_cash_rub=cash, revenue_collected_ex_vat_rub=revenue,
                            variable_cost_paid_rub=variable, fixed_opex_paid_rub=fixed,
                            usn_cash_provision_rub=tax, net_operating_cf_rub=net,
                            closing_cash_rub=cash+net))
        cash += net
    return monthly


class FinanceTests(unittest.TestCase):
    def test_metrics_and_seasonality(self):
        rows = model()
        result = metrics(rows)
        self.assertAlmostEqual(result['npv_rub'], sum(r['present_value_rub'] for r in rows))
        self.assertEqual(payback([-100, 40, 60]), 2)
        self.assertIsNone(payback([-100, 40, 59]))
        monthly = seasonality(rows[1])
        self.assertEqual(len(monthly), 12)
        self.assertAlmostEqual(sum(m['net_operating_cf_rub'] for m in monthly), 1432000)
        self.assertAlmostEqual(monthly[-1]['closing_cash_rub'], 3432000)
        self.assertAlmostEqual(min(m['closing_cash_rub'] for m in monthly), 950000)

    def test_accounting_tax_and_stress(self):
        rows = model()
        self.assertEqual([r['depreciation_rub'] for r in rows[1:]], [1400000, 1500000, 1700000, 2000000, 2100000])
        self.assertEqual(rows[1]['vat_collected_rub'], 0)
        self.assertEqual(rows[2]['vat_collected_rub'], 0)
        self.assertEqual(rows[3]['vat_collected_rub'], 0)
        self.assertAlmostEqual(rows[4]['vat_collected_rub'], rows[4]['revenue_rub']*.05)
        for r in rows[1:]:
            self.assertAlmostEqual(r['free_cash_flow_rub'], r['ebitda_rub']-r['usn_tax_rub']-r['additional_capex_rub']+r['reserve_return_rub'])
            self.assertEqual(r['vat_collected_rub'], r['vat_remitted_rub'])
        self.assertEqual(model(delay=True)[1]['free_cash_flow_rub'], -900000)
        for kwargs in [{'price_factor': .9}, {'volume_factor': .75}, {'variable_factor': 1.2}, {'delay': True}]:
            self.assertLess(metrics(model(**kwargs))['npv_rub'], metrics(rows)['npv_rub'])

    def test_base_cashflow(self):
        rows = model()
        self.assertEqual(len(rows), 6)
        self.assertEqual(rows[0]['free_cash_flow_rub'], -9_500_000)
        self.assertAlmostEqual(rows[1]['revenue_rub'], 8_800_000)
        self.assertAlmostEqual(rows[1]['ebitda_rub'], 1_960_000)
        self.assertAlmostEqual(rows[1]['free_cash_flow_rub'], 1_432_000)


def write_csv(name, records):
    path = ROOT / 'data' / name
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('w', encoding='utf-8', newline='') as stream:
        writer = csv.DictWriter(stream, fieldnames=list(records[0]), lineterminator='\n')
        writer.writeheader()
        writer.writerows({k: round(v, 8) if isinstance(v, float) else v for k, v in row.items()} for row in records)


def build():
    rows = model()
    monthly = seasonality(rows[1])
    scenarios = [('base', {}), ('price_minus_10pct', {'price_factor': .9}),
                 ('volume_minus_25pct', {'volume_factor': .75}),
                 ('variable_plus_20pct', {'variable_factor': 1.2}),
                 ('one_year_delay', {'delay': True}),
                 ('treatment_price_900', {'treatment_price_factor': 900/1300})]
    sensitivity = []
    scenario_rows = {}
    for name, kwargs in scenarios:
        scenario_rows[name] = model(**kwargs)
        sensitivity.append(dict(scenario=name, **metrics(scenario_rows[name])))
    units = []
    for r in rows[1:]:
        for service, prefix in [('treatment', 'treatment'), ('monitoring', 'monitoring')]:
            price, cost = r[prefix+'_price_ex_vat_rub'], r[prefix+'_variable_unit_rub']
            units.append(dict(year=r['year'], service=service,
                              unit='ha_pass' if service == 'treatment' else 'ha_survey',
                              price_ex_vat_rub=price, variable_cost_rub=cost,
                              contribution_before_usn_rub=price-cost,
                              usn_per_unit_rub=price*.06,
                              contribution_after_usn_rub=price*.94-cost,
                              status='[РАСЧЕТ]; фиксированные расходы не распределены'))
    assumptions = []
    def assumption(key, value, unit, note):
        assumptions.append(dict(parameter=key, value=value, unit=unit,
                                status='[ДОПУЩЕНИЕ]; проектный параметр', note=note))
    for key, value, unit in [('initial_capex',7000000,'RUB'), ('launch_expense_t0',500000,'RUB'),
                             ('constant_working_reserve',2000000,'RUB'), ('discount_rate',RATE,'fraction'),
                             ('price_indexation',.04,'fraction/year'), ('cost_indexation',.05,'fraction/year'),
                             ('usn_revenue_tax',.06,'fraction'), ('vat_surcharge',.05,'fraction'),
                             ('treatment_price_2027',1300,'RUB/ha_pass'), ('monitoring_price_2027',250,'RUB/ha_survey'),
                             ('treatment_variable_2027',400,'RUB/ha_pass'), ('monitoring_variable_2027',60,'RUB/ha_survey'),
                             ('crew_count',1,'crew'), ('terminal_equipment_value',0,'RUB')]:
        assumption(key,value,unit,'Разработано для модели; не факт заказчика; цены без НДС; номинальные RUB')
    for key, value in FIXED.items():
        assumption('fixed_'+key, value, 'RUB/year_2027', 'Индексируется на 5% в год')
    for i in range(5):
        for key, value, unit in [('treatment_volume',TREATMENT[i],'ha_pass'),
                                 ('monitoring_volume',MONITORING[i],'ha_survey'),
                                 ('replacement_capex',REPLACEMENTS[i],'RUB'),
                                 ('vat_threshold',[20000000,20000000,20000000,15000000,10000000][i],'RUB')]:
            assumption(f'{key}_{2027+i}',value,unit,'CAPEX задан номинально, без дополнительной индексации; НДС — бюджетное допущение')
    for i, share in enumerate(SEASON, 1):
        assumption(f'season_share_2027_{i:02}',share,'fraction','Одинаковый профиль услуг; оплата в месяц выполнения; фиксированные расходы равномерно')
    assumption('initial_depreciation_years',5,'years','Линейно с Y1, включая капитализированный запас')
    assumption('replacement_depreciation_years',3,'years','Линейно с года покупки')
    assumption('delay_year1_fixed',900000,'RUB','Выручка ноль; объемы и замены сдвинуты, инфляция календарная; t0 не сдвигается')
    assumption('collection_delay',0,'months','Нет дебиторской задолженности; нужно проверять контрактами')
    assumption('chemicals_customer_supplied',1,'boolean','Стоимость препаратов не включена в тариф/переменные: клиент предоставляет; подтвердить договорами')
    assumption('vat_timing','previous_year_then_next_month','model_rule','Пороги источники12;13; проверка прошлого года и начисление со следующего месяца превышения; сезонность SEASON')
    for key,value,unit,note in [
        ('treatment_fuel_logistics',150,'RUB/ha_pass','Часть400; топливо и логистика'),
        ('treatment_energy',50,'RUB/ha_pass','Часть400; энергия зарядки'),
        ('treatment_routine_service',130,'RUB/ha_pass','Часть400; текущие расходники; не АКБ из CAPEX замен'),
        ('treatment_ppe_disposal',70,'RUB/ha_pass','Часть400; СИЗ, мойка, обращение с отходами'),
        ('monitoring_logistics',30,'RUB/ha_survey','Часть60'),
        ('monitoring_energy_service',20,'RUB/ha_survey','Часть60; не амортизация'),
        ('monitoring_data',10,'RUB/ha_survey','Часть60; переменная обработка данных'),
        ('compliance_insurance',120000,'RUB/year','Часть300000'),
        ('compliance_legal',60000,'RUB/year','Часть300000'),
        ('compliance_agronomist',120000,'RUB/year','Часть300000; экономный внешний ретейнер'),
        ('admin_accounting',120000,'RUB/year','Часть180000'),
        ('admin_bank_other',60000,'RUB/year','Часть180000'),
        ('tam_farms',2000,'farms','Неподтвержденная корзина спроса; не статистика РФ'),
        ('sam_farms',200,'farms','Нужна адресная CRM база в радиусе150км'),
        ('farm_area',2000,'ha','Среднее сценарной корзины'),
        ('eligible_share',.15,'fraction','Фильтр агрономии, права и спроса'),
        ('annual_passes_eligible',1.5,'passes/ha','Сценарный TAM/SAM'),
        ('annual_monitoring_eligible',.5,'surveys/ha','Сценарный TAM/SAM'),
        ('processing_day_capacity',130,'ha_pass/day','Включает переезды и подготовку'),
        ('monitoring_day_capacity',500,'ha_survey/day','Последовательно с внесением'),
        ('season_weeks',22,'weeks','154дня;100обработка+20мониторинг+34резерв'),
        ('founder_funding',2500000,'RUB','Не определяет корпоративную долю'),
        ('investor_funding',7000000,'RUB','Неподтвержденное долевое финансирование')]:
        assumption(key,value,unit,note)
    summary = dict(project='АгроВектор', currency='RUB', classification='Все входные данные — допущения; выходные показатели — расчеты',
                   base=metrics(rows), sensitivity=sensitivity, yearly_scenarios=scenario_rows,
                   first_year_monthly=monthly,
                   cash_runway=dict(initial_operating_cash_rub=2000000, capex_recharged=False,
                                    no_revenue_fixed_burn_per_month_rub=350000,
                                    no_revenue_runway_months=2000000/350000,
                                    no_revenue_first_negative_month=6,
                                    base_minimum_month_end_cash_rub=min(m['closing_cash_rub'] for m in monthly),
                                    base_first_negative_month=next((m['month'] for m in monthly if m['closing_cash_rub']<0),None),
                                    base_year_end_cash_rub=monthly[-1]['closing_cash_rub'],
                                    base_additional_funding_required_rub=max(0,-min(m['closing_cash_rub'] for m in monthly))),
                   break_even_2027=dict(after_usn_contribution_rub=rows[1]['contribution_rub']-rows[1]['usn_tax_rub'],
                                        proportional_volume_factor=4200000/(rows[1]['contribution_rub']-rows[1]['usn_tax_rub']),
                                        treatment_ha_passes=6000*4200000/5632000,
                                        monitoring_ha_surveys=4000*4200000/5632000),
                   limitations=['Нет подтвержденных контрактов, цен поставщиков и производительности бригады',
                                'НДС сверх цены должен приниматься покупателем; нет моделирования снижения спроса из-за НДС',
                                'Погода, регуляторные ограничения, логистика и дебиторка могут сорвать сезон',
                                'Возврат резерва в Y5 предполагает завершение проекта; нет стоимости продолжающегося бизнеса',
                                'Однофакторные стрессы не заменяют комбинированный кризисный сценарий'])
    for row in rows:
        row['classification']='[РАСЧЕТ]; входные параметры [ДОПУЩЕНИЕ]'
    write_csv('financial_model_5y.csv', rows)
    write_csv('monthly_cashflow_2027.csv', monthly)
    write_csv('assumptions.csv', assumptions)
    write_csv('equipment.csv', [dict(item=name, quantity=qty, total_cost_rub=cost,
                                   depreciation_years=5, annual_depreciation_rub=cost/5,
                                   status='[ДОПУЩЕНИЕ]; бюджет, не коммерческое предложение') for name,qty,cost in EQUIPMENT])
    write_csv('unit_economics.csv', units)
    write_csv('sensitivity.csv', sensitivity)
    (ROOT/'data/finance_summary.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2, allow_nan=False)+'\n', encoding='utf-8')
    method = '''# АгроВектор: методика воспроизводимой финансовой модели

## Статус данных и запуск
[ДОПУЩЕНИЕ] Цены, объемы, комплектация и сезонность — проектные гипотезы модели, а не предоставленные заказчиком факты. Выходные показатели помечаются [РАСЧЕТ]. Коммерческие предложения и договоры отсутствуют. Правовые источники и пределы налоговой применимости приведены в FINAL_BUSINESS_PLAN.md и 09_SOURCES.md; упрощенный бюджетный алгоритм не является налоговым календарем.

Запуск из корня: `python3 scripts/build_finance.py`; проверка без записи: `python3 scripts/build_finance.py --test`. Python 3.10+, только стандартная библиотека. Скрипт является источником истины, генерирует пять CSV, JSON и этот документ. CSV: UTF-8, разделитель запятая, десятичная точка, RUB; расчеты без промежуточного округления. JSON содержит полный помесячный план и все годовые сценарии.

## Границы проекта и t0
Одна мобильная бригада; два агродрона — основной и резервный, **не две параллельные бригады**; отдельный мониторинговый аппарат. Га-проходы и га-обследования — разные единицы, не уникальные гектары и не заказы. Препараты предоставляет заказчик (дополнительное допущение, требует договора).

Начальные основные средства 7 000 000 RUB, запуск 500 000 RUB, оборотный резерв 2 000 000 RUB. CF0 = −9 500 000 RUB; строка 2026 обозначает условный t0 непосредственно перед 2027, не полный операционный год. Расход запуска полностью учтен в t0, не амортизируется. Капитализированный запас 400 000 RUB включен в 7 млн и в пятилетнюю амортизацию как допущение учета.

## Годовые формулы
Цена = цена_2027 × 1,04^(год−2027); удельные и фиксированные расходы = база × 1,05^(год−2027). Замены 0/0,3/0,6/0,9/0,6 млн заданы номинально, повторно не индексируются.

Выручка = га-проходы × тариф обработки + га-обследования × тариф мониторинга. EBITDA = выручка − переменные − фиксированный OPEX. УСН = 6% выручки без НДС, без уменьшения на взносы, даже при убытке. EBIT = EBITDA − амортизация; расчетная бухгалтерская прибыль = EBIT − УСН (не регламентированная отчетность).

Первоначальная амортизация 7 млн / 5 ежегодно. Каждая замена амортизируется за 3 года с года покупки. Остаточная балансовая стоимость не является притоком, терминальная стоимость оборудования равна нулю. Амортизация не уменьшает налог УСН «доходы» и не вычитается из CF второй раз.

CF = EBITDA − УСН − CAPEX замен + возврат резерва. Резерв 2 млн финансируется ровно один раз в t0, номинально неизменен и полностью возвращается **в конце Y5**. Его использование внутри сезона не означает второй расход на оборотный капитал. Это модель завершения проекта в 2031; для продолжающегося бизнеса возврат резерва неуместен. Долга, грантов, субсидий, процентов и терминального бизнеса нет.

NPV = CF0 + Σ CFt / 1,24^t. Годовые денежные потоки дисконтируются на конец года. Окупаемость: первое пересечение нуля накопленным CF (либо дисконтированным CF); доля года — условная линейная интерполяция годового операционного потока, не точная календарная дата. Возврат резерва не размазывается по Y5: если без него пересечения нет, окупаемость только в 5,00. `null` в JSON / пусто в CSV означает «не достигнута за 5 лет», экстраполяции нет.

## НДС: бюджет, не налоговый календарь
[ФАКТ] Опубликованный в июле 2026 года график освобождения от НДС на УСН: 2027–2029 — 20 млн, 2030 — 15 млн, 2031 — 10 млн руб.; источники [12][13] в реестре. [ДОПУЩЕНИЕ] Ставка после утраты освобождения — 5% сверх цены, без вычета входного НДС, при сохранении права на этот режим. [РАСЧЕТ] Алгоритм проверяет выручку предыдущего года на начало года; затем накопленную месячную выручку и начисляет НДС со следующего месяца после превышения. В базе 2027–2029 освобождены, с января2030 возникает обязанность из-за выручки2029 выше15млн. Месячная сезонность для теста порога одинакова во все годы. Сроки перечисления налога и авансовые счета-фактуры упрощены; финальный налоговый календарь утверждает бухгалтер.

НДС добавляется сверх тарифа: платеж клиента = выручка × 1,05 в облагаемые годы. Собранный НДС = перечисленный НДС, транзит исключен из выручки, базы УСН, EBITDA и CF. Входящие расходы — полный денежный расход, возмещения входного НДС нет. Такой транзит не снижает маржу, но перенос роста цены на клиента не гарантирован. Сроки перечисления НДС не моделируются.

## Стрессы и единичная экономика
Каждый стресс независим: тарифы обоих услуг −10% во все годы; оба объема −25%; обе переменные ставки +20% (фиксированные неизменны). Задержка: t0 сохраняется, 2027 без выручки и с OPEX 0,9 млн; в 2028–2031 объемы базовых Y1–Y4; замены сдвигаются на год (0/0/0,3/0,6/0,9 млн), индексация по календарю. Начальные активы стареют с Y1 даже в простое. Резерв возвращается в исходном Y5, шестого года не добавляем. Дополнительное финансирование при убытках не скрыто в грантах.

Единичная экономика показывает маржу до/после УСН без произвольного распределения постоянных расходов. Безубыточность рассчитана при пропорциональном сохранении структуры обоих услуг; это не независимые пороги по каждой услуге.

## Месячная сезонность и cash runway
[ДОПУЩЕНИЕ] Апрель/май/июнь/июль/август/сентябрь = 10/20/25/20/15/10% годовых объемов обоих услуг, остальные месяцы ноль; соответствует сезону 12 апреля — 12 сентября в Ганте. Деньги получены в месяц выполнения, переменные оплачены тогда же; фиксированные 350 000 RUB/месяц. УСН резервируется ежемесячно (не календарь платежей). Первый год освобожден от НДС по бюджетному правилу. Нет дебиторки, отсрочек, авансов и ежемесячного CAPEX. Старт месячной таблицы — **2 млн после оплаты CAPEX и запуска**, повторного вычитания 7,5 млн нет. Остатки проверены на конец месяца, внутримесячный кассовый разрыв не исключен.

Runway без выручки = 2 млн / 350 тыс.; это отдельный статический стресс, не срок жизни базового бизнеса. Месячные остатки включают накопленную прибыль, а не требуют держать все 2 млн заблокированными каждый день. При полном отсутствии поступлений деньги заканчиваются в шестом месяце. Полный возврат резерва в Y5 — предположение годового проекта, не дополнительный доход первого года.

## Рассчитанные результаты
'''
    fmt = lambda x: 'не достигнута за 5 лет' if x is None else f'{x:.3f}'
    base = summary['base']
    method += f"\nИнвестиция: {base['initial_investment_rub']:,.0f} RUB; NPV(24%): **{base['npv_rub']:,.2f} RUB**. Простая окупаемость: {fmt(base['simple_payback_years'])} года; дисконтированная: {fmt(base['discounted_payback_years'])}.\n\n"
    method += '| Год | Выручка без НДС | EBITDA | УСН | CF |\n|---|---:|---:|---:|---:|\n'
    for r in rows[1:]:
        method += '| '+str(r['year'])+' | '+' | '.join(f'{r[k]:,.2f}' for k in ['revenue_rub','ebitda_rub','usn_tax_rub','free_cash_flow_rub'])+' |\n'
    method += '\n| Сценарий | NPV, RUB | Простая окупаемость, лет | Дисконтированная, лет |\n|---|---:|---:|---:|\n'
    for s in sensitivity:
        method += f"| {s['scenario']} | {s['npv_rub']:,.2f} | {fmt(s['simple_payback_years'])} | {fmt(s['discounted_payback_years'])} |\n"
    method += '\n| Месяц 2027 | Доля | Поступления без НДС | Операционный CF | Остаток денег |\n|---|---:|---:|---:|---:|\n'
    for m in monthly:
        method += f"| {m['month']} | {m['share']:.0%} | {m['revenue_collected_ex_vat_rub']:,.0f} | {m['net_operating_cf_rub']:,.0f} | {m['closing_cash_rub']:,.0f} |\n"
    method += f"\nRunway без выручки: {2000000/350000:.3f} месяца; минимум базового остатка: {summary['cash_runway']['base_minimum_month_end_cash_rub']:,.0f} RUB; конец Y1: {monthly[-1]['closing_cash_rub']:,.0f} RUB.\n"
    method += '\n## Слабые места\n' + '\n'.join('- '+x+'.' for x in summary['limitations'])
    method += '\n- Положительный NPV базы (если получен) не доказывает инвестиционную готовность: необходимо подтвердить объемы, приемлемость тарифов и сезонную пропускную способность одной бригады. Отрицательные стрессовые NPV не сглаживаются.\n- Резерв не индексируется и не растет с выручкой; дебиторка или перенос начала сезона способны потребовать дополнительного финансирования.\n'
    (ROOT/'docs').mkdir(exist_ok=True)
    (ROOT/'docs/FINANCE_METHOD.md').write_text(method, encoding='utf-8')
    # Read back generated artifacts to verify row counts and reconciliation.
    for name, expected in [('financial_model_5y.csv',6),('assumptions.csv',len(assumptions)),
                           ('equipment.csv',6),('unit_economics.csv',10),('sensitivity.csv',6)]:
        with (ROOT/'data'/name).open(encoding='utf-8',newline='') as stream:
            assert len(list(csv.DictReader(stream))) == expected, name
    loaded = json.loads((ROOT/'data/finance_summary.json').read_text(encoding='utf-8'))
    assert math.isclose(loaded['base']['npv_rub'], sum(r['present_value_rub'] for r in loaded['yearly_scenarios']['base']))
    assert math.isclose(sum(x[2] for x in EQUIPMENT),7000000)
    assert math.isclose(sum(SEASON),1)
    print(json.dumps({'verified': True, 'base': summary['base'], 'sensitivity': sensitivity,
                      'cash_runway': summary['cash_runway']},ensure_ascii=False,indent=2))


if __name__ == '__main__':
    import sys
    suite = unittest.defaultTestLoader.loadTestsFromTestCase(FinanceTests)
    if not unittest.TextTestRunner(verbosity=2).run(suite).wasSuccessful():
        sys.exit(1)
    if '--test' not in sys.argv:
        build()
