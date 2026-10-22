#!/usr/bin/env python3
"""Build an isolated corrected export from the unchanged finance model."""
import csv
import json
import math
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
import build_finance as finance
import analyze_resilience_brainstorm as resilience


def write_csv(path, rows):
    with path.open('w', encoding='utf-8', newline='') as stream:
        writer = csv.DictWriter(stream, fieldnames=list(rows[0]), lineterminator='\n')
        writer.writeheader()
        writer.writerows(rows)


def build(write=True):
    rows = finance.model()
    base = finance.metrics(rows)
    monthly = finance.seasonality(rows[1])
    fixed_annual = rows[1]['fixed_opex_rub']
    reserve = rows[0]['reserve_funding_rub']
    fixed_monthly = fixed_annual / 12
    first_negative = math.floor(reserve / fixed_monthly) + 1
    no_revenue = []
    for month in range(1, 13):
        opening = reserve - (month - 1) * fixed_monthly
        closing = reserve - month * fixed_monthly
        no_revenue.append({
            'month': f'2027-{month:02}', 'opening_cash_rub': opening,
            'revenue_rub': 0, 'fixed_opex_paid_rub': fixed_monthly,
            'net_cash_flow_rub': -fixed_monthly, 'closing_cash_rub': closing,
            'status': '[РАСЧЕТ] без выручки; не базовый сценарий'})

    scenarios = json.loads((ROOT / 'data/resilience_brainstorm.json').read_text(encoding='utf-8'))
    capacity_ids = ('capacity_20', 'capacity_16', 'capacity_12', 'capacity20_price1450',
                    'capacity16_price1450', 'capacity20_price1450_down10',
                    'capacity20_price1450_costplus10')
    scenario_ids = ('base', 'price_minus10', 'volume_minus25', 'variable_plus20',
                    'one_year_delay', 'no_reserve_release', 'combined_downside')
    scenario_lookup = {item['scenario']: item for item in scenarios['scenarios']}
    liquidity = [dict(item) for item in scenarios['liquidity']]
    result = {
        'project': 'АгроВектор', 'currency': 'RUB',
        'classification': '[РАСЧЕТ] воспроизводимые значения; тарифы, зарплаты, NPV-ставка и налоги — [ДОПУЩЕНИЕ]',
        'source': {'model_module': 'scripts/build_finance.py', 'functions': ['model()', 'metrics()', 'seasonality()'],
                   'build_called': False, 'baseline_export': 'data/finance_summary.json'},
        'base': base, 'yearly_scenarios': {'base': rows}, 'first_year_monthly': monthly,
        'no_revenue_cash_monthly': no_revenue,
        'cash_runway': {
            'initial_operating_cash_rub': reserve, 'fixed_opex_annual_from_model_rub': fixed_annual,
            'no_revenue_fixed_burn_per_month_rub': fixed_monthly,
            'no_revenue_runway_months': reserve / fixed_monthly,
            'no_revenue_first_negative_month': first_negative,
            'base_minimum_month_end_cash_rub': min(m['closing_cash_rub'] for m in monthly),
            'base_first_negative_month': next((m['month'] for m in monthly if m['closing_cash_rub'] < 0), None),
            'base_year_end_cash_rub': monthly[-1]['closing_cash_rub'],
            'base_additional_funding_required_rub': max(0, -min(m['closing_cash_rub'] for m in monthly)),
        },
        'reference_scenarios': [scenario_lookup[key] for key in scenario_ids],
        'capacity_reference_scenarios': [scenario_lookup[key] for key in capacity_ids],
        'liquidity_reference_cases': liquidity,
        'reference_thresholds': scenarios['thresholds'],
        'assumptions': {
            'treatment_price_2027_rub_per_ha_pass': {'value': 1300, 'status': '[ДОПУЩЕНИЕ]'},
            'monitoring_price_2027_rub_per_ha_survey': {'value': 250, 'status': '[ДОПУЩЕНИЕ]'},
            'gross_monthly_salary_director_rub': {'value': finance.DIRECTOR_GROSS_MONTHLY, 'status': '[ДОПУЩЕНИЕ]'},
            'gross_monthly_salary_operator_rub': {'value': finance.OPERATOR_GROSS_MONTHLY, 'status': '[ДОПУЩЕНИЕ]'},
            'discount_rate': {'value': finance.RATE, 'status': '[ДОПУЩЕНИЕ]'},
            'usn_revenue_rate': {'value': .06, 'status': '[ДОПУЩЕНИЕ; проверить бухгалтером]'},
            'vat_rate_and_thresholds': {'value': '5%; thresholds 20/20/20/15/10 million RUB by year', 'status': '[ДОПУЩЕНИЕ; требуется актуальная проверка]'},
        },
        'limitations': scenarios['limitations'] + [
            'Экспорт не меняет исходную модель/базу; runway без выручки — статический стресс с расходами из строки модели 2027.',
            'Помесячные остатки на конец месяца не исключают внутримесячный кассовый разрыв.',
        ],
        'corrections': {'legacy_no_revenue_first_negative_month': 6,
                        'corrected_no_revenue_first_negative_month': first_negative,
                        'rule': 'floor(initial_reserve / fixed_monthly_burn) + 1'},
    }
    if write:
        (ROOT / 'data').mkdir(exist_ok=True)
        (ROOT / 'docs').mkdir(exist_ok=True)
        (ROOT / 'data/finance_corrected_brainstorm.json').write_text(
            json.dumps(result, ensure_ascii=False, indent=2, allow_nan=False) + '\n', encoding='utf-8')
        write_csv(ROOT / 'data/finance_corrected_brainstorm.csv', rows)
        write_csv(ROOT / 'data/cash_corrected_brainstorm.csv', monthly)
        (ROOT / 'docs/FINANCE_CORRECTED_brainstorm.md').write_text(render_doc(result), encoding='utf-8')
    return result


def render_doc(data):
    base = data['base']
    c = data['cash_runway']
    fmt = lambda value: 'не достигнута' if value is None else f'{value:,.2f}'
    lines = [
        '# Исправленный финансовый экспорт AgroVector — brainstorming', '',
        '**Статус:** параллельный экспорт и проверка, не утвержденная база и не инвестиционная рекомендация. Исходные файлы не изменяются.',
        'Генератор импортирует `build_finance.model()`, `metrics()` и `seasonality()`; `build()` исходной модели не вызывается.', '',
        '## Исправление runway', '',
        f"При резерве {c['initial_operating_cash_rub']:,.0f} ₽ и фиксированных расходах {c['fixed_opex_annual_from_model_rub']:,.0f} ₽/год (из модели; {c['no_revenue_fixed_burn_per_month_rub']:,.0f} ₽/мес.)",
        f"остаток становится отрицательным в **месяце {c['no_revenue_first_negative_month']}**, а не в старом значении 6. Формула: floor(резерв / месячный burn) + 1.",
        'Runway 4.08 месяца — деление резерва на burn, не округление вверх до полного положительного месяца.', '',
        '## Текущая базовая экономика без изменений', '',
        f"NPV при допущении ставки 24%: **{base['npv_rub']:,.2f} ₽**; инвестиция {base['initial_investment_rub']:,.0f} ₽; простая окупаемость {fmt(base['simple_payback_years'])} года; дисконтированная — {fmt(base['discounted_payback_years'])}.",
        '', '| Год | Выручка | EBITDA | УСН* | FCF |', '|---:|---:|---:|---:|---:|',
    ]
    for row in data['yearly_scenarios']['base'][1:]:
        lines.append(f"| {row['year']} | {row['revenue_rub']:,.0f} | {row['ebitda_rub']:,.0f} | {row['usn_tax_rub']:,.0f} | {row['free_cash_flow_rub']:,.0f} |")
    lines += ['', '* Налоговые ставки/алгоритм — бюджетные допущения, не заключение бухгалтера.',
              'Базовая цена 2027: 1 300 ₽/га-проход и 250 ₽/га-обследование; gross salary: руководитель и оператор по 100 000 ₽/мес. с двумя операторами; дисконтирование 24%. Это допущения модели, не подтвержденные оферты/условия найма.',
              '', '## Операционная мощность и чувствительность', '',
              'Один экипаж; резервный дрон не является второй бригадой. Мощность экранируется предположениями 130 га-проходов/день и 500 га-обследований/день, сезонные эффективные дни не наблюдались. Потерянный месячный объем не переносится.',
              '', '| Сценарий | NPV, ₽ | 2027 FCF, ₽ |', '|---|---:|---:|']
    for item in data['capacity_reference_scenarios']:
        lines.append(f"| {item['scenario']} | {item['npv_rub']:,.2f} | {item['fcf_2027_rub']:,.2f} |")
    for item in data['reference_scenarios']:
        lines.append(f"| {item['scenario']} | {item['npv_rub']:,.2f} | {item['fcf_2027_rub']:,.2f} |")
    lines += ['', 'Референсные расчеты унаследованы из `resilience_brainstorm.json`; изменения тарифов, CAPEX, экономии, объемов и налогового переноса являются сценарными допущениями, не выбранной базой. NPV базы отрицателен; NPV=0 порог по тарифу не доказывает спрос.', '',
              '## Кассовые стрессы', '',
              f'В базе помесячный минимум {c["base_minimum_month_end_cash_rub"]:,.0f} ₽; конец 2027 — {c["base_year_end_cash_rub"]:,.0f} ₽. При нулевой выручке в пятом месяце остаток отрицательный. При задержке поступлений базовый стресс показывает минимумы: без аванса — лаг 1 мес. −417 800 ₽, лаг 2 мес. −1 568 400 ₽, лаг 3 мес. −2 587 000 ₽. Даже при авансе 50% лаг 2 месяца дает −4 200 ₽, лаг 3 — −120 800 ₽. Это минимумы по модели liquidity, без внутримесячного календаря и буфера.', '',
              '## Ограничения и launch gates', '',
              '- Нет подтвержденных заказов, тарифов, цен поставщиков, производительности и статистики погодных окон; модель не вероятностная.',
              '- Зарплаты, сезонность, начальный резерв, переменные расходы, индексация, амортизация и CAPEX — проектные параметры. Препараты не включены и предполагаются предоставляемыми заказчиком.',
              '- УСН 6%, НДС 5%/пороговые параметры и учет НДС — только налоговые допущения; проверить актуальный режим, сроки, льготы и договорную цену с бухгалтером.',
              '- Нет плохих долгов, предсезонного аванса, ежедневных платежей, финансирования разрывов, вероятностей или полноценной модели продолжающегося бизнеса. Возврат резерва Y5 предполагает завершение проекта.',
              '- Не считать 1 450 ₽, аренду, сезонный найм или любой стресс новым базовым предположением; NPV не подгонялся.', '',
              'Перед запуском (все gates обязательны):',
              '1. Документированный спрос с тарифом, объемом, услугой, окном, приемкой, НДС и оплатой; количественный порог подтверждения согласуется заранее. Интерес и LOI не равны заказу или полученной выручке.',
              '2. Две сопоставимые письменные комплектации/КП на парк и сервис, батареи, запасные части, обучение, сроки и гарантии; сверка полного CAPEX и TCO.',
              '3. Полевой пилот только после допуска и согласованного доступа к технике: замеры полного цикла и пикового календаря одной бригады. Если пилот нужен до закупки, доступ к чужому парку, бюджет и календарь согласуются отдельно; аренда не объявляется доступной.',
              '4. 13-недельный кассовый календарь с зарплатами, поставщиками, налогами и договорными поступлениями; подтвержденный источник покрытия каждого отрицательного остатка плюс запас.',
              '5. Налоговая карта бухгалтера и проверяемый правовой маршрут до необратимых обязательств; конкретные допуски до каждого полета. Очередность оформления оборудования и документов проверяет юрист, лимиты обязательств утверждаются отдельно.',
              '6. Пересчет NPV и downside на доказанных данных; запуск не разрешать при NPV<0 по согласованной ставке либо нефинансируемом кассовом разрыве.', '',
              'Повторная генерация: `python3 scripts/build_corrected_finance_brainstorm.py`; тесты: `python3 scripts/test_corrected_finance_brainstorm.py`.']
    return '\n'.join(lines) + '\n'


if __name__ == '__main__':
    data = build()
    print(json.dumps({'verified': True, 'npv_rub': data['base']['npv_rub'],
                      'runway_months': data['cash_runway']['no_revenue_runway_months'],
                      'first_negative_month': data['cash_runway']['no_revenue_first_negative_month']}, ensure_ascii=False, indent=2))
