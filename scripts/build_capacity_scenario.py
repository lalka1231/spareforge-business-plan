"""Explicit one-crew resource stress; does not overwrite base model."""
import json
from pathlib import Path
from build_finance import model, metrics, SEASON, write_csv
ROOT=Path(__file__).resolve().parents[1]
def build():
 rows=model(); monthly=[]
 for row in rows[1:]:
  required=row['treatment_ha_passes']/130+row['monitoring_ha_surveys']/500
  retained=0
  for month,share in enumerate(SEASON,1):
   days=required*share
   scale=min(1,20/days) if days else 1
   retained+=share*scale
   monthly.append(dict(year=row['year'],month=month,planned_field_days=days,available_field_days_assumed=20,retained_fraction=scale,delivered_field_days=days*scale,treatment_ha_passes=row['treatment_ha_passes']*share*scale,monitoring_ha_surveys=row['monitoring_ha_surveys']*share*scale))
  for key in ('treatment_ha_passes','monitoring_ha_surveys','treatment_revenue_rub','monitoring_revenue_rub','revenue_rub','variable_cost_rub'):
   row[key]*=retained
  row['contribution_rub']=row['revenue_rub']-row['variable_cost_rub']
  row['ebitda_rub']=row['contribution_rub']-row['fixed_opex_rub']
  row['usn_tax_rub']=row['revenue_rub']*.06
  row['ebit_rub']=row['ebitda_rub']-row['depreciation_rub']
  row['accounting_profit_rub']=row['ebit_rub']-row['usn_tax_rub']
  row['free_cash_flow_rub']=row['ebitda_rub']-row['usn_tax_rub']-row['additional_capex_rub']+row['reserve_return_rub']
  row['present_value_rub']=row['free_cash_flow_rub']*row['discount_factor']
 # VAT is assumed passed through, not revenue or project cash flow. Recompute timing.
 previous=0
 for row in rows[1:]:
  threshold=row['vat_threshold_assumed_rub']; active=previous>threshold; cumulative=vat=0
  delivered=[m for m in monthly if m['year']==row['year']]
  for m in delivered:
   revenue=m['treatment_ha_passes']*row['treatment_price_ex_vat_rub']+m['monitoring_ha_surveys']*row['monitoring_price_ex_vat_rub']
   if active:vat+=revenue*.05
   cumulative+=revenue
   if cumulative>threshold:active=True
  previous=row['revenue_rub'];row['vat_collected_rub']=row['vat_remitted_rub']=vat
  row['customer_billings_including_vat_rub']=row['revenue_rub']+vat
 cash=discounted=0
 for row in rows:
  cash+=row['free_cash_flow_rub'];discounted+=row['present_value_rub']
  row['cumulative_cf_rub']=cash;row['cumulative_discounted_cf_rub']=discounted
 for m in monthly:assert m['delivered_field_days']<=20+1e-8
 write_csv('capacity_scenario_monthly.csv',monthly)
 write_csv('capacity_scenario_annual.csv',rows)
 result={'classification':'[ДОПУЩЕНИЕ] 20 эффективных полевых дней в месяц; не наблюденная погода и не гарантия мощности','formula':'d=(treatment/130+monitoring/500)*month_share; k=min(1,20/d); both monthly volumes multiplied by k','limitations':['Одна бригада; резервный борт не увеличивает мощность','Не выполненный в месяце объем не переносится; это потеря реализации','Постоянные расходы, замены и возврат резерва сохранены','Налоги и полный перенос НДС клиенту остаются допущениями базовой модели','Не доказаны логистика, агрономические окна и фактическая производительность'],**metrics(rows)}
 (ROOT/'data/capacity_scenario_summary.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
 print(json.dumps(result,ensure_ascii=False,indent=2))
if __name__=='__main__':build()
