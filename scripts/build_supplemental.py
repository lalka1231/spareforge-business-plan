"""Supplemental scenarios, receivables and resource checks from existing model."""
from pathlib import Path
import csv,json,calendar
from build_finance import model,metrics,seasonality,write_csv,SEASON
ROOT=Path(__file__).resolve().parents[1]
def build():
 scenarios={
  'alternative_sam_mix':('Мониторинг = внесение / 3 во все пять лет; не базовая выручка',dict(monitoring_ratio=1/3)),
  'procurement_2990000':('Два публичных ценника по 2,99 млн вместо лимита 4 млн; прочее неизменно; не полное КП',dict(initial_capex=8980000)),
  'combined_price_volume_procurement':('Оба тарифа −10%; оба объема −25%; CAPEX +1,98 млн; исходный горизонт',dict(price_factor=.9,volume_factor=.75,initial_capex=8980000))}
 result={};annual=[]
 for key,(desc,kwargs) in scenarios.items():
  rr=model(**kwargs);result[key]={'description':desc,'parameters':kwargs,**metrics(rr),'first_year':rr[1]}
  annual.extend(dict(scenario=key,**r) for r in rr)
 write_csv('supplemental_model.csv',annual)
 monthly=seasonality(model()[1]);cashstress=[];delay_summary=[]
 for lag in (1,2,3):
  cash=2000000;minimum=cash
  for i,m in enumerate(monthly):
   receipts=monthly[i-lag]['revenue_collected_ex_vat_rub'] if i>=lag else 0
   net=receipts*.94-m['variable_cost_paid_rub']-m['fixed_opex_paid_rub']
   cash+=net;minimum=min(minimum,cash)
   cashstress.append(dict(delay_months=lag,month=m['month'],receipts_rub=receipts,usn_rub=receipts*.06,variable_rub=m['variable_cost_paid_rub'],fixed_rub=m['fixed_opex_paid_rub'],net_cf_rub=net,closing_cash_rub=cash))
  delay_summary.append(dict(delay_months=lag,min_cash_rub=minimum,additional_liquidity_rub=max(0,-minimum),year_end_cash_rub=cash))
 write_csv('collection_delay.csv',cashstress)
 capacity=[]
 for r in model()[1:]:
  for month,share in enumerate(SEASON,1):
   if not share:continue
   days=(r['treatment_ha_passes']/130+r['monitoring_ha_surveys']/500)*share
   cal=calendar.monthrange(r['year'],month)[1]
   # 5/7 is a planning screen, not a legal determination or observed weather.
   capacity.append(dict(year=r['year'],month=month,share=share,required_field_days=days,calendar_days=cal,planning_workdays_5_of_7=cal*5/7,excess_vs_screen_days=max(0,days-cal*5/7),status='[РАСЧЕТ] на допущениях; погодный запас не включен; ресурсная выполнимость не подтверждена'))
 write_csv('capacity_check.csv',capacity)
 a={r['parameter']:float(r['value']) for r in csv.DictReader((ROOT/'data/assumptions.csv').open()) if r['value'].replace('.','',1).replace('-','',1).isdigit()}
 basket=a['farm_area']*a['eligible_share']*(a['annual_passes_eligible']*a['treatment_price_2027']+a['annual_monitoring_eligible']*a['monitoring_price_2027'])
 market={'classification':'[РАСЧЕТ] национальный сценарий; не установленная емкость рынка','tam_rub':a['tam_russia_sown_area_2025']*a['eligible_share']*(a['annual_passes_eligible']*a['treatment_price_2027']+a['annual_monitoring_eligible']*a['monitoring_price_2027']),'sam_rub':a['sam_farms']*basket,'som_rub':model()[1]['revenue_rub'],'alternative_y1_rub':result['alternative_sam_mix']['first_year']['revenue_rub']}
 output={'revision':'2026-10-20','scenarios':result,'collection_delay':delay_summary,'market':market,'resource_warning':'Июнь 2031 требует 30 полевых дней из 30 календарных; отдых и погода не обеспечены. Поздняя база — непроверенный верхний сценарий.','limitations':['Закупочный тест не переоценивает сервис и последующие замены; комплектность и 60/50 л требуют КП','Отсрочка без авансов; УСН резервируется при поступлении; месячные остатки не показывают внутримесячные разрывы','Пять рабочих дней из семи — условный ресурсный экран; не утверждение о фактическом режиме труда']}
 (ROOT/'data/supplemental_summary.json').write_text(json.dumps(output,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
 # Independent read-back identities, valid under the current payroll scenario.
 base_rows=model();base_npv=sum(r['free_cash_flow_rub']/(1.24**r['period']) for r in base_rows)
 assert abs(result['procurement_2990000']['npv_rub']-(base_npv-1980000))<.02
 assert abs(result['alternative_sam_mix']['first_year']['revenue_rub']-8300000)<.02
 assert result['alternative_sam_mix']['npv_rub']<base_npv
 assert result['combined_price_volume_procurement']['npv_rub']<base_npv-1980000
 # Recompute delayed receipts without reusing the cashstress loop.
 receipts=[0,0]+[r['revenue_collected_ex_vat_rub'] for r in monthly[:-2]]
 cash=2000000;balances=[]
 for got,m in zip(receipts,monthly):
  cash+=got*.94-m['variable_cost_paid_rub']-m['fixed_opex_paid_rub'];balances.append(cash)
 assert abs(delay_summary[1]['min_cash_rub']-min([2000000]+balances))<.02
 assert abs(market['tam_rub']-24681502500)<.02
 print(json.dumps(output,ensure_ascii=False,indent=2))
if __name__=='__main__':build()
