"""Independent Decimal read-back; does not import either generator."""
import csv,json
from decimal import Decimal as D
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
def read(name):return list(csv.DictReader((ROOT/'data'/name).open()))
base=read('financial_model_5y.csv'); actual=read('capacity_scenario_annual.csv')
shares=list(map(D,['0','0','0','.10','.20','.25','.20','.15','.10','0','0','0']))
npv=D('-9500000')
for i,(b,a) in enumerate(zip(base[1:],actual[1:])):
 t=D(b['treatment_ha_passes']);m=D(b['monitoring_ha_surveys'])
 needed=t/D(130)+m/D(500)
 ksum=sum(s*min(D(1),D(20)/(s*needed)) for s in shares if s)
 revenue=(t*D(b['treatment_price_ex_vat_rub'])+m*D(b['monitoring_price_ex_vat_rub']))*ksum
 variable=(t*D(b['treatment_variable_unit_rub'])+m*D(b['monitoring_variable_unit_rub']))*ksum
 ebitda=revenue-variable-D(b['fixed_opex_rub'])
 cf=ebitda-revenue*D('.06')-D(b['additional_capex_rub'])+D(b['reserve_return_rub'])
 for key,value in [('treatment_ha_passes',t*ksum),('monitoring_ha_surveys',m*ksum),('revenue_rub',revenue),('ebitda_rub',ebitda),('free_cash_flow_rub',cf)]:
  assert abs(D(a[key])-value)<D('.01'),(b['year'],key,a[key],value)
 npv+=cf/D('1.24')**(i+1)
s=json.loads((ROOT/'data/capacity_scenario_summary.json').read_text())
assert abs(D(str(s['npv_rub']))-npv)<D('.01')
for r in read('capacity_scenario_monthly.csv'):
 assert D(r['delivered_field_days'])<=D('20.00000001')
print(json.dumps({'verified':True,'independent_npv_rub':str(npv),'annual_rows':len(actual),'monthly_rows':len(read('capacity_scenario_monthly.csv'))}))
