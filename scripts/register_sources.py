"""Register retrieved evidence before drafting source-bearing prose."""
import json,subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
TOOL=Path.home()/'.hermes/skills/research/grounded-citations/scripts/sources.py'
LEDGER=ROOT/'data/source_ledger.json'
def run(*args):
    r=subprocess.run([sys.executable,str(TOOL),'--ledger',str(LEDGER),*args],text=True,capture_output=True)
    if r.returncode:raise RuntimeError(r.stderr+r.stdout)
    return r.stdout.strip()
if __name__=='__main__':
    evidence=ROOT/'docs/evidence';evidence.mkdir(exist_ok=True)
    mapping=json.loads((ROOT/'data/source_map.json').read_text()) if (ROOT/'data/source_map.json').exists() else []
    for filename in sys.argv[1:]:
        records=json.loads(Path(filename).read_text())
        if isinstance(records,dict):records=records.get('sources',records.get('results',[]))
        for rec in records:
            idout=run('add',rec['url'],'--title',rec['title']);print(idout)
            import re
            match=re.search(r'\[(\d+)\]',idout)
            if match is None: raise ValueError(idout)
            sid=int(match.group(1))
            file=evidence/f'{sid:02d}.txt';file.write_text(rec['content'],encoding='utf-8')
            print(run('quote',str(sid),'--text',rec['quote'],'--from',str(file)))
            file.write_text(rec['title']+'\n'+rec['url']+'\nПроверено: 2026-09-05\n\n'+rec['quote']+'\n',encoding='utf-8')
            mapping=[m for m in mapping if m['id']!=sid]
            mapping.append({'id':sid,'title':rec['title'],'url':rec['url'],'quote':rec['quote']})
    (ROOT/'data/source_map.json').write_text(json.dumps(mapping,ensure_ascii=False,indent=2),encoding='utf-8')
    (ROOT/'docs/09_SOURCES.md').write_text('# Реестр источников АгроВектор\n\nДата проверки: 05.09.2026. Цитаты ниже подтверждают только указанное содержание; рекламные показатели не равны независимому аудиту. Плановые числа проекта хранятся отдельно в assumptions.csv.\n\n'+run('render','--style','evidence')+'\n',encoding='utf-8')
