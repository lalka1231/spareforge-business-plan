"""Generate references from stable local ledger IDs cited in the report body."""
from pathlib import Path
import json,re
ROOT=Path(__file__).resolve().parents[1]
# Resolve the current generated sections, not the previous Markdown export.
text='\n'.join(json.loads((ROOT/'data/sections.json').read_text()).values())
ids={int(i) for i in re.findall(r'\[(\d+)\]',text)}
ledger={s['id']:s for s in json.loads((ROOT/'data/source_ledger.json').read_text())['sources']}
assert ids<=ledger.keys()
block='## Sources\n\n'+'\n'.join(f"[{i}] {ledger[i]['url']} — {ledger[i]['title']}" for i in sorted(ids))+'\n'
(ROOT/'data/references_block.md').write_text(block,encoding='utf-8')
print('Reference IDs:',','.join(map(str,sorted(ids))))
