# Воспроизводимость АгроВектора

## Среда

Python 3.11+, `python-pptx==1.0.2`. Финансовый генератор использует только стандартную библиотеку. Для независимой проверки PPTX и сборки презентации установите зависимости в изолированную среду:

```bash
uv venv .venv
uv pip install --python .venv/bin/python -r requirements.txt
.venv/bin/python scripts/build_finance.py
.venv/bin/python scripts/build_plan.py
.venv/bin/python scripts/build_content.py
.venv/bin/python scripts/build_documents.py
.venv/bin/python scripts/build_presentation.py
.venv/bin/python scripts/verify.py
```

`build_finance.py` — источник параметров и формул; он пересоздает CSV, месячный ДДС и JSON. `data/assumptions.csv` — экспорт параметров, а не независимый калькулятор. `data/narrative_sections.json` — редактируемая нефинансовая часть. `build_content.py` добавляет текущие финансовые таблицы. `build_documents.py` собирает 22 раздела и вычисленный график. Заметки слайдов обновляются из итогового плана.

`data/references_block.md` — зафиксированный блок ссылок, сформированный из citation ledger по цитатам итогового документа. При добавлении новых внешних утверждений нужно сначала обновить реестр и доказательства, затем этот блок; простой пересчет финансов не меняет источников.

При наличии Hermes дополнительная проверка происхождения:

```bash
python ~/.hermes/skills/research/grounded-citations/scripts/sources.py --ledger data/source_ledger.json verify docs/FINAL_BUSINESS_PLAN.md --evidence
```

`register_sources.py` — вспомогательный импорт собранных исследований для среды Hermes, не необходим для пересчета готовой модели. Он проверяет точные выдержки до сохранения, а в репозиторий помещает только короткие цитаты.

## Ограничения

Источники проверены на 05.09.2026. До инвестиционного решения требуется обновление правового и коммерческого due diligence. Публичная цена не подтверждает доступность услуги в регионе; бюджет не является КП.

PowerPoint проверен структурно. Для настоящего рендеринга в PDF/PNG нужны LibreOffice и Poppler; в текущей среде они отсутствуют. Не считать проверку координат гарантией визуального результата в любом редакторе.
