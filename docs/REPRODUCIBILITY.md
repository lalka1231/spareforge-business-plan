> **Редакция 21.10.2026:** актуальны [карта исполнения](REFERENCE_REVISION_2026_10_21.md), [README](../README.md) и `presentations/REFERENCE_REVISION_ACCEPTANCE.json`. Нижележащие сведения о прежних числах и протоколах — история; новая зарплатная база дает отрицательныйNPV.

# Воспроизводимость AgroVector — 20.10.2026

## Проверенная среда и сборка

Python 3.11+, `uv`; зависимости из существующих requirements. В текущем запуске проверены Python 3.11.16, python-pptx 1.0.2, python-docx 1.2.0, markdown-it-py 4.2.0, Pillow 12.3.0. Финансовые генераторы используют стандартную библиотеку. Сборка выполняется из корня существующего проекта:

```bash
uv venv .venv
uv pip install --python .venv/bin/python -r requirements.txt -r documents/requirements-docx.txt
.venv/bin/python scripts/build_finance.py
.venv/bin/python scripts/build_supplemental.py
.venv/bin/python scripts/build_capacity_scenario.py
.venv/bin/python scripts/build_competitors.py
.venv/bin/python scripts/build_plan.py
.venv/bin/python scripts/build_content.py
.venv/bin/python scripts/build_documents.py
.venv/bin/python scripts/build_references.py
.venv/bin/python scripts/build_documents.py
.venv/bin/python scripts/build_docx.py
.venv/bin/python scripts/build_presentation.py
.venv/bin/python scripts/verify.py
.venv/bin/python scripts/verify_capacity_scenario.py
```

Два вызова build_documents.py необходимы: первый формирует тело с текущими цитатами, build_references.py вычисляет список ссылок из стабильного локального ledger, второй подставляет новый список. Финансовые результаты не вводятся вручную в DOCX/PPTX. `build_competitors.py` формирует только экспорт единого CSV; Word/Markdown и слайдовая проекция используют тот же модуль.

Повторный тест модели без записи: `.venv/bin/python scripts/build_finance.py --test`. Основная проверка не импортирует финансовый генератор: независимо читает CSV с Decimal, сверяет годовые формулы, базовый и дополнительные NPV, помесячные деньги, TAM/SAM, SMART/Гант, 22 раздела/слайда и все ячейки полной конкурентной матрицы. У нее отдельный структурный статус, не фиктивное подтверждение визуального контроля.

## Реальный PDF-рендер

В этой среде найден LibreOffice `/opt/homebrew/bin/soffice`; Poppler отсутствует. Для растрирования используется PyMuPDF, не имитация страниц средствами python-docx:

```bash
uv pip install --python .venv/bin/python -r requirements-render.txt
.venv/bin/python scripts/render_review.py
```

Скрипт запускает LibreOffice с отдельным профилем, экспортирует настоящие DOCX/PPTX в PDF, создает PNG всех страниц и контактные листы в `review/2026-10-20/`; сохраняет фактические числа страниц и SHA исходников в `data/render_validation.json`. Проверка текста за краями страницы не доказывает отсутствие внутренних наложений. Контактные листы проверяют общую пагинацию, плотные таблицы и слайды дополнительно просматриваются отдельно. Результат и пределы — `docs/RECONCILIATION_2026_10_20.md`.

Шрифт диаграмм берется из macOS Arial по пути в build_docx.py; на иной ОС необходимо явно выбрать доступный кириллический шрифт. Не запускать исторический revise_teacher_feedback.mjs: он не входит в сборку и может вернуть старые таблицы. render_contact_sheets.py — прежний скрипт с жестким каталогом; актуальный render_review.py от него не зависит.

## Данные и доказательства

- `build_finance.py` — финансовые формулы, экспорт допущений, основной пятилетний сценарий и пять стандартных стрессов.
- `build_supplemental.py` — альтернативная структура услуг, закупочный и комбинированный стрессы, отсрочка оплаты и ресурсный экран. Поздняя мощность не подтверждена; параметры экрана не юридический режим труда.
- `narrative_sections.json` — редактируемая нефинансовая основа; `pitch_standard.json` — SMART, без дублирующей конкурентной базы.
- `competitors.csv` — единственный реестр шести участников, критериев, источников и дат. Слайд 6 сокращен для читаемости; полная матрица в его заметках и Word/Markdown.
- `source_ledger.json` / `source_map.json` / `docs/evidence/` — стабильные ID, короткие выдержки и индивидуальные даты. `build_references.py` не выполняет новый интернет-фактчекинг.
- `register_revision_sources.py` фиксирует конкретные результаты чтения 20.10.2026; `revise_20261020.py` — историческая миграция после синтеза A–E. Это не ежедневное автообновление источников и не часть обычного пересчета.

При наличии Hermes дополнительная проверка:

```bash
.venv/bin/python ~/.hermes/skills/research/grounded-citations/scripts/sources.py --ledger data/source_ledger.json verify docs/FINAL_BUSINESS_PLAN.md --evidence
```

## Открытые границы

Дата редакции не равна дате проверки всех источников. Коммерческие страницы проверены 20.10.2026 с указанными ограничениями; главная BeeUAV не прочитана. Новый официальный материал Россельхознадзора и карточка УП-813 прочитаны; полный актуальный официальный текст ФАП/СП, исключения регионального режима и налоговые первоисточники полностью не перепроверены. Архивные снимки не объявлены свежими. Маркетинговые заявления не подтверждают качество, выручку или разрешения.

Нет подтвержденных заказов, капитала, полного КП, производительности и календаря траншей. Сценарные TAM/SAM и годовой NPV не заменяют эти доказательства. Точное соответствие двум фотографиям не проверено — сами изображения не были доступны. Ни сборка, ни проверки не публикуют файлы и не изменяют GitHub.
