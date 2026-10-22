# AgroVector — исправленная редакция после brainstorming

**Ветка:** `AgroVector_After_brainstorm`. Родитель: `brainstorm_version`, коммит `428ed6b1ea89876a6a910299f843efaa8507062e`. Утвержденная историческая база: `7c05732803c7bedff0ce15f35647195e61de0790`.

Это реализация согласованного пакета **документальных исправлений, расчетных ограничений и условий запуска без подмены базовых допущений**. Не подтверждение готовности коммерческого бизнеса. В этой ветке новые материалы имеют суффикс `_brainstorm`; исходные файлы и исторический brainstorm сохранены как контрольные снимки.

## Актуальные материалы этой редакции

| Материал | Файл |
|---|---|
| Полный письменный бизнес-план: 22 раздела | [FINAL_BUSINESS_PLAN_brainstorm.md](docs/FINAL_BUSINESS_PLAN_brainstorm.md) |
| Полный Word с расчетами и приложениями | [AGROVECTOR_FULL_BUSINESS_PLAN_brainstorm.docx](documents/AGROVECTOR_FULL_BUSINESS_PLAN_brainstorm.docx) |
| Презентация: 22 слайда, сохраненные Morph/анимации | [AGROVECTOR_AFTER_brainstorm.pptx](presentations/AGROVECTOR_AFTER_brainstorm.pptx) |
| PDF презентации — статическая проверка LibreOffice | [AGROVECTOR_AFTER_brainstorm.pdf](presentations/AGROVECTOR_AFTER_brainstorm.pdf) |
| PDF Word — экспорт LibreOffice | [AGROVECTOR_FULL_BUSINESS_PLAN_brainstorm.pdf](documents/AGROVECTOR_FULL_BUSINESS_PLAN_brainstorm.pdf) |
| Исправленный финансовый экспорт | [finance_corrected_brainstorm.json](data/finance_corrected_brainstorm.json) |
| Полный годовой поток / базовый месячный поток | [Годы](data/finance_corrected_brainstorm.csv) / [Месяцы](data/cash_corrected_brainstorm.csv) |
| Финансовые исправления, ограничения и сценарии | [FINANCE_CORRECTED_brainstorm.md](docs/FINANCE_CORRECTED_brainstorm.md) |
| Исправленная методика | [FINANCE_METHOD_brainstorm.md](docs/FINANCE_METHOD_brainstorm.md) |
| Решения по всем 37 записям риска | [CHANGE_DECISIONS_brainstorm.md](docs/CHANGE_DECISIONS_brainstorm.md) |
| Проверки и ограничения приемки | [AFTER_ACCEPTANCE_brainstorm.md](docs/AFTER_ACCEPTANCE_brainstorm.md) |

**Начинайте с этих файлов.** `README.md`, `README_brainstorm.md`, документы/модель без суффикса и прежний `AGROVECTOR_brainstorm.md` — исторические версии. Старые отчеты PASS не являются приемкой новой редакции. Не запускайте старые генераторы `build_finance.py` / `assemble_brainstorm.py` в режиме записи для пересборки новой версии.

## Что изменено

- Подпись фиксированных расходов: **490 600 ₽/месяц**, из годового OPEX модели.
- Отдельный стресс без выручки: **первый отрицательный остаток — в 5-м месяце**; непрерывный runway ≈4,08 месяца. Это не означает пять обеспеченных месяцев и не описывает базовую кассу с апрельскими поступлениями.
- Раскрыты незаполненные финансовые шаблоны `{fmt(...)}` в полном плане; исправлены формулировки недостигнутой окупаемости.
- Добавлены уточнения в каждый из 22 разделов, ресурсно-ценовые и кассовые стрессы, условия GO/NO-GO и статусы внешних проверок. Word и заметки PPTX синхронизированы с новым текстом.
- Сохранены первые два визуальных слайда и XML timing/transition всех 22 слайдов. Статический PDF не является доказательством плавности нативного показа.

## Что не менялось и не доказано

База: одна бригада, три сотрудника, тарифы 2027 **1 300 / 250 ₽**, CAPEX **7 млн ₽**, стартовая потребность **9,5 млн ₽**, ставка **24%**. NPV остается **−1 067 653,16 ₽**, EBITDA 2027 **272 800 ₽**, FCF 2027 **−255 200 ₽**. Поднятие цены до 1 450 ₽, аренда и сезонный найм — только варианты для проверки.

Реальные заказы, КП, налоговая/правовая применимость, квалификация и доступность людей, полевой хронометраж, финансирование и платежи не получены/не подтверждены этой редакцией. Согласование текста не закрывает внешние риски. Архивные внешние источники сохранены с прежними датами, не выданы за новое исследование.

## Воспроизведение

Из существующего окружения проекта:

```sh
.venv/bin/python scripts/build_corrected_finance_brainstorm.py
.venv/bin/python scripts/build_after_documents_brainstorm.py
.venv/bin/python scripts/build_after_deck_brainstorm.py
.venv/bin/python scripts/test_corrected_finance_brainstorm.py
.venv/bin/python scripts/test_after_documents_brainstorm.py
.venv/bin/python scripts/verify_after_brainstorm.py
```

Новые генераторы используют существующую модель и генератор Word. Word собирается в изолированном временном каталоге; записи в исторические файлы отсутствуют. После изменения Office-файлов PDF следует экспортировать заново реальным LibreOffice/Office и повторить визуальную проверку. Полный показ в PowerPoint остается отдельной проверкой.
