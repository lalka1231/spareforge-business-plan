# AgroVector — нативная анимация chart и отдельного PNG, исследование

## Статус и границы проверки

Проект прочитан, **не изменён**. PowerPoint-презентация не запускалась, не сохранялась и не редактировалась. **Native slideshow runtime НЕ проверен.** Ни XSD, ни компиляция AppleScript не доказывают визуальный рост столбцов/раскрытие секторов.

Результаты реальных проверок конечных примеров:

1. `/tmp/agrovector-columns.applescript` → `/tmp/agrovector-columns.scpt`: `osacompile`, exit 0. Команда AddEffect с Wipe, `chart by series elements`, `up`, duration принята установленным словарём.
2. `/tmp/agrovector-pie.applescript` → `/tmp/agrovector-pie.scpt`: `osacompile`, exit 0. AddEffect с Wheel, `chart by category` принята установленным словарём.
3. `/tmp/agrovector-columns-timing.xml`: PASS full ECMA transitional XSD после помещения timing в минимальный p:sld.
4. `/tmp/agrovector-pie-timing.xml`: PASS тем же комплектом XSD.

XML-проверки воспроизводятся `python3 /tmp/agrovector-check-xml.py`; схемы скачаны в `/tmp/agrovector-xsd/`. XSD не проверяет существование shape 42, соответствие данных series/category или исполнение эффекта. Эти XML — структурно валидные исследовательские образцы одного шага, НЕ PowerPoint-exported goldens и НЕ готовая runtime-проверенная замена всего slide timing.

## Что обнаружено в проекте

`scripts/build_cinematic_presentation.py`:
- В строках 77–93 создаются настоящие редактируемые chart через `add_chart`: слайд 17 `COLUMN_CLUSTERED`, слайд 18 `BAR_CLUSTERED` (горизонтальные полосы).
- Круговая диаграмма этим скриптом не создаётся.
- В строках 105–121 правится только media timing и push transition; поэлементной анимации chart здесь нет.
- Каждый слайд получает MP4. Для отдельного PNG-drone это надо менять в будущей реализации, но здесь ничего не менялось.
- Для горизонтального BAR на слайде 18 логичен Wipe вправо (`right`), не вверх. Wipe вверх подходит положительным вертикальным столбцам. Отрицательные значения требуют отдельного направления от нулевой линии; одинаковый Wipe вверх не обеспечивает физически корректный рост всех знаков.

## Главный вывод: правильные пространства имён

Для современного DrawingML chart путь цели — **`p:spTgt/p:graphicEl/a:chart`**. Не `p:chart` и не `c:chart`! `c:chart` в graphicData — ссылка на chart part, а `a:chart` в timing — адрес подэлемента анимации.

Аналогично метаданные построения:

```xml
<p:bldLst>
  <p:bldGraphic spid="42" grpId="0">
    <p:bldSub><a:bldChart bld="seriesEl" animBg="0"/></p:bldSub>
  </p:bldGraphic>
</p:bldLst>
```

**`p:bldChart` не является правильным элементом этого пути.** `p:bldOleChart` — отдельная модель старого OLE chart, не замена для native chart, созданного python-pptx. В `bldLst` нужны `p:bldGraphic` и вложенный **`a:bldChart`**.

Namespaces:

```xml
xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
```

## Конкретные элементы/значения

| Назначение | Значение |
|---|---|
| Entrance Wipe | `p:cTn presetID="22" presetClass="entr"` |
| Entrance Wheel | `p:cTn presetID="21" presetClass="entr"` |
| Исполняющий фильтр Wipe вверх | `p:animEffect transition="in" filter="wipe(up)"` |
| Исполняющий Wheel, одна спица | `p:animEffect transition="in" filter="wheel(1)"` |
| Вся серия | `a:chart bldStep="series" seriesIdx="0" categoryIdx="-1"` |
| Вся категория | `a:chart bldStep="category" seriesIdx="-1" categoryIdx="0"` |
| Один элемент, порядок по сериям | `a:chart bldStep="ptInSeries" seriesIdx="0" categoryIdx="0"` |
| Один элемент, порядок по категориям | `a:chart bldStep="ptInCategory" seriesIdx="0" categoryIdx="0"` |
| Фон/сетка/легенда | `bldStep="gridLegend"` |
| Все точки | `bldStep="allPts"` |

`seriesIdx/categoryIdx` — индексы подэлементов, не shape id; `-1` — schema default для невыбранной размерности. Для обычных последовательных данных python-pptx адресовать с нуля. При нестандартных c:idx/c:order, пропусках и комбинированных chart сверить реальную нативную сериализацию, не угадывать.

`a:bldChart/@bld`: `allAtOnce`, `series`, `category`, `seriesEl`, `categoryEl` (есть также общий вариант `one`). `seriesEl`: по одному point внутри серии; `categoryEl`: по одному point внутри категории. `series/category`: группа целиком. Для single-series pie `series` означает весь пирог, **не отдельные сектора**; нужен `category` либо point targeting.

`presetID` — метаданные эффекта. Без реального `animEffect` и правильной subtarget-цели они ничего не «выращивают». `presetSubtype` в образцах намеренно опущен: схема допускает это, а непроверенное число направления не выдаётся за установленный факт. Направление задаёт исполняющий filter; при нативном создании правильные metadata запишет PowerPoint.

## Образец XML: один вертикальный столбец

Полный XSD-проверенный timing находится в `/tmp/agrovector-columns-timing.xml`. Его существенная часть:

```xml
<p:cTn id="2" presetID="22" presetClass="entr" grpId="0"
       nodeType="withEffect" fill="hold">
  <p:stCondLst><p:cond delay="0"/></p:stCondLst>
  <p:childTnLst>
    <p:set>
      <p:cBhvr>
        <p:cTn id="3" dur="1" fill="hold">
          <p:stCondLst><p:cond delay="0"/></p:stCondLst>
        </p:cTn>
        <p:tgtEl><p:spTgt spid="42"><p:graphicEl>
          <a:chart bldStep="ptInSeries" seriesIdx="0" categoryIdx="0"/>
        </p:graphicEl></p:spTgt></p:tgtEl>
        <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
      </p:cBhvr>
      <p:to><p:strVal val="visible"/></p:to>
    </p:set>
    <p:animEffect transition="in" filter="wipe(up)">
      <p:cBhvr>
        <p:cTn id="4" dur="550" fill="hold"/>
        <p:tgtEl><p:spTgt spid="42"><p:graphicEl>
          <a:chart bldStep="ptInSeries" seriesIdx="0" categoryIdx="0"/>
        </p:graphicEl></p:spTgt></p:tgtEl>
      </p:cBhvr>
    </p:animEffect>
  </p:childTnLst>
</p:cTn>
```

Повторять effect-поддерево для каждой пары `(seriesIdx, categoryIdx)` с уникальными timing id, строить последовательность/задержки, затем добавить приведённый `bldLst`. Один приведённый шаг анимирует **один** столбец, не все. Наличие только `bldLst` не заменяет расписание исполняющих behaviors.

Обязательно одинаковая subtarget в `set` и `animEffect`: если `set` бездумно показывает chart целиком, получится раннее появление будущих элементов. Не скрывать shape chart целиком, если сетка должна быть видна сразу; будущие points должны получать entrance-семантику. Именно скрытие будущих элементов/первый кадр требуют отдельной runtime-проверки.

## Pie: сектор, а не whole-chart fade

Полный XSD-проверенный `/tmp/agrovector-pie-timing.xml` имеет ту же структуру, но:

```xml
<p:cTn id="2" presetID="21" presetClass="entr" grpId="0"
       nodeType="withEffect" fill="hold"> <!-- ... --> </p:cTn>
<!-- И в set, и в animEffect: -->
<p:tgtEl><p:spTgt spid="42"><p:graphicEl>
  <a:chart bldStep="category" seriesIdx="0" categoryIdx="0"/>
</p:graphicEl></p:spTgt></p:tgtEl>
<!-- Исполняющий эффект: -->
<p:animEffect transition="in" filter="wheel(1)"> <!-- cBhvr --> </p:animEffect>
<!-- Метаданные: -->
<p:bldLst><p:bldGraphic spid="42" grpId="0"><p:bldSub>
  <a:bldChart bld="category" animBg="0"/>
</p:bldSub></p:bldGraphic></p:bldLst>
```

Для следующего сектора меняется categoryIdx и timing ids. Для однозначного point targeting допустим также `ptInSeries`, seriesIdx=0, categoryIdx=k с `bld="seriesEl"`.

**Ограничение геометрии:** Wheel — вращающаяся маска раскрытия, не tween исходных чисел и не гарантированное наращивание угла каждого сектора от его собственного начала. Раскрытие может рассчитываться по общей области или области подэлемента — конкретная реализация Mac PowerPoint здесь не измерена. Wipe по сектору — линейное раскрытие сектора, а не радиальное заполнение. Если условие строго «каждый сектор плавно растёт вокруг центра пирога», нужен нативный reference deck и просмотр; нельзя честно обещать этот результат по одному валидному XML. Но category/point targeting — необходимое отличие от whole-chart fade.

## Предпочтительный путь на этом Mac: настоящий AddEffect из установленного sdef

Источник: `/Applications/Microsoft PowerPoint.app/Contents/Resources/PowerPoint.sdef`.

Словарь действительно содержит:
- `add effect`: direct parameter `sequence`, параметры `for` shape, `fx` MsoAnimEffect, `level` MsoAnimateByLevel, `trigger`, `index`; возвращает effect.
- `animation type wipe`, `animation type wheel`.
- levels `chart by series elements`, `chart by category`, `chart by series`.
- **Опечатку API** `chart by ctageory elements` — сохранять буквально, если использовать этот level. Исправленное написание не является именем из установленного словаря.
- `effect parameters.direction`, `timing.duration`, `timing.autoreverse`, `timing.repeat count`, `timing.repeat duration`.

Компилирующийся пример столбцов (не исполнялся):

```applescript
using terms from application "Microsoft PowerPoint"
 tell application "Microsoft PowerPoint"
  set s to slide 17 of active presentation
  set chartShape to shape "REPLACE_WITH_CHART_SHAPE_NAME" of s
  set seq to main sequence of timeline of s
  set fx to add effect seq for chartShape fx animation type wipe level chart by series elements trigger with previous
  set direction of effect parameters of fx to up
  set duration of timing of fx to 0.55
 end tell
end using terms from
```

Компилирующийся пример pie (не исполнялся; pie надо сначала реально создать):

```applescript
using terms from application "Microsoft PowerPoint"
 tell application "Microsoft PowerPoint"
  set s to slide 1 of active presentation
  set chartShape to shape "REPLACE_WITH_PIE_CHART_SHAPE_NAME" of s
  set seq to main sequence of timeline of s
  set fx to add effect seq for chartShape fx animation type wheel level chart by category trigger with previous
  set duration of timing of fx to 0.65
 end tell
end using terms from
```

Имена shape в примерах — явные placeholders; нельзя исполнять как есть или применять к активной презентации пользователя без проверки exact target. `shape N` в AppleScript — позиция в коллекции, не обязательно OOXML spid. В реальной реализации сначала получить имя/id нужного chart и работать с точным документом.

Преимущество AddEffect: PowerPoint сам строит своё расписание и build metadata; это безопаснее ручного сочинения mainSeq. Компиляция подтвердила существование команды/терминов, **не runtime-поддержку уровня данным типом chart**. Read-only `build by level` находится у `effect information`: не пытаться `set build by level`; level передаётся при `add effect`.

## Hover отдельного прозрачного PNG drone, без видео

Это эффект парения, не mouse-over trigger. PNG должен быть отдельным `p:pic` с alpha, а не частью фоновой картинки или MP4. Native `p:animMotion` работает с обычным `p:spTgt` без graphicEl. Малый вертикальный ход, autoreverse и indefinite repeat дают возврат без скачка.

Пример behavior (вставляется в childTnLst самостоятельно запускаемого par, не целый slide timing):

```xml
<p:animMotion origin="layout" path="M 0 0 L 0 -0.008 E" pathEditMode="relative">
  <p:cBhvr additive="base">
    <p:cTn id="901" dur="1600" autoRev="1" repeatCount="indefinite"
           accel="50000" decel="50000" fill="hold">
      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
    </p:cTn>
    <p:tgtEl><p:spTgt spid="77"/></p:tgtEl>
    <p:attrNameLst>
      <p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName>
    </p:attrNameLst>
  </p:cBhvr>
</p:animMotion>
```

Здесь 0.008 — относительное малое смещение пути, а не EMU. Это кандидат, **не runtime-проверенный эффект**; амплитуду/единицы конкретной реализации проверить на слайде. autoRev возвращает к базовой позиции, repeatCount=indefinite зацикливает. Не использовать accumulate=always: иначе дрон будет уплывать. Если повторяется весь родитель, а не motion, убедиться, что входной Wipe/Fade дрона не перезапускается каждый цикл. Бесконечное движение лучше держать параллельной веткой timeline, не блокирующей `after previous` для chart и заголовков. У slide transition не задавать слишком ранний auto-advance.

Native AppleScript-словарь также содержит `add behavior`, `motion effect.path`, `timing.autoreverse` и repeat count (integer). Но значение «до конца слайда» как число repeatCount здесь не установлено; не подставлять выдуманную магическую константу. OOXML `repeatCount="indefinite"` выражает намерение явно. Малое вращение PNG возможно отдельным animRot с autoRev, но не обязательно; не путать hover с эффектами Pulse/GrowShrink, которые масштабируют аппарат.

## Что нужно проверить при будущей реализации

1. Создать отдельный тестовый deck с двумя положительными столбцами и single-series pie на три сектора; использовать AddEffect и сохранить копию.
2. Прочитать ZIP: проверить subtargets и bld metadata, что это не голый spTgt всего chart; проверить сохранённые preset/filter, уникальные id, условия старта и порядок.
3. Открыть копию, запустить именно native slideshow, записать несколько кадров: до старта, середина столбца, после первого/до второго, середина сектора. PDF/статичный PNG не проверяют анимацию.
4. Убедиться, что сетка видна, будущие элементы скрыты, каждый столбец раскрывается с правильной стороны, а pie не просто появляется целиком. Проверить hover минимум на двух полных возвратах, без рывка и без влияния на дальнейший timeline.
5. После native save повторно извлечь exact slide XML; PowerPoint вправе нормализовать или отбросить вручную внесённые узлы.

## Источники

- Установленный PowerPoint.sdef (команды и буквальные имена перечислений); локально прочитан XML, не пересказ VBA Windows API.
- ECMA-376 5th, Transitional `pml.xsd` и `dml-main.xsd`, фактически скачаны и использованы для XSD-проверок:
  https://github.com/QtExcel/ecma-376-5th/tree/master/ECMA-376/OfficeOpenXML-XMLSchema-Transitional
  Типы `CT_TLShapeTargetElement`, `CT_TLGraphicalObjectBuild`, `CT_AnimationChartElement`, `ST_ChartBuildStep`, `CT_AnimationChartBuildProperties`.
- LibreOffice OOXML import preset mapping: `commontimenodecontext.cxx` фактически скачан/прочитан; ENTRANCE 21=wheel, 22=wipe:
  https://github.com/LibreOffice/core/blob/master/oox/source/ppt/commontimenodecontext.cxx
- LibreOffice effect library, фактически прочитана: oneBlade Wheel; Wipe from-bottom как reverse topToBottom:
  https://github.com/LibreOffice/core/blob/master/sd/xml/effects.xml
- Справочная страница chart:
  https://ooxml.info/docs/20/20.1/20.1.2/20.1.2.2/20.1.2.2.3/
- Справочная bldChart:
  https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_bldChart_topic_ID0ESZSIB.html

Web extraction некоторых справочных страниц вернул 403; вместо доверия сниппетам взяты настоящие XSD и открытые исходники. Документация/XSD устанавливают синтаксис и модель таргетинга; особенности фильтра, видимость до входа и геометрию маски на Mac PowerPoint подтверждает только ещё не выполненный native runtime-тест.
