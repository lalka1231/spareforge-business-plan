# FINALnative video autoplay — адресная read-only приемка

**Итог: PASS структурной приемки. Full pixel playback НЕ проверен.**

- Файл: `/Users/daniilvazdaev/Desktop/уник/Бизнес-планирование/AgroVector/presentations/AGROVECTOR_CINEMATIC.pptx`.
- Дата проверки: 2026-10-20T23:43:50.399460+03:00.
- SHA-256 до и после чтения: `d6c6e0bb691d3419942574af272b28cc90d2ec53a0077640273bd6baef609ce0`.
- Неизменность SHA: True. PPTX и файлы проекта не изменялись; PowerPoint не запускался и пересохранение не выполнялось.
- В исходном сообщении путь содержал `AgroVector,presentations`; фактический файл найден по подтвержденному историческим отчетом пути `AgroVector/presentations`.
- История `/tmp/agrovector-cinematic-audit.md` прочитана. Ее прежние ошибки не перенесены на текущий артефакт.

## Проверки текущего файла

- Слайдов в presentation/sldIdLst: 22; проверены все в порядке презентации.
- У каждого слайда ровно один native mainSeq; все имеющиеся start conditions внутри mainSeq имеют delay=0, включая внешние обертки и mediacall.
- В mainSeq каждого слайда mediacall имеет nodeType=withEffect и команду call playFrom(0.0), направленную на существующий media shape.
- clickEffect внутри mainSeq: 0. Отдельная interactiveSeq с clickEffect/togglePause не является click-зависимостью mainSeq и не трактуется как дефект autoplay.
- Все media cTn: repeatCount=indefinite.
- Shape IDs и timing cTn IDs уникальны внутри каждого слайда; animation targets существуют. Совпадение IDs между разными слайдами допустимо.
- Встроенных MP4: 5. Push вверх: 22/22.
- ZIP CRC: OK.
- Нормализованный media delay=indefinite не считается ошибкой: проверен путь автоматической нативной команды, а не устаревший буквальный критерий media delay=0.

## Послайдовый протокол

|Слайд|mainSeq delays|withEffect + playFrom|clickEffect в mainSeq|loop indefinite|unique shape/timing IDs|targets|push ↑|Итог|
|---:|---|---|---:|---|---|---|---|---|
|1|0,0,0|True|0|True|True/True|True|True|PASS|
|2|0,0,0|True|0|True|True/True|True|True|PASS|
|3|0,0,0|True|0|True|True/True|True|True|PASS|
|4|0,0,0|True|0|True|True/True|True|True|PASS|
|5|0,0,0|True|0|True|True/True|True|True|PASS|
|6|0,0,0|True|0|True|True/True|True|True|PASS|
|7|0,0,0|True|0|True|True/True|True|True|PASS|
|8|0,0,0|True|0|True|True/True|True|True|PASS|
|9|0,0,0|True|0|True|True/True|True|True|PASS|
|10|0,0,0|True|0|True|True/True|True|True|PASS|
|11|0,0,0|True|0|True|True/True|True|True|PASS|
|12|0,0,0|True|0|True|True/True|True|True|PASS|
|13|0,0,0|True|0|True|True/True|True|True|PASS|
|14|0,0,0|True|0|True|True/True|True|True|PASS|
|15|0,0,0|True|0|True|True/True|True|True|PASS|
|16|0,0,0|True|0|True|True/True|True|True|PASS|
|17|0,0,0|True|0|True|True/True|True|True|PASS|
|18|0,0,0|True|0|True|True/True|True|True|PASS|
|19|0,0,0|True|0|True|True/True|True|True|PASS|
|20|0,0,0|True|0|True|True/True|True|True|PASS|
|21|0,0,0|True|0|True|True/True|True|True|PASS|
|22|0,0,0|True|0|True|True/True|True|True|PASS|

## Встроенные MP4

- `ppt/media/media2.mp4`
- `ppt/media/media3.mp4`
- `ppt/media/media4.mp4`
- `ppt/media/media5.mp4`
- `ppt/media/media1.mp4`

## Границы приемки

Сообщенное пользователем подтверждение native PowerPoint play on entry / loop=true после открытия и сохранения принято как контекст, но этим read-only аудитом повторно не измерялось. По движущимся пикселям не проверены автозапуск всех роликов, непрерывность движения, вращение роторов, качество петли и звук. Структурный PASS не является full playback / full visual PASS. Содержание, canonical-данные и статическая верстка не входят в эту адресную повторную проверку.

Создан только этот отчет и вспомогательный проверочный скрипт `/tmp/agrovector-autoplay-readonly-check.py`; проект не изменен.

## Финальная косметическая правка координатора

После адресного аудита изменены только положение/высота таблицы6 и вынесена легенда диаграммы17: наложения обнаружены увеличенным просмотром и устранены. Тексты ячеек, значения диаграмм и видеотайминг сохранены. Все структурные и медиапроверки повторно пройдены, native API подтвердил22/22 autoplay+loop; актуальный SHA и результаты — presentations/CINEMATIC_ACCEPTANCE.json. Последний SHA: d4ff42e7c90d7b13b624d514391f7dba960ad8f7f6051de502d5db2868d14fd7.
