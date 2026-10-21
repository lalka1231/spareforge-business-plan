# AgroVector: независимый технический review анимации

## Итог и границы проверки

Прочитаны `scripts/build_motion.py` и `scripts/build_presentation_design.py`. Проект не изменялся. Проверены установленная библиотека и исходники генерации movie/timing python-pptx; прочитан ZIP существующего MOTION. **Нативный показ в PowerPoint не запускался; воспроизведение, зацикливание и визуальная плавность не подтверждены.** XML/ZIP-readback — только структурная проверка.

Рекомендуемый надежный путь: короткий H.264 MP4 без аудио, встроенный через `add_movie`, один video timing с autoplay/loop на сцену; переходы Push вниз между сценами; native motion-path у отдельного сгруппированного дрона. Morph только для намеренно согласованных объектов, не как способ непрерывно проигрывать видео между слайдами. На финальной машине обязательно выполнить native slideshow acceptance ниже. Максимально надежная сериализация сложных эффектов — эталонный слайд, сохраненный установленным PowerPoint после настройки через UI, с последующим тиражированием и remap shape/timing IDs.

## Фактическое окружение

- `/Applications/Microsoft PowerPoint.app`: установлен, CFBundleShortVersionString **16.114**.
- `/Users/daniilvazdaev/Desktop/уник/Бизнес-планирование/AgroVector/.venv/bin/python`: существует; `python-pptx 1.0.2`, lxml импортируется.
- `ffmpeg`: не найден в PATH, `/opt/homebrew/bin/ffmpeg`, `/usr/local/bin/ffmpeg`; Spotlight `mdfind` совпадений не вернул. Это не доказательство отсутствия бинарника на всех дисках, но доступной команды сейчас нет. Устанавливать без согласования не стал.
- `presentations/AGROVECTOR_22_SLIDES_MOTION.pptx`: существует, ZIP testzip OK, embedded `.mp4`/`.mov` нет.
- `/tmp/agrovector-cinematic-assets.md` во время чтения отсутствовал: лицензии, пригодность и кодеки внешних материалов не исследованы.
- `sdef` CLI не работает с активным CommandLineTools вместо полного Xcode. Альтернатива проверена: словарь доступен непосредственно в `/Applications/Microsoft PowerPoint.app/Contents/Resources/PowerPoint.sdef`; есть команды `run slide show`, `go to slide`, `exit slide show`. Точный AppleScript синтаксис требует просмотра параметров словаря; не следует выдумывать COM/Windows API для macOS.

## Что делает имеющийся код и где опасно

1. `build_motion.py:19–29`: только fade входа групп по кликам. `delay="indefinite"` внешней группы и `nodeType="clickEffect"` — не автоматическая hover-анимация. `animEffect filter="fade"` не движение по траектории.
2. `build_motion.py:51–58`: **удаляет весь `p:timing`**. Если ранее добавить MP4 через python-pptx, этот участок сотрет его media timing: видео может остаться на слайде постером/по клику. Требуется единое построение timing после добавления всех movie и фигур; либо сохранение media-веток с перенумерацией ID. Это главный blocker.
3. `build_presentation_design.py:44–47`: дрон состоит из отдельных линий, роторов и корпуса, не group. Анимация только корпуса оставит роторы на месте. Сначала сгруппировать элементы в один `p:grpSp` (с корректным group transform) или создать отдельный прозрачный PNG/SVG-дрон и анимировать его как единую фигуру. Не двигать контент/таблицы с ним.
4. Существующий Morph связывает прежде всего progress rail/fill и `!!AV_editorial_edge`; это декоративное изменение, не кинематографический flyover.
5. При добавлении видео **после** имеющегося timing python-pptx тоже может перестроить timing: docstring `CT_Slide.get_or_add_childTnLst` прямо предупреждает о возможном удалении сложной структуры. Порядок: shapes и movies → единый финальный timing/transition patch → save ZIP, после этого не вызывать add_movie на этих слайдах.

## MP4: упаковка, autoplay, loop

### Подготовка (команды после установки ffmpeg)

```bash
ffmpeg -i source.mov -an -vf 'scale=1920:1080:force_original_aspect_ratio=increase,crop=1920:1080,fps=30' -c:v libx264 -pix_fmt yuv420p -crf 20 -preset medium -movflags +faststart scene.mp4
ffmpeg -i scene.mp4 -frames:v 1 poster.jpg
ffprobe -v error -show_entries stream=codec_name,codec_type,pix_fmt,width,height,r_frame_rate -show_entries format=duration,size -of json scene.mp4
```

Это команды для будущего выполнения, не выполненный encoding. 1080p — безопасный старт для macOS, 4K только после тестов. Loop должен быть смонтирован бесшовным: флаг repeat не устраняет скачок между последним и первым кадром. Одинаковые границы crop для постера и видео. Статичный постер нужен на время декодирования/перехода. Видео mute/без audio-track; не включать Play Full Screen для фона, иначе закроются надписи.

```python
movie = slide.shapes.add_movie(
    'scene.mp4', 0, 0, prs.slide_width, prs.slide_height,
    poster_frame_image='poster.jpg', mime_type='video/mp4')
# Shape tree z-order: перенести именно movie._element в начало области shapes,
# после nvGrpSpPr и grpSpPr, но не перед ними.
spTree = slide.shapes._spTree
spTree.remove(movie._element)
spTree.insert(2, movie._element)
```

Библиотека формирует `a:videoFile r:link="rIdVideo"` и `p14:media r:embed="rIdMedia"`, плюс poster `a:blip`. `r:link` сам по себе НЕ означает внешний URL: нужно читать `.rels`; video и media relationship должны разрешаться в `ppt/media/*.mp4` без `TargetMode="External"`. Не заменять movie обычной картинкой с гиперссылкой. Не прописывать только p14:media без соответствующих rels и content types.

### Минимальная media-ветка timing

Под `p:timing/p:tnLst/p:par/p:cTn[@nodeType='tmRoot']/p:childTnLst` добавить ветку ниже. Она параллельна mainSeq для кликовых/одноразовых эффектов, а НЕ помещена внутрь последовательности, которая ждет конца бесконечного видео. `SPID` — фактический `movie.shape_id`; все cTn IDs уникальны в пределах слайда.

```xml
<p:video xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cMediaNode vol="0">
    <p:cTn id="30" fill="hold" display="0" repeatCount="indefinite">
      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
    </p:cTn>
    <p:tgtEl><p:spTgt spid="SPID"/></p:tgtEl>
  </p:cMediaNode>
</p:video>
```

Это точечное изменение default python-pptx: у него `vol="80000"`, `delay="indefinite"`, отсутствует repeatCount. Loop относится к `p:cTn`, **не** к `p:video`, `p:cMediaNode` или `p14:media`. Не задавать видео `dur="indefinite"` вслепую: естественная длительность MP4 должна завершать один проход; повторяется именно проход. Корневой tmRoot имеет `dur="indefinite" restart="never"`. `fill="hold"` не включает loop. Удалить дублирующий click-to-play effect/interactiveSeq, если он остался из эталона; hyperlink media на picture сама по себе не заменяет autoplay.

Важно: предложенная OOXML-конфигурация технически обоснована исходниками библиотеки и моделью timing; **ее runtime на установленном PowerPoint еще надо подтвердить**. Если Mac нормализует ее иначе, приоритет имеет XML после native save эталона с Playback → Start Automatically → Loop until Stopped, а не предположение о разметке.

## Push вниз и Morph

Стандартный Push, движение содержимого вниз:

```xml
<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med" advClick="1">
  <p:push dir="d"/>
</p:transition>
```

`d` — направление движения; если под «вниз» имеется в виду прокрутка камеры к следующему экрану, визуально потребуется `dir="u"`. Зафиксировать смысл при native просмотре. Отсутствие `advTm` означает отсутствие авто-перехода: видео играет автоматически, слайды — по клику. Если нужен полностью автономный показ, задавать `advTm` отдельно; это не настройка loop видео.

Morph с fallback Push можно сериализовать на уровне слайда через MC (не добавлять рядом еще один обычный transition):

```xml
<mc:AlternateContent
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main">
 <mc:Choice Requires="p159">
  <p:transition spd="med" advClick="1"><p159:morph option="byObject"/></p:transition>
 </mc:Choice>
 <mc:Fallback>
  <p:transition spd="med" advClick="1"><p:push dir="d"/></p:transition>
 </mc:Fallback>
</mc:AlternateContent>
```

Ставить после `p:clrMapOvr` (если есть), до `p:timing` и `p:extLst`. Исходный код использует MC внутри transition; это не следует называть доказанной ошибкой само по себе: MC допускает подстановку содержимого. Эталон из native PowerPoint предпочтительнее и для расположения MC, и для extension duration. Namespace Morph в исходнике правильный. `!!` в именах — механизм forced object matching; одинаковое имя должно быть уникально на каждом слайде и относиться к одному логическому объекту.

Morph не гарантирует непрерывное playback видео через границу слайдов: каждая сцена запускает свой media instance с начала. Для seamless единого полета по всей презентации нужен заранее смонтированный общий MP4, а не обещание непрерывности Morph. Fallback Push и запуск ролика могут дать краткий постер/скачок — это отдельный acceptance test.

## Native hover и flyover

Здесь hover = физическое зависание дрона. **Не mouse-over trigger**: наведение мыши не является надежным штатным триггером произвольного motion path в обычной PPTX. Для движений нужны `p:animMotion` и target отдельного дрона; анимация содержится в PPTX, а не в GIF/MP4.

Схематический motion behavior с точными атрибутами OOXML (вставляется под effect/par в корневой параллельной ветке; id уникален):

```xml
<p:animMotion xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 origin="layout" pathEditMode="relative" path="M 0 0 L 0 -0.006 E">
 <p:cBhvr additive="base">
  <p:cTn id="41" dur="1400" autoRev="1" repeatCount="indefinite" fill="hold">
   <p:stCondLst><p:cond delay="0"/></p:stCondLst>
  </p:cTn>
  <p:tgtEl><p:spTgt spid="DRONE_SPID"/></p:tgtEl>
  <p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst>
 </p:cBhvr>
</p:animMotion>
```

Координаты path — нормализованные доли слайда, не EMU/пиксели. Отрицательный Y — вверх. Для hover малое смещение, reverse и повтор; при необходимости easing `accel/decel` на cTn после проверки native поведения. Нельзя делать loop родительского par с длительностью indefinite и ожидать, что конечный дочерний путь начнет повторяться сам: repeat должен стоять у behavior либо конечной обертки.

Flyover: отдельная сцена/другая копия дрона, тот же тип behavior, например `path="M 0 0 L 0.75 -0.12 E"`, `dur="6000"`, без autoRev/repeat, `fill="hold"`. Исходную позицию дрона задать у левого края. Сложный изогнутый путь лучше записать в UI и извлечь, чем угадывать синтаксис Bézier. Не запускать одновременно два независимых motion behaviors с `additive="base"` на одном shape (hover + flyover): они могут перетирать координаты. Использовать последовательные эффекты или внешний group для полета и вложенный shape для bobbing; последний требует отдельного теста вложенных transform на Mac.

Для действительно проверяемого native результата: создать двухслайдовый контрольный файл в `/tmp`, на первом hovering, на втором flyover, применить Motion Paths в PowerPoint (With Previous, repeat Until End of Slide для hover), сохранить. Прочитать native XML, использовать его как golden fixture. Не подменять native motion видеозаписью дрона.

## Обязательная проверка результата

### Структура — автоматизировать до запуска PowerPoint

- ZipFile.testzip, все media/video/image relationships разрешаются внутри ZIP; mp4 content-type `video/mp4`; размер > 0, ffprobe извлеченного embedded файла, не только исходного.
- На каждом предполагаемом видео-слайде movie shape присутствует, video target spid существует, один autoplay media cTn с delay 0 и repeatCount indefinite; уникальные IDs timing. Проверять соответствие scene → media, а не только общее количество MP4 (дедупликация допустима).
- Нет неожиданного TargetMode External у media; анимационный target не body-only и не отсутствующая shape.
- Ровно один эффективный transition после MC-processing, transition перед timing; после ZIP patch python-pptx readback не считается XSD validation. Полезен Open XML SDK Validator с поддержкой выбранной Office версии; наличие lxml и XML parse не равно проверке по схеме.
- Сравнить тексты, данные диаграмм/таблиц, notes; исходные проверки текста/notes в build_motion не покрывают все данные и animations runtime.

### Настоящий acceptance в установленном Microsoft PowerPoint

Открыть именно конечный файл: `open -a 'Microsoft PowerPoint' '/absolute/path/final.pptx'`. Открытие/запуск команды само по себе ничего не доказывает. Не сохранять поверх источника. При repair alert — fail, сохранить сообщение/repair log.

1. Slide Show, а не Normal/Preview/PDF. Войти на сцену и не кликать: MP4 начинает реально меняться; дрон движется независимо от видео. Animation Pane содержит Motion Path, а не только Fade.
2. Оставаться минимум на **двух полных длительностях** MP4: виден второй цикл, не замороженный последний кадр. Записать видео экрана с таймингом, включая границу циклов. Для каждого слайда фиксировать начало/loop; проверка одного не доказывает работу остальных.
3. Клик перехода с бесконечным media/hover все же переводит на следующий слайд, не требует сначала десятков кликов и не зависает на loop. Повторить Back → Forward и прямой jump на сцену.
4. Push действительно движется вниз; Morph совпадает по именованным объектам, не превращает весь фон в случайное crossfade. Смотрите стык видео, poster и перехода на реальной частоте экрана.
5. Hover: роторы и корпус движутся вместе и возвращаются без дрейфа; минимум три периода. Flyover: один непрерывный проход, сохранение конечной позиции/исчезновение соответствует сценарию; нет возвращения при следующем клике.
6. Закрыть, заново открыть сохраненную копию, снова показать офлайн. Проверить после копирования PPTX в другую папку без исходных MP4. Это подтверждает embedded/offline именно runtime, а не только пути в ZIP.
7. Зафиксировать build/version Mac PowerPoint, имя/хеш проверенного PPTX, слайды, реальные записи и исход pass/fail. Не обозначать экспорт PDF, thumbnails, LibreOffice render или XML assertions как playback verification.

## Источники и ограничения исследования

- Исходный код установленной `python-pptx 1.0.2`: `CT_Picture._pic_video_tmpl`, `CT_TimeNodeList.add_video`, `SlideShapes._add_video_timing`, `CT_Slide.get_or_add_childTnLst` — прочитан через inspect в проектном venv.
- Microsoft autoplay UI: https://support.microsoft.com/en-us/powerpoint/play-a-video-automatically-in-a-slide-show (подтвержден поисковой выдачей).
- Morph namespace/element, прочитано напрямую urllib после отказа extract: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/68d26d78-f7f5-47ab-835d-4e6c82ff39f0
- Morph type: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/41ca8fbf-efc8-49ac-8a32-7bd0856544bd
- Media node reference: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commonmedianode?view=openxml-3.0.1 (найден, полный extract недоступен).

Web extractor возвратил 403 для трех справочных страниц; прямой urllib успешно получил Morph specification. Никакие synthetic media, придуманные результаты slideshow или ссылки на несуществующую запись не использовались. Примеры XML/команды в отчете — спецификация реализации и дальнейшего acceptance, не заявление об их уже выполненном нативном тесте.
