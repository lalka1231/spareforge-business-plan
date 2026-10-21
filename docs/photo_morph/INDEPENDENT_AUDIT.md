# Read-only приемка AGROVECTOR_PHOTO_MORPH.pptx

Файл: `/Users/daniilvazdaev/Desktop/уник/Бизнес-планирование/AgroVector/presentations/AGROVECTOR_PHOTO_MORPH.pptx`

SHA-256: `5aff50b46ef832598d5bfc0b38ffa2db26922b34fd89d0302c84dba82fed0f04`

## Вердикт
Структурные требования в основном соблюдены, но буквальное имя `!!AV_droneMorph` НЕ соблюдено: в файле `!!AV_drone`. Независимое подтверждение сохранения именно этой финальной версии через native PowerPoint отсутствует: Application=Microsoft Macintosh PowerPoint — лишь метаданные, которые могут наследоваться; повторное сохранение в read-only приемке не выполнялось. Runtime-пиксели не проверены.

## Видео и Morph
MP4: 0; video/videoFile XML: 0. MMClips=22 в app.xml — устаревшая статистика, не наличие видео.

Слайд 1: 2 motion paths.
- [('!!AV_background', {'origin': 'layout', 'path': 'M 0 0 L 0 -0.004 E', 'pathEditMode': 'relative'}, [{'id': '4', 'dur': '5000', 'autoRev': '1', 'repeatCount': 'indefinite', 'accel': '40000', 'decel': '40000', 'fill': 'hold'}])]
- [('!!AV_drone', {'origin': 'layout', 'path': 'M 0 0 L 0 -0.012 E', 'pathEditMode': 'relative'}, [{'id': '6', 'dur': '2200', 'autoRev': '1', 'repeatCount': 'indefinite', 'accel': '40000', 'decel': '40000', 'fill': 'hold'}])]

Слайд 2: 2 motion paths.
- [('!!AV_drone', {'origin': 'layout', 'path': 'M 0 0 L -0.12 0 E', 'pathEditMode': 'relative'}, [{'id': '4', 'dur': '4000', 'autoRev': '1', 'repeatCount': 'indefinite', 'accel': '40000', 'decel': '40000', 'fill': 'hold'}])]
- [('!!AV_spray', {'origin': 'layout', 'path': 'M 0 0 L -0.12 0 E', 'pathEditMode': 'relative'}, [{'id': '6', 'dur': '4000', 'autoRev': '1', 'repeatCount': 'indefinite', 'accel': '40000', 'decel': '40000', 'fill': 'hold'}])]

График 11: 5 отдельных точек; фильтры ['wipe(up)']; каждый animEffect содержит graphicEl/a:chart c ptInSeries и индексами series/category; whole-chart fade отсутствует. Значения: [[54.15384615384615, 81.23076923076923, 100.61538461538461, 110.3076923076923, 120.0]]

График 17: 15 отдельных точек; фильтры ['wipe(up)']; каждый animEffect содержит graphicEl/a:chart c ptInSeries и индексами series/category; whole-chart fade отсутствует. Значения: [[8.8, 13.728, 17.630080000000003, 20.078822400000004, 22.695256064000002], [1.96, 5.16, 7.619380000000004, 9.035079900000003, 10.540193564000003], [1.432, 4.03632, 5.961575200000003, 6.930350556000001, 10.578478200160003]]

График 18: 6 отдельных точек; фильтры ['wipe(left)', 'wipe(right)']; каждый animEffect содержит graphicEl/a:chart c ptInSeries и индексами series/category; whole-chart fade отсутствует. Значения: [[3.946443104442745, 0.04487293632982997, -2.6388362920894077, 1.4115262854426354, -1.4330412442990532, -6.605370129557806]]

График 19: 3 отдельных точек; фильтры ['wheel(1)']; каждый animEffect содержит graphicEl/a:chart c ptInSeries и индексами series/category; whole-chart fade отсутствует. Значения: [[7.0, 0.5, 2.0]]

График 20: 2 отдельных точек; фильтры ['wheel(1)']; каждый animEffect содержит graphicEl/a:chart c ptInSeries и индексами series/category; whole-chart fade отсутствует. Значения: [[2.5, 7.0]]

Всего Morph: 22/22 (включая первый); графики: [11, 17, 18, 19, 20].

## Данные
PASS: все 5 графиков сопоставлены программно с financial_model_5y.csv, sensitivity.csv, assumptions.csv, допуск 1e-8 в отображаемых единицах. Слайд 11: внесение/130 + мониторинг/500; 17: выручка/EBITDA/свободный CF в млн руб.; 18: NPV в млн руб.; 19: 7 + 0,5 + 2; 20: 2,5 + 7 млн руб. Это проектные суммы, не проценты рынка/капитала.

## Отдельный PNG-дрон
[(1, '!!AV_drone', 'ppt/media/image8.png', 'e0ef369529734f56d9d3e4bfe9cfaba7dec80e42501dfa955af984abb7b3a530', [{'x': '6583680', 'y': '1234440'}, {'cx': '5029200', 'cy': '3143250'}]), (2, '!!AV_drone', 'ppt/media/image8.png', 'e0ef369529734f56d9d3e4bfe9cfaba7dec80e42501dfa955af984abb7b3a530', [{'x': '8229600', 'y': '1920240'}, {'cx': '2788920', 'cy': '1743075'}])]
Одинаковый PNG и SHA на 1/2; положение и размер различаются. Общего одинакового имени !!AV_drone достаточно для сопоставления объектов Morph, но заданного буквального !!AV_droneMorph нет.

## Notes
PASS: все 22 исходных notes из AGROVECTOR_22_SLIDES.pptx целиком содержатся в соответствующих notes финала; добавлены пояснения визуальной редакции. Проверка по тексту, не побайтовому равенству XML.

## Фотографии
22 background JPEG, 22 различных SHA-256, все совпадают с реально встроенными фонами; 10 уникальных source_url в ledger. Корректное описание: 22 индивидуальные композиции из 10 источников, НЕ 22 разных съемки. Повторные источники кадрированы иначе (это подтверждается кодом build_photo_morph.py и различными хешами результатов, не является независимой визуальной оценкой). Первый фон — отдельный статичный кадр из видеoисточника Pexels, видео в PPTX не встроено.

| Слайд | SHA-256 background | Совпадает с PPTX |
|---|---|---|
| 1 | `f639f77208583eeb8b9f02fcbf2453f0915c7be2b00104c3cab623192923ba61` | True |
| 2 | `6d56dafe19ac54788ec3201add2d39115922313f173bcfc5aa174c26ebc4f1b7` | True |
| 3 | `3c1ecf2839fc2bac4f5947abcb321082c0f16845dd3040bda8532d359e45c445` | True |
| 4 | `10d864b893cd8a77efd73d0191c5f86b0b675ec7ea4b8453b3b79f48e3e31332` | True |
| 5 | `46f5d4e87f637d41b083832fee0c8d3a986a637877e9488598d738f169a921a4` | True |
| 6 | `7e0e3e40d26209449756ff3e360619d30c7481bad4727139bb7c1791be718e47` | True |
| 7 | `4570f120a5d2c2647aad1236d0698459155f17d82e43c8074c79f230a4409acf` | True |
| 8 | `2f499971a36cd8b9b6e87c7cd4eef4e698ffab2b8c50d0d0af32f437dc91eb74` | True |
| 9 | `88838a91154e8997d1d825b4816c7d0b843bf19e5cbeaf7d6d847f09cdeb1b02` | True |
| 10 | `2a67e19cdd11ffa8274f644bd3bdf0c413ca239f2c07dfba62d73262c8911d37` | True |
| 11 | `d4bb12e758f8ecc8468ba8b1ba6bd0edbf02c2fbcc3f6135ed14d4620d28eb55` | True |
| 12 | `e431c6e5bc39b3e5c768a42c41c1fc341487171c645dcbd553a36622b9da494b` | True |
| 13 | `6dc1dbc28f132d9acc159878c921b9f8309bfa9f3fd20b2de009b40202400d49` | True |
| 14 | `3565caed5f5616130a74a7d36ff7cfe54ccec157f1ebe9151659008ad1b2993e` | True |
| 15 | `0c5a69e59bcb8bd71b8c81e2b18f583bdf1ed445f52b663aa2e1a9290e963f8d` | True |
| 16 | `43117ee74835ad1ab122b8cee2c3411c520231821e93ba6d54df4c2e40823b1a` | True |
| 17 | `fcdfa2772a800396d8f27c8d6e0a2d142b838a57cc426bc75b0f4512b88c716f` | True |
| 18 | `77aaaa5b17605d9781413e2d1899c0b9c23036339bffd78a19b055542afa101e` | True |
| 19 | `e4b8bf16225522d47fd4e58edb2f836c73c7609ebd71818041f229f77c4343d7` | True |
| 20 | `4f979cd1a852bdb14429f3b4700ff5eec98995fc1fcd619110c662d86d9a6dd3` | True |
| 21 | `1455dede9596db7614df52aedb56c82c377ffe96da028a943c38a12ec26dfce4` | True |
| 22 | `4f6def2bbc1a9baf21ceeba84cbeaf341a0a56a0b5be7474ad39a5f7129c6c7a` | True |

## Границы приемки
PPTX не изменен, SHA до/после совпадает. Не запускались native-save, slideshow, видеоэкспорт или PDF-анимации. Наличие XML-анимаций доказывает настройки, но не фактическое движение пикселей/визуальную плавность. Сеть не использовалась.
