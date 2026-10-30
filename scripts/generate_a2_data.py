#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate complete curriculum data for Level A2 (TRKI-A2 / Базовый уровень)
Skills: Reading (5 lessons), Listening (3 lessons), Speaking (3 lessons), Writing (3 lessons)
"""
import json
import os

BASE_DIR = "public/data"

# ==============================================================================
# 1. A2 READING LESSONS
# ==============================================================================
a2_reading_lessons = [
    {
        "id": 1,
        "level": "A2",
        "skill": "reading",
        "title": "Московское метро: ориентирование и подземные дворцы",
        "badge": "A2 Đọc hiểu",
        "readingTime": "15 phút",
        "source": "Giáo trình TRKI-A2 Chuẩn Quốc Tế",
        "topic": "Giao thông công cộng & Chỉ đường trong đô thị",
        "vocabCategories": [
            {
                "title": "Từ vựng Giao thông & Chỉ đường (Транспорт и ориентирование)",
                "rows": [
                    {"word": "ста́нция метро́", "gram": "danh từ giống cái", "meaning": "ga tàu điện ngầm", "example": "Эта станция метро очень красивая."},
                    {"word": "переса́дка", "gram": "danh từ giống cái", "meaning": "điểm chuyển tuyến", "example": "Здесь удобная пересадка на другую линию."},
                    {"word": "эскала́тор", "gram": "danh từ giống đực", "meaning": "thang cuốn", "example": "Вставайте справа на эскалаторе."},
                    {"word": "вы́ход в го́род", "gram": "cụm danh từ", "meaning": "lối ra thành phố", "example": "Где находится выход в город?"},
                    {"word": "поверну́ть", "gram": "động từ hoàn thành", "meaning": "rẽ, quẹo", "example": "Поверните направо возле аптеки."}
                ]
            },
            {
                "title": "Động từ chuyển động & Giới từ (Глаголы движения)",
                "rows": [
                    {"word": "дое́хать до", "gram": "động từ + Cách 2", "meaning": "đi xe đến tận", "example": "Как доехать до центра города?"},
                    {"word": "пересе́сть на", "gram": "động từ + Cách 4", "meaning": "đổi sang tuyến/xe khác", "example": "Вам нужно пересесть на Кольцевую линию."},
                    {"word": "пройти́ ми́мо", "gram": "động từ + Cách 2", "meaning": "đi bộ ngang qua", "example": "Пройдите мимо большого парка."}
                ]
            }
        ],
        "readingText": [
            "Московский метрополитен — это не просто удобный общественный транспорт, но и настоящий подземный музей. Он открылся в 1935 году и сегодня является одной из самых загруженных и красивых систем метро в мире.",
            "Каждая станция имеет свой уникальный архитектурный стиль. Например, станция «Маяковская» украшена мозаиками, а на станции «Площадь Революции» можно увидеть 76 бронзовых скульптур. Студенты перед экзаменами часто гладят нос бронзовой собаки на удачу.",
            "Ориентироваться в московском метро очень просто. Каждая линия имеет свой номер, название и цвет на карте. Самая известная — Кольцевая линия, она коричневого цвета. Если вам нужно сделать пересадку, следуйте за указателями на полу и стенах.",
            "Чтобы оплатить проезд, пассажиры используют пластиковую карту «Тройка» или прикладывают банковскую карту к турникету. Поезда ходят очень часто: в часы пик интервал между ними составляет всего 90 секунд."
        ],
        "questions": [
            {
                "id": "a2_r1_q1",
                "lang": "ru",
                "prompt": "В каком году было открыто московское метро?",
                "hint": "Tìm thông tin về năm khai trương ở đoạn 1.",
                "options": [
                    {"key": "A", "text": "В 1917 году"},
                    {"key": "B", "text": "В 1935 году"},
                    {"key": "C", "text": "В 1945 году"},
                    {"key": "D", "text": "В 2000 году"}
                ],
                "ans": "B",
                "explain": "Đoạn 1 nêu rõ: «Он открылся в 1935 году...»."
            },
            {
                "id": "a2_r1_q2",
                "lang": "ru",
                "prompt": "Что делают студенты на станции «Площадь Революции» перед экзаменами?",
                "hint": "Xem hành động đem lại may mắn cho sinh viên ở đoạn 2.",
                "options": [
                    {"key": "A", "text": "Читают конспекты на эскалаторе"},
                    {"key": "B", "text": "Покупают счастливый билет"},
                    {"key": "C", "text": "Гладят нос бронзовой собаки на удачу"},
                    {"key": "D", "text": "Фотографируются с преподавателем"}
                ],
                "ans": "C",
                "explain": "Đoạn 2 viết: «Студенты перед экзаменами часто гладят нос бронзовой собаки на удачу»."
            },
            {
                "id": "a2_r1_q3",
                "lang": "ru",
                "prompt": "Какого цвета Кольцевая линия на карте метро?",
                "hint": "Tìm màu sắc của tuyến đường tròn (Кольцевая) ở đoạn 3.",
                "options": [
                    {"key": "A", "text": "Красного"},
                    {"key": "B", "text": "Синего"},
                    {"key": "C", "text": "Коричневого"},
                    {"key": "D", "text": "Зелёного"}
                ],
                "ans": "C",
                "explain": "Đoạn 3 khẳng định: «Самая известная — Кольцевая линия, она коричневого цвета»."
            },
            {
                "id": "a2_r1_q4",
                "lang": "ru",
                "prompt": "Как часто ходят поезда в часы пик?",
                "hint": "Tìm khoảng thời gian (интервал) giữa các chuyến tàu ở cuối bài.",
                "options": [
                    {"key": "A", "text": "Каждые 5 минут"},
                    {"key": "B", "text": "Каждые 90 секунд"},
                    {"key": "C", "text": "Каждые 10 минут"},
                    {"key": "D", "text": "Раз в полчаса"}
                ],
                "ans": "B",
                "explain": "Bài viết nêu: «в часы пик интервал между ними составляет всего 90 секунд»."
            }
        ]
    },
    {
        "id": 2,
        "level": "A2",
        "skill": "reading",
        "title": "Покупки: Супермаркет или традиционный рынок?",
        "badge": "A2 Đọc hiểu",
        "readingTime": "15 phút",
        "source": "Đời sống thường nhật nước Nga",
        "topic": "Mua sắm, Thực phẩm & Chi tiêu hàng ngày",
        "vocabCategories": [
            {
                "title": "Mua sắm & Thực phẩm (Продукты и покупки)",
                "rows": [
                    {"word": "суперма́ркет", "gram": "danh từ giống đực", "meaning": "siêu thị", "example": "Рядом с домом открылся новый супермаркет."},
                    {"word": "ры́нок", "gram": "danh từ giống đực", "meaning": "chợ dân sinh", "example": "На рынке всегда свежие овощи и фрукты."},
                    {"word": "ски́дка", "gram": "danh từ giống cái", "meaning": "giảm giá, khuyến mãi", "example": "Сегодня в магазине большая скидка на чай."},
                    {"word": "нали́чные", "gram": "danh từ số nhiều", "meaning": "tiền mặt", "example": "Вы платите картой или наличными?"},
                    {"word": "сда́ча", "gram": "danh từ giống cái", "meaning": "tiền thối lại", "example": "Возьмите чек и вашу сдачу."}
                ]
            },
            {
                "title": "Đơn vị đo lường & Chọn lựa (Меры и выбор)",
                "rows": [
                    {"word": "килогра́мм", "gram": "danh từ giống đực", "meaning": "ki-lô-gam (kg)", "example": "Взвесьте мне, пожалуйста, два килограмма яблок."},
                    {"word": "упако́вка", "gram": "danh từ giống cái", "meaning": "hộp, gói đóng sẵn", "example": "Упаковка творога стоит сто рублей."},
                    {"word": "све́жий", "gram": "tính từ", "meaning": "tươi, mới", "example": "У вас хлеб свежий?"}
                ]
            }
        ],
        "readingText": [
            "В современных российских городах люди покупают продукты в двух главных местах: в сетевых супермаркетах и на продуктовых рынках. У каждого места есть свои плюсы и особенности.",
            "В супермаркетах всё быстро и удобно. Цены указаны на ценниках, а постоянным покупателям предлагают карты лояльности со скидками. На кассе можно заплатить смартфоном или банковской картой. Многие супермаркеты работают круглосуточно.",
            "На рынок люди ходят за особой атмосферой và thực phẩm tươi sống. Здесь продают домашний творог, свежее мясо, солёные огурцы и сезонные ягоды. На рынке можно попробовать товар перед покупкой и даже немного поторговаться с продавцом.",
            "Большинство молодых россиян выбирают супермаркеты и онлайн-доставку, а пожилые люди и гурманы предпочитают ходить на рынок в субботу утром."
        ],
        "questions": [
            {
                "id": "a2_r2_q1",
                "lang": "ru",
                "prompt": "Чем удобны супермаркеты для покупателей?",
                "hint": "Tìm ưu điểm của siêu thị ở đoạn 2.",
                "options": [
                    {"key": "A", "text": "Там можно торговаться с продавцом"},
                    {"key": "B", "text": "Там всё быстро, указаны цены и есть скидки"},
                    {"key": "C", "text": "Там всегда продают только домашние продукты"},
                    {"key": "D", "text": "Там нет очередей на кассе"}
                ],
                "ans": "B",
                "explain": "Đoạn 2 nêu rõ sự tiện lợi: nhanh, có bảng giá rõ ràng và thẻ giảm giá."
            },
            {
                "id": "a2_r2_q2",
                "lang": "ru",
                "prompt": "Что разрешается делать покупателям на рынке перед покупкой?",
                "hint": "Đọc câu thứ 3 của đoạn 3 về trải nghiệm ở chợ.",
                "options": [
                    {"key": "A", "text": "Забрать товар бесплатно"},
                    {"key": "B", "text": "Попробовать товар на вкус"},
                    {"key": "C", "text": "Самостоятельно приготовить блюдо"},
                    {"key": "D", "text": "Позвонить директору рынка"}
                ],
                "ans": "B",
                "explain": "Đoạn 3 viết: «На рынке можно попробовать товар перед покупкой...»."
            },
            {
                "id": "a2_r2_q3",
                "lang": "ru",
                "prompt": "Кто чаще всего заказывает продукты онлайн и покупает в супермаркетах?",
                "hint": "Xem đoạn kết về thói quen mua sắm của các nhóm tuổi.",
                "options": [
                    {"key": "A", "text": "Пожилые пенсионеры"},
                    {"key": "B", "text": "Молодые россияне"},
                    {"key": "C", "text": "Иностранные туристы"},
                    {"key": "D", "text": "Продавцы рынка"}
                ],
                "ans": "B",
                "explain": "Đoạn cuối khẳng định: «Большинство молодых россиян выбирают супермаркеты и онлайн-доставку»."
            }
        ]
    },
    {
        "id": 3,
        "level": "A2",
        "skill": "reading",
        "title": "Русская зима и четыре времени года в России",
        "badge": "A2 Đọc hiểu",
        "readingTime": "15 phút",
        "source": "Khám phá Văn hóa & Địa lý Nga",
        "topic": "Thời tiết, Khí hậu & Bốn mùa tại Nga",
        "vocabCategories": [
            {
                "title": "Thời tiết & Thiên nhiên (Погода и природа)",
                "rows": [
                    {"word": "моро́з", "gram": "danh từ giống đực", "meaning": "giá rét, băng giá", "example": "Зимой в Сибири сильный мороз."},
                    {"word": "снегопа́д", "gram": "danh từ giống đực", "meaning": "tuyết rơi dày", "example": "После снегопада город стал белым."},
                    {"word": "гра́дус", "gram": "danh từ giống đực", "meaning": "độ (C)", "example": "Сегодня двадцать градусов мороза (-20°C)."},
                    {"word": "прогно́з пого́ды", "gram": "cụm danh từ", "meaning": "dự báo thời tiết", "example": "По прогнозу погоды завтра будет солнце."},
                    {"word": "та́ять", "gram": "động từ chưa hoàn thành", "meaning": "tan chảy (tuyết, đá)", "example": "Весной снег начинает таять."}
                ]
            },
            {
                "title": "Trang phục mùa đông (Зимняя одежда)",
                "rows": [
                    {"word": "тёплая ку́ртка", "gram": "cụm danh từ", "meaning": "áo khoác ấm", "example": "Обязательно наденьте тёплую куртку."},
                    {"word": "перча́тки", "gram": "danh từ số nhiều", "meaning": "găng tay", "example": "Не забудьте шерстяные перчатки."},
                    {"word": "отопле́ние", "gram": "danh từ giống trung", "meaning": "hệ thống sưởi ấm", "example": "В домах зимой работает центральное отопление."}
                ]
            }
        ],
        "readingText": [
            "Россия — огромная страна, и климат здесь очень разнообразный. Однако для большинства иностранцев Россия ассоциируется прежде всего с настоящей снежной зимой.",
            "Зима в центральной России длится с конца ноября по март. Температура часто опускается ниже нуля, иногда до минус двадцати градусов. В домах, школах и офисах очень тепло, потому что отлично работает центральное отопление.",
            "Русские люди любят зиму за активный отдых. В парках заливают катки, и тысячи людей катаются на коньках. За городом катаются на беговых и горных лыжах, а дети лепят снеговиков и играют в снежки.",
            "Весной снег тает, и россияне празднуют Масленицу — праздник проводов зимы с блинами. Лето в России тёплое и солнечное, а осень восхищает золотыми красками листьев."
        ],
        "questions": [
            {
                "id": "a2_r3_q1",
                "lang": "ru",
                "prompt": "Сколько месяцев обычно длится зима в центральной России?",
                "hint": "Đoạn 2 nêu các tháng từ cuối tháng 11 đến tháng 3.",
                "options": [
                    {"key": "A", "text": "Около двух месяцев"},
                    {"key": "B", "text": "С конца ноября по март (около 4 месяцев)"},
                    {"key": "C", "text": "Почти полгода (8 месяцев)"},
                    {"key": "D", "text": "Только в январе"}
                ],
                "ans": "B",
                "explain": "Đoạn 2 viết: «Зима в центральной России длится с конца ноября по март»."
            },
            {
                "id": "a2_r3_q2",
                "lang": "ru",
                "prompt": "Почему в русских квартирах тепло, даже когда на улице сильный мороз?",
                "hint": "Tìm nguyên nhân ở đoạn 2.",
                "options": [
                    {"key": "A", "text": "Люди включают газовые плиты"},
                    {"key": "B", "text": "Отлично работает центральное отопление"},
                    {"key": "C", "text": "В домах очень маленькие окна"},
                    {"key": "D", "text": "Люди спят в тёплых куртках"}
                ],
                "ans": "B",
                "explain": "Đoạn 2 giải thích: «потому что отлично работает центральное отопление»."
            },
            {
                "id": "a2_r3_q3",
                "lang": "ru",
                "prompt": "Какой традиционный праздник отмечают в конце зимы с блинами?",
                "hint": "Tìm tên lễ hội tiễn mùa đông ở đoạn cuối.",
                "options": [
                    {"key": "A", "text": "Новый год"},
                    {"key": "B", "text": "Масленица"},
                    {"key": "C", "text": "Рождество"},
                    {"key": "D", "text": "День победы"}
                ],
                "ans": "B",
                "explain": "Đoạn cuối nêu rõ: «россияне празднуют Масленицу — праздник проводов зимы с блинами»."
            }
        ]
    },
    {
        "id": 4,
        "level": "A2",
        "skill": "reading",
        "title": "Здоровье: Визит к врачу и покупка лекарств",
        "badge": "A2 Đọc hiểu",
        "readingTime": "15 phút",
        "source": "Tình huống giao tiếp y tế thực tế",
        "topic": "Sức khỏe, Triệu chứng bệnh & Nhà thuốc",
        "vocabCategories": [
            {
                "title": "Triệu chứng & Cảm giác (Симптомы и самочувствие)",
                "rows": [
                    {"word": "просту́да", "gram": "danh từ giống cái", "meaning": "cảm lạnh", "example": "Осенью многие болеют простудой."},
                    {"word": "температу́ра", "gram": "danh từ giống cái", "meaning": "nhiệt độ cơ thể, sốt", "example": "У меня высокая температура — 38 градусов."},
                    {"word": "боле́ть", "gram": "động từ", "meaning": "đau, nhức", "example": "У меня сильно болит горло и голова."},
                    {"word": "ка́шель", "gram": "danh từ giống đực", "meaning": "cơn ho", "example": "Принимайте этот сироп от кашля."}
                ]
            },
            {
                "title": "Khám bệnh & Thuốc men (Врач и лекарства)",
                "rows": [
                    {"word": "поликли́ника", "gram": "danh từ giống cái", "meaning": "phòng khám đa khoa", "example": "Запишитесь на приём в поликлинику."},
                    {"word": "реце́пт", "gram": "danh từ giống đực", "meaning": "đơn thuốc", "example": "Врач выписал рецепт на антибиотики."},
                    {"word": "табле́тки", "gram": "danh từ số nhiều", "meaning": "thuốc viên", "example": "Пейте эти таблетки два раза в день после еды."}
                ]
            }
        ],
        "readingText": [
            "Каждый иностранец, который живёт или учится в России, должен знать, как действовать, если он заболел. В России развита система бесплатной и платной медицины.",
            "Студенты обычно обращаются в университетскую поликлинику. Сначала нужно прийти к врачу-терапевту. Доктор слушает лёгкие, измеряет давление и температуру, осматривает горло и расспрашивает о симптомах.",
            "Если у пациента обычная простуда или грипп, врач выписывает больничный лист и рецепт на лекарства. Доктор рекомендует соблюдать постельный режим, пить горячий чай с лимоном и мёдом и принимать витамины.",
            "В аптеке большинство простых лекарств (жаропонижающие, капли в нос, леденцы от горла) можно купить без рецепта. Фармацевт всегда поможет выбрать подходящее средство."
        ],
        "questions": [
            {
                "id": "a2_r4_q1",
                "lang": "ru",
                "prompt": "К какому специалисту обращаются в поликлинике в первую очередь?",
                "hint": "Tìm tên bác sĩ đa khoa ban đầu ở đoạn 2.",
                "options": [
                    {"key": "A", "text": "К хирургу"},
                    {"key": "B", "text": "К терапевту"},
                    {"key": "C", "text": "К стоматологу"},
                    {"key": "D", "text": "К окулисту"}
                ],
                "ans": "B",
                "explain": "Đoạn 2 nêu rõ: «Сначала нужно прийти к врачу-терапевту»."
            },
            {
                "id": "a2_r4_q2",
                "lang": "ru",
                "prompt": "Что рекомендует врач при обычной простуде?",
                "hint": "Tìm lời khuyên của bác sĩ ở đoạn 3.",
                "options": [
                    {"key": "A", "text": "Идти на работу и заниматься спортом"},
                    {"key": "B", "text": "Постельный режим, тёплый чай с мёдом и витамины"},
                    {"key": "C", "text": "Срочно сделать операцию"},
                    {"key": "D", "text": "Пить холодную воду со льдом"}
                ],
                "ans": "B",
                "explain": "Đoạn 3 viết: nghỉ ngơi tại giường (постельный режим), uống trà nóng với mật ong và chanh."
            }
        ]
    },
    {
        "id": 5,
        "level": "A2",
        "skill": "reading",
        "title": "Русская дача: Традиция загородного отдыха",
        "badge": "A2 Đọc hiểu",
        "readingTime": "15 phút",
        "source": "Văn hóa & Đời sống người Nga",
        "topic": "Sở thích, Kỳ nghỉ cuối tuần & Văn hóa Dacha",
        "vocabCategories": [
            {
                "title": "Nghỉ ngơi & Nông thôn (Отдых за городом)",
                "rows": [
                    {"word": "да́ча", "gram": "danh từ giống cái", "meaning": "nhà vườn ngoại ô", "example": "Летом мы каждые выходные ездим на дачу."},
                    {"word": "приро́да", "gram": "danh từ giống cái", "meaning": "thiên nhiên", "example": "В деревне красивая и чистая природа."},
                    {"word": "шашлы́к", "gram": "danh từ giống đực", "meaning": "thịt nướng xiên kiểu Nga", "example": "На даче мы жарим вкусный шашлык."},
                    {"word": "ба́ня", "gram": "danh từ giống cái", "meaning": "nhà xông hơi truyền thống", "example": "Русская баня полезна для здоровья."}
                ]
            },
            {
                "title": "Làm vườn & Thu hoạch (Сад и огород)",
                "rows": [
                    {"word": "огоро́д", "gram": "danh từ giống đực", "meaning": "vườn rau", "example": "Бабушка выращивает помидоры на огороде."},
                    {"word": "я́годы", "gram": "danh từ số nhiều", "meaning": "quả mọng (dâu, mâm xôi)", "example": "Летом в саду созревают сладкие ягоды."},
                    {"word": "урожа́й", "gram": "danh từ giống đực", "meaning": "mùa thu hoạch", "example": "Осенью мы собираем богатый урожай яблок."}
                ]
            }
        ],
        "readingText": [
            "Дача — это уникальное явление русской культуры. Почти у каждой второй семьи в России есть небольшой загородный дом с земельным участком.",
            "История дачи началась ещё в XVIII веке, но особенно популярной она стала в советское время. В пятницу вечером на выездах из больших городов собираются длинные пробки: миллионы людей едут на дачу отдыхать от шума мегаполиса.",
            "На даче у каждого своё любимое занятие. Старшее поколение с удовольствием работает в саду и огороде: сажают цветы, картофель, огурцы и собирают клубнику. Молодёжь предпочитает жарить шашлыки, купаться в речке, париться в настоящей русской бане и играть на гитаре у костра.",
            "Осенью дачники варят домашнее варенье из ягод и делают солёные заготовки на всю долгую зиму."
        ],
        "questions": [
            {
                "id": "a2_r5_q1",
                "lang": "ru",
                "prompt": "Почему в пятницу вечером на выездах из городов образуются пробки?",
                "hint": "Tìm nguyên nhân kẹt xe cuối tuần ở đoạn 2.",
                "options": [
                    {"key": "A", "text": "Люди едут в аэропорт"},
                    {"key": "B", "text": "Миллионы горожан едут на дачи отдыхать"},
                    {"key": "C", "text": "Ремонтируют все загородные дороги"},
                    {"key": "D", "text": "Проходит футбольный матч"}
                ],
                "ans": "B",
                "explain": "Đoạn 2 nêu rõ: hàng triệu người đổ về nhà vườn dacha nghỉ ngơi cuối tuần."
            },
            {
                "id": "a2_r5_q2",
                "lang": "ru",
                "prompt": "Чем обычно занимается молодёжь на даче?",
                "hint": "Xem hoạt động của giới trẻ ở đoạn 3.",
                "options": [
                    {"key": "A", "text": "Только поливают помидоры"},
                    {"key": "B", "text": "Жарят шашлыки, купаются в реке и парятся в бане"},
                    {"key": "C", "text": "Спят целый день в доме"},
                    {"key": "D", "text": "Пишут научные статьи"}
                ],
                "ans": "B",
                "explain": "Đoạn 3 kể: «Молодёжь предпочитает жарить шашлыки, купаться в речке, париться в настоящей русской бане...»."
            }
        ]
    }
]

# Write Reading A2
os.makedirs(f"{BASE_DIR}/reading/a2", exist_ok=True)
manifest_r = []
for l in a2_reading_lessons:
    filename = f"lesson-{l['id']:02d}.json"
    with open(f"{BASE_DIR}/reading/a2/{filename}", "w", encoding="utf-8") as f:
        json.dump(l, f, ensure_ascii=False, indent=2)
    manifest_r.append({
        "id": l["id"],
        "file": filename,
        "title": l["title"],
        "badge": l["badge"],
        "readingTime": l["readingTime"],
        "topic": l["topic"],
        "questionCount": len(l["questions"])
    })
with open(f"{BASE_DIR}/reading/a2/manifest.json", "w", encoding="utf-8") as f:
    json.dump(manifest_r, f, ensure_ascii=False, indent=2)

print("A2 Reading generated successfully.")

# ==============================================================================
# 2. A2 LISTENING LESSONS
# ==============================================================================
a2_listening_lessons = [
    {
        "id": 1,
        "level": "A2",
        "skill": "listening",
        "title": "Как доехать до центра: Покупка билета и ориентирование",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Hỏi đường đi tàu điện ngầm và mua thẻ giao thông",
        "dialogue": [
            {"speaker": "Турист", "text": "Извините, пожалуйста! Вы не подскажете, как доехать до Красной площади?"},
            {"speaker": "Прохожий", "text": "Конечно! Вам нужно спуститься в метро на станцию «Университет»."},
            {"speaker": "Турист", "text": "А какую линию мне выбрать?"},
            {"speaker": "Прохожий", "text": "Садитесь на красную ветку и поезжайте прямо до станции «Охотный ряд». Это центр города."},
            {"speaker": "Турист", "text": "А сколько станций нужно проехать?"},
            {"speaker": "Прохожий", "text": "Примерно шесть остановок, это займёт около пятнадцати минут."},
            {"speaker": "Турист", "text": "Спасибо огромное! А где можно купить билет?"},
            {"speaker": "Прохожий", "text": "Прямо в вестибюле метро в кассе или в электронном автомате. Возьмите карту «Тройка» — это выгодно."},
            {"speaker": "Турист", "text": "Благодарю вас за помощь! Хорошего дня!"}
        ],
        "dictation": {
            "title": "Chép chính tả (Диктант)",
            "instruction": "Nghe kỹ câu sau và gõ lại chính xác bằng tiếng Nga vào ô bên dưới:",
            "targetText": "Садитесь на красную ветку и поезжайте прямо до центра.",
            "hint": "Chú ý viết hoa chữ cái đầu và dùng đúng thể động từ mệnh lệnh «Садитесь», «поезжайте»."
        },
        "questions": [
            {
                "id": "a2_l1_q1",
                "lang": "ru",
                "prompt": "Куда хочет доехать турист?",
                "hint": "Xác định điểm đến ở câu chào hỏi đầu tiên.",
                "options": [
                    {"key": "A", "text": "В аэропорт"},
                    {"key": "B", "text": "На Красную площадь"},
                    {"key": "C", "text": "На вокзал"},
                    {"key": "D", "text": "В университет"}
                ],
                "ans": "B",
                "explain": "Khách hỏi: «как доехать до Красной площади?»."
            },
            {
                "id": "a2_l1_q2",
                "lang": "ru",
                "prompt": "Сколько времени займёт поездка на метро?",
                "hint": "Người qua đường ước lượng thời gian đi tàu.",
                "options": [
                    {"key": "A", "text": "Около пяти минут"},
                    {"key": "B", "text": "Около пятнадцати минут"},
                    {"key": "C", "text": "Больше часа"},
                    {"key": "D", "text": "Два часа"}
                ],
                "ans": "B",
                "explain": "Đoạn thoại nêu: «это займёт около пятнадцати минут»."
            }
        ]
    },
    {
        "id": 2,
        "level": "A2",
        "skill": "listening",
        "title": "Заказ столика и ужин в ресторане русской кухни",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Đặt bàn và gọi các món ăn truyền thống Nga",
        "dialogue": [
            {"speaker": "Администратор", "text": "Добрый вечер! Добро пожаловать в ресторан «Тайга»! Вы бронировали столик?"},
            {"speaker": "Гость", "text": "Добрый вечер! Да, на фамилию Нгуен, на двоих на семь часов."},
            {"speaker": "Администратор", "text": "Да, всё верно. Проходите, пожалуйста, ваш столик у окна. Вот наше меню."},
            {"speaker": "Официант", "text": "Здравствуйте! Готовы сделать заказ или подсказать что-нибудь фирменное?"},
            {"speaker": "Гость", "text": "Посоветуйте, пожалуйста, традиционные русские супы."},
            {"speaker": "Официант", "text": "Очень рекомендую сибирский борщ с пампушками и грибную солянку."},
            {"speaker": "Гость", "text": "Отлично! Принесите, пожалуйста, две порции борща со сметаной и пельмени с говядиной."},
            {"speaker": "Официант", "text": "А из напитков что будете пить?"},
            {"speaker": "Гость", "text": "Кувшин домашнего клюквенного морса, пожалуйста."},
            {"speaker": "Официант", "text": "Прекрасный выбор! Заказ будет готов через двадцать минут."}
        ],
        "dictation": {
            "title": "Chép chính tả (Диктант)",
            "instruction": "Nghe kỹ câu sau và gõ lại chính xác bằng tiếng Nga vào ô bên dưới:",
            "targetText": "Принесите, пожалуйста, две порции горячего борща со сметаной.",
            "hint": "Gợi ý: Chú ý dấu phẩy sau «пожалуйста» và danh từ giống cái biến cách «две порции»."
        },
        "questions": [
            {
                "id": "a2_l2_q1",
                "lang": "ru",
                "prompt": "На сколько человек забронирован столик?",
                "hint": "Xem số lượng người ở lượt thoại thứ 2.",
                "options": [
                    {"key": "A", "text": "На одного"},
                    {"key": "B", "text": "На двоих"},
                    {"key": "C", "text": "На четверых"},
                    {"key": "D", "text": "На большую компанию"}
                ],
                "ans": "B",
                "explain": "Khách nói rõ: «на двоих на семь часов»."
            },
            {
                "id": "a2_l2_q2",
                "lang": "ru",
                "prompt": "Какой напиток заказали гости?",
                "hint": "Tìm đồ uống truyền thống được chọn ở cuối bài đối thoại.",
                "options": [
                    {"key": "A", "text": "Чёрный кофе"},
                    {"key": "B", "text": "Домашний клюквенный морс"},
                    {"key": "C", "text": "Минеральную воду с газом"},
                    {"key": "D", "text": "Зелёный чай"}
                ],
                "ans": "B",
                "explain": "Khách chọn: «Кувшин домашнего клюквенного морса, пожалуйста»."
            }
        ]
    },
    {
        "id": 3,
        "level": "A2",
        "skill": "listening",
        "title": "Телефонный разговор: Самочувствие друга и визит врача",
        "listeningTime": "12 phút",
        "topic": "Hội thoại qua điện thoại: Hỏi thăm khi bạn bị ốm",
        "dialogue": [
            {"speaker": "Анна", "text": "Алло, Антон, привет! Почему ты сегодня не пришёл на лекцию по грамматике?"},
            {"speaker": "Антон", "text": "Привет, Анна! К сожалению, я вчера сильно простудился. У меня болит горло и температура тридцать восемь."},
            {"speaker": "Анна", "text": "Ой, как жаль! Ты уже вызывал врача на дом?"},
            {"speaker": "Антон", "text": "Да, утром приходил участковый терапевт, послушал меня и выписал рецепт."},
            {"speaker": "Анна", "text": "Какие лекарства тебе нужно купить? Я могу зайти в аптеку после занятий."},
            {"speaker": "Антон", "text": "Спасибо огромное за заботу! Мне нужны капли в нос и сироп от кашля."},
            {"speaker": "Анна", "text": "Хорошо, я всё куплю и занесу тебе вечером. А ещё я принесу мёд и лимон. Выздоравливай скорей!"},
            {"speaker": "Антон", "text": "Анна, ты настоящий друг! До встречи!"}
        ],
        "dictation": {
            "title": "Chép chính tả (Диктант)",
            "instruction": "Nghe kỹ câu sau và gõ lại chính xác bằng tiếng Nga vào ô bên dưới:",
            "targetText": "Я зайду в аптеку и принесу тебе лекарства вечером.",
            "hint": "Gợi ý: Chú ý động từ chuyển động có tiền tố «зайду» và «принесу»."
        },
        "questions": [
            {
                "id": "a2_l3_q1",
                "lang": "ru",
                "prompt": "Почему Антон не пришёл на университетскую лекцию?",
                "hint": "Xem lý do vắng mặt ở lượt thoại thứ 2.",
                "options": [
                    {"key": "A", "text": "Он проспал будильник"},
                    {"key": "B", "text": "Он заболел и простудился"},
                    {"key": "C", "text": "Он уехал в другой город"},
                    {"key": "D", "text": "Он готовился к экзамену"}
                ],
                "ans": "B",
                "explain": "Anton nói: «я вчера сильно простудился. У меня болит горло и температура...»."
            },
            {
                "id": "a2_l3_q2",
                "lang": "ru",
                "prompt": "Что Анна пообещала принести Антону вечером?",
                "hint": "Tìm các món đồ và thuốc Anna sẽ mua hộ.",
                "options": [
                    {"key": "A", "text": "Лекарства из аптеки, мёд и лимон"},
                    {"key": "B", "text": "Новый учебник и тетрадь"},
                    {"key": "C", "text": "Пиццу и колу"},
                    {"key": "D", "text": "Билеты в кино"}
                ],
                "ans": "A",
                "explain": "Anna hứa: mua thuốc ở tiệm thuốc kèm theo mật ong và chanh."
            }
        ]
    }
]

# Write Listening A2
os.makedirs(f"{BASE_DIR}/listening/a2", exist_ok=True)
manifest_l = []
for l in a2_listening_lessons:
    filename = f"lesson-{l['id']:02d}.json"
    with open(f"{BASE_DIR}/listening/a2/{filename}", "w", encoding="utf-8") as f:
        json.dump(l, f, ensure_ascii=False, indent=2)
    manifest_l.append({
        "id": l["id"],
        "file": filename,
        "title": l["title"],
        "badge": "A2 Nghe hiểu",
        "listeningTime": l["listeningTime"],
        "topic": l["topic"],
        "questionCount": len(l["questions"])
    })
with open(f"{BASE_DIR}/listening/a2/manifest.json", "w", encoding="utf-8") as f:
    json.dump(manifest_l, f, ensure_ascii=False, indent=2)

print("A2 Listening generated successfully.")

# ==============================================================================
# 3. A2 SPEAKING LESSONS
# ==============================================================================
a2_speaking_lessons = [
    {
        "id": 1,
        "level": "A2",
        "skill": "speaking",
        "title": "Мой родной город и любимые места",
        "speakingTime": "15 phút",
        "topic": "Miêu tả thành phố quê hương, danh lam thắng cảnh và phương tiện",
        "intro": "Luyện nói về nơi chôn rau cắt rốn hoặc thành phố bạn đang sinh sống. Sử dụng các cấu trúc chỉ vị trí, phương tiện giao thông và bày tỏ cảm xúc.",
        "shadowingPhrases": [
            {
                "id": "a2_spk_1",
                "text": "Я родился и вырос в красивом и древнем городе Ханое.",
                "meaning": "Tôi sinh ra và lớn lên ở thủ đô Hà Nội cổ kính và tươi đẹp.",
                "phonetic": "Ya rodilsya i vyros v krasivom i drevnem gorode Khanoye.",
                "tips": "«Родился» nhấn vào âm đầu /ro-/, «вырос» âm /y/ đọc sâu trong vòm họng."
            },
            {
                "id": "a2_spk_2",
                "text": "В моём городе много парков, зелёных озёр и старинных улиц.",
                "meaning": "Ở thành phố của tôi có nhiều công viên, hồ nước xanh và các con phố cổ.",
                "phonetic": "V moyom gorode mnogo parkov, zelyonykh ozyor i starinnykh ulits.",
                "tips": "Cách 2 số nhiều sau «много»: «парков», «озёр», «улиц»."
            },
            {
                "id": "a2_spk_3",
                "text": "Здесь очень развит общественный транспорт: автобусы и метро.",
                "meaning": "Ở đây giao thông công cộng rất phát triển: xe buýt và tàu điện ngầm.",
                "phonetic": "Zdes' ochen' razvit obshchestvennyy transport: avtobusy i metro.",
                "tips": "«Общественный» đọc âm /shch/ bật mạnh và dứt khoát."
            },
            {
                "id": "a2_spk_4",
                "text": "Я очень люблю гулять по набережной вечером с друзьями.",
                "meaning": "Tôi rất thích đi dạo dọc bờ kè/bờ sông vào buổi tối cùng bạn bè.",
                "phonetic": "Ya ochen' lyublyu gulyat' po naberezhnoy vecherom s druz'yami.",
                "tips": "Cụm «по набережной» dùng Cách 3 giống cái."
            }
        ],
        "freeSpeakingPrompt": {
            "title": "Luyện nói tự do (Монолог о городе)",
            "task": "Hãy thuyết trình ngắn (1 - 2 phút, khoảng 6 - 8 câu) kể về thành phố quê hương của bạn hoặc một thành phố bạn từng đến thăm.",
            "sampleGuidance": [
                "Как называется ваш город и где он находится? (Tên thành phố và vị trí địa lý)",
                "Какая там погода зимой и летом? (Khí hậu các mùa)",
                "Какие интересные достопримечательности там есть? (Danh lam thắng cảnh nổi bật)",
                "За что вы любите этот город? (Vì sao bạn yêu mến thành phố này?)"
            ]
        }
    },
    {
        "id": 2,
        "level": "A2",
        "skill": "speaking",
        "title": "Мой обычный день: Распорядок и привычки",
        "speakingTime": "15 phút",
        "topic": "Kể về lịch trình thường nhật, giờ giấc sinh hoạt và học tập",
        "intro": "Luyện sử dụng các động từ chỉ hành động thường ngày, cách nói giờ trong tiếng Nga và các liên từ chỉ thứ tự thời gian.",
        "shadowingPhrases": [
            {
                "id": "a2_spk_5",
                "text": "Обычно я просыпаюсь в семь часов утра и делаю зарядку.",
                "meaning": "Thông thường tôi thức dậy lúc 7 giờ sáng và tập thể dục.",
                "phonetic": "Obychno ya prosypayus' v sem' chasov utra i delayu zaryadku.",
                "tips": "«Просыпаюсь» đuôi /-s'/ đọc êm, không kéo dài."
            },
            {
                "id": "a2_spk_6",
                "text": "После плотного завтрака я еду на учёбу на автобусе.",
                "meaning": "Sau bữa ăn sáng no nê tôi đi xe buýt đến trường học.",
                "phonetic": "Posle plotnogo zavtraka ya yedu na uchyobu na avtobuse.",
                "tips": "Sau giới từ «после» dùng danh từ và tính từ ở Cách 2."
            },
            {
                "id": "a2_spk_7",
                "text": "Занятия в университете начинаются ровно в девять часов.",
                "meaning": "Các tiết học ở trường đại học bắt đầu đúng vào 9 giờ.",
                "phonetic": "Zanyatiya v universitete nachinayutsya rovno v devyat' chasov.",
                "tips": "«Начинаются» đọc đuôi «-тся» thành âm /-tsa/."
            },
            {
                "id": "a2_spk_8",
                "text": "Вечером я делаю домашние задания и повторяю новые слова.",
                "meaning": "Buổi tối tôi làm bài tập về nhà và ôn lại các từ vựng mới.",
                "phonetic": "Vecherom ya delayu domashnie zadaniya i povtoryayu novye slova.",
                "tips": "«Повторяю» nhấn trọng âm ở âm tiết cuối."
            }
        ],
        "freeSpeakingPrompt": {
            "title": "Luyện nói tự do (Мой распорядок дня)",
            "task": "Hãy mô tả một ngày bình thường của bạn từ sáng đến tối (khoảng 6–8 câu).",
            "sampleGuidance": [
                "Во сколько вы обычно встаёте? (Mấy giờ bạn thức dậy?)",
                "Что вы едите на завтрак, обед и ужин? (Bạn ăn gì các bữa?)",
                "Что вы делаете днём на учёбе или работе? (Ban ngày bạn học tập/làm việc ra sao?)",
                "Как вы отдыхаете перед сном? (Bạn thư giãn thế nào trước khi đi ngủ?)"
            ]
        }
    },
    {
        "id": 3,
        "level": "A2",
        "skill": "speaking",
        "title": "Планы на отпуск и путешествия",
        "speakingTime": "15 phút",
        "topic": "Bàn về kế hoạch du lịch, mua vé và đặt phòng khách sạn",
        "intro": "Luyện tập diễn đạt các kế hoạch trong tương lai với động từ hoàn thành (СВ) và cấu trúc «хотеть / планировать + động từ nguyên mẫu».",
        "shadowingPhrases": [
            {
                "id": "a2_spk_9",
                "text": "Следующим летом я планирую поехать в Санкт-Петербург.",
                "meaning": "Mùa hè tới tôi dự định sẽ đi du lịch Saint Petersburg.",
                "phonetic": "Sleduyushchim letom ya planiruyu poyekhat' v Sankt-Peterburg.",
                "tips": "«Следующим летом» dùng Cách 5 chỉ thời gian diễn ra hành động."
            },
            {
                "id": "a2_spk_10",
                "text": "Я хочу увидеть знаменитые белые ночи и разводные мосты.",
                "meaning": "Tôi muốn chiêm ngưỡng hiện tượng Đêm trắng nổi tiếng và cầu mở.",
                "phonetic": "Ya khochu uvidet' znamenitye belye nochi i razvodnye mosty.",
                "tips": "Động từ hoàn thành «увидеть» thể hiện kết quả mong muốn."
            },
            {
                "id": "a2_spk_11",
                "text": "Мы заранее забронировали уютный номер в гостинице рядом с центром.",
                "meaning": "Chúng tôi đã đặt trước một phòng khách sạn ấm cúng gần trung tâm.",
                "phonetic": "My zaranee zabronirovali uyutnyy nomer v gostinitse ryadom s tsentrom.",
                "tips": "Trọng âm từ «заброни́ровали» rơi vào âm tiết thứ 3."
            }
        ],
        "freeSpeakingPrompt": {
            "title": "Luyện nói tự do (Планы на путешествие)",
            "task": "Kể về chuyến du lịch lý tưởng mà bạn đang ấp ủ hoặc chuẩn bị thực hiện.",
            "sampleGuidance": [
                "Куда именно вы хотите поехать? (Bạn muốn đi đâu?)",
                "С кем вы поедете и на каком транспорте? (Đi cùng ai và bằng phương tiện gì?)",
                "Что вы собираетесь там посмотреть? (Bạn dự định ngắm nhìn những gì?)",
                "Сколько дней продлится поездка? (Chuyến đi kéo dài bao nhiêu ngày?)"
            ]
        }
    }
]

# Write Speaking A2
os.makedirs(f"{BASE_DIR}/speaking/a2", exist_ok=True)
manifest_s = []
for l in a2_speaking_lessons:
    filename = f"lesson-{l['id']:02d}.json"
    with open(f"{BASE_DIR}/speaking/a2/{filename}", "w", encoding="utf-8") as f:
        json.dump(l, f, ensure_ascii=False, indent=2)
    manifest_s.append({
        "id": l["id"],
        "file": filename,
        "title": l["title"],
        "badge": "A2 Nói",
        "speakingTime": l["speakingTime"],
        "topic": l["topic"],
        "questionCount": len(l["shadowingPhrases"])
    })
with open(f"{BASE_DIR}/speaking/a2/manifest.json", "w", encoding="utf-8") as f:
    json.dump(manifest_s, f, ensure_ascii=False, indent=2)

print("A2 Speaking generated successfully.")

# ==============================================================================
# 4. A2 WRITING LESSONS
# ==============================================================================
a2_writing_lessons = [
    {
        "id": 1,
        "level": "A2",
        "skill": "writing",
        "title": "Поздравительная открытка другу с днём рождения",
        "writingTime": "25 phút",
        "topic": "Viết thiệp chúc mừng sinh nhật cho bạn bè người Nga",
        "targetWordRange": [50, 75],
        "task": "Hãy viết một bưu thiếp chúc mừng sinh nhật (50–75 từ) gửi cho bạn cùng lớp người Nga tên là Denis.",
        "requirements": [
            "1. Lời chúc mừng mở đầu trang trọng hoặc thân mật (Дорогой Денис!).",
            "2. Chúc sức khỏe, thành công trong học tập và niềm vui (Здоровья, успехов в учёбе).",
            "3. Nhắc lại một kỷ niệm đẹp hoặc món quà bạn gửi tặng.",
            "4. Lời mời gặp mặt ăn mừng cuối tuần.",
            "5. Lời kết và ký tên thân mật."
        ],
        "cliches": [
            {"ru": "От всей души поздравляю тебя с днём рождения!", "vi": "Từ tận đáy lòng xin chúc mừng sinh nhật bạn!"},
            {"ru": "Желаю тебе крепкого здоровья, счастья и удачи!", "vi": "Chúc bạn dồi dào sức khỏe, hạnh phúc và may mắn!"},
            {"ru": "Пусть все твои мечты сбываются!", "vi": "Chúc cho mọi ước mơ của bạn đều trở thành hiện thực!"},
            {"ru": "С нетерпением жду нашей встречи в субботу.", "vi": "Háo hức mong chờ buổi gặp mặt của chúng mình vào thứ Bảy."}
        ],
        "modelAnswer": "Дорогой Денис!\n\nОт всей души поздравляю тебя с днём рождения! Желаю тебе крепкого здоровья, отличного настроения и больших успехов в учёбе. Пусть все твои заветные мечты обязательно сбудутся!\n\nЯ приготовил для тебя небольшой подарок из Вьетнама. Надеюсь, мы увидимся в субботу и отпразднуем этот день вместе.\n\nТвой друг, Ан.",
        "modelWordCount": 54,
        "rubric": [
            {"criteria": "Bố cục thiệp chúc mừng", "desc": "Có lời xưng hô mở đầu, lời chúc chính và chữ ký kết thúc."},
            {"criteria": "Cấu trúc lời chúc (Cách 2)", "desc": "Sử dụng chính xác cấu trúc «Желаю тебе + Cách 2»."},
            {"criteria": "Ngữ pháp & Chính tả", "desc": "Chia đúng thì và không mắc lỗi chính tả căn bản."}
        ]
    },
    {
        "id": 2,
        "level": "A2",
        "skill": "writing",
        "title": "Электронное письмо другу: Приглашение на выставку",
        "writingTime": "25 phút",
        "topic": "Viết email thân mật rủ bạn đi xem triển lãm nghệ thuật",
        "targetWordRange": [60, 85],
        "task": "Viết một bức thư điện tử (60–85 từ) rủ bạn cùng phòng đi xem triển lãm tranh tại bảo tàng vào Chủ nhật.",
        "requirements": [
            "1. Lời chào và hỏi thăm sức khỏe.",
            "2. Nêu mục đích viết thư: rủ đi triển lãm tranh.",
            "3. Cung cấp thông tin: địa điểm, thời gian và giá vé sinh viên.",
            "4. Hẹn địa điểm gặp nhau (trước cửa ga metro).",
            "5. Đề nghị bạn phản hồi sớm."
        ],
        "cliches": [
            {"ru": "Как твои дела? Надеюсь, ты хорошо себя чувствуешь.", "vi": "Bạn thế nào? Mong bạn vẫn khỏe."},
            {"ru": "Я пишу, чтобы пригласить тебя на новую выставку.", "vi": "Tôi viết thư này để mời bạn cùng đi xem triển lãm mới."},
            {"ru": "Выставка проходит в центральном музее.", "vi": "Triển lãm diễn ra ở bảo tàng trung tâm."},
            {"ru": "Давай встретимся около входа в метро в два часа.", "vi": "Chúng mình hẹn gặp nhau gần lối vào metro lúc 2 giờ nhé."}
        ],
        "modelAnswer": "Привет, Максим!\n\nКак твои дела? Надеюсь, ты свободен в эти выходные. Я пишу тебе, чтобы пригласить на интересную художественную выставку в центре города. Там показывают редкие картины современных художников.\n\nДля студентов вход стоит всего двести рублей. Давай встретимся в воскресенье в два часа дня возле станции метро «Театральная». Напиши мне, сможешь ли ты пойти!\n\nДо встречи,\nМинх.",
        "modelWordCount": 63,
        "rubric": [
            {"criteria": "Nội dung & Mục đích", "desc": "Truyền đạt đầy đủ thời gian, địa điểm và đề nghị rủ bạn."},
            {"criteria": "Văn phong thân mật", "desc": "Sử dụng đại từ «ты», lối hành văn tự nhiên và các liên từ mạch lạc."},
            {"criteria": "Từ vựng & Ngữ pháp A2", "desc": "Sử dụng đúng giới từ chỉ thời gian và địa điểm."}
        ]
    },
    {
        "id": 3,
        "level": "A2",
        "skill": "writing",
        "title": "Короткое эссе: Моё любимое время года",
        "writingTime": "30 phút",
        "topic": "Viết đoạn văn ngắn miêu tả mùa yêu thích trong năm",
        "targetWordRange": [70, 95],
        "task": "Viết một bài văn ngắn (70–95 từ) kể về mùa mà bạn thích nhất trong năm và lý do vì sao.",
        "requirements": [
            "1. Giới thiệu mùa yêu thích của bạn.",
            "2. Miêu tả đặc điểm thời tiết và cảnh sắc thiên nhiên vào mùa đó.",
            "3. Kể những hoạt động bạn thường làm khi mùa đó đến.",
            "4. Cảm xúc và ấn tượng đặc biệt của bạn về mùa này."
        ],
        "cliches": [
            {"ru": "Моё самое любимое время года — это...", "vi": "Mùa mà tôi yêu thích nhất trong năm là..."},
            {"ru": "В это время года на улице стоит прекрасная погода.", "vi": "Vào mùa này thời tiết ngoài trời rất tuyệt vời."},
            {"ru": "Деревья становятся золотыми и жёлтыми.", "vi": "Cây cối ngả sang màu vàng ươm rực rỡ."},
            {"ru": "Я очень люблю гулять в парке и делать фотографии.", "vi": "Tôi rất thích đi dạo trong công viên và chụp ảnh."}
        ],
        "modelAnswer": "Моё самое любимое время года в России — это золотая осень. Она начинается в сентябре и продолжается до середины октября.\n\nВ это время года погода ещё достаточно тёплая, а воздух удивительно свежий и прозрачный. Все парки и леса становятся невероятно красивыми: листья на деревьях окрашиваются в жёлтый, оранжевый и красный цвета.\n\nЯ обожаю гулять по старым аллеям, шуршать опавшими листьями и фотографировать природу. Осень вдохновляет меня на творчество и учёбу.",
        "modelWordCount": 71,
        "rubric": [
            {"criteria": "Chủ đề & Bố cục", "desc": "Bài viết có mở bài, thân bài miêu tả và cảm nghĩ kết bài."},
            {"criteria": "Tính từ & Từ vựng miêu tả", "desc": "Sử dụng phong phú tính từ chỉ màu sắc, cảm xúc và thời tiết."},
            {"criteria": "Độ dài & Ngữ pháp", "desc": "Đạt số lượng từ yêu cầu, sử dụng đúng các thì của động từ."}
        ]
    }
]

# Write Writing A2
os.makedirs(f"{BASE_DIR}/writing/a2", exist_ok=True)
manifest_w = []
for l in a2_writing_lessons:
    filename = f"lesson-{l['id']:02d}.json"
    with open(f"{BASE_DIR}/writing/a2/{filename}", "w", encoding="utf-8") as f:
        json.dump(l, f, ensure_ascii=False, indent=2)
    manifest_w.append({
        "id": l["id"],
        "file": filename,
        "title": l["title"],
        "badge": "A2 Viết",
        "writingTime": l["writingTime"],
        "topic": l["topic"],
        "targetWordRange": l["targetWordRange"]
    })
with open(f"{BASE_DIR}/writing/a2/manifest.json", "w", encoding="utf-8") as f:
    json.dump(manifest_w, f, ensure_ascii=False, indent=2)

print("A2 Writing generated successfully.")

