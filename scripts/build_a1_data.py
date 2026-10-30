# -*- coding: utf-8 -*-
import json

reading_a1 = {
    1: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "студе́нт", "gram": "danh từ giống đực", "meaning": "nam sinh viên", "example": "Антон — иностранный студент."},
            {"word": "студе́нтка", "gram": "danh từ giống cái", "meaning": "nữ sinh viên", "example": "Анна — прилежная студентка."},
            {"word": "университе́т", "gram": "danh từ giống đực", "meaning": "trường đại học", "example": "Мы учимся в московском университете."},
            {"word": "преподава́тель", "gram": "danh từ giống đực", "meaning": "giảng viên, thầy giáo", "example": "Наш преподаватель очень умный человек."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "знако́миться", "gram": "động từ chưa hoàn thành", "meaning": "làm quen", "example": "Давайте познакомимся!"},
            {"word": "познако́миться", "gram": "động từ hoàn thành", "meaning": "làm quen (đã xong)", "example": "Очень приятно с вами познакомиться."},
            {"word": "учи́ться", "gram": "động từ chưa hoàn thành", "meaning": "học tập", "example": "Они учатся на первом курсе."},
            {"word": "звать", "gram": "động từ chưa hoàn thành", "meaning": "gọi tên", "example": "Как вас зовут? — Меня зовут Антон."}
        ]}
    ],
    2: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "стол", "gram": "danh từ giống đực", "meaning": "cái bàn", "example": "Это большой письменный стол."},
            {"word": "стул", "gram": "danh từ giống đực", "meaning": "cái ghế", "example": "Около стола стоит удобный стул."},
            {"word": "кни́га", "gram": "danh từ giống cái", "meaning": "cuốn sách", "example": "На столе лежит интересная книга."},
            {"word": "слова́рь", "gram": "danh từ giống đực", "meaning": "cuốn từ điển", "example": "Это новый русско-вьетнамский словарь."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "лежа́ть", "gram": "động từ chưa hoàn thành", "meaning": "nằm ở", "example": "Тетрадь и ручка лежат на столе."},
            {"word": "стоя́ть", "gram": "động từ chưa hoàn thành", "meaning": "đứng ở", "example": "Лампа стоит на тумбочке."},
            {"word": "сиде́ть", "gram": "động từ chưa hoàn thành", "meaning": "ngồi", "example": "Студенты сидят в светлой аудитории."},
            {"word": "чита́ть", "gram": "động từ chưa hoàn thành", "meaning": "đọc", "example": "Мы читаем русский текст."}
        ]}
    ],
    3: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "рюкза́к", "gram": "danh từ giống đực", "meaning": "ba lô", "example": "Это мой синий рюкзак."},
            {"word": "ру́чка", "gram": "danh từ giống cái", "meaning": "cây bút bi", "example": "Чья это чёрная ручка?"},
            {"word": "я́блоко", "gram": "danh từ giống trung", "meaning": "quả táo", "example": "На тарелке лежит сладкое яблоко."},
            {"word": "кре́сло", "gram": "danh từ giống trung", "meaning": "ghế bành tựa", "example": "В комнате стоит мягкое кресло."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "ку́шать", "gram": "động từ chưa hoàn thành", "meaning": "ăn (thân mật)", "example": "Кушай яблоко на здоровье!"},
            {"word": "ви́деть", "gram": "động từ chưa hoàn thành", "meaning": "nhìn thấy", "example": "Я вижу твой дом издалека."},
            {"word": "брать", "gram": "động từ chưa hoàn thành", "meaning": "lấy, cầm lấy", "example": "Студент берёт учебник в библиотеке."},
            {"word": "писа́ть", "gram": "động từ chưa hoàn thành", "meaning": "viết", "example": "Он пишет домашнее задание ручкой."}
        ]}
    ],
    4: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "аудито́рия", "gram": "danh từ giống cái", "meaning": "giảng đường", "example": "Наша аудитория чистая и светлая."},
            {"word": "доска́", "gram": "danh từ giống cái", "meaning": "bảng viết", "example": "Учитель пишет правила на доске."},
            {"word": "окно́", "gram": "danh từ giống trung", "meaning": "cửa sổ", "example": "Около окна стоят парты."},
            {"word": "дверь", "gram": "danh từ giống cái", "meaning": "cửa ra vào", "example": "Справа находится входная дверь."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "висе́ть", "gram": "động từ chưa hoàn thành", "meaning": "treo ở", "example": "На стене висит географическая карта."},
            {"word": "находи́ться", "gram": "động từ chưa hoàn thành", "meaning": "tọa lạc, nằm ở", "example": "Где находится лаборатория?"},
            {"word": "занима́ться", "gram": "động từ chưa hoàn thành", "meaning": "học tập, làm bài", "example": "Здесь студенты занимаются каждый день."},
            {"word": "слу́шать", "gram": "động từ chưa hoàn thành", "meaning": "lắng nghe", "example": "Мы внимательно слушаем лекцию."}
        ]}
    ],
    5: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "инжене́р", "gram": "danh từ giống đực", "meaning": "kỹ sư", "example": "Борис — опытный инженер."},
            {"word": "врач", "gram": "danh từ giống đực", "meaning": "bác sĩ", "example": "Елена работает врачом в больнице."},
            {"word": "заво́д", "gram": "danh từ giống đực", "meaning": "nhà máy", "example": "Он работает на автомобильном заводе."},
            {"word": "больни́ца", "gram": "danh từ giống cái", "meaning": "bệnh viện", "example": "Больница находится рядом с метро."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "рабо́тать", "gram": "động từ chưa hoàn thành", "meaning": "làm việc", "example": "Кем вы работаете?"},
            {"word": "лечи́ть", "gram": "động từ chưa hoàn thành", "meaning": "chữa bệnh", "example": "Врач лечит больных детей."},
            {"word": "мечта́ть", "gram": "động từ chưa hoàn thành", "meaning": "mơ ước", "example": "Я мечтаю стать программистом."},
            {"word": "строи́ть", "gram": "động từ chưa hoàn thành", "meaning": "xây dựng", "example": "Инженер строит современные дома."}
        ]}
    ],
    6: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "день", "gram": "danh từ giống đực", "meaning": "ngày, ban ngày", "example": "Каждый день мы изучаем новые слова."},
            {"word": "у́тро", "gram": "danh từ giống trung", "meaning": "buổi sáng", "example": "Утром я пью зелёный чай."},
            {"word": "ве́чер", "gram": "danh từ giống đực", "meaning": "buổi tối", "example": "Вечером семья собирается вместе."},
            {"word": "за́втрак", "gram": "danh từ giống đực", "meaning": "bữa ăn sáng", "example": "На завтрак мы едим кашу и яйца."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "де́лать", "gram": "động từ chưa hoàn thành", "meaning": "làm", "example": "Что вы делаете вечером?"},
            {"word": "повторя́ть", "gram": "động từ chưa hoàn thành", "meaning": "nhắc lại, ôn tập", "example": "Студенты повторяют грамматику дома."},
            {"word": "за́втракать", "gram": "động từ chưa hoàn thành", "meaning": "ăn bữa sáng", "example": "Я завтракаю в восемь часов."},
            {"word": "отдыха́ть", "gram": "động từ chưa hoàn thành", "meaning": "nghỉ ngơi", "example": "В воскресенье мы отдыхаем на даче."}
        ]}
    ],
    7: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "язы́к", "gram": "danh từ giống đực", "meaning": "ngôn ngữ, tiếng", "example": "Русский язык очень богатый и красивый."},
            {"word": "сло́во", "gram": "danh từ giống trung", "meaning": "từ ngữ", "example": "Как пишется это трудное слово?"},
            {"word": "фра́за", "gram": "danh từ giống cái", "meaning": "cụm từ, câu nói", "example": "Повторите эту полезную фразу."},
            {"word": "диало́г", "gram": "danh từ giống đực", "meaning": "đoạn hội thoại", "example": "Студенты читают диалог по ролям."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "говори́ть", "gram": "động từ chưa hoàn thành", "meaning": "nói", "example": "Вы говорите по-русски?"},
            {"word": "понима́ть", "gram": "động từ chưa hoàn thành", "meaning": "hiểu", "example": "Я хорошо понимаю этот текст."},
            {"word": "изуча́ть", "gram": "động từ chưa hoàn thành", "meaning": "nghiên cứu, học sâu", "example": "Мы изучаем русскую литературу."},
            {"word": "спра́шивать", "gram": "động từ chưa hoàn thành", "meaning": "hỏi", "example": "Ученик спрашивает учителя о правиле."}
        ]}
    ],
    8: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "семья́", "gram": "danh từ giống cái", "meaning": "gia đình", "example": "У меня большая и дружная семья."},
            {"word": "роди́тели", "gram": "danh từ số nhiều", "meaning": "cha mẹ, phụ huynh", "example": "Мои родители живут в Ханое."},
            {"word": "брат", "gram": "danh từ giống đực", "meaning": "anh/em trai", "example": "Мой старший брат работает программистом."},
            {"word": "сестра́", "gram": "danh từ giống cái", "meaning": "chị/em gái", "example": "Моя младшая сестра учится в школе."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "жить", "gram": "động từ chưa hoàn thành", "meaning": "sống, cư trú", "example": "Мы живём в уютной квартире."},
            {"word": "люби́ть", "gram": "động từ chưa hoàn thành", "meaning": "yêu quý, thích", "example": "Я очень люблю своих родителей."},
            {"word": "помога́ть", "gram": "động từ chưa hoàn thành", "meaning": "giúp đỡ", "example": "Брат помогает мне учить уроки."},
            {"word": "забо́титься", "gram": "động từ chưa hoàn thành", "meaning": "chăm sóc", "example": "Мама заботится обо всей семье."}
        ]}
    ],
    9: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "дом", "gram": "danh từ giống đực", "meaning": "ngôi nhà, tòa nhà", "example": "Наш дом находится в тихом районе."},
            {"word": "кварти́ра", "gram": "danh từ giống cái", "meaning": "căn hộ", "example": "Это трёхкомнатная светлая квартира."},
            {"word": "ку́хня", "gram": "danh từ giống cái", "meaning": "phòng bếp", "example": "На кухне мама готовит обед."},
            {"word": "эта́ж", "gram": "danh từ giống đực", "meaning": "tầng lầu", "example": "Мы живём на пятом этаже."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "гото́вить", "gram": "động từ chưa hoàn thành", "meaning": "nấu nướng", "example": "Бабушка готовит вкусный пирог."},
            {"word": "убира́ть", "gram": "động từ chưa hoàn thành", "meaning": "dọn dẹp", "example": "В субботу я убираю свою комнату."},
            {"word": "находи́ться", "gram": "động từ chưa hoàn thành", "meaning": "tọa lạc", "example": "Метро находится рядом с домом."},
            {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua sắm", "example": "Мы покупаем мебель для гостиной."}
        ]}
    ],
    10: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "го́род", "gram": "danh từ giống đực", "meaning": "thành phố", "example": "Москва — огромный красивый город."},
            {"word": "страна́", "gram": "danh từ giống cái", "meaning": "quốc gia, đất nước", "example": "Россия — самая большая страна в мире."},
            {"word": "столи́ца", "gram": "danh từ giống cái", "meaning": "thủ đô", "example": "Ханой — столица Вьетнама."},
            {"word": "у́лица", "gram": "danh từ giống cái", "meaning": "đường phố", "example": "На нашей улице много зелёных деревьев."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "приезжа́ть", "gram": "động từ chưa hoàn thành", "meaning": "đến nơi (bằng xe/tàu)", "example": "Иностранцы приезжают в Москву учиться."},
            {"word": "гуля́ть", "gram": "động từ chưa hoàn thành", "meaning": "đi dạo", "example": "Мы гуляем по Красной площади."},
            {"word": "путеше́ствовать", "gram": "động từ chưa hoàn thành", "meaning": "du lịch", "example": "Летом он любит путешествовать на поезде."},
            {"word": "знать", "gram": "động từ chưa hoàn thành", "meaning": "biết", "example": "Я хорошо знаю этот древний город."}
        ]}
    ],
    11: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "расписа́ние", "gram": "danh từ giống trung", "meaning": "thời khóa biểu", "example": "В расписании сегодня три пары."},
            {"word": "ле́кция", "gram": "danh từ giống cái", "meaning": "tiết giảng, bài giảng", "example": "Лекция по истории начинается в девять."},
            {"word": "обе́д", "gram": "danh từ giống đực", "meaning": "bữa trưa", "example": "После занятий студенты идут на обед."},
            {"word": "о́тдых", "gram": "danh từ giống đực", "meaning": "sự nghỉ ngơi", "example": "Вечер — лучшее время для отдыха."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "просыпа́ться", "gram": "động từ chưa hoàn thành", "meaning": "thức giấc", "example": "Я просыпаюсь рано утром."},
            {"word": "встава́ть", "gram": "động từ chưa hoàn thành", "meaning": "thức dậy", "example": "Он встаёт в семь часов утра."},
            {"word": "идти́", "gram": "động từ chưa hoàn thành", "meaning": "đi bộ", "example": "Студенты идут в университет пешком."},
            {"word": "возвраща́ться", "gram": "động từ chưa hoàn thành", "meaning": "quay trở về", "example": "Вечером я возвращаюсь в общежитие."}
        ]}
    ],
    12: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "магази́н", "gram": "danh từ giống đực", "meaning": "cửa hàng, siêu thị", "example": "Продуктовый магазин работает до ночи."},
            {"word": "хлеб", "gram": "danh từ giống đực", "meaning": "bánh mì", "example": "Купите, пожалуйста, свежий белый хлеб."},
            {"word": "молоко́", "gram": "danh từ giống trung", "meaning": "sữa", "example": "Дети пьют тёплое молоко с мёдом."},
            {"word": "сыр", "gram": "danh từ giống đực", "meaning": "phô mai", "example": "В магазине продаётся вкусный сыр."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua", "example": "Что вы хотите купить?"},
            {"word": "продава́ть", "gram": "động từ chưa hoàn thành", "meaning": "bán", "example": "Здесь продают свежие овощи."},
            {"word": "сто́ить", "gram": "động từ chưa hoàn thành", "meaning": "có giá tiền", "example": "Сколько стоит этот сыр?"},
            {"word": "плати́ть", "gram": "động từ chưa hoàn thành", "meaning": "thanh toán", "example": "Я плачу банковской картой."}
        ]}
    ],
    13: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "рестора́н", "gram": "danh từ giống đực", "meaning": "nhà hàng", "example": "Мы обедаем в русском ресторане."},
            {"word": "борщ", "gram": "danh từ giống đực", "meaning": "súp củ dền borscht", "example": "Русский борщ со сметаной очень вкусный."},
            {"word": "пирожо́к", "gram": "danh từ giống đực", "meaning": "bánh nướng nhỏ pirozhok", "example": "Горячие пирожки с мясом и капустой."},
            {"word": "счёт", "gram": "danh từ giống đực", "meaning": "hóa đơn", "example": "Официант, принесите счёт, пожалуйста."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "обе́дать", "gram": "động từ chưa hoàn thành", "meaning": "ăn trưa", "example": "Мы обедаем в час дня."},
            {"word": "зака́зывать", "gram": "động từ chưa hoàn thành", "meaning": "đặt món, gọi món", "example": "Что вы будете заказывать?"},
            {"word": "про́бовать", "gram": "động từ chưa hoàn thành", "meaning": "nếm thử", "example": "Она пробует традиционный напиток квас."},
            {"word": "пить", "gram": "động từ chưa hoàn thành", "meaning": "uống", "example": "Я пью чёрный чай с лимоном."}
        ]}
    ],
    14: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "друг", "gram": "danh từ giống đực", "meaning": "người bạn", "example": "Денис — мой лучший друг."},
            {"word": "му́зыка", "gram": "danh từ giống cái", "meaning": "âm nhạc", "example": "Он слушает классическую музыку."},
            {"word": "спорт", "gram": "danh từ giống đực", "meaning": "thể thao", "example": "Спорт помогает быть здоровым."},
            {"word": "гита́ра", "gram": "danh từ giống cái", "meaning": "đàn ghi-ta", "example": "Вечером он играет на гитаре."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "дружи́ть", "gram": "động từ chưa hoàn thành", "meaning": "kết bạn, thân thiết", "example": "Мы дружим уже много лет."},
            {"word": "игра́ть", "gram": "động từ chưa hoàn thành", "meaning": "chơi (nhạc cụ / thể thao)", "example": "Студенты играют в футбол на стадионе."},
            {"word": "увлека́ться", "gram": "động từ chưa hoàn thành", "meaning": "đam mê, say mê", "example": "Она увлекается фотографией."},
            {"word": "петь", "gram": "động từ chưa hoàn thành", "meaning": "hát", "example": "Друзья поют русские песни под гитару."}
        ]}
    ],
    15: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "пого́да", "gram": "danh từ giống cái", "meaning": "thời tiết", "example": "Сегодня прекрасная солнечная погода."},
            {"word": "зима́", "gram": "danh từ giống cái", "meaning": "mùa đông", "example": "Зимой в Москве идёт белый снег."},
            {"word": "весна́", "gram": "danh từ giống cái", "meaning": "mùa xuân", "example": "Весной природа оживает и цветёт."},
            {"word": "ле́то", "gram": "danh từ giống trung", "meaning": "mùa hè", "example": "Летом у студентов долгие каникулы."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "свети́ть", "gram": "động từ chưa hoàn thành", "meaning": "chiếu sáng", "example": "На небе ярко светит солнце."},
            {"word": "идти́ (дождь/снег)", "gram": "cụm động từ", "meaning": "rơi (mưa/tuyết)", "example": "На улице идёт тихий снег."},
            {"word": "тая́ть", "gram": "động từ chưa hoàn thành", "meaning": "tan chảy", "example": "Весной снег быстро тает на солнце."},
            {"word": "замерза́ть", "gram": "động từ chưa hoàn thành", "meaning": "đóng băng", "example": "Зимой река полностью замерзает."}
        ]}
    ],
    16: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "вре́мя", "gram": "danh từ giống trung", "meaning": "thời gian", "example": "У меня есть свободное время вечером."},
            {"word": "парк", "gram": "danh từ giống đực", "meaning": "công viên", "example": "Мы гуляем в городском парке."},
            {"word": "фильм", "gram": "danh từ giống đực", "meaning": "bộ phim", "example": "Мы смотрим интересную кинокомедию."},
            {"word": "вы́ставка", "gram": "danh từ giống cái", "meaning": "buổi triển lãm", "example": "В музее открылась новая выставка."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "смотре́ть", "gram": "động từ chưa hoàn thành", "meaning": "xem, ngắm", "example": "Они смотрят новый фильм в кино."},
            {"word": "гуля́ть", "gram": "động từ chưa hoàn thành", "meaning": "đi dạo tản bộ", "example": "Семья гуляет в парке по аллеям."},
            {"word": "фотографи́ровать", "gram": "động từ chưa hoàn thành", "meaning": "chụp ảnh", "example": "Туристы фотографируют старинные здания."},
            {"word": "отдыха́ть", "gram": "động từ chưa hoàn thành", "meaning": "nghỉ ngơi", "example": "После работы приятно отдохнуть дома."}
        ]}
    ],
    17: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "воскресе́нье", "gram": "danh từ giống trung", "meaning": "ngày chủ nhật", "example": "В воскресенье мы идём на прогулку."},
            {"word": "музе́й", "gram": "danh từ giống đực", "meaning": "viện bảo tàng", "example": "В историческом музее много экспонатов."},
            {"word": "теа́тр", "gram": "danh từ giống đực", "meaning": "nhà hát", "example": "Вечером мы идём в Большой театр."},
            {"word": "биле́т", "gram": "danh từ giống đực", "meaning": "vé vào cửa", "example": "Я купил два билета на балет."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "ходи́ть", "gram": "động từ chưa hoàn thành", "meaning": "đi lại thường xuyên", "example": "Каждую неделю мы ходим в театр."},
            {"word": "посеща́ть", "gram": "động từ chưa hoàn thành", "meaning": "thăm quan, dự", "example": "Туристы посещают Третьяковскую галерею."},
            {"word": "встреча́ться", "gram": "động từ chưa hoàn thành", "meaning": "gặp gỡ nhau", "example": "Друзья встречаются около станции метро."},
            {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua", "example": "Мы покупаем билеты в кассе."}
        ]}
    ],
    18: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "письмо́", "gram": "danh từ giống trung", "meaning": "lá thư", "example": "Я получил длинное письмо от друга."},
            {"word": "а́дрес", "gram": "danh từ giống đực", "meaning": "địa chỉ", "example": "На конверте написан московский адрес."},
            {"word": "но́вость", "gram": "danh từ giống cái", "meaning": "tin tức, tin mới", "example": "У меня есть отличная новость для вас."},
            {"word": "отве́т", "gram": "danh từ giống đực", "meaning": "câu trả lời, lời hồi đáp", "example": "Я с нетерпением жду твоего ответа."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "отправля́ть", "gram": "động từ chưa hoàn thành", "meaning": "gửi đi", "example": "Она отправляет письмо по почте."},
            {"word": "получа́ть", "gram": "động từ chưa hoàn thành", "meaning": "nhận được", "example": "Студент получает стипендию каждый месяц."},
            {"word": "отвеча́ть", "gram": "động từ chưa hoàn thành", "meaning": "trả lời, hồi âm", "example": "Я отвечу на твоё письмо завтра."},
            {"word": "расска́зывать", "gram": "động từ chưa hoàn thành", "meaning": "kể chuyện", "example": "В письме он рассказывает о своей учёбе."}
        ]}
    ],
    19: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "пра́здник", "gram": "danh từ giống đực", "meaning": "ngày lễ hội", "example": "Новый год — самый волшебный праздник."},
            {"word": "пода́рок", "gram": "danh từ giống đực", "meaning": "món quà", "example": "Родители приготовили подарки для детей."},
            {"word": "ёлка", "gram": "danh từ giống cái", "meaning": "cây thông noel", "example": "В комнате стоит красивая новогодняя ёлка."},
            {"word": "тради́ция", "gram": "danh từ giống cái", "meaning": "truyền thống", "example": "Семейные традиции объединяют поколения."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "пра́здновать", "gram": "động từ chưa hoàn thành", "meaning": "ăn mừng, tổ chức lễ", "example": "Мы празднуем Новый год тридцать первого декабря."},
            {"word": "поздравля́ть", "gram": "động từ chưa hoàn thành", "meaning": "chúc mừng", "example": "Друзья поздравляют друг друга с праздником."},
            {"word": "дари́ть", "gram": "động từ chưa hoàn thành", "meaning": "tặng quà", "example": "Дед Мороз дарит подарки детям."},
            {"word": "жела́ть", "gram": "động từ chưa hoàn thành", "meaning": "chúc ước", "example": "Мы желаем вам счастья и крепкого здоровья!"}
        ]}
    ],
    20: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "тест", "gram": "danh từ giống đực", "meaning": "bài kiểm tra", "example": "Итоговый тест проверяет чтение и грамматику."},
            {"word": "зна́ние", "gram": "danh từ giống trung", "meaning": "kiến thức", "example": "Знания помогают уверенно общаться."},
            {"word": "у́ровень", "gram": "danh từ giống đực", "meaning": "trình độ cấp độ", "example": "Успешно освоен базовый уровень А1."},
            {"word": "результа́т", "gram": "danh từ giống đực", "meaning": "kết quả", "example": "Студент получил отличный результат."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "проверя́ть", "gram": "động từ chưa hoàn thành", "meaning": "kiểm tra", "example": "Учитель проверяет письменные работы."},
            {"word": "повторя́ть", "gram": "động từ chưa hoàn thành", "meaning": "ôn luyện lại", "example": "Перед экзаменом нужно повторить правила."},
            {"word": "отвеча́ть", "gram": "động từ chưa hoàn thành", "meaning": "trả lời câu hỏi", "example": "Студент без ошибок отвечает на вопросы теста."},
            {"word": "понима́ть", "gram": "động từ chưa hoàn thành", "meaning": "thấu hiểu", "example": "Теперь я понимаю тексты на русском языке."}
        ]}
    ],
    21: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "гости́ница", "gram": "danh từ giống cái", "meaning": "khách sạn", "example": "Мы остановились в трёхзвёздочной гостинице."},
            {"word": "но́мер", "gram": "danh từ giống đực", "meaning": "phòng khách sạn", "example": "Наш номер чистый, с видом на реку."},
            {"word": "ключ", "gram": "danh từ giống đực", "meaning": "chìa khóa", "example": "Администратор дал электронный ключ от номера."},
            {"word": "па́спорт", "gram": "danh từ giống đực", "meaning": "hộ chiếu", "example": "Для регистрации покажите ваш паспорт."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "брони́ровать", "gram": "động từ chưa hoàn thành", "meaning": "đặt chỗ trước", "example": "Я бронирую номер через интернет."},
            {"word": "жить", "gram": "động từ chưa hoàn thành", "meaning": "cư trú, ở", "example": "Туристы живут в гостинице пять дней."},
            {"word": "плати́ть", "gram": "động từ chưa hoàn thành", "meaning": "trả tiền", "example": "Мы платим за номер при выезде."},
            {"word": "выезжа́ть", "gram": "động từ chưa hoàn thành", "meaning": "trả phòng, rời đi", "example": "Гости выезжают до двенадцати часов дня."}
        ]}
    ],
    22: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "по́чта", "gram": "danh từ giống cái", "meaning": "bưu điện", "example": "Почта находится на соседней улице."},
            {"word": "откры́тка", "gram": "danh từ giống cái", "meaning": "bưu thiếp", "example": "Анна отправляет открытку бабушке."},
            {"word": "посы́лка", "gram": "danh từ giống cái", "meaning": "kiện hàng bưu phẩm", "example": "Родители прислали тёплую посылку."},
            {"word": "ма́рка", "gram": "danh từ giống cái", "meaning": "con tem bưu chính", "example": "Купите две красивые марки на конверт."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "отправля́ть", "gram": "động từ chưa hoàn thành", "meaning": "gửi đi", "example": "Он отправляет посылку в другой город."},
            {"word": "получа́ть", "gram": "động từ chưa hoàn thành", "meaning": "nhận về", "example": "Я получаю заказное письмо по паспорту."},
            {"word": "подпи́сывать", "gram": "động từ chưa hoàn thành", "meaning": "ký tên, ghi địa chỉ", "example": "Она подписывает открытку с праздником."},
            {"word": "взве́шивать", "gram": "động từ chưa hoàn thành", "meaning": "cân trọng lượng", "example": "Оператор взвешивает тяжёлую коробку."}
        ]}
    ],
    23: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "кинотеа́тр", "gram": "danh từ giống đực", "meaning": "rạp chiếu phim", "example": "В субботу мы идём в современный кинотеатр."},
            {"word": "сеа́нс", "gram": "danh từ giống đực", "meaning": "suất chiếu", "example": "Вечерний сеанс начинается в девятнадцать часов."},
            {"word": "экран", "gram": "danh từ giống đực", "meaning": "màn ảnh", "example": "На большом экране показывают премьеру."},
            {"word": "ряд", "gram": "danh từ giống đực", "meaning": "hàng ghế", "example": "Наши места в пятом ряду в центре."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "смотре́ть", "gram": "động từ chưa hoàn thành", "meaning": "xem phim", "example": "Друзья смотрят фильм с субтитрами."},
            {"word": "выбира́ть", "gram": "động từ chưa hoàn thành", "meaning": "lựa chọn", "example": "Мы выбираем лучшие места в зале."},
            {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua vé", "example": "Антон покупает билеты и попкорн."},
            {"word": "обсужда́ть", "gram": "động từ chưa hoàn thành", "meaning": "thảo luận", "example": "После фильма мы обсуждаем сюжет в кафе."}
        ]}
    ],
    24: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "дру́жба", "gram": "danh từ giống cái", "meaning": "tình bạn bè", "example": "Крепкая дружба проверяется временем."},
            {"word": "встре́ча", "gram": "danh từ giống cái", "meaning": "buổi gặp mặt", "example": "Долгожданная встреча старых друзей."},
            {"word": "разгово́р", "gram": "danh từ giống đực", "meaning": "cuộc trò chuyện", "example": "Интересный разговор за чашкой чая."},
            {"word": "по́мощь", "gram": "danh từ giống cái", "meaning": "sự giúp đỡ", "example": "Друг всегда готов прийти на помощь."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "обща́ться", "gram": "động từ chưa hoàn thành", "meaning": "giao tiếp, trò chuyện", "example": "Мы часто общаемся в мессенджере."},
            {"word": "дели́ться", "gram": "động từ chưa hoàn thành", "meaning": "chia sẻ", "example": "Друзья делятся радостными новостями."},
            {"word": "подде́рживать", "gram": "động từ chưa hoàn thành", "meaning": "ủng hộ, động viên", "example": "Он всегда поддерживает меня в трудную минуту."},
            {"word": "по́мнить", "gram": "động từ chưa hoàn thành", "meaning": "ghi nhớ", "example": "Я помню нашу первую встречу в университете."}
        ]}
    ],
    25: [
        {"title": "Danh từ trong bài đọc (Существительные)", "rows": [
            {"word": "успе́х", "gram": "danh từ giống đực", "meaning": "thành công, tiến bộ", "example": "Большой успех в изучении русского языка."},
            {"word": "сертифика́т", "gram": "danh từ giống đực", "meaning": "chứng chỉ", "example": "Студент получил сертификат уровня А1."},
            {"word": "цель", "gram": "danh từ giống cái", "meaning": "mục tiêu phấn đấu", "example": "Моя новая цель — сдать экзамен А2."},
            {"word": "горизо́нт", "gram": "danh từ giống đực", "meaning": "chân trời mới", "example": "Перед нами открываются новые горизонты."}
        ]},
        {"title": "Động từ then chốt (Глаголы)", "rows": [
            {"word": "заверша́ть", "gram": "động từ chưa hoàn thành", "meaning": "hoàn thành", "example": "Мы завершаем курс начального уровня."},
            {"word": "достига́ть", "gram": "động từ chưa hoàn thành", "meaning": "đạt được", "example": "Трудолюбивый человек достигает своей цели."},
            {"word": "продолжа́ть", "gram": "động từ chưa hoàn thành", "meaning": "tiếp tục", "example": "Я буду продолжать учить русский язык."},
            {"word": "совершенствова́ть", "gram": "động từ chưa hoàn thành", "meaning": "hoàn thiện, nâng cao", "example": "Мы совершенствуем свои разговорные навыки."}
        ]}
    ]
}

with open('scripts/reading_a1_vocab.json', 'w', encoding='utf-8') as f:
    json.dump(reading_a1, f, ensure_ascii=False, indent=2)
print("Saved reading_a1_vocab.json (25 lessons, 200 words)")
