#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Expand A1 to full 25 lessons for each skill (Reading, Listening, Speaking, Writing).
Adds lessons 21-25 to existing 20 lessons.
"""
import json
import os

BASE_DIR = "public/data"

# ==============================================================================
# A1 READING (Lessons 21 - 25)
# ==============================================================================
a1_reading_extra = [
    {
        "id": 21,
        "level": "A1",
        "skill": "reading",
        "title": "В гостинице: Регистрация и проживание",
        "badge": "A1 Đọc hiểu",
        "readingTime": "12 phút",
        "source": "Giáo trình tiếng Nga A1 Chuẩn Quốc Tế (TRKI)",
        "topic": "Khách sạn, Nhận phòng & Tiện nghi phòng",
        "vocabCategories": [
            {
                "title": "Từ vựng Khách sạn (Гостиница)",
                "rows": [
                    {"word": "гости́ница", "gram": "danh từ giống cái", "meaning": "khách sạn", "example": "Мы живём в гостинице «Космос»."},
                    {"word": "но́мер", "gram": "danh từ giống đực", "meaning": "phòng khách sạn", "example": "Наш номер находится на пятом этаже."},
                    {"word": "ключ", "gram": "danh từ giống đực", "meaning": "chìa khóa phòng", "example": "Вот ваш электронный ключ от номера."},
                    {"word": "па́спорт", "gram": "danh từ giống đực", "meaning": "hộ chiếu", "example": "Покажите, пожалуйста, ваш паспорт."}
                ]
            }
        ],
        "readingText": [
            "Здравствуйте! Меня зовут Дэвид. Я приехал в Москву как турист.",
            "Сейчас я нахожусь в гостинице. На ресепшене работает очень вежливый администратор.",
            "— Здравствуйте! Я бронировал одноместный номер на три дня. Вот мой паспорт.\n— Добрый день, господин Дэвид! Да, ваша бронь подтверждена. Ваш номер — 405 на четвёртом этаже.",
            "В номере есть большая кровать, шкаф, стол, бесплатный Wi-Fi и телевизор. Завтрак включён в стоимость и проходит с семи до десяти утра на первом этаже."
        ],
        "questions": [
            {
                "id": "a1_r21_q1",
                "lang": "ru",
                "prompt": "Какой номер комнаты получил Дэвид?",
                "hint": "Xem số phòng ở câu thoại của nhân viên lễ tân.",
                "options": [
                    {"key": "A", "text": "Номер 100"},
                    {"key": "B", "text": "Номер 405"},
                    {"key": "C", "text": "Номер 504"},
                    {"key": "D", "text": "Номер 301"}
                ],
                "ans": "B",
                "explain": "Lễ tân nói rõ: «Ваш номер — 405 на четвёртом этаже»."
            },
            {
                "id": "a1_r21_q2",
                "lang": "ru",
                "prompt": "Во сколько начинается завтрак в гостинице?",
                "hint": "Tìm thời gian phục vụ bữa sáng ở đoạn cuối.",
                "options": [
                    {"key": "A", "text": "В шесть утра"},
                    {"key": "B", "text": "В семь утра"},
                    {"key": "C", "text": "В восемь утра"},
                    {"key": "D", "text": "В одиннадцать утра"}
                ],
                "ans": "B",
                "explain": "Đoạn cuối viết: «Завтрак включён в стоимость и проходит с семи до десяти утра...»."
            }
        ]
    },
    {
        "id": 22,
        "level": "A1",
        "skill": "reading",
        "title": "На почте: Открытки и посылки",
        "badge": "A1 Đọc hiểu",
        "readingTime": "12 phút",
        "source": "Giáo trình tiếng Nga A1 Chuẩn Quốc Tế (TRKI)",
        "topic": "Bưu điện: Gửi bưu thiếp, thư từ & bưu phẩm",
        "vocabCategories": [
            {
                "title": "Từ vựng Bưu điện (Почта)",
                "rows": [
                    {"word": "по́чта", "gram": "danh từ giống cái", "meaning": "bưu điện", "example": "Почта открывается в восемь утра."},
                    {"word": "писько́", "gram": "danh từ giống trung", "meaning": "lá thư", "example": "Я пишу письмо родителям."},
                    {"word": "откры́тка", "gram": "danh từ giống cái", "meaning": "bưu thiếp", "example": "Купите красивые открытки с видами Москвы."},
                    {"word": "ма́рка", "gram": "danh từ giống cái", "meaning": "con tem", "example": "Наклейте марку на конверт."}
                ]
            }
        ],
        "readingText": [
            "Сегодня Анна идёт на почту. Почта находится рядом с её домом.",
            "Анна хочет отправить красивую открытку своей подруге во Вьетнам. На открытке изображён Зимний дворец в Санкт-Петербурге.",
            "— Здравствуйте! Дайте, пожалуйста, две открытки и три марки для отправки за границу.\n— Пожалуйста! С вас сто пятьдесят рублей.\n— Вот деньги. Скажите, где здесь почтовый ящик?\n— Синий почтовый ящик висит прямо на улице у входа.",
            "Анна аккуратно пишет адрес и опускает открытку в ящик."
        ],
        "questions": [
            {
                "id": "a1_r22_q1",
                "lang": "ru",
                "prompt": "Куда Анна отправляет открытку?",
                "hint": "Tìm tên đất nước nơi bạn của Anna sống.",
                "options": [
                    {"key": "A", "text": "В Россию"},
                    {"key": "B", "text": "Во Вьетнам"},
                    {"key": "C", "text": "Во Францию"},
                    {"key": "D", "text": "В Китай"}
                ],
                "ans": "B",
                "explain": "Đoạn 2 nêu: «отправить красивую открытку своей подруге во Вьетнам»."
            }
        ]
    },
    {
        "id": 23,
        "level": "A1",
        "skill": "reading",
        "title": "В кинотеатре: Мой любимый фильм",
        "badge": "A1 Đọc hiểu",
        "readingTime": "12 phút",
        "source": "Giáo trình tiếng Nga A1 Chuẩn Quốc Tế (TRKI)",
        "topic": "Giải trí: Đi xem phim và rạp chiếu phim",
        "vocabCategories": [
            {
                "title": "Từ vựng Điện ảnh (Кино)",
                "rows": [
                    {"word": "кинотеа́тр", "gram": "danh từ giống đực", "meaning": "rạp chiếu phim", "example": "Мы идём в кинотеатр в субботу."},
                    {"word": "биле́т", "gram": "danh từ giống đực", "meaning": "vé", "example": "Сколько стоит билет в кино?"},
                    {"word": "ряд", "gram": "danh từ giống đực", "meaning": "hàng ghế", "example": "Наши места в пятом ряду."},
                    {"word": "попко́рн", "gram": "danh từ giống đực", "meaning": "bỏng ngô", "example": "Перед фильмом мы купили сладкий попкорн."}
                ]
            }
        ],
        "readingText": [
            "В субботу вечером студенты Антон и Мария решили пойти в кинотеатр «Октябрь».",
            "Они выбрали новую российскую комедию. Сеанс начинается в 18:30.",
            "— Здравствуйте! Дайте, пожалуйста, два билета на комедию на полседьмого.\n— Какие места вы предпочитаете?\n— Шестой ряд, места десять и одиннадцать, пожалуйста.\n— С вас шестьсот рублей.",
            "Фильм был очень весёлый и интересный. После кино Антон и Мария пошли в кафе обсуждать сюжет."
        ],
        "questions": [
            {
                "id": "a1_r23_q1",
                "lang": "ru",
                "prompt": "В какой день недели Антон и Мария пошли в кино?",
                "hint": "Xem thứ trong tuần ở câu đầu tiên.",
                "options": [
                    {"key": "A", "text": "В понедельник"},
                    {"key": "B", "text": "В субботу"},
                    {"key": "C", "text": "В среду"},
                    {"key": "D", "text": "В пятницу"}
                ],
                "ans": "B",
                "explain": "Câu mở đầu nêu: «В субботу вечером...»."
            }
        ]
    },
    {
        "id": 24,
        "level": "A1",
        "skill": "reading",
        "title": "Дружба и совместные увлечения",
        "badge": "A1 Đọc hiểu",
        "readingTime": "12 phút",
        "source": "Giáo trình tiếng Nga A1 Chuẩn Quốc Tế (TRKI)",
        "topic": "Tình bạn & Các hoạt động giải trí yêu thích",
        "vocabCategories": [
            {
                "title": "Tình bạn & Sở thích (Дружба и хобби)",
                "rows": [
                    {"word": "друг / подру́га", "gram": "danh từ", "meaning": "bạn trai / bạn gái", "example": "Максим — мой лучший друг."},
                    {"word": "вме́сте", "gram": "phó từ", "meaning": "cùng nhau", "example": "Мы часто учимся и гуляем вместе."},
                    {"word": "спорт", "gram": "danh từ giống đực", "meaning": "thể thao", "example": "Спорт помогает быть здоровым."},
                    {"word": "му́зыка", "gram": "danh từ giống cái", "meaning": "âm nhạc", "example": "Я люблю слушать классическую музыку."}
                ]
            }
        ],
        "readingText": [
            "У меня есть хороший друг. Его зовут Максим. Мы познакомились в университете на первом курсе.",
            "Максим учится на экономическом факультете, а я изучаю филологию. Несмотря на разные специальности, у нас много общих увлечений.",
            "По выходным мы играем в футбол на стадионе или катаемся на велосипедах в парке. А по вечерам мы любим пить чай, слушать музыку и говорить обо всём на свете.",
            "Хороший друг — это настоящее сокровище в жизни каждого человека."
        ],
        "questions": [
            {
                "id": "a1_r24_q1",
                "lang": "ru",
                "prompt": "Где познакомились друзья?",
                "hint": "Xem địa điểm quen nhau ở đoạn 1.",
                "options": [
                    {"key": "A", "text": "В школе"},
                    {"key": "B", "text": "В университете"},
                    {"key": "C", "text": "На стадионе"},
                    {"key": "D", "text": "В кафе"}
                ],
                "ans": "B",
                "explain": "Đoạn 1 viết: «Мы познакомились в университете на первом курсе»."
            }
        ]
    },
    {
        "id": 25,
        "level": "A1",
        "skill": "reading",
        "title": "Итоги A1: Я готов говорить и учиться дальше!",
        "badge": "A1 Đọc hiểu",
        "readingTime": "15 phút",
        "source": "Giáo trình tiếng Nga A1 Chuẩn Quốc Tế (TRKI)",
        "topic": "Tổng kết A1: Đánh giá thành quả và bước tiếp lên A2",
        "vocabCategories": [
            {
                "title": "Tổng kết & Thành tựu (Итоги и достижения)",
                "rows": [
                    {"word": "результа́т", "gram": "danh từ giống đực", "meaning": "kết quả", "example": "У меня отличный результат на тесте."},
                    {"word": "зна́ния", "gram": "danh từ số nhiều", "meaning": "kiến thức", "example": "Я получил хорошие базовые знания."},
                    {"word": "прогре́сс", "gram": "danh từ giống đực", "meaning": "sự tiến bộ", "example": "Преподаватель похвалил мой прогресс."},
                    {"word": "уве́ренно", "gram": "phó từ", "meaning": "tự tin", "example": "Теперь я могу уверенно говорить по-русски."}
                ]
            }
        ],
        "readingText": [
            "Поздравляем! Вы успешно завершили двадцать пять уроков чтения на уровне A1.",
            "Теперь вы знаете русский алфавит, основные правила чтения, умеете здороваться, рассказывать о себе, своей семье, профессии, покупках, времени и погоде.",
            "Вы освоили более пятисот базовых слов và các cấu trúc câu giao tiếp thường nhật. Впереди вас ждёт уровень A2 (Базовый уровень), где вы научитесь свободно ориентироваться в городе, заказывать билеты и общаться с врачом.",
            "Учите русский язык с радостью вместе с TAIGA TALK!"
        ],
        "questions": [
            {
                "id": "a1_r25_q1",
                "lang": "ru",
                "prompt": "Какой уровень ждёт студентов впереди после A1?",
                "hint": "Tìm tên trình độ tiếp theo ở đoạn 3.",
                "options": [
                    {"key": "A", "text": "Уровень A0"},
                    {"key": "B", "text": "Уровень A2 (Базовый уровень)"},
                    {"key": "C", "text": "Уровень C2"},
                    {"key": "D", "text": "Сразу аспирантура"}
                ],
                "ans": "B",
                "explain": "Đoạn 3 nêu: «Впереди вас ждёт уровень A2 (Базовый уровень)...»."
            }
        ]
    }
]

# Update Reading A1
for l in a1_reading_extra:
    filename = f"lesson-{l['id']:02d}.json"
    with open(f"{BASE_DIR}/reading/a1/{filename}", "w", encoding="utf-8") as f:
        json.dump(l, f, ensure_ascii=False, indent=2)

# Update manifest
with open(f"{BASE_DIR}/reading/a1/manifest.json", "r", encoding="utf-8") as f:
    manifest_r = json.load(f)
existing_ids = {m["id"] for m in manifest_r}
for l in a1_reading_extra:
    if l["id"] not in existing_ids:
        manifest_r.append({
            "id": l["id"],
            "file": f"lesson-{l['id']:02d}.json",
            "title": l["title"],
            "badge": l["badge"],
            "readingTime": l["readingTime"],
            "source": l["source"],
            "topic": l["topic"],
            "questionCount": len(l["questions"])
        })
with open(f"{BASE_DIR}/reading/a1/manifest.json", "w", encoding="utf-8") as f:
    json.dump(manifest_r, f, ensure_ascii=False, indent=2)

print("A1 Reading expanded to 25 lessons.")

# ==============================================================================
# A1 LISTENING (Lessons 21 - 25)
# ==============================================================================
a1_listening_extra = [
    {
        "id": 21,
        "title": "На вокзале: Какой поезд едет в Петербург?",
        "listeningTime": "10 phút",
        "topic": "Hội thoại: Hỏi thông tin tại nhà ga xe lửa",
        "dialogue": [
            {"speaker": "Пассажир", "text": "Скажите, пожалуйста, с какой платформы отправляется скорый поезд в Санкт-Петербург?"},
            {"speaker": "Дежурный", "text": "Поезд номер семьсот два отправляется со второй платформы в шестнадцать тридцать."},
            {"speaker": "Пассажир", "text": "Спасибо! А выход на вторую платформу через подземный переход?"},
            {"speaker": "Дежурный", "text": "Да, поверните направо и спуститесь вниз по лестнице."},
            {"speaker": "Пассажир", "text": "Большое спасибо за информацию!"}
        ],
        "dictation": {
            "title": "Chép chính tả (Диктант)",
            "instruction": "Nghe kỹ và gõ lại câu sau:",
            "targetText": "Поезд отправляется со второй платформы в четыре часа.",
            "hint": "Chú ý viết hoa chữ đầu và dùng giới từ «со второй платформы»."
        },
        "questions": [
            {
                "id": "a1_l21_q1",
                "lang": "ru",
                "prompt": "С какой платформы отправляется поезд?",
                "hint": "Xem số hiệu sân ga ở câu trả lời của nhân viên.",
                "options": [
                    {"key": "A", "text": "С первой платформы"},
                    {"key": "B", "text": "Со второй платформы"},
                    {"key": "C", "text": "С пятой платформы"},
                    {"key": "D", "text": "С десятой платформы"}
                ],
                "ans": "B",
                "explain": "Nhân viên trả lời: «отправляется со второй платформы»."
            }
        ]
    },
    {
        "id": 22,
        "title": "В аптеке: Витамины и чай с лимоном",
        "listeningTime": "10 phút",
        "topic": "Hội thoại: Mua vitamin và trà thảo mộc tại quầy thuốc",
        "dialogue": [
            {"speaker": "Покупатель", "text": "Здравствуйте! У вас есть витамин С и травяной чай от усталости?"},
            {"speaker": "Фармацевт", "text": "Да, конечно. Вот растворимый витамин С со вкусом апельсина и натуральный чай."},
            {"speaker": "Покупатель", "text": "Сколько это стоит?"},
            {"speaker": "Фармацевт", "text": "Витамины стоят двести рублей, а чай — сто двадцать рублей. Всего триста двадцать рублей."},
            {"speaker": "Покупатель", "text": "Можно оплатить картой?\n— Да, приложите карту к терминалу. Спасибо за покупку!"}
        ],
        "dictation": {
            "title": "Chép chính tả (Диктант)",
            "instruction": "Nghe kỹ và gõ lại câu sau:",
            "targetText": "Дайте, пожалуйста, витамины и травяной чай.",
            "hint": "Dấu phẩy sau «пожалуйста»."
        },
        "questions": [
            {
                "id": "a1_l22_q1",
                "lang": "ru",
                "prompt": "Сколько всего стоят покупки?",
                "hint": "Tính tổng tiền ở lượt thoại thứ 4.",
                "options": [
                    {"key": "A", "text": "100 рублей"},
                    {"key": "B", "text": "320 рублей"},
                    {"key": "C", "text": "500 рублей"},
                    {"key": "D", "text": "200 рублей"}
                ],
                "ans": "B",
                "explain": "Dược sĩ tính: 200 + 120 = 320 rúp."
            }
        ]
    },
    {
        "id": 23,
        "title": "В музее: Покупка студенческого билета",
        "listeningTime": "10 phút",
        "topic": "Hội thoại: Mua vé tham quan bảo tàng cho sinh viên",
        "dialogue": [
            {"speaker": "Студент", "text": "Добрый день! Скажите, есть ли у вас скидки для иностранных студентов?"},
            {"speaker": "Кассир", "text": "Да, покажите ваш студенческий билет. Для студентов билет стоит двести рублей вместо пятисот."},
            {"speaker": "Студент", "text": "Вот мой студенческий билет. Дайте, пожалуйста, один билет."},
            {"speaker": "Кассир", "text": "Возьмите билет и план залов музея. Аудиогид можно взять на втором этаже."}
        ],
        "dictation": {
            "title": "Chép chính tả (Диктант)",
            "instruction": "Nghe kỹ và gõ lại câu sau:",
            "targetText": "Для студентов билет стоит двести рублей.",
            "hint": "Chú ý giới từ «Для» và số đếm «двести»."
        },
        "questions": [
            {
                "id": "a1_l23_q1",
                "lang": "ru",
                "prompt": "Сколько стоит билет для студентов?",
                "hint": "Xem giá vé ưu đãi của sinh viên.",
                "options": [
                    {"key": "A", "text": "500 рублей"},
                    {"key": "B", "text": "200 рублей"},
                    {"key": "C", "text": "Бесплатно"},
                    {"key": "D", "text": "1000 рублей"}
                ],
                "ans": "B",
                "explain": "Thu ngân báo giá: «Для студентов билет стоит двести рублей»."
            }
        ]
    },
    {
        "id": 24,
        "title": "В парке: Прокат велосипедов",
        "listeningTime": "10 phút",
        "topic": "Hội thoại: Thuê xe đạp dạo chơi công viên",
        "dialogue": [
            {"speaker": "Клиент", "text": "Здравствуйте! Мы хотим взять два велосипеда на один час."},
            {"speaker": "Сотрудник", "text": "Отлично! Прокат одного велосипеда стоит двести пятьдесят рублей в час. Нужен документ для залога."},
            {"speaker": "Клиент", "text": "Вот моё водительское удостоверение. Мы вернём велосипеды в пять часов вечера."},
            {"speaker": "Сотрудник", "text": "Хорошо! Вот ваши ключи и шлемы. Приятной прогулки!"}
        ],
        "dictation": {
            "title": "Chép chính tả (Диктант)",
            "instruction": "Nghe kỹ và gõ lại câu sau:",
            "targetText": "Мы хотим взять два велосипеда на один час.",
            "hint": "Động từ «взять» và cụm thời gian «на один час»."
        },
        "questions": [
            {
                "id": "a1_l24_q1",
                "lang": "ru",
                "prompt": "На сколько часов клиенты берут велосипеды?",
                "hint": "Xem thời gian thuê ở câu đầu tiên.",
                "options": [
                    {"key": "A", "text": "На один час"},
                    {"key": "B", "text": "На целый день"},
                    {"key": "C", "text": "На пять часов"},
                    {"key": "D", "text": "На выходные"}
                ],
                "ans": "A",
                "explain": "Khách nói: «Мы хотим взять два велосипеда на один час»."
            }
        ]
    },
    {
        "id": 25,
        "title": "Итоговое аудирование A1: Приглашение на праздник",
        "listeningTime": "12 phút",
        "topic": "Hội thoại tổng kết: Nghe lời mời dự tiệc và chuẩn bị quà",
        "dialogue": [
            {"speaker": "Ольга", "text": "Привет, Виктор! В субботу у меня день рождения. Я приглашаю тебя в гости в шесть вечера."},
            {"speaker": "Виктор", "text": "Ольга, спасибо огромное за приглашение! С удовольствием приду. Что тебе подарить?"},
            {"speaker": "Ольга", "text": "Главный подарок — твоё присутствие! Но если хочешь, принеси торт или цветы."},
            {"speaker": "Виктор", "text": "Договорились! Я куплю вкусный шоколадный торт. До субботы!"}
        ],
        "dictation": {
            "title": "Chép chính tả (Диктант)",
            "instruction": "Nghe kỹ và gõ lại câu sau:",
            "targetText": "Я приглашаю тебя в гости в субботу вечером.",
            "hint": "Động từ «приглашаю» và trạng từ «вечером»."
        },
        "questions": [
            {
                "id": "a1_l25_q1",
                "lang": "ru",
                "prompt": "Что Виктор пообещал принести на день рождения?",
                "hint": "Tìm món quà Viktor hứa đem tới.",
                "options": [
                    {"key": "A", "text": "Книгу"},
                    {"key": "B", "text": "Шоколадный торт"},
                    {"key": "C", "text": "Пиццу"},
                    {"key": "D", "text": "Фрукты"}
                ],
                "ans": "B",
                "explain": "Viktor nói: «Я куплю вкусный шоколадный торт»."
            }
        ]
    }
]

# Write Listening A1 extra
for l in a1_listening_extra:
    filename = f"lesson-{l['id']:02d}.json"
    with open(f"{BASE_DIR}/listening/a1/{filename}", "w", encoding="utf-8") as f:
        json.dump(l, f, ensure_ascii=False, indent=2)

with open(f"{BASE_DIR}/listening/a1/manifest.json", "r", encoding="utf-8") as f:
    manifest_l = json.load(f)
existing_ids = {m["id"] for m in manifest_l}
for l in a1_listening_extra:
    if l["id"] not in existing_ids:
        manifest_l.append({
            "id": l["id"],
            "file": f"lesson-{l['id']:02d}.json",
            "title": l["title"],
            "badge": "A1 Nghe hiểu",
            "listeningTime": l["listeningTime"],
            "topic": l["topic"],
            "questionCount": len(l["questions"])
        })
with open(f"{BASE_DIR}/listening/a1/manifest.json", "w", encoding="utf-8") as f:
    json.dump(manifest_l, f, ensure_ascii=False, indent=2)

print("A1 Listening expanded to 25 lessons.")

# ==============================================================================
# A1 SPEAKING (Lessons 21 - 25)
# ==============================================================================
a1_speaking_extra = [
    {
        "id": 21,
        "level": "A1",
        "skill": "speaking",
        "title": "В кафе: Заказ любимых блюд",
        "speakingTime": "15 phút",
        "topic": "Luyện nói: Gọi đồ ăn, đồ uống và thanh toán",
        "intro": "Luyện tập các cấu trúc gọi món ăn cơ bản và giao tiếp lịch sự với người phục vụ.",
        "shadowingPhrases": [
            {"id": "spk_21_1", "text": "Дайте, пожалуйста, чашку чёрного кофе с сахаром.", "meaning": "Cho tôi một tách cà phê đen có đường.", "phonetic": "Dayte, pozhaluysta, chashku chornogo kofe s sakharom.", "tips": "Âm «чёрного» phát âm /chornava/."},
            {"id": "spk_21_2", "text": "Сколько стоит этот яблочный пирог?", "meaning": "Bánh táo này giá bao nhiêu?", "phonetic": "Skol'ko stoit etot yablochnyy pirog?", "tips": "Trọng âm từ «яблочный» rơi vào âm tiết đầu."}
        ],
        "freeSpeakingPrompt": {
            "title": "Nói tự do",
            "task": "Đóng vai khách hàng gọi món tại quán ăn nhanh kiểu Nga.",
            "sampleGuidance": ["Здравствуйте!", "Я буду...", "Сколько с меня?", "Спасибо!"]
        }
    },
    {
        "id": 22,
        "level": "A1",
        "skill": "speaking",
        "title": "Спрашиваем дорогу на улице",
        "speakingTime": "15 phút",
        "topic": "Hỏi đường đi đến ga tàu điện ngầm và trạm xe buýt",
        "intro": "Luyện mẫu câu hỏi phương hướng và chỉ đường đơn giản.",
        "shadowingPhrases": [
            {"id": "spk_22_1", "text": "Скажите, пожалуйста, где здесь метро?", "meaning": "Xin cho hỏi ở đây ga metro ở đâu?", "phonetic": "Skazhite, pozhaluysta, gde zdes' metro?", "tips": "«Скажите» phát âm mềm mại."},
            {"id": "spk_22_2", "text": "Идите прямо и поверните направо.", "meaning": "Hãy đi thẳng rồi rẽ phải.", "phonetic": "Idite pryamo i povernite napravo.", "tips": "«Прямо» nhấn âm đầu."}
        ],
        "freeSpeakingPrompt": {
            "title": "Nói tự do",
            "task": "Hướng dẫn một người bạn đi từ cổng trường đến trạm xe buýt gần nhất.",
            "sampleGuidance": ["Метро находится...", "Идите прямо...", "Потом поверните..."]
        }
    },
    {
        "id": 23,
        "level": "A1",
        "skill": "speaking",
        "title": "Поздравления и пожелания",
        "speakingTime": "15 phút",
        "topic": "Chúc mừng ngày lễ và sinh nhật bằng tiếng Nga",
        "intro": "Luyện các câu chúc mừng quen thuộc trong văn hóa Nga.",
        "shadowingPhrases": [
            {"id": "spk_23_1", "text": "Поздравляю с днём рождения! Желаю счастья!", "meaning": "Chúc mừng sinh nhật! Chúc bạn hạnh phúc!", "phonetic": "Pozdravlyayu s dnyom rozhdeniya! Zhelayu schast'ya!", "tips": "«Счастья» đọc âm /shchastya/."},
            {"id": "spk_23_2", "text": "С Новым годом и Рождеством!", "meaning": "Chúc mừng Năm Mới và Giáng Sinh!", "phonetic": "S Novym godom i Rozhdestvom!", "tips": "Chúc mừng Năm mới dùng Cách 5."}
        ],
        "freeSpeakingPrompt": {
            "title": "Nói tự do",
            "task": "Nói 3–4 câu chúc mừng sinh nhật gửi tới một người bạn thân.",
            "sampleGuidance": ["Дорогой друг!", "Поздравляю тебя...", "Желаю тебе здоровья...", "Всего самого доброго!"]
        }
    },
    {
        "id": 24,
        "level": "A1",
        "skill": "speaking",
        "title": "Мои планы на выходные",
        "speakingTime": "15 phút",
        "topic": "Kể về dự định làm gì vào thứ Bảy và Chủ nhật",
        "intro": "Luyện diễn đạt kế hoạch cuối tuần với động từ nguyên mẫu.",
        "shadowingPhrases": [
            {"id": "spk_24_1", "text": "В субботу я хочу пойти в парк с друзьями.", "meaning": "Thứ Bảy tôi muốn đi công viên với bạn bè.", "phonetic": "V subbotu ya khochu poyti v park s druz'yami.", "tips": "«В субботу» dùng Cách 4."},
            {"id": "spk_24_2", "text": "В воскресенье я буду отдыхать дома.", "meaning": "Chủ nhật tôi sẽ nghỉ ngơi ở nhà.", "phonetic": "V voskresen'ye ya budu otdykhat' doma.", "tips": "«Буду отдыхать» thể hiện tương lai chưa hoàn thành."}
        ],
        "freeSpeakingPrompt": {
            "title": "Nói tự do",
            "task": "Kể về kế hoạch cuối tuần này của bạn.",
            "sampleGuidance": ["В субботу утром я...", "Днём мы будем...", "В воскресенье вечером я..."]
        }
    },
    {
        "id": 25,
        "level": "A1",
        "skill": "speaking",
        "title": "Итоговое говорение A1: Я говорю по-русски!",
        "speakingTime": "15 phút",
        "topic": "Tự giới thiệu bản thân toàn diện và chia sẻ cảm nghĩ học tiếng Nga",
        "intro": "Tổng kết kỹ năng nói A1: Tự tin nói trôi chảy một đoạn văn hoàn chỉnh.",
        "shadowingPhrases": [
            {"id": "spk_25_1", "text": "Я успешно закончил первый уровень изучения русского языка.", "meaning": "Tôi đã hoàn thành xuất sắc trình độ đầu tiên học tiếng Nga.", "phonetic": "Ya uspeshno zakonchil pervyy uroven' izucheniya russkogo yazyka.", "tips": "«Успешно закончил» phát âm rành rọt."},
            {"id": "spk_25_2", "text": "Я люблю русский язык и буду учиться дальше!", "meaning": "Tôi yêu tiếng Nga và sẽ tiếp tục học lên cao!", "phonetic": "Ya lyublyu russkiy yazyk i budu uchit'sya dal'she!", "tips": "«Дальше» âm /l'/ mềm."}
        ],
        "freeSpeakingPrompt": {
            "title": "Nói tự do",
            "task": "Tự giới thiệu bản thân 5–6 câu tổng kết những gì bạn đã học ở A1.",
            "sampleGuidance": ["Меня зовут...", "Я изучаю русский язык...", "Я умею...", "Моя цель — уровень A2!"]
        }
    }
]

for l in a1_speaking_extra:
    filename = f"lesson-{l['id']:02d}.json"
    with open(f"{BASE_DIR}/speaking/a1/{filename}", "w", encoding="utf-8") as f:
        json.dump(l, f, ensure_ascii=False, indent=2)

with open(f"{BASE_DIR}/speaking/a1/manifest.json", "r", encoding="utf-8") as f:
    manifest_s = json.load(f)
existing_ids = {m["id"] for m in manifest_s}
for l in a1_speaking_extra:
    if l["id"] not in existing_ids:
        manifest_s.append({
            "id": l["id"],
            "file": f"lesson-{l['id']:02d}.json",
            "title": l["title"],
            "badge": "A1 Nói",
            "speakingTime": l["speakingTime"],
            "topic": l["topic"],
            "questionCount": len(l["shadowingPhrases"])
        })
with open(f"{BASE_DIR}/speaking/a1/manifest.json", "w", encoding="utf-8") as f:
    json.dump(manifest_s, f, ensure_ascii=False, indent=2)

print("A1 Speaking expanded to 25 lessons.")

# ==============================================================================
# A1 WRITING (Lessons 21 - 25)
# ==============================================================================
a1_writing_extra = [
    {
        "id": 21,
        "level": "A1",
        "skill": "writing",
        "title": "Записка соседу по комнате",
        "writingTime": "20 phút",
        "topic": "Viết mẩu giấy nhắn ngắn cho bạn cùng phòng ký túc xá",
        "targetWordRange": [30, 50],
        "task": "Viết một mẩu giấy nhắn (30–50 từ) để lại trên bàn cho bạn cùng phòng báo rằng bạn đã đi siêu thị và nhờ bạn ấy nấu cơm.",
        "requirements": [
            "1. Lời xưng hô tên bạn cùng phòng.",
            "2. Báo bạn đi đâu và khi nào về.",
            "3. Lời nhờ vả ngắn gọn.",
            "4. Lời cảm ơn và ký tên."
        ],
        "cliches": [
            {"ru": "Привет, Саша!", "vi": "Chào Sasha!"},
            {"ru": "Я ушёл в супермаркет за продуктами.", "vi": "Tôi đã đi siêu thị mua thực phẩm."},
            {"ru": "Я вернусь через час.", "vi": "Tôi sẽ về sau 1 tiếng."},
            {"ru": "Свари, пожалуйста, рис на ужин.", "vi": "Bạn nấu giúp cơm bữa tối nhé."}
        ],
        "modelAnswer": "Привет, Саша!\n\nЯ ушёл в супермаркет купить хлеб, молоко и овощи. Я вернусь примерно в шесть вечера. Свари, пожалуйста, рис на ужин. Деньги лежат на столе.\n\nСпасибо большое!\nАнтон.",
        "modelWordCount": 35,
        "rubric": [{"criteria": "Mục đích giao tiếp", "desc": "Nêu rõ tin nhắn cần để lại."}, {"criteria": "Ngữ pháp A1", "desc": "Đúng thì quá khứ và mệnh lệnh thức."}]
    },
    {
        "id": 22,
        "level": "A1",
        "skill": "writing",
        "title": "Открытка маме на 8 Марта",
        "writingTime": "20 phút",
        "topic": "Viết bưu thiếp chúc mừng mẹ ngày Quốc tế Phụ nữ",
        "targetWordRange": [35, 55],
        "task": "Viết một tấm thiệp ngắn (35–55 từ) chúc mừng mẹ nhân ngày 8/3 bằng tiếng Nga.",
        "requirements": [
            "1. Lời chào mẹ yêu thương (Дорогая мама!).",
            "2. Lời chúc mừng ngày 8/3.",
            "3. Lời chúc sức khỏe, nụ cười và hạnh phúc.",
            "4. Lời bày tỏ tình yêu thương (Я тебя очень люблю!)."
        ],
        "cliches": [
            {"ru": "Дорогая и любимая мама!", "vi": "Mẹ yêu quý của con!"},
            {"ru": "Поздравляю тебя с праздником Восьмого марта!", "vi": "Chúc mừng mẹ ngày 8 tháng 3!"},
            {"ru": "Желаю тебе крепкого здоровья и радости!", "vi": "Chúc mẹ dồi dào sức khỏe và niềm vui!"},
            {"ru": "Спасибо за твою заботу и любовь!", "vi": "Cảm ơn mẹ vì sự chăm sóc và tình yêu thương!"}
        ],
        "modelAnswer": "Дорогая, любимая мама!\n\nОт всего сердца поздравляю тебя с праздником Восьмого марта! Желаю тебе крепкого здоровья, весеннего настроения, улыбок и счастья! Спасибо за твою бесконечную заботу. Я тебя очень сильно люблю!\n\nТвой сын, Минх.",
        "modelWordCount": 38,
        "rubric": [{"criteria": "Tình cảm & Ngữ cảnh", "desc": "Diễn đạt ấm áp, đúng thể thức chúc mừng."}, {"criteria": "Cấu trúc chúc", "desc": "Dùng đúng «Желаю тебе + Cách 2»."}]
    },
    {
        "id": 23,
        "level": "A1",
        "skill": "writing",
        "title": "Анкета для библиотеки",
        "writingTime": "15 phút",
        "topic": "Điền tờ khai thông tin cá nhân làm thẻ thư viện",
        "targetWordRange": [25, 45],
        "task": "Viết đoạn văn ngắn (25–45 từ) cung cấp thông tin cá nhân để đăng ký thẻ thư viện trường đại học.",
        "requirements": [
            "1. Họ và tên, quốc tịch.",
            "2. Năm sinh, số điện thoại.",
            "3. Khoa và ngành đang học."
        ],
        "cliches": [
            {"ru": "Моя фамилия и имя...", "vi": "Họ và tên của tôi..."},
            {"ru": "Я гражданин Вьетнама.", "vi": "Tôi là công dân Việt Nam."},
            {"ru": "Я студент первого курса филологического факультета.", "vi": "Tôi là sinh viên năm nhất khoa ngữ văn."}
        ],
        "modelAnswer": "Меня зовут Нгуен Ван Ан. Я гражданин Вьетнама. Я родился в 2004 году в Ханое. Сейчас я студент первого курса филологического факультета. Мой номер телефона: +7-999-123-45-67. Мой адрес: улица Ленина, дом 10, комната 204.",
        "modelWordCount": 37,
        "rubric": [{"criteria": "Tính chính xác thông tin", "desc": "Cung cấp đầy đủ các trường thông tin theo yêu cầu."}]
    },
    {
        "id": 24,
        "level": "A1",
        "skill": "writing",
        "title": "Мой любимый день недели",
        "writingTime": "20 phút",
        "topic": "Viết đoạn văn ngắn kể về thứ Bảy yêu thích",
        "targetWordRange": [40, 60],
        "task": "Viết đoạn văn (40–60 từ) kể về ngày thứ Bảy của bạn: bạn làm gì và tại sao bạn thích ngày này.",
        "requirements": [
            "1. Nêu ngày yêu thích trong tuần.",
            "2. Kể các hoạt động buổi sáng, chiều, tối.",
            "3. Cảm xúc của bạn."
        ],
        "cliches": [
            {"ru": "Мой любимый день недели — это суббота.", "vi": "Ngày yêu thích trong tuần của tôi là thứ Bảy."},
            {"ru": "В этот день мне не нужно рано вставать.", "vi": "Vào ngày này tôi không cần dậy sớm."},
            {"ru": "Вечером мы с друзьями смотрим кино.", "vi": "Buổi tối tôi và bạn bè xem phim."}
        ],
        "modelAnswer": "Мой самый любимый день недели — это суббота. В этот день мне не нужно рано вставать на лекции. Утром я долго сплю и вкусно завтракаю. Днём я гуляю в красивом парке или катаюсь на велосипеде. А вечером мы с друзьями пьём чай и смотрим комедии. Суббота — это лучший день!",
        "modelWordCount": 51,
        "rubric": [{"criteria": "Mạch lạc", "desc": "Có liên kết thời gian buổi sáng, chiều, tối."}]
    },
    {
        "id": 25,
        "level": "A1",
        "skill": "writing",
        "title": "Итоговое эссе A1: Мой путь в изучении русского языка",
        "writingTime": "25 phút",
        "topic": "Bài luận tổng kết A1: Chặng đường đã qua và mục tiêu chinh phục A2",
        "targetWordRange": [50, 70],
        "task": "Viết một bài luận ngắn (50–70 từ) tổng kết quá trình học tiếng Nga trình độ A1 của bạn.",
        "requirements": [
            "1. Thời gian bạn bắt đầu học tiếng Nga.",
            "2. Điều bạn thấy thú vị và khó khăn nhất.",
            "3. Những gì bạn đã có thể tự tin nói.",
            "4. Mục tiêu tiếp theo ở cấp độ A2."
        ],
        "cliches": [
            {"ru": "Я начал изучать русский язык...", "vi": "Tôi bắt đầu học tiếng Nga..."},
            {"ru": "Сначала было трудно, но очень интересно.", "vi": "Ban đầu hơi khó nhưng rất thú vị."},
            {"ru": "Теперь я могу читать и говорить простые фразы.", "vi": "Bây giờ tôi đã có thể đọc và nói các câu đơn giản."},
            {"ru": "Моя следующая цель — успешно сдать экзамен A2.", "vi": "Mục tiêu tiếp theo của tôi là thi đỗ A2."}
        ],
        "modelAnswer": "Я начал изучать русский язык три месяца назад на платформе TAIGA TALK. Сначала грамматика казалась сложной, но уроки были очень увлекательными. Теперь я знаю алфавит, могу читать тексты, делать покупки и рассказывать о себе. Я очень рад своим успехам! Моя следующая цель — уверенно перейти на уровень A2.",
        "modelWordCount": 52,
        "rubric": [{"criteria": "Nội dung tổng kết", "desc": "Nêu được quá trình và cam kết tiếp tục học lên A2."}]
    }
]

for l in a1_writing_extra:
    filename = f"lesson-{l['id']:02d}.json"
    with open(f"{BASE_DIR}/writing/a1/{filename}", "w", encoding="utf-8") as f:
        json.dump(l, f, ensure_ascii=False, indent=2)

with open(f"{BASE_DIR}/writing/a1/manifest.json", "r", encoding="utf-8") as f:
    manifest_w = json.load(f)
existing_ids = {m["id"] for m in manifest_w}
for l in a1_writing_extra:
    if l["id"] not in existing_ids:
        manifest_w.append({
            "id": l["id"],
            "file": f"lesson-{l['id']:02d}.json",
            "title": l["title"],
            "badge": "A1 Viết",
            "writingTime": l["writingTime"],
            "topic": l["topic"],
            "targetWordRange": l["targetWordRange"]
        })
with open(f"{BASE_DIR}/writing/a1/manifest.json", "w", encoding="utf-8") as f:
    json.dump(manifest_w, f, ensure_ascii=False, indent=2)

print("A1 Writing expanded to 25 lessons.")
