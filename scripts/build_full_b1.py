# -*- coding: utf-8 -*-
"""
Build complete 100 lessons for Level B1 (Reading, Listening, Speaking, Writing).
Each lesson gets:
- Category 1: Danh từ trong bài học (Существительные) - 4-5 words
- Category 2: Động từ then chốt (Глаголы) - 4 words
Total 8-10 words per lesson = ~850 words for B1.
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import scripts.vocab_data_b1 as b1_old

# Reading B1 verbs for lessons 6-25
reading_verbs_b1 = {
    6: [
        {"word": "возводи́ть", "gram": "động từ chưa hoàn thành", "meaning": "xây dựng đền đài", "example": "Древние мастера возводили белокаменные соборы."},
        {"word": "сохраня́ть", "gram": "động từ chưa hoàn thành", "meaning": "bảo tồn di sản", "example": "Город бережно сохраняет архитектурный облик."},
        {"word": "любова́ться", "gram": "động từ chưa hoàn thành", "meaning": "chiêm ngưỡng mái vòm", "example": "Туристы любуются куполами храмов Суздаля."},
        {"word": "иссле́довать", "gram": "động từ chưa hoàn thành", "meaning": "nghiên cứu lịch sử", "example": "Историки исследуют летописи Золотого кольца."}
    ],
    7: [
        {"word": "осва́ивать", "gram": "động từ chưa hoàn thành", "meaning": "chinh phục không gian", "example": "Человечество активно осваивает околоземную орбиту."},
        {"word": "запуска́ть", "gram": "động từ chưa hoàn thành", "meaning": "phóng tên lửa", "example": "Учёные запускают исследовательские спутники."},
        {"word": "пилоти́ровать", "gram": "động từ chưa hoàn thành", "meaning": "lái tàu vũ trụ", "example": "Космонавты пилотируют корабли к орбитальной станции."},
        {"word": "иссле́довать", "gram": "động từ chưa hoàn thành", "meaning": "khám phá vũ trụ", "example": "Аппараты исследуют далёкие планеты Солнечной системы."}
    ],
    8: [
        {"word": "воплоща́ть", "gram": "động từ chưa hoàn thành", "meaning": "hóa thân, thể hiện hình tượng", "example": "Балерина тонко воплощает образ Одетты на сцене."},
        {"word": "исполня́ть", "gram": "động từ chưa hoàn thành", "meaning": "trình diễn vũ đạo", "example": "Танцовщики виртуозно исполняют сложные фуэте."},
        {"word": "восхища́ть", "gram": "động từ chưa hoàn thành", "meaning": "làm say đắm khán giả", "example": "Русский балет восхищает ценителей по всему миру."},
        {"word": "аплоди́ровать", "gram": "động từ chưa hoàn thành", "meaning": "vỗ tay tán thưởng rực rỡ", "example": "Зал стоя аплодирует солистам Большого театра."}
    ],
    9: [
        {"word": "размышля́ть", "gram": "động từ chưa hoàn thành", "meaning": "suy ngẫm nhân sinh", "example": "Толстой глубоко размышляет о смысле жизни человека."},
        {"word": "иссле́довать", "gram": "động từ chưa hoàn thành", "meaning": "mổ xẻ tâm lý", "example": "Достоевский исследует психологию совести и греха."},
        {"word": "отража́ть", "gram": "động từ chưa hoàn thành", "meaning": "phản ánh thời đại", "example": "Классическая литература отражает духовные поиски эпохи."},
        {"word": "вдохновля́ть", "gram": "động từ chưa hoàn thành", "meaning": "truyền cảm hứng sống", "example": "Гуманистические идеи вдохновляют читателей на добро."}
    ],
    10: [
        {"word": "внедря́ть", "gram": "động từ chưa hoàn thành", "meaning": "áp dụng công nghệ sạch", "example": "Предприятия внедряют безотходные технологии переработки."},
        {"word": "сокраща́ть", "gram": "động từ chưa hoàn thành", "meaning": "cắt giảm khí thải", "example": "Страны стремятся сокращать выбросы углекислого газа."},
        {"word": "переходи́ть", "gram": "động từ chưa hoàn thành", "meaning": "chuyển sang năng lượng tái tạo", "example": "Энергетика переходит на солнце и силу ветра."},
        {"word": "беречь", "gram": "động từ chưa hoàn thành", "meaning": "gìn giữ hệ sinh thái", "example": "Мы обязаны беречь природные ресурсы планеты."}
    ],
    11: [
        {"word": "реформи́ровать", "gram": "động từ chưa hoàn thành", "meaning": "cải cách toàn diện", "example": "Пётр Первый решительно реформировал государство и флот."},
        {"word": "осно́вывать", "gram": "động từ chưa hoàn thành", "meaning": "sáng lập đô thành", "example": "В 1703 году царь основал новую столицу на Неве."},
        {"word": "модернизи́ровать", "gram": "động từ chưa hoàn thành", "meaning": "hiện đại hóa kinh tế", "example": "Реформы помогли модернизировать русскую армию."},
        {"word": "влия́ть", "gram": "động từ chưa hoàn thành", "meaning": "tác động lịch sử", "example": "Преобразования повлияли на весь ход европейской истории."}
    ],
    12: [
        {"word": "сочиня́ть", "gram": "động từ chưa hoàn thành", "meaning": "sáng tác cổ tích", "example": "Народ веками сочинял мудрые сказки и былины."},
        {"word": "передава́ть", "gram": "động từ chưa hoàn thành", "meaning": "truyền miệng qua đời", "example": "Сказания устно передавались от дедов к внукам."},
        {"word": "воспева́ть", "gram": "động từ chưa hoàn thành", "meaning": "ngợi ca lòng dũng cảm", "example": "Былины воспевают подвиги русских богатырей."},
        {"word": "храни́ть", "gram": "động từ chưa hoàn thành", "meaning": "lưu giữ tâm hồn dân tộc", "example": "Фольклор бережно хранит вековую мудрость народа."}
    ],
    13: [
        {"word": "разрабатыва́ть", "gram": "động từ chưa hoàn thành", "meaning": "phát triển thuật toán", "example": "Инженеры разрабатывают нейросети нового поколения."},
        {"word": "анализи́ровать", "gram": "động từ chưa hoàn thành", "meaning": "phân tích dữ liệu lớn", "example": "Алгоритмы мгновенно анализируют терабайты информации."},
        {"word": "автоматизи́ровать", "gram": "động từ chưa hoàn thành", "meaning": "tự động hóa quy trình", "example": "Системы ИИ автоматизируют сложную рутинную работу."},
        {"word": "контроли́ровать", "gram": "động từ chưa hoàn thành", "meaning": "kiểm soát độ an toàn", "example": "Человек обязан контролировать этические границы технологий."}
    ],
    14: [
        {"word": "торгова́ть", "gram": "động từ chưa hoàn thành", "meaning": "buôn bán giao thương", "example": "Купцы торговали ценными мехами, льном и шёлком."},
        {"word": "прокла́дывать", "gram": "động từ chưa hoàn thành", "meaning": "mở tuyến hành lang", "example": "Торговые караваны прокладывали пути через континенты."},
        {"word": "соединя́ть", "gram": "động từ chưa hoàn thành", "meaning": "kết nối Đông Tây", "example": "Великий шёлковый путь соединял разные цивилизации."},
        {"word": "обме́ниваться", "gram": "động từ chưa hoàn thành", "meaning": "trao đổi văn hóa hàng hóa", "example": "Народы активно обменивались товарами и научными знаниями."}
    ],
    15: [
        {"word": "перевоплоща́ться", "gram": "động từ chưa hoàn thành", "meaning": "hóa thân trọn vẹn", "example": "По системе Станиславского актёр должен искренне перевоплощаться."},
        {"word": "пережива́ть", "gram": "động từ chưa hoàn thành", "meaning": "sống cùng nhân vật", "example": "Артист глубоко переживает чувства своего героя на сцене."},
        {"word": "ве́рить", "gram": "động từ chưa hoàn thành", "meaning": "tin vào chân thực", "example": "Зритель должен безусловно верить происходящему на подмостках."},
        {"word": "ста́вить (спектакль)", "gram": "cụm động từ", "meaning": "dàn dựng tác phẩm", "example": "Режиссёр ставит психологическую драму Чехова."}
    ],
    16: [
        {"word": "иссле́довать", "gram": "động từ chưa hoàn thành", "meaning": "thám hiểm núi lửa", "example": "Геологи исследуют действующие вулканы Камчатки."},
        {"word": "изверга́ться", "gram": "động từ chưa hoàn thành", "meaning": "phun trào nham thạch", "example": "Вулкан извергается потоками раскалённой лавы и пепла."},
        {"word": "изуча́ть", "gram": "động từ chưa hoàn thành", "meaning": "quan trắc địa chấn", "example": "Сейсмологи изучают колебания земной коры в реальном времени."},
        {"word": "охраня́ть", "gram": "động từ chưa hoàn thành", "meaning": "bảo vệ di sản tự nhiên", "example": "Государство строго охраняет дикую природу заповедников."}
    ],
    17: [
        {"word": "сочиня́ть", "gram": "động từ chưa hoàn thành", "meaning": "soạn nhạc giao hưởng", "example": "Чайковский сочинял гениальные симфонии и балеты."},
        {"word": "дирижи́ровать", "gram": "động từ chưa hoàn thành", "meaning": "chỉ huy dàn nhạc", "example": "Маэстро виртуозно дирижирует сводным симфоническим оркестром."},
        {"word": "звуча́ть", "gram": "động từ chưa hoàn thành", "meaning": "vang lên âm hưởng", "example": "В зале филармонии звучит бессмертная классическая музыка."},
        {"word": "покоря́ть", "gram": "động từ chưa hoàn thành", "meaning": "chinh phục thính giả", "example": "Русская музыка покоряет сердца слушателей на всех континентах."}
    ],
    18: [
        {"word": "зава́ривать", "gram": "động từ chưa hoàn thành", "meaning": "hãm trà thơm", "example": "Хозяйка заваривает крепкий чай с сушёными травами."},
        {"word": "собира́ться", "gram": "động từ chưa hoàn thành", "meaning": "quây quần bên ấm samovar", "example": "Вся семья вечерами собирается у кипящего самовара."},
        {"word": "угоща́ть", "gram": "động từ chưa hoàn thành", "meaning": "mời trà bánh", "example": "Гостей угощают чаем с липовым мёдом и сушками."},
        {"word": "бесе́довать", "gram": "động từ chưa hoàn thành", "meaning": "hàn huyên tâm sự", "example": "За долгим чаепитием приятно вести душевные неспешные беседы."}
    ],
    19: [
        {"word": "открыва́ть (закон)", "gram": "cụm động từ", "meaning": "phát minh định luật", "example": "Менделеев открыл фундаментальный Периодический закон химических элементов."},
        {"word": "систематизи́ровать", "gram": "động từ chưa hoàn thành", "meaning": "hệ thống hóa bảng tuần hoàn", "example": "Учёный гениально систематизировал свойства всех известных элементов."},
        {"word": "предска́зывать", "gram": "động từ chưa hoàn thành", "meaning": "tiên đoán chất mới", "example": "Таблица позволила точно предсказать свойства ещё не открытых веществ."},
        {"word": "прославля́ть", "gram": "động từ chưa hoàn thành", "meaning": "làm rạng danh khoa học", "example": "Великие открытия прославляют отечественную научную мысль."}
    ],
    20: [
        {"word": "помога́ть", "gram": "động từ chưa hoàn thành", "meaning": "hỗ trợ người yếu thế", "example": "Волонтёры бескорыстно помогают одиноким пожилым людям."},
        {"word": "организо́вывать", "gram": "động từ chưa hoàn thành", "meaning": "tổ chức phong trào", "example": "Молодёжь организовывает благотворительные акции и сборы."},
        {"word": "уча́ствовать", "gram": "động từ chưa hoàn thành", "meaning": "tham gia dọn rác", "example": "Студенты активно участвуют в очистке берегов Байкала."},
        {"word": "объединя́ть", "gram": "động từ chưa hoàn thành", "meaning": "kết nối cộng đồng", "example": "Добрые дела объединяют людей разных возрастов и профессий."}
    ],
    21: [
        {"word": "проекти́ровать", "gram": "động từ chưa hoàn thành", "meaning": "thiết kế quy hoạch", "example": "Урбанисты проектируют удобные пешеходные улицы и парки."},
        {"word": "благоустра́ивать", "gram": "động từ chưa hoàn thành", "meaning": "chỉnh trang không gian", "example": "Власти благоустраивают набережные и дворовые территории."},
        {"word": "создава́ть (среду)", "gram": "cụm động từ", "meaning": "xây dựng môi trường sống", "example": "Город стремится создавать комфортную среду для каждого жителя."},
        {"word": "разгружа́ть", "gram": "động từ chưa hoàn thành", "meaning": "giảm tải giao thông", "example": "Новые линии метро разгружают дороги от автомобильных пробок."}
    ],
    22: [
        {"word": "опи́сывать", "gram": "động từ chưa hoàn thành", "meaning": "khắc họa chân thực", "example": "Чехов лаконично описывает психологию «маленького человека»."},
        {"word": "сопережива́ть", "gram": "động từ chưa hoàn thành", "meaning": "đồng cảm trắc ẩn", "example": "Читатель искренне сопереживает непростой судьбе чеховских персонажей."},
        {"word": "высме́ивать", "gram": "động từ chưa hoàn thành", "meaning": "châm biếm thói xấu", "example": "В коротких сатирических рассказах автор высмеивает пошлость и глупость."},
        {"word": "заставля́ть (думать)", "gram": "cụm động từ", "meaning": "thức tỉnh suy ngẫm", "example": "Тонкая проза заставляет глубоко задуматься о подлинных ценностях."}
    ],
    23: [
        {"word": "соблюда́ть (баланс)", "gram": "cụm động từ", "meaning": "duy trì cân bằng", "example": "Успешные люди соблюдают баланс между карьерой и личной жизнью."},
        {"word": "восстана́вливать", "gram": "động từ chưa hoàn thành", "meaning": "phục hồi năng lượng", "example": "Прогулки на свежем воздухе восстанавливают душевные силы."},
        {"word": "предотвраща́ть", "gram": "động từ chưa hoàn thành", "meaning": "ngăn chặn kiệt sức", "example": "Правильный режим сна предотвращает эмоциональное перегорание."},
        {"word": "забо́титься", "gram": "động từ chưa hoàn thành", "meaning": "chăm sóc tinh thần", "example": "Важно ежедневно заботиться о своём психологическом благополучии."}
    ],
    24: [
        {"word": "стро́ить (без гвоздей)", "gram": "cụm động từ", "meaning": "dựng chùa không đinh", "example": "Северные зодчие строили деревянные храмы Кижей без единого гвоздя."},
        {"word": "сохраня́ть", "gram": "động từ chưa hoàn thành", "meaning": "gìn giữ nguyên vẹn", "example": "Реставраторы бережно сохраняют уникальные памятники зодчества."},
        {"word": "восхища́ть", "gram": "động từ chưa hoàn thành", "meaning": "làm kinh ngạc", "example": "Многоглавые деревянные соборы восхищают туристов своей гармонией."},
        {"word": "защища́ть", "gram": "động từ chưa hoàn thành", "meaning": "bảo vệ di sản thế giới", "example": "ЮНЕСКО защищает Кижский погост как мировое достояние человечества."}
    ],
    25: [
        {"word": "подтвержда́ть", "gram": "động từ chưa hoàn thành", "meaning": "khẳng định năng lực", "example": "Сертификат ТРКИ-1 официально подтверждает владение языком на уровне B1."},
        {"word": "сдава́ть (субтесты)", "gram": "cụm động từ", "meaning": "thi đỗ các phần thi", "example": "Кандидаты успешно сдают субтесты по чтению, письму и говорению."},
        {"word": "демонстри́ровать", "gram": "động từ chưa hoàn thành", "meaning": "thể hiện vốn từ phong phú", "example": "Выпускники демонстрируют свободное владение сложной грамматикой."},
        {"word": "открыва́ть (путь)", "gram": "cụm động từ", "meaning": "mở ra tương lai", "example": "Знание русского языка открывает путь к обучению в лучших вузах России."}
    ]
}

# Reading B1 lessons
reading_b1 = {}
for i in range(1, 26):
    fname = f'public/data/reading/b1/lesson-{i:02d}.json'
    with open(fname, 'r', encoding='utf-8') as f:
        d = json.load(f)
    existing_cats = d.get('vocabCategories', [])
    if i in reading_verbs_b1:
        noun_cat = existing_cats[0]
        noun_cat["title"] = "Danh từ trong bài đọc (Существительные)"
        verb_cat = {
            "title": "Động từ then chốt (Глаголы)",
            "rows": reading_verbs_b1[i]
        }
        reading_b1[i] = [noun_cat, verb_cat]
    else:
        reading_b1[i] = existing_cats

# Listening B1 verbs (Lessons 1-25)
listening_verbs_b1 = {
    1: [
        {"word": "поступа́ть", "gram": "động từ chưa hoàn thành", "meaning": "thi tuyển nghiên cứu sinh", "example": "Иностранцы поступают в аспирантуру российских университетов."},
        {"word": "защища́ть (диссертацию)", "gram": "cụm động từ", "meaning": "bảo vệ luận án tiến sĩ", "example": "Аспирант успешно защищает диссертацию перед учёным советом."},
        {"word": "публикова́ть", "gram": "động từ chưa hoàn thành", "meaning": "công bố bài báo khoa học", "example": "Молодой учёный публикует результаты исследований в профильном журнале."},
        {"word": "стажирова́ться", "gram": "động từ chưa hoàn thành", "meaning": "thực tập trao đổi", "example": "Студенты стажируются в ведущих лабораториях страны."}
    ],
    2: [
        {"word": "путеше́ствовать", "gram": "động từ chưa hoàn thành", "meaning": "du ngoạn xuyên lục địa", "example": "Мы путешествуем по легендарной Транссибирской магистрали."},
        {"word": "пересека́ть (часовые пояса)", "gram": "cụm động từ", "meaning": "băng qua các múi giờ", "example": "Поезд пересекает семь часовых поясов от Москвы до Владивостока."},
        {"word": "любова́ться", "gram": "động từ chưa hoàn thành", "meaning": "ngắm nhìn vẻ đẹp Siberia", "example": "Пассажиры любуются бескрайними таёжными просторами из окна."},
        {"word": "остана́вливаться", "gram": "động từ chưa hoàn thành", "meaning": "dừng chân tại các ga", "example": "Состав останавливается на крупных станциях на полчаса."}
    ],
    3: [
        {"word": "сортирова́ть", "gram": "động từ chưa hoàn thành", "meaning": "phân loại rác tại nguồn", "example": "Сознательные граждане сортируют пластик, стекло и бумагу."},
        {"word": "перераба́тывать", "gram": "động từ chưa hoàn thành", "meaning": "tái chế phế thải", "example": "Завод экологично перерабатывает вторичное сырьё в новую продукцию."},
        {"word": "подде́рживать (инициативы)", "gram": "cụm động từ", "meaning": "ủng hộ sáng kiến xanh", "example": "Город поддерживает экологические инициативы молодёжных движений."},
        {"word": "беречь", "gram": "động từ chưa hoàn thành", "meaning": "bảo vệ thiên nhiên", "example": "Мы обязаны беречь природу ради будущих поколений."}
    ],
    4: [
        {"word": "осва́ивать", "gram": "động từ chưa hoàn thành", "meaning": "khai phá không gian", "example": "Россия первой открыла дорогу к освоению космического пространства."},
        {"word": "выходи́ть (на орбиту)", "gram": "cụm động từ", "meaning": "tiến vào quỹ đạo Trái Đất", "example": "Космический корабль успешно выходит на расчётную околоземную орбиту."},
        {"word": "стыкова́ться", "gram": "động từ chưa hoàn thành", "meaning": "ghép nối với trạm ISS", "example": "Грузовой модуль в автоматическом режиме стыкуется с МКС."},
        {"word": "проводи́ть (эксперименты)", "gram": "cụm động từ", "meaning": "tiến hành thí nghiệm vũ trụ", "example": "Космонавты проводят уникальные биологические эксперименты в невесомости."}
    ],
    5: [
        {"word": "запуска́ть (стартап)", "gram": "cụm động từ", "meaning": "khởi nghiệp dự án IT", "example": "Молодые программисты запускают инновационный образовательный стартап."},
        {"word": "привлека́ть (инвестиции)", "gram": "cụm động từ", "meaning": "thu hút vốn đầu tư", "example": "Проект успешно привлекает инвестиции от венчурных фондов."},
        {"word": "разрабатыва́ть", "gram": "động từ chưa hoàn thành", "meaning": "phát triển ứng dụng", "example": "Команда разрабатывает мобильное приложение для изучения языков."},
        {"word": "выходи́ть (на рынок)", "gram": "cụm động từ", "meaning": "tiếp cận thị trường quốc tế", "example": "Компания уверенно выходит на глобальный технологический рынок."}
    ],
    6: [
        {"word": "выставля́ть", "gram": "động từ chưa hoàn thành", "meaning": "trưng bày kiệt tác hội họa", "example": "В залах Третьяковской галереи выставляют шедевры русских мастеров."},
        {"word": "осма́тривать", "gram": "động từ chưa hoàn thành", "meaning": "ngắm nhìn các tác phẩm", "example": "Посетители подолгу осматривают полотна Репина и Сурикова."},
        {"word": "собира́ть (коллекцию)", "gram": "cụm động từ", "meaning": "sưu tầm tranh quý", "example": "Павел Третьяков всю жизнь собирал лучшую коллекцию русской живописи."},
        {"word": "восхища́ть", "gram": "động từ chưa hoàn thành", "meaning": "làm rung động lòng người", "example": "Русское изобразительное искусство восхищает ценителей со всего мира."}
    ],
    7: [
        {"word": "меня́ться", "gram": "động từ chưa hoàn thành", "meaning": "biến đổi khí hậu toàn cầu", "example": "Климат на планете стремительно меняется из-за деятельности человека."},
        {"word": "тая́ть", "gram": "động từ chưa hoàn thành", "meaning": "tan chảy các sông băng", "example": "Вековые ледники тают, приводя к постепенному подъёму уровня океана."},
        {"word": "сокраща́ть (выбросы)", "gram": "cụm động từ", "meaning": "giảm thiểu khí thải độc", "example": "Мировое сообщество договаривается сокращать промышленные выбросы в атмосферу."},
        {"word": "предотвраща́ть", "gram": "động từ chưa hoàn thành", "meaning": "ngăn chặn thảm họa sinh thái", "example": "Учёные ищут решения, чтобы предотвратить глобальный экологический кризис."}
    ],
    8: [
        {"word": "арендова́ть", "gram": "động từ chưa hoàn thành", "meaning": "thuê mặt bằng kinh doanh", "example": "Компания планирует арендовать офис в современном деловом центре."},
        {"word": "заключа́ть (договор)", "gram": "cụm động từ", "meaning": "ký kết hợp đồng thuê", "example": "Арендатор и собственник заключают официальный договор аренды помещения."},
        {"word": "опла́чивать (коммуналку)", "gram": "cụm động từ", "meaning": "trả phí dịch vụ điện nước", "example": "Арендатор обязан своевременно оплачивать коммунальные платежи."},
        {"word": "расторга́ть", "gram": "động từ chưa hoàn thành", "meaning": "chấm dứt thỏa thuận", "example": "При нарушении условий стороны вправе расторгнуть соглашение."}
    ],
    9: [
        {"word": "обуча́ться", "gram": "động từ chưa hoàn thành", "meaning": "học tập từ xa trực tuyến", "example": "Студенты эффективно обучаются на современных цифровых образовательных платформах."},
        {"word": "проводи́ть (вебинары)", "gram": "cụm động từ", "meaning": "tổ chức lớp học online", "example": "Профессор еженедельно проводит интерактивные вебинары для заочников."},
        {"word": "осва́ивать (материал)", "gram": "cụm động từ", "meaning": "tiếp thu bài giảng số", "example": "Онлайн-формат требует умения самостоятельно осваивать учебный материал."},
        {"word": "взаимоде́йствовать", "gram": "động từ chưa hoàn thành", "meaning": "tương tác qua mạng", "example": "Слушатели активно взаимодействуют с преподавателем в чате курса."}
    ],
    10: [
        {"word": "демонстри́ровать", "gram": "động từ chưa hoàn thành", "meaning": "công chiếu tác phẩm mới", "example": "На кинофестивале демонстрируют премьерные работы независимых режиссёров."},
        {"word": "оце́нивать", "gram": "động từ chưa hoàn thành", "meaning": "chấm điểm tranh tài", "example": "Международное жюри строго оценивает режиссуру и актёрскую игру."},
        {"word": "вруча́ть (награды)", "gram": "cụm động từ", "meaning": "trao giải thưởng cao quý", "example": "На торжественной церемонии председатель жюри вручает гран-при фестиваля."},
        {"word": "номини́ровать", "gram": "động từ chưa hoàn thành", "meaning": "đề cử hạng mục", "example": "Историческую драму номинировали на приз за лучший сценарий."}
    ],
    11: [
        {"word": "подпи́сывать (контракт)", "gram": "cụm động từ", "meaning": "ký hợp đồng làm việc", "example": "Перед началом работы сотрудник подписывает трудовой договор с работодателем."},
        {"word": "соблюда́ть (права)", "gram": "cụm động từ", "meaning": "tuân thủ quyền lợi lao động", "example": "Организация обязана строго соблюдать трудовые права каждого специалиста."},
        {"word": "консульти́ровать", "gram": "động từ chưa hoàn thành", "meaning": "tư vấn điều khoản luật", "example": "Корпоративный юрист подробно консультирует персонал по вопросам выплат."},
        {"word": "выполня́ть (обязанности)", "gram": "cụm động từ", "meaning": "hoàn thành nghĩa vụ công tác", "example": "Работник добросовестно выполняет свои должностные обязанности."}
    ],
    12: [
        {"word": "любова́ться", "gram": "động từ chưa hoàn thành", "meaning": "thưởng ngoạn đền cổ Suzdal", "example": "Туристы подолгу любуются белокаменными соборами древнего Suzdal."},
        {"word": "сохраня́ть (наследие)", "gram": "cụm động từ", "meaning": "bảo tồn di tích vô giá", "example": "Музей-заповедник бережно сохраняет шедевры древнерусского зодчества."},
        {"word": "знако́мить (с историей)", "gram": "cụm động từ", "meaning": "giới thiệu lịch sử hào hùng", "example": "Экскурсовод увлекательно знакомит гостей с эпохой княжеской Руси."},
        {"word": "реставри́ровать", "gram": "động từ chưa hoàn thành", "meaning": "trùng tu công trình cổ", "example": "Специалисты бережно реставрируют старинные белокаменные храмы."}
    ],
    13: [
        {"word": "обсужда́ть (повесть)", "gram": "cụm động từ", "meaning": "thảo luận tác phẩm Pushkin", "example": "В книжном клубе читатели оживлённо обсуждают повесть Пушкина «Метель»."},
        {"word": "анализи́ровать", "gram": "động từ chưa hoàn thành", "meaning": "phân tích diễn biến tâm lý", "example": "Участники встречи глубоко анализируют поступки и судьбу главных героев."},
        {"word": "переосмы́сливать", "gram": "động từ chưa hoàn thành", "meaning": "nhìn nhận lại thông điệp", "example": "Современное поколение по-новому переосмысливает классические сюжеты."},
        {"word": "дели́ться (мнением)", "gram": "cụm động từ", "meaning": "bày tỏ quan điểm cá nhân", "example": "Каждый читатель с увлечением делится своим пониманием финала книги."}
    ],
    14: [
        {"word": "стартова́ть", "gram": "động từ chưa hoàn thành", "meaning": "rời bệ phóng vũ trụ", "example": "Ракета-носитель с пилотируемым кораблём успешно стартует с космодрома."},
        {"word": "пилоти́ровать", "gram": "động từ chưa hoàn thành", "meaning": "điều khiển tàu bay", "example": "Опытный экипаж уверенно пилотирует корабль на пути к международной станции."},
        {"word": "стыкова́ться (с МКС)", "gram": "cụm động từ", "meaning": "kết nối ghép tàu lên trạm", "example": "Космический аппарат плавно стыкуется с герметичным шлюзом МКС."},
        {"word": "трансли́ровать", "gram": "động từ chưa hoàn thành", "meaning": "truyền hình trực tiếp", "example": "Телеканалы в прямом эфире транслируют запуск космической экспедиции."}
    ],
    15: [
        {"word": "утвержда́ть (тему)", "gram": "cụm động từ", "meaning": "phê duyệt đề tài luận án", "example": "Научный совет кафедры утверждает тему диссертационного исследования."},
        {"word": "формули́ровать", "gram": "động từ chưa hoàn thành", "meaning": "xác định mục tiêu đề tài", "example": "Аспирант чётко формулирует цели, задачи и научную новизну работы."},
        {"word": "сове́товаться", "gram": "động từ chưa hoàn thành", "meaning": "trao đổi giáo viên hướng dẫn", "example": "Исследователь регулярно советуется со своим научным руководителем."},
        {"word": "обосно́вывать", "gram": "động từ chưa hoàn thành", "meaning": "chứng minh tính cấp thiết", "example": "Автор убедительно обосновывает практическую актуальность проекта."}
    ],
    16: [
        {"word": "проекти́ровать", "gram": "động từ chưa hoàn thành", "meaning": "thiết kế phong cách kiến tạo", "example": "Архитекторы конструктивизма проектировали здания с чистой геометрией."},
        {"word": "отража́ть (эпоху)", "gram": "cụm động từ", "meaning": "thể hiện tinh thần tiền phong", "example": "Советский авангард ярко отражал революционный дух 1920-х годов."},
        {"word": "сочета́ть", "gram": "động từ chưa hoàn thành", "meaning": "kết hợp thẩm mỹ và công năng", "example": "Зодчие гармонично сочетали строгую форму и высокую функциональность."},
        {"word": "вдохновля́ть", "gram": "động từ chưa hoàn thành", "meaning": "truyền cảm hứng kiến trúc sư", "example": "Памятники конструктивизма вдохновляют современных мировых дизайнеров."}
    ],
    17: [
        {"word": "создава́ть (блюдо)", "gram": "cụm động từ", "meaning": "sáng tạo món ăn độc đáo", "example": "Шеф-повар ресторана создаёт авторские блюда современной русской кухни."},
        {"word": "сочета́ть (вкусы)", "gram": "cụm động từ", "meaning": "phối ngẫu hương vị", "example": "Мастер смело сочетает традиционные северные ягоды и свежую рыбу."},
        {"word": "переосмы́сливать", "gram": "động từ chưa hoàn thành", "meaning": "làm mới công thức cổ xưa", "example": "Гастрономия переосмысливает старинные купеческие рецепты XIX века."},
        {"word": "презентова́ть", "gram": "động từ chưa hoàn thành", "meaning": "trình diễn thực đơn mới", "example": "Ресторан с гордостью презентует гостям своё новое сезонное меню."}
    ],
    18: [
        {"word": "автоматизи́ровать", "gram": "động từ chưa hoàn thành", "meaning": "tự động hóa sản xuất", "example": "Алгоритмы машинного обучения автоматизируют сложную аналитическую работу."},
        {"word": "обраба́тывать (данные)", "gram": "cụm động từ", "meaning": "xử lý dữ liệu lớn", "example": "Нейронные сети мгновенно обрабатывают колоссальные объёмы информации."},
        {"word": "генери́ровать", "gram": "động từ chưa hoàn thành", "meaning": "tạo sinh nội dung thông minh", "example": "Генеративный искусственный интеллект создаёт качественные тексты и код."},
        {"word": "контроли́ровать", "gram": "động từ chưa hoàn thành", "meaning": "giám sát độ an toàn", "example": "Специалисты обязаны строго контролировать этичность принимаемых решений."}
    ],
    19: [
        {"word": "охраня́ть (фауну)", "gram": "cụm động từ", "meaning": "bảo vệ động vật hoang dã", "example": "Сибирский заповедник строго охраняет редкие популяции амурских тигров."},
        {"word": "изуча́ть (экосистемы)", "gram": "cụm động từ", "meaning": "nghiên cứu rừng taiga", "example": "Биологи наблюдают за естественным развитием нетронутых экосистем тайги."},
        {"word": "восстана́вливать", "gram": "động từ chưa hoàn thành", "meaning": "phục hồi các loài quý hiếm", "example": "Учёные восстанавливают численность исчезающих видов флоры и фауны."},
        {"word": "предотвраща́ть", "gram": "động từ chưa hoàn thành", "meaning": "chống nạn săn bắt trộm", "example": "Инспекторы патрулируют территорию, чтобы предотвращать браконьерство."}
    ],
    20: [
        {"word": "поставля́ть (оборудование)", "gram": "cụm động từ", "meaning": "cung ứng máy móc thiết bị", "example": "Завод обязуется поставлять высокотехнологичное оборудование в срок."},
        {"word": "согласо́вывать", "gram": "động từ chưa hoàn thành", "meaning": "thống nhất thông số kỹ thuật", "example": "Инженеры согласуют техническую спецификацию производственной линии."},
        {"word": "вести́ (переговоры)", "gram": "cụm động từ", "meaning": "đàm phán thương vụ lớn", "example": "Руководители ведут конструктивные переговоры о долгосрочном партнёрстве."},
        {"word": "гаранти́ровать", "gram": "động từ chưa hoàn thành", "meaning": "bảo hành chất lượng", "example": "Поставщик гарантирует сервисное обслуживание на протяжении трёх лет."}
    ],
    21: [
        {"word": "предотвраща́ть", "gram": "động từ chưa hoàn thành", "meaning": "phòng tránh suy nhược", "example": "Грамотное планирование дня помогает предотвращать хроническое переутомление."},
        {"word": "укрепля́ть (иммунитет)", "gram": "cụm động từ", "meaning": "tăng cường đề kháng", "example": "Полноценный сон и сбалансированное питание укрепляют защитные силы организма."},
        {"word": "распределя́ть (нагрузку)", "gram": "cụm động từ", "meaning": "phân bổ khối lượng học tập", "example": "Студенты учатся равномерно распределять умственную нагрузку в сессию."},
        {"word": "восстана́вливаться", "gram": "động từ chưa hoàn thành", "meaning": "hồi phục sau mệt mỏi", "example": "Регулярный отдых на природе позволяет быстро восстанавливаться психологически."}
    ],
    22: [
        {"word": "ста́вить (премьеру)", "gram": "cụm động từ", "meaning": "công diễn vở kịch lớn", "example": "Известный драматический театр ставит премьеру по пьесе классика."},
        {"word": "гастроли́ровать", "gram": "động từ chưa hoàn thành", "meaning": "đi lưu diễn quốc tế", "example": "Театральная труппа с успехом гастролирует по городам Европы и Азии."},
        {"word": "покоря́ть (зрителей)", "gram": "cụm động từ", "meaning": "chinh phục khán trường", "example": "Актёры покоряют зрителей искренностью переживаний и глубиной игры."},
        {"word": "собира́ть (аншлаги)", "gram": "cụm động từ", "meaning": "cháy vé mọi đêm diễn", "example": "Новая постановка собирает полные аншлаги на протяжении всего сезона."}
    ],
    23: [
        {"word": "проходи́ть (собеседование)", "gram": "cụm động từ", "meaning": "tham dự phỏng vấn ứng tuyển", "example": "Выпускник успешно проходит собеседование в престижную международную компанию."},
        {"word": "демонстри́ровать", "gram": "động từ chưa hoàn thành", "meaning": "thể hiện năng lực chuyên môn", "example": "Кандидат уверенно демонстрирует свои профессиональные и гибкие навыки."},
        {"word": "претендова́ть", "gram": "động từ chưa hoàn thành", "meaning": "ứng cử vị trí công việc", "example": "Молодой специалист претендует на вакансию ведущего аналитика данных."},
        {"word": "трудоустра́иваться", "gram": "động từ chưa hoàn thành", "meaning": "nhận việc chính thức", "example": "Ярмарка вакансий помогает сотням студентов быстро трудоустроиться."}
    ],
    24: [
        {"word": "путеше́ствовать", "gram": "động từ chưa hoàn thành", "meaning": "du lịch sinh thái dã ngoại", "example": "Горожане любят путешествовать на байдарках по чистым озёрам Селигера."},
        {"word": "ста́вить (палатку)", "gram": "cụm động từ", "meaning": "dựng lều cắm trại ven hồ", "example": "Туристы ставят палатки на живописном песчаном берегу соснового бора."},
        {"word": "любова́ться (природой)", "gram": "cụm động từ", "meaning": "thưởng ngoạn thiên nhiên hoang sơ", "example": "Мы часами любуемся зеркальной гладью воды и закатами над озером."},
        {"word": "беречь", "gram": "động từ chưa hoàn thành", "meaning": "giữ sạch môi trường rừng", "example": "Экотуристы всегда убирают за собой мусор и берегут чистоту природы."}
    ],
    25: [
        {"word": "сдава́ть (экзамен ТРКИ)", "gram": "cụm động từ", "meaning": "hoàn thành kỳ thi quốc gia", "example": "Иностранные учащиеся успешно сдают сертификационный экзамен ТРКИ-1."},
        {"word": "подтвержда́ть (уровень)", "gram": "cụm động từ", "meaning": "chứng minh trình độ B1", "example": "Высокие баллы тестирования официально подтверждают уровень B1 по шкале CEFR."},
        {"word": "демонстри́ровать", "gram": "động từ chưa hoàn thành", "meaning": "thể hiện năng lực ngôn ngữ", "example": "Студенты демонстрируют глубокое понимание живой разговорной речи на слух."},
        {"word": "горди́ться (успехами)", "gram": "cụm động từ", "meaning": "tự hào thành tích đạt được", "example": "Мы искренне гордимся получением государственного сертификата по русскому языку."}
    ]
}

listening_b1 = {}
for i in range(1, 26):
    noun_cat = b1_old.B1_LISTENING_VOCAB[i][0]
    noun_cat["title"] = "Danh từ trong bài nghe (Существительные)"
    verb_cat = {
        "title": "Động từ then chốt (Глаголы)",
        "rows": listening_verbs_b1[i]
    }
    listening_b1[i] = [noun_cat, verb_cat]

# Speaking B1 verbs (Lessons 1-25)
speaking_verbs_b1 = {
    1: [
        {"word": "изуча́ть", "gram": "động từ chưa hoàn thành", "meaning": "chuyên tâm học ngoại ngữ", "example": "В современном глобальном мире важно изучать несколько иностранных языков."},
        {"word": "преодолева́ть (барьер)", "gram": "cụm động từ", "meaning": "vượt qua rào cản ngôn ngữ", "example": "Постоянная практика помогает преодолевать языковой барьер в общении."},
        {"word": "обща́ться", "gram": "động từ chưa hoàn thành", "meaning": "giao lưu không ranh giới", "example": "Знание языков даёт свободу уверенно общаться с коллегами по всему миру."},
        {"word": "расширя́ть (кругозор)", "gram": "cụm động từ", "meaning": "mở rộng tầm nhìn tri thức", "example": "Изучение русской культуры расширяет интеллектуальный кругозор человека."}
    ],
    2: [
        {"word": "адапти́роваться", "gram": "động từ chưa hoàn thành", "meaning": "thích ứng nhịp sống đô thị", "example": "Жизнь в мегаполисе требует умения быстро адаптироваться к быстрому ритму."},
        {"word": "выбира́ть", "gram": "động từ chưa hoàn thành", "meaning": "chọn nơi định cư", "example": "Многие выбирают спокойную провинцию ради тишины и чистой экологии."},
        {"word": "наслажда́ться", "gram": "động từ chưa hoàn thành", "meaning": "tận hưởng tiện ích thành phố", "example": "Горожане наслаждаются развитой инфраструктурой, парками и музеями."},
        {"word": "устава́ть (от суеты)", "gram": "cụm động từ", "meaning": "mệt mỏi vì ồn ào", "example": "Жители больших городов порой устают от постоянной суеты и дорожных пробок."}
    ],
    3: [
        {"word": "влия́ть", "gram": "động từ chưa hoàn thành", "meaning": "tác động đến nhân sinh quan", "example": "Шедевры литературы и кино глубоко влияют на наше мировоззрение."},
        {"word": "переосмы́сливать", "gram": "động từ chưa hoàn thành", "meaning": "nhìn nhận lại giá trị sống", "example": "После прочтения этой книги человек переосмысливает свои поступки."},
        {"word": "вдохновля́ть", "gram": "động từ chưa hoàn thành", "meaning": "truyền cảm hứng sáng tạo", "example": "Сильные образы героев вдохновляют читателя на самосовершенствование."},
        {"word": "заставля́ть (задуматься)", "gram": "cụm động từ", "meaning": "thức tỉnh lương tri", "example": "Глубокое произведение заставляет задуматься о смысле человеческой жизни."}
    ],
    4: [
        {"word": "трансформи́роваться", "gram": "động từ chưa hoàn thành", "meaning": "chuyển biến thời đại", "example": "Институт семьи постепенно трансформируется под влиянием современной эпохи."},
        {"word": "сохраня́ть (традиции)", "gram": "cụm động từ", "meaning": "gìn giữ giá trị thiêng liêng", "example": "Крепкая семья бережно сохраняет преемственность поколений и обычаи."},
        {"word": "находи́ть (понимание)", "gram": "cụm động từ", "meaning": "thấu hiểu lẫn nhau", "example": "Родители и дети учатся находить взаимопонимание в открытом диалоге."},
        {"word": "подде́рживать", "gram": "động từ chưa hoàn thành", "meaning": "chăm sóc thương yêu", "example": "Близкие люди всегда искренне поддерживают друг друга в любых испытаниях."}
    ],
    5: [
        {"word": "осознава́ть", "gram": "động từ chưa hoàn thành", "meaning": "nhận thức trách nhiệm", "example": "Каждый житель Земли обязан осознавать свой личный экологический след."},
        {"word": "сокраща́ть (потребление)", "gram": "cụm động từ", "meaning": "giảm tiêu dùng phung phí", "example": "Мы можем разумно сокращать потребление одноразового пластика в быту."},
        {"word": "перераба́тывать", "gram": "động từ chưa hoàn thành", "meaning": "tái chế bảo vệ môi trường", "example": "Сортировка бытовых отходов позволяет перерабатывать полезные ресурсы."},
        {"word": "беречь", "gram": "động từ chưa hoàn thành", "meaning": "bảo vệ ngôi nhà chung", "example": "Беречь планету — это не модный тренд, а жизненная необходимость."}
    ],
    6: [
        {"word": "обогаща́ть (душу)", "gram": "cụm động từ", "meaning": "làm giàu tâm hồn", "example": "Настоящее высокое искусство духовно обогащает внутренний мир человека."},
        {"word": "воспи́тывать", "gram": "động từ chưa hoàn thành", "meaning": "bồi dưỡng mỹ cảm", "example": "Классическая музыка воспитывает чуткость, сострадание и эстетический вкус."},
        {"word": "находи́ть (гармонию)", "gram": "cụm động từ", "meaning": "tìm kiếm sự hài hòa", "example": "В картинах великих художников человек находит гармонию и душевный покой."},
        {"word": "сопережива́ть", "gram": "động từ chưa hoàn thành", "meaning": "đồng cảm rung động", "example": "Театр учит зрителей искренне сопереживать чужой боли и радости."}
    ],
    7: [
        {"word": "развива́ться", "gram": "động từ chưa hoàn thành", "meaning": "phát triển du lịch đại trà", "example": "Массовый туризм стремительно развивается благодаря доступности перелётов."},
        {"word": "испы́тывать (нагрузку)", "gram": "cụm động từ", "meaning": "chịu áp lực quá tải", "example": "Исторические города испытывают колоссальную нагрузку от потока туристов."},
        {"word": "сохраня́ть (самобытность)", "gram": "cụm động từ", "meaning": "gìn giữ bản sắc chân thật", "example": "Очень важно сохранять культурную аутентичность древних сёл и памятников."},
        {"word": "уважа́ть", "gram": "động từ chưa hoàn thành", "meaning": "tôn trọng phong tục địa phương", "example": "Путешественники обязаны с уважением относиться к обычаям страны пребывания."}
    ],
    8: [
        {"word": "преодолева́ть (страх)", "gram": "cụm động từ", "meaning": "vượt qua sự e dè", "example": "Практика спонтанной речи помогает преодолевать стеснение и страх ошибок."},
        {"word": "приобрета́ть (уверенность)", "gram": "cụm động từ", "meaning": "xây dựng sự tự tin", "example": "Регулярное общение с носителями позволяет приобретать уверенность в себе."},
        {"word": "формули́ровать (мысли)", "gram": "cụm động từ", "meaning": "diễn đạt ý tưởng trôi chảy", "example": "На уровне B1 учащийся умеет свободно формулировать мысли на русском языке."},
        {"word": "совершенствова́ть", "gram": "động từ chưa hoàn thành", "meaning": "hoàn thiện ngữ điệu phát âm", "example": "Мы совершенствуем произношение, стараясь говорить без сильного акцента."}
    ],
    9: [
        {"word": "предпочита́ть", "gram": "động từ chưa hoàn thành", "meaning": "chuộng sách in truyền thống", "example": "Многие читатели предпочитают бумажные книги из-за тактильных ощущений."},
        {"word": "по́льзоваться", "gram": "động từ chưa hoàn thành", "meaning": "sử dụng máy đọc sách số", "example": "В путешествиях очень удобно пользоваться компактным электронным ридером."},
        {"word": "запомина́ть", "gram": "động từ chưa hoàn thành", "meaning": "ghi nhớ chiều sâu nội dung", "example": "Чтение на бумаге помогает лучше запоминать и осмысливать сложную информацию."},
        {"word": "эконо́мить", "gram": "động từ chưa hoàn thành", "meaning": "tiết kiệm không gian lưu trữ", "example": "Цифровые библиотеки позволяют экономить место в квартире и быстро находить текст."}
    ],
    10: [
        {"word": "вдохновля́ть", "gram": "động từ chưa hoàn thành", "meaning": "truyền cảm hứng lý tưởng", "example": "Выдающиеся исторические личности вдохновляют нас своим мужеством и трудом."},
        {"word": "соверша́ть (подвиг)", "gram": "cụm động từ", "meaning": "lập nên kỳ tích lịch sử", "example": "Юрий Гагарин совершил беспримерный подвиг, открыв эру пилотируемой космонавтики."},
        {"word": "вноси́ть (вклад)", "gram": "cụm động từ", "meaning": "đóng góp lớn cho nhân loại", "example": "Великие учёные внесли неоценимый вклад в развитие цивилизации."},
        {"word": "по́мнить", "gram": "động từ chưa hoàn thành", "meaning": "khắc ghi di sản tiền nhân", "example": "Благодарное человечество свято помнит имена своих героев и первопроходцев."}
    ],
    11: [
        {"word": "осва́ивать (навыки)", "gram": "cụm động từ", "meaning": "trau dồi kỹ năng mềm", "example": "Специалисты будущего обязаны осваивать гибкие навыки и цифровую грамотность."},
        {"word": "адапти́роваться", "gram": "động từ chưa hoàn thành", "meaning": "thích nghi với tự động hóa", "example": "Рынок труда меняется, требуя умения быстро адаптироваться к инновациям."},
        {"word": "быть (востребованным)", "gram": "cụm động từ", "meaning": "được thị trường săn đón", "example": "Специалисты со знанием языков и ИИ всегда будут востребованы работодателями."},
        {"word": "развива́ть", "gram": "động từ chưa hoàn thành", "meaning": "rèn luyện tư duy phản biện", "example": "Очень важно постоянно развивать критическое мышление и креативность."}
    ],
    12: [
        {"word": "ста́лкиваться (с шоком)", "gram": "cụm động từ", "meaning": "đối diện sốc văn hóa", "example": "При переезде за границу иностранец сталкивается с культурным шоком."},
        {"word": "адапти́роваться", "gram": "động từ chưa hoàn thành", "meaning": "hòa nhập môi trường mới", "example": "Участие в жизни университета помогает быстрее адаптироваться в новой стране."},
        {"word": "понима́ть (менталитет)", "gram": "cụm động từ", "meaning": "thấu cảm tâm lý bản xứ", "example": "Искреннее желание понимать местный менталитет разрушает любые барьеры."},
        {"word": "уважа́ть", "gram": "động từ chưa hoàn thành", "meaning": "tôn trọng nếp sống sở tại", "example": "Гость обязан с уважением относиться к культурным нормам и традициям народа."}
    ],
    13: [
        {"word": "формирова́ть (характер)", "gram": "cụm động từ", "meaning": "rèn luyện nhân cách ý chí", "example": "Регулярный спорт формирует сильный характер, дисциплину и целеустремлённость."},
        {"word": "развива́ть (выносливость)", "gram": "cụm động từ", "meaning": "nâng cao sức chịu đựng", "example": "Тренировки на выносливость учат не сдаваться перед лицом трудностей."},
        {"word": "побежда́ть", "gram": "động từ chưa hoàn thành", "meaning": "chiến thắng chính bản thân", "example": "Главная победа в спорте — это умение побеждать собственную лень."},
        {"word": "подде́рживать (команду)", "gram": "cụm động từ", "meaning": "phát huy tinh thần đồng đội", "example": "Командный дух учит поддерживать товарища и отвечать за общий результат."}
    ],
    14: [
        {"word": "мечта́ть (о поездке)", "gram": "cụm động từ", "meaning": "mơ về chuyến Trans-Siberian", "example": "Я давно мечтаю отправиться в путешествие по железной дороге через всю Россию."},
        {"word": "открыва́ть (горизонты)", "gram": "cụm động từ", "meaning": "khám phá vùng đất mới", "example": "Эта поездка открывает перед путешественником величие природы и просторов."},
        {"word": "знако́миться", "gram": "động từ chưa hoàn thành", "meaning": "giao lưu cùng người địa phương", "example": "В пути можно познакомиться с удивительными людьми из разных регионов."},
        {"word": "вдохновля́ться", "gram": "động từ chưa hoàn thành", "meaning": "đắm say trước cảnh sắc", "example": "Красота заснеженных лесов и кристального Байкала вдохновляет на всю жизнь."}
    ],
    15: [
        {"word": "проекти́ровать", "gram": "động từ chưa hoàn thành", "meaning": "quy hoạch không gian nhân văn", "example": "Современные архитекторы проектируют комфортную среду для каждого жителя."},
        {"word": "благоустра́ивать", "gram": "động từ chưa hoàn thành", "meaning": "chỉnh trang đô thị xanh", "example": "Власти благоустраивают общественные пространства, создавая пешеходные зоны."},
        {"word": "создава́ть (удобство)", "gram": "cụm động từ", "meaning": "đem lại cuộc sống tiện nghi", "example": "Умная урбанистика создаёт безопасные и доступные городские кварталы."},
        {"word": "забо́титься (о жителях)", "gram": "cụm động từ", "meaning": "chăm lo cho người đi bộ", "example": "В современном городе в первую очередь заботятся о комфорте пешеходов и экологии."}
    ],
    16: [
        {"word": "иссле́довать", "gram": "động từ chưa hoàn thành", "meaning": "nghiên cứu giới hạn khoa học", "example": "Учёные исследуют генетические механизмы лечения тяжёлых болезней."},
        {"word": "осознава́ть (ответственность)", "gram": "cụm động từ", "meaning": "ý thức trách nhiệm đạo đức", "example": "Научное сообщество обязано осознавать колоссальную ответственность за опыты."},
        {"word": "дискути́ровать", "gram": "động từ chưa hoàn thành", "meaning": "tranh luận vấn đề nhân bản", "example": "Общество бурно дискутирует о моральных границах клонирования человека."},
        {"word": "контроли́ровать", "gram": "động từ chưa hoàn thành", "meaning": "giám sát chặt chẽ thử nghiệm", "example": "Государство должно контролировать безопасность биологических технологий."}
    ],
    17: [
        {"word": "приде́рживаться (ЗОЖ)", "gram": "cụm động từ", "meaning": "theo đuổi lối sống lành mạnh", "example": "Всё больше людей осознанно придерживаются здорового образа жизни."},
        {"word": "пита́ться (правильно)", "gram": "cụm động từ", "meaning": "ăn uống cân bằng dinh dưỡng", "example": "Правильно питаться — значит обеспечивать организм энергией и витаминами."},
        {"word": "занима́ться (физкультурой)", "gram": "cụm động từ", "meaning": "vận động cơ thể mỗi ngày", "example": "Ежедневная двигательная активность дарит бодрость и продлевает молодость."},
        {"word": "отка́зываться (от вредного)", "gram": "cụm động từ", "meaning": "từ bỏ thói quen xấu", "example": "Молодёжь сознательно отказывается от курения и фастфуда ради здоровья."}
    ],
    18: [
        {"word": "пра́здновать", "gram": "động từ chưa hoàn thành", "meaning": "tổ chức lễ truyền thống", "example": "Русские праздники на протяжении веков празднуют с размахом и душой."},
        {"word": "раскрыва́ть (смысл)", "gram": "cụm động từ", "meaning": "hiểu sâu sắc tinh thần lễ", "example": "Рождество и Пасха раскрывают глубокий духовный смысл любви и милосердия."},
        {"word": "соблюда́ть (обычаи)", "gram": "cụm động từ", "meaning": "tuân thủ phong tục tập quán", "example": "В каждой семье бережно соблюдают обычаи гостеприимства и застолья."},
        {"word": "объединя́ть", "gram": "động từ chưa hoàn thành", "meaning": "gắn kết mọi thành viên", "example": "Народные торжества объединяют поколения за общим праздничным столом."}
    ],
    19: [
        {"word": "создава́ть (произведения)", "gram": "cụm động từ", "meaning": "sáng tác văn học đỉnh cao", "example": "Мой любимый писатель создавал непревзойдённые по глубине романы."},
        {"word": "иссле́довать (характеры)", "gram": "cụm động từ", "meaning": "khắc họa tâm hồn nhân vật", "example": "В своих произведениях автор исследует противоречия человеческой души."},
        {"word": "воспева́ть", "gram": "động từ chưa hoàn thành", "meaning": "ngợi ca tình yêu đất nước", "example": "Поэт вдохновенно воспевает красоту родной природы и силу духа народа."},
        {"word": "пробужда́ть (чувства)", "gram": "cụm động từ", "meaning": "đánh thức xúc cảm tốt đẹp", "example": "Высокая литература пробуждает в сердцах читателей сострадание и любовь."}
    ],
    20: [
        {"word": "влия́ть (на ценности)", "gram": "cụm động từ", "meaning": "định hình thang giá trị", "example": "Качественный кинематограф сильно влияет на нравственные ценности молодёжи."},
        {"word": "раскрыва́ть (образы)", "gram": "cụm động từ", "meaning": "xây dựng hình tượng nhân văn", "example": "Режиссёр мастерски раскрывает образы героев, заставляя сопереживать им."},
        {"word": "заставля́ть (размышлять)", "gram": "cụm động từ", "meaning": "thúc đẩy tư duy công lý", "example": "Серьёзное кино заставляет размышлять о справедливости, долге и выборе."},
        {"word": "воспи́тывать", "gram": "động từ chưa hoàn thành", "meaning": "giáo dục đạo đức xã hội", "example": "Шедевры мирового кино воспитывают уважение к человеку и идеалам добра."}
    ],
    21: [
        {"word": "учи́ться (всю жизнь)", "gram": "cụm động từ", "meaning": "học tập không ngừng nghỉ", "example": "Концепция непрерывного образования призывает учиться на протяжении всей жизни."},
        {"word": "повыша́ть (квалификацию)", "gram": "cụm động từ", "meaning": "nâng cao tay nghề trình độ", "example": "Профессионалы регулярно повышают квалификацию на профильных курсах."},
        {"word": "осва́ивать (новое)", "gram": "cụm động từ", "meaning": "tiếp cận kỹ năng số", "example": "Способность осваивать новые сферы помогает оставаться востребованным."},
        {"word": "развива́ться", "gram": "động từ chưa hoàn thành", "meaning": "hoàn thiện bản thân liên tục", "example": "Стремление гармонично развиваться открывает перед человеком новые горизонты."}
    ],
    22: [
        {"word": "помога́ть (бескорыстно)", "gram": "cụm động từ", "meaning": "giúp đỡ bằng cả trái tim", "example": "Волонтёры бескорыстно помогают нуждающимся, сиротам и пожилым людям."},
        {"word": "уча́ствовать (в благотворительности)", "gram": "cụm động từ", "meaning": "làm công tác thiện nguyện", "example": "Каждый сознательный гражданин может участвовать в благотворительных проектах."},
        {"word": "дели́ться (добротой)", "gram": "cụm động từ", "meaning": "lan tỏa tình thương yêu", "example": "Делиться теплом и заботой — это естественная потребность зрелой личности."},
        {"word": "менять (мир к лучшему)", "gram": "cụm động từ", "meaning": "thay đổi xã hội tốt đẹp hơn", "example": "Маленькие добрые дела способны реально менять окружающий мир к лучшему."}
    ],
    23: [
        {"word": "вести́ (диалог)", "gram": "cụm động từ", "meaning": "đối thoại liên văn hóa", "example": "В современном обществе необходимо уметь вести конструктивный диалог культур."},
        {"word": "проявля́ть (толерантность)", "gram": "cụm động từ", "meaning": "thể hiện lòng khoan dung cởi mở", "example": "Важно проявлять толерантность и уважение к чужим обычаям и взглядам."},
        {"word": "предотвраща́ть (конфликты)", "gram": "cụm động từ", "meaning": "xóa bỏ bất đồng hiểu lầm", "example": "Взаимное уважение помогает предотвращать недопонимание и бытовые конфликты."},
        {"word": "находи́ть (компромисс)", "gram": "cụm động từ", "meaning": "tìm tiếng nói chung hòa bình", "example": "Дипломатия учит находить компромисс ради мирного и гармоничного развития."}
    ],
    24: [
        {"word": "развива́ть (инновации)", "gram": "cụm động từ", "meaning": "ứng dụng Internet vạn vật", "example": "Мегаполисы активно развивают концепцию умного города и цифровых сервисов."},
        {"word": "оптимизи́ровать", "gram": "động từ chưa hoàn thành", "meaning": "tối ưu hóa năng lượng giao thông", "example": "Датчики мониторинга оптимизируют освещение и потоки общественного транспорта."},
        {"word": "обеспе́чивать (безопасность)", "gram": "cụm động từ", "meaning": "bảo đảm an toàn cho dân cư", "example": "Умные системы видеонаблюдения обеспечивают высокий уровень безопасности в городе."},
        {"word": "упроща́ть (жизнь)", "gram": "cụm động từ", "meaning": "làm cuộc sống thuận tiện hơn", "example": "Цифровые технологии реально упрощают повседневную жизнь каждого горожанина."}
    ],
    25: [
        {"word": "подводи́ть (итоги)", "gram": "cụm động từ", "meaning": "tổng kết thành quả khóa học", "example": "В выпускной речи я с волнением и гордостью подвожу итоги уровня B1."},
        {"word": "благодари́ть (наставников)", "gram": "cụm động từ", "meaning": "tri ân thầy cô giáo", "example": "Мы от всей души благодарим преподавателей за мудрость, терпение и теплоту."},
        {"word": "горди́ться (достижениями)", "gram": "cụm động từ", "meaning": "tự hào vượt qua thử thách", "example": "Каждый из нас по праву гордится своими успехами в освоении русского языка."},
        {"word": "стреми́ться (к высотам)", "gram": "cụm động từ", "meaning": "vươn tới những đỉnh cao mới", "example": "Мы смело смотрим в будущее и стремимся покорять новые горизонты знаний!"}
    ]
}

speaking_b1 = {}
for i in range(1, 26):
    noun_cat = b1_old.B1_SPEAKING_VOCAB[i][0]
    noun_cat["title"] = "Danh từ trong bài nói (Существительные)"
    verb_cat = {
        "title": "Động từ then chốt (Глаголы)",
        "rows": speaking_verbs_b1[i]
    }
    speaking_b1[i] = [noun_cat, verb_cat]

# Writing B1 verbs (Lessons 1-25)
writing_verbs_b1 = {
    1: [
        {"word": "подава́ть (заявление)", "gram": "cụm động từ", "meaning": "nộp đơn chính thức", "example": "Студент подаёт официальное заявление в международный отдел университета."},
        {"word": "хода́тайствовать", "gram": "động từ chưa hoàn thành", "meaning": "kiến nghị xin gia hạn visa", "example": "Деканат ходатайствует о продлении учебной визы иностранному учащемуся."},
        {"word": "прилага́ть (документы)", "gram": "cụm động từ", "meaning": "đính kèm văn bản xác nhận", "example": "К заявлению я прилагаю академическую справку и копию паспорта."},
        {"word": "проси́ть (рассмотреть)", "gram": "cụm động từ", "meaning": "kính đề nghị xem xét", "example": "Убедительно прошу рассмотреть моё обращение в установленный регламентом срок."}
    ],
    2: [
        {"word": "рассужда́ть", "gram": "động từ chưa hoàn thành", "meaning": "nghị luận hiện tượng xã hội", "example": "В своём эссе автор рассуждает о влиянии социальных сетей на молодёжь."},
        {"word": "анализи́ровать", "gram": "động từ chưa hoàn thành", "meaning": "phân tích hệ lụy mạng ảo", "example": "Психологи анализируют проблему цифровой зависимости и дефицита живого общения."},
        {"word": "аргументи́ровать", "gram": "động từ chưa hoàn thành", "meaning": "lập luận bảo vệ quan điểm", "example": "Автор убедительно аргументирует свою позицию фактами и исследованиями."},
        {"word": "приходи́ть (к выводу)", "gram": "cụm động từ", "meaning": "rút ra kết luận thấu đáo", "example": "В заключении я прихожу к выводу о необходимости цифровой осознанности."}
    ],
    3: [
        {"word": "выража́ть (благодарность)", "gram": "cụm động từ", "meaning": "bày tỏ lòng biết ơn sâu sắc", "example": "Выражаю искреннюю благодарность коллегам за поддержку и ценные советы."},
        {"word": "проходи́ть (стажировку)", "gram": "cụm động từ", "meaning": "hoàn thành kỳ thực tập", "example": "За время стажировки я приобрёл бесценный практический профессиональный опыт."},
        {"word": "цени́ть (опыт)", "gram": "cụm động từ", "meaning": "trân quý sự chỉ dạy của đồng nghiệp", "example": "Я высоко ценю доброжелательную атмосферу и профессионализм всего коллектива."},
        {"word": "наде́яться (на сотрудничество)", "gram": "cụm động từ", "meaning": "hy vọng hợp tác lâu dài", "example": "Искренне надеюсь на продолжение плодотворного сотрудничества в будущем."}
    ],
    4: [
        {"word": "направля́ть (запрос)", "gram": "cụm động từ", "meaning": "gửi công văn liên hệ hội thảo", "example": "Направляю официальный запрос в оргкомитет международной научной конференции."},
        {"word": "уточня́ть (регламент)", "gram": "cụm động từ", "meaning": "hỏi rõ quy chế tham luận", "example": "Прошу уточнить регламент устного выступления и требования к оформлению тезисов."},
        {"word": "подтвержда́ть (участие)", "gram": "cụm động từ", "meaning": "xác nhận dự hội nghị", "example": "С радостью подтверждаю своё очное участие в пленарном заседании."},
        {"word": "выступа́ть (с докладом)", "gram": "cụm động từ", "meaning": "báo cáo đề tài nghiên cứu", "example": "Я планирую выступить с докладом по инновационным методам преподавания."}
    ],
    5: [
        {"word": "рассужда́ть (о доступности)", "gram": "cụm động từ", "meaning": "bàn về bình đẳng giáo dục", "example": "В эссе рассматривается вопрос: должно ли высшее образование быть бесплатным?"},
        {"word": "гаранти́ровать", "gram": "động từ chưa hoàn thành", "meaning": "bảo đảm quyền học tập", "example": "Государство обязано гарантировать равный доступ к качественному образованию."},
        {"word": "финанси́ровать", "gram": "động từ chưa hoàn thành", "meaning": "rót vốn ngân sách cho trường", "example": "Необходимо стабильно финансировать фундаментальные научные кафедры и лаборатории."},
        {"word": "стимули́ровать", "gram": "động từ chưa hoàn thành", "meaning": "khuyến khích người tài", "example": "Бесплатные бюджетные места стимулируют талантливых студентов отлично учиться."}
    ],
    6: [
        {"word": "отклика́ться (на вакансию)", "gram": "cụm động từ", "meaning": "ứng tuyển vị trí công việc", "example": "С большим интересом откликаюсь на открытую вакансию младшего аналитика данных."},
        {"word": "обладать (компетенциями)", "gram": "cụm động từ", "meaning": "có đủ năng lực chuyên môn", "example": "Я обладаю необходимыми знаниями в сфере программирования и анализа данных."},
        {"word": "прилага́ть (резюме)", "gram": "cụm động từ", "meaning": "gửi kèm CV và hồ sơ", "example": "В приложении к сопроводительному письму направляю резюме и портфолио проектов."},
        {"word": "рассчи́тывать (на встречу)", "gram": "cụm động từ", "meaning": "mong được mời phỏng vấn", "example": "Буду рад возможности лично встретиться на собеседовании для обсуждения деталей."}
    ],
    7: [
        {"word": "рецензи́ровать", "gram": "động từ chưa hoàn thành", "meaning": "viết bài phê bình kịch", "example": "Критик профессионально рецензирует новую постановку классической драмы."},
        {"word": "оце́нивать (режиссуру)", "gram": "cụm động từ", "meaning": "đánh giá tài năng đạo diễn", "example": "В статье высоко оценивается смелая режиссёрская концепция спектакля."},
        {"word": "отмеча́ть (игру актёров)", "gram": "cụm động từ", "meaning": "khen ngợi diễn xuất tinh tế", "example": "Рецензент особо отмечает психологическую глубину и искренность актёрской игры."},
        {"word": "рекомендова́ть (к просмотру)", "gram": "cụm động từ", "meaning": "khuyên khán giả nên xem", "example": "Я горячо рекомендую этот вдохновляющий спектакль всем ценителям искусства."}
    ],
    8: [
        {"word": "составля́ть (записку)", "gram": "cụm động từ", "meaning": "lập tờ trình dự án mới", "example": "Руководитель группы составляет докладную записку на имя генерального директора."},
        {"word": "предлага́ть (модернизацию)", "gram": "cụm động từ", "meaning": "đề xuất cải tiến hệ thống", "example": "Мы предлагаем внедрить новую систему автоматизации внутреннего документооборота."},
        {"word": "обосно́вывать (рентабельность)", "gram": "cụm động từ", "meaning": "chứng minh hiệu quả kinh tế", "example": "Экономические расчёты убедительно обосновывают быструю окупаемость проекта."},
        {"word": "оптимизи́ровать", "gram": "động từ chưa hoàn thành", "meaning": "tối ưu hóa quy trình làm việc", "example": "Реализация предложения позволит оптимизировать рабочее время сотрудников отдела."}
    ],
    9: [
        {"word": "сосуществова́ть", "gram": "động từ chưa hoàn thành", "meaning": "chung sống hòa hợp cùng tự nhiên", "example": "Человек обязан научиться гармонично сосуществовать с окружающей природой."},
        {"word": "преодолева́ть (кризис)", "gram": "cụm động từ", "meaning": "vượt qua khủng hoảng sinh thái", "example": "Преодоление экологического кризиса требует совместных скоординированных усилий всех стран."},
        {"word": "растра́чивать (ресурсы)", "gram": "cụm động từ", "meaning": "không phung phí tài nguyên", "example": "Мы не должны безответственно растрачивать невосполнимые богатства Земли."},
        {"word": "сохраня́ть (баланс)", "gram": "cụm động từ", "meaning": "gìn giữ thế cân bằng tự nhiên", "example": "Цель человечества — сохранить хрупкое природное равновесие для будущих потомков."}
    ],
    10: [
        {"word": "направля́ть (жалобу)", "gram": "cụm động từ", "meaning": "gửi đơn khiếu nại dịch vụ", "example": "Турист направляет официальную жалобу руководству туристической компании."},
        {"word": "ука́зывать (на нарушения)", "gram": "cụm động từ", "meaning": "nêu rõ vi phạm hợp đồng", "example": "В тексте претензии подробно указаны факты нарушения условий договора оказания услуг."},
        {"word": "тре́бовать (компенсацию)", "gram": "cụm động từ", "meaning": "đòi bồi thường thiệt hại", "example": "Потребитель требует выплаты соразмерной денежной компенсации за сорванную экскурсию."},
        {"word": "оставля́ть (за собой право)", "gram": "cụm động từ", "meaning": "bảo lưu quyền khởi kiện", "example": "В случае отказа я оставляю за собой право обратиться в суд за защитой своих прав."}
    ],
    11: [
        {"word": "дели́ться (советами)", "gram": "cụm động từ", "meaning": "chia sẻ bí quyết học đại học", "example": "Старшекурсники в студенческой газете делятся полезными советами с первокурсниками."},
        {"word": "адапти́роваться (в вузе)", "gram": "cụm động từ", "meaning": "hòa nhập môi trường giảng đường", "example": "Участие в студенческом совете помогает новичкам быстрее адаптироваться к учёбе."},
        {"word": "организо́вывать (время)", "gram": "cụm động từ", "meaning": "sắp xếp thời gian khoa học", "example": "Важно с первых дней научиться грамотно организовывать время и готовиться к сессии."},
        {"word": "преодолева́ть (трудности)", "gram": "cụm động từ", "meaning": "vượt qua thử thách đầu đời", "example": "Дружба и взаимовыручка помогают преодолевать любые академические сложности."}
    ],
    12: [
        {"word": "докла́дывать (о статусе)", "gram": "cụm động từ", "meaning": "báo cáo tiến độ cho thầy", "example": "Докладываю научному руководителю о текущем статусе подготовки дипломной работы."},
        {"word": "направля́ть (черновик)", "gram": "cụm động từ", "meaning": "gửi bản thảo xin ý kiến", "example": "Направляю на проверку черновик теоретической и практической глав исследования."},
        {"word": "вноси́ть (исправления)", "gram": "cụm động từ", "meaning": "sửa chữa theo góp ý", "example": "Я внимательно изучил все ваши замечания и внёс необходимые исправления в текст."},
        {"word": "согласо́вывать (график)", "gram": "cụm động từ", "meaning": "thống nhất lịch gặp hướng dẫn", "example": "Прошу согласовать дату и время очной консультации на следующей неделе."}
    ],
    13: [
        {"word": "сопоставля́ть", "gram": "động từ chưa hoàn thành", "meaning": "so sánh trung tâm và ngoại ô", "example": "В эссе автор сопоставляет преимущества и недостатки жизни в центре и на окраине."},
        {"word": "выбира́ть (район)", "gram": "cụm động từ", "meaning": "chọn nơi an cư lý tưởng", "example": "Семьи с детьми всё чаще выбирают зелёные спальные районы с чистым воздухом."},
        {"word": "ценить (тишину)", "gram": "cụm động từ", "meaning": "chuộng không gian yên bình", "example": "Вдали от шумных городских магистралей люди ценят покой, безопасность и уют."},
        {"word": "учитыва́ть (доступность)", "gram": "cụm động từ", "meaning": "cân nhắc giao thông tiện lợi", "example": "При покупке квартиры важно учитывать близость к метро и транспортную доступность."}
    ],
    14: [
        {"word": "обосно́вывать (мотивацию)", "gram": "cụm động từ", "meaning": "trình bày khát khao học bổng", "example": "В мотивационном письме кандидат подробно обосновывает выбор данной программы."},
        {"word": "подтвержда́ть (успехи)", "gram": "cụm động từ", "meaning": "chứng minh thành tích học tập", "example": "Мои высокие академические успехи подтверждаются дипломом с отличием и наградами."},
        {"word": "стреми́ться (к исследованиям)", "gram": "cụm động từ", "meaning": "say mê nghiên cứu khoa học", "example": "Я стремлюсь внести личный вклад в разработку инновационных биотехнологий."},
        {"word": "реализовывать (потенциал)", "gram": "cụm động từ", "meaning": "phát huy tối đa năng lực", "example": "Грант позволит мне полностью реализовать свой исследовательский потенциал в России."}
    ],
    15: [
        {"word": "выступа́ть (в защиту)", "gram": "cụm động từ", "meaning": "lên tiếng bảo vệ di tích cổ", "example": "Общественность города открыто выступает в защиту старинной деревянной усадьбы."},
        {"word": "сохраня́ть (облик)", "gram": "cụm động từ", "meaning": "gìn giữ diện mạo lịch sử", "example": "Мы призываем сохранить уникальный исторический облик старого квартала города."},
        {"word": "реставри́ровать", "gram": "động từ chưa hoàn thành", "meaning": "trùng tu khoa học tôn tạo", "example": "Вместо сноса особняк необходимо бережно отреставрировать и открыть музей."},
        {"word": "привлека́ть (внимание)", "gram": "cụm động từ", "meaning": "thu hút dư luận quan tâm", "example": "Открытое письмо призвано привлечь внимание властей к проблеме охраны памятников."}
    ],
    16: [
        {"word": "рассужда́ть (о роли книги)", "gram": "cụm động từ", "meaning": "nghị luận giá trị sách thời số", "example": "В цифровую эпоху автор рассуждает о непреходящей ценности вдумчивого чтения."},
        {"word": "развива́ть (мышление)", "gram": "cụm động từ", "meaning": "bồi đắp tư duy phản biện", "example": "Чтение качественной литературы развивает критическое мышление и воображение."},
        {"word": "противосто́ять (шуму)", "gram": "cụm động từ", "meaning": "chống lại nhiễu loạn thông tin", "example": "Глубокая книга помогает противостоять поверхностному информационному шуму сетей."},
        {"word": "сохраня́ть (мудрость)", "gram": "cụm động từ", "meaning": "lưu giữ kho tàng tri thức", "example": "Книга веками сохраняет и передаёт потомкам духовную мудрость человечества."}
    ],
    17: [
        {"word": "знако́мить (с достопримечательностями)", "gram": "cụm động từ", "meaning": "giới thiệu danh lam thắng cảnh", "example": "Текст экскурсии увлекательно знакомит гостей с архитектурными сокровищами Москвы."},
        {"word": "проводи́ть (маршрут)", "gram": "cụm động từ", "meaning": "dẫn đường theo cung lịch sử", "example": "Гид проводит туристов по древним переулкам Замоскворечья и Красной площади."},
        {"word": "расска́зывать (о зодчестве)", "gram": "cụm động từ", "meaning": "thuyết minh kiến trúc cổ tích", "example": "Экскурсовод вдохновенно рассказывает об истории возведения храма Василия Блаженного."},
        {"word": "любова́ться (панорамой)", "gram": "cụm động từ", "meaning": "ngắm nhìn toàn cảnh thủ đô", "example": "Со смотровой площадки Воробьёвых гор туристы любуются грандиозной панорамой города."}
    ],
    18: [
        {"word": "проси́ть (об отпуске)", "gram": "cụm động từ", "meaning": "làm đơn xin bảo lưu học tập", "example": "Прошу предоставить мне академический отпуск по непредвиденным обстоятельствам."},
        {"word": "прилага́ть (справку)", "gram": "cụm động từ", "meaning": "nộp kèm chứng nhận y tế", "example": "К заявлению прилагаю заключение клинико-экспертной комиссии установленного образца."},
        {"word": "ука́зывать (причины)", "gram": "cụm động từ", "meaning": "nêu rõ lý do chính đáng", "example": "В обращении студент чётко указывает веские причины для временного перерыва в учёбе."},
        {"word": "возобновля́ть (обучение)", "gram": "cụm động từ", "meaning": "cam kết trở lại giảng đường", "example": "Обязуюсь возобновить обучение на профильном курсе с первого сентября следующего года."}
    ],
    19: [
        {"word": "осознава́ть (сопричастность)", "gram": "cụm động từ", "meaning": "ý thức trách nhiệm toàn cầu", "example": "Быть гражданином мира значит осознавать свою сопричастность к судьбе всей планеты."},
        {"word": "проявля́ть (открытость)", "gram": "cụm động từ", "meaning": "mở rộng tấm lòng đón nhận", "example": "Современный человек проявляет открытость и уважение ко всем культурам мира."},
        {"word": "реша́ть (глобальные вызовы)", "gram": "cụm động từ", "meaning": "chung tay vượt qua thách thức", "example": "Только вместе человечество способно решать климатические вызовы и побеждать бедность."},
        {"word": "способствовать (миру)", "gram": "cụm động từ", "meaning": "đóng góp cho hòa bình thế giới", "example": "Каждый из нас может способствовать укреплению дружбы и согласия между народами."}
    ],
    20: [
        {"word": "рекомендова́ть (кандидата)", "gram": "cụm động từ", "meaning": "viết thư giới thiệu học trò", "example": "С большим удовольствием рекомендую своего студента для поступления в аспирантуру."},
        {"word": "характеризова́ть", "gram": "động từ chưa hoàn thành", "meaning": "đánh giá phẩm chất đạo đức", "example": "Профессор характеризует соискателя как исключительно прилежного исследователя."},
        {"word": "отмеча́ть (потенциал)", "gram": "cụm động từ", "meaning": "khen ngợi tiềm năng khoa học", "example": "Особо отмечаю его самостоятельность, научную честность и высокий творческий потенциал."},
        {"word": "руча́ться (за успехи)", "gram": "cụm động từ", "meaning": "bảo lãnh cho tương lai", "example": "Уверенно ручаюсь за то, что студент блестяще справится с исследовательской программой."}
    ],
    21: [
        {"word": "анализи́ровать (рынок труда)", "gram": "cụm động từ", "meaning": "nghiên cứu thị trường việc làm", "example": "В аналитической записке подробно анализируются тенденции современного рынка труда."},
        {"word": "прогнози́ровать (спрос)", "gram": "cụm động từ", "meaning": "dự báo nhu cầu nhân lực", "example": "Эксперты прогнозируют взрывной рост спроса на специалистов в сфере биотехнологий."},
        {"word": "выявля́ть (тенденции)", "gram": "cụm động từ", "meaning": "chỉ ra xu thế chuyển dịch", "example": "Исследование выявляет высокую популярность удалённого и проектного форматов занятости."},
        {"word": "рекомендова́ть (навыки)", "gram": "cụm động từ", "meaning": "khuyên trau dồi chuyên môn", "example": "Выпускникам рекомендуется развивать смежные компетенции для повышения конкурентоспособности."}
    ],
    22: [
        {"word": "размышля́ть (о технологиях)", "gram": "cụm động từ", "meaning": "nghị luận công nghệ số", "example": "В эссе автор размышляет над дилеммой: технологии объединяют или разобщают людей?"},
        {"word": "создава́ть (иллюзию)", "gram": "cụm động từ", "meaning": "tạo ảo giác kết nối", "example": "Гаджеты создают иллюзию близости, но нередко ведут к эмоциональному одиночеству."},
        {"word": "укрепля́ть (солидарность)", "gram": "cụm động từ", "meaning": "kết nối người cùng chí hướng", "example": "С другой стороны, сетевые сообщества объединяют единомышленников по всей планете."},
        {"word": "находи́ть (баланс)", "gram": "cụm động từ", "meaning": "cân bằng thế giới thực và ảo", "example": "Человек должен находить баланс, отдавая приоритет тёплому живому общению."}
    ],
    23: [
        {"word": "анонси́ровать (акцию)", "gram": "cụm động từ", "meaning": "thông báo sự kiện thiện nguyện", "example": "В пресс-релизе фонд анонсирует проведение масштабной благотворительной культурной акции."},
        {"word": "собира́ть (средства)", "gram": "cụm động từ", "meaning": "gây quỹ bảo tồn thư viện", "example": "Цель акции — собрать необходимые средства на реставрацию старинной сельской библиотеки."},
        {"word": "привлека́ть (спонсоров)", "gram": "cụm động từ", "meaning": "kêu gọi nhà tài trợ", "example": "Организаторам удалось привлечь ведущих меценатов и волонтёрские организации города."},
        {"word": "призыва́ть (к участию)", "gram": "cụm động từ", "meaning": "vận động cộng đồng hưởng ứng", "example": "Пресс-служба призывает всех неравнодушных граждан принять участие в добром деле."}
    ],
    24: [
        {"word": "направля́ть (приглашение)", "gram": "cụm động từ", "meaning": "gửi thư mời danh dự", "example": "Ректорат направляет официальное приглашение профессору для чтения открытого лекционного курса."},
        {"word": "предлага́ть (программу)", "gram": "cụm động từ", "meaning": "đề xuất lịch trình làm việc", "example": "Мы с радостью предлагаем насыщенную академическую программу трёхдневного визита."},
        {"word": "обеспе́чивать (приём)", "gram": "cụm động từ", "meaning": "bảo đảm điều kiện đón tiếp", "example": "Университет берёт на себя все расходы по проживанию и комфортному пребыванию учёного."},
        {"word": "наде́яться (на визит)", "gram": "cụm động từ", "meaning": "mong đợi sự có mặt của giáo sư", "example": "Искренне надеемся на ваш визит, который укрепит научное партнёрство наших институтов."}
    ],
    25: [
        {"word": "демонстри́ровать (речь)", "gram": "cụm động từ", "meaning": "thể hiện văn phong học thuật", "example": "В итоговом эссе кандидат демонстрирует богатство лексики и зрелость суждений."},
        {"word": "аргументи́ровать (тезис)", "gram": "cụm động từ", "meaning": "lập luận thuyết phục chặt chẽ", "example": "Каждое положение работы автор логично аргументирует примерами из классики и жизни."},
        {"word": "соблюда́ть (нормы)", "gram": "cụm động từ", "meaning": "tuân thủ chuẩn ngữ pháp B1", "example": "Текст написан с безупречным соблюдением всех норм грамматики и стилистики русского языка."},
        {"word": "формули́ровать (вывод)", "gram": "cụm động từ", "meaning": "khái quát kết luận sâu sắc", "example": "В заключительной части автор глубоко и лаконично формулирует главный вывод работы."}
    ]
}

writing_b1 = {}
for i in range(1, 26):
    noun_cat = b1_old.B1_WRITING_VOCAB[i][0]
    noun_cat["title"] = "Danh từ trong bài viết (Существительные)"
    verb_cat = {
        "title": "Động từ then chốt (Глаголы)",
        "rows": writing_verbs_b1[i]
    }
    writing_b1[i] = [noun_cat, verb_cat]

b1_master = {
    'reading': reading_b1,
    'listening': listening_b1,
    'speaking': speaking_b1,
    'writing': writing_b1
}

with open('scripts/vocab_master_b1.json', 'w', encoding='utf-8') as f:
    json.dump(b1_master, f, ensure_ascii=False, indent=2)

print("SUCCESS: scripts/vocab_master_b1.json built successfully!")
print("Counts: Reading:", len(reading_b1), "Listening:", len(listening_b1), "Speaking:", len(speaking_b1), "Writing:", len(writing_b1))
