# -*- coding: utf-8 -*-
"""
Build complete 100 lessons for A1 (Reading, Listening, Speaking, Writing).
Each lesson gets:
- Category 1: Danh từ trong bài học (Существительные) - 4 words
- Category 2: Động từ then chốt (Глаголы) - 4 words
Total 8 words per lesson = 800 words for A1.
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

with open('scripts/reading_a1_vocab.json', 'r', encoding='utf-8') as f:
    reading_a1 = json.load(f)

# Convert string keys to int
reading_a1 = {int(k): v for k, v in reading_a1.items()}

# Listening A1 verbs mapping
listening_verbs = {
    1: [
        {"word": "зака́зывать", "gram": "động từ chưa hoàn thành", "meaning": "đặt món, gọi món", "example": "Что вы будете заказывать сегодня?"},
        {"word": "принести́", "gram": "động từ hoàn thành", "meaning": "mang đến", "example": "Принесите, пожалуйста, чашку чая."},
        {"word": "пить", "gram": "động từ chưa hoàn thành", "meaning": "uống", "example": "Максим пьёт чёрный кофе без сахара."},
        {"word": "выбира́ть", "gram": "động từ chưa hoàn thành", "meaning": "lựa chọn", "example": "Клиент долго выбирает десерт."}
    ],
    2: [
        {"word": "знако́миться", "gram": "động từ chưa hoàn thành", "meaning": "làm quen", "example": "Давайте познакомимся поближе."},
        {"word": "звать", "gram": "động từ chưa hoàn thành", "meaning": "gọi tên", "example": "Как вас зовут? — Меня зовут Анна."},
        {"word": "прие́хать", "gram": "động từ hoàn thành", "meaning": "đến nơi (xe/tàu)", "example": "Откуда вы приехали в Москву?"},
        {"word": "понима́ть", "gram": "động từ chưa hoàn thành", "meaning": "thấu hiểu", "example": "Я понимаю русскую речь."}
    ],
    3: [
        {"word": "чита́ть", "gram": "động từ chưa hoàn thành", "meaning": "đọc", "example": "Студенты читают научные статьи."},
        {"word": "брать", "gram": "động từ chưa hoàn thành", "meaning": "mượn, lấy", "example": "Можно взять этот словарь домой?"},
        {"word": "иска́ть", "gram": "động từ chưa hoàn thành", "meaning": "tìm kiếm", "example": "Я ищу учебник по грамматике."},
        {"word": "возвраща́ть", "gram": "động từ chưa hoàn thành", "meaning": "trả lại", "example": "Книги нужно возвращать вовремя."}
    ],
    4: [
        {"word": "ме́рить", "gram": "động từ chưa hoàn thành", "meaning": "thử đồ (quần áo)", "example": "Где можно примерить эту куртку?"},
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua sắm", "example": "Анна покупает тёплый свитер."},
        {"word": "подходи́ть", "gram": "động từ chưa hoàn thành", "meaning": "vừa vặn, hợp", "example": "Этот размер мне идеально подходит."},
        {"word": "сто́ить", "gram": "động từ chưa hoàn thành", "meaning": "có giá bán", "example": "Сколько стоит эта зимняя куртка?"}
    ],
    5: [
        {"word": "боле́ть", "gram": "động từ chưa hoàn thành", "meaning": "bị đau, bị ốm", "example": "У меня сильно болит горло."},
        {"word": "принима́ть", "gram": "động từ chưa hoàn thành", "meaning": "uống (thuốc)", "example": "Принимайте эти таблетки после еды."},
        {"word": "сове́товать", "gram": "động từ chưa hoàn thành", "meaning": "khuyên bảo", "example": "Фармацевт советует хорошее лекарство."},
        {"word": "выздора́вливать", "gram": "động từ chưa hoàn thành", "meaning": "bình phục", "example": "Выздоравливайте скорее!"}
    ],
    6: [
        {"word": "просыпа́ться", "gram": "động từ chưa hoàn thành", "meaning": "thức giấc", "example": "Я просыпаюсь в семь часов утра."},
        {"word": "за́втракать", "gram": "động từ chưa hoàn thành", "meaning": "ăn sáng", "example": "Мы быстро завтракаем на кухне."},
        {"word": "е́хать", "gram": "động từ chưa hoàn thành", "meaning": "đi (bằng xe)", "example": "Утром я еду на работу на метро."},
        {"word": "отдыха́ть", "gram": "động từ chưa hoàn thành", "meaning": "nghỉ ngơi", "example": "Вечером я отдыхаю и смотрю новости."}
    ],
    7: [
        {"word": "прибыва́ть", "gram": "động từ chưa hoàn thành", "meaning": "đến ga", "example": "Поезд прибывает на первый путь."},
        {"word": "отправля́ться", "gram": "động từ chưa hoàn thành", "meaning": "khởi hành", "example": "Скорый поезд отправляется через пять минут."},
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua", "example": "Пассажир покупает билет в кассе."},
        {"word": "ждать", "gram": "động từ chưa hoàn thành", "meaning": "chờ đợi", "example": "Люди ждут посадки на платформе."}
    ],
    8: [
        {"word": "брони́ровать", "gram": "động từ chưa hoàn thành", "meaning": "đặt phòng trước", "example": "Я хочу забронировать одноместный номер."},
        {"word": "жить", "gram": "động từ chưa hoàn thành", "meaning": "ở trọ, cư trú", "example": "Туристы живут в гостинице в центре."},
        {"word": "плати́ть", "gram": "động từ chưa hoàn thành", "meaning": "trả tiền", "example": "Можно заплатить кредитной картой?"},
        {"word": "выезжа́ть", "gram": "động từ chưa hoàn thành", "meaning": "trả phòng rời đi", "example": "Мы выезжаем завтра утром."}
    ],
    9: [
        {"word": "идти́", "gram": "động từ chưa hoàn thành", "meaning": "đi bộ", "example": "Идите прямо до перекрёстка."},
        {"word": "повора́чивать", "gram": "động từ chưa hoàn thành", "meaning": "rẽ hướng", "example": "Поверните направо около аптеки."},
        {"word": "спра́шивать", "gram": "động từ chưa hoàn thành", "meaning": "hỏi thăm", "example": "Турист спрашивает дорогу к музею."},
        {"word": "подска́зывать", "gram": "động từ chưa hoàn thành", "meaning": "chỉ bảo, hướng dẫn", "example": "Подскажите, пожалуйста, где остановка?"}
    ],
    10: [
        {"word": "лете́ть", "gram": "động từ chưa hoàn thành", "meaning": "bay trên không", "example": "Самолёт летит в Санкт-Петербург."},
        {"word": "регистри́роваться", "gram": "động từ chưa hoàn thành", "meaning": "làm thủ tục chuyến bay", "example": "Пассажиры регистрируются на рейс."},
        {"word": "сдава́ть", "gram": "động từ chưa hoàn thành", "meaning": "ký gửi (hành lý)", "example": "Я сдаю чемодан в багаж."},
        {"word": "сади́ться", "gram": "động từ chưa hoàn thành", "meaning": "lên (máy bay/xe)", "example": "Объявлена посадка на рейс номер сто."}
    ],
    11: [
        {"word": "осма́тривать", "gram": "động từ chưa hoàn thành", "meaning": "thăm khám bệnh", "example": "Врач внимательно осматривает пациента."},
        {"word": "лечи́ть", "gram": "động từ chưa hoàn thành", "meaning": "điều trị chữa bệnh", "example": "Доктор лечит простуду и грипп."},
        {"word": "жа́ловаться", "gram": "động từ chưa hoàn thành", "meaning": "than phiền kêu đau", "example": "Пациент жалуется на сильный кашель."},
        {"word": "выпи́сывать", "gram": "động từ chưa hoàn thành", "meaning": "kê đơn thuốc", "example": "Врач выписывает эффективный сироп."}
    ],
    12: [
        {"word": "выбира́ть", "gram": "động từ chưa hoàn thành", "meaning": "lựa chọn", "example": "Туристы выбирают сувениры на память."},
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua sắm", "example": "Я покупаю деревянную матрёшку."},
        {"word": "дари́ть", "gram": "động từ chưa hoàn thành", "meaning": "tặng quà", "example": "Этот сувенир я подарю другу."},
        {"word": "нра́виться", "gram": "động từ chưa hoàn thành", "meaning": "làm cho thích", "example": "Мне очень нравится эта шкатулка."}
    ],
    13: [
        {"word": "гото́вить", "gram": "động từ chưa hoàn thành", "meaning": "nấu ăn", "example": "Мама готовит праздничный ужин."},
        {"word": "есть", "gram": "động từ chưa hoàn thành", "meaning": "ăn uống", "example": "Вся семья с аппетитом ест горячий суп."},
        {"word": "угоща́ть", "gram": "động từ chưa hoàn thành", "meaning": "mời ăn, thết đãi", "example": "Хозяйка угощает гостей пирогами."},
        {"word": "благодари́ть", "gram": "động từ chưa hoàn thành", "meaning": "cảm ơn", "example": "Гости благодарят за вкусный ужин."}
    ],
    14: [
        {"word": "свети́ть", "gram": "động từ chưa hoàn thành", "meaning": "chiếu sáng", "example": "На небе ярко светит солнце."},
        {"word": "идти́ (дождь)", "gram": "cụm động từ", "meaning": "trời mưa", "example": "На улице идёт сильный дождь."},
        {"word": "тепле́ть", "gram": "động từ chưa hoàn thành", "meaning": "ấm áp dần lên", "example": "Весной погода быстро теплеет."},
        {"word": "холода́ть", "gram": "động từ chưa hoàn thành", "meaning": "lạnh dần đi", "example": "К вечеру на улице холодает."}
    ],
    15: [
        {"word": "меня́ть", "gram": "động từ chưa hoàn thành", "meaning": "đổi (tiền tệ)", "example": "Я хочу поменять доллары на рубли."},
        {"word": "снима́ть", "gram": "động từ chưa hoàn thành", "meaning": "rút tiền mặt", "example": "Он снимает наличные деньги в банкомате."},
        {"word": "открыва́ть", "gram": "động từ chưa hoàn thành", "meaning": "mở tài khoản", "example": "Студент открывает счёт в банке."},
        {"word": "подпи́сывать", "gram": "động từ chưa hoàn thành", "meaning": "ký giấy tờ", "example": "Подпишите квитанцию здесь, пожалуйста."}
    ],
    16: [
        {"word": "осма́тривать", "gram": "động từ chưa hoàn thành", "meaning": "tham quan ngắm nhìn", "example": "Туристы осматривают шедевры живописи."},
        {"word": "слу́шать", "gram": "động từ chưa hoàn thành", "meaning": "lắng nghe", "example": "Экскурсанты слушают рассказ гида."},
        {"word": "расска́зывать", "gram": "động từ chưa hoàn thành", "meaning": "thuyết minh", "example": "Экскурсовод рассказывает историю музея."},
        {"word": "восхища́ться", "gram": "động từ chưa hoàn thành", "meaning": "trầm trồ thán phục", "example": "Все восхищаются старинными залами."}
    ],
    17: [
        {"word": "игра́ть", "gram": "động từ chưa hoàn thành", "meaning": "chơi (thể thao)", "example": "Футболисты играют решающий матч."},
        {"word": "боле́ть", "gram": "động từ chưa hoàn thành", "meaning": "cổ vũ đội bóng", "example": "Мы болеем за команду «Спартак»."},
        {"word": "выи́грывать", "gram": "động từ chưa hoàn thành", "meaning": "chiến thắng", "example": "Наша команда выигрывает со счётом 2:0."},
        {"word": "тренирова́ться", "gram": "động từ chưa hoàn thành", "meaning": "tập luyện", "example": "Спортсмены тренируются каждый день."}
    ],
    18: [
        {"word": "выбира́ть", "gram": "động từ chưa hoàn thành", "meaning": "chọn mua", "example": "Покупатель выбирает новый смартфон."},
        {"word": "проверя́ть", "gram": "động từ chưa hoàn thành", "meaning": "kiểm tra tính năng", "example": "Продавец проверяет работу экрана."},
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua hàng", "example": "Я покупаю современный ноутбук для учёбы."},
        {"word": "рабо́тать", "gram": "động từ chưa hoàn thành", "meaning": "hoạt động tốt", "example": "Новый телефон работает быстро."}
    ],
    19: [
        {"word": "взве́шивать", "gram": "động từ chưa hoàn thành", "meaning": "cân đo", "example": "Продавец взвешивает два килограмма яблок."},
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua", "example": "Бабушка покупает свежие овощи на рынке."},
        {"word": "сто́ить", "gram": "động từ chưa hoàn thành", "meaning": "có giá bán", "example": "Сколько стоят эти сладкие помидоры?"},
        {"word": "торгова́ться", "gram": "động từ chưa hoàn thành", "meaning": "mặc cả giá", "example": "На восточном рынке принято торговаться."}
    ],
    20: [
        {"word": "сдава́ть", "gram": "động từ chưa hoàn thành", "meaning": "dự thi, thi cử", "example": "Студент сдаёт экзамен по русскому языку."},
        {"word": "спра́шивать", "gram": "động từ chưa hoàn thành", "meaning": "đặt câu hỏi", "example": "Экзаменатор спрашивает о граматических правилах."},
        {"word": "отвеча́ть", "gram": "động từ chưa hoàn thành", "meaning": "trả lời rành mạch", "example": "Антон уверенно отвечает на все вопросы."},
        {"word": "объясня́ть", "gram": "động từ chưa hoàn thành", "meaning": "giải thích", "example": "Преподаватель объясняет сложный пример."}
    ],
    21: [
        {"word": "выступа́ть", "gram": "động từ chưa hoàn thành", "meaning": "biểu diễn sân khấu", "example": "Актёры выступают на сцене театра."},
        {"word": "смотре́ть", "gram": "động từ chưa hoàn thành", "meaning": "xem kịch", "example": "Зрители смотрят классический спектакль."},
        {"word": "хло́пать", "gram": "động từ chưa hoàn thành", "meaning": "vỗ tay tán thưởng", "example": "Зал громко хлопает любимым артистам."},
        {"word": "смея́ться", "gram": "động từ chưa hoàn thành", "meaning": "cười vui", "example": "Дети весело смеются во время представления."}
    ],
    22: [
        {"word": "печь", "gram": "động từ chưa hoàn thành", "meaning": "nướng, rán bánh", "example": "Бабушка печёт тонкие русские блины."},
        {"word": "про́бовать", "gram": "động từ chưa hoàn thành", "meaning": "nếm thử", "example": "Попробуйте горячий блин со сгущёнкой!"},
        {"word": "зака́зывать", "gram": "động từ chưa hoàn thành", "meaning": "gọi món", "example": "Мы заказываем блины с красной икрой."},
        {"word": "угоща́ть", "gram": "động từ chưa hoàn thành", "meaning": "chiêu đãi", "example": "Друзья угощают друг друга на Масленицу."}
    ],
    23: [
        {"word": "вызыва́ть", "gram": "động từ chưa hoàn thành", "meaning": "gọi xe (qua app)", "example": "Я вызываю такси через приложение в телефоне."},
        {"word": "е́хать", "gram": "động từ chưa hoàn thành", "meaning": "đi xe", "example": "Мы едем в аэропорт без пробок."},
        {"word": "плати́ть", "gram": "động từ chưa hoàn thành", "meaning": "trả tiền cước", "example": "Пассажир платит за поездку картой."},
        {"word": "остана́вливаться", "gram": "động từ chưa hoàn thành", "meaning": "dừng xe", "example": "Водитель останавливается около входа."}
    ],
    24: [
        {"word": "гуля́ть", "gram": "động từ chưa hoàn thành", "meaning": "đi dạo tản bộ", "example": "Вечером мы гуляем по тихим аллеям парка."},
        {"word": "сиде́ть", "gram": "động từ chưa hoàn thành", "meaning": "ngồi nghỉ", "example": "Старики сидят на деревянных скамейках."},
        {"word": "дыша́ть", "gram": "động từ chưa hoàn thành", "meaning": "hít thở", "example": "Приятно дышать свежим лесным воздухом."},
        {"word": "фотографи́ровать", "gram": "động từ chưa hoàn thành", "meaning": "chụp ảnh", "example": "Туристы фотографируют осенний пейзаж."}
    ],
    25: [
        {"word": "приглаша́ть", "gram": "động từ chưa hoàn thành", "meaning": "mời tham dự", "example": "Друг приглашает меня на свой день рождения."},
        {"word": "пра́здновать", "gram": "động từ chưa hoàn thành", "meaning": "ăn mừng tổ chức tiệc", "example": "Мы празднуем окончание курса А1."},
        {"word": "приходи́ть", "gram": "động từ chưa hoàn thành", "meaning": "đến dự", "example": "Гости приходят точно в назначенное время."},
        {"word": "поздравля́ть", "gram": "động từ chưa hoàn thành", "meaning": "chúc mừng", "example": "Все поздравляют именинника с праздником."}
    ]
}

# Load existing A1 listening, speaking, writing to keep their nouns and add the verbs!
import scripts.vocab_data_a1 as a1_old

listening_a1 = {}
for i in range(1, 26):
    noun_cat = a1_old.A1_LISTENING_VOCAB[i][0]
    noun_cat["title"] = "Danh từ trong bài nghe (Существительные)"
    verb_cat = {
        "title": "Động từ then chốt (Глаголы)",
        "rows": listening_verbs[i]
    }
    listening_a1[i] = [noun_cat, verb_cat]

# Speaking A1 verbs mapping
speaking_verbs = {
    1: [
        {"word": "звать", "gram": "động từ chưa hoàn thành", "meaning": "gọi tên", "example": "Как вас зовут? — Меня зовут Минх."},
        {"word": "жить", "gram": "động từ chưa hoàn thành", "meaning": "sống, cư ngụ", "example": "Я живу в красивом городе Дананг."},
        {"word": "рабо́тать", "gram": "động từ chưa hoàn thành", "meaning": "làm việc", "example": "Мой отец работает инженером."},
        {"word": "учи́ться", "gram": "động từ chưa hoàn thành", "meaning": "học tập", "example": "Я учусь в политехническом университете."}
    ],
    2: [
        {"word": "ви́деть", "gram": "động từ chưa hoàn thành", "meaning": "nhìn thấy", "example": "Что вы видите на этой картине?"},
        {"word": "называ́ть", "gram": "động từ chưa hoàn thành", "meaning": "gọi tên, nêu danh", "example": "Назовите предметы в вашей комнате."},
        {"word": "пока́зывать", "gram": "động từ chưa hoàn thành", "meaning": "chỉ ra, cho xem", "example": "Покажите, пожалуйста, этот словарь."},
        {"word": "знать", "gram": "động từ chưa hoàn thành", "meaning": "biết", "example": "Я знаю, как называется этот предмет."}
    ],
    3: [
        {"word": "люби́ть", "gram": "động từ chưa hoàn thành", "meaning": "yêu quý", "example": "Я очень люблю свою семью."},
        {"word": "расска́зывать", "gram": "động từ chưa hoàn thành", "meaning": "kể chuyện", "example": "Расскажите о ваших родителях."},
        {"word": "помога́ть", "gram": "động từ chưa hoàn thành", "meaning": "giúp đỡ", "example": "Дети помогают родителям по дому."},
        {"word": "звони́ть", "gram": "động từ chưa hoàn thành", "meaning": "gọi điện thoại", "example": "Каждое воскресенье я звоню маме."}
    ],
    4: [
        {"word": "жить", "gram": "động từ chưa hoàn thành", "meaning": "ở trọ, cư ngụ", "example": "Иностранные студенты живут в общежитии."},
        {"word": "дели́ть", "gram": "động từ chưa hoàn thành", "meaning": "chia sẻ chung", "example": "Мы делим комнату с соседом из Китая."},
        {"word": "гото́вить", "gram": "động từ chưa hoàn thành", "meaning": "nấu nướng", "example": "На общей кухне мы готовим ужин."},
        {"word": "убира́ть", "gram": "động từ chưa hoàn thành", "meaning": "dọn dẹp", "example": "По субботам мы убираем комнату."}
    ],
    5: [
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua sắm", "example": "Я часто покупаю продукты в супермаркете."},
        {"word": "сто́ить", "gram": "động từ chưa hoàn thành", "meaning": "có giá bán", "example": "Сколько стоит один килограмм яблок?"},
        {"word": "плати́ть", "gram": "động từ chưa hoàn thành", "meaning": "trả tiền", "example": "Я плачу за покупки наличными деньгами."},
        {"word": "эконо́мить", "gram": "động từ chưa hoàn thành", "meaning": "tiết kiệm", "example": "Скидки помогают экономить бюджет."}
    ],
    6: [
        {"word": "спеши́ть", "gram": "động từ chưa hoàn thành", "meaning": "vội vã", "example": "Я спешу на занятия в университет."},
        {"word": "начина́ться", "gram": "động từ chưa hoàn thành", "meaning": "bắt đầu", "example": "Первая пара начинается ровно в восемь утра."},
        {"word": "зака́нчиваться", "gram": "động từ chưa hoàn thành", "meaning": "kết thúc", "example": "Уроки заканчиваются в два часа дня."},
        {"word": "опа́здывать", "gram": "động từ chưa hoàn thành", "meaning": "muộn, trễ giờ", "example": "Студенты никогда не опаздывают на лекции."}
    ],
    7: [
        {"word": "гото́вить", "gram": "động từ chưa hoàn thành", "meaning": "nấu ăn", "example": "Что вы любите готовить дома?"},
        {"word": "про́бовать", "gram": "động từ chưa hoàn thành", "meaning": "nếm thử món", "example": "Я впервые пробую горячий борщ."},
        {"word": "предпочита́ть", "gram": "động từ chưa hoàn thành", "meaning": "chuộng, thích hơn", "example": "Я предпочитаю зелёный чай кофе."},
        {"word": "зака́зывать", "gram": "động từ chưa hoàn thành", "meaning": "gọi món", "example": "Мы заказываем суп и салат на обед."}
    ],
    8: [
        {"word": "идти́", "gram": "động từ chưa hoàn thành", "meaning": "đi bộ", "example": "Идите прямо по этой широкой улице."},
        {"word": "е́хать", "gram": "động từ chưa hoàn thành", "meaning": "đi phương tiện", "example": "В центр города лучше ехать на метро."},
        {"word": "повора́чивать", "gram": "động từ chưa hoàn thành", "meaning": "rẽ hướng", "example": "Около музея поверните налево."},
        {"word": "находи́ть", "gram": "động từ chưa hoàn thành", "meaning": "tìm thấy", "example": "С картой легко найти нужный адрес."}
    ],
    9: [
        {"word": "встава́ть", "gram": "động từ chưa hoàn thành", "meaning": "thức dậy", "example": "Я встаю рано и делаю зарядку."},
        {"word": "занима́ться", "gram": "động từ chưa hoàn thành", "meaning": "học bài, ôn tập", "example": "Студенты занимаются в библиотеке."},
        {"word": "обе́дать", "gram": "động từ chưa hoàn thành", "meaning": "ăn trưa", "example": "Мы обедаем в студенческой столовой."},
        {"word": "ложи́ться (спать)", "gram": "cụm động từ", "meaning": "đi ngủ", "example": "Я ложусь спать в одиннадцать часов вечера."}
    ],
    10: [
        {"word": "увлека́ться", "gram": "động từ chưa hoàn thành", "meaning": "say mê, đam mê", "example": "Чем вы увлекаетесь в свободное время?"},
        {"word": "игра́ть", "gram": "động từ chưa hoàn thành", "meaning": "chơi nhạc/thể thao", "example": "Мой друг отлично играет в шахматы."},
        {"word": "чита́ть", "gram": "động từ chưa hoàn thành", "meaning": "đọc sách", "example": "Я люблю читать исторические романы."},
        {"word": "рисова́ть", "gram": "động từ chưa hoàn thành", "meaning": "vẽ tranh", "example": "Сестра красиво рисует акварелью."}
    ],
    11: [
        {"word": "боле́ть", "gram": "động từ chưa hoàn thành", "meaning": "bị ốm, đau", "example": "Сегодня я не могу прийти, я болею."},
        {"word": "принима́ть", "gram": "động từ chưa hoàn thành", "meaning": "uống thuốc", "example": "Нужно принимать витамины три раза в день."},
        {"word": "чу́вствовать", "gram": "động từ chưa hoàn thành", "meaning": "cảm thấy", "example": "Как вы себя чувствуете сегодня?"},
        {"word": "выздора́вливать", "gram": "động từ chưa hoàn thành", "meaning": "bình phục", "example": "Я уже чувствую себя лучше и выздоравливаю."}
    ],
    12: [
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua vé", "example": "Один билет до Санкт-Петербурга, пожалуйста."},
        {"word": "е́хать", "gram": "động từ chưa hoàn thành", "meaning": "đi tàu", "example": "Мы едем в купейном вагоне."},
        {"word": "отправля́ться", "gram": "động từ chưa hoàn thành", "meaning": "khởi hành", "example": "Поезд отправляется точно по расписанию."},
        {"word": "плати́ть", "gram": "động từ chưa hoàn thành", "meaning": "thanh toán", "example": "Я плачу за билет через банковское приложение."}
    ],
    13: [
        {"word": "одева́ться", "gram": "động từ chưa hoàn thành", "meaning": "mặc quần áo", "example": "Зимой нужно одеваться очень тепло."},
        {"word": "мёрзнуть", "gram": "động từ chưa hoàn thành", "meaning": "bị cóng, lạnh rét", "example": "Без тёплых перчаток быстро мёрзнут руки."},
        {"word": "носи́ть", "gram": "động từ chưa hoàn thành", "meaning": "mặc, mang theo", "example": "В холодные дни я ношу шерстяной шарф."},
        {"word": "гуля́ть", "gram": "động từ chưa hoàn thành", "meaning": "đi dạo", "example": "В солнечную погоду приятно гулять в парке."}
    ],
    14: [
        {"word": "брони́ровать", "gram": "động từ chưa hoàn thành", "meaning": "đặt bàn trước", "example": "Мы хотим забронировать столик на четверых."},
        {"word": "зака́зывать", "gram": "động từ chưa hoàn thành", "meaning": "gọi món", "example": "Что вы посоветуете заказать на горячее?"},
        {"word": "про́бовать", "gram": "động từ chưa hoàn thành", "meaning": "nếm thử ẩm thực", "example": "Обязательно попробуйте фирменный пирог."},
        {"word": "благодари́ть", "gram": "động từ chưa hoàn thành", "meaning": "cảm ơn phục vụ", "example": "Клиенты благодарят официанта за отличный сервис."}
    ],
    15: [
        {"word": "смотре́ть", "gram": "động từ chưa hoàn thành", "meaning": "xem phim chiếu rạp", "example": "Мы смотрим премьеру нового фильма."},
        {"word": "выбира́ть", "gram": "động từ chưa hoàn thành", "meaning": "chọn chỗ ngồi", "example": "Давай выберем места в центре зала."},
        {"word": "идти́", "gram": "động từ chưa hoàn thành", "meaning": "đi xem phim", "example": "В субботу мы идём в кинотеатр."},
        {"word": "обсужда́ть", "gram": "động từ chưa hoàn thành", "meaning": "bình luận thảo luận", "example": "После сеанса зрители обсуждают финал картины."}
    ],
    16: [
        {"word": "встреча́ться", "gram": "động từ chưa hoàn thành", "meaning": "gặp gỡ bạn bè", "example": "Мы встречаемся в уютном кафе в шесть вечера."},
        {"word": "обща́ться", "gram": "động từ chưa hoàn thành", "meaning": "trò chuyện tâm sự", "example": "Приятно общаться с близкими друзьями."},
        {"word": "расска́зывать", "gram": "động từ chưa hoàn thành", "meaning": "kể chuyện vui", "example": "Друг рассказывает смешную историю из жизни."},
        {"word": "смея́ться", "gram": "động từ chưa hoàn thành", "meaning": "cười rạng rỡ", "example": "Все весело смеются и пьют чай."}
    ],
    17: [
        {"word": "путеше́ствовать", "gram": "động từ chưa hoàn thành", "meaning": "đi du lịch", "example": "Я мечтаю путешествовать по городам Золотого кольца."},
        {"word": "посеща́ть", "gram": "động từ chưa hoàn thành", "meaning": "thăm quan di tích", "example": "Туристы посещают старинные кремли и храмы."},
        {"word": "осма́тривать", "gram": "động từ chưa hoàn thành", "meaning": "ngắm nhìn danh lam", "example": "Мы осматриваем исторический центр города."},
        {"word": "фотографи́ровать", "gram": "động từ chưa hoàn thành", "meaning": "chụp ảnh lưu niệm", "example": "Я фотографирую красивые памятники архитектуры."}
    ],
    18: [
        {"word": "корми́ть", "gram": "động từ chưa hoàn thành", "meaning": "cho ăn", "example": "Утром я кормлю своего любимого кота."},
        {"word": "гуля́ть", "gram": "động từ chưa hoàn thành", "meaning": "dắt đi dạo", "example": "Брат гуляет с собакой в парке каждый день."},
        {"word": "игра́ть", "gram": "động từ chưa hoàn thành", "meaning": "chơi đùa", "example": "Щенок любит играть с резиновым мячиком."},
        {"word": "забо́титься", "gram": "động từ chưa hoàn thành", "meaning": "chăm sóc yêu thương", "example": "Мы с любовью заботимся о домашних животных."}
    ],
    19: [
        {"word": "плани́ровать", "gram": "động từ chưa hoàn thành", "meaning": "lên kế hoạch", "example": "Что вы планируете делать на выходных?"},
        {"word": "е́хать", "gram": "động từ chưa hoàn thành", "meaning": "đi dã ngoại", "example": "В субботу мы едем на загородную дачу."},
        {"word": "отдыха́ть", "gram": "động từ chưa hoàn thành", "meaning": "nghỉ ngơi thư giãn", "example": "Я хочу отдохнуть от шумного города на природе."},
        {"word": "собира́ться", "gram": "động từ chưa hoàn thành", "meaning": "chuẩn bị, tụ họp", "example": "Друзья собираются поехать на озеро на пикник."}
    ],
    20: [
        {"word": "говори́ть", "gram": "động từ chưa hoàn thành", "meaning": "nói lưu loát", "example": "Теперь я свободно говорю на простые темы по-русски."},
        {"word": "расска́зывать", "gram": "động từ chưa hoàn thành", "meaning": "thuyết trình kể", "example": "Студент рассказывает о себе и своих планах."},
        {"word": "дели́ться", "gram": "động từ chưa hoàn thành", "meaning": "chia sẻ cảm xúc", "example": "Я хочу поделиться радостью моих успехов в учёбе."},
        {"word": "ра́доваться", "gram": "động từ chưa hoàn thành", "meaning": "vui mừng hoan hỉ", "example": "Мы радуемся отличным результатам курса."}
    ],
    21: [
        {"word": "бе́гать", "gram": "động từ chưa hoàn thành", "meaning": "chạy bộ rèn sức", "example": "По утрам я бегаю на школьном стадионе."},
        {"word": "пла́вать", "gram": "động từ chưa hoàn thành", "meaning": "bơi lội", "example": "Два раза в неделю я плаваю в бассейне."},
        {"word": "тренирова́ться", "gram": "động từ chưa hoàn thành", "meaning": "tập luyện thể lực", "example": "Спортсмены упорно тренируются в зале."},
        {"word": "укрепля́ть", "gram": "động từ chưa hoàn thành", "meaning": "nâng cao sức khỏe", "example": "Регулярный спорт укрепляет иммунитет."}
    ],
    22: [
        {"word": "поздравля́ть", "gram": "động từ chưa hoàn thành", "meaning": "chúc mừng", "example": "Поздравляю вас с Новым годом и Рождеством!"},
        {"word": "пра́здновать", "gram": "động từ chưa hoàn thành", "meaning": "tổ chức liên hoan", "example": "Мы празднуем день рождения в кругу друзей."},
        {"word": "жела́ть", "gram": "động từ chưa hoàn thành", "meaning": "chúc điều may mắn", "example": "Желаю вам крепкого здоровья и больших успехов!"},
        {"word": "дари́ть", "gram": "động từ chưa hoàn thành", "meaning": "trao tặng quà", "example": "Родители дарят детям красивые книги и сладости."}
    ],
    23: [
        {"word": "чита́ть", "gram": "động từ chưa hoàn thành", "meaning": "nghiên cứu đọc", "example": "В читальном зале студенты читают редкие книги."},
        {"word": "учи́ть", "gram": "động từ chưa hoàn thành", "meaning": "học thuộc lòng", "example": "Я учу новые русские слова каждый день."},
        {"word": "конспекти́ровать", "gram": "động từ chưa hoàn thành", "meaning": "ghi chép tóm tắt", "example": "Она аккуратно конспектирует лекции профессора."},
        {"word": "повторя́ть", "gram": "động từ chưa hoàn thành", "meaning": "ôn lại bài", "example": "Перед зачётом необходимо повторить пройденный материал."}
    ],
    24: [
        {"word": "носи́ть", "gram": "động từ chưa hoàn thành", "meaning": "mặc đồ giữ ấm", "example": "Зимой в России носят тёплые пуховики."},
        {"word": "надева́ть", "gram": "động từ chưa hoàn thành", "meaning": "mặc vào, xỏ vào", "example": "Наденьте тёплую шапку перед выходом на улицу."},
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua sắm", "example": "Мы покупаем зимние ботинки из натуральной кожи."},
        {"word": "примеря́ть", "gram": "động từ chưa hoàn thành", "meaning": "thử áo khoác", "example": "Покупатель примеряет шерстяное пальто в зеркале."}
    ],
    25: [
        {"word": "мечта́ть", "gram": "động từ chưa hoàn thành", "meaning": "mơ ước khát khao", "example": "Я мечтаю свободно владеть русским языком."},
        {"word": "плани́ровать", "gram": "động từ chưa hoàn thành", "meaning": "dự định tương lai", "example": "Мы планируем поступить в магистратуру в Москве."},
        {"word": "стреми́ться", "gram": "động từ chưa hoàn thành", "meaning": "nỗ lực vươn tới", "example": "Трудолюбивый студент стремится к высшим баллам."},
        {"word": "достига́ть", "gram": "động từ chưa hoàn thành", "meaning": "đạt tới đỉnh cao", "example": "Упорством и трудом мы достигаем поставленных целей."}
    ]
}

speaking_a1 = {}
for i in range(1, 26):
    noun_cat = a1_old.A1_SPEAKING_VOCAB[i][0]
    noun_cat["title"] = "Danh từ trong bài nói (Существительные)"
    verb_cat = {
        "title": "Động từ then chốt (Глаголы)",
        "rows": speaking_verbs[i]
    }
    speaking_a1[i] = [noun_cat, verb_cat]

# Writing A1 verbs mapping
writing_verbs = {
    1: [
        {"word": "жить", "gram": "động từ chưa hoàn thành", "meaning": "sống, cư trú", "example": "Я живу в столице Вьетнама — Ханое."},
        {"word": "учи́ться", "gram": "động từ chưa hoàn thành", "meaning": "học tập", "example": "Я студент и учусь на первом курсе."},
        {"word": "люби́ть", "gram": "động từ chưa hoàn thành", "meaning": "yêu thích", "example": "Я очень люблю изучать иностранные языки."},
        {"word": "писа́ть", "gram": "động từ chưa hoàn thành", "meaning": "viết thư", "example": "Я пишу тебе это письмо на русском языке."}
    ],
    2: [
        {"word": "рабо́тать", "gram": "động từ chưa hoàn thành", "meaning": "làm việc", "example": "Мой папа работает главным инженером."},
        {"word": "люби́ть", "gram": "động từ chưa hoàn thành", "meaning": "yêu quý gia đình", "example": "Мы любим собираться вместе за ужином."},
        {"word": "забо́титься", "gram": "động từ chưa hoàn thành", "meaning": "chăm sóc chu đáo", "example": "Родители заботятся о нашем образовании."},
        {"word": "помога́ть", "gram": "động từ chưa hoàn thành", "meaning": "giúp đỡ", "example": "Я всегда помогаю маме на кухне."}
    ],
    3: [
        {"word": "стоя́ть", "gram": "động từ chưa hoàn thành", "meaning": "đứng ở", "example": "Около окна стоит мой большой рабочий стол."},
        {"word": "лежа́ть", "gram": "động từ chưa hoàn thành", "meaning": "nằm ở", "example": "На кровати лежит мягкий тёплый плед."},
        {"word": "находи́ться", "gram": "động từ chưa hoàn thành", "meaning": "tọa lạc, nằm ở", "example": "Моя комната находится на третьем этаже."},
        {"word": "убира́ть", "gram": "động từ chưa hoàn thành", "meaning": "dọn dẹp ngăn nắp", "example": "Каждую субботу я убираю свою комнату."}
    ],
    4: [
        {"word": "просыпа́ться", "gram": "động từ chưa hoàn thành", "meaning": "thức giấc", "example": "Я просыпаюсь в семь часов утра."},
        {"word": "идти́", "gram": "động từ chưa hoàn thành", "meaning": "đi học", "example": "В восемь часов я иду в университет на лекции."},
        {"word": "возвраща́ться", "gram": "động từ chưa hoàn thành", "meaning": "trở về nhà", "example": "Вечером я возвращаюсь в общежитие."},
        {"word": "отдыха́ть", "gram": "động từ chưa hoàn thành", "meaning": "nghỉ ngơi", "example": "Перед сном я отдыхаю и читаю книгу."}
    ],
    5: [
        {"word": "обе́дать", "gram": "động từ chưa hoàn thành", "meaning": "dùng bữa trưa", "example": "В час дня я обычно обедаю в студенческом кафе."},
        {"word": "зака́зывать", "gram": "động từ chưa hoàn thành", "meaning": "gọi món", "example": "Я часто заказываю традиционный русский борщ."},
        {"word": "люби́ть", "gram": "động từ chưa hoàn thành", "meaning": "ưa thích món ăn", "example": "Я очень люблю свежую домашнюю выпечку."},
        {"word": "гото́вить", "gram": "động từ chưa hoàn thành", "meaning": "tự nấu ăn", "example": "На выходных я сам готовлю вьетнамский суп фо."}
    ],
    6: [
        {"word": "жить", "gram": "động từ chưa hoàn thành", "meaning": "sinh sống", "example": "Я родился и живу в древнем городе Ханой."},
        {"word": "расти́", "gram": "động từ chưa hoàn thành", "meaning": "phát triển lớn lên", "example": "Наш город быстро растёт и модернизируется."},
        {"word": "гуля́ть", "gram": "động từ chưa hoàn thành", "meaning": "đi dạo ven hồ", "example": "Вечером горожане любят гулять вокруг озера."},
        {"word": "люби́ть", "gram": "động từ chưa hoàn thành", "meaning": "yêu quý quê hương", "example": "Я всем сердцем люблю свой родной город."}
    ],
    7: [
        {"word": "дружи́ть", "gram": "động từ chưa hoàn thành", "meaning": "kết bạn gắn bó", "example": "Мы дружим со школьных лет."},
        {"word": "обща́ться", "gram": "động từ chưa hoàn thành", "meaning": "trò chuyện tâm tình", "example": "Мы каждый день общаемся и делимся мыслями."},
        {"word": "помога́ть", "gram": "động từ chưa hoàn thành", "meaning": "tương trợ lẫn nhau", "example": "Друг всегда готов помочь в трудную минуту."},
        {"word": "понима́ть", "gram": "động từ chưa hoàn thành", "meaning": "thấu hiểu bạn bè", "example": "Мы понимаем друг друга с полуслова."}
    ],
    8: [
        {"word": "отдыха́ть", "gram": "động từ chưa hoàn thành", "meaning": "nghỉ ngơi cuối tuần", "example": "В воскресенье я наконец-то отдыхаю от учёбы."},
        {"word": "ходи́ть", "gram": "động từ chưa hoàn thành", "meaning": "đi chơi", "example": "Мы ходим в кино или в художественную галерею."},
        {"word": "встреча́ться", "gram": "động từ chưa hoàn thành", "meaning": "hẹn gặp bạn bè", "example": "Днём я встречаюсь с друзьями в парке."},
        {"word": "смотре́ть", "gram": "động từ chưa hoàn thành", "meaning": "xem phim giải trí", "example": "Вечером мы смотрим интересный сериал."}
    ],
    9: [
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua thực phẩm", "example": "В субботу мама покупает свежие продукты на неделю."},
        {"word": "тра́тить", "gram": "động từ chưa hoàn thành", "meaning": "chi tiêu tiền bạc", "example": "Я стараюсь разумно тратить деньги."},
        {"word": "выбира́ть", "gram": "động từ chưa hoàn thành", "meaning": "chọn lựa hàng hóa", "example": "Мы тщательно выбираем спелые фрукты и овощи."},
        {"word": "гото́вить", "gram": "động từ chưa hoàn thành", "meaning": "chế biến món ăn", "example": "Из купленных продуктов мы готовим вкусный обед."}
    ],
    10: [
        {"word": "изуча́ть", "gram": "động từ chưa hoàn thành", "meaning": "học tập chuyên sâu", "example": "Я изучаю русский язык уже шесть месяцев."},
        {"word": "понима́ть", "gram": "động từ chưa hoàn thành", "meaning": "hiểu ngôn ngữ", "example": "Я хочу понимать русскую литературу в оригинале."},
        {"word": "говори́ть", "gram": "động từ chưa hoàn thành", "meaning": "nói chuyện giao tiếp", "example": "Моя цель — свободно говорить без акцента."},
        {"word": "мечта́ть", "gram": "động từ chưa hoàn thành", "meaning": "mơ ước du học", "example": "Я мечтаю учиться в университете в Москве."}
    ],
    11: [
        {"word": "свети́ть", "gram": "động từ chưa hoàn thành", "meaning": "tỏa sáng", "example": "Летом в моем городе ярко светит солнце."},
        {"word": "идти́ (дождь)", "gram": "cụm động từ", "meaning": "đổ mưa", "example": "Осенью часто идёт тёплый тропический дождь."},
        {"word": "люби́ть", "gram": "động từ chưa hoàn thành", "meaning": "yêu thích mùa hè", "example": "Больше всего я люблю мягкую весну."},
        {"word": "меня́ться", "gram": "động từ chưa hoàn thành", "meaning": "thay đổi theo mùa", "example": "Погода быстро меняется в течение одного дня."}
    ],
    12: [
        {"word": "отдыха́ть", "gram": "động từ chưa hoàn thành", "meaning": "nghỉ dưỡng", "example": "Летом я мечтаю отдыхать на тёплом морском побережье."},
        {"word": "купа́ться", "gram": "động từ chưa hoàn thành", "meaning": "tắm biển", "example": "Мы с удовольствием купаемся в прозрачной воде."},
        {"word": "загора́ть", "gram": "động từ chưa hoàn thành", "meaning": "tắm nắng", "example": "Туристы загорают на белом песке под солнцем."},
        {"word": "лета́ть", "gram": "động từ chưa hoàn thành", "meaning": "bay du lịch", "example": "Я люблю летать на самолёте в новые страны."}
    ],
    13: [
        {"word": "жить", "gram": "động từ chưa hoàn thành", "meaning": "sinh sống hòa thuận", "example": "Вся наша семья дружно живёт в уютной квартире."},
        {"word": "украша́ть", "gram": "động từ chưa hoàn thành", "meaning": "trang hoàng nhà cửa", "example": "К празднику мы украшаем гостиную живыми цветами."},
        {"word": "приглаша́ть", "gram": "động từ chưa hoàn thành", "meaning": "mời khách quý", "example": "Мы с радостью приглашаем друзей в наш дом."},
        {"word": "люби́ть", "gram": "động từ chưa hoàn thành", "meaning": "yêu mái ấm", "example": "Я люблю наш дом за теплоту и уютную атмосферу."}
    ],
    14: [
        {"word": "дари́ть", "gram": "động từ chưa hoàn thành", "meaning": "trao quà tặng", "example": "На день рождения друзья дарят мне полезные книги."},
        {"word": "поздравля́ть", "gram": "động từ chưa hoàn thành", "meaning": "chúc mừng sinh nhật", "example": "Все тепло поздравляют именинника и поют песню."},
        {"word": "пра́здновать", "gram": "động từ chưa hoàn thành", "meaning": "tổ chức kỷ niệm", "example": "Мы весело празднуем этот радостный день."},
        {"word": "ра́доваться", "gram": "động từ chưa hoàn thành", "meaning": "hạnh phúc hoan hỉ", "example": "Я искренне радуюсь встрече со всеми друзьями."}
    ],
    15: [
        {"word": "гото́виться", "gram": "động từ chưa hoàn thành", "meaning": "ôn thi, chuẩn bị", "example": "Я усердно готовлюсь к зимней экзаменационной сессии."},
        {"word": "чита́ть", "gram": "động từ chưa hoàn thành", "meaning": "nghiên cứu tài liệu", "example": "В читальном зале я читаю дополнительные статьи."},
        {"word": "писа́ть", "gram": "động từ chưa hoàn thành", "meaning": "viết bài tổng kết", "example": "Студент пишет подробный конспект по грамматике."},
        {"word": "сдава́ть", "gram": "động từ chưa hoàn thành", "meaning": "thi đỗ hoàn thành", "example": "Я надеюсь сдать все зачёты на отлично."}
    ],
    16: [
        {"word": "смотре́ть", "gram": "động từ chưa hoàn thành", "meaning": "thưởng thức phim", "example": "Я советую всем посмотреть эту замечательную комедию."},
        {"word": "по́мнить", "gram": "động từ chưa hoàn thành", "meaning": "nhớ mãi cốt truyện", "example": "Я навсегда запомнил трогательный финал этой картины."},
        {"word": "сове́товать", "gram": "động từ chưa hoàn thành", "meaning": "khuyên nhủ giới thiệu", "example": "Друг советует посмотреть новый исторический фильм."},
        {"word": "нра́виться", "gram": "động từ chưa hoàn thành", "meaning": "làm rung động", "example": "Мне очень понравилась великолепная игра актёров."}
    ],
    17: [
        {"word": "фотографи́ровать", "gram": "động từ chưa hoàn thành", "meaning": "chụp ảnh nghệ thuật", "example": "Моё главное увлечение — фотографировать природу."},
        {"word": "снима́ть", "gram": "động từ chưa hoàn thành", "meaning": "bấm máy chụp cảnh", "example": "Я люблю снимать старинные улочки на закате."},
        {"word": "сохраня́ть", "gram": "động từ chưa hoàn thành", "meaning": "lưu giữ kỷ niệm", "example": "Фотографии сохраняют самые счастливые моменты."},
        {"word": "пока́зывать", "gram": "động từ chưa hoàn thành", "meaning": "khoe hình ảnh", "example": "Я показываю свои снимки друзьям в соцсетях."}
    ],
    18: [
        {"word": "благодари́ть", "gram": "động từ chưa hoàn thành", "meaning": "cảm ơn chân thành", "example": "Искренне благодарю вас за ценную помощь и совет."},
        {"word": "помога́ть", "gram": "động từ chưa hoàn thành", "meaning": "tận tình giúp đỡ", "example": "Преподаватель всегда терпеливо помогает студентам."},
        {"word": "по́мнить", "gram": "động từ chưa hoàn thành", "meaning": "khắc ghi ân tình", "example": "Я всегда буду помнить вашу доброту и поддержку."},
        {"word": "цени́ть", "gram": "động từ chưa hoàn thành", "meaning": "trân trọng biết ơn", "example": "Мы высоко ценим ваш огромный педагогический труд."}
    ],
    19: [
        {"word": "украша́ть", "gram": "động từ chưa hoàn thành", "meaning": "trang hoàng cây thông", "example": "Дети с восторгом украшают ёлку блестящими шарами."},
        {"word": "отмеча́ть", "gram": "động từ chưa hoàn thành", "meaning": "đón mừng năm mới", "example": "В полночь вся страна торжественно отмечает Новый год."},
        {"word": "жела́ть", "gram": "động từ chưa hoàn thành", "meaning": "cầu chúc an lành", "example": "Мы желаем родным счастья, мира и крепкого здоровья."},
        {"word": "собира́ться", "gram": "động từ chưa hoàn thành", "meaning": "quây quần đoàn viên", "example": "Все поколения семьи собираются за праздничным столом."}
    ],
    20: [
        {"word": "вы́учить", "gram": "động từ hoàn thành", "meaning": "học thuộc lòng vững", "example": "За курс А1 я выучил более пятисот русских слов."},
        {"word": "понима́ть", "gram": "động từ chưa hoàn thành", "meaning": "hiểu thông thạo", "example": "Теперь я без труда понимаю простые тексты и диалоги."},
        {"word": "говори́ть", "gram": "động từ chưa hoàn thành", "meaning": "tự tin giao tiếp", "example": "Я умею рассказывать о себе и поддерживать разговор."},
        {"word": "горди́ться", "gram": "động từ chưa hoàn thành", "meaning": "tự hào về bản thân", "example": "Я горжусь своими первыми успехами в изучении языка."}
    ],
    21: [
        {"word": "е́хать", "gram": "động từ chưa hoàn thành", "meaning": "lên đường du hành", "example": "На каникулах мы едем в Санкт-Петербург на поезде."},
        {"word": "собира́ть", "gram": "động từ chưa hoàn thành", "meaning": "thu xếp đồ đạc", "example": "Я аккуратно собираю чемодан перед дальней поездкой."},
        {"word": "смотре́ть", "gram": "động từ chưa hoàn thành", "meaning": "ngắm nhìn phong cảnh", "example": "Из окна вагона приятно смотреть на заснеженные леса."},
        {"word": "возвраща́ться", "gram": "động từ chưa hoàn thành", "meaning": "trở về bình an", "example": "С яркими впечатлениями мы возвращаемся домой."}
    ],
    22: [
        {"word": "писа́ть", "gram": "động từ chưa hoàn thành", "meaning": "viết thiệp gửi bạn", "example": "Я пишу праздничную открытку своему другу в Россию."},
        {"word": "отправля́ть", "gram": "động từ chưa hoàn thành", "meaning": "gửi bưu phẩm", "example": "Она отправляет красивую открытку с видами Ханоя."},
        {"word": "поздравля́ть", "gram": "động từ chưa hoàn thành", "meaning": "chúc mừng dịp lễ", "example": "От всей души поздравляю тебя с наступающим праздником!"},
        {"word": "жела́ть", "gram": "động từ chưa hoàn thành", "meaning": "chúc phúc", "example": "Желаю тебе радости, вдохновения и отличных оценок!"}
    ],
    23: [
        {"word": "есть", "gram": "động từ chưa hoàn thành", "meaning": "ăn uống điều độ", "example": "Я стараюсь есть больше свежих овощей и фруктов."},
        {"word": "пить", "gram": "động từ chưa hoàn thành", "meaning": "uống đủ nước", "example": "Каждый день я пью не менее полутора литров чистой воды."},
        {"word": "гото́вить", "gram": "động từ chưa hoàn thành", "meaning": "tự nấu ăn sạch", "example": "Домашняя еда всегда полезнее и вкуснее фастфуда."},
        {"word": "следи́ть", "gram": "động từ chưa hoàn thành", "meaning": "chú ý sức khỏe", "example": "Очень важно следить за своим ежедневным рационом."}
    ],
    24: [
        {"word": "бере́чь", "gram": "động từ chưa hoàn thành", "meaning": "gìn giữ trân trọng", "example": "Мы должны беречь зелёную природу нашего родного города."},
        {"word": "убира́ть", "gram": "động từ chưa hoàn thành", "meaning": "thu dọn rác thải", "example": "Студенты убирают мусор в городском парке на субботнике."},
        {"word": "сажа́ть", "gram": "động từ chưa hoàn thành", "meaning": "trồng thêm cây xanh", "example": "Весной школьники сажают молодые деревья около школы."},
        {"word": "защища́ть", "gram": "động từ chưa hoàn thành", "meaning": "bảo vệ môi trường", "example": "Каждый из нас может защищать экологию простыми делами."}
    ],
    25: [
        {"word": "переходи́ть", "gram": "động từ chưa hoàn thành", "meaning": "chuyển cấp tiến lên", "example": "Я уверенно перехожу на уровень А2 с отличной базой."},
        {"word": "продолжа́ть", "gram": "động từ chưa hoàn thành", "meaning": "kiên trì tiếp tục", "example": "Я буду ежедневно продолжать учить новые грамматические темы."},
        {"word": "совершенствова́ть", "gram": "động từ chưa hoàn thành", "meaning": "trau dồi kỹ năng", "example": "Моя цель — совершенствовать разговорную речь и слух."},
        {"word": "мечта́ть", "gram": "động từ chưa hoàn thành", "meaning": "hoài bão lớn lao", "example": "Я мечтаю свободно читать классиков русской литературы."}
    ]
}

writing_a1 = {}
for i in range(1, 26):
    noun_cat = a1_old.A1_WRITING_VOCAB[i][0]
    noun_cat["title"] = "Danh từ trong bài viết (Существительные)"
    verb_cat = {
        "title": "Động từ then chốt (Глаголы)",
        "rows": writing_verbs[i]
    }
    writing_a1[i] = [noun_cat, verb_cat]

a1_master = {
    'reading': reading_a1,
    'listening': listening_a1,
    'speaking': speaking_a1,
    'writing': writing_a1
}

with open('scripts/vocab_master_a1.json', 'w', encoding='utf-8') as f:
    json.dump(a1_master, f, ensure_ascii=False, indent=2)

print("SUCCESS: scripts/vocab_master_a1.json built successfully!")
print("Counts: Reading:", len(reading_a1), "Listening:", len(listening_a1), "Speaking:", len(speaking_a1), "Writing:", len(writing_a1))
