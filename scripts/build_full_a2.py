# -*- coding: utf-8 -*-
"""
Build complete 100 lessons for Level A2 (Reading, Listening, Speaking, Writing).
Each lesson gets:
- Category 1: Danh từ trong bài học (Существительные) - 4-5 words
- Category 2: Động từ then chốt (Глаголы) - 4 words
Total 8-10 words per lesson = ~850 words for A2.
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import scripts.vocab_data_a2 as a2_old

# 1. Reading A2
reading_verbs_a2 = {
    8: [
        {"word": "выступа́ть", "gram": "động từ chưa hoàn thành", "meaning": "biểu diễn", "example": "Актёры выступают на сцене театра."},
        {"word": "аплоди́ровать", "gram": "động từ chưa hoàn thành", "meaning": "vỗ tay tán thưởng", "example": "Зрители восторженно аплодируют артистам."},
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua vé", "example": "Мы заранее покупаем билеты в партер."},
        {"word": "восхища́ться", "gram": "động từ chưa hoàn thành", "meaning": "ngưỡng mộ thán phục", "example": "Публика восхищается мастерством балерины."}
    ],
    9: [
        {"word": "снима́ть (жильё)", "gram": "cụm động từ", "meaning": "thuê (nhà ở)", "example": "Студенты снимают квартиру рядом с институтом."},
        {"word": "плати́ть", "gram": "động từ chưa hoàn thành", "meaning": "trả tiền thuê", "example": "Они ежемесячно платят за аренду и коммунальные услуги."},
        {"word": "подпи́сывать", "gram": "động từ chưa hoàn thành", "meaning": "ký hợp đồng", "example": "Стороны подписывают официальный договор найма."},
        {"word": "переезжа́ть", "gram": "động từ chưa hoàn thành", "meaning": "chuyển nhà", "example": "Семья переезжает в новую просторную квартиру."}
    ],
    10: [
        {"word": "тренирова́ться", "gram": "động từ chưa hoàn thành", "meaning": "tập luyện thể lực", "example": "Молодёжь активно тренируется в фитнес-клубах."},
        {"word": "бе́гать", "gram": "động từ chưa hoàn thành", "meaning": "chạy bộ", "example": "По утрам многие бегают по дорожкам парка."},
        {"word": "ката́ться (на лыжах)", "gram": "cụm động từ", "meaning": "trượt tuyết", "example": "Зимой мы любим кататься на беговых лыжах в лесу."},
        {"word": "укрепля́ть", "gram": "động từ chưa hoàn thành", "meaning": "tăng cường sức khỏe", "example": "Регулярный спорт укрепляет сердце и мышцы."}
    ],
    11: [
        {"word": "создава́ть", "gram": "động từ chưa hoàn thành", "meaning": "sáng tác, làm ra", "example": "Мастера создают уникальные деревянные изделия."},
        {"word": "распи́сывать", "gram": "động từ chưa hoàn thành", "meaning": "vẽ trang trí hoa văn", "example": "Художники расписывают посуду хохломскими узорами."},
        {"word": "храни́ть", "gram": "động từ chưa hoàn thành", "meaning": "gìn giữ lưu truyền", "example": "Народ бережно хранит вековые традиции ремёсел."},
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua đồ lưu niệm", "example": "Иностранцы с радостью покупают расписные матрешки."}
    ],
    12: [
        {"word": "приглаша́ть", "gram": "động từ chưa hoàn thành", "meaning": "mời đến chơi", "example": "Хозяева радушно приглашают гостей в дом."},
        {"word": "угоща́ть", "gram": "động từ chưa hoàn thành", "meaning": "chiêu đãi tiệc", "example": "Гостей щедро угощают горячим чаем и пирогами."},
        {"word": "дари́ть", "gram": "động từ chưa hoàn thành", "meaning": "tặng quà", "example": "В гости принято приходить с тортом или цветами."},
        {"word": "благодари́ть", "gram": "động từ chưa hoàn thành", "meaning": "cảm ơn nồng hậu", "example": "Гости искренне благодарят за тёплый приём."}
    ],
    13: [
        {"word": "отправля́ть", "gram": "động từ chưa hoàn thành", "meaning": "gửi bưu phẩm", "example": "Она отправляет заказное письмо с уведомлением."},
        {"word": "получа́ть", "gram": "động từ chưa hoàn thành", "meaning": "nhận hàng", "example": "Я получаю посылку по коду из смс."},
        {"word": "заполня́ть", "gram": "động từ chưa hoàn thành", "meaning": "điền biểu mẫu", "example": "Отправитель заполняет почтовый бланк."},
        {"word": "доставля́ть", "gram": "động từ chưa hoàn thành", "meaning": "vận chuyển phát", "example": "Курьер быстро доставляет бандероль до двери."}
    ],
    14: [
        {"word": "гото́вить", "gram": "động từ chưa hoàn thành", "meaning": "chế biến món ăn", "example": "Русские хозяйки часто готовят пельмени всей семьёй."},
        {"word": "вари́ть", "gram": "động từ chưa hoàn thành", "meaning": "ninh, luộc, nấu", "example": "В чугунке варят наваристые щи из свежей капусты."},
        {"word": "печь", "gram": "động từ chưa hoàn thành", "meaning": "nướng bánh", "example": "В печи пекут пышные караваи и расстегаи."},
        {"word": "про́бовать", "gram": "động từ chưa hoàn thành", "meaning": "thưởng thức nếm", "example": "Обязательно попробуйте традиционный морс из клюквы."}
    ],
    15: [
        {"word": "заводи́ть (питомца)", "gram": "cụm động từ", "meaning": "nuôi thú cưng", "example": "Семья решила завести породистую собаку."},
        {"word": "гуля́ть", "gram": "động từ chưa hoàn thành", "meaning": "dắt đi dạo", "example": "Владельцы гуляют с собаками два раза в день."},
        {"word": "корми́ть", "gram": "động từ chưa hoàn thành", "meaning": "cho ăn thức ăn", "example": "Кошку нужно кормить сбалансированным кормом."},
        {"word": "забо́титься", "gram": "động từ chưa hoàn thành", "meaning": "chăm sóc tận tụy", "example": "Люди с любовью заботятся о бездомных животных."}
    ],
    16: [
        {"word": "одева́ться", "gram": "động từ chưa hoàn thành", "meaning": "mặc ấm giữ nhiệt", "example": "В сильные морозы важно одеваться многослойно."},
        {"word": "надева́ть", "gram": "động từ chưa hoàn thành", "meaning": "đeo khăn mang găng", "example": "Перед выходом на мороз надевайте шапку и варежки."},
        {"word": "выбира́ть", "gram": "động từ chưa hoàn thành", "meaning": "chọn trang phục", "example": "Она выбирает пальто из натуральной шерсти."},
        {"word": "согрева́ть", "gram": "động từ chưa hoàn thành", "meaning": "sưởi ấm cơ thể", "example": "Тёплый пуховик отлично согревает в метель."}
    ],
    17: [
        {"word": "чита́ть", "gram": "động từ chưa hoàn thành", "meaning": "đọc say mê", "example": "Многие читают электронные книги в метро по пути на работу."},
        {"word": "обсужда́ть", "gram": "động từ chưa hoàn thành", "meaning": "thảo luận sách", "example": "В книжном клубе читатели обсуждают сюжет романа."},
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua sách in", "example": "Я люблю покупать книги в старых букинистических лавках."},
        {"word": "рекомендова́ть", "gram": "động từ chưa hoàn thành", "meaning": "giới thiệu sách hay", "example": "Друг рекомендует прочитать рассказы Чехова."}
    ],
    18: [
        {"word": "рабо́тать", "gram": "động từ chưa hoàn thành", "meaning": "làm việc cống hiến", "example": "Он работает ведущим специалистом в IT-компании."},
        {"word": "иска́ть (работу)", "gram": "cụm động từ", "meaning": "tìm kiếm việc làm", "example": "Выпускники вузов ищут работу по специальности."},
        {"word": "учи́ться", "gram": "động từ chưa hoàn thành", "meaning": "học hỏi nâng cao", "example": "Современный профессионал учится на протяжении всей жизни."},
        {"word": "зараба́тывать", "gram": "động từ chưa hoàn thành", "meaning": "kiếm thu nhập", "example": "Специалисты высокой квалификации хорошо зарабатывают."}
    ],
    19: [
        {"word": "печь (блины)", "gram": "cụm động từ", "meaning": "rán bánh xèo blini", "example": "Всю масленичную неделю хозяйки пекут золотистые блины."},
        {"word": "провожа́ть (зиму)", "gram": "cụm động từ", "meaning": "tiễn đưa mùa đông", "example": "Народными гуляниями народ провожает холодную зиму."},
        {"word": "встреча́ть (весну)", "gram": "cụm động từ", "meaning": "đón chào mùa xuân", "example": "С песнями и хороводами люди встречают тёплую весну."},
        {"word": "весели́ться", "gram": "động từ chưa hoàn thành", "meaning": "vui chơi thỏa thích", "example": "На городских площадях все веселятся и участвуют в играх."}
    ],
    20: [
        {"word": "пра́здновать", "gram": "động từ chưa hoàn thành", "meaning": "kỷ niệm trang nghiêm", "example": "Девятого мая вся страна празднует священный День Победы."},
        {"word": "поздравля́ть", "gram": "động từ chưa hoàn thành", "meaning": "tri ân cựu chiến binh", "example": "Школьники с цветами поздравляют ветеранов войны."},
        {"word": "вспомина́ть", "gram": "động từ chưa hoàn thành", "meaning": "tưởng nhớ hy sinh", "example": "Мы вспоминаем подвиг героев минутой молчания."},
        {"word": "смотре́ть (салют)", "gram": "cụm động từ", "meaning": "xem pháo hoa", "example": "Вечером миллионы людей смотрят праздничный салют."}
    ],
    21: [
        {"word": "гуля́ть", "gram": "động từ chưa hoàn thành", "meaning": "dạo bước ngắm cảnh", "example": "Туристы гуляют по гранитным набережным реки Невы."},
        {"word": "разводи́ть (мосты)", "gram": "cụm động từ", "meaning": "mở cầu xoay", "example": "Ночью в Петербурге разводят огромные разводные мосты."},
        {"word": "любова́ться", "gram": "động từ chưa hoàn thành", "meaning": "chiêm ngưỡng vẻ đẹp", "example": "Все любуются белыми ночами и дворцами города."},
        {"word": "фотографи́ровать", "gram": "động từ chưa hoàn thành", "meaning": "chụp ảnh kiến trúc", "example": "Я фотографирую величественный Исаакиевский собор."}
    ],
    22: [
        {"word": "посеща́ть", "gram": "động từ chưa hoàn thành", "meaning": "thăm quan bảo tàng", "example": "Любители искусства посещают Эрмитаж и Русский музей."},
        {"word": "выставля́ть", "gram": "động từ chưa hoàn thành", "meaning": "trưng bày kiệt tác", "example": "В залах выставляют редкие полотна великих художников."},
        {"word": "понима́ть", "gram": "động từ chưa hoàn thành", "meaning": "cảm thụ nghệ thuật", "example": "Экскурсия помогает глубже понимать замысел живописца."},
        {"word": "восхища́ться", "gram": "động từ chưa hoàn thành", "meaning": "thán phục tài hoa", "example": "Зрители восхищаются реализмом русских пейзажей."}
    ],
    23: [
        {"word": "забо́титься", "gram": "động từ chưa hoàn thành", "meaning": "chăm lo thể chất", "example": "Необходимо с молодости заботиться о своём здоровье."},
        {"word": "соблюда́ть (режим)", "gram": "cụm động từ", "meaning": "tuân thủ giờ giấc", "example": "Врачи советуют соблюдать режим сна и отдыха."},
        {"word": "пита́ться", "gram": "động từ chưa hoàn thành", "meaning": "ăn uống lành mạnh", "example": "Для долголетия важно правильно и разнообразно питаться."},
        {"word": "отка́зываться", "gram": "động từ chưa hoàn thành", "meaning": "từ bỏ thói xấu", "example": "Многие отказываются от вредных привычек ради спорта."}
    ],
    24: [
        {"word": "дружи́ть", "gram": "động từ chưa hoàn thành", "meaning": "kết bạn chung thủy", "example": "Настоящие друзья дружат бескорыстно и искренне."},
        {"word": "помога́ть", "gram": "động từ chưa hoàn thành", "meaning": "tương thân tương ái", "example": "Верный товарищ всегда готов помочь без лишних слов."},
        {"word": "дове́рять", "gram": "động từ chưa hoàn thành", "meaning": "tin tưởng bạn bè", "example": "В крепкой дружбе очень важно полностью доверять друг другу."},
        {"word": "подде́рживать", "gram": "động từ chưa hoàn thành", "meaning": "sát cánh động viên", "example": "Друзья поддерживают друг друга во всех начинаниях."}
    ],
    25: [
        {"word": "осва́ивать", "gram": "động từ chưa hoàn thành", "meaning": "lĩnh hội kiến thức", "example": "Мы успешно осваиваем грамматику продвинутого уровня."},
        {"word": "достига́ть", "gram": "động từ chưa hoàn thành", "meaning": "chinh phục mục tiêu", "example": "Благодаря упорному труду студент достигает уровня B1."},
        {"word": "сдава́ть (экзамен)", "gram": "cụm động từ", "meaning": "thi đỗ chứng chỉ", "example": "Я готов сдать международный сертификационный экзамен."},
        {"word": "продолжа́ть", "gram": "động từ chưa hoàn thành", "meaning": "không ngừng học tập", "example": "Мы продолжаем погружаться в богатую русскую культуру."}
    ]
}

# Reading A2 lessons
reading_a2 = {}
for i in range(1, 26):
    fname = f'public/data/reading/a2/lesson-{i:02d}.json'
    with open(fname, 'r', encoding='utf-8') as f:
        d = json.load(f)
    existing_cats = d.get('vocabCategories', [])
    if i in reading_verbs_a2:
        # Keep noun cat, add/replace verb cat
        noun_cat = existing_cats[0]
        noun_cat["title"] = "Danh từ trong bài đọc (Существительные)"
        verb_cat = {
            "title": "Động từ then chốt (Глаголы)",
            "rows": reading_verbs_a2[i]
        }
        reading_a2[i] = [noun_cat, verb_cat]
    else:
        reading_a2[i] = existing_cats

# Listening A2 verbs
listening_verbs_a2 = {
    1: [
        {"word": "дое́хать", "gram": "động từ hoàn thành", "meaning": "đi xe đến đích", "example": "Как доехать до Красной площади?"},
        {"word": "переса́живаться", "gram": "động từ chưa hoàn thành", "meaning": "chuyển sang tuyến khác", "example": "На станции «Охотный ряд» нужно пересаживаться."},
        {"word": "спуска́ться", "gram": "động từ chưa hoàn thành", "meaning": "đi xuống ga", "example": "Пассажиры спускаются в метро на эскалаторе."},
        {"word": "подска́зывать", "gram": "động từ chưa hoàn thành", "meaning": "chỉ đường mách bảo", "example": "Прохожий вежливо подсказывает правильный маршрут."}
    ],
    2: [
        {"word": "брони́ровать", "gram": "động từ chưa hoàn thành", "meaning": "đặt bàn trước", "example": "Я хочу забронировать столик на вечер пятницы."},
        {"word": "зака́зывать", "gram": "động từ chưa hoàn thành", "meaning": "gọi món ăn", "example": "Мы заказываем суп уху и запечённую рыбу."},
        {"word": "рекомендова́ть", "gram": "động từ chưa hoàn thành", "meaning": "gợi ý món ngon", "example": "Официант рекомендует фирменные сибирские пельмени."},
        {"word": "плати́ть", "gram": "động từ chưa hoàn thành", "meaning": "thanh toán hóa đơn", "example": "Гости платят по счёту кредитной картой."}
    ],
    3: [
        {"word": "боле́ть", "gram": "động từ chưa hoàn thành", "meaning": "bị sốt cảm", "example": "Друг заболел и не пришёл на лекции."},
        {"word": "вызыва́ть (врача)", "gram": "cụm động từ", "meaning": "gọi bác sĩ đến nhà", "example": "При высокой температуре нужно вызвать врача на дом."},
        {"word": "измеря́ть", "gram": "động từ chưa hoàn thành", "meaning": "đo nhiệt độ", "example": "Больной измеряет температуру градусником."},
        {"word": "принима́ть", "gram": "động từ chưa hoàn thành", "meaning": "uống thuốc trị bệnh", "example": "Принимайте жаропонижающие таблетки после еды."}
    ],
    4: [
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua vé tàu", "example": "Пассажир покупает билет на скорый поезд."},
        {"word": "выбира́ть", "gram": "động từ chưa hoàn thành", "meaning": "chọn loại toa", "example": "Мы выбираем нижнюю полку в купейном вагоне."},
        {"word": "отправля́ться", "gram": "động từ chưa hoàn thành", "meaning": "rời ga xuất phát", "example": "Фирменный поезд отправляется ровно в полночь."},
        {"word": "предъявля́ть", "gram": "động từ chưa hoàn thành", "meaning": "xuất trình vé giấy tờ", "example": "При посадке предъявите паспорт и билет проводнику."}
    ],
    5: [
        {"word": "снима́ть", "gram": "động từ chưa hoàn thành", "meaning": "thuê căn hộ", "example": "Студент хочет снять однокомнатную квартиру на длительный срок."},
        {"word": "звони́ть", "gram": "động từ chưa hoàn thành", "meaning": "gọi điện hỏi tin", "example": "Я звоню по объявлению об аренде жилья."},
        {"word": "осма́тривать", "gram": "động từ chưa hoàn thành", "meaning": "đến xem phòng", "example": "Вечером арендатор осматривает состояние квартиры."},
        {"word": "догова́риваться", "gram": "động từ chưa hoàn thành", "meaning": "thỏa thuận giá cả", "example": "Стороны договариваются о размере ежемесячной оплаты."}
    ],
    6: [
        {"word": "регистри́роваться", "gram": "động từ chưa hoàn thành", "meaning": "làm thủ tục bay", "example": "Пассажиры регистрируются на стойке в терминале."},
        {"word": "сдава́ть (багаж)", "gram": "cụm động từ", "meaning": "gửi hành lý", "example": "Мы сдаём тяжёлые чемоданы в багажное отделение."},
        {"word": "взве́шивать", "gram": "động từ chưa hoàn thành", "meaning": "cân xách tay", "example": "Сотрудник взвешивает ручную кладь перед посадкой."},
        {"word": "лете́ть", "gram": "động từ chưa hoàn thành", "meaning": "bay chuyển cánh", "example": "Самолёт летит прямым рейсом в Новосибирск."}
    ],
    7: [
        {"word": "открыва́ть (счёт)", "gram": "cụm động từ", "meaning": "mở tài khoản ngân hàng", "example": "Иностранец открывает рублёвый счёт в банке."},
        {"word": "меня́ть", "gram": "động từ chưa hoàn thành", "meaning": "quy đổi ngoại tệ", "example": "Клиент меняет доллары по текущему курсу валют."},
        {"word": "переводи́ть", "gram": "động từ chưa hoàn thành", "meaning": "chuyển khoản tiền", "example": "Я перевожу деньги родителям через мобильный банк."},
        {"word": "снима́ть", "gram": "động từ chưa hoàn thành", "meaning": "rút tiền mặt ATM", "example": "Он снимает нужную сумму в банкомате без комиссии."}
    ],
    8: [
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua vé kịch", "example": "Мы покупаем билеты на балет «Лебединое озеро»."},
        {"word": "сдава́ть (в гардероб)", "gram": "cụm động từ", "meaning": "gửi áo khoác", "example": "Зрители сдают пальто в гардероб перед спектаклем."},
        {"word": "занима́ть (место)", "gram": "cụm động từ", "meaning": "ngồi vào chỗ", "example": "После третьего звонка зрители занимают свои места."},
        {"word": "аплоди́ровать", "gram": "động từ chưa hoàn thành", "meaning": "nhiệt liệt vỗ tay", "example": "Зал стоя аплодирует выдающимся артистам."}
    ],
    9: [
        {"word": "ходи́ть (на приём)", "gram": "cụm động từ", "meaning": "đến khám nha khoa", "example": "Пациент ходит на профилактический приём раз в полгода."},
        {"word": "жа́ловаться", "gram": "động từ chưa hoàn thành", "meaning": "kêu đau răng", "example": "Она жалуется на острую зубную боль по ночам."},
        {"word": "лечи́ть", "gram": "động từ chưa hoàn thành", "meaning": "hàn trám điều trị", "example": "Опытный стоматолог безболезненно лечит зуб."},
        {"word": "полоска́ть", "gram": "động từ chưa hoàn thành", "meaning": "súc miệng sát trùng", "example": "Врач советует полоскать рот травяным настоем."}
    ],
    10: [
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua thẻ tập gym", "example": "Студент покупает годовой абонемент в фитнес-клуб."},
        {"word": "занима́ться", "gram": "động từ chưa hoàn thành", "meaning": "tập với máy", "example": "Он занимается на современных силовых тренажёрах."},
        {"word": "пла́вать", "gram": "động từ chưa hoàn thành", "meaning": "bơi trong bể", "example": "После тренировки приятно поплавать в бассейне."},
        {"word": "бе́гать", "gram": "động từ chưa hoàn thành", "meaning": "chạy trên thảm", "example": "Она бегает на беговой дорожке по двадцать минут."}
    ],
    11: [
        {"word": "выбира́ть", "gram": "động từ chưa hoàn thành", "meaning": "chọn quà thủ công", "example": "Туристы выбирают традиционные сувениры с ярмарки."},
        {"word": "распи́сывать", "gram": "động từ chưa hoàn thành", "meaning": "chạm vẽ thủ công", "example": "Мастера тонко расписывают деревянную шкатулку."},
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua đồ gốm sứ", "example": "Я покупаю фарфоровую чашку с гжельским узором."},
        {"word": "упако́вывать", "gram": "động từ chưa hoàn thành", "meaning": "đóng gói hộp quà", "example": "Продавец аккуратно упаковывает хрупкий сувенир."}
    ],
    12: [
        {"word": "приходи́ть (в гости)", "gram": "cụm động từ", "meaning": "đến chơi nhà bạn", "example": "Друзья приходят в гости с тёплыми улыбками."},
        {"word": "дари́ть", "gram": "động từ chưa hoàn thành", "meaning": "trao hộp bánh", "example": "Гость дарит хозяйке красивый букет цветов."},
        {"word": "уса́живать", "gram": "động từ chưa hoàn thành", "meaning": "mời ngồi vào bàn", "example": "Хозяева радушно усаживают всех за праздничный стол."},
        {"word": "налива́ть", "gram": "động từ chưa hoàn thành", "meaning": "rót trà nóng", "example": "Она наливает ароматный чай из пузатого самовара."}
    ],
    13: [
        {"word": "отправля́ть", "gram": "động từ chưa hoàn thành", "meaning": "gửi bưu phẩm", "example": "Я отправляю посылку с зимними вещами брату."},
        {"word": "взве́шивать", "gram": "động từ chưa hoàn thành", "meaning": "cân kiện hàng", "example": "Оператор почты взвешивает коробку на весах."},
        {"word": "опла́чивать", "gram": "động từ chưa hoàn thành", "meaning": "thanh toán cước phí", "example": "Клиент оплачивает доставку в окне кассы."},
        {"word": "отсле́живать", "gram": "động từ chưa hoàn thành", "meaning": "tra cứu vận đơn", "example": "По трек-номеру удобно отслеживать посылку онлайн."}
    ],
    14: [
        {"word": "зака́зывать", "gram": "động từ chưa hoàn thành", "meaning": "gọi bánh xèo Nga", "example": "Мы заказываем сытные блины с мясным фаршем."},
        {"word": "про́бовать", "gram": "động từ chưa hoàn thành", "meaning": "thưởng thức ẩm thực", "example": "Иностранец впервые пробует сладкие сырники со сметаной."},
        {"word": "печь", "gram": "động từ chưa hoàn thành", "meaning": "rán nướng nhanh", "example": "Повар быстро печёт горячие блины на круглой сковороде."},
        {"word": "сове́товать", "gram": "động từ chưa hoàn thành", "meaning": "khuyên thử món", "example": "Друг советует взять ягодный морс к десерту."}
    ],
    15: [
        {"word": "приводи́ть (к ветеринару)", "gram": "cụm động từ", "meaning": "đưa đi khám thú y", "example": "Хозяин приводит заболевшую собаку в ветклинику."},
        {"word": "осма́тривать", "gram": "động từ chưa hoàn thành", "meaning": "thăm khám thú cưng", "example": "Ветеринар внимательно осматривает лапу щенка."},
        {"word": "де́лать (прививку)", "gram": "cụm động từ", "meaning": "tiêm vắc-xin", "example": "Доктору нужно сделать плановую прививку котенку."},
        {"word": "выпи́сывать", "gram": "động từ chưa hoàn thành", "meaning": "kê đơn thuốc nhỏ", "example": "Врач выписывает витамины для шерсти и здоровья."}
    ],
    16: [
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua áo ấm lông vũ", "example": "Мы покупаем непромокаемый пуховик на зиму."},
        {"word": "примеря́ть", "gram": "động từ chưa hoàn thành", "meaning": "thử áo khoác gương", "example": "Покупатель примеряет зимнюю куртку в примерочной."},
        {"word": "выбира́ть", "gram": "động từ chưa hoàn thành", "meaning": "chọn màu hợp mốt", "example": "Она выбирает стильный шерстяной шарф к пальто."},
        {"word": "греть", "gram": "động từ chưa hoàn thành", "meaning": "giữ ấm tuyệt vời", "example": "Натуральный пух отлично греет в тридцатиградусный мороз."}
    ],
    17: [
        {"word": "выбира́ть", "gram": "động từ chưa hoàn thành", "meaning": "chọn cuốn sách hay", "example": "Покупатель долго выбирает роман в книжном магазине."},
        {"word": "рекомендова́ть", "gram": "động từ chưa hoàn thành", "meaning": "tư vấn sách mới", "example": "Консультант рекомендует бестселлер известного автора."},
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua ấn phẩm in", "example": "Я покупаю иллюстрированное подарочное издание."},
        {"word": "чита́ть", "gram": "động từ chưa hoàn thành", "meaning": "say sưa đọc sách", "example": "Он с увлечением читает книгу по вечерам."}
    ],
    18: [
        {"word": "проходи́ть (собеседование)", "gram": "cụm động từ", "meaning": "tham dự phỏng vấn", "example": "Кандидат проходит собеседование на должность менеджера."},
        {"word": "расска́зывать", "gram": "động từ chưa hoàn thành", "meaning": "trình bày kinh nghiệm", "example": "Соискатель рассказывает о своих профессиональных навыках."},
        {"word": "задава́ть (вопросы)", "gram": "cụm động từ", "meaning": "đặt câu hỏi kiểm tra", "example": "Работодатель задаёт вопросы о предыдущих проектах."},
        {"word": "предлага́ть", "gram": "động từ chưa hoàn thành", "meaning": "đưa lời mời làm việc", "example": "Компания предлагает отличные условия труда и зарплату."}
    ],
    19: [
        {"word": "печь (блины)", "gram": "cụm động từ", "meaning": "làm bánh đãi bạn", "example": "На Масленицу студенты пекут блины на кухне общежития."},
        {"word": "угоща́ть", "gram": "động từ chưa hoàn thành", "meaning": "mời bạn ăn bánh", "example": "Анна щедро угощает однокурсников домашними блинами."},
        {"word": "гуля́ть", "gram": "động từ chưa hoàn thành", "meaning": "dạo chơi hội xuân", "example": "Молодёжь гуляет в парке и смотрит праздничные забавы."},
        {"word": "сжигать (чучело)", "gram": "cụm động từ", "meaning": "đốt hình nộm rơm", "example": "В воскресенье по традиции сжигают соломенное чучело зимы."}
    ],
    20: [
        {"word": "встреча́ть", "gram": "động từ chưa hoàn thành", "meaning": "đón chào ngày lễ", "example": "Город торжественно встречает праздник Великой Победы."},
        {"word": "возлага́ть (цветы)", "gram": "cụm động từ", "meaning": "dâng hoa tưởng niệm", "example": "Люди возлагают красные гвоздики к Вечному огню."},
        {"word": "поздравля́ть", "gram": "động từ chưa hoàn thành", "meaning": "chúc mừng cựu binh", "example": "Волонтёры поздравляют ветеранов и вручают подарки."},
        {"word": "смотре́ть (парад)", "gram": "cụm động từ", "meaning": "theo dõi duyệt binh", "example": "Миллионы семей смотрят военный парад по телевизору."}
    ],
    21: [
        {"word": "покупа́ть (экскурсию)", "gram": "cụm động từ", "meaning": "mua tour du thuyền", "example": "Туристы покупают ночную водную экскурсию по каналам."},
        {"word": "плыть", "gram": "động từ chưa hoàn thành", "meaning": "đi thuyền trên sông", "example": "Теплоход плавно плывёт по реке Неве мимо дворцов."},
        {"word": "смотре́ть (развод мостов)", "gram": "cụm động từ", "meaning": "ngắm mở cầu xoay", "example": "С набережной все смотрят на завораживающий развод мостов."},
        {"word": "любова́ться", "gram": "động từ chưa hoàn thành", "meaning": "thưởng ngoạn đêm trắng", "example": "Гости города любуются романтическими белыми ночами."}
    ],
    22: [
        {"word": "покупа́ть (билет)", "gram": "cụm động từ", "meaning": "mua vé tham quan", "example": "Мы покупаем билет в кассе музея со студенческой скидкой."},
        {"word": "брать (аудиогид)", "gram": "cụm động từ", "meaning": "thuê máy thuyết minh", "example": "Турист берёт персональный аудиогид на русском языке."},
        {"word": "осма́тривать", "gram": "động từ chưa hoàn thành", "meaning": "ngắm nhìn gian trưng bày", "example": "Посетители внимательно осматривают залы с картинами."},
        {"word": "восхища́ться", "gram": "động từ chưa hoàn thành", "meaning": "ngợi ca kiệt tác", "example": "Все восхищаются непревзойдённым талантом живописца."}
    ],
    23: [
        {"word": "бе́гать", "gram": "động từ chưa hoàn thành", "meaning": "chạy bộ rèn sức", "example": "Студенты регулярно бегают на стадионе по вечерам."},
        {"word": "размина́ться", "gram": "động từ chưa hoàn thành", "meaning": "khởi động cơ bắp", "example": "Перед забегом необходимо тщательно размяться."},
        {"word": "уча́ствовать", "gram": "động từ chưa hoàn thành", "meaning": "tham gia giải chạy", "example": "Друзья участвуют в городском благотворительном марафоне."},
        {"word": "побежда́ть", "gram": "động từ chưa hoàn thành", "meaning": "chiến thắng vinh quang", "example": "Благодаря упорству команда побеждает в эстафете."}
    ],
    24: [
        {"word": "обща́ться", "gram": "động từ chưa hoàn thành", "meaning": "trò chuyện gắn bó", "example": "Мы часто общаемся за чашкой чая в уютной обстановке."},
        {"word": "дели́ться", "gram": "động từ chưa hoàn thành", "meaning": "chia sẻ bí mật", "example": "Близкие друзья делятся самыми сокровенными переживаниями."},
        {"word": "помога́ть", "gram": "động từ chưa hoàn thành", "meaning": "giúp đỡ vô tư", "example": "Верный товарищ всегда поможет в любой сложной ситуации."},
        {"word": "ценить", "gram": "động từ chưa hoàn thành", "meaning": "trân quý tình bạn", "example": "Я высоко ценю нашу многолетнюю искреннюю дружбу."}
    ],
    25: [
        {"word": "гото́виться", "gram": "động từ chưa hoàn thành", "meaning": "ôn luyện thi cử", "example": "Студент активно готовится к сдаче сертификационного теста."},
        {"word": "повторя́ть", "gram": "động từ chưa hoàn thành", "meaning": "ôn lại từ vựng", "example": "Перед экзаменом нужно повторить все глаголы движения."},
        {"word": "сдава́ть", "gram": "động từ chưa hoàn thành", "meaning": "hoàn thành bài thi", "example": "Все студенты успешно сдают устную и письменную части."},
        {"word": "горди́ться", "gram": "động từ chưa hoàn thành", "meaning": "tự hào kết quả", "example": "Мы гордимся получением официального сертификата ТРКИ."}
    ]
}

listening_a2 = {}
for i in range(1, 26):
    noun_cat = a2_old.A2_LISTENING_VOCAB[i][0]
    noun_cat["title"] = "Danh từ trong bài nghe (Существительные)"
    verb_cat = {
        "title": "Động từ then chốt (Глаголы)",
        "rows": listening_verbs_a2[i]
    }
    listening_a2[i] = [noun_cat, verb_cat]

# Speaking A2 verbs
speaking_verbs_a2 = {
    1: [
        {"word": "жить", "gram": "động từ chưa hoàn thành", "meaning": "sinh sống lớn lên", "example": "Я родился и всю жизнь живу в древнем городе."},
        {"word": "гуля́ть", "gram": "động từ chưa hoàn thành", "meaning": "đi dạo ven sông", "example": "Моё любимое занятие — гулять по тихим старинным улицам."},
        {"word": "пока́зывать", "gram": "động từ chưa hoàn thành", "meaning": "hướng dẫn khách xem", "example": "Я люблю показывать свой город иностранным гостям."},
        {"word": "горди́ться", "gram": "động từ chưa hoàn thành", "meaning": "tự hào về quê hương", "example": "Жители искренне гордятся богатой историей города."}
    ],
    2: [
        {"word": "просыпа́ться", "gram": "động từ chưa hoàn thành", "meaning": "thức giấc mỗi sáng", "example": "В будние дни я просыпаюсь в семь часов утра по будильнику."},
        {"word": "за́втракать", "gram": "động từ chưa hoàn thành", "meaning": "dùng bữa sáng", "example": "Я завтракаю овсяной кашей и чашкой ароматного кофе."},
        {"word": "учи́ться", "gram": "động từ chưa hoàn thành", "meaning": "học tại trường", "example": "Днём я учусь в университете и посещаю лекции."},
        {"word": "отдыха́ть", "gram": "động từ chưa hoàn thành", "meaning": "thư giãn buổi tối", "example": "Вечером я отдыхаю дома, читаю книги и слушаю музыку."}
    ],
    3: [
        {"word": "плани́ровать", "gram": "động từ chưa hoàn thành", "meaning": "lên kế hoạch đi chơi", "example": "Летом я планирую отправиться в путешествие на Байкал."},
        {"word": "путеше́ствовать", "gram": "động từ chưa hoàn thành", "meaning": "du lịch bốn phương", "example": "Путешествовать на поезде очень интересно и романтично."},
        {"word": "покупа́ть (билеты)", "gram": "cụm động từ", "meaning": "mua vé hành trình", "example": "Мы заранее покупаем билеты на самолёт со скидкой."},
        {"word": "осма́тривать", "gram": "động từ chưa hoàn thành", "meaning": "ngắm cảnh danh thắng", "example": "В поездке туристы осматривают уникальные памятники природы."}
    ],
    4: [
        {"word": "собира́ться", "gram": "động từ chưa hoàn thành", "meaning": "sum họp quây quần", "example": "По воскресеньям вся семья собирается за большим столом."},
        {"word": "помога́ть", "gram": "động từ chưa hoàn thành", "meaning": "tương trợ lẫn nhau", "example": "Родители и дети всегда помогают друг другу во всём."},
        {"word": "храни́ть (традиции)", "gram": "cụm động từ", "meaning": "gìn giữ nếp nhà", "example": "В нашей семье бережно хранят старинные кулинарные рецепты."},
        {"word": "любить", "gram": "động từ chưa hoàn thành", "meaning": "yêu quý ruột thịt", "example": "Я бесконечно люблю и уважаю своих дорогих родителей."}
    ],
    5: [
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "sắm sửa thực phẩm", "example": "Я покупаю свежие овощи и зелень на фермерском рынке."},
        {"word": "выбира́ть", "gram": "động từ chưa hoàn thành", "meaning": "lựa chọn kỹ càng", "example": "Хозяйка тщательно выбирает спелые сезонные фрукты."},
        {"word": "плати́ть", "gram": "động từ chưa hoàn thành", "meaning": "thanh toán tiện lợi", "example": "В супермаркете я всегда плачу бесконтактной картой."},
        {"word": "эконо́мить", "gram": "động từ chưa hoàn thành", "meaning": "tiết kiệm chi tiêu", "example": "Акции и скидочные купоны помогают существенно экономить бюджет."}
    ],
    6: [
        {"word": "ориенти́роваться", "gram": "động từ chưa hoàn thành", "meaning": "định hướng giao thông", "example": "По мобильной карте легко ориентироваться в незнакомом мегаполисе."},
        {"word": "добира́ться", "gram": "động từ chưa hoàn thành", "meaning": "đi đến nơi cần đến", "example": "Быстрее всего добираться до аэропорта на аэроэкспрессе."},
        {"word": "переса́живаться", "gram": "động từ chưa hoàn thành", "meaning": "chuyển đổi tuyến xe", "example": "Вам нужно пересесть на Кольцевую линию метрополитена."},
        {"word": "спра́шивать", "gram": "động từ chưa hoàn thành", "meaning": "hỏi đường người đi", "example": "Не стесняйтесь спрашивать дорогу у вежливых прохожих."}
    ],
    7: [
        {"word": "гото́вить", "gram": "động từ chưa hoàn thành", "meaning": "nấu món truyền thống", "example": "Я умею готовить настоящий украинский борщ с пампушками."},
        {"word": "про́бовать", "gram": "động từ chưa hoàn thành", "meaning": "thử đặc sản Nga", "example": "В поездке туристы с аппетитом пробуют сибирские пельмени."},
        {"word": "зака́зывать", "gram": "động từ chưa hoàn thành", "meaning": "gọi món ẩm thực", "example": "В ресторане русской кухни мы заказываем уху и блины."},
        {"word": "угоща́ть", "gram": "động từ chưa hoàn thành", "meaning": "chiêu đãi bạn bè", "example": "Хозяин щедро угощает гостей домашним ягодным пирогом."}
    ],
    8: [
        {"word": "боле́ть", "gram": "động từ chưa hoàn thành", "meaning": "mắc bệnh cảm cúm", "example": "Осенью многие болеют сезонными простудными заболеваниями."},
        {"word": "обраща́ться (к врачу)", "gram": "cụm động từ", "meaning": "đi khám bác sĩ", "example": "При первых симптомах болезни нужно обращаться к терапевту."},
        {"word": "принима́ть (лекарства)", "gram": "cụm động từ", "meaning": "uống thuốc theo đơn", "example": "Пациент строго принимает прописанные врачом таблетки."},
        {"word": "выздора́вливать", "gram": "động từ chưa hoàn thành", "meaning": "hồi phục sức lực", "example": "Благодаря правильному лечению больной быстро выздоравливает."}
    ],
    9: [
        {"word": "отдыха́ть", "gram": "động từ chưa hoàn thành", "meaning": "nghỉ ngơi nhà vườn", "example": "Летом городские жители любят отдыхать на загородной даче."},
        {"word": "выра́щивать", "gram": "động từ chưa hoàn thành", "meaning": "trồng hoa quả", "example": "Бабушка с любовью выращивает помидоры и клубнику в саду."},
        {"word": "собира́ть (урожай)", "gram": "cụm động từ", "meaning": "thu hoạch mùa màng", "example": "Осенью дачники собирают богатый урожай спелых яблок."},
        {"word": "дыша́ть", "gram": "động từ chưa hoàn thành", "meaning": "hít thở khí trời", "example": "В сосновом лесу приятно дышать чистым целебным воздухом."}
    ],
    10: [
        {"word": "занима́ться (спортом)", "gram": "cụm động từ", "meaning": "tập luyện thể dục", "example": "Для поддержания формы я занимаюсь спортом три раза в неделю."},
        {"word": "бе́гать", "gram": "động từ chưa hoàn thành", "meaning": "chạy bộ giữ dáng", "example": "Утром я бегаю в городском парке по пять километров."},
        {"word": "пла́вать", "gram": "động từ chưa hoàn thành", "meaning": "bơi trong làn nước", "example": "Плавание в бассейне отлично снимает усталость и стресс."},
        {"word": "укрепля́ть", "gram": "động từ chưa hoàn thành", "meaning": "tăng cường đề kháng", "example": "Активный образ жизни укрепляет здоровье и дарит бодрость."}
    ],
    11: [
        {"word": "создава́ть", "gram": "động từ chưa hoàn thành", "meaning": "làm đồ thủ công", "example": "Народные умельцы создают удивительные деревянные игрушки."},
        {"word": "распи́сывать", "gram": "động từ chưa hoàn thành", "meaning": "vẽ hoa văn gốm", "example": "Художницы вручную расписывают посуду гжельскими розами."},
        {"word": "любова́ться", "gram": "động từ chưa hoàn thành", "meaning": "thưởng lãm nghệ thuật", "example": "Посетители выставки любуются золотыми хохломскими ложками."},
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua quà truyền thống", "example": "Я покупаю палехскую лаковую миниатюру в подарок другу."}
    ],
    12: [
        {"word": "приглаша́ть", "gram": "động từ chưa hoàn thành", "meaning": "mời đến ăn cơm", "example": "Русские люди с открытой душой приглашают гостей в дом."},
        {"word": "накрыва́ть (стол)", "gram": "cụm động từ", "meaning": "bày biện bàn tiệc", "example": "К приходу друзей хозяйка накрывает богатый праздничный стол."},
        {"word": "угоща́ть", "gram": "động từ chưa hoàn thành", "meaning": "thết đãi nhiệt tình", "example": "Гостей принято угощать лучшими домашними угощениями."},
        {"word": "благодари́ть", "gram": "động từ chưa hoàn thành", "meaning": "cảm ơn hiếu khách", "example": "Уходя, гости сердечно благодарят хозяев за гостеприимство."}
    ],
    13: [
        {"word": "отправля́ть", "gram": "động từ chưa hoàn thành", "meaning": "gửi thiếp bưu điện", "example": "К праздникам я отправляю поздравительные открытки родным."},
        {"word": "получа́ть", "gram": "động từ chưa hoàn thành", "meaning": "nhận kiện bưu phẩm", "example": "Студент получает долгожданную посылку от родителей."},
        {"word": "зака́зывать", "gram": "động từ chưa hoàn thành", "meaning": "đặt mua trực tuyến", "example": "Сейчас многие заказывают товары с доставкой в пункт выдачи."},
        {"word": "опла́чивать", "gram": "động từ chưa hoàn thành", "meaning": "trả cước bưu chính", "example": "Клиент оплачивает пересылку письма в окне оператора."}
    ],
    14: [
        {"word": "гото́вить", "gram": "động từ chưa hoàn thành", "meaning": "nấu ăn tại nhà", "example": "По праздникам мы любим готовить старинные русские блюда."},
        {"word": "печь", "gram": "động từ chưa hoàn thành", "meaning": "làm bánh truyền thống", "example": "Бабушка печёт румяные пирожки с яблоками и капустой."},
        {"word": "про́бовать", "gram": "động từ chưa hoàn thành", "meaning": "thưởng thức hương vị", "example": "Иностранным студентам очень нравится пробовать квас и морс."},
        {"word": "дели́ться (рецептом)", "gram": "cụm động từ", "meaning": "chia sẻ công thức nấu", "example": "Мама с удовольствием делится семейным рецептом кулича."}
    ],
    15: [
        {"word": "забо́титься", "gram": "động từ chưa hoàn thành", "meaning": "yêu thương săn sóc", "example": "Домашний питомец учит детей заботиться и быть добрыми."},
        {"word": "гуля́ть", "gram": "động từ chưa hoàn thành", "meaning": "dắt chó đi dạo", "example": "Каждое утро я с радостью гуляю с собакой на свежем воздухе."},
        {"word": "корми́ть", "gram": "động từ chưa hoàn thành", "meaning": "cho ăn đúng giờ", "example": "Кошку нужно кормить качественной и полезной пищей."},
        {"word": "игра́ть", "gram": "động từ chưa hoàn thành", "meaning": "vui đùa cùng thú cưng", "example": "Щенок любит весело играть с мячиком на траве."}
    ],
    16: [
        {"word": "носи́ть", "gram": "động từ chưa hoàn thành", "meaning": "mặc đồ giữ ấm", "example": "Зимой в России обязательно носят тёплую шапку и шарф."},
        {"word": "одева́ться", "gram": "động từ chưa hoàn thành", "meaning": "mặc đồ hợp mùa", "example": "При низкой температуре нужно одеваться тепло и удобно."},
        {"word": "выбира́ть", "gram": "động từ chưa hoàn thành", "meaning": "chọn áo khoác tốt", "example": "Я выбираю длинный пуховик с капюшоном от ветра."},
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua giày ủng đông", "example": "Мы покупаем качественные зимние сапоги с мехом."}
    ],
    17: [
        {"word": "чита́ть", "gram": "động từ chưa hoàn thành", "meaning": "đọc tác phẩm kinh điển", "example": "Я люблю читать рассказы русских классиков вслух."},
        {"word": "обсужда́ть", "gram": "động từ chưa hoàn thành", "meaning": "bàn luận về cuốn sách", "example": "На уроках литературы мы обсуждаем поступки главных героев."},
        {"word": "выбира́ть", "gram": "động từ chưa hoàn thành", "meaning": "chọn đầu sách mới", "example": "В библиотеке студент выбирает книгу по рекомендации педагога."},
        {"word": "учи́ться", "gram": "động từ chưa hoàn thành", "meaning": "học qua trang sách", "example": "Хорошая литература учит мудрости, добру и справедливости."}
    ],
    18: [
        {"word": "рабо́тать", "gram": "động từ chưa hoàn thành", "meaning": "công tác chuyên môn", "example": "После окончания института я мечтаю работать в IT-сфере."},
        {"word": "выбира́ть (профессию)", "gram": "cụm động từ", "meaning": "chọn lựa hướng đi", "example": "Молодёжь выбирает современные и востребованные профессии."},
        {"word": "развива́ть (навыки)", "gram": "cụm động từ", "meaning": "rèn luyện kỹ năng", "example": "Для успешной карьеры важно развивать коммуникабельность."},
        {"word": "достига́ть (успеха)", "gram": "cụm động từ", "meaning": "chạm tới thành công", "example": "Трудолюбием и упорством человек достигает высоких результатов."}
    ],
    19: [
        {"word": "пра́здновать", "gram": "động từ chưa hoàn thành", "meaning": "ăn mừng tuần lễ bánh", "example": "Всю неделю народ весело празднует старинную Масленицу."},
        {"word": "печь (блины)", "gram": "cụm động từ", "meaning": "rán bánh mời khách", "example": "Хозяйки пекут тонкие румяные блины с маслом и мёдом."},
        {"word": "угоща́ть", "gram": "động từ chưa hoàn thành", "meaning": "mời ăn ngày hội", "example": "На площади всех бесплатно угощают горячим чаем и блинами."},
        {"word": "проща́ться (с зимой)", "gram": "cụm động từ", "meaning": "tạm biệt giá lạnh", "example": "С хороводами и песнями люди прощаются с долгой зимой."}
    ],
    20: [
        {"word": "поздравля́ть", "gram": "động từ chưa hoàn thành", "meaning": "tri ân ngày chiến thắng", "example": "Мы от всего сердца поздравляем дорогих ветеранов с 9 Мая."},
        {"word": "по́мнить", "gram": "động từ chưa hoàn thành", "meaning": "khắc ghi công ơn", "example": "Потомки свято помнят подвиг солдат, защитивших Родину."},
        {"word": "смотре́ть (парад)", "gram": "cụm động từ", "meaning": "xem duyệt binh kỷ niệm", "example": "Вся страна с гордостью смотрит Парад Победы на Красной площади."},
        {"word": "жела́ть", "gram": "động từ chưa hoàn thành", "meaning": "cầu chúc hòa bình", "example": "Мы искренне желаем чистого неба и вечного мира на всей планете."}
    ],
    21: [
        {"word": "путеше́ствовать", "gram": "động từ chưa hoàn thành", "meaning": "du ngoạn phương bắc", "example": "Я мечтаю путешествовать по Санкт-Петербургу белыми ночами."},
        {"word": "любова́ться", "gram": "động từ chưa hoàn thành", "meaning": "chiêm ngưỡng danh lam", "example": "Туристы с набережных любуются величественной Невой и мостами."},
        {"word": "посеща́ть", "gram": "động từ chưa hoàn thành", "meaning": "thăm quan cung điện", "example": "В Северной столице обязательно нужно посетить Эрмитаж и Петергоф."},
        {"word": "фотографи́ровать", "gram": "động từ chưa hoàn thành", "meaning": "chụp ảnh lưu niệm", "example": "Я фотографирую прекрасные памятники архитектуры и соборы."}
    ],
    22: [
        {"word": "посеща́ть (музей)", "gram": "cụm động từ", "meaning": "thưởng thức bảo tàng", "example": "В выходной день мы посещаем Третьяковскую галерею в Москве."},
        {"word": "осма́тривать", "gram": "động từ chưa hoàn thành", "meaning": "ngắm kiệt tác hội họa", "example": "Посетители подолгу осматривают шедевры русских художников."},
        {"word": "понима́ть (искусство)", "gram": "cụm động từ", "meaning": "thấu hiểu mỹ thuật", "example": "Классическая живопись помогает глубже понимать душу народа."},
        {"word": "вдохновля́ться", "gram": "động từ chưa hoàn thành", "meaning": "tìm thấy nguồn cảm hứng", "example": "Многие вдохновляются прекрасными картинами для своего творчества."}
    ],
    23: [
        {"word": "забо́титься (о здоровье)", "gram": "cụm động từ", "meaning": "giữ gìn thể lực", "example": "Каждый день я забочусь о своём здоровье и самочувствии."},
        {"word": "пита́ться", "gram": "động từ chưa hoàn thành", "meaning": "ăn uống khoa học", "example": "Очень важно регулярно питаться полезной домашней пищей."},
        {"word": "занима́ться", "gram": "động từ chưa hoàn thành", "meaning": "luyện tập thể thao", "example": "Утренняя зарядка и бег заряжают энергией на целый рабочий день."},
        {"word": "отдыха́ть", "gram": "động từ chưa hoàn thành", "meaning": "ngủ đủ giấc", "example": "Организму необходим полноценный сон для восстановления сил."}
    ],
    24: [
        {"word": "дружи́ть", "gram": "động từ chưa hoàn thành", "meaning": "gắn bó keo sơn", "example": "Настоящие друзья дружат годами и никогда не предают."},
        {"word": "подде́рживать", "gram": "động từ chưa hoàn thành", "meaning": "sát cánh bên bạn", "example": "Верный товарищ поддержит в любой радости и в трудную минуту."},
        {"word": "обща́ться", "gram": "động từ chưa hoàn thành", "meaning": "chia sẻ chân thành", "example": "Мы можем часами общаться и обсуждать любые жизненные вопросы."},
        {"word": "ра́доваться", "gram": "động từ chưa hoàn thành", "meaning": "mừng vui cho bạn", "example": "Истинный друг искренне радуется твоим успехам и достижениям."}
    ],
    25: [
        {"word": "достига́ть", "gram": "động từ chưa hoàn thành", "meaning": "chạm mốc cấp độ", "example": "Успешно освоив уровень А2, мы достигаем новой ступени в языке."},
        {"word": "говори́ть", "gram": "động từ chưa hoàn thành", "meaning": "nói chuyện lưu loát", "example": "Теперь я могу свободно говорить и рассуждать на многие темы."},
        {"word": "продолжа́ть", "gram": "động từ chưa hoàn thành", "meaning": "hướng tới B1", "example": "Я планирую уверенно продолжать обучение на уровне B1."},
        {"word": "совершенствова́ть", "gram": "động từ chưa hoàn thành", "meaning": "nâng cao năng lực", "example": "Мы совершенствуем свои навыки общения с носителями языка."}
    ]
}

speaking_a2 = {}
for i in range(1, 26):
    noun_cat = a2_old.A2_SPEAKING_VOCAB[i][0]
    noun_cat["title"] = "Danh từ trong bài nói (Существительные)"
    verb_cat = {
        "title": "Động từ then chốt (Глаголы)",
        "rows": speaking_verbs_a2[i]
    }
    speaking_a2[i] = [noun_cat, verb_cat]

# Writing A2 verbs
writing_verbs_a2 = {
    1: [
        {"word": "поздравля́ть", "gram": "động từ chưa hoàn thành", "meaning": "chúc mừng sinh nhật", "example": "От всей души поздравляю тебя с днём рождения!"},
        {"word": "жела́ть", "gram": "động từ chưa hoàn thành", "meaning": "chúc điều tốt đẹp", "example": "Желаю тебе крепкого здоровья, счастья и больших успехов в учёбе!"},
        {"word": "дари́ть", "gram": "động từ chưa hoàn thành", "meaning": "trao gửi món quà", "example": "Я дарю тебе эту редкую книгу на добрую память."},
        {"word": "наде́яться", "gram": "động từ chưa hoàn thành", "meaning": "hy vọng sớm gặp", "example": "Очень надеюсь скоро увидеться с тобой на каникулах."}
    ],
    2: [
        {"word": "приглаша́ть", "gram": "động từ chưa hoàn thành", "meaning": "mời bạn dự triển lãm", "example": "Приглашаю тебя пойти в субботу на выставку современной живописи."},
        {"word": "смотре́ть", "gram": "động từ chưa hoàn thành", "meaning": "thưởng lãm tranh", "example": "Там можно посмотреть уникальные работы молодых художников."},
        {"word": "встреча́ться", "gram": "động từ chưa hoàn thành", "meaning": "hẹn gặp tại sảnh", "example": "Давай встретимся около входа в музей в двенадцать часов."},
        {"word": "жда́ть", "gram": "động từ chưa hoàn thành", "meaning": "ngóng chờ câu trả lời", "example": "С нетерпением жду твоего скорого ответа по электронной почте."}
    ],
    3: [
        {"word": "люби́ть", "gram": "động từ chưa hoàn thành", "meaning": "yêu thích mùa vàng", "example": "Больше всего в году я люблю золотую русскую осень."},
        {"word": "гуля́ть", "gram": "động từ chưa hoàn thành", "meaning": "tản bộ công viên", "example": "В ясные осенние дни приятно гулять по ковру из жёлтых листьев."},
        {"word": "наслажда́ться", "gram": "động từ chưa hoàn thành", "meaning": "tận hưởng cảnh sắc", "example": "Я наслаждаюсь тишиной, прохладой и красотой увядающей природы."},
        {"word": "меня́ться", "gram": "động từ chưa hoàn thành", "meaning": "chuyển biến thời tiết", "example": "Осенью погода постепенно меняется и воздух становится свежим."}
    ],
    4: [
        {"word": "переезжа́ть", "gram": "động từ chưa hoàn thành", "meaning": "chuyển sang nhà mới", "example": "На прошлой неделе мы переехали в новую двухкомнатную квартиру."},
        {"word": "жить", "gram": "động từ chưa hoàn thành", "meaning": "ở nơi tiện ích", "example": "Теперь мы живём в спокойном районе рядом с большим зелёным парком."},
        {"word": "обустра́ивать", "gram": "động từ chưa hoàn thành", "meaning": "bài trí nội thất", "example": "Мы с удовольствием обустраиваем комнаты и покупаем новую мебель."},
        {"word": "приглаша́ть", "gram": "động từ chưa hoàn thành", "meaning": "mời tân gia", "example": "Приглашаю тебя в гости на новоселье в следующие выходные!"}
    ],
    5: [
        {"word": "проси́ть (о помощи)", "gram": "cụm động từ", "meaning": "nhờ vả xóm giềng", "example": "Обращаюсь к вам с вежливой просьбой о небольшой помощи."},
        {"word": "уезжа́ть", "gram": "động từ chưa hoàn thành", "meaning": "đi vắng vài ngày", "example": "Я уезжаю в командировку на три дня до конца недели."},
        {"word": "полива́ть (цветы)", "gram": "cụm động từ", "meaning": "tưới cây trong phòng", "example": "Пожалуйста, полейте комнатные цветы один раз в среду."},
        {"word": "благодари́ть", "gram": "động từ chưa hoàn thành", "meaning": "cảm ơn trước", "example": "Заранее сердечно благодарю вас за добрососедскую помощь!"}
    ],
    6: [
        {"word": "путеше́ствовать", "gram": "động từ chưa hoàn thành", "meaning": "du ngoạn tàu hỏa", "example": "Путешествовать по железной дороге — это лучший способ увидеть Россию."},
        {"word": "смотре́ть", "gram": "động từ chưa hoàn thành", "meaning": "ngắm cảnh qua cửa sổ", "example": "Из окна купе пассажиры смотрят на бесконечные просторы тайги."},
        {"word": "знако́миться", "gram": "động từ chưa hoàn thành", "meaning": "kết bạn trên tàu", "example": "В пути приятно знакомиться и пить чай с попутчиками."},
        {"word": "вспомина́ть", "gram": "động từ chưa hoàn thành", "meaning": "hoài niệm kỷ niệm đẹp", "example": "Я часто с теплотой вспоминаю эту незабываемую поездку."}
    ],
    7: [
        {"word": "пра́здновать", "gram": "động từ chưa hoàn thành", "meaning": "đón mừng giao thừa", "example": "Новый год мы празднуем в кругу самых близких людей дома."},
        {"word": "украша́ть", "gram": "động từ chưa hoàn thành", "meaning": "trang hoàng cây thông", "example": "Вся семья дружно украшает пушистую ёлку яркими гирляндами."},
        {"word": "гото́вить", "gram": "động từ chưa hoàn thành", "meaning": "làm món salad Nga", "example": "Мама готовит традиционный новогодний салат «Оливье»."},
        {"word": "поздравля́ть", "gram": "động từ chưa hoàn thành", "meaning": "chúc tụng thời khắc giao thừa", "example": "Под бой курантов мы поздравляем друг друга и загадываем желания."}
    ],
    8: [
        {"word": "посеща́ть (театр)", "gram": "cụm động từ", "meaning": "đi xem kịch nghệ", "example": "Вчера мы посетили премьеру спектакля по пьесе классика."},
        {"word": "восхища́ться", "gram": "động từ chưa hoàn thành", "meaning": "khen ngợi diễn xuất", "example": "Зрители искренне восхищались великолепной игрой драматических актёров."},
        {"word": "сове́товать", "gram": "động từ chưa hoàn thành", "meaning": "khuyên nên đi xem", "example": "Я горячо советую всем друзьям посмотреть эту постановку."},
        {"word": "аплоди́ровать", "gram": "động từ chưa hoàn thành", "meaning": "vỗ tay không ngớt", "example": "В финале весь зал долго аплодировал артистам стоя."}
    ],
    9: [
        {"word": "иска́ть (жильё)", "gram": "cụm động từ", "meaning": "tìm thuê phòng", "example": "Студент ищет комнату в тихом районе недалеко от университета."},
        {"word": "снима́ть", "gram": "động từ chưa hoàn thành", "meaning": "thuê trọ dài hạn", "example": "Я хочу снять уютную квартиру с мебелью и стиральной машиной."},
        {"word": "гаранти́ровать", "gram": "động từ chưa hoàn thành", "meaning": "cam kết trật tự", "example": "Гарантирую своевременную оплату, идеальную чистоту и тишину."},
        {"word": "проси́ть (звонить)", "gram": "cụm động từ", "meaning": "nhờ liên lạc lại", "example": "Прошу звонить по указанному номеру телефона после шести вечера."}
    ],
    10: [
        {"word": "занима́ться", "gram": "động từ chưa hoàn thành", "meaning": "luyện tập thể hình", "example": "Я регулярно занимаюсь спортом для поддержания бодрости и здоровья."},
        {"word": "бе́гать", "gram": "động từ chưa hoàn thành", "meaning": "chạy bộ mỗi sáng", "example": "Каждое утро я с удовольствием бегаю по аллеям парка."},
        {"word": "чувствовать", "gram": "động từ chưa hoàn thành", "meaning": "thấy người khỏe khoắn", "example": "После физических упражнений я чувствую прилив сил и энергии."},
        {"word": "сове́товать", "gram": "động từ chưa hoàn thành", "meaning": "khuyên vận động", "example": "Советую каждому человеку уделять спорту хотя бы полчаса в день."}
    ],
    11: [
        {"word": "восхища́ться", "gram": "động từ chưa hoàn thành", "meaning": "trầm trồ đồ thủ công", "example": "Я восхищаюсь мастерством русских резчиков по дереву."},
        {"word": "распи́сывать", "gram": "động từ chưa hoàn thành", "meaning": "vẽ men gốm", "example": "Каждое изделие художники расписывают вручную с душой."},
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua quà lưu niệm", "example": "В подарок друзьям я купил расписной поднос из Жостово."},
        {"word": "беречь", "gram": "động từ chưa hoàn thành", "meaning": "gìn giữ tinh hoa", "example": "Такие уникальные народные ремёсла нужно беречь для потомков."}
    ],
    12: [
        {"word": "благодари́ть", "gram": "động từ chưa hoàn thành", "meaning": "cảm ơn bữa cơm ấm", "example": "Хочу от всего сердца поблагодарить вашу семью за тёплый приём."},
        {"word": "вспомина́ть", "gram": "động từ chưa hoàn thành", "meaning": "nhớ mãi món ăn ngon", "example": "Я с восхищением вспоминаю ваши вкуснейшие домашние пироги."},
        {"word": "цени́ть", "gram": "động từ chưa hoàn thành", "meaning": "trân quý lòng tốt", "example": "Я очень ценю вашу доброту, открытость и заботу об иностранцах."},
        {"word": "приглаша́ть", "gram": "động từ chưa hoàn thành", "meaning": "mời lại nhà mình", "example": "Буду счастлив пригласить вас в гости с ответным визитом!"}
    ],
    13: [
        {"word": "отправля́ть", "gram": "động từ chưa hoàn thành", "meaning": "gửi bưu thiếp từ xa", "example": "Я отправляю тебе эту открытку прямо из заснеженной Москвы."},
        {"word": "гуля́ть", "gram": "động từ chưa hoàn thành", "meaning": "dạo chơi trung tâm", "example": "Сегодня мы гуляли по Красной площади и Александровскому саду."},
        {"word": "передава́ть (привет)", "gram": "cụm động từ", "meaning": "gửi lời chào thăm", "example": "Передавай горячий привет родителям и всем школьным друзьям!"},
        {"word": "скуча́ть", "gram": "động từ chưa hoàn thành", "meaning": "nhớ nhung bạn bè", "example": "Я очень скучаю по нашим душевным разговорам за чашкой чая."}
    ],
    14: [
        {"word": "дели́ться", "gram": "động từ chưa hoàn thành", "meaning": "chia sẻ công thức", "example": "В этом письме я хочу поделиться рецептом русского супа ухи."},
        {"word": "вари́ть", "gram": "động từ chưa hoàn thành", "meaning": "nấu nước dùng cá", "example": "Сначала нужно сварить ароматный рыбный бульон с луком и морковью."},
        {"word": "добавля́ть", "gram": "động từ chưa hoàn thành", "meaning": "cho thêm gia vị", "example": "В конце варки обязательно добавьте лавровый лист и свежий укроп."},
        {"word": "угоща́ть", "gram": "động từ chưa hoàn thành", "meaning": "mời cả nhà cùng ăn", "example": "Приготовьте этот суп и угостите своих близких вкусным обедом!"}
    ],
    15: [
        {"word": "забо́титься", "gram": "động từ chưa hoàn thành", "meaning": "chăm sóc chó cưng", "example": "У меня дома живёт умная собака, о которой я с радостью забочусь."},
        {"word": "гуля́ть", "gram": "động từ chưa hoàn thành", "meaning": "đi dạo mỗi sáng chiều", "example": "Каждый день мы гуляем в парке и играем на свежем воздухе."},
        {"word": "понима́ть", "gram": "động từ chưa hoàn thành", "meaning": "hiểu lệnh của chủ", "example": "Мой преданный пёс прекрасно понимает команды с одного взгляда."},
        {"word": "люби́ть", "gram": "động từ chưa hoàn thành", "meaning": "yêu quý như bạn", "example": "Четвероногий питомец стал настоящим и верным членом нашей семьи."}
    ],
    16: [
        {"word": "сове́товать", "gram": "động từ chưa hoàn thành", "meaning": "khuyên bạn chuẩn bị đồ", "example": "Советую тебе заранее купить тёплые вещи перед поездкой в Россию."},
        {"word": "выбира́ть", "gram": "động từ chưa hoàn thành", "meaning": "chọn áo phao lông", "example": "Выбирай длинный пуховик из плотной ветрозащитной ткани."},
        {"word": "покупа́ть", "gram": "động từ chưa hoàn thành", "meaning": "mua ủng ấm chống trơn", "example": "Обязательно купи непромокаемую зимнюю обувь на толстой подошве."},
        {"word": "носи́ть", "gram": "động từ chưa hoàn thành", "meaning": "mặc giữ nhiệt tốt", "example": "В морозную погоду комфортно носить термобельё и шерстяные носки."}
    ],
    17: [
        {"word": "чита́ть", "gram": "động từ chưa hoàn thành", "meaning": "đọc tác phẩm sâu sắc", "example": "Недавно я прочитал трогательную повесть о верности и благородстве."},
        {"word": "рекомендова́ть", "gram": "động từ chưa hoàn thành", "meaning": "khuyên đọc truyện này", "example": "Я искренне рекомендую эту замечательную книгу всем моим друзьям."},
        {"word": "заду́мываться", "gram": "động từ chưa hoàn thành", "meaning": "suy ngẫm cuộc đời", "example": "Это глубокое произведение заставляет задуматься о ценности дружбы."},
        {"word": "вдохновля́ть", "gram": "động từ chưa hoàn thành", "meaning": "truyền cảm hứng sống", "example": "Поступки главных героев вдохновляют на добрые и смелые дела."}
    ],
    18: [
        {"word": "составля́ть (резюме)", "gram": "cụm động từ", "meaning": "viết CV xin việc", "example": "Я аккуратно составляю резюме для прохождения летней стажировки."},
        {"word": "ука́зывать", "gram": "động từ chưa hoàn thành", "meaning": "nêu rõ bằng cấp", "example": "В анкете соискатель указывает уровень владения языками и навыки."},
        {"word": "стреми́ться", "gram": "động từ chưa hoàn thành", "meaning": "học hỏi phát triển", "example": "Я стремлюсь получить ценный практический опыт в международной фирме."},
        {"word": "наде́яться", "gram": "động từ chưa hoàn thành", "meaning": "hy vọng nhận việc", "example": "Надеюсь на приглашение на собеседование в удобное для вас время."}
    ],
    19: [
        {"word": "опи́сывать", "gram": "động từ chưa hoàn thành", "meaning": "miêu tả lễ hội", "example": "В статье я описываю старинные русские обычаи встречи весны."},
        {"word": "печь (блины)", "gram": "cụm động từ", "meaning": "làm bánh thơm giòn", "example": "В масленичные дни в каждом доме пекут ароматные круглые блины."},
        {"word": "собира́ться", "gram": "động từ chưa hoàn thành", "meaning": "tụ họp xóm giềng", "example": "Люди выходят на городские площади, чтобы петь песни и водить хороводы."},
        {"word": "жела́ть", "gram": "động từ chưa hoàn thành", "meaning": "chúc điều an lành", "example": "В Прощёное воскресенье принято просить прощения и желать добра."}
    ],
    20: [
        {"word": "посвяща́ть", "gram": "động từ chưa hoàn thành", "meaning": "dành bài viết tri ân", "example": "Своё эссе я посвящаю подвигу защитников Родины в годы войны."},
        {"word": "по́мнить", "gram": "động từ chưa hoàn thành", "meaning": "nhớ mãi công ơn", "example": "Наш долг — свято помнить имена павших воинов и ценить мирную жизнь."},
        {"word": "передава́ть (память)", "gram": "cụm động từ", "meaning": "truyền lại ngọn lửa", "example": "Мы должны бережно передавать память о Победе следующим поколениям."},
        {"word": "горди́ться", "gram": "động từ chưa hoàn thành", "meaning": "tự hào lịch sử", "example": "Каждая российская семья гордится своими дедами и прадедами."}
    ],
    21: [
        {"word": "дели́ться (впечатлениями)", "gram": "cụm động từ", "meaning": "kể lại ấn tượng", "example": "Спешу поделиться своими незабываемыми впечатлениями от поездки в Петербург."},
        {"word": "восхища́ться", "gram": "động từ chưa hoàn thành", "meaning": "say đắm cảnh đẹp", "example": "Я был по-настоящему очарован красотой гранитных набережных и каналов."},
        {"word": "гуля́ть", "gram": "động từ chưa hoàn thành", "meaning": "tản bộ đêm trắng", "example": "Мы часами гуляли по Невскому проспекту в романтичные белые ночи."},
        {"word": "сове́товать", "gram": "động từ chưa hoàn thành", "meaning": "khuyên nên ghé thăm", "example": "Обязательно посети этот великолепный город при первой возможности!"}
    ],
    22: [
        {"word": "посеща́ть (галерею)", "gram": "cụm động từ", "meaning": "thăm phòng triển lãm", "example": "Вчера я посетил персональную выставку известного пейзажиста."},
        {"word": "осма́тривать", "gram": "động từ chưa hoàn thành", "meaning": "chiêm ngưỡng đường nét", "example": "Я внимательно осматривал каждую картину, поражаясь игре красок и света."},
        {"word": "чувствовать", "gram": "động từ chưa hoàn thành", "meaning": "rung cảm tâm hồn", "example": "Полотна передают дыхание живой природы и дарят умиротворение."},
        {"word": "рекомендова́ть", "gram": "động từ chưa hoàn thành", "meaning": "giới thiệu người xem", "example": "Рекомендую всем ценителям прекрасного не пропустить эту экспозицию."}
    ],
    23: [
        {"word": "рассужда́ть", "gram": "động từ chưa hoàn thành", "meaning": "nghị luận lối sống", "example": "В своём сочинении я рассуждаю о важности здорового образа жизни."},
        {"word": "соблюда́ть (баланс)", "gram": "cụm động từ", "meaning": "cân bằng học nghỉ", "example": "Очень важно соблюдать разумный баланс между учёбой и отдыхом."},
        {"word": "занима́ться", "gram": "động từ chưa hoàn thành", "meaning": "vận động thường xuyên", "example": "Регулярная физическая активность помогает сохранять ясность ума."},
        {"word": "призыва́ть", "gram": "động từ chưa hoàn thành", "meaning": "kêu gọi sống lành mạnh", "example": "Я призываю сверстников больше гулять на свежем воздухе и заниматься спортом."}
    ],
    24: [
        {"word": "рассужда́ть", "gram": "động từ chưa hoàn thành", "meaning": "bàn về tình bạn", "example": "Что такое истинная дружба? Об этом я размышляю в своём эссе."},
        {"word": "подде́рживать", "gram": "động từ chưa hoàn thành", "meaning": "sát cánh động viên", "example": "Верный друг бескорыстно поддерживает тебя в любых жизненных испытаниях."},
        {"word": "дове́рять", "gram": "động từ chưa hoàn thành", "meaning": "tin cậy tâm giao", "example": "Настоящая дружба строится на искренности, уважении и взаимном доверии."},
        {"word": "беречь", "gram": "động từ chưa hoàn thành", "meaning": "gìn giữ mối quan hệ", "example": "Крепкую дружбу нужно беречь как самое драгоценное сокровище в жизни."}
    ],
    25: [
        {"word": "подводи́ть (итоги)", "gram": "cụm động từ", "meaning": "tổng kết chặng đường", "example": "Завершая курс А2, я с гордостью подвожу итоги своей упорной работы."},
        {"word": "осва́ивать", "gram": "động từ chưa hoàn thành", "meaning": "lĩnh hội kiến thức mới", "example": "За это время я освоил множество сложных грамматических конструкций."},
        {"word": "благодари́ть", "gram": "động từ chưa hoàn thành", "meaning": "tri ân thầy cô", "example": "Искренне благодарю преподавателей за терпение, вдохновение и поддержку."},
        {"word": "стреми́ться", "gram": "động từ chưa hoàn thành", "meaning": "chinh phục cấp độ B1", "example": "Я с энтузиазмом стремлюсь к новым горизонтам и покорению уровня B1!"}
    ]
}

writing_a2 = {}
for i in range(1, 26):
    noun_cat = a2_old.A2_WRITING_VOCAB[i][0]
    noun_cat["title"] = "Danh từ trong bài viết (Существительные)"
    verb_cat = {
        "title": "Động từ then chốt (Глаголы)",
        "rows": writing_verbs_a2[i]
    }
    writing_a2[i] = [noun_cat, verb_cat]

a2_master = {
    'reading': reading_a2,
    'listening': listening_a2,
    'speaking': speaking_a2,
    'writing': writing_a2
}

with open('scripts/vocab_master_a2.json', 'w', encoding='utf-8') as f:
    json.dump(a2_master, f, ensure_ascii=False, indent=2)

print("SUCCESS: scripts/vocab_master_a2.json built successfully!")
print("Counts: Reading:", len(reading_a2), "Listening:", len(listening_a2), "Speaking:", len(speaking_a2), "Writing:", len(writing_a2))
