# -*- coding: utf-8 -*-
"""
Generator for A2 Listening Lessons 4 to 25.
"""

A2_LISTENING_4_TO_25 = [
    {
        "id": 4,
        "level": "A2",
        "skill": "listening",
        "title": "Покупка железнодорожного билета в кассе",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Mua vé tàu hỏa đi du lịch tại ga đường sắt",
        "dialogue": [
            {"speaker": "Пассажир", "ru": "Здравствуйте! Мне нужен один билет до Санкт-Петербурга на пятницу вечер.", "vi": "Xin chào! Tôi cần một vé đi Saint Petersburg vào tối thứ Sáu."},
            {"speaker": "Кассир", "ru": "Добрый день. На какой поезд вы хотите: скорый «Сапсан» или ночной поезд?", "vi": "Chào bạn. Bạn muốn đi tàu nào: tàu cao tốc Sapsan hay tàu đêm?"},
            {"speaker": "Пассажир", "ru": "Лучше ночной, чтобы поспать. Какие места есть в купе?", "vi": "Tốt hơn là tàu đêm để tôi ngủ được. Có chỗ nào ở khoang coupe không ạ?"},
            {"speaker": "Кассир", "ru": "Есть нижняя полка в вагоне номер шесть. Отправление в 23:55 с Ленинградского вокзала.", "vi": "Có một giường tầng dưới ở toa số 6. Tàu khởi hành lúc 23:55 từ ga Leningradsky."},
            {"speaker": "Пассажир", "ru": "Отлично, беру. Сколько стоит билет?", "vi": "Tuyệt vời, tôi lấy chỗ đó. Vé giá bao nhiêu ạ?"},
            {"speaker": "Кассир", "ru": "Три тысячи восемьсот рублей с постельным бельём. Ваш паспорт, пожалуйста.", "vi": "3800 rúp đã bao gồm ga gối. Cho tôi xin hộ chiếu của bạn."},
            {"speaker": "Пассажир", "ru": "Вот паспорт и банковская карта. Спасибо!", "vi": "Đây là hộ chiếu và thẻ ngân hàng. Cảm ơn chị!"}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "Поезд отправляется в двадцать три пятьдесят пять с вокзала.",
            "hint": "Chú ý số đếm giờ khởi hành của chuyến tàu đêm."
        },
        "questions": [
            {
                "id": "a2_l4_q1",
                "prompt": "Какой тип поезда и полку выбрал пассажир?",
                "hint": "Xem lựa chọn của hành khách trong hội thoại.",
                "options": [
                    {"key": "A", "text": "Дневной «Сапсан» и сидячее место"},
                    {"key": "B", "text": "Ночной поезд и нижнюю полку в купе"},
                    {"key": "C", "text": "Верхнюю полку в плацкарте"},
                    {"key": "D", "text": "Автобусный рейс"}
                ],
                "ans": "B",
                "explain": "Пассажир выбрал ночной поезд и нижнюю полку в купейном вагоне."
            }
        ]
    },
    {
        "id": 5,
        "level": "A2",
        "skill": "listening",
        "title": "Аренда жилья: Разговор с хозяйкой квартиры",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Gọi điện thoại hỏi thuê phòng trọ sinh viên",
        "dialogue": [
            {"speaker": "Студент", "ru": "Алло, здравствуйте! Я звоню по объявлению об аренде однокомнатной квартиры на проспекте Мира.", "vi": "A lô, xin chào cô! Cháu gọi theo tin đăng thuê căn hộ 1 phòng trên đại lộ Mira."},
            {"speaker": "Хозяйка", "ru": "Здравствуйте! Да, квартира пока свободна. Вы студент?", "vi": "Chào cháu! Đúng rồi, căn hộ vẫn đang trống. Cháu là sinh viên à?"},
            {"speaker": "Студент", "ru": "Да, учусь на втором курсе университета. Скажите, пожалуйста, сколько минут пешком до метро?", "vi": "Vâng, cháu học năm thứ hai đại học. Xin cô cho biết đi bộ ra ga metro mất mấy phút ạ?"},
            {"speaker": "Хозяйка", "ru": "Буквально пять минут спокойным шагом. В квартире есть стиральная машина, холодильник и Wi-Fi.", "vi": "Đúng 5 phút đi bộ thong thả. Trong nhà có máy giặt, tủ lạnh và Wi-Fi."},
            {"speaker": "Студент", "ru": "А коммунальные услуги включены в стоимость?", "vi": "Thế tiền dịch vụ điện nước đã bao gồm trong giá thuê chưa ạ?"},
            {"speaker": "Хозяйка", "ru": "Аренда 35 тысяч рублей, а свет и воду вы оплачиваете по счётчикам.", "vi": "Giá thuê 35 nghìn rúp, còn tiền điện nước cháu tự trả theo đồng hồ đo."}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "До станции метро можно дойти пешком за пять минут.",
            "hint": "Chú ý cự ly thời gian đi bộ ra ga tàu điện ngầm."
        },
        "questions": [
            {
                "id": "a2_l5_q1",
                "prompt": "Как оплачиваются коммунальные услуги за квартиру?",
                "hint": "Xem câu trả lời của chủ nhà ở cuối đoạn.",
                "options": [
                    {"key": "A", "text": "Они полностью бесплатны"},
                    {"key": "B", "text": "Оплачиваются отдельно по показаниям счётчиков"},
                    {"key": "C", "text": "Входят в фиксированную стоимость"},
                    {"key": "D", "text": "Их оплачивает университет"}
                ],
                "ans": "B",
                "explain": "Хозяйка пояснила: свет и вода оплачиваются жильцом отдельно по счётчикам."
            }
        ]
    },
    {
        "id": 6,
        "level": "A2",
        "skill": "listening",
        "title": "Покупка зимней одежды в торговом центре",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Mua áo khoác lông vũ tại trung tâm thương mại",
        "dialogue": [
            {"speaker": "Консультант", "ru": "Здравствуйте! Чем я могу вам помочь?", "vi": "Xin chào quý khách! Tôi có thể giúp gì cho bạn?"},
            {"speaker": "Покупатель", "ru": "Добрый день. Я ищу тёплый зимний пуховик. Мне нужен сорок восьмой размер.", "vi": "Chào bạn. Tôi đang tìm một chiếc áo phao lông vũ ấm áp mùa đông. Tôi cần size 48."},
            {"speaker": "Консультант", "ru": "Посмотрите эту удлинённую модель тёмно-синего цвета. Наполнитель — натуральный пух, выдерживает до -25 градусов.", "vi": "Bạn hãy xem mẫu dáng dài màu xanh thẫm này nhé. Lớp lót lông vũ tự nhiên, chịu được nhiệt độ tới -25 độ C."},
            {"speaker": "Покупатель", "ru": "Очень красивый. Где я могу его примерить?", "vi": "Đẹp quá. Tôi có thể thử áo ở đâu?"},
            {"speaker": "Консультант", "ru": "Примерочные кабины прямо по коридору направо. Пожалуйста, проходите.", "vi": "Phòng thử đồ ở ngay hành lang bên phải. Mời bạn vào trong."}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "Примерочные кабины находятся прямо по коридору направо.",
            "hint": "Chú ý vị trí chỉ dẫn phòng thay đồ của nhân viên."
        },
        "questions": [
            {
                "id": "a2_l6_q1",
                "prompt": "Какой размер пуховика нужен покупателю?",
                "hint": "Xem số đo kích cỡ trong lời nói của khách hàng.",
                "options": [
                    {"key": "A", "text": "42-й"},
                    {"key": "B", "text": "48-й"},
                    {"key": "C", "text": "54-й"},
                    {"key": "D", "text": "Универсальный"}
                ],
                "ans": "B",
                "explain": "Покупатель сообщил консультанту: «Мне нужен сорок восьмой размер»."
            }
        ]
    },
    {
        "id": 7,
        "level": "A2",
        "skill": "listening",
        "title": "Разговор на почте: Отправка заказного письма",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Gửi thư bảo đảm và mua phong bì tại bưu cục",
        "dialogue": [
            {"speaker": "Посетитель", "ru": "Здравствуйте! Мне нужно отправить заказное письмо с уведомлением о вручении.", "vi": "Xin chào! Tôi cần gửi một lá thư bảo đảm có báo phát."},
            {"speaker": "Оператор", "ru": "Пожалуйста, положите письмо на весы. Адрес получателя и индекс заполнили?", "vi": "Mời bạn đặt thư lên cân. Bạn đã điền địa chỉ người nhận và mã bưu điện chưa?"},
            {"speaker": "Посетитель", "ru": "Да, индекс 190000, Санкт-Петербург. Вот бланк.", "vi": "Vâng, mã 190000, Saint Petersburg. Đây là tờ khai ạ."},
            {"speaker": "Оператор", "ru": "Отлично. Стоимость отправления — 140 рублей. Вот ваш чек с трек-номером для отслеживания.", "vi": "Tốt rồi. Cước phí gửi thư là 140 rúp. Đây là hóa đơn có mã vận đơn để bạn tra cứu."}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "Вот ваш чек с трек-номером для отслеживания письма.",
            "hint": "Lắng nghe từ ngữ về hóa đơn bưu chính."
        },
        "questions": [
            {
                "id": "a2_l7_q1",
                "prompt": "Какую сумму заплатил посетитель за отправку письма?",
                "hint": "Xem cước phí trong lời của nhân viên bưu cục.",
                "options": [
                    {"key": "A", "text": "50 рублей"},
                    {"key": "B", "text": "140 рублей"},
                    {"key": "C", "text": "300 рублей"},
                    {"key": "D", "text": "Бесплатно"}
                ],
                "ans": "B",
                "explain": "Оператор озвучил стоимость заказного отправления: 140 рублей."
            }
        ]
    },
    {
        "id": 8,
        "level": "A2",
        "skill": "listening",
        "title": "Запись на приём к врачу-стоматологу",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Gọi điện tới phòng khám nha khoa đặt lịch khám",
        "dialogue": [
            {"speaker": "Администратор", "ru": "Стоматологическая клиника «Здоровье», добрый день! Чем могу помочь?", "vi": "Phòng khám nha khoa 'Sức Khỏe', xin chào! Tôi có thể giúp gì ạ?"},
            {"speaker": "Пациент", "ru": "Здравствуйте! У меня со вчерашнего дня сильно болит зуб. Можно записаться к врачу на сегодня?", "vi": "Chào cô! Từ hôm qua cháu bị đau răng dữ dội. Cháu có thể đặt lịch khám hôm nay được không ạ?"},
            {"speaker": "Администратор", "ru": "С острой болью доктор Смирнов сможет принять вас сегодня в 16:30. Вам удобно?", "vi": "Với trường hợp đau cấp, bác sĩ Smirnov có thể nhận khám vào 16:30 hôm nay. Giờ đó có tiện cho bạn không?"},
            {"speaker": "Пациент", "ru": "Да, замечательно! Запишите меня, пожалуйста.", "vi": "Dạ rất tiện ạ! Cô ghi tên cho cháu với nhé."}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "Доктор сможет принять вас сегодня в шестнадцать тридцать.",
            "hint": "Chú ý mốc thời gian giờ hẹn khám bệnh."
        },
        "questions": [
            {
                "id": "a2_l8_q1",
                "prompt": "На какое время пациента записали к врачу?",
                "hint": "Xem giờ hẹn trong câu nói của lễ tân.",
                "options": [
                    {"key": "A", "text": "На утро в 9:00"},
                    {"key": "B", "text": "На сегодня в 16:30"},
                    {"key": "C", "text": "На завтра в 14:00"},
                    {"key": "D", "text": "Через неделю"}
                ],
                "ans": "B",
                "explain": "Администратор назначил приём на 16:30 текущего дня."
            }
        ]
    },
    {
        "id": 9,
        "level": "A2",
        "skill": "listening",
        "title": "В фитнес-клубе: Оформление абонемента",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Đăng ký thẻ hội viên câu lạc bộ thể hình",
        "dialogue": [
            {"speaker": "Клиент", "ru": "Здравствуйте! Я хочу узнать об условиях абонемента в ваш тренажёрный зал.", "vi": "Xin chào! Tôi muốn tìm hiểu về các điều kiện gói tập gym bên mình."},
            {"speaker": "Менеджер", "ru": "Добрый день! У нас есть абонементы на 3 месяца, полгода и год. В годовую карту входят тренажёрный зал, бассейн и сауна.", "vi": "Chào bạn! Bên mình có gói 3 tháng, 6 tháng và 1 năm. Gói năm bao gồm phòng gym, hồ bơi và phòng xông hơi."},
            {"speaker": "Клиент", "ru": "Отлично, бассейн для меня очень важен. А вводное занятие с тренером есть?", "vi": "Tuyệt quá, bể bơi rất quan trọng với tôi. Có buổi tập kèm định hướng với HLV không?"},
            {"speaker": "Менеджер", "ru": "Да, в карту включены два вводных персональных инструктажа.", "vi": "Có chứ, thẻ tập đã bao gồm 2 buổi hướng dẫn cá nhân cùng huấn luyện viên."}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "В годовую карту входят тренажёрный зал и бассейн.",
            "hint": "Chú ý tiện ích bao gồm trong gói tập."
        },
        "questions": [
            {
                "id": "a2_l9_q1",
                "prompt": "Что входит в годовую карту фитнес-клуба?",
                "hint": "Xem các tiện ích được quản lý liệt kê.",
                "options": [
                    {"key": "A", "text": "Только утренний зал"},
                    {"key": "B", "text": "Тренажёрный зал, бассейн и сауна"},
                    {"key": "C", "text": "Только теннисный корт"},
                    {"key": "D", "text": "Массаж и ресторан"}
                ],
                "ans": "B",
                "explain": "В годовой абонемент включены зал, бассейн, сауна и персональные инструктажи."
            }
        ]
    },
    {
        "id": 10,
        "level": "A2",
        "skill": "listening",
        "title": "Покупка экскурсионного тура в турагентстве",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Đặt tour du lịch tham quan Vành đai Vàng",
        "dialogue": [
            {"speaker": "Турист", "ru": "Здравствуйте! Мы хотим поехать на экскурсию по Золотому кольцу на майские праздники.", "vi": "Xin chào! Chúng tôi muốn đi tour du lịch Vành đai Vàng vào dịp nghỉ lễ tháng Năm."},
            {"speaker": "Агент", "ru": "Прекрасный выбор! Могу предложить трёхдневный автобусный тур: Владимир — Суздаль — Ярославль.", "vi": "Lựa chọn tuyệt vời! Tôi xin gợi ý tour ô tô 3 ngày: Vladimir - Suzdal - Yaroslavl."},
            {"speaker": "Турист", "ru": "Что включено в стоимость тура?", "vi": "Giá tour đã bao gồm những gì ạ?"},
            {"speaker": "Агент", "ru": "Проезд на комфортабельном автобусе, проживание в гостинице, завтраки и все экскурсии с гидом.", "vi": "Xe du lịch tiện nghi, phòng khách sạn, các bữa sáng và toàn bộ chương trình hướng dẫn viên."}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "Трёхдневный тур включает проживание в гостинице и завтраки.",
            "hint": "Nghe về quyền lợi chỗ ở và ăn sáng trong chuyến đi."
        },
        "questions": [
            {
                "id": "a2_l10_q1",
                "prompt": "Какие города входят в предложенный маршрут тура?",
                "hint": "Xem 3 thành phố cổ được đại lý du lịch đề xuất.",
                "options": [
                    {"key": "A", "text": "Владимир, Суздаль, Ярославль"},
                    {"key": "B", "text": "Сочи и Казань"},
                    {"key": "C", "text": "Владивосток и Хабаровск"},
                    {"key": "D", "text": "Только Санкт-Петербург"}
                ],
                "ans": "A",
                "explain": "Турагент предложил популярный классический маршрут: Владимир — Суздаль — Ярославль."
            }
        ]
    },
    {
        "id": 11,
        "level": "A2",
        "skill": "listening",
        "title": "Встреча друга в аэропорту Шереметьево",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Đón bạn tại sảnh nhà ga quốc tế sân bay",
        "dialogue": [
            {"speaker": "Антон", "ru": "Алло, Тхань! Ты уже прошёл паспортный контроль и получил багаж?", "vi": "A lô, Thành ơi! Cậu đã qua cửa kiểm tra hộ chiếu và lấy hành lý chưa?"},
            {"speaker": "Тхань", "ru": "Привет, Антон! Да, только что забрал свой чемодан. Выхожу в зал прилёта.", "vi": "Chào Anton! Ừ, mình vừa lấy vali xong. Đang bước ra sảnh đến đây."},
            {"speaker": "Антон", "ru": "Я стою возле информационного табло у выхода номер три. У меня в руках жёлтый шарф.", "vi": "Mình đang đứng cạnh bảng thông tin điện tử ở cửa số 3. Tay cầm khăn quàng màu vàng nhé."},
            {"speaker": "Тхань", "ru": "Вижу тебя! Иду к тебе!", "vi": "Mình nhìn thấy cậu rồi! Đang qua đây!"}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "Я стою возле информационного табло у выхода номер три.",
            "hint": "Vị trí cột mốc hẹn gặp nhau tại sảnh sân bay."
        },
        "questions": [
            {
                "id": "a2_l11_q1",
                "prompt": "По какому предмету Тхань смог быстро узнать Антона?",
                "hint": "Xem chi tiết đồ vật trên tay bạn ở sảnh đón.",
                "options": [
                    {"key": "A", "text": "По жёлтому шарфу"},
                    {"key": "B", "text": "По букету цветов"},
                    {"key": "C", "text": "По чёрному зонту"},
                    {"key": "D", "text": "По большой табличке"}
                ],
                "ans": "A",
                "explain": "Антон сообщил, что держит в руках жёлтый шарф для лёгкого ориентирования."
            }
        ]
    },
    {
        "id": 12,
        "level": "A2",
        "skill": "listening",
        "title": "Заказ такси и разговор с водителем",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Đi xe taxi và hướng dẫn đường đi tới trường",
        "dialogue": [
            {"speaker": "Водитель", "ru": "Здравствуйте! Вы заказывали такси до главного корпуса университета?", "vi": "Xin chào! Bạn vừa đặt chuyến xe tới tòa nhà chính của trường đại học phải không?"},
            {"speaker": "Пассажир", "ru": "Добрый день! Да, всё верно. Скажите, пожалуйста, на дорогах сейчас есть пробки?", "vi": "Chào anh! Dạ đúng rồi. Bác tài cho em hỏi trên đường bây giờ có bị tắc xe không ạ?"},
            {"speaker": "Водитель", "ru": "На Садовом кольце плотное движение, но навигатор ведёт по набережной. Доедем минут за двадцать.", "vi": "Đường Vành đai Sadovoye xe đông, nhưng định vị đang dẫn vòng men theo bờ kè. Tầm 20 phút là tới nơi."},
            {"speaker": "Пассажир", "ru": "Отлично, у меня лекция начинается через полчаса, как раз успеваю.", "vi": "Thế thì tốt quá, nửa tiếng nữa em vào giờ giảng, vừa kịp giờ luôn ạ."}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "Навигатор ведёт по набережной, доедем минут за двадцать.",
            "hint": "Chú ý đường đi của xe qua định vị vệ tinh."
        },
        "questions": [
            {
                "id": "a2_l12_q1",
                "prompt": "Сколько времени займёт поездка по расчёту навигатора?",
                "hint": "Xem thời gian di chuyển trong câu của tài xế.",
                "options": [
                    {"key": "A", "text": "Около 20 минут"},
                    {"key": "B", "text": "Больше двух часов"},
                    {"key": "C", "text": "Пять минут"},
                    {"key": "D", "text": "Ровно один час"}
                ],
                "ans": "A",
                "explain": "Водитель уточнил, что маршрут по набережной займёт около 20 минут."
            }
        ]
    },
    {
        "id": 13,
        "level": "A2",
        "skill": "listening",
        "title": "В парикмахерской: Обсуждение стрижки",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Trao đổi kiểu tóc tại tiệm cắt tóc",
        "dialogue": [
            {"speaker": "Мастер", "ru": "Здравствуйте! Присаживайтесь в кресло. Как будем стричься?", "vi": "Xin chào! Mời bạn ngồi vào ghế. Hôm nay chúng ta cắt kiểu gì đây bạn?"},
            {"speaker": "Клиент", "ru": "Здравствуйте! Сделайте, пожалуйста, покороче по бокам и сзади, а сверху оставьте немного длины.", "vi": "Chào anh! Cắt giùm em ngắn hai bên mai và sau gáy, còn phần trên đỉnh đầu để dài một chút nhé."},
            {"speaker": "Мастер", "ru": "Хорошо, виски прямые или косые делать?", "vi": "Được rồi, phần thái dương cắt vuông thẳng hay vát xéo?"},
            {"speaker": "Клиент", "ru": "Прямые, пожалуйста. И помойте голову перед стрижкой.", "vi": "Cắt thẳng ạ. Và nhờ anh gội đầu cho em trước khi cắt nhé."}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "Сделайте покороче по бокам, а сверху оставьте длину.",
            "hint": "Lắng nghe yêu cầu của khách hàng về độ dài tóc."
        },
        "questions": [
            {
                "id": "a2_l13_q1",
                "prompt": "Какую стрижку попросил клиент?",
                "hint": "Xem yêu cầu về hai bên và phần trên đỉnh đầu.",
                "options": [
                    {"key": "A", "text": "Побрить наголо полностью"},
                    {"key": "B", "text": "Покороче по бокам и сзади, оставить длину сверху"},
                    {"key": "C", "text": "Только покрасить волосы"},
                    {"key": "D", "text": "Ничего не стричь"}
                ],
                "ans": "B",
                "explain": "Клиент попросил классическую аккуратную форму: короче по бокам и длиннее сверху."
            }
        ]
    },
    {
        "id": 14,
        "level": "A2",
        "skill": "listening",
        "title": "Поздравление друга по телефону с Новым годом",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Gọi điện thoại gửi lời chúc Tết bạn thân",
        "dialogue": [
            {"speaker": "Максим", "ru": "Алло, Тхань! С наступающим Новым годом тебя и всю твою семью!", "vi": "A lô Thành ơi! Chúc mừng năm mới sắp đến với cậu và cả gia đình nhé!"},
            {"speaker": "Тхань", "ru": "Привет, Максим! Огромное спасибо! И тебя от всей души поздравляю!", "vi": "Chào Maksim! Cảm ơn cậu rất nhiều! Mình cũng chúc mừng năm mới cậu hết lòng!"},
            {"speaker": "Максим", "ru": "Желаю тебе крепкого здоровья, удачи в учёбе и успешной сдачи всех экзаменов на отлично!", "vi": "Chúc cậu dồi dào sức khỏe, may mắn trong học tập và thi cử toàn điểm xuất sắc nhé!"},
            {"speaker": "Тхань", "ru": "Спасибо за тёплые слова! Встретимся на каникулах!", "vi": "Cảm ơn những lời chúc ấm áp của cậu! Hẹn gặp nhau đợt nghỉ này nhé!"}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "Желаю тебе крепкого здоровья и успешной сдачи экзаменов.",
            "hint": "Lời chúc tốt lành quen thuộc dịp năm mới."
        },
        "questions": [
            {
                "id": "a2_l14_q1",
                "prompt": "Чего пожелал Максим своему другу?",
                "hint": "Xem các điều ước được nhắc tới trong lời chúc.",
                "options": [
                    {"key": "A", "text": "Здоровья, удачи в учёбе и отличных оценок на экзаменах"},
                    {"key": "B", "text": "Купить новый автомобиль"},
                    {"key": "C", "text": "Уехать в другую страну"},
                    {"key": "D", "text": "Больше спать"}
                ],
                "ans": "A",
                "explain": "Максим пожелал здоровья и учебных успехов в новом году."
            }
        ]
    },
    {
        "id": 15,
        "level": "A2",
        "skill": "listening",
        "title": "Обсуждение планов на выходные с коллегой",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Bàn kế hoạch đi dã ngoại cuối tuần tại công viên",
        "dialogue": [
            {"speaker": "Ольга", "ru": "Сергей, синоптики обещают тёплую и солнечную погоду в эту субботу. Какие планы?", "vi": "Sergey, dự báo thời tiết hứa hẹn thứ Bảy này trời ấm và có nắng đấy. Cậu có dự định gì chưa?"},
            {"speaker": "Сергей", "ru": "Я думал поехать в парк «Царицыно» покататься на велосипедах и устроить пикник. Поедешь со мной?", "vi": "Mình tính tới công viên Tsaritsyno đạp xe và làm chuyến dã ngoại. Cậu đi cùng mình không?"},
            {"speaker": "Ольга", "ru": "С удовольствием! Я испеку яблочный пирог и возьму термос с ароматным чаем.", "vi": "Rất sẵn lòng! Mình sẽ nướng bánh táo và mang theo bình giữ nhiệt trà thơm."},
            {"speaker": "Сергей", "ru": "Договорились. Встретимся в субботу в одиннадцать утра у главного входа.", "vi": "Thống nhất nhé. Hẹn gặp nhau lúc 11 giờ sáng thứ Bảy tại cổng chính."}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "Встретимся в субботу в одиннадцать утра у главного входа.",
            "hint": "Thời gian và địa điểm hẹn gặp mặt cuối tuần."
        },
        "questions": [
            {
                "id": "a2_l15_q1",
                "prompt": "Что Ольга пообещала взять с собой на пикник?",
                "hint": "Xem món đồ ăn thức uống mà bạn gái chuẩn bị.",
                "options": [
                    {"key": "A", "text": "Яблочный пирог и термос с чаем"},
                    {"key": "B", "text": "Бутерброды с колбасой"},
                    {"key": "C", "text": "Шашлыки и мангал"},
                    {"key": "D", "text": "Фруктовый сок"}
                ],
                "ans": "A",
                "explain": "Ольга вызвалась испечь пирог и взять термос с горячим чаем."
            }
        ]
    },
    {
        "id": 16,
        "level": "A2",
        "skill": "listening",
        "title": "Покупка сувениров в сувенирном магазине",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Mua búp bê Matryoshka và hộp sứ lưu niệm",
        "dialogue": [
            {"speaker": "Продавец", "ru": "Здравствуйте! Ищете традиционные русские сувениры? Обратите внимание на наши матрёшки ручной росписи.", "vi": "Xin chào! Bạn đang tìm quà lưu niệm truyền thống của Nga ạ? Xin mời xem những búp bê Matryoshka vẽ tay này nhé."},
            {"speaker": "Покупатель", "ru": "Добрый день! Очень красивые. А сколько мест внутри этой большой матрёшки?", "vi": "Chào bác! Đẹp quá ạ. Trong búp bê to này lồng được bao nhiêu con nhỏ vậy bác?"},
            {"speaker": "Продавец", "ru": "Здесь семь кукол, каждая расписана сюжетами русских народных сказок.", "vi": "Mẫu này có 7 con, mỗi con vẽ một tích trong truyện cổ tích dân gian Nga."},
            {"speaker": "Покупатель", "ru": "Прекрасно! Заверните её, пожалуйста, в подарочную коробку.", "vi": "Tuyệt quá! Nhờ bác gói giúp cháu vào hộp quà nhé ạ."}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "Заверните сувенир в красивую подарочную коробку.",
            "hint": "Nghe yêu cầu đóng gói bọc quà lưu niệm."
        },
        "questions": [
            {
                "id": "a2_l16_q1",
                "prompt": "Сколько кукол спрятано внутри выбранной матрёшки?",
                "hint": "Số lượng con búp bê lồng bên trong.",
                "options": [
                    {"key": "A", "text": "3 куклы"},
                    {"key": "B", "text": "7 кукол"},
                    {"key": "C", "text": "10 кукол"},
                    {"key": "D", "text": "15 кукол"}
                ],
                "ans": "B",
                "explain": "Продавец пояснил, что в этой модели ровно семь кукол."
            }
        ]
    },
    {
        "id": 17,
        "level": "A2",
        "skill": "listening",
        "title": "В книжном магазине: Поиск учебника русского языка",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Tìm mua sách ngữ pháp và từ điển trong nhà sách",
        "dialogue": [
            {"speaker": "Покупатель", "ru": "Здравствуйте! Подскажите, где у вас отдел учебников русского языка как иностранного?", "vi": "Xin chào! Làm ơn cho tôi hỏi khu vực giáo trình tiếng Nga cho người nước ngoài ở đâu ạ?"},
            {"speaker": "Консультант", "ru": "Добрый день! Это на втором этаже, секция номер двенадцать, рядом со словарями.", "vi": "Chào bạn! Ở trên tầng hai nhé, quầy số 12, ngay bên cạnh khu từ điển."},
            {"speaker": "Покупатель", "ru": "А есть ли у вас сборник упражнений по грамматике для уровня A2?", "vi": "Bên mình có cuốn bài tập thực hành ngữ pháp trình độ A2 không ạ?"},
            {"speaker": "Консультант", "ru": "Да, только вчера поступил отличный практический курс с аудиодиском и ключами.", "vi": "Có chứ, vừa hôm qua mới về bộ bài tập kèm đĩa nghe và đáp án rất hay."}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "Отдел учебников находится на втором этаже в двенадцатой секции.",
            "hint": "Chỉ dẫn vị trí quầy sách trên tầng 2."
        },
        "questions": [
            {
                "id": "a2_l17_q1",
                "prompt": "Где находится отдел литературы по русскому языку как иностранному?",
                "hint": "Xem số tầng và số quầy ở câu chỉ dẫn.",
                "options": [
                    {"key": "A", "text": "На первом этаже у кассы"},
                    {"key": "B", "text": "На втором этаже, секция 12"},
                    {"key": "C", "text": "В подвальном помещении"},
                    {"key": "D", "text": "Такого отдела нет"}
                ],
                "ans": "B",
                "explain": "Консультант направил покупателя на второй этаж в секцию 12."
            }
        ]
    },
    {
        "id": 18,
        "level": "A2",
        "skill": "listening",
        "title": "Разговор в автосервисе: Замена колеса",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Đưa xe vào tiệm sửa chữa và thay bánh xe",
        "dialogue": [
            {"speaker": "Водитель", "ru": "Здравствуйте! На трассе я пробил переднее правое колесо. Можете посмотреть?", "vi": "Xin chào! Trên đường cao tốc tôi bị thủng lốp trước bên phải. Các anh kiểm tra giúp được không?"},
            {"speaker": "Механик", "ru": "Давайте загоним машину на подъёмник. Сейчас снимем колесо и проверим давление.", "vi": "Bác đánh xe lên cầu nâng nhé. Giờ tụi em tháo bánh ra và đo áp suất lốp ngay."},
            {"speaker": "Водитель", "ru": "Сколько времени займёт ремонт и балансировка?", "vi": "Sửa và cân mâm hết bao nhiêu thời gian anh?"},
            {"speaker": "Механик", "ru": "Примерно тридцать минут. Можете пока выпить кофе в комнате для клиентов.", "vi": "Tầm 30 phút thôi ạ. Bác vào phòng chờ uống cà phê nghỉ ngơi nhé."}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "Ремонт и балансировка колеса займут примерно тридцать минут.",
            "hint": "Thời lượng hoàn thành việc sửa chữa lốp xe."
        },
        "questions": [
            {
                "id": "a2_l18_q1",
                "prompt": "Какое колесо повредил водитель автомобиля?",
                "hint": "Xem vị trí bánh xe bị hỏng ở đầu hội thoại.",
                "options": [
                    {"key": "A", "text": "Переднее правое колесо"},
                    {"key": "B", "text": "Заднее левое колесо"},
                    {"key": "C", "text": "Запасное колесо в багажнике"},
                    {"key": "D", "text": "Все четыре колеса"}
                ],
                "ans": "A",
                "explain": "Водитель сообщил: пробил переднее правое колесо."
            }
        ]
    },
    {
        "id": 19,
        "level": "A2",
        "skill": "listening",
        "title": "Бронирование столика на день рождения",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Đặt bàn tiệc sinh nhật tại nhà hàng",
        "dialogue": [
            {"speaker": "Хостес", "ru": "Ресторан «Суздаль», добрый вечер! Слушаю вас.", "vi": "Nhà hàng 'Suzdal', xin chào quý khách! Tôi có thể hỗ trợ gì ạ?"},
            {"speaker": "Гость", "ru": "Добрый вечер! Я хочу забронировать столик на субботу на шесть человек на 19:00.", "vi": "Chào chị! Tôi muốn đặt một bàn cho 6 người vào tối thứ Bảy lúc 19:00."},
            {"speaker": "Хостес", "ru": "Прекрасно. У вас праздник? Мы дарим именинникам десерт и скидку 10%.", "vi": "Tuyệt quá. Nhà mình có dịp lễ gì không ạ? Nhà hàng xin tặng món tráng miệng và giảm 10% cho tiệc sinh nhật ạ."},
            {"speaker": "Гость", "ru": "Да, у моей подруги день рождения! Забронируйте, пожалуйста, столик у окна.", "vi": "Vâng, hôm đó là sinh nhật bạn tôi! Nhờ chị xếp giúp một bàn cạnh cửa sổ nhé."}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "Мы дарим именинникам праздничный десерт и скидку десять процентов.",
            "hint": "Ưu đãi đặc biệt của nhà hàng dành cho tiệc sinh nhật."
        },
        "questions": [
            {
                "id": "a2_l19_q1",
                "prompt": "Какой подарок ресторан делает гостям в честь дня рождения?",
                "hint": "Xem ưu đãi của quản lý nhà hàng.",
                "options": [
                    {"key": "A", "text": "Бесплатный торт и скидку 10%"},
                    {"key": "B", "text": "Билеты в кино"},
                    {"key": "C", "text": "Бесплатный обед на месяц"},
                    {"key": "D", "text": "Никаких подарков"}
                ],
                "ans": "A",
                "explain": "Ресторан дарит десерт и 10% скидку именинникам."
            }
        ]
    },
    {
        "id": 20,
        "level": "A2",
        "skill": "listening",
        "title": "Обсуждение фильма после похода в кино",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Trao đổi cảm nhận về tác phẩm điện ảnh vừa xem",
        "dialogue": [
            {"speaker": "Марина", "ru": "Ну как тебе фильм? Мне очень понравилась игра главных актёров!", "vi": "Cậu thấy phim thế nào? Mình rất ưng diễn xuất của các diễn viên chính!"},
            {"speaker": "Виктор", "ru": "Сюжет действительно захватывающий, особенно неожиданная развязка в финале.", "vi": "Cốt truyện thực sự cuốn hút, nhất là cái kết bất ngờ không thể đoán trước ở đoạn cuối."},
            {"speaker": "Марина", "ru": "Да, и музыкальное сопровождение создаёт невероятную атмосферу.", "vi": "Chuẩn luôn, cả phần âm nhạc nền cũng tạo ra bầu không khí tuyệt vời."},
            {"speaker": "Виктор", "ru": "Обязательно посоветую этот фильм своим университетским друзьям.", "vi": "Nhất định mình sẽ giới thiệu bộ phim này cho nhóm bạn ở trường đại học."}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "Сюжет фильма действительно захватывающий и очень динамичный.",
            "hint": "Đánh giá về cốt truyện phim điện ảnh."
        },
        "questions": [
            {
                "id": "a2_l20_q1",
                "prompt": "Что особенно впечатлило Виктора в фильме?",
                "hint": "Xem nhận xét của Viktor về kịch bản.",
                "options": [
                    {"key": "A", "text": "Захватывающий сюжет и неожиданная развязка"},
                    {"key": "B", "text": "Короткая продолжительность"},
                    {"key": "C", "text": "Удобные кресла в кинотеатре"},
                    {"key": "D", "text": "Вкусный попкорн"}
                ],
                "ans": "A",
                "explain": "Виктор отметил захватывающий сюжет и неожиданный финал картины."
            }
        ]
    },
    {
        "id": 21,
        "level": "A2",
        "skill": "listening",
        "title": "В банке: Обмен валюты и открытие карты",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Đổi ngoại tệ và mở thẻ ngân hàng tại quầy giao dịch",
        "dialogue": [
            {"speaker": "Сотрудник", "ru": "Здравствуйте! Возьмите талон электронной очереди. Какая операция вас интересует?", "vi": "Xin chào! Mời bạn lấy phiếu số thứ tự điện tử. Bạn cần thực hiện giao dịch gì ạ?"},
            {"speaker": "Клиент", "ru": "Добрый день! Мне нужно обменять 300 долларов на рубли и открыть дебетовую карту «Мир».", "vi": "Chào chị! Tôi cần đổi 300 đô la sang rúp và mở một thẻ ghi nợ nội địa Mir."},
            {"speaker": "Сотрудник", "ru": "Курс обмена сегодня девяносто два рубля за доллар. Предъявите, пожалуйста, ваш паспорт и миграционную карту.", "vi": "Tỷ giá quy đổi hôm nay là 92 rúp một USD. Cho tôi xin hộ chiếu và thẻ di trú của bạn nhé."},
            {"speaker": "Клиент", "ru": "Вот все документы. Карта будет готова сегодня?", "vi": "Dạ đủ giấy tờ đây ạ. Thẻ có lấy được ngay trong hôm nay không chị?"},
            {"speaker": "Сотрудник", "ru": "Моментальную неименную карту выдадим прямо сейчас за 10 минут.", "vi": "Thẻ phát hành nhanh không in tên thì cấp cho bạn ngay bây giờ trong 10 phút ạ."}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "Моментальную банковскую карту выдадут прямо сейчас за десять минут.",
            "hint": "Thời gian phát hành thẻ ngân hàng lấy ngay."
        },
        "questions": [
            {
                "id": "a2_l21_q1",
                "prompt": "Сколько времени займёт получение моментальной карты?",
                "hint": "Xem thời gian nhân viên ngân hàng thông báo.",
                "options": [
                    {"key": "A", "text": "Около 10 минут"},
                    {"key": "B", "text": "Две недели"},
                    {"key": "C", "text": "Три рабочих дня"},
                    {"key": "D", "text": "Один месяц"}
                ],
                "ans": "A",
                "explain": "Сотрудник банка сообщил, что моментальную карту выдадут за 10 минут."
            }
        ]
    },
    {
        "id": 22,
        "level": "A2",
        "skill": "listening",
        "title": "Заказ доставки еды на дом",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Gọi điện thoại xác nhận đơn đặt đồ ăn giao tận nhà",
        "dialogue": [
            {"speaker": "Оператор", "ru": "Служба доставки «БыстроЕд», здравствуйте! Подтверждаем ваш заказ номер 412.", "vi": "Dịch vụ giao đồ ăn 'BystroEd' xin chào! Chúng tôi xin xác nhận đơn hàng số 412 của bạn."},
            {"speaker": "Заказчик", "ru": "Добрый вечер! Да, я заказывал две порции пельменей, борщ со сметаной и морс.", "vi": "Chào bạn! Vâng, mình đặt 2 suất sủi cảo pelmeni, súp củ dền kèm kem chua và nước ép dâu dại mors."},
            {"speaker": "Оператор", "ru": "Всё верно. Адрес доставки: улица Ленина, дом 15, квартира 48, верно?", "vi": "Đúng rồi bạn. Địa chỉ nhận: phố Lênin, nhà số 15, căn hộ 48, đúng không ạ?"},
            {"speaker": "Заказчик", "ru": "Да, код домофона 48К. Курьер успеет приехать в течение получаса?", "vi": "Chuẩn rồi, mã chuông cửa là 48K. Shipper có kịp giao tới trong vòng nửa tiếng không?"},
            {"speaker": "Оператор", "ru": "Курьер уже выехал, доставка займёт около 25 минут.", "vi": "Shipper đã xuất phát rồi ạ, giao hàng mất tầm 25 phút nữa thôi."}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "Курьер уже выехал, доставка займёт около двадцати пяти минут.",
            "hint": "Thời gian shipper giao đồ ăn đến căn hộ."
        },
        "questions": [
            {
                "id": "a2_l22_q1",
                "prompt": "Какие блюда заказал клиент?",
                "hint": "Xem danh sách món ăn trong lời đáp của khách.",
                "options": [
                    {"key": "A", "text": "Пельмени, борщ со сметаной и морс"},
                    {"key": "B", "text": "Пиццу и суши"},
                    {"key": "C", "text": "Только десерт"},
                    {"key": "D", "text": "Гамбургер и картошку фри"}
                ],
                "ans": "A",
                "explain": "В заказе были традиционные русские блюда: пельмени, борщ и ягодный морс."
            }
        ]
    },
    {
        "id": 23,
        "level": "A2",
        "skill": "listening",
        "title": "Потеря вещей: Разговор в бюро находок",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Tìm đồ thất lạc trên tàu điện ngầm tại quầy thông tin",
        "dialogue": [
            {"speaker": "Пассажир", "ru": "Здравствуйте! Кажется, я забыл свой синий рюкзак в вагоне метро на станции «Курская».", "vi": "Xin chào! Hình như tôi để quên chiếc ba lô màu xanh trong toa tàu ở ga Kurskaya."},
            {"speaker": "Дежурный", "ru": "Добрый день. Опишите подробно, что находилось внутри рюкзака?", "vi": "Chào bạn. Xin bạn mô tả chi tiết xem bên trong ba lô có những đồ đạc gì ạ?"},
            {"speaker": "Пассажир", "ru": "Там был ноутбук, синяя тетрадь по русскому языку и связка ключей с брелоком в виде матрёшки.", "vi": "Trong đó có một chiếc laptop, quyển vở học tiếng Nga màu xanh và chùm chìa khóa móc búp bê matryoshka."},
            {"speaker": "Дежурный", "ru": "Минуту, проверяю по базе данных... Да, дежурный по станции передал похожий рюкзак. Предъявите паспорт для возврата.", "vi": "Đợi tôi một phút để tra cơ sở dữ liệu... Có rồi, nhân viên trực ga đã bàn giao một chiếc ba lô y như vậy. Xin mời xuất trình hộ chiếu để nhận lại đồ."}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "Дежурный по станции передал найденный рюкзак в бюро находок.",
            "hint": "Thủ tục bàn giao đồ thất lạc tìm thấy trong ga tàu."
        },
        "questions": [
            {
                "id": "a2_l23_q1",
                "prompt": "Какой отличительный брелок был на связке ключей?",
                "hint": "Xem món đồ lưu niệm gắn trên chùm chìa khóa.",
                "options": [
                    {"key": "A", "text": "Брелок в виде матрёшки"},
                    {"key": "B", "text": "Футбольный мяч"},
                    {"key": "C", "text": "Автомобильный значок"},
                    {"key": "D", "text": "Маленький фонарик"}
                ],
                "ans": "A",
                "explain": "Пассажир описал связку ключей с брелоком в виде русской матрёшки."
            }
        ]
    },
    {
        "id": 24,
        "level": "A2",
        "skill": "listening",
        "title": "Подготовка к экзамену по русскому языку",
        "listeningTime": "12 phút",
        "topic": "Hội thoại: Đôi bạn cùng ôn luyện ngữ pháp trước kỳ thi TRKI",
        "dialogue": [
            {"speaker": "Лена", "ru": "Привет, Нам! Завтра у нас тест по глаголам движения с приставками. Ты всё повторил?", "vi": "Chào Nam! Ngày mai tụi mình thi kiểm tra động từ chuyển động có tiền tố đấy. Cậu ôn hết chưa?"},
            {"speaker": "Нам", "ru": "Привет, Лена! Я хорошо выучил приставки «по-», «в-» и «вы-», но иногда путаю «дойти» и «перейти».", "vi": "Chào Lena! Mình đã nắm chắc các tiền tố po-, v- và vy-, nhưng đôi khi vẫn nhầm lẫn giữa dojti và perejti."},
            {"speaker": "Лена", "ru": "Смотри: «дойти» означает достичь цели, а «перейти» — пересечь улицу или реку. Давай сделаем пару упражнений вместе!", "vi": "Nhìn này: dojti là đi đến đích điểm, còn perejti là băng qua một con đường hay qua sông. Nào, tụi mình cùng làm vài bài tập nhé!"},
            {"speaker": "Нам", "ru": "Отличная идея, так учиться намного понятнее и веселее!", "vi": "Ý tưởng tuyệt quá, cùng học thế này dễ hiểu và vui hơn hẳn!"}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "Глагол перейти означает пересечь улицу или дорогу с одной стороны на другую.",
            "hint": "Quy tắc ngữ pháp về ý nghĩa của động từ có tiền tố pere-."
        },
        "questions": [
            {
                "id": "a2_l24_q1",
                "prompt": "В чём разница между «дойти» и «перейти» по объяснению Лены?",
                "hint": "Xem giải thích của Lena về mục tiêu và phương hướng di chuyển.",
                "options": [
                    {"key": "A", "text": "«Дойти» — достичь цели, «перейти» — пересечь улицу"},
                    {"key": "B", "text": "Они означают одно и то же"},
                    {"key": "C", "text": "«Перейти» значит лететь на самолёте"},
                    {"key": "D", "text": "«Дойти» используется только зимой"}
                ],
                "ans": "A",
                "explain": "Лена объяснила: «дойти» — дойти до цели, «перейти» — пересечь преграду или дорогу."
            }
        ]
    },
    {
        "id": 25,
        "level": "A2",
        "skill": "listening",
        "title": "Итоговый диалог: Планы на каникулы и уровень B1",
        "listeningTime": "12 phút",
        "topic": "Hội thoại tổng kết: Kỳ nghỉ hè và mục tiêu chinh phục B1",
        "dialogue": [
            {"speaker": "Преподаватель", "ru": "Дорогие студенты! Поздравляю вас с успешным окончанием курса аудирования уровня A2!", "vi": "Các em sinh viên thân mến! Thầy/cô xin chúc mừng các em đã hoàn thành xuất sắc khóa luyện nghe trình độ A2!"},
            {"speaker": "Студент", "ru": "Спасибо вам огромное за интересные уроки и поддержку! Мы стали гораздо лучше понимать русскую речь на слух.", "vi": "Chúng em cảm ơn thầy/cô rất nhiều vì những bài giảng bổ ích và sự đồng hành tận tình! Chúng em nghe hiểu tiếng Nga tự tin hơn nhiều rồi ạ."},
            {"speaker": "Преподаватель", "ru": "Какие у вас планы на лето? Будете продолжать занятия?", "vi": "Mùa hè này các em có dự định gì? Có tiếp tục học nâng cao không?"},
            {"speaker": "Студент", "ru": "Да, летом я поеду на практику в Санкт-Петербург, а с осени начну подготовку к сертификату первого уровня B1!", "vi": "Dạ có ạ, mùa hè em sẽ đi thực tập tại Saint Petersburg, còn từ mùa thu em sẽ bắt đầu ôn thi chứng chỉ B1 bậc 1 ạ!"}
        ],
        "dictation": {
            "title": "Chính tả nghe hiểu (Диктант)",
            "instruction": "Nghe và điền chính xác câu sau:",
            "targetText": "С осени мы начинаем интенсивную подготовку к уровню B1.",
            "hint": "Mục tiêu học thuật mới sau khi hoàn tất chặng A2."
        },
        "questions": [
            {
                "id": "a2_l25_q1",
                "prompt": "Куда студент планирует поехать летом на практику?",
                "hint": "Xem địa danh được sinh viên nhắc tới ở câu kết.",
                "options": [
                    {"key": "A", "text": "В Санкт-Петербург"},
                    {"key": "B", "text": "В Сочи"},
                    {"key": "C", "text": "Домой без практики"},
                    {"key": "D", "text": "В Лондон"}
                ],
                "ans": "A",
                "explain": "Студент сообщил преподавателю о планах пройти летнюю практику в Санкт-Петербурге."
            }
        ]
    }
]
