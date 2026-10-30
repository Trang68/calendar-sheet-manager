# -*- coding: utf-8 -*-
"""
Generator for A2 Writing Lessons 4 to 25.
"""

A2_WRITING_4_TO_25 = [
    {
        "id": 4,
        "level": "A2",
        "skill": "writing",
        "title": "Письмо другу о переезде в новую квартиру",
        "writingTime": "25 phút",
        "topic": "Viết thư kể cho bạn thân về việc dọn sang căn hộ mới",
        "targetWordRange": [50, 80],
        "task": "Hãy viết một bức thư ngắn (50–80 từ) gửi cho bạn thân kể về căn phòng mới mà bạn vừa chuyển đến thuê.",
        "requirements": [
            "Lời chào thân mật mở đầu thư",
            "Nêu lý do chuyển nhà và vị trí của căn hộ mới (gần ga tàu điện ngầm, công viên)",
            "Miêu tả không gian phòng và tiện nghi (bàn học, ban công, máy giặt)",
            "Lời mời bạn tới chơi và lời chúc sức khỏe kết thư"
        ],
        "cliches": [
            {"ru": "Привет, Антон!", "vi": "Chào Anton!"},
            {"ru": "Хочу поделиться отличной новостью: я переехал!", "vi": "Mình muốn chia sẻ một tin vui: mình vừa chuyển nhà rồi!"},
            {"ru": "Квартира очень светлая и уютная.", "vi": "Căn hộ rất sáng sủa và ấm cúng."},
            {"ru": "Приходи ко мне в гости на выходных!", "vi": "Cuối tuần này sang nhà mình chơi nhé!"}
        ],
        "modelAnswer": "Привет, Антон!\n\nХочу поделиться радостной новостью: я наконец переехал в новую квартиру! Она находится всего в пяти минутах от станции метро «Университет».\n\nКомната очень светлая и просторная. Здесь есть большой письменный стол, удобная кровать và ban công nhìn ra công viên cây xanh.\n\nПриглашаю тебя в гости в эту субботу на чай с пирогами!\n\nТвой друг,\nНам",
        "modelWordCount": 58,
        "rubric": [
            {"criteria": "Nội dung & Hoàn thành yêu cầu đề bài", "weight": "40%", "desc": "Nêu đầy đủ thông tin vị trí, tiện nghi và lời mời chân thành."},
            {"criteria": "Từ vựng & Cấu trúc ngữ pháp A2", "weight": "35%", "desc": "Sử dụng đúng biến cách giới từ, tính từ miêu tả phòng ốc."},
            {"criteria": "Bố cục & Phong cách thư từ", "weight": "25%", "desc": "Lời chào, ngắt đoạn logic và ký tên chuẩn văn phong thân mật."}
        ]
    },
    {
        "id": 5,
        "level": "A2",
        "skill": "writing",
        "title": "Записка соседу с просьбой о помощи",
        "writingTime": "20 phút",
        "topic": "Mẩu ghi chú ngắn nhờ hàng xóm nhận hộ bưu phẩm",
        "targetWordRange": [40, 65],
        "task": "Bạn vắng nhà vào chiều nay và shipper sẽ giao bưu kiện. Hãy để lại một mẩu giấy ghi chú gửi người hàng xóm thân thiết nhờ nhận giùm.",
        "requirements": [
            "Lời chào và xưng hô lịch thiệp",
            "Nêu rõ lý do bận vắng nhà và nhờ nhận gói hàng từ người giao hàng",
            "Cam kết thời gian sang lấy lại và gửi lời cảm ơn"
        ],
        "cliches": [
            {"ru": "Здравствуйте, уважаемый Алексей!", "vi": "Xin chào bác Aleksey kính mến!"},
            {"ru": "У меня к вам небольшая просьба.", "vi": "Cháu có một lời nhờ vả nhỏ này gửi tới bác ạ."},
            {"ru": "Буду очень благодарен за помощь!", "vi": "Cháu sẽ vô cùng biết ơn sự giúp đỡ của bác!"}
        ],
        "modelAnswer": "Здравствуйте, уважаемый Алексей!\n\nИзвините за беспокойство. Сегодня с 14:00 до 16:00 курьер должен привезти мне посылку с книгами. К сожалению, я задерживаюсь на лекциях в университете.\n\nНе могли бы вы принять посылку? Я зайду к вам вечером в 19:00 и всё заберу.\n\nОгромное спасибо за помощь!\n\nС уважением,\nТхань (кв. 42)",
        "modelWordCount": 52,
        "rubric": [
            {"criteria": "Mục đích giao tiếp rõ ràng", "weight": "40%", "desc": "Trình bày cụ thể việc cần nhờ và giờ giấc hẹn gặp."},
            {"criteria": "Tính lịch sự & Ngữ pháp chuẩn xác", "weight": "35%", "desc": "Sử dụng lối nói nhã nhặn 'не могли бы вы', 'извините'."},
            {"criteria": "Định dạng mẩu ghi chú", "weight": "25%", "desc": "Ghi rõ số phòng căn hộ và ký tên người gửi."}
        ]
    },
    {
        "id": 6,
        "level": "A2",
        "skill": "writing",
        "title": "Короткий отзыв о ресторане или кафе",
        "writingTime": "25 phút",
        "topic": "Viết bài đánh giá ngắn gọn về nhà hàng, quán ăn trên mạng",
        "targetWordRange": [50, 75],
        "task": "Hãy viết một bài nhận xét ngắn (50–75 từ) trên ứng dụng bản đồ về một quán ăn Nga mà bạn vừa trải nghiệm.",
        "requirements": [
            "Tên quán ăn và thời điểm bạn ghé thăm",
            "Đánh giá về món ăn (hương vị, độ tươi ngon)",
            "Đánh giá về thái độ phục vụ của nhân viên và không gian quán",
            "Đưa ra lời khuyên hoặc thang điểm đánh giá"
        ],
        "cliches": [
            {"ru": "Вчера мы ужинали в кафе...", "vi": "Hôm qua chúng tôi dùng bữa tối tại quán cà phê..."},
            {"ru": "Еда была очень вкусной и свежей.", "vi": "Đồ ăn rất ngon miệng và tươi mới."},
            {"ru": "Официанты вежливые и внимательные.", "vi": "Nhân viên phục vụ nhã nhặn và chu đáo."},
            {"ru": "Рекомендую всем посетить это место!", "vi": "Tôi khuyên mọi người nên ghé thăm nơi này!"}
        ],
        "modelAnswer": "Вчера мы ужинали в ресторане русской кухни «Теремок». Заказывали традиционный борщ со сметаной и горячие блины с красной рыбой.\n\nВсе блюда были невероятно вкусными, свежими и ароматными! Персонал очень внимательный и доброжелательный, заказ принесли быстро. В зале играет приятная музыка и царит уютная атмосфера.\n\nСтавлю пять звёзд из пяти и обязательно вернусь сюда снова!",
        "modelWordCount": 56,
        "rubric": [
            {"criteria": "Nội dung nhận xét", "weight": "40%", "desc": "Đầy đủ các khía cạnh: món ăn, phục vụ, không gian, mức điểm."},
            {"criteria": "Vốn từ vựng ẩm thực", "weight": "35%", "desc": "Các tính từ đánh giá chính xác: свежий, вкусный, доброжелательный."},
            {"criteria": "Tính mạch lạc", "weight": "25%", "desc": "Câu văn trôi chảy, sử dụng tốt các liên từ."}
        ]
    },
    {
        "id": 7,
        "level": "A2",
        "skill": "writing",
        "title": "Письмо родителям о жизни и учёбе в России",
        "writingTime": "25 phút",
        "topic": "Viết thư gửi bố mẹ cập nhật tình hình ăn ở và học tập nơi xứ người",
        "targetWordRange": [60, 90],
        "task": "Hãy viết một bức thư (60–90 từ) gửi về quê nhà cho bố mẹ, kể về cuộc sống du học sinh của bạn.",
        "requirements": [
            "Lời chào yêu thương mở đầu gửi bố mẹ",
            "Báo cáo tình hình sức khỏe và thời tiết nơi bạn đang ở",
            "Kể về việc học tiếng Nga trên giảng đường và bạn bè cùng lớp",
            "Lời dặn dò bố mẹ giữ gìn sức khỏe"
        ],
        "cliches": [
            {"ru": "Дорогие мама и папа!", "vi": "Bố mẹ yêu quý của con!"},
            {"ru": "У меня всё в порядке.", "vi": "Con ở đây mọi sự đều tốt đẹp cả."},
            {"ru": "Я успешно сдал первые зачёты.", "vi": "Con đã vượt qua những bài thi điều kiện đầu tiên."},
            {"ru": "Крепко вас обнимаю!", "vi": "Con ôm bố mẹ thật chặt!"}
        ],
        "modelAnswer": "Дорогие мама и папа!\n\nПишу вам из Москвы. У меня всё хорошо со здоровьем, питаюсь регулярно и тепло одеваюсь, так как на улице уже выпал первый снег.\n\nЗанятия в университете очень интересные. Мы каждый день учим новые русские слова, читаем тексты и тренируем диалоги. Мои преподаватели очень добрые и всегда помогают.\n\nНе волнуйтесь за меня. Берегите себя и не болейте!\n\nКрепко вас обнимаю,\nВаш сын Нам",
        "modelWordCount": 65,
        "rubric": [
            {"criteria": "Tình cảm & Nội dung bài viết", "weight": "40%", "desc": "Thể hiện sự hiếu thảo, cung cấp tin tức chân thực về học tập."},
            {"criteria": "Ngữ pháp & Chia động từ", "weight": "35%", "desc": "Sử dụng đúng thì hiện tại và các phó từ chỉ tần suất."},
            {"criteria": "Thể thức thư từ gia đình", "weight": "25%", "desc": "Mở đầu và kết thúc ấm áp đúng văn phong người Việt viết tiếng Nga."}
        ]
    },
    {
        "id": 8,
        "level": "A2",
        "skill": "writing",
        "title": "Объявление о пропаже ключей или вещей",
        "writingTime": "20 phút",
        "topic": "Viết mẩu tin thông báo tìm đồ đạc thất lạc dán tại bảng tin",
        "targetWordRange": [40, 60],
        "task": "Bạn bị rơi một chùm chìa khóa tại sảnh trường đại học. Hãy viết mẩu thông báo tìm kiếm dán tại bảng tin sinh viên.",
        "requirements": [
            "Tiêu đề thông báo nổi bật",
            "Mô tả chi tiết đồ vật bị mất (màu sắc, số lượng, đặc điểm nhận dạng)",
            "Nơi và thời điểm có khả năng bị rơi",
            "Thông tin liên lạc và lời cảm ơn kèm hậu tạ"
        ],
        "cliches": [
            {"ru": "Внимание! Потеряны ключи!", "vi": "Chú ý! Bị rơi chùm chìa khóa!"},
            {"ru": "Нашедшего просьба позвонить по номеру...", "vi": "Ai nhặt được xin vui lòng gọi theo số..."},
            {"ru": "Вознаграждение гарантируется!", "vi": "Xin đảm bảo có quà hậu tạ!"}
        ],
        "modelAnswer": "ВНИМАНИЕ! ПОТЕРЯНЫ КЛЮЧИ!\n\nВчера, 15 октября, около 14:00 в холле главного корпуса была утеряна связка ключей (3 штуки) с металлическим брелоком в виде Эйфелевой башни.\n\nНашедшего очень прошу вернуть за вознаграждение! Пожалуйста, позвоните или напишите в Telegram по номеру: +7 (999) 123-45-67.\n\nЗаранее большое спасибо!",
        "modelWordCount": 46,
        "rubric": [
            {"criteria": "Đầy đủ thông tin thiết yếu", "weight": "40%", "desc": "Nêu rõ thời gian, địa điểm, đặc điểm đồ vật và số liên hệ."},
            {"criteria": "Văn phong thông báo", "weight": "35%", "desc": "Súc tích, rõ ràng, câu cú ngắn gọn chuẩn mực."},
            {"criteria": "Định dạng văn bản", "weight": "25%", "desc": "Tiêu đề in hoa, cách dòng hợp lý."}
        ]
    },
    {
        "id": 9,
        "level": "A2",
        "skill": "writing",
        "title": "Заявление преподавателю о пропуске занятий",
        "writingTime": "20 phút",
        "topic": "Viết đơn xin phép vắng học ngắn gọn gửi thầy cô giáo",
        "targetWordRange": [45, 65],
        "task": "Bạn bị sốt cảm cúm không thể đến lớp vào sáng mai. Hãy viết thư điện tử gửi giảng viên xin phép nghỉ học và xin bài tập tự ôn.",
        "requirements": [
            "Lời chào kính trọng tới giảng viên",
            "Trình bày lý do vắng mặt (bị ốm, cảm sốt) và thời gian nghỉ",
            "Xin tài liệu bài giảng và bài tập về nhà để tự học",
            "Lời cảm ơn và ký tên đầy đủ mã lớp"
        ],
        "cliches": [
            {"ru": "Уважаемая Анна Сергеевна!", "vi": "Kính gửi cô Anna Sergeevna!"},
            {"ru": "К сожалению, я заболел и не смогу прийти...", "vi": "Rất tiếc là em bị ốm nên không thể đến lớp..."},
            {"ru": "Обязательно выполню все домашние задания.", "vi": "Em nhất định sẽ hoàn thành toàn bộ bài tập về nhà."},
            {"ru": "С уважением...", "vi": "Kính thư..."}
        ],
        "modelAnswer": "Уважаемая Анна Сергеевна!\n\nК сожалению, я сильно простудился и заболел, поэтому не смогу присутствовать на практических занятиях по русскому языку завтра, 20 ноября.\n\nВрач назначил мне постельный режим на три дня. Справку из поликлиники я обязательно принесу на следующий урок. Вышлите, пожалуйста, список домашних заданий.\n\nС уважением,\nСтудент группы 102 Нам",
        "modelWordCount": 51,
        "rubric": [
            {"criteria": "Tính trang trọng & Phép lịch sự", "weight": "40%", "desc": "Xưng hô đúng tên đệm (Имя Отчество), câu văn nhã nhặn."},
            {"criteria": "Ngữ pháp lý do & Thời gian", "weight": "35%", "desc": "Sử dụng đúng liên từ chỉ nguyên nhân (поэтому, так как)."},
            {"criteria": "Cấu trúc thư tín học vụ", "weight": "25%", "desc": "Đề ngày tháng, nêu rõ lớp và cam kết nộp giấy tờ."}
        ]
    },
    {
        "id": 10,
        "level": "A2",
        "skill": "writing",
        "title": "Рассказ о прошедших выходных на природе",
        "writingTime": "25 phút",
        "topic": "Bài văn ngắn kể lại chuyến dã ngoại ngoại ô cuối tuần",
        "targetWordRange": [55, 80],
        "task": "Hãy viết một đoạn văn (55–80 từ) kể về ngày cuối tuần tuyệt vời bạn đã trải qua cùng bạn bè tại vùng ngoại ô.",
        "requirements": [
            "Giới thiệu thời gian và địa điểm chuyến đi",
            "Các hoạt động đã tham gia (đi dạo trong rừng thông, nướng thịt, chụp ảnh)",
            "Cảm xúc sau chuyến đi đối với tinh thần và sức khỏe"
        ],
        "cliches": [
            {"ru": "В прошлую субботу мы решили поехать...", "vi": "Thứ Bảy tuần trước chúng tôi quyết định đi..."},
            {"ru": "Погода была просто великолепная!", "vi": "Thời tiết hôm ấy tuyệt vời làm sao!"},
            {"ru": "Мы отлично отдохнули и набрались сил.", "vi": "Chúng tôi đã nghỉ ngơi tuyệt vời và nạp lại đầy năng lượng."}
        ],
        "modelAnswer": "В прошлые выходные мы с друзьями решили отдохнуть от городской суеты и поехали за город на природу. Погода стояла тёплая и ясная.\n\nМы гуляли по красивому сосновому лесу, дышали чистым воздухом, делали яркие фотографии и готовили шашлыки на костре. Мы много смеялись và hát những bài ca quen thuộc.\n\nЭто был незабываемый день! Мы прекрасно отдохнули и зарядились энергией на всю неделю.",
        "modelWordCount": 61,
        "rubric": [
            {"criteria": "Sử dụng thì quá khứ hoàn thành", "weight": "40%", "desc": "Sử dụng chuẩn xác thể hoàn thành và chưa hoàn thành của động từ."},
            {"criteria": "Từ vựng miêu tả thiên nhiên", "weight": "35%", "desc": "Các cụm từ: за город, на костре, чистый воздух."},
            {"criteria": "Tính biểu cảm", "weight": "25%", "desc": "Truyền tải được cảm xúc hào hứng và thư thái."}
        ]
    },
    {
        "id": 11,
        "level": "A2",
        "skill": "writing",
        "title": "Поздравление коллеге с Новым годом",
        "writingTime": "20 phút",
        "topic": "Viết thiệp chúc Tết gửi đồng nghiệp cùng nhóm làm việc",
        "targetWordRange": [45, 65],
        "task": "Hãy viết một tấm thiệp chúc mừng năm mới (45–65 từ) gửi cho một người bạn hoặc đồng nghiệp cơ quan.",
        "requirements": [
            "Lời chúc mừng năm mới ấm áp",
            "Lời chúc sức khỏe, hạnh phúc và thành công trong công việc",
            "Cảm ơn vì sự hợp tác tốt đẹp trong năm qua",
            "Ký tên trang trọng"
        ],
        "cliches": [
            {"ru": "Поздравляю тебя с наступающим Новым годом!", "vi": "Chúc mừng bạn nhân dịp năm mới sắp sang!"},
            {"ru": "Желаю крепкого здоровья, радости и успехов!", "vi": "Chúc bạn dồi dào sức khỏe, niềm vui và thành tựu!"},
            {"ru": "Пусть все твои мечты сбудутся!", "vi": "Mong mọi ước mơ của bạn đều trở thành hiện thực!"}
        ],
        "modelAnswer": "Дорогой Алексей!\n\nОт всей души поздравляю тебя с наступающим Новым годом и Рождеством! Искренне благодарю тебя за помощь и отличную совместную работу в уходящем году.\n\nЖелаю тебе и твоей семье крепкого сибирского здоровья, неиссякаемой энергии, финансового благополучия и новых профессиональных побед! Пусть в доме всегда царят уют и радость!\n\nС наилучшими пожеланиями,\nТвой коллега Тхань",
        "modelWordCount": 54,
        "rubric": [
            {"criteria": "Văn phong chúc mừng trang trọng", "weight": "40%", "desc": "Cấu trúc chúc tụng: от всей души поздравляю, желаю чего."},
            {"criteria": "Cách danh từ đi sau желаю", "weight": "35%", "desc": "Đúng cách Sinh cách (Родительный падеж): здоровья, энергии, побед."},
            {"criteria": "Thể thức thiệp mừng", "weight": "25%", "desc": "Trình bày đẹp mắt, lời chúc ý nghĩa."}
        ]
    },
    {
        "id": 12,
        "level": "A2",
        "skill": "writing",
        "title": "Письмо с вопросами о курсах русского языка",
        "writingTime": "25 phút",
        "topic": "Viết email hỏi thông tin khóa học tiếng Nga tại trung tâm ngoại ngữ",
        "targetWordRange": [55, 80],
        "task": "Bạn muốn đăng ký khóa học tiếng Nga giao tiếp mùa hè. Hãy viết email gửi trung tâm ngôn ngữ để hỏi chi tiết về lớp học.",
        "requirements": [
            "Lời chào chính thức gửi ban giám đốc trung tâm",
            "Nêu trình độ hiện tại của bạn và nguyện vọng theo học",
            "Hỏi 2–3 câu về: thời khóa biểu, học phí và tài liệu học tập",
            "Lời kết lịch thiệp mong nhận phản hồi"
        ],
        "cliches": [
            {"ru": "Здравствуйте! Меня интересует курс...", "vi": "Xin chào! Tôi đang quan tâm đến khóa học..."},
            {"ru": "Подскажите, пожалуйста, сколько стоит обучение?", "vi": "Xin cho tôi biết học phí là bao nhiêu?"},
            {"ru": "Заранее спасибо за информацию.", "vi": "Xin cảm ơn trước về thông tin cung cấp."}
        ],
        "modelAnswer": "Здравствуйте, уважаемая администрация центра «Русское слово»!\n\nМеня зовут Нам, я студент из Вьетнама. Сейчас я закончил уровень A1 и хочу продолжить изучение русского языка на летнем интенсивном курсе A2.\n\nПодскажите, пожалуйста, когда начинаются занятия и в какие дни недели проходят уроки? Какова стоимость курса и предоставляются ли учебные материалы?\n\nЗаранее благодарю за ответ.\n\nС уважением,\nНам",
        "modelWordCount": 57,
        "rubric": [
            {"criteria": "Mục tiêu hỏi đáp cụ thể", "weight": "40%", "desc": "Đặt câu hỏi trực tiếp, lịch sự về lịch học và chi phí."},
            {"criteria": "Ngữ pháp câu hỏi gián tiếp", "weight": "35%", "desc": "Sử dụng khéo léo cấu trúc 'подскажите, пожалуйста, когда...'."},
            {"criteria": "Văn phong thư điện tử", "weight": "25%", "desc": "Lời mở đầu và kết thư trang nhã chuẩn mực."}
        ]
    },
    {
        "id": 13,
        "level": "A2",
        "skill": "writing",
        "title": "Описание своей рабочей или учебной недели",
        "writingTime": "25 phút",
        "topic": "Bài viết miêu tả lịch trình một tuần làm việc và học tập bận rộn",
        "targetWordRange": [60, 85],
        "task": "Hãy viết bài văn ngắn (60–85 từ) miêu tả một tuần bận rộn nhưng thú vị của bạn tại trường học hoặc công ty.",
        "requirements": [
            "Phân chia các hoạt động từ thứ Hai đến thứ Sáu",
            "Nhắc đến các thói quen học tập, thể thao giữa tuần",
            "Cảm giác mong đợi ngày cuối tuần để nạp lại năng lượng"
        ],
        "cliches": [
            {"ru": "Моя неделя проходит очень динамично.", "vi": "Một tuần của tôi trôi qua rất năng động."},
            {"ru": "С понедельника по пятницу я...", "vi": "Từ thứ Hai đến thứ Sáu tôi..."},
            {"ru": "Вечером я нахожу время для спорта.", "vi": "Buổi tối tôi vẫn dành thời gian cho thể thao."}
        ],
        "modelAnswer": "Моя учебная неделя проходит очень насыщенно и динамично. С понедельника по пятницу я встаю в семь утра и еду на занятия в университет.\n\nКаждый день у нас по три или четыре пары лекций и семинаров. По вторникам и четвергам после пар я обязательно хожу в бассейн, чтобы снять усталость.\n\nК пятнице я чувствую приятную усталость и с нетерпением жду субботы, чтобы выспаться и встретиться с близкими друзьями.",
        "modelWordCount": 65,
        "rubric": [
            {"criteria": "Tính liền mạch của bài viết", "weight": "40%", "desc": "Trình tự thời gian tuần hoàn logic từ đầu tuần đến cuối tuần."},
            {"criteria": "Từ nối & Trạng ngữ thời gian", "weight": "35%", "desc": "Dùng chuẩn: с понедельника по пятницу, по вторникам, к пятнице."},
            {"criteria": "Vốn từ phong phú", "weight": "25%", "desc": "Từ ngữ biểu đạt trạng thái: насыщенно, усталость, с нетерпением."}
        ]
    },
    {
        "id": 14,
        "level": "A2",
        "skill": "writing",
        "title": "Отзыв о гостинице после поездки",
        "writingTime": "25 phút",
        "topic": "Nhận xét dịch vụ khách sạn trên trang đặt phòng du lịch",
        "targetWordRange": [50, 75],
        "task": "Hãy viết một bài đánh giá khách quan về khách sạn bạn đã lưu trú trong chuyến du lịch Saint Petersburg vừa qua.",
        "requirements": [
            "Tên khách sạn và vị trí địa lý",
            "Đánh giá về phòng ốc (độ sạch sẽ, tiện nghi phòng, giường ngủ)",
            "Đánh giá về bữa sáng buffet và nhân viên lễ tân",
            "Lời khuyên cho các du khách khác"
        ],
        "cliches": [
            {"ru": "Мы останавливались в отеле...", "vi": "Chúng tôi đã lưu trú tại khách sạn..."},
            {"ru": "Номер чистый и комфортный.", "vi": "Phòng ốc sạch sẽ và tiện nghi."},
            {"ru": "Завтрак типа «шведский стол» был отличным.", "vi": "Bữa sáng buffet tự chọn rất tuyệt vời."}
        ],
        "modelAnswer": "Мы останавливались в отеле «Северная пальмира» в Санкт-Петербурге на четыре дня. Отель расположен идеально — всего 10 минут пешком до Эрмитажа.\n\nНомер оказался очень чистым, тёплым и тихим. Постельное бельё белоснежное, уборка проводилась каждый день. Завтраки вкусные và đa dạng món tráng miệng. Персонал на ресепшене был очень приветливым и помог заказать такси.\n\nОбязательно рекомендую этот отель для семейного отдыха!",
        "modelWordCount": 58,
        "rubric": [
            {"criteria": "Tiêu chí đánh giá khách quan", "weight": "40%", "desc": "Nhận xét cụ thể về vị trí, phòng, bữa ăn và dịch vụ."},
            {"criteria": "Độ chính xác ngữ pháp A2", "weight": "35%", "desc": "Tính từ bổ nghĩa cho danh từ: тёплый, чистый, приветливый."},
            {"criteria": "Tính ứng dụng thực tế", "weight": "25%", "desc": "Bố cục đánh giá dịch vụ chuẩn mực trên các app du lịch."}
        ]
    },
    {
        "id": 15,
        "level": "A2",
        "skill": "writing",
        "title": "Сообщение другу с предложением сходить в театр",
        "writingTime": "20 phút",
        "topic": "Tin nhắn rủ bạn cùng mua vé đi xem kịch tại nhà hát",
        "targetWordRange": [45, 65],
        "task": "Bạn muốn rủ người bạn Nga cùng đi xem một vở kịch vào cuối tuần. Hãy soạn một tin nhắn rủ bạn đi cùng.",
        "requirements": [
            "Lời chào thân mật",
            "Gợi ý tên vở kịch hoặc nhà hát, ngày giờ biểu diễn",
            "Hỏi xem bạn có rảnh không và đề xuất đặt vé",
            "Hẹn gặp lại"
        ],
        "cliches": [
            {"ru": "Привет, Маша! Как твои дела?", "vi": "Chào Masha! Cậu khỏe không?"},
            {"ru": "Не хочешь пойти со мной в театр?", "vi": "Cậu có muốn cùng mình đi xem kịch không?"},
            {"ru": "Дай знать, если у тебя получится!", "vi": "Báo cho mình biết nếu cậu đi được nhé!"}
        ],
        "modelAnswer": "Привет, Маша! Как твои дела?\n\nВ следующее воскресенье в Малом театре показывают замечательный спектакль «Ревизор». Я знаю, что ты очень любишь классическую русскую литературу.\n\nТы свободна в воскресенье в 18:00? Если хочешь пойти вместе, я могу купить нам билеты в партер онлайн прямо сейчас. Напиши мне, пожалуйста, до вечера!\n\nЖду ответа,\nТхань",
        "modelWordCount": 54,
        "rubric": [
            {"criteria": "Mục đích mời mọc rõ ràng", "weight": "40%", "desc": "Thời gian, địa điểm, tên vở diễn và đề xuất mua vé cụ thể."},
            {"criteria": "Văn phong hội thoại tin nhắn", "weight": "35%", "desc": "Tự nhiên, thân tình, dùng tốt các câu nghi vấn rủ rê."},
            {"criteria": "Ngữ pháp câu điều kiện 'если'", "weight": "25%", "desc": "Vận dụng câu điều kiện chân thực: nếu rảnh thì mình mua vé."}
        ]
    },
    {
        "id": 16,
        "level": "A2",
        "skill": "writing",
        "title": "Короткое эссе: Почему спорт важен для каждого?",
        "writingTime": "25 phút",
        "topic": "Tiểu luận nghị luận ngắn: Lợi ích của thể thao đối với đời sống",
        "targetWordRange": [60, 85],
        "task": "Hãy viết bài văn ngắn (60–85 từ) bày tỏ quan điểm của bạn về tầm quan trọng của việc rèn luyện thể thao hàng ngày.",
        "requirements": [
            "Mở bài: Khẳng định vai trò của thể thao với sức khỏe",
            "Thân bài: Thể thao giúp tăng cường miễn dịch, giảm căng thẳng và tạo kỷ luật",
            "Kết bài: Lời khuyên mọi người nên dành ít nhất 30 phút mỗi ngày vận động"
        ],
        "cliches": [
            {"ru": "Спорт играет важную роль в жизни человека.", "vi": "Thể thao đóng vai trò quan trọng trong đời sống con người."},
            {"ru": "В здоровом теле — здоровый дух.", "vi": "Một tinh thần minh mẫn trong một cơ thể cường tráng."},
            {"ru": "Регулярные тренировки помогают...", "vi": "Việc tập luyện đều đặn giúp ích..."}
        ],
        "modelAnswer": "Сегодня спорт играет ключевую роль в жизни современного человека. Известная поговорка гласит: «В здоровом теле — здоровый дух».\n\nВо-первых, физические упражнения укрепляют иммунитет, сердце и мышцы. Во-вторых, спорт отлично снимает стресс после учёбы или напряжённой работы и заряжает отличным настроением. Даже простая утренняя пробежка или плавание делают нас бодрыми.\n\nПоэтому я считаю, что каждый человек должен находить время для физической активности каждый день.",
        "modelWordCount": 63,
        "rubric": [
            {"criteria": "Cấu trúc lập luận", "weight": "40%", "desc": "Có mở bài, thân bài với từ nối (во-первых, во-вторых) và kết bài."},
            {"criteria": "Vốn từ học thuật A2", "weight": "35%", "desc": "Cụm từ: физические упражнения, снимает стресс, играет ключевую роль."},
            {"criteria": "Tính mạch lạc", "weight": "25%", "desc": "Luận điểm rõ ràng, thuyết phục."}
        ]
    },
    {
        "id": 17,
        "level": "A2",
        "skill": "writing",
        "title": "Письмо другу с рецептом традиционного блюда",
        "writingTime": "25 phút",
        "topic": "Viết thư hướng dẫn bạn ngoại quốc nấu món Phở hoặc Nem rán",
        "targetWordRange": [60, 85],
        "task": "Người bạn Nga rất thích ẩm thực Việt Nam và muốn học nấu món nem rán (вьетнамские блинчики). Hãy viết thư hướng dẫn cách làm đơn giản.",
        "requirements": [
            "Lời chào bạn và bày tỏ sự hào hứng",
            "Liệt kê các nguyên liệu chính (bánh đa nem, thịt xay, miến, mộc nhĩ, trứng)",
            "Hướng dẫn các bước: trộn nhân, cuộn bánh và rán giòn trên chảo dầu",
            "Lời chúc ngon miệng kết thư"
        ],
        "cliches": [
            {"ru": "Рад, что тебе понравилась наша кухня!", "vi": "Rất vui vì cậu thích ẩm thực nước mình!"},
            {"ru": "Приготовить это блюдо совсем не сложно.", "vi": "Nấu món ăn này không hề phức tạp chút nào."},
            {"ru": "Приятного аппетита!", "vi": "Chúc cậu ăn ngon miệng nhé!"}
        ],
        "modelAnswer": "Привет, Саша!\n\nОчень рад, что ты хочешь приготовить наши вьетнамские блинчики — нэм! Это очень просто и вкусно.\n\nТебе понадобятся: рисовая бумага, мясной фарш, стеклянная лапша, сушёные грибы, морковь и одно яйцо. Сначала мелко нарежь грибы и лапшу, смешай их с фаршем и яйцом. Затем заверни начинку в тонкую рисовую бумагу и обжарь блинчики на сковороде в горячем масле до золотистой корочки.\n\nПодавай горячими с соусом! Приятного аппетита!\n\nТвой друг Нам",
        "modelWordCount": 68,
        "rubric": [
            {"criteria": "Trình tự hướng dẫn hành động", "weight": "40%", "desc": "Dùng chuẩn các động từ mệnh lệnh: нарежь, смешай, заверни, обжарь."},
            {"criteria": "Từ vựng ẩm thực", "weight": "35%", "desc": "Nguyên liệu và cách chế biến chính xác."},
            {"criteria": "Tính thân mật và dễ hiểu", "weight": "25%", "desc": "Câu văn trực quan, sinh động."}
        ]
    },
    {
        "id": 18,
        "level": "A2",
        "skill": "writing",
        "title": "Заказ товара в интернет-магазине с уточнениями",
        "writingTime": "20 phút",
        "topic": "Soạn thư hỏi nhân viên hỗ trợ khách hàng về thời gian giao hàng",
        "targetWordRange": [45, 65],
        "task": "Bạn vừa đặt mua một chiếc máy đọc sách online nhưng cần thay đổi địa chỉ giao hàng. Hãy viết thư gửi bộ phận chăm sóc khách hàng.",
        "requirements": [
            "Mã đơn hàng và tên sản phẩm đã đặt",
            "Lý do cần đổi địa chỉ hoặc thời gian nhận hàng",
            "Cung cấp địa chỉ mới chính xác kèm số điện thoại",
            "Lời cảm ơn và mong được hỗ trợ"
        ],
        "cliches": [
            {"ru": "Здравствуйте! Мой номер заказа...", "vi": "Xin chào! Mã số đơn hàng của tôi là..."},
            {"ru": "Прошу вас изменить адрес доставки.", "vi": "Kính đề nghị các bạn đổi địa chỉ giao hàng."},
            {"ru": "Спасибо за понимание и содействие.", "vi": "Cảm ơn sự thấu hiểu và hỗ trợ của các bạn."}
        ],
        "modelAnswer": "Здравствуйте, служба поддержки интернет-магазина!\n\nВчера я оформил заказ № 84920 на покупку электронной книги. К сожалению, в пятницу меня не будет по указанному адресу.\n\nПрошу вас изменить адрес доставки на следующий: улица Академика Королёва, дом 8, квартира 24. Мой контактный номер телефона не изменился: +7 (900) 555-33-22.\n\nПодтвердите, пожалуйста, изменение адреса ответным письмом. Большое спасибо за помощь!\n\nС уважением,\nНам",
        "modelWordCount": 54,
        "rubric": [
            {"criteria": "Chính xác thông tin nghiệp vụ", "weight": "40%", "desc": "Nêu rõ số đơn, địa chỉ cũ - mới và thông tin xác nhận."},
            {"criteria": "Phong cách giao dịch thương mại", "weight": "35%", "desc": "Lịch sự, mạch lạc, dùng chuẩn mẫu câu 'прошу вас'."},
            {"criteria": "Ngữ pháp tiếng Nga", "weight": "25%", "desc": "Giới từ và cách chỉ địa điểm chính xác."}
        ]
    },
    {
        "id": 19,
        "level": "A2",
        "skill": "writing",
        "title": "Эссе: Мой любимый праздник в году",
        "writingTime": "25 phút",
        "topic": "Bài văn ngắn: Ngày lễ em háo hức mong đợi nhất trong năm",
        "targetWordRange": [60, 85],
        "task": "Hãy viết đoạn văn ngắn (60–85 từ) kể về ngày lễ bạn yêu thích nhất trong năm và cách bạn kỷ niệm ngày đó.",
        "requirements": [
            "Tên ngày lễ và thời điểm diễn ra",
            "Lý do bạn yêu mến ngày lễ này (được quây quần, nhận quà, nghỉ ngơi)",
            "Các hoạt động và món ăn đặc sắc gắn liền với ngày lễ"
        ],
        "cliches": [
            {"ru": "Мой самый любимый праздник — это...", "vi": "Ngày lễ yêu thích nhất của tôi là..."},
            {"ru": "В этот день вся семья собирается вместе.", "vi": "Vào ngày này cả gia đình lại sum họp đông đủ."},
            {"ru": "Этот день дарит ощущение сказки.", "vi": "Ngày này mang lại cảm giác như trong truyện cổ tích."}
        ],
        "modelAnswer": "Мой самый любимый праздник в году — это Новый год. Я люблю его с детства за сказочную атмосферу и запах свежей хвои и мандаринов.\n\nВ канун праздника мы всей семьёй украшаем пушистую ёлку разноцветными гирляндами и готовим праздничные блюда. В полночь под бой часов мы поздравляем друг друга, загадываем самые сокровенные желания и дарим подарки.\n\nЭтот праздник дарит веру в чудеса, надежду на лучшее и объединяет самых дорогих людей.",
        "modelWordCount": 65,
        "rubric": [
            {"criteria": "Nội dung & Hình ảnh biểu cảm", "weight": "40%", "desc": "Khắc họa được không khí ấm áp, thiêng liêng của lễ hội."},
            {"criteria": "Ngữ pháp câu phức A2", "weight": "35%", "desc": "Sử dụng tốt các liên từ 'за что', 'когда', 'чтобы'."},
            {"criteria": "Vốn từ văn hóa", "weight": "25%", "desc": "Các từ vựng đặc trưng: атмосфера, канун праздника, чудеса."}
        ]
    },
    {
        "id": 20,
        "level": "A2",
        "skill": "writing",
        "title": "Сообщение в чат дома о правилах тишины",
        "writingTime": "20 phút",
        "topic": "Thông báo trong nhóm cư dân chung cư về việc sửa nhà và giữ trật tự",
        "targetWordRange": [45, 65],
        "task": "Bạn sắp sửa chữa lại bàn ghế trong phòng vào cuối tuần. Hãy gửi lời nhắn lên nhóm chat chung cư xin lỗi trước về tiếng ồn và cam kết giờ giấc.",
        "requirements": [
            "Lời chào các cư dân cùng tòa nhà",
            "Thông báo về việc sửa chữa trong căn hộ số mấy",
            "Cam kết chỉ làm ồn trong khung giờ quy định (10:00–18:00)",
            "Lời xin lỗi vì sự bất tiện và cảm ơn mọi người"
        ],
        "cliches": [
            {"ru": "Уважаемые соседи, добрый день!", "vi": "Kính chào các bác hàng xóm láng giềng!"},
            {"ru": "Приношу извинения за временные неудобства.", "vi": "Cháu xin cáo lỗi vì sự bất tiện tạm thời này ạ."},
            {"ru": "Заранее спасибо за понимание!", "vi": "Xin cảm ơn sự thấu hiểu của mọi người trước ạ!"}
        ],
        "modelAnswer": "Уважаемые соседи, добрый день!\n\nПишет вам жилец квартиры № 36. В эту субботу, 12 ноября, у нас в квартире будут проходить небольшие ремонтные работы по сборке мебели.\n\nМы будем шуметь строго в разрешённые часы: с 10:00 до 13:00 и с 15:00 до 18:00, соблюдая закон о тишине. Заранее приносим искренние извинения за временный шум и благодарим за понимание!\n\nС уважением,\nНам",
        "modelWordCount": 57,
        "rubric": [
            {"criteria": "Tính văn minh & Chuẩn mực giao tiếp", "weight": "40%", "desc": "Lời lẽ lịch thiệp, tôn trọng cộng đồng dân cư."},
            {"criteria": "Tính chính xác của thông tin", "weight": "35%", "desc": "Giờ giấc, số phòng và cam kết rõ ràng."},
            {"criteria": "Ngữ pháp tiếng Nga", "weight": "25%", "desc": "Cấu trúc: строго в разрешённые часы, приносим извинения."}
        ]
    },
    {
        "id": 21,
        "level": "A2",
        "skill": "writing",
        "title": "Письмо благодарности другу за тёплый приём",
        "writingTime": "20 phút",
        "topic": "Bức thư cảm ơn bạn bè đã tiếp đãi chu đáo khi đến chơi nhà",
        "targetWordRange": [50, 75],
        "task": "Bạn vừa có chuyến thăm nhà một người bạn Nga ở ngoại ô. Hãy viết thư cảm ơn bạn và gia đình bạn vì sự đón tiếp nồng hậu.",
        "requirements": [
            "Lời chào thân mật",
            "Cảm ơn vì bữa tối thịnh soạn và sự chu đáo của bố mẹ bạn",
            "Nhắc lại kỷ niệm đáng nhớ (uống trà samovar, đi dạo)",
            "Mời bạn đến thăm lại nhà mình trong dịp tới"
        ],
        "cliches": [
            {"ru": "Хочу сказать огромное спасибо за гостеприимство!", "vi": "Mình muốn nói lời cảm ơn sâu sắc vì sự hiếu khách tuyệt vời!"},
            {"ru": "Я отлично провёл время у вас дома.", "vi": "Mình đã có một khoảng thời gian tuyệt vời tại nhà cậu."},
            {"ru": "Передавай большой привет родителям!", "vi": "Cho mình gửi lời chào hỏi thăm tới hai bác nhé!"}
        ],
        "modelAnswer": "Дорогой Иван!\n\nЯ благополучно доехал до дома и хочу от всего сердца поблагодарить тебя и твоих родителей за тёплый и радушный приём!\n\nВаша загородная дача просто чудесна, а домашние пироги твоей мамы были самыми вкусными, что я пробовал в России. Наша вечерняя беседа у камина за чаем навсегда останется в моей памяти.\n\nПередавай огромный привет родителям! В следующий раз жду тебя в гости во Вьетнаме!\n\nТвой друг,\nНам",
        "modelWordCount": 64,
        "rubric": [
            {"criteria": "Tính chân thành & Tình bạn", "weight": "40%", "desc": "Lời cảm ơn ấm áp, thể hiện sự trân trọng văn hóa gia đình Nga."},
            {"criteria": "Cấu trúc biểu cảm A2", "weight": "35%", "desc": "Cụm từ: от всего сердца поблагодарить, радушный приём."},
            {"criteria": "Phong cách thư từ", "weight": "25%", "desc": "Mở đầu và kết thúc tự nhiên, đúng ngữ điệu bè bạn."}
        ]
    },
    {
        "id": 22,
        "level": "A2",
        "skill": "writing",
        "title": "Рассказ о запомнившейся поездке на экскурсию",
        "writingTime": "25 phút",
        "topic": "Viết bài chia sẻ kỷ niệm chuyến đi tour danh thắng lịch sử",
        "targetWordRange": [60, 85],
        "task": "Hãy viết đoạn văn ngắn (60–85 từ) kể lại chuyến tham quan thành phố cổ Suzdal cùng lớp học tiếng Nga.",
        "requirements": [
            "Thời điểm chuyến đi và phương tiện di chuyển",
            "Những địa danh nổi bật bạn đã chiêm ngưỡng (nhà thờ gỗ, tu viện cổ)",
            "Ấn tượng về không gian văn hóa lịch sử và ẩm thực địa phương"
        ],
        "cliches": [
            {"ru": "На прошлых выходных мы ездили на экскурсию в...", "vi": "Cuối tuần trước chúng tôi đã đi tham quan ở..."},
            {"ru": "Город поразил нас своей древней архитектурой.", "vi": "Thành phố khiến chúng tôi sững sờ bởi kiến trúc cổ kính."},
            {"ru": "Мы сделали много красивых снимков.", "vi": "Chúng tôi đã chụp được rất nhiều bức ảnh đẹp."}
        ],
        "modelAnswer": "В прошлые выходные наш студенческий курс отправился на автобусную экскурсию в старинный город Суздаль. Дорога из Москвы заняла около четырёх часов.\n\nСуздаль буквально дышит живой историей: здесь нет высотных зданий, зато повсюду стоят белокаменные храмы и деревянные дома с резными окнами. Мы посетили кремль, попробовали местную медовуху и свежие огурцы. Экскурсовод рассказал нам много древних легенд.\n\nЭта поездка помогла мне лучше узнать богатую русскую историю.",
        "modelWordCount": 65,
        "rubric": [
            {"criteria": "Miêu tả không gian & Lịch sử", "weight": "40%", "desc": "Khắc họa được nét cổ kính, nét đẹp kiến trúc của Suzdal."},
            {"criteria": "Sử dụng thì quá khứ chuẩn xác", "weight": "35%", "desc": "Động từ hoàn thành thể chỉ chuỗi sự kiện: отправился, посетили, помогла."},
            {"criteria": "Tính mạch lạc bài viết", "weight": "25%", "desc": "Trình tự chuyến đi hấp dẫn, giàu cảm xúc."}
        ]
    },
    {
        "id": 23,
        "level": "A2",
        "skill": "writing",
        "title": "Описание своего рабочего места и любимых предметов",
        "writingTime": "20 phút",
        "topic": "Miêu tả góc học tập, bàn làm việc và các vật dụng gắn bó thân quen",
        "targetWordRange": [50, 75],
        "task": "Hãy viết đoạn văn ngắn (50–75 từ) miêu tả góc bàn học tập của bạn tại nhà hoặc ký túc xá.",
        "requirements": [
            "Vị trí đặt bàn học trong phòng (gần cửa sổ, góc phòng)",
            "Các đồ vật chính trên bàn (máy tính xách tay, đèn học, từ điển, cốc trà)",
            "Ý nghĩa của không gian này đối với sự tập trung học tập"
        ],
        "cliches": [
            {"ru": "Моё рабочее место находится...", "vi": "Nơi làm việc của tôi nằm ở..."},
            {"ru": "На столе всегда царит идеальный порядок.", "vi": "Trên bàn bao giờ cũng duy trì trật tự ngăn nắp."},
            {"ru": "Здесь мне легко сосредоточиться на учёбе.", "vi": "Ở đây tôi rất dễ tập trung vào việc học."}
        ],
        "modelAnswer": "Моё рабочее место находится у широкого окна, поэтому днём здесь всегда много солнечного света. На деревянном столе стоит мой ноутбук, лампа с тёплым светом и подставка для книг.\n\nСправа всегда лежат тетради по русскому языку и большой словарь, а слева стоит любимая керамическая кружка для чая. На стене висит календарь с планами на месяц.\n\nЭто мой любимый уголок, где я могу часами спокойно учиться и читать книги.",
        "modelWordCount": 64,
        "rubric": [
            {"criteria": "Trật tự không gian rõ nét", "weight": "40%", "desc": "Dùng từ định vị chuẩn xác: у окна, справа, слева, на стене."},
            {"criteria": "Tính từ miêu tả đồ vật", "weight": "35%", "desc": "Phối hợp tính từ và danh từ nhịp nhàng."},
            {"criteria": "Tính cảm xúc", "weight": "25%", "desc": "Thể hiện được sự gắn bó và thư thái của góc học tập."}
        ]
    },
    {
        "id": 24,
        "level": "A2",
        "skill": "writing",
        "title": "Письмо другу о планах на летний отпуск",
        "writingTime": "25 phút",
        "topic": "Bức thư bàn về dự định đi nghỉ mát và du lịch mùa hè sắp tới",
        "targetWordRange": [55, 80],
        "task": "Kỳ thi học kỳ sắp kết thúc. Hãy viết một bức thư (55–80 từ) kể cho bạn thân về những dự định thú vị của bạn trong kỳ nghỉ hè.",
        "requirements": [
            "Lời chào thân mật",
            "Nêu kế hoạch sau khi thi xong (về quê, đi biển, du lịch)",
            "Đề xuất hai bạn có thể cùng gặp nhau ở một địa điểm nào đó",
            "Hỏi thăm kế hoạch nghỉ ngơi của bạn"
        ],
        "cliches": [
            {"ru": "Экзамены почти позади!", "vi": "Kỳ thi coi như gần xong xuôi rồi!"},
            {"ru": "Этим летом я планирую поехать...", "vi": "Mùa hè này mình dự định sẽ đi..."},
            {"ru": "А какие у тебя планы на лето?", "vi": "Thế còn cậu có kế hoạch gì cho mùa hè chưa?"}
        ],
        "modelAnswer": "Привет, Дима!\n\nУра, экзаменационная сессия почти закончилась! Остался всего один зачёт по истории, и начнутся долгожданные летние каникулы.\n\nВ июле я собираюсь поехать домой во Вьетнам к родителям на три недели, а в августе вернусь в Россию и мечтаю съездить на озеро Байкал. Я слышал, что там невероятная дикая природа и кристально чистая вода.\n\nА как ты планируешь провести отпуск? Может быть, поедем на Байкал вместе?\n\nЖду твоего ответа,\nНам",
        "modelWordCount": 67,
        "rubric": [
            {"criteria": "Sử dụng thì tương lai & Dự định", "weight": "40%", "desc": "Cấu trúc: планирую поехать, собираюсь, мечтаю съездить."},
            {"criteria": "Tính liên kết đoạn văn", "weight": "35%", "desc": "Chuyển tiếp tự nhiên giữa kết quả học tập và kỳ nghỉ."},
            {"criteria": "Lời mời gợi mở", "weight": "25%", "desc": "Khích lệ bạn bè cùng tham gia chuyến phiêu lưu."}
        ]
    },
    {
        "id": 25,
        "level": "A2",
        "skill": "writing",
        "title": "Итоговое эссе: Чему я научился за курс A2",
        "writingTime": "30 phút",
        "topic": "Bài viết tổng kết: Những tiến bộ vượt bậc sau khóa học tiếng Nga A2",
        "targetWordRange": [65, 95],
        "task": "Hãy viết một bài văn ngắn tổng kết (65–95 từ) chia sẻ về những kết quả bạn đã đạt được và sự tự tin sau khi hoàn thành trình độ A2.",
        "requirements": [
            "Khẳng định việc hoàn thành xuất sắc 25 bài học cấp độ A2",
            "Nêu những kỹ năng đã tiến bộ vượt trội (viết thư, giao tiếp hàng ngày, đọc báo)",
            "Bày tỏ lòng biết ơn đối với người dạy và tinh thần học hỏi",
            "Mục tiêu hướng tới cấp độ B1 trong giai đoạn tới"
        ],
        "cliches": [
            {"ru": "За время обучения я сделал большой шаг вперёд.", "vi": "Trong suốt thời gian học tập tôi đã tiến một bước dài về phía trước."},
            {"ru": "Теперь я могу уверенно выражать свои мысли.", "vi": "Giờ đây tôi đã có thể tự tin biểu đạt suy nghĩ của bản thân."},
            {"ru": "Моя главная цель — успешно сдать экзамен ТРКИ-1.", "vi": "Mục tiêu trọng tâm của tôi là thi đỗ kỳ thi TRKI-1."}
        ],
        "modelAnswer": "Завершение курса русского языка уровня A2 стало для меня важным личным достижением. За эти 25 уроков я проделал огромную работу и сделал большой шаг вперёд в освоении языка.\n\nТеперь я могу уверенно ориентироваться в городе, бронировать билеты, общаться с врачами и друзьями, а также писать грамотные письма и эссе. Русская грамматика больше не кажется мне непреодолимой преградой.\n\nЯ искренne благодарен преподавателям за терпение и вдохновение. Моя следующая цель — перейти на уровень B1 и сдать государственный сертификационный экзамен ТРКИ-1!",
        "modelWordCount": 77,
        "rubric": [
            {"criteria": "Tính tổng kết & Giá trị truyền cảm hứng", "weight": "40%", "desc": "Trình bày sâu sắc chặng đường nỗ lực và thành quả đạt được."},
            {"criteria": "Ngữ pháp phong phú & Từ vựng nâng cao", "weight": "35%", "desc": "Dùng từ chuẩn mực: личное достижение, непреодолимая преграда, освоение языка."},
            {"criteria": "Bố cục hoàn chỉnh", "weight": "25%", "desc": "Mở đầu tự hào, thân bài cụ thể, kết bài quyết tâm hướng tới B1."}
        ]
    }
]
