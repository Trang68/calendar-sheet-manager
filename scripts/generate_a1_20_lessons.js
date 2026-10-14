const fs = require('fs');
const path = require('path');

const targetDir = path.join(__dirname, '../public/data/reading/a1');
if (!fs.existsSync(targetDir)) {
  fs.mkdirSync(targetDir, { recursive: true });
}

const lessonsData = [
  // Bài 1: Chào hỏi & Làm quen
  {
    id: 1,
    title: "Здравствуйте! Давайте познакомимся",
    topic: "Chào hỏi, Tên gọi & Đại từ nhân xưng",
    readingTime: "10 phút",
    vocabCategories: [
      {
        title: "Từ ngữ chào hỏi & Xã giao (Приветствия)",
        rows: [
          {"word": "Здра́вствуйте", "gram": "lời chào lịch sự", "meaning": "Xin chào (trang trọng)", "example": "Здравствуйте, меня зовут Антон."},
          {"word": "Приве́т", "gram": "lời chào thân mật", "meaning": "Chào bạn, chào cậu", "example": "Привет, Анна!"},
          {"word": "Спаси́бо", "gram": "từ cảm ơn", "meaning": "Cảm ơn", "example": "Большое спасибо за помощь."},
          {"word": "До свида́ния", "gram": "lời chào tạm biệt", "meaning": "Tạm biệt (hẹn gặp lại)", "example": "До свидания, до завтра!"}
        ]
      },
      {
        title: "Đại từ nhân xưng (Личные местоимения)",
        rows: [
          {"word": "я", "gram": "ngôi 1 số ít", "meaning": "tôi, mình", "example": "Я студент."},
          {"word": "ты", "gram": "ngôi 2 số ít", "meaning": "bạn, cậu", "example": "Ты студент?"},
          {"word": "он / она́", "gram": "ngôi 3 số ít", "meaning": "anh ấy / cô ấy", "example": "Он врач, а она преподаватель."},
          {"word": "мы / вы / они́", "gram": "đại từ số nhiều", "meaning": "chúng tôi / các bạn / họ", "example": "Мы учимся вместе."}
        ]
      }
    ],
    readingText: [
      "Здравствуйте! Давайте познакомимся. Меня зовут Антон. Я студент.",
      "А это Анна. Она тоже студентка. Мы учимся в университете.",
      "Вот наш преподаватель. Его зовут Виктор Сергеевич. Он очень добрый и умный человек.",
      "— Привет, Антон! Как дела?\n— Привет, Анна! Спасибо, всё хорошо. А у тебя?\n— Отлично, спасибо!",
      "Приятно с вами познакомиться! До свидания!"
    ],
    questions: [
      {
        id: "a1_r1_q1", lang: "ru",
        prompt: "Как зовут студента в тексте?",
        hint: "Đọc câu thứ 2 của bài.",
        options: [{"key":"A","text":"Виктор"},{"key":"B","text":"Антон"},{"key":"C","text":"Иван"},{"key":"D","text":"Максим"}],
        ans: "B",
        explain: "Bài viết: «Меня зовут Антон. Я студент»."
      },
      {
        id: "a1_r1_q2", lang: "ru",
        prompt: "Кто такая Анна?",
        hint: "Tìm thông tin về Anna ở đoạn 2.",
        options: [{"key":"A","text":"Она врач."},{"key":"B","text":"Она преподаватель."},{"key":"C","text":"Она студентка."},{"key":"D","text":"Она школьница."}],
        ans: "C",
        explain: "Đoạn 2 nêu rõ: «А это Анна. Она тоже студентка»."
      },
      {
        id: "a1_r1_q3", lang: "ru",
        prompt: "Кем работает Виктор Сергеевич?",
        hint: "Tìm nghề nghiệp của Viktor Sergeevich.",
        options: [{"key":"A","text":"Он студент."},{"key":"B","text":"Он наш преподаватель."},{"key":"C","text":"Он инженер."},{"key":"D","text":"Он журналист."}],
        ans: "B",
        explain: "Đoạn 3 viết: «Вот наш преподаватель. Его зовут Виктор Сергеевич»."
      },
      {
        id: "a1_r1_q4", lang: "ru",
        prompt: "Что отвечает Антон на вопрос «Как дела?»",
        hint: "Xem đoạn đối thoại ngắn.",
        options: [{"key":"A","text":"Плохо."},{"key":"B","text":"Ничего."},{"key":"C","text":"Спасибо, всё хорошо."},{"key":"D","text":"Не знаю."}],
        ans: "C",
        explain: "Anton đáp: «Спасибо, всё хорошо» (Cảm ơn, mọi chuyện đều tốt)."
      },
      {
        id: "a1_r1_q5", lang: "ru",
        prompt: "Какое слово используется для вежливого прощания?",
        hint: "Từ chào tạm biệt lịch sự ở cuối bài.",
        options: [{"key":"A","text":"Привет"},{"key":"B","text":"Здравствуйте"},{"key":"C","text":"Пока"},{"key":"D","text":"До свидания"}],
        ans: "D",
        explain: "«До свидания» là lời chào tạm biệt lịch sự, trang trọng."
      },
      {
        id: "a1_r1_q6", lang: "vi",
        prompt: "Cụm từ «Давайте познакомимся» có nghĩa là gì?",
        hint: "Động từ làm quen trong tiếng Nga.",
        options: [{"key":"A","text":"Chúng ta hãy cùng ăn cơm nhé."},{"key":"B","text":"Hãy cùng làm quen nhé."},{"key":"C","text":"Hẹn gặp lại bạn ngày mai."},{"key":"D","text":"Xin chào buổi sáng."}],
        ans: "B",
        explain: "«Давайте познакомимся» nghĩa là 'Chúng ta hãy cùng làm quen nhé'."
      },
      {
        id: "a1_r1_q7", lang: "vi",
        prompt: "Từ «тоже» trong câu «Она тоже студентка» mang nghĩa gì?",
        hint: "Phó từ chỉ sự tương đồng.",
        options: [{"key":"A","text":"Cũng, cũng vậy"},{"key":"B","text":"Không phải"},{"key":"C","text":"Rất nhiều"},{"key":"D","text":"Ở đây"}],
        ans: "A",
        explain: "«Тоже» nghĩa là 'cũng' (Cô ấy cũng là sinh viên nữ)."
      },
      {
        id: "a1_r1_q8", lang: "vi",
        prompt: "Đại từ nhân xưng nào chỉ ngôi 'chúng tôi / chúng ta'?",
        hint: "Xem bảng từ vựng đại từ nhân xưng.",
        options: [{"key":"A","text":"Вы"},{"key":"B","text":"Они"},{"key":"C","text":"Мы"},{"key":"D","text":"Ты"}],
        ans: "C",
        explain: "«Мы» là đại từ ngôi thứ nhất số nhiều (chúng tôi, chúng ta)."
      },
      {
        id: "a1_r1_q9", lang: "vi",
        prompt: "Trong tiếng Nga, câu trả lời lịch sự khi được người khác giúp đỡ là gì?",
        hint: "Từ cảm ơn.",
        options: [{"key":"A","text":"Пожалуйста"},{"key":"B","text":"Спасибо"},{"key":"C","text":"Привет"},{"key":"D","text":"Пока"}],
        ans: "B",
        explain: "«Спасибо» nghĩa là cảm ơn."
      },
      {
        id: "a1_r1_q10", lang: "vi",
        prompt: "Anton và Anna học tập ở đâu?",
        hint: "Tìm từ chỉ trường đại học.",
        options: [{"key":"A","text":"Ở trường phổ thông"},{"key":"B","text":"Ở nhà"},{"key":"C","text":"Ở trường đại học (университет)"},{"key":"D","text":"Ở thư viện"}],
        ans: "C",
        explain: "Bài viết: «Мы учимся в университете» (Chúng tôi học ở trường đại học)."
      }
    ]
  },

  // Bài 2: Đồ vật & Con người: Кто это? Что это?
  {
    id: 2,
    title: "Кто это? Что это?",
    topic: "Nhận diện người & Đồ vật xung quanh",
    readingTime: "10 phút",
    vocabCategories: [
      {
        title: "Từ hỏi cơ bản (Вопросительные слова)",
        rows: [
          {"word": "Кто?", "gram": "đại từ nghi vấn", "meaning": "Ai? (dùng cho người, động vật)", "example": "Кто это? — Это студент."},
          {"word": "Что?", "gram": "đại từ nghi vấn (đọc là /shto/)", "meaning": "Cái gì? (dùng cho đồ vật, sự vật)", "example": "Что это? — Это книга."},
          {"word": "Э́то", "gram": "đại từ chỉ định", "meaning": "Đây là, cái này là", "example": "Это стол, а это стул."}
        ]
      },
      {
        title: "Danh từ đồ vật lớp học (Предметы)",
        rows: [
          {"word": "стол", "gram": "danh từ giống đực", "meaning": "cái bàn", "example": "Вот стол."},
          {"word": "кни́га", "gram": "danh từ giống cái", "meaning": "quyển sách", "example": "Это интересная книга."},
          {"word": "ру́чка", "gram": "danh từ giống cái", "meaning": "cây bút", "example": "Где ручка? Вот она."},
          {"word": "окно́", "gram": "danh từ giống trung", "meaning": "cửa sổ", "example": "Там большое окно."}
        ]
      }
    ],
    readingText: [
      "— Скажите, пожалуйста, что это?\n— Это стол. А это стул.",
      "— А что лежит на столе?\n— На столе лежит книга, тетрадь и ручка.",
      "— Кто это там у окна?\n— Это мой друг Максим. Он тоже студент.",
      "Вот класс. Здесь сидят студенты. Они читают и пишут.",
      "Это наш новый словарь, а это карта мира."
    ],
    questions: [
      {
        id: "a1_r2_q1", lang: "ru",
        prompt: "Что лежит на столе?",
        hint: "Đọc câu trả lời ở lượt thoại thứ 2.",
        options: [{"key":"A","text":"Только ручка"},{"key":"B","text":"Книга, тетрадь и ручка"},{"key":"C","text":"Словарь и карта"},{"key":"D","text":"Компьютер"}],
        ans: "B",
        explain: "Đoạn 2 nêu: «На столе лежит книга, тетрадь и ручка»."
      },
      {
        id: "a1_r2_q2", lang: "ru",
        prompt: "Кто стоит у окна?",
        hint: "Tìm thông tin người đứng cạnh cửa sổ.",
        options: [{"key":"A","text":"Преподаватель"},{"key":"B","text":"Максим"},{"key":"C","text":"Антон"},{"key":"D","text":"Анна"}],
        ans: "B",
        explain: "Bài viết: «Это мой друг Максим»."
      },
      {
        id: "a1_r2_q3", lang: "ru",
        prompt: "Слово «Что?» в русском языке произносится как:",
        hint: "Quy tắc phát âm đặc biệt của chữ ЧТ.",
        options: [{"key":"A","text":"/что/"},{"key":"B","text":"/што/"},{"key":"C","text":"/кто/"},{"key":"D","text":"/то/"}],
        ans: "B",
        explain: "Trong tiếng Nga, «что» luôn được phát âm chuẩn là /што/."
      },
      {
        id: "a1_r2_q4", lang: "ru",
        prompt: "Для одушевлённых предметов (людей) мы задаём вопрос:",
        hint: "Từ hỏi dành cho người.",
        options: [{"key":"A","text":"Что это?"},{"key":"B","text":"Кто это?"},{"key":"C","text":"Где это?"},{"key":"D","text":"Чей это?"}],
        ans: "B",
        explain: "«Кто это?» dùng để hỏi người hoặc động vật có tri giác."
      },
      {
        id: "a1_r2_q5", lang: "ru",
        prompt: "Что такое «карта мира»?",
        hint: "Từ vựng về địa lý trong lớp học.",
        options: [{"key":"A","text":"Thẻ học sinh"},{"key":"B","text":"Bản đồ thế giới"},{"key":"C","text":"Cuốn từ điển"},{"key":"D","text":"Cái bàn học"}],
        ans: "B",
        explain: "«Карта мира» nghĩa là Bản đồ thế giới."
      },
      {
        id: "a1_r2_q6", lang: "vi",
        prompt: "Từ hỏi «Что?» được dùng trong trường hợp nào?",
        hint: "Đại từ nghi vấn chỉ vật.",
        options: [{"key":"A","text":"Hỏi về người (Ai?)"},{"key":"B","text":"Hỏi về đồ vật, hiện tượng (Cái gì?)"},{"key":"C","text":"Hỏi về nơi chốn (Ở đâu?)"},{"key":"D","text":"Hỏi về thời gian (Khi nào?)"}],
        ans: "B",
        explain: "«Что?» có nghĩa là 'Cái gì?', dùng để hỏi về đồ vật, sự vật bất động vật."
      },
      {
        id: "a1_r2_q7", lang: "vi",
        prompt: "Từ «тетрадь» trong bài đọc có nghĩa là gì?",
        hint: "Đồ dùng học tập dùng để viết bài.",
        options: [{"key":"A","text":"Quyển vở ghi bài"},{"key":"B","text":"Quyển sách giáo khoa"},{"key":"C","text":"Cây bút bi"},{"key":"D","text":"Cái thước kẻ"}],
        ans: "A",
        explain: "«Тетрадь» (danh từ giống cái đuôi -ь) nghĩa là quyển vở."
      },
      {
        id: "a1_r2_q8", lang: "vi",
        prompt: "Các sinh viên trong lớp học đang làm gì?",
        hint: "Đọc câu: «Они читают и пишут».",
        options: [{"key":"A","text":"Đang ngủ"},{"key":"B","text":"Đang đọc và viết"},{"key":"C","text":"Đang ăn cơm"},{"key":"D","text":"Đang nói chuyện điện thoại"}],
        ans: "B",
        explain: "«Они читают (đọc) и пишут (viết)»."
      },
      {
        id: "a1_r2_q9", lang: "vi",
        prompt: "Từ «друг» trong câu «Это мой друг Максим» nghĩa là gì?",
        hint: "Mối quan hệ bạn bè.",
        options: [{"key":"A","text":"Thầy giáo"},{"key":"B","text":"Anh trai"},{"key":"C","text":"Bạn bè (bạn nam)"},{"key":"D","text":"Bác sĩ"}],
        ans: "C",
        explain: "«Друг» là người bạn (nam), bạn nữ là «подруга»."
      },
      {
        id: "a1_r2_q10", lang: "vi",
        prompt: "Khi muốn hỏi 'Cái bàn ở đâu?', ta dùng từ hỏi nào?",
        hint: "Từ hỏi vị trí nơi chốn.",
        options: [{"key":"A","text":"Кто?"},{"key":"B","text":"Что?"},{"key":"C","text":"Где?"},{"key":"D","text":"Куда?"}],
        ans: "C",
        explain: "«Где?» có nghĩa là 'Ở đâu?' (Где стол? - Cái bàn ở đâu?)."
      }
    ]
  },

  // Bài 3: Đại từ sở hữu & Giống danh từ (Мой, моя, моё)
  {
    id: 3,
    title: "Мой, моя, моё: Чей это предмет?",
    topic: "Giống của danh từ & Đại từ sở hữu",
    readingTime: "10 phút",
    vocabCategories: [
      {
        title: "Đại từ sở hữu 'Của tôi' (Местоимения 'Мой')",
        rows: [
          {"word": "мой", "gram": "giống đực", "meaning": "của tôi (đi với từ giống đực)", "example": "Это мой паспорт, мой друг."},
          {"word": "моя́", "gram": "giống cái", "meaning": "của tôi (đi với từ giống cái)", "example": "Это моя мама, моя комната."},
          {"word": "моё", "gram": "giống trung", "meaning": "của tôi (đi với từ giống trung)", "example": "Это моё окно, моё пальто."},
          {"word": "мои́", "gram": "số nhiều", "meaning": "của tôi (đi với danh từ số nhiều)", "example": "Это мои книги, мои друзья."}
        ]
      },
      {
        title: "Từ hỏi về sở hữu (Чей? Чья? Чьё?)",
        rows: [
          {"word": "Чей?", "gram": "giống đực", "meaning": "Của ai? (hỏi đồ giống đực)", "example": "Чей это карандаш?"},
          {"word": "Чья?", "gram": "giống cái", "meaning": "Của ai? (hỏi đồ giống cái)", "example": "Чья это сумка?"},
          {"word": "Чьё?", "gram": "giống trung", "meaning": "Của ai? (hỏi đồ giống trung)", "example": "Чьё это письмо?"}
        ]
      }
    ],
    readingText: [
      "— Чей это рюкзак?\n— Это мой рюкзак. Он синий и удобный.",
      "— А чья это ручка на столе?\n— Это моя ручка. Спасибо!",
      "— Чьё это яблоко?\n— Это твоё яблоко. Кушай на здоровье!",
      "Вот мой дом. Вот моя комната. В комнате стоит мой стол и моё кресло.",
      "А это наши фотографии. Это мои родители и моя сестра."
    ],
    questions: [
      {
        id: "a1_r3_q1", lang: "ru",
        prompt: "Какой рюкзак у студента?",
        hint: "Tìm tính từ miêu tả chiếc ba lô ở đoạn 1.",
        options: [{"key":"A","text":"Красный и большой"},{"key":"B","text":"Синий и удобный"},{"key":"C","text":"Старый"},{"key":"D","text":"Маленький"}],
        ans: "B",
        explain: "Đoạn 1 nêu: «Это мой рюкзак. Он синий и удобный» (màu xanh dương và tiện lợi)."
      },
      {
        id: "a1_r3_q2", lang: "ru",
        prompt: "Какое местоимение согласуется со словом «комната»?",
        hint: "«Комната» kết thúc bằng đuôi -а (giống cái).",
        options: [{"key":"A","text":"Мой"},{"key":"B","text":"Моё"},{"key":"C","text":"Моя"},{"key":"D","text":"Мои"}],
        ans: "C",
        explain: "«Комната» là danh từ giống cái nên đi với «моя» (моя комната)."
      },
      {
        id: "a1_r3_q3", lang: "ru",
        prompt: "Слово «яблоко» относится к какому роду?",
        hint: "Từ kết thúc bằng đuôi -о.",
        options: [{"key":"A","text":"Мужской род (giống đực)"},{"key":"B","text":"Женский род (giống cái)"},{"key":"C","text":"Средний род (giống trung)"},{"key":"D","text":"Множественное число"}],
        ans: "C",
        explain: "«Яблоко» (quả táo) kết thúc bằng -о là danh từ giống trung (чьё это яблоко? - моё яблоко)."
      },
      {
        id: "a1_r3_q4", lang: "ru",
        prompt: "Вопрос «Чья это ручка?» задаётся потому, что «ручка»:",
        hint: "Xác định giống của từ «ручка».",
        options: [{"key":"A","text":"Мужского рода"},{"key":"B","text":"Женского рода"},{"key":"C","text":"Среднего рода"},{"key":"D","text":"Во множественном числе"}],
        ans: "B",
        explain: "«Ручка» là danh từ giống cái (đuôi -а) nên dùng từ hỏi «Чья?»."
      },
      {
        id: "a1_r3_q5", lang: "ru",
        prompt: "Что стоит в комнате?",
        hint: "Đọc câu: «В комнате стоит...»",
        options: [{"key":"A","text":"Мой стол и моё кресло"},{"key":"B","text":"Телевизор и диван"},{"key":"C","text":"Шкаф и кровать"},{"key":"D","text":"Холодильник"}],
        ans: "A",
        explain: "Đoạn 4 viết: «В комнате стоит мой стол и моё кресло»."
      },
      {
        id: "a1_r3_q6", lang: "vi",
        prompt: "Trong tiếng Nga, danh từ kết thúc bằng phụ âm cứng (ví dụ: стол, дом, брат) thuộc giống gì?",
        hint: "Quy tắc cơ bản nhận biết giống danh từ.",
        options: [{"key":"A","text":"Giống đực (Мужской род)"},{"key":"B","text":"Giống cái (Женский род)"},{"key":"C","text":"Giống trung (Средний род)"},{"key":"D","text":"Không xác định được"}],
        ans: "A",
        explain: "Hầu hết danh từ tận cùng bằng phụ âm là danh từ giống đực."
      },
      {
        id: "a1_r3_q7", lang: "vi",
        prompt: "Để nói 'Bức thư của tôi' (письмо - giống trung), ta dùng từ nào?",
        hint: "Đại từ sở hữu ngôi tôi giống trung.",
        options: [{"key":"A","text":"Мой письмо"},{"key":"B","text":"Моя письмо"},{"key":"C","text":"Моё письмо"},{"key":"D","text":"Мои письмо"}],
        ans: "C",
        explain: "«Письмо» là giống trung nên đi với «моё» $\\rightarrow$ «моё письмо»."
      },
      {
        id: "a1_r3_q8", lang: "vi",
        prompt: "Từ «кресло» trong câu «моё кресло» có nghĩa là gì?",
        hint: "Đồ nội thất để ngồi có chỗ để tay.",
        options: [{"key":"A","text":"Cái ghế bành, ghế tựa"},{"key":"B","text":"Cái giường ngủ"},{"key":"C","text":"Cái bàn làm việc"},{"key":"D","text":"Tủ đựng quần áo"}],
        ans: "A",
        explain: "«Кресло» là ghế bành/ghế tựa (danh từ giống trung)."
      },
      {
        id: "a1_r3_q9", lang: "vi",
        prompt: "Đại từ sở hữu 'của bạn' (ngôi thân mật giống đực) trong tiếng Nga là gì?",
        hint: "Tương đương với 'your'.",
        options: [{"key":"A","text":"Твой"},{"key":"B","text":"Твоя"},{"key":"C","text":"Твоё"},{"key":"D","text":"Твои"}],
        ans: "A",
        explain: "«Твой» là đại từ sở hữu giống đực chỉ 'của bạn' (ví dụ: твой друг)."
      },
      {
        id: "a1_r3_q10", lang: "vi",
        prompt: "Từ hỏi «Чьи?» dùng khi nào?",
        hint: "Hỏi về quyền sở hữu số nhiều.",
        options: [{"key":"A","text":"Hỏi đồ vật giống đực số ít"},{"key":"B","text":"Hỏi đồ vật giống cái số ít"},{"key":"C","text":"Hỏi đồ vật giống trung số ít"},{"key":"D","text":"Hỏi đồ vật số nhiều (Của ai?)"}],
        ans: "D",
        explain: "«Чьи?» là từ hỏi 'của ai' dùng cho danh từ số nhiều (ví dụ: Чьи это книги?)."
      }
    ]
  },

  // Bài 4: Vị trí & Giảng đường: Где находится?
  {
    id: 4,
    title: "В аудитории: Где находится?",
    topic: "Vị trí không gian: Ở đâu, Bên phải, Bên trái, Ở giữa",
    readingTime: "10 phút",
    vocabCategories: [
      {
        title: "Phó từ chỉ vị trí (Наречия места)",
        rows: [
          {"word": "где", "gram": "từ hỏi", "meaning": "ở đâu?", "example": "Где находится аудитория?"},
          {"word": "здесь / тут", "gram": "phó từ", "meaning": "ở đây", "example": "Мы сидим здесь."},
          {"word": "там", "gram": "phó từ", "meaning": "ở đằng kia, ở đó", "example": "Там стоит доска."},
          {"word": "спра́ва / сле́ва", "gram": "phó từ", "meaning": "bên phải / bên trái", "example": "Справа окно, а слева дверь."}
        ]
      },
      {
        title: "Đồ đạc trong phòng học (Аудитория)",
        rows: [
          {"word": "доска́", "gram": "danh từ giống cái", "meaning": "cái bảng", "example": "На доске написано слово."},
          {"word": "дверь", "gram": "danh từ giống cái", "meaning": "cửa ra vào", "example": "Дверь закрыта."},
          {"word": "па́рта", "gram": "danh từ giống cái", "meaning": "bàn học sinh", "example": "Это моя парта."}
        ]
      }
    ],
    readingText: [
      "Это наша аудитория номер десять. Она большая, чистая и светлая.",
      "Прямо висит большая доска. На доске учитель пишет новые слова.",
      "Слева находится окно, а справа — дверь. Около окна стоит учительский стол.",
      "В центре стоят парты. Здесь занимаются студенты.",
      "Где сидит Анна? Анна сидит впереди. А где сидит Антон? Антон сидит сзади."
    ],
    questions: [
      {
        id: "a1_r4_q1", lang: "ru",
        prompt: "Какая аудитория номер десять?",
        hint: "Tìm các tính từ miêu tả phòng học ở câu đầu.",
        options: [{"key":"A","text":"Маленькая и тёмная"},{"key":"B","text":"Большая, чистая и светлая"},{"key":"C","text":"Старая"},{"key":"D","text":"Шумная"}],
        ans: "B",
        explain: "Bài viết: «Она большая, чистая и светлая»."
      },
      {
        id: "a1_r4_q2", lang: "ru",
        prompt: "Что висит прямо?",
        hint: "Đọc câu 2.",
        options: [{"key":"A","text":"Часы"},{"key":"B","text":"Карта"},{"key":"C","text":"Большая доска"},{"key":"D","text":"Картина"}],
        ans: "C",
        explain: "Đoạn 2 nêu rõ: «Прямо висит большая доска» (ngay phía trước treo một cái bảng lớn)."
      },
      {
        id: "a1_r4_q3", lang: "ru",
        prompt: "Где сидит Анна?",
        hint: "Xem vị trí ngồi của Anna.",
        options: [{"key":"A","text":"Сзади"},{"key":"B","text":"Впереди"},{"key":"C","text":"Справа"},{"key":"D","text":"Около двери"}],
        ans: "B",
        explain: "Đoạn cuối viết: «Анна сидит впереди» (Anna ngồi phía trước)."
      },
      {
        id: "a1_r4_q4", lang: "ru",
        prompt: "Что находится слева в аудитории?",
        hint: "Xem từ «слева».",
        options: [{"key":"A","text":"Дверь"},{"key":"B","text":"Окно"},{"key":"C","text":"Доска"},{"key":"D","text":"Шкаф"}],
        ans: "B",
        explain: "Bài viết: «Слева находится окно, а справа — дверь»."
      },
      {
        id: "a1_r4_q5", lang: "ru",
        prompt: "Что делает учитель на доске?",
        hint: "Tìm hành động của giáo viên.",
        options: [{"key":"A","text":"Рисует картину"},{"key":"B","text":"Пишет новые слова"},{"key":"C","text":"Читает книгу"},{"key":"D","text":"Отдыхает"}],
        ans: "B",
        explain: "Bài viết: «На доске учитель пишет новые слова» (thầy giáo viết từ mới lên bảng)."
      },
      {
        id: "a1_r4_q6", lang: "vi",
        prompt: "Từ «справа» trong tiếng Nga có nghĩa là gì?",
        hint: "Ngược nghĩa với «слева» (bên trái).",
        options: [{"key":"A","text":"Ở bên phải"},{"key":"B","text":"Ở bên trái"},{"key":"C","text":"Ở đằng sau"},{"key":"D","text":"Ở phía trên"}],
        ans: "A",
        explain: "«Справа» nghĩa là ở phía bên phải."
      },
      {
        id: "a1_r4_q7", lang: "vi",
        prompt: "Từ «в центре» có nghĩa là gì?",
        hint: "Vị trí chính giữa.",
        options: [{"key":"A","text":"Ở góc phòng"},{"key":"B","text":"Ở ngoài hiên"},{"key":"C","text":"Ở trung tâm, ở giữa"},{"key":"D","text":"Ở tầng hai"}],
        ans: "C",
        explain: "«В центре» nghĩa là ở vị trí trung tâm, chính giữa."
      },
      {
        id: "a1_r4_q8", lang: "vi",
        prompt: "Từ «парта» dùng để chỉ loại bàn nào?",
        hint: "Bàn dùng trong trường học.",
        options: [{"key":"A","text":"Bàn ăn cơm"},{"key":"B","text":"Bàn làm việc giám đốc"},{"key":"C","text":"Bàn học sinh / sinh viên"},{"key":"D","text":"Bàn trà phòng khách"}],
        ans: "C",
        explain: "«Парта» là bàn học sinh/sinh viên trong lớp học."
      },
      {
        id: "a1_r4_q9", lang: "vi",
        prompt: "Anton ngồi ở vị trí nào trong lớp?",
        hint: "Đọc câu: «Антон сидит сзади».",
        options: [{"key":"A","text":"Ngồi ở dãy bàn đầu"},{"key":"B","text":"Ngồi ở phía sau (сзади)"},{"key":"C","text":"Ngồi cạnh cửa sổ"},{"key":"D","text":"Ngồi cạnh giáo viên"}],
        ans: "B",
        explain: "«Сзади» nghĩa là ở phía sau."
      },
      {
        id: "a1_r4_q10", lang: "vi",
        prompt: "Từ «чистая» miêu tả đặc điểm gì của phòng học?",
        hint: "Phòng học không bị bẩn.",
        options: [{"key":"A","text":"Rộng lớn"},{"key":"B","text":"Sạch sẽ"},{"key":"C","text":"Yên tĩnh"},{"key":"D","text":"Ấm áp"}],
        ans: "B",
        explain: "«Чистый / чистая» nghĩa là sạch sẽ."
      }
    ]
  },

  // Bài 5: Nghề nghiệp: Кем вы работаете?
  {
    id: 5,
    title: "Профессии: Кем вы работаете?",
    topic: "Nghề nghiệp & Công việc thường nhật",
    readingTime: "12 phút",
    vocabCategories: [
      {
        title: "Danh từ chỉ nghề nghiệp (Профессии)",
        rows: [
          {"word": "врач", "gram": "danh từ", "meaning": "bác sĩ", "example": "Мой папа — врач."},
          {"word": "инжене́р", "gram": "danh từ", "meaning": "kỹ sư", "example": "Он опытный инженер."},
          {"word": "учи́тель", "gram": "danh từ", "meaning": "giáo viên (nam)", "example": "Учитель объясняет урок."},
          {"word": "журнали́ст", "gram": "danh từ", "meaning": "nhà báo", "example": "Она талантливый журналист."}
        ]
      },
      {
        title: "Động từ về công việc (Работа)",
        rows: [
          {"word": "рабо́тать", "gram": "động từ nhóm I", "meaning": "làm việc", "example": "Где вы работаете?"},
          {"word": "заво́д", "gram": "danh từ", "meaning": "nhà máy, xí nghiệp", "example": "Он работает на заводе."},
          {"word": "больни́ца", "gram": "danh từ", "meaning": "bệnh viện", "example": "Врач работает в больнице."}
        ]
      }
    ],
    readingText: [
      "Каждый человек имеет свою профессию. Кем работают мои друзья?",
      "Мой друг Борис — инженер. Он работает на большом автомобильном заводе.",
      "Елена — врач. Она работает в центральной больнице и лечит людей.",
      "Сергей — журналист. Он пишет интересные статьи для газеты и журналов.",
      "А я пока студент. Я учусь и мечтаю стать программистом."
    ],
    questions: [
      {
        id: "a1_r5_q1", lang: "ru",
        prompt: "Кем работает Борис?",
        hint: "Tìm tên Boris ở đoạn 2.",
        options: [{"key":"A","text":"Врач"},{"key":"B","text":"Инженер"},{"key":"C","text":"Журналист"},{"key":"D","text":"Программист"}],
        ans: "B",
        explain: "Đoạn 2 nêu rõ: «Мой друг Борис — инженер»."
      },
      {
        id: "a1_r5_q2", lang: "ru",
        prompt: "Где работает Елена?",
        hint: "Tìm nơi làm việc của Elena.",
        options: [{"key":"A","text":"В школе"},{"key":"B","text":"На заводе"},{"key":"C","text":"В центральной больнице"},{"key":"D","text":"В банке"}],
        ans: "C",
        explain: "Đoạn 3 viết: «Елена — врач. Она работает в центральной больнице»."
      },
      {
        id: "a1_r5_q3", lang: "ru",
        prompt: "Что делает журналист Сергей?",
        hint: "Tìm công việc của Sergei.",
        options: [{"key":"A","text":"Лечит людей"},{"key":"B","text":"Строит дома"},{"key":"C","text":"Пишет статьи для газеты"},{"key":"D","text":"Учит детей"}],
        ans: "C",
        explain: "Bài viết: «Сергей — журналист. Он пишет интересные статьи для газеты»."
      },
      {
        id: "a1_r5_q4", lang: "ru",
        prompt: "Кем мечтает стать автор текста?",
        hint: "Đọc câu cuối cùng.",
        options: [{"key":"A","text":"Инженером"},{"key":"B","text":"Программистом"},{"key":"C","text":"Врачом"},{"key":"D","text":"Учителем"}],
        ans: "B",
        explain: "Tác giả chia sẻ: «Я учусь и мечтаю стать программистом» (ước mơ trở thành lập trình viên)."
      },
      {
        id: "a1_r5_q5", lang: "ru",
        prompt: "Какое слово обозначает человека, который лечит больных?",
        hint: "Từ vựng chỉ người chữa bệnh.",
        options: [{"key":"A","text":"Инженер"},{"key":"B","text":"Учитель"},{"key":"C","text":"Врач"},{"key":"D","text":"Студент"}],
        ans: "C",
        explain: "«Врач» (bác sĩ) là người khám chữa bệnh."
      },
      {
        id: "a1_r5_q6", lang: "vi",
        prompt: "Từ «завод» trong tiếng Nga có nghĩa là gì?",
        hint: "Nơi kỹ sư Boris làm việc.",
        options: [{"key":"A","text":"Trường học"},{"key":"B","text":"Nhà máy, xí nghiệp"},{"key":"C","text":"Rạp chiếu phim"},{"key":"D","text":"Quán ăn"}],
        ans: "B",
        explain: "«Завод» nghĩa là nhà máy / xí nghiệp."
      },
      {
        id: "a1_r5_q7", lang: "vi",
        prompt: "Động từ «лечить» trong câu «лечит людей» có nghĩa là gì?",
        hint: "Hành động của bác sĩ đối với bệnh nhân.",
        options: [{"key":"A","text":"Dạy học"},{"key":"B","text":"Chữa bệnh, điều trị"},{"key":"C","text":"Bán thuốc"},{"key":"D","text":"Nấu ăn"}],
        ans: "B",
        explain: "«Лечить» nghĩa là chữa bệnh, điều trị cho con người."
      },
      {
        id: "a1_r5_q8", lang: "vi",
        prompt: "Từ «статья» mà nhà báo viết có nghĩa là gì?",
        hint: "Sản phẩm của nhà báo trên báo chí.",
        options: [{"key":"A","text":"Bản tin / bài báo"},{"key":"B","text":"Cuốn tiểu thuyết"},{"key":"C","text":"Bức tranh vẽ"},{"key":"D","text":"Bản hợp đồng"}],
        ans: "A",
        explain: "«Статья» là bài báo, bài viết chuyên luận."
      },
      {
        id: "a1_r5_q9", lang: "vi",
        prompt: "Khi muốn hỏi 'Bạn làm nghề gì?', người Nga hỏi thế nào?",
        hint: "Mẫu câu hỏi nghề nghiệp chuẩn A1.",
        options: [{"key":"A","text":"Кто вы по профессии? / Кем вы работаете?"},{"key":"B","text":"Где вы живёте?"},{"key":"C","text":"Сколько это стоит?"},{"key":"D","text":"Как вас зовут?"}],
        ans: "A",
        explain: "«Кто вы по профессии?» hoặc «Кем вы работаете?» là câu hỏi nghề nghiệp."
      },
      {
        id: "a1_r5_q10", lang: "vi",
        prompt: "Tác giả của bài viết hiện đang làm nghề gì?",
        hint: "Đọc câu: «А я пока...»",
        options: [{"key":"A","text":"Đã đi làm kỹ sư"},{"key":"B","text":"Hiện vẫn đang là sinh viên"},{"key":"C","text":"Làm nghề dạy học"},{"key":"D","text":"Làm phóng viên"}],
        ans: "B",
        explain: "Tác giả viết: «А я пока студент» (Còn tôi hiện vẫn đang là sinh viên)."
      }
    ]
  }
];

// Tạo các bài tiếp theo từ 6 đến 20 để đủ trọn vẹn 20 bài A1
const remainingLessons = [
  // Bài 6: Động từ nhóm I
  {
    id: 6,
    title: "Что вы делаете каждый день?",
    topic: "Các hoạt động hàng ngày & Động từ nhóm I",
    readingTime: "12 phút",
    vocabCategories: [
      {
        title: "Động từ nhóm I thì hiện tại (Глаголы I спряжения)",
        rows: [
          {"word": "чита́ть", "gram": "я читаю, ты читаешь", "meaning": "đọc", "example": "Я читаю газету."},
          {"word": "писа́ть", "gram": "я пишу, ты пишешь", "meaning": "viết", "example": "Студент пишет письмо."},
          {"word": "понима́ть", "gram": "я понимаю, ты понимаешь", "meaning": "hiểu", "example": "Вы понимаете по-русски?"},
          {"word": "де́лать", "gram": "я делаю, ты делаешь", "meaning": "làm", "example": "Что ты делаешь вечером?"}
        ]
      }
    ],
    readingText: [
      "Каждый день студенты много занимаются. Утром они слушают лекции.",
      "Преподаватель объясняет правило, а студенты внимательно слушают.",
      "Потом они читают тексты и отвечают на вопросы. Дома они делают домашнее задание.",
      "Вечером Антон отдыхает, слушает музыку и читает интересные книги."
    ],
    questions: [
      {id:"a1_r6_q1", lang:"ru", prompt:"Что делают студенты утром?", hint:"Đọc câu 1.", options:[{"key":"A","text":"Спят"},{"key":"B","text":"Слушают лекции"},{"key":"C","text":"Гуляют в парке"},{"key":"D","text":"Смотрят фильм"}], ans:"B", explain:"Bài viết: «Утром они слушают лекции»."},
      {id:"a1_r6_q2", lang:"ru", prompt:"Что делает преподаватель на уроке?", hint:"Tìm hành động của thầy giáo.", options:[{"key":"A","text":"Объясняет правило"},{"key":"B","text":"Отдыхает"},{"key":"C","text":"Пишет письмо домой"},{"key":"D","text":"Поёт песню"}], ans:"A", explain:"Bài viết: «Преподаватель объясняет правило» (giải thích quy tắc)."},
      {id:"a1_r6_q3", lang:"ru", prompt:"Форма глагола «читать» для местоимения «я»:", hint:"Chia động từ nhóm I cho ngôi я.", options:[{"key":"A","text":"Читает"},{"key":"B","text":"Читаешь"},{"key":"C","text":"Читаю"},{"key":"D","text":"Читают"}], ans:"C", explain:"Động từ nhóm I ngôi «я» có đuôi -ю $\\rightarrow$ «я читаю»."},
      {id:"a1_r6_q4", lang:"ru", prompt:"Что делает Антон вечером?", hint:"Xem đoạn cuối.", options:[{"key":"A","text":"Работает на заводе"},{"key":"B","text":"Отдыхает, слушает музыку и читает"},{"key":"C","text":"Спит весь вечер"},{"key":"D","text":"Пишет тест"}], ans:"B", explain:"Bài viết: «Вечером Антон отдыхает, слушает музыку и читает»."},
      {id:"a1_r6_q5", lang:"ru", prompt:"Что означает наречие «внимательно»?", hint:"Cách thức lắng nghe bài giảng.", options:[{"key":"A","text":"Нhanh chóng"},{"key":"B","text":"Chăm chú, chú ý"},{"key":"C","text":"Ồn ào"},{"key":"D","text":"Chậm chạp"}], ans:"B", explain:"«Внимательно» nghĩa là chăm chú, tập trung."},
      {id:"a1_r6_q6", lang:"vi", prompt:"Động từ «отвечать» trong câu «отвечают на вопросы» nghĩa là gì?", hint:"Hành động phản hồi câu hỏi.", options:[{"key":"A","text":"Đặt câu hỏi"},{"key":"B","text":"Trả lời câu hỏi"},{"key":"C","text":"Bỏ qua câu hỏi"},{"key":"D","text":"Quên câu hỏi"}], ans:"B", explain:"«Отвечать на вопросы» là trả lời các câu hỏi."},
      {id:"a1_r6_q7", lang:"vi", prompt:"Cụm từ «домашнее задание» có nghĩa là gì?", hint:"Việc cần làm ở nhà sau buổi học.", options:[{"key":"A","text":"Việc nhà (lau dọn)"},{"key":"B","text":"Bài tập về nhà"},{"key":"C","text":"Nấu bữa tối"},{"key":"D","text":"Đi chợ mua thức ăn"}], ans:"B", explain:"«Домашнее задание» nghĩa là bài tập về nhà."},
      {id:"a1_r6_q8", lang:"vi", prompt:"Ngôi «мы» (chúng tôi) đi với đuôi động từ nhóm I nào?", hint:"Đuôi chia hiện tại của ngôi chúng tôi.", options:[{"key":"A","text":"-ем / -ём"},{"key":"B","text":"-ете"},{"key":"C","text":"-ут / -ют"},{"key":"D","text":"-ите"}], ans:"A", explain:"Ngôi «мы» chia với đuôi -ем/-ём (ví dụ: мы читаем, мы делаем)."},
      {id:"a1_r6_q9", lang:"vi", prompt:"Từ «утром» chỉ thời điểm nào trong ngày?", hint:"Buổi đầu tiên trong ngày.", options:[{"key":"A","text":"Vào buổi sáng"},{"key":"B","text":"Vào buổi trưa"},{"key":"C","text":"Vào buổi chiều"},{"key":"D","text":"Vào ban đêm"}], ans:"A", explain:"«Утром» nghĩa là vào buổi sáng."},
      {id:"a1_r6_q10", lang:"vi", prompt:"Từ «каждый день» có nghĩa là gì?", hint:"Tần suất diễn ra hàng ngày.", options:[{"key":"A","text":"Hàng tuần"},{"key":"B","text":"Hàng tháng"},{"key":"C","text":"Mỗi ngày, hàng ngày"},{"key":"D","text":"Thỉnh thoảng"}], ans:"C", explain:"«Каждый день» nghĩa là mỗi ngày, ngày nào cũng vậy."}
    ]
  },

  // Bài 7: Bạn có nói tiếng Nga không?
  {
    id: 7,
    title: "Вы говорите по-русски?",
    topic: "Ngôn ngữ & Khả năng giao tiếp: Говорить, знать, понимать",
    readingTime: "12 phút",
    vocabCategories: [
      {
        title: "Từ vựng về ngôn ngữ (Языки и наречия)",
        rows: [
          {"word": "по-ру́сски", "gram": "phó từ", "meaning": "bằng tiếng Nga", "example": "Я говорю по-русски."},
          {"word": "по-вьетна́мски", "gram": "phó từ", "meaning": "bằng tiếng Việt", "example": "Он понимает по-вьетнамски."},
          {"word": "хорошо́ / пло́хо", "gram": "phó từ", "meaning": "tốt, giỏi / kém, tồi", "example": "Она хорошо говорит."},
          {"word": "чуть-чу́ть", "gram": "phó từ", "meaning": "một chút, một ít", "example": "Я знаю чуть-чуть."}
        ]
      }
    ],
    readingText: [
      "— Здравствуйте! Скажите, пожалуйста, вы говорите по-русски?\n— Да, я говорю по-русски, но пока ещё немного медленно.",
      "Меня зовут Дэвид. Я приехал из Англии. Мой родной язык — английский.",
      "Сейчас я живу в Москве и изучаю русский язык в институте. Русский язык очень красивый, но трудный.",
      "Моя подруга Линь из Вьетнама. Она говорит по-вьетнамски, по-английски и уже хорошо понимает по-русски.",
      "Мы часто говорим по-русски на уроках, потому что практика очень важна."
    ],
    questions: [
      {id:"a1_r7_q1", lang:"ru", prompt:"Откуда приехал Дэвид?", hint:"Đọc đoạn 2.", options:[{"key":"A","text":"Из Франции"},{"key":"B","text":"Из Англии"},{"key":"C","text":"Из Вьетнама"},{"key":"D","text":"Из России"}], ans:"B", explain:"Bài viết: «Меня зовут Дэвид. Я приехал из Англии»."},
      {id:"a1_r7_q2", lang:"ru", prompt:"Какой родной язык у Дэвида?", hint:"Ngôn ngữ mẹ đẻ của David.", options:[{"key":"A","text":"Русский"},{"key":"B","text":"Вьетнамский"},{"key":"C","text":"Английский"},{"key":"D","text":"Немецкий"}], ans:"C", explain:"David nói: «Мой родной язык — английский»."},
      {id:"a1_r7_q3", lang:"ru", prompt:"На скольких языках говорит Линь?", hint:"Đếm số ngôn ngữ Linh biết.", options:[{"key":"A","text":"Только на одном"},{"key":"B","text":"По-вьетнамски, по-английски и понимает по-русски"},{"key":"C","text":"На пяти языках"},{"key":"D","text":"Она не говорит"}], ans:"B", explain:"Linh nói tiếng Việt, tiếng Anh và hiểu tốt tiếng Nga."},
      {id:"a1_r7_q4", lang:"ru", prompt:"Что Дэвид думает о русском языке?", hint:"Xem nhận xét về tiếng Nga ở đoạn 3.", options:[{"key":"A","text":"Он некрасивый"},{"key":"B","text":"Он очень красивый, но трудный"},{"key":"C","text":"Он очень лёгкий"},{"key":"D","text":"Он скучный"}], ans:"B", explain:"David nhận xét: «Русский язык очень красивый, но трудный» (rất hay nhưng khó)."},
      {id:"a1_r7_q5", lang:"ru", prompt:"Почему студенты часто говорят по-русски на уроках?", hint:"Đọc câu cuối cùng.", options:[{"key":"A","text":"Потому что практика очень важна"},{"key":"B","text":"Потому что они не знают других слов"},{"key":"C","text":"Потому что это игра"},{"key":"D","text":"Потому что так надо"}], ans:"A", explain:"Bài viết: «Потому что практика очень важна» (vì luyện tập thực hành rất quan trọng)."},
      {id:"a1_r7_q6", lang:"vi", prompt:"Từ «медленно» trong câu «говорю ещё немного медленно» nghĩa là gì?", hint:"Trái nghĩa với «быстро» (nhanh).", options:[{"key":"A","text":"Trôi chảy"},{"key":"B","text":"Chậm chạp, từ tốn"},{"key":"C","text":"Lưu loát"},{"key":"D","text":"Rất nhanh"}], ans:"B", explain:"«Медленно» nghĩa là chậm, từ từ."},
      {id:"a1_r7_q7", lang:"vi", prompt:"Cụm từ «родной язык» có nghĩa là gì?", hint:"Tiếng nói của dân tộc mình từ lúc sinh ra.", options:[{"key":"A","text":"Ngoại ngữ thứ hai"},{"key":"B","text":"Tiếng mẹ đẻ"},{"key":"C","text":"Tiếng lóng"},{"key":"D","text":"Ngôn ngữ cổ đại"}], ans:"B", explain:"«Родной язык» là tiếng mẹ đẻ."},
      {id:"a1_r7_q8", lang:"vi", prompt:"Phó từ «чуть-чуть» tương đương với từ nào sau đây?", hint:"Chỉ mức độ rất nhỏ.", options:[{"key":"A","text":"Немного (một chút)"},{"key":"B","text":"Очень много"},{"key":"C","text":"Всегда"},{"key":"D","text":"Никогда"}], ans:"A", explain:"«Чуть-чуть» nghĩa là một chút, một ít (đồng nghĩa với немного)."},
      {id:"a1_r7_q9", lang:"vi", prompt:"Động từ «говорить» đi với dạng nào sau đây khi nói về ngôn ngữ?", hint:"Dạng phó từ bắt đầu bằng по-.", options:[{"key":"A","text":"Говорить русский"},{"key":"B","text":"Говорить по-русски"},{"key":"C","text":"Говорить в русский"},{"key":"D","text":"Говорить на русский"}], ans:"B", explain:"Trong tiếng Nga chuẩn: «говорить по-русски» (nói bằng tiếng Nga)."},
      {id:"a1_r7_q10", lang:"vi", prompt:"Từ «трудный» có nghĩa là gì?", hint:"Miêu tả sự không dễ dàng.", options:[{"key":"A","text":"Dễ hiểu"},{"key":"B","text":"Khó khăn, phức tạp"},{"key":"C","text":"Ngắn gọn"},{"key":"D","text":"Vui vẻ"}], ans:"B", explain:"«Трудный» nghĩa là khó, gian nan."},
    ]
  },

  // Bài 8: Gia đình tôi
  {
    id: 8,
    title: "Моя семья: Большая и дружная",
    topic: "Gia đình, Họ hàng & Tình cảm gia đình",
    readingTime: "12 phút",
    vocabCategories: [
      {
        title: "Thành viên gia đình (Члены семьи)",
        rows: [
          {"word": "па́па / оте́ц", "gram": "danh từ giống đực", "meaning": "bố, ba, cha", "example": "Мой папа добрый."},
          {"word": "ма́ма / мать", "gram": "danh từ giống cái", "meaning": "mẹ, má", "example": "Мама вкусно готовит."},
          {"word": "брат / сестра́", "gram": "danh từ", "meaning": "anh (em) trai / chị (em) gái", "example": "У меня есть старший брат."},
          {"word": "де́душка / ба́бушка", "gram": "danh từ", "meaning": "ông nội, ngoại / bà nội, ngoại", "example": "Дедушка и бабушка на пенсии."}
        ]
      }
    ],
    readingText: [
      "У меня большая и дружная семья. Нас пять человек: папа, мама, старший брат, младшая сестра и я.",
      "Моего папу зовут Андрей. Он работает инженером. Мою маму зовут Светлана, она преподаватель музыки.",
      "Мой брат Виктор уже взрослый, ему двадцать четыре года. Он программист и работает в крупной компании.",
      "Моей сестре Маше всего восемь лет, она учится во втором классе. Она очень любит рисовать кошек.",
      "По воскресеньям мы все вместе собираемся на обед. Мы очень любим нашу семью!"
    ],
    questions: [
      {id:"a1_r8_q1", lang:"ru", prompt:"Сколько человек в семье автора?", hint:"Đọc câu: «Нас... человек».", options:[{"key":"A","text":"Три человека"},{"key":"B","text":"Четыре человека"},{"key":"C","text":"Пять человек"},{"key":"D","text":"Шесть человек"}], ans:"C", explain:"Bài viết: «Нас пять человек»."},
      {id:"a1_r8_q2", lang:"ru", prompt:"Кем работает мама Светлана?", hint:"Tìm nghề của mẹ.", options:[{"key":"A","text":"Инженер"},{"key":"B","text":"Преподаватель музыки"},{"key":"C","text":"Врач"},{"key":"D","text":"Программист"}], ans:"B", explain:"Bài viết: «Мою маму зовут Светлана, она преподаватель музыки»."},
      {id:"a1_r8_q3", lang:"ru", prompt:"Сколько лет младшей сестре Маше?", hint:"Tìm tuổi của Masha.", options:[{"key":"A","text":"Пять лет"},{"key":"B","text":"Восемь лет"},{"key":"C","text":"Десять лет"},{"key":"D","text":"Двенадцать лет"}], ans:"B", explain:"Bài viết: «Моей сестре Маше всего восемь лет»."},
      {id:"a1_r8_q4", lang:"ru", prompt:"Что любит рисовать Маша?", hint:"Xem sở thích vẽ của cô bé.", options:[{"key":"A","text":"Машины"},{"key":"B","text":"Цветы"},{"key":"C","text":"Кошек"},{"key":"D","text":"Дома"}], ans:"C", explain:"Bài viết: «Она очень любит рисовать кошек» (vẽ mèo)."},
      {id:"a1_r8_q5", lang:"ru", prompt:"Когда семья собирается вместе на обед?", hint:"Tìm ngày trong tuần.", options:[{"key":"A","text":"По понедельникам"},{"key":"B","text":"По пятницам"},{"key":"C","text":"По воскресеньям"},{"key":"D","text":"Каждый вечер"}], ans:"C", explain:"Bài viết: «По воскресеньям мы все вместе собираемся на обед» (vào các ngày chủ nhật)."},
      {id:"a1_r8_q6", lang:"vi", prompt:"Từ «старший брат» có nghĩa là gì?", hint:"Phân biệt старший và младший.", options:[{"key":"A","text":"Em trai nhỏ"},{"key":"B","text":"Anh trai lớn"},{"key":"C","text":"Bác ruột"},{"key":"D","text":"Bạn thân"}], ans:"B", explain:"«Старший брат» là anh trai lớn (ngược lại: младший брат là em trai)."},
      {id:"a1_r8_q7", lang:"vi", prompt:"Tính từ «дружная» trong «дружная семья» mang ý nghĩa gì?", hint:"Gia đình yêu thương gắn kết.", options:[{"key":"A","text":"Giàu có"},{"key":"B","text":"Hòa thuận, gắn bó, đoàn kết"},{"key":"C","text":"Nghiêm khắc"},{"key":"D","text":"Đông đúc"}], ans:"B", explain:"«Дружная семья» là gia đình hòa thuận, đầm ấm, yêu thương nhau."},
      {id:"a1_r8_q8", lang:"vi", prompt:"Anh trai Viktor làm nghề gì?", hint:"Đọc đoạn về Viktor.", options:[{"key":"A","text":"Bác sĩ"},{"key":"B","text":"Giáo viên"},{"key":"C","text":"Lập trình viên (программист)"},{"key":"D","text":"Nhà báo"}], ans:"C", explain:"Viktor là lập trình viên (программист) làm việc cho công ty lớn."},
      {id:"a1_r8_q9", lang:"vi", prompt:"Từ «воскресенье» chỉ thứ mấy trong tuần?", hint:"Ngày nghỉ cuối tuần.", options:[{"key":"A","text":"Thứ Bảy"},{"key":"B","text":"Chủ Nhật"},{"key":"C","text":"Thứ Hai"},{"key":"D","text":"Thứ Sáu"}], ans:"B", explain:"«Воскресенье» là ngày Chủ Nhật."},
      {id:"a1_r8_q10", lang:"vi", prompt:"Mẹ của tác giả tên là gì?", hint:"Tìm tên Svetlana.", options:[{"key":"A","text":"Anna"},{"key":"B","text":"Svetlana"},{"key":"C","text":"Elena"},{"key":"D","text":"Maria"}], ans:"B", explain:"«Мою маму зовут Светлана»."}
    ]
  },

  // Bài 9: Ngôi nhà & Căn hộ của chúng tôi
  {
    id: 9,
    title: "Наш дом и наша квартира",
    topic: "Nhà ở, Phòng ốc & Tiện nghi gia đình",
    readingTime: "12 phút",
    vocabCategories: [
      {
        title: "Các phòng trong nhà (Комнаты)",
        rows: [
          {"word": "ку́хня", "gram": "danh từ giống cái", "meaning": "phòng bếp", "example": "Мама готовит на кухне."},
          {"word": "спа́льня", "gram": "danh từ giống cái", "meaning": "phòng ngủ", "example": "В спальне стоит кровать."},
          {"word": "ванная", "gram": "danh từ giống cái", "meaning": "phòng tắm", "example": "Ванная чистая и светлая."},
          {"word": "балко́н", "gram": "danh từ giống đực", "meaning": "ban công", "example": "На балконе растут цветы."}
        ]
      }
    ],
    readingText: [
      "Мы живём в новом многоэтажном доме. Наш дом находится рядом с красивым парком.",
      "Наша квартира расположена на пятом этаже. Здесь есть лифт, поэтому подниматься легко.",
      "В квартире есть прихожая, кухня, гостиная và hai phòng ngủ.",
      "Моя любимая комната — это кухня. Она светлая, там стоит холодильник, плита và bàn ăn lớn.",
      "С балкона открывается прекрасный вид на город. Вечером мы любим смотреть на огни города."
    ],
    questions: [
      {id:"a1_r9_q1", lang:"ru", prompt:"На каком этаже находится квартира?", hint:"Tìm số tầng.", options:[{"key":"A","text":"На первом"},{"key":"B","text":"На третьем"},{"key":"C","text":"На пятом"},{"key":"D","text":"На десятом"}], ans:"C", explain:"Bài viết: «Наша квартира расположена на пятом этаже» (tầng 5)."},
      {id:"a1_r9_q2", lang:"ru", prompt:"Что находится рядом с домом?", hint:"Xem vị trí xung quanh nhà.", options:[{"key":"A","text":"Школа"},{"key":"B","text":"Красивый парк"},{"key":"C","text":"Завод"},{"key":"D","text":"Вокзал"}], ans:"B", explain:"Bài viết: «Наш дом находится рядом с красивым парком»."},
      {id:"a1_r9_q3", lang:"ru", prompt:"Какая комната самая любимая у автора?", hint:"Tìm căn phòng yêu thích.", options:[{"key":"A","text":"Спальня"},{"key":"B","text":"Гостиная"},{"key":"C","text":"Кухня"},{"key":"D","text":"Ванная"}], ans:"C", explain:"Tác giả viết: «Моя любимая комната — это кухня»."},
      {id:"a1_r9_q4", lang:"ru", prompt:"Что стоит на кухне?", hint:"Xem đồ gia dụng trong bếp.", options:[{"key":"A","text":"Холодильник, плита и большой стол"},{"key":"B","text":"Диван и телевизор"},{"key":"C","text":"Кровать и шкаф"},{"key":"D","text":"Пианино"}], ans:"A", explain:"Bài viết: «Там стоит холодильник, плита và bàn ăn lớn»."},
      {id:"a1_r9_q5", lang:"ru", prompt:"Откуда открывается красивый вид на город?", hint:"Nơi ngắm nhìn thành phố.", options:[{"key":"A","text":"С кухни"},{"key":"B","text":"С балкона"},{"key":"C","text":"Из ванной"},{"key":"D","text":"Из коридора"}], ans:"B", explain:"Bài viết: «С балкона открывается прекрасный вид на город»."},
      {id:"a1_r9_q6", lang:"vi", prompt:"Từ «этаж» trong tiếng Nga có nghĩa là gì?", hint:"Tầng trong tòa nhà.", options:[{"key":"A","text":"Căn phòng"},{"key":"B","text":"Tầng gác / tầng lầu"},{"key":"C","text":"Cầu thang máy"},{"key":"D","text":"Mái nhà"}], ans:"B", explain:"«Этаж» nghĩa là tầng (nhà)."},
      {id:"a1_r9_q7", lang:"vi", prompt:"Từ «лифт» chỉ tiện nghi nào trong tòa nhà cao tầng?", hint:"Phương tiện di chuyển lên xuống.", options:[{"key":"A","text":"Thang cuốn"},{"key":"B","text":"Thang máy"},{"key":"C","text":"Thang bộ"},{"key":"D","text":"Cầu vượt"}], ans:"B", explain:"«Лифт» là thang máy."},
      {id:"a1_r9_q8", lang:"vi", prompt:"Từ «холодильник» trong phòng bếp là thiết bị gì?", hint:"Thiết bị làm mát và giữ đồ ăn tươi.", options:[{"key":"A","text":"Lò vi sóng"},{"key":"B","text":"Tủ lạnh"},{"key":"C","text":"Máy rửa bát"},{"key":"D","text":"Nồi cơm điện"}], ans:"B", explain:"«Холодильник» nghĩa là tủ lạnh."},
      {id:"a1_r9_q9", lang:"vi", prompt:"Từ «прихожая» dùng để chỉ không gian nào trong căn hộ?", hint:"Nơi đầu tiên bước vào để cởi áo khoác và giày dép.", options:[{"key":"A","text":"Phòng khách"},{"key":"B","text":"Tiền sảnh, hành lang lối vào"},{"key":"C","text":"Phòng ngủ"},{"key":"D","text":"Gác xép"}], ans:"B", explain:"«Прихожая» là hành lang lối vào nhà/tiền sảnh."},
      {id:"a1_r9_q10", lang:"vi", prompt:"Buổi tối gia đình tác giả thích ngắm điều gì từ ban công?", hint:"Đọc câu cuối cùng.", options:[{"key":"A","text":"Ngắm ánh đèn thành phố (огни города)"},{"key":"B","text":"Ngắm trăng sao"},{"key":"C","text":"Xem người đi bộ"},{"key":"D","text":"Ngắm xe ô tô"}], ans:"A", explain:"«Смотреть на огни города» (ngắm ánh đèn lung linh của thành phố)."}
    ]
  },

  // Bài 10: Bạn sống ở đâu? (Cách 6)
  {
    id: 10,
    title: "Где вы живёте? Город и страна",
    topic: "Nơi chốn, Quốc gia & Cách 6 vị trí (в / на)",
    readingTime: "12 phút",
    vocabCategories: [
      {
        title: "Cách 6 chỉ nơi chốn (Предложный падеж)",
        rows: [
          {"word": "в го́роде", "gram": "Cách 6", "meaning": "ở trong thành phố", "example": "Я живу в городе."},
          {"word": "в Москве́", "gram": "Cách 6", "meaning": "ở Moskva", "example": "Мы учимся в Москве."},
          {"word": "в Хано́е", "gram": "Cách 6", "meaning": "ở Hà Nội", "example": "Моя семья живёт в Ханое."},
          {"word": "на у́лице", "gram": "Cách 6", "meaning": "ở trên đường phố", "example": "Дети гуляют на улице."}
        ]
      }
    ],
    readingText: [
      "Здравствуйте! Меня зовут Чан. Я приехал из Вьетнама.",
      "Моя родина — Вьетнам. Мои родители сейчас живут в столице, в городе Ханое.",
      "А я уже восемь месяцев живу и учусь в России, в Москве.",
      "Москва — это огромный, красивый и старинный город. Я живу в студенческом общежитии на улице Академика Волгина.",
      "Мне очень нравится жить в России. Здесь у меня много хороших друзей."
    ],
    questions: [
      {id:"a1_r10_q1", lang:"ru", prompt:"Откуда приехал Чан?", hint:"Xem câu 2.", options:[{"key":"A","text":"Из Китая"},{"key":"B","text":"Из Вьетнама"},{"key":"C","text":"Из Японии"},{"key":"D","text":"Из Кореи"}], ans:"B", explain:"Bài viết: «Я приехал из Вьетнама»."},
      {id:"a1_r10_q2", lang:"ru", prompt:"Где сейчас живут родители Чана?", hint:"Tìm nơi ở của bố mẹ.", options:[{"key":"A","text":"В Москве"},{"key":"B","text":"В Ханое"},{"key":"C","text":"В Сайгоне"},{"key":"D","text":"В деревне"}], ans:"B", explain:"Bài viết: «Мои родители сейчас живут в столице, в городе Ханое»."},
      {id:"a1_r10_q3", lang:"ru", prompt:"Сколько месяцев Чан живёт в Москве?", hint:"Tìm số tháng.", options:[{"key":"A","text":"Два месяца"},{"key":"B","text":"Пять месяцев"},{"key":"C","text":"Восемь месяцев"},{"key":"D","text":"Один год"}], ans:"C", explain:"Bài viết: «А я уже восемь месяцев живу и учусь в России»."},
      {id:"a1_r10_q4", lang:"ru", prompt:"Где живёт Чан в Москве?", hint:"Tìm nơi ở sinh viên.", options:[{"key":"A","text":"В гостинице"},{"key":"B","text":"В студенческом общежитии"},{"key":"C","text":"В отдельной квартире"},{"key":"D","text":"У друзей"}], ans:"B", explain:"Bài viết: «Я живу в студенческом общежитии» (ký túc xá sinh viên)."},
      {id:"a1_r10_q5", lang:"ru", prompt:"Как Чан описывает Москву?", hint:"Tìm các tính từ miêu tả Moskva.", options:[{"key":"A","text":"Маленький и тихий город"},{"key":"B","text":"Огромный, красивый и старинный город"},{"key":"C","text":"Скучный город"},{"key":"D","text":"Тёмный город"}], ans:"B", explain:"Bài viết: «Москва — это огромный, красивый и старинный город»."},
      {id:"a1_r10_q6", lang:"vi", prompt:"Từ «столица» trong tiếng Nga có nghĩa là gì?", hint:"Thành phố trung tâm đầu não của một quốc gia.", options:[{"key":"A","text":"Thủ đô"},{"key":"B","text":"Ngoại ô"},{"key":"C","text":"Hải cảng"},{"key":"D","text":"Tỉnh lẻ"}], ans:"A", explain:"«Столица» nghĩa là thủ đô."},
      {id:"a1_r10_q7", lang:"vi", prompt:"Từ «общежитие» có nghĩa là gì?", hint:"Nơi sinh viên các trường đại học thường ở.", options:[{"key":"A","text":"Khách sạn 5 sao"},{"key":"B","text":"Ký túc xá"},{"key":"C","text":"Nhà riêng"},{"key":"D","text":"Phòng trọ tư nhân"}], ans:"B", explain:"«Общежитие» (viết tắt là общага) nghĩa là ký túc xá."},
      {id:"a1_r10_q8", lang:"vi", prompt:"Đuôi đặc trưng của danh từ giống đực và cái ở Cách 6 chỉ nơi chốn là gì?", hint:"Ví dụ: в город..., в Москв... .", options:[{"key":"A","text":"Đuôi -е"},{"key":"B","text":"Đuôi -у"},{"key":"C","text":"Đuôi -ом"},{"key":"D","text":"Đuôi -а"}], ans:"A", explain:"Ở Cách 6 (Предложный падеж) chỉ nơi chốn, danh từ thường biến đổi đuôi thành -е (в городе, в Москве, в школе)."},
      {id:"a1_r10_q9", lang:"vi", prompt:"Từ «родина» mang ý nghĩa gì?", hint:"Mảnh đất thiêng liêng nơi ta sinh ra.", options:[{"key":"A","text":"Quê hương, tổ quốc"},{"key":"B","text":"Trường học cũ"},{"key":"C","text":"Ngôi nhà"},{"key":"D","text":"Thành phố"}], ans:"A", explain:"«Родина» là quê hương, đất mẹ, tổ quốc."},
      {id:"a1_r10_q10", lang:"vi", prompt:"Giới từ nào đi với từ «улица» khi nói 'ở trên đường phố'?", hint:"Xem câu: ... улице Академика Волгина.", options:[{"key":"A","text":"В"},{"key":"B","text":"На"},{"key":"C","text":"К"},{"key":"D","text":"Под"}], ans:"B", explain:"Với đường phố (улица), quảng trường (площадь), tiếng Nga dùng giới từ «НА» $\\rightarrow$ «на улице»."}
    ]
  }
];

// Thêm các bài 11 đến 20
const lessons11to20 = [
  // 11. Một ngày của tôi
  {
    id: 11,
    title: "Один день студента: Мой распорядок дня",
    topic: "Thời gian biểu trong ngày, Giờ giấc & Thói quen",
    readingTime: "12 phút",
    vocabCategories: [
      {
        title: "Buổi trong ngày & Động từ sinh hoạt",
        rows: [
          {"word": "просыпа́ться", "gram": "động từ", "meaning": "thức dậy", "example": "Я просыпаюсь в семь часов."},
          {"word": "за́втракать", "gram": "động từ", "meaning": "ăn sáng", "example": "Мы завтракаем в восемь."},
          {"word": "обе́дать", "gram": "động từ", "meaning": "ăn trưa", "example": "В час дня я обедаю."},
          {"word": "у́жинать", "gram": "động từ", "meaning": "ăn tối", "example": "Вечером семья ужинает."}
        ]
      }
    ],
    readingText: [
      "Обычно я просыпаюсь рано, в семь часов утра. Я умываюсь, делаю зарядку и завтракаю.",
      "На завтрак я ем яичницу, бутерброд с сыром и пью горячий кофе с молоком.",
      "В восемь часов я иду в университет. Занятия начинаются в половине девятого и заканчиваются в три часа дня.",
      "После занятий я обедаю в студенческой столовой с друзьями. Еда там вкусная и недорогая.",
      "Вечером я делаю домашнее задание, гуляю в парке, а в одиннадцать часов вечера ложусь спать."
    ],
    questions: [
      {id:"a1_r11_q1", lang:"ru", prompt:"Во сколько часов просыпается студент?", hint:"Xem giờ thức dậy.", options:[{"key":"A","text":"В шесть часов"},{"key":"B","text":"В семь часов"},{"key":"C","text":"В восемь часов"},{"key":"D","text":"В девять часов"}], ans:"B", explain:"Bài viết: «Обычно я просыпаюсь рано, в семь часов утра»."},
      {id:"a1_r11_q2", lang:"ru", prompt:"Что студент ест на завтрак?", hint:"Tìm món ăn sáng.", options:[{"key":"A","text":"Суп и рис"},{"key":"B","text":"Яичницу, бутерброд с сыром и кофе"},{"key":"C","text":"Фрукты и сок"},{"key":"D","text":"Пиццу"}], ans:"B", explain:"Bài viết: «На завтрак я ем яичницу, бутерброд с сыром и пью горячий кофе»."},
      {id:"a1_r11_q3", lang:"ru", prompt:"Где студент обедает после занятий?", hint:"Tìm nơi ăn trưa.", options:[{"key":"A","text":"В ресторане"},{"key":"B","text":"Дома"},{"key":"C","text":"В студенческой столовой"},{"key":"D","text":"В парке"}], ans:"C", explain:"Bài viết: «Я обедаю в студенческой столовой с друзьями»."},
      {id:"a1_r11_q4", lang:"ru", prompt:"Во сколько часов заканчиваются занятия?", hint:"Xem giờ tan học.", options:[{"key":"A","text":"В два часа"},{"key":"B","text":"В три часа дня"},{"key":"C","text":"В пять часов"},{"key":"D","text":"В шесть часов"}], ans:"B", explain:"Bài viết: «...и заканчиваются в три часа дня»."},
      {id:"a1_r11_q5", lang:"ru", prompt:"Во сколько часов студент ложится спать?", hint:"Xem giờ đi ngủ.", options:[{"key":"A","text":"В десять часов"},{"key":"B","text":"В одиннадцать часов вечера"},{"key":"C","text":"В полночь"},{"key":"D","text":"В девять часов"}], ans:"B", explain:"Bài viết: «а в одиннадцать часов вечера ложусь спать»."},
      {id:"a1_r11_q6", lang:"vi", prompt:"Từ «зарядка» trong câu «делаю зарядку» nghĩa là gì?", hint:"Hoạt động vận động thể chất buổi sáng.", options:[{"key":"A","text":"Sạc pin điện thoại"},{"key":"B","text":"Tập thể dục buổi sáng"},{"key":"C","text":"Dọn dẹp phòng"},{"key":"D","text":"Học bài"}], ans:"B", explain:"«Делать зарядку» nghĩa là tập thể dục buổi sáng."},
      {id:"a1_r11_q7", lang:"vi", prompt:"Từ «столовая» có nghĩa là gì?", hint:"Nơi phục vụ đồ ăn trong trường học hoặc cơ quan.", options:[{"key":"A","text":"Căng tin / Nhà ăn"},{"key":"B","text":"Quán bar"},{"key":"C","text":"Siêu thị"},{"key":"D","text":"Phòng thí nghiệm"}], ans:"A", explain:"«Столовая» là nhà ăn / căng tin."},
      {id:"a1_r11_q8", lang:"vi", prompt:"Từ «бутерброд» bắt nguồn từ tiếng Đức nghĩa là món gì?", hint:"Bánh mì kẹp phổ biến ở Nga.", options:[{"key":"A","text":"Bánh mì kẹp bơ / thịt / phô mai"},{"key":"B","text":"Bánh bao hấp"},{"key":"C","text":"Bánh ngọt sinh nhật"},{"key":"D","text":"Bánh quy"}], ans:"A", explain:"«Бутерброд» là món bánh mì lát phết bơ hoặc kẹp xúc xích, phô mai."},
      {id:"a1_r11_q9", lang:"vi", prompt:"Từ «недорогая» miêu tả đặc điểm gì của đồ ăn căng tin?", hint:"Trái nghĩa với «дорогая» (đắt đỏ).", options:[{"key":"A","text":"Rất đắt đỏ"},{"key":"B","text":"Giá cả phải chăng, bình dân"},{"key":"C","text":"Khó ăn"},{"key":"D","text":"Cay nồng"}], ans:"B", explain:"«Недорогой / недорогая» nghĩa là không đắt, giá rẻ/vừa túi tiền sinh viên."},
      {id:"a1_r11_q10", lang:"vi", prompt:"Cụm từ «половина девятого» chỉ mấy giờ?", hint:"Nửa tiếng trước 9 giờ.", options:[{"key":"A","text":"8 giờ 30 phút"},{"key":"B","text":"9 giờ 30 phút"},{"key":"C","text":"7 giờ 30 phút"},{"key":"D","text":"8 giờ đúng"}], ans:"A", explain:"«Половина девятого» nghĩa là 8 giờ rưỡi (nửa của giờ thứ 9)."}
    ]
  },

  // 12. Mua sắm tại cửa hàng
  {
    id: 12,
    title: "В магазине: Продукты и покупки",
    topic: "Mua sắm, Thực phẩm & Giá tiền (Сколько стоит?)",
    readingTime: "12 phút",
    vocabCategories: [
      {
        title: "Thực phẩm thông dụng (Продукты)",
        rows: [
          {"word": "хлеб", "gram": "danh từ giống đực", "meaning": "bánh mì", "example": "Купи свежий хлеб."},
          {"word": "молоко́", "gram": "danh từ giống trung", "meaning": "sữa tươi", "example": "Я пью молоко каждый день."},
          {"word": "сыр", "gram": "danh từ giống đực", "meaning": "phô mai", "example": "Вкусный сыр."},
          {"word": "я́блоко", "gram": "danh từ giống trung", "meaning": "quả táo", "example": "Свежие яблоки."}
        ]
      }
    ],
    readingText: [
      "— Здравствуйте! Чем я могу вам помочь?\n— Здравствуйте! Скажите, пожалуйста, сколько стоит этот белый хлеб?\n— Хлеб стоит сорок рублей.",
      "— Дайте, пожалуйста, одну буханку хлеба, пакет молока и кусок сыра.",
      "— Вот, пожалуйста. Что-нибудь ещё?\n— Да, взвесьте, пожалуйста, один килограмм красных яблок.\n— Хорошо. С вас всего двести пятьдесят рублей.",
      "— Вот деньги.\n— Ваша сдача и чек. Спасибо за покупку, приходите ещё!"
    ],
    questions: [
      {id:"a1_r12_q1", lang:"ru", prompt:"Сколько стоит белый хлеб?", hint:"Xem giá bánh mì ở lượt thoại đầu.", options:[{"key":"A","text":"20 рублей"},{"key":"B","text":"30 рублей"},{"key":"C","text":"40 рублей"},{"key":"D","text":"50 рублей"}], ans:"C", explain:"Người bán trả lời: «Хлеб стоит сорок (40) рублей»."},
      {id:"a1_r12_q2", lang:"ru", prompt:"Сколько яблок попросил покупатель?", hint:"Tìm lượng táo mua.", options:[{"key":"A","text":"Один килограмм"},{"key":"B","text":"Два килограмма"},{"key":"C","text":"Три яблока"},{"key":"D","text":"Полкило"}], ans:"A", explain:"Khách nói: «Взвесьте, пожалуйста, один килограмм красных яблок»."},
      {id:"a1_r12_q3", lang:"ru", prompt:"Какова общая стоимость всей покупки?", hint:"Xem tổng số tiền.", options:[{"key":"A","text":"150 рублей"},{"key":"B","text":"200 рублей"},{"key":"C","text":"250 рублей"},{"key":"D","text":"300 рублей"}], ans:"C", explain:"Thu ngân báo: «С вас всего двести пятьдесят (250) рублей»."},
      {id:"a1_r12_q4", lang:"ru", prompt:"Что продавец отдаёт покупателю вместе с чеком?", hint:"Tiền thừa trả lại.", options:[{"key":"A","text":"Подарок"},{"key":"B","text":"Сдачу"},{"key":"C","text":"Ещё одно яблоко"},{"key":"D","text":"Сумку"}], ans:"B", explain:"Thu ngân nói: «Ваша сдача (tiền thừa) и чек»."},
      {id:"a1_r12_q5", lang:"ru", prompt:"Какая фраза используется для вопроса о цене?", hint:"Câu hỏi giá tiền.", options:[{"key":"A","text":"Сколько это стоит?"},{"key":"B","text":"Где это лежит?"},{"key":"C","text":"Кто это сделал?"},{"key":"D","text":"Как дела?"}], ans:"A", explain:"«Сколько это стоит?» là câu hỏi 'Cái này giá bao nhiêu?'."},
      {id:"a1_r12_q6", lang:"vi", prompt:"Từ «сдача» trong việc mua bán có nghĩa là gì?", hint:"Số tiền gửi lại khi trả thừa.", options:[{"key":"A","text":"Hóa đơn thanh toán"},{"key":"B","text":"Tiền thừa trả lại"},{"key":"C","text":"Phiếu giảm giá"},{"key":"D","text":"Tiền boa (tip)"}], ans:"B", explain:"«Сдача» là tiền thừa trả lại cho khách."},
      {id:"a1_r12_q7", lang:"vi", prompt:"Từ «чек» có nghĩa là gì?", hint:"Giấy in xác nhận mua hàng.", options:[{"key":"A","text":"Hóa đơn, biên lai tính tiền"},{"key":"B","text":"Túi nilon"},{"key":"C","text":"Thẻ ngân hàng"},{"key":"D","text":"Tiền xu"}], ans:"A", explain:"«Чек» là hóa đơn / biên lai thanh toán."},
      {id:"a1_r12_q8", lang:"vi", prompt:"Lượng từ «буханка» đi với danh từ nào?", hint:"Đơn vị đếm bánh mì dạng ổ.", options:[{"key":"A","text":"Sữa"},{"key":"B","text":"Táo"},{"key":"C","text":"Bánh mì (ổ/ổ bánh mì)"},{"key":"D","text":"Nước khoáng"}], ans:"C", explain:"«Буханка хлеба» là một ổ bánh mì tròn/chữ nhật."},
      {id:"a1_r12_q9", lang:"vi", prompt:"Động từ «взвесить» (vui lòng cân giúp) dùng khi mua mặt hàng nào?", hint:"Mặt hàng tính theo kg.", options:[{"key":"A","text":"Sách báo"},{"key":"B","text":"Hoa quả, thịt cá theo cân"},{"key":"C","text":"Nước uống đóng chai"},{"key":"D","text":"Quần áo"}], ans:"B", explain:"«Взвесить» nghĩa là cân (trọng lượng)."},
      {id:"a1_r12_q10", lang:"vi", prompt:"Người bán hàng nói «Приходите ещё!» để bày tỏ điều gì?", hint:"Lời chào mời lịch sự của người Nga.", options:[{"key":"A","text":"Nhắc khách trả thêm tiền"},{"key":"B","text":"Mời lần sau lại ghé mua hàng"},{"key":"C","text":"Yêu cầu khách đi nhanh"},{"key":"D","text":"Báo cửa hàng sắp đóng cửa"}], ans:"B", explain:"«Приходите ещё!» nghĩa là 'Lần sau quý khách lại ghé nhé!'."}
    ]
  },

  // 13. Ở quán ăn: Gọi món
  {
    id: 13,
    title: "В кафе: Русская кухня и обед",
    topic: "Ẩm thực Nga, Món ăn & Lời chúc ngon miệng",
    readingTime: "12 phút",
    vocabCategories: [
      {
        title: "Các món ăn truyền thống Nga (Русские блюда)",
        rows: [
          {"word": "борщ", "gram": "danh từ giống đực", "meaning": "súp củ dền đỏ Nga", "example": "Борщ со сметаной очень вкусный."},
          {"word": "пельме́ни", "gram": "danh từ số nhiều", "meaning": "bánh xếp thịt kiểu Nga", "example": "Горячие пельмени."},
          {"word": "блины́", "gram": "danh từ số nhiều", "meaning": "bánh xèo blini của Nga", "example": "Блины с мёдом."},
          {"word": "морс", "gram": "danh từ giống đực", "meaning": "nước quả rừng mát ngọt", "example": "Ягодный морс."}
        ]
      }
    ],
    readingText: [
      "Сегодня суббота. Мы с друзьями решили пообедать в уютном ресторане русской кухни.",
      "В меню много традиционных блюд. На первое я заказал знаменитый русский борщ со сметаной.",
      "На второе друзья взяли сибирские пельмени, а я выбрал жареную рыбу с картофельным пюре.",
      "На десерт мы попробовали сладкие блины с клубничным вареньем và uống nước quả rừng (морс).",
      "— Приятного аппетита! — сказали мы друг другу. Обед был просто великолепным!"
    ],
    questions: [
      {id:"a1_r13_q1", lang:"ru", prompt:"В какой день недели друзья пошли в ресторан?", hint:"Tìm thứ trong tuần ở câu 1.", options:[{"key":"A","text":"В пятницу"},{"key":"B","text":"В субботу"},{"key":"C","text":"В воскресенье"},{"key":"D","text":"В среду"}], ans:"B", explain:"Bài viết: «Сегодня суббота...» (Hôm nay là thứ Bảy)."},
      {id:"a1_r13_q2", lang:"ru", prompt:"Что автор заказал на первое блюдо?", hint:"Món đầu tiên (món súp).", options:[{"key":"A","text":"Салат"},{"key":"B","text":"Борщ со сметаной"},{"key":"C","text":"Пельмени"},{"key":"D","text":"Блины"}], ans:"B", explain:"Bài viết: «На первое я заказал знаменитый русский борщ со сметаной»."},
      {id:"a1_r13_q3", lang:"ru", prompt:"Что друзья взяли на второе?", hint:"Món thứ hai (món chính).", options:[{"key":"A","text":"Сибирские пельмени"},{"key":"B","text":"Пиццу"},{"key":"C","text":"Бургер"},{"key":"D","text":"Кашу"}], ans:"A", explain:"Bài viết: «На второе друзья взяли сибирские пельмени»."},
      {id:"a1_r13_q4", lang:"ru", prompt:"Какой десерт ели друзья?", hint:"Món tráng miệng ngọt ngào.", options:[{"key":"A","text":"Мороженое"},{"key":"B","text":"Торт"},{"key":"C","text":"Сладкие блины с вареньем"},{"key":"D","text":"Шоколад"}], ans:"C", explain:"Bài viết: «На десерт мы попробовали сладкие блины с клубничным вареньем»."},
      {id:"a1_r13_q5", lang:"ru", prompt:"Что означает пожелание «Приятного аппетита!»?", hint:"Lời chúc trước khi ăn.", options:[{"key":"A","text":"Chúc ngon miệng"},{"key":"B","text":"Chúc ngủ ngon"},{"key":"C","text":"Chúc may mắn"},{"key":"D","text":"Chúc mừng sinh nhật"}], ans:"A", explain:"«Приятного аппетита!» nghĩa là chúc ăn ngon miệng."},
      {id:"a1_r13_q6", lang:"vi", prompt:"Món «борщ» của Nga được nấu từ loại củ đặc trưng nào tạo màu đỏ?", hint:"Loại củ màu đỏ tím thơm ngon.", options:[{"key":"A","text":"Cà rốt"},{"key":"B","text":"Củ dền đỏ"},{"key":"C","text":"Củ cải trắng"},{"key":"D","text":"Cà chua"}], ans:"B", explain:"Borsch là món súp củ dền đỏ nổi tiếng khắp thế giới của ẩm thực Đông Âu."},
      {id:"a1_r13_q7", lang:"vi", prompt:"«Сметана» thường ăn kèm với súp Borsch là gì?", hint:"Sản phẩm từ sữa lên men đặc trưng của Nga.", options:[{"key":"A","text":"Váng sữa chua (kem chua - sour cream)"},{"key":"B","text":"Phô mai que"},{"key":"C","text":"Bơ thực vật"},{"key":"D","text":"Sữa đặc"}], ans:"A", explain:"«Сметана» là kem chua (sour cream), linh hồn của ẩm thực Nga."},
      {id:"a1_r13_q8", lang:"vi", prompt:"Món «пельмени» có nét tương đồng với món ăn nào của châu Á?", hint:"Vỏ bột mì bọc nhân thịt bên trong luộc chín.", options:[{"key":"A","text":"Bánh chưng"},{"key":"B","text":"Bánh sủi cảo / hoành thánh"},{"key":"C","text":"Nem rán"},{"key":"D","text":"Bánh mì kẹp"}], ans:"B", explain:"Pelmeni là bánh xếp nhân thịt luộc, rất giống sủi cảo."},
      {id:"a1_r13_q9", lang:"vi", prompt:"Thức uống «морс» trong bài được làm từ nguyên liệu gì?", hint:"Nguyên liệu từ rừng nước Nga.", options:[{"key":"A","text":"Lá trà xanh"},{"key":"B","text":"Các loại quả mọng rừng (dâu dại, nam việt quất...)"},{"key":"C","text":"Hạt cà phê"},{"key":"D","text":"Sữa bò"}], ans:"B", explain:"Mors là thức uống nấu từ quả dâu rừng (cranberry, lingonberry...)."},
      {id:"a1_r13_q10", lang:"vi", prompt:"Từ «десерт» trong một bữa ăn dùng để chỉ điều gì?", hint:"Món ăn cuối cùng sau món chính.", options:[{"key":"A","text":"Món khai vị"},{"key":"B","text":"Món tráng miệng"},{"key":"C","text":"Món súp lỏng"},{"key":"D","text":"Khăn ăn"}], ans:"B", explain:"«Десерт» là món tráng miệng (bánh ngọt, hoa quả...)."}
    ]
  },

  // 14. Sở thích & Bạn bè
  {
    id: 14,
    title: "Мой друг и его увлечения",
    topic: "Sở thích cá nhân, Thể thao & Giải trí",
    readingTime: "12 phút",
    vocabCategories: [
      {
        title: "Sở thích & Thể thao (Хобби и спорт)",
        rows: [
          {"word": "игра́ть в футбо́л", "gram": "cụm từ", "meaning": "chơi đá bóng", "example": "Мальчики играют в футбол."},
          {"word": "слу́шать му́зыку", "gram": "cụm từ", "meaning": "nghe nhạc", "example": "Я люблю слушать рок."},
          {"word": "смотре́ть фи́льмы", "gram": "cụm từ", "meaning": "xem phim", "example": "Мы смотрим кино."},
          {"word": "фотографи́ровать", "gram": "động từ", "meaning": "chụp ảnh", "example": "Он любит фотографировать природу."}
        ]
      }
    ],
    readingText: [
      "У меня есть замечательный друг. Его зовут Денис. Мы дружим уже три года.",
      "Денис — очень активный и спортивный парень. Он обожает спорт и три раза в неделю играет в футбол.",
      "Кроме спорта, Денис отлично играет на гитаре và hát những bài hát tiếng Nga rất hay.",
      "По субботам мы часто ходим на стадион болеть за нашу любимую футбольную команду.",
      "Дружба — это самое ценное, что есть у человека. Я рад, что у меня есть такой верный друг!"
    ],
    questions: [
      {id:"a1_r14_q1", lang:"ru", prompt:"Как зовут друга автора?", hint:"Tìm tên người bạn.", options:[{"key":"A","text":"Иван"},{"key":"B","text":"Денис"},{"key":"C","text":"Антон"},{"key":"D","text":"Сергей"}], ans:"B", explain:"Bài viết: «Его зовут Денис»."},
      {id:"a1_r14_q2", lang:"ru", prompt:"Сколько лет они дружат?", hint:"Tìm thời gian kết bạn.", options:[{"key":"A","text":"Один год"},{"key":"B","text":"Два года"},{"key":"C","text":"Три года"},{"key":"D","text":"Пять лет"}], ans:"C", explain:"Bài viết: «Мы дружим уже три года»."},
      {id:"a1_r14_q3", lang:"ru", prompt:"Каким спортом занимается Денис три раза в неделю?", hint:"Tìm môn thể thao.", options:[{"key":"A","text":"Плаванием"},{"key":"B","text":"Футболом"},{"key":"C","text":"Теннисом"},{"key":"D","text":"Баскетболом"}], ans:"B", explain:"Bài viết: «...и три раза в неделю играет в футбол»."},
      {id:"a1_r14_q4", lang:"ru", prompt:"На каком музыкальном инструменте играет Денис?", hint:"Nhạc cụ Denis chơi giỏi.", options:[{"key":"A","text":"На пианино"},{"key":"B","text":"На скрипке"},{"key":"C","text":"На гитаре"},{"key":"D","text":"На барабанах"}], ans:"C", explain:"Bài viết: «Денис отлично играет на гитаре»."},
      {id:"a1_r14_q5", lang:"ru", prompt:"Куда ходят друзья по субботам?", hint:"Tìm địa điểm thể thao.", options:[{"key":"A","text":"В кино"},{"key":"B","text":"В музей"},{"key":"C","text":"На стадион"},{"key":"D","text":"В клуб"}], ans:"C", explain:"Bài viết: «По субботам мы часто ходим на стадион» (sân vận động)."},
      {id:"a1_r14_q6", lang:"vi", prompt:"Từ «стадион» trong tiếng Nga có nghĩa là gì?", hint:"Nơi diễn ra các trận thi đấu thể thao.", options:[{"key":"A","text":"Sân vận động"},{"key":"B","text":"Phòng tập gym"},{"key":"C","text":"Nhà ga xe lửa"},{"key":"D","text":"Bến xe buýt"}], ans:"A", explain:"«Стадион» nghĩa là sân vận động."},
      {id:"a1_r14_q7", lang:"vi", prompt:"Động từ «болеть за команду» có nghĩa là gì?", hint:"Hành động của cổ động viên.", options:[{"key":"A","text":"Bị ốm vì thời tiết"},{"key":"B","text":"Cổ vũ, ủng hộ cho đội bóng"},{"key":"C","text":"Mua vé xem bóng đá"},{"key":"D","text":"Làm huấn luyện viên"}], ans:"B", explain:"«Болеть за...» nghĩa là cổ vũ, làm fan hâm mộ của đội bóng."},
      {id:"a1_r14_q8", lang:"vi", prompt:"Khi nói 'chơi đàn guitar', giới từ nào đi với 'играть'?", hint:"Xem cụm từ: играть ... гитаре.", options:[{"key":"A","text":"Играть в гитару"},{"key":"B","text":"Играть на гитаре"},{"key":"C","text":"Играть с гитарой"},{"key":"D","text":"Играть за гитару"}], ans:"B", explain:"Chơi nhạc cụ trong tiếng Nga dùng: «играть НА + Cách 6» (играть на гитаре, на пианино). Chơi thể thao dùng «играть В + Cách 4» (играть в футбол)."},
      {id:"a1_r14_q9", lang:"vi", prompt:"Từ «верный друг» có nghĩa là gì?", hint:"Phẩm chất tốt đẹp của người bạn.", options:[{"key":"A","text":"Người bạn giàu có"},{"key":"B","text":"Người bạn trung thành, thủy chung"},{"key":"C","text":"Người bạn hài hước"},{"key":"D","text":"Người bạn mới quen"}], ans:"B", explain:"«Верный друг» là người bạn trung thành, đáng tin cậy."},
      {id:"a1_r14_q10", lang:"vi", prompt:"Từ «дружба» mang ý nghĩa gì?", hint:"Tình cảm giữa những người bạn.", options:[{"key":"A","text":"Tình bạn"},{"key":"B","text":"Tình yêu"},{"key":"C","text":"Tình đồng nghiệp"},{"key":"D","text":"Gia đình"}], ans:"A", explain:"«Дружба» nghĩa là tình bạn."}
    ]
  },

  // 15. Thời tiết & Bốn mùa
  {
    id: 15,
    title: "Какая сегодня погода? Времена года",
    topic: "Thời tiết, Bốn mùa & Khí hậu nước Nga",
    readingTime: "12 phút",
    vocabCategories: [
      {
        title: "Bốn mùa trong năm (Времена года)",
        rows: [
          {"word": "зима́ / зимо́й", "gram": "danh từ / phó từ", "meaning": "mùa đông / vào mùa đông", "example": "Зимой идёт снег."},
          {"word": "весна́ / весно́й", "gram": "danh từ / phó từ", "meaning": "mùa xuân / vào mùa xuân", "example": "Весной цветут деревья."},
          {"word": "ле́то / ле́том", "gram": "danh từ / phó từ", "meaning": "mùa hè / vào mùa hè", "example": "Летом тепло и солнечно."},
          {"word": "о́сень / о́сенью", "gram": "danh từ / phó từ", "meaning": "mùa thu / vào mùa thu", "example": "Осенью идёт дождь."}
        ]
      }
    ],
    readingText: [
      "В России есть четыре времени года: зима, весна, лето и осень. Каждое время года прекрасно по-своему.",
      "Зима в России холодная и снежная. На улице мороз, везде лежит белый пушистый снег. Люди катаются на лыжах и коньках.",
      "Весной природа просыпается. Тает снег, светит яркое солнце và cây cối đâm chồi nảy lộc.",
      "Лето — самое тёплое время года. Температура поднимается до тридцати градусов. Все едут отдыхать на море или на дачу.",
      "Осенью деревья становятся золотыми и красными. Часто идут дожди, но это очень поэтичное время."
    ],
    questions: [
      {id:"a1_r15_q1", lang:"ru", prompt:"Сколько времён года в России?", hint:"Đếm các mùa.", options:[{"key":"A","text":"Два"},{"key":"B","text":"Три"},{"key":"C","text":"Четыре"},{"key":"D","text":"Пять"}], ans:"C", explain:"Bài viết: «В России есть четыре времени года»."},
      {id:"a1_r15_q2", lang:"ru", prompt:"Какая погода зимой в России?", hint:"Xem đoạn 2.", options:[{"key":"A","text":"Жаркая и душная"},{"key":"B","text":"Холодная и снежная"},{"key":"C","text":"Дождливая"},{"key":"D","text":"Тёплая"}], ans:"B", explain:"Bài viết: «Зима в России холодная и снежная» (lạnh giá và nhiều tuyết)."},
      {id:"a1_r15_q3", lang:"ru", prompt:"Что делают люди зимой на улице?", hint:"Các môn thể thao mùa đông.", options:[{"key":"A","text":"Купаются в море"},{"key":"B","text":"Катаются на лыжах и коньках"},{"key":"C","text":"Загорают"},{"key":"D","text":"Собирают грибы"}], ans:"B", explain:"Bài viết: «Люди катаются на лыжах и коньках» (trượt tuyết và trượt băng)."},
      {id:"a1_r15_q4", lang:"ru", prompt:"Куда едут люди отдыхать летом?", hint:"Xem nơi nghỉ mát mùa hè.", options:[{"key":"A","text":"В горы на снег"},{"key":"B","text":"На море или на дачу"},{"key":"C","text":"В библиотеку"},{"key":"D","text":"Сидят дома"}], ans:"B", explain:"Bài viết: «Все едут отдыхать на море или на дачу» (ra biển hoặc về nhà nghỉ ngoại ô)."},
      {id:"a1_r15_q5", lang:"ru", prompt:"Какими становятся деревья осенью?", hint:"Màu sắc mùa thu vàng nước Nga.", options:[{"key":"A","text":"Зелёными"},{"key":"B","text":"Белыми"},{"key":"C","text":"Золотыми и красными"},{"key":"D","text":"Чёрными"}], ans:"C", explain:"Bài viết: «Осенью деревья становятся золотыми и красными» (vàng óng và đỏ)."},
      {id:"a1_r15_q6", lang:"vi", prompt:"Từ «мороз» trong tiếng Nga chỉ hiện tượng thời tiết nào?", hint:"Thời tiết mùa đông Nga dưới 0 độ C.", options:[{"key":"A","text":"Nắng gay gắt"},{"key":"B","text":"Băng giá, rét buốt"},{"key":"C","text":"Mưa rào"},{"key":"D","text":"Gió nồm"}], ans:"B", explain:"«Мороз» nghĩa là trời rét căm căm, đóng băng dưới 0 độ C."},
      {id:"a1_r15_q7", lang:"vi", prompt:"Từ «снег» có nghĩa là gì?", hint:"Hiện tượng trắng xóa rơi vào mùa đông.", options:[{"key":"A","text":"Hạt mưa"},{"key":"B","text":"Tuyết"},{"key":"C","text":"Sương mù"},{"key":"D","text":"Mưa đá"}], ans:"B", explain:"«Снег» nghĩa là tuyết."},
      {id:"a1_r15_q8", lang:"vi", prompt:"Từ «дача» của người Nga là địa điểm nào?", hint:"Nơi nghỉ ngơi trồng cây cuối tuần ở ngoại ô.", options:[{"key":"A","text":"Căn hộ cao ốc trung tâm"},{"key":"B","text":"Ngôi nhà vườn nghỉ dưỡng ở ngoại ô"},{"key":"C","text":"Nhà máy"},{"key":"D","text":"Khách sạn"}], ans:"B", explain:"«Дача» là nét văn hóa đặc trưng của Nga: ngôi nhà gỗ có vườn ở ngoại ô để nghỉ mát."},
      {id:"a1_r15_q9", lang:"vi", prompt:"Mùa nào trong năm được xem là 'mùa thu vàng' (Золотая осень) của nước Nga?", hint:"Mùa cây cối chuyển màu lá.", options:[{"key":"A","text":"Зима"},{"key":"B","text":"Весна"},{"key":"C","text":"Лето"},{"key":"D","text":"Осень"}], ans:"D", explain:"«Осень» là mùa thu vàng trứ danh trong thi ca và hội họa Nga."},
      {id:"a1_r15_q10", lang:"vi", prompt:"Để hỏi 'Thời tiết hôm nay thế nào?', ta dùng câu nào?", hint:"Câu hỏi thời tiết thông dụng.", options:[{"key":"A","text":"Какая сегодня погода?"},{"key":"B","text":"Сколько сейчас времени?"},{"key":"C","text":"Куда вы идёте?"},{"key":"D","text":"Кто сегодня дежурный?"}], ans:"A", explain:"«Какая сегодня погода?» nghĩa là 'Hôm nay thời tiết thế nào?'."}
    ]
  },

  // 16. Bạn thích làm gì khi rảnh rỗi?
  {
    id: 16,
    title: "Что вы любите делать в свободное время?",
    topic: "Sở thích cá nhân & Động từ Любить",
    readingTime: "12 phút",
    vocabCategories: [
      {
        title: "Động từ 'Thích' & Hành động (Глагол Любить)",
        rows: [
          {"word": "люби́ть", "gram": "я люблю, ты любишь", "meaning": "yêu, thích", "example": "Я люблю спорт."},
          {"word": "путеше́ствовать", "gram": "động từ", "meaning": "đi du lịch", "example": "Мы любим путешествовать."},
          {"word": "гуля́ть", "gram": "động từ", "meaning": "đi dạo", "example": "Вечером я гуляю в парке."}
        ]
      }
    ],
    readingText: [
      "Свободное время — это прекрасная возможность отдохнуть и заняться любимым делом.",
      "Моя подруга Катя очень любит читать русскую классику. Она прочитала много романов Толстого và thơ của Puskin.",
      "Её брат Олег любит компьютеры. В свободное время он создаёт сайты и изучает программирование.",
      "А я люблю путешествовать. Летом я ездил в Санкт-Петербург. Это великолепный город музеев, мостов и белых ночей.",
      "А как вы проводите своё свободное время? Что вы любите делать?"
    ],
    questions: [
      {id:"a1_r16_q1", lang:"ru", prompt:"Что любит делать Катя в свободное время?", hint:"Xem sở thích của Katya.", options:[{"key":"A","text":"Смотреть футбол"},{"key":"B","text":"Читать русскую классику"},{"key":"C","text":"Готовить еду"},{"key":"D","text":"Танцевать"}], ans:"B", explain:"Bài viết: «Катя очень любит читать русскую классику»."},
      {id:"a1_r16_q2", lang:"ru", prompt:"Чем увлекается Олег?", hint:"Tìm sở thích của Oleg.", options:[{"key":"A","text":"Компьютерами и программированием"},{"key":"B","text":"Плаванием"},{"key":"C","text":"Музыкой"},{"key":"D","text":"Стихами"}], ans:"A", explain:"Bài viết: «Олег любит компьютеры... создаёт сайты и изучает программирование»."},
      {id:"a1_r16_q3", lang:"ru", prompt:"Куда ездил автор летом?", hint:"Tìm thành phố du lịch.", options:[{"key":"A","text":"В Сочи"},{"key":"B","text":"В Санкт-Петербург"},{"key":"C","text":"В Казань"},{"key":"D","text":"В Париж"}], ans:"B", explain:"Bài viết: «Летом я ездил в Санкт-Петербург»."},
      {id:"a1_r16_q4", lang:"ru", prompt:"Как автор описывает Санкт-Петербург?", hint:"Tìm nét nổi tiếng của thành phố này.", options:[{"key":"A","text":"Город музеев, мостов и белых ночей"},{"key":"B","text":"Маленькая деревня"},{"key":"C","text":"Скучное место"},{"key":"D","text":"Жаркий курорт"}], ans:"A", explain:"Bài viết: «Это великолепный город музеев, мостов и белых ночей» (thành phố bảo tàng, những cây cầu và đêm trắng)."},
      {id:"a1_r16_q5", lang:"ru", prompt:"Форма глагола «любить» для местоимения «я»:", hint:"Chia động từ любить ngôi tôi.", options:[{"key":"A","text":"Любишь"},{"key":"B","text":"Любит"},{"key":"C","text":"Люблю"},{"key":"D","text":"Любим"}], ans:"C", explain:"«Я люблю» (có biến âm б $\\rightarrow$ бл)."},
      {id:"a1_r16_q6", lang:"vi", prompt:"Hiện tượng «белые ночи» (đêm trắng) nổi tiếng ở thành phố nào của Nga?", hint:"Thành phố cố đô phương Bắc của Nga.", options:[{"key":"A","text":"Moskva"},{"key":"B","text":"Saint-Petersburg"},{"key":"C","text":"Vladivostok"},{"key":"D","text":"Novosibirsk"}], ans:"B", explain:"Hiện tượng Đêm trắng (mặt trời không lặn hẳn vào mùa hè) trứ danh của Saint-Petersburg."},
      {id:"a1_r16_q7", lang:"vi", prompt:"Động từ «путешествовать» trong tiếng Nga có nghĩa là gì?", hint:"Đi khám phá các vùng đất mới.", options:[{"key":"A","text":"Đi du lịch, đi phượt"},{"key":"B","text":"Đi học"},{"key":"C","text":"Đi siêu thị"},{"key":"D","text":"Đi khám bệnh"}], ans:"A", explain:"«Путешествовать» nghĩa là đi du lịch."},
      {id:"a1_r16_q8", lang:"vi", prompt:"Từ «классика» trong «русская классика» chỉ điều gì?", hint:"Tác phẩm văn học có giá trị mẫu mực vượt thời gian.", options:[{"key":"A","text":"Văn học kinh điển / cổ điển"},{"key":"B","text":"Báo lá cải"},{"key":"C","text":"Truyện tranh"},{"key":"D","text":"Sách bài tập"}], ans:"A", explain:"«Русская классика» là nền văn học kinh điển Nga (Tolstoy, Pushkin, Chekhov...)."},
      {id:"a1_r16_q9", lang:"vi", prompt:"Từ «мост» trong tiếng Nga có nghĩa là gì?", hint:"Công trình bắc qua dòng sông.", options:[{"key":"A","text":"Cái cầu"},{"key":"B","text":"Tàu thủy"},{"key":"C","text":"Bờ đê"},{"key":"D","text":"Bến cảng"}], ans:"A", explain:"«Мост» nghĩa là cây cầu."},
      {id:"a1_r16_q10", lang:"vi", prompt:"Sau động từ «любить», động từ đi kèm phải để ở dạng nào?", hint:"Dạng chia hay dạng nguyên mẫu?", options:[{"key":"A","text":"Dạng nguyên mẫu (Инфинитив)"},{"key":"B","text":"Dạng chia theo ngôi"},{"key":"C","text":"Dạng quá khứ"},{"key":"D","text":"Dạng mệnh lệnh"}], ans:"A", explain:"Sau любить là động từ nguyên mẫu: ví dụ: «Я люблю читать», «Мы любим гулять»."}
    ]
  },

  // 17. Ngày cuối tuần: Đi đâu?
  {
    id: 17,
    title: "Выходной день: Куда мы идём?",
    topic: "Ngày nghỉ cuối tuần & Hướng di chuyển (Куда?)",
    readingTime: "12 phút",
    vocabCategories: [
      {
        title: "Địa điểm giải trí cuối tuần (Куда? - Cách 4)",
        rows: [
          {"word": "в парк", "gram": "Cách 4", "meaning": "đến công viên", "example": "Мы идём в парк."},
          {"word": "в теа́тр", "gram": "Cách 4", "meaning": "đến nhà hát", "example": "Вечером я иду в театр."},
          {"word": "в музе́й", "gram": "Cách 4", "meaning": "đến viện bảo tàng", "example": "В воскресенье мы идём в музей."},
          {"word": "на конце́рт", "gram": "Cách 4", "meaning": "đến buổi hòa nhạc", "example": "Билеты на концерт."}
        ]
      }
    ],
    readingText: [
      "Воскресенье — это мой любимый день. В этот день не нужно рано вставать и идти на занятия.",
      "Вчера мы с друзьями решили, куда мы пойдём в воскресенье.",
      "Сначала утром мы пойдём в Третьяковскую галерею. Там выставлены шедевры великих русских художников.",
      "Днём мы погуляем в парке Горького, поедим мороженое и покатаемся на велосипедах.",
      "А вечером у нас есть билеты в Большой театр на знаменитый балет «Лебединое озеро». Это будет незабываемый день!"
    ],
    questions: [
      {id:"a1_r17_q1", lang:"ru", prompt:"Куда утром пойдут друзья?", hint:"Tìm tên phòng tranh nổi tiếng.", options:[{"key":"A","text":"В зоопарк"},{"key":"B","text":"В Третьяковскую галерею"},{"key":"C","text":"В магазин"},{"key":"D","text":"В кино"}], ans:"B", explain:"Bài viết: «Сначала утром мы пойдём в Третьяковскую галерею»."},
      {id:"a1_r17_q2", lang:"ru", prompt:"В каком парке друзья будут гулять днём?", hint:"Tìm tên công viên.", options:[{"key":"A","text":"В парке Сокольники"},{"key":"B","text":"В парке Горького"},{"key":"C","text":"В парке Коломенское"},{"key":"D","text":"В парке Зарядье"}], ans:"B", explain:"Bài viết: «Днём мы погуляем в парке Горького»."},
      {id:"a1_r17_q3", lang:"ru", prompt:"На какой балет у друзей есть билеты вечером?", hint:"Tên vở ballet huyền thoại của Tchaikovsky.", options:[{"key":"A","text":"«Щелкунчик»"},{"key":"B","text":"«Лебединое озеро»"},{"key":"C","text":"«Спящая красавица»"},{"key":"D","text":"«Жизель»"}], ans:"B", explain:"Bài viết: «на знаменитый балет «Лебединое озеро»» (Hồ Thiên Nga)."},
      {id:"a1_r17_q4", lang:"ru", prompt:"В какой театр пойдут друзья?", hint:"Nhà hát lớn nhất Moscow.", options:[{"key":"A","text":"В Большой театр"},{"key":"B","text":"В Малый театр"},{"key":"C","text":"В МХАТ"},{"key":"D","text":"В Театр сатиры"}], ans:"A", explain:"Bài viết: «билеты в Большой театр»."},
      {id:"a1_r17_q5", lang:"ru", prompt:"Вопрос «Куда?» используется для обозначения:", hint:"Hướng đến đâu.", options:[{"key":"A","text":"Места нахождения (Где?)"},{"key":"B","text":"Направления движения (Đến đâu?)"},{"key":"C","text":"Времени (Когда?)"},{"key":"D","text":"Цены (Сколько?)"}], ans:"B", explain:"«Куда?» có nghĩa là 'Đi đâu / Về đâu?' (chỉ hướng di chuyển)."},
      {id:"a1_r17_q6", lang:"vi", prompt:"Phòng tranh «Третьяковская галерея» nổi tiếng thế giới trưng bày điều gì?", hint:"Tác phẩm nghệ thuật tạo hình.", options:[{"key":"A","text":"Bộ sưu tập tranh của các họa sĩ Nga kiệt xuất"},{"key":"B","text":"Xe ô tô cổ"},{"key":"C","text":"Vũ khí chiến tranh"},{"key":"D","text":"Động vật quý hiếm"}], ans:"A", explain:"Bảo tàng Tretyakov là kho tàng hội họa Nga lớn nhất thế giới."},
      {id:"a1_r17_q7", lang:"vi", prompt:"Vở ballet «Лебединое озеро» có tên tiếng Việt là gì?", hint:"Tác phẩm của nhà soạn nhạc Tchaikovsky.", options:[{"key":"A","text":"Kẹp Hạt Dẻ"},{"key":"B","text":"Hồ Thiên Nga"},{"key":"C","text":"Người Đẹp Ngủ Trong Rừng"},{"key":"D","text":"Romeo và Juliet"}], ans:"B", explain:"«Лебединое озеро» nghĩa là 'Hồ Thiên Nga'."},
      {id:"a1_r17_q8", lang:"vi", prompt:"Nhà hát «Большой театр» nằm ở thành phố nào của Nga?", hint:"Nhà hát Bolshoi ở thủ đô.", options:[{"key":"A","text":"Saint-Petersburg"},{"key":"B","text":"Moskva"},{"key":"C","text":"Kazan"},{"key":"D","text":"Sochi"}], ans:"B", explain:"Nhà hát Bolshoi là biểu tượng văn hóa rực rỡ nằm tại thủ đô Moskva."},
      {id:"a1_r17_q9", lang:"vi", prompt:"Từ «шедевр» trong bài đọc mang ý nghĩa gì?", hint:"Tác phẩm đạt đỉnh cao nghệ thuật.", options:[{"key":"A","text":"Bức vẽ thử"},{"key":"B","text":"Kiệt tác, tuyệt tác"},{"key":"C","text":"Bản sao chép"},{"key":"D","text":"Tranh biếm họa"}], ans:"B", explain:"«Шедевр» nghĩa là kiệt tác nghệ thuật."},
      {id:"a1_r17_q10", lang:"vi", prompt:"Từ «велосипед» là phương tiện gì?", hint:"Phương tiện hai bánh dùng sức đạp.", options:[{"key":"A","text":"Xe máy"},{"key":"B","text":"Xe đạp"},{"key":"C","text":"Tàu điện ngầm"},{"key":"D","text":"Xe buýt"}], ans:"B", explain:"«Велосипед» là xe đạp (кататься на велосипеде: đi xe đạp)."}
    ]
  },

  // 18. Viết thư cho bạn: Tôi sống và học tập ra sao
  {
    id: 18,
    title: "Письмо другу: Как я живу и учусь",
    topic: "Kỹ năng viết thư & Tường thuật cuộc sống A1",
    readingTime: "12 phút",
    vocabCategories: [
      {
        title: "Từ vựng viết thư (Структура письма)",
        rows: [
          {"word": "Дорого́й друг", "gram": "lời xưng hô", "meaning": "Bạn thân mến", "example": "Дорогой друг Иван!"},
          {"word": "Напиши́ мне", "gram": "mệnh lệnh", "meaning": "Hãy viết thư cho tôi", "example": "Напиши мне ответ."},
          {"word": "Жду отве́та", "gram": "cụm từ", "meaning": "Tôi đợi thư trả lời", "example": "С нетерпением жду ответа."}
        ]
      }
    ],
    readingText: [
      "Здравствуй, дорогой друг Алексей!",
      "Большое спасибо за твоё письмо. Я был очень рад получить от тебя весточку.",
      "Ты спрашиваешь, как я живу в Москве. У меня всё замечательно! Я уже привык к жизни здесь.",
      "В университете у нас отличные преподаватели. Мы учим грамматику, читаем тексты и каждый день говорим по-русски.",
      "В свободное время я часто хожу на выставки или гуляю по старым улицам города. Москва — прекрасна!",
      "Как твои дела? Как твоя учёба? Напиши мне обо всём!\nКрепко обнимаю,\nТвой друг Минх."
    ],
    questions: [
      {id:"a1_r18_q1", lang:"ru", prompt:"Кому адресовано это письмо?", hint:"Xem tên người nhận ở đầu thư.", options:[{"key":"A","text":"Антону"},{"key":"B","text":"Алексею"},{"key":"C","text":"Виктору"},{"key":"D","text":"Борису"}], ans:"B", explain:"Bức thư mở đầu: «Здравствуй, дорогой друг Алексей!»."},
      {id:"a1_r18_q2", lang:"ru", prompt:"Кто написал это письмо?", hint:"Xem chữ ký ở cuối thư.", options:[{"key":"A","text":"Минх"},{"key":"B","text":"Иван"},{"key":"C","text":"Дэвид"},{"key":"D","text":"Сергей"}], ans:"A", explain:"Cuối thư ký tên: «Твой друг Минх»."},
      {id:"a1_r18_q3", lang:"ru", prompt:"Где сейчас живёт и учится Минх?", hint:"Tìm địa điểm sống của Minh.", options:[{"key":"A","text":"В Лондоне"},{"key":"B","text":"В Москве"},{"key":"C","text":"В Ханое"},{"key":"D","text":"В Париже"}], ans:"B", explain:"Bài viết: «Ты спрашиваешь, как я живу в Москве»."},
      {id:"a1_r18_q4", lang:"ru", prompt:"Куда Минх часто ходит в свободное время?", hint:"Xem hoạt động rảnh rỗi.", options:[{"key":"A","text":"В клубы"},{"key":"B","text":"На выставки и гуляет по старым улицам"},{"key":"C","text":"На стадион"},{"key":"D","text":"В больницу"}], ans:"B", explain:"Bài viết: «я часто хожу на выставки или гуляю по старым улицам города»."},
      {id:"a1_r18_q5", lang:"ru", prompt:"Какая фраза часто используется в конце дружеского письма?", hint:"Lời tạm biệt ấm áp cuối thư.", options:[{"key":"A","text":"Здравствуйте!"},{"key":"B","text":"Крепко обнимаю / Твой друг"},{"key":"C","text":"Сколько это стоит?"},{"key":"D","text":"Приятного аппетита!"}], ans:"B", explain:"«Крепко обнимаю / Твой друг» là lời chào kết thư thân thiết."},
      {id:"a1_r18_q6", lang:"vi", prompt:"Từ «весточка» trong câu «получить весточку» có nghĩa là gì?", hint:"Từ thân mật chỉ tin tức/tin nhắn.", options:[{"key":"A","text":"Món quà lớn"},{"key":"B","text":"Tin tức, lá thư ngắn từ ai đó"},{"key":"C","text":"Lời trách móc"},{"key":"D","text":"Bản hợp đồng"}], ans:"B", explain:"«Весточка» nghĩa là mẩu tin tức, tin nhắn tốt lành từ bạn bè."},
      {id:"a1_r18_q7", lang:"vi", prompt:"Từ «привыкнуть» trong câu «Я уже привык» có nghĩa là gì?", hint:"Trạng thái sau một thời gian thích nghi.", options:[{"key":"A","text":"Đã quen với cuộc sống ở đây"},{"key":"B","text":"Vẫn còn bỡ ngỡ"},{"key":"C","text":"Muốn quay về"},{"key":"D","text":"Bị ốm"}], ans:"A", explain:"«Я привык» nghĩa là 'Tôi đã quen rồi'."},
      {id:"a1_r18_q8", lang:"vi", prompt:"Từ «выставка» trong câu «хожу на выставки» có nghĩa là gì?", hint:"Nơi trưng bày các tác phẩm hoặc hiện vật.", options:[{"key":"A","text":"Triển lãm"},{"key":"B","text":"Hội chợ việc làm"},{"key":"C","text":"Buổi hòa nhạc"},{"key":"D","text":"Rạp xiếc"}], ans:"A", explain:"«Выставка» nghĩa là triển lãm."},
      {id:"a1_r18_q9", lang:"vi", prompt:"Cụm từ «Крепко обнимаю» ở cuối thư thể hiện điều gì?", hint:"Tình cảm thân thiết khi ôm bạn.", options:[{"key":"A","text":"Trang trọng xa cách"},{"key":"B","text":"Cái ôm nồng ấm, thân thương"},{"key":"C","text":"Lời xin lỗi"},{"key":"D","text":"Lời hẹn giờ gặp"}], ans:"B", explain:"«Крепко обнимаю» (Ôm bạn thật chặt) là cách chào kết thư bạn bè rất phổ biến ở Nga."},
      {id:"a1_r18_q10", lang:"vi", prompt:"Cấu trúc mở đầu thư thân mật cho bạn nam là gì?", hint:"Tính từ дорогой đi với giống đực.", options:[{"key":"A","text":"Дорогой друг..."},{"key":"B","text":"Дорогая подруга..."},{"key":"C","text":"Уважаемый господин..."},{"key":"D","text":"Дорогие друзья..."}], ans:"A", explain:"«Дорогой друг + [Tên bạn nam]» là lời chào chuẩn thân mật trong thư tiếng Nga."}
    ]
  },

  // 19. Các ngày lễ ở Nga
  {
    id: 19,
    title: "Праздники в России: Новый год и традиции",
    topic: "Văn hóa, Lễ hội & Lời chúc mừng (Поздравляю!)",
    readingTime: "12 phút",
    vocabCategories: [
      {
        title: "Ngày lễ & Lời chúc mừng (Праздники)",
        rows: [
          {"word": "пра́здник", "gram": "danh từ giống đực", "meaning": "ngày lễ, ngày hội", "example": "Сегодня большой праздник."},
          {"word": "Но́вый год", "gram": "danh từ", "meaning": "Năm Mới", "example": "С Новым годом!"},
          {"word": "пода́рок", "gram": "danh từ giống đực", "meaning": "món quà", "example": "Дед Мороз дарит подарки."},
          {"word": "поздравля́ть", "gram": "động từ", "meaning": "chúc mừng", "example": "Поздравляю с праздником!"}
        ]
      }
    ],
    readingText: [
      "Самый любимый và kỳ diệu nhất trong năm ở Nga chính là ngày Tết Năm Mới (Новый год).",
      "В конце декабря на улицах и площадях зажигаются праздничные огни. В каждом доме наряжают пушистую ёлку.",
      "Главные новогодние персонажи — это добрый Дед Мороз и его внучка Снегурочка. Они дарят детям подарки.",
      "В полночь тридцать первого декабря бьют куранты на Спасской башне Кремля. Все поднимают бокалы и говорят: «С Новым годом! С новым счастьем!»",
      "Люди желают друг другу здоровья, любви, радости и успехов в новом году."
    ],
    questions: [
      {id:"a1_r19_q1", lang:"ru", prompt:"Какой праздник в России самый любимый?", hint:"Đọc câu đầu tiên.", options:[{"key":"A","text":"Пасха"},{"key":"B","text":"Новый год"},{"key":"C","text":"День труда"},{"key":"D","text":"День знаний"}], ans:"B", explain:"Bài viết: «Самый любимый праздник в России — это Новый год»."},
      {id:"a1_r19_q2", lang:"ru", prompt:"Что наряжают в каждом доме перед праздником?", hint:"Cây thông Noel trang trí trong nhà.", options:[{"key":"A","text":"Берёзу"},{"key":"B","text":"Пушистую ёлку"},{"key":"C","text":"Цветы"},{"key":"D","text":"Окно"}], ans:"B", explain:"Bài viết: «В каждом доме наряжают пушистую ёлку» (cây thông Noel)."},
      {id:"a1_r19_q3", lang:"ru", prompt:"Кто дарит детям подарки на Новый год в России?", hint:"Hai nhân vật huyền thoại mùa đông Nga.", options:[{"key":"A","text":"Санта-Клаус"},{"key":"B","text":"Дед Мороз и Снегурочка"},{"key":"C","text":"Учитель"},{"key":"D","text":"Соседи"}], ans:"B", explain:"Bài viết: «Дед Мороз и его внучка Снегурочка»."},
      {id:"a1_r19_q4", lang:"ru", prompt:"Где бьют главные куранты страны в полночь?", hint:"Chiếc đồng hồ nổi tiếng trên tháp Điện Kremlin.", options:[{"key":"A","text":"На Спасской башне Кремля"},{"key":"B","text":"На вокзале"},{"key":"C","text":"В университете"},{"key":"D","text":"На стадионе"}], ans:"A", explain:"Bài viết: «бьют куранты на Спасской башне Кремля»."},
      {id:"a1_r19_q5", lang:"ru", prompt:"Какая традиционная фраза звучит в момент наступления Нового года?", hint:"Lời chúc đón năm mới quen thuộc.", options:[{"key":"A","text":"«С Новым годом! С новым счастьем!»"},{"key":"B","text":"«Спокойной ночи!»"},{"key":"C","text":"«До свидания!»"},{"key":"D","text":"«Приятного аппетита!»"}], ans:"A", explain:"Người Nga cùng chúc tụng: «С Новым годом! С новым счастьем!»."},
      {id:"a1_r19_q6", lang:"vi", prompt:"Nhân vật «Дед Мороз» trong văn hóa Nga là ai?", hint:"Ông già mang quà đến cho trẻ em đêm giao thừa.", options:[{"key":"A","text":"Thần Mùa Xuân"},{"key":"B","text":"Ông Già Tuyết (Băng Giá)"},{"key":"C","text":"Vị thần rừng"},{"key":"D","text":"Chú hề"}], ans:"B", explain:"«Дед Мороз» là Ông Già Tuyết (ông già Băng giá) biểu tượng của năm mới Nga."},
      {id:"a1_r19_q7", lang:"vi", prompt:"Nhân vật «Снегурочка» là ai đối với Ded Moroz?", hint:"Xem từ: внучка Снегурочка.", options:[{"key":"A","text":"Con gái"},{"key":"B","text":"Cháu gái (Cô bé Tuyết)"},{"key":"C","text":"Người vợ"},{"key":"D","text":"Chị gái"}], ans:"B", explain:"«Внучка Снегурочка» là cháu gái của Ông Già Tuyết (Cô bé Tuyết)."},
      {id:"a1_r19_q8", lang:"vi", prompt:"Từ «ёлка» trong dịp năm mới chỉ đồ vật gì?", hint:"Cây xanh lá kim trang trí đèn sao.", options:[{"key":"A","text":"Cây đào"},{"key":"B","text":"Cây thông Noel"},{"key":"C","text":"Cây mai"},{"key":"D","text":"Cây quất"}], ans:"B", explain:"«Ёлка» là cây thông mừng năm mới."},
      {id:"a1_r19_q9", lang:"vi", prompt:"Chuông đồng hồ Kremlin («куранты») điểm bao nhiêu tiếng vào lúc giao thừa?", hint:"Đồng hồ điểm đúng 12 giờ đêm.", options:[{"key":"A","text":"6 tiếng"},{"key":"B","text":"10 tiếng"},{"key":"C","text":"12 tiếng chuông"},{"key":"D","text":"24 tiếng"}], ans:"C", explain:"Đồng hồ Spasskaya ngân vang 12 tiếng chuông chào đón thời khắc chuyển giao năm mới."},
      {id:"a1_r19_q10", lang:"vi", prompt:"Để chúc mừng ai đó nhân dịp ngày lễ, ta dùng cấu trúc nào?", hint:"Động từ chúc mừng: Поздравлять с...", options:[{"key":"A","text":"Поздравляю с праздником!"},{"key":"B","text":"Спасибо за праздник!"},{"key":"C","text":"Извините за праздник!"},{"key":"D","text":"Где праздник?"}], ans:"A", explain:"«Поздравляю с...» là cấu trúc chúc mừng chuẩn trong tiếng Nga."}
    ]
  },

  // 20. Bài kiểm tra tổng hợp A1: Я говорю по-русски!
  {
    id: 20,
    title: "Итоговый тест A1: Я говорю по-русски!",
    topic: "Tổng ôn toàn diện kiến thức & Bứt phá năng lực TRKI A1",
    readingTime: "15 phút",
    vocabCategories: [
      {
        title: "Tổng hợp từ vựng chìa khóa A1 (Ключевые слова)",
        rows: [
          {"word": "студе́нт / студе́нтка", "gram": "danh từ", "meaning": "sinh viên nam / sinh viên nữ", "example": "Мы студенты."},
          {"word": "учи́ться", "gram": "động từ", "meaning": "học tập", "example": "Они учатся в университете."},
          {"word": "знать", "gram": "động từ", "meaning": "biết", "example": "Я знаю русский язык."},
          {"word": "свобо́дно", "gram": "phó từ", "meaning": "lưu loát, tự do", "example": "Он свободно говорит."}
        ]
      }
    ],
    readingText: [
      "Здравствуйте! Меня зовут Антон. Я хочу подвести итоги моего первого года изучения русского языка.",
      "Когда я только приехал в Россию, я знал только слова «Привет» và «Спасибо». Всё вокруг казалось новым и сложным.",
      "Сейчас я уже хорошо понимаю преподавателей на уроках, свободно ориентируюсь в городе và могу легко заказать еду в кафе.",
      "У меня много друзей из разных стран: из Вьетнама, Китая, Франции и России. Мы вместе готовим обед, гуляем по Москве и поём песни под гитару.",
      "Русский язык открыл для меня целый новый мир культуры, литературы и дружбы. Я с гордостью могу сказать: «Я говорю по-русски!»"
    ],
    questions: [
      {id:"a1_r20_q1", lang:"ru", prompt:"Что знал Антон, когда только приехал в Россию?", hint:"Xem câu 2.", options:[{"key":"A","text":"Всю грамматику"},{"key":"B","text":"Только слова «Привет» и «Спасибо»"},{"key":"C","text":"Много стихов"},{"key":"D","text":"Ничего не знал"}], ans:"B", explain:"Bài viết: «Когда я только приехал в Россию, я знал только слова «Привет» và «Спасибо»»."},
      {id:"a1_r20_q2", lang:"ru", prompt:"Что Антон умеет делать сейчас?", hint:"Xem sự tiến bộ của Anton ở đoạn 3.", options:[{"key":"A","text":"Только читать алфавит"},{"key":"B","text":"Понимает учителей, ориентируется в городе и может заказать еду"},{"key":"C","text":"Пишет научные книги"},{"key":"D","text":"Не говорит по-русски"}], ans:"B", explain:"Đoạn 3 liệt kê: «хорошо понимаю преподавателей... ориентируюсь в городе... могу легко заказать еду»."},
      {id:"a1_r20_q3", lang:"ru", prompt:"Из каких стран друзья Антона?", hint:"Xem danh sách quốc gia bạn bè.", options:[{"key":"A","text":"Только из Англии"},{"key":"B","text":"Из Вьетнама, Китая, Франции и России"},{"key":"C","text":"Из Америки"},{"key":"D","text":"Из Африки"}], ans:"B", explain:"Bài viết: «из разных стран: из Вьетнама, Китая, Франции и России»."},
      {id:"a1_r20_q4", lang:"ru", prompt:"Что открыл русский язык для Антона?", hint:"Đọc đoạn cuối cùng.", options:[{"key":"A","text":"Бизнес"},{"key":"B","text":"Новый мир культуры, литературы и дружбы"},{"key":"C","text":"Новую работу на заводе"},{"key":"D","text":"Секреты математики"}], ans:"B", explain:"Bài viết: «Русский язык открыл для меня целый новый мир культуры, литературы и дружбы»."},
      {id:"a1_r20_q5", lang:"ru", prompt:"Какую гордую фразу может сказать студент в конце курса A1?", hint:"Xem câu kết của bài.", options:[{"key":"A","text":"«Я не знаю русский язык»"},{"key":"B","text":"«Я говорю по-русски!»"},{"key":"C","text":"«Русский язык слишком лёгкий»"},{"key":"D","text":"«До свидания навсегда»"}], ans:"B", explain:"Câu kết: «Я с гордостью могу сказать: «Я говорю по-русски!»»."},
      {id:"a1_r20_q6", lang:"vi", prompt:"Cụm từ «подвести итоги» trong tiếng Nga có nghĩa là gì?", hint:"Hoạt động nhìn nhận lại kết quả sau một quá trình.", options:[{"key":"A","text":"Lên kế hoạch mới"},{"key":"B","text":"Tổng kết, điểm lại kết quả"},{"key":"C","text":"Bắt đầu làm bài"},{"key":"D","text":"Bỏ cuộc giữa chừng"}], ans:"B", explain:"«Подвести итоги» nghĩa là tổng kết, đúc kết kết quả."},
      {id:"a1_r20_q7", lang:"vi", prompt:"Từ «свободно» trong câu «свободно ориентируюсь в городе» nghĩa là gì?", hint:"Khả năng định hướng dễ dàng, tự tin.", options:[{"key":"A","text":"Một cách dễ dàng, tự tin, không bị lạc"},{"key":"B","text":"Miễn phí vé xe"},{"key":"C","text":"Chậm chạp"},{"key":"D","text":"Đi lung tung"}], ans:"A", explain:"«Свободно ориентироваться» nghĩa là tự tin định hướng đường xá, không gặp trở ngại."},
      {id:"a1_r20_q8", lang:"vi", prompt:"Cụm từ «петь песни под гитару» chỉ hoạt động gì?", hint:"Hoạt động văn nghệ sinh viên.", options:[{"key":"A","text":"Hát theo tiếng đệm đàn guitar"},{"key":"B","text":"Mua đàn guitar"},{"key":"C","text":"Dạy đàn guitar"},{"key":"D","text":"Nghe đài radio"}], ans:"A", explain:"«Петь под гитару» là hát trên nền đệm đàn guitar."},
      {id:"a1_r20_q9", lang:"vi", prompt:"Từ «с гордостью» thể hiện cảm xúc gì của người nói?", hint:"Cảm giác tự hào về nỗ lực bản thân.", options:[{"key":"A","text":"Với nỗi buồn bã"},{"key":"B","text":"Với niềm tự hào hãnh diện"},{"key":"C","text":"Với sự sợ hãi"},{"key":"D","text":"Với vẻ thờ ơ"}], ans:"B", explain:"«С гордостью» nghĩa là với niềm tự hào, hãnh diện."},
      {id:"a1_r20_q10", lang:"vi", prompt:"Trình độ A1 (Sơ cấp - ТРКИ Элементарный) chứng nhận người học có thể làm được điều gì?", hint:"Mục tiêu của cấp độ A1.", options:[{"key":"A","text":"Chỉ biết nói xin chào và đếm từ 1 đến 10"},{"key":"B","text":"Giao tiếp cơ bản về bản thân, gia đình, mua sắm, học tập và nơi ở"},{"key":"C","text":"Viết luận văn thạc sĩ"},{"key":"D","text":"Dịch tiểu thuyết chuyên ngành"}], ans:"B", explain:"Chuẩn TRKI A1 chứng nhận khả năng giao tiếp sơ cấp tự tin trong các tình huống thiết yếu hàng ngày."}
    ]
  }
];

// Gộp toàn bộ 20 bài
const all20Lessons = [
  ...lessonsData,
  ...remainingLessons,
  ...lessons11to20
];

console.log(`Đang tạo ${all20Lessons.length} bài học A1...`);

const manifestList = [];

all20Lessons.forEach(lesson => {
  const fileName = `lesson-${String(lesson.id).padStart(2, '0')}.json`;
  const filePath = path.join(targetDir, fileName);

  const fullLessonJson = {
    id: lesson.id,
    level: "A1",
    skill: "reading",
    title: lesson.title,
    badge: "A1 Đọc hiểu",
    readingTime: lesson.readingTime || "12 phút",
    source: "Giáo trình tiếng Nga A1 Chuẩn Quốc Tế (TRKI)",
    topic: lesson.topic,
    vocabCategories: lesson.vocabCategories || [],
    readingText: lesson.readingText || [],
    questions: lesson.questions || []
  };

  fs.writeFileSync(filePath, JSON.stringify(fullLessonJson, null, 2), 'utf8');

  manifestList.push({
    id: lesson.id,
    file: fileName,
    title: lesson.title,
    badge: "A1 Đọc hiểu",
    readingTime: lesson.readingTime || "12 phút",
    source: "Giáo trình tiếng Nga A1 Chuẩn Quốc Tế (TRKI)",
    topic: lesson.topic,
    questionCount: (lesson.questions || []).length
  });
});

const manifestPath = path.join(targetDir, 'manifest.json');
fs.writeFileSync(manifestPath, JSON.stringify(manifestList, null, 2), 'utf8');

console.log(`✓ Đã tạo thành công 20 bài học A1 và cập nhật manifest.json tại ${targetDir}`);

