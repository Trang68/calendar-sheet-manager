const fs = require('fs');
const path = require('path');

// ============================================================================
// LESSON 1 DATA (20 câu: 10 Nga + 10 Việt; 10 Từ vựng + 10 Đọc hiểu/Ngữ pháp)
// ============================================================================
const lesson1 = {
  id: 1,
  title: "Семья московской писательницы Татьяны Устиновой",
  badge: "BÀI ĐỌC 1",
  time: "25-30 phút",
  source: "Газета «Труд»",
  topic: "Văn học & Gia đình",
  paragraphs: [
    "Недавно в газете «Труд» была напечатана статья, в которой рассказывалось о семье московской писательницы Татьяны Устиновой. Вся её дружная семья живёт за городом в большом старом доме. На кухне постоянно готовится еда, дети играют, а в саду весело бегает собака Кузя. В этой семье все друг друга любят и всегда помогают друг другу. Романы создаются здесь же, в этом уютном доме, и печатаются на старом компьютере, который писательница купила когда-то очень давно у знакомых.",
    "Этот дом, сад и эта семья, где все кричат, ссорятся и мирятся, появляются в каждом романе Татьяны Устиновой. Её герои — нормальные, живые люди, с которыми мы встречаемся каждый день. Поэтому книги Татьяны Устиновой читаются очень легко и с удовольствием."
  ],
  caption: "📷 <strong>Chú thích ảnh:</strong> Т. Устинова и её загородный дом.",
  vocabBoxTitle: "Từ vựng & Cấu trúc ngữ pháp trọng tâm (Bài 1)",
  vocabCategories: [
    {
      title: "Động từ & Dạng phản thân bị động (-ся)",
      badge: "5 Động từ quan trọng",
      badgeClass: "cat-verb",
      headers: ["Động từ", "Ý nghĩa tiếng Việt", "Ví dụ trong bài"],
      rows: [
        { word: "ссо́риться", gram: "НСВ (с кем?)", meaning: "Cãi nhau, xung đột", example: "...где все кричат, <strong>ссорятся</strong>..." },
        { word: "мири́ться", gram: "НСВ (с кем?)", meaning: "Làm hòa, giảng hòa", example: "...ссорятся và <strong>мирятся</strong>..." },
        { word: "печа́таться", gram: "НСВ (bị động -ся)", meaning: "Được in ấn, xuất bản", example: "...романы <strong>печатаются</strong> на компьютере" },
        { word: "создава́ться", gram: "НСВ (bị động -ся)", meaning: "Được sáng tác, tạo ra", example: "Романы <strong>создаются</strong> здесь же..." },
        { word: "чита́ться", gram: "НСВ (bị động)", meaning: "Được đón đọc (dễ đọc)", example: "...книги <strong>читаются</strong> очень легко" }
      ]
    },
    {
      title: "Danh từ & Tính từ",
      badge: "Danh - Tính từ",
      badgeClass: "cat-noun",
      headers: ["Từ vựng", "Ý nghĩa tiếng Việt", "Ví dụ trong bài"],
      rows: [
        { word: "писа́тельница", gram: "Danh từ (Giống cái)", meaning: "Nữ nhà văn", example: "...о семье московской <strong>писательницы</strong>..." },
        { word: "рома́н", gram: "Danh từ (Giống đực)", meaning: "Tiểu thuyết", example: "<strong>Романы</strong> создаются здесь же..." },
        { word: "дру́жная (семья)", gram: "Tính từ (Giống cái)", meaning: "Hòa thuận, gắn kết", example: "Вся её <strong>druzhная семья</strong> живёт..." },
        { word: "ую́тный (дом)", gram: "Tính từ (Giống đực)", meaning: "Ấm cúng, thoải mái", example: "...в этом <strong>уютном доме</strong>..." },
        { word: "живы́е (люди)", gram: "Tính từ (Số nhiều)", meaning: "Chân thực, sống động", example: "Её герои — нормальные, <strong>живые люди</strong>..." }
      ]
    },
    {
      title: "Cụm từ cố định & Quản trị cách",
      badge: "Cụm từ hay gặp",
      badgeClass: "cat-expr",
      headers: ["Cụm từ tiếng Nga", "Nghĩa tiếng Việt", "Ngữ pháp"],
      rows: [
        { word: "за́ городом", gram: "Giới từ за + Cách 5", meaning: "Ở ngoại ô, ngoài thành phố" },
        { word: "друг дру́га / друг дру́гу", gram: "Cách 4 (любить) / Cách 3 (помогать)", meaning: "Lẫn nhau, nhau" },
        { word: "с удово́льствием", gram: "Giới từ с + Cách 5", meaning: "Một cách thích thú, hào hứng" },
        { word: "когда́-то о́чень давно́", gram: "Hậu tố bất định -то", meaning: "Từ một thời rất xa xưa" }
      ]
    }
  ],
  questions: [
    // 1-10: CÂU HỎI TIẾNG NGA (РУССКИЙ ЯЗЫК)
    {
      id: "l1_q1",
      lang: "ru",
      prompt: "Что означает глагол <strong>«ссориться»</strong> в контексте текста?",
      options: [
        { key: "A", text: "A. Жить в полном согласии и никогда не спорить" },
        { key: "B", text: "B. Конфликтовать, спорить и выражать разногласия", correct: true },
        { key: "C", text: "C. Вместе готовить вкусную еду на кухне" },
        { key: "D", text: "D. Восстанавливать мирные дружеские отношения" }
      ],
      ans: "B",
      expl: {
        B: "Правильно! «Ссориться» означает спорить, конфликтовать và выражать взаимные претензии.",
        other: "Неверно. «Ссориться» — это вступать в ссору, конфликтовать."
      }
    },
    {
      id: "l1_q2",
      lang: "ru",
      prompt: "Какой глагол является антонимом (противоположным по значению) к слову <strong>«ссориться»</strong>?",
      options: [
        { key: "A", text: "A. кричать" },
        { key: "B", text: "B. мириться", correct: true },
        { key: "C", text: "C. печатать" },
        { key: "D", text: "D. играть" }
      ],
      ans: "B",
      expl: {
        B: "Правильно! Антонимом к глаголу «ссориться» является «мириться» (восстанавливать согласие).",
        other: "Неверно. Противоположным по смыслу является глагол «мириться»."
      }
    },
    {
      id: "l1_q3",
      lang: "ru",
      prompt: "Какое значение имеет прилагательное <strong>«уютный»</strong> в словосочетании «уютный дом»?",
      options: [
        { key: "A", text: "A. Огромный и очень дорогой дворец" },
        { key: "B", text: "B. Пустой дом, где никто не живёт" },
        { key: "C", text: "C. Тёплый, удобный, располагающий к спокойному отдыху", correct: true },
        { key: "D", text: "D. Старый и разрушенный дом без отопления" }
      ],
      ans: "C",
      expl: {
        C: "Правильно! «Уютный» — удобный, благоустроенный, приятный для жизни.",
        other: "Неверно. «Уютный дом» — это тёплый, приятный и удобный для жизни дом."
      }
    },
    {
      id: "l1_q4",
      lang: "ru",
      prompt: "Как вы понимаете выражение <strong>«дружная семья»</strong>?",
      options: [
        { key: "A", text: "A. Семья, члены которой живут в согласии, любят и поддерживают друг друга", correct: true },
        { key: "B", text: "B. Семья, в которой все люди работают в одной компании" },
        { key: "C", text: "C. Семья, у которой очень много богатых друзей" },
        { key: "D", text: "D. Семья, которая часто устраивает шумные праздники для соседей" }
      ],
      ans: "A",
      expl: {
        A: "Правильно! Дружная семья — это сплочённая семья, где царят любовь, взаимопомощь и согласие.",
        other: "Неверно. Дружная семья — это семья, сплочённая взаимной любовью и поддержкой."
      }
    },
    {
      id: "l1_q5",
      lang: "ru",
      prompt: "Вставьте правильную форму местоимений: «В этой семье все любят ________ и помогают ________.»",
      hint: "Помните управление: любить (кого? Винительный пад.) — помогать (кому? Дательный пад.)",
      options: [
        { key: "A", text: "A. друг другу / друг друга" },
        { key: "B", text: "B. друг друга / друг другу", correct: true },
        { key: "C", text: "C. друг с другом / друг о друге" },
        { key: "D", text: "D. друг от друга / друг к другу" }
      ],
      ans: "B",
      expl: {
        B: "Правильно! Любить кого? -> друг друга (В.п.); Помогать кому? -> друг другу (Д.п.).",
        other: "Ошибка в падеже: любить кого? (друг друга), помогать кому? (друг другу)."
      }
    },
    {
      id: "l1_q6",
      lang: "ru",
      prompt: "Выберите правильную форму пассивного глагола: «Романы писательницы ________ здесь же, в загородном доме.»",
      options: [
        { key: "A", text: "A. создаются", correct: true },
        { key: "B", text: "B. создаётся" },
        { key: "C", text: "C. создают" },
        { key: "D", text: "D. создал" }
      ],
      ans: "A",
      expl: {
        A: "Правильно! Подлежащее «романы» во множественном числе, поэтому сказуемое — «создаются».",
        other: "Неверно. Подлежащее «романы» (мн.ч.) требует форму глагола «создаются»."
      }
    },
    {
      id: "l1_q7",
      lang: "ru",
      prompt: "Вставьте союзное слово: «Её герои — живые люди, с ________ мы встречаемся каждый день.»",
      hint: "Конструкция: с кем? (Творительный падеж множественного числа)",
      options: [
        { key: "A", text: "A. которые" },
        { key: "B", text: "B. которых" },
        { key: "C", text: "C. которыми", correct: true },
        { key: "D", text: "D. которым" }
      ],
      ans: "C",
      expl: {
        C: "Правильно! Предлог «с» с одушевлёнными существительными требует Творительного падежа: с кем? -> с которыми.",
        other: "Ошибка в падеже. С предлогом «с» (с кем?) используется форма Т.п. мн.ч. «с которыми»."
      }
    },
    {
      id: "l1_q8",
      lang: "ru",
      prompt: "Где живёт семья Татьяны Устиновой согласно тексту?",
      options: [
        { key: "A", text: "A. В современной квартире в центре Москвы" },
        { key: "B", text: "B. За городом в большом старом доме", correct: true },
        { key: "C", text: "C. В гостинице в Санкт-Петербурге" },
        { key: "D", text: "D. На даче у своих знакомых" }
      ],
      ans: "B",
      expl: {
        B: "Правильно! В тексте прямо сказано: «Вся её дружная семья живёт за городом в большом старом доме».",
        other: "Неверно. По тексту семья живёт за городом в большом старом доме."
      }
    },
    {
      id: "l1_q9",
      lang: "ru",
      prompt: "На чём писательница печатает свои романы?",
      options: [
        { key: "A", text: "A. На новейшем дорогом ноутбуке" },
        { key: "B", text: "B. На старой механической пишущей машинке" },
        { key: "C", text: "C. На старом компьютере, купленном когда-то у знакомых", correct: true },
        { key: "D", text: "D. Она пишет все романы от руки в тетрадях" }
      ],
      ans: "C",
      expl: {
        C: "Правильно! Текст: «...и печатаются на старом компьютере, который писательница купила когда-то очень давно у знакомых».",
        other: "Неверно. Романы печатаются на старом компьютере, купленном у знакомых."
      }
    },
    {
      id: "l1_q10",
      lang: "ru",
      prompt: "Почему книги Татьяны Устиновой читаются очень легко и с удовольствием?",
      options: [
        { key: "A", text: "A. Потому что в них мало страниц и много цветных картинок" },
        { key: "B", text: "B. Потому что её герои — нормальные, живые люди, похожие на нас", correct: true },
        { key: "C", text: "C. Потому что книги продаются по очень низкой цене" },
        { key: "D", text: "D. Потому что сюжеты рассказывают только о жизни за городом" }
      ],
      ans: "B",
      expl: {
        B: "Правильно! В тексте: «Её герои — нормальные, живые люди, с которыми мы встречаемся каждый день. Поэтому книги... читаются очень легко».",
        other: "Неверно. Причина в том, что персонажи — настоящие, живые люди из реальной жизни."
      }
    },

    // 11-20: CÂU HỎI TIẾNG VIỆT
    {
      id: "l1_q11",
      lang: "vi",
      prompt: "Ý nghĩa của cụm từ <strong>«за городом»</strong> trong câu «Вся её дружная семья живёт за городом» là gì?",
      options: [
        { key: "A", text: "A. Trong trung tâm thủ đô sầm uất" },
        { key: "B", text: "B. Ở vùng ngoại ô, ngoài thành phố", correct: true },
        { key: "C", text: "C. Ở nước ngoài" },
        { key: "D", text: "D. Dưới một khu phố ngầm" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! «За городом» nghĩa là ở vùng ngoại ô, ngoài phạm vi thành phố.",
        other: "Chưa đúng. «За городом» là cụm từ chỉ vị trí ở ngoại ô, ngoại thành."
      }
    },
    {
      id: "l1_q12",
      lang: "vi",
      prompt: "Cụm thành ngữ cố định <strong>«с удовольствием»</strong> thể hiện thái độ hay tâm trạng gì?",
      options: [
        { key: "A", text: "A. Một cách miễn cưỡng, bắt buộc" },
        { key: "B", text: "B. Một cách thờ ơ, không quan tâm" },
        { key: "C", text: "C. Một cách hào hứng, say mê và thích thú", correct: true },
        { key: "D", text: "D. Một cách vội vã, lo âu" }
      ],
      ans: "C",
      expl: {
        C: "Chính xác! «С удовольствием» mang nghĩa là với niềm vui sướng, say mê, thích thú.",
        other: "Chưa đúng. «С удовольствием» diễn tả việc làm gì đó một cách rất hào hứng, thích thú."
      }
    },
    {
      id: "l1_q13",
      lang: "vi",
      prompt: "Trong cụm từ <strong>«когда-то очень давно»</strong>, hậu tố <strong>«-то»</strong> mang ý nghĩa ngữ pháp gì?",
      options: [
        { key: "A", text: "A. Phủ định hoàn toàn thời gian" },
        { key: "B", text: "B. Chỉ một thời điểm bất định trong quá khứ mà người nói không nhớ rõ chính xác", correct: true },
        { key: "C", text: "C. Nhấn mạnh một mốc thời gian tương lai chắc chắn xảy ra" },
        { key: "D", text: "D. Biểu thị điều kiện giả định" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! Hậu tố «-то» biểu thị ý nghĩa bất định (không xác định chính xác thời điểm trong quá khứ).",
        other: "Chưa đúng. «Когда-то» là đại từ/phó từ bất định chỉ một lúc nào đó trong quá khứ mà người nói không nhớ cụ thể."
      }
    },
    {
      id: "l1_q14",
      lang: "vi",
      prompt: "Động từ phản thân <strong>«читаются»</strong> trong câu «книги читаются очень легко» biểu đạt ý nghĩa gì?",
      options: [
        { key: "A", text: "A. Những cuốn sách tự biết đọc to cho mọi người nghe" },
        { key: "B", text: "B. Đặc tính của cuốn sách được viết rất cuốn hút, dễ đọc và dễ tiếp nhận", correct: true },
        { key: "C", text: "C. Sách chỉ dành riêng cho trẻ em tập đọc" },
        { key: "D", text: "D. Cuốn sách bị ép buộc phải đọc" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! Dạng động từ phản thân «читается» đi kèm trạng từ «легко» chỉ đặc tính dễ đọc, văn phong trôi chảy.",
        other: "Chưa đúng. Đây là dạng động từ phản thân mang ý nghĩa miêu tả phẩm chất/đặc tính dễ đọc của tác phẩm."
      }
    },
    {
      id: "l1_q15",
      lang: "vi",
      prompt: "Từ <strong>«живые люди»</strong> khi dùng để miêu tả các nhân vật của Ustinova có nghĩa là:",
      options: [
        { key: "A", text: "A. Những người còn đang sống và thở về mặt sinh học" },
        { key: "B", text: "B. Những nhân vật chân thực, có cảm xúc thật và gần gũi với đời thường chứ không sáo rỗng", correct: true },
        { key: "C", text: "C. Những người có sức khỏe vượt trội và bất tử" },
        { key: "D", text: "D. Những nhân vật hoạt hình vui nhộn" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! «Живые люди» trong văn học là phép ẩn dụ chỉ những nhân vật sinh động, chân thực như bước ra từ đời thực.",
        other: "Chưa đúng. Trong ngữ cảnh văn học, «живые люди» mang nghĩa là nhân vật sống động, chân thực và gần gũi."
      }
    },
    {
      id: "l1_q16",
      lang: "vi",
      prompt: "Chi tiết nào cho thấy không khí sinh hoạt trong gia đình nữ nhà văn vô cùng sôi nổi và náo nhiệt?",
      options: [
        { key: "A", text: "A. Trong nhà hoàn toàn im ắng và mọi người chỉ giao tiếp bằng tin nhắn" },
        { key: "B", text: "B. Bếp liên tục nấu nướng, trẻ em nô đùa, chú chó Kuzya chạy nhảy trong vườn", correct: true },
        { key: "C", text: "C. Mọi người không bao giờ ăn cơm cùng nhau" },
        { key: "D", text: "D. Gia đình chỉ mở cửa vào các dịp lễ tết" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! Bài viết miêu tả: bếp luôn nấu ăn, trẻ con chơi đùa, chó Kuzya chạy ngoài vườn, mọi người to tiếng rồi lại làm lành.",
        other: "Chưa đúng. Đoạn văn nêu rõ: «На кухне постоянно готовится еда, дети играют, а в саду весело бегает собака Кузя»."
      }
    },
    {
      id: "l1_q17",
      lang: "vi",
      prompt: "Nguồn cảm hứng sáng tác cho các cuốn tiểu thuyết của T. Ustinova chủ yếu đến từ đâu?",
      options: [
        { key: "A", text: "A. Những chuyến du lịch thám hiểm đến các vùng đất hoang sơ" },
        { key: "B", text: "B. Chính ngôi nhà, khu vườn và những tình huống hằng ngày trong gia đình bà", correct: true },
        { key: "C", text: "C. Các vụ án mạng bí ẩn đăng trên báo nước ngoài" },
        { key: "D", text: "D. Trí tưởng tượng viễn tưởng về tương lai" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! Bài viết nêu rõ: «Этот дом, сад и эта семья, где все кричат, ссорятся и мирятся, появляются в каждом романе...».",
        other: "Chưa đúng theo bài đọc. Chính ngôi nhà, khu vườn và gia đình bà là chất liệu xuất hiện trong mỗi cuốn tiểu thuyết."
      }
    },
    {
      id: "l1_q18",
      lang: "vi",
      prompt: "Từ <strong>«статья»</strong> trong câu «была напечатана статья» thuộc giống gì và có nghĩa là gì?",
      options: [
        { key: "A", text: "A. Giống đực - Một cuốn từ điển" },
        { key: "B", text: "B. Giống cái - Một bài báo / bài viết", correct: true },
        { key: "C", text: "C. Giống trung - Một bức ảnh phóng sự" },
        { key: "D", text: "D. Số nhiều - Những tờ rơi quảng cáo" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! «Статья» là danh từ giống cái (đuôi -я), có nghĩa là bài báo, bài viết.",
        other: "Chưa đúng. «Статья» là danh từ giống cái, mang nghĩa là bài báo/bài viết."
      }
    },
    {
      id: "l1_q19",
      lang: "vi",
      prompt: "Tên tờ báo đưa tin về gia đình nhà văn Tatiana Ustinova là gì?",
      options: [
        { key: "A", text: "A. Газета «Правда»" },
        { key: "B", text: "B. Газета «Труд»", correct: true },
        { key: "C", text: "C. Газета «Известия»" },
        { key: "D", text: "D. Газета «Культура»" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! Câu đầu tiên của bài đọc: «Недавно в газете «Труд» была напечатана статья...».",
        other: "Chưa đúng. Ngay câu mở đầu văn bản có ghi: «в газете «Труд»»."
      }
    },
    {
      id: "l1_q20",
      lang: "vi",
      prompt: "Tổng kết thông điệp chính của văn bản về bí quyết thành công của nhà văn T. Ustinova là:",
      options: [
        { key: "A", text: "A. Sử dụng kỹ thuật công nghệ máy tính tiên tiến nhất để viết sách" },
        { key: "B", text: "B. Sống tách biệt và không tiếp xúc với người bình thường" },
        { key: "C", text: "C. Bắt nguồn từ cuộc sống gia đình ấm áp, tràn ngập tình yêu thương và thấu hiểu con người thực tế", correct: true },
        { key: "D", text: "D. Chi tiêu thật nhiều tiền cho các chiến dịch quảng cáo rầm rộ" }
      ],
      ans: "C",
      expl: {
        C: "Chính xác! Bài viết nhấn mạnh sự gắn kết gia đình, sự ấm cúng và sự chân thực đời thường chính là chìa khóa khiến sách của bà được đón nhận nồng nhiệt.",
        other: "Chưa đúng. Thông điệp cốt lõi là tình cảm gia đình hòa thuận và nhân vật gần gũi đời thường đã làm nên sức hút tác phẩm."
      }
    }
  ]
};

// ============================================================================
// LESSON 2 DATA (20 câu: 10 Nga + 10 Việt; 10 Từ vựng + 10 Đọc hiểu/Ngữ pháp)
// ============================================================================
const lesson2 = {
  id: 2,
  title: "Центры планирования семьи: сколько детей должно быть в семье?",
  badge: "BÀI ĐỌC 2",
  time: "25-30 phút",
  source: "Xã hội & Gia đình hiện đại",
  topic: "Xã hội & Kế hoạch hóa gia đình",
  paragraphs: [
    "Сейчас почти во всех больших городах России создаются центры планирования семьи. В эти центры приходят молодые люди, чтобы получить консультацию опытных психологов, врачей и педагогов. В этих центрах обсуждаются наиболее важные проблемы молодой семьи: материальные проблемы, проблема жилья, проблема рождения и воспитания ребёнка.",
    "В последнее время психологами и педагогами часто обсуждается вопрос: <strong>сколько детей должно быть в семье — один или несколько?</strong> В каждой семье эта проблема решается по-своему.",
    "Многодетные семьи сейчас «вышли из моды». Считается, что один ребёнок — это удобно. Одному ребёнку в семье уделяется больше времени и внимания, он лучше обеспечивается материально. Но единственные дети, особенно подростки, часто говорят о своём одиночестве. Такие дети могут быть неуверенными в себе, в своих силах, потому что постоянно опекаются своими папами, мамами, бабушками и дедушками.",
    "Конечно, если в семье много детей, родители не могут проводить много времени с каждым ребёнком. Но в многодетных семьях младшие дети воспитываются старшими братьями и сёстрами, которые играют и занимаются с ними. Дети в больших семьях быстро становятся самостоятельными и уверенными в себе."
  ],
  caption: "🕊️ <strong>Chủ đề trọng tâm:</strong> Kế hoạch hóa gia đình & Lựa chọn số lượng con (1 con vs. đông con).",
  vocabBoxTitle: "Từ vựng & Cấu trúc quản trị cách (Bài 2)",
  vocabCategories: [
    {
      title: "Động từ & Cấu trúc quản trị cách (Глагольное управление)",
      badge: "Cốt lõi bài đọc",
      badgeClass: "cat-verb",
      headers: ["Động từ", "Cấu trúc cách & Ý nghĩa", "Ví dụ trong bài"],
      rows: [
        { word: "уделя́ть / удели́ть", gram: "НСВ / СВ", meaning: "Dành (thời gian, sự chú ý) cho ai (Cách 3)", example: "Одному ребёнку <strong>уделяется</strong> больше времени..." },
        { word: "опека́ть", gram: "НСВ", meaning: "Bao bọc, chở che, giám sát quá mức (Cách 4)", example: "...постоянно <strong>опекаются</strong> папами и мамами..." },
        { word: "воспи́тывать / воспита́ть", gram: "НСВ / СВ", meaning: "Nuôi dạy, giáo dục ai (Cách 4)", example: "...младшие дети <strong>воспитываются</strong> старшими братьями..." },
        { word: "обеспе́чивать / обеспечи́ть", gram: "НСВ / СВ", meaning: "Chu cấp cho ai cái gì: кого? (4) чем? (5)", example: "...он лучше <strong>обеспечивается</strong> материально." },
        { word: "обсужда́ться / реши́ться", gram: "Dạng bị động (-ся)", meaning: "Được thảo luận / Được giải quyết", example: "В каждой семье проблема <strong>решается</strong> по-своему." }
      ]
    },
    {
      title: "Danh từ & Tính từ chủ đề",
      badge: "Từ vựng gia đình",
      badgeClass: "cat-noun",
      headers: ["Từ vựng", "Ý nghĩa tiếng Việt", "Ngữ cảnh"],
      rows: [
        { word: "плани́рование семьи́", gram: "Danh từ ghép", meaning: "Kế hoạch hóa gia đình", example: "Центры планирования семьи" },
        { word: "многоде́тная семья́", gram: "Tính từ + Danh từ", meaning: "Gia đình đông con", example: "Многодетные семьи вышли из моды" },
        { word: "одино́чество", gram: "Danh từ (Giống trung)", meaning: "Sự cô đơn, lẻ loi", example: "...говорят о своём одиночестве" },
        { word: "самостоя́тельный", gram: "Tính từ", meaning: "Tự lập, tự giác", example: "Дети быстро становятся самостоятельными" },
        { word: "уве́ренный в себе́", gram: "Cụm tính từ", meaning: "Tự tin vào bản thân", example: "неуверенные в себе ↔ уверенные в себе" }
      ]
    },
    {
      title: "Thành ngữ & Cụm từ cố định",
      badge: "Idioms & Expressions",
      badgeClass: "cat-expr",
      headers: ["Cụm từ", "Nghĩa tiếng Việt", "Ghi chú"],
      rows: [
        { word: "вы́йти из мо́ды", gram: "Thành ngữ", meaning: "Không còn là xu hướng, lỗi thời" },
        { word: "по-сво́ему", gram: "Trạng từ", meaning: "Theo cách riêng của mình" },
        { word: "в свои́х си́лах", gram: "Cụm giới từ", meaning: "Vào năng lực/sức lực của mình" }
      ]
    }
  ],
  questions: [
    // 1-10: CÂU HỎI TIẾNG NGA
    {
      id: "l2_q1",
      lang: "ru",
      prompt: "Какая главная проблема обсуждается в этом тексте?",
      options: [
        { key: "A", text: "A. Нехватка детских садов и школ в крупных городах России" },
        { key: "B", text: "B. Проблема жилья и ипотечных кредитов для молодых специалистов" },
        { key: "C", text: "C. Вопрос о том, сколько детей должно быть в семье, и особенности их воспитания", correct: true },
        { key: "D", text: "D. Трудности в работе центров психологической помощи" }
      ],
      ans: "C",
      expl: {
        C: "Правильно! Основная тема текста — сколько детей должно быть в семье (один или несколько) и как это влияет на детей.",
        other: "Неверно. Главный вопрос текста — количество детей в семье и специфика их развития."
      }
    },
    {
      id: "l2_q2",
      lang: "ru",
      prompt: "Что означает метафора <strong>«вышли из моды»</strong> в отношении многодетных семей?",
      options: [
        { key: "A", text: "A. Большие семьи перестали модно одеваться" },
        { key: "B", text: "B. Иметь много детей стало непопулярным и непривычным явлением", correct: true },
        { key: "C", text: "C. Государство запретило иметь больше двух детей" },
        { key: "D", text: "D. Многодетные семьи стали популярнее, чем раньше" }
      ],
      ans: "B",
      expl: {
        B: "Правильно! «Выйти из моды» означает потерять былую популярность, перестать быть общепринятым трендом.",
        other: "Неверно. В данном контексте это означает потерю популярности и актуальности в обществе."
      }
    },
    {
      id: "l2_q3",
      lang: "ru",
      prompt: "Какого падежа требует глагол <strong>«уделять (внимание)»</strong>? Выберите правильный вариант: «Родители уделяют много внимания ________.»",
      options: [
        { key: "A", text: "A. детям (Дательный падеж)", correct: true },
        { key: "B", text: "B. детей (Родительный / Винительный падеж)" },
        { key: "C", text: "C. детьми (Творительный падеж)" },
        { key: "D", text: "D. о детях (Предложный падеж)" }
      ],
      ans: "A",
      expl: {
        A: "Правильно! Глагол «уделять внимание» требует Дательного падежа: кому? -> детям.",
        other: "Ошибка в падеже. Уделять внимание кому? — Дательный падеж (детям)."
      }
    },
    {
      id: "l2_q4",
      lang: "ru",
      prompt: "Выберите правильную форму: «Родители чрезмерно опекают ________.»",
      hint: "Опекать (кого? Винительный падеж одушевлённых существительных)",
      options: [
        { key: "A", text: "A. единственному сыну" },
        { key: "B", text: "B. единственным сыном" },
        { key: "C", text: "C. единственного сына", correct: true },
        { key: "D", text: "D. о единственном сыне" }
      ],
      ans: "C",
      expl: {
        C: "Правильно! Глагол «опекать» управляет Винительным падежом: кого? -> единственного сына.",
        other: "Неверно. Опекать кого? — Винительный падеж (единственного сына)."
      }
    },
    {
      id: "l2_q5",
      lang: "ru",
      prompt: "Какое грамматическое управление имеет глагол <strong>«обеспечивать / обеспечить»</strong>?",
      options: [
        { key: "A", text: "A. кому? (3) что? (4)" },
        { key: "B", text: "B. кого? (4) чем? (5)", correct: true },
        { key: "C", text: "C. кого? (4) о чём? (6)" },
        { key: "D", text: "D. кого? (2) к чему? (3)" }
      ],
      ans: "B",
      expl: {
        B: "Правильно! Обеспечивать кого? (Винительный падеж) чем? (Творительный падеж).",
        other: "Ошибка. Стандартное управление: обеспечить кого? (4) чем? (5)."
      }
    },
    {
      id: "l2_q6",
      lang: "ru",
      prompt: "Вставьте правильную форму: «Семья стремится обеспечить ребёнка качественным ________.»",
      options: [
        { key: "A", text: "A. образование (В.п.)" },
        { key: "B", text: "B. образованием (Т.п.)", correct: true },
        { key: "C", text: "C. образования (Р.п.)" },
        { key: "D", text: "D. к образованию (Д.п.)" }
      ],
      ans: "B",
      expl: {
        B: "Правильно! Обеспечить кого? (ребёнка) чем? -> качественным образованием (Творительный падеж).",
        other: "Неверно. Глагол «обеспечить» требует Творительного падежа: чем? (образованием)."
      }
    },
    {
      id: "l2_q7",
      lang: "ru",
      prompt: "Кем воспитываются младшие дети в многодетных семьях согласно тексту?",
      options: [
        { key: "A", text: "A. Только школьными учителями" },
        { key: "B", text: "B. Нянями и гувернантками" },
        { key: "C", text: "C. Старшими братьями и сёстрами", correct: true },
        { key: "D", text: "D. Специалистами центров планирования" }
      ],
      ans: "C",
      expl: {
        C: "Правильно! Текст: «...в многодетных семьях младшие дети воспитываются старшими братьями и сёстрами».",
        other: "Неверно. По тексту младшие дети воспитываются старшими братьями и сёстрами."
      }
    },
    {
      id: "l2_q8",
      lang: "ru",
      prompt: "Почему единственные дети в семье могут чувствовать себя неуверенными в себе?",
      options: [
        { key: "A", text: "A. Потому что им не хватает денег на карманные расходы" },
        { key: "B", text: "B. Потому что они постоянно опекаются взрослыми членами семьи", correct: true },
        { key: "C", text: "C. Потому что родители никогда не уделяют им времени" },
        { key: "D", text: "D. Потому что они живут далеко от школы" }
      ],
      ans: "B",
      expl: {
        B: "Правильно! Причина в гиперопеке: они постоянно опекаются папами, мамами, бабушками и дедушками.",
        other: "Неверно. Текст указывает на чрезмерную опеку со стороны всех взрослых родственников."
      }
    },
    {
      id: "l2_q9",
      lang: "ru",
      prompt: "Каким качеством быстрее овладевают дети, выросшие в больших семьях?",
      options: [
        { key: "A", text: "A. Они быстрее становятся самостоятельными", correct: true },
        { key: "B", text: "B. Они становятся равнодушными к другим людям" },
        { key: "C", text: "C. Они всегда выбирают только творческие профессии" },
        { key: "D", text: "D. Они не любят общаться с людьми" }
      ],
      ans: "A",
      expl: {
        A: "Правильно! В тексте: «Дети в больших семьях быстро становятся самостоятельными и уверенными в себе».",
        other: "Неверно. Главное преимущество — быстрое развитие самостоятельности и уверенности."
      }
    },
    {
      id: "l2_q10",
      lang: "ru",
      prompt: "Зачем молодые люди обращаются в центры планирования семьи?",
      options: [
        { key: "A", text: "A. Чтобы получить консультацию опытных психологов, врачей и педагогов", correct: true },
        { key: "B", text: "B. Чтобы бесплатно получить квартиру от государства" },
        { key: "C", text: "C. Чтобы найти работу для всех членов семьи" },
        { key: "D", text: "D. Чтобы купить детские товары со скидкой" }
      ],
      ans: "A",
      expl: {
        A: "Правильно! Текст: «...чтобы получить консультацию опытных психологов, врачей и педагогов».",
        other: "Неверно. Цель — получить консультацию квалифицированных специалистов."
      }
    },

    // 11-20: CÂU HỎI TIẾNG VIỆT
    {
      id: "l2_q11",
      lang: "vi",
      prompt: "Khái niệm <strong>«центры планирования семьи»</strong> dịch chuẩn xác nhất sang tiếng Việt là gì?",
      options: [
        { key: "A", text: "A. Trung tâm môi giới hôn nhân và gia đình" },
        { key: "B", text: "B. Trung tâm kế hoạch hóa gia đình", correct: true },
        { key: "C", text: "C. Trung tâm bảo trợ xã hội trẻ mồ côi" },
        { key: "D", text: "D. Trung tâm tư vấn tài chính bất động sản" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! «Центр планирования семьи» là trung tâm kế hoạch hóa gia đình.",
        other: "Chưa đúng. Đây là thuật ngữ y tế - xã hội chỉ Trung tâm kế hoạch hóa gia đình."
      }
    },
    {
      id: "l2_q12",
      lang: "vi",
      prompt: "Từ <strong>«одиночество»</strong> (giống trung) có nghĩa là gì?",
      options: [
        { key: "A", text: "A. Sự ích kỷ, tham lam" },
        { key: "B", text: "B. Sự cô đơn, lẻ loi, thiếu người chia sẻ", correct: true },
        { key: "C", text: "C. Sự tự tin thái quá" },
        { key: "D", text: "D. Sự bận rộn trong công việc" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! «Одиночество» nghĩa là cảm giác cô đơn, cô độc.",
        other: "Chưa đúng. «Одиночество» bắt nguồn từ gốc «один» (một mình), chỉ sự cô đơn, lẻ loi."
      }
    },
    {
      id: "l2_q13",
      lang: "vi",
      prompt: "Từ <strong>«самостоятельный»</strong> miêu tả một đứa trẻ có phẩm chất gì?",
      options: [
        { key: "A", text: "A. Luôn dựa dẫm và chờ người khác làm hộ" },
        { key: "B", text: "B. Tự lập, tự giác, có khả năng tự giải quyết công việc của mình", correct: true },
        { key: "C", text: "C. Nhút nhát và sợ chỗ đông người" },
        { key: "D", text: "D. Hay cãi lời và bướng bỉnh" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! «Самостоятельный» (tự lập, độc lập) là đứa trẻ tự biết chăm sóc và giải quyết việc của mình.",
        other: "Chưa đúng. «Самостоятельный» bắt nguồn từ «сам» (tự mình) + «стоять» (đứng vững), nghĩa là tự lập."
      }
    },
    {
      id: "l2_q14",
      lang: "vi",
      prompt: "Động từ <strong>«опекать»</strong> trong bài nói về việc cha mẹ nuôi dạy con một mang sắc thái ý nghĩa:",
      options: [
        { key: "A", text: "A. Bỏ bê con cái, không quan tâm chăm sóc" },
        { key: "B", text: "B. Bao bọc quá mức (giám sát, làm thay mọi việc khiến con mất tự lập)", correct: true },
        { key: "C", text: "C. Rèn luyện kỷ luật nghiêm khắc như quân đội" },
        { key: "D", text: "D. Cho con hoàn toàn tự do muốn làm gì tùy thích" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! «Опекать» trong tâm lý học giáo dục thường chỉ sự bao bọc quá mức (гиперопека).",
        other: "Chưa đúng. Ngữ cảnh bài viết chỉ sự chở che, bao bọc quá mức của người lớn khiến trẻ thiếu tự tin."
      }
    },
    {
      id: "l2_q15",
      lang: "vi",
      prompt: "Trạng từ <strong>«по-своему»</strong> trong câu «эта проблема решается по-своему» có nghĩa là gì?",
      options: [
        { key: "A", text: "A. Bắt buộc theo cùng một công thức chung giống hệt nhau" },
        { key: "B", text: "B. Theo cách riêng, quan điểm riêng của từng gia đình", correct: true },
        { key: "C", text: "C. Theo chỉ thị bắt buộc của chính quyền" },
        { key: "D", text: "D. Theo lời khuyên duy nhất của bác sĩ" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! «По-своему» có nghĩa là theo cách riêng, lối riêng của mỗi người/mỗi nhà.",
        other: "Chưa đúng. «По-своему» biểu thị hành động diễn ra theo cách thức riêng của chủ thể."
      }
    },
    {
      id: "l2_q16",
      lang: "vi",
      prompt: "Ưu điểm lớn nhất của việc sinh duy nhất một con (con một) được đề cập trong bài là gì?",
      options: [
        { key: "A", text: "A. Đứa trẻ không bao giờ bị ốm đau" },
        { key: "B", text: "B. Tiện lợi, cha mẹ có nhiều thời gian quan tâm và chu cấp điều kiện vật chất tốt hơn cho con", correct: true },
        { key: "C", text: "C. Cha mẹ không cần phải đi làm kiếm tiền" },
        { key: "D", text: "D. Đứa trẻ chắc chắn sẽ trở thành thiên tài" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! Bài viết chỉ ra: «Одному ребёнку в семье уделяется больше времени и внимания, он лучше обеспечивается материально».",
        other: "Chưa đúng. Ưu điểm được nêu trong bài là đứa trẻ nhận được nhiều thời gian quan tâm và điều kiện vật chất tốt hơn."
      }
    },
    {
      id: "l2_q17",
      lang: "vi",
      prompt: "Mặt hạn chế (nhược điểm) mà cha mẹ đông con thường phải đối mặt là gì?",
      options: [
        { key: "A", text: "A. Các con không bao giờ chịu nói chuyện với nhau" },
        { key: "B", text: "B. Cha mẹ không thể dành nhiều thời gian riêng cho từng đứa trẻ", correct: true },
        { key: "C", text: "C. Trẻ em trong nhà luôn luôn nhút nhát" },
        { key: "D", text: "D. Không có trường học nào nhận dạy trẻ gia đình đông con" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! Bài viết: «...родители не могут проводить много времени с каждым ребёнком».",
        other: "Chưa đúng. Nhược điểm của gia đình đông con là cha mẹ khó dành nhiều thời gian riêng cho từng đứa con."
      }
    },
    {
      id: "l2_q18",
      lang: "vi",
      prompt: "Đối tượng nào trong bài đọc thường hay bộc lộ cảm giác cô đơn nhất?",
      options: [
        { key: "A", text: "A. Các bậc phụ huynh đi làm xa" },
        { key: "B", text: "B. Những đứa trẻ là con một, đặc biệt là lứa tuổi thiếu niên (подростки)", correct: true },
        { key: "C", text: "C. Những đứa em út trong gia đình đông con" },
        { key: "D", text: "D. Các bác sĩ tâm lý tại trung tâm" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! Trích dẫn bài đọc: «Но единственные дети, особенно подростки, часто говорят о своём одиночестве».",
        other: "Chưa đúng theo văn bản. Đối tượng hay chia sẻ về sự cô đơn là con một, đặc biệt là lứa tuổi thiếu niên."
      }
    },
    {
      id: "l2_q19",
      lang: "vi",
      prompt: "Cụm danh từ <strong>«материальные проблемы»</strong> chỉ nhóm vấn đề nào sau đây?",
      options: [
        { key: "A", text: "A. Vấn đề sức khỏe tinh thần" },
        { key: "B", text: "B. Vấn đề tài chính, tiền bạc, kinh tế gia đình", correct: true },
        { key: "C", text: "C. Mâu thuẫn giữa mẹ chồng và nàng dâu" },
        { key: "D", text: "D. Vấn đề bất đồng ngôn ngữ" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! «Материальные проблемы» là các khó khăn liên quan đến vật chất, tài chính, tiền bạc.",
        other: "Chưa đúng. «Материальный» chỉ phương diện vật chất, kinh tế, tài chính."
      }
    },
    {
      id: "l2_q20",
      lang: "vi",
      prompt: "Kết luận logic từ bài đọc: Việc quyết định sinh 1 hay nhiều con nên dựa trên điều gì?",
      options: [
        { key: "A", text: "A. Chạy theo xu hướng số đông của xã hội" },
        { key: "B", text: "B. Hoàn cảnh cụ thể, khả năng kinh tế và sự cân nhắc kỹ lưỡng giữa mặt tích cực và hạn chế của mỗi gia đình", correct: true },
        { key: "C", text: "C. Chỉ dựa vào ý kiến của ông bà" },
        { key: "D", text: "D. Bắt buộc phải có từ 3 con trở lên" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! Bài viết nhấn mạnh «В каждой семье эта проблема решается по-своему», cả hai mô hình đều có ưu và nhược điểm riêng cần cân nhắc.",
        other: "Chưa đúng. Thông điệp bài đọc là mỗi gia đình tự quyết định dựa trên hoàn cảnh và sự thấu hiểu điểm mạnh - điểm yếu của từng mô hình."
      }
    }
  ]
};

// ============================================================================
// LESSON 3 DATA (20 câu: 10 Nga + 10 Việt; 10 Từ vựng + 10 Đọc hiểu/Ngữ pháp)
// ============================================================================
const lesson3 = {
  id: 3,
  title: "Пьеса «Чайка» А. П. Чехова и символ Московского Художественного театра",
  badge: "BÀI ĐỌC 3",
  time: "25-30 phút",
  source: "УРОК 1. ЛЮДЕЙ НЕИНТЕРЕСНЫХ В МИРЕ НЕТ... (Trang 17)",
  topic: "Văn học & Nghệ thuật kịch nghệ Nga",
  paragraphs: [
    "Пьеса «Чайка» была написана Антоном Павловичем Чеховым в 1896 году в Мелихове. Эта пьеса была особенно дорога Чехову, потому что в ней писатель впервые рассказал о нелёгкой судьбе людей, которые, как и сам Чехов, выбрали творческую профессию. Герои пьесы — девушка, которая мечтает стать великой актрисой, и молодой начинающий писатель. Оба они мечтают о славе, верят в свой талант и в свои силы, но ещё не знают, какие трудности встретятся им на их нелёгком пути.",
    "Впервые «Чайка» была поставлена в 1896 году в Санкт-Петербурге. К сожалению, пьеса не имела успеха, потому что не была понята зрителями. Антон Павлович Чехов очень переживал эту неудачу. Поэтому, когда режиссёр Московского Художественного театра Константин Сергеевич Станиславский предложил Чехову поставить «Чайку» в Москве, Антон Павлович согласился не сразу. Он пришёл в театр, познакомился с артистами, читал им свою пьесу, был на репетициях спектакля и убедился, что идея его пьесы была правильно понята режиссёром и молодыми талантливыми артистами. Так началась творческая дружба писателя с театром.",
    "В 1898 году чеховская пьеса «Чайка» была поставлена на сцене Московского Художественного театра. Пьеса имела огромный успех. Она была хорошо встречена зрителями. Они долго аплодировали, бросали цветы на сцену и не хотели уходить из театра. С тех пор летящая чайка стала символом Московского Художественного театра. Этот символ и сейчас можно увидеть на здании этого театра, на его занавесе и на театральных билетах."
  ],
  caption: "🕊️ <strong>Câu hỏi trọng tâm (а):</strong> Символ МХТ — летящая чайка (увидеть: на здании театра, на занавесе, на театральных билетах).",
  vocabBoxTitle: "Từ vựng, Thuật ngữ Sân khấu & Ngữ pháp (Bài 3)",
  vocabCategories: [
    {
      title: "Động từ, Thể bị động & Quản trị cách",
      badge: "Глаголы & Управление",
      badgeClass: "cat-verb",
      headers: ["Động từ", "Nghĩa & Quản trị cách", "Ví dụ trong bài"],
      rows: [
        { word: "поста́вить (пьесу)", gram: "СВ (bị động: поставлена)", meaning: "Dàn dựng, công diễn vở kịch", example: "...была <strong>поставлена</strong> на сцене МХТ" },
        { word: "мечта́ть", gram: "НСВ (о ком? о чём?)", meaning: "Mơ ước về điều gì (Cách 6)", example: "Оба они <strong>мечтают о славе</strong>..." },
        { word: "ве́рить", gram: "НСВ (в кого? в что?)", meaning: "Tin tưởng vào điều gì (Cách 4)", example: "...<strong>верят в свой талант</strong> и в свои силы" },
        { word: "пережива́ть", gram: "НСВ (что? Cách 4)", meaning: "Đau buồn, trăn trở về điều gì", example: "Чехов очень <strong>переживал эту неудачу</strong>" },
        { word: "убеди́ться", gram: "СВ (в чём?)", meaning: "Tin chắc, kiểm chứng chắc chắn", example: "...и <strong>убедился</strong>, что идея была понята" },
        { word: "аплоди́ровать", gram: "НСВ (кому? Cách 3)", meaning: "Vỗ tay hoan nghênh ai", example: "Они долго <strong>аплодировали</strong>..." }
      ]
    },
    {
      title: "Thuật ngữ Sân khấu & Từ vựng chủ đề",
      badge: "Театр и литература",
      badgeClass: "cat-noun",
      headers: ["Từ vựng", "Ý nghĩa tiếng Việt", "Ngữ cảnh"],
      rows: [
        { word: "пье́са", gram: "Danh từ (Giống cái)", meaning: "Vở kịch bản sân khấu", example: "Пьеса «Чайка» А. П. Чехова" },
        { word: "режиссёр", gram: "Danh từ (Giống đực)", meaning: "Đạo diễn sân khấu", example: "Режиссёр К. С. Станиславский" },
        { word: "репети́ция", gram: "Danh từ (Giống cái)", meaning: "Buổi tập dượt, tổng duyệt", example: "...был на репетициях спектакля" },
        { word: "за́навес", gram: "Danh từ (Giống đực)", meaning: "Màn sân khấu kịch", example: "...на занавесе театра" },
        { word: "летя́щая ча́йка", gram: "Phân từ + Danh từ", meaning: "Chim hải âu đang sải cánh bay", example: "Символ МХТ — летящая чайка" },
        { word: "нелёгкая судьба́", gram: "Tính từ + Danh từ", meaning: "Số phận gian nan, trắc trở", example: "...о нелёгкой судьбе творческих людей" }
      ]
    },
    {
      title: "Cụm từ cố định & Thành ngữ trong bài",
      badge: "Коллокации",
      badgeClass: "cat-expr",
      headers: ["Cụm từ", "Nghĩa tiếng Việt", "Ngữ pháp"],
      rows: [
        { word: "име́ть успе́х", gram: "Thành ngữ", meaning: "Đạt được thành công vang dội" },
        { word: "с тех пор", gram: "Liên từ thời gian", meaning: "Kể từ đó trở đi" },
        { word: "быть доро́гим кому?", gram: "Tính từ ngắn + Cách 3", meaning: "Rất thân thương, quý giá đối với ai" }
      ]
    }
  ],
  questions: [
    // 1-10: CÂU HỎI TIẾNG NGA
    {
      id: "l3_q1",
      lang: "ru",
      prompt: "Что представляет собой жанр <strong>«пьеса»</strong>?",
      options: [
        { key: "A", text: "A. Лирическое стихотворение о природе" },
        { key: "B", text: "B. Драматическое произведение, предназначенное для постановки на сцене театра", correct: true },
        { key: "C", text: "C. Научно-популярная статья об истории искусства" },
        { key: "D", text: "D. Музыкальное инструментальное сочинение без слов" }
      ],
      ans: "B",
      expl: {
        B: "Правильно! Пьеса — это литературное произведение драматического рода, написанное для театральной сцены.",
        other: "Неверно. Пьеса — это драматическое произведение для театра."
      }
    },
    {
      id: "l3_q2",
      lang: "ru",
      prompt: "Какой символ имеет Московский Художественный театр (МХТ) согласно тексту?",
      options: [
        { key: "A", text: "A. Белый лебедь на воде" },
        { key: "B", text: "B. Золотая театральная маска" },
        { key: "C", text: "C. Летящая чайка", correct: true },
        { key: "D", text: "D. Горящий факел искусства" }
      ],
      ans: "C",
      expl: {
        C: "Правильно! Текст: «С тех пор летящая чайка стала символом Московского Художественного театра».",
        other: "Неверно. Символ МХТ — летящая чайка."
      }
    },
    {
      id: "l3_q3",
      lang: "ru",
      prompt: "Где и сегодня можно увидеть символ летящей чайки?",
      options: [
        { key: "A", text: "A. На здании театра, на его занавесе и на театральных билетах", correct: true },
        { key: "B", text: "B. Только на обложках чеховских книг" },
        { key: "C", text: "C. Исключительно на картинах в Третьяковской галерее" },
        { key: "D", text: "D. На фасадах всех театров Санкт-Петербурга" }
      ],
      ans: "A",
      expl: {
        A: "Правильно! В тексте: «Этот символ и сейчас можно увидеть на здании этого театра, на его занавесе и на театральных билетах».",
        other: "Неверно. Символ изображён на здании, на занавесе и на билетах театра."
      }
    },
    {
      id: "l3_q4",
      lang: "ru",
      prompt: "Какую роль в театре выполняет <strong>режиссёр</strong>?",
      options: [
        { key: "A", text: "A. Продаёт билеты зрителям в кассе театра" },
        { key: "B", text: "B. Руководит созданием спектакля, работает с артистами над воплощением замысла пьесы", correct: true },
        { key: "C", text: "C. Шьёт костюмы и наносит грим актёрам" },
        { key: "D", text: "D. Пишет критические отзывы в газеты" }
      ],
      ans: "B",
      expl: {
        B: "Правильно! Режиссёр — это творческий руководитель постановки спектакля.",
        other: "Неверно. Режиссёр руководит постановкой спектакля и работой актёров."
      }
    },
    {
      id: "l3_q5",
      lang: "ru",
      prompt: "Что означает слово <strong>«репетиция»</strong> в театральном искусстве?",
      options: [
        { key: "A", text: "A. Праздничный банкет после премьеры" },
        { key: "B", text: "B. Предварительный подготовительный процесс подготовки и отработки спектакля", correct: true },
        { key: "C", text: "C. Закрытие театрального сезона" },
        { key: "D", text: "D. Продажа сувениров перед началом действия" }
      ],
      ans: "B",
      expl: {
        B: "Правильно! Репетиция — предварительное исполнение спектакля для подготовки к публичному выступлению.",
        other: "Неверно. Репетиция — это подготовка и отработка спектакля артистами."
      }
    },
    {
      id: "l3_q6",
      lang: "ru",
      prompt: "Какого предлога и падежа требует глагол <strong>«мечтать»</strong>? Выберите правильный вариант: «Герои мечтают ________.»",
      options: [
        { key: "A", text: "A. великую славу (В.п.)" },
        { key: "B", text: "B. о великой славе (о + П.п.)", correct: true },
        { key: "C", text: "C. к великой славе (к + Д.п.)" },
        { key: "D", text: "D. великой славой (Т.п.)" }
      ],
      ans: "B",
      expl: {
        B: "Правильно! Глагол «мечтать» требует конструкции: о ком? о чём? — Предложный падеж (о славе).",
        other: "Ошибка в падеже: мечтать о чём? — Предложный падеж."
      }
    },
    {
      id: "l3_q7",
      lang: "ru",
      prompt: "Вставьте правильную форму: «Молодые артисты свято верят ________.»",
      options: [
        { key: "A", text: "A. в свой талант (в + В.п.)", correct: true },
        { key: "B", text: "B. о своём таланте (о + П.п.)" },
        { key: "C", text: "C. своему таланту (Д.п.)" },
        { key: "D", text: "D. с своим талантом (Т.п.)" }
      ],
      ans: "A",
      expl: {
        A: "Правильно! Верить в кого? в что? требует Винительного падежа: «в свой талант».",
        other: "Ошибка. Глагол «верить в...» требует Винительного падежа."
      }
    },
    {
      id: "l3_q8",
      lang: "ru",
      prompt: "Какого падежа требует глагол <strong>«аплодировать»</strong>? Выберите правильный вариант: «Зрители долго аплодировали ________.»",
      options: [
        { key: "A", text: "A. талантливым артистам (Дательный падеж)", correct: true },
        { key: "B", text: "B. талантливых артистов (Винительный падеж)" },
        { key: "C", text: "C. талантливыми артистами (Творительный падеж)" },
        { key: "D", text: "D. о талантливых артистах (Предложный падеж)" }
      ],
      ans: "A",
      expl: {
        A: "Правильно! Аплодировать кому? — Дательный падеж: артистам.",
        other: "Ошибка в падеже. Аплодировать кому? — Дательный падеж."
      }
    },
    {
      id: "l3_q9",
      lang: "ru",
      prompt: "Почему первая постановка «Чайки» в Санкт-Петербурге в 1896 году потерпела неудачу?",
      options: [
        { key: "A", text: "A. Актёры забыли слова роли во время спектакля" },
        { key: "B", text: "B. Пьеса была новой и необычной и не была понята зрителями", correct: true },
        { key: "C", text: "C. Билеты были слишком дорогими для публики" },
        { key: "D", text: "D. В зале было слишком холодно из-за зимы" }
      ],
      ans: "B",
      expl: {
        B: "Правильно! В тексте прямо сказано: «К сожалению, пьеса не имела успеха, потому что не была понята зрителями».",
        other: "Неверно. Причина провала заключалась в том, что зрители не поняли новаторскую пьесу."
      }
    },
    {
      id: "l3_q10",
      lang: "ru",
      prompt: "Как отреагировали зрители на московскую премьеру «Чайки» в МХТ в 1898 году?",
      options: [
        { key: "A", text: "A. Ушли из зала до окончания действия" },
        { key: "B", text: "B. Долго аплодировали, бросали цветы на сцену и не хотели уходить", correct: true },
        { key: "C", text: "C. Свистели и выражали недовольство" },
        { key: "D", text: "D. Сохраняли полное молчание после спектакля" }
      ],
      ans: "B",
      expl: {
        B: "Правильно! Пьеса имела огромный успех: зрители аплодировали, бросали цветы и не хотели уходить.",
        other: "Неверно. Спектакль в Москве имел триумфальный успех."
      }
    },

    // 11-20: CÂU HỎI TIẾNG VIỆT
    {
      id: "l3_q11",
      lang: "vi",
      prompt: "Từ <strong>«занавес»</strong> trong kịch nghệ Nga có nghĩa là gì?",
      options: [
        { key: "A", text: "A. Bức rèm/màn sân khấu che khuất sàn diễn", correct: true },
        { key: "B", text: "B. Tấm vé vào cổng rạp hát" },
        { key: "C", text: "C. Bục đứng diễn thuyết của đạo diễn" },
        { key: "D", text: "D. Tờ chương trình giới thiệu diễn viên" }
      ],
      ans: "A",
      expl: {
        A: "Chính xác! «Занавес» là bức màn sân khấu của nhà hát.",
        other: "Chưa đúng. «Занавес» là màn sân khấu kịch."
      }
    },
    {
      id: "l3_q12",
      lang: "vi",
      prompt: "Ý nghĩa của cụm danh từ <strong>«нелёгкая судьба»</strong> là:",
      options: [
        { key: "A", text: "A. Cuộc sống an nhàn, bình lặng" },
        { key: "B", text: "B. Số phận gian nan, chông gai, đầy thăng trầm trắc trở", correct: true },
        { key: "C", text: "C. Con đường công danh thăng tiến thần tốc" },
        { key: "D", text: "D. Cuộc đời ngắn ngủi vô vị" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! «Нелёгкая судьба» miêu tả cuộc đời vất vả, gian truân của những người làm nghệ thuật.",
        other: "Chưa đúng. «Нелёгкая судьба» nghĩa là số phận gian nan, trắc trở."
      }
    },
    {
      id: "l3_q13",
      lang: "vi",
      prompt: "Động từ <strong>«убедиться»</strong> trong câu «убедился, что идея была правильно понята» mang nghĩa gì?",
      options: [
        { key: "A", text: "A. Hoài nghi và không tin tưởng" },
        { key: "B", text: "B. Tận mắt chứng kiến và tin chắc, yên tâm tuyệt đối", correct: true },
        { key: "C", text: "C. Từ chối hợp tác" },
        { key: "D", text: "D. Quên đi một sự việc" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! «Убедиться» nghĩa là tin chắc, kiểm chứng chắc chắn điều gì đó.",
        other: "Chưa đúng. «Убедиться» có nghĩa là tin chắc, nắm chắc điều gì."
      }
    },
    {
      id: "l3_q14",
      lang: "vi",
      prompt: "Cụm từ <strong>«творческая профессия»</strong> chỉ nhóm nghề nghiệp nào sau đây?",
      options: [
        { key: "A", text: "A. Kỹ sư xây dựng cầu đường" },
        { key: "B", text: "B. Nghề sáng tạo nghệ thuật (nhà văn, diễn viên, họa sĩ, nhạc sĩ...)", correct: true },
        { key: "C", text: "C. Kế toán ngân hàng" },
        { key: "D", text: "D. Công nhân sản xuất cơ khí" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! «Творческая профессия» là nghề nghiệp mang tính sáng tạo nghệ thuật (như nhà văn, diễn viên kịch).",
        other: "Chưa đúng. «Творческая профессия» là các nghề sáng tạo nghệ thuật."
      }
    },
    {
      id: "l3_q15",
      lang: "vi",
      prompt: "Tại sao ban đầu khi đạo diễn K. S. Stanislavski ngỏ lời đưa vở «Чайка» về Moskva, Chekhov lại do dự không đồng ý ngay?",
      options: [
        { key: "A", text: "A. Vì ông đòi mức thù lao quá cao" },
        { key: "B", text: "B. Vì ông còn rất đau buồn và ám ảnh sau thất bại đau đớn tại Peterburg năm 1896", correct: true },
        { key: "C", text: "C. Vì ông bận đi du lịch nước ngoài" },
        { key: "D", text: "D. Vì ông muốn giải nghệ không viết kịch nữa" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! Bài viết: «Антон Павлович Чехов очень переживал эту неудачу. Поэтому... согласился не сразу».",
        other: "Chưa đúng. Chekhov rất trăn trở và đau lòng trước thất bại trước đó nên ông không dám nhận lời ngay."
      }
    },
    {
      id: "l3_q16",
      lang: "vi",
      prompt: "Hai nhân vật chính trong vở «Чайка» đại diện cho những hình tượng nào?",
      options: [
        { key: "A", text: "A. Một vị tướng quân đội và một thương nhân giàu có" },
        { key: "B", text: "B. Một cô gái khao khát thành minh tinh màn bạc và một chàng trai trẻ mới bước vào nghiệp văn chương", correct: true },
        { key: "C", text: "C. Hai nhà khoa học tự nhiên" },
        { key: "D", text: "D. Đôi vợ chồng nông dân già ở nông thôn" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! Trích dẫn: «девушка, которая мечтает стать великой актрисой, и молодой начинающий писатель».",
        other: "Chưa đúng theo văn bản. Nhân vật là một cô gái mơ làm diễn viên lớn và một nhà văn trẻ mới vào nghề."
      }
    },
    {
      id: "l3_q17",
      lang: "vi",
      prompt: "Thành ngữ <strong>«с тех пор»</strong> trong tiếng Nga tương đương với từ nối nào trong tiếng Việt?",
      options: [
        { key: "A", text: "A. Kể từ thời điểm đó trở đi", correct: true },
        { key: "B", text: "B. Trước đó rất lâu" },
        { key: "C", text: "C. Trong một khoảng thời gian ngắn ngủi" },
        { key: "D", text: "D. Mãi cho đến tận ngày mai" }
      ],
      ans: "A",
      expl: {
        A: "Chính xác! «С тех пор» nghĩa là kể từ đó, từ mốc thời gian đó trở đi.",
        other: "Chưa đúng. «С тех пор» tương đương với 'kể từ đó trở đi'."
      }
    },
    {
      id: "l3_q18",
      lang: "vi",
      prompt: "Hành động <strong>«бросали цветы на сцену»</strong> của khán giả Moskva biểu thị điều gì?",
      options: [
        { key: "A", text: "A. Thể hiện sự phản đối kịch liệt" },
        { key: "B", text: "B. Bày tỏ lòng ái mộ, biết ơn và tôn vinh sâu sắc đối với các nghệ sĩ", correct: true },
        { key: "C", text: "C. Yêu cầu nhà hát dọn dẹp vệ sinh" },
        { key: "D", text: "D. Nghi thức tôn giáo bắt buộc" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! Tung hoa lên sân khấu kịch là cử chỉ văn hóa đẹp thể hiện sự mến mộ và thán phục tài năng diễn viên.",
        other: "Chưa đúng. Đây là nét đẹp văn hóa sân khấu thể hiện lòng cảm phục và tôn vinh người nghệ sĩ."
      }
    },
    {
      id: "l3_q19",
      lang: "vi",
      prompt: "Vở kịch «Чайка» được Anton Chekhov sáng tác vào năm nào và ở đâu?",
      options: [
        { key: "A", text: "A. Năm 1898 tại Moskva" },
        { key: "B", text: "B. Năm 1896 tại làng Melikhovo", correct: true },
        { key: "C", text: "C. Năm 1904 tại Badenweiler" },
        { key: "D", text: "D. Năm 1890 tại đảo Sakhalin" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! Câu đầu tiên: «Пьеса «Чайка» была написана Антоном Павловичем Чеховым в 1896 году в Мелихове».",
        other: "Chưa đúng theo câu mở đầu bài đọc: năm 1896 tại Melikhovo."
      }
    },
    {
      id: "l3_q20",
      lang: "vi",
      prompt: "Ý nghĩa biểu tượng của cánh hải âu «летящая чайка» đối với lịch sử Nhà hát Nghệ thuật Moskva (MXT) là gì?",
      options: [
        { key: "A", text: "A. Là loài chim đặc trưng sống gần con sông Moskva" },
        { key: "B", text: "B. Biểu tượng cho bước ngoặt hồi sinh vĩ đại của nhà hát, tình bạn nghệ thuật bất hủ giữa Chekhov và MXT", correct: true },
        { key: "C", text: "C. Biểu trưng cho ước mơ di cư ra nước ngoài" },
        { key: "D", text: "D. Là hình ảnh linh vật may mắn của riêng đạo diễn Stanislavski" }
      ],
      ans: "B",
      expl: {
        B: "Chính xác! «Чайка» là tác phẩm mang lại danh tiếng lẫy lừng cho MXT, gắn liền sự ra đời và trường tồn của trường phái kịch nghệ Chekhov - Stanislavski.",
        other: "Chưa đúng. Cánh hải âu tượng trưng cho bước ngoặt lịch sử và tình bạn nghệ thuật vĩ đại giữa Chekhov và MXT."
      }
    }
  ]
};

// ============================================================================
// HTML BUILDER HELPER
// ============================================================================
function renderVocabSection(categories) {
  let html = '';
  categories.forEach(cat => {
    html += `
              <div class="vocab-category-card">
                <div class="vocab-cat-header">
                  <div class="vocab-cat-title">
                    <span>${cat.badgeClass === 'cat-verb' ? '⚡' : (cat.badgeClass === 'cat-noun' ? '🏷️' : '🔗')}</span> ${cat.title}
                  </div>
                  <span class="cat-badge ${cat.badgeClass}">${cat.badge}</span>
                </div>
                <div class="table-responsive-wrapper">
                  <table class="vocab-table">
                    <thead>
                      <tr>
                        <th style="width: 32%;">${cat.headers[0]}</th>
                        <th style="width: 33%;">${cat.headers[1]}</th>
                        <th style="width: 35%;">${cat.headers[2]}</th>
                      </tr>
                    </thead>
                    <tbody>`;
    cat.rows.forEach(r => {
      html += `
                      <tr>
                        <td>
                          <span class="vocab-word">${r.word}</span>
                          <span class="vocab-gram">${r.gram || ''}</span>
                        </td>
                        <td><span class="vocab-meaning">${r.meaning}</span></td>
                        <td><span class="vocab-example">${r.example || ''}</span></td>
                      </tr>`;
    });
    html += `
                    </tbody>
                  </table>
                </div>
              </div>`;
  });
  return html;
}

function renderLessonLayout(lesson) {
  const pTags = lesson.paragraphs.map(p => `<p>${p}</p>`).join('\n            ');
  const vocabHtml = renderVocabSection(lesson.vocabCategories);

  // Group questions into 2 cards:
  // Card 1: 10 câu Tiếng Nga
  // Card 2: 10 câu Tiếng Việt
  const ruQuestions = lesson.questions.filter(q => q.lang === 'ru');
  const viQuestions = lesson.questions.filter(q => q.lang === 'vi');

  function renderQuestionCard(title, badgeText, qList) {
    let cardHtml = `
          <div class="quiz-card">
            <div class="quiz-card-header">
              <h3 class="quiz-section-title">
                ${title}
              </h3>
              <span style="font-size: 13px; font-weight: 700; color: var(--brand);">${badgeText}</span>
            </div>`;
    qList.forEach((q, idx) => {
      const qNum = lesson.questions.indexOf(q) + 1;
      const langBadge = q.lang === 'ru' 
        ? `<span class="lang-badge lang-ru">🇷🇺 RU</span>` 
        : `<span class="lang-badge lang-vi">🇻🇳 VI</span>`;

      cardHtml += `
            <div class="question-block" data-q="${q.id}" data-ans="${q.ans}">
              <div class="question-prompt">
                ${langBadge} <strong>Câu ${qNum}.</strong> ${q.prompt}
                ${q.hint ? `<span class="question-hint">💡 ${q.hint}</span>` : ''}
              </div>
              <div class="options-list">`;
      q.options.forEach(opt => {
        cardHtml += `
                <label class="option-item">
                  <input type="radio" name="${q.id}" value="${opt.key}" />
                  <span>${opt.text}</span>
                </label>`;
      });
      cardHtml += `
              </div>
              <div class="feedback-msg" id="msg-${q.id}"></div>
            </div>`;
    });
    cardHtml += `
          </div>`;
    return cardHtml;
  }

  const card1Html = renderQuestionCard(
    `<span>🇷🇺</span> Phần 1: Câu Hỏi & Từ Vựng Tiếng Nga (10 câu)`,
    `10 câu RU`,
    ruQuestions
  );

  const card2Html = renderQuestionCard(
    `<span>🇻🇳</span> Phần 2: Đọc Hiểu & Phân Tích Tiếng Việt (10 câu)`,
    `10 câu VI`,
    viQuestions
  );

  const submitIdSuffix = lesson.id === 1 ? '' : (lesson.id === 2 ? 'L2' : 'L3');
  const scoreValId = `scoreValue${submitIdSuffix}`;
  const scoreStatusId = `scoreStatus${submitIdSuffix}`;
  const scoreDetailId = `scoreDetail${submitIdSuffix}`;
  const btnResetId = `btnReset${submitIdSuffix}`;
  const btnSubmitId = `btnSubmit${submitIdSuffix}`;

  return `
      <!-- MAIN CONTENT: B2 - READING - LESSON ${lesson.id} -->
      <div id="contentB2Reading${lesson.id}" class="test-layout" style="${lesson.id === 1 ? '' : 'display: none;'}">
        <!-- LEFT COLUMN: ĐOẠN VĂN ĐỌC HIỂU & TỪ VỰNG -->
        <section class="reading-panel">
          <div class="panel-header">
            <div style="display: flex; gap: 8px;">
              <span class="badge-tag badge-b2">TRÌNH ĐỘ B2</span>
              <span class="badge-tag badge-reading">${lesson.badge}</span>
            </div>
            <span style="font-size: 13px; color: var(--muted); font-weight: 600;">Thời gian: ${lesson.time}</span>
          </div>

          <h2 class="reading-title">${lesson.title}</h2>

          <div class="reading-meta">
            <span>📰 Nguồn: ${lesson.source}</span>
            <span>✍️ Chủ đề: ${lesson.topic}</span>
          </div>

          <div class="reading-text-card">
            ${pTags}
            <div class="reading-photo-caption">
              ${lesson.caption}
            </div>
          </div>

          <!-- TỪ VỰNG & CẤU TRÚC NGỮ PHÁP TRỌNG TÂM -->
          <div class="vocab-focus-box">
            <div class="vocab-focus-title">
              <span>💡</span> ${lesson.vocabBoxTitle}
            </div>
            <div class="vocab-breakdown">
              ${vocabHtml}
            </div>
          </div>
        </section>

        <!-- RIGHT COLUMN: 20 CÂU HỎI KIỂM TRA -->
        <section class="quiz-panel">
          ${card1Html}
          ${card2Html}

          <!-- SUBMIT & SCORE ACTION BAR -->
          <div class="submit-card">
            <div class="score-display">
              <div class="score-circle" id="${scoreValId}">--</div>
              <div class="score-text">
                <strong id="${scoreStatusId}">Chưa nộp bài</strong>
                <span id="${scoreDetailId}">Hãy hoàn thành tất cả 20 câu hỏi (10 Nga + 10 Việt) để xem kết quả và giải thích chi tiết.</span>
              </div>
            </div>
            <div class="btn-group">
              <button class="btn-action btn-reset" id="${btnResetId}">🔄 Làm lại</button>
              <button class="btn-action btn-submit" id="${btnSubmitId}">🚀 Chấm Điểm Ngay</button>
            </div>
          </div>
        </section>
      </div>`;
}

// Generate explanations map JSON string
function generateExplanationsJs(lessonVarName, lessonObj) {
  const map = {};
  lessonObj.questions.forEach(q => {
    map[q.id] = q.expl;
  });
  return `const ${lessonVarName} = ${JSON.stringify(map, null, 2)};`;
}

const allLessons = [lesson1, lesson2, lesson3];
const allHtmlLayouts = allLessons.map(renderLessonLayout).join('\n\n');

const explanationsJsL1 = generateExplanationsJs('quizExplanationsL1', lesson1);
const explanationsJsL2 = generateExplanationsJs('quizExplanationsL2', lesson2);
const explanationsJsL3 = generateExplanationsJs('quizExplanationsL3', lesson3);

const fullHtml = `<!doctype html>
<html lang="vi">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Luyện Tập & Kiểm Tra Tiếng Nga | TAIGA Center</title>
    <meta
      name="description"
      content="Hệ thống bài test và luyện tập tiếng Nga theo trình độ A1, A2, B1, B2, C1 với 4 kỹ năng Nghe, Nói, Đọc, Viết tại TAIGA Center."
    />
    <link rel="preconnect" href="https://fonts.googleapis.com" />
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
    <link
      href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&family=Space+Grotesk:wght@500;600;700;800&display=swap"
      rel="stylesheet"
    />
    <link rel="stylesheet" href="/test.css" />
  </head>
  <body>
    <div class="shell">
      <!-- HEADER / TOPBAR -->
      <header class="topbar">
        <a href="/home" class="brand-logo">
          <span class="brand-flag">🇷🇺</span>
          <span class="brand-name">TAIGA Center</span>
          <span class="brand-badge">Kiểm Tra</span>
        </a>

        <!-- Hamburger toggle button on mobile -->
        <button class="nav-toggle-btn" id="navToggleBtn" aria-label="Mở menu" aria-expanded="false">
          <span class="hamburger-bar"></span>
          <span class="hamburger-bar"></span>
          <span class="hamburger-bar"></span>
        </button>

        <nav class="nav-links" id="navLinks">
          <!-- <a href="/home" class="nav-item">Trang chủ</a> -->
          <!-- <a href="/roadmap" class="nav-item">Lộ trình học</a> -->
          <!-- <a href="/articles" class="nav-item">Bài viết</a> -->
          <a href="/test" class="nav-item active">Kiểm tra (Test)</a>
          <!-- <a href="/contact" class="nav-item">Liên hệ</a> -->
          <!-- <a href="/app" class="nav-item nav-btn-login">Đăng nhập</a> -->
        </nav>
      </header>

      <!-- HERO SECTION -->
      <section class="test-hero">
        <div class="hero-tag">
          <span>🎯</span>
          <span>HỆ THỐNG LUYỆN TẬP TƯƠNG TÁC 4 KỸ NĂNG CHUẨN QUỐC TẾ</span>
        </div>
        <h1 class="hero-title">
          Kho Bài Test & Luyện Đề <br />
          Chuẩn Quốc Tế <span class="highlight">A1 - A2 - B1 - B2 - C1</span>
        </h1>
        <p class="hero-desc">
          Rèn luyện phản xạ, củng cố ngữ pháp và ghi nhớ từ vựng sâu với 20 câu hỏi tương tác logic (50% tiếng Nga + 50% tiếng Việt) cho mỗi bài đọc hiểu B2.
        </p>
      </section>

      <!-- LEVEL & SKILL SELECTOR -->
      <div class="level-tabs-container">
        <!-- Trình độ: A1 -> C1 -->
        <div class="level-tabs" id="levelTabs">
          <button class="level-btn" data-level="A1">A1 <span class="desc">Sơ cấp</span></button>
          <button class="level-btn" data-level="A2">A2 <span class="desc">Cơ bản</span></button>
          <button class="level-btn" data-level="B1">B1 <span class="desc">Trung cấp 1</span></button>
          <button class="level-btn active" data-level="B2">B2 <span class="desc">Trung cấp 2</span></button>
          <button class="level-btn" data-level="C1">C1 <span class="desc">Cao cấp</span></button>
        </div>

        <!-- 4 Kỹ năng: Nghe, Nói, Đọc, Viết -->
        <div class="skill-tabs" id="skillTabs">
          <button class="skill-btn" data-skill="listening">🎧 Nghe (Аудирование)</button>
          <button class="skill-btn" data-skill="speaking">🗣️ Nói (Говорение)</button>
          <button class="skill-btn active" data-skill="reading">📖 Đọc (Чтение)</button>
          <button class="skill-btn" data-skill="writing">✍️ Viết (Письмо)</button>
        </div>

        <!-- Chọn bài học (Lessons) -->
        <div class="lesson-picker">
          <div style="font-weight: 700; font-size: 14px; color: var(--ink);">
            Danh sách bài kiểm tra (<span id="currentSelectionLabel">B2 - Đọc (Bài 1)</span>):
          </div>
          <div class="lesson-chips" id="lessonChips">
            <button class="lesson-chip active" data-lesson="1">Bài 1: Семья писательницы Т. Устиновой (20 câu)</button>
            <button class="lesson-chip" data-lesson="2">Bài 2: Планирование семьи (20 câu)</button>
            <button class="lesson-chip" data-lesson="3">Bài 3: Пьеса «Чайка» А. П. Чехова и символ МХТ (20 câu)</button>
          </div>
        </div>
      </div>

${allHtmlLayouts}

      <!-- PLACEHOLDER FOR OTHER SELECTIONS -->
      <div id="placeholderEmpty" class="placeholder-card">
        <div class="placeholder-icon">🚧</div>
        <h3 class="placeholder-title" id="placeholderTitle">Đang cập nhật nội dung</h3>
        <p class="placeholder-desc" id="placeholderDesc">
          Nội dung cho cấp độ và kỹ năng này đang được ban chuyên môn hoàn thiện. Bạn hãy quay lại kiểm tra mục <strong>B2 - Đọc</strong> nhé!
        </p>
        <button class="btn-action btn-submit" id="btnBackToB2">
          👉 Đến bài B2 - Đọc (Bài 1)
        </button>
      </div>
    </div>

    <!-- JAVASCRIPT XỬ LÝ TƯƠNG TÁC -->
    <script>
      ${explanationsJsL1}

      ${explanationsJsL2}

      ${explanationsJsL3}

      // Hàm gắn sự kiện click cho các option-item
      function bindOptionEvents() {
        document.querySelectorAll('.option-item').forEach(item => {
          item.onclick = function() {
            const radio = this.querySelector('input[type="radio"]');
            if (radio) {
              radio.checked = true;
              const parent = this.closest('.options-list');
              parent.querySelectorAll('.option-item').forEach(el => el.classList.remove('selected'));
              this.classList.add('selected');
            }
          };
        });
      }
      bindOptionEvents();

      // Submit chấm điểm Bài 1
      document.getElementById('btnSubmit').addEventListener('click', function() {
        evaluateQuiz('#contentB2Reading1', quizExplanationsL1, 'scoreValue', 'scoreStatus', 'scoreDetail', 20);
      });

      // Submit chấm điểm Bài 2
      document.getElementById('btnSubmitL2').addEventListener('click', function() {
        evaluateQuiz('#contentB2Reading2', quizExplanationsL2, 'scoreValueL2', 'scoreStatusL2', 'scoreDetailL2', 20);
      });

      // Submit chấm điểm Bài 3
      document.getElementById('btnSubmitL3').addEventListener('click', function() {
        evaluateQuiz('#contentB2Reading3', quizExplanationsL3, 'scoreValueL3', 'scoreStatusL3', 'scoreDetailL3', 20);
      });

      // Hàm chấm điểm chung
      function evaluateQuiz(containerSelector, explanations, valId, statusId, detailId, totalQuestions) {
        const container = document.querySelector(containerSelector);
        const questionBlocks = container.querySelectorAll('.question-block');
        let score = 0;
        const total = totalQuestions || questionBlocks.length;

        questionBlocks.forEach(block => {
          const qId = block.getAttribute('data-q');
          const correctAns = block.getAttribute('data-ans');
          const selectedRadio = block.querySelector('input[name="' + qId + '"]:checked');
          const msgEl = document.getElementById('msg-' + qId);

          block.querySelectorAll('.option-item').forEach(opt => opt.classList.remove('correct', 'incorrect'));

          if (selectedRadio) {
            const userAns = selectedRadio.value;
            const selectedItem = selectedRadio.closest('.option-item');

            if (userAns === correctAns) {
              score++;
              selectedItem.classList.add('correct');
              msgEl.className = 'feedback-msg show feedback-correct';
              msgEl.innerHTML = '✅ ' + (explanations[qId] && explanations[qId][correctAns] ? explanations[qId][correctAns] : 'Chính xác!');
            } else {
              selectedItem.classList.add('incorrect');
              const correctRadio = block.querySelector('input[name="' + qId + '"][value="' + correctAns + '"]');
              if (correctRadio) {
                correctRadio.closest('.option-item').classList.add('correct');
              }
              msgEl.className = 'feedback-msg show feedback-wrong';
              msgEl.innerHTML = '❌ ' + (explanations[qId] && explanations[qId].other ? explanations[qId].other : 'Chưa chính xác!');
            }
          } else {
            msgEl.className = 'feedback-msg show feedback-wrong';
            msgEl.innerHTML = '⚠️ Bạn chưa trả lời câu này! Đáp án đúng là <strong>' + correctAns + '</strong>.';
            const correctRadio = block.querySelector('input[name="' + qId + '"][value="' + correctAns + '"]');
            if (correctRadio) {
              correctRadio.closest('.option-item').classList.add('correct');
            }
          }
        });

        const scoreCircle = document.getElementById(valId);
        const scoreStatus = document.getElementById(statusId);
        const scoreDetail = document.getElementById(detailId);

        scoreCircle.textContent = score + '/' + total;
        const ratio = score / total;
        if (ratio >= 0.8) {
          scoreStatus.textContent = "Xuất sắc! (Отлично)";
          scoreStatus.style.color = "#16a34a";
          scoreDetail.textContent = 'Bạn đã nắm vững toàn bộ từ vựng, cấu trúc quản trị cách và ý nghĩa bài đọc!';
        } else if (ratio >= 0.5) {
          scoreStatus.textContent = "Khá tốt! (Хорошо)";
          scoreStatus.style.color = "#d97706";
          scoreDetail.textContent = 'Bạn hiểu bài khá tốt, hãy xem lại các câu giải thích ngữ pháp và từ vựng để nhớ lâu hơn nhé.';
        } else {
          scoreStatus.textContent = "Cần ôn tập thêm! (Повторите)";
          scoreStatus.style.color = "#dc2626";
          scoreDetail.textContent = 'Hãy đọc kỹ bảng từ vựng bên trái, sau đó bấm Làm lại nhé!';
        }
      }

      // Reset bài 1
      document.getElementById('btnReset').addEventListener('click', function() {
        resetQuiz('#contentB2Reading1', 'scoreValue', 'scoreStatus', 'scoreDetail', 20);
      });

      // Reset bài 2
      document.getElementById('btnResetL2').addEventListener('click', function() {
        resetQuiz('#contentB2Reading2', 'scoreValueL2', 'scoreStatusL2', 'scoreDetailL2', 20);
      });

      // Reset bài 3
      document.getElementById('btnResetL3').addEventListener('click', function() {
        resetQuiz('#contentB2Reading3', 'scoreValueL3', 'scoreStatusL3', 'scoreDetailL3', 20);
      });

      function resetQuiz(containerSelector, valId, statusId, detailId, totalQuestions) {
        const container = document.querySelector(containerSelector);
        container.querySelectorAll('input[type="radio"]').forEach(r => r.checked = false);
        container.querySelectorAll('.option-item').forEach(opt => opt.classList.remove('selected', 'correct', 'incorrect'));
        container.querySelectorAll('.feedback-msg').forEach(msg => {
          msg.className = 'feedback-msg';
          msg.innerHTML = '';
        });
        document.getElementById(valId).textContent = '--';
        document.getElementById(statusId).textContent = 'Chưa nộp bài';
        document.getElementById(statusId).style.color = '';
        document.getElementById(detailId).textContent = 'Hãy hoàn thành ' + (totalQuestions || 20) + ' câu hỏi để xem kết quả và giải thích chi tiết.';
      }

      // State bộ lọc: Level, Skill, Lesson
      let currentLevel = 'B2';
      let currentSkill = 'reading';
      let currentLesson = '1';

      function updateView() {
        const isB2Reading = (currentLevel === 'B2' && currentSkill === 'reading');
        const contentB2L1 = document.getElementById('contentB2Reading1');
        const contentB2L2 = document.getElementById('contentB2Reading2');
        const contentB2L3 = document.getElementById('contentB2Reading3');
        const placeholder = document.getElementById('placeholderEmpty');
        const labelEl = document.getElementById('currentSelectionLabel');

        const skillNames = {
          listening: 'Nghe',
          speaking: 'Nói',
          reading: 'Đọc',
          writing: 'Viết'
        };

        labelEl.textContent = currentLevel + ' - ' + (skillNames[currentSkill] || currentSkill) + ' (Bài ' + currentLesson + ')';

        if (isB2Reading) {
          placeholder.classList.remove('active');
          contentB2L1.style.display = (currentLesson === '1') ? 'grid' : 'none';
          contentB2L2.style.display = (currentLesson === '2') ? 'grid' : 'none';
          contentB2L3.style.display = (currentLesson === '3') ? 'grid' : 'none';
        } else {
          contentB2L1.style.display = 'none';
          contentB2L2.style.display = 'none';
          contentB2L3.style.display = 'none';
          placeholder.classList.add('active');
          document.getElementById('placeholderTitle').textContent = currentLevel + ' - ' + (skillNames[currentSkill] || currentSkill) + ' (Đang cập nhật)';
          document.getElementById('placeholderDesc').innerHTML = 'Hệ thống đang chuẩn bị các bộ đề cho trình độ <strong>' + currentLevel + '</strong> phần <strong>' + skillNames[currentSkill] + '</strong>. Bạn có thể chọn mục <strong>B2 - Đọc</strong> để luyện tập.';
        }
      }

      // Chuyển bài học (Lesson Chips)
      document.querySelectorAll('#lessonChips .lesson-chip').forEach(chip => {
        chip.addEventListener('click', function() {
          if (this.classList.contains('disabled')) return;
          document.querySelectorAll('#lessonChips .lesson-chip').forEach(c => c.classList.remove('active'));
          this.classList.add('active');
          currentLesson = this.getAttribute('data-lesson');
          updateView();
        });
      });

      // Level buttons click
      document.querySelectorAll('#levelTabs .level-btn').forEach(btn => {
        btn.addEventListener('click', function() {
          document.querySelectorAll('#levelTabs .level-btn').forEach(b => b.classList.remove('active'));
          this.classList.add('active');
          currentLevel = this.getAttribute('data-level');
          updateView();
        });
      });

      // Skill buttons click
      document.querySelectorAll('#skillTabs .skill-btn').forEach(btn => {
        btn.addEventListener('click', function() {
          document.querySelectorAll('#skillTabs .skill-btn').forEach(b => b.classList.remove('active'));
          this.classList.add('active');
          currentSkill = this.getAttribute('data-skill');
          updateView();
        });
      });

      // Back to B2 button click
      document.getElementById('btnBackToB2').addEventListener('click', function() {
        document.querySelectorAll('#levelTabs .level-btn').forEach(b => {
          b.classList.toggle('active', b.getAttribute('data-level') === 'B2');
        });
        document.querySelectorAll('#skillTabs .skill-btn').forEach(b => {
          b.classList.toggle('active', b.getAttribute('data-skill') === 'reading');
        });
        currentLevel = 'B2';
        currentSkill = 'reading';
        currentLesson = '1';
        updateView();
      });

      // Mobile Hamburger Menu Toggle
      const navToggleBtn = document.getElementById('navToggleBtn');
      const navLinks = document.getElementById('navLinks');
      if (navToggleBtn && navLinks) {
        navToggleBtn.addEventListener('click', function() {
          const isOpen = navLinks.classList.toggle('open');
          navToggleBtn.classList.toggle('active', isOpen);
          navToggleBtn.setAttribute('aria-expanded', String(isOpen));
        });

        navLinks.querySelectorAll('.nav-item').forEach(link => {
          link.addEventListener('click', () => {
            navLinks.classList.remove('open');
            navToggleBtn.classList.remove('active');
            navToggleBtn.setAttribute('aria-expanded', 'false');
          });
        });
      }
    </script>
  </body>
</html>
`;

fs.writeFileSync(path.join(__dirname, '../public/test.html'), fullHtml, 'utf8');
console.log('Successfully generated public/test.html with 20 questions each (total 60 questions)!');
