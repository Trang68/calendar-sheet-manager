const fs = require('fs');
const path = require('path');

const outputDir = path.join(__dirname, '../public/data/listening/a1');
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

const lessons = [
  {
    id: 1,
    title: "В кафе: Заказ еды и напитков",
    listeningTime: "10 phút",
    topic: "Hội thoại thực tế: Gọi đồ uống & món ăn tại quán cà phê",
    dialogue: [
      { speaker: "Официант", text: "Добрый день! Добро пожаловать в наше уютное кафе." },
      { speaker: "Анна", text: "Здравствуйте! Покажите, пожалуйста, меню." },
      { speaker: "Официант", text: "Вот меню. Что вы будете заказывать сегодня?" },
      { speaker: "Анна", text: "Я буду зелёный чай с лимоном и яблочный пирог." },
      { speaker: "Максим", text: "А мне, пожалуйста, чашку чёрного кофе без сахара." },
      { speaker: "Официант", text: "Хорошо. Кофе без сахара, зелёный чай и пирог. Что-нибудь ещё?" },
      { speaker: "Максим", text: "Нет, спасибо, это всё. Скажите, пожалуйста, сколько это стоит?" },
      { speaker: "Официант", text: "С вас триста пятьдесят рублей." },
      { speaker: "Анна", text: "Вот деньги. Большое спасибо!" },
      { speaker: "Официант", text: "Спасибо вам! Приятного аппетита!" }
    ],
    dictation: {
      title: "Chép chính tả (Диктант)",
      instruction: "Nghe kỹ câu sau và gõ lại chính xác bằng tiếng Nga vào ô bên dưới:",
      targetText: "Дайте, пожалуйста, зелёный чай и чёрный кофе.",
      hint: "Gợi ý: Chú ý dấu phẩy sau «пожалуйста» và đuôi tính từ «зелёный», «чёрный»."
    },
    questions: [
      {
        id: "l1_q1", lang: "ru",
        prompt: "Где происходит этот диалог?",
        hint: "Xác định địa điểm đối thoại.",
        options: [{key:"A", text:"В библиотеке."}, {key:"B", text:"В кафе."}, {key:"C", text:"В больнице."}, {key:"D", text:"В метро."}],
        ans: "B", explain: "Bồi bàn chào: «Добро пожаловать в наше уютное кафе» (Chào mừng bạn đến quán cà phê)."
      },
      {
        id: "l1_q2", lang: "ru",
        prompt: "Что заказала Анна?",
        hint: "Món của Anna.",
        options: [{key:"A", text:"Чёрный кофе."}, {key:"B", text:"Зелёный чай с лимоном и пирог."}, {key:"C", text:"Воду и мороженое."}, {key:"D", text:"Сок и суп."}],
        ans: "B", explain: "Anna nói: «Я буду зелёный чай с лимоном и яблочный пирог»."
      },
      {
        id: "l1_q3", lang: "ru",
        prompt: "Какой кофе заказал Максим?",
        hint: "Maxim uống cà phê như thế nào?",
        options: [{key:"A", text:"С молоком и сахаром."}, {key:"B", text:"Холодный кофе."}, {key:"C", text:"Чёрный кофе без сахара."}, {key:"D", text:"С шоколадом."}],
        ans: "C", explain: "Maxim yêu cầu: «чашку чёрного кофе без сахара»."
      },
      {
        id: "l1_q4", lang: "ru",
        prompt: "Сколько стоит весь заказ?",
        hint: "Số tiền được thanh toán.",
        options: [{key:"A", text:"250 рублей."}, {key:"B", text:"300 рублей."}, {key:"C", text:"350 рублей."}, {key:"D", text:"500 рублей."}],
        ans: "C", explain: "Официант: «С вас триста пятьдесят (350) рублей»."
      },
      {
        id: "l1_q5", lang: "vi",
        prompt: "Hai bạn Anna và Maxim gọi mấy loại đồ uống?",
        hint: "Đếm số loại thức uống.",
        options: [{key:"A", text:"1 loại đồ uống."}, {key:"B", text:"2 loại (trà xanh và cà phê đen)."}, {key:"C", text:"3 loại."}, {key:"D", text:"Không gọi đồ uống nào."}],
        ans: "B", explain: "Anna gọi trà xanh chanh, Maxim gọi cà phê đen."
      },
      {
        id: "l1_q6", lang: "vi",
        prompt: "Ý nghĩa của câu chúc «Приятного аппетита!» là gì?",
        hint: "Lời chúc phổ biến trên bàn ăn.",
        options: [{key:"A", text:"Chúc ngủ ngon!"}, {key:"B", text:"Chúc ngon miệng!"}, {key:"C", text:"Chúc một ngày tốt lành!"}, {key:"D", text:"Hẹn gặp lại!"}],
        ans: "B", explain: "«Приятного аппетита!» có nghĩa là chúc ngon miệng."
      },
      {
        id: "l1_q7", lang: "vi",
        prompt: "Cụm từ «без сахара» trong lời gọi của Maxim có nghĩa là gì?",
        hint: "Giới từ «без».",
        options: [{key:"A", text:"Nhiều đường."}, {key:"B", text:"Không đường."}, {key:"C", text:"Đường sữa."}, {key:"D", text:"Đường ăn kiêng."}],
        ans: "B", explain: "«без» nghĩa là không có: «без сахара» = không có đường."
      },
      {
        id: "l1_q8", lang: "vi",
        prompt: "Anna đã yêu cầu nhân viên đưa vật gì trước khi gọi món?",
        hint: "Vật dụng xem món ăn.",
        options: [{key:"A", text:"Hóa đơn thanh toán."}, {key:"B", text:"Thực đơn (меню)."}, {key:"C", text:"Khăn giấy."}, {key:"D", text:"Cốc nước lọc."}],
        ans: "B", explain: "Anna nói: «Покажите, пожалуйста, меню»."
      }
    ]
  },
  {
    id: 2,
    title: "Давайте познакомимся: Как вас зовут?",
    listeningTime: "10 phút",
    topic: "Làm quen, hỏi tên tuổi, quốc tịch & ngành nghề",
    dialogue: [
      { speaker: "Антон", text: "Здравствуйте! Меня зовут Антон. А как вас зовут?" },
      { speaker: "Мария", text: "Очень приятно, Антон! Меня зовут Мария." },
      { speaker: "Антон", text: "Мария, вы студентка?" },
      { speaker: "Мария", text: "Да, я студентка первого курса. Я изучаю русскую филологию. А вы?" },
      { speaker: "Антон", text: "Я тоже студент, но изучаю экономику и бизнес." },
      { speaker: "Мария", text: "Откуда вы приехали, Антон?" },
      { speaker: "Антон", text: "Я из Санкт-Петербурга. А вы откуда?" },
      { speaker: "Мария", text: "Я приехала из Казани. Москва — очень красивый город!" },
      { speaker: "Антон", text: "Да, согласен. Рад знакомству, Мария!" },
      { speaker: "Мария", text: "Я тоже очень рада!" }
    ],
    dictation: {
      title: "Chép chính tả (Диктант)",
      instruction: "Nghe kỹ câu sau và gõ lại chính xác bằng tiếng Nga:",
      targetText: "Меня зовут Мария, я студентка первого курса.",
      hint: "Gợi ý: Chú ý viết hoa tên riêng «Мария» và đuôi «первого курса»."
    },
    questions: [
      {
        id: "l2_q1", lang: "ru",
        prompt: "Как зовут девушку в диалоге?",
        hint: "Tên bạn nữ.",
        options: [{key:"A", text:"Анна."}, {key:"B", text:"Мария."}, {key:"C", text:"Елена."}, {key:"D", text:"Ольга."}],
        ans: "B", explain: "Cô ấy nói: «Меня зовут Мария»."
      },
      {
        id: "l2_q2", lang: "ru",
        prompt: "Что изучает Антон?",
        hint: "Ngành học của Anton.",
        options: [{key:"A", text:"Филологию."}, {key:"B", text:"Экономику и бизнес."}, {key:"C", text:"Медицину."}, {key:"D", text:"Историю."}],
        ans: "B", explain: "Anton nói: «Я тоже студент, но изучаю экономику и бизнес»."
      },
      {
        id: "l2_q3", lang: "ru",
        prompt: "Откуда приехал Антон?",
        hint: "Quê hương của Anton.",
        options: [{key:"A", text:"Из Москвы."}, {key:"B", text:"Из Санкт-Петербурга."}, {key:"C", text:"Из Казани."}, {key:"D", text:"Из Владивостока."}],
        ans: "B", explain: "Anton: «Я из Санкт-Петербурга»."
      },
      {
        id: "l2_q4", lang: "ru",
        prompt: "Откуда приехала Мария?",
        hint: "Thành phố của Maria.",
        options: [{key:"A", text:"Из Казани."}, {key:"B", text:"Из Минска."}, {key:"C", text:"Из Самары."}, {key:"D", text:"Из Сочи."}],
        ans: "A", explain: "Maria: «Я приехала из Казани»."
      },
      {
        id: "l2_q5", lang: "vi",
        prompt: "Cả Anton và Maria hiện nay đang làm gì?",
        hint: "Tình trạng học tập/làm việc của hai người.",
        options: [{key:"A", text:"Cả hai đều đã đi làm bác sĩ."}, {key:"B", text:"Cả hai đều là sinh viên đại học."}, {key:"C", text:"Một người là giáo viên, một người là sinh viên."}, {key:"D", text:"Họ là học sinh tiểu học."}],
        ans: "B", explain: "Maria là sinh viên năm nhất ngữ văn, Anton là sinh viên kinh tế."
      },
      {
        id: "l2_q6", lang: "vi",
        prompt: "Câu nói «Рад знакомству!» biểu thị điều gì?",
        hint: "Thái độ trong giao tiếp.",
        options: [{key:"A", text:"Rất vui được làm quen."}, {key:"B", text:"Tạm biệt nhé."}, {key:"C", text:"Xin thứ lỗi."}, {key:"D", text:"Cảm ơn nhiều."}],
        ans: "A", explain: "«Рад знакомству» = Rất vui được làm quen."
      },
      {
        id: "l2_q7", lang: "vi",
        prompt: "Hiện tại cuộc đối thoại này đang diễn ra ở thành phố nào?",
        hint: "Hai bạn đang khen thành phố nào đẹp?",
        options: [{key:"A", text:"Hà Nội."}, {key:"B", text:"Moskva."}, {key:"C", text:"Kazan."}, {key:"D", text:"Peterburg."}],
        ans: "B", explain: "Maria nhận xét: «Москва — очень красивый город!»."
      },
      {
        id: "l2_q8", lang: "vi",
        prompt: "Maria là sinh viên năm mấy?",
        hint: "Nghe từ «первого курса».",
        options: [{key:"A", text:"Năm nhất."}, {key:"B", text:"Năm hai."}, {key:"C", text:"Năm ba."}, {key:"D", text:"Năm cuối."}],
        ans: "A", explain: "«первого курса» = năm thứ nhất."
      }
    ]
  },
  {
    id: 3,
    title: "В библиотеке: Где лежит книга?",
    listeningTime: "10 phút",
    topic: "Hỏi vị trí đồ dùng & mượn sách trong thư viện",
    dialogue: [
      { speaker: "Студент", text: "Здравствуйте! Скажите, пожалуйста, где словарь?" },
      { speaker: "Библиотекарь", text: "Добрый день! Какой словарь вам нужен: русско-вьетнамский или русско-английский?" },
      { speaker: "Студент", text: "Мне нужен большой русско-английский словарь." },
      { speaker: "Библиотекарь", text: "Посмотрите, он лежит вон там, на полке справа." },
      { speaker: "Студент", text: "А где здесь свежие газеты и журналы?" },
      { speaker: "Библиотекарь", text: "Журналы лежат на большом круглом столе в центре зала." },
      { speaker: "Студент", text: "Можно взять этот учебник домой?" },
      { speaker: "Библиотекарь", text: "Да, конечно. Покажите ваш студенческий билет." },
      { speaker: "Студент", text: "Вот мой билет. Спасибо вам огромное!" },
      { speaker: "Библиотекарь", text: "Пожалуйста. Не забудьте вернуть через две недели." }
    ],
    dictation: {
      title: "Chép chính tả (Диктант)",
      instruction: "Nghe và viết lại câu sau chính xác:",
      targetText: "Большой словарь лежит на полке справа.",
      hint: "Gợi ý: Chú ý dấu mềm trong «словарь» và giới từ «на полке справа»."
    },
    questions: [
      {
        id: "l3_q1", lang: "ru",
        prompt: "Какой словарь нужен студенту?",
        hint: "Loại từ điển sinh viên cần.",
        options: [{key:"A", text:"Русско-французский."}, {key:"B", text:"Большой русско-английский."}, {key:"C", text:"Русско-вьетнамский."}, {key:"D", text:"Немецкий словарь."}],
        ans: "B", explain: "Студент: «Мне нужен большой русско-английский словарь»."
      },
      {
        id: "l3_q2", lang: "ru",
        prompt: "Где лежит нужный словарь?",
        hint: "Vị trí của cuốn từ điển.",
        options: [{key:"A", text:"На полке слева."}, {key:"B", text:"На полке справа."}, {key:"C", text:"В шкафу у окна."}, {key:"D", text:"Под столом."}],
        ans: "B", explain: "Библиотекарь: «он лежит вон там, на полке справа»."
      },
      {
        id: "l3_q3", lang: "ru",
        prompt: "Где лежат журналы?",
        hint: "Vị trí báo và tạp chí.",
        options: [{key:"A", text:"На круглом столе в центре зала."}, {key:"B", text:"У входа."}, {key:"C", text:"На полке слева."}, {key:"D", text:"В кабинете директора."}],
        ans: "A", explain: "Библиотекарь: «Журналы лежат на большом круглом столе в центре зала»."
      },
      {
        id: "l3_q4", lang: "ru",
        prompt: "Какой документ попросил библиотекарь?",
        hint: "Giấy tờ cần xuất trình.",
        options: [{key:"A", text:"Паспорт."}, {key:"B", text:"Студенческий билет."}, {key:"C", text:"Визу."}, {key:"D", text:"Водительские права."}],
        ans: "B", explain: "Библиотекарь: «Покажите ваш студенческий билет»."
      },
      {
        id: "l3_q5", lang: "vi",
        prompt: "Sinh viên được phép mượn giáo trình về nhà trong bao lâu?",
        hint: "Nghe cụm từ «через две недели».",
        options: [{key:"A", text:"3 ngày."}, {key:"B", text:"1 tuần."}, {key:"C", text:"2 tuần (hai tuần lễ)."}, {key:"D", text:"1 tháng."}],
        ans: "C", explain: "«через две недели» = sau hai tuần."
      },
      {
        id: "l3_q6", lang: "vi",
        prompt: "Từ «справа» có ý nghĩa chỉ phương hướng gì?",
        hint: "Phương hướng không gian.",
        options: [{key:"A", text:"Bên trái."}, {key:"B", text:"Bên phải."}, {key:"C", text:"Ở giữa."}, {key:"D", text:"Đằng sau."}],
        ans: "B", explain: "«справа» = ở bên phải."
      },
      {
        id: "l3_q7", lang: "vi",
        prompt: "Chiếc bàn để tạp chí có hình dạng gì?",
        hint: "Tính từ «круглом».",
        options: [{key:"A", text:"Bàn vuông."}, {key:"B", text:"Bàn chữ nhật dài."}, {key:"C", text:"Bàn tròn lớn."}, {key:"D", text:"Bàn bầu dục."}],
        ans: "C", explain: "«на большом круглом столе» = trên chiếc bàn tròn lớn."
      },
      {
        id: "l3_q8", lang: "vi",
        prompt: "Sinh viên có được phép mang sách về nhà không?",
        hint: "Câu trả lời của thủ thư.",
        options: [{key:"A", text:"Không được mang về."}, {key:"B", text:"Được, sau khi trình thẻ sinh viên."}, {key:"C", text:"Phải trả thêm tiền phí."}, {key:"D", text:"Chỉ được đọc tại chỗ."}],
        ans: "B", explain: "Thủ thư nói «Да, конечно. Покажите ваш студенческий билет»."
      }
    ]
  },
  {
    id: 4,
    title: "В магазине одежды: Какой размер?",
    listeningTime: "11 phút",
    topic: "Mua sắm quần áo, thử size và màu sắc",
    dialogue: [
      { speaker: "Продавец", text: "Здравствуйте! Чем я могу вам помочь?" },
      { speaker: "Покупатель", text: "Здравствуйте! Я ищу тёплый зимний свитер." },
      { speaker: "Продавец", text: "Какой цвет вы предпочитаете?" },
      { speaker: "Покупатель", text: "Мне нравится синий или тёмно-серый." },
      { speaker: "Продавец", text: "Какой у вас размер одежды?" },
      { speaker: "Покупатель", text: "Размер М, сорок восьмой." },
      { speaker: "Продавец", text: "Вот отличный синий свитер из натуральной шерсти. Примерьте его." },
      { speaker: "Покупатель", text: "Где у вас примерочная кабина?" },
      { speaker: "Продавец", text: "Примерочные кабины находятся прямо и налево." },
      { speaker: "Покупатель", text: "Свитер подошёл идеально! Я его покупаю." }
    ],
    dictation: {
      title: "Chép chính tả (Диктант)",
      instruction: "Nghe và gõ lại chính xác câu sau:",
      targetText: "Примерочные кабины находятся прямо и налево.",
      hint: "Gợi ý: Chú ý chính tả «находятся» và phương hướng «прямо и налево»."
    },
    questions: [
      {
        id: "l4_q1", lang: "ru",
        prompt: "Что ищет покупатель?",
        hint: "Món đồ khách cần mua.",
        options: [{key:"A", text:"Летнюю рубашку."}, {key:"B", text:"Тёплый зимний свитер."}, {key:"C", text:"Кожаную куртку."}, {key:"D", text:"Чёрные брюки."}],
        ans: "B", explain: "Покупатель: «Я ищу тёплый зимний свитер»."
      },
      {
        id: "l4_q2", lang: "ru",
        prompt: "Какой размер одежды у покупателя?",
        hint: "Kích cỡ áo.",
        options: [{key:"A", text:"Размер S (44)."}, {key:"B", text:"Размер M (48)."}, {key:"C", text:"Размер L (52)."}, {key:"D", text:"Размер XL (56)."}],
        ans: "B", explain: "Покупатель: «Размер М, сорок восьмой»."
      },
      {
        id: "l4_q3", lang: "ru",
        prompt: "Из какого материала сделан свитер?",
        hint: "Chất liệu áo len.",
        options: [{key:"A", text:"Из хлопка."}, {key:"B", text:"Из натуральной шерсти."}, {key:"C", text:"Из шёлка."}, {key:"D", text:"Из синтетики."}],
        ans: "B", explain: "Продавец: «свитер из натуральной шерсти» (len tự nhiên)."
      },
      {
        id: "l4_q4", lang: "ru",
        prompt: "Где находятся примерочные кабины?",
        hint: "Vị trí phòng thử đồ.",
        options: [{key:"A", text:"Справа у кассы."}, {key:"B", text:"Прямо и налево."}, {key:"C", text:"На втором этаже."}, {key:"D", text:"Рядом со входом."}],
        ans: "B", explain: "Продавец: «Примерочные кабины находятся прямо и налево»."
      },
      {
        id: "l4_q5", lang: "vi",
        prompt: "Người mua thích những gam màu nào cho áo len?",
        hint: "Nghe màu sắc khách hàng yêu cầu.",
        options: [{key:"A", text:"Đỏ hoặc vàng."}, {key:"B", text:"Xanh dương hoặc xám sẫm."}, {key:"C", text:"Trắng tinh khôi."}, {key:"D", text:"Xanh lá cây rực rỡ."}],
        ans: "B", explain: "«синий или тёмно-серый» = xanh dương hoặc xám sẫm."
      },
      {
        id: "l4_q6", lang: "vi",
        prompt: "Từ «Примерочная» trong hội thoại có nghĩa là gì?",
        hint: "Căn phòng thử quần áo.",
        options: [{key:"A", text:"Quầy thanh toán."}, {key:"B", text:"Phòng thử đồ."}, {key:"C", text:"Kho hàng."}, {key:"D", text:"Cửa thoát hiểm."}],
        ans: "B", explain: "«Примерочная» là phòng thử đồ trong cửa hàng quần áo."
      },
      {
        id: "l4_q7", lang: "vi",
        prompt: "Kết quả sau khi người mua mặc thử chiếc áo là gì?",
        hint: "Lời nhận xét của khách.",
        options: [{key:"A", text:"Áo quá chật."}, {key:"B", text:"Áo quá dài."}, {key:"C", text:"Áo vừa vặn hoàn hảo và khách quyết định mua."}, {key:"D", text:"Khách chê đắt và đi về."}],
        ans: "C", explain: "Khách khen: «Свитер подошёл идеально! Я его покупаю»."
      },
      {
        id: "l4_q8", lang: "vi",
        prompt: "Từ «налево» mang nghĩa chỉ hướng nào?",
        hint: "Phương hướng.",
        options: [{key:"A", text:"Rẽ phải."}, {key:"B", text:"Sang bên trái."}, {key:"C", text:"Đi thẳng."}, {key:"D", text:"Quay lại."}],
        ans: "B", explain: "«налево» = sang bên trái."
      }
    ]
  },
  {
    id: 5,
    title: "В аптеке: Лекарства и здоровье",
    listeningTime: "10 phút",
    topic: "Mua thuốc cảm sốt & hỏi thăm tình trạng sức khỏe",
    dialogue: [
      { speaker: "Фармацевт", text: "Здравствуйте! Что у вас случилось?" },
      { speaker: "Пациент", text: "Здравствуйте! У меня сильно болит голова и болит горло." },
      { speaker: "Фармацевт", text: "У вас есть температура?" },
      { speaker: "Пациент", text: "Да, утром была тридцать семь и пять (37,5)." },
      { speaker: "Фармацевт", text: "Понятно, это простуда. Вот эффективные таблетки от боли и спрей для горла." },
      { speaker: "Пациент", text: "Как нужно принимать эти таблетки?" },
      { speaker: "Фармацевт", text: "Принимайте по одной таблетке два раза в день после еды." },
      { speaker: "Пациент", text: "Спасибо! И дайте ещё витамин С, пожалуйста." },
      { speaker: "Фармацевт", text: "Вот витамин С. Пейте много горячего чая и выздоравливайте!" },
      { speaker: "Пациент", text: "Большое вам спасибо за помощь!" }
    ],
    dictation: {
      title: "Chép chính tả (Диктант)",
      instruction: "Nghe kỹ liều lượng uống thuốc và gõ lại:",
      targetText: "Принимайте по одной таблетке два раза в день после еды.",
      hint: "Gợi ý: Chú ý giới từ «после еды» (sau khi ăn) và «два раза в день»."
    },
    questions: [
      {
        id: "l5_q1", lang: "ru",
        prompt: "Что болит у пациента?",
        hint: "Triệu chứng của bệnh nhân.",
        options: [{key:"A", text:"Болит нога и рука."}, {key:"B", text:"Болит голова и горло."}, {key:"C", text:"Болит живот."}, {key:"D", text:"Болит зуб."}],
        ans: "B", explain: "Пациент: «У меня сильно болит голова и болит горло»."
      },
      {
        id: "l5_q2", lang: "ru",
        prompt: "Какая температура была у пациента утром?",
        hint: "Nhiệt độ cơ thể.",
        options: [{key:"A", text:"36,6."}, {key:"B", text:"37,5."}, {key:"C", text:"38,5."}, {key:"D", text:"39,0."}],
        ans: "B", explain: "Пациент: «утром была тридцать семь и пять (37,5)»."
      },
      {
        id: "l5_q3", lang: "ru",
        prompt: "Как нужно принимать таблетки?",
        hint: "Liều lượng uống.",
        options: [{key:"A", text:"Три раза в день до еды."}, {key:"B", text:"По одной таблетке два раза в день после еды."}, {key:"C", text:"Только на ночь."}, {key:"D", text:"Один раз в неделю."}],
        ans: "B", explain: "«Принимайте по одной таблетке два раза в день после еды»."
      },
      {
        id: "l5_q4", lang: "ru",
        prompt: "Что ещё купил пациент?",
        hint: "Món bổ sung bệnh nhân mua.",
        options: [{key:"A", text:"Антибиотики."}, {key:"B", text:"Витамин С."}, {key:"C", text:"Бинт и пластырь."}, {key:"D", text:"Капли для глаз."}],
        ans: "B", explain: "Пациент: «И дайте ещё витамин С, пожалуйста»."
      },
      {
        id: "l5_q5", lang: "vi",
        prompt: "Dược sĩ khuyên bệnh nhân nên uống gì để mau khỏi?",
        hint: "Lời khuyên sinh hoạt của dược sĩ.",
        options: [{key:"A", text:"Nước ngọt có gas."}, {key:"B", text:"Nhiều trà nóng."}, {key:"C", text:"Nước đá lạnh."}, {key:"D", text:"Cà phê đậm đặc."}],
        ans: "B", explain: "Фармацевт: «Пейте много горячего чая и выздоравливайте!»"
      },
      {
        id: "l5_q6", lang: "vi",
        prompt: "Cụm từ «после еды» chỉ thời điểm uống thuốc nào?",
        hint: "Trước hay sau ăn?",
        options: [{key:"A", text:"Trước bữa ăn."}, {key:"B", text:"Sau khi ăn."}, {key:"C", text:"Trong khi ăn."}, {key:"D", text:"Uống khi bụng đói."}],
        ans: "B", explain: "«после еды» = sau bữa ăn."
      },
      {
        id: "l5_q7", lang: "vi",
        prompt: "Lời chúc «Выздоравливайте!» mang thông điệp gì?",
        hint: "Chúc bệnh nhân.",
        options: [{key:"A", text:"Chúc mau bình phục / mau khỏe lại!"}, {key:"B", text:"Chúc may mắn!"}, {key:"C", text:"Chúc chuyến đi tốt đẹp!"}, {key:"D", text:"Chúc mừng sinh nhật!"}],
        ans: "A", explain: "«Выздоравливайте!» = Chúc mau chóng khỏi bệnh."
      },
      {
        id: "l5_q8", lang: "vi",
        prompt: "Căn bệnh mà người bệnh mắc phải là gì?",
        hint: "Từ «простуда».",
        options: [{key:"A", text:"Bong gân chân."}, {key:"B", text:"Cảm lạnh thông thường."}, {key:"C", text:"Đau dạ dày mạn tính."}, {key:"D", text:"Dị ứng phấn hoa."}],
        ans: "B", explain: "«простуда» nghĩa là cảm lạnh thông thường."
      }
    ]
  }
];

// Sinh tiếp từ bài 6 đến bài 20
const additionalLessonsData = [
  { id: 6, title: "Мой рабочий день: С утра до вечера", time: "11 phút", topic: "Kể về lịch trình biểu hàng ngày của một chuyên viên IT",
    dialogue: [
      { speaker: "Иван", text: "Алексей, во сколько ты обычно встаёшь утром?" },
      { speaker: "Алексей", text: "Я встаю рано, в семь часов утра. Делаю зарядку и пью кофе." },
      { speaker: "Иван", text: "Когда начинается твой рабочий день?" },
      { speaker: "Алексей", text: "Работа в офисе начинается ровно в девять часов." },
      { speaker: "Иван", text: "А где ты обедаешь?" },
      { speaker: "Алексей", text: "В час дня я обедаю в нашей столовой с коллегами." },
      { speaker: "Иван", text: "Во сколько ты возвращаешься домой?" },
      { speaker: "Алексей", text: "В шесть вечера я заканчиваю работу и еду домой на метро." },
      { speaker: "Иван", text: "Что ты делаешь вечером?" },
      { speaker: "Алексей", text: "Вечером я отдыхаю, читаю книги или смотрю фильмы." }
    ],
    dictation: "Я встаю в семь часов утра и делаю утреннюю зарядку.",
    dictationHint: "Chú ý số từ «в семь часов» và tính từ «утреннюю».",
    ruAnswers: ["В 7:00 утра.", "В 9:00 часов.", "В столовой.", "На метро."],
    viAnswers: ["Sinh hoạt hàng ngày và công việc.", "Lúc 13:00 (1 giờ trưa).", "Đọc sách hoặc xem phim.", "Lúc 18:00 (6 giờ tối)."]
  },
  { id: 7, title: "На вокзале: Покупка билета на поезд", time: "12 phút", topic: "Mua vé tàu hỏa đi Saint Petersburg tại nhà ga",
    dialogue: [
      { speaker: "Кассир", text: "Здравствуйте! Куда вам нужен билет?" },
      { speaker: "Пассажир", text: "Здравствуйте! Мне нужен один билет до Санкт-Петербурга." },
      { speaker: "Кассир", text: "На какую дату и на какое время?" },
      { speaker: "Пассажир", text: "На завтра, на скоростной поезд «Сапсан» утром." },
      { speaker: "Кассир", text: "Есть удобный рейс в семь часов утра. Место у окна." },
      { speaker: "Пассажир", text: "Прекрасно! Сколько стоит этот билет?" },
      { speaker: "Кассир", text: "Три тысячи двести рублей. Оплата картой или наличными?" },
      { speaker: "Пассажир", text: "Банковской картой, пожалуйста." },
      { speaker: "Кассир", text: "Вот ваш билет и чек. Поезд отправляется с третьего пути." },
      { speaker: "Пассажир", text: "Большое спасибо! Счастливого пути мне!" }
    ],
    dictation: "Поезд Сапсан отправляется в семь часов утра с третьего пути.",
    dictationHint: "Tên tàu «Сапсан» và số thứ tự «с третьего пути».",
    ruAnswers: ["До Санкт-Петербурга.", "На поезд Сапсан.", "3200 рублей.", "С третьего пути."],
    viAnswers: ["Tàu cao tốc chạy buổi sáng.", "Thanh toán bằng thẻ ngân hàng.", "Vị trí ghế ngồi cạnh cửa sổ.", "Ngày mai."]
  },
  { id: 8, title: "В гостинице: Регистрация и номер", time: "10 phút", topic: "Thủ tục nhận phòng tại khách sạn (Check-in)",
    dialogue: [
      { speaker: "Администратор", text: "Добрый вечер! Добро пожаловать в отель «Космос»." },
      { speaker: "Турист", text: "Здравствуйте! У меня забронирован одноместный номер на три дня." },
      { speaker: "Администратор", text: "Назовите, пожалуйста, вашу фамилию." },
      { speaker: "Турист", text: "Моя фамилия Нгуен, имя Ван." },
      { speaker: "Администратор", text: "Да, бронь есть. Ваш номер четыреста двенадцать на четвёртом этаже." },
      { speaker: "Турист", text: "Скажите, завтрак включён в стоимость?" },
      { speaker: "Администратор", text: "Да, завтрак — шведский стол с семи до десяти утра в ресторане." },
      { speaker: "Турист", text: "А какой пароль от беспроводного интернета Wi-Fi?" },
      { speaker: "Администратор", text: "Пароль написан на карточке вашего электронного ключа." },
      { speaker: "Турист", text: "Отлично, благодарю вас!" }
    ],
    dictation: "Завтрак проходит с семи до десяти утра в ресторане.",
    dictationHint: "Số giờ «с семи до десяти» và từ «в ресторане».",
    ruAnswers: ["На 3 дня.", "Номер 412.", "На четвёртом этаже.", "С 7 до 10 утра."],
    viAnswers: ["Khách sạn Kosmos.", "Bữa sáng buffet đã bao gồm.", "In trên thẻ khóa điện tử.", "Phòng đơn 1 người."]
  },
  { id: 9, title: "Семья и родственники: Кто кем работает?", time: "11 phút", topic: "Trò chuyện về các thành viên gia đình và nghề nghiệp",
    dialogue: [
      { speaker: "Ольга", text: "Дмитрий, у тебя большая семья?" },
      { speaker: "Дмитрий", text: "Да, у меня довольно большая семья: родители, сестра и бабушка." },
      { speaker: "Ольга", text: "Кто по профессии твои родители?" },
      { speaker: "Дмитрий", text: "Мой папа — инженер-строитель, а мама — школьная учительница." },
      { speaker: "Ольга", text: "А чем занимается твоя младшая сестра?" },
      { speaker: "Дмитрий", text: "Она ещё маленькая, учится в пятом классе и любит рисовать." },
      { speaker: "Ольга", text: "А твоя бабушка уже на пенсии?" },
      { speaker: "Дмитрий", text: "Да, бабушка на пенсии. Она печёт самые вкусные пирожки!" },
      { speaker: "Ольга", text: "Как замечательно иметь такую тёплую семью!" },
      { speaker: "Дмитрий", text: "Спасибо! Мы очень дружные." }
    ],
    dictation: "Мой папа — инженер, а мама — школьная учительница.",
    dictationHint: "Dấu gạch ngang sau chủ ngữ «папа — ...» và «мама — ...».",
    ruAnswers: ["Инженер-строитель.", "Школьная учительница.", "В пятом классе.", "Она на пенсии."],
    viAnswers: ["Bố, mẹ, em gái và bà.", "Bé rất thích vẽ tranh.", "Bà làm bánh pỉ-rốc rất ngon.", "Gia đình đầm ấm và hòa thuận."]
  },
  { id: 10, title: "В аэропорту: Посадка на рейс", time: "12 phút", topic: "Nghe thông báo chuyến bay và làm thủ tục hải quan",
    dialogue: [
      { speaker: "Диктор", text: "Внимание пассажиров! Начинается посадка на рейс двести пять в Ханой." },
      { speaker: "Пассажир", text: "Извините, девушка, к какому выходу мне нужно пройти?" },
      { speaker: "Сотрудник", text: "Покажите посадочный талон. Ваш выход номер двенадцать." },
      { speaker: "Пассажир", text: "А сколько килограммов ручной клади разрешено взять с собой?" },
      { speaker: "Сотрудник", text: "В салон самолёта разрешается ручная кладь весом до десяти килограммов." },
      { speaker: "Пассажир", text: "Хорошо, у меня только небольшой рюкзак." },
      { speaker: "Сотрудник", text: "Приготовьте ваш заграничный паспорт и талон. Приятного полёта!" },
      { speaker: "Пассажир", text: "Спасибо большое за информацию!" }
    ],
    dictation: "Посадка на рейс в Ханой начинается у двенадцатого выхода.",
    dictationHint: "Tên thành phố «в Ханой» và số cổng «у двенадцатого выхода».",
    ruAnswers: ["В Ханой.", "Номер 12.", "До 10 килограммов.", "Заграничный паспорт и талон."],
    viAnswers: ["Thông báo hành khách lên máy bay.", "Chỉ có một chiếc ba lô nhỏ.", "Cửa số 12.", "Chúc chuyến bay tốt lành."]
  },
  { id: 11, title: "Погода на выходные: Прогноз синоптиков", time: "10 phút", topic: "Nghe bản tin dự báo thời tiết mùa đông và cuối tuần",
    dialogue: [
      { speaker: "Диктор", text: "Здравствуйте! В эфире прогноз погоды на предстоящие выходные." },
      { speaker: "Ведущий", text: "В субботу в столичном регионе ожидается облачная погода со снегом." },
      { speaker: "Ведущий", text: "Температура воздуха днём составит минус пять градусов." },
      { speaker: "Ведущий", text: "Ветер умеренный, северный, три-пять метров в секунду." },
      { speaker: "Ведущий", text: "В воскресенье выглянет яркое солнце, но похолодает до минус десяти." },
      { speaker: "Ведущий", text: "На дорогах возможна сильная гололедица. Будьте осторожны!" },
      { speaker: "Диктор", text: "Одевайтесь теплее и приятного всем зимнего отдыха на свежем воздухе!" }
    ],
    dictation: "В субботу ожидается облачная погода со снегом и мороз.",
    dictationHint: "Từ «облачная погода» và giới từ «со снегом».",
    ruAnswers: ["Минус 5 градусов.", "Северный ветер.", "Минус 10 градусов.", "Сильная гололедица."],
    viAnswers: ["Có mây và tuyết rơi.", "Chủ nhật trời có nắng nhưng lạnh hơn.", "Trơn trượt do băng tuyết trên đường.", "Mặc ấm khi ra ngoài."]
  },
  { id: 12, title: "В музее: Экскурсия по залам", time: "11 phút", topic: "Nghe hướng dẫn viên thuyết minh tại bảo tàng nghệ thuật",
    dialogue: [
      { speaker: "Гид", text: "Здравствуйте, дорогие гости! Мы находимся в историческом зале музея." },
      { speaker: "Турист", text: "Скажите, пожалуйста, разрешено ли фотографировать картины?" },
      { speaker: "Гид", text: "Да, можно делать фотографии, но строго без вспышки." },
      { speaker: "Турист", text: "Чья это знаменитая картина перед нами?" },
      { speaker: "Гид", text: "Это шедевр известного русского художника девятнадцатого века." },
      { speaker: "Турист", text: "Сколько времени продлится наша экскурсия?" },
      { speaker: "Гид", text: "Экскурсия продлится ровно один час. Пройдёмте в следующий зал!" }
    ],
    dictation: "В залах музея разрешено фотографировать картины без вспышки.",
    dictationHint: "Cụm từ «без вспышки» (không dùng đèn flash).",
    ruAnswers: ["В историческом зале.", "Да, но строго без вспышки.", "Художника 19-го века.", "Один час."],
    viAnswers: ["Được chụp nhưng cấm dùng đèn flash.", "Thời gian tham quan 1 tiếng.", "Họa sĩ Nga thế kỷ 19.", "Đi sang phòng trưng bày kế tiếp."]
  },
  { id: 13, title: "У врача: Осмотр и рекомендации", time: "11 phút", topic: "Khám sức khỏe tổng quát và đơn thuốc bác sĩ",
    dialogue: [
      { speaker: "Врач", text: "Добрый день! На что вы жалуетесь?" },
      { speaker: "Больной", text: "Здравствуйте, доктор. У меня уже три дня сильный кашель." },
      { speaker: "Врач", text: "Давайте я послушаю ваше дыхание. Дышите глубже... Не дышите..." },
      { speaker: "Больной", text: "Доктор, это грипп или бронхит?" },
      { speaker: "Врач", text: "В лёгких всё чисто. Это трахеит. Я выпишу вам сироп от кашля." },
      { speaker: "Больной", text: "Мне нужен больничный лист для работы?" },
      { speaker: "Врач", text: "Да, я выпишу больничный лист на пять дней. Соблюдайте постельный режим!" }
    ],
    dictation: "Врач выписал сироп от кашля и больничный лист на пять дней.",
    dictationHint: "Từ «больничный лист» (giấy nghỉ ốm hưởng BHXH).",
    ruAnswers: ["Сильный кашель 3 дня.", "Дышите глубже.", "Сироп от кашля.", "На 5 дней."],
    viAnswers: ["Bệnh nhân ho nhiều trong 3 ngày.", "Nghe nhịp thở phổi của người bệnh.", "Nghỉ ngơi tại giường bệnh.", "Cấp giấy nghỉ ốm 5 ngày."]
  },
  { id: 14, title: "Спорт и хобби: Чем ты увлекаешься?", time: "10 phút", topic: "Nói về môn thể thao yêu thích và rèn luyện thân thể",
    dialogue: [
      { speaker: "Денис", text: "Виктор, ты занимаешься спортом?" },
      { speaker: "Виктор", text: "Да, спорт — это моя жизнь! Я плаваю в бассейне три раза в неделю." },
      { speaker: "Денис", text: "А по выходным ты играешь в футбол?" },
      { speaker: "Виктор", text: "По субботам мы с друзьями играем в футбол на стадионе возле дома." },
      { speaker: "Денис", text: "Здорово! А зимой ты катаешься на лыжах или на коньках?" },
      { speaker: "Виктор", text: "Зимой я больше люблю кататься на коньках на Красной площади." },
      { speaker: "Денис", text: "Пойдём вместе в эту субботу на каток!" },
      { speaker: "Виктор", text: "С удовольствием! Договорились." }
    ],
    dictation: "Виктор плавает в бассейне три раза в неделю.",
    dictationHint: "Động từ «плавает» và số lần «три раза в неделю».",
    ruAnswers: ["3 раза в неделю.", "На стадионе возле дома.", "На коньках.", "В субботу."],
    viAnswers: ["Bơi lội ở hồ bơi.", "Chơi bóng đá cùng bạn bè.", "Trượt băng trên Quảng trường Đỏ.", "Đồng ý cùng đi trượt băng."]
  },
  { id: 15, title: "В банке: Обмен валюты", time: "10 phút", topic: "Đổi tiền ngoại tệ và mở tài khoản thẻ tại ngân hàng",
    dialogue: [
      { speaker: "Оператор", text: "Добрый день! Какая услуга вас интересует?" },
      { speaker: "Клиент", text: "Здравствуйте! Я хочу обменять сто долларов на рубли." },
      { speaker: "Оператор", text: "Текущий курс покупки: девяносто два рубля за один доллар." },
      { speaker: "Клиент", text: "Хорошо, меня устраивает этот курс." },
      { speaker: "Оператор", text: "Предоставьте ваш паспорт для оформления операции." },
      { speaker: "Клиент", text: "Пожалуйста, вот мой паспорт." },
      { speaker: "Оператор", text: "Получите девять тысяч двести рублей и квитанцию." },
      { speaker: "Клиент", text: "Спасибо большое за быстрое обслуживание!" }
    ],
    dictation: "Курс обмена валюты составляет девяносто два рубля за доллар.",
    dictationHint: "Số từ «девяносто два» và từ «валюты».",
    ruAnswers: ["100 долларов.", "92 рубля за доллар.", "Паспорт.", "9200 рублей."],
    viAnswers: ["Đổi 100 USD ra đồng rúp.", "92 rúp đổi 1 đô la.", "Xuất trình hộ chiếu.", "Nhận tiền và biên lai hóa đơn."]
  },
  { id: 16, title: "В такси: Как проехать в центр?", time: "10 phút", topic: "Gọi xe taxi và chỉ đường đến Nhà hát Lớn Bolshoi",
    dialogue: [
      { speaker: "Водитель", text: "Добрый вечер! Вы заказывали такси до Театральной площади?" },
      { speaker: "Пассажир", text: "Да, добрый вечер! Нам нужно к главному входу Большого театра." },
      { speaker: "Водитель", text: "Садитесь, пожалуйста. Доедем примерно за двадцать минут." },
      { speaker: "Пассажир", text: "Сейчас в центре города есть пробки?" },
      { speaker: "Водитель", text: "Движение плотное, но на Тверской улице свободно." },
      { speaker: "Пассажир", text: "Отлично, спектакль начинается в семь вечера, мы успеваем." },
      { speaker: "Водитель", text: "Не волнуйтесь, приедем вовремя. Пристегните ремни безопасности." }
    ],
    dictation: "Пристегните ремни безопасности, мы приедем вовремя.",
    dictationHint: "Cụm từ quen thuộc khi lên xe «пристегните ремни безопасности».",
    ruAnswers: ["К Большому театру.", "Около 20 минут.", "В 19:00 (семь вечера).", "Пристегнуть ремни."],
    viAnswers: ["Đến cổng chính Nhà hát Bolshoi.", "Ước tính khoảng 20 phút.", "Bắt đầu lúc 7 giờ tối.", "Cần thắt dây an toàn."]
  },
  { id: 17, title: "В ресторане: Традиционный русский обед", time: "12 phút", topic: "Thưởng thức các món súp củ dền borscht, bánh kếp blini",
    dialogue: [
      { speaker: "Официант", text: "Здравствуйте! Вы готовы сделать заказ?" },
      { speaker: "Гость", text: "Да. Посоветуйте, пожалуйста, традиционные русские блюда." },
      { speaker: "Официант", text: "Рекомендую классический борщ со сметаной и домашние пельмени." },
      { speaker: "Гость", text: "Прекрасно! Принесите одну порцию борща и пельмени с говядиной." },
      { speaker: "Официант", text: "Что будете пить? У нас есть ягодный морс и квас." },
      { speaker: "Гость", text: "Один бокал холодного морса, пожалуйста." },
      { speaker: "Официант", text: "А на десерт желаете горячие русские блины с мёдом?" },
      { speaker: "Гость", text: "Да, с удовольствием попробую!" }
    ],
    dictation: "Гость заказал классический борщ со сметаной и ягодный морс.",
    dictationHint: "Từ «борщ со сметаной» và thức uống «ягодный морс».",
    ruAnswers: ["Борщ и пельмени.", "С говядиной.", "Ягодный морс.", "Блины с мёдом."],
    viAnswers: ["Súp borscht ăn kèm kem chua smetana.", "Sủi cảo nhân thịt bò.", "Nước dâu rừng (mors).", "Bánh kếp blini với mật ong."]
  },
  { id: 18, title: "Телефонный разговор: Приглашение в кино", time: "10 phút", topic: "Gọi điện thoại mời bạn bè đi xem phim chiếu rạp",
    dialogue: [
      { speaker: "Артём", text: "Алло! Привет, Даша! Это Артём." },
      { speaker: "Даша", text: "Привет, Артём! Рада тебя слышать. Что нового?" },
      { speaker: "Артём", text: "Ты свободна в пятницу вечером? В кинотеатре идёт новый фильм." },
      { speaker: "Даша", text: "В пятницу я свободна после шести часов. А какой это фильм?" },
      { speaker: "Артём", text: "Это очень интересная фантастическая комедия." },
      { speaker: "Даша", text: "Здорово! Во сколько начинается сеанс?" },
      { speaker: "Артём", text: "Сеанс в семь вечера. Давай встретимся у входа в кинотеатр в без двадцати семь." },
      { speaker: "Даша", text: "Договорились! До встречи в пятницу!" }
    ],
    dictation: "Сеанс начинается в семь вечера у входа в кинотеатр.",
    dictationHint: "Từ «сеанс» và giờ «в семь вечера».",
    ruAnswers: ["В пятницу вечером.", "Фантастическая комедия.", "В 19:00 (в 7 вечера).", "В 18:40 (без двадцати семь)."],
    viAnswers: ["Gọi điện rủ bạn gái đi xem phim.", "Rảnh rỗi sau 18 giờ.", "Phim hài viễn tưởng.", "Hẹn gặp lúc 18h40 trước rạp."]
  },
  { id: 19, title: "Праздник Новый год: Поздравления и подарки", time: "11 phút", topic: "Không khí đón năm mới, cây thông và ông già tuyết Ded Moroz",
    dialogue: [
      { speaker: "Марина", text: "С наступающим Новым годом, Сергей!" },
      { speaker: "Сергей", text: "Спасибо, Марина! И тебя с праздником! Как ты отмечаешь?" },
      { speaker: "Марина", text: "Мы отмечаем всей семьёй дома. Уже нарядили ёлку и приготовили салат оливье." },
      { speaker: "Сергей", text: "А подарки под ёлку уже положили?" },
      { speaker: "Марина", text: "Да, дети очень ждут Деда Мороза и Снегурочку." },
      { speaker: "Сергей", text: "В двенадцать ночи под бой курантов загадай заветное желание!" },
      { speaker: "Марина", text: "Обязательно! Желаю тебе крепкого здоровья, счастья и успехов в новом году!" },
      { speaker: "Сергей", text: "Спасибо огромное! Пусть все мечты сбудутся!" }
    ],
    dictation: "С наступающим Новым годом! Желаю крепкого здоровья и счастья!",
    dictationHint: "Lời chúc chuẩn «С наступающим Новым годом!» và «крепкого здоровья».",
    ruAnswers: ["Дома всей семьёй.", "Салат оливье.", "Деда Мороза и Снегурочку.", "В 12 часов ночи."],
    viAnswers: ["Trang trí cây thông noel và làm salad olivier.", "Trẻ em háo hức chờ Ded Moroz và Snegurochka.", "Ước nguyện khi chuông đồng hồ điểm 12 tiếng.", "Sức khỏe dồi dào và hạnh phúc."]
  },
  { id: 20, title: "Итоговый экзамен: Диалог с преподавателем", time: "12 phút", topic: "Phần thi vấn đáp kết thúc khóa học tiếng Nga trình độ A1",
    dialogue: [
      { speaker: "Преподаватель", text: "Здравствуйте! Потяните, пожалуйста, экзаменационный билет." },
      { speaker: "Студент", text: "Здравствуйте! Мой билет номер семь: тема «Моя учёба и увлечения»." },
      { speaker: "Преподаватель", text: "Отлично. Расскажите о ваших успехах в изучении русского языка." },
      { speaker: "Студент", text: "Я изучаю русский язык четыре месяца. Я уже умею читать, писать и говорить на базовые темы." },
      { speaker: "Преподаватель", text: "Какой раздел грамматики был для вас самым интересным?" },
      { speaker: "Студент", text: "Мне очень понравилась система падежей и спряжение русских глаголов." },
      { speaker: "Преподаватель", text: "Прекрасный ответ и отличное произношение! Ваша оценка — отлично!" },
      { speaker: "Студент", text: "Огромное спасибо за ваши уроки и поддержку!" }
    ],
    dictation: "Студент сдал экзамен по русскому языку на оценку отлично.",
    dictationHint: "Từ «экзамен» và mức điểm cao nhất «на оценку отлично».",
    ruAnswers: ["Билет номер 7.", "4 месяца.", "Падежи и глаголы.", "Отлично."],
    viAnswers: ["Bốc thăm đề thi số 7.", "Học tiếng Nga được 4 tháng.", "Hệ thống các cách và chia động từ.", "Đạt điểm Giỏi (отлично)."]
  }
];

// Combine all 20 lessons
additionalLessonsData.forEach(item => {
  const qList = [
    {
      id: `l${item.id}_q1`, lang: "ru",
      prompt: `Вопрос 1 по диалогу: что упомянуто?`,
      hint: `Tìm chi tiết trong lời thoại.`,
      options: [
        { key: "A", text: item.ruAnswers[0] },
        { key: "B", text: "Другой вариант ответа." },
        { key: "C", text: "Неверная информация." },
        { key: "D", text: "В тексте не сказано." }
      ],
      ans: "A", explain: `Chi tiết được nêu chính xác trong lời thoại: «${item.ruAnswers[0]}».`
    },
    {
      id: `l${item.id}_q2`, lang: "ru",
      prompt: `Вопрос 2: уточните факт из беседы.`,
      hint: `Chi tiết thứ hai trong đối thoại.`,
      options: [
        { key: "A", text: "Не подходит по смыслу." },
        { key: "B", text: item.ruAnswers[1] },
        { key: "C", text: "Ошибочное утверждение." },
        { key: "D", text: "Никто не говорил об этом." }
      ],
      ans: "B", explain: `Câu trả lời chuẩn theo bài nghe: «${item.ruAnswers[1]}».`
    },
    {
      id: `l${item.id}_q3`, lang: "ru",
      prompt: `Вопрос 3: выберите правильное действие или предмет.`,
      hint: `Quan sát thông tin người nói.`,
      options: [
        { key: "A", text: "Совершенно другой факт." },
        { key: "B", text: "Ошибка в деталях." },
        { key: "C", text: item.ruAnswers[2] },
        { key: "D", text: "Ничего подобного." }
      ],
      ans: "C", explain: `Chính xác theo nội dung hội thoại: «${item.ruAnswers[2]}».`
    },
    {
      id: `l${item.id}_q4`, lang: "ru",
      prompt: `Вопрос 4: каков финал или важный вывод разговора?`,
      hint: `Phần kết của đoạn hội thoại.`,
      options: [
        { key: "A", text: "Планы отменились." },
        { key: "B", text: "Собеседники поспорили." },
        { key: "C", text: "Не удалось договориться." },
        { key: "D", text: item.ruAnswers[3] }
      ],
      ans: "D", explain: `Kết thúc đoạn hội thoại xác nhận: «${item.ruAnswers[3]}».`
    },
    {
      id: `l${item.id}_q5`, lang: "vi",
      prompt: `Chủ đề chính của đoạn hội thoại nghe được là gì?`,
      hint: `Khái quát bối cảnh giao tiếp.`,
      options: [
        { key: "A", text: item.viAnswers[0] },
        { key: "B", text: "Tranh cãi về thời tiết." },
        { key: "C", text: "Học ngữ pháp nâng cao C1." },
        { key: "D", text: "Đặt vé máy bay khẩn cấp." }
      ],
      ans: "A", explain: `Hội thoại xoay quanh chủ đề: ${item.viAnswers[0]}.`
    },
    {
      id: `l${item.id}_q6`, lang: "vi",
      prompt: `Thông tin chi tiết nào sau đây là đúng với nội dung bạn vừa nghe?`,
      hint: `Đối chiếu lời người nói.`,
      options: [
        { key: "A", text: "Các nhân vật không hiểu ý nhau." },
        { key: "B", text: item.viAnswers[1] },
        { key: "C", text: "Họ hẹn gặp vào tháng sau." },
        { key: "D", text: "Cuộc gặp bị hủy bỏ vì trời mưa." }
      ],
      ans: "B", explain: `Nội dung nghe được xác nhận: ${item.viAnswers[1]}.`
    },
    {
      id: `l${item.id}_q7`, lang: "vi",
      prompt: `Yếu tố quan trọng nào được các nhân vật đồng thuận trong hội thoại?`,
      hint: `Tìm điểm mấu chốt được nhấn mạnh.`,
      options: [
        { key: "A", text: "Không có ai đồng ý." },
        { key: "B", text: "Phải trả thêm tiền phạt." },
        { key: "C", text: item.viAnswers[2] },
        { key: "D", text: "Rời khỏi nơi gặp mặt ngay lập tức." }
      ],
      ans: "C", explain: `Chi tiết mấu chốt: ${item.viAnswers[2]}.`
    },
    {
      id: `l${item.id}_q8`, lang: "vi",
      prompt: `Ý nghĩa ứng dụng thực tế của đoạn hội thoại giao tiếp này là gì?`,
      hint: `Bài học giao tiếp đời sống thực tiễn.`,
      options: [
        { key: "A", text: "Giúp người học biết cách xử lý tình huống: " + item.viAnswers[3] },
        { key: "B", text: "Không có tính ứng dụng trong đời sống." },
        { key: "C", text: "Chỉ dùng cho văn bản hành chính." },
        { key: "D", text: "Chỉ dùng khi tranh luận tòa án." }
      ],
      ans: "A", explain: `Giúp người học ghi nhớ và áp dụng mẫu câu: ${item.viAnswers[3]}.`
    }
  ];

  lessons.push({
    id: item.id,
    level: "A1",
    skill: "listening",
    title: item.title,
    listeningTime: item.time,
    topic: item.topic,
    dialogue: item.dialogue,
    dictation: {
      title: "Chép chính tả (Диктант)",
      instruction: "Nghe câu audio và viết lại chính xác bằng tiếng Nga vào ô bên dưới:",
      targetText: item.dictation,
      hint: `Gợi ý: ${item.dictationHint}`
    },
    questions: qList
  });
});

// Ghi 20 file lesson-01..20.json
const manifest = [];
lessons.forEach(l => {
  const fileName = `lesson-${String(l.id).padStart(2, '0')}.json`;
  const filePath = path.join(outputDir, fileName);
  fs.writeFileSync(filePath, JSON.stringify(l, null, 2), 'utf8');

  manifest.push({
    id: l.id,
    file: fileName,
    title: l.title,
    badge: "A1 Nghe hiểu",
    listeningTime: l.listeningTime,
    topic: l.topic,
    questionCount: l.questions.length
  });
});

// Ghi manifest.json
const manifestPath = path.join(outputDir, 'manifest.json');
fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2), 'utf8');

console.log(`✓ Đã tạo thành công 20 bài học Listening A1 và manifest.json tại ${outputDir}`);
