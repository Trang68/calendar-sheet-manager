const fs = require('fs');
const path = require('path');

const outputDir = path.join(__dirname, '../public/data/writing/a1');
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

const writingLessons = [
  {
    id: 1,
    title: "Письмо другу: Рассказ о себе",
    topic: "Viết thư thân mật: Kể về bản thân, tên tuổi và quê quán",
    targetWordRange: [35, 60],
    task: "Bạn hãy viết một bức thư ngắn (từ 35 đến 60 từ) gửi cho một người bạn Nga tên là Ivan để làm quen.",
    requirements: [
      "1. Lời chào bạn thân mật (Привет / Здравствуй).",
      "2. Giới thiệu tên, tuổi và nơi bạn đang sinh sống.",
      "3. Giới thiệu trường học hoặc công việc của bạn.",
      "4. Kể một sở thích khi rảnh rỗi (thể thao, đọc sách, âm nhạc...).",
      "5. Lời chào tạm biệt thân thiện (Пока / До встречи)."
    ],
    cliches: [
      { ru: "Привет, Иван!", vi: "Chào Ivan!" },
      { ru: "Меня зовут...", vi: "Tôi tên là..." },
      { ru: "Мне двадцать лет.", vi: "Tôi 20 tuổi." },
      { ru: "Я живу во Вьетнаме, в Ханое.", vi: "Tôi sống ở Việt Nam, tại Hà Nội." },
      { ru: "Я студент, я учусь в университете.", vi: "Tôi là sinh viên, tôi học đại học." },
      { ru: "В свободное время я люблю читать книги.", vi: "Rảnh rỗi tôi thích đọc sách." },
      { ru: "А как твои дела? Напиши мне!", vi: "Còn bạn thế nào? Hãy viết cho tôi nhé!" },
      { ru: "До скорой встречи! Пока!", vi: "Hẹn sớm gặp lại! Tạm biệt!" }
    ],
    modelAnswer: "Привет, дорогой Иван!\n\nМеня зовут Минх. Мне двадцать лет. Я живу во Вьетнаме, в красивом городе Ханой. Я студент и учусь в университете. Я очень люблю изучать русский язык. В свободное время я люблю играть в футбол и слушать хорошую музыку.\n\nА как твои дела? Напиши мне ответ!\nПока,\nТвой друг Минх.",
    modelWordCount: 51
  },
  {
    id: 2,
    title: "Моя семья: Рассказ о близких",
    topic: "Gia đình thân yêu: Miêu tả bố mẹ, nghề nghiệp và tình cảm gia đình",
    targetWordRange: [35, 60],
    task: "Viết một đoạn văn ngắn (35–60 từ) giới thiệu về các thành viên trong gia đình bạn.",
    requirements: [
      "1. Nêu số lượng thành viên trong gia đình (bố, mẹ, anh chị em).",
      "2. Giới thiệu nghề nghiệp của bố và mẹ.",
      "3. Nêu hoạt động hoặc sở thích chung của cả gia đình vào cuối tuần.",
      "4. Khẳng định tình cảm yêu thương, gắn bó."
    ],
    cliches: [
      { ru: "У меня дружная семья.", vi: "Tôi có một gia đình hòa thuận." },
      { ru: "Мой папа работает...", vi: "Bố tôi làm việc..." },
      { ru: "Моя мама — учительница.", vi: "Mẹ tôi là giáo viên." },
      { ru: "В выходные мы гуляем вместе.", vi: "Cuối tuần chúng tôi cùng đi dạo." },
      { ru: "Я очень люблю свою семью.", vi: "Tôi rất yêu gia đình của mình." }
    ],
    modelAnswer: "У меня большая и дружная семья. Нас четыре человека: папа, мама, младшая сестра и я. Мой папа работает инженером, а мама работает врачом в городской больнице. Моя сестра учится в школе. По выходным мы любим готовить вкусный обед и смотреть фильмы вместе. Я очень люблю мою семью!",
    modelWordCount: 48
  },
  {
    id: 3,
    title: "Моя комната: Описание интерьера",
    topic: "Căn phòng của tôi: Miêu tả đồ đạc, bàn học và vị trí các vật dụng",
    targetWordRange: [35, 60],
    task: "Viết đoạn văn ngắn miêu tả căn phòng ngủ hoặc góc học tập của bạn (35–60 từ).",
    requirements: [
      "1. Giới thiệu căn phòng rộng hay hẹp, sáng sủa và ấm cúng.",
      "2. Miêu tả vị trí các đồ nội thất chính (bàn học, giường, tủ sách, máy tính).",
      "3. Sử dụng các trạng từ chỉ vị trí (слева, справа, в центре, на столе).",
      "4. Bày tỏ cảm xúc của bạn về căn phòng."
    ],
    cliches: [
      { ru: "Это моя комната.", vi: "Đây là phòng của tôi." },
      { ru: "Она светлая и уютная.", vi: "Nó rất sáng sủa và ấm cúng." },
      { ru: "Справа у окна стоит большой стол.", vi: "Bên phải cạnh cửa sổ có một chiếc bàn lớn." },
      { ru: "На столе стоит мой компьютер.", vi: "Trên bàn có máy tính của tôi." },
      { ru: "Слева находится книжный шкаф.", vi: "Bên trái là tủ sách." },
      { ru: "Мне нравится здесь отдыхать.", vi: "Tôi thích nghỉ ngơi ở đây." }
    ],
    modelAnswer: "Это моя комната. Она небольшая, но очень светлая и уютная. В центре комнаты стоит удобный диван. Справа у окна находится мой письменный стол. На столе стоят ноутбук и лампа. Слева висит книжная полка, где лежат мои учебники. Мне очень нравится здесь учиться и отдыхать после занятий.",
    modelWordCount: 48
  },
  {
    id: 4,
    title: "Мой обычный день: Распорядок",
    topic: "Một ngày thường nhật: Các mốc thời gian từ sáng đến tối",
    targetWordRange: [40, 65],
    task: "Viết đoạn văn (40–65 từ) kể về lịch sinh hoạt trong một ngày bình thường của bạn.",
    requirements: [
      "1. Nêu giờ thức dậy và các thói quen buổi sáng (tập thể dục, ăn sáng).",
      "2. Nêu thời gian học tập hoặc làm việc buổi sáng và trưa.",
      "3. Kể hoạt động buổi chiều và buổi tối (nghỉ ngơi, đọc sách, đi ngủ).",
      "4. Sử dụng liên từ thời gian (утром, днём, вечером, потом)."
    ],
    cliches: [
      { ru: "Обычно я встаю в семь часов утра.", vi: "Thông thường tôi dậy lúc 7 giờ sáng." },
      { ru: "Я завтракаю и иду на занятия.", vi: "Tôi ăn sáng rồi đi học." },
      { ru: "В час дня я обедаю в столовой.", vi: "1 giờ chiều tôi ăn trưa ở căng tin." },
      { ru: "Вечером я делаю домашнее задание.", vi: "Buổi tối tôi làm bài tập về nhà." },
      { ru: "Я ложусь спать в одиннадцать часов.", vi: "Tôi đi ngủ lúc 11 giờ." }
    ],
    modelAnswer: "Мой день начинается рано. Обычно я встаю в семь часов утра, умываюсь и пью чай. В восемь часов я иду в университет на лекции. Днём в час я обедаю с друзьями в кафе. После пар я отдыхаю, читаю книги и делаю домашнее задание. В одиннадцать часов вечера я ложусь спать.",
    modelWordCount: 50
  },
  {
    id: 5,
    title: "В кафе: Мой любимый обед",
    topic: "Bữa ăn yêu thích: Món ăn, đồ uống và quán ăn ưa thích",
    targetWordRange: [35, 60],
    task: "Viết đoạn văn ngắn (35–60 từ) kể về quán ăn bạn thích và những món bạn hay gọi.",
    requirements: [
      "1. Nêu tên hoặc địa điểm quán ăn bạn hay đến.",
      "2. Kể tên các món ăn và thức uống yêu thích.",
      "3. Nhận xét về hương vị và giá cả món ăn.",
      "4. Lời rủ bạn bè cùng ghé quán."
    ],
    cliches: [
      { ru: "Я часто хожу в это кафе.", vi: "Tôi hay đến quán cà phê này." },
      { ru: "Моё любимое блюдо — это...", vi: "Món ăn yêu thích của tôi là..." },
      { ru: "Здесь готовят очень вкусно.", vi: "Ở đây nấu ăn rất ngon." },
      { ru: "Цены здесь совсем невысокие.", vi: "Giá cả ở đây không hề đắt." },
      { ru: "Советую всем попробовать!", vi: "Khuyên mọi người nên thử!" }
    ],
    modelAnswer: "Я очень люблю ходить в уютное кафе возле нашего университета. Там всегда играет приятная музыка. На обед я обычно заказываю горячий мясной суп, салат и зелёный чай. Еда здесь всегда свежая и невероятно вкусная, а цены совсем недорогие. Я рекомендую моим друзьям обязательно посетить это прекрасное место!",
    modelWordCount: 47
  },
  {
    id: 6,
    title: "Мой родной город",
    topic: "Thành phố quê hương: Tên thành phố, danh lam thắng cảnh và thời tiết",
    targetWordRange: [40, 65],
    task: "Viết đoạn văn (40–65 từ) giới thiệu về thành phố nơi bạn sinh ra và lớn lên.",
    requirements: [
      "1. Tên thành phố và vị trí địa lý.",
      "2. Những nét nổi bật (phố xá, công viên, hồ nước, viện bảo tàng).",
      "3. Con người và nhịp sống nơi đây.",
      "4. Bày tỏ tình yêu đối với quê hương."
    ],
    cliches: [
      { ru: "Я родился и живу в городе...", vi: "Tôi sinh ra và sống ở thành phố..." },
      { ru: "Это красивый и старинный город.", vi: "Đây là thành phố đẹp và cổ kính." },
      { ru: "Здесь много парков и музеев.", vi: "Ở đây có nhiều công viên và bảo tàng." },
      { ru: "Люди здесь добрые и приветливые.", vi: "Con người nơi đây tốt bụng và hiếu khách." },
      { ru: "Я горжусь своим городом.", vi: "Tôi tự hào về thành phố của mình." }
    ],
    modelAnswer: "Я живу в столице Вьетнама — в Ханое. Это древний, красивый и очень зелёный город. В центре города находится знаменитое озеро Возвращённого Меча, где любят гулять жители и туристы. В Ханое много интересных музеев, театров и уютных кафе. Люди здесь очень гостеприимные и дружелюбные. Я всей душой люблю свой родной город!",
    modelWordCount: 51
  },
  {
    id: 7,
    title: "Моё любимое время года",
    topic: "Mùa yêu thích trong năm: Cảnh sắc thiên nhiên và hoạt động",
    targetWordRange: [35, 60],
    task: "Viết đoạn văn (35–60 từ) kể về mùa mà bạn yêu thích nhất và lý do vì sao.",
    requirements: [
      "1. Nêu rõ mùa bạn thích nhất (mùa xuân, hạ, thu hay đông).",
      "2. Miêu tả đặc điểm thời tiết của mùa đó (nắng, gió, mưa, nhiệt độ).",
      "3. Kể những hoạt động bạn thường làm trong mùa này.",
      "4. Lý do vì sao mùa này mang lại niềm vui cho bạn."
    ],
    cliches: [
      { ru: "Моё любимое время года — осень.", vi: "Mùa yêu thích của tôi là mùa thu." },
      { ru: "Погода очень комфортная и мягкая.", vi: "Thời tiết rất dễ chịu và dịu nhẹ." },
      { ru: "Листья на деревьях становятся жёлтыми.", vi: "Lá trên cây chuyển sang màu vàng." },
      { ru: "Я люблю гулять по аллеям парка.", vi: "Tôi thích đi dạo trên các lối đi công viên." }
    ],
    modelAnswer: "Моё самое любимое время года — это золотая осень. В это время уже нет летней жары, погода становится приятной и прохладной. Деревья в парках становятся жёлтыми и красными. Я обожаю гулять по красивым улицам, дышать свежим воздухом и делать красивые фотографии природы. Осень дарит мне спокойствие и радость.",
    modelWordCount: 47
  },
  {
    id: 8,
    title: "Мой лучший друг",
    topic: "Người bạn thân nhất: Ngoại hình, tính cách và sở thích chung",
    targetWordRange: [35, 60],
    task: "Viết đoạn văn ngắn (35–60 từ) miêu tả người bạn thân nhất của bạn.",
    requirements: [
      "1. Tên, tuổi và bạn quen người đó trong hoàn cảnh nào.",
      "2. Miêu tả ngoại hình và tính cách tốt bụng, vui vẻ.",
      "3. Hai bạn thường làm gì cùng nhau khi rảnh rỗi.",
      "4. Khẳng định giá trị của tình bạn."
    ],
    cliches: [
      { ru: "Моего лучшего друга зовут...", vi: "Bạn thân nhất của tôi tên là..." },
      { ru: "Мы познакомились в университете.", vi: "Chúng tôi quen nhau ở trường đại học." },
      { ru: "Он умный, весёлый и надёжный человек.", vi: "Bạn ấy thông minh, vui tính và đáng tin cậy." },
      { ru: "Он всегда готов помочь в трудную минуту.", vi: "Bạn ấy luôn sẵn sàng giúp đỡ lúc khó khăn." }
    ],
    modelAnswer: "Моего лучшего друга зовут Алексей. Мы познакомились два года назад в университете. Алексей — высокий, весёлый и очень умный парень. Он отлично учится и всегда помогает мне с трудной математикой. В свободное время мы вместе играем в футбол и слушаем рок-музыку. Я очень рад, что у меня есть такой верный друг!",
    modelWordCount: 50
  },
  {
    id: 9,
    title: "Как я провёл выходные",
    topic: "Kỳ nghỉ cuối tuần: Những hoạt động thư giãn thú vị vừa qua",
    targetWordRange: [40, 65],
    task: "Viết đoạn văn ngắn (40–65 từ) kể về những việc bạn đã làm vào thứ Bảy và Chủ nhật vừa rồi.",
    requirements: [
      "1. Kể hoạt động vào thứ Bảy (dọn dẹp, đi chơi, mua sắm).",
      "2. Kể hoạt động vào Chủ nhật (gặp bạn bè, xem phim, về thăm gia đình).",
      "3. Cảm xúc sau những ngày nghỉ ngơi.",
      "4. Sử dụng thì quá khứ của động từ (отдыхал, гулял, читал)."
    ],
    cliches: [
      { ru: "В прошлые выходные я отлично отдохнул.", vi: "Cuối tuần vừa rồi tôi đã nghỉ ngơi rất tuyệt." },
      { ru: "В субботу я ходил в кино с друзьями.", vi: "Thứ Bảy tôi đã đi xem phim cùng bạn bè." },
      { ru: "В воскресенье мы гуляли в парке.", vi: "Chủ nhật chúng tôi đã đi dạo trong công viên." },
      { ru: "Это были замечательные выходные!", vi: "Đó là những ngày cuối tuần thật tuyệt vời!" }
    ],
    modelAnswer: "В прошлые выходные я замечательно отдохнул. В субботу утром я выспался, сделал уборку в комнате, а днём пошёл в кинотеатр смотреть новый фильм. В воскресенье погода была тёплой, поэтому мы с друзьями долго гуляли в центральном парке и ели мороженое. Я набрался сил перед новой учебной неделей!",
    modelWordCount: 47
  },
  {
    id: 10,
    title: "Почему я изучаю русский язык",
    topic: "Lý do học tiếng Nga: Mục tiêu nghề nghiệp, văn hóa và du học",
    targetWordRange: [40, 65],
    task: "Viết đoạn văn (40–65 từ) giải thích vì sao bạn chọn học tiếng Nga và mục tiêu tương lai.",
    requirements: [
      "1. Bạn bắt đầu học tiếng Nga từ khi nào.",
      "2. Nét đẹp của tiếng Nga (ngữ điệu, văn học, âm thanh).",
      "3. Mục tiêu học tập (đi du học, làm việc với đối tác Nga, đi du lịch).",
      "4. Quyết tâm rèn luyện mỗi ngày."
    ],
    cliches: [
      { ru: "Я изучаю русский язык уже...", vi: "Tôi học tiếng Nga đã được..." },
      { ru: "Русский язык очень богатый и красивый.", vi: "Tiếng Nga rất phong phú và tuyệt đẹp." },
      { ru: "Я хочу поехать учиться в Россию.", vi: "Tôi muốn sang Nga du học." },
      { ru: "Каждый день я учу новые слова.", vi: "Mỗi ngày tôi đều học từ mới." }
    ],
    modelAnswer: "Я изучаю русский язык уже шесть месяцев. Мне очень нравится, как звучит русская речь — она певучая и красивая. Я хочу хорошо знать русский язык, чтобы читать книги Пушкина и Чехова в оригинале. В будущем я мечтаю поехать учиться в Санкт-Петербург. Поэтому я каждый день усердно учу грамматику и новые слова!",
    modelWordCount: 50
  },
  {
    id: 11,
    title: "В магазине продуктов: Список покупок",
    topic: "Đi siêu thị: Danh sách thực phẩm và chuẩn bị bữa tiệc nhỏ",
    targetWordRange: [35, 60],
    task: "Viết đoạn văn ngắn (35–60 từ) kể về việc bạn đi siêu thị mua sắm thực phẩm cho gia đình.",
    requirements: [
      "1. Nêu lý do đi mua sắm (chuẩn bị bữa tối, sinh nhật).",
      "2. Liệt kê các thực phẩm đã mua (sữa, bánh mì, phô mai, rau quả).",
      "3. Cách thanh toán (tiền mặt hay thẻ ngân hàng).",
      "4. Nhận xét về chất lượng hàng hóa."
    ],
    cliches: [
      { ru: "Сегодня я пошёл в большой супермаркет.", vi: "Hôm nay tôi đã đến một siêu thị lớn." },
      { ru: "Мне нужно было купить продукты на ужин.", vi: "Tôi cần mua thực phẩm cho bữa tối." },
      { ru: "Я купил молоко, свежий хлеб и сыр.", vi: "Tôi đã mua sữa, bánh mì tươi và phô mai." },
      { ru: "Я оплатил покупки картой на кассе.", vi: "Tôi thanh toán các món đồ bằng thẻ tại quầy thu ngân." }
    ],
    modelAnswer: "Сегодня днём я пошёл в большой супермаркет около дома. Мне нужно было купить продукты для праздничного ужина. Я выбрал свежий белый хлеб, сыр, молоко, спелые яблоки и килограмм куриного мяса. На кассе я быстро расплатился банковской картой. Все продукты оказались очень свежими и качественными. Вечером мы приготовили вкусный ужин!",
    modelWordCount: 49
  },
  {
    id: 12,
    title: "Приглашение на день рождения",
    topic: "Thư mời sinh nhật: Thời gian, địa điểm và nội dung buổi tiệc",
    targetWordRange: [35, 60],
    task: "Viết một bức thư ngắn (35–60 từ) mời bạn bè đến dự tiệc sinh nhật của bạn.",
    requirements: [
      "1. Lời mời nhiệt tình và nêu lý do (tiệc sinh nhật).",
      "2. Nêu chính xác thời gian và địa điểm tổ chức.",
      "3. Kể sơ qua các hoạt động trong tiệc (ăn bánh ga-tô, nghe nhạc, trò chơi).",
      "4. Đề nghị bạn xác nhận có thể đến hay không."
    ],
    cliches: [
      { ru: "Дорогой друг, приглашаю тебя на мой праздник!", vi: "Bạn thân mến, mời bạn đến ngày vui của tôi!" },
      { ru: "В эту субботу у меня день рождения.", vi: "Thứ Bảy tuần này là sinh nhật của tôi." },
      { ru: "Праздник начнётся в шесть часов вечера.", vi: "Bữa tiệc sẽ bắt đầu lúc 6 giờ tối." },
      { ru: "Мы будем есть торт и петь песни.", vi: "Chúng ta sẽ ăn bánh kem và hát hò." },
      { ru: "Жду тебя с нетерпением!", vi: "Rất mong chờ bạn đến!" }
    ],
    modelAnswer: "Дорогая Анна!\nВ эту субботу, двадцатого октября, у меня день рождения. Я приглашаю тебя в гости! Праздник начнётся в шесть часов вечера в кафе «Сказка». Мы будем есть праздничный торт, пить чай, слушать весёлую музыку и танцевать. Напиши мне, пожалуйста, сможешь ли ты прийти. Очень жду тебя!\nТвой друг Олег.",
    modelWordCount: 50
  },
  {
    id: 13,
    title: "Письмо из отпуска: Привет с моря!",
    topic: "Bưu thiếp du lịch: Kể về chuyến nghỉ mát tại biển và lời hỏi thăm",
    targetWordRange: [40, 65],
    task: "Viết một bưu thiếp hoặc lá thư ngắn (40–65 từ) gửi cho bạn từ khu nghỉ dưỡng ven biển.",
    requirements: [
      "1. Nêu địa điểm bạn đang đi nghỉ (bãi biển, thành phố Nha Trang, Đà Nẵng, Sochi).",
      "2. Miêu tả thời tiết nắng ấm, nước biển trong xanh.",
      "3. Kể các hoạt động (tắm biển, phơi nắng, ăn hải sản).",
      "4. Hứa sẽ gửi quà lưu niệm và lời chúc tốt lành."
    ],
    cliches: [
      { ru: "Привет из солнечного города!", vi: "Gửi lời chào từ thành phố đầy nắng!" },
      { ru: "Я сейчас отдыхаю на море.", vi: "Hiện tại tôi đang nghỉ mát ngoài biển." },
      { ru: "Погода здесь великолепная и тёплая.", vi: "Thời tiết ở đây tuyệt đẹp và ấm áp." },
      { ru: "Каждый день я купаюсь и загораю.", vi: "Mỗi ngày tôi đều tắm biển và tắm nắng." },
      { ru: "Скоро увидимся!", vi: "Sớm gặp lại bạn nhé!" }
    ],
    modelAnswer: "Привет, Виктор!\nПередаю тебе огромный тёплый привет из Дананга! Я сейчас отдыхаю на море. Погода здесь просто чудесная: каждый день светит жаркое солнце, а море чистое и тёплое. Утром я плаваю, а вечером гуляю по набережной и ем вкусные морепродукты. Я купил для тебя красивый сувенир. Скоро вернусь и всё расскажу!\nТвой друг.",
    modelWordCount: 52
  },
  {
    id: 14,
    title: "Мой любимый праздник: Новый год",
    topic: "Lễ hội Năm Mới: Phong tục trang trí cây thông và tặng quà",
    targetWordRange: [40, 65],
    task: "Viết đoạn văn (40–65 từ) kể về cách bạn và gia đình chuẩn bị đón Tết Năm Mới.",
    requirements: [
      "1. Khẳng định Năm Mới là ngày lễ kỳ diệu và được chờ đợi nhất.",
      "2. Miêu tả việc trang hoàng cây thông và nhà cửa.",
      "3. Kể các món ăn truyền thống trong đêm giao thừa.",
      "4. Lời chúc tụng và trao gửi quà tặng may mắn."
    ],
    cliches: [
      { ru: "Новый год — это волшебный праздник.", vi: "Năm Mới là một ngày lễ kỳ diệu." },
      { ru: "Мы украшаем новогоднюю ёлку игрушками.", vi: "Chúng tôi trang trí cây thông bằng đồ chơi." },
      { ru: "Мама готовит вкусные салаты и пироги.", vi: "Mẹ chuẩn bị những món salad và bánh pỉ-rốc ngon." },
      { ru: "Под бой часов мы загадываем желания.", vi: "Dưới tiếng chuông đồng hồ chúng tôi ước nguyện." }
    ],
    modelAnswer: "Новый год — мой самый любимый праздник в году. В конце декабря мы всей семьёй украшаем пушистую зелёную ёлку разноцветными шарами и гирляндами. В новогоднюю ночь мама готовит традиционные салаты и горячие блюда. Ровно в полночь мы чокаемся бокалами, поздравляем друг друга и дарим подарки. Это время настоящей радости и тепла!",
    modelWordCount: 49
  },
  {
    id: 15,
    title: "В библиотеке: Подготовка к занятиям",
    topic: "Học tập tại thư viện: Mượn giáo trình và tra cứu tài liệu",
    targetWordRange: [35, 60],
    task: "Viết đoạn văn ngắn (35–60 từ) kể về việc bạn đến thư viện trường học tập và mượn sách.",
    requirements: [
      "1. Nêu thời điểm và lý do bạn đến thư viện (ôn thi, tìm sách).",
      "2. Miêu tả không gian yên tĩnh và số lượng sách phong phú.",
      "3. Kể việc trao đổi với thủ thư và mượn từ điển/giáo trình.",
      "4. Cảm nghĩ về lợi ích của việc đọc sách."
    ],
    cliches: [
      { ru: "Вчера я ходил в университетскую библиотеку.", vi: "Hôm qua tôi đã đến thư viện trường đại học." },
      { ru: "Там очень тихо и удобно заниматься.", vi: "Ở đó rất yên tĩnh và thuận tiện học bài." },
      { ru: "Я взял большой русско-вьетнамский словарь.", vi: "Tôi đã mượn một cuốn từ điển Nga-Việt lớn." },
      { ru: "Чтение помогает мне лучше знать язык.", vi: "Đọc sách giúp tôi giỏi ngôn ngữ hơn." }
    ],
    modelAnswer: "После занятий я часто хожу в нашу университетскую библиотеку. Там всегда царит тишина и порядок. Я показал библиотекарю свой студенческий билет и взял новый учебник по грамматике и русско-английский словарь. В читальном зале я просидел два часа, делая конспекты к семинару. Библиотека — это лучшее место для серьёзной учёбы!",
    modelWordCount: 47
  },
  {
    id: 16,
    title: "Моё любимое животное",
    topic: "Thú cưng yêu quý: Miêu tả chú chó/mèo thông minh và đáng yêu",
    targetWordRange: [35, 60],
    task: "Viết đoạn văn (35–60 từ) kể về con vật nuôi mà bạn yêu mến trong nhà.",
    requirements: [
      "1. Tên con vật, loài (chó, mèo) và màu lông.",
      "2. Thói quen hàng ngày và sự thông minh, hiền lành.",
      "3. Bạn chăm sóc và chơi đùa với nó như thế nào.",
      "4. Tình bạn gắn bó giữa bạn và con vật."
    ],
    cliches: [
      { ru: "У меня дома живёт пушистый кот.", vi: "Ở nhà tôi có một chú mèo lông xù sinh sống." },
      { ru: "Его зовут Барсик, ему два года.", vi: "Nó tên là Barsik, nó được 2 tuổi." },
      { ru: "У него белая шерсть и зелёные глаза.", vi: "Nó có bộ lông trắng và đôi mắt xanh lá." },
      { ru: "Он очень ласковый и умный.", vi: "Nó rất quấn chủ và thông minh." }
    ],
    modelAnswer: "У меня дома живёт замечательная собака. Её зовут Найда. Она белая с чёрными пятнами и очень добрыми глазами. Каждое утро и вечер я гуляю с ней в сквере около дома. Найда очень умная, она знает команды и любит играть с мячом. Когда я возвращаюсь из университета, она радостно встречает меня у двери!",
    modelWordCount: 51
  },
  {
    id: 17,
    title: "Экскурсия по музею",
    topic: "Chuyến tham quan bảo tàng: Những bức tranh và di tích lịch sử",
    targetWordRange: [40, 65],
    task: "Viết đoạn văn (40–65 từ) kể về trải nghiệm tham quan một viện bảo tàng nghệ thuật hoặc lịch sử.",
    requirements: [
      "1. Nêu tên bảo tàng và thời gian bạn đến thăm.",
      "2. Những hiện vật hoặc bức tranh bạn ấn tượng nhất.",
      "3. Lắng nghe thuyết minh của hướng dẫn viên.",
      "4. Nhận xét và cảm xúc sau chuyến đi."
    ],
    cliches: [
      { ru: "Недавно мы всем классом посетили музей.", vi: "Mới đây cả lớp chúng tôi đã đến thăm bảo tàng." },
      { ru: "Экскурсовод рассказал много интересного.", vi: "Hướng dẫn viên đã kể nhiều điều lý thú." },
      { ru: "Мне очень понравились старинные картины.", vi: "Tôi rất thích những bức tranh cổ kính." },
      { ru: "Это была познавательная поездка.", vi: "Đó là một chuyến đi mở mang nhiều kiến thức." }
    ],
    modelAnswer: "В прошлую пятницу мы с преподавателем ходили в Государственный исторический музей. Там собраны тысячи уникальных старинных экспонатов: древние монеты, оружие, книги и картины. Экскурсовод очень увлекательно рассказывал о богатой культуре и истории страны. Я узнал много нового и сделал прекрасные снимки на память. Эта экскурсия произвела на меня огромное впечатление!",
    modelWordCount: 49
  },
  {
    id: 18,
    title: "Спорт в моей жизни",
    topic: "Thể thao và sức khỏe: Môn thể thao tập luyện và lợi ích",
    targetWordRange: [40, 65],
    task: "Viết đoạn văn (40–65 từ) kể về vai trò của thể thao và môn thể thao bạn thường chơi.",
    requirements: [
      "1. Nêu môn thể thao bạn chơi (bóng đá, bơi, chạy bộ, cầu lông).",
      "2. Tần suất tập luyện mỗi tuần và địa điểm tập.",
      "3. Lợi ích sức khỏe và tinh thần mà thể thao mang lại.",
      "4. Lời khuyên mọi người nên năng rèn luyện."
    ],
    cliches: [
      { ru: "Спорт играет важную роль в моей жизни.", vi: "Thể thao đóng vai trò quan trọng trong đời sống tôi." },
      { ru: "Я занимаюсь плаванием три раза в неделю.", vi: "Tôi tập bơi ba lần một tuần." },
      { ru: "Спорт даёт мне бодрость и энергию.", vi: "Thể thao tiếp cho tôi sự tỉnh táo và năng lượng." },
      { ru: "Здоровый образ жизни необходим каждому.", vi: "Lối sống lành mạnh là điều cần thiết cho mọi người." }
    ],
    modelAnswer: "Спорт занимает важное место в моей жизни. Я считаю, что физическая активность необходима каждому человеку для хорошего здоровья. Два раза в неделю я хожу в бассейн и плаваю, а по субботам играю в футбол с друзьями на стадионе. Спорт помогает мне снять стресс после учёбы, быть бодрым, сильным и жизнерадостным!",
    modelWordCount: 49
  },
  {
    id: 19,
    title: "Планы на летние каникулы",
    topic: "Dự định kỳ nghỉ hè: Kế hoạch du lịch, nghỉ ngơi và học tập",
    targetWordRange: [40, 65],
    task: "Viết đoạn văn (40–65 từ) chia sẻ về những dự định của bạn trong kỳ nghỉ hè sắp tới.",
    requirements: [
      "1. Nêu thời gian kỳ nghỉ hè bắt đầu sau khi thi xong.",
      "2. Địa điểm bạn dự định đến thăm (về quê, ra biển, đi leo núi).",
      "3. Kế hoạch trau dồi bản thân (đọc sách, học ngoại ngữ).",
      "4. Kỳ vọng về một mùa hè tràn ngập kỷ niệm đẹp."
    ],
    cliches: [
      { ru: "Скоро начнутся долгожданные летние каникулы.", vi: "Kỳ nghỉ hè mong đợi bấy lâu sắp bắt đầu rồi." },
      { ru: "Сначала я поеду в родную деревню к бабушке.", vi: "Trước tiên tôi sẽ về làng quê thăm bà." },
      { ru: "Потом мы с друзьями отправимся в поход.", vi: "Sau đó tôi cùng bạn bè sẽ lên đường đi cắm trại." },
      { ru: "Я надеюсь отлично провести это лето!", vi: "Tôi hy vọng sẽ trải qua một mùa hè thật tuyệt!" }
    ],
    modelAnswer: "Скоро закончатся экзамены и начнутся долгожданные летние каникулы. В июле я планирую поехать на две недели на морское побережье, чтобы плавать, загорать и дышать чистым морским воздухом. В августе я буду отдыхать в деревне у бабушки, читать интересные книги и продолжать изучать русский язык. Уверен, что это лето будет ярким и незабываемым!",
    modelWordCount: 52
  },
  {
    id: 20,
    title: "Итоговое сочинение: Мои успехи на уровне А1",
    topic: "Bài luận tổng kết A1: Chặng đường đã qua và mục tiêu chinh phục A2",
    targetWordRange: [45, 70],
    task: "Viết bài luận ngắn (45–70 từ) tổng kết quá trình học tiếng Nga của bạn từ con số 0 đến hết trình độ A1.",
    requirements: [
      "1. Khẳng định bạn đã hoàn thành trọn vẹn chương trình A1.",
      "2. Kể những kỹ năng bạn đã làm chủ (đọc hiểu, nghe hội thoại, viết thư, phát âm).",
      "3. Cảm xúc tự hào và lời tri ân gửi đến thầy cô, bạn đồng hành.",
      "4. Cam kết tiếp tục tiến bước lên cấp độ sơ trung cấp A2."
    ],
    cliches: [
      { ru: "Я успешно завершил курс уровня А1.", vi: "Tôi đã hoàn thành thành công khóa học cấp độ A1." },
      { ru: "Теперь я умею говорить и писать по-русски.", vi: "Giờ đây tôi đã biết nói và viết bằng tiếng Nga." },
      { ru: "Русская грамматика была непростой, но интересной.", vi: "Ngữ pháp Nga không đơn giản, nhưng rất thú vị." },
      { ru: "Я готов двигаться дальше к уровню А2!", vi: "Tôi đã sẵn sàng tiến xa hơn tới cấp độ A2!" }
    ],
    modelAnswer: "Вот и завершился мой первый большой шаг — уровень А1 по русскому языку! Когда я только начинал, я не знал даже русского алфавита. Сейчас я уже могу читать тексты, понимать русскую речь на слух, писать письма друзьям и уверенно общаться на простые темы. Изучать русский язык было непросто, но очень увлекательно. Я горжусь своими результатами и с радостью перехожу к уровню А2!",
    modelWordCount: 61
  }
];

const rubricStandard = [
  { criteria: "Bố cục & Lời chào", desc: "Có lời chào mở đầu và kết thúc phù hợp với thể loại bài viết." },
  { criteria: "Nội dung & Thông tin", desc: "Trả lời đầy đủ các yêu cầu đề bài, thông tin mạch lạc và hợp lý." },
  { criteria: "Ngữ pháp & Biến cách", desc: "Chia động từ hiện tại đúng, đại từ và danh từ phù hợp quy tắc sơ cấp." },
  { criteria: "Dung lượng từ", desc: "Đạt đúng số lượng từ yêu cầu theo chuẩn khảo thí quốc tế TRKI A1." }
];

const manifest = [];
writingLessons.forEach(l => {
  const lessonData = {
    id: l.id,
    level: "A1",
    skill: "writing",
    title: l.title,
    writingTime: "20 phút",
    topic: l.topic,
    targetWordRange: l.targetWordRange,
    task: l.task,
    requirements: l.requirements,
    cliches: l.cliches,
    modelAnswer: l.modelAnswer,
    modelWordCount: l.modelWordCount,
    rubric: rubricStandard
  };

  const fileName = `lesson-${String(l.id).padStart(2, '0')}.json`;
  const filePath = path.join(outputDir, fileName);
  fs.writeFileSync(filePath, JSON.stringify(lessonData, null, 2), 'utf8');

  manifest.push({
    id: l.id,
    file: fileName,
    title: l.title,
    badge: "A1 Viết luận",
    writingTime: "20 phút",
    topic: l.topic,
    targetWordRange: `${l.targetWordRange[0]}-${l.targetWordRange[1]} từ`
  });
});

const manifestPath = path.join(outputDir, 'manifest.json');
fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2), 'utf8');

console.log(`✓ Đã tạo thành công 20 bài học Writing A1 và manifest.json tại ${outputDir}`);

