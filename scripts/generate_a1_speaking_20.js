const fs = require('fs');
const path = require('path');

const outputDir = path.join(__dirname, '../public/data/speaking/a1');
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

const speakingLessons = [
  {
    id: 1,
    title: "Знакомство и рассказ о себе",
    topic: "Giao tiếp cơ bản: Tự giới thiệu bản thân, tên tuổi và ngành học",
    phrases: [
      { text: "Здравствуйте! Меня зовут Антон.", meaning: "Xin chào! Tôi tên là Anton.", phonetic: "Zdravstvuyte! Menya zovut Anton.", tips: "Nuốt nhẹ chữ «в» đầu tiên: /zdrastvuytye/." },
      { text: "Я студент, я учусь в университете.", meaning: "Tôi là sinh viên, tôi học ở trường đại học.", phonetic: "Ya student, ya uchus' v universitete.", tips: "Âm «сь» trong «учусь» phát âm mềm mại và êm tai." },
      { text: "Я изучаю русский язык уже один год.", meaning: "Tôi học tiếng Nga đã được một năm.", phonetic: "Ya izuchayu russkiy yazyk uzhe odin god.", tips: "Trọng âm rơi vào «язы́к», «русский» kéo dài âm /s/." },
      { text: "В свободное время я люблю читать книги.", meaning: "Vào thời gian rảnh rỗi tôi thích đọc sách.", phonetic: "V svobodnoye vremya ya lyublyu chitat' knigi.", tips: "«Люблю» môi chu tròn, âm /b/ bật nhẹ." },
      { text: "Очень приятно с вами познакомиться!", meaning: "Rất vui được làm quen với bạn!", phonetic: "Ochen' priyatno s vami poznakomit'sya!", tips: "Đuôi «-ться» trong «познакомиться» đọc như âm /-tsa/." }
    ],
    task: "Hãy tự giới thiệu về bản thân bằng tiếng Nga (3–5 câu): tên bạn, bạn làm nghề gì, sống ở đâu và bạn thích làm gì trong thời gian rảnh.",
    guidance: [
      "Меня зовут... (Tôi tên là...)",
      "Я живу в городе... (Tôi sống ở thành phố...)",
      "Я работаю... / Я учусь... (Tôi làm việc... / Tôi học...)",
      "В свободное время я люблю... (Thời gian rảnh tôi thích...)"
    ]
  },
  {
    id: 2,
    title: "Кто это? Что это? Предметы вокруг нас",
    topic: "Nhận diện người & đồ vật: Luyện câu hỏi Кто это? / Что это?",
    phrases: [
      { text: "Скажите, пожалуйста, кто это?", meaning: "Làm ơn cho tôi hỏi, đây là ai?", phonetic: "Skazhite, pozhaluysta, kto eto?", tips: "Ngữ điệu câu hỏi có từ để hỏi (IK-2), nhấn mạnh vào «кто»." },
      { text: "Это наш новый преподаватель физики.", meaning: "Đây là giảng viên môn vật lý mới của chúng tôi.", phonetic: "Eto nash novyy prepodavatel' fiziki.", tips: "«Преподаватель» nhấn trọng âm ở âm tiết cuối «-тель» mềm." },
      { text: "А что лежит на столе справа?", meaning: "Còn cái gì đang nằm ở trên bàn phía bên phải?", phonetic: "A chto lezhit na stole sprava?", tips: "«что» phát âm là /shto/, không đọc là /ch-to/." },
      { text: "Там лежат учебник, тетрадь и ручка.", meaning: "Ở đó có sách giáo khoa, vở ghi và bút mực.", phonetic: "Tam lezhat uchebnik, tetrad' i ruchka.", tips: "«Тетрадь» kết thúc bằng chữ d mềm, đọc là /t'etrat'/." },
      { text: "Чей это красивый синий рюкзак?", meaning: "Chiếc ba lô màu xanh dương đẹp đẽ này là của ai?", phonetic: "Chey eto krasivyy siniy ryukzak?", tips: "«Рюкзак» đọc biến âm /r'ugzak/, ngữ điệu nhấn ở «чей»." }
    ],
    task: "Hãy mô tả 3–4 đồ vật đang ở xung quanh bạn trong phòng học hoặc trên bàn làm việc bằng tiếng Nga.",
    guidance: [
      "Это мой рабочий стол... (Đây là bàn làm việc của tôi...)",
      "Здесь лежат книга и тетрадь... (Ở đây có sách và vở...)",
      "Справа стоит компьютер... (Bên phải có máy tính...)",
      "Это очень нужные вещи... (Đây là những vật dụng rất cần thiết...)"
    ]
  },
  {
    id: 3,
    title: "Моя семья: Родители, братья и сёстры",
    topic: "Giới thiệu gia đình: Kể về bố mẹ, anh chị em và tình cảm gia đình",
    phrases: [
      { text: "У меня большая и очень дружная семья.", meaning: "Tôi có một gia đình lớn và rất hòa thuận.", phonetic: "U menya bol'shaya i ochen' druzhnaya sem'ya.", tips: "«Семья» có dấu mềm tách âm, đọc là /s'em-ya/." },
      { text: "Мой папа работает врачом в больнице.", meaning: "Bố tôi làm bác sĩ ở bệnh viện.", phonetic: "Moy papa rabotayet vrachom v bol'nitse.", tips: "«Врачом» là cách 5 công cụ chỉ nghề nghiệp." },
      { text: "Моя мама преподаёт литературу в школе.", meaning: "Mẹ tôi giảng dạy môn văn học ở trường phổ thông.", phonetic: "Moya mama prepodayot literaturu v shkole.", tips: "Chữ «ё» trong «преподаёт» luôn luôn mang trọng âm." },
      { text: "Мой старший брат учится в институте.", meaning: "Anh trai tôi đang học tại học viện.", phonetic: "Moy starshiy brat uchitsya v institute.", tips: "«Старший» có phụ âm «ш» luôn cứng." },
      { text: "Мы все любим собираться вместе по вечерам.", meaning: "Tất cả chúng tôi thích tụ họp cùng nhau vào các buổi tối.", phonetic: "My vse lyubim sobirat'sya vmeste po vecheram.", tips: "«Вместе» phát âm chữ «е» mềm: /vm'est'e/." }
    ],
    task: "Hãy kể về gia đình của bạn (3–5 câu): Gia đình bạn có bao nhiêu người, bố mẹ bạn làm gì và cả nhà thường làm gì vào cuối tuần.",
    guidance: [
      "Моя семья состоит из четырёх человек... (Gia đình tôi gồm 4 người...)",
      "Мои родители работают... (Bố mẹ tôi làm việc...)",
      "У меня есть брат / сестра... (Tôi có anh/chị/em...)",
      "В выходные мы любим гулять... (Cuối tuần chúng tôi thích đi dạo...)"
    ]
  },
  {
    id: 4,
    title: "В ресторане: Заказ вкусного обеда",
    topic: "Giao tiếp ẩm thực: Gọi món, hỏi thực đơn và thanh toán",
    phrases: [
      { text: "Официант, принесите нам меню, пожалуйста.", meaning: "Bồi bàn ơi, làm ơn mang cho chúng tôi thực đơn.", phonetic: "Ofitsiant, prinesite nam menyu, pozhaluysta.", tips: "Từ «меню» là danh từ bất biến giống trung." },
      { text: "Что вы нам посоветуете на первое?", meaning: "Bạn gợi ý cho chúng tôi món khai vị / món súp nào?", phonetic: "Chto vy nam posovetuyete na pervoye?", tips: "«На первое» mang nghĩa là món súp đầu tiên trong bữa ăn Nga." },
      { text: "Я буду горячий борщ со сметаной.", meaning: "Tôi sẽ dùng món súp củ dền nóng với kem chua.", phonetic: "Ya budu goryachiy borshch so smetanoy.", tips: "Âm «щ» trong «борщ» là âm xuýt mềm dài: /borshch/." },
      { text: "Дайте, пожалуйста, минеральную воду без газа.", meaning: "Làm ơn cho tôi một chai nước khoáng không ga.", phonetic: "Dayte, pozhaluysta, mineral'nuyu vodu bez gaza.", tips: "«Без газа» là cách 2 sau giới từ «без»." },
      { text: "Принесите, пожалуйста, счёт. Спасибо!", meaning: "Làm ơn mang hóa đơn thanh toán cho tôi. Xin cảm ơn!", phonetic: "Prinesite, pozhaluysta, schyot. Spasibo!", tips: "«Счёт» đọc phát âm như chữ «щ»: /shchot/." }
    ],
    task: "Đóng vai khách hàng tại nhà hàng Nga, hãy gọi 1 món súp, 1 món chính, 1 đồ uống và yêu cầu tính tiền.",
    guidance: [
      "Здравствуйте! Дайте мне, пожалуйста... (Xin chào! Cho tôi xin...)",
      "На первое я хочу заказать... (Món đầu tiên tôi muốn gọi...)",
      "А пить я буду... (Còn uống thì tôi sẽ dùng...)",
      "Сколько это стоит? Принесите счёт... (Bao nhiêu tiền? Cho tôi xin hóa đơn...)"
    ]
  },
  {
    id: 5,
    title: "В магазине: Покупки и цены",
    topic: "Mua sắm thường nhật: Hỏi giá tiền, số lượng và trả lời",
    phrases: [
      { text: "Здравствуйте! Сколько стоит килограмм яблок?", meaning: "Xin chào! Một cân táo này giá bao nhiêu?", phonetic: "Zdravstvuyte! Skol'ko stoit kilogramm yablok?", tips: "Ngữ điệu hỏi giá tiền nhấn mạnh vào «Сколько»." },
      { text: "Килограмм красных яблок стоит сто рублей.", meaning: "Một cân táo đỏ có giá một trăm rúp.", phonetic: "Kilogramm krasnykh yablok stoit sto rubley.", tips: "«Рублей» là dạng cách 2 số nhiều sau số 100." },
      { text: "Взвесьте мне, пожалуйста, полкило сыра.", meaning: "Làm ơn cân giúp tôi nửa cân phô mai.", phonetic: "Vzves'te mne, pozhaluysta, polkilo syra.", tips: "«Полкило» là cách nói tắt thông dụng của nửa kilogram." },
      { text: "Можно оплатить покупку банковской картой?", meaning: "Tôi có thể thanh toán bằng thẻ ngân hàng được không?", phonetic: "Mozhno oplatit' pokupku bankovskoy kartoy?", tips: "«Банковской картой» ở cách 5 công cụ." },
      { text: "Конечно, приложите карту к терминалу.", meaning: "Tất nhiên rồi, mời bạn chạm thẻ vào máy thanh toán.", phonetic: "Konechno, prilozhite kartu k terminalu.", tips: "«Конечно» đọc biến âm «чн» thành /shn/: /kan'eshna/." }
    ],
    task: "Hãy mô tả bạn đi chợ mua hoa quả và đồ ăn: hỏi giá 2 mặt hàng và phương thức thanh toán bằng thẻ.",
    guidance: [
      "Скажите, сколько стоит... (Cho tôi hỏi, giá bao nhiêu...)",
      "Дайте мне, пожалуйста, два килограмма... (Cho tôi xin 2 cân...)",
      "У вас есть свежий хлеб? (Ở đây có bánh mì tươi không?)",
      "Я оплачу картой... (Tôi sẽ trả bằng thẻ...)"
    ]
  },
  {
    id: 6,
    title: "Мой день: Утро, день и вечер",
    topic: "Lịch trình hàng ngày: Thức dậy, làm việc, tập thể thao và nghỉ ngơi",
    phrases: [
      { text: "Обычно я просыпаюсь в семь часов утра.", meaning: "Thông thường tôi thức dậy lúc bảy giờ sáng.", phonetic: "Obychno ya prosypayus' v sem' chasov utra.", tips: "«Обычно» mang nghĩa thói quen lặp lại thường xuyên." },
      { text: "Я чищу зубы, умываюсь и делаю зарядку.", meaning: "Tôi đánh răng, rửa mặt và tập thể dục buổi sáng.", phonetic: "Ya chishchu zuby, umyvayus' i delayu zaryadku.", tips: "«Чищу» biến âm /chishchu/, «умываюсь» âm mềm." },
      { text: "В восемь часов я завтракаю и иду на работу.", meaning: "Lúc tám giờ tôi ăn sáng và đi làm việc.", phonetic: "V vosem' chasov ya zavtrakayu i idu na rabotu.", tips: "«Иду на работу» dùng giới từ «на» với công việc." },
      { text: "После работы я хожу в тренажёрный зал.", meaning: "Sau giờ làm tôi thường đến phòng tập thể hình.", phonetic: "Posle raboty ya khozhu v trenazhyornyy zal.", tips: "«Тренажёрный» trọng âm vào «ё»." },
      { text: "Вечером я отдыхаю и ложусь спать в одиннадцать.", meaning: "Buổi tối tôi nghỉ ngơi và đi ngủ lúc mười một giờ.", phonetic: "Vecherom ya otdykhayu i lozhus' spat' v odinnadtsat'.", tips: "«Ложусь спать» đọc liền mạch tự nhiên." }
    ],
    task: "Kể ngắn gọn về một ngày điển hình của bạn từ sáng đến tối (3–5 câu).",
    guidance: [
      "Утром я встаю в... (Buổi sáng tôi thức dậy lúc...)",
      "Днём я работаю / учусь... (Ban ngày tôi làm việc / học...)",
      "В час дня я обедаю... (Lúc 1 giờ trưa tôi ăn trưa...)",
      "Вечером я люблю... (Buổi tối tôi thích...)"
    ]
  },
  {
    id: 7,
    title: "В городе: Как пройти в центр?",
    topic: "Hỏi đường & Phương hướng: Đi thẳng, rẽ trái, rẽ phải, trạm xe buýt",
    phrases: [
      { text: "Извините, скажите, где станция метро?", meaning: "Xin lỗi, cho tôi hỏi trạm tàu điện ngầm ở đâu?", phonetic: "Izvinite, skazhite, gde stantsiya metro?", tips: "«Извините» dùng lịch sự để bắt đầu hỏi đường." },
      { text: "Идите прямо до перекрёстка и поверните направо.", meaning: "Hãy đi thẳng đến ngã tư rồi rẽ sang bên phải.", phonetic: "Idite pryamo do perekryostka i povernite napravo.", tips: "«Направо» = rẽ phải, trọng âm âm tiết 2." },
      { text: "Большой театр находится на противоположной стороне.", meaning: "Nhà hát Bolshoi nằm ở phía đối diện bên kia đường.", phonetic: "Bol'shoy teatr nakhoditsya na protivopolozhnoy storone.", tips: "«Стороне» trọng âm ở chữ e cuối cùng." },
      { text: "Это далеко отсюда или можно дойти пешком?", meaning: "Chỗ đó có xa đây không hay có thể đi bộ tới được?", phonetic: "Eto daleko otsyuda ili mozhno doyti peshkom?", tips: "«Пешком» nghĩa là đi bằng hai chân (đi bộ)." },
      { text: "Пешком всего пять минут, это совсем рядом.", meaning: "Đi bộ chỉ mất năm phút thôi, ngay gần đây thôi.", phonetic: "Peshkom vsego pyat' minut, eto sovsem ryadom.", tips: "«Совсем рядом» = cực kỳ gần." }
    ],
    task: "Đóng vai người hỏi đường đến một viện bảo tàng hoặc quảng trường nổi tiếng tại Nga.",
    guidance: [
      "Скажите, пожалуйста, как доехать до... (Làm ơn chỉ giúp cách đi đến...)",
      "Где находится остановка автобуса? (Trạm dừng xe buýt ở đâu?)",
      "Сколько минут ехать на метро? (Đi tàu điện ngầm mất mấy phút?)",
      "Большое спасибо за помощь! (Cảm ơn bạn rất nhiều vì đã giúp đỡ!)"
    ]
  },
  {
    id: 8,
    title: "Времена года и погода",
    topic: "Thời tiết và bốn mùa: Miêu tả nắng mưa, tuyết trắng và nhiệt độ",
    phrases: [
      { text: "Какая сегодня прекрасная тёплая погода!", meaning: "Thời tiết hôm nay thật ấm áp và tuyệt vời làm sao!", phonetic: "Kakaya segodnya prekrasnaya tyoplaya pogoda!", tips: "«Сегодня» phát âm là /s'evodn'a/ (chữ «г» đọc như «в»)." },
      { text: "Летом в Москве тепло и светит яркое солнце.", meaning: "Vào mùa hè ở Moskva ấm áp và mặt trời chiếu sáng rực rỡ.", phonetic: "Letom v Moskve teplo i svetit yarkoye solntse.", tips: "«Солнце» nuốt chữ «л», phát âm là /sontse/." },
      { text: "Зимой часто идёт снег и дует холодный ветер.", meaning: "Mùa đông tuyết thường rơi và gió lạnh thổi mạnh.", phonetic: "Zimoy chasto idyot sneg i duyet kholodnyy veter.", tips: "«Снег» tận cùng đọc biến âm /sn'ek/." },
      { text: "Сегодня на улице минус десять градусов мороза.", meaning: "Hôm nay ngoài trời âm mười độ lạnh cóng.", phonetic: "Segodnya na ulitse minus desyat' gradusov moroza.", tips: "«Градусов» là cách 2 số nhiều sau số 10." },
      { text: "Осенью идут дожди, а деревья становятся жёлтыми.", meaning: "Mùa thu thường có mưa và cây cối ngả sang màu vàng óng.", phonetic: "Osen'yu idut dozhdi, a derev'ya stanovyatsya zhyoltymi.", tips: "«Осенью» có dấu mềm tách âm /os'en'yu/." }
    ],
    task: "Mô tả thời tiết của mùa bạn thích nhất trong năm bằng tiếng Nga (3–4 câu).",
    guidance: [
      "Моё любимое время года — это... (Mùa yêu thích của tôi là...)",
      "В это время года на улице... (Vào mùa này ngoài trời...)",
      "Светит солнце / идёт снег... (Trời có nắng / tuyết rơi...)",
      "Я очень люблю гулять, потому что... (Tôi rất thích đi dạo vì...)"
    ]
  },
  {
    id: 9,
    title: "В гостинице: Размещение и сервис",
    topic: "Lưu trú khách sạn: Nhận chìa khóa phòng, hỏi giờ ăn sáng và tiện ích",
    phrases: [
      { text: "Добрый день! У меня забронирован номер.", meaning: "Chào bạn! Tôi đã đặt trước một phòng.", phonetic: "Dobryy den'! U menya zabronirovan nomer.", tips: "«Забронирован» nhấn trọng âm ở âm tiết «-ни́-»." },
      { text: "Вот мой заграничный паспорт и подтверждение брони.", meaning: "Đây là hộ chiếu và giấy xác nhận đặt phòng của tôi.", phonetic: "Vot moy zagranichnyy pasport i podtverzhdeniye broni.", tips: "«Подтверждение» phát âm rõ /pattv'erzhd'en'iye/." },
      { text: "На каком этаже находится мой номер?", meaning: "Phòng của tôi nằm ở tầng mấy vậy?", phonetic: "Na kakom etazhe nakhoditsya moy nomer?", tips: "«На каком этаже» dùng cách 6 với giới từ «на»." },
      { text: "Во сколько начинается утренний завтрак?", meaning: "Bữa điểm tâm sáng bắt đầu từ mấy giờ?", phonetic: "Vo skol'ko nachinayetsya utrenniy zavtrak?", tips: "«Во сколько» là câu hỏi giờ giấc thông dụng nhất." },
      { text: "В номере есть кондиционер и фен?", meaning: "Trong phòng có máy lạnh và máy sấy tóc không?", phonetic: "V nomere yest' konditsioner i fen?", tips: "«Фен» phát âm âm /e/ mở nhẹ." }
    ],
    task: "Thực hành đoạn hội thoại làm thủ tục check-in tại quầy lễ tân khách sạn Nga.",
    guidance: [
      "Здравствуйте, я хочу зарегистрироваться... (Chào bạn, tôi muốn làm thủ tục nhận phòng...)",
      "Мой номер одноместный... (Phòng của tôi là phòng đơn...)",
      "Где находится лифт? (Thang máy nằm ở đâu?)",
      "Спасибо за теплый приём! (Cảm ơn vì sự đón tiếp nồng hậu!)"
    ]
  },
  {
    id: 10,
    title: "Мои увлечения и хобби",
    topic: "Sở thích cá nhân: Nghe nhạc, chơi đàn guitar, thể thao và du lịch",
    phrases: [
      { text: "Какое у тебя самое главное хобби?", meaning: "Sở thích lớn nhất của bạn là gì?", phonetic: "Kakoye u tebya samoye glavnoye khobbi?", tips: "«Хобби» là từ ngoại lai giống trung, không biến cách." },
      { text: "Я очень люблю играть на классической гитаре.", meaning: "Tôi rất thích chơi đàn guitar cổ điển.", phonetic: "Ya ochen' lyublyu igrat' na klassicheskoy gitare.", tips: "Chơi nhạc cụ dùng cấu trúc «играть на + Cách 6»." },
      { text: "По выходным я играю с друзьями в футбол.", meaning: "Vào cuối tuần tôi chơi bóng đá với bạn bè.", phonetic: "Po vykhodnym ya igrayu s druz'yami v futbol.", tips: "Chơi thể thao dùng cấu trúc «играть в + Cách 4»." },
      { text: "Моя подруга увлекается фотографией и живописью.", meaning: "Bạn gái tôi say mê môn nhiếp ảnh và hội họa.", phonetic: "Moya podruga uvlekayetsya fotografiyey i zhivopis'yu.", tips: "«Увлекается» đi với Cách 5 công cụ." },
      { text: "Путешествия по миру дарят мне море вдохновения.", meaning: "Những chuyến du lịch vòng quanh thế giới cho tôi vô vàn cảm hứng.", phonetic: "Puteshestviya po miru daryat mne more vdokhnoveniya.", tips: "«Вдохновения» có trọng âm ở âm tiết «-ве́-»." }
    ],
    task: "Kể về 2 sở thích cá nhân mà bạn yêu thích nhất khi có thời gian rảnh rỗi.",
    guidance: [
      "В свободное время я увлекаюсь... (Thời gian rảnh tôi say mê...)",
      "Я люблю слушать русскую музыку... (Tôi thích nghe nhạc Nga...)",
      "Это помогает мне отдохнуть... (Điều này giúp tôi thư giãn...)",
      "А чем увлекаешься ты? (Còn bạn say mê điều gì?)"
    ]
  },
  {
    id: 11,
    title: "В аптеке: Описание симптомов",
    topic: "Sức khỏe: Nói về triệu chứng cảm sốt, đau họng, mua thuốc",
    phrases: [
      { text: "Здравствуйте! У меня сильно болит горло и голова.", meaning: "Chào dược sĩ! Tôi bị đau họng và đau đầu dữ dội.", phonetic: "Zdravstvuyte! U menya sil'no bolit gorlo i golova.", tips: "«Болит» dùng cho danh từ số ít (болит горло, болит голова)." },
      { text: "Какая у вас сейчас температура тела?", meaning: "Hiện tại thân nhiệt của bạn là bao nhiêu độ?", phonetic: "Kakaya u vas seychas temperatura tela?", tips: "Trọng âm «температу́ра»." },
      { text: "У меня кашель и сильный насморк.", meaning: "Tôi bị ho và sổ mũi rất nặng.", phonetic: "U menya kashel' i sil'nyy nasmork.", tips: "«Кашель» khi biến cách bị nuốt nguyên âm «е» (нет кашля)." },
      { text: "Дайте мне эффективное лекарство от простуды.", meaning: "Hãy cho tôi một loại thuốc hiệu quả trị cảm lạnh.", phonetic: "Dayte mne effektivnoye lekarstvo ot prostudy.", tips: "«От простуды» dùng Cách 2 sau giới từ «от»." },
      { text: "Принимайте этот сироп по чайной ложке после еды.", meaning: "Hãy uống si-rô này mỗi lần một thìa cà phê sau bữa ăn.", phonetic: "Prinimayte etot sirop po chaynoy lozhke posle edy.", tips: "«По чайной ложке» = mỗi lần một muỗng cà phê." }
    ],
    task: "Đóng vai người bệnh mô tả triệu chứng sốt và nhờ dược sĩ tư vấn thuốc phù hợp.",
    guidance: [
      "Я плохо себя чувствую... (Tôi cảm thấy không được khỏe...)",
      "У меня температура тридцать восемь... (Tôi bị sốt 38 độ...)",
      "Посоветуйте хорошее средство... (Hãy khuyên tôi phương thuốc tốt...)",
      "Спасибо, я буду лечиться! (Cảm ơn, tôi sẽ điều trị theo lời dặn!)"
    ]
  },
  {
    id: 12,
    title: "На вокзале: Поездка на поезде",
    topic: "Đi lại: Mua vé tàu hỏa, hỏi giờ khởi hành và tìm toa xe",
    phrases: [
      { text: "Скажите, во сколько отправляется поезд в Казань?", meaning: "Cho tôi hỏi tàu hỏa đi Kazan khởi hành lúc mấy giờ?", phonetic: "Skazhite, vo skol'ko otpravlyayetsya poyezd v Kazan'?", tips: "«Отправля́ется» nhấn trọng âm ở âm tiết thứ 3." },
      { text: "Мне нужен один билет в купейный вагон.", meaning: "Tôi cần một vé ở toa giường nằm khoang 4 chỗ.", phonetic: "Mne nuzhen odin bilet v kupeynyy vagon.", tips: "«Купейный вагон» là toa tàu có phòng khép kín thoải mái." },
      { text: "Какое у вас место: нижнее или верхнее?", meaning: "Chỗ nằm của bạn là giường tầng dưới hay tầng trên?", phonetic: "Kakoye u vas mesto: nizhneye ili verkhneye?", tips: "«Нижнее» (tầng dưới) luôn được khách đi tàu ưa chuộng hơn." },
      { text: "С какого пути и с какой платформы отходит состав?", meaning: "Đoàn tàu xuất phát từ đường ray số mấy và sân ga nào?", phonetic: "S kakogo puti i s kakoy platformy otkhodit sostav?", tips: "«Пути» là cách 2 bất quy tắc của danh từ giống đực «путь»." },
      { text: "Желаем вам приятной и лёгкой дороги!", meaning: "Chúc quý khách có một chuyến hành trình êm ái và vui vẻ!", phonetic: "Zhelayem vam priyatnoy i lyogkoy dorogi!", tips: "«Лёгкой» phát âm là /lyokhkay/ (biến âm «г» trước «к»)." }
    ],
    task: "Hỏi mua 1 vé tàu đi Saint Petersburg và hỏi số đường ray tàu xuất phát.",
    guidance: [
      "Здравствуйте! Дайте билет до... (Xin chào! Cho tôi vé đi đến...)",
      "Поезд отходит утром или вечером? (Tàu chạy buổi sáng hay tối?)",
      "Сколько часов ехать? (Đi hết bao nhiêu tiếng?)",
      "Где находится справочное бюро? (Bàn thông tin ở đâu?)"
    ]
  },
  {
    id: 13,
    title: "В гостях: Русский чай и традиции",
    topic: "Văn hóa tiếp khách: Đến thăm nhà, uống trà samovar và khen ngợi",
    phrases: [
      { text: "Проходите в дом, чувствуйте себя как дома!", meaning: "Mời vào trong nhà, bạn hãy cứ tự nhiên như ở nhà mình nhé!", phonetic: "Prokhodite v dom, chuvstvuyte sebya kak doma!", tips: "«Чувствуйте» đọc nuốt âm «в» đầu tiên: /chustvuyte/." },
      { text: "Спасибо за приглашение! Это вам цветы и торт.", meaning: "Cảm ơn vì lời mời! Đây là hoa và bánh kem gửi tặng gia đình.", phonetic: "Spasibo za priglasheniye! Eto vam tsvety i tort.", tips: "Người Nga có văn hóa mang quà nhỏ khi đến chơi nhà." },
      { text: "Садитесь за стол, чайник уже закипел.", meaning: "Mời ngồi vào bàn, ấm nước đã sôi rồi đấy.", phonetic: "Sadites' za stol, chaynink uzhe zakipel.", tips: "«Чайник» đọc êm dịu, trọng âm «ча́йник»." },
      { text: "Попробуйте домашние пироги с яблоками и мёдом.", meaning: "Bạn hãy nếm thử món bánh pỉ-rốc táo mật ong nhà làm nhé.", phonetic: "Poprobuyte domashniye pirogi s yablokami i myodom.", tips: "«Пироги́» có trọng âm ở chữ «и» cuối." },
      { text: "Всё очень вкусно, вы замечательная хозяйка!", meaning: "Mọi món đều rất ngon, bác thực sự là một nữ chủ nhà tuyệt vời!", phonetic: "Vsyo ochen' vkusno, vy zamechatel'naya khozyayka!", tips: "Lời khen ngợi lịch thiệp dành cho nữ chủ nhà." }
    ],
    task: "Nói lời cảm ơn và khen ngợi lòng mến khách khi được mời đến chơi nhà một người bạn Nga.",
    guidance: [
      "Большое спасибо за тёплый приём... (Cảm ơn nhiều vì sự đón tiếp ấm áp...)",
      "У вас очень красивый и уютный дом... (Nhà bạn rất đẹp và ấm cúng...)",
      "Пирог был невероятно вкусным... (Bánh pỉ-rốc ngon không thể tin nổi...)",
      "Приглашаю вас в гости ко мне! (Tôi mời bạn dịp tới ghé thăm nhà tôi!)"
    ]
  },
  {
    id: 14,
    title: "В банке: Финансы и карты",
    topic: "Ngân hàng: Mở thẻ thanh toán, rút tiền ATM và đổi tiền",
    phrases: [
      { text: "Здравствуйте! Я хочу открыть счёт в банке.", meaning: "Xin chào! Tôi muốn mở một tài khoản tại ngân hàng.", phonetic: "Zdravstvuyte! Ya khochu otkryt' schyot v banke.", tips: "«Открыть счёт» = mở tài khoản thanh toán." },
      { text: "Заполните, пожалуйста, это заявление и анкету.", meaning: "Làm ơn điền vào mẫu đơn đề nghị và tờ khai này.", phonetic: "Zapolnite, pozhaluysta, eto zayavleniye i anketu.", tips: "«Заявление» là đơn từ hành chính." },
      { text: "Где находится ближайший банкомат для снятия наличных?", meaning: "Cây ATM gần nhất để rút tiền mặt nằm ở đâu vậy?", phonetic: "Gde nakhoditsya blizhayshiy bankomat dlya snyatiya nalichnykh?", tips: "«Снятия наличных» = rút tiền mặt." },
      { text: "Комиссия за перевод денег отсутствует.", meaning: "Giao dịch chuyển tiền hoàn toàn miễn phí hoa hồng.", phonetic: "Komissiya za perevod deneg otsutstvuyet.", tips: "«Отсутствует» nuốt âm mềm, phát âm là /atsutstvuyet/." },
      { text: "Ваша карта будет готова через три рабочих дня.", meaning: "Thẻ của quý khách sẽ hoàn thành sau ba ngày làm việc.", phonetic: "Vasha karta budet gotova cherez tri rabochikh dnya.", tips: "«Рабочих дня» là cách 2 số ít sau số 3." }
    ],
    task: "Đóng vai khách hàng hỏi nhân viên ngân hàng về cách rút tiền tại cây ATM và phí dịch vụ.",
    guidance: [
      "Скажите, как работает этот банкомат? (Cho tôi hỏi máy ATM này hoạt động thế nào?)",
      "Какая комиссия за снятие? (Phí rút tiền là bao nhiêu?)",
      "Можно ли перевести деньги через телефон? (Có chuyển tiền qua điện thoại được không?)",
      "Благодарю за консультацию! (Cảm ơn bạn đã tư vấn!)"
    ]
  },
  {
    id: 15,
    title: "В кинотеатре: Покупка билетов",
    topic: "Giải trí: Chọn phim, chọn hàng ghế và mua bỏng ngô",
    phrases: [
      { text: "Какой интересный фильм сейчас идёт на экране?", meaning: "Bộ phim hấp dẫn nào đang được chiếu trên màn ảnh hiện nay?", phonetic: "Kakoy interesnyy fil'm seychas idyot na ekrane?", tips: "«Идёт фильм» là thành ngữ chỉ phim đang chiếu rạp." },
      { text: "Дайте, пожалуйста, два билета на вечерний сеанс.", meaning: "Làm ơn cho tôi hai vé xem suất chiếu buổi tối.", phonetic: "Dayte, pozhaluysta, dva bileta na vecherniy seans.", tips: "«Сеанс» mượn từ tiếng Pháp, chỉ suất chiếu." },
      { text: "Выберите места: в центре зала, седьмой ряд.", meaning: "Hãy chọn chỗ ngồi: ở giữa phòng chiếu, hàng ghế số 7.", phonetic: "Vyberite mesta: v tsentre zala, sed'moy ryad.", tips: "«Ряд» mang nghĩa hàng ghế." },
      { text: "Мы хотим купить большой попкорн и колу со льдом.", meaning: "Chúng tôi muốn mua một hộp bắp rang bơ lớn và nước ngọt có đá.", phonetic: "My khotim kupit' bol'shoy popkorn i kolu so l'dom.", tips: "«Со льдом» là cách 5 sau giới từ «со»." },
      { text: "Приятного вам просмотра фильма!", meaning: "Chúc quý khách xem phim vui vẻ!", phonetic: "Priyatnogo vam prosmotra fil'ma!", tips: "Lời chúc phổ biến trước khi vào rạp chiếu." }
    ],
    task: "Mời bạn đi xem phim và nói về bộ phim bạn muốn xem cùng thời gian suất chiếu.",
    guidance: [
      "Пойдём сегодня вечером в кино! (Tối nay chúng mình đi xem phim đi!)",
      "Там идёт новая комедия... (Ở đó đang chiếu bộ phim hài mới...)",
      "Сеанс начинается в восемь вечера... (Suất chiếu bắt đầu lúc 8 giờ tối...)",
      "Я куплю билеты онлайн... (Tôi sẽ mua vé trực tuyến...)"
    ]
  },
  {
    id: 16,
    title: "В университете: На занятиях",
    topic: "Học tập: Giảng đường, giờ học пара, kỳ thi và giảng viên",
    phrases: [
      { text: "Во сколько начинается наша первая пара?", meaning: "Tiết học đôi (ca học) đầu tiên của chúng ta bắt đầu lúc mấy giờ?", phonetic: "Vo skol'ko nachinayetsya nasha pervaya para?", tips: "«Пара» trong trường đại học Nga là một ca học gồm 2 tiết (90 phút)." },
      { text: "Сегодня у нас лекция по русской грамматике.", meaning: "Hôm nay chúng ta có buổi nghe giảng về ngữ pháp tiếng Nga.", phonetic: "Segodnya u nas lektsiya po russkoy grammatike.", tips: "«Лекция по + Cách 3»." },
      { text: "Профессор объясняет сложную тему очень понятно.", meaning: "Giáo sư giảng giải chủ đề phức tạp rất dễ hiểu.", phonetic: "Professor ob'yasnyayet slozhnuyu temu ochen' ponyatno.", tips: "Dấu cứng «ъ» trong «объясняет» ngắt âm rành mạch: /ab-yasnyayet/." },
      { text: "Я готовлюсь к зимней сессии и учу новые слова.", meaning: "Tôi đang chuẩn bị cho kỳ thi học kỳ mùa đông và học từ vựng mới.", phonetic: "Ya gotovlyus' k zimney sessii i uchu novyye slova.", tips: "«Сессия» là đợt thi học kỳ tập trung tại các trường Nga." },
      { text: "Желаю тебе успешно сдать все зачёты и экзамены!", meaning: "Chúc bạn vượt qua thành công tất cả các bài kiểm tra và kỳ thi!", phonetic: "Zhelayu tebe uspeshno sdat' vse zachyoty i ekzameny!", tips: "«Зачёт» là bài kiểm tra điều kiện đạt/không đạt (pass/fail)." }
    ],
    task: "Chia sẻ về một ngày học tập tại trường đại học của bạn bằng tiếng Nga.",
    guidance: [
      "Каждый день у нас три пары... (Mỗi ngày chúng tôi học 3 ca...)",
      "Мы внимательно слушаем преподавателя... (Chúng tôi chăm chú lắng nghe giảng viên...)",
      "После пар мы идём в библиотеку... (Sau giờ học chúng tôi vào thư viện...)",
      "Учиться трудно, но очень интересно! (Học tuy khó nhưng vô cùng thú vị!)"
    ]
  },
  {
    id: 17,
    title: "В аэропорту: Регистрация на рейс",
    topic: "Du lịch hàng không: Cân hành lý, vé máy bay và qua cửa an ninh",
    phrases: [
      { text: "Предъявите, пожалуйста, ваш паспорт и билет.", meaning: "Làm ơn xuất trình hộ chiếu và vé máy bay của quý khách.", phonetic: "Pred'yavite, pozhaluysta, vash pasport i bilet.", tips: "«Предъявите» có dấu cứng «ъ», đọc ngắt giọng." },
      { text: "Поставьте ваш чемодан на весы для взвешивания.", meaning: "Mời bạn đặt va ly lên bàn cân để kiểm tra trọng lượng.", phonetic: "Postav'te vash chemodan na vesy dlya vzveshivaniya.", tips: "«Весы́» mang nghĩa cái cân, trọng âm âm tiết 2." },
      { text: "У вас нет перевеса, багаж весит двадцать килограммов.", meaning: "Hành lý không bị quá cước, nặng đúng hai mươi kilogram.", phonetic: "U vas net perevesa, bagazh vesit dvadtsat' kilogrammov.", tips: "«Перевес» = hành lý quá cân quy định." },
      { text: "Вот ваш посадочный талон. Посадка у выхода номер шесть.", meaning: "Đây là thẻ lên máy bay. Cửa khởi hành là cửa số sáu.", phonetic: "Vot vash posadochnyy talon. Posadka u vykhoda nomer shest'.", tips: "«Посадочный талон» = Boarding pass." },
      { text: "Желаем вам приятного полёта и мягкой посадки!", meaning: "Chúc quý khách có một chuyến bay êm đẹp và hạ cánh an toàn!", phonetic: "Zhelayem vam priyatnogo polyota i myagkoy posadki!", tips: "«Мягкой» biến âm «г» thành /kh/: /m'yakhkay/." }
    ],
    task: "Thực hành hội thoại với nhân viên quầy thủ tục sân bay khi gửi hành lý.",
    guidance: [
      "Здравствуйте, я лечу в Ханой... (Xin chào, tôi bay về Hà Nội...)",
      "Можно место возле иллюминатора? (Cho tôi chỗ ngồi cạnh cửa sổ được không?)",
      "Сколько мест багажа разрешено? (Được gửi mấy kiện hành lý?)",
      "Спасибо, всего доброго! (Cảm ơn, chúc mọi điều tốt lành!)"
    ]
  },
  {
    id: 18,
    title: "Спорт: Здоровый образ жизни",
    topic: "Lối sống lành mạnh: Chạy bộ, tập yoga, ăn uống bổ dưỡng",
    phrases: [
      { text: "Я веду активный и здоровый образ жизни.", meaning: "Tôi duy trì lối sống năng động và lành mạnh.", phonetic: "Ya vedu aktivnyy i zdorovyy obraz zhizni.", tips: "Cụm từ cố định «здоровый образ жизни» (ЗОЖ)." },
      { text: "Каждое утро я бегаю в парке полчаса.", meaning: "Mỗi buổi sáng tôi chạy bộ trong công viên nửa tiếng.", phonetic: "Kazhdoye utro ya begayu v parke polchasa.", tips: "«Полчаса» = nửa giờ đồng hồ." },
      { text: "Правильное питание очень важно для здоровья.", meaning: "Chế độ dinh dưỡng đúng đắn rất quan trọng đối với sức khỏe.", phonetic: "Pravil'noye pitaniye ochen' vazhno dlya zdorov'ya.", tips: "«Питание» = dinh dưỡng, ăn uống." },
      { text: "Я стараюсь пить два литра чистой воды в день.", meaning: "Tôi cố gắng uống hai lít nước tinh khiết mỗi ngày.", phonetic: "Ya starayus' pit' dva litra chistoy vody v den'.", tips: "«Два литра» là cách 2 số ít sau số 2." },
      { text: "Спорт помогает мне быть бодрым и энергичным.", meaning: "Thể thao giúp tôi luôn sảng khoái và tràn đầy năng lượng.", phonetic: "Sport pomogayet mne byt' bodrym i energichnym.", tips: "«Энергичным» đọc phụ âm mềm «г» /en'erg'ichnym/." }
    ],
    task: "Kể về các thói quen rèn luyện sức khỏe mỗi ngày của bạn bằng tiếng Nga.",
    guidance: [
      "Для меня спорт очень важен... (Đối với tôi thể thao rất quan trọng...)",
      "Я люблю плавать и кататься на велосипеде... (Tôi thích bơi và đi xe đạp...)",
      "Я не ем вредную еду... (Tôi không ăn đồ ăn có hại cho sức khỏe...)",
      "Чувствую себя отлично! (Tôi cảm thấy cơ thể tuyệt vời!)"
    ]
  },
  {
    id: 19,
    title: "Праздники в России: Новый год и Масленица",
    topic: "Lễ hội truyền thống: Đón năm mới, tiễn mùa đông Maslenitsa",
    phrases: [
      { text: "Новый год — самый любимый праздник в России.", meaning: "Năm Mới là ngày lễ được yêu thích nhất ở nước Nga.", phonetic: "Novyy god — samyy lyubimyy prazdnik v Rossii.", tips: "«Праздник» nuốt chữ «д», phát âm là /prazn'ik/." },
      { text: "Мы украшаем ёлку яркими стеклянными игрушками.", meaning: "Chúng tôi trang hoàng cây thông bằng những quả cầu thủy tinh rực rỡ.", phonetic: "My ukrashayem yolku yarkimi steklyannymi igrushkami.", tips: "«Украшаем» nhấn trọng âm ở âm tiết thứ 3." },
      { text: "На Масленицу все пекут золотистые круглые блины.", meaning: "Vào lễ Maslenitsa mọi người đều rán những chiếc bánh xèo blini tròn vàng óng.", phonetic: "Na Maslenitsu vse pekut zolotistyye kruglyye bliny.", tips: "«Блины» tượng trưng cho vầng thái dương ấm áp." },
      { text: "Дети с радостью получают новогодние подарки.", meaning: "Trẻ nhỏ vui sướng nhận những phần quà mừng năm mới.", phonetic: "Deti s radost'yu poluchayut novogodniye podarki.", tips: "«С радостью» mang ý nghĩa ngập tràn hân hoan." },
      { text: "Поздравляю вас от всей души с праздником!", meaning: "Tôi xin gửi lời chúc mừng ngày lễ chân thành nhất tự đáy lòng!", phonetic: "Pozdravlyayu vas ot vsey dushi s prazdnikom!", tips: "Lời chúc trang trọng và ấm cúng nhất của người Nga." }
    ],
    task: "Kể về một ngày lễ truyền thống mà bạn thích nhất (Tết cổ truyền hoặc Tết Năm Mới).",
    guidance: [
      "Мой любимый праздник — это Новый год... (Lễ hội yêu thích của tôi là Năm Mới...)",
      "В этот день мы собираемся семьёй... (Vào ngày này cả gia đình tụ họp...)",
      "Мы готовим вкусные традиционные блюда... (Chúng tôi nấu những món ăn truyền thống ngon...)",
      "Все дарят друг другу подарки... (Mọi người tặng quà cho nhau...)"
    ]
  },
  {
    id: 20,
    title: "Итоговая самопрезентация: Я говорю по-русски!",
    topic: "Vấn đáp A1 tổng kết: Thuyết trình toàn diện về bản thân và ước mơ",
    phrases: [
      { text: "Здравствуйте, уважаемые экзаменаторы!", meaning: "Kính chào quý vị giám khảo đáng kính!", phonetic: "Zdravstvuyte, uvazhayemyye ekzamenatory!", tips: "Mở đầu bài thi vấn đáp trang trọng và đĩnh đạc." },
      { text: "Сегодня я с радостью хочу рассказать о себе.", meaning: "Hôm nay em rất vui mừng được trình bày về bản thân mình.", phonetic: "Segodnya ya s radost'yu khochu rasskazat' o sebe.", tips: "«Рассказать о себе» dùng Cách 6." },
      { text: "Я успешно изучил все темы начального уровня А1.", meaning: "Em đã hoàn thành xuất sắc tất cả các chủ đề của trình độ sơ cấp A1.", phonetic: "Ya uspeshno izuchil vse temy nachal'nogo urovnya A1.", tips: "«Уровня» là dạng cách 2 của «уровень»." },
      { text: "Я могу свободно общаться в простых жизненных ситуациях.", meaning: "Em có thể giao tiếp tự tin trong các tình huống đời sống thường nhật.", phonetic: "Ya mogu svobodno obshchat'sya v prostykh zhiznennykh situatsiyakh.", tips: "«Общаться» đuôi mềm «-ться» phát âm /-tsa/." },
      { text: "Я продолжаю учить русский язык и мечтаю поехать в Россию!", meaning: "Em sẽ tiếp tục học tiếng Nga và mơ ước một ngày được đặt chân đến nước Nga!", phonetic: "Ya prodolzhayu uchit' russkiy yazyk i mechtayu poyekhat' v Rossiyu!", tips: "«В Россию» dùng Cách 4 chỉ phương hướng sau «поехать»." }
    ],
    task: "Bài thi nói tổng kết A1: Hãy thực hiện bài thuyết trình tự giới thiệu bản thân trong 1 phút (tên, xuất thân, quá trình học tiếng Nga và lý do bạn yêu thích nước Nga).",
    guidance: [
      "Меня зовут... Я приехал из Вьетнама... (Tên em là... Em đến từ Việt Nam...)",
      "Я изучаю русский язык, потому что... (Em học tiếng Nga bởi vì...)",
      "Русский язык очень богатый и красивый... (Tiếng Nga rất phong phú và tuyệt đẹp...)",
      "Спасибо за внимание и поддержку! (Xin cảm ơn sự chú ý lắng nghe và đồng hành của quý thầy cô!)"
    ]
  }
];

const manifest = [];
speakingLessons.forEach(l => {
  const lessonData = {
    id: l.id,
    level: "A1",
    skill: "speaking",
    title: l.title,
    speakingTime: "15 phút",
    topic: l.topic,
    intro: "Luyện phát âm chuẩn xác từng âm tiết tiếng Nga với công nghệ AI Web Speech Recognition. Nhấn nút Loa 🔊 để nghe phát âm chuẩn, sau đó nhấn nút Mic 🎙️ để phát âm và nhận phản hồi chấm điểm trực tiếp!",
    shadowingPhrases: l.phrases.map((p, idx) => ({
      id: `spk_${idx + 1}`,
      text: p.text,
      meaning: p.meaning,
      phonetic: p.phonetic,
      tips: p.tips
    })),
    freeSpeakingPrompt: {
      title: "Luyện nói tự do (Монолог)",
      task: l.task,
      sampleGuidance: l.guidance
    }
  };

  const fileName = `lesson-${String(l.id).padStart(2, '0')}.json`;
  const filePath = path.join(outputDir, fileName);
  fs.writeFileSync(filePath, JSON.stringify(lessonData, null, 2), 'utf8');

  manifest.push({
    id: l.id,
    file: fileName,
    title: l.title,
    badge: "A1 Nói & Vấn đáp",
    speakingTime: "15 phút",
    topic: l.topic,
    phraseCount: l.phrases.length
  });
});

const manifestPath = path.join(outputDir, 'manifest.json');
fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2), 'utf8');

console.log(`✓ Đã tạo thành công 20 bài học Speaking A1 và manifest.json tại ${outputDir}`);
