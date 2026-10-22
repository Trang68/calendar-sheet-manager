/**
 * TAIGA TALK — 25 CHỦ ĐỀ TỪ VỰNG & FLASHCARDS TIẾNG NGA
 * Hệ thống 25 chủ đề thực tế từ A1 đến B2 với phiên âm, từ loại, nghĩa tiếng Việt và câu ví dụ.
 */

window.VOCAB_25_TOPICS = [
  {
    id: "topic-1",
    num: "01",
    icon: "💬",
    titleVi: "Chào hỏi & Làm quen",
    titleRu: "Приветствия и знакомство",
    desc: "Các mẫu câu chào hỏi trang trọng, thân mật, tự giới thiệu tên tuổi, quốc tịch và nghề nghiệp.",
    count: 15,
    cards: [
      {
        word: "здравствуйте",
        ipa: "[ˈzdrastvʊjtʲe]",
        phonetic: "zdrast-vuy-tye",
        gram: "Thán từ (chào trang trọng)",
        meaning: "Xin chào (lịch sự, kính trọng)",
        exampleRu: "Здравствуйте! Меня зовут Анна.",
        exampleVi: "Xin chào! Tôi tên là Anna."
      },
      {
        word: "привет",
        ipa: "[prʲɪˈvʲet]",
        phonetic: "pri-vyet",
        gram: "Thán từ (thân mật)",
        meaning: "Chào bạn (dùng với bạn bè, người quen)",
        exampleRu: "Привет, Антон! Как дела?",
        exampleVi: "Chào Anton! Dạo này thế nào?"
      },
      {
        word: "доброе утро",
        ipa: "[ˈdobrəjə ˈutrə]",
        phonetic: "dô-bra-ye u-tra",
        gram: "Cụm chào buổi sáng",
        meaning: "Chào buổi sáng!",
        exampleRu: "Доброе утро, коллега!",
        exampleVi: "Chào buổi sáng, đồng nghiệp!"
      },
      {
        word: "добрый день",
        ipa: "[ˈdobrɨj dʲenʲ]",
        phonetic: "dô-bry den'",
        gram: "Cụm chào ban ngày",
        meaning: "Chào buổi trưa / chiều!",
        exampleRu: "Добрый день! Рад вас видеть.",
        exampleVi: "Chào bạn! Rất vui được gặp bạn."
      },
      {
        word: "добрый вечер",
        ipa: "[ˈdobrɨj ˈvʲetɕɪr]",
        phonetic: "dô-bry vye-cher",
        gram: "Cụm chào buổi tối",
        meaning: "Chào buổi tối!",
        exampleRu: "Добрый вечер, друзья!",
        exampleVi: "Chào buổi tối, các bạn!"
      },
      {
        word: "до свидания",
        ipa: "[də svʲɪˈdanʲɪjə]",
        phonetic: "da svi-da-ni-ya",
        gram: "Cụm tạm biệt lịch sự",
        meaning: "Tạm biệt (hẹn gặp lại)",
        exampleRu: "Спасибо за урок, до свидания!",
        exampleVi: "Cảm ơn vì bài học, tạm biệt thầy/cô!"
      },
      {
        word: "пока",
        ipa: "[pɐˈka]",
        phonetic: "pa-ka",
        gram: "Tạm biệt thân mật",
        meaning: "Tạm biệt nhé (bạn bè)",
        exampleRu: "Пока! Увидимся завтра.",
        exampleVi: "Tạm biệt! Hẹn gặp lại ngày mai."
      },
      {
        word: "спасибо",
        ipa: "[spɐˈsʲibə]",
        phonetic: "spa-si-ba",
        gram: "Lời cảm ơn",
        meaning: "Cảm ơn",
        exampleRu: "Большое спасибо за помощь!",
        exampleVi: "Cảm ơn bạn rất nhiều vì đã giúp đỡ!"
      },
      {
        word: "пожалуйста",
        ipa: "[pɐˈʐalʊstə]",
        phonetic: "pa-zha-luy-sta",
        gram: "Thán từ lịch sự",
        meaning: "Xin mời / Không có gì (đáp lại cảm ơn)",
        exampleRu: "— Спасибо! — Пожалуйста!",
        exampleVi: "— Cảm ơn bạn! — Không có gì!"
      },
      {
        word: "извините",
        ipa: "[ɪzvʲɪˈnʲitʲe]",
        phonetic: "iz-vi-ni-tye",
        gram: "Lời xin lỗi lịch sự",
        meaning: "Xin lỗi / Xin phép làm phiền",
        exampleRu: "Извините, вы говорите по-английски?",
        exampleVi: "Xin lỗi, bạn có nói tiếng Anh không?"
      },
      {
        word: "меня зовут",
        ipa: "[mʲɪˈnʲa zɐˈvut]",
        phonetic: "me-nya za-vut",
        gram: "Cụm giới thiệu tên",
        meaning: "Tôi tên là...",
        exampleRu: "Меня зовут Виктор, я из Ханоя.",
        exampleVi: "Tôi tên là Viktor, tôi đến từ Hà Nội."
      },
      {
        word: "очень приятно",
        ipa: "[ˈotɕɪnʲ prʲɪˈjatnə]",
        phonetic: "ô-chin pri-yat-na",
        gram: "Cụm biểu cảm",
        meaning: "Rất hân hạnh được làm quen",
        exampleRu: "Очень приятно с вами познакомиться!",
        exampleVi: "Rất vui được làm quen với bạn!"
      }
    ]
  },
  {
    id: "topic-2",
    num: "02",
    icon: "👨‍👩‍👧",
    titleVi: "Gia đình & Bản thân",
    titleRu: "Семья и о себе",
    desc: "Từ vựng về bố mẹ, anh chị em, họ hàng, tình trạng hôn nhân và cuộc sống gia đình.",
    count: 14,
    cards: [
      {
        word: "семья",
        ipa: "[sʲɪmʲˈja]",
        phonetic: "sem-ya",
        gram: "Danh từ (giống cái)",
        meaning: "Gia đình",
        exampleRu: "У меня большая и дружная семья.",
        exampleVi: "Tôi có một gia đình lớn và hòa thuận."
      },
      {
        word: "родители",
        ipa: "[rɐˈdʲitʲɪlʲɪ]",
        phonetic: "ra-di-tye-li",
        gram: "Danh từ số nhiều",
        meaning: "Cha mẹ, phụ huynh",
        exampleRu: "Мои родители живут во Вьетнаме.",
        exampleVi: "Bố mẹ tôi sống ở Việt Nam."
      },
      {
        word: "отец",
        ipa: "[ɐˈtʲet͡s]",
        phonetic: "a-tyets",
        gram: "Danh từ (giống đực)",
        meaning: "Người cha (bố)",
        exampleRu: "Мой отец работает инженером.",
        exampleVi: "Bố tôi làm kỹ sư."
      },
      {
        word: "мать",
        ipa: "[matʲ]",
        phonetic: "mat'",
        gram: "Danh từ (giống cái)",
        meaning: "Người mẹ (mẹ)",
        exampleRu: "Моя мама — учительница русского языка.",
        exampleVi: "Mẹ tôi là giáo viên tiếng Nga."
      },
      {
        word: "брат",
        ipa: "[brat]",
        phonetic: "brat",
        gram: "Danh từ (giống đực)",
        meaning: "Anh trai / em trai",
        exampleRu: "Старший брат учится в университете.",
        exampleVi: "Anh trai lớn đang học đại học."
      },
      {
        word: "сестра",
        ipa: "[sʲɪˈstra]",
        phonetic: "syest-ra",
        gram: "Danh từ (giống cái)",
        meaning: "Chị gái / em gái",
        exampleRu: "Моя младшая сестра любит рисовать.",
        exampleVi: "Em gái út của tôi thích vẽ tranh."
      },
      {
        word: "дедушка",
        ipa: "[ˈdʲedʊʂkə]",
        phonetic: "dye-dush-ka",
        gram: "Danh từ (nghĩa giống đực)",
        meaning: "Ông nội / ông ngoại",
        exampleRu: "Дедушка читает газету каждое утро.",
        exampleVi: "Ông đọc báo mỗi buổi sáng."
      },
      {
        word: "бабушка",
        ipa: "[ˈbabʊʂkə]",
        phonetic: "ba-bush-ka",
        gram: "Danh từ (giống cái)",
        meaning: "Bà nội / bà ngoại",
        exampleRu: "Бабушка готовит очень вкусный борщ.",
        exampleVi: "Bà nấu súp củ dền borscht rất ngon."
      }
    ]
  },
  {
    id: "topic-3",
    num: "03",
    icon: "🍲",
    titleVi: "Ẩm thực & Nhà hàng",
    titleRu: "Еда и рестораны",
    desc: "Món ăn truyền thống Nga, thực phẩm hàng ngày, gọi món, tính tiền và khẩu vị.",
    count: 15,
    cards: [
      {
        word: "борщ",
        ipa: "[borɕː]",
        phonetic: "borshch",
        gram: "Danh từ (giống đực)",
        meaning: "Súp củ dền truyền thống Nga",
        exampleRu: "Русский борщ обычно едят со сметаной.",
        exampleVi: "Súp borscht Nga thường ăn kèm kem chua smetana."
      },
      {
        word: "пельмени",
        ipa: "[pʲɪlʲˈmʲenʲɪ]",
        phonetic: "pyel'-mye-ni",
        gram: "Danh từ số nhiều",
        meaning: "Bánh sủi cảo nhân thịt kiểu Nga",
        exampleRu: "Сибирские пельмени очень вкусные.",
        exampleVi: "Bánh pelmeni vùng Siberia rất ngon."
      },
      {
        word: "хлеб",
        ipa: "[xlʲep]",
        phonetic: "khlep",
        gram: "Danh từ (giống đực)",
        meaning: "Bánh mì (lương thực chính)",
        exampleRu: "Хлеб — всему голова.",
        exampleVi: "Bánh mì là đầu mối của sự sống (tục ngữ Nga)."
      },
      {
        word: "завтрак",
        ipa: "[ˈzaftrək]",
        phonetic: "zaf-trak",
        gram: "Danh từ (giống đực)",
        meaning: "Bữa ăn sáng",
        exampleRu: "Что у нас сегодня на завтрак?",
        exampleVi: "Hôm nay chúng ta ăn gì cho bữa sáng?"
      },
      {
        word: "обед",
        ipa: "[ɐˈbʲet]",
        phonetic: "a-byet",
        gram: "Danh từ (giống đực)",
        meaning: "Bữa trưa",
        exampleRu: "В час дня мы идём на обед.",
        exampleVi: "Lúc 1 giờ chiều chúng tôi đi ăn trưa."
      },
      {
        word: "ужин",
        ipa: "[ˈuʐɨn]",
        phonetic: "u-zhin",
        gram: "Danh từ (giống đực)",
        meaning: "Bữa tối",
        exampleRu: "Семья собралась за ужином.",
        exampleVi: "Cả gia đình quây quần bên bữa tối."
      },
      {
        word: "счёт",
        ipa: "[ɕːot]",
        phonetic: "shchyot",
        gram: "Danh từ (giống đực)",
        meaning: "Hóa đơn thanh toán",
        exampleRu: "Официант, принесите счёт, пожалуйста!",
        exampleVi: "Phục vụ ơi, vui lòng mang hóa đơn giúp tôi!"
      }
    ]
  },
  {
    id: "topic-4",
    num: "04",
    icon: "🛒",
    titleVi: "Mua sắm & Đi siêu thị",
    titleRu: "Покупки и магазины",
    desc: "Hỏi giá cả, mua hàng tạp hóa, quần áo, mặc cả và phương thức thanh toán thẻ.",
    count: 12,
    cards: [
      {
        word: "магазин",
        ipa: "[məɡɐˈzʲin]",
        phonetic: "ma-ga-zin",
        gram: "Danh từ (giống đực)",
        meaning: "Cửa hàng, tiệm tạp hóa",
        exampleRu: "Я иду в магазин за продуктами.",
        exampleVi: "Tôi đi đến cửa hàng mua thực phẩm."
      },
      {
        word: "сколько стоит",
        ipa: "[ˈskolʲkə ˈstoɪt]",
        phonetic: "skol'-ka sto-it",
        gram: "Cụm hỏi giá tiền",
        meaning: "Cái này giá bao nhiêu?",
        exampleRu: "Скажите, сколько стоит эта книга?",
        exampleVi: "Xin cho hỏi quyển sách này giá bao nhiêu?"
      },
      {
        word: "скидка",
        ipa: "[ˈskʲitkə]",
        phonetic: "skit-ka",
        gram: "Danh từ (giống cái)",
        meaning: "Mức giảm giá, chiết khấu",
        exampleRu: "Студентам предоставляется скидка 15%.",
        exampleVi: "Sinh viên được giảm giá 15%."
      },
      {
        word: "наличные",
        ipa: "[nɐˈlʲitɕnɨjə]",
        phonetic: "na-lich-ny-ye",
        gram: "Danh từ số nhiều",
        meaning: "Tiền mặt",
        exampleRu: "Вы платите наличными или картой?",
        exampleVi: "Bạn thanh toán bằng tiền mặt hay thẻ?"
      }
    ]
  },
  {
    id: "topic-5",
    num: "05",
    icon: "🏠",
    titleVi: "Nhà cửa & Phòng ốc",
    titleRu: "Дом и квартира",
    desc: "Căn hộ, ký túc xá, đồ nội thất, thuê nhà và tiện nghi sinh hoạt.",
    count: 12,
    cards: [
      {
        word: "квартира",
        ipa: "[kvɐrˈtʲirə]",
        phonetic: "kvar-ti-ra",
        gram: "Danh từ (giống cái)",
        meaning: "Căn hộ chung cư",
        exampleRu: "Мы арендуем двухкомнатную квартиру.",
        exampleVi: "Chúng tôi thuê một căn hộ 2 phòng ngủ."
      },
      {
        word: "общежитие",
        ipa: "[əpɕːɪˈʐɨtʲɪjə]",
        phonetic: "ap-shche-zhi-ti-ye",
        gram: "Danh từ (giống trung)",
        meaning: "Ký túc xá sinh viên",
        exampleRu: "Иностранные студенты живут в общежитии.",
        exampleVi: "Các sinh viên quốc tế sống trong ký túc xá."
      },
      {
        word: "комната",
        ipa: "[ˈkomnətə]",
        phonetic: "kom-na-ta",
        gram: "Danh từ (giống cái)",
        meaning: "Căn phòng",
        exampleRu: "Моя комната очень светлая и уютная.",
        exampleVi: "Phòng của tôi rất sáng và ấm cúng."
      }
    ]
  },
  {
    id: "topic-6",
    num: "06",
    icon: "🚌",
    titleVi: "Đi lại & Phương tiện",
    titleRu: "Транспорт и передвижение",
    desc: "Tàu điện ngầm Metro Moskva, xe buýt, taxi, mua vé và trạm dừng chân.",
    count: 12,
    cards: [
      {
        word: "метро",
        ipa: "[mʲɪˈtro]",
        phonetic: "mye-tro",
        gram: "Danh từ bất biến",
        meaning: "Tàu điện ngầm (Metro)",
        exampleRu: "Московское метро — одно из самых красивых в мире.",
        exampleVi: "Metro Moskva là một trong những hệ thống đẹp nhất thế giới."
      },
      {
        word: "станция",
        ipa: "[ˈstantsɨjə]",
        phonetic: "stan-tsi-ya",
        gram: "Danh từ (giống cái)",
        meaning: "Nhà ga, trạm dừng",
        exampleRu: "Следующая станция — Театральная.",
        exampleVi: "Ga tiếp theo là ga Teatralnaya."
      },
      {
        word: "билет",
        ipa: "[bʲɪˈlʲet]",
        phonetic: "bi-lyet",
        gram: "Danh từ (giống đực)",
        meaning: "Vé tàu, vé xe",
        exampleRu: "Один билет на автобус, пожалуйста.",
        exampleVi: "Cho tôi một vé xe buýt, cảm ơn!"
      }
    ]
  },
  {
    id: "topic-7",
    num: "07",
    icon: "🏨",
    titleVi: "Khách sạn & Du lịch",
    titleRu: "Гостиница и туризм",
    desc: "Đặt phòng khách sạn, làm thủ tục check-in, hộ chiếu và dịch vụ phòng.",
    count: 10,
    cards: [
      {
        word: "гостиница",
        ipa: "[ɡɐˈsʲtʲinʲɪtsə]",
        phonetic: "gas-ti-ni-tsa",
        gram: "Danh từ (giống cái)",
        meaning: "Khách sạn (hotel)",
        exampleRu: "Мы забронировали номер в гостинице.",
        exampleVi: "Chúng tôi đã đặt một phòng ở khách sạn."
      },
      {
        word: "паспорт",
        ipa: "[ˈpaspərt]",
        phonetic: "pas-part",
        gram: "Danh từ (giống đực)",
        meaning: "Hộ chiếu, giấy tờ tùy thân",
        exampleRu: "Пожалуйста, предъявите ваш паспорт.",
        exampleVi: "Vui lòng xuất trình hộ chiếu của bạn."
      }
    ]
  },
  {
    id: "topic-8",
    num: "08",
    icon: "🗺️",
    titleVi: "Thành phố & Chỉ đường",
    titleRu: "Город и ориентирование",
    desc: "Hỏi đường, rẽ trái, rẽ phải, đi thẳng, khoảng cách và địa danh nổi tiếng.",
    count: 10,
    cards: [
      {
        word: "прямо",
        ipa: "[ˈprʲamə]",
        phonetic: "prya-ma",
        gram: "Phó từ chỉ hướng",
        meaning: "Đi thẳng",
        exampleRu: "Идите прямо двести метров.",
        exampleVi: "Hãy đi thẳng hai trăm mét."
      },
      {
        word: "налево",
        ipa: "[nɐˈlʲevə]",
        phonetic: "na-lye-va",
        gram: "Phó từ chỉ hướng",
        meaning: "Rẽ sang trái",
        exampleRu: "Поверните налево на перекрёстке.",
        exampleVi: "Hãy rẽ trái tại ngã tư."
      },
      {
        word: "направо",
        ipa: "[nɐˈpravə]",
        phonetic: "na-pra-va",
        gram: "Phó từ chỉ hướng",
        meaning: "Rẽ sang phải",
        exampleRu: "Аптека находится направо.",
        exampleVi: "Hiệu thuốc nằm ở phía bên phải."
      }
    ]
  },
  {
    id: "topic-9",
    num: "09",
    icon: "☀️",
    titleVi: "Thời tiết & 4 Mùa",
    titleRu: "Погода и времена года",
    desc: "Băng tuyết mùa đông Nga, cái nóng mùa hè, dự báo nhiệt độ và thời tiết.",
    count: 11,
    cards: [
      {
        word: "зима",
        ipa: "[zʲɪˈma]",
        phonetic: "zi-ma",
        gram: "Danh từ (giống cái)",
        meaning: "Mùa đông",
        exampleRu: "В России зима холодная и снежная.",
        exampleVi: "Ở Nga mùa đông lạnh và nhiều tuyết."
      },
      {
        word: "снег",
        ipa: "[sʲnʲek]",
        phonetic: "snyek",
        gram: "Danh từ (giống đực)",
        meaning: "Tuyết rơi",
        exampleRu: "На улице идёт красивый белый снег.",
        exampleVi: "Ngoài trời đang rơi những bông tuyết trắng tuyệt đẹp."
      },
      {
        word: "холодно",
        ipa: "[ˈxolədnə]",
        phonetic: "kho-lad-na",
        gram: "Vị từ chỉ trạng thái",
        meaning: "Trời lạnh",
        exampleRu: "Сегодня очень холодно, минус двадцать градусов.",
        exampleVi: "Hôm nay trời rất lạnh, âm hai mươi độ C."
      }
    ]
  },
  {
    id: "topic-10",
    num: "10",
    icon: "⏰",
    titleVi: "Thời gian & Lịch",
    titleRu: "Время, дни и календарь",
    desc: "Các ngày trong tuần, 12 tháng, hỏi giờ và hẹn gặp mặt đúng giờ.",
    count: 12,
    cards: [
      {
        word: "который час",
        ipa: "[kɐˈtorɨj tɕas]",
        phonetic: "ka-to-ry chas",
        gram: "Cụm hỏi giờ",
        meaning: "Mấy giờ rồi?",
        exampleRu: "Скажите, пожалуйста, который час?",
        exampleVi: "Làm ơn cho hỏi, bây giờ là mấy giờ?"
      },
      {
        word: "понедельник",
        ipa: "[pənʲɪˈdʲelʲnʲɪk]",
        phonetic: "pa-nye-del'-nik",
        gram: "Danh từ (giống đực)",
        meaning: "Thứ Hai",
        exampleRu: "В понедельник начинается новая рабочая неделя.",
        exampleVi: "Thứ Hai bắt đầu một tuần làm việc mới."
      }
    ]
  },
  {
    id: "topic-11",
    num: "11",
    icon: "💰",
    titleVi: "Số đếm & Tiền tệ",
    titleRu: "Числа и деньги",
    desc: "Số đếm cơ bản, đồng Rúp Nga (Рубль), đổi ngoại tệ và tính tiền.",
    count: 10,
    cards: [
      {
        word: "рубль",
        ipa: "[rublʲ]",
        phonetic: "rubl'",
        gram: "Danh từ (giống đực)",
        meaning: "Đồng Rúp Nga (tiền tệ)",
        exampleRu: "Сто рублей стоит чашка чая.",
        exampleVi: "Một tách trà có giá một trăm rúp."
      },
      {
        word: "деньги",
        ipa: "[ˈdʲenʲɡʲɪ]",
        phonetic: "dyen'-gi",
        gram: "Danh từ số nhiều",
        meaning: "Tiền bạc",
        exampleRu: "Где здесь можно обменять деньги?",
        exampleVi: "Ở đâu gần đây có thể đổi tiền được?"
      }
    ]
  },
  {
    id: "topic-12",
    num: "12",
    icon: "💊",
    titleVi: "Sức khỏe & Hiệu thuốc",
    titleRu: "Здоровье и аптека",
    desc: "Triệu chứng ốm đau, đi khám bác sĩ, mua thuốc theo đơn tại hiệu thuốc Apteka.",
    count: 11,
    cards: [
      {
        word: "аптека",
        ipa: "[ɐpˈtʲekə]",
        phonetic: "ap-tye-ka",
        gram: "Danh từ (giống cái)",
        meaning: "Hiệu thuốc Tây",
        exampleRu: "Где находится круглосуточная аптека?",
        exampleVi: "Hiệu thuốc mở cửa 24/7 ở đâu?"
      },
      {
        word: "врач",
        ipa: "[vratɕ]",
        phonetic: "vrach",
        gram: "Danh từ (giống đực)",
        meaning: "Bác sĩ",
        exampleRu: "Врач осмотрел пациента и выписал рецепт.",
        exampleVi: "Bác sĩ đã khám bệnh nhân và kê đơn thuốc."
      }
    ]
  },
  {
    id: "topic-13",
    num: "13",
    icon: "👤",
    titleVi: "Cơ thể con người",
    titleRu: "Тело человека",
    desc: "Đầu, mắt, mũi, tay chân và các bộ phận cơ thể cần thiết khi giao tiếp y tế.",
    count: 10,
    cards: [
      {
        word: "голова",
        ipa: "[ɡəlɐˈva]",
        phonetic: "ga-la-va",
        gram: "Danh từ (giống cái)",
        meaning: "Đầu (bộ phận)",
        exampleRu: "У меня сегодня сильно болит голова.",
        exampleVi: "Hôm nay tôi bị đau đầu dữ dội."
      },
      {
        word: "рука",
        ipa: "[rʊˈka]",
        phonetic: "ru-ka",
        gram: "Danh từ (giống cái)",
        meaning: "Bàn tay / cánh tay",
        exampleRu: "Помой руки перед едой.",
        exampleVi: "Hãy rửa tay trước khi ăn."
      }
    ]
  },
  {
    id: "topic-14",
    num: "14",
    icon: "😊",
    titleVi: "Cảm xúc & Tính cách",
    titleRu: "Эмоции и характер",
    desc: "Vui vẻ, buồn bã, mệt mỏi, thân thiện, tốt bụng và phẩm chất con người.",
    count: 10,
    cards: [
      {
        word: "радость",
        ipa: "[ˈradəsʲtʲ]",
        phonetic: "ra-dast'",
        gram: "Danh từ (giống cái)",
        meaning: "Niềm vui, sự vui mừng",
        exampleRu: "Эта новость принесла нам большую радость.",
        exampleVi: "Tin tức này mang lại cho chúng tôi niềm vui lớn."
      },
      {
        word: "добрый",
        ipa: "[ˈdobrɨj]",
        phonetic: "dô-bry",
        gram: "Tính từ",
        meaning: "Tốt bụng, hiền hậu",
        exampleRu: "Он очень добрый и отзывчивый человек.",
        exampleVi: "Anh ấy là một người rất tốt bụng và nhiệt tình."
      }
    ]
  },
  {
    id: "topic-15",
    num: "15",
    icon: "👗",
    titleVi: "Quần áo & Thời trang",
    titleRu: "Одежда и обувь",
    desc: "Áo khoác ấm mùa đông, khăn len, găng tay, giày dép và kích cỡ.",
    count: 10,
    cards: [
      {
        word: "пальто",
        ipa: "[pɐlʲˈto]",
        phonetic: "pal'-to",
        gram: "Danh từ bất biến",
        meaning: "Áo khoác măng-tô",
        exampleRu: "Зимой в Москве нужно носить тёплое пальто.",
        exampleVi: "Mùa đông ở Moskva cần mặc áo khoác dày ấm."
      },
      {
        word: "шапка",
        ipa: "[ˈʂapkə]",
        phonetic: "shap-ka",
        gram: "Danh từ (giống cái)",
        meaning: "Mũ len, mũ mùa đông",
        exampleRu: "Не забудь надеть тёплую шапку!",
        exampleVi: "Đừng quên đội mũ len ấm nhé!"
      }
    ]
  },
  {
    id: "topic-16",
    num: "16",
    icon: "🎨",
    titleVi: "Màu sắc & Kích thước",
    titleRu: "Цвета и размеры",
    desc: "Trắng, đen, đỏ, xanh, vàng, lớn, nhỏ, dài ngắn và mô tả đặc điểm vật thể.",
    count: 10,
    cards: [
      {
        word: "красный",
        ipa: "[ˈkrasnɨj]",
        phonetic: "kras-ny",
        gram: "Tính từ",
        meaning: "Màu đỏ",
        exampleRu: "Красная площадь — сердце Москвы.",
        exampleVi: "Quảng trường Đỏ là trái tim của Moskva."
      },
      {
        word: "белый",
        ipa: "[ˈbʲelɨj]",
        phonetic: "bye-ly",
        gram: "Tính từ",
        meaning: "Màu trắng",
        exampleRu: "Белые ночи в Санкт-Петербурге прекрасны.",
        exampleVi: "Đêm trắng ở Saint Petersburg thật tuyệt đẹp."
      }
    ]
  },
  {
    id: "topic-17",
    num: "17",
    icon: "💼",
    titleVi: "Nghề nghiệp & Nơi làm việc",
    titleRu: "Профессии и работа",
    desc: "Bác sĩ, giáo viên, kỹ sư, lập trình viên, phỏng vấn và môi trường làm việc.",
    count: 11,
    cards: [
      {
        word: "учитель",
        ipa: "[ʊˈtɕitʲɪlʲ]",
        phonetic: "u-chi-tyel'",
        gram: "Danh từ (giống đực)",
        meaning: "Giáo viên",
        exampleRu: "Учитель объясняет грамматику урока.",
        exampleVi: "Thầy giáo đang giải thích ngữ pháp bài học."
      },
      {
        word: "работа",
        ipa: "[rɐˈbotə]",
        phonetic: "ra-bo-ta",
        gram: "Danh từ (giống cái)",
        meaning: "Công việc, chỗ làm",
        exampleRu: "Я люблю свою интересную работу.",
        exampleVi: "Tôi yêu công việc thú vị của mình."
      }
    ]
  },
  {
    id: "topic-18",
    num: "18",
    icon: "🎓",
    titleVi: "Trường học & Học tập",
    titleRu: "Школа и учеба",
    desc: "Đại học, bài giảng, thi cử, sách giáo khoa, ghi chú và bài tập về nhà.",
    count: 11,
    cards: [
      {
        word: "университет",
        ipa: "[ʊnʲɪvʲɪrsʲɪˈtʲet]",
        phonetic: "u-ni-vyer-si-tyet",
        gram: "Danh từ (giống đực)",
        meaning: "Trường đại học",
        exampleRu: "МГУ — старейший университет России.",
        exampleVi: "MGU là trường đại học lâu đời nhất nước Nga."
      },
      {
        word: "студент",
        ipa: "[stʊˈdʲent]",
        phonetic: "stu-dyent",
        gram: "Danh từ (giống đực)",
        meaning: "Sinh viên nam",
        exampleRu: "Студенты готовятся к сессии.",
        exampleVi: "Các sinh viên đang chuẩn bị cho kỳ thi học kỳ."
      }
    ]
  },
  {
    id: "topic-19",
    num: "19",
    icon: "🇷🇺",
    titleVi: "Du học & Đời sống tại Nga",
    titleRu: "Жизнь в России",
    desc: "Visa du học, đăng ký tạm trú, thời tiết mùa đông và hòa nhập văn hóa bản xứ.",
    count: 10,
    cards: [
      {
        word: "виза",
        ipa: "[ˈvʲizə]",
        phonetic: "vi-za",
        gram: "Danh từ (giống cái)",
        meaning: "Thị thực, Visa",
        exampleRu: "Студенческая виза продлевается каждый год.",
        exampleVi: "Visa sinh viên được gia hạn hàng năm."
      },
      {
        word: "регистрация",
        ipa: "[rʲɪɡʲɪˈstratsɨjə]",
        phonetic: "rye-gis-tra-tsi-ya",
        gram: "Danh từ (giống cái)",
        meaning: "Đăng ký tạm trú",
        exampleRu: "После приезда нужно оформить регистрацию.",
        exampleVi: "Sau khi đến nơi bạn cần làm thủ tục đăng ký tạm trú."
      }
    ]
  },
  {
    id: "topic-20",
    num: "20",
    icon: "⚽",
    titleVi: "Thể thao & Giải trí",
    titleRu: "Спорт и отдых",
    desc: "Bóng đá, trượt băng, bơi lội, rạp chiếu phim và sở thích cuối tuần.",
    count: 10,
    cards: [
      {
        word: "спорт",
        ipa: "[sport]",
        phonetic: "sport",
        gram: "Danh từ (giống đực)",
        meaning: "Thể thao",
        exampleRu: "Заниматься спортом полезно для здоровья.",
        exampleVi: "Tập luyện thể thao rất có lợi cho sức khỏe."
      },
      {
        word: "коньки",
        ipa: "[kɐnʲˈkʲi]",
        phonetic: "kan'-ki",
        gram: "Danh từ số nhiều",
        meaning: "Giày trượt băng",
        exampleRu: "Зимой мы часто катаемся на коньках.",
        exampleVi: "Vào mùa đông chúng tôi thường đi trượt băng."
      }
    ]
  },
  {
    id: "topic-21",
    num: "21",
    icon: "🐻",
    titleVi: "Động vật & Thiên nhiên Nga",
    titleRu: "Природа и животные",
    desc: "Rừng Taiga, gấu nâu, chó sói, hồ Baikal và cảnh sắc thiên nhiên hùng vĩ.",
    count: 10,
    cards: [
      {
        word: "тайга",
        ipa: "[tɐjˈɡa]",
        phonetic: "tay-ga",
        gram: "Danh từ (giống cái)",
        meaning: "Rừng Taiga (rừng lá kim xứ lạnh)",
        exampleRu: "Сибирская тайга занимает огромную территорию.",
        exampleVi: "Rừng Taiga ở Siberia bao phủ một vùng lãnh thổ khổng lồ."
      },
      {
        word: "медведь",
        ipa: "[mʲɪdˈvʲetʲ]",
        phonetic: "myed-vyet'",
        gram: "Danh từ (giống đực)",
        meaning: "Con gấu (biểu tượng nước Nga)",
        exampleRu: "Бурый медведь — символ русской природы.",
        exampleVi: "Gấu nâu là biểu tượng của thiên nhiên nước Nga."
      }
    ]
  },
  {
    id: "topic-22",
    num: "22",
    icon: "💻",
    titleVi: "Công nghệ & Internet",
    titleRu: "Технологии и интернет",
    desc: "Điện thoại, ứng dụng Telegram, Yandex, máy tính và kết nối mạng.",
    count: 10,
    cards: [
      {
        word: "компьютер",
        ipa: "[kɐmˈpʲjutər]",
        phonetic: "kam-pyu-ter",
        gram: "Danh từ (giống đực)",
        meaning: "Máy vi tính",
        exampleRu: "Я работаю за компьютером каждый день.",
        exampleVi: "Tôi làm việc với máy tính mỗi ngày."
      },
      {
        word: "приложение",
        ipa: "[prʲɪlɐˈʐenʲɪjə]",
        phonetic: "pri-la-zhe-ni-ye",
        gram: "Danh từ (giống trung)",
        meaning: "Ứng dụng (App)",
        exampleRu: "Я скачал удобное приложение для изучения слов.",
        exampleVi: "Tôi đã tải ứng dụng tiện ích để học từ vựng."
      }
    ]
  },
  {
    id: "topic-23",
    num: "23",
    icon: "🏦",
    titleVi: "Bưu điện & Ngân hàng",
    titleRu: "Почта и банк",
    desc: "Gửi bưu phẩm, chuyển tiền, mở tài khoản Sberbank và giao dịch thủ tục.",
    count: 10,
    cards: [
      {
        word: "банк",
        ipa: "[bank]",
        phonetic: "bank",
        gram: "Danh từ (giống đực)",
        meaning: "Ngân hàng",
        exampleRu: "Мне нужно открыть счёт в Сбербанке.",
        exampleVi: "Tôi cần mở tài khoản tại ngân hàng Sberbank."
      },
      {
        word: "посылка",
        ipa: "[pɐˈsɨlkə]",
        phonetic: "pa-syl-ka",
        gram: "Danh từ (giống cái)",
        meaning: "Kiện hàng bưu phẩm",
        exampleRu: "Я получил посылку из Вьетнама.",
        exampleVi: "Tôi đã nhận được kiện bưu phẩm gửi từ Việt Nam."
      }
    ]
  },
  {
    id: "topic-24",
    num: "24",
    icon: "🚨",
    titleVi: "Tình huống khẩn cấp",
    titleRu: "Экстренные ситуации",
    desc: "Cấp cứu 112, cứu hỏa, cảnh sát, lạc đường và mẫu câu kêu cứu bảo vệ bản thân.",
    count: 10,
    cards: [
      {
        word: "помогите",
        ipa: "[pəmɐˈɡʲitʲe]",
        phonetic: "pa-ma-gi-tye",
        gram: "Mệnh lệnh thức",
        meaning: "Cứu tôi với! Giúp tôi với!",
        exampleRu: "Помогите, пожалуйста, я потерял сумку!",
        exampleVi: "Làm ơn giúp tôi với, tôi bị mất túi xách!"
      },
      {
        word: "полиция",
        ipa: "[pɐˈlʲitsɨjə]",
        phonetic: "pa-li-tsi-ya",
        gram: "Danh từ (giống cái)",
        meaning: "Cảnh sát (số gọi 102/112)",
        exampleRu: "Срочно вызовите полицию!",
        exampleVi: "Hãy gọi cảnh sát khẩn cấp!"
      }
    ]
  },
  {
    id: "topic-25",
    num: "25",
    icon: "🤝",
    titleVi: "Tiếng Nga thương mại",
    titleRu: "Деловой русский язык",
    desc: "Đàm phán hợp đồng, trao đổi thư từ email đối tác, báo giá và ký kết.",
    count: 12,
    cards: [
      {
        word: "договор",
        ipa: "[dəɡɐˈvor]",
        phonetic: "da-ga-vor",
        gram: "Danh từ (giống đực)",
        meaning: "Hợp đồng kinh tế",
        exampleRu: "Стороны подписали взаимовыгодный договор.",
        exampleVi: "Hai bên đã ký kết hợp đồng cùng có lợi."
      },
      {
        word: "сотрудничество",
        ipa: "[səˈtrudnʲɪtɕəstvə]",
        phonetic: "sa-trud-ni-chyest-va",
        gram: "Danh từ (giống trung)",
        meaning: "Sự hợp tác lâu dài",
        exampleRu: "Благодарим вас за плодотворное сотрудничество!",
        exampleVi: "Xin chân thành cảm ơn sự hợp tác hiệu quả của quý vị!"
      }
    ]
  }
];

