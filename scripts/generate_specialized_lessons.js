const fs = require('fs');
const path = require('path');

const baseDir = path.join(__dirname, '../public/data/specialized');

const domains = [
  {
    key: "tour_guide",
    name: "HDV Du Lịch",
    icon: "🎙️",
    badge: "Chuyên Ngành Du Lịch",
    lessons: [
      {
        id: 1,
        title: "Встреча туристов в аэропорту и трансфер в отель",
        readingTime: "12 phút",
        source: "Cẩm nang Hướng dẫn viên du lịch đón khách quốc tế",
        topic: "Nghiệp vụ HDV: Đón đoàn tại sân bay, làm thủ tục xe đưa đón và lưu ý an toàn",
        vocabCategories: [
          {
            title: "Thuật ngữ đón khách & Phương tiện (Трансфер и встреча)",
            rows: [
              { word: "Встреча́ющий гид", gram: "danh từ nam", meaning: "Hướng dẫn viên đón đoàn", example: "Гид держит табличку с названием туроператора." },
              { word: "Тра́нсфер", gram: "danh từ nam", meaning: "Xe đưa đón sân bay về khách sạn", example: "Наш комфортабельный автобус готов к трансферу." },
              { word: "Ва́учер", gram: "danh từ nam", meaning: "Phiếu xác nhận dịch vụ du lịch", example: "Пожалуйста, предъявите ваш туристический ваучер." },
              { word: "Бага́жное отделе́ние", gram: "danh từ trung", meaning: "Khoang chứa hành lý xe bus", example: "Положите чемоданы в багажное отделение автобуса." },
              { word: "Заселе́ние в оте́ль", gram: "danh từ trung", meaning: "Làm thủ tục nhận phòng (Check-in)", example: "Заселение в номера начнётся ровно в четырнадцать часов." }
            ]
          },
          {
            title: "Mẫu câu hướng dẫn & Dặn dò (Инструктаж туристов)",
            rows: [
              { word: "Добро́ пожа́ловать!", gram: "thành ngữ chào đón", meaning: "Nhiệt liệt chào mừng quý khách!", example: "Добро пожаловать в солнечный Вьетнам!" },
              { word: "Разни́ца во вре́мени", gram: "cụm danh từ", meaning: "Độ chênh lệch múi giờ", example: "Разница во времени с Москвой составляет плюс четыре часа." },
              { word: "Обме́н валю́ты", gram: "cụm danh từ", meaning: "Đổi tiền ngoại tệ", example: "Вы можете выгодно обменять доллары и рубли на вьетнамские донги на ресепшн." },
              { word: "Сим-ка́рта с интерне́том", gram: "cụm từ", meaning: "SIM điện thoại 4G/5G", example: "Купить местную сим-карту можно прямо в терминале аэропорта." }
            ]
          }
        ],
        readingText: [
          "Здравствуйте, дорогие гости! Добро пожаловать во Вьетнам! Меня зовут Нгуен, я ваш сопровождающий гид от туристической компании. Поздравляю вас с успешным прибытием в международный аэропорт Камрань!",
          "Сейчас мы пройдём на парковку, где нас ждёт большой кондиционированный автобус номер восемь. Пожалуйста, проверьте наличие всех ваших сумок и чемоданов перед тем, как сложить их в багажное отделение. Дорога до курортного города Нячанг займёт примерно сорок пять минут.",
          "Обратите внимание: местное время опережает московское на четыре часа. Официальная валюта нашей страны — вьетнамский донг. Менять валюту в аэропорту лучше в официальных обменных пунктах или в банках города. Не рекомендуем менять деньги у частных лиц на улице.",
          "Завтра в девять часов утра в холле вашего отеля состоится важная информационная встреча. Я подробно расскажу о правилах безопасности, медицинских страховках, а также о самых захватывающих экскурсиях: морской прогулке по островам, водопадах и грязелечебницах.",
          "Желаю вам незабываемого и яркого отдыха на лазурном побережье Вьетнама! Если у вас возникнут любые вопросы, я круглосуточно на связи."
        ],
        questions: [
          {
            id: "tg1_q1", lang: "ru",
            prompt: "Кем работает автор приветственного обращения?",
            hint: "Nghề nghiệp của người nói.",
            options: [
              { key: "A", text: "Сотрудником таможни аэропорта." },
              { key: "B", text: "Сопровождающим туристическим гидом." },
              { key: "C", text: "Администратором отеля." },
              { key: "D", text: "Водителем такси." }
            ],
            ans: "B", explain: "Người nói tự giới thiệu: «Меня зовут Нгуен, я ваш сопровождающий гид от туристической компании»."
          },
          {
            id: "tg1_q2", lang: "ru",
            prompt: "Сколько времени займёт поездка из аэропорта до Нячанга?",
            hint: "Thời gian xe di chuyển.",
            options: [
              { key: "A", text: "Около двадцати минут." },
              { key: "B", text: "Примерно сорок пять минут." },
              { key: "C", text: "Два часа." },
              { key: "D", text: "Полтора часа." }
            ],
            ans: "B", explain: "Bài nói: «Дорога до курортного города Нячанг займёт примерно сорок пять минут»."
          },
          {
            id: "tg1_q3", lang: "ru",
            prompt: "На сколько часов местное время во Вьетнаме опережает московское?",
            hint: "Chênh lệch múi giờ.",
            options: [
              { key: "A", text: "На 2 часа." },
              { key: "B", text: "На 3 часа." },
              { key: "C", text: "На 4 часа." },
              { key: "D", text: "Время совпадает." }
            ],
            ans: "C", explain: "HDV nêu: «местное время опережает московское на четыре часа (+4)»."
          },
          {
            id: "tg1_q4", lang: "ru",
            prompt: "Во сколько состоится информационная встреча в отеле?",
            hint: "Giờ họp đoàn sáng mai.",
            options: [
              { key: "A", text: "В 8:00 утра." },
              { key: "B", text: "В 9:00 утра." },
              { key: "C", text: "В 12:00 дня." },
              { key: "D", text: "В 18:00 вечера." }
            ],
            ans: "B", explain: "HDV dặn: «Завтра в девять часов утра в холле вашего отеля состоится информационная встреча»."
          },
          {
            id: "tg1_q5", lang: "ru",
            prompt: "Где гид НЕ рекомендует туристам менять валюту?",
            hint: "Nơi không nên đổi tiền.",
            options: [
              { key: "A", text: "В банках города." },
              { key: "B", text: "В официальных обменниках аэропорта." },
              { key: "C", text: "У частных лиц на улице." },
              { key: "D", text: "На ресепшн отеля." }
            ],
            ans: "C", explain: "HDV cảnh báo: «Не рекомендуем менять деньги у частных лиц на улице»."
          },
          {
            id: "tg1_q6", lang: "vi",
            prompt: "Từ «Трансфер» trong nghiệp vụ du lịch lữ hành có nghĩa là gì?",
            hint: "Dịch vụ di chuyển giữa các điểm.",
            options: [
              { key: "A", text: "Dịch vụ giặt ủi tại phòng khách sạn." },
              { key: "B", text: "Dịch vụ xe đưa đón khách giữa sân bay và khách sạn." },
              { key: "C", text: "Vé máy bay khứ hồi." },
              { key: "D", text: "Thủ tục xin visa thị thực tại cửa khẩu." }
            ],
            ans: "B", explain: "«Трансфер» là thuật ngữ chuyên ngành chỉ dịch vụ vận chuyển đưa đón du khách giữa sân bay và nơi lưu trú."
          },
          {
            id: "tg1_q7", lang: "vi",
            prompt: "Trước khi hành khách xếp va ly vào khoang xe bus, HDV đã yêu cầu điều gì?",
            hint: "Kiểm tra tư trang.",
            options: [
              { key: "A", text: "Đóng thêm phí cước hành lý." },
              { key: "B", text: "Mở khóa va ly để kiểm tra an ninh." },
              { key: "C", text: "Kiểm tra kỹ lưỡng đầy đủ số lượng túi xách và va ly." },
              { key: "D", text: "Ký tên vào danh sách nhận phòng trước." }
            ],
            ans: "C", explain: "«Пожалуйста, проверьте наличие всех ваших сумок и чемоданов перед тем, как сложить их в багажное отделение»."
          },
          {
            id: "tg1_q8", lang: "vi",
            prompt: "Buổi họp thông tin (Инфо-встреча) sáng hôm sau nhằm mục đích chính gì?",
            hint: "Nội dung buổi phổ biến.",
            options: [
              { key: "A", text: "Thu lại hộ chiếu của du khách." },
              { key: "B", text: "Phổ biến quy tắc an toàn, bảo hiểm và giới thiệu các tour tham quan." },
              { key: "C", text: "Đổi tiền Việt Nam đồng bắt buộc." },
              { key: "D", text: "Kiểm tra sức khỏe đầu vào của du khách." }
            ],
            ans: "B", explain: "Nội dung cuộc họp: «расскажу о правилах безопасности, медицинских страховках, а также о самых захватывающих экскурсиях»."
          },
          {
            id: "tg1_q9", lang: "vi",
            prompt: "Cụm từ «круглосуточно на связи» của HDV thể hiện cam kết gì?",
            hint: "Thời gian hỗ trợ khách.",
            options: [
              { key: "A", text: "Chỉ nghe máy trong giờ hành chính từ 8h đến 17h." },
              { key: "B", text: "Luôn sẵn sàng hỗ trợ khách 24/7 (ngày lẫn đêm)." },
              { key: "C", text: "Chỉ nhắn tin qua mạng xã hội sau 20h." },
              { key: "D", text: "Khách phải trả phí khi gọi điện tư vấn." }
            ],
            ans: "B", explain: "«круглосуточно на связи» nghĩa là giữ liên lạc và trực hỗ trợ 24 giờ mỗi ngày (24/7)."
          },
          {
            id: "tg1_q10", lang: "vi",
            prompt: "Thành phố biển được nhắc đến làm điểm đến trong bài là thành phố nào?",
            hint: "Địa danh du lịch nổi tiếng miền Trung.",
            options: [
              { key: "A", text: "Đà Nẵng." },
              { key: "B", text: "Nha Trang." },
              { key: "C", text: "Vũng Tàu." },
              { key: "D", text: "Hạ Long." }
            ],
            ans: "B", explain: "Sân bay Cam Ranh và cung đường đến thành phố nghỉ dưỡng Nha Trang."
          }
        ]
      }
    ]
  },
  {
    key: "trade_logistics",
    name: "Thương Mại XNK",
    icon: "🚢",
    badge: "Chuyên Ngành Xuất Nhập Khẩu",
    lessons: [
      {
        id: 1,
        title: "Переговоры по внешнеторговому контракту и условия Инкотермс",
        readingTime: "14 phút",
        source: "Tài liệu Nghiệp vụ Ngoại thương & Xuất nhập khẩu (ВЭД)",
        topic: "Đàm phán hợp đồng thương mại quốc tế, điều kiện giao hàng CIF/FOB và phương thức thanh toán L/C",
        vocabCategories: [
          {
            title: "Thuật ngữ Ngoại thương & Hải quan (ВЭД и таможня)",
            rows: [
              { word: "Вне́шнеторго́вый контра́кт", gram: "cụm danh từ", meaning: "Hợp đồng ngoại thương / mua bán quốc tế", example: "Стороны подписали внешнеторговый контракт на поставку кофе." },
              { word: "Усло́вия поста́вки", gram: "cụm danh từ", meaning: "Điều kiện giao hàng (IncoTerms)", example: "Поставка товара осуществляется на условиях CIF Владивосток." },
              { word: "Тамо́женная по́шлина", gram: "cụm từ", meaning: "Thuế hải quan / thuế nhập khẩu", example: "Ставка ввозной таможенной пошлины составляет пять процентов." },
              { word: "Коносаме́нт (Bill of Lading)", gram: "danh từ nam", meaning: "Vận đơn đường biển", example: "Морской коносамент подтверждает приём груза к перевозке." },
              { word: "Сертифика́т происхожде́ния (C/O)", gram: "cụm từ", meaning: "Chứng nhận xuất xứ hàng hóa", example: "Сертификат формы E подтверждает вьетнамское происхождение товара." }
            ]
          },
          {
            title: "Thanh toán & Chứng từ (Оплата и документы)",
            rows: [
              { word: "Аккредити́в (L/C)", gram: "danh từ nam", meaning: "Thư tín dụng chứng từ", example: "Оплата производится посредством безотзывного документарного аккредитива." },
              { word: "Инво́йс (Счёт-фактура)", gram: "danh từ nam", meaning: "Hóa đơn thương mại", example: "Коммерческий инвойс содержит точную стоимость и объём партии." },
              { word: "Упако́вочный лист (Packing List)", gram: "cụm từ", meaning: "Phiếu đóng gói hàng hóa", example: "В упаковочном листе указан вес брутто и нетто каждого места." }
            ]
          }
        ],
        readingText: [
          "Вьетнамская экспортная компания успешно завершила переговоры с российскими партнёрами о регулярных поставках натурального кофе робуста и сухофруктов на рынок Российской Федерации.",
          "В ходе деловых переговоров стороны согласовали ключевые условия внешнеторгового контракта. Базисом поставки был выбран международный термин Инкотермс 2020 — FOB порт Хошимин (FOB Ho Chi Minh Port). Это означает, что продавец обязуется доставить груз в порт и погрузить его на борт морского судна, после чего все риски переходят к покупателю.",
          "Общий объём первой пробной партии составляет пять сорокафутовых контейнеров. В качестве надёжной формы международных расчётов стороны договорились использовать безотзывный покрытый аккредитив (Letter of Credit), открытый первоклассным международным банком.",
          "Для успешного таможенного оформления экспортер обязуется предоставить полный пакет товаросопроводительных документов: коммерческий инвойс, упаковочный лист, фитосанитарный сертификат, коносамент и сертификат происхождения товара по форме Вьетнам-ЕАЭС для получения тарифных преференций.",
          "Соблюдение всех международных стандартов гарантирует стабильное развитие торговых отношений между Вьетнамом и Россией."
        ],
        questions: [
          {
            id: "tl1_q1", lang: "ru",
            prompt: "Какой базис поставки согласовали стороны контракта?",
            hint: "Điều kiện giao hàng IncoTerms.",
            options: [
              { key: "A", text: "EXW (самовывоз с фабрики)." },
              { key: "B", text: "FOB порт Хошимин." },
              { key: "C", text: "DDP склад покупателя в Москве." },
              { key: "D", text: "CIF Санкт-Петербург." }
            ],
            ans: "B", explain: "Văn bản nêu: «Базисом поставки был выбран термин Инкотермс 2020 — FOB порт Хошимин»."
          },
          {
            id: "tl1_q2", lang: "ru",
            prompt: "Какую продукцию поставляет вьетнамская компания?",
            hint: "Mặt hàng xuất khẩu.",
            options: [
              { key: "A", text: "Электронику и автомобили." },
              { key: "B", text: "Натуральный кофе робуста и сухофрукты." },
              { key: "C", text: "Одежду и обувь из кожи." },
              { key: "D", text: "Нефть и природный газ." }
            ],
            ans: "B", explain: "«о регулярных поставках натурального кофе робуста и сухофруктов»."
          },
          {
            id: "tl1_q3", lang: "ru",
            prompt: "Какая форма расчётов выбрана партнёрами?",
            hint: "Phương thức thanh toán quốc tế.",
            options: [
              { key: "A", text: "Оплата наличными курьеру." },
              { key: "B", text: "Безотзывный аккредитив (L/C)." },
              { key: "C", text: "Банковский чек через полгода." },
              { key: "D", text: "Бартерный обмен на зерно." }
            ],
            ans: "B", explain: "«договорились использовать безотзывный покрытый аккредитив (Letter of Credit)»."
          },
          {
            id: "tl1_q4", lang: "ru",
            prompt: "Каков объём первой пробной партии груза?",
            hint: "Số lượng container.",
            options: [
              { key: "A", text: "1 контейнер 20 футов." },
              { key: "B", text: "5 сорокафутовых (40-ft) контейнеров." },
              { key: "C", text: "10 вагонов поезда." },
              { key: "D", text: "100 ящиков." }
            ],
            ans: "B", explain: "«Общий объём первой пробной партии составляет пять сорокафутовых контейнеров»."
          },
          {
            id: "tl1_q5", lang: "ru",
            prompt: "Для чего необходим сертификат происхождения Вьетнам-ЕАЭС?",
            hint: "Mục đích giấy C/O.",
            options: [
              { key: "A", text: "Для получения тарифных преференций (снижения пошлин)." },
              { key: "B", text: "Для проверки веса контейнера." },
              { key: "C", text: "Для оплаты морского фрахта." },
              { key: "D", text: "Для рекламы товара в соцсетях." }
            ],
            ans: "A", explain: "Giấy chứng nhận xuất xứ form EAEU giúp hàng hóa hưởng ưu đãi thuế quan."
          },
          {
            id: "tl1_q6", lang: "vi",
            prompt: "Điều kiện giao hàng FOB (Free on Board) quy định trách nhiệm của người bán đến thời điểm nào?",
            hint: "Ranh giới chuyển giao rủi ro.",
            options: [
              { key: "A", text: "Giao hàng tới tận kho người mua ở Nga." },
              { key: "B", text: "Giao hàng và bốc hàng an toàn qua lan can tàu biển tại cảng đi." },
              { key: "C", text: "Chỉ cần để hàng tại xưởng của người bán." },
              { key: "D", text: "Người bán phải chịu tiền cước tàu biển quốc tế." }
            ],
            ans: "B", explain: "FOB quy định người bán hoàn tất nghĩa vụ khi hàng đã được bốc lên boong tàu tại cảng bốc hàng chỉ định."
          },
          {
            id: "tl1_q7", lang: "vi",
            prompt: "Phương thức thanh toán L/C (Аккредитив) mang lại lợi ích lớn nhất nào cho cả hai bên?",
            hint: "Tính đảm bảo thanh toán qua ngân hàng.",
            options: [
              { key: "A", text: "Không mất phí ngân hàng." },
              { key: "B", text: "Được nhận tiền mặt ngay lập tức tại cảng." },
              { key: "C", text: "Ngân hàng đứng ra bảo lãnh thanh toán khi người bán xuất trình đủ bộ chứng từ hợp lệ." },
              { key: "D", text: "Có thể đổi trả hàng vô điều kiện sau 1 năm." }
            ],
            ans: "C", explain: "Thư tín dụng L/C là cam kết thanh toán có điều kiện của ngân hàng, bảo vệ quyền lợi tối đa cho cả nhà xuất khẩu và nhập khẩu."
          },
          {
            id: "tl1_q8", lang: "vi",
            prompt: "Chứng từ nào sau đây là bằng chứng xác nhận việc hàng hóa đã được xếp lên tàu biển?",
            hint: "Vận đơn đường biển.",
            options: [
              { key: "A", text: "Коносамент (Bill of Lading)." },
              { key: "B", text: "Коммерческий инвойс." },
              { key: "C", text: "Упаковочный лист." },
              { key: "D", text: "Фитосанитарный сертификат." }
            ],
            ans: "A", explain: "Коносамент (Vận đơn đường biển) là bằng chứng xác nhận hãng tàu đã nhận hàng để vận chuyển."
          },
          {
            id: "tl1_q9", lang: "vi",
            prompt: "Chứng từ kiểm dịch thực vật (Фитосанитарный сертификат) bắt buộc đối với loại mặt hàng nào?",
            hint: "Nông sản thực vật.",
            options: [
              { key: "A", text: "Linh kiện máy tính điện tử." },
              { key: "B", text: "Nông sản như cà phê hạt, hoa quả sấy khô." },
              { key: "C", text: "Quần áo may sẵn." },
              { key: "D", text: "Dầu mỏ thô." }
            ],
            ans: "B", explain: "Mặt hàng có nguồn gốc thực vật (cà phê, hoa quả sấy) bắt buộc phải có giấy chứng nhận kiểm dịch thực vật để thông quan."
          },
          {
            id: "tl1_q10", lang: "vi",
            prompt: "Từ viết tắt ВЭД trong tiếng Nga đại diện cho khái niệm gì?",
            hint: "Hoạt động kinh tế đối ngoại.",
            options: [
              { key: "A", text: "Hội chợ triển lãm trong nước." },
              { key: "B", text: "Văn phòng quản lý thuế nội địa." },
              { key: "C", text: "Hoạt động kinh tế đối ngoại / Xuất nhập khẩu (Внешнеэкономическая деятельность)." },
              { key: "D", text: "Hệ thống bảo hiểm xã hội." }
            ],
            ans: "C", explain: "ВЭД là viết tắt của «Внешнеэкономическая деятельность» (Hoạt động kinh tế đối ngoại / Ngoại thương)."
          }
        ]
      }
    ]
  },
  {
    key: "retail_sales",
    name: "Bán Hàng & Chợ",
    icon: "🛍️",
    badge: "Chuyên Ngành Bán Hàng Thực Chiến",
    lessons: [
      {
        id: 1,
        title: "Торговля одеждой на рынке «Садовод»: Размеры, опт и розница",
        readingTime: "11 phút",
        source: "Sổ tay giao tiếp bán hàng thực tế tại các trung tâm thương mại & chợ Nga",
        topic: "Kỹ năng bán hàng: Tư vấn size số, mặc cả giá sỉ/lẻ, chất liệu vải và đóng gói vận chuyển",
        vocabCategories: [
          {
            title: "Thuật ngữ bán hàng & Giá cả (Продажи и цены)",
            rows: [
              { word: "Опто́вая цена́", gram: "cụm danh từ", meaning: "Giá bán buôn / Giá sỉ", example: "Оптовая цена действует при заказе от одного размерного ряда." },
              { word: "Ро́зница (В ро́зницу)", gram: "cụm danh từ", meaning: "Bán lẻ", example: "В розницу эта модель стоит на триста рублей дороже." },
              { word: "Разме́рный ряд (лине́йка)", gram: "cụm danh từ", meaning: "Dây size (đủ các cỡ từ S đến XL)", example: "В размерном ряду пять штук: сорок четвёртый по пятьдесят второй." },
              { word: "Ски́дка за объём", gram: "cụm từ", meaning: "Chiết khấu giảm giá theo số lượng", example: "Если возьмёте три упаковки, сделаю хорошую скидку." },
              { word: "Натура́льный хло́пок", gram: "cụm từ", meaning: "Cotton tự nhiên 100%", example: "Это дышащая ткань, сто процентов натуральный хлопок." }
            ]
          },
          {
            title: "Giao dịch & Thanh toán tại chợ (Рыночные сделки)",
            rows: [
              { word: "Приме́рочная кабина", gram: "cụm từ", meaning: "Phòng thử đồ cho khách", example: "Примерочная находится прямо за шторкой." },
              { word: "Това́рный чек / Накладна́я", gram: "cụm từ", meaning: "Hóa đơn bán hàng / Biên lai", example: "Я выпишу вам товарный чек с печатью." },
              { word: "Отпра́вка ка́рго", gram: "cụm từ", meaning: "Gửi xe chở hàng liên tỉnh / cargo", example: "Мы можем упаковать коробки и отправить через карго на ваш город." }
            ]
          }
        ],
        readingText: [
          "Рынок «Садовод» и торговый комплекс «Москва» в Люблино — крупнейшие центры оптовой и розничной торговли одеждой, обувью и текстилем в столице России. Здесь работают тысячи предпринимателей, в том числе представители вьетнамской диаспоры.",
          "Для успешных продаж критически важно свободно владеть торговой лексикой. Когда покупатель подходит к торговому павильону, продавец вежливо приветствует его: «Добрый день! Что вас интересует? У нас новое поступление отличных зимних курток и тёплых толстовок!»",
          "Если покупатель оптовик, разговор строится вокруг условий партии: «Какой объём вас интересует? Оптом от пяти штук одного артикула цена тысяча двести рублей. Если забираете всю размерную сетку — отдам по тысяче сто. Качество фабричное, наполнитель био-пух, ткань водоотталкивающая».",
          "Для розничных покупателей важен сервис и возможность примерки: «Какой размер вы носите? Попробуйте сорок восьмой, эта модель садится идеально по фигуре. Зеркало и примерочная рядом. Оплатить можно наличными или переводом на карту».",
          "Умение найти общий язык с клиентом, грамотно презентовать товар и уступить в цене — залог высокого торгового оборота."
        ],
        questions: [
          {
            id: "rs1_q1", lang: "ru",
            prompt: "Какие торговые комплексы названы крупнейшими в Москве?",
            hint: "Tên hai khu chợ lớn.",
            options: [
              { key: "A", text: "ГУМ и ЦУМ." },
              { key: "B", text: "Рынок «Садовод» и ТЦ «Москва» в Люблино." },
              { key: "C", text: "Ашан и Мега." },
              { key: "D", text: "Горбушкин двор." }
            ],
            ans: "B", explain: "«Рынок «Садовод» и торговый комплекс «Москва» в Люблино — крупнейшие центры»."
          },
          {
            id: "tl1_q2", lang: "ru",
            prompt: "От скольких штук обычно действует оптовая цена в тексте?",
            hint: "Số lượng tối thiểu mua sỉ.",
            options: [
              { key: "A", text: "От 1 штуки." },
              { key: "B", text: "От 5 штук одного артикула." },
              { key: "C", text: "От 100 штук." },
              { key: "D", text: "Только от 500 штук." }
            ],
            ans: "B", explain: "«Оптом от пяти штук одного артикула цена тысяча двести рублей»."
          },
          {
            id: "tl1_q3", lang: "ru",
            prompt: "Какой материал наполнителя курток упоминается в диалоге?",
            hint: "Chất liệu lót bên trong áo khoác.",
            options: [
              { key: "A", text: "Натуральная шерсть." },
              { key: "B", text: "Био-пух с водоотталкивающей тканью." },
              { key: "C", text: "Синтепон низкого качества." },
              { key: "D", text: "Поролон." }
            ],
            ans: "B", explain: "«Качество фабричное, наполнитель био-пух, ткань водоотталкивающая»."
          },
          {
            id: "tl1_q4", lang: "ru",
            prompt: "Что предложит продавец покупателю, если тот заберёт всю размерную сетку?",
            hint: "Ưu đãi khi mua nguyên ri/dây.",
            options: [
              { key: "A", text: "Повысит цену." },
              { key: "B", text: "Снизит цену до 1100 рублей за штуку." },
              { key: "C", text: "Откажется продавать." },
              { key: "D", text: "Потребует предоплату за год." }
            ],
            ans: "B", explain: "«Если забираете всю размерную сетку — отдам по тысяче сто (1100)»."
          },
          {
            id: "tl1_q5", lang: "ru",
            prompt: "Какие способы оплаты товара доступны покупателям?",
            hint: "Hình thức trả tiền.",
            options: [
              { key: "A", text: "Только золотыми монетами." },
              { key: "B", text: "Наличными или переводом на банковскую карту." },
              { key: "C", text: "Только валютой евро." },
              { key: "D", text: "В кредит на 5 лет." }
            ],
            ans: "B", explain: "«Оплатить можно наличными или переводом на карту»."
          },
          {
            id: "tl1_q6", lang: "vi",
            prompt: "Cụm từ «размерный ряд» (hoặc «линейка») trong buôn bán quần áo nghĩa là gì?",
            hint: "Dây đủ các size.",
            options: [
              { key: "A", text: "Một chiếc áo duy nhất cỡ ngoại cỡ." },
              { key: "B", text: "Một dây/ri hàng gồm đầy đủ các size từ nhỏ đến lớn của cùng một mẫu." },
              { key: "C", text: "Thước dây dùng để đo quần áo." },
              { key: "D", text: "Nhãn mác may trên cổ áo." }
            ],
            ans: "B", explain: "«Размерный ряд» trong bán buôn là nguyên dây (ri) sản phẩm gồm đầy đủ các size từ nhỏ đến lớn."
          },
          {
            id: "tl1_q7", lang: "vi",
            prompt: "Sự khác biệt căn bản giữa «опт» và «розница» là gì?",
            hint: "Bán buôn và bán lẻ.",
            options: [
              { key: "A", text: "Опт là bán lẻ từng cái, розница là bán buôn số lượng lớn." },
              { key: "B", text: "Опт là bán sỉ/buôn số lượng nhiều giá tốt; розница là bán lẻ từng sản phẩm giá cao hơn." },
              { key: "C", text: "Cả hai đều cùng một nghĩa không khác nhau." },
              { key: "D", text: "Опт chỉ dành cho hàng lỗi thanh lý." }
            ],
            ans: "B", explain: "«Опт» = bán sỉ, bán buôn số lượng lớn; «Розница» = bán lẻ từng chiếc."
          },
          {
            id: "tl1_q8", lang: "vi",
            prompt: "Khi khách chê giá đắt, người bán hàng có kinh nghiệm thường làm gì?",
            hint: "Nghệ thuật thương lượng.",
            options: [
              { key: "A", text: "Mắng khách và mời khách đi chỗ khác." },
              { key: "B", text: "Nhấn mạnh chất lượng vải, đường may đẹp và linh hoạt giảm nhẹ giá nếu khách lấy thêm số lượng." },
              { key: "C", text: "Giảm giá 90% ngay lập tức." },
              { key: "D", text: "Đổi sang một mặt hàng rách hỏng." }
            ],
            ans: "B", explain: "Kỹ năng bán hàng: giới thiệu ưu điểm chất liệu (vải cotton, chống thấm, phom dáng) và đề xuất chiết khấu khi mua số lượng."
          },
          {
            id: "tl1_q9", lang: "vi",
            prompt: "Từ «карго» (cargo) trong giới tiểu thương ở Nga thường dùng để chỉ việc gì?",
            hint: "Vận tải hàng hóa.",
            options: [
              { key: "A", text: "Dịch vụ đóng hàng vào bao/thùng và thuê xe tải chuyển hàng đi các tỉnh, thành phố khác." },
              { key: "B", text: "Đi ăn trưa tại căng tin chợ." },
              { key: "C", text: "Treo ma-nơ-canh trước cửa hàng." },
              { key: "D", text: "Thu tiền thuê quầy bán hàng hàng tháng." }
            ],
            ans: "A", explain: "«Карго» là dịch vụ vận chuyển hàng hóa trọn gói từ chợ đầu mối đi các vùng miền hoặc quốc tế."
          },
          {
            id: "tl1_q10", lang: "vi",
            prompt: "Cụm từ «фабричное качество» mang ý nghĩa đảm bảo gì cho khách mua?",
            hint: "Nguồn gốc may mặc.",
            options: [
              { key: "A", text: "Hàng tự may thủ công chắp vá." },
              { key: "B", text: "Hàng sản xuất chuẩn nhà máy quy chuẩn, đường kim mũi chỉ đều và chắc chắn." },
              { key: "C", text: "Hàng lỗi đã qua sử dụng." },
              { key: "D", text: "Hàng tồn kho lâu năm bị mốc." }
            ],
            ans: "B", explain: "«Фабричное качество» là hàng xuất xứ nhà máy chuẩn xịn, chất lượng ổn định."
          }
        ]
      }
    ]
  },
  {
    key: "medicine",
    name: "Y Học & Dược",
    icon: "🏥",
    badge: "Chuyên Ngành Y Dược",
    lessons: [
      {
        id: 1,
        title: "Осмотр пациента терапевтом: Симптомы, анамнез и лечение",
        readingTime: "13 phút",
        source: "Giáo trình Tiếng Nga Y Khoa & Giao tiếp Bác sĩ - Bệnh nhân",
        topic: "Khám lâm sàng: Thăm khám triệu chứng, đo huyết áp, tiền sử bệnh và kê đơn điều trị",
        vocabCategories: [
          {
            title: "Thuật ngữ triệu chứng & Bệnh học (Симптомы и патология)",
            rows: [
              { word: "Жа́лоба пацие́нта", gram: "cụm danh từ", meaning: "Lời khai triệu chứng của người bệnh", example: "Основная жалоба пациента — острая боль в области груди." },
              { word: "Анамне́з боле́зни", gram: "cụm từ", meaning: "Tiền sử bệnh học", example: "Врач тщательно собрал анамнез и историю аллергических реакций." },
              { word: "Артериа́льное давле́ние", gram: "cụm từ", meaning: "Huyết áp động mạch", example: "У пациента повышенное артериальное давление: сто сорок на девяносто." },
              { word: "Воспали́тельный проце́сс", gram: "cụm từ", meaning: "Quá trình viêm nhiễm", example: "Анализ крови показал наличие острого воспалительного процесса." },
              { word: "Диа́гноз", gram: "danh từ nam", meaning: "Chẩn đoán bệnh", example: "Предварительный диагноз — острый бронхит." }
            ]
          },
          {
            title: "Đơn thuốc & Điều trị (Рецепт и терапия)",
            rows: [
              { word: "Реце́пт на лека́рство", gram: "cụm từ", meaning: "Đơn thuốc y khoa", example: "Доктор выписал рецепт на антибиотики широкого спектра действия." },
              { word: "Дозиро́вка препара́та", gram: "cụm từ", meaning: "Liều lượng thuốc dùng", example: "Строго соблюдайте назначенную дозировку: пятьсот миллиграммов в сутки." },
              { word: "Побо́чные эффе́кты", gram: "cụm từ", meaning: "Tác dụng phụ không mong muốn", example: "Среди возможных побочных эффектов отмечается лёгкая сонливость." }
            ]
          }
        ],
        readingText: [
          "В кабинет врача-терапевта городской клинической больницы обратился пациент сорока пяти лет с жалобами на общее недомогание, фебрильную температуру тела до тридцати восьми с половиной градусов и сухой надсадный кашель, продолжающийся четвёртый день.",
          "Врач начал приём со сбора подробного анамнеза. Он уточнил, когда появились первые признаки заболевания, есть ли у пациента сопутствующие хронические заболевания сердечно-сосудистой системы и аллергия на медикаментозные препараты.",
          "Затем доктор провёл физикальный осмотр: измерил артериальное давление и пульс, провёл аускультацию лёгких с помощью стетоскопа. При выслушивании дыхания в нижних отделах правого лёгкого были обнаружены влажные мелкопузырчатые хрипы.",
          "Для верификации клинического диагноза пациенту было назначено рентгенологическое исследование органов грудной клетки и общий клинический анализ крови с лейкоцитарной формулой.",
          "До получения результатов рентгенографии врач назначил антибактериальную терапию, отхаркивающие средства, обильное тёплое питьё и строгий постельный режим в течение пяти дней."
        ],
        questions: [
          {
            id: "med1_q1", lang: "ru",
            prompt: "С какими основными жалобами обратился пациент к врачу?",
            hint: "Triệu chứng của bệnh nhân.",
            options: [
              { key: "A", text: "Боль в колене и хромота." },
              { key: "B", text: "Недомогание, температура 38,5 и сухой кашель." },
              { key: "C", text: "Ухудшение зрения и сыпь на коже." },
              { key: "D", text: "Острая зубная боль." }
            ],
            ans: "B", explain: "«с жалобами на общее недомогание, фебрильную температуру тела до тридцати восьми с половиной градусов и сухой кашель»."
          },
          {
            id: "med1_q2", lang: "ru",
            prompt: "С чего терапевт начал приём больного?",
            hint: "Bước đầu tiên khi khám.",
            options: [
              { key: "A", text: "С немедленной операции." },
              { key: "B", text: "Со сбора подробного анамнеза болезни." },
              { key: "C", text: "С выписки больничного листа без осмотра." },
              { key: "D", text: "С направления в санаторий." }
            ],
            ans: "B", explain: "«Врач начал приём со сбора подробного анамнеза»."
          },
          {
            id: "med1_q3", lang: "ru",
            prompt: "Какой диагностический прибор использовал врач для аускультации лёгких?",
            hint: "Dụng cụ nghe nhịp thở.",
            options: [
              { key: "A", text: "Термометр." },
              { key: "B", text: "Стетоскоп (фонендоскоп)." },
              { key: "C", text: "Шприц." },
              { key: "D", text: "Скальпель." }
            ],
            ans: "B", explain: "«провёл аускультацию лёгких с помощью стетоскопа»."
          },
          {
            id: "med1_q4", lang: "ru",
            prompt: "Какие дополнительные обследования были назначены для уточнения диагноза?",
            hint: "Chỉ định cận lâm sàng.",
            options: [
              { key: "A", text: "МРТ головного мозга." },
              { key: "B", text: "Рентген грудной клетки и клинический анализ крови." },
              { key: "C", text: "УЗИ органов брюшной полости." },
              { key: "D", text: "Гастроскопия желудка." }
            ],
            ans: "B", explain: "«рентгенологическое исследование органов грудной клетки и общий клинический анализ крови»."
          },
          {
            id: "med1_q5", lang: "ru",
            prompt: "Какой режим был предписан больному на ближайшие пять дней?",
            hint: "Chế độ sinh hoạt.",
            options: [
              { key: "A", text: "Занятия в спортзале." },
              { key: "B", text: "Строгий постельный режим и обильное тёплое питьё." },
              { key: "C", text: "Прогулки на морозе." },
              { key: "D", text: "Сверхурочная работа в офисе." }
            ],
            ans: "B", explain: "«обильное тёплое питьё и строгий постельный режим в течение пяти дней»."
          },
          {
            id: "med1_q6", lang: "vi",
            prompt: "Từ «Анамнез» trong thuật ngữ y khoa có nghĩa là gì?",
            hint: "Lịch sử bệnh án.",
            options: [
              { key: "A", text: "Tên một loại thuốc kháng sinh mạnh." },
              { key: "B", text: "Quá trình hỏi để thu thập tiền sử bệnh, triệu chứng và cơ địa dị ứng của bệnh nhân." },
              { key: "C", text: "Hóa đơn viện phí xuất viện." },
              { key: "D", text: "Biên bản phẫu thuật khẩn cấp." }
            ],
            ans: "B", explain: "«Анамнез» (Bệnh sử) là thông tin y khoa do người bệnh cung cấp về diễn tiến bệnh tật và cơ địa trước đó."
          },
          {
            id: "med1_q7", lang: "vi",
            prompt: "Thao tác «Аускультация» (nghe) phổi được thực hiện nhằm phát hiện điều gì?",
            hint: "Âm thanh trong lồng ngực.",
            options: [
              { key: "A", text: "Kiểm tra xem xương sườn có gãy không." },
              { key: "B", text: "Phát hiện tiếng rale ran bệnh lý (khò khè, ẩm, nổ) và rì rào phế nang." },
              { key: "C", text: "Đo dung tích máu trong tim." },
              { key: "D", text: "Kiểm tra nhiệt độ da lồng ngực." }
            ],
            ans: "B", explain: "Thao tác nghe phổi bằng ống nghe (стетоскоп) giúp phát hiện ran rít, ran ngáy, ran ẩm trong các bệnh lý hô hấp."
          },
          {
            id: "med1_q8", lang: "vi",
            prompt: "Vì sao bác sĩ bắt buộc phải hỏi về «аллергия на препараты» (dị ứng thuốc)?",
            hint: "An toàn dùng thuốc.",
            options: [
              { key: "A", text: "Để tăng thêm giá tiền đơn thuốc." },
              { key: "B", text: "Để tránh nguy cơ sốc phản vệ hoặc tai biến dị ứng nguy hiểm đến tính mạng khi kê đơn." },
              { key: "C", text: "Để bệnh nhân tự chọn thuốc theo ý thích." },
              { key: "D", text: "Không có ý nghĩa quan trọng." }
            ],
            ans: "B", explain: "Khai thác tiền sử dị ứng thuốc là nguyên tắc an toàn tối thượng để ngăn ngừa sốc phản vệ."
          },
          {
            id: "med1_q9", lang: "vi",
            prompt: "Chỉ số huyết áp «140/90 mmHg» trong bài phản ánh tình trạng gì?",
            hint: "Huyết áp bình thường hay cao?",
            options: [
              { key: "A", text: "Huyết áp thấp nguy kịch." },
              { key: "B", text: "Huyết áp tăng cao (tiền tăng huyết áp / tăng huyết áp độ 1)." },
              { key: "C", text: "Huyết áp hoàn toàn tối ưu của người trẻ." },
              { key: "D", text: "Mất mạch hoàn toàn." }
            ],
            ans: "B", explain: "Mức 140/90 mmHg là ngưỡng chẩn đoán tăng huyết áp theo chuẩn y tế quốc tế."
          },
          {
            id: "med1_q10", lang: "vi",
            prompt: "Cụm từ «побочные эффекты» trên tờ hướng dẫn sử dụng thuốc có nghĩa là gì?",
            hint: "Tác dụng kèm theo ngoài ý muốn.",
            options: [
              { key: "A", text: "Công dụng chính chữa khỏi bệnh." },
              { key: "B", text: "Tác dụng phụ không mong muốn của dược phẩm." },
              { key: "C", text: "Hạn sử dụng ghi trên vỏ hộp." },
              { key: "D", text: "Tên công ty sản xuất." }
            ],
            ans: "B", explain: "«Побочные эффекты» là tác dụng phụ ngoài ý muốn của thuốc (như buồn ngủ, khô miệng, dị ứng...)."
          }
        ]
      }
    ]
  },
  {
    key: "natural_sciences",
    name: "Khoa Học Tự Nhiên",
    icon: "📐",
    badge: "Chuyên Ngành KHTN (Toán - Lý - Hóa)",
    lessons: [
      {
        id: 1,
        title: "Основы точных наук: Математика, законы механики и строение вещества",
        readingTime: "14 phút",
        source: "Giáo trình Tiếng Nga Khoa học Kỹ thuật dành cho sinh viên đại học",
        topic: "Thuật ngữ KHTN: Phân số, đạo hàm, phương trình, định luật Newton và bảng tuần hoàn nguyên tố",
        vocabCategories: [
          {
            title: "Thuật ngữ Toán học (Математические термины)",
            rows: [
              { word: "Уравне́ние", gram: "danh từ trung", meaning: "Phương trình toán học", example: "Решите квадратное уравнение с одной переменной." },
              { word: "Дробь (числи́тель и знамена́тель)", gram: "cụm từ", meaning: "Phân số (tử số và mẫu số)", example: "В правильной дроби числитель меньше знаменателя." },
              { word: "Произво́дная фу́нкции", gram: "cụm từ", meaning: "Đạo hàm của hàm số", example: "Производная показывает скорость изменения функции." },
              { word: "Интегра́л", gram: "danh từ nam", meaning: "Tích phân", example: "Вычислите определённый интеграл на отрезке от нуля до единицы." }
            ]
          },
          {
            title: "Thuật ngữ Vật lý & Hóa học (Физика и химия)",
            rows: [
              { word: "Зако́н всемирного тяготе́ния", gram: "cụm từ", meaning: "Định luật vạn vật hấp dẫn", example: "Исаак Ньютон сформулировал закон всемирного тяготения." },
              { word: "Ско́рость све́та", gram: "cụm từ", meaning: "Vận tốc ánh sáng trong chân không", example: "Скорость света в вакууме составляет около трёхсот тысяч километров в секунду." },
              { word: "Периоди́ческая табли́ца Менделе́ева", gram: "cụm từ", meaning: "Bảng tuần hoàn các nguyên tố hóa học Mendeleev", example: "Дмитрий Менделеев открыл фундаментальный периодический закон." },
              { word: "Молекуля́рная ма́сса", gram: "cụm từ", meaning: "Khối lượng phân tử", example: "Относительная молекулярная масса воды равна восемнадцати." }
            ]
          }
        ],
        readingText: [
          "Естественные науки — математика, физика и химия — составляют фундамент современного научно-технического прогресса. Для успешной учёбы в технических вузах России иностранным студентам необходимо освоить специальный научный стиль речи.",
          "В математике базовыми понятиями являются переменная величина, функция, предел и производная. Например, квадратное уравнение вида a x квадрат плюс b x плюс c равно нулю решается через нахождение дискриминанта. Если дискриминант больше нуля, уравнение имеет два различных действительных корня.",
          "Классическая механика опирается на три фундаментальных закона Ньютона. Первый закон — закон инерции — гласит, что тело сохраняет состояние покоя или равномерного прямолинейного движения, если на него не действуют внешние силы. Сила равна произведению массы тела на его ускорение: F равно m умножить на a.",
          "В химии величайшим открытием девятнадцатого века стала Периодическая система химических элементов Д. И. Менделеева. Свойства простых веществ и химических соединений находятся в периодической зависимости от заряда их атомных ядер. Вода, состоящая из двух атомов водорода и одного атома кислорода, является универсальным растворителем на Земле.",
          "Овладение языком точных наук открывает путь к передовым инженерным и исследовательским технологиям."
        ],
        questions: [
          {
            id: "ns1_q1", lang: "ru",
            prompt: "Через что решается квадратное уравнение в школьной алгебре?",
            hint: "Công thức giải phương trình bậc 2.",
            options: [
              { key: "A", text: "Через нахождение дискриминанта (D)." },
              { key: "B", text: "Через таблицу умножения." },
              { key: "C", text: "Через интеграл Римана." },
              { key: "D", text: "Через формулу химической реакции." }
            ],
            ans: "A", explain: "«квадратное уравнение решается через нахождение дискриминанта (D = b^2 - 4ac)»."
          },
          {
            id: "ns1_q2", lang: "ru",
            prompt: "Сколько корней имеет квадратное уравнение, если дискриминант строго больше нуля (D > 0)?",
            hint: "Số nghiệm khi D > 0.",
            options: [
              { key: "A", text: "Не имеет корней." },
              { key: "B", text: "Один корень." },
              { key: "C", text: "Два различных действительных корня." },
              { key: "D", text: "Бесконечно много корней." }
            ],
            ans: "C", explain: "«Если дискриминант больше нуля, уравнение имеет два различных действительных корня»."
          },
          {
            id: "ns1_q3", lang: "ru",
            prompt: "Как формулируется второй закон Ньютона (F = ma)?",
            hint: "Công thức lực và gia tốc.",
            options: [
              { key: "A", text: "Сила равна делению массы на скорость." },
              { key: "B", text: "Сила равна произведению массы тела на его ускорение." },
              { key: "C", text: "Сила равна нулю всегда." },
              { key: "D", text: "Масса равна температуре тела." }
            ],
            ans: "B", explain: "«Сила равна произведению массы тела на его ускорение: F = m * a»."
          },
          {
            id: "ns1_q4", lang: "ru",
            prompt: "Кто открыл Периодический закон химических элементов?",
            hint: "Nhà bác học vĩ đại người Nga.",
            options: [
              { key: "A", text: "Исаак Ньютон." },
              { key: "B", text: "Альберт Эйнштейн." },
              { key: "C", text: "Д. И. Менделеев." },
              { key: "D", text: "Михаил Ломоносов." }
            ],
            ans: "C", explain: "«Периодическая система химических элементов Д. И. Менделеева»."
          },
          {
            id: "ns1_q5", lang: "ru",
            prompt: "Из каких атомов состоит молекула воды (H2O)?",
            hint: "Cấu tạo phân tử nước.",
            options: [
              { key: "A", text: "Из двух атомов углерода и одного азота." },
              { key: "B", text: "Из двух атомов водорода и одного атома кислорода." },
              { key: "C", text: "Из железа и серы." },
              { key: "D", text: "Из хлора и натрия." }
            ],
            ans: "B", explain: "«Вода, состоящая из двух атомов водорода (H2) и одного атома кислорода (O)»."
          },
          {
            id: "ns1_q6", lang: "vi",
            prompt: "Trong một phân số toán học (дробь), phần nằm phía trên dấu gạch ngang gọi là gì?",
            hint: "Числитель.",
            options: [
              { key: "A", text: "Mẫu số (знаменатель)." },
              { key: "B", text: "Tử số (числитель)." },
              { key: "C", text: "Thương số." },
              { key: "D", text: "Số dư." }
            ],
            ans: "B", explain: "«Числитель» là tử số, «знаменатель» là mẫu số trong phân số."
          },
          {
            id: "ns1_q7", lang: "vi",
            prompt: "Khái niệm «Производная» trong môn Giải tích thể hiện điều gì?",
            hint: "Đạo hàm.",
            options: [
              { key: "A", text: "Tổng diện tích hình thang cong." },
              { key: "B", text: "Tốc độ biến thiên tức thời của hàm số tại một điểm." },
              { key: "C", text: "Giá trị trung bình cộng." },
              { key: "D", text: "Số lượng nghiệm nguyên." }
            ],
            ans: "B", explain: "Đạo hàm (производная) thể hiện tốc độ biến thiên của hàm số theo biến số."
          },
          {
            id: "ns1_q8", lang: "vi",
            prompt: "Định luật 1 Newton còn có tên gọi phổ biến trong đời sống là định luật gì?",
            hint: "Quán tính.",
            options: [
              { key: "A", text: "Định luật bảo toàn năng lượng." },
              { key: "B", text: "Định luật quán tính (закон инерции)." },
              { key: "C", text: "Định luật phản xạ ánh sáng." },
              { key: "D", text: "Định luật phân rã phóng xạ." }
            ],
            ans: "B", explain: "Định luật 1 Newton là định luật quán tính: vật giữ nguyên trạng thái đứng yên hoặc chuyển động thẳng đều khi không chịu ngoại lực tác dụng."
          },
          {
            id: "ns1_q9", lang: "vi",
            prompt: "Vận tốc ánh sáng trong chân không (Скорость света в вакууме) đạt giá trị xấp xỉ bao nhiêu?",
            hint: "Hằng số vật lý nổi tiếng c.",
            options: [
              { key: "A", text: "Khoảng 300.000 km/h." },
              { key: "B", text: "Khoảng 300.000 km/s (300 nghìn km mỗi giây)." },
              { key: "C", text: "Khoảng 1.000 m/s." },
              { key: "D", text: "Bằng vận tốc âm thanh." }
            ],
            ans: "B", explain: "Hằng số vận tốc ánh sáng c xấp xỉ 300.000 km/s."
          },
          {
            id: "ns1_q10", lang: "vi",
            prompt: "Bảng tuần hoàn Mendeleev sắp xếp các nguyên tố hóa học dựa trên quy luật nào?",
            hint: "Quy luật điện tích hạt nhân.",
            options: [
              { key: "A", text: "Sắp xếp ngẫu nhiên theo bảng chữ cái." },
              { key: "B", text: "Sự biến thiên tuần hoàn theo điện tích hạt nhân nguyên tử (số hiệu nguyên tử Z)." },
              { key: "C", text: "Theo màu sắc bên ngoài của chất." },
              { key: "D", text: "Theo giá thành trên thị trường." }
            ],
            ans: "B", explain: "Tính chất các nguyên tố biến đổi tuần hoàn theo điện tích hạt nhân nguyên tử."
          }
        ]
      }
    ]
  },
  {
    key: "social_sciences",
    name: "Khoa Học Xã Hội",
    icon: "🏛️",
    badge: "Chuyên Ngành KH Xã Hội (Văn - Sử - Luật)",
    lessons: [
      {
        id: 1,
        title: "Русская культура и литература XIX века: Золотой век и великие классики",
        readingTime: "13 phút",
        source: "Giáo trình Văn hóa, Văn học & Lịch sử Nga",
        topic: "Văn hóa & Khoa học xã hội: Thời đại hoàng kim văn học Nga, Pushkin, Tolstoy và giá trị nhân văn",
        vocabCategories: [
          {
            title: "Thuật ngữ Văn học & Nghệ thuật (Литература и искусство)",
            rows: [
              { word: "Литерату́рное направле́ние", gram: "cụm danh từ", meaning: "Trường phái / Trào lưu văn học", example: "Реализм стал ведущим литературным направлением второй половины XIX века." },
              { word: "Произведе́ние иску́сства", gram: "cụm danh từ", meaning: "Tác phẩm nghệ thuật", example: "Роман «Война и мир» — бессмертное произведение мирового искусства." },
              { word: "Главный персона́ж (геро́й)", gram: "cụm từ", meaning: "Nhân vật chính của tác phẩm", example: "Евгений Онегин — типичный герой своего времени." },
              { word: "Сюже́тная ли́ния", gram: "cụm từ", meaning: "Cốt truyện / Tuyến diễn biến", example: "Сюжетная линия повести держит читателя в напряжении." }
            ]
          },
          {
            title: "Lịch sử & Xã hội học (История и общество)",
            rows: [
              { word: "Культу́рное насле́дие", gram: "cụm từ", meaning: "Di sản văn hóa dân tộc", example: "Русский язык является величайшим культурным наследием народа." },
              { word: "Истори́ческая эпо́ха", gram: "cụm từ", meaning: "Thời đại / Kỷ nguyên lịch sử", example: "Эпоха реформ девятнадцатого века изменила облик страны." },
              { word: "Гражда́нское о́бщество", gram: "cụm từ", meaning: "Xã hội dân sự", example: "Развитие институтов гражданского общества и верховенство закона." }
            ]
          }
        ],
        readingText: [
          "Девятнадцатый век по праву называют «Золотым веком» русской словесности и культуры. В эту эпоху были созданы шедевры мировой литературы, оказавшие колоссальное влияние на духовное развитие человечества.",
          "Родоначальником современного русского литературного языка считается великий поэт Александр Сергеевич Пушкин. Его роман в стихах «Евгений Онегин» знаменитый критик В. Г. Белинский назвал «энциклопедией русской жизни». Пушкин объединил богатейший народный язык с высокими европейскими поэтическими традициями.",
          "Вторая половина столетия ознаменовалась расцветом русского психологического реализма. Романы Льва Николаевича Толстого «Война и мир» и «Анна Каренина», а также глубокие философские драмы Фёдора Михайловича Достоевского «Преступление и наказание» и «Братья Карамазовы» исследуют вечные вопросы морали, совести, веры и человеческой души.",
          "Драматургия Антона Павловича Чехова совершила революцию в мировом театральном искусстве, открыв тонкий подтекст и поэзию обыденной человеческой жизни.",
          "Изучение русской литературы и общественных наук позволяет глуб любого понять национальный менталитет, историю и богатство русского языка."
        ],
        questions: [
          {
            id: "ss1_q1", lang: "ru",
            prompt: "Как называют XIX век в истории русской словесности?",
            hint: "Tên gọi thế kỷ 19.",
            options: [
              { key: "A", text: "«Серебряным веком»." },
              { key: "B", text: "«Золотым веком» русской литературы." },
              { key: "C", text: "«Бронзовым веком»." },
              { key: "D", text: "«Эпохой застоя»." }
            ],
            ans: "B", explain: "«Девятнадцатый век по праву называют «Золотым веком» русской словесности»."
          },
          {
            id: "ss1_q2", lang: "ru",
            prompt: "Кто считается основоположником современного русского литературного языка?",
            hint: "Mặt trời của thi ca Nga.",
            options: [
              { key: "A", text: "Н. В. Гоголь." },
              { key: "B", text: "А. С. Пушкин." },
              { key: "C", text: "И. С. Тургенев." },
              { key: "D", text: "М. Ю. Лермонтов." }
            ],
            ans: "B", explain: "«Родоначальником современного русского литературного языка считается А. С. Пушкин»."
          },
          {
            id: "ss1_q3", lang: "ru",
            prompt: "Какое произведение критик Белинский назвал «энциклопедией русской жизни»?",
            hint: "Tác phẩm nổi tiếng của Pushkin.",
            options: [
              { key: "A", text: "«Капитанская дочка»." },
              { key: "B", text: "«Евгений Онегин»." },
              { key: "C", text: "«Медный всадник»." },
              { key: "D", text: "«Руслан и Людмила»." }
            ],
            ans: "B", explain: "«Его роман в стихах «Евгений Онегин» Белинский назвал «энциклопедией русской жизни»»."
          },
          {
            id: "ss1_q4", lang: "ru",
            prompt: "Какое литературное направление стало ведущим во второй половине XIX века?",
            hint: "Trào lưu văn học.",
            options: [
              { key: "A", text: "Классицизм." },
              { key: "B", text: "Психологический реализм." },
              { key: "C", text: "Сюрреализм." },
              { key: "D", text: "Футуризм." }
            ],
            ans: "B", explain: "«Вторая половина столетия ознаменовалась расцветом русского психологического реализма»."
          },
          {
            id: "ss1_q5", lang: "ru",
            prompt: "Какие бессмертные романы создал Лев Николаевич Толстой?",
            hint: "Đại tác phẩm của Lev Tolstoy.",
            options: [
              { key: "A", text: "«Война и мир» и «Анна Каренина»." },
              { key: "B", text: "«Преступление и наказание»." },
              { key: "C", text: "«Вишнёвый сад»." },
              { key: "D", text: "«Мёртвые души»." }
            ],
            ans: "A", explain: "«Романы Льва Николаевича Толстого «Война и мир» и «Анна Каренина»»."
          },
          {
            id: "ss1_q6", lang: "vi",
            prompt: "Ý nghĩa của thuật ngữ «Реализм» (Chủ nghĩa hiện thực) trong văn học là gì?",
            hint: "Phương pháp sáng tác.",
            options: [
              { key: "A", text: "Tưởng tượng ra những thế giới phép thuật không có thật." },
              { key: "B", text: "Miêu tả chân thực, khách quan đời sống xã hội và nội tâm con người đúng như nó tồn tại." },
              { key: "C", text: "Chỉ viết về cuộc đời các vị vua chúa cổ đại." },
              { key: "D", text: "Chỉ sáng tác bằng các thể thơ siêu thực trừu tượng." }
            ],
            ans: "B", explain: "Chủ nghĩa hiện thực (Реализм) hướng tới phản ánh khách quan, chân thực và sâu sắc các quy luật của hiện thực đời sống và tính cách điển hình trong hoàn cảnh điển hình."
          },
          {
            id: "ss1_q7", lang: "vi",
            prompt: "Tác phẩm «Tội ác và trừng phạt» (Преступление и наказание) của Dostoevsky xoay quanh vấn đề gì?",
            hint: "Tư tưởng triết học & lương tâm.",
            options: [
              { key: "A", text: "Cách trồng hoa trong trang trại." },
              { key: "B", text: "Sự giằng xé lương tâm, tội lỗi và sự cứu rỗi linh hồn của nhân vật Raskolnikov." },
              { key: "C", text: "Kỹ thuật hàng hải thời kỳ cận đại." },
              { key: "D", text: "Lịch sử kiến trúc thành phố Moskva." }
            ],
            ans: "B", explain: "«Tội ác và trừng phạt» phân tích sâu sắc bi kịch lương tâm, tư tưởng siêu nhân và con đường chuộc tội của Raskolnikov."
          },
          {
            id: "ss1_q8", lang: "vi",
            prompt: "Nhà văn Anton Chekhov đã tạo nên cuộc cách mạng trong lĩnh vực nghệ thuật nào?",
            hint: "Sân khấu kịch.",
            options: [
              { key: "A", text: "Điêu khắc đá cẩm thạch." },
              { key: "B", text: "Nghệ thuật sân khấu kịch nói thế giới (kịch tâm lý đời thường)." },
              { key: "C", text: "Hội họa sơn dầu phong cảnh." },
              { key: "D", text: "Múa ba-lê cổ điển." }
            ],
            ans: "B", explain: "Chekhov là bậc thầy cách tân nghệ thuật kịch hiện đại (với các vở như Cậu Vanya, Vườn anh đào, Hải âu)."
          },
          {
            id: "ss1_q9", lang: "vi",
            prompt: "Cụm từ «Культурное наследие» có ý nghĩa xã hội gì?",
            hint: "Di sản văn hóa.",
            options: [
              { key: "A", text: "Tiền gửi tiết kiệm ngân hàng." },
              { key: "B", text: "Di sản văn hóa vật thể và phi vật thể do các thế hệ tiền nhân để lại cho hậu thế." },
              { key: "C", text: "Hàng hóa tiêu dùng ngắn hạn." },
              { key: "D", text: "Quy tắc lái xe trên đường cao tốc." }
            ],
            ans: "B", explain: "«Культурное наследие» (Di sản văn hóa) là những giá trị tinh thần, ngôn ngữ, nghệ thuật được gìn giữ và kế thừa qua các thế hệ."
          },
          {
            id: "ss1_q10", lang: "vi",
            prompt: "Vì sao việc đọc tác phẩm văn học nguyên bản tiếng Nga lại quan trọng đối với người học chuyên sâu?",
            hint: "Cảm thụ ngôn ngữ.",
            options: [
              { key: "A", text: "Để lấy bằng lái xe." },
              { key: "B", text: "Giúp cảm nhận được chiều sâu tư duy, sắc thái ngôn ngữ tinh tế và tâm hồn Nga." },
              { key: "C", text: "Chỉ để học thuộc lòng số liệu thống kê." },
              { key: "D", text: "Không đem lại lợi ích gì thêm." }
            ],
            ans: "B", explain: "Đọc nguyên tác văn học giúp nắm bắt linh hồn ngôn ngữ, vốn từ phong phú và chiều sâu văn hóa bản địa."
          }
        ]
      }
    ]
  }
];

// Tiến hành tạo thư mục và file cho từng chuyên ngành
domains.forEach(d => {
  const domainDir = path.join(baseDir, d.key);
  if (!fs.existsSync(domainDir)) {
    fs.mkdirSync(domainDir, { recursive: true });
  }

  const manifest = [];
  d.lessons.forEach(l => {
    const lessonData = {
      id: l.id,
      level: "specialized",
      domain: d.key,
      domainName: d.name,
      badge: d.badge,
      title: l.title,
      readingTime: l.readingTime,
      source: l.source,
      topic: l.topic,
      vocabCategories: l.vocabCategories,
      readingText: l.readingText,
      questions: l.questions
    };

    const fileName = `lesson-${String(l.id).padStart(2, '0')}.json`;
    const filePath = path.join(domainDir, fileName);
    fs.writeFileSync(filePath, JSON.stringify(lessonData, null, 2), 'utf8');

    manifest.push({
      id: l.id,
      file: fileName,
      title: l.title,
      badge: d.badge,
      readingTime: l.readingTime,
      source: l.source,
      topic: l.topic,
      questionCount: l.questions.length
    });
  });

  const manifestPath = path.join(domainDir, 'manifest.json');
  fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2), 'utf8');
  console.log(`✓ Đã tạo chuyên ngành [${d.name}]: ${d.lessons.length} bài tại ${domainDir}`);
});

console.log('✓ Hoàn tất tạo 6 chuyên ngành tiếng Nga thực chiến!');

