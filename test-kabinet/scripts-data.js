/**
 * Тест «1С:Кабинет сотрудника» — ГК Форус
 * 15 видов специалистов по триггерам (автоматизация / КЭДО / обращения).
 * Должность только подставляется в заголовок: «Кадровик на детоксе», «Бухгалтер на детоксе».
 */
window.QUIZ_SCRIPTS = {
  brand: "1С:Кабинет сотрудника",
  company: "ГК Форус",
  logoSrc: "assets/brand/logo-forus.png",
  publicPath: "test-kabinet",
  shareBaseUrl:
    "https://htmlpreview.github.io/?https://cdn.jsdelivr.net/gh/daaanilNikonov/for-lisa@cursor/quiz-landing-scripts-55dd/test-kabinet/index.html",

  tempImages: [
    "assets/results/temp-1.jpeg",
    "assets/results/temp-2.jpeg",
    "assets/results/temp-3.jpeg",
  ],

  roleTitles: {
    kadrovik: "Кадровик",
    buh_zp: "Бухгалтер",
    glb: "Главный бухгалтер",
    ruk_otdela: "Руководитель отдела",
    ruk_kompanii: "Руководитель компании",
    drugoe: "Специалист",
  },

  questions: [
    {
      id: "role",
      title: "Ваша должность:",
      options: [
        { id: "kadrovik", label: "кадровик" },
        { id: "buh_zp", label: "бухгалтер по расчету ЗП" },
        { id: "glb", label: "главный бухгалтер" },
        { id: "ruk_otdela", label: "руководитель отдела" },
        { id: "ruk_kompanii", label: "руководитель компании" },
        { id: "drugoe", label: "другое" },
      ],
    },
    {
      id: "size",
      title: "Численность персонала в вашей компании:",
      options: [
        { id: "s50", label: "до 50" },
        { id: "s50_100", label: "50–100" },
        { id: "s100_500", label: "100–500" },
        { id: "s500_2000", label: "500–2000" },
        { id: "s2000", label: "2000 и более" },
      ],
    },
    {
      id: "docs",
      title: "Как вы работаете с расчетными и приказами:",
      options: [
        { id: "paper", label: "все документы на бумаге" },
        { id: "email", label: "передача по электронной почте или мессенджеры" },
        { id: "kedo", label: "есть КЭДО" },
        { id: "none", label: "нет таких документов" },
      ],
    },
    {
      id: "certs",
      title: "Как часто ко мне приходят сотрудники за справками:",
      options: [
        { id: "never", label: "не обращаются" },
        { id: "daily", label: "обращаются ежедневно" },
        { id: "rare", label: "очень редко" },
      ],
    },
    {
      id: "auto",
      title: "Уровень автоматизации моей работы:",
      options: [
        { id: "low", label: "до 40% автоматизировано" },
        { id: "mid", label: "40–80% автоматизировано" },
        { id: "high", label: "более 80% автоматизировано" },
      ],
    },
    {
      id: "overtime",
      title: "Часто ли вам приходится задерживаться на работе, чтобы доделать задачи:",
      options: [
        { id: "always", label: "почти всегда" },
        { id: "monthly", label: "раз в месяц" },
        { id: "rare", label: "редко" },
        { id: "never", label: "никогда" },
      ],
    },
    {
      id: "report",
      title: "Когда сдаете кадровую отчетность:",
      options: [
        { id: "now", label: "сразу" },
        { id: "month", label: "в течение месяца" },
        { id: "last_day", label: "в последний день" },
      ],
    },
  ],

  /**
   * Яндекс.Форма «Форма сбора заявок на КЭДО»
   * https://forms.yandex.ru/u/6a83b54feb614605f98e10ee
   *
   * Отправка с лендинга идёт через proxyUrl (CSRF Яндекса нельзя пройти из браузера).
   * Локально: node proxy/server.mjs → proxyUrl: "/api/yandex-submit"
   * Пункт 4 «Должность» заполняется из ответа на 1-й вопрос квиза.
   */
  yandexForm: {
    url: "https://forms.yandex.ru/u/6a83b54feb614605f98e10ee",
    surveyId: "6a83b54feb614605f98e10ee",
    /** Относительный путь, если лендинг открыт через proxy/server.mjs */
    proxyUrl: "", // опционально: "/api/yandex-submit" при запуске proxy/server.mjs
    fields: {
      name: "answer_short_text_9008986271269920",
      inn: "answer_short_text_9008986271288132",
      phone: "answer_short_text_9008986271302340",
      position: "answer_short_text_9008986271334516",
    },
  },

  kedoVendorQuestion: {
    id: "kedo_vendor",
    title: "Каким КЭДО вы пользуетесь?",
    options: [
      { id: "kabinet", label: "1С:Кабинет сотрудника" },
      { id: "nopaper", label: "Nopaper" },
      { id: "topfactor", label: "Топ Фактор" },
      { id: "directum", label: "Directum" },
      { id: "drugoe", label: "Другое" },
    ],
    allowCustom: true,
    customPlaceholder: "Если другое — впишите название КЭДО",
  },

  /**
   * 15 видов. match не зависит от должности.
   * text — текст в плашке результата
   * imagePrompt — описание для нейросети (генерация картинки)
   */
  scenarios: [
    {
      id: "detoks",
      shortName: "на детоксе",
      titlePattern: "{role} на детоксе",
      priority: 10,
      triggersNote: ">80% автоматизации + КЭДО + справки редко/не обращаются",
      text:
        "Процессы дышат сами: документы уходят в цифре, сотрудники почти не дергают, а вы успеваете думать о развитии, а не о стопке заявлений. Вы на редком «детоксе» от рутины — и это ваш главный актив.",
      imagePrompt:
        "Яркая современная иллюстрация офиса в дружелюбном cartoon-стиле. Специалист спокойно сидит за чистым столом у большого монитора с дашбордом 1С и зелёными галочками КЭДО, рядом аккуратный ноутбук и чашка кофе. На фоне коллеги с телефонами показывают подтверждения подписи, мало бумаг, много света, растения, ощущение порядка и лёгкости. Без логотипов брендов, горизонтальный формат 16:9.",
      imageSrc: "assets/results/detoks.jpg",
      match(a) {
        return (
          a.auto === "high" &&
          a.docs === "kedo" &&
          (a.certs === "rare" || a.certs === "never")
        );
      },
    },
    {
      id: "nevidimka",
      shortName: "невидимка сервиса",
      titlePattern: "{role} — невидимка сервиса",
      priority: 15,
      triggersNote: ">80% автоматизации + за справками не обращаются + без КЭДО",
      text:
        "Вас почти не видно в очереди у кабинета — и это комплимент системе. Самообслуживание и автоматизация забрали поток обращений, а вы работаете «за кулисами», где процессы уже не требуют постоянного присутствия героя у принтера.",
      imagePrompt:
        "Иллюстрация: специалист в полупустом светлом open-space работает за двумя мониторами с графиками и статусами заявок «выполнено», кабинетная дверь закрыта, в коридоре никого нет. На столе минимум бумаг, мягкий дневной свет, спокойная композиция, cartoon-корпоративный стиль, 16:9.",
      imageSrc: "assets/results/nevidimka.jpg",
      match(a) {
        return a.auto === "high" && a.certs === "never" && a.docs !== "kedo";
      },
    },
    {
      id: "favorit",
      shortName: "офисный фаворит",
      titlePattern: "{role} — офисный фаворит",
      priority: 20,
      triggersNote: ">80% автоматизации + справки ежедневно",
      text:
        "Цифра у вас уже есть, но люди всё равно идут «к вам лично». Вы — точка доверия офиса: автоматизация закрыла часть задач, а человеческий магнит остался. Пора направить поток в сервис, не теряя статуса надёжного человека.",
      imagePrompt:
        "Тёплый cartoon-офис: улыбчивый специалист за аккуратным столом с современным компьютером (высокий процент автоматизации на экране), вокруг стоят 3–4 коллеги с вопросами и бланками справок, атмосфера доброжелательная, не хаотичная. Дневной свет, растения, 16:9.",
      imageSrc: "assets/results/favorit.jpg",
      match(a) {
        return a.auto === "high" && a.certs === "daily";
      },
    },
    {
      id: "tsifrovoy_dirizher",
      shortName: "цифровой дирижёр",
      titlePattern: "{role} — цифровой дирижёр",
      priority: 25,
      triggersNote: ">80% автоматизации + есть КЭДО",
      text:
        "Вы дирижируете электронным контуром: КЭДО, статусы, маршруты. Рутина уже не рушит день, и фокус смещается на контроль качества процессов. Следующий уровень — чтобы оркестр играл без постоянных взмахов вашей руки.",
      imagePrompt:
        "Специалист за широким столом как дирижёр цифровой сцены: на мониторах схемы согласований КЭДО, прогресс-бары, зелёные статусы. На заднем плане полупрозрачные иконки документов и телефоны сотрудников с подписями. Чистый современный офис, cartoon-стиль, 16:9.",
      imageSrc: "assets/results/tsifrovoy_dirizher.jpg",
      match(a) {
        return a.auto === "high" && a.docs === "kedo";
      },
    },
    {
      id: "nochnoy_dirizher",
      shortName: "ночной дирижёр",
      titlePattern: "{role} — ночной дирижёр",
      priority: 30,
      triggersNote: "Задерживается почти всегда + автоматизация не высокая",
      text:
        "Когда остальные уходят, у вас только начинается вторая смена. Вы «дирижёр ночи»: сводите хвосты дня, пока офис пустеет. Это не героизм навсегда — сигнал, что процесс держится на вашем личном времени.",
      imagePrompt:
        "Ночной офис за окном тёмный город, настенные часы около 21:00, специалист за столом в мягком свете монитора, вокруг стопки папок и стикеры с задачами, усталость без гротеска, сосредоточенный взгляд. Cartoon-иллюстрация, тёплые лампы и холодный свет экрана, 16:9.",
      imageSrc: "assets/results/nochnoy_dirizher.jpg",
      match(a) {
        return a.overtime === "always" && a.auto !== "high";
      },
    },
    {
      id: "khranitel_dedlayna",
      shortName: "хранитель дедлайна",
      titlePattern: "{role} — хранитель дедлайна",
      priority: 35,
      triggersNote: "Автоматизация до 40% + отчётность в последний день + ежедневно/почти всегда",
      text:
        "Вы живёте у края календаря: отчётность в последний день, поток обращений не стихает, автоматизация почти не страхует. Это режим «успеть любой ценой». Вам нужен не ещё один подвиг, а система, которая переносит дедлайн с ваших плеч на процесс.",
      imagePrompt:
        "Динамичная иллюстрация: специалист в окружении красных отметок «сегодня крайний срок», календарь с обведённой датой, стопки отчётов, коллеги с срочными просьбами, на экране 1С низкий процент автоматизации. Напряжение, но герой собран. Cartoon-стиль, 16:9.",
      imageSrc: "assets/results/khranitel_dedlayna.jpg",
      match(a) {
        return (
          a.auto === "low" &&
          a.report === "last_day" &&
          (a.certs === "daily" || a.overtime === "always")
        );
      },
    },
    {
      id: "bumazhnyy_alkhimik",
      shortName: "бумажный алхимик",
      titlePattern: "{role} — бумажный алхимик",
      priority: 40,
      triggersNote: "Все документы на бумаге + автоматизация не высокая",
      text:
        "Вы превращаете кипы листов в результат почти вручную: подписи, обходы, архив. В этом есть мастерство — и огромная цена времени. Цифровой контур здесь не про «модно», а про то, чтобы алхимия не съедала весь рабочий день.",
      imagePrompt:
        "Кабинет-архив: высокие стеллажи с папками, стол завален документами и печатями, специалист подписывает бумаги, рядом лоток «на подпись», принтер, мало цифры на экране. Тёплые тона бумаги и дерева, cartoon-стиль, 16:9.",
      imageSrc: "assets/results/bumazhnyy_alkhimik.jpg",
      match(a) {
        return a.docs === "paper" && a.auto !== "high";
      },
    },
    {
      id: "okhotnik_podpisey",
      shortName: "охотник за подписями",
      titlePattern: "{role} — охотник за подписями",
      priority: 45,
      triggersNote: "Бумага или мессенджеры + справки ежедневно",
      text:
        "Ваш день — квест «найти подписанта». Справки каждый день, документы гуляют по кабинетам или чатам. Вы мастер коротких рейдов по этажам и переписке. Кабинет сотрудника и КЭДО как раз про то, чтобы охота закончилась, а сервис остался.",
      imagePrompt:
        "Специалист с планшетом/папкой идёт по офисному коридору, коллеги выглядывают из кабинетов с ручками и документами на подпись, на фоне мессенджер-уведомления и стопка заявлений. Живая динамика, cartoon, 16:9.",
      imageSrc: "assets/results/okhotnik_podpisey.jpg",
      match(a) {
        return (
          (a.docs === "paper" || a.docs === "email") && a.certs === "daily"
        );
      },
    },
    {
      id: "letopisets",
      shortName: "летописец заявок",
      titlePattern: "{role} — летописец заявок",
      priority: 50,
      triggersNote: "Автоматизация 40–80% + справки ежедневно + отчётность сразу",
      text:
        "Поток заявок вы ведёте как летопись: всё фиксируете, всё успеваете «сразу», хотя люди приходят каждый день. Средняя автоматизация держит вас на плаву, но не снимает ручной слой. Пора, чтобы заявки писались в систему, а не только в вашу память.",
      imagePrompt:
        "Специалист за компьютером, экран 1С с прогрессом около 60%, вокруг очередь коллег со справками, на столе блокнот-журнал заявок и аккуратные стопки папок. Вечерний свет из окна, cartoon-стиль, 16:9.",
      imageSrc: "assets/results/letopisets.jpg",
      match(a) {
        return a.auto === "mid" && a.certs === "daily" && a.report === "now";
      },
    },
    {
      id: "messenger_orkestr",
      shortName: "мессенджер-оркестр",
      titlePattern: "{role} — мессенджер-оркестр",
      priority: 55,
      triggersNote: "Документы через email/мессенджеры",
      text:
        "Ваш основной канал — чаты и почта: быстро, привычно и… хаотично. Важные файлы живут в переписке, а не в процессе. Оркестр уведомлений звучит громко; нужен единый зал — сервис, где документы не теряются между стикерами и тредами.",
      imagePrompt:
        "Специалист за ноутбуком, вокруг парят окна мессенджеров и email с вложениями PDF, телефон вибрирует, на столе наушники и стикеры «срочно». Современный офис, лёгкий хаос цифровых окон, cartoon, 16:9.",
      imageSrc: "assets/results/messenger_orkestr.jpg",
      match(a) {
        return a.docs === "email";
      },
    },
    {
      id: "mayak_kraynego_dnya",
      shortName: "маяк крайнего дня",
      titlePattern: "{role} — маяк крайнего дня",
      priority: 60,
      triggersNote: "Отчётность в последний день",
      text:
        "Отчётность вспыхивает у вас как маяк в последний день срока: весь свет и всё внимание — туда. Так можно жить годами, но цена — нервы и риск ошибки. Стабильный ритм «в течение месяца / сразу» начинается с автоматизации подготовки данных.",
      imagePrompt:
        "Специалист у календаря с крупной красной датой «крайний день», на столе отчётные формы, таймер, чашка остывшего чая, за окном вечер. Напряжённая, но героическая сцена, cartoon-корпоративный стиль, 16:9.",
      imageSrc: "assets/results/mayak_kraynego_dnya.jpg",
      match(a) {
        return a.report === "last_day";
      },
    },
    {
      id: "sprinter",
      shortName: "спринтер отчётности",
      titlePattern: "{role} — спринтер отчётности",
      priority: 65,
      triggersNote: "Отчётность сразу + автоматизация не максимальная",
      text:
        "Вы сдаёте отчётность сразу — дисциплина сильная. Но без полной автоматизации каждый спринт всё равно требует ручного разгона. Сохраним вашу скорость и уберём лишние круги по данным и справкам.",
      imagePrompt:
        "Динамичная композиция: специалист уверенно нажимает «отправить отчёт» на мониторе, рядом чеклист с галочками, лёгкое движение бумаг и интерфейсных окон. Светлый офис, энергия без хаоса, cartoon, 16:9.",
      imageSrc: "assets/results/sprinter.jpg",
      match(a) {
        return a.report === "now" && a.auto !== "high";
      },
    },
    {
      id: "balansir",
      shortName: "балансир на канате",
      titlePattern: "{role} — балансир на канате",
      priority: 70,
      triggersNote: "Автоматизация 40–80%",
      text:
        "Вы посередине пути: часть процессов уже в системе, часть ещё на ручнике. Это тонкий канат — можно сорваться в аврал или дойти до устойчивой цифры. Вам нужен следующий понятный шаг автоматизации, а не революция за одну ночь.",
      imagePrompt:
        "Метафоричная офисная иллюстрация: специалист идёт по канату между стопкой бумаг слева и светящимся экраном 1С справа, баланс, сосредоточенность. Современный cartoon, мягкие цвета, 16:9.",
      imageSrc: "assets/results/balansir.jpg",
      match(a) {
        return a.auto === "mid";
      },
    },
    {
      id: "tihaya_gavan",
      shortName: "тихая гавань",
      titlePattern: "{role} — тихая гавань",
      priority: 68,
      triggersNote: "Справки редко + задержки редко/никогда + не последний день",
      text:
        "У вас относительно тихо: редкие обращения, редкие задержки, отчётность не горит в последний день. Это хорошая база. Чтобы гавань не превратилась в штиль без развития, зафиксируйте процессы в сервисе самообслуживания и КЭДО.",
      imagePrompt:
        "Спокойный светлый кабинет, специалист пьёт чай у монитора с зелёными статусами задач, на подоконнике растение, за дверью пустой коридор, минимум бумаг. Умиротворённая cartoon-иллюстрация, 16:9.",
      imageSrc: "assets/results/tihaya_gavan.jpg",
      match(a) {
        return (
          (a.certs === "rare" || a.certs === "never") &&
          (a.overtime === "rare" || a.overtime === "never") &&
          a.report !== "last_day"
        );
      },
    },
    {
      id: "issledovatel",
      shortName: "исследователь процессов",
      titlePattern: "{role} — исследователь процессов",
      priority: 100,
      triggersNote: "Запасной сценарий / смешанный профиль",
      text:
        "Ваш профиль смешанный: вы уже пробуете разные способы работы с документами и людьми. Это позиция исследователя — идеальный момент выбрать единый контур: меньше ручных маршрутов, больше понятного сервиса для сотрудников.",
      imagePrompt:
        "Специалист за столом с ноутбуком, лупой над схемой процессов (блок-схема HR/зарплата), рядом немного бумаг и планшет, поза любопытная и вовлечённая. Светлый офис, cartoon-стиль, 16:9.",
      imageSrc: "assets/results/issledovatel.jpg",
      match() {
        return true;
      },
    },
  ],
};

function painScore(a) {
  let s = 0;
  if (a.certs === "daily") s += 1;
  if (a.overtime === "always") s += 1;
  if (a.report === "last_day") s += 1;
  return s;
}

function calmScore(a) {
  let s = 0;
  if (a.auto === "high") s += 1;
  if (a.docs === "kedo") s += 1;
  if (a.report === "now") s += 1;
  if (a.certs === "rare" || a.certs === "never") s += 1;
  return s;
}

function formatScenarioTitle(scenario, roleId) {
  const roleLabel =
    window.QUIZ_SCRIPTS.roleTitles[roleId] ||
    window.QUIZ_SCRIPTS.roleTitles.drugoe;
  const pattern =
    scenario.titlePattern || "{role} — " + (scenario.shortName || scenario.id);
  return pattern.replaceAll("{role}", roleLabel);
}

function buildQuestionQueue(answers) {
  const queue = [];
  for (const q of window.QUIZ_SCRIPTS.questions) {
    queue.push(q);
    if (q.id === "docs" && answers.docs === "kedo") {
      queue.push(window.QUIZ_SCRIPTS.kedoVendorQuestion);
    }
  }
  return queue;
}

window.QUIZ_MATCH = {
  painScore,
  calmScore,
  formatScenarioTitle,
  buildQuestionQueue,
  resolve(answers) {
    const list = window.QUIZ_SCRIPTS.scenarios
      .filter((sc) => sc.match(answers))
      .sort((a, b) => a.priority - b.priority);
    return list[0] || window.QUIZ_SCRIPTS.scenarios.at(-1);
  },
};

window.QUIZ_PROFILE_LABELS = {
  role: Object.fromEntries(
    window.QUIZ_SCRIPTS.questions[0].options.map((o) => [o.id, o.label])
  ),
  size: Object.fromEntries(
    window.QUIZ_SCRIPTS.questions[1].options.map((o) => [o.id, o.label])
  ),
  docs: {
    paper: "Документы на бумаге",
    email: "Документы по email/мессенджерам",
    kedo: "Уже есть КЭДО",
    none: "Нет расчётных/приказов в работе",
  },
  kedo_vendor: {
    kabinet: "КЭДО: 1С:Кабинет сотрудника",
    nopaper: "КЭДО: Nopaper",
    topfactor: "КЭДО: Топ Фактор",
    directum: "КЭДО: Directum",
    drugoe: "КЭДО: другое",
  },
  certs: {
    never: "За справками не обращаются",
    daily: "Справки ежедневно",
    rare: "Справки очень редко",
  },
  auto: {
    low: "Автоматизация до 40%",
    mid: "Автоматизация 40–80%",
    high: "Автоматизация более 80%",
  },
  overtime: {
    always: "Задерживаются почти всегда",
    monthly: "Задерживаются раз в месяц",
    rare: "Задерживаются редко",
    never: "Не задерживаются",
  },
  report: {
    now: "Отчётность сразу",
    month: "Отчётность в течение месяца",
    last_day: "Отчётность в последний день",
  },
};
