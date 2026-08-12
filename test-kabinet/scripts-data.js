/**
 * Тест «1С:Кабинет сотрудника» — ГК Форус
 * 10 сценариев по триггерам. Название сценария зависит от должности (Q1).
 */
window.QUIZ_SCRIPTS = {
  brand: "1С:Кабинет сотрудника",
  company: "ГК Форус",
  logoSrc: "assets/brand/logo-forus.png",
  publicPath: "test-kabinet",
  shareBaseUrl:
    "https://htmlpreview.github.io/?https://github.com/daaanilNikonov/for-lisa/blob/cursor/quiz-landing-scripts-55dd/test-kabinet/index.html",

  tempImages: [
    "assets/results/temp-1.jpeg",
    "assets/results/temp-2.jpeg",
    "assets/results/temp-3.jpeg",
  ],

  /** Короткие роли для названия сценария */
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
        { id: "glb", label: "глб" },
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

  /** Условный вопрос после «есть КЭДО» */
  kedoVendorQuestion: {
    id: "kedo_vendor",
    title: "Каким КЭДО вы пользуетесь?",
    options: [
      { id: "v1", label: "1" },
      { id: "v2", label: "2" },
      { id: "v3", label: "3" },
      { id: "v4", label: "4" },
    ],
    allowCustom: true,
    customPlaceholder: "Или впишите название вашей системы КЭДО",
  },

  /**
   * titlePattern: {role} подставляется из roleTitles.
   * shortName — хвост персонажа.
   */
  scenarios: [
    {
      id: "detoks",
      shortName: "на детоксе",
      titlePattern: "{role} на детоксе",
      group: "hr",
      priority: 10,
      text: "",
      imageSrc: "",
      triggersNote:
        "Автоматизация >80% + есть КЭДО + справки редко/не обращаются. Численность не важна.",
      match(a) {
        return (
          isHrLike(a.role) &&
          a.auto === "high" &&
          a.docs === "kedo" &&
          (a.certs === "rare" || a.certs === "never")
        );
      },
    },
    {
      id: "lyubimchik",
      shortName: "любимчик",
      titlePattern: "{role}-любимчик",
      group: "hr",
      priority: 20,
      text: "",
      imageSrc: "",
      triggersNote:
        "Автоматизация >80%, но сотрудники всё равно часто приходят (ежедневно) или почти всегда задержки.",
      match(a) {
        return (
          isHrLike(a.role) &&
          a.auto === "high" &&
          (a.certs === "daily" || a.overtime === "always")
        );
      },
    },
    {
      id: "sorvigolova",
      shortName: "сорвиголова",
      titlePattern: "{role}-сорвиголова",
      group: "hr",
      priority: 30,
      text: "",
      imageSrc: "",
      triggersNote:
        "Автоматизация до 40% + частые обращения/задержки + отчётность в последний день.",
      match(a) {
        return (
          isHrLike(a.role) &&
          a.auto === "low" &&
          (a.certs === "daily" || a.overtime === "always") &&
          a.report === "last_day"
        );
      },
    },
    {
      id: "knizhnyy_cherv",
      shortName: "книжный червь",
      titlePattern: "{role} — книжный червь",
      group: "hr",
      priority: 40,
      text: "",
      imageSrc: "",
      triggersNote:
        "Часто обращаются за справками, отчётность сдаёт сразу, автоматизация 40–80%.",
      match(a) {
        return (
          isHrLike(a.role) &&
          a.auto === "mid" &&
          a.certs === "daily" &&
          a.report === "now"
        );
      },
    },
    {
      id: "pozharnyy",
      shortName: "пожарный",
      titlePattern: "{role}-пожарный",
      group: "hr",
      priority: 50,
      text: "",
      imageSrc: "",
      triggersNote:
        "Почти всегда задерживается и/или сдаёт в последний день при низкой/средней автоматизации.",
      match(a) {
        return (
          isHrLike(a.role) &&
          a.auto !== "high" &&
          (a.overtime === "always" || a.report === "last_day") &&
          (a.certs === "daily" || a.overtime === "always")
        );
      },
    },
    {
      id: "bumazhnyy_mag",
      shortName: "бумажный маг",
      titlePattern: "{role} — бумажный маг",
      group: "hr",
      priority: 60,
      text: "",
      imageSrc: "",
      triggersNote: "Все документы на бумаге, автоматизация не высокая.",
      match(a) {
        return isHrLike(a.role) && a.docs === "paper" && a.auto !== "high";
      },
    },
    {
      id: "tihaya_gavan",
      shortName: "тихая гавань",
      titlePattern: "{role} — тихая гавань",
      group: "hr",
      priority: 70,
      text: "",
      imageSrc: "",
      triggersNote:
        "Справки редко, задержки редко/никогда, отчётность не в последний день.",
      match(a) {
        return (
          isHrLike(a.role) &&
          (a.certs === "rare" || a.certs === "never") &&
          (a.overtime === "rare" || a.overtime === "never") &&
          a.report !== "last_day"
        );
      },
    },
    {
      id: "kontroler",
      shortName: "контролёр",
      titlePattern: "{role}-контролёр",
      group: "boss",
      priority: 80,
      text: "",
      imageSrc: "",
      triggersNote:
        "Руководитель + спокойные триггеры (КЭДО / высокая автоматизация / отчётность сразу).",
      match(a) {
        return isBoss(a.role) && calmScore(a) >= 2;
      },
    },
    {
      id: "kapitan_shtorma",
      shortName: "капитан шторма",
      titlePattern: "{role} — капитан шторма",
      group: "boss",
      priority: 90,
      text: "",
      imageSrc: "",
      triggersNote:
        "Руководитель + триггеры напряжения (ежедневно / почти всегда / последний день).",
      match(a) {
        return isBoss(a.role) && painScore(a) >= 2;
      },
    },
    {
      id: "novichok",
      shortName: "новичок в цифре",
      titlePattern: "{role} — новичок в цифре",
      group: "any",
      priority: 100,
      text: "",
      imageSrc: "",
      triggersNote: "Универсальный сценарий / запасной вариант, если остальные не сработали.",
      match() {
        return true;
      },
    },
  ],
};

function isHrLike(role) {
  return ["kadrovik", "buh_zp", "glb", "drugoe"].includes(role);
}

function isBoss(role) {
  return role === "ruk_otdela" || role === "ruk_kompanii";
}

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
  const pattern = scenario.titlePattern || "{role} — " + (scenario.shortName || scenario.id);
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
  isHrLike,
  isBoss,
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
    v1: "КЭДО: вариант 1",
    v2: "КЭДО: вариант 2",
    v3: "КЭДО: вариант 3",
    v4: "КЭДО: вариант 4",
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
