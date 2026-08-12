/**
 * Тест «1С:Кабинет сотрудника» — ГК Форус
 *
 * 10 сценариев выбираются по ТРИГГЕРАМ (не по численности).
 * Численность сохраняется в профиле, но на канал не влияет.
 *
 * Триггеры «напряжения» (к авральным сценариям):
 *   — справки ежедневно
 *   — задерживаются почти всегда
 *   — отчётность в последний день
 *
 * Триггеры «спокойствия» (к «детокс»/цифровым сценариям):
 *   — автоматизация >80%
 *   — есть КЭДО
 *   — отчётность сразу
 *   — справки очень редко / не обращаются
 */
window.QUIZ_SCRIPTS = {
  brand: "1С:Кабинет сотрудника",
  company: "ГК Форус",
  logoSrc: "assets/brand/logo-forus.png",

  /** Временные картинки (рандомный пул, пока нет финальных 10) */
  tempImages: [
    "assets/results/temp-1.jpeg",
    "assets/results/temp-2.jpeg",
    "assets/results/temp-3.jpeg",
  ],

  questions: [
    {
      id: "role",
      number: 1,
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
      number: 2,
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
      number: 3,
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
      number: 4,
      title: "Как часто ко мне приходят сотрудники за справками:",
      options: [
        { id: "never", label: "не обращаются" },
        { id: "daily", label: "обращаются ежедневно" },
        { id: "rare", label: "очень редко" },
      ],
    },
    {
      id: "auto",
      number: 5,
      title: "Уровень автоматизации моей работы:",
      options: [
        { id: "low", label: "до 40% автоматизировано" },
        { id: "mid", label: "40–80% автоматизировано" },
        { id: "high", label: "более 80% автоматизировано" },
      ],
    },
    {
      id: "overtime",
      number: 6,
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
      number: 7,
      title: "Когда сдаете кадровую отчетность:",
      options: [
        { id: "now", label: "сразу" },
        { id: "month", label: "в течение месяца" },
        { id: "last_day", label: "в последний день" },
      ],
    },
  ],

  /**
   * 10 основных сценариев.
   * text пока пустой; imageSrc — временно из пула tempImages (см. app.js assignTempImages).
   * matches(answers) — правила триггеров; первый подходящий по порядку приоритета побеждает.
   */
  scenarios: [
    {
      id: "detoks",
      name: "Бухгалтер/кадровик на детоксе",
      shortName: "На детоксе",
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
      name: "Бухгалтер/кадровик — любимчик",
      shortName: "Любимчик",
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
      name: "Бухгалтер/кадровик — сорвиголова",
      shortName: "Сорвиголова",
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
      name: "Книжный червь",
      shortName: "Книжный червь",
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
      name: "Пожарный",
      shortName: "Пожарный",
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
      name: "Бумажный маг",
      shortName: "Бумажный маг",
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
      name: "Тихая гавань",
      shortName: "Тихая гавань",
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
      name: "Руководитель-контролёр",
      shortName: "Контролёр",
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
      name: "Капитан шторма",
      shortName: "Капитан шторма",
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
      name: "Новичок в цифре",
      shortName: "Новичок в цифре",
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

window.QUIZ_MATCH = {
  isHrLike,
  isBoss,
  painScore,
  calmScore,
  /** Вернуть сценарий по ответам (минимальный priority среди совпавших) */
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
