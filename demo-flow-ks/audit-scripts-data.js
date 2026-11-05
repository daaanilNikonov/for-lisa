/**
 * Аудит процессов «1С:Кабинет сотрудника» — ГК Форус
 * Без картинок: итог — список находок аудита + код сценария демо (С1…С6).
 */
window.AUDIT_SCRIPTS = {
  brand: "1С:Кабинет сотрудника",
  company: "ГК Форус",
  logoSrc: "assets/brand/logo-forus.png",
  productLine: "Аудит кадровых процессов",

  /**
   * Google Form — подставьте URL формы и entry.* id полей.
   * Как получить entry id: откройте форму → «Отправить» → «ссылка с предзаполнением»
   * → заполните пробные значения → в URL будут entry.123456=...
   *
   * Пока id пустые, заявка сохраняется локально + код сценария копируется,
   * а кнопка открывает форму (если указан только url).
   */
  googleForm: {
    url: "https://docs.google.com/forms/d/e/REPLACE_FORM_ID/viewform",
    fields: {
      name: "entry.REPLACE_NAME",
      phone: "entry.REPLACE_PHONE",
      inn: "entry.REPLACE_INN",
      company: "entry.REPLACE_COMPANY",
      scenario: "entry.REPLACE_SCENARIO",
      audit: "entry.REPLACE_AUDIT",
    },
  },

  questions: [
    {
      id: "slips",
      title: "Как сотрудники получают расчётные листки?",
      options: [
        { id: "paper", label: "На бумаге" },
        { id: "mail", label: "По почте / в мессенджере" },
        { id: "other_service", label: "Через другой сервис" },
        { id: "already_1c", label: "Уже через 1С" },
      ],
    },
    {
      id: "certs_get",
      title: "Как сотрудники получают справки?",
      options: [
        { id: "to_hr", label: "Обращаются к кадровику / бухгалтеру" },
        { id: "mail", label: "Запрашивают по почте / в мессенджере" },
        { id: "separate_service", label: "Через отдельный сервис" },
        { id: "self_system", label: "Самостоятельно в системе" },
      ],
    },
    {
      id: "certs_types",
      title: "Какие справки запрашивают чаще всего?",
      options: [
        { id: "income", label: "О доходах" },
        { id: "workplace", label: "С места работы" },
        { id: "seniority", label: "О стаже / периоде работы" },
        { id: "various", label: "Разные виды справок" },
        { id: "rare", label: "Запрашивают редко" },
      ],
    },
    {
      id: "vacation_ask",
      title: "Как часто сотрудники спрашивают про отпуск?",
      options: [
        { id: "rarely", label: "Почти не спрашивают" },
        { id: "days_left", label: "Спрашивают, сколько дней осталось" },
        { id: "how_to_apply", label: "Спрашивают, как оформить заявление" },
        {
          id: "both_often",
          label: "Часто спрашивают и про остаток, и про оформление",
        },
      ],
    },
    {
      id: "vacation_apply",
      title: "Как сотрудники подают заявление на отпуск?",
      options: [
        { id: "paper", label: "На бумаге" },
        { id: "mail", label: "По почте / в мессенджере" },
        { id: "separate", label: "Через отдельную систему" },
        { id: "already_1c", label: "Уже через 1С" },
      ],
    },
    {
      id: "vacation_approve",
      title: "Как согласовывается отпуск?",
      options: [
        { id: "via_hr", label: "Напрямую через кадровика / бухгалтера" },
        { id: "manager", label: "Нужно согласование руководителя" },
        { id: "multi", label: "Согласуют несколько лиц" },
        { id: "oral", label: "Устно" },
        { id: "automated", label: "Автоматизировано" },
      ],
    },
    {
      id: "vacation_record",
      title: "Как фиксируется согласование отпуска?",
      options: [
        { id: "none", label: "Никак" },
        { id: "mail", label: "Письмом / в мессенджере" },
        { id: "paper", label: "На бумаге" },
        { id: "separate", label: "В отдельной системе" },
        { id: "auto_1c", label: "Автоматически в 1С" },
      ],
    },
    {
      id: "trips",
      title: "Как оформляются командировки?",
      options: [
        { id: "never", label: "Не ездят в командировки" },
        { id: "rare", label: "Ездят нечасто" },
        { id: "with_report", label: "Ездят, предоставляют отчёт" },
        { id: "no_report", label: "Ездят, без отчётов" },
      ],
    },
    {
      id: "docs_get",
      title: "Как сотрудники получают кадровые документы?",
      options: [
        { id: "paper", label: "На бумаге" },
        { id: "mail", label: "По почте / в мессенджере" },
        { id: "other_service", label: "Через другой сервис" },
        { id: "electronic_1c", label: "Уже электронно в 1С" },
      ],
    },
    {
      id: "docs_find",
      title: "Как сотрудник находит свой документ?",
      options: [
        { id: "ask_hr", label: "Обращается к кадровику" },
        { id: "hr_searches", label: "Кадровик ищет и отправляет" },
        { id: "paper_archive", label: "Ищет в бумажном архиве" },
        { id: "other_service", label: "Ищет в другом сервисе" },
        { id: "self_system", label: "Самостоятельно в системе" },
      ],
    },
  ],

  /**
   * Правила аудита: если match(answers) → пункт попадает в итоговый список.
   * Порядок = приоритет показа.
   */
  auditRules: [
    {
      id: "vacation_not_recorded",
      match: (a) => a.vacation_record === "none",
      text: "Не согласуются отпуска — согласование никак не фиксируется",
    },
    {
      id: "vacation_oral",
      match: (a) => a.vacation_approve === "oral",
      text: "Отпуск согласуют устно — нет прозрачного маршрута",
    },
    {
      id: "vacation_manual_record",
      match: (a) =>
        a.vacation_record === "mail" || a.vacation_record === "paper",
      text: "Согласование отпуска фиксируется вручную (бумага / переписка)",
    },
    {
      id: "vacation_apply_offline",
      match: (a) =>
        a.vacation_apply === "paper" || a.vacation_apply === "mail",
      text: "Заявления на отпуск подают на бумаге или в мессенджере",
    },
    {
      id: "vacation_questions",
      match: (a) =>
        a.vacation_ask === "days_left" ||
        a.vacation_ask === "how_to_apply" ||
        a.vacation_ask === "both_often",
      text: "Сотрудники часто спрашивают про остаток отпуска и оформление",
    },
    {
      id: "slips_offline",
      match: (a) => a.slips && a.slips !== "already_1c",
      text: "Расчётные листки выдаются вне 1С",
    },
    {
      id: "certs_manual",
      match: (a) => a.certs_get === "to_hr" || a.certs_get === "mail",
      text: "Справки выдаются вручную через кадровика / переписку",
    },
    {
      id: "certs_separate",
      match: (a) => a.certs_get === "separate_service",
      text: "Справки — через отдельный сервис (контур не в 1С)",
    },
    {
      id: "certs_demand",
      match: (a) =>
        a.certs_types &&
        a.certs_types !== "rare" &&
        a.certs_get !== "self_system",
      text: "Справки запрашивают регулярно, самообслуживания нет",
    },
    {
      id: "trips_active",
      match: (a) => a.trips === "with_report" || a.trips === "no_report",
      text: "Есть регулярные командировки — нужен цифровой контур",
    },
    {
      id: "trips_no_report",
      match: (a) => a.trips === "no_report",
      text: "Командировки без отчётности",
    },
    {
      id: "docs_offline",
      match: (a) => a.docs_get && a.docs_get !== "electronic_1c",
      text: "Кадровые документы выдаются вне электронного контура 1С",
    },
    {
      id: "docs_search_manual",
      match: (a) =>
        a.docs_find === "ask_hr" ||
        a.docs_find === "hr_searches" ||
        a.docs_find === "paper_archive",
      text: "Документы ищут через кадровика или бумажный архив",
    },
  ],

  /**
   * Сценарии демо. Код С1…С6 менеджер использует на демонстрации.
   * pick: первый совпавший (по порядку массива).
   */
  scenarios: [
    {
      code: "С3",
      id: "vacation_chaos",
      title: "Хаос в отпусках",
      summary:
        "Согласование и фиксация отпуска слабые или отсутствуют. На демо делаем акцент на заявлении, маршруте согласования и остатке дней в Кабинете сотрудника.",
      demoHint: "Показать: заявление → согласование → приказ / остаток отпуска",
      pick: (a) =>
        a.vacation_record === "none" ||
        a.vacation_approve === "oral" ||
        (a.vacation_apply === "paper" && a.vacation_ask === "both_often"),
    },
    {
      code: "С4",
      id: "certs_bottleneck",
      title: "Узкое место — справки",
      summary:
        "Справки запрашивают часто и получают вручную. На демо — самообслуживание справок и разгрузка кадровика/бухгалтера.",
      demoHint: "Показать: запрос справки сотрудником без обращения к HR",
      pick: (a) =>
        (a.certs_get === "to_hr" || a.certs_get === "mail") &&
        a.certs_types !== "rare",
    },
    {
      code: "С5",
      id: "trips_focus",
      title: "Командировки в фокусе",
      summary:
        "Сотрудники ездят в командировки. В сценарии демо добавляем оформление командировок и связанные документы.",
      demoHint: "Показать: командировка + связанные кадровые документы",
      pick: (a) => a.trips === "with_report" || a.trips === "no_report",
    },
    {
      code: "С6",
      id: "docs_archive",
      title: "Архив и поиск документов",
      summary:
        "Документы ищут через кадровика или бумажный архив. Демо строится вокруг электронного доступа сотрудника к своим документам.",
      demoHint: "Показать: личный кабинет документов и поиск без HR",
      pick: (a) =>
        a.docs_find === "paper_archive" ||
        a.docs_find === "ask_hr" ||
        a.docs_find === "hr_searches",
    },
    {
      code: "С1",
      id: "digital_mature",
      title: "Уже в цифре — расширение контура",
      summary:
        "Базовые процессы уже в 1С или близко к этому. Демо: углубление, непроявленные зоны и единый контур вместо разрозненных сервисов.",
      demoHint: "Показать: единый кабинет вместо отдельных каналов",
      pick: (a) => {
        let score = 0;
        if (a.slips === "already_1c") score += 1;
        if (a.certs_get === "self_system") score += 1;
        if (a.vacation_apply === "already_1c") score += 1;
        if (a.vacation_record === "auto_1c" || a.vacation_approve === "automated")
          score += 1;
        if (a.docs_get === "electronic_1c") score += 1;
        if (a.docs_find === "self_system") score += 1;
        return score >= 4;
      },
    },
    {
      code: "С2",
      id: "classic_mix",
      title: "Классический смешанный контур",
      summary:
        "Бумага, почта и мессенджеры вперемешку. Базовый сценарий демо Кабинета сотрудника: листки, справки, отпуск, документы в одном окне.",
      demoHint: "Показать полный путь сотрудника: листки → справки → отпуск → документы",
      pick: () => true,
    },
  ],
};

window.AUDIT_MATCH = {
  buildAudit(answers) {
    const findings = [];
    for (const rule of window.AUDIT_SCRIPTS.auditRules) {
      try {
        if (rule.match(answers)) {
          findings.push({ id: rule.id, text: rule.text });
        }
      } catch (err) {
        console.warn(rule.id, err);
      }
    }
    return findings;
  },

  pickScenario(answers) {
    for (const sc of window.AUDIT_SCRIPTS.scenarios) {
      try {
        if (sc.pick(answers)) return sc;
      } catch (err) {
        console.warn(sc.code, err);
      }
    }
    return window.AUDIT_SCRIPTS.scenarios[window.AUDIT_SCRIPTS.scenarios.length - 1];
  },

  maturityScore(answers) {
    let score = 0;
    if (answers.slips === "already_1c") score += 1;
    if (answers.certs_get === "self_system") score += 1;
    if (answers.vacation_apply === "already_1c") score += 1;
    if (answers.vacation_record === "auto_1c") score += 1;
    if (answers.vacation_approve === "automated") score += 1;
    if (answers.docs_get === "electronic_1c") score += 1;
    if (answers.docs_find === "self_system") score += 1;
    return score;
  },
};
