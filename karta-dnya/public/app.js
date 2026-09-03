(() => {
  "use strict";

  const state = {
    today: null,
    tomorrow: null,
    managers: [],
    forms: [],
    boards: {},
    kpiDefs: [],
    transfers: [],
    stickerPack: [],
    analytics: [],
    dashboard: null,
    tomorrowPreview: [],
    monthlyKpis: {},
    kpiCalendar: null,
    selectedId: "all",
    phase: "morning",
    planDate: null,
    weekDrafts: {},
    tomorrowCollapsed: false,
    weekView: "all",
    draftTasks: [],
    monthlyDraft: [],
    stickerSaveTimer: null,
  };

  const COLORS = ["cyan", "amber", "mint", "rose", "violet", "pink"];
  const LINE_COLORS = ["#26a6e0", "#fecf68", "#7dcea0", "#f0a3b0", "#a9b4f5", "#ff8bd2"];
  const WEEKDAY_SHORT = ["вс", "пн", "вт", "ср", "чт", "пт", "сб"];

  const el = {
    todayLabel: document.getElementById("todayLabel"),
    managerTabs: document.getElementById("managerTabs"),
    overviewView: document.getElementById("overviewView"),
    singleView: document.getElementById("singleView"),
    checklistTitle: document.getElementById("checklistTitle"),
    checklistSub: document.getElementById("checklistSub"),
    boardTitle: document.getElementById("boardTitle"),
    boardSub: document.getElementById("boardSub"),
    weekDays: document.getElementById("weekDays"),
    saveWeekBtn: document.getElementById("saveWeekBtn"),
    tasksEditor: document.getElementById("tasksEditor"),
    stickerBoard: document.getElementById("stickerBoard"),
    stickerPack: document.getElementById("stickerPack"),
    archiveList: document.getElementById("archiveList"),
    saveFormBtn: document.getElementById("saveFormBtn"),
    formHint: document.getElementById("formHint"),
    tabMorning: document.getElementById("tabMorning"),
    tabEvening: document.getElementById("tabEvening"),
    addTaskBtn: document.getElementById("addTaskBtn"),
    addTaskMenu: document.getElementById("addTaskMenu"),
    weekChecklist: document.getElementById("weekChecklist"),
    openWeekTaskPanelBtn: document.getElementById("openWeekTaskPanelBtn"),
    weekTaskModal: document.getElementById("weekTaskModal"),
    weekTaskForm: document.getElementById("weekTaskForm"),
    weekTaskText: document.getElementById("weekTaskText"),
    weekTaskTarget: document.getElementById("weekTaskTarget"),
    weekTaskUnit: document.getElementById("weekTaskUnit"),
    weekTaskMetrics: document.getElementById("weekTaskMetrics"),
    weekTaskDayGrid: document.getElementById("weekTaskDayGrid"),
    weekTaskSelectAll: document.getElementById("weekTaskSelectAll"),
    weekTaskSelectOne: document.getElementById("weekTaskSelectOne"),
    weekTaskMoveFrom: document.getElementById("weekTaskMoveFrom"),
    dayPlanTitle: document.getElementById("dayPlanTitle"),
    dayPlanSub: document.getElementById("dayPlanSub"),
    kpiSetupBanner: document.getElementById("kpiSetupBanner"),
    monthlyKpiBox: document.getElementById("monthlyKpiBox"),
    monthlyKpiSub: document.getElementById("monthlyKpiSub"),
    monthlyKpiEditor: document.getElementById("monthlyKpiEditor"),
    addMonthlyKpiBtn: document.getElementById("addMonthlyKpiBtn"),
    saveMonthlyKpiBtn: document.getElementById("saveMonthlyKpiBtn"),
    reportWeekStart: document.getElementById("reportWeekStart"),
    buildReportBtn: document.getElementById("buildReportBtn"),
    addManagerBtn: document.getElementById("addManagerBtn"),
    renameManagerBtn: document.getElementById("renameManagerBtn"),
    addSticker: document.getElementById("addSticker"),
    managerStats: document.getElementById("managerStats"),
    kpiTable: document.querySelector("#kpiTable tbody"),
    chartLegend: document.getElementById("chartLegend"),
    convChart: document.getElementById("convChart"),
    analyticsCards: document.getElementById("analyticsCards"),
    transfersList: document.getElementById("transfersList"),
    tomorrowWindow: document.getElementById("tomorrowWindow"),
    tomorrowTitle: document.getElementById("tomorrowTitle"),
    tomorrowDate: document.getElementById("tomorrowDate"),
    tomorrowList: document.getElementById("tomorrowList"),
    tomorrowClose: document.getElementById("tomorrowClose"),
    tomorrowOpenDay: document.getElementById("tomorrowOpenDay"),
    tomorrowFab: document.getElementById("tomorrowFab"),
    tomorrowFabCount: document.getElementById("tomorrowFabCount"),
    promptModal: document.getElementById("promptModal"),
    promptForm: document.getElementById("promptForm"),
    promptTitle: document.getElementById("promptTitle"),
    promptLabel: document.getElementById("promptLabel"),
    promptInput: document.getElementById("promptInput"),
    toast: document.getElementById("toast"),
  };

  async function api(path, options = {}) {
    const res = await fetch(path, {
      headers: { "Content-Type": "application/json", ...(options.headers || {}) },
      ...options,
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || "Ошибка запроса");
    return data;
  }

  function toast(message) {
    el.toast.textContent = message;
    el.toast.style.zIndex = "2147483647";
    el.toast.classList.add("is-on");
    clearTimeout(toast._t);
    toast._t = setTimeout(() => el.toast.classList.remove("is-on"), 3400);
  }

  function formatDateRu(iso) {
    if (!iso) return "";
    const [y, m, d] = iso.split("-").map(Number);
    return new Date(y, m - 1, d).toLocaleDateString("ru-RU", {
      day: "numeric",
      month: "long",
      year: "numeric",
    });
  }

  function addDaysISO(iso, days) {
    const [y, m, d] = iso.split("-").map(Number);
    const dt = new Date(y, m - 1, d);
    dt.setDate(dt.getDate() + days);
    return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, "0")}-${String(dt.getDate()).padStart(2, "0")}`;
  }

  function todaysForm(managerId, date = state.planDate || state.today) {
    return state.forms.find((f) => f.managerId === managerId && f.date === date);
  }

  function weekdayShort(iso) {
    const [y, m, d] = iso.split("-").map(Number);
    return WEEKDAY_SHORT[new Date(y, m - 1, d).getDay()];
  }

  function weekDateList() {
    return workdaysFromToday(5);
  }

  /** Next N workdays (Mon–Fri), starting from today (or next Monday if weekend). */
  function workdaysFromToday(count = 5) {
    const out = [];
    let cur = state.today;
    // if weekend, jump to next Monday
    const [y0, m0, d0] = cur.split("-").map(Number);
    let dow = new Date(y0, m0 - 1, d0).getDay();
    if (dow === 0) cur = addDaysISO(cur, 1);
    else if (dow === 6) cur = addDaysISO(cur, 2);
    let guard = 0;
    while (out.length < count && guard < 20) {
      const [y, m, d] = cur.split("-").map(Number);
      const day = new Date(y, m - 1, d).getDay();
      if (day >= 1 && day <= 5) out.push(cur);
      cur = addDaysISO(cur, 1);
      guard += 1;
    }
    return out;
  }


  function draftKey(managerId, date) {
    return `${managerId}:${date}`;
  }

  function stashCurrentDraft() {
    if (!state.planDate || !state.selectedId || state.selectedId === "all") return;
    state.weekDrafts[draftKey(state.selectedId, state.planDate)] = state.draftTasks.map((t) => ({
      ...t,
    }));
  }

  function tasksForDate(managerId, date) {
    const key = draftKey(managerId, date);
    if (state.weekDrafts[key]) return state.weekDrafts[key].map((t) => ({ ...t }));
    const form = state.forms.find((f) => f.managerId === managerId && f.date === date);
    if (form?.tasks?.length) {
      return form.tasks.map((t) => ({
        ...t,
        transferDate: t.transferDate || addDaysISO(date, 1),
      }));
    }
    return blankTaskRow(date);
  }

  function statusFor(managerId) {
    const form = state.forms.find((f) => f.managerId === managerId && f.date === state.today);
    if (!form) return { key: "idle", label: "Ещё не начато" };
    if (form.status === "completed") return { key: "done", label: "В архиве" };
    if (form.status === "morning") return { key: "morning", label: "Чеклист задан" };
    return { key: "idle", label: "Ещё не начато" };
  }

  function progressPct(task) {
    const target = Math.max(1, Number(task.target) || 1);
    const done = Math.max(0, Number(task.doneCount) || 0);
    return Math.round((done / target) * 100);
  }

  function askPrompt({ title, label, placeholder = "", initial = "" }) {
    return new Promise((resolve) => {
      el.promptTitle.textContent = title;
      el.promptLabel.textContent = label;
      el.promptInput.value = initial;
      el.promptInput.placeholder = placeholder;
      const onClose = () => {
        el.promptModal.removeEventListener("close", onClose);
        resolve(el.promptModal.returnValue === "ok" ? el.promptInput.value.trim() : null);
      };
      el.promptModal.addEventListener("close", onClose);
      el.promptModal.showModal();
      requestAnimationFrame(() => el.promptInput.focus());
    });
  }

  async function loadState() {
    const data = await api("/api/state");
    state.today = data.today;
    state.tomorrow = data.tomorrow || addDaysISO(data.today, 1);
    if (!state.planDate) state.planDate = state.today;
    state.managers = data.managers || [];
    state.forms = data.forms || [];
    state.boards = data.boards || {};
    state.kpiDefs = data.kpiDefs || [];
    state.transfers = data.transfers || [];
    state.stickerPack = data.stickerPack || [];
    state.analytics = data.analytics || [];
    state.dashboard = data.dashboard || null;
    state.tomorrowPreview = data.tomorrowPreview || [];
    el.todayLabel.textContent = `Сегодня · ${formatDateRu(state.today)}`;
    if (state.selectedId !== "all" && !state.managers.some((m) => m.id === state.selectedId)) {
      state.selectedId = "all";
    }
    renderAll();
  }

  function renderAll() {
    renderTabs();
    const single = state.selectedId !== "all";
    el.renameManagerBtn.classList.toggle("is-hidden", !single);
    if (!single) {
      el.overviewView.classList.remove("is-hidden");
      el.singleView.classList.add("is-hidden");
      renderOverview();
      hideTomorrowWindow();
    } else {
      el.overviewView.classList.add("is-hidden");
      el.singleView.classList.remove("is-hidden");
      loadManagerWorkspace(state.selectedId);
    }
    renderDashboard();
    renderArchive();
  }

  function renderTabs() {
    el.managerTabs.innerHTML = "";
    const allBtn = document.createElement("button");
    allBtn.type = "button";
    allBtn.className = `manager-tab${state.selectedId === "all" ? " is-active" : ""}`;
    allBtn.textContent = "Все";
    allBtn.addEventListener("click", () => {
      state.selectedId = "all";
      renderAll();
    });
    el.managerTabs.appendChild(allBtn);

    state.managers.forEach((manager) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = `manager-tab${state.selectedId === manager.id ? " is-active" : ""}`;
      const status = statusFor(manager.id);
      btn.innerHTML = `<span></span><small></small>`;
      btn.querySelector("span").textContent = manager.name;
      btn.querySelector("small").textContent = status.label;
      btn.title = "Клик — открыть · двойной клик — переименовать";
      btn.addEventListener("click", () => {
        state.selectedId = manager.id;
        state.planDate = state.today;
        state.weekDrafts = {};
        state.monthlyDraft = [];
        renderAll();
      });
      btn.addEventListener("dblclick", (e) => {
        e.preventDefault();
        renameManager(manager);
      });
      el.managerTabs.appendChild(btn);
    });
  }

  function renderOverview() {
    el.overviewView.innerHTML = "";
    state.managers.forEach((manager) => {
      const form = state.forms.find((f) => f.managerId === manager.id && f.date === state.today);
      const status = statusFor(manager.id);
      const stats = state.analytics.find((a) => a.managerId === manager.id);
      const tomorrowCount =
        (state.forms.find((f) => f.managerId === manager.id && f.date === state.tomorrow)?.tasks || [])
          .length ||
        state.tomorrowPreview.find((t) => t.managerId === manager.id)?.taskCount ||
        0;
      const card = document.createElement("article");
      card.className = "overview-card";
      card.innerHTML = `
        <header class="overview-card-head">
          <div>
            <h3></h3>
            <span class="manager-status"></span>
          </div>
          <button type="button" class="btn btn-ghost btn-sm" data-open>Открыть</button>
        </header>
        <p class="tomorrow-teaser"></p>
        <div class="overview-grid">
          <div>
            <p class="overview-label">Чеклист</p>
            <ul class="overview-tasks"></ul>
          </div>
          <div>
            <p class="overview-label">Аналитика</p>
            <div class="mini-stats"></div>
          </div>
        </div>`;
      card.querySelector("h3").textContent = manager.name;
      const st = card.querySelector(".manager-status");
      st.textContent = status.label;
      st.classList.add(status.key === "done" ? "is-done" : status.key === "morning" ? "is-morning" : "");
      card.querySelector(".tomorrow-teaser").textContent = tomorrowCount
        ? `Завтра ждёт: ${tomorrowCount} ${tomorrowCount === 1 ? "задача" : "задач"}`
        : "Завтра: план пока пуст";
      const list = card.querySelector(".overview-tasks");
      const tasks = form?.tasks || [];
      if (!tasks.length) {
        list.innerHTML = `<li class="muted">Пока пусто</li>`;
      } else {
        tasks.slice(0, 5).forEach((t) => {
          const li = document.createElement("li");
          li.textContent = `${progressPct(t)}% · ${t.text}${t.carriedTo ? " → перенос" : ""}`;
          list.appendChild(li);
        });
      }
      const mini = card.querySelector(".mini-stats");
      mini.innerHTML = `
        <div><b>${stats?.tasksFullyDone || 0}</b><span>сделано</span></div>
        <div><b>${stats?.tasksCarried || 0}</b><span>перенесено</span></div>
        <div><b>${stats?.conversion || 0}%</b><span>конверсия</span></div>`;
      card.querySelector("[data-open]").addEventListener("click", () => {
        state.selectedId = manager.id;
        state.planDate = state.today;
        renderAll();
      });
      el.overviewView.appendChild(card);
    });
  }

  function loadManagerWorkspace(managerId) {
    const manager = state.managers.find((m) => m.id === managerId);
    if (!manager) return;
    if (!state.planDate) state.planDate = state.today;
    const form = todaysForm(managerId, state.planDate);
    const isToday = state.planDate === state.today;
    el.checklistTitle.textContent = `Чеклист · ${manager.name}`;
    el.checklistSub.textContent = `${formatDateRu(state.planDate)} · «${manager.name} ${state.planDate}»`;
    el.boardTitle.textContent = "Основные приоритеты в работе";
    el.boardSub.textContent = `Доска · ${manager.name}`;

    state.draftTasks = tasksForDate(managerId, state.planDate);

    if (form?.status === "completed") {
      state.phase = "evening";
    } else if (isToday && form?.status === "morning") {
      state.phase = state.phase === "evening" ? "evening" : "morning";
    } else {
      state.phase = "morning";
    }

    setPhase(state.phase);
    renderWeekDays(managerId);
    renderTasksEditor();
    renderManagerStats(managerId);
    renderStickerPack();
    renderStickers(managerId);
    renderTomorrowWindow(managerId);
    renderMonthlyKpiEditor(managerId);
    renderKpiSetupBanner();
    if (el.reportWeekStart && !el.reportWeekStart.value && state.today) {
      const [y, m, d] = state.today.split("-").map(Number);
      const dt = new Date(y, m - 1, d);
      const day = dt.getDay();
      const diff = day === 0 ? -6 : 1 - day;
      dt.setDate(dt.getDate() + diff);
      el.reportWeekStart.value = `${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,"0")}-${String(dt.getDate()).padStart(2,"0")}`;
    }
  }

  function renderKpiSetupBanner() {
    if (!el.kpiSetupBanner) return;
    const cal = state.kpiCalendar;
    if (!cal) {
      el.kpiSetupBanner.classList.add("is-hidden");
      return;
    }
    const setup = cal.setupDay;
    const isToday = state.today === setup;
    el.kpiSetupBanner.classList.toggle("is-hidden", false);
    el.kpiSetupBanner.classList.toggle("is-hot", isToday);
    el.kpiSetupBanner.textContent = isToday
      ? `Сегодня день установки KPI на месяц (${setup}). Пропишите обязательные KPI ниже — они будут каждый будний день.`
      : `День установки KPI в этом месяце: ${setup}. До этого действуют KPI прошлого месяца (если заданы).`;
  }

  function renderMonthlyKpiEditor(managerId) {
    if (!el.monthlyKpiEditor) return;
    const month = monthKey(state.today);
    const [y, m] = month.split("-").map(Number);
    const setup = kpiSetupDayForMonth(y, m);
    const saved = state.monthlyKpis?.[managerId]?.[month];
    if (!state.monthlyDraft.length) {
      state.monthlyDraft = (saved?.items || state.kpiDefs || []).map((k) => ({
        id: k.id || `mkpi-${k.name}`,
        name: k.name || k.text || "",
        target: k.target || k.defaultTarget || 1,
        unit: k.unit || "",
      }));
      if (!state.monthlyDraft.length) {
        state.monthlyDraft = [{ id: `mkpi-${Date.now()}`, name: "", target: 1, unit: "" }];
      }
    }
    if (el.monthlyKpiSub) {
      el.monthlyKpiSub.textContent = `Месяц ${month} · установка ${setup}${saved ? " · уже сохранено" : " · ещё не заданы"}`;
    }
    el.monthlyKpiEditor.innerHTML = "";
    state.monthlyDraft.forEach((item, index) => {
      const row = document.createElement("div");
      row.className = "monthly-kpi-row";
      row.innerHTML = `
        <input type="text" class="mk-name" placeholder="Название KPI" />
        <input type="number" class="mk-target" min="1" step="1" title="План" />
        <input type="text" class="mk-unit" placeholder="ед." />
        <button type="button" class="task-remove" title="Удалить">×</button>`;
      const name = row.querySelector(".mk-name");
      const target = row.querySelector(".mk-target");
      const unit = row.querySelector(".mk-unit");
      name.value = item.name || "";
      target.value = item.target || 1;
      unit.value = item.unit || "";
      const sync = () => {
        item.name = name.value;
        item.target = Math.max(1, Number(target.value) || 1);
        item.unit = unit.value.trim();
      };
      [name, target, unit].forEach((n) => n.addEventListener("input", sync));
      row.querySelector(".task-remove").addEventListener("click", () => {
        state.monthlyDraft.splice(index, 1);
        if (!state.monthlyDraft.length) {
          state.monthlyDraft = [{ id: `mkpi-${Date.now()}`, name: "", target: 1, unit: "" }];
        }
        renderMonthlyKpiEditor(managerId);
      });
      el.monthlyKpiEditor.appendChild(row);
    });
    if (!state.monthlyDraft.length) {
      const empty = document.createElement("p");
      empty.className = "panel-sub";
      empty.textContent = "Пока пусто — нажмите «+ Добавить KPI».";
      el.monthlyKpiEditor.appendChild(empty);
    }
  }

  function addMonthlyKpiRow() {
    if (!state.selectedId || state.selectedId === "all") {
      toast("Сначала откройте сотрудника");
      return;
    }
    state.monthlyDraft.push({ id: `mkpi-${Date.now()}`, name: "", target: 1, unit: "" });
    renderMonthlyKpiEditor(state.selectedId);
    const inputs = el.monthlyKpiEditor.querySelectorAll(".mk-name");
    const last = inputs[inputs.length - 1];
    if (last) last.focus();
  }

  async function saveMonthlyKpi() {
    const managerId = state.selectedId;
    if (!managerId || managerId === "all") return;
    const month = monthKey(state.today);
    const items = state.monthlyDraft
      .map((k) => ({
        id: k.id,
        name: String(k.name || "").trim(),
        target: Math.max(1, Number(k.target) || 1),
        unit: String(k.unit || "").trim(),
      }))
      .filter((k) => k.name);
    if (!items.length) {
      toast("Укажите хотя бы один KPI");
      return;
    }
    try {
      const data = await api(`/api/managers/${managerId}/monthly-kpi`, {
        method: "PUT",
        body: JSON.stringify({ month, items }),
      });
      if (!state.monthlyKpis[managerId]) state.monthlyKpis[managerId] = {};
      state.monthlyKpis[managerId] = data.monthlyKpis;
      toast(data.message || "KPI месяца сохранены");
      renderMonthlyKpiEditor(managerId);
    } catch (err) {
      toast(err.message);
    }
  }


  function blankTaskRow(date = state.planDate || state.today) {
    return [
      {
        id: `local-${Date.now()}`,
        text: "",
        target: 1,
        doneCount: 0,
        unit: "",
        mandatory: false,
        kind: "numeric",
        done: false,
        transferDate: addDaysISO(date, 1),
      },
    ];
  }

  function setPhase(phase) {
    const isToday = state.planDate === state.today;
    if (phase === "evening" && !isToday) {
      phase = "morning";
      toast("Вечернее закрытие доступно только для сегодняшнего дня");
    }
    state.phase = phase;
    el.tabMorning.classList.toggle("is-active", phase === "morning");
    el.tabEvening.classList.toggle("is-active", phase === "evening");
    el.tabEvening.disabled = !isToday;
    el.tabEvening.title = isToday ? "" : "Сначала дождитесь этого дня";

    const form = todaysForm(state.selectedId, state.planDate);
    const completed = form?.status === "completed";
    const hasMorning = Boolean(form && (form.status === "morning" || form.status === "completed"));

    if (phase === "morning") {
      el.saveFormBtn.textContent = completed
        ? "Уже в архиве"
        : isToday
          ? "Сохранить утро"
          : "Сохранить план дня";
      el.saveFormBtn.disabled = completed;
      el.addTaskBtn.hidden = completed;
            el.formHint.textContent = completed
        ? "День уже закрыт. Можно смотреть прогресс и переносы."
        : isToday
          ? "Утро: план и обязательные KPI. Можно задать цель числом (например, 2 демо)."
          : "Планируете вперёд: задачи будут ждать менеджера в выбранный день. Вечернее закрытие — только в сам день.";
    } else {
      el.saveFormBtn.textContent = completed ? "Уже в архиве" : "Закрыть день · перенести остатки";
      el.saveFormBtn.disabled = completed || !hasMorning;
      el.addTaskBtn.hidden = true;
            el.formHint.textContent = !hasMorning
        ? "Сначала сохраните утренний чеклист."
        : completed
          ? "День в архиве."
          : "Вечер: факт (сколько сделано). Остаток автоматически уйдёт на выбранную дату переноса.";
    }
  }

  function renderWeekDays(managerId) {
    el.weekDays.innerHTML = "";
    workdaysFromToday(5).forEach((date) => {
      const form = state.forms.find((f) => f.managerId === managerId && f.date === date);
      const draft = state.weekDrafts[draftKey(managerId, date)];
      const count = draft
        ? draft.filter((t) => String(t.text || "").trim()).length
        : (form?.tasks || []).length;
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "week-day";
      if (date === state.planDate) btn.classList.add("is-active");
      if (date === state.today) btn.classList.add("is-today");
      if (date === state.tomorrow) btn.classList.add("is-tomorrow");
      const [yy, mm] = date.split("-").map(Number);
      if (date === kpiSetupDayForMonth(yy, mm)) btn.classList.add("is-kpi-setup");
      if (form?.status === "completed") btn.classList.add("is-done");
      else if (count > 0) btn.classList.add("has-tasks");
      const planned = date > state.today;
      btn.innerHTML = `
        <span class="wd-name">${weekdayShort(date)}</span>
        <span class="wd-num">${date.slice(8)}</span>
        <span class="wd-count">${count ? `${count} зад.` : planned ? "план" : "пусто"}</span>`;
      btn.title = planned
        ? "Заранее: задачи сохранятся и откроются в этот день"
        : "Сегодняшний рабочий день";
      btn.addEventListener("click", () => {
        selectPlanDate(date);
        openWeekTaskPanel({ preselect: [date] });
      });
      el.weekDays.appendChild(btn);
    });
    renderWeekChecklist(managerId);
    updateDayPlanLabels();
  }

  function updateDayPlanLabels() {
    const date = state.planDate || state.today;
    if (el.dayPlanTitle) {
      el.dayPlanTitle.textContent =
        date === state.today
          ? `План дня · сегодня (${weekdayShort(date)})`
          : `Черновик на ${weekdayShort(date)}, ${formatDateRu(date)}`;
    }
    if (el.dayPlanSub) {
      el.dayPlanSub.textContent =
        date === state.today
          ? "Можно закрывать вечер и переносить остатки"
          : "Задачи уже сохраняются заранее; утро/вечер этого дня откроются только в сам день";
    }
  }

  function renderWeekChecklist(managerId) {
    if (!el.weekChecklist) return;
    el.weekChecklist.innerHTML = "";
    const days = workdaysFromToday(5);
    let shown = 0;
    days.forEach((date) => {
      const tasks = tasksForDate(managerId, date).filter((t) => String(t.text || "").trim());
      if (state.weekView === "filled" && !tasks.length) return;
      shown += 1;
      const block = document.createElement("article");
      block.className = `week-check-day${date === state.planDate ? " is-active" : ""}`;
      const form = state.forms.find((f) => f.managerId === managerId && f.date === date);
      const status =
        form?.status === "completed" ? "закрыт" : date > state.today ? "заранее" : form ? "в работе" : "пусто";
      block.innerHTML = `
        <header>
          <button type="button" class="week-check-open"></button>
          <span class="week-check-status"></span>
        </header>
        <ul></ul>
        <div class="week-check-actions">
          <button type="button" class="btn btn-ghost btn-sm" data-act="add">+ Задача</button>
          <button type="button" class="btn btn-ghost btn-sm" data-act="open">Открыть день</button>
        </div>`;
      block.querySelector(".week-check-open").textContent =
        `${weekdayShort(date).toUpperCase()} ${date.slice(8)}.${date.slice(5, 7)} · ${tasks.length} зад.`;
      block.querySelector(".week-check-status").textContent = status;
      const ul = block.querySelector("ul");
      if (!tasks.length) {
        ul.innerHTML = `<li class="muted">Нет задач — нажмите «+ Задача»</li>`;
      } else {
        tasks.forEach((t) => {
          const li = document.createElement("li");
          const kind = t.kind === "check" ? "да/нет" : t.kind === "kpi" ? "KPI" : (t.unit || "шт.");
          li.textContent = t.kind === "check"
            ? `${t.text} · ${kind}`
            : `${t.text} · ${kind} ${t.target || 1}`;
          ul.appendChild(li);
        });
      }
      block.querySelector(".week-check-open").addEventListener("click", () => selectPlanDate(date));
      block.querySelector('[data-act="open"]').addEventListener("click", () => selectPlanDate(date));
      block.querySelector('[data-act="add"]').addEventListener("click", () => {
        selectPlanDate(date);
        openWeekTaskPanel({ preselect: [date] });
      });
      el.weekChecklist.appendChild(block);
    });
    if (!shown) {
      el.weekChecklist.innerHTML =
        `<div class="archive-empty">Пока нет задач на будни. Нажмите «+ Добавить задачи».</div>`;
    }
  }

  function selectPlanDate(date) {
    if (!state.selectedId || state.selectedId === "all") return;
    stashCurrentDraft();
    state.planDate = date;
    state.draftTasks = tasksForDate(state.selectedId, date);
    if (date !== state.today) state.phase = "morning";
    else if (todaysForm(state.selectedId, date)?.status === "completed") state.phase = "evening";
    setPhase(state.phase);
    renderWeekDays(state.selectedId);
    renderTasksEditor();
    updateDayPlanLabels();
    const manager = state.managers.find((m) => m.id === state.selectedId);
    if (manager) {
      el.checklistSub.textContent = `${formatDateRu(date)} · «${manager.name} ${date}»`;
    }
  }

  function fillWeekTaskDayGrid(preselect = []) {
    if (!el.weekTaskDayGrid) return;
    el.weekTaskDayGrid.innerHTML = "";
    const selected = new Set(preselect.length ? preselect : [state.planDate || state.today]);
    workdaysFromToday(5).forEach((date) => {
      const label = document.createElement("label");
      label.className = "week-day-check";
      label.innerHTML = `<input type="checkbox" value="${date}" /><span>${weekdayShort(date)} ${date.slice(8)}.${date.slice(5, 7)}</span>`;
      const input = label.querySelector("input");
      input.checked = selected.has(date);
      el.weekTaskDayGrid.appendChild(label);
    });
  }

  function fillWeekTaskMoveFrom() {
    if (!el.weekTaskMoveFrom) return;
    const managerId = state.selectedId;
    el.weekTaskMoveFrom.innerHTML = `<option value="">— не переносить, создать новую —</option>`;
    workdaysFromToday(5).forEach((date) => {
      tasksForDate(managerId, date)
        .filter((t) => String(t.text || "").trim())
        .forEach((t) => {
          const opt = document.createElement("option");
          opt.value = `${date}::${t.id}`;
          opt.textContent = `${weekdayShort(date)} ${date.slice(8)} · ${t.text}`;
          el.weekTaskMoveFrom.appendChild(opt);
        });
    });
  }

  function openWeekTaskPanel({ preselect = [] } = {}) {
    if (!el.weekTaskModal) return;
    if (!state.selectedId || state.selectedId === "all") {
      toast("Сначала откройте сотрудника");
      return;
    }
    el.weekTaskText.value = "";
    el.weekTaskTarget.value = "1";
    el.weekTaskUnit.value = "";
    const kindRadio = el.weekTaskForm.querySelector('input[name="weekTaskKind"][value="numeric"]');
    if (kindRadio) kindRadio.checked = true;
    syncWeekTaskMetricsVisibility();
    fillWeekTaskDayGrid(preselect);
    fillWeekTaskMoveFrom();
    el.weekTaskModal.showModal();
    requestAnimationFrame(() => el.weekTaskText.focus());
  }

  function syncWeekTaskMetricsVisibility() {
    const kind = el.weekTaskForm?.querySelector('input[name="weekTaskKind"]:checked')?.value || "numeric";
    if (el.weekTaskMetrics) {
      el.weekTaskMetrics.classList.toggle("is-hidden", kind === "check");
    }
  }

  function selectedWeekTaskDates() {
    return [...(el.weekTaskDayGrid?.querySelectorAll("input:checked") || [])].map((n) => n.value);
  }

  async function submitWeekTaskFromPanel() {
    const managerId = state.selectedId;
    const text = String(el.weekTaskText.value || "").trim();
    if (!text) {
      toast("Введите текст задачи");
      return false;
    }
    const dates = selectedWeekTaskDates();
    if (!dates.length) {
      toast("Выберите хотя бы один день");
      return false;
    }
    const kind = el.weekTaskForm.querySelector('input[name="weekTaskKind"]:checked')?.value || "numeric";
    const target = kind === "check" ? 1 : Math.max(1, Number(el.weekTaskTarget.value) || 1);
    const unit = kind === "check" ? "" : String(el.weekTaskUnit.value || "").trim();
    const moveFrom = String(el.weekTaskMoveFrom.value || "");

    stashCurrentDraft();

    // optional move: remove from source day
    if (moveFrom) {
      const [fromDate, taskId] = moveFrom.split("::");
      const source = tasksForDate(managerId, fromDate).filter((t) => t.id !== taskId);
      state.weekDrafts[draftKey(managerId, fromDate)] = source.length ? source : blankTaskRow(fromDate);
    }

    const baseTask = {
      text,
      target,
      doneCount: 0,
      unit,
      mandatory: kind === "kpi",
      kind,
      kpiId: kind === "kpi" ? `kpi-custom-${Date.now()}` : null,
    };

    dates.forEach((date) => {
      const list = tasksForDate(managerId, date).filter((t) => String(t.text || "").trim());
      list.push({
        ...baseTask,
        id: `local-${Date.now()}-${date}-${Math.random().toString(16).slice(2, 6)}`,
        transferDate: addDaysISO(date, 1),
      });
      state.weekDrafts[draftKey(managerId, date)] = list;
    });

    // persist immediately so Friday→Monday tasks survive
    try {
      await persistWeekDrafts(managerId, dates);
      toast(
        dates.length === 1
          ? `Задача добавлена на ${weekdayShort(dates[0])}`
          : `Задача добавлена на ${dates.length} дн.`
      );
    } catch (err) {
      toast(err.message);
      return false;
    }

    state.planDate = dates.includes(state.planDate) ? state.planDate : dates[0];
    state.draftTasks = tasksForDate(managerId, state.planDate);
    renderWeekDays(managerId);
    renderTasksEditor();
    updateDayPlanLabels();
    return true;
  }

  async function persistWeekDrafts(managerId, dates) {
    const days = [];
    const allDates = new Set([...(dates || []), ...workdaysFromToday(5)]);
    for (const date of allDates) {
      const form = state.forms.find((f) => f.managerId === managerId && f.date === date);
      if (form?.status === "completed") continue;
      const tasksSource = state.weekDrafts[draftKey(managerId, date)] || form?.tasks || [];
      const tasks = tasksSource
        .map((t) => ({
          id: t.id,
          text: String(t.text || "").trim(),
          target: Math.max(1, Number(t.target) || 1),
          doneCount: Math.max(0, Number(t.doneCount) || 0),
          unit: t.unit || "",
          mandatory: Boolean(t.mandatory) || t.kind === "kpi",
          kpiId: t.kpiId || null,
          kind: t.kind || "numeric",
          carriedFrom: t.carriedFrom || null,
        }))
        .filter((t) => t.text);
      if (tasks.length) days.push({ date, tasks });
    }
    if (!days.length) {
      throw new Error("Нет задач для сохранения — добавьте через «+ Добавить задачи»");
    }
    const data = await api("/api/forms/week", {
      method: "POST",
      body: JSON.stringify({ managerId, from: state.today, days }),
    });
    (data.forms || []).forEach(upsertForm);
    if (data.tomorrowPreview) state.tomorrowPreview = data.tomorrowPreview;
    // clear drafts for saved dates — reload from forms
    days.forEach((d) => {
      delete state.weekDrafts[draftKey(managerId, d.date)];
    });
  }

  function hideTomorrowWindow() {
    el.tomorrowWindow.hidden = true;
    el.tomorrowFab.classList.add("is-hidden");
  }

  function renderTomorrowWindow(managerId) {
    const preview =
      state.tomorrowPreview.find((t) => t.managerId === managerId) || {
        date: state.tomorrow,
        tasks: [],
        taskCount: 0,
        managerName: "",
      };
    // prefer live draft / forms over stale preview
    const liveForm = state.forms.find((f) => f.managerId === managerId && f.date === state.tomorrow);
    const liveDraft = state.weekDrafts[draftKey(managerId, state.tomorrow)];
    const tasks = (liveDraft || liveForm?.tasks || preview.tasks || []).filter((t) =>
      String(t.text || "").trim()
    );

    el.tomorrowFabCount.textContent = String(tasks.length);
    el.tomorrowTitle.textContent = tasks.length
      ? `Завтра ждёт ${tasks.length} ${tasks.length === 1 ? "задача" : tasks.length < 5 ? "задачи" : "задач"}`
      : "На завтра пока пусто";
    el.tomorrowDate.textContent = formatDateRu(state.tomorrow);
    el.tomorrowList.innerHTML = "";
    if (!tasks.length) {
      el.tomorrowList.innerHTML =
        `<li class="tomorrow-empty">Составьте план на завтра в ленте «Неделя вперёд»</li>`;
    } else {
      tasks.forEach((t) => {
        const li = document.createElement("li");
        const carried = t.carriedFrom ? " · перенос" : "";
        li.innerHTML = `<strong></strong><span></span>`;
        li.querySelector("strong").textContent = t.text;
        li.querySelector("span").textContent =
          `план ${t.target || 1} ${t.unit || "шт."}${t.mandatory ? " · KPI" : ""}${carried}`;
        el.tomorrowList.appendChild(li);
      });
    }

    if (state.tomorrowCollapsed) {
      el.tomorrowWindow.hidden = true;
      el.tomorrowFab.classList.remove("is-hidden");
    } else {
      el.tomorrowWindow.hidden = false;
      el.tomorrowFab.classList.add("is-hidden");
    }
  }

  function renderTasksEditor() {
    el.tasksEditor.innerHTML = "";
    const evening = state.phase === "evening";
    const form = todaysForm(state.selectedId, state.planDate);
    const locked = form?.status === "completed";
    const day = state.planDate || state.today;

    state.draftTasks.forEach((task, index) => {
      if (!task.kind) {
        task.kind = task.mandatory || task.kpiId ? "kpi" : task.target > 1 || task.unit ? "numeric" : "check";
      }
      const isCheck = task.kind === "check";
      const row = document.createElement("div");
      row.className = `task-row task-row-rich kind-${task.kind}`;
      const pct = progressPct(task);
      row.innerHTML = `
        <div class="task-main">
          <input type="text" class="task-input" placeholder="Задача / KPI" />
          <span class="kind-badge"></span>
        </div>
        <div class="task-metrics">
          <label class="metric-plan">План <input type="number" class="task-target" min="1" step="1" /></label>
          <label class="metric-done">Факт <input type="number" class="task-done" min="0" step="1" /></label>
          <label class="metric-unit">Ед. <input type="text" class="task-unit" placeholder="звонков" /></label>
          <label class="metric-check ${isCheck ? "" : "is-hidden"}">
            <input type="checkbox" class="task-check-done" /> Сделано
          </label>
          <span class="task-pct">${isCheck ? (task.doneCount >= 1 ? "✓" : "—") : pct + "%"}</span>
        </div>
        <div class="task-transfer ${evening ? "" : "is-hidden"}">
          <label>Перенос остатка на
            <input type="date" class="task-transfer-date" />
          </label>
          <span class="transfer-hint"></span>
        </div>
        <div class="task-meta-line"></div>
        <button type="button" class="task-remove" title="Удалить">×</button>
      `;

      const textInput = row.querySelector(".task-input");
      const badge = row.querySelector(".kind-badge");
      const targetInput = row.querySelector(".task-target");
      const doneInput = row.querySelector(".task-done");
      const unitInput = row.querySelector(".task-unit");
      const checkDone = row.querySelector(".task-check-done");
      const transferDate = row.querySelector(".task-transfer-date");
      const pctEl = row.querySelector(".task-pct");
      const hint = row.querySelector(".transfer-hint");
      const meta = row.querySelector(".task-meta-line");

      badge.textContent = task.kind === "kpi" ? "KPI" : isCheck ? "да/нет" : "число";
      textInput.value = task.text || "";
      targetInput.value = Math.max(1, Number(task.target) || 1);
      doneInput.value = Math.max(0, Number(task.doneCount) || 0);
      unitInput.value = task.unit || "";
      checkDone.checked = Number(task.doneCount) >= 1;
      transferDate.value = task.transferDate || addDaysISO(day, 1);

      if (isCheck) {
        row.querySelector(".metric-plan").classList.add("is-hidden");
        row.querySelector(".metric-done").classList.add("is-hidden");
        row.querySelector(".metric-unit").classList.add("is-hidden");
      }

      textInput.readOnly = locked || evening;
      targetInput.readOnly = locked || evening || task.kind === "kpi";
      doneInput.readOnly = locked || !evening;
      unitInput.readOnly = locked || evening || task.kind === "kpi";
      checkDone.disabled = locked || !evening;
      transferDate.disabled = locked || !evening;
      row.querySelector(".task-remove").hidden = locked || evening || task.kind === "kpi";

      if (task.carriedFrom) {
        meta.innerHTML = `↩ из ${task.carriedFrom.date}: остаток ${task.carriedFrom.amount}`;
      }
      if (task.carriedTo) {
        meta.innerHTML += `${meta.innerHTML ? " · " : ""}→ перенесено на ${task.carriedTo.date} (${task.carriedTo.amount})`;
      }

      const sync = () => {
        task.text = textInput.value;
        if (isCheck) {
          task.target = 1;
          task.unit = "";
          if (evening) task.doneCount = checkDone.checked ? 1 : 0;
          else task.doneCount = Math.max(0, Math.min(1, Number(task.doneCount) || 0));
        } else {
          task.target = Math.max(1, Number(targetInput.value) || 1);
          task.doneCount = Math.max(0, Math.min(task.target, Number(doneInput.value) || 0));
          task.unit = unitInput.value.trim();
        }
        task.mandatory = task.kind === "kpi";
        task.transferDate = transferDate.value || addDaysISO(day, 1);
        task.done = task.doneCount >= task.target;
        if (isCheck) pctEl.textContent = task.done ? "✓" : "—";
        else {
          const p = progressPct(task);
          pctEl.textContent = `${p}%`;
          pctEl.classList.toggle("is-good", p >= 100);
          pctEl.classList.toggle("is-mid", p > 0 && p < 100);
        }
        const rem = task.target - task.doneCount;
        hint.textContent =
          rem > 0
            ? `Остаток ${rem}${task.unit ? " " + task.unit : ""} → ${task.transferDate}`
            : "Всё закрыто, перенос не нужен";
        stashCurrentDraft();
        clearTimeout(renderTasksEditor._previewTimer);
        renderTasksEditor._previewTimer = setTimeout(() => {
          if (state.selectedId && state.selectedId !== "all") {
            renderWeekDays(state.selectedId);
            renderTomorrowWindow(state.selectedId);
          }
        }, 180);
      };
      sync();

      [textInput, targetInput, doneInput, unitInput, transferDate, checkDone].forEach((node) => {
        node.addEventListener("input", sync);
        node.addEventListener("change", sync);
      });

      row.querySelector(".task-remove").addEventListener("click", () => {
        state.draftTasks.splice(index, 1);
        if (!state.draftTasks.length) state.draftTasks = blankTaskRow(day);
        stashCurrentDraft();
        renderTasksEditor();
        renderWeekDays(state.selectedId);
        renderTomorrowWindow(state.selectedId);
      });

      el.tasksEditor.appendChild(row);
    });
  }

  function renderManagerStats(managerId) {
    const stats = state.analytics.find((a) => a.managerId === managerId);
    if (!stats) {
      el.managerStats.innerHTML = "";
      return;
    }
    el.managerStats.innerHTML = `
      <div class="stat"><b>${stats.tasksFullyDone}</b><span>сделано полностью</span></div>
      <div class="stat"><b>${stats.tasksPartial}</b><span>частично</span></div>
      <div class="stat"><b>${stats.tasksCarried}</b><span>перенесено</span></div>
      <div class="stat"><b>${stats.conversion}%</b><span>конверсия ед.</span></div>`;
  }

  function renderStickerPack() {
    el.stickerPack.innerHTML = "";
    state.stickerPack.forEach((pack) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = `pack-chip vibe-${pack.vibe || "cyan"}`;
      btn.innerHTML = `<span class="pack-emoji">${pack.emoji || ""}</span><span>${pack.text}</span>`;
      btn.title = "Добавить на доску";
      btn.addEventListener("click", () => addPackSticker(pack.packId));
      el.stickerPack.appendChild(btn);
    });
  }

  async function addPackSticker(packId) {
    const managerId = state.selectedId;
    if (!managerId || managerId === "all") return;
    try {
      const data = await api(`/api/boards/${managerId}/stickers`, {
        method: "POST",
        body: JSON.stringify({ packId }),
      });
      if (!state.boards[managerId]) state.boards[managerId] = [];
      state.boards[managerId].push(data.sticker);
      el.stickerBoard.appendChild(createStickerNode(managerId, data.sticker));
    } catch (err) {
      toast(err.message);
    }
  }

  function renderStickers(managerId) {
    [...el.stickerBoard.querySelectorAll(".sticker")].forEach((n) => n.remove());
    (state.boards[managerId] || []).forEach((sticker) => {
      el.stickerBoard.appendChild(createStickerNode(managerId, sticker));
    });
  }

  function createStickerNode(managerId, sticker, { startEditing = false } = {}) {
    const node = document.createElement("div");
    const isPack = sticker.kind === "pack";
    node.className = `sticker sticker-${sticker.color || "cyan"}${isPack ? " sticker-pack-item" : ""}`;
    if (sticker.vibe) node.classList.add(`vibe-${sticker.vibe}`);
    node.dataset.id = sticker.id;
    node.style.left = `${sticker.x}px`;
    node.style.top = `${sticker.y}px`;
    node.style.transform = `rotate(${sticker.rotation || 0}deg)`;
    node.innerHTML = `
      ${isPack ? `<div class="pack-emoji-lg">${sticker.emoji || ""}</div>` : ""}
      <textarea class="sticker-text" rows="${isPack ? 2 : 4}" spellcheck="false" readonly></textarea>
      <div class="sticker-tools">
        ${isPack ? "" : '<button type="button" data-act="edit" title="Редактировать">✎</button>'}
        ${isPack ? "" : '<button type="button" data-act="color" title="Цвет">◐</button>'}
        <button type="button" data-act="del" title="Удалить">×</button>
      </div>`;
    const textEl = node.querySelector(".sticker-text");
    textEl.value = sticker.text || "";

    const finishEdit = () => {
      node.classList.remove("is-editing");
      textEl.readOnly = true;
      sticker.text = textEl.value.trim() || "Приоритет";
      textEl.value = sticker.text;
      queueStickersSave(managerId);
    };
    const startEdit = () => {
      if (isPack) return;
      node.classList.add("is-editing");
      textEl.readOnly = false;
      textEl.focus();
      textEl.select();
    };

    textEl.addEventListener("dblclick", (e) => {
      e.stopPropagation();
      startEdit();
    });
    textEl.addEventListener("pointerdown", (e) => {
      if (!textEl.readOnly) e.stopPropagation();
    });
    textEl.addEventListener("blur", () => {
      if (!textEl.readOnly) finishEdit();
    });
    textEl.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        textEl.blur();
      }
    });

    const editBtn = node.querySelector('[data-act="edit"]');
    if (editBtn) editBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      startEdit();
    });
    const colorBtn = node.querySelector('[data-act="color"]');
    if (colorBtn) {
      colorBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        const idx = COLORS.indexOf(sticker.color);
        sticker.color = COLORS[(idx + 1) % COLORS.length];
        node.className = `sticker sticker-${sticker.color}${node.classList.contains("is-editing") ? " is-editing" : ""}`;
        queueStickersSave(managerId);
      });
    }
    node.querySelector('[data-act="del"]').addEventListener("click", async (e) => {
      e.stopPropagation();
      try {
        await api(`/api/boards/${managerId}/stickers/${sticker.id}`, { method: "DELETE" });
        state.boards[managerId] = (state.boards[managerId] || []).filter((s) => s.id !== sticker.id);
        node.remove();
      } catch (err) {
        toast(err.message);
      }
    });

    enableDrag(node, sticker, managerId);
    if (startEditing) requestAnimationFrame(startEdit);
    return node;
  }

  function enableDrag(node, sticker, managerId) {
    let dragging = false;
    let startX = 0;
    let startY = 0;
    let origX = 0;
    let origY = 0;
    node.addEventListener("pointerdown", (e) => {
      if (e.target.closest(".sticker-tools") || node.classList.contains("is-editing")) return;
      dragging = true;
      node.setPointerCapture(e.pointerId);
      startX = e.clientX;
      startY = e.clientY;
      origX = sticker.x;
      origY = sticker.y;
      node.style.zIndex = "8";
    });
    node.addEventListener("pointermove", (e) => {
      if (!dragging) return;
      const board = el.stickerBoard.getBoundingClientRect();
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;
      const maxX = board.width - node.offsetWidth - 4;
      const maxY = board.height - node.offsetHeight - 4;
      sticker.x = Math.max(4, Math.min(maxX, origX + dx));
      sticker.y = Math.max(4, Math.min(maxY, origY + dy));
      node.style.left = `${sticker.x}px`;
      node.style.top = `${sticker.y}px`;
    });
    const end = () => {
      if (!dragging) return;
      dragging = false;
      node.style.zIndex = "";
      queueStickersSave(managerId);
    };
    node.addEventListener("pointerup", end);
    node.addEventListener("pointercancel", end);
  }

  function queueStickersSave(managerId) {
    clearTimeout(state.stickerSaveTimer);
    state.stickerSaveTimer = setTimeout(async () => {
      try {
        const data = await api(`/api/boards/${managerId}`, {
          method: "PUT",
          body: JSON.stringify({ stickers: state.boards[managerId] || [] }),
        });
        state.boards[managerId] = data.stickers;
      } catch (err) {
        toast(err.message);
      }
    }, 280);
  }

  function upsertForm(form) {
    const idx = state.forms.findIndex((f) => f.id === form.id);
    if (idx >= 0) state.forms[idx] = form;
    else state.forms.push(form);
  }

  async function saveForm() {
    const managerId = state.selectedId;
    if (!managerId || managerId === "all") return;
    const date = state.planDate || state.today;
    try {
      if (state.phase === "morning") {
        stashCurrentDraft();
        const tasks = state.draftTasks
          .map((t) => ({
            id: t.id,
            text: String(t.text || "").trim(),
            target: Math.max(1, Number(t.target) || 1),
            doneCount: 0,
            unit: t.unit || "",
            mandatory: Boolean(t.mandatory) || t.kind === "kpi",
            kpiId: t.kpiId || null,
            kind: t.kind || "numeric",
            carriedFrom: t.carriedFrom || null,
          }))
          .filter((t) => t.text);
        if (!tasks.length) {
          toast("Добавьте хотя бы одну задачу");
          return;
        }
        const data = await api("/api/forms/morning", {
          method: "POST",
          body: JSON.stringify({ managerId, date, tasks, seedKpis: false }),
        });
        upsertForm(data.form);
        if (data.tomorrowPreview) state.tomorrowPreview = data.tomorrowPreview;
        delete state.weekDrafts[draftKey(managerId, date)];
        toast(data.message || "План сохранён");
        if (date === state.today) {
          state.phase = "evening";
          state.draftTasks = data.form.tasks.map((t) => ({
            ...t,
            transferDate: t.transferDate || addDaysISO(date, 1),
          }));
        } else {
          state.draftTasks = data.form.tasks.map((t) => ({ ...t }));
        }
        await refreshAnalytics();
        renderAll();
      } else {
        if (date !== state.today) {
          toast("Закрывать день можно только сегодня");
          return;
        }
        const tasks = state.draftTasks.map((t) => ({
          id: t.id,
          text: String(t.text || "").trim(),
          target: Math.max(1, Number(t.target) || 1),
          doneCount: Math.max(0, Number(t.doneCount) || 0),
          unit: t.unit || "",
          mandatory: Boolean(t.mandatory) || t.kind === "kpi",
          kpiId: t.kpiId || null,
          kind: t.kind || "numeric",
          transferDate: t.transferDate || addDaysISO(state.today, 1),
          carriedFrom: t.carriedFrom || null,
        }));
        const data = await api("/api/forms/evening", {
          method: "POST",
          body: JSON.stringify({ managerId, date: state.today, tasks }),
        });
        upsertForm(data.form);
        if (data.destForm) upsertForm(data.destForm);
        (data.transfers || []).forEach((tr) => {
          state.transfers.unshift(tr);
        });
        if (data.analytics) state.analytics = data.analytics;
        if (data.dashboard) state.dashboard = data.dashboard;
        toast(data.message || "День закрыт");
        state.weekDrafts = {};
        await loadState();
      }
    } catch (err) {
      toast(err.message);
    }
  }

  async function saveWeekPlan() {
    const managerId = state.selectedId;
    if (!managerId || managerId === "all") return;
    stashCurrentDraft();
    try {
      await persistWeekDrafts(managerId, workdaysFromToday(5));
      toast("План на будни сохранён");
      await loadState();
      if (state.selectedId === managerId) {
        state.planDate = state.planDate || state.today;
        state.draftTasks = tasksForDate(managerId, state.planDate);
        renderWeekDays(managerId);
        renderTasksEditor();
      }
    } catch (err) {
      toast(err.message || "Нечего сохранять — добавьте задачи");
    }
  }

  async function refreshAnalytics() {
    try {
      const data = await api("/api/analytics");
      state.analytics = data.analytics || [];
      state.dashboard = data.dashboard || null;
      state.transfers = data.transfers || state.transfers;
    } catch (_) {
      /* ignore */
    }
  }

  function renderDashboard() {
    // KPI table
    el.kpiTable.innerHTML = "";
    const rows = state.dashboard?.table || [];
    if (!rows.length) {
      el.kpiTable.innerHTML = `<tr><td colspan="5" class="muted">Пока нет данных за сегодня</td></tr>`;
    } else {
      rows.forEach((r) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${r.managerName}</td>
          <td>${r.kpiName}</td>
          <td>${r.target} ${r.unit || ""}</td>
          <td>${r.doneCount} ${r.unit || ""}</td>
          <td><strong class="${r.conversion >= 80 ? "good" : r.conversion >= 50 ? "mid" : "low"}">${r.conversion}%</strong>
            ${r.hasData ? "" : '<span class="muted"> · нет в чеклисте</span>'}</td>`;
        el.kpiTable.appendChild(tr);
      });
    }

    // analytics cards
    el.analyticsCards.innerHTML = `<h3 class="dash-title">Сводка по менеджерам</h3><div class="analytics-grid"></div>`;
    const grid = el.analyticsCards.querySelector(".analytics-grid");
    (state.analytics || []).forEach((a) => {
      const card = document.createElement("div");
      card.className = "analytics-card";
      card.innerHTML = `
        <h4></h4>
        <div class="mini-stats">
          <div><b>${a.tasksFullyDone}</b><span>сделано</span></div>
          <div><b>${a.tasksCarried}</b><span>перенесено</span></div>
          <div><b>${a.conversion}%</b><span>конверсия</span></div>
        </div>
        <p class="muted">${a.unitsDone}/${a.unitsPlanned} ед. факта</p>`;
      card.querySelector("h4").textContent = a.managerName;
      grid.appendChild(card);
    });

    // transfers journal
    el.transfersList.innerHTML = "";
    const transfers = [...(state.transfers || [])].sort((a, b) =>
      String(b.createdAt || "").localeCompare(String(a.createdAt || ""))
    );
    if (!transfers.length) {
      el.transfersList.innerHTML = `<div class="archive-empty">Переносов пока нет</div>`;
    } else {
      transfers.slice(0, 40).forEach((tr) => {
        const item = document.createElement("div");
        item.className = "transfer-item";
        item.innerHTML = `
          <strong></strong>
          <span></span>
          <em></em>`;
        item.querySelector("strong").textContent = tr.managerName;
        item.querySelector("span").textContent =
          `${tr.text}: остаток ${tr.amount} ${tr.unit || ""}`.trim();
        item.querySelector("em").textContent = `${tr.fromDate} → ${tr.toDate}`;
        el.transfersList.appendChild(item);
      });
    }

    drawChart();
  }

  function drawChart() {
    const canvas = el.convChart;
    if (!canvas || !state.dashboard) return;
    const ctx = canvas.getContext("2d");
    const dpr = window.devicePixelRatio || 1;
    const cssW = canvas.clientWidth || 720;
    const cssH = 320;
    canvas.width = cssW * dpr;
    canvas.height = cssH * dpr;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, cssW, cssH);

    const pad = { l: 42, r: 16, t: 18, b: 36 };
    const w = cssW - pad.l - pad.r;
    const h = cssH - pad.t - pad.b;
    const dates = state.dashboard.dates || [];
    const series = state.dashboard.series || [];

    // grid
    ctx.strokeStyle = "rgba(38,166,224,0.15)";
    ctx.fillStyle = "#9aa8b5";
    ctx.font = "12px Manrope, sans-serif";
    for (let i = 0; i <= 4; i += 1) {
      const y = pad.t + (h * i) / 4;
      ctx.beginPath();
      ctx.moveTo(pad.l, y);
      ctx.lineTo(pad.l + w, y);
      ctx.stroke();
      ctx.fillText(`${100 - i * 25}%`, 4, y + 4);
    }

    el.chartLegend.innerHTML = "";
    series.forEach((s, idx) => {
      const color = LINE_COLORS[idx % LINE_COLORS.length];
      const legend = document.createElement("span");
      legend.className = "legend-item";
      legend.innerHTML = `<i style="background:${color}"></i>${s.managerName}`;
      el.chartLegend.appendChild(legend);

      const pts = s.points
        .map((p, i) => ({
          x: pad.l + (dates.length <= 1 ? w / 2 : (w * i) / (dates.length - 1)),
          y: p.conversion == null ? null : pad.t + h * (1 - p.conversion / 100),
          raw: p,
        }))
        .filter((p) => p.y != null);

      if (pts.length < 2) {
        // single point
        pts.forEach((p) => {
          ctx.fillStyle = color;
          ctx.beginPath();
          ctx.arc(p.x, p.y, 4, 0, Math.PI * 2);
          ctx.fill();
        });
        return;
      }

      ctx.strokeStyle = color;
      ctx.lineWidth = 2.5;
      ctx.beginPath();
      pts.forEach((p, i) => {
        if (i === 0) ctx.moveTo(p.x, p.y);
        else ctx.lineTo(p.x, p.y);
      });
      ctx.stroke();
      pts.forEach((p) => {
        ctx.fillStyle = color;
        ctx.beginPath();
        ctx.arc(p.x, p.y, 3.5, 0, Math.PI * 2);
        ctx.fill();
      });
    });

    // x labels sparse
    ctx.fillStyle = "#9aa8b5";
    dates.forEach((d, i) => {
      if (dates.length > 8 && i % 2 !== 0 && i !== dates.length - 1) return;
      const x = pad.l + (dates.length <= 1 ? w / 2 : (w * i) / (dates.length - 1));
      const label = d.slice(5);
      ctx.fillText(label, x - 14, cssH - 12);
    });
  }

  function renderArchive() {
    const archive = state.forms
      .filter((f) => f.status === "completed")
      .sort(
        (a, b) =>
          String(b.date).localeCompare(String(a.date)) || a.title.localeCompare(b.title, "ru")
      );
    el.archiveList.innerHTML = "";
    if (!archive.length) {
      el.archiveList.innerHTML =
        `<div class="archive-empty">Пока пусто. После вечернего сохранения формы появятся здесь.</div>`;
      return;
    }
    archive.forEach((form) => {
      const item = document.createElement("article");
      item.className = "archive-item";
      const unitsDone = form.tasks.reduce((s, t) => s + (Number(t.doneCount) || 0), 0);
      const unitsPlan = form.tasks.reduce((s, t) => s + Math.max(1, Number(t.target) || 1), 0);
      item.innerHTML = `
        <div>
          <h3></h3>
          <ul></ul>
        </div>
        <div class="archive-side">
          <span class="archive-badge"></span>
          <button type="button" class="btn btn-ghost btn-sm btn-danger" data-del>Удалить</button>
        </div>`;
      item.querySelector("h3").textContent = form.title;
      item.querySelector(".archive-badge").textContent =
        `${unitsDone}/${unitsPlan} ед. · ${unitsPlan ? Math.round((unitsDone / unitsPlan) * 100) : 0}%`;
      const ul = item.querySelector("ul");
      form.tasks.forEach((t) => {
        const li = document.createElement("li");
        li.textContent =
          `${t.text}: ${t.doneCount}/${t.target} ${t.unit || ""} (${progressPct(t)}%)`.trim() +
          (t.carriedTo ? ` → ${t.carriedTo.date}` : "");
        if (t.done) li.className = "done";
        ul.appendChild(li);
      });
      item.querySelector("[data-del]").addEventListener("click", async () => {
        if (!confirm(`Удалить «${form.title}»?`)) return;
        try {
          await api(`/api/forms/${form.id}`, { method: "DELETE" });
          state.forms = state.forms.filter((f) => f.id !== form.id);
          toast(`Удалено: ${form.title}`);
          renderAll();
        } catch (err) {
          toast(err.message);
        }
      });
      el.archiveList.appendChild(item);
    });
  }

  async function renameManager(manager) {
    const name = await askPrompt({
      title: "Переименовать менеджера",
      label: "ФИО",
      initial: manager.name,
    });
    if (!name || name === manager.name) return;
    try {
      const data = await api(`/api/managers/${manager.id}`, {
        method: "PATCH",
        body: JSON.stringify({ name }),
      });
      manager.name = data.manager.name;
      state.forms.forEach((f) => {
        if (f.managerId === manager.id) {
          f.managerName = manager.name;
          f.title = `${manager.name} ${f.date}`;
        }
      });
      toast(`Имя обновлено: ${manager.name}`);
      renderAll();
    } catch (err) {
      toast(err.message);
    }
  }

  el.addManagerBtn.addEventListener("click", async () => {
    const name = await askPrompt({
      title: "Новый менеджер",
      label: "Имя сотрудника",
      placeholder: "Фамилия Имя",
    });
    if (!name) return;
    try {
      const data = await api("/api/managers", {
        method: "POST",
        body: JSON.stringify({ name }),
      });
      state.managers.push(data.manager);
      state.boards[data.manager.id] = [];
      state.selectedId = data.manager.id;
      toast(`Добавлен: ${data.manager.name}`);
      renderAll();
    } catch (err) {
      toast(err.message);
    }
  });

  el.renameManagerBtn.addEventListener("click", () => {
    const manager = state.managers.find((m) => m.id === state.selectedId);
    if (manager) renameManager(manager);
  });

  el.addSticker.addEventListener("click", async () => {
    const managerId = state.selectedId;
    if (!managerId || managerId === "all") return;
    try {
      const data = await api(`/api/boards/${managerId}/stickers`, {
        method: "POST",
        body: JSON.stringify({ text: "Новый приоритет", color: "cyan" }),
      });
      if (!state.boards[managerId]) state.boards[managerId] = [];
      state.boards[managerId].push(data.sticker);
      el.stickerBoard.appendChild(
        createStickerNode(managerId, data.sticker, { startEditing: true })
      );
    } catch (err) {
      toast(err.message);
    }
  });


  el.tabMorning.addEventListener("click", () => {
    const form = todaysForm(state.selectedId, state.planDate);
    if (form) state.draftTasks = form.tasks.map((t) => ({ ...t }));
    setPhase("morning");
    renderTasksEditor();
  });

  el.tabEvening.addEventListener("click", () => {
    if (state.planDate !== state.today) {
      toast("Вечер доступен только для сегодня");
      return;
    }
    const form = todaysForm(state.selectedId, state.today);
    if (form) {
      state.draftTasks = form.tasks.map((t) => ({
        ...t,
        transferDate: t.transferDate || addDaysISO(state.today, 1),
      }));
    }
    setPhase("evening");
    renderTasksEditor();
  });

  function addTaskOfKind(kind) {
    const day = state.planDate || state.today;
    if (kind === "kpi") {
      const existing = new Set(state.draftTasks.filter((t) => t.kpiId).map((t) => t.kpiId));
      const month = monthKey(day);
      const monthly = state.monthlyKpis?.[state.selectedId]?.[month]?.items || [];
      const source = monthly.length ? monthly : state.kpiDefs;
      let added = 0;
      source.forEach((k) => {
        const id = k.id || k.name;
        if (existing.has(id)) return;
        state.draftTasks.push({
          id: `local-${Date.now()}-${id}`,
          text: k.name,
          target: k.target || k.defaultTarget || 1,
          doneCount: 0,
          unit: k.unit || "",
          mandatory: true,
          kpiId: id,
          kind: "kpi",
          transferDate: addDaysISO(day, 1),
        });
        added += 1;
      });
      state.draftTasks = state.draftTasks.filter((t) => t.text || t.kpiId);
      stashCurrentDraft();
      renderTasksEditor();
      renderWeekDays(state.selectedId);
      toast(added ? `Добавлено KPI: ${added}` : "KPI уже в списке");
      return;
    }
    state.draftTasks.push({
      id: `local-${Date.now()}`,
      text: "",
      target: kind === "check" ? 1 : 1,
      doneCount: 0,
      unit: kind === "check" ? "" : "",
      mandatory: false,
      kind,
      transferDate: addDaysISO(day, 1),
    });
    stashCurrentDraft();
    renderTasksEditor();
    renderWeekDays(state.selectedId);
    if (state.planDate && state.planDate !== state.today) {
      persistWeekDrafts(state.selectedId, [state.planDate]).catch((err) => toast(err.message));
    }
  }

  function bindAddTaskMenu(button, menu) {
    if (!button || !menu) return;
    button.addEventListener("click", (e) => {
      e.stopPropagation();
      el.addTaskMenu?.classList.add("is-hidden");
      menu.classList.toggle("is-hidden");
    });
    menu.querySelectorAll("button[data-kind]").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        addTaskOfKind(btn.dataset.kind);
        menu.classList.add("is-hidden");
      });
    });
  }
  bindAddTaskMenu(el.addTaskBtn, el.addTaskMenu);
  document.addEventListener("click", () => {
    el.addTaskMenu?.classList.add("is-hidden");
  });

  el.addMonthlyKpiBtn?.addEventListener("click", addMonthlyKpiRow);
  el.saveMonthlyKpiBtn?.addEventListener("click", saveMonthlyKpi);

  el.buildReportBtn?.addEventListener("click", () => {
    const managerId = state.selectedId;
    if (!managerId || managerId === "all") {
      toast("Сначала откройте сотрудника");
      return;
    }
    const weekStart = el.reportWeekStart?.value || state.today;
    const url = `/report?managerId=${encodeURIComponent(managerId)}&weekStart=${encodeURIComponent(weekStart)}`;
    window.open(url, "_blank", "noopener");
  });

  el.saveFormBtn.addEventListener("click", saveForm);
  el.openWeekTaskPanelBtn?.addEventListener("click", () => {
    openWeekTaskPanel({ preselect: [state.planDate || state.today] });
  });
  el.weekTaskForm?.addEventListener("click", async (e) => {
    const btn = e.target.closest("button[value]");
    if (!btn) return;
    if (btn.value === "ok") {
      e.preventDefault();
      e.stopPropagation();
      const ok = await submitWeekTaskFromPanel();
      if (ok) el.weekTaskModal.close("ok");
      return;
    }
    el.weekTaskModal.returnValue = btn.value;
  });
  el.weekTaskForm?.querySelectorAll('input[name="weekTaskKind"]').forEach((r) => {
    r.addEventListener("change", syncWeekTaskMetricsVisibility);
  });
  el.weekTaskSelectAll?.addEventListener("click", () => {
    el.weekTaskDayGrid?.querySelectorAll("input").forEach((n) => {
      n.checked = true;
    });
  });
  el.weekTaskSelectOne?.addEventListener("click", () => {
    const one = state.planDate || state.today;
    el.weekTaskDayGrid?.querySelectorAll("input").forEach((n) => {
      n.checked = n.value === one;
    });
  });
  document.querySelectorAll(".week-view-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      state.weekView = btn.dataset.view || "all";
      document.querySelectorAll(".week-view-btn").forEach((b) => {
        b.classList.toggle("is-active", b === btn);
      });
      if (state.selectedId && state.selectedId !== "all") {
        renderWeekChecklist(state.selectedId);
      }
    });
  });

  el.saveWeekBtn.addEventListener("click", saveWeekPlan);
  el.tomorrowClose.addEventListener("click", () => {
    state.tomorrowCollapsed = true;
    if (state.selectedId && state.selectedId !== "all") renderTomorrowWindow(state.selectedId);
  });
  el.tomorrowFab.addEventListener("click", () => {
    state.tomorrowCollapsed = false;
    if (state.selectedId && state.selectedId !== "all") renderTomorrowWindow(state.selectedId);
  });
  el.tomorrowOpenDay.addEventListener("click", () => {
    if (!state.selectedId || state.selectedId === "all") return;
    selectPlanDate(state.tomorrow);
    state.tomorrowCollapsed = false;
    renderTomorrowWindow(state.selectedId);
  });
  el.promptForm.addEventListener("click", (e) => {
    const btn = e.target.closest("button[value]");
    if (btn) el.promptModal.returnValue = btn.value;
  });
  window.addEventListener("resize", () => drawChart());

  loadState().catch((err) => toast(err.message || "Не удалось загрузить данные"));
})();
