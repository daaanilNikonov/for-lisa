(() => {
  "use strict";

  const state = {
    today: null,
    managers: [],
    forms: [],
    boards: {},
    kpiDefs: [],
    transfers: [],
    stickerPack: [],
    analytics: [],
    dashboard: null,
    selectedId: "all",
    phase: "morning",
    draftTasks: [],
    stickerSaveTimer: null,
  };

  const COLORS = ["cyan", "amber", "mint", "rose", "violet", "pink"];
  const LINE_COLORS = ["#26a6e0", "#fecf68", "#7dcea0", "#f0a3b0", "#a9b4f5", "#ff8bd2"];

  const el = {
    todayLabel: document.getElementById("todayLabel"),
    managerTabs: document.getElementById("managerTabs"),
    overviewView: document.getElementById("overviewView"),
    singleView: document.getElementById("singleView"),
    checklistTitle: document.getElementById("checklistTitle"),
    checklistSub: document.getElementById("checklistSub"),
    boardTitle: document.getElementById("boardTitle"),
    boardSub: document.getElementById("boardSub"),
    tasksEditor: document.getElementById("tasksEditor"),
    stickerBoard: document.getElementById("stickerBoard"),
    stickerPack: document.getElementById("stickerPack"),
    archiveList: document.getElementById("archiveList"),
    saveFormBtn: document.getElementById("saveFormBtn"),
    formHint: document.getElementById("formHint"),
    tabMorning: document.getElementById("tabMorning"),
    tabEvening: document.getElementById("tabEvening"),
    addTaskBtn: document.getElementById("addTaskBtn"),
    seedKpiBtn: document.getElementById("seedKpiBtn"),
    addManagerBtn: document.getElementById("addManagerBtn"),
    renameManagerBtn: document.getElementById("renameManagerBtn"),
    addSticker: document.getElementById("addSticker"),
    managerStats: document.getElementById("managerStats"),
    kpiTable: document.querySelector("#kpiTable tbody"),
    chartLegend: document.getElementById("chartLegend"),
    convChart: document.getElementById("convChart"),
    analyticsCards: document.getElementById("analyticsCards"),
    transfersList: document.getElementById("transfersList"),
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

  function todaysForm(managerId, date = state.today) {
    return state.forms.find((f) => f.managerId === managerId && f.date === date);
  }

  function statusFor(managerId) {
    const form = todaysForm(managerId);
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
    state.managers = data.managers || [];
    state.forms = data.forms || [];
    state.boards = data.boards || {};
    state.kpiDefs = data.kpiDefs || [];
    state.transfers = data.transfers || [];
    state.stickerPack = data.stickerPack || [];
    state.analytics = data.analytics || [];
    state.dashboard = data.dashboard || null;
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
      const form = todaysForm(manager.id);
      const status = statusFor(manager.id);
      const stats = state.analytics.find((a) => a.managerId === manager.id);
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
        renderAll();
      });
      el.overviewView.appendChild(card);
    });
  }

  function loadManagerWorkspace(managerId) {
    const manager = state.managers.find((m) => m.id === managerId);
    if (!manager) return;
    const form = todaysForm(managerId);
    el.checklistTitle.textContent = `Чеклист · ${manager.name}`;
    el.checklistSub.textContent = `${formatDateRu(state.today)} · «${manager.name} ${state.today}»`;
    el.boardTitle.textContent = "Основные приоритеты в работе";
    el.boardSub.textContent = `Доска · ${manager.name}`;

    if (form?.status === "completed") {
      state.phase = "evening";
      state.draftTasks = form.tasks.map((t) => ({ ...t }));
    } else if (form?.status === "morning") {
      state.phase = "evening";
      state.draftTasks = form.tasks.map((t) => ({ ...t }));
    } else {
      state.phase = "morning";
      state.draftTasks = blankTaskRow();
    }

    setPhase(state.phase);
    renderTasksEditor();
    renderManagerStats(managerId);
    renderStickerPack();
    renderStickers(managerId);
  }

  function blankTaskRow() {
    return [
      {
        id: `local-${Date.now()}`,
        text: "",
        target: 1,
        doneCount: 0,
        unit: "",
        mandatory: false,
        done: false,
        transferDate: addDaysISO(state.today, 1),
      },
    ];
  }

  function setPhase(phase) {
    state.phase = phase;
    el.tabMorning.classList.toggle("is-active", phase === "morning");
    el.tabEvening.classList.toggle("is-active", phase === "evening");
    const form = todaysForm(state.selectedId);
    const completed = form?.status === "completed";
    const hasMorning = Boolean(form && (form.status === "morning" || form.status === "completed"));

    if (phase === "morning") {
      el.saveFormBtn.textContent = completed ? "Уже в архиве" : "Сохранить утро";
      el.saveFormBtn.disabled = completed;
      el.addTaskBtn.hidden = completed;
      el.seedKpiBtn.hidden = completed;
      el.formHint.textContent = completed
        ? "День уже закрыт. Можно смотреть прогресс и переносы."
        : "Утро: план и обязательные KPI. Можно задать цель числом (например, 2 демо).";
    } else {
      el.saveFormBtn.textContent = completed ? "Уже в архиве" : "Закрыть день · перенести остатки";
      el.saveFormBtn.disabled = completed || !hasMorning;
      el.addTaskBtn.hidden = true;
      el.seedKpiBtn.hidden = true;
      el.formHint.textContent = !hasMorning
        ? "Сначала сохраните утренний чеклист."
        : completed
          ? "День в архиве."
          : "Вечер: факт (сколько сделано). Остаток автоматически уйдёт на выбранную дату переноса.";
    }
  }

  function renderTasksEditor() {
    el.tasksEditor.innerHTML = "";
    const evening = state.phase === "evening";
    const form = todaysForm(state.selectedId);
    const locked = form?.status === "completed";

    state.draftTasks.forEach((task, index) => {
      const row = document.createElement("div");
      row.className = "task-row task-row-rich";
      const pct = progressPct(task);
      row.innerHTML = `
        <div class="task-main">
          <input type="text" class="task-input" placeholder="Задача / KPI" />
          <label class="chip-toggle"><input type="checkbox" class="task-mandatory" /> KPI</label>
        </div>
        <div class="task-metrics">
          <label>План <input type="number" class="task-target" min="1" step="1" /></label>
          <label>Факт <input type="number" class="task-done" min="0" step="1" /></label>
          <label>Ед. <input type="text" class="task-unit" placeholder="звонков" /></label>
          <span class="task-pct">${pct}%</span>
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
      const mandatory = row.querySelector(".task-mandatory");
      const targetInput = row.querySelector(".task-target");
      const doneInput = row.querySelector(".task-done");
      const unitInput = row.querySelector(".task-unit");
      const transferDate = row.querySelector(".task-transfer-date");
      const pctEl = row.querySelector(".task-pct");
      const hint = row.querySelector(".transfer-hint");
      const meta = row.querySelector(".task-meta-line");

      textInput.value = task.text || "";
      mandatory.checked = Boolean(task.mandatory);
      targetInput.value = Math.max(1, Number(task.target) || 1);
      doneInput.value = Math.max(0, Number(task.doneCount) || 0);
      unitInput.value = task.unit || "";
      transferDate.value = task.transferDate || addDaysISO(state.today, 1);

      textInput.readOnly = locked || evening;
      mandatory.disabled = locked || evening;
      targetInput.readOnly = locked || evening;
      doneInput.readOnly = locked || !evening;
      unitInput.readOnly = locked || evening;
      transferDate.disabled = locked || !evening;
      row.querySelector(".task-remove").hidden = locked || evening;

      if (task.carriedFrom) {
        meta.innerHTML = `↩ из ${task.carriedFrom.date}: остаток ${task.carriedFrom.amount} (было ${task.carriedFrom.doneCount}/${task.carriedFrom.originalTarget})`;
      }
      if (task.carriedTo) {
        meta.innerHTML += `${meta.innerHTML ? " · " : ""}→ перенесено на ${task.carriedTo.date} (${task.carriedTo.amount})`;
      }

      const sync = () => {
        task.text = textInput.value;
        task.mandatory = mandatory.checked;
        task.target = Math.max(1, Number(targetInput.value) || 1);
        task.doneCount = Math.max(0, Math.min(task.target, Number(doneInput.value) || 0));
        task.unit = unitInput.value.trim();
        task.transferDate = transferDate.value || addDaysISO(state.today, 1);
        task.done = task.doneCount >= task.target;
        const p = progressPct(task);
        pctEl.textContent = `${p}%`;
        pctEl.classList.toggle("is-good", p >= 100);
        pctEl.classList.toggle("is-mid", p > 0 && p < 100);
        const rem = task.target - task.doneCount;
        hint.textContent =
          rem > 0
            ? `Остаток ${rem} ${task.unit || "шт."} → ${task.transferDate}`
            : "Всё закрыто, перенос не нужен";
      };
      sync();

      [textInput, mandatory, targetInput, doneInput, unitInput, transferDate].forEach((node) => {
        node.addEventListener("input", sync);
        node.addEventListener("change", sync);
      });

      row.querySelector(".task-remove").addEventListener("click", () => {
        state.draftTasks.splice(index, 1);
        if (!state.draftTasks.length) state.draftTasks = blankTaskRow();
        renderTasksEditor();
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
    try {
      if (state.phase === "morning") {
        const tasks = state.draftTasks
          .map((t) => ({
            id: t.id,
            text: String(t.text || "").trim(),
            target: Math.max(1, Number(t.target) || 1),
            doneCount: 0,
            unit: t.unit || "",
            mandatory: Boolean(t.mandatory),
            kpiId: t.kpiId || null,
            carriedFrom: t.carriedFrom || null,
          }))
          .filter((t) => t.text);
        if (!tasks.length) {
          toast("Добавьте хотя бы одну задачу");
          return;
        }
        const data = await api("/api/forms/morning", {
          method: "POST",
          body: JSON.stringify({ managerId, date: state.today, tasks, seedKpis: false }),
        });
        upsertForm(data.form);
        toast(data.message || "Утро сохранено");
        state.phase = "evening";
        state.draftTasks = data.form.tasks.map((t) => ({
          ...t,
          transferDate: t.transferDate || addDaysISO(state.today, 1),
        }));
        await refreshAnalytics();
        renderAll();
      } else {
        const tasks = state.draftTasks.map((t) => ({
          id: t.id,
          text: String(t.text || "").trim(),
          target: Math.max(1, Number(t.target) || 1),
          doneCount: Math.max(0, Number(t.doneCount) || 0),
          unit: t.unit || "",
          mandatory: Boolean(t.mandatory),
          kpiId: t.kpiId || null,
          transferDate: t.transferDate || addDaysISO(state.today, 1),
          carriedFrom: t.carriedFrom || null,
        }));
        const data = await api("/api/forms/evening", {
          method: "POST",
          body: JSON.stringify({ managerId, date: state.today, tasks }),
        });
        upsertForm(data.form);
        if (data.destForm) upsertForm(data.destForm);
        // refresh forms that received transfers
        (data.transfers || []).forEach((tr) => {
          state.transfers.unshift(tr);
        });
        if (data.analytics) state.analytics = data.analytics;
        if (data.dashboard) state.dashboard = data.dashboard;
        toast(data.message || "День закрыт");
        await loadState();
      }
    } catch (err) {
      toast(err.message);
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

  el.seedKpiBtn.addEventListener("click", () => {
    const existingKpis = new Set(state.draftTasks.filter((t) => t.kpiId).map((t) => t.kpiId));
    const toAdd = state.kpiDefs.filter((k) => !existingKpis.has(k.id));
    if (!toAdd.length) {
      toast("Обязательные KPI уже в списке");
      return;
    }
    toAdd.forEach((k) => {
      state.draftTasks.push({
        id: `local-${Date.now()}-${k.id}`,
        text: k.name,
        target: k.defaultTarget,
        doneCount: 0,
        unit: k.unit,
        mandatory: true,
        kpiId: k.id,
        transferDate: addDaysISO(state.today, 1),
      });
    });
    // remove empty placeholder rows
    state.draftTasks = state.draftTasks.filter((t) => t.text || t.kpiId);
    renderTasksEditor();
    toast("Добавлены обязательные KPI");
  });

  el.tabMorning.addEventListener("click", () => {
    const form = todaysForm(state.selectedId);
    if (form) state.draftTasks = form.tasks.map((t) => ({ ...t }));
    setPhase("morning");
    renderTasksEditor();
  });

  el.tabEvening.addEventListener("click", () => {
    const form = todaysForm(state.selectedId);
    if (form) {
      state.draftTasks = form.tasks.map((t) => ({
        ...t,
        transferDate: t.transferDate || addDaysISO(state.today, 1),
      }));
    }
    setPhase("evening");
    renderTasksEditor();
  });

  el.addTaskBtn.addEventListener("click", () => {
    state.draftTasks.push({
      id: `local-${Date.now()}`,
      text: "",
      target: 1,
      doneCount: 0,
      unit: "",
      mandatory: false,
      transferDate: addDaysISO(state.today, 1),
    });
    renderTasksEditor();
  });

  el.saveFormBtn.addEventListener("click", saveForm);
  el.promptForm.addEventListener("click", (e) => {
    const btn = e.target.closest("button[value]");
    if (btn) el.promptModal.returnValue = btn.value;
  });
  window.addEventListener("resize", () => drawChart());

  loadState().catch((err) => toast(err.message || "Не удалось загрузить данные"));
})();
