(() => {
  "use strict";

  const state = {
    today: null,
    managers: [],
    forms: [],
    boards: {},
    selectedId: "all",
    phase: "morning",
    draftTasks: [],
    stickerSaveTimer: null,
  };

  const COLORS = ["cyan", "amber", "mint", "rose", "violet"];

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
    archiveList: document.getElementById("archiveList"),
    saveFormBtn: document.getElementById("saveFormBtn"),
    formHint: document.getElementById("formHint"),
    tabMorning: document.getElementById("tabMorning"),
    tabEvening: document.getElementById("tabEvening"),
    addTaskBtn: document.getElementById("addTaskBtn"),
    addManagerBtn: document.getElementById("addManagerBtn"),
    addSticker: document.getElementById("addSticker"),
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
    toast._t = setTimeout(() => el.toast.classList.remove("is-on"), 3200);
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

  function todaysForm(managerId) {
    return state.forms.find((f) => f.managerId === managerId && f.date === state.today);
  }

  function statusFor(managerId) {
    const form = todaysForm(managerId);
    if (!form) return { key: "idle", label: "Ещё не начато" };
    if (form.status === "completed") return { key: "done", label: "В архиве" };
    if (form.status === "morning") return { key: "morning", label: "Чеклист задан" };
    return { key: "idle", label: "Ещё не начато" };
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
    el.todayLabel.textContent = `Сегодня · ${formatDateRu(state.today)}`;
    if (state.selectedId !== "all" && !state.managers.some((m) => m.id === state.selectedId)) {
      state.selectedId = "all";
    }
    renderAll();
  }

  function renderAll() {
    renderTabs();
    if (state.selectedId === "all") {
      el.overviewView.classList.remove("is-hidden");
      el.singleView.classList.add("is-hidden");
      renderOverview();
    } else {
      el.overviewView.classList.add("is-hidden");
      el.singleView.classList.remove("is-hidden");
      loadManagerWorkspace(state.selectedId);
    }
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
      btn.addEventListener("click", () => {
        state.selectedId = manager.id;
        renderAll();
      });
      el.managerTabs.appendChild(btn);
    });
  }

  function renderOverview() {
    el.overviewView.innerHTML = "";
    if (!state.managers.length) {
      el.overviewView.innerHTML = `<div class="archive-empty">Добавьте менеджера, чтобы начать</div>`;
      return;
    }

    state.managers.forEach((manager) => {
      const form = todaysForm(manager.id);
      const status = statusFor(manager.id);
      const stickers = state.boards[manager.id] || [];
      const card = document.createElement("article");
      card.className = "overview-card";
      card.innerHTML = `
        <header class="overview-card-head">
          <div>
            <h3></h3>
            <span class="manager-status ${
              status.key === "done" ? "is-done" : status.key === "morning" ? "is-morning" : ""
            }"></span>
          </div>
          <button type="button" class="btn btn-ghost btn-sm" data-open>Открыть</button>
        </header>
        <div class="overview-grid">
          <div>
            <p class="overview-label">Чеклист на день</p>
            <ul class="overview-tasks"></ul>
          </div>
          <div>
            <p class="overview-label">Доска приоритетов</p>
            <div class="overview-stickers"></div>
          </div>
        </div>
      `;
      card.querySelector("h3").textContent = manager.name;
      card.querySelector(".manager-status").textContent = status.label;
      const list = card.querySelector(".overview-tasks");
      const tasks = form?.tasks || [];
      if (!tasks.length) {
        const li = document.createElement("li");
        li.className = "muted";
        li.textContent = "Пока пусто — менеджер ещё не создал чеклист";
        list.appendChild(li);
      } else {
        tasks.forEach((t) => {
          const li = document.createElement("li");
          li.className = t.done ? "done" : "";
          li.textContent = `${t.done ? "✓ " : "○ "}${t.text}`;
          list.appendChild(li);
        });
      }
      const stickersBox = card.querySelector(".overview-stickers");
      if (!stickers.length) {
        stickersBox.innerHTML = `<span class="muted">Стикеров нет</span>`;
      } else {
        stickers.slice(0, 4).forEach((s) => {
          const chip = document.createElement("span");
          chip.className = `sticker-chip sticker-${s.color || "cyan"}`;
          chip.textContent = s.text || "—";
          stickersBox.appendChild(chip);
        });
        if (stickers.length > 4) {
          const more = document.createElement("span");
          more.className = "muted";
          more.textContent = `+${stickers.length - 4}`;
          stickersBox.appendChild(more);
        }
      }
      card.querySelector("[data-open]").addEventListener("click", () => {
        state.selectedId = manager.id;
        renderAll();
        document.getElementById("workspace")?.scrollIntoView({ behavior: "smooth" });
      });
      el.overviewView.appendChild(card);
    });
  }

  function loadManagerWorkspace(managerId) {
    const manager = state.managers.find((m) => m.id === managerId);
    if (!manager) return;

    const form = todaysForm(managerId);
    el.checklistTitle.textContent = `Чеклист · ${manager.name}`;
    el.checklistSub.textContent = `${formatDateRu(state.today)} · форма «${manager.name} ${state.today}»`;
    el.boardTitle.textContent = "Основные приоритеты в работе";
    el.boardSub.textContent = `Доска · ${manager.name}`;

    if (form?.status === "completed") {
      state.phase = "evening";
      state.draftTasks = form.tasks.map((t) => ({ ...t }));
    } else if (form?.status === "morning") {
      state.phase = "evening";
      state.draftTasks = form.tasks.map((t) => ({ ...t, done: Boolean(t.done) }));
    } else {
      state.phase = "morning";
      state.draftTasks = [{ id: `local-${Date.now()}`, text: "", done: false }];
    }

    setPhase(state.phase);
    renderTasksEditor();
    renderStickers(managerId);
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
      el.formHint.textContent = completed
        ? "Этот день уже в архиве. Можно удалить запись ниже, если это тест."
        : "Создайте чеклист задач на день. В архив он попадёт только после вечерних галочек.";
    } else {
      el.saveFormBtn.textContent = completed ? "Уже в архиве" : "Сохранить и в архив";
      el.saveFormBtn.disabled = completed || !hasMorning;
      el.addTaskBtn.hidden = true;
      el.formHint.textContent = !hasMorning
        ? "Сначала сохраните утренний чеклист — затем проставьте галочки."
        : completed
          ? "Чеклист уже в архиве."
          : "Отметьте выполненное и нажмите сохранить — форма уйдёт в архив.";
    }
  }

  function renderTasksEditor() {
    el.tasksEditor.innerHTML = "";
    const evening = state.phase === "evening";
    const form = todaysForm(state.selectedId);
    const locked = form?.status === "completed";

    state.draftTasks.forEach((task, index) => {
      const row = document.createElement("div");
      row.className = "task-row";
      row.innerHTML = `
        <input type="checkbox" ${task.done ? "checked" : ""} ${evening && !locked ? "" : "disabled"} />
        <input type="text" class="task-input" placeholder="Задача на день" ${
          !evening && !locked ? "" : "readonly"
        } />
        <button type="button" class="task-remove" title="Удалить" ${
          !evening && !locked ? "" : "hidden"
        }>×</button>
      `;
      const textInput = row.querySelector(".task-input");
      const check = row.querySelector('input[type="checkbox"]');
      textInput.value = task.text || "";
      textInput.addEventListener("input", () => {
        task.text = textInput.value;
      });
      check.addEventListener("change", () => {
        task.done = check.checked;
      });
      row.querySelector(".task-remove").addEventListener("click", () => {
        state.draftTasks.splice(index, 1);
        if (!state.draftTasks.length) {
          state.draftTasks.push({ id: `local-${Date.now()}`, text: "", done: false });
        }
        renderTasksEditor();
      });
      el.tasksEditor.appendChild(row);
    });
  }

  function renderStickers(managerId) {
    [...el.stickerBoard.querySelectorAll(".sticker")].forEach((n) => n.remove());
    const stickers = state.boards[managerId] || [];
    stickers.forEach((sticker) => {
      el.stickerBoard.appendChild(createStickerNode(managerId, sticker));
    });
  }

  function createStickerNode(managerId, sticker, { startEditing = false } = {}) {
    const node = document.createElement("div");
    node.className = `sticker sticker-${sticker.color || "cyan"}`;
    node.dataset.id = sticker.id;
    node.style.left = `${sticker.x}px`;
    node.style.top = `${sticker.y}px`;
    node.style.transform = `rotate(${sticker.rotation || 0}deg)`;
    node.innerHTML = `
      <textarea class="sticker-text" rows="4" spellcheck="false" readonly></textarea>
      <div class="sticker-tools">
        <button type="button" data-act="edit" title="Редактировать">✎</button>
        <button type="button" data-act="color" title="Цвет">◐</button>
        <button type="button" data-act="del" title="Удалить">×</button>
      </div>
    `;
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
      if (e.key === "Escape") {
        textEl.value = sticker.text || "";
        textEl.blur();
      }
    });

    node.querySelector('[data-act="edit"]').addEventListener("click", (e) => {
      e.stopPropagation();
      startEdit();
    });
    node.querySelector('[data-act="color"]').addEventListener("click", (e) => {
      e.stopPropagation();
      const idx = COLORS.indexOf(sticker.color);
      sticker.color = COLORS[(idx + 1) % COLORS.length];
      node.className = `sticker sticker-${sticker.color}${
        node.classList.contains("is-editing") ? " is-editing" : ""
      }`;
      queueStickersSave(managerId);
    });
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
          .map((t) => ({ id: t.id, text: String(t.text || "").trim(), done: false }))
          .filter((t) => t.text);
        if (!tasks.length) {
          el.formHint.textContent = "Введите текст задачи — пустые строки не сохраняются.";
          toast("Добавьте хотя бы одну задачу");
          el.tasksEditor.querySelector(".task-input")?.focus();
          return;
        }
        const data = await api("/api/forms/morning", {
          method: "POST",
          body: JSON.stringify({ managerId, date: state.today, tasks }),
        });
        upsertForm(data.form);
        toast(data.message || "Утро сохранено");
        state.phase = "evening";
        state.draftTasks = data.form.tasks.map((t) => ({ ...t }));
        renderAll();
      } else {
        const tasks = state.draftTasks.map((t) => ({
          id: t.id,
          text: String(t.text || "").trim(),
          done: Boolean(t.done),
        }));
        const data = await api("/api/forms/evening", {
          method: "POST",
          body: JSON.stringify({ managerId, date: state.today, tasks }),
        });
        upsertForm(data.form);
        toast(data.message || "В архиве");
        renderAll();
      }
    } catch (err) {
      toast(err.message);
    }
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
      const empty = document.createElement("div");
      empty.className = "archive-empty";
      empty.textContent =
        "Пока пусто. Чеклист менеджера попадает сюда после вечерних галочек и «Сохранить».";
      el.archiveList.appendChild(empty);
      return;
    }

    archive.forEach((form) => {
      const item = document.createElement("article");
      item.className = "archive-item";
      const doneCount = form.tasks.filter((t) => t.done).length;
      item.innerHTML = `
        <div>
          <h3></h3>
          <ul></ul>
        </div>
        <div class="archive-side">
          <span class="archive-badge"></span>
          <button type="button" class="btn btn-ghost btn-sm btn-danger" data-del>Удалить</button>
        </div>
      `;
      item.querySelector("h3").textContent = form.title;
      item.querySelector(".archive-badge").textContent = `${doneCount}/${form.tasks.length} выполнено`;
      const ul = item.querySelector("ul");
      form.tasks.forEach((t) => {
        const li = document.createElement("li");
        if (t.done) li.className = "done";
        li.textContent = t.text;
        ul.appendChild(li);
      });
      item.querySelector("[data-del]").addEventListener("click", async () => {
        if (!confirm(`Удалить «${form.title}» из архива?`)) return;
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

  el.addManagerBtn.addEventListener("click", async () => {
    const name = await askPrompt({
      title: "Новый менеджер",
      label: "Имя сотрудника",
      placeholder: "Например, Мария",
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

  el.addSticker.addEventListener("click", async () => {
    const managerId = state.selectedId;
    if (!managerId || managerId === "all") return;
    try {
      const color = COLORS[Math.floor(Math.random() * COLORS.length)];
      const data = await api(`/api/boards/${managerId}/stickers`, {
        method: "POST",
        body: JSON.stringify({
          text: "Новый приоритет",
          color,
          x: 36 + Math.random() * 140,
          y: 36 + Math.random() * 100,
        }),
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
    const form = todaysForm(state.selectedId);
    if (form?.status === "completed") {
      state.draftTasks = form.tasks.map((t) => ({ ...t }));
    } else if (form?.status === "morning") {
      state.draftTasks = form.tasks.map((t) => ({ ...t, done: false }));
    }
    setPhase("morning");
    renderTasksEditor();
  });

  el.tabEvening.addEventListener("click", () => {
    const form = todaysForm(state.selectedId);
    if (form && (form.status === "morning" || form.status === "completed")) {
      state.draftTasks = form.tasks.map((t) => ({ ...t, done: Boolean(t.done) }));
    }
    setPhase("evening");
    renderTasksEditor();
  });

  el.addTaskBtn.addEventListener("click", () => {
    state.draftTasks.push({ id: `local-${Date.now()}`, text: "", done: false });
    renderTasksEditor();
    const inputs = el.tasksEditor.querySelectorAll(".task-input");
    inputs[inputs.length - 1]?.focus();
  });

  el.saveFormBtn.addEventListener("click", saveForm);

  el.promptForm.addEventListener("click", (e) => {
    const btn = e.target.closest("button[value]");
    if (btn) el.promptModal.returnValue = btn.value;
  });

  loadState().catch((err) => toast(err.message || "Не удалось загрузить данные"));
})();
