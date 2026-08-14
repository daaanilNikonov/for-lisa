(() => {
  "use strict";

  const state = {
    today: null,
    managers: [],
    forms: [],
    stickers: [],
    checklist: { date: null, items: [] },
    activeManagerId: null,
    phase: "morning",
    draftTasks: [],
    stickerSaveTimer: null,
    checklistSaveTimer: null,
  };

  const el = {
    todayLabel: document.getElementById("todayLabel"),
    checklistList: document.getElementById("checklistList"),
    stickerBoard: document.getElementById("stickerBoard"),
    managersGrid: document.getElementById("managersGrid"),
    archiveList: document.getElementById("archiveList"),
    managerModal: document.getElementById("managerModal"),
    modalTitle: document.getElementById("modalTitle"),
    modalMeta: document.getElementById("modalMeta"),
    tasksEditor: document.getElementById("tasksEditor"),
    saveFormBtn: document.getElementById("saveFormBtn"),
    formHint: document.getElementById("formHint"),
    tabMorning: document.getElementById("tabMorning"),
    tabEvening: document.getElementById("tabEvening"),
    addTaskBtn: document.getElementById("addTaskBtn"),
    addManagerBtn: document.getElementById("addManagerBtn"),
    addChecklistItem: document.getElementById("addChecklistItem"),
    addSticker: document.getElementById("addSticker"),
    promptModal: document.getElementById("promptModal"),
    promptForm: document.getElementById("promptForm"),
    promptTitle: document.getElementById("promptTitle"),
    promptLabel: document.getElementById("promptLabel"),
    promptInput: document.getElementById("promptInput"),
    toast: document.getElementById("toast"),
  };

  const COLORS = ["cyan", "amber", "mint", "rose", "violet"];

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
    // ensure visible above <dialog>
    el.toast.style.zIndex = "2147483647";
    el.toast.classList.add("is-on");
    clearTimeout(toast._t);
    toast._t = setTimeout(() => el.toast.classList.remove("is-on"), 3200);
  }

  function formatDateRu(iso) {
    if (!iso) return "";
    const [y, m, d] = iso.split("-").map(Number);
    const date = new Date(y, m - 1, d);
    return date.toLocaleDateString("ru-RU", {
      day: "numeric",
      month: "long",
      year: "numeric",
    });
  }

  function todaysForm(managerId) {
    return state.forms.find(
      (f) => f.managerId === managerId && f.date === state.today
    );
  }

  function statusFor(managerId) {
    const form = todaysForm(managerId);
    if (!form) return { key: "idle", label: "Ещё не начато" };
    if (form.status === "completed") return { key: "done", label: "День закрыт" };
    if (form.status === "morning") return { key: "morning", label: "Задачи заданы" };
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
        if (el.promptModal.returnValue === "ok") {
          resolve(el.promptInput.value.trim());
        } else {
          resolve(null);
        }
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
    state.stickers = data.stickers || [];
    state.checklist = data.checklist || { items: [] };
    el.todayLabel.textContent = `Сегодня · ${formatDateRu(state.today)}`;
    renderAll();
  }

  function renderAll() {
    renderChecklist();
    renderStickers();
    renderManagers();
    renderArchive();
  }

  function renderChecklist() {
    el.checklistList.innerHTML = "";
    (state.checklist.items || []).forEach((item) => {
      const li = document.createElement("li");
      li.className = `checklist-item${item.done ? " is-done" : ""}`;
      li.innerHTML = `
        <input type="checkbox" ${item.done ? "checked" : ""} aria-label="Отметить пункт" />
        <span class="checklist-text" contenteditable="true" spellcheck="false"></span>
      `;
      li.querySelector(".checklist-text").textContent = item.text;
      const checkbox = li.querySelector('input[type="checkbox"]');
      checkbox.addEventListener("change", () => {
        item.done = checkbox.checked;
        li.classList.toggle("is-done", item.done);
        queueChecklistSave();
      });
      const textEl = li.querySelector(".checklist-text");
      textEl.addEventListener("blur", () => {
        const next = textEl.textContent.trim();
        if (next && next !== item.text) {
          item.text = next;
          queueChecklistSave();
        } else {
          textEl.textContent = item.text;
        }
      });
      textEl.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          textEl.blur();
        }
      });
      el.checklistList.appendChild(li);
    });
  }

  function queueChecklistSave() {
    clearTimeout(state.checklistSaveTimer);
    state.checklistSaveTimer = setTimeout(async () => {
      try {
        const data = await api("/api/checklist", {
          method: "PUT",
          body: JSON.stringify({ items: state.checklist.items }),
        });
        state.checklist = data.checklist;
      } catch (err) {
        toast(err.message);
      }
    }, 350);
  }

  function renderStickers() {
    [...el.stickerBoard.querySelectorAll(".sticker")].forEach((n) => n.remove());
    state.stickers.forEach((sticker) => {
      el.stickerBoard.appendChild(createStickerNode(sticker));
    });
  }

  function createStickerNode(sticker, { startEditing = false } = {}) {
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
      queueStickersSave();
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
      node.className = `sticker sticker-${sticker.color}${node.classList.contains("is-editing") ? " is-editing" : ""}`;
      queueStickersSave();
    });
    node.querySelector('[data-act="del"]').addEventListener("click", async (e) => {
      e.stopPropagation();
      try {
        await api(`/api/stickers/${sticker.id}`, { method: "DELETE" });
        state.stickers = state.stickers.filter((s) => s.id !== sticker.id);
        node.remove();
      } catch (err) {
        toast(err.message);
      }
    });

    enableDrag(node, sticker);
    if (startEditing) requestAnimationFrame(startEdit);
    return node;
  }

  function enableDrag(node, sticker) {
    let dragging = false;
    let startX = 0;
    let startY = 0;
    let origX = 0;
    let origY = 0;

    const onPointerDown = (e) => {
      if (e.target.closest(".sticker-tools") || node.classList.contains("is-editing")) {
        return;
      }
      dragging = true;
      node.setPointerCapture(e.pointerId);
      startX = e.clientX;
      startY = e.clientY;
      origX = sticker.x;
      origY = sticker.y;
      node.style.zIndex = "8";
    };

    const onPointerMove = (e) => {
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
    };

    const onPointerUp = () => {
      if (!dragging) return;
      dragging = false;
      node.style.zIndex = "";
      queueStickersSave();
    };

    node.addEventListener("pointerdown", onPointerDown);
    node.addEventListener("pointermove", onPointerMove);
    node.addEventListener("pointerup", onPointerUp);
    node.addEventListener("pointercancel", onPointerUp);
  }

  function queueStickersSave() {
    clearTimeout(state.stickerSaveTimer);
    state.stickerSaveTimer = setTimeout(async () => {
      try {
        const data = await api("/api/stickers", {
          method: "PUT",
          body: JSON.stringify({ stickers: state.stickers }),
        });
        state.stickers = data.stickers;
      } catch (err) {
        toast(err.message);
      }
    }, 280);
  }

  function renderManagers() {
    el.managersGrid.innerHTML = "";
    state.managers.forEach((manager) => {
      const form = todaysForm(manager.id);
      const status = statusFor(manager.id);
      const card = document.createElement("article");
      card.className = "manager-card";
      const previewTasks = (form?.tasks || []).slice(0, 3);
      card.innerHTML = `
        <h3 class="manager-name"></h3>
        <span class="manager-status ${
          status.key === "done" ? "is-done" : status.key === "morning" ? "is-morning" : ""
        }"></span>
        <ul class="manager-tasks-preview"></ul>
        <div class="manager-actions">
          <button type="button" class="btn btn-primary" data-act="open">Карта дня</button>
          <button type="button" class="btn btn-ghost" data-act="rename">Имя</button>
        </div>
      `;
      card.querySelector(".manager-name").textContent = manager.name;
      card.querySelector(".manager-status").textContent = status.label;
      const list = card.querySelector(".manager-tasks-preview");
      if (!previewTasks.length) {
        const li = document.createElement("li");
        li.textContent = "Задач пока нет";
        list.appendChild(li);
      } else {
        previewTasks.forEach((t) => {
          const li = document.createElement("li");
          li.textContent = `${t.done ? "✓ " : "• "}${t.text}`;
          list.appendChild(li);
        });
      }
      card.querySelector('[data-act="open"]').addEventListener("click", () => openManagerForm(manager.id));
      card.querySelector('[data-act="rename"]').addEventListener("click", () => renameManager(manager));
      el.managersGrid.appendChild(card);
    });
  }

  function renderArchive() {
    const archive = state.forms
      .filter((f) => f.status === "completed")
      .sort((a, b) => String(b.date).localeCompare(String(a.date)) || a.title.localeCompare(b.title, "ru"));

    el.archiveList.innerHTML = "";
    if (!archive.length) {
      const empty = document.createElement("div");
      empty.className = "archive-empty";
      empty.textContent =
        "Пока пусто. Форма попадает сюда только после вечерней отметки галочками.";
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
        <span class="archive-badge"></span>
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
      el.archiveList.appendChild(item);
    });
  }

  function openManagerForm(managerId) {
    const manager = state.managers.find((m) => m.id === managerId);
    if (!manager) return;
    state.activeManagerId = managerId;
    const form = todaysForm(managerId);
    const completed = form?.status === "completed";

    if (completed) {
      state.phase = "evening";
      state.draftTasks = (form.tasks || []).map((t) => ({ ...t }));
    } else if (form?.status === "morning") {
      state.phase = "evening";
      state.draftTasks = (form.tasks || []).map((t) => ({ ...t, done: Boolean(t.done) }));
    } else {
      state.phase = "morning";
      state.draftTasks = [{ id: `local-${Date.now()}`, text: "", done: false }];
    }

    el.modalTitle.textContent = `Карта дня · ${manager.name}`;
    el.modalMeta.textContent = `${formatDateRu(state.today)} · форма «${manager.name} ${state.today}»`;
    setPhase(state.phase, { completed });
    renderTasksEditor();
    el.managerModal.showModal();
  }

  function setPhase(phase, { completed = false } = {}) {
    state.phase = phase;
    el.tabMorning.classList.toggle("is-active", phase === "morning");
    el.tabEvening.classList.toggle("is-active", phase === "evening");

    const form = todaysForm(state.activeManagerId);
    const hasMorning = Boolean(form && (form.status === "morning" || form.status === "completed"));

    if (phase === "morning") {
      el.saveFormBtn.textContent = completed ? "День уже в архиве" : "Сохранить утро";
      el.saveFormBtn.disabled = completed;
      el.addTaskBtn.hidden = completed;
      el.formHint.textContent = completed
        ? "Эта форма уже сохранена в хранилище."
        : "Первое заполнение: пропишите задачи на день. В архив форма ещё не попадёт.";
    } else {
      el.saveFormBtn.textContent = completed ? "Уже сохранено" : "Закрыть день и в архив";
      el.saveFormBtn.disabled = completed || !hasMorning;
      el.addTaskBtn.hidden = true;
      el.formHint.textContent = !hasMorning
        ? "Сначала сохраните утренние задачи — затем можно проставить галочки."
        : completed
          ? "Форма уже в хранилище."
          : "Второе заполнение: отметьте выполненное. После сохранения форма попадёт в хранилище.";
    }
  }

  function renderTasksEditor() {
    el.tasksEditor.innerHTML = "";
    const evening = state.phase === "evening";
    const form = todaysForm(state.activeManagerId);
    const locked = form?.status === "completed";

    state.draftTasks.forEach((task, index) => {
      const row = document.createElement("div");
      row.className = "task-row";
      row.innerHTML = `
        <input type="checkbox" ${task.done ? "checked" : ""} ${evening && !locked ? "" : "disabled"} />
        <input type="text" placeholder="Задача на день" ${!evening && !locked ? "" : "readonly"} />
        <button type="button" class="task-remove" title="Удалить" ${!evening && !locked ? "" : "hidden"}>×</button>
      `;
      const textInput = row.querySelector('input[type="text"]');
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

  async function saveForm() {
    const managerId = state.activeManagerId;
    if (!managerId) return;

    try {
      if (state.phase === "morning") {
        const tasks = state.draftTasks
          .map((t) => ({ id: t.id, text: String(t.text || "").trim(), done: false }))
          .filter((t) => t.text);
        if (!tasks.length) {
          el.formHint.textContent = "Введите текст задачи — пустые строки не сохраняются.";
          toast("Добавьте хотя бы одну задачу");
          const first = el.tasksEditor.querySelector('input[type="text"]');
          first?.focus();
          return;
        }
        const data = await api("/api/forms/morning", {
          method: "POST",
          body: JSON.stringify({ managerId, date: state.today, tasks }),
        });
        upsertForm(data.form);
        toast(data.message || "Утро сохранено");
        el.managerModal.close();
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
        toast(data.message || "Форма в архиве");
        el.managerModal.close();
        renderAll();
      }
    } catch (err) {
      toast(err.message);
    }
  }

  function upsertForm(form) {
    const idx = state.forms.findIndex((f) => f.id === form.id);
    if (idx >= 0) state.forms[idx] = form;
    else state.forms.push(form);
  }

  async function renameManager(manager) {
    const name = await askPrompt({
      title: "Имя менеджера",
      label: "Как отображать в карте дня",
      initial: manager.name,
    });
    if (!name || name === manager.name) return;
    try {
      const data = await api(`/api/managers/${manager.id}`, {
        method: "PATCH",
        body: JSON.stringify({ name }),
      });
      manager.name = data.manager.name;
      renderManagers();
      renderArchive();
      toast("Имя обновлено");
    } catch (err) {
      toast(err.message);
    }
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
      renderManagers();
      toast(`Добавлен: ${data.manager.name}`);
    } catch (err) {
      toast(err.message);
    }
  });

  el.addChecklistItem.addEventListener("click", async () => {
    const text = await askPrompt({
      title: "Пункт чеклиста",
      label: "Что добавить в работу группы",
      placeholder: "Например, созвон по эстафете баз",
    });
    if (!text) return;
    try {
      const data = await api("/api/checklist/items", {
        method: "POST",
        body: JSON.stringify({ text }),
      });
      state.checklist = data.checklist;
      renderChecklist();
    } catch (err) {
      toast(err.message);
    }
  });

  el.addSticker.addEventListener("click", async () => {
    try {
      const color = COLORS[Math.floor(Math.random() * COLORS.length)];
      const data = await api("/api/stickers", {
        method: "POST",
        body: JSON.stringify({
          text: "Новый приоритет",
          color,
          x: 36 + Math.random() * 140,
          y: 36 + Math.random() * 100,
        }),
      });
      state.stickers.push(data.sticker);
      el.stickerBoard.appendChild(createStickerNode(data.sticker, { startEditing: true }));
    } catch (err) {
      toast(err.message);
    }
  });

  el.tabMorning.addEventListener("click", () => {
    const form = todaysForm(state.activeManagerId);
    if (form?.status === "completed") {
      state.draftTasks = form.tasks.map((t) => ({ ...t }));
    } else if (form?.status === "morning" && state.phase === "evening") {
      // keep draft checkboxes when switching back? reset to morning editable copy
      state.draftTasks = form.tasks.map((t) => ({ ...t, done: false }));
    }
    setPhase("morning", { completed: form?.status === "completed" });
    renderTasksEditor();
  });

  el.tabEvening.addEventListener("click", () => {
    const form = todaysForm(state.activeManagerId);
    if (!form || (form.status !== "morning" && form.status !== "completed")) {
      setPhase("evening");
      renderTasksEditor();
      return;
    }
    if (state.phase === "morning" && form.status === "morning") {
      state.draftTasks = form.tasks.map((t) => ({ ...t, done: Boolean(t.done) }));
    }
    setPhase("evening", { completed: form.status === "completed" });
    renderTasksEditor();
  });

  el.addTaskBtn.addEventListener("click", () => {
    state.draftTasks.push({ id: `local-${Date.now()}`, text: "", done: false });
    renderTasksEditor();
    const inputs = el.tasksEditor.querySelectorAll('input[type="text"]');
    inputs[inputs.length - 1]?.focus();
  });

  el.saveFormBtn.addEventListener("click", saveForm);

  el.promptForm.addEventListener("click", (e) => {
    const btn = e.target.closest("button[value]");
    if (btn) el.promptModal.returnValue = btn.value;
  });

  loadState().catch((err) => {
    toast(err.message || "Не удалось загрузить данные");
  });
})();
