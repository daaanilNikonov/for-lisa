(() => {
  const STORAGE_KEY = "forus-quiz-scenario-content-v1";
  const data = window.QUIZ_SCRIPTS;
  const labels = window.QUIZ_PROFILE_LABELS;

  const state = {
    step: 0,
    answers: {},
  };

  const els = {
    quizCard: document.getElementById("quizCard"),
    quizStep: document.getElementById("quizStep"),
    progressBar: document.getElementById("progressBar"),
    btnBack: document.getElementById("btnBack"),
    btnNext: document.getElementById("btnNext"),
    btnRestart: document.getElementById("btnRestart"),
    resultSection: document.getElementById("result"),
    quizSection: document.getElementById("quiz"),
    resultTitle: document.getElementById("resultTitle"),
    resultMeta: document.getElementById("resultMeta"),
    resultText: document.getElementById("resultText"),
    resultNote: document.getElementById("resultNote"),
    resultImage: document.getElementById("resultImage"),
    resultImagePlaceholder: document.getElementById("resultImagePlaceholder"),
    resultTriggers: document.getElementById("resultTriggers"),
    resultProfile: document.getElementById("resultProfile"),
    scenariosGrid: document.getElementById("scenariosGrid"),
    editorList: document.getElementById("editorList"),
    editorStatus: document.getElementById("editorStatus"),
    btnSaveEditor: document.getElementById("btnSaveEditor"),
    btnExportEditor: document.getElementById("btnExportEditor"),
    btnResetEditor: document.getElementById("btnResetEditor"),
  };

  /** Назначить временные картинки по кругу, если imageSrc пустой */
  function assignTempImages() {
    const pool = data.tempImages || [];
    data.scenarios.forEach((sc, i) => {
      if (!sc.imageSrc && pool.length) {
        sc.imageSrc = pool[i % pool.length];
      }
    });
  }

  function loadOverrides() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw);
      data.scenarios.forEach((sc) => {
        const o = parsed[sc.id];
        if (!o) return;
        if (typeof o.text === "string") sc.text = o.text;
        if (typeof o.imageSrc === "string" && o.imageSrc) sc.imageSrc = o.imageSrc;
      });
    } catch (err) {
      console.warn("Не удалось прочитать сохранённый контент", err);
    }
  }

  function collectEditorPayload() {
    const payload = {};
    data.scenarios.forEach((sc) => {
      payload[sc.id] = {
        name: sc.name,
        text: sc.text || "",
        imageSrc: sc.imageSrc || "",
      };
    });
    return payload;
  }

  function saveOverrides() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(collectEditorPayload()));
    els.editorStatus.textContent = "Сохранено в браузере";
    renderScenariosGrid();
  }

  function getQuestion(index) {
    return data.questions[index];
  }

  function activeTriggers(answers) {
    const list = [];
    if (answers.certs === "daily") list.push("Триггер напряжения: справки ежедневно");
    if (answers.overtime === "always") list.push("Триггер напряжения: почти всегда задерживается");
    if (answers.report === "last_day") list.push("Триггер напряжения: отчётность в последний день");
    if (answers.auto === "high") list.push("Триггер спокойствия: автоматизация >80%");
    if (answers.docs === "kedo") list.push("Триггер спокойствия: есть КЭДО");
    if (answers.report === "now") list.push("Триггер спокойствия: отчётность сразу");
    if (answers.certs === "rare" || answers.certs === "never") {
      list.push("Триггер спокойствия: справки редко / не обращаются");
    }
    return list;
  }

  function renderQuestion() {
    const q = getQuestion(state.step);
    const total = data.questions.length;
    const selected = state.answers[q.id];

    els.quizStep.textContent = `Вопрос ${q.number} из ${total}`;
    els.progressBar.style.width = `${(state.step / total) * 100}%`;
    els.btnBack.disabled = state.step === 0;
    els.btnNext.disabled = !selected;
    els.btnNext.textContent = state.step === total - 1 ? "Показать результат" : "Далее";

    els.quizCard.innerHTML = `
      <h2>${q.title}</h2>
      <div class="options" role="listbox" aria-label="${q.title}">
        ${q.options
          .map(
            (opt) => `
          <button
            type="button"
            class="option ${selected === opt.id ? "is-selected" : ""}"
            data-option="${opt.id}"
            role="option"
            aria-selected="${selected === opt.id}"
          >
            <span class="option-bullet" aria-hidden="true"></span>
            <span class="option-label">${opt.label}</span>
          </button>`
          )
          .join("")}
      </div>
    `;

    els.quizCard.querySelectorAll(".option").forEach((btn) => {
      btn.addEventListener("click", () => {
        state.answers[q.id] = btn.dataset.option;
        renderQuestion();
      });
    });
  }

  function showResult() {
    const scenario = window.QUIZ_MATCH.resolve(state.answers);
    els.progressBar.style.width = "100%";
    els.quizSection.classList.add("is-hidden");
    els.resultSection.classList.remove("is-hidden");

    els.resultTitle.textContent = scenario.name;
    els.resultMeta.textContent = `id: ${scenario.id} · приоритет ${scenario.priority}`;
    els.resultText.textContent = scenario.text || "";
    els.resultNote.textContent = scenario.triggersNote || "";

    if (scenario.imageSrc) {
      els.resultImage.hidden = false;
      els.resultImage.src = scenario.imageSrc;
      els.resultImage.alt = scenario.name;
      els.resultImagePlaceholder.hidden = true;
    } else {
      els.resultImage.hidden = true;
      els.resultImage.removeAttribute("src");
      els.resultImagePlaceholder.hidden = false;
    }

    const triggers = activeTriggers(state.answers);
    els.resultTriggers.innerHTML = triggers.length
      ? triggers.map((t) => `<li>${t}</li>`).join("")
      : "<li>Явных триггеров нет — запасной сценарий</li>";

    const keys = ["role", "size", "docs", "certs", "auto", "overtime", "report"];
    els.resultProfile.innerHTML = keys
      .map((key) => {
        const value = state.answers[key];
        const label = labels[key]?.[value] || value;
        return `<li>${label}</li>`;
      })
      .join("");

    els.resultSection.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function restart() {
    state.step = 0;
    state.answers = {};
    els.resultSection.classList.add("is-hidden");
    els.quizSection.classList.remove("is-hidden");
    renderQuestion();
    els.quizSection.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function renderScenariosGrid() {
    els.scenariosGrid.innerHTML = data.scenarios
      .map(
        (sc) => `
      <article class="scenario-card" id="scenario-${sc.id}">
        <h3>${sc.name}</h3>
        <div class="scenario-id">${sc.id}</div>
        ${
          sc.imageSrc
            ? `<img class="scenario-thumb" src="${sc.imageSrc}" alt="${sc.name}" />`
            : `<div class="scenario-thumb" role="img" aria-label="Нет картинки"></div>`
        }
        <p class="scenario-triggers">${sc.triggersNote}</p>
      </article>`
      )
      .join("");
  }

  function fileToDataUrl(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  function renderEditor() {
    els.editorList.innerHTML = data.scenarios
      .map(
        (sc) => `
      <article class="editor-card" data-id="${sc.id}">
        <img class="editor-preview" src="${sc.imageSrc || ""}" alt="${sc.name}" />
        <div class="editor-fields">
          <h3>${sc.name}</h3>
          <p class="hint">${sc.triggersNote}</p>
          <label>
            Картинка сценария
            <input type="file" accept="image/*" data-field="image" />
          </label>
          <label>
            Текст результата
            <textarea data-field="text" placeholder="Сюда вставьте текст для этого сценария">${
              sc.text || ""
            }</textarea>
          </label>
        </div>
      </article>`
      )
      .join("");

    els.editorList.querySelectorAll(".editor-card").forEach((card) => {
      const id = card.dataset.id;
      const scenario = data.scenarios.find((s) => s.id === id);
      const preview = card.querySelector(".editor-preview");
      const fileInput = card.querySelector('input[data-field="image"]');
      const textArea = card.querySelector('textarea[data-field="text"]');

      fileInput.addEventListener("change", async () => {
        const file = fileInput.files?.[0];
        if (!file || !scenario) return;
        const url = await fileToDataUrl(file);
        scenario.imageSrc = url;
        preview.src = url;
        els.editorStatus.textContent = `Картинка выбрана: ${scenario.shortName}`;
      });

      textArea.addEventListener("input", () => {
        if (!scenario) return;
        scenario.text = textArea.value;
      });
    });
  }

  els.btnBack.addEventListener("click", () => {
    if (state.step === 0) return;
    state.step -= 1;
    renderQuestion();
  });

  els.btnNext.addEventListener("click", () => {
    const q = getQuestion(state.step);
    if (!state.answers[q.id]) return;
    if (state.step === data.questions.length - 1) {
      showResult();
      return;
    }
    state.step += 1;
    renderQuestion();
  });

  els.btnRestart.addEventListener("click", restart);

  els.btnSaveEditor.addEventListener("click", saveOverrides);

  els.btnExportEditor.addEventListener("click", () => {
    const blob = new Blob([JSON.stringify(collectEditorPayload(), null, 2)], {
      type: "application/json",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "forus-quiz-scenarios.json";
    a.click();
    URL.revokeObjectURL(url);
    els.editorStatus.textContent = "JSON скачан";
  });

  els.btnResetEditor.addEventListener("click", () => {
    localStorage.removeItem(STORAGE_KEY);
    location.reload();
  });

  assignTempImages();
  loadOverrides();
  renderQuestion();
  renderScenariosGrid();
  renderEditor();
})();
