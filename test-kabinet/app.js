(() => {
  const STORAGE_KEY = "forus-quiz-scenario-content-v1";
  const LEAD_KEY = "forus-quiz-leads-v1";
  const data = window.QUIZ_SCRIPTS;
  const labels = window.QUIZ_PROFILE_LABELS;

  const state = {
    step: 0,
    answers: {},
    lastScenario: null,
    shareDataUrl: "",
  };

  const els = {
    quizCard: document.getElementById("quizCard"),
    quizStep: document.getElementById("quizStep"),
    progressBar: document.getElementById("progressBar"),
    btnBack: document.getElementById("btnBack"),
    btnNext: document.getElementById("btnNext"),
    btnRestart: document.getElementById("btnRestart"),
    btnShare: document.getElementById("btnShare"),
    btnCopyLink: document.getElementById("btnCopyLink"),
    btnDownloadShot: document.getElementById("btnDownloadShot"),
    shareBox: document.getElementById("shareBox"),
    sharePreview: document.getElementById("sharePreview"),
    shareLink: document.getElementById("shareLink"),
    shareStatus: document.getElementById("shareStatus"),
    shareCard: document.getElementById("shareCard"),
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
    leadForm: document.getElementById("leadForm"),
    leadStatus: document.getElementById("leadStatus"),
  };

  function queue() {
    return window.QUIZ_MATCH.buildQuestionQueue(state.answers);
  }

  function currentQuestion() {
    return queue()[state.step];
  }

  function assignTempImages() {
    const pool = data.tempImages || [];
    data.scenarios.forEach((sc, i) => {
      if (!sc.imageSrc && pool.length) sc.imageSrc = pool[i % pool.length];
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
      console.warn(err);
    }
  }

  function collectEditorPayload() {
    const payload = {};
    data.scenarios.forEach((sc) => {
      payload[sc.id] = {
        shortName: sc.shortName,
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

  function isQuestionAnswered(q) {
    if (!q) return false;
    if (q.id === "kedo_vendor") {
      return Boolean(state.answers.kedo_vendor || state.answers.kedo_vendor_custom);
    }
    return Boolean(state.answers[q.id]);
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

  function landingUrl() {
    if (location.protocol.startsWith("http") && !location.hostname.includes("htmlpreview")) {
      return `${location.origin}${location.pathname.replace(/\/[^/]*$/, "/")}`;
    }
    return data.shareBaseUrl;
  }

  function renderQuestion() {
    const qList = queue();
    if (state.step >= qList.length) state.step = qList.length - 1;
    const q = qList[state.step];
    const total = qList.length;
    const selected = state.answers[q.id];

    els.quizStep.textContent = `Вопрос ${state.step + 1} из ${total}`;
    els.progressBar.style.width = `${(state.step / total) * 100}%`;
    els.btnBack.disabled = state.step === 0;
    els.btnNext.disabled = !isQuestionAnswered(q);
    els.btnNext.textContent =
      state.step === total - 1 ? "Показать результат" : "Далее";

    const customBlock =
      q.allowCustom
        ? `<div class="custom-field">
            <label>
              Окошко для вписания
              <input
                type="text"
                id="kedoCustomInput"
                placeholder="${q.customPlaceholder || "Впишите свой вариант"}"
                value="${state.answers.kedo_vendor_custom || ""}"
              />
            </label>
          </div>`
        : "";

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
      ${customBlock}
    `;

    els.quizCard.querySelectorAll(".option").forEach((btn) => {
      btn.addEventListener("click", () => {
        state.answers[q.id] = btn.dataset.option;
        if (q.id === "docs" && btn.dataset.option !== "kedo") {
          delete state.answers.kedo_vendor;
          delete state.answers.kedo_vendor_custom;
        }
        renderQuestion();
      });
    });

    const customInput = document.getElementById("kedoCustomInput");
    if (customInput) {
      customInput.addEventListener("input", () => {
        state.answers.kedo_vendor_custom = customInput.value.trim();
        els.btnNext.disabled = !isQuestionAnswered(q);
      });
    }
  }

  function showResult() {
    const scenario = window.QUIZ_MATCH.resolve(state.answers);
    const title = window.QUIZ_MATCH.formatScenarioTitle(scenario, state.answers.role);
    state.lastScenario = { ...scenario, displayName: title };

    els.progressBar.style.width = "100%";
    els.quizSection.classList.add("is-hidden");
    els.resultSection.classList.remove("is-hidden");
    els.shareBox.classList.add("is-hidden");
    state.shareDataUrl = "";
    els.sharePreview.hidden = true;

    els.resultTitle.textContent = title;
    els.resultMeta.textContent = `сценарий · ${scenario.id}`;
    els.resultText.textContent = scenario.text || "";
    els.resultNote.textContent = scenario.triggersNote || "";

    if (scenario.imageSrc) {
      els.resultImage.hidden = false;
      els.resultImage.src = scenario.imageSrc;
      els.resultImage.alt = title;
      els.resultImagePlaceholder.hidden = true;
    } else {
      els.resultImage.hidden = true;
      els.resultImage.removeAttribute("src");
      els.resultImagePlaceholder.hidden = false;
    }

    els.resultTriggers.innerHTML = activeTriggers(state.answers)
      .map((t) => `<li>${t}</li>`)
      .join("") || "<li>Явных триггеров нет — запасной сценарий</li>";

    const keys = [
      "role",
      "size",
      "docs",
      "kedo_vendor",
      "certs",
      "auto",
      "overtime",
      "report",
    ];
    els.resultProfile.innerHTML = keys
      .map((key) => {
        if (key === "kedo_vendor") {
          if (state.answers.docs !== "kedo") return "";
          const custom = state.answers.kedo_vendor_custom;
          const opt = state.answers.kedo_vendor;
          const label = custom
            ? `КЭДО: ${custom}`
            : labels.kedo_vendor?.[opt] || (opt ? `КЭДО: ${opt}` : "");
          return label ? `<li>${label}</li>` : "";
        }
        const value = state.answers[key];
        if (!value) return "";
        return `<li>${labels[key]?.[value] || value}</li>`;
      })
      .filter(Boolean)
      .join("");

    els.shareLink.value = landingUrl();
    els.resultSection.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function restart() {
    state.step = 0;
    state.answers = {};
    state.lastScenario = null;
    els.resultSection.classList.add("is-hidden");
    els.quizSection.classList.remove("is-hidden");
    renderQuestion();
    els.quizSection.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  async function makeScreenshot() {
    if (typeof html2canvas !== "function") {
      throw new Error("Библиотека скриншота не загрузилась");
    }
    const canvas = await html2canvas(els.shareCard, {
      backgroundColor: "#ffffff",
      scale: 2,
      useCORS: true,
    });
    return canvas.toDataURL("image/png");
  }

  async function shareResults() {
    els.shareStatus.textContent = "Готовим скриншот…";
    els.shareBox.classList.remove("is-hidden");
    els.shareLink.value = landingUrl();
    try {
      const dataUrl = await makeScreenshot();
      state.shareDataUrl = dataUrl;
      els.sharePreview.src = dataUrl;
      els.sharePreview.hidden = false;
      els.shareStatus.textContent = "Скриншот готов — скачайте или скопируйте ссылку";
      els.shareBox.scrollIntoView({ behavior: "smooth", block: "nearest" });
    } catch (err) {
      console.warn(err);
      els.shareStatus.textContent =
        "Не удалось сделать скриншот. Ссылку на лендинг всё равно можно скопировать.";
    }
  }

  function renderScenariosGrid() {
    const sampleRole = "kadrovik";
    els.scenariosGrid.innerHTML = data.scenarios
      .map((sc) => {
        const title = window.QUIZ_MATCH.formatScenarioTitle(sc, sampleRole);
        return `
      <article class="scenario-card" id="scenario-${sc.id}">
        <h3>${title}</h3>
        <div class="scenario-id">${sc.id} · пример для кадровика</div>
        ${
          sc.imageSrc
            ? `<img class="scenario-thumb" src="${sc.imageSrc}" alt="${title}" />`
            : `<div class="scenario-thumb" role="img" aria-label="Нет картинки"></div>`
        }
        <p class="scenario-triggers">${sc.triggersNote}</p>
      </article>`;
      })
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
        <img class="editor-preview" src="${sc.imageSrc || ""}" alt="${sc.shortName}" />
        <div class="editor-fields">
          <h3>${sc.titlePattern}</h3>
          <p class="hint">${sc.triggersNote}</p>
          <label>
            Картинка сценария
            <input type="file" accept="image/*" data-field="image" />
          </label>
          <label>
            Текст результата
            <textarea data-field="text" placeholder="Текст для этого сценария">${
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
        els.editorStatus.textContent = `Картинка: ${scenario.shortName}`;
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
    const q = currentQuestion();
    if (!isQuestionAnswered(q)) return;

    const qList = queue();
    if (state.step === qList.length - 1) {
      showResult();
      return;
    }

    // Recalculate queue after docs answer (may insert kedo question)
    state.step += 1;
    const nextList = queue();
    if (state.step >= nextList.length) {
      showResult();
      return;
    }
    renderQuestion();
  });

  els.btnRestart.addEventListener("click", restart);
  els.btnShare.addEventListener("click", shareResults);

  els.btnCopyLink.addEventListener("click", async () => {
    const value = els.shareLink.value;
    try {
      await navigator.clipboard.writeText(value);
      els.shareStatus.textContent = "Ссылка скопирована";
    } catch {
      els.shareLink.select();
      document.execCommand("copy");
      els.shareStatus.textContent = "Ссылка скопирована";
    }
  });

  els.btnDownloadShot.addEventListener("click", async () => {
    try {
      if (!state.shareDataUrl) state.shareDataUrl = await makeScreenshot();
      const a = document.createElement("a");
      a.href = state.shareDataUrl;
      const slug = (state.lastScenario?.id || "result").replace(/[^\w-]+/g, "-");
      a.download = `forus-test-${slug}.png`;
      a.click();
      els.shareStatus.textContent = "Скриншот скачан";
    } catch (err) {
      els.shareStatus.textContent = "Не удалось скачать скриншот";
    }
  });

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

  els.leadForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const payload = {
      name: document.getElementById("leadName").value.trim(),
      phone: document.getElementById("leadPhone").value.trim(),
      inn: document.getElementById("leadInn").value.trim(),
      scenarioId: state.lastScenario?.id || null,
      scenarioName: state.lastScenario?.displayName || null,
      answers: { ...state.answers },
      createdAt: new Date().toISOString(),
    };

    if (!/^\d{10}$|^\d{12}$/.test(payload.inn)) {
      els.leadStatus.textContent = "Проверьте ИНН: нужно 10 или 12 цифр";
      return;
    }

    const existing = JSON.parse(localStorage.getItem(LEAD_KEY) || "[]");
    existing.push(payload);
    localStorage.setItem(LEAD_KEY, JSON.stringify(existing));
    els.leadStatus.textContent =
      "Заявка сохранена. Мы свяжемся с вами и подскажем, как усилить процессы.";
    els.leadForm.reset();
  });

  assignTempImages();
  loadOverrides();
  renderQuestion();
  renderScenariosGrid();
  renderEditor();
})();
