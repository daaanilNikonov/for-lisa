(() => {
  const STORAGE_KEY = "forus-quiz-scenario-content-v2";
  const LEAD_KEY = "forus-quiz-leads-v1";
  const data = window.QUIZ_SCRIPTS;
  const isAdmin = new URLSearchParams(location.search).has("admin");

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
    btnShowLead: document.getElementById("btnShowLead"),
    btnCopyLink: document.getElementById("btnCopyLink"),
    btnDownloadShot: document.getElementById("btnDownloadShot"),
    btnExportLeads: document.getElementById("btnExportLeads"),
    shareBox: document.getElementById("shareBox"),
    sharePreview: document.getElementById("sharePreview"),
    shareLink: document.getElementById("shareLink"),
    shareStatus: document.getElementById("shareStatus"),
    shareCard: document.getElementById("shareCard"),
    leadPanel: document.getElementById("leadPanel"),
    resultSection: document.getElementById("result"),
    quizSection: document.getElementById("quiz"),
    resultTitle: document.getElementById("resultTitle"),
    resultText: document.getElementById("resultText"),
    resultImage: document.getElementById("resultImage"),
    resultImagePlaceholder: document.getElementById("resultImagePlaceholder"),
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

  /** Картинка вида: override → явный imageSrc → assets/results/{id}.jpg/.jpeg/.png → пусто */
  function resolveImageSrc(sc) {
    if (sc.imageSrc) return sc.imageSrc;
    return `assets/results/${sc.id}.jpg`;
  }

  function loadOverrides() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw);
      data.scenarios.forEach((sc) => {
        const o = parsed[sc.id];
        if (!o) return;
        if (typeof o.text === "string" && o.text) sc.text = o.text;
        if (typeof o.imagePrompt === "string") sc.imagePrompt = o.imagePrompt;
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
        imagePrompt: sc.imagePrompt || "",
        imageSrc: sc.imageSrc || "",
      };
    });
    return payload;
  }

  function saveOverrides() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(collectEditorPayload()));
    if (els.editorStatus) els.editorStatus.textContent = "Сохранено в браузере";
  }

  function isQuestionAnswered(q) {
    if (!q) return false;
    if (q.id === "kedo_vendor") {
      return Boolean(state.answers.kedo_vendor || state.answers.kedo_vendor_custom);
    }
    return Boolean(state.answers[q.id]);
  }

  function landingUrl() {
    if (location.protocol.startsWith("http") && !location.hostname.includes("htmlpreview")) {
      return `${location.origin}${location.pathname.replace(/\/[^/]*$/, "/")}`;
    }
    return data.shareBaseUrl;
  }

  function kedoLabel() {
    if (state.answers.docs !== "kedo") return "";
    if (state.answers.kedo_vendor_custom) return state.answers.kedo_vendor_custom;
    const id = state.answers.kedo_vendor;
    const opt = data.kedoVendorQuestion.options.find((o) => o.id === id);
    return opt?.label || id || "";
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

    const customBlock = q.allowCustom
      ? `<div class="custom-field">
            <label>
              Свой вариант
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
        if (q.id === "kedo_vendor") {
          // выбор из списка — можно очистить кастом, если хотят только список
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
    const title = window.QUIZ_MATCH.formatScenarioTitle(
      scenario,
      state.answers.role
    );
    state.lastScenario = { ...scenario, displayName: title };

    els.progressBar.style.width = "100%";
    els.quizSection.classList.add("is-hidden");
    els.resultSection.classList.remove("is-hidden");
    els.shareBox.classList.add("is-hidden");
    els.leadPanel.classList.add("is-hidden");
    state.shareDataUrl = "";
    els.sharePreview.hidden = true;

    els.resultTitle.textContent = title;
    els.resultText.textContent = scenario.text || "";

    const src = resolveImageSrc(scenario);
    els.resultImage.hidden = false;
    els.resultImagePlaceholder.hidden = true;
    els.resultImage.alt = title;
    els.resultImage.onerror = () => {
      els.resultImage.hidden = true;
      els.resultImagePlaceholder.hidden = false;
    };
    els.resultImage.onload = () => {
      els.resultImage.hidden = false;
      els.resultImagePlaceholder.hidden = true;
    };
    els.resultImage.src = src;

    els.shareLink.value = landingUrl();
    els.resultSection.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function restart() {
    state.step = 0;
    state.answers = {};
    state.lastScenario = null;
    els.resultSection.classList.add("is-hidden");
    els.leadPanel.classList.add("is-hidden");
    els.quizSection.classList.remove("is-hidden");
    renderQuestion();
    els.quizSection.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  async function makeScreenshot() {
    if (typeof html2canvas !== "function") {
      throw new Error("Библиотека скриншота не загрузилась");
    }
    return (
      await html2canvas(els.shareCard, {
        backgroundColor: "#ffffff",
        scale: 2,
        useCORS: true,
      })
    ).toDataURL("image/png");
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
      els.shareStatus.textContent = "Скриншот готов";
    } catch (err) {
      console.warn(err);
      els.shareStatus.textContent =
        "Скриншот не удалось создать — ссылку всё равно можно скопировать.";
    }
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
    if (!els.editorList) return;
    els.editorList.innerHTML = data.scenarios
      .map((sc) => {
        const preview = sc.imageSrc || `assets/results/${sc.id}.jpg`;
        return `
      <article class="editor-card" data-id="${sc.id}">
        <img class="editor-preview" src="${preview}" alt="${sc.shortName}" onerror="this.style.opacity=0.25" />
        <div class="editor-fields">
          <h3>${sc.titlePattern.replace("{role}", "…")}</h3>
          <p class="hint">id: <code>${sc.id}</code> · файл: assets/results/${sc.id}.jpg</p>
          <label>
            Загрузить картинку для этого вида
            <input type="file" accept="image/*" data-field="image" />
          </label>
        </div>
      </article>`;
      })
      .join("");

    els.editorList.querySelectorAll(".editor-card").forEach((card) => {
      const id = card.dataset.id;
      const scenario = data.scenarios.find((s) => s.id === id);
      const preview = card.querySelector(".editor-preview");
      const fileInput = card.querySelector('input[data-field="image"]');

      fileInput.addEventListener("change", async () => {
        const file = fileInput.files?.[0];
        if (!file || !scenario) return;
        const url = await fileToDataUrl(file);
        scenario.imageSrc = url;
        preview.src = url;
        preview.style.opacity = 1;
        els.editorStatus.textContent = `Загружено: ${scenario.shortName}`;
        saveOverrides();
      });
    });
  }

  function csvEscape(value) {
    const s = String(value ?? "");
    if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
    return s;
  }

  function leadsToCsv(rows) {
    const header = "created_at,name,phone,inn,scenario_id,scenario_name,kedo,source";
    const lines = rows.map((r) =>
      [
        r.createdAt,
        r.name,
        r.phone,
        r.inn,
        r.scenarioId,
        r.scenarioName,
        r.kedo,
        r.source || "landing",
      ]
        .map(csvEscape)
        .join(",")
    );
    return [header, ...lines].join("\n");
  }

  function loadLeads() {
    try {
      return JSON.parse(localStorage.getItem(LEAD_KEY) || "[]");
    } catch {
      return [];
    }
  }

  function submitToYandexForm(payload) {
    const cfg = data.yandexForm || {};
    if (!cfg.url) return false;

    const form = document.createElement("form");
    form.method = "POST";
    form.action = cfg.url;
    form.target = "yandexFormFrame";
    form.acceptCharset = "UTF-8";
    form.style.display = "none";

    const map = cfg.fields || {};
    const values = {
      [map.name || "name"]: payload.name,
      [map.phone || "phone"]: payload.phone,
      [map.inn || "inn"]: payload.inn,
      [map.scenario || "scenario"]: payload.scenarioName || "",
    };

    Object.entries(values).forEach(([name, value]) => {
      const input = document.createElement("input");
      input.type = "hidden";
      input.name = name;
      input.value = value;
      form.appendChild(input);
    });

    document.body.appendChild(form);
    form.submit();
    form.remove();
    return true;
  }

  /** Строка для дозаписи в leads/submissions.csv (репозиторий) */
  function buildRepoCsvRow(payload) {
    return [
      payload.createdAt,
      payload.name,
      payload.phone,
      payload.inn,
      payload.scenarioId,
      payload.scenarioName,
      payload.kedo,
      "landing",
    ]
      .map(csvEscape)
      .join(",");
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
    state.step += 1;
    if (state.step >= queue().length) {
      showResult();
      return;
    }
    renderQuestion();
  });

  els.btnRestart.addEventListener("click", restart);
  els.btnShare.addEventListener("click", shareResults);
  els.btnShowLead.addEventListener("click", () => {
    els.leadPanel.classList.remove("is-hidden");
    els.leadPanel.scrollIntoView({ behavior: "smooth", block: "start" });
  });

  els.btnCopyLink.addEventListener("click", async () => {
    try {
      await navigator.clipboard.writeText(els.shareLink.value);
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
      a.download = `forus-test-${state.lastScenario?.id || "result"}.png`;
      a.click();
      els.shareStatus.textContent = "Скриншот скачан";
    } catch {
      els.shareStatus.textContent = "Не удалось скачать скриншот";
    }
  });

  if (els.btnSaveEditor) els.btnSaveEditor.addEventListener("click", saveOverrides);
  if (els.btnExportEditor) {
    els.btnExportEditor.addEventListener("click", () => {
      const blob = new Blob([JSON.stringify(collectEditorPayload(), null, 2)], {
        type: "application/json",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "forus-quiz-images.json";
      a.click();
      URL.revokeObjectURL(url);
      els.editorStatus.textContent = "JSON скачан";
    });
  }
  if (els.btnResetEditor) {
    els.btnResetEditor.addEventListener("click", () => {
      localStorage.removeItem(STORAGE_KEY);
      location.reload();
    });
  }
  if (els.btnExportLeads) {
    els.btnExportLeads.addEventListener("click", () => {
      const csv = leadsToCsv(loadLeads());
      const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "submissions.csv";
      a.click();
      URL.revokeObjectURL(url);
      els.editorStatus.textContent =
        "CSV скачан — допишите в test-kabinet/leads/submissions.csv и закоммитьте";
    });
  }

  els.leadForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const payload = {
      name: document.getElementById("leadName").value.trim(),
      phone: document.getElementById("leadPhone").value.trim(),
      inn: document.getElementById("leadInn").value.trim(),
      scenarioId: state.lastScenario?.id || "",
      scenarioName: state.lastScenario?.displayName || "",
      kedo: kedoLabel(),
      createdAt: new Date().toISOString(),
      source: "landing",
      repoCsvRow: "",
    };

    if (!/^\d{10}$|^\d{12}$/.test(payload.inn)) {
      els.leadStatus.textContent = "Проверьте ИНН: нужно 10 или 12 цифр";
      return;
    }

    payload.repoCsvRow = buildRepoCsvRow(payload);

    const existing = loadLeads();
    existing.push(payload);
    localStorage.setItem(LEAD_KEY, JSON.stringify(existing));

    const sent = submitToYandexForm(payload);
    if (sent) {
      els.leadStatus.textContent =
        "Заявка отправлена в Яндекс.Форму и сохранена. Мы свяжемся с вами.";
    } else {
      els.leadStatus.textContent =
        "Заявка сохранена. Укажите URL Яндекс.Формы в настройках, чтобы отправка шла туда автоматически.";
    }

    // Строка для таблицы репозитория — копируем в буфер
    navigator.clipboard?.writeText(payload.repoCsvRow).catch(() => {});

    els.leadForm.reset();
  });

  // Admin UI
  if (isAdmin) {
    document.querySelectorAll(".admin-only").forEach((el) => {
      el.classList.remove("is-hidden");
    });
    renderEditor();
  }

  loadOverrides();
  renderQuestion();
})();
