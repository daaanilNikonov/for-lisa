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

  /** База ассетов для встраивания в Тильду / CDN (см. tilda/) */
  function assetUrl(path) {
    if (!path) return path;
    if (/^(https?:|data:|blob:)/i.test(path)) return path;
    const base = window.QUIZ_ASSET_BASE || "";
    return base + String(path).replace(/^\//, "");
  }

  /** Картинка вида: override → явный imageSrc → assets/results/{id}.jpg */
  function resolveImageSrc(sc) {
    if (sc.imageSrc) return assetUrl(sc.imageSrc);
    return assetUrl(`assets/results/${sc.id}.jpg`);
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
      if (els.resultImagePlaceholder) els.resultImagePlaceholder.hidden = true;
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

  function roleLabel() {
    const roleId = state.answers.role;
    if (!roleId) return data.roleTitles.drugoe || "Специалист";
    if (data.roleTitles[roleId]) return data.roleTitles[roleId];
    const q = data.questions.find((item) => item.id === "role");
    const opt = q?.options?.find((o) => o.id === roleId);
    return opt?.label || roleId;
  }

  /** Телефон для Яндекс.Формы: 11–12 цифр, обычно 79XXXXXXXXX */
  function normalizePhone(raw) {
    let digits = String(raw || "").replace(/\D+/g, "");
    if (digits.length === 11 && digits.startsWith("8")) {
      digits = `7${digits.slice(1)}`;
    }
    if (digits.length === 10) digits = `7${digits}`;
    return digits;
  }

  function yandexCfg() {
    return data.yandexForm || {};
  }

  function yandexFieldMap() {
    return (
      yandexCfg().fields || {
        name: "answer_short_text_9008986271269920",
        inn: "answer_short_text_9008986271288132",
        phone: "answer_short_text_9008986271302340",
        position: "answer_short_text_9008986271334516",
      }
    );
  }

  /** Текст КЭДО для Яндекс.Формы из ответов квиза */
  function kedoForYandex() {
    const vendor = kedoLabel();
    if (vendor) return vendor;
    const docs = state.answers.docs;
    if (docs === "kedo") return "есть КЭДО (система не указана)";
    if (docs === "paper") return "нет КЭДО (документы на бумаге)";
    if (docs === "email") return "нет КЭДО (email/мессенджеры)";
    if (docs === "none") return "нет кадровых документов";
    return "";
  }

  function buildYandexValues(payload) {
    const map = yandexFieldMap();
    const kedo = (payload.kedo || "").trim();
    const position = (payload.position || "").trim();
    const values = {
      [map.name]: payload.name,
      [map.inn]: payload.inn,
      [map.phone]: payload.phone,
    };
    if (map.kedo) {
      values[map.position] = position;
      values[map.kedo] = kedo;
    } else {
      // В форме пока нет отдельного поля КЭДО — пишем в «Должность»
      values[map.position] = kedo ? `${position} · КЭДО: ${kedo}` : position;
    }
    return values;
  }

  function yandexFormBaseUrl() {
    const cfg = yandexCfg();
    const id = cfg.surveyId || "6a83b54feb614605f98e10ee";
    return (cfg.url || `https://forms.yandex.ru/u/${id}`).replace(/\/?$/, "/");
  }

  /** URL с предзаполнением (в т.ч. должность из Q1) */
  function buildYandexPrefillUrl(payload) {
    const values = buildYandexValues(payload);
    const params = new URLSearchParams({ iframe: "1", ...values });
    return `${yandexFormBaseUrl()}?${params.toString()}`;
  }

  function currentLeadDraft() {
    return {
      name: document.getElementById("leadName")?.value.trim() || "",
      phone: normalizePhone(document.getElementById("leadPhone")?.value || ""),
      inn: document.getElementById("leadInn")?.value.trim() || "",
      position: roleLabel(),
      kedo: kedoForYandex(),
    };
  }

  function leadDraftValid(draft) {
    if (!draft.name) return "Укажите имя";
    if (!/^\d{10}$|^\d{12}$/.test(draft.inn)) return "Проверьте ИНН: нужно 10 или 12 цифр";
    if (draft.phone.length < 11 || draft.phone.length > 12) {
      return "Проверьте телефон: нужно 11–12 цифр, например +7 900 123-45-67";
    }
    return "";
  }

  /**
   * Яндекс.Формы нельзя заполнить fetch-ом с лендинга (CSRF/CORS/капча),
   * поэтому синхронизируем красивую форму лендинга в iframe и клик
   * по «Отправить заявку» реально нажимает кнопку Яндекс.Формы.
   */
  const yandexBridge = {
    ready: false,
    loading: false,
    lastUrl: "",
    buttonTop: 360,
    buttonLeft: 0,
  };

  function ensureYandexBridge() {
    if (document.getElementById("yandexHitbox")) return;
    const wrap = document.createElement("div");
    wrap.className = "lead-actions";
    wrap.innerHTML = `
      <p class="lead-position-hint" id="leadPositionHint"></p>
      <div class="yandex-hitbox" id="yandexHitbox" aria-label="Отправить заявку">
        <span class="yandex-hitbox-label" id="yandexHitboxLabel">Отправить заявку</span>
        <iframe class="yandex-hitbox-frame" id="yandexHitFrame" title="Яндекс Форма" tabindex="-1" scrolling="no"></iframe>
        <button type="button" class="btn btn-primary yandex-hitbox-blocker" id="yandexHitBlocker">
          Отправить заявку
        </button>
      </div>
    `;
    const oldBtn = els.leadForm.querySelector('button[type="submit"]');
    const status = document.getElementById("leadStatus");
    if (oldBtn) oldBtn.replaceWith(wrap);
    else if (status) els.leadForm.insertBefore(wrap, status);
    else els.leadForm.appendChild(wrap);

    const blocker = document.getElementById("yandexHitBlocker");
    blocker.addEventListener("click", () => {
      const err = leadDraftValid(currentLeadDraft());
      els.leadStatus.textContent = err || "Заполните поля — затем нажмите «Отправить заявку» ещё раз";
      if (!err) refreshYandexBridge(true);
    });

    window.addEventListener("message", onYandexMessage);
    window.addEventListener("resize", () => {
      if (document.getElementById("yandexHitFrame")) positionYandexFrame();
    });
  }

  function onYandexMessage(event) {
    if (!String(event.origin || "").includes("forms.yandex.ru")) return;
    let payload = event.data;
    if (typeof payload === "string") {
      try {
        payload = JSON.parse(payload);
      } catch {
        return;
      }
    }
    if (!payload || typeof payload !== "object") return;

    if (typeof payload["iframe-height"] === "number") {
      // Для этой 4-польной формы кнопка стабильно ~368px сверху при iframe=1.
      // Не сдвигаем crop по высоте контейнера — иначе промахиваемся мимо Submit.
      positionYandexFrame();
    }

    if (payload.message === "sent") {
      onYandexSent(payload.answer_key || "");
    }
  }

  function positionYandexFrame() {
    const frame = document.getElementById("yandexHitFrame");
    const hit = document.getElementById("yandexHitbox");
    if (!frame || !hit) return;
    // Кнопка Яндекса ~97×36 @ (12, 368). Масштабируем под размер CTA лендинга.
    const btnLeft = 12;
    const btnTop = 368;
    const btnW = 97;
    const btnH = 36;
    const scale = Math.max(hit.offsetWidth / btnW, hit.offsetHeight / btnH, 1.8);
    const scaledH = btnH * scale;
    const topPad = Math.max(0, (hit.offsetHeight - scaledH) / 2);
    frame.style.transform = `scale(${scale})`;
    frame.style.transformOrigin = "0 0";
    frame.style.left = `${-btnLeft * scale}px`;
    frame.style.top = `${topPad - btnTop * scale}px`;
    yandexBridge.buttonTop = btnTop;
    yandexBridge.buttonLeft = btnLeft;
    yandexBridge.scale = scale;
  }

  function setHitboxEnabled(enabled) {
    const blocker = document.getElementById("yandexHitBlocker");
    const label = document.getElementById("yandexHitboxLabel");
    const hit = document.getElementById("yandexHitbox");
    if (!blocker || !label) return;
    blocker.classList.toggle("is-hidden", enabled);
    label.textContent = "Отправить заявку";
    label.classList.toggle("is-ready", enabled);
    if (hit) hit.classList.toggle("is-ready", enabled);
    yandexBridge.ready = enabled;
    if (enabled) positionYandexFrame();
  }

  function refreshYandexBridge(force) {
    ensureYandexBridge();
    const hint = document.getElementById("leadPositionHint");
    const draft = currentLeadDraft();
    if (hint) {
      hint.textContent = draft.kedo
        ? `В заявку: ${draft.position} · КЭДО: ${draft.kedo}`
        : `В заявку: ${draft.position}`;
    }
    const err = leadDraftValid(draft);
    const frame = document.getElementById("yandexHitFrame");
    if (!frame) return;

    if (err) {
      setHitboxEnabled(false);
      return;
    }

    const url = buildYandexPrefillUrl(draft);
    if (!force && url === yandexBridge.lastUrl && yandexBridge.ready) return;

    yandexBridge.loading = true;
    yandexBridge.ready = false;
    setHitboxEnabled(false);
    els.leadStatus.textContent = "Подготавливаем отправку в Яндекс.Форму…";
    yandexBridge.lastUrl = url;
    positionYandexFrame();
    frame.onload = () => {
      yandexBridge.loading = false;
      // Дополнительно проталкиваем значения postMessage (на случай кэша query)
      const values = buildYandexValues(draft);
      Object.entries(values).forEach(([slug, value]) => {
        frame.contentWindow?.postMessage(
          { message: "set-question-value", slug, value },
          "https://forms.yandex.ru"
        );
      });
      setTimeout(() => {
        setHitboxEnabled(true);
        document.getElementById("yandexHitbox")?.scrollIntoView({
          block: "center",
          behavior: "smooth",
        });
        els.leadStatus.textContent =
          "Должность и контакты подставлены. Нажмите «Отправить» — заявка уйдёт в Яндекс.Форму.";
      }, 900);
    };
    frame.src = url;
  }

  function onYandexSent(answerKey) {
    const draft = currentLeadDraft();
    const payload = {
      ...draft,
      roleId: state.answers.role || "",
      scenarioId: state.lastScenario?.id || "",
      scenarioName: state.lastScenario?.displayName || "",
      kedo: kedoLabel(),
      createdAt: new Date().toISOString(),
      source: "landing",
      yandexAnswerKey: answerKey || "",
      yandexOk: true,
      repoCsvRow: "",
    };
    payload.repoCsvRow = buildRepoCsvRow(payload);
    const existing = loadLeads();
    existing.push(payload);
    localStorage.setItem(LEAD_KEY, JSON.stringify(existing));
    navigator.clipboard?.writeText(payload.repoCsvRow).catch(() => {});

    els.leadStatus.textContent =
      "Заявка отправлена в Яндекс.Форму. Мы свяжемся с вами.";
    els.leadForm.reset();
    yandexBridge.lastUrl = "";
    yandexBridge.ready = false;
    setHitboxEnabled(false);
    const hint = document.getElementById("leadPositionHint");
    if (hint) {
      const kedo = kedoForYandex();
      hint.textContent = kedo
        ? `В заявку: ${roleLabel()} · КЭДО: ${kedo}`
        : `В заявку: ${roleLabel()}`;
    }
  }

  /** Опциональный серверный прокси (если открыт proxy/server.mjs и нет капчи) */
  async function submitToYandexProxy(payload) {
    const cfg = yandexCfg();
    const endpoint = cfg.proxyUrl;
    if (!endpoint) return { ok: false, reason: "no-proxy" };
    try {
      const res = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          surveyId: cfg.surveyId || "6a83b54feb614605f98e10ee",
          values: buildYandexValues(payload),
        }),
      });
      const json = await res.json().catch(() => ({}));
      if (!res.ok || json.ok === false || !json.result?.answer_id) {
        return { ok: false, reason: "proxy-error", detail: json.error || res.status };
      }
      return { ok: true, result: json.result };
    } catch (err) {
      return { ok: false, reason: "network", detail: String(err.message || err) };
    }
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
  // lead panel open handled with Yandex bridge below

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

  els.leadForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const draft = currentLeadDraft();
    const err = leadDraftValid(draft);
    if (err) {
      els.leadStatus.textContent = err;
      setHitboxEnabled(false);
      return;
    }

    // Сначала пробуем прокси (если доступен и отвечает answer_id)
    els.leadStatus.textContent = "Отправляем заявку…";
    const viaProxy = await submitToYandexProxy(draft);
    if (viaProxy.ok) {
      onYandexSent(viaProxy.result?.answer_key || "");
      return;
    }

    // Основной путь: синхронизация в iframe Яндекс.Формы + клик по её Submit
    refreshYandexBridge(true);
  });

  ["leadName", "leadPhone", "leadInn"].forEach((id) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener("input", () => {
      const err = leadDraftValid(currentLeadDraft());
      if (!err) refreshYandexBridge(false);
      else setHitboxEnabled(false);
    });
  });

  const _showLead = () => {
    ensureYandexBridge();
    const hint = document.getElementById("leadPositionHint");
    if (hint) {
      const kedo = kedoForYandex();
      hint.textContent = kedo
        ? `В заявку: ${roleLabel()} · КЭДО: ${kedo}`
        : `В заявку: ${roleLabel()}`;
    }
    refreshYandexBridge(false);
  };
  els.btnShowLead.addEventListener("click", () => {
    els.leadPanel.classList.remove("is-hidden");
    els.leadPanel.scrollIntoView({ behavior: "smooth", block: "start" });
    _showLead();
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
