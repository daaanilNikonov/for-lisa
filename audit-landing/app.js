(() => {
  const LEAD_KEY = "forus-audit-leads-v1";
  const data = window.AUDIT_SCRIPTS;
  const match = window.AUDIT_MATCH;

  const state = {
    step: 0,
    answers: {},
    findings: [],
    scenario: null,
  };

  const els = {
    quizCard: document.getElementById("quizCard"),
    quizStep: document.getElementById("quizStep"),
    progressBar: document.getElementById("progressBar"),
    btnBack: document.getElementById("btnBack"),
    btnNext: document.getElementById("btnNext"),
    btnRestart: document.getElementById("btnRestart"),
    btnShowLead: document.getElementById("btnShowLead"),
    btnCopyCode: document.getElementById("btnCopyCode"),
    leadPanel: document.getElementById("leadPanel"),
    resultSection: document.getElementById("result"),
    quizSection: document.getElementById("quiz"),
    scenarioCode: document.getElementById("scenarioCode"),
    scenarioTitle: document.getElementById("scenarioTitle"),
    scenarioSummary: document.getElementById("scenarioSummary"),
    scenarioDemoHint: document.getElementById("scenarioDemoHint"),
    auditList: document.getElementById("auditList"),
    auditEmpty: document.getElementById("auditEmpty"),
    leadForm: document.getElementById("leadForm"),
    leadStatus: document.getElementById("leadStatus"),
    leadHint: document.getElementById("leadHint"),
    leadScenarioHint: document.getElementById("leadScenarioHint"),
    btnOpenGoogle: document.getElementById("btnOpenGoogle"),
  };

  function questions() {
    return data.questions;
  }

  function normalizePhone(raw) {
    let digits = String(raw || "").replace(/\D+/g, "");
    if (digits.length === 11 && digits.startsWith("8")) {
      digits = `7${digits.slice(1)}`;
    }
    if (digits.length === 10) digits = `7${digits}`;
    return digits;
  }

  function isAnswered() {
    const q = questions()[state.step];
    return Boolean(q && state.answers[q.id]);
  }

  function optionLabel(questionId, optionId) {
    const q = questions().find((item) => item.id === questionId);
    const opt = q?.options?.find((o) => o.id === optionId);
    return opt?.label || optionId || "";
  }

  function answersSummary() {
    return questions()
      .map((q) => `${q.title} → ${optionLabel(q.id, state.answers[q.id])}`)
      .join("\n");
  }

  function renderQuestion() {
    const list = questions();
    if (state.step >= list.length) state.step = list.length - 1;
    const q = list[state.step];
    const total = list.length;
    const selected = state.answers[q.id];

    els.quizStep.textContent = `Вопрос ${state.step + 1} из ${total}`;
    els.progressBar.style.width = `${(state.step / total) * 100}%`;
    els.btnBack.disabled = state.step === 0;
    els.btnNext.disabled = !isAnswered();
    els.btnNext.textContent =
      state.step === total - 1 ? "Показать аудит" : "Далее";

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
    state.findings = match.buildAudit(state.answers);
    state.scenario = match.pickScenario(state.answers);
    const sc = state.scenario;

    els.scenarioCode.textContent = sc.code;
    els.scenarioTitle.textContent = sc.title;
    els.scenarioSummary.textContent = sc.summary;
    els.scenarioDemoHint.textContent = `На демо: ${sc.demoHint}`;
    if (els.leadScenarioHint) els.leadScenarioHint.textContent = sc.code;

    els.auditList.innerHTML = state.findings
      .map((f) => `<li>${f.text}</li>`)
      .join("");
    const empty = state.findings.length === 0;
    els.auditEmpty.classList.toggle("is-hidden", !empty);
    els.auditList.classList.toggle("is-hidden", empty);

    els.quizSection.classList.add("is-hidden");
    els.resultSection.classList.remove("is-hidden");
    els.leadPanel.classList.add("is-hidden");
    els.resultSection.scrollIntoView({ behavior: "smooth", block: "start" });
    syncGoogleLink();
  }

  function restart() {
    state.step = 0;
    state.answers = {};
    state.findings = [];
    state.scenario = null;
    els.resultSection.classList.add("is-hidden");
    els.quizSection.classList.remove("is-hidden");
    els.leadPanel.classList.add("is-hidden");
    if (els.leadStatus) {
      els.leadStatus.textContent = "";
      els.leadStatus.className = "lead-status";
    }
    renderQuestion();
    els.quizSection.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function googleCfg() {
    return data.googleForm || {};
  }

  function googleConfigured() {
    const cfg = googleCfg();
    const map = cfg.fields || {};
    if (!cfg.url || cfg.url.includes("REPLACE_FORM_ID")) return false;
    return Object.values(map).every(
      (v) => typeof v === "string" && v && !v.includes("REPLACE_")
    );
  }

  function buildGooglePrefillUrl(payload) {
    const cfg = googleCfg();
    const map = cfg.fields || {};
    const url = new URL(cfg.url.replace("/formResponse", "/viewform"));
    if (map.name) url.searchParams.set(map.name, payload.name || "");
    if (map.phone) url.searchParams.set(map.phone, payload.phone || "");
    if (map.inn) url.searchParams.set(map.inn, payload.inn || "");
    if (map.company) url.searchParams.set(map.company, payload.company || "");
    if (map.scenario) url.searchParams.set(map.scenario, payload.scenario || "");
    if (map.audit) url.searchParams.set(map.audit, payload.audit || "");
    return url.toString();
  }

  function currentLeadDraft() {
    return {
      name: document.getElementById("leadName")?.value.trim() || "",
      company: document.getElementById("leadCompany")?.value.trim() || "",
      phone: normalizePhone(document.getElementById("leadPhone")?.value || ""),
      inn: document.getElementById("leadInn")?.value.trim() || "",
      scenario: state.scenario?.code || "",
      audit: state.findings.map((f) => f.text).join("; "),
      answers: answersSummary(),
    };
  }

  function leadDraftValid(draft) {
    if (!draft.name) return "Укажите имя";
    if (!/^\d{10}$|^\d{12}$/.test(draft.inn)) {
      return "Проверьте ИНН: нужно 10 или 12 цифр";
    }
    if (draft.phone.length < 11 || draft.phone.length > 12) {
      return "Проверьте телефон: нужно 11–12 цифр, например +7 900 123-45-67";
    }
    return "";
  }

  function saveLeadLocal(payload) {
    const rows = JSON.parse(localStorage.getItem(LEAD_KEY) || "[]");
    rows.push({
      ...payload,
      created_at: new Date().toISOString(),
    });
    localStorage.setItem(LEAD_KEY, JSON.stringify(rows));
  }

  function syncGoogleLink() {
    if (!els.btnOpenGoogle) return;
    const draft = currentLeadDraft();
    const cfg = googleCfg();
    if (!cfg.url || cfg.url.includes("REPLACE_FORM_ID")) {
      els.btnOpenGoogle.classList.add("is-hidden");
      if (els.leadHint) {
        els.leadHint.textContent =
          "Google Form ещё не подключена: укажите url и entry.* в scripts-data.js → googleForm. Заявка сохранится в браузере, код сценария можно скопировать.";
      }
      return;
    }
    els.btnOpenGoogle.classList.remove("is-hidden");
    try {
      els.btnOpenGoogle.href = buildGooglePrefillUrl(draft);
    } catch (err) {
      els.btnOpenGoogle.href = cfg.url;
    }
    if (els.leadHint) {
      els.leadHint.textContent = googleConfigured()
        ? `В форму уйдёт код ${draft.scenario} и краткий аудит.`
        : "Откройте Google Form и проверьте поля — часть entry.* ещё с REPLACE_.";
    }
  }

  async function submitLead(event) {
    event.preventDefault();
    const draft = currentLeadDraft();
    const err = leadDraftValid(draft);
    if (err) {
      els.leadStatus.textContent = err;
      els.leadStatus.className = "lead-status is-err";
      return;
    }

    saveLeadLocal(draft);
    syncGoogleLink();

    const cfg = googleCfg();
    if (cfg.url && !cfg.url.includes("REPLACE_FORM_ID")) {
      const prefill = buildGooglePrefillUrl(draft);
      window.open(prefill, "_blank", "noopener");
      els.leadStatus.textContent =
        "Заявка сохранена локально. Открыли Google Form с предзаполнением — нажмите «Отправить» в форме.";
      els.leadStatus.className = "lead-status is-ok";
      return;
    }

    els.leadStatus.textContent =
      "Заявка сохранена в браузере. Подключите Google Form в scripts-data.js, чтобы отправлять на сервер.";
    els.leadStatus.className = "lead-status is-ok";
  }

  async function copyScenarioCode() {
    const code = state.scenario?.code || "";
    if (!code) return;
    try {
      await navigator.clipboard.writeText(code);
      els.btnCopyCode.textContent = "Скопировано";
      setTimeout(() => {
        els.btnCopyCode.textContent = "Скопировать код";
      }, 1600);
    } catch (err) {
      prompt("Скопируйте код сценария:", code);
    }
  }

  els.btnBack.addEventListener("click", () => {
    if (state.step > 0) {
      state.step -= 1;
      renderQuestion();
    }
  });

  els.btnNext.addEventListener("click", () => {
    if (!isAnswered()) return;
    if (state.step >= questions().length - 1) {
      showResult();
      return;
    }
    state.step += 1;
    renderQuestion();
  });

  els.btnRestart.addEventListener("click", restart);
  els.btnShowLead.addEventListener("click", () => {
    els.leadPanel.classList.remove("is-hidden");
    syncGoogleLink();
    els.leadPanel.scrollIntoView({ behavior: "smooth", block: "start" });
  });
  els.btnCopyCode.addEventListener("click", copyScenarioCode);
  els.leadForm.addEventListener("submit", submitLead);
  ["leadName", "leadCompany", "leadPhone", "leadInn"].forEach((id) => {
    document.getElementById(id)?.addEventListener("input", syncGoogleLink);
  });

  renderQuestion();
})();
