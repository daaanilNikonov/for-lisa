(() => {
  const data = window.QUIZ_SCRIPTS;
  const profileLabels = window.QUIZ_PROFILE_LABELS;

  const state = {
    step: 0,
    answers: {},
    filterRole: "all",
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
    resultImagePath: document.getElementById("resultImagePath"),
    resultProfile: document.getElementById("resultProfile"),
    scriptsFilters: document.getElementById("scriptsFilters"),
    scriptsGrid: document.getElementById("scriptsGrid"),
  };

  function getQuestion(index) {
    return data.questions[index];
  }

  function getResult(role, size) {
    return data.results.find((r) => r.role === role && r.size === size);
  }

  function optionLabel(questionId, optionId) {
    const q = data.questions.find((item) => item.id === questionId);
    return q?.options.find((o) => o.id === optionId)?.label || optionId;
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
    const result = getResult(state.answers.role, state.answers.size);
    if (!result) return;

    els.progressBar.style.width = "100%";
    els.quizSection.classList.add("is-hidden");
    els.resultSection.classList.remove("is-hidden");

    els.resultTitle.textContent = result.title;
    els.resultMeta.textContent = `${optionLabel("role", result.role)} · ${optionLabel(
      "size",
      result.size
    )} · id: ${result.id}`;

    els.resultText.textContent = result.text || "Здесь будет текст";
    els.resultNote.textContent = result.contentNote || "";
    els.resultImagePath.textContent = result.imageSlot;

    if (result.imageSrc) {
      els.resultImage.hidden = false;
      els.resultImage.src = result.imageSrc;
      els.resultImage.alt = result.imageAlt || "";
      els.resultImagePlaceholder.hidden = true;
    } else {
      els.resultImage.hidden = true;
      els.resultImage.removeAttribute("src");
      els.resultImagePlaceholder.hidden = false;
    }

    const profileKeys = ["docs", "certs", "auto", "overtime", "report"];
    els.resultProfile.innerHTML = profileKeys
      .map((key) => {
        const value = state.answers[key];
        const label = profileLabels[key]?.[value] || value;
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

  function renderFilters() {
    const roles = [
      { id: "all", label: "Все" },
      ...data.questions[0].options.map((o) => ({ id: o.id, label: o.label })),
    ];

    els.scriptsFilters.innerHTML = roles
      .map(
        (role) => `
      <button type="button" class="chip ${
        state.filterRole === role.id ? "is-active" : ""
      }" data-role="${role.id}">${role.label}</button>`
      )
      .join("");

    els.scriptsFilters.querySelectorAll(".chip").forEach((chip) => {
      chip.addEventListener("click", () => {
        state.filterRole = chip.dataset.role;
        renderFilters();
        renderScriptsGrid();
      });
    });
  }

  function renderScriptsGrid() {
    const list =
      state.filterRole === "all"
        ? data.results
        : data.results.filter((r) => r.role === state.filterRole);

    els.scriptsGrid.innerHTML = list
      .map(
        (r) => `
      <article class="script-card" id="script-${r.id}">
        <h3>${r.label}</h3>
        <div class="script-id">${r.id}</div>
        <div class="slots">
          <div class="mini-slot">
            <strong>Текст:</strong> ${r.text || "Здесь будет текст"}
          </div>
          <div class="mini-slot">
            <strong>Картинка:</strong> ${r.imageSrc ? r.imageSrc : `слот · ${r.imageSlot}`}
          </div>
        </div>
      </article>`
      )
      .join("");
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

  renderQuestion();
  renderFilters();
  renderScriptsGrid();
})();
