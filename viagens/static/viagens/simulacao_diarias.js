(function () {
  const form = document.getElementById("simulacaoDiariasForm");
  if (!form) {
    return;
  }

  const periodsList = document.getElementById("simPeriodosList");
  const addPeriodoBtn = document.getElementById("addPeriodoBtn");
  const periodTemplate = document.getElementById("simPeriodoTemplate");
  const periodsPayloadInput = document.getElementById("simPeriodsPayload");
  const qtdServidoresInput = form.querySelector("input[name='quantidade_servidores']");

  const calcBtn = document.getElementById("simCalcularBtn");
  const calcBtnLabel = document.getElementById("simCalcularBtnLabel");
  const novaBtn = document.getElementById("simNovaBtn");
  const errorEl = document.getElementById("simCalcError");
  const resultCard = document.getElementById("simResultadoCard");

  const tableBody = document.getElementById("simTableBody");
  const totalEl = document.getElementById("simTotal");
  const totalValorEl = document.getElementById("simTotalValor");
  const valorExtensoEl = document.getElementById("simValorExtenso");
  const qtdEl = document.getElementById("simQtd");
  const horasEl = document.getElementById("simHoras");
  const porServidorQtdEl = document.getElementById("simPorServidorQtd");
  const porServidorValorEl = document.getElementById("simPorServidorValor");
  const valorUnitarioEl = document.getElementById("simValorUnitario");

  const initialPeriodsNode = document.getElementById("simInitialPeriods");
  let initialPeriods = [];
  if (initialPeriodsNode?.textContent) {
    try {
      initialPeriods = JSON.parse(initialPeriodsNode.textContent);
    } catch (_err) {
      initialPeriods = [];
    }
  }

  const emptyMarkers = { total: "-", extenso: "-", qtd: "-", horas: "-" };

  const getCsrfToken = () =>
    form.querySelector("input[name='csrfmiddlewaretoken']")?.value || "";

  const clearError = () => {
    if (!errorEl) {
      return;
    }
    errorEl.hidden = true;
    errorEl.textContent = "";
  };

  const setError = (message) => {
    if (!errorEl) {
      return;
    }
    errorEl.hidden = !message;
    errorEl.textContent = message || "";
  };

  const setLoading = (isLoading) => {
    if (!calcBtn) {
      return;
    }
    calcBtn.disabled = isLoading;
    calcBtn.classList.toggle("is-loading", isLoading);
    if (calcBtnLabel) {
      calcBtnLabel.textContent = isLoading ? "Calculando..." : "Calcular";
    }
  };

  const formatMoney = (value) => {
    const raw = String(value || "").trim();
    if (!raw) {
      return "-";
    }
    return raw.startsWith("R$") ? raw : `R$ ${raw}`;
  };

  const formatHours = (value) => {
    if (value === undefined || value === null || value === "") {
      return "-";
    }
    return `${String(value).replace(".", ",")}h`;
  };

  const tipoClass = (tipo) => {
    const normalized = String(tipo || "").toUpperCase();
    if (normalized === "CAPITAL") {
      return "diarias-chip diarias-chip--capital";
    }
    if (normalized === "BRASILIA") {
      return "diarias-chip diarias-chip--brasilia";
    }
    return "diarias-chip diarias-chip--interior";
  };

  const tipoLabel = (tipo) => {
    const normalized = String(tipo || "").toUpperCase();
    if (normalized === "CAPITAL") {
      return "CAPITAL";
    }
    if (normalized === "BRASILIA") {
      return "BRASILIA";
    }
    return "INTERIOR";
  };

  const hideResult = () => {
    if (!resultCard || resultCard.hidden) {
      return;
    }
    resultCard.classList.remove("is-visible");
    window.setTimeout(() => {
      if (!resultCard.classList.contains("is-visible")) {
        resultCard.hidden = true;
      }
    }, 220);
  };

  const showResult = () => {
    if (!resultCard) {
      return;
    }
    resultCard.hidden = false;
    requestAnimationFrame(() => {
      resultCard.classList.add("is-visible");
    });
  };

  const resetResult = () => {
    if (tableBody) {
      tableBody.innerHTML = "";
    }
    if (totalEl) {
      totalEl.textContent = emptyMarkers.total;
    }
    if (totalValorEl) {
      totalValorEl.textContent = emptyMarkers.total;
    }
    if (valorExtensoEl) {
      valorExtensoEl.textContent = emptyMarkers.extenso;
    }
    if (qtdEl) {
      qtdEl.textContent = emptyMarkers.qtd;
    }
    if (horasEl) {
      horasEl.textContent = emptyMarkers.horas;
    }
    if (porServidorQtdEl) {
      porServidorQtdEl.textContent = emptyMarkers.qtd;
    }
    if (porServidorValorEl) {
      porServidorValorEl.textContent = emptyMarkers.total;
    }
    if (valorUnitarioEl) {
      valorUnitarioEl.textContent = emptyMarkers.total;
    }
    hideResult();
  };

  const addCell = (row, value, className = "") => {
    const cell = document.createElement("td");
    cell.textContent = value || "-";
    if (className) {
      cell.className = className;
    }
    row.appendChild(cell);
  };

  const renderResult = (payload) => {
    const periodos = Array.isArray(payload?.periodos) ? payload.periodos : [];
    const totais = payload?.totais || {};
    if (tableBody) {
      tableBody.innerHTML = "";
      periodos.forEach((periodo) => {
        const row = document.createElement("tr");
        const tipoCell = document.createElement("td");
        const chip = document.createElement("span");
        chip.className = tipoClass(periodo.tipo);
        chip.textContent = tipoLabel(periodo.tipo);
        tipoCell.appendChild(chip);
        row.appendChild(tipoCell);

        addCell(row, periodo.data_saida || "");
        addCell(row, periodo.hora_saida || "");
        addCell(row, periodo.data_chegada || "");
        addCell(row, periodo.hora_chegada || "");
        addCell(row, String(periodo.n_diarias ?? ""), "diarias-cell--number");
        addCell(
          row,
          formatHours(periodo.horas_adicionais).replace("h", ""),
          "diarias-cell--number"
        );
        addCell(row, formatMoney(periodo.valor_diaria), "diarias-cell--money");
        addCell(row, formatMoney(periodo.subtotal), "diarias-cell--money");
        tableBody.appendChild(row);
      });
    }

    if (totalEl) {
      totalEl.textContent = formatMoney(totais.total_valor);
    }
    if (totalValorEl) {
      totalValorEl.textContent = formatMoney(totais.total_valor);
    }
    if (valorExtensoEl) {
      valorExtensoEl.textContent = totais.valor_extenso || "-";
    }
    if (qtdEl) {
      qtdEl.textContent = totais.total_diarias || "-";
    }
    if (horasEl) {
      horasEl.textContent = formatHours(totais.total_horas);
    }
    if (porServidorQtdEl) {
      porServidorQtdEl.textContent = totais.diarias_por_servidor || "-";
    }
    if (porServidorValorEl) {
      porServidorValorEl.textContent = formatMoney(totais.valor_por_servidor);
    }
    if (valorUnitarioEl) {
      valorUnitarioEl.textContent = formatMoney(totais.valor_unitario_referencia);
    }
    showResult();
  };

  const getPeriodFields = (card) => {
    const get = (field) =>
      (card.querySelector(`[data-period-field='${field}']`)?.value || "").trim();
    return {
      tipo: get("tipo").toUpperCase(),
      start_date: get("start_date"),
      start_time: get("start_time"),
      end_date: get("end_date"),
      end_time: get("end_time"),
    };
  };

  const setPeriodError = (card, message) => {
    const errorElPeriod = card.querySelector("[data-period-error]");
    if (message) {
      card.classList.add("is-invalid");
    } else {
      card.classList.remove("is-invalid");
    }
    if (!errorElPeriod) {
      return;
    }
    errorElPeriod.hidden = !message;
    errorElPeriod.textContent = message || "";
  };

  const validatePeriodCard = (card, index) => {
    const fields = getPeriodFields(card);
    if (
      !fields.tipo ||
      !fields.start_date ||
      !fields.start_time ||
      !fields.end_date ||
      !fields.end_time
    ) {
      setPeriodError(card, `Periodo ${index + 1}: preencha todos os campos.`);
      return null;
    }
    const start = new Date(`${fields.start_date}T${fields.start_time}:00`);
    const end = new Date(`${fields.end_date}T${fields.end_time}:00`);
    if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime()) || end <= start) {
      setPeriodError(card, `Periodo ${index + 1}: inicio deve ser anterior ao fim.`);
      return null;
    }
    setPeriodError(card, "");
    return fields;
  };

  const updateRemoveButtonsVisibility = () => {
    if (!periodsList) {
      return;
    }
    const cards = Array.from(periodsList.querySelectorAll("[data-periodo-item]"));
    cards.forEach((card) => {
      const removeBtn = card.querySelector("[data-remove-periodo]");
      if (!removeBtn) {
        return;
      }
      removeBtn.hidden = cards.length <= 1;
      removeBtn.disabled = cards.length <= 1;
    });
  };

  const bindCardValidation = (card) => {
    card.querySelectorAll("[data-period-field]").forEach((field) => {
      field.addEventListener("input", () => {
        const cards = Array.from(periodsList.querySelectorAll("[data-periodo-item]"));
        const idx = cards.indexOf(card);
        validatePeriodCard(card, idx >= 0 ? idx : 0);
      });
      field.addEventListener("change", () => {
        const cards = Array.from(periodsList.querySelectorAll("[data-periodo-item]"));
        const idx = cards.indexOf(card);
        validatePeriodCard(card, idx >= 0 ? idx : 0);
      });
    });
  };

  const createPeriodoCard = (data = {}) => {
    if (!periodTemplate || !periodsList) {
      return;
    }
    const fragment = periodTemplate.content.cloneNode(true);
    const card = fragment.querySelector("[data-periodo-item]");
    if (!card) {
      return;
    }
    const setValue = (field, value) => {
      const input = card.querySelector(`[data-period-field='${field}']`);
      if (input) {
        input.value = value || "";
      }
    };
    setValue("tipo", String(data.tipo || "INTERIOR").toUpperCase());
    setValue("start_date", data.start_date || "");
    setValue("start_time", data.start_time || "");
    setValue("end_date", data.end_date || "");
    setValue("end_time", data.end_time || "");

    const removeBtn = card.querySelector("[data-remove-periodo]");
    if (removeBtn) {
      removeBtn.addEventListener("click", () => {
        const cards = periodsList.querySelectorAll("[data-periodo-item]");
        if (cards.length <= 1) {
          return;
        }
        card.remove();
        updateRemoveButtonsVisibility();
      });
    }

    bindCardValidation(card);
    periodsList.appendChild(fragment);
    if (window.lucide) {
      window.lucide.createIcons({ attrs: { "stroke-width": 2 } });
    }
    updateRemoveButtonsVisibility();
  };

  const resetPeriodos = (seed = []) => {
    if (!periodsList) {
      return;
    }
    periodsList.innerHTML = "";
    if (Array.isArray(seed) && seed.length) {
      seed.forEach((periodo) => createPeriodoCard(periodo));
      return;
    }
    createPeriodoCard();
  };

  const getPeriodsPayload = () => {
    if (!periodsList) {
      return [];
    }
    const cards = Array.from(periodsList.querySelectorAll("[data-periodo-item]"));
    if (!cards.length) {
      throw new Error("Adicione ao menos um periodo para calcular.");
    }
    const payload = [];
    let hasInvalid = false;
    cards.forEach((card, index) => {
      const validated = validatePeriodCard(card, index);
      if (!validated) {
        hasInvalid = true;
        return;
      }
      payload.push(validated);
    });
    if (hasInvalid) {
      throw new Error("Corrija os periodos destacados para calcular.");
    }
    return payload;
  };

  const resetForNewSimulation = () => {
    form.reset();
    if (qtdServidoresInput) {
      qtdServidoresInput.value = "1";
    }
    resetPeriodos();
    clearError();
    resetResult();
  };

  const calculate = async () => {
    clearError();
    let periods = [];
    try {
      periods = getPeriodsPayload();
    } catch (validationError) {
      setError(validationError?.message || "Revise os periodos.");
      return;
    }
    if (periodsPayloadInput) {
      periodsPayloadInput.value = JSON.stringify(periods);
    }

    setLoading(true);
    try {
      const response = await fetch(form.dataset.calcUrl || "", {
        method: "POST",
        credentials: "same-origin",
        headers: {
          "X-CSRFToken": getCsrfToken(),
        },
        body: new FormData(form),
      });
      let payload = {};
      try {
        payload = await response.json();
      } catch (_err) {
        payload = {};
      }
      if (!response.ok) {
        throw new Error(payload?.error || "Nao foi possivel calcular as diarias.");
      }
      renderResult(payload);
    } catch (error) {
      setError(error?.message || "Nao foi possivel calcular as diarias.");
      resetResult();
    } finally {
      setLoading(false);
    }
  };

  form.addEventListener("submit", (event) => {
    event.preventDefault();
    calculate();
  });

  if (addPeriodoBtn) {
    addPeriodoBtn.addEventListener("click", () => createPeriodoCard());
  }
  if (novaBtn) {
    novaBtn.addEventListener("click", resetForNewSimulation);
  }

  resetPeriodos(initialPeriods);
  clearError();
  resetResult();
})();
