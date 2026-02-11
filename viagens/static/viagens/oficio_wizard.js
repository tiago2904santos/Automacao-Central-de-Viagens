function updateElementIndex(element, index, prefix) {
  const regex = new RegExp(`${prefix}-(\\d+|__prefix__)-`, "g");
  element.querySelectorAll("[name], [id], [for]").forEach((item) => {
    if (item.name) {
      item.name = item.name.replace(regex, `${prefix}-${index}-`);
    }
    if (item.id) {
      item.id = item.id.replace(regex, `${prefix}-${index}-`);
    }
    if (item.htmlFor) {
      item.htmlFor = item.htmlFor.replace(regex, `${prefix}-${index}-`);
    }
  });
  element.dataset.index = String(index);
}

function initRoteiroFormset() {
  const trechosList = document.querySelector("[data-trechos-list]");
  const totalFormsInput = document.getElementById("id_trechos-TOTAL_FORMS");
  const destinoList = document.getElementById("destinosList");
  const destinoTemplate = document.getElementById("destinoTemplate");
  const addDestinoBtn = document.getElementById("addDestinoBtn");
  const destinosOrderInput = document.getElementById("destinosOrder");
  const destinosTotalInput = document.getElementById("destinosTotalForms");
  const roteiroForm = document.getElementById("roteiroForm");
  const retornoSaidaCidadeInput = document.getElementById("retornoSaidaCidade");
  const retornoChegadaCidadeInput = document.getElementById("retornoChegadaCidade");
  const retornoSaidaDataInput = document.getElementById("retornoSaidaData");
  const retornoSaidaHoraInput = document.getElementById("retornoSaidaHora");
  const retornoChegadaDataInput = document.getElementById("retornoChegadaData");
  const retornoChegadaHoraInput = document.getElementById("retornoChegadaHora");
  const tipoDestinoInput = document.getElementById("tipoDestino");
  const quantidadeDiariasInput = document.getElementById("quantidadeDiarias");
  const valorDiariasInput = document.getElementById("valorDiarias");
  const valorDiariasExtensoInput = document.getElementById("valorDiariasExtenso");
  const servidoresSelect = document.getElementById("servidoresSelect");
  const diariasPanel = document.getElementById("diariasPanel");
  const diariasCta = document.getElementById("diariasCta");
  const diariasResults = document.getElementById("diariasResults");
  const calcularDiariasBtn = document.getElementById("calcularDiariasBtn");
  const recalcularDiariasBtn = document.getElementById("recalcularDiariasBtn");
  const calcularDiariasBtnLabel = document.getElementById("calcularDiariasBtnLabel");
  const calculoDiariasMensagem = document.getElementById("calculoDiariasMensagem");
  const statusEl = document.getElementById("diariasStatus");
  const diariasTableBody = document.getElementById("diariasTableBody");
  const diariasTotalCard = document.getElementById("diariasTotal");
  const diariasExtensoCard = document.getElementById("diariasExtenso");
  const diariasQtdCard = document.getElementById("diariasQtd");
  const diariasHorasCard = document.getElementById("diariasHoras");
  const diariasTotalQtd = document.getElementById("diariasTotalQtd");
  const diariasTotalHoras = document.getElementById("diariasTotalHoras");
  const diariasTotalValor = document.getElementById("diariasTotalValor");
  const diariasCalcError = document.getElementById("diariasCalcError");
  const sedeUfSelect = document.getElementById("sedeUf");
  const sedeCidadeSelect = document.getElementById("sedeCidade");
  let diariasState = {
    hasResult: false,
    isStale: false,
    isLoading: false,
  };

  if (!trechosList || !totalFormsInput || !destinoList || !destinoTemplate) {
    return;
  }

  const trechoTemplate = document.getElementById("trechoTemplate");
  const sampleTrechoCard = trechosList.querySelector(
    ".trecho-card[data-index]:not([data-retorno-card]):not([data-diarias-card]):not([data-diarias-preview])"
  );

  const prefix = "trechos";
  const getCards = () => Array.from(trechosList.querySelectorAll(".trecho-card"));
  const getDestinosItems = () => Array.from(destinoList.querySelectorAll(".destino-item"));

  const getSelectLabel = (select) => {
    if (!select) {
      return "";
    }
    const input = select._autocompleteInput || select.parentElement?.querySelector(".autocomplete-input");
    if (input && input.value) {
      return input.value.trim();
    }
    const option = select.selectedOptions?.[0];
    return option ? option.textContent.trim() : "";
  };

  const formatLabel = (cidade, estado) => {
    const cidadeTrim = (cidade || "").trim();
    const estadoTrim = (estado || "").trim();
    if (cidadeTrim && cidadeTrim.includes("/")) {
      return cidadeTrim;
    }
    if (cidadeTrim && estadoTrim) {
      return `${cidadeTrim}/${estadoTrim}`;
    }
    return cidadeTrim || estadoTrim || "";
  };

  const CAPITAIS = new Set([
    "ARACAJU",
    "BELEM",
    "BELO HORIZONTE",
    "BOA VISTA",
    "BRASILIA",
    "CAMPO GRANDE",
    "CUIABA",
    "CURITIBA",
    "FLORIANOPOLIS",
    "FORTALEZA",
    "GOIANIA",
    "JOAO PESSOA",
    "MACAPA",
    "MACEIO",
    "MANAUS",
    "NATAL",
    "PALMAS",
    "PORTO ALEGRE",
    "PORTO VELHO",
    "RECIFE",
    "RIO BRANCO",
    "RIO DE JANEIRO",
    "SALVADOR",
    "SAO LUIS",
    "SAO PAULO",
    "TERESINA",
    "VITORIA",
  ]);

  const normalizeCityName = (value) =>
    (value || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toUpperCase()
      .trim();

  const inferTipoDestino = () => {
    const destinos = collectDestinosData().filter((dest) => dest.uf && dest.cidade);
    if (!destinos.length) {
      return "";
    }
    let temCapital = false;
    for (const destino of destinos) {
      const uf = (destino.uf || "").toUpperCase();
      const cidadeBase = normalizeCityName((destino.label || "").split("/")[0]);
      if (uf === "DF" && cidadeBase === "BRASILIA") {
        return "BRASILIA";
      }
      if (CAPITAIS.has(cidadeBase)) {
        temCapital = true;
      }
    }
    return temCapital ? "CAPITAL" : "INTERIOR";
  };

  const syncTipoDestino = () => {
    const tipo = inferTipoDestino();
    if (tipoDestinoInput) {
      tipoDestinoInput.value = tipo;
    }
    return tipo;
  };

  const getSedeLabel = () => {
    if (!sedeCidadeSelect) {
      return formatLabel("", sedeUfSelect?.value || "");
    }
    const cidadeLabel = getSelectLabel(sedeCidadeSelect);
    const estadoSigla = sedeUfSelect?.value || "";
    return formatLabel(cidadeLabel, estadoSigla);
  };

  const collectDestinosData = () =>
    getDestinosItems().map((item) => {
      const estado = item.querySelector("[data-role='destino-estado']");
      const cidade = item.querySelector("[data-role='destino-cidade']");
      const cidadeLabel = getSelectLabel(cidade);
      const estadoSigla = estado?.value || "";
      return {
        uf: estadoSigla,
        cidade: cidade?.value || "",
        label: formatLabel(cidadeLabel, estadoSigla),
        element: item,
      };
    });

  const setCalculoMensagem = (text) => {
    if (calculoDiariasMensagem) {
      calculoDiariasMensagem.textContent = text || "";
    }
  };

  const setCalculoErro = (text) => {
    if (!diariasCalcError) {
      return;
    }
    const hasError = Boolean(text);
    diariasCalcError.hidden = !hasError;
    diariasCalcError.textContent = hasError ? text : "";
  };

  const setButtonsLoading = (isLoading) => {
    [calcularDiariasBtn, recalcularDiariasBtn].forEach((btn) => {
      if (!btn) {
        return;
      }
      btn.disabled = isLoading;
      btn.classList.toggle("is-loading", isLoading);
    });
    if (calcularDiariasBtnLabel) {
      calcularDiariasBtnLabel.textContent = isLoading
        ? "Calculando..."
        : "Calcular diarias";
    }
  };

  const clearDiariasStatus = () => {
    if (!statusEl) {
      return;
    }
    statusEl.classList.remove(
      "bg-success",
      "bg-danger",
      "bg-warning",
      "bg-secondary",
      "bg-info",
      "status-danger-soft"
    );
    statusEl.textContent = "";
    statusEl.hidden = true;
  };

  const updateResultsVisibility = (showResults) => {
    if (diariasCta) {
      diariasCta.hidden = showResults;
    }
    if (!diariasResults) {
      return;
    }
    if (!showResults) {
      diariasResults.classList.remove("is-visible");
      diariasResults.hidden = true;
      return;
    }
    diariasResults.hidden = false;
    requestAnimationFrame(() => {
      diariasResults.classList.add("is-visible");
    });
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

  const formatTipoLabel = (tipo) => {
    const normalized = String(tipo || "").toUpperCase();
    if (normalized === "BRASILIA") {
      return "BRASILIA";
    }
    if (normalized === "CAPITAL") {
      return "CAPITAL";
    }
    return "INTERIOR";
  };

  const tipoClassName = (tipo) => {
    const normalized = String(tipo || "").toUpperCase();
    if (normalized === "BRASILIA") {
      return "diarias-chip diarias-chip--brasilia";
    }
    if (normalized === "CAPITAL") {
      return "diarias-chip diarias-chip--capital";
    }
    return "diarias-chip diarias-chip--interior";
  };

  const resetDiariasData = (clearHiddenFields) => {
    if (diariasTableBody) {
      diariasTableBody.innerHTML = "";
    }
    if (diariasTotalCard) {
      diariasTotalCard.textContent = "-";
    }
    if (diariasExtensoCard) {
      diariasExtensoCard.textContent = "-";
    }
    if (diariasQtdCard) {
      diariasQtdCard.textContent = "-";
    }
    if (diariasHorasCard) {
      diariasHorasCard.textContent = "-";
    }
    if (diariasTotalQtd) {
      diariasTotalQtd.textContent = "-";
    }
    if (diariasTotalHoras) {
      diariasTotalHoras.textContent = "-";
    }
    if (diariasTotalValor) {
      diariasTotalValor.textContent = "-";
    }
    if (clearHiddenFields) {
      if (quantidadeDiariasInput) {
        quantidadeDiariasInput.value = "";
      }
      if (valorDiariasInput) {
        valorDiariasInput.value = "";
      }
      if (valorDiariasExtensoInput) {
        valorDiariasExtensoInput.value = "";
      }
    }
  };

  const setPanelDatasetState = (stateName) => {
    if (diariasPanel) {
      diariasPanel.dataset.state = stateName;
    }
  };

  const setDiariasIdle = () => {
    diariasState = {
      hasResult: false,
      isStale: false,
      isLoading: false,
    };
    setButtonsLoading(false);
    setPanelDatasetState("idle");
    updateResultsVisibility(false);
    clearDiariasStatus();
    setCalculoMensagem("Clique para calcular com base no roteiro.");
  };

  const setDiariasLoading = () => {
    diariasState = {
      hasResult: diariasState.hasResult,
      isStale: false,
      isLoading: true,
    };
    setButtonsLoading(true);
    setPanelDatasetState("loading");
    updateResultsVisibility(diariasState.hasResult);
    clearDiariasStatus();
    if (statusEl) {
      statusEl.classList.add("bg-info");
      statusEl.textContent = "Calculando...";
      statusEl.hidden = false;
    }
    setCalculoMensagem("Calculando diarias...");
  };

  const setDiariasDone = () => {
    diariasState = {
      hasResult: true,
      isStale: false,
      isLoading: false,
    };
    setButtonsLoading(false);
    setPanelDatasetState("done");
    updateResultsVisibility(true);
    clearDiariasStatus();
    if (statusEl) {
      statusEl.classList.add("bg-success");
      statusEl.textContent = "Calculo atualizado";
      statusEl.hidden = false;
    }
    setCalculoMensagem("Calculo atualizado.");
  };

  const setDiariasStale = () => {
    if (!diariasState.hasResult) {
      setDiariasIdle();
      return;
    }
    diariasState = {
      hasResult: true,
      isStale: true,
      isLoading: false,
    };
    setButtonsLoading(false);
    setPanelDatasetState("stale");
    updateResultsVisibility(true);
    clearDiariasStatus();
    if (statusEl) {
      statusEl.classList.add("bg-danger");
      statusEl.textContent = "Calculo desatualizado";
      statusEl.hidden = false;
    }
    setCalculoMensagem("Calculo desatualizado. Clique em Recalcular.");
  };

  const addCell = (row, value, className = "") => {
    const cell = document.createElement("td");
    cell.textContent = value || "-";
    if (className) {
      cell.className = className;
    }
    row.appendChild(cell);
  };

  const renderDiariasResultado = (payload) => {
    const periodos = Array.isArray(payload?.periodos) ? payload.periodos : [];
    const totais = payload?.totais || {};

    if (diariasTableBody) {
      diariasTableBody.innerHTML = "";
      periodos.forEach((periodo) => {
        const row = document.createElement("tr");
        const tipoCell = document.createElement("td");
        const chip = document.createElement("span");
        chip.className = tipoClassName(periodo.tipo);
        chip.textContent = formatTipoLabel(periodo.tipo);
        tipoCell.appendChild(chip);
        row.appendChild(tipoCell);

        const saidaLabel = [periodo.data_saida, periodo.hora_saida]
          .filter(Boolean)
          .join(" ");
        const chegadaLabel = [periodo.data_chegada, periodo.hora_chegada]
          .filter(Boolean)
          .join(" ");
        addCell(row, saidaLabel);
        addCell(row, chegadaLabel);
        addCell(row, String(periodo.n_diarias ?? ""), "diarias-cell--number");
        addCell(
          row,
          formatHours(periodo.horas_adicionais).replace("h", ""),
          "diarias-cell--number"
        );
        addCell(row, formatMoney(periodo.valor_diaria), "diarias-cell--money");
        addCell(row, formatMoney(periodo.subtotal), "diarias-cell--money");
        diariasTableBody.appendChild(row);
      });
    }

    if (diariasTotalCard) {
      diariasTotalCard.textContent = formatMoney(totais.total_valor);
    }
    if (diariasExtensoCard) {
      diariasExtensoCard.textContent = totais.valor_extenso || "-";
    }
    if (diariasQtdCard) {
      diariasQtdCard.textContent = totais.total_diarias || "-";
    }
    if (diariasHorasCard) {
      diariasHorasCard.textContent = formatHours(totais.total_horas);
    }
    if (diariasTotalQtd) {
      diariasTotalQtd.textContent = totais.total_diarias || "-";
    }
    if (diariasTotalHoras) {
      diariasTotalHoras.textContent = formatHours(totais.total_horas);
    }
    if (diariasTotalValor) {
      diariasTotalValor.textContent = formatMoney(totais.total_valor);
    }
    if (quantidadeDiariasInput) {
      quantidadeDiariasInput.value = totais.total_diarias || "";
    }
    if (valorDiariasInput) {
      valorDiariasInput.value = totais.total_valor || "";
    }
    if (valorDiariasExtensoInput && totais.valor_extenso) {
      valorDiariasExtensoInput.value = totais.valor_extenso;
    }
    if (tipoDestinoInput && payload?.tipo_destino) {
      tipoDestinoInput.value = payload.tipo_destino;
    }
    setCalculoErro("");
    setDiariasDone();
  };

  const getCsrfToken = () => {
    const tokenInput = roteiroForm?.querySelector("input[name='csrfmiddlewaretoken']");
    return tokenInput?.value || "";
  };

  const getServidoresCount = () => {
    if (servidoresSelect) {
      return servidoresSelect.selectedOptions.length;
    }
    return Number(roteiroForm?.dataset.servidoresCount || "0");
  };

  const requestDiariasCalculation = async () => {
    if (!roteiroForm) {
      return;
    }
    const calcUrl = roteiroForm.dataset.calcDiariasUrl || "";
    if (!calcUrl) {
      return;
    }

    trimTrailingEmptyDestinos();
    regenerateTrechos();
    syncTipoDestino();
    syncDestinosOrder();

    const formData = new FormData(roteiroForm);
    formData.set("quantidade_servidores", String(getServidoresCount()));

    const hadResultBefore = diariasState.hasResult;
    setCalculoErro("");
    setDiariasLoading();

    try {
      const response = await fetch(calcUrl, {
        method: "POST",
        credentials: "same-origin",
        headers: {
          "X-CSRFToken": getCsrfToken(),
        },
        body: formData,
      });
      let payload = {};
      try {
        payload = await response.json();
      } catch (_err) {
        payload = {};
      }
      if (!response.ok) {
        const message = payload?.error || "Preencha datas e horas para calcular.";
        throw new Error(message);
      }
      renderDiariasResultado(payload);
    } catch (error) {
      if (hadResultBefore) {
        diariasState.hasResult = true;
        setDiariasStale();
      } else {
        resetDiariasData(false);
        setDiariasIdle();
      }
      setCalculoErro(error?.message || "Preencha datas e horas para calcular.");
    }
  };

  const invalidateDiarias = () => {
    syncTipoDestino();
    setCalculoErro("");
    if (diariasState.isLoading) {
      return;
    }
    if (diariasState.hasResult) {
      setDiariasStale();
      return;
    }
    setDiariasIdle();
  };

  const syncTrechosTotal = () => {
    totalFormsInput.value = String(getCards().length);
  };

  const clearCitySelect = (select) => {
    if (!select) {
      return;
    }
    select.value = "";
    select.dataset.selected = "";
    window.syncAutocompleteDisplay?.(select);
  };

  const setRouteData = (card, route) => {
    const origemEstado = card.querySelector("[data-trecho-origem-estado]");
    const origemCidade = card.querySelector("[data-trecho-origem-cidade]");
    const destinoEstado = card.querySelector("[data-trecho-destino-estado]");
    const destinoCidade = card.querySelector("[data-trecho-destino-cidade]");
    if (origemEstado) {
      origemEstado.value = route?.origem_estado || "";
    }
    if (origemCidade) {
      origemCidade.value = route?.origem_cidade || "";
    }
    if (destinoEstado) {
      destinoEstado.value = route?.destino_estado || "";
    }
    if (destinoCidade) {
      destinoCidade.value = route?.destino_cidade || "";
    }
  };

  const resetClonedTrechoCard = (card) => {
    card.querySelectorAll("input, select, textarea").forEach((field) => {
      if (field.type === "checkbox" || field.type === "radio") {
        field.checked = false;
      } else {
        field.value = "";
      }
    });
    card.querySelectorAll("[data-trecho-route-preview]").forEach((preview) => {
      preview.textContent = "Origem -> Destino";
    });
    card.querySelectorAll("[data-trecho-origem-preview], [data-trecho-destino-preview]").forEach((preview) => {
      preview.textContent = "-";
    });
  };

  const updateRoutePreview = (card, origemLabel, destinoLabel) => {
    const preview = card.querySelector("[data-trecho-route-preview]");
    if (preview) {
      if (origemLabel || destinoLabel) {
        preview.textContent = `${origemLabel || "-"} -> ${destinoLabel || "-"}`;
      } else {
        preview.textContent = "Origem -> Destino";
      }
    }
    const origemPreview = card.querySelector("[data-trecho-origem-preview]");
    const destinoPreview = card.querySelector("[data-trecho-destino-preview]");
    if (origemPreview) {
      origemPreview.textContent = origemLabel || "-";
    }
    if (destinoPreview) {
      destinoPreview.textContent = destinoLabel || "-";
    }
  };

  const regenerateTrechos = () => {
    const destinos = collectDestinosData();
    const validDestinos = destinos.filter((dest) => dest.uf && dest.cidade);
    const origemBase = {
      uf: sedeUfSelect?.value || "",
      cidade: sedeCidadeSelect?.value || "",
    };
    const preservedDates = getCards().map((card) => ({
      saida_data: card.querySelector("[data-role='saida-data']")?.value || "",
      saida_hora: card.querySelector("[data-role='saida-hora']")?.value || "",
      chegada_data: card.querySelector("[data-role='chegada-data']")?.value || "",
      chegada_hora: card.querySelector("[data-role='chegada-hora']")?.value || "",
    }));
    const targetCount = Math.max(1, validDestinos.length);
    while (getCards().length < targetCount) {
      if (trechoTemplate) {
        const fragment = trechoTemplate.content.cloneNode(true);
        const newCard = fragment.querySelector(".trecho-card");
        if (!newCard) {
          break;
        }
        updateElementIndex(newCard, getCards().length, prefix);
        trechosList.appendChild(fragment);
      } else if (sampleTrechoCard) {
        const cloned = sampleTrechoCard.cloneNode(true);
        resetClonedTrechoCard(cloned);
        updateElementIndex(cloned, getCards().length, prefix);
        trechosList.appendChild(cloned);
      } else {
        break;
      }
    }
    while (getCards().length > targetCount) {
      const cards = getCards();
      cards[cards.length - 1].remove();
    }
    getCards().forEach((card, index) => {
      updateElementIndex(card, index, prefix);
      setupCardListeners(card);
    });
    const cards = getCards();
    cards.forEach((card, index) => {
      const destino = validDestinos[index];
      const origemLabel =
        index === 0 ? getSedeLabel() : validDestinos[index - 1]?.label || "";
      const destinoLabel = destino?.label || "";
      const route = destino
        ? {
            origem_estado: index === 0 ? origemBase.uf : validDestinos[index - 1]?.uf || "",
            origem_cidade: index === 0 ? origemBase.cidade : validDestinos[index - 1]?.cidade || "",
            destino_estado: destino.uf,
            destino_cidade: destino.cidade,
          }
        : {
            origem_estado: "",
            origem_cidade: "",
            destino_estado: "",
            destino_cidade: "",
          };
      setRouteData(card, route);
      updateRoutePreview(card, origemLabel, destinoLabel);
      const preserved = preservedDates[index] || {};
      const saidaData = card.querySelector("[data-role='saida-data']");
      const saidaHora = card.querySelector("[data-role='saida-hora']");
      const chegadaData = card.querySelector("[data-role='chegada-data']");
      const chegadaHora = card.querySelector("[data-role='chegada-hora']");
      if (saidaData) {
        saidaData.value = preserved.saida_data || "";
      }
      if (saidaHora) {
        saidaHora.value = preserved.saida_hora || "";
      }
      if (chegadaData) {
        chegadaData.value = preserved.chegada_data || "";
      }
      if (chegadaHora) {
        chegadaHora.value = preserved.chegada_hora || "";
      }
    });
    syncTrechosTotal();
    updateRetornoFields();
    invalidateDiarias();
  };

  const updateRetornoFields = () => {
    if (!retornoSaidaCidadeInput && !retornoChegadaCidadeInput) {
      return;
    }
    const validDestinos = collectDestinosData().filter((dest) => dest.uf && dest.cidade);
    const lastDestino = validDestinos[validDestinos.length - 1];
    const sedeLabel = getSedeLabel();
    if (retornoSaidaCidadeInput) {
      retornoSaidaCidadeInput.value = lastDestino?.label || "";
    }
    if (retornoChegadaCidadeInput) {
      retornoChegadaCidadeInput.value = sedeLabel;
    }
  };

  const updateDestinosTotal = () => {
    if (destinosTotalInput) {
      destinosTotalInput.value = String(getDestinosItems().length);
    }
  };

  const syncDestinosOrder = () => {
    if (!destinosOrderInput) {
      return;
    }
    const order = getDestinosItems()
      .map((item) => item.getAttribute("data-index") || "")
      .filter(Boolean);
    destinosOrderInput.value = order.join(",");
  };

  const syncDestinos = () => {
    updateDestinosTotal();
    syncDestinosOrder();
    validateAllDestinos();
  };

  const updateDestinoValidation = (item) => {
    const cidade = item.querySelector("[data-role='destino-cidade']");
    const error = item.querySelector("[data-destino-error]");
    const isInvalid = !cidade?.value;
    item.classList.toggle("is-invalid", isInvalid);
    if (error) {
      error.hidden = !isInvalid;
    }
    return !isInvalid;
  };

  const validateAllDestinos = () => {
    getDestinosItems().forEach(updateDestinoValidation);
  };

  const isDestinoEmpty = (item) => {
    const estado = item.querySelector("[data-role='destino-estado']");
    const cidade = item.querySelector("[data-role='destino-cidade']");
    return !(estado?.value || cidade?.value);
  };

  const trimTrailingEmptyDestinos = () => {
    let items = getDestinosItems();
    while (items.length > 1 && isDestinoEmpty(items[items.length - 1])) {
      items[items.length - 1].remove();
      items = getDestinosItems();
    }
    syncDestinos();
  };

  const setupCardListeners = (card) => {
    window.initializeAutocompleteSelects?.(card);
    const saidaData = card.querySelector("[data-role='saida-data']");
    const saidaHora = card.querySelector("[data-role='saida-hora']");
    const chegadaData = card.querySelector("[data-role='chegada-data']");
    const chegadaHora = card.querySelector("[data-role='chegada-hora']");
    if (saidaData) {
      saidaData.addEventListener("change", invalidateDiarias);
    }
    if (saidaHora) {
      saidaHora.addEventListener("change", invalidateDiarias);
    }
    if (chegadaData) {
      chegadaData.addEventListener("change", invalidateDiarias);
    }
    if (chegadaHora) {
      chegadaHora.addEventListener("change", invalidateDiarias);
    }
  };

  const setupDestinoItem = (item) => {
    window.initializeAutocompleteSelects?.(item);
    const estado = item.querySelector("[data-role='destino-estado']");
    const cidade = item.querySelector("[data-role='destino-cidade']");
    if (estado && cidade) {
      estado.addEventListener("change", () => {
        clearCitySelect(cidade);
        updateDestinoValidation(item);
        regenerateTrechos();
      });
      cidade.addEventListener("change", () => {
        updateDestinoValidation(item);
        regenerateTrechos();
      });
    }
    const removeBtn = item.querySelector("[data-action='remove-destino']");
    if (removeBtn) {
      removeBtn.addEventListener("click", () => {
        const items = getDestinosItems();
        if (items.length <= 1) {
          if (estado) {
            estado.value = "";
          }
          if (cidade) {
            cidade.value = "";
            cidade.dispatchEvent(new Event("change"));
          }
          return;
        }
        item.remove();
        // After removal, reindex to keep destinos-0..N-1 contiguous for POST parsing.
        getDestinosItems().forEach((destinoItem, index) =>
          updateElementIndex(destinoItem, index, "destinos")
        );
        syncDestinos();
        regenerateTrechos();
      });
    }
    updateDestinoValidation(item);
  };

  const addDestinoItem = () => {
    const fragment = destinoTemplate.content.cloneNode(true);
    const newItem = fragment.querySelector(".destino-item");
    const index = getDestinosItems().length;
    updateElementIndex(newItem, index, "destinos");
    destinoList.appendChild(fragment);
    setupDestinoItem(newItem);
    syncDestinos();
    regenerateTrechos();
  };

  const initDestinos = () => {
    const items = getDestinosItems();
    if (!items.length) {
      addDestinoItem();
      return;
    }
    items.forEach((item, index) => {
      updateElementIndex(item, index, "destinos");
      setupDestinoItem(item);
    });
    syncDestinos();
  };

  const initDestinosSorting = () => {
    let dragItem = null;

    const handleDragStart = (event) => {
      const handle = event.target.closest(".drag-handle");
      if (!handle) {
        return;
      }
      const item = handle.closest(".destino-item");
      if (!item) {
        return;
      }
      dragItem = item;
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", item.getAttribute("data-index") || "");
      item.classList.add("is-dragging");
    };

    const handleDragOver = (event) => {
      event.preventDefault();
      const overItem = event.target.closest(".destino-item");
      if (!overItem || !dragItem || overItem === dragItem) {
        return;
      }
      const rect = overItem.getBoundingClientRect();
      const after = event.clientY - rect.top > rect.height / 2;
      if (after) {
        overItem.after(dragItem);
      } else {
        overItem.before(dragItem);
      }
    };

    const handleDragEnd = () => {
      if (dragItem) {
        dragItem.classList.remove("is-dragging");
      }
      dragItem = null;
      // Keep destino indices stable during reorder; the backend uses `destinos-order`
      // to apply the user-defined order without renaming all form fields on drag.
      syncDestinos();
      regenerateTrechos();
    };

    destinoList.addEventListener("dragstart", handleDragStart);
    destinoList.addEventListener("dragover", handleDragOver);
    destinoList.addEventListener("dragend", handleDragEnd);
  };

  initDestinos();
  regenerateTrechos();
  initDestinosSorting();

  if (addDestinoBtn) {
    addDestinoBtn.addEventListener("click", addDestinoItem);
  }
  if (calcularDiariasBtn) {
    calcularDiariasBtn.addEventListener("click", requestDiariasCalculation);
  }
  if (recalcularDiariasBtn) {
    recalcularDiariasBtn.addEventListener("click", requestDiariasCalculation);
  }
  resetDiariasData(false);
  setDiariasIdle();

  if (sedeUfSelect) {
    sedeUfSelect.addEventListener("change", () => {
      clearCitySelect(sedeCidadeSelect);
      regenerateTrechos();
    });
  }
  if (sedeCidadeSelect) {
    sedeCidadeSelect.addEventListener("change", regenerateTrechos);
  }

  if (retornoChegadaDataInput) {
    retornoChegadaDataInput.addEventListener("change", invalidateDiarias);
  }
  if (retornoChegadaHoraInput) {
    retornoChegadaHoraInput.addEventListener("change", invalidateDiarias);
  }
  if (retornoSaidaDataInput) {
    retornoSaidaDataInput.addEventListener("change", invalidateDiarias);
  }
  if (retornoSaidaHoraInput) {
    retornoSaidaHoraInput.addEventListener("change", invalidateDiarias);
  }
  if (servidoresSelect) {
    servidoresSelect.addEventListener("change", invalidateDiarias);
  }

  if (roteiroForm) {
    roteiroForm.addEventListener("submit", () => {
      trimTrailingEmptyDestinos();
      regenerateTrechos();
      syncTipoDestino();
      syncDestinosOrder();
    });
  }
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initRoteiroFormset);
} else {
  initRoteiroFormset();
}
