function updateElementIndex(element, index, prefix) {
  const regex = new RegExp(`${prefix}-(\d+|__prefix__)-`, "g");
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
  const tipoDestinoSelect = document.getElementById("tipoDestino");
  const quantidadeDiariasInput = document.getElementById("quantidadeDiarias");
  const valorDiariasInput = document.getElementById("valorDiarias");
  const servidoresSelect = document.getElementById("servidoresSelect");
  const diariasPreviewText = document.getElementById("diariasPreviewText");
  const diariasPreviewPorServidor = document.getElementById("diariasPreviewPorServidor");
  const diariasPreviewTotal = document.getElementById("diariasPreviewTotal");
  const diariasPreviewHoras = document.getElementById("diariasPreviewHoras");
  const sedeUfSelect = document.getElementById("sedeUf");
  const sedeCidadeSelect = document.getElementById("sedeCidade");

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

  const parseDateTime = (dateValue, timeValue) => {
    if (!dateValue) {
      return null;
    }
    const time = timeValue || "00:00";
    const dt = new Date(`${dateValue}T${time}:00`);
    return Number.isNaN(dt.getTime()) ? null : dt;
  };

  const calcularDiarias = (tipoDestino, saidaDt, chegadaDt, servidores) => {
    if (!tipoDestino || !saidaDt || !chegadaDt) {
      return {
        quantidade: "",
        valorTotal: "",
      };
    }
    const totalMs = chegadaDt.getTime() - saidaDt.getTime();
    if (totalMs <= 0) {
      return { quantidade: "", valorTotal: "" };
    }
    const totalHoras = totalMs / (1000 * 60 * 60);
    let diasInteiros = Math.floor(totalHoras / 24);
    let restoMs = totalMs - diasInteiros * 24 * 60 * 60 * 1000;
    let parcial = 0;

    if (
      saidaDt.toDateString() !== chegadaDt.toDateString() &&
      totalMs < 24 * 60 * 60 * 1000
    ) {
      diasInteiros = 1;
      restoMs = 0;
    } else if (restoMs <= 6 * 60 * 60 * 1000) {
      parcial = 0;
    } else if (restoMs <= 8 * 60 * 60 * 1000) {
      parcial = 15;
    } else {
      parcial = 30;
    }

    const tabela = {
      INTERIOR: { full: 290.55, p15: 43.58, p30: 87.17 },
      CAPITAL: { full: 371.26, p15: 55.69, p30: 111.38 },
      BRASILIA: { full: 468.12, p15: 70.22, p30: 140.43 },
    };
    const valores = tabela[tipoDestino] || { full: 0, p15: 0, p30: 0 };
    const valorParcial = parcial === 15 ? valores.p15 : parcial === 30 ? valores.p30 : 0;
    const valor1 = diasInteiros * valores.full + valorParcial;
    const total = valor1 * (servidores || 0);

    const partes = [];
    if (diasInteiros > 0) {
      partes.push(`${diasInteiros} x 100%`);
    }
    if (parcial > 0) {
      partes.push(`1 x ${parcial}%`);
    }

    return {
      quantidade: partes.join(" + "),
      valorTotal: total.toFixed(2).replace(".", ","),
    };
  };

  const formatCurrency = (value) => {
    return value.toFixed(2);
  };

  const updateDiarias = () => {
    if (!tipoDestinoSelect || !quantidadeDiariasInput || !valorDiariasInput) {
      return;
    }
    const first = getCards()[0];
    const saidaData = first?.querySelector("[data-role='saida-data']")?.value || "";
    const saidaHora = first?.querySelector("[data-role='saida-hora']")?.value || "";
    const chegadaData = retornoChegadaDataInput?.value || "";
    const chegadaHora = retornoChegadaHoraInput?.value || "";
    const saidaDt = parseDateTime(saidaData, saidaHora);
    const chegadaDt = parseDateTime(chegadaData, chegadaHora);
    const servidores = servidoresSelect
      ? servidoresSelect.selectedOptions.length
      : Number(roteiroForm?.dataset.servidoresCount || "0");
    const resultado = calcularDiarias(
      tipoDestinoSelect.value,
      saidaDt,
      chegadaDt,
      servidores
    );
    quantidadeDiariasInput.value = resultado.quantidade || "";
    valorDiariasInput.value = resultado.valorTotal || "";
  };

  const updateDiariasPreview = () => {
    if (
      !diariasPreviewText ||
      !diariasPreviewPorServidor ||
      !diariasPreviewTotal ||
      !diariasPreviewHoras
    ) {
      return;
    }
    const first = getCards()[0];
    const saidaData = first?.querySelector("[data-role='saida-data']")?.value || "";
    const saidaHora = first?.querySelector("[data-role='saida-hora']")?.value || "";
    const chegadaData = retornoChegadaDataInput?.value || "";
    const chegadaHora = retornoChegadaHoraInput?.value || "";
    const saidaDt = parseDateTime(saidaData, saidaHora);
    const chegadaDt = parseDateTime(chegadaData, chegadaHora);
    const servidores = servidoresSelect
      ? servidoresSelect.selectedOptions.length
      : Number(roteiroForm?.dataset.servidoresCount || "0");

    if (!tipoDestinoSelect?.value || !saidaDt || !chegadaDt || servidores <= 0) {
      diariasPreviewText.textContent = "Preencha saída/chegada para calcular";
      diariasPreviewPorServidor.textContent = "-";
      diariasPreviewTotal.textContent = "-";
      diariasPreviewHoras.textContent = "-";
      return;
    }

    const totalMs = chegadaDt.getTime() - saidaDt.getTime();
    if (totalMs <= 0) {
      diariasPreviewText.textContent = "Preencha saída/chegada para calcular";
      diariasPreviewPorServidor.textContent = "-";
      diariasPreviewTotal.textContent = "-";
      diariasPreviewHoras.textContent = "-";
      return;
    }

    const totalHoras = totalMs / (1000 * 60 * 60);
    let diasInteiros = Math.floor(totalHoras / 24);
    let restoHoras = totalHoras - diasInteiros * 24;
    let parcial = 0;

    if (
      saidaDt.toDateString() !== chegadaDt.toDateString() &&
      totalHoras < 24
    ) {
      diasInteiros = 1;
      restoHoras = 0;
    } else if (restoHoras <= 6) {
      parcial = 0;
    } else if (restoHoras <= 8) {
      parcial = 15;
    } else {
      parcial = 30;
    }

    const tabela = {
      INTERIOR: { full: 290.55, p15: 43.58, p30: 87.17 },
      CAPITAL: { full: 371.26, p15: 55.69, p30: 111.38 },
      BRASILIA: { full: 468.12, p15: 70.22, p30: 140.43 },
    };
    const valores = tabela[tipoDestinoSelect.value] || {
      full: 0,
      p15: 0,
      p30: 0,
    };
    const valorParcial =
      parcial === 15 ? valores.p15 : parcial === 30 ? valores.p30 : 0;
    const valorPorServidor = diasInteiros * valores.full + valorParcial;
    const valorTotal = valorPorServidor * servidores;

    const partes = [];
    if (diasInteiros > 0) {
      partes.push(`${diasInteiros} x 100%`);
    }
    if (parcial > 0) {
      partes.push(`1 x ${parcial}%`);
    }

    diariasPreviewText.textContent = partes.join(" + ") || "0";
    diariasPreviewPorServidor.textContent = formatCurrency(valorPorServidor);
    diariasPreviewTotal.textContent = formatCurrency(valorTotal);
    diariasPreviewHoras.textContent = `${totalHoras.toFixed(2)}h`;
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
      preview.textContent = "Origem → Destino";
    });
    card.querySelectorAll("[data-trecho-origem-preview], [data-trecho-destino-preview]").forEach((preview) => {
      preview.textContent = "-";
    });
  };

  const updateRoutePreview = (card, origemLabel, destinoLabel) => {
    const preview = card.querySelector("[data-trecho-route-preview]");
    if (preview) {
      if (origemLabel || destinoLabel) {
        preview.textContent = `${origemLabel || "-"} → ${destinoLabel || "-"}`;
      } else {
        preview.textContent = "Origem → Destino";
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
    updateDiarias();
    updateDiariasPreview();
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

  const updateDestinosOrder = () => {
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
    updateDestinosOrder();
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
      saidaData.addEventListener("change", updateDiarias);
      saidaData.addEventListener("change", updateDiariasPreview);
    }
    if (saidaHora) {
      saidaHora.addEventListener("change", updateDiarias);
      saidaHora.addEventListener("change", updateDiariasPreview);
    }
    if (chegadaData) {
      chegadaData.addEventListener("change", updateDiarias);
      chegadaData.addEventListener("change", updateDiariasPreview);
    }
    if (chegadaHora) {
      chegadaHora.addEventListener("change", updateDiarias);
      chegadaHora.addEventListener("change", updateDiariasPreview);
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
      getDestinosItems().forEach((item, index) => updateElementIndex(item, index, "destinos"));
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

  if (sedeUfSelect) {
    sedeUfSelect.addEventListener("change", () => {
      clearCitySelect(sedeCidadeSelect);
      regenerateTrechos();
    });
  }
  if (sedeCidadeSelect) {
    sedeCidadeSelect.addEventListener("change", regenerateTrechos);
  }

  if (tipoDestinoSelect) {
    tipoDestinoSelect.addEventListener("change", updateDiarias);
    tipoDestinoSelect.addEventListener("change", updateDiariasPreview);
  }
  if (retornoChegadaDataInput) {
    retornoChegadaDataInput.addEventListener("change", updateDiarias);
    retornoChegadaDataInput.addEventListener("change", updateDiariasPreview);
  }
  if (retornoChegadaHoraInput) {
    retornoChegadaHoraInput.addEventListener("change", updateDiarias);
    retornoChegadaHoraInput.addEventListener("change", updateDiariasPreview);
  }
  if (retornoSaidaDataInput) {
    retornoSaidaDataInput.addEventListener("change", updateDiarias);
    retornoSaidaDataInput.addEventListener("change", updateDiariasPreview);
  }
  if (retornoSaidaHoraInput) {
    retornoSaidaHoraInput.addEventListener("change", updateDiarias);
    retornoSaidaHoraInput.addEventListener("change", updateDiariasPreview);
  }
  if (servidoresSelect) {
    servidoresSelect.addEventListener("change", updateDiarias);
    servidoresSelect.addEventListener("change", updateDiariasPreview);
  }

  if (roteiroForm) {
    roteiroForm.addEventListener("submit", () => {
      trimTrailingEmptyDestinos();
      regenerateTrechos();
    });
  }
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initRoteiroFormset);
} else {
  initRoteiroFormset();
}
