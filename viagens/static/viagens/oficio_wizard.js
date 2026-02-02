async function carregarCidades(estadoSelect, cidadeSelect) {
  const estado = estadoSelect?.value || "";
  cidadeSelect.innerHTML = '<option value="">Selecione</option>';
  cidadeSelect.value = "";
  window.syncAutocompleteDisplay?.(cidadeSelect);
  if (!estado) {
    return;
  }

  const response = await fetch(`/api/cidades/?estado=${encodeURIComponent(estado)}`);
  const data = await response.json();
  const selectedId = cidadeSelect.dataset.selected || "";

  data.cidades.forEach((cidade) => {
    const option = document.createElement("option");
    option.value = String(cidade.id);
    option.textContent = cidade.nome;
    if (selectedId && String(cidade.id) === String(selectedId)) {
      option.selected = true;
    }
    cidadeSelect.appendChild(option);
  });

  window.syncAutocompleteDisplay?.(cidadeSelect);
}

function setupCidadeSelect(estadoSelect, cidadeSelect) {
  if (!estadoSelect || !cidadeSelect) {
    return;
  }

  estadoSelect.addEventListener("change", () => {
    cidadeSelect.dataset.selected = "";
    carregarCidades(estadoSelect, cidadeSelect);
  });

  if (estadoSelect.value) {
    carregarCidades(estadoSelect, cidadeSelect);
  }
}

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

function toggleLock(field, locked) {
  if (!field) {
    return;
  }
  field.classList.toggle("is-locked", locked);
  field.setAttribute("aria-readonly", locked ? "true" : "false");
  if (locked) {
    field.tabIndex = -1;
  } else {
    field.removeAttribute("tabindex");
  }
}

function initRoteiroFormset() {
  const trechosList = document.querySelector("[data-trechos-list]");
  const template = document.getElementById("trechoTemplate");
  const totalFormsInput = document.getElementById("id_trechos-TOTAL_FORMS");
  const addTrechoBtn = document.getElementById("addTrechoBtn");
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
  const diariasPreviewPorServidor = document.getElementById(
    "diariasPreviewPorServidor"
  );
  const diariasPreviewTotal = document.getElementById("diariasPreviewTotal");
  const diariasPreviewHoras = document.getElementById("diariasPreviewHoras");

  if (!trechosList || !template || !totalFormsInput) {
    return;
  }

  const prefix = "trechos";

  const getCards = () => Array.from(trechosList.querySelectorAll(".trecho-card"));

  const getSelectLabel = (select) => {
    if (!select) {
      return "";
    }
    const input =
      select._autocompleteInput ||
      select.parentElement?.querySelector(".autocomplete-input");
    if (input && input.value) {
      return input.value.trim();
    }
    const option = select.selectedOptions?.[0];
    return option ? option.textContent.trim() : "";
  };

  const formatLocal = (cidadeSelect, estadoSelect) => {
    const cidade = getSelectLabel(cidadeSelect);
    const estado = estadoSelect?.value || "";
    if (cidade && cidade.includes("/")) {
      return cidade;
    }
    if (cidade && estado) {
      return `${cidade}/${estado}`;
    }
    return cidade || estado || "";
  };

  const getLastDestinoCard = () => {
    const cards = getCards();
    for (let i = cards.length - 1; i >= 0; i -= 1) {
      const destinoEstado =
        cards[i].querySelector("[data-role='destino-estado']")?.value || "";
      const destinoCidade =
        cards[i].querySelector("[data-role='destino-cidade']")?.value || "";
      if (destinoEstado && destinoCidade) {
        return cards[i];
      }
    }
    return cards[0] || null;
  };

  const updateRetornoFields = () => {
    if (!retornoSaidaCidadeInput && !retornoChegadaCidadeInput) {
      return;
    }
    const cards = getCards();
    const first = cards[0];
    const last = getLastDestinoCard();
    if (!first || !last) {
      return;
    }
    const sedeCidade = formatLocal(
      first.querySelector("[data-role='origem-cidade']"),
      first.querySelector("[data-role='origem-estado']")
    );
    const destinoCidade = formatLocal(
      last.querySelector("[data-role='destino-cidade']"),
      last.querySelector("[data-role='destino-estado']")
    );
    if (retornoSaidaCidadeInput) {
      retornoSaidaCidadeInput.value = destinoCidade;
    }
    if (retornoChegadaCidadeInput) {
      retornoChegadaCidadeInput.value = sedeCidade;
    }
  };

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

  const formatCurrency = (value) => {
    return value.toFixed(2);
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

  const reindexCards = () => {
    getCards().forEach((card, index) => updateElementIndex(card, index, prefix));
  };

  const updateTotals = () => {
    totalFormsInput.value = String(getCards().length);
  };

  const getSede = () => {
    const first = getCards()[0];
    if (!first) {
      return { estado: "", cidade: "" };
    }
    return {
      estado: first.querySelector("[data-role='origem-estado']")?.value || "",
      cidade: first.querySelector("[data-role='origem-cidade']")?.value || "",
    };
  };

  const setOrigin = (card, estado, cidade) => {
    const origemEstado = card.querySelector("[data-role='origem-estado']");
    const origemCidade = card.querySelector("[data-role='origem-cidade']");
    if (origemEstado) {
      origemEstado.value = estado || "";
      window.syncAutocompleteDisplay?.(origemEstado);
    }
    if (origemCidade) {
      origemCidade.dataset.selected = cidade || "";
      if (cidade) {
        const existing = Array.from(origemCidade.options).find(
          (opt) => String(opt.value) === String(cidade)
        );
        if (!existing) {
          const option = document.createElement("option");
          option.value = String(cidade);
          option.textContent = "Carregando...";
          option.selected = true;
          origemCidade.appendChild(option);
        } else {
          origemCidade.value = String(cidade);
        }
      }
      if (origemEstado && origemEstado.value) {
        carregarCidades(origemEstado, origemCidade).then(() => {
          window.syncAutocompleteDisplay?.(origemCidade);
        });
      } else {
        origemCidade.innerHTML = '<option value="">Selecione</option>';
        origemCidade.value = "";
        window.syncAutocompleteDisplay?.(origemCidade);
      }
    }
  };

  const applyCardLabels = () => {
    getCards().forEach((card, index) => {
      card.classList.toggle("is-first", index === 0);
      const badge = card.querySelector(".badge");
      if (badge) {
        badge.textContent = `Trecho ${index + 1}${index === 0 ? " (Ida)" : ""}`;
      }
      const titles = card.querySelectorAll(".sub-card-title");
      if (titles.length) {
        titles[0].textContent = index === 0 ? "Origem (Sede)" : "Origem";
      }
      const removeBtn = card.querySelector("[data-action='remove']");
      if (removeBtn) {
        const disabled = index === 0;
        removeBtn.disabled = disabled;
        removeBtn.setAttribute("aria-disabled", disabled ? "true" : "false");
      }
    });
  };

  const lockOrigins = () => {
    getCards().forEach((card, index) => {
      const origemEstado = card.querySelector("[data-role='origem-estado']");
      const origemCidade = card.querySelector("[data-role='origem-cidade']");
      const shouldLock = index !== 0;
      toggleLock(origemEstado, shouldLock);
      toggleLock(origemCidade, shouldLock);
    });
  };

  const syncOrigins = () => {
    const cards = getCards();
    cards.forEach((card, index) => {
      if (index === 0) {
        return;
      }
      const prevCard = cards[index - 1];
      const prevEstado =
        prevCard.querySelector("[data-role='destino-estado']")?.value || "";
      const prevCidade =
        prevCard.querySelector("[data-role='destino-cidade']")?.value || "";
      setOrigin(card, prevEstado, prevCidade);
    });
    updateRetornoFields();
  };

  const removeCardsAfter = (index) => {
    getCards().forEach((card, idx) => {
      if (idx > index) {
        card.remove();
      }
    });
    reindexCards();
    updateTotals();
    updateRetornoFields();
  };

  const addCard = (origemEstado, origemCidade) => {
    const fragment = template.content.cloneNode(true);
    const newCard = fragment.querySelector(".trecho-card");
    const index = getCards().length;
    updateElementIndex(newCard, index, prefix);
    trechosList.appendChild(fragment);
    window.initializeAutocompleteSelects?.(newCard);
    setupCard(newCard);
    setOrigin(newCard, origemEstado, origemCidade);
    reindexCards();
    updateTotals();
    applyCardLabels();
    lockOrigins();
    updateRetornoFields();
    return newCard;
  };

  const isCardComplete = (card) => {
    const destinoEstado =
      card.querySelector("[data-role='destino-estado']")?.value || "";
    const destinoCidade =
      card.querySelector("[data-role='destino-cidade']")?.value || "";
    return Boolean(destinoEstado && destinoCidade);
  };

  const isAutoPlaceholder = (card) => {
    if (isCardComplete(card)) {
      return false;
    }
    const destinoEstado =
      card.querySelector("[data-role='destino-estado']")?.value || "";
    const destinoCidade =
      card.querySelector("[data-role='destino-cidade']")?.value || "";
    if (destinoEstado || destinoCidade) {
      return false;
    }
    const saidaData = card.querySelector("[data-role='saida-data']")?.value || "";
    const saidaHora = card.querySelector("[data-role='saida-hora']")?.value || "";
    const chegadaData =
      card.querySelector("[data-role='chegada-data']")?.value || "";
    const chegadaHora =
      card.querySelector("[data-role='chegada-hora']")?.value || "";
    return !saidaData && !saidaHora && !chegadaData && !chegadaHora;
  };


  const handleDestinoChange = (card) => {
    const cards = getCards();
    const index = cards.indexOf(card);
    if (index < 0) {
      return;
    }
    removeCardsAfter(index);
    syncOrigins();
    applyCardLabels();
    lockOrigins();
    updateRetornoFields();
    updateDiarias();
    updateDiariasPreview();
  };

  const handleSedeChange = () => {
    removeCardsAfter(0);
    syncOrigins();
    applyCardLabels();
    lockOrigins();
    const first = getCards()[0];
    if (first) {
      handleDestinoChange(first);
    }
    updateRetornoFields();
    updateDiarias();
    updateDiariasPreview();
  };

  const setupCard = (card) => {
    window.initializeAutocompleteSelects?.(card);
    const origemEstado = card.querySelector("[data-role='origem-estado']");
    const origemCidade = card.querySelector("[data-role='origem-cidade']");
    const destinoEstado = card.querySelector("[data-role='destino-estado']");
    const destinoCidade = card.querySelector("[data-role='destino-cidade']");
    const saidaData = card.querySelector("[data-role='saida-data']");
    const saidaHora = card.querySelector("[data-role='saida-hora']");

    if (origemEstado && origemCidade) {
      setupCidadeSelect(origemEstado, origemCidade);
      origemEstado.addEventListener("change", handleSedeChange);
      origemCidade.addEventListener("change", handleSedeChange);
    }
    if (destinoEstado && destinoCidade) {
      setupCidadeSelect(destinoEstado, destinoCidade);
      destinoEstado.addEventListener("change", () => {
        destinoCidade.dataset.selected = "";
        handleDestinoChange(card);
      });
      destinoCidade.addEventListener("change", () => handleDestinoChange(card));
    }
    if (saidaData) {
      saidaData.addEventListener("change", updateDiarias);
    }
    if (saidaHora) {
      saidaHora.addEventListener("change", updateDiarias);
    }
    if (saidaData) {
      saidaData.addEventListener("change", updateDiariasPreview);
    }
    if (saidaHora) {
      saidaHora.addEventListener("change", updateDiariasPreview);
    }

    const removeBtn = card.querySelector("[data-action='remove']");
    if (removeBtn) {
      removeBtn.addEventListener("click", () => {
        const cards = getCards();
        const index = cards.indexOf(card);
        if (index <= 0) {
          return;
        }
        if (
          !window.confirm(
            "Remover este trecho? Os trechos seguintes serao ajustados."
          )
        ) {
          return;
        }
        card.remove();
        reindexCards();
        updateTotals();
        syncOrigins();
        applyCardLabels();
        lockOrigins();
        updateRetornoFields();
        updateDiarias();
        updateDiariasPreview();
      });
    }
  };

  const trimTrailingEmptyCards = () => {
    let cards = getCards();
    while (cards.length > 1 && isAutoPlaceholder(cards[cards.length - 1])) {
      cards[cards.length - 1].remove();
      cards = getCards();
    }
    reindexCards();
    updateTotals();
    syncOrigins();
    applyCardLabels();
    lockOrigins();
    updateRetornoFields();
    updateDiarias();
    updateDiariasPreview();
  };

  getCards().forEach((card) => {
    window.initializeAutocompleteSelects?.(card);
    setupCard(card);
  });
  syncOrigins();
  applyCardLabels();
  lockOrigins();
  updateTotals();
  updateRetornoFields();
  updateDiarias();
  updateDiariasPreview();

  if (!totalFormsInput.value || Number(totalFormsInput.value) < 1) {
    totalFormsInput.value = "1";
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
      trimTrailingEmptyCards();
    });
  }

  if (addTrechoBtn) {
    addTrechoBtn.addEventListener("click", () => {
      const cards = getCards();
      const last = cards[cards.length - 1];
      if (!last) {
        return;
      }
      const destinoEstado =
        last.querySelector("[data-role='destino-estado']")?.value || "";
      const destinoCidade =
        last.querySelector("[data-role='destino-cidade']")?.value || "";
      if (!destinoEstado || !destinoCidade) {
        window.alert(
          "Selecione o destino do trecho atual antes de adicionar um novo."
        );
        return;
      }
      const sede = getSede();
      const destinoEhSede =
        destinoEstado === sede.estado && destinoCidade === sede.cidade;
      if (destinoEhSede) {
        window.alert(
          "O roteiro ja retornou a sede. Nao e necessario adicionar novos trechos."
        );
        return;
      }
      addCard(destinoEstado, destinoCidade);
    });
  }
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initRoteiroFormset);
} else {
  initRoteiroFormset();
}
