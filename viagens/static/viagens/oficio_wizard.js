const motoristaSelect = document.getElementById("motoristaSelect");
const motoristaNome = document.getElementById("motoristaNome");

if (motoristaSelect && motoristaNome) {
  const selected = motoristaSelect.selectedOptions[0];
  if (selected && selected.value && !motoristaNome.value) {
    motoristaNome.value = selected.textContent.trim();
  }
  motoristaSelect.addEventListener("change", () => {
    const selectedOption = motoristaSelect.selectedOptions[0];
    if (selectedOption && selectedOption.value) {
      motoristaNome.value = selectedOption.textContent.trim();
    }
  });
}

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

  if (!trechosList || !template || !totalFormsInput) {
    return;
  }

  const prefix = "trechos";

  const getCards = () => Array.from(trechosList.querySelectorAll(".trecho-card"));
  const getLastCard = () => getCards()[getCards().length - 1];

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
  };

  const removeCardsAfter = (index) => {
    getCards().forEach((card, idx) => {
      if (idx > index) {
        card.remove();
      }
    });
    reindexCards();
    updateTotals();
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
    const saidaData = card.querySelector("[data-role='saida-data']")?.value || "";
    const saidaHora = card.querySelector("[data-role='saida-hora']")?.value || "";
    const chegadaData =
      card.querySelector("[data-role='chegada-data']")?.value || "";
    const chegadaHora =
      card.querySelector("[data-role='chegada-hora']")?.value || "";
    return !saidaData && !saidaHora && !chegadaData && !chegadaHora;
  };

  const maybeAddNext = (currentCard) => {
    if (currentCard !== getLastCard()) {
      return;
    }
    const destinoEstado =
      currentCard.querySelector("[data-role='destino-estado']")?.value || "";
    const destinoCidade =
      currentCard.querySelector("[data-role='destino-cidade']")?.value || "";
    if (!destinoEstado || !destinoCidade) {
      return;
    }
    const sede = getSede();
    const destinoEhSede =
      destinoEstado === sede.estado && destinoCidade === sede.cidade;
    if (!destinoEhSede) {
      addCard(destinoEstado, destinoCidade);
    }
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
    maybeAddNext(card);
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
  };

  const setupCard = (card) => {
    window.initializeAutocompleteSelects?.(card);
    const origemEstado = card.querySelector("[data-role='origem-estado']");
    const origemCidade = card.querySelector("[data-role='origem-cidade']");
    const destinoEstado = card.querySelector("[data-role='destino-estado']");
    const destinoCidade = card.querySelector("[data-role='destino-cidade']");

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
        const last = getLastCard();
        if (last) {
          maybeAddNext(last);
        }
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
  };

  getCards().forEach((card) => {
    window.initializeAutocompleteSelects?.(card);
    setupCard(card);
  });
  syncOrigins();
  applyCardLabels();
  lockOrigins();
  updateTotals();

  const firstCard = getCards()[0];
  if (firstCard) {
    maybeAddNext(firstCard);
  }

  const roteiroForm = document.getElementById("roteiroForm");
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
