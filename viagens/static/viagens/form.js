const viajantesPreview = document.getElementById("viajantesPreview");
const placaInput = document.getElementById("plateInput");
const modeloInput = document.getElementById("modelInput");
const combustivelInput = document.getElementById("fuelInput");
const tipoViaturaInput = document.getElementById("tipoViaturaInput");
const combustivelHiddenInput = document.querySelector("[data-hidden-combustivel]");
const servidoresSelect = document.getElementById("servidoresSelect");
const viajantesList = document.getElementById("viajantesList");
const motoristaSelect = document.getElementById("motoristaSelect");
const motoristaChipList = document.getElementById("motoristaChipList");
const motoristaNome = document.getElementById("motoristaNome");
const motoristaOficio = document.getElementById("motoristaOficio");
const motoristaProtocolo = document.getElementById("motoristaProtocolo");
const caronaFields = document.getElementById("caronaFields");

function debounce(fn, wait = 260) {
  let timer = null;
  return (...args) => {
    window.clearTimeout(timer);
    timer = window.setTimeout(() => fn(...args), wait);
  };
}

function idsSelecionados() {
  if (servidoresSelect) {
    return Array.from(servidoresSelect.selectedOptions)
      .map((option) => option.value)
      .filter(Boolean);
  }
  return Array.from(
    document.querySelectorAll('input[name="viajantes_ids"], select[name="servidores"] option:checked')
  )
    .map((input) => input.value)
    .filter(Boolean);
}

function getPreviewField(name) {
  return document.querySelector(`[data-preview-field='${name}']`);
}

function setPreviewValue(name, value) {
  const el = getPreviewField(name);
  if (!el) {
    return;
  }
  const text = (value || "").toString().trim();
  el.textContent = text || "-";
}

function updatePreviewFromInputs() {
  document.querySelectorAll("[data-preview-target]").forEach((input) => {
    const name = input.getAttribute("data-preview-target");
    if (!name) {
      return;
    }
    setPreviewValue(name, input.value);
  });
}

function getCustosSelect() {
  return document.querySelector("[data-custos-select]");
}

function getCustosGroup(select) {
  return select?.closest("[data-custos-group]") || document.body;
}

function getCustosInstitutionInput(select) {
  const group = getCustosGroup(select);
  return group.querySelector("[data-custos-instituicao]");
}

function buildCustosPreviewText(select) {
  if (!select) {
    return "";
  }
  const option = select.selectedOptions?.[0];
  const label = option ? option.textContent.trim() : select.value;
  const instituicao = (getCustosInstitutionInput(select)?.value || "").trim();
  if (select.value === "OUTRA_INSTITUICAO" && instituicao) {
    return `${label} – ${instituicao}`;
  }
  return label;
}

function updateCustosPreview() {
  const select = getCustosSelect();
  const preview = buildCustosPreviewText(select);
  setPreviewValue("custos", preview || "-");
}

function toggleCustosInstitutionField(select) {
  const group = getCustosGroup(select);
  const container = group.querySelector("[data-custos-field]");
  if (!container) {
    return;
  }
  const show = select?.value === "OUTRA_INSTITUICAO";
  container.classList.toggle("is-hidden", !show);
  const input = container.querySelector("[data-custos-instituicao]");
  if (input) {
    input.disabled = !show;
    if (!show) {
      input.value = "";
    }
  }
}

function initCustosControls() {
  document.querySelectorAll("[data-custos-select]").forEach((select) => {
    select.addEventListener("change", () => {
      toggleCustosInstitutionField(select);
      updateCustosPreview();
    });
    toggleCustosInstitutionField(select);
  });
  document.querySelectorAll("[data-custos-instituicao]").forEach((input) => {
    input.addEventListener("input", updateCustosPreview);
  });
  updateCustosPreview();
}

function renderPreviewViajantes(viajantes) {
  const container = getPreviewField("viajantes");
  const countEl = getPreviewField("viajantes-count");
  if (countEl) {
    countEl.textContent = String(viajantes.length || 0);
  }
  if (!container) {
    return;
  }
  if (!viajantes.length) {
    container.classList.add("empty");
    container.textContent = "Nenhum servidor selecionado.";
    return;
  }
  container.classList.remove("empty");
  container.innerHTML = viajantes
    .map((v) => {
      const detalhes = [];
      if (v.rg) {
        detalhes.push(`RG: ${v.rg}`);
      }
      if (v.cpf) {
        detalhes.push(`CPF: ${v.cpf}`);
      }
      const sub = detalhes.length
        ? `<span class="preview-sub">${detalhes.join(" | ")}</span>`
        : "";
      return `
        <div class="preview-line">
          <span class="preview-label">${v.nome || "Servidor"}</span>
          <span class="preview-value">${v.cargo || "-"}${sub}</span>
        </div>
      `;
    })
    .join("");
}

async function carregarViajantesPreview() {
  const ids = idsSelecionados();
  if (!ids.length) {
    if (viajantesPreview) {
      viajantesPreview.innerHTML = "<em>Nenhum servidor selecionado.</em>";
    }
    renderPreviewViajantes([]);
    return;
  }

  const response = await fetch(`/api/viajantes/?ids=${ids.join(",")}`);
  const data = await response.json();
  const viajantes = data.viajantes || [];

  if (viajantesPreview) {
    if (!viajantes.length) {
      viajantesPreview.innerHTML = "<em>Nenhum dado encontrado.</em>";
    } else {
      const rows = viajantes
        .map((viajante) => {
          const telefone = viajante.telefone ? ` | Tel: ${viajante.telefone}` : "";
          return `<div><strong>${viajante.nome}</strong> - RG: ${viajante.rg} | CPF: ${viajante.cpf} | Cargo: ${viajante.cargo}${telefone}</div>`;
        })
        .join("");
      viajantesPreview.innerHTML = rows;
    }
  }

  renderPreviewViajantes(viajantes);
}

function filtrarViajantesLocal(termo) {
  if (!viajantesList) {
    return;
  }
  const term = termo.toLowerCase();
  viajantesList.querySelectorAll(".chip").forEach((item) => {
    const texto = item.textContent.toLowerCase();
    item.style.display = texto.includes(term) ? "" : "none";
  });
}

function toggleChipEmptyState() {
  if (!viajantesList) {
    return;
  }
  const hasItems = servidoresSelect
    ? servidoresSelect.selectedOptions.length > 0
    : Boolean(viajantesList.querySelector(".chip"));
  viajantesList.classList.toggle("is-empty", !hasItems);
}

function addServidorChip(viajante) {
  if (!viajante) {
    return;
  }
  const id = String(viajante.id || "");
  if (!id) {
    return;
  }
  if (servidoresSelect) {
    let option = Array.from(servidoresSelect.options).find(
      (item) => String(item.value) === id
    );
    if (!option) {
      option = document.createElement("option");
      option.value = id;
      option.textContent = viajante.nome || viajante.label || "Servidor";
      servidoresSelect.appendChild(option);
    }
    if (option.selected) {
      window.showToast?.("Servidor ja selecionado.", "info");
      return;
    }
    option.selected = true;
    servidoresSelect.dispatchEvent(new Event("change", { bubbles: true }));
    return;
  }
  if (!viajantesList) {
    return;
  }
}

function removeServidorChip(id) {
  if (!id) {
    return;
  }
  if (servidoresSelect) {
    const option = Array.from(servidoresSelect.options).find(
      (item) => String(item.value) === String(id)
    );
    if (option) {
      option.selected = false;
      servidoresSelect.dispatchEvent(new Event("change", { bubbles: true }));
    }
    return;
  }
  if (!viajantesList) {
    return;
  }
}

function renderServidoresChips() {
  if (!viajantesList) {
    return;
  }
  viajantesList.innerHTML = "";
  if (!servidoresSelect) {
    toggleChipEmptyState();
    return;
  }
  const options = Array.from(servidoresSelect.selectedOptions);
  options.forEach((option) => {
    const id = option.value;
    const label = option.textContent?.trim() || "Servidor";
    const chip = document.createElement("span");
    chip.className = "chip";
    chip.dataset.id = id;
    chip.innerHTML = `
      ${label}
      <button type="button" class="chip-remove" data-remove-id="${id}" aria-label="Remover">
        &times;
      </button>
    `;
    viajantesList.appendChild(chip);
  });
  toggleChipEmptyState();
}

function attachSearchSuggestions(input, fetchFn, onSelect, options = {}) {
  if (!input) {
    return;
  }
  const row = input.closest(".field-row") || input.parentElement;
  if (!row) {
    return;
  }
  row.classList.add("has-suggestions");
  const suggestions = document.createElement("div");
  suggestions.className = "search-suggestions";
  row.appendChild(suggestions);

  const { allowEmpty = false, minChars = 1 } = options;
  let items = [];
  let activeIndex = -1;

  function clear() {
    suggestions.innerHTML = "";
    items = [];
    activeIndex = -1;
    row.classList.remove("is-open");
  }

  function render() {
    if (!items.length) {
      suggestions.innerHTML = '<div class="autocomplete-hint">Nenhum resultado.</div>';
      row.classList.add("is-open");
      return;
    }
    suggestions.innerHTML = items
      .map((item, index) => {
        const label = item.label || item.nome || item.placa || "Item";
        const active = index === activeIndex ? " is-active" : "";
        return `<button type="button" class="search-suggestion${active}" data-index="${index}">${label}</button>`;
      })
      .join("");
    row.classList.add("is-open");
  }

  async function load(term) {
    if (!term && !allowEmpty) {
      clear();
      return;
    }
    if (!allowEmpty && term.length < minChars) {
      clear();
      return;
    }
    items = await fetchFn(term);
    activeIndex = items.length ? 0 : -1;
    render();
  }

  const debounced = debounce((term) => {
    load(term).catch(() => clear());
  }, 260);

  input.addEventListener("input", () => {
    const term = input.value.trim();
    debounced(term);
  });

  input.addEventListener("focus", () => {
    const term = input.value.trim();
    if (allowEmpty || term.length >= minChars) {
      debounced(term);
    }
  });

  input.addEventListener("keydown", (event) => {
    if (!items.length) {
      return;
    }
    if (event.key === "ArrowDown") {
      event.preventDefault();
      activeIndex = Math.min(items.length - 1, activeIndex + 1);
      render();
      return;
    }
    if (event.key === "ArrowUp") {
      event.preventDefault();
      activeIndex = Math.max(0, activeIndex - 1);
      render();
      return;
    }
    if (event.key === "Enter") {
      event.preventDefault();
      const item = items[activeIndex];
      if (item) {
        onSelect(item);
      }
      clear();
      return;
    }
    if (event.key === "Escape") {
      clear();
    }
  });

  suggestions.addEventListener("click", (event) => {
    const btn = event.target.closest(".search-suggestion");
    if (!btn) {
      return;
    }
    const index = Number(btn.getAttribute("data-index"));
    const item = items[index];
    if (item) {
      onSelect(item);
    }
    clear();
  });

  document.addEventListener("click", (event) => {
    if (!row.contains(event.target)) {
      clear();
    }
  });
}

async function fetchResults(url, fallback = []) {
  try {
    const response = await fetch(url, {
      headers: { "X-Requested-With": "XMLHttpRequest" },
      credentials: "same-origin",
    });
    if (!response.ok) {
      return fallback;
    }
    const contentType = response.headers.get("content-type") || "";
    if (!contentType.includes("application/json")) {
      return fallback;
    }
    const data = await response.json();
    return data.results || fallback;
  } catch (err) {
    return fallback;
  }
}

async function buscarVeiculos(term) {
  const url = new URL("/api/veiculos/", window.location.origin);
  url.searchParams.set("placa", term);
  return fetchResults(url.toString(), []);
}

function aplicarVeiculo(veiculo) {
  if (!veiculo) {
    return;
  }

  if (placaInput) {
    placaInput.value = veiculo.placa || placaInput.value;
  }
  if (modeloInput) {
    modeloInput.value = veiculo.modelo || modeloInput.value;
  }
  if (combustivelInput) {
    const valor = veiculo.combustivel || combustivelInput.value;
    if (combustivelInput.tagName === "SELECT") {
      const option = Array.from(combustivelInput.options).find(
        (opt) => opt.value === valor
      );
      if (option) {
        combustivelInput.value = valor;
      } else {
        combustivelInput.value = "";
        window.showToast?.("Combustivel fora da lista, selecione manualmente.", "info");
      }
    } else {
      combustivelInput.value = valor;
    }
    if (combustivelHiddenInput) {
      combustivelHiddenInput.value = combustivelInput.value || valor;
    }
    combustivelInput.dispatchEvent(new Event("change", { bubbles: true }));
  }
  if (tipoViaturaInput) {
    tipoViaturaInput.value = veiculo.tipo_viatura || tipoViaturaInput.value;
  }
  updatePreviewFromInputs();
  window.showToast?.("Veiculo selecionado.", "success");
}

async function carregarVeiculoPorPlaca() {
  const placa = placaInput?.value.trim();
  if (!placa) {
    return;
  }
  const response = await fetch(`/api/veiculo/?plate=${encodeURIComponent(placa)}`);
  const data = await response.json();
  if (data.found) {
    aplicarVeiculo(data);
  }
}

let currentMotorista = null;

function isMotoristaCarona() {
  const motoristaId = motoristaSelect?.value || "";
  const manual = motoristaNome?.value.trim() || "";
  if (!motoristaId && !manual) {
    return false;
  }
  if (motoristaId) {
    return !idsSelecionados().includes(motoristaId);
  }
  return true;
}

function updateCaronaVisibility() {
  if (!caronaFields) {
    return;
  }
  const isCarona = isMotoristaCarona();
  caronaFields.classList.toggle("is-visible", isCarona);
  const refSelect = caronaFields.querySelector("[data-carona-ref]");
  if (refSelect) {
    refSelect.disabled = !isCarona;
    if (!isCarona) {
      refSelect.value = "";
    }
  }
  document.querySelectorAll("[data-carona-preview]").forEach((el) => {
    el.classList.toggle("is-hidden", !isCarona);
  });
}

function renderMotoristaChip(motorista) {
  if (!motoristaChipList) {
    return;
  }
  motoristaChipList.innerHTML = "";
  if (!motorista) {
    return;
  }
  const id = motorista.id ? String(motorista.id) : "manual";
  const chip = document.createElement("span");
  chip.className = "chip";
  chip.dataset.id = id;
  chip.innerHTML = `
    ${motorista.nome || motorista.label || "Motorista"}
    <button type="button" class="chip-remove" data-remove-motorista="1" aria-label="Remover">
      &times;
    </button>
  `;
  motoristaChipList.appendChild(chip);
}

function setMotorista(motorista) {
  if (!motorista) {
    return;
  }
  currentMotorista = motorista;
  if (motoristaSelect) {
    const id = motorista.id ? String(motorista.id) : "";
    if (id) {
      let option = Array.from(motoristaSelect.options).find(
        (item) => String(item.value) === id
      );
      if (!option) {
        option = document.createElement("option");
        option.value = id;
        option.textContent = motorista.nome || motorista.label || "Motorista";
        motoristaSelect.appendChild(option);
      }
      motoristaSelect.value = id;
      window.syncAutocompleteDisplay?.(motoristaSelect);
    } else {
      motoristaSelect.value = "";
      window.syncAutocompleteDisplay?.(motoristaSelect);
    }
  }
  if (motoristaNome) {
    motoristaNome.value = "";
  }
  renderMotoristaChip(motorista);
  updateMotoristaPreview();
}

function clearMotorista() {
  currentMotorista = null;
  if (motoristaSelect) {
    motoristaSelect.value = "";
    window.syncAutocompleteDisplay?.(motoristaSelect);
  }
  if (motoristaChipList) {
    motoristaChipList.innerHTML = "";
  }
  updateMotoristaPreview();
}

function updateMotoristaPreview() {
  if (!motoristaNome) {
    return;
  }
  const manual = motoristaNome.value.trim();
  const isCarona = isMotoristaCarona();
  const nome = manual || currentMotorista?.nome || currentMotorista?.label || "";
  const suffix = nome && isCarona ? " (Carona)" : "";
  setPreviewValue("motorista_nome", nome ? `${nome}${suffix}` : "-");
  setPreviewValue("motorista_cpf", currentMotorista?.cpf || "-");
  setPreviewValue("motorista_rg", currentMotorista?.rg || "-");
  setPreviewValue("motorista_cargo", currentMotorista?.cargo || "-");
  setPreviewValue("motorista_oficio", motoristaOficio?.value || "-");
  setPreviewValue("motorista_protocolo", motoristaProtocolo?.value || "-");
  updateCaronaVisibility();
}

function initBindings() {
  if (idsSelecionados().length) {
    carregarViajantesPreview();
  }
  renderServidoresChips();
  toggleChipEmptyState();

  document.querySelectorAll("[data-preview-target]").forEach((input) => {
    input.addEventListener("input", updatePreviewFromInputs);
    input.addEventListener("change", updatePreviewFromInputs);
  });
  updatePreviewFromInputs();
  initCustosControls();

  if (motoristaNome) {
    motoristaNome.addEventListener("input", () => {
      const nomeManual = motoristaNome.value.trim();
      if (motoristaSelect?.value) {
        motoristaSelect.value = "";
        window.syncAutocompleteDisplay?.(motoristaSelect);
        currentMotorista = null;
        if (nomeManual) {
          renderMotoristaChip({ nome: nomeManual });
        } else {
          renderMotoristaChip(null);
        }
        updateMotoristaPreview();
        return;
      }
      if (!nomeManual) {
        renderMotoristaChip(null);
        updateMotoristaPreview();
        return;
      }
      renderMotoristaChip({ nome: nomeManual });
      updateMotoristaPreview();
    });
  }
  motoristaOficio?.addEventListener("input", updateMotoristaPreview);
  motoristaProtocolo?.addEventListener("input", updateMotoristaPreview);
  updateMotoristaPreview();

  if (viajantesList) {
    viajantesList.addEventListener("click", (event) => {
      const btn = event.target.closest("[data-remove-id]");
      if (!btn) {
        return;
      }
      const id = btn.getAttribute("data-remove-id");
      removeServidorChip(id);
      updateMotoristaPreview();
    });
  }

  if (servidoresSelect) {
    servidoresSelect.addEventListener("change", () => {
      renderServidoresChips();
      carregarViajantesPreview();
      updateMotoristaPreview();
    });
  }

  if (motoristaChipList) {
    motoristaChipList.addEventListener("click", (event) => {
      const btn = event.target.closest("[data-remove-motorista]");
      if (!btn) {
        return;
      }
      if (motoristaNome) {
        motoristaNome.value = "";
      }
      clearMotorista();
    });
  }

  if (motoristaSelect) {
    motoristaSelect.addEventListener("change", () => {
      const motoristaId = motoristaSelect.value || "";
      if (!motoristaId) {
        currentMotorista = null;
        if (!motoristaNome?.value.trim()) {
          renderMotoristaChip(null);
        }
        updateMotoristaPreview();
        return;
      }
      if (motoristaNome) {
        motoristaNome.value = "";
      }
      fetch(`/api/motoristas/${motoristaId}/`)
        .then((response) => response.json())
        .then((data) => {
          if (data && data.id) {
            currentMotorista = data;
            renderMotoristaChip(data);
            updateMotoristaPreview();
          }
        })
        .catch(() => {
          currentMotorista = null;
          updateMotoristaPreview();
        });
    });
  }

  if (placaInput) {
    attachSearchSuggestions(placaInput, buscarVeiculos, aplicarVeiculo, { minChars: 2 });
    placaInput.addEventListener("blur", () => {
      carregarVeiculoPorPlaca().catch(() => {});
    });
  }

  if (motoristaSelect && motoristaSelect.value) {
    fetch(`/api/motoristas/${motoristaSelect.value}/`)
      .then((response) => response.json())
      .then((data) => {
        if (data && data.id) {
          currentMotorista = data;
          renderMotoristaChip(data);
          updateMotoristaPreview();
        }
      })
      .catch(() => {});
  }
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initBindings);
} else {
  initBindings();
}

window.addServidorChip = addServidorChip;
window.setMotoristaFromViajante = setMotorista;
