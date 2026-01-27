const viajantesPreview = document.getElementById("viajantesPreview");
const placaInput = document.getElementById("plateInput");
const modeloInput = document.getElementById("modelInput");
const combustivelInput = document.getElementById("fuelInput");
const viajantesSearch = document.getElementById("viajantesSearch");
const viajantesList = document.getElementById("viajantesList");
const motoristaSelect = document.getElementById("motoristaSelect");
const motoristaNome = document.getElementById("motoristaNome");

function debounce(fn, wait = 260) {
  let timer = null;
  return (...args) => {
    window.clearTimeout(timer);
    timer = window.setTimeout(() => fn(...args), wait);
  };
}

function idsSelecionados() {
  return Array.from(document.querySelectorAll('input[name="viajantes_ids"]:checked')).map(
    (checkbox) => checkbox.value
  );
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
  viajantesList.querySelectorAll(".checkbox-item").forEach((item) => {
    const texto = item.textContent.toLowerCase();
    item.style.display = texto.includes(term) ? "" : "none";
  });
}

function ensureCheckbox(viajante) {
  if (!viajantesList || !viajante) {
    return;
  }
  const existing = viajantesList.querySelector(`input[value='${viajante.id}']`);
  if (existing) {
    existing.checked = true;
    existing.dispatchEvent(new Event("change", { bubbles: true }));
    return;
  }
  const label = document.createElement("label");
  label.className = "checkbox-item";
  label.innerHTML = `
    <input type="checkbox" name="viajantes_ids" value="${viajante.id}" checked />
    <span>${viajante.nome || viajante.label || "Novo viajante"}</span>
  `;
  viajantesList.prepend(label);
  const input = label.querySelector("input");
  input?.dispatchEvent(new Event("change", { bubbles: true }));
}

function attachSearchSuggestions(input, fetchFn, onSelect) {
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
      clear();
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
    if (!term) {
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

async function buscarServidores(term) {
  const url = new URL("/api/servidores/", window.location.origin);
  url.searchParams.set("q", term);
  const response = await fetch(url.toString(), {
    headers: { "X-Requested-With": "XMLHttpRequest" },
  });
  const data = await response.json();
  return data.results || [];
}

async function buscarVeiculos(term) {
  const url = new URL("/api/veiculos/", window.location.origin);
  url.searchParams.set("placa", term);
  const response = await fetch(url.toString(), {
    headers: { "X-Requested-With": "XMLHttpRequest" },
  });
  const data = await response.json();
  return data.results || [];
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
    combustivelInput.value = veiculo.combustivel || combustivelInput.value;
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

function updateMotoristaPreview() {
  if (!motoristaSelect && !motoristaNome) {
    return;
  }
  const manual = motoristaNome?.value.trim() || "";
  if (manual) {
    setPreviewValue("motorista_nome", manual);
    setPreviewValue("motorista_cpf", "-");
    setPreviewValue("motorista_rg", "-");
    setPreviewValue("motorista_cargo", "-");
    return;
  }
  const selected = motoristaSelect?.selectedOptions?.[0];
  const nome = selected && selected.value ? selected.textContent.trim() : "";
  setPreviewValue("motorista_nome", nome);
  setPreviewValue("motorista_cpf", selected?.dataset.cpf || "");
  setPreviewValue("motorista_rg", selected?.dataset.rg || "");
  setPreviewValue("motorista_cargo", selected?.dataset.cargo || "");
}

function initBindings() {
  document.addEventListener("change", (event) => {
    if (event.target.matches('input[name="viajantes_ids"]')) {
      carregarViajantesPreview();
    }
  });
  if (document.querySelector('input[name="viajantes_ids"]')) {
    carregarViajantesPreview();
  }

  document.querySelectorAll("[data-preview-target]").forEach((input) => {
    input.addEventListener("input", updatePreviewFromInputs);
    input.addEventListener("change", updatePreviewFromInputs);
  });
  updatePreviewFromInputs();

  if (motoristaSelect) {
    motoristaSelect.addEventListener("change", updateMotoristaPreview);
  }
  if (motoristaNome) {
    motoristaNome.addEventListener("input", updateMotoristaPreview);
  }
  updateMotoristaPreview();

  if (viajantesSearch && viajantesList) {
    viajantesSearch.addEventListener("input", () => {
      const termo = viajantesSearch.value.trim();
      filtrarViajantesLocal(termo);
    });

    attachSearchSuggestions(
      viajantesSearch,
      buscarServidores,
      (servidor) => {
        ensureCheckbox(servidor);
        viajantesSearch.value = servidor.nome || servidor.label || "";
        filtrarViajantesLocal("");
      }
    );
  }

  if (placaInput) {
    attachSearchSuggestions(placaInput, buscarVeiculos, aplicarVeiculo);
    placaInput.addEventListener("blur", () => {
      carregarVeiculoPorPlaca().catch(() => {});
    });
  }
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initBindings);
} else {
  initBindings();
}
