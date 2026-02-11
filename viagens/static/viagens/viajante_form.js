function getCookie(name) {
  const value = `; ${document.cookie}`;
  const parts = value.split(`; ${name}=`);
  if (parts.length === 2) {
    return parts.pop().split(";").shift();
  }
  return "";
}

function normalizeCargoKey(value) {
  return (value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function rebuildCargoOptions(select) {
  if (!select) return;
  const current = select.value || "";
  const seen = new Set();
  const options = [];
  let placeholderText = null;

  Array.from(select.options).forEach((option) => {
    const value = (option.value || "").trim();
    const label = (option.textContent || "").trim() || value;
    if (!value) {
      if (placeholderText === null) {
        placeholderText = label || "Selecione";
      }
      return;
    }
    const key = normalizeCargoKey(value || label);
    if (!key || seen.has(key)) return;
    seen.add(key);
    options.push({ value, label });
  });

  select.innerHTML = "";
  if (placeholderText !== null) {
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = placeholderText;
    select.appendChild(placeholder);
  }
  options.forEach((opt) => {
    const option = document.createElement("option");
    option.value = opt.value;
    option.textContent = opt.label;
    select.appendChild(option);
  });

  if (current) {
    const match = Array.from(select.options).find(
      (option) => normalizeCargoKey(option.value) === normalizeCargoKey(current)
    );
    if (match) {
      select.value = match.value;
    }
  }
}

function ensureCargoOption(select, nome) {
  if (!select || !nome) return;
  rebuildCargoOptions(select);
  const exists = Array.from(select.options).find(
    (option) => option.value.toLowerCase() === nome.toLowerCase()
  );
  if (exists) {
    select.value = exists.value;
    select.dispatchEvent(new Event("change", { bubbles: true }));
    return;
  }
  const option = document.createElement("option");
  option.value = nome;
  option.textContent = nome;
  select.appendChild(option);
  select.value = nome;
  select.dispatchEvent(new Event("change", { bubbles: true }));
}

let cargoListenerAttached = false;

function initCargoCreate() {
  if (cargoListenerAttached) return;
  cargoListenerAttached = true;
  document.addEventListener("click", async (event) => {
    const button = event.target.closest("[data-add-cargo]");
    if (!button) return;
    const container =
      button.closest("label") || button.closest(".field-row") || button.parentElement;
    if (!container) return;
    const select = container.querySelector("[data-cargo-select]");
    if (!select) return;
    const nome = window.prompt("Informe o novo cargo:");
    if (!nome) return;
    const trimmed = nome.trim();
    if (!trimmed) return;
    try {
      const response = await fetch("/cargos/criar/", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "X-CSRFToken": getCookie("csrftoken"),
          "X-Requested-With": "XMLHttpRequest",
        },
        body: JSON.stringify({ nome: trimmed }),
      });
      if (!response.ok) {
        const error = await response.json().catch(() => ({}));
        const message = error.error || "Nao foi possivel criar o cargo.";
        window.showToast?.(message, "error");
        return;
      }
      const data = await response.json();
      ensureCargoOption(select, data.nome || trimmed);
      const cargoNovo = container.querySelector("[data-cargo-novo]");
      if (cargoNovo) {
        cargoNovo.value = "";
      }
      window.showToast?.("Cargo criado com sucesso.", "success");
    } catch (err) {
      window.showToast?.("Nao foi possivel criar o cargo.", "error");
    }
  });
}

function initViajanteForm() {
  const root = document;
  window.initInputMasks?.(root);
  root.querySelectorAll("[data-cargo-select]").forEach((select) => {
    rebuildCargoOptions(select);
  });
  initCargoCreate();
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initViajanteForm);
} else {
  initViajanteForm();
}
