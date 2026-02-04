function onlyDigits(value) {
  return (value || "").replace(/\D/g, "");
}

function maskCPF(value) {
  const digits = onlyDigits(value).slice(0, 11);
  if (digits.length <= 3) return digits;
  if (digits.length <= 6) return `${digits.slice(0, 3)}.${digits.slice(3)}`;
  if (digits.length <= 9)
    return `${digits.slice(0, 3)}.${digits.slice(3, 6)}.${digits.slice(6)}`;
  return `${digits.slice(0, 3)}.${digits.slice(3, 6)}.${digits.slice(6, 9)}-${digits.slice(9)}`;
}

function maskRG(value) {
  const digits = onlyDigits(value).slice(0, 9);
  if (digits.length <= 2) return digits;
  if (digits.length <= 5) return `${digits.slice(0, 2)}.${digits.slice(2)}`;
  if (digits.length <= 8)
    return `${digits.slice(0, 2)}.${digits.slice(2, 5)}.${digits.slice(5)}`;
  return `${digits.slice(0, 2)}.${digits.slice(2, 5)}.${digits.slice(5, 8)}-${digits.slice(8)}`;
}

function maskTelefone(value) {
  const digits = onlyDigits(value).slice(0, 11);
  if (digits.length <= 2) return digits;
  if (digits.length <= 6) return `(${digits.slice(0, 2)}) ${digits.slice(2)}`;
  if (digits.length <= 10)
    return `(${digits.slice(0, 2)}) ${digits.slice(2, 6)}-${digits.slice(6)}`;
  return `(${digits.slice(0, 2)}) ${digits.slice(2, 7)}-${digits.slice(7)}`;
}

function applyMasks(root) {
  root.querySelectorAll("[data-mask='cpf']").forEach((input) => {
    input.addEventListener("input", () => {
      input.value = maskCPF(input.value);
    });
  });

  root.querySelectorAll("[data-mask='rg']").forEach((input) => {
    input.addEventListener("input", () => {
      input.value = maskRG(input.value);
    });
  });

  root.querySelectorAll("[data-mask='telefone']").forEach((input) => {
    input.addEventListener("input", () => {
      input.value = maskTelefone(input.value);
    });
  });

  root.querySelectorAll("[data-uppercase='true']").forEach((input) => {
    input.addEventListener("input", () => {
      input.value = input.value.toUpperCase();
    });
  });
}

function getCookie(name) {
  const value = `; ${document.cookie}`;
  const parts = value.split(`; ${name}=`);
  if (parts.length === 2) {
    return parts.pop().split(";").shift();
  }
  return "";
}

function ensureCargoOption(select, nome) {
  if (!select || !nome) return;
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
  applyMasks(root);
  initCargoCreate();
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initViajanteForm);
} else {
  initViajanteForm();
}
