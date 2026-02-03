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

function initCargoToggle(root) {
  root.querySelectorAll("[data-add-cargo]").forEach((button) => {
    button.addEventListener("click", () => {
      const container = button.closest("label");
      if (!container) return;
      const select = container.querySelector("[data-cargo-select]");
      const input = container.querySelector("[data-cargo-novo]");
      if (!select || !input) return;
      const isHidden = input.hasAttribute("hidden");
      if (isHidden) {
        input.removeAttribute("hidden");
        input.focus();
        select.value = "";
      } else {
        input.setAttribute("hidden", "hidden");
        input.value = "";
      }
    });
  });
}

function initViajanteForm() {
  const root = document;
  applyMasks(root);
  initCargoToggle(root);
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initViajanteForm);
} else {
  initViajanteForm();
}
