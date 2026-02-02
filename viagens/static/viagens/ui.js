(function () {
  const toastRoot = document.getElementById("toastRoot");

  function showToast(message, type = "info") {
    if (!toastRoot) {
      return;
    }
    const toast = document.createElement("div");
    toast.className = `toast toast-${type}`;
    toast.textContent = message;
    toastRoot.appendChild(toast);
    requestAnimationFrame(() => toast.classList.add("is-visible"));
    setTimeout(() => {
      toast.classList.remove("is-visible");
      setTimeout(() => toast.remove(), 220);
    }, 2800);
  }

  window.showToast = showToast;
})();

(function () {
  const modalRoot = document.getElementById("modalRoot");
  if (!modalRoot) {
    return;
  }

  function closeModal() {
    modalRoot.innerHTML = "";
    document.body.classList.remove("modal-open");
  }

  function addOptionToSelect(select, item) {
    if (!select || !item) {
      return;
    }
    const exists = Array.from(select.options).find(
      (option) => String(option.value) === String(item.id)
    );
    const option = exists || document.createElement("option");
    option.value = String(item.id);
    option.textContent =
      item.text || item.label || item.nome || item.placa || "Item";
    if (item.cpf) {
      option.dataset.cpf = item.cpf;
    }
    if (item.rg) {
      option.dataset.rg = item.rg;
    }
    if (item.cargo) {
      option.dataset.cargo = item.cargo;
    }
    if (!exists) {
      select.appendChild(option);
    }
    if (select.multiple) {
      option.selected = true;
    } else {
      select.value = String(item.id);
    }
    select.dispatchEvent(new Event("change", { bubbles: true }));
  }

  function addCheckboxToList(listEl, item) {
    if (!listEl || !item) {
      return;
    }
    if (window.addServidorChip) {
      window.addServidorChip(item);
      return;
    }
    const existing = listEl.querySelector(`[data-hidden-id='${item.id}']`);
    if (existing) {
      return;
    }
    const chip = document.createElement("span");
    chip.className = "chip";
    chip.dataset.id = String(item.id);
    chip.innerHTML = `
      ${item.nome || item.label || "Novo viajante"}
      <button type="button" class="chip-remove" data-remove-id="${item.id}" aria-label="Remover">
        &times;
      </button>
    `;
    const input = document.createElement("input");
    input.type = "hidden";
    input.name = "viajantes_ids";
    input.value = String(item.id);
    input.dataset.hiddenId = String(item.id);
    listEl.appendChild(chip);
    listEl.appendChild(input);
  }

  function handleModalSuccess(trigger, item) {
    const kind = trigger?.getAttribute("data-modal-kind") || "";

    if (kind === "viajante") {
      const listSelector = trigger.getAttribute("data-target-list");
      const selectSelector = trigger.getAttribute("data-target-select");
      const motoristaTarget = trigger.getAttribute("data-target-motorista");
      if (listSelector) {
        addCheckboxToList(document.querySelector(listSelector), item);
      }
      if (selectSelector) {
        addOptionToSelect(document.querySelector(selectSelector), item);
      }
      if (motoristaTarget && window.setMotoristaFromViajante) {
        window.setMotoristaFromViajante(item);
      }
      window.showToast?.("Viajante cadastrado com sucesso.", "success");
      return;
    }

    if (kind === "veiculo") {
      const plateId = trigger.getAttribute("data-target-plate");
      const modelId = trigger.getAttribute("data-target-model");
      const fuelId = trigger.getAttribute("data-target-fuel");
      if (plateId) {
        const plate = document.getElementById(plateId);
        if (plate) {
          plate.value = item.placa || "";
          plate.dispatchEvent(new Event("input", { bubbles: true }));
        }
      }
      if (modelId) {
        const model = document.getElementById(modelId);
        if (model) {
          model.value = item.modelo || "";
          model.dispatchEvent(new Event("input", { bubbles: true }));
        }
      }
      if (fuelId) {
        const fuel = document.getElementById(fuelId);
        if (fuel) {
          const valor = item.combustivel || "";
          if (fuel.tagName === "SELECT") {
            const option = Array.from(fuel.options).find(
              (opt) => opt.value === valor
            );
            if (option) {
              fuel.value = valor;
            } else {
              fuel.value = "";
              window.showToast?.("Combustivel fora da lista, selecione manualmente.", "info");
            }
            fuel.dispatchEvent(new Event("change", { bubbles: true }));
          } else {
            fuel.value = valor;
            fuel.dispatchEvent(new Event("input", { bubbles: true }));
          }
        }
      }
      window.showToast?.("Veiculo cadastrado e preenchido.", "success");
    }
  }

  function bindModalEvents(modalEl, trigger) {
    modalEl.querySelectorAll("[data-modal-close]").forEach((btn) => {
      btn.addEventListener("click", closeModal);
    });

    modalEl.addEventListener("click", (event) => {
      if (event.target === modalEl) {
        closeModal();
      }
    });

    const form = modalEl.querySelector("form");
    if (!form) {
      return;
    }

    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      const submitBtn = form.querySelector("[type='submit']");
      submitBtn?.setAttribute("data-loading", "true");

      const response = await fetch(form.action, {
        method: "POST",
        headers: {
          "X-Requested-With": "XMLHttpRequest",
        },
        body: new FormData(form),
      });

      const contentType = response.headers.get("content-type") || "";
      if (contentType.includes("application/json")) {
        const data = await response.json();
        if (data.success) {
          handleModalSuccess(trigger, data.item);
          closeModal();
        } else {
          submitBtn?.removeAttribute("data-loading");
        }
        return;
      }

      const html = await response.text();
      modalRoot.innerHTML = html;
      const newModal = modalRoot.querySelector("[data-modal]");
      if (newModal) {
        bindModalEvents(newModal, trigger);
      }
    });
  }

  async function openModal(trigger) {
    const url = trigger.getAttribute("data-modal-url");
    if (!url) {
      return;
    }
    modalRoot.innerHTML =
      '<div class="modal-backdrop" data-modal><div class="modal-card"><div class="modal-loading">Carregando...</div></div></div>';
    document.body.classList.add("modal-open");
    try {
      const response = await fetch(url, {
        headers: { "X-Requested-With": "XMLHttpRequest" },
      });
      const html = await response.text();
      modalRoot.innerHTML = html;
      const modalEl = modalRoot.querySelector("[data-modal]");
      if (modalEl) {
        bindModalEvents(modalEl, trigger);
      }
    } catch (err) {
      modalRoot.innerHTML = "";
      document.body.classList.remove("modal-open");
      window.showToast?.("Nao foi possivel abrir o modal.", "error");
    }
  }

  document.addEventListener("click", (event) => {
    const trigger = event.target.closest("[data-modal-url]");
    if (!trigger) {
      return;
    }
    event.preventDefault();
    openModal(trigger);
  });
})();

(function () {
  if (document.querySelector(".fab-stack")) {
    document.body.classList.add("has-fab");
  }
  const backToTop = document.createElement("button");
  backToTop.type = "button";
  backToTop.className = "back-to-top";
  backToTop.setAttribute("aria-label", "Voltar ao topo");
  backToTop.innerHTML = "&#8593;";
  document.body.appendChild(backToTop);

  const onScroll = () => {
    const show = window.scrollY > 320;
    backToTop.classList.toggle("is-visible", show);
  };

  backToTop.addEventListener("click", () => {
    window.scrollTo({ top: 0, behavior: "smooth" });
  });

  window.addEventListener("scroll", onScroll, { passive: true });
  onScroll();
})();

(function () {
  function openConfirm(modal) {
    if (!modal) {
      return;
    }
    modal.classList.add("is-visible");
    document.body.classList.add("modal-open");
  }

  function closeConfirm(modal) {
    if (!modal) {
      return;
    }
    modal.classList.remove("is-visible");
    document.body.classList.remove("modal-open");
  }

  document.addEventListener("click", (event) => {
    const openBtn = event.target.closest("[data-confirm-open]");
    if (openBtn) {
      const selector = openBtn.getAttribute("data-confirm-open");
      const modal = selector ? document.querySelector(selector) : null;
      openConfirm(modal);
      return;
    }

    const closeBtn = event.target.closest("[data-confirm-close]");
    if (closeBtn) {
      const modal = closeBtn.closest("[data-confirm-modal]");
      closeConfirm(modal);
      return;
    }

    const backdrop = event.target.closest("[data-confirm-modal]");
    if (backdrop && event.target === backdrop) {
      closeConfirm(backdrop);
    }
  });
})();

(function () {
  function ensureGotoInput(form) {
    let input = form.querySelector("input[name='goto_step']");
    if (!input) {
      input = document.createElement("input");
      input.type = "hidden";
      input.name = "goto_step";
      form.appendChild(input);
    }
    return input;
  }

  function bindStepper(stepper) {
    const formId = stepper.getAttribute("data-stepper-form");
    const form = formId ? document.getElementById(formId) : null;
    if (!form) {
      return;
    }
    const gotoInput = ensureGotoInput(form);
    stepper.querySelectorAll("[data-step]").forEach((btn) => {
      btn.addEventListener("click", (event) => {
        const step = btn.getAttribute("data-step");
        const mode = btn.getAttribute("data-step-mode");
        if (!step || mode === "link") {
          return;
        }
        event.preventDefault();
        gotoInput.value = step;
        if (typeof form.requestSubmit === "function") {
          form.requestSubmit();
          return;
        }
        if (typeof form.checkValidity === "function" && !form.checkValidity()) {
          form.reportValidity?.();
          return;
        }
        form.submit();
      });
    });
  }

  document.querySelectorAll("[data-stepper-form]").forEach((stepper) => bindStepper(stepper));
})();

(function () {
  const debounce = (fn, wait = 250) => {
    let timer = null;
    return (...args) => {
      window.clearTimeout(timer);
      timer = window.setTimeout(() => fn(...args), wait);
    };
  };

  function createAutocomplete(select) {
    if (select.dataset.autocompleteReady === "1") {
      return;
    }
    select.dataset.autocompleteReady = "1";
    const url = select.getAttribute("data-autocomplete-url");
    if (!url) {
      return;
    }
    const wrapper = document.createElement("div");
    wrapper.className = "autocomplete";
    const input = document.createElement("input");
    input.type = "text";
    input.className = "autocomplete-input";
    input.setAttribute("autocomplete", "off");
    input.setAttribute("aria-expanded", "false");
    input.setAttribute("role", "combobox");
    const list = document.createElement("div");
    list.className = "autocomplete-list";
    list.setAttribute("role", "listbox");

    select.classList.add("autocomplete-select");
    select.parentElement?.insertBefore(wrapper, select);
    wrapper.appendChild(input);
    wrapper.appendChild(list);
    wrapper.appendChild(select);
    select._autocompleteInput = input;

    const type = select.getAttribute("data-autocomplete-type") || "";
    const isMultiple = select.multiple;
    let items = [];
    let activeIndex = -1;

    function relatedUf() {
      if (type !== "cidade") {
        return "";
      }
      const role = select.getAttribute("data-role") || "";
      const card = select.closest(".trecho-card");
      if (!card) {
        return "";
      }
      if (role.includes("origem")) {
        return card.querySelector("[data-role='origem-estado']")?.value || "";
      }
      return card.querySelector("[data-role='destino-estado']")?.value || "";
    }

    function setExpanded(expanded) {
      input.setAttribute("aria-expanded", expanded ? "true" : "false");
      wrapper.classList.toggle("is-open", expanded);
    }

    function clearList() {
      list.innerHTML = "";
      items = [];
      activeIndex = -1;
      setExpanded(false);
    }

    function selectItem(item) {
      if (!item) {
        return;
      }
      const value = type === "uf" ? item.sigla || item.id : item.id;
      const label =
        type === "uf"
          ? item.label || item.sigla || item.id || ""
          : item.text || item.label || item.nome || "";
      if (isMultiple) {
        if (!value) {
          return;
        }
        const exists = Array.from(select.options).find(
          (option) => String(option.value) === String(value)
        );
        const selectedAlready = Array.from(select.selectedOptions).some(
          (option) => String(option.value) === String(value)
        );
        if (selectedAlready) {
          clearList();
          input.value = "";
          return;
        }
        const option = exists || document.createElement("option");
        option.value = String(value || "");
        option.textContent = label;
        if (item.cpf) {
          option.dataset.cpf = item.cpf;
        }
        if (item.rg) {
          option.dataset.rg = item.rg;
        }
        if (item.cargo) {
          option.dataset.cargo = item.cargo;
        }
        if (!exists) {
          select.appendChild(option);
        }
        option.selected = true;
        select.dispatchEvent(new Event("change", { bubbles: true }));
        input.value = "";
        clearList();
        return;
      }
      const exists = Array.from(select.options).find(
        (option) => String(option.value) === String(value)
      );
      if (!exists) {
        const option = document.createElement("option");
        option.value = String(value || "");
        option.textContent = label;
        if (item.cpf) {
          option.dataset.cpf = item.cpf;
        }
        if (item.rg) {
          option.dataset.rg = item.rg;
        }
        if (item.cargo) {
          option.dataset.cargo = item.cargo;
        }
        select.appendChild(option);
      }
      select.value = String(value || "");
      select.dataset.selected = String(value || "");
      select.dispatchEvent(new Event("change", { bubbles: true }));
      input.value = label;
      clearList();
    }

    function renderItems() {
      if (!items.length) {
        clearList();
        return;
      }
      const selectedValues = isMultiple
        ? new Set(
            Array.from(select.selectedOptions).map((option) => String(option.value))
          )
        : null;
      list.innerHTML = items
        .map((item, index) => {
          const label =
            item.text || item.label || item.nome || item.sigla || item.id;
          const isSelected =
            isMultiple && selectedValues?.has(String(item.id));
          const activeClass = index === activeIndex ? " is-active" : "";
          const disabledAttr = isSelected ? " disabled aria-disabled=\"true\"" : "";
          const selectedClass = isSelected ? " is-selected" : "";
          return `<button type="button" class="autocomplete-item${activeClass}${selectedClass}" data-index="${index}" role="option"${disabledAttr}>${label}</button>`;
        })
        .join("");
      setExpanded(true);
    }

    async function fetchItems(term) {
      const uf = relatedUf();
      if (type === "cidade" && !uf) {
        list.innerHTML =
          '<div class="autocomplete-hint">Selecione a UF primeiro.</div>';
        setExpanded(true);
        return;
      }
      const urlObj = new URL(url, window.location.origin);
      urlObj.searchParams.set("q", term);
      if (type === "cidade") {
        urlObj.searchParams.set("uf", uf);
      }
      const response = await fetch(urlObj.toString(), {
        headers: { "X-Requested-With": "XMLHttpRequest" },
      });
      const data = await response.json();
      items = data.results || [];
      activeIndex = items.length ? 0 : -1;
      renderItems();
    }

    const debouncedFetch = debounce((term) => {
      fetchItems(term).catch(() => {
        clearList();
      });
    }, 260);

    input.addEventListener("focus", () => {
      if (!isMultiple) {
        const selectedOption = select.selectedOptions[0];
        if (selectedOption && selectedOption.value) {
          input.value = selectedOption.textContent.trim();
        }
      }
      debouncedFetch(input.value.trim());
    });

    input.addEventListener("input", () => {
      debouncedFetch(input.value.trim());
    });

    input.addEventListener("keydown", (event) => {
      if (!items.length) {
        return;
      }
      if (event.key === "ArrowDown") {
        event.preventDefault();
        activeIndex = Math.min(items.length - 1, activeIndex + 1);
        renderItems();
        return;
      }
      if (event.key === "ArrowUp") {
        event.preventDefault();
        activeIndex = Math.max(0, activeIndex - 1);
        renderItems();
        return;
      }
      if (event.key === "Enter") {
        event.preventDefault();
        selectItem(items[activeIndex]);
        return;
      }
      if (event.key === "Escape") {
        clearList();
      }
    });

    list.addEventListener("click", (event) => {
      const btn = event.target.closest(".autocomplete-item");
      if (!btn) {
        return;
      }
      if (btn.hasAttribute("disabled")) {
        return;
      }
      const index = Number(btn.getAttribute("data-index"));
      selectItem(items[index]);
    });

    document.addEventListener("click", (event) => {
      if (!wrapper.contains(event.target)) {
        clearList();
      }
    });

    select.addEventListener("change", () => {
      if (isMultiple) {
        input.value = "";
        return;
      }
      const selectedOption = select.selectedOptions[0];
      input.value =
        selectedOption && selectedOption.value
          ? selectedOption.textContent.trim()
          : "";
    });

    const selected = select.selectedOptions[0];
    if (selected && selected.value && !isMultiple) {
      input.value = selected.textContent.trim();
    }
  }

  function syncAutocompleteDisplay(select) {
    if (!select) {
      return;
    }
    const input =
      select._autocompleteInput ||
      select.parentElement?.querySelector(".autocomplete-input");
    if (!input) {
      return;
    }
    if (select.multiple) {
      input.value = "";
      return;
    }
    const option = select.selectedOptions?.[0];
    input.value = option && option.value ? option.textContent.trim() : "";
  }

  function initializeAutocompleteSelects(root = document) {
    root.querySelectorAll("select[data-autocomplete-url]").forEach((select) => {
      createAutocomplete(select);
    });
  }

  window.syncAutocompleteDisplay = syncAutocompleteDisplay;
  window.initializeAutocompleteSelects = initializeAutocompleteSelects;
  initializeAutocompleteSelects(document);
})();
