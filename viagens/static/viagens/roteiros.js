(function () {
  function closeModalElement(modalElement) {
    if (!modalElement) {
      return;
    }
    const backdrop = modalElement.closest("[data-modal]") || modalElement;
    backdrop.remove();
    document.body.classList.remove("modal-open");
  }

  class RoteiroSelector {
    constructor(config) {
      this.oficioId = config?.oficioId;
      this.searchInput = document.getElementById("roteiro-search-input");
      this.searchResultsContainer = document.getElementById("roteiro-search-results");
      this.selectRoteiroCallback = config?.selectRoteiroCallback;
      this.modalElement = document.getElementById("modalSelecionarRoteiro");
      this.debounceTimer = null;
      this.initialized = false;
    }

    init() {
      if (this.initialized || !this.searchInput || !this.searchResultsContainer) {
        return;
      }
      this.initialized = true;
      this.searchInput.addEventListener("input", () => {
        window.clearTimeout(this.debounceTimer);
        this.debounceTimer = window.setTimeout(() => {
          this.buscarRoteiros(this.searchInput.value);
        }, 300);
      });
    }

    async buscarRoteiros(query) {
      const normalizedQuery = String(query || "").trim();
      if (normalizedQuery.length < 3) {
        this.searchResultsContainer.innerHTML =
          '<div class="empty">Digite ao menos 3 caracteres para buscar roteiros.</div>';
        return;
      }

      try {
        const response = await fetch(
          "/api/roteiros/buscar/?q=" + encodeURIComponent(normalizedQuery)
        );
        const data = await response.json();
        const roteiros = Array.isArray(data)
          ? data
          : Array.isArray(data?.roteiros)
            ? data.roteiros
            : [];
        this.renderResultados(roteiros);
      } catch (error) {
        console.error("Erro ao buscar roteiros:", error);
        this.searchResultsContainer.innerHTML =
          '<div class="error">Erro ao carregar roteiros.</div>';
      }
    }

    renderResultados(roteiros) {
      this.searchResultsContainer.innerHTML = "";
      if (!roteiros.length) {
        this.searchResultsContainer.innerHTML =
          '<div class="empty">Nenhum roteiro encontrado.</div>';
        return;
      }

      roteiros.forEach((roteiro) => {
        const item = document.createElement("div");
        item.className = "card";
        item.style.marginBottom = "12px";
        item.innerHTML = `
          <div class="page-header" style="margin-bottom: 0;">
            <div>
              <div class="sub-card-title">${roteiro.nome}</div>
              <p class="page-subtitle" style="margin-bottom: 0;">
                ${roteiro.origem} &rarr; ${roteiro.destino} (${roteiro.tipo})
              </p>
            </div>
            <div class="page-actions">
              <button
                type="button"
                class="btn-primary use-roteiro-btn"
                data-roteiro-id="${roteiro.id}"
                data-roteiro-nome="${roteiro.nome}"
              >
                Usar
              </button>
            </div>
          </div>
        `;
        this.searchResultsContainer.appendChild(item);
      });

      this.searchResultsContainer
        .querySelectorAll(".use-roteiro-btn")
        .forEach((button) => {
          button.addEventListener("click", (event) => {
            const currentButton = event.currentTarget;
            this.selecionarRoteiro(
              currentButton.dataset.roteiroId,
              currentButton.dataset.roteiroNome
            );
          });
        });

      if (window.lucide) {
        window.lucide.createIcons();
      }
    }

    async selecionarRoteiro(roteiroId, roteiroNome) {
      if (!this.oficioId) {
        window.alert("Nao foi possivel identificar o oficio atual.");
        return;
      }

      const csrfToken =
        document.querySelector("[name=csrfmiddlewaretoken]")?.value || "";

      try {
        const vincularResponse = await fetch(
          "/oficios/" + this.oficioId + "/roteiros/vincular/",
          {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
              "X-CSRFToken": csrfToken,
            },
            body: JSON.stringify({ roteiro: roteiroId, observacao: "" }),
          }
        );
        const vincularData = await vincularResponse.json();

        if (!vincularResponse.ok || !vincularData.success) {
          throw new Error(vincularData.message || "Erro ao vincular roteiro.");
        }

        const roteiroResponse = await fetch("/api/roteiros/" + roteiroId + "/json/");
        const roteiroData = await roteiroResponse.json();

        if (typeof this.selectRoteiroCallback === "function") {
          this.selectRoteiroCallback(roteiroData);
        }

        if (this.modalElement) {
          closeModalElement(this.modalElement);
        }

        if (window.showToast) {
          window.showToast(
            "Roteiro '" + roteiroNome + "' vinculado com sucesso.",
            "success"
          );
        } else {
          window.alert(
            "Roteiro '" + roteiroNome + "' vinculado com sucesso."
          );
        }
      } catch (error) {
        console.error("Erro ao selecionar roteiro:", error);
        window.alert(error.message || "Erro ao selecionar roteiro.");
      }
    }

    limpar() {
      if (!this.searchInput || !this.searchResultsContainer) {
        return;
      }
      this.searchInput.value = "";
      this.searchResultsContainer.innerHTML =
        '<div class="empty">Digite para buscar roteiros.</div>';
    }
  }

  class TrechoFormset {
    constructor(prefix) {
      this.prefix = prefix;
      this.formsetContainer = document.getElementById(prefix + "-formset-container");
      this.addTrechoBtn = document.getElementById("add-trecho-btn");
      this.emptyFormTemplate = document.getElementById("empty-trecho-form");
      this.totalFormsInput = document.querySelector(
        'input[name="' + prefix + '-TOTAL_FORMS"]'
      );
      this.initialized = false;
    }

    init() {
      if (this.initialized) {
        return;
      }
      if (
        !this.formsetContainer ||
        !this.addTrechoBtn ||
        !this.emptyFormTemplate ||
        !this.totalFormsInput
      ) {
        console.error("Elementos do formset nao encontrados. Verifique IDs.");
        return;
      }

      this.initialized = true;
      this.addTrechoBtn.addEventListener("click", () => this.addTrecho());
      this.formsetContainer.addEventListener("click", (event) => {
        const removeButton = event.target.closest(".remove-trecho-btn");
        if (!removeButton) {
          return;
        }
        const row = removeButton.closest("[data-trecho-row]") || removeButton.closest(".card");
        if (row) {
          this.removeTrecho(row);
        }
      });
      this.updateOrders();
    }

    addTrecho() {
      const currentForms = Number.parseInt(this.totalFormsInput.value || "0", 10);
      const newFormHtml = this.emptyFormTemplate.innerHTML.replace(
        /__prefix__/g,
        String(currentForms)
      );
      const wrapper = document.createElement("div");
      wrapper.innerHTML = newFormHtml.trim();
      const newElement = wrapper.firstElementChild;
      if (!newElement) {
        return;
      }

      this.formsetContainer.appendChild(newElement);
      this.totalFormsInput.value = String(currentForms + 1);
      this.updateOrders();
      if (window.lucide) {
        window.lucide.createIcons();
      }
    }

    removeTrecho(formRow) {
      if (!window.confirm("Tem certeza que deseja remover este trecho?")) {
        return;
      }

      const deleteInput = formRow.querySelector('input[id$="-DELETE"]');
      if (deleteInput) {
        deleteInput.checked = true;
        formRow.hidden = true;
      } else {
        formRow.remove();
      }
      this.updateOrders();
    }

    updateOrders() {
      let order = 1;
      this.formsetContainer.querySelectorAll("[data-trecho-row]").forEach((row) => {
        if (row.hidden) {
          return;
        }
        const orderInput = row.querySelector('input[id$="-ordem"]');
        const orderSpan = row.querySelector(".trecho-ordem");
        if (orderInput) {
          orderInput.value = String(order);
        }
        if (orderSpan) {
          orderSpan.textContent = String(order);
        }
        order += 1;
      });
    }
  }

  window.RoteiroSelector = RoteiroSelector;
  window.TrechoFormset = TrechoFormset;
})();
