/**
 * roteiro_selector.js
 * Central de Viagens ASCOM — PCPR
 * Integra roteiros reutilizaveis na etapa 3 do oficio.
 */

(function () {
  "use strict";

  const searchInput = document.getElementById("roteiro-search-input");
  const searchResults = document.getElementById("roteiro-search-results");
  const btnBuscar = document.getElementById("btn-buscar-roteiro");
  const btnUsar = document.getElementById("btn-usar-roteiro");
  const btnLimpar = document.getElementById("btn-limpar-roteiro");
  const btnSalvar = document.getElementById("btn-salvar-roteiro-inline");
  const nomeNovoInput = document.getElementById("roteiro-novo-nome");
  const preview = document.getElementById("roteiro-selecionado-preview");
  const previewNome = document.getElementById("preview-nome-roteiro");
  const previewDestinos = document.getElementById("preview-destinos-roteiro");
  const saveFeedback = document.getElementById("roteiro-save-feedback");
  const saveMsg = document.getElementById("roteiro-save-msg");
  const roteiroSelecionadoInput = document.getElementById("roteiro-selecionado-id");
  const roteiroOrigemInput = document.getElementById("roteiro-origem-id");
  const sedeUfSelect = document.getElementById("sedeUf");
  const sedeCidadeSelect = document.getElementById("sedeCidade");
  const addDestinoBtn = document.getElementById("addDestinoBtn");
  const destinosList = document.getElementById("destinosList");
  const tipoDestinoInput = document.getElementById("tipoDestino");
  const trechosSection = document.getElementById("trechos-gerados-section");

  if (!searchInput || !searchResults || !destinosList) {
    return;
  }

  const urlBuscar = searchInput.dataset.urlBuscar || "";
  const urlDetalhe = searchInput.dataset.urlDetalhe || "";
  const urlCriar = searchInput.dataset.urlCriar || "";
  const csrfToken =
    document.querySelector("[name=csrfmiddlewaretoken]")?.value || "";

  let roteiroSelecionado = null;
  let debounceTimer = null;

  function hideResults() {
    searchResults.hidden = true;
  }

  function showResults() {
    searchResults.hidden = false;
  }

  function hidePreview() {
    if (preview) {
      preview.hidden = true;
    }
  }

  function showPreview() {
    if (preview) {
      preview.hidden = false;
    }
  }

  function getSelectedLabel(select) {
    if (!select) {
      return "";
    }
    const option = select.selectedOptions && select.selectedOptions[0];
    return option ? option.textContent.trim() : "";
  }

  function createOrSelectOption(select, value, label, extraDataset) {
    if (!select || value === undefined || value === null || value === "") {
      return;
    }
    let option = Array.from(select.options).find(
      (item) => String(item.value) === String(value)
    );
    if (!option) {
      option = document.createElement("option");
      option.value = String(value);
      option.textContent = label || String(value);
      if (extraDataset && typeof extraDataset === "object") {
        Object.keys(extraDataset).forEach((key) => {
          option.dataset[key] = String(extraDataset[key]);
        });
      }
      select.appendChild(option);
    } else if (label) {
      option.textContent = label;
    }
    select.value = String(value);
    select.dataset.selected = String(value);
    window.syncAutocompleteDisplay?.(select);
  }

  function getDestinoCards() {
    return Array.from(
      destinosList.querySelectorAll(".destino-card, .destino-item")
    );
  }

  function clearDestinoCard(card) {
    if (!card) {
      return;
    }
    const ufSelect = card.querySelector("[data-role='destino-estado']");
    const cidadeSelect = card.querySelector("[data-role='destino-cidade']");
    if (ufSelect) {
      ufSelect.value = "";
      window.syncAutocompleteDisplay?.(ufSelect);
    }
    if (cidadeSelect) {
      cidadeSelect.value = "";
      cidadeSelect.dataset.selected = "";
      Array.from(cidadeSelect.options)
        .filter((option) => option.value)
        .forEach((option) => option.remove());
      window.syncAutocompleteDisplay?.(cidadeSelect);
      cidadeSelect.dispatchEvent(new Event("change", { bubbles: true }));
    }
  }

  function ensureDestinoCount(targetCount) {
    let cards = getDestinoCards();
    const normalizedTarget = Math.max(targetCount, 1);

    while (cards.length < normalizedTarget && addDestinoBtn) {
      addDestinoBtn.click();
      cards = getDestinoCards();
    }

    while (cards.length > normalizedTarget && cards.length > 1) {
      const lastCard = cards[cards.length - 1];
      const removeBtn = lastCard.querySelector("[data-action='remove-destino']");
      if (removeBtn) {
        removeBtn.click();
      } else {
        lastCard.remove();
      }
      cards = getDestinoCards();
    }

    return cards;
  }

  function getCamposDestinos() {
    return getDestinoCards()
      .map((card) => {
        const ufSelect = card.querySelector("[data-role='destino-estado']");
        const cidadeSelect = card.querySelector("[data-role='destino-cidade']");
        if (!ufSelect || !cidadeSelect || !cidadeSelect.value) {
          return null;
        }
        return {
          uf_destino: ufSelect.value,
          cidade_destino_id: cidadeSelect.value,
          cidade_destino_nome: getSelectedLabel(cidadeSelect),
        };
      })
      .filter(Boolean);
  }

  function getCamposSede() {
    const sedeUfOption = sedeUfSelect?.selectedOptions?.[0];
    return {
      uf_sede_sigla: sedeUfSelect ? sedeUfSelect.value : "",
      uf_sede_id: sedeUfOption?.dataset?.estadoId || "",
      cidade_sede_id: sedeCidadeSelect ? sedeCidadeSelect.value : "",
      cidade_sede_nome: getSelectedLabel(sedeCidadeSelect),
    };
  }

  function fillSede(roteiro) {
    const ufSigla = roteiro.uf_sede_sigla || roteiro.uf_origem || "";
    if (sedeUfSelect && ufSigla) {
      const option = Array.from(sedeUfSelect.options).find(
        (item) => item.value === ufSigla
      );
      if (option) {
        sedeUfSelect.value = ufSigla;
        window.syncAutocompleteDisplay?.(sedeUfSelect);
      }
      sedeUfSelect.dispatchEvent(new Event("change", { bubbles: true }));
    }

    if (sedeCidadeSelect && roteiro.cidade_sede_id) {
      const label = roteiro.cidade_sede_nome
        ? (roteiro.uf_sede_sigla || ufSigla)
          ? `${roteiro.cidade_sede_nome}/${roteiro.uf_sede_sigla || ufSigla}`
          : roteiro.cidade_sede_nome
        : roteiro.cidade_origem || "";
      createOrSelectOption(
        sedeCidadeSelect,
        roteiro.cidade_sede_id,
        label
      );
      sedeCidadeSelect.dispatchEvent(new Event("change", { bubbles: true }));
    }
  }

  function fillDestinos(roteiro) {
    const trechos = Array.isArray(roteiro.trechos) ? roteiro.trechos : [];
    const cards = ensureDestinoCount(trechos.length);

    if (!trechos.length && cards[0]) {
      clearDestinoCard(cards[0]);
      return;
    }

    cards.forEach((card, index) => {
      const trecho = trechos[index];
      if (!trecho) {
        clearDestinoCard(card);
        return;
      }

      const ufSelect = card.querySelector("[data-role='destino-estado']");
      const cidadeSelect = card.querySelector("[data-role='destino-cidade']");
      if (ufSelect) {
        const ufValue = trecho.uf_destino || "";
        if (ufValue) {
          const option = Array.from(ufSelect.options).find(
            (item) => item.value === ufValue
          );
          if (option) {
            ufSelect.value = ufValue;
          } else {
            createOrSelectOption(ufSelect, ufValue, ufValue);
          }
          window.syncAutocompleteDisplay?.(ufSelect);
          ufSelect.dispatchEvent(new Event("change", { bubbles: true }));
        }
      }

      if (cidadeSelect && trecho.cidade_destino_id) {
        const cidadeLabel = trecho.cidade_destino_nome
          ? `${trecho.cidade_destino_nome}/${trecho.uf_destino || ""}`.replace(/\/$/, "")
          : trecho.cidade_destino || "";
        createOrSelectOption(
          cidadeSelect,
          trecho.cidade_destino_id,
          cidadeLabel
        );
        cidadeSelect.dispatchEvent(new Event("change", { bubbles: true }));
      }
    });
  }

  function buildCardsPayload(roteiro) {
    const trechos = Array.isArray(roteiro?.trechos) ? roteiro.trechos : [];
    if (!trechos.length) {
      return [];
    }

    const cards = trechos.map((trecho, index) => ({
      tipo: `Trecho ${index + 1} (Ida)`,
      origem_estado: trecho.uf_origem || roteiro.uf_sede_sigla || roteiro.uf_origem || "",
      origem_cidade: trecho.cidade_origem || roteiro.cidade_origem || "",
      destino_estado: trecho.uf_destino || "",
      destino_cidade:
        trecho.cidade_destino_nome || trecho.cidade_destino || "",
      card_index: index + 1,
      is_retorno: false,
    }));

    const ultimoTrecho = trechos[trechos.length - 1];
    cards.push({
      tipo: "Retorno",
      origem_estado: ultimoTrecho.uf_destino || "",
      origem_cidade:
        ultimoTrecho.cidade_destino_nome || ultimoTrecho.cidade_destino || "",
      destino_estado: roteiro.uf_sede_sigla || roteiro.uf_origem || "",
      destino_cidade: roteiro.cidade_sede_nome || roteiro.cidade_origem || "",
      card_index: cards.length + 1,
      is_retorno: true,
    });

    return cards;
  }

  function preencherCamposComRoteiro(roteiro) {
    fillSede(roteiro);
    fillDestinos(roteiro);

    roteiroSelecionado = roteiro;
    if (roteiroSelecionadoInput) {
      roteiroSelecionadoInput.value = roteiro.id || "";
    }
    if (roteiroOrigemInput) {
      roteiroOrigemInput.value = roteiro.id || "";
    }
    if (previewNome) {
      previewNome.textContent = roteiro.nome || "—";
    }
    if (previewDestinos) {
      previewDestinos.textContent =
        roteiro.destinos ||
        (Array.isArray(roteiro.trechos)
          ? roteiro.trechos
              .map((trecho) => trecho.cidade_destino_nome || trecho.cidade_destino || "")
              .filter(Boolean)
              .join(" -> ")
          : "—");
    }

    showPreview();
    hideResults();
    if (searchInput) {
      searchInput.value = "";
    }
    if (btnUsar) {
      btnUsar.hidden = true;
    }

    document.dispatchEvent(
      new CustomEvent("roteiro:selecionado", {
        detail: roteiro,
        bubbles: true,
      })
    );
    document.dispatchEvent(
      new CustomEvent("oficio:roteiro_preenchido", {
        detail: { roteiroId: roteiro.id || "", roteiro: roteiro },
        bubbles: true,
      })
    );
    document.dispatchEvent(
      new CustomEvent("roteiro:aplicado", {
        detail: {
          roteiro_id: roteiro.id || "",
          roteiro_nome: roteiro.nome || "",
          cards: buildCardsPayload(roteiro),
          roteiro: roteiro,
        },
        bubbles: true,
      })
    );
    if (trechosSection) {
      trechosSection.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }

  async function buscarRoteiros(q) {
    const query = String(q || "").trim();
    if (query.length < 2) {
      searchResults.innerHTML = "";
      hideResults();
      return;
    }

    try {
      const response = await fetch(
        `${urlBuscar}?q=${encodeURIComponent(query)}&limit=10`,
        {
          headers: { "X-Requested-With": "XMLHttpRequest" },
        }
      );
      const data = await response.json();
      const roteiros = Array.isArray(data?.roteiros) ? data.roteiros : [];

      searchResults.innerHTML = "";
      if (!roteiros.length) {
        searchResults.innerHTML =
          '<div style="padding:10px 12px;color:#6b7280;font-size:0.85rem;">Nenhum roteiro encontrado.</div>';
        showResults();
        return;
      }

      roteiros.forEach((roteiro) => {
        const item = document.createElement("button");
        item.type = "button";
        item.style.cssText =
          "display:block;width:100%;text-align:left;padding:10px 12px;border:none;border-bottom:1px solid #eef2f7;background:#fff;cursor:pointer;";
        item.innerHTML = `
          <div style="font-weight:700;color:#111827;">${roteiro.nome}</div>
          <div style="font-size:0.78rem;color:#6b7280;">${roteiro.destinos || "—"}</div>
        `;
        item.addEventListener("mouseenter", () => {
          item.style.background = "#f0fdf4";
        });
        item.addEventListener("mouseleave", () => {
          item.style.background = "#fff";
        });
        item.addEventListener("click", () => {
          selecionarRoteiro(roteiro);
        });
        searchResults.appendChild(item);
      });

      showResults();
    } catch (error) {
      console.error("[RoteiroSelector] Erro na busca:", error);
      searchResults.innerHTML =
        '<div style="padding:10px 12px;color:#991b1b;font-size:0.85rem;">Erro ao buscar roteiros.</div>';
      showResults();
    }
  }

  async function selecionarRoteiro(roteiroBasico) {
    try {
      const detailUrl = urlDetalhe.replace("/0/json/", `/${roteiroBasico.id}/json/`);
      const response = await fetch(detailUrl, {
        headers: { "X-Requested-With": "XMLHttpRequest" },
      });
      const roteiro = await response.json();
      if (!response.ok) {
        throw new Error(roteiro.erro || "Erro ao carregar detalhes do roteiro.");
      }
      preencherCamposComRoteiro(roteiro);
    } catch (error) {
      console.error("[RoteiroSelector] Erro ao carregar roteiro:", error);
      window.alert(error.message || "Erro ao carregar detalhes do roteiro.");
    }
  }

  async function salvarRoteiroInline() {
    const nome = String(nomeNovoInput?.value || "").trim();
    if (!nome) {
      nomeNovoInput?.focus();
      return;
    }

    const sede = getCamposSede();
    const destinos = getCamposDestinos();
    if (!sede.uf_sede_id || !sede.cidade_sede_id) {
      window.alert("Preencha a UF e Cidade sede antes de salvar o roteiro.");
      return;
    }
    if (!destinos.length) {
      window.alert("Adicione ao menos um destino antes de salvar o roteiro.");
      return;
    }

    const originalLabel = btnSalvar ? btnSalvar.textContent : "";
    const payload = {
      nome: nome,
      uf_sede_id: sede.uf_sede_id,
      cidade_sede_id: sede.cidade_sede_id,
      tipo_destino: tipoDestinoInput?.value || "",
      trechos: destinos.map((destino, index) => ({
        ordem: index + 1,
        uf_destino: destino.uf_destino,
        cidade_destino_id: destino.cidade_destino_id,
        modal: "veiculo_proprio",
      })),
    };

    try {
      if (btnSalvar) {
        btnSalvar.disabled = true;
        btnSalvar.textContent = "Salvando...";
      }

      const response = await fetch(urlCriar, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "X-CSRFToken": csrfToken,
          "X-Requested-With": "XMLHttpRequest",
        },
        body: JSON.stringify(payload),
      });
      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.erro || "Erro ao salvar o roteiro.");
      }

      if (saveMsg) {
        saveMsg.textContent = data.mensagem || "Roteiro salvo na biblioteca!";
      }
      if (saveFeedback) {
        saveFeedback.hidden = false;
      }
      if (nomeNovoInput) {
        nomeNovoInput.value = "";
      }
      if (roteiroSelecionadoInput) {
        roteiroSelecionadoInput.value = data.roteiro_id || "";
      }
      if (roteiroOrigemInput) {
        roteiroOrigemInput.value = data.roteiro_id || "";
      }
      if (previewNome) {
        previewNome.textContent = data.nome || nome;
      }
      if (previewDestinos) {
        previewDestinos.textContent =
          destinos
            .map((destino) => destino.cidade_destino_nome || "")
            .filter(Boolean)
            .join(" -> ") || "â€”";
      }
      showPreview();
      window.setTimeout(() => {
        if (saveFeedback) {
          saveFeedback.hidden = true;
        }
      }, 4000);
    } catch (error) {
      console.error("[RoteiroSelector] Erro ao salvar:", error);
      window.alert(error.message || "Erro inesperado ao salvar o roteiro.");
    } finally {
      if (btnSalvar) {
        btnSalvar.disabled = false;
        btnSalvar.textContent = originalLabel || "Salvar na Biblioteca";
      }
    }
  }

  searchInput.addEventListener("input", (event) => {
    window.clearTimeout(debounceTimer);
    debounceTimer = window.setTimeout(() => {
      buscarRoteiros(event.target.value.trim());
    }, 300);
  });

  btnBuscar?.addEventListener("click", () => {
    buscarRoteiros(searchInput.value.trim());
  });

  btnLimpar?.addEventListener("click", () => {
    roteiroSelecionado = null;
    if (roteiroSelecionadoInput) {
      roteiroSelecionadoInput.value = "";
    }
    if (roteiroOrigemInput) {
      roteiroOrigemInput.value = "";
    }
    hidePreview();
    hideResults();
    searchInput.value = "";
    document.dispatchEvent(
      new CustomEvent("roteiro:limpo", {
        bubbles: true,
      })
    );
    document.dispatchEvent(
      new CustomEvent("roteiro:limpar", {
        bubbles: true,
      })
    );
  });

  btnSalvar?.addEventListener("click", salvarRoteiroInline);

  document.addEventListener("click", (event) => {
    if (!event.target.closest("#painel-roteiro")) {
      hideResults();
    }
  });

  if (window.lucide) {
    window.lucide.createIcons();
  }
})();
