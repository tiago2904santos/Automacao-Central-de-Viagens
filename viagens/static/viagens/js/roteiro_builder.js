(function () {
  "use strict";

  function getCookie(name) {
    let cookieValue = null;
    if (document.cookie && document.cookie !== "") {
      const cookies = document.cookie.split(";");
      for (let i = 0; i < cookies.length; i += 1) {
        const cookie = cookies[i].trim();
        if (cookie.substring(0, name.length + 1) === `${name}=`) {
          cookieValue = decodeURIComponent(cookie.substring(name.length + 1));
          break;
        }
      }
    }
    return cookieValue;
  }

  function mostrarAlerta(type, message) {
    const container = document.getElementById("alert-container");
    if (!container) {
      window.alert(message);
      return;
    }

    const className =
      type === "success" ? "empty" : type === "warning" ? "empty" : "error";
    container.innerHTML = `<div class="${className}">${message}</div>`;
    window.setTimeout(() => {
      container.innerHTML = "";
    }, 5000);
  }

  function gerarOptionsUF(defaultUF) {
    const ufs = [
      ["AC", "Acre"],
      ["AL", "Alagoas"],
      ["AP", "Amapa"],
      ["AM", "Amazonas"],
      ["BA", "Bahia"],
      ["CE", "Ceara"],
      ["DF", "Distrito Federal"],
      ["ES", "Espirito Santo"],
      ["GO", "Goias"],
      ["MA", "Maranhao"],
      ["MT", "Mato Grosso"],
      ["MS", "Mato Grosso do Sul"],
      ["MG", "Minas Gerais"],
      ["PA", "Para"],
      ["PB", "Paraiba"],
      ["PE", "Pernambuco"],
      ["PI", "Piaui"],
      ["PR", "Parana"],
      ["RJ", "Rio de Janeiro"],
      ["RN", "Rio Grande do Norte"],
      ["RO", "Rondonia"],
      ["RR", "Roraima"],
      ["RS", "Rio Grande do Sul"],
      ["SC", "Santa Catarina"],
      ["SE", "Sergipe"],
      ["SP", "Sao Paulo"],
      ["TO", "Tocantins"],
    ];

    return ufs
      .map(
        ([sigla, nome]) =>
          `<option value="${sigla}"${sigla === defaultUF ? " selected" : ""}>${sigla} - ${nome}</option>`
      )
      .join("");
  }

  function pickElement(ids) {
    for (const id of ids) {
      const element = document.getElementById(id);
      if (element) {
        return element;
      }
    }
    return null;
  }

  async function safeJsonFromResponse(response) {
    const contentType = response.headers.get("content-type") || "";
    if (!contentType.includes("application/json")) {
      const rawText = await response.text();
      if (rawText && rawText.trim().startsWith("<")) {
        throw new Error(
          "Servidor retornou HTML em vez de JSON. Verifique autenticacao, URL e logs."
        );
      }
      throw new Error("Resposta inesperada do servidor.");
    }
    return response.json();
  }

  async function fetchCidadesPorUf(uf) {
    const normalizedUf = String(uf || "PR").trim().toUpperCase() || "PR";
    const response = await fetch(`/api/cidades/?uf=${encodeURIComponent(normalizedUf)}`, {
      headers: {
        "X-Requested-With": "XMLHttpRequest",
      },
    });
    const data = await safeJsonFromResponse(response);
    return Array.isArray(data?.cidades) ? data.cidades : [];
  }

  function initRoteiroBuilder(arg1, arg2, arg3) {
    const isInline = typeof arg1 === "boolean" ? arg1 : Boolean(arg3);
    const nomeInput = pickElement(
      isInline ? ["inline-roteiro-nome"] : ["nomeRoteiro", "id_nome_roteiro"]
    );
    const ufSedeEl = pickElement(
      isInline ? ["inline-roteiro-uf-sede"] : ["ufSede", "id_uf_sede_roteiro"]
    );
    const cidadeSedeEl = pickElement(
      isInline
        ? ["inline-roteiro-cidade-sede"]
        : ["cidadeSede", "id_cidade_sede_roteiro"]
    );
    const listaDestinos = pickElement(
      isInline ? ["inline-destinos-container"] : ["listaDestinos", "lista-destinos"]
    );
    const previewContainer = pickElement(
      isInline
        ? ["inline-preview-cards-lista"]
        : ["previewTrechos", "preview-trechos-container"]
    );
    const previewSection = pickElement(
      isInline ? ["inline-preview-cards"] : ["cardPreviewTrechos", "preview-cards"]
    );
    const addBtn = pickElement(
      isInline ? ["inline-btn-add-destino"] : ["btnAdicionarDestino", "btn-add-destino"]
    );
    const saveBtn = pickElement(
      isInline ? ["inline-btn-salvar-roteiro"] : ["btnSalvar", "btn-salvar-roteiro"]
    );
    const erroNome = pickElement(["erroNome"]);
    const templateDestino = document.getElementById("templateDestino");
    const templatePreviewTrecho = document.getElementById("templatePreviewTrecho");

    if (
      !nomeInput ||
      !ufSedeEl ||
      !cidadeSedeEl ||
      !listaDestinos ||
      !previewContainer ||
      !addBtn ||
      !saveBtn
    ) {
      return null;
    }

    let sortableInstance = null;

    function setErroNome(message) {
      if (!erroNome) {
        return;
      }
      if (message) {
        erroNome.hidden = false;
        erroNome.textContent = message;
      } else {
        erroNome.hidden = true;
        erroNome.textContent = "";
      }
    }

    function populateUfSelect(selectEl, selectedUf) {
      if (!selectEl) {
        return;
      }
      selectEl.innerHTML = gerarOptionsUF(selectedUf || "PR");
    }

    function populateCidadeSelect(selectEl, cidades, selectedValue) {
      if (!selectEl) {
        return;
      }
      const hasBlank = !selectedValue;
      selectEl.innerHTML = "";

      const blankOption = document.createElement("option");
      blankOption.value = "";
      blankOption.textContent = "Selecione";
      blankOption.selected = hasBlank;
      selectEl.appendChild(blankOption);

      cidades.forEach((cidade) => {
        const option = document.createElement("option");
        option.value = String(cidade.id);
        option.textContent = cidade.nome;
        if (selectedValue && String(selectedValue) === String(cidade.id)) {
          option.selected = true;
        }
        selectEl.appendChild(option);
      });
    }

    async function reloadCidadeOptions(selectEl, uf, selectedValue) {
      if (!selectEl || selectEl.tagName !== "SELECT") {
        return;
      }
      const cidades = await fetchCidadesPorUf(uf);
      populateCidadeSelect(selectEl, cidades, selectedValue);
    }

    function readSelectedOptionLabel(selectEl) {
      if (!selectEl || selectEl.tagName !== "SELECT") {
        return String(selectEl?.value || "").trim();
      }
      const option = selectEl.selectedOptions && selectEl.selectedOptions[0];
      return option ? option.textContent.trim() : "";
    }

    function togglePreview(show) {
      if (!previewSection) {
        return;
      }
      if ("hidden" in previewSection) {
        previewSection.hidden = !show;
      } else {
        previewSection.style.display = show ? "block" : "none";
      }
    }

    function getDestinoCards() {
      return Array.from(listaDestinos.querySelectorAll(".destino-card, .destino-item"));
    }

    function renumerarDestinos() {
      getDestinoCards().forEach((card, index) => {
        card.dataset.ordem = String(index + 1);
        const label = card.querySelector(".destino-label, .destino-badge, .badge");
        if (label) {
          label.textContent = `Destino ${index + 1}`;
        }
      });
    }

    function setDestinoError(card, show) {
      const error = card.querySelector(".destino-error, [data-destino-error]");
      if (!error) {
        return;
      }
      error.hidden = !show;
    }

    function coletarDestinos(validar) {
      const destinos = [];
      let hasError = false;

      getDestinoCards().forEach((card) => {
        const ufInput = card.querySelector(".uf-destino, [data-role='destino-estado']");
        const cidadeInput = card.querySelector(
          ".cidade-destino, [data-role='destino-cidade']"
        );

        const uf = String(ufInput?.value || "PR").trim().toUpperCase() || "PR";
        const cidadeId = cidadeInput && cidadeInput.tagName === "SELECT" ? cidadeInput.value : "";
        const cidade =
          cidadeInput && cidadeInput.tagName === "SELECT"
            ? readSelectedOptionLabel(cidadeInput)
            : String(cidadeInput?.value || "").trim();

        const isFilled = Boolean(cidade) && (cidadeInput?.tagName !== "SELECT" || Boolean(cidadeId));

        if (validar && !isFilled) {
          hasError = true;
          card.classList.add("is-invalid");
          cidadeInput?.classList.add("is-invalid");
          setDestinoError(card, true);
        } else {
          card.classList.remove("is-invalid");
          cidadeInput?.classList.remove("is-invalid");
          setDestinoError(card, false);
        }

        if (isFilled) {
          destinos.push({
            uf: uf,
            cidade: cidade,
            cidade_id: cidadeId || "",
          });
        }
      });

      return hasError ? null : destinos;
    }

    function renderPreview(trechos) {
      previewContainer.innerHTML = "";
      if (!trechos.length) {
        togglePreview(false);
        return;
      }

      togglePreview(true);
      trechos.forEach((trecho) => {
        let element = null;
        if (templatePreviewTrecho?.content?.firstElementChild) {
          element = templatePreviewTrecho.content.cloneNode(true).firstElementChild;
          const tipo = element.querySelector(".preview-tipo");
          const rota = element.querySelector(".preview-rota");
          if (tipo) {
            tipo.textContent = trecho.tipo;
            if (trecho.isRetorno) {
              tipo.style.background = "#6b7280";
            }
          }
          if (rota) {
            rota.textContent = trecho.rota;
          }
        } else {
          element = document.createElement("div");
          element.className = "sub-card";
          element.style.margin = "0";
          element.innerHTML = `
            <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;">
              <span class="badge">${trecho.tipo}</span>
              <span class="field-help" style="margin:0;">${trecho.rota}</span>
            </div>
          `;
        }
        previewContainer.appendChild(element);
      });
    }

    function atualizarPreviewTrechos() {
      const cidadeSede = readSelectedOptionLabel(cidadeSedeEl);
      const ufSede = String(ufSedeEl.value || "PR").trim().toUpperCase() || "PR";
      const destinos = coletarDestinos(false);

      if (!cidadeSede || !destinos.length) {
        renderPreview([]);
        return;
      }

      const trechos = [];
      let origemAtual = { cidade: cidadeSede, uf: ufSede };

      destinos.forEach((destino, index) => {
        trechos.push({
          tipo: `Trecho ${index + 1} (Ida)`,
          rota: `${origemAtual.cidade}/${origemAtual.uf} -> ${destino.cidade}/${destino.uf}`,
          isRetorno: false,
        });
        origemAtual = { cidade: destino.cidade, uf: destino.uf };
      });

      trechos.push({
        tipo: "Retorno",
        rota: `${origemAtual.cidade}/${origemAtual.uf} -> ${cidadeSede}/${ufSede}`,
        isRetorno: true,
      });

      renderPreview(trechos);
    }

    function bindDestinoCard(card) {
      const ufSelect = card.querySelector(".uf-destino, [data-role='destino-estado']");
      const cidadeSelect = card.querySelector(
        ".cidade-destino, [data-role='destino-cidade']"
      );
      const removeBtn = card.querySelector(
        ".btn-remover-destino, .btn-remove-destino, [data-action='remove-destino']"
      );

      ufSelect?.addEventListener("change", async function () {
        if (cidadeSelect && cidadeSelect.tagName === "SELECT") {
          await reloadCidadeOptions(cidadeSelect, this.value);
        }
        atualizarPreviewTrechos();
      });

      cidadeSelect?.addEventListener("change", function () {
        this.classList.remove("is-invalid");
        card.classList.remove("is-invalid");
        setDestinoError(card, false);
        atualizarPreviewTrechos();
      });

      removeBtn?.addEventListener("click", function () {
        if (getDestinoCards().length <= 1) {
          mostrarAlerta("warning", "O roteiro precisa ter pelo menos um destino.");
          return;
        }
        card.remove();
        renumerarDestinos();
        atualizarPreviewTrechos();
      });
    }

    async function adicionarDestino(uf, cidadeId) {
      let card = null;
      if (templateDestino?.content?.firstElementChild) {
        card = templateDestino.content.cloneNode(true).firstElementChild;
      } else {
        const wrapper = document.createElement("div");
        wrapper.innerHTML = `
          <div class="destino-item destino-card">
            <div class="destino-header">
              <div class="destino-info">
                <span class="badge destino-label">Destino 1</span>
                <div class="trecho-subtitle">UF + cidade</div>
              </div>
              <div class="destino-actions">
                <span class="drag-handle" title="Arrastar para reordenar">&#9776;</span>
                <button type="button" class="btn-danger btn-remover-destino">Remover</button>
              </div>
            </div>
            <div class="grid">
              <label>
                UF
                <select class="input-field uf-destino"></select>
              </label>
              <label>
                Cidade
                <select class="input-field cidade-destino"></select>
              </label>
            </div>
            <span class="field-errors destino-error" hidden>Informe a cidade para continuar.</span>
          </div>
        `;
        card = wrapper.firstElementChild;
      }

      const ufSelect = card.querySelector(".uf-destino, [data-role='destino-estado']");
      const cidadeSelect = card.querySelector(
        ".cidade-destino, [data-role='destino-cidade']"
      );

      if (ufSelect) {
        populateUfSelect(ufSelect, uf || "PR");
      }
      if (cidadeSelect && cidadeSelect.tagName === "SELECT") {
        await reloadCidadeOptions(cidadeSelect, uf || "PR", cidadeId || "");
      }

      bindDestinoCard(card);
      listaDestinos.appendChild(card);
      renumerarDestinos();
      atualizarPreviewTrechos();
      cidadeSelect?.focus();
      return card;
    }

    async function atualizarCidadesDaSede() {
      if (cidadeSedeEl.tagName !== "SELECT") {
        atualizarPreviewTrechos();
        return;
      }
      const selectedValue = cidadeSedeEl.value;
      await reloadCidadeOptions(cidadeSedeEl, ufSedeEl.value, selectedValue);
      atualizarPreviewTrechos();
    }

    async function salvarRoteiro() {
      const nome = String(nomeInput.value || "").trim();
      const destinos = coletarDestinos(true);
      const cidadeSede = readSelectedOptionLabel(cidadeSedeEl);

      setErroNome("");

      if (nome.length < 3) {
        setErroNome("Informe um nome de roteiro com pelo menos 3 caracteres.");
        nomeInput.focus();
        return;
      }

      if (!cidadeSedeEl.value || !cidadeSede) {
        mostrarAlerta("danger", "Selecione a cidade da sede.");
        cidadeSedeEl.focus();
        return;
      }

      if (!destinos) {
        mostrarAlerta("danger", "Preencha a cidade de todos os destinos.");
        return;
      }

      if (!destinos.length) {
        mostrarAlerta("warning", "Adicione pelo menos um destino para o roteiro.");
        return;
      }

      const originalLabel = saveBtn.innerHTML;
      saveBtn.disabled = true;
      saveBtn.innerHTML =
        '<span class="spinner-border spinner-border-sm" aria-hidden="true"></span> Salvando...';

      try {
        const response = await fetch("/roteiros/salvar/", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "X-CSRFToken": getCookie("csrftoken"),
            "X-Requested-With": "XMLHttpRequest",
          },
          body: JSON.stringify({
            nome: nome,
            uf_sede: String(ufSedeEl.value || "PR").trim().toUpperCase() || "PR",
            cidade_sede: cidadeSede,
            destinos: destinos.map((destino) => ({
              uf: destino.uf,
              cidade: destino.cidade,
            })),
          }),
        });

        const data = await safeJsonFromResponse(response);
        if (!response.ok || (!data.ok && !data.sucesso)) {
          throw new Error(data.erro || data.error || "Erro ao salvar roteiro.");
        }

        mostrarAlerta("success", data.mensagem || "Roteiro salvo com sucesso.");
        saveBtn.innerHTML = "Salvo! ✓";

        window.setTimeout(() => {
          if (isInline) {
            document.dispatchEvent(
              new CustomEvent("roteiro:selecionado", {
                detail: { roteiroId: data.roteiro_id || data.id || "" },
                bubbles: true,
              })
            );
            saveBtn.disabled = false;
            saveBtn.innerHTML = originalLabel;
          } else {
            window.location.href = "/roteiros/";
          }
        }, 900);
      } catch (error) {
        mostrarAlerta("danger", error.message || "Erro de conexao ao salvar o roteiro.");
        saveBtn.disabled = false;
        saveBtn.innerHTML = originalLabel;
      }
    }

    addBtn.addEventListener("click", function () {
      adicionarDestino("PR", "");
    });
    saveBtn.addEventListener("click", salvarRoteiro);
    ufSedeEl.addEventListener("change", atualizarCidadesDaSede);
    cidadeSedeEl.addEventListener(
      cidadeSedeEl.tagName === "SELECT" ? "change" : "input",
      atualizarPreviewTrechos
    );

    if (window.Sortable) {
      sortableInstance = window.Sortable.create(listaDestinos, {
        animation: 150,
        handle: ".drag-handle",
        onEnd: function () {
          renumerarDestinos();
          atualizarPreviewTrechos();
        },
      });
    }

    if (getDestinoCards().length) {
      getDestinoCards().forEach((card) => {
        const ufSelect = card.querySelector(".uf-destino, [data-role='destino-estado']");
        if (ufSelect) {
          populateUfSelect(ufSelect, ufSelect.value || "PR");
        }
        bindDestinoCard(card);
      });
      renumerarDestinos();
      atualizarPreviewTrechos();
    } else {
      adicionarDestino("PR", "");
    }

    const api = {
      addDestino: adicionarDestino,
      coletarDestinos: coletarDestinos,
      atualizarPreviewTrechos: atualizarPreviewTrechos,
      sortable: sortableInstance,
    };
    window._roteiroBuilder = api;
    return api;
  }

  window.getCookie = getCookie;
  window.mostrarAlerta = mostrarAlerta;
  window.gerarOptionsUF = gerarOptionsUF;
  window.initRoteiroBuilder = initRoteiroBuilder;
  window.adicionarDestino = function () {
    return window._roteiroBuilder?.addDestino("PR", "");
  };
  window.coletarDestinos = function (validar) {
    return window._roteiroBuilder?.coletarDestinos(validar);
  };
  window.atualizarPreviewTrechos = function () {
    return window._roteiroBuilder?.atualizarPreviewTrechos();
  };
  window.removerDestino = function (button) {
    button
      ?.closest(".destino-card")
      ?.querySelector(".btn-remover-destino, .btn-remove-destino, [data-action='remove-destino']")
      ?.click();
  };
})();
