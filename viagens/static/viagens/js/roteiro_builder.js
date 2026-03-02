function initRoteiroBuilder(
  containerId = "destinos-container",
  previewId = "preview-cards-lista",
  isInline = false
) {
  const destinosContainer = document.getElementById(containerId);
  const previewCardsLista = document.getElementById(previewId);
  const roteiroNomeInput = document.getElementById(
    isInline ? "inline-roteiro-nome" : "roteiro-nome"
  );
  const roteiroUfSedeSelect = document.getElementById(
    isInline ? "inline-roteiro-uf-sede" : "roteiro-uf-sede"
  );
  const roteiroCidadeSedeInput = document.getElementById(
    isInline ? "inline-roteiro-cidade-sede" : "roteiro-cidade-sede"
  );
  const btnAddDestino = document.getElementById(
    isInline ? "inline-btn-add-destino" : "btn-add-destino"
  );
  const btnSalvarRoteiro = document.getElementById(
    isInline ? "inline-btn-salvar-roteiro" : "btn-salvar-roteiro"
  );
  const previewCardsSection = document.getElementById(
    isInline ? "inline-preview-cards" : "preview-cards"
  );
  const listaRoteirosContainer = document.getElementById("lista-roteiros-container");
  const noRoteirosMessage = document.getElementById("no-roteiros-message");

  if (
    !destinosContainer ||
    !previewCardsLista ||
    !roteiroNomeInput ||
    !roteiroUfSedeSelect ||
    !roteiroCidadeSedeInput ||
    !btnAddDestino ||
    !btnSalvarRoteiro ||
    !previewCardsSection
  ) {
    return;
  }

  const apiSalvar = "/api/roteiros/salvar/";
  const apiListar = "/api/roteiros/listar/";
  let destinoCounter = 0;

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

  function collectDestinos() {
    return Array.from(destinosContainer.querySelectorAll(".destino-item"))
      .map((item) => {
        const ufInput = item.querySelector(".destino-uf");
        const cidadeInput = item.querySelector(".destino-cidade");
        const uf = (ufInput?.value || "PR").trim().toUpperCase();
        const cidade = (cidadeInput?.value || "").trim();
        if (!cidade) {
          return null;
        }
        return { uf, cidade };
      })
      .filter(Boolean);
  }

  function renderPreview() {
    const sedeUf = (roteiroUfSedeSelect.value || "PR").trim().toUpperCase();
    const sedeCidade = (roteiroCidadeSedeInput.value || "Curitiba").trim();
    const destinos = collectDestinos();

    previewCardsLista.innerHTML = "";
    if (!destinos.length) {
      previewCardsSection.style.display = "none";
      return;
    }

    previewCardsSection.style.display = "block";
    const sequencia = [{ cidade: sedeCidade, uf: sedeUf }, ...destinos, { cidade: sedeCidade, uf: sedeUf }];

    for (let i = 0; i < sequencia.length - 1; i += 1) {
      const origem = sequencia[i];
      const destino = sequencia[i + 1];
      const isRetorno = i === sequencia.length - 2;
      const card = document.createElement("div");
      card.className = "sub-card";
      card.style.margin = "0";
      card.innerHTML = `
        <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;">
          <span class="badge">${isRetorno ? "Retorno" : `Trecho ${i + 1} (Ida)`}</span>
          <span class="field-help" style="margin:0;">${origem.cidade}/${origem.uf} -> ${destino.cidade}/${destino.uf}</span>
        </div>
      `;
      previewCardsLista.appendChild(card);
    }
  }

  function updateSortable() {
    if (!window.Sortable || destinosContainer.dataset.sortableInitialized === "true") {
      return;
    }
    window.Sortable.create(destinosContainer, {
      animation: 150,
      handle: ".drag-handle",
      onEnd: renderPreview,
    });
    destinosContainer.dataset.sortableInitialized = "true";
  }

  function addDestino(uf = "PR", cidade = "") {
    destinoCounter += 1;
    const wrapper = document.createElement("div");
    wrapper.className = "destino-item";
    wrapper.dataset.index = String(destinoCounter);
    wrapper.innerHTML = `
      <div class="destino-header">
        <div class="destino-info">
          <span class="badge">Destino ${destinoCounter}</span>
          <div class="trecho-subtitle">UF + cidade</div>
        </div>
        <div class="destino-actions">
          <span class="drag-handle" title="Arraste para reordenar" aria-label="Arraste para reordenar">&#9776;</span>
          <button type="button" class="btn-danger destino-remove">Remover</button>
        </div>
      </div>
      <div class="grid">
        <label>
          UF
          <select class="input-field destino-uf">
            <option value="PR"${uf === "PR" ? " selected" : ""}>PR</option>
          </select>
        </label>
        <label>
          Cidade
          <input type="text" class="input-field destino-cidade" value="${cidade}" placeholder="Cidade destino" />
        </label>
      </div>
    `;

    const ufInput = wrapper.querySelector(".destino-uf");
    const cidadeInput = wrapper.querySelector(".destino-cidade");
    const removeBtn = wrapper.querySelector(".destino-remove");

    ufInput?.addEventListener("change", renderPreview);
    cidadeInput?.addEventListener("input", renderPreview);
    removeBtn?.addEventListener("click", () => {
      wrapper.remove();
      renderPreview();
    });

    destinosContainer.appendChild(wrapper);
    updateSortable();
    renderPreview();
    cidadeInput?.focus();
  }

  async function carregarRoteiros() {
    if (!listaRoteirosContainer || !noRoteirosMessage) {
      return;
    }

    try {
      const query = new URLSearchParams(window.location.search).get("q") || "";
      const listUrl = query ? `${apiListar}?q=${encodeURIComponent(query)}` : apiListar;
      const response = await fetch(listUrl, {
        headers: { "X-Requested-With": "XMLHttpRequest" },
      });
      const roteiros = await response.json();
      listaRoteirosContainer.innerHTML = "";

      if (!Array.isArray(roteiros) || !roteiros.length) {
        noRoteirosMessage.hidden = false;
        return;
      }

      noRoteirosMessage.hidden = true;
      roteiros.forEach((roteiro) => {
        const card = document.createElement("div");
        card.className = "card";
        card.innerHTML = `
          <div class="page-header" style="margin:0 0 12px;">
            <div>
              <h3 class="sub-card-title" style="margin-bottom:4px;">${roteiro.nome}</h3>
              <p class="field-help" style="margin:0;">${roteiro.cidade_sede}/${roteiro.uf_sede} -> ${roteiro.cidades_destino || "-"}</p>
            </div>
            <span class="badge">${roteiro.total_cidades} destino${roteiro.total_cidades === 1 ? "" : "s"}</span>
          </div>
          <div class="wizard-actions" style="margin-top:0;">
            <a href="/roteiros/${roteiro.id}/detalhe/" class="btn-clear">Ver</a>
            <a href="/roteiros/${roteiro.id}/editar/" class="btn-clear">Editar</a>
          </div>
        `;
        listaRoteirosContainer.appendChild(card);
      });
    } catch (error) {
      listaRoteirosContainer.innerHTML =
        '<div class="error">Erro ao carregar os roteiros salvos.</div>';
    }
  }

  async function salvarRoteiro() {
    const nome = roteiroNomeInput.value.trim();
    const ufSede = (roteiroUfSedeSelect.value || "PR").trim().toUpperCase();
    const cidadeSede = (roteiroCidadeSedeInput.value || "Curitiba").trim();
    const destinos = collectDestinos();

    if (!nome) {
      roteiroNomeInput.focus();
      window.alert("Por favor, preencha o nome do roteiro.");
      return;
    }
    if (!destinos.length) {
      window.alert("Por favor, adicione pelo menos um destino ao roteiro.");
      return;
    }

    const originalLabel = btnSalvarRoteiro.textContent;
    btnSalvarRoteiro.disabled = true;
    btnSalvarRoteiro.textContent = "Salvando...";

    try {
      const response = await fetch(apiSalvar, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "X-CSRFToken": getCookie("csrftoken"),
          "X-Requested-With": "XMLHttpRequest",
        },
        body: JSON.stringify({
          nome: nome,
          uf_sede: ufSede,
          cidade_sede: cidadeSede,
          destinos: destinos,
        }),
      });

      const result = await response.json();
      if (!response.ok) {
        throw new Error(result.error || "Erro ao salvar roteiro.");
      }

      if (isInline) {
        document.dispatchEvent(
          new CustomEvent("roteiro:selecionado", {
            detail: { roteiroId: result.id || result.roteiro_id || "" },
            bubbles: true,
          })
        );
      } else {
        await carregarRoteiros();
      }

      roteiroNomeInput.value = "";
      destinosContainer.innerHTML = "";
      previewCardsLista.innerHTML = "";
      previewCardsSection.style.display = "none";
      destinoCounter = 0;
      window.alert("Roteiro salvo com sucesso.");
    } catch (error) {
      window.alert(error.message || "Erro de rede ao salvar roteiro.");
    } finally {
      btnSalvarRoteiro.disabled = false;
      btnSalvarRoteiro.textContent = originalLabel;
    }
  }

  btnAddDestino.addEventListener("click", () => addDestino());
  btnSalvarRoteiro.addEventListener("click", salvarRoteiro);
  roteiroUfSedeSelect.addEventListener("change", renderPreview);
  roteiroCidadeSedeInput.addEventListener("input", renderPreview);

  updateSortable();
  if (!isInline) {
    carregarRoteiros();
  }
}

window.initRoteiroBuilder = initRoteiroBuilder;
