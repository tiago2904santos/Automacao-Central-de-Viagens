(function () {
  "use strict";

  const CIDADES_URL = "/api/cidades/";
  const LAST_SAIDA_HORA_KEY = "roteiro_last_saida_hora";

  const state = {
    cidadesCache: {},
    destinos: [],
    initialTrechos: Array.isArray(window.ROTEIRO_INICIAL) ? window.ROTEIRO_INICIAL : [],
    initialRetorno:
      window.RETORNO_INICIAL && typeof window.RETORNO_INICIAL === "object"
        ? window.RETORNO_INICIAL
        : {},
    initialTempoViagem: String(window.TEMPO_VIAGEM_INICIAL || "").trim(),
    sortable: null,
  };

  function $(id) {
    return document.getElementById(id);
  }

  function getForm() {
    return $("roteiroForm");
  }

  function getCsrfToken() {
    const fromForm = getForm()?.querySelector('input[name="csrfmiddlewaretoken"]')?.value;
    if (fromForm) {
      return fromForm;
    }
    const match = document.cookie.match(/csrftoken=([^;]+)/);
    return match ? match[1] : "";
  }

  function setMessage(element, message) {
    if (!element) {
      return;
    }
    element.hidden = !message;
    element.textContent = message || "";
  }

  function getLastSaidaHora() {
    try {
      return window.localStorage.getItem(LAST_SAIDA_HORA_KEY) || "";
    } catch (error) {
      return "";
    }
  }

  function setLastSaidaHora(value) {
    if (!value) {
      return;
    }
    try {
      window.localStorage.setItem(LAST_SAIDA_HORA_KEY, value);
    } catch (error) {
      // Ignora falhas de storage no navegador.
    }
  }

  async function fetchJson(url) {
    const response = await fetch(url, {
      headers: { "X-Requested-With": "XMLHttpRequest" },
    });
    const contentType = response.headers.get("content-type") || "";
    if (!contentType.includes("application/json")) {
      const raw = await response.text();
      throw new Error(raw && raw.trim().startsWith("<") ? "O servidor retornou HTML em vez de JSON." : "Resposta invalida do servidor.");
    }
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.erro || data.error || "Falha ao carregar dados.");
    }
    return data;
  }

  async function loadCities(uf) {
    const normalizedUf = String(uf || "PR").trim().toUpperCase() || "PR";
    if (state.cidadesCache[normalizedUf]) {
      return state.cidadesCache[normalizedUf];
    }
    const data = await fetchJson(`${CIDADES_URL}?uf=${encodeURIComponent(normalizedUf)}`);
    const cidades = Array.isArray(data.cidades)
      ? data.cidades.map((item) => String(item.nome || "").trim()).filter(Boolean)
      : [];
    state.cidadesCache[normalizedUf] = cidades;
    return cidades;
  }

  function populateCitySelect(select, cities, selectedValue) {
    if (!select) {
      return;
    }
    select.innerHTML = '<option value="">Selecione</option>';
    cities.forEach((cityName) => {
      const option = document.createElement("option");
      option.value = cityName;
      option.textContent = cityName;
      if (cityName === selectedValue) {
        option.selected = true;
      }
      select.appendChild(option);
    });
    if (
      selectedValue &&
      !Array.from(select.options).some((option) => option.value === selectedValue)
    ) {
      const fallback = document.createElement("option");
      fallback.value = selectedValue;
      fallback.textContent = selectedValue;
      fallback.selected = true;
      select.appendChild(fallback);
    }
  }

  async function hydrateCitySelect(select, uf, selectedValue) {
    const cities = await loadCities(uf);
    populateCitySelect(select, cities, selectedValue);
  }

  function getSedeUf() {
    return String($("sedeUf")?.value || "PR").trim().toUpperCase() || "PR";
  }

  function getSedeCidade() {
    return String($("sedeCidade")?.value || "").trim();
  }

  function getTempoViagem() {
    return String($("tempoViagem")?.value || "").trim();
  }

  function getSedeLabel() {
    const cidade = getSedeCidade();
    const uf = getSedeUf();
    return cidade ? `${cidade}/${uf}` : "-";
  }

  function buildRouteKey(origemEstado, origemCidade, destinoEstado, destinoCidade) {
    return [
      String(origemEstado || "").trim().toUpperCase(),
      String(origemCidade || "").trim(),
      String(destinoEstado || "").trim().toUpperCase(),
      String(destinoCidade || "").trim(),
    ].join("|");
  }

  function collectTrechoDrafts() {
    const drafts = {};
    document.querySelectorAll('[data-role="trecho-card"]').forEach((card) => {
      const key = buildRouteKey(
        card.dataset.origemEstado,
        card.dataset.origemCidade,
        card.dataset.destinoEstado,
        card.dataset.destinoCidade
      );
      if (!drafts[key]) {
        drafts[key] = [];
      }
      drafts[key].push({
        saida_data: card.querySelector('[data-role="saida-data"]')?.value || "",
        saida_hora: card.querySelector('[data-role="saida-hora"]')?.value || "",
        chegada_data: card.querySelector('[data-role="chegada-data"]')?.value || "",
        chegada_hora: card.querySelector('[data-role="chegada-hora"]')?.value || "",
      });
    });
    return drafts;
  }

  function captureRetorno() {
    return {
      saida_data: $("retornoSaidaData")?.value || "",
      saida_hora: $("retornoSaidaHora")?.value || "",
      chegada_data: $("retornoChegadaData")?.value || "",
      chegada_hora: $("retornoChegadaHora")?.value || "",
    };
  }

  function applyRetorno(retorno) {
    const origemPreview = document.querySelector('[data-role="retorno-origem-local"]');
    const destinoPreview = document.querySelector('[data-role="retorno-destino-local"]');
    const validDestinos = state.destinos.filter((item) => item.cidade);
    const lastDestino = validDestinos.length ? validDestinos[validDestinos.length - 1] : null;
    const retornoCard = $("retornoCard");

    if (!retornoCard) {
      return;
    }

    retornoCard.hidden = !lastDestino || !getSedeCidade();
    if (!lastDestino || !getSedeCidade()) {
      if (origemPreview) {
        origemPreview.textContent = "-";
      }
      if (destinoPreview) {
        destinoPreview.textContent = "-";
      }
      if ($("retornoSaidaData")) {
        $("retornoSaidaData").value = "";
      }
      if ($("retornoSaidaHora")) {
        $("retornoSaidaHora").value = "";
      }
      if ($("retornoChegadaData")) {
        $("retornoChegadaData").value = "";
      }
      if ($("retornoChegadaHora")) {
        $("retornoChegadaHora").value = "";
      }
      return;
    }

    if (origemPreview) {
      origemPreview.textContent = `${lastDestino.cidade}/${lastDestino.uf}`;
    }
    if (destinoPreview) {
      destinoPreview.textContent = getSedeLabel();
    }

    $("retornoSaidaData").value = retorno.saida_data || "";
    $("retornoSaidaHora").value = retorno.saida_hora || getLastSaidaHora() || "";
    $("retornoChegadaData").value = retorno.chegada_data || "";
    $("retornoChegadaHora").value = retorno.chegada_hora || "";
    if ((!retorno.chegada_data || !retorno.chegada_hora) && getTempoViagem()) {
      applyArrivalToRetorno();
    }
  }

  function showTrechosEmpty(message) {
    const list = $("trechosList");
    if (!list) {
      return;
    }
    list.innerHTML = "";
    const paragraph = document.createElement("p");
    paragraph.className = "field-help";
    paragraph.id = "trechosEmptyMsg";
    paragraph.style.margin = "0";
    paragraph.textContent = message;
    list.appendChild(paragraph);
  }

  function buildLegs() {
    const sedeCidade = getSedeCidade();
    const sedeUf = getSedeUf();
    const validDestinos = state.destinos.filter((item) => item.cidade);
    if (!sedeCidade || !validDestinos.length) {
      return [];
    }

    const legs = [];
    let origem = { uf: sedeUf, cidade: sedeCidade };
    validDestinos.forEach((destino) => {
      legs.push({
        origem: { uf: origem.uf, cidade: origem.cidade },
        destino: { uf: destino.uf, cidade: destino.cidade },
      });
      origem = { uf: destino.uf, cidade: destino.cidade };
    });
    return legs;
  }

  function calcularChegada(saidaDataVal, saidaHoraVal, tempoViagemVal) {
    const saidaHora = String(saidaHoraVal || "").trim();
    const tempoViagem = String(tempoViagemVal || "").trim();
    if (!saidaHora || !tempoViagem) {
      return null;
    }

    const [sh, sm] = saidaHora.split(":").map(Number);
    const [th, tm] = tempoViagem.split(":").map(Number);
    if ([sh, sm, th, tm].some((value) => Number.isNaN(value))) {
      return null;
    }

    const baseDate = String(saidaDataVal || "2000-01-01").trim();
    const saida = new Date(
      `${baseDate}T${String(sh).padStart(2, "0")}:${String(sm).padStart(2, "0")}:00`
    );
    if (Number.isNaN(saida.getTime())) {
      return null;
    }

    const tempoMs = ((th * 60) + tm) * 60000;
    const chegada = new Date(saida.getTime() + tempoMs);
    return {
      chegada_hora: chegada.toTimeString().slice(0, 5),
      chegada_data: saidaDataVal ? chegada.toISOString().slice(0, 10) : "",
    };
  }

  function applyArrivalToCard(card) {
    if (!card) {
      return;
    }

    const chegadaDataInput = card.querySelector('[data-role="chegada-data"]');
    const chegadaHoraInput = card.querySelector('[data-role="chegada-hora"]');
    const result = calcularChegada(
      card.querySelector('[data-role="saida-data"]')?.value || "",
      card.querySelector('[data-role="saida-hora"]')?.value || "",
      getTempoViagem()
    );

    if (!result) {
      if (chegadaDataInput) {
        chegadaDataInput.value = "";
      }
      if (chegadaHoraInput) {
        chegadaHoraInput.value = "";
      }
      return;
    }

    if (chegadaDataInput) {
      chegadaDataInput.value = result.chegada_data;
    }
    if (chegadaHoraInput) {
      chegadaHoraInput.value = result.chegada_hora;
    }
  }

  function applyArrivalToRetorno() {
    const retornoCard = $("retornoCard");
    if (!retornoCard || retornoCard.hidden) {
      return;
    }
    applyArrivalToCard(retornoCard);
  }

  function recalculateAllArrivals() {
    document.querySelectorAll('[data-role="trecho-card"]').forEach((card) => {
      applyArrivalToCard(card);
    });
    applyArrivalToRetorno();
  }

  function bindArrivalAutoCalculation(card) {
    if (!card || card.dataset.autoCalcBound === "1") {
      return;
    }

    card.dataset.autoCalcBound = "1";
    const saidaDataInput = card.querySelector('[data-role="saida-data"]');
    const saidaHoraInput = card.querySelector('[data-role="saida-hora"]');
    const recalculate = function () {
      if (card.id === "retornoCard") {
        applyArrivalToRetorno();
        return;
      }
      applyArrivalToCard(card);
    };

    saidaDataInput?.addEventListener("change", recalculate);
    saidaDataInput?.addEventListener("blur", recalculate);

    saidaHoraInput?.addEventListener("change", function () {
      if (this.value) {
        setLastSaidaHora(this.value);
      }
      recalculate();
    });
    saidaHoraInput?.addEventListener("blur", function () {
      if (this.value) {
        setLastSaidaHora(this.value);
      }
      recalculate();
    });
  }

  function renderTrechos() {
    const list = $("trechosList");
    const template = $("trechoTemplate");
    if (!list || !template) {
      return;
    }

    const existingDrafts = collectTrechoDrafts();
    const retornoDraft = captureRetorno();
    const legs = buildLegs();

    if (!legs.length) {
      showTrechosEmpty("Selecione a sede e adicione destinos para gerar os trechos automaticamente.");
      applyRetorno(retornoDraft);
      return;
    }

    list.innerHTML = "";
    const lastSaidaHora = getLastSaidaHora();

    legs.forEach((leg, index) => {
      const wrapper = document.createElement("div");
      wrapper.innerHTML = template.innerHTML.trim();
      const card = wrapper.firstElementChild;
      const routeKey = buildRouteKey(
        leg.origem.uf,
        leg.origem.cidade,
        leg.destino.uf,
        leg.destino.cidade
      );
      const draft =
        (existingDrafts[routeKey] && existingDrafts[routeKey].shift()) ||
        state.initialTrechos[index] ||
        {};

      card.dataset.origemEstado = leg.origem.uf;
      card.dataset.origemCidade = leg.origem.cidade;
      card.dataset.destinoEstado = leg.destino.uf;
      card.dataset.destinoCidade = leg.destino.cidade;

      const badge = card.querySelector('[data-role="trecho-badge"]');
      const preview = card.querySelector('[data-role="trecho-preview"]');
      const origemPreview = card.querySelector('[data-role="origem-local"]');
      const destinoPreview = card.querySelector('[data-role="destino-local"]');

      if (badge) {
        badge.textContent = index === 0 ? "Trecho 1 (Ida)" : `Trecho ${index + 1}`;
      }
      if (preview) {
        preview.textContent = `${leg.origem.cidade}/${leg.origem.uf} -> ${leg.destino.cidade}/${leg.destino.uf}`;
      }
      if (origemPreview) {
        origemPreview.textContent = `${leg.origem.cidade}/${leg.origem.uf}`;
      }
      if (destinoPreview) {
        destinoPreview.textContent = `${leg.destino.cidade}/${leg.destino.uf}`;
      }

      const saidaDataInput = card.querySelector('[data-role="saida-data"]');
      const saidaHoraInput = card.querySelector('[data-role="saida-hora"]');
      const chegadaDataInput = card.querySelector('[data-role="chegada-data"]');
      const chegadaHoraInput = card.querySelector('[data-role="chegada-hora"]');

      if (saidaDataInput) {
        saidaDataInput.value = draft.saida_data || "";
      }
      if (saidaHoraInput) {
        saidaHoraInput.value = draft.saida_hora || lastSaidaHora || "";
      }
      if (chegadaDataInput) {
        chegadaDataInput.value = draft.chegada_data || "";
      }
      if (chegadaHoraInput) {
        chegadaHoraInput.value = draft.chegada_hora || "";
      }

      bindArrivalAutoCalculation(card);
      list.appendChild(card);
    });

    state.initialTrechos = [];
    applyRetorno(
      retornoDraft.saida_data ||
        retornoDraft.saida_hora ||
        retornoDraft.chegada_data ||
        retornoDraft.chegada_hora
        ? retornoDraft
        : state.initialRetorno
    );
    if (getTempoViagem()) {
      recalculateAllArrivals();
    }
  }

  function destroySortable() {
    if (state.sortable) {
      state.sortable.destroy();
      state.sortable = null;
    }
  }

  async function renderDestinos() {
    const list = $("destinosList");
    const template = $("destinoTemplate");
    if (!list || !template) {
      return;
    }

    destroySortable();
    list.innerHTML = "";

    for (let index = 0; index < state.destinos.length; index += 1) {
      const destino = state.destinos[index];
      const wrapper = document.createElement("div");
      wrapper.innerHTML = template.innerHTML.trim();
      const card = wrapper.firstElementChild;
      const badge = card.querySelector('[data-role="destino-badge"]');
      const preview = card.querySelector('[data-role="destino-preview"]');
      const ufSelect = card.querySelector('[data-role="destino-uf"]');
      const cidadeSelect = card.querySelector('[data-role="destino-cidade"]');
      const removeButton = card.querySelector('[data-role="remove-destino"]');
      const errorLabel = card.querySelector('[data-role="destino-error"]');

      if (badge) {
        badge.textContent = `Destino ${index + 1}`;
      }
      if (ufSelect) {
        ufSelect.value = destino.uf || "PR";
      }

      try {
        await hydrateCitySelect(cidadeSelect, destino.uf || "PR", destino.cidade || "");
      } catch (error) {
        populateCitySelect(cidadeSelect, [], destino.cidade || "");
      }

      if (preview) {
        preview.textContent = destino.cidade ? `${destino.cidade}/${destino.uf}` : "UF + cidade";
      }

      ufSelect?.addEventListener("change", async function () {
        state.destinos[index].uf = String(this.value || "PR").trim().toUpperCase() || "PR";
        state.destinos[index].cidade = "";
        if (errorLabel) {
          errorLabel.hidden = true;
        }
        card.classList.remove("is-invalid");
        try {
          await hydrateCitySelect(cidadeSelect, state.destinos[index].uf, "");
        } catch (error) {
          populateCitySelect(cidadeSelect, [], "");
        }
        if (preview) {
          preview.textContent = "UF + cidade";
        }
        renderTrechos();
      });

      cidadeSelect?.addEventListener("change", function () {
        state.destinos[index].cidade = String(this.value || "").trim();
        if (errorLabel) {
          errorLabel.hidden = !!state.destinos[index].cidade;
        }
        card.classList.toggle("is-invalid", !state.destinos[index].cidade);
        if (preview) {
          preview.textContent = state.destinos[index].cidade
            ? `${state.destinos[index].cidade}/${state.destinos[index].uf}`
            : "UF + cidade";
        }
        renderTrechos();
      });

      removeButton?.addEventListener("click", function () {
        state.destinos.splice(index, 1);
        void renderDestinos().then(renderTrechos);
      });

      list.appendChild(card);
    }

    if (state.destinos.length > 1 && typeof window.Sortable === "function") {
      state.sortable = new window.Sortable(list, {
        animation: 150,
        handle: ".drag-handle",
        onEnd(event) {
          if (event.oldIndex === event.newIndex) {
            return;
          }
          const [moved] = state.destinos.splice(event.oldIndex, 1);
          state.destinos.splice(event.newIndex, 0, moved);
          void renderDestinos().then(renderTrechos);
        },
      });
    }
  }

  function validateForm() {
    if (!getSedeCidade()) {
      return "Informe a cidade da sede.";
    }
    if (!state.destinos.length) {
      return "Adicione ao menos um destino.";
    }

    let invalidDestino = false;
    document.querySelectorAll(".destino-card").forEach((card, index) => {
      const hasCity = Boolean(state.destinos[index] && state.destinos[index].cidade);
      const errorLabel = card.querySelector('[data-role="destino-error"]');
      card.classList.toggle("is-invalid", !hasCity);
      if (errorLabel) {
        errorLabel.hidden = hasCity;
      }
      if (!hasCity) {
        invalidDestino = true;
      }
    });

    if (invalidDestino) {
      return "Preencha todos os destinos antes de salvar.";
    }

    if (!document.querySelectorAll('[data-role="trecho-card"]').length) {
      return "Nenhum trecho foi gerado.";
    }

    return "";
  }

  function collectPayload() {
    const destinos = state.destinos
      .filter((item) => item.cidade)
      .map((item) => ({
        uf: item.uf,
        cidade: item.cidade,
      }));

    const trechos = Array.from(document.querySelectorAll('[data-role="trecho-card"]')).map((card) => ({
      origem_estado: card.dataset.origemEstado || getSedeUf(),
      origem_cidade: card.dataset.origemCidade || getSedeCidade(),
      destino_estado: card.dataset.destinoEstado || "PR",
      destino_cidade: card.dataset.destinoCidade || "",
      saida_data: card.querySelector('[data-role="saida-data"]')?.value || "",
      saida_hora: card.querySelector('[data-role="saida-hora"]')?.value || "",
      chegada_data: card.querySelector('[data-role="chegada-data"]')?.value || "",
      chegada_hora: card.querySelector('[data-role="chegada-hora"]')?.value || "",
    }));

    return {
      sede_uf: getSedeUf(),
      sede_cidade: getSedeCidade(),
      tempo_viagem: getTempoViagem(),
      destinos: destinos,
      trechos: trechos,
      retorno: captureRetorno(),
    };
  }

  async function salvarRoteiro() {
    const feedback = $("roteiro-save-feedback");
    const errorBox = $("roteiro-save-error");
    const button = $("btnSalvarRoteiro");
    const validationError = validateForm();

    setMessage(feedback, "");
    setMessage(errorBox, "");

    if (validationError) {
      setMessage(errorBox, validationError);
      return;
    }

    button.disabled = true;
    button.textContent = "Salvando...";

    try {
      const response = await fetch(window.ROTEIRO_FORM_SAVE_URL || "/roteiros/salvar/", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "X-CSRFToken": getCsrfToken(),
          "X-Requested-With": "XMLHttpRequest",
        },
        body: JSON.stringify(collectPayload()),
      });

      const contentType = response.headers.get("content-type") || "";
      if (!contentType.includes("application/json")) {
        const raw = await response.text();
        throw new Error(raw && raw.trim().startsWith("<") ? "O servidor retornou HTML em vez de JSON." : "Resposta invalida ao salvar o roteiro.");
      }

      const data = await response.json();
      if (!response.ok || !data.ok) {
        throw new Error(data.erro || data.error || "Nao foi possivel salvar o roteiro.");
      }

      setMessage(feedback, data.message || "Roteiro salvo com sucesso.");
      window.location.href =
        data.redirect_url || `/roteiros/${data.roteiro_id || data.id}/detalhe/`;
    } catch (error) {
      setMessage(errorBox, error.message || "Nao foi possivel salvar o roteiro.");
      button.disabled = false;
      button.textContent = button.dataset.originalLabel || "Salvar Roteiro";
    }
  }

  async function hydrateSede() {
    const sedeUfSelect = $("sedeUf");
    const sedeCidadeSelect = $("sedeCidade");
    if (!sedeUfSelect || !sedeCidadeSelect) {
      return;
    }

    const targetUf = String(window.SEDE_UF_INICIAL || sedeUfSelect.value || "PR")
      .trim()
      .toUpperCase();
    const targetCidade = String(window.SEDE_CIDADE_INICIAL || sedeCidadeSelect.value || "").trim();

    sedeUfSelect.value = targetUf || "PR";
    try {
      await hydrateCitySelect(sedeCidadeSelect, sedeUfSelect.value, targetCidade);
    } catch (error) {
      populateCitySelect(sedeCidadeSelect, [], targetCidade);
    }
    if (targetCidade) {
      sedeCidadeSelect.value = targetCidade;
    }
  }

  function hydrateDestinosFromInitial() {
    if (state.initialTrechos.length) {
      state.destinos = state.initialTrechos.map((trecho) => ({
        uf: String(trecho.destino_estado || "PR").trim().toUpperCase() || "PR",
        cidade: String(trecho.destino_cidade || "").trim(),
      }));
      return;
    }
    state.destinos = [];
  }

  async function init() {
    const form = getForm();
    if (!form) {
      return;
    }

    const saveButton = $("btnSalvarRoteiro");
    if (saveButton) {
      saveButton.dataset.originalLabel = saveButton.textContent.trim();
      saveButton.addEventListener("click", salvarRoteiro);
    }
    if ($("tempoViagem")) {
      $("tempoViagem").value = state.initialTempoViagem;
      $("tempoViagem").addEventListener("change", recalculateAllArrivals);
      $("tempoViagem").addEventListener("blur", recalculateAllArrivals);
    }
    bindArrivalAutoCalculation($("retornoCard"));

    await hydrateSede();
    hydrateDestinosFromInitial();
    await renderDestinos();
    renderTrechos();

    $("addDestinoBtn")?.addEventListener("click", async function () {
      state.destinos.push({ uf: "PR", cidade: "" });
      await renderDestinos();
      renderTrechos();
    });

    $("sedeUf")?.addEventListener("change", async function () {
      window.SEDE_UF_INICIAL = this.value;
      try {
        await hydrateCitySelect($("sedeCidade"), this.value, "");
      } catch (error) {
        populateCitySelect($("sedeCidade"), [], "");
      }
      $("sedeCidade").value = "";
      renderTrechos();
    });

    $("sedeCidade")?.addEventListener("change", function () {
      window.SEDE_CIDADE_INICIAL = this.value;
      renderTrechos();
    });
  }

  window.calcularChegada = calcularChegada;
  window.salvarRoteiro = salvarRoteiro;

  document.addEventListener("DOMContentLoaded", function () {
    void init();
  });
})();
