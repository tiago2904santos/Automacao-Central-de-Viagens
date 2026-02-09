(function () {
  const form = document.getElementById("configOficioForm");
  if (!form) {
    return;
  }

  const unidadeInput = form.querySelector("#id_unidade_nome");
  const origemInput = form.querySelector("#id_origem_nome");
  const cepInput = form.querySelector("#id_cep");
  const logradouroInput = form.querySelector("#id_logradouro");
  const bairroInput = form.querySelector("#id_bairro");
  const cidadeInput = form.querySelector("#id_cidade");
  const ufInput = form.querySelector("#id_uf");
  const numeroInput = form.querySelector("#id_numero");
  const complementoInput = form.querySelector("#id_complemento");
  const telefoneInput = form.querySelector("#id_telefone");
  const emailInput = form.querySelector("#id_email");
  const assinanteSelect = form.querySelector("#id_assinante");
  const cepStatus = document.getElementById("cepStatus");

  const previewUnidade = document.getElementById("previewUnidade");
  const previewOrigem = document.getElementById("previewOrigem");
  const previewRodapeUnidade = document.getElementById("previewRodapeUnidade");
  const previewEndereco = document.getElementById("previewEndereco");
  const previewTelefone = document.getElementById("previewTelefone");
  const previewEmail = document.getElementById("previewEmail");
  const previewAssinanteNome = document.getElementById("previewAssinanteNome");
  const previewAssinanteCargo = document.getElementById("previewAssinanteCargo");

  const particles = new Set(["da", "de", "do", "dos", "das", "e"]);

  const titleCasePt = (text) => {
    if (!text) return "";
    const words = text
      .trim()
      .toLowerCase()
      .split(/\s+/)
      .filter(Boolean);
    return words
      .map((word, idx) => {
        if (idx > 0 && particles.has(word)) {
          return word;
        }
        return word.charAt(0).toUpperCase() + word.slice(1);
      })
      .join(" ");
  };

  const setPreview = (el, value) => {
    if (!el) return;
    el.textContent = value ? value : "-";
  };

  const forceUppercase = (input) => {
    if (!input) return;
    const start = input.selectionStart;
    const end = input.selectionEnd;
    const upper = input.value.toUpperCase();
    if (upper !== input.value) {
      input.value = upper;
      if (start !== null && end !== null) {
        input.setSelectionRange(start, end);
      }
    }
  };

  const formatEndereco = () => {
    const logradouro = (logradouroInput?.value || "").trim();
    const numero = (numeroInput?.value || "").trim();
    const complemento = (complementoInput?.value || "").trim();
    const bairro = (bairroInput?.value || "").trim();
    const cidade = (cidadeInput?.value || "").trim();
    const uf = (ufInput?.value || "").trim();
    const cep = (cepInput?.value || "").trim();

    const numeroParte = numero
      ? `${numero}${complemento ? " " + complemento : ""}`
      : "";
    const primeiraParte = [logradouro, numeroParte].filter(Boolean).join(", ");
    const cidadeUf = cidade && uf ? `${cidade}/${uf}` : cidade || uf;

    let endereco = [primeiraParte, bairro, cidadeUf].filter(Boolean).join(" - ");
    if (cep) {
      endereco = endereco ? `${endereco} - CEP ${cep}` : `CEP ${cep}`;
    }
    return endereco;
  };

  const updatePreview = () => {
    const unidade = (unidadeInput?.value || "").trim().toUpperCase();
    const origem = (origemInput?.value || "").trim().toUpperCase();
    setPreview(previewUnidade, unidade);
    setPreview(previewOrigem, origem);
    setPreview(previewRodapeUnidade, titleCasePt(unidade));
    setPreview(previewEndereco, formatEndereco());
    setPreview(previewTelefone, (telefoneInput?.value || "").trim());
    setPreview(previewEmail, (emailInput?.value || "").trim());

    const selectedOption = assinanteSelect?.selectedOptions?.[0];
    const hasValue = selectedOption && selectedOption.value;
    const label = hasValue ? selectedOption.textContent.trim() : "";
    if (!label) {
      setPreview(previewAssinanteNome, "");
      setPreview(previewAssinanteCargo, "");
      return;
    }
    const parts = label.split(" - ");
    const nome = parts[0] || "";
    const cargo = parts.slice(1).join(" - ");
    setPreview(previewAssinanteNome, titleCasePt(nome));
    setPreview(previewAssinanteCargo, titleCasePt(cargo));
  };

  const setCepStatus = (message, isError) => {
    if (!cepStatus) return;
    cepStatus.textContent = message || "";
    cepStatus.style.color = isError ? "#ef4444" : "";
  };

  const debounce = (fn, wait = 400) => {
    let timer = null;
    return (...args) => {
      window.clearTimeout(timer);
      timer = window.setTimeout(() => fn(...args), wait);
    };
  };

  const fetchCep = async () => {
    if (!cepInput) return;
    const digits = (cepInput.value || "").replace(/\D/g, "");
    if (!digits) {
      setCepStatus("", false);
      updatePreview();
      return;
    }
    if (digits.length !== 8) {
      setCepStatus("CEP deve conter 8 digitos.", true);
      updatePreview();
      return;
    }
    const formatted = `${digits.slice(0, 5)}-${digits.slice(5)}`;
    cepInput.value = formatted;
    setCepStatus("Buscando CEP...", false);
    try {
      const response = await fetch(`/api/cep/${digits}/`);
      if (!response.ok) {
        throw new Error("CEP nao encontrado");
      }
      const data = await response.json();
      if (data.error) {
        throw new Error("CEP nao encontrado");
      }
      if (logradouroInput) logradouroInput.value = data.logradouro || "";
      if (bairroInput) bairroInput.value = data.bairro || "";
      if (cidadeInput) cidadeInput.value = data.cidade || "";
      if (ufInput) ufInput.value = data.uf || "";
      setCepStatus("CEP atualizado.", false);
    } catch (err) {
      setCepStatus("CEP nao encontrado.", true);
    }
    updatePreview();
  };

  const bindUppercase = (input) => {
    if (!input) return;
    input.addEventListener("input", () => {
      forceUppercase(input);
      updatePreview();
    });
  };

  bindUppercase(unidadeInput);
  bindUppercase(origemInput);

  [
    cepInput,
    logradouroInput,
    bairroInput,
    cidadeInput,
    ufInput,
    numeroInput,
    complementoInput,
    telefoneInput,
    emailInput,
  ].forEach((input) => {
    if (!input) return;
    input.addEventListener("input", updatePreview);
  });

  if (assinanteSelect) {
    assinanteSelect.addEventListener("change", updatePreview);
  }

  if (cepInput) {
    cepInput.addEventListener("blur", fetchCep);
    cepInput.addEventListener("change", fetchCep);
    const debouncedFetch = debounce(fetchCep, 400);
    cepInput.addEventListener("input", () => {
      const digits = (cepInput.value || "").replace(/\D/g, "");
      if (digits.length >= 8) {
        debouncedFetch();
      }
    });
  }

  forceUppercase(unidadeInput);
  forceUppercase(origemInput);
  updatePreview();
})();
