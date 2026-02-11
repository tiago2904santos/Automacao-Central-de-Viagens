(function () {
  function toDigits(value) {
    return (value || "").replace(/\D+/g, "");
  }

  function formatCpf(value) {
    const digits = toDigits(value).slice(0, 11);
    if (digits.length <= 3) return digits;
    if (digits.length <= 6) return `${digits.slice(0, 3)}.${digits.slice(3)}`;
    if (digits.length <= 9) {
      return `${digits.slice(0, 3)}.${digits.slice(3, 6)}.${digits.slice(6)}`;
    }
    return `${digits.slice(0, 3)}.${digits.slice(3, 6)}.${digits.slice(6, 9)}-${digits.slice(9)}`;
  }

  function formatPhone(value) {
    const digits = toDigits(value).slice(0, 11);
    if (digits.length <= 2) return digits;
    if (digits.length <= 6) return `(${digits.slice(0, 2)}) ${digits.slice(2)}`;
    if (digits.length <= 10) {
      return `(${digits.slice(0, 2)}) ${digits.slice(2, 6)}-${digits.slice(6)}`;
    }
    return `(${digits.slice(0, 2)}) ${digits.slice(2, 7)}-${digits.slice(7)}`;
  }

  function normalizeRg(value) {
    const cleaned = (value || "").toUpperCase().replace(/[^0-9X]+/g, "");
    if (!cleaned) return "";

    const trailingX = cleaned.endsWith("X");
    const digits = cleaned.replace(/X/g, "");
    if (trailingX) {
      return `${digits.slice(0, 9)}X`;
    }
    return digits.slice(0, 10);
  }

  function formatBase(base, firstGroupLen) {
    if (!base) return "";
    if (firstGroupLen === 1) {
      const b = base.slice(0, 7);
      if (b.length <= 1) return b;
      if (b.length <= 4) return `${b.slice(0, 1)}.${b.slice(1)}`;
      return `${b.slice(0, 1)}.${b.slice(1, 4)}.${b.slice(4)}`;
    }

    const b = base.slice(0, 8);
    if (b.length <= 2) return b;
    if (b.length <= 5) return `${b.slice(0, 2)}.${b.slice(2)}`;
    return `${b.slice(0, 2)}.${b.slice(2, 5)}.${b.slice(5)}`;
  }

  function formatRg(value) {
    const canon = normalizeRg(value);
    if (!canon) return "";

    if (canon.length <= 7) {
      return formatBase(canon, 1);
    }
    if (canon.length === 8) {
      const base = canon.slice(0, -1);
      const dv = canon.slice(-1);
      return `${formatBase(base, 1)}-${dv}`;
    }
    if (canon.length === 9) {
      const base = canon.slice(0, -1);
      const dv = canon.slice(-1);
      if (dv === "X") {
        return `${formatBase(base, 1)}-${dv}`;
      }
      return `${formatBase(base, 2)}-${dv}`;
    }

    const base = canon.slice(0, -1);
    const dv = canon.slice(-1);
    return `${formatBase(base, 2)}-${dv}`;
  }

  function normalizeUpper(value) {
    return (value || "").toUpperCase();
  }

  function normalizeOficio(value) {
    const raw = (value || "").replace(/[^\d/]+/g, "");
    if (!raw) return "";

    const slashIndex = raw.indexOf("/");
    if (slashIndex < 0) {
      return toDigits(raw);
    }

    const left = toDigits(raw.slice(0, slashIndex));
    const right = toDigits(raw.slice(slashIndex + 1)).slice(0, 4);
    if (!right && raw.endsWith("/")) {
      return `${left}/`;
    }
    if (!left && right) {
      return right;
    }
    return right ? `${left}/${right}` : left;
  }

  function formatProtocolo(value) {
    const digits = toDigits(value).slice(0, 9);
    if (digits.length <= 2) return digits;
    if (digits.length <= 5) return `${digits.slice(0, 2)}.${digits.slice(2)}`;
    if (digits.length <= 8) {
      return `${digits.slice(0, 2)}.${digits.slice(2, 5)}.${digits.slice(5)}`;
    }
    return `${digits.slice(0, 2)}.${digits.slice(2, 5)}.${digits.slice(5, 8)}-${digits.slice(8)}`;
  }

  function countTokenChars(text, end, tokenRegex) {
    let count = 0;
    const safeEnd = Math.max(0, Math.min(end, text.length));
    for (let i = 0; i < safeEnd; i += 1) {
      tokenRegex.lastIndex = 0;
      if (tokenRegex.test(text.charAt(i))) {
        count += 1;
      }
    }
    return count;
  }

  function positionFromTokenCount(text, tokenCount, tokenRegex) {
    if (tokenCount <= 0) {
      return 0;
    }
    let seen = 0;
    for (let i = 0; i < text.length; i += 1) {
      tokenRegex.lastIndex = 0;
      if (tokenRegex.test(text.charAt(i))) {
        seen += 1;
      }
      if (seen >= tokenCount) {
        return i + 1;
      }
    }
    return text.length;
  }

  function applyWithCaret(input, formatter, tokenRegex) {
    if (!input || typeof formatter !== "function") {
      return;
    }
    const before = input.value || "";
    const start = input.selectionStart ?? before.length;
    const tokenCount = tokenRegex
      ? countTokenChars(before, start, tokenRegex)
      : start;
    const after = formatter(before);
    if (after === before) {
      return;
    }
    input.value = after;
    const nextPos = tokenRegex
      ? positionFromTokenCount(after, tokenCount, tokenRegex)
      : Math.min(start, after.length);
    try {
      input.setSelectionRange(nextPos, nextPos);
    } catch (err) {
      // ignore input types that do not support selection
    }
  }

  function applyMask(input) {
    const mask = input.getAttribute("data-mask") || "";
    if (mask === "cpf") {
      applyWithCaret(input, formatCpf, /\d/);
      return;
    }
    if (mask === "telefone") {
      applyWithCaret(input, formatPhone, /\d/);
      return;
    }
    if (mask === "rg") {
      applyWithCaret(input, formatRg, /[A-Za-z0-9]/);
      return;
    }
    if (mask === "digits") {
      applyWithCaret(input, toDigits, /\d/);
      return;
    }
    if (mask === "oficio-num") {
      applyWithCaret(input, toDigits, /\d/);
      return;
    }
    if (mask === "oficio") {
      applyWithCaret(input, normalizeOficio, /[\d/]/);
      return;
    }
    if (mask === "protocolo") {
      applyWithCaret(input, formatProtocolo, /\d/);
      return;
    }
    if (input.getAttribute("data-uppercase") === "true") {
      applyWithCaret(input, normalizeUpper);
    }
  }

  function bindMaskInput(input) {
    if (!input || input.dataset.maskBound === "1") {
      return;
    }
    input.dataset.maskBound = "1";
    input.addEventListener("input", () => {
      applyMask(input);
    });
    applyMask(input);
  }

  function initInputMasks(root) {
    const scope = root && root.querySelectorAll ? root : document;
    scope
      .querySelectorAll("[data-mask], [data-uppercase='true']")
      .forEach(bindMaskInput);
  }

  window.initInputMasks = initInputMasks;

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => initInputMasks(document));
  } else {
    initInputMasks(document);
  }

  const observer = new MutationObserver((mutations) => {
    mutations.forEach((mutation) => {
      mutation.addedNodes.forEach((node) => {
        if (node.nodeType === Node.ELEMENT_NODE) {
          initInputMasks(node);
        }
      });
    });
  });

  if (document.body) {
    observer.observe(document.body, { childList: true, subtree: true });
  }
})();
