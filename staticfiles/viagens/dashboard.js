function formatNumber(value) {
  const num = Number(value || 0);
  return num.toLocaleString("pt-BR");
}

function renderBarChart(container, series) {
  if (!container) {
    return;
  }
  const data = Array.isArray(series) ? series : [];
  if (!data.length) {
    container.innerHTML = '<div class="empty">Sem dados no periodo.</div>';
    return;
  }

  const max = Math.max(...data.map((item) => Number(item.total || 0)), 1);
  const width = 720;
  const height = 220;
  const padding = { top: 16, right: 12, bottom: 28, left: 28 };
  const chartWidth = width - padding.left - padding.right;
  const chartHeight = height - padding.top - padding.bottom;
  const barGap = 6;
  const barWidth = Math.max(8, (chartWidth - barGap * (data.length - 1)) / data.length);

  let bars = "";
  let labels = "";
  data.forEach((item, index) => {
    const total = Number(item.total || 0);
    const x = padding.left + index * (barWidth + barGap);
    const barHeight = (total / max) * chartHeight;
    const y = padding.top + (chartHeight - barHeight);
    const labelY = height - 6;
    const label = item.label || item.dia || "";
    bars += `\n      <rect class="chart-bar" x="${x}" y="${y}" width="${barWidth}" height="${barHeight}" rx="4" ry="4">\n        <title>${label}: ${total}</title>\n      </rect>`;
    if (data.length <= 12 || index % Math.ceil(data.length / 8) === 0) {
      labels += `\n        <text class="chart-label" x="${x + barWidth / 2}" y="${labelY}" text-anchor="middle">${label}</text>`;
    }
  });

  const gridLines = [0.25, 0.5, 0.75, 1]
    .map((ratio) => {
      const y = padding.top + chartHeight * (1 - ratio);
      return `<line class="chart-grid" x1="${padding.left}" x2="${width - padding.right}" y1="${y}" y2="${y}" />`;
    })
    .join("");

  container.innerHTML = `
    <svg class="chart-svg" viewBox="0 0 ${width} ${height}" role="img" aria-hidden="true">
      ${gridLines}
      ${bars}
      ${labels}
    </svg>
  `;
}

async function fetchDashboardData(endpoint, year) {
  const url = new URL(endpoint, window.location.origin);
  url.searchParams.set("ano", String(year));
  const response = await fetch(url.toString(), {
    headers: { "X-Requested-With": "XMLHttpRequest" },
    signal: fetchDashboardData._controller?.signal,
  });
  if (!response.ok) {
    throw new Error("Falha ao carregar dados do painel.");
  }
  return response.json();
}

function updateKpis(root, payload) {
  const kpis = payload.kpis || {};
  root.querySelectorAll("[data-kpi]").forEach((card) => {
    const key = card.getAttribute("data-kpi");
    const data = kpis[key] || { total: 0, periodo: 0, rotulo_periodo: "" };
    const valueEl = card.querySelector("[data-kpi-value]");
    const metaEl = card.querySelector("[data-kpi-meta]");
    if (valueEl) {
      valueEl.textContent = formatNumber(data.total);
    }
    if (metaEl) {
      metaEl.textContent = `${formatNumber(data.periodo)} ${data.rotulo_periodo}`.trim();
    }
  });
}

function updateCharts(root, payload) {
  const series = payload.series || {};
  renderBarChart(root.querySelector('[data-chart="oficios"]'), series.oficios);
  renderBarChart(root.querySelector('[data-chart="ranking"]'), series.ranking);
}

function initDashboard() {
  const root = document.querySelector("[data-dashboard]");
  if (!root) {
    return;
  }
  const endpoint = root.getAttribute("data-endpoint");
  const initialYear = Number(root.getAttribute("data-year") || new Date().getFullYear());
  if (!endpoint) {
    return;
  }

  const initialEl = document.getElementById("dashboard-initial");
  const initialPayload = initialEl ? JSON.parse(initialEl.textContent || "{}") : null;
  let currentYear = initialPayload?.year || initialYear;
  if (initialPayload) {
    updateKpis(root, initialPayload);
    updateCharts(root, initialPayload);
  }

  const updateAll = (payload) => {
    updateKpis(root, payload);
    updateCharts(root, payload);
  };

  const loadYear = async (year) => {
    fetchDashboardData._controller?.abort?.();
    fetchDashboardData._controller = new AbortController();
    currentYear = year;
    root.classList.add("is-loading");
    try {
      const payload = await fetchDashboardData(endpoint, year);
      updateAll(payload);
    } catch (err) {
      if (window.showToast) {
        window.showToast("Nao foi possivel atualizar o painel.", "error");
      }
    } finally {
      root.classList.remove("is-loading");
      fetchDashboardData._controller = null;
    }
  };

  if (root.hasAttribute("data-refresh")) {
    loadYear(currentYear);
  }
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initDashboard);
} else {
  initDashboard();
}
