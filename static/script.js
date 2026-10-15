(() => {
  const ACCENT = "#073b9f";
  const GRID_COLOR = "#d7e1ef";
  const LABEL_COLOR = "#5f6c7e";
  const CHART_COLORS = [
    "#073b9f",
    "#ff9933",
    "#138808",
    "#2f6fc7",
    "#d43f3a",
    "#20a7b8",
    "#f4b740",
    "#5f6c7e",
    "#7d5cc6",
    "#2d9d78",
  ];

  let columns = [];
  let numericColumns = [];
  let chartInstances = {};
  let currentPage = 1;
  let totalPages = 1;
  let activeFilters = {};
  let printMode = false;
  let refreshVersion = 0;

  const loadingOverlay = document.getElementById("loading-overlay");
  const topBar = document.querySelector(".top-bar");
  const filterBar = document.getElementById("filter-bar");
  const filterControls = document.getElementById("filter-controls");
  const currentFileEl = document.getElementById("current-file");
  const filterToggleBtn = document.getElementById("filter-toggle");
  const exportPdfBtn = document.getElementById("export-pdf");
  const uploadBtn = document.getElementById("upload-btn");
  const fileInput = document.getElementById("file-input");
  const refreshDataBtn = document.getElementById("refresh-data");
  const DEFAULT_CHARTS = [
    { x: "Affiliation", y: "Count", type: "bar" },
    { x: "Year", y: "Count", type: "line" },
    { x: "Division (DRDO)", y: "Count", type: "doughnut" },
    { x: "Branch Category", y: "Count", type: "bar" },
    { x: "State", y: "Count", type: "bar" },
    { x: "Scientist Name", y: "Count", type: "pie" },
  ];

  const X_AXIS_PREFERENCES = [
    "Year",
    "Affiliation",
    "Division (DRDO)",
    "Branch Category",
    "State",
    "Scientist Name",
    "Mode of Training",
    "Certificate Issued",
  ];
  const Y_AXIS_PREFERENCES = ["Count", "Time in Weeks"];

  function chooseDefaultOption(options, preferences) {
    const normalized = options.map((c) => c.toLowerCase());
    for (const pref of preferences) {
      const matchIndex = normalized.indexOf(pref.toLowerCase());
      if (matchIndex !== -1) {
        return options[matchIndex];
      }
    }
    return options.length ? options[0] : "";
  }

  function showLoading(show) {
    loadingOverlay.classList.toggle("hidden", !show);
  }

  function updateTopBarHeight() {
    if (!topBar) return;
    const height = Math.ceil(topBar.getBoundingClientRect().height);
    document.documentElement.style.setProperty("--top-bar-height", `${height}px`);
  }

function updateHeaderScrollState() {
    if (!topBar) return;
    
    const isCurrentlyCompact = topBar.classList.contains("is-compact");
    
    // This proves the new code is running! 
    // You can see this by right-clicking your page -> Inspect -> Console
    console.log("Current Scroll:", Math.round(window.scrollY), "| Compact:", isCurrentlyCompact);
    
    // Pushed the shrink threshold all the way to 250px to absorb the massive height change
    if (!isCurrentlyCompact && window.scrollY > 250) {
      topBar.classList.add("is-compact");
      requestAnimationFrame(updateTopBarHeight);
    } else if (isCurrentlyCompact && window.scrollY <= 20) {
      topBar.classList.remove("is-compact");
      requestAnimationFrame(updateTopBarHeight);
    }
  }

  function setFilterPanelOpen(open) {
    if (open) updateTopBarHeight();
    filterBar.classList.toggle("hidden", !open);
    filterToggleBtn.classList.toggle("active", open);
    filterToggleBtn.setAttribute("aria-expanded", String(open));
  }

  function updateReportHeader() {
    document.getElementById("report-file").textContent =
      currentFileEl.textContent || "No file loaded";
    document.getElementById("report-generated").textContent =
      `Generated ${new Date().toLocaleString()}`;
  }

  async function exportDashboardPdf() {
    showLoading(true);
    try {
      await refreshAllCharts();
      updateReportHeader();
      setChartsPrintMode(true);
      await waitForPrintLayout();
      window.print();
    } catch (err) {
      console.error(err);
      alert(err.message || "Export failed");
    } finally {
      setTimeout(() => showLoading(false), 250);
    }
  }

  function formatNumber(n) {
    if (n == null || Number.isNaN(n)) return "—";
    return Number(n).toLocaleString(undefined, {
      maximumFractionDigits: 2,
    });
  }

  function filterQueryString() {
    const params = new URLSearchParams();
    Object.entries(activeFilters).forEach(([key, val]) => {
      if (val && val !== "All") params.set(key, val);
    });
    return params.toString();
  }

  function withCacheBuster(url) {
    const sep = url.includes("?") ? "&" : "?";
    return `${url}${sep}_t=${Date.now()}`;
  }

  async function fetchJson(url, options = {}) {
    const res = await fetch(withCacheBuster(url), {
      cache: "no-store",
      ...options,
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Request failed");
    return data;
  }

  async function loadMeta() {
    const meta = await fetchJson("/meta");
    const displayPath = meta.filepath || meta.filename;
    currentFileEl.textContent = displayPath || "No file loaded";
    currentFileEl.title = displayPath || "No file loaded";
  }

  async function loadFilters() {
    const filters = await fetchJson("/filters");
    filterControls.innerHTML = "";

    const priority = [
      "Year",
      "Gender",
      "State",
      "Branch Category",
      "Field / Branch of Engg.",
      "Division (DRDO)",
      "Mode of Training",
      "Certificate Issued",
    ];
    const keys = [
      ...priority.filter((k) => filters[k]),
      ...Object.keys(filters).filter((k) => !priority.includes(k)),
    ];

    keys.forEach((col) => {
      const group = document.createElement("div");
      group.className = "filter-group";

      const label = document.createElement("label");
      label.textContent = col;
      label.setAttribute("for", `filter-${col}`);

      const select = document.createElement("select");
      select.id = `filter-${col}`;
      select.dataset.column = col;

      const allOpt = document.createElement("option");
      allOpt.value = "All";
      allOpt.textContent = "All";
      select.appendChild(allOpt);

      filters[col].forEach((val) => {
        const opt = document.createElement("option");
        opt.value = val;
        opt.textContent = val;
        select.appendChild(opt);
      });

      if (activeFilters[col]) {
        select.value = activeFilters[col];
      }

      group.append(label, select);
      filterControls.appendChild(group);
    });
  }

  async function loadColumns() {
    const data = await fetchJson("/columns");
    columns = data.columns || [];
    numericColumns = data.numeric || [];

    document.querySelectorAll(".chart-card").forEach((card, idx) => {
      const xSel = card.querySelector(".x-axis");
      const ySel = card.querySelector(".y-axis");
      const typeSel = card.querySelector(".chart-type");
      const defaults = DEFAULT_CHARTS[idx] || DEFAULT_CHARTS[0];

      xSel.innerHTML = columns
        .map((c) => `<option value="${c}">${c}</option>`)
        .join("");
      ySel.innerHTML = numericColumns
        .map((c) => `<option value="${c}">${c}</option>`)
        .join("");

      const fallbackX = chooseDefaultOption(columns, X_AXIS_PREFERENCES);
      const fallbackY = chooseDefaultOption(numericColumns, Y_AXIS_PREFERENCES);

      if (columns.includes(defaults.x)) {
        xSel.value = defaults.x;
      } else {
        xSel.value = fallbackX;
      }
      if (numericColumns.includes(defaults.y)) {
        ySel.value = defaults.y;
      } else {
        ySel.value = fallbackY;
      }
      typeSel.value = defaults.type;

      xSel.addEventListener("change", () => refreshChart(card));
      ySel.addEventListener("change", () => refreshChart(card));
      typeSel.addEventListener("change", () => refreshChart(card));
      card.querySelector(".refresh-chart").addEventListener("click", () =>
        refreshChart(card)
      );
    });
  }

  async function loadKpis() {
    const qs = filterQueryString();
    const data = await fetchJson(qs ? `/kpi?${qs}` : "/kpi");
    document.getElementById("kpi-total-students").textContent = formatNumber(
      data.total_students
    );
    document.getElementById("kpi-female-interns").textContent = formatNumber(
      data.female_interns
    );
    document.getElementById("kpi-male-interns").textContent = formatNumber(
      data.male_interns
    );
    document.getElementById("kpi-branches-involved").textContent = formatNumber(
      data.branches_involved
    );
    document.getElementById("kpi-institutes-involved").textContent = formatNumber(
      data.institutes_involved
    );
    document.getElementById("kpi-drdo-divisions").textContent = formatNumber(
      data.drdo_divisions
    );
    document.getElementById("kpi-certificates-issued").textContent = formatNumber(
      data.certificates_issued
    );
    document.getElementById("kpi-average-weeks").textContent = formatNumber(
      data.average_weeks
    );
  }

  async function loadTable(page = 1) {
    currentPage = page;
    const qs = filterQueryString();
    const data = await fetchJson(`/data?page=${page}&per_page=10${qs ? `&${qs}` : ""}`);

    totalPages = data.pages;
    const head = document.getElementById("table-head");
    const body = document.getElementById("table-body");

    if (!data.data.length && data.total === 0) {
      head.innerHTML = "";
      body.innerHTML =
        '<tr><td colspan="99" style="text-align:center;padding:2rem;color:#9aa3b5">No data</td></tr>';
    } else {
      const cols = Object.keys(data.data[0]);
      head.innerHTML = `<tr>${cols.map((c) => `<th>${c}</th>`).join("")}</tr>`;
      body.innerHTML = data.data
        .map(
          (row) =>
            `<tr>${cols.map((c) => `<td>${row[c] ?? ""}</td>`).join("")}</tr>`
        )
        .join("");
    }

    document.getElementById("page-info").textContent =
      `Page ${data.page} of ${data.pages}`;
    document.getElementById("prev-page").disabled = data.page <= 1;
    document.getElementById("next-page").disabled = data.page >= data.pages;
  }

  function chartConfig(type, labels, values, title) {
    const isCircular = type === "pie" || type === "doughnut";
    const dataset = {
      label: title,
      data: values,
      backgroundColor: CHART_COLORS,
      borderColor: isCircular ? "#ffffff" : ACCENT,
      borderWidth: isCircular ? 2 : 2,
      tension: 0.3,
      fill: type === "line",
    };

    return {
      type,
      data: {
        labels,
        datasets: [dataset],
      },
      options: {
  responsive: true,
  maintainAspectRatio: false,
  animation: {
    duration: 0,
  },
  animations: false,
        transitions: {
          active: {
            animation: {
              duration: 0,
            },
          },
          resize: {
            animation: {
              duration: 0,
            },
          },
        },
        plugins: {
          legend: {
            display: isCircular,
            position: "right",
            labels: {
              color: printMode ? "#374151" : LABEL_COLOR,
              boxWidth: printMode ? 8 : 12,
              font: {
                size: printMode ? 8 : 12,
              },
            },
          },
        },
        scales: isCircular
          ? {}
          : {
              x: {
                ticks: { color: LABEL_COLOR, maxRotation: 45 },
                grid: { color: GRID_COLOR },
              },
              y: {
                ticks: { color: LABEL_COLOR },
                grid: { color: GRID_COLOR },
                beginAtZero: true,
              },
            },
      },
    };
  }

  function setChartsPrintMode(enabled) {
    printMode = enabled;
    Object.values(chartInstances).forEach((chart) => {
      const isCircular = chart.config.type === "pie" || chart.config.type === "doughnut";
      chart.canvas.closest(".chart-card")?.classList.toggle("circular-chart", isCircular);
      chart.options.animation = { duration: 0 };
      chart.options.animations = false;
      if (isCircular) {
        chart.options.plugins.legend.display = true;
        chart.options.plugins.legend.position = "right";
        chart.options.plugins.legend.labels.color = enabled ? "#374151" : LABEL_COLOR;
        chart.options.plugins.legend.labels.boxWidth = enabled ? 8 : 12;
        chart.options.plugins.legend.labels.font = { size: enabled ? 8 : 12 };
      }
      chart.resize();
      chart.update("none");
    });
  }

  function waitForPrintLayout() {
    return new Promise((resolve) => {
      requestAnimationFrame(() => {
        requestAnimationFrame(() => {
          setTimeout(resolve, 500);
        });
      });
    });
  }

  async function refreshChart(card){
  const version = refreshVersion;
    const x = card.querySelector(".x-axis").value;
    const y = card.querySelector(".y-axis").value;
    const type = card.querySelector(".chart-type").value;
    const canvas = card.querySelector("canvas");
    const id = card.dataset.chartId;
    card.classList.toggle("circular-chart", type === "pie" || type === "doughnut");

    if (!x || !y) {
      if (chartInstances[id]) {
        chartInstances[id].destroy();
        delete chartInstances[id];
      }
      card.querySelector("[data-title]").textContent = "Select X and Y axis";
      return;
    }

    const title = `${y} by ${x}`;
    card.querySelector("[data-title]").textContent = title;

    const params = new URLSearchParams({ x_axis: x, y_axis: y });
    Object.entries(activeFilters).forEach(([k, v]) => {
      if (v && v !== "All") params.set(k, v);
    });

    const chartData = await fetchJson(`/chart-data?${params}`);
if (version !== refreshVersion) return;
if (!Array.isArray(chartData.labels)) {
      throw new Error("Failed to load chart data");
    }

    const existingChart = chartInstances[id];
if (existingChart && existingChart.config.type === type) {
  existingChart.data.labels = chartData.labels;
  existingChart.data.datasets[0].label = title;
  existingChart.data.datasets[0].data = chartData.values;
  existingChart.update("none");
  return;
}

if (existingChart) {
  existingChart.destroy();
}

chartInstances[id] = new Chart(
  canvas,
  chartConfig(type, chartData.labels, chartData.values, title)
);
  }

  async function refreshAllCharts() {
    const cards = document.querySelectorAll(".chart-card");
    await Promise.all([...cards].map((card) => refreshChart(card)));
  }

  async function refreshDashboard({ showOverlay = false } = {}) {
  refreshVersion += 1;
  if (showOverlay) showLoading(true);
    filterBar.classList.add("updating");
    try {
      await Promise.all([
        loadKpis(),
        loadTable(currentPage),
        refreshAllCharts(),
      ]);
    } finally {
      filterBar.classList.remove("updating");
      if (showOverlay) showLoading(false);
    }
  }

  async function init() {
    updateTopBarHeight();
    updateHeaderScrollState();
    showLoading(true);
    try {
      await loadMeta();
      await loadFilters();
      await loadColumns();
      await refreshDashboard();
    } catch (err) {
      console.error(err);
      alert(err.message || "Failed to load dashboard");
    } finally {
      showLoading(false);
    }
  }

  async function refreshFromSource() {
    showLoading(true);
    try {
      await fetchJson("/reload");
      activeFilters = {};
      currentPage = 1;
      await loadMeta();
      await loadFilters();
      await loadColumns();
      await refreshDashboard();
    } catch (err) {
      console.error(err);
      alert(err.message || "Refresh failed");
    } finally {
      showLoading(false);
    }
  }

  document.getElementById("apply-filters").addEventListener("click", () => {
    activeFilters = {};
    filterControls.querySelectorAll("select").forEach((sel) => {
      activeFilters[sel.dataset.column] = sel.value;
    });
    currentPage = 1;
    refreshDashboard();
  });

  document.getElementById("reset-filters").addEventListener("click", () => {
    activeFilters = {};
    filterControls.querySelectorAll("select").forEach((sel) => {
      sel.value = "All";
    });
    currentPage = 1;
    refreshDashboard();
  });

  filterToggleBtn.addEventListener("click", () => {
    setFilterPanelOpen(filterBar.classList.contains("hidden"));
  });

  exportPdfBtn.addEventListener("click", exportDashboardPdf);
  window.addEventListener("beforeprint", () => setChartsPrintMode(true));
  window.addEventListener("afterprint", () => setChartsPrintMode(false));
  window.addEventListener("resize", updateTopBarHeight);
  window.addEventListener("scroll", updateHeaderScrollState, { passive: true });
  document.querySelectorAll(".brand-logo img").forEach((img) => {
    img.addEventListener("load", updateTopBarHeight);
  });

  document.getElementById("prev-page").addEventListener("click", () => {
    if (currentPage > 1) loadTable(currentPage - 1);
  });

  document.getElementById("next-page").addEventListener("click", () => {
    if (currentPage < totalPages) loadTable(currentPage + 1);
  });

  refreshDataBtn.addEventListener("click", refreshFromSource);

  uploadBtn.addEventListener("click", async () => {
    showLoading(true);
    try {
      await fetchJson("/pick-file", { method: "POST" });
      activeFilters = {};
      currentPage = 1;
      await loadMeta();
      await loadFilters();
      await loadColumns();
      await refreshDashboard();
    } catch (err) {
      // Fallback: if server file picker is unavailable, use browser upload.
      console.warn("Server file picker fallback:", err);
      fileInput.click();
    } finally {
      showLoading(false);
    }
  });

  fileInput.addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    showLoading(true);
    const form = new FormData();
    form.append("file", file);

    try {
      await fetchJson("/upload", { method: "POST", body: form });

      activeFilters = {};
      currentPage = 1;
      await loadMeta();
      await loadFilters();
      await loadColumns();
      await refreshDashboard();
    } catch (err) {
      alert(err.message || "Upload failed");
    } finally {
      showLoading(false);
      fileInput.value = "";
    }
  });

  init();
})();
