const PRIMARY = "#1f77b4";
const SUCCESS = "#2ca02c";
const WARNING = "#ff9f1c";
const DANGER = "#d62728";
const NEUTRAL = "#7f8c8d";
const SLA_THRESHOLD = 0.8;

const COL_START = 6;
const COL_END = 14;

function secondsFromTime(value) {
  if (value == null || value === "") return 0;
  if (typeof value === "number") {
    return value < 1 ? value * 86400 : value;
  }
  if (value instanceof Date) {
    return value.getHours() * 3600 + value.getMinutes() * 60 + value.getSeconds();
  }
  const text = String(value).trim();
  const parts = text.split(":").map((p) => parseInt(p, 10));
  if (parts.some((p) => Number.isNaN(p))) return 0;
  while (parts.length < 3) parts.unshift(0);
  const [h, m, s] = parts.slice(-3);
  return h * 3600 + m * 60 + s;
}

function formatSeconds(totalSeconds) {
  const total = Math.round(Number(totalSeconds) || 0);
  const h = Math.floor(total / 3600);
  const rem = total % 3600;
  const m = Math.floor(rem / 60);
  const s = rem % 60;
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}:${String(s).padStart(2, "0")}`;
}

function toNumber(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function excelDateToString(value) {
  if (value == null || value === "") return null;
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return String(value);
    const d = new Date(parsed.y, parsed.m - 1, parsed.d);
    return d.toISOString().slice(0, 10);
  }
  return String(value).trim();
}

function chartHeight(df) {
  return Math.max(440, df.length * 38 + 120);
}

function plotLayout(title, extra = {}) {
  return {
    title: { text: title, x: 0.02, xanchor: "left", font: { size: 16 } },
    font: { family: "Tahoma, Segoe UI, sans-serif", size: 12 },
    margin: { t: 60, b: 48, l: 8, r: 48 },
    plot_bgcolor: "#ffffff",
    paper_bgcolor: "#ffffff",
    showlegend: true,
    legend: { orientation: "h", yanchor: "bottom", y: 1.02, xanchor: "left", x: 0 },
    ...extra,
  };
}

function channelYAxis(channels) {
  return {
    automargin: true,
    tickfont: { size: 12 },
    ticklabelposition: "outside",
    type: "category",
    categoryorder: "array",
    categoryarray: channels,
    title: "",
  };
}

function valueXAxis(title, extra = {}) {
  return {
    title,
    gridcolor: "#eee",
    automargin: true,
    zeroline: false,
    ...extra,
  };
}

function cellValue(sheet, row, col) {
  const cell = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
  if (!cell) return null;
  return cell.v;
}

function normalizeKey(key) {
  const trimmed = String(key || "").trim();
  const compact = trimmed.replace(/\s+/g, "").toLowerCase();
  const aliases = {
    averageresponse: "Average Response",
    servicelevel: "Service Level",
  };
  return aliases[compact] || trimmed;
}

function parseServiceLevel(value) {
  if (value == null || value === "") return 0;
  if (typeof value === "number") return value > 1 ? value / 100 : value;
  const text = String(value).trim().replace("%", "");
  const n = Number(text);
  if (!Number.isFinite(n)) return 0;
  return n > 1 ? n / 100 : n;
}

function readChannelsReportSheet(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
  const headerRow = 2;
  const headers = [];
  for (let col = COL_START; col <= COL_END; col++) {
    const raw = cellValue(sheet, headerRow, col);
    headers.push(raw == null ? `col_${col - COL_START}` : normalizeKey(raw));
  }

  const records = [];
  for (let row = headerRow + 1; row <= range.e.r; row++) {
    const obj = {};
    headers.forEach((name, idx) => {
      obj[name] = cellValue(sheet, row, COL_START + idx);
    });
    if (obj.Channel == null || String(obj.Channel).trim() === "") continue;
    records.push(obj);
  }
  return records;
}

function readDataSheet(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
  const headers = [];
  for (let col = 0; col <= range.e.c; col++) {
    const raw = cellValue(sheet, 0, col);
    headers.push(raw == null ? `col_${col}` : normalizeKey(raw));
  }

  const records = [];
  for (let row = 1; row <= range.e.r; row++) {
    const obj = {};
    headers.forEach((name, idx) => {
      obj[name] = cellValue(sheet, row, idx);
    });
    if (obj.Channel == null || String(obj.Channel).trim() === "") continue;
    records.push(obj);
  }
  return records;
}

function finalizeRecords(records) {
  let lastDate = null;
  for (const row of records) {
    if (row.Date != null && row.Date !== "") lastDate = row.Date;
    else row.Date = lastDate;
    row.Date = excelDateToString(row.Date);
    row.Channel = String(row.Channel).trim();

    const extraKey = Object.keys(row).find((k) =>
      /^col_\d+$/i.test(k) || k.toLowerCase().startsWith("unnamed")
    );
    if (extraKey) {
      row.Pending = toNumber(row.Pending) + toNumber(row[extraKey]);
      delete row[extraKey];
    }

    ["Incoming", "Closed", "Pending", "Backlog"].forEach((col) => {
      row[col] = Math.round(toNumber(row[col]));
    });

    if (row.ResponseSeconds != null && row["Average Response"] == null) {
      row["Average Response"] = row.ResponseSeconds;
    }
    if (row.AverageResponse != null && row["Average Response"] == null) {
      row["Average Response"] = row.AverageResponse;
    }

    row["Service Level"] = parseServiceLevel(
      row["Service Level"] ?? row.ServiceLevel
    );
    row.ResponseSeconds = secondsFromTime(row["Average Response"]);
    row.AverageResponseText = formatSeconds(row.ResponseSeconds);
    row.ClosureRate = row.Incoming > 0 ? row.Closed / row.Incoming : 0;
  }
  return records;
}

function loadChannelsReport(workbook) {
  if (workbook.Sheets["Channels Report"]) {
    return finalizeRecords(readChannelsReportSheet(workbook.Sheets["Channels Report"]));
  }
  if (workbook.Sheets.Data) {
    return finalizeRecords(readDataSheet(workbook.Sheets.Data));
  }
  const fallback = workbook.SheetNames[0];
  if (!fallback) throw new Error("The workbook has no sheets.");
  throw new Error(
    "Could not find a 'Channels Report' or 'Data' sheet. Available sheets: " +
      workbook.SheetNames.join(", ")
  );
}

function renderChart(id, figure) {
  const el = document.getElementById(id);
  const height = figure.layout?.height || 440;
  el.style.height = `${height}px`;
  const card = el.closest(".chart-card");
  if (card) card.style.minHeight = `${height + 48}px`;
  Plotly.newPlot(el, figure.data, figure.layout, {
    responsive: true,
    displayModeBar: false,
  });
}

function chartIncomingVsClosed(df) {
  const sorted = [...df].sort((a, b) => a.Incoming - b.Incoming);
  const channels = sorted.map((r) => r.Channel);
  return {
    data: [
      {
        type: "bar",
        orientation: "h",
        name: "Incoming",
        y: channels,
        x: sorted.map((r) => r.Incoming),
        marker: { color: PRIMARY },
        hovertemplate: "<b>%{y}</b><br>Incoming: %{x}<extra></extra>",
      },
      {
        type: "bar",
        orientation: "h",
        name: "Closed",
        y: channels,
        x: sorted.map((r) => r.Closed),
        marker: { color: SUCCESS },
        hovertemplate: "<b>%{y}</b><br>Closed: %{x}<extra></extra>",
      },
    ],
    layout: plotLayout("Incoming vs closed by channel", {
      height: chartHeight(df),
      barmode: "group",
      bargap: 0.22,
      bargroupgap: 0.08,
      xaxis: valueXAxis("Ticket count"),
      yaxis: channelYAxis(channels),
    }),
  };
}

function chartServiceLevel(df) {
  const sorted = [...df].sort((a, b) => a["Service Level"] - b["Service Level"]);
  const channels = sorted.map((r) => r.Channel);
  return {
    data: [
      {
        type: "bar",
        orientation: "h",
        x: sorted.map((r) => r["Service Level"] * 100),
        y: channels,
        marker: {
          color: sorted.map((r) =>
            r["Service Level"] >= SLA_THRESHOLD ? SUCCESS : DANGER
          ),
        },
        text: sorted.map((r) => `${Math.round(r["Service Level"] * 100)}%`),
        textposition: "outside",
        cliponaxis: false,
        textfont: { size: 11 },
        hovertemplate: "<b>%{y}</b><br>SLA: %{x:.0f}%<extra></extra>",
        name: "Service level",
      },
    ],
    layout: plotLayout("Service level by channel", {
      height: chartHeight(df),
      showlegend: false,
      shapes: [
        {
          type: "line",
          x0: SLA_THRESHOLD * 100,
          x1: SLA_THRESHOLD * 100,
          y0: 0,
          y1: 1,
          yref: "paper",
          line: { dash: "dash", color: NEUTRAL },
        },
      ],
      annotations: [
        {
          x: SLA_THRESHOLD * 100,
          y: 1,
          yref: "paper",
          text: `Target ${Math.round(SLA_THRESHOLD * 100)}%`,
          showarrow: false,
          yanchor: "bottom",
        },
      ],
      xaxis: valueXAxis("Service level (%)", { range: [0, 115] }),
      yaxis: channelYAxis(channels),
    }),
  };
}

function chartResponseTime(df) {
  const sorted = [...df].sort((a, b) => a.ResponseSeconds - b.ResponseSeconds);
  const channels = sorted.map((r) => r.Channel);
  return {
    data: [
      {
        type: "bar",
        orientation: "h",
        y: channels,
        x: sorted.map((r) => r.ResponseSeconds / 60),
        marker: { color: WARNING },
        text: sorted.map((r) => r.AverageResponseText),
        textposition: "outside",
        cliponaxis: false,
        textfont: { size: 11 },
        hovertemplate: "<b>%{y}</b><br>Avg: %{text}<extra></extra>",
        name: "Avg response",
      },
    ],
    layout: plotLayout("Average response time (minutes)", {
      height: chartHeight(df),
      showlegend: false,
      xaxis: valueXAxis("Minutes"),
      yaxis: channelYAxis(channels),
    }),
  };
}

function chartPendingBacklog(df) {
  const sorted = [...df].sort((a, b) => {
    const totalA = a.Backlog + a.Pending;
    const totalB = b.Backlog + b.Pending;
    return totalA - totalB;
  });
  const channels = sorted.map((r) => r.Channel);
  return {
    data: [
      {
        type: "bar",
        orientation: "h",
        name: "Pending",
        y: channels,
        x: sorted.map((r) => r.Pending),
        marker: { color: WARNING },
        hovertemplate: "<b>%{y}</b><br>Pending: %{x}<extra></extra>",
      },
      {
        type: "bar",
        orientation: "h",
        name: "Backlog",
        y: channels,
        x: sorted.map((r) => r.Backlog),
        marker: { color: DANGER },
        hovertemplate: "<b>%{y}</b><br>Backlog: %{x}<extra></extra>",
      },
    ],
    layout: plotLayout("Pending and backlog by channel", {
      height: chartHeight(df),
      barmode: "stack",
      xaxis: valueXAxis("Ticket count"),
      yaxis: channelYAxis(channels),
    }),
  };
}

function chartVolumePie(df) {
  const pie = df.filter((r) => r.Incoming > 0);
  const legendRows = Math.ceil(pie.length / 2);
  return {
    data: [
      {
        type: "pie",
        labels: pie.map((r) => r.Channel),
        values: pie.map((r) => r.Incoming),
        hole: 0.42,
        textinfo: "percent",
        textposition: "inside",
        insidetextorientation: "horizontal",
        textfont: { size: 11 },
        hovertemplate: "<b>%{label}</b><br>Tickets: %{value:,}<br>%{percent}<extra></extra>",
      },
    ],
    layout: plotLayout("Incoming volume by channel", {
      height: Math.max(440, legendRows * 22 + 200),
      showlegend: true,
      legend: {
        orientation: "v",
        x: 1.02,
        xanchor: "left",
        y: 0.5,
        yanchor: "middle",
        font: { size: 11 },
        traceorder: "normal",
      },
      margin: { t: 60, b: 40, l: 20, r: 160 },
    }),
  };
}

function showError(message) {
  const box = document.getElementById("error-box");
  box.textContent = message;
  box.classList.remove("hidden");
  document.getElementById("dashboard").classList.add("hidden");
}

function hideError() {
  document.getElementById("error-box").classList.add("hidden");
}

function renderDashboard(df) {
  hideError();
  const dashboard = document.getElementById("dashboard");
  dashboard.classList.remove("hidden");

  const reportDate = df.find((r) => r.Date)?.Date;
  document.getElementById("report-date").textContent = reportDate
    ? `📅 Report date: ${reportDate}`
  : "";

  const totalBacklog = df.reduce((sum, r) => sum + r.Backlog, 0);
  document.getElementById("kpi-backlog").textContent = totalBacklog.toLocaleString("en-US");

  renderChart("chart-incoming", chartIncomingVsClosed(df));
  renderChart("chart-sla", chartServiceLevel(df));
  renderChart("chart-response", chartResponseTime(df));
  renderChart("chart-pending", chartPendingBacklog(df));
  renderChart("chart-pie", chartVolumePie(df));
}

function handleFile(file) {
  document.getElementById("file-name").textContent = file ? file.name : "";
  document.getElementById("welcome").classList.toggle("hidden", Boolean(file));
  if (!file) {
    document.getElementById("dashboard").classList.add("hidden");
    hideError();
    return;
  }

  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: "array", cellDates: true });
      const df = loadChannelsReport(workbook);
      if (!df.length) {
        showError("The file does not contain valid channel data.");
        return;
      }
      renderDashboard(df);
    } catch (err) {
      console.error(err);
      showError(
        "❌ Could not read the file. Make sure it includes a 'Channels Report' or 'Data' sheet with the expected structure."
      );
    }
  };
  reader.readAsArrayBuffer(file);
}

document.getElementById("excel-input").addEventListener("change", (e) => {
  const file = e.target.files?.[0];
  handleFile(file || null);
});
