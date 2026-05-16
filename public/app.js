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

function plotLayout(title, extra = {}) {
  return {
    title: { text: title, x: 0.98, xanchor: "right", font: { size: 16 } },
    font: { family: "Tahoma, Segoe UI, sans-serif", size: 12 },
    margin: { t: 60, b: 40, l: 20, r: 20 },
    plot_bgcolor: "#ffffff",
    paper_bgcolor: "#ffffff",
    showlegend: true,
    legend: { orientation: "h", yanchor: "bottom", y: 1.02, xanchor: "right", x: 1 },
    ...extra,
  };
}

function loadChannelsReport(workbook) {
  const sheet = workbook.Sheets["Channels Report"];
  if (!sheet) throw new Error("ورقة 'Channels Report' غير موجودة");

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: null,
    raw: false,
  });

  const headerRow = rows[2] || [];
  const dataRows = rows.slice(3);

  const records = dataRows
    .map((row) => {
      const slice = row.slice(COL_START, COL_END + 1);
      const obj = {};
      headerRow.slice(COL_START, COL_END + 1).forEach((key, idx) => {
        const name = key == null ? "" : String(key).trim();
        obj[name || `col_${idx}`] = slice[idx];
      });
      return obj;
    })
    .filter((row) => row.Channel != null && String(row.Channel).trim() !== "");

  let lastDate = null;
  for (const row of records) {
    if (row.Date != null && row.Date !== "") lastDate = row.Date;
    else row.Date = lastDate;
    row.Date = excelDateToString(row.Date);
    row.Channel = String(row.Channel).trim();

    const extraKey = Object.keys(row).find((k) => k.toLowerCase().startsWith("unnamed"));
    if (extraKey) {
      row.Pending = toNumber(row.Pending) + toNumber(row[extraKey]);
      delete row[extraKey];
    }

    ["Incoming", "Closed", "Pending", "Backlog"].forEach((col) => {
      row[col] = Math.round(toNumber(row[col]));
    });

    row["Service Level"] = toNumber(row["Service Level"]);
    row.ResponseSeconds = secondsFromTime(row["Average Response"]);
    row.AverageResponseText = formatSeconds(row.ResponseSeconds);
    row.ClosureRate =
      row.Incoming > 0 ? row.Closed / row.Incoming : 0;
  }

  return records;
}

function renderChart(id, figure) {
  Plotly.newPlot(id, figure.data, figure.layout, { responsive: true, displayModeBar: false });
}

function chartIncomingVsClosed(df) {
  const sorted = [...df].sort((a, b) => b.Incoming - a.Incoming);
  return {
    data: [
      {
        type: "bar",
        name: "الوارد",
        x: sorted.map((r) => r.Channel),
        y: sorted.map((r) => r.Incoming),
        marker: { color: PRIMARY },
        text: sorted.map((r) => r.Incoming),
        textposition: "outside",
      },
      {
        type: "bar",
        name: "المُغلق",
        x: sorted.map((r) => r.Channel),
        y: sorted.map((r) => r.Closed),
        marker: { color: SUCCESS },
        text: sorted.map((r) => r.Closed),
        textposition: "outside",
      },
    ],
    layout: plotLayout("الوارد مقابل المُغلق لكل قناة", {
      barmode: "group",
      xaxis: { tickangle: -35 },
      yaxis: { title: "عدد التذاكر", gridcolor: "#eee" },
    }),
  };
}

function chartServiceLevel(df) {
  const sorted = [...df].sort((a, b) => a["Service Level"] - b["Service Level"]);
  return {
    data: [
      {
        type: "bar",
        orientation: "h",
        x: sorted.map((r) => r["Service Level"] * 100),
        y: sorted.map((r) => r.Channel),
        marker: {
          color: sorted.map((r) =>
            r["Service Level"] >= SLA_THRESHOLD ? SUCCESS : DANGER
          ),
        },
        text: sorted.map((r) => `${Math.round(r["Service Level"] * 100)}%`),
        textposition: "outside",
        name: "مستوى الخدمة",
      },
    ],
    layout: plotLayout("مستوى الخدمة لكل قناة", {
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
          text: `الهدف ${Math.round(SLA_THRESHOLD * 100)}%`,
          showarrow: false,
          yanchor: "bottom",
        },
      ],
      xaxis: { title: "مستوى الخدمة (٪)", range: [0, 110], gridcolor: "#eee" },
      yaxis: { title: "" },
    }),
  };
}

function chartResponseTime(df) {
  const sorted = [...df].sort((a, b) => b.ResponseSeconds - a.ResponseSeconds);
  return {
    data: [
      {
        type: "bar",
        x: sorted.map((r) => r.Channel),
        y: sorted.map((r) => r.ResponseSeconds / 60),
        marker: { color: WARNING },
        text: sorted.map((r) => r.AverageResponseText),
        textposition: "outside",
        name: "متوسط الاستجابة",
      },
    ],
    layout: plotLayout("متوسط وقت الاستجابة (دقيقة)", {
      showlegend: false,
      xaxis: { tickangle: -35 },
      yaxis: { title: "الدقائق", gridcolor: "#eee" },
    }),
  };
}

function chartPendingBacklog(df) {
  const sorted = [...df].sort((a, b) => {
    if (b.Backlog !== a.Backlog) return b.Backlog - a.Backlog;
    return b.Pending - a.Pending;
  });
  return {
    data: [
      {
        type: "bar",
        name: "المُعلّق",
        x: sorted.map((r) => r.Channel),
        y: sorted.map((r) => r.Pending),
        marker: { color: WARNING },
      },
      {
        type: "bar",
        name: "المتراكم",
        x: sorted.map((r) => r.Channel),
        y: sorted.map((r) => r.Backlog),
        marker: { color: DANGER },
      },
    ],
    layout: plotLayout("المُعلّق والمتراكم لكل قناة", {
      barmode: "stack",
      xaxis: { tickangle: -35 },
      yaxis: { title: "عدد التذاكر", gridcolor: "#eee" },
    }),
  };
}

function chartVolumePie(df) {
  const pie = df.filter((r) => r.Incoming > 0);
  return {
    data: [
      {
        type: "pie",
        labels: pie.map((r) => r.Channel),
        values: pie.map((r) => r.Incoming),
        hole: 0.45,
        textposition: "inside",
        textinfo: "percent+label",
      },
    ],
    layout: plotLayout("توزيع الوارد على القنوات"),
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
    ? `📅 تاريخ التقرير: ${reportDate}`
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
        showError("الملف لا يحتوي على بيانات قنوات صالحة.");
        return;
      }
      renderDashboard(df);
    } catch (err) {
      console.error(err);
      showError(
        "❌ تعذّر قراءة الملف. تأكد أن الملف يحتوي على ورقة باسم 'Channels Report' بنفس الهيكل المتوقع."
      );
    }
  };
  reader.readAsArrayBuffer(file);
}

document.getElementById("excel-input").addEventListener("change", (e) => {
  const file = e.target.files?.[0];
  handleFile(file || null);
});
