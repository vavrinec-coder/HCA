import "./taskpane.css";

const DEFAULT_BACKEND_URL =
  import.meta.env.VITE_BACKEND_URL || "https://hca-calc-engine.onrender.com";

const elements = {
  button: document.querySelector("#payroll-recalc"),
  clearLog: document.querySelector("#clear-log"),
  status: document.querySelector("#status"),
  configState: document.querySelector("#config-state"),
  dataSheet: document.querySelector("#data-sheet"),
  dataRange: document.querySelector("#data-range"),
  filterRule: document.querySelector("#filter-rule"),
  forecastPeriod: document.querySelector("#forecast-period"),
  calculationMonths: document.querySelector("#calculation-months"),
  includedRows: document.querySelector("#included-rows"),
  loadSpeed: document.querySelector("#load-speed"),
  backendResult: document.querySelector("#backend-result"),
  activityLog: document.querySelector("#activity-log"),
  backendUrl: document.querySelector("#backend-url"),
  backendState: document.querySelector("#backend-state"),
  cloudDot: document.querySelector("#cloud-dot"),
};

Office.onReady((info) => {
  elements.backendUrl.value =
    localStorage.getItem("xf1.backendUrl") || DEFAULT_BACKEND_URL;
  elements.backendState.textContent = safeHost(elements.backendUrl.value);
  addLog("Task pane ready.");

  if (info.host !== Office.HostType.Excel) {
    setStatus("Open this add-in inside Excel Desktop.", "error");
    elements.button.disabled = true;
    return;
  }

  elements.button.addEventListener("click", handlePayrollRecalc);
  elements.clearLog.addEventListener("click", () => {
    elements.activityLog.innerHTML = "";
    addLog("Activity log cleared.");
  });
  elements.backendUrl.addEventListener("change", () => {
    const value = elements.backendUrl.value.trim();
    localStorage.setItem("xf1.backendUrl", value);
    elements.backendState.textContent = value ? safeHost(value) : "Backend";
    checkBackendHealth();
  });

  checkBackendHealth();
});

async function handlePayrollRecalc() {
  const startedAt = performance.now();
  setBusy(true);
  setStatus("Reading Payroll data from workbook...");
  addLog("Payroll Recalc started.");

  try {
    const payload = await buildPayrollPayload(startedAt);
    updateConfigUi(payload.source, payload.model);
    updateMetrics(payload.metrics);
    setStatus("Workbook load complete. Sending preview to backend...");

    const backendSummary = await sendLoadPreview(payload);
    setBackendUi("Success", "connected");
    elements.backendResult.textContent = backendSummary.status || "Success";
    elements.backendResult.className = "is-success";
    setStatus("Payroll load preview complete.", "success");
    addLog(
      `Payroll preview completed. ${payload.metrics.includedRows.toLocaleString()} rows included.`
    );
  } catch (error) {
    setBackendUi("Error", "error");
    elements.backendResult.textContent = "Error";
    elements.backendResult.className = "is-error";
    setStatus(error.message || String(error), "error");
    addLog(error.message || String(error));
  } finally {
    setBusy(false);
  }
}

async function buildPayrollPayload(startedAt) {
  return Excel.run(async (context) => {
    const configSheet = context.workbook.worksheets.getItem("Config");
    const configRange = configSheet.getUsedRange(true);
    configRange.load("values");

    await context.sync();

    const config = parseConfig(configRange.values);
    const dataSheet = context.workbook.worksheets.getItem(config.payroll.dataLoadSheet);
    const headerRange = dataSheet.getRange(config.payroll.headers);
    const dataRange = dataSheet.getRange(config.payroll.cellRange);

    headerRange.load("values");
    dataRange.load("values");

    await context.sync();

    const headers = normalizeHeaders(headerRange.values[0] || []);
    const rows = dataRange.values || [];
    const filterOffset = getFilterOffset(
      config.payroll.cellRange,
      config.payroll.filterColumn
    );
    const included = rows
      .filter((row) => isIncluded(row[filterOffset]))
      .map((row) => rowToObject(headers, row));

    const loadTimeMs = Math.round((performance.now() - startedAt) * 10) / 10;

    return {
      section: "Payroll",
      model: config.model,
      source: {
        sheet: config.payroll.dataLoadSheet,
        headerRange: config.payroll.headers,
        dataRange: config.payroll.cellRange,
        filterColumn: config.payroll.filterColumn,
      },
      metrics: {
        totalRows: rows.length,
        includedRows: included.length,
        loadTimeMs,
      },
      headers,
      rows: included,
    };
  });
}

function parseConfig(values) {
  if (!values || values.length < 2) {
    throw new Error("Config sheet does not contain the expected config table.");
  }

  const headerRow = values[0].map((value) => normalizeKey(value));
  const columns = {
    section: headerRow.indexOf("section"),
    setting: headerRow.indexOf("setting"),
    value: headerRow.indexOf("value"),
  };

  for (const [name, index] of Object.entries(columns)) {
    if (index === -1) {
      throw new Error(`Config sheet is missing required column: ${name}`);
    }
  }

  const settings = {};
  values.slice(1).forEach((row) => {
    const section = normalizeKey(row[columns.section]);
    const setting = normalizeKey(row[columns.setting]);
    if (!section || !setting) {
      return;
    }
    settings[`${section}.${setting}`] = row[columns.value];
  });

  const lastActualsDate = parseExcelDate(
    requiredSetting(settings, "model.last actuals date"),
    "Last actuals date"
  );
  const modelEndDate = parseExcelDate(
    requiredSetting(settings, "model.model end date"),
    "Model end date"
  );
  const financialYearEndMonth = parseMonthNumber(
    requiredSetting(settings, "model.financial year end month"),
    "Financial year end month"
  );
  const timeline = buildModelTimeline(
    lastActualsDate,
    modelEndDate,
    financialYearEndMonth
  );

  const payroll = {
    dataLoadSheet: requiredSetting(settings, "payroll.data load sheet"),
    cellRange: requiredSetting(settings, "payroll.cell range"),
    headers: requiredSetting(settings, "payroll.headers"),
    filterColumn: requiredSetting(settings, "payroll.filter column"),
  };

  return {
    model: {
      lastActualsDate: formatIsoDate(lastActualsDate),
      modelEndDate: formatIsoDate(modelEndDate),
      calculationStartDate: timeline.calculationStartDate,
      calculationEndDate: timeline.calculationEndDate,
      calculationMonths: timeline.periods.length,
      financialYearEndMonth,
      periods: timeline.periods,
    },
    payroll,
  };
}

function normalizeHeaders(headers) {
  return headers.map((header, index) => {
    const normalized = String(header ?? "").trim();
    return normalized || `Column ${index + 1}`;
  });
}

function rowToObject(headers, row) {
  return headers.reduce((record, header, index) => {
    record[header] = row[index] ?? null;
    return record;
  }, {});
}

function isIncluded(value) {
  if (typeof value === "number") {
    return value === 1;
  }

  return String(value ?? "").trim() === "1";
}

function getFilterOffset(rangeAddress, filterColumn) {
  const startColumn = extractStartColumn(rangeAddress);
  const startIndex = columnToNumber(startColumn);
  const filterIndex = columnToNumber(filterColumn);
  const offset = filterIndex - startIndex;

  if (offset < 0) {
    throw new Error(
      `Filter column ${filterColumn} is outside data range ${rangeAddress}.`
    );
  }

  return offset;
}

function extractStartColumn(rangeAddress) {
  const cleaned = String(rangeAddress).split("!").pop().replace(/\$/g, "");
  const match = cleaned.match(/^([A-Z]+)\d+/i);

  if (!match) {
    throw new Error(`Could not read start column from range: ${rangeAddress}`);
  }

  return match[1];
}

function columnToNumber(columnLetters) {
  return String(columnLetters)
    .trim()
    .toUpperCase()
    .split("")
    .reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0);
}

function requiredSetting(settings, key) {
  const value = settings[key];
  if (value === undefined || value === null || String(value).trim() === "") {
    throw new Error(`Config value is blank or missing: ${key}`);
  }
  return value;
}

function normalizeKey(value) {
  return String(value ?? "").trim().toLowerCase();
}

function parseExcelDate(value, label) {
  if (typeof value === "number") {
    return new Date(Date.UTC(1899, 11, 30 + value));
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()));
  }

  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) {
    throw new Error(`Config date is invalid: ${label}`);
  }

  return new Date(
    Date.UTC(parsed.getFullYear(), parsed.getMonth(), parsed.getDate())
  );
}

function parseMonthNumber(value, label) {
  const month = Number(value);
  if (!Number.isInteger(month) || month < 1 || month > 12) {
    throw new Error(`${label} must be a number from 1 to 12.`);
  }
  return month;
}

function buildModelTimeline(lastActualsDate, modelEndDate, financialYearEndMonth) {
  const startDate = endOfMonth(addMonths(lastActualsDate, 1));
  const endDate = endOfMonth(modelEndDate);

  if (startDate > endDate) {
    throw new Error("Model end date must be after Last actuals date.");
  }

  const periods = [];
  for (
    let cursor = startDate;
    cursor <= endDate;
    cursor = endOfMonth(addMonths(cursor, 1))
  ) {
    periods.push({
      date: formatIsoDate(cursor),
      label: formatMonthLabel(cursor),
      financialYear: getFinancialYear(cursor, financialYearEndMonth),
    });
  }

  return {
    calculationStartDate: formatIsoDate(startDate),
    calculationEndDate: formatIsoDate(endDate),
    periods,
  };
}

function addMonths(date, months) {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + months, 1));
}

function endOfMonth(date) {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + 1, 0));
}

function getFinancialYear(date, financialYearEndMonth) {
  const month = date.getUTCMonth() + 1;
  const year = date.getUTCFullYear();
  return month <= financialYearEndMonth ? year : year + 1;
}

function formatIsoDate(date) {
  return date.toISOString().slice(0, 10);
}

function formatMonthLabel(date) {
  return date.toLocaleString("en-US", {
    month: "short",
    year: "numeric",
    timeZone: "UTC",
  });
}

async function sendLoadPreview(payload) {
  const baseUrl = elements.backendUrl.value.trim().replace(/\/$/, "");
  localStorage.setItem("xf1.backendUrl", baseUrl);
  elements.backendState.textContent = safeHost(baseUrl);

  const response = await fetch(`${baseUrl}/payroll/load-preview`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Backend returned ${response.status}: ${body}`);
  }

  return response.json();
}

function updateMetrics(metrics) {
  const rowsPerSecond =
    metrics.loadTimeMs > 0 ? Math.round(metrics.totalRows / (metrics.loadTimeMs / 1000)) : 0;
  const percent =
    metrics.totalRows > 0
      ? ((metrics.includedRows / metrics.totalRows) * 100).toFixed(2)
      : "0.00";

  elements.includedRows.textContent = `${metrics.includedRows.toLocaleString()} of ${metrics.totalRows.toLocaleString()} (${percent}%)`;
  elements.loadSpeed.textContent = `${rowsPerSecond.toLocaleString()} rows/sec - ${(
    metrics.loadTimeMs / 1000
  ).toFixed(2)}s`;
}

function updateConfigUi(source, model) {
  elements.configState.textContent = "Config";
  elements.configState.className = "is-success";
  elements.dataSheet.textContent = source.sheet;
  elements.dataRange.textContent = source.dataRange;
  elements.filterRule.textContent = `${source.filterColumn} = 1`;
  elements.forecastPeriod.textContent = `${model.periods[0].label} to ${
    model.periods[model.periods.length - 1].label
  }`;
  elements.calculationMonths.textContent = `${model.calculationMonths} months`;
  addLog("Config loaded successfully.");
}

function setBusy(isBusy) {
  elements.button.disabled = isBusy;
  elements.button.textContent = isBusy ? "Loading..." : "Payroll Recalc";
}

function setStatus(message, type = "") {
  elements.status.textContent = message;
  elements.status.className = `status${type ? ` is-${type}` : ""}`;
}

async function checkBackendHealth() {
  const baseUrl = elements.backendUrl.value.trim().replace(/\/$/, "");
  if (!baseUrl) {
    setBackendUi("Missing", "error");
    return;
  }

  try {
    const response = await fetch(`${baseUrl}/health`);
    if (!response.ok) {
      throw new Error(`Health check returned ${response.status}`);
    }
    setBackendUi("Connected", "connected");
    elements.backendResult.textContent = "Ready";
    elements.backendResult.className = "is-success";
    addLog("Connection to backend verified.");
  } catch {
    setBackendUi("Offline", "error");
    elements.backendResult.textContent = "Offline";
    elements.backendResult.className = "is-error";
  }
}

function setBackendUi(label, state) {
  elements.backendState.textContent = label;
  elements.backendState.className = `state-pill is-${state}`;
  elements.cloudDot.className = `cloud-dot is-${state}`;
}

function addLog(message) {
  const item = document.createElement("li");
  const time = document.createElement("time");
  const text = document.createElement("span");
  time.textContent = new Date().toLocaleTimeString([], {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
  text.textContent = message;
  item.append(time, text);
  elements.activityLog.append(item);

  while (elements.activityLog.children.length > 8) {
    elements.activityLog.firstElementChild.remove();
  }
}

function safeHost(value) {
  try {
    return new URL(value).host;
  } catch {
    return "Backend";
  }
}
