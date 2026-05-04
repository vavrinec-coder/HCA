import "./taskpane.css";
import { getConfigNamedRange, getFilterOffset, parseConfig } from "./config.js";

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
    await clearOutputRange(payload.output);
    setStatus("Workbook load complete. Sending preview to backend...");

    const backendSummary = await sendLoadPreview(payload);
    await writePayrollOutputs(payload.output, backendSummary.outputs);
    setBackendUi("Success", "connected");
    elements.backendResult.textContent = backendSummary.status || "Success";
    elements.backendResult.className = "is-success";
    setStatus("Payroll headcount output complete.", "success");
    addLog(
      `Payroll headcount completed. ${payload.metrics.includedRows.toLocaleString()} rows included.`
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
    const configRange = getConfigNamedRange(context);
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
      output: config.output,
      assumptions: config.assumptions,
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

async function clearOutputRange(outputConfig) {
  const cleared = await Excel.run(async (context) => {
    const outputSheet = context.workbook.worksheets.getItem(outputConfig.sheet);
    const usedRange = outputSheet.getUsedRangeOrNullObject();
    usedRange.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);

    await context.sync();

    if (usedRange.isNullObject) {
      addLog(`No existing output found on ${outputConfig.sheet}.`);
      return false;
    }

    const firstRowToClear = Math.max(usedRange.rowIndex, 1);
    const firstColumnToClear = Math.max(usedRange.columnIndex, 1);
    const lastRow = usedRange.rowIndex + usedRange.rowCount - 1;
    const lastColumn = usedRange.columnIndex + usedRange.columnCount - 1;
    const rowCount = lastRow - firstRowToClear + 1;
    const columnCount = lastColumn - firstColumnToClear + 1;

    if (rowCount <= 0 || columnCount <= 0) {
      addLog(`No clearable output found on ${outputConfig.sheet}.`);
      return false;
    }

    outputSheet
      .getRangeByIndexes(firstRowToClear, firstColumnToClear, rowCount, columnCount)
      .clear(Excel.ClearApplyTo.all);

    await context.sync();
    return true;
  });

  if (cleared) {
    addLog(`Cleared stale output from ${outputConfig.sheet}.`);
  }
}

async function writePayrollOutputs(outputConfig, outputs) {
  if (!outputs?.headcount?.table?.length) {
    throw new Error("Backend did not return a headcount output table.");
  }
  if (!outputs?.baseSalary?.total?.table?.length) {
    throw new Error("Backend did not return base salary output tables.");
  }

  await Excel.run(async (context) => {
    const outputSheet = context.workbook.worksheets.getItem(outputConfig.sheet);
    writeOutputTable(
      outputSheet,
      outputConfig.headcountStartCell,
      outputs.headcount.table,
      "0.00"
    );
    writeOutputTable(
      outputSheet,
      outputConfig.baseSalaryTotalStartCell,
      outputs.baseSalary.total.table,
      "#,##0"
    );
    writeOutputTable(
      outputSheet,
      outputConfig.baseSalaryDomesticStartCell,
      outputs.baseSalary.domestic.table,
      "#,##0"
    );
    writeOutputTable(
      outputSheet,
      outputConfig.baseSalaryInternationalStartCell,
      outputs.baseSalary.international.table,
      "#,##0"
    );
    writeOutputTable(
      outputSheet,
      outputConfig.baseSalaryCogsStartCell,
      outputs.baseSalary.cogs.table,
      "#,##0"
    );
    writeOutputTable(
      outputSheet,
      outputConfig.medicalStartCell,
      outputs.benefits.medical.table,
      "#,##0"
    );
    writeOutputTable(
      outputSheet,
      outputConfig.retirement401kStartCell,
      outputs.benefits.retirement401k.table,
      "#,##0"
    );
    writeOutputTable(
      outputSheet,
      outputConfig.otherBenefitsStartCell,
      outputs.benefits.otherBenefits.table,
      "#,##0"
    );

    await context.sync();
  });

  addLog(
    `Payroll outputs written to ${outputConfig.sheet}.`
  );
}

function writeOutputTable(outputSheet, startCell, table, numberFormat) {
  const startRange = outputSheet.getRange(startCell);
  const targetRange = startRange.getResizedRange(
    table.length - 1,
    table[0].length - 1
  );

  targetRange.values = table;
  targetRange.numberFormat = buildOutputNumberFormat(table, numberFormat);
  targetRange.format.autofitColumns();
}

function buildOutputNumberFormat(table, numberFormat) {
  return table.map((row, rowIndex) =>
    row.map((_, columnIndex) => {
      if (rowIndex === 0 || columnIndex === 0) {
        return "@";
      }
      return numberFormat;
    })
  );
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
