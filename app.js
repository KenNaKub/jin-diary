const STORAGE_KEY = "jin-diary-entries";
const DELETED_STORAGE_KEY = "jin-diary-deleted-entry-ids";
const SETTINGS_KEY = "jin-diary-settings";
const DEFAULT_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbywIWTWnpCDGlCDRJ0aNNnWUw-rz72KpgXL9lwdWYWPem-sRbE0_bu0XfV2NaiCvqPi/exec";
const APP_TIME_ZONE = "Asia/Bangkok";
const BANGKOK_OFFSET = "+07:00";
const DEFAULT_ACTIVITY = "feed";
const DAILY_SUMMARY_LIMIT = 10;
const CHART_DAYS = 30;

const scriptTemplate = `const SHEET_NAME = "entries";

function doGet(e) {
  const action = e.parameter.action || "list";
  const callback = e.parameter.callback;
  const output = handleAction(action, e.parameter);
  return respond(output, callback);
}

function doPost(e) {
  const params = e.parameter || {};
  const body = e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
  const action = params.action || body.action || "append";
  const output = handleAction(action, { ...params, ...body });
  return respond(output, params.callback);
}

function handleAction(action, params) {
  if (action === "append") return appendEntry(params);
  if (action === "delete" || action === "remove" || action === "delete_entry") return deleteEntry(params);
  return listEntries();
}

function appendEntry(params) {
  const sheet = getSheet();
  const entry = typeof params.entry === "string" ? JSON.parse(params.entry || "{}") : params.entry || params;
  const row = [
    entry.id,
    entry.happenedAt,
    entry.type,
    entry.amount,
    entry.unit,
    entry.note,
    entry.createdAt || bangkokTimestamp()
  ];
  const existingRow = findRowById(sheet, entry.id);
  if (existingRow) {
    sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return { ok: true, entry };
}

function bangkokTimestamp() {
  return Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd'T'HH:mm:ss") + "+07:00";
}

function listEntries() {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1).filter(row => row[0]);
  const seen = {};
  return {
    ok: true,
    entries: rows
      .filter(row => {
        const id = String(row[0]);
        if (seen[id]) return false;
        seen[id] = true;
        return true;
      })
      .map(row => ({
        id: row[0],
        happenedAt: row[1],
        type: row[2],
        amount: row[3],
        unit: row[4],
        note: row[5]
      }))
  };
}

function deleteEntry(params) {
  const sheet = getSheet();
  const id = String(params.id || params.entryId || "");
  const values = sheet.getDataRange().getValues();
  let deleted = false;
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][0]) === id) {
      sheet.deleteRow(i + 1);
      deleted = true;
    }
  }
  if (deleted) return { ok: true, id };
  return { ok: false, id, error: "Entry not found" };
}

function findRowById(sheet, id) {
  if (!id) return 0;
  const values = sheet.getDataRange().getValues();
  const target = String(id);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === target) return i + 1;
  }
  return 0;
}

function getSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["id", "happenedAt", "type", "amount", "unit", "note", "createdAt"]);
  }
  return sheet;
}

function respond(payload, callback) {
  const body = callback ? callback + "(" + JSON.stringify(payload) + ")" : JSON.stringify(payload);
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JAVASCRIPT);
}`;

let deletedEntryIds = loadDeletedEntryIds();
let entries = loadEntries();
let settings = loadSettings();
let chartType = "feed";
const sessionPendingEntryIds = new Set();

const els = {
  tabs: document.querySelectorAll("[data-view-button]"),
  views: document.querySelectorAll("[data-view]"),
  form: document.querySelector("#entryForm"),
  type: document.querySelector("#activityType"),
  happenedAt: document.querySelector("#happenedAt"),
  amount: document.querySelector("#amount"),
  unit: document.querySelector("#unit"),
  note: document.querySelector("#note"),
  minusTenButton: document.querySelector("#minusTenButton"),
  nowButton: document.querySelector("#nowButton"),
  sampleButton: document.querySelector("#sampleButton"),
  syncMainButton: document.querySelector("#syncMainButton"),
  clearLocalButton: document.querySelector("#clearLocalButton"),
  refreshButton: document.querySelector("#refreshButton"),
  todayTimeline: document.querySelector("#todayTimeline"),
  todayLabel: document.querySelector("#todayLabel"),
  dayGrid: document.querySelector("#dayGrid"),
  summaryChart: document.querySelector("#summaryChart"),
  chartToggles: document.querySelectorAll("[data-chart-type]"),
  syncStatuses: document.querySelectorAll(".sync-status"),
  scriptUrl: document.querySelector("#scriptUrl"),
  saveSettingsButton: document.querySelector("#saveSettingsButton"),
  syncButton: document.querySelector("#syncButton"),
  scriptSnippet: document.querySelector("#scriptSnippet"),
  timelineTemplate: document.querySelector("#timelineItemTemplate"),
};

function isLocalTestHost(hostname = window.location.hostname) {
  return hostname === "localhost" || hostname === "127.0.0.1" || hostname === "";
}

function configureLocalOnlyUi() {
  if (isLocalTestHost()) {
    document.documentElement.classList.add("local-test");
    document.querySelector('[data-view-button="settings"]')?.removeAttribute("hidden");
    document.querySelector('[data-view="settings"]')?.removeAttribute("hidden");
    return;
  }

  document.querySelector('[data-view-button="settings"]')?.setAttribute("hidden", "");
  document.querySelector('[data-view="settings"]')?.setAttribute("hidden", "");
  els.sampleButton?.setAttribute("hidden", "");
}

function localDateTimeValue(date = new Date()) {
  return bangkokDateTimeParts(date).inputValue;
}

function localDateTimeFromInput(value) {
  return value ? new Date(parseLocalInput(value)) : new Date();
}

function parseLocalInput(value) {
  return `${value}:00${BANGKOK_OFFSET}`;
}

function currentBangkokTimestamp() {
  return `${localDateTimeValue()}:00${BANGKOK_OFFSET}`;
}

function bangkokDateTimeParts(date = new Date()) {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: APP_TIME_ZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
    hourCycle: "h23",
  }).formatToParts(date);
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return {
    date: `${values.year}-${values.month}-${values.day}`,
    inputValue: `${values.year}-${values.month}-${values.day}T${values.hour}:${values.minute}`,
  };
}

function formatTime(value) {
  return new Intl.DateTimeFormat("en-US", {
    timeZone: APP_TIME_ZONE,
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).format(new Date(value));
}

function formatDay(value) {
  return new Intl.DateTimeFormat("en-US", {
    timeZone: APP_TIME_ZONE,
    weekday: "short",
    day: "numeric",
    month: "short",
  }).format(new Date(value));
}

function dayKey(value) {
  return bangkokDateTimeParts(new Date(value)).date;
}

function activityLabel(type) {
  const labels = {
    breast: "Breast",
    pump: "Pump",
    feed: "Feed",
    weight: "Weight",
    height: "Height",
    diaper: "Diaper",
    sleep: "Sleep",
    note: "Note",
  };
  return labels[type] || type;
}

function amountLabel(entry) {
  if (entry.amount === "" || entry.amount == null) return "";
  return `${Number(entry.amount).toLocaleString("en-US", { maximumFractionDigits: 3 })}${entry.unit ? ` ${entry.unit}` : ""}`;
}

function sortedEntries() {
  return dedupeEntries(entries).sort((a, b) => new Date(b.happenedAt) - new Date(a.happenedAt));
}

function render() {
  renderStatus();
  renderToday();
  renderSummary();
}

function renderStatus(message) {
  const hasUrl = Boolean(settings.scriptUrl);
  const invalidUrl = hasUrl && isGoogleusercontentUrl(settings.scriptUrl);
  const text = message || (invalidUrl ? "Need /exec URL" : hasUrl ? "Sheet ready" : "Local only");
  els.syncStatuses.forEach((status) => {
    status.textContent = text;
    status.classList.toggle("online", text === "Synced");
    status.classList.toggle("pending", hasUrl && !invalidUrl && text !== "Synced" && !text.toLowerCase().includes("failed"));
    status.classList.toggle("error", text.toLowerCase().includes("failed") || invalidUrl);
  });
}

function renderToday() {
  const today = dayKey(new Date());
  const todaysEntries = sortedEntries().filter((entry) => dayKey(entry.happenedAt) === today);
  els.todayTimeline.replaceChildren();
  els.todayLabel.textContent = todaysEntries.length ? `${todaysEntries.length} entries today` : "No entries yet";

  if (!todaysEntries.length) {
    els.todayTimeline.append(emptyState("No activity saved for today."));
    return;
  }

  todaysEntries.forEach((entry) => els.todayTimeline.append(renderTimelineItem(entry)));
}

function renderTimelineItem(entry) {
  const node = els.timelineTemplate.content.firstElementChild.cloneNode(true);
  node.dataset.type = entry.type;
  node.querySelector(".time").textContent = formatTime(entry.happenedAt);
  node.querySelector("strong").textContent = activityLabel(entry.type);
  node.querySelector("span").textContent = amountLabel(entry);
  node.querySelector("p").textContent = entry.note || "";
  node.querySelector("button").addEventListener("click", (event) => deleteEntryRecord(entry, event.currentTarget));
  return node;
}

async function deleteEntryRecord(entry, button) {
  if (!entry?.id) return false;
  if (button) button.disabled = true;

  const localOnly = !settings.scriptUrl || entry.pendingSync;
  const deletedFromSheet = localOnly ? true : await deleteFromSheet(entry.id).catch(() => false);

  if (!deletedFromSheet) {
    renderStatus("Sheet delete failed");
    if (button) button.disabled = false;
    return false;
  }

  rememberDeletedEntry(entry.id);
  entries = entries.filter((item) => item.id !== entry.id);
  saveEntries();
  render();
  renderStatus(localOnly ? "Deleted locally" : "Deleted from sheet");
  return true;
}

function renderSummary() {
  const groups = groupByDay(sortedEntries());
  const visibleGroups = groups.slice(0, DAILY_SUMMARY_LIMIT);
  els.dayGrid.replaceChildren();

  if (!groups.length) {
    renderSummaryChart([]);
    els.dayGrid.append(emptyState("Daily cards appear after you save activities."));
    return;
  }

  renderSummaryChart(groups);

  visibleGroups.forEach(([key, dayEntries]) => {
    const metrics = calculateDay(dayEntries);
    const card = document.createElement("article");
    card.className = "day-card";
    card.innerHTML = `
      <h3>${formatDay(dayEntries[0].happenedAt)}</h3>
      <div class="metrics">
        <div class="metric"><span>Feed</span><strong>${metrics.feed.count}</strong><small>${metrics.feed.amount}</small></div>
        <div class="metric"><span>Pump</span><strong>${metrics.pump.count}</strong><small>${metrics.pump.amount}</small></div>
        <div class="metric"><span>Breast</span><strong>${metrics.breast.count}</strong><small>${metrics.breast.amount}</small></div>
        <div class="metric"><span>Weight</span><strong>${metrics.weight}</strong></div>
        <div class="metric"><span>Height</span><strong>${metrics.height}</strong></div>
      </div>
      <details class="day-details">
        <summary>Details</summary>
        <div class="compact-list"></div>
      </details>
    `;

    const list = card.querySelector(".compact-list");
    dayEntries
      .slice()
      .sort((a, b) => new Date(a.happenedAt) - new Date(b.happenedAt))
      .forEach((entry) => {
        const line = document.createElement("div");
        line.className = "compact-list-item";
        const amount = amountLabel(entry);
        const text = document.createElement("span");
        text.textContent = `${formatTime(entry.happenedAt)} ${activityLabel(entry.type)}${amount ? ` (${amount})` : ""}${entry.note ? ` - ${entry.note}` : ""}`;
        const deleteButton = document.createElement("button");
        deleteButton.className = "icon danger compact-delete";
        deleteButton.type = "button";
        deleteButton.setAttribute("aria-label", `Delete ${activityLabel(entry.type)} at ${formatTime(entry.happenedAt)}`);
        deleteButton.textContent = "x";
        deleteButton.addEventListener("click", () => deleteEntryRecord(entry, deleteButton));
        line.append(text, deleteButton);
        list.append(line);
      });

    els.dayGrid.append(card);
  });
}

function renderSummaryChart(groups) {
  if (!els.summaryChart) return;
  els.chartToggles.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.chartType === chartType);
  });

  els.summaryChart.replaceChildren();
  els.summaryChart.classList.toggle("line-chart", isMeasurementChart(chartType));
  els.summaryChart.classList.toggle("height-chart", chartType === "height");
  if (!groups.length) {
    els.summaryChart.append(emptyState("Chart appears after you save activities."));
    return;
  }

  if (isMeasurementChart(chartType)) {
    renderMeasurementChart(groups);
    return;
  }

  renderBarChart(groups);
}

function renderBarChart(groups) {
  const groupsByKey = new Map(groups);
  const rows = lookbackDayKeys(CHART_DAYS).map((key) => {
    const dayEntries = groupsByKey.get(key) || [];
    const total = sumEntries(dayEntries, chartType);
    const unit = unitFor(dayEntries, chartType, defaultUnitFor(chartType));
    return {
      key,
      label: chartDayLabel(key),
      total,
      unit,
    };
  });
  const max = Math.max(...rows.map((row) => row.total), 1);
  const { width, height, pad } = chartFrame();
  const plotWidth = width - pad.left - pad.right;
  const plotHeight = height - pad.top - pad.bottom;
  const gap = 3;
  const barWidth = Math.max((plotWidth - gap * (rows.length - 1)) / rows.length, 2);
  const yMax = max * 1.12;
  const yFor = (value) => pad.top + ((yMax - value) / yMax) * plotHeight;
  const valueUnit = rows.find((row) => row.unit)?.unit || "oz";

  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
  svg.setAttribute("role", "img");
  svg.setAttribute("aria-label", `${activityLabel(chartType)} chart, ${CHART_DAYS} days`);
  svg.innerHTML = `
    ${chartGridLines(yMax, yFor, pad, plotWidth, valueUnit)}
    <line class="chart-axis" x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${pad.top + plotHeight}"></line>
    <line class="chart-axis" x1="${pad.left}" y1="${pad.top + plotHeight}" x2="${pad.left + plotWidth}" y2="${pad.top + plotHeight}"></line>
    ${rows
      .map((row, index) => {
        const x = pad.left + index * (barWidth + gap);
        const y = yFor(row.total);
        const barHeight = pad.top + plotHeight - y;
        const displayHeight = row.total ? Math.max(barHeight, 2) : 0;
        return `
          <rect class="bar-chart-fill" x="${x.toFixed(2)}" y="${(pad.top + plotHeight - displayHeight).toFixed(2)}" width="${barWidth.toFixed(2)}" height="${displayHeight.toFixed(2)}" rx="2">
            <title>${row.label}: ${formatMetric(row.total, row.unit)}</title>
          </rect>
          ${
            index % 5 === 0
              ? `<text class="chart-x" x="${(x + barWidth / 2).toFixed(2)}" y="${height - 10}" text-anchor="middle">${row.label}</text>`
              : ""
          }
        `;
      })
      .join("")}
  `;
  els.summaryChart.append(svg);
}

function renderMeasurementChart(groups) {
  const groupsByKey = new Map(groups);
  const fallbackUnit = chartType === "height" ? "cm" : "kg";
  const points = lookbackDayKeys(CHART_DAYS)
    .map((key, index) => {
      const measurement = latestMeasurement(groupsByKey.get(key) || [], chartType);
      if (!measurement) return null;
      return {
        index,
        key,
        label: chartDayLabel(key),
        value: Number(measurement.amount),
        unit: measurement.unit || fallbackUnit,
      };
    })
    .filter(Boolean);

  if (!points.length) {
    els.summaryChart.append(emptyState(`No ${activityLabel(chartType).toLowerCase()} saved in the last ${CHART_DAYS} days.`));
    return;
  }

  const { width, height, pad } = chartFrame();
  const max = Math.max(...points.map((point) => point.value), chartType === "height" ? 1 : 0.2);
  const yMax = max * 1.12;
  const plotWidth = width - pad.left - pad.right;
  const plotHeight = height - pad.top - pad.bottom;
  const xFor = (index) => pad.left + (index / (CHART_DAYS - 1)) * plotWidth;
  const yFor = (value) => pad.top + ((yMax - value) / yMax) * plotHeight;
  const path = points.map((point, index) => `${index ? "L" : "M"} ${xFor(point.index).toFixed(2)} ${yFor(point.value).toFixed(2)}`).join(" ");
  const valueUnit = points.at(-1)?.unit || fallbackUnit;
  const labels = points.length === 1 ? [points[0]] : [points[0], points.at(-1)];

  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
  svg.setAttribute("role", "img");
  svg.setAttribute("aria-label", `${activityLabel(chartType)} chart, ${points.length} measurements in ${CHART_DAYS} days`);
  svg.innerHTML = `
    ${chartGridLines(yMax, yFor, pad, plotWidth, valueUnit)}
    <line class="chart-axis" x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${pad.top + plotHeight}"></line>
    <line class="chart-axis" x1="${pad.left}" y1="${pad.top + plotHeight}" x2="${pad.left + plotWidth}" y2="${pad.top + plotHeight}"></line>
    <path class="line-chart-path" d="${path}"></path>
    ${points
      .map(
        (point) => `
          <circle class="line-chart-point" cx="${xFor(point.index).toFixed(2)}" cy="${yFor(point.value).toFixed(2)}" r="4">
            <title>${point.label}: ${formatMetric(point.value, point.unit)}</title>
          </circle>
        `,
      )
      .join("")}
    ${labels
      .map(
        (point) => `
          <text class="chart-x" x="${xFor(point.index).toFixed(2)}" y="${height - 10}" text-anchor="middle">${point.label}</text>
        `,
      )
      .join("")}
  `;
  els.summaryChart.append(svg);
}

function chartGridLines(yMax, yFor, pad, plotWidth, valueUnit) {
  return [1, 0.75, 0.5, 0.25, 0]
    .map((fraction) => {
      const value = yMax * fraction;
      const y = yFor(value);
      const line =
        fraction > 0 && fraction < 1
          ? `<line class="chart-grid-line" x1="${pad.left}" y1="${y.toFixed(2)}" x2="${pad.left + plotWidth}" y2="${y.toFixed(2)}"></line>`
          : "";
      return `
        ${line}
        <text class="chart-y" x="${pad.left - 8}" y="${y + 4}" text-anchor="end">${formatChartNumber(value)} ${valueUnit}</text>
      `;
    })
    .join("");
}

function chartFrame() {
  const chartWidth = Math.round(els.summaryChart.getBoundingClientRect().width);
  return {
    width: Math.max(chartWidth || 600, 320),
    height: 180,
    pad: { top: 16, right: 14, bottom: 28, left: 58 },
  };
}

function lookbackDayKeys(days) {
  const today = dayKey(new Date());
  const cursor = new Date(`${today}T00:00:00${BANGKOK_OFFSET}`);
  const keys = [];
  for (let i = days - 1; i >= 0; i--) {
    const date = new Date(cursor);
    date.setDate(cursor.getDate() - i);
    keys.push(dayKey(date));
  }
  return keys;
}

function chartDayLabel(key) {
  return new Intl.DateTimeFormat("en-US", {
    timeZone: APP_TIME_ZONE,
    month: "short",
    day: "numeric",
  }).format(new Date(`${key}T00:00:00${BANGKOK_OFFSET}`));
}

function groupByDay(items) {
  const groups = new Map();
  items.forEach((entry) => {
    const key = dayKey(entry.happenedAt);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(entry);
  });
  return [...groups.entries()].sort(([a], [b]) => b.localeCompare(a));
}

function calculateDay(items) {
  const entriesFor = (type) => items.filter((entry) => entry.type === type);
  const weight = latestMeasurement(items, "weight");
  const height = latestMeasurement(items, "height");

  return {
    feed: countAndAmount(entriesFor("feed"), sumEntries(items, "feed"), unitFor(items, "feed", "oz")),
    pump: countAndAmount(entriesFor("pump"), sumEntries(items, "pump"), unitFor(items, "pump", "oz")),
    breast: countAndAmount(entriesFor("breast"), sumEntries(items, "breast"), unitFor(items, "breast", defaultUnitFor("breast"))),
    weight: weight ? amountLabel(weight) : "-",
    height: height ? amountLabel(height) : "-",
  };
}

function isMeasurementChart(type) {
  return type === "weight" || type === "height";
}

function latestMeasurement(items, type) {
  return items
    .filter((entry) => entry.type === type && entry.amount !== "" && entry.amount != null)
    .sort((a, b) => new Date(b.happenedAt) - new Date(a.happenedAt))[0];
}

function sumEntries(items, type) {
  return items
    .filter((entry) => entry.type === type)
    .reduce((total, entry) => total + Number(entry.amount || 0), 0);
}

function unitFor(items, type, fallback) {
  return items.find((entry) => entry.type === type && entry.unit)?.unit || fallback;
}

function countAndAmount(items, totalAmount, unit) {
  return {
    count: `${items.length} ${items.length === 1 ? "time" : "times"}`,
    amount: `Total ${formatMetric(totalAmount, unit)}`,
  };
}

function formatMetric(value, unit) {
  if (!value) return "-";
  return `${value.toLocaleString("en-US", { maximumFractionDigits: 2 })} ${unit}`;
}

function formatChartNumber(value) {
  return Number(value).toLocaleString("en-US", { maximumFractionDigits: 1 });
}

function emptyState(text) {
  const node = document.createElement("div");
  node.className = "empty-state";
  node.textContent = text;
  return node;
}

function createEntry() {
  const entry = {
    id: crypto.randomUUID(),
    happenedAt: parseLocalInput(els.happenedAt.value),
    type: els.type.value,
    amount: els.amount.value === "" ? "" : Number(els.amount.value),
    unit: els.unit.value,
    note: els.note.value.trim(),
    createdAt: currentBangkokTimestamp(),
    pendingSync: true,
  };
  sessionPendingEntryIds.add(entry.id);
  return entry;
}

function saveEntries() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(entries));
}

function loadEntries() {
  try {
    const loaded = JSON.parse(localStorage.getItem(STORAGE_KEY)) || [];
    return loaded.filter((entry) => !isDeletedEntry(entry));
  } catch {
    return [];
  }
}

function rememberDeletedEntry(id) {
  if (!id) return;
  sessionPendingEntryIds.delete(String(id));
  deletedEntryIds.add(String(id));
  saveDeletedEntryIds();
}

function forgetDeletedEntries(ids) {
  let changed = false;
  ids.forEach((id) => {
    if (deletedEntryIds.delete(String(id))) changed = true;
  });
  if (changed) saveDeletedEntryIds();
}

function isDeletedEntry(entry) {
  return entry?.id && deletedEntryIds.has(String(entry.id));
}

function saveDeletedEntryIds() {
  localStorage.setItem(DELETED_STORAGE_KEY, JSON.stringify([...deletedEntryIds]));
}

function loadDeletedEntryIds() {
  try {
    return new Set(JSON.parse(localStorage.getItem(DELETED_STORAGE_KEY)) || []);
  } catch {
    return new Set();
  }
}

function saveSettings() {
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
}

function loadSettings() {
  try {
    const loaded = JSON.parse(localStorage.getItem(SETTINGS_KEY)) || {};
    if (!loaded.scriptUrl || isGoogleusercontentUrl(loaded.scriptUrl) || loaded.scriptUrl !== DEFAULT_SCRIPT_URL) {
      return { ...loaded, scriptUrl: DEFAULT_SCRIPT_URL };
    }
    return loaded;
  } catch {
    return { scriptUrl: DEFAULT_SCRIPT_URL };
  }
}

function jsonp(action, params = {}) {
  if (!settings.scriptUrl) return Promise.resolve({ ok: false });
  if (isGoogleusercontentUrl(settings.scriptUrl)) {
    return Promise.reject(new Error("Use the script.google.com /exec URL, not the googleusercontent echo URL."));
  }

  return new Promise((resolve, reject) => {
    const callback = `jinDiaryCallback_${Date.now()}_${Math.floor(Math.random() * 100000)}`;
    const url = new URL(settings.scriptUrl);
    url.searchParams.set("action", action);
    url.searchParams.set("callback", callback);
    Object.entries(params).forEach(([key, value]) => url.searchParams.set(key, value));

    const script = document.createElement("script");
    const cleanup = () => {
      delete window[callback];
      script.remove();
    };

    window[callback] = (payload) => {
      cleanup();
      resolve(payload);
    };

    script.onerror = () => {
      cleanup();
      reject(new Error("Sheet request failed"));
    };

    script.src = url.toString();
    document.body.append(script);
  });
}

function isGoogleusercontentUrl(value) {
  try {
    return new URL(value).hostname === "script.googleusercontent.com";
  } catch {
    return false;
  }
}

async function appendToSheet(entry, options = {}) {
  if (!settings.scriptUrl) return false;
  if (options.showStatus !== false) renderStatus("Saving to sheet...");
  const payload = await jsonp("append", { entry: JSON.stringify(entry) });
  const saved = payload.ok && payload.entry?.id === entry.id;
  if (saved) markEntrySynced(entry.id);
  if (options.showStatus !== false) renderStatus(saved ? "Saved to sheet" : "Sheet save failed");
  return saved;
}

function markEntrySynced(id) {
  if (!id) return;
  sessionPendingEntryIds.delete(String(id));
  entries = entries.map((entry) => (entry.id === id ? { ...entry, pendingSync: false } : entry));
  saveEntries();
}

async function deleteFromSheet(id) {
  if (!settings.scriptUrl) return true;
  renderStatus("Deleting from sheet...");
  const payload = await jsonp("delete", { id });
  const notFound = payload.error === "Entry not found";
  const deleted = String(payload.id) === String(id) && (payload.ok || notFound);
  renderStatus(deleted ? (notFound ? "Already deleted from sheet" : "Deleted from sheet") : "Sheet delete failed");
  return deleted;
}

async function syncFromSheet() {
  if (!settings.scriptUrl) {
    renderStatus("Local only");
    return false;
  }

  renderStatus("Syncing...");
  const payload = await jsonp("list");
  if (payload.ok && Array.isArray(payload.entries)) {
    const rawSheetEntries = dedupeEntries(payload.entries.map(normalizeEntry));
    forgetDeletedEntries(rawSheetEntries.map((entry) => entry.id).filter(Boolean));
    const sheetEntries = rawSheetEntries.filter((entry) => !isDeletedEntry(entry));
    const sheetIds = new Set(sheetEntries.map((entry) => entry.id));
    const localOnlyEntries = entries.filter(
      (entry) => entry.pendingSync && entry.id && sessionPendingEntryIds.has(entry.id) && !sheetIds.has(entry.id)
    );

    if (localOnlyEntries.length) {
      renderStatus(`Uploading ${localOnlyEntries.length} local ${localOnlyEntries.length === 1 ? "entry" : "entries"}...`);
      for (const entry of localOnlyEntries) {
        const { pendingSync, ...entryForSheet } = entry;
        const saved = await appendToSheet(entryForSheet, { showStatus: false });
        if (!saved) {
          renderStatus("Sync failed");
          return false;
        }
      }

      const refreshed = await jsonp("list");
      if (!refreshed.ok || !Array.isArray(refreshed.entries)) {
        renderStatus("Sync failed");
        return false;
      }

      const refreshedEntries = dedupeEntries(refreshed.entries.map(normalizeEntry));
      forgetDeletedEntries(refreshedEntries.map((entry) => entry.id).filter(Boolean));
      entries = refreshedEntries.filter((entry) => !isDeletedEntry(entry));
    } else {
      entries = sheetEntries;
    }

    saveEntries();
    render();
    renderStatus("Synced");
    return true;
  } else {
    renderStatus("Sync failed");
    return false;
  }
}

function normalizeEntry(entry) {
  return {
    id: entry.id || crypto.randomUUID(),
    happenedAt: entry.happenedAt,
    type: entry.type || "note",
    amount: entry.amount === "" || entry.amount == null ? "" : Number(entry.amount),
    unit: entry.unit || "",
    note: entry.note || "",
    pendingSync: false,
  };
}

function dedupeEntries(items) {
  const byId = new Map();
  items.forEach((entry) => {
    if (!entry.id) {
      byId.set(crypto.randomUUID(), entry);
      return;
    }
    if (!byId.has(entry.id)) byId.set(entry.id, entry);
  });
  return [...byId.values()];
}

function loadSampleDay() {
  const sampleDate = new Date();
  const datePart = localDateTimeValue(sampleDate).slice(0, 10);
  const samples = [
    ["06:00", "breast", "", "", ""],
    ["07:30", "breast", "", "", ""],
    ["08:30", "breast", "", "", ""],
    ["12:30", "breast", "", "", ""],
    ["02:40", "pump", 4, "oz", ""],
    ["06:00", "pump", 1.2, "oz", ""],
    ["11:50", "pump", 3, "oz", ""],
    ["15:20", "pump", 3, "oz", ""],
    ["18:20", "pump", 4, "oz", ""],
    ["04:15", "feed", 2.5, "oz", ""],
    ["15:05", "feed", 1, "oz", ""],
    ["16:25", "feed", 3, "oz", ""],
    ["18:15", "feed", 1, "oz", ""],
    ["18:33", "feed", 3.5, "oz", ""],
    ["19:30", "feed", 2, "oz", ""],
    ["20:35", "feed", 2, "oz", ""],
  ];

  entries = samples.map(([time, type, amount, unit, note]) => ({
    id: crypto.randomUUID(),
    happenedAt: parseLocalInput(`${datePart}T${time}`),
    type,
    amount,
    unit,
    note,
    pendingSync: false,
  }));
  saveEntries();
  render();
}

els.tabs.forEach((button) => {
  button.addEventListener("click", () => {
    const view = button.dataset.viewButton;
    els.tabs.forEach((item) => item.classList.toggle("is-active", item === button));
    els.views.forEach((item) => item.classList.toggle("is-active", item.dataset.view === view));
    if (view === "summary") renderSummary();
  });
});

els.form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const entry = createEntry();
  entries.push(entry);
  saveEntries();
  render();
  els.amount.value = "";
  els.note.value = "";
  els.happenedAt.value = localDateTimeValue();

  try {
    const saved = await appendToSheet(entry);
    if (saved) await syncFromSheet();
  } catch {
    renderStatus("Sheet save failed");
  }
});

els.type.addEventListener("change", () => {
  setDefaultUnit();
});

els.nowButton.addEventListener("click", () => {
  els.happenedAt.value = localDateTimeValue();
});

els.minusTenButton.addEventListener("click", () => {
  const date = localDateTimeFromInput(els.happenedAt.value);
  date.setMinutes(date.getMinutes() - 10);
  els.happenedAt.value = localDateTimeValue(date);
});

els.sampleButton.addEventListener("click", loadSampleDay);

els.chartToggles.forEach((button) => {
  button.addEventListener("click", () => {
    chartType = button.dataset.chartType;
    renderSummary();
  });
});

els.syncMainButton.addEventListener("click", () => {
  syncFromSheet().catch(() => renderStatus("Sync failed"));
});

els.clearLocalButton.addEventListener("click", () => {
  entries = [];
  saveEntries();
  render();
});

els.refreshButton.addEventListener("click", () => {
  syncFromSheet().catch(() => renderStatus("Sync failed"));
});

els.saveSettingsButton.addEventListener("click", () => {
  settings.scriptUrl = els.scriptUrl.value.trim();
  saveSettings();
  render();
});

els.syncButton.addEventListener("click", () => {
  syncFromSheet().catch(() => renderStatus("Sync failed"));
});

els.happenedAt.value = localDateTimeValue();
els.type.value = DEFAULT_ACTIVITY;
setDefaultUnit();
els.scriptUrl.value = settings.scriptUrl || "";
els.scriptSnippet.textContent = scriptTemplate;
configureLocalOnlyUi();
render();
syncFromSheet().catch(() => renderStatus("Sync failed"));

function setDefaultUnit() {
  els.unit.value = defaultUnitFor(els.type.value);
}

function defaultUnitFor(type) {
  const defaults = {
    breast: "min",
    pump: "oz",
    feed: "oz",
    weight: "kg",
    height: "cm",
    diaper: "",
    sleep: "min",
    note: "",
  };
  return defaults[type] || "";
}
