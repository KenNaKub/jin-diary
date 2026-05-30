const STORAGE_KEY = "jin-diary-entries";
const SETTINGS_KEY = "jin-diary-settings";
const DEFAULT_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbywIWTWnpCDGlCDRJ0aNNnWUw-rz72KpgXL9lwdWYWPem-sRbE0_bu0XfV2NaiCvqPi/exec";
const APP_TIME_ZONE = "Asia/Bangkok";
const BANGKOK_OFFSET = "+07:00";

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

let entries = loadEntries();
let settings = loadSettings();
let chartType = "feed";

const els = {
  tabs: document.querySelectorAll("[data-view-button]"),
  views: document.querySelectorAll("[data-view]"),
  form: document.querySelector("#entryForm"),
  type: document.querySelector("#activityType"),
  happenedAt: document.querySelector("#happenedAt"),
  amount: document.querySelector("#amount"),
  unit: document.querySelector("#unit"),
  note: document.querySelector("#note"),
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
  if (isLocalTestHost()) return;

  document.querySelector('[data-view-button="settings"]')?.setAttribute("hidden", "");
  document.querySelector('[data-view="settings"]')?.setAttribute("hidden", "");
  els.sampleButton?.setAttribute("hidden", "");
}

function localDateTimeValue(date = new Date()) {
  return bangkokDateTimeParts(date).inputValue;
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
  node.querySelector("button").addEventListener("click", async () => {
    const deleted = await deleteFromSheet(entry.id).catch(() => false);
    if (!deleted) return;
    entries = entries.filter((item) => item.id !== entry.id);
    saveEntries();
    render();
  });
  return node;
}

function renderSummary() {
  const groups = groupByDay(sortedEntries());
  els.dayGrid.replaceChildren();

  if (!groups.length) {
    renderSummaryChart([]);
    els.dayGrid.append(emptyState("Daily cards appear after you save activities."));
    return;
  }

  renderSummaryChart(groups);

  groups.forEach(([key, dayEntries]) => {
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
        const amount = amountLabel(entry);
        line.textContent = `${formatTime(entry.happenedAt)} ${activityLabel(entry.type)}${amount ? ` (${amount})` : ""}${entry.note ? ` - ${entry.note}` : ""}`;
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
  if (!groups.length) {
    els.summaryChart.append(emptyState("Chart appears after you save activities."));
    return;
  }

  const groupsByKey = new Map(groups);
  const rows = lookbackDayKeys(30).map((key) => {
    const dayEntries = groupsByKey.get(key) || [];
    const total = sumEntries(dayEntries, chartType);
    const unit = unitFor(dayEntries, chartType, "oz");
    return {
      key,
      label: chartDayLabel(key),
      total,
      unit,
    };
  });
  const max = Math.max(...rows.map((row) => row.total), 1);

  rows.forEach((row) => {
    const item = document.createElement("div");
    item.className = "chart-bar";
    const value = formatMetric(row.total, row.unit);
    const height = row.total ? Math.max((row.total / max) * 100, 3) : 0;
    item.title = `${row.label}: ${value}`;
    item.innerHTML = `
      <div class="chart-bar-track" aria-label="${row.label}, ${value}">
        <div class="chart-bar-fill" style="height: ${height}%"></div>
      </div>
      <span>${row.label}</span>
    `;
    els.summaryChart.append(item);
  });
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
  const weight = items.filter((entry) => entry.type === "weight").at(-1);

  return {
    feed: countAndAmount(entriesFor("feed"), sumEntries(items, "feed"), unitFor(items, "feed", "oz")),
    pump: countAndAmount(entriesFor("pump"), sumEntries(items, "pump"), unitFor(items, "pump", "oz")),
    breast: countAndAmount(entriesFor("breast"), sumEntries(items, "breast"), unitFor(items, "breast", "oz")),
    weight: weight ? amountLabel(weight) : "-",
  };
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

function emptyState(text) {
  const node = document.createElement("div");
  node.className = "empty-state";
  node.textContent = text;
  return node;
}

function createEntry() {
  return {
    id: crypto.randomUUID(),
    happenedAt: parseLocalInput(els.happenedAt.value),
    type: els.type.value,
    amount: els.amount.value === "" ? "" : Number(els.amount.value),
    unit: els.unit.value,
    note: els.note.value.trim(),
    createdAt: currentBangkokTimestamp(),
    pendingSync: true,
  };
}

function saveEntries() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(entries));
}

function loadEntries() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY)) || [];
  } catch {
    return [];
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
  if (options.showStatus !== false) renderStatus(saved ? "Saved to sheet" : "Sheet save failed");
  return saved;
}

async function deleteFromSheet(id) {
  if (!settings.scriptUrl) return true;
  renderStatus("Deleting from sheet...");
  const payload = await jsonp("delete", { id });
  const deleted = payload.ok && payload.id === id;
  renderStatus(deleted ? "Deleted from sheet" : "Sheet delete failed");
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
    const sheetEntries = dedupeEntries(payload.entries.map(normalizeEntry));
    const sheetIds = new Set(sheetEntries.map((entry) => entry.id));
    const localOnlyEntries = entries.filter((entry) => entry.pendingSync && entry.id && !sheetIds.has(entry.id));

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

      entries = dedupeEntries(refreshed.entries.map(normalizeEntry));
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
    ["21:00", "weight", 7.355, "kg", ""],
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
  const defaults = {
    breast: "",
    pump: "oz",
    feed: "oz",
    weight: "kg",
    diaper: "",
    sleep: "min",
    note: "",
  };
  els.unit.value = defaults[els.type.value] || "";
});

els.nowButton.addEventListener("click", () => {
  els.happenedAt.value = localDateTimeValue();
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
els.scriptUrl.value = settings.scriptUrl || "";
els.scriptSnippet.textContent = scriptTemplate;
configureLocalOnlyUi();
render();
syncFromSheet().catch(() => renderStatus("Sync failed"));
