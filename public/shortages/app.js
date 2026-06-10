const state = {
  preview: null,
  sessions: [],
  daySessions: [],
  workSessions: [],
  activeDaySession: null,
  activeSessionId: null,
  editMode: "new",
  shortagesOnly: true,
  currentPage: "sessions",
  selectedMancoSessionIds: [],
};

const ACTIVE_WORK_SESSION_STORAGE_KEY = "orderflow-active-work-session-id";
const HAVI_UIEN_SETTINGS_KEY = "orderflow-havi-uien-settings";
const pathname = window.location.pathname;
const API_PREFIX = pathname === "/shortage" || pathname.startsWith("/shortage/") ? "/shortage/api" : "/api";

function apiPath(path) {
  return `${API_PREFIX}${path}`;
}

const sessionForm = document.getElementById("sessionForm");
const sessionDate = document.getElementById("sessionDate");
const sessionName = document.getElementById("sessionName");
const daySessionForm = document.getElementById("daySessionForm");
const daySessionDate = document.getElementById("daySessionDate");
const daySessionName = document.getElementById("daySessionName");
const daySessionWeekday = document.getElementById("daySessionWeekday");
const daySessionList = document.getElementById("daySessionList");
const activeSessionTitle = document.getElementById("activeSessionTitle");
const ordersSessionPanel = document.getElementById("ordersSessionPanel");
const quickUpdateButton = document.getElementById("quickUpdateButton");
const backToHistoryButton = document.getElementById("backToHistoryButton");
const exportMancosButton = document.getElementById("exportMancosButton");
const exportSelectedMancosButton = document.getElementById("exportSelectedMancosButton");
const emailInput = document.getElementById("emailInput");
const excelInput = document.getElementById("excelInput");
const selectedWeekday = document.getElementById("selectedWeekday");
const saveSessionButton = document.getElementById("saveSessionButton");
const statusText = document.getElementById("statusText");
const clientTitle = document.getElementById("clientTitle");
const confidenceBadge = document.getElementById("confidenceBadge");
const deliveryPoint = document.getElementById("deliveryPoint");
const orderedTotal = document.getElementById("orderedTotal");
const shortageTotal = document.getElementById("shortageTotal");
const shortageRate = document.getElementById("shortageRate");
const itemsBody = document.getElementById("itemsBody");
const dcHeaderCell = document.getElementById("dcHeaderCell");
const shortageOnlyToggle = document.getElementById("shortageOnlyToggle");
const historyDate = document.getElementById("historyDate");
const historyDateList = document.getElementById("historyDateList");
const historyList = document.getElementById("historyList");
const analyticsButton = document.getElementById("analyticsButton");
const analyticsPage = document.getElementById("analyticsPage");
const analyticsTitle = document.getElementById("analyticsTitle");
const exportDayAnalyticsButton = document.getElementById("exportDayAnalyticsButton");
const backFromAnalyticsButton = document.getElementById("backFromAnalyticsButton");
const clientOverviewBody = document.getElementById("clientOverviewBody");
const clientRateBars = document.getElementById("clientRateBars");
const topArticleValue = document.getElementById("topArticleValue");
const topArticleBars = document.getElementById("topArticleBars");
const analyticsLegend = document.getElementById("analyticsLegend");
const topArticlesBody = document.getElementById("topArticlesBody");
const sessionsPage = document.getElementById("sessionsPage");
const workspacePage = document.getElementById("workspacePage");
const historyPage = document.getElementById("historyPage");
const sessionsNav = document.getElementById("sessionsNav");
const ordersNav = document.getElementById("ordersNav");
const historyNav = document.getElementById("historyNav");

daySessionDate.value = new Date().toISOString().slice(0, 10);
setDefaultDaySessionName();

window.addEventListener("load", loadInitialData);
setInterval(syncSharedData, 8000);
daySessionDate.addEventListener("change", () => {
  setDefaultDaySessionName();
  updateDaySessionWeekday();
});
daySessionForm.addEventListener("submit", createDaySession);
emailInput.addEventListener("change", handleGreenOpsEmailUpload);
excelInput.addEventListener("change", handleUpload);
sessionForm.addEventListener("submit", saveSession);
quickUpdateButton.addEventListener("click", saveSession);
backToHistoryButton.addEventListener("click", () => showPage("history"));
exportMancosButton.addEventListener("click", exportMancos);
exportSelectedMancosButton?.addEventListener("click", exportSelectedMancos);
historyDate.addEventListener("input", renderHistory);
itemsBody.addEventListener("input", handleShortageInput);
shortageOnlyToggle.addEventListener("change", () => {
  state.shortagesOnly = shortageOnlyToggle.checked;
  renderPreview();
});
analyticsButton.addEventListener("click", () => showPage("analytics"));
exportDayAnalyticsButton?.addEventListener("click", exportDayAnalytics);
backFromAnalyticsButton.addEventListener("click", () => showPage("history"));
sessionsNav?.addEventListener("click", () => showPage("sessions"));
ordersNav?.addEventListener("click", () => showPage("workspace"));
historyNav?.addEventListener("click", () => showPage("history"));
daySessionList.addEventListener("click", (event) => {
  const deleteButton = event.target.closest("[data-delete-day-session-id]");
  if (deleteButton) {
    deleteDaySession(deleteButton.dataset.deleteDaySessionId);
    return;
  }
  const openButton = event.target.closest("[data-day-session-id]");
  if (!openButton) return;
  openDaySession(openButton.dataset.daySessionId);
});
historyDateList.addEventListener("click", (event) => {
  const button = event.target.closest("[data-history-date]");
  if (!button) return;
  historyDate.value = button.dataset.historyDate;
  renderHistory();
});
historyList.addEventListener("change", (event) => {
  const checkbox = event.target.closest("[data-select-manco-session-id]");
  if (!checkbox) return;
  toggleMancoSessionSelection(checkbox.dataset.selectMancoSessionId, checkbox.checked);
});

initCustomDatePickers();
updateDaySessionWeekday();
renderPreview();

async function handleGreenOpsEmailUpload() {
  if (!state.activeDaySession) {
    statusText.textContent = "Open an OrderFlow session before uploading an email order.";
    showPage("sessions");
    return;
  }
  const file = emailInput.files[0];
  if (!file) return;

  statusText.textContent = "Reading OrderFlow email and saving it to the active Manco session...";
  saveSessionButton.disabled = true;

  const formData = new FormData();
  formData.append("email", file);
  formData.append("date", state.activeDaySession.date || "");
  formData.append("name", state.activeDaySession.name || "");
  formData.append("workSessionId", state.activeDaySession.id || "");

  const response = await fetch(apiPath("/orders/ingest"), { method: "POST", body: formData });
  const payload = await response.json();
  if (!response.ok) {
    state.preview = null;
    renderPreview();
    statusText.textContent = payload.error || "Could not read the OrderFlow email order.";
    return;
  }

  state.sessions = normalizeDisplayObject(payload.sessions || state.sessions);
  const savedSessions = normalizeDisplayObject(payload.savedSessions || []);
  if (savedSessions.length) {
    const firstSession = savedSessions[0];
    state.activeSessionId = firstSession.id;
    state.editMode = "existing";
    state.activeDaySession = {
      id: firstSession.workSessionId || state.activeDaySession?.id || `saved-${firstSession.date}`,
      date: firstSession.date,
      weekday: firstSession.weekday,
      name: firstSession.name,
      source: "work",
    };
    state.preview = clonePreview(firstSession.preview || {});
    applyActiveDaySession();
    renderPreview();
    saveSessionButton.textContent = "Update delivery point";
    saveSessionButton.disabled = false;
    quickUpdateButton.disabled = false;
  }

  renderHistory();
  renderDaySessions();
  statusText.textContent = `${file.name} saved to ${state.activeDaySession.name}. ${savedSessions.length} order${savedSessions.length === 1 ? "" : "s"} ready for Manco control.`;
}

async function handleUpload() {
  if (!state.activeDaySession) {
    statusText.textContent = "Open an OrderFlow session before uploading an order file.";
    showPage("sessions");
    return;
  }
  const file = excelInput.files[0];
  if (!file) return;

  statusText.textContent = "Reading Excel file and calculating Manco´s...";
  saveSessionButton.disabled = true;

  const formData = new FormData();
  formData.append("file", file);

  const response = await fetch(apiPath("/shortages/parse"), { method: "POST", body: formData });
  const payload = await response.json();
  if (!response.ok) {
    state.preview = null;
    renderPreview();
    statusText.textContent = payload.error || "Could not read the uploaded file.";
    return;
  }

  state.preview = normalizeEditablePreview(normalizeDisplayObject(payload.preview));
  state.activeSessionId = null;
  state.editMode = "new";
  applyActiveDaySession();
  renderPreview();
  saveSessionButton.disabled = false;
  saveSessionButton.textContent = "Save session";
  statusText.textContent = `${file.name} loaded. Review the result and save the session.`;
}

async function saveSession(event) {
  event.preventDefault();
  if (!state.preview) return;

  statusText.textContent = "Saving session...";
  const isUpdate = Boolean(state.activeSessionId);
  const previewForSave = buildPreviewForDataTransfer(state.preview);
  const response = await fetch(apiPath("/sessions"), {
    method: isUpdate ? "PUT" : "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      id: state.activeSessionId,
      date: sessionDate.value,
      name: sessionName.value,
      workSessionId: state.activeDaySession?.id || "",
      preview: previewForSave,
    }),
  });
  const payload = await response.json();
  if (!response.ok) {
    statusText.textContent = payload.error || "Could not save the session.";
    return;
  }

  state.sessions = normalizeDisplayObject(payload.sessions || []);
  state.activeSessionId = payload.session?.id || state.activeSessionId;
  statusText.textContent = isUpdate ? "Delivery point updated." : "Session saved.";
  saveSessionButton.textContent = "Update delivery point";
  saveSessionButton.disabled = false;
  quickUpdateButton.disabled = false;
  renderHistory();
}

async function loadInitialData() {
  await Promise.all([loadOrderFlowSessions(), loadSessions()]);
  selectOrderFlowSessionForShortages();
  renderHistory();
  showPage("history");
}

async function syncSharedData() {
  if (state.currentPage === "workspace" && state.preview) return;
  await Promise.all([loadOrderFlowSessions(), loadSessions()]);
  selectOrderFlowSessionForShortages();
  if (state.currentPage === "history") renderHistory();
  if (state.currentPage === "analytics") renderAnalytics();
  if (state.currentPage === "sessions") renderDaySessions();
}

async function loadSessions() {
  const response = await fetch(apiPath("/sessions"));
  const payload = await response.json();
  if (!response.ok) {
    statusText.textContent = payload.error || "Could not load saved sessions.";
    return;
  }
  state.sessions = normalizeDisplayObject(payload.sessions || []);
  renderHistory();
  renderDaySessions();
}

async function loadDaySessions() {
  const response = await fetch(apiPath("/day-sessions"));
  const payload = await response.json();
  if (!response.ok) {
    statusText.textContent = payload.error || "Could not load daily sessions.";
    return;
  }
  state.daySessions = payload.daySessions || [];
  renderDaySessions();
}

async function loadOrderFlowSessions() {
  try {
    const response = await fetch("/api/work_sessions", { credentials: "include" });
    const payload = await response.json();
    if (!response.ok) {
      state.workSessions = [];
      await loadDaySessions();
      return;
    }
    state.workSessions = Array.isArray(payload.sessions) ? payload.sessions : [];
    renderDaySessions();
  } catch {
    state.workSessions = [];
    await loadDaySessions();
  }
}

async function createDaySession(event) {
  event.preventDefault();
  const response = await fetch(apiPath("/day-sessions"), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      date: daySessionDate.value,
      name: daySessionName.value,
    }),
  });
  const payload = await response.json();
  if (!response.ok) {
    daySessionList.innerHTML = `<p class="empty-cell">${escapeHtml(payload.error || "Could not create session.")}</p>`;
    return;
  }
  state.daySessions = payload.daySessions || [];
  renderDaySessions();
  openDaySession(payload.daySession.id);
}

function showPage(page) {
  state.currentPage = page;
  const showSessions = page === "sessions";
  const showWorkspace = page === "workspace";
  const showHistory = page === "history";
  const showAnalytics = page === "analytics";
  sessionsPage.classList.toggle("hidden", !showSessions);
  workspacePage.classList.toggle("hidden", !showWorkspace);
  historyPage.classList.toggle("hidden", !showHistory);
  analyticsPage.classList.toggle("hidden", !showAnalytics);
  sessionsNav?.classList.toggle("active", showSessions);
  ordersNav?.classList.toggle("active", showWorkspace);
  historyNav?.classList.toggle("active", showHistory || showAnalytics);
  renderWorkspaceMode();
  if (showHistory) renderHistory();
  if (showAnalytics) renderAnalytics();
  if (showSessions) renderDaySessions();
}

function selectOrderFlowSessionForShortages() {
  if (state.activeDaySession?.id) return;

  const storedId = window.localStorage.getItem(ACTIVE_WORK_SESSION_STORAGE_KEY) || "";
  const daySessions = visibleDaySessions();
  const storedSession = daySessions.find((session) => session.id === storedId);
  const sessionWithOrders = daySessions.find((session) => state.sessions.some((entry) => sessionMatchesWorkSession(entry, session)));
  const selected = storedSession || sessionWithOrders || daySessions[0] || null;
  if (!selected) return;

  state.activeDaySession = selected;
  applyActiveDaySession();
  historyDate.value = selected.date || "";
}

function renderPreview() {
  const preview = state.preview;
  const isHaviUien = isHaviUienPreview(preview);
  renderItemsTableHeader(isHaviUien);
  if (!preview) {
    shortageOnlyToggle.checked = state.shortagesOnly;
    shortageOnlyToggle.disabled = true;
    clientTitle.textContent = "No file loaded";
    confidenceBadge.textContent = "0%";
    deliveryPoint.textContent = "-";
    orderedTotal.textContent = "0";
    shortageTotal.textContent = "0";
    shortageRate.textContent = "0%";
    saveSessionButton.textContent = "Save session";
    saveSessionButton.disabled = true;
    quickUpdateButton.classList.add("hidden");
    quickUpdateButton.disabled = true;
    exportMancosButton.classList.add("hidden");
    exportMancosButton.disabled = true;
    itemsBody.innerHTML = `<tr><td colspan="${itemsTableColumnCount(false)}" class="empty-cell">Upload a file to preview Manco lines.</td></tr>`;
    return;
  }

  clientTitle.textContent = repairDisplayText(preview.client);
  const hasMancos = hasShortageLines();
  const showMancosOnly = state.shortagesOnly && hasMancos;
  shortageOnlyToggle.disabled = false;
  shortageOnlyToggle.checked = showMancosOnly;
  confidenceBadge.textContent = `${preview.confidence}%`;
  confidenceBadge.classList.toggle("warning", preview.confidence < 50);
  deliveryPoint.textContent = repairDisplayText(preview.deliveryPoint);
  orderedTotal.textContent = formatNumber(preview.orderedTotal);
  shortageTotal.textContent = formatNumber(preview.shortageTotal);
  shortageRate.textContent = `${preview.shortagePercentage}%`;
  saveSessionButton.disabled = false;
  quickUpdateButton.disabled = false;
  exportMancosButton.disabled = !hasShortageLines();
  renderWorkspaceMode();

  const itemEntries = (preview.items || []).map((item, index) => ({
    item: isHaviUien ? mapHaviUienMancoItem(item) : item,
    index,
  }));
  const visibleItems = showMancosOnly
    ? itemEntries.filter(({ item }) => Number(item.shortageQuantity || 0) > 0)
    : itemEntries;

  if (!visibleItems.length) {
    itemsBody.innerHTML = `<tr><td colspan="${itemsTableColumnCount(isHaviUien)}" class="empty-cell">${
      showMancosOnly ? "No articles with Manco´s found." : "No order lines with Manco data were found."
    }</td></tr>`;
    return;
  }

  itemsBody.innerHTML = visibleItems
    .map(
      ({ item, index }) => `
        <tr>
          ${isHaviUien ? `<td>${escapeHtml(item.dc || "-")}</td>` : ""}
          <td>${escapeHtml(item.article)}</td>
          <td>${escapeHtml(item.description || "-")}</td>
          <td>${formatNumber(item.orderedQuantity)}</td>
          <td>${formatNumber(item.deliveredQuantity)}</td>
          <td>
            <input
              class="shortage-input"
              type="number"
              min="0"
              max="${Number(item.orderedQuantity || 0)}"
              step="1"
              value="${item.shortageQuantity}"
              data-index="${index}"
              aria-label="Manco for ${escapeHtml(item.article)}"
            >
          </td>
          <td>${item.shortagePercentage}%</td>
          <td></td>
        </tr>
      `,
    )
    .join("");
}

function handleShortageInput(event) {
  if (!event.target.classList.contains("shortage-input") || !state.preview) return;
  const index = Number(event.target.dataset.index);
  const item = state.preview.items[index];
  if (!item) return;

  const ordered = Number(item.orderedQuantity || 0);
  const shortage = clampNumber(event.target.value, 0, ordered);
  item.shortageQuantity = cleanNumber(shortage);
  item.deliveredQuantity = cleanNumber(Math.max(ordered - shortage, 0));
  item.shortagePercentage = ordered ? roundPercent((shortage / ordered) * 100) : 0;
  recalculatePreviewTotals();
  const row = event.target.closest("tr");
  if (row) {
    const isHaviUien = isHaviUienPreview(state.preview);
    row.children[isHaviUien ? 4 : 3].textContent = formatNumber(item.deliveredQuantity);
    row.children[isHaviUien ? 6 : 5].textContent = `${item.shortagePercentage}%`;
  }
  orderedTotal.textContent = formatNumber(state.preview.orderedTotal);
  shortageTotal.textContent = formatNumber(state.preview.shortageTotal);
  shortageRate.textContent = `${state.preview.shortagePercentage}%`;
  exportMancosButton.disabled = !hasShortageLines();
}

function normalizeEditablePreview(preview) {
  const normalized = { ...preview };
  normalized.items = (preview.items || []).map((item) => {
    const ordered = Number(item.orderedQuantity || 0);
    return {
      ...item,
      deliveredQuantity: cleanNumber(ordered),
      shortageQuantity: 0,
      shortagePercentage: 0,
    };
  });
  recalculateTotals(normalized);
  return normalized;
}

function renderItemsTableHeader(isHaviUien = false) {
  dcHeaderCell?.classList.toggle("hidden", !isHaviUien);
}

function itemsTableColumnCount(isHaviUien = isHaviUienPreview(state.preview)) {
  return isHaviUien ? 8 : 7;
}

function getHaviUienSettings() {
  try {
    const raw = window.localStorage.getItem(HAVI_UIEN_SETTINGS_KEY);
    const parsed = raw ? JSON.parse(raw) : {};
    return {
      article: String(parsed.article || "").trim(),
      description: String(parsed.description || "").trim(),
    };
  } catch {
    return { article: "", description: "" };
  }
}

function isHaviUienPreview(preview) {
  if (!preview) return false;
  const text = normalizeText(`${preview.client || ""} ${preview.deliveryPoint || ""} ${preview.greenopsFatrans || ""} ${preview.greenopsCustomer || ""}`);
  return text.includes("havi") && (text.includes("uien") || text.includes("onion"));
}

function mapHaviUienMancoItem(item) {
  const settings = getHaviUienSettings();
  return {
    ...item,
    dc: buildHaviUienDcLabel(item),
    orderNumber: String(item?.orderNumber || item?.description || item?.reference || "").trim(),
    article: settings.article,
    description: settings.description,
  };
}

function buildHaviUienDcLabel(item) {
  const explicit = String(item?.dc || "").trim();
  if (explicit) return explicit;
  return [item?.article, item?.description]
    .map((value) => String(value || "").trim())
    .filter(Boolean)
    .join(" - ");
}

function buildPreviewForDataTransfer(preview) {
  const cloned = clonePreview(preview || {});
  if (!isHaviUienPreview(cloned)) return cloned;
  cloned.items = (cloned.items || []).map((item) => mapHaviUienMancoItem(item));
  return cloned;
}

function recalculatePreviewTotals() {
  recalculateTotals(state.preview);
}

function recalculateTotals(preview) {
  const ordered = (preview.items || []).reduce((sum, item) => sum + Number(item.orderedQuantity || 0), 0);
  const shortage = (preview.items || []).reduce((sum, item) => sum + Number(item.shortageQuantity || 0), 0);
  preview.orderedTotal = cleanNumber(ordered);
  preview.shortageTotal = cleanNumber(shortage);
  preview.shortagePercentage = ordered ? roundPercent((shortage / ordered) * 100) : 0;
}

function renderHistory() {
  renderHistoryDates();
  const filterDate = historyDate.value || mostRecentSessionDate();
  const sessions = getVisibleHistorySessions(filterDate);
  syncSelectedMancoSessions(sessions);
  updateSelectedMancosExportButton();
  if (filterDate && historyDate.value !== filterDate) {
    historyDate.value = filterDate;
  }

  if (!sessions.length) {
    const sessionName = state.activeDaySession?.name || "this session";
    historyList.innerHTML = `<p class="empty-cell">No orders have been added to ${escapeHtml(sessionName)} yet.</p>`;
    updateSelectedMancosExportButton();
    return;
  }

  const grouped = groupSessionsByClient(sessions);
  historyList.innerHTML = Object.entries(grouped)
    .map(([client, clientSessions]) => {
      const ordered = clientSessions.reduce((sum, session) => sum + Number(session.preview?.orderedTotal || 0), 0);
      const shortage = clientSessions.reduce((sum, session) => sum + Number(session.preview?.shortageTotal || 0), 0);
      const rate = ordered ? roundPercent((shortage / ordered) * 100) : 0;
      return `
        <article class="client-history-card">
          <div class="client-history-header">
            <div>
              <p class="label">Client</p>
              <h3>${escapeHtml(client)}</h3>
            </div>
            <div class="client-history-totals">
              <span>${clientSessions.length} delivery point${clientSessions.length === 1 ? "" : "s"}</span>
              <strong>${rate}%</strong>
            </div>
          </div>
          <div class="delivery-point-list">
            ${clientSessions.map(renderDeliveryPointSession).join("")}
          </div>
        </article>
      `;
    })
    .join("");
  updateSelectedMancosExportButton();
}

function getVisibleHistorySessions(filterDate = historyDate.value || mostRecentSessionDate()) {
  const activeWorkSessionId = state.activeDaySession?.id || "";
  return state.sessions.filter((session) => {
    if (activeWorkSessionId && session.workSessionId) {
      return session.workSessionId === activeWorkSessionId;
    }
    return filterDate ? session.date === filterDate : true;
  });
}

function renderAnalytics() {
  const filterDate = historyDate.value || mostRecentSessionDate();
  const sessions = filterDate ? state.sessions.filter((session) => session.date === filterDate) : state.sessions;
  analyticsTitle.textContent = filterDate ? `Session analytics · ${filterDate}` : "Session analytics";

  if (!sessions.length) {
    clientOverviewBody.innerHTML = '<tr><td colspan="5" class="empty-cell">No sessions found for this date.</td></tr>';
    clientRateBars.innerHTML = '<p class="empty-cell">No data to chart.</p>';
    topArticleBars.innerHTML = '<p class="empty-cell">No article Manco´s found.</p>';
    topArticlesBody.innerHTML = '<tr><td colspan="4" class="empty-cell">No article Manco´s found.</td></tr>';
    analyticsLegend.innerHTML = "";
    topArticleValue.textContent = "-";
    return;
  }

  const clientRows = buildClientAnalytics(sessions);
  const articleRows = buildArticleAnalytics(sessions);
  renderClientOverview(clientRows);
  renderClientRateBars(clientRows);
  renderTopArticles(articleRows, clientRows);
}

async function exportDayAnalytics() {
  const filterDate = historyDate.value || mostRecentSessionDate();
  if (!filterDate) {
    statusText.textContent = "Choose a day before exporting analytics.";
    return;
  }

  exportDayAnalyticsButton.disabled = true;
  statusText.textContent = "Generating day analytics export...";
  try {
    const response = await fetch(apiPath("/export-day-analytics"), {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        date: filterDate,
        workSessionId: state.activeDaySession?.id || "",
        haviUienSettings: getHaviUienSettings(),
      }),
    });

    if (!response.ok) {
      const payload = await response.json();
      statusText.textContent = payload.error || "Could not export day analytics.";
      return;
    }

    const blob = await response.blob();
    const disposition = response.headers.get("Content-Disposition") || "";
    const match = disposition.match(/filename="(.+)"/);
    const fileName = match ? match[1] : `Manco Analytics - ${filterDate}.xlsx`;
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.click();
    URL.revokeObjectURL(url);
    statusText.textContent = `${fileName} downloaded.`;
  } catch (error) {
    console.error("Day analytics export error:", error);
    statusText.textContent = "Could not export day analytics.";
  } finally {
    exportDayAnalyticsButton.disabled = false;
  }
}

function buildClientAnalytics(sessions) {
  const clients = new Map();
  for (const session of sessions) {
    const preview = session.preview || {};
    const client = repairDisplayText(preview.client || "Unknown client");
    const current = clients.get(client) || { client, ordered: 0, delivered: 0, shortage: 0 };
    for (const item of preview.items || []) {
      const ordered = Number(item.orderedQuantity || 0);
      const shortage = Number(item.shortageQuantity || 0);
      current.ordered += ordered;
      current.shortage += shortage;
      current.delivered += Math.max(ordered - shortage, 0);
    }
    clients.set(client, current);
  }
  return Array.from(clients.values())
    .map((row) => ({
      ...row,
      deliveryRate: row.ordered ? roundPercent((row.delivered / row.ordered) * 100) : 0,
      shortageRate: row.ordered ? roundPercent((row.shortage / row.ordered) * 100) : 0,
    }))
    .sort((a, b) => b.shortage - a.shortage);
}

function buildArticleAnalytics(sessions) {
  const articles = new Map();
  for (const session of sessions) {
    const preview = buildPreviewForDataTransfer(session.preview || {});
    const client = repairDisplayText(preview.client || "Unknown client");
    for (const item of preview.items || []) {
      const shortage = Number(item.shortageQuantity || 0);
      if (shortage <= 0) continue;
      const key = `${item.article}::${item.description || ""}`;
      const current = articles.get(key) || {
        article: item.article,
        description: repairDisplayText(item.description || ""),
        shortage: 0,
        clients: {},
      };
      current.shortage += shortage;
      current.clients[client] = (current.clients[client] || 0) + shortage;
      articles.set(key, current);
    }
  }
  return Array.from(articles.values()).sort((a, b) => b.shortage - a.shortage).slice(0, 10);
}

function renderClientOverview(rows) {
  if (!rows.length) {
    clientOverviewBody.innerHTML = '<tr><td colspan="5" class="empty-cell">No client data found.</td></tr>';
    return;
  }
  clientOverviewBody.innerHTML = rows
    .map(
      (row) => `
        <tr>
          <td><strong>${escapeHtml(row.client)}</strong></td>
          <td>${formatNumber(row.ordered)}</td>
          <td>${formatNumber(row.delivered)}</td>
          <td class="shortage-text">${formatNumber(row.shortage)}</td>
          <td><span class="delivery-badge">${row.deliveryRate}%</span></td>
        </tr>
      `,
    )
    .join("");
}

function renderClientRateBars(rows) {
  if (!rows.length) {
    clientRateBars.innerHTML = '<p class="empty-cell">No client data found.</p>';
    return;
  }
  const maxRate = Math.max(...rows.map((row) => row.shortageRate), 1);
  clientRateBars.innerHTML = rows
    .map((row, index) => {
      const height = Math.max((row.shortageRate / maxRate) * 100, row.shortageRate > 0 ? 8 : 2);
      return `
        <div class="vertical-bar-item">
          <span>${row.shortageRate}%</span>
          <div class="vertical-bar-track">
            <div class="vertical-bar-fill client-color-${index % 6}" style="height: ${height}%"></div>
          </div>
          <strong>${escapeHtml(row.client)}</strong>
        </div>
      `;
    })
    .join("");
}

function renderTopArticles(rows, clientRows) {
  if (!rows.length) {
    topArticleBars.innerHTML = '<p class="empty-cell">No article Manco´s found.</p>';
    topArticlesBody.innerHTML = '<tr><td colspan="4" class="empty-cell">No article Manco´s found.</td></tr>';
    analyticsLegend.innerHTML = "";
    topArticleValue.textContent = "-";
    return;
  }

  const clients = clientRows.map((row) => row.client);
  const colorMap = Object.fromEntries(clients.map((client, index) => [client, index % 6]));
  const maxShortage = Math.max(...rows.map((row) => row.shortage), 1);
  topArticleValue.textContent = rows[0].article || "-";

  topArticleBars.innerHTML = rows
    .map((row) => {
      const segments = Object.entries(row.clients)
        .map(([client, qty]) => {
          const width = (qty / maxShortage) * 100;
          return `<span class="stack-segment client-color-${colorMap[client] ?? 0}" style="width: ${width}%" title="${escapeHtml(client)}: ${formatNumber(qty)}"></span>`;
        })
        .join("");
      return `
        <div class="horizontal-bar-row">
          <span class="bar-label">${escapeHtml(row.article)}</span>
          <div class="stacked-bar">${segments}</div>
          <strong>${formatNumber(row.shortage)}</strong>
        </div>
      `;
    })
    .join("");

  analyticsLegend.innerHTML = clients
    .map((client) => `<span><i class="client-color-${colorMap[client]}"></i>${escapeHtml(client)}</span>`)
    .join("");

  topArticlesBody.innerHTML = rows
    .map((row, index) => {
      const clientText = Object.entries(row.clients)
        .map(([client, qty]) => `${escapeHtml(client)}: ${formatNumber(qty)}`)
        .join("<br>");
      return `
        <tr>
          <td>#${index + 1}</td>
          <td><strong>${escapeHtml(row.article)}</strong><br><span class="muted-cell">${escapeHtml(row.description || "-")}</span></td>
          <td class="shortage-text">${formatNumber(row.shortage)}</td>
          <td>${clientText}</td>
        </tr>
      `;
    })
    .join("");
}

function renderDaySessions() {
  const daySessions = visibleDaySessions();
  if (!daySessions.length) {
    daySessionList.innerHTML = '<p class="empty-cell">No OrderFlow sessions found. Create a session on the OrderFlow home page first.</p>';
    return;
  }
  daySessionList.innerHTML = daySessions
    .map((session) => {
      const orderCount = state.sessions.filter((entry) => sessionMatchesWorkSession(entry, session)).length;
      const deleteButton = session.source === "legacy"
        ? `<button class="delete-session-button" type="button" data-delete-day-session-id="${escapeHtml(session.id)}">Delete</button>`
        : "";
      return `
        <article class="day-session-card">
          <div>
            <h3>${escapeHtml(session.name || "OrderFlow session")}</h3>
            <p>${escapeHtml(session.weekday || "")}, ${escapeHtml(session.date || "")}</p>
          </div>
          <div class="history-stat">
            <span>Orders</span>
            <strong>${orderCount}</strong>
          </div>
          <div class="day-session-actions">
            <button class="open-session-button" type="button" data-day-session-id="${escapeHtml(session.id)}">Open session</button>
            ${deleteButton}
          </div>
        </article>
      `;
    })
    .join("");
}

async function deleteDaySession(sessionId) {
  const session = visibleDaySessions().find((entry) => entry.id === sessionId);
  if (!session) return;
  const confirmed = window.confirm(`Delete ${session.name}? This will also remove saved orders for ${session.date}.`);
  if (!confirmed) return;

  const response = await fetch(apiPath("/day-sessions"), {
    method: "DELETE",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ id: session.id, date: session.date }),
  });
  const payload = await response.json();
  if (!response.ok) {
    daySessionList.innerHTML = `<p class="empty-cell">${escapeHtml(payload.error || "Could not delete session.")}</p>`;
    return;
  }
  state.daySessions = payload.daySessions || [];
  state.sessions = normalizeDisplayObject(payload.sessions || []);
  if (state.activeDaySession?.date === session.date) {
    state.activeDaySession = null;
    state.activeSessionId = null;
    state.preview = null;
    renderPreview();
  }
  renderDaySessions();
  renderHistory();
}

function openDaySession(sessionId) {
  const session = visibleDaySessions().find((entry) => entry.id === sessionId);
  if (!session) return;
  state.activeDaySession = session;
  state.activeSessionId = null;
  state.preview = null;
  state.editMode = "new";
  applyActiveDaySession();
  renderPreview();
  statusText.textContent = `${session.name} opened. Upload an OrderFlow email order or Manco Excel file.`;
  showPage("workspace");
}

function visibleDaySessions() {
  if (state.workSessions.length) {
    return [...state.workSessions]
      .filter((session) => session && session.id && session.date)
      .map((session) => ({
        id: session.id,
        date: session.date,
        weekday: weekdayFromDate(session.date),
        name: session.name || `${weekdayFromDate(session.date)} OrderFlow session`,
        source: "work",
        createdAt: session.createdAt || "",
      }))
      .sort((a, b) => String(b.createdAt || b.date || "").localeCompare(String(a.createdAt || a.date || "")));
  }
  const byDate = new Map();
  for (const session of state.daySessions) {
    if (!session.date) continue;
    byDate.set(session.date, { ...session, source: "legacy" });
  }
  for (const session of state.sessions) {
    if (!session.date || byDate.has(session.date)) continue;
    byDate.set(session.date, {
      id: `saved-${session.date}`,
      date: session.date,
      weekday: session.weekday || weekdayFromDate(session.date),
      name: session.name || `${weekdayFromDate(session.date)} Manco review`,
      source: "legacy",
    });
  }
  return Array.from(byDate.values()).sort((a, b) => String(b.date || "").localeCompare(String(a.date || "")));
}

function sessionMatchesWorkSession(entry, workSession) {
  if (!entry || !workSession) return false;
  if (entry.workSessionId && workSession.id) {
    return entry.workSessionId === workSession.id;
  }
  return entry.date === workSession.date;
}

function applyActiveDaySession() {
  if (!state.activeDaySession) return;
  if (state.activeDaySession.id) {
    window.localStorage.setItem(ACTIVE_WORK_SESSION_STORAGE_KEY, state.activeDaySession.id);
  }
  sessionDate.value = state.activeDaySession.date || "";
  sessionName.value = state.activeDaySession.name || "";
  activeSessionTitle.textContent = state.activeDaySession.name || "Active session";
  const title = document.getElementById("shortageSessionTitle");
  if (title) {
    title.textContent = state.activeDaySession.name || "Orders added in this session";
  }
  updateWeekday();
}

function renderHistoryDates() {
  const dates = Array.from(
    state.sessions.reduce((map, session) => {
      if (!session.date) return map;
      const current = map.get(session.date) || { date: session.date, weekday: session.weekday || "", count: 0 };
      current.count += 1;
      map.set(session.date, current);
      return map;
    }, new Map()).values(),
  ).sort((a, b) => b.date.localeCompare(a.date));

  if (!dates.length) {
    historyDateList.innerHTML = "";
    return;
  }

  historyDateList.innerHTML = dates
    .map(
      (entry) => `
        <button class="date-chip ${entry.date === (historyDate.value || mostRecentSessionDate()) ? "active" : ""}" type="button" data-history-date="${escapeHtml(entry.date)}">
          <span>${escapeHtml(entry.weekday || "Session")}</span>
          <strong>${escapeHtml(entry.date)}</strong>
          <em>${entry.count} saved</em>
        </button>
      `,
    )
    .join("");
}

function mostRecentSessionDate() {
  return state.sessions.map((session) => session.date).filter(Boolean).sort().reverse()[0] || "";
}

function groupSessionsByClient(sessions) {
  return sessions.reduce((groups, session) => {
    const client = repairDisplayText(session.preview?.client || "Unknown client");
    if (!groups[client]) groups[client] = [];
    groups[client].push(session);
    return groups;
  }, {});
}

function renderDeliveryPointSession(session) {
  const preview = session.preview || {};
  const reference = preview.greenopsReference ? ` - ${preview.greenopsReference}` : "";
  const isSelected = state.selectedMancoSessionIds.includes(session.id);
  return `
    <article class="delivery-point-card ${isSelected ? "selected" : ""}">
      <label class="manco-session-selector" title="Select order for Manco export">
        <input
          type="checkbox"
          data-select-manco-session-id="${escapeHtml(session.id)}"
          ${isSelected ? "checked" : ""}
        >
        <span></span>
      </label>
      <div>
        <h4>${escapeHtml(`${preview.deliveryPoint || "Unknown delivery point"}${reference}`)}</h4>
        <p>${escapeHtml(session.weekday || "")}, ${escapeHtml(session.date || "")} · ${escapeHtml(session.name || "Manco session")}</p>
      </div>
      <div class="history-stat">
        <span>Ordered</span>
        <strong>${formatNumber(preview.orderedTotal || 0)}</strong>
      </div>
      <div class="history-stat">
        <span>Manco</span>
        <strong>${formatNumber(preview.shortageTotal || 0)}</strong>
      </div>
      <div class="history-stat">
        <span>Rate</span>
        <strong>${preview.shortagePercentage || 0}%</strong>
      </div>
      <button class="open-session-button" type="button" data-session-id="${escapeHtml(session.id)}">Open Order</button>
    </article>
  `;
}

function sessionHasMancoLines(session) {
  return Boolean(session?.preview?.items?.some((item) => Number(item.shortageQuantity || 0) > 0));
}

function toggleMancoSessionSelection(sessionId, checked) {
  const selected = new Set(state.selectedMancoSessionIds);
  if (checked) {
    selected.add(sessionId);
  } else {
    selected.delete(sessionId);
  }
  state.selectedMancoSessionIds = Array.from(selected);
  updateSelectedMancosExportButton();
  renderHistory();
}

function syncSelectedMancoSessions(sessions = getVisibleHistorySessions()) {
  const validIds = new Set(sessions.map((session) => session.id));
  state.selectedMancoSessionIds = state.selectedMancoSessionIds.filter((id) => validIds.has(id));
}

function getSelectedMancoSessions() {
  const selectedIds = new Set(state.selectedMancoSessionIds);
  return getVisibleHistorySessions().filter((session) => selectedIds.has(session.id));
}

function updateSelectedMancosExportButton() {
  if (!exportSelectedMancosButton) return;
  const count = getSelectedMancoSessions().length;
  exportSelectedMancosButton.disabled = count === 0;
  exportSelectedMancosButton.textContent = count
    ? `Export selected Manco´s (${count})`
    : "Export selected Manco´s";
}

historyList.addEventListener("click", (event) => {
  const button = event.target.closest(".open-session-button");
  if (!button) return;
  openSavedSession(button.dataset.sessionId);
});

function openSavedSession(sessionId) {
  const session = state.sessions.find((entry) => entry.id === sessionId);
  if (!session) return;

  state.activeSessionId = session.id;
  state.editMode = "existing";
  state.activeDaySession = {
    id: session.workSessionId || `saved-${session.date}`,
    date: session.date,
    weekday: session.weekday,
    name: session.name,
    source: session.workSessionId ? "work" : "legacy",
  };
  state.preview = clonePreview(session.preview || {});
  applyActiveDaySession();
  renderPreview();
  saveSessionButton.textContent = "Update delivery point";
  saveSessionButton.disabled = false;
  statusText.textContent = `${state.preview.client} ${state.preview.deliveryPoint} opened for editing.`;
  showPage("workspace");
  document.querySelector(".preview-panel").scrollIntoView({ behavior: "smooth", block: "start" });
}

function renderWorkspaceMode() {
  const editingExisting = state.editMode === "existing" && Boolean(state.preview);
  ordersSessionPanel.classList.toggle("hidden", editingExisting);
  quickUpdateButton.classList.toggle("hidden", !editingExisting);
  backToHistoryButton.classList.toggle("hidden", !editingExisting);
  exportMancosButton.classList.toggle("hidden", !editingExisting);
  exportMancosButton.disabled = !editingExisting || !hasShortageLines();
}

function hasShortageLines() {
  return Boolean(state.preview?.items?.some((item) => Number(item.shortageQuantity || 0) > 0));
}

async function exportMancos() {
  if (!state.preview || !hasShortageLines()) {
    statusText.textContent = "There are no Manco lines to export.";
    return;
  }
  exportMancosButton.disabled = true;
  const previewForExport = buildPreviewForDataTransfer(state.preview);
  await exportMancoPayload(
    { previews: [{ preview: previewForExport, date: sessionDate.value, name: sessionName.value, id: state.activeSessionId || "" }] },
    () => {
      exportMancosButton.disabled = !hasShortageLines();
    },
  );
  return;
  const response = await fetch(apiPath("/export-mancos"), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ preview: state.preview }),
  });

  if (!response.ok) {
    const payload = await response.json();
    statusText.textContent = payload.error || "Could not export Manco´s.";
    exportMancosButton.disabled = false;
    return;
  }

  const blob = await response.blob();
  const disposition = response.headers.get("Content-Disposition") || "";
  const match = disposition.match(/filename="(.+)"/);
  const fileName = match ? match[1] : "mancos.xlsx";
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
  statusText.textContent = `${fileName} exported.`;
  exportMancosButton.disabled = !hasShortageLines();
}

async function exportSelectedMancos() {
  const selectedSessions = getSelectedMancoSessions();
  if (!selectedSessions.length) {
    statusText.textContent = "Select at least one order with Manco´s before exporting.";
    return;
  }
  exportSelectedMancosButton.disabled = true;
  await exportMancoPayload(
    {
      previews: selectedSessions.map((session) => ({
        id: session.id,
        date: session.date,
        name: session.name,
        preview: buildPreviewForDataTransfer(session.preview || {}),
      })),
    },
    updateSelectedMancosExportButton,
  );
}

async function exportMancoPayload(payload, afterComplete) {
  const response = await fetch(apiPath("/export-mancos"), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const payload = await response.json();
    statusText.textContent = payload.error || "Could not export Manco´s.";
    afterComplete?.();
    return;
  }

  const blob = await response.blob();
  const disposition = response.headers.get("Content-Disposition") || "";
  const match = disposition.match(/filename="(.+)"/);
  const fileName = match ? match[1] : "mancos.xlsx";
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
  statusText.textContent = `${fileName} exported.`;
  afterComplete?.();
}

function clonePreview(preview) {
  return normalizeDisplayObject(JSON.parse(JSON.stringify(preview)));
}

function setDefaultSessionName() {
  const weekday = weekdayFromDate(sessionDate.value);
  if (!sessionName.value || /Manco review$/i.test(sessionName.value) || /^\w+ - \d{2}-\d{2}-\d{4}$/i.test(sessionName.value)) {
    sessionName.value = `${weekday} - ${formatDateName(sessionDate.value)}`;
  }
}

function setDefaultDaySessionName() {
  const weekday = weekdayFromDate(daySessionDate.value);
  if (!daySessionName.value || /Manco review$/i.test(daySessionName.value) || /^\w+ - \d{2}-\d{2}-\d{4}$/i.test(daySessionName.value)) {
    daySessionName.value = `${weekday} - ${formatDateName(daySessionDate.value)}`;
  }
}

function updateWeekday() {
  selectedWeekday.textContent = weekdayFromDate(sessionDate.value);
}

function updateDaySessionWeekday() {
  daySessionWeekday.textContent = weekdayFromDate(daySessionDate.value);
}

const CUSTOM_DATE_MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const CUSTOM_DATE_WEEKDAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
let activeDatePicker = null;

function initCustomDatePickers() {
  document.querySelectorAll('input[type="date"]').forEach((input) => {
    if (input.dataset.customDatePicker === "1" || input.readOnly || input.disabled) return;
    input.dataset.customDatePicker = "1";
    input.type = "text";
    input.autocomplete = "off";
    input.inputMode = "none";
    input.classList.add("custom-date-input");
    input.setAttribute("placeholder", "YYYY-MM-DD");
    input.addEventListener("focus", () => openCustomDatePicker(input));
    input.addEventListener("click", () => openCustomDatePicker(input));
    input.addEventListener("keydown", (event) => {
      if (["Enter", " ", "ArrowDown"].includes(event.key)) {
        event.preventDefault();
        openCustomDatePicker(input);
      }
      if (event.key === "Escape") closeCustomDatePicker();
    });
  });
}

function openCustomDatePicker(input) {
  closeCustomDatePicker();
  const selectedDate = parseDateValue(input.value) || new Date();
  const view = { year: selectedDate.getFullYear(), month: selectedDate.getMonth() };
  const popover = document.createElement("div");
  popover.className = "custom-date-picker";
  document.body.appendChild(popover);
  activeDatePicker = { input, popover };

  const render = () => {
    popover.innerHTML = renderCustomDatePicker(input, view);
  };
  render();
  positionCustomDatePicker(input, popover);

  popover.addEventListener("click", (event) => {
    const prev = event.target.closest("[data-date-prev]");
    const next = event.target.closest("[data-date-next]");
    const dateButton = event.target.closest("[data-date-value]");
    const todayButton = event.target.closest("[data-date-today]");
    const clearButton = event.target.closest("[data-date-clear]");
    if (prev || next) {
      view.month += next ? 1 : -1;
      if (view.month < 0) {
        view.month = 11;
        view.year -= 1;
      }
      if (view.month > 11) {
        view.month = 0;
        view.year += 1;
      }
      render();
      return;
    }
    if (dateButton) {
      setDatePickerInputValue(input, dateButton.dataset.dateValue);
      closeCustomDatePicker();
      return;
    }
    if (todayButton) {
      setDatePickerInputValue(input, formatDateForInput(new Date()));
      closeCustomDatePicker();
      return;
    }
    if (clearButton) {
      setDatePickerInputValue(input, "");
      closeCustomDatePicker();
    }
  });

  setTimeout(() => {
    document.addEventListener("mousedown", handleDatePickerOutsideClick);
    window.addEventListener("resize", closeCustomDatePicker, { once: true });
    window.addEventListener("scroll", closeCustomDatePicker, { once: true, capture: true });
  });
}

function renderCustomDatePicker(input, view) {
  const selectedValue = input.value;
  const todayValue = formatDateForInput(new Date());
  const firstDay = new Date(view.year, view.month, 1);
  const start = new Date(view.year, view.month, 1 - firstDay.getDay());
  const cells = [];
  for (let index = 0; index < 42; index += 1) {
    const date = new Date(start);
    date.setDate(start.getDate() + index);
    const value = formatDateForInput(date);
    const outside = date.getMonth() !== view.month;
    cells.push(`
      <button type="button" class="custom-date-day ${outside ? "outside" : ""} ${value === selectedValue ? "selected" : ""} ${value === todayValue ? "today" : ""}" data-date-value="${value}">
        ${date.getDate()}
      </button>
    `);
  }
  return `
    <div class="custom-date-header">
      <button type="button" data-date-prev aria-label="Previous month">‹</button>
      <strong>${CUSTOM_DATE_MONTHS[view.month]} ${view.year}</strong>
      <button type="button" data-date-next aria-label="Next month">›</button>
    </div>
    <div class="custom-date-weekdays">
      ${CUSTOM_DATE_WEEKDAYS.map((day) => `<span>${day}</span>`).join("")}
    </div>
    <div class="custom-date-grid">${cells.join("")}</div>
    <div class="custom-date-footer">
      <button type="button" data-date-clear>Clear</button>
      <button type="button" data-date-today>Today</button>
    </div>
  `;
}

function positionCustomDatePicker(input, popover) {
  const rect = input.getBoundingClientRect();
  const width = 360;
  const left = Math.min(Math.max(12, rect.left + window.scrollX), window.scrollX + window.innerWidth - width - 12);
  popover.style.left = `${left}px`;
  popover.style.top = `${rect.bottom + window.scrollY + 10}px`;
}

function handleDatePickerOutsideClick(event) {
  if (!activeDatePicker) return;
  if (activeDatePicker.popover.contains(event.target) || activeDatePicker.input.contains(event.target)) return;
  closeCustomDatePicker();
}

function closeCustomDatePicker() {
  if (!activeDatePicker) return;
  document.removeEventListener("mousedown", handleDatePickerOutsideClick);
  activeDatePicker.popover.remove();
  activeDatePicker = null;
}

function setDatePickerInputValue(input, value) {
  input.value = value;
  input.dispatchEvent(new Event("input", { bubbles: true }));
  input.dispatchEvent(new Event("change", { bubbles: true }));
}

function parseDateValue(value) {
  const [year, month, day] = String(value || "").split("-").map(Number);
  if (!year || !month || !day) return null;
  const date = new Date(year, month - 1, day);
  return Number.isNaN(date.getTime()) ? null : date;
}

function formatDateForInput(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function weekdayFromDate(value) {
  if (!value) return "No date selected";
  const date = new Date(`${value}T12:00:00`);
  return date.toLocaleDateString("en-GB", { weekday: "long" });
}

function formatDateName(value) {
  if (!value) return "";
  const [year, month, day] = value.split("-");
  return `${day}-${month}-${year}`;
}

function formatNumber(value) {
  return new Intl.NumberFormat("en-GB", { maximumFractionDigits: 2 }).format(Number(value || 0));
}

function clampNumber(value, min, max) {
  const number = Number(String(value).replace(",", "."));
  if (!Number.isFinite(number)) return min;
  return Math.min(Math.max(number, min), max);
}

function cleanNumber(value) {
  return Math.abs(value - Math.round(value)) < 0.0001 ? Math.round(value) : Number(value.toFixed(2));
}

function roundPercent(value) {
  return Number(value.toFixed(2));
}

function escapeHtml(value) {
  return repairDisplayText(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function normalizeDisplayObject(value) {
  if (typeof value === "string") return repairDisplayText(value);
  if (Array.isArray(value)) return value.map((entry) => normalizeDisplayObject(entry));
  if (value && typeof value === "object") {
    return Object.fromEntries(Object.entries(value).map(([key, entry]) => [key, normalizeDisplayObject(entry)]));
  }
  return value;
}

function normalizeText(value) {
  return repairDisplayText(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function repairDisplayText(value) {
  return String(value ?? "")
    .replaceAll("K��ln", "Köln")
    .replaceAll("KÃ¶ln", "Köln")
    .replaceAll("Gro��beeren", "Großbeeren")
    .replaceAll("GroÃŸbeeren", "Großbeeren")
    .replaceAll("K��ge", "Køge")
    .replaceAll("KÃ¸ge", "Køge")
    .replaceAll("G��nzburg", "Günzburg")
    .replaceAll("GÃ¼nzburg", "Günzburg")
    .replaceAll("H??nchen", "Hähnchen")
    .replaceAll("H??hnchen", "Hähnchen")
    .replaceAll("H��nchen", "Hähnchen")
    .replaceAll("H��hnchen", "Hähnchen")
    .replaceAll("Hï¿½ï¿½nchen", "Hähnchen")
    .replaceAll("HÃ¤hnchen", "Hähnchen")
    .replaceAll("K??se", "Käse")
    .replaceAll("K��se", "Käse")
    .replaceAll("Kï¿½ï¿½se", "Käse")
    .replaceAll("KÃ¤se", "Käse")
    .replaceAll("Ziegenk��se", "Ziegenkäse")
    .replaceAll("ZiegenkÃ¤se", "Ziegenkäse")
    .replaceAll("M��ckm��hl", "Möckmühl")
    .replaceAll("MÃ¶ckmÃ¼hl", "Möckmühl")
    .replaceAll("Â·", "-");
}

