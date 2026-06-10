const state = {
  preview: null,
  sessions: [],
  daySessions: [],
  activeDaySession: null,
  activeSessionId: null,
  editMode: "new",
  shortagesOnly: false,
  currentPage: "sessions",
};

const API_PREFIX = window.location.pathname.startsWith("/shortage") ? "/shortage/api" : "/api";

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
const shortageOnlyToggle = document.getElementById("shortageOnlyToggle");
const historyDate = document.getElementById("historyDate");
const historyDateList = document.getElementById("historyDateList");
const historyList = document.getElementById("historyList");
const analyticsButton = document.getElementById("analyticsButton");
const analyticsPage = document.getElementById("analyticsPage");
const analyticsTitle = document.getElementById("analyticsTitle");
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
historyDate.addEventListener("input", renderHistory);
itemsBody.addEventListener("input", handleShortageInput);
shortageOnlyToggle.addEventListener("change", () => {
  state.shortagesOnly = shortageOnlyToggle.checked;
  renderPreview();
});
analyticsButton.addEventListener("click", () => showPage("analytics"));
backFromAnalyticsButton.addEventListener("click", () => showPage("history"));
sessionsNav.addEventListener("click", () => showPage("sessions"));
ordersNav.addEventListener("click", () => showPage("workspace"));
historyNav.addEventListener("click", () => showPage("history"));
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

updateDaySessionWeekday();
renderPreview();

async function handleGreenOpsEmailUpload() {
  if (!state.activeDaySession) {
    statusText.textContent = "Open or create a daily session before uploading a GreenOps email order.";
    showPage("sessions");
    return;
  }
  const file = emailInput.files[0];
  if (!file) return;

  statusText.textContent = "Reading GreenOps email and saving it to the active shortage session...";
  saveSessionButton.disabled = true;

  const formData = new FormData();
  formData.append("email", file);
  formData.append("date", state.activeDaySession.date || "");
  formData.append("name", state.activeDaySession.name || "");

  const response = await fetch(apiPath("/orders/ingest"), { method: "POST", body: formData });
  const payload = await response.json();
  if (!response.ok) {
    state.preview = null;
    renderPreview();
    statusText.textContent = payload.error || "Could not read the GreenOps email order.";
    return;
  }

  state.sessions = payload.sessions || state.sessions;
  const savedSessions = payload.savedSessions || [];
  if (savedSessions.length) {
    const firstSession = savedSessions[0];
    state.activeSessionId = firstSession.id;
    state.editMode = "existing";
    state.activeDaySession = {
      id: `saved-${firstSession.date}`,
      date: firstSession.date,
      weekday: firstSession.weekday,
      name: firstSession.name,
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
  statusText.textContent = `${file.name} saved to ${state.activeDaySession.name}. ${savedSessions.length} order${savedSessions.length === 1 ? "" : "s"} ready for shortage control.`;
}

async function handleUpload() {
  if (!state.activeDaySession) {
    statusText.textContent = "Open or create a daily session before uploading an order file.";
    showPage("sessions");
    return;
  }
  const file = excelInput.files[0];
  if (!file) return;

  statusText.textContent = "Reading Excel file and calculating shortages...";
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

  state.preview = normalizeEditablePreview(payload.preview);
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
  const response = await fetch(apiPath("/sessions"), {
    method: isUpdate ? "PUT" : "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      id: state.activeSessionId,
      date: sessionDate.value,
      name: sessionName.value,
      preview: state.preview,
    }),
  });
  const payload = await response.json();
  if (!response.ok) {
    statusText.textContent = payload.error || "Could not save the session.";
    return;
  }

  state.sessions = payload.sessions;
  state.activeSessionId = payload.session?.id || state.activeSessionId;
  statusText.textContent = isUpdate ? "Delivery point updated." : "Session saved.";
  saveSessionButton.textContent = "Update delivery point";
  saveSessionButton.disabled = false;
  quickUpdateButton.disabled = false;
  renderHistory();
}

async function loadInitialData() {
  await Promise.all([loadDaySessions(), loadSessions()]);
  showPage("sessions");
}

async function syncSharedData() {
  if (state.currentPage === "workspace" && state.preview) return;
  await Promise.all([loadDaySessions(), loadSessions()]);
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
  state.sessions = payload.sessions || [];
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
  sessionsNav.classList.toggle("active", showSessions);
  ordersNav.classList.toggle("active", showWorkspace);
  historyNav.classList.toggle("active", showHistory || showAnalytics);
  renderWorkspaceMode();
  if (showHistory) renderHistory();
  if (showAnalytics) renderAnalytics();
  if (showSessions) renderDaySessions();
}

function renderPreview() {
  const preview = state.preview;
  if (!preview) {
    shortageOnlyToggle.checked = false;
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
    itemsBody.innerHTML = '<tr><td colspan="7" class="empty-cell">Upload a file to preview shortage lines.</td></tr>';
    return;
  }

  clientTitle.textContent = preview.client;
  shortageOnlyToggle.disabled = false;
  shortageOnlyToggle.checked = state.shortagesOnly;
  confidenceBadge.textContent = `${preview.confidence}%`;
  confidenceBadge.classList.toggle("warning", preview.confidence < 50);
  deliveryPoint.textContent = preview.deliveryPoint;
  orderedTotal.textContent = formatNumber(preview.orderedTotal);
  shortageTotal.textContent = formatNumber(preview.shortageTotal);
  shortageRate.textContent = `${preview.shortagePercentage}%`;
  saveSessionButton.disabled = false;
  quickUpdateButton.disabled = false;
  exportMancosButton.disabled = !hasShortageLines();
  renderWorkspaceMode();

  const visibleItems = state.shortagesOnly
    ? preview.items.filter((item) => Number(item.shortageQuantity || 0) > 0)
    : preview.items;

  if (!visibleItems.length) {
    itemsBody.innerHTML = `<tr><td colspan="7" class="empty-cell">${
      state.shortagesOnly ? "No articles with shortages found." : "No order lines with shortage data were found."
    }</td></tr>`;
    return;
  }

  itemsBody.innerHTML = visibleItems
    .map(
      (item) => `
        <tr>
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
              data-index="${preview.items.indexOf(item)}"
              aria-label="Shortage for ${escapeHtml(item.article)}"
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
    row.children[3].textContent = formatNumber(item.deliveredQuantity);
    row.children[5].textContent = `${item.shortagePercentage}%`;
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
  const sessions = filterDate ? state.sessions.filter((session) => session.date === filterDate) : state.sessions;
  if (filterDate && historyDate.value !== filterDate) {
    historyDate.value = filterDate;
  }

  if (!sessions.length) {
    historyList.innerHTML = '<p class="empty-cell">No saved sessions found for this date.</p>';
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
}

function renderAnalytics() {
  const filterDate = historyDate.value || mostRecentSessionDate();
  const sessions = filterDate ? state.sessions.filter((session) => session.date === filterDate) : state.sessions;
  analyticsTitle.textContent = filterDate ? `Session analytics · ${filterDate}` : "Session analytics";

  if (!sessions.length) {
    clientOverviewBody.innerHTML = '<tr><td colspan="5" class="empty-cell">No sessions found for this date.</td></tr>';
    clientRateBars.innerHTML = '<p class="empty-cell">No data to chart.</p>';
    topArticleBars.innerHTML = '<p class="empty-cell">No article shortages found.</p>';
    topArticlesBody.innerHTML = '<tr><td colspan="4" class="empty-cell">No article shortages found.</td></tr>';
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

function buildClientAnalytics(sessions) {
  const clients = new Map();
  for (const session of sessions) {
    const preview = session.preview || {};
    const client = preview.client || "Unknown client";
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
    const preview = session.preview || {};
    const client = preview.client || "Unknown client";
    for (const item of preview.items || []) {
      const shortage = Number(item.shortageQuantity || 0);
      if (shortage <= 0) continue;
      const key = `${item.article}::${item.description || ""}`;
      const current = articles.get(key) || {
        article: item.article,
        description: item.description || "",
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
    topArticleBars.innerHTML = '<p class="empty-cell">No article shortages found.</p>';
    topArticlesBody.innerHTML = '<tr><td colspan="4" class="empty-cell">No article shortages found.</td></tr>';
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
    daySessionList.innerHTML = '<p class="empty-cell">No daily sessions created yet.</p>';
    return;
  }
  daySessionList.innerHTML = daySessions
    .map((session) => {
      const orderCount = state.sessions.filter((entry) => entry.date === session.date).length;
      return `
        <article class="day-session-card">
          <div>
            <h3>${escapeHtml(session.name || "Daily session")}</h3>
            <p>${escapeHtml(session.weekday || "")}, ${escapeHtml(session.date || "")}</p>
          </div>
          <div class="history-stat">
            <span>Orders</span>
            <strong>${orderCount}</strong>
          </div>
          <div class="day-session-actions">
            <button class="open-session-button" type="button" data-day-session-id="${escapeHtml(session.id)}">Open session</button>
            <button class="delete-session-button" type="button" data-delete-day-session-id="${escapeHtml(session.id)}">Delete</button>
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
  state.sessions = payload.sessions || [];
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
  statusText.textContent = `${session.name} opened. Upload a GreenOps email order or shortage Excel file.`;
  showPage("workspace");
}

function visibleDaySessions() {
  const byDate = new Map();
  for (const session of state.sessions) {
    if (!session.date || byDate.has(session.date)) continue;
    byDate.set(session.date, {
      id: `saved-${session.date}`,
      date: session.date,
      weekday: session.weekday || weekdayFromDate(session.date),
      name: session.name || `${weekdayFromDate(session.date)} shortage review`,
    });
  }
  for (const session of state.daySessions) {
    if (!session.date) continue;
    byDate.set(session.date, session);
  }
  return Array.from(byDate.values()).sort((a, b) => String(b.date || "").localeCompare(String(a.date || "")));
}

function applyActiveDaySession() {
  if (!state.activeDaySession) return;
  sessionDate.value = state.activeDaySession.date || "";
  sessionName.value = state.activeDaySession.name || "";
  activeSessionTitle.textContent = state.activeDaySession.name || "Active session";
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
    const client = session.preview?.client || "Unknown client";
    if (!groups[client]) groups[client] = [];
    groups[client].push(session);
    return groups;
  }, {});
}

function renderDeliveryPointSession(session) {
  const preview = session.preview || {};
  return `
    <article class="delivery-point-card">
      <div>
        <h4>${escapeHtml(preview.deliveryPoint || "Unknown delivery point")}</h4>
        <p>${escapeHtml(session.weekday || "")}, ${escapeHtml(session.date || "")} · ${escapeHtml(session.name || "Shortage session")}</p>
      </div>
      <div class="history-stat">
        <span>Ordered</span>
        <strong>${formatNumber(preview.orderedTotal || 0)}</strong>
      </div>
      <div class="history-stat">
        <span>Shortage</span>
        <strong>${formatNumber(preview.shortageTotal || 0)}</strong>
      </div>
      <div class="history-stat">
        <span>Rate</span>
        <strong>${preview.shortagePercentage || 0}%</strong>
      </div>
      <button class="open-session-button" type="button" data-session-id="${escapeHtml(session.id)}">Open list</button>
    </article>
  `;
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
    id: `saved-${session.date}`,
    date: session.date,
    weekday: session.weekday,
    name: session.name,
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
    statusText.textContent = "There are no shortage lines to export.";
    return;
  }
  exportMancosButton.disabled = true;
  const response = await fetch(apiPath("/export-mancos"), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ preview: state.preview }),
  });

  if (!response.ok) {
    const payload = await response.json();
    statusText.textContent = payload.error || "Could not export Manco's.";
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

function clonePreview(preview) {
  return JSON.parse(JSON.stringify(preview));
}

function setDefaultSessionName() {
  const weekday = weekdayFromDate(sessionDate.value);
  if (!sessionName.value || /shortage review$/i.test(sessionName.value) || /^\w+ - \d{2}-\d{2}-\d{4}$/i.test(sessionName.value)) {
    sessionName.value = `${weekday} - ${formatDateName(sessionDate.value)}`;
  }
}

function setDefaultDaySessionName() {
  const weekday = weekdayFromDate(daySessionDate.value);
  if (!daySessionName.value || /shortage review$/i.test(daySessionName.value) || /^\w+ - \d{2}-\d{2}-\d{4}$/i.test(daySessionName.value)) {
    daySessionName.value = `${weekday} - ${formatDateName(daySessionDate.value)}`;
  }
}

function updateWeekday() {
  selectedWeekday.textContent = weekdayFromDate(sessionDate.value);
}

function updateDaySessionWeekday() {
  daySessionWeekday.textContent = weekdayFromDate(daySessionDate.value);
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
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
