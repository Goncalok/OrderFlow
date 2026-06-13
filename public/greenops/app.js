const LEVERSCHEMA_STORAGE_KEY = "email-order-leverschema-results";
const LAADSCHEMA_STORAGE_KEY = "orderflow-laadschema-data";
const WORK_SESSION_STORAGE_KEY = "orderflow-work-sessions";
const ACTIVE_WORK_SESSION_STORAGE_KEY = "orderflow-active-work-session-id";
const CLIENT_ACTION_SETTINGS_KEY = "orderflow-client-action-settings";
const HAVI_UIEN_SETTINGS_KEY = "orderflow-havi-uien-settings";
const LAADSCHEMA_DAYS = ["Zondag", "Maandag", "Dinsdag", "Woensdag", "Donderdag", "Vrijdag", "Zaterdag"];
const LAADSCHEMA_ISO_DAY_OFFSETS = {
  Maandag: 0,
  Dinsdag: 1,
  Woensdag: 2,
  Donderdag: 3,
  Vrijdag: 4,
  Zaterdag: 5,
  Zondag: 6,
};
const DASHBOARD_CLIENTS = [
  "Carrefour",
  "Colruyt",
  "Colruyt Saturday",
  "Denemark",
  "Edeka",
  "Globus",
  "Hanos",
  "HAVI",
  "Havi Duisburg Saturday",
  "HelloFresh",
  "Heeren",
  "NettoMD",
  "Rewe",
  "Penny",
];

const CLIENT_WORKSPACE_ALIASES = {
  Carrefour: ["Carrefour FIF", "Carrefour KDC"],
  Edeka: ["Edeka Laatzen", "Edeka Mochmuhl"],
  HAVI: [
    "Havi NL",
    "Havi BE",
    "Havi UK",
    "Havi DE",
    "Havi\nDC Duisburg\nDC Wunstorf\nDC Neu Wulmstorf",
    "Havi DE Saturday",
  ],
};

const ACTION_VISIBILITY_OPTIONS = [
  { key: "export", label: "Export to Excel", description: "Standard single-order Excel export." },
  { key: "merge", label: "Merge and Export", description: "Merge repeated customer orders into one file." },
  { key: "special", label: "Export HAVI DE UIEN", description: "Special export flow for Havi DC (Onions)." },
  { key: "nettoMd", label: "NettoMD Orderpicking", description: "Special orderpicking export for NettoMD." },
  { key: "leverschema", label: "Leverschema", description: "Show sheet selection, save PP result, and Leverschema summary." },
  { key: "printCmr", label: "Print CMR", description: "Export CMR document for Carrefour, Colruyt, Colruyt Saturday, Denemark, Edeka, Globus, HAVI, and Heeren." },
];

const state = {
  user: null,
  file: null,
  fileRestoreToken: 0,
  mode: null,
  preview: null,
  clientWorkspaces: {},
  selectedClient: DASHBOARD_CLIENTS[0],
  selectedIndex: 0,
  orderSelectionActive: false,
  activeDeliveryPointKey: null,
  leverschemaIncludedIndexes: [],
  leverschemaResults: {},
  laadschemaData: {},
  laadschemaCustomTrucks: {},
  teamSessions: [],
  clientActionSettings: loadClientActionSettings(),
  haviUienSettings: loadHaviUienSettings(),
  activeWorkSession: null,
  stockItems: [],
  currentPage: "dashboard",
  currentLeverschemaSheet: "Monday-Thursday",
  settingsClient: DASHBOARD_CLIENTS[0],
};

const MASTER_SHEETS = {
  "Monday-Thursday": [
    { key: "havi_duisburg", klant: "Havi Duisburg", dateOffset: 1, time: "15:30", carrier: "Hendrikx/SVZ", note: "prognoose" },
    { key: "carrefour_fif", klant: "Carrefour FIF", dateOffset: 0, time: "11:00", carrier: "Van Tilburg", note: "prognoose" },
    { key: "colruyt", klant: "Colruyt", dateOffset: 2, time: "01:30", carrier: "Hendrikx", note: "prognoose" },
    { key: "heeren", klant: "Heeren", dateOffset: 2, time: "05:15", carrier: "SVZ", note: "prognoose" },
    { key: "carrefour_kdc", klant: "Carrefour KDC", dateOffset: 1, time: "03:15", carrier: "Van Tilburg", note: "prognoose" },
  ],
  Friday: [
    { key: "havi_duisburg", klant: "Havi Duisburg", dateOffset: 2, time: "15:30", carrier: "Hendrikx/SVZ", note: "prognoose" },
    { key: "havi_duisburg_saturday", klant: "Havi Duisburg", dateOffset: 3, time: "15:30", carrier: "Hendrikx/SVZ", note: "prognoose" },
    { key: "carrefour_fif", klant: "Carrefour FIF", dateOffset: 0, time: "11:00", carrier: "Van Tilburg", note: "prognoose" },
    { key: "colruyt", klant: "Colruyt", dateOffset: 3, time: "01:30", carrier: "Hendrikx", note: "prognoose" },
    { key: "colruyt_saturday", klant: "Colruyt", dateOffset: 4, time: "01:30", carrier: "Hendrikx", note: "prognoose" },
    { key: "heeren", klant: "Heeren", dateOffset: 3, time: "05:15", carrier: "SVZ", note: "prognoose" },
    { key: "carrefour_kdc", klant: "Carrefour KDC", dateOffset: 1, time: "03:15", carrier: "Van Tilburg", note: "prognoose" },
  ],
  Saturday: [
    { key: "havi_duisburg", klant: "Havi Duisburg", dateOffset: 2, time: "15:30", carrier: "Hendrikx/SVZ", note: "prognoose" },
    { key: "carrefour_fif", klant: "Carrefour FIF", dateOffset: 0, time: "11:00", carrier: "Van Tilburg", note: "prognoose" },
    { key: "colruyt", klant: "Colruyt", dateOffset: 3, time: "01:30", carrier: "Hendrikx", note: "prognoose" },
    { key: "heeren", klant: "Heeren", dateOffset: 3, time: "05:15", carrier: "SVZ", note: "prognoose" },
    { key: "carrefour_kdc", klant: "Carrefour KDC", dateOffset: 2, time: "03:15", carrier: "Van Tilburg", note: "prognoose" },
  ],
};

const CARREFOUR_BOX_TYPES = {
  296158: "M",
  296159: "M",
  296157: "M",
  296156: "M",
  296178: "M",
  296182: "M",
  296106: "M",
  296179: "M",
  296147: "M",
  296181: "M",
  296180: "M",
  231023: "M",
  231022: "M",
  296195: "M",
  296196: "M",
  296197: "M",
};

const HAVI_NETHERLANDS_FULL_PALLET = {
  297071: { collo: 108, stuk: 864 },
  297045: { collo: 64, stuk: 320 },
  297072: { collo: 108, stuk: 864 },
  297068: { collo: 160, stuk: 1600 },
  297075: { collo: 40, stuk: 800 },
  297090: { collo: 200, stuk: 800 },
  297092: { collo: 200, stuk: 800 },
  297073: { collo: 56, stuk: 448 },
  297048: { collo: 64, stuk: 384 },
  297083: { collo: 600, stuk: 600 },
};

const HAVI_BELGIUM_FULL_PALLET = {
  297057: { collo: 64, stuk: 384 },
  297045: { collo: 64, stuk: 320 },
  297072: { collo: 108, stuk: 864 },
  297068: { collo: 160, stuk: 1600 },
  297076: { collo: 40, stuk: 800 },
  297077: { collo: 40, stuk: 800 },
  297092: { collo: 200, stuk: 800 },
  297073: { collo: 56, stuk: 448 },
  297021: { collo: 200, stuk: 800 },
};

const appLogo = document.getElementById("appLogo");
const authScreen = document.getElementById("authScreen");
const appShell = document.getElementById("appShell");
const appHero = document.querySelector(".hero");
const loginForm = document.getElementById("loginForm");
const loginEmail = document.getElementById("loginEmail");
const loginPassword = document.getElementById("loginPassword");
const loginButton = document.getElementById("loginButton");
const loginError = document.getElementById("loginError");
const sessionUserName = document.getElementById("sessionUserName");
const sessionUserEmail = document.getElementById("sessionUserEmail");
const headerSessionName = document.getElementById("headerSessionName");
const headerSessionDate = document.getElementById("headerSessionDate");
const settingsButton = document.getElementById("settingsButton");
const logoutButton = document.getElementById("logoutButton");
const emailInput = document.getElementById("emailInput");
const exportButton = document.getElementById("exportButton");
const mergeButton = document.getElementById("mergeButton");
const saveOrderButton = document.getElementById("saveOrderButton");
const addItemButton = document.getElementById("addItemButton");
const specialButton = document.getElementById("specialButton");
const nettoMdButton = document.getElementById("nettoMdButton");
const leverschemaButton = document.getElementById("leverschemaButton");
const printCmrButton = document.getElementById("printCmrButton");
const haviDcCmrButtons = document.getElementById("haviDcCmrButtons");
const printCmrDuisburgButton = document.getElementById("printCmrDuisburgButton");
const printCmrWunstorfButton = document.getElementById("printCmrWunstorfButton");
const printCmrNeuWulmstorfButton = document.getElementById("printCmrNeuWulmstorfButton");
const nettomdCmrButtons = document.getElementById("nettomdCmrButtons");
const nettomdCmrToggleButton = document.getElementById("nettomdCmrToggleButton");
const nettomdCmrMenu = document.getElementById("nettomdCmrMenu");
const cmrWithPakbonButton = document.getElementById("cmrWithPakbonButton");
const cmrWithoutPakbonButton = document.getElementById("cmrWithoutPakbonButton");
const printCmrKerpenButton = document.getElementById("printCmrKerpenButton");
const printCmrHodenhagenButton = document.getElementById("printCmrHodenhagenButton");
const printCmrHenstedtButton = document.getElementById("printCmrHenstedtButton");
const printCmrHammButton = document.getElementById("printCmrHammButton");
const printCmrGanderkeseeButton = document.getElementById("printCmrGanderkeseeButton");
const printCmrBottropButton = document.getElementById("printCmrBottropButton");
const printCmrKrefeldButton = document.getElementById("printCmrKrefeldButton");
const exportSheetSelect = document.getElementById("exportSheetSelect");
const orderSelect = document.getElementById("orderSelect");
const deliveryPointList = document.getElementById("deliveryPointList");
const deliveryReferenceList = document.getElementById("deliveryReferenceList");
const customersFound = document.getElementById("customersFound");
const itemsInSelection = document.getElementById("itemsInSelection");
const deliveryDate = document.getElementById("deliveryDate");
const detailCustomer = document.getElementById("detailCustomer");
const detailReference = document.getElementById("detailReference");
const detailFatrans = document.getElementById("detailFatrans");
const tableHead = document.getElementById("tableHead");
const tableBody = document.getElementById("tableBody");
const statusText = document.getElementById("statusText");
const dashboardPage = document.getElementById("dashboardPage");
const clientsPage = document.getElementById("clientsPage");
const settingsPage = document.getElementById("settingsPage");
const stockPage = document.getElementById("stockPage");
const clientCards = document.getElementById("clientCards");
const openLeverschemaFromClientsButton = document.getElementById("openLeverschemaFromClientsButton");
const exportLeverschemaFromClientsButton = document.getElementById("exportLeverschemaFromClientsButton");
const openLaadschemaButton = document.getElementById("openLaadschemaButton");
const openStockButton = document.getElementById("openStockButton");
const backToClientsFromStockButton = document.getElementById("backToClientsFromStockButton");
const stockFileInput = document.getElementById("stockFileInput");
const exportStockButton = document.getElementById("exportStockButton");
const stockStatusText = document.getElementById("stockStatusText");
const stockRowsCount = document.getElementById("stockRowsCount");
const stockTableBody = document.getElementById("stockTableBody");
const backToClientsFromLaadschemaButton = document.getElementById("backToClientsFromLaadschemaButton");
const exportLaadschemaButton = document.getElementById("exportLaadschemaButton");
const addTruckButton = document.getElementById("addTruckButton");
const clearLaadschemaButton = document.getElementById("clearLaadschemaButton");
const clearAllLaadschemaButton = document.getElementById("clearAllLaadschemaButton");
const laadschemaDaySelect = document.getElementById("laadschemaDaySelect");
const laadschemaDateInput = document.getElementById("laadschemaDateInput");
const laadschemaWeekInput = document.getElementById("laadschemaWeekInput");
const laadschemaTable = document.getElementById("laadschemaTable");
const laadschemaTableHead = document.getElementById("laadschemaTableHead");
const laadschemaTableBody = document.getElementById("laadschemaTableBody");
const settingsClientSelect = document.getElementById("settingsClientSelect");
const settingsActionList = document.getElementById("settingsActionList");
const haviUienSettingsCard = document.getElementById("haviUienSettingsCard");
const haviUienArticleInput = document.getElementById("haviUienArticleInput");
const haviUienDescriptionInput = document.getElementById("haviUienDescriptionInput");
const backFromSettingsButton = document.getElementById("backFromSettingsButton");
const ordersClientTitle = document.getElementById("ordersClientTitle");
const backToClientsButton = document.getElementById("backToClientsButton");
const workSessionForm = document.getElementById("workSessionForm");
const workSessionDate = document.getElementById("workSessionDate");
const workSessionName = document.getElementById("workSessionName");
const createWorkSessionButton = document.getElementById("createWorkSessionButton");
const sessionHistoryList = document.getElementById("sessionHistoryList");
const ordersPage = document.getElementById("ordersPage");
const leverschemaPage = document.getElementById("leverschemaPage");
const laadschemaPage = document.getElementById("laadschemaPage");
const backFromLeverschemaButton = document.getElementById("backFromLeverschemaButton");
const leverschemaTableTitle = document.getElementById("leverschemaTableTitle");
const leverschemaTableNote = document.getElementById("leverschemaTableNote");
const leverschemaMasterHead = document.getElementById("leverschemaMasterHead");
const leverschemaMasterBody = document.getElementById("leverschemaMasterBody");
const exportLeverschemaButton = document.getElementById("exportLeverschemaButton");
const confirmationModal = document.getElementById("confirmationModal");
const confirmationModalEyebrow = document.getElementById("confirmationModalEyebrow");
const confirmationModalTitle = document.getElementById("confirmationModalTitle");
const confirmationModalMessage = document.getElementById("confirmationModalMessage");
const confirmationCancelButton = document.getElementById("confirmationCancelButton");
const confirmationConfirmButton = document.getElementById("confirmationConfirmButton");
const clearLeverschemaButton = document.getElementById("clearLeverschemaButton");
const leverschemaSavedSheet = document.getElementById("leverschemaSavedSheet");
const leverschemaRule = document.getElementById("leverschemaRule");
const leverschemaTotal = document.getElementById("leverschemaTotal");
const leverschemaPlaces = document.getElementById("leverschemaPlaces");
const leverschemaNote = document.getElementById("leverschemaNote");
const leverschemaIncludeCard = document.getElementById("leverschemaIncludeCard");
const leverschemaSummaryCard = document.getElementById("leverschemaSummaryCard");
const leverschemaIncludeList = document.getElementById("leverschemaIncludeList");
const sheetTabButtons = Array.from(document.querySelectorAll("[data-sheet-tab]"));

loginForm.addEventListener("submit", handleLogin);
appLogo.addEventListener("click", () => switchPage("dashboard"));
settingsButton.addEventListener("click", openSettingsPage);
logoutButton.addEventListener("click", handleLogout);
workSessionForm.addEventListener("submit", handleWorkSessionCreate);
workSessionDate.addEventListener("change", syncWorkSessionNameFromDate);
emailInput.addEventListener("change", handleUpload);
tableBody.addEventListener("change", handleQuantityEdit);
tableBody.addEventListener("change", handleItemFieldEdit);
tableBody.addEventListener("click", handleDeleteItem);
leverschemaIncludeList.addEventListener("change", handleLeverschemaIncludeChange);
deliveryPointList?.addEventListener("click", handleDeliveryPointClick);
deliveryPointList?.addEventListener("change", handleDeliveryReferenceToggle);
deliveryPointList?.addEventListener("click", handleDeliveryOrderDelete);
deliveryPointList?.addEventListener("click", handleDeliveryReferenceClick);
deliveryReferenceList?.addEventListener("change", handleDeliveryReferenceToggle);
deliveryReferenceList?.addEventListener("click", handleDeliveryOrderDelete);
deliveryReferenceList?.addEventListener("click", handleDeliveryReferenceClick);
leverschemaMasterBody.addEventListener("click", handleLeverschemaRowClear);
orderSelect.addEventListener("change", () => {
  state.selectedIndex = Number(orderSelect.value || 0);
  state.orderSelectionActive = true;
  state.activeDeliveryPointKey = getDeliveryPointKeyForIndex(state.selectedIndex);
  syncLeverschemaIncludedIndexes();
  syncCurrentClientWorkspace();
  renderPreview();
  // Only update Havi DC buttons if printCmr is enabled
  const preferences = getClientActionPreferences(state.selectedClient);
  const isHavi3DC = isHavi3DCClient(normalizeText(state.selectedClient));
  if (isHavi3DC && preferences.printCmr) {
    updateHaviDCButtonsVisibility();
  } else if (isHavi3DC && !preferences.printCmr) {
    // Hide all individual Havi DC buttons when printCmr is disabled
    printCmrDuisburgButton.classList.add("hidden");
    printCmrWunstorfButton.classList.add("hidden");
    printCmrNeuWulmstorfButton.classList.add("hidden");
  }
});
exportButton.addEventListener("click", () => downloadExport("selected"));
mergeButton.addEventListener("click", () => downloadExport("merge"));
saveOrderButton?.addEventListener("click", handleSaveOrder);
addItemButton?.addEventListener("click", handleAddItem);
specialButton.addEventListener("click", () => downloadExport("special"));
nettoMdButton.addEventListener("click", () => downloadExport("netto_md"));
leverschemaButton.addEventListener("click", saveLeverschemaResult);
// printCmrButton handler is set at the end of the file to handle pakbon modal
printCmrDuisburgButton.addEventListener("click", () => downloadExport("print_cmr", "Duisburg"));
printCmrWunstorfButton.addEventListener("click", () => downloadExport("print_cmr", "Wunstorf"));
printCmrNeuWulmstorfButton.addEventListener("click", () => downloadExport("print_cmr", "Neu Wulmstorf"));
nettomdCmrToggleButton?.addEventListener("click", toggleNettoMdCmrMenu);
cmrWithPakbonButton?.addEventListener("click", handleCmrWithPakbonChoice);
cmrWithoutPakbonButton?.addEventListener("click", handleCmrWithoutPakbonChoice);
printCmrKerpenButton.addEventListener("click", () => downloadExport("print_cmr", "Kerpen"));
printCmrHodenhagenButton.addEventListener("click", () => downloadExport("print_cmr", "Hodenhagen"));
printCmrHenstedtButton.addEventListener("click", () => downloadExport("print_cmr", "Henstedt-Ulzburg"));
printCmrHammButton.addEventListener("click", () => downloadExport("print_cmr", "Hamm"));
printCmrGanderkeseeButton.addEventListener("click", () => downloadExport("print_cmr", "Ganderkesee"));
printCmrBottropButton.addEventListener("click", () => downloadExport("print_cmr", "Bottrop"));
printCmrKrefeldButton.addEventListener("click", () => downloadExport("print_cmr", "Krefeld"));
clearLeverschemaButton.addEventListener("click", clearLeverschemaMemory);
exportLeverschemaButton.addEventListener("click", exportLeverschemaWorkbook);
exportSheetSelect?.addEventListener("change", () => {
  syncCurrentClientWorkspace();
});
clientCards.addEventListener("click", handleClientCardClick);
openLeverschemaFromClientsButton.addEventListener("click", () => switchPage("leverschema"));
exportLeverschemaFromClientsButton.addEventListener("click", exportLeverschemaWorkbook);
openStockButton?.addEventListener("click", () => switchPage("stock"));
backToClientsFromStockButton?.addEventListener("click", () => switchPage("clients"));
stockFileInput?.addEventListener("change", handleStockUpload);
exportStockButton?.addEventListener("click", exportStockWorkbook);
openLaadschemaButton.addEventListener("click", () => {
  switchPage("laadschema");
  initializeLaadschema();
});
backToClientsFromLaadschemaButton.addEventListener("click", () => switchPage("clients"));
exportLaadschemaButton.addEventListener("click", exportLaadschemaData);
addTruckButton.addEventListener("click", openAddTruckModal);
clearLaadschemaButton.addEventListener("click", clearLaadschemaDay);
clearAllLaadschemaButton.addEventListener("click", clearAllLaadschemaData);
laadschemaDaySelect.addEventListener("change", updateLaadschemaDate);
laadschemaDateInput.addEventListener("change", updateLaadschemaWeek);
laadschemaWeekInput.addEventListener("change", updateLaadschemaFromWeek);
backToClientsButton.addEventListener("click", () => switchPage("clients"));
backFromSettingsButton.addEventListener("click", () => switchPage(state.activeWorkSession ? "clients" : "dashboard"));
backFromLeverschemaButton.addEventListener("click", () => switchPage("clients"));
settingsClientSelect.addEventListener("change", handleSettingsClientChange);
settingsActionList.addEventListener("change", handleSettingsActionChange);
haviUienArticleInput?.addEventListener("input", handleHaviUienSettingsInput);
haviUienDescriptionInput?.addEventListener("input", handleHaviUienSettingsInput);
window.addEventListener("focus", () => {
  if (state.user) refreshTeamState();
});
document.addEventListener("visibilitychange", () => {
  if (!document.hidden && state.user) refreshTeamState();
});
sheetTabButtons.forEach((button) => {
  button.addEventListener("click", () => {
    state.currentLeverschemaSheet = button.dataset.sheetTab;
    renderLeverschemaWorkbook();
  });
});

// IndexedDB: offline cache for the same team state synced via /api/work-sessions
function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open("OrderFlow_SessionsDB", 2);
    request.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains("sessions")) {
        db.createObjectStore("sessions", { keyPath: "id" });
      }
      if (!db.objectStoreNames.contains("workspaceEmails")) {
        db.createObjectStore("workspaceEmails", { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function buildWorkspaceEmailId(sessionId, client) {
  return `${sessionId || "no-session"}::${client || "no-client"}`;
}

function idbSaveWorkspaceEmail(id, file) {
  if (!id || !file) return Promise.resolve();
  return openDatabase().then((db) => {
    return new Promise((resolve, reject) => {
      const tx = db.transaction("workspaceEmails", "readwrite");
      const store = tx.objectStore("workspaceEmails");
      store.put({
        id,
        name: file.name || "email.eml",
        type: file.type || "message/rfc822",
        blob: file,
        updatedAt: new Date().toISOString(),
      });
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  });
}

function idbGetWorkspaceEmail(id) {
  if (!id) return Promise.resolve(null);
  return openDatabase().then((db) => {
    return new Promise((resolve, reject) => {
      const tx = db.transaction("workspaceEmails", "readonly");
      const store = tx.objectStore("workspaceEmails");
      const req = store.get(id);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
  });
}

function getTeamSessionsSorted() {
  return [...state.teamSessions].sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
}

function idbGetAllSessionsSorted() {
  return openDatabase().then((db) => {
    return new Promise((resolve, reject) => {
      const tx = db.transaction("sessions", "readonly");
      const store = tx.objectStore("sessions");
      const req = store.getAll();
      req.onsuccess = () => {
        const all = req.result || [];
        all.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
        resolve(all);
      };
      req.onerror = () => reject(req.error);
    });
  });
}

function mirrorSessionsToIndexedDB(sessions) {
  return openDatabase().then((db) => {
    return new Promise((resolve, reject) => {
      const tx = db.transaction("sessions", "readwrite");
      const store = tx.objectStore("sessions");
      store.clear();
      for (const s of sessions) {
        store.put(s);
      }
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  });
}

let teamStatePersistTimer = null;
let teamStateSyncInterval = null;
let deletedTeamSessionIds = [];
let workSessionCreateInFlight = false;

function cancelScheduledTeamStatePersist() {
  if (teamStatePersistTimer) {
    clearTimeout(teamStatePersistTimer);
    teamStatePersistTimer = null;
  }
}

function schedulePersistTeamState() {
  if (teamStatePersistTimer) {
    clearTimeout(teamStatePersistTimer);
  }
  teamStatePersistTimer = setTimeout(() => {
    teamStatePersistTimer = null;
    flushTeamStateToServer();
  }, 1500);
}

async function persistTeamStateImmediately() {
  cancelScheduledTeamStatePersist();
  await flushTeamStateToServer();
}

async function flushTeamStateToServer() {
  const body = JSON.stringify({
    sessions: state.teamSessions,
    leverschemaResults: state.leverschemaResults,
    laadschemaData: state.laadschemaData,
    laadschemaCustomTrucks: state.laadschemaCustomTrucks,
    deletedSessionIds: deletedTeamSessionIds,
  });
  
  console.log('📤 Sending team state to server...');
  console.log('Laadschema data size:', Object.keys(state.laadschemaData).length, 'days');
  console.log('Custom trucks size:', Object.keys(state.laadschemaCustomTrucks).length, 'days');
  
  try {
    const response = await fetch("/api/work_sessions", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      credentials: "include",
      body,
    });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }
    deletedTeamSessionIds = [];
    console.log('✅ Team state sent successfully');
  } catch (err) {
    console.error("Team state cloud sync failed:", err);
  }
  try {
    await mirrorSessionsToIndexedDB(state.teamSessions);
  } catch (err) {
    console.error("IndexedDB mirror failed:", err);
  }
  persistLeverschemaLocalCache();
}

function readLeverschemaFromLocalStorage() {
  try {
    const raw = window.localStorage.getItem(LEVERSCHEMA_STORAGE_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return typeof parsed === "object" && parsed !== null ? parsed : {};
  } catch {
    return {};
  }
}

function persistLeverschemaLocalCache() {
  try {
    window.localStorage.setItem(LEVERSCHEMA_STORAGE_KEY, JSON.stringify(state.leverschemaResults));
  } catch {
    // Ignore storage issues and keep in-memory state.
  }
}

function loadLaadschemaData() {
  try {
    const raw = window.localStorage.getItem(LAADSCHEMA_STORAGE_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return typeof parsed === "object" && parsed !== null ? parsed : {};
  } catch {
    return {};
  }
}

function loadLaadschemaCustomTrucks() {
  try {
    const raw = window.localStorage.getItem('laadschemaCustomTrucks');
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return typeof parsed === "object" && parsed !== null ? parsed : {};
  } catch {
    return {};
  }
}

function persistLaadschemaData() {
  try {
    window.localStorage.setItem(LAADSCHEMA_STORAGE_KEY, JSON.stringify(state.laadschemaData));
    // Also sync to server
    schedulePersistTeamState();
  } catch {
    // Ignore storage issues and keep in-memory state.
  }
}

async function tryMigrateIndexedDBToCloud() {
  if (state.teamSessions.length > 0) return;
  const local = await idbGetAllSessionsSorted();
  const lev = readLeverschemaFromLocalStorage();
  if (local.length === 0 && Object.keys(lev).length === 0) return;
  state.teamSessions = local;
  state.leverschemaResults = { ...lev, ...state.leverschemaResults };
  cancelScheduledTeamStatePersist();
  await flushTeamStateToServer();
}

async function refreshTeamState() {
  if (teamStatePersistTimer) {
    console.log("Skipping team refresh while local changes are pending.");
    return;
  }
  try {
    const response = await fetch("/api/work_sessions", { credentials: "include", cache: "no-store" });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }
    const data = await response.json();
    
    // Store old state for comparison
    const oldSessionsStr = JSON.stringify(state.teamSessions);
    const oldLeverschemaStr = JSON.stringify(state.leverschemaResults);
    const oldLaadschemaStr = JSON.stringify(state.laadschemaData);
    const oldCustomTrucksStr = JSON.stringify(state.laadschemaCustomTrucks);
    
    state.teamSessions = Array.isArray(data.sessions) ? data.sessions : [];
    state.leverschemaResults =
      data.leverschemaResults && typeof data.leverschemaResults === "object" ? data.leverschemaResults : {};
    
    // Load Laadschema data from server
    if (data.laadschemaData && typeof data.laadschemaData === "object") {
      state.laadschemaData = data.laadschemaData;
    }
    if (data.laadschemaCustomTrucks && typeof data.laadschemaCustomTrucks === "object") {
      state.laadschemaCustomTrucks = data.laadschemaCustomTrucks;
    }
    
    // Check if there are changes
    const sessionsChanged = oldSessionsStr !== JSON.stringify(state.teamSessions);
    const leverschemaChanged = oldLeverschemaStr !== JSON.stringify(state.leverschemaResults);
    const laadschemaChanged = oldLaadschemaStr !== JSON.stringify(state.laadschemaData);
    const customTrucksChanged = oldCustomTrucksStr !== JSON.stringify(state.laadschemaCustomTrucks);
    
    if (sessionsChanged) {
      renderSessionHistoryList();
      updateHeaderAccountInfo();
    }

    // If Laadschema data changed, re-render the table
    if (laadschemaChanged || customTrucksChanged) {
      console.log('🔄 Syncing Laadschema changes from server...');
      console.log('Laadschema data:', state.laadschemaData);
      console.log('Custom trucks:', state.laadschemaCustomTrucks);
      console.log('Current page:', state.currentPage);
      
      persistLaadschemaData(); // Save to localStorage
      localStorage.setItem('laadschemaCustomTrucks', JSON.stringify(state.laadschemaCustomTrucks));
      
      // Re-render Laadschema table if on that page
      if (state.currentPage === 'laadschema') {
        console.log('Re-rendering Laadschema table...');
        renderLaadschemaTable();
        console.log('✅ Laadschema synchronized successfully');
      } else {
        console.log('✅ Laadschema data synchronized (will render when page is opened)');
      }
    }
    
    // If there are changes and we have an active session, reload the workspace
    if ((sessionsChanged || leverschemaChanged) && state.activeWorkSession) {
      const updatedSession = state.teamSessions.find(s => s.id === state.activeWorkSession.id);
      if (updatedSession) {
        console.log('🔄 Syncing Leverschema changes from server...');
        console.log('Updated session:', updatedSession);
        
        // Update active session reference
        state.activeWorkSession = updatedSession;
        
        // Update client workspaces from the synced session
        state.clientWorkspaces = migrateClientWorkspaces(updatedSession.workspaces || {});
        state.activeWorkSession.workspaces = state.clientWorkspaces;
        
        // Reload current client workspace to reflect changes
        const currentWorkspace = state.clientWorkspaces[canonicalClientName(state.selectedClient)];
        if (currentWorkspace) {
          console.log('Loading workspace for client:', state.selectedClient);
          console.log('Workspace data:', currentWorkspace);
          
          // Update state from workspace
          state.mode = currentWorkspace.mode;
          state.preview = currentWorkspace.preview;
          state.selectedIndex = currentWorkspace.selectedIndex || 0;
          state.leverschemaIncludedIndexes = [...(currentWorkspace.leverschemaIncludedIndexes || [])];
          
          // Re-render UI
          renderPreview();
          renderLeverschemaSummary();
          renderLeverschemaWorkbook();
          
          console.log('✅ Leverschema synchronized successfully');
        }
      } else if (sessionsChanged) {
        state.activeWorkSession = null;
        state.clientWorkspaces = {};
        updateHeaderAccountInfo();
        if (state.currentPage === "clients") {
          switchPage("dashboard");
        }
      }
    }
    
    await tryMigrateIndexedDBToCloud();
    await mirrorSessionsToIndexedDB(state.teamSessions);
    persistLeverschemaLocalCache();
  } catch (err) {
    console.warn("Using offline session cache (server unavailable or not configured).", err);
    try {
      state.teamSessions = await idbGetAllSessionsSorted();
    } catch {
      state.teamSessions = [];
    }
    state.leverschemaResults = loadLeverschemaResults();
  }
}

function startTeamStateSync() {
  // Stop any existing sync
  stopTeamStateSync();
  
  // Sync every 8 seconds so sessions created on another computer appear quickly.
  teamStateSyncInterval = setInterval(async () => {
    if (state.user) {
      await refreshTeamState();
    }
  }, 8000);
  
  console.log('Team state sync started (polling every 8 seconds)');
}

function stopTeamStateSync() {
  if (teamStateSyncInterval) {
    clearInterval(teamStateSyncInterval);
    teamStateSyncInterval = null;
    console.log('Team state sync stopped');
  }
}

function upsertTeamSessionAndPersist(session) {
  session.updatedAt = new Date().toISOString();
  const idx = state.teamSessions.findIndex((s) => s.id === session.id);
  if (idx >= 0) {
    state.teamSessions[idx] = session;
  } else {
    state.teamSessions.unshift(session);
  }
  schedulePersistTeamState();
}

function deleteTeamSessionAndPersist(id) {
  deletedTeamSessionIds = [...new Set([...deletedTeamSessionIds, id])];
  state.teamSessions = state.teamSessions.filter((s) => s.id !== id);
  schedulePersistTeamState();
}

async function bootstrapApp() {
  renderClientTabs();
  renderSettingsPage();
  renderClientsLeverschemaCard();
  renderLeverschemaSummary();
  renderLeverschemaWorkbook();
  try {
    const response = await fetch("/api/session");
    if (!response.ok) {
      showAuthScreen();
      return;
    }
    const payload = await response.json();
    await showAppShell(payload.user);
  } catch {
    showAuthScreen();
  }
}

async function handleLogin(event) {
  event.preventDefault();
  const email = loginEmail.value.trim();
  const password = loginPassword.value;
  loginButton.disabled = true;
  loginError.classList.add("hidden");
  loginError.textContent = "";

  try {
    console.log("Attempting login with email:", email);
    const response = await fetch("/api/login", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      credentials: "include",
      body: JSON.stringify({ email, password }),
    });
    console.log("Login response status:", response.status);
    const payload = await response.json();
    console.log("Login response payload:", payload);
    if (!response.ok) {
      loginError.textContent = payload.error || "Unable to sign in.";
      loginError.classList.remove("hidden");
      return;
    }
    loginPassword.value = "";
    console.log("Login successful, showing app shell");
    await showAppShell(payload.user);
  } catch (err) {
    console.error("Login error:", err);
    loginError.textContent = "Unable to reach the login service.";
    loginError.classList.remove("hidden");
  } finally {
    loginButton.disabled = false;
  }
}

async function handleLogout() {
  cancelScheduledTeamStatePersist();
  try {
    await fetch("/api/logout", { method: "POST" });
  } catch {
    // Ignore logout transport errors and clear the local screen anyway.
  }
  resetPreviewState();
  showAuthScreen();
}

async function showAppShell(user) {
  state.user = user;
  authScreen.classList.add("hidden");
  appShell.classList.remove("hidden");
  updateHeaderAccountInfo();
  await hydrateWorkSessionForUser();
  startTeamStateSync(); // Start syncing team state
  switchPage(getInitialOrderFlowPage());
}

function getInitialOrderFlowPage() {
  if (wantsOrderFlowClientsPage() && state.activeWorkSession) {
    return "clients";
  }
  return "dashboard";
}

function wantsOrderFlowClientsPage() {
  const params = new URLSearchParams(window.location.search);
  const page = (params.get("page") || params.get("view") || "").trim().toLowerCase();
  const hash = window.location.hash.replace("#", "").trim().toLowerCase();
  return page === "clients" || page === "session-clients" || hash === "clients" || hash === "session-clients";
}

function showAuthScreen() {
  stopTeamStateSync(); // Stop syncing when logged out
  cancelScheduledTeamStatePersist();
  state.user = null;
  state.activeWorkSession = null;
  state.teamSessions = [];
  state.leverschemaResults = {};
  window.localStorage.removeItem(ACTIVE_WORK_SESSION_STORAGE_KEY);
  authScreen.classList.remove("hidden");
  appShell.classList.add("hidden");
  updateHeaderAccountInfo();
}

function getActiveSessionSheet() {
  if (!state.activeWorkSession?.date) return "Monday-Thursday";
  const sessionDate = parseInputDate(state.activeWorkSession.date);
  if (!sessionDate) return "Monday-Thursday";
  const day = sessionDate.getDay();
  if (day === 5) return "Friday";
  if (day === 6) return "Saturday";
  return "Monday-Thursday";
}

function createEmptyClientWorkspace() {
  return {
    mode: null,
    preview: null,
    selectedIndex: 0,
    orderSelectionActive: false,
    activeDeliveryPointKey: null,
    leverschemaIncludedIndexes: [],
    exportSheet: "Paderborn",
    emailCacheId: null,
    emailFileName: "",
  };
}

function canonicalClientName(client) {
  const normalized = normalizeText(client);
  if (normalized === "carrefour fif" || normalized === "carrefour kdc") return "Carrefour";
  if (normalized === "edeka laatzen" || normalized === "edeka mochmuhl") return "Edeka";
  if (normalized === "colruyt saturday") return "Colruyt Saturday";
  if (normalized === "havi duisburg saturday") return "Havi Duisburg Saturday";
  if (normalized === "havi" || normalized.startsWith("havi ")) return "HAVI";
  return client || DASHBOARD_CLIENTS[0];
}

function migrateClientWorkspaces(workspaces) {
  const migrated = { ...(workspaces || {}) };
  Object.entries(CLIENT_WORKSPACE_ALIASES).forEach(([canonical, aliases]) => {
    let combined = migrated[canonical] ? cloneWorkspace(migrated[canonical]) : null;
    aliases.forEach((alias) => {
      if (!migrated[alias]) return;
      combined = mergeClientWorkspaces(combined, migrated[alias]);
      delete migrated[alias];
    });
    if (combined) migrated[canonical] = combined;
  });
  return migrated;
}

function cloneWorkspace(workspace) {
  return JSON.parse(JSON.stringify(workspace || createEmptyClientWorkspace()));
}

function mergeClientWorkspaces(target, source) {
  if (!source) return target || null;
  if (!target) return cloneWorkspace(source);
  const merged = cloneWorkspace(target);
  const incoming = cloneWorkspace(source);

  if (merged.mode === incoming.mode && Array.isArray(merged.preview?.orders) && Array.isArray(incoming.preview?.orders)) {
    const orders = [...merged.preview.orders];
    incoming.preview.orders.forEach((order) => {
      const duplicateIndex = orders.findIndex((existing) => buildOrderIdentity(existing) === buildOrderIdentity(order));
      if (duplicateIndex >= 0) {
        orders[duplicateIndex] = order;
      } else {
        orders.push(order);
      }
    });
    merged.preview = {
      ...merged.preview,
      deliveryDate: merged.preview.deliveryDate || incoming.preview.deliveryDate || "",
      customerCount: orders.length,
      orders,
      canMerge: orders.length > 0,
    };
    merged.leverschemaIncludedIndexes = normalizeIncludedIndexes(merged.leverschemaIncludedIndexes, orders.length);
    return merged;
  }

  if (merged.mode === "special" && incoming.mode === "special" && merged.preview?.items && incoming.preview?.items) {
    merged.preview = mergeSpecialPreviews(merged.preview, incoming.preview);
    merged.selectedIndex = Math.max(0, (merged.preview.specialOrders || []).length - 1);
    merged.leverschemaIncludedIndexes = [merged.selectedIndex];
    return merged;
  }

  if (!merged.preview && incoming.preview) {
    return incoming;
  }
  return merged;
}

function normalizeIncludedIndexes(indexes, orderCount) {
  const valid = (indexes || []).filter((index) => Number.isInteger(index) && index >= 0 && index < orderCount);
  return valid.length ? [...new Set(valid)] : orderCount ? [0] : [];
}

function getClientWorkspace(client) {
  const canonicalClient = canonicalClientName(client);
  if (!state.clientWorkspaces[canonicalClient]) {
    state.clientWorkspaces[canonicalClient] = createEmptyClientWorkspace();
  }
  return state.clientWorkspaces[canonicalClient];
}

function syncCurrentClientWorkspace() {
  state.selectedClient = canonicalClientName(state.selectedClient);
  const workspace = getClientWorkspace(state.selectedClient);
  workspace.mode = state.mode;
  workspace.preview = state.preview;
  workspace.selectedIndex = state.selectedIndex;
  workspace.orderSelectionActive = state.orderSelectionActive;
  workspace.activeDeliveryPointKey = state.activeDeliveryPointKey;
  workspace.leverschemaIncludedIndexes = [...state.leverschemaIncludedIndexes];
  workspace.exportSheet = exportSheetSelect?.value || workspace.exportSheet || "Paderborn";
  workspace.emailFileName = state.file?.name || workspace.emailFileName || "";
  workspace.emailCacheId = state.activeWorkSession
    ? buildWorkspaceEmailId(state.activeWorkSession.id, state.selectedClient)
    : null;

  if (workspace.emailCacheId && state.file) {
    idbSaveWorkspaceEmail(workspace.emailCacheId, state.file).catch((err) => {
      console.error("Workspace email cache save failed:", err);
    });
  }

  if (state.activeWorkSession) {
    state.activeWorkSession.workspaces = state.clientWorkspaces;
    upsertTeamSessionAndPersist(state.activeWorkSession);
  }
}

function loadClientWorkspace(client) {
  const canonicalClient = canonicalClientName(client);
  state.selectedClient = canonicalClient;
  const workspace = getClientWorkspace(canonicalClient);
  state.file = null;
  state.mode = workspace.mode;
  state.preview = workspace.preview;
  state.selectedIndex = workspace.selectedIndex || 0;
  state.orderSelectionActive = workspace.orderSelectionActive === true;
  state.activeDeliveryPointKey = workspace.activeDeliveryPointKey || null;
  state.leverschemaIncludedIndexes = [...(workspace.leverschemaIncludedIndexes || [])];
  if (exportSheetSelect) {
    exportSheetSelect.value = workspace.exportSheet || "Paderborn";
  }
  if (emailInput) {
    emailInput.value = "";
  }
  renderPreview();
  restoreWorkspaceEmail(canonicalClient, workspace);
}

function resetPreviewState() {
  state.file = null;
  state.mode = null;
  state.preview = null;
  state.clientWorkspaces = {};
  state.selectedIndex = 0;
  state.orderSelectionActive = false;
  state.activeDeliveryPointKey = null;
  state.leverschemaIncludedIndexes = [];
  clearPreviewDisplay();
}

function clearPreviewDisplay() {
  if (emailInput) emailInput.value = "";
  customersFound.textContent = "0";
  itemsInSelection.textContent = "0";
  deliveryDate.textContent = "-";
  detailCustomer.textContent = "-";
  detailReference.textContent = "-";
  detailFatrans.textContent = "-";
  tableHead.innerHTML = "";
  tableBody.innerHTML = "";
  orderSelect.innerHTML = "";
  renderDeliverySidebar();
  if (exportSheetSelect) {
    exportSheetSelect.value = "Paderborn";
  }
  exportButton.disabled = true;
  mergeButton.disabled = true;
  if (saveOrderButton) saveOrderButton.disabled = true;
  setAddItemEnabled(false);
  specialButton.disabled = true;
  nettoMdButton.disabled = true;
  leverschemaButton.disabled = true;
  printCmrButton.disabled = true;
  setNettoMdCmrMenuOpen(false);
}

async function restoreWorkspaceEmail(client, workspace) {
  const cacheId =
    workspace?.emailCacheId ||
    (state.activeWorkSession ? buildWorkspaceEmailId(state.activeWorkSession.id, client) : null);
  const restoreToken = ++state.fileRestoreToken;

  if (!cacheId || !workspace?.preview) {
    return;
  }

  try {
    const cached = await idbGetWorkspaceEmail(cacheId);
    if (!cached?.blob) {
      if (
        restoreToken === state.fileRestoreToken &&
        state.selectedClient === client &&
        state.preview
      ) {
        statusText.textContent = `Preview restored for ${client}. Reload the email file to export again.`;
      }
      return;
    }

    const restoredFile = new File(
      [cached.blob],
      cached.name || workspace.emailFileName || "email.eml",
      { type: cached.type || "message/rfc822" }
    );

    if (restoreToken !== state.fileRestoreToken || state.selectedClient !== client) {
      return;
    }

    state.file = restoredFile;
    statusText.textContent = `Email restored for ${client}. Ready to export.`;
  } catch (err) {
    console.error("Workspace email restore failed:", err);
    if (
      restoreToken === state.fileRestoreToken &&
      state.selectedClient === client &&
      state.preview
    ) {
      statusText.textContent = `Preview restored for ${client}. Reload the email file to export again.`;
    }
  }
}

async function hydrateWorkSessionForUser() {
  await refreshTeamState();
  const today = formatDateForInput(new Date());
  const sessions = getTeamSessionsSorted();
  const storedSessionId = window.localStorage.getItem(ACTIVE_WORK_SESSION_STORAGE_KEY) || "";
  const storedSession = sessions.find((s) => s.id === storedSessionId);
  const todaysSession = sessions.find((s) => s.date === today);
  state.activeWorkSession = storedSession || todaysSession || (wantsOrderFlowClientsPage() ? sessions[0] : null) || null;

  const seedDate = state.activeWorkSession?.date || today;
  const seedName = state.activeWorkSession?.name || buildDefaultSessionName(seedDate);
  workSessionDate.value = seedDate;
  workSessionDate.dataset.previousValue = seedDate;
  workSessionName.value = seedName;
  workSessionName.dataset.autoName = seedName === buildDefaultSessionName(seedDate) ? "true" : "false";

  if (state.activeWorkSession) {
    state.clientWorkspaces = migrateClientWorkspaces(state.activeWorkSession.workspaces || {});
    state.activeWorkSession.workspaces = state.clientWorkspaces;
  } else {
    state.clientWorkspaces = {};
  }

  updateHeaderAccountInfo();
  loadClientWorkspace(state.selectedClient || DASHBOARD_CLIENTS[0]);
  renderSessionHistoryList();
}

async function handleWorkSessionCreate(event) {
  event.preventDefault();
  if (!state.user) return;
  if (workSessionCreateInFlight) return;
  workSessionCreateInFlight = true;

  const sessionDate = workSessionDate.value || formatDateForInput(new Date());
  const sessionName = (workSessionName.value || buildDefaultSessionName(sessionDate)).trim();
  createWorkSessionButton.disabled = true;
  statusText.textContent = "Creating session...";

  try {
    const response = await fetch("/api/work_sessions", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      credentials: "include",
      body: JSON.stringify({ date: sessionDate, name: sessionName }),
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(payload.error || `Session creation failed (HTTP ${response.status}).`);
    }

    const session = payload.session;
    if (!session?.id) {
      throw new Error("Session creation failed: the server did not return a session.");
    }

    state.teamSessions = Array.isArray(payload.sessions) ? payload.sessions : [session, ...state.teamSessions];
    state.leverschemaResults =
      payload.leverschemaResults && typeof payload.leverschemaResults === "object" ? payload.leverschemaResults : state.leverschemaResults;
    state.laadschemaData =
      payload.laadschemaData && typeof payload.laadschemaData === "object" ? payload.laadschemaData : state.laadschemaData;
    state.laadschemaCustomTrucks =
      payload.laadschemaCustomTrucks && typeof payload.laadschemaCustomTrucks === "object"
        ? payload.laadschemaCustomTrucks
        : state.laadschemaCustomTrucks;

    const storedSession = state.teamSessions.find((entry) => entry.id === session.id) || session;
    state.activeWorkSession = storedSession;
    state.clientWorkspaces = migrateClientWorkspaces(storedSession.workspaces || {});
    state.activeWorkSession.workspaces = state.clientWorkspaces;
    await mirrorSessionsToIndexedDB(state.teamSessions);

    updateHeaderAccountInfo();
    loadClientWorkspace(state.selectedClient || DASHBOARD_CLIENTS[0]);
    renderSessionHistoryList();

    statusText.textContent = `${session.name} created.`;
    switchPage("clients");
  } catch (err) {
    console.error("Work session creation failed:", err);
    statusText.textContent = err.message || "Session creation failed.";
    sessionHistoryList.innerHTML = `<p class="history-empty state-meta">${escapeHtml(statusText.textContent)}</p>`;
  } finally {
    workSessionCreateInFlight = false;
    createWorkSessionButton.disabled = false;
  }
}

function syncWorkSessionNameFromDate() {
  const current = workSessionName.value.trim();
  const autoName = workSessionName.dataset.autoName === "true";
  const previousDefault = buildDefaultSessionName(workSessionDate.dataset.previousValue || formatDateForInput(new Date()));
  if (autoName || !current || current === previousDefault) {
    workSessionName.value = buildDefaultSessionName(workSessionDate.value);
    workSessionName.dataset.autoName = "true";
  } else {
    workSessionName.dataset.autoName = "false";
  }
  workSessionDate.dataset.previousValue = workSessionDate.value;
}

workSessionName.addEventListener("input", () => {
  const defaultName = buildDefaultSessionName(workSessionDate.value || formatDateForInput(new Date()));
  workSessionName.dataset.autoName = workSessionName.value.trim() === defaultName ? "true" : "false";
});

function renderSessionHistoryList() {
  const sessions = getTeamSessionsSorted();

  sessionHistoryList.innerHTML = "";
  if (sessions.length === 0) {
    sessionHistoryList.innerHTML = `<p class="history-empty state-meta">No sessions saved yet. Create your first session above.</p>`;
    return;
  }

  sessions.forEach(session => {
    const isActive = state.activeWorkSession?.id === session.id;
    const card = document.createElement("div");
    card.className = `history-item-card ${isActive ? "is-active" : ""}`;
    card.innerHTML = `
      <div class="history-item-info">
        <strong>${escapeHtml(session.name)}</strong>
        <span>Date: ${formatSessionDate(session.date)} • Saved clients: ${Object.keys(session.workspaces || {}).length}</span>
      </div>
      <div class="history-item-actions">
        <button type="button" class="btn-delete" onclick="deleteHistoricalSession('${session.id}')">Delete</button>
        <button type="button" class="btn-open" onclick="openHistoricalSession('${session.id}')">Open</button>
      </div>
    `;
    sessionHistoryList.appendChild(card);
  });
}

function openHistoricalSession(id) {
  const session = state.teamSessions.find((s) => s.id === id);
  if (!session) return;

  state.activeWorkSession = session;
  state.clientWorkspaces = migrateClientWorkspaces(session.workspaces || {});
  state.activeWorkSession.workspaces = state.clientWorkspaces;

  workSessionDate.value = session.date;
  workSessionDate.dataset.previousValue = session.date;
  workSessionName.value = session.name;
  workSessionName.dataset.autoName = "false";

  updateHeaderAccountInfo();
  loadClientWorkspace(state.selectedClient || DASHBOARD_CLIENTS[0]);
  renderSessionHistoryList();

  statusText.textContent = `${session.name} opened.`;
  switchPage("clients");
}

async function deleteHistoricalSession(id) {
  const shouldDelete = await showConfirmationDialog({
    eyebrow: "Delete session",
    title: "Delete this session?",
    message: "This will remove the session and its saved workspace data. This action cannot be undone.",
    confirmLabel: "Delete session",
    cancelLabel: "Cancel",
  });
  if (!shouldDelete) return;
  deleteTeamSessionAndPersist(id);
  if (state.activeWorkSession?.id === id) {
    state.activeWorkSession = null;
    state.clientWorkspaces = {};
    updateHeaderAccountInfo();
    loadClientWorkspace(state.selectedClient || DASHBOARD_CLIENTS[0]);
  }
  renderSessionHistoryList();
}

window.openHistoricalSession = openHistoricalSession;
window.deleteHistoricalSession = deleteHistoricalSession;

function showConfirmationDialog(options = {}) {
  if (!confirmationModal) {
    return Promise.resolve(window.confirm(options.message || "Are you sure?"));
  }
  confirmationModalEyebrow.textContent = options.eyebrow || "Confirm action";
  confirmationModalTitle.textContent = options.title || "Are you sure?";
  confirmationModalMessage.textContent = options.message || "This action cannot be undone.";
  confirmationConfirmButton.textContent = options.confirmLabel || "Confirm";
  confirmationCancelButton.textContent = options.cancelLabel || "Cancel";
  confirmationConfirmButton.classList.toggle("danger", options.danger !== false);
  confirmationModal.classList.remove("hidden");
  confirmationConfirmButton.focus();

  return new Promise((resolve) => {
    const cleanup = (result) => {
      confirmationModal.classList.add("hidden");
      confirmationConfirmButton.removeEventListener("click", onConfirm);
      confirmationCancelButton.removeEventListener("click", onCancel);
      confirmationModal.removeEventListener("click", onBackdrop);
      document.removeEventListener("keydown", onKeyDown);
      resolve(result);
    };
    const onConfirm = () => cleanup(true);
    const onCancel = () => cleanup(false);
    const onBackdrop = (event) => {
      if (event.target.closest("[data-confirmation-cancel]")) cleanup(false);
    };
    const onKeyDown = (event) => {
      if (event.key === "Escape") cleanup(false);
    };
    confirmationConfirmButton.addEventListener("click", onConfirm);
    confirmationCancelButton.addEventListener("click", onCancel);
    confirmationModal.addEventListener("click", onBackdrop);
    document.addEventListener("keydown", onKeyDown);
  });
}

function openSettingsPage() {
  state.settingsClient = state.selectedClient || DASHBOARD_CLIENTS[0];
  renderSettingsPage();
  switchPage("settings");
}

function renderClientTabs() {
  const activeSheet = getActiveSessionSheet();
  const clientsToShow = DASHBOARD_CLIENTS.filter((client) => {
    if (client.includes("Saturday") && activeSheet !== "Friday") return false;
    return true;
  });

  clientCards.innerHTML = clientsToShow.map((client) => {
    const parts = client.split("\n");
    const title = parts[0];
    const subtext = parts.slice(1).join("\n");
    
    return `
      <button class="client-card-button ${client === state.selectedClient ? "active" : ""}" type="button" data-client-card="${escapeHtml(client)}">
        <span class="client-card-title">${escapeHtml(title)}</span>
        ${subtext ? `<span class="client-card-subtext">${escapeHtml(subtext)}</span>` : ""}
      </button>
    `;
  }).join("");
}

function handleClientCardClick(event) {
  const button = event.target.closest("[data-client-card]");
  if (!button) return;
  syncCurrentClientWorkspace();
  state.selectedClient = button.dataset.clientCard;
  renderClientTabs();
  loadClientWorkspace(state.selectedClient);
  statusText.textContent = `${state.selectedClient} selected.`;
  switchPage("orders");
}

function renderSettingsPage() {
  if (!settingsClientSelect || !settingsActionList) return;

  const activeSheet = getActiveSessionSheet();
  const clientsToShow = DASHBOARD_CLIENTS.filter((client) => {
    if (client.includes("Saturday") && activeSheet !== "Friday") return false;
    return true;
  });

  settingsClientSelect.innerHTML = clientsToShow.map((client) => `
    <option value="${escapeHtml(client)}" ${client === state.settingsClient ? "selected" : ""}>${escapeHtml(client)}</option>
  `).join("");

  const preferences = getClientActionPreferences(state.settingsClient);
  settingsActionList.innerHTML = ACTION_VISIBILITY_OPTIONS.map((option) => `
    <label class="settings-action-option">
      <input type="checkbox" data-settings-action="${escapeHtml(option.key)}" ${preferences[option.key] ? "checked" : ""}>
      <div>
        <strong>${escapeHtml(option.label)}</strong>
        <span>${escapeHtml(option.description)}</span>
      </div>
    </label>
  `).join("");

  renderHaviUienSettingsCard();
}

function handleSettingsClientChange() {
  state.settingsClient = settingsClientSelect.value || DASHBOARD_CLIENTS[0];
  renderSettingsPage();
}

function handleSettingsActionChange(event) {
  const checkbox = event.target.closest("[data-settings-action]");
  if (!checkbox) return;

  state.settingsClient = settingsClientSelect.value || DASHBOARD_CLIENTS[0];
  const actionKey = checkbox.dataset.settingsAction;
  const preferences = getClientActionPreferences(state.settingsClient);
  preferences[actionKey] = checkbox.checked;
  state.clientActionSettings[state.settingsClient] = preferences;
  persistClientActionSettings();
  renderSettingsPage();
  updateOrdersClientMode();
  statusText.textContent = `${state.settingsClient} settings updated.`;
}

function renderHaviUienSettingsCard() {
  if (!haviUienSettingsCard || !haviUienArticleInput || !haviUienDescriptionInput) return;
  const normalizedClient = normalizeText(state.settingsClient);
  const isHavi = normalizedClient === "havi" || normalizedClient === "havi duisburg saturday";
  haviUienSettingsCard.classList.toggle("hidden", !isHavi);
  const settings = getHaviUienSettings();
  haviUienArticleInput.value = settings.article;
  haviUienDescriptionInput.value = settings.description;
}

function handleHaviUienSettingsInput() {
  state.haviUienSettings = {
    article: haviUienArticleInput?.value.trim() || "",
    description: haviUienDescriptionInput?.value.trim() || "",
  };
  persistHaviUienSettings();
  if (isSelectedHaviDeUienOrder()) {
    renderPreview();
  }
  statusText.textContent = "HAVI DE UIEN settings updated.";
}

function renderClientsLeverschemaCard() {
  const entries = getCurrentSessionLeverschemaEntries();
  const totalEntries = entries.length;

  if (!totalEntries) {
    openLeverschemaFromClientsButton.disabled = true;
    exportLeverschemaFromClientsButton.disabled = true;
    return;
  }

  openLeverschemaFromClientsButton.disabled = false;
  exportLeverschemaFromClientsButton.disabled = false;
}

function isNettoMdClientSelected() {
  return normalizeText(state.selectedClient) === "nettomd";
}

function isReweOrPennyClientSelected() {
  const normalized = normalizeText(state.selectedClient);
  return normalized === "rewe" || normalized === "penny";
}

function getSelectedExportClient(type = "") {
  if (type !== "print_cmr") return state.selectedClient || "";
  const normalizedClient = normalizeText(state.selectedClient);
  const order = getCmrContextOrder() || getCurrentOrder();
  const orderText = normalizeText(`${order?.customer || ""} ${order?.fatrans || ""} ${order?.deliveryPoint || ""}`);
  const activePoint = normalizeText(getCmrActiveDeliveryPoint() || "");

  if (normalizedClient === "carrefour") {
    return activePoint.includes("kdc") || orderText.includes("kdc") ? "Carrefour KDC" : "Carrefour FIF";
  }
  if (normalizedClient === "edeka") {
    return orderText.includes("mochmuhl") || orderText.includes("mockmuhl") || orderText.includes("mockmuehl")
      ? "Edeka Mochmuhl"
      : "Edeka Laatzen";
  }
  if (normalizedClient === "havi") {
    if (orderText.includes("duisburg") || orderText.includes("wunstorf")) return "Havi Duisburg";
    if (orderText.includes("belg") || /\bbe\b/.test(orderText)) return "Havi BE";
    if (orderText.includes("nether") || orderText.includes("nederland") || /\bnl\b/.test(orderText)) return "Havi NL";
    if (orderText.includes("de ") || orderText.includes("gmbh") || orderText.includes("uien")) return "Havi DE";
    return "Havi DE";
  }
  return state.selectedClient || "";
}

function getSelectedHaviDcName() {
  const order = getCurrentOrder();
  const orderText = normalizeText(`${order?.customer || ""} ${order?.fatrans || ""} ${order?.deliveryPoint || ""}`);
  if (orderText.includes("neu wulmstorf")) return "Neu Wulmstorf";
  if (orderText.includes("wunstorf")) return "Wunstorf";
  if (orderText.includes("duisburg")) return "Duisburg";
  return "";
}

function applyStandardClientDisplayRules(order) {
  if (!order) return order;
  const normalizedClient = normalizeText(state.selectedClient);
  if (normalizedClient === "rewe") {
    return {
      ...order,
      // Keep the original customer (location), just update fatrans
      fatrans: "Rewe",
    };
  }
  if (normalizedClient === "penny") {
    return {
      ...order,
      // Keep the original customer (location), just update fatrans
      fatrans: "Penny",
    };
  }
  return order;
}

function enableCMRButtonForSpecialClients() {
  // Enable Print CMR for specific clients even without email
  const normalizedClient = normalizeText(state.selectedClient);
  if (isPrintCmrClient(normalizedClient)) {
    printCmrButton.disabled = false;
    // Also ensure the button is visible
    printCmrButton.classList.remove("hidden");
  }
  
  // Handle Havi 3 DCs - enable all 3 buttons
  if (isHavi3DCClient(normalizedClient)) {
    printCmrDuisburgButton.disabled = false;
    printCmrWunstorfButton.disabled = false;
    printCmrNeuWulmstorfButton.disabled = false;
  }
}

function isPrintCmrClient(normalizedClient) {
  return [
    "carrefour",
    "carrefour fif",
    "carrefour kdc",
    "colruyt",
    "colruyt saturday",
    "denemark",
    "edeka",
    "edeka laatzen",
    "edeka mochmuhl",
    "globus",
    "havi",
    "havi nl",
    "havi be",
    "havi de",
    "havi de saturday",
    "havi duisburg saturday",
    "heeren",
    "nettomd",
    "rewe",
    "penny",
    "hanos",
  ].includes(normalizedClient) || isHavi3DCClient(normalizedClient);
}

function isHavi3DCClient(normalizedClient) {
  const result = normalizedClient === "havi dc duisburg dc wunstorf dc neu wulmstorf";
  console.log("isHavi3DCClient check:", {
    input: normalizedClient,
    expected: "havi dc duisburg dc wunstorf dc neu wulmstorf",
    result: result
  });
  return result;
}

function isNettoMDClient(normalizedClient) {
  return normalizedClient === "nettomd";
}

function updateHaviDCButtonsVisibility() {
  const isHavi3DC = isHavi3DCClient(normalizeText(state.selectedClient));
  if (!isHavi3DC) return;

  // Get the currently selected customer from the dropdown
  const selectedOrder = getCurrentOrder();
  const customerName = selectedOrder?.customer || "";
  
  console.log("updateHaviDCButtonsVisibility:", {
    selectedOrder: selectedOrder,
    customerName: customerName
  });

  // Hide all DC buttons first
  printCmrDuisburgButton.classList.add("hidden");
  printCmrWunstorfButton.classList.add("hidden");
  printCmrNeuWulmstorfButton.classList.add("hidden");

  // Show the correct button based on customer name
  if (customerName.toLowerCase().includes("duisburg")) {
    printCmrDuisburgButton.classList.remove("hidden");
  } else if (customerName.toLowerCase().includes("wunstorf")) {
    printCmrWunstorfButton.classList.remove("hidden");
  } else if (customerName.toLowerCase().includes("neu wulmstorf")) {
    printCmrNeuWulmstorfButton.classList.remove("hidden");
  } else {
    // Default to Duisburg if no specific DC is detected
    printCmrDuisburgButton.classList.remove("hidden");
  }
}

function toggleNettoMdCmrMenu() {
  const isOpen = nettomdCmrToggleButton?.getAttribute("aria-expanded") === "true";
  setNettoMdCmrMenuOpen(!isOpen);
}

function setNettoMdCmrMenuOpen(isOpen) {
  nettomdCmrToggleButton?.setAttribute("aria-expanded", isOpen ? "true" : "false");
  if (nettomdCmrToggleButton) nettomdCmrToggleButton.textContent = isOpen ? "Print CMR v" : "Print CMR >";
  nettomdCmrMenu?.classList.toggle("hidden", !isOpen);
}

function updateOrdersClientMode() {
  const preferences = getClientActionPreferences(state.selectedClient);
  const isNettoMd = isNettoMdClientSelected();
  const needsSheetSelection = isReweOrPennyClientSelected();
  const isHavi3DC = isHavi3DCClient(normalizeText(state.selectedClient));
  const isNettoMDCMR = isNettoMDClient(normalizeText(state.selectedClient));
  
  console.log("updateOrdersClientMode:", {
    selectedClient: state.selectedClient,
    normalizedClient: normalizeText(state.selectedClient),
    isHavi3DC: isHavi3DC,
    isNettoMDCMR: isNettoMDCMR,
    preferences: preferences,
    printCmrHidden: isNettoMd || !preferences.printCmr || isHavi3DC || isNettoMDCMR
  });
  
  exportButton.classList.toggle("hidden", isNettoMd || !preferences.export);
  mergeButton.classList.toggle("hidden", isNettoMd || !preferences.merge);
  specialButton.classList.toggle("hidden", isNettoMd || !preferences.special);
  nettoMdButton.classList.toggle("hidden", !preferences.nettoMd);
  leverschemaButton.classList.toggle("hidden", isNettoMd || !preferences.leverschema);
  printCmrButton.classList.toggle("hidden", isNettoMd || !preferences.printCmr || isHavi3DC || isNettoMDCMR);
  
  // Show/hide Havi 3 DC CMR buttons container
  haviDcCmrButtons.classList.toggle("hidden", !isHavi3DC || !preferences.printCmr);
  
  // Show/hide NettoMD CMR buttons container
  nettomdCmrButtons.classList.toggle("hidden", !isNettoMDCMR || !preferences.printCmr);
  if (!isNettoMDCMR || !preferences.printCmr) {
    setNettoMdCmrMenuOpen(false);
  }
  
  // Update individual DC button visibility
  if (isHavi3DC && preferences.printCmr) {
    updateHaviDCButtonsVisibility();
  } else if (isHavi3DC && !preferences.printCmr) {
    // Hide all individual Havi DC buttons when printCmr is disabled
    printCmrDuisburgButton.classList.add("hidden");
    printCmrWunstorfButton.classList.add("hidden");
    printCmrNeuWulmstorfButton.classList.add("hidden");
  }
  
  // Enable all NettoMD buttons if NettoMD client
  if (isNettoMDCMR && preferences.printCmr) {
    nettomdCmrToggleButton.disabled = false;
    printCmrKerpenButton.disabled = false;
    printCmrHodenhagenButton.disabled = false;
    printCmrHenstedtButton.disabled = false;
    printCmrHammButton.disabled = false;
    printCmrGanderkeseeButton.disabled = false;
    printCmrBottropButton.disabled = false;
    printCmrKrefeldButton.disabled = false;
  } else {
    nettomdCmrToggleButton.disabled = true;
  }
  
  exportSheetSelect?.classList.toggle("hidden", !needsSheetSelection);
  exportSheetSelect?.classList.toggle("sheet-select", needsSheetSelection);
  leverschemaSummaryCard.classList.toggle("hidden", isNettoMd || !preferences.leverschema);
  if (isNettoMd || !preferences.leverschema) {
    leverschemaIncludeCard.classList.add("hidden");
  }
}

async function handleUpload() {
  const file = emailInput.files[0];
  if (!file) return;

  state.file = file;
  statusText.textContent = "Uploading and parsing email...";

  const formData = new FormData();
  formData.append("email", file);

  try {
    const response = await fetch("/api/parse", { method: "POST", body: formData });
    if (response.status === 401) {
      resetPreviewState();
      showAuthScreen();
      return;
    }
    const payload = await response.json();
    if (!response.ok) {
      statusText.textContent = payload.error || "Failed to parse the email.";
      return;
    }

    const addResult = addParsedPreviewToCurrentWorkspace(payload.mode, payload.preview, file);
    state.mode = addResult.mode;
    state.preview = addResult.preview;
    state.selectedIndex = addResult.selectedIndex;
    state.orderSelectionActive = false;
    state.activeDeliveryPointKey = null;
    state.leverschemaIncludedIndexes = addResult.includedIndexes;

    syncCurrentClientWorkspace();
    await persistTeamStateImmediately();
    renderPreview();
    const isSaturdayCarryover = Boolean(getSaturdayCarryoverConfig());
    const shortageMessage = isSaturdayCarryover ? "" : await saveOrdersForShortages(file);
    const carryoverMessage = await propagateSaturdayOrdersToNextSession(file);
    statusText.textContent = [addResult.message, shortageMessage, carryoverMessage].filter(Boolean).join(" ");
  } catch {
    statusText.textContent = "Failed to reach the parser.";
  }
}

function addParsedPreviewToCurrentWorkspace(mode, preview, file) {
  if (!preview) {
    return {
      mode: null,
      preview: null,
      selectedIndex: 0,
      includedIndexes: [],
      message: "No order data found in this email.",
    };
  }

  const incomingPreview = normalizePreviewForSession(mode, preview, file);
  if (isHaviWorkspaceSelected() && state.preview && (state.mode === "special" || mode === "special")) {
    const existingOrders = previewToWorkspaceOrders(state.mode, state.preview);
    const incomingOrders = previewToWorkspaceOrders(mode, incomingPreview);
    const merged = mergeWorkspaceOrders(existingOrders, incomingOrders);
    const selectedIndex = Math.max(0, Math.min(merged.firstChangedIndex, merged.orders.length - 1));
    const nextPreview = {
      ...state.preview,
      ...incomingPreview,
      deliveryDate: incomingPreview.deliveryDate || state.preview.deliveryDate || "",
      customerCount: merged.orders.length,
      orders: merged.orders,
      canMerge: merged.orders.length > 0,
    };
    delete nextPreview.items;
    delete nextPreview.specialOrders;

    const includedIndexes = getSameDeliveryPointIndexes(merged.orders, selectedIndex);
    return {
      mode: "standard",
      preview: nextPreview,
      selectedIndex,
      includedIndexes: includedIndexes.length ? includedIndexes : [selectedIndex],
      message: `${merged.orders.length} HAVI order${merged.orders.length === 1 ? "" : "s"} saved to this session.`,
    };
  }

  if (mode === "special" && state.mode === "special" && state.preview?.items?.length) {
    const mergedSpecial = mergeSpecialPreviews(state.preview, incomingPreview);
    const selectedIndex = Math.max(0, (mergedSpecial.specialOrders || []).length - 1);
    return {
      mode,
      preview: selectSpecialPreview(mergedSpecial, selectedIndex),
      selectedIndex,
      includedIndexes: [selectedIndex],
      message: `${mergedSpecial.specialOrders.length} HAVI order${mergedSpecial.specialOrders.length === 1 ? "" : "s"} saved to this session.`,
    };
  }

  const canAppendOrders =
    Array.isArray(incomingPreview?.orders) &&
    Array.isArray(state.preview?.orders) &&
    state.mode === mode;

  if (!canAppendOrders) {
    const selectedIndex = 0;
    return {
      mode,
      preview: incomingPreview,
      selectedIndex,
      includedIndexes: mode === "special" ? [0] : [selectedIndex],
      message: "Order saved to this session.",
    };
  }

  const existingOrders = state.preview.orders || [];
  const incomingOrders = incomingPreview.orders || [];
  const combinedOrders = [...existingOrders];
  let firstNewIndex = combinedOrders.length;
  let addedCount = 0;
  let updatedCount = 0;

  incomingOrders.forEach((order) => {
    const duplicateIndex = combinedOrders.findIndex((existing) => buildOrderIdentity(existing) === buildOrderIdentity(order));
    if (duplicateIndex >= 0) {
      combinedOrders[duplicateIndex] = order;
      updatedCount += 1;
      if (addedCount === 0) firstNewIndex = duplicateIndex;
      return;
    }
    combinedOrders.push(order);
    addedCount += 1;
  });

  const selectedIndex = Math.max(0, Math.min(firstNewIndex, combinedOrders.length - 1));
  const nextPreview = {
    ...state.preview,
    deliveryDate: incomingPreview.deliveryDate || state.preview.deliveryDate || "",
    customerCount: combinedOrders.length,
    dcName: incomingPreview.dcName || state.preview.dcName || "",
    orders: combinedOrders,
    canMerge: combinedOrders.length > 0,
  };

  const includedIndexes = mode === "standard"
    ? getSameDeliveryPointIndexes(combinedOrders, selectedIndex)
    : [selectedIndex];

  const messageParts = [];
  if (addedCount) messageParts.push(`${addedCount} new order${addedCount === 1 ? "" : "s"}`);
  if (updatedCount) messageParts.push(`${updatedCount} updated`);
  return {
    mode,
    preview: nextPreview,
    selectedIndex,
    includedIndexes: includedIndexes.length ? includedIndexes : [selectedIndex],
    message: `${messageParts.join(", ") || "Order"} saved to this session.`,
  };
}

function normalizePreviewForSession(mode, preview, file) {
  if (mode === "special") {
    const sourceFileName = file?.name || preview.sourceFileName || "";
    return {
      ...preview,
      sourceFileName,
      sourceOrderIndex: Number.isInteger(preview.sourceOrderIndex) ? preview.sourceOrderIndex : 0,
      deliveryPoint: preview.deliveryPoint || getSpecialDeliveryPoint(preview, sourceFileName),
      specialOrders: Array.isArray(preview.specialOrders) ? preview.specialOrders : undefined,
    };
  }
  if (!preview?.orders?.length) return preview;
  return {
    ...preview,
    customerCount: preview.orders.length,
    orders: preview.orders.map((order, index) => ({
      ...order,
      deliveryDate: order.deliveryDate || preview.deliveryDate || "",
      sourceFileName: file?.name || order.sourceFileName || "",
      sourceOrderIndex: Number.isInteger(order.sourceOrderIndex) ? order.sourceOrderIndex : index,
    })),
  };
}

function getSameDeliveryPointIndexes(orders, selectedIndex) {
  const selectedDeliveryPoint = getDeliveryPointKey(orders[selectedIndex]);
  return orders
    .map((order, index) => ({ order, index }))
    .filter(({ order }) => getDeliveryPointKey(order) === selectedDeliveryPoint)
    .map(({ index }) => index);
}

function mergeSpecialPreviews(existingPreview, incomingPreview) {
  const existingOrders = Array.isArray(existingPreview.specialOrders)
    ? existingPreview.specialOrders
    : [stripSpecialOrders(existingPreview)];
  const incomingOrder = stripSpecialOrders(incomingPreview);
  const combined = [...existingOrders];
  const duplicateIndex = combined.findIndex((entry) => buildSpecialOrderIdentity(entry) === buildSpecialOrderIdentity(incomingOrder));
  if (duplicateIndex >= 0) {
    combined[duplicateIndex] = incomingOrder;
  } else {
    combined.push(incomingOrder);
  }
  return {
    ...incomingOrder,
    customerCount: combined.length,
    specialOrders: combined,
  };
}

function selectSpecialPreview(preview, selectedIndex = 0) {
  const specialOrders = Array.isArray(preview.specialOrders) ? preview.specialOrders : [stripSpecialOrders(preview)];
  const selected = specialOrders[Math.max(0, Math.min(selectedIndex, specialOrders.length - 1))] || specialOrders[0];
  return {
    ...selected,
    customerCount: specialOrders.length,
    specialOrders,
  };
}

function stripSpecialOrders(preview) {
  const { specialOrders, ...singlePreview } = preview || {};
  return singlePreview;
}

function buildSpecialOrderIdentity(preview) {
  const reference = normalizeText(preview?.reference || "");
  const deliveryDate = normalizeText(preview?.deliveryDate || "");
  const fileName = normalizeText(preview?.sourceFileName || "");
  const itemSignature = (preview?.items || [])
    .map((item) => `${normalizeText(item.primary || "")}:${normalizeText(item.secondary || "")}:${normalizeText(item.quantity || "")}`)
    .join("|");
  return `${reference}|${deliveryDate}|${fileName}|${itemSignature}`;
}

function getSpecialDeliveryPoint(preview, fileName = "") {
  const itemText = (preview?.items || []).map((item) => `${item.primary || ""} ${item.secondary || ""}`).join(" ");
  const text = normalizeText(`${preview?.deliveryPoint || ""} ${preview?.customer || ""} ${preview?.reference || ""} ${preview?.fatrans || ""} ${fileName} ${itemText}`);
  if (text.includes("uien") || text.includes("onion")) return "HAVI DE UIEN";
  if (text.includes("neu wulmstorf")) return "Neu Wulmstorf";
  if (text.includes("duisburg")) return "Duisburg";
  if (text.includes("wunstorf")) return "Wunstorf";
  if (text.includes("havi nl") || text.includes("netherlands") || text.includes("nederland")) return "HAVI NL";
  if (text.includes("havi be") || text.includes("belg")) return "HAVI BE";
  if (text.includes("havi de") || text.includes("gmbh")) return "HAVI DE";
  return "HAVI";
}

function isHaviWorkspaceSelected() {
  const normalized = normalizeText(state.selectedClient);
  return normalized === "havi" || normalized === "havi duisburg saturday";
}

function isSelectedHaviDeUienOrder() {
  if (!state.preview || !state.orderSelectionActive) return false;

  if (state.mode === "special") {
    const specialPoint = getSpecialDeliveryPoint(state.preview, state.preview.sourceFileName || "");
    return normalizeText(specialPoint) === "havi de uien";
  }

  if (state.mode !== "standard" || !isHaviWorkspaceSelected()) return false;

  const selected = applyStandardClientDisplayRules(state.preview.orders?.[state.selectedIndex]);
  if (!selected) return false;

  const deliveryPoint = getDeliveryPointKey(selected);
  const combined = normalizeText(
    `${deliveryPoint} ${selected.customer || ""} ${selected.fatrans || ""} ${selected.label || ""} ${selected.reference || ""}`,
  );
  return normalizeText(deliveryPoint) === "havi de uien" || combined.includes("havi de uien") || combined.includes("uien");
}

function previewToWorkspaceOrders(mode, preview) {
  if (!preview) return [];
  if (mode === "special") return specialPreviewToOrders(preview);
  return Array.isArray(preview.orders) ? preview.orders : [];
}

function specialPreviewToOrders(preview) {
  const specialOrders = Array.isArray(preview.specialOrders) ? preview.specialOrders : [stripSpecialOrders(preview)];
  return specialOrders.map((order, index) => {
    const deliveryPoint = order.deliveryPoint || getSpecialDeliveryPoint(order, order.sourceFileName || "");
    const reference = getSpecialOrderReference(order, index);
    return {
      label: `${deliveryPoint} (${reference})`,
      customer: order.customer || "Havi Logistics GmbH",
      reference,
      fatrans: deliveryPoint,
      deliveryPoint,
      deliveryDate: order.deliveryDate || "",
      sourceFileName: order.sourceFileName || "",
      sourceOrderIndex: Number.isInteger(order.sourceOrderIndex) ? order.sourceOrderIndex : index,
      items: order.items || [],
    };
  });
}

function mergeWorkspaceOrders(existingOrders, incomingOrders) {
  const orders = [...existingOrders];
  let firstChangedIndex = orders.length;
  incomingOrders.forEach((order) => {
    const duplicateIndex = orders.findIndex((existing) => buildOrderIdentity(existing) === buildOrderIdentity(order));
    if (duplicateIndex >= 0) {
      orders[duplicateIndex] = order;
      if (firstChangedIndex === existingOrders.length) firstChangedIndex = duplicateIndex;
      return;
    }
    orders.push(order);
  });
  return {
    orders,
    firstChangedIndex: orders.length ? Math.min(firstChangedIndex, orders.length - 1) : 0,
  };
}

function buildOrderIdentity(order) {
  const reference = normalizeText(order?.reference || "");
  const customer = normalizeText(order?.customer || "");
  const fatrans = normalizeText(order?.fatrans || "");
  return reference ? `${reference}|${customer}|${fatrans}` : `${customer}|${fatrans}|${normalizeText(order?.label || "")}`;
}

async function handleSaveOrder() {
  if (!state.preview) return;
  syncCurrentClientWorkspace();
  const isSaturdayCarryover = Boolean(getSaturdayCarryoverConfig());
  const shortageMessage = isSaturdayCarryover ? "" : await saveOrdersForShortages(state.file);
  const carryoverMessage = await propagateSaturdayOrdersToNextSession(state.file);
  statusText.textContent = [
    isSaturdayCarryover
      ? `${state.selectedClient} orders saved to the next session.`
      : `${state.selectedClient} orders saved to the active session.`,
    shortageMessage,
    carryoverMessage,
  ].filter(Boolean).join(" ");
  if (saveOrderButton) saveOrderButton.disabled = false;
}

async function saveOrdersForShortages(file, options = {}) {
  const targetPreview = options.preview || buildPreviewForDataTransfer();
  const targetMode = options.mode || state.mode || "";
  const targetSession = options.session || state.activeWorkSession;
  const manageButton = options.manageButton !== false;
  if (!targetPreview) return "";
  if (!targetSession?.date || !targetSession?.id) {
    return "Create or open a work session so this order can appear in Manco´s.";
  }

  if (saveOrderButton) saveOrderButton.disabled = true;
  const formData = new FormData();
  if (file) {
    formData.append("email", file);
  }
  formData.append("mode", targetMode);
  formData.append("preview", JSON.stringify(targetPreview));
  formData.append("date", targetSession.date || "");
  formData.append("name", targetSession.name || "");
  formData.append("workSessionId", targetSession.id || "");
  formData.append("selectedClient", options.selectedClient || state.selectedClient || "");

  try {
    const response = await fetch("/api/orders/ingest", {
      method: "POST",
      credentials: "include",
      body: formData,
    });
    const payload = await response.json();
    if (response.status === 401) {
      resetPreviewState();
      showAuthScreen();
      return "Sign in again before saving to Manco´s.";
    }
    if (!response.ok) {
      return payload.error || "Could not save this order to Manco´s.";
    }
    const savedCount = Array.isArray(payload.savedSessions) ? payload.savedSessions.length : 0;
    return `${savedCount} order${savedCount === 1 ? "" : "s"} available in Manco´s.`;
  } catch {
    return "Could not reach Manco´s sync.";
  } finally {
    if (manageButton && saveOrderButton) saveOrderButton.disabled = false;
  }
}

async function propagateSaturdayOrdersToNextSession(file) {
  const config = getSaturdayCarryoverConfig();
  if (!config || !state.preview || !state.activeWorkSession?.date) return "";

  const nextDate = addDaysToDateValue(state.activeWorkSession.date, 1);
  const targetSession = ensureWorkSessionForDate(nextDate);
  const sourcePreview = buildPreviewForDataTransfer();
  const sourceOrders = previewToWorkspaceOrders(state.mode, sourcePreview)
    .map((order) => normalizeSaturdayCarryoverOrder(order, config));
  if (!sourceOrders.length) return "";

  targetSession.workspaces = migrateClientWorkspaces(targetSession.workspaces || {});
  const targetWorkspace = mergeOrdersIntoCarryoverWorkspace(
    targetSession.workspaces[config.targetClient] || createEmptyClientWorkspace(),
    sourceOrders,
    sourcePreview,
  );
  targetSession.workspaces[config.targetClient] = targetWorkspace;

  saveCarryoverLeverschemaResult(targetSession, config, sourceOrders);
  upsertTeamSessionAndPersist(targetSession);
  persistLeverschemaResults();
  await persistTeamStateImmediately();

  await saveOrdersForShortages(file, {
    session: targetSession,
    mode: targetWorkspace.mode,
    preview: targetWorkspace.preview,
    manageButton: false,
  });

  if (state.activeWorkSession?.id === targetSession.id) {
    state.clientWorkspaces = migrateClientWorkspaces(targetSession.workspaces || {});
    state.activeWorkSession.workspaces = state.clientWorkspaces;
  }
  renderSessionHistoryList();
  renderClientsLeverschemaCard();
  return `${config.sourceLabel} copied to ${config.targetClient} in ${targetSession.name}, with PP saved to Leverschema.`;
}

function getSaturdayCarryoverConfig() {
  const selected = normalizeText(state.selectedClient);
  if (selected === "havi duisburg saturday") {
    return {
      sourceLabel: "Havi Duisburg Saturday",
      targetClient: "HAVI",
      targetCustomer: "Havi Duisburg",
      targetFatrans: "Havi Duisburg",
      targetDeliveryPoint: "Havi Duisburg",
      masterKey: "havi_duisburg",
    };
  }
  if (selected === "colruyt saturday") {
    return {
      sourceLabel: "Colruyt Saturday",
      targetClient: "Colruyt",
      targetCustomer: "Colruyt",
      targetFatrans: "Colruyt",
      targetDeliveryPoint: "Colruyt",
      masterKey: "colruyt",
    };
  }
  return null;
}

function ensureWorkSessionForDate(dateValue) {
  const existing = state.teamSessions.find((session) => session.date === dateValue);
  if (existing) {
    existing.workspaces = migrateClientWorkspaces(existing.workspaces || {});
    return existing;
  }
  return {
    id: createLocalId(),
    createdBy: state.user?.email || "",
    date: dateValue,
    name: buildDefaultSessionName(dateValue),
    createdAt: new Date().toISOString(),
    workspaces: {},
  };
}

function createLocalId() {
  return typeof crypto !== "undefined" && typeof crypto.randomUUID === "function"
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function addDaysToDateValue(dateValue, days) {
  const parsed = parseInputDate(dateValue) || new Date();
  parsed.setDate(parsed.getDate() + days);
  return formatDateForInput(parsed);
}

function normalizeSaturdayCarryoverOrder(order, config) {
  const next = JSON.parse(JSON.stringify(order || {}));
  const currentPoint = getDeliveryPointKey(next);
  const keepHaviUien = config.targetClient === "HAVI" && isHaviUienDeliveryPoint(currentPoint);
  next.customer = keepHaviUien ? next.customer || "Havi Logistics GmbH" : config.targetCustomer;
  next.fatrans = keepHaviUien ? next.fatrans || "HAVI DE UIEN" : config.targetFatrans;
  next.deliveryPoint = keepHaviUien ? "HAVI DE UIEN" : config.targetDeliveryPoint;
  next.label = next.reference ? `${next.deliveryPoint} (${next.reference})` : next.deliveryPoint;
  next.items = (next.items || []).map((item) => ({
    ...item,
    quantity: parseNumber(item.quantity),
  }));
  return next;
}

function mergeOrdersIntoCarryoverWorkspace(workspace, incomingOrders, sourcePreview) {
  const existingOrders = workspace.mode === "standard"
    ? workspace.preview?.orders || []
    : previewToWorkspaceOrders(workspace.mode, workspace.preview);
  const merged = mergeWorkspaceOrders(existingOrders, incomingOrders);
  const orders = merged.orders;
  return {
    ...createEmptyClientWorkspace(),
    mode: "standard",
    preview: {
      deliveryDate: sourcePreview.deliveryDate || workspace.preview?.deliveryDate || "",
      customerCount: orders.length,
      orders,
      canMerge: orders.length > 0,
    },
    selectedIndex: Math.max(0, Math.min(merged.firstChangedIndex, orders.length - 1)),
    orderSelectionActive: false,
    activeDeliveryPointKey: null,
    leverschemaIncludedIndexes: incomingOrders
      .map((incoming) => orders.findIndex((order) => buildOrderIdentity(order) === buildOrderIdentity(incoming)))
      .filter((index) => index >= 0),
    exportSheet: workspace.exportSheet || "Paderborn",
    emailCacheId: null,
    emailFileName: workspace.emailFileName || "",
  };
}

function saveCarryoverLeverschemaResult(targetSession, config, sourceOrders) {
  const sheet = getSessionSheetForDate(targetSession.date);
  const definitions = MASTER_SHEETS[sheet] || [];
  const definition = definitions.find((entry) => entry.key === config.masterKey);
  const deliveryDate = definition ? formatWorkbookDate(definition.dateOffset, targetSession.date) : "";
  const order = {
    customer: config.targetCustomer,
    reference: sourceOrders.map((entry) => entry.reference).filter(Boolean).join(" + "),
    fatrans: config.targetFatrans,
    deliveryPoint: config.targetDeliveryPoint,
    deliveryDate,
    items: sourceOrders.flatMap((entry) =>
      (entry.items || []).map((item) => ({
        primary: item.primary,
        secondary: item.secondary,
        quantity: parseNumber(item.quantity),
        unit: item.unit,
      })),
    ),
  };
  if (!order.items.length) return;

  const result = calculateLeverschemaForClient(order, config.targetClient, targetSession);
  const deliveryKey = normalizeWorkbookDate(order.deliveryDate).replaceAll("/", "-") || "no-date";
  const storageKey = `${targetSession.id}::${sheet}::${config.masterKey}::${deliveryKey}`;
  state.leverschemaResults[storageKey] = {
    sessionId: targetSession.id,
    sessionDate: targetSession.date,
    sheet,
    customer: order.customer,
    reference: order.reference,
    deliveryDate: order.deliveryDate || "",
    masterKey: config.masterKey,
    ruleName: result.ruleName,
    totalPallets: result.totalPallets,
    totalPalletPlaces: result.totalPalletPlaces,
    rows: result.rows,
    includedIndexes: sourceOrders.map((_, index) => index),
    savedAt: new Date().toISOString(),
  };
}

function calculateLeverschemaForClient(order, client, session) {
  const previousClient = state.selectedClient;
  const previousMode = state.mode;
  const previousSession = state.activeWorkSession;
  try {
    state.selectedClient = client;
    state.mode = "standard";
    state.activeWorkSession = session;
    return calculateLeverschema(order);
  } finally {
    state.selectedClient = previousClient;
    state.mode = previousMode;
    state.activeWorkSession = previousSession;
  }
}

function getSessionSheetForDate(dateValue) {
  const sessionDate = parseInputDate(dateValue);
  if (!sessionDate) return "Monday-Thursday";
  const day = sessionDate.getDay();
  if (day === 5) return "Friday";
  if (day === 6) return "Saturday";
  return "Monday-Thursday";
}

function renderPreview() {
  if (!state.preview) {
    clearPreviewDisplay();
    updateOrdersClientMode();
    renderLeverschemaSummary();
    renderLeverschemaWorkbook();
    enableCMRButtonForSpecialClients();
    statusText.textContent = `No email loaded for ${state.selectedClient}.`;
    return;
  }

  customersFound.textContent = state.preview.customerCount;
  deliveryDate.textContent = state.preview.deliveryDate || "-";
  leverschemaButton.disabled = false;
  printCmrButton.disabled = false;
  if (saveOrderButton) saveOrderButton.disabled = false;

  if (state.mode === "special") {
    exportButton.disabled = true;
    mergeButton.disabled = true;
    specialButton.disabled = !state.orderSelectionActive;
    nettoMdButton.disabled = true;
    orderSelect.innerHTML = '<option value="0">Havi Logistics GmbH</option>';
    if (!state.orderSelectionActive) {
      clearOrderOverview();
    } else {
      setAddItemEnabled(true);
      detailCustomer.textContent = state.preview.customer;
      detailReference.textContent = state.preview.reference;
      detailFatrans.textContent = "HAVI DE UIEN";
      itemsInSelection.textContent = state.preview.items.length;
      setTable(
        ["DC", "Article", "Description", "Quantity", "Unit", ""],
        state.preview.items.map((item, index) => renderSpecialItemRow(item, index)),
      );
    }
    statusText.textContent = "Special HAVI DE UIEN email detected.";
  } else if (state.mode === "netto_md") {
    const orders = state.preview.orders;
    exportButton.disabled = true;
    mergeButton.disabled = true;
    specialButton.disabled = true;
    nettoMdButton.disabled = false;
    leverschemaButton.disabled = true;
    printCmrButton.disabled = true;
    orderSelect.innerHTML = orders.map((order, index) => `<option value="${index}">${escapeHtml(order.label)}</option>`).join("");
    orderSelect.value = String(state.selectedIndex);

    if (!state.orderSelectionActive) {
      clearOrderOverview();
    } else {
      setAddItemEnabled(true);
      const selected = orders[state.selectedIndex];
      detailCustomer.textContent = selected.customer;
      detailReference.textContent = selected.reference;
      detailFatrans.textContent = state.preview.dcName || selected.fatrans;
      itemsInSelection.textContent = selected.items.length;
      setTable(
        ["Article", "Description", "Quantity", "Unit", ""],
        selected.items.map((item, index) => renderStandardItemRow(item, index)),
      );
    }
    statusText.textContent = "NettoMD orderpicking email detected.";
  } else {
    const orders = state.preview.orders;
    syncLeverschemaIncludedIndexes();
    exportButton.disabled = !state.orderSelectionActive;
    mergeButton.disabled = !state.orderSelectionActive || !canMergeSelectedOrders();
    specialButton.disabled = !isSelectedHaviDeUienOrder();
    nettoMdButton.disabled = true;
    const normalizedClient = normalizeText(state.selectedClient);
    const shouldEnableCmr = isPrintCmrClient(normalizedClient);
    
    console.log("Print CMR Debug:", {
      selectedClient: state.selectedClient,
      normalizedClient: normalizedClient,
      shouldEnableCmr: shouldEnableCmr
    });
    
    printCmrButton.disabled = !shouldEnableCmr;
    orderSelect.innerHTML = orders.map((order, index) => `<option value="${index}">${escapeHtml(order.label)}</option>`).join("");
    orderSelect.value = String(state.selectedIndex);

    if (!state.orderSelectionActive) {
      clearOrderOverview();
      printCmrButton.disabled = !shouldEnableCmr;
    } else {
      setAddItemEnabled(true);
      const selected = applyStandardClientDisplayRules(orders[state.selectedIndex]);
      detailCustomer.textContent = selected.customer;
      detailReference.textContent = selected.reference;
      detailFatrans.textContent = selected.fatrans;
      itemsInSelection.textContent = selected.items.length;
      if (isHaviUienDeliveryPoint(getDeliveryPointKey(selected))) {
        setTable(
          ["DC", "Article", "Description", "Quantity", "Unit", ""],
          selected.items.map((item, index) => renderHaviUienStandardItemRow(item, index)),
        );
      } else {
        setTable(
          ["Article", "Description", "Quantity", "Unit", ""],
          selected.items.map((item, index) => renderStandardItemRow(item, index)),
        );
      }
    }
    statusText.textContent = "Preview ready.";
  }

  // Always call these functions to update UI state
  updateOrdersClientMode();
  enableCMRButtonForSpecialClients();
  // Only update Havi DC buttons if printCmr is enabled
  const preferences = getClientActionPreferences(state.selectedClient);
  const isHavi3DC = isHavi3DCClient(normalizeText(state.selectedClient));
  if (isHavi3DC && preferences.printCmr) {
    updateHaviDCButtonsVisibility();
  } else if (isHavi3DC && !preferences.printCmr) {
    // Hide all individual Havi DC buttons when printCmr is disabled
    printCmrDuisburgButton.classList.add("hidden");
    printCmrWunstorfButton.classList.add("hidden");
    printCmrNeuWulmstorfButton.classList.add("hidden");
  }

  renderLeverschemaSummary();
  renderLeverschemaIncludeOptions();
  renderDeliverySidebar();
  updateOrdersClientMode();
  updateLeverschemaPageVisibility();
  renderLeverschemaWorkbook();
}

function setTable(headers, rows) {
  tableHead.innerHTML = `<tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr>`;
  tableBody.innerHTML = rows
    .map(
      (row) =>
        `<tr>${row
          .map((value) => `<td>${isEditableTableControl(value) ? value : escapeHtml(String(value ?? ""))}</td>`)
          .join("")}</tr>`,
    )
    .join("");
}

function isEditableTableControl(value) {
  return typeof value === "string" && (
    value.includes("data-quantity-input") ||
    value.includes("data-item-field-input") ||
    value.includes("data-delete-item")
  );
}

function renderStandardItemRow(item, index) {
  return [
    renderItemFieldInput(index, "primary", item.primary, "Article"),
    renderItemFieldInput(index, "secondary", item.secondary, "Description"),
    renderQuantityInput(index, item.quantity),
    renderItemFieldInput(index, "unit", item.unit || "Collo", "Unit"),
    renderDeleteItemButton(index),
  ];
}

function renderHaviUienStandardItemRow(item, index) {
  const settings = getHaviUienSettings();
  return [
    buildHaviUienDcLabel(item),
    settings.article,
    settings.description,
    renderQuantityInput(index, item.quantity),
    renderItemFieldInput(index, "unit", item.unit || "cases", "Unit"),
    renderDeleteItemButton(index),
  ];
}

function renderSpecialItemRow(item, index) {
  const settings = getHaviUienSettings();
  return [
    buildHaviUienDcLabel(item),
    settings.article,
    settings.description,
    renderQuantityInput(index, item.quantity),
    renderItemFieldInput(index, "unit", item.unit || "Collo", "Unit"),
    renderDeleteItemButton(index),
  ];
}

function clearOrderOverview() {
  detailCustomer.textContent = "-";
  detailReference.textContent = "-";
  detailFatrans.textContent = "-";
  itemsInSelection.textContent = "0";
  tableHead.innerHTML = "";
  tableBody.innerHTML = "";
  setAddItemEnabled(false);
}

function renderQuantityInput(index, value) {
  return `<input class="quantity-input" data-quantity-input="1" data-item-index="${index}" type="number" min="0" step="1" value="${escapeHtml(String(value ?? ""))}">`;
}

function renderItemFieldInput(index, field, value, label, type = "text") {
  const numericAttributes = type === "number" ? ' min="0" step="1"' : "";
  return `<input class="item-field-input" data-item-field-input="1" data-item-index="${index}" data-item-field="${escapeHtml(field)}" type="${type}"${numericAttributes} value="${escapeHtml(String(value ?? ""))}" aria-label="${escapeHtml(label)}">`;
}

function renderDeleteItemButton(index) {
  return `<button class="delete-item-button" type="button" data-delete-item="1" data-item-index="${index}" aria-label="Delete item">Delete</button>`;
}

function isHaviUienDeliveryPoint(value) {
  return normalizeText(value || "") === "havi de uien";
}

function buildHaviUienDcLabel(item) {
  return [item?.primary, item?.secondary]
    .map((value) => String(value || "").trim())
    .filter(Boolean)
    .join(" - ");
}

function mapHaviUienManualItem(item) {
  const settings = getHaviUienSettings();
  return {
    primary: settings.article,
    secondary: settings.description,
    dc: buildHaviUienDcLabel(item),
    quantity: parseNumber(item.quantity),
    unit: item.unit,
  };
}

function buildPreviewForDataTransfer() {
  if (!state.preview) return state.preview;
  const preview = JSON.parse(JSON.stringify(state.preview));
  if (state.mode !== "standard" || !Array.isArray(preview.orders)) {
    return preview;
  }
  preview.orders = preview.orders.map((rawOrder) => {
    const order = applyStandardClientDisplayRules(rawOrder);
    if (!isHaviUienDeliveryPoint(getDeliveryPointKey(order))) {
      return rawOrder;
    }
    return {
      ...rawOrder,
      deliveryPoint: "HAVI DE UIEN",
      items: (rawOrder.items || []).map((item) => mapHaviUienManualItem(item)),
    };
  });
  return preview;
}

function handleQuantityEdit(event) {
  const input = event.target.closest("[data-quantity-input]");
  if (!input || !state.preview) return;
  const itemIndex = Number(input.dataset.itemIndex);
  const quantity = sanitizeEditedQuantity(input.value);

  if (state.mode === "special") {
    if (state.preview.items[itemIndex]) {
      state.preview.items[itemIndex].quantity = quantity;
    }
  } else {
    const order = state.preview.orders?.[state.selectedIndex];
    if (order?.items?.[itemIndex]) {
      order.items[itemIndex].quantity = quantity;
    }
  }

  input.value = quantity;
  syncCurrentClientWorkspace();
  itemsInSelection.textContent = getCurrentOrder()?.items.length || "0";
  renderLeverschemaSummary();
  renderLeverschemaWorkbook();
}

function handleItemFieldEdit(event) {
  const input = event.target.closest("[data-item-field-input]");
  if (!input || !state.preview) return;

  const itemIndex = Number(input.dataset.itemIndex);
  const field = input.dataset.itemField;
  const item = getSelectedOrderItems()?.[itemIndex];
  if (!item || !field) return;

  item[field] = input.type === "number" ? sanitizeEditedQuantity(input.value) : input.value.trim();
  input.value = item[field];
  syncCurrentClientWorkspace();
  renderLeverschemaSummary();
  renderLeverschemaWorkbook();
}

function handleAddItem() {
  const items = getSelectedOrderItems();
  if (!items || !state.orderSelectionActive) return;

  const newItem = state.mode === "special"
    ? { primary: "", secondary: "", slicesQuantity: 0, quantity: 0, unit: "Collo" }
    : { primary: "", secondary: "", quantity: 0, unit: "Collo" };
  items.push(newItem);
  syncCurrentClientWorkspace();
  renderPreview();

  const nextInput = tableBody.querySelector(`[data-item-index="${items.length - 1}"][data-item-field-input]`);
  nextInput?.focus();
}

function handleDeleteItem(event) {
  const button = event.target.closest("[data-delete-item]");
  if (!button || !state.preview) return;

  const items = getSelectedOrderItems();
  const itemIndex = Number(button.dataset.itemIndex);
  if (!items || itemIndex < 0 || itemIndex >= items.length) return;

  items.splice(itemIndex, 1);
  syncCurrentClientWorkspace();
  renderPreview();
  itemsInSelection.textContent = getCurrentOrder()?.items.length || "0";
  renderLeverschemaSummary();
  renderLeverschemaWorkbook();
}

function getSelectedOrderItems() {
  if (!state.preview) return null;
  if (state.mode === "special") return state.preview.items || null;
  return state.preview.orders?.[state.selectedIndex]?.items || null;
}

function setAddItemEnabled(enabled) {
  if (addItemButton) addItemButton.disabled = !enabled;
}

function handleLeverschemaIncludeChange(event) {
  const checkbox = event.target.closest("[data-leverschema-order-index]");
  if (!checkbox || !state.preview || state.mode === "special" || state.mode === "netto_md") return;
  const index = Number(checkbox.dataset.leverschemaOrderIndex);
  const checked = checkbox.checked;
  const next = new Set(state.leverschemaIncludedIndexes);
  if (checked) {
    next.add(index);
  } else {
    next.delete(index);
  }
  state.leverschemaIncludedIndexes = Array.from(next).sort((left, right) => left - right);
  syncCurrentClientWorkspace();
  renderLeverschemaSummary();
}

function handleDeliveryPointClick(event) {
  const button = event.target.closest("[data-delivery-point]");
  if (!button || !state.preview) return;

  if (state.mode === "special") {
    const group = getSpecialDeliveryPointGroups().find((entry) => entry.key === button.dataset.deliveryPoint);
    if (!group?.orders?.length) return;
    state.activeDeliveryPointKey = state.activeDeliveryPointKey === group.key ? null : group.key;
    state.orderSelectionActive = false;
    syncCurrentClientWorkspace();
    renderPreview();
    return;
  }

  if (!state.preview?.orders?.length) return;

  const group = getDeliveryPointGroups().find((entry) => entry.key === button.dataset.deliveryPoint);
  if (!group?.orders?.length) return;

  state.activeDeliveryPointKey = state.activeDeliveryPointKey === group.key ? null : group.key;
  state.orderSelectionActive = false;
  syncCurrentClientWorkspace();
  renderPreview();
}

function handleDeliveryReferenceToggle(event) {
  const checkbox = event.target.closest("[data-delivery-order-index]");
  if (!checkbox || !state.preview || state.mode === "netto_md") return;

  const index = Number(checkbox.dataset.deliveryOrderIndex);
  const next = new Set(state.leverschemaIncludedIndexes);
  if (checkbox.checked) {
    next.add(index);
  } else {
    next.delete(index);
  }
  state.leverschemaIncludedIndexes = Array.from(next).sort((left, right) => left - right);
  syncCurrentClientWorkspace();
  renderPreview();
}

function handleDeliveryReferenceClick(event) {
  const button = event.target.closest("[data-delivery-reference-index]");
  if (!button || !state.preview) return;

  state.selectedIndex = Number(button.dataset.deliveryReferenceIndex);
  state.orderSelectionActive = true;
  state.activeDeliveryPointKey = getDeliveryPointKeyForIndex(state.selectedIndex);
  if (state.mode === "special") {
    state.preview = selectSpecialPreview(state.preview, state.selectedIndex);
  }
  syncCurrentClientWorkspace();
  renderPreview();
}

async function handleDeliveryOrderDelete(event) {
  const button = event.target.closest("[data-delete-delivery-order-index]");
  if (!button || !state.preview) return;
  event.preventDefault();
  event.stopPropagation();

  const index = Number(button.dataset.deleteDeliveryOrderIndex);
  const order = getWorkspaceOrderAtIndex(index);
  if (!order) return;
  const reference = getOrderReferenceForRemoval(order, index);
  const shouldDelete = await showConfirmationDialog({
    eyebrow: "Delete order",
    title: "Delete this order?",
    message: `${reference} will be removed from this card and from Manco's.`,
    confirmLabel: "Delete order",
    cancelLabel: "Cancel",
  });
  if (!shouldDelete) return;

  const removalIdentity = buildMancoRemovalIdentity(order, index);
  removeWorkspaceOrderAtIndex(index);
  syncCurrentClientWorkspace();
  await persistTeamStateImmediately();
  renderPreview();
  renderLeverschemaSummary();
  renderLeverschemaWorkbook();

  const mancoMessage = await removeOrderFromMancos(removalIdentity);
  statusText.textContent = [`${reference} deleted.`, mancoMessage].filter(Boolean).join(" ");
}

function getWorkspaceOrderAtIndex(index) {
  if (!state.preview || !Number.isInteger(index) || index < 0) return null;
  if (state.mode === "special") {
    const entries = getSpecialOrderEntries();
    return entries[index]?.order || null;
  }
  return state.preview.orders?.[index] || null;
}

function removeWorkspaceOrderAtIndex(index) {
  if (state.mode === "special") {
    const specialOrders = Array.isArray(state.preview.specialOrders)
      ? [...state.preview.specialOrders]
      : [stripSpecialOrders(state.preview)];
    specialOrders.splice(index, 1);
    if (!specialOrders.length) {
      state.preview = null;
      state.mode = null;
      state.selectedIndex = 0;
      state.orderSelectionActive = false;
      state.activeDeliveryPointKey = null;
      state.leverschemaIncludedIndexes = [];
      return;
    }
    state.selectedIndex = Math.min(state.selectedIndex, specialOrders.length - 1);
    state.preview = selectSpecialPreview({ ...state.preview, specialOrders }, state.selectedIndex);
    state.leverschemaIncludedIndexes = remapIndexesAfterRemoval(state.leverschemaIncludedIndexes, index);
    return;
  }

  const orders = Array.isArray(state.preview?.orders) ? [...state.preview.orders] : [];
  orders.splice(index, 1);
  if (!orders.length) {
    state.preview = null;
    state.mode = null;
    state.selectedIndex = 0;
    state.orderSelectionActive = false;
    state.activeDeliveryPointKey = null;
    state.leverschemaIncludedIndexes = [];
    return;
  }
  state.preview = {
    ...state.preview,
    orders,
    customerCount: orders.length,
    canMerge: orders.length > 0,
  };
  state.selectedIndex = Math.min(state.selectedIndex >= index ? Math.max(state.selectedIndex - 1, 0) : state.selectedIndex, orders.length - 1);
  state.leverschemaIncludedIndexes = remapIndexesAfterRemoval(state.leverschemaIncludedIndexes, index);
  const selectedKey = getDeliveryPointKeyForIndex(state.selectedIndex);
  const groupStillExists = getDeliveryPointGroups().some((group) => group.key === state.activeDeliveryPointKey);
  state.activeDeliveryPointKey = groupStillExists ? state.activeDeliveryPointKey : selectedKey;
}

function remapIndexesAfterRemoval(indexes, removedIndex) {
  return [...new Set((indexes || [])
    .filter((index) => index !== removedIndex)
    .map((index) => (index > removedIndex ? index - 1 : index)))]
    .sort((left, right) => left - right);
}

function getOrderReferenceForRemoval(order, index) {
  if (state.mode === "special") return getSpecialOrderReference(order, index);
  return String(order?.reference || order?.label || `Order ${index + 1}`).trim();
}

function buildMancoRemovalIdentity(order, index) {
  const displayOrder = state.mode === "standard" ? applyStandardClientDisplayRules(order) : order;
  const reference = getOrderReferenceForRemoval(order, index);
  const hasReference = Boolean(reference && !/^order\s+\d+$/i.test(reference));
  return {
    date: state.activeWorkSession?.date || "",
    workSessionId: state.activeWorkSession?.id || "",
    reference,
    customer: hasReference ? "" : String(order?.customer || displayOrder?.customer || "").trim(),
    fatrans: hasReference ? "" : String(order?.fatrans || displayOrder?.fatrans || "").trim(),
    deliveryPoint: hasReference ? "" : getDeliveryPointKey(displayOrder),
  };
}

async function removeOrderFromMancos(identity) {
  if (!identity?.workSessionId && !identity?.date) return "";
  try {
    const response = await fetch("/api/orders/remove", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      credentials: "include",
      body: JSON.stringify(identity),
    });
    const payload = await response.json().catch(() => ({}));
    if (response.status === 401) {
      resetPreviewState();
      showAuthScreen();
      return "Sign in again to update Manco's.";
    }
    if (!response.ok) {
      return payload.error || "Could not remove this order from Manco's.";
    }
    const removed = Number(payload.removed || 0);
    return removed ? "Removed from Manco's." : "No matching Manco's order was found.";
  } catch {
    return "Could not reach Manco's sync.";
  }
}

function renderDeliverySidebar() {
  if (!deliveryPointList || !deliveryReferenceList) return;

  if (!state.preview) {
    deliveryPointList.innerHTML = `<p class="delivery-empty">Load an email to see delivery points.</p>`;
    deliveryReferenceList.innerHTML = "";
    return;
  }

  if (state.mode === "special") {
    renderSpecialDeliverySidebar();
    return;
  }

  const groups = getDeliveryPointGroups();
  if (!groups.length) {
    deliveryPointList.innerHTML = `<p class="delivery-empty">No delivery points found in this email.</p>`;
    deliveryReferenceList.innerHTML = "";
    return;
  }

  const activeGroup = groups.find((entry) => entry.key === state.activeDeliveryPointKey) || null;

  deliveryPointList.innerHTML = groups
    .map((group) => renderDeliveryPointAccordion(group, activeGroup, renderStandardDeliveryReference))
    .join("");
  deliveryReferenceList.innerHTML = "";
}

function renderDeliveryPointAccordion(group, activeGroup, renderReference) {
  const isActive = activeGroup ? group.key === activeGroup.key : false;
  return `
    <div class="delivery-point-group ${isActive ? "active" : ""}">
      <button class="delivery-point-button ${isActive ? "active" : ""}" type="button" data-delivery-point="${escapeHtml(group.key)}">
        <span class="delivery-point-name">${escapeHtml(group.label)}</span>
        <span class="delivery-point-meta">
          <strong>${group.orders.length}</strong>
          <span class="delivery-point-arrow" aria-hidden="true">&gt;</span>
        </span>
      </button>
      ${isActive ? `<div class="delivery-reference-list inline">${group.orders.map(renderReference).join("")}</div>` : ""}
    </div>
  `;
}

function getDeliveryPointKeyForIndex(index) {
  if (state.mode === "special") {
    const entry = getSpecialOrderEntries()[index];
    return entry ? getSpecialDeliveryPoint(entry.order, entry.order.sourceFileName || "") : null;
  }
  const order = state.preview?.orders?.[index];
  return order ? getDeliveryPointKey(state.mode === "standard" ? applyStandardClientDisplayRules(order) : order) : null;
}

function renderStandardDeliveryReference({ order, index }) {
  const isChecked = state.leverschemaIncludedIndexes.includes(index);
  const reference = order.reference || order.label || `Order ${index + 1}`;
  return `
    <div class="delivery-reference-option ${index === state.selectedIndex ? "active" : ""}">
      <input type="checkbox" data-delivery-order-index="${index}" ${isChecked ? "checked" : ""} aria-label="Include ${escapeHtml(reference)}">
      <button type="button" data-delivery-reference-index="${index}">
        <strong>${escapeHtml(reference)}</strong>
        <span>${escapeHtml(`${order.items?.length || 0} items`)}</span>
      </button>
      <button class="delivery-reference-delete" type="button" data-delete-delivery-order-index="${index}" aria-label="Delete ${escapeHtml(reference)}">Delete</button>
    </div>
  `;
}

function getDeliveryPointGroups() {
  const orders = state.preview?.orders || [];
  const groups = new Map();
  orders.forEach((rawOrder, index) => {
    const order = state.mode === "standard" ? applyStandardClientDisplayRules(rawOrder) : rawOrder;
    const key = getDeliveryPointKey(order);
    if (!groups.has(key)) {
      groups.set(key, { key, label: key, orders: [] });
    }
    groups.get(key).orders.push({ order, index });
  });
  return Array.from(groups.values());
}

function renderSpecialDeliverySidebar() {
  const groups = getSpecialDeliveryPointGroups();
  if (!groups.length) {
    deliveryPointList.innerHTML = `<p class="delivery-empty">No delivery points found in this email.</p>`;
    deliveryReferenceList.innerHTML = "";
    return;
  }

  const activeGroup = groups.find((entry) => entry.key === state.activeDeliveryPointKey) || null;

  deliveryPointList.innerHTML = groups
    .map((group) => renderDeliveryPointAccordion(group, activeGroup, renderSpecialDeliveryReference))
    .join("");
  deliveryReferenceList.innerHTML = "";
}

function renderSpecialDeliveryReference({ order, index }) {
  const isChecked = state.leverschemaIncludedIndexes.includes(index);
  const reference = getSpecialOrderReference(order, index);
  return `
    <div class="delivery-reference-option ${index === state.selectedIndex ? "active" : ""}">
      <input type="checkbox" data-delivery-order-index="${index}" ${isChecked ? "checked" : ""} aria-label="Include ${escapeHtml(reference)}">
      <button type="button" data-delivery-reference-index="${index}">
        <strong>${escapeHtml(reference)}</strong>
        <span>${escapeHtml(`${order.items?.length || 0} items`)}</span>
      </button>
      <button class="delivery-reference-delete" type="button" data-delete-delivery-order-index="${index}" aria-label="Delete ${escapeHtml(reference)}">Delete</button>
    </div>
  `;
}

function getSpecialDeliveryPointGroups() {
  const groups = new Map();
  getSpecialOrderEntries().forEach(({ order, index }) => {
    const key = getSpecialDeliveryPoint(order, order.sourceFileName || "");
    if (!groups.has(key)) {
      groups.set(key, { key, label: key, orders: [] });
    }
    groups.get(key).orders.push({ order, index });
  });
  return Array.from(groups.values());
}

function getSpecialOrderEntries() {
  const specialOrders = Array.isArray(state.preview?.specialOrders)
    ? state.preview.specialOrders
    : state.preview
      ? [stripSpecialOrders(state.preview)]
      : [];
  return specialOrders.map((order, index) => ({ order, index }));
}

function getSpecialOrderReference(order, index) {
  const reference = String(order?.reference || "").trim();
  if (reference && normalizeText(reference) !== "multiple vcso references") return reference;
  const itemReference = (order?.items || [])
    .map((item) => String(item.secondary || "").trim())
    .find((value) => /^VCSO/i.test(value));
  return itemReference || `HAVI order ${index + 1}`;
}

function getDeliveryPointKey(order) {
  const customer = String(order?.customer || "");
  const fatrans = String(order?.fatrans || "");
  const label = String(order?.label || "");
  const explicitDeliveryPoint = String(order?.deliveryPoint || "").trim();
  if (explicitDeliveryPoint && !isGenericClientDeliveryPoint(explicitDeliveryPoint)) {
    return explicitDeliveryPoint;
  }
  const combined = normalizeText(`${customer} ${fatrans} ${label} ${explicitDeliveryPoint}`);
  const normalizedClient = normalizeText(state.selectedClient);

  if (normalizedClient === "havi" || normalizedClient === "havi duisburg saturday") {
    if (combined.includes("uien") || combined.includes("onion")) return "HAVI DE UIEN";
    if (combined.includes("havi nl") || combined.includes("netherlands") || combined.includes("nederland")) return "HAVI NL";
    if (combined.includes("havi be") || combined.includes("belg")) return "HAVI BE";
    if (combined.includes("havi de") || combined.includes("gmbh")) return "HAVI DE";
  }

  if (normalizedClient === "rewe" || normalizedClient === "penny") {
    const customerPoint = cleanDeliveryPointName(customer);
    if (customerPoint && !isGenericClientDeliveryPoint(customerPoint)) {
      return customerPoint;
    }
  }

  if (normalizedClient === "denemark" || normalizeText(fatrans) === "denemark") {
    const customerPoint = cleanDeliveryPointName(customer || label);
    if (customerPoint && !isGenericClientDeliveryPoint(customerPoint)) {
      return customerPoint;
    }
  }

  const aliases = [
    ["fif", "FIF"],
    ["kdc", "KDC"],
    ["duisburg", "Duisburg"],
    ["wunstorf", "Wunstorf"],
    ["neu wulmstorf", "Neu Wulmstorf"],
    ["kerpen", "Kerpen"],
    ["hodenhagen", "Hodenhagen"],
    ["henstedt", "Henstedt-Ulzburg"],
    ["ganderkesee", "Ganderkesee"],
    ["bottrop", "Bottrop"],
    ["krefeld", "Krefeld"],
    ["hamm", "Hamm"],
  ];
  const matched = aliases.find(([needle]) => combined.includes(needle));
  if (matched) return matched[1];

  const cleanedFatrans = fatrans.trim();
  if (cleanedFatrans && !isGenericClientDeliveryPoint(cleanedFatrans)) {
    const firstToken = cleanedFatrans.split(/\s+/)[0];
    if (/^[A-Z0-9]{2,5}$/.test(firstToken)) return firstToken;
    return cleanedFatrans;
  }
  const cleanedCustomer = cleanDeliveryPointName(customer);
  if (cleanedCustomer && !isGenericClientDeliveryPoint(cleanedCustomer)) return cleanedCustomer;
  return cleanedCustomer || state.selectedClient || "Delivery point";
}

function cleanDeliveryPointName(value) {
  return String(value || "")
    .replace(/^(rewe|penny|carrefour|havi|netto|edeka)\s+/i, "")
    .trim();
}

function isGenericClientDeliveryPoint(value) {
  const normalized = normalizeText(value);
  return ["", "rewe", "penny", "rewe penny", "rewe/penny", "carrefour", "havi", "netto", "edeka"].includes(normalized);
}

function ensureSelectedDeliveryPointIndexes(group) {
  const groupIndexes = group.orders.map(({ index }) => index);
  const hasSelected = state.leverschemaIncludedIndexes.some((index) => groupIndexes.includes(index));
  if (!hasSelected) {
    state.leverschemaIncludedIndexes = [...groupIndexes];
  }
}

function canMergeSelectedOrders() {
  if (state.mode !== "standard" || !state.preview?.orders?.length) return false;
  return getSelectedExportOrderIndexes().length > 0;
}

function getSelectedExportOrderIndexes() {
  const orders = state.preview?.orders || [];
  if (!orders.length) return [];
  const eligibleIndexes = getLeverschemaEligibleOrders().map(({ index }) => index);
  const selected = state.leverschemaIncludedIndexes
    .filter((index) => Number.isInteger(index) && index >= 0 && index < orders.length)
    .filter((index) => !eligibleIndexes.length || eligibleIndexes.includes(index));
  return selected.length
    ? selected
    : [state.selectedIndex].filter((index) => index >= 0 && index < orders.length);
}

function getCurrentOrder() {
  if (!state.preview) return null;

  if (state.mode === "special") {
    return {
      customer: state.preview.customer,
      reference: state.preview.reference,
      fatrans: "HAVI DE UIEN",
      deliveryPoint: state.preview.deliveryPoint || getSpecialDeliveryPoint(state.preview, state.preview.sourceFileName || ""),
      deliveryDate: state.preview.deliveryDate || "",
      items: state.preview.items.map((item) => ({
        primary: item.primary,
        secondary: item.secondary,
        quantity: parseNumber(item.quantity),
        unit: item.unit,
        slicesQuantity: parseNumber(item.slicesQuantity),
      })),
    };
  }

  if (state.mode === "netto_md") {
    const selectedNetto = state.preview.orders?.[state.selectedIndex];
    if (!selectedNetto) return null;

    return {
      customer: selectedNetto.customer,
      reference: selectedNetto.reference,
      fatrans: state.preview.dcName || selectedNetto.fatrans,
      deliveryDate: state.preview.deliveryDate || "",
      items: selectedNetto.items.map((item) => ({
        primary: item.primary,
        secondary: item.secondary,
        quantity: parseNumber(item.quantity),
        unit: item.unit,
      })),
    };
  }

  const selected = applyStandardClientDisplayRules(state.preview.orders?.[state.selectedIndex]);
  if (!selected) return null;
  const isHaviUien = isHaviUienDeliveryPoint(getDeliveryPointKey(selected));

  return {
    customer: selected.customer,
    reference: selected.reference,
    fatrans: selected.fatrans,
    deliveryPoint: getDeliveryPointKey(selected),
    deliveryDate: selected.deliveryDate || state.preview.deliveryDate || "",
    items: selected.items.map((item) => (
      isHaviUien
        ? mapHaviUienManualItem(item)
        : {
            primary: item.primary,
            secondary: item.secondary,
            quantity: parseNumber(item.quantity),
            unit: item.unit,
          }
    )),
  };
}

function getLeverschemaOrder() {
  if (!state.preview) return null;
  if (state.mode === "special") {
    return getCurrentOrder();
  }

  const orders = state.preview.orders || [];
  const rawSelected = orders[state.selectedIndex];
  if (!rawSelected) return null;

  const selected = applyStandardClientDisplayRules(rawSelected);
  const selectedDeliveryPoint = getDeliveryPointKey(selected);

  const includedOrders = getLeverschemaEligibleOrders()
    .filter(({ index }) => state.leverschemaIncludedIndexes.includes(index))
    .map(({ order }) => order);

  if (!includedOrders.length) {
    return {
      customer: selected.customer,
      reference: "",
      fatrans: selected.fatrans,
      deliveryPoint: selectedDeliveryPoint,
      deliveryDate: selected.deliveryDate || state.preview.deliveryDate || "",
      items: [],
    };
  }

  return {
    customer: selected.customer,
    reference: includedOrders.map((order) => order.reference).filter(Boolean).join(" + "),
    fatrans: selected.fatrans,
    deliveryPoint: selectedDeliveryPoint,
    deliveryDate: selected.deliveryDate || state.preview.deliveryDate || "",
    items: includedOrders.flatMap((order) =>
      order.items.map((item) => ({
        primary: item.primary,
        secondary: item.secondary,
        quantity: parseNumber(item.quantity),
        unit: item.unit,
      })),
    ),
  };
}

function getLeverschemaEligibleOrders() {
  if (state.mode === "special" || state.mode === "netto_md" || !state.preview?.orders?.length) {
    if (state.mode === "special") {
      return getSpecialOrderEntries().map(({ order, index }) => ({
        order: {
          customer: order.customer,
          reference: getSpecialOrderReference(order, index),
          fatrans: "HAVI DE UIEN",
          deliveryPoint: order.deliveryPoint || getSpecialDeliveryPoint(order, order.sourceFileName || ""),
          deliveryDate: order.deliveryDate || "",
          items: (order.items || []).map((item) => ({
            primary: item.primary,
            secondary: item.secondary,
            quantity: parseNumber(item.quantity),
            unit: item.unit,
            slicesQuantity: parseNumber(item.slicesQuantity),
          })),
        },
        index,
      }));
    }
    return [];
  }
  const selected = state.preview.orders[state.selectedIndex];
  if (!selected) return [];
  const selectedDeliveryPoint = getDeliveryPointKey(applyStandardClientDisplayRules(selected));
  return state.preview.orders
    .map((order, index) => ({ order: applyStandardClientDisplayRules(order), index }))
    .filter(({ order }) => getDeliveryPointKey(order) === selectedDeliveryPoint)
    .map(({ order, index }) => ({
      order: isHaviUienDeliveryPoint(selectedDeliveryPoint)
        ? {
            ...order,
            items: (order.items || []).map((item) => mapHaviUienManualItem(item)),
          }
        : order,
      index,
    }));
}

function syncLeverschemaIncludedIndexes() {
  if (state.mode === "special") {
    state.leverschemaIncludedIndexes = [0];
    return;
  }
  const eligibleIndexes = getLeverschemaEligibleOrders().map(({ index }) => index);
  const kept = state.leverschemaIncludedIndexes.filter((index) => eligibleIndexes.includes(index));
  if (kept.length) {
    state.leverschemaIncludedIndexes = kept;
    return;
  }
  state.leverschemaIncludedIndexes = eligibleIndexes.includes(state.selectedIndex) ? [state.selectedIndex] : eligibleIndexes.slice(0, 1);
}

function renderLeverschemaIncludeOptions() {
  if (state.mode === "special" || state.mode === "netto_md" || isNettoMdClientSelected()) {
    leverschemaIncludeCard.classList.add("hidden");
    leverschemaIncludeList.innerHTML = "";
    return;
  }

  const eligibleOrders = getLeverschemaEligibleOrders();
  const showCard = eligibleOrders.length > 1;
  leverschemaIncludeCard.classList.toggle("hidden", !showCard);
  if (!showCard) {
    leverschemaIncludeList.innerHTML = "";
    return;
  }

  leverschemaIncludeList.innerHTML = eligibleOrders
    .map(
      ({ order, index }) => `
        <label class="leverschema-include-option">
          <input type="checkbox" data-leverschema-order-index="${index}" ${state.leverschemaIncludedIndexes.includes(index) ? "checked" : ""}>
          <div>
            <strong>${escapeHtml(order.reference || order.label || `Order ${index + 1}`)}</strong>
            <span>${escapeHtml(`${order.items.length} items`)}</span>
          </div>
        </label>
      `,
    )
    .join("");
}

function saveLeverschemaResult() {
  const order = getLeverschemaOrder();
  if (!order) return;

  const sheet = getActiveSessionSheet();
  const result = calculateLeverschema(order);
  const storageKey = buildLeverschemaOrderKey(order, sheet);
  const masterKey = resolveMasterCustomerKey(order.customer, result.ruleName, buildLeverschemaRoutingContext(order));

  state.leverschemaResults[storageKey] = {
    sessionId: getCurrentLeverschemaSessionId(),
    sessionDate: getCurrentLeverschemaBaseDate(),
    sheet,
    customer: order.customer,
    reference: order.reference,
    deliveryDate: order.deliveryDate || "",
    masterKey,
    ruleName: result.ruleName,
    totalPallets: result.totalPallets,
    totalPalletPlaces: result.totalPalletPlaces,
    rows: result.rows,
    includedIndexes: [...state.leverschemaIncludedIndexes],
    savedAt: new Date().toISOString(),
  };

  persistLeverschemaResults();
  syncCurrentClientWorkspace();
  renderLeverschemaSummary();
  updateLeverschemaPageVisibility();
  renderLeverschemaWorkbook();
  statusText.textContent = `${order.customer} saved to Leverschema for ${sheet}.`;
}

function renderLeverschemaSummary() {
  const order = getLeverschemaOrder();
  if (!order) {
    leverschemaSavedSheet.textContent = "-";
    leverschemaRule.textContent = "-";
    leverschemaTotal.textContent = "-";
    leverschemaPlaces.textContent = "-";
    leverschemaNote.textContent = "Load an email and click Leverschema to save the result for the selected sheet.";
    leverschemaButton.disabled = true;
    return;
  }

  if (state.mode === "netto_md") {
    leverschemaSavedSheet.textContent = "-";
    leverschemaRule.textContent = "Not available";
    leverschemaTotal.textContent = "-";
    leverschemaPlaces.textContent = "-";
    leverschemaNote.textContent = "Leverschema is not used for NettoMD Orderpicking exports.";
    return;
  }

  const sheet = getActiveSessionSheet();
  const storageKey = buildLeverschemaOrderKey(order, sheet);
  const saved = state.leverschemaResults[storageKey];
  if (!order.items.length) {
    leverschemaSavedSheet.textContent = sheet;
    leverschemaRule.textContent = "-";
    leverschemaTotal.textContent = "-";
    leverschemaPlaces.textContent = "-";
    leverschemaNote.textContent = "Choose at least one order to include in the Leverschema count.";
    return;
  }

  if (!saved) {
    const previewResult = calculateLeverschema(order);
    leverschemaSavedSheet.textContent = sheet;
    leverschemaRule.textContent = previewResult.ruleName;
    leverschemaTotal.textContent = formatNumber(previewResult.totalPallets);
    leverschemaPlaces.textContent = formatNumber(previewResult.totalPalletPlaces);
    leverschemaNote.textContent = "Preview shown for the selected sheet. Click Leverschema to save this result for later Excel export.";
    return;
  }

  leverschemaSavedSheet.textContent = saved.sheet;
  leverschemaRule.textContent = saved.ruleName;
  leverschemaTotal.textContent = formatNumber(saved.totalPallets);
  leverschemaPlaces.textContent = formatNumber(saved.totalPalletPlaces);
  leverschemaNote.textContent = `Saved for later export on ${formatSavedAt(saved.savedAt)}.`;
}

async function handleStockUpload(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  stockStatusText.textContent = "Reading Stock workbook...";
  exportStockButton.disabled = true;

  const formData = new FormData();
  formData.append("file", file);
  try {
    const response = await fetch("/api/stock/parse", {
      method: "POST",
      credentials: "include",
      body: formData,
    });

    if (response.status === 401) {
      showAuthScreen();
      return;
    }

    const payload = await response.json();
    if (!response.ok) {
      stockStatusText.textContent = payload.error || "Could not read the Stock workbook.";
      state.stockItems = [];
      renderStockTable();
      return;
    }

    state.stockItems = Array.isArray(payload.items) ? payload.items : [];
    stockStatusText.textContent = `${state.stockItems.length} Stock rows loaded.`;
    renderStockTable();
  } catch (error) {
    console.error("Stock upload error:", error);
    stockStatusText.textContent = "Could not read the Stock workbook.";
    state.stockItems = [];
    renderStockTable();
  } finally {
    event.target.value = "";
  }
}

function renderStockTable() {
  if (!stockTableBody) return;
  const items = Array.isArray(state.stockItems) ? state.stockItems : [];
  stockRowsCount.textContent = String(items.length);
  exportStockButton.disabled = items.length === 0;

  if (!items.length) {
    stockTableBody.innerHTML = `
      <tr>
        <td colspan="4" class="empty-table-message">Upload an Excel file to show Stock rows.</td>
      </tr>
    `;
    return;
  }

  stockTableBody.innerHTML = items
    .map((item) => `
      <tr>
        <td>${escapeHtml(item.itemNumber)}</td>
        <td>${escapeHtml(item.productName)}</td>
        <td>${escapeHtml(item.quantity)}</td>
        <td>${escapeHtml(item.tht)}</td>
      </tr>
    `)
    .join("");
}

async function exportStockWorkbook() {
  const items = Array.isArray(state.stockItems) ? state.stockItems : [];
  if (!items.length) {
    stockStatusText.textContent = "Upload a Stock file before exporting.";
    return;
  }
  stockStatusText.textContent = "Generating Stock export...";
  exportStockButton.disabled = true;
  try {
    const response = await fetch("/api/stock/export", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      credentials: "include",
      body: JSON.stringify({ items }),
    });

    if (response.status === 401) {
      showAuthScreen();
      return;
    }

    if (!response.ok) {
      const payload = await response.json();
      stockStatusText.textContent = payload.error || "Stock export failed.";
      return;
    }

    const blob = await response.blob();
    const disposition = response.headers.get("Content-Disposition") || "";
    const match = disposition.match(/filename="(.+)"/);
    const fileName = match ? match[1] : "Stock.xlsx";
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.click();
    URL.revokeObjectURL(url);
    stockStatusText.textContent = `${fileName} downloaded.`;
  } catch (error) {
    console.error("Stock export error:", error);
    stockStatusText.textContent = "Stock export failed.";
  } finally {
    exportStockButton.disabled = items.length === 0;
  }
}

function switchPage(page) {
  state.currentPage = page;
  const showDashboard = page === "dashboard";
  const showClients = page === "clients";
  const showSettings = page === "settings";
  const showOrders = page === "orders";
  const showLeverschema = page === "leverschema";
  const showLaadschema = page === "laadschema";
  const showStock = page === "stock";
  const showHero = showDashboard || showClients;
  if (ordersClientTitle) {
    ordersClientTitle.textContent = state.selectedClient || "Orders";
  }
  appHero?.classList.toggle("hidden", !showHero);
  dashboardPage.classList.toggle("hidden", !showDashboard);
  clientsPage.classList.toggle("hidden", !showClients);
  settingsPage.classList.toggle("hidden", !showSettings);
  stockPage?.classList.toggle("hidden", !showStock);
  ordersPage.classList.toggle("hidden", !showOrders);
  leverschemaPage.classList.toggle("hidden", !showLeverschema);
  laadschemaPage.classList.toggle("hidden", !showLaadschema);
  if (showClients) {
    renderClientTabs();
  }
  if (showSettings) {
    renderSettingsPage();
  }
  if (showOrders) {
    updateOrdersClientMode();
  }
  if (showLeverschema) {
    state.currentLeverschemaSheet = getActiveSessionSheet();
    renderLeverschemaWorkbook();
  }
  if (showStock) {
    renderStockTable();
  }
}

function updateLeverschemaPageVisibility() {
  const hasData = getCurrentSessionLeverschemaEntries().length > 0;
  clearLeverschemaButton.disabled = !hasData;
  exportLeverschemaButton.disabled = !hasData;
  renderClientsLeverschemaCard();
  if (!hasData && state.currentPage === "leverschema") {
    switchPage("clients");
  }
}

function renderLeverschemaWorkbook() {
  updateLeverschemaPageVisibility();

  const sheet = state.currentLeverschemaSheet;
  const definitions = MASTER_SHEETS[sheet] || [];
  leverschemaTableTitle.textContent = sheet;
  leverschemaTableNote.textContent = "Saved Leverschema results are shown here in the same structure as the workbook sheet.";
  sheetTabButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.sheetTab === sheet);
    button.classList.toggle("hidden", button.dataset.sheetTab !== getActiveSessionSheet());
  });

  leverschemaMasterHead.innerHTML = `
    <tr>
      <th>Klant</th>
      <th>Palletplaatsen (PP)</th>
      <th>Losdatum</th>
      <th>Levertijd</th>
      <th>Vervoerder</th>
      <th>Opmerking</th>
    </tr>
  `;

  const entries = Object.entries(state.leverschemaResults)
    .map(([storageKey, entry]) => ({ storageKey, ...entry }))
    .filter((entry) => getLeverschemaEntrySessionId(entry.storageKey, entry) === getCurrentLeverschemaSessionId())
    .filter((entry) => entry.sheet === sheet)
    .map((entry) => ({
      ...entry,
      masterKey: entry.masterKey || resolveMasterCustomerKey(entry.customer, entry.ruleName),
    }))
    .sort((left, right) => String(left.savedAt).localeCompare(String(right.savedAt)));

  const queueByKey = {};
  for (const entry of entries) {
    if (!entry.masterKey) continue;
    if (!queueByKey[entry.masterKey]) queueByKey[entry.masterKey] = [];
    queueByKey[entry.masterKey].push(entry);
  }

  const rows = definitions.map((definition) => {
    const queue = queueByKey[definition.key] || [];
    const targetDate = formatWorkbookDate(definition.dateOffset, getCurrentLeverschemaBaseDate());
    const matchingIndex = queue.findIndex((entry) => normalizeWorkbookDate(entry.deliveryDate) === normalizeWorkbookDate(targetDate));
    const entry = matchingIndex >= 0 ? queue.splice(matchingIndex, 1)[0] : queue.length ? queue.shift() : null;
    return {
      klant: definition.klant,
      palletPlaces: entry ? formatNumber(entry.totalPalletPlaces) : "",
      storageKey: entry ? entry.storageKey : "",
      losdatum: targetDate,
      levertijd: definition.time,
      vervoerder: definition.carrier,
      opmerking: definition.note,
    };
  });

  leverschemaMasterBody.innerHTML = rows
    .map(
      (row) => `
        <tr>
          <td>${escapeHtml(row.klant)}</td>
          <td>${renderPpCell(row.palletPlaces, row.storageKey)}</td>
          <td>${escapeHtml(row.losdatum)}</td>
          <td>${escapeHtml(row.levertijd)}</td>
          <td>${escapeHtml(row.vervoerder)}</td>
          <td>${escapeHtml(row.opmerking)}</td>
        </tr>
      `,
    )
    .join("");
}

function clearLeverschemaMemory() {
  const sessionId = getCurrentLeverschemaSessionId();
  state.leverschemaResults = Object.fromEntries(
    Object.entries(state.leverschemaResults).filter(([storageKey, entry]) => getLeverschemaEntrySessionId(storageKey, entry) !== sessionId),
  );
  persistLeverschemaResults();
  renderLeverschemaSummary();
  renderLeverschemaIncludeOptions();
  renderLeverschemaWorkbook();
  renderClientsLeverschemaCard();
  statusText.textContent = "Saved pallet memory cleared for the active session.";
}

function handleLeverschemaRowClear(event) {
  const button = event.target.closest("[data-clear-leverschema-key]");
  if (!button) return;
  const storageKey = button.dataset.clearLeverschemaKey;
  if (!storageKey || !state.leverschemaResults[storageKey]) return;
  delete state.leverschemaResults[storageKey];
  persistLeverschemaResults();
  renderLeverschemaSummary();
  renderLeverschemaIncludeOptions();
  renderLeverschemaWorkbook();
  renderClientsLeverschemaCard();
  statusText.textContent = "Saved PP entry removed.";
}

async function exportLeverschemaWorkbook() {
  const sessionResults = getCurrentSessionLeverschemaResults();
  const hasData = Object.keys(sessionResults).length > 0;
  if (!hasData) return;

  statusText.textContent = "Generating Leverschema workbook...";
  try {
    const response = await fetch("/api/export_leverschema", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        results: sessionResults,
        selectedSheet: state.currentLeverschemaSheet,
        sessionDate: getCurrentLeverschemaBaseDate(),
      }),
    });

    if (response.status === 401) {
      resetPreviewState();
      showAuthScreen();
      return;
    }

    if (!response.ok) {
      const payload = await response.json();
      statusText.textContent = payload.error || "Leverschema export failed.";
      return;
    }

    const blob = await response.blob();
    const disposition = response.headers.get("Content-Disposition") || "";
    const match = disposition.match(/filename="(.+)"/);
    const fileName = match ? match[1] : "Leverschema_OOH.xlsm";
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.click();
    URL.revokeObjectURL(url);
    statusText.textContent = `${fileName} downloaded.`;
  } catch {
    statusText.textContent = "Leverschema export failed.";
  }
}

function calculateLeverschema(order) {
  const config = resolveLeverschemaConfig(order);
  if (!config) {
    return {
      ruleName: "Not available",
      totalPallets: 0,
      totalPalletPlaces: 0,
      rows: [],
    };
  }

  const rows = order.items.map((item) => config.computeRow(item));
  return {
    ruleName: config.name,
    totalPallets: config.computeTotalPallets(rows),
    totalPalletPlaces: config.computeTotalPalletPlaces(rows),
    rows,
  };
}

function resolveLeverschemaConfig(order) {
  if (state.mode === "special") {
    return {
      name: "Havi Duisburg",
      computeRow: (item) => {
        const fullPalletQty = normalizeUnit(item.unit) === "slices" ? 800 : 40;
        const palletPlaces = fullPalletQty ? Math.ceil(item.quantity / fullPalletQty) : 0;
        return { ...item, fullPalletQty, palletPlaces };
      },
      computeTotalPallets: (rows) => sumBy(rows, "palletPlaces"),
      computeTotalPalletPlaces: (rows) => sumBy(rows, "palletPlaces"),
    };
  }

  const deliveryPoint = normalizeText(order.deliveryPoint || getDeliveryPointKey(order));
  const haystack = normalizeText(`${order.customer} ${order.fatrans} ${order.deliveryPoint || ""} ${state.selectedClient}`);
  const isCarrefour = haystack.includes("carrefour") || ["fif", "kdc"].includes(deliveryPoint);
  if (isCarrefour && (deliveryPoint === "kdc" || haystack.includes("carrefour kdc"))) {
    return {
      name: "Carrefour KDC",
      computeRow: (item) => {
        const boxType = CARREFOUR_BOX_TYPES[Number(item.primary)] || "";
        const palletMetric = calculateCarrefourPalletMetric(item.quantity, boxType);
        return { ...item, boxType, palletMetric };
      },
      computeTotalPallets: (rows) => Math.ceil(sumBy(rows, "palletMetric") / 22),
      computeTotalPalletPlaces: (rows) => Math.ceil(sumBy(rows, "palletMetric") / 22),
    };
  }
  if (isCarrefour && (deliveryPoint === "fif" || haystack.includes("carrefour fif"))) {
    return {
      name: "Carrefour FIF",
      computeRow: (item) => {
        const boxType = CARREFOUR_BOX_TYPES[Number(item.primary)] || "";
        const palletMetric = calculateCarrefourPalletMetric(item.quantity, boxType);
        return { ...item, boxType, palletMetric };
      },
      computeTotalPallets: (rows) => Math.ceil(sumBy(rows, "palletMetric") / 22),
      computeTotalPalletPlaces: (rows) => Math.ceil(sumBy(rows, "palletMetric") / 22),
    };
  }
  if (haystack.includes("colruyt")) {
    return {
      name: "Colruyt",
      computeRow: (item) => ({ ...item, palletPlaces: Math.ceil(item.quantity / 132) }),
      computeTotalPallets: (rows) => sumBy(rows, "palletPlaces"),
      computeTotalPalletPlaces: (rows) => sumBy(rows, "palletPlaces"),
    };
  }
  if (haystack.includes("heeren")) {
    return {
      name: "Heeren",
      computeRow: (item) => ({ ...item, palletPlaces: Math.ceil(item.quantity / 90) }),
      computeTotalPallets: (rows) => sumBy(rows, "palletPlaces"),
      computeTotalPalletPlaces: (rows) => Math.ceil(rows.reduce((sum, row) => sum + row.quantity / 90, 0)),
    };
  }
  if (haystack.includes("havi belg")) {
    return createHaviConfig("Havi Belgium", HAVI_BELGIUM_FULL_PALLET);
  }
  if (haystack.includes("havi nether") || haystack.includes("havi nederland")) {
    return createHaviConfig("Havi Netherlands", HAVI_NETHERLANDS_FULL_PALLET);
  }
  if (
    haystack.includes("havi duisburg") ||
    haystack.includes("havi de uien") ||
    (haystack.includes("havi logistics gmbh") && deliveryPoint !== "havi nl" && deliveryPoint !== "nl")
  ) {
    return createHaviDuisburgConfig();
  }

  return null;
}

function resolveMasterCustomerKey(customer, ruleName, selectedClient = "") {
  const haystack = normalizeText(`${customer} ${ruleName} ${selectedClient}`);
  console.log('🔍 Resolving master key:', { customer, ruleName, selectedClient, haystack });
  
  if (haystack.includes("carrefour kdc")) return "carrefour_kdc";
  if (haystack.includes(" kdc") || haystack.endsWith("kdc")) return "carrefour_kdc";
  if (haystack.includes("carrefour fif")) return "carrefour_fif";
  if (haystack.includes(" fif") || haystack.endsWith("fif")) return "carrefour_fif";
  if (haystack.includes("colruyt saturday")) return "colruyt_saturday";
  if (haystack.includes("colruyt")) return "colruyt";
  if (haystack.includes("heeren")) return "heeren";
  
  // Check for Saturday variants BEFORE checking the regular ones
  if (haystack.includes("havi duisburg saturday") || (haystack.includes("havi de saturday") && haystack.includes("havi logistics gmbh"))) {
    console.log('✅ Matched: havi_duisburg_saturday');
    return "havi_duisburg_saturday";
  }
  
  if (haystack.includes("havi duisburg") || haystack.includes("havi logistics gmbh")) {
    console.log('✅ Matched: havi_duisburg');
    return "havi_duisburg";
  }
  
  console.log('⚠️ No match found, returning empty string');
  return "";
}

function createHaviConfig(name, mapping) {
  return {
    name,
    computeRow: (item) => {
      const fullPalletQty = resolveHaviFullPalletQty(mapping, item.primary, item.unit);
      const palletPlaces = fullPalletQty ? Math.ceil(item.quantity / fullPalletQty) : 0;
      return { ...item, fullPalletQty, palletPlaces };
    },
    computeTotalPallets: (rows) => sumBy(rows, "palletPlaces"),
    computeTotalPalletPlaces: (rows) => sumBy(rows, "palletPlaces"),
  };
}

function createHaviDuisburgConfig() {
  return {
    name: "Havi Duisburg",
    computeRow: (item) => {
      const fullPalletQty = normalizeUnit(item.unit) === "cases" ? 40 : normalizeUnit(item.unit) === "slices" ? 800 : 0;
      const palletPlaces = fullPalletQty ? Math.ceil(item.quantity / fullPalletQty) : 0;
      return { ...item, fullPalletQty, palletPlaces };
    },
    computeTotalPallets: (rows) => sumBy(rows, "palletPlaces"),
    computeTotalPalletPlaces: (rows) => sumBy(rows, "palletPlaces"),
  };
}

function resolveHaviFullPalletQty(mapping, article, unit) {
  const config = mapping[Number(article)];
  if (!config) return 0;
  const normalizedUnit = normalizeUnit(unit);
  if (normalizedUnit === "collo") return config.collo || 0;
  if (normalizedUnit === "stuk" || normalizedUnit === "stuck") return config.stuk || 0;
  return 0;
}

function calculateCarrefourPalletMetric(quantity, boxType) {
  if (boxType === "S") return Math.ceil(quantity / 8) + Math.ceil(quantity / 64) * 2;
  if (boxType === "M") return Math.ceil(quantity / 8) * 2 + Math.ceil(quantity / 40) * 2;
  if (boxType === "L") return Math.ceil(quantity / 4) * 2 + Math.ceil(quantity / 20) * 2;
  return 0;
}

function buildLeverschemaRoutingContext(order) {
  return [state.selectedClient, order.fatrans, order.deliveryPoint].filter(Boolean).join(" ");
}

function buildLeverschemaOrderKey(order, sheet) {
  const masterKey = resolveMasterCustomerKey(order.customer, order.customer, buildLeverschemaRoutingContext(order)) || normalizeText(order.customer || "leverschema");
  const deliveryKey = normalizeWorkbookDate(order.deliveryDate).replaceAll("/", "-") || "no-date";
  const sid = getCurrentLeverschemaSessionId();
  return `${sid}::${sheet}::${masterKey}::${deliveryKey}`;
}

function loadLeverschemaResults() {
  try {
    const raw = window.localStorage.getItem(LEVERSCHEMA_STORAGE_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    for (const [key, entry] of Object.entries(parsed)) {
      if (entry && !entry.masterKey) {
        entry.masterKey = resolveMasterCustomerKey(entry.customer, entry.ruleName);
        parsed[key] = entry;
      }
      if (entry && !entry.sessionId) {
        entry.sessionId = getLeverschemaEntrySessionId(key, entry);
        parsed[key] = entry;
      }
    }
    return parsed;
  } catch {
    return {};
  }
}

function getCurrentLeverschemaSessionId() {
  return state.activeWorkSession?.id || "default";
}

function getCurrentLeverschemaBaseDate() {
  return state.activeWorkSession?.date || formatDateForInput(new Date());
}

function getLeverschemaEntrySessionId(storageKey, entry = {}) {
  if (entry?.sessionId) return entry.sessionId;
  const [sessionId = "default"] = String(storageKey || "").split("::");
  return sessionId || "default";
}

function getCurrentSessionLeverschemaEntries() {
  const sessionId = getCurrentLeverschemaSessionId();
  return Object.entries(state.leverschemaResults)
    .map(([storageKey, entry]) => ({ storageKey, ...entry }))
    .filter((entry) => getLeverschemaEntrySessionId(entry.storageKey, entry) === sessionId);
}

function getCurrentSessionLeverschemaResults() {
  const sessionId = getCurrentLeverschemaSessionId();
  return Object.fromEntries(
    Object.entries(state.leverschemaResults).filter(([storageKey, entry]) => getLeverschemaEntrySessionId(storageKey, entry) === sessionId),
  );
}

function loadClientActionSettings() {
  try {
    const raw = window.localStorage.getItem(CLIENT_ACTION_SETTINGS_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function persistClientActionSettings() {
  try {
    window.localStorage.setItem(CLIENT_ACTION_SETTINGS_KEY, JSON.stringify(state.clientActionSettings));
  } catch {
    // Ignore storage issues and keep in-memory state.
  }
}

function loadHaviUienSettings() {
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

function persistHaviUienSettings() {
  try {
    window.localStorage.setItem(HAVI_UIEN_SETTINGS_KEY, JSON.stringify(getHaviUienSettings()));
  } catch {
    // Ignore storage issues and keep in-memory state.
  }
}

function getHaviUienSettings() {
  return {
    article: String(state.haviUienSettings?.article || "").trim(),
    description: String(state.haviUienSettings?.description || "").trim(),
  };
}

function getClientActionPreferences(client) {
  const stored = state.clientActionSettings[client] || {};
  const defaults = {
    export: true,
    merge: true,
    special: true,
    nettoMd: true,
    leverschema: true,
    printCmr: false,
  };

  if (normalizeText(client) === "nettomd") {
    defaults.export = false;
    defaults.merge = false;
    defaults.special = false;
    defaults.nettoMd = true;
    defaults.leverschema = false;
  }

  const cmrDefaultClients = ["carrefour", "carrefour fif", "carrefour kdc", "colruyt", "colruyt saturday", "denemark", "edeka", "edeka laatzen", "edeka mochmuhl", "globus", "havi", "havi nl", "havi be", "havi de", "havi de saturday", "havi duisburg saturday", "havi dc duisburg dc wunstorf dc neu wulmstorf", "heeren", "nettomd", "rewe", "penny", "hanos"];
  const normalizedClient = normalizeText(client);
  const shouldHaveCMR = cmrDefaultClients.includes(normalizedClient);
  
  // Force printCmr to true for CMR clients BEFORE merging with stored settings
  if (shouldHaveCMR) {
    defaults.printCmr = true;
    // Also remove printCmr from stored settings if it's false
    if (stored.printCmr === false) {
      delete stored.printCmr;
    }
  }

  const result = { ...defaults, ...stored };
  
  // Double-check: Force printCmr to true for clients that should have CMR
  if (shouldHaveCMR) {
    result.printCmr = true;
  }
  
  return result;
}


function updateHeaderAccountInfo() {
  if (sessionUserName) {
    sessionUserName.textContent = state.user?.name || "-";
  }
  if (sessionUserEmail) {
    sessionUserEmail.textContent = state.user?.email || "-";
  }
  if (headerSessionName) {
    headerSessionName.textContent = state.activeWorkSession?.name || "No active session";
  }
  if (headerSessionDate) {
    headerSessionDate.textContent = state.activeWorkSession?.date
      ? `Session date: ${formatSessionDate(state.activeWorkSession.date)}`
      : "Create or open a session.";
  }
  if (state.activeWorkSession?.id) {
    window.localStorage.setItem(ACTIVE_WORK_SESSION_STORAGE_KEY, state.activeWorkSession.id);
  } else {
    window.localStorage.removeItem(ACTIVE_WORK_SESSION_STORAGE_KEY);
  }
}

function formatDateForInput(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function parseInputDate(dateValue) {
  const [year, month, day] = String(dateValue || "").split("-").map(Number);
  if (!year || !month || !day) return null;
  const date = new Date(year, month - 1, day);
  return Number.isNaN(date.getTime()) ? null : date;
}

function buildDefaultSessionName(dateValue) {
  return `Plan for ${formatSessionDate(dateValue)}`;
}

function formatSessionDate(dateValue) {
  const [year, month, day] = String(dateValue || "").split("-");
  if (!year || !month || !day) return "-";
  return `${day}/${month}/${year}`;
}

function persistLeverschemaResults() {
  persistLeverschemaLocalCache();
  schedulePersistTeamState();
}

function parseNumber(value) {
  const normalized = String(value ?? "").replace(",", ".").trim();
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function formatNumber(value) {
  const numeric = Number(value) || 0;
  return Number.isInteger(numeric) ? String(numeric) : numeric.toFixed(2);
}

function formatSavedAt(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "just now";
  return date.toLocaleString();
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
  const selectedDate = parseInputDate(input.value) || new Date();
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

function formatWorkbookDate(offset, baseDateValue = getCurrentLeverschemaBaseDate()) {
  const date = baseDateValue ? new Date(`${baseDateValue}T00:00:00`) : new Date();
  if (Number.isNaN(date.getTime())) {
    date.setTime(Date.now());
  }
  date.setHours(0, 0, 0, 0);
  date.setDate(date.getDate() + Number(offset || 0));
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
}

function normalizeWorkbookDate(value) {
  const text = String(value || "").trim().replaceAll(".", "/").replaceAll("-", "/");
  const parts = text.split("/").filter(Boolean);
  if (parts.length !== 3) return "";
  const [day, month, year] = parts;
  return `${String(day).padStart(2, "0")}/${String(month).padStart(2, "0")}/${year}`;
}

function sumBy(items, key) {
  return items.reduce((sum, item) => sum + (Number(item[key]) || 0), 0);
}

function normalizeText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\n/g, " ")  // Replace newlines with spaces
    .toLowerCase();
}

function normalizeUnit(value) {
  return normalizeText(value).replace(/\s+/g, "");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

// Laadschema functionality
function initializeLaadschema() {
  // Load saved data from localStorage
  state.laadschemaData = loadLaadschemaData();
  state.laadschemaCustomTrucks = loadLaadschemaCustomTrucks();
  
  // Set current date and week
  const today = new Date();
  const currentWeek = getWeekNumber(today);
  
  laadschemaDateInput.value = formatDateForInput(today);
  laadschemaWeekInput.value = currentWeek;
  
  // Set current day
  laadschemaDaySelect.value = getLaadschemaDayName(today);
  
  renderLaadschemaTable();
}

function getWeekNumber(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

function getLaadschemaDayName(date) {
  return LAADSCHEMA_DAYS[date.getDay()] || "Maandag";
}

function getIsoWeekStart(date) {
  const start = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const day = start.getDay() || 7;
  start.setDate(start.getDate() - day + 1);
  start.setHours(0, 0, 0, 0);
  return start;
}

function getDateForIsoWeekDay(year, week, dayName) {
  const weekNumber = Number(week);
  const offset = LAADSCHEMA_ISO_DAY_OFFSETS[dayName] ?? 0;
  const fourthOfJanuary = new Date(year, 0, 4);
  const firstWeekStart = getIsoWeekStart(fourthOfJanuary);
  const date = new Date(firstWeekStart);
  date.setDate(firstWeekStart.getDate() + (weekNumber - 1) * 7 + offset);
  return date;
}

function getSelectedLaadschemaDate() {
  return parseInputDate(laadschemaDateInput.value) || new Date();
}

function getLaadschemaStorageKey() {
  const selectedDate = getSelectedLaadschemaDate();
  const key = formatDateForInput(selectedDate);
  if (laadschemaDateInput.value !== key) {
    laadschemaDateInput.value = key;
  }
  return key;
}

function getLaadschemaDayKey() {
  return laadschemaDaySelect.value || getLaadschemaDayName(getSelectedLaadschemaDate());
}

function getCurrentLaadschemaData() {
  const dateKey = getLaadschemaStorageKey();
  const dayKey = getLaadschemaDayKey();
  return state.laadschemaData?.[dateKey] || state.laadschemaData?.[dayKey] || {};
}

function ensureCurrentLaadschemaDataStore() {
  const dateKey = getLaadschemaStorageKey();
  const dayKey = getLaadschemaDayKey();
  if (!state.laadschemaData[dateKey]) {
    state.laadschemaData[dateKey] = state.laadschemaData[dayKey]
      ? JSON.parse(JSON.stringify(state.laadschemaData[dayKey]))
      : {};
  }
  return dateKey;
}

function getCurrentLaadschemaCustomTrucks() {
  const dateKey = getLaadschemaStorageKey();
  const dayKey = getLaadschemaDayKey();
  return state.laadschemaCustomTrucks?.[dateKey] || state.laadschemaCustomTrucks?.[dayKey] || [];
}

function ensureCurrentLaadschemaCustomTruckStore() {
  const dateKey = getLaadschemaStorageKey();
  const dayKey = getLaadschemaDayKey();
  if (!state.laadschemaCustomTrucks[dateKey]) {
    state.laadschemaCustomTrucks[dateKey] = Array.isArray(state.laadschemaCustomTrucks[dayKey])
      ? JSON.parse(JSON.stringify(state.laadschemaCustomTrucks[dayKey]))
      : [];
  }
  return dateKey;
}

function updateLaadschemaDate() {
  const selectedDate = getSelectedLaadschemaDate();
  const weekStart = getIsoWeekStart(selectedDate);
  const dayOffset = LAADSCHEMA_ISO_DAY_OFFSETS[laadschemaDaySelect.value] ?? 0;
  const matchingDate = new Date(weekStart);
  matchingDate.setDate(weekStart.getDate() + dayOffset);
  laadschemaDateInput.value = formatDateForInput(matchingDate);
  laadschemaWeekInput.value = getWeekNumber(matchingDate);
  renderLaadschemaTable();
}

function updateLaadschemaWeek() {
  // When date changes, calculate and update week
  const date = parseInputDate(laadschemaDateInput.value);
  if (date) {
    const week = getWeekNumber(date);
    laadschemaWeekInput.value = week;
    laadschemaDaySelect.value = getLaadschemaDayName(date);
  }
  renderLaadschemaTable();
}

function updateLaadschemaFromWeek() {
  // When week changes, keep the selected weekday and move the calendar date to the matching day in that week.
  const selectedDate = getSelectedLaadschemaDate();
  const year = selectedDate.getFullYear();
  const week = parseInt(laadschemaWeekInput.value);
  
  if (week >= 1 && week <= 53) {
    const matchingDate = getDateForIsoWeekDay(year, week, laadschemaDaySelect.value);
    laadschemaDateInput.value = formatDateForInput(matchingDate);
    laadschemaDaySelect.value = getLaadschemaDayName(matchingDate);
  }
  renderLaadschemaTable();
}

function renderLaadschemaTable() {
  // Get current selected day
  const currentDay = laadschemaDaySelect?.value || 'Maandag';
  
  // Define different configurations for Sunday vs other days
  let clientsOriginal, schedulesOriginal, plannedTimesOriginal;
  
  if (currentDay === 'Zondag') {
    // Sunday configuration (from Excel file)
    clientsOriginal = [
      'McD NL', 'McD UK', 'DK auto 1', 'DK auto 2', 'DK auto 3', 'DK auto 4',
      'Rewe / Penny', 'Rewe / Penny', 'Rewe / Penny',
      'Hanos for Monday', 'McD Duisburg for Monday'
    ];
    
    schedulesOriginal = [
      'Mon-Sat', 'Mon-Sat', 'Mon-Sat', 'Mon-Sat', 'Mon-Sat', 'Mon-Sat',
      'Mon-Sat', 'Mon-Sat', 'Mon-Sat',
      'Mon-Sat', 'Mon-Sat'
    ];
    
    plannedTimesOriginal = [
      '13:00', '16:00', '16:30', '16:30', '17:00', '17:00',
      '', '', '',
      '09:00', '10:00'
    ];
  } else {
    // Regular days configuration (starting from column E, skipping column D)
    clientsOriginal = [
      'Hanos', 'McD Duisburg', 'MCD BE', 'Carrefour FIF', 
      'McD NL', 'McD DE', 'McD UK', 'DK auto 1', 'DK auto 2', 
      'Rewe / Penny', 'Rewe / Penny', 'Rewe / Penny', 'Rewe / Penny', 
      'Netto MD', 'Carrefour KDC', 'Colruyt', 'Hello Fresh'
    ];
    
    schedulesOriginal = [
      'Mon-Sat', 'Mon-Sat', 'Mon-Sat', 'Mon-Sat',
      'Mon-Sat', 'Mon-Sat', 'Mon-Sat', 'Mon-Sat', 'Mon-Sat',
      'Mon-Sat', 'Mon-Sat', 'Mon-Sat', 'Mon-Sat', '',
      'Mon-Sat', 'Mon-Sat', 'Mon Tue Thu Fri Sat'
    ];
    
    plannedTimesOriginal = [
      '09:00', '10:00', '11:00', '12:30', '13:00', '15:30', '16:00',
      '16:30', '17:30', '17:00', '17:30', '18:00', '19:00', '22:00',
      '04:00', '01:30', '03:00'
    ];
  }
  
  // Add custom trucks for this day
  const currentCustomTrucks = getCurrentLaadschemaCustomTrucks();
  if (Array.isArray(currentCustomTrucks)) {
    currentCustomTrucks.forEach(truck => {
      clientsOriginal.push(truck.clientName);
      schedulesOriginal.push('Extra');
      plannedTimesOriginal.push(truck.plannedTime);
    });
  }
  
  // Create array of indices with their times
  const clientIndices = clientsOriginal.map((client, idx) => ({
    idx,
    client,
    time: plannedTimesOriginal[idx]
  }));
  
  // Separate clients that depart after midnight
  const afterMidnightClients = ['Carrefour KDC', 'Colruyt', 'Hello Fresh'];
  const beforeMidnight = clientIndices.filter(c => !afterMidnightClients.includes(c.client));
  const afterMidnight = clientIndices.filter(c => afterMidnightClients.includes(c.client));
  
  // Sort before midnight by time, keeping original order for same times
  beforeMidnight.sort((a, b) => {
    const timeA = new Date(`2000-01-01 ${a.time}`);
    const timeB = new Date(`2000-01-01 ${b.time}`);
    const timeDiff = timeA - timeB;
    // If times are equal, maintain original order by index
    return timeDiff !== 0 ? timeDiff : a.idx - b.idx;
  });
  
  // Sort after midnight by time (chronological order), keeping original order for same times
  afterMidnight.sort((a, b) => {
    const timeA = new Date(`2000-01-01 ${a.time}`);
    const timeB = new Date(`2000-01-01 ${b.time}`);
    const timeDiff = timeA - timeB;
    // If times are equal, maintain original order by index
    return timeDiff !== 0 ? timeDiff : a.idx - b.idx;
  });
  
  // Combine: before midnight sorted + after midnight
  const sortedIndices = [...beforeMidnight, ...afterMidnight];
  
  // Create sorted arrays
  const clients = sortedIndices.map(c => c.client);
  const schedules = sortedIndices.map(c => schedulesOriginal[c.idx]);
  const plannedTimes = sortedIndices.map(c => c.time);
  
  // Create header
  laadschemaTableHead.innerHTML = `
    <tr>
      <th class="row-header">Freight</th>
      ${clients.map(client => `<th>${escapeHtml(client)}</th>`).join('')}
    </tr>
  `;
  
  // Create rows
  const rows = [
    {
      key: 'remark',
      label: 'Remark',
      info: '',
      values: schedules
    },
    {
      key: 'delivery-note-checked',
      label: 'Delivery note checked',
      info: 'Name',
      values: clients.map(() => '')
    },
    {
      key: 'vehicle-inspection',
      label: 'Physical check + vehicle inspection',
      info: 'Name',
      values: clients.map(() => '')
    },
    {
      key: 'measured-temperature',
      label: 'Measured temperature before loading',
      info: '°C',
      inputType: 'number',
      values: clients.map(() => '')
    },
    {
      key: 'set-temperature',
      label: 'Set temperature before loading',
      info: '°C',
      inputType: 'number',
      values: clients.map(() => '3')
    },
    {
      key: 'logger-activated',
      label: 'Logger activated',
      info: 'Time',
      values: clients.map(() => 'x')
    },
    {
      key: 'wheel-chocks',
      label: 'Wheel chocks placed',
      info: 'Check',
      values: clients.map(() => '')
    },
    {
      key: 'planned-departure',
      label: 'Planned departure time',
      info: '',
      inputType: 'time',
      values: plannedTimes
    },
    {
      key: 'actual-departure',
      label: 'Actual departure time',
      info: '',
      inputType: 'time',
      values: clients.map(() => '')
    },
    {
      key: 'driver-arrival',
      label: 'Driver arrival time',
      info: '',
      inputType: 'time',
      values: clients.map(() => '')
    },
    {
      key: 'delay-reason',
      label: 'Delay reason',
      info: '',
      values: clients.map(() => '')
    },
    {
      key: 'extra-products',
      label: 'Extra products loaded?',
      info: 'Extra products loaded',
      values: clients.map(() => '')
    },
    {
      key: 'total-pallet-places',
      label: 'Total pallet places',
      info: 'PP',
      values: clients.map(() => '0')
    }
  ];
  
  // Get current day for loading saved data (already declared at function start)
  const savedData = getCurrentLaadschemaData();
  
  laadschemaTableBody.innerHTML = rows.map((row, rowIndex) => {
    // Combine label with info if info exists
    const displayLabel = row.info ? `${row.label} (${row.info})` : row.label;
    
    // Check if this is the "Actual leaving time" row
    const isActualLeavingTimeRow = row.key === 'actual-departure';
    // Check if this is the "Arrival time" row
    const isArrivalTimeRow = row.key === 'driver-arrival';
    
    return `
    <tr>
      <td class="row-header">${escapeHtml(displayLabel)}</td>
      ${row.values.map((value, colIndex) => {
        // Load saved value if it exists
        const savedValue = savedData[rowIndex] && savedData[rowIndex][colIndex] !== undefined 
          ? savedData[rowIndex][colIndex] 
          : value;
        
        let cellStyle = '';
        
        // If this is the actual leaving time row, compare with planned time
        if (isActualLeavingTimeRow) {
          const plannedTime = rows[7].values[colIndex]; // Row 7 is "Planned departure time"
          const actualTime = savedValue;
          
          // Compare times if both exist
          if (plannedTime && actualTime) {
            const planned = new Date(`2000-01-01 ${plannedTime}`);
            const actual = new Date(`2000-01-01 ${actualTime}`);
            
            if (actual > planned) {
              cellStyle = 'background-color: #ff6b6b;'; // Red for late
            } else {
              cellStyle = 'background-color: #51cf66;'; // Green for on time or early
            }
          }
        }
        
        return `
        <td class="editable-cell laadschema-time-cell" style="${cellStyle}">
          <div class="laadschema-time-input-wrapper">
            <input type="text" value="${escapeHtml(savedValue)}" 
                   onchange="updateLaadschemaCell(${rowIndex}, ${colIndex}, this.value)"
                   ${row.inputType === 'number' ? 'type="number"' : ''}
                   ${row.inputType === 'time' ? 'type="time"' : ''}>
            ${isActualLeavingTimeRow && (!savedValue || savedValue.trim() === '') ? `<button type="button" class="laadschema-in-button" onclick="setCurrentTime(${rowIndex}, ${colIndex}, 'in')">Out</button>` : ''}
            ${isActualLeavingTimeRow && savedValue && savedValue.trim() !== '' ? `<button type="button" class="laadschema-in-button" onclick="setCurrentTime(${rowIndex}, ${colIndex}, 'in')" style="display: none;">Out</button>` : ''}
            ${isArrivalTimeRow && (!savedValue || savedValue.trim() === '') ? `<button type="button" class="laadschema-out-button" onclick="setCurrentTime(${rowIndex}, ${colIndex}, 'out')">In</button>` : ''}
            ${isArrivalTimeRow && savedValue && savedValue.trim() !== '' ? `<button type="button" class="laadschema-out-button" onclick="setCurrentTime(${rowIndex}, ${colIndex}, 'out')" style="display: none;">In</button>` : ''}
          </div>
        </td>
      `;
      }).join('')}
    </tr>
  `;
  }).join('');
}

function updateLaadschemaCell(rowIndex, colIndex, value) {
  // Store data by exact calendar date so each weekday/date combination has its own plan.
  const currentKey = ensureCurrentLaadschemaDataStore();
  
  // Initialize storage structure if needed
  if (!state.laadschemaData[currentKey]) {
    state.laadschemaData[currentKey] = {};
  }
  if (!state.laadschemaData[currentKey][rowIndex]) {
    state.laadschemaData[currentKey][rowIndex] = {};
  }
  
  // Store the updated value
  state.laadschemaData[currentKey][rowIndex][colIndex] = value;
  console.log(`Updated cell [${rowIndex}][${colIndex}] to: ${value}`);
  
  // Persist to localStorage
  persistLaadschemaData();
  
  // Toggle button visibility based on value
  const inButton = document.querySelector(`#laadschemaTableBody tr:nth-child(${rowIndex + 1}) td:nth-child(${colIndex + 2}) .laadschema-in-button`);
  const outButton = document.querySelector(`#laadschemaTableBody tr:nth-child(${rowIndex + 1}) td:nth-child(${colIndex + 2}) .laadschema-out-button`);
  
  if (inButton) {
    if (value && value.trim() !== '') {
      inButton.style.display = 'none';
    } else {
      inButton.style.display = 'block';
    }
  }
  
  if (outButton) {
    if (value && value.trim() !== '') {
      outButton.style.display = 'none';
    } else {
      outButton.style.display = 'block';
    }
  }
  
  // If this is the actual leaving time row (row 8), update the color
  if (rowIndex === 8) {
    const plannedTimeCell = document.querySelector(`#laadschemaTableBody tr:nth-child(${rowIndex + 1}) td:nth-child(${colIndex + 2})`);
    const actualTimeCell = document.querySelector(`#laadschemaTableBody tr:nth-child(${rowIndex + 1}) td:nth-child(${colIndex + 2})`);
    
    // Get the planned time from row 7
    const plannedTimeInput = document.querySelector(`#laadschemaTableBody tr:nth-child(8) td:nth-child(${colIndex + 2}) input`);
    const actualTimeInput = document.querySelector(`#laadschemaTableBody tr:nth-child(9) td:nth-child(${colIndex + 2}) input`);
    
    if (plannedTimeInput && actualTimeInput) {
      const plannedTime = plannedTimeInput.value;
      const actualTime = actualTimeInput.value;
      
      if (plannedTime && actualTime) {
        const planned = new Date(`2000-01-01 ${plannedTime}`);
        const actual = new Date(`2000-01-01 ${actualTime}`);
        
        const cell = actualTimeInput.closest('td');
        if (actual > planned) {
          cell.style.backgroundColor = '#ff6b6b'; // Red for late
        } else {
          cell.style.backgroundColor = '#51cf66'; // Green for on time or early
        }
      } else {
        // Clear background color if no time
        const cell = actualTimeInput.closest('td');
        cell.style.backgroundColor = '';
      }
    }
  }
}

function setCurrentTime(rowIndex, colIndex, buttonType) {
  // Get current time in HH:MM format
  const now = new Date();
  const hours = String(now.getHours()).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  const currentTime = `${hours}:${minutes}`;
  
  // Find the input field and set the value
  const input = document.querySelector(`#laadschemaTableBody tr:nth-child(${rowIndex + 1}) td:nth-child(${colIndex + 2}) input`);
  const button = document.querySelector(`#laadschemaTableBody tr:nth-child(${rowIndex + 1}) td:nth-child(${colIndex + 2}) .laadschema-${buttonType}-button`);
  
  if (input) {
    input.value = currentTime;
    // Hide the button
    if (button) {
      button.style.display = 'none';
    }
    // Trigger the onchange event to update colors
    updateLaadschemaCell(rowIndex, colIndex, currentTime);
  }
}

async function exportLaadschemaData() {
  const currentDay = laadschemaDaySelect.value;
  const currentDate = laadschemaDateInput.value;
  const currentWeek = laadschemaWeekInput.value;
  const dayData = getCurrentLaadschemaData();
  const customTrucks = getCurrentLaadschemaCustomTrucks();
  
  console.log("Export - Current day:", currentDay);
  console.log("Export - Laadschema data:", state.laadschemaData);
  console.log("Export - Day data:", dayData);
  console.log("Export - Custom trucks:", state.laadschemaCustomTrucks);
  
  // Check if there's any data to export
  if (!dayData || Object.keys(dayData).length === 0) {
    statusText.textContent = "No data to export for this date.";
    return;
  }
  
  statusText.textContent = "Generating Laadschema export...";
  
  try {
    const payload = {
      data: { [currentDay]: dayData },
      customTrucks: { [currentDay]: Array.isArray(customTrucks) ? customTrucks : [] },
      selectedDay: currentDay,
      selectedDate: currentDate,
      selectedWeek: currentWeek,
    };
    
    console.log("Export payload:", JSON.stringify(payload, null, 2));
    
    const response = await fetch("/api/export_laadschema", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      credentials: "include",
      body: JSON.stringify(payload),
    });

    if (response.status === 401) {
      resetPreviewState();
      showAuthScreen();
      return;
    }

    if (!response.ok) {
      const payload = await response.json();
      console.error("Export error:", payload);
      statusText.textContent = payload.error || "Laadschema export failed.";
      return;
    }

    const blob = await response.blob();
    const disposition = response.headers.get("Content-Disposition") || "";
    const match = disposition.match(/filename="(.+)"/);
    const fileName = match ? match[1] : "Laadschema.xlsx";
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.click();
    URL.revokeObjectURL(url);
    statusText.textContent = `${fileName} downloaded.`;
  } catch (err) {
    console.error("Laadschema export error:", err);
    statusText.textContent = "Laadschema export failed.";
  }
}

function getSelectedLaadschemaDayLabel() {
  const selectedOption = laadschemaDaySelect?.selectedOptions?.[0];
  return selectedOption?.textContent?.trim() || laadschemaDaySelect?.value || "selected day";
}

async function clearLaadschemaDay() {
  const currentDay = laadschemaDaySelect.value;
  const currentKey = getLaadschemaStorageKey();
  const dataKeyToClear = state.laadschemaData[currentKey] ? currentKey : currentDay;
  const truckKeyToClear = state.laadschemaCustomTrucks[currentKey] ? currentKey : currentDay;

  const shouldClear = await showConfirmationDialog({
    eyebrow: "Loading schedule",
    title: "Clear this day?",
    message: `Are you sure you want to clear all data for ${getSelectedLaadschemaDayLabel()} ${currentKey}? This action cannot be undone.`,
    confirmLabel: "Clear day",
    cancelLabel: "Cancel",
  });
  if (!shouldClear) {
    return;
  }
  
  // Clear data for the exact selected date
  if (state.laadschemaData[dataKeyToClear]) {
    delete state.laadschemaData[dataKeyToClear];
    persistLaadschemaData();
  }
  if (state.laadschemaCustomTrucks[truckKeyToClear]) {
    delete state.laadschemaCustomTrucks[truckKeyToClear];
    localStorage.setItem('laadschemaCustomTrucks', JSON.stringify(state.laadschemaCustomTrucks));
    schedulePersistTeamState();
  }
  
  // Re-render table to show empty cells
  renderLaadschemaTable();
  
  console.log(`Cleared Laadschema data for ${currentDay} ${currentKey}`);
}

async function clearAllLaadschemaData() {
  const shouldClear = await showConfirmationDialog({
    eyebrow: "Loading schedule",
    title: "Clear all days?",
    message: "Are you sure you want to clear all loading schedule data for every day? This action cannot be undone.",
    confirmLabel: "Clear all",
    cancelLabel: "Cancel",
  });
  if (!shouldClear) {
    return;
  }
  
  // Clear all days data
  state.laadschemaData = {};
  persistLaadschemaData();
  
  // Clear all custom trucks
  state.laadschemaCustomTrucks = {};
  localStorage.setItem('laadschemaCustomTrucks', JSON.stringify(state.laadschemaCustomTrucks));
  schedulePersistTeamState();
  
  // Re-render table to show empty cells with only standard columns
  renderLaadschemaTable();
  
  console.log('Cleared all Laadschema data and custom trucks');
}

function openAddTruckModal() {
  const modal = document.getElementById("addTruckModal");
  const clientNameInput = document.getElementById("addTruckClientName");
  const plannedTimeInput = document.getElementById("addTruckPlannedTime");
  const statusDiv = document.getElementById("addTruckStatus");
  
  // Reset form
  clientNameInput.value = "";
  plannedTimeInput.value = "";
  statusDiv.classList.add("hidden");
  statusDiv.textContent = "";
  
  modal.classList.remove("hidden");
}

function closeAddTruckModal() {
  const modal = document.getElementById("addTruckModal");
  modal.classList.add("hidden");
}

function confirmAddTruck() {
  const clientNameInput = document.getElementById("addTruckClientName");
  const plannedTimeInput = document.getElementById("addTruckPlannedTime");
  const statusDiv = document.getElementById("addTruckStatus");
  
  const clientName = clientNameInput.value.trim();
  const plannedTime = plannedTimeInput.value;
  
  // Validate inputs
  if (!clientName) {
    statusDiv.textContent = "Please enter a client name.";
    statusDiv.className = "status-message error";
    statusDiv.classList.remove("hidden");
    return;
  }
  
  if (!plannedTime) {
    statusDiv.textContent = "Please enter a planned departure time.";
    statusDiv.className = "status-message error";
    statusDiv.classList.remove("hidden");
    return;
  }
  
  // Get current selected date
  const currentKey = ensureCurrentLaadschemaCustomTruckStore();
  
  // Initialize custom trucks storage if needed
  if (!state.laadschemaCustomTrucks) {
    state.laadschemaCustomTrucks = {};
  }
  if (!state.laadschemaCustomTrucks[currentKey]) {
    state.laadschemaCustomTrucks[currentKey] = [];
  }
  
  // Add new truck
  state.laadschemaCustomTrucks[currentKey].push({
    clientName: clientName,
    plannedTime: plannedTime
  });
  
  // Persist to localStorage
  localStorage.setItem('laadschemaCustomTrucks', JSON.stringify(state.laadschemaCustomTrucks));
  
  schedulePersistTeamState();
  console.log(`Added custom truck: ${clientName} at ${plannedTime} for ${currentKey}`);
  
  // Re-render table
  renderLaadschemaTable();
  
  // Close modal
  closeAddTruckModal();
}

async function downloadExport(type, dcName = null, options = {}) {
  // Allow Print CMR for Denemark and Edeka Laatzen without email
  const normalizedClient = normalizeText(state.selectedClient);
  const exportClient = getSelectedExportClient(type);
  const allowWithoutEmail = type === "print_cmr" && isPrintCmrClient(normalizedClient);
  const canExportFromPreview = ["selected", "merge"].includes(type) && state.mode === "standard" && state.preview?.orders?.length;

  if (!state.file && !allowWithoutEmail && !canExportFromPreview) {
    statusText.textContent = `The original email file is not available for ${state.selectedClient}. Reload the email to export.`;
    return;
  }
  statusText.textContent = "Generating Excel file...";

  const formData = new FormData();
  const shouldAttachEmail = state.file && !(type === "print_cmr" && options.blankCmr);
  if (shouldAttachEmail) {
    formData.append("email", state.file);
  }
  formData.append("quantityOverrides", JSON.stringify(buildQuantityOverrides()));
  if (canExportFromPreview) {
    formData.append("previewExport", JSON.stringify(buildPreviewForDataTransfer()));
  }
  if (type === "merge") {
    formData.append("mergeOrderIndexes", JSON.stringify(getSelectedExportOrderIndexes()));
  }
  formData.append("selectedClient", exportClient || "");
  const haviUienSettings = getHaviUienSettings();
  formData.append("haviUienArticle", haviUienSettings.article);
  formData.append("haviUienDescription", haviUienSettings.description);
  if (type === "print_cmr") {
    const cmrMeta = buildCmrExportMeta();
    formData.append("cmrReferences", JSON.stringify(cmrMeta.references));
    formData.append("cmrPalletPlaces", cmrMeta.palletPlaces);
    // Add dcName for Havi 3 DCs
    const resolvedDcName = dcName || (normalizedClient === "havi" ? getSelectedHaviDcName() : "");
    if (resolvedDcName) {
      formData.append("dcName", resolvedDcName);
    }
  }
  const exportSheet = isReweOrPennyClientSelected() ? exportSheetSelect?.value || "Paderborn" : "";

  try {
    const orderIndex = type === "print_cmr" ? getCmrContextOrderIndex() : state.selectedIndex;
    const response = await fetch(`/api/export?type=${encodeURIComponent(type)}&orderIndex=${orderIndex}&exportSheet=${encodeURIComponent(exportSheet)}&selectedClient=${encodeURIComponent(exportClient || "")}`, {
      method: "POST",
      body: formData,
    });

    if (response.status === 401) {
      resetPreviewState();
      showAuthScreen();
      return;
    }

    if (!response.ok) {
      const payload = await response.json();
      statusText.textContent = payload.error || "Export failed.";
      return;
    }

    const blob = await response.blob();
    const disposition = response.headers.get("Content-Disposition") || "";
    const match = disposition.match(/filename="(.+)"/);
    const fileName = match ? match[1] : "export.xlsx";
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.click();
    URL.revokeObjectURL(url);
    statusText.textContent = `${fileName} downloaded.`;
  } catch {
    statusText.textContent = "Export failed.";
  }
}

function buildCmrExportMeta() {
  const order = getCmrContextOrder() || getLeverschemaOrder() || getCurrentOrder();
  const references = String(order?.reference || "")
    .split("+")
    .map((value) => value.trim())
    .filter(Boolean);

  let palletPlaces = "";
  if (order) {
    const sheet = getActiveSessionSheet();
    const storageKey = buildLeverschemaOrderKey(order, sheet);
    const saved = state.leverschemaResults?.[storageKey];
    if (saved?.totalPalletPlaces !== undefined && saved?.totalPalletPlaces !== null && saved.totalPalletPlaces !== "") {
      palletPlaces = String(saved.totalPalletPlaces);
    } else {
      const calculated = calculateLeverschema(order);
      if (calculated?.totalPalletPlaces !== undefined && calculated?.totalPalletPlaces !== null) {
        palletPlaces = String(calculated.totalPalletPlaces);
      }
    }
  }

  return {
    references,
    palletPlaces,
  };
}

function getCmrContextOrder() {
  if (!state.preview || state.mode !== "standard" || !state.preview.orders?.length) {
    return null;
  }

  const activePoint = getCmrActiveDeliveryPoint();
  if (!activePoint) return null;

  const groupEntries = getCmrContextOrderEntries(activePoint);
  if (!groupEntries.length) return null;

  const checkedEntries = groupEntries.filter(({ index }) => state.leverschemaIncludedIndexes.includes(index));
  const selectedEntries = checkedEntries.length
    ? checkedEntries
    : state.orderSelectionActive
      ? groupEntries.filter(({ index }) => index === state.selectedIndex)
      : groupEntries;
  const entries = selectedEntries.length ? selectedEntries : [groupEntries[0]];
  const base = entries[0].order;

  return {
    customer: base.customer,
    reference: entries.map(({ order }) => order.reference).filter(Boolean).join(" + "),
    fatrans: base.fatrans,
    deliveryPoint: activePoint,
    deliveryDate: base.deliveryDate || state.preview.deliveryDate || "",
    items: entries.flatMap(({ order }) =>
      (order.items || []).map((item) => ({
        primary: item.primary,
        secondary: item.secondary,
        quantity: parseNumber(item.quantity),
        unit: item.unit,
      })),
    ),
  };
}

function getCmrActiveDeliveryPoint() {
  return state.activeDeliveryPointKey || getDeliveryPointKeyForIndex(state.selectedIndex);
}

function getCmrContextOrderEntries(activePoint = getCmrActiveDeliveryPoint()) {
  if (!state.preview?.orders?.length || !activePoint) return [];
  return state.preview.orders
    .map((order, index) => ({ order: applyStandardClientDisplayRules(order), index }))
    .filter(({ order }) => getDeliveryPointKey(order) === activePoint);
}

function getCmrContextOrderIndex() {
  if (!state.preview || state.mode !== "standard" || !state.preview.orders?.length) {
    return state.selectedIndex;
  }
  const groupEntries = getCmrContextOrderEntries();
  if (!groupEntries.length) return state.selectedIndex;
  const checkedEntry = groupEntries.find(({ index }) => state.leverschemaIncludedIndexes.includes(index));
  if (checkedEntry) return checkedEntry.index;
  const selectedEntry = groupEntries.find(({ index }) => index === state.selectedIndex);
  return (selectedEntry || groupEntries[0]).index;
}

function buildQuantityOverrides() {
  if (!state.preview) return {};
  if (state.mode === "special") {
    return {
      items: state.preview.items.map((item) => sanitizeEditedQuantity(item.quantity)),
      slices: state.preview.items.map((item) => sanitizeEditedQuantity(item.slicesQuantity)),
    };
  }

  return {
    orders: state.preview.orders.map((order) => ({
      items: order.items.map((item) => sanitizeEditedQuantity(item.quantity)),
    })),
  };
}

function sanitizeEditedQuantity(value) {
  const parsed = parseNumber(value);
  return Number.isInteger(parsed) ? String(parsed) : String(Math.round(parsed));
}

function renderPpCell(value, storageKey) {
  if (!value) return "";
  return `
    <div class="pp-cell">
      <span>${escapeHtml(value)}</span>
      <button class="pp-clear-button" type="button" data-clear-leverschema-key="${escapeHtml(storageKey)}">Clear</button>
    </div>
  `;
}

initCustomDatePickers();
bootstrapApp();


// Pakbon Upload Functionality
let pakbonParsedItems = null;

function openCmrChoiceModal() {
  document.getElementById("cmrChoiceModal")?.classList.remove("hidden");
}

function closeCmrChoiceModal() {
  document.getElementById("cmrChoiceModal")?.classList.add("hidden");
}

function shouldUsePakbonModalForCurrentClient() {
  const normalizedClient = normalizeText(state.selectedClient);
  return normalizedClient === "carrefour" || normalizedClient === "carrefour fif" || normalizedClient === "carrefour kdc" || normalizedClient === "colruyt";
}

function handleCmrWithPakbonChoice() {
  closeCmrChoiceModal();
  if (shouldUsePakbonModalForCurrentClient()) {
    openPakbonModal();
    return;
  }
  downloadExport("print_cmr");
}

function handleCmrWithoutPakbonChoice() {
  closeCmrChoiceModal();
  downloadExport("print_cmr", null, { blankCmr: true });
}

function openPakbonModal() {
  const modal = document.getElementById("pakbonModal");
  const fileInput = document.getElementById("pakbonFileInput");
  const fileList = document.getElementById("pakbonFileList");
  const status = document.getElementById("pakbonStatus");
  const preview = document.getElementById("pakbonItemsPreview");
  
  // Reset modal
  fileInput.value = "";
  fileList.innerHTML = "";
  status.innerHTML = "";
  status.className = "status-message";
  preview.classList.add("hidden");
  pakbonParsedItems = null;
  
  modal.classList.remove("hidden");
}

function closePakbonModal() {
  const modal = document.getElementById("pakbonModal");
  modal.classList.add("hidden");
}

async function uploadPakbonFiles() {
  const fileInput = document.getElementById("pakbonFileInput");
  const status = document.getElementById("pakbonStatus");
  const preview = document.getElementById("pakbonItemsPreview");
  const uploadButton = document.getElementById("pakbonUploadButton");
  
  if (!fileInput.files || fileInput.files.length === 0) {
    status.className = "status-message error";
    status.textContent = "Please select at least one PDF file.";
    return;
  }
  
  uploadButton.disabled = true;
  status.className = "status-message";
  status.textContent = "Parsing pakbon files...";
  
  try {
    const formData = new FormData();
    for (let i = 0; i < fileInput.files.length; i++) {
      formData.append(`file${i}`, fileInput.files[i]);
    }
    
    const response = await fetch("/api/parse_pakbon", {
      method: "POST",
      credentials: "include",
      body: formData,
    });
    
    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error || "Failed to parse pakbon files");
    }
    
    const result = await response.json();
    pakbonParsedItems = result.items;
    
    // Show success message
    status.className = "status-message success";
    status.textContent = `Successfully parsed ${result.totalItems} unique items from ${fileInput.files.length} file(s).`;
    
    // Show preview
    const tableBody = document.getElementById("pakbonItemsTable");
    tableBody.innerHTML = result.items.map(item => `
      <tr>
        <td>${escapeHtml(item.articleNumber)}</td>
        <td>${escapeHtml(item.description)}</td>
        <td>${item.quantity}</td>
        <td>${escapeHtml(item.unit)}</td>
      </tr>
    `).join("");
    preview.classList.remove("hidden");
    
    // Now export CMR with the parsed data
    await exportCMRWithPakbonData(result.items, result.goederenTotal, result.chepTotal);
    
  } catch (error) {
    status.className = "status-message error";
    status.textContent = `Error: ${error.message}`;
    uploadButton.disabled = false;
  }
}

async function exportCMRWithPakbonData(items, goederenTotal, chepTotal) {
  const status = document.getElementById("pakbonStatus");
  
  try {
    status.className = "status-message";
    status.textContent = "Generating CMR with pakbon data...";
    
    // Calculate total pallet places from items
    const totalQuantity = items.reduce((sum, item) => sum + item.quantity, 0);
    
    // Prepare CMR export with pakbon data
    const formData = new FormData();
    
    // Add pakbon items as JSON
    formData.append("pakbonItems", JSON.stringify(items));
    formData.append("pakbonTotalQuantity", String(totalQuantity));
    
    // Add CMR metadata
    const cmrMeta = buildCmrExportMeta();
    formData.append("cmrReferences", JSON.stringify(cmrMeta.references));
    
    // Handle PP based on client:
    // - Carrefour KDC: use chepTotal instead of Leverschema PP
    // - Others: use Leverschema PP
    const normalizedClient = normalizeText(state.selectedClient);
    const exportClient = getSelectedExportClient("print_cmr");
    if (normalizeText(exportClient) === "carrefour kdc" && chepTotal !== undefined) {
      formData.append("cmrPalletPlaces", String(chepTotal));
    } else {
      formData.append("cmrPalletPlaces", cmrMeta.palletPlaces);  // Use PP from Leverschema
    }
    
    // For Colruyt, also send goederenTotal for E41
    if (normalizedClient === "colruyt" && goederenTotal !== undefined) {
      formData.append("goederenTotal", String(goederenTotal));
    }
    
    // Add email file if available (optional for CMR)
    if (state.file) {
      formData.append("email", state.file);
    }
    
    // Build query string with export parameters
    const params = new URLSearchParams({
      type: "print_cmr",
      selectedClient: exportClient || state.selectedClient || "",
      orderIndex: String(getCmrContextOrderIndex()),
    });
    
    const response = await fetch(`/api/export?${params.toString()}`, {
      method: "POST",
      credentials: "include",
      body: formData,
    });
    
    if (response.status === 401) {
      resetPreviewState();
      showAuthScreen();
      return;
    }
    
    if (!response.ok) {
      const error = await response.json();
      console.error("CMR Export Error:", error);
      throw new Error(error.error || "CMR export failed");
    }
    
    const blob = await response.blob();
    const disposition = response.headers.get("Content-Disposition") || "";
    const match = disposition.match(/filename="(.+)"/);
    const fileName = match ? match[1] : "CMR.xlsm";
    
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.click();
    URL.revokeObjectURL(url);
    
    status.className = "status-message success";
    status.textContent = `CMR exported successfully: ${fileName}`;
    
    // Close modal after 2 seconds
    setTimeout(() => {
      closePakbonModal();
    }, 2000);
    
  } catch (error) {
    status.className = "status-message error";
    status.textContent = `Export error: ${error.message}`;
  } finally {
    document.getElementById("pakbonUploadButton").disabled = false;
  }
}

// Set up Print CMR button handler to choose whether Pakbon data should be used.
printCmrButton.addEventListener("click", function() {
  openCmrChoiceModal();
});
