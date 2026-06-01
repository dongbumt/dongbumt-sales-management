const STORAGE_KEY = "dongbuEnterpriseSupply.v1";
const SUPABASE_SESSION_KEY = "dongbuEnterpriseSupply.supabaseSession.v1";
const SUPABASE_AUTOSAVE_KEY = "dongbuEnterpriseSupply.supabaseAutoSave.v1";
const SUPABASE_URL = "https://omgxpgcariemnyyobruw.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_dddFKmf-Yn4JFj2B24HWSQ_UF5XgKiL";
const SUPABASE_STATE_ID = "main";
const FILE_DB_NAME = "dongbuEnterpriseSupply.files";
const FILE_DB_STORE = "handles";
const FILE_HANDLE_KEY = "mainDataFile";

const monthByQuarter = {
  Q1: ["01", "02", "03"],
  Q2: ["04", "05", "06"],
  Q3: ["07", "08", "09"],
  Q4: ["10", "11", "12"],
};
const ALL_QUARTER = "ALL";

const viewTitles = {
  dashboard: "대시보드",
  performance: "실적 관리",
  goals: "분기 목표",
  actions: "실행계획",
  accounts: "거래처 분석",
  master: "거래처 마스터",
  items: "품목 분석",
  report: "보고서",
  data: "데이터 관리",
};

const currentDate = new Date();
const currentYear = currentDate.getFullYear();
const currentQuarter = `Q${Math.floor(currentDate.getMonth() / 3) + 1}`;

let state = loadState();
let fileHandle = null;
let fileSaveTimer = null;
let needsFilePermission = false;
let lastFileSavedAt = "";
let supabaseSession = loadSupabaseSession();
let supabaseSaveTimer = null;
let supabaseLastSyncedAt = "";
let supabaseAutoSave = localStorage.getItem(SUPABASE_AUTOSAVE_KEY) !== "false";
const tableSorts = {};
let excelRows = [];
let excelAnalysis = null;
let excelMeta = null;

const el = {
  viewTitle: document.querySelector("#viewTitle"),
  yearSelect: document.querySelector("#yearSelect"),
  quarterSelect: document.querySelector("#quarterSelect"),
  storageStatus: document.querySelector("#storageStatus"),
  kpiGrid: document.querySelector("#kpiGrid"),
  monthlyChart: document.querySelector("#monthlyChart"),
  dashboardPeriodButtons: document.querySelectorAll("[data-dashboard-period]"),
  itemPeriodButtons: document.querySelectorAll("[data-item-period]"),
  trendTitle: document.querySelector("#trendTitle"),
  trendSubtitle: document.querySelector("#trendSubtitle"),
  accountMixSubtitle: document.querySelector("#accountMixSubtitle"),
  dashboardActionsTitle: document.querySelector("#dashboardActionsTitle"),
  dashboardActionsSubtitle: document.querySelector("#dashboardActionsSubtitle"),
  accountMix: document.querySelector("#accountMix"),
  gapBox: document.querySelector("#gapBox"),
  dashboardActions: document.querySelector("#dashboardActions"),
  monthlyForm: document.querySelector("#monthlyForm"),
  monthInput: document.querySelector("#monthInput"),
  accountInput: document.querySelector("#accountInput"),
  channelInput: document.querySelector("#channelInput"),
  salesInput: document.querySelector("#salesInput"),
  costInput: document.querySelector("#costInput"),
  itemsInput: document.querySelector("#itemsInput"),
  accountFilter: document.querySelector("#accountFilter"),
  monthlyRows: document.querySelector("#monthlyRows"),
  monthlySummary: document.querySelector("#monthlySummary"),
  excelAccountInput: document.querySelector("#excelAccountInput"),
  excelMonthInput: document.querySelector("#excelMonthInput"),
  excelFileInput: document.querySelector("#excelFileInput"),
  excelOverwriteInput: document.querySelector("#excelOverwriteInput"),
  excelSupportMessage: document.querySelector("#excelSupportMessage"),
  excelMappingPanel: document.querySelector("#excelMappingPanel"),
  excelResultPanel: document.querySelector("#excelResultPanel"),
  mapDateSelect: document.querySelector("#mapDateSelect"),
  mapAccountSelect: document.querySelector("#mapAccountSelect"),
  mapItemSelect: document.querySelector("#mapItemSelect"),
  mapSalesSelect: document.querySelector("#mapSalesSelect"),
  mapCostSelect: document.querySelector("#mapCostSelect"),
  mapProfitSelect: document.querySelector("#mapProfitSelect"),
  analyzeExcelButton: document.querySelector("#analyzeExcelButton"),
  importExcelResultButton: document.querySelector("#importExcelResultButton"),
  excelResultSummary: document.querySelector("#excelResultSummary"),
  excelKpis: document.querySelector("#excelKpis"),
  excelAnalysisNote: document.querySelector("#excelAnalysisNote"),
  excelResultRows: document.querySelector("#excelResultRows"),
  goalForm: document.querySelector("#goalForm"),
  goalSalesInput: document.querySelector("#goalSalesInput"),
  goalProfitInput: document.querySelector("#goalProfitInput"),
  goalMarginInput: document.querySelector("#goalMarginInput"),
  goalProgress: document.querySelector("#goalProgress"),
  allocationForm: document.querySelector("#allocationForm"),
  allocationEditIdInput: document.querySelector("#allocationEditIdInput"),
  allocationAccountInput: document.querySelector("#allocationAccountInput"),
  allocationSalesInput: document.querySelector("#allocationSalesInput"),
  allocationProfitInput: document.querySelector("#allocationProfitInput"),
  allocationSubmitButton: document.querySelector("#allocationSubmitButton"),
  allocationCancelButton: document.querySelector("#allocationCancelButton"),
  allocationRows: document.querySelector("#allocationRows"),
  managerGoalRows: document.querySelector("#managerGoalRows"),
  extraGoalForm: document.querySelector("#extraGoalForm"),
  extraGoalEditIdInput: document.querySelector("#extraGoalEditIdInput"),
  extraGoalCategoryInput: document.querySelector("#extraGoalCategoryInput"),
  extraGoalTitleInput: document.querySelector("#extraGoalTitleInput"),
  extraGoalTargetInput: document.querySelector("#extraGoalTargetInput"),
  extraGoalActualInput: document.querySelector("#extraGoalActualInput"),
  extraGoalUnitInput: document.querySelector("#extraGoalUnitInput"),
  extraGoalOwnerInput: document.querySelector("#extraGoalOwnerInput"),
  extraGoalPlanInput: document.querySelector("#extraGoalPlanInput"),
  extraGoalStatusInput: document.querySelector("#extraGoalStatusInput"),
  extraGoalSubmitButton: document.querySelector("#extraGoalSubmitButton"),
  extraGoalCancelButton: document.querySelector("#extraGoalCancelButton"),
  extraGoalRows: document.querySelector("#extraGoalRows"),
  actionForm: document.querySelector("#actionForm"),
  actionTitleInput: document.querySelector("#actionTitleInput"),
  actionAccountInput: document.querySelector("#actionAccountInput"),
  actionTypeInput: document.querySelector("#actionTypeInput"),
  actionOwnerInput: document.querySelector("#actionOwnerInput"),
  actionDueInput: document.querySelector("#actionDueInput"),
  actionSalesInput: document.querySelector("#actionSalesInput"),
  actionProfitInput: document.querySelector("#actionProfitInput"),
  actionStatusInput: document.querySelector("#actionStatusInput"),
  statusFilter: document.querySelector("#statusFilter"),
  actionRows: document.querySelector("#actionRows"),
  actionSummary: document.querySelector("#actionSummary"),
  accountChartAccountSelect: document.querySelector("#accountChartAccountSelect"),
  accountChartSubtitle: document.querySelector("#accountChartSubtitle"),
  accountTrendChart: document.querySelector("#accountTrendChart"),
  accountAnalysisSummary: document.querySelector("#accountAnalysisSummary"),
  accountRows: document.querySelector("#accountRows"),
  managerSalesRows: document.querySelector("#managerSalesRows"),
  itemMonthInput: document.querySelector("#itemMonthInput"),
  itemAccountFilter: document.querySelector("#itemAccountFilter"),
  cleanOrphanItemsButton: document.querySelector("#cleanOrphanItemsButton"),
  itemAnalysisKpis: document.querySelector("#itemAnalysisKpis"),
  itemAnalysisSummary: document.querySelector("#itemAnalysisSummary"),
  itemAnalysisRows: document.querySelector("#itemAnalysisRows"),
  accountForm: document.querySelector("#accountForm"),
  accountNameInput: document.querySelector("#accountNameInput"),
  accountOwnerInput: document.querySelector("#accountOwnerInput"),
  accountAliasInput: document.querySelector("#accountAliasInput"),
  accountMasterSummary: document.querySelector("#accountMasterSummary"),
  accountMasterNote: document.querySelector("#accountMasterNote"),
  accountMasterRows: document.querySelector("#accountMasterRows"),
  recanonicalizeButton: document.querySelector("#recanonicalizeButton"),
  reportBody: document.querySelector("#reportBody"),
  printButton: document.querySelector("#printButton"),
  supabaseConnectionCard: document.querySelector("#supabaseConnectionCard"),
  supabaseEmailInput: document.querySelector("#supabaseEmailInput"),
  supabasePasswordInput: document.querySelector("#supabasePasswordInput"),
  supabaseLoginButton: document.querySelector("#supabaseLoginButton"),
  supabaseLogoutButton: document.querySelector("#supabaseLogoutButton"),
  supabaseLoadButton: document.querySelector("#supabaseLoadButton"),
  supabaseSaveButton: document.querySelector("#supabaseSaveButton"),
  supabaseAutoSaveInput: document.querySelector("#supabaseAutoSaveInput"),
  fileConnectionCard: document.querySelector("#fileConnectionCard"),
  createDataFileButton: document.querySelector("#createDataFileButton"),
  openDataFileButton: document.querySelector("#openDataFileButton"),
  reloadDataFileButton: document.querySelector("#reloadDataFileButton"),
  disconnectDataFileButton: document.querySelector("#disconnectDataFileButton"),
  exportJsonButton: document.querySelector("#exportJsonButton"),
  importJsonInput: document.querySelector("#importJsonInput"),
  dataStats: document.querySelector("#dataStats"),
  toastRoot: document.querySelector("#toastRoot"),
};

init();

function init() {
  fillYearOptions();
  el.yearSelect.value = state.activeYear;
  el.quarterSelect.value = state.activeQuarter;
  el.monthInput.value = getDefaultMonthInputValue();
  el.actionDueInput.value = getDefaultDueDate();

  document.querySelectorAll(".nav-item").forEach((button) => {
    button.addEventListener("click", () => setView(button.dataset.view));
  });

  el.dashboardPeriodButtons.forEach((button) => {
    button.addEventListener("click", () => {
      state.dashboardPeriod = button.dataset.dashboardPeriod;
      saveState({ localOnly: true });
      renderAll();
    });
  });

  bindAnalysisPeriodControls("item", el.itemPeriodButtons, el.itemMonthInput, renderItemAnalysis);
  bindSortableTables();

  el.yearSelect.addEventListener("change", () => {
    state.activeYear = Number(el.yearSelect.value);
    syncAnalysisMonthYears();
    saveState();
    renderAll();
  });

  el.quarterSelect.addEventListener("change", () => {
    state.activeQuarter = el.quarterSelect.value;
    el.monthInput.value = getDefaultMonthInputValue();
    saveState();
    renderAll();
  });

  el.accountFilter.addEventListener("change", renderMonthly);
  el.accountChartAccountSelect.addEventListener("change", () => {
    state.accountChartAccount = el.accountChartAccountSelect.value;
    saveState({ localOnly: true });
    renderAccounts();
  });
  el.itemAccountFilter.addEventListener("change", renderItemAnalysis);
  el.statusFilter.addEventListener("change", renderActions);
  el.printButton.addEventListener("click", () => window.print());
  el.excelFileInput.addEventListener("change", handleExcelFile);
  el.analyzeExcelButton.addEventListener("click", analyzeExcelRows);
  el.importExcelResultButton.addEventListener("click", importExcelAnalysis);

  bindForms();
  bindDataActions();
  bindAccountMaster();
  seedAccountsIfEmpty();
  renderAll();
  restoreFileHandle();
}

function bindAnalysisPeriodControls(kind, buttons, monthInput, renderFn) {
  buttons.forEach((button) => {
    button.addEventListener("click", () => {
      state[`${kind}AnalysisPeriod`] = button.dataset[`${kind}Period`];
      saveState({ localOnly: true });
      renderAll();
    });
  });

  monthInput.addEventListener("change", () => {
    state[`${kind}AnalysisMonth`] = monthInput.value || getDefaultAnalysisMonth();
    saveState({ localOnly: true });
    renderFn();
  });
}

function bindSortableTables() {
  document.addEventListener("click", (event) => {
    const header = event.target.closest("th[data-sort-table][data-sort-key]");
    if (!header) return;
    const table = header.dataset.sortTable;
    const key = header.dataset.sortKey;
    const current = tableSorts[table];
    tableSorts[table] = {
      key,
      direction: current?.key === key && current.direction === "asc" ? "desc" : "asc",
      type: header.dataset.sortType || "text",
    };
    renderSortedTable(table);
    updateSortHeaders();
  });
}

function renderSortedTable(table) {
  const renderers = {
    monthly: renderMonthly,
    excelResult: renderExcelAnalysis,
    allocation: renderGoals,
    managerGoal: renderGoals,
    extraGoals: renderGoals,
    actions: renderActions,
    accounts: renderAccounts,
    managerSales: renderAccounts,
    accountMaster: renderAccountMaster,
    itemAnalysis: renderItemAnalysis,
  };
  renderers[table]?.();
}

function syncAnalysisMonthYears() {
  ["account", "item"].forEach((kind) => {
    const key = `${kind}AnalysisMonth`;
    if (!String(state[key] || "").startsWith(`${state.activeYear}-`)) {
      state[key] = getDefaultAnalysisMonth();
    }
  });
}

function getDefaultAnalysisMonth() {
  return getDefaultMonthInputValue();
}

function getQuarterMonthCodes(quarter) {
  if (quarter === ALL_QUARTER) return Object.values(monthByQuarter).flat();
  return monthByQuarter[quarter] || monthByQuarter[currentQuarter];
}

function getSelectedMonthKeys() {
  return getQuarterMonthCodes(state.activeQuarter).map((month) => `${state.activeYear}-${month}`);
}

function getDefaultMonthInputValue() {
  return `${state.activeYear}-${getQuarterMonthCodes(state.activeQuarter)[0]}`;
}

function getSelectedPeriodLabel() {
  return `${state.activeYear}년 ${quarterLabel(state.activeQuarter)}`;
}

function loadState() {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (saved) {
    return prepareState(JSON.parse(saved));
  }

  return prepareState(createDefaultState());
}

function createDefaultState() {
  return {
    version: "1.1",
    lastUpdated: "",
    activeYear: currentYear,
    activeQuarter: currentQuarter,
    dashboardPeriod: "quarter",
    accountAnalysisPeriod: "quarter",
    accountAnalysisMonth: `${currentYear}-${monthByQuarter[currentQuarter][0]}`,
    accountChartAccount: "전체",
    itemAnalysisPeriod: "quarter",
    itemAnalysisMonth: `${currentYear}-${monthByQuarter[currentQuarter][0]}`,
    accounts: [],
    records: [
      sampleRecord("2026-04", "한빛푸드", "대기업납품", 118000000, 101200000, "돈육 전지, 삼겹"),
      sampleRecord("2026-04", "태성FS", "대기업납품", 82000000, 72000000, "우삼겹, 목심"),
      sampleRecord("2026-04", "그린유통", "품목확대", 57000000, 48600000, "돈육 후지"),
      sampleRecord("2026-05", "한빛푸드", "대기업납품", 126000000, 107700000, "돈육 전지, 삼겹"),
      sampleRecord("2026-05", "태성FS", "대기업납품", 87000000, 75400000, "우삼겹, 목심"),
      sampleRecord("2026-05", "그린유통", "품목확대", 62000000, 52600000, "돈육 후지"),
      sampleRecord("2026-06", "한빛푸드", "대기업납품", 132000000, 112500000, "돈육 전지, 삼겹"),
      sampleRecord("2026-06", "태성FS", "대기업납품", 91000000, 79100000, "우삼겹, 목심"),
      sampleRecord("2026-06", "그린유통", "신규제안", 69000000, 58400000, "돈육 후지, 갈비"),
    ],
    itemRecords: [],
    goals: {
      "2026-Q2": {
        sales: 900000000,
        profit: 112000000,
        margin: 12.4,
        allocations: [
          { id: cryptoId(), account: "한빛푸드", sales: 390000000, profit: 50000000 },
          { id: cryptoId(), account: "태성FS", sales: 270000000, profit: 32000000 },
          { id: cryptoId(), account: "그린유통", sales: 210000000, profit: 27000000 },
        ],
        extraGoals: [],
      },
    },
    actions: [
      sampleAction("한빛푸드 월 발주량 8% 확대", "한빛푸드", "발주량 확대", "영업1팀", "2026-06-12", 30000000, 4200000, "진행중"),
      sampleAction("태성FS 냉동 우삼겹 단가 재협상", "태성FS", "단가 재협상", "영업2팀", "2026-06-18", 0, 3500000, "예정"),
      sampleAction("그린유통 갈비 품목 정규 편성", "그린유통", "신규 품목 제안", "상품기획", "2026-06-25", 24000000, 3100000, "진행중"),
    ],
  };
}

function normalizeState(input) {
  const fallback = createDefaultState();
  const source = input && typeof input === "object" ? input : fallback;
  const hasKnownQuarter = source.activeQuarter === ALL_QUARTER || Boolean(monthByQuarter[source.activeQuarter]);
  const activeQuarter = hasKnownQuarter ? source.activeQuarter : fallback.activeQuarter;
  const activeYear = Number(source.activeYear) || fallback.activeYear;

  const normalized = {
    version: source.version || "1.1",
    lastUpdated: source.lastUpdated || "",
    activeYear,
    activeQuarter,
    dashboardPeriod: ["quarter", "year", "fiveYear"].includes(source.dashboardPeriod) ? source.dashboardPeriod : fallback.dashboardPeriod,
    accountAnalysisPeriod: ["month", "quarter", "year"].includes(source.accountAnalysisPeriod) ? source.accountAnalysisPeriod : fallback.accountAnalysisPeriod,
    accountAnalysisMonth: normalizeAnalysisMonth(source.accountAnalysisMonth, activeYear, activeQuarter),
    accountChartAccount: source.accountChartAccount ? String(source.accountChartAccount) : fallback.accountChartAccount,
    itemAnalysisPeriod: ["month", "quarter", "year"].includes(source.itemAnalysisPeriod) ? source.itemAnalysisPeriod : fallback.itemAnalysisPeriod,
    itemAnalysisMonth: normalizeAnalysisMonth(source.itemAnalysisMonth, activeYear, activeQuarter),
    records: Array.isArray(source.records) ? source.records.map(normalizeRecord).filter(Boolean) : fallback.records,
    itemRecords: Array.isArray(source.itemRecords) ? source.itemRecords.map(normalizeItemRecord).filter(Boolean) : fallback.itemRecords,
    goals: source.goals && typeof source.goals === "object" && !Array.isArray(source.goals) ? source.goals : fallback.goals,
    actions: Array.isArray(source.actions) ? source.actions.map(normalizeAction).filter(Boolean) : fallback.actions,
    accounts: Array.isArray(source.accounts) ? source.accounts.map(normalizeAccount).filter(Boolean) : [],
  };
  // 거래처 마스터가 비어 있으면 실적/목표/액션의 거래처명으로 자동 채움(모든 로드 경로 공통)
  if (!normalized.accounts.length) {
    normalized.accounts = deriveAccountsFromData(normalized);
  }
  return normalized;
}

function prepareState(input) {
  return reconcileRecordsWithItemDetails(normalizeState(input));
}

function reconcileRecordsWithItemDetails(targetState = state) {
  if (!targetState || !Array.isArray(targetState.records) || !Array.isArray(targetState.itemRecords)) {
    return targetState;
  }

  const detailMap = new Map();
  targetState.itemRecords.forEach((item) => {
    const account = getCanonicalAccountName(targetState, item.account);
    const key = `${item.month}__${normalizeAccountKey(account)}`;
    if (!detailMap.has(key)) {
      detailMap.set(key, {
        month: item.month,
        account,
        sales: 0,
        cost: 0,
        items: new Map(),
      });
    }
    const detail = detailMap.get(key);
    const sales = Number(item.sales) || 0;
    const cost = Number(item.cost) || 0;
    const itemKey = normalizeItemKey(item.item);
    const itemSummary = detail.items.get(itemKey) || { name: item.item, sales: 0 };
    detail.sales += sales;
    detail.cost += cost;
    itemSummary.sales += sales;
    detail.items.set(itemKey, itemSummary);
  });

  if (!detailMap.size) return targetState;

  const usedDetailKeys = new Set();
  const nextRecords = [];

  targetState.records.forEach((record) => {
    const account = getCanonicalAccountName(targetState, record.account);
    const key = `${record.month}__${normalizeAccountKey(account)}`;
    const detail = detailMap.get(key);

    if (!detail) {
      nextRecords.push(record);
      return;
    }

    if (usedDetailKeys.has(key)) return;
    usedDetailKeys.add(key);
    nextRecords.push({
      ...record,
      account: detail.account,
      sales: Math.round(detail.sales),
      cost: Math.round(detail.cost),
      items: summarizeDetailItems(detail.items),
    });
  });

  detailMap.forEach((detail, key) => {
    if (usedDetailKeys.has(key)) return;
    nextRecords.push({
      id: cryptoId(),
      month: detail.month,
      account: detail.account,
      channel: "\uB300\uAE30\uC5C5\uB0A9\uD488",
      sales: Math.round(detail.sales),
      cost: Math.round(detail.cost),
      items: summarizeDetailItems(detail.items),
    });
  });

  targetState.records = nextRecords;
  return targetState;
}

function getCanonicalAccountName(targetState, raw) {
  const text = getCellText(raw);
  const key = normalizeAccountKey(text);
  if (!key) return text;
  const accounts = Array.isArray(targetState.accounts) ? targetState.accounts : [];

  for (const account of accounts) {
    if (normalizeAccountKey(account.name) === key) return account.name;
    if ((account.aliases || []).some((alias) => normalizeAccountKey(alias) === key)) return account.name;
  }

  for (const account of accounts) {
    if (accountsLookSame(normalizeAccountKey(account.name), key)) return account.name;
    if ((account.aliases || []).some((alias) => accountsLookSame(normalizeAccountKey(alias), key))) return account.name;
  }

  return text;
}

function summarizeDetailItems(items) {
  return [...items.values()]
    .sort((a, b) => b.sales - a.sales || a.name.localeCompare(b.name, "ko"))
    .slice(0, 5)
    .map((item) => item.name)
    .join(", ");
}

// 실적/품목/목표배분/액션에 등장하는 거래처명을 모아 마스터 초기 목록 생성
function deriveAccountsFromData(s) {
  const names = new Set();
  (s.records || []).forEach((r) => r.account && names.add(String(r.account).trim()));
  (s.itemRecords || []).forEach((r) => r.account && names.add(String(r.account).trim()));
  Object.values(s.goals || {}).forEach((goal) => (goal.allocations || []).forEach((a) => a.account && names.add(String(a.account).trim())));
  (s.actions || []).forEach((a) => a.account && names.add(String(a.account).trim()));
  return [...names].filter(Boolean).sort((a, b) => a.localeCompare(b, "ko")).map((name) => ({ id: cryptoId(), name, owner: "", aliases: [] }));
}

function normalizeAccount(account) {
  if (!account || !account.name) return null;
  const aliases = Array.isArray(account.aliases) ? [...new Set(account.aliases.map((a) => getCellText(a)).filter(Boolean))] : [];
  return {
    id: account.id || cryptoId(),
    name: String(account.name).trim(),
    owner: getCellText(account.owner || account.manager),
    aliases,
  };
}

// 마스터가 비어 있으면 기존 실적/목표/액션의 거래처명으로 자동 시드
function seedAccountsIfEmpty() {
  if (Array.isArray(state.accounts) && state.accounts.length) return;
  state.accounts = deriveAccountsFromData(state);
}

function normalizeAnalysisMonth(value, year, quarter) {
  const text = String(value || "");
  const match = text.match(/^(\d{4})-(\d{2})$/);
  if (match && Number(match[2]) >= 1 && Number(match[2]) <= 12) return text;
  return `${year}-${getQuarterMonthCodes(quarter)[0]}`;
}

function normalizeRecord(record) {
  if (!record || !record.month || !record.account) return null;
  return {
    id: record.id || cryptoId(),
    month: record.month,
    account: String(record.account),
    channel: record.channel || "대기업납품",
    sales: Number(record.sales) || 0,
    cost: Number(record.cost) || 0,
    items: record.items || "",
  };
}

function normalizeItemRecord(record) {
  if (!record || !record.month || !record.account || !record.item) return null;
  const sales = Number(record.sales) || 0;
  const rawCost = Number(record.cost);
  const rawProfit = Number(record.profit);
  const cost = Number.isFinite(rawCost) ? rawCost : Number.isFinite(rawProfit) ? sales - rawProfit : sales;
  return {
    id: record.id || cryptoId(),
    month: record.month,
    account: String(record.account),
    item: String(record.item),
    sales,
    cost,
  };
}

function normalizeAction(action) {
  if (!action || !action.title) return null;
  return {
    id: action.id || cryptoId(),
    title: String(action.title),
    account: action.account || "",
    type: action.type || "발주량 확대",
    owner: action.owner || "",
    due: action.due || `${currentYear}-${monthByQuarter[currentQuarter][2]}-25`,
    expectedSales: Number(action.expectedSales) || 0,
    expectedProfit: Number(action.expectedProfit) || 0,
    status: action.status || "예정",
  };
}

function normalizeExtraGoal(goal) {
  if (!goal || !goal.title) return null;
  return {
    id: goal.id || cryptoId(),
    category: goal.category || "기타",
    title: String(goal.title).trim(),
    target: Number(goal.target) || 0,
    actual: Number(goal.actual) || 0,
    unit: getCellText(goal.unit) || "건",
    owner: getCellText(goal.owner),
    plan: getCellText(goal.plan),
    status: goal.status || "예정",
  };
}

function saveState(options = {}) {
  syncAllGoalTargets();
  state.version = "1.1";
  state.lastUpdated = formatTimestamp();
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  updateStorageUi();
  if (fileHandle && !options.localOnly) {
    queueFileSave();
  }
  if (supabaseSession && supabaseAutoSave && !options.localOnly && !options.skipCloud) {
    queueSupabaseSave();
  }
}

function sampleRecord(month, account, channel, sales, cost, items) {
  return { id: cryptoId(), month, account, channel, sales, cost, items };
}

function sampleAction(title, account, type, owner, due, expectedSales, expectedProfit, status) {
  return { id: cryptoId(), title, account, type, owner, due, expectedSales, expectedProfit, status };
}

function cryptoId() {
  if (window.crypto?.randomUUID) {
    return window.crypto.randomUUID();
  }
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function fillYearOptions() {
  const years = new Set([currentYear - 1, currentYear, currentYear + 1, ...state.records.map((record) => Number(record.month.slice(0, 4)))]);
  el.yearSelect.innerHTML = [...years]
    .sort((a, b) => a - b)
    .map((year) => `<option value="${year}">${year}년</option>`)
    .join("");
}

function bindForms() {
  el.monthlyForm.addEventListener("submit", (event) => {
    event.preventDefault();
    state.records.push({
      id: cryptoId(),
      month: el.monthInput.value,
      account: ensureAccountFromImport(el.accountInput.value.trim()),
      channel: el.channelInput.value,
      sales: Number(el.salesInput.value),
      cost: Number(el.costInput.value),
      items: el.itemsInput.value.trim(),
    });
    state.activeYear = Number(el.monthInput.value.slice(0, 4));
    state.activeQuarter = quarterFromMonth(el.monthInput.value);
    el.yearSelect.value = state.activeYear;
    el.quarterSelect.value = state.activeQuarter;
    el.monthlyForm.reset();
    el.monthInput.value = getDefaultMonthInputValue();
    saveState();
    fillYearOptions();
    renderAll();
  });

  el.goalForm.addEventListener("submit", (event) => {
    event.preventDefault();
    syncGoalTargets(getGoal());
    saveState();
    renderAll();
  });

  el.allocationForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const goal = getGoal();
    const editId = el.allocationEditIdInput.value;
    const payload = {
      account: ensureAccountFromImport(el.allocationAccountInput.value.trim()),
      sales: Number(el.allocationSalesInput.value),
      profit: Number(el.allocationProfitInput.value),
    };
    const existing = editId ? goal.allocations.find((allocation) => allocation.id === editId) : null;
    if (existing) {
      Object.assign(existing, payload);
    } else {
      goal.allocations.push({ id: cryptoId(), ...payload });
    }
    syncGoalTargets(goal);
    resetAllocationForm();
    saveState();
    renderAll();
  });

  el.allocationCancelButton.addEventListener("click", resetAllocationForm);

  if (el.extraGoalForm) {
    el.extraGoalForm.addEventListener("submit", (event) => {
      event.preventDefault();
      const goal = getGoal();
      const editId = el.extraGoalEditIdInput.value;
      const payload = normalizeExtraGoal({
        id: editId || cryptoId(),
        category: el.extraGoalCategoryInput.value,
        title: el.extraGoalTitleInput.value.trim(),
        target: Number(el.extraGoalTargetInput.value),
        actual: Number(el.extraGoalActualInput.value),
        unit: el.extraGoalUnitInput.value.trim(),
        owner: el.extraGoalOwnerInput.value.trim(),
        plan: el.extraGoalPlanInput.value.trim(),
        status: el.extraGoalStatusInput.value,
      });
      if (!payload) return;
      const existing = editId ? goal.extraGoals.find((item) => item.id === editId) : null;
      if (existing) {
        Object.assign(existing, payload);
      } else {
        goal.extraGoals.push(payload);
      }
      resetExtraGoalForm();
      saveState();
      renderAll();
    });
    el.extraGoalCancelButton.addEventListener("click", resetExtraGoalForm);
  }

  el.actionForm.addEventListener("submit", (event) => {
    event.preventDefault();
    state.actions.push({
      id: cryptoId(),
      title: el.actionTitleInput.value.trim(),
      account: ensureAccountFromImport(el.actionAccountInput.value.trim()),
      type: el.actionTypeInput.value,
      owner: el.actionOwnerInput.value.trim(),
      due: el.actionDueInput.value,
      expectedSales: Number(el.actionSalesInput.value),
      expectedProfit: Number(el.actionProfitInput.value),
      status: el.actionStatusInput.value,
    });
    el.actionForm.reset();
    el.actionDueInput.value = getDefaultDueDate();
    saveState();
    renderAll();
  });
}

function bindDataActions() {
  el.supabaseLoginButton.addEventListener("click", loginSupabase);
  el.supabaseLogoutButton.addEventListener("click", logoutSupabase);
  el.supabaseLoadButton.addEventListener("click", () => loadStateFromSupabase({ showMessage: true }));
  el.supabaseSaveButton.addEventListener("click", () => saveStateToSupabase({ showMessage: true }));
  el.supabaseAutoSaveInput.addEventListener("change", () => {
    supabaseAutoSave = el.supabaseAutoSaveInput.checked;
    localStorage.setItem(SUPABASE_AUTOSAVE_KEY, String(supabaseAutoSave));
    renderData();
  });
  el.createDataFileButton.addEventListener("click", connectNewDataFile);
  el.openDataFileButton.addEventListener("click", connectExistingDataFile);
  el.reloadDataFileButton.addEventListener("click", () => loadStateFromFile({ showMessage: true }));
  el.disconnectDataFileButton.addEventListener("click", disconnectDataFile);
  el.exportJsonButton.addEventListener("click", exportJsonBackup);
  el.importJsonInput.addEventListener("change", () => {
    const file = el.importJsonInput.files?.[0];
    if (file) importJsonBackup(file);
    el.importJsonInput.value = "";
  });
}

function setView(viewId) {
  document.querySelectorAll(".nav-item").forEach((button) => button.classList.toggle("active", button.dataset.view === viewId));
  document.querySelectorAll(".view").forEach((view) => view.classList.toggle("active", view.id === viewId));
  el.viewTitle.textContent = viewTitles[viewId];
  renderAll();
}

function renderAll() {
  ensureGoal();
  renderDashboard();
  renderMonthly();
  renderExcel();
  renderGoals();
  renderActions();
  renderAccounts();
  renderAccountMaster();
  renderItemAnalysis();
  renderReport();
  renderData();
  updateStorageUi();
  updateSortHeaders();
}

function getSelectedKey() {
  return `${state.activeYear}-${state.activeQuarter}`;
}

function ensureGoal() {
  const key = getSelectedKey();
  if (!state.goals[key]) {
    state.goals[key] = { sales: 0, profit: 0, margin: 0, allocations: [], extraGoals: [] };
  }
  state.goals[key].allocations = Array.isArray(state.goals[key].allocations) ? state.goals[key].allocations : [];
  state.goals[key].extraGoals = Array.isArray(state.goals[key].extraGoals) ? state.goals[key].extraGoals.map(normalizeExtraGoal).filter(Boolean) : [];
}

function getGoal() {
  ensureGoal();
  return state.goals[getSelectedKey()];
}

function getGoalTargets(goal = getGoal()) {
  const allocations = Array.isArray(goal.allocations) ? goal.allocations : [];
  const sales = sum(allocations, "sales");
  const profit = sum(allocations, "profit");
  return {
    ...goal,
    sales,
    profit,
    margin: margin(sales, profit),
  };
}

function syncGoalTargets(goal = getGoal()) {
  goal.allocations = Array.isArray(goal.allocations) ? goal.allocations : [];
  goal.extraGoals = Array.isArray(goal.extraGoals) ? goal.extraGoals.map(normalizeExtraGoal).filter(Boolean) : [];
  const target = getGoalTargets(goal);
  goal.sales = target.sales;
  goal.profit = target.profit;
  goal.margin = target.margin;
  return target;
}

function syncAllGoalTargets() {
  Object.values(state.goals || {}).forEach((goal) => syncGoalTargets(goal));
}

function getQuarterRecords() {
  const months = getSelectedMonthKeys();
  return state.records.filter((record) => months.includes(record.month));
}

function getQuarterItemRecords() {
  const months = getSelectedMonthKeys();
  return state.itemRecords.filter((record) => months.includes(record.month));
}

function getAnalysisRecords(records, period, month) {
  if (period === "month") {
    return records.filter((record) => record.month === month);
  }
  if (period === "year") {
    return records.filter((record) => record.month.startsWith(`${state.activeYear}-`));
  }
  const months = getSelectedMonthKeys();
  return records.filter((record) => months.includes(record.month));
}

function getAnalysisLabel(period, month) {
  if (period === "month") return formatMonth(month);
  if (period === "year") return `${state.activeYear}년`;
  return `${state.activeYear}년 ${quarterLabel(state.activeQuarter)}`;
}

function renderAnalysisControls(kind, buttons, monthInput) {
  const period = state[`${kind}AnalysisPeriod`] || "quarter";
  const month = state[`${kind}AnalysisMonth`] || getDefaultAnalysisMonth();
  buttons.forEach((button) => {
    button.classList.toggle("active", button.dataset[`${kind}Period`] === period);
  });
  monthInput.value = month;
  monthInput.disabled = period !== "month";
}

function getQuarterActions() {
  const months = getSelectedMonthKeys();
  return state.actions.filter((action) => months.includes(action.due.slice(0, 7)));
}

function sum(records, field) {
  return records.reduce((total, record) => total + Number(record[field] || 0), 0);
}

function getTotals(records = getQuarterRecords()) {
  const sales = sum(records, "sales");
  const cost = sum(records, "cost");
  const profit = sales - cost;
  const margin = sales > 0 ? (profit / sales) * 100 : 0;
  return { sales, cost, profit, margin };
}

function getAccountOwner(accountName) {
  const resolved = resolveAccountName(accountName);
  return resolved.account?.owner || "미지정";
}

function createManagerBucket(owner) {
  return {
    owner: owner || "미지정",
    accountCount: 0,
    accountSet: new Set(),
    sales: 0,
    cost: 0,
    profit: 0,
    margin: 0,
    salesTarget: 0,
    profitTarget: 0,
  };
}

function groupRecordsByManager(records) {
  return records.reduce((map, record) => {
    const owner = getAccountOwner(record.account);
    if (!map[owner]) map[owner] = createManagerBucket(owner);
    const row = map[owner];
    row.sales += Number(record.sales) || 0;
    row.cost += Number(record.cost) || 0;
    row.profit += (Number(record.sales) || 0) - (Number(record.cost) || 0);
    row.accountSet.add(resolveAccountName(record.account).name || record.account);
    row.accountCount = row.accountSet.size;
    row.margin = margin(row.sales, row.profit);
    return map;
  }, {});
}

function renderDashboard() {
  const context = getDashboardContext();
  const goal = context.goal;
  const totals = getTotals(context.records);
  const actions = context.actions;
  context.totals = totals;
  const salesRate = rate(totals.sales, goal.sales);
  const profitRate = rate(totals.profit, goal.profit);
  const completed = actions.filter((action) => action.status === "완료").length;
  const pending = actions.filter((action) => action.status !== "완료");

  el.kpiGrid.innerHTML = [
    kpiCard(`${context.kpiPrefix} 매출`, won(totals.sales), `${won(goal.sales)} 목표 · ${salesRate.toFixed(1)}%`),
    kpiCard(`${context.kpiPrefix} 매출이익`, won(totals.profit), `${won(goal.profit)} 목표 · ${profitRate.toFixed(1)}%`),
    kpiCard("기간 마진율", `${totals.margin.toFixed(1)}%`, `${goal.margin.toFixed(1)}% 목표`),
    kpiCard("실행계획 완료", `${completed}/${actions.length}`, `${won(sum(pending, "expectedSales"))} 미완료 기대효과`),
  ].join("");

  el.dashboardPeriodButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.dashboardPeriod === context.period);
  });
  el.trendTitle.textContent = context.trendTitle;
  el.trendSubtitle.textContent = context.trendSubtitle;
  el.accountMixSubtitle.textContent = context.mixSubtitle;
  el.dashboardActionsTitle.textContent = context.actionTitle;
  el.dashboardActionsSubtitle.textContent = context.actionSubtitle;

  renderTrendChart(context);
  renderAccountMix(context);
  renderGapBox(context);
  renderDashboardActions(context);
}

function kpiCard(label, value, meta) {
  return `<article class="kpi-card"><span>${label}</span><strong>${value}</strong><p>${meta}</p></article>`;
}

function getDashboardContext() {
  const period = state.dashboardPeriod || "quarter";

  if (period === "year") {
    const monthKeys = Array.from({ length: 12 }, (_, index) => `${state.activeYear}-${String(index + 1).padStart(2, "0")}`);
    const records = state.records.filter((record) => record.month.startsWith(`${state.activeYear}-`));
    const actions = state.actions.filter((action) => action.due.startsWith(`${state.activeYear}-`));
    return {
      period,
      label: `${state.activeYear}년`,
      kpiPrefix: "연도",
      records,
      actions,
      goal: aggregateGoals([`${state.activeYear}-Q1`, `${state.activeYear}-Q2`, `${state.activeYear}-Q3`, `${state.activeYear}-Q4`]),
      trendTitle: "월별 실적 흐름",
      trendSubtitle: `${state.activeYear}년 월별 매출액과 매출이익`,
      mixSubtitle: "연도 매출 기준",
      actionTitle: "올해 핵심 실행계획",
      actionSubtitle: "미완료 항목 우선",
      emptyLabel: "선택한 연도",
      trendItems: monthKeys.map((key) => ({
        key,
        label: formatMonth(key),
        ...getTotals(records.filter((record) => record.month === key)),
      })),
    };
  }

  if (period === "fiveYear") {
    const endYear = state.activeYear;
    const years = Array.from({ length: 5 }, (_, index) => endYear - 4 + index);
    const records = state.records.filter((record) => years.includes(Number(record.month.slice(0, 4))));
    const actions = state.actions.filter((action) => years.includes(Number(action.due.slice(0, 4))));
    return {
      period,
      label: `${years[0]}~${endYear}년`,
      kpiPrefix: "최근 5년",
      records,
      actions,
      goal: aggregateGoals(years.flatMap((year) => [`${year}-Q1`, `${year}-Q2`, `${year}-Q3`, `${year}-Q4`])),
      trendTitle: "연도별 실적 흐름",
      trendSubtitle: `${years[0]}~${endYear}년 매출액과 매출이익`,
      mixSubtitle: "최근 5년 매출 기준",
      actionTitle: "최근 5년 핵심 실행계획",
      actionSubtitle: "미완료 항목 우선",
      emptyLabel: "최근 5년",
      trendItems: years.map((year) => ({
        key: String(year),
        label: `${year}년`,
        ...getTotals(records.filter((record) => Number(record.month.slice(0, 4)) === year)),
      })),
    };
  }

  const monthKeys = getSelectedMonthKeys();
  const records = state.records.filter((record) => monthKeys.includes(record.month));
  const actions = state.actions.filter((action) => monthKeys.includes(action.due.slice(0, 7)));
  const isAllPeriod = state.activeQuarter === ALL_QUARTER;
  const label = getSelectedPeriodLabel();
  return {
    period: "quarter",
    label,
    kpiPrefix: isAllPeriod ? "전체" : "분기",
    records,
    actions,
    goal: isAllPeriod ? aggregateGoals(getYearGoalKeys(state.activeYear)) : getGoalTargets(getGoal()),
    trendTitle: "월별 실적 흐름",
    trendSubtitle: `${label} 매출액과 매출이익`,
    mixSubtitle: isAllPeriod ? "전체 매출 기준" : "분기 매출 기준",
    actionTitle: isAllPeriod ? "선택 연도 핵심 실행계획" : "이번 분기 핵심 실행계획",
    actionSubtitle: "미완료 항목 우선",
    emptyLabel: isAllPeriod ? "선택한 연도" : "선택한 분기",
    trendItems: monthKeys.map((key) => ({
      key,
      label: formatMonth(key),
      ...getTotals(records.filter((record) => record.month === key)),
    })),
  };
}

function aggregateGoals(keys) {
  const totals = keys.reduce(
    (result, key) => {
      const goal = state.goals[key];
      if (!goal) return result;
      const target = getGoalTargets(goal);
      result.sales += target.sales;
      result.profit += target.profit;
      return result;
    },
    { sales: 0, profit: 0 },
  );
  return {
    sales: totals.sales,
    profit: totals.profit,
    margin: margin(totals.sales, totals.profit),
  };
}

function getYearGoalKeys(year) {
  return ["Q1", "Q2", "Q3", "Q4"].map((quarter) => `${year}-${quarter}`);
}

function renderTrendChart(context) {
  const maxSales = Math.max(...context.trendItems.map((item) => item.sales), 1);

  el.monthlyChart.innerHTML = context.trendItems
    .map((item) => {
      const salesHeight = Math.max(8, (item.sales / maxSales) * 145);
      const profitHeight = Math.max(8, (item.profit / maxSales) * 145);
      return `
        <article class="bar-card">
          <div class="bar-stack" aria-label="${item.label} 실적">
            <div class="bar sales" style="height: ${salesHeight}px"></div>
            <div class="bar profit" style="height: ${profitHeight}px"></div>
          </div>
          <footer>
            <strong>${item.label}</strong>
            <span>매출 ${won(item.sales)}</span>
            <span>이익 ${won(item.profit)}</span>
          </footer>
        </article>
      `;
    })
    .join("");
}

function renderAccountMix(context) {
  const records = groupByAccount(context.records);
  const totals = context.totals || getTotals(context.records);
  const rows = Object.values(records)
    .sort((a, b) => b.sales - a.sales)
    .slice(0, 5);

  el.accountMix.innerHTML = rows.length
    ? rows
        .map((row) => {
          const mix = rate(row.sales, totals.sales);
          return `
            <article class="rank-item">
              <div class="rank-row"><strong>${escapeHtml(row.account)}</strong><span>${mix.toFixed(1)}%</span></div>
              <div class="meter ${meterClass(mix, 70, 40)}"><span style="width: ${Math.min(mix, 100)}%"></span></div>
              <p>${won(row.sales)} · 마진 ${row.margin.toFixed(1)}%</p>
            </article>
          `;
        })
        .join("")
    : emptyState(`${context.emptyLabel} 실적이 없습니다.`);
}

function renderGapBox(context) {
  const goal = context.goal;
  const totals = context.totals || getTotals(context.records);
  const salesGap = Math.max(goal.sales - totals.sales, 0);
  const profitGap = Math.max(goal.profit - totals.profit, 0);
  const pending = context.actions.filter((action) => action.status !== "완료");
  const pendingSales = sum(pending, "expectedSales");
  const pendingProfit = sum(pending, "expectedProfit");

  el.gapBox.innerHTML = `
    <article class="gap-item">
      <div class="gap-row"><span>부족 매출</span><strong>${won(salesGap)}</strong></div>
      <p>미완료 계획 반영 시 ${won(Math.max(salesGap - pendingSales, 0))} 부족</p>
    </article>
    <article class="gap-item">
      <div class="gap-row"><span>부족 이익</span><strong>${won(profitGap)}</strong></div>
      <p>미완료 계획 반영 시 ${won(Math.max(profitGap - pendingProfit, 0))} 부족</p>
    </article>
  `;
}

function renderDashboardActions(context) {
  const actions = context.actions
    .filter((action) => action.status !== "완료")
    .sort((a, b) => a.due.localeCompare(b.due))
    .slice(0, 3);

  el.dashboardActions.innerHTML = actions.length
    ? actions
        .map(
          (action) => `
          <article class="action-card">
            <span class="pill ${statusClass(action.status)}">${action.status}</span>
            <h3>${escapeHtml(action.title)}</h3>
            <p>${escapeHtml(action.account)} · ${action.due}</p>
            <p>매출 ${won(action.expectedSales)} / 이익 ${won(action.expectedProfit)}</p>
          </article>
        `,
        )
        .join("")
    : emptyState(`${context.emptyLabel} 미완료 실행계획이 없습니다.`);
}

function sortTableRows(rows, table, accessors, fallbackSort) {
  const sorted = [...rows];
  const sort = tableSorts[table];
  if (!sort || !accessors[sort.key]) {
    return fallbackSort ? sorted.sort(fallbackSort) : sorted;
  }

  sorted.sort((a, b) => {
    const result = compareSortValues(accessors[sort.key](a), accessors[sort.key](b), sort.type);
    if (result !== 0) return sort.direction === "asc" ? result : -result;
    return fallbackSort ? fallbackSort(a, b) : 0;
  });
  return sorted;
}

function compareSortValues(a, b, type = "text") {
  if (type === "number") {
    return (Number(a) || 0) - (Number(b) || 0);
  }
  return String(a ?? "").localeCompare(String(b ?? ""), "ko", { numeric: true, sensitivity: "base" });
}

function updateSortHeaders() {
  document.querySelectorAll("th[data-sort-table][data-sort-key]").forEach((header) => {
    const sort = tableSorts[header.dataset.sortTable];
    const active = sort?.key === header.dataset.sortKey;
    header.classList.add("sortable");
    header.classList.toggle("sort-asc", active && sort.direction === "asc");
    header.classList.toggle("sort-desc", active && sort.direction === "desc");
    header.setAttribute("aria-sort", active ? (sort.direction === "asc" ? "ascending" : "descending") : "none");
  });
}

function renderMonthly() {
  fillAccountFilter();
  const selectedAccount = el.accountFilter.value || "전체";
  let records = sortTableRows(getQuarterRecords(), "monthly", {
    month: (record) => record.month,
    account: (record) => record.account,
    channel: (record) => record.channel,
    items: (record) => record.items || "",
    sales: (record) => record.sales,
    profit: (record) => record.sales - record.cost,
    margin: (record) => margin(record.sales, record.sales - record.cost),
  }, (a, b) => a.month.localeCompare(b.month) || a.account.localeCompare(b.account));
  if (selectedAccount !== "전체") {
    records = records.filter((record) => record.account === selectedAccount);
  }
  const totals = getTotals(records);
  el.monthlySummary.textContent = `${records.length}건 · 매출 ${won(totals.sales)} · 이익 ${won(totals.profit)} · 마진 ${totals.margin.toFixed(1)}%`;

  el.monthlyRows.innerHTML = records.length
    ? records
        .map((record) => {
          const profit = record.sales - record.cost;
          return `
            <tr>
              <td>${formatMonth(record.month)}</td>
              <td>${escapeHtml(record.account)}</td>
              <td>${escapeHtml(record.channel)}</td>
              <td>${escapeHtml(record.items || "-")}</td>
              <td class="num">${won(record.sales)}</td>
              <td class="num">${won(profit)}</td>
              <td class="num">${margin(record.sales, profit).toFixed(1)}%</td>
              <td>
                <div class="table-actions">
                  <button class="danger-button" type="button" data-delete-record="${record.id}">삭제</button>
                </div>
              </td>
            </tr>
          `;
        })
        .join("")
    : `<tr><td colspan="8">${emptyState("선택한 조건의 실적이 없습니다.")}</td></tr>`;

  document.querySelectorAll("[data-delete-record]").forEach((button) => {
    button.addEventListener("click", () => deleteRecord(button.dataset.deleteRecord));
  });
}

function fillAccountFilter() {
  const currentValue = el.accountFilter.value || "전체";
  const accounts = [...new Set(getQuarterRecords().map((record) => record.account))].sort();
  el.accountFilter.innerHTML = ["전체", ...accounts].map((account) => `<option value="${escapeAttr(account)}">${escapeHtml(account)}</option>`).join("");
  el.accountFilter.value = accounts.includes(currentValue) ? currentValue : "전체";
}

function deleteRecord(id) {
  const target = state.records.find((record) => record.id === id);
  if (!target) return;
  const sameKey = (month, account) => month === target.month && normalizeAccountKey(account) === normalizeAccountKey(target.account);
  // 같은 월·거래처의 다른 실적이 남아 있으면 품목 데이터는 그 실적에 속하므로 보존
  const otherSameRecord = state.records.some((record) => record.id !== id && sameKey(record.month, record.account));
  const itemMatches = otherSameRecord ? [] : state.itemRecords.filter((item) => sameKey(item.month, item.account));

  const message = itemMatches.length
    ? `이 실적을 삭제하면 연결된 품목 데이터 ${itemMatches.length}건도 함께 삭제됩니다. 삭제할까요?`
    : "이 실적을 삭제할까요?";
  if (!confirm(message)) return;

  state.records = state.records.filter((record) => record.id !== id);
  if (!otherSameRecord) {
    state.itemRecords = state.itemRecords.filter((item) => !sameKey(item.month, item.account));
  }
  saveState();
  renderAll();
}

function renderExcel() {
  const supported = typeof window.XLSX !== "undefined";
  el.excelSupportMessage.classList.toggle("error", !supported);
  el.excelSupportMessage.textContent = supported
    ? "엑셀 파일을 선택하면 거래명세서와 월별 매출이익 파일까지 감안해 데이터 시작 행과 컬럼을 자동 인식합니다. 날짜가 없는 파일은 기준 월을 사용합니다."
    : "엑셀 업로드 기능을 불러오지 못했습니다. lib/xlsx.full.min.js 파일이 있는지 확인해 주세요.";

  el.excelFileInput.disabled = !supported;
  el.analyzeExcelButton.disabled = !supported || excelRows.length === 0;
  el.importExcelResultButton.disabled = !excelAnalysis || !excelAnalysis.rows.length;
  el.excelMappingPanel.hidden = excelRows.length === 0;
  el.excelResultPanel.hidden = !excelAnalysis;

  if (!excelAnalysis) {
    el.excelResultSummary.textContent = excelRows.length
      ? `${excelRows.length}건을 읽었습니다. 컬럼 매칭을 확인한 뒤 분석 실행을 눌러주세요.`
      : "엑셀 파일을 선택하면 결과가 표시됩니다.";
    el.excelKpis.innerHTML = "";
    el.excelAnalysisNote.innerHTML = "";
    el.excelResultRows.innerHTML = `<tr><td colspan="8">${emptyState("분석 결과가 아직 없습니다.")}</td></tr>`;
  }
}

async function handleExcelFile() {
  const file = el.excelFileInput.files?.[0];
  if (!file) return;
  if (typeof window.XLSX === "undefined") {
    showToast("엑셀 업로드 기능을 불러오지 못했습니다.", "error");
    return;
  }

  try {
    const buffer = await file.arrayBuffer();
    const workbook = window.XLSX.read(buffer, { type: "array", cellDates: true });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const matrix = window.XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
    const parsed = buildExcelRows(matrix, file.name);
    excelRows = parsed.rows;
    excelMeta = parsed.meta;
    excelAnalysis = null;

    if (!excelRows.length) {
      showToast("읽을 수 있는 거래내역이 없습니다. 날짜가 들어 있는 데이터 행을 찾지 못했습니다.", "error");
      renderExcel();
      return;
    }

    if (!el.excelAccountInput.value.trim() && parsed.meta.accountName) {
      el.excelAccountInput.value = parsed.meta.accountName;
    }
    if (parsed.meta.defaultMonth) {
      el.excelMonthInput.value = parsed.meta.defaultMonth;
    }
    setupExcelMapping(parsed.headers, parsed.detected);
    renderExcel();
    showToast(`${sheetName} 시트에서 ${excelRows.length}건을 읽었습니다.`, "success");
  } catch (error) {
    excelRows = [];
    excelMeta = null;
    excelAnalysis = null;
    renderExcel();
    showToast(`엑셀 파일을 읽지 못했습니다: ${error.message}`, "error");
  }
}

function buildExcelRows(matrix, fileName = "") {
  const cleanMatrix = matrix.filter((row) => Array.isArray(row) && row.some((cell) => getCellText(cell)));
  const headerRowIndex = findHeaderRow(cleanMatrix);
  const datedDataStartIndex = findDataStartRow(cleanMatrix);
  const useHeaderMode = headerRowIndex !== -1 && (datedDataStartIndex === -1 || headerRowIndex < datedDataStartIndex);
  const dataStartIndex = useHeaderMode ? getDataStartAfterHeader(cleanMatrix, headerRowIndex) : datedDataStartIndex;
  if (dataStartIndex === -1) {
    return { headers: [], rows: [], detected: {}, meta: { accountName: extractAccountName(cleanMatrix), defaultMonth: extractDefaultMonth(cleanMatrix, fileName), dataStartIndex: -1 } };
  }

  const maxCols = Math.max(...cleanMatrix.map((row) => row.length));
  const headerRows = useHeaderMode
    ? cleanMatrix.slice(headerRowIndex, dataStartIndex)
    : cleanMatrix.slice(Math.max(0, dataStartIndex - 3), dataStartIndex);
  const headers = Array.from({ length: maxCols }, (_, index) => makeColumnHeader(headerRows, index));
  const rows = cleanMatrix.slice(dataStartIndex).map((row) =>
    headers.reduce((result, header, index) => {
      result[header] = row[index] ?? "";
      return result;
    }, {}),
  );

  return {
    headers,
    rows: rows.filter((row) => Object.values(row).some((value) => getCellText(value))),
    detected: detectExcelColumns(headers, rows),
    meta: {
      accountName: extractAccountName(cleanMatrix),
      defaultMonth: extractDefaultMonth(cleanMatrix, fileName),
      mode: useHeaderMode ? "summary" : "ledger",
      dataStartIndex,
    },
  };
}

function findDataStartRow(matrix) {
  return matrix.findIndex((row) => row.some((cell) => parseMonthValue(cell)) && row.filter(isLikelyAmountCell).length >= 1);
}

function findHeaderRow(matrix) {
  let bestIndex = -1;
  let bestScore = 0;
  matrix.forEach((row, index) => {
    const text = row.map((cell) => normalizeHeader(cell)).join("|");
    let score = 0;
    if (/월\/일|거래일|일자|날짜/.test(text)) score += 2;
    if (/거래처|거래처명|매출처/.test(text)) score += 2;
    if (/상품명|품명|품목|제품명/.test(text)) score += 2;
    if (/매출금액|매출합계액|공급가액|판매금액/.test(text)) score += 3;
    if (/매입금액|매입합계액|매출원가|원가/.test(text)) score += 3;
    if (/매출이익|이익\(%\)|이익률/.test(text)) score += 3;
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });
  return bestScore >= 4 ? bestIndex : -1;
}

function getDataStartAfterHeader(matrix, headerRowIndex) {
  let index = headerRowIndex + 1;
  if (index < matrix.length && !rowLooksLikeData(matrix[index]) && rowHasSubHeaders(matrix[index])) {
    index += 1;
  }
  return index < matrix.length ? index : -1;
}

function rowLooksLikeData(row) {
  return row.filter(isLikelyAmountCell).length >= 3 || row.some((cell) => parseMonthValue(cell));
}

function rowHasSubHeaders(row) {
  const textCells = row.map(getCellText).filter(Boolean);
  if (!textCells.length) return false;
  return textCells.length <= 8 && row.filter(isLikelyAmountCell).length < 2;
}

function makeColumnHeader(headerRows, index) {
  const parts = headerRows.map((row) => getCellText(row[index])).filter(Boolean);
  const text = [...new Set(parts)].join(" ");
  return `${columnLabel(index)}: ${text || `열${index + 1}`}`;
}

function columnLabel(index) {
  let label = "";
  let number = index + 1;
  while (number > 0) {
    const remainder = (number - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    number = Math.floor((number - 1) / 26);
  }
  return label;
}

function extractAccountName(matrix) {
  for (const row of matrix.slice(0, 8)) {
    for (const cell of row) {
      const text = getCellText(cell);
      const match = text.match(/거래(?:처|명)\s*[:：]\s*(.+)$/);
      if (match) return match[1].trim();
    }
  }
  return "";
}

function extractDefaultMonth(matrix, fileName = "") {
  const candidates = [fileName];
  matrix.slice(0, 8).forEach((row) => row.forEach((cell) => candidates.push(getCellText(cell))));
  for (const candidate of candidates) {
    const month = parseMonthValue(candidate);
    if (month) return month;
  }
  return "";
}

function detectExcelColumns(headers, rows) {
  const date = detectColumnByValues(headers, rows, "date") || detectColumn(headers, ["거래일자", "매출일자", "일자", "날짜", "출고일", "년월", "월/일"]);
  return {
    date,
    account: detectColumn(headers, ["거래처", "매출처", "납품처", "고객", "업체명", "회사명"]),
    item: detectItemColumn(headers),
    sales: detectSalesColumn(headers, rows, date),
    cost: detectCostColumn(headers, rows),
    profit: detectColumn(headers, ["매출이익", "이익", "마진", "매출총이익", "손익"]),
  };
}

function detectItemColumn(headers) {
  return (
    headers.find((header) => /상품명|품명|품목명|제품명/.test(normalizeHeader(header)) && !/코드/.test(normalizeHeader(header))) ||
    headers.find((header) => /품목|상품|제품|부위|적요|내용/.test(normalizeHeader(header)) && !/코드/.test(normalizeHeader(header))) ||
    ""
  );
}

function detectColumnByValues(headers, rows, type) {
  const sample = rows.slice(0, 80);
  let best = "";
  let bestScore = 0;
  headers.forEach((header) => {
    const score = sample.reduce((count, row) => {
      if (type === "date") return count + (parseMonthValue(row[header]) ? 1 : 0);
      return count;
    }, 0);
    if (score > bestScore) {
      bestScore = score;
      best = header;
    }
  });
  return bestScore ? best : "";
}

function detectSalesColumn(headers, rows, dateHeader) {
  const candidates = rankMoneyColumns(headers, rows, [dateHeader]).filter((item) => {
    const header = normalizeHeader(item.header);
    if (/(이익|마진|손익|grossprofit|profit)/i.test(header)) return false;
    if (/(미수|잔액|누계|번호|코드|box|kg|단가|부가세|vat)/i.test(header)) return false;
    return /(매출|판매|공급가액|금액|총액|합계|공급)/i.test(header);
  });
  candidates.sort((a, b) => salesHeaderScore(b.header) - salesHeaderScore(a.header) || b.count - a.count || b.sum - a.sum);
  return candidates[0]?.header || rankMoneyColumns(headers, rows, [dateHeader])[0]?.header || "";
}

function salesHeaderScore(header) {
  const text = normalizeHeader(header);
  if (/(이익|마진|손익|grossprofit|profit)/i.test(text)) return -1;
  if (/매출금액|매출합계액|매출액/.test(text)) return 5;
  if (/판매금액|판매액/.test(text)) return 4;
  if (/매출.*공급가액|공급가액.*매출/.test(text)) return 3;
  if (/매출/.test(text)) return 2;
  if (/공급가액|금액|합계/.test(text)) return 1;
  return 0;
}

function detectCostColumn(headers, rows) {
  const candidates = rankMoneyColumns(headers, rows).filter((item) => {
    const header = normalizeHeader(item.header);
    return /(원가|매입|원료|재료)/i.test(header) && !/(단가|번호|코드)/i.test(header);
  });
  return candidates[0]?.header || "";
}

function rankMoneyColumns(headers, rows, excluded = []) {
  const sample = rows.slice(0, 120);
  return headers
    .filter((header) => header && !excluded.includes(header))
    .map((header) => {
      const values = sample.map((row) => parseMoney(row[header])).filter((value) => value !== null && value !== 0);
      return {
        header,
        count: values.length,
        sum: values.reduce((total, value) => total + Math.abs(value), 0),
      };
    })
    .filter((item) => item.count > 0)
    .sort((a, b) => b.count - a.count || b.sum - a.sum);
}

function setupExcelMapping(headers, detected = {}) {
  setMappingOptions(el.mapDateSelect, headers, detected.date || detectColumn(headers, ["거래일자", "매출일자", "일자", "날짜", "출고일", "년월", "월/일"]), true);
  setMappingOptions(el.mapAccountSelect, headers, detected.account || detectColumn(headers, ["거래처", "매출처", "납품처", "고객", "업체명", "회사명"]), true);
  setMappingOptions(el.mapItemSelect, headers, detected.item || detectItemColumn(headers), true);
  setMappingOptions(el.mapSalesSelect, headers, detected.sales || detectColumn(headers, ["매출액", "공급가액", "판매금액", "금액", "매출", "총액", "합계"]));
  setMappingOptions(el.mapCostSelect, headers, detected.cost || detectColumn(headers, ["매출원가", "원가", "매입금액", "매입액", "원료비", "재료비", "매입"]), true);
  setMappingOptions(el.mapProfitSelect, headers, detected.profit || detectColumn(headers, ["매출이익", "이익", "마진", "매출총이익", "손익"]), true);
}

function setMappingOptions(select, headers, selected, allowBlank = false) {
  const options = allowBlank ? [`<option value="">없음</option>`] : [];
  options.push(...headers.map((header) => `<option value="${escapeAttr(header)}">${escapeHtml(header)}</option>`));
  select.innerHTML = options.join("");
  select.value = selected || (allowBlank ? "" : headers[0] || "");
}

function detectColumn(headers, candidates) {
  const normalizedCandidates = candidates.map(normalizeHeader);
  return headers.find((header) => normalizedCandidates.some((candidate) => normalizeHeader(header).includes(candidate))) || "";
}

function normalizeHeader(value) {
  return String(value).replace(/\s/g, "").toLowerCase();
}

function analyzeExcelRows() {
  if (!excelRows.length) {
    showToast("먼저 엑셀 파일을 선택해 주세요.", "error");
    return;
  }

  const mapping = {
    date: el.mapDateSelect.value,
    account: el.mapAccountSelect.value,
    item: el.mapItemSelect.value,
    sales: el.mapSalesSelect.value,
    cost: el.mapCostSelect.value,
    profit: el.mapProfitSelect.value,
  };
  const fallbackMonth = el.excelMonthInput.value || excelMeta?.defaultMonth || "";

  if (!mapping.date && !fallbackMonth) {
    showToast("날짜 컬럼이 없는 파일은 기준 월을 입력해야 합니다.", "error");
    return;
  }
  if (!mapping.sales) {
    showToast("매출액 컬럼은 반드시 필요합니다.", "error");
    return;
  }

  const groups = {};
  const itemSales = {};
  const itemDetails = {};
  let skipped = 0;
  let summarySkipped = 0;
  let costDetailCount = 0;
  let costEstimated = 0;
  let costMissing = 0;
  let lastAccount = el.excelAccountInput.value.trim() || excelMeta?.accountName || "";

  excelRows.forEach((row) => {
    if (isExcelSummaryRow(row, mapping)) {
      summarySkipped += 1;
      return;
    }

    const month = mapping.date ? parseMonthValue(row[mapping.date]) || fallbackMonth : fallbackMonth;
    const sales = parseMoney(row[mapping.sales]);
    if (!month || sales === null) {
      skipped += 1;
      return;
    }

    const profitValue = mapping.profit ? parseMoney(row[mapping.profit]) : null;
    let cost = mapping.cost ? parseMoney(row[mapping.cost]) : null;
    if (cost !== null || profitValue !== null) {
      costDetailCount += 1;
    }
    if (cost === null && profitValue !== null) {
      cost = sales - profitValue;
      costEstimated += 1;
    }
    if (cost === null) {
      cost = sales;
      costMissing += 1;
    }

    const rowAccount = getCellText(mapping.account ? row[mapping.account] : "");
    if (rowAccount) lastAccount = rowAccount;
    const account = rowAccount || lastAccount || "미지정 거래처";
    const item = getCellText(mapping.item ? row[mapping.item] : "");
    const key = `${month}__${account}`;
    if (!groups[key]) {
      groups[key] = { month, account, sales: 0, cost: 0, count: 0, items: new Set() };
    }
    groups[key].sales += sales;
    groups[key].cost += cost;
    groups[key].count += 1;
    if (item) {
      groups[key].items.add(item);
      itemSales[item] = (itemSales[item] || 0) + sales;
      const itemKey = `${month}__${normalizeAccountKey(account)}__${normalizeItemKey(item)}`;
      if (!itemDetails[itemKey]) {
        itemDetails[itemKey] = { month, account, item, sales: 0, cost: 0, count: 0 };
      }
      itemDetails[itemKey].sales += sales;
      itemDetails[itemKey].cost += cost;
      itemDetails[itemKey].count += 1;
    }
  });

  const rows = Object.values(groups)
    .map((row) => ({
      ...row,
      profit: row.sales - row.cost,
      margin: margin(row.sales, row.sales - row.cost),
      itemsText: [...row.items].slice(0, 5).join(", "),
    }))
    .sort((a, b) => a.month.localeCompare(b.month) || a.account.localeCompare(b.account));

  excelAnalysis = {
    rows,
    skipped,
    summarySkipped,
    costDetailCount,
    costEstimated,
    costMissing,
    sourceCount: excelRows.length,
    itemSales,
    itemDetails: Object.values(itemDetails).map((row) => ({
      ...row,
      profit: row.sales - row.cost,
      margin: margin(row.sales, row.sales - row.cost),
    })),
    meta: excelMeta,
  };
  renderExcelAnalysis();
}

function isExcelSummaryRow(row, mapping) {
  const accountText = getCellText(mapping.account ? row[mapping.account] : "");
  const itemText = getCellText(mapping.item ? row[mapping.item] : "");
  if (isSummaryLabel(accountText) || isSummaryLabel(itemText)) return true;

  const hasSummaryLabel = Object.values(row).some((value) => isSummaryLabel(getCellText(value)));
  if (!hasSummaryLabel) return false;

  // Summary rows in vendor reports usually carry totals but no product name.
  return !itemText;
}

function isSummaryLabel(value) {
  const text = normalizeHeader(value).replace(/[:：]/g, "");
  return /^(소계|합계|총계|누계|월계|전체합계|총합계|grandtotal|subtotal|total)$/.test(text);
}

function renderExcelAnalysis() {
  if (!excelAnalysis) return;
  const totals = getTotalsFromAggregates(excelAnalysis.rows);
  const bestMonth = [...excelAnalysis.rows].sort((a, b) => b.sales - a.sales)[0];
  const weakMargin = [...excelAnalysis.rows].filter((row) => row.sales > 0).sort((a, b) => a.margin - b.margin)[0];
  const topItem = Object.entries(excelAnalysis.itemSales).sort((a, b) => b[1] - a[1])[0];

  el.excelResultSummary.textContent = `${excelAnalysis.sourceCount}건 중 ${excelAnalysis.rows.length}개 월별 실적으로 집계 · 제외 ${excelAnalysis.skipped}건`;
  el.excelKpis.innerHTML = [
    excelKpi("총 매출", won(totals.sales)),
    excelKpi("총 매출이익", won(totals.profit)),
    excelKpi("평균 마진율", `${totals.margin.toFixed(1)}%`),
    excelKpi("집계 기간", getAnalysisRange(excelAnalysis.rows)),
  ].join("");

  const notes = [];
  if (bestMonth) notes.push(`매출이 가장 큰 월은 ${formatMonth(bestMonth.month)}이며, ${bestMonth.account} 기준 매출 ${won(bestMonth.sales)}입니다.`);
  if (weakMargin) notes.push(`마진율이 가장 낮은 월은 ${formatMonth(weakMargin.month)} ${weakMargin.account}이며, 마진율 ${weakMargin.margin.toFixed(1)}%입니다.`);
  if (topItem) notes.push(`거래내역상 매출 기여가 가장 큰 품목은 ${topItem[0]}이며, 집계 매출은 ${won(topItem[1])}입니다.`);
  if (excelAnalysis.costEstimated) notes.push(`원가 컬럼이 없어 매출이익으로 원가를 역산한 거래가 ${excelAnalysis.costEstimated}건 있습니다.`);
  if (excelAnalysis.costMissing) notes.push(`원가/이익 컬럼이 없어 ${excelAnalysis.costMissing}건은 원가를 매출액과 동일하게 임시 처리했습니다. 이 경우 이익과 마진율은 보수적으로 0으로 표시됩니다.`);
  if (excelAnalysis.summarySkipped) notes.push(`소계/합계 행 ${excelAnalysis.summarySkipped}건은 상세 품목과 중복되어 제외했습니다.`);
  if (excelAnalysis.skipped) notes.push(`날짜 또는 매출액 값이 부족해 제외된 거래가 ${excelAnalysis.skipped}건 있습니다.`);
  el.excelAnalysisNote.innerHTML = notes.map((note) => `<p>${escapeHtml(note)}</p>`).join("");

  el.excelResultRows.innerHTML = excelAnalysis.rows.length
    ? sortTableRows(excelAnalysis.rows, "excelResult", {
        month: (row) => row.month,
        account: (row) => row.account,
        items: (row) => row.itemsText || "",
        count: (row) => row.count,
        sales: (row) => row.sales,
        cost: (row) => row.cost,
        profit: (row) => row.profit,
        margin: (row) => row.margin,
      }, (a, b) => a.month.localeCompare(b.month) || a.account.localeCompare(b.account))
        .map(
          (row) => `
            <tr>
              <td>${formatMonth(row.month)}</td>
              <td>${escapeHtml(row.account)}</td>
              <td>${escapeHtml(row.itemsText || "-")}</td>
              <td class="num">${row.count.toLocaleString("ko-KR")}</td>
              <td class="num">${won(row.sales)}</td>
              <td class="num">${won(row.cost)}</td>
              <td class="num">${won(row.profit)}</td>
              <td class="num">${row.margin.toFixed(1)}%</td>
            </tr>
          `,
        )
        .join("")
    : `<tr><td colspan="8">${emptyState("집계할 수 있는 거래내역이 없습니다.")}</td></tr>`;

  renderExcel();
}

function excelKpi(label, value) {
  return `<article class="excel-kpi"><span>${label}</span><strong>${escapeHtml(value)}</strong></article>`;
}

function getTotalsFromAggregates(rows) {
  const sales = sum(rows, "sales");
  const cost = sum(rows, "cost");
  const profit = sales - cost;
  return { sales, cost, profit, margin: margin(sales, profit) };
}

function getAnalysisRange(rows) {
  if (!rows.length) return "-";
  const months = rows.map((row) => row.month).sort();
  return `${formatMonth(months[0])} ~ ${formatMonth(months[months.length - 1])}`;
}

function importExcelAnalysis() {
  if (!excelAnalysis || !excelAnalysis.rows.length) {
    showToast("반영할 분석 결과가 없습니다.", "error");
    return;
  }

  const overwrite = el.excelOverwriteInput.checked;
  const hasCostDetail = excelAnalysis.costDetailCount > 0;
  let updated = 0;
  let added = 0;
  let merged = 0;
  let itemUpdated = 0;
  let itemAdded = 0;

  excelAnalysis.rows.forEach((row) => {
    row.account = ensureAccountFromImport(row.account); // 마스터 정식명으로 치환(미매칭 시 자동 등록)
    const matches = findImportRecordMatches(row);
    const existing = pickImportRecordMatch(matches, row);
    const payload = {
      month: row.month,
      account: row.account,
      channel: "대기업납품",
      sales: Math.round(row.sales),
      cost: getImportCost(row, existing, hasCostDetail),
      items: row.itemsText,
    };

    if (existing && (overwrite || hasCostDetail)) {
      Object.assign(existing, payload);
      const duplicateRecords = matches.filter((record) => record !== existing);
      if (duplicateRecords.length) {
        const duplicateRecordSet = new Set(duplicateRecords);
        state.records = state.records.filter((record) => !duplicateRecordSet.has(record));
        merged += duplicateRecords.length;
      }
      updated += 1;
    } else {
      state.records.push({ id: cryptoId(), ...payload });
      added += 1;
    }
  });

  if (hasCostDetail && excelAnalysis.itemDetails?.length) {
    const itemResult = importExcelItemDetails(excelAnalysis.itemDetails);
    itemUpdated = itemResult.updated;
    itemAdded = itemResult.added;
    reconcileRecordsWithItemDetails();
  }

  const firstMonth = excelAnalysis.rows[0].month;
  state.activeYear = Number(firstMonth.slice(0, 4));
  state.activeQuarter = quarterFromMonth(firstMonth);
  fillYearOptions();
  el.yearSelect.value = state.activeYear;
  el.quarterSelect.value = state.activeQuarter;
  saveState();
  renderAll();
  const mergeText = merged ? `, 중복정리 ${merged}건` : "";
  const itemText = itemAdded || itemUpdated ? `, 품목 추가 ${itemAdded}건, 품목 수정 ${itemUpdated}건` : "";
  showToast(`월별 실적에 반영했습니다. 추가 ${added}건, 수정 ${updated}건${mergeText}${itemText}`, "success");
}

function importExcelItemDetails(details) {
  let added = 0;
  let updated = 0;

  details.forEach((detail) => {
    if (!detail.item || !detail.month || !detail.account) return;
    detail.account = ensureAccountFromImport(detail.account); // 마스터 정식명으로 치환
    const matches = findItemRecordMatches(detail);
    const existing = matches[0] || null;
    const payload = {
      month: detail.month,
      account: detail.account,
      item: detail.item,
      sales: Math.round(detail.sales),
      cost: Math.round(detail.cost),
    };

    if (existing) {
      Object.assign(existing, payload);
      const duplicateRecords = matches.filter((record) => record !== existing);
      if (duplicateRecords.length) {
        const duplicateRecordSet = new Set(duplicateRecords);
        state.itemRecords = state.itemRecords.filter((record) => !duplicateRecordSet.has(record));
      }
      updated += 1;
    } else {
      state.itemRecords.push({ id: cryptoId(), ...payload });
      added += 1;
    }
  });

  return { added, updated };
}

function findItemRecordMatches(detail) {
  const accountKey = normalizeAccountKey(detail.account);
  const itemKey = normalizeItemKey(detail.item);
  return state.itemRecords.filter((record) => {
    if (record.month !== detail.month) return false;
    if (!accountsLookSame(normalizeAccountKey(record.account), accountKey)) return false;
    return normalizeItemKey(record.item) === itemKey;
  });
}

function findImportRecordMatches(row) {
  const targetKey = normalizeAccountKey(row.account);
  return state.records.filter((record) => {
    if (record.month !== row.month) return false;
    if (record.account === row.account) return true;
    return accountsLookSame(normalizeAccountKey(record.account), targetKey);
  });
}

function pickImportRecordMatch(matches, row) {
  if (!matches.length) return null;
  return matches.find((record) => record.account === row.account) || matches[0];
}

function getImportCost(row, existing, hasCostDetail) {
  if (hasCostDetail || !existing || !recordHasCostDetail(existing)) {
    return Math.round(row.cost);
  }
  return Math.round(existing.cost);
}

function recordHasCostDetail(record) {
  const sales = Number(record.sales);
  const cost = Number(record.cost);
  return Number.isFinite(sales) && Number.isFinite(cost) && Math.round(sales) !== Math.round(cost);
}

function accountsLookSame(left, right) {
  if (!left || !right) return false;
  if (left === right) return true;
  if (left.length < 4 || right.length < 4) return false;
  return left.includes(right) || right.includes(left);
}

function normalizeAccountKey(value) {
  return getCellText(value)
    .toLowerCase()
    .replace(/주식회사|유한회사|\(주\)|㈜|\(유\)|（주）/g, "")
    .replace(/\(?(20\d{2})\)?/g, "")
    .replace(/[^0-9a-z가-힣]/g, "");
}

function normalizeItemKey(value) {
  return getCellText(value)
    .toLowerCase()
    .replace(/[^0-9a-z가-힣]/g, "");
}

function parseMonthValue(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}`;
  }
  if (typeof value === "number" && value >= 30000 && value <= 60000) {
    const date = excelSerialToDate(value);
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
  }

  const text = String(value || "").trim();
  if (!text) return "";
  if (/^-?\(?\d{1,3}(,\d{3})+(\.\d+)?\)?$/.test(text)) return "";
  const compact = text.replace(/\D/g, "");
  if (compact.length >= 6) {
    const year = Number(compact.slice(0, 4));
    const month = Number(compact.slice(4, 6));
    if (year >= 2000 && year <= 2100 && month >= 1 && month <= 12) return `${year}-${String(month).padStart(2, "0")}`;
  }

  const yearMonth = text.match(/(20\d{2})\D+(\d{1,2})/);
  if (yearMonth) {
    const year = Number(yearMonth[1]);
    const month = Number(yearMonth[2]);
    if (year >= 2000 && year <= 2100 && month >= 1 && month <= 12) return `${yearMonth[1]}-${String(month).padStart(2, "0")}`;
  }

  const monthOnly = text.match(/(\d{1,2})\s*월/);
  if (monthOnly) {
    const month = Number(monthOnly[1]);
    if (month >= 1 && month <= 12) return `${state.activeYear}-${String(month).padStart(2, "0")}`;
  }
  return "";
}

function excelSerialToDate(serial) {
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  return new Date(utcValue * 1000);
}

function parseMoney(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const text = String(value).trim();
  if (!text) return null;
  const negative = /^\(.*\)$/.test(text) || text.includes("△");
  const cleaned = text.replace(/[^\d.-]/g, "");
  if (!cleaned || cleaned === "-" || cleaned === ".") return null;
  const number = Number(cleaned);
  if (!Number.isFinite(number)) return null;
  return negative ? -Math.abs(number) : number;
}

function isLikelyAmountCell(value) {
  if (typeof value === "number" && Number.isFinite(value)) return true;
  const text = String(value || "").trim();
  if (!text) return false;
  const compact = text.replace(/\D/g, "");
  if ((/[\/.~년월일-]/.test(text) || text.includes("~")) && compact.length >= 6) return false;
  const numericShape = text.replace(/[,\s₩원]/g, "");
  if (!/^\(?-?\d+(\.\d+)?\)?$/.test(numericShape)) return false;
  return parseMoney(text) !== null;
}

function getCellText(value) {
  return String(value || "").trim();
}

function renderGoals() {
  const goal = getGoal();
  const target = getGoalTargets(goal);
  const totals = getTotals();
  el.goalSalesInput.value = target.sales;
  el.goalProfitInput.value = target.profit;
  el.goalMarginInput.value = target.margin.toFixed(1);

  const progressItems = [
    ["매출 달성률", totals.sales, target.sales, rate(totals.sales, target.sales)],
    ["이익 달성률", totals.profit, target.profit, rate(totals.profit, target.profit)],
    ["마진율 차이", totals.margin, target.margin, target.margin > 0 ? (totals.margin / target.margin) * 100 : 0],
  ];

  el.goalProgress.innerHTML = progressItems
    .map(([label, actual, target, value]) => {
      const isMargin = label.includes("마진");
      const actualText = isMargin ? `${Number(actual).toFixed(1)}%` : won(actual);
      const targetText = isMargin ? `${Number(target).toFixed(1)}%` : won(target);
      return `
        <article class="progress-item">
          <div class="progress-row"><strong>${label}</strong><span>${value.toFixed(1)}%</span></div>
          <div class="meter ${meterClass(value, 100, 70)}"><span style="width: ${Math.min(value, 100)}%"></span></div>
          <p>${actualText} / ${targetText}</p>
        </article>
      `;
    })
    .join("");

  renderAllocations(goal);
  renderManagerGoalSummary(goal);
  renderExtraGoals(goal);
}

function renderAllocations(goal) {
  const accountData = groupByAccount(getQuarterRecords());
  const editingId = el.allocationEditIdInput.value;
  const allocations = sortTableRows(goal.allocations, "allocation", {
    account: (allocation) => allocation.account,
    sales: (allocation) => allocation.sales,
    actualSales: (allocation) => (accountData[allocation.account] || { sales: 0 }).sales,
    salesRate: (allocation) => rate((accountData[allocation.account] || { sales: 0 }).sales, allocation.sales),
    profit: (allocation) => allocation.profit,
    actualProfit: (allocation) => (accountData[allocation.account] || { profit: 0 }).profit,
  }, (a, b) => a.account.localeCompare(b.account, "ko"));

  el.allocationRows.innerHTML = goal.allocations.length
    ? allocations
        .map((allocation) => {
          const actual = accountData[allocation.account] || { sales: 0, profit: 0 };
          return `
            <tr>
              <td>${escapeHtml(allocation.account)}</td>
              <td class="num">${won(allocation.sales)}</td>
              <td class="num">${won(actual.sales)}</td>
              <td class="num">${rate(actual.sales, allocation.sales).toFixed(1)}%</td>
              <td class="num">${won(allocation.profit)}</td>
              <td class="num">${won(actual.profit)}</td>
              <td>
                <div class="table-actions">
                  <button class="secondary-button" type="button" data-edit-allocation="${allocation.id}">${editingId === allocation.id ? "수정중" : "수정"}</button>
                  <button class="danger-button" type="button" data-delete-allocation="${allocation.id}">삭제</button>
                </div>
              </td>
            </tr>
          `;
        })
        .join("")
    : `<tr><td colspan="7">${emptyState("거래처별 목표 배분이 없습니다.")}</td></tr>`;

  document.querySelectorAll("[data-delete-allocation]").forEach((button) => {
    button.addEventListener("click", () => deleteAllocation(button.dataset.deleteAllocation));
  });
  document.querySelectorAll("[data-edit-allocation]").forEach((button) => {
    button.addEventListener("click", () => editAllocation(button.dataset.editAllocation));
  });
}

function renderManagerGoalSummary(goal) {
  if (!el.managerGoalRows) return;
  const accountData = groupByAccount(getQuarterRecords());
  const managerMap = {};

  (goal.allocations || []).forEach((allocation) => {
    const owner = getAccountOwner(allocation.account);
    if (!managerMap[owner]) managerMap[owner] = createManagerBucket(owner);
    managerMap[owner].salesTarget += Number(allocation.sales) || 0;
    managerMap[owner].profitTarget += Number(allocation.profit) || 0;
  });

  Object.entries(accountData).forEach(([account, actual]) => {
    const owner = getAccountOwner(account);
    if (!managerMap[owner]) managerMap[owner] = createManagerBucket(owner);
    managerMap[owner].sales += Number(actual.sales) || 0;
    managerMap[owner].profit += Number(actual.profit) || 0;
    managerMap[owner].margin = margin(managerMap[owner].sales, managerMap[owner].profit);
  });

  const rows = sortTableRows(Object.values(managerMap), "managerGoal", {
    owner: (row) => row.owner,
    salesTarget: (row) => row.salesTarget,
    actualSales: (row) => row.sales,
    salesRate: (row) => rate(row.sales, row.salesTarget),
    profitTarget: (row) => row.profitTarget,
    actualProfit: (row) => row.profit,
    profitRate: (row) => rate(row.profit, row.profitTarget),
  }, (a, b) => b.salesTarget - a.salesTarget || b.sales - a.sales || a.owner.localeCompare(b.owner, "ko"));

  el.managerGoalRows.innerHTML = rows.length
    ? rows
        .map(
          (row) => `
            <tr>
              <td>${escapeHtml(row.owner)}</td>
              <td class="num">${won(row.salesTarget)}</td>
              <td class="num">${won(row.sales)}</td>
              <td class="num">${rate(row.sales, row.salesTarget).toFixed(1)}%</td>
              <td class="num">${won(row.profitTarget)}</td>
              <td class="num">${won(row.profit)}</td>
              <td class="num">${rate(row.profit, row.profitTarget).toFixed(1)}%</td>
            </tr>
          `,
        )
        .join("")
    : `<tr><td colspan="7">${emptyState("담당자별 목표 또는 실적이 없습니다.")}</td></tr>`;
}

function renderExtraGoals(goal) {
  if (!el.extraGoalRows) return;
  const editingId = el.extraGoalEditIdInput.value;
  const rows = sortTableRows(goal.extraGoals || [], "extraGoals", {
    category: (row) => row.category,
    title: (row) => row.title,
    target: (row) => row.target,
    actual: (row) => row.actual,
    rate: (row) => rate(row.actual, row.target),
    unit: (row) => row.unit,
    owner: (row) => row.owner,
    plan: (row) => row.plan,
    status: (row) => row.status,
  }, (a, b) => a.category.localeCompare(b.category, "ko") || a.title.localeCompare(b.title, "ko"));

  el.extraGoalRows.innerHTML = rows.length
    ? rows
        .map(
          (row) => `
            <tr>
              <td>${escapeHtml(row.category)}</td>
              <td><strong>${escapeHtml(row.title)}</strong></td>
              <td class="num">${formatNumber(row.target)}</td>
              <td class="num">${formatNumber(row.actual)}</td>
              <td class="num">${rate(row.actual, row.target).toFixed(1)}%</td>
              <td>${escapeHtml(row.unit || "-")}</td>
              <td>${escapeHtml(row.owner || "-")}</td>
              <td>${escapeHtml(row.plan || "-")}</td>
              <td>
                <select class="status-select" data-extra-goal-status="${row.id}">
                  ${["예정", "진행중", "완료", "보류"].map((status) => `<option value="${status}" ${status === row.status ? "selected" : ""}>${status}</option>`).join("")}
                </select>
              </td>
              <td>
                <div class="table-actions">
                  <button class="secondary-button" type="button" data-edit-extra-goal="${row.id}">${editingId === row.id ? "수정중" : "수정"}</button>
                  <button class="danger-button" type="button" data-delete-extra-goal="${row.id}">삭제</button>
                </div>
              </td>
            </tr>
          `,
        )
        .join("")
    : `<tr><td colspan="10">${emptyState("등록된 부가 목표가 없습니다.")}</td></tr>`;

  document.querySelectorAll("[data-extra-goal-status]").forEach((select) => {
    select.addEventListener("change", () => updateExtraGoalStatus(select.dataset.extraGoalStatus, select.value));
  });
  document.querySelectorAll("[data-edit-extra-goal]").forEach((button) => {
    button.addEventListener("click", () => editExtraGoal(button.dataset.editExtraGoal));
  });
  document.querySelectorAll("[data-delete-extra-goal]").forEach((button) => {
    button.addEventListener("click", () => deleteExtraGoal(button.dataset.deleteExtraGoal));
  });
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString("ko-KR", { maximumFractionDigits: 2 });
}

function editExtraGoal(id) {
  const goal = getGoal();
  const item = (goal.extraGoals || []).find((row) => row.id === id);
  if (!item) return;
  el.extraGoalEditIdInput.value = item.id;
  el.extraGoalCategoryInput.value = item.category;
  el.extraGoalTitleInput.value = item.title;
  el.extraGoalTargetInput.value = item.target;
  el.extraGoalActualInput.value = item.actual;
  el.extraGoalUnitInput.value = item.unit;
  el.extraGoalOwnerInput.value = item.owner;
  el.extraGoalPlanInput.value = item.plan;
  el.extraGoalStatusInput.value = item.status;
  el.extraGoalSubmitButton.textContent = "수정 저장";
  el.extraGoalCancelButton.hidden = false;
  renderExtraGoals(goal);
  el.extraGoalTitleInput.focus();
}

function resetExtraGoalForm() {
  if (!el.extraGoalForm) return;
  el.extraGoalForm.reset();
  el.extraGoalEditIdInput.value = "";
  el.extraGoalActualInput.value = "0";
  el.extraGoalSubmitButton.textContent = "목표 추가";
  el.extraGoalCancelButton.hidden = true;
}

function updateExtraGoalStatus(id, status) {
  const item = (getGoal().extraGoals || []).find((row) => row.id === id);
  if (!item) return;
  item.status = status;
  saveState();
  renderAll();
}

function deleteExtraGoal(id) {
  if (!confirm("이 부가 목표를 삭제할까요?")) return;
  const goal = getGoal();
  goal.extraGoals = (goal.extraGoals || []).filter((row) => row.id !== id);
  if (el.extraGoalEditIdInput.value === id) resetExtraGoalForm();
  saveState();
  renderAll();
}

function editAllocation(id) {
  const goal = getGoal();
  const allocation = goal.allocations.find((item) => item.id === id);
  if (!allocation) return;
  el.allocationEditIdInput.value = allocation.id;
  el.allocationAccountInput.value = allocation.account;
  el.allocationSalesInput.value = allocation.sales;
  el.allocationProfitInput.value = allocation.profit;
  el.allocationSubmitButton.textContent = "수정 저장";
  el.allocationCancelButton.hidden = false;
  renderAllocations(goal);
  el.allocationAccountInput.focus();
}

function resetAllocationForm() {
  el.allocationForm.reset();
  el.allocationEditIdInput.value = "";
  el.allocationSubmitButton.textContent = "배분 추가";
  el.allocationCancelButton.hidden = true;
}

function deleteAllocation(id) {
  if (!confirm("이 목표 배분을 삭제할까요?")) return;
  const goal = getGoal();
  goal.allocations = goal.allocations.filter((allocation) => allocation.id !== id);
  if (el.allocationEditIdInput.value === id) resetAllocationForm();
  syncGoalTargets(goal);
  saveState();
  renderAll();
}

function renderActions() {
  const selectedStatus = el.statusFilter.value || "전체";
  let actions = sortTableRows(getQuarterActions(), "actions", {
    title: (action) => action.title,
    account: (action) => action.account,
    type: (action) => action.type,
    owner: (action) => action.owner,
    due: (action) => action.due,
    expectedSales: (action) => action.expectedSales,
    expectedProfit: (action) => action.expectedProfit,
    status: (action) => action.status,
  }, (a, b) => a.due.localeCompare(b.due));
  if (selectedStatus !== "전체") {
    actions = actions.filter((action) => action.status === selectedStatus);
  }

  const expectedSales = sum(actions, "expectedSales");
  const expectedProfit = sum(actions, "expectedProfit");
  el.actionSummary.textContent = `${actions.length}건 · 예상 매출 ${won(expectedSales)} · 예상 이익 ${won(expectedProfit)}`;

  el.actionRows.innerHTML = actions.length
    ? actions
        .map(
          (action) => `
          <tr>
            <td>${escapeHtml(action.title)}</td>
            <td>${escapeHtml(action.account)}</td>
            <td>${escapeHtml(action.type)}</td>
            <td>${escapeHtml(action.owner)}</td>
            <td>${action.due}</td>
            <td class="num">${won(action.expectedSales)}</td>
            <td class="num">${won(action.expectedProfit)}</td>
            <td>
              <select class="status-select" data-action-status="${action.id}">
                ${["예정", "진행중", "완료", "보류"].map((status) => `<option value="${status}" ${status === action.status ? "selected" : ""}>${status}</option>`).join("")}
              </select>
            </td>
            <td>
              <div class="table-actions">
                <button class="danger-button" type="button" data-delete-action="${action.id}">삭제</button>
              </div>
            </td>
          </tr>
        `,
        )
        .join("")
    : `<tr><td colspan="9">${emptyState("선택한 조건의 실행계획이 없습니다.")}</td></tr>`;

  document.querySelectorAll("[data-action-status]").forEach((select) => {
    select.addEventListener("change", () => updateActionStatus(select.dataset.actionStatus, select.value));
  });

  document.querySelectorAll("[data-delete-action]").forEach((button) => {
    button.addEventListener("click", () => deleteAction(button.dataset.deleteAction));
  });
}

function updateActionStatus(id, status) {
  const action = state.actions.find((item) => item.id === id);
  if (action) {
    action.status = status;
    saveState();
    renderAll();
  }
}

function deleteAction(id) {
  if (!confirm("이 실행계획을 삭제할까요?")) return;
  state.actions = state.actions.filter((action) => action.id !== id);
  saveState();
  renderAll();
}

function renderAccounts() {
  const monthKeys = getSelectedMonthKeys();
  const records = state.records.filter((record) => monthKeys.includes(record.month));
  const label = getSelectedPeriodLabel();
  const totals = getTotals(records);
  const rows = sortTableRows(Object.values(groupByAccount(records)), "accounts", {
    account: (row) => row.account,
    owner: (row) => getAccountOwner(row.account),
    sales: (row) => row.sales,
    profit: (row) => row.profit,
    margin: (row) => row.margin,
    mix: (row) => rate(row.sales, totals.sales),
    strategy: (row) => strategyLabel(row, rate(row.sales, totals.sales)),
    items: (row) => [...row.items].join(", "),
  }, (a, b) => b.sales - a.sales);
  const accounts = rows.map((row) => row.account);
  const selectedAccount = accounts.includes(state.accountChartAccount) ? state.accountChartAccount : "전체";
  state.accountChartAccount = selectedAccount;

  el.accountChartAccountSelect.innerHTML = ["전체", ...accounts].map((account) => `<option value="${escapeAttr(account)}">${escapeHtml(account)}</option>`).join("");
  el.accountChartAccountSelect.value = selectedAccount;
  renderAccountTrendChart(records, selectedAccount, monthKeys, label);
  renderManagerSalesSummary(records, totals);

  el.accountAnalysisSummary.innerHTML = rows.length
    ? `<p>${label} 기준 거래처 ${rows.length.toLocaleString("ko-KR")}곳 · 매출 ${won(totals.sales)} · 이익 ${won(totals.profit)} · 마진 ${totals.margin.toFixed(1)}%</p>`
    : `<p>${label} 기준 분석할 거래처 실적이 없습니다.</p>`;

  el.accountRows.innerHTML = rows.length
    ? rows
        .map((row) => {
          const mix = rate(row.sales, totals.sales);
          return `
            <tr>
              <td>${escapeHtml(row.account)}</td>
              <td>${escapeHtml(getAccountOwner(row.account))}</td>
              <td class="num">${won(row.sales)}</td>
              <td class="num">${won(row.profit)}</td>
              <td class="num">${row.margin.toFixed(1)}%</td>
              <td class="num">${mix.toFixed(1)}%</td>
              <td><span class="pill ${strategyClass(row)}">${strategyLabel(row, mix)}</span></td>
              <td>${escapeHtml([...row.items].join(", ") || "-")}</td>
            </tr>
          `;
        })
        .join("")
    : `<tr><td colspan="8">${emptyState("분석할 실적이 없습니다.")}</td></tr>`;
}

function renderManagerSalesSummary(records, totals) {
  if (!el.managerSalesRows) return;
  const rows = sortTableRows(Object.values(groupRecordsByManager(records)), "managerSales", {
    owner: (row) => row.owner,
    accountCount: (row) => row.accountCount,
    sales: (row) => row.sales,
    profit: (row) => row.profit,
    margin: (row) => row.margin,
    mix: (row) => rate(row.sales, totals.sales),
  }, (a, b) => b.sales - a.sales || a.owner.localeCompare(b.owner, "ko"));

  el.managerSalesRows.innerHTML = rows.length
    ? rows
        .map((row) => {
          const mix = rate(row.sales, totals.sales);
          return `
            <tr>
              <td>${escapeHtml(row.owner)}</td>
              <td class="num">${row.accountCount.toLocaleString("ko-KR")}</td>
              <td class="num">${won(row.sales)}</td>
              <td class="num">${won(row.profit)}</td>
              <td class="num">${row.margin.toFixed(1)}%</td>
              <td class="num">${mix.toFixed(1)}%</td>
            </tr>
          `;
        })
        .join("")
    : `<tr><td colspan="6">${emptyState("담당자별로 집계할 실적이 없습니다.")}</td></tr>`;
}

function renderAccountTrendChart(records, selectedAccount, monthKeys, label) {
  const chartRecords = selectedAccount === "전체" ? records : records.filter((record) => record.account === selectedAccount);
  const items = monthKeys.map((key) => ({
    key,
    label: formatMonth(key),
    ...getTotals(chartRecords.filter((record) => record.month === key)),
  }));
  const totals = getTotals(chartRecords);
  const maxSales = Math.max(...items.map((item) => item.sales), 1);
  const allAccountLabel = state.activeQuarter === ALL_QUARTER ? `${state.activeYear}년 전체 거래처` : `${label} 전체 거래처`;

  el.accountChartSubtitle.textContent =
    selectedAccount === "전체"
      ? `${allAccountLabel} 월별 매출과 매출이익`
      : `${label} ${selectedAccount} 월별 매출과 매출이익 · 매출 ${won(totals.sales)} · 마진 ${totals.margin.toFixed(1)}%`;

  el.accountTrendChart.innerHTML = items
    .map((item) => {
      const salesHeight = Math.max(8, (item.sales / maxSales) * 145);
      const profitHeight = Math.max(8, (item.profit / maxSales) * 145);
      return `
        <article class="bar-card">
          <div class="bar-stack" aria-label="${item.label} 실적">
            <div class="bar sales" style="height: ${salesHeight}px"></div>
            <div class="bar profit" style="height: ${profitHeight}px"></div>
          </div>
          <footer>
            <strong>${item.label}</strong>
            <span>매출 ${won(item.sales)}</span>
            <span>이익 ${won(item.profit)}</span>
          </footer>
        </article>
      `;
    })
    .join("");
}

function renderItemAnalysis() {
  const period = state.itemAnalysisPeriod || "quarter";
  const month = state.itemAnalysisMonth || getDefaultAnalysisMonth();
  const records = getAnalysisRecords(state.itemRecords, period, month);
  const label = getAnalysisLabel(period, month);
  const currentValue = el.itemAccountFilter.value || "전체";
  const accounts = [...new Set(records.map((record) => record.account))].sort();
  el.itemAccountFilter.innerHTML = ["전체", ...accounts].map((account) => `<option value="${escapeAttr(account)}">${escapeHtml(account)}</option>`).join("");
  el.itemAccountFilter.value = accounts.includes(currentValue) ? currentValue : "전체";
  renderAnalysisControls("item", el.itemPeriodButtons, el.itemMonthInput);

  const filtered = el.itemAccountFilter.value === "전체" ? records : records.filter((record) => record.account === el.itemAccountFilter.value);
  const rawRows = Object.values(groupItemRecords(filtered));
  const rawTotals = getItemTotals(rawRows);
  const rows = sortTableRows(rawRows, "itemAnalysis", {
    account: (row) => row.account,
    item: (row) => row.item,
    monthCount: (row) => row.monthCount,
    sales: (row) => row.sales,
    cost: (row) => row.cost,
    profit: (row) => row.sales - row.cost,
    margin: (row) => margin(row.sales, row.sales - row.cost),
    mix: (row) => rate(row.sales, rawTotals.sales),
  }, (a, b) => b.sales - a.sales || a.account.localeCompare(b.account) || a.item.localeCompare(b.item));
  const totals = getItemTotals(rows);
  const weakItem = rows.filter((row) => row.sales > 0).sort((a, b) => a.margin - b.margin)[0];

  el.itemAnalysisKpis.innerHTML = [
    excelKpi("분석 품목", `${rows.length.toLocaleString("ko-KR")}개`),
    excelKpi("품목 매출", won(totals.sales)),
    excelKpi("품목 매출이익", won(totals.profit)),
    excelKpi("평균 마진율", `${totals.margin.toFixed(1)}%`),
  ].join("");

  el.itemAnalysisSummary.innerHTML = rows.length
    ? `<p>${label} 기준 ${el.itemAccountFilter.value === "전체" ? "전체 거래처" : escapeHtml(el.itemAccountFilter.value)} 품목 실적입니다.${weakItem ? ` 최저 마진 품목은 ${escapeHtml(weakItem.item)}(${weakItem.margin.toFixed(1)}%)입니다.` : ""}</p>`
    : `<p>${label} 기준 품목 데이터가 없습니다. 매출이익 파일을 실적 관리에서 반영하면 품목별 마진 데이터가 여기에 쌓입니다.</p>`;

  el.itemAnalysisRows.innerHTML = rows.length
    ? rows
        .map((row) => {
          const profit = row.sales - row.cost;
          const rowMargin = margin(row.sales, profit);
          const mix = rate(row.sales, totals.sales);
          return `
            <tr>
              <td>${escapeHtml(row.account)}</td>
              <td>${escapeHtml(row.item)}</td>
              <td class="num">${row.monthCount.toLocaleString("ko-KR")}</td>
              <td class="num">${won(row.sales)}</td>
              <td class="num">${won(row.cost)}</td>
              <td class="num">${won(profit)}</td>
              <td class="num">${rowMargin.toFixed(1)}%</td>
              <td class="num">${mix.toFixed(1)}%</td>
            </tr>
          `;
        })
        .join("")
    : `<tr><td colspan="8">${emptyState("분석할 품목 데이터가 없습니다.")}</td></tr>`;
}

function groupItemRecords(records) {
  return records.reduce((map, record) => {
    const key = `${normalizeAccountKey(record.account)}__${normalizeItemKey(record.item)}`;
    if (!map[key]) {
      map[key] = { account: record.account, item: record.item, sales: 0, cost: 0, months: new Set(), margin: 0 };
    }
    map[key].sales += Number(record.sales) || 0;
    map[key].cost += Number(record.cost) || 0;
    map[key].months.add(record.month);
    map[key].monthCount = map[key].months.size;
    map[key].margin = margin(map[key].sales, map[key].sales - map[key].cost);
    return map;
  }, {});
}

function getItemTotals(rows) {
  const sales = sum(rows, "sales");
  const cost = sum(rows, "cost");
  const profit = sales - cost;
  return { sales, cost, profit, margin: margin(sales, profit) };
}

/* ============ Supabase 서버 저장 ============ */

function loadSupabaseSession() {
  try {
    const saved = JSON.parse(localStorage.getItem(SUPABASE_SESSION_KEY) || "null");
    if (!saved || !saved.access_token) return null;
    return saved;
  } catch {
    return null;
  }
}

function persistSupabaseSession(session) {
  supabaseSession = session;
  if (session) {
    localStorage.setItem(SUPABASE_SESSION_KEY, JSON.stringify(session));
  } else {
    localStorage.removeItem(SUPABASE_SESSION_KEY);
  }
}

async function loginSupabase() {
  const email = el.supabaseEmailInput.value.trim();
  const password = el.supabasePasswordInput.value;
  if (!email || !password) {
    showToast("Supabase 이메일과 비밀번호를 입력해 주세요.", "error");
    return;
  }

  try {
    const response = await supabaseRequest("/auth/v1/token?grant_type=password", {
      method: "POST",
      body: { email, password },
      auth: false,
    });
    persistSupabaseSession({
      access_token: response.access_token,
      refresh_token: response.refresh_token,
      expires_at: response.expires_at,
      user: response.user,
    });
    el.supabasePasswordInput.value = "";
    renderData();
    showToast("Supabase에 로그인했습니다.", "success");
  } catch (error) {
    showToast(`Supabase 로그인 실패: ${error.message}`, "error");
  }
}

async function logoutSupabase() {
  clearTimeout(supabaseSaveTimer);
  persistSupabaseSession(null);
  supabaseLastSyncedAt = "";
  renderData();
  showToast("Supabase 연결을 해제했습니다.", "success");
}

function isSupabaseSessionExpiring() {
  if (!supabaseSession?.expires_at) return false;
  return Date.now() / 1000 > Number(supabaseSession.expires_at) - 60;
}

async function refreshSupabaseSession() {
  if (!supabaseSession?.refresh_token) throw new Error("다시 로그인이 필요합니다.");
  const response = await supabaseRequest("/auth/v1/token?grant_type=refresh_token", {
    method: "POST",
    body: { refresh_token: supabaseSession.refresh_token },
    auth: false,
  });
  persistSupabaseSession({
    access_token: response.access_token,
    refresh_token: response.refresh_token,
    expires_at: response.expires_at,
    user: response.user,
  });
}

function queueSupabaseSave() {
  clearTimeout(supabaseSaveTimer);
  supabaseSaveTimer = setTimeout(() => saveStateToSupabase(), 900);
}

async function saveStateToSupabase({ showMessage = false } = {}) {
  if (!supabaseSession) {
    if (showMessage) showToast("먼저 Supabase에 로그인해 주세요.", "error");
    return;
  }

  try {
    syncAllGoalTargets();
    const currentServerRow = await fetchSupabaseStateRow();
    const serverChanged = Boolean(currentServerRow?.updated_at && currentServerRow.updated_at !== supabaseLastSyncedAt);
    const dataToSave = serverChanged ? mergeStateForSupabase(currentServerRow.data, state) : state;
    if (serverChanged) {
      state = dataToSave;
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    }
    const payload = {
      id: SUPABASE_STATE_ID,
      data: dataToSave,
      updated_at: new Date().toISOString(),
      updated_by: supabaseSession.user?.email || "",
    };
    const rows = await supabaseRequest("/rest/v1/app_state?on_conflict=id", {
      method: "POST",
      body: payload,
      headers: { Prefer: "resolution=merge-duplicates,return=representation" },
    });
    supabaseLastSyncedAt = rows?.[0]?.updated_at || payload.updated_at;
    updateStorageUi();
    if (serverChanged) {
      renderAll();
    } else {
      renderData();
    }
    if (showMessage) showToast(serverChanged ? "서버의 최신 변경과 합쳐서 저장했습니다." : "현재 데이터를 Supabase 서버에 저장했습니다.", "success");
  } catch (error) {
    showToast(`Supabase 저장 실패: ${error.message}`, "error");
  }
}

async function loadStateFromSupabase({ showMessage = false } = {}) {
  if (!supabaseSession) {
    showToast("먼저 Supabase에 로그인해 주세요.", "error");
    return;
  }

  try {
    const serverRow = await fetchSupabaseStateRow();
    if (!serverRow) {
      showToast("Supabase 서버에 저장된 데이터가 없습니다. 먼저 현재 데이터를 서버에 저장해 주세요.", "error");
      return;
    }
    state = prepareState(serverRow.data);
    supabaseLastSyncedAt = serverRow.updated_at || "";
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    fillYearOptions();
    el.yearSelect.value = state.activeYear;
    el.quarterSelect.value = state.activeQuarter;
    el.monthInput.value = getDefaultMonthInputValue();
    renderAll();
    if (showMessage) showToast("Supabase 서버에서 데이터를 불러왔습니다.", "success");
  } catch (error) {
    showToast(`Supabase 불러오기 실패: ${error.message}`, "error");
  }
}

async function fetchSupabaseStateRow() {
  const rows = await supabaseRequest(`/rest/v1/app_state?id=eq.${encodeURIComponent(SUPABASE_STATE_ID)}&select=id,data,updated_at,updated_by&limit=1`, {
    method: "GET",
  });
  return rows[0] || null;
}

function mergeStateForSupabase(serverData, localData) {
  const server = prepareState(serverData);
  const local = prepareState(localData);
  const merged = {
    ...server,
    ...local,
    records: mergeById(server.records, local.records),
    itemRecords: mergeById(server.itemRecords, local.itemRecords),
    actions: mergeById(server.actions, local.actions),
    accounts: mergeAccountsForSync(server.accounts, local.accounts),
    goals: mergeGoalsForSync(server.goals, local.goals),
  };
  return prepareState(merged);
}

function mergeById(serverItems = [], localItems = []) {
  const map = new Map();
  serverItems.forEach((item) => map.set(item.id || cryptoId(), item));
  localItems.forEach((item) => map.set(item.id || cryptoId(), item));
  return [...map.values()];
}

function mergeAccountsForSync(serverAccounts = [], localAccounts = []) {
  const map = new Map();
  [...serverAccounts, ...localAccounts].forEach((account) => {
    const key = account.id || normalizeAccountKey(account.name);
    const existing = map.get(key);
    if (!existing) {
      map.set(key, { ...account, aliases: [...(account.aliases || [])] });
      return;
    }
    const aliases = new Set([...(existing.aliases || []), ...(account.aliases || [])]);
    map.set(key, { ...existing, ...account, owner: account.owner || existing.owner || "", aliases: [...aliases] });
  });
  return [...map.values()];
}

function mergeGoalsForSync(serverGoals = {}, localGoals = {}) {
  const result = { ...serverGoals };
  Object.entries(localGoals).forEach(([key, localGoal]) => {
    const serverGoal = result[key];
    if (!serverGoal) {
      result[key] = localGoal;
      return;
    }
    result[key] = {
      ...serverGoal,
      ...localGoal,
      allocations: mergeById(serverGoal.allocations || [], localGoal.allocations || []),
      extraGoals: mergeById(serverGoal.extraGoals || [], localGoal.extraGoals || []),
    };
  });
  Object.values(result).forEach((goal) => syncGoalTargets(goal));
  return result;
}

async function supabaseRequest(path, options = {}) {
  const headers = {
    apikey: SUPABASE_PUBLISHABLE_KEY,
    ...options.headers,
  };
  if (options.body !== undefined) {
    headers["Content-Type"] = "application/json";
  }
  if (options.auth !== false) {
    if (!supabaseSession?.access_token) throw new Error("로그인이 필요합니다.");
    if (isSupabaseSessionExpiring()) {
      await refreshSupabaseSession();
    }
    headers.Authorization = `Bearer ${supabaseSession.access_token}`;
  }

  const response = await fetch(`${SUPABASE_URL}${path}`, {
    method: options.method || "GET",
    headers,
    body: options.body !== undefined ? JSON.stringify(options.body) : undefined,
  });
  const text = await response.text();
  let payload = null;
  try {
    payload = text ? JSON.parse(text) : null;
  } catch {
    payload = text;
  }
  if (!response.ok) {
    throw new Error(payload?.message || payload?.error_description || payload?.error || response.statusText);
  }
  return payload;
}

/* ============ 거래처 마스터 / 매칭 ============ */

// 들어온 거래처명을 마스터의 정식명으로 해석. {name, account, matched}
function resolveAccountName(raw) {
  const text = getCellText(raw);
  const key = normalizeAccountKey(text);
  if (!key) return { name: text, account: null, matched: false };

  // 1순위: 정규화한 정식명/별칭 완전 일치
  for (const account of state.accounts) {
    if (normalizeAccountKey(account.name) === key) return { name: account.name, account, matched: true };
    if ((account.aliases || []).some((alias) => normalizeAccountKey(alias) === key)) return { name: account.name, account, matched: true };
  }
  // 2순위: 부분 포함(4자 이상) 유연 매칭
  for (const account of state.accounts) {
    if (accountsLookSame(normalizeAccountKey(account.name), key)) return { name: account.name, account, matched: true };
    if ((account.aliases || []).some((alias) => accountsLookSame(normalizeAccountKey(alias), key))) return { name: account.name, account, matched: true };
  }
  return { name: text, account: null, matched: false };
}

// 임포트용: 해석 실패 시 신규 거래처 자동 등록 후 정식명 반환
function ensureAccountFromImport(raw) {
  const text = getCellText(raw);
  if (!text) return text;
  const resolved = resolveAccountName(text);
  if (resolved.matched) return resolved.name;
  state.accounts.push({ id: cryptoId(), name: text, owner: "", aliases: [] });
  return text;
}

function bindAccountMaster() {
  el.accountForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const name = el.accountNameInput.value.trim();
    if (!name) return;
    const owner = el.accountOwnerInput.value.trim();
    const aliases = el.accountAliasInput.value
      .split(",")
      .map((alias) => alias.trim())
      .filter(Boolean);
    const existing = state.accounts.find((account) => normalizeAccountKey(account.name) === normalizeAccountKey(name));
    if (existing) {
      if (owner) existing.owner = owner;
      aliases.forEach((alias) => addAliasToAccount(existing, alias));
      showToast("이미 있는 거래처라 담당자와 별칭을 반영했습니다.", "success");
    } else {
      const account = { id: cryptoId(), name, owner, aliases: [] };
      aliases.forEach((alias) => addAliasToAccount(account, alias));
      state.accounts.push(account);
      showToast("거래처를 추가했습니다.", "success");
    }
    el.accountForm.reset();
    saveState();
    renderAll();
  });

  el.recanonicalizeButton.addEventListener("click", recanonicalizeAll);
  if (el.cleanOrphanItemsButton) el.cleanOrphanItemsButton.addEventListener("click", cleanOrphanItemRecords);
}

// 대응하는 실적(records)이 없는 품목 데이터(itemRecords)를 제거
function cleanOrphanItemRecords() {
  const recordKeys = new Set(state.records.map((r) => `${r.month}__${normalizeAccountKey(r.account)}`));
  const orphans = state.itemRecords.filter((item) => !recordKeys.has(`${item.month}__${normalizeAccountKey(item.account)}`));
  if (!orphans.length) {
    showToast("정리할 품목 데이터가 없습니다. (모든 품목이 실적과 연결돼 있습니다)", "success");
    return;
  }
  if (!confirm(`실적이 없는 품목 데이터 ${orphans.length}건을 삭제할까요?`)) return;
  const orphanSet = new Set(orphans);
  state.itemRecords = state.itemRecords.filter((item) => !orphanSet.has(item));
  saveState();
  renderAll();
  showToast(`품목 데이터 ${orphans.length}건을 정리했습니다.`, "success");
}

function addAliasToAccount(account, alias) {
  const text = getCellText(alias);
  if (!text) return;
  account.aliases = account.aliases || [];
  if (text === getCellText(account.name)) return;
  if (account.aliases.some((a) => getCellText(a) === text)) return;
  account.aliases.push(text);
}

function countAccountRecords(account) {
  const keys = new Set([normalizeAccountKey(account.name), ...(account.aliases || []).map(normalizeAccountKey)]);
  const matches = (name) => keys.has(normalizeAccountKey(name));
  const recordCount = state.records.filter((r) => matches(r.account)).length;
  const itemCount = state.itemRecords.filter((r) => matches(r.account)).length;
  return { recordCount, itemCount };
}

function renderAccountMaster() {
  const accounts = sortTableRows([...state.accounts], "accountMaster", {
    name: (account) => account.name,
    owner: (account) => account.owner || "",
    aliases: (account) => (account.aliases || []).join(", "),
    recordCount: (account) => countAccountRecords(account).recordCount,
  }, (a, b) => a.name.localeCompare(b.name, "ko"));
  const aliasTotal = accounts.reduce((total, account) => total + (account.aliases ? account.aliases.length : 0), 0);
  const ownerTotal = new Set(accounts.map((account) => account.owner).filter(Boolean)).size;
  el.accountMasterSummary.textContent = `거래처 ${accounts.length.toLocaleString("ko-KR")}곳 · 담당자 ${ownerTotal.toLocaleString("ko-KR")}명 · 등록 별칭 ${aliasTotal.toLocaleString("ko-KR")}개`;
  el.accountMasterNote.innerHTML =
    "<p>엑셀을 반영하면 거래처명이 여기 정식명으로 자동 치환됩니다. 표기가 갈라진 거래처는 별칭을 추가하거나 병합으로 하나로 합칠 수 있습니다. 기존 데이터를 마스터에 맞춰 정리하려면 ‘기존 실적 재매칭’을 누르세요.</p>";

  el.accountMasterRows.innerHTML = accounts.length
    ? accounts
        .map((account) => {
          const counts = countAccountRecords(account);
          const aliasChips = (account.aliases || []).length
            ? account.aliases
                .map(
                  (alias, index) =>
                    `<span class="alias-chip">${escapeHtml(alias)}<button type="button" title="별칭 삭제" data-remove-alias="${account.id}" data-alias-index="${index}">×</button></span>`,
                )
                .join("")
            : `<span class="muted">-</span>`;
          const mergeTargets = accounts
            .filter((other) => other.id !== account.id)
            .map((other) => `<option value="${escapeAttr(other.id)}">${escapeHtml(other.name)}</option>`)
            .join("");
          return `
            <tr>
              <td>
                <strong>${escapeHtml(account.name)}</strong>
                <button class="link-button" type="button" data-edit-account="${account.id}">이름수정</button>
              </td>
              <td>
                <strong>${escapeHtml(account.owner || "미지정")}</strong>
                <button class="link-button" type="button" data-edit-owner="${account.id}">담당수정</button>
              </td>
              <td>
                <div class="alias-cell">${aliasChips}</div>
                <button class="link-button" type="button" data-add-alias="${account.id}">+ 별칭추가</button>
              </td>
              <td class="num">${counts.recordCount.toLocaleString("ko-KR")}건${counts.itemCount ? ` <span class="muted">(품목 ${counts.itemCount})</span>` : ""}</td>
              <td>
                <div class="merge-cell">
                  <select data-merge-target="${account.id}">
                    <option value="">대상 선택</option>
                    ${mergeTargets}
                  </select>
                  <button class="secondary-button" type="button" data-merge-from="${account.id}">병합</button>
                </div>
              </td>
              <td>
                <div class="table-actions">
                  <button class="danger-button" type="button" data-delete-account="${account.id}">삭제</button>
                </div>
              </td>
            </tr>
          `;
        })
        .join("")
    : `<tr><td colspan="6">${emptyState("등록된 거래처가 없습니다. 위에서 추가하거나 엑셀을 반영하면 자동 등록됩니다.")}</td></tr>`;

  document.querySelectorAll("[data-add-alias]").forEach((button) => {
    button.addEventListener("click", () => promptAddAlias(button.dataset.addAlias));
  });
  document.querySelectorAll("[data-remove-alias]").forEach((button) => {
    button.addEventListener("click", () => removeAlias(button.dataset.removeAlias, Number(button.dataset.aliasIndex)));
  });
  document.querySelectorAll("[data-edit-account]").forEach((button) => {
    button.addEventListener("click", () => promptEditAccountName(button.dataset.editAccount));
  });
  document.querySelectorAll("[data-edit-owner]").forEach((button) => {
    button.addEventListener("click", () => promptEditAccountOwner(button.dataset.editOwner));
  });
  document.querySelectorAll("[data-merge-from]").forEach((button) => {
    button.addEventListener("click", () => {
      const select = document.querySelector(`[data-merge-target="${button.dataset.mergeFrom}"]`);
      mergeAccounts(button.dataset.mergeFrom, select ? select.value : "");
    });
  });
  document.querySelectorAll("[data-delete-account]").forEach((button) => {
    button.addEventListener("click", () => deleteAccountMaster(button.dataset.deleteAccount));
  });
}

function promptAddAlias(id) {
  const account = state.accounts.find((a) => a.id === id);
  if (!account) return;
  const input = prompt(`'${account.name}'에 추가할 별칭을 입력하세요. 여러 개는 쉼표로 구분할 수 있습니다.`);
  if (input === null) return;
  const before = (account.aliases || []).length;
  input
    .split(",")
    .map((alias) => alias.trim())
    .filter(Boolean)
    .forEach((alias) => addAliasToAccount(account, alias));
  if ((account.aliases || []).length > before) {
    saveState();
    renderAll();
    showToast("별칭을 추가했습니다.", "success");
  } else {
    showToast("이미 등록된 별칭이거나 정식명과 완전히 같은 이름입니다.", "error");
  }
}

function removeAlias(id, index) {
  const account = state.accounts.find((a) => a.id === id);
  if (!account || !Array.isArray(account.aliases)) return;
  if (index < 0 || index >= account.aliases.length) return;
  account.aliases.splice(index, 1);
  saveState();
  renderAll();
}

function promptEditAccountName(id) {
  const account = state.accounts.find((a) => a.id === id);
  if (!account) return;
  const input = prompt("정식 거래처명을 수정합니다.", account.name);
  if (input === null) return;
  const name = input.trim();
  if (!name) {
    showToast("거래처명은 비울 수 없습니다.", "error");
    return;
  }
  const duplicate = state.accounts.find((a) => a.id !== id && normalizeAccountKey(a.name) === normalizeAccountKey(name));
  if (duplicate) {
    showToast("같은 이름의 거래처가 이미 있습니다. 병합을 사용하세요.", "error");
    return;
  }
  const oldName = account.name;
  account.name = name;
  // 실적/목표/액션의 이 거래처명도 함께 변경
  repointAccountName(oldName, name);
  saveState();
  renderAll();
  showToast("거래처명을 수정했습니다.", "success");
}

function promptEditAccountOwner(id) {
  const account = state.accounts.find((a) => a.id === id);
  if (!account) return;
  const input = prompt(`'${account.name}' 담당자를 입력하세요. 비워두면 미지정으로 표시됩니다.`, account.owner || "");
  if (input === null) return;
  account.owner = input.trim();
  saveState();
  renderAll();
  showToast("담당자를 수정했습니다.", "success");
}

// fromName(정확히 일치하는 문자열)을 toName으로 일괄 변경
function repointAccountName(fromName, toName) {
  const fromKey = normalizeAccountKey(fromName);
  const apply = (item) => {
    if (item.account && normalizeAccountKey(item.account) === fromKey) item.account = toName;
  };
  state.records.forEach(apply);
  state.itemRecords.forEach(apply);
  state.actions.forEach(apply);
  Object.values(state.goals || {}).forEach((goal) => (goal.allocations || []).forEach(apply));
}

function mergeAccounts(fromId, toId) {
  if (!toId) {
    showToast("병합할 대상 거래처를 선택하세요.", "error");
    return;
  }
  if (fromId === toId) return;
  const from = state.accounts.find((a) => a.id === fromId);
  const to = state.accounts.find((a) => a.id === toId);
  if (!from || !to) return;
  if (!confirm(`'${from.name}'을(를) '${to.name}'에 병합할까요?\n'${from.name}'의 모든 실적이 '${to.name}'으로 합쳐지고, '${from.name}'은 별칭으로 남습니다.`)) return;

  // from의 이름+별칭을 to의 별칭으로 흡수
  if (!to.owner && from.owner) to.owner = from.owner;
  addAliasToAccount(to, from.name);
  (from.aliases || []).forEach((alias) => addAliasToAccount(to, alias));
  // from에 연결된 모든 데이터의 거래처명을 to 정식명으로 변경
  const fromKeys = new Set([normalizeAccountKey(from.name), ...(from.aliases || []).map(normalizeAccountKey)]);
  const apply = (item) => {
    if (item.account && fromKeys.has(normalizeAccountKey(item.account))) item.account = to.name;
  };
  state.records.forEach(apply);
  state.itemRecords.forEach(apply);
  state.actions.forEach(apply);
  Object.values(state.goals || {}).forEach((goal) => (goal.allocations || []).forEach(apply));
  // from 제거
  state.accounts = state.accounts.filter((a) => a.id !== fromId);
  saveState();
  renderAll();
  showToast(`'${from.name}'을 '${to.name}'에 병합했습니다.`, "success");
}

function deleteAccountMaster(id) {
  const account = state.accounts.find((a) => a.id === id);
  if (!account) return;
  const counts = countAccountRecords(account);
  if (counts.recordCount || counts.itemCount) {
    if (!confirm(`'${account.name}'에는 실적 ${counts.recordCount}건(품목 ${counts.itemCount}건)이 연결돼 있습니다.\n거래처 마스터에서만 삭제하며 실적 데이터는 그대로 남습니다. 계속할까요?`)) return;
  } else if (!confirm(`'${account.name}'을(를) 마스터에서 삭제할까요?`)) {
    return;
  }
  state.accounts = state.accounts.filter((a) => a.id !== id);
  saveState();
  renderAll();
  showToast("거래처를 삭제했습니다.", "success");
}

// 기존 모든 데이터의 거래처명을 마스터 정식명으로 일괄 재정리
function recanonicalizeAll() {
  if (!state.accounts.length) {
    showToast("먼저 거래처를 등록하세요.", "error");
    return;
  }
  if (!confirm("기존 실적·목표·실행계획의 거래처명을 마스터 정식명으로 일괄 정리할까요?")) return;
  let changed = 0;
  const apply = (item) => {
    if (!item.account) return;
    const resolved = resolveAccountName(item.account);
    if (resolved.matched && resolved.name !== item.account) {
      item.account = resolved.name;
      changed += 1;
    }
  };
  state.records.forEach(apply);
  state.itemRecords.forEach(apply);
  state.actions.forEach(apply);
  Object.values(state.goals || {}).forEach((goal) => (goal.allocations || []).forEach(apply));
  saveState();
  renderAll();
  showToast(`재매칭 완료: ${changed}건의 거래처명을 정리했습니다.`, "success");
}

function renderReport() {
  const goal = getGoalTargets(getGoal());
  const totals = getTotals();
  const accounts = Object.values(groupByAccount(getQuarterRecords())).sort((a, b) => b.sales - a.sales);
  const actions = getQuarterActions();
  const pending = actions.filter((action) => action.status !== "완료");
  const salesGap = Math.max(goal.sales - totals.sales, 0);
  const profitGap = Math.max(goal.profit - totals.profit, 0);
  const pendingSales = sum(pending, "expectedSales");
  const pendingProfit = sum(pending, "expectedProfit");

  el.reportBody.innerHTML = `
    <section class="report-section">
      <h3>${getSelectedPeriodLabel()} 대기업납품 실적</h3>
      <div class="report-metrics">
        <article class="report-metric"><span>매출</span><strong>${won(totals.sales)}</strong></article>
        <article class="report-metric"><span>매출이익</span><strong>${won(totals.profit)}</strong></article>
        <article class="report-metric"><span>마진율</span><strong>${totals.margin.toFixed(1)}%</strong></article>
      </div>
    </section>
    <section class="report-section">
      <h3>목표 대비</h3>
      <p>매출 목표 ${won(goal.sales)} 중 ${rate(totals.sales, goal.sales).toFixed(1)}% 달성, 매출이익 목표 ${won(goal.profit)} 중 ${rate(totals.profit, goal.profit).toFixed(1)}% 달성.</p>
      <p>현재 부족분은 매출 ${won(salesGap)}, 매출이익 ${won(profitGap)}이며 미완료 실행계획의 기대효과는 매출 ${won(pendingSales)}, 이익 ${won(pendingProfit)}입니다.</p>
    </section>
    <section class="report-section">
      <h3>주요 거래처</h3>
      ${accounts.length ? `<ol>${accounts.slice(0, 5).map((row) => `<li>${escapeHtml(row.account)}: 매출 ${won(row.sales)}, 이익 ${won(row.profit)}, 마진 ${row.margin.toFixed(1)}%</li>`).join("")}</ol>` : "<p>등록된 실적이 없습니다.</p>"}
    </section>
    <section class="report-section">
      <h3>다음 실행</h3>
      ${pending.length ? `<ol>${pending.slice(0, 5).map((action) => `<li>${escapeHtml(action.title)} · ${escapeHtml(action.account)} · ${action.due} · ${action.status}</li>`).join("")}</ol>` : "<p>미완료 실행계획이 없습니다.</p>"}
    </section>
  `;
}

function renderData() {
  const supported = isFileSystemSupported();
  const accounts = new Set(state.records.map((record) => record.account));
  const currentMode = fileHandle ? "파일 자동저장" : "브라우저 저장";
  const fileName = fileHandle?.name || "-";
  const cloudUser = supabaseSession?.user?.email || "";
  const cloudMode = supabaseSession ? "Supabase 서버 연결" : "Supabase 미연결";
  const cloudSavedText = supabaseLastSyncedAt ? formatDateTime(supabaseLastSyncedAt) : "-";
  el.supabaseConnectionCard.innerHTML = `
    <h3>${cloudMode}</h3>
    <p>프로젝트: <strong>${SUPABASE_URL.replace("https://", "")}</strong></p>
    <p>로그인 계정: <strong>${escapeHtml(cloudUser || "-")}</strong></p>
    <p>마지막 서버 저장/불러오기: <strong>${escapeHtml(cloudSavedText)}</strong></p>
  `;
  el.supabaseEmailInput.disabled = Boolean(supabaseSession);
  el.supabasePasswordInput.disabled = Boolean(supabaseSession);
  el.supabaseLoginButton.disabled = Boolean(supabaseSession);
  el.supabaseLogoutButton.disabled = !supabaseSession;
  el.supabaseLoadButton.disabled = !supabaseSession;
  el.supabaseSaveButton.disabled = !supabaseSession;
  el.supabaseAutoSaveInput.checked = supabaseAutoSave;
  el.supabaseAutoSaveInput.disabled = !supabaseSession;

  const supportText = supported ? "이 브라우저는 파일 자동저장을 지원합니다." : "이 브라우저에서는 파일 자동저장이 지원되지 않습니다. 수동 백업/복원은 사용할 수 있습니다.";
  const permissionText = needsFilePermission ? "브라우저 권한 재승인이 필요합니다. 기존 파일 연결을 다시 진행해 주세요." : "변경사항은 먼저 브라우저에 저장되고, 파일 연결 시 선택한 JSON 파일에도 자동 저장됩니다.";

  el.fileConnectionCard.innerHTML = `
    <h3>${currentMode}</h3>
    <p>연결 파일: <strong>${escapeHtml(fileName)}</strong></p>
    <p>${supportText}</p>
    <p>${permissionText}</p>
  `;

  el.createDataFileButton.disabled = !supported;
  el.openDataFileButton.disabled = !supported;
  el.reloadDataFileButton.disabled = !fileHandle;
  el.disconnectDataFileButton.disabled = !fileHandle;

  el.dataStats.innerHTML = [
    dataStat("저장 방식", currentMode),
    dataStat("서버 저장", cloudMode),
    dataStat("거래처", `${accounts.size}곳`),
    dataStat("월별 실적", `${state.records.length}건`),
    dataStat("품목 실적", `${state.itemRecords.length}건`),
    dataStat("실행계획", `${state.actions.length}건`),
    dataStat("목표 분기", `${Object.keys(state.goals).length}개`),
    dataStat("마지막 저장", state.lastUpdated || "-"),
    dataStat("파일 저장", lastFileSavedAt || "-"),
    dataStat("데이터 버전", state.version || "1.1"),
  ].join("");
}

function dataStat(label, value) {
  return `<article class="data-stat"><span>${label}</span><strong>${escapeHtml(value)}</strong></article>`;
}

function isFileSystemSupported() {
  return typeof window.showSaveFilePicker === "function" && typeof window.showOpenFilePicker === "function";
}

async function connectNewDataFile() {
  if (!isFileSystemSupported()) {
    showToast("현재 브라우저에서는 파일 자동저장이 지원되지 않습니다.", "error");
    return;
  }

  try {
    const handle = await window.showSaveFilePicker({
      suggestedName: `동부엠티_대기업납품_데이터_${todayString()}.json`,
      types: [{ description: "JSON 데이터 파일", accept: { "application/json": [".json"] } }],
    });
    if (!(await verifyFilePermission(handle, true))) {
      showToast("파일 저장 권한이 필요합니다.", "error");
      return;
    }
    fileHandle = handle;
    needsFilePermission = false;
    await dbSet(FILE_HANDLE_KEY, handle);
    await saveStateToFile(true);
    renderAll();
    showToast("저장파일이 연결됐습니다.", "success");
  } catch (error) {
    if (error.name !== "AbortError") showToast(`저장파일 연결에 실패했습니다: ${error.message}`, "error");
  }
}

async function connectExistingDataFile() {
  if (!isFileSystemSupported()) {
    showToast("현재 브라우저에서는 파일 자동저장이 지원되지 않습니다.", "error");
    return;
  }

  try {
    const [handle] = await window.showOpenFilePicker({
      types: [{ description: "JSON 데이터 파일", accept: { "application/json": [".json"] } }],
      multiple: false,
    });
    if (!(await verifyFilePermission(handle, true))) {
      showToast("파일 저장 권한이 필요합니다.", "error");
      return;
    }
    const hasData = state.records.length || state.actions.length || Object.keys(state.goals).length;
    if (hasData && !confirm("선택한 파일 내용으로 현재 데이터를 바꿀까요? 현재 브라우저 데이터는 덮어쓰기됩니다.")) return;
    fileHandle = handle;
    needsFilePermission = false;
    await dbSet(FILE_HANDLE_KEY, handle);
    await loadStateFromFile({ showMessage: true });
  } catch (error) {
    if (error.name !== "AbortError") showToast(`파일 불러오기에 실패했습니다: ${error.message}`, "error");
  }
}

async function loadStateFromFile({ showMessage = false } = {}) {
  if (!fileHandle) return;
  try {
    if (!(await verifyFilePermission(fileHandle, false))) {
      needsFilePermission = true;
      updateStorageUi();
      showToast("파일 읽기 권한이 필요합니다.", "error");
      return;
    }
    const file = await fileHandle.getFile();
    const text = await file.text();
    if (!text.trim()) {
      await saveStateToFile(true);
      return;
    }
    state = prepareState(JSON.parse(text));
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    needsFilePermission = false;
    renderAll();
    if (showMessage) showToast("파일에서 데이터를 불러왔습니다.", "success");
  } catch (error) {
    showToast(`파일을 읽지 못했습니다: ${error.message}`, "error");
  }
}

async function disconnectDataFile() {
  if (!fileHandle) return;
  if (!confirm("저장파일 연결을 해제할까요? 브라우저에 저장된 데이터는 유지됩니다.")) return;
  clearTimeout(fileSaveTimer);
  fileHandle = null;
  needsFilePermission = false;
  lastFileSavedAt = "";
  await dbDelete(FILE_HANDLE_KEY);
  renderAll();
  showToast("저장파일 연결을 해제했습니다.", "success");
}

function queueFileSave() {
  clearTimeout(fileSaveTimer);
  fileSaveTimer = setTimeout(() => saveStateToFile(false), 700);
}

async function saveStateToFile(immediate) {
  if (!fileHandle) return;
  try {
    if (!(await verifyFilePermission(fileHandle, true))) {
      needsFilePermission = true;
      updateStorageUi();
      if (immediate) showToast("파일 저장 권한이 필요합니다.", "error");
      return;
    }
    syncAllGoalTargets();
    const writable = await fileHandle.createWritable();
    await writable.write(JSON.stringify(state, null, 2));
    await writable.close();
    lastFileSavedAt = formatTimestamp();
    needsFilePermission = false;
    updateStorageUi();
  } catch (error) {
    showToast(`파일 저장에 실패했습니다: ${error.message}`, "error");
  }
}

async function verifyFilePermission(handle, withWrite) {
  const options = withWrite ? { mode: "readwrite" } : { mode: "read" };
  if (typeof handle.queryPermission !== "function") return true;
  if ((await handle.queryPermission(options)) === "granted") return true;
  if (typeof handle.requestPermission !== "function") return false;
  return (await handle.requestPermission(options)) === "granted";
}

function exportJsonBackup() {
  syncAllGoalTargets();
  const payload = JSON.stringify(state, null, 2);
  const blob = new Blob([payload], { type: "application/json" });
  const link = document.createElement("a");
  const url = URL.createObjectURL(blob);
  link.href = url;
  link.download = `동부엠티_대기업납품_백업_${todayString()}.json`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  showToast("JSON 백업 파일을 만들었습니다.", "success");
}

function importJsonBackup(file) {
  if (!confirm("백업 파일 내용으로 현재 데이터를 바꿀까요?")) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      state = prepareState(JSON.parse(reader.result));
      saveState();
      renderAll();
      showToast("백업 데이터를 불러왔습니다.", "success");
    } catch (error) {
      showToast(`백업 파일을 읽지 못했습니다: ${error.message}`, "error");
    }
  };
  reader.readAsText(file);
}

async function restoreFileHandle() {
  if (!isFileSystemSupported()) {
    updateStorageUi();
    return;
  }

  try {
    const handle = await dbGet(FILE_HANDLE_KEY);
    if (!handle) {
      updateStorageUi();
      return;
    }
    const permission = typeof handle.queryPermission === "function" ? await handle.queryPermission({ mode: "readwrite" }) : "granted";
    if (permission === "granted") {
      fileHandle = handle;
      needsFilePermission = false;
      await loadStateFromFile();
    } else {
      fileHandle = handle;
      needsFilePermission = true;
      updateStorageUi();
    }
  } catch {
    updateStorageUi();
  }
}

function openFileDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(FILE_DB_NAME, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(FILE_DB_STORE);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function dbGet(key) {
  const db = await openFileDb();
  return new Promise((resolve, reject) => {
    const request = db.transaction(FILE_DB_STORE, "readonly").objectStore(FILE_DB_STORE).get(key);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function dbSet(key, value) {
  const db = await openFileDb();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(FILE_DB_STORE, "readwrite");
    transaction.objectStore(FILE_DB_STORE).put(value, key);
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error);
  });
}

async function dbDelete(key) {
  const db = await openFileDb();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(FILE_DB_STORE, "readwrite");
    transaction.objectStore(FILE_DB_STORE).delete(key);
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error);
  });
}

function updateStorageUi() {
  if (!el.storageStatus) return;
  if (supabaseSession && fileHandle) {
    el.storageStatus.textContent = needsFilePermission ? "Supabase + 파일 권한 필요" : "Supabase + 파일 자동저장";
  } else if (supabaseSession) {
    el.storageStatus.textContent = supabaseAutoSave ? "Supabase 자동저장" : "Supabase 연결";
  } else if (fileHandle) {
    el.storageStatus.textContent = needsFilePermission ? "파일 권한 필요" : `파일 자동저장: ${fileHandle.name}`;
  } else {
    el.storageStatus.textContent = "브라우저 저장";
  }
}

function formatTimestamp() {
  const date = new Date();
  const year = String(date.getFullYear()).slice(2);
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hour = String(date.getHours()).padStart(2, "0");
  const minute = String(date.getMinutes()).padStart(2, "0");
  return `${year}-${month}-${day} ${hour}:${minute}`;
}

function formatDateTime(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value || "-");
  const year = String(date.getFullYear()).slice(2);
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hour = String(date.getHours()).padStart(2, "0");
  const minute = String(date.getMinutes()).padStart(2, "0");
  return `${year}-${month}-${day} ${hour}:${minute}`;
}

function todayString() {
  const date = new Date();
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function showToast(message, type = "") {
  const toast = document.createElement("div");
  toast.className = `toast ${type}`.trim();
  toast.textContent = message;
  el.toastRoot.appendChild(toast);
  setTimeout(() => toast.remove(), 2600);
}

function groupByAccount(records) {
  return records.reduce((map, record) => {
    if (!map[record.account]) {
      map[record.account] = { account: record.account, sales: 0, cost: 0, profit: 0, margin: 0, items: new Set() };
    }
    map[record.account].sales += record.sales;
    map[record.account].cost += record.cost;
    map[record.account].profit += record.sales - record.cost;
    record.items
      .split(",")
      .map((item) => item.trim())
      .filter(Boolean)
      .forEach((item) => map[record.account].items.add(item));
    map[record.account].margin = margin(map[record.account].sales, map[record.account].profit);
    return map;
  }, {});
}

function completedActions() {
  return getQuarterActions().filter((action) => action.status === "완료").length;
}

function pendingImpactText() {
  const pending = getQuarterActions().filter((action) => action.status !== "완료");
  return `${won(sum(pending, "expectedSales"))}`;
}

function quarterFromMonth(monthValue) {
  const month = Number(monthValue.slice(5, 7));
  return `Q${Math.floor((month - 1) / 3) + 1}`;
}

function getDefaultDueDate() {
  const months = getQuarterMonthCodes(state.activeQuarter);
  const lastMonth = months[months.length - 1];
  return `${state.activeYear}-${lastMonth}-25`;
}

function formatMonth(monthValue) {
  const [year, month] = monthValue.split("-");
  return `${year}.${month}`;
}

function quarterLabel(quarter) {
  if (quarter === ALL_QUARTER) return "전체";
  return `${quarter.replace("Q", "")}분기`;
}

function won(value) {
  const number = Number(value) || 0;
  const sign = number < 0 ? "-" : "";
  const man = Math.round(Math.abs(number) / 10000);
  return `${sign}${man.toLocaleString("ko-KR")}만`;
}

function rate(actual, target) {
  return target > 0 ? (actual / target) * 100 : 0;
}

function margin(sales, profit) {
  return sales > 0 ? (profit / sales) * 100 : 0;
}

function meterClass(value, good, warn) {
  if (value >= good) return "good";
  if (value >= warn) return "warn";
  return "bad";
}

function statusClass(status) {
  if (status === "완료") return "done";
  if (status === "보류") return "hold";
  if (status === "진행중") return "progress";
  return "";
}

function strategyLabel(row, mix) {
  if (row.margin < 10) return "마진 개선";
  if (mix >= 35 && row.margin >= 13) return "핵심 확대";
  if (mix >= 20) return "유지 확대";
  return "품목 제안";
}

function strategyClass(row) {
  if (row.margin < 10) return "hold";
  if (row.margin >= 13) return "done";
  return "progress";
}

function emptyState(text) {
  return `<div class="empty-state">${text}</div>`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeAttr(value) {
  return escapeHtml(value).replaceAll("`", "&#096;");
}
