const state = {
  step: 1,
  selection: {
    type: "Economy",
    speed: "2-4 working days",
    price: 5.9,
  },
  quantity: 1,
  labels: [],
  activeLabelIndex: 0,
  csvMode: false,
  csvEditable: false,
  csvRows: [],
  info: {
    senderName: "",
    senderStreet: "",
    senderCity: "",
    senderState: "",
    senderZip: "",
    recipientName: "",
    recipientStreet: "",
    recipientCountry: "",
    recipientCity: "",
    recipientState: "",
    recipientZip: "",
    packageWeight: "",
    packageDims: "",
    trackingId: "",
    labelId: "",
  },
  customs: {
    summaryDescription: "",
    shipmentType: "",
    senderReference: "",
    importerReference: "",
    items: [
      {
        description: "",
        quantity: "1",
        weightKg: "",
        valueEur: "",
        originCountry: "",
        hsCode: "",
      },
    ],
    completedFor: "",
  },
};

const SUPABASE_URL = "https://pxcqxubehvnyaubqjcrf.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_MfL9s44GmmR9peD1jouetw_aZOa8xVa";
const HISTORY_TABLE = "label_generations";
const ACCOUNT_SETTINGS_TABLE = "account_settings";
const HISTORY_LIMIT = 25;
const HISTORY_FETCH_RETRY_DELAYS_MS = [0, 220, 600];
const WAREHOUSE_MAX_COUNT = 10;
const WAREHOUSE_REMOVE_ANIMATION_MS = 440;
const SHELL_TRANSITION_IN_MS = 340;
const MAIN_VIEW_TRANSITION_OUT_MS = 130;
const MAIN_VIEW_TRANSITION_IN_MS = 320;
const CSV_MODAL_STEP_SWITCH_DELAY_MS = 160;
const AUTH_PARTICLES_WARMUP_STEPS = 42;
const VAT_RATE = 0.21;
const TOTAL_STEPS = 4;
const CUSTOMS_MAX_ITEMS = 12;
const ROUTE_PATHS = {
  login: "/login",
  account: "/account",
  reports: "/reports",
};
const STEP_ROUTE_PATHS = {
  1: "/label-info",
  2: "/label-selection",
  3: "/invoice-confirm",
  4: "/label-pdf",
};
const ROUTE_SUFFIXES = [
  "/index.html",
  ROUTE_PATHS.login,
  ROUTE_PATHS.account,
  ROUTE_PATHS.reports,
  ...Object.values(STEP_ROUTE_PATHS),
];
// Toggle to compare auth logo layouts quickly:
// "inline" => logo replaces Shipr header block inside card.
// "stack" => logo above the card (previous behavior).
const AUTH_LOGO_LAYOUT = "stack";
const ROUTE_BASE_PATH = detectRouteBasePath(window.location.pathname);
const SUPABASE_AUTH_STORAGE_KEY = "sb-pxcqxubehvnyaubqjcrf-auth-token-shipr";

function createSupabaseAuthLock() {
  return async (...args) => {
    const callback = args[args.length - 1];
    if (typeof callback !== "function") {
      return null;
    }

    const lockName = typeof args[0] === "string" ? args[0] : "sb-auth-lock";
    const requestedTimeout = Number(args[1]);
    const timeoutMs = Number.isFinite(requestedTimeout) && requestedTimeout > 0
      ? requestedTimeout
      : 10000;

    const canUseNavigatorLocks =
      typeof navigator !== "undefined" &&
      navigator.locks &&
      typeof navigator.locks.request === "function";

    if (!canUseNavigatorLocks) {
      return callback();
    }

    let timerId = 0;
    let controller = null;
    if (typeof AbortController !== "undefined") {
      controller = new AbortController();
      timerId = window.setTimeout(() => {
        controller.abort("timeout");
      }, timeoutMs + 120);
    }

    try {
      return await navigator.locks.request(
        lockName,
        controller
          ? { mode: "exclusive", signal: controller.signal }
          : { mode: "exclusive" },
        async () => callback()
      );
    } catch (_error) {
      // Fallback: run auth critical section without Web Lock if lock acquisition fails.
      return callback();
    } finally {
      if (timerId) {
        window.clearTimeout(timerId);
      }
    }
  };
}

const supabaseClient =
  window.supabase?.createClient &&
  SUPABASE_URL &&
  SUPABASE_PUBLISHABLE_KEY
    ? window.supabase.createClient(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY, {
        auth: {
          autoRefreshToken: true,
          persistSession: true,
          detectSessionInUrl: true,
          storageKey: SUPABASE_AUTH_STORAGE_KEY,
          lock: createSupabaseAuthLock(),
        },
      })
    : null;

const authGate = document.getElementById("authGate");
const appPage = document.getElementById("appPage");
const authForm = document.getElementById("authForm");
const authEmail = document.getElementById("authEmail");
const authPassword = document.getElementById("authPassword");
const authError = document.getElementById("authError");
const authSignIn = document.getElementById("authSignIn");
const authSignUp = document.getElementById("authSignUp");
const authForgotPassword = document.getElementById("authForgotPassword");
const authLogoLottie = document.getElementById("authLogoLottie");
const authBrand = document.getElementById("authBrand");
const authBrandOriginal = document.getElementById("authBrandOriginal");
const authBrandInlineLogo = document.getElementById("authBrandInlineLogo");
const appLogoLottie = document.getElementById("appLogoLottie");
const accountChip = document.getElementById("accountChip");
const signOutButton = document.getElementById("signOutButton");
const openAccountPageButton = document.getElementById("openAccountPage");
const openReportsPageButton = document.getElementById("openReportsPage");
const closeAccountPageButton = document.getElementById("closeAccountPage");
const closeReportsPageButton = document.getElementById("closeReportsPage");
const builderPage = document.getElementById("builderPage");
const accountPageSection = document.getElementById("accountPageSection");
const reportsPageSection = document.getElementById("reportsPageSection");
const accountCompanyName = document.getElementById("accountCompanyName");
const accountContactName = document.getElementById("accountContactName");
const accountContactEmail = document.getElementById("accountContactEmail");
const accountContactPhone = document.getElementById("accountContactPhone");
const accountBillingAddress = document.getElementById("accountBillingAddress");
const accountTaxId = document.getElementById("accountTaxId");
const accountCustomerId = document.getElementById("accountCustomerId");
const accountPlan = document.getElementById("accountPlan");
const warehouseStatus = document.getElementById("warehouseStatus");
const warehouseList = document.getElementById("warehouseList");
const warehouseAddButton = document.getElementById("warehouseAdd");
const warehouseSaveButton = document.getElementById("warehouseSave");
const accountHistoryStatus = document.getElementById("accountHistoryStatus");
const accountHistoryList = document.getElementById("accountHistoryList");
const accountPreviewMeta = document.getElementById("accountPreviewMeta");
const accountDownloadPdf = document.getElementById("accountDownloadPdf");
const openReceiptModalButton = document.getElementById("openReceiptModal");
const accountBatchPanel = document.getElementById("accountBatchPanel");
const accountBatchList = document.getElementById("accountBatchList");
const accountBatchPreview = document.getElementById("accountBatchPreview");
const accountPdfFrame = document.getElementById("accountPdfFrame");
const receiptModal = document.getElementById("receiptModal");
const receiptModalClose = document.getElementById("receiptModalClose");
const receiptSummary = document.getElementById("receiptSummary");
const receiptTableBody = document.querySelector("#receiptTable tbody");
const receiptDownloadPdf = document.getElementById("receiptDownloadPdf");
const reportsRangeButtons = document.querySelectorAll("[data-report-range]");
const reportsCustomRange = document.getElementById("reportsCustomRange");
const reportsStartDate = document.getElementById("reportsStartDate");
const reportsEndDate = document.getElementById("reportsEndDate");
const reportsApplyCustom = document.getElementById("reportsApplyCustom");
const reportSavingsValue = document.getElementById("reportSavingsValue");
const reportSavingsMeta = document.getElementById("reportSavingsMeta");
const reportPendingReturnsValue = document.getElementById("reportPendingReturnsValue");
const reportPendingReturnsMeta = document.getElementById("reportPendingReturnsMeta");
const reportPendingRefundsValue = document.getElementById("reportPendingRefundsValue");
const reportPendingRefundsMeta = document.getElementById("reportPendingRefundsMeta");
const reportAccountBalanceValue = document.getElementById("reportAccountBalanceValue");
const reportTotalShippingCostsValue = document.getElementById("reportTotalShippingCostsValue");
const reportTotalShippingCostsMeta = document.getElementById("reportTotalShippingCostsMeta");
const reportCarrierAdjustmentsValue = document.getElementById("reportCarrierAdjustmentsValue");
const reportCarrierAdjustmentsMeta = document.getElementById("reportCarrierAdjustmentsMeta");
const reportTotalPaymentsValue = document.getElementById("reportTotalPaymentsValue");
const reportPaymentsChart = document.getElementById("reportPaymentsChart");
const reportPaymentsTooltip = document.getElementById("reportPaymentsTooltip");
const reportPaymentsAxis = document.getElementById("reportPaymentsAxis");
const reportAverageCostValue = document.getElementById("reportAverageCostValue");
const reportAverageCostChart = document.getElementById("reportAverageCostChart");
const reportAverageCostTooltip = document.getElementById("reportAverageCostTooltip");
const reportAverageCostAxis = document.getElementById("reportAverageCostAxis");
const reportAvgDomesticValue = document.getElementById("reportAvgDomesticValue");
const reportAvgDomesticMeta = document.getElementById("reportAvgDomesticMeta");
const reportAvgInternationalValue = document.getElementById("reportAvgInternationalValue");
const reportAvgInternationalMeta = document.getElementById("reportAvgInternationalMeta");
const reportNewShipmentsValue = document.getElementById("reportNewShipmentsValue");
const reportNewShipmentsMeta = document.getElementById("reportNewShipmentsMeta");
const reportDeliveryIssuesValue = document.getElementById("reportDeliveryIssuesValue");
const reportDeliveryIssuesMeta = document.getElementById("reportDeliveryIssuesMeta");
const reportTotalShipmentsValue = document.getElementById("reportTotalShipmentsValue");
const reportShipmentsChart = document.getElementById("reportShipmentsChart");
const reportShipmentsTooltip = document.getElementById("reportShipmentsTooltip");
const reportShipmentsAxis = document.getElementById("reportShipmentsAxis");
const reportServicesDonut = document.getElementById("reportServicesDonut");
const reportServicesLegend = document.getElementById("reportServicesLegend");
const reportServicesTotal = document.getElementById("reportServicesTotal");
const reportZones = document.getElementById("reportZones");
const reportTopRegionsBody = document.getElementById("reportTopRegionsBody");
const reportTopCountriesBody = document.getElementById("reportTopCountriesBody");
const reportRegionsMap = document.getElementById("reportRegionsMap");
const reportRegionsMapTooltip = document.getElementById("reportRegionsMapTooltip");
const reportCountriesMap = document.getElementById("reportCountriesMap");
const reportCountriesMapTooltip = document.getElementById("reportCountriesMapTooltip");

const stepPanels = document.querySelectorAll(".step-panel");
const stepperItems = document.querySelectorAll(".stepper-item");
const stepper = document.getElementById("stepper");
const labelCards = document.querySelectorAll(".label-card");
const summaryService = document.getElementById("summaryService");
const summaryPrice = document.getElementById("summaryPrice");
const summaryQty = document.getElementById("summaryQty");
const summaryTotal = document.getElementById("summaryTotal");
const summaryTracking = document.getElementById("summaryTracking");
const summaryLabelId = document.getElementById("summaryLabelId");
const paymentService = document.getElementById("paymentService");
const paymentTotal = document.getElementById("paymentTotal");
const invoiceQty = document.getElementById("invoiceQty");
const invoiceSubtotal = document.getElementById("invoiceSubtotal");
const invoiceVat = document.getElementById("invoiceVat");
const invoiceCompany = document.getElementById("invoiceCompany");
const invoiceContact = document.getElementById("invoiceContact");
const invoiceEmail = document.getElementById("invoiceEmail");
const invoicePhone = document.getElementById("invoicePhone");
const invoiceAddress = document.getElementById("invoiceAddress");
const invoiceTaxId = document.getElementById("invoiceTaxId");
const quantityInput = document.getElementById("labelQuantity");
const batchPanel = document.getElementById("batchPanel");
const batchList = document.getElementById("batchList");
const batchPreview = document.getElementById("batchPreview");
const autoCsvButton = document.getElementById("autoCsv");
const csvEditToggle = document.getElementById("csvEditToggle");
const csvSection = document.getElementById("csvSection");
const csvTableBody = document.querySelector("#csvTable tbody");
const labelForm = document.getElementById("labelForm");
const pdfFrame = document.getElementById("pdfFrame");
const downloadPdf = document.getElementById("downloadPdf");
const startOver = document.getElementById("startOver");
const payButton = document.getElementById("payButton");
const autoFillButton = document.getElementById("autoFill");
const labelError = document.getElementById("labelError");
const customsGhostPanel = document.getElementById("customsGhostPanel");
const customsScopeMeta = document.getElementById("customsScopeMeta");
const customsError = document.getElementById("customsError");
const customsSummaryDescriptionInput = document.getElementById("customsSummaryDescription");
const customsShipmentTypeInput = document.getElementById("customsShipmentType");
const customsSenderReferenceInput = document.getElementById("customsSenderReference");
const customsImporterReferenceInput = document.getElementById("customsImporterReference");
const customsItemsList = document.getElementById("customsItemsList");
const customsAddItemButton = document.getElementById("customsAddItem");
const customsBackButton = document.getElementById("customsBack");
const customsContinueButton = document.getElementById("customsContinue");
const openRestrictedGoodsModalButton = document.getElementById("openRestrictedGoodsModal");
const restrictedGoodsModal = document.getElementById("restrictedGoodsModal");
const restrictedGoodsModalClose = document.getElementById("restrictedGoodsModalClose");

// Provider dropdown
const providerDropdown = document.getElementById("providerDropdown");
const providerTrigger = document.getElementById("providerTrigger");
const providerMenu = document.getElementById("providerMenu");
const providerStatus = document.getElementById("providerStatus");

// CSV modal
const csvModal = document.getElementById("csvModal");
const csvModalClose = document.getElementById("csvModalClose");
const csvUploadTrigger = document.getElementById("csvUploadTrigger");
const csvUploadStep = document.getElementById("csvUploadStep");
const csvMappingStep = document.getElementById("csvMappingStep");
const csvMapMeta = document.getElementById("csvMapMeta");
const csvMappingBody = document.getElementById("csvMappingBody");
const csvMapBack = document.getElementById("csvMapBack");
const csvMapApply = document.getElementById("csvMapApply");
const csvMapError = document.getElementById("csvMapError");
const csvDropzone = document.getElementById("csvDropzone");
const csvFileInput = document.getElementById("csvFileInput");

const inputMap = {
  senderName: document.getElementById("senderName"),
  senderStreet: document.getElementById("senderStreet"),
  senderCity: document.getElementById("senderCity"),
  senderState: document.getElementById("senderState"),
  senderZip: document.getElementById("senderZip"),
  recipientName: document.getElementById("recipientName"),
  recipientStreet: document.getElementById("recipientStreet"),
  recipientCountry: document.getElementById("recipientCountry"),
  recipientCity: document.getElementById("recipientCity"),
  recipientState: document.getElementById("recipientState"),
  recipientZip: document.getElementById("recipientZip"),
  packageWeight: document.getElementById("packageWeight"),
  packageDims: document.getElementById("packageDims"),
  trackingId: document.getElementById("trackingId"),
  labelId: document.getElementById("labelId"),
};

const countryFlag = document.getElementById("countryFlag");

const labelRequiredFields = [
  "senderName",
  "senderStreet",
  "senderCity",
  "senderState",
  "senderZip",
  "recipientName",
  "recipientStreet",
  "recipientCountry",
  "recipientCity",
  "recipientState",
  "recipientZip",
  "packageWeight",
  "packageDims",
];

const previewService = document.getElementById("previewService");
const previewTracking = document.getElementById("previewTracking");
const previewFrom = document.getElementById("previewFrom");
const previewTo = document.getElementById("previewTo");
const previewWeight = document.getElementById("previewWeight");
const previewDims = document.getElementById("previewDims");

let currentPdfUrl = "";
let currentBatchPdfUrl = "";
let currentUser = null;
let historyRecords = [];
let historyStore = "supabase";
let accountActiveRecord = null;
let accountActiveHistoryIndex = -1;
let accountActiveLabelIndex = 0;
let accountLabels = [];
let accountBatchPdfUrl = "";
let warehouseRecords = [];
let warehouseDirty = false;
let warehouseSaving = false;
let warehouseLoadRequestToken = 0;
let warehouseEnteringIds = new Set();
let authParticlesStarted = false;
let currentMainView = "builder";
let authShellTransitionToken = 0;
let mainViewTransitionToken = 0;
let csvMappingDraft = null;
let csvModalStepState = "upload";
let csvModalStepTransitionToken = 0;
let reportsRangeDays = 30;
let reportsRangeMode = "preset";
let reportsCustomStart = "";
let reportsCustomEnd = "";
let reportsActiveServiceIndex = -1;
let reportsActiveCountryKey = "";
let reportsActiveRegionKey = "";
let reportsDomesticGeoJson = null;
let reportsWorldGeoJson = null;
let reportsGeoLoadPromise = null;
let historyLoadRequestToken = 0;
let customsGhostVisible = false;
let providerStatusTimer = 0;
let shopifyConnection = null;

const DOMESTIC_COUNTRY_ALIASES = new Set([
  "domestic",
  "home",
  "home market",
  "local",
  "belgium",
  "belgie",
  "belgique",
  "be",
]);

const EU_COUNTRY_CODES = new Set([
  "AT",
  "BE",
  "BG",
  "HR",
  "CY",
  "CZ",
  "DK",
  "EE",
  "FI",
  "FR",
  "DE",
  "GR",
  "HU",
  "IE",
  "IT",
  "LV",
  "LT",
  "LU",
  "MT",
  "NL",
  "PL",
  "PT",
  "RO",
  "SK",
  "SI",
  "ES",
  "SE",
]);

const COUNTRY_ALIAS_TO_CODE = {
  ue: "EU",
  eu: "EU",
  "european union": "EU",
  "union europeenne": "EU",
  "union européenne": "EU",
  "pays bas": "NL",
  nederland: "NL",
  hollande: "NL",
  allemagne: "DE",
  espagne: "ES",
  italie: "IT",
  autriche: "AT",
  suede: "SE",
  "suède": "SE",
  grece: "GR",
  "grèce": "GR",
  tchequie: "CZ",
  "république tchèque": "CZ",
  slovaquie: "SK",
  "belgie": "BE",
  belgique: "BE",
};

const NORMALIZED_COUNTRY_CODE_MAP = (() => {
  if (typeof COUNTRY_CODES !== "object" || !COUNTRY_CODES) return {};
  const map = {};
  Object.entries(COUNTRY_CODES).forEach(([name, code]) => {
    map[normalizeNameKey(name)] = String(code || "").toUpperCase();
  });
  return map;
})();

function resolveCountryCode(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  if (/^[A-Za-z]{2}$/.test(raw)) {
    return raw.toUpperCase();
  }
  const lowered = raw.toLowerCase();
  const normalized = normalizeNameKey(raw);
  const aliasCode = COUNTRY_ALIAS_TO_CODE[normalized] || COUNTRY_ALIAS_TO_CODE[lowered];
  if (aliasCode) return aliasCode;

  if (typeof COUNTRY_CODES === "object" && COUNTRY_CODES) {
    const direct = COUNTRY_CODES[lowered];
    if (direct) return String(direct).toUpperCase();
  }

  const normalizedCode = NORMALIZED_COUNTRY_CODE_MAP[normalized];
  if (normalizedCode) return normalizedCode;
  return "";
}

function isEuCountry(value) {
  const code = resolveCountryCode(value);
  if (!code) return false;
  if (code === "EU") return true;
  return EU_COUNTRY_CODES.has(code);
}

function createEmptyCustomsItem() {
  return {
    description: "",
    quantity: "1",
    weightKg: "",
    valueEur: "",
    originCountry: "",
    hsCode: "",
  };
}

function getShipmentCountryValues() {
  if (state.csvMode && Array.isArray(state.csvRows) && state.csvRows.length) {
    return state.csvRows.map((row) => String(row?.recipientCountry || "").trim()).filter(Boolean);
  }
  return [String(state.info.recipientCountry || inputMap.recipientCountry?.value || "").trim()].filter(
    Boolean
  );
}

function requiresCustomsDeclaration() {
  const countries = getShipmentCountryValues();
  if (!countries.length) return false;
  return countries.some((country) => !isEuCountry(country));
}

function getCustomsRequirementSignature() {
  const countries = getShipmentCountryValues()
    .map((country) => resolveCountryCode(country) || normalizeNameKey(country))
    .filter(Boolean)
    .sort();
  return `${state.csvMode ? "csv" : "single"}::${countries.join("|")}`;
}

function invalidateCustomsCompletion() {
  state.customs.completedFor = "";
}

const RETAIL_BASE_RATE = {
  Economy: 7.1,
  Priority: 12.2,
  "International Express": 21.8,
};

const REPORT_SERVICE_COLOR_MAP = {
  Economy: "#7747e3",
  Priority: "#8f66ed",
  "International Express": "#a18af0",
};

const REPORT_SERVICE_FALLBACK_COLORS = [
  "#7e58de",
  "#9b77ef",
  "#6f4fcb",
  "#b09af2",
  "#8b6dde",
];

const REPORT_MOCK_COUNTRIES = [
  { name: "France", count: 26, spend: 201.4 },
  { name: "Netherlands", count: 22, spend: 168.3 },
  { name: "Germany", count: 19, spend: 150.5 },
  { name: "Luxembourg", count: 12, spend: 92.1 },
  { name: "Spain", count: 9, spend: 74.8 },
  { name: "Italy", count: 7, spend: 61.3 },
];

const REPORT_MOCK_REGIONS = [
  { name: "North Region", count: 56, spend: 436.8 },
  { name: "South Region", count: 31, spend: 241.8 },
  { name: "Capital Region", count: 18, spend: 140.4 },
];

const REPORT_DOMESTIC_GEOJSON_URL = "assets/belgium-provinces.geojson";
const REPORT_WORLD_GEOJSON_URL = "assets/world-countries.geojson";

function normalizePathname(pathname) {
  if (!pathname) return "/";
  const withLeadingSlash = pathname.startsWith("/") ? pathname : `/${pathname}`;
  const trimmed = withLeadingSlash.replace(/\/+$/, "");
  return trimmed || "/";
}

function detectRouteBasePath(pathname) {
  const normalized = normalizePathname(pathname);
  for (const suffix of ROUTE_SUFFIXES) {
    if (!suffix) continue;
    if (normalized === suffix) return "";
    if (normalized.endsWith(suffix)) {
      return normalized.slice(0, -suffix.length);
    }
  }
  const looksLikeFile = /\.[a-z0-9]+$/i.test(normalized.split("/").pop() || "");
  if (normalized !== "/" && !looksLikeFile) {
    return normalized;
  }
  return "";
}

function clampStep(value) {
  const maxSteps = stepPanels?.length || TOTAL_STEPS;
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return 1;
  return Math.max(1, Math.min(maxSteps, Math.trunc(numeric)));
}

function buildRoutePath(suffix) {
  const normalizedSuffix = normalizePathname(suffix || "/");
  const base = ROUTE_BASE_PATH || "";
  const fullPath = `${base}${normalizedSuffix}`;
  return normalizePathname(fullPath);
}

function getRelativeRoutePath(pathname) {
  const normalized = normalizePathname(pathname);
  const base = normalizePathname(ROUTE_BASE_PATH || "/");
  if (base === "/") return normalized;
  if (normalized === base) return "/";
  if (normalized.startsWith(`${base}/`)) {
    return normalized.slice(base.length);
  }
  return normalized;
}

function parseRouteFromLocation() {
  const path = getRelativeRoutePath(window.location.pathname);
  if (path === ROUTE_PATHS.login) {
    return { view: "login" };
  }
  if (path === ROUTE_PATHS.account) {
    return { view: "account" };
  }
  if (path === ROUTE_PATHS.reports) {
    return { view: "reports" };
  }

  const stepEntry = Object.entries(STEP_ROUTE_PATHS).find(([, routePath]) => routePath === path);
  if (stepEntry) {
    return { view: "builder", step: clampStep(Number(stepEntry[0])) };
  }

  const hash = window.location.hash || "";
  if (hash === "#account") {
    return { view: "account" };
  }
  if (hash === "#login") {
    return { view: "login" };
  }
  const hashStep = hash.match(/step-(\d+)/);
  if (hashStep) {
    return { view: "builder", step: clampStep(Number(hashStep[1])) };
  }

  return { view: "builder", step: clampStep(state.step) };
}

function routeToPath(route) {
  if (!route || route.view === "builder") {
    const step = clampStep(route?.step || state.step);
    return buildRoutePath(STEP_ROUTE_PATHS[step] || STEP_ROUTE_PATHS[1]);
  }
  if (route.view === "login") {
    return buildRoutePath(ROUTE_PATHS.login);
  }
  if (route.view === "account") {
    return buildRoutePath(ROUTE_PATHS.account);
  }
  if (route.view === "reports") {
    return buildRoutePath(ROUTE_PATHS.reports);
  }
  return buildRoutePath(STEP_ROUTE_PATHS[clampStep(state.step)]);
}

function routeToState(route) {
  if (!route || route.view === "builder") {
    const step = clampStep(route?.step || state.step);
    return { view: "builder", step, customs: Boolean(route?.customs && step === 1) };
  }
  return { view: route.view };
}

function updateRoute(route, options = {}) {
  const { replace = false } = options;
  const nextPath = routeToPath(route);
  const nextState = routeToState(route);

  const currentPath = normalizePathname(window.location.pathname);
  const samePath = currentPath === normalizePathname(nextPath);
  const sameState =
    history.state?.view === nextState.view &&
    Number(history.state?.step || 0) === Number(nextState.step || 0) &&
    Boolean(history.state?.customs) === Boolean(nextState.customs);

  if (replace) {
    if (!samePath || !sameState) {
      history.replaceState(nextState, "", nextPath);
    }
    return;
  }
  if (samePath && sameState) return;
  history.pushState(nextState, "", nextPath);
}

async function ensureReportsGeoDataLoaded() {
  if (reportsDomesticGeoJson && reportsWorldGeoJson) return;
  if (reportsGeoLoadPromise) {
    await reportsGeoLoadPromise;
    return;
  }
  reportsGeoLoadPromise = Promise.all([
    fetch(REPORT_DOMESTIC_GEOJSON_URL).then((response) => {
      if (!response.ok) {
        throw new Error(`Failed to load domestic regions GeoJSON (${response.status})`);
      }
      return response.json();
    }),
    fetch(REPORT_WORLD_GEOJSON_URL).then((response) => {
      if (!response.ok) {
        throw new Error(`Failed to load world GeoJSON (${response.status})`);
      }
      return response.json();
    }),
  ])
    .then(([domesticData, worldData]) => {
      reportsDomesticGeoJson = domesticData;
      reportsWorldGeoJson = worldData;
    })
    .catch((error) => {
      console.warn("Reports map data failed to load:", error);
      reportsDomesticGeoJson = null;
      reportsWorldGeoJson = null;
    })
    .finally(() => {
      reportsGeoLoadPromise = null;
    });

  await reportsGeoLoadPromise;
}

function normalizeNameKey(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function setProviderStatus(message = "", options = {}) {
  if (!providerStatus) return;
  const { kind = "info", persist = false } = options;
  providerStatus.textContent = String(message || "").trim();
  providerStatus.classList.remove("is-success", "is-error");
  if (kind === "success") {
    providerStatus.classList.add("is-success");
  } else if (kind === "error") {
    providerStatus.classList.add("is-error");
  }
  if (providerStatusTimer) {
    window.clearTimeout(providerStatusTimer);
    providerStatusTimer = 0;
  }
  if (!persist && providerStatus.textContent) {
    providerStatusTimer = window.setTimeout(() => {
      if (!shopifyConnection) {
        providerStatus.textContent = "";
        providerStatus.classList.remove("is-success", "is-error");
      } else {
        providerStatus.textContent = `Shopify connected: ${shopifyConnection.shop}`;
        providerStatus.classList.add("is-success");
      }
      providerStatusTimer = 0;
    }, 2800);
  }
}

function normalizeShopDomain(value) {
  let normalized = String(value || "")
    .trim()
    .toLowerCase()
    .replace(/^https?:\/\//, "")
    .replace(/\/.*$/, "");
  if (!normalized) return "";
  if (!normalized.endsWith(".myshopify.com") && /^[a-z0-9][a-z0-9-]*$/.test(normalized)) {
    normalized = `${normalized}.myshopify.com`;
  }
  if (!/^[a-z0-9][a-z0-9-]*\.myshopify\.com$/.test(normalized)) {
    return "";
  }
  return normalized;
}

async function getAuthAccessToken() {
  if (!supabaseClient) return "";
  const {
    data: { session },
  } = await supabaseClient.auth.getSession();
  return String(session?.access_token || "");
}

async function fetchApiWithAuth(path, options = {}) {
  const { timeoutMs = 15000, ...requestOptions } = options;
  const token = await getAuthAccessToken();
  if (!token) {
    throw new Error("You must be signed in to use provider import.");
  }
  const headers = new Headers(requestOptions.headers || {});
  headers.set("Authorization", `Bearer ${token}`);
  if (requestOptions.body && !headers.has("Content-Type")) {
    headers.set("Content-Type", "application/json");
  }
  const controller = new AbortController();
  const timeoutId = window.setTimeout(() => {
    controller.abort("request-timeout");
  }, Math.max(3000, Number(timeoutMs) || 15000));
  try {
    const response = await fetch(path, {
      ...requestOptions,
      headers,
      signal: controller.signal,
    });
    const contentType = response.headers.get("content-type") || "";
    const payload = contentType.includes("application/json")
      ? await response.json().catch(() => ({}))
      : await response.text().catch(() => "");
    if (!response.ok) {
      const message =
        (payload && typeof payload === "object" && payload.error) ||
        (typeof payload === "string" ? payload : "") ||
        `Request failed (${response.status})`;
      throw new Error(message);
    }
    return payload;
  } catch (error) {
    if (error?.name === "AbortError") {
      throw new Error("Shopify request timed out. Check Worker route and secrets.");
    }
    throw error;
  } finally {
    window.clearTimeout(timeoutId);
  }
}

function updateShopifyProviderStatus() {
  if (!providerStatus) return;
  if (!currentUser) {
    setProviderStatus("", { persist: true });
    return;
  }
  if (shopifyConnection) {
    setProviderStatus(`Shopify connected: ${shopifyConnection.shop}`, {
      kind: "success",
      persist: true,
    });
    return;
  }
  setProviderStatus("", { persist: true });
}

async function loadShopifyConnectionStatus(options = {}) {
  const { quiet = false } = options;
  if (!currentUser) {
    shopifyConnection = null;
    updateShopifyProviderStatus();
    return null;
  }
  try {
    const data = await fetchApiWithAuth("/api/shopify/connection");
    const connection =
      data && typeof data === "object" && data.connection && typeof data.connection === "object"
        ? data.connection
        : null;
    shopifyConnection = connection
      ? {
          shop: String(connection.shop || "").trim(),
          connectedAt: String(connection.connectedAt || "").trim(),
          scopes: String(connection.scopes || "").trim(),
        }
      : null;
    updateShopifyProviderStatus();
    return shopifyConnection;
  } catch (error) {
    shopifyConnection = null;
    updateShopifyProviderStatus();
    if (!quiet) {
      setProviderStatus(error.message || "Could not load Shopify connection.", {
        kind: "error",
      });
    }
    return null;
  }
}

function consumeShopifyCallbackParams() {
  const params = new URLSearchParams(window.location.search || "");
  if (params.get("provider") !== "shopify") return;

  const status = params.get("shopify");
  if (status === "connected") {
    const shop = normalizeShopDomain(params.get("shop"));
    if (shop) {
      setProviderStatus(`Shopify connected: ${shop}`, { kind: "success" });
    } else {
      setProviderStatus("Shopify connected.", { kind: "success" });
    }
  } else if (status === "error") {
    const message = String(params.get("message") || "").trim() || "Could not connect Shopify.";
    setProviderStatus(message, { kind: "error" });
  }

  const cleanUrl = `${window.location.pathname}${window.location.hash || ""}`;
  history.replaceState(history.state, "", cleanUrl);
}

function mapShopifyImportRows(rows) {
  if (!Array.isArray(rows)) return [];
  return rows
    .map((row) => {
      if (!row || typeof row !== "object") return null;
      return {
        senderName: String(row.senderName || "").trim(),
        senderStreet: String(row.senderStreet || "").trim(),
        senderCity: String(row.senderCity || "").trim(),
        senderState: String(row.senderState || "").trim(),
        senderZip: String(row.senderZip || "").trim(),
        recipientName: String(row.recipientName || "").trim(),
        recipientStreet: String(row.recipientStreet || "").trim(),
        recipientCity: String(row.recipientCity || "").trim(),
        recipientState: String(row.recipientState || "").trim(),
        recipientZip: String(row.recipientZip || "").trim(),
        recipientCountry: String(row.recipientCountry || "").trim(),
        packageWeight: String(row.packageWeight || "").trim(),
        packageDims: String(row.packageDims || "").trim(),
      };
    })
    .filter(Boolean);
}

function applyImportedRows(rows, sourceLabel) {
  if (!Array.isArray(rows) || rows.length === 0) {
    throw new Error(`No usable ${sourceLabel} rows found.`);
  }
  clearBatchState();
  state.csvRows = rows;
  setCsvMode(true);
  setCsvEditMode(false);
  renderCsvTable();
  if (labelError) {
    labelError.classList.remove("is-visible");
  }
  updatePreview();
  updateSummary();
  updatePayment();
}

async function beginShopifyInstall() {
  const defaultShop = shopifyConnection?.shop || "";
  const typed = window.prompt(
    "Enter your Shopify store domain (example: your-store.myshopify.com)",
    defaultShop
  );
  if (typed === null) return;
  const shop = normalizeShopDomain(typed);
  if (!shop) {
    setProviderStatus("Enter a valid .myshopify.com domain.", { kind: "error" });
    return;
  }
  try {
    setProviderStatus("Connecting Shopify...", { persist: true });
    const data = await fetchApiWithAuth("/api/shopify/install-link", {
      method: "POST",
      body: JSON.stringify({ shop }),
    });
    if (!data || typeof data !== "object" || !data.url) {
      throw new Error("Shopify install URL was not returned.");
    }
    window.location.assign(String(data.url));
  } catch (error) {
    setProviderStatus(error.message || "Could not start Shopify connect flow.", {
      kind: "error",
    });
  }
}

async function importShopifyOrders(shop) {
  const normalizedShop = normalizeShopDomain(shop);
  if (!normalizedShop) {
    setProviderStatus("Connect Shopify before importing.", { kind: "error" });
    return;
  }

  try {
    setProviderStatus(`Importing orders from ${normalizedShop}...`, { persist: true });
    const data = await fetchApiWithAuth("/api/shopify/import-orders", {
      method: "POST",
      body: JSON.stringify({
        shop: normalizedShop,
        limit: 50,
      }),
    });
    const rows = mapShopifyImportRows(data?.rows);
    applyImportedRows(rows, "Shopify");
    setProviderStatus(`Imported ${rows.length} orders from ${normalizedShop}.`, {
      kind: "success",
    });
    const triggerText = providerTrigger?.querySelector("span");
    if (triggerText) {
      triggerText.textContent = `Imported ${rows.length} Shopify orders`;
      window.setTimeout(() => {
        triggerText.textContent = "Import from provider";
      }, 2200);
    }
  } catch (error) {
    const message = String(error?.message || "Shopify import failed.");
    if (
      /connection expired|reconnect shopify|invalid api key or access token|unrecognized login|wrong password/i.test(
        message
      )
    ) {
      shopifyConnection = null;
      updateShopifyProviderStatus();
      setProviderStatus("Shopify connection expired. Click Shopify again to reconnect.", {
        kind: "error",
        persist: true,
      });
      return;
    }
    setProviderStatus(message, { kind: "error" });
  }
}

async function handleShopifyProviderAction() {
  if (!currentUser) {
    setProviderStatus("Sign in before importing from Shopify.", { kind: "error" });
    return;
  }
  if (!shopifyConnection?.shop) {
    await beginShopifyInstall();
    return;
  }
  await importShopifyOrders(shopifyConnection.shop);
}

function getFeaturePolygons(geometry) {
  if (!geometry || !geometry.type) return [];
  if (geometry.type === "Polygon") {
    return [geometry.coordinates];
  }
  if (geometry.type === "MultiPolygon") {
    return geometry.coordinates;
  }
  if (geometry.type === "GeometryCollection") {
    return (geometry.geometries || []).flatMap((item) => getFeaturePolygons(item));
  }
  return [];
}

function getFeatureBounds(features) {
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;

  features.forEach((feature) => {
    const polygons = getFeaturePolygons(feature?.geometry);
    polygons.forEach((polygon) => {
      polygon.forEach((ring) => {
        ring.forEach((coord) => {
          const x = Number(coord?.[0]);
          const y = Number(coord?.[1]);
          if (!Number.isFinite(x) || !Number.isFinite(y)) return;
          minX = Math.min(minX, x);
          minY = Math.min(minY, y);
          maxX = Math.max(maxX, x);
          maxY = Math.max(maxY, y);
        });
      });
    });
  });

  if (!Number.isFinite(minX) || !Number.isFinite(minY)) {
    return null;
  }
  return { minX, minY, maxX, maxY };
}

function createGeoProjector(bounds, width, height, padding = 8, options = {}) {
  const { isGeographic = false } = options;
  if (!bounds) {
    return () => [width / 2, height / 2];
  }
  const midLat = (bounds.minY + bounds.maxY) / 2;
  const lonScale = isGeographic ? Math.cos((midLat * Math.PI) / 180) : 1;
  const minX = bounds.minX * lonScale;
  const maxX = bounds.maxX * lonScale;
  const minY = bounds.minY;
  const maxY = bounds.maxY;

  const spanX = Math.max(1e-9, maxX - minX);
  const spanY = Math.max(1e-9, maxY - minY);
  const drawWidth = Math.max(1, width - padding * 2);
  const drawHeight = Math.max(1, height - padding * 2);
  const scale = Math.min(drawWidth / spanX, drawHeight / spanY);
  const offsetX = (width - spanX * scale) / 2;
  const offsetY = (height - spanY * scale) / 2;
  return (coord) => {
    const lon = Number(coord?.[0]) || 0;
    const lat = Number(coord?.[1]) || 0;
    const x = offsetX + (lon * lonScale - minX) * scale;
    const y = offsetY + (maxY - lat) * scale;
    return [x, y];
  };
}

function geometryToPath(geometry, projector) {
  const polygons = getFeaturePolygons(geometry);
  const parts = [];
  polygons.forEach((polygon) => {
    polygon.forEach((ring) => {
      if (!Array.isArray(ring) || ring.length < 2) return;
      ring.forEach((coord, index) => {
        const [x, y] = projector(coord);
        parts.push(`${index === 0 ? "M" : "L"}${x.toFixed(2)} ${y.toFixed(2)}`);
      });
      parts.push("Z");
    });
  });
  return parts.join(" ");
}

function getCountryFeatureName(feature) {
  const properties = feature?.properties || {};
  return (
    properties.name ||
    properties.admin ||
    properties.name_long ||
    properties.NAME ||
    properties.COUNTRY ||
    "Unknown"
  );
}

function getDomesticProvinceName(feature) {
  const properties = feature?.properties || {};
  return properties.province || properties.NAME_2 || properties.name || "Unknown Province";
}

function getDomesticRegionName(feature) {
  const properties = feature?.properties || {};
  const raw =
    properties.region ||
    properties.NAME_1 ||
    properties.region_name ||
    properties.regio ||
    "";
  const normalized = normalizeNameKey(raw);
  if (normalized.includes("flanders") || normalized.includes("vlaanderen")) return "North Region";
  if (normalized.includes("wallonia") || normalized.includes("wallonie")) return "South Region";
  if (normalized.includes("brussels") || normalized.includes("bruxelles")) {
    return "Capital Region";
  }
  return "North Region";
}

const sampleProfiles = [
  {
    sender: {
      name: "Harper Clarke",
      street: "Westblaak 84",
      city: "Rotterdam",
      state: "South Holland",
      zip: "3012",
      country: "Netherlands",
    },
    recipient: {
      name: "Jordan Reyes",
      street: "Rue de Rivoli 240",
      city: "Paris",
      state: "Ile-de-France",
      zip: "75001",
      country: "France",
    },
  },
  {
    sender: {
      name: "Avery Grant",
      street: "Avenue Foch 12",
      city: "Lille",
      state: "Hauts-de-France",
      zip: "59800",
      country: "France",
    },
    recipient: {
      name: "Riley Chen",
      street: "Neumarkt 18",
      city: "Cologne",
      state: "North Rhine-Westphalia",
      zip: "50667",
      country: "Germany",
    },
  },
  {
    sender: {
      name: "Logan Pierce",
      street: "Via Torino 47",
      city: "Milan",
      state: "Lombardy",
      zip: "20123",
      country: "Italy",
    },
    recipient: {
      name: "Morgan Diaz",
      street: "Carrer de Balmes 125",
      city: "Barcelona",
      state: "Catalonia",
      zip: "08008",
      country: "Spain",
    },
  },
];

const csvColumns = [
  { key: "senderName", label: "From Name" },
  { key: "senderStreet", label: "From Street" },
  { key: "senderCity", label: "From City" },
  { key: "senderState", label: "From Region" },
  { key: "senderZip", label: "From Postal Code" },
  { key: "recipientName", label: "To Name" },
  { key: "recipientStreet", label: "To Street" },
  { key: "recipientCity", label: "To City" },
  { key: "recipientState", label: "To Region" },
  { key: "recipientZip", label: "To Postal Code" },
  { key: "recipientCountry", label: "Country" },
  { key: "packageWeight", label: "Weight" },
  { key: "packageDims", label: "Dims" },
];

const csvReviewColumns = csvColumns.filter(
  (column) =>
    column.key !== "senderName" &&
    column.key !== "senderStreet" &&
    column.key !== "senderCity" &&
    column.key !== "senderState" &&
    column.key !== "senderZip"
);

const CSV_REQUIRED_FIELDS = new Set([
  "recipientName",
  "recipientStreet",
  "recipientCity",
  "recipientState",
  "recipientZip",
  "recipientCountry",
  "packageWeight",
  "packageDims",
]);

const CSV_HEADER_ALIASES = {
  senderName: [
    "from_name",
    "sender_name",
    "ship_from_name",
    "origin_name",
    "warehouse_name",
    "location_name",
    "location",
  ],
  senderStreet: [
    "from_street",
    "sender_street",
    "from_address",
    "from_address1",
    "from_address_1",
    "from_address_line_1",
    "from_line_1",
    "sender_address",
    "sender_address1",
    "sender_address_1",
    "sender_address_line_1",
    "address_line_1",
    "line_1",
    "ship_from_address1",
    "ship_from_address_1",
    "origin_address1",
    "origin_address_1",
    "location_address1",
    "location_address_1",
    "warehouse_address1",
    "warehouse_address_1",
  ],
  senderCity: [
    "from_city",
    "sender_city",
    "ship_from_city",
    "origin_city",
    "location_city",
    "warehouse_city",
  ],
  senderState: [
    "from_state",
    "from_region",
    "sender_state",
    "sender_region",
    "sender_province",
    "ship_from_state",
    "ship_from_region",
    "ship_from_province",
    "origin_state",
    "origin_region",
    "origin_province",
    "location_state",
    "location_region",
    "location_province",
  ],
  senderZip: [
    "from_zip",
    "from_postcode",
    "from_postal_code",
    "sender_zip",
    "sender_postcode",
    "sender_postal_code",
    "ship_from_zip",
    "ship_from_postcode",
    "ship_from_postal_code",
    "origin_zip",
    "origin_postcode",
    "origin_postal_code",
    "location_zip",
    "location_postcode",
    "location_postal_code",
  ],
  recipientName: [
    "to_name",
    "recipient_name",
    "shipping_name",
    "ship_name",
    "ship_to_name",
    "delivery_name",
    "shipping_full_name",
    "customer_name",
    "buyer_name",
    "consignee_name",
    "company",
    "shipping_company",
  ],
  recipientStreet: [
    "to_street",
    "to_address",
    "to_address1",
    "to_address_1",
    "to_address_line_1",
    "to_line_1",
    "recipient_street",
    "recipient_address",
    "recipient_address1",
    "recipient_address_1",
    "recipient_address_line_1",
    "shipping_address",
    "shipping_address1",
    "shipping_address_1",
    "shipping_address_line_1",
    "ship_address",
    "ship_address1",
    "ship_address_1",
    "ship_address_line_1",
    "ship_to_address",
    "ship_to_address1",
    "ship_to_address_1",
    "ship_to_address_line_1",
    "delivery_address",
    "delivery_address_1",
    "delivery_address_line_1",
    "address_line_1",
    "line_1",
    "address1",
    "address_1",
  ],
  recipientCity: [
    "to_city",
    "recipient_city",
    "shipping_city",
    "ship_city",
    "ship_to_city",
    "delivery_city",
    "city",
  ],
  recipientState: [
    "to_state",
    "to_region",
    "recipient_state",
    "recipient_region",
    "recipient_province",
    "shipping_state",
    "shipping_region",
    "shipping_province",
    "ship_state",
    "ship_region",
    "ship_province",
    "ship_to_state",
    "ship_to_region",
    "ship_to_province",
    "delivery_state",
    "delivery_region",
    "delivery_province",
    "state",
    "province",
    "region",
  ],
  recipientZip: [
    "to_zip",
    "to_postcode",
    "to_postal_code",
    "recipient_zip",
    "recipient_postcode",
    "recipient_postal_code",
    "shipping_zip",
    "shipping_postcode",
    "shipping_postal_code",
    "ship_zip",
    "ship_postcode",
    "ship_postal_code",
    "ship_to_zip",
    "ship_to_postcode",
    "ship_to_postal_code",
    "delivery_zip",
    "delivery_postcode",
    "delivery_postal_code",
    "zip",
    "postcode",
    "postal_code",
  ],
  recipientCountry: [
    "to_country",
    "to_country_code",
    "recipient_country",
    "recipient_country_code",
    "shipping_country",
    "shipping_country_code",
    "ship_country",
    "ship_country_code",
    "ship_to_country",
    "ship_to_country_code",
    "delivery_country",
    "delivery_country_code",
    "country",
    "country_code",
  ],
  packageWeight: [
    "weight",
    "package_weight",
    "parcel_weight",
    "shipment_weight",
    "total_weight",
    "lineitem_weight",
    "line_item_weight",
    "item_weight",
    "grams",
    "weight_grams",
    "weight_g",
    "weight_kg",
    "weight_lb",
    "weight_oz",
    "package_weight_oz",
    "package_weight_lb",
    "package_weight_kg",
    "weight_in_oz",
    "weight_in_lb",
    "weight_in_kg",
    "weight_lbs",
  ],
  packageDims: [
    "dims",
    "dim",
    "dimensions",
    "package_dimensions",
    "parcel_dimensions",
    "shipment_dimensions",
    "item_dimensions",
    "box_dimensions",
  ],
};

const CSV_NAME_PART_ALIASES = {
  senderFirst: ["sender_first_name", "from_first_name", "origin_first_name", "ship_from_first_name"],
  senderLast: ["sender_last_name", "from_last_name", "origin_last_name", "ship_from_last_name"],
  recipientFirst: [
    "recipient_first_name",
    "to_first_name",
    "shipping_first_name",
    "ship_first_name",
    "ship_to_first_name",
    "customer_first_name",
    "buyer_first_name",
    "first_name",
    "firstname",
  ],
  recipientLast: [
    "recipient_last_name",
    "to_last_name",
    "shipping_last_name",
    "ship_last_name",
    "ship_to_last_name",
    "customer_last_name",
    "buyer_last_name",
    "last_name",
    "lastname",
  ],
};

const CSV_ADDRESS2_ALIASES = {
  senderStreet2: [
    "from_street_2",
    "from_address_2",
    "from_address_line_2",
    "from_line_2",
    "sender_address2",
    "sender_address_2",
    "sender_address_line_2",
    "ship_from_address2",
    "ship_from_address_2",
    "origin_address2",
    "origin_address_2",
    "location_address2",
    "location_address_2",
    "warehouse_address2",
    "warehouse_address_2",
  ],
  recipientStreet2: [
    "to_street_2",
    "to_address_2",
    "to_address_line_2",
    "to_line_2",
    "recipient_address2",
    "recipient_address_2",
    "recipient_address_line_2",
    "shipping_address2",
    "shipping_address_2",
    "shipping_address_line_2",
    "ship_address2",
    "ship_address_2",
    "ship_address_line_2",
    "ship_to_address2",
    "ship_to_address_2",
    "ship_to_address_line_2",
    "delivery_address2",
    "delivery_address_2",
    "delivery_address_line_2",
    "address_line_2",
    "line_2",
    "address2",
    "address_2",
  ],
};

const CSV_DIMENSION_PART_ALIASES = {
  length: ["length", "package_length", "parcel_length", "box_length", "item_length"],
  width: ["width", "package_width", "parcel_width", "box_width", "item_width"],
  height: ["height", "package_height", "parcel_height", "box_height", "item_height"],
};

const CSV_WAREHOUSE_ALIASES = [
  "location",
  "location_name",
  "warehouse",
  "warehouse_name",
  "fulfillment_location",
  "fulfillment_center",
  "ship_from_location",
  "origin_location",
];

const mockPlans = ["Starter", "Growth", "Scale", "Enterprise"];
const mockCompanySuffixes = [
  "Logistics",
  "Commerce",
  "Fulfillment",
  "Distribution",
  "Operations",
];
const mockFirstNames = [
  "Avery",
  "Jordan",
  "Taylor",
  "Casey",
  "Morgan",
  "Riley",
  "Parker",
  "Dakota",
];
const mockLastNames = [
  "Reed",
  "Carter",
  "Bennett",
  "Hayes",
  "Morgan",
  "Quinn",
  "Parker",
  "Sullivan",
];
const mockStreetNames = [
  "Rue du Port",
  "Avenue Centrale",
  "Kadeweg",
  "Via Mercato",
  "Carrer Major",
  "Boulevard Europe",
];
const mockCityRegions = [
  ["Amsterdam", "North Holland", "Netherlands"],
  ["Paris", "Ile-de-France", "France"],
  ["Cologne", "North Rhine-Westphalia", "Germany"],
  ["Milan", "Lombardy", "Italy"],
  ["Barcelona", "Catalonia", "Spain"],
  ["Vienna", "Vienna", "Austria"],
];

function historyStorageKey(userId) {
  return `shipr-history-${userId}`;
}

function historyServerCacheKey(userId) {
  return `shipr-history-server-cache-${userId}`;
}

function warehouseStorageKey(userId) {
  return `shipr-warehouses-${userId}`;
}

function wait(ms) {
  return new Promise((resolve) => {
    window.setTimeout(resolve, ms);
  });
}

async function fetchSupabaseHistoryRows(userId) {
  if (!supabaseClient || !userId) {
    return { data: [], error: new Error("Supabase unavailable") };
  }

  let lastError = null;
  for (let attempt = 0; attempt < HISTORY_FETCH_RETRY_DELAYS_MS.length; attempt += 1) {
    const delay = HISTORY_FETCH_RETRY_DELAYS_MS[attempt];
    if (delay > 0) {
      await wait(delay);
    }

    const { data, error } = await supabaseClient
      .from(HISTORY_TABLE)
      .select("id, created_at, service_type, quantity, total_price, payload")
      .eq("user_id", userId)
      .order("created_at", { ascending: false })
      .limit(HISTORY_LIMIT);

    if (!error && Array.isArray(data)) {
      return { data, error: null };
    }

    lastError = error || new Error("Unknown history fetch error");
  }

  return { data: [], error: lastError };
}

function formatHistoryDate(value) {
  if (!value) return "--";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "--";
  return date.toLocaleString([], {
    year: "numeric",
    month: "short",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function formatHistoryHeadlineParts(value) {
  if (!value) {
    return { dateText: "--", timeText: "" };
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return { dateText: "--", timeText: "" };
  }
  const day = String(date.getDate()).padStart(2, "0");
  const month = date
    .toLocaleString("en-GB", { month: "short" })
    .replace(".", "");
  const year = date.getFullYear();
  const hour = String(date.getHours()).padStart(2, "0");
  const minute = String(date.getMinutes()).padStart(2, "0");
  return {
    dateText: `${day} ${month}. ${year}`,
    timeText: `\u00A0-\u00A0${hour}:${minute}`,
  };
}

function formatMoney(value) {
  return `€${Number(value || 0).toFixed(2)}`;
}

function hashString(value) {
  const text = String(value || "");
  let hash = 2166136261;
  for (let i = 0; i < text.length; i += 1) {
    hash ^= text.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return hash >>> 0;
}

function pickBySeed(list, seed, offset = 0) {
  if (!Array.isArray(list) || list.length === 0) return "";
  return list[(seed + offset * 97) % list.length];
}

function toTitleCase(text) {
  return String(text || "")
    .split(/\s+/)
    .filter(Boolean)
    .map((chunk) => chunk.charAt(0).toUpperCase() + chunk.slice(1).toLowerCase())
    .join(" ");
}

function buildMockAccountProfile(user) {
  if (!user) return null;
  const email = (user.email || "").toLowerCase();
  const localPart = email.split("@")[0] || "shipr";
  const seed = hashString(user.id || email || localPart);
  const cleanedBase = localPart
    .replace(/[^a-z0-9]+/gi, " ")
    .trim();
  const companyBase = toTitleCase(cleanedBase || "Shipr");
  const companySuffix = pickBySeed(mockCompanySuffixes, seed, 1);
  const contactFirst = pickBySeed(mockFirstNames, seed, 2);
  const contactLast = pickBySeed(mockLastNames, seed, 3);
  const [city, stateCode, country] = pickBySeed(mockCityRegions, seed, 4);
  const streetNumber = 100 + (seed % 8800);
  const streetName = pickBySeed(mockStreetNames, seed, 5);
  const zip = String(1000 + ((seed >> 3) % 8999)).padStart(4, "0");
  const mobilePrefix = String(450 + (seed % 500)).padStart(3, "0");
  const lineA = String(10 + ((seed >> 5) % 90)).padStart(2, "0");
  const lineB = String(10 + ((seed >> 9) % 90)).padStart(2, "0");
  const lineC = String(10 + ((seed >> 13) % 90)).padStart(2, "0");
  const taxCore = String(100000000 + (seed % 900000000)).padStart(9, "0");
  const customerCore = String(100000 + (seed % 900000)).padStart(6, "0");

  return {
    companyName: `${companyBase} ${companySuffix}`.trim(),
    contactName: `${contactFirst} ${contactLast}`,
    contactEmail: email || `${localPart}@example.com`,
    contactPhone: `+32 ${mobilePrefix} ${lineA} ${lineB} ${lineC}`,
    billingAddress: `${streetNumber} ${streetName}, ${city}, ${stateCode} ${zip}, ${country}`,
    taxId: `EU-${taxCore}`,
    customerId: `CUST-${customerCore}`,
    plan: pickBySeed(mockPlans, seed, 6),
  };
}

function renderAccountProfile(user) {
  const profile = buildMockAccountProfile(user);
  const values = profile
    ? {
        companyName: profile.companyName,
        contactName: profile.contactName,
        contactEmail: profile.contactEmail,
        contactPhone: profile.contactPhone,
        billingAddress: profile.billingAddress,
        taxId: profile.taxId,
        customerId: profile.customerId,
        plan: profile.plan,
      }
    : {
        companyName: "--",
        contactName: "--",
        contactEmail: "--",
        contactPhone: "--",
        billingAddress: "--",
        taxId: "--",
        customerId: "--",
        plan: "--",
      };

  if (accountCompanyName) accountCompanyName.textContent = values.companyName;
  if (accountContactName) accountContactName.textContent = values.contactName;
  if (accountContactEmail) accountContactEmail.textContent = values.contactEmail;
  if (accountContactPhone) accountContactPhone.textContent = values.contactPhone;
  if (accountBillingAddress) accountBillingAddress.textContent = values.billingAddress;
  if (accountTaxId) accountTaxId.textContent = values.taxId;
  if (accountCustomerId) accountCustomerId.textContent = values.customerId;
  if (accountPlan) accountPlan.textContent = values.plan;

  if (invoiceCompany) invoiceCompany.textContent = values.companyName;
  if (invoiceContact) invoiceContact.textContent = values.contactName;
  if (invoiceEmail) invoiceEmail.textContent = values.contactEmail;
  if (invoicePhone) invoicePhone.textContent = values.contactPhone;
  if (invoiceAddress) invoiceAddress.textContent = values.billingAddress;
  if (invoiceTaxId) invoiceTaxId.textContent = values.taxId;
}

function generateWarehouseId() {
  if (window.crypto?.randomUUID) {
    return window.crypto.randomUUID();
  }
  return `wh-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}

function normalizeWarehouseRecord(record, fallback = {}) {
  const country = String(record?.country || fallback.country || "France").trim() || "France";
  return {
    id: String(record?.id || fallback.id || generateWarehouseId()),
    name: String(record?.name || fallback.name || "").trim(),
    senderName: String(
      record?.senderName ?? record?.sender_name ?? fallback.senderName ?? ""
    ).trim(),
    street: String(record?.street || fallback.street || "").trim(),
    city: String(record?.city || fallback.city || "").trim(),
    region: String(record?.region ?? record?.state ?? fallback.region ?? "").trim(),
    postalCode: String(
      record?.postalCode ?? record?.postal_code ?? fallback.postalCode ?? ""
    ).trim(),
    country,
    isDefault: Boolean(record?.isDefault ?? record?.is_default ?? fallback.isDefault),
  };
}

function normalizeWarehouseRecords(records) {
  const list = Array.isArray(records) ? records.slice(0, WAREHOUSE_MAX_COUNT) : [];
  const normalized = list.map((record, index) =>
    normalizeWarehouseRecord(record, {
      name: `Warehouse ${index + 1}`,
    })
  );
  if (!normalized.length) return [];
  let defaultIndex = normalized.findIndex((item) => item.isDefault);
  if (defaultIndex === -1) {
    defaultIndex = 0;
  }
  return normalized.map((item, index) => ({
    ...item,
    isDefault: index === defaultIndex,
  }));
}

function buildDefaultWarehouseRecords(user) {
  const profile = buildMockAccountProfile(user);
  const billing = parseAccountBillingAddress(profile?.billingAddress);
  return normalizeWarehouseRecords([
    {
      id: generateWarehouseId(),
      name: "Warehouse 1",
      senderName: String(profile?.companyName || profile?.contactName || "").trim(),
      street: billing.street,
      city: billing.city,
      region: billing.state,
      postalCode: billing.zip,
      country: billing.country || "France",
      isDefault: true,
    },
  ]);
}

function getDefaultWarehouseRecord() {
  return warehouseRecords.find((origin) => origin.isDefault) || warehouseRecords[0] || null;
}

function areSenderFieldsEmpty() {
  return (
    !String(inputMap.senderName?.value || "").trim() &&
    !String(inputMap.senderStreet?.value || "").trim() &&
    !String(inputMap.senderCity?.value || "").trim() &&
    !String(inputMap.senderState?.value || "").trim() &&
    !String(inputMap.senderZip?.value || "").trim()
  );
}

function applyWarehouseToSender(origin, options = {}) {
  const { announce = true } = options;
  if (!origin) return;
  if (inputMap.senderName) inputMap.senderName.value = origin.senderName || origin.name || "";
  if (inputMap.senderStreet) inputMap.senderStreet.value = origin.street || "";
  if (inputMap.senderCity) inputMap.senderCity.value = origin.city || "";
  if (inputMap.senderState) inputMap.senderState.value = origin.region || "";
  if (inputMap.senderZip) inputMap.senderZip.value = origin.postalCode || "";
  syncInfoState();
  updatePreview();
  if (announce) {
    setWarehouseStatus(`Applied ${origin.name || "warehouse"} to sender details.`, {
      tone: "success",
    });
  }
}

function maybeApplyDefaultWarehouseToSender() {
  if (!areSenderFieldsEmpty()) return;
  const defaultOrigin = getDefaultWarehouseRecord();
  if (!defaultOrigin) return;
  applyWarehouseToSender(defaultOrigin, { announce: false });
}

function setWarehouseStatus(message, options = {}) {
  const { tone = "muted" } = options;
  if (!warehouseStatus) return;
  warehouseStatus.textContent = String(message || "").trim();
  warehouseStatus.classList.remove("is-error", "is-success");
  if (tone === "error") {
    warehouseStatus.classList.add("is-error");
  } else if (tone === "success") {
    warehouseStatus.classList.add("is-success");
  }
}

function updateWarehouseControls() {
  const disabledBase = !currentUser || warehouseSaving;
  if (warehouseAddButton) {
    warehouseAddButton.disabled =
      disabledBase || warehouseRecords.length >= WAREHOUSE_MAX_COUNT;
  }
  if (warehouseSaveButton) {
    warehouseSaveButton.disabled =
      disabledBase || !warehouseDirty || warehouseRecords.length === 0;
    const label = warehouseSaveButton.querySelector("span");
    if (label) {
      label.textContent = warehouseSaving ? "Saving..." : "Save Origins";
    }
  }
}

function setWarehouseDirty(isDirty, options = {}) {
  const { announce = true } = options;
  warehouseDirty = Boolean(isDirty);
  updateWarehouseControls();
  if (warehouseDirty && announce) {
    setWarehouseStatus("Unsaved changes in shipping origins.");
  }
}

function createWarehouseField(label, fieldKey, value, options = {}) {
  const { wide = false, placeholder = "" } = options;
  const field = document.createElement("label");
  field.className = `warehouse-field${wide ? " is-wide" : ""}`;

  const title = document.createElement("span");
  title.className = "warehouse-field-label";
  title.textContent = label;

  const input = document.createElement("input");
  input.type = "text";
  input.value = value || "";
  input.placeholder = placeholder;
  input.dataset.warehouseField = fieldKey;

  field.appendChild(title);
  if (fieldKey === "country") {
    const wrap = document.createElement("div");
    wrap.className = "warehouse-country-input";
    const icon = document.createElement("span");
    icon.className = "warehouse-country-flag";
    icon.innerHTML = getCountryIcon(input.value);
    wrap.appendChild(icon);
    wrap.appendChild(input);
    field.appendChild(wrap);
  } else {
    field.appendChild(input);
  }
  return field;
}

function renderWarehouseList() {
  if (!warehouseList) return;
  warehouseList.innerHTML = "";
  warehouseList.classList.toggle("is-multi", warehouseRecords.length > 1);

  if (!currentUser) {
    const empty = document.createElement("div");
    empty.className = "warehouse-empty";
    empty.textContent = "Sign in to manage account shipping origins.";
    warehouseList.appendChild(empty);
    updateWarehouseControls();
    return;
  }

  if (!warehouseRecords.length) {
    const empty = document.createElement("div");
    empty.className = "warehouse-empty";
    empty.textContent = "No shipping origins configured yet.";
    warehouseList.appendChild(empty);
    updateWarehouseControls();
    return;
  }

  warehouseRecords.forEach((origin, index) => {
    const card = document.createElement("article");
    card.className = `warehouse-card${origin.isDefault ? " is-default" : ""}`;
    card.dataset.warehouseId = origin.id;
    if (warehouseEnteringIds.has(origin.id)) {
      card.classList.add("is-entering");
    }

    const head = document.createElement("div");
    head.className = "warehouse-card-head";

    const title = document.createElement("div");
    title.className = "warehouse-card-title";

    const name = document.createElement("span");
    name.className = "warehouse-card-name";
    name.textContent = origin.name || `Warehouse ${index + 1}`;

    title.appendChild(name);

    head.appendChild(title);

    const defaultToggle = document.createElement("button");
    defaultToggle.type = "button";
    defaultToggle.className = `warehouse-default-toggle${origin.isDefault ? " is-active" : ""}`;
    defaultToggle.dataset.warehouseAction = "default";
    defaultToggle.disabled = warehouseSaving;

    const defaultBox = document.createElement("span");
    defaultBox.className = "warehouse-default-box";
    defaultBox.setAttribute("aria-hidden", "true");

    const defaultText = document.createElement("span");
    defaultText.textContent = "Default";

    defaultToggle.appendChild(defaultBox);
    defaultToggle.appendChild(defaultText);
    head.appendChild(defaultToggle);

    const grid = document.createElement("div");
    grid.className = "warehouse-grid";
    grid.appendChild(
      createWarehouseField("Warehouse Name", "name", origin.name, {
        placeholder: `Warehouse ${index + 1}`,
      })
    );
    grid.appendChild(
      createWarehouseField("Sender Name", "senderName", origin.senderName, {
        placeholder: "Sender/Company name",
      })
    );
    grid.appendChild(createWarehouseField("Street", "street", origin.street, { wide: true }));
    grid.appendChild(createWarehouseField("City", "city", origin.city));
    grid.appendChild(createWarehouseField("Region", "region", origin.region));
    grid.appendChild(createWarehouseField("Postal Code", "postalCode", origin.postalCode));
    grid.appendChild(createWarehouseField("Country", "country", origin.country));

    const controls = document.createElement("div");
    controls.className = "warehouse-card-actions";

    const controlGroup = document.createElement("div");
    controlGroup.className = "warehouse-card-controls";

    const applyButton = document.createElement("button");
    applyButton.type = "button";
    applyButton.className = "btn btn-secondary btn-sm";
    applyButton.dataset.warehouseAction = "apply";
    applyButton.textContent = "Apply";
    applyButton.disabled = warehouseSaving;

    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.className = "btn btn-ghost btn-sm";
    removeButton.dataset.warehouseAction = "remove";
    removeButton.textContent = "Remove";
    removeButton.disabled = warehouseSaving || warehouseRecords.length <= 1;

    controlGroup.appendChild(applyButton);
    controlGroup.appendChild(removeButton);
    controls.appendChild(controlGroup);

    card.appendChild(head);
    card.appendChild(grid);
    card.appendChild(controls);
    warehouseList.appendChild(card);
  });

  const inputs = warehouseList.querySelectorAll("input[data-warehouse-field]");
  inputs.forEach((input) => {
    input.disabled = warehouseSaving;
  });

  if (warehouseEnteringIds.size > 0) {
    const enteringCards = Array.from(
      warehouseList.querySelectorAll(".warehouse-card.is-entering")
    );
    window.setTimeout(() => {
      enteringCards.forEach((card) => card.classList.remove("is-entering"));
      warehouseEnteringIds.clear();
    }, 620);
  }
  updateWarehouseControls();
}

function findWarehouseIndexById(warehouseId) {
  if (!warehouseId) return -1;
  return warehouseRecords.findIndex((origin) => origin.id === warehouseId);
}

function removeWarehouseRecord(warehouseId) {
  warehouseRecords = warehouseRecords.filter((origin) => origin.id !== warehouseId);
  warehouseRecords = normalizeWarehouseRecords(warehouseRecords);
  renderWarehouseList();
  setWarehouseDirty(true);
}

function removeWarehouseWithAnimation(warehouseId) {
  if (!warehouseList) {
    removeWarehouseRecord(warehouseId);
    return;
  }
  const card = warehouseList.querySelector(`[data-warehouse-id="${warehouseId}"]`);
  if (!card) {
    removeWarehouseRecord(warehouseId);
    return;
  }
  if (card.classList.contains("is-removing")) return;

  card.classList.add("is-removing");
  const controls = card.querySelectorAll("button, input");
  controls.forEach((element) => {
    element.disabled = true;
  });

  window.setTimeout(() => {
    removeWarehouseRecord(warehouseId);
  }, WAREHOUSE_REMOVE_ANIMATION_MS);
}

function buildNewWarehouseRecord() {
  const nextNumber = warehouseRecords.length + 1;
  const senderNameSeed = String(accountCompanyName?.textContent || "").trim();
  return normalizeWarehouseRecord({
    id: generateWarehouseId(),
    name: `Warehouse ${nextNumber}`,
    senderName: senderNameSeed === "--" ? "" : senderNameSeed,
    country: "France",
    isDefault: warehouseRecords.length === 0,
  });
}

function validateWarehouseRecords(records) {
  if (!Array.isArray(records) || !records.length) {
    return { ok: false, message: "Add at least one warehouse origin.", records: [] };
  }

  const normalized = normalizeWarehouseRecords(records);
  for (let i = 0; i < normalized.length; i += 1) {
    const origin = normalized[i];
    const prefix = origin.name || `Warehouse ${i + 1}`;
    if (!origin.name) {
      return { ok: false, message: `${prefix}: warehouse name is required.`, records: normalized };
    }
    if (!origin.senderName) {
      return { ok: false, message: `${prefix}: sender name is required.`, records: normalized };
    }
    if (!origin.street) {
      return { ok: false, message: `${prefix}: street is required.`, records: normalized };
    }
    if (!origin.city) {
      return { ok: false, message: `${prefix}: city is required.`, records: normalized };
    }
    if (!origin.postalCode) {
      return { ok: false, message: `${prefix}: postal code is required.`, records: normalized };
    }
    if (!origin.country) {
      return { ok: false, message: `${prefix}: country is required.`, records: normalized };
    }
  }

  return { ok: true, records: normalized };
}

function toWarehouseOriginsPayload(records) {
  return normalizeWarehouseRecords(records).map((origin) => ({
    id: origin.id,
    name: origin.name,
    senderName: origin.senderName,
    street: origin.street,
    city: origin.city,
    region: origin.region,
    postalCode: origin.postalCode,
    country: origin.country,
    isDefault: origin.isDefault,
  }));
}

async function fetchSupabaseWarehouseOrigins(userId) {
  if (!supabaseClient || !userId) {
    return { origins: [], error: new Error("Supabase unavailable") };
  }
  const { data, error } = await supabaseClient
    .from(ACCOUNT_SETTINGS_TABLE)
    .select("warehouse_origins")
    .eq("user_id", userId)
    .maybeSingle();

  if (error) {
    return { origins: [], error };
  }
  const origins = Array.isArray(data?.warehouse_origins) ? data.warehouse_origins : [];
  return { origins, error: null };
}

async function saveSupabaseWarehouseOrigins(userId, origins) {
  if (!supabaseClient || !userId) {
    return { origins: [], error: new Error("Supabase unavailable") };
  }
  const payload = toWarehouseOriginsPayload(origins);
  const { data, error } = await supabaseClient
    .from(ACCOUNT_SETTINGS_TABLE)
    .upsert(
      {
        user_id: userId,
        warehouse_origins: payload,
      },
      { onConflict: "user_id" }
    )
    .select("warehouse_origins")
    .single();

  if (error) {
    return { origins: [], error };
  }
  return {
    origins: Array.isArray(data?.warehouse_origins) ? data.warehouse_origins : payload,
    error: null,
  };
}

async function loadWarehouseSettings(options = {}) {
  const { quiet = false } = options;
  const requestToken = ++warehouseLoadRequestToken;
  const userId = currentUser?.id || null;

  if (!warehouseList) return;

  if (!userId) {
    warehouseRecords = [];
    warehouseDirty = false;
    warehouseSaving = false;
    renderWarehouseList();
    setWarehouseStatus("Sign in to manage shipping origins.");
    return;
  }

  if (!quiet) {
    setWarehouseStatus("Loading shipping origins...");
  }

  let nextRecords = [];
  let source = "default";

  if (supabaseClient) {
    const { origins, error } = await fetchSupabaseWarehouseOrigins(userId);
    if (requestToken !== warehouseLoadRequestToken || currentUser?.id !== userId) {
      return;
    }
    if (!error) {
      if (origins.length) {
        nextRecords = origins;
        source = "supabase";
      }
    } else {
      console.warn("Warehouse settings sync failed; falling back to local data.", error);
    }
  }

  if (!nextRecords.length) {
    const localRecords = loadLocalWarehouses(userId);
    if (localRecords.length) {
      nextRecords = localRecords;
      source = "local";
    }
  }

  if (!nextRecords.length) {
    nextRecords = buildDefaultWarehouseRecords(currentUser);
    source = "default";
  }

  warehouseRecords = normalizeWarehouseRecords(nextRecords);
  warehouseDirty = false;
  warehouseSaving = false;
  warehouseEnteringIds.clear();
  saveLocalWarehouses(userId, warehouseRecords);
  renderWarehouseList();

  if (source === "supabase") {
    setWarehouseStatus("Shipping origins synced from your account.");
  } else if (source === "local") {
    setWarehouseStatus("Showing browser-saved origins. Click Save Origins to sync.");
  } else {
    setWarehouseStatus("Add your shipping origin and click Save Origins.");
  }

  maybeApplyDefaultWarehouseToSender();
}

async function saveWarehouseSettings() {
  if (!currentUser || warehouseSaving) return;

  const validation = validateWarehouseRecords(warehouseRecords);
  if (!validation.ok) {
    setWarehouseStatus(validation.message, { tone: "error" });
    return;
  }

  warehouseRecords = validation.records;
  warehouseSaving = true;
  renderWarehouseList();
  setWarehouseStatus("Saving shipping origins...");

  const userId = currentUser.id;
  let savedToSupabase = false;
  let savedOrigins = toWarehouseOriginsPayload(warehouseRecords);

  if (supabaseClient) {
    const { origins, error } = await saveSupabaseWarehouseOrigins(userId, savedOrigins);
    if (!error) {
      savedOrigins = Array.isArray(origins) && origins.length ? origins : savedOrigins;
      savedToSupabase = true;
    } else {
      console.warn("Failed to sync shipping origins to account.", error);
    }
  }

  warehouseRecords = normalizeWarehouseRecords(savedOrigins);
  saveLocalWarehouses(userId, warehouseRecords);
  warehouseSaving = false;
  warehouseDirty = false;
  warehouseEnteringIds.clear();
  renderWarehouseList();

  if (savedToSupabase) {
    setWarehouseStatus("Shipping origins saved to your account.", { tone: "success" });
  } else {
    setWarehouseStatus("Saved locally. Account sync is currently unavailable.", {
      tone: "error",
    });
  }

  maybeApplyDefaultWarehouseToSender();
}

function resetWarehouseState() {
  warehouseLoadRequestToken += 1;
  warehouseRecords = [];
  warehouseDirty = false;
  warehouseSaving = false;
  warehouseEnteringIds.clear();
  renderWarehouseList();
  setWarehouseStatus("Sign in to manage shipping origins.");
}

function calculateRecordTotals(record) {
  const quantity = Math.max(1, Number(record?.quantity) || 1);
  let totalExVat = Number(record?.total_price);
  if (!Number.isFinite(totalExVat)) {
    const unitPrice = Number(record?.payload?.selection?.price) || 0;
    totalExVat = Number((unitPrice * quantity).toFixed(2));
  }
  const vatAmount = Number((totalExVat * VAT_RATE).toFixed(2));
  const totalIncVat = Number((totalExVat + vatAmount).toFixed(2));
  const unitExVat = Number((totalExVat / quantity).toFixed(2));
  const unitIncVat = Number((totalIncVat / quantity).toFixed(2));
  return { quantity, totalExVat, vatAmount, totalIncVat, unitExVat, unitIncVat };
}

function prefersReducedMotion() {
  return Boolean(window.matchMedia?.("(prefers-reduced-motion: reduce)")?.matches);
}

function clearTransitionFlags(element) {
  if (!element) return;
  element.classList.remove(
    "is-shell-enter-start",
    "is-shell-enter-active",
    "is-shell-exit",
    "is-app-transition-out",
    "is-app-transition-in"
  );
}

function transitionShellVisibility(isAuthed, options = {}) {
  const { animate = true } = options;
  if (!authGate || !appPage) return;

  const showElement = isAuthed ? appPage : authGate;
  const hideElement = isAuthed ? authGate : appPage;
  const shouldAnimate = animate && !prefersReducedMotion();

  clearTransitionFlags(showElement);
  clearTransitionFlags(hideElement);

  if (!shouldAnimate) {
    showElement.classList.remove("is-hidden");
    hideElement.classList.add("is-hidden");
    return;
  }

  const transitionToken = ++authShellTransitionToken;
  showElement.classList.remove("is-hidden");
  showElement.classList.add("is-shell-enter-start");
  hideElement.classList.add("is-shell-exit");

  requestAnimationFrame(() => {
    if (transitionToken !== authShellTransitionToken) return;
    showElement.classList.add("is-shell-enter-active");
  });

  window.setTimeout(() => {
    if (transitionToken !== authShellTransitionToken) return;
    hideElement.classList.add("is-hidden");
    hideElement.classList.remove("is-shell-exit");
    showElement.classList.remove("is-shell-enter-start", "is-shell-enter-active");
  }, SHELL_TRANSITION_IN_MS);
}

function runMainViewTransition(mutate, options = {}) {
  const { animate = true } = options;
  if (!appPage || appPage.classList.contains("is-hidden") || !animate || prefersReducedMotion()) {
    mutate();
    return;
  }

  const transitionToken = ++mainViewTransitionToken;
  appPage.classList.remove("is-app-transition-in");
  appPage.classList.add("is-app-transition-out");

  window.setTimeout(() => {
    if (transitionToken !== mainViewTransitionToken) return;
    mutate();
    appPage.classList.remove("is-app-transition-out");
    appPage.classList.add("is-app-transition-in");

    window.setTimeout(() => {
      if (transitionToken !== mainViewTransitionToken) return;
      appPage.classList.remove("is-app-transition-in");
    }, MAIN_VIEW_TRANSITION_IN_MS);
  }, MAIN_VIEW_TRANSITION_OUT_MS);
}

function setMainView(view, options = {}) {
  const { push = true, replace = false, animate = true } = options;
  const nextView =
    view === "account" || view === "reports" || view === "builder" ? view : "builder";
  const viewChanged = nextView !== currentMainView;

  const applyView = () => {
    currentMainView = nextView;
    if (builderPage) {
      builderPage.classList.toggle("is-hidden", nextView !== "builder");
    }
    if (accountPageSection) {
      accountPageSection.classList.toggle("is-hidden", nextView !== "account");
    }
    if (reportsPageSection) {
      reportsPageSection.classList.toggle("is-hidden", nextView !== "reports");
    }
    if (nextView !== "account") {
      setReceiptModalOpen(false);
    }
    if (nextView !== "builder") {
      setRestrictedGoodsModalOpen(false);
    }
    if (nextView === "reports") {
      renderReportsDashboard();
      ensureReportsGeoDataLoaded().then(() => {
        if (currentMainView === "reports") {
          renderReportsDashboard();
        }
      });
    }
  };

  runMainViewTransition(applyView, { animate: animate && viewChanged });

  if (push) {
    if (nextView === "account") {
      updateRoute({ view: "account" }, { replace });
      return;
    }
    if (nextView === "reports") {
      updateRoute({ view: "reports" }, { replace });
      return;
    }
    updateRoute({ view: "builder", step: state.step }, { replace });
  }
}

function setAccountPageVisible(visible, options = {}) {
  setMainView(visible ? "account" : "builder", options);
}

function setReportsPageVisible(visible, options = {}) {
  setMainView(visible ? "reports" : "builder", options);
}

function setReceiptModalOpen(open) {
  if (!receiptModal) return;
  receiptModal.classList.toggle("is-closed", !open);
}

function setRestrictedGoodsModalOpen(open) {
  if (!restrictedGoodsModal) return;
  restrictedGoodsModal.classList.toggle("is-closed", !open);
}

function loadLocalHistory(userId) {
  if (!userId) return [];
  try {
    const raw = localStorage.getItem(historyStorageKey(userId));
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (_error) {
    return [];
  }
}

function loadServerHistoryCache(userId) {
  if (!userId) return [];
  try {
    const raw = localStorage.getItem(historyServerCacheKey(userId));
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.slice(0, HISTORY_LIMIT) : [];
  } catch (_error) {
    return [];
  }
}

function saveServerHistoryCache(userId, records) {
  if (!userId) return;
  try {
    localStorage.setItem(
      historyServerCacheKey(userId),
      JSON.stringify(Array.isArray(records) ? records.slice(0, HISTORY_LIMIT) : [])
    );
  } catch (_error) {
    // Ignore cache write failures (quota/private mode)
  }
}

function saveLocalHistory(userId, records) {
  if (!userId) return;
  try {
    localStorage.setItem(historyStorageKey(userId), JSON.stringify(records));
  } catch (_error) {
    // Ignore local storage errors silently (quota/private mode)
  }
}

function loadLocalWarehouses(userId) {
  if (!userId) return [];
  try {
    const raw = localStorage.getItem(warehouseStorageKey(userId));
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (_error) {
    return [];
  }
}

function saveLocalWarehouses(userId, warehouses) {
  if (!userId) return;
  try {
    localStorage.setItem(warehouseStorageKey(userId), JSON.stringify(warehouses));
  } catch (_error) {
    // Ignore local storage errors silently (quota/private mode)
  }
}

function getHistorySignature(record) {
  const labels = record?.payload?.labels;
  if (!Array.isArray(labels) || !labels.length) return "";
  return labels.map((item) => item?.labelId || "").join("|");
}

function historyRecordHasLabels(record) {
  const labels = record?.payload?.labels;
  return Array.isArray(labels) && labels.length > 0;
}

function findFirstPreviewableHistoryIndex() {
  return historyRecords.findIndex((record) => historyRecordHasLabels(record));
}

function findHistoryRecordIndexById(recordId) {
  if (!recordId) return -1;
  return historyRecords.findIndex(
    (record) => record?.id === recordId && historyRecordHasLabels(record)
  );
}

function revokeAccountPdfUrls() {
  const urls = new Set();
  accountLabels.forEach((label) => {
    if (label.pdfUrl) {
      urls.add(label.pdfUrl);
    }
  });
  if (accountBatchPdfUrl) {
    urls.add(accountBatchPdfUrl);
  }
  urls.forEach((url) => URL.revokeObjectURL(url));
  accountBatchPdfUrl = "";
}

function resetAccountPreview() {
  revokeAccountPdfUrls();
  accountLabels = [];
  accountActiveRecord = null;
  accountActiveHistoryIndex = -1;
  accountActiveLabelIndex = 0;
  if (accountBatchPanel) {
    accountBatchPanel.classList.add("is-hidden");
  }
  if (accountBatchList) {
    accountBatchList.innerHTML = "";
  }
  if (accountBatchPreview) {
    accountBatchPreview.classList.add("is-single");
    accountBatchPreview.classList.add("is-empty");
  }
  if (accountPdfFrame) {
    accountPdfFrame.src = "";
  }
  if (accountPreviewMeta) {
    accountPreviewMeta.textContent = "Select a generation to preview labels and receipt details.";
  }
  if (accountDownloadPdf) {
    accountDownloadPdf.disabled = true;
  }
  if (openReceiptModalButton) {
    openReceiptModalButton.disabled = true;
  }
  if (receiptSummary) {
    receiptSummary.innerHTML = "";
  }
  if (receiptTableBody) {
    receiptTableBody.innerHTML = "";
  }
}

function renderAccountHistoryList() {
  if (!accountHistoryList) return;
  accountHistoryList.innerHTML = "";

  if (!currentUser) {
    const empty = document.createElement("div");
    empty.className = "account-history-empty";
    empty.textContent = "Sign in to view previous generations.";
    accountHistoryList.appendChild(empty);
    return;
  }

  if (!historyRecords.length) {
    const empty = document.createElement("div");
    empty.className = "account-history-empty";
    empty.textContent = "No generations yet.";
    accountHistoryList.appendChild(empty);
    return;
  }

  historyRecords.forEach((record, index) => {
    const totals = calculateRecordTotals(record);
    const item = document.createElement("button");
    item.type = "button";
    item.className = `account-record${index === accountActiveHistoryIndex ? " is-active" : ""}`;
    item.dataset.historyIndex = String(index);
    item.style.animationDelay = `${Math.min(index, 12) * 0.04}s`;

    const head = document.createElement("div");
    head.className = "account-record-head";

    const service = document.createElement("span");
    service.className = "account-record-service";
    service.textContent = record.service_type || "Label generation";

    const date = document.createElement("span");
    date.className = "account-record-date";
    const headlineParts = formatHistoryHeadlineParts(record.created_at);
    const dateMain = document.createElement("span");
    dateMain.className = "account-record-date-main";
    dateMain.textContent = headlineParts.dateText;
    const dateTime = document.createElement("span");
    dateTime.className = "account-record-date-time";
    dateTime.textContent = headlineParts.timeText;
    date.appendChild(dateMain);
    date.appendChild(dateTime);

    head.appendChild(date);
    head.appendChild(service);

    const meta = document.createElement("div");
    meta.className = "account-record-meta mono";
    meta.textContent = `${totals.quantity} labels • EX ${formatMoney(
      totals.totalExVat
    )} • INCL ${formatMoney(totals.totalIncVat)}`;

    item.appendChild(head);
    item.appendChild(meta);
    accountHistoryList.appendChild(item);
  });
}

function syncAccountHistorySelection() {
  if (!accountHistoryList) return;
  const items = accountHistoryList.querySelectorAll("[data-history-index]");
  items.forEach((item) => {
    const index = Number(item.getAttribute("data-history-index"));
    item.classList.toggle("is-active", index === accountActiveHistoryIndex);
  });
}

function setAccountHistoryStatus(message) {
  if (!accountHistoryStatus) return;
  const text = String(message || "").trim();
  accountHistoryStatus.textContent = text;
  accountHistoryStatus.classList.toggle("is-hidden", !text);
}

async function loadGenerationHistory(options = {}) {
  const { preferLatest = false } = options;
  if (!accountHistoryStatus) return;
  const requestToken = ++historyLoadRequestToken;
  const requestedUserId = currentUser?.id || null;
  const priorRecords = Array.isArray(historyRecords) ? [...historyRecords] : [];
  const isStale = () =>
    requestToken !== historyLoadRequestToken || currentUser?.id !== requestedUserId;

  const applyLoadedRecords = (records, _loadedStatus, emptyStatus) => {
    if (isStale()) return;
    historyRecords = Array.isArray(records) ? records : [];
    setAccountHistoryStatus(historyRecords.length ? "" : emptyStatus);
    renderAccountHistoryList();

    if (!historyRecords.length) {
      resetAccountPreview();
      refreshReportsIfVisible();
      return;
    }

    let selectionIndex = previousRecordId
      ? findHistoryRecordIndexById(previousRecordId)
      : -1;
    if (selectionIndex === -1) {
      selectionIndex = findFirstPreviewableHistoryIndex();
    }

    if (selectionIndex === -1) {
      resetAccountPreview();
      setAccountHistoryStatus("No preview-ready generations yet.");
      refreshReportsIfVisible();
      return;
    }

    selectAccountRecord(selectionIndex);
    refreshReportsIfVisible();
  };

  if (!currentUser) {
    historyRecords = [];
    historyStore = "supabase";
    setAccountHistoryStatus("Sign in to view previous generations.");
    renderAccountHistoryList();
    resetAccountPreview();
    renderReportsDashboard();
    return;
  }

  setAccountHistoryStatus("Loading previous generations...");
  const previousRecordId = !preferLatest ? accountActiveRecord?.id : null;

  if (supabaseClient) {
    const { data, error } = await fetchSupabaseHistoryRows(requestedUserId);

    if (isStale()) return;
    if (!error && Array.isArray(data)) {
      historyStore = "supabase";
      saveServerHistoryCache(requestedUserId, data);
      applyLoadedRecords(
        data,
        "Select a generation to view PDFs and receipt details.",
        "No previous generations yet."
      );
      return;
    }

    console.warn("History sync failed; falling back to cached/local history.", error);
    if (isStale()) return;
    const serverCachedRecords = loadServerHistoryCache(requestedUserId);
    if (serverCachedRecords.length) {
      historyStore = "supabase-cache";
      applyLoadedRecords(
        serverCachedRecords,
        "Sync delayed. Showing last synced history.",
        "No previous generations yet."
      );
      return;
    }
  }

  if (isStale()) return;

  const localRecords = loadLocalHistory(requestedUserId);
  if (isStale()) return;
  if (localRecords.length) {
    historyStore = "local";
    applyLoadedRecords(
      localRecords,
      "Showing browser-saved history for this account.",
      "No browser history saved for this account yet."
    );
    return;
  }

  if (isStale()) return;
  if (priorRecords.length) {
    historyStore = "memory";
    applyLoadedRecords(
      priorRecords,
      "Sync delayed. Showing previously loaded history.",
      "No previous generations yet."
    );
    return;
  }

  if (isStale()) return;
  historyStore = "local";
  applyLoadedRecords(
    [],
    "Showing browser-saved history for this account.",
    "Could not load history right now. Please refresh in a moment."
  );
}

function buildHistoryRecord() {
  if (!currentUser || !state.labels.length) return null;
  const payloadLabels = state.labels.map((label) => ({
    index: label.index,
    labelId: label.labelId,
    trackingId: label.trackingId,
    data: label.data,
  }));
  const customsRequired = requiresCustomsDeclaration();
  const customsPayload = customsRequired
    ? {
        summaryDescription: state.customs.summaryDescription,
        shipmentType: state.customs.shipmentType,
        senderReference: state.customs.senderReference,
        importerReference: state.customs.importerReference,
        items: buildCustomsItemsPayload(),
      }
    : null;
  return {
    service_type: state.selection.type,
    quantity: payloadLabels.length,
    total_price: Number((state.selection.price * payloadLabels.length).toFixed(2)),
    payload: {
      selection: { ...state.selection },
      labels: payloadLabels,
      customs: customsPayload,
    },
  };
}

async function persistGenerationHistory() {
  const record = buildHistoryRecord();
  if (!record) return;

  const latestSignature = getHistorySignature(historyRecords[0]);
  const nextSignature = getHistorySignature(record);
  if (latestSignature && latestSignature === nextSignature) {
    refreshReportsIfVisible();
    return;
  }

  if (supabaseClient && currentUser?.id) {
    const { data, error } = await supabaseClient
      .from(HISTORY_TABLE)
      .insert({
        user_id: currentUser.id,
        service_type: record.service_type,
        quantity: record.quantity,
        total_price: record.total_price,
        payload: record.payload,
      })
      .select("id, created_at, service_type, quantity, total_price, payload")
      .single();

    if (!error && data) {
      historyStore = "supabase";
      historyRecords = [data, ...historyRecords].slice(0, HISTORY_LIMIT);
      saveServerHistoryCache(currentUser.id, historyRecords);
      setAccountHistoryStatus("");
      renderAccountHistoryList();
      refreshReportsIfVisible();
      return;
    }

    console.warn("History insert failed; saving this generation locally.", error);
    historyStore = "local";
    setAccountHistoryStatus("");
  }

  const localRecord = {
    id: `local-${Date.now()}`,
    created_at: new Date().toISOString(),
    ...record,
  };
  historyRecords = [localRecord, ...historyRecords].slice(0, HISTORY_LIMIT);
  saveLocalHistory(currentUser.id, historyRecords);
  renderAccountHistoryList();
  refreshReportsIfVisible();
}

function renderAccountBatchList() {
  if (!accountBatchPanel || !accountBatchList) return;
  if (accountLabels.length <= 1) {
    accountBatchPanel.classList.add("is-hidden");
    accountBatchList.innerHTML = "";
    if (accountBatchPreview) {
      accountBatchPreview.classList.add("is-single");
    }
    return;
  }

  accountBatchPanel.classList.remove("is-hidden");
  if (accountBatchPreview) {
    accountBatchPreview.classList.remove("is-single");
  }

  accountBatchList.innerHTML = "";
  accountLabels.forEach((label, index) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `batch-item${index === accountActiveLabelIndex ? " is-active" : ""}`;
    button.dataset.accountLabelIndex = String(index);
    button.innerHTML = `
      <div class="batch-index">Label ${label.index}</div>
      <div class="batch-meta mono">${label.trackingId}</div>
      <div class="batch-meta mono">${label.labelId}</div>
    `;
    accountBatchList.appendChild(button);
  });
}

function selectAccountLabel(index) {
  if (!accountLabels[index]) return;
  accountActiveLabelIndex = index;
  const activeLabel = accountLabels[index];
  if (accountPdfFrame) {
    accountPdfFrame.src = buildPdfViewerUrl(activeLabel.pdfUrl, activeLabel.pdfPage || index + 1);
  }

  if (accountBatchList) {
    Array.from(accountBatchList.children).forEach((child, idx) => {
      child.classList.toggle("is-active", idx === index);
    });
  }
}

function renderReceiptDetails(record) {
  if (!receiptSummary || !receiptTableBody || !record) return;
  const totals = calculateRecordTotals(record);
  const serviceType = record.payload?.selection?.type || record.service_type || "--";
  const labels = accountLabels.length ? accountLabels : record.payload?.labels || [];

  receiptSummary.innerHTML = "";
  const summaryRows = [
    ["Service", serviceType],
    ["Date", formatHistoryDate(record.created_at)],
    ["Order", record.id || "--"],
    ["Quantity", String(totals.quantity)],
    ["Subtotal (ex VAT)", formatMoney(totals.totalExVat)],
    [`VAT (${Math.round(VAT_RATE * 100)}%)`, formatMoney(totals.vatAmount)],
    ["Total (incl VAT)", formatMoney(totals.totalIncVat)],
    ["Unit (ex VAT)", formatMoney(totals.unitExVat)],
    ["Unit (incl VAT)", formatMoney(totals.unitIncVat)],
  ];

  summaryRows.forEach(([label, value]) => {
    const row = document.createElement("div");
    row.className = "receipt-summary-row";
    const key = document.createElement("span");
    key.textContent = label;
    const val = document.createElement("strong");
    val.className = "mono";
    val.textContent = value;
    row.appendChild(key);
    row.appendChild(val);
    receiptSummary.appendChild(row);
  });

  receiptTableBody.innerHTML = "";
  labels.forEach((label) => {
    const tr = document.createElement("tr");
    const columns = [
      label.labelId || "--",
      label.trackingId || "--",
      formatMoney(totals.unitExVat),
      formatMoney(totals.unitIncVat),
    ];
    columns.forEach((text) => {
      const td = document.createElement("td");
      td.textContent = text;
      tr.appendChild(td);
    });
    receiptTableBody.appendChild(tr);
  });
}

function triggerFileDownload(url, filename) {
  if (!url) return;
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.target = "_blank";
  anchor.rel = "noopener";
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
}

function downloadActiveAccountLabelPdf() {
  if (!accountBatchPdfUrl) return;
  const filenameBase = accountActiveRecord?.id || "label-batch";
  triggerFileDownload(accountBatchPdfUrl, `${filenameBase}.pdf`);
}

function downloadReceiptPdfFile() {
  if (!accountActiveRecord) return;
  const totals = calculateRecordTotals(accountActiveRecord);
  const serviceType =
    accountActiveRecord.payload?.selection?.type ||
    accountActiveRecord.service_type ||
    "Label generation";
  const labels = accountLabels.length
    ? accountLabels
    : accountActiveRecord.payload?.labels || [];

  const lines = [
    "SHIPR RECEIPT",
    `Order: ${accountActiveRecord.id || "--"}`,
    `Date: ${formatHistoryDate(accountActiveRecord.created_at)}`,
    `Service: ${serviceType}`,
    `Quantity: ${totals.quantity}`,
    `Subtotal (EX VAT): ${formatMoney(totals.totalExVat)}`,
    `VAT (${Math.round(VAT_RATE * 100)}%): ${formatMoney(totals.vatAmount)}`,
    `Total (INCL VAT): ${formatMoney(totals.totalIncVat)}`,
    "",
    "LABEL BREAKDOWN",
  ];

  labels.forEach((label, index) => {
    lines.push(
      `${index + 1}. ${label.labelId || "--"} | ${label.trackingId || "--"} | EX ${formatMoney(
        totals.unitExVat
      )} | INCL ${formatMoney(totals.unitIncVat)}`
    );
  });

  const blob = buildPdf(lines, {
    pageWidth: 612,
    pageHeight: 792,
    marginX: 40,
    marginTop: 48,
    marginBottom: 48,
    fontSize: 11,
    lineHeight: 14,
  });
  const url = URL.createObjectURL(blob);
  triggerFileDownload(url, `receipt-${accountActiveRecord.id || "label-order"}.pdf`);
  setTimeout(() => URL.revokeObjectURL(url), 0);
}

function selectAccountRecord(index) {
  const record = historyRecords[index];
  const labels = record?.payload?.labels;
  if (!record || !Array.isArray(labels) || !labels.length) return;

  accountActiveRecord = record;
  accountActiveHistoryIndex = index;
  accountActiveLabelIndex = 0;
  revokeAccountPdfUrls();

  const serviceType = record.payload?.selection?.type || record.service_type || "Label";
  accountLabels = labels.map((saved, idx) => {
    const label = {
      index: saved.index || idx + 1,
      labelId: saved.labelId || generateLabelId(),
      trackingId: saved.trackingId || generateTracking(),
      data: saved.data || {},
    };
    return {
      ...label,
      pdfUrl: "",
      pdfPage: idx + 1,
    };
  });
  const batchBlob = createLabelBatchPdfBlob(accountLabels, serviceType);
  accountBatchPdfUrl = URL.createObjectURL(batchBlob);
  accountLabels = accountLabels.map((label) => ({
    ...label,
    pdfUrl: accountBatchPdfUrl,
  }));

  const totals = calculateRecordTotals(record);
  if (accountPreviewMeta) {
    accountPreviewMeta.textContent = `${serviceType} • ${totals.quantity} labels • EX ${formatMoney(
      totals.totalExVat
    )} • VAT ${formatMoney(totals.vatAmount)} • INCL ${formatMoney(totals.totalIncVat)}`;
  }

  if (accountDownloadPdf) {
    accountDownloadPdf.disabled = false;
  }
  if (openReceiptModalButton) {
    openReceiptModalButton.disabled = false;
  }
  if (accountBatchPreview) {
    accountBatchPreview.classList.remove("is-empty");
  }

  syncAccountHistorySelection();
  renderAccountBatchList();
  selectAccountLabel(0);
  renderReceiptDetails(record);
}

function toDayKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(
    date.getDate()
  ).padStart(2, "0")}`;
}

function toDayStart(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return null;
  date.setHours(0, 0, 0, 0);
  return date;
}

function formatReportAxisDate(date) {
  if (!date) return "--";
  return date.toLocaleDateString([], { month: "numeric", day: "numeric" });
}

function formatReportTooltipDate(date) {
  if (!date) return "--";
  return date.toLocaleDateString([], { month: "short", day: "numeric" });
}

function formatDateInputValue(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(
    date.getDate()
  ).padStart(2, "0")}`;
}

function normalizeCountryLabel(value) {
  const raw = String(value || "").trim();
  if (!raw) return "Unknown";
  const lowered = raw.toLowerCase();
  if (DOMESTIC_COUNTRY_ALIASES.has(lowered)) return "Domestic";
  return toTitleCase(raw);
}

function isDomesticCountry(value) {
  const normalized = String(value || "").trim().toLowerCase();
  return DOMESTIC_COUNTRY_ALIASES.has(normalized);
}

function resolveDomesticRegion(labelData) {
  const stateText = String(labelData?.recipientState || "").trim().toLowerCase();
  if (stateText.includes("flanders") || stateText.includes("vlaanderen")) return "North Region";
  if (stateText.includes("wallonia") || stateText.includes("wallonie")) return "South Region";
  if (stateText.includes("brussels") || stateText.includes("bruxelles")) {
    return "Capital Region";
  }

  const zipDigits = String(labelData?.recipientZip || "").replace(/\D/g, "");
  if (!zipDigits) return "North Region";
  const prefix = Number(zipDigits.slice(0, 1));
  if (prefix === 1) return "Capital Region";
  if ([2, 3, 8, 9].includes(prefix)) return "North Region";
  if ([4, 5, 6, 7].includes(prefix)) return "South Region";
  return "North Region";
}

function getRecordCreatedDay(record) {
  return toDayStart(record?.created_at);
}

function getRecordServiceType(record) {
  return record?.payload?.selection?.type || record?.service_type || "Unknown Service";
}

function getRecordLabels(record) {
  const labels = record?.payload?.labels;
  if (!Array.isArray(labels)) return [];
  return labels.filter((label) => label && typeof label === "object");
}

function getRetailRateForService(serviceType, fallbackUnitPrice) {
  const baseRate = RETAIL_BASE_RATE[serviceType];
  if (Number.isFinite(baseRate)) return baseRate;
  const fallback = Number(fallbackUnitPrice) || 0;
  return Number((fallback * 1.18).toFixed(2));
}

function getServiceColor(serviceName, index = 0) {
  const mapped = REPORT_SERVICE_COLOR_MAP[serviceName];
  if (mapped) return mapped;
  return REPORT_SERVICE_FALLBACK_COLORS[index % REPORT_SERVICE_FALLBACK_COLORS.length];
}

function parseZipPrefix(value) {
  const digits = String(value || "").replace(/\D/g, "");
  if (!digits.length) return null;
  return Number(digits.slice(0, 1));
}

function estimateZoneFromLabelData(labelData) {
  const senderPrefix = parseZipPrefix(labelData?.senderZip);
  const recipientPrefix = parseZipPrefix(labelData?.recipientZip);
  if (!Number.isFinite(senderPrefix) || !Number.isFinite(recipientPrefix)) {
    return 4;
  }
  const distance = Math.abs(senderPrefix - recipientPrefix);
  return Math.max(1, Math.min(9, distance + 1));
}

function incrementReportsCounter(store, key, count, spend) {
  const safeKey = key || "Unknown";
  if (!store.has(safeKey)) {
    store.set(safeKey, { count: 0, spend: 0 });
  }
  const target = store.get(safeKey);
  target.count += count;
  target.spend += spend;
}

function round2(value) {
  return Number((Number(value) || 0).toFixed(2));
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formatSignedMoney(value) {
  const amount = round2(value);
  const absolute = formatMoney(Math.abs(amount));
  if (amount > 0) return `+${absolute}`;
  if (amount < 0) return `-${absolute}`;
  return absolute;
}

function setReportsCustomRangeVisible(visible) {
  if (!reportsCustomRange) return;
  reportsCustomRange.classList.toggle("is-hidden", !visible);
}

function getReportsRangeSettings() {
  const today = toDayStart(new Date());
  if (reportsRangeMode === "custom") {
    const start = toDayStart(reportsCustomStart);
    const end = toDayStart(reportsCustomEnd);
    if (start && end) {
      const boundedEnd = end > today ? today : end;
      if (start <= boundedEnd) {
        const daySpan = Math.max(1, Math.round((boundedEnd - start) / 86400000) + 1);
        return {
          mode: "custom",
          start,
          end: boundedEnd,
          days: daySpan,
          rangeToken: "custom",
        };
      }
    }
  }

  const safeRange = [7, 30, 90, 365].includes(Number(reportsRangeDays))
    ? Number(reportsRangeDays)
    : 30;
  const start = new Date(today);
  start.setDate(start.getDate() - (safeRange - 1));
  return {
    mode: "preset",
    start,
    end: today,
    days: safeRange,
    rangeToken: safeRange,
  };
}

function applyMockCountriesIfEmpty(rows) {
  const source = Array.isArray(rows) ? rows.map((item) => ({ ...item })) : [];
  const byName = new Map(source.map((item) => [item.name, item]));
  REPORT_MOCK_COUNTRIES.forEach((mockItem) => {
    if (!byName.has(mockItem.name)) {
      byName.set(mockItem.name, { ...mockItem });
    }
  });
  return Array.from(byName.values())
    .sort((a, b) => b.count - a.count || b.spend - a.spend)
    .slice(0, 8);
}

function applyMockRegionsIfEmpty(rows) {
  if (Array.isArray(rows) && rows.length > 0) return rows;
  return REPORT_MOCK_REGIONS.map((item) => ({ ...item }));
}

function buildReportsModel() {
  const rangeSettings = getReportsRangeSettings();
  const { start, end, days } = rangeSettings;

  const buckets = [];
  const bucketByKey = new Map();
  for (let offset = 0; offset < days; offset += 1) {
    const date = new Date(start);
    date.setDate(start.getDate() + offset);
    const key = toDayKey(date);
    const bucket = {
      key,
      date,
      payments: 0,
      shipments: 0,
      savings: 0,
      averageCost: 0,
    };
    bucketByKey.set(key, bucket);
    buckets.push(bucket);
  }

  const serviceTotals = new Map();
  const regionTotals = new Map();
  const countryTotals = new Map();
  const zoneTotals = Array.from({ length: 9 }, () => 0);

  let totalShippingCosts = 0;
  let totalShipments = 0;
  let totalSavings = 0;
  let pendingReturnsCount = 0;
  let pendingReturnsAmount = 0;
  let pendingRefundsCount = 0;
  let pendingRefundsAmount = 0;
  let adjustmentEvents = 0;
  let carrierAdjustments = 0;
  let deliveryIssues = 0;
  let domesticCount = 0;
  let domesticSpend = 0;
  let internationalCount = 0;
  let internationalSpend = 0;

  historyRecords.forEach((record) => {
    const day = getRecordCreatedDay(record);
    if (!day || day < start || day > end) return;

    const key = toDayKey(day);
    const bucket = bucketByKey.get(key);
    if (!bucket) return;

    const totals = calculateRecordTotals(record);
    const quantity = Math.max(1, Number(totals.quantity) || 1);
    const totalExVat = round2(totals.totalExVat);
    const unitExVat = quantity > 0 ? round2(totalExVat / quantity) : 0;
    const serviceType = getRecordServiceType(record);
    const retailRate = getRetailRateForService(serviceType, unitExVat);
    const savingsAmount = round2(Math.max(0, (retailRate - unitExVat) * quantity));
    const signature = hashString(
      `${record?.id || "unknown"}|${record?.created_at || "0"}|${serviceType}`
    );
    const returnFlags = signature % 13 === 0 ? Math.max(1, Math.round(quantity * 0.34)) : 0;
    const refundFlags = signature % 17 === 0 ? Math.max(1, Math.round(quantity * 0.2)) : 0;
    const issueFlags = signature % 11 === 0 ? Math.max(1, Math.round(quantity * 0.15)) : 0;
    const hasAdjustment = signature % 9 === 0;
    const adjustmentAmount = hasAdjustment
      ? round2((signature % 2 === 0 ? 1 : -1) * (0.3 + (signature % 4) * 0.12))
      : 0;

    bucket.payments = round2(bucket.payments + totalExVat);
    bucket.shipments += quantity;
    bucket.savings = round2(bucket.savings + savingsAmount);

    totalShippingCosts = round2(totalShippingCosts + totalExVat);
    totalShipments += quantity;
    totalSavings = round2(totalSavings + savingsAmount);

    pendingReturnsCount += returnFlags;
    pendingReturnsAmount = round2(pendingReturnsAmount + returnFlags * unitExVat);
    pendingRefundsCount += refundFlags;
    pendingRefundsAmount = round2(pendingRefundsAmount + refundFlags * unitExVat);
    deliveryIssues += issueFlags;

    if (hasAdjustment) {
      adjustmentEvents += 1;
      carrierAdjustments = round2(carrierAdjustments + adjustmentAmount);
    }

    incrementReportsCounter(serviceTotals, serviceType, quantity, totalExVat);

    const labels = getRecordLabels(record);
    const distributedCount = labels.length || quantity;
    const distributedCost = distributedCount ? round2(totalExVat / distributedCount) : 0;
    const sourceItems =
      labels.length > 0
        ? labels
        : [{ data: record?.payload?.labels?.[0]?.data || record?.payload?.info || {} }];

    sourceItems.forEach((labelItem) => {
      const labelData = labelItem?.data || {};
      const recipientCountry = normalizeCountryLabel(labelData.recipientCountry);
      const domestic = isDomesticCountry(recipientCountry);
      const zone = estimateZoneFromLabelData(labelData);

      zoneTotals[zone - 1] += 1;

      if (domestic) {
        domesticCount += 1;
        domesticSpend = round2(domesticSpend + distributedCost);
        const region = resolveDomesticRegion(labelData);
        incrementReportsCounter(regionTotals, region, 1, distributedCost);
      } else {
        internationalCount += 1;
        internationalSpend = round2(internationalSpend + distributedCost);
        incrementReportsCounter(countryTotals, recipientCountry, 1, distributedCost);
      }
    });
  });

  buckets.forEach((bucket) => {
    bucket.averageCost = bucket.shipments > 0 ? round2(bucket.payments / bucket.shipments) : 0;
  });

  const serviceRows = Array.from(serviceTotals.entries())
    .map(([name, stats], index) => ({
      name,
      count: stats.count,
      spend: round2(stats.spend),
      color: getServiceColor(name, index),
    }))
    .sort((a, b) => b.count - a.count || b.spend - a.spend);

  const topRegions = Array.from(regionTotals.entries())
    .map(([name, stats]) => ({ name, count: stats.count, spend: round2(stats.spend) }))
    .sort((a, b) => b.count - a.count || b.spend - a.spend)
    .slice(0, 8);

  const topCountriesRaw = Array.from(countryTotals.entries())
    .map(([name, stats]) => ({ name, count: stats.count, spend: round2(stats.spend) }))
    .sort((a, b) => b.count - a.count || b.spend - a.spend)
    .slice(0, 8);

  const topCountries = applyMockCountriesIfEmpty(topCountriesRaw);
  const topRegionsWithFallback = applyMockRegionsIfEmpty(topRegions);

  const totalPayments = round2(totalShippingCosts + carrierAdjustments);
  const deliveryIssueRate = totalShipments ? (deliveryIssues / totalShipments) * 100 : 0;
  const safeDomesticCount =
    domesticCount || topRegionsWithFallback.reduce((sum, row) => sum + row.count, 0);
  const safeDomesticSpend =
    domesticSpend || topRegionsWithFallback.reduce((sum, row) => sum + row.spend, 0);
  const safeInternationalCount =
    internationalCount || topCountries.reduce((sum, row) => sum + row.count, 0);
  const safeInternationalSpend =
    internationalSpend || topCountries.reduce((sum, row) => sum + row.spend, 0);

  return {
    rangeDays: days,
    rangeMode: rangeSettings.mode,
    rangeToken: rangeSettings.rangeToken,
    rangeStart: start,
    rangeEnd: end,
    buckets,
    serviceRows,
    zoneTotals,
    topRegions: topRegionsWithFallback,
    topCountries,
    totalSavings,
    totalShippingCosts,
    totalPayments,
    totalShipments,
    pendingReturnsCount,
    pendingReturnsAmount,
    pendingRefundsCount,
    pendingRefundsAmount,
    accountBalance: 0,
    carrierAdjustments,
    adjustmentEvents,
    avgCost: totalShipments ? round2(totalShippingCosts / totalShipments) : 0,
    avgDomestic: safeDomesticCount ? round2(safeDomesticSpend / safeDomesticCount) : 0,
    avgDomesticCount: safeDomesticCount,
    avgInternational: safeInternationalCount
      ? round2(safeInternationalSpend / safeInternationalCount)
      : 0,
    avgInternationalCount: safeInternationalCount,
    newShipments: totalShipments,
    deliveryIssues,
    deliveryIssueRate,
  };
}

function buildReportTickIndices(length, maxTicks = 7) {
  if (length <= 0) return [];
  if (length <= maxTicks) return Array.from({ length }, (_, index) => index);
  const ticks = [];
  const step = (length - 1) / (maxTicks - 1);
  for (let i = 0; i < maxTicks; i += 1) {
    ticks.push(Math.round(i * step));
  }
  return Array.from(new Set(ticks)).sort((a, b) => a - b);
}

function renderReportAxis(axisEl, points) {
  if (!axisEl) return;
  const tickIndices = buildReportTickIndices(points.length, 7);
  axisEl.style.setProperty("--ticks", String(Math.max(1, tickIndices.length)));
  axisEl.innerHTML = tickIndices
    .map(
      (index) =>
        `<span class="reports-axis-label">${formatReportAxisDate(points[index]?.date)}</span>`
    )
    .join("");
}

function renderReportLineChart({
  svgEl,
  axisEl,
  tooltipEl,
  points,
  valueKey,
  valueFormatter,
  tooltipLabel,
}) {
  if (!svgEl || !axisEl || !Array.isArray(points)) return;
  const safePoints = points.length ? points : [{ date: new Date(), [valueKey]: 0 }];
  const values = safePoints.map((point) => Number(point[valueKey]) || 0);
  const width = 900;
  const height = 220;
  const paddingX = 24;
  const paddingTop = 18;
  const paddingBottom = 28;
  const innerWidth = width - paddingX * 2;
  const innerHeight = height - paddingTop - paddingBottom;
  const maxValue = Math.max(1, ...values);
  const baselineY = paddingTop + innerHeight;
  const stepX = safePoints.length > 1 ? innerWidth / (safePoints.length - 1) : 0;

  const coords = values.map((value, index) => {
    const x = paddingX + index * stepX;
    const y = baselineY - (value / maxValue) * innerHeight;
    return { x, y };
  });

  const linePath = coords
    .map((coord, index) => `${index === 0 ? "M" : "L"} ${coord.x.toFixed(2)} ${coord.y.toFixed(2)}`)
    .join(" ");

  const fillPath = `${linePath} L ${coords[coords.length - 1].x.toFixed(2)} ${baselineY.toFixed(
    2
  )} L ${coords[0].x.toFixed(2)} ${baselineY.toFixed(2)} Z`;

  const gridCount = 4;
  const gridLines = Array.from({ length: gridCount }, (_, index) => {
    const ratio = index / (gridCount - 1);
    const y = paddingTop + ratio * innerHeight;
    return `<line class="reports-line-grid" x1="${paddingX}" y1="${y.toFixed(
      2
    )}" x2="${(width - paddingX).toFixed(2)}" y2="${y.toFixed(2)}"></line>`;
  }).join("");

  const dots = coords
    .map(
      (coord, index) =>
        `<circle class="reports-line-dot" data-chart-index="${index}" cx="${coord.x.toFixed(
          2
        )}" cy="${coord.y.toFixed(2)}" r="3.6" tabindex="0" role="button" aria-label="${tooltipLabel} ${formatReportTooltipDate(
          safePoints[index].date
        )}: ${valueFormatter(values[index])}"></circle>`
    )
    .join("");

  svgEl.innerHTML = `
    ${gridLines}
    <path class="reports-line-fill" d="${fillPath}"></path>
    <path class="reports-line-path" d="${linePath}"></path>
    ${dots}
  `;

  renderReportAxis(axisEl, safePoints);

  if (!tooltipEl) return;
  tooltipEl.classList.add("is-hidden");
  const pointEls = Array.from(svgEl.querySelectorAll("[data-chart-index]"));

  const showTooltip = (index) => {
    const safeIndex = Math.max(0, Math.min(pointEls.length - 1, Number(index) || 0));
    const point = safePoints[safeIndex];
    const value = values[safeIndex];
    const coord = coords[safeIndex];
    tooltipEl.textContent = `${formatReportTooltipDate(point.date)} • ${valueFormatter(value)}`;
    tooltipEl.style.left = `${((coord.x / width) * 100).toFixed(2)}%`;
    tooltipEl.style.top = `${((coord.y / height) * 100).toFixed(2)}%`;
    tooltipEl.classList.remove("is-hidden");
    pointEls.forEach((el, idx) => {
      el.classList.toggle("is-active", idx === safeIndex);
    });
  };

  const hideTooltip = () => {
    tooltipEl.classList.add("is-hidden");
    pointEls.forEach((el) => el.classList.remove("is-active"));
  };

  pointEls.forEach((pointEl) => {
    pointEl.addEventListener("mouseenter", () => {
      showTooltip(pointEl.getAttribute("data-chart-index"));
    });
    pointEl.addEventListener("focus", () => {
      showTooltip(pointEl.getAttribute("data-chart-index"));
    });
    pointEl.addEventListener("mouseleave", hideTooltip);
    pointEl.addEventListener("blur", hideTooltip);
  });

  svgEl.onmouseleave = hideTooltip;
}

function setReportsServiceHighlight(index) {
  const normalizedIndex = Number.isInteger(index) ? index : -1;
  const segments = reportServicesDonut?.querySelectorAll("[data-service-index]");
  const legendItems = reportServicesLegend?.querySelectorAll("[data-service-index]");
  if (!segments || !legendItems) return;

  segments.forEach((segment) => {
    const itemIndex = Number(segment.getAttribute("data-service-index"));
    const muted = normalizedIndex >= 0 && itemIndex !== normalizedIndex;
    segment.classList.toggle("is-muted", muted);
  });
  legendItems.forEach((item) => {
    const itemIndex = Number(item.getAttribute("data-service-index"));
    const muted = normalizedIndex >= 0 && itemIndex !== normalizedIndex;
    item.classList.toggle("is-muted", muted);
  });
}

function renderReportsServices(model) {
  if (!reportServicesDonut || !reportServicesLegend || !reportServicesTotal) return;
  const services = Array.isArray(model.serviceRows) ? model.serviceRows : [];
  const total = services.reduce((sum, service) => sum + service.count, 0);
  reportServicesTotal.textContent = String(total || 0);

  if (!services.length || total === 0) {
    reportServicesDonut.innerHTML = `
      <circle class="reports-donut-track" cx="110" cy="110" r="72"></circle>
      <circle class="reports-donut-segment" cx="110" cy="110" r="72" stroke="#333a47" stroke-dasharray="${(
        2 * Math.PI * 72
      ).toFixed(2)} 0" stroke-dashoffset="0"></circle>
    `;
    reportServicesLegend.innerHTML = `<div class="reports-empty">No service data in this range.</div>`;
    reportsActiveServiceIndex = -1;
    return;
  }

  const radius = 72;
  const circumference = 2 * Math.PI * radius;
  let offset = 0;

  const segmentsMarkup = services
    .map((service, index) => {
      const ratio = service.count / total;
      const segmentLength = Math.max(0, ratio * circumference);
      const markup = `<circle class="reports-donut-segment" data-service-index="${index}" cx="110" cy="110" r="${radius}" stroke="${service.color}" stroke-dasharray="${segmentLength.toFixed(
        2
      )} ${(circumference - segmentLength).toFixed(2)}" stroke-dashoffset="${(-offset).toFixed(
        2
      )}" transform="rotate(-90 110 110)"></circle>`;
      offset += segmentLength;
      return markup;
    })
    .join("");

  reportServicesDonut.innerHTML = `
    <circle class="reports-donut-track" cx="110" cy="110" r="${radius}"></circle>
    ${segmentsMarkup}
  `;

  reportServicesLegend.innerHTML = "";
  services.forEach((service, index) => {
    const row = document.createElement("button");
    row.type = "button";
    row.className = "reports-legend-item";
    row.setAttribute("data-service-index", String(index));
    row.innerHTML = `
      <span class="reports-legend-left">
        <span class="reports-legend-dot" style="background:${service.color}"></span>
        <span class="reports-legend-name">${service.name}</span>
      </span>
      <span class="reports-legend-value mono">${service.count} • ${formatMoney(service.spend)}</span>
    `;
    row.addEventListener("mouseenter", () => setReportsServiceHighlight(index));
    row.addEventListener("mouseleave", () => setReportsServiceHighlight(reportsActiveServiceIndex));
    row.addEventListener("click", () => {
      reportsActiveServiceIndex = reportsActiveServiceIndex === index ? -1 : index;
      setReportsServiceHighlight(reportsActiveServiceIndex);
    });
    reportServicesLegend.appendChild(row);
  });

  Array.from(reportServicesDonut.querySelectorAll("[data-service-index]")).forEach((segment) => {
    const index = Number(segment.getAttribute("data-service-index"));
    segment.addEventListener("mouseenter", () => setReportsServiceHighlight(index));
    segment.addEventListener("mouseleave", () => setReportsServiceHighlight(reportsActiveServiceIndex));
    segment.addEventListener("click", () => {
      reportsActiveServiceIndex = reportsActiveServiceIndex === index ? -1 : index;
      setReportsServiceHighlight(reportsActiveServiceIndex);
    });
  });

  setReportsServiceHighlight(reportsActiveServiceIndex);
}

function renderReportsZones(model) {
  if (!reportZones) return;
  const zones = Array.isArray(model.zoneTotals) ? model.zoneTotals : Array.from({ length: 9 }, () => 0);
  const maxZone = Math.max(1, ...zones);
  reportZones.innerHTML = "";

  zones.forEach((count, index) => {
    const zoneIndex = index + 1;
    const zone = document.createElement("button");
    zone.type = "button";
    zone.className = "reports-zone";
    zone.innerHTML = `
      <div class="reports-zone-bar"></div>
      <div class="reports-zone-label mono">Zone ${zoneIndex}</div>
      <div class="reports-zone-value mono">${count}</div>
    `;
    const bar = zone.querySelector(".reports-zone-bar");
    const scale = Math.max(0.06, count / maxZone);
    bar.style.transform = `scaleY(${scale.toFixed(3)})`;
    zone.title = `Zone ${zoneIndex}: ${count} labels`;
    zone.addEventListener("mouseenter", () => zone.classList.add("is-active"));
    zone.addEventListener("mouseleave", () => zone.classList.remove("is-active"));
    reportZones.appendChild(zone);
  });
}

function renderReportsLocationTable(
  bodyEl,
  rows,
  emptyMessage,
  { withFlags = false, rowKey = (row) => row.name, onHover = null } = {}
) {
  if (!bodyEl) return;
  bodyEl.innerHTML = "";
  if (!Array.isArray(rows) || rows.length === 0) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 3;
    td.className = "reports-empty";
    td.textContent = emptyMessage;
    tr.appendChild(td);
    bodyEl.appendChild(tr);
    return;
  }

  rows.forEach((row) => {
    const key = String(rowKey(row) || row.name || "");
    const tr = document.createElement("tr");
    tr.dataset.rowKey = key;
    const nameCell = document.createElement("td");
    if (withFlags) {
      const wrapper = document.createElement("span");
      wrapper.style.display = "inline-flex";
      wrapper.style.alignItems = "center";
      wrapper.style.gap = "8px";
      const flag = document.createElement("span");
      flag.textContent = resolveCountryFlag(row.name);
      const text = document.createElement("span");
      text.textContent = row.name;
      wrapper.appendChild(flag);
      wrapper.appendChild(text);
      nameCell.appendChild(wrapper);
    } else {
      nameCell.textContent = row.name;
    }

    const countCell = document.createElement("td");
    countCell.className = "mono";
    countCell.textContent = String(row.count);
    const spendCell = document.createElement("td");
    spendCell.className = "mono";
    spendCell.textContent = formatMoney(row.spend);

    tr.appendChild(nameCell);
    tr.appendChild(countCell);
    tr.appendChild(spendCell);
    if (typeof onHover === "function") {
      tr.addEventListener("mouseenter", () => onHover(key));
      tr.addEventListener("mouseleave", () => onHover(""));
    }
    bodyEl.appendChild(tr);
  });
}

function showMapTooltip(tooltipEl, svgEl, x, y, text) {
  if (!tooltipEl || !svgEl) return;
  const box = svgEl.viewBox.baseVal;
  const left = ((x - box.x) / box.width) * 100;
  const top = ((y - box.y) / box.height) * 100;
  tooltipEl.textContent = text;
  tooltipEl.style.left = `${left.toFixed(2)}%`;
  tooltipEl.style.top = `${top.toFixed(2)}%`;
  tooltipEl.classList.remove("is-hidden");
}

function hideMapTooltip(tooltipEl) {
  if (!tooltipEl) return;
  tooltipEl.classList.add("is-hidden");
}

function setReportsCountryHighlight(key) {
  reportsActiveCountryKey = key || "";
  const shapeEls = reportCountriesMap?.querySelectorAll("[data-country-key]");
  const rowEls = reportTopCountriesBody?.querySelectorAll("[data-row-key]");
  shapeEls?.forEach((el) => {
    el.classList.toggle(
      "is-active",
      Boolean(reportsActiveCountryKey) && el.dataset.countryKey === reportsActiveCountryKey
    );
  });
  rowEls?.forEach((row) => {
    row.classList.toggle(
      "is-active",
      Boolean(reportsActiveCountryKey) && row.dataset.rowKey === reportsActiveCountryKey
    );
  });
}

function setReportsRegionHighlight(key) {
  reportsActiveRegionKey = key || "";
  const shapeEls = reportRegionsMap?.querySelectorAll("[data-region-key]");
  const rowEls = reportTopRegionsBody?.querySelectorAll("[data-row-key]");
  shapeEls?.forEach((el) => {
    el.classList.toggle(
      "is-active",
      Boolean(reportsActiveRegionKey) && el.dataset.regionKey === reportsActiveRegionKey
    );
  });
  rowEls?.forEach((row) => {
    row.classList.toggle(
      "is-active",
      Boolean(reportsActiveRegionKey) && row.dataset.rowKey === reportsActiveRegionKey
    );
  });
}

function renderReportsCountriesMap(model) {
  if (!reportCountriesMap) return;
  const countries = Array.isArray(model.topCountries) ? model.topCountries : [];
  if (!reportsWorldGeoJson?.features?.length) {
    reportCountriesMap.innerHTML = `
      <rect class="reports-map-bg" x="0" y="0" width="720" height="300"></rect>
      <text x="360" y="150" text-anchor="middle" class="reports-map-label">World map data unavailable</text>
    `;
    return;
  }

  const width = 720;
  const height = 300;
  const features = reportsWorldGeoJson.features.filter(
    (feature) => normalizeNameKey(getCountryFeatureName(feature)) !== "antarctica"
  );
  const bounds = getFeatureBounds(features);
  const project = createGeoProjector(bounds, width, height, 6, { isGeographic: true });
  const byCountryKey = new Map(
    countries.map((item) => [normalizeNameKey(item.name), item])
  );
  const maxCount = Math.max(1, ...countries.map((item) => item.count || 0));

  const mapMarkup = features
    .map((feature) => {
      const countryName = getCountryFeatureName(feature);
      const countryKey = normalizeNameKey(countryName);
      const stats = byCountryKey.get(countryKey);
      const pathData = geometryToPath(feature.geometry, project);
      if (!pathData) return "";
      let fillOpacity = 0.08;
      if (stats) {
        fillOpacity = 0.18 + ((stats.count || 0) / maxCount) * 0.44;
      }
      const fill = `rgba(119, 71, 227, ${fillOpacity.toFixed(3)})`;
      return `
        <path
          class="reports-map-country-shape${stats ? " reports-map-country-highlight" : ""}"
          data-country-key="${escapeHtml(countryKey)}"
          data-country-name="${escapeHtml(countryName)}"
          d="${pathData}"
          style="--country-fill:${fill}"
        ></path>
      `;
    })
    .join("");

  reportCountriesMap.innerHTML = `
    <rect class="reports-map-bg" x="0" y="0" width="${width}" height="${height}"></rect>
    ${mapMarkup}
  `;

  const shapeEls = reportCountriesMap.querySelectorAll("[data-country-key]");
  shapeEls.forEach((shapeEl) => {
    const key = shapeEl.getAttribute("data-country-key") || "";
    const name = shapeEl.getAttribute("data-country-name") || "Unknown";
    const row = byCountryKey.get(key);
    shapeEl.addEventListener("mouseenter", () => {
      setReportsCountryHighlight(key);
      const bbox = shapeEl.getBBox();
      const x = bbox.x + bbox.width / 2;
      const y = bbox.y + bbox.height / 2;
      const line = row
        ? `${name} • ${row.count} labels • ${formatMoney(row.spend)}`
        : `${name} • 0 labels • ${formatMoney(0)}`;
      showMapTooltip(reportCountriesMapTooltip, reportCountriesMap, x, y, line);
    });
    shapeEl.addEventListener("mouseleave", () => {
      setReportsCountryHighlight("");
      hideMapTooltip(reportCountriesMapTooltip);
    });
  });

  reportCountriesMap.onmouseleave = () => {
    setReportsCountryHighlight("");
    hideMapTooltip(reportCountriesMapTooltip);
  };
}

function renderReportsRegionsMap(model) {
  if (!reportRegionsMap) return;
  const rows = Array.isArray(model.topRegions) ? model.topRegions : [];
  if (!reportsDomesticGeoJson?.features?.length) {
    reportRegionsMap.innerHTML = `
      <rect class="reports-map-bg" x="0" y="0" width="360" height="220"></rect>
      <text x="180" y="112" text-anchor="middle" class="reports-map-label">Domestic map data unavailable</text>
    `;
    return;
  }

  const width = 360;
  const height = 220;
  const features = reportsDomesticGeoJson.features;
  const bounds = getFeatureBounds(features);
  const project = createGeoProjector(bounds, width, height, 8, { isGeographic: true });
  const rowMap = new Map(rows.map((row) => [row.name, row]));
  const maxCount = Math.max(1, ...rows.map((row) => row.count || 0));

  const paths = features
    .map((feature) => {
      const province = getDomesticProvinceName(feature);
      const region = getDomesticRegionName(feature);
      const stats = rowMap.get(region) || { count: 0, spend: 0 };
      const pathData = geometryToPath(feature.geometry, project);
      if (!pathData) return "";
      const intensity = Math.max(0.14, (stats.count || 0) / maxCount);
      const fill = `rgba(119, 71, 227, ${(0.16 + intensity * 0.42).toFixed(3)})`;
      return `
        <path
          class="reports-map-region-shape"
          data-region-key="${escapeHtml(region)}"
          data-province-name="${escapeHtml(province)}"
          d="${pathData}"
          style="--region-fill:${fill}"
        ></path>
      `;
    })
    .join("");

  reportRegionsMap.innerHTML = `
    <rect class="reports-map-bg" x="0" y="0" width="${width}" height="${height}"></rect>
    ${paths}
  `;

  const shapes = reportRegionsMap.querySelectorAll(".reports-map-region-shape");
  shapes.forEach((shapeEl) => {
    const region = shapeEl.getAttribute("data-region-key") || "";
    const province = shapeEl.getAttribute("data-province-name") || "";
    const row = rowMap.get(region) || { count: 0, spend: 0 };
    shapeEl.addEventListener("mouseenter", () => {
      setReportsRegionHighlight(region);
      const bbox = shapeEl.getBBox();
      const x = bbox.x + bbox.width / 2;
      const y = bbox.y + bbox.height / 2;
      showMapTooltip(
        reportRegionsMapTooltip,
        reportRegionsMap,
        x,
        y,
        `${province} (${region}) • ${row.count} labels • ${formatMoney(row.spend)}`
      );
    });
    shapeEl.addEventListener("mouseleave", () => {
      setReportsRegionHighlight("");
      hideMapTooltip(reportRegionsMapTooltip);
    });
  });

  reportRegionsMap.onmouseleave = () => {
    setReportsRegionHighlight("");
    hideMapTooltip(reportRegionsMapTooltip);
  };
}

function setReportRangeButtons(rangeToken) {
  reportsRangeButtons.forEach((button) => {
    const token = button.dataset.reportRange;
    const buttonToken = token === "custom" ? "custom" : Number(token);
    button.classList.toggle("is-active", buttonToken === rangeToken);
  });
}

function seedReportsCustomDatesFromPreset() {
  const today = toDayStart(new Date());
  const safeRange = [7, 30, 90, 365].includes(Number(reportsRangeDays))
    ? Number(reportsRangeDays)
    : 30;
  const start = new Date(today);
  start.setDate(start.getDate() - (safeRange - 1));
  reportsCustomStart = formatDateInputValue(start);
  reportsCustomEnd = formatDateInputValue(today);
  if (reportsStartDate) reportsStartDate.value = reportsCustomStart;
  if (reportsEndDate) reportsEndDate.value = reportsCustomEnd;
}

function activateCustomReportsRangePicker() {
  reportsRangeMode = "custom";
  if (!reportsCustomStart || !reportsCustomEnd) {
    seedReportsCustomDatesFromPreset();
  }
  setReportRangeButtons("custom");
  setReportsCustomRangeVisible(true);
}

function applyCustomReportsRange() {
  if (!reportsStartDate || !reportsEndDate) return;
  const nextStart = reportsStartDate.value;
  const nextEnd = reportsEndDate.value;
  const startDate = toDayStart(nextStart);
  const endDate = toDayStart(nextEnd);
  const isValid = Boolean(startDate && endDate && startDate <= endDate);
  reportsStartDate.classList.toggle("is-invalid", !isValid);
  reportsEndDate.classList.toggle("is-invalid", !isValid);
  if (!isValid) return;
  reportsStartDate.classList.remove("is-invalid");
  reportsEndDate.classList.remove("is-invalid");
  reportsCustomStart = nextStart;
  reportsCustomEnd = nextEnd;
  reportsRangeMode = "custom";
  renderReportsDashboard();
}

function renderReportsDashboard() {
  if (!reportsPageSection) return;
  const model = buildReportsModel();
  reportsActiveCountryKey = "";
  reportsActiveRegionKey = "";
  setReportRangeButtons(model.rangeToken);
  setReportsCustomRangeVisible(model.rangeMode === "custom");
  if (reportsStartDate && reportsEndDate) {
    if (!reportsCustomStart) {
      reportsCustomStart = formatDateInputValue(model.rangeStart);
    }
    if (!reportsCustomEnd) {
      reportsCustomEnd = formatDateInputValue(model.rangeEnd);
    }
    reportsStartDate.value = reportsCustomStart;
    reportsEndDate.value = reportsCustomEnd;
  }

  if (reportSavingsValue) reportSavingsValue.textContent = formatMoney(model.totalSavings);
  if (reportSavingsMeta) {
    reportSavingsMeta.textContent = `${model.totalShipments} labels analyzed`;
  }

  if (reportPendingReturnsValue) {
    reportPendingReturnsValue.textContent = `${model.pendingReturnsCount} labels`;
  }
  if (reportPendingReturnsMeta) {
    reportPendingReturnsMeta.textContent = `${formatMoney(model.pendingReturnsAmount)} pending`;
  }

  if (reportPendingRefundsValue) {
    reportPendingRefundsValue.textContent = `${model.pendingRefundsCount} labels`;
  }
  if (reportPendingRefundsMeta) {
    reportPendingRefundsMeta.textContent = `${formatMoney(model.pendingRefundsAmount)} pending`;
  }

  if (reportAccountBalanceValue) {
    reportAccountBalanceValue.textContent = formatMoney(model.accountBalance);
  }
  if (reportTotalShippingCostsValue) {
    reportTotalShippingCostsValue.textContent = formatMoney(model.totalShippingCosts);
  }
  if (reportTotalShippingCostsMeta) {
    reportTotalShippingCostsMeta.textContent = `EX VAT • INCL ${formatMoney(
      model.totalShippingCosts * (1 + VAT_RATE)
    )}`;
  }
  if (reportCarrierAdjustmentsValue) {
    reportCarrierAdjustmentsValue.textContent = formatSignedMoney(model.carrierAdjustments);
  }
  if (reportCarrierAdjustmentsMeta) {
    reportCarrierAdjustmentsMeta.textContent = `${model.adjustmentEvents} events`;
  }

  if (reportTotalPaymentsValue) {
    reportTotalPaymentsValue.textContent = formatMoney(model.totalPayments);
  }
  if (reportAverageCostValue) {
    reportAverageCostValue.textContent = formatMoney(model.avgCost);
  }
  if (reportAvgDomesticValue) {
    reportAvgDomesticValue.textContent = formatMoney(model.avgDomestic);
  }
  if (reportAvgDomesticMeta) {
    reportAvgDomesticMeta.textContent = `${model.avgDomesticCount} labels`;
  }
  if (reportAvgInternationalValue) {
    reportAvgInternationalValue.textContent = formatMoney(model.avgInternational);
  }
  if (reportAvgInternationalMeta) {
    reportAvgInternationalMeta.textContent = `${model.avgInternationalCount} labels`;
  }
  if (reportNewShipmentsValue) {
    reportNewShipmentsValue.textContent = String(model.newShipments);
  }
  if (reportNewShipmentsMeta) {
    reportNewShipmentsMeta.textContent = `${formatMoney(model.totalShippingCosts)} spend`;
  }
  if (reportDeliveryIssuesValue) {
    reportDeliveryIssuesValue.textContent = String(model.deliveryIssues);
  }
  if (reportDeliveryIssuesMeta) {
    reportDeliveryIssuesMeta.textContent = `${model.deliveryIssueRate.toFixed(2)}% rate`;
  }
  if (reportTotalShipmentsValue) {
    reportTotalShipmentsValue.textContent = String(model.totalShipments);
  }

  renderReportLineChart({
    svgEl: reportPaymentsChart,
    axisEl: reportPaymentsAxis,
    tooltipEl: reportPaymentsTooltip,
    points: model.buckets,
    valueKey: "payments",
    valueFormatter: formatMoney,
    tooltipLabel: "Payments",
  });
  renderReportLineChart({
    svgEl: reportAverageCostChart,
    axisEl: reportAverageCostAxis,
    tooltipEl: reportAverageCostTooltip,
    points: model.buckets,
    valueKey: "averageCost",
    valueFormatter: formatMoney,
    tooltipLabel: "Average Cost",
  });
  renderReportLineChart({
    svgEl: reportShipmentsChart,
    axisEl: reportShipmentsAxis,
    tooltipEl: reportShipmentsTooltip,
    points: model.buckets,
    valueKey: "shipments",
    valueFormatter: (value) => `${Math.round(value)} labels`,
    tooltipLabel: "Shipments",
  });
  renderReportsServices(model);
  renderReportsZones(model);
  renderReportsLocationTable(
    reportTopRegionsBody,
    model.topRegions,
    "No domestic region data yet.",
    {
      rowKey: (row) => row.name,
      onHover: (key) => setReportsRegionHighlight(key),
    }
  );
  renderReportsLocationTable(
    reportTopCountriesBody,
    model.topCountries,
    "No international country data yet.",
    {
      withFlags: true,
      rowKey: (row) => normalizeNameKey(row.name),
      onHover: (key) => setReportsCountryHighlight(key),
    }
  );
  renderReportsRegionsMap(model);
  renderReportsCountriesMap(model);
}

function refreshReportsIfVisible() {
  if (currentMainView === "reports") {
    renderReportsDashboard();
  }
}

function setAuthMessage(message, { isError = true } = {}) {
  if (!authError) return;
  authError.textContent = message || "";
  authError.classList.toggle("is-visible", Boolean(message));
  authError.classList.toggle("is-info", Boolean(message) && !isError);
}

function setAuthBusy(isBusy, signInLabel = "Sign In", signUpLabel = "Create Account") {
  if (authEmail) authEmail.disabled = isBusy;
  if (authPassword) authPassword.disabled = isBusy;
  if (authSignIn) {
    authSignIn.disabled = isBusy;
    authSignIn.textContent = signInLabel;
  }
  if (authSignUp) {
    authSignUp.disabled = isBusy;
    authSignUp.textContent = signUpLabel;
  }
}

function setAuthView(session, options = {}) {
  const { animate = true } = options;
  currentUser = session?.user || null;
  const isAuthed = Boolean(currentUser);
  renderAccountProfile(currentUser);
  if (isAuthed) {
    loadWarehouseSettings({ quiet: true });
    loadShopifyConnectionStatus({ quiet: true });
  } else {
    resetWarehouseState();
    shopifyConnection = null;
    updateShopifyProviderStatus();
  }
  transitionShellVisibility(isAuthed, { animate });
  if (accountChip) {
    accountChip.classList.toggle("is-hidden", !isAuthed);
    accountChip.innerHTML = "";
    if (isAuthed) {
      const dot = document.createElement("span");
      dot.className = "account-chip-dot";
      const text = document.createElement("span");
      text.className = "account-chip-text";
      text.textContent = currentUser?.email || "";
      accountChip.appendChild(dot);
      accountChip.appendChild(text);
    }
  }
  if (signOutButton) {
    signOutButton.classList.toggle("is-hidden", !isAuthed);
  }
  if (openAccountPageButton) {
    openAccountPageButton.classList.toggle("is-hidden", !isAuthed);
  }
  if (openReportsPageButton) {
    openReportsPageButton.classList.toggle("is-hidden", !isAuthed);
  }
  if (!isAuthed) {
    setReceiptModalOpen(false);
    setMainView("builder", { push: false, animate: false });
    resetAccountPreview();
    updateRoute({ view: "login" }, { replace: true });
    return;
  }

  const route = parseRouteFromLocation();
  if (route.view === "login") {
    updateRoute({ view: "builder", step: state.step }, { replace: true });
  }
}

async function signInWithPassword() {
  if (!supabaseClient) {
    setAuthMessage("Supabase auth is not configured.");
    return;
  }
  const email = authEmail?.value.trim().toLowerCase() || "";
  const password = authPassword?.value || "";
  if (!email || !password) {
    setAuthMessage("Email and password are required.");
    return;
  }
  setAuthMessage("");
  setAuthBusy(true, "Signing In...", "Create Account");
  const { error } = await supabaseClient.auth.signInWithPassword({
    email,
    password,
  });
  setAuthBusy(false, "Sign In", "Create Account");
  if (error) {
    setAuthMessage(error.message);
  }
}

async function signUpWithPassword() {
  if (!supabaseClient) {
    setAuthMessage("Supabase auth is not configured.");
    return;
  }
  const email = authEmail?.value.trim().toLowerCase() || "";
  const password = authPassword?.value || "";
  if (!email || !password) {
    setAuthMessage("Email and password are required.");
    return;
  }
  setAuthMessage("");
  setAuthBusy(true, "Sign In", "Creating...");
  const { error } = await supabaseClient.auth.signUp({
    email,
    password,
  });
  setAuthBusy(false, "Sign In", "Create Account");
  if (error) {
    setAuthMessage(error.message);
    return;
  }
  setAuthMessage(
    "Account created. If confirmation is enabled, approve the email then sign in.",
    { isError: false }
  );
}

async function resetPassword() {
  if (!supabaseClient) {
    setAuthMessage("Supabase auth is not configured.");
    return;
  }
  const email = authEmail?.value.trim().toLowerCase() || "";
  if (!email) {
    setAuthMessage("Enter your email address, then click Forgot password.");
    return;
  }
  setAuthMessage("");
  setAuthBusy(true, "Sign In", "Create Account");
  const { error } = await supabaseClient.auth.resetPasswordForEmail(email, {
    redirectTo: window.location.origin,
  });
  setAuthBusy(false, "Sign In", "Create Account");
  if (error) {
    setAuthMessage(error.message);
    return;
  }
  setAuthMessage("If that email exists, a reset link has been sent. Check your inbox.", {
    isError: false,
  });
}

async function initializeAuth() {
  if (!supabaseClient) {
    setAuthMessage("Supabase client failed to initialize.");
    setAuthView(null, { animate: false });
    return;
  }
  const {
    data: { session },
    error,
  } = await supabaseClient.auth.getSession();
  if (error) {
    setAuthMessage(error.message);
  }
  setAuthView(session, { animate: false });
  if (session?.user) {
    consumeShopifyCallbackParams();
  }
  const initialRoute = parseRouteFromLocation();
  const isAuthed = Boolean(session);
  if (isAuthed && initialRoute.view === "account") {
    setMainView("account", { push: false, animate: false });
  } else if (isAuthed && initialRoute.view === "reports") {
    setMainView("reports", { push: false, animate: false });
  } else {
    setMainView("builder", { push: false, animate: false });
    if (isAuthed && initialRoute.view === "builder") {
      goToStep(clampStep(initialRoute.step), { push: false, regenerate: false });
    } else if (isAuthed && initialRoute.view !== "builder") {
      updateRoute({ view: "builder", step: state.step }, { replace: true });
    }
  }

  await loadGenerationHistory({
    preferLatest: isAuthed && (initialRoute.view === "account" || initialRoute.view === "reports"),
  });
  supabaseClient.auth.onAuthStateChange(async (_event, updatedSession) => {
    setAuthMessage("");
    setAuthView(updatedSession, { animate: true });
    if (!updatedSession) {
      resetAll();
      goToStep(1, { push: false, regenerate: false });
      setMainView("builder", { push: false, animate: false });
      await loadGenerationHistory({ preferLatest: false });
      updateRoute({ view: "login" }, { replace: true });
      return;
    }

    const route = parseRouteFromLocation();
    if (route.view === "account") {
      setMainView("account", { push: false, animate: false });
    } else if (route.view === "reports") {
      setMainView("reports", { push: false, animate: false });
    } else {
      setMainView("builder", { push: false, animate: false });
      if (route.view === "builder") {
        goToStep(clampStep(route.step), { push: false, regenerate: false });
      } else {
        updateRoute({ view: "builder", step: state.step }, { replace: true });
      }
    }
    await loadGenerationHistory({ preferLatest: route.view === "account" || route.view === "reports" });
  });
}

function applyAuthLogoLayout(layout) {
  const useInline = layout === "inline";
  if (authLogoLottie) {
    authLogoLottie.classList.toggle("is-hidden", useInline);
  }
  if (authBrandInlineLogo) {
    authBrandInlineLogo.classList.toggle("is-hidden", !useInline);
  }
  if (authBrandOriginal) {
    authBrandOriginal.classList.toggle("is-hidden", useInline);
  }
  if (authBrand) {
    authBrand.classList.toggle("is-logo-only", useInline);
  }
}

function initializeAuthLogo() {
  const layout = AUTH_LOGO_LAYOUT === "stack" ? "stack" : "inline";
  applyAuthLogoLayout(layout);
  if (!window.lottie) return;
  const target = layout === "inline" ? authBrandInlineLogo : authLogoLottie;
  if (!target) return;
  target.style.pointerEvents = "none";
  const anim = window.lottie.loadAnimation({
    container: target,
    renderer: "svg",
    loop: false,
    autoplay: true,
    path: "assets/flow-logo.json",
  });
  anim.addEventListener("DOMLoaded", () => {
    const svg = target.querySelector("svg");
    if (svg) svg.style.pointerEvents = "none";
  });
}

function initializeAppLogo() {
  if (!window.lottie || !appLogoLottie) return;
  window.lottie.loadAnimation({
    container: appLogoLottie,
    renderer: "svg",
    loop: false,
    autoplay: true,
    path: "assets/flow-logo.json",
  });
}

function initializeAuthParticles() {
  if (authParticlesStarted || !authGate || !window.THREE) return;
  authParticlesStarted = true;

  const movementSpeed = 15;
  const totalObjects = 1100;
  const objectSize = 8;
  const dampingFactor = 0.992;
  const minAlpha = 0.2;
  const maxAlpha = 0.8;
  const alphaShimmerSpeed = 0.005;
  const targetSpreadForFullOpacity = 400;
  const minSceneOpacity = 0;
  const baseColors = [
    new THREE.Color(0x7747e3),
    new THREE.Color(0x8f66ed),
    new THREE.Color(0x6b2fcc),
    new THREE.Color(0xa18af0),
    new THREE.Color(0x5c29b8),
  ];

  const renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true });
  renderer.setPixelRatio(Math.min(window.devicePixelRatio || 1, 1.6));
  renderer.domElement.className = "auth-particles-canvas";
  renderer.domElement.style.opacity = "0";
  renderer.domElement.style.pointerEvents = "none";
  authGate.prepend(renderer.domElement);

  const scene = new THREE.Scene();
  const camera = new THREE.PerspectiveCamera(75, 1, 1, 10000);
  camera.position.z = 800;

  const geometry = new THREE.BufferGeometry();
  const positions = new Float32Array(totalObjects * 3);
  const colors = new Float32Array(totalObjects * 4);
  const directions = [];
  const startX = 0;
  const startY = 0;

  for (let i = 0; i < totalObjects; i += 1) {
    const pIndex = i * 3;
    const cIndex = i * 4;
    positions[pIndex] = startX;
    positions[pIndex + 1] = startY;
    positions[pIndex + 2] = 0;

    directions.push({
      x: Math.random() * movementSpeed - movementSpeed / 2,
      y: Math.random() * movementSpeed - movementSpeed / 2,
      z: Math.random() * movementSpeed - movementSpeed / 2,
      alphaDir: (Math.random() > 0.5 ? 1 : -1) * alphaShimmerSpeed,
    });

    const baseColor = baseColors[Math.floor(Math.random() * baseColors.length)];
    const initialAlpha = Math.random() * (maxAlpha - minAlpha) + minAlpha;
    colors[cIndex] = baseColor.r;
    colors[cIndex + 1] = baseColor.g;
    colors[cIndex + 2] = baseColor.b;
    colors[cIndex + 3] = initialAlpha;
  }

  geometry.setAttribute("position", new THREE.BufferAttribute(positions, 3));
  geometry.setAttribute("color", new THREE.BufferAttribute(colors, 4));

  const material = new THREE.PointsMaterial({
    size: objectSize,
    vertexColors: true,
    transparent: true,
    sizeAttenuation: true,
    depthWrite: false,
  });
  const points = new THREE.Points(geometry, material);
  scene.add(points);

  const updateParticles = () => {
    let currentMaxDistanceSq = 0;
    for (let i = 0; i < totalObjects; i += 1) {
      const pIndex = i * 3;
      const cIndex = i * 4;
      const vector = directions[i];

      vector.x *= dampingFactor;
      vector.y *= dampingFactor;
      vector.z *= dampingFactor;

      positions[pIndex] += vector.x;
      positions[pIndex + 1] += vector.y;
      positions[pIndex + 2] += vector.z;

      const dx = positions[pIndex] - startX;
      const dy = positions[pIndex + 1] - startY;
      const distanceSq = dx * dx + dy * dy;
      if (distanceSq > currentMaxDistanceSq) {
        currentMaxDistanceSq = distanceSq;
      }

      let alpha = colors[cIndex + 3] + vector.alphaDir;
      if (alpha > maxAlpha || alpha < minAlpha) {
        vector.alphaDir *= -1;
        alpha = Math.max(minAlpha, Math.min(maxAlpha, alpha));
      }
      colors[cIndex + 3] = alpha;
    }

    geometry.attributes.position.needsUpdate = true;
    geometry.attributes.color.needsUpdate = true;
    return Math.sqrt(currentMaxDistanceSq);
  };

  for (let i = 0; i < AUTH_PARTICLES_WARMUP_STEPS; i += 1) {
    updateParticles();
  }

  const onResize = () => {
    const width = window.innerWidth;
    const height = window.innerHeight;
    renderer.setSize(width, height);
    camera.aspect = width / height;
    camera.updateProjectionMatrix();
  };

  onResize();
  window.addEventListener("resize", onResize);
  let particlesVisible = false;
  requestAnimationFrame(() => {
    particlesVisible = true;
    authGate.classList.add("is-ready");
  });

  const animate = () => {
    requestAnimationFrame(animate);
    if (authGate.classList.contains("is-hidden")) {
      return;
    }
    const maxSpread = updateParticles();
    const sceneOpacity = Math.min(
      1,
      Math.max(minSceneOpacity, maxSpread / targetSpreadForFullOpacity)
    );
    renderer.domElement.style.opacity = particlesVisible ? sceneOpacity.toFixed(2) : "0";
    renderer.render(scene, camera);
  };

  animate();
}

function getCountryIcon(country) {
  return resolveCountryFlag(country);
}

function updateCountryFlag() {
  if (!countryFlag) return;
  const value = inputMap.recipientCountry.value.trim();
  countryFlag.innerHTML = getCountryIcon(value);
}

function setCsvMode(enabled) {
  state.csvMode = enabled;
  invalidateCustomsCompletion();
  if (labelForm) {
    labelForm.classList.toggle("is-hidden", enabled);
  }
  if (csvSection) {
    csvSection.classList.toggle("is-visible", enabled);
  }
  if (csvEditToggle) {
    csvEditToggle.disabled = !enabled;
  }
  if (quantityInput) {
    quantityInput.disabled = enabled;
  }
  if (enabled) {
    updateQuantityFromCsv();
  }
}

function setCsvEditMode(enabled) {
  state.csvEditable = enabled;
  if (csvEditToggle) {
    csvEditToggle.classList.toggle("is-active", enabled);
    const label = csvEditToggle.querySelector("span");
    if (label) label.textContent = enabled ? "Editing" : "Edit";
  }
  renderCsvTable();
}

function updateQuantityFromCsv() {
  if (!quantityInput) return;
  const count = Math.max(1, state.csvRows.length || 1);
  quantityInput.value = count;
  state.quantity = count;
  updateSummary();
  updatePayment();
}

function buildCsvRow(profile) {
  return {
    senderName: profile.sender.name,
    senderStreet: profile.sender.street,
    senderCity: profile.sender.city,
    senderState: profile.sender.state,
    senderZip: profile.sender.zip,
    recipientName: profile.recipient.name,
    recipientStreet: profile.recipient.street,
    recipientCountry: profile.recipient.country || "France",
    recipientCity: profile.recipient.city,
    recipientState: profile.recipient.state,
    recipientZip: profile.recipient.zip,
    packageWeight: (Math.random() * 2.7 + 0.3).toFixed(1),
    packageDims: `${Math.floor(Math.random() * 17 + 16)} x ${Math.floor(
      Math.random() * 12 + 12
    )} x ${Math.floor(Math.random() * 8 + 4)}`,
  };
}

function generateCsvRows(count = 3) {
  const rows = [];
  for (let i = 0; i < count; i += 1) {
    const profile = sampleProfiles[i % sampleProfiles.length];
    rows.push(buildCsvRow(profile));
  }
  return rows;
}

function renderCsvTable() {
  if (!csvTableBody) return;
  csvTableBody.innerHTML = "";
  state.csvRows.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");
    csvReviewColumns.forEach((column) => {
      const td = document.createElement("td");
      const input = document.createElement("input");
      input.type = "text";
      input.value = row[column.key] ?? "";
      input.dataset.row = String(rowIndex);
      input.dataset.key = column.key;
      input.disabled = !state.csvEditable;
      input.addEventListener("input", handleCsvInput);

      if (column.key === "recipientCountry") {
        const wrapper = document.createElement("div");
        wrapper.className = "csv-country";
        const icon = document.createElement("span");
        icon.className = "csv-country-icon";
        icon.innerHTML = getCountryIcon(input.value);
        wrapper.appendChild(icon);
        wrapper.appendChild(input);
        td.appendChild(wrapper);
      } else {
        td.appendChild(input);
      }
      tr.appendChild(td);
    });
    csvTableBody.appendChild(tr);
  });
}

function handleCsvInput(event) {
  const input = event.target;
  const rowIndex = Number(input.dataset.row);
  const key = input.dataset.key;
  if (Number.isNaN(rowIndex) || !key) return;
  if (!state.csvRows[rowIndex]) return;
  state.csvRows[rowIndex][key] = input.value.trim();
  invalidateCustomsCompletion();
  input.classList.remove("is-invalid");
  if (labelError) {
    labelError.classList.remove("is-visible");
  }
  if (key === "recipientCountry") {
    const icon = input.parentElement?.querySelector(".csv-country-icon");
    if (icon) {
      icon.innerHTML = getCountryIcon(input.value.trim());
    }
  }
  updateSummary();
  updatePreview();
}

function validateCsvRows() {
  if (!csvTableBody || state.csvRows.length === 0) {
    return false;
  }
  let isValid = true;
  const inputs = csvTableBody.querySelectorAll("input");
  inputs.forEach((input) => {
    const key = String(input.dataset.key || "");
    if (!CSV_REQUIRED_FIELDS.has(key)) {
      input.classList.remove("is-invalid");
      return;
    }
    const value = input.value.trim();
    const isFieldValid = value !== "";
    input.classList.toggle("is-invalid", !isFieldValid);
    if (!isFieldValid) {
      isValid = false;
    }
  });
  return isValid;
}

function selectLabel(card) {
  labelCards.forEach((item) => item.classList.remove("is-selected"));
  card.classList.add("is-selected");

  state.selection.type = card.dataset.type;
  state.selection.speed = card.dataset.speed;
  state.selection.price = Number(card.dataset.price);

  updateSummary();
  updatePayment();
  updatePreview();
}

labelCards.forEach((card) => {
  card.addEventListener("click", () => selectLabel(card));
});

// Click stepper items to navigate back to completed/current steps
stepperItems.forEach((item) => {
  item.style.cursor = "pointer";
  item.addEventListener("click", () => {
    const targetStep = Number(item.dataset.step);
    // Allow navigating to completed or current steps, never forward past current
    // Block navigation away from step 4 (labels generated) — use "Start new label" instead
    if (targetStep >= state.step) return;
    if (state.step === 4 && state.labels.length > 0) return;
    goToStep(targetStep);
  });
});

function updateSummary() {
  summaryService.textContent = state.selection.type;
  const quantity = getQuantity();
  const total = state.selection.price * quantity;
  const activeLabel = getActiveLabel();
  const activeData = activeLabel.data || state.info;
  summaryPrice.textContent = formatMoney(state.selection.price);
  summaryQty.textContent = String(quantity);
  summaryTotal.textContent = formatMoney(total);
  summaryTracking.textContent = activeLabel.trackingId || "--";
  summaryLabelId.textContent = activeLabel.labelId || "--";
}

function updatePayment() {
  const quantity = getQuantity();
  const subtotal = Number((state.selection.price * quantity).toFixed(2));
  const vatAmount = Number((subtotal * VAT_RATE).toFixed(2));
  const total = Number((subtotal + vatAmount).toFixed(2));

  if (paymentService) {
    paymentService.textContent = state.selection.type;
  }
  if (invoiceQty) {
    invoiceQty.textContent = String(quantity);
  }
  if (invoiceSubtotal) {
    invoiceSubtotal.textContent = formatMoney(subtotal);
  }
  if (invoiceVat) {
    invoiceVat.textContent = formatMoney(vatAmount);
  }
  if (paymentTotal) {
    paymentTotal.textContent = formatMoney(total);
  }
}

function updatePreview() {
  const activeLabel = getActiveLabel();
  const activeData = activeLabel.data || state.info;
  previewService.textContent = state.selection.type;
  previewTracking.textContent = activeLabel.trackingId || "TRACKING";

  previewFrom.textContent = formatAddress(
    activeData.senderName,
    activeData.senderStreet,
    activeData.senderCity,
    activeData.senderState,
    activeData.senderZip,
    ""
  );

  previewTo.textContent = formatAddress(
    activeData.recipientName,
    activeData.recipientStreet,
    activeData.recipientCity,
    activeData.recipientState,
    activeData.recipientZip,
    activeData.recipientCountry
  );

  previewWeight.textContent = activeData.packageWeight
    ? `${activeData.packageWeight} kg`
    : "--";
  previewDims.textContent = activeData.packageDims || "--";
}

function renderActiveBuilderPanel() {
  stepPanels.forEach((panel) => {
    const isActive = !customsGhostVisible && Number(panel.dataset.panel) === state.step;
    panel.classList.toggle("is-active", isActive);
  });
  if (customsGhostPanel) {
    customsGhostPanel.classList.toggle("is-active", customsGhostVisible);
  }
}

function setCustomsError(message = "") {
  if (!customsError) return;
  customsError.textContent = message;
  customsError.classList.toggle("is-visible", Boolean(message));
}

function setCustomsFieldInvalid(input, invalid) {
  if (!input) return;
  input.classList.toggle("is-invalid", Boolean(invalid));
}

function setCustomsItemFieldInvalid(index, fieldKey, invalid) {
  if (!customsItemsList) return;
  const input = customsItemsList.querySelector(
    `[data-customs-item-index="${index}"] [data-customs-item-field="${fieldKey}"]`
  );
  setCustomsFieldInvalid(input, invalid);
}

function syncCustomsStateFromInputs() {
  if (customsSummaryDescriptionInput) {
    state.customs.summaryDescription = customsSummaryDescriptionInput.value.trim();
  }
  if (customsShipmentTypeInput) {
    state.customs.shipmentType = customsShipmentTypeInput.value.trim();
  }
  if (customsSenderReferenceInput) {
    state.customs.senderReference = customsSenderReferenceInput.value.trim();
  }
  if (customsImporterReferenceInput) {
    state.customs.importerReference = customsImporterReferenceInput.value.trim();
  }
}

function syncCustomsInputsFromState() {
  if (customsSummaryDescriptionInput) {
    customsSummaryDescriptionInput.value = state.customs.summaryDescription || "";
  }
  if (customsShipmentTypeInput) {
    customsShipmentTypeInput.value = state.customs.shipmentType || "";
  }
  if (customsSenderReferenceInput) {
    customsSenderReferenceInput.value = state.customs.senderReference || "";
  }
  if (customsImporterReferenceInput) {
    customsImporterReferenceInput.value = state.customs.importerReference || "";
  }
}

function seedCustomsDefaults() {
  if (!Array.isArray(state.customs.items) || !state.customs.items.length) {
    state.customs.items = [createEmptyCustomsItem()];
  }
  const first = state.customs.items[0];
  const weightSource = state.csvMode
    ? String(state.csvRows?.[0]?.packageWeight || "").trim()
    : String(state.info.packageWeight || inputMap.packageWeight?.value || "").trim();
  if (!first.weightKg && weightSource) {
    first.weightKg = weightSource;
  }
  const defaultOriginCountry =
    String(getDefaultWarehouseRecord()?.country || "").trim() || "France";
  if (!first.originCountry) {
    first.originCountry = defaultOriginCountry;
  }
}

function updateCustomsScopeMeta() {
  if (!customsScopeMeta) return;
  const uniqueCountries = Array.from(new Set(getShipmentCountryValues()));
  const outsideEu = uniqueCountries.filter((country) => !isEuCountry(country));
  if (!outsideEu.length) {
    customsScopeMeta.textContent =
      "Complete this declaration for destinations that require customs processing.";
    return;
  }
  if (outsideEu.length === 1) {
    customsScopeMeta.textContent = `${outsideEu[0]} is outside the customs union. Complete declaration details before service selection.`;
    return;
  }
  customsScopeMeta.textContent = `${outsideEu.length} destinations are outside the customs union. Complete declaration details before service selection.`;
}

function buildCustomsItemsPayload() {
  return (state.customs.items || []).map((item) => ({
    description: String(item?.description || "").trim(),
    quantity: String(item?.quantity || "").trim(),
    weightKg: String(item?.weightKg || "").trim(),
    valueEur: String(item?.valueEur || "").trim(),
    originCountry: String(item?.originCountry || "").trim(),
    hsCode: String(item?.hsCode || "").trim(),
  }));
}

function renderCustomsItems() {
  if (!customsItemsList) return;
  customsItemsList.innerHTML = "";
  const items = buildCustomsItemsPayload();
  state.customs.items = items.length ? items : [createEmptyCustomsItem()];

  state.customs.items.forEach((item, index) => {
    const card = document.createElement("article");
    card.className = "customs-item";
    card.dataset.customsItemIndex = String(index);

    const head = document.createElement("div");
    head.className = "customs-item-head";

    const title = document.createElement("div");
    title.className = "customs-item-title";
    title.textContent = `Item ${index + 1}`;
    head.appendChild(title);

    if (state.customs.items.length > 1) {
      const remove = document.createElement("button");
      remove.type = "button";
      remove.className = "btn btn-ghost btn-sm customs-item-remove";
      remove.dataset.customsItemAction = "remove";
      remove.dataset.customsItemIndex = String(index);
      remove.innerHTML =
        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><line x1="5" y1="12" x2="19" y2="12"/></svg><span>Remove</span>';
      head.appendChild(remove);
    }

    const note = document.createElement("div");
    note.className = "customs-note";
    note.innerHTML =
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M12 9v4"/><circle cx="12" cy="17" r="1" fill="currentColor" stroke="none"/><path d="M10.3 3.9L1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/></svg><span>Provide a precise product description and material/composition.</span>';

    const descriptionField = document.createElement("label");
    descriptionField.className = "field";
    descriptionField.innerHTML = `<span>Item Description</span><input type="text" data-customs-item-field="description" value="${escapeHtml(
      item.description
    )}" />`;

    const rowThree = document.createElement("div");
    rowThree.className = "field-row customs-row-three";

    const quantityField = document.createElement("label");
    quantityField.className = "field";
    quantityField.innerHTML = `<span>Number of Items</span><input type="number" min="1" step="1" data-customs-item-field="quantity" value="${escapeHtml(
      item.quantity
    )}" />`;

    const weightField = document.createElement("label");
    weightField.className = "field";
    weightField.innerHTML = `<span>Weight (kg)</span><input type="number" min="0.01" step="0.01" data-customs-item-field="weightKg" value="${escapeHtml(
      item.weightKg
    )}" />`;

    const valueField = document.createElement("label");
    valueField.className = "field";
    valueField.innerHTML = `<span>Value (EUR)</span><div class="customs-currency"><span class="customs-currency-prefix">€</span><input type="number" min="0" step="0.01" data-customs-item-field="valueEur" value="${escapeHtml(
      item.valueEur
    )}" /></div>`;

    rowThree.appendChild(quantityField);
    rowThree.appendChild(weightField);
    rowThree.appendChild(valueField);

    const rowTwo = document.createElement("div");
    rowTwo.className = "field-row customs-row-two";

    const originField = document.createElement("label");
    originField.className = "field";
    originField.innerHTML =
      '<span>Origin of Items</span><div class="country-input-wrap customs-origin-wrap"><span class="customs-origin-flag"></span><input type="text" data-customs-item-field="originCountry" /></div>';
    const originInput = originField.querySelector("input");
    const originFlag = originField.querySelector(".customs-origin-flag");
    if (originInput) originInput.value = item.originCountry;
    if (originFlag) originFlag.innerHTML = getCountryIcon(item.originCountry);

    const hsField = document.createElement("label");
    hsField.className = "field";
    hsField.innerHTML = `<span>HS Tariff Code (optional)</span><input type="text" data-customs-item-field="hsCode" value="${escapeHtml(
      item.hsCode
    )}" />`;

    rowTwo.appendChild(originField);
    rowTwo.appendChild(hsField);

    card.appendChild(head);
    card.appendChild(note);
    card.appendChild(descriptionField);
    card.appendChild(rowThree);
    card.appendChild(rowTwo);
    customsItemsList.appendChild(card);
  });

  if (customsAddItemButton) {
    customsAddItemButton.disabled = state.customs.items.length >= CUSTOMS_MAX_ITEMS;
  }
}

function setCustomsGhostVisible(visible, options = {}) {
  const { push = true, replace = false } = options;
  customsGhostVisible = Boolean(visible);
  if (customsGhostVisible) {
    seedCustomsDefaults();
    syncCustomsInputsFromState();
    renderCustomsItems();
    updateCustomsScopeMeta();
  } else {
    setCustomsError("");
  }
  renderActiveBuilderPanel();
  if (push) {
    updateRoute(
      { view: "builder", step: state.step, customs: customsGhostVisible && state.step === 1 },
      { replace }
    );
  }
}

function maybeOpenCustomsGhostBeforeSelection() {
  if (state.step !== 1) return false;
  if (!requiresCustomsDeclaration()) return false;
  const signature = getCustomsRequirementSignature();
  if (state.customs.completedFor && state.customs.completedFor === signature) {
    return false;
  }
  setCustomsGhostVisible(true);
  return true;
}

function validateCustomsDeclaration() {
  syncCustomsStateFromInputs();
  const items = buildCustomsItemsPayload();
  state.customs.items = items.length ? items : [createEmptyCustomsItem()];

  let isValid = true;
  const hasSummary = Boolean(state.customs.summaryDescription);
  const hasShipmentType = Boolean(state.customs.shipmentType);
  setCustomsFieldInvalid(customsSummaryDescriptionInput, !hasSummary);
  setCustomsFieldInvalid(customsShipmentTypeInput, !hasShipmentType);
  if (!hasSummary || !hasShipmentType) {
    isValid = false;
  }

  state.customs.items.forEach((item, index) => {
    const quantity = Number(item.quantity);
    const weight = Number(item.weightKg);
    const value = Number(item.valueEur);
    const checks = {
      description: Boolean(item.description),
      quantity: Number.isFinite(quantity) && quantity > 0,
      weightKg: Number.isFinite(weight) && weight > 0,
      valueEur: Number.isFinite(value) && value >= 0,
      originCountry: Boolean(item.originCountry),
    };
    Object.entries(checks).forEach(([key, ok]) => {
      setCustomsItemFieldInvalid(index, key, !ok);
      if (!ok) isValid = false;
    });
  });

  if (!isValid) {
    setCustomsError("Complete required customs declaration fields before continuing.");
    return false;
  }
  setCustomsError("");
  state.customs.completedFor = getCustomsRequirementSignature();
  return true;
}

function resetCustomsDeclaration() {
  state.customs.summaryDescription = "";
  state.customs.shipmentType = "";
  state.customs.senderReference = "";
  state.customs.importerReference = "";
  state.customs.items = [createEmptyCustomsItem()];
  state.customs.completedFor = "";
  syncCustomsInputsFromState();
  renderCustomsItems();
  setCustomsError("");
}

function getQuantity() {
  if (state.csvMode) {
    return Math.max(1, state.csvRows.length || 1);
  }
  return Math.max(1, Number(state.quantity) || 1);
}

function getActiveLabel() {
  if (state.labels[state.activeLabelIndex]) {
    return state.labels[state.activeLabelIndex];
  }

  if (state.csvMode && state.csvRows.length) {
    const data = state.csvRows[state.activeLabelIndex] || state.csvRows[0];
    return {
      trackingId: state.info.trackingId,
      labelId: state.info.labelId,
      pdfUrl: currentPdfUrl,
      data,
    };
  }

  return {
    trackingId: state.info.trackingId,
    labelId: state.info.labelId,
    pdfUrl: currentPdfUrl,
    data: state.info,
  };
}

function formatAddress(name, street, city, stateCode, zip, country) {
  const line1 = name || "--";
  const line2 = street || "";
  const line3 = formatCityRegionPostal(city, stateCode, zip);
  const line4 = country || "";
  return [line1, line2, line3, line4].filter(Boolean).join("\n");
}

function goToStep(step, options = {}) {
  const { push = true, regenerate = true, replace = false } = options;
  if (step === state.step && push) return;
  state.step = step;
  customsGhostVisible = false;
  renderActiveBuilderPanel();

  stepperItems.forEach((item) => {
    const itemStep = Number(item.dataset.step);
    item.classList.toggle("is-active", itemStep === step);
    item.classList.toggle("is-complete", itemStep < step);
    item.classList.toggle("is-disabled", itemStep > step);
  });

  if (stepper) {
    const totalSteps = stepperItems.length;
    const progress = totalSteps > 1 ? (step - 1) / (totalSteps - 1) : 0;
    stepper.style.setProperty("--progress", progress);
  }

  if (step === 2) {
    ensureIdentifiers();
  }

  if (step === 3) {
    renderAccountProfile(currentUser);
    updatePayment();
  }

  if (step === 4 && regenerate) {
    generatePdf();
  }

  if (push) {
    updateRoute({ view: "builder", step, customs: false }, { replace });
  }
}

function ensureIdentifiers() {
  if (!inputMap.trackingId.value) {
    inputMap.trackingId.value = generateTracking();
  }
  if (!inputMap.labelId.value) {
    inputMap.labelId.value = generateLabelId();
  }
  syncInfoState();
}

function generateTracking() {
  return `TRK-${Math.floor(100000000 + Math.random() * 900000000)}`;
}

function generateLabelId() {
  return `LBL-${Math.floor(100000 + Math.random() * 900000)}`;
}

function syncInfoState() {
  state.info.senderName = inputMap.senderName.value.trim();
  state.info.senderStreet = inputMap.senderStreet.value.trim();
  state.info.senderCity = inputMap.senderCity.value.trim();
  state.info.senderState = inputMap.senderState.value.trim();
  state.info.senderZip = inputMap.senderZip.value.trim();
  state.info.recipientName = inputMap.recipientName.value.trim();
  state.info.recipientStreet = inputMap.recipientStreet.value.trim();
  state.info.recipientCountry = inputMap.recipientCountry.value.trim();
  state.info.recipientCity = inputMap.recipientCity.value.trim();
  state.info.recipientState = inputMap.recipientState.value.trim();
  state.info.recipientZip = inputMap.recipientZip.value.trim();
  state.info.packageWeight = inputMap.packageWeight.value.trim();
  state.info.packageDims = inputMap.packageDims.value.trim();
  state.info.trackingId = inputMap.trackingId.value.trim();
  state.info.labelId = inputMap.labelId.value.trim();

  invalidateCustomsCompletion();
  updateCountryFlag();
  updateSummary();
  updatePreview();
}

function isLabelFieldValid(id, value) {
  if (id === "packageWeight") {
    return Number(value) > 0;
  }
  return value !== "";
}

function validateFields(fieldIds, lookup, errorEl, validator) {
  let firstInvalid = null;
  fieldIds.forEach((id) => {
    const input = lookup[id];
    const value = input.value.trim();
    const isValid = validator(id, value);
    input.classList.toggle("is-invalid", !isValid);
    if (!isValid && !firstInvalid) {
      firstInvalid = input;
    }
  });

  if (firstInvalid) {
    if (errorEl) {
      errorEl.classList.add("is-visible");
    }
    firstInvalid.focus();
    return false;
  }

  if (errorEl) {
    errorEl.classList.remove("is-visible");
  }
  return true;
}

function validateLabelInfo() {
  if (state.csvMode) {
    const ok = validateCsvRows();
    if (!ok && labelError) {
      labelError.classList.add("is-visible");
    }
    if (ok && labelError) {
      labelError.classList.remove("is-visible");
    }
    return ok;
  }
  return validateFields(labelRequiredFields, inputMap, labelError, isLabelFieldValid);
}

Object.entries(inputMap).forEach(([_, input]) => {
  input.addEventListener("input", () => {
    syncInfoState();
    input.classList.remove("is-invalid");
    if (labelError) {
      labelError.classList.remove("is-visible");
    }
  });
});

if (authForm) {
  authForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    await signInWithPassword();
  });
}

if (authForgotPassword) {
  authForgotPassword.addEventListener("click", async () => {
    await resetPassword();
  });
}

if (authSignUp) {
  authSignUp.addEventListener("click", async () => {
    await signUpWithPassword();
  });
}

if (signOutButton) {
  signOutButton.addEventListener("click", async () => {
    if (!supabaseClient) return;
    await supabaseClient.auth.signOut();
  });
}

if (openAccountPageButton) {
  openAccountPageButton.addEventListener("click", async () => {
    setAccountPageVisible(true);
    await loadWarehouseSettings({ quiet: true });
    await loadGenerationHistory({ preferLatest: true });
  });
}

if (openReportsPageButton) {
  openReportsPageButton.addEventListener("click", async () => {
    setReportsPageVisible(true);
    await loadGenerationHistory({ preferLatest: true });
    renderReportsDashboard();
  });
}

if (closeAccountPageButton) {
  closeAccountPageButton.addEventListener("click", () => {
    setAccountPageVisible(false);
  });
}

if (closeReportsPageButton) {
  closeReportsPageButton.addEventListener("click", () => {
    setReportsPageVisible(false);
  });
}

if (warehouseAddButton) {
  warehouseAddButton.addEventListener("click", () => {
    if (!currentUser || warehouseSaving) return;
    if (warehouseRecords.length >= WAREHOUSE_MAX_COUNT) {
      setWarehouseStatus(`Maximum ${WAREHOUSE_MAX_COUNT} warehouses reached.`, {
        tone: "error",
      });
      return;
    }
    const newOrigin = buildNewWarehouseRecord();
    warehouseEnteringIds.add(newOrigin.id);
    warehouseRecords = [...warehouseRecords, newOrigin];
    renderWarehouseList();
    setWarehouseDirty(true);
  });
}

if (warehouseSaveButton) {
  warehouseSaveButton.addEventListener("click", async () => {
    await saveWarehouseSettings();
  });
}

if (warehouseList) {
  warehouseList.addEventListener("input", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) return;
    const field = target.dataset.warehouseField;
    if (!field) return;
    const card = target.closest("[data-warehouse-id]");
    if (!card) return;
    const index = findWarehouseIndexById(card.dataset.warehouseId);
    if (index === -1) return;
    warehouseRecords[index][field] = target.value.trim();

    if (field === "name") {
      const label = card.querySelector(".warehouse-card-name");
      if (label) {
        label.textContent = target.value.trim() || `Warehouse ${index + 1}`;
      }
    }

    if (field === "country") {
      const icon = target.parentElement?.querySelector(".warehouse-country-flag");
      if (icon) {
        icon.innerHTML = getCountryIcon(target.value.trim());
      }
    }

    setWarehouseDirty(true);
  });

  warehouseList.addEventListener("click", (event) => {
    const actionButton = event.target.closest("[data-warehouse-action]");
    if (!actionButton) return;
    const card = actionButton.closest("[data-warehouse-id]");
    if (!card) return;
    const warehouseId = card.dataset.warehouseId;
    const action = actionButton.dataset.warehouseAction;
    const index = findWarehouseIndexById(warehouseId);
    if (index === -1) return;

    if (action === "apply") {
      applyWarehouseToSender(warehouseRecords[index]);
      return;
    }

    if (action === "default") {
      if (!warehouseRecords[index].isDefault) {
        warehouseRecords = normalizeWarehouseRecords(
          warehouseRecords.map((origin) => ({
            ...origin,
            isDefault: origin.id === warehouseId,
          }))
        );
        renderWarehouseList();
        setWarehouseDirty(true);
      }
      return;
    }

    if (action === "remove") {
      if (warehouseRecords.length <= 1) {
        setWarehouseStatus("At least one warehouse origin is required.", {
          tone: "error",
        });
        return;
      }
      removeWarehouseWithAnimation(warehouseId);
    }
  });
}

[
  customsSummaryDescriptionInput,
  customsShipmentTypeInput,
  customsSenderReferenceInput,
  customsImporterReferenceInput,
].forEach((input) => {
  if (!input) return;
  const eventName = input.tagName === "SELECT" ? "change" : "input";
  input.addEventListener(eventName, () => {
    syncCustomsStateFromInputs();
    setCustomsFieldInvalid(input, false);
    setCustomsError("");
  });
});

if (customsItemsList) {
  customsItemsList.addEventListener("input", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) return;
    const card = target.closest("[data-customs-item-index]");
    if (!card) return;
    const index = Number(card.dataset.customsItemIndex);
    if (Number.isNaN(index) || !state.customs.items[index]) return;
    const key = target.dataset.customsItemField;
    if (!key) return;
    state.customs.items[index][key] = target.value.trim();
    target.classList.remove("is-invalid");
    if (key === "originCountry") {
      const flag = target.parentElement?.querySelector(".customs-origin-flag");
      if (flag) {
        flag.innerHTML = getCountryIcon(target.value.trim());
      }
    }
    setCustomsError("");
  });

  customsItemsList.addEventListener("click", (event) => {
    const removeButton = event.target.closest("[data-customs-item-action='remove']");
    if (!removeButton) return;
    const index = Number(removeButton.dataset.customsItemIndex);
    if (Number.isNaN(index)) return;
    if (state.customs.items.length <= 1) return;
    state.customs.items.splice(index, 1);
    renderCustomsItems();
    setCustomsError("");
  });
}

if (customsAddItemButton) {
  customsAddItemButton.addEventListener("click", () => {
    if (state.customs.items.length >= CUSTOMS_MAX_ITEMS) return;
    state.customs.items.push(createEmptyCustomsItem());
    renderCustomsItems();
    setCustomsError("");
  });
}

if (customsBackButton) {
  customsBackButton.addEventListener("click", () => {
    setCustomsGhostVisible(false, { replace: true });
  });
}

if (customsContinueButton) {
  customsContinueButton.addEventListener("click", () => {
    if (!validateCustomsDeclaration()) return;
    setCustomsGhostVisible(false, { push: false });
    goToStep(2);
  });
}

reportsRangeButtons.forEach((button) => {
  button.addEventListener("click", () => {
    const token = button.dataset.reportRange;
    if (token === "custom") {
      activateCustomReportsRangePicker();
      return;
    }
    const nextRange = Number(token);
    if (![7, 30, 90, 365].includes(nextRange)) return;
    reportsRangeMode = "preset";
    reportsRangeDays = nextRange;
    setReportsCustomRangeVisible(false);
    if (reportsStartDate) reportsStartDate.classList.remove("is-invalid");
    if (reportsEndDate) reportsEndDate.classList.remove("is-invalid");
    renderReportsDashboard();
  });
});

if (reportsApplyCustom) {
  reportsApplyCustom.addEventListener("click", () => {
    applyCustomReportsRange();
  });
}

if (reportsStartDate && reportsEndDate) {
  [reportsStartDate, reportsEndDate].forEach((input) => {
    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        applyCustomReportsRange();
      }
    });
  });
}

if (accountHistoryList) {
  accountHistoryList.addEventListener("click", (event) => {
    const target = event.target.closest("[data-history-index]");
    if (!target) return;
    const index = Number(target.dataset.historyIndex);
    if (Number.isNaN(index)) return;
    selectAccountRecord(index);
  });
}

if (accountBatchList) {
  accountBatchList.addEventListener("click", (event) => {
    const target = event.target.closest("[data-account-label-index]");
    if (!target) return;
    const index = Number(target.dataset.accountLabelIndex);
    if (Number.isNaN(index)) return;
    selectAccountLabel(index);
  });
}

if (accountDownloadPdf) {
  accountDownloadPdf.addEventListener("click", (event) => {
    event.preventDefault();
    downloadActiveAccountLabelPdf();
  });
}

if (openReceiptModalButton) {
  openReceiptModalButton.addEventListener("click", () => {
    if (!accountActiveRecord) return;
    renderReceiptDetails(accountActiveRecord);
    setReceiptModalOpen(true);
  });
}

if (receiptModalClose) {
  receiptModalClose.addEventListener("click", (event) => {
    event.preventDefault();
    setReceiptModalOpen(false);
  });
}

if (receiptModal) {
  receiptModal.addEventListener("click", (event) => {
    if (event.target === receiptModal) {
      setReceiptModalOpen(false);
    }
  });
}

if (openRestrictedGoodsModalButton) {
  openRestrictedGoodsModalButton.addEventListener("click", (event) => {
    event.preventDefault();
    setRestrictedGoodsModalOpen(true);
  });
}

if (restrictedGoodsModalClose) {
  restrictedGoodsModalClose.addEventListener("click", (event) => {
    event.preventDefault();
    setRestrictedGoodsModalOpen(false);
  });
}

if (restrictedGoodsModal) {
  restrictedGoodsModal.addEventListener("click", (event) => {
    if (event.target === restrictedGoodsModal) {
      setRestrictedGoodsModalOpen(false);
    }
  });
}

if (receiptDownloadPdf) {
  receiptDownloadPdf.addEventListener("click", (event) => {
    event.preventDefault();
    downloadReceiptPdfFile();
  });
}

document.addEventListener("keydown", (event) => {
  if (event.key !== "ArrowUp" && event.key !== "ArrowDown") return;
  if (isTextEntryElement(event.target)) return;
  if (receiptModal && !receiptModal.classList.contains("is-closed")) return;
  if (restrictedGoodsModal && !restrictedGoodsModal.classList.contains("is-closed")) return;

  const direction = event.key === "ArrowDown" ? 1 : -1;
  const isAccountOpen = accountPageSection && !accountPageSection.classList.contains("is-hidden");
  const isBuilderOpen = builderPage && !builderPage.classList.contains("is-hidden");

  if (isAccountOpen) {
    const active = document.activeElement;
    const inBatch =
      active?.closest?.("#accountBatchList") || event.target?.closest?.("#accountBatchList");
    const handled = inBatch
      ? navigateAccountBatch(direction)
      : navigateAccountHistory(direction);
    if (handled) {
      event.preventDefault();
    }
    return;
  }

  if (isBuilderOpen && state.step === 4 && state.labels.length > 1) {
    const handled = navigateBuilderBatch(direction);
    if (handled) {
      event.preventDefault();
    }
  }
});

function autoFill() {
  clearBatchState();
  setCsvMode(false);
  setCsvEditMode(false);
  state.csvRows = [];
  if (labelError) {
    labelError.classList.remove("is-visible");
  }
  const profile = sampleProfiles[Math.floor(Math.random() * sampleProfiles.length)];
  inputMap.senderName.value = profile.sender.name;
  inputMap.senderStreet.value = profile.sender.street;
  inputMap.senderCity.value = profile.sender.city;
  inputMap.senderState.value = profile.sender.state;
  inputMap.senderZip.value = profile.sender.zip;

  inputMap.recipientName.value = profile.recipient.name;
  inputMap.recipientStreet.value = profile.recipient.street;
  inputMap.recipientCountry.value = profile.recipient.country || "France";
  inputMap.recipientCity.value = profile.recipient.city;
  inputMap.recipientState.value = profile.recipient.state;
  inputMap.recipientZip.value = profile.recipient.zip;

  inputMap.packageWeight.value = (Math.random() * 2.7 + 0.3).toFixed(1);
  inputMap.packageDims.value = `${Math.floor(Math.random() * 17 + 16)} x ${
    Math.floor(Math.random() * 12 + 12)
  } x ${Math.floor(Math.random() * 8 + 4)}`;

  inputMap.trackingId.value = generateTracking();
  inputMap.labelId.value = generateLabelId();

  syncInfoState();
}

if (autoFillButton) {
  autoFillButton.addEventListener("click", (event) => {
    event.preventDefault();
    autoFill();
  });
}

if (autoCsvButton) {
  autoCsvButton.addEventListener("click", (event) => {
    event.preventDefault();
    clearBatchState();
    state.csvRows = generateCsvRows(3);
    setCsvMode(true);
    setCsvEditMode(false);
    renderCsvTable();
    if (labelError) {
      labelError.classList.remove("is-visible");
    }
    updatePreview();
    updateSummary();
    updatePayment();
  });
}

if (csvEditToggle) {
  csvEditToggle.addEventListener("click", (event) => {
    event.preventDefault();
    if (!state.csvMode) return;
    setCsvEditMode(!state.csvEditable);
  });
}

quantityInput.addEventListener("input", () => {
  if (state.csvMode) {
    updateQuantityFromCsv();
    return;
  }
  const value = Number.parseInt(quantityInput.value, 10);
  state.quantity = Number.isNaN(value) ? 1 : Math.max(1, value);
  quantityInput.value = state.quantity;
  updateSummary();
  updatePayment();
  if (state.step === 4) {
    generatePdf();
  }
});

function handleNext(step) {
  if (step === 2 && !validateLabelInfo()) {
    return;
  }
  if (step === 2 && maybeOpenCustomsGhostBeforeSelection()) {
    return;
  }
  goToStep(step);
}

document.querySelectorAll("[data-next]").forEach((button) => {
  button.addEventListener("click", () => {
    const nextStep = Number(button.dataset.next);
    handleNext(nextStep);
  });
});

document.querySelectorAll("[data-back]").forEach((button) => {
  button.addEventListener("click", () => {
    const prevStep = Number(button.dataset.back);
    goToStep(prevStep);
  });
});

if (payButton) {
  payButton.addEventListener("click", (event) => {
    event.preventDefault();

    payButton.disabled = true;
    const originalMarkup = payButton.innerHTML;
    payButton.innerHTML = "Preparing Label...";

    setTimeout(() => {
      payButton.disabled = false;
      payButton.innerHTML = originalMarkup;
      goToStep(4);
    }, 900);
  });
}

function pdfEscape(value) {
  return value.replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)");
}

function buildPdfViewerUrl(pdfUrl, pageNumber = 1) {
  if (!pdfUrl) return "";
  const safePage = Math.max(1, Number(pageNumber) || 1);
  return `${pdfUrl}#page=${safePage}`;
}

function buildPdf(lines, options = {}) {
  const pageWidth = Number(options.pageWidth) || 288;
  const pageHeight = Number(options.pageHeight) || 432;
  const marginX = Number(options.marginX) || 20;
  const marginTop = Number(options.marginTop) || 32;
  const marginBottom = Number(options.marginBottom) || 24;
  const fontSize = Number(options.fontSize) || 12;
  const lineHeight = Number(options.lineHeight) || 16;

  const usableHeight = pageHeight - marginTop - marginBottom;
  const linesPerPage = Math.max(1, Math.floor(usableHeight / lineHeight));
  const paginateLines = (sourceLines) => {
    const safeLines = Array.isArray(sourceLines)
      ? sourceLines.map((line) => String(line ?? ""))
      : [];
    const chunked = [];
    for (let i = 0; i < safeLines.length; i += linesPerPage) {
      chunked.push(safeLines.slice(i, i + linesPerPage));
    }
    return chunked;
  };

  const pages = [];
  if (Array.isArray(lines) && lines.length && lines.every((item) => Array.isArray(item))) {
    lines.forEach((pageLines) => {
      const pageChunks = paginateLines(pageLines);
      if (!pageChunks.length) {
        pages.push([""]);
      } else {
        pages.push(...pageChunks);
      }
    });
  } else {
    pages.push(...paginateLines(lines));
  }

  if (!pages.length) {
    pages.push([""]);
  }

  const objects = [];
  const fontObjectId = 3;
  let nextObjectId = 4;
  const pageIds = [];
  const contentEntries = [];

  pages.forEach((pageLines) => {
    const contentObjectId = nextObjectId;
    const pageObjectId = nextObjectId + 1;
    nextObjectId += 2;
    pageIds.push(pageObjectId);

    const startY = pageHeight - marginTop;
    const contentLines = ["BT", `/F1 ${fontSize} Tf`, "0 0 0 rg", `${marginX} ${startY} Td`];
    pageLines.forEach((line, index) => {
      if (index > 0) {
        contentLines.push(`0 -${lineHeight} Td`);
      }
      contentLines.push(`(${pdfEscape(line)}) Tj`);
    });
    contentLines.push("ET");
    const content = contentLines.join("\n");
    contentEntries.push({ contentObjectId, pageObjectId, content });
  });

  objects.push({ id: 1, value: "1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj" });
  objects.push({
    id: 2,
    value: `2 0 obj << /Type /Pages /Kids [${pageIds
      .map((id) => `${id} 0 R`)
      .join(" ")}] /Count ${pageIds.length} >> endobj`,
  });
  objects.push({
    id: 3,
    value: "3 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj",
  });

  contentEntries.forEach(({ contentObjectId, content }) => {
    const stream = `<< /Length ${content.length} >> stream\n${content}\nendstream`;
    objects.push({ id: contentObjectId, value: `${contentObjectId} 0 obj ${stream} endobj` });
  });

  contentEntries.forEach(({ contentObjectId, pageObjectId }) => {
    objects.push({
      id: pageObjectId,
      value: `${pageObjectId} 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pageWidth} ${pageHeight}] /Contents ${contentObjectId} 0 R /Resources << /Font << /F1 ${fontObjectId} 0 R >> >> >> endobj`,
    });
  });

  objects.sort((a, b) => a.id - b.id);
  const maxObjectId = objects.reduce((maxId, entry) => Math.max(maxId, entry.id), 0);

  let pdf = "%PDF-1.4\n";
  const offsets = Array(maxObjectId + 1).fill(0);
  objects.forEach((entry) => {
    offsets[entry.id] = pdf.length;
    pdf += `${entry.value}\n`;
  });

  const xrefStart = pdf.length;
  pdf += `xref\n0 ${maxObjectId + 1}\n`;
  pdf += "0000000000 65535 f \n";
  for (let i = 1; i <= maxObjectId; i += 1) {
    const offset = offsets[i] || 0;
    const marker = offset > 0 ? "n" : "f";
    pdf += `${String(offset).padStart(10, "0")} 00000 ${marker} \n`;
  }
  pdf += `trailer << /Size ${maxObjectId + 1} /Root 1 0 R >>\nstartxref\n${xrefStart}\n%%EOF`;

  return new Blob([pdf], { type: "application/pdf" });
}

function createLabelBatchPdfBlob(labels, serviceType = state.selection.type) {
  const labelPages = Array.isArray(labels)
    ? labels.map((label) => buildLabelLines(label, serviceType))
    : [];
  return buildPdf(labelPages);
}

function formatCityRegionPostal(city, region, postalCode) {
  const place = [city, region].filter(Boolean).join(", ");
  return [postalCode, place].filter(Boolean).join(" ");
}

function buildLabelLines(label, serviceType = state.selection.type) {
  const data = label.data || state.info;
  return [
    "SHIPR LABEL",
    `Service: ${serviceType}`,
    `Tracking: ${label.trackingId}`,
    `Label ID: ${label.labelId}`,
    "",
    "FROM:",
    data.senderName,
    data.senderStreet,
    formatCityRegionPostal(data.senderCity, data.senderState, data.senderZip),
    "",
    "TO:",
    data.recipientName,
    data.recipientStreet,
    formatCityRegionPostal(data.recipientCity, data.recipientState, data.recipientZip),
    data.recipientCountry ? data.recipientCountry : null,
    "",
    `Weight: ${data.packageWeight} kg`,
    `Dims: ${data.packageDims} cm`,
  ].filter((line) => line !== undefined && line !== null);
}


function revokePdfUrls() {
  const urls = new Set();
  state.labels.forEach((label) => {
    if (label.pdfUrl) {
      urls.add(label.pdfUrl);
    }
  });
  if (currentBatchPdfUrl) {
    urls.add(currentBatchPdfUrl);
  }
  urls.forEach((url) => URL.revokeObjectURL(url));
  currentBatchPdfUrl = "";
}

function clearBatchState() {
  revokePdfUrls();
  state.labels = [];
  state.activeLabelIndex = 0;
  currentPdfUrl = "";
  if (batchPanel) {
    batchPanel.classList.add("is-hidden");
  }
  if (batchList) {
    batchList.innerHTML = "";
  }
  if (batchPreview) {
    batchPreview.classList.add("is-single");
  }
  if (pdfFrame) {
    pdfFrame.src = "";
  }
}

function renderBatchList() {
  const count = getQuantity();
  if (!batchPanel || !batchList) return;

  if (count <= 1) {
    batchPanel.classList.add("is-hidden");
    batchList.innerHTML = "";
    if (batchPreview) {
      batchPreview.classList.add("is-single");
    }
    return;
  }

  batchPanel.classList.remove("is-hidden");
  if (batchPreview) {
    batchPreview.classList.remove("is-single");
  }
  batchList.innerHTML = "";

  state.labels.forEach((label, index) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `batch-item${index === state.activeLabelIndex ? " is-active" : ""}`;
    button.dataset.batchIndex = String(index);
    button.style.animationDelay = `${index * 0.04}s`;
    button.innerHTML = `
      <div class="batch-index">Label ${label.index}</div>
      <div class="batch-meta mono">${label.trackingId}</div>
      <div class="batch-meta mono">${label.labelId}</div>
    `;
    button.addEventListener("click", () => {
      selectBatchLabel(index);
    });
    batchList.appendChild(button);
  });
}

function selectBatchLabel(index) {
  state.activeLabelIndex = index;
  const activeLabel = getActiveLabel();
  currentPdfUrl = activeLabel.pdfUrl || "";
  pdfFrame.src = buildPdfViewerUrl(currentPdfUrl, activeLabel.pdfPage || index + 1);

  if (batchList) {
    Array.from(batchList.children).forEach((child, idx) => {
      child.classList.toggle("is-active", idx === index);
    });
  }

  downloadPdf.onclick = () => {
    if (!currentBatchPdfUrl) return;
    triggerFileDownload(currentBatchPdfUrl, "labels-batch.pdf");
  };

  updateSummary();
  updatePreview();
}

function isTextEntryElement(element) {
  if (!element) return false;
  const tag = element.tagName;
  return (
    tag === "INPUT" ||
    tag === "TEXTAREA" ||
    tag === "SELECT" ||
    Boolean(element.isContentEditable)
  );
}

function focusHistoryItem(index) {
  const target = accountHistoryList?.querySelector(`[data-history-index="${index}"]`);
  if (target) {
    target.focus({ preventScroll: true });
  }
}

function focusAccountBatchItem(index) {
  const target = accountBatchList?.querySelector(
    `[data-account-label-index="${index}"]`
  );
  if (target) {
    target.focus({ preventScroll: true });
  }
}

function focusBuilderBatchItem(index) {
  const target = batchList?.querySelector(`[data-batch-index="${index}"]`);
  if (target) {
    target.focus({ preventScroll: true });
  }
}

function navigateAccountHistory(direction) {
  if (!historyRecords.length) return false;
  let index =
    accountActiveHistoryIndex >= 0
      ? accountActiveHistoryIndex
      : findFirstPreviewableHistoryIndex();
  if (index < 0) return false;

  let next = index + direction;
  while (next >= 0 && next < historyRecords.length) {
    if (historyRecordHasLabels(historyRecords[next])) {
      selectAccountRecord(next);
      requestAnimationFrame(() => focusHistoryItem(next));
      return true;
    }
    next += direction;
  }
  return true;
}

function navigateAccountBatch(direction) {
  if (accountLabels.length <= 1) return false;
  const next = Math.max(
    0,
    Math.min(accountLabels.length - 1, accountActiveLabelIndex + direction)
  );
  if (next === accountActiveLabelIndex) return true;
  selectAccountLabel(next);
  requestAnimationFrame(() => focusAccountBatchItem(next));
  return true;
}

function navigateBuilderBatch(direction) {
  if (state.labels.length <= 1) return false;
  const next = Math.max(0, Math.min(state.labels.length - 1, state.activeLabelIndex + direction));
  if (next === state.activeLabelIndex) return true;
  selectBatchLabel(next);
  requestAnimationFrame(() => focusBuilderBatchItem(next));
  return true;
}

function generatePdf() {
  if (!state.csvMode) {
    syncInfoState();
  }
  ensureIdentifiers();
  revokePdfUrls();
  const dataList =
    state.csvMode && state.csvRows.length
      ? state.csvRows
      : Array.from({ length: getQuantity() }, () => ({ ...state.info }));

  state.labels = dataList.map((data, index) => {
    const labelId =
      index === 0 ? state.info.labelId || generateLabelId() : generateLabelId();
    const trackingId =
      index === 0 ? state.info.trackingId || generateTracking() : generateTracking();
    return {
      index: index + 1,
      labelId,
      trackingId,
      data,
      pdfUrl: "",
      pdfPage: index + 1,
    };
  });
  const batchBlob = createLabelBatchPdfBlob(state.labels, state.selection.type);
  currentBatchPdfUrl = URL.createObjectURL(batchBlob);
  state.labels = state.labels.map((label) => ({
    ...label,
    pdfUrl: currentBatchPdfUrl,
  }));

  renderBatchList();
  selectBatchLabel(0);
  persistGenerationHistory();
}

function resetAll() {
  clearBatchState();
  resetCustomsDeclaration();
  customsGhostVisible = false;
  state.quantity = 1;
  state.csvMode = false;
  state.csvEditable = false;
  state.csvRows = [];
  if (labelForm) {
    labelForm.classList.remove("is-hidden");
  }
  if (csvSection) {
    csvSection.classList.remove("is-visible");
  }
  if (csvEditToggle) {
    csvEditToggle.disabled = true;
    csvEditToggle.classList.remove("is-active");
    const label = csvEditToggle.querySelector("span");
    if (label) label.textContent = "Edit";
  }
  if (quantityInput) {
    quantityInput.value = 1;
    quantityInput.disabled = false;
  }
  Object.values(inputMap).forEach((input) => {
    if (!input.hasAttribute("readonly")) {
      input.value = "";
    }
  });

  inputMap.trackingId.value = generateTracking();
  inputMap.labelId.value = generateLabelId();
  syncInfoState();
  selectLabel(labelCards[0]);
  updatePayment();
}

startOver.addEventListener("click", () => {
  resetAll();
  goToStep(1);
});

// ── Provider dropdown ──

if (providerTrigger) {
  providerTrigger.addEventListener("click", (e) => {
    e.stopPropagation();
    providerDropdown.classList.toggle("is-open");
  });
}

document.addEventListener("click", (e) => {
  if (providerDropdown && !providerDropdown.contains(e.target)) {
    providerDropdown.classList.remove("is-open");
  }
});

document.querySelectorAll(".provider-option").forEach((opt) => {
  opt.addEventListener("click", async () => {
    const provider = opt.dataset.provider;
    providerDropdown.classList.remove("is-open");
    if (provider === "shopify") {
      await handleShopifyProviderAction();
      return;
    }

    autoFill();
    const triggerText = providerTrigger?.querySelector("span");
    if (triggerText) {
      triggerText.textContent = `Imported from ${opt.textContent.trim()}`;
      window.setTimeout(() => {
        triggerText.textContent = "Import from provider";
      }, 2000);
    }
  });
});

// ── CSV upload modal ──

function setCsvMapError(message = "") {
  if (!csvMapError) return;
  csvMapError.textContent = message;
  csvMapError.classList.toggle("is-visible", Boolean(message));
}

function setCsvModalStep(step, options = {}) {
  const { animate = true } = options;
  const nextStep = step === "mapping" ? "mapping" : "upload";
  const isMapping = nextStep === "mapping";
  const shouldAnimate = animate && !prefersReducedMotion();
  const modalTitle = document.querySelector(".csv-modal-title");

  if (!csvUploadStep || !csvMappingStep) {
    if (modalTitle) {
      modalTitle.textContent = isMapping ? "Map CSV columns" : "Upload CSV file";
    }
    csvModalStepState = nextStep;
    return;
  }

  if (!shouldAnimate) {
    csvUploadStep.classList.toggle("is-active", !isMapping);
    csvMappingStep.classList.toggle("is-active", isMapping);
    if (modalTitle) {
      modalTitle.textContent = isMapping ? "Map CSV columns" : "Upload CSV file";
    }
    csvModalStepState = nextStep;
    return;
  }

  const currentStep = csvModalStepState;
  if (currentStep === nextStep) {
    if (modalTitle) {
      modalTitle.textContent = isMapping ? "Map CSV columns" : "Upload CSV file";
    }
    return;
  }

  const currentEl = currentStep === "mapping" ? csvMappingStep : csvUploadStep;
  const nextEl = nextStep === "mapping" ? csvMappingStep : csvUploadStep;
  const transitionToken = ++csvModalStepTransitionToken;

  currentEl.classList.remove("is-active");
  window.setTimeout(() => {
    if (transitionToken !== csvModalStepTransitionToken) return;
    nextEl.classList.add("is-active");
    csvModalStepState = nextStep;
  }, CSV_MODAL_STEP_SWITCH_DELAY_MS);

  if (modalTitle) {
    window.setTimeout(() => {
      if (transitionToken !== csvModalStepTransitionToken) return;
      modalTitle.textContent = isMapping ? "Map CSV columns" : "Upload CSV file";
    }, Math.floor(CSV_MODAL_STEP_SWITCH_DELAY_MS * 0.85));
  }
}

function resetCsvMappingFlow() {
  csvMappingDraft = null;
  csvModalStepTransitionToken += 1;
  if (csvMappingBody) {
    csvMappingBody.innerHTML = "";
  }
  if (csvMapMeta) {
    csvMapMeta.textContent = "Auto-mapped columns ready for review";
  }
  setCsvMapError("");
  setCsvModalStep("upload", { animate: false });
}

function syncCsvFileInputState() {
  if (!csvFileInput || !csvModal) return;
  csvFileInput.disabled = csvModal.classList.contains("is-closed");
}

function openCsvModal() {
  if (!csvModal) return;
  resetCsvMappingFlow();
  csvModal.classList.remove("is-closed");
  syncCsvFileInputState();
}

function closeCsvModal() {
  if (!csvModal) return;
  csvModal.classList.add("is-closed");
  syncCsvFileInputState();
  window.setTimeout(() => {
    if (csvModal.classList.contains("is-closed")) {
      resetCsvMappingFlow();
    }
  }, 360);
}

if (csvUploadTrigger) {
  csvUploadTrigger.addEventListener("click", () => {
    openCsvModal();
  });
}

if (csvModalClose) {
  csvModalClose.addEventListener("click", () => {
    closeCsvModal();
  });
}

if (csvModal) {
  csvModal.addEventListener("click", (e) => {
    if (e.target === csvModal) {
      closeCsvModal();
    }
  });
}

if (csvMapBack) {
  csvMapBack.addEventListener("click", () => {
    setCsvMapError("");
    setCsvModalStep("upload");
  });
}

if (csvDropzone) {
  csvDropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    csvDropzone.classList.add("is-dragover");
  });

  csvDropzone.addEventListener("dragleave", () => {
    csvDropzone.classList.remove("is-dragover");
  });

  csvDropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    csvDropzone.classList.remove("is-dragover");
    const file = e.dataTransfer.files[0];
    if (file) handleCsvFile(file);
  });
}

if (csvFileInput) {
  csvFileInput.addEventListener("change", () => {
    const file = csvFileInput.files[0];
    if (file) handleCsvFile(file);
    csvFileInput.value = "";
  });
}

if (csvMapApply) {
  csvMapApply.addEventListener("click", () => {
    applyCsvMapping();
  });
}

syncCsvFileInputState();

function handleCsvFile(file) {
  if (!/\.csv$/i.test(file.name || "")) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    const text = String(e.target.result || "");
    const analysis = analyzeCsvText(text);
    if (!analysis) {
      setCsvModalStep("mapping");
      setCsvMapError("Could not parse this CSV. Please check the file format and try again.");
      if (csvMappingBody) {
        csvMappingBody.innerHTML = "";
      }
      if (csvMapMeta) {
        csvMapMeta.textContent = "No rows detected";
      }
      return;
    }

    csvMappingDraft = {
      fileName: file.name || "upload.csv",
      analysis,
      mapping: { ...analysis.autoColumnIndices },
    };
    renderCsvMappingTable();
    setCsvMapError("");
    setCsvModalStep("mapping");
  };
  reader.readAsText(file);
}

function normalizeCsvHeaderToken(value) {
  return String(value || "")
    .replace(/^\uFEFF/, "")
    .trim()
    .toLowerCase()
    .replace(/["']/g, "")
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function countDelimiterOutsideQuotes(line, delimiter) {
  let inQuotes = false;
  let count = 0;
  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];
    if (char === '"') {
      if (inQuotes && line[i + 1] === '"') {
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }
    if (!inQuotes && char === delimiter) {
      count += 1;
    }
  }
  return count;
}

function detectCsvDelimiter(text) {
  const sampleLines = String(text || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .slice(0, 8);
  if (!sampleLines.length) return ",";

  const candidates = [",", ";", "\t", "|"];
  let best = ",";
  let bestScore = -1;

  candidates.forEach((delimiter) => {
    const score = sampleLines.reduce(
      (total, line) => total + countDelimiterOutsideQuotes(line, delimiter),
      0
    );
    if (score > bestScore) {
      bestScore = score;
      best = delimiter;
    }
  });

  return bestScore <= 0 ? "," : best;
}

function parseCsvMatrix(text, delimiter) {
  const rows = [];
  let row = [];
  let value = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];

    if (char === '"') {
      if (inQuotes && text[i + 1] === '"') {
        value += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (!inQuotes && char === delimiter) {
      row.push(value);
      value = "";
      continue;
    }

    if (!inQuotes && (char === "\n" || char === "\r")) {
      if (char === "\r" && text[i + 1] === "\n") {
        i += 1;
      }
      row.push(value);
      if (row.some((cell) => String(cell || "").trim() !== "")) {
        rows.push(row);
      }
      row = [];
      value = "";
      continue;
    }

    value += char;
  }

  row.push(value);
  if (row.some((cell) => String(cell || "").trim() !== "")) {
    rows.push(row);
  }

  return rows;
}

function findHeaderIndexByAliases(headerMap, aliases) {
  if (!headerMap || !Array.isArray(aliases)) return -1;
  for (const alias of aliases) {
    const token = normalizeCsvHeaderToken(alias);
    if (!token) continue;
    const index = headerMap.get(token);
    if (Number.isInteger(index)) {
      return index;
    }
  }
  return -1;
}

function getCsvValueAt(row, index) {
  if (!Array.isArray(row) || index < 0 || index >= row.length) return "";
  return String(row[index] || "").trim();
}

function joinCsvParts(parts) {
  return parts
    .map((part) => String(part || "").trim())
    .filter(Boolean)
    .join(" ");
}

function joinCsvAddress(line1, line2) {
  if (!line1 && !line2) return "";
  if (!line2) return line1;
  if (!line1) return line2;
  return `${line1}, ${line2}`;
}

function composeCsvDimensions(existingDims, lengthValue, widthValue, heightValue) {
  if (existingDims) return existingDims;
  if (!lengthValue || !widthValue || !heightValue) return "";
  return `${lengthValue} x ${widthValue} x ${heightValue}`;
}

function buildCsvHeaderMap(headers) {
  const headerMap = new Map();
  headers.forEach((header, index) => {
    const normalized = normalizeCsvHeaderToken(header);
    if (normalized && !headerMap.has(normalized)) {
      headerMap.set(normalized, index);
    }
  });
  return headerMap;
}

function parseAccountBillingAddress(value) {
  const text = String(value || "").trim();
  if (!text) {
    return { street: "", city: "", state: "", zip: "", country: "" };
  }

  const parts = text.split(",").map((part) => part.trim()).filter(Boolean);
  const street = parts[0] || "";
  const city = parts[1] || "";
  let state = "";
  let zip = "";
  const regionAndZip = parts[2] || "";
  if (regionAndZip) {
    const match = regionAndZip.match(/^(.+?)\s+([A-Za-z0-9-]{3,})$/);
    if (match) {
      state = match[1].trim();
      zip = match[2].trim();
    } else {
      state = regionAndZip;
    }
  }
  const country = parts[3] || "";
  return { street, city, state, zip, country };
}

function getDefaultWarehouseSenderDefaults() {
  const defaultOrigin = warehouseRecords.find((origin) => origin.isDefault) || warehouseRecords[0];
  if (!defaultOrigin) {
    return {
      senderName: "",
      senderStreet: "",
      senderCity: "",
      senderState: "",
      senderZip: "",
    };
  }
  return {
    senderName: String(defaultOrigin.senderName || defaultOrigin.name || "").trim(),
    senderStreet: String(defaultOrigin.street || "").trim(),
    senderCity: String(defaultOrigin.city || "").trim(),
    senderState: String(defaultOrigin.region || "").trim(),
    senderZip: String(defaultOrigin.postalCode || "").trim(),
  };
}

function getCsvSenderDefaults() {
  const direct = {
    senderName: String(inputMap.senderName?.value || state.info.senderName || "").trim(),
    senderStreet: String(inputMap.senderStreet?.value || state.info.senderStreet || "").trim(),
    senderCity: String(inputMap.senderCity?.value || state.info.senderCity || "").trim(),
    senderState: String(inputMap.senderState?.value || state.info.senderState || "").trim(),
    senderZip: String(inputMap.senderZip?.value || state.info.senderZip || "").trim(),
  };

  const profile = buildMockAccountProfile(currentUser);
  const billing = parseAccountBillingAddress(profile?.billingAddress);
  const fromWarehouse = getDefaultWarehouseSenderDefaults();
  const fromProfile = {
    senderName: String(profile?.companyName || profile?.contactName || "").trim(),
    senderStreet: billing.street,
    senderCity: billing.city,
    senderState: billing.state,
    senderZip: billing.zip,
  };

  return {
    senderName: direct.senderName || fromWarehouse.senderName || fromProfile.senderName,
    senderStreet: direct.senderStreet || fromWarehouse.senderStreet || fromProfile.senderStreet,
    senderCity: direct.senderCity || fromWarehouse.senderCity || fromProfile.senderCity,
    senderState: direct.senderState || fromWarehouse.senderState || fromProfile.senderState,
    senderZip: direct.senderZip || fromWarehouse.senderZip || fromProfile.senderZip,
  };
}

function parseCsvNumber(value) {
  const normalized = String(value || "")
    .trim()
    .replace(",", ".")
    .replace(/[^\d.+-]/g, "");
  const numeric = Number(normalized);
  return Number.isFinite(numeric) ? numeric : null;
}

function normalizeCsvWeight(value, sourceHeaderToken) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const numeric = parseCsvNumber(raw);
  if (!Number.isFinite(numeric) || numeric <= 0) return raw;

  const token = normalizeCsvHeaderToken(sourceHeaderToken);
  let kilograms = numeric;

  if (/(_|^)oz(s)?(_|$)/.test(token) || token.includes("ounce")) {
    kilograms = numeric * 0.0283495;
  } else if (/(_|^)kg(_|$)/.test(token)) {
    kilograms = numeric;
  } else if (token === "grams" || /(_|^)g(rams)?(_|$)/.test(token)) {
    kilograms = numeric / 1000;
  } else if (/(_|^)lb(s)?(_|$)/.test(token)) {
    kilograms = numeric * 0.45359237;
  }

  return Number(kilograms.toFixed(2)).toString();
}

function getCsvColumnLabel(key) {
  return csvColumns.find((column) => column.key === key)?.label || key;
}

function getCsvColumnsRequiredFirst() {
  const required = csvColumns.filter((column) => CSV_REQUIRED_FIELDS.has(column.key));
  const optional = csvColumns.filter((column) => !CSV_REQUIRED_FIELDS.has(column.key));
  return [...required, ...optional];
}

function detectAutoCsvColumnIndices(headerMap) {
  const columnIndices = {};
  csvColumns.forEach((column) => {
    const aliases = [
      ...(CSV_HEADER_ALIASES[column.key] || []),
      column.key,
      normalizeCsvHeaderToken(column.label),
    ];
    columnIndices[column.key] = findHeaderIndexByAliases(headerMap, aliases);
  });
  return columnIndices;
}

function normalizeAutoCsvColumnIndices(autoColumnIndices) {
  const used = new Set();
  const normalized = { ...autoColumnIndices };
  getCsvColumnsRequiredFirst().forEach((column) => {
    const key = column.key;
    const mappedIndex = Number(normalized[key]);
    if (!Number.isInteger(mappedIndex) || mappedIndex < 0) {
      normalized[key] = -1;
      return;
    }
    if (used.has(mappedIndex)) {
      normalized[key] = -1;
      return;
    }
    used.add(mappedIndex);
    normalized[key] = mappedIndex;
  });
  return normalized;
}

function detectCsvExtraIndices(headerMap) {
  return {
    senderFirst: findHeaderIndexByAliases(headerMap, CSV_NAME_PART_ALIASES.senderFirst),
    senderLast: findHeaderIndexByAliases(headerMap, CSV_NAME_PART_ALIASES.senderLast),
    recipientFirst: findHeaderIndexByAliases(headerMap, CSV_NAME_PART_ALIASES.recipientFirst),
    recipientLast: findHeaderIndexByAliases(headerMap, CSV_NAME_PART_ALIASES.recipientLast),
    senderStreet2: findHeaderIndexByAliases(headerMap, CSV_ADDRESS2_ALIASES.senderStreet2),
    recipientStreet2: findHeaderIndexByAliases(headerMap, CSV_ADDRESS2_ALIASES.recipientStreet2),
    dimLength: findHeaderIndexByAliases(headerMap, CSV_DIMENSION_PART_ALIASES.length),
    dimWidth: findHeaderIndexByAliases(headerMap, CSV_DIMENSION_PART_ALIASES.width),
    dimHeight: findHeaderIndexByAliases(headerMap, CSV_DIMENSION_PART_ALIASES.height),
    warehouse: findHeaderIndexByAliases(headerMap, CSV_WAREHOUSE_ALIASES),
  };
}

function analyzeCsvText(text) {
  const delimiter = detectCsvDelimiter(text);
  const matrix = parseCsvMatrix(text, delimiter);
  if (matrix.length < 2) return null;

  const headers = matrix[0].map((header) => String(header || "").trim());
  const rows = matrix
    .slice(1)
    .filter((row) => Array.isArray(row) && row.some((cell) => String(cell || "").trim() !== ""));
  if (!rows.length) return null;

  const headerMap = buildCsvHeaderMap(headers);
  const autoColumnIndices = normalizeAutoCsvColumnIndices(
    detectAutoCsvColumnIndices(headerMap)
  );
  const extraIndices = detectCsvExtraIndices(headerMap);

  return {
    delimiter,
    headers,
    rows,
    headerMap,
    autoColumnIndices,
    extraIndices,
  };
}

function getCsvCompositeFallbackValue(key, row, extraIndices) {
  if (!row || !extraIndices) return "";
  if (key === "recipientName") {
    return joinCsvParts([
      getCsvValueAt(row, extraIndices.recipientFirst),
      getCsvValueAt(row, extraIndices.recipientLast),
    ]);
  }
  if (key === "senderName") {
    return (
      joinCsvParts([
        getCsvValueAt(row, extraIndices.senderFirst),
        getCsvValueAt(row, extraIndices.senderLast),
      ]) || getCsvValueAt(row, extraIndices.warehouse)
    );
  }
  if (key === "packageDims") {
    return composeCsvDimensions(
      "",
      getCsvValueAt(row, extraIndices.dimLength),
      getCsvValueAt(row, extraIndices.dimWidth),
      getCsvValueAt(row, extraIndices.dimHeight)
    );
  }
  return "";
}

function findCsvSampleValue(analysis, key, mappingIndex) {
  if (!analysis || !Array.isArray(analysis.rows)) return "";
  for (const row of analysis.rows) {
    const direct = getCsvValueAt(row, mappingIndex);
    if (direct) return direct;
    if (mappingIndex < 0) {
      const composite = getCsvCompositeFallbackValue(key, row, analysis.extraIndices);
      if (composite) return composite;
    }
  }
  return "";
}

function renderCsvMappingTable() {
  if (!csvMappingBody || !csvMappingDraft) return;
  const { analysis, mapping, fileName } = csvMappingDraft;
  csvMappingBody.innerHTML = "";
  if (csvMapMeta) {
    csvMapMeta.textContent = `${fileName} • ${analysis.rows.length} rows • ${analysis.headers.length} columns`;
  }

  const orderedColumns = getCsvColumnsRequiredFirst();
  orderedColumns.forEach((column) => {
    const tr = document.createElement("tr");

    const fieldCell = document.createElement("td");
    const fieldWrap = document.createElement("div");
    fieldWrap.className = "csv-map-field";
    const fieldLabel = document.createElement("strong");
    fieldLabel.textContent = column.label;
    fieldWrap.appendChild(fieldLabel);
    if (CSV_REQUIRED_FIELDS.has(column.key)) {
      const badge = document.createElement("span");
      badge.className = "csv-map-required";
      badge.textContent = "Required";
      fieldWrap.appendChild(badge);
    }
    fieldCell.appendChild(fieldWrap);

    const selectCell = document.createElement("td");
    const select = document.createElement("select");
    select.className = "csv-map-select";
    select.dataset.mapKey = column.key;

    const noneOption = document.createElement("option");
    noneOption.value = "-1";
    noneOption.textContent = "Not mapped";
    select.appendChild(noneOption);

    analysis.headers.forEach((header, index) => {
      const option = document.createElement("option");
      option.value = String(index);
      option.textContent = header || `Column ${index + 1}`;
      select.appendChild(option);
    });

    const mappedIndex = Number(mapping[column.key]);
    select.value = Number.isInteger(mappedIndex) ? String(mappedIndex) : "-1";
    select.addEventListener("change", () => {
      if (!csvMappingDraft) return;
      csvMappingDraft.mapping[column.key] = Number(select.value);
      updateCsvMappingSample(column.key);
      setCsvMapError("");
    });
    selectCell.appendChild(select);

    const sampleCell = document.createElement("td");
    const sample = document.createElement("span");
    sample.className = "csv-map-sample";
    sample.dataset.sampleKey = column.key;
    sampleCell.appendChild(sample);

    tr.appendChild(fieldCell);
    tr.appendChild(selectCell);
    tr.appendChild(sampleCell);
    csvMappingBody.appendChild(tr);
  });

  orderedColumns.forEach((column) => updateCsvMappingSample(column.key));
}

function updateCsvMappingSample(key) {
  if (!csvMappingDraft || !csvMappingBody) return;
  const sampleEl = csvMappingBody.querySelector(`[data-sample-key="${key}"]`);
  if (!sampleEl) return;

  const mappedIndex = Number(csvMappingDraft.mapping[key]);
  const sampleValue = findCsvSampleValue(csvMappingDraft.analysis, key, mappedIndex);
  if (!sampleValue) {
    sampleEl.textContent = "No sample";
    sampleEl.classList.add("is-empty");
    return;
  }
  sampleEl.textContent = sampleValue;
  sampleEl.classList.remove("is-empty");
}

function getMissingRequiredCsvMappings(mapping, analysis) {
  const missing = [];
  CSV_REQUIRED_FIELDS.forEach((key) => {
    const mappedIndex = Number(mapping[key]);
    if (mappedIndex >= 0) return;

    if (key === "recipientName") {
      const hasNameParts =
        analysis.extraIndices.recipientFirst >= 0 && analysis.extraIndices.recipientLast >= 0;
      if (hasNameParts) return;
    }

    if (key === "packageDims") {
      const hasDimsParts =
        analysis.extraIndices.dimLength >= 0 &&
        analysis.extraIndices.dimWidth >= 0 &&
        analysis.extraIndices.dimHeight >= 0;
      if (hasDimsParts) return;
    }

    missing.push(getCsvColumnLabel(key));
  });
  return missing;
}

function getRequiredCsvMappingConflicts(mapping, analysis) {
  const indexToKey = new Map();
  const conflicts = [];

  CSV_REQUIRED_FIELDS.forEach((key) => {
    const mappedIndex = Number(mapping[key]);
    if (mappedIndex < 0) {
      if (
        key === "recipientName" &&
        analysis.extraIndices.recipientFirst >= 0 &&
        analysis.extraIndices.recipientLast >= 0
      ) {
        return;
      }
      if (
        key === "packageDims" &&
        analysis.extraIndices.dimLength >= 0 &&
        analysis.extraIndices.dimWidth >= 0 &&
        analysis.extraIndices.dimHeight >= 0
      ) {
        return;
      }
      return;
    }

    if (!indexToKey.has(mappedIndex)) {
      indexToKey.set(mappedIndex, key);
      return;
    }

    const previousKey = indexToKey.get(mappedIndex);
    conflicts.push([previousKey, key]);
  });

  return conflicts;
}

function applyCsvMapping() {
  if (!csvMappingDraft) return;
  const missingRequired = getMissingRequiredCsvMappings(
    csvMappingDraft.mapping,
    csvMappingDraft.analysis
  );
  if (missingRequired.length) {
    setCsvMapError(`Map required fields: ${missingRequired.join(", ")}.`);
    return;
  }

  const requiredConflicts = getRequiredCsvMappingConflicts(
    csvMappingDraft.mapping,
    csvMappingDraft.analysis
  );
  if (requiredConflicts.length) {
    const [firstConflict] = requiredConflicts;
    const conflictLabels = firstConflict.map((key) => getCsvColumnLabel(key));
    setCsvMapError(
      `Required fields must map to separate CSV columns (${conflictLabels.join(" + ")}).`
    );
    return;
  }

  const rows = buildCsvRowsFromAnalysis(csvMappingDraft.analysis, csvMappingDraft.mapping);
  if (!rows.length) {
    setCsvMapError("This mapping produced no usable rows. Adjust mapping and try again.");
    return;
  }

  clearBatchState();
  state.csvRows = rows;
  setCsvMode(true);
  setCsvEditMode(false);
  renderCsvTable();
  if (labelError) labelError.classList.remove("is-visible");
  updatePreview();
  updateSummary();
  updatePayment();
  closeCsvModal();
}

function buildCsvRowsFromAnalysis(analysis, selectedColumnIndices = {}) {
  if (!analysis || !Array.isArray(analysis.rows)) return [];
  const senderDefaults = getCsvSenderDefaults();
  const parcelDefaults = {
    packageWeight: String(inputMap.packageWeight?.value || state.info.packageWeight || "").trim(),
    packageDims: String(inputMap.packageDims?.value || state.info.packageDims || "").trim(),
  };

  const columnIndices = {};
  csvColumns.forEach((column) => {
    const mapped = Number(selectedColumnIndices[column.key]);
    const autoMapped = Number(analysis.autoColumnIndices?.[column.key]);
    columnIndices[column.key] = Number.isInteger(mapped)
      ? mapped
      : Number.isInteger(autoMapped)
        ? autoMapped
        : -1;
  });

  const parsedRows = [];
  analysis.rows.forEach((row) => {
    if (!Array.isArray(row) || row.every((cell) => String(cell || "").trim() === "")) {
      return;
    }

    const senderNameDirect = getCsvValueAt(row, columnIndices.senderName);
    const senderFirst = getCsvValueAt(row, analysis.extraIndices.senderFirst);
    const senderLast = getCsvValueAt(row, analysis.extraIndices.senderLast);
    const warehouseName = getCsvValueAt(row, analysis.extraIndices.warehouse);
    const senderName = senderNameDirect || joinCsvParts([senderFirst, senderLast]) || warehouseName;

    const recipientNameDirect = getCsvValueAt(row, columnIndices.recipientName);
    const recipientFirst = getCsvValueAt(row, analysis.extraIndices.recipientFirst);
    const recipientLast = getCsvValueAt(row, analysis.extraIndices.recipientLast);
    const recipientName = recipientNameDirect || joinCsvParts([recipientFirst, recipientLast]);

    const senderStreet = joinCsvAddress(
      getCsvValueAt(row, columnIndices.senderStreet),
      getCsvValueAt(row, analysis.extraIndices.senderStreet2)
    );
    const recipientStreet = joinCsvAddress(
      getCsvValueAt(row, columnIndices.recipientStreet),
      getCsvValueAt(row, analysis.extraIndices.recipientStreet2)
    );

    const existingDims = getCsvValueAt(row, columnIndices.packageDims);
    const packageDims = composeCsvDimensions(
      existingDims,
      getCsvValueAt(row, analysis.extraIndices.dimLength),
      getCsvValueAt(row, analysis.extraIndices.dimWidth),
      getCsvValueAt(row, analysis.extraIndices.dimHeight)
    );

    const recipientCountry = getCsvValueAt(row, columnIndices.recipientCountry) || "France";
    const packageWeightHeaderToken =
      columnIndices.packageWeight >= 0
        ? normalizeCsvHeaderToken(analysis.headers[columnIndices.packageWeight])
        : "";
    const rawWeight = getCsvValueAt(row, columnIndices.packageWeight);
    const packageWeight =
      normalizeCsvWeight(rawWeight, packageWeightHeaderToken) || parcelDefaults.packageWeight;

    const parsed = {
      senderName: senderName || senderDefaults.senderName,
      senderStreet: senderStreet || senderDefaults.senderStreet,
      senderCity: getCsvValueAt(row, columnIndices.senderCity) || senderDefaults.senderCity,
      senderState: getCsvValueAt(row, columnIndices.senderState) || senderDefaults.senderState,
      senderZip: getCsvValueAt(row, columnIndices.senderZip) || senderDefaults.senderZip,
      recipientName,
      recipientStreet,
      recipientCity: getCsvValueAt(row, columnIndices.recipientCity),
      recipientState: getCsvValueAt(row, columnIndices.recipientState),
      recipientZip: getCsvValueAt(row, columnIndices.recipientZip),
      recipientCountry,
      packageWeight,
      packageDims: packageDims || parcelDefaults.packageDims,
    };

    const hasUsefulData = csvColumns.some((column) => {
      return String(parsed[column.key] || "").trim() !== "";
    });
    if (!hasUsefulData) return;
    parsedRows.push(parsed);
  });

  return parsedRows;
}

function parseCsvText(text) {
  const analysis = analyzeCsvText(text);
  if (!analysis) return [];
  return buildCsvRowsFromAnalysis(analysis, analysis.autoColumnIndices);
}

ensureIdentifiers();
state.quantity = Number(quantityInput.value) || 1;
updateCountryFlag();
updateSummary();
updatePayment();
updatePreview();
const initialRoute = parseRouteFromLocation();
const initialStep =
  initialRoute.view === "builder" ? clampStep(initialRoute.step) : clampStep(state.step);

goToStep(initialStep, { push: false, regenerate: false });
updateRoute(
  initialRoute.view === "builder" ? { view: "builder", step: initialStep } : initialRoute,
  { replace: true }
);

window.addEventListener("popstate", (event) => {
  const route =
    event.state && typeof event.state.view === "string"
      ? event.state
      : parseRouteFromLocation();

  if (!currentUser) {
    setMainView("builder", { push: false });
    updateRoute({ view: "login" }, { replace: true });
    return;
  }

  if (route.view === "account") {
    setMainView("account", { push: false });
    loadGenerationHistory({ preferLatest: true });
    return;
  }

  if (route.view === "reports") {
    setMainView("reports", { push: false });
    loadGenerationHistory({ preferLatest: true });
    return;
  }

  if (route.view === "builder") {
    setMainView("builder", { push: false });
    const targetStep = clampStep(route.step);
    goToStep(targetStep, { push: false, regenerate: false });
    if (route.customs && targetStep === 1) {
      setCustomsGhostVisible(true, { push: false });
    }
    return;
  }

  setMainView("builder", { push: false });
  updateRoute({ view: "builder", step: state.step }, { replace: true });
});
if (batchPreview) {
  batchPreview.classList.add("is-single");
}

resetWarehouseState();
initializeAuthLogo();
initializeAppLogo();
initializeAuthParticles();
initializeAuth();
