const DATA_FIELDS = [
  {
    key: "shipment_date",
    label: "Shipment date",
    aliases: ["date", "shipment_date", "created_at", "order_date", "shipping_date", "label_date"],
  },
  {
    key: "origin_postcode",
    label: "Origin postcode",
    aliases: ["origin_postcode", "origin_postal_code", "from_postcode", "from_zip", "sender_zip", "warehouse_postcode"],
  },
  {
    key: "origin_country",
    label: "Origin country",
    aliases: ["origin_country", "from_country", "sender_country", "warehouse_country"],
  },
  {
    key: "destination_country",
    label: "Destination country",
    aliases: ["destination_country", "dest_country", "to_country", "shipping_country", "recipient_country", "country"],
  },
  {
    key: "destination_postcode",
    label: "Destination postcode",
    aliases: ["destination_postcode", "destination_postal_code", "dest_zip", "to_zip", "shipping_zip", "recipient_postcode", "postcode", "postal_code"],
  },
  {
    key: "weight",
    label: "Weight",
    aliases: ["weight", "package_weight", "parcel_weight", "shipment_weight", "weight_kg", "weight_g", "poids"],
  },
  {
    key: "weight_unit",
    label: "Weight unit",
    aliases: ["weight_unit", "unit_weight", "weight_units"],
  },
  {
    key: "length",
    label: "Length",
    aliases: ["length", "package_length", "parcel_length", "dim_length"],
  },
  {
    key: "width",
    label: "Width",
    aliases: ["width", "package_width", "parcel_width", "dim_width"],
  },
  {
    key: "height",
    label: "Height",
    aliases: ["height", "package_height", "parcel_height", "dim_height"],
  },
  {
    key: "dimensions",
    label: "Dimensions",
    aliases: ["dimensions", "dims", "size", "package_dims", "parcel_dimensions"],
  },
  {
    key: "dimension_unit",
    label: "Dimension unit",
    aliases: ["dimension_unit", "dims_unit", "length_unit", "unit_dimension"],
  },
  {
    key: "service_type",
    label: "Service type",
    aliases: ["service", "service_type", "shipping_method", "carrier_service", "delivery_method", "transporteur"],
  },
  {
    key: "quantity",
    label: "Quantity",
    aliases: ["quantity", "qty", "labels", "label_count", "number_of_labels"],
  },
];

const EXTRACTED_FIELDS = [
  {
    key: "origin_postcode",
    label: "Origin postcode",
    sourceType: "origin_address",
  },
  {
    key: "origin_country",
    label: "Origin country",
    sourceType: "origin_address",
  },
  {
    key: "destination_postcode",
    label: "Destination postcode",
    sourceType: "destination_address",
  },
  {
    key: "destination_country",
    label: "Destination country",
    sourceType: "destination_address",
  },
];

const REMOVE_PATTERNS = [
  /(^|_)(name|first_name|last_name|full_name|customer|buyer|consignee)(_|$)/,
  /(^|_)(email|e_mail|mail)(_|$)/,
  /(^|_)(phone|tel|mobile|gsm)(_|$)/,
  /(^|_)(address|street|address1|address_1|line_1|line1|address2|line_2|line2)(_|$)/,
  /(^|_)(city|town|commune)(_|$)/,
  /(^|_)(company|organization|organisation)(_|$)/,
  /(^|_)(order|order_id|order_number|tracking|tracking_number|reference|note|notes|comment)(_|$)/,
];

const state = {
  fileName: "",
  headers: [],
  rows: [],
  mapping: {},
  step: "upload",
  transitionToken: 0,
};

const STEP_OUT_MS = 220;
const STEP_RESIZE_MS = 520;
const STEP_IN_MS = 260;

const els = {
  panels: {
    upload: document.getElementById("uploadPanel"),
    mapping: document.getElementById("mappingPanel"),
    review: document.getElementById("reviewPanel"),
    thanks: document.getElementById("thanksPanel"),
  },
  progress: Array.from(document.querySelectorAll("[data-progress-step]")),
  stepLabel: document.getElementById("cleanerStepLabel"),
  dropzone: document.getElementById("cleanerDropzone"),
  fileInput: document.getElementById("cleanerFileInput"),
  mappingBody: document.getElementById("cleanerMappingBody"),
  title: document.getElementById("cleanerTitle"),
  previewHead: document.getElementById("cleanerPreviewHead"),
  previewBody: document.getElementById("cleanerPreviewBody"),
  card: document.querySelector(".cleaner-card"),
  toastStack: document.getElementById("toastStack"),
  confirmModal: document.getElementById("cleanerConfirmModal"),
  confirmCancel: document.getElementById("cleanerConfirmCancel"),
  confirmApprove: document.getElementById("cleanerConfirmApprove"),
  truthConfirm: document.getElementById("cleanerTruthConfirm"),
  backToUpload: document.getElementById("cleanerBackToUpload"),
  backToMapping: document.getElementById("cleanerBackToMapping"),
  continueBtn: document.getElementById("cleanerContinue"),
  download: document.getElementById("cleanerDownload"),
  downloadFromMap: document.getElementById("cleanerDownloadFromMap"),
  submit: document.getElementById("cleanerSubmit"),
  pixelCanvas: document.getElementById("cleanerPixelCanvas"),
};

function setStatus(message = "", tone = "") {
  const text = String(message || "").trim();
  if (!text) return;
  const normalizedTone = tone === "success" ? "success" : tone === "error" ? "error" : "info";
  showStatusToast(text, { tone: normalizedTone });
}

function getToastIconSvg(tone) {
  if (tone === "success") {
    return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M8.5 12.4l2.3 2.3 4.7-5.1"/></svg>';
  }
  if (tone === "error") {
    return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3.7 21 19.3H3L12 3.7Z"/><path d="M12 9v4.6"/><circle cx="12" cy="16.9" r="0.8" fill="currentColor" stroke="none"/></svg>';
  }
  return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 10.2v5"/><circle cx="12" cy="7.3" r="0.8" fill="currentColor" stroke="none"/></svg>';
}

function updateToastStackLayout() {
  if (!els.toastStack) return;
  const toasts = Array.from(els.toastStack.querySelectorAll(".toast"));
  toasts.forEach((toast, index) => {
    const depth = Math.min(index, 3);
    const offset = depth * 14;
    const scale = 1 - depth * 0.035;
    const opacity = 1 - depth * 0.14;
    toast.style.setProperty("--stack-offset", `${offset}px`);
    toast.style.setProperty("--stack-scale", `${scale}`);
    toast.style.setProperty("--stack-opacity", `${Math.max(0.45, opacity)}`);
    toast.style.setProperty("--stack-z", `${100 - index}`);
  });
}

let lastStatusToast = {
  message: "",
  tone: "",
  title: "",
  at: 0,
};

function showStatusToast(message, options = {}) {
  const text = String(message || "").trim();
  if (!text || !els.toastStack) return;
  const title = String(options?.title || "").trim();
  const rawTone = String(options?.tone || "info").trim().toLowerCase();
  const tone = rawTone === "success" ? "success" : rawTone === "error" || rawTone === "warning" ? "error" : "info";
  const now = Date.now();
  if (
    lastStatusToast.message === text &&
    lastStatusToast.tone === tone &&
    lastStatusToast.title === title &&
    now - lastStatusToast.at < 900
  ) {
    return;
  }
  lastStatusToast = { message: text, tone, title, at: now };
  showToast(text, { ...options, title, tone });
}

function showToast(message, options = {}) {
  if (!els.toastStack) return;
  const text = String(message || "").trim();
  if (!text) return;
  const rawTone = String(options?.tone || "info").trim().toLowerCase();
  const tone = rawTone === "success" ? "success" : rawTone === "error" || rawTone === "warning" ? "error" : "info";
  const { title = "", duration = 3400 } = options;
  const toast = document.createElement("div");
  toast.className = `toast is-${tone}`;

  const main = document.createElement("div");
  main.className = "toast-main";

  const icon = document.createElement("span");
  icon.className = `toast-icon is-${tone}`;
  icon.setAttribute("aria-hidden", "true");
  icon.innerHTML = getToastIconSvg(tone);

  const body = document.createElement("div");
  body.className = "toast-body";

  if (title) {
    const titleEl = document.createElement("div");
    titleEl.className = "toast-title";
    titleEl.textContent = String(title).trim();
    body.appendChild(titleEl);
  }

  const messageEl = document.createElement("div");
  messageEl.className = "toast-message";
  messageEl.textContent = text;
  body.appendChild(messageEl);

  main.appendChild(icon);
  main.appendChild(body);
  toast.appendChild(main);
  els.toastStack.prepend(toast);
  updateToastStackLayout();

  window.requestAnimationFrame(() => {
    window.requestAnimationFrame(() => {
      toast.classList.add("is-visible");
    });
  });

  const dismiss = () => {
    toast.classList.remove("is-visible");
    toast.classList.add("is-exit");
    window.setTimeout(() => {
      toast.remove();
      updateToastStackLayout();
    }, 420);
  };

  window.setTimeout(dismiss, Math.max(1800, Number(duration) || 3400));
}

function getStepPanel(step) {
  return els.panels[step] || els.panels.upload;
}

function getStepWidth(step) {
  const shellWidth = els.card?.parentElement?.getBoundingClientRect().width || window.innerWidth;
  if (step === "upload") return Math.min(640, shellWidth);
  if (step === "thanks") return Math.min(760, shellWidth);
  return Math.min(1040, shellWidth);
}

function measurePanelHeight(panel) {
  if (!panel || !els.card) return 0;
  const wasActive = panel.classList.contains("is-active");
  const previousStyles = {
    position: panel.style.position,
    inset: panel.style.inset,
    width: panel.style.width,
    visibility: panel.style.visibility,
    pointerEvents: panel.style.pointerEvents,
    zIndex: panel.style.zIndex,
  };

  panel.classList.add("is-active");
  panel.classList.remove(
    "is-step-transition-exit",
    "is-step-transition-hidden",
    "is-step-transition-enter-start"
  );
  panel.style.position = "absolute";
  panel.style.inset = "0 auto auto 0";
  panel.style.width = `${Math.floor(els.card.getBoundingClientRect().width)}px`;
  panel.style.visibility = "hidden";
  panel.style.pointerEvents = "none";
  panel.style.zIndex = "-1";

  const height = panel.getBoundingClientRect().height;

  panel.style.position = previousStyles.position;
  panel.style.inset = previousStyles.inset;
  panel.style.width = previousStyles.width;
  panel.style.visibility = previousStyles.visibility;
  panel.style.pointerEvents = previousStyles.pointerEvents;
  panel.style.zIndex = previousStyles.zIndex;

  if (!wasActive) {
    panel.classList.remove("is-active");
  }

  return height;
}

function applyStepView(step) {
  els.card?.classList.remove("is-step-upload", "is-step-mapping", "is-step-review", "is-step-thanks");
  els.card?.classList.add(`is-step-${step}`);
  Object.entries(els.panels).forEach(([key, panel]) => {
    panel.classList.toggle("is-active", key === step);
    panel.classList.remove(
      "is-step-transition-exit",
      "is-step-transition-hidden",
      "is-step-transition-enter-start"
    );
  });
  const stepIndex = step === "upload" ? 1 : step === "mapping" ? 2 : 3;
  const titles = {
    upload: "Clean your shipping data before sending it to Shipide.",
    mapping: "Review detected columns",
    review: "Cleaned file preview",
    thanks: "Submission complete",
  };
  if (els.title) {
    els.title.textContent = titles[step] || titles.upload;
  }
  els.progress.forEach((item) => {
    item.classList.toggle("is-active", step === "thanks" || Number(item.dataset.progressStep) <= stepIndex);
  });
  els.stepLabel.textContent = step === "thanks" ? "Complete" : `Step ${stepIndex} of 3`;
}

function resetStepTransitionState() {
  state.transitionToken += 1;
  els.card?.classList.remove("is-step-transitioning");
  if (els.card) {
    els.card.style.width = "";
    els.card.style.height = "";
  }
  Object.values(els.panels).forEach((panel) => {
    panel.classList.remove(
      "is-step-transition-exit",
      "is-step-transition-hidden",
      "is-step-transition-enter-start"
    );
  });
}

function setStep(step, options = {}) {
  const nextStep = ["upload", "mapping", "review", "thanks"].includes(step) ? step : "upload";
  const previousStep = state.step || "upload";
  const shouldAnimate =
    options.animate !== false &&
    previousStep !== nextStep &&
    els.card &&
    !window.matchMedia("(prefers-reduced-motion: reduce)").matches;

  if (!shouldAnimate) {
    resetStepTransitionState();
    state.step = nextStep;
    applyStepView(nextStep);
    return;
  }

  const transitionToken = ++state.transitionToken;
  const currentPanel = getStepPanel(previousStep);
  const incomingPanel = getStepPanel(nextStep);
  const cardRect = els.card.getBoundingClientRect();
  const currentPanelHeight = Math.max(1, currentPanel.getBoundingClientRect().height);
  const incomingPanelHeight = Math.max(1, measurePanelHeight(incomingPanel));
  const chromeHeight = Math.max(0, cardRect.height - currentPanelHeight);
  const targetHeight = Math.max(1, chromeHeight + incomingPanelHeight);
  const targetWidth = getStepWidth(nextStep);

  els.card.style.width = `${Math.round(cardRect.width)}px`;
  els.card.style.height = `${Math.round(cardRect.height)}px`;
  els.card.classList.add("is-step-transitioning");
  currentPanel.classList.add("is-step-transition-exit");

  window.setTimeout(() => {
    if (transitionToken !== state.transitionToken) return;
    currentPanel.classList.remove("is-step-transition-exit", "is-active");
    currentPanel.classList.add("is-step-transition-hidden");
    state.step = nextStep;
    applyStepView(nextStep);
    incomingPanel.classList.add("is-step-transition-hidden");
    els.card.style.width = `${Math.round(targetWidth)}px`;
    els.card.style.height = `${Math.round(targetHeight)}px`;
  }, STEP_OUT_MS);

  window.setTimeout(() => {
    if (transitionToken !== state.transitionToken) return;
    incomingPanel.classList.remove("is-step-transition-hidden");
    incomingPanel.classList.add("is-step-transition-enter-start");
    window.requestAnimationFrame(() => {
      if (transitionToken !== state.transitionToken) return;
      incomingPanel.classList.remove("is-step-transition-enter-start");
    });
  }, STEP_OUT_MS + STEP_RESIZE_MS);

  window.setTimeout(() => {
    if (transitionToken !== state.transitionToken) return;
    resetStepTransitionState();
    state.step = nextStep;
    applyStepView(nextStep);
  }, STEP_OUT_MS + STEP_RESIZE_MS + STEP_IN_MS);
}

function normalizeHeader(value) {
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
      if (inQuotes && line[i + 1] === '"') i += 1;
      else inQuotes = !inQuotes;
      continue;
    }
    if (!inQuotes && char === delimiter) count += 1;
  }
  return count;
}

function detectDelimiter(text) {
  const lines = String(text || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .slice(0, 8);
  const candidates = [",", ";", "\t", "|"];
  let best = ",";
  let score = -1;
  candidates.forEach((delimiter) => {
    const nextScore = lines.reduce((sum, line) => sum + countDelimiterOutsideQuotes(line, delimiter), 0);
    if (nextScore > score) {
      best = delimiter;
      score = nextScore;
    }
  });
  return score <= 0 ? "," : best;
}

function parseDelimitedText(text, delimiter) {
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
      if (char === "\r" && text[i + 1] === "\n") i += 1;
      row.push(value);
      if (row.some((cell) => String(cell || "").trim())) rows.push(row);
      row = [];
      value = "";
      continue;
    }
    value += char;
  }
  row.push(value);
  if (row.some((cell) => String(cell || "").trim())) rows.push(row);
  return rows;
}

function parseHtmlTable(text) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(text, "text/html");
  const table = doc.querySelector("table");
  if (!table) return null;
  return Array.from(table.querySelectorAll("tr"))
    .map((tr) => Array.from(tr.children).map((cell) => cell.textContent.trim()))
    .filter((row) => row.some(Boolean));
}

function detectField(header) {
  const normalized = normalizeHeader(header);
  if (isOriginAddressHeader(normalized)) {
    return { action: "remove", reason: "Likely personal or identifying data." };
  }
  if (isDestinationAddressHeader(normalized)) {
    return {
      action: "remove",
      reason: "Likely personal or identifying data.",
    };
  }
  const direct = DATA_FIELDS.find((field) => {
    const aliases = [field.key, field.label, ...field.aliases].map(normalizeHeader);
    return aliases.includes(normalized);
  });
  if (direct) {
    return { action: direct.key, reason: "Matched useful shipment field." };
  }
  const fuzzy = DATA_FIELDS.find((field) => {
    const aliases = [field.key, ...field.aliases].map(normalizeHeader);
    return aliases.some((alias) => alias && (normalized.includes(alias) || alias.includes(normalized)));
  });
  if (fuzzy) {
    return { action: fuzzy.key, reason: "Looks like shipment-analysis data." };
  }
  if (REMOVE_PATTERNS.some((pattern) => pattern.test(normalized))) {
    return { action: "remove", reason: "Likely personal or identifying data." };
  }
  return { action: "remove", reason: "Looks like shipment-analysis data." };
}

function isOriginAddressHeader(normalized) {
  return (
    normalized.includes("shipped_from_address") ||
    normalized.includes("ship_from_address") ||
    normalized.includes("from_address") ||
    normalized.includes("sender_address") ||
    normalized.includes("origin_address")
  );
}

function isDestinationAddressHeader(normalized) {
  return (
    normalized.includes("shipped_to_address") ||
    normalized.includes("ship_to_address") ||
    normalized.includes("to_address") ||
    normalized.includes("recipient_address") ||
    normalized.includes("destination_address") ||
    normalized.includes("delivery_address") ||
    normalized.includes("shipping_address")
  );
}

function getAddressSourceType(header) {
  const normalized = normalizeHeader(header);
  if (isOriginAddressHeader(normalized)) return "origin_address";
  if (isDestinationAddressHeader(normalized)) return "destination_address";
  return "";
}

function parseCombinedAddress(value) {
  const parts = String(value || "")
    .split(",")
    .map((part) => part.trim())
    .filter(Boolean);
  const country = parts.length ? parts[parts.length - 1] : "";
  const cityPostcodePart = parts.length >= 2 ? parts[parts.length - 2] : String(value || "");
  const postcodeMatch =
    cityPostcodePart.match(/\b([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})\b/i) ||
    cityPostcodePart.match(/\b(\d{4,6}(?:-\d{3,4})?)\b/);
  return {
    postcode: postcodeMatch ? postcodeMatch[1].trim() : "",
    country,
  };
}

function getExtractedValue(row, fieldKey) {
  for (let index = 0; index < state.headers.length; index += 1) {
    const sourceType = getAddressSourceType(state.headers[index]);
    if (!sourceType) continue;
    if ((fieldKey === "origin_postcode" || fieldKey === "origin_country") && sourceType !== "origin_address") {
      continue;
    }
    if (
      (fieldKey === "destination_postcode" || fieldKey === "destination_country") &&
      sourceType !== "destination_address"
    ) {
      continue;
    }
    const parsed = parseCombinedAddress(row[index]);
    if (fieldKey === "origin_postcode" || fieldKey === "destination_postcode") {
      if (parsed.postcode) return parsed.postcode;
    }
    if ((fieldKey === "origin_country" || fieldKey === "destination_country") && parsed.country) return parsed.country;
  }
  return "";
}

function getAvailableExtractedFields() {
  const hasOriginAddress = state.headers.some((header) => getAddressSourceType(header) === "origin_address");
  const hasDestinationAddress = state.headers.some((header) => getAddressSourceType(header) === "destination_address");
  const directlyMapped = new Set(
    Object.values(state.mapping)
      .map((item) => item.action)
      .filter((action) => action && action !== "remove")
  );
  return EXTRACTED_FIELDS.filter((field) => {
    if (directlyMapped.has(field.key)) return false;
    if (field.sourceType === "origin_address") return hasOriginAddress;
    if (field.sourceType === "destination_address") return hasDestinationAddress;
    return false;
  });
}

function getKeptFieldKeys() {
  const direct = Object.values(state.mapping)
    .map((item) => item.action)
    .filter((action) => action && action !== "remove");
  return Array.from(new Set([...direct, ...getAvailableExtractedFields().map((field) => field.key)]));
}

function getCleanedDataReadiness() {
  const cleaned = buildCleanRows();
  return {
    cleaned,
    isReady: cleaned.headers.length > 0 && cleaned.rows.length > 0,
  };
}

function getSampleForColumn(index) {
  for (const row of state.rows.slice(0, 20)) {
    const value = String(row[index] || "").trim();
    if (value) return value;
  }
  return "No sample";
}

function createMappingRow(header, index) {
  const row = document.createElement("tr");
  const nameCell = document.createElement("td");
  const name = document.createElement("div");
  name.className = "cleaner-column-name";
  name.innerHTML = `<strong></strong><span class="mono"></span>`;
  name.querySelector("strong").textContent = header || `Column ${index + 1}`;
  name.querySelector("span").textContent = normalizeHeader(header) || "untitled";
  nameCell.appendChild(name);

  const actionCell = document.createElement("td");
  const select = document.createElement("select");
  select.className = `cleaner-map-select ${state.mapping[index]?.action === "remove" ? "is-remove" : "is-keep"}`;
  select.dataset.columnIndex = String(index);
  const removeOption = new Option("Remove from submission", "remove");
  select.appendChild(removeOption);
  DATA_FIELDS.forEach((field) => select.appendChild(new Option(`Keep as ${field.label}`, field.key)));
  select.value = state.mapping[index]?.action || "remove";
  select.addEventListener("change", () => {
    const nextField = DATA_FIELDS.find((field) => field.key === select.value);
    state.mapping[index] = {
      action: select.value,
      reason:
        select.value === "remove"
          ? "Likely personal or identifying data."
          : nextField
            ? "Matched useful shipment field."
            : "Looks like shipment-analysis data.",
    };
    renderMapping();
  });
  actionCell.appendChild(select);

  const sampleCell = document.createElement("td");
  const sample = document.createElement("span");
  sample.className = "cleaner-sample";
  sample.textContent = getSampleForColumn(index);
  sampleCell.appendChild(sample);

  const reasonCell = document.createElement("td");
  reasonCell.className = "cleaner-reason";
  reasonCell.textContent = state.mapping[index]?.reason || "";

  row.append(nameCell, actionCell, sampleCell, reasonCell);
  return row;
}

function createExtractedMappingRow(field) {
  const row = document.createElement("tr");
  const nameCell = document.createElement("td");
  const name = document.createElement("div");
  name.className = "cleaner-column-name";
  name.innerHTML = `<strong></strong><span class="mono"></span>`;
  name.querySelector("strong").textContent = field.label;
  name.querySelector("span").textContent = "extracted, not raw address";
  nameCell.appendChild(name);

  const actionCell = document.createElement("td");
  const select = document.createElement("select");
  select.className = "cleaner-map-select is-keep";
  select.disabled = true;
  select.appendChild(new Option(`Keep as ${field.label}`, field.key));
  select.value = field.key;
  actionCell.appendChild(select);

  const sampleCell = document.createElement("td");
  const sample = document.createElement("span");
  sample.className = "cleaner-sample";
  sample.textContent = getExtractedValue(state.rows[0] || [], field.key) || "No sample";
  sampleCell.appendChild(sample);

  const reasonCell = document.createElement("td");
  reasonCell.className = "cleaner-reason";
  reasonCell.textContent = "Looks like shipment-analysis data.";

  row.append(nameCell, actionCell, sampleCell, reasonCell);
  return row;
}

function renderMapping() {
  els.mappingBody.innerHTML = "";
  const uploadedColumns = state.headers.map((header, index) => ({ header, index }));
  uploadedColumns
    .filter((item) => state.mapping[item.index]?.action !== "remove")
    .forEach((item) => els.mappingBody.appendChild(createMappingRow(item.header, item.index)));
  getAvailableExtractedFields().forEach((field) => {
    els.mappingBody.appendChild(createExtractedMappingRow(field));
  });
  uploadedColumns
    .filter((item) => state.mapping[item.index]?.action === "remove")
    .forEach((item) => els.mappingBody.appendChild(createMappingRow(item.header, item.index)));
}

function buildCleanRows() {
  const kept = Object.entries(state.mapping)
    .map(([index, config]) => ({ index: Number(index), key: config.action }))
    .filter((item) => item.key && item.key !== "remove");
  const extracted = getAvailableExtractedFields();
  const outputHeaders = [];
  const seen = new Map();
  [...kept, ...extracted.map((field) => ({ key: field.key, extracted: true }))].forEach((item) => {
    const count = seen.get(item.key) || 0;
    seen.set(item.key, count + 1);
    outputHeaders.push(count ? `${item.key}_${count + 1}` : item.key);
  });
  const rows = state.rows.map((row) => [
    ...kept.map((item) => String(row[item.index] || "").trim()),
    ...extracted.map((field) => getExtractedValue(row, field.key)),
  ]);
  return { headers: outputHeaders, rows };
}

function escapeCsv(value) {
  const text = String(value ?? "");
  if (/[",\n\r;]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
  return text;
}

function buildCleanCsv() {
  const cleaned = buildCleanRows();
  return [cleaned.headers, ...cleaned.rows]
    .map((row) => row.map(escapeCsv).join(","))
    .join("\n");
}

function renderPreview() {
  const { cleaned } = getCleanedDataReadiness();
  els.previewHead.innerHTML = "";
  els.previewBody.innerHTML = "";
  const headRow = document.createElement("tr");
  cleaned.headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headRow.appendChild(th);
  });
  els.previewHead.appendChild(headRow);
  cleaned.rows.slice(0, 8).forEach((row) => {
    const tr = document.createElement("tr");
    row.forEach((value) => {
      const td = document.createElement("td");
      td.textContent = value;
      tr.appendChild(td);
    });
    els.previewBody.appendChild(tr);
  });
}

function downloadCleanCsv() {
  const { isReady } = getCleanedDataReadiness();
  if (!isReady) {
    setStatus("Select at least one shipment-analysis field. Dimensions, service type, date, and quantity are optional.", "error");
    return;
  }
  const csv = buildCleanCsv();
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const base = state.fileName.replace(/\.[^.]+$/, "") || "shipping-data";
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `${base}-shipide-cleaned.csv`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function setConfirmModalOpen(open) {
  if (!els.confirmModal) return;
  els.confirmModal.classList.toggle("is-closed", !open);
  if (open) {
    if (els.truthConfirm) els.truthConfirm.checked = false;
    if (els.confirmApprove) els.confirmApprove.disabled = true;
    window.setTimeout(() => {
      els.truthConfirm?.focus();
    }, 50);
  }
}

function requestSubmitConfirmation() {
  const { isReady } = getCleanedDataReadiness();
  if (!isReady) {
    setStatus("Select at least one shipment-analysis field before submitting. Missing dimensions or service type will not block submission.", "error");
    return;
  }
  setConfirmModalOpen(true);
}

async function submitCleanData() {
  const { cleaned, isReady } = getCleanedDataReadiness();
  if (!isReady) {
    setStatus("Select at least one shipment-analysis field before submitting. Missing dimensions or service type will not block submission.", "error");
    return;
  }
  if (!els.truthConfirm?.checked) {
    setStatus("Confirm the information is truthful before submitting.", "error");
    return;
  }
  els.submit.disabled = true;
  if (els.confirmApprove) els.confirmApprove.disabled = true;
  showStatusToast("Submitting cleaned data to Shipide...", { tone: "info" });
  try {
    const response = await fetch("/api/public/shipping-data-cleaner/submit", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        fileName: state.fileName,
        headers: cleaned.headers,
        rows: cleaned.rows,
        removedColumns: state.headers.filter((_, index) => state.mapping[index]?.action === "remove"),
      }),
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(payload?.error || "Could not submit the cleaned data.");
    setConfirmModalOpen(false);
    setStep("thanks");
  } catch (error) {
    setStatus(`${error.message || "Submission failed."} Download the cleaned CSV and send it to Shipide if needed.`, "error");
  } finally {
    els.submit.disabled = false;
    if (els.confirmApprove) {
      els.confirmApprove.disabled = !els.truthConfirm?.checked;
    }
  }
}

function handleMatrix(matrix, fileName) {
  if (!matrix || matrix.length < 2) {
    throw new Error("No usable table was detected. Export the file as CSV or TXT and try again.");
  }
  state.fileName = fileName || "shipping-export.csv";
  state.headers = matrix[0].map((header, index) => String(header || `Column ${index + 1}`).trim());
  state.rows = matrix
    .slice(1)
    .filter((row) => row.some((cell) => String(cell || "").trim()));
  state.mapping = {};
  state.headers.forEach((header, index) => {
    state.mapping[index] = detectField(header);
  });
  renderMapping();
  setStep("mapping");
  setStatus("");
}

function handleFile(file) {
  if (!file) return;
  const name = file.name || "shipping-export.csv";
  const extension = name.split(".").pop().toLowerCase();
  const reader = new FileReader();
  reader.onerror = () => setStatus("Could not read the file. Please try again.", "error");
  reader.onload = () => {
    try {
      const text = String(reader.result || "");
      let matrix = null;
      if (extension === "xls" || /^\s*</.test(text)) {
        matrix = parseHtmlTable(text);
      }
      if (!matrix) {
        matrix = parseDelimitedText(text, detectDelimiter(text));
      }
      handleMatrix(matrix, name);
    } catch (error) {
      setStatus(error.message || "Could not parse this file.", "error");
    }
  };
  if (extension === "xlsx") {
    setStatus("XLSX parsing is not enabled in this first version. Export as CSV or TXT and upload that file.", "error");
    return;
  }
  reader.readAsText(file);
}

function drawBackground() {
  const canvas = els.pixelCanvas;
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  const resize = () => {
    const rect = canvas.getBoundingClientRect();
    canvas.width = Math.max(1, Math.floor(rect.width * window.devicePixelRatio));
    canvas.height = Math.max(1, Math.floor(rect.height * window.devicePixelRatio));
  };
  resize();
  window.addEventListener("resize", resize);
  const pixels = Array.from({ length: 110 }, () => ({
    x: Math.random(),
    y: Math.random(),
    size: 3 + Math.random() * 8,
    speed: 0.00008 + Math.random() * 0.00018,
    alpha: 0.22 + Math.random() * 0.58,
  }));
  const render = (time) => {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = "#00060f";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    pixels.forEach((pixel) => {
      const x = ((pixel.x + time * pixel.speed) % 1) * canvas.width;
      const y = pixel.y * canvas.height;
      ctx.fillStyle = `rgba(119, 71, 227, ${pixel.alpha})`;
      ctx.fillRect(x, y, pixel.size * window.devicePixelRatio, pixel.size * window.devicePixelRatio);
    });
    requestAnimationFrame(render);
  };
  requestAnimationFrame(render);
}

els.dropzone.addEventListener("dragover", (event) => {
  event.preventDefault();
  els.dropzone.classList.add("is-dragover");
});
els.dropzone.addEventListener("dragleave", () => els.dropzone.classList.remove("is-dragover"));
els.dropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  els.dropzone.classList.remove("is-dragover");
  handleFile(event.dataTransfer.files[0]);
});
els.fileInput.addEventListener("change", () => {
  handleFile(els.fileInput.files[0]);
  els.fileInput.value = "";
});
els.backToUpload.addEventListener("click", () => setStep("upload"));
els.backToMapping.addEventListener("click", () => setStep("mapping"));
els.continueBtn.addEventListener("click", () => {
  const { isReady } = getCleanedDataReadiness();
  if (!isReady) {
    setStatus("Select at least one shipment-analysis field. Missing optional fields are fine.", "error");
    return;
  }
  setStatus("");
  renderPreview();
  setStep("review");
});
els.download.addEventListener("click", downloadCleanCsv);
els.downloadFromMap.addEventListener("click", downloadCleanCsv);
els.submit.addEventListener("click", requestSubmitConfirmation);
els.confirmCancel?.addEventListener("click", () => setConfirmModalOpen(false));
els.confirmApprove?.addEventListener("click", submitCleanData);
els.truthConfirm?.addEventListener("change", () => {
  if (els.confirmApprove) {
    els.confirmApprove.disabled = !els.truthConfirm.checked;
  }
});
els.confirmModal?.addEventListener("click", (event) => {
  if (event.target === els.confirmModal) {
    setConfirmModalOpen(false);
  }
});

drawBackground();
setStep("upload");
