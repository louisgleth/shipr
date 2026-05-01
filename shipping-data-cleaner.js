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
};

const els = {
  panels: {
    upload: document.getElementById("uploadPanel"),
    mapping: document.getElementById("mappingPanel"),
    review: document.getElementById("reviewPanel"),
  },
  progress: Array.from(document.querySelectorAll("[data-progress-step]")),
  stepLabel: document.getElementById("cleanerStepLabel"),
  dropzone: document.getElementById("cleanerDropzone"),
  fileInput: document.getElementById("cleanerFileInput"),
  fileMeta: document.getElementById("cleanerFileMeta"),
  mappingBody: document.getElementById("cleanerMappingBody"),
  keptCount: document.getElementById("keptCount"),
  removedCount: document.getElementById("removedCount"),
  rowCount: document.getElementById("rowCount"),
  previewHead: document.getElementById("cleanerPreviewHead"),
  previewBody: document.getElementById("cleanerPreviewBody"),
  reviewMeta: document.getElementById("cleanerReviewMeta"),
  contactEmail: document.getElementById("cleanerContactEmail"),
  status: document.getElementById("cleanerStatus"),
  backToUpload: document.getElementById("cleanerBackToUpload"),
  backToMapping: document.getElementById("cleanerBackToMapping"),
  continueBtn: document.getElementById("cleanerContinue"),
  download: document.getElementById("cleanerDownload"),
  downloadFromMap: document.getElementById("cleanerDownloadFromMap"),
  submit: document.getElementById("cleanerSubmit"),
  pixelCanvas: document.getElementById("cleanerPixelCanvas"),
};

function setStatus(message = "", tone = "") {
  els.status.textContent = message;
  els.status.classList.toggle("is-error", tone === "error");
  els.status.classList.toggle("is-success", tone === "success");
}

function setStep(step) {
  Object.entries(els.panels).forEach(([key, panel]) => {
    panel.classList.toggle("is-active", key === step);
  });
  const stepIndex = step === "upload" ? 1 : step === "mapping" ? 2 : 3;
  els.progress.forEach((item) => {
    item.classList.toggle("is-active", Number(item.dataset.progressStep) <= stepIndex);
  });
  els.stepLabel.textContent = `Step ${stepIndex} of 3`;
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
  return { action: "remove", reason: "Not needed for shipping-rate analysis." };
}

function getSampleForColumn(index) {
  for (const row of state.rows.slice(0, 20)) {
    const value = String(row[index] || "").trim();
    if (value) return value;
  }
  return "No sample";
}

function renderMapping() {
  els.mappingBody.innerHTML = "";
  state.headers.forEach((header, index) => {
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
    select.className = "cleaner-map-select";
    select.dataset.columnIndex = String(index);
    const removeOption = new Option("Remove from submission", "remove");
    select.appendChild(removeOption);
    DATA_FIELDS.forEach((field) => select.appendChild(new Option(`Keep as ${field.label}`, field.key)));
    select.value = state.mapping[index]?.action || "remove";
    select.addEventListener("change", () => {
      const nextField = DATA_FIELDS.find((field) => field.key === select.value);
      state.mapping[index] = {
        action: select.value,
        reason: select.value === "remove" ? "Manually removed." : `Manually kept as ${nextField?.label || select.value}.`,
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
    els.mappingBody.appendChild(row);
  });
  updateCounters();
}

function updateCounters() {
  const actions = Object.values(state.mapping).map((item) => item.action);
  els.keptCount.textContent = String(actions.filter((action) => action !== "remove").length);
  els.removedCount.textContent = String(actions.filter((action) => action === "remove").length);
  els.rowCount.textContent = String(state.rows.length);
}

function buildCleanRows() {
  const kept = Object.entries(state.mapping)
    .map(([index, config]) => ({ index: Number(index), key: config.action }))
    .filter((item) => item.key && item.key !== "remove");
  const outputHeaders = [];
  const seen = new Map();
  kept.forEach((item) => {
    const count = seen.get(item.key) || 0;
    seen.set(item.key, count + 1);
    outputHeaders.push(count ? `${item.key}_${count + 1}` : item.key);
  });
  const rows = state.rows.map((row) => kept.map((item) => String(row[item.index] || "").trim()));
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
  const cleaned = buildCleanRows();
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
  els.reviewMeta.textContent = `${cleaned.rows.length} rows, ${cleaned.headers.length} sanitized columns.`;
}

function downloadCleanCsv() {
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

async function submitCleanData() {
  const cleaned = buildCleanRows();
  if (!cleaned.headers.length || !cleaned.rows.length) {
    setStatus("Keep at least one useful column before submitting.", "error");
    return;
  }
  els.submit.disabled = true;
  setStatus("Submitting cleaned data to Shipide...");
  try {
    const response = await fetch("/api/public/shipping-data-cleaner/submit", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        fileName: state.fileName,
        contactEmail: els.contactEmail.value.trim(),
        headers: cleaned.headers,
        rows: cleaned.rows,
        removedColumns: state.headers.filter((_, index) => state.mapping[index]?.action === "remove"),
      }),
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(payload?.error || "Could not submit the cleaned data.");
    setStatus("Cleaned data submitted. We only received the sanitized columns shown above.", "success");
  } catch (error) {
    setStatus(`${error.message || "Submission failed."} Download the cleaned CSV and send it to Shipide if needed.`, "error");
  } finally {
    els.submit.disabled = false;
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
  els.fileMeta.textContent = `${state.fileName} - ${state.rows.length} rows - ${state.headers.length} columns`;
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
  renderPreview();
  setStep("review");
});
els.download.addEventListener("click", downloadCleanCsv);
els.downloadFromMap.addEventListener("click", downloadCleanCsv);
els.submit.addEventListener("click", submitCleanData);

drawBackground();
