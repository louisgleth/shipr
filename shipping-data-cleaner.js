const DATA_FIELDS = [
  {
    key: "shipment_date",
    label: "Shipment date",
    aliases: [
      "date",
      "shipment_date",
      "created_at",
      "order_date",
      "shipping_date",
      "label_date",
      "verzenddatum",
      "verzend_datum",
      "datum",
      "date_expedition",
      "date_d_expedition",
      "date_envoi",
    ],
  },
  {
    key: "origin_postcode",
    label: "Origin postcode",
    aliases: [
      "origin_postcode",
      "origin_postal_code",
      "from_postcode",
      "from_zip",
      "postal_code_from",
      "postcode_from",
      "zip_from",
      "shipping_from_zip",
      "shipping_from_postcode",
      "shipping_from_postal_code",
      "ship_from_postcode",
      "ship_from_postal_code",
      "ship_from_zip",
      "shipping_origin_postcode",
      "origin_zip",
      "sender_zip",
      "warehouse_postcode",
      "postcode_van_herkomst",
      "postcode_herkomst",
      "herkomst_postcode",
      "afzender_postcode",
      "code_postal_origine",
      "code_postal_expediteur",
      "cp_origine",
    ],
  },
  {
    key: "origin_country",
    label: "Origin country",
    aliases: [
      "origin_country",
      "from_country",
      "country_from",
      "shipping_from_country",
      "ship_from_country",
      "shipping_origin_country",
      "sender_country",
      "warehouse_country",
      "land_van_herkomst",
      "land_herkomst",
      "herkomst_land",
      "afzender_land",
      "pays_origine",
      "pays_d_origine",
      "pays_expediteur",
    ],
  },
  {
    key: "destination_country",
    label: "Destination country",
    aliases: [
      "destination_country",
      "dest_country",
      "to_country",
      "shipping_country",
      "recipient_country",
      "country",
      "land_bestemming",
      "bestemmingsland",
      "bestemming_land",
      "ontvanger_land",
      "pays_destination",
      "pays_de_destination",
      "pays_destinataire",
    ],
  },
  {
    key: "destination_postcode",
    label: "Destination postcode",
    aliases: [
      "destination_postcode",
      "destination_postal_code",
      "dest_zip",
      "to_zip",
      "shipping_zip",
      "recipient_postcode",
      "postcode",
      "postal_code",
      "postcode_bestemming",
      "bestemming_postcode",
      "postcode_van_bestemming",
      "ontvanger_postcode",
      "code_postal_destination",
      "code_postal_destinataire",
      "cp_destination",
    ],
  },
  {
    key: "weight",
    label: "Weight",
    aliases: [
      "weight",
      "package_weight",
      "parcel_weight",
      "shipment_weight",
      "weight_kg",
      "weight_g",
      "gewicht",
      "pakketgewicht",
      "zendinggewicht",
      "poids",
      "poids_colis",
      "poids_expedition",
    ],
  },
  {
    key: "weight_unit",
    label: "Weight unit",
    aliases: ["weight_unit", "unit_weight", "weight_units", "gewichtseenheid", "eenheid_gewicht", "unite_poids"],
  },
  {
    key: "length",
    label: "Length",
    aliases: ["length", "package_length", "parcel_length", "dim_length", "lengte", "longueur"],
  },
  {
    key: "width",
    label: "Width",
    aliases: ["width", "package_width", "parcel_width", "dim_width", "breedte", "largeur"],
  },
  {
    key: "height",
    label: "Height",
    aliases: ["height", "package_height", "parcel_height", "dim_height", "hoogte", "hauteur"],
  },
  {
    key: "dimensions",
    label: "Dimensions",
    aliases: [
      "dimensions",
      "dims",
      "size",
      "package_dims",
      "parcel_dimensions",
      "afmetingen",
      "pakketafmetingen",
      "dimensies",
    ],
  },
  {
    key: "dimension_unit",
    label: "Dimension unit",
    aliases: ["dimension_unit", "dims_unit", "length_unit", "unit_dimension", "afmetingseenheid", "unite_dimension"],
  },
  {
    key: "service_type",
    label: "Service type",
    aliases: [
      "service",
      "service_type",
      "shipping_method",
      "carrier_service",
      "delivery_method",
      "verzendmethode",
      "verzendservice",
      "dienst",
      "transporteur",
      "type_service",
      "mode_expedition",
      "methode_expedition",
    ],
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

const LOCALE = (() => {
  const preferredLanguage = Array.isArray(navigator.languages) && navigator.languages.length
    ? navigator.languages[0]
    : navigator.language || navigator.userLanguage || "en";
  const language = String(preferredLanguage).toLowerCase();
  if (language.startsWith("nl")) return "nl";
  if (language.startsWith("fr")) return "fr";
  return "en";
})();

const COPY = {
  en: {
    documentTitle: "Shipide Data Cleaner",
    step: "Step {index} of 3",
    complete: "Complete",
    brand: "Data Cleaner",
    kicker: "Shipment extract sanitization",
    uploadTitle: "Clean your shipping data before sending it to Shipide.",
    mappingTitle: "Review detected columns",
    originTitleStep: "Add ship-from address",
    reviewTitle: "Cleaned file preview",
    thanksTitleStep: "Submission complete",
    intro:
      "Upload your carrier, marketplace, or store export. We keep shipment-analysis fields and strip names, full addresses, emails, phones, notes, and identifiers before submission.",
    downloadCsv: "Download CSV",
    dropStrong: "Drop your CSV, TXT, TSV, XLS export here",
    dropSub: "or browse from your computer",
    uploadAria: "Upload shipping export",
    keptByDefault: "Kept by default",
    keptByDefaultCopy: "Date, origin and destination country/postcode, weight, dimensions, service type.",
    removedByDefault: "Removed by default",
    removedByDefaultCopy: "Names, street addresses, email, phone, company, order IDs, notes, tracking references.",
    uploadedColumn: "Uploaded column",
    action: "Action",
    sample: "Sample",
    reason: "Reason",
    changeFile: "Change file",
    previewCleanedData: "Preview cleaned data",
    editMapping: "Edit mapping",
    editOriginAddress: "Edit ship-from address",
    submitToShipide: "Submit to Shipide",
    originCopy:
      "We could not identify a ship-from address in this export. Add the address you usually ship from so Shipide can analyze your shipment profile correctly.",
    originPlaceholder: "Company, street, postcode, city, country",
    addOriginAddress: "Add another ship-from address",
    removeOriginAddress: "Remove",
    originRequired: "Add at least one ship-from address before continuing.",
    dataReceived: "Data received",
    thankYou: "Thank you.",
    thanksCopy: "We received the cleaned shipment data. We’ll review your shipment profile and get back to you as soon as possible.",
    submissionCheck: "Submission Check",
    confirmTitle: "Have you checked that the provided information is truthful?",
    confirmNote: "Please confirm that the sanitized shipment data you are sending is accurate and provided in good faith.",
    confirmCheckbox: "I confirm the information provided is truthful and accurate.",
    checked: "I've Checked",
    reviewAgain: "Review Again",
    removeFromSubmission: "Remove from submission",
    keepAs: "Keep as {label}",
    extractedSource: "extracted, not raw address",
    noSample: "No sample",
    untitled: "untitled",
    column: "Column {index}",
    likelyPersonal: "Likely personal or identifying data.",
    matchedUseful: "Matched useful shipment field.",
    looksUseful: "Looks like shipment-analysis data.",
    selectFieldOptional: "Select at least one shipment-analysis field. Dimensions, service type, and date are optional.",
    selectFieldBeforeSubmit: "Select at least one shipment-analysis field before submitting. Missing dimensions or service type will not block submission.",
    confirmTruthful: "Confirm the information is truthful before submitting.",
    invalidLink: "This cleaner link is invalid or expired. Ask Shipide for a new link.",
    submitting: "Submitting cleaned data to Shipide...",
    submitFailedSuffix: "Download the cleaned CSV and send it to Shipide if needed.",
    submissionFailed: "Submission failed.",
    noUsableTable: "No usable table was detected. Export the file as CSV or TXT and try again.",
    linkNotReady: "This cleaner link is not ready. Ask Shipide for a valid shipment extract link.",
    readFailed: "Could not read the file. Please try again.",
    parseFailed: "Could not parse this file.",
    xlsxDisabled: "XLSX parsing is not enabled in this first version. Export as CSV or TXT and upload that file.",
    tokenRequired: "This cleaner page requires a client-specific Shipide link.",
    linkInvalidOrExpired: "This shipment extract link is invalid or expired.",
    selectAtLeastOne: "Select at least one shipment-analysis field. Missing optional fields are fine.",
    submitFailed: "Could not submit the cleaned data.",
    fields: {
      shipment_date: "Shipment date",
      origin_postcode: "Origin postcode",
      origin_country: "Origin country",
      destination_country: "Destination country",
      destination_postcode: "Destination postcode",
      weight: "Weight",
      weight_unit: "Weight unit",
      length: "Length",
      width: "Width",
      height: "Height",
      dimensions: "Dimensions",
      dimension_unit: "Dimension unit",
      service_type: "Service type",
    },
  },
  fr: {
    documentTitle: "Nettoyeur de données Shipide",
    step: "Étape {index} sur 3",
    complete: "Terminé",
    brand: "Nettoyeur de données",
    kicker: "Nettoyage d’extrait d’expédition",
    uploadTitle: "Nettoyez vos données d’expédition avant de les envoyer à Shipide.",
    mappingTitle: "Vérifier les colonnes détectées",
    originTitleStep: "Ajouter l’adresse d’expédition",
    reviewTitle: "Aperçu du fichier nettoyé",
    thanksTitleStep: "Envoi terminé",
    intro:
      "Importez votre export transporteur, marketplace ou boutique. Nous conservons les champs utiles à l’analyse d’expédition et retirons les noms, adresses complètes, e-mails, téléphones, notes et identifiants avant l’envoi.",
    downloadCsv: "Télécharger le CSV",
    dropStrong: "Déposez votre export CSV, TXT, TSV ou XLS ici",
    dropSub: "ou parcourez votre ordinateur",
    uploadAria: "Importer un export d’expédition",
    keptByDefault: "Conservé par défaut",
    keptByDefaultCopy: "Date, pays/code postal d’origine et de destination, poids, dimensions, type de service.",
    removedByDefault: "Retiré par défaut",
    removedByDefaultCopy: "Noms, adresses, e-mails, téléphones, société, numéros de commande, notes, références de suivi.",
    uploadedColumn: "Colonne importée",
    action: "Action",
    sample: "Exemple",
    reason: "Raison",
    changeFile: "Changer le fichier",
    previewCleanedData: "Prévisualiser les données nettoyées",
    editMapping: "Modifier le mapping",
    editOriginAddress: "Modifier l’adresse d’expédition",
    submitToShipide: "Envoyer à Shipide",
    originCopy:
      "Nous n’avons pas identifié d’adresse d’expédition dans cet export. Ajoutez l’adresse depuis laquelle vous expédiez habituellement afin que Shipide puisse analyser correctement votre profil d’expédition.",
    originPlaceholder: "Société, rue, code postal, ville, pays",
    addOriginAddress: "Ajouter une autre adresse d’expédition",
    removeOriginAddress: "Retirer",
    originRequired: "Ajoutez au moins une adresse d’expédition avant de continuer.",
    dataReceived: "Données reçues",
    thankYou: "Merci.",
    thanksCopy: "Nous avons reçu les données d’expédition nettoyées. Nous analyserons votre profil d’expédition et reviendrons vers vous dès que possible.",
    submissionCheck: "Vérification de l’envoi",
    confirmTitle: "Avez-vous vérifié que les informations fournies sont exactes ?",
    confirmNote: "Veuillez confirmer que les données d’expédition nettoyées que vous envoyez sont exactes et fournies de bonne foi.",
    confirmCheckbox: "Je confirme que les informations fournies sont exactes.",
    checked: "J’ai vérifié",
    reviewAgain: "Revoir",
    removeFromSubmission: "Retirer de l’envoi",
    keepAs: "Conserver comme {label}",
    extractedSource: "extrait, pas l’adresse brute",
    noSample: "Aucun exemple",
    untitled: "sans titre",
    column: "Colonne {index}",
    likelyPersonal: "Données probablement personnelles ou identifiantes.",
    matchedUseful: "Champ utile d’expédition reconnu.",
    looksUseful: "Semble être une donnée utile d’expédition.",
    selectFieldOptional: "Sélectionnez au moins un champ d’analyse d’expédition. Les dimensions, le service et la date sont optionnels.",
    selectFieldBeforeSubmit: "Sélectionnez au moins un champ d’analyse d’expédition avant l’envoi. Les dimensions ou le service manquants ne bloquent pas l’envoi.",
    confirmTruthful: "Confirmez que les informations sont exactes avant l’envoi.",
    invalidLink: "Ce lien de nettoyage est invalide ou expiré. Demandez un nouveau lien à Shipide.",
    submitting: "Envoi des données nettoyées à Shipide...",
    submitFailedSuffix: "Téléchargez le CSV nettoyé et envoyez-le à Shipide si nécessaire.",
    submissionFailed: "Échec de l’envoi.",
    noUsableTable: "Aucun tableau exploitable n’a été détecté. Exportez le fichier en CSV ou TXT puis réessayez.",
    linkNotReady: "Ce lien de nettoyage n’est pas prêt. Demandez un lien d’extrait valide à Shipide.",
    readFailed: "Impossible de lire le fichier. Veuillez réessayer.",
    parseFailed: "Impossible d’analyser ce fichier.",
    xlsxDisabled: "L’analyse XLSX n’est pas activée dans cette première version. Exportez en CSV ou TXT puis importez ce fichier.",
    tokenRequired: "Cette page nécessite un lien Shipide spécifique au client.",
    linkInvalidOrExpired: "Ce lien d’extrait d’expédition est invalide ou expiré.",
    selectAtLeastOne: "Sélectionnez au moins un champ d’analyse d’expédition. Les champs optionnels manquants ne posent pas problème.",
    submitFailed: "Impossible d’envoyer les données nettoyées.",
    fields: {
      shipment_date: "Date d’expédition",
      origin_postcode: "Code postal d’origine",
      origin_country: "Pays d’origine",
      destination_country: "Pays de destination",
      destination_postcode: "Code postal de destination",
      weight: "Poids",
      weight_unit: "Unité de poids",
      length: "Longueur",
      width: "Largeur",
      height: "Hauteur",
      dimensions: "Dimensions",
      dimension_unit: "Unité de dimension",
      service_type: "Type de service",
    },
  },
  nl: {
    documentTitle: "Shipide Data Cleaner",
    step: "Stap {index} van 3",
    complete: "Voltooid",
    brand: "Data Cleaner",
    kicker: "Opschonning van verzendextract",
    uploadTitle: "Schoon je verzenddata op voordat je die naar Shipide stuurt.",
    mappingTitle: "Gedetecteerde kolommen controleren",
    originTitleStep: "Verzendadres toevoegen",
    reviewTitle: "Voorbeeld van opgeschoond bestand",
    thanksTitleStep: "Indiening voltooid",
    intro:
      "Upload je export van vervoerder, marketplace of webshop. We bewaren velden voor verzendanalyse en verwijderen namen, volledige adressen, e-mails, telefoons, notities en identificatoren vóór verzending.",
    downloadCsv: "CSV downloaden",
    dropStrong: "Sleep je CSV-, TXT-, TSV- of XLS-export hierheen",
    dropSub: "of kies een bestand op je computer",
    uploadAria: "Verzendexport uploaden",
    keptByDefault: "Standaard behouden",
    keptByDefaultCopy: "Datum, land/postcode van oorsprong en bestemming, gewicht, afmetingen, servicetype.",
    removedByDefault: "Standaard verwijderd",
    removedByDefaultCopy: "Namen, straatadressen, e-mail, telefoon, bedrijf, order-ID’s, notities, trackingreferenties.",
    uploadedColumn: "Geüploade kolom",
    action: "Actie",
    sample: "Voorbeeld",
    reason: "Reden",
    changeFile: "Bestand wijzigen",
    previewCleanedData: "Opgeschoonde data bekijken",
    editMapping: "Mapping aanpassen",
    editOriginAddress: "Verzendadres aanpassen",
    submitToShipide: "Naar Shipide sturen",
    originCopy:
      "We konden geen verzendadres van oorsprong in deze export vinden. Voeg het adres toe vanwaar je meestal verzendt, zodat Shipide je verzendprofiel correct kan analyseren.",
    originPlaceholder: "Bedrijf, straat, postcode, stad, land",
    addOriginAddress: "Nog een verzendadres toevoegen",
    removeOriginAddress: "Verwijderen",
    originRequired: "Voeg minstens één verzendadres toe voordat je verdergaat.",
    dataReceived: "Data ontvangen",
    thankYou: "Bedankt.",
    thanksCopy: "We hebben de opgeschoonde verzenddata ontvangen. We bekijken je verzendprofiel en komen zo snel mogelijk bij je terug.",
    submissionCheck: "Controle voor verzending",
    confirmTitle: "Heb je gecontroleerd dat de verstrekte informatie waarheidsgetrouw is?",
    confirmNote: "Bevestig dat de opgeschoonde verzenddata die je verzendt correct is en te goeder trouw wordt aangeleverd.",
    confirmCheckbox: "Ik bevestig dat de verstrekte informatie waarheidsgetrouw en correct is.",
    checked: "Ik heb gecontroleerd",
    reviewAgain: "Opnieuw bekijken",
    removeFromSubmission: "Verwijderen uit inzending",
    keepAs: "Behouden als {label}",
    extractedSource: "geëxtraheerd, niet het ruwe adres",
    noSample: "Geen voorbeeld",
    untitled: "zonder titel",
    column: "Kolom {index}",
    likelyPersonal: "Waarschijnlijk persoonlijke of identificerende data.",
    matchedUseful: "Nuttig verzendveld herkend.",
    looksUseful: "Lijkt op verzendanalyse-data.",
    selectFieldOptional: "Selecteer minstens één verzendanalyseveld. Afmetingen, servicetype en datum zijn optioneel.",
    selectFieldBeforeSubmit: "Selecteer minstens één verzendanalyseveld voordat je indient. Ontbrekende afmetingen of servicetype blokkeren de inzending niet.",
    confirmTruthful: "Bevestig dat de informatie waarheidsgetrouw is voordat je indient.",
    invalidLink: "Deze cleaner-link is ongeldig of verlopen. Vraag Shipide om een nieuwe link.",
    submitting: "Opgeschoonde data naar Shipide verzenden...",
    submitFailedSuffix: "Download de opgeschoonde CSV en stuur die indien nodig naar Shipide.",
    submissionFailed: "Indienen mislukt.",
    noUsableTable: "Er werd geen bruikbare tabel gevonden. Exporteer het bestand als CSV of TXT en probeer opnieuw.",
    linkNotReady: "Deze cleaner-link is niet klaar. Vraag Shipide om een geldige extractlink.",
    readFailed: "Het bestand kon niet worden gelezen. Probeer opnieuw.",
    parseFailed: "Dit bestand kon niet worden verwerkt.",
    xlsxDisabled: "XLSX-verwerking is niet ingeschakeld in deze eerste versie. Exporteer als CSV of TXT en upload dat bestand.",
    tokenRequired: "Deze pagina vereist een klant-specifieke Shipide-link.",
    linkInvalidOrExpired: "Deze verzendextractlink is ongeldig of verlopen.",
    selectAtLeastOne: "Selecteer minstens één verzendanalyseveld. Ontbrekende optionele velden zijn geen probleem.",
    submitFailed: "De opgeschoonde data kon niet worden ingediend.",
    fields: {
      shipment_date: "Verzenddatum",
      origin_postcode: "Postcode oorsprong",
      origin_country: "Land oorsprong",
      destination_country: "Land bestemming",
      destination_postcode: "Postcode bestemming",
      weight: "Gewicht",
      weight_unit: "Gewichtseenheid",
      length: "Lengte",
      width: "Breedte",
      height: "Hoogte",
      dimensions: "Afmetingen",
      dimension_unit: "Afmetingseenheid",
      service_type: "Servicetype",
    },
  },
};

const copy = COPY[LOCALE] || COPY.en;

function t(key, replacements = {}) {
  const value = key.includes(".")
    ? key.split(".").reduce((acc, part) => (acc && acc[part] != null ? acc[part] : undefined), copy)
    : copy[key];
  const fallback = key.includes(".")
    ? key.split(".").reduce((acc, part) => (acc && acc[part] != null ? acc[part] : undefined), COPY.en)
    : COPY.en[key];
  return String(value || fallback || key).replace(/\{(\w+)\}/g, (_, name) =>
    replacements[name] == null ? "" : String(replacements[name])
  );
}

function fieldLabel(key) {
  return t(`fields.${key}`);
}

const REMOVE_PATTERNS = [
  /(^|_)(name|first_name|last_name|full_name|customer|buyer|consignee)(_|$)/,
  /(^|_)(email|e_mail|mail)(_|$)/,
  /(^|_)(phone|tel|mobile|gsm)(_|$)/,
  /(^|_)(address|street|address1|address_1|line_1|line1|address2|line_2|line2)(_|$)/,
  /(^|_)(city|town|commune)(_|$)/,
  /(^|_)(company|organization|organisation)(_|$)/,
  /(^|_)(order|order_id|order_number|tracking|tracking_number|reference|note|notes|comment)(_|$)/,
];

const SHOPIFY_EXACT_FIELD_MAP = {
  created_at: "shipment_date",
  shipping_method: "service_type",
  shipping_zip: "destination_postcode",
  shipping_country: "destination_country",
};

const SHOPIFY_EXACT_REMOVE_HEADERS = new Set([
  "name",
  "email",
  "financial_status",
  "paid_at",
  "fulfillment_status",
  "fulfilled_at",
  "accepts_marketing",
  "currency",
  "subtotal",
  "shipping",
  "taxes",
  "total",
  "discount_code",
  "discount_amount",
  "lineitem_name",
  "lineitem_price",
  "lineitem_compare_at_price",
  "lineitem_sku",
  "lineitem_requires_shipping",
  "lineitem_taxable",
  "lineitem_fulfillment_status",
  "cancelled_at",
  "payment_method",
  "payment_reference",
  "refunded_amount",
  "vendor",
  "outstanding_balance",
  "employee",
  "location",
  "device_id",
  "id",
  "tags",
  "risk_level",
  "source",
  "lineitem_discount",
  "receipt_number",
  "duties",
  "payment_id",
  "payment_terms_name",
  "next_payment_due_at",
  "payment_references",
]);

const SHOPIFY_REMOVE_PREFIXES = ["billing_", "tax_"];

const CLEANER_TOKEN_STORAGE_KEY = "shipide-shipment-extract-token";

function readCleanerTokenFromUrl() {
  const params = new URLSearchParams(window.location.search);
  const directToken = params.get("token") || params.get("extract") || "";
  if (directToken) return String(directToken).trim();
  const hash = String(window.location.hash || "").replace(/^#/, "");
  if (!hash) return "";
  const hashQuery = hash.includes("?") ? hash.slice(hash.indexOf("?") + 1) : hash;
  return String(new URLSearchParams(hashQuery).get("token") || "").trim();
}

function getCurrentCleanerToken() {
  const urlToken = readCleanerTokenFromUrl();
  if (urlToken) {
    try {
      window.sessionStorage?.setItem(CLEANER_TOKEN_STORAGE_KEY, urlToken);
    } catch (_error) {}
    return urlToken;
  }
  try {
    return String(window.sessionStorage?.getItem(CLEANER_TOKEN_STORAGE_KEY) || "").trim();
  } catch (_error) {
    return "";
  }
}

const state = {
  fileName: "",
  headers: [],
  rows: [],
  mapping: {},
  extractedMapping: {},
  manualOriginAddresses: [""],
  mappingOrder: [],
  extractedInsertAfter: 0,
  step: "upload",
  transitionToken: 0,
  requestToken: getCurrentCleanerToken(),
  requestEmail: "",
  requestReady: false,
};

const STEP_OUT_MS = 220;
const STEP_RESIZE_MS = 520;
const STEP_IN_MS = 260;
const STEP_CARD_CLASSES = [
  "is-step-upload",
  "is-step-mapping",
  "is-step-origin",
  "is-step-review",
  "is-step-thanks",
];

const els = {
  panels: {
    upload: document.getElementById("uploadPanel"),
    mapping: document.getElementById("mappingPanel"),
    origin: document.getElementById("originPanel"),
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
  backToMappingFromOrigin: document.getElementById("cleanerBackToMappingFromOrigin"),
  continueBtn: document.getElementById("cleanerContinue"),
  originContinue: document.getElementById("cleanerOriginContinue"),
  originInputs: document.getElementById("cleanerOriginInputs"),
  addOrigin: document.getElementById("cleanerAddOrigin"),
  download: document.getElementById("cleanerDownload"),
  downloadFromMap: document.getElementById("cleanerDownloadFromMap"),
  submit: document.getElementById("cleanerSubmit"),
  pixelCanvas: document.getElementById("cleanerPixelCanvas"),
};

function setText(selector, value) {
  const element = document.querySelector(selector);
  if (element) element.textContent = value;
}

function applyLocalization() {
  document.documentElement.lang = LOCALE;
  document.title = t("documentTitle");
  setText(".cleaner-brand strong", t("brand"));
  if (els.fileInput) els.fileInput.setAttribute("aria-label", t("uploadAria"));
  setText(".cleaner-kicker", t("kicker"));
  const kickerDot = document.createElement("span");
  document.querySelector(".cleaner-kicker")?.prepend(kickerDot);
  const intro = document.getElementById("cleanerIntroCopy");
  if (intro) intro.textContent = t("intro");
  document.querySelectorAll("#cleanerDownloadFromMap span, #cleanerDownload span").forEach((node) => {
    node.textContent = t("downloadCsv");
  });
  setText(".cleaner-dropzone strong", t("dropStrong"));
  setText(".cleaner-dropzone span", t("dropSub"));
  const specs = document.querySelectorAll(".cleaner-spec-item");
  if (specs[0]) {
    specs[0].querySelector(".cleaner-spec-label").textContent = t("keptByDefault");
    specs[0].querySelector("p").textContent = t("keptByDefaultCopy");
  }
  if (specs[1]) {
    specs[1].querySelector(".cleaner-spec-label").textContent = t("removedByDefault");
    specs[1].querySelector("p").textContent = t("removedByDefaultCopy");
  }
  const mapHeaders = document.querySelectorAll(".cleaner-map-table th");
  [t("uploadedColumn"), t("action"), t("sample"), t("reason")].forEach((label, index) => {
    if (mapHeaders[index]) mapHeaders[index].textContent = label;
  });
  setText("#cleanerBackToUpload span", t("changeFile"));
  setText("#cleanerContinue span", t("previewCleanedData"));
  setText("#cleanerBackToMapping span", t("editMapping"));
  setText("#cleanerBackToMappingFromOrigin span", t("editMapping"));
  setText("#cleanerSubmit span", t("submitToShipide"));
  setText("#cleanerOriginContinue span", t("previewCleanedData"));
  setText("#cleanerOriginCopy", t("originCopy"));
  setText("#cleanerAddOrigin span:last-child", t("addOriginAddress"));
  setText(".cleaner-thanks-kicker", t("dataReceived"));
  const thanksKickerDot = document.createElement("span");
  document.querySelector(".cleaner-thanks-kicker")?.prepend(thanksKickerDot);
  setText(".cleaner-thanks-copy h2", t("thankYou"));
  setText(".cleaner-thanks-copy p:not(.cleaner-thanks-kicker)", t("thanksCopy"));
  setText(".cleaner-confirm-kicker span:last-child", t("submissionCheck"));
  setText("#cleanerConfirmTitle", t("confirmTitle"));
  setText(".cleaner-confirm-note", t("confirmNote"));
  setText(".cleaner-confirm-check > span:last-child", t("confirmCheckbox"));
  setText("#cleanerConfirmApprove span", t("checked"));
  setText("#cleanerConfirmCancel span", t("reviewAgain"));
}

function setUploadEnabled(enabled) {
  const isEnabled = Boolean(enabled);
  if (els.fileInput) {
    els.fileInput.disabled = !isEnabled;
  }
  if (els.dropzone) {
    els.dropzone.classList.toggle("is-disabled", !isEnabled);
    els.dropzone.setAttribute("aria-disabled", String(!isEnabled));
  }
}

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
  if (step === "origin") return Math.min(760, shellWidth);
  return Math.min(1040, shellWidth);
}

function getStepTitle(step) {
  const titles = {
    upload: t("uploadTitle"),
    mapping: t("mappingTitle"),
    origin: t("originTitleStep"),
    review: t("reviewTitle"),
    thanks: t("thanksTitleStep"),
  };
  return titles[step] || titles.upload;
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

function measureStepCardHeight(step, panel, width) {
  if (!panel || !els.card) return 0;
  const previousCardClasses = STEP_CARD_CLASSES.filter((className) => els.card.classList.contains(className));
  const previousCardWidth = els.card.style.width;
  const previousCardHeight = els.card.style.height;
  const previousTitle = els.title?.textContent || "";
  const previousPanelClasses = new Map(
    Object.entries(els.panels).map(([key, item]) => [key, item.className])
  );

  els.card.classList.remove(...STEP_CARD_CLASSES);
  els.card.classList.add(`is-step-${step}`);
  els.card.style.width = `${Math.round(width)}px`;
  els.card.style.height = "auto";
  if (els.title) els.title.textContent = getStepTitle(step);

  Object.entries(els.panels).forEach(([key, item]) => {
    item.className = previousPanelClasses.get(key) || item.className;
    item.classList.toggle("is-active", item === panel);
    item.classList.remove(
      "is-step-transition-exit",
      "is-step-transition-hidden",
      "is-step-transition-enter-start"
    );
  });

  const height = Math.max(1, els.card.getBoundingClientRect().height);

  els.card.classList.remove(...STEP_CARD_CLASSES);
  previousCardClasses.forEach((className) => els.card.classList.add(className));
  els.card.style.width = previousCardWidth;
  els.card.style.height = previousCardHeight;
  if (els.title) els.title.textContent = previousTitle;
  Object.entries(els.panels).forEach(([key, item]) => {
    item.className = previousPanelClasses.get(key) || item.className;
  });

  return height;
}

function applyStepView(step) {
  els.card?.classList.remove(...STEP_CARD_CLASSES);
  els.card?.classList.add(`is-step-${step}`);
  Object.entries(els.panels).forEach(([key, panel]) => {
    panel.classList.toggle("is-active", key === step);
    panel.classList.remove(
      "is-step-transition-exit",
      "is-step-transition-hidden",
      "is-step-transition-enter-start"
    );
  });
  const stepIndex = step === "upload" ? 1 : step === "mapping" || step === "origin" ? 2 : 3;
  const stepTotal = 3;
  if (els.title) {
    els.title.textContent = getStepTitle(step);
  }
  const reviewBackLabel = step === "review" && needsOrigin ? t("editOriginAddress") : t("editMapping");
  setText("#cleanerBackToMapping span", reviewBackLabel);
  els.progress.forEach((item) => {
    const itemStep = Number(item.dataset.progressStep);
    item.hidden = itemStep > stepTotal;
    item.classList.toggle("is-active", step === "thanks" || itemStep <= stepIndex);
  });
  els.stepLabel.textContent = step === "thanks" ? t("complete") : t("step", { index: stepIndex });
  document.querySelector(".cleaner-progress")?.style.setProperty("--cleaner-progress-steps", String(stepTotal));
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
  const nextStep = ["upload", "mapping", "origin", "review", "thanks"].includes(step) ? step : "upload";
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
  const targetWidth = getStepWidth(nextStep);
  const targetHeight = measureStepCardHeight(nextStep, incomingPanel, targetWidth);

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

function detectShopifyField(normalized) {
  if (SHOPIFY_EXACT_FIELD_MAP[normalized]) {
    return { action: SHOPIFY_EXACT_FIELD_MAP[normalized], reason: t("matchedUseful") };
  }
  if (
    SHOPIFY_EXACT_REMOVE_HEADERS.has(normalized) ||
    SHOPIFY_REMOVE_PREFIXES.some((prefix) => normalized.startsWith(prefix))
  ) {
    return { action: "remove", reason: t("likelyPersonal") };
  }
  return null;
}

function detectField(header) {
  const normalized = normalizeHeader(header);
  const shopifyField = detectShopifyField(normalized);
  if (shopifyField) return shopifyField;
  if (isOriginAddressHeader(normalized)) {
    return { action: "remove", reason: t("likelyPersonal") };
  }
  if (isDestinationAddressHeader(normalized)) {
    return {
      action: "remove",
      reason: t("likelyPersonal"),
    };
  }
  const direct = DATA_FIELDS.find((field) => {
    const aliases = [field.key, field.label, ...field.aliases].map(normalizeHeader);
    return aliases.includes(normalized);
  });
  if (direct) {
    return { action: direct.key, reason: t("matchedUseful") };
  }
  const fuzzy = DATA_FIELDS.find((field) => {
    const aliases = [field.key, ...field.aliases].map(normalizeHeader);
    return aliases.some((alias) => alias && alias.length >= 4 && normalized.includes(alias));
  });
  if (fuzzy) {
    return { action: fuzzy.key, reason: t("looksUseful") };
  }
  if (REMOVE_PATTERNS.some((pattern) => pattern.test(normalized))) {
    return { action: "remove", reason: t("likelyPersonal") };
  }
  return { action: "remove", reason: t("looksUseful") };
}

function isOriginAddressHeader(normalized) {
  return (
    normalized === "shipping_from" ||
    normalized === "ship_from" ||
    normalized.includes("shipped_from_address") ||
    normalized.includes("shipping_from_address") ||
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

function getExtractedFieldCandidates() {
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

function getAvailableExtractedFields() {
  return getExtractedFieldCandidates().filter((field) => {
    const action = state.extractedMapping[field.key]?.action || field.key;
    return action !== "remove";
  });
}

function getKeptFieldKeys() {
  const direct = Object.values(state.mapping)
    .map((item) => item.action)
    .filter((action) => action && action !== "remove");
  return Array.from(new Set([...direct, ...getAvailableExtractedFields().map((field) => field.key)]));
}

function getManualOriginAddresses() {
  return state.manualOriginAddresses.map((address) => String(address || "").trim()).filter(Boolean);
}

function needsManualOriginAddress() {
  const kept = new Set(getKeptFieldKeys());
  return !kept.has("origin_postcode") || !kept.has("origin_country");
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
  return t("noSample");
}

function createMappingRow(header, index) {
  const row = document.createElement("tr");
  const nameCell = document.createElement("td");
  const name = document.createElement("div");
  name.className = "cleaner-column-name";
  name.innerHTML = `<strong></strong><span class="mono"></span>`;
  name.querySelector("strong").textContent = header || t("column", { index: index + 1 });
  name.querySelector("span").textContent = normalizeHeader(header) || t("untitled");
  nameCell.appendChild(name);

  const actionCell = document.createElement("td");
  const select = document.createElement("select");
  select.className = `cleaner-map-select ${state.mapping[index]?.action === "remove" ? "is-remove" : "is-keep"}`;
  select.dataset.columnIndex = String(index);
  const removeOption = new Option(t("removeFromSubmission"), "remove");
  select.appendChild(removeOption);
  DATA_FIELDS.forEach((field) => select.appendChild(new Option(t("keepAs", { label: fieldLabel(field.key) }), field.key)));
  select.value = state.mapping[index]?.action || "remove";
  select.addEventListener("change", () => {
    const nextField = DATA_FIELDS.find((field) => field.key === select.value);
    state.mapping[index] = {
      action: select.value,
      reason:
        select.value === "remove"
          ? t("likelyPersonal")
          : nextField
            ? t("matchedUseful")
            : t("looksUseful"),
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
  name.querySelector("strong").textContent = fieldLabel(field.key);
  name.querySelector("span").textContent = t("extractedSource");
  nameCell.appendChild(name);

  const actionCell = document.createElement("td");
  const select = document.createElement("select");
  const currentAction = state.extractedMapping[field.key]?.action || field.key;
  select.className = `cleaner-map-select ${currentAction === "remove" ? "is-remove" : "is-keep"}`;
  select.appendChild(new Option(t("keepAs", { label: fieldLabel(field.key) }), field.key));
  select.appendChild(new Option(t("removeFromSubmission"), "remove"));
  select.value = currentAction;
  select.addEventListener("change", () => {
    state.extractedMapping[field.key] = {
      action: select.value,
      reason:
        select.value === "remove"
          ? t("likelyPersonal")
          : t("looksUseful"),
    };
    renderMapping();
  });
  actionCell.appendChild(select);

  const sampleCell = document.createElement("td");
  const sample = document.createElement("span");
  sample.className = "cleaner-sample";
  sample.textContent = getExtractedValue(state.rows[0] || [], field.key) || t("noSample");
  sampleCell.appendChild(sample);

  const reasonCell = document.createElement("td");
  reasonCell.className = "cleaner-reason";
  reasonCell.textContent =
    state.extractedMapping[field.key]?.reason || t("looksUseful");

  row.append(nameCell, actionCell, sampleCell, reasonCell);
  return row;
}

function renderMapping() {
  els.mappingBody.innerHTML = "";
  const uploadedColumns = state.headers.map((header, index) => ({ header, index }));
  const orderedColumns = state.mappingOrder.length
    ? state.mappingOrder
        .map((index) => uploadedColumns.find((item) => item.index === index))
        .filter(Boolean)
    : uploadedColumns;
  orderedColumns.forEach((item, position) => {
    if (position === state.extractedInsertAfter) {
      getExtractedFieldCandidates().forEach((field) => {
        els.mappingBody.appendChild(createExtractedMappingRow(field));
      });
    }
    els.mappingBody.appendChild(createMappingRow(item.header, item.index));
  });
  if (state.extractedInsertAfter >= orderedColumns.length) {
    getExtractedFieldCandidates().forEach((field) => {
      els.mappingBody.appendChild(createExtractedMappingRow(field));
    });
  }
}

function renderOriginInputs() {
  if (!els.originInputs) return;
  if (!state.manualOriginAddresses.length) state.manualOriginAddresses = [""];
  els.originInputs.innerHTML = "";
  state.manualOriginAddresses.forEach((address, index) => {
    const row = document.createElement("div");
    row.className = "cleaner-origin-row";

    const input = document.createElement("input");
    input.type = "text";
    input.value = address;
    input.placeholder = t("originPlaceholder");
    input.autocomplete = "street-address";
    input.dataset.originIndex = String(index);
    input.addEventListener("input", () => {
      state.manualOriginAddresses[index] = input.value;
    });
    row.appendChild(input);

    if (state.manualOriginAddresses.length > 1) {
      const remove = document.createElement("button");
      remove.type = "button";
      remove.className = "cleaner-origin-remove";
      remove.setAttribute("aria-label", t("removeOriginAddress"));
      remove.title = t("removeOriginAddress");
      remove.innerHTML = `
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
          <path d="M4 7h16" />
          <path d="M10 11v6" />
          <path d="M14 11v6" />
          <path d="M6.5 7l.8 13h9.4l.8-13" />
          <path d="M9 7V4h6v3" />
        </svg>
      `;
      remove.addEventListener("click", () => {
        row.classList.add("is-removing");
        window.setTimeout(() => {
          state.manualOriginAddresses.splice(index, 1);
          renderOriginInputs();
        }, 170);
      });
      row.appendChild(remove);
    }

    els.originInputs.appendChild(row);
  });
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
    setStatus(t("selectFieldOptional"), "error");
    return;
  }
  if (needsManualOriginAddress() && !getManualOriginAddresses().length) {
    setStatus(t("originRequired"), "error");
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
    setStatus(t("selectFieldBeforeSubmit"), "error");
    return;
  }
  if (needsManualOriginAddress() && !getManualOriginAddresses().length) {
    setStatus(t("originRequired"), "error");
    return;
  }
  setConfirmModalOpen(true);
}

async function submitCleanData() {
  const { cleaned, isReady } = getCleanedDataReadiness();
  if (!isReady) {
    setStatus(t("selectFieldBeforeSubmit"), "error");
    return;
  }
  if (needsManualOriginAddress() && !getManualOriginAddresses().length) {
    setStatus(t("originRequired"), "error");
    return;
  }
  if (!els.truthConfirm?.checked) {
    setStatus(t("confirmTruthful"), "error");
    return;
  }
  const requestToken = getCurrentCleanerToken();
  state.requestToken = requestToken;
  if (!state.requestReady || !requestToken) {
    setStatus(t("invalidLink"), "error");
    return;
  }
  els.submit.disabled = true;
  if (els.confirmApprove) els.confirmApprove.disabled = true;
  showStatusToast(t("submitting"), { tone: "info" });
  try {
    const response = await fetch("/api/public/clean-data/submit", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        token: requestToken,
        fileName: state.fileName,
        headers: cleaned.headers,
        rows: cleaned.rows,
        removedColumns: state.headers.filter((_, index) => state.mapping[index]?.action === "remove"),
        originAddresses: needsManualOriginAddress() ? getManualOriginAddresses() : [],
      }),
    });
    await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(t("submitFailed"));
    setConfirmModalOpen(false);
    setStep("thanks");
  } catch (error) {
    setStatus(`${error.message || t("submissionFailed")} ${t("submitFailedSuffix")}`, "error");
  } finally {
    els.submit.disabled = false;
    if (els.confirmApprove) {
      els.confirmApprove.disabled = !els.truthConfirm?.checked;
    }
  }
}

function handleMatrix(matrix, fileName) {
  if (!matrix || matrix.length < 2) {
    throw new Error(t("noUsableTable"));
  }
  state.fileName = fileName || "shipping-export.csv";
  state.headers = matrix[0].map((header, index) => String(header || t("column", { index: index + 1 })).trim());
  state.rows = matrix
    .slice(1)
    .filter((row) => row.some((cell) => String(cell || "").trim()));
  state.mapping = {};
  state.extractedMapping = {};
  state.manualOriginAddresses = [""];
  state.headers.forEach((header, index) => {
    state.mapping[index] = detectField(header);
  });
  state.mappingOrder = state.headers
    .map((_, index) => index)
    .sort((left, right) => {
      const leftRemoved = state.mapping[left]?.action === "remove";
      const rightRemoved = state.mapping[right]?.action === "remove";
      if (leftRemoved === rightRemoved) return left - right;
      return leftRemoved ? 1 : -1;
    });
  state.extractedInsertAfter = state.mappingOrder.filter(
    (index) => state.mapping[index]?.action !== "remove"
  ).length;
  renderMapping();
  setStep("mapping");
  setStatus("");
}

function handleFile(file) {
  if (!state.requestReady) {
    setStatus(t("linkNotReady"), "error");
    return;
  }
  if (!file) return;
  const name = file.name || "shipping-export.csv";
  const extension = name.split(".").pop().toLowerCase();
  const reader = new FileReader();
  reader.onerror = () => setStatus(t("readFailed"), "error");
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
      setStatus(error.message || t("parseFailed"), "error");
    }
  };
  if (extension === "xlsx") {
    setStatus(t("xlsxDisabled"), "error");
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
    speed: 0.00002 + Math.random() * 0.000045,
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

async function validateShipmentExtractRequest() {
  setUploadEnabled(false);
  state.requestToken = getCurrentCleanerToken();
  if (!state.requestToken) {
    setStatus(t("tokenRequired"), "error");
    return;
  }
  try {
    const response = await fetch(
      `/api/public/clean-data/request?token=${encodeURIComponent(state.requestToken)}`,
      { headers: { Accept: "application/json" } }
    );
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(t("linkInvalidOrExpired"));
    }
    state.requestReady = true;
    state.requestEmail = String(payload?.request?.clientEmail || "").trim();
    setUploadEnabled(true);
  } catch (error) {
    state.requestReady = false;
    setUploadEnabled(false);
    setStatus(error?.message || t("linkInvalidOrExpired"), "error");
  }
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
els.backToMapping.addEventListener("click", () => setStep(needsManualOriginAddress() ? "origin" : "mapping"));
els.backToMappingFromOrigin?.addEventListener("click", () => setStep("mapping"));
els.continueBtn.addEventListener("click", () => {
  const { isReady } = getCleanedDataReadiness();
  if (!isReady) {
    setStatus(t("selectAtLeastOne"), "error");
    return;
  }
  setStatus("");
  if (needsManualOriginAddress()) {
    renderOriginInputs();
    setStep("origin");
    return;
  }
  renderPreview();
  setStep("review");
});
els.addOrigin?.addEventListener("click", () => {
  state.manualOriginAddresses.push("");
  renderOriginInputs();
  window.setTimeout(() => {
    const inputs = els.originInputs?.querySelectorAll("input");
    inputs?.[inputs.length - 1]?.focus();
  }, 20);
});
els.originContinue?.addEventListener("click", () => {
  if (!getManualOriginAddresses().length) {
    setStatus(t("originRequired"), "error");
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

applyLocalization();
drawBackground();
setStep("upload");
void validateShipmentExtractRequest();
