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
  csvPage: 1,
  csvValidationAttempted: false,
  csvSource: "none",
  shipFromOriginId: "",
  shipFromLockedByProvider: false,
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
const WAREHOUSE_RETURN_FIELD_MAP = {
  senderName: "returnSenderName",
  street: "returnStreet",
  city: "returnCity",
  region: "returnRegion",
  postalCode: "returnPostalCode",
  country: "returnCountry",
};
const WAREHOUSE_SENDER_FIELDS = new Set([
  "senderName",
  "street",
  "city",
  "region",
  "postalCode",
  "country",
]);
const SHELL_TRANSITION_IN_MS = 340;
const MAIN_VIEW_TRANSITION_OUT_MS = 130;
const MAIN_VIEW_TRANSITION_IN_MS = 320;
const AUTH_REGISTER_STEP_OUT_MS = 220;
const AUTH_REGISTER_STEP_RESIZE_MS = 520;
const AUTH_REGISTER_STEP_IN_MS = 260;
const CSV_MODAL_STEP_SWITCH_DELAY_MS = 160;
const CSV_TABLE_PAGE_SIZE = 10;
const VAT_RATE = 0;
const TOTAL_STEPS = 4;
const CUSTOMS_MAX_ITEMS = 12;
const REPORT_RANGE_PRESETS = [7, 30, 90, 365];
const AUTH_AGREEMENT_PDF_WORKER_URL =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
const ROUTE_PATHS = {
  login: "/login",
  register: "/register",
  signupPreview: "/signup-preview",
  account: "/account",
  admin: "/admin",
  leads: "/leads",
  history: "/history",
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
  ROUTE_PATHS.register,
  ROUTE_PATHS.signupPreview,
  ROUTE_PATHS.account,
  ROUTE_PATHS.admin,
  ROUTE_PATHS.leads,
  ROUTE_PATHS.history,
  ROUTE_PATHS.reports,
  ...Object.values(STEP_ROUTE_PATHS),
];
const AUTH_SIGNUP_PREVIEW_TOKEN = "local-signup-preview";
const FLOW_LOGO_JSON_URL = "assets/flow-logo.json";
const AUTH_BACKGROUND_VARIANT_STORAGE_KEY = "shipide-auth-bg-variant";
const LEAD_STACK_META = {
  shopify: {
    label: "Shopify",
    icon: "assets/shopify.svg",
    className: "is-shopify",
  },
  woocommerce: {
    label: "WooCommerce",
    icon: "woocommerce-logo.png",
    className: "is-woocommerce",
  },
  wix: {
    label: "Wix eCommerce",
    icon: "Wix-icon.svg",
    className: "is-wix",
  },
};
const LEAD_OUTCOME_META = {
  new: {
    label: "New",
    summaryBucket: "ready",
    className: "is-new",
  },
  interested_follow_up: {
    label: "Interested",
    summaryBucket: "followUp",
    className: "is-interested",
  },
  mild_follow_up: {
    label: "Mildly interested",
    summaryBucket: "followUp",
    className: "is-mild",
  },
  not_interested: {
    label: "Not interested",
    summaryBucket: "notInterested",
    className: "is-not-interested",
  },
};
const MOCK_LEAD_PROSPECTS = (() => {
  const firstNames = [
    "Claire",
    "Pieter",
    "Amandine",
    "Mateo",
    "Sanne",
    "Marta",
    "Elias",
    "Juliette",
    "Tom",
    "Elena",
    "Noah",
    "Camille",
    "Arthur",
    "Sofia",
    "Luca",
    "Ines",
    "Jonas",
    "Leonie",
    "Nils",
    "Eva",
    "Thiago",
    "Marina",
    "Anton",
    "Lucie",
    "Ruben",
  ];
  const lastNames = [
    "Dupont",
    "de Vries",
    "Leroy",
    "Esposito",
    "Jansen",
    "Henriques",
    "Nordin",
    "Bernard",
    "Verbruggen",
    "Rossi",
    "Moreau",
    "van den Berg",
    "Costa",
    "Lefevre",
    "Svensson",
    "Dubois",
    "Silva",
    "Ribeiro",
    "Schmidt",
    "Martens",
    "Fontaine",
    "Peeters",
    "Lindholm",
    "Navarro",
    "Ferreira",
  ];
  const companyPrefixes = [
    "Atelier",
    "North Harbor",
    "Lune",
    "Verde",
    "Noordline",
    "Rivage",
    "Serein",
    "Fjord",
    "Meridian",
    "Maison",
    "Studio",
    "Nova",
    "Forme",
    "Collective",
    "Harbor",
    "Foundry",
    "Works",
    "Supply",
    "Market",
    "Commerce",
  ];
  const companySuffixes = [
    "Goods",
    "Home",
    "Cycle",
    "Atelier",
    "Collective",
    "Works",
    "Supply",
    "Studio",
    "Living",
    "Essentials",
    "Store",
    "Line",
    "Market",
    "Lab",
    "Partners",
    "Merch",
  ];
  const stacks = ["shopify", "woocommerce", "wix"];
  const countryDialing = [
    { code: "+32", base: "470000000" },
    { code: "+31", base: "610000000" },
    { code: "+33", base: "640000000" },
    { code: "+39", base: "330000000" },
    { code: "+351", base: "910000000" },
    { code: "+46", base: "700000000" },
  ];
  const dispositions = [
    "new",
    "new",
    "new",
    "new",
    "new",
    "interested_follow_up",
    "mild_follow_up",
    "new",
    "not_interested",
    "new",
  ];

  const formatPhone = (dialing, index) => {
    const digits = String(Number(dialing.base) + index * 173).padStart(9, "0");
    return `${dialing.code} ${digits.slice(0, 3)} ${digits.slice(3, 5)} ${digits.slice(5, 7)} ${digits.slice(7, 9)}`;
  };

  return Array.from({ length: 100 }, (_, index) => {
    const first = firstNames[index % firstNames.length];
    const last = lastNames[(index * 3) % lastNames.length];
    const companyPrefix = companyPrefixes[index % companyPrefixes.length];
    const companySuffix = companySuffixes[(index * 5) % companySuffixes.length];
    const techStack = stacks[index % stacks.length];
    const dialing = countryDialing[index % countryDialing.length];
    const month = ((index * 5) % 12) + 1;
    const day = ((index * 7) % 27) + 1;
    const year = 2016 + (index % 9);
    const age = 27 + ((index * 7) % 23);
    const estimatedMonthlyRevenue = 18000 + ((index * 9300) % 162000);
    const estimatedParcelsPerMonth = 140 + ((index * 81) % 1480);
    const companyName = `${companyPrefix} ${companySuffix}`;

    return {
      id: `lead-${index + 1}`,
      name: `${first} ${last}`,
      age,
      phone: formatPhone(dialing, index + 1),
      companyName,
      techStack,
      storeOnlineSince: `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`,
      estimatedMonthlyRevenue,
      estimatedParcelsPerMonth,
      disposition: dispositions[index % dispositions.length],
    };
  });
})();
const leadProspectRepository = createLeadProspectRepository();
const AUTH_SIGNUP_PREVIEW_DATA = {
  inviteToken: AUTH_SIGNUP_PREVIEW_TOKEN,
  credentials: {
    email: "claire@ateliermeridian.com",
    password: "PreviewPass123!",
  },
  invite: {
    companyName: "Atelier Meridian",
    contactName: "Claire Dupont",
    contactEmail: "operations@ateliermeridian.com",
    contactPhone: "+32 2 555 01 29",
    billingAddress: "Avenue Louise 120, 1050 Brussels, Belgium",
    taxId: "BE0123456789",
    customerId: "SHIP-24018",
  },
  contract: {
    id: "local-signup-preview-contract",
    title: "Commercial Agreement",
    version: "v2.0",
    hash: "local-signup-preview-contract-v2.0",
    pdfUrl: "assets/contracts/commAgreement-v1.pdf",
    bodyText:
      "Preview mode only.\n\nThis route is intended for local UI review of the signup flow. The agreement PDF above is loaded from the local project so you can validate the layout, scroll behavior, and acceptance state without creating a live account.",
  },
};
const LANGUAGE_STORAGE_KEY = "shipr-language";
const SUPPORTED_LANGUAGES = new Set(["en", "fr", "nl"]);
const LANGUAGE_LOCALE = {
  en: "en-GB",
  fr: "fr-BE",
  nl: "nl-BE",
};
const TRANSLATION_ATTRS = ["placeholder", "title", "aria-label"];
const TRANSLATIONS = {
  "Portal": { fr: "Portail", nl: "Portaal" },
  "Shipide logo": { fr: "Logo Shipide", nl: "Shipide-logo" },
  "Sign in to continue": { fr: "Connectez-vous pour continuer", nl: "Log in om verder te gaan" },
  "Complete your account registration": {
    fr: "Finalisez l’inscription de votre compte",
    nl: "Voltooi je accountregistratie",
  },
  "Use your email and password.": {
    fr: "Utilisez votre e-mail et votre mot de passe.",
    nl: "Gebruik je e-mailadres en wachtwoord.",
  },
  "Use your invite link to create your secure account and billing profile.": {
    fr: "Utilisez votre lien d’invitation pour créer votre compte sécurisé et votre profil de facturation.",
    nl: "Gebruik je uitnodigingslink om je beveiligde account en facturatieprofiel aan te maken.",
  },
  "A quick 2-step setup to activate your account.": {
    fr: "Une activation rapide en 2 etapes pour votre compte.",
    nl: "Een snelle 2-stappen setup om je account te activeren.",
  },
  "Use your email and password. New users can create an account below.": {
    fr: "Utilisez votre e-mail et votre mot de passe. Les nouveaux utilisateurs peuvent créer un compte ci-dessous.",
    nl: "Gebruik je e-mailadres en wachtwoord. Nieuwe gebruikers kunnen hieronder een account aanmaken.",
  },
  "Email": { fr: "E-mail", nl: "E-mail" },
  "Password": { fr: "Mot de passe", nl: "Wachtwoord" },
  "Confirm Password": { fr: "Confirmer le mot de passe", nl: "Bevestig wachtwoord" },
  "Credentials": { fr: "Identifiants", nl: "Inloggegevens" },
  "Access & Contact": { fr: "Accès & Contact", nl: "Toegang & Contact" },
  "Credentials and primary account contact.": {
    fr: "Identifiants et contact principal du compte.",
    nl: "Inloggegevens en hoofdcontact van het account.",
  },
  "Sign-in credentials and primary billing contact.": {
    fr: "Identifiants de connexion et contact principal de facturation.",
    nl: "Inloggegevens en primair facturatiecontact.",
  },
  "Billing & Compliance": { fr: "Facturation & Conformité", nl: "Facturatie & Compliance" },
  "Monthly invoice profile and tax details.": {
    fr: "Profil de facturation mensuelle et informations fiscales.",
    nl: "Maandelijks factuurprofiel en fiscale gegevens.",
  },
  "Invoice identity and tax details for monthly billing.": {
    fr: "Identité de facturation et informations fiscales pour la facturation mensuelle.",
    nl: "Factuuridentiteit en fiscale gegevens voor maandelijkse facturatie.",
  },
  "Service Agreement": { fr: "Accord de service", nl: "Serviceovereenkomst" },
  "Scroll through the full agreement before accepting to finish registration.": {
    fr: "Faites defiler l accord complet avant d accepter pour finaliser l inscription.",
    nl: "Scroll door de volledige overeenkomst voordat je accepteert om registratie af te ronden.",
  },
  "Loading agreement...": {
    fr: "Chargement de l accord...",
    nl: "Overeenkomst laden...",
  },
  "Preparing agreement...": {
    fr: "Preparation de l accord...",
    nl: "Overeenkomst voorbereiden...",
  },
  "Account Setup": { fr: "Configuration du compte", nl: "Accountinstelling" },
  "Enter account and billing details to continue.": {
    fr: "Entrez les informations du compte et de facturation pour continuer.",
    nl: "Voer account- en facturatiegegevens in om door te gaan.",
  },
  "View Agreement PDF": {
    fr: "Voir le PDF de l accord",
    nl: "Bekijk overeenkomst-PDF",
  },
  "Confirm your password to continue.": {
    fr: "Confirmez votre mot de passe pour continuer.",
    nl: "Bevestig je wachtwoord om verder te gaan.",
  },
  "Scroll to the end of the agreement to unlock acceptance.": {
    fr: "Faites defiler jusqu en bas de l accord pour debloquer l acceptation.",
    nl: "Scroll naar het einde van de overeenkomst om accepteren te ontgrendelen.",
  },
  "I have read the full agreement and I agree to the terms.": {
    fr: "J ai lu l accord complet et j accepte les conditions.",
    nl: "Ik heb de volledige overeenkomst gelezen en ga akkoord met de voorwaarden.",
  },
  "Agreement version {version}": {
    fr: "Version de l accord {version}",
    nl: "Overeenkomstversie {version}",
  },
  "Review complete. You can now confirm agreement.": {
    fr: "Lecture complete. Vous pouvez maintenant confirmer l accord.",
    nl: "Lezen voltooid. Je kunt nu de overeenkomst bevestigen.",
  },
  "Agreement accepted. You can create your account.": {
    fr: "Accord accepte. Vous pouvez creer votre compte.",
    nl: "Overeenkomst geaccepteerd. Je kunt je account aanmaken.",
  },
  "You must review and accept the agreement before registering.": {
    fr: "Vous devez lire et accepter l accord avant de vous inscrire.",
    nl: "Je moet de overeenkomst lezen en accepteren voordat je registreert.",
  },
  "Could not load registration agreement.": {
    fr: "Impossible de charger l accord d inscription.",
    nl: "Kon registratieovereenkomst niet laden.",
  },
  "Could not prepare agreement preview.": {
    fr: "Impossible de preparer l apercu de l accord.",
    nl: "Kon overeenkomstvoorbeeld niet voorbereiden.",
  },
  "Registration link required.": {
    fr: "Lien d’inscription requis.",
    nl: "Registratielink vereist.",
  },
  "This registration link is invalid or expired.": {
    fr: "Ce lien d’inscription est invalide ou expiré.",
    nl: "Deze registratielink is ongeldig of verlopen.",
  },
  "Invitation verified. Complete your details to activate access.": {
    fr: "Invitation vérifiée. Complétez vos informations pour activer l’accès.",
    nl: "Uitnodiging bevestigd. Vul je gegevens in om toegang te activeren.",
  },
  "Local signup preview mode. Sample data is loaded and no live account will be created.": {
    fr: "Mode aperçu local de l’inscription. Des données d’exemple sont chargées et aucun compte live ne sera créé.",
    nl: "Lokale previewmodus voor registratie. Voorbeelddata is geladen en er wordt geen live account aangemaakt.",
  },
  "Preview mode only. Registration is disabled on this local route.": {
    fr: "Mode aperçu uniquement. L’inscription est désactivée sur cette route locale.",
    nl: "Alleen previewmodus. Registratie is uitgeschakeld op deze lokale route.",
  },
  "Validating invitation...": { fr: "Validation de l’invitation...", nl: "Uitnodiging valideren..." },
  "Register Account": { fr: "Créer le compte", nl: "Account registreren" },
  "Creating account...": { fr: "Création du compte...", nl: "Account aanmaken..." },
  "Account created. Signing you in...": {
    fr: "Compte créé. Connexion en cours...",
    nl: "Account aangemaakt. Je wordt ingelogd...",
  },
  "Could not create account.": {
    fr: "Impossible de créer le compte.",
    nl: "Kon account niet aanmaken.",
  },
  "Registration is disabled.": { fr: "L’inscription est désactivée.", nl: "Registratie is uitgeschakeld." },
  "Company Name": { fr: "Nom de l’entreprise", nl: "Bedrijfsnaam" },
  "Point of Contact": { fr: "Personne de contact", nl: "Contactpersoon" },
  "Contact Email": { fr: "E-mail contact", nl: "E-mail contact" },
  "Phone": { fr: "Téléphone", nl: "Telefoon" },
  "Billing Address": { fr: "Adresse de facturation", nl: "Factuuradres" },
  "Tax ID": { fr: "Identifiant fiscal", nl: "Fiscaal nummer" },
  "Customer ID": { fr: "ID client", nl: "Klant-ID" },
  "All registration fields are required except Customer ID.": {
    fr: "Tous les champs d’inscription sont requis sauf l’ID client.",
    nl: "Alle registratievelden zijn verplicht behalve Klant-ID.",
  },
  "All registration fields are required.": {
    fr: "Tous les champs d’inscription sont requis.",
    nl: "Alle registratievelden zijn verplicht.",
  },
  "Password must be at least {min} characters.": {
    fr: "Le mot de passe doit comporter au moins {min} caractères.",
    nl: "Het wachtwoord moet minstens {min} tekens bevatten.",
  },
  "Passwords do not match.": {
    fr: "Les mots de passe ne correspondent pas.",
    nl: "Wachtwoorden komen niet overeen.",
  },
  "Invite email does not match this registration link.": {
    fr: "L’e-mail ne correspond pas à ce lien d’inscription.",
    nl: "Het e-mailadres komt niet overeen met deze registratielink.",
  },
  "Agreement acceptance is required.": {
    fr: "L acceptation de l accord est requise.",
    nl: "Acceptatie van de overeenkomst is vereist.",
  },
  "You must scroll through the full agreement before accepting.": {
    fr: "Vous devez faire defiler l accord complet avant d accepter.",
    nl: "Je moet door de volledige overeenkomst scrollen voordat je accepteert.",
  },
  "Agreement version mismatch. Refresh and try again.": {
    fr: "Version de l accord invalide. Actualisez et reessayez.",
    nl: "Overeenkomstversie komt niet overeen. Vernieuw en probeer opnieuw.",
  },
  "Agreement integrity check failed. Refresh and try again.": {
    fr: "Controle d integrite de l accord echouee. Actualisez et reessayez.",
    nl: "Integriteitscontrole van de overeenkomst mislukt. Vernieuw en probeer opnieuw.",
  },
  "Agreement identifier mismatch. Refresh and try again.": {
    fr: "Identifiant de l accord invalide. Actualisez et reessayez.",
    nl: "Overeenkomst-id komt niet overeen. Vernieuw en probeer opnieuw.",
  },
  "Agreement timestamps are missing.": {
    fr: "Horodatages de l accord manquants.",
    nl: "Tijdstempels van de overeenkomst ontbreken.",
  },
  "Agreement timestamps are invalid.": {
    fr: "Horodatages de l accord invalides.",
    nl: "Tijdstempels van de overeenkomst zijn ongeldig.",
  },
  "Agreement scroll checkpoint is invalid.": {
    fr: "Point de controle de defilement invalide.",
    nl: "Scrollcontrolepunt van overeenkomst is ongeldig.",
  },
  "Agreement confirmation expired. Please review and accept again.": {
    fr: "Confirmation d accord expiree. Veuillez relire et accepter a nouveau.",
    nl: "Overeenkomstbevestiging is verlopen. Lees en accepteer opnieuw.",
  },
  "Agreement timestamp is in the future.": {
    fr: "Horodatage de l accord dans le futur.",
    nl: "Tijdstempel van overeenkomst ligt in de toekomst.",
  },
  "Agreement scroll requirement was not completed.": {
    fr: "Condition de defilement de l accord non respectee.",
    nl: "Vereiste scroll voor overeenkomst is niet voltooid.",
  },
  "Click-wrap schema missing. Run supabase_clickwrap.sql in Supabase SQL editor.": {
    fr: "Schema click-wrap manquant. Executez supabase_clickwrap.sql dans l editeur SQL Supabase.",
    nl: "Click-wrap-schema ontbreekt. Voer supabase_clickwrap.sql uit in de Supabase SQL-editor.",
  },
  "Creating invite...": { fr: "Création de l’invitation...", nl: "Uitnodiging maken..." },
  "Sign In": { fr: "Se connecter", nl: "Inloggen" },
  "Create Account": { fr: "Créer un compte", nl: "Account aanmaken" },
  "Forgot Password?": { fr: "Mot de passe oublié ?", nl: "Wachtwoord vergeten?" },
  "Account": { fr: "Compte", nl: "Account" },
  "History": { fr: "Historique", nl: "Historiek" },
  "Reports": { fr: "Rapports", nl: "Rapporten" },
  "Admin Panel": { fr: "Panneau admin", nl: "Adminpaneel" },
  "Language": { fr: "Langue", nl: "Taal" },
  "Fill Mock Data": { fr: "Charger des données fictives", nl: "Mockdata laden" },
  "Restore Live Data": { fr: "Restaurer les données live", nl: "Live data herstellen" },
  "Mock data loaded (frontend only).": {
    fr: "Données fictives chargées (frontend uniquement).",
    nl: "Mockdata geladen (alleen frontend).",
  },
  "Live admin data restored.": {
    fr: "Données admin live restaurées.",
    nl: "Live admindata hersteld.",
  },
  "Payment Methods": { fr: "Modes de paiement", nl: "Betaalmethodes" },
  "Invoice": { fr: "Facture", nl: "Factuur" },
  "Card": { fr: "Carte", nl: "Kaart" },
  "Instant": { fr: "Instant", nl: "Instant" },
  "Payment Tracking": { fr: "Suivi du paiement", nl: "Betalingsopvolging" },
  "Invoice Tracking": { fr: "Suivi facture", nl: "Factuuropvolging" },
  "Invoice billing": { fr: "Facturation sur facture", nl: "Facturatie via factuur" },
  "Card auto-charge": { fr: "Prélèvement carte auto", nl: "Automatische kaartbetaling" },
  "Hybrid billing": { fr: "Facturation hybride", nl: "Hybride facturatie" },
  "No billable activity": { fr: "Aucune activité facturable", nl: "Geen factureerbare activiteit" },
  "Invoice sent · pending": { fr: "Facture envoyée · en attente", nl: "Factuur verzonden · in afwachting" },
  "Reminder sent (D-1)": { fr: "Rappel envoyé (J-1)", nl: "Herinnering verzonden (D-1)" },
  "Client billing updated.": {
    fr: "Configuration de paiement client mise à jour.",
    nl: "Klantbetalingsinstelling bijgewerkt.",
  },
  "At least one payment method must remain enabled.": {
    fr: "Au moins un mode de paiement doit rester activé.",
    nl: "Minstens één betaalmethode moet ingeschakeld blijven.",
  },
  "Could not update client billing settings.": {
    fr: "Impossible de mettre à jour la configuration de paiement client.",
    nl: "Kon klantbetalingsinstellingen niet bijwerken.",
  },
  "Build, confirm, and preview your postage label in minutes.": {
    fr: "Créez, confirmez et prévisualisez votre étiquette d’expédition en quelques minutes.",
    nl: "Maak, bevestig en bekijk je verzendlabel in enkele minuten.",
  },
  "A minimal, high-trust label pipeline: enter details, choose a service, confirm invoice details, and receive a crisp PDF ready to print.": {
    fr: "Un flux minimal et fiable : saisissez les informations, choisissez un service, confirmez la facture et obtenez un PDF prêt à imprimer.",
    nl: "Een minimale, betrouwbare flow: voer gegevens in, kies een service, bevestig factuurgegevens en ontvang een scherpe printklare PDF.",
  },
  "Label Info": { fr: "Infos Étiquette", nl: "Labelinfo" },
  "Shipment Details": { fr: "Détails d’expédition", nl: "Zendingsdetails" },
  "Sender, recipient, parcel": { fr: "Expéditeur, destinataire, colis", nl: "Afzender, ontvanger, pakket" },
  "Label Selection": { fr: "Choix Étiquette", nl: "Labelkeuze" },
  "Service Selection": { fr: "Choix du service", nl: "Servicekeuze" },
  "Carrier & Service": { fr: "Transporteur & Service", nl: "Vervoerder & Service" },
  "Service, speed, cost": { fr: "Service, délai, coût", nl: "Service, snelheid, kost" },
  "Carrier, speed, price": { fr: "Transporteur, délai, prix", nl: "Vervoerder, snelheid, prijs" },
  "Delivery option and pricing": {
    fr: "Option de livraison et tarification",
    nl: "Leveringsoptie en prijs",
  },
  "Invoice Confirm": { fr: "Confirmation Facture", nl: "Factuurbevestiging" },
  "Invoice Review": { fr: "Vérification facture", nl: "Factuurcontrole" },
  "Billing details, monthly invoice": { fr: "Coordonnées de facturation, facture mensuelle", nl: "Factuurgegevens, maandfactuur" },
  "Billing profile, monthly invoice": { fr: "Profil de facturation, facture mensuelle", nl: "Factuurprofiel, maandfactuur" },
  "Billing profile and approval": {
    fr: "Profil de facturation et validation",
    nl: "Factuurprofiel en bevestiging",
  },
  "Payment": { fr: "Paiement", nl: "Betaling" },
  "Wallet or card payment": {
    fr: "Paiement par solde ou carte",
    nl: "Betaling via saldo of kaart",
  },
  "Label PDF": { fr: "PDF Étiquette", nl: "Label-PDF" },
  "Label Output": { fr: "Sortie étiquettes", nl: "Labeloutput" },
  "Labels & PDF": { fr: "Étiquettes & PDF", nl: "Labels & PDF" },
  "Preview and download": { fr: "Prévisualiser et télécharger", nl: "Bekijken en downloaden" },
  "Order Summary": { fr: "Résumé de commande", nl: "Besteloverzicht" },
  "Your Account": { fr: "Votre compte", nl: "Jouw account" },
  "Service": { fr: "Service", nl: "Service" },
  "Carrier": { fr: "Transporteur", nl: "Vervoerder" },
  "Effective Discount": { fr: "Remise effective", nl: "Effectieve korting" },
  "Next Invoice Amount": { fr: "Montant prochaine facture", nl: "Volgend factuurbedrag" },
  "Next Invoice Date": { fr: "Date prochaine facture", nl: "Volgende factuurdatum" },
  "Need Help?": { fr: "Besoin d’aide ?", nl: "Hulp nodig?" },
  "Chat with us": { fr: "Chatter avec nous", nl: "Chat met ons" },
  "Economy Carrier": { fr: "Transporteur Économie", nl: "Economy-vervoerder" },
  "Priority Carrier": { fr: "Transporteur Prioritaire", nl: "Prioriteit-vervoerder" },
  "International Carrier": { fr: "Transporteur International", nl: "Internationale vervoerder" },
  "Carrier Network": { fr: "Réseau transporteur", nl: "Vervoerdersnetwerk" },
  "Unit Price": { fr: "Prix unitaire", nl: "Eenheidsprijs" },
  "Quantity": { fr: "Quantité", nl: "Aantal" },
  "Total": { fr: "Total", nl: "Totaal" },
  "Tracking": { fr: "Suivi", nl: "Tracking" },
  "Label ID": { fr: "ID étiquette", nl: "Label-ID" },
  "Choose your label service": { fr: "Choisissez votre service d’étiquette", nl: "Kies je labelservice" },
  "Select your delivery service. Pricing updates instantly.": { fr: "Sélectionnez votre service de livraison. Le prix se met à jour instantanément.", nl: "Selecteer je leveringsservice. De prijs wordt meteen bijgewerkt." },
  "Economy": { fr: "Économie", nl: "Economy" },
  "Reliable, cost-efficient": { fr: "Fiable, économique", nl: "Betrouwbaar, kostenefficiënt" },
  "TRACKING INCLUDED": { fr: "SUIVI INCLUS", nl: "TRACKING INBEGREPEN" },
  "4x6 PDF": { fr: "PDF 4x6", nl: "4x6 PDF" },
  "Priority": { fr: "Prioritaire", nl: "Prioriteit" },
  "Fast, time-sensitive": { fr: "Rapide, urgent", nl: "Snel, tijdsgevoelig" },
  "SIGNATURE READY": { fr: "SIGNATURE PRÊTE", nl: "HANDTEKENING KLAAR" },
  "INSURED": { fr: "ASSURÉ", nl: "VERZEKERD" },
  "International Express": { fr: "Express International", nl: "Internationaal Express" },
  "Cross-border, priority routing": { fr: "Transfrontalier, routage prioritaire", nl: "Grensoverschrijdend, prioritaire routing" },
  "CUSTOMS READY": { fr: "DOUANE PRÊTE", nl: "DOUANE KLAAR" },
  "GLOBAL ETA": { fr: "DÉLAI MONDIAL", nl: "GLOBALE ETA" },
  "Label Quantity": { fr: "Quantité d’étiquettes", nl: "Aantal labels" },
  "Back": { fr: "Retour", nl: "Terug" },
  "Continue": { fr: "Continuer", nl: "Doorgaan" },
  "Enter label information": { fr: "Saisir les informations de l’étiquette", nl: "Labelinformatie invoeren" },
  "Add sender, recipient, and parcel details — or import from a provider.": {
    fr: "Ajoutez les informations expéditeur, destinataire et colis — ou importez depuis un fournisseur.",
    nl: "Voeg afzender-, ontvanger- en pakketgegevens toe — of importeer via een provider.",
  },
  "Import from provider": { fr: "Importer depuis un fournisseur", nl: "Importeren vanuit provider" },
  "Connected": { fr: "Connecté", nl: "Verbonden" },
  "Amazon Seller": { fr: "Vendeur Amazon", nl: "Amazon Seller" },
  "Upload CSV": { fr: "Importer CSV", nl: "CSV uploaden" },
  "Auto-fill": { fr: "Auto-remplir", nl: "Auto-invullen" },
  "Auto CSV": { fr: "Auto CSV", nl: "Auto CSV" },
  "Complete all required fields or use Auto-generate CSV.": { fr: "Complétez tous les champs requis ou utilisez la génération CSV automatique.", nl: "Vul alle verplichte velden in of gebruik Auto-generate CSV." },
  "Sender": { fr: "Expéditeur", nl: "Afzender" },
  "Recipient": { fr: "Destinataire", nl: "Ontvanger" },
  "Name": { fr: "Nom", nl: "Naam" },
  "Street": { fr: "Rue", nl: "Straat" },
  "Country": { fr: "Pays", nl: "Land" },
  "City": { fr: "Ville", nl: "Stad" },
  "Region / Province": { fr: "Région / Province", nl: "Regio / Provincie" },
  "Postal Code": { fr: "Code postal", nl: "Postcode" },
  "Parcel": { fr: "Colis", nl: "Pakket" },
  "Weight (kg)": { fr: "Poids (kg)", nl: "Gewicht (kg)" },
  "Weight": { fr: "Poids", nl: "Gewicht" },
  "Dimensions (cm)": { fr: "Dimensions (cm)", nl: "Afmetingen (cm)" },
  "Dims (L x W x H, cm)": { fr: "Dims (L x l x H, cm)", nl: "Afm. (L x B x H, cm)" },
  "CSV Batch Upload": { fr: "Import CSV en lot", nl: "CSV-batch upload" },
  "Imported rows ready for review.": { fr: "Lignes importées prêtes à être vérifiées.", nl: "Geïmporteerde rijen klaar voor controle." },
  "Edit": { fr: "Modifier", nl: "Bewerken" },
  "Editing": { fr: "Édition", nl: "Bewerken" },
  "Ship From": { fr: "Expédier depuis", nl: "Verzenden vanuit" },
  "Customs Declaration": { fr: "Déclaration en douane", nl: "Douaneaangifte" },
  "Complete this declaration for destinations that require customs processing.": { fr: "Complétez cette déclaration pour les destinations nécessitant un traitement douanier.", nl: "Vul deze aangifte in voor bestemmingen die douaneafhandeling vereisen." },
  "Complete required customs declaration fields before continuing.": { fr: "Complétez les champs obligatoires de la déclaration douanière avant de continuer.", nl: "Vul de verplichte douanevelden in voordat je verdergaat." },
  "Parcel Summary": { fr: "Résumé du colis", nl: "Pakketoverzicht" },
  "Content Description": { fr: "Description du contenu", nl: "Inhoudsbeschrijving" },
  "Shipment Type": { fr: "Type d’expédition", nl: "Type verzending" },
  "Select type": { fr: "Sélectionner un type", nl: "Type selecteren" },
  "Sale of goods": { fr: "Vente de marchandises", nl: "Verkoop van goederen" },
  "Gift": { fr: "Cadeau", nl: "Geschenk" },
  "Commercial sample": { fr: "Échantillon commercial", nl: "Commercieel staal" },
  "Returned goods": { fr: "Marchandises retournées", nl: "Retourgoederen" },
  "Documents": { fr: "Documents", nl: "Documenten" },
  "Other": { fr: "Autre", nl: "Andere" },
  "Sender Customs Reference (optional)": { fr: "Référence douanière expéditeur (optionnel)", nl: "Douanereferentie afzender (optioneel)" },
  "Importer / Recipient Reference (optional)": { fr: "Référence importateur / destinataire (optionnel)", nl: "Referentie importeur / ontvanger (optioneel)" },
  "Verify that the shipment content is allowed.": { fr: "Vérifiez que le contenu de l’envoi est autorisé.", nl: "Controleer of de inhoud van de zending is toegestaan." },
  "Review restricted goods list": { fr: "Consulter la liste des marchandises interdites", nl: "Bekijk lijst met beperkte goederen" },
  "Declared Items": { fr: "Articles déclarés", nl: "Aangegeven items" },
  "Add Item": { fr: "Ajouter un article", nl: "Item toevoegen" },
  "Top up by IBAN": { fr: "Recharger par IBAN", nl: "Opwaarderen via IBAN" },
  "Generate Transfer Reference": {
    fr: "Générer la référence de virement",
    nl: "Overschrijvingsreferentie genereren",
  },
  "Generate New Reference": {
    fr: "Générer une nouvelle référence",
    nl: "Nieuwe referentie genereren",
  },
  "Fund your account balance": { fr: "Alimentez votre solde de compte", nl: "Vul je accountsaldo aan" },
  "Generate a transfer reference, then send your bank transfer. Funds are credited once received.": {
    fr: "Générez une référence de virement puis envoyez votre transfert bancaire. Les fonds sont crédités à réception.",
    nl: "Genereer een overschrijvingsreferentie en verstuur je bankoverschrijving. Fondsen worden gecrediteerd na ontvangst.",
  },
  "Beneficiary": { fr: "Bénéficiaire", nl: "Begunstigde" },
  "Reference": { fr: "Référence", nl: "Referentie" },
  "ETA": { fr: "Délai", nl: "ETA" },
  "Copy Reference": { fr: "Copier la référence", nl: "Referentie kopiëren" },
  "Copy IBAN": { fr: "Copier l’IBAN", nl: "IBAN kopiëren" },
  "Close": { fr: "Fermer", nl: "Sluiten" },
  "Confirm invoice information": { fr: "Confirmer les informations de facturation", nl: "Factuurinformatie bevestigen" },
  "No charge now. These details are used for your month-end invoice.": { fr: "Aucun débit maintenant. Ces informations seront utilisées pour votre facture de fin de mois.", nl: "Nu geen kosten. Deze gegevens worden gebruikt voor je maandfactuur." },
  "Invoice Recipient": { fr: "Destinataire de la facture", nl: "Factuurontvanger" },
  "Company": { fr: "Entreprise", nl: "Bedrijf" },
  "Point of Contact": { fr: "Personne de contact", nl: "Contactpersoon" },
  "Phone": { fr: "Téléphone", nl: "Telefoon" },
  "Phone Number": { fr: "Numéro de téléphone", nl: "Telefoonnummer" },
  "Billing Address": { fr: "Adresse de facturation", nl: "Factuuradres" },
  "Tax ID": { fr: "Identifiant fiscal", nl: "Fiscaal nummer" },
  "Account Manager": { fr: "Account manager", nl: "Accountmanager" },
  "Monthly Invoice Preview": { fr: "Aperçu de facture mensuelle", nl: "Voorbeeld maandfactuur" },
  "Label service": { fr: "Service d’étiquette", nl: "Labelservice" },
  "Subtotal": { fr: "Sous-total", nl: "Subtotaal" },
  "Subtotal (ex VAT)": { fr: "Sous-total", nl: "Subtotaal" },
  "Subtotal (ex. vat)": { fr: "Sous-total", nl: "Subtotaal" },
  "Subtotal (EX. VAT)": { fr: "Sous-total", nl: "Subtotaal" },
  "VAT (21%)": { fr: "Taxe", nl: "Belasting" },
  "Total": { fr: "Total", nl: "Totaal" },
  "Total (incl VAT)": { fr: "Total", nl: "Totaal" },
  "Total (incl. vat)": { fr: "Total", nl: "Totaal" },
  "Total (INCL. VAT)": { fr: "Total", nl: "Totaal" },
  "This order will be added to your monthly invoice. If you want your information to be changed, please contact your account manager.": {
    fr: "Cette commande sera ajoutée à votre facture mensuelle. Si vous souhaitez modifier vos informations, contactez votre account manager.",
    nl: "Deze bestelling wordt toegevoegd aan je maandfactuur. Als je gegevens moeten wijzigen, neem contact op met je accountmanager.",
  },
  "Choose payment method": { fr: "Choisissez un mode de paiement", nl: "Kies betaalmethode" },
  "Order summary": { fr: "Récapitulatif", nl: "Besteloverzicht" },
  "Top up": { fr: "Recharger", nl: "Opwaarderen" },
  "IBAN Transfer": { fr: "Virement IBAN", nl: "IBAN-overschrijving" },
  "Reference History": { fr: "Historique des références", nl: "Referentiegeschiedenis" },
  "No transfer references yet.": {
    fr: "Aucune référence de virement pour l’instant.",
    nl: "Nog geen overschrijvingsreferenties.",
  },
  "Pending": { fr: "En attente", nl: "In afwachting" },
  "Received": { fr: "Reçu", nl: "Ontvangen" },
  "Credited": { fr: "Crédité", nl: "Gecrediteerd" },
  "Cancelled": { fr: "Annulé", nl: "Geannuleerd" },
  "Failed": { fr: "Échec", nl: "Mislukt" },
  "Card and wallet are instant. Bank transfer top-up is credited after receipt.": {
    fr: "La carte et le solde sont instantanés. Le virement est crédité après réception.",
    nl: "Kaart en saldo zijn direct. Overschrijving wordt na ontvangst gecrediteerd.",
  },
  "Select an instant payment method to generate your labels now.": {
    fr: "Sélectionnez un mode de paiement instantané pour générer vos étiquettes maintenant.",
    nl: "Kies een directe betaalmethode om je labels nu te genereren.",
  },
  "Pay & Generate": { fr: "Payer & Générer", nl: "Betalen & Genereren" },
  "No open invoices.": { fr: "Aucune facture ouverte.", nl: "Geen open facturen." },
  "No paid invoices yet.": { fr: "Aucune facture payée pour l’instant.", nl: "Nog geen betaalde facturen." },
  "Current Balance": { fr: "Solde actuel", nl: "Huidig saldo" },
  "Order Total": { fr: "Total commande", nl: "Besteltotaal" },
  "Order Total (INCL. VAT)": { fr: "Total commande", nl: "Besteltotaal" },
  "Bank transfer top-ups are asynchronous. Add funds now and use your balance once credited.": {
    fr: "Les recharges par virement sont asynchrones. Ajoutez des fonds maintenant et utilisez votre solde une fois crédité.",
    nl: "Opwaarderingen via bankoverschrijving zijn asynchroon. Voeg nu saldo toe en gebruik het zodra het is gecrediteerd.",
  },
  "Account balance": { fr: "Solde du compte", nl: "Accountsaldo" },
  "Pay instantly from your available balance.": {
    fr: "Payez instantanément depuis votre solde disponible.",
    nl: "Betaal direct vanuit je beschikbare saldo.",
  },
  "Credit / Debit card": { fr: "Carte de crédit / débit", nl: "Creditcard / debetkaart" },
  "Charge immediately and generate labels now.": {
    fr: "Débit immédiat et génération des étiquettes.",
    nl: "Direct afrekenen en labels meteen genereren.",
  },
  "Status": { fr: "Statut", nl: "Status" },
  "Available": { fr: "Disponible", nl: "Beschikbaar" },
  "Disabled": { fr: "Désactivé", nl: "Uitgeschakeld" },
  "IBAN": { fr: "IBAN", nl: "IBAN" },
  "Reference required": { fr: "Référence obligatoire", nl: "Referentie verplicht" },
  "1-2 business days": { fr: "1-2 jours ouvrés", nl: "1-2 werkdagen" },
  "Transfers are credited once received (typically 1-2 business days).": {
    fr: "Les virements sont crédités dès réception (généralement 1-2 jours ouvrés).",
    nl: "Overschrijvingen worden gecrediteerd na ontvangst (meestal 1-2 werkdagen).",
  },
  "Enter a valid top-up amount.": {
    fr: "Entrez un montant de recharge valide.",
    nl: "Voer een geldig opwaardeerbedrag in.",
  },
  "Creating top-up request...": {
    fr: "Création de la demande de recharge...",
    nl: "Opwaardeeraanvraag wordt aangemaakt...",
  },
  "Preparing transfer instructions...": {
    fr: "Préparation des informations de virement...",
    nl: "Overschrijvingsgegevens worden voorbereid...",
  },
  "Transfer reference generated. Use it as communication.": {
    fr: "Référence de virement générée. Utilisez-la en communication.",
    nl: "Overschrijvingsreferentie gegenereerd. Gebruik die als mededeling.",
  },
  "IBAN top-up request created.": {
    fr: "Demande de recharge IBAN créée.",
    nl: "IBAN-opwaardeeraanvraag aangemaakt.",
  },
  "Could not create top-up request.": {
    fr: "Impossible de créer la demande de recharge.",
    nl: "Kon opwaardeeraanvraag niet maken.",
  },
  "Reference copied to clipboard.": {
    fr: "Référence copiée dans le presse-papiers.",
    nl: "Referentie naar klembord gekopieerd.",
  },
  "Could not copy reference.": {
    fr: "Impossible de copier la référence.",
    nl: "Kon referentie niet kopiëren.",
  },
  "IBAN copied to clipboard.": {
    fr: "IBAN copié dans le presse-papiers.",
    nl: "IBAN naar klembord gekopieerd.",
  },
  "Could not copy IBAN.": {
    fr: "Impossible de copier l’IBAN.",
    nl: "Kon IBAN niet kopiëren.",
  },
  "Select a payment method to continue.": {
    fr: "Sélectionnez un mode de paiement pour continuer.",
    nl: "Kies een betaalmethode om verder te gaan.",
  },
  "Insufficient wallet balance. Top up by IBAN or use card.": {
    fr: "Solde insuffisant. Rechargez via IBAN ou utilisez la carte.",
    nl: "Onvoldoende saldo. Waardeer op via IBAN of gebruik kaart.",
  },
  "Could not process checkout.": {
    fr: "Impossible de traiter le paiement.",
    nl: "Kon betaling niet verwerken.",
  },
  "Paid with account balance. Remaining balance: {amount}": {
    fr: "Payé via le solde du compte. Solde restant : {amount}",
    nl: "Betaald via accountsaldo. Resterend saldo: {amount}",
  },
  "Card payment approved. Generating labels.": {
    fr: "Paiement carte approuvé. Génération des étiquettes.",
    nl: "Kaartbetaling goedgekeurd. Labels worden gegenereerd.",
  },
  "You must be signed in.": {
    fr: "Vous devez être connecté.",
    nl: "Je moet ingelogd zijn.",
  },
  "Confirm & Generate": { fr: "Confirmer & Générer", nl: "Bevestigen & Genereren" },
  "Your label is ready": { fr: "Votre étiquette est prête", nl: "Je label is klaar" },
  "Preview the PDF and download the label for printing.": { fr: "Prévisualisez le PDF et téléchargez l’étiquette à imprimer.", nl: "Bekijk de PDF en download het label om te printen." },
  "Download PDF": { fr: "Télécharger PDF", nl: "PDF downloaden" },
  "Start New Label": { fr: "Nouvelle Étiquette", nl: "Nieuw label starten" },
  "PDF Preview": { fr: "Aperçu PDF", nl: "PDF-voorbeeld" },
  "Batch Queue": { fr: "File d’attente lot", nl: "Batchwachtrij" },
  "Label Summary": { fr: "Résumé d’étiquette", nl: "Labelsamenvatting" },
  "TRACKING": { fr: "SUIVI", nl: "TRACKING" },
  "FROM": { fr: "DE", nl: "VAN" },
  "TO": { fr: "VERS", nl: "NAAR" },
  "WEIGHT": { fr: "POIDS", nl: "GEWICHT" },
  "DIMENSIONS": { fr: "DIMENSIONS", nl: "AFMETINGEN" },
  "Account Overview": { fr: "Vue du compte", nl: "Accountoverzicht" },
  "Review account profile details and manage your shipping-origin presets.": { fr: "Consultez les informations du compte et gérez vos adresses d’expédition.", nl: "Bekijk accountgegevens en beheer je verzendorigines." },
  "Back to Builder": { fr: "Retour au builder", nl: "Terug naar builder" },
  "Sign Out": { fr: "Se déconnecter", nl: "Uitloggen" },
  "Account Information": { fr: "Informations du compte", nl: "Accountinformatie" },
  "Billing & Balance": { fr: "Facturation & Solde", nl: "Facturatie & Saldo" },
  "Track wallet funding, invoice status, and payment readiness.": {
    fr: "Suivez le financement du solde, le statut des factures et la disponibilité au paiement.",
    nl: "Volg saldo-opwaarderingen, factuurstatus en betaalgereedheid.",
  },
  "Available Balance": { fr: "Solde disponible", nl: "Beschikbaar saldo" },
  "Pending Top-ups": { fr: "Recharges en attente", nl: "Open opwaarderingen" },
  "Open Invoices": { fr: "Factures ouvertes", nl: "Open facturen" },
  "Paid Invoices": { fr: "Factures payées", nl: "Betaalde facturen" },
  "Open invoices": { fr: "Factures ouvertes", nl: "Open facturen" },
  "Paid invoices": { fr: "Factures payées", nl: "Betaalde facturen" },
  "IBAN transfer details available in Top up by IBAN.": {
    fr: "Détails du virement IBAN disponibles dans Recharge par IBAN.",
    nl: "IBAN-overschrijvingsgegevens beschikbaar in Opwaarderen via IBAN.",
  },
  "Company Name": { fr: "Nom de l’entreprise", nl: "Bedrijfsnaam" },
  "Contact Email": { fr: "E-mail contact", nl: "E-mail contact" },
  "Customer ID": { fr: "ID client", nl: "Klant-ID" },
  "Plan": { fr: "Plan", nl: "Plan" },
  "To update account information, please contact your account manager.": { fr: "Pour modifier les informations du compte, veuillez contacter votre account manager.", nl: "Neem contact op met je accountmanager om accountinformatie te wijzigen." },
  "Shipping Origins": { fr: "Origines d’expédition", nl: "Verzendorigines" },
  "Add Warehouse": { fr: "Ajouter un entrepôt", nl: "Magazijn toevoegen" },
  "Save Origins": { fr: "Enregistrer les origines", nl: "Origines opslaan" },
  "Save": { fr: "Enregistrer", nl: "Opslaan" },
  "Sign in to manage shipping origins.": { fr: "Connectez-vous pour gérer les origines d’expédition.", nl: "Log in om verzendorigines te beheren." },
  "Client Invitations": { fr: "Invitations client", nl: "Klantuitnodigingen" },
  "Commercial Settings": { fr: "Paramètres commerciaux", nl: "Commerciële instellingen" },
  "Commercial Discount Settings": {
    fr: "Paramètres de remise commerciale",
    nl: "Commerciële kortingsinstellingen",
  },
  "Create Client": { fr: "Créer client", nl: "Klant aanmaken" },
  "Client Email": { fr: "E-mail client", nl: "Klant-e-mail" },
  "Client Email (optional)": { fr: "E-mail client (optionnel)", nl: "Klant-e-mail (optioneel)" },
  "Link Expiry": { fr: "Expiration du lien", nl: "Linkverval" },
  "7 days": { fr: "7 jours", nl: "7 dagen" },
  "14 days": { fr: "14 jours", nl: "14 dagen" },
  "30 days": { fr: "30 jours", nl: "30 dagen" },
  "60 days": { fr: "60 jours", nl: "60 dagen" },
  "Registration URL": { fr: "URL d’inscription", nl: "Registratie-URL" },
  "Copy URL": { fr: "Copier l’URL", nl: "URL kopiëren" },
  "Invite link created.": { fr: "Lien d’invitation créé.", nl: "Uitnodigingslink aangemaakt." },
  "Invite URL copied.": { fr: "URL d’invitation copiée.", nl: "Uitnodigings-URL gekopieerd." },
  "Invalid client email format.": { fr: "Format d’e-mail client invalide.", nl: "Ongeldig formaat voor klant-e-mail." },
  "You are not allowed to create client invites.": { fr: "Vous n’êtes pas autorisé à créer des invitations client.", nl: "Je bent niet gemachtigd om klantuitnodigingen te maken." },
  "Recent links": { fr: "Liens récents", nl: "Recente links" },
  "Recent Links": { fr: "Liens récents", nl: "Recente links" },
  "Create controlled registration links and share them instantly.": {
    fr: "Créez des liens d’inscription contrôlés et partagez-les instantanément.",
    nl: "Maak gecontroleerde registratielinks en deel ze meteen.",
  },
  "Define the carrier discount you receive and the share you pass through to clients.": {
    fr: "Définissez la remise transporteur reçue et la part répercutée aux clients.",
    nl: "Bepaal de vervoerderskorting die je ontvangt en welk deel je aan klanten doorgeeft.",
  },
  "No invites yet.": { fr: "Aucune invitation pour l’instant.", nl: "Nog geen uitnodigingen." },
  "Loading invites...": { fr: "Chargement des invitations...", nl: "Uitnodigingen laden..." },
  "No email lock": { fr: "Sans e-mail verrouillé", nl: "Zonder e-mailvergrendeling" },
  "Created {date}": { fr: "Créé {date}", nl: "Aangemaakt {date}" },
  "Expires {date}": { fr: "Vervalt {date}", nl: "Verloopt {date}" },
  "Claimed {date}": { fr: "Réclamé {date}", nl: "Geclaimd {date}" },
  "Open": { fr: "Ouvert", nl: "Open" },
  "Claimed": { fr: "Réclamé", nl: "Geclaimd" },
  "Expired": { fr: "Expiré", nl: "Verlopen" },
  "Generation History": { fr: "Historique des générations", nl: "Generatiehistoriek" },
  "Review previous generations, preview label PDFs, and export receipts.": { fr: "Consultez les générations précédentes, prévisualisez les PDF et exportez les reçus.", nl: "Bekijk eerdere generaties, label-PDF’s en exporteer bonnen." },
  "Generations": { fr: "Générations", nl: "Generaties" },
  "Sign in to view previous generations.": { fr: "Connectez-vous pour voir les générations précédentes.", nl: "Log in om vorige generaties te bekijken." },
  "Generation Preview": { fr: "Aperçu de génération", nl: "Generatievoorbeeld" },
  "Select a generation to preview labels and receipt details.": { fr: "Sélectionnez une génération pour voir les étiquettes et le reçu.", nl: "Selecteer een generatie om labels en bondetails te bekijken." },
  "Select a generation to view PDFs and receipt details.": { fr: "Sélectionnez une génération pour afficher les PDF et les détails du reçu.", nl: "Selecteer een generatie om PDF's en bondetails te bekijken." },
  "Download Label PDF": { fr: "Télécharger PDF étiquette", nl: "Label-PDF downloaden" },
  "Download Labels (.PDF)": { fr: "Télécharger les étiquettes (.PDF)", nl: "Labels downloaden (.PDF)" },
  "View Receipt": { fr: "Voir le reçu", nl: "Bon bekijken" },
  "Download Invoice": { fr: "Télécharger la facture", nl: "Factuur downloaden" },
  "WooCommerce": { fr: "WooCommerce", nl: "WooCommerce" },
  "Wix eCommerce": { fr: "Wix eCommerce", nl: "Wix eCommerce" },
  "Coming soon": { fr: "Bientôt", nl: "Binnenkort" },
  "{provider} Settings": { fr: "Paramètres {provider}", nl: "{provider}-instellingen" },
  "{provider} import and settings are coming soon.": {
    fr: "L’import et les paramètres {provider} arrivent bientôt.",
    nl: "{provider}-import en instellingen komen binnenkort.",
  },
  "No label PDF selected yet. Choose a generation to preview.": { fr: "Aucun PDF d’étiquette sélectionné. Choisissez une génération pour prévisualiser.", nl: "Nog geen label-PDF geselecteerd. Kies een generatie om te bekijken." },
  "Shipping Reports": { fr: "Rapports d’expédition", nl: "Verzendrapporten" },
  "Interactive visibility into savings, payments, shipment volume, services, and destinations.": { fr: "Vue interactive des économies, paiements, volumes, services et destinations.", nl: "Interactief inzicht in besparingen, betalingen, volume, services en bestemmingen." },
  "All Time": { fr: "Toute la période", nl: "Alle tijd" },
  "Last 7 Days": { fr: "7 derniers jours", nl: "Laatste 7 dagen" },
  "Last 30 Days": { fr: "30 derniers jours", nl: "Laatste 30 dagen" },
  "Last 90 Days": { fr: "90 derniers jours", nl: "Laatste 90 dagen" },
  "Last Year": { fr: "Dernière année", nl: "Afgelopen jaar" },
  "Custom": { fr: "Personnalisé", nl: "Aangepast" },
  "Custom Range": { fr: "Période personnalisée", nl: "Aangepast bereik" },
  "Time frame": { fr: "Période", nl: "Tijdsperiode" },
  "From": { fr: "Du", nl: "Van" },
  "To": { fr: "Au", nl: "Tot" },
  "Apply Range": { fr: "Appliquer", nl: "Bereik toepassen" },
  "Savings vs Retail": { fr: "Économies vs public", nl: "Besparing vs retail" },
  "Estimated discount captured": { fr: "Remise estimée capturée", nl: "Geschatte gerealiseerde korting" },
  "Pending Returns": { fr: "Retours en attente", nl: "Openstaande retouren" },
  "Return labels waiting action": { fr: "Étiquettes retour en attente", nl: "Retourlabels wachten op actie" },
  "Pending Label Refunds": { fr: "Remboursements en attente", nl: "Openstaande labelterugbetalingen" },
  "Unused or voided labels": { fr: "Étiquettes non utilisées ou annulées", nl: "Ongebruikte of geannuleerde labels" },
  "Account Balance": { fr: "Solde du compte", nl: "Rekeningsaldo" },
  "Monthly billing cycle": { fr: "Cycle de facturation mensuel", nl: "Maandelijkse facturatiecyclus" },
  "Settled at month end": { fr: "Soldé en fin de mois", nl: "Verrekend op einde maand" },
  "Total Shipping Costs": { fr: "Coûts totaux d’expédition", nl: "Totale verzendkosten" },
  "EX VAT": { fr: "Total", nl: "Totaal" },
  "ex. vat": { fr: "total", nl: "totaal" },
  "incl. vat": { fr: "total", nl: "totaal" },
  "EX. VAT": { fr: "Total", nl: "Totaal" },
  "INCL. VAT": { fr: "Total", nl: "Totaal" },
  "Label history": { fr: "Historique des étiquettes", nl: "Labelhistoriek" },
  "Selected range": { fr: "Période sélectionnée", nl: "Geselecteerde periode" },
  "Service charges in selected range": { fr: "Frais de service sur la période", nl: "Servicekosten in geselecteerde periode" },
  "Carrier Adjustments": { fr: "Ajustements transporteur", nl: "Carrier-correcties" },
  "Post-label carrier corrections": { fr: "Corrections transporteur post-étiquette", nl: "Carrier-correcties na label" },
  "Total Payments": { fr: "Paiements totaux", nl: "Totale betalingen" },
  "Daily spend trend": { fr: "Tendance des dépenses quotidiennes", nl: "Dagelijkse uitgaventrend" },
  "Daily spend trend (ex VAT)": { fr: "Tendance des dépenses quotidiennes", nl: "Dagelijkse uitgaventrend" },
  "Daily spend trend (ex. vat)": {
    fr: "Tendance des dépenses quotidiennes",
    nl: "Dagelijkse uitgaventrend",
  },
  "Daily spend trend (EX. VAT)": {
    fr: "Tendance des dépenses quotidiennes",
    nl: "Dagelijkse uitgaventrend",
  },
  "Average Cost": { fr: "Coût moyen", nl: "Gemiddelde kost" },
  "Average label price by day": { fr: "Prix moyen d’étiquette par jour", nl: "Gemiddelde labelprijs per dag" },
  "Avg. Domestic": { fr: "Moy. domestique", nl: "Gem. binnenlands" },
  "Recipient zone: domestic": { fr: "Zone destinataire : domestique", nl: "Ontvangerszone: binnenlands" },
  "Avg. International": { fr: "Moy. international", nl: "Gem. internationaal" },
  "Recipient zone: international": { fr: "Zone destinataire : international", nl: "Ontvangerszone: internationaal" },
  "New Shipments": { fr: "Nouveaux envois", nl: "Nieuwe zendingen" },
  "Created in selected range": { fr: "Créés sur la période sélectionnée", nl: "Aangemaakt in geselecteerde periode" },
  "Delivery Issues": { fr: "Problèmes de livraison", nl: "Leveringsproblemen" },
  "Flagged as delayed / exception": { fr: "Signalés en retard / exception", nl: "Gemarkeerd als vertraagd / uitzondering" },
  "Total Shipments": { fr: "Total envois", nl: "Totale zendingen" },
  "Daily label volume": { fr: "Volume quotidien d’étiquettes", nl: "Dagelijks labelvolume" },
  "Services": { fr: "Services", nl: "Services" },
  "Label mix by service type": { fr: "Répartition par type de service", nl: "Labelmix per servicetype" },
  "labels": { fr: "étiquettes", nl: "labels" },
  "0 labels analyzed": { fr: "0 étiquettes analysées", nl: "0 labels geanalyseerd" },
  "Recipient Zones": { fr: "Zones destinataire", nl: "Ontvangerszones" },
  "Distribution from Zone 1 to Zone 9": { fr: "Répartition de la Zone 1 à la Zone 9", nl: "Verdeling van Zone 1 tot Zone 9" },
  "Top Domestic Regions": { fr: "Principales régions domestiques", nl: "Top binnenlandse regio’s" },
  "Most frequent domestic destinations by region": { fr: "Destinations domestiques les plus fréquentes par région", nl: "Meest voorkomende binnenlandse bestemmingen per regio" },
  "Labels": { fr: "Étiquettes", nl: "Labels" },
  "Spend": { fr: "Dépenses", nl: "Spend" },
  "Top International Countries": { fr: "Principaux pays internationaux", nl: "Top internationale landen" },
  "Most frequent non-domestic destinations": { fr: "Destinations non domestiques les plus fréquentes", nl: "Meest voorkomende niet-binnenlandse bestemmingen" },
  "Shopify Fulfillment Settings": { fr: "Paramètres Shopify Fulfillment", nl: "Shopify fulfillment-instellingen" },
  "Select Fulfillment Locations": { fr: "Sélectionner les emplacements de fulfillment", nl: "Fulfillmentlocaties selecteren" },
  "Select one location to override imported orders. Select all locations to mirror your Shopify dashboard routing.": {
    fr: "Sélectionnez un emplacement pour forcer les commandes importées. Sélectionnez tous les emplacements pour refléter le routage Shopify.",
    nl: "Selecteer één locatie om geïmporteerde bestellingen te overschrijven. Selecteer alle locaties om Shopify-routing te volgen.",
  },
  "Locations": { fr: "Emplacements", nl: "Locaties" },
  "Loading locations...": { fr: "Chargement des emplacements...", nl: "Locaties laden..." },
  "Cancel": { fr: "Annuler", nl: "Annuleren" },
  "Save Settings": { fr: "Enregistrer", nl: "Instellingen opslaan" },
  "Upload CSV file": { fr: "Importer un fichier CSV", nl: "CSV-bestand uploaden" },
  "Drop your CSV here, or": { fr: "Déposez votre CSV ici, ou", nl: "Sleep je CSV hierheen, of" },
  "browse": { fr: "parcourir", nl: "bladeren" },
  ".csv files only": { fr: "fichiers .csv uniquement", nl: "alleen .csv-bestanden" },
  "Expected Fields": { fr: "Champs attendus", nl: "Verwachte velden" },
  "Name, street, city, region, postal code, country": { fr: "Nom, rue, ville, région, code postal, pays", nl: "Naam, straat, stad, regio, postcode, land" },
  "Weight and dimensions (or length, width, height)": { fr: "Poids et dimensions (ou longueur, largeur, hauteur)", nl: "Gewicht en afmetingen (of lengte, breedte, hoogte)" },
  "Optional warehouse / origin fields": { fr: "Champs optionnels d’entrepôt / d’origine", nl: "Optionele magazijn-/originevelden" },
  "Map Columns": { fr: "Mapper les colonnes", nl: "Kolommen mappen" },
  "Auto-mapped columns ready for review": { fr: "Colonnes mappées automatiquement prêtes à vérifier", nl: "Automatisch gemapte kolommen klaar voor controle" },
  "Expected Field": { fr: "Champ attendu", nl: "Verwacht veld" },
  "CSV Column": { fr: "Colonne CSV", nl: "CSV-kolom" },
  "Sample": { fr: "Exemple", nl: "Voorbeeld" },
  "Previous": { fr: "Précédent", nl: "Vorige" },
  "Next": { fr: "Suivant", nl: "Volgende" },
  "Rows {start}-{end} • Page {page}/{pages}": {
    fr: "Lignes {start}-{end} • Page {page}/{pages}",
    nl: "Rijen {start}-{end} • Pagina {page}/{pages}",
  },
  "No rows loaded": { fr: "Aucune ligne chargée", nl: "Geen rijen geladen" },
  "Upload a valid CSV with at least one data row.": { fr: "Importez un CSV valide avec au moins une ligne de données.", nl: "Upload een geldige CSV met minstens één gegevensrij." },
  "Apply Mapping": { fr: "Appliquer le mapping", nl: "Mapping toepassen" },
  "Receipt": { fr: "Reçu", nl: "Bon" },
  "Operational Receipt": { fr: "Reçu opérationnel", nl: "Operationeel ontvangstbewijs" },
  "Reference copy": { fr: "Copie de référence", nl: "Referentiekopie" },
  "Label": { fr: "Étiquette", nl: "Label" },
  "Price ex VAT": { fr: "Prix hors TVA", nl: "Prijs excl. btw" },
  "Price incl VAT": { fr: "Prix TVA incluse", nl: "Prijs incl. btw" },
  "Price ex. vat": { fr: "Prix hors TVA", nl: "Prijs excl. btw" },
  "Price incl. vat": { fr: "Prix TVA incluse", nl: "Prijs incl. btw" },
  "Price EX. VAT": { fr: "Prix hors TVA", nl: "Prijs excl. btw" },
  "Price INCL. VAT": { fr: "Prix TVA incluse", nl: "Prijs incl. btw" },
  "Download Receipt PDF": { fr: "Télécharger le reçu PDF", nl: "Bon-PDF downloaden" },
  "Preparing invoice PDF...": {
    fr: "Préparation du PDF de facture...",
    nl: "Factuur-PDF wordt voorbereid..."
  },
  "Could not generate invoice PDF.": {
    fr: "Impossible de générer le PDF de facture.",
    nl: "Kon de factuur-PDF niet genereren."
  },
  "Invoice PDF ready.": { fr: "PDF de facture prêt.", nl: "Factuur-PDF klaar." },
  "Prepared for the shipping batch shown below.": {
    fr: "Préparé pour le lot d’expédition affiché ci-dessous.",
    nl: "Opgesteld voor de hieronder getoonde verzendbatch.",
  },
  "Receipt Details": { fr: "Détails du reçu", nl: "Ontvangstbewijsdetails" },
  "Issued To": { fr: "Émis pour", nl: "Uitgegeven aan" },
  "Issued By": { fr: "Émis par", nl: "Uitgegeven door" },
  "Receipt No.": { fr: "N° reçu", nl: "Bonnr." },
  "Issued": { fr: "Émis le", nl: "Uitgegeven" },
  "Batch Size": { fr: "Taille du lot", nl: "Batchgrootte" },
  "Unit Rate": { fr: "Tarif unitaire", nl: "Eenheidstarief" },
  "Average Unit Rate": { fr: "Tarif unitaire moyen", nl: "Gemiddeld eenheidstarief" },
  "Billing Profile": { fr: "Profil de facturation", nl: "Facturatieprofiel" },
  "Contact": { fr: "Contact", nl: "Contact" },
  "Tax Profile": { fr: "Profil fiscal", nl: "Fiscaal profiel" },
  "Recipient Batch": { fr: "Lot destinataires", nl: "Ontvangersbatch" },
  "One line per generated label in this order.": {
    fr: "Une ligne par étiquette générée dans cette commande.",
    nl: "Eén regel per gegenereerd label in deze bestelling.",
  },
  "Destination": { fr: "Destination", nl: "Bestemming" },
  "Parcel": { fr: "Colis", nl: "Pakket" },
  "Weight": { fr: "Poids", nl: "Gewicht" },
  "Prepared from Shipide Portal history.": {
    fr: "Préparé depuis l’historique Shipide Portal.",
    nl: "Opgesteld vanuit de Shipide Portal-historiek.",
  },
  "Document generated {date}": { fr: "Document généré {date}", nl: "Document gegenereerd {date}" },
  "Questions about billing should be handled through your monthly invoice workflow.": {
    fr: "Les questions de facturation doivent être traitées via votre processus de facture mensuelle.",
    nl: "Vragen over facturatie moeten via je maandelijkse factuurproces worden afgehandeld.",
  },
  "Company Information": { fr: "Informations société", nl: "Bedrijfsinformatie" },
  "Billing Notice": { fr: "Mention de facturation", nl: "Factureringsmelding" },
  "This receipt is provided for operational reference only. It is not valid for tax or accounting purposes. Your invoice is the only valid billing document for bookkeeping.": {
    fr: "Ce reçu est fourni à titre de référence opérationnelle uniquement. Il n’est pas valable à des fins fiscales ou comptables. Votre facture est le seul document de facturation valable pour la comptabilité.",
    nl: "Dit ontvangstbewijs wordt alleen ter operationele referentie verstrekt. Het is niet geldig voor fiscale of boekhoudkundige doeleinden. Je factuur is het enige geldige facturatiedocument voor je boekhouding.",
  },
  "Preparing receipt PDF...": {
    fr: "Préparation du PDF du reçu...",
    nl: "PDF van ontvangstbewijs voorbereiden...",
  },
  "Receipt PDF ready.": { fr: "PDF du reçu prêt.", nl: "PDF van ontvangstbewijs klaar." },
  "Could not generate receipt PDF.": {
    fr: "Impossible de générer le PDF du reçu.",
    nl: "Kon de PDF van het ontvangstbewijs niet genereren.",
  },
  "Summary": { fr: "Résumé", nl: "Samenvatting" },
  "Generated Labels": { fr: "Étiquettes générées", nl: "Gegenereerde labels" },
  "Compact operational summary for this shipment batch.": {
    fr: "Résumé opérationnel compact pour ce lot d’expédition.",
    nl: "Compact operationeel overzicht voor deze verzendbatch.",
  },
  "One line per generated label.": {
    fr: "Une ligne par étiquette générée.",
    nl: "Eén regel per gegenereerd label.",
  },
  "Amount": { fr: "Montant", nl: "Bedrag" },
  "Restricted Goods": { fr: "Marchandises interdites", nl: "Beperkte goederen" },
  "Shopify fulfillment settings": { fr: "paramètres Shopify fulfillment", nl: "Shopify fulfillment-instellingen" },
  "e.g. Cotton t-shirt, 100% cotton": { fr: "ex. T-shirt coton, 100% coton", nl: "bv. Katoenen T-shirt, 100% katoen" },
  "Label PDF preview": { fr: "Aperçu PDF étiquette", nl: "Label-PDF voorbeeld" },
  "Account label PDF preview": { fr: "Aperçu PDF compte", nl: "Account label-PDF voorbeeld" },
  "Total payments chart": { fr: "Graphique paiements totaux", nl: "Grafiek totale betalingen" },
  "Average cost chart": { fr: "Graphique coût moyen", nl: "Grafiek gemiddelde kost" },
  "Total shipments chart": { fr: "Graphique total envois", nl: "Grafiek totale zendingen" },
  "Service distribution chart": { fr: "Graphique répartition services", nl: "Grafiek serviceverdeling" },
  "Domestic regions map": { fr: "Carte des régions domestiques", nl: "Kaart binnenlandse regio's" },
  "International countries map": { fr: "Carte des pays internationaux", nl: "Kaart internationale landen" },
  "Current postal legislation prohibits sending mail or parcels with the following contents:": {
    fr: "La législation postale actuelle interdit l’envoi de courrier ou colis contenant les éléments suivants :",
    nl: "De huidige postwetgeving verbiedt het verzenden van post of pakketten met de volgende inhoud:",
  },
  "Dangerous goods are items whose shape, nature, or packaging can endanger people or damage other shipments, postal equipment, or third-party goods. Some everyday products are also prohibited in the postal network, including lighters, perfumes, manicure products, glue, paint, flammable liquids, aerosols, pressurized deodorants, hair dyes, and lithium batteries.": {
    fr: "Les marchandises dangereuses sont des objets dont la forme, la nature ou l’emballage peuvent mettre des personnes en danger ou endommager d’autres envois, l’équipement postal ou des biens tiers. Certains produits du quotidien sont aussi interdits dans le réseau postal, notamment les briquets, parfums, produits de manucure, colle, peinture, liquides inflammables, aérosols, déodorants sous pression, colorations capillaires et batteries au lithium.",
    nl: "Gevaarlijke goederen zijn items waarvan de vorm, aard of verpakking mensen in gevaar kan brengen of andere zendingen, postapparatuur of goederen van derden kan beschadigen. Sommige alledaagse producten zijn ook verboden in het postnetwerk, waaronder aanstekers, parfums, manicureproducten, lijm, verf, brandbare vloeistoffen, aerosols, onder druk staande deodoranten, haarkleurmiddelen en lithiumbatterijen.",
  },
  "Dangerous Goods": { fr: "Marchandises dangereuses", nl: "Gevaarlijke goederen" },
  "A specific exception exists for certain lithium batteries. Refer to the dangerous goods overview for the complete rules and detailed list.": {
    fr: "Une exception spécifique existe pour certaines batteries au lithium. Consultez l’aperçu des marchandises dangereuses pour les règles complètes.",
    nl: "Voor bepaalde lithiumbatterijen bestaat een specifieke uitzondering. Raadpleeg het overzicht van gevaarlijke goederen voor alle regels.",
  },
  "Prohibited Items": { fr: "Objets interdits", nl: "Verboden artikelen" },
  "Drugs, narcotics, and psychotropic substances.": { fr: "Drogues, stupéfiants et substances psychotropes.", nl: "Drugs, verdovende middelen en psychotrope stoffen." },
  "Weapons, essential weapon components, and ammunition; imitation weapons; knives, swords, daggers, and other sharpened or cutting objects.": {
    fr: "Armes, composants essentiels d’armes et munitions ; armes factices ; couteaux, épées, poignards et autres objets tranchants ou coupants.",
    nl: "Wapens, essentiële wapenonderdelen en munitie; imitatie-wapens; messen, zwaarden, dolken en andere geslepen of snijdende voorwerpen.",
  },
  "Items, writings, or substances whose import, export, production, circulation, distribution, use, possession, sale, or transport is generally prohibited by law.": {
    fr: "Objets, écrits ou substances dont l’importation, l’exportation, la production, la circulation, la distribution, l’usage, la détention, la vente ou le transport est généralement interdit par la loi.",
    nl: "Items, geschriften of stoffen waarvan invoer, uitvoer, productie, circulatie, distributie, gebruik, bezit, verkoop of vervoer wettelijk verboden is.",
  },
  "Items with markings or inscriptions contrary to public order or accepted standards of decency.": {
    fr: "Objets portant des inscriptions contraires à l’ordre public ou aux bonnes mœurs.",
    nl: "Items met markeringen of opschriften die strijdig zijn met openbare orde of goede zeden.",
  },
  "Shipments whose destination or contents are prohibited under trade law.": {
    fr: "Envois dont la destination ou le contenu est interdit par le droit du commerce.",
    nl: "Zendingen waarvan bestemming of inhoud verboden is onder handelswetgeving.",
  },
  "Shipments containing counterfeit goods.": { fr: "Envois contenant des marchandises contrefaites.", nl: "Zendingen met namaakgoederen." },
  "Coins, banknotes, and bearer securities.": { fr: "Pièces de monnaie, billets de banque et titres au porteur.", nl: "Munten, bankbiljetten en effecten aan toonder." },
  "Jewelry (except costume jewelry without gold, silver, or precious stones), artworks, collections, animal skins, and other high-value materials.": {
    fr: "Bijoux (sauf bijoux fantaisie sans or, argent ou pierres précieuses), œuvres d’art, collections, peaux d’animaux et autres matériaux de grande valeur.",
    nl: "Juwelen (behalve fantasiejuwelen zonder goud, zilver of edelstenen), kunstwerken, verzamelingen, dierenhuiden en andere hoogwaardige materialen.",
  },
  "Parcels with a declared value above €5,000.": { fr: "Colis avec une valeur déclarée supérieure à 5 000 €.", nl: "Pakketten met een aangegeven waarde boven €5.000." },
  "Supabase auth is not configured.": { fr: "L’authentification Supabase n’est pas configurée.", nl: "Supabase-auth is niet geconfigureerd." },
  "Email and password are required.": { fr: "E-mail et mot de passe requis.", nl: "E-mail en wachtwoord zijn verplicht." },
  "Signing In...": { fr: "Connexion...", nl: "Inloggen..." },
  "Creating...": { fr: "Création...", nl: "Aanmaken..." },
  "Account created. If confirmation is enabled, approve the email then sign in.": { fr: "Compte créé. Si la confirmation est activée, validez l’e-mail puis connectez-vous.", nl: "Account aangemaakt. Als bevestiging actief is, bevestig je e-mail en log daarna in." },
  "Enter your email address, then click Forgot password.": { fr: "Entrez votre e-mail, puis cliquez sur Mot de passe oublié.", nl: "Vul je e-mailadres in en klik daarna op Wachtwoord vergeten." },
  "If that email exists, a reset link has been sent. Check your inbox.": { fr: "Si cet e-mail existe, un lien de réinitialisation a été envoyé. Vérifiez votre boîte mail.", nl: "Als dit e-mailadres bestaat, is een resetlink verzonden. Controleer je inbox." },
  "Supabase client failed to initialize.": { fr: "Échec d’initialisation du client Supabase.", nl: "Supabase-client kon niet worden geïnitialiseerd." },
  "Sign in before configuring Shopify locations.": { fr: "Connectez-vous avant de configurer les emplacements Shopify.", nl: "Log in voordat je Shopify-locaties configureert." },
  "Connect Shopify before opening fulfillment settings.": { fr: "Connectez Shopify avant d’ouvrir les paramètres de fulfillment.", nl: "Verbind Shopify voordat je fulfillmentinstellingen opent." },
  "Connect Shopify before saving settings.": { fr: "Connectez Shopify avant d’enregistrer les paramètres.", nl: "Verbind Shopify voordat je instellingen opslaat." },
  "Loading Shopify locations...": { fr: "Chargement des emplacements Shopify...", nl: "Shopify-locaties laden..." },
  "Could not load Shopify locations.": { fr: "Impossible de charger les emplacements Shopify.", nl: "Kan Shopify-locaties niet laden." },
  "Connect Shopify first.": { fr: "Connectez Shopify d’abord.", nl: "Verbind eerst Shopify." },
  "Select at least one location.": { fr: "Sélectionnez au moins un emplacement.", nl: "Selecteer minstens één locatie." },
  "Shopify fulfillment settings updated.": { fr: "Paramètres Shopify fulfillment mis à jour.", nl: "Shopify fulfillment-instellingen bijgewerkt." },
  "Settings saved.": { fr: "Paramètres enregistrés.", nl: "Instellingen opgeslagen." },
  "Could not save Shopify settings.": { fr: "Impossible d’enregistrer les paramètres Shopify.", nl: "Kan Shopify-instellingen niet opslaan." },
  "Saving...": { fr: "Enregistrement...", nl: "Opslaan..." },
  "{count} selected": { fr: "{count} sélectionné(s)", nl: "{count} geselecteerd" },
  "{count} selected (all)": { fr: "{count} sélectionné(s) (tous)", nl: "{count} geselecteerd (alle)" },
  "No locations available": { fr: "Aucun emplacement disponible", nl: "Geen locaties beschikbaar" },
  "All Locations": { fr: "Tous les emplacements", nl: "Alle locaties" },
  "Mirror Shopify location routing": { fr: "Reproduire le routage Shopify", nl: "Shopify-locatierouting volgen" },
  "No address details": { fr: "Aucune adresse", nl: "Geen adresgegevens" },
  "You must be signed in to use provider import.": { fr: "Vous devez être connecté pour utiliser l’import fournisseur.", nl: "Je moet ingelogd zijn om providerimport te gebruiken." },
  "Shopify request timed out. Check Worker route and secrets.": { fr: "La requête Shopify a expiré. Vérifiez la route Worker et les secrets.", nl: "Shopify-verzoek time-out. Controleer Worker-route en secrets." },
  "Request timed out. Please try again.": {
    fr: "La requête a expiré. Veuillez réessayer.",
    nl: "Verzoek time-out. Probeer opnieuw.",
  },
  "Could not load Shopify connection.": { fr: "Impossible de charger la connexion Shopify.", nl: "Kan Shopify-verbinding niet laden." },
  "Shopify connected: {shop}": { fr: "Shopify connecté : {shop}", nl: "Shopify verbonden: {shop}" },
  "Shopify connected.": { fr: "Shopify connecté.", nl: "Shopify verbonden." },
  "Could not connect Shopify.": { fr: "Impossible de connecter Shopify.", nl: "Kon Shopify niet verbinden." },
  "No usable {sourceLabel} rows found.": { fr: "Aucune ligne {sourceLabel} exploitable trouvée.", nl: "Geen bruikbare {sourceLabel}-rijen gevonden." },
  "Enter your Shopify store domain (example: your-store.myshopify.com)": { fr: "Entrez le domaine de votre boutique Shopify (exemple : your-store.myshopify.com)", nl: "Voer je Shopify-winkeldomein in (voorbeeld: your-store.myshopify.com)" },
  "Enter a valid .myshopify.com domain.": { fr: "Entrez un domaine .myshopify.com valide.", nl: "Voer een geldig .myshopify.com-domein in." },
  "Connecting Shopify...": { fr: "Connexion à Shopify...", nl: "Shopify verbinden..." },
  "Shopify install URL was not returned.": { fr: "L’URL d’installation Shopify n’a pas été renvoyée.", nl: "Shopify-installatie-URL werd niet teruggegeven." },
  "Could not start Shopify connect flow.": { fr: "Impossible de démarrer la connexion Shopify.", nl: "Kon Shopify-verbindingsflow niet starten." },
  "Shopify locations endpoint is not live yet. Deploy the latest API worker (or restart the local Node server) and try again.": {
    fr: "L’endpoint Shopify locations n’est pas encore en ligne. Déployez le dernier worker API (ou redémarrez le serveur local Node) puis réessayez.",
    nl: "De Shopify-locaties endpoint is nog niet live. Deploy de nieuwste API-worker (of herstart de lokale Node-server) en probeer opnieuw.",
  },
  "Connect Shopify before importing.": { fr: "Connectez Shopify avant d’importer.", nl: "Verbind Shopify vóór import." },
  "Importing orders from {shop}...": { fr: "Import des commandes depuis {shop}...", nl: "Bestellingen importeren van {shop}..." },
  "Imported {count} orders from {shop}.": { fr: "{count} commandes importées depuis {shop}.", nl: "{count} bestellingen geïmporteerd van {shop}." },
  "Imported {count} Shopify orders": { fr: "{count} commandes Shopify importées", nl: "{count} Shopify-bestellingen geïmporteerd" },
  "Shopify connection expired. Click Shopify again to reconnect.": { fr: "Connexion Shopify expirée. Cliquez à nouveau sur Shopify pour reconnecter.", nl: "Shopify-verbinding verlopen. Klik opnieuw op Shopify om te herverbinden." },
  "Sign in before importing from Shopify.": { fr: "Connectez-vous avant d’importer depuis Shopify.", nl: "Log in vóór importeren vanuit Shopify." },
  "Shopify import failed.": { fr: "Échec de l’import Shopify.", nl: "Shopify-import mislukt." },
  "Sign in to use account shipping origins.": { fr: "Connectez-vous pour utiliser les origines d’expédition du compte.", nl: "Log in om account-verzendorigines te gebruiken." },
  "Sign in to manage account shipping origins.": { fr: "Connectez-vous pour gérer les origines d’expédition du compte.", nl: "Log in om account-verzendorigines te beheren." },
  "Add at least one shipping origin in Account settings.": { fr: "Ajoutez au moins une origine d’expédition dans les paramètres du compte.", nl: "Voeg minstens één verzendorigine toe in accountinstellingen." },
  "Add at least one shipping origin in Account settings before continuing.": { fr: "Ajoutez au moins une origine d’expédition dans les paramètres du compte avant de continuer.", nl: "Voeg minstens één verzendorigine toe in accountinstellingen voordat je verdergaat." },
  "Sender is fixed to your only saved shipping origin.": { fr: "L’expéditeur est fixé à votre seule origine enregistrée.", nl: "Afzender staat vast op je enige opgeslagen oorsprong." },
  "No saved ship-from origin. Add one in Account settings.": { fr: "Aucune origine d’expédition enregistrée. Ajoutez-en une dans les paramètres du compte.", nl: "Geen opgeslagen verzendorigine. Voeg er een toe in accountinstellingen." },
  "Ship from is controlled by your Shopify": { fr: "L’origine d’expédition est contrôlée par vos", nl: "Ship from wordt beheerd door je Shopify" },
  "fulfillment settings": { fr: "paramètres fulfillment", nl: "fulfillment-instellingen" },
  "for this import.": { fr: "pour cet import.", nl: "voor deze import." },
  "Ship from is controlled by your Shopify fulfillment settings for this import.": { fr: "L’origine d’expédition est pilotée par vos paramètres Shopify fulfillment pour cet import.", nl: "Ship from wordt bepaald door je Shopify fulfillment-instellingen voor deze import." },
  "Using your only saved ship-from origin.": { fr: "Utilisation de votre seule origine enregistrée.", nl: "Je enige opgeslagen verzendorigine wordt gebruikt." },
  "Choose which origin to apply to this batch.": { fr: "Choisissez l’origine à appliquer à ce lot.", nl: "Kies welke oorsprong op deze batch wordt toegepast." },
  "Unsaved changes in shipping origins.": { fr: "Modifications non enregistrées dans les origines d’expédition.", nl: "Niet-opgeslagen wijzigingen in verzendorigines." },
  "Origin": { fr: "Origine", nl: "Oorsprong" },
  "Warehouse": { fr: "Entrepôt", nl: "Magazijn" },
  "warehouse": { fr: "entrepôt", nl: "magazijn" },
  "Applied {name} to sender details.": { fr: "{name} appliqué aux données expéditeur.", nl: "{name} toegepast op afzendergegevens." },
  "Physical": { fr: "Physique", nl: "Fysiek" },
  "Return": { fr: "Retour", nl: "Retour" },
  "Same as physical": { fr: "Identique au physique", nl: "Zelfde als fysiek" },
  "Not set": { fr: "Non défini", nl: "Niet ingesteld" },
  "Origin Label": { fr: "Nom de l’origine", nl: "Naam oorsprong" },
  "Sender Name": { fr: "Nom expéditeur", nl: "Naam afzender" },
  "Sender/Company name": { fr: "Nom expéditeur/entreprise", nl: "Naam afzender/bedrijf" },
  "Default": { fr: "Par défaut", nl: "Standaard" },
  "Same Return Address": { fr: "Même adresse de retour", nl: "Zelfde retouradres" },
  "Return Name": { fr: "Nom de retour", nl: "Retournaam" },
  "Return Street": { fr: "Rue de retour", nl: "Retourstraat" },
  "Return City": { fr: "Ville de retour", nl: "Retourstad" },
  "Return Region": { fr: "Région de retour", nl: "Retourregio" },
  "Return Postal Code": { fr: "Code postal retour", nl: "Retourpostcode" },
  "Return Country": { fr: "Pays de retour", nl: "Retourland" },
  "Return sender/company": { fr: "Nom retour/entreprise", nl: "Retourafzender/bedrijf" },
  "Apply": { fr: "Appliquer", nl: "Toepassen" },
  "Remove": { fr: "Supprimer", nl: "Verwijderen" },
  "Add at least one warehouse origin.": { fr: "Ajoutez au moins une origine d’entrepôt.", nl: "Voeg minstens één magazijnoorsprong toe." },
  "Maximum {max} warehouses reached.": { fr: "Maximum de {max} entrepôts atteint.", nl: "Maximum {max} magazijnen bereikt." },
  "At least one warehouse origin is required.": { fr: "Au moins une origine d’entrepôt est requise.", nl: "Minstens één magazijnoorsprong is vereist." },
  "Loading shipping origins...": { fr: "Chargement des origines d’expédition...", nl: "Verzendorigines laden..." },
  "Shipping origins synced from your account.": { fr: "Origines d’expédition synchronisées depuis votre compte.", nl: "Verzendorigines gesynchroniseerd vanuit je account." },
  "Showing browser-saved origins. Click Save to sync.": { fr: "Origines locales affichées. Cliquez sur Enregistrer pour synchroniser.", nl: "Browser-opgeslagen origines getoond. Klik op Opslaan om te synchroniseren." },
  "Add your shipping origin and click Save.": { fr: "Ajoutez votre origine d’expédition puis cliquez sur Enregistrer.", nl: "Voeg je verzendorigine toe en klik op Opslaan." },
  "Saving shipping origins...": { fr: "Enregistrement des origines d’expédition...", nl: "Verzendorigines opslaan..." },
  "Could not sync to account. Supabase client is unavailable.": { fr: "Impossible de synchroniser au compte. Client Supabase indisponible.", nl: "Kan niet met account synchroniseren. Supabase-client is niet beschikbaar." },
  "Could not sync shipping origins: {error}": { fr: "Impossible de synchroniser les origines d’expédition : {error}", nl: "Kan verzendorigines niet synchroniseren: {error}" },
  "unknown error": { fr: "erreur inconnue", nl: "onbekende fout" },
  "Shipping origins saved to your account.": { fr: "Origines d’expédition enregistrées sur votre compte.", nl: "Verzendorigines opgeslagen in je account." },
  "No shipping origins configured yet.": { fr: "Aucune origine d’expédition configurée pour l’instant.", nl: "Nog geen verzendorigines geconfigureerd." },
  "{prefix}: origin label is required.": { fr: "{prefix} : le nom de l’origine est requis.", nl: "{prefix}: oorsprongsnaam is verplicht." },
  "{prefix}: sender name is required.": { fr: "{prefix} : le nom expéditeur est requis.", nl: "{prefix}: afzendernaam is verplicht." },
  "{prefix}: street is required.": { fr: "{prefix} : la rue est requise.", nl: "{prefix}: straat is verplicht." },
  "{prefix}: city is required.": { fr: "{prefix} : la ville est requise.", nl: "{prefix}: stad is verplicht." },
  "{prefix}: postal code is required.": { fr: "{prefix} : le code postal est requis.", nl: "{prefix}: postcode is verplicht." },
  "{prefix}: country is required.": { fr: "{prefix} : le pays est requis.", nl: "{prefix}: land is verplicht." },
  "{prefix}: return sender name is required when using a custom return address.": { fr: "{prefix} : le nom retour est requis avec une adresse retour personnalisée.", nl: "{prefix}: retourafzender is verplicht bij een aangepast retouradres." },
  "{prefix}: return street is required when using a custom return address.": { fr: "{prefix} : la rue retour est requise avec une adresse retour personnalisée.", nl: "{prefix}: retourstraat is verplicht bij een aangepast retouradres." },
  "{prefix}: return city is required when using a custom return address.": { fr: "{prefix} : la ville retour est requise avec une adresse retour personnalisée.", nl: "{prefix}: retourstad is verplicht bij een aangepast retouradres." },
  "{prefix}: return postal code is required when using a custom return address.": { fr: "{prefix} : le code postal retour est requis avec une adresse retour personnalisée.", nl: "{prefix}: retourpostcode is verplicht bij een aangepast retouradres." },
  "{prefix}: return country is required when using a custom return address.": { fr: "{prefix} : le pays retour est requis avec une adresse retour personnalisée.", nl: "{prefix}: retourland is verplicht bij een aangepast retouradres." },
  "No generations yet.": { fr: "Aucune génération pour l’instant.", nl: "Nog geen generaties." },
  "Label generation": { fr: "Génération d’étiquette", nl: "Labelgeneratie" },
  "{count} labels • ex. vat {ex} • incl. vat {incl}": {
    fr: "{count} étiquettes • Total {incl}",
    nl: "{count} labels • Totaal {incl}",
  },
  "Loading previous generations...": { fr: "Chargement des générations précédentes...", nl: "Vorige generaties laden..." },
  "No preview-ready generations yet.": { fr: "Aucune génération prête pour prévisualisation.", nl: "Nog geen generaties klaar voor preview." },
  "No previous generations yet.": { fr: "Aucune génération précédente.", nl: "Nog geen vorige generaties." },
  "Sync delayed. Showing last synced history.": { fr: "Synchronisation en retard. Affichage du dernier historique synchronisé.", nl: "Synchronisatie vertraagd. Laatst gesynchroniseerde historiek getoond." },
  "Showing browser-saved history for this account.": { fr: "Affichage de l’historique enregistré dans le navigateur pour ce compte.", nl: "Browser-opgeslagen historiek voor dit account wordt getoond." },
  "No browser history saved for this account yet.": { fr: "Aucun historique navigateur enregistré pour ce compte.", nl: "Nog geen browserhistoriek opgeslagen voor dit account." },
  "Sync delayed. Showing previously loaded history.": { fr: "Synchronisation en retard. Affichage de l’historique déjà chargé.", nl: "Synchronisatie vertraagd. Eerder geladen historiek wordt getoond." },
  "Could not load history right now. Please refresh in a moment.": { fr: "Impossible de charger l’historique pour le moment. Veuillez actualiser.", nl: "Kan historiek nu niet laden. Vernieuw zo meteen." },
  "Date": { fr: "Date", nl: "Datum" },
  "Order": { fr: "Commande", nl: "Bestelling" },
  "EX": { fr: "Sous-total", nl: "Subtotaal" },
  "INCL": { fr: "Total", nl: "Totaal" },
  "VAT ({vat}%)": { fr: "Taxe ({vat}%)", nl: "Belasting ({vat}%)" },
  "Unit Price": { fr: "Prix unitaire", nl: "Eenheidsprijs" },
  "Unit (ex VAT)": { fr: "Prix unitaire", nl: "Eenheidsprijs" },
  "Unit (incl VAT)": { fr: "Prix unitaire", nl: "Eenheidsprijs" },
  "Unit (ex. vat)": { fr: "Prix unitaire", nl: "Eenheidsprijs" },
  "Unit (incl. vat)": { fr: "Prix unitaire", nl: "Eenheidsprijs" },
  "Unit (EX. VAT)": { fr: "Prix unitaire", nl: "Eenheidsprijs" },
  "Unit (INCL. VAT)": { fr: "Prix unitaire", nl: "Eenheidsprijs" },
  "SHIPR RECEIPT": { fr: "REÇU SHIPR", nl: "SHIPR BON" },
  "LABEL BREAKDOWN": { fr: "DÉTAIL DES ÉTIQUETTES", nl: "LABELOVERZICHT" },
  "label-batch": { fr: "lot-etiquettes", nl: "label-batch" },
  "label-order": { fr: "commande-etiquette", nl: "label-order" },
  "{service} • {count} labels • ex. vat {ex} • vat {vat} • incl. vat {incl}": {
    fr: "{service} • {count} étiquettes • Total {incl}",
    nl: "{service} • {count} labels • Totaal {incl}",
  },
  "{service} • {count} labels • EX. VAT {ex} • VAT {vat} • INCL. VAT {incl}": {
    fr: "{service} • {count} étiquettes • Total {incl}",
    nl: "{service} • {count} labels • Totaal {incl}",
  },
  "{count} labels • EX. VAT {ex} • INCL. VAT {incl}": {
    fr: "{count} étiquettes • Total {incl}",
    nl: "{count} labels • Totaal {incl}",
  },
  "{service} • {count} labels • Total {amount}": {
    fr: "{service} • {count} étiquettes • Total {amount}",
    nl: "{service} • {count} labels • Totaal {amount}",
  },
  "{count} labels • Total {amount}": {
    fr: "{count} étiquettes • Total {amount}",
    nl: "{count} labels • Totaal {amount}",
  },
  "{count} labels analyzed": { fr: "{count} étiquettes analysées", nl: "{count} labels geanalyseerd" },
  "{count} labels": { fr: "{count} étiquettes", nl: "{count} labels" },
  "{amount} pending": { fr: "{amount} en attente", nl: "{amount} in behandeling" },
  "ex. vat • incl. vat {amount}": {
    fr: "Total {amount}",
    nl: "Totaal {amount}",
  },
  "EX. VAT • INCL. VAT {amount}": {
    fr: "Total {amount}",
    nl: "Totaal {amount}",
  },
  "Total {amount}": { fr: "Total {amount}", nl: "Totaal {amount}" },
  "7 days": { fr: "7 jours", nl: "7 dagen" },
  "30 days": { fr: "30 jours", nl: "30 dagen" },
  "90 days": { fr: "90 jours", nl: "90 dagen" },
  "1 year": { fr: "1 an", nl: "1 jaar" },
  "{count} events": { fr: "{count} événements", nl: "{count} events" },
  "{amount} spend": { fr: "{amount} dépenses", nl: "{amount} spend" },
  "{rate}% rate": { fr: "{rate}% taux", nl: "{rate}% ratio" },
  "Payments": { fr: "Paiements", nl: "Betalingen" },
  "Average Cost": { fr: "Coût moyen", nl: "Gemiddelde kost" },
  "Shipments": { fr: "Envois", nl: "Zendingen" },
  "No service data in this range.": { fr: "Aucune donnée de service sur cette période.", nl: "Geen servicedata in dit bereik." },
  "Zone {index}": { fr: "Zone {index}", nl: "Zone {index}" },
  "No domestic region data yet.": { fr: "Aucune donnée régionale domestique pour l’instant.", nl: "Nog geen binnenlandse regiogegevens." },
  "No international country data yet.": { fr: "Aucune donnée pays internationale pour l’instant.", nl: "Nog geen internationale landgegevens." },
  "Unknown": { fr: "Inconnu", nl: "Onbekend" },
  "Unknown Service": { fr: "Service inconnu", nl: "Onbekende service" },
  "North Region": { fr: "Région Nord", nl: "Noordelijke regio" },
  "South Region": { fr: "Région Sud", nl: "Zuidelijke regio" },
  "Capital Region": { fr: "Région Capitale", nl: "Hoofdstedelijke regio" },
  "Region": { fr: "Région", nl: "Regio" },
  "Could not parse this CSV. Please check the file format and try again.": { fr: "Impossible de lire ce CSV. Vérifiez le format et réessayez.", nl: "Kon deze CSV niet verwerken. Controleer het formaat en probeer opnieuw." },
  "No rows detected": { fr: "Aucune ligne détectée", nl: "Geen rijen gedetecteerd" },
  "{file} • {rows} rows • {columns} columns": { fr: "{file} • {rows} lignes • {columns} colonnes", nl: "{file} • {rows} rijen • {columns} kolommen" },
  "Required": { fr: "Obligatoire", nl: "Verplicht" },
  "Not mapped": { fr: "Non mappé", nl: "Niet gemapt" },
  "Column {index}": { fr: "Colonne {index}", nl: "Kolom {index}" },
  "No sample": { fr: "Aucun exemple", nl: "Geen voorbeeld" },
  "Map required fields: {fields}.": { fr: "Mappez les champs obligatoires : {fields}.", nl: "Map verplichte velden: {fields}." },
  "Required fields must map to separate CSV columns ({fields}).": { fr: "Les champs obligatoires doivent être mappés sur des colonnes CSV distinctes ({fields}).", nl: "Verplichte velden moeten op aparte CSV-kolommen worden gemapt ({fields})." },
  "This mapping produced no usable rows. Adjust mapping and try again.": { fr: "Ce mapping ne produit aucune ligne exploitable. Ajustez puis réessayez.", nl: "Deze mapping leverde geen bruikbare rijen op. Pas mapping aan en probeer opnieuw." },
  "Map CSV columns": { fr: "Mapper les colonnes CSV", nl: "CSV-kolommen mappen" },
  "Preparing Label...": { fr: "Préparation de l’étiquette...", nl: "Label voorbereiden..." },
  "{country} is outside the customs union. Complete declaration details before service selection.": {
    fr: "{country} est hors union douanière. Complétez la déclaration avant de choisir le service.",
    nl: "{country} ligt buiten de douane-unie. Vul de aangifte in vóór je een service kiest.",
  },
  "{count} destinations are outside the customs union. Complete declaration details before service selection.": {
    fr: "{count} destinations sont hors union douanière. Complétez la déclaration avant de choisir le service.",
    nl: "{count} bestemmingen liggen buiten de douane-unie. Vul de aangifte in vóór je een service kiest.",
  },
  "Item {index}": { fr: "Article {index}", nl: "Item {index}" },
  "Provide a precise product description and material/composition.": {
    fr: "Indiquez une description précise du produit et sa matière/composition.",
    nl: "Geef een precieze productbeschrijving en materiaal/samenstelling.",
  },
  "Item Description": { fr: "Description de l’article", nl: "Itemomschrijving" },
  "Number of Items": { fr: "Nombre d’articles", nl: "Aantal items" },
  "Value (EUR)": { fr: "Valeur (EUR)", nl: "Waarde (EUR)" },
  "Origin of Items": { fr: "Origine des articles", nl: "Oorsprong van items" },
  "HS Tariff Code (optional)": { fr: "Code tarifaire SH (optionnel)", nl: "HS-tariefcode (optioneel)" },
  "Imported from {provider}": { fr: "Importé depuis {provider}", nl: "Geïmporteerd van {provider}" },
  "Failed to load domestic regions GeoJSON ({status})": { fr: "Échec du chargement GeoJSON régions ({status})", nl: "Laden van GeoJSON-regio's mislukt ({status})" },
  "Failed to load world GeoJSON ({status})": { fr: "Échec du chargement GeoJSON monde ({status})", nl: "Laden van wereld-GeoJSON mislukt ({status})" },
  "Unknown history fetch error": { fr: "Erreur inconnue lors du chargement de l’historique", nl: "Onbekende fout bij laden van historiek" },
  "Supabase unavailable": { fr: "Supabase indisponible", nl: "Supabase niet beschikbaar" },
  "Could not load this page preview.": { fr: "Impossible de charger cet aperçu.", nl: "Kan deze preview niet laden." },
  "Domestic map data unavailable": { fr: "Données cartographiques domestiques indisponibles", nl: "Binnenlandse kaartdata niet beschikbaar" },
  "Request failed ({status})": { fr: "Requête échouée ({status})", nl: "Verzoek mislukt ({status})" },
};
// Toggle to compare auth logo layouts quickly:
// "inline" => logo replaces Shipr header block inside card.
// "stack" => logo above the card (previous behavior).
const AUTH_LOGO_LAYOUT = "stack";
const ROUTE_BASE_PATH = detectRouteBasePath(window.location.pathname);
const SUPABASE_AUTH_STORAGE_KEY = "sb-pxcqxubehvnyaubqjcrf-auth-token-shipr";

function createSupabaseAuthLock() {
  const lockQueues = new Map();
  return async (...args) => {
    const callback = args[args.length - 1];
    if (typeof callback !== "function") {
      return null;
    }

    const lockName = typeof args[0] === "string" ? args[0] : "sb-auth-lock";
    const previous = lockQueues.get(lockName) || Promise.resolve();
    let releaseLock = null;
    const current = new Promise((resolve) => {
      releaseLock = resolve;
    });
    lockQueues.set(
      lockName,
      previous.then(
        () => current,
        () => current
      )
    );

    await previous.catch(() => {});
    try {
      return await callback();
    } finally {
      releaseLock();
      if (lockQueues.get(lockName) === current) {
        lockQueues.delete(lockName);
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
const authPixelCanvas = document.getElementById("authPixelCanvas");
const authBlastCanvas = document.getElementById("authBlastCanvas");
const authBgSwitcher = document.getElementById("authBgSwitcher");
const authBgVariantButtons = authBgSwitcher
  ? Array.from(authBgSwitcher.querySelectorAll("[data-auth-bg-variant]"))
  : [];
const authStack = authGate?.querySelector(".auth-stack") || null;
const authCard = authGate?.querySelector(".auth-card") || null;
const authCardFrame = document.getElementById("authCardFrame");
const authForm = document.getElementById("authForm");
const authStepOneShell = authForm?.querySelector(".auth-step-one-shell") || null;
const authTitle = document.getElementById("authTitle");
const authSubtitle = document.getElementById("authSubtitle");
const authEmail = document.getElementById("authEmail");
const authPassword = document.getElementById("authPassword");
const authPasswordConfirm = document.getElementById("authPasswordConfirm");
const authRegisterOnlyFields = document.querySelectorAll(".auth-register-field");
const authRegisterProgress = document.getElementById("authRegisterProgress");
const authRegisterProgressItems = authRegisterProgress
  ? Array.from(authRegisterProgress.querySelectorAll(".auth-register-progress-item"))
  : [];
const authRegisterStepInfo = document.getElementById("authRegisterStepInfo");
const authCompanyName = document.getElementById("authCompanyName");
const authContactName = document.getElementById("authContactName");
const authContactEmail = document.getElementById("authContactEmail");
const authContactPhone = document.getElementById("authContactPhone");
const authBillingStreet = document.getElementById("authBillingStreet");
const authBillingCity = document.getElementById("authBillingCity");
const authBillingPostalCode = document.getElementById("authBillingPostalCode");
const authBillingCountry = document.getElementById("authBillingCountry");
const authBillingCountryFlag = document.getElementById("authBillingCountryFlag");
const authBillingAddress = document.getElementById("authBillingAddress");
const authTaxId = document.getElementById("authTaxId");
const authCustomerId = document.getElementById("authCustomerId");
const authAgreementGroup = document.getElementById("authAgreementGroup");
const authAgreementVersion = document.getElementById("authAgreementVersion");
const authAgreementDownload = document.getElementById("authAgreementDownload");
const authAgreementScroll = document.getElementById("authAgreementScroll");
const authAgreementPages = document.getElementById("authAgreementPages");
const authAgreementLens = document.getElementById("authAgreementLens");
const authAgreementLensInner = document.getElementById("authAgreementLensInner");
const authAgreementAccept = document.getElementById("authAgreementAccept");
const authAgreementStatus = document.getElementById("authAgreementStatus");
const authInviteStatus = document.getElementById("authInviteStatus");
const authError = document.getElementById("authError");
const authSignIn = document.getElementById("authSignIn");
const authSignUp = document.getElementById("authSignUp");
const authRegisterBack = document.getElementById("authRegisterBack");
const authForgotPassword = document.getElementById("authForgotPassword");
const authLogoLottie = document.getElementById("authLogoLottie");
const authBrand = document.getElementById("authBrand");
const authBrandOriginal = document.getElementById("authBrandOriginal");
const authBrandInlineLogo = document.getElementById("authBrandInlineLogo");
const appLogoLottie = document.getElementById("appLogoLottie");
const flowLogoAnimationDataPromise =
  typeof window !== "undefined" && typeof window.fetch === "function"
    ? window
        .fetch(FLOW_LOGO_JSON_URL, { cache: "force-cache" })
        .then((response) => (response.ok ? response.json() : null))
        .catch(() => null)
    : Promise.resolve(null);
const accountChip = document.getElementById("accountChip");
const signOutButton = document.getElementById("signOutButton");
const openAccountPageButton = document.getElementById("openAccountPage");
const openBuilderPageButton = document.getElementById("openBuilderPage");
const openAdminPageButton = document.getElementById("openAdminPage");
const openLeadsPageButton = document.getElementById("openLeadsPage");
const openHistoryPageButton = document.getElementById("openHistoryPage");
const openReportsPageButton = document.getElementById("openReportsPage");
const closeAccountPageButton = document.getElementById("closeAccountPage");
const closeAdminPageButton = document.getElementById("closeAdminPage");
const closeLeadsPageButton = document.getElementById("closeLeadsPage");
const closeHistoryPageButton = document.getElementById("closeHistoryPage");
const closeReportsPageButton = document.getElementById("closeReportsPage");
const openAdminSettingsModalButton = document.getElementById("openAdminSettingsModal");
const builderPage = document.getElementById("builderPage");
const accountPageSection = document.getElementById("accountPageSection");
const adminPageSection = document.getElementById("adminPageSection");
const leadsPageSection = document.getElementById("leadsPageSection");
const historyPageSection = document.getElementById("historyPageSection");
const reportsPageSection = document.getElementById("reportsPageSection");
const accountHistoryPanel = historyPageSection?.querySelector(".account-history-panel") || null;
const accountPreviewPanel = historyPageSection?.querySelector(".account-preview-panel") || null;
const accountCompanyName = document.getElementById("accountCompanyName");
const accountContactName = document.getElementById("accountContactName");
const accountContactEmail = document.getElementById("accountContactEmail");
const accountContactPhone = document.getElementById("accountContactPhone");
const accountBillingAddress = document.getElementById("accountBillingAddress");
const accountTaxId = document.getElementById("accountTaxId");
const accountCustomerId = document.getElementById("accountCustomerId");
const accountAccountManager = document.getElementById("accountAccountManager");
const languageSelect = document.getElementById("languageSelect");
const warehouseStatus = document.getElementById("warehouseStatus");
const warehouseStatusRow = warehouseStatus?.closest(".warehouse-status-row") || null;
const warehouseList = document.getElementById("warehouseList");
const warehouseAddButton = document.getElementById("warehouseAdd");
const warehouseSaveButton = document.getElementById("warehouseSave");
const clientInviteEmailInput = document.getElementById("clientInviteEmail");
const clientInviteExpirySelect = document.getElementById("clientInviteExpiry");
const clientInviteCreateButton = document.getElementById("clientInviteCreateButton");
const clientInviteResult = document.getElementById("clientInviteResult");
const clientInviteUrlInput = document.getElementById("clientInviteUrl");
const clientInviteCopyButton = document.getElementById("clientInviteCopyButton");
const clientInviteStatus = document.getElementById("clientInviteStatus");
const clientInviteHistoryEmpty = document.getElementById("clientInviteHistoryEmpty");
const clientInviteHistoryList = document.getElementById("clientInviteHistoryList");
const clientInviteResultEmail = document.getElementById("clientInviteResultEmail");
const clientInviteResultExpiry = document.getElementById("clientInviteResultExpiry");
const openClientInviteHistoryModalButton = document.getElementById("openClientInviteHistoryModal");
const clientInviteHistoryModal = document.getElementById("clientInviteHistoryModal");
const clientInviteHistoryClose = document.getElementById("clientInviteHistoryClose");
const clientInviteHistoryCancel = document.getElementById("clientInviteHistoryCancel");
const adminSummaryClients = document.getElementById("adminSummaryClients");
const adminSummaryActiveClients = document.getElementById("adminSummaryActiveClients");
const adminSummaryInvites = document.getElementById("adminSummaryInvites");
const adminSummaryRevenue = document.getElementById("adminSummaryRevenue");
const adminSummaryProfit = document.getElementById("adminSummaryProfit");
const adminSummaryProfitMeta = document.getElementById("adminSummaryProfitMeta");
const adminSettingsModal = document.getElementById("adminSettingsModal");
const adminSettingsClose = document.getElementById("adminSettingsClose");
const adminSettingsCancel = document.getElementById("adminSettingsCancel");
const adminCarrierDiscountInput = document.getElementById("adminCarrierDiscount");
const adminClientDiscountInput = document.getElementById("adminClientDiscount");
const adminSettingsPreview = document.getElementById("adminSettingsPreview");
const adminSettingsSaveButton = document.getElementById("adminSettingsSaveButton");
const leadsReadyCount = document.getElementById("leadsReadyCount");
const leadsFollowUpCount = document.getElementById("leadsFollowUpCount");
const leadsNotInterestedCount = document.getElementById("leadsNotInterestedCount");
const leadsSearchInput = document.getElementById("leadsSearchInput");
const leadsStackFilter = document.getElementById("leadsStackFilter");
const reloadLeadsButton = document.getElementById("reloadLeadsButton");
const leadsStatus = document.getElementById("leadsStatus");
const leadsEmpty = document.getElementById("leadsEmpty");
const leadsTableWrap = document.getElementById("leadsTableWrap");
const leadsTableBody = document.getElementById("leadsTableBody");
const leadCallOutcomeModal = document.getElementById("leadCallOutcomeModal");
const leadCallOutcomeClose = document.getElementById("leadCallOutcomeClose");
const leadCallOutcomeCancel = document.getElementById("leadCallOutcomeCancel");
const leadCallOutcomeLeadName = document.getElementById("leadCallOutcomeLeadName");
const leadCallOutcomeLeadCompany = document.getElementById("leadCallOutcomeLeadCompany");
const leadCallOutcomeLeadPhone = document.getElementById("leadCallOutcomeLeadPhone");
const leadCallOutcomeActions = document.getElementById("leadCallOutcomeActions");
const adminClientSearchInput = document.getElementById("adminClientSearch");
const adminClientFilterSelect = document.getElementById("adminClientFilter");
const adminClientSortSelect = document.getElementById("adminClientSort");
const adminMockDataButton = document.getElementById("adminMockDataButton");
const adminBillingRunPreviewButton = document.getElementById("adminBillingRunPreview");
const adminBillingRunCreateButton = document.getElementById("adminBillingRunCreate");
const adminBillingRunSendButton = document.getElementById("adminBillingRunSend");
const adminBillingRunStatus = document.getElementById("adminBillingRunStatus");
const adminBillingRunResult = document.getElementById("adminBillingRunResult");
const adminBillingTestEmailInput = document.getElementById("adminBillingTestEmail");
const adminBillingSendTestButton = document.getElementById("adminBillingSendTest");
const adminBillingSendTestSequenceButton = document.getElementById("adminBillingSendTestSequence");
const adminReportsSendTestButton = document.getElementById("adminReportsSendTest");
const adminAgreementPreviewButton = document.getElementById("adminAgreementPreview");
const adminAgreementSendTestButton = document.getElementById("adminAgreementSendTest");
const adminInvoiceList = document.getElementById("adminInvoiceList");
const adminInvoiceEmpty = document.getElementById("adminInvoiceEmpty");
const adminClientsEmpty = document.getElementById("adminClientsEmpty");
const adminClientsList = document.getElementById("adminClientsList");
const accountHistoryStatus = document.getElementById("accountHistoryStatus");
const accountHistoryList = document.getElementById("accountHistoryList");
const accountPreviewMeta = document.getElementById("accountPreviewMeta");
const accountDownloadPdf = document.getElementById("accountDownloadPdf");
const accountDownloadInvoice = document.getElementById("accountDownloadInvoice");
const openReceiptModalButton = document.getElementById("openReceiptModal");
const accountBatchPanel = document.getElementById("accountBatchPanel");
const accountBatchList = document.getElementById("accountBatchList");
const accountBatchPreview = document.getElementById("accountBatchPreview");
const accountPdfFrame = document.getElementById("accountPdfFrame");
const receiptModal = document.getElementById("receiptModal");
const receiptModalClose = document.getElementById("receiptModalClose");
const receiptDocument = document.getElementById("receiptDocument");
const receiptDownloadPdf = document.getElementById("receiptDownloadPdf");
const toastStack = document.getElementById("toastStack");
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
const step3Title = document.querySelector('.stepper-item[data-step="3"] .step-title');
const step3Sub = document.querySelector('.stepper-item[data-step="3"] .step-sub');
const labelCards = document.querySelectorAll(".label-card");
const summaryService = document.getElementById("summaryService");
const summaryPrice = document.getElementById("summaryPrice");
const summaryQty = document.getElementById("summaryQty");
const summaryTotal = document.getElementById("summaryTotal");
const summaryTracking = document.getElementById("summaryTracking");
const summaryChatButton = document.getElementById("summaryChatButton");
const checkoutStepTitle = document.getElementById("checkoutStepTitle");
const checkoutStepSubtitle = document.getElementById("checkoutStepSubtitle");
const invoiceReviewWrap = document.getElementById("invoiceReviewWrap");
const directPaymentWrap = document.getElementById("directPaymentWrap");
const paymentMethodList = document.getElementById("paymentMethodList");
const paymentMethodWallet = document.getElementById("paymentMethodWallet");
const paymentMethodCard = document.getElementById("paymentMethodCard");
const paymentCardExpand = document.getElementById("paymentCardExpand");
const paymentCardBrandPill = document.getElementById("paymentCardBrandPill");
const cardNumberInput = document.getElementById("cardNumberInput");
const cardExpiryInput = document.getElementById("cardExpiryInput");
const cardCvcInput = document.getElementById("cardCvcInput");
const cardHolderInput = document.getElementById("cardHolderInput");
const paymentMethodError = document.getElementById("paymentMethodError");
const paymentService = document.getElementById("paymentService");
const paymentTotal = document.getElementById("paymentTotal");
const directPaymentTotal = document.getElementById("directPaymentTotal");
const invoiceQty = document.getElementById("invoiceQty");
const invoiceSubtotal = document.getElementById("invoiceSubtotal");
const invoiceVat = document.getElementById("invoiceVat");
const invoiceCompany = document.getElementById("invoiceCompany");
const invoiceContact = document.getElementById("invoiceContact");
const invoiceEmail = document.getElementById("invoiceEmail");
const invoicePhone = document.getElementById("invoicePhone");
const invoiceAddress = document.getElementById("invoiceAddress");
const invoiceTaxId = document.getElementById("invoiceTaxId");
const walletAvailableInline = document.getElementById("walletAvailableInline");
const walletBalanceValue = document.getElementById("walletBalanceValue");
const walletPendingValue = document.getElementById("walletPendingValue");
const directPaymentHint = document.getElementById("directPaymentHint");
const payButtonLabel = document.getElementById("payButtonLabel");
const openIbanTopupFromStep = document.getElementById("openIbanTopupFromStep");
const quantityInput = document.getElementById("labelQuantity");
const batchPanel = document.getElementById("batchPanel");
const batchList = document.getElementById("batchList");
const batchPreview = document.getElementById("batchPreview");
const adminOnlyTestTools = document.getElementById("adminOnlyTestTools");
const autoCsvButton = document.getElementById("autoCsv");
const csvEditToggle = document.getElementById("csvEditToggle");
const csvSection = document.getElementById("csvSection");
const csvTableBody = document.querySelector("#csvTable tbody");
const csvPagePrev = document.getElementById("csvPagePrev");
const csvPageNext = document.getElementById("csvPageNext");
const csvPageMeta = document.getElementById("csvPageMeta");
const csvShipFromCard = document.getElementById("csvShipFromCard");
const csvShipFromNote = document.getElementById("csvShipFromNote");
const csvShipFromSelector = document.getElementById("csvShipFromSelector");
const labelForm = document.getElementById("labelForm");
const senderOriginSelector = document.getElementById("senderOriginSelector");
const senderOriginNote = document.getElementById("senderOriginNote");
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
const openIbanTopupFromAccount = document.getElementById("openIbanTopupFromAccount");
const openWalletHistoryFromAccount = document.getElementById("openWalletHistoryFromAccount");
const accountWalletBalance = document.getElementById("accountWalletBalance");
const accountPendingTopups = document.getElementById("accountPendingTopups");
const accountReferenceHistoryList = document.getElementById("accountReferenceHistoryList");
const ibanTopupModal = document.getElementById("ibanTopupModal");
const ibanTopupClose = document.getElementById("ibanTopupClose");
const ibanTopupCancel = document.getElementById("ibanTopupCancel");
const ibanTopupRequest = document.getElementById("ibanTopupRequest");
const ibanTopupStatus = document.getElementById("ibanTopupStatus");
const ibanTopupLoading = document.getElementById("ibanTopupLoading");
const ibanTopupLoadingTitle = document.getElementById("ibanTopupLoadingTitle");
const ibanTopupLoadingSub = document.getElementById("ibanTopupLoadingSub");
const ibanTopupResult = document.getElementById("ibanTopupResult");
const ibanResultBeneficiary = document.getElementById("ibanResultBeneficiary");
const ibanResultIban = document.getElementById("ibanResultIban");
const ibanResultBic = document.getElementById("ibanResultBic");
const ibanResultReference = document.getElementById("ibanResultReference");
const ibanResultEta = document.getElementById("ibanResultEta");
const ibanResultNote = document.getElementById("ibanResultNote");
const ibanCopyReference = document.getElementById("ibanCopyReference");
const ibanCopyIban = document.getElementById("ibanCopyIban");
const ibanTopupModalNote = document.getElementById("ibanTopupModalNote");
const walletHistoryModal = document.getElementById("walletHistoryModal");
const walletHistoryClose = document.getElementById("walletHistoryClose");
const walletHistoryCancel = document.getElementById("walletHistoryCancel");

// Provider dropdown
const providerDropdown = document.getElementById("providerDropdown");
const providerTrigger = document.getElementById("providerTrigger");
const providerMenu = document.getElementById("providerMenu");
const providerStatus = document.getElementById("providerStatus");
const shopifyProviderOption = document.querySelector('.provider-option[data-provider="shopify"]');
const woocommerceSettingsTrigger = document.getElementById("woocommerceSettingsTrigger");
const wixSettingsTrigger = document.getElementById("wixSettingsTrigger");
const shopifySettingsTrigger = document.getElementById("shopifySettingsTrigger");
const providerSettingsModal = document.getElementById("providerSettingsModal");
const providerSettingsTitle = document.getElementById("providerSettingsTitle");
const providerSettingsName = document.getElementById("providerSettingsName");
const providerSettingsNote = document.getElementById("providerSettingsNote");
const providerSettingsClose = document.getElementById("providerSettingsClose");
const providerSettingsCancel = document.getElementById("providerSettingsCancel");
const shopifySettingsModal = document.getElementById("shopifySettingsModal");
const shopifySettingsClose = document.getElementById("shopifySettingsClose");
const shopifySettingsCancel = document.getElementById("shopifySettingsCancel");
const shopifySettingsSave = document.getElementById("shopifySettingsSave");
const shopifySettingsStatus = document.getElementById("shopifySettingsStatus");
const shopifyLocationsSummary = document.getElementById("shopifyLocationsSummary");
const shopifyLocationsList = document.getElementById("shopifyLocationsList");

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
const defaultLabelErrorMessage = String(labelError?.textContent || "").trim()
  || "Complete all required fields or use Auto-generate CSV.";
const defaultPaymentMethodErrorMessage = String(paymentMethodError?.textContent || "").trim()
  || "Select a payment method to continue.";

[
  authAgreementStatus,
  authInviteStatus,
  authError,
  providerStatus,
  labelError,
  customsError,
  paymentMethodError,
  warehouseStatus,
  clientInviteStatus,
  adminBillingRunStatus,
  accountHistoryStatus,
  shopifySettingsStatus,
  ibanTopupStatus,
  csvMapError,
].forEach((element) => {
  if (element) {
    element.hidden = true;
  }
});

if (warehouseStatusRow) {
  warehouseStatusRow.hidden = true;
}

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
let cachedAuthAccessToken = "";
let activeLanguage = "en";
let authMode = "login";
let authInviteToken = "";
let authInviteData = null;
let authInviteValidationToken = 0;
let authInviteIsValid = false;
let authIsBusy = false;
let authAgreementContract = null;
let authAgreementHasReachedEnd = false;
let authAgreementAccepted = false;
let authAgreementScrolledAt = "";
let authAgreementAgreedAt = "";
let authAgreementMaxProgress = 0;
let authAgreementPdfRenderToken = 0;
let authAgreementRendering = false;
let authRegisterStep = 1;
let authRegisterStepTransitionToken = 0;
let authAgreementMagnifierPage = null;
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
let authBackgroundStarted = false;
let authBackgroundVariant = "matrix";
let authBlastBackground = null;
let authBlastBackgroundLoading = null;
let currentMainView = "builder";
let adminAccessAllowed = false;
let adminAccessStatusPromise = null;
let adminAccessStatusRequestId = 0;
let adminDashboardLoading = false;
let adminDashboardLoaded = false;
let adminDashboardState = null;
let adminClients = [];
let leadProspects = [];
let leadProspectsLoading = false;
let leadProspectsLoaded = false;
let leadProspectsLoadPromise = null;
let leadProspectsLoadRequestToken = 0;
let leadSearchQuery = "";
let leadStackFilterValue = "all";
let leadCallOutcomeLeadId = "";
let leadCallOutcomeSaving = false;
let adminClientSearch = "";
let adminClientFilter = "all";
let adminClientSort = "recent";
let adminBillingBusy = false;
let adminBillingInvoices = [];
let adminSettingsDraft = {
  carrier_discount_pct: 25,
  client_discount_pct: 20,
};
let adminSettingsSaved = {
  carrier_discount_pct: 25,
  client_discount_pct: 20,
};
let adminInviteActionBusyIds = new Set();
let adminClientBillingBusyIds = new Set();
let adminMockModeEnabled = false;
let adminMockSnapshot = null;
let authShellTransitionToken = 0;
let mainViewTransitionToken = 0;
let historyPanelSyncFrame = 0;
let csvMappingDraft = null;
let csvModalStepState = "upload";
let csvModalStepTransitionToken = 0;
let reportsRangeDays = 30;
let reportsRangeMode = "all";
let reportsCustomStart = "";
let reportsCustomEnd = "";
let reportsActiveServiceIndex = -1;
let reportsActiveCountryKey = "";
let reportsActiveRegionKey = "";
let reportsDomesticGeoJson = null;
let reportsWorldGeoJson = null;
let reportsGeoLoadPromise = null;
let historyLoadRequestToken = 0;
let billingOverview = null;
let checkoutPaymentMethod = "invoice";
const IBAN_TOPUP_REFERENCE_REUSE_MAX_AGE_MS = 14 * 24 * 60 * 60 * 1000;
let ibanTopupDraft = null;
let ibanTopupRequestInFlight = false;
let ibanTopupRequestPromise = null;
let customsGhostVisible = false;
let providerStatusTimer = 0;
let shopifyConnection = null;
let shopifyLocationsCache = [];
let shopifyLocationDraftSelection = new Set();
let shopifySavedLocationSelection = [];
let shopifySettingsBusy = false;
let clientInviteBusy = false;
let clientInviteHistory = [];
let authKeepAliveTimer = 0;
const translationTextNodeBase = new WeakMap();
const translationAttrBase = new WeakMap();
const AUTH_AGREEMENT_LENS_WIDTH = 214;
const AUTH_AGREEMENT_LENS_HEIGHT = 132;
const AUTH_AGREEMENT_LENS_ZOOM = 2.35;

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

function getInviteTokenFromLocation(location = window.location) {
  const params = new URLSearchParams(location.search || "");
  const queryToken = String(params.get("invite") || params.get("token") || "").trim();
  if (queryToken) return queryToken;
  const relativePath = getRelativeRoutePath(location.pathname || "/");
  if (relativePath.startsWith(`${ROUTE_PATHS.register}/`)) {
    return decodeURIComponent(relativePath.slice(ROUTE_PATHS.register.length + 1)).trim();
  }
  return "";
}

function isSignupPreviewRoute(location = window.location) {
  const params = new URLSearchParams(location.search || "");
  const previewFlag = String(params.get("signupPreview") || "").trim().toLowerCase();
  if (["1", "true", "yes"].includes(previewFlag)) {
    return true;
  }
  return getRelativeRoutePath(location.pathname || "/") === ROUTE_PATHS.signupPreview;
}

function getSignupPreviewStepFromLocation(location = window.location) {
  const params = new URLSearchParams(location.search || "");
  return Number(params.get("previewStep") || params.get("step") || "") === 2 ? 2 : 1;
}

function getReportRangeFromLocation(location = window.location) {
  const params = new URLSearchParams(location.search || "");
  return String(params.get("range") || params.get("reportRange") || "")
    .trim()
    .toLowerCase();
}

function applyReportRangeFromToken(token = "") {
  const normalized = String(token || "").trim().toLowerCase();
  if (!normalized) return;
  if (["month", "monthly", "30d", "30", "last30"].includes(normalized)) {
    reportsRangeMode = "preset";
    reportsRangeDays = 30;
    setReportsCustomRangeVisible(false);
    return;
  }
  if (["all", "alltime", "all-time"].includes(normalized)) {
    reportsRangeMode = "all";
    setReportsCustomRangeVisible(false);
  }
}

function parseRouteFromLocation() {
  const path = getRelativeRoutePath(window.location.pathname);
  if (isSignupPreviewRoute(window.location)) {
    return {
      view: "register",
      inviteToken: AUTH_SIGNUP_PREVIEW_TOKEN,
      previewRegistration: true,
      previewStep: getSignupPreviewStepFromLocation(window.location),
    };
  }
  if (path === ROUTE_PATHS.login) {
    return { view: "login" };
  }
  if (path === ROUTE_PATHS.register || path.startsWith(`${ROUTE_PATHS.register}/`)) {
    return { view: "register", inviteToken: getInviteTokenFromLocation(window.location) };
  }
  if (path === ROUTE_PATHS.account) {
    return { view: "account" };
  }
  if (path === ROUTE_PATHS.admin) {
    return { view: "admin" };
  }
  if (path === ROUTE_PATHS.leads) {
    return { view: "leads" };
  }
  if (path === ROUTE_PATHS.history) {
    return { view: "history" };
  }
  if (path === ROUTE_PATHS.reports) {
    const reportRangeFromQuery = getReportRangeFromLocation(window.location);
    const reportRangeFromState = String(history.state?.reportRange || "").trim().toLowerCase();
    return { view: "reports", reportRange: reportRangeFromQuery || reportRangeFromState };
  }

  const stepEntry = Object.entries(STEP_ROUTE_PATHS).find(([, routePath]) => routePath === path);
  if (stepEntry) {
    return { view: "builder", step: clampStep(Number(stepEntry[0])) };
  }

  const hash = window.location.hash || "";
  if (hash === "#account") {
    return { view: "account" };
  }
  if (hash === "#admin") {
    return { view: "admin" };
  }
  if (hash === "#leads") {
    return { view: "leads" };
  }
  if (hash === "#history") {
    return { view: "history" };
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
  if (route.view === "register") {
    if (route.previewRegistration) {
      const previewStep = Number(route?.previewStep) === 2 ? 2 : 1;
      const previewPath = buildRoutePath(ROUTE_PATHS.signupPreview);
      return previewStep === 2 ? `${previewPath}?step=2` : previewPath;
    }
    return buildRoutePath(ROUTE_PATHS.register);
  }
  if (route.view === "account") {
    return buildRoutePath(ROUTE_PATHS.account);
  }
  if (route.view === "admin") {
    return buildRoutePath(ROUTE_PATHS.admin);
  }
  if (route.view === "leads") {
    return buildRoutePath(ROUTE_PATHS.leads);
  }
  if (route.view === "history") {
    return buildRoutePath(ROUTE_PATHS.history);
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
  if (route.view === "register") {
    return {
      view: "register",
      inviteToken: String(route?.inviteToken || "").trim(),
      previewRegistration: Boolean(route?.previewRegistration),
      previewStep: Number(route?.previewStep) === 2 ? 2 : 1,
    };
  }
  if (route.view === "reports") {
    return { view: "reports", reportRange: String(route?.reportRange || "").trim() };
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
    Boolean(history.state?.customs) === Boolean(nextState.customs) &&
    String(history.state?.inviteToken || "") === String(nextState?.inviteToken || "") &&
    Boolean(history.state?.previewRegistration) === Boolean(nextState?.previewRegistration) &&
    Number(history.state?.previewStep || 1) === Number(nextState?.previewStep || 1) &&
    String(history.state?.reportRange || "") === String(nextState?.reportRange || "");

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
        throw new Error(
          tr("Failed to load domestic regions GeoJSON ({status})", { status: response.status })
        );
      }
      return response.json();
    }),
    fetch(REPORT_WORLD_GEOJSON_URL).then((response) => {
      if (!response.ok) {
        throw new Error(
          tr("Failed to load world GeoJSON ({status})", { status: response.status })
        );
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

function normalizeLanguageCode(value) {
  const code = String(value || "").trim().toLowerCase();
  if (SUPPORTED_LANGUAGES.has(code)) return code;
  if (code.startsWith("fr")) return "fr";
  if (code.startsWith("nl")) return "nl";
  if (code.startsWith("en")) return "en";
  return "en";
}

function setTextInputCompositionState(target, composing) {
  if (!(target instanceof HTMLElement)) return;
  if (composing) {
    target.dataset.composing = "true";
  } else {
    delete target.dataset.composing;
  }
}

function isTextInputComposing(event, target = event?.target) {
  return Boolean(
    event?.isComposing ||
      (target instanceof HTMLElement && target.dataset.composing === "true")
  );
}

function bindCompositionAwareInput(input, handler) {
  if (!input || typeof handler !== "function") return;

  let skipNextInput = false;

  input.addEventListener("compositionstart", () => {
    setTextInputCompositionState(input, true);
  });

  input.addEventListener("compositionend", (event) => {
    setTextInputCompositionState(input, false);
    skipNextInput = true;
    handler(event);
    window.setTimeout(() => {
      skipNextInput = false;
    }, 0);
  });

  input.addEventListener("input", (event) => {
    if (skipNextInput) {
      skipNextInput = false;
      return;
    }
    if (isTextInputComposing(event, input)) {
      return;
    }
    handler(event);
  });
}

function bindDelegatedCompositionAwareInput(container, selector, handler) {
  if (!container || !selector || typeof handler !== "function") return;

  const skipNextInputs = new WeakSet();

  const resolveTarget = (event) => {
    const rawTarget = event.target;
    if (!(rawTarget instanceof Element)) return null;
    const target = rawTarget.closest(selector);
    if (!(target instanceof HTMLElement) || !container.contains(target)) {
      return null;
    }
    return target;
  };

  container.addEventListener("compositionstart", (event) => {
    const target = resolveTarget(event);
    if (!target) return;
    setTextInputCompositionState(target, true);
  });

  container.addEventListener("compositionend", (event) => {
    const target = resolveTarget(event);
    if (!target) return;
    setTextInputCompositionState(target, false);
    skipNextInputs.add(target);
    handler(target, event);
    window.setTimeout(() => {
      skipNextInputs.delete(target);
    }, 0);
  });

  container.addEventListener("input", (event) => {
    const target = resolveTarget(event);
    if (!target) return;
    if (skipNextInputs.has(target)) {
      skipNextInputs.delete(target);
      return;
    }
    if (isTextInputComposing(event, target)) {
      return;
    }
    handler(target, event);
  });
}

function getUiLocale() {
  return LANGUAGE_LOCALE[activeLanguage] || LANGUAGE_LOCALE.en;
}

function normalizeTranslationKey(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim();
}

function interpolateTranslation(template, vars = {}) {
  return String(template || "").replace(/\{(\w+)\}/g, (_match, key) => {
    if (vars && Object.prototype.hasOwnProperty.call(vars, key)) {
      return String(vars[key]);
    }
    return `{${key}}`;
  });
}

function tr(baseText, vars = {}) {
  const baseKey = normalizeTranslationKey(baseText);
  if (!baseKey) return "";
  const entry = TRANSLATIONS[baseKey];
  const template =
    activeLanguage === "en"
      ? baseKey
      : entry && typeof entry === "object" && entry[activeLanguage]
        ? entry[activeLanguage]
        : baseKey;
  return interpolateTranslation(template, vars);
}

function translateServiceName(service) {
  const normalized = String(service || "").trim();
  if (!normalized) return tr("Label generation");
  return tr(normalized);
}

function translateDomesticRegionName(name) {
  const normalized = normalizeNameKey(name);
  if (normalized === normalizeNameKey("North Region")) return tr("North Region");
  if (normalized === normalizeNameKey("South Region")) return tr("South Region");
  if (normalized === normalizeNameKey("Capital Region")) return tr("Capital Region");
  return name;
}

function applyStaticTranslations(root = document.body) {
  if (!root) return;
  const walker = document.createTreeWalker(
    root,
    NodeFilter.SHOW_TEXT,
    {
      acceptNode(node) {
        if (!node || !node.nodeValue || !node.parentElement) {
          return NodeFilter.FILTER_REJECT;
        }
        const parentTag = node.parentElement.tagName;
        if (parentTag === "SCRIPT" || parentTag === "STYLE") {
          return NodeFilter.FILTER_REJECT;
        }
        return NodeFilter.FILTER_ACCEPT;
      },
    },
    false
  );

  let textNode = walker.nextNode();
  while (textNode) {
    const baseValue = translationTextNodeBase.has(textNode)
      ? translationTextNodeBase.get(textNode)
      : textNode.nodeValue;
    if (!translationTextNodeBase.has(textNode)) {
      translationTextNodeBase.set(textNode, baseValue);
    }
    const source = String(baseValue || "");
    const leading = source.match(/^\s*/)?.[0] || "";
    const trailing = source.match(/\s*$/)?.[0] || "";
    const core = source.trim();
    if (core) {
      textNode.nodeValue = `${leading}${tr(core)}${trailing}`;
    }
    textNode = walker.nextNode();
  }

  const elements = root.querySelectorAll("*");
  elements.forEach((element) => {
    let attrCache = translationAttrBase.get(element);
    if (!attrCache) {
      attrCache = {};
      translationAttrBase.set(element, attrCache);
    }
    TRANSLATION_ATTRS.forEach((attrName) => {
      if (!element.hasAttribute(attrName)) return;
      if (!Object.prototype.hasOwnProperty.call(attrCache, attrName)) {
        attrCache[attrName] = element.getAttribute(attrName) || "";
      }
      const source = String(attrCache[attrName] || "");
      if (!source.trim()) return;
      element.setAttribute(attrName, tr(source));
    });
  });
}

function getStoredLanguagePreference() {
  try {
    return normalizeLanguageCode(localStorage.getItem(LANGUAGE_STORAGE_KEY));
  } catch (_error) {
    return "en";
  }
}

function setStoredLanguagePreference(language) {
  try {
    localStorage.setItem(LANGUAGE_STORAGE_KEY, normalizeLanguageCode(language));
  } catch (_error) {
    // Ignore storage write failures.
  }
}

async function saveLanguagePreference(language) {
  if (!supabaseClient || !currentUser?.id) return;
  const nextLanguage = normalizeLanguageCode(language);
  const existing = normalizeLanguageCode(currentUser?.user_metadata?.preferred_language);
  if (existing === nextLanguage) return;
  const metadata = { ...(currentUser?.user_metadata || {}), preferred_language: nextLanguage };
  const { data, error } = await supabaseClient.auth.updateUser({ data: metadata });
  if (!error && data?.user) {
    currentUser = data.user;
  }
}

function resolvePreferredLanguage(user) {
  const userPref = normalizeLanguageCode(user?.user_metadata?.preferred_language);
  if (SUPPORTED_LANGUAGES.has(userPref)) return userPref;
  const localPref = getStoredLanguagePreference();
  if (SUPPORTED_LANGUAGES.has(localPref)) return localPref;
  const browserPref = normalizeLanguageCode(navigator.language || "en");
  return browserPref;
}

function refreshTranslatedRuntime() {
  updateSummary();
  if (paymentService && state?.selection?.type) {
    paymentService.textContent = translateServiceName(state.selection.type);
  }
  if (previewService && state?.selection?.type) {
    previewService.textContent = translateServiceName(state.selection.type);
  }
  renderSenderOriginSelector();
  renderCsvShipFromSelector();
  renderCsvTable();
  renderWarehouseList();
  renderCustomsItems();
  updateCustomsScopeMeta();
  renderAccountHistoryList();
  renderAccountBatchList();
  if (accountActiveRecord) {
    renderReceiptDetails(accountActiveRecord);
  }
  renderReportsDashboard();
  renderShopifySettingsLocations();
  renderClientInviteHistory(clientInviteHistory);
  renderAdminSummary(adminDashboardState?.summary || {});
  renderAdminSettingsPreview();
  renderAdminMockDataButton();
  renderAdminClientsList();
  if (csvMappingDraft) {
    renderCsvMappingTable();
  }
}

async function setLanguage(language, options = {}) {
  const { persist = false } = options;
  const nextLanguage = normalizeLanguageCode(language);
  if (nextLanguage === activeLanguage && !persist) {
    if (languageSelect) {
      languageSelect.value = nextLanguage;
    }
    return;
  }
  activeLanguage = nextLanguage;
  document.documentElement.setAttribute("lang", nextLanguage);
  setStoredLanguagePreference(nextLanguage);
  if (languageSelect) {
    languageSelect.value = nextLanguage;
  }
  applyStaticTranslations(document.body);
  refreshTranslatedRuntime();
  if (persist) {
    await saveLanguagePreference(nextLanguage);
  }
}

function showToast(message, options = {}) {
  if (!toastStack) return;
  const text = String(message || "").trim();
  if (!text) return;
  const rawTone = String(options?.tone || "info").trim().toLowerCase();
  const tone = rawTone === "success" ? "success" : rawTone === "error" || rawTone === "warning" ? "error" : "";
  if (!tone) return;
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
  toastStack.prepend(toast);
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

let lastStatusToast = {
  message: "",
  tone: "",
  title: "",
  at: 0,
};

function showStatusToast(message, options = {}) {
  const text = String(message || "").trim();
  const title = String(options?.title || "").trim();
  const tone = String(options?.tone || "info").trim() || "info";
  if (!text) return;
  const now = Date.now();
  if (
    lastStatusToast.message === text
    && lastStatusToast.tone === tone
    && lastStatusToast.title === title
    && now - lastStatusToast.at < 900
  ) {
    return;
  }
  lastStatusToast = {
    message: text,
    tone,
    title,
    at: now,
  };
  showToast(text, { ...options, title, tone });
}

function setInlineFormErrorToast(errorEl, message = "", tone = "error") {
  const text = String(message || "").trim();
  if (errorEl) {
    errorEl.hidden = true;
    errorEl.textContent = "";
    errorEl.classList.remove("is-visible", "is-error", "is-success", "is-info");
  }
  if (text) {
    showStatusToast(text, { tone });
  }
}

function updateToastStackLayout() {
  if (!toastStack) return;
  const toasts = Array.from(toastStack.querySelectorAll(".toast"));
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

function getToastIconSvg(tone) {
  if (tone === "success") {
    return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M8.5 12.4l2.3 2.3 4.7-5.1"/></svg>';
  }
  if (tone === "error") {
    return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3.7 21 19.3H3L12 3.7Z"/><path d="M12 9v4.6"/><circle cx="12" cy="16.9" r="0.8" fill="currentColor" stroke="none"/></svg>';
  }
  return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 10.2v5"/><circle cx="12" cy="7.3" r="0.8" fill="currentColor" stroke="none"/></svg>';
}

function setProviderStatus(message = "", options = {}) {
  const { kind = "info", toast = Boolean(message) } = options;
  const text = String(message || "").trim();
  if (providerStatus) {
    providerStatus.hidden = true;
    providerStatus.textContent = "";
    providerStatus.classList.remove("is-success", "is-error");
  }
  if (providerStatusTimer) {
    window.clearTimeout(providerStatusTimer);
    providerStatusTimer = 0;
  }
  if (toast && text) {
    showStatusToast(text, { tone: kind });
  }
}

function setShopifyProviderConnectedState(isConnected) {
  if (!shopifyProviderOption) return;
  shopifyProviderOption.classList.toggle("is-connected", Boolean(isConnected));
}

function setShopifySettingsModalOpen(open) {
  if (!shopifySettingsModal) return;
  shopifySettingsModal.classList.toggle("is-closed", !open);
}

function setGenericProviderSettingsModalOpen(open) {
  if (!providerSettingsModal) return;
  providerSettingsModal.classList.toggle("is-closed", !open);
}

function openGenericProviderSettings(providerLabel) {
  const label = String(providerLabel || "").trim() || "Provider";
  if (providerSettingsTitle) {
    providerSettingsTitle.textContent = tr("{provider} Settings", { provider: label });
  }
  if (providerSettingsName) {
    providerSettingsName.textContent = label;
  }
  if (providerSettingsNote) {
    providerSettingsNote.textContent = tr("{provider} import and settings are coming soon.", {
      provider: label,
    });
  }
  setGenericProviderSettingsModalOpen(true);
}

function setShopifySettingsStatus(message = "", options = {}) {
  const { kind = "info", toast = Boolean(message) } = options;
  const text = String(message || "").trim();
  if (shopifySettingsStatus) {
    shopifySettingsStatus.hidden = true;
    shopifySettingsStatus.textContent = "";
    shopifySettingsStatus.classList.remove("is-success", "is-error");
  }
  if (toast && text) {
    showStatusToast(text, { tone: kind });
  }
}

function normalizeShopifyLocationId(value) {
  const asString = String(value ?? "").trim();
  if (!asString) return "";
  const numeric = Number(asString);
  if (Number.isFinite(numeric) && numeric > 0) {
    return String(Math.trunc(numeric));
  }
  return asString;
}

function normalizeShopifyLocationIdList(values) {
  if (!Array.isArray(values)) return [];
  const seen = new Set();
  values.forEach((value) => {
    const id = normalizeShopifyLocationId(value);
    if (!id) return;
    seen.add(id);
  });
  return Array.from(seen);
}

async function fetchShopifySavedSettings(shopDomain) {
  const shop = normalizeShopDomain(shopDomain);
  if (!shop) return [];
  const query = new URLSearchParams();
  query.set("shop", shop);
  let data = null;
  try {
    data = await fetchApiWithAuth(`/api/shopify/settings?${query.toString()}`);
  } catch (error) {
    const message = String(error?.message || "");
    if (/api route not found/i.test(message)) {
      data = await fetchApiWithAuth(`/api/shopify/setting?${query.toString()}`);
    } else {
      throw error;
    }
  }
  return normalizeShopifyLocationIdList(data?.settings?.selectedLocationIds);
}

async function saveShopifySavedSettings(shopDomain, selectedLocationIds) {
  const shop = normalizeShopDomain(shopDomain);
  if (!shop) {
    throw new Error(tr("Connect Shopify before saving settings."));
  }
  let data = null;
  const payload = {
    shop,
    selectedLocationIds: normalizeShopifyLocationIdList(selectedLocationIds),
  };
  try {
    data = await fetchApiWithAuth("/api/shopify/settings", {
      method: "POST",
      body: JSON.stringify(payload),
    });
  } catch (error) {
    const message = String(error?.message || "");
    if (/api route not found/i.test(message)) {
      data = await fetchApiWithAuth("/api/shopify/setting", {
        method: "POST",
        body: JSON.stringify(payload),
      });
    } else {
      throw error;
    }
  }
  return normalizeShopifyLocationIdList(data?.settings?.selectedLocationIds);
}

function getShopifySelectedLocationSummary(locations, selectedIds) {
  if (!Array.isArray(locations) || locations.length === 0) {
    return tr("No locations available");
  }
  const selectedSet = new Set(normalizeShopifyLocationIdList(selectedIds));
  if (!selectedSet.size) return tr("{count} selected", { count: 0 });

  const selected = locations.filter((location) => selectedSet.has(location.id));
  if (selected.length === locations.length) {
    return tr("{count} selected (all)", { count: selected.length });
  }
  return tr("{count} selected", { count: selected.length });
}

function setShopifySettingsBusy(isBusy) {
  shopifySettingsBusy = Boolean(isBusy);
  if (shopifySettingsSave) {
    shopifySettingsSave.disabled = shopifySettingsBusy;
    const label = shopifySettingsSave.querySelector("span");
    if (label) {
      label.textContent = shopifySettingsBusy ? tr("Saving...") : tr("Save Settings");
    }
  }
  if (shopifySettingsCancel) {
    shopifySettingsCancel.disabled = shopifySettingsBusy;
  }
  if (shopifySettingsClose) {
    shopifySettingsClose.disabled = shopifySettingsBusy;
  }
}

function getShopifyLocationMeta(location) {
  const city = String(location?.city || "").trim();
  const region = String(location?.region || "").trim();
  const country = String(location?.country || "").trim();
  return [city, region, country].filter(Boolean).join(", ");
}

function renderShopifySettingsLocations() {
  if (!shopifyLocationsList || !shopifyLocationsSummary) return;
  const locations = Array.isArray(shopifyLocationsCache) ? shopifyLocationsCache : [];
  const validIds = new Set(locations.map((location) => location.id));
  const nextSelection = new Set();
  shopifyLocationDraftSelection.forEach((id) => {
    if (validIds.has(id)) {
      nextSelection.add(id);
    }
  });
  shopifyLocationDraftSelection = nextSelection;

  shopifyLocationsSummary.textContent = getShopifySelectedLocationSummary(
    locations,
    Array.from(shopifyLocationDraftSelection)
  );
  shopifyLocationsList.innerHTML = "";

  if (!locations.length) {
    const empty = document.createElement("div");
    empty.className = "shopify-location-empty";
    empty.textContent = tr("No locations available");
    shopifyLocationsList.appendChild(empty);
    return;
  }

  const allRow = document.createElement("button");
  allRow.type = "button";
  const allChecked = shopifyLocationDraftSelection.size === locations.length;
  allRow.className = `shopify-location-row${allChecked ? " is-selected" : ""}`;
  allRow.dataset.role = "all";
  allRow.setAttribute("aria-pressed", allChecked ? "true" : "false");
  allRow.innerHTML = `
      <span class="shopify-location-check" aria-hidden="true"><span class="shopify-location-check-fill"></span></span>
      <div>
        <div class="shopify-location-item-title">${tr("All Locations")}</div>
        <div class="shopify-location-item-meta">${tr("Mirror Shopify location routing")}</div>
      </div>
  `;
  shopifyLocationsList.appendChild(allRow);

  locations.forEach((location) => {
    const row = document.createElement("button");
    row.type = "button";
    const isChecked = shopifyLocationDraftSelection.has(location.id);
    row.className = `shopify-location-row${isChecked ? " is-selected" : ""}`;
    row.dataset.locationId = location.id;
    row.setAttribute("aria-pressed", isChecked ? "true" : "false");
    const meta = getShopifyLocationMeta(location);
    row.innerHTML = `
      <span class="shopify-location-check" aria-hidden="true"><span class="shopify-location-check-fill"></span></span>
      <div>
        <div class="shopify-location-item-title">${location.name}</div>
        <div class="shopify-location-item-meta">${meta || tr("No address details")}</div>
      </div>
    `;
    shopifyLocationsList.appendChild(row);
  });
}

async function fetchShopifyLocations(shopDomain) {
  const shop = normalizeShopDomain(shopDomain);
  if (!shop) {
    throw new Error(tr("Connect Shopify before opening fulfillment settings."));
  }
  const query = new URLSearchParams();
  query.set("shop", shop);
  let data = null;
  try {
    data = await fetchApiWithAuth(`/api/shopify/locations?${query.toString()}`);
  } catch (error) {
    const message = String(error?.message || "");
    if (/api route not found/i.test(message)) {
      try {
        data = await fetchApiWithAuth(`/api/shopify/location?${query.toString()}`);
      } catch (_fallbackError) {
        throw new Error(
          tr(
            "Shopify locations endpoint is not live yet. Deploy the latest API worker (or restart the local Node server) and try again."
          )
        );
      }
    } else {
      throw error;
    }
  }
  const locations = Array.isArray(data?.locations) ? data.locations : [];
  return locations
    .map((location) => ({
      id: normalizeShopifyLocationId(location?.id),
      name: String(location?.name || "").trim(),
      city: String(location?.city || "").trim(),
      region: String(location?.region || "").trim(),
      country: String(location?.country || "").trim(),
    }))
    .filter((location) => location.id && location.name);
}

async function openShopifySettingsModal() {
  if (!currentUser?.id) {
    setProviderStatus(tr("Sign in before configuring Shopify locations."), { kind: "error" });
    return;
  }
  if (!shopifyConnection?.shop) {
    setProviderStatus(tr("Connect Shopify before opening fulfillment settings."), {
      kind: "error",
    });
    return;
  }

  setShopifySettingsModalOpen(true);
  setShopifySettingsBusy(true);
  setShopifySettingsStatus(tr("Loading Shopify locations..."));

  try {
    const [locations, savedSelection] = await Promise.all([
      fetchShopifyLocations(shopifyConnection.shop),
      fetchShopifySavedSettings(shopifyConnection.shop),
    ]);
    shopifyLocationsCache = locations;
    shopifySavedLocationSelection = savedSelection;
    const savedSet = new Set(savedSelection);
    const validSaved = locations
      .map((location) => location.id)
      .filter((id) => savedSet.has(id));

    shopifyLocationDraftSelection = new Set(
      validSaved.length ? validSaved : locations.map((location) => location.id)
    );
    renderShopifySettingsLocations();
    setShopifySettingsStatus("");
  } catch (error) {
    const message = String(error?.message || tr("Could not load Shopify locations."));
    if (
      /connection expired|reconnect shopify|invalid api key or access token|unrecognized login|wrong password/i.test(
        message
      )
    ) {
      shopifyConnection = null;
      updateShopifyProviderStatus();
    }
    shopifyLocationsCache = [];
    shopifyLocationDraftSelection = new Set();
    shopifySavedLocationSelection = [];
    renderShopifySettingsLocations();
    setShopifySettingsStatus(message, { kind: "error" });
  } finally {
    setShopifySettingsBusy(false);
  }
}

function closeShopifySettingsModal() {
  if (shopifySettingsBusy) return;
  setShopifySettingsModalOpen(false);
}

function closeGenericProviderSettingsModal() {
  setGenericProviderSettingsModalOpen(false);
}

async function saveShopifySettings() {
  if (!currentUser?.id || !shopifyConnection?.shop) {
    setShopifySettingsStatus(tr("Connect Shopify first."), { kind: "error" });
    return;
  }
  const selectedIds = normalizeShopifyLocationIdList(Array.from(shopifyLocationDraftSelection));
  if (!selectedIds.length) {
    setShopifySettingsStatus(tr("Select at least one location."), { kind: "error" });
    return;
  }

  setShopifySettingsBusy(true);
  try {
    const savedSelection = await saveShopifySavedSettings(shopifyConnection.shop, selectedIds);
    shopifySavedLocationSelection = savedSelection;
    setProviderStatus(tr("Shopify fulfillment settings updated."), { kind: "success" });
    setShopifySettingsStatus(tr("Settings saved."), { kind: "success" });
    window.setTimeout(() => {
      closeShopifySettingsModal();
    }, 150);
  } catch (error) {
    setShopifySettingsStatus(error?.message || tr("Could not save Shopify settings."), {
      kind: "error",
    });
  } finally {
    setShopifySettingsBusy(false);
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
  try {
    const {
      data: { session },
      error,
    } = await supabaseClient.auth.getSession();
    if (error) {
      throw error;
    }
    const token = String(session?.access_token || "");
    if (token) {
      cachedAuthAccessToken = token;
      return token;
    }
  } catch (_error) {
    // Fallback to cached token for transient session-store hiccups.
  }
  return cachedAuthAccessToken;
}

async function refreshAuthAccessToken() {
  if (!supabaseClient) return "";
  try {
    const { data, error } = await supabaseClient.auth.refreshSession();
    if (error) {
      throw error;
    }
    const token = String(data?.session?.access_token || "");
    if (token) {
      cachedAuthAccessToken = token;
    }
    return token;
  } catch (_error) {
    return "";
  }
}

function stopAuthKeepAlive() {
  if (authKeepAliveTimer) {
    window.clearInterval(authKeepAliveTimer);
    authKeepAliveTimer = 0;
  }
}

function startAuthKeepAlive() {
  stopAuthKeepAlive();
  if (!supabaseClient) return;
  authKeepAliveTimer = window.setInterval(() => {
    if (!currentUser) return;
    void refreshAuthAccessToken();
  }, 4 * 60 * 1000);
}

async function fetchApiWithAuth(path, options = {}) {
  const { timeoutMs = 15000, ...requestOptions } = options;
  const sendWithToken = async (token) => {
    if (!token) {
      throw new Error(tr("You must be signed in."));
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
          tr("Request failed ({status})", { status: response.status });
        const error = new Error(message);
        error.status = response.status;
        throw error;
      }
      return payload;
    } catch (error) {
      if (error?.name === "AbortError") {
        throw new Error(tr("Request timed out. Please try again."));
      }
      throw error;
    } finally {
      window.clearTimeout(timeoutId);
    }
  };

  let token = await getAuthAccessToken();
  try {
    return await sendWithToken(token);
  } catch (error) {
    if (Number(error?.status) !== 401) {
      throw error;
    }
    token = await refreshAuthAccessToken();
    if (!token) {
      cachedAuthAccessToken = "";
      throw new Error(tr("You must be signed in."));
    }
    return sendWithToken(token);
  }
}

async function fetchApiBlobWithAuth(path, options = {}) {
  const { timeoutMs = 15000, ...requestOptions } = options;
  const sendWithToken = async (token) => {
    if (!token) {
      throw new Error(tr("You must be signed in."));
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
      if (!response.ok) {
        const contentType = response.headers.get("content-type") || "";
        const payload = contentType.includes("application/json")
          ? await response.json().catch(() => ({}))
          : await response.text().catch(() => "");
        const message =
          (payload && typeof payload === "object" && payload.error) ||
          (typeof payload === "string" ? payload : "") ||
          tr("Request failed ({status})", { status: response.status });
        const error = new Error(message);
        error.status = response.status;
        throw error;
      }
      return {
        blob: await response.blob(),
        headers: response.headers,
      };
    } catch (error) {
      if (error?.name === "AbortError") {
        throw new Error(tr("Request timed out. Please try again."));
      }
      throw error;
    } finally {
      window.clearTimeout(timeoutId);
    }
  };

  let token = await getAuthAccessToken();
  try {
    return await sendWithToken(token);
  } catch (error) {
    if (Number(error?.status) !== 401) {
      throw error;
    }
    token = await refreshAuthAccessToken();
    if (!token) {
      cachedAuthAccessToken = "";
      throw new Error(tr("You must be signed in."));
    }
    return sendWithToken(token);
  }
}

function isValidEmailFormat(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || "").trim());
}

function isInviteExpired(invite) {
  const expiresAt = String(invite?.expires_at || "").trim();
  if (!expiresAt) return true;
  const ms = Date.parse(expiresAt);
  if (!Number.isFinite(ms)) return true;
  return ms < Date.now();
}

function isInviteRevoked(invite) {
  return Boolean(String(invite?.revoked_at || "").trim());
}

function getClientInviteStatus(invite) {
  if (invite?.claimed_at) return "claimed";
  if (isInviteRevoked(invite)) return "revoked";
  if (isInviteExpired(invite)) return "expired";
  return "open";
}

function formatInviteStatusLabel(status, invite) {
  if (status === "claimed") {
    return invite?.claimed_at
      ? tr("Claimed {date}", { date: formatHistoryDate(invite.claimed_at) })
      : tr("Claimed");
  }
  if (status === "revoked") {
    return invite?.revoked_at
      ? tr("Revoked {date}", { date: formatHistoryDate(invite.revoked_at) })
      : tr("Revoked");
  }
  if (status === "expired") return tr("Expired");
  return tr("Open");
}

function canRevokeInvite(invite) {
  return getClientInviteStatus(invite) === "open";
}

function renderClientInviteHistory(invites = []) {
  if (!clientInviteHistoryList || !clientInviteHistoryEmpty) return;
  clientInviteHistory = Array.isArray(invites) ? invites.slice() : [];
  clientInviteHistoryList.innerHTML = "";
  if (!Array.isArray(invites) || invites.length === 0) {
    clientInviteHistoryEmpty.textContent = tr("No invites yet.");
    clientInviteHistoryEmpty.classList.remove("is-hidden");
    return;
  }

  clientInviteHistoryEmpty.classList.add("is-hidden");
  invites.forEach((invite) => {
    const item = document.createElement("div");
    item.className = "invite-history-item";

    const main = document.createElement("div");
    main.className = "invite-history-main";
    const status = getClientInviteStatus(invite);

    const top = document.createElement("div");
    top.className = "invite-history-top";

    const email = document.createElement("div");
    email.className = "invite-history-email";
    email.textContent = invite?.invited_email ? String(invite.invited_email) : "--";

    const badge = document.createElement("div");
    badge.className = `invite-history-badge is-${status}`;
    badge.textContent = formatInviteStatusLabel(status, invite);
    top.appendChild(email);
    top.appendChild(badge);

    const urlRow = document.createElement("div");
    urlRow.className = "invite-history-url-row";

    const inviteUrl = String(invite?.invite_url || "").trim();
    const urlValue = document.createElement("div");
    urlValue.className = `invite-history-url${inviteUrl ? "" : " is-unavailable"}`;
    urlValue.textContent = inviteUrl || tr("Stored URL unavailable for this invite.");
    urlRow.appendChild(urlValue);

    const copyButton = document.createElement("button");
    copyButton.type = "button";
    copyButton.className = "btn btn-secondary btn-sm";
    copyButton.dataset.inviteCopy = String(invite?.id || "");
    copyButton.disabled = !inviteUrl;
    copyButton.innerHTML = `
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><rect x="9" y="9" width="12" height="12" rx="1"/><rect x="3" y="3" width="12" height="12" rx="1"/></svg>
      <span>${tr("Copy URL")}</span>
    `;
    urlRow.appendChild(copyButton);

    const revokeButton = document.createElement("button");
    revokeButton.type = "button";
    revokeButton.className = "btn btn-ghost btn-sm";
    revokeButton.dataset.inviteRevoke = String(invite?.id || "");
    revokeButton.disabled =
      !canRevokeInvite(invite) || adminInviteActionBusyIds.has(String(invite?.id || ""));
    revokeButton.innerHTML = `
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M4 7h16"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M6 7l1 12a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2l1-12"/><path d="M9 7V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v3"/></svg>
      <span>${adminInviteActionBusyIds.has(String(invite?.id || "")) ? tr("Revoking...") : tr("Revoke")}</span>
    `;
    urlRow.appendChild(revokeButton);

    const meta = document.createElement("div");
    meta.className = "invite-history-meta";
    const createdText = tr("Created {date}", { date: formatHistoryDate(invite?.created_at) });
    const expiresText = tr("Expires {date}", { date: formatHistoryDate(invite?.expires_at) });
    const extraText =
      status === "claimed" && invite?.claimed_email
        ? tr("Claimed by {email}", { email: invite.claimed_email })
        : status === "revoked"
          ? tr("Invite revoked")
          : "";
    meta.textContent = [createdText, expiresText, extraText].filter(Boolean).join(" • ");

    main.appendChild(top);
    main.appendChild(urlRow);
    main.appendChild(meta);

    item.appendChild(main);
    clientInviteHistoryList.appendChild(item);
  });
}

async function loadClientInviteHistory(options = {}) {
  const { quiet = false } = options;
  if (!currentUser) {
    renderClientInviteHistory([]);
    return;
  }
  if (!quiet) {
    renderClientInviteHistory([]);
    if (clientInviteHistoryEmpty) {
      clientInviteHistoryEmpty.textContent = tr("Loading invites...");
      clientInviteHistoryEmpty.classList.remove("is-hidden");
    }
  }
  try {
    const payload = await fetchApiWithAuth("/api/auth/invites?limit=20");
    const invites = Array.isArray(payload?.invites) ? payload.invites : [];
    renderClientInviteHistory(invites);
  } catch (error) {
    showStatusToast(error?.message || tr("Could not load invites."), { tone: "error" });
    renderClientInviteHistory([]);
    if (clientInviteHistoryEmpty) {
      clientInviteHistoryEmpty.textContent = tr("No invites yet.");
      clientInviteHistoryEmpty.classList.remove("is-hidden");
    }
  }
}

function setClientInviteStatus(message = "", options = {}) {
  const { tone = "info", toast = Boolean(message) } = options;
  const text = String(message || "").trim();
  if (clientInviteStatus) {
    clientInviteStatus.hidden = true;
    clientInviteStatus.textContent = "";
    clientInviteStatus.classList.remove("is-error", "is-success");
  }
  if (toast && text) {
    showStatusToast(text, { tone: tone === "muted" ? "info" : tone });
  }
}

function setClientInviteResult(url = "", meta = {}) {
  if (!clientInviteResult || !clientInviteUrlInput) return;
  const value = String(url || "").trim();
  if (!value) {
    clientInviteResult.classList.add("is-hidden");
    clientInviteUrlInput.value = "";
    if (clientInviteResultEmail) {
      clientInviteResultEmail.textContent = "--";
    }
    if (clientInviteResultExpiry) {
      clientInviteResultExpiry.textContent = "--";
    }
    if (clientInviteCopyButton) {
      clientInviteCopyButton.disabled = true;
    }
    return;
  }
  clientInviteUrlInput.value = value;
  if (clientInviteResultEmail) {
    clientInviteResultEmail.textContent = String(meta?.invitedEmail || "").trim() || "--";
  }
  if (clientInviteResultExpiry) {
    clientInviteResultExpiry.textContent = meta?.expiresAt ? formatHistoryDate(meta.expiresAt) : "--";
  }
  clientInviteResult.classList.remove("is-hidden");
  if (clientInviteCopyButton) {
    clientInviteCopyButton.disabled = false;
  }
}

function setClientInviteBusy(isBusy) {
  clientInviteBusy = Boolean(isBusy);
  if (clientInviteEmailInput) {
    clientInviteEmailInput.disabled = clientInviteBusy;
  }
  if (clientInviteExpirySelect) {
    clientInviteExpirySelect.disabled = clientInviteBusy;
  }
  if (clientInviteCreateButton) {
    clientInviteCreateButton.disabled = clientInviteBusy;
    const label = clientInviteCreateButton.querySelector("span");
    const nextLabel = clientInviteBusy ? tr("Creating invite...") : tr("Create Client");
    if (label) {
      label.textContent = nextLabel;
    } else {
      clientInviteCreateButton.textContent = nextLabel;
    }
  }
  if (clientInviteCopyButton) {
    clientInviteCopyButton.disabled =
      clientInviteBusy || !String(clientInviteUrlInput?.value || "").trim();
  }
}

async function createClientInvite() {
  if (clientInviteBusy) return;
  const invitedEmail = String(clientInviteEmailInput?.value || "").trim().toLowerCase();
  if (!invitedEmail) {
    setClientInviteStatus(tr("Client email is required."), { tone: "error" });
    return;
  }
  if (!isValidEmailFormat(invitedEmail)) {
    setClientInviteStatus(tr("Invalid client email format."), { tone: "error" });
    return;
  }
  const expiresInDays = Math.max(1, Math.min(90, Number(clientInviteExpirySelect?.value) || 14));
  setClientInviteStatus("");
  setClientInviteBusy(true);
  try {
    const payload = await fetchApiWithAuth("/api/auth/invites", {
      method: "POST",
      body: JSON.stringify({
        invitedEmail: invitedEmail || null,
        expiresInDays,
      }),
    });
    const inviteUrl = String(payload?.inviteUrl || "").trim();
    if (!inviteUrl) {
      throw new Error(tr("Could not create account."));
    }
    setClientInviteResult(inviteUrl, {
      invitedEmail,
      expiresAt: payload?.expiresAt,
    });
    setClientInviteStatus(tr("Invite link created."), { tone: "success" });
    if (clientInviteEmailInput) {
      clientInviteEmailInput.value = "";
    }
    await loadAdminDashboard({ quiet: true });
  } catch (error) {
    setClientInviteStatus(error?.message || tr("Could not create account."), { tone: "error" });
  } finally {
    setClientInviteBusy(false);
  }
}

async function copyClientInviteUrl() {
  const inviteUrl = String(clientInviteUrlInput?.value || "").trim();
  if (!inviteUrl) return;
  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(inviteUrl);
    } else {
      clientInviteUrlInput.focus();
      clientInviteUrlInput.select();
      document.execCommand("copy");
    }
    setClientInviteStatus(tr("Invite URL copied."), { tone: "success" });
  } catch (_error) {
    setClientInviteStatus(inviteUrl, { tone: "info" });
  }
}

async function copyInviteHistoryUrl(inviteId) {
  const invite = clientInviteHistory.find((entry) => String(entry?.id || "") === String(inviteId || ""));
  const inviteUrl = String(invite?.invite_url || "").trim();
  if (!inviteUrl) return;
  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(inviteUrl);
    } else {
      setClientInviteResult(inviteUrl);
      await copyClientInviteUrl();
      return;
    }
    showToast(tr("Invite URL copied."), { tone: "success" });
  } catch (_error) {
    showToast(inviteUrl, { tone: "info" });
  }
}

function setAdminInviteActionBusy(inviteId, isBusy) {
  const safeId = String(inviteId || "").trim();
  if (!safeId) return;
  if (isBusy) {
    adminInviteActionBusyIds.add(safeId);
  } else {
    adminInviteActionBusyIds.delete(safeId);
  }
  renderClientInviteHistory(clientInviteHistory);
}

async function revokeClientInvite(inviteId) {
  const safeId = String(inviteId || "").trim();
  if (!safeId || adminInviteActionBusyIds.has(safeId)) return;
  setAdminInviteActionBusy(safeId, true);
  try {
    await fetchApiWithAuth("/api/admin/invites/revoke", {
      method: "POST",
      body: JSON.stringify({ inviteId: safeId }),
    });
    showToast(tr("Invite revoked."), { tone: "success" });
    await loadAdminDashboard({ quiet: true });
  } catch (error) {
    showToast(error?.message || tr("Could not revoke invite."), { tone: "error" });
  } finally {
    setAdminInviteActionBusy(safeId, false);
  }
}

function formatPercent(value) {
  return `${Number(value || 0).toFixed(1).replace(/\.0$/, "")}%`;
}

function getAdminSettingsDraft() {
  return {
    carrier_discount_pct: Math.max(
      0,
      Math.min(100, Number(adminCarrierDiscountInput?.value) || adminSettingsDraft.carrier_discount_pct || 0)
    ),
    client_discount_pct: Math.max(
      0,
      Math.min(100, Number(adminClientDiscountInput?.value) || adminSettingsDraft.client_discount_pct || 0)
    ),
  };
}

function hasAdminSettingsChanges() {
  const draft = getAdminSettingsDraft();
  return (
    Number(draft.carrier_discount_pct).toFixed(2) !==
      Number(adminSettingsSaved.carrier_discount_pct || 0).toFixed(2) ||
    Number(draft.client_discount_pct).toFixed(2) !==
      Number(adminSettingsSaved.client_discount_pct || 0).toFixed(2)
  );
}

function renderAdminSettingsPreview() {
  if (!adminSettingsPreview) return;
  const draft = getAdminSettingsDraft();
  const retainedDiscount = Math.max(0, draft.carrier_discount_pct - draft.client_discount_pct);
  adminSettingsPreview.innerHTML = `
    <div>${tr("Carrier keeps {carrier} and clients receive {client}. Retained margin: {retained}.", {
      carrier: formatPercent(draft.carrier_discount_pct),
      client: formatPercent(draft.client_discount_pct),
      retained: formatPercent(retainedDiscount),
    })}</div>
    <div class="admin-settings-meter">
      <span class="admin-settings-meter-fill" style="width:${Math.min(100, draft.carrier_discount_pct)}%"></span>
      <span class="admin-settings-meter-fill is-secondary" style="width:${Math.min(100, draft.client_discount_pct)}%"></span>
    </div>
  `;
  if (adminSettingsSaveButton) {
    adminSettingsSaveButton.disabled = adminDashboardLoading || !hasAdminSettingsChanges();
  }
}

function applyAdminSettings(settings = {}) {
  adminSettingsSaved = {
    carrier_discount_pct: Number(settings?.carrier_discount_pct) || 0,
    client_discount_pct: Number(settings?.client_discount_pct) || 0,
  };
  adminSettingsDraft = { ...adminSettingsSaved };
  if (adminCarrierDiscountInput) {
    adminCarrierDiscountInput.value = String(adminSettingsSaved.carrier_discount_pct);
  }
  if (adminClientDiscountInput) {
    adminClientDiscountInput.value = String(adminSettingsSaved.client_discount_pct);
  }
  renderAdminSettingsPreview();
}

function renderAdminSummary(summary = {}) {
  if (adminSummaryClients) {
    adminSummaryClients.textContent = String(summary?.total_clients ?? 0);
  }
  if (adminSummaryActiveClients) {
    adminSummaryActiveClients.textContent = tr("{count} active", {
      count: Number(summary?.active_clients) || 0,
    });
  }
  if (adminSummaryInvites) {
    adminSummaryInvites.textContent = String(summary?.open_invites ?? 0);
  }
  if (adminSummaryRevenue) {
    adminSummaryRevenue.textContent = formatMoney(summary?.total_revenue_ex_vat || 0);
  }
  if (adminSummaryProfit) {
    adminSummaryProfit.textContent = formatMoney(summary?.total_profit_ex_vat || 0);
    adminSummaryProfit.classList.toggle(
      "is-negative",
      Number(summary?.total_profit_ex_vat || 0) < 0
    );
  }
  if (adminSummaryProfitMeta) {
    adminSummaryProfitMeta.textContent = tr("Based on current discount rules");
  }
}

function getAdminClientProfile(client) {
  return buildMockAccountProfile(client?.user || null) || {
    companyName: "--",
    contactName: "--",
    contactEmail: client?.user?.email || "--",
    contactPhone: "--",
    billingAddress: "--",
    taxId: "--",
    customerId: "--",
    accountManager: "--",
  };
}

function normalizeAdminClientBilling(client) {
  const raw = client?.billing && typeof client.billing === "object" ? client.billing : {};
  let invoiceEnabled = raw.invoice_enabled === true;
  let cardEnabled = raw.card_enabled === false ? false : true;
  if (!invoiceEnabled && !cardEnabled) {
    cardEnabled = true;
  }
  return {
    invoice_enabled: invoiceEnabled,
    card_enabled: cardEnabled,
  };
}

function getAdminClientPaymentMode(client) {
  const billing = normalizeAdminClientBilling(client);
  if (billing.invoice_enabled && billing.card_enabled) return "hybrid";
  if (billing.card_enabled) return "card";
  return "invoice";
}

function getAdminPaymentModeLabel(mode) {
  if (mode === "hybrid") return `${tr("Invoice")} + ${tr("Card")}`;
  if (mode === "card") return tr("Card");
  return tr("Invoice");
}

function getAdminActivityLabel(status) {
  if (status === "active") return tr("Active");
  if (status === "quiet") return tr("Quiet");
  return tr("Dormant");
}

function getAdminClientValueClass(value) {
  if (Number(value || 0) < 0) return "admin-client-value is-negative";
  if (Number(value || 0) > 0) return "admin-client-value is-positive";
  return "admin-client-value";
}

function isAdminClientBillingBusy(userId) {
  return adminClientBillingBusyIds.has(String(userId || ""));
}

function setAdminClientBillingBusy(userId, isBusy) {
  const safeUserId = String(userId || "").trim();
  if (!safeUserId) return;
  if (isBusy) {
    adminClientBillingBusyIds.add(safeUserId);
  } else {
    adminClientBillingBusyIds.delete(safeUserId);
  }
  renderAdminClientsList();
}

function refreshAdminOnlyTestTools() {
  const canUseAdminTestTools = Boolean(currentUser && adminAccessAllowed);
  if (adminOnlyTestTools) {
    adminOnlyTestTools.classList.toggle("is-hidden", !canUseAdminTestTools);
  }
  if (autoFillButton) {
    autoFillButton.disabled = !canUseAdminTestTools;
  }
  if (autoCsvButton) {
    autoCsvButton.disabled = !canUseAdminTestTools;
  }
}

function renderAdminMockDataButton() {
  if (!adminMockDataButton) return;
  const label = adminMockDataButton.querySelector("span");
  const nextLabel = adminMockModeEnabled ? tr("Restore Live Data") : tr("Fill Mock Data");
  if (label) {
    label.textContent = nextLabel;
  } else {
    adminMockDataButton.textContent = nextLabel;
  }
}

function renderAdminClientsList() {
  if (!adminClientsList || !adminClientsEmpty) return;
  adminClientsList.innerHTML = "";
  let filtered = Array.isArray(adminClients) ? adminClients.slice() : [];

  const search = String(adminClientSearch || "").trim().toLowerCase();
  if (search) {
    filtered = filtered.filter((client) => {
      const profile = getAdminClientProfile(client);
      const haystack = [
        profile.companyName,
        profile.contactName,
        profile.contactEmail,
        profile.customerId,
        profile.accountManager,
      ]
        .join(" ")
        .toLowerCase();
      return haystack.includes(search);
    });
  }

  if (adminClientFilter !== "all") {
    filtered = filtered.filter(
      (client) => String(client?.metrics?.activity_status || "dormant") === adminClientFilter
    );
  }

  filtered.sort((left, right) => {
    if (adminClientSort === "revenue") {
      return (Number(right?.metrics?.total_revenue_ex_vat) || 0) - (Number(left?.metrics?.total_revenue_ex_vat) || 0);
    }
    if (adminClientSort === "profit") {
      return (Number(right?.metrics?.total_profit_ex_vat) || 0) - (Number(left?.metrics?.total_profit_ex_vat) || 0);
    }
    if (adminClientSort === "mrr") {
      return (Number(right?.metrics?.mrr_ex_vat) || 0) - (Number(left?.metrics?.mrr_ex_vat) || 0);
    }
    if (adminClientSort === "signup") {
      return Date.parse(right?.user?.created_at || 0) - Date.parse(left?.user?.created_at || 0);
    }
    if (adminClientSort === "volume") {
      return (Number(right?.metrics?.avg_parcels_per_month) || 0) - (Number(left?.metrics?.avg_parcels_per_month) || 0);
    }
    return Date.parse(right?.metrics?.last_generation_at || right?.user?.created_at || 0) -
      Date.parse(left?.metrics?.last_generation_at || left?.user?.created_at || 0);
  });

  if (!filtered.length) {
    adminClientsEmpty.textContent = adminDashboardLoading
      ? tr("Loading client accounts...")
      : tr("No client accounts match the current filter.");
    adminClientsEmpty.classList.remove("is-hidden");
    return;
  }

  adminClientsEmpty.classList.add("is-hidden");
  filtered.forEach((client) => {
    const profile = getAdminClientProfile(client);
    const metrics = client?.metrics || {};
    const billing = normalizeAdminClientBilling(client);
    const paymentMode = metrics.payment_mode || getAdminClientPaymentMode(client);
    const paymentTrackingLabel =
      paymentMode === "invoice" ? tr("Invoice Tracking") : tr("Payment Tracking");
    const paymentTimingLabel =
      paymentMode === "card" && Number(metrics.avg_payment_days) === 0
        ? tr("Instant")
        : metrics.avg_payment_days != null
          ? `${metrics.avg_payment_days} ${tr("days")}`
          : "--";
    const userId = String(client?.user?.id || "").trim();
    const billingBusy = isAdminClientBillingBusy(userId);
    const row = document.createElement("article");
    row.className = "admin-client-row";
    row.innerHTML = `
      <div class="admin-client-identity">
        <div class="admin-client-name">${profile.companyName}</div>
        <div class="admin-client-meta">${profile.contactName} • ${profile.contactEmail} • ${profile.contactPhone}</div>
        <div class="admin-client-address">${profile.billingAddress}</div>
        <div class="admin-client-submeta mono">${profile.customerId} • ${profile.accountManager}</div>
      </div>
      <div class="admin-client-stack">
        <div class="admin-client-metric">
          <span class="admin-client-key">${tr("Total Revenue")}</span>
          <span class="admin-client-value">${formatMoney(metrics.total_revenue_ex_vat || 0)}</span>
        </div>
        <div class="admin-client-metric">
          <span class="admin-client-key">${tr("Total Profit")}</span>
          <span class="${getAdminClientValueClass(metrics.total_profit_ex_vat)}">${formatMoney(metrics.total_profit_ex_vat || 0)}</span>
        </div>
      </div>
      <div class="admin-client-stack">
        <div class="admin-client-metric">
          <span class="admin-client-key">MRR</span>
          <span class="admin-client-value">${formatMoney(metrics.mrr_ex_vat || 0)}</span>
        </div>
        <div class="admin-client-metric">
          <span class="admin-client-key">MRP</span>
          <span class="${getAdminClientValueClass(metrics.mrp_ex_vat)}">${formatMoney(metrics.mrp_ex_vat || 0)}</span>
        </div>
      </div>
      <div class="admin-client-stack">
        <div class="admin-client-metric">
          <span class="admin-client-key">${tr("Parcels / Month")}</span>
          <span class="admin-client-value">${Number(metrics.avg_parcels_per_month || 0).toFixed(1)}</span>
        </div>
        <div class="admin-client-metric">
          <span class="admin-client-key">${tr("Last Generation")}</span>
          <span class="admin-client-value">${metrics.last_generation_at ? formatHistoryDate(metrics.last_generation_at) : "--"}</span>
        </div>
      </div>
      <div class="admin-client-stack">
        <div class="admin-client-metric">
          <span class="admin-client-key">${tr("Registered")}</span>
          <span class="admin-client-value">${client?.user?.created_at ? formatHistoryDate(client.user.created_at) : "--"}</span>
        </div>
        <div class="admin-client-metric">
          <span class="admin-client-key">${tr("Avg. Payment Time")}</span>
          <span class="admin-client-value">${paymentTimingLabel}</span>
        </div>
      </div>
      <div class="admin-client-stack">
        <div class="admin-client-metric">
          <span class="admin-client-key">${tr("Payment Methods")}</span>
          <div class="admin-client-methods">
            <button
              type="button"
              class="admin-client-method-toggle${billing.invoice_enabled ? " is-active" : ""}"
              data-admin-billing-toggle="invoice"
              data-admin-client-id="${userId}"
              ${billingBusy ? "disabled" : ""}
            >
              <span class="admin-client-method-box" aria-hidden="true"></span>
              <span>${tr("Invoice")}</span>
            </button>
            <button
              type="button"
              class="admin-client-method-toggle${billing.card_enabled ? " is-active" : ""}"
              data-admin-billing-toggle="card"
              data-admin-client-id="${userId}"
              ${billingBusy ? "disabled" : ""}
            >
              <span class="admin-client-method-box" aria-hidden="true"></span>
              <span>${tr("Card")}</span>
            </button>
          </div>
          <span class="admin-client-method-mode mono">${getAdminPaymentModeLabel(paymentMode)}</span>
        </div>
      </div>
      <div class="admin-client-stack">
        <div class="admin-client-metric">
          <span class="admin-client-key">${paymentTrackingLabel}</span>
          <div class="admin-client-pills">
            <span class="admin-client-pill is-${metrics.activity_status || "dormant"}">${getAdminActivityLabel(metrics.activity_status || "dormant")}</span>
            <span class="admin-client-pill is-billing">${formatAdminTrackingStatus(metrics.last_invoice_tracking)}</span>
          </div>
        </div>
      </div>
    `;
    adminClientsList.appendChild(row);
  });
}

function buildAdminSummaryFromClients(clients = [], openInvites = 0) {
  const rows = Array.isArray(clients) ? clients : [];
  const activeClients = rows.filter((client) => client?.metrics?.activity_status === "active").length;
  const totalRevenue = rows.reduce(
    (sum, client) => sum + Math.max(0, Number(client?.metrics?.total_revenue_ex_vat) || 0),
    0
  );
  const totalProfit = rows.reduce(
    (sum, client) => sum + Number(client?.metrics?.total_profit_ex_vat || 0),
    0
  );
  return {
    total_clients: rows.length,
    active_clients: activeClients,
    open_invites: Number(openInvites) || 0,
    total_revenue_ex_vat: Number(totalRevenue.toFixed(2)),
    total_profit_ex_vat: Number(totalProfit.toFixed(2)),
  };
}

function normalizeBillingForMetrics(billing) {
  const normalized = normalizeAdminClientBilling({ billing });
  return {
    invoice_enabled: normalized.invoice_enabled,
    card_enabled: normalized.card_enabled,
  };
}

function applyBillingToClientMetrics(metrics = {}, billing) {
  const normalizedBilling = normalizeBillingForMetrics(billing);
  const nextMetrics = { ...(metrics || {}) };
  const paymentMode =
    normalizedBilling.invoice_enabled && normalizedBilling.card_enabled
      ? "hybrid"
      : normalizedBilling.card_enabled
        ? "card"
        : "invoice";
  const hasRevenue = Number(nextMetrics.total_revenue_ex_vat || 0) > 0;
  nextMetrics.payment_mode = paymentMode;
  nextMetrics.invoice_enabled = normalizedBilling.invoice_enabled;
  nextMetrics.card_enabled = normalizedBilling.card_enabled;
  if (paymentMode === "card") {
    nextMetrics.avg_payment_days = hasRevenue ? 0 : null;
    nextMetrics.last_invoice_tracking = hasRevenue ? "Card auto-charge" : "No billable activity";
  } else if (paymentMode === "hybrid") {
    nextMetrics.last_invoice_tracking = hasRevenue ? "Hybrid billing" : "No billable activity";
  } else if (!String(nextMetrics.last_invoice_tracking || "").trim()) {
    nextMetrics.last_invoice_tracking = hasRevenue ? "Billing not live" : "No billable activity";
  }
  return nextMetrics;
}

function updateAdminClientBillingState(userId, billing) {
  const safeUserId = String(userId || "").trim();
  if (!safeUserId) return false;
  let updated = false;
  const applyUpdate = (client) => {
    if (String(client?.user?.id || "").trim() !== safeUserId) {
      return client;
    }
    updated = true;
    return {
      ...client,
      billing: normalizeBillingForMetrics(billing),
      metrics: applyBillingToClientMetrics(client?.metrics || {}, billing),
    };
  };
  adminClients = adminClients.map(applyUpdate);
  if (adminDashboardState && Array.isArray(adminDashboardState.clients)) {
    adminDashboardState = {
      ...adminDashboardState,
      clients: adminDashboardState.clients.map(applyUpdate),
    };
  }
  return updated;
}

function formatAdminTrackingStatus(status) {
  const value = String(status || "").trim();
  if (!value) return tr("Billing not live");
  return tr(value);
}

function buildMockAdminClients() {
  const now = Date.now();
  const rows = [
    {
      companyName: "Nordlane Commerce",
      contactName: "Sophie Van Dijk",
      email: "ops@nordlane.eu",
      customerId: "CUST-502194",
      accountManager: "A. Lambert",
      billingAddress: "Cantersteen 12, 1000 Brussels, Belgium",
      payment: { invoice_enabled: true, card_enabled: false },
      metrics: {
        total_revenue_ex_vat: 19640.5,
        total_profit_ex_vat: 4880.45,
        mrr_ex_vat: 3273.42,
        mrp_ex_vat: 813.41,
        avg_parcels_per_month: 284.6,
        avg_payment_days: 18,
        last_invoice_tracking: "Invoice sent · pending",
        activity_status: "active",
      },
    },
    {
      companyName: "Kite Retail Group",
      contactName: "Thomas Meers",
      email: "finance@kiteretail.eu",
      customerId: "CUST-874300",
      accountManager: "N. Martin",
      billingAddress: "Avenue Louise 221, 1050 Brussels, Belgium",
      payment: { invoice_enabled: false, card_enabled: true },
      metrics: {
        total_revenue_ex_vat: 8470.2,
        total_profit_ex_vat: 2123.5,
        mrr_ex_vat: 1411.7,
        mrp_ex_vat: 353.92,
        avg_parcels_per_month: 112.3,
        avg_payment_days: 0,
        last_invoice_tracking: "Card auto-charge",
        activity_status: "active",
      },
    },
    {
      companyName: "Arta Home",
      contactName: "Camille Dupont",
      email: "shipping@arta-home.com",
      customerId: "CUST-719255",
      accountManager: "A. Lambert",
      billingAddress: "Rue des Guillemins 18, 4000 Liège, Belgium",
      payment: { invoice_enabled: true, card_enabled: true },
      metrics: {
        total_revenue_ex_vat: 12324.1,
        total_profit_ex_vat: 2840.12,
        mrr_ex_vat: 1027.01,
        mrp_ex_vat: 236.68,
        avg_parcels_per_month: 96.7,
        avg_payment_days: 8,
        last_invoice_tracking: "Hybrid billing",
        activity_status: "quiet",
      },
    },
    {
      companyName: "Veldmark Health",
      contactName: "Lina Verbeek",
      email: "logistics@veldmark.eu",
      customerId: "CUST-402881",
      accountManager: "M. Rossi",
      billingAddress: "Koning Albertlaan 64, 9000 Ghent, Belgium",
      payment: { invoice_enabled: true, card_enabled: false },
      metrics: {
        total_revenue_ex_vat: 4380.0,
        total_profit_ex_vat: 911.44,
        mrr_ex_vat: 547.5,
        mrp_ex_vat: 113.93,
        avg_parcels_per_month: 44.8,
        avg_payment_days: 29,
        last_invoice_tracking: "Reminder sent (D-1)",
        activity_status: "quiet",
      },
    },
    {
      companyName: "Heliox Parts",
      contactName: "Noah Jacobs",
      email: "labels@helioxparts.com",
      customerId: "CUST-638512",
      accountManager: "N. Martin",
      billingAddress: "Industrieweg 10, 3001 Leuven, Belgium",
      payment: { invoice_enabled: false, card_enabled: true },
      metrics: {
        total_revenue_ex_vat: 2160.4,
        total_profit_ex_vat: 422.12,
        mrr_ex_vat: 180.03,
        mrp_ex_vat: 35.18,
        avg_parcels_per_month: 21.1,
        avg_payment_days: 0,
        last_invoice_tracking: "Card auto-charge",
        activity_status: "dormant",
      },
    },
  ];

  return rows.map((row, index) => {
    const createdAt = new Date(now - (index + 1) * 1000 * 60 * 60 * 24 * 35).toISOString();
    const lastGenerationAt = new Date(now - (index + 1) * 1000 * 60 * 60 * 24 * (index * 6 + 2)).toISOString();
    const metrics = applyBillingToClientMetrics(
      {
        ...row.metrics,
        last_generation_at: lastGenerationAt,
      },
      row.payment
    );
    return {
      user: {
        id: `mock-client-${index + 1}`,
        email: row.email,
        created_at: createdAt,
        user_metadata: {
          company_name: row.companyName,
          contact_name: row.contactName,
          contact_email: row.email,
          contact_phone: "+32 2 555 01 0" + index,
          billing_address: row.billingAddress,
          customer_id: row.customerId,
          account_manager: row.accountManager,
        },
      },
      billing: normalizeBillingForMetrics(row.payment),
      metrics,
    };
  });
}

function toggleAdminMockData() {
  if (adminDashboardLoading) return;
  if (!adminMockModeEnabled) {
    adminMockSnapshot = {
      clients: JSON.parse(JSON.stringify(adminClients || [])),
      summary: JSON.parse(JSON.stringify(adminDashboardState?.summary || {})),
    };
    const mockClients = buildMockAdminClients();
    adminClients = mockClients;
    adminMockModeEnabled = true;
    const openInvites = Number(adminDashboardState?.summary?.open_invites || 0);
    adminDashboardState = {
      ...(adminDashboardState || {}),
      clients: mockClients,
      summary: buildAdminSummaryFromClients(mockClients, openInvites),
    };
    renderAdminSummary(adminDashboardState.summary || {});
    renderAdminClientsList();
    renderAdminMockDataButton();
    showToast(tr("Mock data loaded (frontend only)."), { tone: "info" });
    return;
  }

  adminMockModeEnabled = false;
  if (adminMockSnapshot) {
    adminClients = Array.isArray(adminMockSnapshot.clients) ? adminMockSnapshot.clients : [];
    adminDashboardState = {
      ...(adminDashboardState || {}),
      clients: adminClients,
      summary: adminMockSnapshot.summary || {},
    };
  }
  adminMockSnapshot = null;
  renderAdminSummary(adminDashboardState?.summary || {});
  renderAdminClientsList();
  renderAdminMockDataButton();
  showToast(tr("Live admin data restored."), { tone: "success" });
}

async function updateAdminClientBilling(userId, method) {
  const safeUserId = String(userId || "").trim();
  const normalizedMethod = String(method || "").trim();
  if (!safeUserId || !["invoice", "card"].includes(normalizedMethod)) return;
  if (isAdminClientBillingBusy(safeUserId)) return;

  const client = adminClients.find((entry) => String(entry?.user?.id || "").trim() === safeUserId);
  if (!client) return;
  const current = normalizeAdminClientBilling(client);
  const next = { ...current };
  if (normalizedMethod === "invoice") {
    next.invoice_enabled = !current.invoice_enabled;
  } else {
    next.card_enabled = !current.card_enabled;
  }
  if (!next.invoice_enabled && !next.card_enabled) {
    showToast(tr("At least one payment method must remain enabled."), { tone: "error" });
    return;
  }

  if (adminMockModeEnabled) {
    updateAdminClientBillingState(safeUserId, next);
    renderAdminSummary(adminDashboardState?.summary || {});
    renderAdminClientsList();
    showToast(tr("Client billing updated."), { tone: "success" });
    return;
  }

  setAdminClientBillingBusy(safeUserId, true);
  try {
    const response = await fetchApiWithAuth("/api/admin/client-billing", {
      method: "POST",
      body: JSON.stringify({
        userId: safeUserId,
        invoiceEnabled: next.invoice_enabled,
        cardEnabled: next.card_enabled,
      }),
    });
    const returnedBilling = response?.billing || next;
    updateAdminClientBillingState(safeUserId, returnedBilling);
    renderAdminSummary(adminDashboardState?.summary || {});
    renderAdminClientsList();
    showToast(tr("Client billing updated."), { tone: "success" });
  } catch (error) {
    showToast(error?.message || tr("Could not update client billing settings."), { tone: "error" });
  } finally {
    setAdminClientBillingBusy(safeUserId, false);
  }
}

async function loadAdminAccessStatus(options = {}) {
  const { quiet = false } = options;
  const activeUserId = String(currentUser?.id || "").trim();
  if (!activeUserId) {
    adminAccessStatusRequestId += 1;
    adminAccessStatusPromise = null;
    adminAccessAllowed = false;
    if (openLeadsPageButton) {
      openLeadsPageButton.disabled = false;
      openLeadsPageButton.classList.add("is-hidden");
      openLeadsPageButton.removeAttribute("aria-busy");
    }
    if (openAdminPageButton) {
      openAdminPageButton.disabled = false;
      openAdminPageButton.classList.add("is-hidden");
      openAdminPageButton.removeAttribute("aria-busy");
    }
    refreshAdminOnlyTestTools();
    return false;
  }
  if (adminAccessStatusPromise) {
    return adminAccessStatusPromise;
  }
  const requestId = ++adminAccessStatusRequestId;
  if (openLeadsPageButton) {
    openLeadsPageButton.disabled = true;
    openLeadsPageButton.setAttribute("aria-busy", "true");
  }
  if (openAdminPageButton) {
    openAdminPageButton.disabled = true;
    openAdminPageButton.setAttribute("aria-busy", "true");
  }
  const requestPromise = (async () => {
    let allowed = false;
    try {
      const payload = await fetchApiWithAuth("/api/admin/status", { timeoutMs: 8000 });
      allowed = Boolean(payload?.allowed);
    } catch (error) {
      if (requestId === adminAccessStatusRequestId && String(currentUser?.id || "").trim() === activeUserId) {
        adminAccessAllowed = false;
        if (!quiet) {
          showToast(error?.message || tr("Could not verify admin access."), { tone: "error" });
        }
      }
    }

    const stillSameUser = String(currentUser?.id || "").trim() === activeUserId;
    const isLatestRequest = requestId === adminAccessStatusRequestId;
    if (stillSameUser && isLatestRequest) {
      adminAccessAllowed = allowed;
      if (openLeadsPageButton) {
        openLeadsPageButton.disabled = false;
        openLeadsPageButton.classList.toggle("is-hidden", !adminAccessAllowed);
        openLeadsPageButton.removeAttribute("aria-busy");
      }
      if (openAdminPageButton) {
        openAdminPageButton.disabled = false;
        openAdminPageButton.classList.toggle("is-hidden", !adminAccessAllowed);
        openAdminPageButton.removeAttribute("aria-busy");
      }
      refreshAdminOnlyTestTools();
      if (!adminAccessAllowed && currentMainView === "admin") {
        setAdminPageVisible(false, { replace: true });
      }
      if (!adminAccessAllowed && currentMainView === "leads") {
        setLeadsPageVisible(false, { replace: true });
      }
      return adminAccessAllowed;
    }

    return Boolean(stillSameUser && adminAccessAllowed);
  })();

  let inFlightPromise = null;
  inFlightPromise = requestPromise.finally(() => {
    if (adminAccessStatusPromise === inFlightPromise) {
      adminAccessStatusPromise = null;
    }
    if (openAdminPageButton) {
      openAdminPageButton.disabled = false;
      openAdminPageButton.removeAttribute("aria-busy");
    }
    if (openLeadsPageButton) {
      openLeadsPageButton.disabled = false;
      openLeadsPageButton.removeAttribute("aria-busy");
    }
  });
  adminAccessStatusPromise = inFlightPromise;

  return inFlightPromise;
}

async function loadAdminDashboard(options = {}) {
  const { quiet = false } = options;
  if (!currentUser) return false;
  const hasAccess = adminAccessAllowed || (await loadAdminAccessStatus({ quiet: true }));
  if (!hasAccess) {
    adminDashboardLoaded = false;
    adminDashboardState = null;
    adminClients = [];
    adminBillingInvoices = [];
    adminClientBillingBusyIds = new Set();
    adminMockModeEnabled = false;
    adminMockSnapshot = null;
    renderAdminMockDataButton();
    renderAdminClientsList();
    renderAdminInvoiceList();
    if (currentMainView === "admin") {
      setAdminPageVisible(false, { replace: true });
    }
    if (!quiet) {
      showToast(tr("You are not allowed to access the admin panel."), { tone: "error" });
    }
    return false;
  }
  adminDashboardLoading = true;
  renderAdminClientsList();
  try {
    const payload = await fetchApiWithAuth("/api/admin/dashboard", { timeoutMs: 20000 });
    adminClientBillingBusyIds = new Set();
    adminMockModeEnabled = false;
    adminMockSnapshot = null;
    adminDashboardState = payload && typeof payload === "object" ? payload : {};
    clientInviteHistory = Array.isArray(adminDashboardState?.invites) ? adminDashboardState.invites : [];
    adminClients = Array.isArray(adminDashboardState?.clients) ? adminDashboardState.clients : [];
    syncAdminBillingFromDashboard(adminDashboardState);
    renderAdminSummary(adminDashboardState?.summary || {});
    renderClientInviteHistory(clientInviteHistory);
    applyAdminSettings(adminDashboardState?.settings || {});
    adminDashboardLoaded = true;
    renderAdminMockDataButton();
    renderAdminClientsList();
    return true;
  } catch (error) {
    adminClientBillingBusyIds = new Set();
    adminMockModeEnabled = false;
    adminMockSnapshot = null;
    adminDashboardLoaded = false;
    adminDashboardState = null;
    adminClients = [];
    adminBillingInvoices = [];
    renderAdminMockDataButton();
    renderAdminClientsList();
    renderAdminInvoiceList();
    if (!quiet) {
      showToast(error?.message || tr("Could not load admin dashboard."), { tone: "error" });
    }
    return false;
  } finally {
    adminDashboardLoading = false;
    renderAdminSettingsPreview();
  }
}

async function saveAdminSettings() {
  if (adminDashboardLoading || !hasAdminSettingsChanges()) return;
  const payload = getAdminSettingsDraft();
  adminDashboardLoading = true;
  renderAdminSettingsPreview();
  try {
    const response = await fetchApiWithAuth("/api/admin/settings", {
      method: "POST",
      body: JSON.stringify(payload),
    });
    applyAdminSettings(response?.settings || payload);
    showToast(tr("Discount settings saved."), { tone: "success" });
    await loadAdminDashboard({ quiet: true });
  } catch (error) {
    showToast(error?.message || tr("Could not save admin settings."), { tone: "error" });
  } finally {
    adminDashboardLoading = false;
    renderAdminSettingsPreview();
  }
}

async function loadBillingOverview(options = {}) {
  const { quiet = false } = options;
  if (!currentUser) {
    billingOverview = null;
    updateSummary();
    renderCheckoutStepMode();
    renderAccountBillingOverview();
    return null;
  }
  try {
    const payload = await fetchApiWithAuth("/api/billing/overview", { timeoutMs: 12000 });
    billingOverview = payload && typeof payload === "object" ? payload : null;
    updateSummary();
    renderCheckoutStepMode();
    renderAccountBillingOverview();
    return billingOverview;
  } catch (error) {
    billingOverview = null;
    updateSummary();
    renderCheckoutStepMode();
    renderAccountBillingOverview();
    if (!quiet) {
      showToast(error?.message || tr("Could not load billing overview."), { tone: "error" });
    }
    return null;
  }
}

function setAdminBillingBusy(isBusy) {
  adminBillingBusy = Boolean(isBusy);
  if (adminBillingRunPreviewButton) adminBillingRunPreviewButton.disabled = adminBillingBusy;
  if (adminBillingRunCreateButton) adminBillingRunCreateButton.disabled = adminBillingBusy;
  if (adminBillingRunSendButton) adminBillingRunSendButton.disabled = adminBillingBusy;
  if (adminBillingSendTestButton) adminBillingSendTestButton.disabled = adminBillingBusy;
  if (adminBillingSendTestSequenceButton) {
    adminBillingSendTestSequenceButton.disabled = adminBillingBusy;
  }
  if (adminReportsSendTestButton) adminReportsSendTestButton.disabled = adminBillingBusy;
  if (adminAgreementPreviewButton) adminAgreementPreviewButton.disabled = adminBillingBusy;
  if (adminAgreementSendTestButton) adminAgreementSendTestButton.disabled = adminBillingBusy;
}

function setAdminBillingStatus(message = "", options = {}) {
  const text = String(message || "").trim();
  const tone = String(options?.tone || "info").trim() || "info";
  if (adminBillingRunStatus) {
    adminBillingRunStatus.hidden = true;
    adminBillingRunStatus.textContent = "";
    adminBillingRunStatus.classList.remove("is-error", "is-success");
  }
  if (text) {
    showStatusToast(text, { tone });
  }
}

function renderAdminBillingRunResult(payload = null) {
  if (!adminBillingRunResult) return;
  if (!payload || typeof payload !== "object") {
    adminBillingRunResult.textContent = tr("No billing runs yet.");
    return;
  }
  const run = payload?.run && typeof payload.run === "object" ? payload.run : {};
  const reminders = Array.isArray(payload?.reminders) ? payload.reminders : [];
  const summary = [
    `MODE ${String(run.mode || "--").toUpperCase()}`,
    `PERIOD ${String(run.period_start || "--")} -> ${String(run.period_end || "--")}`,
    `ROWS ${Number(run.rows_scanned || 0)}`,
    `CREATED ${Number(run.invoices_created || 0)}`,
    `UPDATED ${Number(run.invoices_updated || 0)}`,
    `SENT ${Number(run.invoices_sent || 0)}`,
    `REMINDERS ${reminders.length}`,
    `TOTAL ${formatMoney(Number(run?.totals?.total_inc_vat || run?.totals?.subtotal_ex_vat || 0))}`,
  ];
  adminBillingRunResult.textContent = summary.join("   |   ");
}

function getAdminInvoiceStatusClass(status) {
  const normalized = String(status || "").trim().toLowerCase();
  if (["draft", "sent", "overdue", "paid", "cancelled"].includes(normalized)) return normalized;
  return "draft";
}

function renderAdminInvoiceList() {
  if (!adminInvoiceList || !adminInvoiceEmpty) return;
  adminInvoiceList.innerHTML = "";
  const rows = Array.isArray(adminBillingInvoices) ? adminBillingInvoices : [];
  if (!rows.length) {
    adminInvoiceEmpty.classList.remove("is-hidden");
    return;
  }
  adminInvoiceEmpty.classList.add("is-hidden");
  rows.slice(0, 120).forEach((invoice) => {
    const status = String(invoice?.status || "draft").toLowerCase();
    const row = document.createElement("article");
    row.className = "admin-billing-item";
    row.innerHTML = `
      <div class="admin-billing-item-main">
        <div class="admin-billing-item-title">
          <span>${String(invoice?.reference || "--")}</span>
          <span class="admin-billing-status-pill is-${getAdminInvoiceStatusClass(status)}">${escapeHtml(
            String(invoice?.tracking_label || status || "draft")
          )}</span>
        </div>
        <div class="admin-billing-item-sub">
          ${escapeHtml(String(invoice?.company_name || "--"))} • ${escapeHtml(
            String(invoice?.contact_email || "--")
          )} • ${escapeHtml(String(invoice?.period_start || "--"))} - ${escapeHtml(
            String(invoice?.period_end || "--")
          )}
        </div>
        <div class="admin-billing-item-sub mono">
          ${escapeHtml(tr("Total"))} ${formatMoney(
      Number(invoice?.total_inc_vat || invoice?.subtotal_ex_vat || 0)
    )}
        </div>
      </div>
      <div class="admin-billing-item-actions">
        <button type="button" class="btn btn-secondary btn-sm" data-admin-invoice-send="${String(
          invoice?.id || ""
        )}" ${adminBillingBusy || status === "paid" || status === "cancelled" ? "disabled" : ""}>
          <span>${tr("Send")}</span>
        </button>
        <button type="button" class="btn btn-primary btn-sm" data-admin-invoice-paid="${String(
          invoice?.id || ""
        )}" ${adminBillingBusy || status === "paid" || status === "cancelled" ? "disabled" : ""}>
          <span>${tr("Mark Paid")}</span>
        </button>
      </div>
    `;
    adminInvoiceList.appendChild(row);
  });
}

function syncAdminBillingFromDashboard(payload = null) {
  const billingBlock = payload?.billing && typeof payload.billing === "object" ? payload.billing : null;
  adminBillingInvoices = Array.isArray(billingBlock?.invoices) ? billingBlock.invoices : [];
  renderAdminInvoiceList();
}

async function loadAdminInvoices(options = {}) {
  const { quiet = true } = options;
  try {
    const payload = await fetchApiWithAuth("/api/admin/invoices?limit=120", { timeoutMs: 20000 });
    adminBillingInvoices = Array.isArray(payload?.invoices) ? payload.invoices : [];
    renderAdminInvoiceList();
    return adminBillingInvoices;
  } catch (error) {
    if (!quiet) {
      showToast(error?.message || tr("Could not load invoices."), { tone: "error" });
    }
    return [];
  }
}

async function sendAdminInvoiceWithRenderer(invoiceId) {
  const safeInvoiceId = String(invoiceId || "").trim();
  if (!safeInvoiceId) {
    throw new Error(tr("Invoice id is required."));
  }
  return fetchApiWithAuth("/api/admin/invoices/send", {
    method: "POST",
    timeoutMs: 120000,
    body: JSON.stringify({
      invoiceId: safeInvoiceId,
    }),
  });
}

async function runAdminBillingCycle(mode = "preview") {
  if (adminBillingBusy) return;
  setAdminBillingStatus("");
  setAdminBillingBusy(true);
  try {
    const runMode = mode === "send" ? "create" : mode;
    const payload = await fetchApiWithAuth("/api/admin/invoices/run", {
      method: "POST",
      timeoutMs: 60000,
      body: JSON.stringify({
        mode: runMode,
        periodMode: "previous",
        includeReminders: true,
      }),
    });
    if (mode === "send") {
      const invoiceRows = Array.isArray(payload?.run?.invoices) ? payload.run.invoices : [];
      let sentCount = 0;
      let failedCount = 0;
      for (const invoice of invoiceRows) {
        const invoiceId = String(invoice?.id || "").trim();
        if (!invoiceId) continue;
        try {
          await sendAdminInvoiceWithRenderer(invoiceId);
          sentCount += 1;
        } catch (_error) {
          failedCount += 1;
        }
      }
      if (payload?.run && typeof payload.run === "object") {
        payload.run.invoices_sent = sentCount;
      }
      renderAdminBillingRunResult(payload);
      if (failedCount > 0) {
        throw new Error(
          tr("Billing run sent {sent} invoices with {failed} failures.", {
            sent: String(sentCount),
            failed: String(failedCount),
          })
        );
      }
    } else {
      renderAdminBillingRunResult(payload);
    }
    setAdminBillingStatus(tr("Billing run completed."), { tone: "success" });
    await loadAdminDashboard({ quiet: true });
    await loadAdminInvoices({ quiet: true });
    await loadBillingOverview({ quiet: true });
  } catch (error) {
    setAdminBillingStatus(error?.message || tr("Billing run failed."), { tone: "error" });
  } finally {
    setAdminBillingBusy(false);
  }
}

async function sendAdminBillingTestEmail() {
  if (adminBillingBusy) return;
  const toEmail = String(adminBillingTestEmailInput?.value || "").trim();
  if (!toEmail) {
    setAdminBillingStatus(tr("A valid test email is required."), { tone: "error" });
    return;
  }
  setAdminBillingBusy(true);
  setAdminBillingStatus("");
  try {
    const payload = await fetchApiWithAuth("/api/admin/invoices/send-test", {
      method: "POST",
      timeoutMs: 120000,
      body: JSON.stringify({
        toEmail,
      }),
    });
    setAdminBillingStatus(
      tr("Test invoice email sent to {email}.", { email: String(payload?.to || toEmail) }),
      { tone: "success" }
    );
  } catch (error) {
    setAdminBillingStatus(error?.message || tr("Could not send test invoice email."), {
      tone: "error",
    });
  } finally {
    setAdminBillingBusy(false);
  }
}

async function sendAdminBillingTestSequence() {
  if (adminBillingBusy) return;
  const toEmail = String(adminBillingTestEmailInput?.value || "").trim();
  if (!toEmail) {
    setAdminBillingStatus(tr("A valid test email is required."), { tone: "error" });
    return;
  }
  setAdminBillingBusy(true);
  setAdminBillingStatus("");
  try {
    const payload = await fetchApiWithAuth("/api/admin/invoices/send-test-sequence", {
      method: "POST",
      timeoutMs: 240000,
      body: JSON.stringify({
        toEmail,
      }),
    });
    const sent = Number(payload?.sent_count || 0);
    const failed = Number(payload?.failed_count || 0);
    const total = Number(payload?.total_count || sent + failed || 0);
    if (failed > 0) {
      const message = tr("Sequence sent with {failed} failures ({sent}/{total}).", {
        failed: String(failed),
        sent: String(sent),
        total: String(total),
      });
      setAdminBillingStatus(message, { tone: "error" });
      return;
    }
    const message = tr("Follow-up sequence sent to {email} ({count} emails).", {
      email: String(payload?.to || toEmail),
      count: String(total || sent),
    });
    setAdminBillingStatus(message, { tone: "success" });
  } catch (error) {
    setAdminBillingStatus(error?.message || tr("Could not send follow-up sequence."), {
      tone: "error",
    });
  } finally {
    setAdminBillingBusy(false);
  }
}

async function sendAdminReportsTestEmail() {
  if (adminBillingBusy) return;
  const toEmail = String(adminBillingTestEmailInput?.value || "").trim();
  if (!toEmail) {
    setAdminBillingStatus(tr("A valid test email is required."), { tone: "error" });
    return;
  }
  setAdminBillingBusy(true);
  setAdminBillingStatus("");
  try {
    const payload = await fetchApiWithAuth("/api/admin/reports/send-test", {
      method: "POST",
      timeoutMs: 30000,
      body: JSON.stringify({ toEmail }),
    });
    const amount = Number(payload?.profitAmount || 0);
    setAdminBillingStatus(
      tr("Reports test email sent to {email} ({amount}).", {
        email: String(payload?.to || toEmail),
        amount: formatMoney(amount),
      }),
      { tone: "success" }
    );
  } catch (error) {
    setAdminBillingStatus(error?.message || tr("Could not send reports test email."), {
      tone: "error",
    });
  } finally {
    setAdminBillingBusy(false);
  }
}

async function previewAdminAgreementEmail() {
  if (adminBillingBusy) return;
  const toEmail = String(adminBillingTestEmailInput?.value || "").trim();
  if (!toEmail) {
    setAdminBillingStatus(tr("A valid test email is required."), { tone: "error" });
    return;
  }
  const previewWindow = window.open("", "_blank");
  if (!previewWindow) {
    setAdminBillingStatus(tr("Allow pop-ups to preview the agreement email."), {
      tone: "error",
    });
    return;
  }
  previewWindow.document.write("<!doctype html><title>Loading...</title><body style=\"margin:0;background:#050913\"></body>");
  previewWindow.document.close();
  setAdminBillingBusy(true);
  setAdminBillingStatus("");
  try {
    const payload = await fetchApiWithAuth("/api/admin/agreements/preview-test", {
      method: "POST",
      body: JSON.stringify({ toEmail }),
    });
    previewWindow.document.open();
    previewWindow.document.write(String(payload?.html || ""));
    previewWindow.document.close();
    setAdminBillingStatus(
      tr("Agreement email preview opened for {email}.", {
        email: String(payload?.to || toEmail),
      }),
      { tone: "success" }
    );
  } catch (error) {
    previewWindow.close();
    setAdminBillingStatus(error?.message || tr("Could not preview agreement email."), {
      tone: "error",
    });
  } finally {
    setAdminBillingBusy(false);
  }
}

async function sendAdminAgreementTestEmail() {
  if (adminBillingBusy) return;
  const toEmail = String(adminBillingTestEmailInput?.value || "").trim();
  if (!toEmail) {
    setAdminBillingStatus(tr("A valid test email is required."), { tone: "error" });
    return;
  }
  setAdminBillingBusy(true);
  setAdminBillingStatus("");
  try {
    const payload = await fetchApiWithAuth("/api/admin/agreements/send-test", {
      method: "POST",
      timeoutMs: 45000,
      body: JSON.stringify({ toEmail }),
    });
    setAdminBillingStatus(
      tr("Agreement test email sent to {email}.", {
        email: String(payload?.to || toEmail),
      }),
      { tone: "success" }
    );
  } catch (error) {
    setAdminBillingStatus(error?.message || tr("Could not send test agreement email."), {
      tone: "error",
    });
  } finally {
    setAdminBillingBusy(false);
  }
}

async function sendAdminInvoice(invoiceId) {
  const safeInvoiceId = String(invoiceId || "").trim();
  if (!safeInvoiceId || adminBillingBusy) return;
  setAdminBillingBusy(true);
  setAdminBillingStatus("");
  try {
    await sendAdminInvoiceWithRenderer(safeInvoiceId);
    setAdminBillingStatus(tr("Invoice sent."), { tone: "success" });
    await loadAdminInvoices({ quiet: true });
    await loadAdminDashboard({ quiet: true });
  } catch (error) {
    setAdminBillingStatus(error?.message || tr("Could not send invoice."), { tone: "error" });
  } finally {
    setAdminBillingBusy(false);
  }
}

async function markAdminInvoicePaid(invoiceId) {
  const safeInvoiceId = String(invoiceId || "").trim();
  if (!safeInvoiceId || adminBillingBusy) return;
  setAdminBillingBusy(true);
  setAdminBillingStatus("");
  try {
    await fetchApiWithAuth("/api/admin/invoices/mark-paid", {
      method: "POST",
      body: JSON.stringify({ invoiceId: safeInvoiceId }),
    });
    setAdminBillingStatus(tr("Invoice marked as paid."), { tone: "success" });
    await loadAdminInvoices({ quiet: true });
    await loadAdminDashboard({ quiet: true });
    await loadBillingOverview({ quiet: true });
  } catch (error) {
    setAdminBillingStatus(error?.message || tr("Could not mark invoice as paid."), {
      tone: "error",
    });
  } finally {
    setAdminBillingBusy(false);
  }
}

function updateShopifyProviderStatus() {
  setShopifyProviderConnectedState(Boolean(currentUser && shopifyConnection));
  if (!currentUser) {
    setProviderStatus("", { persist: true });
    return;
  }
  setProviderStatus("", { persist: true });
}

async function loadShopifyConnectionStatus(options = {}) {
  const { quiet = false } = options;
  if (!currentUser) {
    shopifyConnection = null;
    shopifySavedLocationSelection = [];
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
    shopifySavedLocationSelection = [];
    updateShopifyProviderStatus();
    return shopifyConnection;
  } catch (error) {
    shopifyConnection = null;
    shopifySavedLocationSelection = [];
    updateShopifyProviderStatus();
    if (!quiet) {
      setProviderStatus(error.message || tr("Could not load Shopify connection."), {
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
      setProviderStatus(tr("Shopify connected: {shop}", { shop }), { kind: "success" });
    } else {
      setProviderStatus(tr("Shopify connected."), { kind: "success" });
    }
  } else if (status === "error") {
    const message = String(params.get("message") || "").trim() || tr("Could not connect Shopify.");
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

function applyImportedRows(rows, sourceLabel, options = {}) {
  const { source = "provider" } = options;
  if (!Array.isArray(rows) || rows.length === 0) {
    throw new Error(tr("No usable {sourceLabel} rows found.", { sourceLabel }));
  }
  clearBatchState();
  state.csvRows = rows;
  state.csvPage = 1;
  state.csvValidationAttempted = false;
  setCsvBatchSource(source);
  setCsvMode(true);
  if (!state.shipFromLockedByProvider) {
    syncCsvRowsWithSelectedOrigin({ rerender: false });
  }
  setCsvEditMode(false);
  setInlineFormErrorToast(labelError, "");
  updatePreview();
  updateSummary();
  updatePayment();
}

async function beginShopifyInstall() {
  const defaultShop = shopifyConnection?.shop || "";
  const typed = window.prompt(
    tr("Enter your Shopify store domain (example: your-store.myshopify.com)"),
    defaultShop
  );
  if (typed === null) return;
  const shop = normalizeShopDomain(typed);
  if (!shop) {
    setProviderStatus(tr("Enter a valid .myshopify.com domain."), { kind: "error" });
    return;
  }
  try {
    setProviderStatus(tr("Connecting Shopify..."), { persist: true });
    const data = await fetchApiWithAuth("/api/shopify/install-link", {
      method: "POST",
      body: JSON.stringify({ shop }),
    });
    if (!data || typeof data !== "object" || !data.url) {
      throw new Error(tr("Shopify install URL was not returned."));
    }
    window.location.assign(String(data.url));
  } catch (error) {
    setProviderStatus(error.message || tr("Could not start Shopify connect flow."), {
      kind: "error",
    });
  }
}

async function importShopifyOrders(shop) {
  const normalizedShop = normalizeShopDomain(shop);
  if (!normalizedShop) {
    setProviderStatus(tr("Connect Shopify before importing."), { kind: "error" });
    return;
  }

  try {
    setProviderStatus(
      tr("Importing orders from {shop}...", {
        shop: normalizedShop,
      }),
      { persist: true }
    );
    const data = await fetchApiWithAuth("/api/shopify/import-orders", {
      method: "POST",
      body: JSON.stringify({
        shop: normalizedShop,
        limit: 50,
      }),
    });
    const rows = mapShopifyImportRows(data?.rows);
    applyImportedRows(rows, "Shopify", { source: "provider-shopify" });
    setProviderStatus(
      tr("Imported {count} orders from {shop}.", {
        count: rows.length,
        shop: normalizedShop,
      }),
      { kind: "success" }
    );
    const triggerText = providerTrigger?.querySelector("span");
    if (triggerText) {
      triggerText.textContent = tr("Imported {count} Shopify orders", {
        count: rows.length,
      });
      window.setTimeout(() => {
        triggerText.textContent = tr("Import from provider");
      }, 2200);
    }
  } catch (error) {
    const message = String(error?.message || tr("Shopify import failed."));
    if (
      /connection expired|reconnect shopify|invalid api key or access token|unrecognized login|wrong password/i.test(
        message
      )
    ) {
      shopifyConnection = null;
      shopifySavedLocationSelection = [];
      updateShopifyProviderStatus();
      setProviderStatus(
        tr("Shopify connection expired. Click Shopify again to reconnect."),
        {
          kind: "error",
          persist: true,
        }
      );
      return;
    }
    setProviderStatus(message, { kind: "error" });
  }
}

async function handleShopifyProviderAction() {
  if (!currentUser) {
    setProviderStatus(tr("Sign in before importing from Shopify."), { kind: "error" });
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

const mockAccountManagers = [
  "Lea Martin",
  "Noah Peeters",
  "Eva Janssens",
  "Milan Dubois",
];
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

function raceWarehouseSupabaseRequest(promise, timeoutMs = 12000) {
  const safeTimeout = Math.max(3000, Number(timeoutMs) || 12000);
  let timeoutId = 0;
  const timeoutPromise = new Promise((_, reject) => {
    timeoutId = window.setTimeout(() => {
      reject(new Error(tr("Request timed out. Please try again.")));
    }, safeTimeout);
  });
  return Promise.race([promise, timeoutPromise]).finally(() => {
    window.clearTimeout(timeoutId);
  });
}

async function fetchSupabaseHistoryRows(userId) {
  if (!supabaseClient || !userId) {
    return { data: [], error: new Error(tr("Supabase unavailable")) };
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

    lastError = error || new Error(tr("Unknown history fetch error"));
  }

  return { data: [], error: lastError };
}

function formatHistoryDate(value) {
  if (!value) return "--";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "--";
  return date.toLocaleString(getUiLocale(), {
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
    .toLocaleString(getUiLocale(), { month: "short" })
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

function createLeadProspectRepository(seedRows = MOCK_LEAD_PROSPECTS) {
  const records = Array.isArray(seedRows)
    ? seedRows.map((row, index) => ({
        ...row,
        orderIndex: index,
        disposition: LEAD_OUTCOME_META[row?.disposition] ? row.disposition : "new",
        last_called_at: String(row?.last_called_at || ""),
      }))
    : [];

  const store = new Map(records.map((row) => [String(row.id || ""), row]));

  return {
    async list() {
      await new Promise((resolve) => window.setTimeout(resolve, 110));
      return Array.from(store.values())
        .sort((left, right) => Number(left.orderIndex || 0) - Number(right.orderIndex || 0))
        .map((row) => ({ ...row }));
    },
    async updateDisposition(id, disposition) {
      const safeId = String(id || "").trim();
      const safeDisposition = LEAD_OUTCOME_META[disposition] ? disposition : "new";
      await new Promise((resolve) => window.setTimeout(resolve, 90));
      const current = store.get(safeId);
      if (!current) {
        throw new Error(tr("Lead not found."));
      }
      const next = {
        ...current,
        disposition: safeDisposition,
        last_called_at: new Date().toISOString(),
      };
      store.set(safeId, next);
      return { ...next };
    },
  };
}

function getLeadStackMeta(stack) {
  return LEAD_STACK_META[String(stack || "").trim().toLowerCase()] || LEAD_STACK_META.shopify;
}

function getLeadOutcomeMeta(outcome) {
  return LEAD_OUTCOME_META[String(outcome || "").trim().toLowerCase()] || LEAD_OUTCOME_META.new;
}

function formatLeadStoreSince(value) {
  if (!value) return "--";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "--";
  return date.toLocaleString(getUiLocale(), {
    month: "short",
    year: "numeric",
  });
}

function formatLeadParcels(value) {
  const count = Math.max(0, Number(value || 0));
  return count.toLocaleString(getUiLocale());
}

function getLeadCallHref(phone) {
  const safePhone = String(phone || "")
    .trim()
    .replace(/[^\d+]/g, "");
  return safePhone ? `tel:${safePhone}` : "";
}

function findLeadProspectById(leadId) {
  const safeId = String(leadId || "").trim();
  return leadProspects.find((lead) => String(lead?.id || "") === safeId) || null;
}

function renderLeadSummaryCards() {
  const counts = {
    ready: 0,
    followUp: 0,
    notInterested: 0,
  };
  leadProspects.forEach((lead) => {
    const bucket = getLeadOutcomeMeta(lead?.disposition).summaryBucket;
    if (bucket === "followUp") {
      counts.followUp += 1;
      return;
    }
    if (bucket === "notInterested") {
      counts.notInterested += 1;
      return;
    }
    counts.ready += 1;
  });
  if (leadsReadyCount) leadsReadyCount.textContent = String(counts.ready);
  if (leadsFollowUpCount) leadsFollowUpCount.textContent = String(counts.followUp);
  if (leadsNotInterestedCount) leadsNotInterestedCount.textContent = String(counts.notInterested);
}

function getFilteredLeadProspects() {
  const query = String(leadSearchQuery || "").trim().toLowerCase();
  const stack = String(leadStackFilterValue || "all").trim().toLowerCase();
  return leadProspects.filter((lead) => {
    const matchesStack = stack === "all" || String(lead?.techStack || "").trim().toLowerCase() === stack;
    if (!matchesStack) return false;
    if (!query) return true;
    const haystack = [
      lead?.name,
      lead?.phone,
      lead?.companyName,
      getLeadStackMeta(lead?.techStack).label,
      formatLeadStoreSince(lead?.storeOnlineSince),
    ]
      .join(" ")
      .toLowerCase();
    return haystack.includes(query);
  });
}

function renderLeadProspects() {
  renderLeadSummaryCards();
  if (!leadsTableBody || !leadsEmpty || !leadsTableWrap || !leadsStatus) return;
  if (reloadLeadsButton) {
    reloadLeadsButton.disabled = leadProspectsLoading;
  }

  const filtered = getFilteredLeadProspects();
  leadsTableBody.innerHTML = "";

  if (leadProspectsLoading) {
    leadsStatus.textContent = tr("Loading leads...");
    if (!leadProspects.length) {
      leadsEmpty.classList.add("is-hidden");
      leadsTableWrap.classList.add("is-hidden");
      return;
    }
  } else if (!leadProspectsLoaded) {
    leadsStatus.textContent = tr("Lead queue unavailable.");
  } else if (filtered.length !== leadProspects.length) {
    leadsStatus.textContent = tr("Showing {shown} of {total} leads.", {
      shown: filtered.length,
      total: leadProspects.length,
    });
  } else {
    leadsStatus.textContent = tr("{count} leads ready in the queue.", {
      count: leadProspects.length,
    });
  }

  if (!filtered.length) {
    leadsEmpty.classList.remove("is-hidden");
    leadsTableWrap.classList.add("is-hidden");
    return;
  }

  leadsEmpty.classList.add("is-hidden");
  leadsTableWrap.classList.remove("is-hidden");

  filtered.forEach((lead) => {
    const stack = getLeadStackMeta(lead?.techStack);
    const outcome = getLeadOutcomeMeta(lead?.disposition);
    const row = document.createElement("tr");
    row.className = "leads-row";
    row.innerHTML = `
      <td>
        <div class="lead-prospect">
          <div class="lead-prospect-name">${escapeHtml(String(lead?.name || "--"))}</div>
          <div class="lead-prospect-phone mono">${escapeHtml(String(lead?.phone || "--"))}</div>
        </div>
      </td>
      <td class="mono leads-table-center">${escapeHtml(String(lead?.age || "--"))}</td>
      <td>
        <div class="lead-company">${escapeHtml(String(lead?.companyName || "--"))}</div>
      </td>
      <td>
        <span class="lead-stack-pill ${stack.className}">
          <img src="${escapeHtml(stack.icon)}" alt="" class="lead-stack-pill-logo" />
          <span>${escapeHtml(stack.label)}</span>
        </span>
      </td>
      <td class="mono">${escapeHtml(formatLeadStoreSince(lead?.storeOnlineSince))}</td>
      <td class="mono leads-table-metric">${escapeHtml(formatMoney(lead?.estimatedMonthlyRevenue || 0))}</td>
      <td class="mono leads-table-metric">${escapeHtml(formatLeadParcels(lead?.estimatedParcelsPerMonth || 0))}</td>
      <td>
        <span class="lead-outcome-pill ${outcome.className}">${escapeHtml(outcome.label)}</span>
      </td>
      <td class="leads-table-actions">
        <button type="button" class="btn btn-secondary btn-sm lead-call-button" data-lead-call="${escapeHtml(
          String(lead?.id || "")
        )}">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M22 16.9v3a2 2 0 0 1-2.2 2 19.8 19.8 0 0 1-8.6-3.1A19.4 19.4 0 0 1 5.2 12.8 19.8 19.8 0 0 1 2.1 4.1 2 2 0 0 1 4 1.9h3a2 2 0 0 1 2 1.7l.5 3a2 2 0 0 1-.6 1.8L7 10.3a16 16 0 0 0 6.7 6.7l1.9-1.9a2 2 0 0 1 1.8-.6l3 .5a2 2 0 0 1 1.6 1.9z"/></svg>
          <span>Call</span>
        </button>
      </td>
    `;
    leadsTableBody.appendChild(row);
  });
}

function setLeadCallOutcomeBusy(isBusy) {
  leadCallOutcomeSaving = Boolean(isBusy);
  if (leadCallOutcomeActions) {
    leadCallOutcomeActions.querySelectorAll("[data-lead-outcome]").forEach((button) => {
      if (button instanceof HTMLButtonElement) {
        button.disabled = leadCallOutcomeSaving;
      }
    });
  }
  if (leadCallOutcomeClose) leadCallOutcomeClose.disabled = leadCallOutcomeSaving;
  if (leadCallOutcomeCancel) leadCallOutcomeCancel.disabled = leadCallOutcomeSaving;
}

function populateLeadCallOutcomeModal(lead) {
  if (leadCallOutcomeLeadName) {
    leadCallOutcomeLeadName.textContent = String(lead?.name || "--");
  }
  if (leadCallOutcomeLeadCompany) {
    leadCallOutcomeLeadCompany.textContent = String(lead?.companyName || "--");
  }
  if (leadCallOutcomeLeadPhone) {
    leadCallOutcomeLeadPhone.textContent = String(lead?.phone || "--");
  }
}

function openLeadCallOutcomeForLead(lead) {
  if (!lead) return;
  leadCallOutcomeLeadId = String(lead.id || "");
  populateLeadCallOutcomeModal(lead);
  setLeadCallOutcomeBusy(false);
  setLeadCallOutcomeModalOpen(true);
}

function closeLeadCallOutcome() {
  if (leadCallOutcomeSaving) return;
  leadCallOutcomeLeadId = "";
  setLeadCallOutcomeModalOpen(false);
}

async function saveLeadCallOutcome(outcome) {
  const safeOutcome = String(outcome || "").trim();
  if (!leadCallOutcomeLeadId || !LEAD_OUTCOME_META[safeOutcome]) return;
  setLeadCallOutcomeBusy(true);
  try {
    const updatedLead = await leadProspectRepository.updateDisposition(leadCallOutcomeLeadId, safeOutcome);
    leadProspects = leadProspects.map((lead) =>
      String(lead?.id || "") === String(updatedLead?.id || "") ? updatedLead : lead
    );
    renderLeadProspects();
    leadCallOutcomeLeadId = "";
    setLeadCallOutcomeModalOpen(false);
    showToast(
      safeOutcome === "not_interested"
        ? tr("Lead marked as not interested.")
        : tr("Follow-up queued for this lead."),
      { tone: "success" }
    );
  } catch (error) {
    showToast(error?.message || tr("Could not save the lead outcome."), { tone: "error" });
  } finally {
    setLeadCallOutcomeBusy(false);
  }
}

function launchLeadCall(lead) {
  const href = getLeadCallHref(lead?.phone);
  if (!href) {
    showToast(tr("Phone number unavailable for this lead."), { tone: "error" });
    return;
  }
  const link = document.createElement("a");
  link.href = href;
  link.setAttribute("aria-hidden", "true");
  link.style.position = "absolute";
  link.style.left = "-9999px";
  document.body.appendChild(link);
  link.click();
  link.remove();
}

function handleLeadCallAction(leadId) {
  const lead = findLeadProspectById(leadId);
  if (!lead) return;
  launchLeadCall(lead);
  window.setTimeout(() => {
    openLeadCallOutcomeForLead(lead);
  }, 120);
}

async function loadLeadProspects(options = {}) {
  const { quiet = false, force = false } = options;
  if (!currentUser) {
    leadProspects = [];
    leadProspectsLoaded = false;
    leadProspectsLoading = false;
    renderLeadProspects();
    return [];
  }
  const hasAccess = adminAccessAllowed || (await loadAdminAccessStatus({ quiet: true }));
  if (!hasAccess) {
    leadProspects = [];
    leadProspectsLoaded = false;
    leadProspectsLoading = false;
    renderLeadProspects();
    if (currentMainView === "leads") {
      setLeadsPageVisible(false, { replace: true });
    }
    if (!quiet) {
      showToast(tr("You are not allowed to access the leads dashboard."), { tone: "error" });
    }
    return [];
  }
  if (!force && leadProspectsLoadPromise) {
    return leadProspectsLoadPromise;
  }

  const requestToken = ++leadProspectsLoadRequestToken;
  leadProspectsLoading = true;
  renderLeadProspects();

  let requestPromise = null;
  requestPromise = (async () => {
    try {
      const rows = await leadProspectRepository.list();
      if (requestToken !== leadProspectsLoadRequestToken) {
        return leadProspects;
      }
      leadProspects = Array.isArray(rows) ? rows : [];
      leadProspectsLoaded = true;
      renderLeadProspects();
      return leadProspects;
    } catch (error) {
      if (requestToken === leadProspectsLoadRequestToken) {
        leadProspects = [];
        leadProspectsLoaded = false;
        renderLeadProspects();
        if (!quiet) {
          showToast(error?.message || tr("Could not load leads."), { tone: "error" });
        }
      }
      return [];
    } finally {
      if (requestToken === leadProspectsLoadRequestToken) {
        leadProspectsLoading = false;
        renderLeadProspects();
      }
      if (leadProspectsLoadPromise === requestPromise) {
        leadProspectsLoadPromise = null;
      }
    }
  })();
  leadProspectsLoadPromise = requestPromise;
  return requestPromise;
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
  const metadata = user?.user_metadata || {};
  const metadataProfile = {
    companyName: String(metadata.company_name || metadata.companyName || "").trim(),
    contactName: String(metadata.contact_name || metadata.contactName || "").trim(),
    contactEmail: String(metadata.contact_email || metadata.contactEmail || "").trim(),
    contactPhone: String(metadata.contact_phone || metadata.contactPhone || "").trim(),
    billingAddress: String(metadata.billing_address || metadata.billingAddress || "").trim(),
    taxId: String(metadata.tax_id || metadata.taxId || "").trim(),
    customerId: String(metadata.customer_id || metadata.customerId || "").trim(),
    accountManager: String(
      metadata.account_manager || metadata.accountManager || ""
    ).trim(),
  };
  const hasMetadataProfile = Object.values(metadataProfile).some((value) =>
    Boolean(String(value || "").trim())
  );
  if (hasMetadataProfile) {
    return {
      companyName: metadataProfile.companyName || "--",
      contactName: metadataProfile.contactName || "--",
      contactEmail: metadataProfile.contactEmail || email || "--",
      contactPhone: metadataProfile.contactPhone || "--",
      billingAddress: metadataProfile.billingAddress || "--",
      taxId: metadataProfile.taxId || "--",
      customerId: metadataProfile.customerId || "--",
      accountManager: metadataProfile.accountManager || pickBySeed(mockAccountManagers, seed, 6),
    };
  }

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
    accountManager: pickBySeed(mockAccountManagers, seed, 6),
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
        accountManager: profile.accountManager,
      }
    : {
        companyName: "--",
        contactName: "--",
        contactEmail: "--",
        contactPhone: "--",
        billingAddress: "--",
        taxId: "--",
        customerId: "--",
        accountManager: "--",
      };

  if (accountCompanyName) accountCompanyName.textContent = values.companyName;
  if (accountContactName) accountContactName.textContent = values.contactName;
  if (accountContactEmail) accountContactEmail.textContent = values.contactEmail;
  if (accountContactPhone) accountContactPhone.textContent = values.contactPhone;
  if (accountBillingAddress) accountBillingAddress.textContent = values.billingAddress;
  if (accountTaxId) accountTaxId.textContent = values.taxId;
  if (accountCustomerId) accountCustomerId.textContent = values.customerId;
  if (accountAccountManager) accountAccountManager.textContent = values.accountManager;

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
  const sameReturnAddress =
    record?.sameReturnAddress !== undefined
      ? Boolean(record.sameReturnAddress)
      : record?.same_return_address !== undefined
        ? Boolean(record.same_return_address)
        : fallback.sameReturnAddress !== undefined
          ? Boolean(fallback.sameReturnAddress)
          : true;

  const returnCountryRaw =
    record?.returnCountry ??
    record?.return_country ??
    fallback.returnCountry ??
    (sameReturnAddress ? country : "");
  const returnCountry =
    String(returnCountryRaw || "").trim() || (sameReturnAddress ? country : "");

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
    sameReturnAddress,
    returnSenderName: String(
      record?.returnSenderName ??
        record?.return_sender_name ??
        fallback.returnSenderName ??
        (sameReturnAddress ? record?.senderName ?? fallback.senderName ?? "" : "")
    ).trim(),
    returnStreet: String(
      record?.returnStreet ??
        record?.return_street ??
        fallback.returnStreet ??
        (sameReturnAddress ? record?.street ?? fallback.street ?? "" : "")
    ).trim(),
    returnCity: String(
      record?.returnCity ??
        record?.return_city ??
        fallback.returnCity ??
        (sameReturnAddress ? record?.city ?? fallback.city ?? "" : "")
    ).trim(),
    returnRegion: String(
      record?.returnRegion ??
        record?.return_region ??
        record?.returnState ??
        record?.return_state ??
        fallback.returnRegion ??
        (sameReturnAddress ? record?.region ?? fallback.region ?? "" : "")
    ).trim(),
    returnPostalCode: String(
      record?.returnPostalCode ??
        record?.return_postal_code ??
        record?.returnZip ??
        record?.return_zip ??
        fallback.returnPostalCode ??
        (sameReturnAddress ? record?.postalCode ?? fallback.postalCode ?? "" : "")
    ).trim(),
    returnCountry,
    isDefault: Boolean(record?.isDefault ?? record?.is_default ?? fallback.isDefault),
  };
}

function normalizeWarehouseRecords(records) {
  const list = Array.isArray(records) ? records.slice(0, WAREHOUSE_MAX_COUNT) : [];
  const normalized = list.map((record, index) =>
    normalizeWarehouseRecord(record, {
      name: `Origin ${index + 1}`,
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
      name: "Origin 1",
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

function getWarehouseRecordById(warehouseId) {
  const id = String(warehouseId || "").trim();
  if (!id) return null;
  return warehouseRecords.find((origin) => origin.id === id) || null;
}

function getSelectedWarehouseRecord() {
  const selected = getWarehouseRecordById(state.shipFromOriginId);
  if (selected) return selected;
  return getDefaultWarehouseRecord();
}

function ensureSelectedWarehouseOrigin() {
  const selected = getSelectedWarehouseRecord();
  state.shipFromOriginId = selected?.id || "";
  return selected;
}

function getWarehouseSenderValues(origin) {
  return {
    senderName: String(origin?.senderName || origin?.name || "").trim(),
    senderStreet: String(origin?.street || "").trim(),
    senderCity: String(origin?.city || "").trim(),
    senderState: String(origin?.region || "").trim(),
    senderZip: String(origin?.postalCode || "").trim(),
  };
}

function applyWarehouseSenderToRows(rows, origin) {
  const sender = getWarehouseSenderValues(origin);
  return rows.map((row) => ({
    ...row,
    senderName: sender.senderName,
    senderStreet: sender.senderStreet,
    senderCity: sender.senderCity,
    senderState: sender.senderState,
    senderZip: sender.senderZip,
  }));
}

function formatWarehouseAddressLine(origin, options = {}) {
  const { useReturn = false } = options;
  const name = String(
    useReturn
      ? origin?.returnSenderName || ""
      : origin?.senderName || origin?.name || ""
  ).trim();
  const street = String(useReturn ? origin?.returnStreet || "" : origin?.street || "").trim();
  const city = String(useReturn ? origin?.returnCity || "" : origin?.city || "").trim();
  const region = String(useReturn ? origin?.returnRegion || "" : origin?.region || "").trim();
  const postalCode = String(
    useReturn ? origin?.returnPostalCode || "" : origin?.postalCode || ""
  ).trim();
  const country = String(useReturn ? origin?.returnCountry || "" : origin?.country || "").trim();
  return [name, street, formatCityRegionPostal(city, region, postalCode), country]
    .filter(Boolean)
    .join(" • ");
}

function buildOriginChoiceButton(origin, options = {}) {
  const { selected = false, disabled = false } = options;
  const button = document.createElement("button");
  button.type = "button";
  button.className = `origin-choice${selected ? " is-selected" : ""}${disabled ? " is-disabled" : ""}`;
  button.dataset.originId = origin.id;
  button.disabled = disabled;
  button.innerHTML = `
    <span class="origin-choice-check" aria-hidden="true"><span class="origin-choice-check-fill"></span></span>
    <span>
      <span class="origin-choice-name">${escapeHtml(
        origin.name || tr("Origin")
      )}</span>
      <span class="origin-choice-meta">${tr("Physical")}: ${escapeHtml(
        formatWarehouseAddressLine(origin)
      )}</span>
      <span class="origin-choice-meta">${tr("Return")}: ${
        origin.sameReturnAddress
          ? tr("Same as physical")
          : escapeHtml(formatWarehouseAddressLine(origin, { useReturn: true }) || tr("Not set"))
      }</span>
    </span>
  `;
  return button;
}

function renderSenderOriginSelector() {
  if (!senderOriginSelector || !senderOriginNote) return;
  senderOriginSelector.innerHTML = "";

  if (!currentUser) {
    senderOriginNote.textContent = tr("Sign in to use account shipping origins.");
    if (inputMap.senderName) inputMap.senderName.value = "";
    if (inputMap.senderStreet) inputMap.senderStreet.value = "";
    if (inputMap.senderCity) inputMap.senderCity.value = "";
    if (inputMap.senderState) inputMap.senderState.value = "";
    if (inputMap.senderZip) inputMap.senderZip.value = "";
    syncInfoState();
    return;
  }

  if (!warehouseRecords.length) {
    senderOriginNote.textContent = tr("Add at least one shipping origin in Account settings.");
    if (inputMap.senderName) inputMap.senderName.value = "";
    if (inputMap.senderStreet) inputMap.senderStreet.value = "";
    if (inputMap.senderCity) inputMap.senderCity.value = "";
    if (inputMap.senderState) inputMap.senderState.value = "";
    if (inputMap.senderZip) inputMap.senderZip.value = "";
    syncInfoState();
    return;
  }

  const selected = ensureSelectedWarehouseOrigin();
  if (selected) {
    applyWarehouseToSender(selected, { announce: false });
  }

  const locked = warehouseRecords.length === 1;
  warehouseRecords.forEach((origin) => {
    const button = buildOriginChoiceButton(origin, {
      selected: origin.id === state.shipFromOriginId,
      disabled: locked,
    });
    senderOriginSelector.appendChild(button);
  });

  senderOriginNote.textContent = locked
    ? tr("Sender is fixed to your only saved shipping origin.")
    : "";
}

function renderCsvShipFromSelector() {
  if (!csvShipFromCard || !csvShipFromSelector || !csvShipFromNote) return;
  csvShipFromCard.classList.toggle("is-hidden", !state.csvMode);
  if (!state.csvMode) {
    csvShipFromSelector.innerHTML = "";
    csvShipFromSelector.classList.remove("is-locked");
    csvShipFromNote.innerHTML = "";
    return;
  }

  csvShipFromSelector.innerHTML = "";
  const lockedByProvider = Boolean(state.shipFromLockedByProvider);
  csvShipFromSelector.classList.toggle("is-locked", lockedByProvider);
  if (lockedByProvider) {
    csvShipFromNote.innerHTML =
      `${tr("Ship from is controlled by your Shopify")} <button type="button" class="csv-ship-from-note-link" data-action="open-shopify-settings">${tr("fulfillment settings")}</button> ${tr("for this import.")}`;
    return;
  }

  csvShipFromNote.innerHTML = "";
  if (!warehouseRecords.length) {
    csvShipFromNote.textContent = tr("No saved ship-from origin. Add one in Account settings.");
    return;
  }

  ensureSelectedWarehouseOrigin();
  const singleOrigin = warehouseRecords.length === 1;

  warehouseRecords.forEach((origin) => {
    const button = buildOriginChoiceButton(origin, {
      selected: origin.id === state.shipFromOriginId,
      disabled: lockedByProvider || singleOrigin,
    });
    csvShipFromSelector.appendChild(button);
  });

  csvShipFromNote.textContent = singleOrigin
    ? tr("Using your only saved ship-from origin.")
    : tr("Choose which origin to apply to this batch.");
}

function syncCsvRowsWithSelectedOrigin(options = {}) {
  const { rerender = true } = options;
  if (!state.csvMode || state.shipFromLockedByProvider) return;
  if (!Array.isArray(state.csvRows) || !state.csvRows.length) return;
  const selected = ensureSelectedWarehouseOrigin();
  if (!selected) return;
  state.csvRows = applyWarehouseSenderToRows(state.csvRows, selected);
  if (rerender) {
    renderCsvTable();
    updatePreview();
    updateSummary();
  }
}

function setSelectedWarehouseOrigin(originId, options = {}) {
  const { syncCsv = true } = options;
  const origin = getWarehouseRecordById(originId);
  if (!origin) return;
  state.shipFromOriginId = origin.id;
  applyWarehouseToSender(origin, { announce: false });
  if (syncCsv) {
    syncCsvRowsWithSelectedOrigin({ rerender: true });
  }
  renderSenderOriginSelector();
  renderCsvShipFromSelector();
}

function maybeApplyDefaultWarehouseToSender() {
  const selected = ensureSelectedWarehouseOrigin();
  if (!selected) return;
  applyWarehouseToSender(selected, { announce: false });
}

function setCsvBatchSource(source = "none") {
  state.csvSource = String(source || "none");
  state.shipFromLockedByProvider = state.csvSource === "provider-shopify";
  renderCsvShipFromSelector();
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
    setWarehouseStatus(
      tr("Applied {name} to sender details.", {
        name: origin.name || tr("warehouse"),
      }),
      {
      tone: "success",
      }
    );
  }
}

function setWarehouseStatus(message, options = {}) {
  const { tone = "muted", toast = Boolean(message) } = options;
  const text = String(message || "").trim();
  if (warehouseStatusRow) {
    warehouseStatusRow.hidden = true;
  }
  if (warehouseStatus) {
    warehouseStatus.hidden = true;
    warehouseStatus.textContent = "";
    warehouseStatus.classList.remove("is-error", "is-success");
  }
  if (toast && text) {
    showStatusToast(text, { tone: tone === "muted" ? "info" : tone });
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
      label.textContent = warehouseSaving ? tr("Saving...") : tr("Save");
    }
  }
}

function setWarehouseDirty(isDirty, options = {}) {
  const { announce = true } = options;
  warehouseDirty = Boolean(isDirty);
  updateWarehouseControls();
  if (warehouseDirty && announce) {
    setWarehouseStatus(tr("Unsaved changes in shipping origins."));
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
  if (fieldKey === "country" || fieldKey === "returnCountry") {
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
    empty.textContent = tr("Sign in to manage account shipping origins.");
    warehouseList.appendChild(empty);
    updateWarehouseControls();
    renderSenderOriginSelector();
    renderCsvShipFromSelector();
    return;
  }

  if (!warehouseRecords.length) {
    const empty = document.createElement("div");
    empty.className = "warehouse-empty";
    empty.textContent = tr("No shipping origins configured yet.");
    warehouseList.appendChild(empty);
    updateWarehouseControls();
    renderSenderOriginSelector();
    renderCsvShipFromSelector();
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
    name.textContent = origin.name || `${tr("Origin")} ${index + 1}`;

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
    defaultText.textContent = tr("Default");

    defaultToggle.appendChild(defaultBox);
    defaultToggle.appendChild(defaultText);
    head.appendChild(defaultToggle);

    const grid = document.createElement("div");
    grid.className = "warehouse-grid";
    grid.appendChild(
      createWarehouseField(tr("Origin Label"), "name", origin.name, {
        placeholder: `${tr("Origin")} ${index + 1}`,
      })
    );
    grid.appendChild(
      createWarehouseField(tr("Sender Name"), "senderName", origin.senderName, {
        placeholder: tr("Sender/Company name"),
      })
    );
    grid.appendChild(createWarehouseField(tr("Street"), "street", origin.street, { wide: true }));
    grid.appendChild(createWarehouseField(tr("City"), "city", origin.city));
    grid.appendChild(createWarehouseField(tr("Region"), "region", origin.region));
    grid.appendChild(createWarehouseField(tr("Postal Code"), "postalCode", origin.postalCode));
    grid.appendChild(createWarehouseField(tr("Country"), "country", origin.country));

    const returnToggle = document.createElement("button");
    returnToggle.type = "button";
    returnToggle.className = `warehouse-return-toggle${origin.sameReturnAddress ? " is-active" : ""}`;
    returnToggle.dataset.warehouseAction = "toggle-return";
    returnToggle.disabled = warehouseSaving;

    const returnToggleBox = document.createElement("span");
    returnToggleBox.className = "warehouse-return-box";
    returnToggleBox.setAttribute("aria-hidden", "true");

    const returnToggleText = document.createElement("span");
    returnToggleText.textContent = tr("Same Return Address");

    returnToggle.appendChild(returnToggleBox);
    returnToggle.appendChild(returnToggleText);

    let returnGrid = null;
    if (!origin.sameReturnAddress) {
      returnGrid = document.createElement("div");
      returnGrid.className = "warehouse-return-grid";
      returnGrid.appendChild(
        createWarehouseField(tr("Return Name"), "returnSenderName", origin.returnSenderName, {
          placeholder: tr("Return sender/company"),
        })
      );
      returnGrid.appendChild(
        createWarehouseField(tr("Return Street"), "returnStreet", origin.returnStreet, {
          wide: true,
        })
      );
      returnGrid.appendChild(
        createWarehouseField(tr("Return City"), "returnCity", origin.returnCity)
      );
      returnGrid.appendChild(
        createWarehouseField(tr("Return Region"), "returnRegion", origin.returnRegion)
      );
      returnGrid.appendChild(
        createWarehouseField(
          tr("Return Postal Code"),
          "returnPostalCode",
          origin.returnPostalCode
        )
      );
      returnGrid.appendChild(
        createWarehouseField(tr("Return Country"), "returnCountry", origin.returnCountry)
      );
    }

    const controls = document.createElement("div");
    controls.className = "warehouse-card-actions";

    const controlGroup = document.createElement("div");
    controlGroup.className = "warehouse-card-controls";

    const saveButton = document.createElement("button");
    saveButton.type = "button";
    saveButton.className = "btn btn-primary btn-sm";
    saveButton.dataset.warehouseAction = "save";
    saveButton.textContent = warehouseSaving ? tr("Saving...") : tr("Save");
    saveButton.disabled = warehouseSaving;

    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.className = "btn btn-ghost btn-sm";
    removeButton.dataset.warehouseAction = "remove";
    removeButton.textContent = tr("Remove");
    removeButton.disabled = warehouseSaving || warehouseRecords.length <= 1;

    controlGroup.appendChild(saveButton);
    controlGroup.appendChild(removeButton);
    controls.appendChild(controlGroup);

    card.appendChild(head);
    card.appendChild(grid);
    card.appendChild(returnToggle);
    if (returnGrid) {
      card.appendChild(returnGrid);
    }
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
  renderSenderOriginSelector();
  renderCsvShipFromSelector();
  if (state.csvMode) {
    syncCsvRowsWithSelectedOrigin({ rerender: true });
  }
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
    name: `${tr("Origin")} ${nextNumber}`,
    senderName: senderNameSeed === "--" ? "" : senderNameSeed,
    country: "France",
    isDefault: warehouseRecords.length === 0,
  });
}

function validateWarehouseRecords(records) {
  if (!Array.isArray(records) || !records.length) {
    return { ok: false, message: tr("Add at least one warehouse origin."), records: [] };
  }

  const normalized = normalizeWarehouseRecords(records);
  for (let i = 0; i < normalized.length; i += 1) {
    const origin = normalized[i];
    const prefix = origin.name || `${tr("Origin")} ${i + 1}`;
    if (!origin.name) {
      return {
        ok: false,
        message: tr("{prefix}: origin label is required.", { prefix }),
        records: normalized,
      };
    }
    if (!origin.senderName) {
      return {
        ok: false,
        message: tr("{prefix}: sender name is required.", { prefix }),
        records: normalized,
      };
    }
    if (!origin.street) {
      return {
        ok: false,
        message: tr("{prefix}: street is required.", { prefix }),
        records: normalized,
      };
    }
    if (!origin.city) {
      return {
        ok: false,
        message: tr("{prefix}: city is required.", { prefix }),
        records: normalized,
      };
    }
    if (!origin.postalCode) {
      return {
        ok: false,
        message: tr("{prefix}: postal code is required.", { prefix }),
        records: normalized,
      };
    }
    if (!origin.country) {
      return {
        ok: false,
        message: tr("{prefix}: country is required.", { prefix }),
        records: normalized,
      };
    }
    if (!origin.sameReturnAddress) {
      if (!origin.returnSenderName) {
        return {
          ok: false,
          message: tr(
            "{prefix}: return sender name is required when using a custom return address.",
            { prefix }
          ),
          records: normalized,
        };
      }
      if (!origin.returnStreet) {
        return {
          ok: false,
          message: tr(
            "{prefix}: return street is required when using a custom return address.",
            { prefix }
          ),
          records: normalized,
        };
      }
      if (!origin.returnCity) {
        return {
          ok: false,
          message: tr("{prefix}: return city is required when using a custom return address.", {
            prefix,
          }),
          records: normalized,
        };
      }
      if (!origin.returnPostalCode) {
        return {
          ok: false,
          message: tr(
            "{prefix}: return postal code is required when using a custom return address.",
            { prefix }
          ),
          records: normalized,
        };
      }
      if (!origin.returnCountry) {
        return {
          ok: false,
          message: tr(
            "{prefix}: return country is required when using a custom return address.",
            { prefix }
          ),
          records: normalized,
        };
      }
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
    sameReturnAddress: origin.sameReturnAddress,
    returnSenderName: origin.returnSenderName,
    returnStreet: origin.returnStreet,
    returnCity: origin.returnCity,
    returnRegion: origin.returnRegion,
    returnPostalCode: origin.returnPostalCode,
    returnCountry: origin.returnCountry,
    isDefault: origin.isDefault,
  }));
}

async function fetchSupabaseWarehouseOrigins(userId) {
  if (!supabaseClient || !userId) {
    return { origins: [], error: new Error(tr("Supabase unavailable")) };
  }
  try {
    const { data, error } = await raceWarehouseSupabaseRequest(
      supabaseClient
        .from(ACCOUNT_SETTINGS_TABLE)
        .select("warehouse_origins")
        .eq("user_id", userId)
        .maybeSingle()
    );

    if (error) {
      return { origins: [], error };
    }
    const origins = Array.isArray(data?.warehouse_origins) ? data.warehouse_origins : [];
    return { origins, error: null };
  } catch (error) {
    return {
      origins: [],
      error: error instanceof Error ? error : new Error(tr("unknown error")),
    };
  }
}

async function saveSupabaseWarehouseOrigins(userId, origins) {
  if (!supabaseClient || !userId) {
    return { origins: [], error: new Error(tr("Supabase unavailable")) };
  }
  const payload = toWarehouseOriginsPayload(origins);
  try {
    const { data, error } = await raceWarehouseSupabaseRequest(
      supabaseClient
        .from(ACCOUNT_SETTINGS_TABLE)
        .upsert(
          {
            user_id: userId,
            warehouse_origins: payload,
          },
          { onConflict: "user_id" }
        )
        .select("warehouse_origins")
        .single()
    );

    if (error) {
      return { origins: [], error };
    }
    return {
      origins: Array.isArray(data?.warehouse_origins) ? data.warehouse_origins : payload,
      error: null,
    };
  } catch (error) {
    return {
      origins: [],
      error: error instanceof Error ? error : new Error(tr("unknown error")),
    };
  }
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
    state.shipFromOriginId = "";
    renderWarehouseList();
    setWarehouseStatus(tr("Sign in to manage shipping origins."));
    return;
  }

  if (!quiet) {
    setWarehouseStatus(tr("Loading shipping origins..."));
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
    setWarehouseStatus(tr("Shipping origins synced from your account."));
  } else if (source === "local") {
    setWarehouseStatus(tr("Showing browser-saved origins. Click Save to sync."));
  } else {
    setWarehouseStatus(tr("Add your shipping origin and click Save."));
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
  setWarehouseStatus(tr("Saving shipping origins..."));

  const userId = currentUser.id;
  let savedOrigins = toWarehouseOriginsPayload(warehouseRecords);
  try {
    if (!supabaseClient) {
      throw new Error(tr("Could not sync to account. Supabase client is unavailable."));
    }

    const { origins, error } = await saveSupabaseWarehouseOrigins(userId, savedOrigins);
    if (error) {
      throw error;
    }
    savedOrigins = Array.isArray(origins) && origins.length ? origins : savedOrigins;

    warehouseRecords = normalizeWarehouseRecords(savedOrigins);
    saveLocalWarehouses(userId, warehouseRecords);
    warehouseDirty = false;
    warehouseEnteringIds.clear();
    setWarehouseStatus(tr("Shipping origins saved to your account."), { tone: "success" });
    maybeApplyDefaultWarehouseToSender();
  } catch (error) {
    setWarehouseStatus(
      tr("Could not sync shipping origins: {error}", {
        error: error instanceof Error ? error.message || tr("unknown error") : tr("unknown error"),
      }),
      {
        tone: "error",
      }
    );
    setWarehouseDirty(true, { announce: false });
  } finally {
    warehouseSaving = false;
    renderWarehouseList();
  }
}

function resetWarehouseState() {
  warehouseLoadRequestToken += 1;
  warehouseRecords = [];
  warehouseDirty = false;
  warehouseSaving = false;
  warehouseEnteringIds.clear();
  state.shipFromOriginId = "";
  renderWarehouseList();
  setWarehouseStatus(tr("Sign in to manage shipping origins."));
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
  const { push = true, replace = false, animate = true, reportRange = "" } = options;
  const nextView =
    view === "account" ||
    view === "admin" ||
    view === "leads" ||
    view === "history" ||
    view === "reports" ||
    view === "builder"
      ? view
      : "builder";
  const viewChanged = nextView !== currentMainView;

  const applyView = () => {
    currentMainView = nextView;
    syncTopbarNavState(nextView);
    if (builderPage) {
      builderPage.classList.toggle("is-hidden", nextView !== "builder");
    }
    if (accountPageSection) {
      accountPageSection.classList.toggle("is-hidden", nextView !== "account");
    }
    if (adminPageSection) {
      adminPageSection.classList.toggle("is-hidden", nextView !== "admin");
    }
    if (leadsPageSection) {
      leadsPageSection.classList.toggle("is-hidden", nextView !== "leads");
    }
    if (historyPageSection) {
      historyPageSection.classList.toggle("is-hidden", nextView !== "history");
    }
    if (reportsPageSection) {
      reportsPageSection.classList.toggle("is-hidden", nextView !== "reports");
    }
    if (nextView !== "history") {
      setReceiptModalOpen(false);
    }
    if (nextView !== "builder") {
      setRestrictedGoodsModalOpen(false);
    }
    if (nextView !== "builder" && nextView !== "account") {
      setIbanTopupModalOpen(false);
    }
    if (nextView !== "admin") {
      setClientInviteHistoryModalOpen(false);
      setAdminSettingsModalOpen(false);
    }
    if (nextView !== "leads") {
      setLeadCallOutcomeModalOpen(false);
    }
    if (nextView === "reports") {
      applyReportRangeFromToken(reportRange || getReportRangeFromLocation(window.location));
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
    if (nextView === "admin") {
      updateRoute({ view: "admin" }, { replace });
      return;
    }
    if (nextView === "leads") {
      updateRoute({ view: "leads" }, { replace });
      return;
    }
    if (nextView === "history") {
      updateRoute({ view: "history" }, { replace });
      return;
    }
    if (nextView === "reports") {
      const reportRangeToken = String(reportRange || "").trim();
      updateRoute(
        reportRangeToken ? { view: "reports", reportRange: reportRangeToken } : { view: "reports" },
        { replace }
      );
      return;
    }
    updateRoute({ view: "builder", step: state.step }, { replace });
  }
}

function syncTopbarNavState(view = currentMainView) {
  const nextView =
    view === "account" ||
    view === "admin" ||
    view === "leads" ||
    view === "history" ||
    view === "reports" ||
    view === "builder"
      ? view
      : "builder";
  const navButtons = [
    [openBuilderPageButton, "builder"],
    [openAccountPageButton, "account"],
    [openLeadsPageButton, "leads"],
    [openAdminPageButton, "admin"],
    [openHistoryPageButton, "history"],
    [openReportsPageButton, "reports"],
  ];
  navButtons.forEach(([button, targetView]) => {
    if (!button) return;
    button.classList.toggle("is-active", nextView === targetView);
  });
}

function setAccountPageVisible(visible, options = {}) {
  setMainView(visible ? "account" : "builder", options);
}

function setAdminPageVisible(visible, options = {}) {
  setMainView(visible ? "admin" : "builder", options);
}

function setLeadsPageVisible(visible, options = {}) {
  setMainView(visible ? "leads" : "builder", options);
}

function setReportsPageVisible(visible, options = {}) {
  setMainView(visible ? "reports" : "builder", options);
}

function setHistoryPageVisible(visible, options = {}) {
  setMainView(visible ? "history" : "builder", options);
  queueHistoryPanelSync();
}

function setReceiptModalOpen(open) {
  if (!receiptModal) return;
  receiptModal.classList.toggle("is-closed", !open);
}

function setClientInviteHistoryModalOpen(open) {
  if (!clientInviteHistoryModal) return;
  clientInviteHistoryModal.classList.toggle("is-closed", !open);
}

function setAdminSettingsModalOpen(open) {
  if (!adminSettingsModal) return;
  adminSettingsModal.classList.toggle("is-closed", !open);
}

function setLeadCallOutcomeModalOpen(open) {
  if (!leadCallOutcomeModal) return;
  leadCallOutcomeModal.classList.toggle("is-closed", !open);
}

function syncHistoryPanelHeights() {
  if (!accountHistoryPanel || !accountPreviewPanel) return;
  if (historyPageSection?.classList.contains("is-hidden") || window.innerWidth <= 1120) {
    accountHistoryPanel.style.height = "";
    return;
  }
  accountHistoryPanel.style.height = "";
  const previewHeight = Math.round(accountPreviewPanel.getBoundingClientRect().height);
  if (previewHeight > 0) {
    accountHistoryPanel.style.height = `${previewHeight}px`;
  }
}

function queueHistoryPanelSync() {
  if (historyPanelSyncFrame) {
    window.cancelAnimationFrame(historyPanelSyncFrame);
  }
  historyPanelSyncFrame = window.requestAnimationFrame(() => {
    historyPanelSyncFrame = 0;
    syncHistoryPanelHeights();
  });
}

function setRestrictedGoodsModalOpen(open) {
  if (!restrictedGoodsModal) return;
  restrictedGoodsModal.classList.toggle("is-closed", !open);
}

function setIbanTopupStatus(message = "", options = {}) {
  const text = String(message || "").trim();
  const tone = String(options?.tone || "info").trim() || "info";
  const toast = options?.toast ?? Boolean(text);
  if (ibanTopupStatus) {
    ibanTopupStatus.hidden = true;
    ibanTopupStatus.textContent = "";
    ibanTopupStatus.classList.remove("is-error", "is-success");
  }
  if (toast && text) {
    showStatusToast(text, { tone });
  }
}

function syncIbanTopupActions() {
  const ready = Boolean(ibanTopupDraft) && !ibanTopupRequestInFlight;
  if (ibanCopyReference) {
    ibanCopyReference.disabled = !ready;
  }
  if (ibanCopyIban) {
    ibanCopyIban.disabled = !ready;
  }
  if (ibanTopupRequest) {
    ibanTopupRequest.disabled = ibanTopupRequestInFlight;
  }
}

function setIbanTopupLoading(loading, options = {}) {
  const isLoading = loading === true;
  if (ibanTopupLoadingTitle) {
    ibanTopupLoadingTitle.textContent =
      (typeof options?.title === "string" && options.title.trim())
      || tr("Preparing transfer instructions...");
  }
  if (ibanTopupLoadingSub) {
    ibanTopupLoadingSub.textContent =
      (typeof options?.subtitle === "string" && options.subtitle.trim())
      || tr("Generating your transfer reference and bank details.");
  }
  if (ibanTopupLoading) {
    ibanTopupLoading.classList.toggle("is-hidden", !isLoading);
  }
  if (isLoading && ibanTopupResult) {
    ibanTopupResult.classList.add("is-hidden");
    ibanTopupResult.classList.remove("is-revealed");
  }
  syncIbanTopupActions();
}

function resetIbanTopupResult(options = {}) {
  if (options?.clearDraft !== false) {
    ibanTopupDraft = null;
  }
  setIbanTopupLoading(false);
  if (ibanTopupResult) {
    ibanTopupResult.classList.add("is-hidden");
    ibanTopupResult.classList.remove("is-revealed");
  }
  if (ibanResultBeneficiary) ibanResultBeneficiary.textContent = "--";
  if (ibanResultIban) ibanResultIban.textContent = "--";
  if (ibanResultBic) ibanResultBic.textContent = "--";
  if (ibanResultReference) ibanResultReference.textContent = "--";
  if (ibanResultEta) ibanResultEta.textContent = "Instant / Up to 3 days";
  if (ibanResultNote) {
    ibanResultNote.textContent = tr("Transfers are credited once received (typically 1-2 business days).");
  }
  syncIbanTopupActions();
}

function setIbanTopupModalOpen(open, options = {}) {
  if (!ibanTopupModal) return;
  ibanTopupModal.classList.toggle("is-closed", !open);
  if (open) {
    const instructions = getIbanInstructionsFromOverview();
    if (ibanTopupModalNote) {
      ibanTopupModalNote.textContent =
        instructions.note ||
        tr("Generate a transfer reference, then send your bank transfer. Funds are credited once received.");
    }
    setIbanTopupStatus("");
    const reusableTopup = getReusableIbanTopupFromOverview();
    if (reusableTopup) {
      resetIbanTopupResult({ clearDraft: false });
      populateIbanTopupResult(reusableTopup);
      return;
    }
    resetIbanTopupResult();
    setIbanTopupLoading(true, {
      title: tr("Preparing transfer instructions..."),
      subtitle: tr("Generating your transfer reference and bank details."),
    });
    void createIbanTopupRequest({ silentToast: true, loadingMessage: tr("Preparing transfer instructions...") });
  } else {
    setIbanTopupLoading(false);
  }
}

function setWalletHistoryModalOpen(open) {
  if (!walletHistoryModal) return;
  walletHistoryModal.classList.toggle("is-closed", !open);
}

function populateIbanTopupResult(payload) {
  ibanTopupDraft = payload && typeof payload === "object" ? payload : ibanTopupDraft;
  const instructions = payload?.instructions && typeof payload.instructions === "object"
    ? payload.instructions
    : getIbanInstructionsFromOverview();
  const topup = payload?.topup && typeof payload.topup === "object" ? payload.topup : null;
  if (ibanResultBeneficiary) {
    ibanResultBeneficiary.textContent = String(instructions.beneficiary || "--");
  }
  if (ibanResultIban) {
    ibanResultIban.textContent = String(instructions.iban || "--");
  }
  if (ibanResultBic) {
    ibanResultBic.textContent = String(instructions.bic || "--");
  }
  if (ibanResultReference) {
    ibanResultReference.textContent = String(instructions.reference || topup?.reference || "--");
  }
  if (ibanResultEta) {
    ibanResultEta.textContent = "Instant / Up to 3 days";
  }
  if (ibanResultNote) {
    ibanResultNote.textContent =
      String(instructions.note || "").trim()
      || tr("Transfers are credited once received (typically 1-2 business days).");
  }
  setIbanTopupLoading(false);
  if (ibanTopupResult) {
    ibanTopupResult.classList.remove("is-hidden");
    ibanTopupResult.classList.remove("is-revealed");
    void ibanTopupResult.offsetWidth;
    ibanTopupResult.classList.add("is-revealed");
  }
  syncIbanTopupActions();
}

async function createIbanTopupRequest(options = {}) {
  const loadingMessage =
    typeof options?.loadingMessage === "string" && options.loadingMessage.trim()
      ? options.loadingMessage
      : tr("Creating top-up request...");
  if (ibanTopupRequestInFlight && ibanTopupRequestPromise) {
    setIbanTopupLoading(true, {
      title: loadingMessage,
      subtitle: tr("Generating your transfer reference and bank details."),
    });
    return ibanTopupRequestPromise;
  }
  ibanTopupRequestInFlight = true;
  syncIbanTopupActions();
  setIbanTopupLoading(true, {
    title: loadingMessage,
    subtitle: tr("Generating your transfer reference and bank details."),
  });
  setIbanTopupStatus(loadingMessage, { tone: "info", toast: false });
  ibanTopupRequestPromise = (async () => {
    try {
      const payload = await fetchApiWithAuth("/api/billing/topups/request", {
        method: "POST",
        body: JSON.stringify({}),
      });
      ibanTopupDraft = payload;
      populateIbanTopupResult(payload);
      setIbanTopupStatus(tr("Transfer reference generated. Use it as communication."), {
        tone: "success",
        toast: options?.silentToast ? false : true,
      });
      await loadBillingOverview({ quiet: true });
      return payload;
    } catch (error) {
      setIbanTopupLoading(false);
      syncIbanTopupActions();
      setIbanTopupStatus(error?.message || tr("Could not create top-up request."), {
        tone: "error",
        toast: true,
      });
      return null;
    } finally {
      ibanTopupRequestInFlight = false;
      ibanTopupRequestPromise = null;
      syncIbanTopupActions();
    }
  })();
  return ibanTopupRequestPromise;
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
    accountPreviewMeta.textContent = tr(
      "Select a generation to preview labels and receipt details."
    );
  }
  if (accountDownloadPdf) {
    accountDownloadPdf.disabled = true;
  }
  if (accountDownloadInvoice) {
    accountDownloadInvoice.disabled = true;
  }
  if (openReceiptModalButton) {
    openReceiptModalButton.disabled = true;
  }
  if (receiptDocument) {
    receiptDocument.innerHTML = "";
  }
  queueHistoryPanelSync();
}

function renderAccountHistoryList() {
  if (!accountHistoryList) return;
  accountHistoryList.innerHTML = "";

  if (!currentUser) {
    const empty = document.createElement("div");
    empty.className = "account-history-empty";
    empty.textContent = tr("Sign in to view previous generations.");
    accountHistoryList.appendChild(empty);
    return;
  }

  if (!historyRecords.length) {
    const empty = document.createElement("div");
    empty.className = "account-history-empty";
    empty.textContent = tr("No generations yet.");
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
    service.textContent = translateServiceName(record.service_type || "Label generation");

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
    meta.textContent = tr("{count} labels • Total {amount}", {
      count: totals.quantity,
      amount: formatMoney(totals.totalIncVat),
    });

    item.appendChild(head);
    item.appendChild(meta);
    accountHistoryList.appendChild(item);
  });
  queueHistoryPanelSync();
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
  const text = String(message || "").trim();
  if (accountHistoryStatus) {
    accountHistoryStatus.hidden = true;
    accountHistoryStatus.textContent = "";
    accountHistoryStatus.classList.add("is-hidden");
  }
  if (text) {
    showStatusToast(text, { tone: "info" });
  }
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
      setAccountHistoryStatus(tr("No preview-ready generations yet."));
      refreshReportsIfVisible();
      return;
    }

    selectAccountRecord(selectionIndex);
    refreshReportsIfVisible();
  };

  if (!currentUser) {
    historyRecords = [];
    historyStore = "supabase";
    setAccountHistoryStatus(tr("Sign in to view previous generations."));
    renderAccountHistoryList();
    resetAccountPreview();
    renderReportsDashboard();
    return;
  }

  setAccountHistoryStatus(tr("Loading previous generations..."));
  const previousRecordId = !preferLatest ? accountActiveRecord?.id : null;

  if (supabaseClient) {
    const { data, error } = await fetchSupabaseHistoryRows(requestedUserId);

    if (isStale()) return;
    if (!error && Array.isArray(data)) {
      historyStore = "supabase";
      saveServerHistoryCache(requestedUserId, data);
      applyLoadedRecords(
        data,
        tr("Select a generation to view PDFs and receipt details."),
        tr("No previous generations yet.")
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
        tr("Sync delayed. Showing last synced history."),
        tr("No previous generations yet.")
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
      tr("Showing browser-saved history for this account."),
      tr("No browser history saved for this account yet.")
    );
    return;
  }

  if (isStale()) return;
  if (priorRecords.length) {
    historyStore = "memory";
    applyLoadedRecords(
      priorRecords,
      tr("Sync delayed. Showing previously loaded history."),
      tr("No previous generations yet.")
    );
    return;
  }

  if (isStale()) return;
  historyStore = "local";
  applyLoadedRecords(
    [],
    tr("Showing browser-saved history for this account."),
    tr("Could not load history right now. Please refresh in a moment.")
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
    loadBillingOverview({ quiet: true });
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
      loadBillingOverview({ quiet: true });
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
  loadBillingOverview({ quiet: true });
}

function renderAccountBatchList() {
  if (!accountBatchPanel || !accountBatchList) return;
  if (accountLabels.length <= 1) {
    accountBatchPanel.classList.add("is-hidden");
    accountBatchList.innerHTML = "";
    if (accountBatchPreview) {
      accountBatchPreview.classList.add("is-single");
    }
    queueHistoryPanelSync();
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
      <div class="batch-index">${tr("Label")} ${label.index}</div>
      <div class="batch-meta mono">${label.trackingId}</div>
      <div class="batch-meta mono">${label.labelId}</div>
    `;
    accountBatchList.appendChild(button);
  });
  queueHistoryPanelSync();
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

function splitReceiptAddressLines(value) {
  const text = String(value || "").trim();
  if (!text) return ["--"];
  const parts = text
    .split(",")
    .map((part) => part.trim())
    .filter(Boolean);
  return parts.length ? parts : [text];
}

function splitInvoiceAddressLines(value) {
  const raw = String(value || "").trim();
  if (!raw) return ["--"];
  const directLines = raw
    .split(/\r?\n+/)
    .map((line) => line.trim())
    .filter(Boolean);
  if (directLines.length > 1) return directLines;
  const parts = raw
    .split(",")
    .map((part) => part.trim())
    .filter(Boolean);
  if (parts.length <= 2) return parts.length ? parts : ["--"];
  return [
    parts[0],
    parts.slice(1, -1).join(", "),
    parts.at(-1),
  ].filter(Boolean);
}

function formatReceiptDestination(data = {}) {
  const destination = [data.recipientCity, data.recipientCountry].filter(Boolean).join(", ");
  return destination || "--";
}

function formatReceiptParcel(data = {}) {
  const parts = [];
  if (data.packageWeight) {
    parts.push(`${data.packageWeight} kg`);
  }
  if (data.packageDims) {
    parts.push(data.packageDims);
  }
  return parts.length ? parts.join(" • ") : "--";
}

function buildReceiptTrackingId(rawId) {
  const raw = String(rawId || "").trim();
  if (!raw) {
    return {
      display: "--",
      slug: "label-order",
    };
  }

  const parts = raw.split("-").filter(Boolean);
  let suffix = "";
  let coreParts = parts;

  if (parts.length > 5 && /^\d+$/.test(parts[parts.length - 1])) {
    suffix = parts[parts.length - 1];
    coreParts = parts.slice(0, -1);
  }

  const compact = coreParts.join("").replace(/[^a-z0-9]/gi, "");
  if (!compact) {
    const safeSlug = raw
      .replace(/[^a-z0-9]+/gi, "-")
      .replace(/^-+|-+$/g, "")
      .toLowerCase();
    return {
      display: raw.toUpperCase(),
      slug: safeSlug || "label-order",
    };
  }

  const lead = compact.slice(0, 6).toUpperCase();
  const tail = compact.length > 6 ? compact.slice(-4).toUpperCase() : "";
  const displayParts = ["RCPT", lead];
  const slugParts = [lead.toLowerCase()];

  if (tail) {
    displayParts.push(tail);
    slugParts.push(tail.toLowerCase());
  }
  if (suffix) {
    displayParts.push(suffix.toUpperCase());
    slugParts.push(suffix.toLowerCase());
  }

  return {
    display: displayParts.join("-"),
    slug: slugParts.join("-"),
  };
}

function buildInvoiceTrackingId(rawId) {
  const receiptLike = buildReceiptTrackingId(rawId);
  const display = String(receiptLike.display || "--").replace(/^RCPT\b/, "INV");
  const slug = String(receiptLike.slug || "invoice").replace(/^receipt-/, "").replace(/^rcpt-/, "");
  return {
    display,
    slug,
  };
}

const INVOICE_ISSUER_PROFILE = Object.freeze({
  legalName: "Cryvelin LLC",
  brandName: "Shipide",
  descriptor: "Operates Shipide",
  jurisdiction: "Delaware, United States",
  email: "billing@shipide.com",
  iban: "BE71 0000 1111 2222",
});

function getHistoryInvoicePaymentMode() {
  const paymentMode = String(billingOverview?.payment_mode || "").trim().toLowerCase();
  if (["invoice", "card", "wallet", "hybrid"].includes(paymentMode)) {
    return paymentMode;
  }
  if (!billingOverview) {
    return "invoice";
  }
  const { invoiceEnabled, cardEnabled, walletEnabled } = getBillingFlags();
  if (invoiceEnabled) return "invoice";
  if (walletEnabled && !cardEnabled) return "wallet";
  if (cardEnabled) return "card";
  return "invoice";
}

function addDaysLocal(dateLike, days) {
  const source = new Date(dateLike);
  if (Number.isNaN(source.getTime())) return null;
  const next = new Date(source);
  next.setDate(next.getDate() + Number(days || 0));
  return next;
}

function startOfLocalDay(dateLike) {
  const date = new Date(dateLike);
  if (Number.isNaN(date.getTime())) return null;
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function invoiceRequiresManualSettlementClient(paymentMode = "") {
  return String(paymentMode || "").trim().toLowerCase() === "invoice";
}

function getInvoicePaymentMethodLabel(paymentMode = "") {
  const mode = String(paymentMode || "").trim().toLowerCase();
  if (mode === "invoice") return "Monthly billing";
  if (mode === "wallet") return "Wallet debit";
  return "Automatic collection";
}

function getInvoiceSettlementBadgeMeta({
  isMonthlyBilling,
  dueDate = null,
  paymentMode = "",
  reminderStage = 0,
} = {}) {
  if (!isMonthlyBilling) {
    if (paymentMode === "wallet") return { label: "Settled via wallet", tone: "success" };
    return { label: "Collected automatically", tone: "success" };
  }
  const stage = Math.max(0, Number(reminderStage) || 0);
  if (stage > 0) {
    const explicitStages = {
      1: { label: "Due in 15 days", tone: "warning" },
      2: { label: "Due in 7 days", tone: "warning" },
      3: { label: "Due tomorrow", tone: "attention" },
      4: { label: "Due today", tone: "danger" },
      5: { label: "3 days overdue", tone: "danger" },
      6: { label: "10 days overdue", tone: "danger" },
      7: { label: "15 days overdue", tone: "danger" },
      8: { label: "30 days overdue", tone: "danger" },
    };
    if (explicitStages[stage]) {
      return explicitStages[stage];
    }
  }
  const today = startOfLocalDay(new Date());
  const dueDay = startOfLocalDay(dueDate);
  if (!today || !dueDay) return { label: "Due in 30 days", tone: "success" };
  const diffDays = Math.round((dueDay.getTime() - today.getTime()) / 86400000);
  if (diffDays > 15) return { label: `Due in ${diffDays} days`, tone: "success" };
  if (diffDays > 1) return { label: `Due in ${diffDays} days`, tone: "warning" };
  if (diffDays === 1) return { label: "Due tomorrow", tone: "attention" };
  if (diffDays === 0) return { label: "Due today", tone: "danger" };
  if (diffDays === -1) return { label: "1 day overdue", tone: "danger" };
  return { label: `${Math.abs(diffDays)} days overdue`, tone: "danger" };
}

function getInvoiceSettlementBadgeLabel(options = {}) {
  return getInvoiceSettlementBadgeMeta(options).label;
}

function getInvoiceSettlementBadgeTone(options = {}) {
  return getInvoiceSettlementBadgeMeta(options).tone;
}

function formatInvoiceDateDisplay(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "--";
  return date.toLocaleDateString(getUiLocale(), {
    year: "numeric",
    month: "short",
    day: "2-digit",
  });
}

function buildReceiptViewModel(record) {
  const totals = calculateRecordTotals(record);
  const profile = buildMockAccountProfile(currentUser);
  const serviceType = translateServiceName(
    record.payload?.selection?.type || record.service_type || "--"
  );
  const labels = accountLabels.length ? accountLabels : record.payload?.labels || [];
  const headlineParts = formatHistoryHeadlineParts(record.created_at);
  const issuedAt = `${headlineParts.dateText}${headlineParts.timeText || ""}`.replace(/\u00A0/g, " ");
  const receiptTracking = buildReceiptTrackingId(record?.id);
  return {
    record,
    totals,
    profile,
    serviceType,
    labels,
    issuedAt,
    receiptNumber: receiptTracking.display,
    receiptSlug: receiptTracking.slug,
    quantity: totals.quantity,
    billingAddressLines: splitReceiptAddressLines(profile?.billingAddress),
  };
}

function buildInvoiceViewModel(record) {
  const totals = calculateRecordTotals(record);
  const profile = buildMockAccountProfile(currentUser);
  const serviceType = translateServiceName(
    record.payload?.selection?.type || record.service_type || "--"
  );
  const labels = accountLabels.length ? accountLabels : record.payload?.labels || [];
  const issuedDate = new Date(record?.created_at || Date.now());
  const issuedAt = formatInvoiceDateDisplay(issuedDate);
  const invoiceTracking = buildInvoiceTrackingId(record?.id);
  const paymentMode = getHistoryInvoicePaymentMode();
  const isMonthlyBilling = invoiceRequiresManualSettlementClient(paymentMode);
  const dueDate = isMonthlyBilling ? formatInvoiceDateDisplay(addDaysLocal(issuedDate, 30)) : "--";
  const paymentMethodLabel = getInvoicePaymentMethodLabel(paymentMode);
  const settlementLines = isMonthlyBilling
    ? [
        `IBAN ${INVOICE_ISSUER_PROFILE.iban}`,
        `Communication ${invoiceTracking.display}`,
      ]
    : [];
  const settlementTransferRows = isMonthlyBilling
    ? [
        { label: "IBAN", value: INVOICE_ISSUER_PROFILE.iban },
        { label: "Communication", value: invoiceTracking.display, mono: true },
      ]
    : [];
  const serviceRows = [
    {
      service: serviceType,
      quantity: totals.quantity,
      rate: formatMoney(totals.unitIncVat),
      total: formatMoney(totals.totalIncVat),
    },
  ];
  const settlementBadgeMeta = getInvoiceSettlementBadgeMeta({
    isMonthlyBilling,
    dueDate: addDaysLocal(issuedDate, 30),
    paymentMode,
  });

  return {
    record,
    totals,
    profile,
    serviceType,
    labels,
    issuedAt,
    issuedIso: issuedDate.toISOString(),
    dueAt: dueDate,
    invoiceNumber: invoiceTracking.display,
    invoiceSlug: invoiceTracking.slug,
    quantity: totals.quantity,
    billingAddressLines: splitInvoiceAddressLines(profile?.billingAddress),
    paymentMode,
    paymentMethodLabel,
    isMonthlyBilling,
    settlementBadge: settlementBadgeMeta.label,
    settlementBadgeTone: settlementBadgeMeta.tone,
    reverseVatNote:
      "VAT not charged. Any reverse-charge VAT is due by the recipient.",
    settlementTitle: "Payment & Summary",
    settlementLines,
    settlementTransferRows,
    bankIban: INVOICE_ISSUER_PROFILE.iban,
    issuer: INVOICE_ISSUER_PROFILE,
    serviceRows,
  };
}

function buildInvoiceServiceRowsFromBillingItems(items = [], invoice = {}) {
  const grouped = new Map();
  const safeItems = Array.isArray(items) && items.length
    ? items
    : [
        {
          service_type: "Shipping labels",
          quantity: Math.max(1, Number(invoice?.labels_count) || 1),
          amount_inc_vat: Number(invoice?.total_inc_vat || invoice?.subtotal_ex_vat || 0),
          sort_index: 0,
        },
      ];

  safeItems.forEach((item, index) => {
    const service = translateServiceName(item?.service_type || "Shipping labels");
    const key = String(service || "Shipping labels").trim().toLowerCase();
    const quantity = Math.max(1, Number(item?.quantity) || 1);
    const totalAmount = Number(item?.amount_inc_vat ?? item?.amount_ex_vat ?? 0) || 0;
    const sortIndex = Number.isFinite(Number(item?.sort_index)) ? Number(item.sort_index) : index;
    const existing = grouped.get(key) || {
      service,
      quantity: 0,
      totalAmount: 0,
      sortIndex,
    };
    existing.quantity += quantity;
    existing.totalAmount += totalAmount;
    existing.sortIndex = Math.min(existing.sortIndex, sortIndex);
    grouped.set(key, existing);
  });

  return Array.from(grouped.values())
    .sort((left, right) => left.sortIndex - right.sortIndex)
    .map((entry) => ({
      service: entry.service,
      quantity: entry.quantity,
      rate: formatMoney(entry.quantity ? entry.totalAmount / entry.quantity : entry.totalAmount),
      total: formatMoney(entry.totalAmount),
    }));
}

function buildBillingInvoiceViewModel(invoice = {}, options = {}) {
  const snapshot =
    invoice?.metadata?.invoice_profile && typeof invoice.metadata.invoice_profile === "object"
      ? invoice.metadata.invoice_profile
      : {};
  const paymentMode = String(invoice?.payment_mode || "invoice").trim().toLowerCase();
  const isMonthlyBilling = invoiceRequiresManualSettlementClient(paymentMode);
  const issuedDate =
    new Date(
      options?.issuedAt
      || invoice?.issued_at
      || invoice?.created_at
      || Date.now()
    );
  const resolvedIssuedDate = Number.isNaN(issuedDate.getTime()) ? new Date() : issuedDate;
  const dueSource =
    options?.dueAt
    || invoice?.due_at
    || (isMonthlyBilling ? addDaysLocal(resolvedIssuedDate, 30) : null);
  const dueDate = dueSource ? new Date(dueSource) : null;
  const resolvedDueDate = dueDate && !Number.isNaN(dueDate.getTime()) ? dueDate : null;
  const reminderStage = Math.max(0, Number(options?.reminderStage ?? invoice?.reminder_stage) || 0);
  const invoiceNumber = String(
    invoice?.reference
    || buildInvoiceTrackingId(invoice?.id).display
    || "--"
  ).trim();
  const serviceRows = buildInvoiceServiceRowsFromBillingItems(invoice?.items || [], invoice);
  const quantity =
    Number(invoice?.labels_count)
    || serviceRows.reduce((sum, row) => sum + (Number(row?.quantity) || 0), 0)
    || 0;
  const billingAddress = String(
    snapshot?.billingAddress
    || invoice?.billing_address
    || ""
  ).trim();
  const profile = {
    companyName: String(invoice?.company_name || snapshot?.companyName || "").trim() || "Client account",
    billingAddress,
    taxId: String(snapshot?.taxId || invoice?.tax_id || "").trim(),
    contactEmail: String(invoice?.contact_email || snapshot?.contactEmail || "").trim(),
  };
  const settlementBadgeMeta = getInvoiceSettlementBadgeMeta({
    isMonthlyBilling,
    dueDate: resolvedDueDate,
    paymentMode,
    reminderStage,
  });

  return {
    issuedAt: formatInvoiceDateDisplay(resolvedIssuedDate),
    issuedIso: resolvedIssuedDate.toISOString(),
    dueAt: resolvedDueDate ? formatInvoiceDateDisplay(resolvedDueDate) : "--",
    invoiceNumber,
    invoiceSlug: buildInvoicePdfFilenameFromReference(invoiceNumber).replace(/^invoice-/, "").replace(/\.pdf$/, ""),
    quantity,
    paymentMode,
    paymentMethodLabel: getInvoicePaymentMethodLabel(paymentMode),
    isMonthlyBilling,
    settlementBadge: settlementBadgeMeta.label,
    settlementBadgeTone: settlementBadgeMeta.tone,
    settlementTitle: "Payment & Summary",
    settlementLines: isMonthlyBilling
      ? [
          `IBAN ${INVOICE_ISSUER_PROFILE.iban}`,
          `Communication ${invoiceNumber}`,
        ]
      : [],
    settlementTransferRows: isMonthlyBilling
      ? [
          { label: "IBAN", value: INVOICE_ISSUER_PROFILE.iban },
          { label: "Communication", value: invoiceNumber, mono: true },
        ]
      : [],
    reverseVatNote:
      "VAT not charged. Any reverse-charge VAT is due by the recipient.",
    issuer: INVOICE_ISSUER_PROFILE,
    profile,
    billingAddressLines: splitInvoiceAddressLines(profile.billingAddress),
    totals: {
      totalIncVat: Number(invoice?.total_inc_vat || invoice?.subtotal_ex_vat || 0),
    },
    serviceRows,
  };
}

function buildReceiptRowsHtml(viewModel, rowIndices = null) {
  const {
    totals,
    labels,
  } = viewModel;
  const scopedLabels = Array.isArray(rowIndices)
    ? rowIndices.map((index) => labels[index]).filter(Boolean)
    : labels;

  return scopedLabels.length
    ? scopedLabels
        .map((label) => {
          const data = label?.data || {};
          const recipientName = data.recipientName || "--";
          const destination = formatReceiptDestination(data);
          const recipientMeta = destination && destination !== "--" ? ` · ${destination}` : "";
          return `
            <tr>
              <td><span class="receipt-cell-primary mono">${escapeHtml(label.labelId || "--")}</span></td>
              <td><span class="receipt-cell-primary mono">${escapeHtml(label.trackingId || "--")}</span></td>
              <td><span class="receipt-cell-primary">${escapeHtml(recipientName)}<span class="receipt-inline-meta">${escapeHtml(recipientMeta)}</span></span></td>
              <td>
                <span class="receipt-cell-primary mono">${escapeHtml(formatMoney(totals.unitIncVat))}</span>
              </td>
            </tr>`;
        })
        .join("")
    : `<tr><td colspan="4"><span class="receipt-cell-primary" style="color:var(--muted)">--</span></td></tr>`;
}

function buildInvoiceRowsHtml(viewModel, rowIndices = null) {
  const scopedRows = Array.isArray(rowIndices)
    ? rowIndices.map((index) => viewModel.serviceRows[index]).filter(Boolean)
    : viewModel.serviceRows;

  return scopedRows.length
    ? scopedRows
        .map((row) => `
            <tr>
              <td><span class="receipt-cell-primary">${escapeHtml(row.service || "--")}</span></td>
              <td><span class="receipt-cell-primary mono">${escapeHtml(String(row.quantity || "--"))}</span></td>
              <td><span class="receipt-cell-primary mono">${escapeHtml(row.rate || "--")}</span></td>
              <td><span class="receipt-cell-primary mono">${escapeHtml(row.total || "--")}</span></td>
            </tr>`)
        .join("")
    : `<tr><td colspan="4"><span class="receipt-cell-primary" style="color:var(--muted)">--</span></td></tr>`;
}

function buildReceiptHeaderHtml(viewModel) {
  const {
    profile,
    serviceType,
    issuedAt,
    quantity,
    billingAddressLines,
    totals,
  } = viewModel;
  const profileName = profile?.companyName || "--";
  const contactLine = [profile?.contactName, profile?.contactEmail].filter(Boolean).join(" • ") || "--";
  const taxLine = [profile?.taxId, profile?.customerId].filter(Boolean).join(" • ") || "--";
  const accountManager = profile?.accountManager || "--";

  return `
    <div class="receipt-topline">
      <div class="receipt-brand">
        <img src="shipide_logo.png" class="receipt-brand-logo" alt="Shipide" crossorigin="anonymous" />
      </div>
      <div class="receipt-top-meta">
        <span class="receipt-chip">${escapeHtml(tr("Receipt"))}</span>
        <div class="receipt-top-meta-lines mono">
          <span>${escapeHtml(issuedAt)}</span>
        </div>
      </div>
    </div>

    <div class="receipt-grid">
      <section class="receipt-panel">
        <div class="receipt-panel-title">${escapeHtml(tr("Issued To"))}</div>
        <div class="receipt-address">
          <div class="receipt-address-block">
            <span class="receipt-address-name">${escapeHtml(profileName)}</span>
            <div class="receipt-address-lines">
              <span>${escapeHtml(contactLine)}</span>
              ${billingAddressLines
                .map((line) => `<span>${escapeHtml(line)}</span>`)
                .join("")}
            </div>
          </div>
          <div class="receipt-address-meta mono">
            <span>${escapeHtml(taxLine)}</span>
            <span>${escapeHtml(profile?.contactPhone || "--")}</span>
            <span>${escapeHtml(tr("Account Manager"))}: ${escapeHtml(accountManager)}</span>
          </div>
        </div>
      </section>

      <section class="receipt-panel receipt-panel-summary">
        <div class="receipt-panel-title">${escapeHtml(tr("Summary"))}</div>
        <div class="receipt-kv-grid receipt-kv-grid-compact">
          <div class="receipt-kv">
            <span class="receipt-kv-key">${escapeHtml(tr("Service"))}</span>
            <span class="receipt-kv-value">${escapeHtml(serviceType)}</span>
          </div>
          <div class="receipt-kv">
            <span class="receipt-kv-key">${escapeHtml(tr("Quantity"))}</span>
            <span class="receipt-kv-value mono">${escapeHtml(String(quantity))}</span>
          </div>
          <div class="receipt-kv">
            <span class="receipt-kv-key">${escapeHtml(tr("Subtotal"))}</span>
            <span class="receipt-kv-value mono">${escapeHtml(formatMoney(totals.totalExVat))}</span>
          </div>
          <div class="receipt-kv receipt-kv-total">
            <span class="receipt-kv-key">${escapeHtml(tr("Total"))}</span>
            <span class="receipt-kv-value mono">${escapeHtml(formatMoney(totals.totalIncVat))}</span>
          </div>
          <div class="receipt-kv">
            <span class="receipt-kv-key">${escapeHtml(tr("Average Unit Rate"))}</span>
            <span class="receipt-kv-value mono">${escapeHtml(formatMoney(totals.unitExVat))}</span>
          </div>
        </div>
      </section>
    </div>
  `;
}

function buildInvoiceHeaderHtml(viewModel) {
  const {
    profile,
    issuedAt,
    dueAt,
    billingAddressLines,
    invoiceNumber,
    paymentMethodLabel,
    isMonthlyBilling,
    issuer,
  } = viewModel;
  const profileName = profile?.companyName || "--";
  const clientLines = [
    ...billingAddressLines,
    profile?.taxId ? `VAT ${profile.taxId}` : "",
  ].filter(Boolean);
  const issuerLines = [
    issuer?.descriptor || "",
    ...splitInvoiceAddressLines(issuer?.jurisdiction || ""),
    issuer?.email || "",
  ].filter(Boolean);

  return `
    <div class="receipt-topline">
      <div class="receipt-brand">
        <img src="shipide_logo.png" class="receipt-brand-logo" alt="Shipide" crossorigin="anonymous" />
      </div>
      <div class="receipt-top-meta">
        <div class="invoice-top-tag">
          <span class="invoice-top-tag-marker" aria-hidden="true"></span>
          <span class="invoice-top-tag-label mono">${escapeHtml(tr("Invoice"))}</span>
        </div>
        <div class="receipt-top-meta-lines mono">
          <span>${escapeHtml(issuedAt)}</span>
          <span>${escapeHtml(invoiceNumber)}</span>
        </div>
      </div>
    </div>

    <div class="invoice-party-grid">
      <section class="invoice-party-column">
        <div class="receipt-panel-title">${escapeHtml("Billed to")}</div>
        <div class="receipt-address">
          <div class="receipt-address-block">
            <span class="receipt-address-name">${escapeHtml(profileName)}</span>
            <div class="receipt-address-lines">
              ${clientLines.map((line) => `<span>${escapeHtml(line)}</span>`).join("")}
            </div>
          </div>
        </div>
      </section>

      <section class="invoice-party-column">
        <div class="receipt-panel-title">${escapeHtml("Billed by")}</div>
        <div class="receipt-address">
          <div class="receipt-address-block">
            <span class="receipt-address-name">${escapeHtml(issuer?.legalName || "--")}</span>
            <div class="receipt-address-lines">
              ${issuerLines.map((line) => `<span>${escapeHtml(line)}</span>`).join("")}
            </div>
          </div>
        </div>
      </section>
    </div>

    <div class="invoice-meta-strip">
      <div class="invoice-meta-item">
        <div class="receipt-panel-title">${escapeHtml("Invoice no.")}</div>
        <div class="invoice-meta-value mono">${escapeHtml(invoiceNumber)}</div>
      </div>
      <div class="invoice-meta-item">
        <div class="receipt-panel-title">${escapeHtml("Invoice date")}</div>
        <div class="invoice-meta-value mono">${escapeHtml(issuedAt)}</div>
      </div>
      <div class="invoice-meta-item">
        <div class="receipt-panel-title">${escapeHtml(isMonthlyBilling ? "Due date" : "Status")}</div>
        <div class="invoice-meta-value">${escapeHtml(isMonthlyBilling ? dueAt : "Paid automatically")}</div>
      </div>
      <div class="invoice-meta-item">
        <div class="receipt-panel-title">${escapeHtml("Payment")}</div>
        <div class="invoice-meta-value">${escapeHtml(paymentMethodLabel)}</div>
      </div>
    </div>
  `;
}

function buildReceiptTableCardHtml(viewModel, options = {}) {
  const {
    rowIndices = null,
    showSectionHead = true,
    showColumnHead = true,
  } = options;
  const receiptRows = buildReceiptRowsHtml(viewModel, rowIndices);
  const quantityValue = Array.isArray(rowIndices) ? rowIndices.length : viewModel.quantity;

  return `
    <section class="receipt-table-card${showSectionHead ? "" : " is-continuation"}">
      ${showSectionHead ? `
        <div class="receipt-table-head">
          <div class="receipt-table-copy">
            <span class="receipt-table-title">${escapeHtml(tr("Generated Labels"))}</span>
            <span class="receipt-table-sub">${escapeHtml(tr("One line per generated label."))}</span>
          </div>
          <span class="receipt-table-badge mono">${escapeHtml(`${quantityValue} ${tr("labels")}`)}</span>
        </div>
      ` : ""}
      <div class="receipt-table-wrap">
        <table class="receipt-table">
          ${showColumnHead ? `
            <thead>
              <tr>
                <th>${escapeHtml(tr("Label"))}</th>
                <th>${escapeHtml(tr("Tracking"))}</th>
                <th>${escapeHtml(tr("Recipient"))}</th>
                <th>${escapeHtml(tr("Amount"))}</th>
              </tr>
            </thead>
          ` : ""}
          <tbody>${receiptRows}</tbody>
        </table>
      </div>
    </section>
  `;
}

function buildInvoiceTableCardHtml(viewModel, options = {}) {
  const {
    rowIndices = null,
    showSectionHead = true,
    showColumnHead = true,
  } = options;
  const invoiceRows = buildInvoiceRowsHtml(viewModel, rowIndices);
  const quantityValue = viewModel.quantity;

  return `
    <section class="receipt-table-card invoice-table-card${showSectionHead ? "" : " is-continuation"}">
      ${showSectionHead ? `
        <div class="receipt-table-head">
          <div class="receipt-table-copy">
            <span class="receipt-table-title">${escapeHtml("Services")}</span>
            <span class="receipt-table-sub">${escapeHtml("Condensed by billed service.")}</span>
          </div>
          <span class="receipt-table-badge mono">${escapeHtml(`${quantityValue} ${tr("labels")}`)}</span>
        </div>
      ` : ""}
      <div class="receipt-table-wrap">
        <table class="receipt-table">
          ${showColumnHead ? `
            <thead>
              <tr>
                <th>${escapeHtml("Services")}</th>
                <th>${escapeHtml("Qty")}</th>
                <th>${escapeHtml("Rate")}</th>
                <th>${escapeHtml("Line total")}</th>
              </tr>
            </thead>
          ` : ""}
          <tbody>${invoiceRows}</tbody>
        </table>
      </div>
    </section>
  `;
}

function buildReceiptDisclaimerHtml() {
  return `
    <section class="receipt-disclaimer">
      <svg class="receipt-disclaimer-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" aria-hidden="true">
        <path d="M12 3 2.8 19a1 1 0 0 0 .87 1.5h16.66A1 1 0 0 0 21.2 19L12 3Z"></path>
        <path d="M12 9v5"></path>
        <circle cx="12" cy="17.2" r="1"></circle>
      </svg>
      <div class="receipt-disclaimer-copy">
        <span class="receipt-disclaimer-title">${escapeHtml(tr("Billing Notice"))}</span>
        <span class="receipt-disclaimer-text">${escapeHtml(
          tr(
            "This receipt is provided for operational reference only. It is not valid for tax or accounting purposes. Your invoice is the only valid billing document for bookkeeping."
          )
        )}</span>
      </div>
    </section>
  `;
}

function buildInvoiceSettlementHtml(viewModel) {
  const title = viewModel.settlementTitle || "Payment & Summary";
  const transferRows = Array.isArray(viewModel.settlementTransferRows)
    ? viewModel.settlementTransferRows
    : [];
  const summaryRows = [
    { label: "Subtotal", value: formatMoney(viewModel.totals.totalIncVat) },
    { label: "Total", value: formatMoney(viewModel.totals.totalIncVat), strong: true },
  ];
  return `
    <section class="invoice-settlement-stack">
      <div class="invoice-settlement-panel">
        <div class="invoice-settlement-grid${transferRows.length ? " has-transfer" : ""}">
          <div class="invoice-payment-block">
            <span class="receipt-panel-title">${escapeHtml(title)}</span>
            <span class="invoice-payment-badge is-${escapeHtml(
              viewModel.settlementBadgeTone || (viewModel.isMonthlyBilling ? "success" : "success")
            )}">${escapeHtml(
              viewModel.settlementBadge || (viewModel.isMonthlyBilling ? "Due in 30 days" : "Collected automatically")
            )}</span>
          </div>
          ${transferRows.length
            ? `<div class="invoice-transfer-block">
                <div class="invoice-transfer-lines">
                  ${transferRows
                    .map(
                      (row, index) => `
                        <div class="invoice-transfer-row${index === transferRows.length - 1 ? " is-total" : ""}">
                          <span>${escapeHtml(row.label || "--")}</span>
                          <span class="invoice-transfer-value${row.mono ? " mono" : ""}">${escapeHtml(
                            row.value || "--"
                          )}</span>
                        </div>`
                    )
                    .join("")}
                </div>
              </div>`
            : ""}
          <div class="invoice-summary-block">
            ${summaryRows
              .map(
                (row) => `
                  <div class="invoice-summary-row${row.strong ? " is-total" : ""}">
                    <span>${escapeHtml(row.label)}</span>
                    <span class="mono">${escapeHtml(row.value)}</span>
                  </div>`
              )
              .join("")}
          </div>
        </div>
      </div>
    </section>
  `;
}

function buildInvoiceFooterHtml(invoiceNumber) {
  return `
    <div class="receipt-footer">
      <span class="receipt-footer-left">Cryvelin LLC · billing@shipide.com</span>
      <span class="receipt-footer-right">${escapeHtml(invoiceNumber)}</span>
    </div>
  `;
}

function buildReceiptFooterHtml(receiptNumber) {
  return `
    <div class="receipt-footer">
      <span class="receipt-footer-left">Shipide · billing@shipide.com</span>
      <span class="receipt-footer-right">${escapeHtml(receiptNumber)}</span>
    </div>
  `;
}

function buildReceiptPageHtml(viewModel, options = {}) {
  const {
    rowIndices = null,
    showHeader = true,
    showTableCard = true,
    showSectionHead = true,
    showColumnHead = true,
    showDisclaimer = true,
  } = options;

  return `
    <div class="receipt-sheet">
      <div class="receipt-sheet-body">
        ${showHeader ? `<div class="receipt-header-region">${buildReceiptHeaderHtml(viewModel)}</div>` : ""}
        ${showTableCard ? buildReceiptTableCardHtml(viewModel, { rowIndices, showSectionHead, showColumnHead }) : ""}
        ${showDisclaimer ? buildReceiptDisclaimerHtml() : ""}
      </div>
      ${buildReceiptFooterHtml(viewModel.receiptNumber)}
    </div>
  `;
}

function buildReceiptDocumentHtml(record) {
  const viewModel = buildReceiptViewModel(record);
  return buildReceiptPageHtml(viewModel, {
    rowIndices: viewModel.labels.map((_, index) => index),
    showHeader: true,
    showTableCard: true,
    showSectionHead: true,
    showColumnHead: true,
    showDisclaimer: true,
  });
}

function buildInvoicePageHtml(viewModel, options = {}) {
  const {
    rowIndices = null,
    showHeader = true,
    showTableCard = true,
    showSectionHead = true,
    showColumnHead = true,
    showSettlement = true,
  } = options;

  return `
    <div class="receipt-sheet invoice-sheet">
      <div class="receipt-sheet-body">
        ${showHeader ? `<div class="receipt-header-region">${buildInvoiceHeaderHtml(viewModel)}</div>` : ""}
        ${showTableCard ? buildInvoiceTableCardHtml(viewModel, { rowIndices, showSectionHead, showColumnHead }) : ""}
        ${showSettlement ? `
          <section class="invoice-lower-stack">
            ${buildInvoiceSettlementHtml(viewModel)}
            ${viewModel.reverseVatNote
              ? `<div class="invoice-tax-note">${escapeHtml(viewModel.reverseVatNote)}</div>`
              : ""}
          </section>
        ` : ""}
      </div>
      ${buildInvoiceFooterHtml(viewModel.invoiceNumber)}
    </div>
  `;
}

function buildInvoiceDocumentHtml(record) {
  const viewModel = buildInvoiceViewModel(record);
  return buildInvoiceDocumentHtmlFromViewModel(viewModel);
}

function buildInvoiceDocumentHtmlFromViewModel(viewModel) {
  return buildInvoicePageHtml(viewModel, {
    rowIndices: viewModel.serviceRows.map((_, index) => index),
    showHeader: true,
    showTableCard: true,
    showSectionHead: true,
    showColumnHead: true,
    showSettlement: true,
  });
}

function renderInvoiceDetails(record) {
  if (!receiptDocument || !record) return;
  receiptDocument.innerHTML = buildInvoiceDocumentHtml(record);
}

function renderInvoiceViewModel(viewModel) {
  if (!receiptDocument || !viewModel) return;
  receiptDocument.innerHTML = buildInvoiceDocumentHtmlFromViewModel(viewModel);
}

function wrapReceiptPageDocumentHtml(pageHtml) {
  return `<div class="receipt-document">${pageHtml}</div>`;
}

function waitForAnimationFrame() {
  return new Promise((resolve) => window.requestAnimationFrame(() => resolve()));
}

async function buildPaginatedInvoicePageConfigs(viewModel) {
  if (!receiptDocument || !viewModel) {
    return [
      {
        firstPage: true,
        rows: Array.isArray(viewModel?.serviceRows)
          ? viewModel.serviceRows.map((_, index) => index)
          : [],
        hasSettlement: true,
      },
    ];
  }

  const allRows = Array.isArray(viewModel.serviceRows)
    ? viewModel.serviceRows.map((_, index) => index)
    : [];
  const previousMarkup = receiptDocument.innerHTML;
  const previousClassName = receiptDocument.className;

  try {
    receiptDocument.className = "receipt-document";
    receiptDocument.innerHTML = buildInvoicePageHtml(viewModel, {
      rowIndices: allRows,
      showHeader: true,
      showTableCard: true,
      showSectionHead: true,
      showColumnHead: true,
      showSettlement: true,
    });
    await waitForAnimationFrame();

    const measurementPage = receiptDocument;
    if (!measurementPage) {
      return [{ firstPage: true, rows: allRows, hasSettlement: true }];
    }

    const regions = measureReceiptRegions(measurementPage, 1);
    const bodyNode = measurementPage.querySelector(".receipt-sheet-body");
    const footerNode = measurementPage.querySelector(".receipt-footer");
    const bodyRect = bodyNode?.getBoundingClientRect();
    const footerRect = footerNode?.getBoundingClientRect();
    const footerGap = 8;
    const contentBudget =
      bodyRect && footerRect
        ? Math.max(0, footerRect.top - bodyRect.top - footerGap)
        : Math.max(0, bodyRect?.height || 0);

    if (!contentBudget || !regions.rows.length) {
      return [{ firstPage: true, rows: allRows, hasSettlement: true }];
    }

    const headerPx = regions.headerH;
    const tableGapPx = regions.tableGapY;
    const tableSectionHeadPx = regions.tableSectionHeadH;
    const theadPx = regions.theadH;
    const settlementPx = regions.disclaimer ? regions.disclaimer.h : 0;

    const pages = [];
    let rowIdx = 0;
    let budget = contentBudget - headerPx - tableGapPx - tableSectionHeadPx - theadPx;
    let pageRows = [];

    while (rowIdx < regions.rows.length) {
      const rowPx = regions.rows[rowIdx]?.h || 0;
      if (rowPx <= budget) {
        pageRows.push(rowIdx);
        budget -= rowPx;
        rowIdx += 1;
      } else {
        pages.push({ firstPage: pages.length === 0, rows: pageRows });
        pageRows = [];
        budget = contentBudget - theadPx;
      }
    }

    const isFirstPage = pages.length === 0;
    if (settlementPx + 4 <= budget) {
      pages.push({ firstPage: isFirstPage, rows: pageRows, hasSettlement: true });
    } else {
      pages.push({ firstPage: isFirstPage, rows: pageRows, hasSettlement: false });
      pages.push({ firstPage: false, rows: [], hasSettlement: true });
    }

    return pages;
  } finally {
    receiptDocument.className = previousClassName;
    receiptDocument.innerHTML = previousMarkup;
  }
}

async function renderInvoicePrintDocumentFromViewModel(viewModel) {
  if (!viewModel) {
    throw new Error("Invoice print view model is required.");
  }
  if (!receiptDocument) {
    throw new Error("Invoice print container is missing.");
  }
  const pages = await buildPaginatedInvoicePageConfigs(viewModel);
  if (pages.length <= 1) {
    const pageConfig = pages[0] || {
      firstPage: true,
      rows: Array.isArray(viewModel.serviceRows)
        ? viewModel.serviceRows.map((_, index) => index)
        : [],
      hasSettlement: true,
    };
    receiptDocument.className = "receipt-document";
    receiptDocument.innerHTML = buildInvoicePageHtml(viewModel, {
      rowIndices: pageConfig.rows,
      showHeader: pageConfig.firstPage,
      showTableCard: pageConfig.firstPage || pageConfig.rows.length > 0,
      showSectionHead: pageConfig.firstPage,
      showColumnHead: true,
      showSettlement: Boolean(pageConfig.hasSettlement),
    });
    return receiptDocument.outerHTML;
  }

  receiptDocument.className = "receipt-print-pages";
  receiptDocument.innerHTML = pages
    .map((pageConfig) =>
      wrapReceiptPageDocumentHtml(
        buildInvoicePageHtml(viewModel, {
          rowIndices: pageConfig.rows,
          showHeader: pageConfig.firstPage,
          showTableCard: pageConfig.firstPage || pageConfig.rows.length > 0,
          showSectionHead: pageConfig.firstPage,
          showColumnHead: true,
          showSettlement: Boolean(pageConfig.hasSettlement),
        })
      )
    )
    .join("");
  return receiptDocument.innerHTML;
}

if (typeof window !== "undefined") {
  window.__shipideInvoiceRenderer = {
    renderViewModel: renderInvoiceViewModel,
    renderPrintViewModel: renderInvoicePrintDocumentFromViewModel,
    buildDocumentHtmlFromViewModel: buildInvoiceDocumentHtmlFromViewModel,
    buildViewModelFromBillingInvoice: buildBillingInvoiceViewModel,
  };
}

function renderReceiptDetails(record) {
  if (!receiptDocument || !record) return;
  receiptDocument.innerHTML = buildReceiptDocumentHtml(record);
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

/* ---- receipt PDF helpers ---- */

function measureReceiptRegions(receiptDoc, scale) {
  const docRect = receiptDoc.getBoundingClientRect();
  const rel = (el) => {
    if (!el) return null;
    const r = el.getBoundingClientRect();
    return { y: (r.top - docRect.top) * scale, h: r.height * scale };
  };

  const headerRegion = rel(receiptDoc.querySelector(".receipt-header-region"));
  const grid = receiptDoc.querySelector(".receipt-grid");
  const gridPos = rel(grid);
  const headerH = headerRegion
    ? headerRegion.y + headerRegion.h
    : (gridPos ? gridPos.y + gridPos.h : 0);

  /* table layout */
  const tableCard = receiptDoc.querySelector(".receipt-table-card");
  const tableSectionHead = receiptDoc.querySelector(".receipt-table-head");
  const thead = receiptDoc.querySelector(".receipt-table thead");
  const tableCardPos = rel(tableCard);
  const tableSectionHeadPos = rel(tableSectionHead);
  const theadPos = rel(thead);
  const tableCardY = tableCardPos ? tableCardPos.y : 0;
  const tableGapY = Math.max(0, tableCardY - headerH);
  const tableSectionHeadH = tableSectionHeadPos ? tableSectionHeadPos.h : 0;
  const theadH = theadPos ? theadPos.h : 0;

  /* individual rows */
  const rows = [];
  const trs = receiptDoc.querySelectorAll(".receipt-table tbody tr");
  trs.forEach((tr) => {
    const p = rel(tr);
    if (p) rows.push(p);
  });

  const disclaimer = rel(
    receiptDoc.querySelector(".invoice-lower-stack")
      || receiptDoc.querySelector(".invoice-settlement-stack")
      || receiptDoc.querySelector(".invoice-settlement-panel")
      || receiptDoc.querySelector(".invoice-tax-note")
      || receiptDoc.querySelector(".receipt-disclaimer")
  );

  /* footer */
  const footer = rel(receiptDoc.querySelector(".receipt-footer"));

  return { headerH, tableGapY, tableSectionHeadH, theadH, rows, disclaimer, footer };
}

function getReceiptPdfFilename() {
  const tracking = buildReceiptTrackingId(accountActiveRecord?.id);
  return `receipt-${tracking.slug || tr("label-order")}.pdf`;
}

function getInvoicePdfFilename() {
  const tracking = buildInvoiceTrackingId(accountActiveRecord?.id);
  return `invoice-${tracking.slug || "shipide"}.pdf`;
}

function buildInvoicePdfFilenameFromReference(reference) {
  const safe = String(reference || "shipide")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9-]+/g, "-")
    .replace(/^-+|-+$/g, "");
  return `invoice-${safe || "shipide"}.pdf`;
}

function buildInvoiceVariantPdfFilename(reference, reminderStage = 0) {
  const base = buildInvoicePdfFilenameFromReference(reference).replace(/\.pdf$/i, "");
  const stage = Math.max(0, Number(reminderStage) || 0);
  return stage > 0 ? `${base}-stage-${stage}.pdf` : `${base}.pdf`;
}

function setPdfActionBusy(triggerButton, isBusy, options = {}) {
  if (!triggerButton) return () => {};
  const idleLabel = String(options?.idleLabel || "").trim();
  const busyLabel = String(options?.busyLabel || "").trim();
  const label = Array.from(triggerButton.children).find(
    (child) => child.tagName === "SPAN" && !child.classList.contains("btn-inline-loader")
  ) || null;
  const icon = triggerButton.querySelector("svg");
  const originalLabel = label ? label.textContent : "";
  let loader = triggerButton.querySelector(".btn-inline-loader");
  triggerButton.disabled = Boolean(isBusy);
  triggerButton.classList.toggle("is-loading", Boolean(isBusy));
  if (icon) {
    icon.hidden = Boolean(isBusy);
  }
  if (!loader) {
    loader = document.createElement("span");
    loader.className = "btn-inline-loader is-hidden";
    loader.setAttribute("aria-hidden", "true");
    for (let index = 0; index < 3; index += 1) {
      const bar = document.createElement("span");
      bar.className = "btn-inline-loader-bar";
      loader.appendChild(bar);
    }
    if (label) {
      triggerButton.insertBefore(loader, label);
    } else {
      triggerButton.appendChild(loader);
    }
  }
  loader.classList.toggle("is-hidden", !isBusy);
  if (label) {
    label.textContent = isBusy ? (busyLabel || originalLabel) : (idleLabel || originalLabel);
  }
  return () => {
    triggerButton.disabled = false;
    triggerButton.classList.remove("is-loading");
    if (icon) {
      icon.hidden = false;
    }
    if (loader) {
      loader.classList.add("is-hidden");
    }
    if (label) {
      label.textContent = idleLabel || originalLabel;
    }
  };
}

function openReceiptPdfTarget(blob, filename) {
  const url = URL.createObjectURL(blob);
  let opened = false;

  try {
    const receiptWindow = window.open(url, "_blank", "noopener,noreferrer");
    if (receiptWindow) {
      receiptWindow.opener = null;
      opened = true;
    }
  } catch (_error) {
    opened = false;
  }

  if (!opened) {
    triggerFileDownload(url, filename);
  }

  window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

async function buildReceiptPdfExport() {
  if (!accountActiveRecord) return;
  renderReceiptDetails(accountActiveRecord);

  try {
    if (
      receiptDocument &&
      window.html2canvas &&
      window.jspdf &&
      typeof window.jspdf.jsPDF === "function"
    ) {
      receiptDocument.classList.add("is-exporting");
      if (document.fonts?.ready) {
        await document.fonts.ready.catch(() => {});
      }
      await new Promise((resolve) => window.requestAnimationFrame(() => resolve()));

      const scaleFactor = Math.max(2, Math.min(window.devicePixelRatio || 1, 3));
      const h2cOpts = {
        backgroundColor: "#0b1018",
        scale: scaleFactor,
        useCORS: true,
        logging: false,
      };

      /* ---- unlock table scroll for export ---- */
      const tableWrap = receiptDocument.querySelector(".receipt-table-wrap");
      const origMaxH = tableWrap ? tableWrap.style.maxHeight : "";
      const origOverflow = tableWrap ? tableWrap.style.overflow : "";
      if (tableWrap) { tableWrap.style.maxHeight = "none"; tableWrap.style.overflow = "visible"; }

      /* ---- wait a frame so layout settles, then measure DOM regions ---- */
      await new Promise((r) => window.requestAnimationFrame(r));
      const regions = measureReceiptRegions(receiptDocument, scaleFactor);
      const fullCanvas = await window.html2canvas(receiptDocument, h2cOpts);

      /* restore table scroll */
      if (tableWrap) { tableWrap.style.maxHeight = origMaxH; tableWrap.style.overflow = origOverflow; }

      /* ---- build PDF ---- */
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({ orientation: "portrait", unit: "pt", format: "a4", compress: true });
      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();
      const margin = 18;
      const cW = pageWidth - margin * 2;
      const usableH = pageHeight - margin * 2;

      const paintBg = () => { pdf.setFillColor(0, 6, 15); pdf.rect(0, 0, pageWidth, pageHeight, "F"); };
      const toPt = (canvasPx) => (canvasPx * cW) / fullCanvas.width;

      const fullPdfH = toPt(fullCanvas.height);

      if (fullPdfH <= usableH) {
        /* ---- single page ---- */
        paintBg();
        pdf.addImage(fullCanvas.toDataURL("image/png"), "PNG", margin, margin, cW, fullPdfH, undefined, "FAST");
      } else {
        /* ---- multi-page: DOM-aware layout ---- */
        const headerPt = toPt(regions.headerH);
        const tableGapPt = toPt(regions.tableGapY);
        const tableSectionHeadPt = toPt(regions.tableSectionHeadH);
        const theadPt = toPt(regions.theadH);
        const footerPt = regions.footer ? toPt(regions.footer.h) : 0;
        const disclaimerPt = regions.disclaimer ? toPt(regions.disclaimer.h) : 0;
        const footerGap = 8;
        const contentBudget = usableH - footerPt - footerGap;

        /* ---- assign rows to pages ---- */
        const pages = [];
        let rowIdx = 0;
        const rowCount = regions.rows.length;

        /* page 1 budget: header + card gap + section header + table columns */
        let budget = contentBudget - headerPt - tableGapPt - tableSectionHeadPt - theadPt;
        let pageRows = [];

        while (rowIdx < rowCount) {
          const rowPt = toPt(regions.rows[rowIdx].h);
          if (rowPt <= budget) {
            pageRows.push(rowIdx);
            budget -= rowPt;
            rowIdx++;
          } else {
            /* current page is full */
            pages.push({ firstPage: pages.length === 0, rows: pageRows });
            pageRows = [];
            /* subsequent page budget: table column head only */
            budget = contentBudget - theadPt;
          }
        }

        /* check if disclaimer fits on current page */
        const isFirstPage = pages.length === 0;
        if (disclaimerPt + 4 <= budget) {
          pages.push({ firstPage: isFirstPage, rows: pageRows, hasDisclaimer: true });
        } else {
          pages.push({ firstPage: isFirstPage, rows: pageRows, hasDisclaimer: false });
          pages.push({ firstPage: false, rows: [], hasDisclaimer: true });
        }

        const viewModel = buildReceiptViewModel(accountActiveRecord);
        const renderExportPage = async (pageConfig) => {
          receiptDocument.innerHTML = buildReceiptPageHtml(viewModel, {
            rowIndices: pageConfig.rows,
            showHeader: pageConfig.firstPage,
            showTableCard: pageConfig.firstPage || pageConfig.rows.length > 0,
            showSectionHead: pageConfig.firstPage,
            showColumnHead: true,
            showDisclaimer: Boolean(pageConfig.hasDisclaimer),
          });
          await new Promise((resolve) => window.requestAnimationFrame(resolve));
          return window.html2canvas(receiptDocument, h2cOpts);
        };

        /* ---- render pages ---- */
        for (let p = 0; p < pages.length; p++) {
          const pageCanvas = await renderExportPage(pages[p]);
          if (p > 0) pdf.addPage();
          paintBg();
          const pagePdfH = (pageCanvas.height * cW) / pageCanvas.width;
          pdf.addImage(
            pageCanvas.toDataURL("image/png"),
            "PNG",
            margin,
            margin,
            cW,
            pagePdfH,
            undefined,
            "FAST"
          );
        }
      }
      return {
        blob: pdf.output("blob"),
        filename: getReceiptPdfFilename(),
      };
    }

    const viewModel = buildReceiptViewModel(accountActiveRecord);
    const fallbackLines = [
      tr("Receipt"),
      `${tr("Receipt No.")}: ${viewModel.receiptNumber}`,
      `${tr("Issued")}: ${viewModel.issuedAt}`,
      `${tr("Service")}: ${viewModel.serviceType}`,
      `${tr("Quantity")}: ${viewModel.quantity}`,
      `${tr("Subtotal")}: ${formatMoney(viewModel.totals.totalExVat)}`,
      `${tr("Total")}: ${formatMoney(viewModel.totals.totalIncVat)}`,
      "",
      tr(
        "This receipt is provided for operational reference only. It is not valid for tax or accounting purposes. Your invoice is the only valid billing document for bookkeeping."
      ),
    ];
    const blob = buildPdf(fallbackLines, {
      pageWidth: 612,
      pageHeight: 792,
      marginX: 40,
      marginTop: 48,
      marginBottom: 48,
      fontSize: 11,
      lineHeight: 14,
    });
    return {
      blob,
      filename: getReceiptPdfFilename(),
    };
  } catch (error) {
    throw error;
  } finally {
    if (receiptDocument) {
      receiptDocument.classList.remove("is-exporting");
    }
  }
}

async function buildInvoicePdfExport() {
  if (!accountActiveRecord) return;
  const viewModel = buildInvoiceViewModel(accountActiveRecord);
  return buildInvoicePdfExportFromViewModel(viewModel, getInvoicePdfFilename());
}

async function requestSelectableInvoicePdf(viewModel, filename) {
  if (
    !viewModel
    || (typeof window !== "undefined" && window.__SHIPIDE_INVOICE_PRINT_MODE__)
  ) {
    return null;
  }
  const response = await fetchApiBlobWithAuth("/api/billing/render-invoice-pdf", {
    method: "POST",
    timeoutMs: 120000,
    body: JSON.stringify({
      viewModel,
      filename,
    }),
  });
  if (!response?.blob) {
    return null;
  }
  return {
    blob: response.blob,
    filename,
  };
}

function base64ToBlob(base64, contentType = "application/pdf") {
  const normalized = String(base64 || "").trim();
  if (!normalized) return null;
  const binary = atob(normalized);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return new Blob([bytes], { type: contentType });
}

async function requestSelectableInvoicePdfBatch(entries = []) {
  if (
    !Array.isArray(entries)
    || !entries.length
    || (typeof window !== "undefined" && window.__SHIPIDE_INVOICE_PRINT_MODE__)
  ) {
    return [];
  }
  const payload = await fetchApiWithAuth("/api/billing/render-invoice-pdf-batch", {
    method: "POST",
    timeoutMs: 180000,
    body: JSON.stringify({
      variants: entries.map((entry) => ({
        reminderStage: entry?.reminderStage,
        filename: entry?.filename,
        viewModel: entry?.viewModel,
      })),
    }),
  });
  const rows = Array.isArray(payload?.variants) ? payload.variants : [];
  return rows.map((entry, index) => ({
    reminderStage: Math.max(0, Number(entry?.reminderStage ?? entries[index]?.reminderStage) || 0),
    filename: String(entry?.filename || entries[index]?.filename || "invoice-shipide.pdf").trim() || "invoice-shipide.pdf",
    blob: base64ToBlob(entry?.pdfBase64, "application/pdf"),
  }));
}

function collectInvoicePdfTextSnapshot(pageNode) {
  if (!(pageNode instanceof HTMLElement)) {
    return {
      width: 0,
      height: 0,
      textRuns: [],
    };
  }
  const pageRect = pageNode.getBoundingClientRect();
  const elements = Array.from(pageNode.querySelectorAll("*"));
  return {
    width: pageRect.width,
    height: pageRect.height,
    textRuns: elements
      .filter((el) => {
        if (!(el instanceof HTMLElement)) return false;
        if (el.children.length > 0) return false;
        return Boolean(String(el.textContent || "").trim());
      })
      .map((el) => {
        const rect = el.getBoundingClientRect();
        const style = window.getComputedStyle(el);
        return {
          text: String(el.textContent || "").replace(/\s+/g, " ").trim(),
          x: rect.left - pageRect.left,
          y: rect.top - pageRect.top,
          width: rect.width,
          height: rect.height,
          fontSize: Number.parseFloat(style.fontSize) || rect.height || 12,
          fontFamily: style.fontFamily || "",
        };
      })
      .filter((item) => item.text),
  };
}

function drawInvoicePdfTextLayer(pdf, snapshot, layout = {}) {
  if (!pdf || !snapshot?.textRuns?.length || !snapshot?.width) return;
  const margin = Number(layout?.margin) || 18;
  const pageHeight = Number(layout?.pageHeight) || 841.89;
  const contentWidth = Number(layout?.contentWidth) || 559.28;
  const scale = contentWidth / snapshot.width;

  snapshot.textRuns.forEach((run) => {
    const fontName = /mono/i.test(run.fontFamily) ? "courier" : "helvetica";
    const fontSize = Math.max(4, run.fontSize * scale);
    const baselineY =
      pageHeight
      - margin
      - ((run.y + run.height) * scale)
      + fontSize * 0.2;
    pdf.setFont(fontName, "normal");
    pdf.setFontSize(fontSize);
    pdf.setTextColor(0, 0, 0);
    pdf.text(run.text, margin + run.x * scale, baselineY, {
      baseline: "alphabetic",
      maxWidth: Math.max(fontSize, run.width * scale),
    });
  });
}

async function buildInvoicePdfExportFromViewModel(viewModel, filename = "invoice-shipide.pdf", options = {}) {
  if (!viewModel) return null;
  const allowRasterFallback = options?.allowRasterFallback !== false;
  const preferServerRender = options?.preferServerRender !== false;
  if (preferServerRender) {
    try {
      const renderedPdf = await requestSelectableInvoicePdf(viewModel, filename);
      if (renderedPdf?.blob) {
        return renderedPdf;
      }
      if (!allowRasterFallback) {
        throw new Error(tr("Could not generate the approved invoice PDF."));
      }
    } catch (error) {
      if (!allowRasterFallback) {
        throw error;
      }
      // Fall back to the local renderer when server-side print rendering is unavailable.
    }
  }

  const previousMarkup = receiptDocument ? receiptDocument.innerHTML : "";
  renderInvoiceViewModel(viewModel);

  try {
    if (
      receiptDocument &&
      window.html2canvas &&
      window.jspdf &&
      typeof window.jspdf.jsPDF === "function"
    ) {
      receiptDocument.classList.add("is-exporting");
      if (document.fonts?.ready) {
        await document.fonts.ready.catch(() => {});
      }
      await new Promise((resolve) => window.requestAnimationFrame(() => resolve()));

      const scaleFactor = Math.max(2, Math.min(window.devicePixelRatio || 1, 3));
      const h2cOpts = {
        backgroundColor: "#0b1018",
        scale: scaleFactor,
        useCORS: true,
        logging: false,
      };

      const tableWrap = receiptDocument.querySelector(".receipt-table-wrap");
      const origMaxH = tableWrap ? tableWrap.style.maxHeight : "";
      const origOverflow = tableWrap ? tableWrap.style.overflow : "";
      if (tableWrap) {
        tableWrap.style.maxHeight = "none";
        tableWrap.style.overflow = "visible";
      }

      await new Promise((r) => window.requestAnimationFrame(r));
      const regions = measureReceiptRegions(receiptDocument, scaleFactor);
      const fullCanvas = await window.html2canvas(receiptDocument, h2cOpts);

      if (tableWrap) {
        tableWrap.style.maxHeight = origMaxH;
        tableWrap.style.overflow = origOverflow;
      }

      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({ orientation: "portrait", unit: "pt", format: "a4", compress: true });
      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();
      const margin = 0;
      const cW = pageWidth - margin * 2;
      const usableH = pageHeight - margin * 2;

      const paintBg = () => {
        pdf.setFillColor(0, 0, 0);
        pdf.rect(0, 0, pageWidth, pageHeight, "F");
      };
      const toPt = (canvasPx) => (canvasPx * cW) / fullCanvas.width;
      const fullPdfH = toPt(fullCanvas.height);

      if (fullPdfH <= usableH) {
        paintBg();
        drawInvoicePdfTextLayer(pdf, collectInvoicePdfTextSnapshot(receiptDocument), {
          margin,
          pageHeight,
          contentWidth: cW,
        });
        pdf.addImage(fullCanvas.toDataURL("image/png"), "PNG", margin, margin, cW, fullPdfH, undefined, "FAST");
      } else {
        const headerPt = toPt(regions.headerH);
        const tableGapPt = toPt(regions.tableGapY);
        const tableSectionHeadPt = toPt(regions.tableSectionHeadH);
        const theadPt = toPt(regions.theadH);
        const footerPt = regions.footer ? toPt(regions.footer.h) : 0;
        const settlementPt = regions.disclaimer ? toPt(regions.disclaimer.h) : 0;
        const footerGap = 8;
        const contentBudget = usableH - footerPt - footerGap;

        const pages = [];
        let rowIdx = 0;
        const rowCount = regions.rows.length;
        let budget = contentBudget - headerPt - tableGapPt - tableSectionHeadPt - theadPt;
        let pageRows = [];

        while (rowIdx < rowCount) {
          const rowPt = toPt(regions.rows[rowIdx].h);
          if (rowPt <= budget) {
            pageRows.push(rowIdx);
            budget -= rowPt;
            rowIdx += 1;
          } else {
            pages.push({ firstPage: pages.length === 0, rows: pageRows });
            pageRows = [];
            budget = contentBudget - theadPt;
          }
        }

        const isFirstPage = pages.length === 0;
        if (settlementPt + 4 <= budget) {
          pages.push({ firstPage: isFirstPage, rows: pageRows, hasSettlement: true });
        } else {
          pages.push({ firstPage: isFirstPage, rows: pageRows, hasSettlement: false });
          pages.push({ firstPage: false, rows: [], hasSettlement: true });
        }

        const renderExportPage = async (pageConfig) => {
          receiptDocument.innerHTML = buildInvoicePageHtml(viewModel, {
            rowIndices: pageConfig.rows,
            showHeader: pageConfig.firstPage,
            showTableCard: pageConfig.firstPage || pageConfig.rows.length > 0,
            showSectionHead: pageConfig.firstPage,
            showColumnHead: true,
            showSettlement: Boolean(pageConfig.hasSettlement),
          });
          await new Promise((resolve) => window.requestAnimationFrame(resolve));
          const snapshot = collectInvoicePdfTextSnapshot(receiptDocument);
          const canvas = await window.html2canvas(receiptDocument, h2cOpts);
          return { canvas, snapshot };
        };

        for (let p = 0; p < pages.length; p += 1) {
          const pageRender = await renderExportPage(pages[p]);
          const pageCanvas = pageRender?.canvas;
          if (p > 0) pdf.addPage();
          paintBg();
          drawInvoicePdfTextLayer(pdf, pageRender?.snapshot, {
            margin,
            pageHeight,
            contentWidth: cW,
          });
          const pagePdfH = (pageCanvas.height * cW) / pageCanvas.width;
          pdf.addImage(
            pageCanvas.toDataURL("image/png"),
            "PNG",
            margin,
            margin,
            cW,
            pagePdfH,
            undefined,
            "FAST"
          );
        }
      }

      return {
        blob: pdf.output("blob"),
        filename,
      };
    }

    const fallbackLines = [
      tr("Invoice"),
      `${tr("Invoice No.")}: ${viewModel.invoiceNumber}`,
      `${tr("Issue date")}: ${viewModel.issuedAt}`,
      `${viewModel.isMonthlyBilling ? tr("Due date") : tr("Status")}: ${viewModel.isMonthlyBilling ? viewModel.dueAt : tr("Paid automatically")}`,
      `${tr("Payment method")}: ${viewModel.paymentMethodLabel}`,
      `${tr("Quantity")}: ${viewModel.quantity}`,
      `${tr("Total")}: ${formatMoney(viewModel.totals.totalIncVat)}`,
      "",
      ...viewModel.settlementLines,
    ];
    const blob = buildPdf(fallbackLines, {
      pageWidth: 612,
      pageHeight: 792,
      marginX: 40,
      marginTop: 48,
      marginBottom: 48,
      fontSize: 11,
      lineHeight: 14,
    });
    return {
      blob,
      filename,
    };
  } finally {
    if (receiptDocument) {
      receiptDocument.classList.remove("is-exporting");
      receiptDocument.innerHTML = previousMarkup;
    }
  }
}

function buildAdminTestBillingInvoiceRecord(toEmail) {
  const invoiceId = typeof crypto?.randomUUID === "function"
    ? crypto.randomUUID()
    : `inv-${Date.now()}`;
  const issuedDate = new Date();
  return {
    id: invoiceId,
    reference: buildInvoiceTrackingId(invoiceId).display,
    company_name: "Atelier Meridian",
    contact_name: "Claire Dupont",
    contact_email: String(toEmail || "").trim(),
    issued_at: issuedDate.toISOString(),
    due_at: addDaysLocal(issuedDate, 30)?.toISOString() || null,
    payment_mode: "invoice",
    total_inc_vat: 120,
    subtotal_ex_vat: 120,
    labels_count: 1,
    metadata: {
      invoice_profile: {
        companyName: "Atelier Meridian",
        contactName: "Claire Dupont",
        contactEmail: String(toEmail || "").trim(),
        contactPhone: "+32 2 555 01 29",
        billingAddress: "Avenue Louise 120, 1050 Brussels, Belgium",
        taxId: "BE0123456789",
        customerId: "SHIPIDE-TEST",
      },
    },
    items: [
      {
        service_type: "Economy",
        quantity: 1,
        amount_inc_vat: 120,
      },
    ],
  };
}

function buildAdminTestInvoiceViewModel(toEmail) {
  return buildBillingInvoiceViewModel(buildAdminTestBillingInvoiceRecord(toEmail), {
    reminderStage: 0,
  });
}

async function fetchAdminInvoiceDetail(invoiceId) {
  const safeInvoiceId = String(invoiceId || "").trim();
  if (!safeInvoiceId) {
    throw new Error(tr("Invoice id is required."));
  }
  const payload = await fetchApiWithAuth(
    `/api/admin/invoices/detail?invoiceId=${encodeURIComponent(safeInvoiceId)}`,
    { timeoutMs: 30000 }
  );
  return payload?.invoice && typeof payload.invoice === "object" ? payload.invoice : null;
}

function getInvoiceVariantStagesForPaymentMode(paymentMode = "") {
  return invoiceRequiresManualSettlementClient(paymentMode) ? [0, 1, 2, 3, 4, 5, 6, 7, 8] : [0];
}

async function buildInvoicePdfVariantsForInvoiceRecord(invoiceRecord) {
  const filenameReference = String(invoiceRecord?.reference || buildInvoiceTrackingId(invoiceRecord?.id).display || "shipide").trim();
  const stages = getInvoiceVariantStagesForPaymentMode(invoiceRecord?.payment_mode);
  const jobs = stages.map((reminderStage) => ({
    reminderStage,
    viewModel: buildBillingInvoiceViewModel(invoiceRecord, { reminderStage }),
    filename: buildInvoiceVariantPdfFilename(filenameReference, reminderStage),
  }));
  const variants = [];
  for (const job of jobs) {
    const exportData = await buildInvoicePdfExportFromViewModel(job.viewModel, job.filename, {
      allowRasterFallback: false,
      preferServerRender: true,
    });
    if (!exportData?.blob) {
      throw new Error(tr("Could not generate invoice PDF."));
    }
    variants.push({
      reminderStage: job.reminderStage,
      filename: exportData.filename,
      pdfBase64: await blobToBase64(exportData.blob),
    });
  }
  return variants;
}

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Could not encode PDF attachment."));
    reader.onload = () => {
      const result = String(reader.result || "");
      const base64 = result.includes(",") ? result.split(",").pop() : result;
      resolve(base64 || "");
    };
    reader.readAsDataURL(blob);
  });
}

async function downloadReceiptPdfFile() {
  if (!accountActiveRecord) return;
  const restoreButton = setPdfActionBusy(
    receiptDownloadPdf,
    true,
    {
      idleLabel: tr("Download Receipt PDF"),
      busyLabel: tr("Preparing receipt PDF..."),
    }
  );
  try {
    const exportData = await buildReceiptPdfExport();
    if (!exportData?.blob) {
      throw new Error(tr("Could not generate receipt PDF."));
    }
    const url = URL.createObjectURL(exportData.blob);
    triggerFileDownload(url, exportData.filename);
    window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
    showToast(tr("Receipt PDF ready."), { tone: "success" });
  } catch (error) {
    showToast(error?.message || tr("Could not generate receipt PDF."), { tone: "error" });
  } finally {
    restoreButton();
  }
}

async function openReceiptPdfFile() {
  if (!accountActiveRecord) return;
  const restoreButton = setPdfActionBusy(openReceiptModalButton, true, {
    idleLabel: tr("View Receipt"),
    busyLabel: tr("Preparing receipt PDF..."),
  });

  try {
    const exportData = await buildReceiptPdfExport();
    if (!exportData?.blob) {
      throw new Error(tr("Could not generate receipt PDF."));
    }
    openReceiptPdfTarget(exportData.blob, exportData.filename);
    showToast(tr("Receipt PDF ready."), { tone: "success" });
  } catch (error) {
    showToast(error?.message || tr("Could not generate receipt PDF."), { tone: "error" });
  } finally {
    restoreButton();
  }
}

async function downloadInvoicePdfFile() {
  if (!accountActiveRecord) return;
  const restoreButton = setPdfActionBusy(
    accountDownloadInvoice,
    true,
    {
      idleLabel: tr("Download Invoice"),
      busyLabel: tr("Preparing invoice PDF..."),
    }
  );
  try {
    const exportData = await buildInvoicePdfExport();
    if (!exportData?.blob) {
      throw new Error(tr("Could not generate invoice PDF."));
    }
    const url = URL.createObjectURL(exportData.blob);
    triggerFileDownload(url, exportData.filename);
    window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
    showToast(tr("Invoice PDF ready."), { tone: "success" });
  } catch (error) {
    showToast(error?.message || tr("Could not generate invoice PDF."), { tone: "error" });
  } finally {
    restoreButton();
  }
}

function selectAccountRecord(index) {
  const record = historyRecords[index];
  const labels = record?.payload?.labels;
  if (!record || !Array.isArray(labels) || !labels.length) return;

  accountActiveRecord = record;
  accountActiveHistoryIndex = index;
  accountActiveLabelIndex = 0;
  revokeAccountPdfUrls();

  const serviceType = translateServiceName(
    record.payload?.selection?.type || record.service_type || "Label"
  );
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
    accountPreviewMeta.textContent = "";
  }

  if (accountDownloadPdf) {
    accountDownloadPdf.disabled = false;
  }
  if (accountDownloadInvoice) {
    accountDownloadInvoice.disabled = false;
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
  queueHistoryPanelSync();
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
  return date.toLocaleDateString(getUiLocale(), { month: "numeric", day: "numeric" });
}

function formatReportTooltipDate(date) {
  if (!date) return "--";
  return date.toLocaleDateString(getUiLocale(), { month: "short", day: "numeric" });
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
  return record?.payload?.selection?.type || record?.service_type || tr("Unknown Service");
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
  reportsCustomRange.classList.toggle("is-open", Boolean(visible));
  reportsCustomRange.setAttribute("aria-hidden", visible ? "false" : "true");
}

function getSafeReportsRangeDays(value = reportsRangeDays) {
  return REPORT_RANGE_PRESETS.includes(Number(value)) ? Number(value) : 30;
}

function getAllTimeReportBounds(today) {
  const normalizedToday = toDayStart(today || new Date());
  const recordDays = historyRecords
    .map((record) => toDayStart(getRecordCreatedDay(record)))
    .filter(Boolean)
    .sort((left, right) => left - right);

  if (!recordDays.length) {
    const fallbackStart = new Date(normalizedToday);
    fallbackStart.setDate(fallbackStart.getDate() - 29);
    return {
      start: fallbackStart,
      end: normalizedToday,
    };
  }

  return {
    start: recordDays[0],
    end: recordDays[recordDays.length - 1] > normalizedToday ? normalizedToday : recordDays[recordDays.length - 1],
  };
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

  if (reportsRangeMode === "all") {
    const { start, end } = getAllTimeReportBounds(today);
    return {
      mode: "all",
      start,
      end,
      days: Math.max(1, Math.round((end - start) / 86400000) + 1),
      rangeToken: "all",
    };
  }

  const safeRange = getSafeReportsRangeDays();
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
    reportServicesLegend.innerHTML = `<div class="reports-empty">${tr(
      "No service data in this range."
    )}</div>`;
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
        <span class="reports-legend-name">${translateServiceName(service.name)}</span>
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
      <div class="reports-zone-label mono">${tr("Zone {index}", { index: zoneIndex })}</div>
      <div class="reports-zone-value mono">${count}</div>
    `;
    const bar = zone.querySelector(".reports-zone-bar");
    const scale = Math.max(0.06, count / maxZone);
    bar.style.transform = `scaleY(${scale.toFixed(3)})`;
    zone.title = `${tr("Zone {index}", { index: zoneIndex })}: ${tr("{count} labels", {
      count,
    })}`;
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
    const rowEl = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 3;
    td.className = "reports-empty";
    td.textContent = emptyMessage;
    rowEl.appendChild(td);
    bodyEl.appendChild(rowEl);
    return;
  }

  rows.forEach((row) => {
    const key = String(rowKey(row) || row.name || "");
    const rowEl = document.createElement("tr");
    rowEl.dataset.rowKey = key;
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
      nameCell.textContent = translateDomesticRegionName(row.name);
    }

    const countCell = document.createElement("td");
    countCell.className = "mono";
    countCell.textContent = String(row.count);
    const spendCell = document.createElement("td");
    spendCell.className = "mono";
    spendCell.textContent = formatMoney(row.spend);

    rowEl.appendChild(nameCell);
    rowEl.appendChild(countCell);
    rowEl.appendChild(spendCell);
    if (typeof onHover === "function") {
      rowEl.addEventListener("mouseenter", () => onHover(key));
      rowEl.addEventListener("mouseleave", () => onHover(""));
    }
    bodyEl.appendChild(rowEl);
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
    const name = shapeEl.getAttribute("data-country-name") || tr("Unknown");
    const row = byCountryKey.get(key);
    shapeEl.addEventListener("mouseenter", () => {
      setReportsCountryHighlight(key);
      const bbox = shapeEl.getBBox();
      const x = bbox.x + bbox.width / 2;
      const y = bbox.y + bbox.height / 2;
      const line = row
        ? `${name} • ${tr("{count} labels", { count: row.count })} • ${formatMoney(row.spend)}`
        : `${name} • ${tr("{count} labels", { count: 0 })} • ${formatMoney(0)}`;
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
      <text x="180" y="112" text-anchor="middle" class="reports-map-label">${tr(
        "Domestic map data unavailable"
      )}</text>
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
        `${province} (${translateDomesticRegionName(region)}) • ${tr("{count} labels", {
          count: row.count,
        })} • ${formatMoney(row.spend)}`
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
    const token = String(button.dataset.reportRange || "");
    button.classList.toggle("is-active", token === String(rangeToken));
  });
}

function seedReportsCustomDatesFromPreset() {
  const currentSettings =
    reportsRangeMode === "all"
      ? (() => {
          const today = toDayStart(new Date());
          const allTime = getAllTimeReportBounds(today);
          return { start: allTime.start, end: allTime.end };
        })()
      : getReportsRangeSettings();
  const start = toDayStart(currentSettings.start);
  const end = toDayStart(currentSettings.end);
  reportsCustomStart = formatDateInputValue(start);
  reportsCustomEnd = formatDateInputValue(end);
  if (reportsStartDate) reportsStartDate.value = reportsCustomStart;
  if (reportsEndDate) reportsEndDate.value = reportsCustomEnd;
}

function activateCustomReportsRangePicker() {
  if (!reportsCustomStart || !reportsCustomEnd) {
    seedReportsCustomDatesFromPreset();
  }
  reportsRangeMode = "custom";
  setReportRangeButtons("custom");
  setReportsCustomRangeVisible(true);
  if (reportsStartDate) reportsStartDate.classList.remove("is-invalid");
  if (reportsEndDate) reportsEndDate.classList.remove("is-invalid");
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
    reportSavingsMeta.textContent = tr("{count} labels analyzed", {
      count: model.totalShipments,
    });
  }

  if (reportPendingReturnsValue) {
    reportPendingReturnsValue.textContent = tr("{count} labels", {
      count: model.pendingReturnsCount,
    });
  }
  if (reportPendingReturnsMeta) {
    reportPendingReturnsMeta.textContent = tr("{amount} pending", {
      amount: formatMoney(model.pendingReturnsAmount),
    });
  }

  if (reportPendingRefundsValue) {
    reportPendingRefundsValue.textContent = tr("{count} labels", {
      count: model.pendingRefundsCount,
    });
  }
  if (reportPendingRefundsMeta) {
    reportPendingRefundsMeta.textContent = tr("{amount} pending", {
      amount: formatMoney(model.pendingRefundsAmount),
    });
  }

  if (reportAccountBalanceValue) {
    reportAccountBalanceValue.textContent = formatMoney(model.accountBalance);
  }
  if (reportTotalShippingCostsValue) {
    reportTotalShippingCostsValue.textContent = formatMoney(model.totalShippingCosts);
  }
  if (reportTotalShippingCostsMeta) {
    reportTotalShippingCostsMeta.textContent = tr("Selected range");
  }
  if (reportCarrierAdjustmentsValue) {
    reportCarrierAdjustmentsValue.textContent = formatSignedMoney(model.carrierAdjustments);
  }
  if (reportCarrierAdjustmentsMeta) {
    reportCarrierAdjustmentsMeta.textContent = tr("{count} events", {
      count: model.adjustmentEvents,
    });
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
    reportAvgDomesticMeta.textContent = tr("{count} labels", {
      count: model.avgDomesticCount,
    });
  }
  if (reportAvgInternationalValue) {
    reportAvgInternationalValue.textContent = formatMoney(model.avgInternational);
  }
  if (reportAvgInternationalMeta) {
    reportAvgInternationalMeta.textContent = tr("{count} labels", {
      count: model.avgInternationalCount,
    });
  }
  if (reportNewShipmentsValue) {
    reportNewShipmentsValue.textContent = String(model.newShipments);
  }
  if (reportNewShipmentsMeta) {
    reportNewShipmentsMeta.textContent = tr("{amount} spend", {
      amount: formatMoney(model.totalShippingCosts),
    });
  }
  if (reportDeliveryIssuesValue) {
    reportDeliveryIssuesValue.textContent = String(model.deliveryIssues);
  }
  if (reportDeliveryIssuesMeta) {
    reportDeliveryIssuesMeta.textContent = tr("{rate}% rate", {
      rate: model.deliveryIssueRate.toFixed(2),
    });
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
    tooltipLabel: tr("Payments"),
  });
  renderReportLineChart({
    svgEl: reportAverageCostChart,
    axisEl: reportAverageCostAxis,
    tooltipEl: reportAverageCostTooltip,
    points: model.buckets,
    valueKey: "averageCost",
    valueFormatter: formatMoney,
    tooltipLabel: tr("Average Cost"),
  });
  renderReportLineChart({
    svgEl: reportShipmentsChart,
    axisEl: reportShipmentsAxis,
    tooltipEl: reportShipmentsTooltip,
    points: model.buckets,
    valueKey: "shipments",
    valueFormatter: (value) => tr("{count} labels", { count: Math.round(value) }),
    tooltipLabel: tr("Shipments"),
  });
  renderReportsServices(model);
  renderReportsZones(model);
  renderReportsLocationTable(
    reportTopRegionsBody,
    model.topRegions,
    tr("No domestic region data yet."),
    {
      rowKey: (row) => row.name,
      onHover: (key) => setReportsRegionHighlight(key),
    }
  );
  renderReportsLocationTable(
    reportTopCountriesBody,
    model.topCountries,
    tr("No international country data yet."),
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

let lastAuthToast = {
  message: "",
  tone: "",
  at: 0,
};

function showAuthToast(message, tone = "info") {
  const text = String(message || "").trim();
  const normalizedTone = String(tone || "info").trim() || "info";
  if (!text) return;
  const now = Date.now();
  if (
    lastAuthToast.message === text
    && lastAuthToast.tone === normalizedTone
    && now - lastAuthToast.at < 900
  ) {
    return;
  }
  lastAuthToast = {
    message: text,
    tone: normalizedTone,
    at: now,
  };
  showToast(text, { tone: normalizedTone });
}

function setAuthMessage(message, options = {}) {
  const { isError = true, tone = isError ? "error" : "info", toast = true } = options;
  const text = String(message || "").trim();
  if (authError) {
    authError.hidden = true;
    authError.textContent = "";
    authError.classList.remove("is-visible", "is-info");
  }
  if (toast && text) {
    showAuthToast(text, tone);
  }
}

function setAuthInviteStatus(message = "", options = {}) {
  const { tone = "info", toast = true } = options;
  const text = String(message || "").trim();
  if (authInviteStatus) {
    authInviteStatus.hidden = true;
    authInviteStatus.textContent = "";
    authInviteStatus.classList.remove("is-visible", "is-error", "is-success");
  }
  if (toast && text) {
    showAuthToast(text, tone);
  }
}

function setAuthAgreementStatus(message = "", options = {}) {
  const { tone = "info", toast = true } = options;
  const text = String(message || "").trim();
  if (authAgreementStatus) {
    authAgreementStatus.hidden = true;
    authAgreementStatus.textContent = "";
    authAgreementStatus.classList.remove("is-success");
  }
  if (toast && text) {
    showAuthToast(text, tone);
  }
}

function updateAuthRegisterSubmitState() {
  const primaryLabel = authMode === "register"
    ? (authRegisterStep === 1 ? tr("Continue") : tr("Register Account"))
    : tr("Sign In");

  if (authSignIn) {
    const label = authSignIn.querySelector("span");
    if (label) {
      label.textContent = primaryLabel;
    } else {
      authSignIn.textContent = primaryLabel;
    }
  }

  if (authMode !== "register") {
    if (authSignIn) {
      authSignIn.disabled = authIsBusy;
    }
    if (authRegisterBack) {
      authRegisterBack.classList.add("is-hidden");
      authRegisterBack.disabled = true;
    }
    return;
  }

  const stepOneValid = getRegistrationStepOneValidationError() === "";
  const agreementReady = Boolean(
    authAgreementContract && authAgreementHasReachedEnd && authAgreementAccepted
  );
  const canSubmit = authRegisterStep === 1 ? stepOneValid : agreementReady;

  if (authSignIn) {
    authSignIn.disabled = authIsBusy || !authInviteIsValid || !canSubmit;
  }
  if (authRegisterBack) {
    const showBack = authRegisterStep === 2;
    authRegisterBack.classList.toggle("is-hidden", !showBack);
    authRegisterBack.disabled = authIsBusy || !showBack;
  }
}

function getRegistrationStepOneValidationError() {
  const email = String(authEmail?.value || "").trim().toLowerCase();
  const password = String(authPassword?.value || "");
  const confirmPassword = String(authPasswordConfirm?.value || "");
  const profile = collectRegistrationProfile();

  if (!email || !password) {
    return tr("Email and password are required.");
  }
  if (!confirmPassword) {
    return tr("Confirm your password to continue.");
  }
  const passwordValidationError = getRegistrationPasswordValidationError({
    password,
    confirmPassword,
  });
  if (passwordValidationError) {
    return passwordValidationError;
  }
  if (!validateRegistrationProfile(profile)) {
    return tr("All registration fields are required.");
  }
  return "";
}

function getRegistrationPasswordValidationError(options = {}) {
  const password = String(options?.password ?? authPassword?.value ?? "");
  const confirmPassword = String(options?.confirmPassword ?? authPasswordConfirm?.value ?? "");
  if (password.length < 10) {
    return tr("Password must be at least {min} characters.", { min: 10 });
  }
  if (confirmPassword && password !== confirmPassword) {
    return tr("Passwords do not match.");
  }
  return "";
}

function maybeShowRegistrationPasswordToast(input) {
  if (authMode !== "register") return;
  if (!input) return;
  const password = String(authPassword?.value || "");
  const confirmPassword = String(authPasswordConfirm?.value || "");

  if (input === authPassword) {
    if (!password) return;
    const validationError = getRegistrationPasswordValidationError({
      password,
      confirmPassword: "",
    });
    if (validationError) {
      setAuthMessage(validationError);
    }
    return;
  }

  if (input === authPasswordConfirm) {
    if (!confirmPassword) return;
    const validationError = getRegistrationPasswordValidationError({
      password,
      confirmPassword,
    });
    if (validationError) {
      setAuthMessage(validationError);
    }
  }
}

function updateAuthRegisterProgressState(step = authRegisterStep) {
  if (!authRegisterProgressItems.length) return;
  authRegisterProgressItems.forEach((item) => {
    const itemStep = Number(item.dataset.step || "1");
    item.classList.toggle("is-active", itemStep === step);
    item.classList.toggle("is-complete", itemStep < step);
  });
}

function applyAuthRegisterStepView(step = authRegisterStep) {
  const nextStep = Number(step) === 2 ? 2 : 1;
  if (authForm) {
    authForm.dataset.registerStep = String(nextStep);
  }
  if (authCard) {
    authCard.classList.toggle("is-register-step-two", authMode === "register" && nextStep === 2);
  }
  if (authStepOneShell) {
    authStepOneShell.classList.toggle("is-hidden", authMode === "register" && nextStep === 2);
    authStepOneShell.classList.remove("auth-step-transition-exit", "auth-step-transition-enter-start");
  }
  if (authAgreementGroup) {
    authAgreementGroup.classList.toggle("is-hidden", authMode !== "register" || nextStep !== 2);
    authAgreementGroup.classList.remove("auth-step-transition-exit", "auth-step-transition-enter-start");
  }
  updateAuthRegisterProgressState(nextStep);
}

function measureAuthRegisterSectionHeight(section) {
  if (!section || !authForm) return 0;

  const wasHidden = section.classList.contains("is-hidden");
  const previousStyles = {
    position: section.style.position,
    inset: section.style.inset,
    width: section.style.width,
    visibility: section.style.visibility,
    pointerEvents: section.style.pointerEvents,
    zIndex: section.style.zIndex,
  };

  section.classList.remove("is-hidden", "auth-step-transition-exit", "auth-step-transition-enter-start");
  section.style.position = "absolute";
  section.style.inset = "0 auto auto 0";
  section.style.width = `${Math.floor(authForm.getBoundingClientRect().width)}px`;
  section.style.visibility = "hidden";
  section.style.pointerEvents = "none";
  section.style.zIndex = "-1";

  const height = section.getBoundingClientRect().height;

  section.style.position = previousStyles.position;
  section.style.inset = previousStyles.inset;
  section.style.width = previousStyles.width;
  section.style.visibility = previousStyles.visibility;
  section.style.pointerEvents = previousStyles.pointerEvents;
  section.style.zIndex = previousStyles.zIndex;

  if (wasHidden) {
    section.classList.add("is-hidden");
  }

  return height;
}

function resetAuthRegisterStepTransitionState() {
  authRegisterStepTransitionToken += 1;
  if (authCard) {
    authCard.classList.remove("is-register-step-transitioning");
    authCard.style.width = "";
    authCard.style.height = "";
  }
  if (authCardFrame) {
    authCardFrame.classList.remove(
      "is-step-transition-exit",
      "is-step-transition-hidden",
      "is-step-transition-enter-start"
    );
  }
  if (authForm) {
    authForm.classList.remove("is-register-step-transitioning");
  }
  if (authStepOneShell) {
    authStepOneShell.classList.remove("auth-step-transition-exit", "auth-step-transition-enter-start");
  }
  if (authAgreementGroup) {
    authAgreementGroup.classList.remove("auth-step-transition-exit", "auth-step-transition-enter-start");
  }
}

function setAuthRegisterStep(step = 1, options = {}) {
  const { resetMessage = false, animate = false } = options;
  const nextStep = Number(step) === 2 ? 2 : 1;
  const previousStep = authRegisterStep === 2 ? 2 : 1;
  const shouldAnimate =
    animate &&
    authMode === "register" &&
    previousStep !== nextStep &&
    authCard &&
    authCardFrame &&
    authForm &&
    authStepOneShell &&
    authAgreementGroup &&
    !prefersReducedMotion();

  authRegisterStep = nextStep;
  updateAuthRegisterProgressState(nextStep);

  if (resetMessage) {
    setAuthMessage("");
  }

  if (!shouldAnimate) {
    resetAuthRegisterStepTransitionState();
    applyAuthRegisterStepView(nextStep);
    if (authMode === "register" && nextStep === 2) {
      window.requestAnimationFrame(() => {
        updateAuthAgreementProgress();
      });
    }
    updateAuthRegisterSubmitState();
    return;
  }

  const transitionToken = ++authRegisterStepTransitionToken;
  const currentSection = previousStep === 2 ? authAgreementGroup : authStepOneShell;
  const incomingSection = nextStep === 2 ? authAgreementGroup : authStepOneShell;

  if (!currentSection || !incomingSection) {
    applyAuthRegisterStepView(nextStep);
    updateAuthRegisterSubmitState();
    return;
  }

  const cardRect = authCard.getBoundingClientRect();
  const currentSectionHeight = Math.max(1, currentSection.getBoundingClientRect().height);
  const incomingSectionHeight = Math.max(1, measureAuthRegisterSectionHeight(incomingSection));
  const chromeHeight = Math.max(0, cardRect.height - currentSectionHeight);
  const targetHeight = Math.max(1, chromeHeight + incomingSectionHeight);
  const containerWidth = authCard.parentElement?.getBoundingClientRect().width || cardRect.width;
  const targetWidth = authMode === "register" && nextStep === 2
    ? Math.min(468, containerWidth)
    : containerWidth;

  authCard.style.width = `${Math.round(cardRect.width)}px`;
  authCard.style.height = `${Math.round(cardRect.height)}px`;
  authCard.classList.add("is-register-step-transitioning");
  authCardFrame.classList.remove("is-step-transition-hidden", "is-step-transition-enter-start");
  authCardFrame.classList.add("is-step-transition-exit");
  authForm.classList.add("is-register-step-transitioning");

  window.setTimeout(() => {
    if (transitionToken !== authRegisterStepTransitionToken) return;

    authCardFrame.classList.remove("is-step-transition-exit");
    authCardFrame.classList.add("is-step-transition-hidden");
    if (authForm) {
      authForm.dataset.registerStep = String(nextStep);
    }
    applyAuthRegisterStepView(nextStep);
    if (authCard) {
      authCard.style.width = `${Math.round(targetWidth)}px`;
      authCard.style.height = `${Math.round(targetHeight)}px`;
    }
  }, AUTH_REGISTER_STEP_OUT_MS);

  window.setTimeout(() => {
    if (transitionToken !== authRegisterStepTransitionToken) return;

    authCardFrame.classList.remove("is-step-transition-hidden");
    authCardFrame.classList.add("is-step-transition-enter-start");
    window.requestAnimationFrame(() => {
      if (transitionToken !== authRegisterStepTransitionToken) return;
      authCardFrame.classList.remove("is-step-transition-enter-start");
      if (authMode === "register" && nextStep === 2) {
        updateAuthAgreementProgress();
      }
    });
  }, AUTH_REGISTER_STEP_OUT_MS + AUTH_REGISTER_STEP_RESIZE_MS);

  window.setTimeout(() => {
    if (transitionToken !== authRegisterStepTransitionToken) return;
    resetAuthRegisterStepTransitionState();
    applyAuthRegisterStepView(nextStep);
    updateAuthRegisterSubmitState();
  }, AUTH_REGISTER_STEP_OUT_MS + AUTH_REGISTER_STEP_RESIZE_MS + AUTH_REGISTER_STEP_IN_MS);
}

function setAuthAgreementPlaceholder(message, { isError = false, isLoading = false } = {}) {
  if (!authAgreementPages) return;
  const placeholder = document.createElement("div");
  placeholder.className = `auth-agreement-placeholder${isError ? " is-error" : ""}${isLoading ? " is-loading" : ""}`;
  if (isLoading) {
    const loader = document.createElement("span");
    loader.className = "auth-agreement-loader";
    loader.setAttribute("aria-hidden", "true");
    for (let index = 0; index < 3; index += 1) {
      const bar = document.createElement("span");
      bar.className = "auth-agreement-loader-bar";
      loader.appendChild(bar);
    }

    const copy = document.createElement("div");
    copy.className = "auth-agreement-placeholder-copy";

    const title = document.createElement("strong");
    title.className = "auth-agreement-placeholder-title";
    title.textContent = String(message || "").trim();

    const sub = document.createElement("span");
    sub.className = "auth-agreement-placeholder-sub";
    sub.textContent = tr("Loading your contract preview...");

    copy.append(title, sub);
    placeholder.append(loader, copy);
  } else {
    placeholder.textContent = String(message || "").trim();
  }
  authAgreementPages.replaceChildren(placeholder);
}

function renderAuthAgreementTextFallback(text) {
  if (!authAgreementPages) return;
  const content = String(text || "").trim();
  if (!content) {
    authAgreementPages.replaceChildren();
    return;
  }
  const fallback = document.createElement("pre");
  fallback.className = "auth-agreement-text-fallback mono";
  fallback.textContent = content;
  authAgreementPages.replaceChildren(fallback);
}

function resolveAuthAgreementPdfUrl(contract) {
  const value = String(contract?.pdfUrl || "").trim();
  if (!value) return "";
  try {
    return new URL(value, window.location.origin).toString();
  } catch (_) {
    return value;
  }
}

function updateAuthAgreementDownload(contract) {
  if (!authAgreementDownload) return;
  const pdfUrl = resolveAuthAgreementPdfUrl(contract);
  if (!pdfUrl) {
    authAgreementDownload.href = "#";
    authAgreementDownload.setAttribute("aria-disabled", "true");
    authAgreementDownload.classList.add("is-disabled");
    authAgreementDownload.removeAttribute("download");
    return;
  }

  let filename =
    String(contract?.downloadName || "").trim()
    || `shipide-agreement-${String(contract?.version || "document").trim() || "document"}.pdf`;
  try {
    const pdfLocation = new URL(pdfUrl, window.location.origin);
    const lastSegment = String(pdfLocation.pathname.split("/").pop() || "").trim();
    if (lastSegment && !String(contract?.downloadName || "").trim()) {
      filename = lastSegment;
    }
  } catch (_) {
    // Keep the generated filename if the URL cannot be parsed cleanly.
  }

  authAgreementDownload.href = pdfUrl;
  authAgreementDownload.setAttribute("download", filename);
  authAgreementDownload.removeAttribute("aria-disabled");
  authAgreementDownload.classList.remove("is-disabled");
}

function hideAuthAgreementMagnifier() {
  if (authAgreementLens) {
    authAgreementLens.classList.remove("is-active");
  }
  const scrollWrap = authAgreementScroll?.closest(".auth-agreement-scroll-wrap");
  if (scrollWrap) {
    scrollWrap.classList.remove("is-magnifier-active");
  }
  authAgreementMagnifierPage = null;
  if (authAgreementLensInner) {
    authAgreementLensInner.replaceChildren();
  }
}

function createAuthAgreementMagnifierGraphic(pageElement) {
  const sourceGraphic = pageElement?.querySelector("svg, canvas");
  if (!sourceGraphic || !authAgreementLensInner) return false;
  const sourceRect = sourceGraphic.getBoundingClientRect();
  const width = Math.max(1, Math.floor(sourceRect.width));
  const height = Math.max(1, Math.floor(sourceRect.height));
  let clone = null;

  if (sourceGraphic instanceof SVGElement) {
    clone = sourceGraphic.cloneNode(true);
  } else if (sourceGraphic instanceof HTMLCanvasElement) {
    clone = document.createElement("canvas");
    clone.width = sourceGraphic.width;
    clone.height = sourceGraphic.height;
    const cloneContext = clone.getContext("2d", { alpha: false });
    if (cloneContext) {
      cloneContext.drawImage(sourceGraphic, 0, 0);
    }
  }

  if (!clone) return false;
  clone.classList.add("auth-agreement-lens-graphic");
  clone.style.width = `${width}px`;
  clone.style.height = `${height}px`;
  authAgreementLensInner.replaceChildren(clone);
  authAgreementMagnifierPage = pageElement;
  return true;
}

function updateAuthAgreementMagnifier(event) {
  if (
    !authAgreementLens ||
    !authAgreementLensInner ||
    !authAgreementScroll ||
    !(event instanceof PointerEvent)
  ) {
    return;
  }
  if (event.pointerType && event.pointerType !== "mouse") {
    hideAuthAgreementMagnifier();
    return;
  }

  const pageElement = event.target instanceof Element
    ? event.target.closest(".auth-agreement-page")
    : null;
  const scrollWrap = authAgreementScroll.closest(".auth-agreement-scroll-wrap");
  if (!pageElement || !scrollWrap) {
    hideAuthAgreementMagnifier();
    return;
  }

  if (
    authAgreementMagnifierPage !== pageElement ||
    !authAgreementLensInner.firstElementChild
  ) {
    const created = createAuthAgreementMagnifierGraphic(pageElement);
    if (!created) {
      hideAuthAgreementMagnifier();
      return;
    }
  }

  const lensGraphic = authAgreementLensInner.firstElementChild;
  if (!(lensGraphic instanceof HTMLElement || lensGraphic instanceof SVGElement)) {
    hideAuthAgreementMagnifier();
    return;
  }

  const sourceGraphic = pageElement.querySelector("svg, canvas");
  const sourceRect = sourceGraphic?.getBoundingClientRect()
    || pageElement.getBoundingClientRect();
  const wrapRect = scrollWrap.getBoundingClientRect();
  const offsetX = Math.max(0, Math.min(sourceRect.width, event.clientX - sourceRect.left));
  const offsetY = Math.max(0, Math.min(sourceRect.height, event.clientY - sourceRect.top));
  const lensHalfWidth = AUTH_AGREEMENT_LENS_WIDTH / 2;
  const lensHalfHeight = AUTH_AGREEMENT_LENS_HEIGHT / 2;
  const wrapPadding = 8;
  const maxLeft = Math.max(wrapPadding, wrapRect.width - AUTH_AGREEMENT_LENS_WIDTH - wrapPadding);
  const maxTop = Math.max(wrapPadding, wrapRect.height - AUTH_AGREEMENT_LENS_HEIGHT - wrapPadding);

  let lensLeft = event.clientX - wrapRect.left - lensHalfWidth;
  let lensTop = event.clientY - wrapRect.top - lensHalfHeight;
  lensLeft = Math.max(wrapPadding, Math.min(maxLeft, lensLeft));
  lensTop = Math.max(wrapPadding, Math.min(maxTop, lensTop));

  authAgreementLens.style.transform = `translate(${Math.round(lensLeft)}px, ${Math.round(lensTop)}px)`;
  lensGraphic.style.transform = `translate(${Math.round(lensHalfWidth - offsetX * AUTH_AGREEMENT_LENS_ZOOM)}px, ${Math.round(lensHalfHeight - offsetY * AUTH_AGREEMENT_LENS_ZOOM)}px) scale(${AUTH_AGREEMENT_LENS_ZOOM})`;
  authAgreementLens.classList.add("is-active");
  scrollWrap.classList.add("is-magnifier-active");
}

async function renderAuthAgreementPdf(contract, renderToken) {
  if (!authAgreementPages || !authAgreementScroll) return;
  const pdfUrl = resolveAuthAgreementPdfUrl(contract);
  const fallbackText = String(contract?.bodyText || "").trim();
  const pdfjsLib = window.pdfjsLib;

  if (
    !pdfUrl ||
    !pdfjsLib ||
    typeof pdfjsLib.getDocument !== "function"
  ) {
    renderAuthAgreementTextFallback(fallbackText);
    return;
  }

  if (pdfjsLib.GlobalWorkerOptions) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = AUTH_AGREEMENT_PDF_WORKER_URL;
  }

  setAuthAgreementPlaceholder(tr("Loading agreement..."), { isLoading: true });

  try {
    const loadingTask = pdfjsLib.getDocument({
      url: pdfUrl,
      withCredentials: false,
    });
    const pdfDocument = await loadingTask.promise;
    if (renderToken !== authAgreementPdfRenderToken) return;

    const fragment = document.createDocumentFragment();
    const availableWidth = Math.max(220, Math.min(284, authAgreementScroll.clientWidth - 18));
    const isGeneratedPreview = Boolean(
      contract?.recordId ||
      contract?.previewPdfHash ||
      /\/api\/auth\/register-contract-preview-file\b/i.test(pdfUrl)
    );
    const canRenderSvg = !isGeneratedPreview && typeof pdfjsLib.SVGGraphics === "function";
    const outputScale = isGeneratedPreview
      ? Math.min(6, Math.max(3, (window.devicePixelRatio || 1) * 3))
      : Math.min(4, Math.max(2, (window.devicePixelRatio || 1) * 2));

    const renderPageToCanvas = async (page, viewport, pageElement) => {
      const canvas = document.createElement("canvas");
      canvas.width = Math.floor(viewport.width * outputScale);
      canvas.height = Math.floor(viewport.height * outputScale);
      canvas.style.width = `${Math.floor(viewport.width)}px`;
      canvas.style.height = `${Math.floor(viewport.height)}px`;
      pageElement.appendChild(canvas);

      const canvasContext = canvas.getContext("2d", { alpha: false });
      if (!canvasContext) return false;
      canvasContext.setTransform(outputScale, 0, 0, outputScale, 0, 0);
      await page.render({ canvasContext, viewport }).promise;
      return true;
    };

    for (let pageNumber = 1; pageNumber <= pdfDocument.numPages; pageNumber += 1) {
      if (renderToken !== authAgreementPdfRenderToken) return;
      const page = await pdfDocument.getPage(pageNumber);
      const baseViewport = page.getViewport({ scale: 1 });
      const scale = Math.max(0.65, Math.min(1.4, availableWidth / baseViewport.width));
      const viewport = page.getViewport({ scale });

      const pageElement = document.createElement("div");
      pageElement.className = "auth-agreement-page";
      pageElement.style.width = `${Math.floor(viewport.width)}px`;

      if (canRenderSvg) {
        try {
          const operatorList = await page.getOperatorList();
          if (renderToken !== authAgreementPdfRenderToken) return;
          const svgGraphics = new pdfjsLib.SVGGraphics(page.commonObjs, page.objs);
          const svg = await svgGraphics.getSVG(operatorList, viewport);
          if (renderToken !== authAgreementPdfRenderToken) return;
          svg.style.display = "block";
          svg.style.width = "100%";
          svg.style.height = "auto";
          pageElement.appendChild(svg);
          fragment.appendChild(pageElement);
          continue;
        } catch (error) {
          console.warn("Falling back to canvas agreement rendering for page", pageNumber, error);
        }
      }

      const rendered = await renderPageToCanvas(page, viewport, pageElement);
      if (!rendered) continue;
      fragment.appendChild(pageElement);
    }

    if (renderToken !== authAgreementPdfRenderToken) return;
    authAgreementPages.replaceChildren(fragment);
  } catch (error) {
    console.error("Failed to render agreement PDF", error);
    if (renderToken !== authAgreementPdfRenderToken) return;
    if (fallbackText) {
      renderAuthAgreementTextFallback(fallbackText);
    } else {
      setAuthAgreementPlaceholder(tr("Could not load registration agreement."), { isError: true });
    }
  }
}

function resetAuthAgreementState({ clearContract = false } = {}) {
  authAgreementPdfRenderToken += 1;
  authAgreementRendering = false;
  authAgreementHasReachedEnd = false;
  authAgreementAccepted = false;
  authAgreementScrolledAt = "";
  authAgreementAgreedAt = "";
  authAgreementMaxProgress = 0;
  hideAuthAgreementMagnifier();
  if (clearContract) {
    authAgreementContract = null;
  }
  if (authAgreementAccept) {
    authAgreementAccept.checked = false;
    authAgreementAccept.disabled = true;
  }
  if (authAgreementScroll) {
    authAgreementScroll.scrollTop = 0;
    const scrollWrap = authAgreementScroll.closest(".auth-agreement-scroll-wrap");
    if (scrollWrap) {
      scrollWrap.classList.remove("is-complete");
    }
  }
  if (authAgreementPages) {
    if (clearContract) {
      setAuthAgreementPlaceholder(tr("Loading agreement..."), { isLoading: true });
    } else {
      authAgreementPages.replaceChildren();
    }
  }
  if (authAgreementVersion) {
    authAgreementVersion.textContent = clearContract
      ? tr("Agreement version {version}", { version: "--" })
      : authAgreementVersion.textContent;
  }
  updateAuthAgreementDownload(clearContract ? null : authAgreementContract);
  setAuthAgreementStatus("");
}

function renderAuthAgreementContract(contract) {
  const normalized =
    contract && typeof contract === "object" ? contract : null;
  authAgreementContract = normalized;
  resetAuthAgreementState({ clearContract: false });
  if (!normalized) {
    setAuthAgreementPlaceholder(tr("Could not load registration agreement."), { isError: true });
    if (authAgreementVersion) {
      authAgreementVersion.textContent = tr("Agreement version {version}", { version: "--" });
    }
    updateAuthAgreementDownload(null);
    setAuthAgreementStatus(tr("Could not load registration agreement."), { tone: "error" });
    updateAuthRegisterSubmitState();
    return;
  }
  if (authAgreementVersion) {
    authAgreementVersion.textContent = tr("Agreement version {version}", {
      version: String(normalized.version || "--"),
    });
  }
  updateAuthAgreementDownload(normalized);
  setAuthAgreementPlaceholder(tr("Loading agreement..."), { isLoading: true });
  setAuthAgreementStatus("");
  const renderToken = ++authAgreementPdfRenderToken;
  authAgreementRendering = true;
  void renderAuthAgreementPdf(normalized, renderToken).finally(() => {
    if (renderToken !== authAgreementPdfRenderToken) return;
    authAgreementRendering = false;
    if (authAgreementScroll) {
      authAgreementScroll.scrollTop = 0;
    }
    window.requestAnimationFrame(() => {
      updateAuthAgreementProgress();
    });
  });
  updateAuthRegisterSubmitState();
}

function updateAuthAgreementProgress() {
  if (!authAgreementScroll) return;
  if (authAgreementRendering) return;
  const maxScrollTop = Math.max(1, authAgreementScroll.scrollHeight - authAgreementScroll.clientHeight);
  const progress = Math.min(1, Math.max(0, authAgreementScroll.scrollTop / maxScrollTop));
  authAgreementMaxProgress = Math.max(authAgreementMaxProgress, progress);
  const reachedEnd = progress >= 0.985 || authAgreementScroll.scrollHeight <= authAgreementScroll.clientHeight + 2;
  const scrollWrap = authAgreementScroll.closest(".auth-agreement-scroll-wrap");
  if (reachedEnd && !authAgreementHasReachedEnd) {
    authAgreementHasReachedEnd = true;
    authAgreementScrolledAt = new Date().toISOString();
    if (authAgreementAccept) {
      authAgreementAccept.disabled = false;
    }
    setAuthAgreementStatus("");
  }
  if (scrollWrap) {
    scrollWrap.classList.toggle("is-complete", authAgreementHasReachedEnd);
  }
  updateAuthRegisterSubmitState();
}

function buildRegistrationAgreementPayload() {
  if (!authAgreementContract) return null;
  const timezone = Intl.DateTimeFormat().resolvedOptions().timeZone || "";
  return {
    contractId: authAgreementContract.id || null,
    contractVersion: String(authAgreementContract.version || "").trim(),
    contractHash: String(authAgreementContract.hash || "").trim(),
    recordId: String(authAgreementContract.recordId || "").trim() || null,
    previewToken: String(authAgreementContract.previewToken || "").trim() || null,
    scrolledToEnd: authAgreementHasReachedEnd,
    scrolledToEndAt: authAgreementScrolledAt || null,
    agreed: authAgreementAccepted,
    agreedAt: authAgreementAgreedAt || null,
    maxScrollProgress: Number(authAgreementMaxProgress.toFixed(4)),
    clientTimezone: timezone,
    clientLocale: navigator.language || "",
  };
}

function setAuthMode(mode, options = {}) {
  const { inviteToken = "" } = options;
  authMode = mode === "register" ? "register" : "login";
  const isRegister = authMode === "register";
  authInviteIsValid = false;
  authInviteToken = String(inviteToken || "").trim();

  if (authGate) {
    authGate.classList.toggle("is-register-mode", isRegister);
  }
  if (authForm) {
    authForm.classList.toggle("is-register-mode", isRegister);
  }
  if (authRegisterOnlyFields?.length) {
    authRegisterOnlyFields.forEach((field) => {
      field.classList.toggle("is-hidden", !isRegister);
    });
  }
  if (authForgotPassword) {
    authForgotPassword.classList.toggle("is-hidden", isRegister);
  }
  if (authTitle) {
    authTitle.textContent = isRegister
      ? tr("Complete your account registration")
      : tr("Sign in to continue");
  }
  if (authSubtitle) {
    authSubtitle.textContent = isRegister
      ? tr("Use your invite link to create your secure account and billing profile.")
      : tr("Use your email and password.");
  }
  if (authRegisterProgress) {
    authRegisterProgress.classList.toggle("is-hidden", !isRegister);
  }
  if (authPassword) {
    authPassword.setAttribute("autocomplete", isRegister ? "new-password" : "current-password");
  }
  if (authPasswordConfirm) {
    authPasswordConfirm.required = isRegister;
    if (!isRegister) {
      authPasswordConfirm.value = "";
    }
  }
  if (authEmail) {
    authEmail.readOnly = false;
  }
  const requiredFields = [
    authCompanyName,
    authContactName,
    authContactEmail,
    authContactPhone,
    authBillingStreet,
    authBillingCity,
    authBillingPostalCode,
    authBillingCountry,
    authTaxId,
  ];
  requiredFields.forEach((input) => {
    if (!input) return;
    input.required = isRegister;
  });
  if (authCustomerId) {
    authCustomerId.required = false;
  }
  if (!isRegister) {
    setAuthInviteStatus("");
    authInviteData = null;
    resetAuthAgreementState({ clearContract: true });
    setAuthRegisterStep(1);
  } else {
    resetAuthAgreementState({ clearContract: true });
    setAuthRegisterStep(1);
  }
  updateAuthRegisterSubmitState();
}

function parseBillingAddressParts(value = "") {
  const normalized = String(value || "").trim();
  const result = {
    street: "",
    city: "",
    postalCode: "",
    country: "",
  };
  if (!normalized) return result;

  const segments = normalized
    .split(",")
    .map((part) => String(part || "").trim())
    .filter(Boolean);

  if (!segments.length) return result;
  result.street = segments[0] || "";
  result.country = segments.length > 2 ? segments[segments.length - 1] : "";
  const citySegment = segments.length > 1
    ? segments.length > 2
      ? segments[segments.length - 2]
      : segments[1]
    : "";

  const cityMatch = citySegment.match(/^([0-9][A-Za-z0-9-]*)\s+(.+)$/);
  if (cityMatch) {
    result.postalCode = String(cityMatch[1] || "").trim();
    result.city = String(cityMatch[2] || "").trim();
  } else if (!result.country && citySegment) {
    result.city = citySegment;
  } else if (citySegment) {
    result.city = citySegment;
  }
  return result;
}

function buildBillingAddressValue() {
  const street = String(authBillingStreet?.value || "").trim();
  const city = String(authBillingCity?.value || "").trim();
  const postalCode = String(authBillingPostalCode?.value || "").trim();
  const country = String(authBillingCountry?.value || "").trim();
  const locality = [postalCode, city].filter(Boolean).join(" ");
  const address = [street, locality, country].filter(Boolean).join(", ");
  if (authBillingAddress) {
    authBillingAddress.value = address;
  }
  return address;
}

function updateAuthBillingCountryFlag(value = authBillingCountry?.value || "") {
  if (!authBillingCountryFlag) return;
  authBillingCountryFlag.innerHTML = getCountryIcon(String(value || "").trim());
}

function applyBillingAddressDefaults(invite = {}) {
  const parsed = parseBillingAddressParts(
    invite.billingAddress || invite.billing_address || ""
  );
  const street = String(invite.billingStreet || invite.billing_street || parsed.street || "").trim();
  const city = String(invite.billingCity || invite.billing_city || parsed.city || "").trim();
  const postalCode = String(
    invite.billingPostalCode || invite.billing_postal_code || parsed.postalCode || ""
  ).trim();
  const country = String(invite.billingCountry || invite.billing_country || parsed.country || "").trim();
  if (authBillingStreet) authBillingStreet.value = street;
  if (authBillingCity) authBillingCity.value = city;
  if (authBillingPostalCode) authBillingPostalCode.value = postalCode;
  if (authBillingCountry) authBillingCountry.value = country;
  updateAuthBillingCountryFlag(country);
  buildBillingAddressValue();
}

function applyInviteDefaults(invite = {}) {
  if (authEmail) {
    authEmail.readOnly = false;
  }
  if (authCompanyName && invite.companyName) authCompanyName.value = invite.companyName;
  if (authContactName && invite.contactName) authContactName.value = invite.contactName;
  if (authContactEmail) {
    authContactEmail.value = invite.contactEmail || authContactEmail.value || "";
  }
  if (authContactPhone && invite.contactPhone) authContactPhone.value = invite.contactPhone;
  applyBillingAddressDefaults(invite);
  if (authTaxId && invite.taxId) authTaxId.value = invite.taxId;
  if (authCustomerId && invite.customerId) authCustomerId.value = invite.customerId;
}

function applySignupPreviewDefaults() {
  applyInviteDefaults(AUTH_SIGNUP_PREVIEW_DATA.invite);
  if (authEmail) {
    authEmail.value = AUTH_SIGNUP_PREVIEW_DATA.credentials.email;
  }
  if (authPassword) {
    authPassword.value = AUTH_SIGNUP_PREVIEW_DATA.credentials.password;
  }
  if (authPasswordConfirm) {
    authPasswordConfirm.value = AUTH_SIGNUP_PREVIEW_DATA.credentials.password;
  }
}

function collectRegistrationProfile() {
  const billingAddress = buildBillingAddressValue();
  return {
    companyName: String(authCompanyName?.value || "").trim(),
    contactName: String(authContactName?.value || "").trim(),
    contactEmail: String(authContactEmail?.value || "").trim().toLowerCase(),
    contactPhone: String(authContactPhone?.value || "").trim(),
    billingAddress,
    taxId: String(authTaxId?.value || "").trim(),
    customerId: String(authCustomerId?.value || "").trim(),
  };
}

async function fetchPublicApi(path, options = {}) {
  const { timeoutMs = 15000, ...requestOptions } = options;
  const headers = new Headers(requestOptions.headers || {});
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
        tr("Request failed ({status})", { status: response.status });
      throw new Error(message);
    }
    return payload;
  } catch (error) {
    if (error?.name === "AbortError") {
      throw new Error(tr("Request failed ({status})", { status: "timeout" }));
    }
    throw error;
  } finally {
    window.clearTimeout(timeoutId);
  }
}

async function prepareRegistrationAgreementPreview(options = {}) {
  const { animate = true } = options;
  const inviteToken = String(authInviteToken || getInviteTokenFromLocation(window.location)).trim();
  const email = String(authEmail?.value || "").trim().toLowerCase();
  const profile = collectRegistrationProfile();
  const validationError = getRegistrationStepOneValidationError();

  if (validationError) {
    setAuthMessage(validationError);
    return false;
  }
  if (!inviteToken) {
    setAuthMessage(tr("Registration link required."));
    return false;
  }
  if (!authInviteIsValid) {
    setAuthMessage(tr("This registration link is invalid or expired."));
    return false;
  }

  setAuthMessage("");
  setAuthBusy(true, tr("Preparing agreement..."), "");
  try {
    const payload = await fetchPublicApi("/api/auth/register-contract-preview", {
      method: "POST",
      body: JSON.stringify({
        token: inviteToken,
        email,
        profile,
      }),
    });
    const contract = payload?.contract || null;
    const isReady = Boolean(contract?.pdfUrl && contract?.version && contract?.hash);
    if (!isReady) {
      throw new Error(tr("Could not prepare agreement preview."));
    }
    renderAuthAgreementContract(contract);
    setAuthRegisterStep(2, { animate });
    return true;
  } catch (error) {
    setAuthMessage(error?.message || tr("Could not prepare agreement preview."));
    return false;
  } finally {
    setAuthBusy(false, tr("Continue"), "");
    updateAuthRegisterSubmitState();
  }
}

async function loadRegistrationInvite(token, options = {}) {
  const { preview = false, previewStep = 1 } = options;
  const requestToken = ++authInviteValidationToken;
  const inviteToken = preview ? AUTH_SIGNUP_PREVIEW_TOKEN : String(token || "").trim();
  authInviteIsValid = false;
  authInviteToken = inviteToken;
  authInviteData = null;
  resetAuthAgreementState({ clearContract: true });

  if (preview) {
    authInviteData = { ...AUTH_SIGNUP_PREVIEW_DATA.invite };
    authInviteToken = AUTH_SIGNUP_PREVIEW_DATA.inviteToken;
    applySignupPreviewDefaults();
    authInviteIsValid = true;
    setAuthInviteStatus(
      tr("Local signup preview mode. Sample data is loaded and no live account will be created."),
      { tone: "success" }
    );
    setAuthMessage("");
    if (requestToken === authInviteValidationToken) {
      if (Number(previewStep) === 2) {
        const prepared = await prepareRegistrationAgreementPreview({ animate: false });
        if (!prepared) {
          renderAuthAgreementContract({ ...AUTH_SIGNUP_PREVIEW_DATA.contract });
          setAuthRegisterStep(1, { resetMessage: true });
        }
      } else {
        renderAuthAgreementContract({ ...AUTH_SIGNUP_PREVIEW_DATA.contract });
        setAuthRegisterStep(1, { resetMessage: true });
      }
      setAuthBusy(false, tr("Register Account"), "");
      updateAuthRegisterSubmitState();
    }
    return;
  }

  if (!inviteToken) {
    setAuthInviteStatus("", { tone: "info" });
    setAuthMessage(tr("Registration link required."));
    updateAuthRegisterSubmitState();
    return;
  }

  setAuthMessage("");
  setAuthInviteStatus(tr("Validating invitation..."));
  setAuthBusy(true, tr("Validating invitation..."), "");

  try {
    const payload = await fetchPublicApi(
      `/api/auth/register-invite?token=${encodeURIComponent(inviteToken)}`
    );
    if (requestToken !== authInviteValidationToken) return;
    authInviteData = payload?.invite || null;
    renderAuthAgreementContract(payload?.contract || null);
    const hasContract = Boolean(payload?.contract?.version && payload?.contract?.hash);
    if (!hasContract) {
      throw new Error(tr("Could not load registration agreement."));
    }
    applyInviteDefaults(authInviteData || {});
    authInviteIsValid = true;
    setAuthInviteStatus(tr("Invitation verified. Complete your details to activate access."), {
      tone: "success",
    });
    setAuthMessage("");
  } catch (error) {
    if (requestToken !== authInviteValidationToken) return;
    authInviteData = null;
    authInviteIsValid = false;
    renderAuthAgreementContract(null);
    setAuthInviteStatus(tr("This registration link is invalid or expired."), {
      tone: "error",
    });
    setAuthMessage("");
  } finally {
    if (requestToken === authInviteValidationToken) {
      setAuthBusy(false, tr("Register Account"), "");
      updateAuthRegisterSubmitState();
    }
  }
}

function setAuthBusy(isBusy, signInLabel = tr("Sign In"), signUpLabel = tr("Create Account")) {
  authIsBusy = Boolean(isBusy);
  if (authEmail) authEmail.disabled = isBusy;
  if (authPassword) authPassword.disabled = isBusy;
  [
    authPasswordConfirm,
    authCompanyName,
    authContactName,
    authContactEmail,
    authContactPhone,
    authBillingStreet,
    authBillingCity,
    authBillingPostalCode,
    authBillingCountry,
    authBillingAddress,
    authTaxId,
    authCustomerId,
  ].forEach((input) => {
    if (input) input.disabled = isBusy;
  });
  if (authAgreementAccept) {
    authAgreementAccept.disabled = isBusy || !authAgreementHasReachedEnd;
  }
  if (authForgotPassword) {
    authForgotPassword.disabled = isBusy;
  }
  if (authSignIn) {
    const label = authSignIn.querySelector("span");
    if (label) {
      label.textContent = signInLabel;
    } else {
      authSignIn.textContent = signInLabel;
    }
  }
  if (authSignUp) {
    authSignUp.disabled = isBusy;
    const label = authSignUp.querySelector("span");
    if (label) {
      label.textContent = signUpLabel;
    } else {
      authSignUp.textContent = signUpLabel;
    }
  }
  updateAuthRegisterSubmitState();
}

function setAuthView(session, options = {}) {
  const { animate = true } = options;
  currentUser = session?.user || null;
  cachedAuthAccessToken = String(session?.access_token || "");
  setClientInviteBusy(false);
  void setLanguage(resolvePreferredLanguage(currentUser), { persist: false });
  const isAuthed = Boolean(currentUser);
  const route = parseRouteFromLocation();
  renderAccountProfile(currentUser);
  if (isAuthed) {
    adminAccessAllowed = false;
    adminAccessStatusRequestId += 1;
    adminAccessStatusPromise = null;
    startAuthKeepAlive();
    loadWarehouseSettings({ quiet: true });
    loadShopifyConnectionStatus({ quiet: true });
    refreshAdminOnlyTestTools();
    if (openLeadsPageButton) {
      openLeadsPageButton.disabled = true;
      openLeadsPageButton.classList.add("is-hidden");
      openLeadsPageButton.setAttribute("aria-busy", "true");
    }
    if (openAdminPageButton) {
      openAdminPageButton.disabled = true;
      openAdminPageButton.classList.add("is-hidden");
      openAdminPageButton.setAttribute("aria-busy", "true");
    }
    void loadAdminAccessStatus({ quiet: true });
    loadBillingOverview({ quiet: true });
    if (!normalizeLanguageCode(currentUser?.user_metadata?.preferred_language)) {
      const localPref = getStoredLanguagePreference();
      if (SUPPORTED_LANGUAGES.has(localPref)) {
        void saveLanguagePreference(localPref);
      }
    }
    if (adminBillingTestEmailInput && !adminBillingTestEmailInput.value.trim()) {
      adminBillingTestEmailInput.value = currentUser?.email || "";
    }
  } else {
    adminAccessStatusRequestId += 1;
    adminAccessStatusPromise = null;
    stopAuthKeepAlive();
    resetWarehouseState();
    shopifyConnection = null;
    updateShopifyProviderStatus();
    adminAccessAllowed = false;
    adminDashboardLoaded = false;
    adminDashboardState = null;
    adminClients = [];
    adminBillingInvoices = [];
    clientInviteHistory = [];
    adminClientBillingBusyIds = new Set();
    adminMockModeEnabled = false;
    adminMockSnapshot = null;
    billingOverview = null;
    refreshAdminOnlyTestTools();
    renderClientInviteHistory([]);
    setClientInviteStatus("");
    setClientInviteResult("");
    renderAdminSummary({});
    renderAdminMockDataButton();
    renderAdminClientsList();
    renderAdminInvoiceList();
    updateSummary();
    renderCheckoutStepMode();
    renderAccountBillingOverview();
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
  if (openBuilderPageButton) {
    openBuilderPageButton.classList.toggle("is-hidden", !isAuthed);
  }
  if (openLeadsPageButton) {
    const adminAccessLoading = Boolean(isAuthed && adminAccessStatusPromise);
    openLeadsPageButton.disabled = adminAccessLoading;
    if (adminAccessLoading) {
      openLeadsPageButton.setAttribute("aria-busy", "true");
    } else {
      openLeadsPageButton.removeAttribute("aria-busy");
    }
    openLeadsPageButton.classList.toggle("is-hidden", !isAuthed || !adminAccessAllowed);
  }
  if (openAdminPageButton) {
    const adminAccessLoading = Boolean(isAuthed && adminAccessStatusPromise);
    openAdminPageButton.disabled = adminAccessLoading;
    if (adminAccessLoading) {
      openAdminPageButton.setAttribute("aria-busy", "true");
    } else {
      openAdminPageButton.removeAttribute("aria-busy");
    }
    openAdminPageButton.classList.toggle("is-hidden", !isAuthed || !adminAccessAllowed);
  }
  refreshAdminOnlyTestTools();
  if (openHistoryPageButton) {
    openHistoryPageButton.classList.toggle("is-hidden", !isAuthed);
  }
  if (openReportsPageButton) {
    openReportsPageButton.classList.toggle("is-hidden", !isAuthed);
  }
  if (!isAuthed) {
    setReceiptModalOpen(false);
    setShopifySettingsModalOpen(false);
    setIbanTopupModalOpen(false);
    setClientInviteHistoryModalOpen(false);
    setAdminSettingsModalOpen(false);
    setLeadCallOutcomeModalOpen(false);
    shopifyLocationsCache = [];
    shopifyLocationDraftSelection = new Set();
    shopifySavedLocationSelection = [];
    setClientInviteResult("");
    setClientInviteStatus("");
    if (clientInviteEmailInput) {
      clientInviteEmailInput.value = "";
    }
    leadProspects = [];
    leadProspectsLoaded = false;
    leadProspectsLoading = false;
    leadCallOutcomeLeadId = "";
    renderLeadProspects();
    if (adminCarrierDiscountInput) {
      adminCarrierDiscountInput.value = "";
    }
    if (adminClientDiscountInput) {
      adminClientDiscountInput.value = "";
    }
    if (adminBillingTestEmailInput) {
      adminBillingTestEmailInput.value = "";
    }
    setMainView("builder", { push: false, animate: false });
    resetAccountPreview();
    if (route.view === "register") {
      setAuthMode("register", { inviteToken: route.inviteToken });
      void loadRegistrationInvite(route.inviteToken, {
        preview: Boolean(route.previewRegistration),
        previewStep: route.previewStep,
      });
    } else {
      setAuthMode("login");
      updateRoute({ view: "login" }, { replace: true });
    }
    return;
  }

  setAuthMode("login");
  if (route.view === "login" || route.view === "register") {
    updateRoute({ view: "builder", step: state.step }, { replace: true });
  }
}

async function signInWithPassword() {
  if (!supabaseClient) {
    setAuthMessage(tr("Supabase auth is not configured."));
    return;
  }
  const email = authEmail?.value.trim().toLowerCase() || "";
  const password = authPassword?.value || "";
  if (!email || !password) {
    setAuthMessage(tr("Email and password are required."));
    return;
  }
  setAuthMessage("");
  setAuthBusy(true, tr("Signing In..."), "");
  const { error } = await supabaseClient.auth.signInWithPassword({
    email,
    password,
  });
  setAuthBusy(false, tr("Sign In"), "");
  if (error) {
    setAuthMessage(error.message);
  }
}

function validateRegistrationProfile(profile) {
  const required = [
    profile.companyName,
    profile.contactName,
    profile.contactEmail,
    profile.contactPhone,
    profile.billingAddress,
    profile.taxId,
  ];
  return required.every((value) => Boolean(String(value || "").trim()));
}

async function registerWithInvite() {
  const route = parseRouteFromLocation();
  if (route?.previewRegistration) {
    setAuthMessage(tr("Preview mode only. Registration is disabled on this local route."), {
      isError: false,
      tone: "info",
    });
    return;
  }
  if (!supabaseClient) {
    setAuthMessage(tr("Supabase auth is not configured."));
    return;
  }
  const inviteToken = String(authInviteToken || getInviteTokenFromLocation(window.location)).trim();
  if (!inviteToken) {
    setAuthMessage(tr("Registration link required."));
    return;
  }
  const email = String(authEmail?.value || "").trim().toLowerCase();
  const password = String(authPassword?.value || "");
  const profile = collectRegistrationProfile();
  const stepOneValidationError = getRegistrationStepOneValidationError();
  if (stepOneValidationError) {
    setAuthRegisterStep(1);
    setAuthMessage(stepOneValidationError);
    return;
  }
  if (!authInviteIsValid || !authAgreementContract) {
    setAuthMessage(tr("Could not load registration agreement."));
    return;
  }
  if (!authAgreementHasReachedEnd || !authAgreementAccepted || !authAgreementAgreedAt) {
    setAuthMessage(tr("You must review and accept the agreement before registering."));
    return;
  }
  const agreement = buildRegistrationAgreementPayload();
  if (!agreement) {
    setAuthMessage(tr("Could not load registration agreement."));
    return;
  }

  setAuthMessage("");
  setAuthBusy(true, tr("Creating account..."), "");
  try {
    await fetchPublicApi("/api/auth/register", {
      method: "POST",
      body: JSON.stringify({
        token: inviteToken,
        email,
        password,
        profile,
        agreement,
        preferredLanguage: activeLanguage,
      }),
    });
    setAuthMessage(tr("Account created. Signing you in..."), {
      isError: false,
      tone: "success",
    });
    const { error } = await supabaseClient.auth.signInWithPassword({
      email,
      password,
    });
    if (error) {
      setAuthMessage(error.message);
    }
  } catch (error) {
    setAuthMessage(error?.message || tr("Registration is disabled."));
  } finally {
    setAuthBusy(false, tr("Register Account"), "");
  }
}

async function signUpWithPassword() {
  if (authMode === "register") {
    await registerWithInvite();
    return;
  }
  setAuthMessage(tr("Registration is disabled."));
}

async function resetPassword() {
  if (!supabaseClient) {
    setAuthMessage(tr("Supabase auth is not configured."));
    return;
  }
  const email = authEmail?.value.trim().toLowerCase() || "";
  if (!email) {
    setAuthMessage(tr("Enter your email address, then click Forgot password."));
    return;
  }
  setAuthMessage("");
  setAuthBusy(true, tr("Sign In"), "");
  const { error } = await supabaseClient.auth.resetPasswordForEmail(email, {
    redirectTo: window.location.origin,
  });
  setAuthBusy(false, tr("Sign In"), "");
  if (error) {
    setAuthMessage(error.message);
    return;
  }
  setAuthMessage(tr("If that email exists, a reset link has been sent. Check your inbox."), {
    isError: false,
    tone: "success",
  });
}

async function initializeAuth() {
  if (!supabaseClient) {
    setAuthMessage(tr("Supabase client failed to initialize."));
    setAuthView(null, { animate: false });
    return;
  }
  let session = null;
  let sessionError = null;
  try {
    const firstAttempt = await supabaseClient.auth.getSession();
    session = firstAttempt?.data?.session || null;
    sessionError = firstAttempt?.error || null;
  } catch (error) {
    sessionError = error;
  }
  if (sessionError && !session) {
    await new Promise((resolve) => window.setTimeout(resolve, 140));
    try {
      const retryAttempt = await supabaseClient.auth.getSession();
      session = retryAttempt?.data?.session || null;
      sessionError = retryAttempt?.error || null;
    } catch (error) {
      sessionError = error;
    }
  }
  if (sessionError && !session) {
    setAuthMessage(sessionError?.message || tr("Supabase unavailable"));
  }
  setAuthView(session, { animate: false });
  if (session?.user) {
    consumeShopifyCallbackParams();
  }
  const initialRoute = parseRouteFromLocation();
  const isAuthed = Boolean(session);
  if (isAuthed && initialRoute.view === "account") {
    setMainView("account", { push: false, animate: false });
  } else if (isAuthed && initialRoute.view === "admin") {
    setMainView("admin", { push: false, animate: false });
  } else if (isAuthed && initialRoute.view === "leads") {
    setMainView("leads", { push: false, animate: false });
  } else if (isAuthed && initialRoute.view === "history") {
    setMainView("history", { push: false, animate: false });
  } else if (isAuthed && initialRoute.view === "reports") {
    setMainView("reports", {
      push: false,
      animate: false,
      reportRange: String(initialRoute?.reportRange || ""),
    });
  } else {
    setMainView("builder", { push: false, animate: false });
    if (isAuthed && initialRoute.view === "builder") {
      goToStep(clampStep(initialRoute.step), { push: false, regenerate: false });
    } else if (isAuthed && initialRoute.view !== "builder") {
      updateRoute({ view: "builder", step: state.step }, { replace: true });
    }
  }

  await loadGenerationHistory({
    preferLatest:
      isAuthed &&
      (initialRoute.view === "account" ||
        initialRoute.view === "admin" ||
        initialRoute.view === "leads" ||
        initialRoute.view === "history" ||
        initialRoute.view === "reports"),
  });
  if (isAuthed && initialRoute.view === "admin") {
    await loadAdminDashboard({ quiet: true });
  }
  if (isAuthed && initialRoute.view === "leads") {
    await loadLeadProspects({ quiet: true });
  }
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
    } else if (route.view === "admin") {
      setMainView("admin", { push: false, animate: false });
      await loadAdminDashboard({ quiet: true });
    } else if (route.view === "leads") {
      setMainView("leads", { push: false, animate: false });
      await loadLeadProspects({ quiet: true });
    } else if (route.view === "history") {
      setMainView("history", { push: false, animate: false });
    } else if (route.view === "reports") {
      setMainView("reports", {
        push: false,
        animate: false,
        reportRange: String(route?.reportRange || ""),
      });
    } else {
      setMainView("builder", { push: false, animate: false });
      if (route.view === "builder") {
        goToStep(clampStep(route.step), { push: false, regenerate: false });
      } else {
        updateRoute({ view: "builder", step: state.step }, { replace: true });
      }
    }
    await loadGenerationHistory({
      preferLatest:
        route.view === "account" ||
        route.view === "admin" ||
        route.view === "leads" ||
        route.view === "history" ||
        route.view === "reports",
    });
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
  void mountFlowLogoAnimation(target);
}

function initializeAppLogo() {
  if (!window.lottie || !appLogoLottie) return;
  void mountFlowLogoAnimation(appLogoLottie);
}

function clamp01(value) {
  return Math.max(0, Math.min(1, value));
}

function parseHexColorToRgb(value) {
  const hex = String(value || "").trim().replace(/^#/, "");
  if (!hex) return { r: 119, g: 71, b: 227 };
  if (hex.length === 3 || hex.length === 4) {
    return {
      r: Number.parseInt(hex[0] + hex[0], 16) || 119,
      g: Number.parseInt(hex[1] + hex[1], 16) || 71,
      b: Number.parseInt(hex[2] + hex[2], 16) || 227,
    };
  }
  if (hex.length === 6 || hex.length === 8) {
    return {
      r: Number.parseInt(hex.slice(0, 2), 16) || 119,
      g: Number.parseInt(hex.slice(2, 4), 16) || 71,
      b: Number.parseInt(hex.slice(4, 6), 16) || 227,
    };
  }
  return { r: 119, g: 71, b: 227 };
}

function mixRgbColors(source, target, amount) {
  const t = clamp01(amount);
  return {
    r: Math.round(source.r + (target.r - source.r) * t),
    g: Math.round(source.g + (target.g - source.g) * t),
    b: Math.round(source.b + (target.b - source.b) * t),
  };
}

function cloneAnimationData(data) {
  if (!data) return null;
  if (typeof structuredClone === "function") {
    return structuredClone(data);
  }
  return JSON.parse(JSON.stringify(data));
}

async function mountFlowLogoAnimation(target) {
  if (!window.lottie || !target) return null;
  target.style.pointerEvents = "none";
  target.textContent = "";
  target.classList.remove("is-ready");
  const animationData = cloneAnimationData(await flowLogoAnimationDataPromise);
  const config = {
    container: target,
    renderer: "svg",
    loop: false,
    autoplay: true,
    rendererSettings: {
      progressiveLoad: true,
    },
  };
  if (animationData) {
    config.animationData = animationData;
  } else {
    config.path = FLOW_LOGO_JSON_URL;
  }
  const anim = window.lottie.loadAnimation(config);
  anim.addEventListener("DOMLoaded", () => {
    const svg = target.querySelector("svg");
    if (svg) svg.style.pointerEvents = "none";
    target.classList.add("is-ready");
  });
  return anim;
}

class PixelCanvasPixel {
  constructor(canvas, context, x, y, color, speed, delay) {
    this.width = canvas.width;
    this.height = canvas.height;
    this.ctx = context;
    this.x = x;
    this.y = y;
    this.color = color;
    this.colorRgb = parseHexColorToRgb(color);
    this.speed = Math.random() * 0.8 * speed + speed * 0.1;
    this.size = 0;
    this.sizeStep = Math.random() * 0.4 + 0.08;
    this.minSize = 0.45;
    this.maxSizeInteger = 2;
    this.maxSize = Math.random() * (this.maxSizeInteger - this.minSize) + this.minSize;
    this.delay = delay;
    this.counter = 0;
    this.counterStep = Math.random() * 4 + (this.width + this.height) * 0.01;
    this.isIdle = false;
    this.isReverse = false;
    this.isShimmer = false;
    this.phase = Math.random() * Math.PI * 2;
  }

  draw(size = this.size, alpha = 1, color = this.color) {
    if (size <= 0 || alpha <= 0) return;
    const centerOffset = this.maxSizeInteger * 0.5 - size * 0.5;
    this.ctx.save();
    this.ctx.globalAlpha = alpha;
    this.ctx.fillStyle = color;
    this.ctx.fillRect(
      this.x + centerOffset,
      this.y + centerOffset,
      size,
      size
    );
    this.ctx.restore();
  }

  appear() {
    this.isIdle = false;
    if (this.counter <= this.delay) {
      this.counter += this.counterStep;
      return;
    }
    if (this.size >= this.maxSize) {
      this.isShimmer = true;
    }
    if (this.isShimmer) {
      this.shimmer();
    } else {
      this.size = Math.min(this.maxSize, this.size + this.sizeStep);
    }
    this.draw();
  }

  disappear() {
    this.isShimmer = false;
    this.counter = 0;
    if (this.size <= 0) {
      this.isIdle = true;
      return;
    }
    this.size = Math.max(0, this.size - 0.12);
    this.draw();
  }

  shimmer() {
    if (this.speed <= 0) {
      this.size = this.maxSize;
      return;
    }
    if (this.size >= this.maxSize) {
      this.isReverse = true;
    } else if (this.size <= this.minSize) {
      this.isReverse = false;
    }
    if (this.isReverse) {
      this.size = Math.max(this.minSize, this.size - this.speed);
    } else {
      this.size = Math.min(this.maxSize, this.size + this.speed);
    }
  }

  renderWave(elapsed, options = {}) {
    const sinceReveal = elapsed - this.delay;
    if (sinceReveal <= 0) return;

    const revealDuration = Number(options.revealDuration) || 170;
    const frontWidth = Number(options.frontWidth) || 120;
    const reducedMotion = Boolean(options.reducedMotion);
    const progress = Math.min(1, sinceReveal / revealDuration);
    const eased = progress * progress * (3 - 2 * progress);
    const localIntro = reducedMotion ? 1 : Math.min(1, sinceReveal / 110);
    const localEase = localIntro * localIntro * (3 - 2 * localIntro);
    const frontBoost = Math.max(0, 1 - sinceReveal / frontWidth);
    const introOpacity = reducedMotion ? 1 : Math.min(1, elapsed / 520);
    const shimmer =
      progress >= 1 && !reducedMotion
        ? 0.96 + Math.sin(elapsed * 0.011 + this.phase) * 0.08
        : 1;
    const settledAlpha = reducedMotion ? 1 : 0.88;
    const rimGlow = Math.pow(frontBoost, 1.15);
    const highlightAmount = reducedMotion ? 0 : rimGlow * 0.72;
    const waveColor =
      highlightAmount > 0
        ? (() => {
            const mixed = mixRgbColors(this.colorRgb, { r: 244, g: 236, b: 255 }, highlightAmount);
            return `rgb(${mixed.r} ${mixed.g} ${mixed.b})`;
          })()
        : this.color;
    const size = Math.min(
      this.maxSizeInteger * 1.45,
      this.maxSize * (0.03 + eased * 0.97) * (1 + frontBoost * 0.16) * shimmer * introOpacity * localEase
    );
    const alpha = Math.min(1, (0.02 + eased * settledAlpha + frontBoost * 0.16) * introOpacity * localEase);
    this.draw(size, alpha, waveColor);
  }
}

class PixelCanvasElement extends HTMLElement {
  static register(tag = "pixel-canvas") {
    if (!("customElements" in window) || customElements.get(tag)) return;
    customElements.define(tag, this);
  }

  get colors() {
    const raw = String(this.dataset.colors || "").trim();
    const colors = raw
      ? raw
          .split(raw.includes("|") ? "|" : ",")
          .map((entry) => entry.trim())
          .filter(Boolean)
      : [];
    return colors.length
      ? colors
      : ["#7747e38c", "#7747e3ad", "#7747e3d1", "#7747e3"];
  }

  get gap() {
    const value = Number(this.dataset.gap || 6);
    return Math.max(4, Math.min(48, Number.isFinite(value) ? value : 6));
  }

  get speed() {
    const value = Number(this.dataset.speed || 20);
    if (!Number.isFinite(value) || value <= 0 || this.reducedMotion) return 0;
    return Math.min(100, value) * 0.001;
  }

  get noFocus() {
    return this.hasAttribute("data-no-focus");
  }

  get autoStart() {
    return this.hasAttribute("data-auto-start");
  }

  get startDelay() {
    const value = Number(this.dataset.startDelay || 0);
    return Number.isFinite(value) ? Math.max(0, value) : 0;
  }

  get wave() {
    return String(this.dataset.wave || "").trim().toLowerCase();
  }

  connectedCallback() {
    if (this.shadowRoot) return;

    const root = this.attachShadow({ mode: "open" });
    const style = document.createElement("style");
    style.textContent = `
      :host {
        display: block;
        inline-size: 100%;
        block-size: 100%;
        overflow: hidden;
      }
      canvas {
        display: block;
        inline-size: 100%;
        block-size: 100%;
      }
    `;
    this.canvas = document.createElement("canvas");
    root.append(style, this.canvas);
    this.ctx = this.canvas.getContext("2d", { alpha: true });
    if (!this.ctx) return;

    this.reducedMotion = prefersReducedMotion();
    this.timeInterval = 1000 / 60;
    this.timePrevious = performance.now();
    this.interactionTarget = this.parentElement || this;
    this.startTimer = null;
    this.isPaused = false;

    this.init();

    this.resizeObserver = new ResizeObserver(() => {
      this.init();
      if (this.reducedMotion) {
        this.renderStatic();
      } else if (this.autoStart) {
        this.scheduleAutoStart();
      }
    });
    this.resizeObserver.observe(this);

    if (!this.autoStart) {
      this.interactionTarget.addEventListener("mouseenter", this);
      this.interactionTarget.addEventListener("mouseleave", this);
      if (!this.noFocus) {
        this.interactionTarget.addEventListener("focusin", this);
        this.interactionTarget.addEventListener("focusout", this);
      }
    }

    if (this.reducedMotion) {
      this.renderStatic();
    } else if (this.autoStart) {
      this.scheduleAutoStart();
    }
  }

  disconnectedCallback() {
    cancelAnimationFrame(this.animation);
    if (this.startTimer) {
      window.clearTimeout(this.startTimer);
      this.startTimer = null;
    }
    if (this.resizeObserver) {
      this.resizeObserver.disconnect();
    }
    if (this.interactionTarget && !this.autoStart) {
      this.interactionTarget.removeEventListener("mouseenter", this);
      this.interactionTarget.removeEventListener("mouseleave", this);
      if (!this.noFocus) {
        this.interactionTarget.removeEventListener("focusin", this);
        this.interactionTarget.removeEventListener("focusout", this);
      }
    }
  }

  handleEvent(event) {
    const handler = this[`on${event.type}`];
    if (typeof handler === "function") {
      handler.call(this, event);
    }
  }

  setPaused(paused) {
    const nextPaused = Boolean(paused);
    if (this.isPaused === nextPaused) return;
    this.isPaused = nextPaused;
    if (nextPaused) {
      cancelAnimationFrame(this.animation);
      if (this.startTimer) {
        window.clearTimeout(this.startTimer);
        this.startTimer = null;
      }
      return;
    }
    this.timePrevious = performance.now();
    if (this.reducedMotion) {
      this.renderStatic();
      return;
    }
    if (this.autoStart) {
      this.scheduleAutoStart();
      return;
    }
    this.handleAnimation("appear");
  }

  onmouseenter() {
    this.handleAnimation("appear");
  }

  onmouseleave() {
    this.handleAnimation("disappear");
  }

  onfocusin(event) {
    if (event.currentTarget.contains(event.relatedTarget)) return;
    this.handleAnimation("appear");
  }

  onfocusout(event) {
    if (event.currentTarget.contains(event.relatedTarget)) return;
    this.handleAnimation("disappear");
  }

  scheduleAutoStart() {
    if (this.startTimer) {
      window.clearTimeout(this.startTimer);
      this.startTimer = null;
    }
    const delay = this.startDelay;
    if (!delay) {
      this.handleAnimation("appear");
      return;
    }
    this.startTimer = window.setTimeout(() => {
      this.startTimer = null;
      this.handleAnimation("appear");
    }, delay);
  }

  init() {
    if (!this.ctx) return;
    const rect = this.getBoundingClientRect();
    const width = Math.max(1, Math.floor(rect.width));
    const height = Math.max(1, Math.floor(rect.height));
    this.canvas.width = width;
    this.canvas.height = height;
    this.canvas.style.width = `${width}px`;
    this.canvas.style.height = `${height}px`;
    this.pixels = [];
    this.updateWaveOrigin();
    this.createPixels();
  }

  getDistanceToCanvasCenter(x, y) {
    const dx = x - this.canvas.width / 2;
    const dy = y - this.canvas.height / 2;
    return Math.sqrt(dx * dx + dy * dy);
  }

  updateWaveOrigin() {
    const fallback = {
      x: this.canvas.width / 2,
      y: this.canvas.height / 2,
    };
    const anchorElement = authStack || authCard;
    if (!anchorElement) {
      this.waveOrigin = fallback;
      this.waveMaxDistance = Math.hypot(this.canvas.width / 2, this.canvas.height / 2);
      return;
    }

    const hostRect = this.getBoundingClientRect();
    const anchorRect = anchorElement.getBoundingClientRect();
    const hasAnchorBox = anchorRect.width > 0 && anchorRect.height > 0;
    const origin = hasAnchorBox
      ? {
          x: anchorRect.left - hostRect.left + anchorRect.width / 2,
          y: anchorRect.top - hostRect.top + anchorRect.height / 2,
        }
      : fallback;

    this.waveOrigin = origin;
    this.waveMaxDistance = Math.max(
      1,
      Math.hypot(origin.x, origin.y),
      Math.hypot(this.canvas.width - origin.x, origin.y),
      Math.hypot(origin.x, this.canvas.height - origin.y),
      Math.hypot(this.canvas.width - origin.x, this.canvas.height - origin.y)
    );
  }

  getPixelDelay(x, y, gap) {
    if (this.reducedMotion) return 0;

    if (this.wave === "radial") {
      const origin = this.waveOrigin || { x: this.canvas.width / 2, y: this.canvas.height / 2 };
      const dx = x - origin.x;
      const dy = y - origin.y;
      const distance = Math.sqrt(dx * dx + dy * dy);
      const maxDistance = Math.max(1, this.waveMaxDistance || Math.hypot(this.canvas.width / 2, this.canvas.height / 2));
      const normalized = clamp01(distance / maxDistance);
      const curved = normalized * normalized * (0.35 + normalized * 0.65);
      return 18 + curved * maxDistance * 1.55 + Math.random() * gap * 1.8;
    }

    if (this.wave === "sweep") {
      const width = Math.max(1, this.canvas.width);
      const height = Math.max(1, this.canvas.height);
      const normalizedX = x / width;
      const normalizedY = y / height;
      const waveBand =
        Math.sin(normalizedY * Math.PI * 2.4 - normalizedX * Math.PI * 0.9) * gap * 5.5;
      const verticalDrift = normalizedY * gap * 3.2;
      const horizontalTravel = normalizedX * width * 0.9;
      return Math.max(0, horizontalTravel + waveBand + verticalDrift);
    }

    return this.getDistanceToCanvasCenter(x, y);
  }

  getEffectiveGap() {
    const area = Math.max(1, this.canvas.width * this.canvas.height);
    const maxPixels = 42000;
    const dynamicGap = Math.ceil(Math.sqrt(area / maxPixels));
    return Math.max(this.gap, Number.isFinite(dynamicGap) ? dynamicGap : this.gap);
  }

  createPixels() {
    const gap = this.getEffectiveGap();
    this.pixelGap = gap;
    for (let x = 0; x < this.canvas.width; x += gap) {
      for (let y = 0; y < this.canvas.height; y += gap) {
        const color = this.colors[Math.floor(Math.random() * this.colors.length)];
        const delay = this.getPixelDelay(x, y, gap);
        this.pixels.push(
          new PixelCanvasPixel(this.canvas, this.ctx, x, y, color, this.speed, delay)
        );
      }
    }
  }

  renderStatic() {
    if (!this.ctx) return;
    cancelAnimationFrame(this.animation);
    this.ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);
    this.pixels.forEach((pixel) => {
      pixel.counter = pixel.delay + 1;
      pixel.size = pixel.maxSize;
      pixel.draw();
    });
  }

  handleAnimation(name) {
    cancelAnimationFrame(this.animation);
    this.timePrevious = performance.now();
    if (this.autoStart && name === "appear") {
      this.updateWaveOrigin();
      if (this.wave === "radial" && this.pixelGap) {
        for (let i = 0; i < this.pixels.length; i += 1) {
          this.pixels[i].delay = this.getPixelDelay(this.pixels[i].x, this.pixels[i].y, this.pixelGap);
        }
      }
      this.waveStartAt = this.timePrevious;
    }
    this.animation = requestAnimationFrame(() => this.animate(name));
  }

  renderWaveFrame(timeNow) {
    const elapsed = Math.max(0, timeNow - (this.waveStartAt || timeNow));
    this.ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);
    for (let i = 0; i < this.pixels.length; i += 1) {
      this.pixels[i].renderWave(elapsed, {
        revealDuration: 235,
        frontWidth: 240,
        reducedMotion: this.reducedMotion,
      });
    }
  }

  animate(fnName) {
    if (this.isPaused) return;
    this.animation = requestAnimationFrame(() => this.animate(fnName));
    const timeNow = performance.now();
    const timePassed = timeNow - this.timePrevious;
    if (timePassed < this.timeInterval) return;

    this.timePrevious = timeNow - (timePassed % this.timeInterval);

    if (this.autoStart && fnName === "appear") {
      this.renderWaveFrame(timeNow);
      return;
    }

    this.ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);

    for (let i = 0; i < this.pixels.length; i += 1) {
      this.pixels[i][fnName]();
    }

    if (this.pixels.every((pixel) => pixel.isIdle)) {
      cancelAnimationFrame(this.animation);
    }
  }
}

function normalizeAuthBackgroundVariant(value) {
  return String(value || "").trim().toLowerCase() === "blast" ? "blast" : "matrix";
}

function readStoredAuthBackgroundVariant() {
  try {
    return normalizeAuthBackgroundVariant(window.localStorage.getItem(AUTH_BACKGROUND_VARIANT_STORAGE_KEY));
  } catch (error) {
    return "matrix";
  }
}

function storeAuthBackgroundVariant(variant) {
  try {
    window.localStorage.setItem(
      AUTH_BACKGROUND_VARIANT_STORAGE_KEY,
      normalizeAuthBackgroundVariant(variant)
    );
  } catch (error) {
    // Ignore storage issues.
  }
}

function updateAuthBackgroundVariantUi(variant) {
  const normalized = normalizeAuthBackgroundVariant(variant);
  if (authGate) {
    authGate.dataset.authBgVariant = normalized;
  }
  authBgVariantButtons.forEach((button) => {
    const isActive = normalizeAuthBackgroundVariant(button.dataset.authBgVariant) === normalized;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-pressed", isActive ? "true" : "false");
  });
}

async function ensureAuthBlastBackground() {
  if (authBlastBackground) return authBlastBackground;
  if (!authBlastCanvas) return null;
  if (authBlastBackgroundLoading) return authBlastBackgroundLoading;
  authBlastBackgroundLoading = import("./auth-pixel-blast.js")
    .then((mod) => mod.createAuthPixelBlastBackground(authBlastCanvas, {
      color: "#7747e3",
      interactionTarget: authGate || document.body,
      variant: "square",
      pixelSize: 2,
      patternScale: 1,
      patternDensity: 1.4,
      pixelSizeJitter: 1.6,
      enableRipples: true,
      rippleSpeed: 0.4,
      rippleThickness: 0.12,
      rippleIntensityScale: 1.5,
      speed: 0.95,
      edgeFade: 0.31,
      transparent: true,
    }))
    .then((instance) => {
      authBlastBackground = instance;
      return instance;
    })
    .catch((error) => {
      authBlastBackgroundLoading = null;
      throw error;
    });
  return authBlastBackgroundLoading;
}

async function setAuthBackgroundVariant(variant, options = {}) {
  const normalized = normalizeAuthBackgroundVariant(variant);
  const { persist = true } = options;
  authBackgroundVariant = normalized;
  updateAuthBackgroundVariantUi(normalized);
  if (persist) {
    storeAuthBackgroundVariant(normalized);
  }
  if (normalized === "blast") {
    try {
      const blast = await ensureAuthBlastBackground();
      blast?.setPaused(false);
      if (authPixelCanvas?.setPaused) {
        authPixelCanvas.setPaused(true);
      }
    } catch (error) {
      authBackgroundVariant = "matrix";
      updateAuthBackgroundVariantUi("matrix");
      if (persist) {
        storeAuthBackgroundVariant("matrix");
      }
      showToast("Could not load the Blast background.", { tone: "error" });
      if (authPixelCanvas?.setPaused) {
        authPixelCanvas.setPaused(false);
      }
    }
    return;
  }
  authBlastBackground?.setPaused(true);
  if (authPixelCanvas?.setPaused) {
    authPixelCanvas.setPaused(false);
  }
}

function initializeAuthBackground() {
  if (authBackgroundStarted || !authGate || !authPixelCanvas) return;
  authBackgroundStarted = true;
  PixelCanvasElement.register();
  authBackgroundVariant = readStoredAuthBackgroundVariant();
  updateAuthBackgroundVariantUi(authBackgroundVariant);
  if (authBackgroundVariant === "blast" && authPixelCanvas?.setPaused) {
    authPixelCanvas.setPaused(true);
  }
  authBgSwitcher?.addEventListener("click", (event) => {
    const button = event.target instanceof Element
      ? event.target.closest("[data-auth-bg-variant]")
      : null;
    if (!(button instanceof HTMLElement)) return;
    const nextVariant = normalizeAuthBackgroundVariant(button.dataset.authBgVariant);
    void setAuthBackgroundVariant(nextVariant);
  });
  requestAnimationFrame(() => {
    authGate.classList.add("is-ready");
    void setAuthBackgroundVariant(authBackgroundVariant, { persist: false });
  });
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
  } else {
    state.shipFromLockedByProvider = false;
    state.csvSource = "none";
    state.csvPage = 1;
    state.csvValidationAttempted = false;
  }
  renderCsvShipFromSelector();
  renderCsvTable();
}

function setCsvEditMode(enabled) {
  state.csvEditable = enabled;
  if (csvEditToggle) {
    csvEditToggle.classList.toggle("is-active", enabled);
    const label = csvEditToggle.querySelector("span");
    if (label) label.textContent = enabled ? tr("Editing") : tr("Edit");
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

function getCsvPageCount() {
  const totalRows = Array.isArray(state.csvRows) ? state.csvRows.length : 0;
  return Math.max(1, Math.ceil(totalRows / CSV_TABLE_PAGE_SIZE));
}

function setCsvPage(page, { render = true } = {}) {
  const totalPages = getCsvPageCount();
  const nextPage = Math.max(1, Math.min(totalPages, Number(page) || 1));
  state.csvPage = nextPage;
  if (render) {
    renderCsvTable();
  }
}

function renderCsvTablePager(totalRows, startIndex, endIndex, totalPages) {
  if (!csvPagePrev || !csvPageNext || !csvPageMeta) return;
  const hasRows = totalRows > 0;
  const activePage = Math.max(1, Math.min(totalPages, state.csvPage || 1));
  csvPagePrev.disabled = !hasRows || activePage <= 1;
  csvPageNext.disabled = !hasRows || activePage >= totalPages;

  if (!hasRows) {
    csvPageMeta.textContent = tr("No rows loaded");
    return;
  }

  csvPageMeta.textContent = tr("Rows {start}-{end} • Page {page}/{pages}", {
    start: startIndex + 1,
    end: endIndex,
    page: activePage,
    pages: totalPages,
  });
}

function renderCsvTable() {
  if (!csvTableBody) return;
  const totalRows = Array.isArray(state.csvRows) ? state.csvRows.length : 0;
  const totalPages = getCsvPageCount();
  const safePage = Math.max(1, Math.min(totalPages, state.csvPage || 1));
  state.csvPage = safePage;

  const startIndex = Math.max(0, (safePage - 1) * CSV_TABLE_PAGE_SIZE);
  const endIndex = Math.min(totalRows, startIndex + CSV_TABLE_PAGE_SIZE);
  const visibleRows = state.csvRows.slice(startIndex, endIndex);

  csvTableBody.innerHTML = "";
  visibleRows.forEach((row, offset) => {
    const rowIndex = startIndex + offset;
    const rowEl = document.createElement("tr");
    csvReviewColumns.forEach((column) => {
      const td = document.createElement("td");
      const input = document.createElement("input");
      input.type = "text";
      input.value = row[column.key] ?? "";
      input.dataset.row = String(rowIndex);
      input.dataset.key = column.key;
      input.disabled = !state.csvEditable;
      if (state.csvValidationAttempted && CSV_REQUIRED_FIELDS.has(column.key)) {
        input.classList.toggle("is-invalid", String(row[column.key] ?? "").trim() === "");
      }
      bindCompositionAwareInput(input, handleCsvInput);

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
      rowEl.appendChild(td);
    });
    csvTableBody.appendChild(rowEl);
  });

  renderCsvTablePager(totalRows, startIndex, endIndex, totalPages);
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
  setInlineFormErrorToast(labelError, "");
  if (key === "recipientCountry") {
    const icon = input.parentElement?.querySelector(".csv-country-icon");
    if (icon) {
      icon.innerHTML = getCountryIcon(input.value.trim());
    }
  }
  if (state.csvValidationAttempted && CSV_REQUIRED_FIELDS.has(key)) {
    input.classList.toggle("is-invalid", input.value.trim() === "");
  }
  updateSummary();
  updatePreview();
}

function validateCsvRows() {
  if (!Array.isArray(state.csvRows) || state.csvRows.length === 0) {
    return false;
  }
  state.csvValidationAttempted = true;
  let isValid = true;
  let firstInvalidRow = -1;

  state.csvRows.forEach((row, rowIndex) => {
    CSV_REQUIRED_FIELDS.forEach((key) => {
      const value = String(row?.[key] ?? "").trim();
      if (value) return;
      isValid = false;
      if (firstInvalidRow < 0) {
        firstInvalidRow = rowIndex;
      }
    });
  });

  if (!isValid && firstInvalidRow >= 0) {
    state.csvPage = Math.floor(firstInvalidRow / CSV_TABLE_PAGE_SIZE) + 1;
  }
  renderCsvTable();
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

function getSummaryCarrierLabel() {
  return "Bpost";
}

function getSummaryDiscountRate() {
  if (billingOverview && Number.isFinite(Number(billingOverview.client_discount_pct))) {
    return Number(billingOverview.client_discount_pct);
  }
  const service = normalizeNameKey(state.selection.type);
  if (service === normalizeNameKey("Economy")) return 21;
  if (service === normalizeNameKey("Priority")) return 16;
  if (service === normalizeNameKey("International Express")) return 13;
  return 15;
}

function formatSummaryInvoiceDate(rawDate) {
  const candidate = new Date(String(rawDate || "").trim());
  if (Number.isNaN(candidate.getTime())) {
    const now = new Date();
    const fallback = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    return fallback.toLocaleDateString(getUiLocale(), {
      day: "2-digit",
      month: "short",
      year: "numeric",
    });
  }
  return candidate.toLocaleDateString(getUiLocale(), {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
}

function updateSummary() {
  const discountRate = getSummaryDiscountRate();
  const nextExVat = Number(billingOverview?.next_invoice_amount_ex_vat);
  const nextDate = String(billingOverview?.next_invoice_date || "").trim();
  const projectedInvoiceAmount = Number.isFinite(nextExVat)
    ? nextExVat
    : Number((42 + state.selection.price * getQuantity() * 3.2).toFixed(2));
  const paymentMode = String(billingOverview?.payment_mode || "").trim().toLowerCase();

  if (summaryService) {
    summaryService.textContent = getSummaryCarrierLabel();
  }
  if (summaryPrice) {
    summaryPrice.textContent = `${discountRate}%`;
  }
  if (summaryQty) {
    summaryQty.textContent = formatMoney(projectedInvoiceAmount);
  }
  if (summaryTotal) {
    summaryTotal.textContent = formatSummaryInvoiceDate(nextDate);
  }
  if (summaryTracking) {
    summaryTracking.textContent = tr("Chat with us");
  }
}

function openSupportChat() {
  window.location.href =
    "mailto:support@shipide.com?subject=Shipide%20Portal%20Support&body=Hi%20Shipide%20team%2C%0A%0AI%20need%20help%20with%20my%20account.";
}

if (summaryChatButton) {
  summaryChatButton.addEventListener("click", openSupportChat);
}

function updatePayment() {
  const { quantity, subtotal, vatAmount, total } = getOrderTotals();

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
  if (directPaymentTotal) {
    directPaymentTotal.textContent = formatMoney(total);
  }
  renderCheckoutStepMode();
}

function getOrderTotals() {
  const quantity = getQuantity();
  const subtotal = Number((state.selection.price * quantity).toFixed(2));
  const vatAmount = Number((subtotal * VAT_RATE).toFixed(2));
  const total = Number((subtotal + vatAmount).toFixed(2));
  return { quantity, subtotal, vatAmount, total };
}

function getIbanInstructionsFromOverview() {
  const fallback = {
    beneficiary: "Shipide",
    iban: "BE68 5390 0754 7034",
    bic: "KREDBEBB",
    note: "",
  };
  const fromOverview = billingOverview?.iban_instructions;
  if (!fromOverview || typeof fromOverview !== "object") return fallback;
  return {
    beneficiary: String(fromOverview.beneficiary || fallback.beneficiary).trim() || fallback.beneficiary,
    iban: String(fromOverview.iban || fallback.iban).trim() || fallback.iban,
    bic: String(fromOverview.bic || fallback.bic).trim() || fallback.bic,
    note: String(fromOverview.note || fallback.note).trim() || fallback.note,
  };
}

function getReusableIbanTopupFromOverview() {
  const topups = Array.isArray(billingOverview?.recent_topups) ? billingOverview.recent_topups : [];
  const topup = topups.find((entry) => String(entry?.status || "").trim().toLowerCase() === "pending");
  if (!topup) return null;
  const requestedAt = Date.parse(String(topup?.requested_at || ""));
  if (Number.isFinite(requestedAt) && Date.now() - requestedAt > IBAN_TOPUP_REFERENCE_REUSE_MAX_AGE_MS) {
    return null;
  }
  const reference = String(topup?.reference || "").trim();
  if (!reference) return null;
  const instructions = getIbanInstructionsFromOverview();
  return {
    topup,
    instructions: {
      ...instructions,
      reference,
      amount_eur: Number.isFinite(Number(topup?.amount_eur)) ? Number(topup.amount_eur) : null,
      currency: String(topup?.currency || "EUR").trim() || "EUR",
    },
  };
}

function getBillingFlags() {
  let invoiceEnabled = billingOverview?.invoice_enabled === true;
  let cardEnabled = billingOverview?.card_enabled === false ? false : true;
  const walletEnabled = billingOverview?.wallet_enabled !== false;
  if (!invoiceEnabled && !cardEnabled) {
    cardEnabled = true;
  }
  return {
    invoiceEnabled,
    cardEnabled,
    walletEnabled,
  };
}

function detectCardBrand(numberDigits) {
  const digits = String(numberDigits || "").replace(/\D+/g, "");
  if (!digits) return "generic";
  if (/^4/.test(digits)) return "visa";
  if (/^(5[1-5]|2(2[2-9]|[3-6]\d|7[01]|720))/.test(digits)) return "mastercard";
  if (/^3[47]/.test(digits)) return "amex";
  return "generic";
}

function getCardBrandLabel(brand) {
  if (brand === "visa") return "VISA";
  if (brand === "mastercard") return "MASTERCARD";
  if (brand === "amex") return "AMEX";
  return "CARD";
}

function getCardNumberMaxLength(brand) {
  return brand === "amex" ? 15 : 16;
}

function getCardCvcMaxLength(brand) {
  return brand === "amex" ? 4 : 3;
}

function formatCardNumberForDisplay(digits, brand) {
  const safe = String(digits || "").replace(/\D+/g, "");
  if (!safe) return "";
  if (brand === "amex") {
    const p1 = safe.slice(0, 4);
    const p2 = safe.slice(4, 10);
    const p3 = safe.slice(10, 15);
    return [p1, p2, p3].filter(Boolean).join(" ");
  }
  return safe.match(/.{1,4}/g)?.join(" ") || safe;
}

function formatCardExpiryForDisplay(raw) {
  const digits = String(raw || "").replace(/\D+/g, "").slice(0, 4);
  if (digits.length <= 2) return digits;
  return `${digits.slice(0, 2)}/${digits.slice(2, 4)}`;
}

function updateCardBrandVisual(brand = "generic") {
  if (!paymentCardBrandPill) return;
  paymentCardBrandPill.classList.remove("is-visa", "is-mastercard", "is-amex", "is-generic");
  paymentCardBrandPill.classList.add(
    brand === "visa" || brand === "mastercard" || brand === "amex" ? `is-${brand}` : "is-generic"
  );
  paymentCardBrandPill.textContent = getCardBrandLabel(brand);
  if (cardCvcInput) {
    const max = getCardCvcMaxLength(brand);
    cardCvcInput.maxLength = max;
    if (String(cardCvcInput.value || "").length > max) {
      cardCvcInput.value = String(cardCvcInput.value || "").slice(0, max);
    }
  }
}

function updateCardFormVisibility() {
  const canExpand =
    checkoutPaymentMethod === "card" &&
    paymentMethodCard &&
    !paymentMethodCard.classList.contains("is-disabled");
  if (paymentMethodCard) {
    paymentMethodCard.classList.toggle("is-card-expanded", Boolean(canExpand));
  }
  if (paymentCardExpand) {
    paymentCardExpand.setAttribute("aria-hidden", canExpand ? "false" : "true");
  }
}

function validateCardDetails() {
  const numberDigits = String(cardNumberInput?.value || "").replace(/\D+/g, "");
  const brand = detectCardBrand(numberDigits);
  const numberLen = numberDigits.length;
  const expectedNumberLen = getCardNumberMaxLength(brand);
  if (numberLen < expectedNumberLen) {
    return {
      ok: false,
      message: tr("Enter a valid card number."),
    };
  }

  const expiryDigits = String(cardExpiryInput?.value || "").replace(/\D+/g, "");
  if (expiryDigits.length < 4) {
    return {
      ok: false,
      message: tr("Enter a valid expiry date."),
    };
  }
  const expMonth = Number(expiryDigits.slice(0, 2));
  const expYear = Number(expiryDigits.slice(2, 4));
  if (!Number.isFinite(expMonth) || expMonth < 1 || expMonth > 12) {
    return {
      ok: false,
      message: tr("Enter a valid expiry month."),
    };
  }
  const now = new Date();
  const nowMonth = now.getMonth() + 1;
  const nowYear = now.getFullYear() % 100;
  if (expYear < nowYear || (expYear === nowYear && expMonth < nowMonth)) {
    return {
      ok: false,
      message: tr("Card has expired."),
    };
  }

  const cvc = String(cardCvcInput?.value || "").replace(/\D+/g, "");
  const cvcExpected = getCardCvcMaxLength(brand);
  if (cvc.length < cvcExpected) {
    return {
      ok: false,
      message: tr("Enter a valid CVC."),
    };
  }

  const holder = String(cardHolderInput?.value || "").trim();
  if (holder.length < 2) {
    return {
      ok: false,
      message: tr("Enter the cardholder name."),
    };
  }

  return {
    ok: true,
    brand,
    last4: numberDigits.slice(-4),
  };
}

function setCheckoutPaymentMethod(method, options = {}) {
  const safeMethod = String(method || "").trim().toLowerCase();
  const { invoiceEnabled, cardEnabled, walletEnabled } = getBillingFlags();
  let nextMethod = "invoice";
  if (safeMethod === "wallet" && walletEnabled) {
    nextMethod = "wallet";
  } else if (safeMethod === "card" && cardEnabled) {
    nextMethod = "card";
  } else if (safeMethod === "invoice" && invoiceEnabled) {
    nextMethod = "invoice";
  } else if (invoiceEnabled) {
    nextMethod = "invoice";
  } else if (walletEnabled) {
    nextMethod = "wallet";
  } else if (cardEnabled) {
    nextMethod = "card";
  }
  checkoutPaymentMethod = nextMethod;

  if (paymentMethodList) {
    paymentMethodList.querySelectorAll("[data-payment-method]").forEach((button) => {
      const card = button;
      const cardMethod = String(card.dataset.paymentMethod || "").trim().toLowerCase();
      const isActive = cardMethod === checkoutPaymentMethod;
      card.classList.toggle("is-selected", isActive);
      const isDisabled =
        cardMethod === "wallet"
          ? !walletEnabled
          : cardMethod === "card"
            ? !cardEnabled
            : false;
      card.classList.toggle("is-disabled", isDisabled);
      card.setAttribute("aria-pressed", isActive ? "true" : "false");
      card.setAttribute("aria-disabled", isDisabled ? "true" : "false");
      card.setAttribute("tabindex", isDisabled ? "-1" : "0");
    });
  }

  if (!options.quiet) {
    setInlineFormErrorToast(paymentMethodError, "");
  }
  updateCardFormVisibility();
}

function getTopupStatusMeta(status) {
  const key = String(status || "").trim().toLowerCase();
  if (key === "credited") return { label: tr("Credited"), className: "is-success" };
  if (key === "received") return { label: tr("Received"), className: "is-info" };
  if (key === "cancelled") return { label: tr("Cancelled"), className: "is-muted" };
  if (key === "failed") return { label: tr("Failed"), className: "is-danger" };
  return { label: tr("Pending"), className: "is-warning" };
}

function renderTopupReferenceHistory() {
  if (!accountReferenceHistoryList) return;
  const rows = Array.isArray(billingOverview?.recent_topups) ? billingOverview.recent_topups : [];
  if (!rows.length) {
    accountReferenceHistoryList.innerHTML = `<div class="account-billing-empty">${escapeHtml(
      tr("No transfer references yet.")
    )}</div>`;
    return;
  }

  accountReferenceHistoryList.innerHTML = "";
  rows.slice(0, 30).forEach((row) => {
    const reference = String(row?.reference || "--").trim();
    const timestamp = formatHistoryDate(row?.requested_at || row?.created_at || "");
    const statusMeta = getTopupStatusMeta(row?.status);
    const item = document.createElement("article");
    item.className = "account-reference-item";
    item.innerHTML = `
      <div class="account-reference-main">
        <div class="account-reference-code mono">${escapeHtml(reference)}</div>
        <div class="account-reference-date">${escapeHtml(timestamp)}</div>
      </div>
      <span class="account-reference-pill ${statusMeta.className}">${escapeHtml(statusMeta.label)}</span>
    `;
    accountReferenceHistoryList.appendChild(item);
  });
}

function renderAccountBillingOverview() {
  const walletBalance = Number(billingOverview?.wallet_balance_eur || 0);
  const pendingTopups = Number(billingOverview?.wallet_pending_topups_eur || 0);
  if (accountWalletBalance) {
    accountWalletBalance.textContent = formatMoney(walletBalance);
  }
  if (accountPendingTopups) {
    accountPendingTopups.textContent = formatMoney(pendingTopups);
  }
  renderTopupReferenceHistory();
}

function renderCheckoutStepMode() {
  const { invoiceEnabled, cardEnabled, walletEnabled } = getBillingFlags();
  const walletBalance = Number(billingOverview?.wallet_balance_eur || 0);
  const pendingTopups = Number(billingOverview?.wallet_pending_topups_eur || 0);
  const { total } = getOrderTotals();
  const walletSufficient = walletBalance + 0.0001 >= total;

  if (walletAvailableInline) {
    walletAvailableInline.textContent = formatMoney(walletBalance);
  }
  if (walletBalanceValue) {
    walletBalanceValue.textContent = formatMoney(walletBalance);
  }
  if (walletPendingValue) {
    walletPendingValue.textContent = formatMoney(pendingTopups);
  }
  if (paymentMethodCard) {
    paymentMethodCard.classList.toggle("is-disabled", !cardEnabled);
  }
  if (paymentMethodWallet) {
    const walletInsufficient = !walletSufficient;
    paymentMethodWallet.classList.toggle("is-disabled", !walletEnabled);
    paymentMethodWallet.classList.toggle("is-insufficient", walletInsufficient);
  }

  if (invoiceEnabled) {
    if (step3Title) step3Title.textContent = tr("Invoice Review");
    if (step3Sub) step3Sub.textContent = tr("Billing profile and approval");
    if (invoiceReviewWrap) invoiceReviewWrap.classList.remove("is-hidden");
    if (directPaymentWrap) directPaymentWrap.classList.add("is-hidden");
    setInlineFormErrorToast(paymentMethodError, "");
    if (checkoutStepTitle) checkoutStepTitle.textContent = tr("Confirm invoice information");
    if (checkoutStepSubtitle) {
      checkoutStepSubtitle.textContent = tr(
        "No charge now. These details are used for your month-end invoice."
      );
    }
    if (payButtonLabel) {
      payButtonLabel.textContent = tr("Confirm & Generate");
    }
    setCheckoutPaymentMethod("invoice", { quiet: true });
  } else {
    if (step3Title) step3Title.textContent = tr("Payment");
    if (step3Sub) step3Sub.textContent = tr("Wallet or card payment");
    if (invoiceReviewWrap) invoiceReviewWrap.classList.add("is-hidden");
    if (directPaymentWrap) directPaymentWrap.classList.remove("is-hidden");
    if (checkoutStepTitle) checkoutStepTitle.textContent = tr("Choose payment method");
    if (checkoutStepSubtitle) {
      checkoutStepSubtitle.textContent = tr(
        "Select an instant payment method to generate your labels now."
      );
    }
    if (payButtonLabel) {
      payButtonLabel.textContent = tr("Pay & Generate");
    }
    const currentMethod =
      checkoutPaymentMethod === "wallet" || checkoutPaymentMethod === "card"
        ? checkoutPaymentMethod
        : "";
    const nextMethod =
      currentMethod === "wallet" && (!walletEnabled || !walletSufficient) && cardEnabled
        ? "card"
        : currentMethod === "card" && !cardEnabled && walletEnabled
          ? "wallet"
          : currentMethod || (walletEnabled && walletSufficient ? "wallet" : cardEnabled ? "card" : "wallet");
    setCheckoutPaymentMethod(nextMethod, { quiet: true });
  }

  renderAccountBillingOverview();
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
  const text = String(message || "").trim();
  if (customsError) {
    customsError.hidden = true;
    customsError.textContent = "";
    customsError.classList.remove("is-visible");
  }
  if (text) {
    showStatusToast(text, { tone: "error" });
  }
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
      tr("Complete this declaration for destinations that require customs processing.");
    return;
  }
  if (outsideEu.length === 1) {
    customsScopeMeta.textContent = tr(
      "{country} is outside the customs union. Complete declaration details before service selection.",
      { country: outsideEu[0] }
    );
    return;
  }
  customsScopeMeta.textContent = tr(
    "{count} destinations are outside the customs union. Complete declaration details before service selection.",
    { count: outsideEu.length }
  );
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
    title.textContent = tr("Item {index}", { index: index + 1 });
    head.appendChild(title);

    if (state.customs.items.length > 1) {
      const remove = document.createElement("button");
      remove.type = "button";
      remove.className = "btn btn-ghost btn-sm customs-item-remove";
      remove.dataset.customsItemAction = "remove";
      remove.dataset.customsItemIndex = String(index);
      remove.innerHTML =
        `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><line x1="5" y1="12" x2="19" y2="12"/></svg><span>${tr(
          "Remove"
        )}</span>`;
      head.appendChild(remove);
    }

    const note = document.createElement("div");
    note.className = "customs-note";
    note.innerHTML =
      `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M12 9v4"/><circle cx="12" cy="17" r="1" fill="currentColor" stroke="none"/><path d="M10.3 3.9L1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/></svg><span>${tr(
        "Provide a precise product description and material/composition."
      )}</span>`;

    const descriptionField = document.createElement("label");
    descriptionField.className = "field";
    descriptionField.innerHTML = `<span>${tr("Item Description")}</span><input type="text" data-customs-item-field="description" value="${escapeHtml(
      item.description
    )}" />`;

    const rowThree = document.createElement("div");
    rowThree.className = "field-row customs-row-three";

    const quantityField = document.createElement("label");
    quantityField.className = "field";
    quantityField.innerHTML = `<span>${tr("Number of Items")}</span><input type="number" min="1" step="1" data-customs-item-field="quantity" value="${escapeHtml(
      item.quantity
    )}" />`;

    const weightField = document.createElement("label");
    weightField.className = "field";
    weightField.innerHTML = `<span>${tr("Weight (kg)")}</span><input type="number" min="0.01" step="0.01" data-customs-item-field="weightKg" value="${escapeHtml(
      item.weightKg
    )}" />`;

    const valueField = document.createElement("label");
    valueField.className = "field";
    valueField.innerHTML = `<span>${tr("Value (EUR)")}</span><div class="customs-currency"><span class="customs-currency-prefix">€</span><input type="number" min="0" step="0.01" data-customs-item-field="valueEur" value="${escapeHtml(
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
      `<span>${tr("Origin of Items")}</span><div class="country-input-wrap customs-origin-wrap"><span class="customs-origin-flag"></span><input type="text" data-customs-item-field="originCountry" /></div>`;
    const originInput = originField.querySelector("input");
    const originFlag = originField.querySelector(".customs-origin-flag");
    if (originInput) originInput.value = item.originCountry;
    if (originFlag) originFlag.innerHTML = getCountryIcon(item.originCountry);

    const hsField = document.createElement("label");
    hsField.className = "field";
    hsField.innerHTML = `<span>${tr("HS Tariff Code (optional)")}</span><input type="text" data-customs-item-field="hsCode" value="${escapeHtml(
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
    setCustomsError(tr("Complete required customs declaration fields before continuing."));
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
    void loadBillingOverview({ quiet: true });
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
    setInlineFormErrorToast(errorEl, defaultLabelErrorMessage, "error");
    firstInvalid.focus();
    return false;
  }

  setInlineFormErrorToast(errorEl, "");
  return true;
}

function validateLabelInfo() {
  if (state.csvMode) {
    const ok = validateCsvRows();
    if (!ok) {
      setInlineFormErrorToast(labelError, defaultLabelErrorMessage, "error");
    } else {
      setInlineFormErrorToast(labelError, "");
    }
    return ok;
  }
  const selectedOrigin = ensureSelectedWarehouseOrigin();
  if (!selectedOrigin) {
    setInlineFormErrorToast(
      labelError,
      tr("Add at least one shipping origin in Account settings before continuing."),
      "error"
    );
    return false;
  }
  applyWarehouseToSender(selectedOrigin, { announce: false });
  return validateFields(labelRequiredFields, inputMap, labelError, isLabelFieldValid);
}

if (!(typeof window !== "undefined" && window.__SHIPIDE_INVOICE_PRINT_MODE__)) {
Object.entries(inputMap).forEach(([_, input]) => {
  bindCompositionAwareInput(input, () => {
    syncInfoState();
    input.classList.remove("is-invalid");
    setInlineFormErrorToast(labelError, "");
  });
});

if (authForm) {
  authForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    if (authMode === "register") {
      if (authRegisterStep === 1) {
        await prepareRegistrationAgreementPreview({ animate: true });
        return;
      }
      await registerWithInvite();
      return;
    }
    await signInWithPassword();
  });
}

[
  authEmail,
  authPassword,
  authPasswordConfirm,
  authCompanyName,
  authContactName,
  authContactEmail,
  authContactPhone,
  authBillingStreet,
  authBillingCity,
  authBillingPostalCode,
  authBillingCountry,
  authBillingAddress,
  authTaxId,
  authCustomerId,
].forEach((input) => {
  if (!input) return;
  bindCompositionAwareInput(input, () => {
    if (input === authBillingCountry) {
      updateAuthBillingCountryFlag(input.value);
    }
    if (authMode === "register") {
      setAuthMessage("");
    }
    updateAuthRegisterSubmitState();
  });
  if (input === authPassword || input === authPasswordConfirm) {
    input.addEventListener("blur", () => {
      maybeShowRegistrationPasswordToast(input);
    });
  }
});

if (authRegisterBack) {
  authRegisterBack.addEventListener("click", () => {
    if (authMode !== "register") return;
    setAuthRegisterStep(1, { resetMessage: true, animate: true });
  });
}

if (authAgreementScroll) {
  authAgreementScroll.addEventListener("scroll", () => {
    hideAuthAgreementMagnifier();
    updateAuthAgreementProgress();
  });
}

if (authAgreementPages) {
  authAgreementPages.addEventListener("pointermove", (event) => {
    updateAuthAgreementMagnifier(event);
  });
  authAgreementPages.addEventListener("pointerleave", () => {
    hideAuthAgreementMagnifier();
  });
}

if (authAgreementAccept) {
  authAgreementAccept.addEventListener("change", () => {
    authAgreementAccepted = authAgreementAccept.checked;
    if (authAgreementAccepted) {
      authAgreementAgreedAt = new Date().toISOString();
      setAuthAgreementStatus(tr("Agreement accepted. You can create your account."), {
        tone: "success",
      });
    } else {
      authAgreementAgreedAt = "";
      setAuthAgreementStatus("");
    }
    updateAuthRegisterSubmitState();
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
    await Promise.all([
      loadWarehouseSettings({ quiet: true }),
      loadBillingOverview({ quiet: true }),
    ]);
  });
}

if (openBuilderPageButton) {
  openBuilderPageButton.addEventListener("click", () => {
    setMainView("builder");
    updateRoute({ view: "builder", step: state.step }, { replace: false });
  });
}

if (openAdminPageButton) {
  openAdminPageButton.addEventListener("click", async () => {
    const hasAccess = await loadAdminAccessStatus({ quiet: true });
    if (!hasAccess) {
      showToast(tr("You are not allowed to access the admin panel."), { tone: "error" });
      return;
    }
    setAdminPageVisible(true);
    await loadAdminDashboard({ quiet: true });
  });
}

if (openLeadsPageButton) {
  openLeadsPageButton.addEventListener("click", async () => {
    const hasAccess = await loadAdminAccessStatus({ quiet: true });
    if (!hasAccess) {
      showToast(tr("You are not allowed to access the leads dashboard."), { tone: "error" });
      return;
    }
    setLeadsPageVisible(true);
    await loadLeadProspects({ quiet: true });
  });
}

if (openHistoryPageButton) {
  openHistoryPageButton.addEventListener("click", async () => {
    setHistoryPageVisible(true);
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

if (closeAdminPageButton) {
  closeAdminPageButton.addEventListener("click", () => {
    setAdminPageVisible(false);
  });
}

if (closeLeadsPageButton) {
  closeLeadsPageButton.addEventListener("click", () => {
    setLeadsPageVisible(false);
  });
}

if (closeReportsPageButton) {
  closeReportsPageButton.addEventListener("click", () => {
    setReportsPageVisible(false);
  });
}

if (closeHistoryPageButton) {
  closeHistoryPageButton.addEventListener("click", () => {
    setHistoryPageVisible(false);
  });
}

if (languageSelect) {
  languageSelect.addEventListener("change", async (event) => {
    const target = event.target;
    if (!(target instanceof HTMLSelectElement)) return;
    await setLanguage(target.value, { persist: true });
  });
}

if (clientInviteCreateButton) {
  clientInviteCreateButton.addEventListener("click", async () => {
    await createClientInvite();
  });
}

if (clientInviteCopyButton) {
  clientInviteCopyButton.addEventListener("click", async () => {
    await copyClientInviteUrl();
  });
}

if (openClientInviteHistoryModalButton) {
  openClientInviteHistoryModalButton.addEventListener("click", () => {
    setClientInviteHistoryModalOpen(true);
    void loadClientInviteHistory({ quiet: clientInviteHistory.length > 0 });
  });
}

if (clientInviteHistoryClose) {
  clientInviteHistoryClose.addEventListener("click", () => {
    setClientInviteHistoryModalOpen(false);
  });
}

if (clientInviteHistoryCancel) {
  clientInviteHistoryCancel.addEventListener("click", () => {
    setClientInviteHistoryModalOpen(false);
  });
}

if (clientInviteHistoryModal) {
  clientInviteHistoryModal.addEventListener("click", (event) => {
    if (event.target === clientInviteHistoryModal) {
      setClientInviteHistoryModalOpen(false);
    }
  });
}

if (leadsSearchInput) {
  bindCompositionAwareInput(leadsSearchInput, (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) return;
    leadSearchQuery = target.value || "";
    renderLeadProspects();
  });
}

if (leadsStackFilter) {
  leadsStackFilter.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLSelectElement)) return;
    leadStackFilterValue = String(target.value || "all");
    renderLeadProspects();
  });
}

if (reloadLeadsButton) {
  reloadLeadsButton.addEventListener("click", async () => {
    await loadLeadProspects({ quiet: true, force: true });
  });
}

if (leadsTableBody) {
  leadsTableBody.addEventListener("click", (event) => {
    const target = event.target instanceof Element ? event.target.closest("[data-lead-call]") : null;
    if (!(target instanceof HTMLButtonElement)) return;
    handleLeadCallAction(target.dataset.leadCall);
  });
}

if (leadCallOutcomeActions) {
  leadCallOutcomeActions.addEventListener("click", async (event) => {
    const target =
      event.target instanceof Element ? event.target.closest("[data-lead-outcome]") : null;
    if (!(target instanceof HTMLButtonElement)) return;
    await saveLeadCallOutcome(target.dataset.leadOutcome);
  });
}

if (leadCallOutcomeClose) {
  leadCallOutcomeClose.addEventListener("click", () => {
    closeLeadCallOutcome();
  });
}

if (leadCallOutcomeCancel) {
  leadCallOutcomeCancel.addEventListener("click", () => {
    closeLeadCallOutcome();
  });
}

if (leadCallOutcomeModal) {
  leadCallOutcomeModal.addEventListener("click", (event) => {
    if (event.target === leadCallOutcomeModal) {
      closeLeadCallOutcome();
    }
  });
}

if (clientInviteEmailInput) {
  clientInviteEmailInput.addEventListener("input", () => {
    setClientInviteStatus("");
  });
}

if (clientInviteExpirySelect) {
  clientInviteExpirySelect.addEventListener("change", () => {
    setClientInviteStatus("");
  });
}

if (clientInviteHistoryList) {
  clientInviteHistoryList.addEventListener("click", async (event) => {
    const target = event.target instanceof Element ? event.target : null;
    if (!target) return;
    const copyButton = target.closest("[data-invite-copy]");
    if (copyButton instanceof HTMLElement) {
      await copyInviteHistoryUrl(copyButton.dataset.inviteCopy);
      return;
    }
    const revokeButton = target.closest("[data-invite-revoke]");
    if (revokeButton instanceof HTMLElement) {
      await revokeClientInvite(revokeButton.dataset.inviteRevoke);
    }
  });
}

if (adminCarrierDiscountInput) {
  adminCarrierDiscountInput.addEventListener("input", () => {
    renderAdminSettingsPreview();
  });
}

if (adminClientDiscountInput) {
  adminClientDiscountInput.addEventListener("input", () => {
    renderAdminSettingsPreview();
  });
}

if (adminSettingsSaveButton) {
  adminSettingsSaveButton.addEventListener("click", async () => {
    await saveAdminSettings();
  });
}

if (openAdminSettingsModalButton) {
  openAdminSettingsModalButton.addEventListener("click", () => {
    renderAdminSettingsPreview();
    setAdminSettingsModalOpen(true);
  });
}

if (adminSettingsClose) {
  adminSettingsClose.addEventListener("click", () => {
    setAdminSettingsModalOpen(false);
  });
}

if (adminSettingsCancel) {
  adminSettingsCancel.addEventListener("click", () => {
    setAdminSettingsModalOpen(false);
  });
}

if (adminSettingsModal) {
  adminSettingsModal.addEventListener("click", (event) => {
    if (event.target === adminSettingsModal) {
      setAdminSettingsModalOpen(false);
    }
  });
}

if (adminClientSearchInput) {
  bindCompositionAwareInput(adminClientSearchInput, (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) return;
    adminClientSearch = target.value || "";
    renderAdminClientsList();
  });
}

if (adminClientFilterSelect) {
  adminClientFilterSelect.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLSelectElement)) return;
    adminClientFilter = String(target.value || "all");
    renderAdminClientsList();
  });
}

if (adminClientSortSelect) {
  adminClientSortSelect.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLSelectElement)) return;
    adminClientSort = String(target.value || "recent");
    renderAdminClientsList();
  });
}

if (adminMockDataButton) {
  adminMockDataButton.addEventListener("click", () => {
    toggleAdminMockData();
  });
}

if (adminBillingRunPreviewButton) {
  adminBillingRunPreviewButton.addEventListener("click", async () => {
    await runAdminBillingCycle("preview");
  });
}

if (adminBillingRunCreateButton) {
  adminBillingRunCreateButton.addEventListener("click", async () => {
    await runAdminBillingCycle("create");
  });
}

if (adminBillingRunSendButton) {
  adminBillingRunSendButton.addEventListener("click", async () => {
    await runAdminBillingCycle("send");
  });
}

if (adminBillingSendTestButton) {
  adminBillingSendTestButton.addEventListener("click", async () => {
    await sendAdminBillingTestEmail();
  });
}

if (adminBillingSendTestSequenceButton) {
  adminBillingSendTestSequenceButton.addEventListener("click", async () => {
    await sendAdminBillingTestSequence();
  });
}

if (adminReportsSendTestButton) {
  adminReportsSendTestButton.addEventListener("click", async () => {
    await sendAdminReportsTestEmail();
  });
}

if (adminAgreementPreviewButton) {
  adminAgreementPreviewButton.addEventListener("click", async () => {
    await previewAdminAgreementEmail();
  });
}

if (adminAgreementSendTestButton) {
  adminAgreementSendTestButton.addEventListener("click", async () => {
    await sendAdminAgreementTestEmail();
  });
}

if (adminInvoiceList) {
  adminInvoiceList.addEventListener("click", async (event) => {
    const target = event.target instanceof Element ? event.target : null;
    if (!target) return;
    const sendButton = target.closest("[data-admin-invoice-send]");
    if (sendButton instanceof HTMLElement) {
      await sendAdminInvoice(sendButton.dataset.adminInvoiceSend);
      return;
    }
    const paidButton = target.closest("[data-admin-invoice-paid]");
    if (paidButton instanceof HTMLElement) {
      await markAdminInvoicePaid(paidButton.dataset.adminInvoicePaid);
    }
  });
}

if (adminClientsList) {
  adminClientsList.addEventListener("click", async (event) => {
    const target = event.target instanceof Element ? event.target : null;
    if (!target) return;
    const toggle = target.closest("[data-admin-billing-toggle]");
    if (!(toggle instanceof HTMLElement)) return;
    const clientId = String(toggle.dataset.adminClientId || "").trim();
    const method = String(toggle.dataset.adminBillingToggle || "").trim();
    await updateAdminClientBilling(clientId, method);
  });
}

if (warehouseAddButton) {
  warehouseAddButton.addEventListener("click", () => {
    if (!currentUser || warehouseSaving) return;
    if (warehouseRecords.length >= WAREHOUSE_MAX_COUNT) {
      setWarehouseStatus(
        tr("Maximum {max} warehouses reached.", {
          max: WAREHOUSE_MAX_COUNT,
        }),
        {
        tone: "error",
        }
      );
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
  const handleWarehouseFieldInput = (target) => {
    if (!(target instanceof HTMLInputElement)) return;
    const field = target.dataset.warehouseField;
    if (!field) return;
    const card = target.closest("[data-warehouse-id]");
    if (!card) return;
    const index = findWarehouseIndexById(card.dataset.warehouseId);
    if (index === -1) return;
    const value = target.value.trim();
    warehouseRecords[index][field] = value;

    const returnField = WAREHOUSE_RETURN_FIELD_MAP[field];
    if (returnField && warehouseRecords[index].sameReturnAddress) {
      warehouseRecords[index][returnField] = value;
    }

    if (field === "name") {
      const label = card.querySelector(".warehouse-card-name");
      if (label) {
        label.textContent = value || `${tr("Warehouse")} ${index + 1}`;
      }
    }

    if (field === "country" || field === "returnCountry") {
      const icon = target.parentElement?.querySelector(".warehouse-country-flag");
      if (icon) {
        icon.innerHTML = getCountryIcon(value);
      }
    }

    if (
      warehouseRecords[index].id === state.shipFromOriginId &&
      WAREHOUSE_SENDER_FIELDS.has(field)
    ) {
      applyWarehouseToSender(warehouseRecords[index], { announce: false });
      syncCsvRowsWithSelectedOrigin({ rerender: true });
    }

    setWarehouseDirty(true);
    renderSenderOriginSelector();
    renderCsvShipFromSelector();
  };

  bindDelegatedCompositionAwareInput(
    warehouseList,
    "input[data-warehouse-field]",
    handleWarehouseFieldInput
  );

  warehouseList.addEventListener("click", async (event) => {
    const actionButton = event.target.closest("[data-warehouse-action]");
    if (!actionButton) return;
    const card = actionButton.closest("[data-warehouse-id]");
    if (!card) return;
    const warehouseId = card.dataset.warehouseId;
    const action = actionButton.dataset.warehouseAction;
    const index = findWarehouseIndexById(warehouseId);
    if (index === -1) return;

    if (action === "save") {
      setSelectedWarehouseOrigin(warehouseId, { syncCsv: true });
      await saveWarehouseSettings();
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

    if (action === "toggle-return") {
      const origin = warehouseRecords[index];
      const nextSameReturn = !origin.sameReturnAddress;
      const nextOrigin = {
        ...origin,
        sameReturnAddress: nextSameReturn,
      };
      if (nextSameReturn) {
        nextOrigin.returnSenderName = origin.senderName;
        nextOrigin.returnStreet = origin.street;
        nextOrigin.returnCity = origin.city;
        nextOrigin.returnRegion = origin.region;
        nextOrigin.returnPostalCode = origin.postalCode;
        nextOrigin.returnCountry = origin.country;
      } else {
        nextOrigin.returnSenderName = origin.returnSenderName || origin.senderName;
        nextOrigin.returnStreet = origin.returnStreet || origin.street;
        nextOrigin.returnCity = origin.returnCity || origin.city;
        nextOrigin.returnRegion = origin.returnRegion || origin.region;
        nextOrigin.returnPostalCode = origin.returnPostalCode || origin.postalCode;
        nextOrigin.returnCountry = origin.returnCountry || origin.country;
      }
      warehouseRecords[index] = normalizeWarehouseRecord(nextOrigin, origin);
      renderWarehouseList();
      setWarehouseDirty(true);
      return;
    }

    if (action === "remove") {
      if (warehouseRecords.length <= 1) {
        setWarehouseStatus(tr("At least one warehouse origin is required."), {
          tone: "error",
        });
        return;
      }
      removeWarehouseWithAnimation(warehouseId);
    }
  });
}

if (senderOriginSelector) {
  senderOriginSelector.addEventListener("click", (event) => {
    const clickTarget = event.target;
    if (!(clickTarget instanceof Element)) return;
    const target = clickTarget.closest("[data-origin-id]");
    if (!target) return;
    const originId = String(target.dataset.originId || "").trim();
    if (!originId) return;
    if (warehouseRecords.length <= 1) return;
    setSelectedWarehouseOrigin(originId, { syncCsv: true });
  });
}

if (csvShipFromSelector) {
  csvShipFromSelector.addEventListener("click", (event) => {
    const clickTarget = event.target;
    if (!(clickTarget instanceof Element)) return;
    const target = clickTarget.closest("[data-origin-id]");
    if (!target) return;
    if (state.shipFromLockedByProvider) return;
    const originId = String(target.dataset.originId || "").trim();
    if (!originId) return;
    if (warehouseRecords.length <= 1) return;
    setSelectedWarehouseOrigin(originId, { syncCsv: true });
  });
}

if (csvShipFromNote) {
  csvShipFromNote.addEventListener("click", async (event) => {
    const target = event.target;
    if (!(target instanceof Element)) return;
    const trigger = target.closest("[data-action='open-shopify-settings']");
    if (!trigger) return;
    event.preventDefault();
    await openShopifySettingsModal();
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
  const handler = () => {
    syncCustomsStateFromInputs();
    setCustomsFieldInvalid(input, false);
    setCustomsError("");
  };
  if (eventName === "input") {
    bindCompositionAwareInput(input, handler);
  } else {
    input.addEventListener(eventName, handler);
  }
});

if (customsItemsList) {
  const handleCustomsItemInput = (target) => {
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
  };

  bindDelegatedCompositionAwareInput(
    customsItemsList,
    "input[data-customs-item-field]",
    handleCustomsItemInput
  );

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
    const token = String(button.dataset.reportRange || "all");
    if (token === "custom") {
      activateCustomReportsRangePicker();
      return;
    }
    if (token === "all") {
      reportsRangeMode = "all";
      setReportsCustomRangeVisible(false);
      if (reportsStartDate) reportsStartDate.classList.remove("is-invalid");
      if (reportsEndDate) reportsEndDate.classList.remove("is-invalid");
      renderReportsDashboard();
      return;
    }
    const nextRange = Number(token);
    if (!REPORT_RANGE_PRESETS.includes(nextRange)) return;
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
  openReceiptModalButton.addEventListener("click", (event) => {
    event.preventDefault();
    void openReceiptPdfFile();
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

if (paymentMethodList) {
  paymentMethodList.addEventListener("click", (event) => {
    const target = event.target instanceof Element ? event.target.closest("[data-payment-method]") : null;
    if (!(target instanceof HTMLElement)) return;
    if (target.classList.contains("is-disabled")) return;
    const method = String(target.dataset.paymentMethod || "").trim().toLowerCase();
    setCheckoutPaymentMethod(method);
    setInlineFormErrorToast(paymentMethodError, "");
  });

  paymentMethodList.addEventListener("keydown", (event) => {
    if (event.key !== "Enter" && event.key !== " ") return;
    if (isTextEntryElement(event.target)) return;
    const target = event.target instanceof Element ? event.target.closest("[data-payment-method]") : null;
    if (!(target instanceof HTMLElement)) return;
    if (target.classList.contains("is-disabled")) return;
    event.preventDefault();
    const method = String(target.dataset.paymentMethod || "").trim().toLowerCase();
    setCheckoutPaymentMethod(method);
    setInlineFormErrorToast(paymentMethodError, "");
  });
}

if (cardNumberInput) {
  cardNumberInput.addEventListener("input", () => {
    const digits = String(cardNumberInput.value || "").replace(/\D+/g, "");
    const detectedBrand = detectCardBrand(digits);
    const maxLen = getCardNumberMaxLength(detectedBrand);
    const clipped = digits.slice(0, maxLen);
    updateCardBrandVisual(detectedBrand);
    cardNumberInput.value = formatCardNumberForDisplay(clipped, detectedBrand);
  });
}

if (cardExpiryInput) {
  cardExpiryInput.addEventListener("input", () => {
    cardExpiryInput.value = formatCardExpiryForDisplay(cardExpiryInput.value);
  });
}

if (cardCvcInput) {
  cardCvcInput.addEventListener("input", () => {
    const brand = detectCardBrand(String(cardNumberInput?.value || ""));
    const max = getCardCvcMaxLength(brand);
    cardCvcInput.value = String(cardCvcInput.value || "")
      .replace(/\D+/g, "")
      .slice(0, max);
  });
}

updateCardBrandVisual("generic");

if (openIbanTopupFromStep) {
  openIbanTopupFromStep.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.stopPropagation();
    }
  });
  openIbanTopupFromStep.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    setIbanTopupModalOpen(true);
  });
}

if (openIbanTopupFromAccount) {
  openIbanTopupFromAccount.addEventListener("click", (event) => {
    event.preventDefault();
    setIbanTopupModalOpen(true);
  });
}

if (openWalletHistoryFromAccount) {
  openWalletHistoryFromAccount.addEventListener("click", (event) => {
    event.preventDefault();
    renderTopupReferenceHistory();
    setWalletHistoryModalOpen(true);
  });
}

if (ibanTopupClose) {
  ibanTopupClose.addEventListener("click", (event) => {
    event.preventDefault();
    setIbanTopupModalOpen(false);
  });
}

if (ibanTopupCancel) {
  ibanTopupCancel.addEventListener("click", (event) => {
    event.preventDefault();
    setIbanTopupModalOpen(false);
  });
}

if (ibanTopupModal) {
  ibanTopupModal.addEventListener("click", (event) => {
    if (event.target === ibanTopupModal) {
      setIbanTopupModalOpen(false);
    }
  });
}

if (walletHistoryClose) {
  walletHistoryClose.addEventListener("click", (event) => {
    event.preventDefault();
    setWalletHistoryModalOpen(false);
  });
}

if (walletHistoryCancel) {
  walletHistoryCancel.addEventListener("click", (event) => {
    event.preventDefault();
    setWalletHistoryModalOpen(false);
  });
}

if (walletHistoryModal) {
  walletHistoryModal.addEventListener("click", (event) => {
    if (event.target === walletHistoryModal) {
      setWalletHistoryModalOpen(false);
    }
  });
}

if (ibanTopupRequest) {
  ibanTopupRequest.addEventListener("click", (event) => {
    event.preventDefault();
    void createIbanTopupRequest();
  });
}

if (ibanCopyReference) {
  ibanCopyReference.addEventListener("click", async (event) => {
    event.preventDefault();
    const value = String(ibanResultReference?.textContent || "").trim();
    if (!value || value === "--") return;
    try {
      await navigator.clipboard.writeText(value);
      showToast(tr("Reference copied to clipboard."), { tone: "success" });
    } catch (_error) {
      showToast(tr("Could not copy reference."), { tone: "error" });
    }
  });
}

if (ibanCopyIban) {
  ibanCopyIban.addEventListener("click", async (event) => {
    event.preventDefault();
    const value = String(ibanResultIban?.textContent || "").trim();
    if (!value || value === "--") return;
    try {
      await navigator.clipboard.writeText(value);
      showToast(tr("IBAN copied to clipboard."), { tone: "success" });
    } catch (_error) {
      showToast(tr("Could not copy IBAN."), { tone: "error" });
    }
  });
}

if (receiptDownloadPdf) {
  receiptDownloadPdf.addEventListener("click", (event) => {
    event.preventDefault();
    downloadReceiptPdfFile();
  });
}

if (accountDownloadInvoice) {
  accountDownloadInvoice.addEventListener("click", (event) => {
    event.preventDefault();
    downloadInvoicePdfFile();
  });
}

document.addEventListener("keydown", (event) => {
  if (event.key !== "ArrowUp" && event.key !== "ArrowDown") return;
  if (isTextEntryElement(event.target)) return;
  if (receiptModal && !receiptModal.classList.contains("is-closed")) return;
  if (restrictedGoodsModal && !restrictedGoodsModal.classList.contains("is-closed")) return;
  if (shopifySettingsModal && !shopifySettingsModal.classList.contains("is-closed")) return;
  if (ibanTopupModal && !ibanTopupModal.classList.contains("is-closed")) return;
  if (walletHistoryModal && !walletHistoryModal.classList.contains("is-closed")) return;
  if (clientInviteHistoryModal && !clientInviteHistoryModal.classList.contains("is-closed")) return;
  if (adminSettingsModal && !adminSettingsModal.classList.contains("is-closed")) return;
  if (leadCallOutcomeModal && !leadCallOutcomeModal.classList.contains("is-closed")) return;

  const direction = event.key === "ArrowDown" ? 1 : -1;
  const isHistoryOpen = historyPageSection && !historyPageSection.classList.contains("is-hidden");
  const isBuilderOpen = builderPage && !builderPage.classList.contains("is-hidden");

  if (isHistoryOpen) {
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

document.addEventListener("keydown", (event) => {
  if (event.key !== "Escape") return;
  if (clientInviteHistoryModal && !clientInviteHistoryModal.classList.contains("is-closed")) {
    setClientInviteHistoryModalOpen(false);
    return;
  }
  if (adminSettingsModal && !adminSettingsModal.classList.contains("is-closed")) {
    setAdminSettingsModalOpen(false);
    return;
  }
  if (leadCallOutcomeModal && !leadCallOutcomeModal.classList.contains("is-closed")) {
    closeLeadCallOutcome();
    return;
  }
  if (ibanTopupModal && !ibanTopupModal.classList.contains("is-closed")) {
    setIbanTopupModalOpen(false);
    return;
  }
  if (walletHistoryModal && !walletHistoryModal.classList.contains("is-closed")) {
    setWalletHistoryModalOpen(false);
    return;
  }
  if (shopifySettingsModal && !shopifySettingsModal.classList.contains("is-closed")) {
    closeShopifySettingsModal();
  }
});

function autoFill() {
  clearBatchState();
  setCsvMode(false);
  setCsvEditMode(false);
  state.csvRows = [];
  state.csvPage = 1;
  state.csvValidationAttempted = false;
  setInlineFormErrorToast(labelError, "");
  const profile = sampleProfiles[Math.floor(Math.random() * sampleProfiles.length)];
  maybeApplyDefaultWarehouseToSender();

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
    if (!currentUser || !adminAccessAllowed) {
      showToast(tr("Only admin accounts can use Auto-fill."), { tone: "error" });
      return;
    }
    autoFill();
  });
}

if (autoCsvButton) {
  autoCsvButton.addEventListener("click", (event) => {
    event.preventDefault();
    if (!currentUser || !adminAccessAllowed) {
      showToast(tr("Only admin accounts can use Auto CSV."), { tone: "error" });
      return;
    }
    clearBatchState();
    state.csvRows = generateCsvRows(3);
    state.csvPage = 1;
    state.csvValidationAttempted = false;
    setCsvBatchSource("csv-auto");
    setCsvMode(true);
    syncCsvRowsWithSelectedOrigin({ rerender: false });
    setCsvEditMode(false);
    setInlineFormErrorToast(labelError, "");
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

if (csvPagePrev) {
  csvPagePrev.addEventListener("click", (event) => {
    event.preventDefault();
    setCsvPage(state.csvPage - 1);
  });
}

if (csvPageNext) {
  csvPageNext.addEventListener("click", (event) => {
    event.preventDefault();
    setCsvPage(state.csvPage + 1);
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
    void handleCheckoutAndGenerate();
  });
}

async function handleCheckoutAndGenerate() {
  if (!payButton) return;
  const { invoiceEnabled, cardEnabled, walletEnabled } = getBillingFlags();
  const { total } = getOrderTotals();
  let cardMeta = null;

  if (!invoiceEnabled) {
    let method = checkoutPaymentMethod;
    if (method !== "card" && method !== "wallet") {
      method = walletEnabled ? "wallet" : cardEnabled ? "card" : "";
    }
    if (!method) {
      setInlineFormErrorToast(paymentMethodError, defaultPaymentMethodErrorMessage, "error");
      return;
    }
    if (method === "wallet") {
      const walletBalance = Number(billingOverview?.wallet_balance_eur || 0);
      if (walletBalance + 0.0001 < total) {
        setInlineFormErrorToast(
          paymentMethodError,
          tr("Insufficient wallet balance. Top up by IBAN or use card."),
          "error"
        );
        return;
      }
    }
    if (method === "card") {
      const cardValidation = validateCardDetails();
      if (!cardValidation.ok) {
        setInlineFormErrorToast(paymentMethodError, cardValidation.message, "error");
        return;
      }
      cardMeta = {
        brand: cardValidation.brand,
        last4: cardValidation.last4,
      };
    }
    setCheckoutPaymentMethod(method, { quiet: true });
  }

  payButton.disabled = true;
  const originalMarkup = payButton.innerHTML;
  payButton.innerHTML = tr("Preparing Label...");

  try {
    if (!invoiceEnabled) {
      const payload = await fetchApiWithAuth("/api/billing/checkout", {
        method: "POST",
        body: JSON.stringify({
          method: checkoutPaymentMethod,
          amount: total,
          labelsCount: getQuantity(),
          service: state.selection.type,
          cardBrand: cardMeta?.brand || null,
          cardLast4: cardMeta?.last4 || null,
        }),
      });
      await loadBillingOverview({ quiet: true });
      if (checkoutPaymentMethod === "wallet") {
        showToast(
          tr("Paid with account balance. Remaining balance: {amount}", {
            amount: formatMoney(Number(payload?.wallet_balance_eur || 0)),
          }),
          { tone: "success" }
        );
      } else if (checkoutPaymentMethod === "card") {
        showToast(tr("Card payment approved. Generating labels."), { tone: "success" });
      }
    }

    await new Promise((resolve) => window.setTimeout(resolve, 480));
    goToStep(4);
  } catch (error) {
    const message = error?.message || tr("Could not process checkout.");
    if (!invoiceEnabled) {
      setInlineFormErrorToast(paymentMethodError, message, "error");
    } else {
      showToast(message, { tone: "error" });
    }
  } finally {
    payButton.disabled = false;
    payButton.innerHTML = originalMarkup;
  }
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
  state.csvSource = "none";
  state.shipFromLockedByProvider = false;
  state.csvPage = 1;
  state.csvValidationAttempted = false;
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
  renderCsvShipFromSelector();
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
  checkoutPaymentMethod = "invoice";
  state.quantity = 1;
  state.csvMode = false;
  state.csvEditable = false;
  state.csvRows = [];
  state.csvPage = 1;
  state.csvValidationAttempted = false;
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
    if (label) label.textContent = tr("Edit");
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
  setInlineFormErrorToast(paymentMethodError, "");
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
    if (opt.disabled || opt.classList.contains("is-disabled")) return;
    const provider = opt.dataset.provider;
    const optionLabel =
      opt.querySelector(".provider-option-main > span:last-child")?.textContent?.trim() ||
      opt.textContent.trim();
    providerDropdown.classList.remove("is-open");
    if (provider === "shopify") {
      await handleShopifyProviderAction();
      return;
    }
    if (provider === "woocommerce" || provider === "wix") {
      openGenericProviderSettings(optionLabel);
      return;
    }

    autoFill();
    const triggerText = providerTrigger?.querySelector("span");
    if (triggerText) {
      triggerText.textContent = tr("Imported from {provider}", { provider: optionLabel });
      window.setTimeout(() => {
        triggerText.textContent = tr("Import from provider");
      }, 2000);
    }
  });
});

[
  [woocommerceSettingsTrigger, tr("WooCommerce")],
  [wixSettingsTrigger, tr("Wix eCommerce")],
].forEach(([trigger, providerLabel]) => {
  if (!trigger) return;
  trigger.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (providerDropdown) {
      providerDropdown.classList.remove("is-open");
    }
    openGenericProviderSettings(providerLabel);
  });
});

if (shopifySettingsTrigger) {
  shopifySettingsTrigger.addEventListener("click", async (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (providerDropdown) {
      providerDropdown.classList.remove("is-open");
    }
    await openShopifySettingsModal();
  });
}

if (providerSettingsClose) {
  providerSettingsClose.addEventListener("click", (event) => {
    event.preventDefault();
    closeGenericProviderSettingsModal();
  });
}

if (providerSettingsCancel) {
  providerSettingsCancel.addEventListener("click", (event) => {
    event.preventDefault();
    closeGenericProviderSettingsModal();
  });
}

if (providerSettingsModal) {
  providerSettingsModal.addEventListener("click", (event) => {
    if (event.target === providerSettingsModal) {
      closeGenericProviderSettingsModal();
    }
  });
}

if (shopifyLocationsList) {
  shopifyLocationsList.addEventListener("click", (event) => {
    const row = event.target?.closest?.(".shopify-location-row");
    if (!row || !shopifyLocationsList.contains(row)) return;
    if (row.dataset.role === "all") {
      const shouldSelectAll = !row.classList.contains("is-selected");
      if (shouldSelectAll) {
        shopifyLocationDraftSelection = new Set(
          shopifyLocationsCache.map((location) => location.id)
        );
      } else {
        shopifyLocationDraftSelection = new Set();
      }
      renderShopifySettingsLocations();
      return;
    }
    const locationId = normalizeShopifyLocationId(row.dataset.locationId);
    if (!locationId) return;
    if (shopifyLocationDraftSelection.has(locationId)) {
      shopifyLocationDraftSelection.delete(locationId);
    } else {
      shopifyLocationDraftSelection.add(locationId);
    }
    renderShopifySettingsLocations();
  });
}

if (shopifySettingsSave) {
  shopifySettingsSave.addEventListener("click", async (event) => {
    event.preventDefault();
    await saveShopifySettings();
  });
}

if (shopifySettingsClose) {
  shopifySettingsClose.addEventListener("click", (event) => {
    event.preventDefault();
    closeShopifySettingsModal();
  });
}

if (shopifySettingsCancel) {
  shopifySettingsCancel.addEventListener("click", (event) => {
    event.preventDefault();
    closeShopifySettingsModal();
  });
}

if (shopifySettingsModal) {
  shopifySettingsModal.addEventListener("click", (event) => {
    if (event.target === shopifySettingsModal) {
      closeShopifySettingsModal();
    }
  });
}

// ── CSV upload modal ──

function setCsvMapError(message = "") {
  const text = String(message || "").trim();
  if (csvMapError) {
    csvMapError.hidden = true;
    csvMapError.textContent = "";
    csvMapError.classList.remove("is-visible");
  }
  if (text) {
    showStatusToast(text, { tone: "error" });
  }
}

function setCsvModalStep(step, options = {}) {
  const { animate = true } = options;
  const nextStep = step === "mapping" ? "mapping" : "upload";
  const isMapping = nextStep === "mapping";
  const shouldAnimate = animate && !prefersReducedMotion();
  const modalTitle = document.querySelector(".csv-modal-title");

  if (!csvUploadStep || !csvMappingStep) {
    if (modalTitle) {
      modalTitle.textContent = isMapping ? tr("Map CSV columns") : tr("Upload CSV file");
    }
    csvModalStepState = nextStep;
    return;
  }

  if (!shouldAnimate) {
    csvUploadStep.classList.toggle("is-active", !isMapping);
    csvMappingStep.classList.toggle("is-active", isMapping);
    if (modalTitle) {
      modalTitle.textContent = isMapping ? tr("Map CSV columns") : tr("Upload CSV file");
    }
    csvModalStepState = nextStep;
    return;
  }

  const currentStep = csvModalStepState;
  if (currentStep === nextStep) {
    if (modalTitle) {
      modalTitle.textContent = isMapping ? tr("Map CSV columns") : tr("Upload CSV file");
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
      modalTitle.textContent = isMapping ? tr("Map CSV columns") : tr("Upload CSV file");
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
    csvMapMeta.textContent = tr("Auto-mapped columns ready for review");
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
  csvDropzone.addEventListener("click", (e) => {
    if (e.target === csvFileInput) return;
    if (csvFileInput) csvFileInput.click();
  });

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
  if (!file) return;
  if (!/\.(csv|tsv|txt)$/i.test(file.name || "")) {
    showToast(tr("Please upload a .csv file."), { tone: "error" });
    return;
  }
  const reader = new FileReader();
  reader.onerror = () => {
    setCsvModalStep("mapping");
    setCsvMapError(tr("Could not read the file. Please try again."));
    if (csvMappingBody) csvMappingBody.innerHTML = "";
    if (csvMapMeta) csvMapMeta.textContent = tr("No rows detected");
  };
  reader.onload = (e) => {
    try {
      const text = String(e.target.result || "");
      const analysis = analyzeCsvText(text);
      if (!analysis) {
        setCsvModalStep("mapping");
        setCsvMapError(tr("Could not parse this CSV. Please check the file format and try again."));
        if (csvMappingBody) {
          csvMappingBody.innerHTML = "";
        }
        if (csvMapMeta) {
          csvMapMeta.textContent = tr("No rows detected");
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
    } catch (err) {
      setCsvModalStep("mapping");
      setCsvMapError(tr("Error processing CSV: {msg}", { msg: err.message || "Unknown error" }));
      if (csvMappingBody) csvMappingBody.innerHTML = "";
      if (csvMapMeta) csvMapMeta.textContent = tr("Error");
    }
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
  const selectedOrigin = getSelectedWarehouseRecord();
  if (!selectedOrigin) {
    return {
      senderName: "",
      senderStreet: "",
      senderCity: "",
      senderState: "",
      senderZip: "",
    };
  }
  return getWarehouseSenderValues(selectedOrigin);
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
    csvMapMeta.textContent = tr("{file} • {rows} rows • {columns} columns", {
      file: fileName,
      rows: analysis.rows.length,
      columns: analysis.headers.length,
    });
  }

  const orderedColumns = getCsvColumnsRequiredFirst();
  orderedColumns.forEach((column) => {
    const rowEl = document.createElement("tr");

    const fieldCell = document.createElement("td");
    const fieldWrap = document.createElement("div");
    fieldWrap.className = "csv-map-field";
    const fieldLabel = document.createElement("strong");
    fieldLabel.textContent = column.label;
    fieldWrap.appendChild(fieldLabel);
    if (CSV_REQUIRED_FIELDS.has(column.key)) {
      const badge = document.createElement("span");
      badge.className = "csv-map-required";
      badge.textContent = tr("Required");
      fieldWrap.appendChild(badge);
    }
    fieldCell.appendChild(fieldWrap);

    const selectCell = document.createElement("td");
    const select = document.createElement("select");
    select.className = "csv-map-select";
    select.dataset.mapKey = column.key;

    const noneOption = document.createElement("option");
    noneOption.value = "-1";
    noneOption.textContent = tr("Not mapped");
    select.appendChild(noneOption);

    analysis.headers.forEach((header, index) => {
      const option = document.createElement("option");
      option.value = String(index);
      option.textContent = header || tr("Column {index}", { index: index + 1 });
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

    rowEl.appendChild(fieldCell);
    rowEl.appendChild(selectCell);
    rowEl.appendChild(sampleCell);
    csvMappingBody.appendChild(rowEl);
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
    sampleEl.textContent = tr("No sample");
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
    setCsvMapError(
      tr("Map required fields: {fields}.", { fields: missingRequired.join(", ") })
    );
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
      tr("Required fields must map to separate CSV columns ({fields}).", {
        fields: conflictLabels.join(" + "),
      })
    );
    return;
  }

  const rows = buildCsvRowsFromAnalysis(csvMappingDraft.analysis, csvMappingDraft.mapping);
  if (!rows.length) {
    setCsvMapError(tr("This mapping produced no usable rows. Adjust mapping and try again."));
    return;
  }

  clearBatchState();
  state.csvRows = rows;
  state.csvPage = 1;
  state.csvValidationAttempted = false;
  setCsvBatchSource("csv-upload");
  setCsvMode(true);
  syncCsvRowsWithSelectedOrigin({ rerender: false });
  setCsvEditMode(false);
  setInlineFormErrorToast(labelError, "");
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

if (!(typeof window !== "undefined" && window.__SHIPIDE_INVOICE_PRINT_MODE__)) {
  ensureIdentifiers();
  state.quantity = Number(quantityInput.value) || 1;
  updateCountryFlag();
  updateSummary();
  updatePayment();
  updatePreview();
  renderLeadProspects();
  const initialRoute = parseRouteFromLocation();
  const initialStep =
    initialRoute.view === "builder" ? clampStep(initialRoute.step) : clampStep(state.step);

  goToStep(initialStep, { push: false, regenerate: false });
  if (initialRoute.view === "register") {
    history.replaceState(
      routeToState(initialRoute),
      "",
      `${window.location.pathname}${window.location.search || ""}`
    );
  } else {
    updateRoute(
      initialRoute.view === "builder" ? { view: "builder", step: initialStep } : initialRoute,
      { replace: true }
    );
  }

  window.addEventListener("popstate", (event) => {
    const route =
      event.state && typeof event.state.view === "string"
        ? event.state
        : parseRouteFromLocation();

    if (!currentUser) {
      setMainView("builder", { push: false });
      if (route.view === "register") {
        setAuthMode("register", { inviteToken: route.inviteToken });
        void loadRegistrationInvite(route.inviteToken, {
          preview: Boolean(route.previewRegistration),
          previewStep: route.previewStep,
        });
      } else {
        setAuthMode("login");
        updateRoute({ view: "login" }, { replace: true });
      }
      return;
    }

    if (route.view === "account") {
      setMainView("account", { push: false });
      loadWarehouseSettings({ quiet: true });
      loadBillingOverview({ quiet: true });
      return;
    }

    if (route.view === "admin") {
      setMainView("admin", { push: false });
      loadAdminDashboard({ quiet: true });
      return;
    }

    if (route.view === "leads") {
      setMainView("leads", { push: false });
      loadLeadProspects({ quiet: true });
      return;
    }

    if (route.view === "history") {
      setMainView("history", { push: false });
      loadGenerationHistory({ preferLatest: true });
      return;
    }

    if (route.view === "reports") {
      setMainView("reports", {
        push: false,
        reportRange: String(route?.reportRange || getReportRangeFromLocation(window.location) || ""),
      });
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

  window.addEventListener("focus", () => {
    if (!currentUser) return;
    void refreshAuthAccessToken();
  });

  if (batchPreview) {
    batchPreview.classList.add("is-single");
  }
  renderCsvTable();

  renderAdminInvoiceList();
  setAdminBillingBusy(false);
  resetWarehouseState();
  initializeAuthLogo();
  initializeAppLogo();
  initializeAuthBackground();
  void setLanguage(resolvePreferredLanguage(null), { persist: false });
  initializeAuth();
}
}
