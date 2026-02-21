// Algorithmic flag emoji: converts any ISO 3166-1 alpha-2 code to flag emoji
function isoToFlag(code) {
  if (!code || code.length !== 2) return "\u{1F30D}";
  const upper = code.toUpperCase();
  const cp1 = 0x1F1E6 + (upper.charCodeAt(0) - 65);
  const cp2 = 0x1F1E6 + (upper.charCodeAt(1) - 65);
  return String.fromCodePoint(cp1, cp2);
}

// Country name / alias -> ISO 3166-1 alpha-2 code
// Built in segments to stay under content limits
const COUNTRY_CODES = {};

function _add(entries) {
  entries.forEach(function (e) {
    COUNTRY_CODES[e[0]] = e[1];
  });
}

_add([
  ["afghanistan", "AF"], ["albania", "AL"], ["algeria", "DZ"],
  ["andorra", "AD"], ["angola", "AO"], ["antigua and barbuda", "AG"],
  ["argentina", "AR"], ["armenia", "AM"], ["australia", "AU"],
  ["austria", "AT"], ["azerbaijan", "AZ"],
]);

_add([
  ["bahamas", "BS"], ["bahrain", "BH"], ["bangladesh", "BD"],
  ["barbados", "BB"], ["belarus", "BY"], ["belgium", "BE"],
  ["belize", "BZ"], ["benin", "BJ"], ["bhutan", "BT"],
  ["bolivia", "BO"], ["bosnia and herzegovina", "BA"], ["botswana", "BW"],
  ["brazil", "BR"], ["brunei", "BN"], ["bulgaria", "BG"],
  ["burkina faso", "BF"], ["burundi", "BI"],
]);

_add([
  ["cabo verde", "CV"], ["cape verde", "CV"], ["cambodia", "KH"],
  ["cameroon", "CM"], ["canada", "CA"], ["central african republic", "CF"],
  ["chad", "TD"], ["chile", "CL"], ["china", "CN"],
  ["colombia", "CO"], ["comoros", "KM"], ["congo", "CG"],
  ["costa rica", "CR"], ["croatia", "HR"], ["cuba", "CU"],
  ["cyprus", "CY"], ["czech republic", "CZ"], ["czechia", "CZ"],
]);

_add([
  ["denmark", "DK"], ["djibouti", "DJ"], ["dominica", "DM"],
  ["dominican republic", "DO"],
  ["dr congo", "CD"], ["democratic republic of the congo", "CD"],
]);

_add([
  ["ecuador", "EC"], ["egypt", "EG"], ["el salvador", "SV"],
  ["equatorial guinea", "GQ"], ["eritrea", "ER"], ["estonia", "EE"],
  ["eswatini", "SZ"], ["swaziland", "SZ"], ["ethiopia", "ET"],
]);

_add([
  ["fiji", "FJ"], ["finland", "FI"], ["france", "FR"],
  ["gabon", "GA"], ["gambia", "GM"], ["georgia", "GE"],
  ["germany", "DE"], ["ghana", "GH"], ["greece", "GR"],
  ["grenada", "GD"], ["guatemala", "GT"], ["guinea", "GN"],
  ["guinea-bissau", "GW"], ["guyana", "GY"],
]);

_add([
  ["haiti", "HT"], ["honduras", "HN"], ["hungary", "HU"],
  ["iceland", "IS"], ["india", "IN"], ["indonesia", "ID"],
  ["iran", "IR"], ["iraq", "IQ"], ["ireland", "IE"],
  ["israel", "IL"], ["italy", "IT"], ["ivory coast", "CI"],
  ["cote d'ivoire", "CI"],
]);

_add([
  ["jamaica", "JM"], ["japan", "JP"], ["jordan", "JO"],
  ["kazakhstan", "KZ"], ["kenya", "KE"], ["kiribati", "KI"],
  ["kosovo", "XK"], ["kuwait", "KW"], ["kyrgyzstan", "KG"],
]);

_add([
  ["laos", "LA"], ["latvia", "LV"], ["lebanon", "LB"],
  ["lesotho", "LS"], ["liberia", "LR"], ["libya", "LY"],
  ["liechtenstein", "LI"], ["lithuania", "LT"], ["luxembourg", "LU"],
]);

_add([
  ["madagascar", "MG"], ["malawi", "MW"], ["malaysia", "MY"],
  ["maldives", "MV"], ["mali", "ML"], ["malta", "MT"],
  ["marshall islands", "MH"], ["mauritania", "MR"], ["mauritius", "MU"],
  ["mexico", "MX"], ["micronesia", "FM"], ["moldova", "MD"],
  ["monaco", "MC"], ["mongolia", "MN"], ["montenegro", "ME"],
  ["morocco", "MA"], ["mozambique", "MZ"], ["myanmar", "MM"],
  ["burma", "MM"],
]);

_add([
  ["namibia", "NA"], ["nauru", "NR"], ["nepal", "NP"],
  ["netherlands", "NL"], ["new zealand", "NZ"], ["nicaragua", "NI"],
  ["niger", "NE"], ["nigeria", "NG"], ["north korea", "KP"],
  ["north macedonia", "MK"], ["macedonia", "MK"], ["norway", "NO"],
]);

_add([
  ["oman", "OM"], ["pakistan", "PK"], ["palau", "PW"],
  ["palestine", "PS"], ["panama", "PA"], ["papua new guinea", "PG"],
  ["paraguay", "PY"], ["peru", "PE"], ["philippines", "PH"],
  ["poland", "PL"], ["portugal", "PT"],
]);

_add([
  ["qatar", "QA"], ["romania", "RO"], ["russia", "RU"],
  ["rwanda", "RW"],
]);

_add([
  ["saint kitts and nevis", "KN"], ["saint lucia", "LC"],
  ["saint vincent and the grenadines", "VC"], ["samoa", "WS"],
  ["san marino", "SM"], ["sao tome and principe", "ST"],
  ["saudi arabia", "SA"], ["senegal", "SN"], ["serbia", "RS"],
  ["seychelles", "SC"], ["sierra leone", "SL"], ["singapore", "SG"],
  ["slovakia", "SK"], ["slovenia", "SI"], ["solomon islands", "SB"],
  ["somalia", "SO"], ["south africa", "ZA"], ["south korea", "KR"],
  ["south sudan", "SS"], ["spain", "ES"], ["sri lanka", "LK"],
  ["sudan", "SD"], ["suriname", "SR"], ["sweden", "SE"],
  ["switzerland", "CH"], ["syria", "SY"],
]);

_add([
  ["taiwan", "TW"], ["tajikistan", "TJ"], ["tanzania", "TZ"],
  ["thailand", "TH"], ["timor-leste", "TL"], ["east timor", "TL"],
  ["togo", "TG"], ["tonga", "TO"], ["trinidad and tobago", "TT"],
  ["tunisia", "TN"], ["turkey", "TR"], ["turkmenistan", "TM"],
  ["tuvalu", "TV"],
]);

_add([
  ["uganda", "UG"], ["ukraine", "UA"],
  ["united arab emirates", "AE"], ["uae", "AE"],
  ["united kingdom", "GB"], ["uk", "GB"], ["great britain", "GB"],
  ["england", "GB"], ["scotland", "GB"], ["wales", "GB"],
  ["united states", "US"], ["usa", "US"],
  ["united states of america", "US"],
  ["uruguay", "UY"], ["uzbekistan", "UZ"],
]);

_add([
  ["vanuatu", "VU"], ["vatican", "VA"], ["vatican city", "VA"],
  ["venezuela", "VE"], ["vietnam", "VN"], ["viet nam", "VN"],
  ["yemen", "YE"], ["zambia", "ZM"], ["zimbabwe", "ZW"],
]);

// Territories and common aliases
_add([
  ["hong kong", "HK"], ["macau", "MO"], ["macao", "MO"],
  ["puerto rico", "PR"], ["guam", "GU"],
  ["us virgin islands", "VI"], ["british virgin islands", "VG"],
  ["bermuda", "BM"], ["cayman islands", "KY"],
  ["curacao", "CW"], ["aruba", "AW"],
  ["greenland", "GL"], ["faroe islands", "FO"],
  ["french polynesia", "PF"], ["new caledonia", "NC"],
  ["guadeloupe", "GP"], ["martinique", "MQ"], ["reunion", "RE"],
  ["mayotte", "YT"], ["french guiana", "GF"],
  ["gibraltar", "GI"], ["isle of man", "IM"],
  ["jersey", "JE"], ["guernsey", "GG"],
]);

// Resolve: name or 2/3 letter code -> flag emoji
function resolveCountryFlag(input) {
  var val = (input || "").trim().toLowerCase();
  if (!val) return "\u{1F30D}";

  // Direct 2-letter ISO code match
  if (val.length === 2) {
    var upper = val.toUpperCase();
    // Check if it's a known code (all alpha-2 codes produce valid regional indicators)
    if (/^[A-Z]{2}$/.test(upper)) {
      return isoToFlag(upper);
    }
  }

  // 3-letter common aliases
  var threeLetterMap = {
    "usa": "US", "gbr": "GB", "can": "CA", "aus": "AU",
    "fra": "FR", "deu": "DE", "ger": "DE", "jpn": "JP",
    "chn": "CN", "ind": "IN", "bra": "BR", "mex": "MX",
    "ita": "IT", "esp": "ES", "kor": "KR", "rus": "RU",
    "nld": "NL", "tur": "TR", "che": "CH", "swe": "SE",
    "nor": "NO", "dnk": "DK", "fin": "FI", "pol": "PL",
    "aut": "AT", "bel": "BE", "prt": "PT", "irl": "IE",
    "nzl": "NZ", "arg": "AR", "col": "CO", "per": "PE",
    "chl": "CL", "ven": "VE", "ury": "UY", "pry": "PY",
    "bol": "BO", "ecu": "EC", "zaf": "ZA", "nga": "NG",
    "ken": "KE", "egy": "EG", "mar": "MA", "gha": "GH",
    "eth": "ET", "tza": "TZ", "uga": "UG", "are": "AE",
    "sau": "SA", "isr": "IL", "pak": "PK", "bgd": "BD",
    "tha": "TH", "idn": "ID", "phl": "PH", "mys": "MY",
    "sgp": "SG", "vnm": "VN", "twn": "TW", "hkg": "HK",
  };
  if (threeLetterMap[val]) {
    return isoToFlag(threeLetterMap[val]);
  }

  // Full name lookup
  if (COUNTRY_CODES[val]) {
    return isoToFlag(COUNTRY_CODES[val]);
  }

  // Partial match: check if input contains or is contained by a known name
  var keys = Object.keys(COUNTRY_CODES);
  for (var i = 0; i < keys.length; i++) {
    if (val.includes(keys[i]) || keys[i].includes(val)) {
      return isoToFlag(COUNTRY_CODES[keys[i]]);
    }
  }

  return "\u{1F30D}";
}
