// ============================================================
// logistics.js — Pure-JS logistics math for the Quince Portal.
// ============================================================
//
// Mirror of `Build/AppScriptV2/LogisticsCompute.gs` and
// `Build/python/logistics_bulk_compute.py`. When shipping team revises
// rates, update all three in lockstep:
//   1. Build/python/constants.py
//   2. Build/AppScriptV2/LogisticsCompute.gs
//   3. Build/web-portal/logistics.js (this file)
// Bump LOGISTICS_RATES_VERSION below.

const LOGISTICS_RATES_VERSION = '2026-05-22';

// ----------------------------------------------------------------------------
// Air tab — per-kg airfreight rates (cells N5-N14 in upstream calculator)
// ----------------------------------------------------------------------------
const LC_AIR_RATES = {
  VN: { LAX: 5.00,              ORD: 5.25 },
  IN: { LAX: 5.25,              ORD: 5.25 },
  BD: { LAX: 7.1 * 0.77 + 0.25, ORD: 7.1 * 0.82 + 0.25 },
  LK: { LAX: 7.0 * 0.77 + 0.25, ORD: 7.0 * 0.80 + 0.25 },
  KH: { LAX: 8.4 * 0.70 + 0.25, ORD: 8.4 * 0.74 + 0.25 }
};

const LC_AIR_MIN_RATIO = {
  SEA: 1.3, IN: 1.25, LK: 1.25, BD: 1.25, KH: 1.25
};

// ----------------------------------------------------------------------------
// Ocean container rates
// ----------------------------------------------------------------------------
const LC_OCEAN_CONTAINER_VN_LA  = 5000;
const LC_OCEAN_CONTAINER_VN_NY  = 0;
const LC_OCEAN_CONTAINER_IND_LA = 5000;
const LC_OCEAN_CONTAINER_IND_NY = 0;
const LC_OCEAN_CBM_PER_CONTAINER = 60;
const LC_OCEAN_KG_PER_CBM = 167;

// ----------------------------------------------------------------------------
// Static line-item constants
// ----------------------------------------------------------------------------
const LC_DIM_FACTOR = 200;
const LC_DIM_TO_GRAMS = 453;
const LC_CUSTOMS_DUTY = 0.05;
const LC_TRUCK_TO_LA = 750;
const LC_TRUCK_PORT_TO_PA_FC = 1500;
const LC_RETRIEVAL_FIXED = 0.15;
const LC_RETRIEVAL_VAR = 0.23;
const LC_PALLETS_PER_TRUCK = 16;
const LC_CARTONS_PER_PALLET = 12;

// ----------------------------------------------------------------------------
// Weight bucket table (VLOOKUP approximate match)
// ----------------------------------------------------------------------------
const LC_WEIGHT_BUCKETS = [
  [100, 10], [200, 10], [300, 12.5], [400, 12.5], [500, 15],
  [600, 15], [700, 15], [800, 17.5], [900, 17.5], [1000, 17.5]
];

// ----------------------------------------------------------------------------
// COO → region maps
// ----------------------------------------------------------------------------
const LC_AIR_REGION_MAP = {
  IN: 'IN', LK: 'LK', BD: 'BD', KH: 'KH',
  VN: 'SEA', CN: 'SEA', ID: 'SEA', PH: 'SEA', MY: 'SEA',
  TH: 'SEA', SG: 'SEA', NP: 'SEA', PK: 'SEA'
};
const LC_AIR_REGION_DEFAULT = 'SEA';

const LC_OCEAN_REGION_MAP = { IN: 'IN' };
const LC_OCEAN_REGION_DEFAULT = 'SEA';

// ----------------------------------------------------------------------------
// Country-name → ISO code (matches Build/python/logistics_bulk_compute.py)
// ----------------------------------------------------------------------------
const COUNTRY_NAME_TO_CODE = {
  'united arab emirates': 'AE', 'bangladesh': 'BD', 'canada': 'CA',
  'china': 'CN', 'indonesia': 'ID', 'israel': 'IL', 'india': 'IN',
  'cambodia': 'KH', 'sri lanka': 'LK', 'morocco': 'MA', 'mexico': 'MX',
  'malaysia': 'MY', 'peru': 'PE', 'philippines': 'PH', 'pakistan': 'PK',
  'singapore': 'SG', 'thailand': 'TH', 'turkey': 'TR',
  'taiwan, province of china': 'TW', 'united states': 'US', 'viet nam': 'VN',
  'vietnam': 'VN', 'bulgaria': 'BG', 'spain': 'ES', 'italy': 'IT',
  'portugal': 'PT', 'hong kong': 'HK', 'jordan': 'JO', 'guatemala': 'GT',
  'republic of korea': 'KR', 'south korea': 'KR', 'korea': 'KR',
  'bahrain': 'BH', 'dominican republic': 'DO', 'el salvador': 'SV',
  'brazil': 'BR', 'ecuador': 'EC', 'uk': 'GB', 'united kingdom': 'GB',
  'nepal': 'NP', 'trinidad and tobago': 'TT', 'japan': 'JP'
};

// Valid 2-letter ISO codes (all values from COUNTRY_NAME_TO_CODE).
const VALID_COO_CODES = (function () {
  const s = new Set();
  Object.values(COUNTRY_NAME_TO_CODE).forEach(c => s.add(c));
  return s;
})();

// Normalize a COO cell value into a 2-letter ISO code, or null if unrecognized.
// Accepted formats:
//   - 2-letter ISO code: 'IN', 'CN', 'VN'
//   - Full country name: 'India', 'Viet Nam'
//   - Combined: 'India | IN'
//   - "System default (Country)" — the Bid Inputs dropdown's default option;
//     extract the parenthesized country and resolve via the lookups above.
function normalizeCOO(raw) {
  if (raw === null || raw === undefined) return null;
  const s = String(raw).trim();
  if (!s) return null;

  // Phase 7.1: "System default (Country)" → extract country inside parens.
  // Only triggers when the string starts with "System default" — narrow rule
  // to avoid mis-parsing arbitrary "Vendor (India)" values.
  if (/^system\s+default\b/i.test(s)) {
    const m = s.match(/\(([^)]+)\)/);
    if (!m) return null;   // bare "System default" with no parens can't be resolved
    return normalizeCOO(m[1].trim());
  }

  // Combined "Name | XX" — take part after |
  if (s.indexOf('|') >= 0) {
    const parts = s.split('|');
    const code = parts[parts.length - 1].trim().toUpperCase();
    return VALID_COO_CODES.has(code) ? code : null;
  }

  // 2-letter ISO code
  if (s.length === 2 && /^[A-Za-z]{2}$/.test(s)) {
    const code = s.toUpperCase();
    return VALID_COO_CODES.has(code) ? code : null;
  }

  // Full country name (case-insensitive)
  const code = COUNTRY_NAME_TO_CODE[s.toLowerCase()];
  return code || null;
}

// ============================================================
// Pure math helpers
// ============================================================

function lc_billedWeightG(weightG, lIn, wIn, hIn) {
  const dimG = (lIn * wIn * hIn) / LC_DIM_FACTOR * LC_DIM_TO_GRAMS;
  return Math.max(weightG, dimG);
}

function lc_avgWeightPerCarton(s10Grams) {
  let avg = 10;
  for (let i = 0; i < LC_WEIGHT_BUCKETS.length; i++) {
    const [k, v] = LC_WEIGHT_BUCKETS[i];
    if (k <= s10Grams) avg = v;
    else break;
  }
  return avg;
}

function lc_unitsPerTruck(weightG, billedG) {
  const s10 = Math.max(Math.min(weightG * 1.3, billedG), 100);
  const s11 = lc_avgWeightPerCarton(s10);
  const s12 = s11 / (s10 / 1000.0);
  return s12 * LC_CARTONS_PER_PALLET * LC_PALLETS_PER_TRUCK;
}

function lc_oceanPerUnitContainer(weightG, billedG, containerRate) {
  if (containerRate === 0) return 0;
  const totalKgs = LC_OCEAN_CBM_PER_CONTAINER * LC_OCEAN_KG_PER_CBM;
  const denomKg = Math.min(weightG * 1.25, billedG) / 1000.0;
  if (denomKg <= 0) return 0;
  const numUnits = Math.ceil(totalKgs / denomKg * 0.80);
  if (numUnits <= 0) return 0;
  return containerRate / numUnits;
}

// ============================================================
// Air compute
// ============================================================
// Sums 4 line items: First Mile, Retrieval, Customs, Truck to FC.
// Averages LA + PA. Region mapping: IN, LK, BD, KH have own lanes;
// everything else (CN, VN, SEA group) → SEA (uses VN rates).
function lc_computeAir(weightG, lIn, wIn, hIn, coo) {
  const billed = lc_billedWeightG(weightG, lIn, wIn, hIn);
  const unitsPerTruck = lc_unitsPerTruck(weightG, billed);

  const region = LC_AIR_REGION_MAP[coo] || LC_AIR_REGION_DEFAULT;
  let rates;
  if (region === 'IN')      rates = LC_AIR_RATES.IN;
  else if (region === 'LK') rates = LC_AIR_RATES.LK;
  else if (region === 'BD') rates = LC_AIR_RATES.BD;
  else if (region === 'KH') rates = LC_AIR_RATES.KH;
  else                       rates = LC_AIR_RATES.VN;   // SEA fallback

  const minRatio = LC_AIR_MIN_RATIO[region] || 1.25;
  const firstMileWeight = Math.min(weightG * minRatio, billed) / 1000.0;
  const firstMileLa = firstMileWeight * rates.LAX;
  const firstMilePa = firstMileWeight * rates.ORD;

  const retrieval = LC_RETRIEVAL_FIXED + LC_RETRIEVAL_VAR * weightG * 1.3 / 1000.0;
  const customs = LC_CUSTOMS_DUTY;
  const truckToFcLa = LC_TRUCK_TO_LA / unitsPerTruck;
  const truckToFcPa = LC_TRUCK_PORT_TO_PA_FC / unitsPerTruck;

  const laTotal = firstMileLa + retrieval + customs + truckToFcLa;
  const paTotal = firstMilePa + retrieval + customs + truckToFcPa;
  return (laTotal + paTotal) / 2.0;
}

// ============================================================
// Ocean compute
// ============================================================
// Sums 4 line items: First Mile, Retrieval, Customs, Truck to FC.
// Averages LA + PA. Region mapping: IN has own lane; everything else → SEA.
function lc_computeOcean(weightG, lIn, wIn, hIn, coo) {
  const billed = lc_billedWeightG(weightG, lIn, wIn, hIn);
  const unitsPerTruck = lc_unitsPerTruck(weightG, billed);

  const region = LC_OCEAN_REGION_MAP[coo] || LC_OCEAN_REGION_DEFAULT;
  let containerLa, containerNy;
  if (region === 'IN') {
    containerLa = LC_OCEAN_CONTAINER_IND_LA;
    containerNy = LC_OCEAN_CONTAINER_IND_NY;
  } else {
    containerLa = LC_OCEAN_CONTAINER_VN_LA;
    containerNy = LC_OCEAN_CONTAINER_VN_NY;
  }

  const firstMileLa = lc_oceanPerUnitContainer(weightG, billed, containerLa);
  const firstMilePa = lc_oceanPerUnitContainer(weightG, billed, containerNy);
  const retrieval = LC_RETRIEVAL_FIXED + LC_RETRIEVAL_VAR * weightG * 1.3 / 1000.0;
  const customs = LC_CUSTOMS_DUTY;
  const truckToFcLa = LC_TRUCK_TO_LA / unitsPerTruck;
  const truckToFcPa = LC_TRUCK_PORT_TO_PA_FC / unitsPerTruck;

  const laTotal = firstMileLa + retrieval + customs + truckToFcLa;
  const paTotal = firstMilePa + retrieval + customs + truckToFcPa;
  return (laTotal + paTotal) / 2.0;
}

// ============================================================
// Public entry point
// ============================================================
function computeLogistics(weightG, lIn, wIn, hIn, cooCode) {
  return {
    ocean: lc_computeOcean(weightG, lIn, wIn, hIn, cooCode),
    air:   lc_computeAir(weightG, lIn, wIn, hIn, cooCode)
  };
}
