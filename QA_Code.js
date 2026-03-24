/**
 * QA dashboard — aligned with Trainer (Code.js): access, points, and manager CSV follow one rule:
 * - App access: QA_MANAGER_EMAILS or Active Roster (Col B = Image Production, Col C contains "QA").
 * - Points & export: only managers plus roster QA; QA (Q/U/AM) and training (G–L) count only for roster emails.
 */
/** QA Dashboard CONFIG — all names prefixed with QA_ to avoid clashes with Code.js in same project **/
const QA_SHEET_NAME = '[Image] Prod';
const QA_TZ = 'America/New_York';

// Column indices (1-based)
const QA_COL_I_TYPE = 9;    // I — type: Eval, RSHF, hlrm, categories (case-insensitive)
const QA_COL_L_COMPLET = 12; // L — completion multiplier (all types: base points * Col L)
const QA_COL_N_FLAG = 14;    // N — if true, add points to agent in P
const QA_COL_P_TS = 16;      // P — timestamp for M / row
const QA_COL_Q_EMAIL = 17;   // Q — agent email (points when M true)
const QA_COL_R_FLAG = 18;    // R — if true, add points to agent in T
const QA_COL_T_TS = 20;      // T — timestamp for Q
const QA_COL_U_EMAIL = 21;   // U — agent email (points when Q true)
const QA_COL_AK_FLAG = 37;   // AK — if true (RSHF/hlrm only), add points to agent in AL
const QA_COL_AL_TS = 38;     // AL — timestamp for AJ (fallback: O)
const QA_COL_AM_EMAIL = 39;  // AM — agent email (points when AJ true)

// Points by type (Col I) and flag. All case-insensitive. All points = base * Col L (completion).
// Eval: N true → 20k*L to Q, R true → 5k*L to U
// RSHF: N true → 30k*L to Q, R true → 15k*L to U, AK true → 10k*L to AM
// hlrm: N true → 60k*L to Q, R true → 20k*L to U, AK true → 60k*L to AM
// categories: N true → 20k*L to Q, R true → 10k*L to U, AK true → 5k*L to AM
// out: R true → 20k*L to U; AK true → 15k*L to AM; N true & !AK → 30k*L to Q; N true & AK → 15k*L to Q
const QA_POINTS_EVAL_M = 20000;
const QA_POINTS_EVAL_Q = 5000;
const QA_POINTS_RSHF_M = 30000;
const QA_POINTS_RSHF_Q = 15000;
const QA_POINTS_RSHF_AJ = 10000;
const QA_POINTS_HLRM_M = 60000;
const QA_POINTS_HLRM_Q = 20000;
const QA_POINTS_HLRM_AJ = 60000;
const QA_POINTS_CATEGORIES_M = 20000;
const QA_POINTS_CATEGORIES_Q = 10000;
const QA_POINTS_CATEGORIES_AJ = 5000;
const QA_POINTS_OUT_R = 20000;      // R true → to U
const QA_POINTS_OUT_AK = 15000;    // AK true → to AM
const QA_POINTS_OUT_N_ONLY = 30000;  // N true, AK false → to Q
const QA_POINTS_OUT_N_AND_AK = 15000; // N true, AK true → to Q

const QA_PERIOD_DAYS = 7;
const QA_MULTIPLIER_NUM_PERIODS = 4;

/** Periods are 7-day windows starting on this date (in QA_TZ). Set to your program start. */
const QA_PERIOD_START_DATE = new Date(2026, 1, 23);

// Pay: 7-day period. Multiplier: 4-period average (forward-looking from today).
const QA_PAY_THRESHOLDS = [
  { points: 1200000, amount: 75 },
  { points: 1550000, amount: 90 },
  { points: 1900000, amount: 100 }
];

const QA_MULTIPLIER_THRESHOLDS = [
  { avgMin: 1200000, value: 1, badgeLabel: '1.0x Multiplier Active' },
  { avgMin: 1550000, value: 1.1, badgeLabel: '1.1x Multiplier Active' },
  { avgMin: 1900000, value: 1.25, badgeLabel: '1.25x Multiplier Active' }
];

/** Managers who can export the full QA report (CSV) and always have QA dashboard access. */
const QA_MANAGER_EMAILS = [
  'lydia.huang@invisible.email',
  'michael.hernandez@invisible.email',
  'rana.traboulsi@invisible.email'
];

const QA_ROSTER_SHEET_NAME = 'Active Roster';
const QA_ROSTER_COL_B_MATCH = 'image production';
/** Col C must contain this substring (case-insensitive), e.g. "QA". */
const QA_ROSTER_COL_C_CONTAINS = 'qa';

/** Training rows use QA period windows (QA_getPeriodBounds_) for date filters; statuses/points match Code.js. */
const QA_TR_ALLOWED_STATUSES = ['Task Submitted', 'Revised'];
const QA_TR_POINTS_RSHF = 75000;
const QA_TR_POINTS_EVALS = 20000;
const QA_TR_POINTS_HLRM = 50000;
const QA_TR_POINTS_CATEGORIES = 30000;
const QA_TR_POINTS_MULTI_OUT = 35000;

function QA_normalizeEmail_(str) {
  const s = String(str || '').trim();
  const match = s.match(/\<([^\>]+)\>/);
  const emailOnly = match ? match[1].trim() : s;
  return emailOnly.toLowerCase();
}

/**
 * Col A emails on Active Roster where Col B is Image Production and Col C contains "QA" (case-insensitive).
 * Returns null if sheet missing; otherwise a Set of normalized emails.
 */
function QA_getActiveRosterQaEmails_(ss) {
  const sh = ss.getSheetByName(QA_ROSTER_SHEET_NAME);
  if (!sh) return null;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return new Set();
  const a = sh.getRange(2, 1, lastRow, 1).getValues();
  const b = sh.getRange(2, 2, lastRow, 2).getValues();
  const c = sh.getRange(2, 3, lastRow, 3).getValues();
  const sub = String(QA_ROSTER_COL_C_CONTAINS || '').toLowerCase();
  const out = new Set();
  for (var i = 0; i < a.length; i++) {
    var bVal = String(b[i][0] || '').trim().toLowerCase();
    var cVal = String(c[i][0] || '').trim().toLowerCase();
    if (bVal !== QA_ROSTER_COL_B_MATCH || cVal.indexOf(sub) === -1) continue;
    var em = QA_normalizeEmail_(a[i][0]);
    if (em) out.add(em);
  }
  return out;
}

/** Map Col I (lower) to one of five batch keys (same order as Code.js basePoints). */
function QA_trainerBatchKeyFromH_(hRaw) {
  if (hRaw.indexOf('rshf') !== -1) return 'RSHF';
  if (hRaw.indexOf('evals') !== -1 || hRaw.indexOf('eval') !== -1 || hRaw.indexOf('ema') !== -1 || (hRaw.indexOf('hrm') !== -1 && hRaw.indexOf('hlrm') === -1)) return 'Eval';
  if (hRaw.indexOf('hlrm') !== -1) return 'HLRM';
  if (hRaw.indexOf('categories') !== -1) return 'Categories';
  if (hRaw.indexOf('multi-out') !== -1) return 'Multi_out';
  return null;
}

/**
 * QA period breakdown + training batch totals → five combined totals (Eval, RSHF, HLRM, Categories, Multi_out).
 */
function QA_mergeCombinedBatchPeriod_(periodBreakdown, trainingBatch) {
  periodBreakdown = periodBreakdown || {};
  trainingBatch = trainingBatch || {};
  var trE = Number(trainingBatch.Eval) || 0;
  var trR = Number(trainingBatch.RSHF) || 0;
  var trH = Number(trainingBatch.HLRM) || 0;
  var trC = Number(trainingBatch.Categories) || 0;
  var trM = Number(trainingBatch.Multi_out) || 0;
  return {
    Eval: (Number(periodBreakdown.Eval_M) || 0) + (Number(periodBreakdown.Eval_Q) || 0) + trE,
    RSHF: (Number(periodBreakdown.RSHF_M) || 0) + (Number(periodBreakdown.RSHF_Q) || 0) + (Number(periodBreakdown.RSHF_AJ) || 0) + trR,
    HLRM: (Number(periodBreakdown.HLRM_M) || 0) + (Number(periodBreakdown.HLRM_Q) || 0) + (Number(periodBreakdown.HLRM_AJ) || 0) + trH,
    Categories: (Number(periodBreakdown.Categories_M) || 0) + (Number(periodBreakdown.Categories_Q) || 0) + (Number(periodBreakdown.Categories_AJ) || 0) + trC,
    Multi_out: (Number(periodBreakdown.Out_R) || 0) + (Number(periodBreakdown.Out_AK) || 0) + (Number(periodBreakdown.Out_N_only) || 0) + (Number(periodBreakdown.Out_N_and_AK) || 0) + trM
  };
}

/**
 * Training rows from [Image] Prod Col G–L + AJ (sheet order), no date filter.
 * Same eligibility as QA_computeTrainerPointsFromProd_ / Code.js trainer parse.
 */
function QA_parseTrainerProdRows_(ss, rosterQaSet) {
  var out = [];
  const sh = ss.getSheetByName(QA_SHEET_NAME);
  if (!sh) return out;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return out;

  var valuesMain = sh.getRange(2, 7, lastRow, 12).getValues();
  var valuesError = sh.getRange(2, 36, lastRow, 36).getValues();

  for (var i = 0; i < valuesMain.length; i++) {
    var main = valuesMain[i];
    var rowEmail = QA_normalizeEmail_(main[1]);
    if (!rowEmail) continue;
    if (!rosterQaSet.has(rowEmail)) continue;

    var ts = main[0];
    if (!(ts instanceof Date)) continue;

    var st = String(main[3] || '').trim();
    if (QA_TR_ALLOWED_STATUSES.indexOf(st) === -1) continue;

    var kVal = Number(main[5] || 0);
    if (!isFinite(kVal)) continue;

    var hRaw = String(main[2] || '').trim().toLowerCase();
    var basePoints = 0;
    if (hRaw.indexOf('rshf') !== -1) basePoints = QA_TR_POINTS_RSHF;
    else if (hRaw.indexOf('evals') !== -1 || hRaw.indexOf('eval') !== -1 || hRaw.indexOf('ema') !== -1 || (hRaw.indexOf('hrm') !== -1 && hRaw.indexOf('hlrm') === -1)) basePoints = QA_TR_POINTS_EVALS;
    else if (hRaw.indexOf('hlrm') !== -1) basePoints = QA_TR_POINTS_HLRM;
    else if (hRaw.indexOf('categories') !== -1) basePoints = QA_TR_POINTS_CATEGORIES;
    else if (hRaw.indexOf('multi-out') !== -1) basePoints = QA_TR_POINTS_MULTI_OUT;
    if (basePoints === 0) continue;

    var pts = basePoints * kVal;
    var c = kVal;
    var errVal = isFinite(Number(valuesError[i][0])) ? Number(valuesError[i][0]) : 0;
    var dayStr = QA_toDateStrInTz_(ts, QA_TZ);
    out.push({ email: rowEmail, dayStr: dayStr, pts: pts, c: c, errVal: errVal, hRaw: hRaw });
  }
  return out;
}

/**
 * Multi-window for period index P matches QA_buildReportPeriods_ (sum periods multiStart..P).
 */
function QA_getTrainerBoundsForReportPeriod_(periodIndex, tz) {
  var multiStart = Math.max(0, periodIndex - QA_MULTIPLIER_NUM_PERIODS + 1);
  var bCur = QA_getPeriodBoundsByIndex_(periodIndex, tz);
  var bMulti = QA_getPeriodBoundsByIndex_(multiStart, tz);
  return {
    periodStartStr: bCur.periodStartStr,
    periodEndStr: bCur.periodEndStr,
    multiPeriodStartStr: bMulti.periodStartStr,
    multiPeriodEndStr: bCur.periodEndStr,
    periodStart: bCur.periodStart,
    periodEnd: bCur.periodEnd
  };
}

/**
 * Same pool / ceiling logic as dashboard; bounds select which rows count for period vs multi.
 */
function QA_accumulateTrainerFromParsed_(parsedRows, bounds) {
  var empty = { periodTotalsByEmail: {}, multiPeriodPointsByEmail: {}, periodBatchByEmail: {} };
  if (!parsedRows || !parsedRows.length) return empty;

  var periodStartStr = bounds.periodStartStr;
  var periodEndStr = bounds.periodEndStr;
  var multiPeriodStartStr = bounds.multiPeriodStartStr;
  var multiPeriodEndStr = bounds.multiPeriodEndStr;

  var inMulti = [];
  for (var i = 0; i < parsedRows.length; i++) {
    var pr = parsedRows[i];
    if (pr.dayStr >= multiPeriodStartStr && pr.dayStr <= multiPeriodEndStr) inMulti.push(pr);
  }

  var validPoolC = Object.create(null);
  var validPoolE = Object.create(null);
  var periodTotalsByEmail = Object.create(null);
  var multiPeriodPointsByEmail = Object.create(null);
  var periodBatchByEmail = Object.create(null);

  for (var j = 0; j < inMulti.length; j++) {
    var r = inMulti[j];
    var rowEmail = r.email;
    var inCurrentPeriod = r.dayStr >= periodStartStr && r.dayStr <= periodEndStr;
    var curC = validPoolC[rowEmail] || 0;
    var curE = validPoolE[rowEmail] || 0;
    var newC = curC + r.c;
    var newE = curE + r.errVal;
    var underCeiling = (newC > 8 && newE <= 0.125 * newC) || (newC <= 8 && newE <= 1);

    if (underCeiling) {
      validPoolC[rowEmail] = newC;
      validPoolE[rowEmail] = newE;
      multiPeriodPointsByEmail[rowEmail] = (multiPeriodPointsByEmail[rowEmail] || 0) + r.pts;
      if (inCurrentPeriod) {
        periodTotalsByEmail[rowEmail] = (periodTotalsByEmail[rowEmail] || 0) + r.pts;
        if (!periodBatchByEmail[rowEmail]) {
          periodBatchByEmail[rowEmail] = { Eval: 0, RSHF: 0, HLRM: 0, Categories: 0, Multi_out: 0 };
        }
        var batchKey = QA_trainerBatchKeyFromH_(r.hRaw);
        if (batchKey) periodBatchByEmail[rowEmail][batchKey] += r.pts;
      }
    }
  }

  return { periodTotalsByEmail: periodTotalsByEmail, multiPeriodPointsByEmail: multiPeriodPointsByEmail, periodBatchByEmail: periodBatchByEmail };
}

/**
 * Training points from [Image] Prod Col G–L + AJ (same row rules / point constants as Code.js).
 * Date filters ONLY use QA 7-day windows (QA_getPeriodBounds_ / QA_PERIOD_DAYS). Do not use Code.js getPeriodBounds_ — the Trainer dashboard (Code.js) keeps its own 9-day competition periods (PERIOD_DAYS = 9) there.
 * If qaBounds is omitted, uses QA_getPeriodBounds_(now).
 */
function QA_computeTrainerPointsFromProd_(ss, rosterQaSet, qaBounds) {
  var empty = { periodTotalsByEmail: {}, multiPeriodPointsByEmail: {}, periodBatchByEmail: {} };
  var parsed = QA_parseTrainerProdRows_(ss, rosterQaSet);
  if (!parsed.length) return empty;
  var bounds = qaBounds || QA_getPeriodBounds_(new Date(), QA_TZ);
  return QA_accumulateTrainerFromParsed_(parsed, bounds);
}

/** Web app entry for QA — use this as the deployment "entry point" for the QA app. */
function doGetQA() {
  return HtmlService.createHtmlOutputFromFile('QA_Index')
    .setTitle('QA Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/** True if the current user is in QA_MANAGER_EMAILS. Use when deployed "Execute as: User". */
function QA_isManager() {
  try {
    var raw = Session.getActiveUser().getEmail();
    var email = raw ? QA_normalizeEmail_(raw) : '';
    if (!email) return false;
    var list = QA_MANAGER_EMAILS.map(function(e) { return QA_normalizeEmail_(e); }).filter(Boolean);
    return list.indexOf(email) !== -1;
  } catch (e) {
    return false;
  }
}

/**
 * Main data endpoint for QA dashboard.
 * Access: Active Roster QA (Image Production + Col C contains "QA") OR QA_MANAGER_EMAILS — same population as manager CSV rows.
 */
function getQADashboardData() {
  var email = '';
  try {
    const raw = Session.getActiveUser().getEmail();
    email = raw ? QA_normalizeEmail_(raw) : '';
    if (!email) { Logger.log('QA getDashboardData: no email'); return { noAccess: true }; }

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(QA_SHEET_NAME);
    if (!sh) { Logger.log('QA getDashboardData: sheet not found ' + QA_SHEET_NAME); return { error: 'Sheet "' + QA_SHEET_NAME + '" not found' }; }

    const rosterQa = QA_getActiveRosterQaEmails_(ss);
    if (rosterQa === null) return { error: 'Sheet "' + QA_ROSTER_SHEET_NAME + '" not found' };

    const managerNorm = QA_MANAGER_EMAILS.map(function(e) { return QA_normalizeEmail_(e); }).filter(Boolean);
    var allowed = rosterQa.has(email) || managerNorm.indexOf(email) !== -1;
    if (!allowed) { Logger.log('QA getDashboardData: access denied ' + email); return { noAccess: true }; }

    const lastRow = sh.getLastRow();
    Logger.log('QA getDashboardData: access ok, loading data for ' + email + ', lastRow=' + lastRow);
    var result = QA_getDashboardDataForEmail_(email);
    Logger.log('QA getDashboardData: success for ' + email);
    return result;
  } catch (err) {
    Logger.log('QA getDashboardData ERROR for ' + email + ': ' + (err && err.message ? err.message : err));
    return { error: String(err && err.message ? err.message : err) };
  }
}

/**
 * Per-user QA data. Caller must have already verified access (roster QA or manager).
 * Roster-only points: QA column scores and training scores apply only when email is on Active Roster QA.
 */
function QA_getDashboardDataForEmail_(email) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(QA_SHEET_NAME);
  if (!sh) throw new Error('Sheet "' + QA_SHEET_NAME + '" not found');

  const rosterQa = QA_getActiveRosterQaEmails_(ss);
  if (rosterQa === null) throw new Error('Sheet "' + QA_ROSTER_SHEET_NAME + '" not found');

  const now = new Date();
  const bounds = QA_getPeriodBounds_(now, QA_TZ);
  const periodStartStr = bounds.periodStartStr;
  const periodEndStr = bounds.periodEndStr;
  const multiPeriodStartStr = bounds.multiPeriodStartStr;
  const multiPeriodEndStr = bounds.multiPeriodEndStr;
  const periodStart = bounds.periodStart;
  const periodEnd = bounds.periodEnd;
  const numPeriodsForAvg = bounds.numPeriodsForAvg || QA_MULTIPLIER_NUM_PERIODS;
  const periodRangeText =
    Utilities.formatDate(periodStart, QA_TZ, 'M/d/yyyy') + ' - ' + Utilities.formatDate(periodEnd, QA_TZ, 'M/d/yyyy');

  function makeBreakdown() {
    return { Eval_M: 0, Eval_Q: 0, RSHF_M: 0, RSHF_Q: 0, RSHF_AJ: 0, HLRM_M: 0, HLRM_Q: 0, HLRM_AJ: 0, Categories_M: 0, Categories_Q: 0, Categories_AJ: 0, Out_R: 0, Out_AK: 0, Out_N_only: 0, Out_N_and_AK: 0 };
  }

  const lastRow = sh.getLastRow();
    if (lastRow < 2) {
    var trainerEmpty = QA_computeTrainerPointsFromProd_(ss, rosterQa, bounds);
    var trPts0 = rosterQa.has(email) ? (trainerEmpty.periodTotalsByEmail[email] || 0) : 0;
    var trMulti0 = rosterQa.has(email) ? (trainerEmpty.multiPeriodPointsByEmail[email] || 0) : 0;
    var trBatch0 = rosterQa.has(email) && trainerEmpty.periodBatchByEmail && trainerEmpty.periodBatchByEmail[email] ? trainerEmpty.periodBatchByEmail[email] : {};
    var bd0 = makeBreakdown();
    return QA_buildResponse_(email, 0, 0, periodRangeText, bd0, bd0, trPts0, trMulti0, numPeriodsForAvg, trBatch0);
  }

  const maxCol = Math.max(QA_COL_P_TS, QA_COL_T_TS, QA_COL_AL_TS, QA_COL_Q_EMAIL, QA_COL_U_EMAIL, QA_COL_AM_EMAIL, QA_COL_L_COMPLET);
  const values = sh.getRange(2, 1, lastRow, maxCol).getValues();

  let periodPoints = 0;
  let multiPeriodTotalPoints = 0;
  var periodBreakdown = makeBreakdown();
  var multiPeriodBreakdown = makeBreakdown();

  /** QA eval points (Q/U/AM) only count for people on Active Roster QA; managers not on roster see $0 here. */
  var countQaSheetPoints = rosterQa.has(email);

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var typeRaw = String(row[QA_COL_I_TYPE - 1] || '').trim().toLowerCase();
    if (!typeRaw) continue; // Col I empty → 0 points for this row
    var completion = isFinite(Number(row[QA_COL_L_COMPLET - 1])) ? Number(row[QA_COL_L_COMPLET - 1]) : 0;
    var isEval = typeRaw.indexOf('eval') !== -1 || (typeRaw.indexOf('hrm') !== -1 && typeRaw.indexOf('hlrm') === -1);
    var isRSHF = typeRaw.indexOf('rshf') !== -1;
    var isHLRM = typeRaw.indexOf('hlrm') !== -1;
    var isCategories = typeRaw.indexOf('categories') !== -1;
    var isOut = typeRaw.indexOf('multi-out') !== -1;

    var mTrue = row[QA_COL_N_FLAG - 1] === true || String(row[QA_COL_N_FLAG - 1] || '').trim().toUpperCase() === 'TRUE';
    var qTrue = row[QA_COL_R_FLAG - 1] === true || String(row[QA_COL_R_FLAG - 1] || '').trim().toUpperCase() === 'TRUE';
    var ajTrue = row[QA_COL_AK_FLAG - 1] === true || String(row[QA_COL_AK_FLAG - 1] || '').trim().toUpperCase() === 'TRUE';

    var tsO = row[QA_COL_P_TS - 1];
    var tsS = row[QA_COL_T_TS - 1];
    var tsAJ = row[QA_COL_AL_TS - 1] instanceof Date ? row[QA_COL_AL_TS - 1] : tsO;

    var agentP, agentT, agentAL;

    if (isEval) {
      if (mTrue && tsO instanceof Date) {
        agentP = QA_normalizeEmail_(row[QA_COL_Q_EMAIL - 1]);
        if (countQaSheetPoints && agentP === email) {
          var ptsM = QA_POINTS_EVAL_M * completion;
          if (QA_isDateInRange_(tsO, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsM; periodBreakdown.Eval_M += ptsM; }
          if (QA_isDateInRange_(tsO, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsM; multiPeriodBreakdown.Eval_M += ptsM; }
        }
      }
      if (qTrue && tsS instanceof Date) {
        agentT = QA_normalizeEmail_(row[QA_COL_U_EMAIL - 1]);
        if (countQaSheetPoints && agentT === email) {
          var ptsQ = QA_POINTS_EVAL_Q * completion;
          if (QA_isDateInRange_(tsS, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsQ; periodBreakdown.Eval_Q += ptsQ; }
          if (QA_isDateInRange_(tsS, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsQ; multiPeriodBreakdown.Eval_Q += ptsQ; }
        }
      }
    }

    if (isRSHF) {
      if (mTrue && tsO instanceof Date) {
        agentP = QA_normalizeEmail_(row[QA_COL_Q_EMAIL - 1]);
        if (countQaSheetPoints && agentP === email) {
          var ptsM = QA_POINTS_RSHF_M * completion;
          if (QA_isDateInRange_(tsO, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsM; periodBreakdown.RSHF_M += ptsM; }
          if (QA_isDateInRange_(tsO, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsM; multiPeriodBreakdown.RSHF_M += ptsM; }
        }
      }
      if (qTrue && tsS instanceof Date) {
        agentT = QA_normalizeEmail_(row[QA_COL_U_EMAIL - 1]);
        if (countQaSheetPoints && agentT === email) {
          var ptsQ = QA_POINTS_RSHF_Q * completion;
          if (QA_isDateInRange_(tsS, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsQ; periodBreakdown.RSHF_Q += ptsQ; }
          if (QA_isDateInRange_(tsS, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsQ; multiPeriodBreakdown.RSHF_Q += ptsQ; }
        }
      }
      if (ajTrue && tsAJ instanceof Date) {
        agentAL = QA_normalizeEmail_(row[QA_COL_AM_EMAIL - 1]);
        if (countQaSheetPoints && agentAL === email) {
          var ptsAJ = QA_POINTS_RSHF_AJ * completion;
          if (QA_isDateInRange_(tsAJ, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsAJ; periodBreakdown.RSHF_AJ += ptsAJ; }
          if (QA_isDateInRange_(tsAJ, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsAJ; multiPeriodBreakdown.RSHF_AJ += ptsAJ; }
        }
      }
    }

    if (isHLRM) {
      if (mTrue && tsO instanceof Date) {
        agentP = QA_normalizeEmail_(row[QA_COL_Q_EMAIL - 1]);
        if (countQaSheetPoints && agentP === email) {
          var ptsM = QA_POINTS_HLRM_M * completion;
          if (QA_isDateInRange_(tsO, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsM; periodBreakdown.HLRM_M += ptsM; }
          if (QA_isDateInRange_(tsO, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsM; multiPeriodBreakdown.HLRM_M += ptsM; }
        }
      }
      if (qTrue && tsS instanceof Date) {
        agentT = QA_normalizeEmail_(row[QA_COL_U_EMAIL - 1]);
        if (countQaSheetPoints && agentT === email) {
          var ptsQ = QA_POINTS_HLRM_Q * completion;
          if (QA_isDateInRange_(tsS, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsQ; periodBreakdown.HLRM_Q += ptsQ; }
          if (QA_isDateInRange_(tsS, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsQ; multiPeriodBreakdown.HLRM_Q += ptsQ; }
        }
      }
      if (ajTrue && tsAJ instanceof Date) {
        agentAL = QA_normalizeEmail_(row[QA_COL_AM_EMAIL - 1]);
        if (countQaSheetPoints && agentAL === email) {
          var ptsAJ = QA_POINTS_HLRM_AJ * completion;
          if (QA_isDateInRange_(tsAJ, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsAJ; periodBreakdown.HLRM_AJ += ptsAJ; }
          if (QA_isDateInRange_(tsAJ, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsAJ; multiPeriodBreakdown.HLRM_AJ += ptsAJ; }
        }
      }
    }

    if (isCategories) {
      if (mTrue && tsO instanceof Date) {
        agentP = QA_normalizeEmail_(row[QA_COL_Q_EMAIL - 1]);
        if (countQaSheetPoints && agentP === email) {
          var ptsM = QA_POINTS_CATEGORIES_M * completion;
          if (QA_isDateInRange_(tsO, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsM; periodBreakdown.Categories_M += ptsM; }
          if (QA_isDateInRange_(tsO, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsM; multiPeriodBreakdown.Categories_M += ptsM; }
        }
      }
      if (qTrue && tsS instanceof Date) {
        agentT = QA_normalizeEmail_(row[QA_COL_U_EMAIL - 1]);
        if (countQaSheetPoints && agentT === email) {
          var ptsQ = QA_POINTS_CATEGORIES_Q * completion;
          if (QA_isDateInRange_(tsS, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsQ; periodBreakdown.Categories_Q += ptsQ; }
          if (QA_isDateInRange_(tsS, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsQ; multiPeriodBreakdown.Categories_Q += ptsQ; }
        }
      }
      if (ajTrue && tsAJ instanceof Date) {
        agentAL = QA_normalizeEmail_(row[QA_COL_AM_EMAIL - 1]);
        if (countQaSheetPoints && agentAL === email) {
          var ptsAJ = QA_POINTS_CATEGORIES_AJ * completion;
          if (QA_isDateInRange_(tsAJ, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsAJ; periodBreakdown.Categories_AJ += ptsAJ; }
          if (QA_isDateInRange_(tsAJ, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsAJ; multiPeriodBreakdown.Categories_AJ += ptsAJ; }
        }
      }
    }

    if (isOut) {
      if (qTrue && tsS instanceof Date) {
        agentT = QA_normalizeEmail_(row[QA_COL_U_EMAIL - 1]);
        if (countQaSheetPoints && agentT === email) {
          var ptsR = QA_POINTS_OUT_R * completion;
          if (QA_isDateInRange_(tsS, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsR; periodBreakdown.Out_R += ptsR; }
          if (QA_isDateInRange_(tsS, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsR; multiPeriodBreakdown.Out_R += ptsR; }
        }
      }
      if (ajTrue && tsAJ instanceof Date) {
        agentAL = QA_normalizeEmail_(row[QA_COL_AM_EMAIL - 1]);
        if (countQaSheetPoints && agentAL === email) {
          var ptsAK = QA_POINTS_OUT_AK * completion;
          if (QA_isDateInRange_(tsAJ, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsAK; periodBreakdown.Out_AK += ptsAK; }
          if (QA_isDateInRange_(tsAJ, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsAK; multiPeriodBreakdown.Out_AK += ptsAK; }
        }
      }
      if (mTrue && tsO instanceof Date) {
        agentP = QA_normalizeEmail_(row[QA_COL_Q_EMAIL - 1]);
        if (countQaSheetPoints && agentP === email) {
          var ptsN = (ajTrue ? QA_POINTS_OUT_N_AND_AK : QA_POINTS_OUT_N_ONLY) * completion;
          if (QA_isDateInRange_(tsO, periodStartStr, periodEndStr, QA_TZ)) {
            periodPoints += ptsN;
            if (ajTrue) periodBreakdown.Out_N_and_AK += ptsN; else periodBreakdown.Out_N_only += ptsN;
          }
          if (QA_isDateInRange_(tsO, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) {
            multiPeriodTotalPoints += ptsN;
            if (ajTrue) multiPeriodBreakdown.Out_N_and_AK += ptsN; else multiPeriodBreakdown.Out_N_only += ptsN;
          }
        }
      }
    }
  }


  var trainerResult = QA_computeTrainerPointsFromProd_(ss, rosterQa, bounds);
  var trainingPeriodPoints = rosterQa.has(email) ? (trainerResult.periodTotalsByEmail[email] || 0) : 0;
  var trainingMultiPeriod = rosterQa.has(email) ? (trainerResult.multiPeriodPointsByEmail[email] || 0) : 0;
  var trainingBatch = rosterQa.has(email) && trainerResult.periodBatchByEmail && trainerResult.periodBatchByEmail[email] ? trainerResult.periodBatchByEmail[email] : {};

  return QA_buildResponse_(email, periodPoints, multiPeriodTotalPoints, periodRangeText, periodBreakdown, multiPeriodBreakdown, trainingPeriodPoints, trainingMultiPeriod, numPeriodsForAvg, trainingBatch);
}

/**
 * Load sheet once and build period points per email per period index (current + last 2 for manager report).
 * allEmails = QA managers ∪ Active Roster QA only. QA column points (Q/U/AM) only accrue for roster QA.
 * Returns { allEmails, periods: [{ periodRangeText, periodPointsByEmail, multiPeriodPointsByEmail, numPeriodsForAvg }] }.
 */
function QA_getAllAgentsDataForReport_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(QA_SHEET_NAME);
  if (!sh) return null;

  var rosterQa = QA_getActiveRosterQaEmails_(ss);
  if (rosterQa === null) return null;

  var managerNorm = QA_MANAGER_EMAILS.map(function(e) { return QA_normalizeEmail_(e); }).filter(Boolean);

  var now = new Date();
  var currentPeriodIndex = QA_getPeriodIndex_(now, QA_TZ);

  var lastRow = sh.getLastRow();
  if (lastRow < 2) {
    var periodsEmpty = QA_buildReportPeriods_({}, currentPeriodIndex, rosterQa, managerNorm);
    var parsedTrainerEmpty = QA_parseTrainerProdRows_(ss, rosterQa);
    for (var pe = 0; pe < periodsEmpty.length; pe++) {
      var pIdxE = periodsEmpty[pe].periodIndex;
      var tbE = QA_getTrainerBoundsForReportPeriod_(pIdxE, QA_TZ);
      var trE = QA_accumulateTrainerFromParsed_(parsedTrainerEmpty, tbE);
      periodsEmpty[pe].trainerPeriodByEmail = trE.periodTotalsByEmail;
      periodsEmpty[pe].trainerMultiPeriodByEmail = trE.multiPeriodPointsByEmail;
    }
    var allEmailsEmpty = [];
    var seenE = {};
    managerNorm.forEach(function(e) { if (!seenE[e]) { seenE[e] = true; allEmailsEmpty.push(e); } });
    rosterQa.forEach(function(e) { if (!seenE[e]) { seenE[e] = true; allEmailsEmpty.push(e); } });
    return { allEmails: allEmailsEmpty, periods: periodsEmpty };
  }

  var maxCol = Math.max(QA_COL_P_TS, QA_COL_T_TS, QA_COL_AL_TS, QA_COL_Q_EMAIL, QA_COL_U_EMAIL, QA_COL_AM_EMAIL, QA_COL_L_COMPLET);
  var values = sh.getRange(2, 1, lastRow, maxCol).getValues();

  // periodPointsByEmailByPeriod[email][periodIndex] = points — only Active Roster QA (same as dashboard QA columns)
  var periodPointsByEmailByPeriod = {};

  function addToAgentPeriod(agent, pts, ts) {
    if (!agent) return;
    if (!rosterQa.has(agent)) return;
    var pIdx = QA_getPeriodIndex_(ts, QA_TZ);
    if (!periodPointsByEmailByPeriod[agent]) periodPointsByEmailByPeriod[agent] = {};
    periodPointsByEmailByPeriod[agent][pIdx] = (periodPointsByEmailByPeriod[agent][pIdx] || 0) + pts;
  }

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var typeRaw = String(row[QA_COL_I_TYPE - 1] || '').trim().toLowerCase();
    if (!typeRaw) continue;
    var completion = isFinite(Number(row[QA_COL_L_COMPLET - 1])) ? Number(row[QA_COL_L_COMPLET - 1]) : 0;
    var isEval = typeRaw.indexOf('eval') !== -1 || (typeRaw.indexOf('hrm') !== -1 && typeRaw.indexOf('hlrm') === -1);
    var isRSHF = typeRaw.indexOf('rshf') !== -1;
    var isHLRM = typeRaw.indexOf('hlrm') !== -1;
    var isCategories = typeRaw.indexOf('categories') !== -1;
    var isOut = typeRaw.indexOf('multi-out') !== -1;

    var mTrue = row[QA_COL_N_FLAG - 1] === true || String(row[QA_COL_N_FLAG - 1] || '').trim().toUpperCase() === 'TRUE';
    var qTrue = row[QA_COL_R_FLAG - 1] === true || String(row[QA_COL_R_FLAG - 1] || '').trim().toUpperCase() === 'TRUE';
    var ajTrue = row[QA_COL_AK_FLAG - 1] === true || String(row[QA_COL_AK_FLAG - 1] || '').trim().toUpperCase() === 'TRUE';

    var tsO = row[QA_COL_P_TS - 1];
    var tsS = row[QA_COL_T_TS - 1];
    var tsAJ = row[QA_COL_AL_TS - 1] instanceof Date ? row[QA_COL_AL_TS - 1] : tsO;

    if (isEval) {
      if (mTrue && tsO instanceof Date) {
        var agentP = QA_normalizeEmail_(row[QA_COL_Q_EMAIL - 1]);
        if (agentP) addToAgentPeriod(agentP, QA_POINTS_EVAL_M * completion, tsO);
      }
      if (qTrue && tsS instanceof Date) {
        var agentT = QA_normalizeEmail_(row[QA_COL_U_EMAIL - 1]);
        if (agentT) addToAgentPeriod(agentT, QA_POINTS_EVAL_Q * completion, tsS);
      }
    }
    if (isRSHF) {
      if (mTrue && tsO instanceof Date) {
        agentP = QA_normalizeEmail_(row[QA_COL_Q_EMAIL - 1]);
        if (agentP) addToAgentPeriod(agentP, QA_POINTS_RSHF_M * completion, tsO);
      }
      if (qTrue && tsS instanceof Date) {
        agentT = QA_normalizeEmail_(row[QA_COL_U_EMAIL - 1]);
        if (agentT) addToAgentPeriod(agentT, QA_POINTS_RSHF_Q * completion, tsS);
      }
      if (ajTrue && tsAJ instanceof Date) {
        var agentAL = QA_normalizeEmail_(row[QA_COL_AM_EMAIL - 1]);
        if (agentAL) addToAgentPeriod(agentAL, QA_POINTS_RSHF_AJ * completion, tsAJ);
      }
    }
    if (isHLRM) {
      if (mTrue && tsO instanceof Date) {
        agentP = QA_normalizeEmail_(row[QA_COL_Q_EMAIL - 1]);
        if (agentP) addToAgentPeriod(agentP, QA_POINTS_HLRM_M * completion, tsO);
      }
      if (qTrue && tsS instanceof Date) {
        agentT = QA_normalizeEmail_(row[QA_COL_U_EMAIL - 1]);
        if (agentT) addToAgentPeriod(agentT, QA_POINTS_HLRM_Q * completion, tsS);
      }
      if (ajTrue && tsAJ instanceof Date) {
        agentAL = QA_normalizeEmail_(row[QA_COL_AM_EMAIL - 1]);
        if (agentAL) addToAgentPeriod(agentAL, QA_POINTS_HLRM_AJ * completion, tsAJ);
      }
    }
    if (isCategories) {
      if (mTrue && tsO instanceof Date) {
        agentP = QA_normalizeEmail_(row[QA_COL_Q_EMAIL - 1]);
        if (agentP) addToAgentPeriod(agentP, QA_POINTS_CATEGORIES_M * completion, tsO);
      }
      if (qTrue && tsS instanceof Date) {
        agentT = QA_normalizeEmail_(row[QA_COL_U_EMAIL - 1]);
        if (agentT) addToAgentPeriod(agentT, QA_POINTS_CATEGORIES_Q * completion, tsS);
      }
      if (ajTrue && tsAJ instanceof Date) {
        agentAL = QA_normalizeEmail_(row[QA_COL_AM_EMAIL - 1]);
        if (agentAL) addToAgentPeriod(agentAL, QA_POINTS_CATEGORIES_AJ * completion, tsAJ);
      }
    }
    if (isOut) {
      if (qTrue && tsS instanceof Date) {
        agentT = QA_normalizeEmail_(row[QA_COL_U_EMAIL - 1]);
        if (agentT) addToAgentPeriod(agentT, QA_POINTS_OUT_R * completion, tsS);
      }
      if (ajTrue && tsAJ instanceof Date) {
        agentAL = QA_normalizeEmail_(row[QA_COL_AM_EMAIL - 1]);
        if (agentAL) addToAgentPeriod(agentAL, QA_POINTS_OUT_AK * completion, tsAJ);
      }
      if (mTrue && tsO instanceof Date) {
        agentP = QA_normalizeEmail_(row[QA_COL_Q_EMAIL - 1]);
        if (agentP) addToAgentPeriod(agentP, (ajTrue ? QA_POINTS_OUT_N_AND_AK : QA_POINTS_OUT_N_ONLY) * completion, tsO);
      }
    }
  }

  var allEmails = [];
  var seen = {};
  managerNorm.forEach(function(e) { if (!seen[e]) { seen[e] = true; allEmails.push(e); } });
  rosterQa.forEach(function(e) { if (!seen[e]) { seen[e] = true; allEmails.push(e); } });

  var periods = QA_buildReportPeriods_(periodPointsByEmailByPeriod, currentPeriodIndex, rosterQa, managerNorm);
  var parsedTrainer = QA_parseTrainerProdRows_(ss, rosterQa);
  for (var pi = 0; pi < periods.length; pi++) {
    var pIdx = periods[pi].periodIndex;
    var tb = QA_getTrainerBoundsForReportPeriod_(pIdx, QA_TZ);
    var tr = QA_accumulateTrainerFromParsed_(parsedTrainer, tb);
    periods[pi].trainerPeriodByEmail = tr.periodTotalsByEmail;
    periods[pi].trainerMultiPeriodByEmail = tr.multiPeriodPointsByEmail;
  }
  return { allEmails: allEmails, periods: periods };
}

function QA_buildReportPeriods_(periodPointsByEmailByPeriod, currentPeriodIndex, rosterQaSet, managerNorm) {
  var reportIndexes = [];
  for (var off = -2; off <= 0; off++) {
    var p = currentPeriodIndex + off;
    if (p >= 0) reportIndexes.push(p);
  }
  if (reportIndexes.length === 0) reportIndexes.push(0);

  var result = [];
  for (var idx = 0; idx < reportIndexes.length; idx++) {
    var periodIndex = reportIndexes[idx];
    var b = QA_getPeriodBoundsByIndex_(periodIndex, QA_TZ);
    var periodRangeText = Utilities.formatDate(b.periodStart, QA_TZ, 'M/d/yyyy') + ' - ' + Utilities.formatDate(b.periodEnd, QA_TZ, 'M/d/yyyy');
    var numPeriodsForAvg = b.numPeriodsForAvg;
    var periodPointsByEmail = {};
    var multiPeriodPointsByEmail = {};
    var multiStart = Math.max(0, periodIndex - QA_MULTIPLIER_NUM_PERIODS + 1);
    var nAvg = periodIndex - multiStart + 1;
    if (nAvg < 1) nAvg = 1;

    var emails = new Set();
    managerNorm.forEach(function(e) { emails.add(e); });
    rosterQaSet.forEach(function(e) { emails.add(e); });
    emails.forEach(function(email) {
      var byP = periodPointsByEmailByPeriod[email] || {};
      var periodPts = byP[periodIndex] || 0;
      var multiTotal = 0;
      for (var k = multiStart; k <= periodIndex; k++) multiTotal += (byP[k] || 0);
      periodPointsByEmail[email] = periodPts;
      multiPeriodPointsByEmail[email] = multiTotal;
    });

    result.push({
      periodIndex: periodIndex,
      periodRangeText: periodRangeText,
      periodPointsByEmail: periodPointsByEmail,
      multiPeriodPointsByEmail: multiPeriodPointsByEmail,
      numPeriodsForAvg: numPeriodsForAvg
    });
  }
  return result;
}

/**
 * Look up numeric totals by email when map keys may be raw sheet text or normalized.
 * Trainer maps use QA_normalizeEmail_(Col H); roster/report emails are normalized.
 */
function QA_trainerMapLookup_(map, email) {
  if (!map) return 0;
  var em = QA_normalizeEmail_(email);
  if (!em) return 0;
  if (Object.prototype.hasOwnProperty.call(map, em)) return Number(map[em]) || 0;
  if (Object.prototype.hasOwnProperty.call(map, email)) return Number(map[email]) || 0;
  for (var k in map) {
    if (Object.prototype.hasOwnProperty.call(map, k) && QA_normalizeEmail_(k) === em) {
      return Number(map[k]) || 0;
    }
  }
  return 0;
}

function QA_csvEscape_(val) {
  var s = String(val == null ? '' : val);
  if (s.indexOf('"') !== -1 || s.indexOf(',') !== -1 || s.indexOf('\n') !== -1) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

/**
 * Manager-only: return full QA report as CSV (current + last 2 periods).
 * Columns: Period, Agent email, Total points of period, Qualified incentive, Average points in past 4 periods, Qualified multiplier, Total payout.
 */
function getQAManagerReport() {
  if (!QA_isManager()) return { error: 'Unauthorized' };
  var data = QA_getAllAgentsDataForReport_();
  if (!data) return { error: 'Sheet not found or no data' };

  var header = ['Period', 'Agent email', 'Total points of period', 'Qualified incentive', 'Average points in past 4 periods', 'Qualified multiplier', 'Total payout'];
  var rows = [header.map(function(c) { return QA_csvEscape_(c); }).join(',')];

  for (var p = 0; p < data.periods.length; p++) {
    var period = data.periods[p];
    var n = period.numPeriodsForAvg || QA_MULTIPLIER_NUM_PERIODS;
    for (var i = 0; i < data.allEmails.length; i++) {
      var email = data.allEmails[i];
      var periodPointsQA = QA_trainerMapLookup_(period.periodPointsByEmail, email);
      var trainerPts = QA_trainerMapLookup_(period.trainerPeriodByEmail, email);
      var periodPointsDisplay = periodPointsQA + trainerPts;
      var multiQA = QA_trainerMapLookup_(period.multiPeriodPointsByEmail, email);
      var multiTr = QA_trainerMapLookup_(period.trainerMultiPeriodByEmail, email);
      var multiTotal = multiQA + multiTr;
      var multiPeriodAvg = n > 0 ? multiTotal / n : 0;
      var pay = QA_additionalPay_(periodPointsDisplay);
      var mult = QA_consistencyMultiplier_(multiPeriodAvg);
      var totalPayout = pay.amount * (mult.value || 0);

      rows.push([
        QA_csvEscape_(period.periodRangeText),
        QA_csvEscape_(email),
        QA_csvEscape_(Math.round(periodPointsDisplay)),
        QA_csvEscape_(pay.amount),
        QA_csvEscape_(Math.floor(multiPeriodAvg)),
        QA_csvEscape_(mult.value),
        QA_csvEscape_(Math.round(totalPayout))
      ].join(','));
    }
  }

  return { csv: rows.join('\r\n') };
}

/** Set this, then run `runQAManagerTrainerTraceFromEditor` (no arguments) — output appears in Execution log. */
var QA_TRACE_EMAIL_FOR_EDITOR = 'elijah.aikens@invisible.email';

/**
 * No parameters: uses QA_TRACE_EMAIL_FOR_EDITOR. Logs JSON to Execution log (View → Logs or Ctrl+Enter after run).
 */
function runQAManagerTrainerTraceFromEditor() {
  return getQAManagerTrainerTrace(QA_TRACE_EMAIL_FOR_EDITOR);
}

/**
 * Manager-only: per-period QA vs training totals and date bounds for one email (debug / support).
 * Return value is also written with Logger.log — open View → Logs (or Execution log) after Run.
 * From editor without args, use runQAManagerTrainerTraceFromEditor or pass an email below in Run config.
 */
function getQAManagerTrainerTrace(rawEmail) {
  function logResult_(obj) {
    try {
      Logger.log(JSON.stringify(obj, null, 2));
    } catch (e) {
      Logger.log(String(obj));
    }
  }
  if (!QA_isManager()) {
    var u = { error: 'Unauthorized' };
    logResult_(u);
    return u;
  }
  var email = QA_normalizeEmail_(rawEmail);
  if (!email) {
    var inv = { error: 'Invalid email' };
    logResult_(inv);
    return inv;
  }
  var data = QA_getAllAgentsDataForReport_();
  if (!data) {
    var nd = { error: 'Sheet not found or no data' };
    logResult_(nd);
    return nd;
  }
  var ss = SpreadsheetApp.getActive();
  var rosterQa = QA_getActiveRosterQaEmails_(ss);
  var rosterSet = rosterQa || new Set();
  var parsed = QA_parseTrainerProdRows_(ss, rosterSet);
  var trainerRowCount = 0;
  var trainerSamples = [];
  for (var ri = 0; ri < parsed.length; ri++) {
    if (parsed[ri].email !== email) continue;
    trainerRowCount++;
    if (trainerSamples.length < 8) {
      trainerSamples.push({ dayStr: parsed[ri].dayStr, pts: Math.round(parsed[ri].pts), c: parsed[ri].c, err: parsed[ri].errVal });
    }
  }
  var periodsOut = [];
  for (var p = 0; p < data.periods.length; p++) {
    var period = data.periods[p];
    var pIdx = period.periodIndex;
    var tb = QA_getTrainerBoundsForReportPeriod_(pIdx, QA_TZ);
    var qaP = QA_trainerMapLookup_(period.periodPointsByEmail, email);
    var qaM = QA_trainerMapLookup_(period.multiPeriodPointsByEmail, email);
    var trP = QA_trainerMapLookup_(period.trainerPeriodByEmail, email);
    var trM = QA_trainerMapLookup_(period.trainerMultiPeriodByEmail, email);
    var n = period.numPeriodsForAvg || QA_MULTIPLIER_NUM_PERIODS;
    periodsOut.push({
      periodIndex: pIdx,
      periodRange: period.periodRangeText,
      trainerBounds: {
        periodWeek: tb.periodStartStr + ' .. ' + tb.periodEndStr,
        multiWindow: tb.multiPeriodStartStr + ' .. ' + tb.multiPeriodEndStr
      },
      qaPointsThisPeriod: Math.round(qaP),
      qaMultiSum: Math.round(qaM),
      trainerPointsThisPeriod: Math.round(trP),
      trainerMultiSum: Math.round(trM),
      csvTotalPointsOfPeriod: Math.round(qaP + trP),
      numPeriodsForAvg: n,
      combinedMultiAvg: n > 0 ? Math.floor((qaM + trM) / n) : 0
    });
  }
  var out = {
    email: email,
    onActiveRosterQa: rosterQa ? rosterQa.has(email) : false,
    trainerEligibleRowsInSheet: trainerRowCount,
    trainerRowSamples: trainerSamples,
    periods: periodsOut
  };
  logResult_(out);
  return out;
}

function QA_buildResponse_(email, qaPeriodPoints, multiPeriodQATotal, periodRangeText, periodBreakdown, multiPeriodBreakdown, trainingPeriodPoints, multiPeriodTrainingTotal, numPeriodsForAvg, trainingBatch) {
  periodBreakdown = periodBreakdown || {};
  multiPeriodBreakdown = multiPeriodBreakdown || {};
  trainingPeriodPoints = trainingPeriodPoints || 0;
  multiPeriodTrainingTotal = multiPeriodTrainingTotal || 0;
  qaPeriodPoints = qaPeriodPoints || 0;
  multiPeriodQATotal = multiPeriodQATotal || 0;
  numPeriodsForAvg = numPeriodsForAvg || QA_MULTIPLIER_NUM_PERIODS;
  trainingBatch = trainingBatch || {};

  var combinedBatchPeriod = QA_mergeCombinedBatchPeriod_(periodBreakdown, trainingBatch);

  var totalPeriodPoints = qaPeriodPoints + trainingPeriodPoints;
  var totalMultiPeriodPoints = multiPeriodQATotal + multiPeriodTrainingTotal;
  var multiPeriodAvg = numPeriodsForAvg > 0 ? totalMultiPeriodPoints / numPeriodsForAvg : 0;

  var pay = QA_additionalPay_(totalPeriodPoints);
  var payProgress = QA_progressToNextPayTier_(totalPeriodPoints);
  var mult = QA_consistencyMultiplier_(multiPeriodAvg);
  var multProgress = QA_progressToNextMultiplierTier_(multiPeriodAvg);

  let earnings = 0;
  if (pay.amount > 0) {
    earnings = mult.value ? pay.amount * mult.value : pay.amount;
  }
  const earningsInt = Math.round(earnings);

  return {
    email,
    weekRangeText: periodRangeText,
    weekCompletions: Math.round(totalPeriodPoints),
    qaPointsOnly: Math.round(qaPeriodPoints),
    trainingPointsPeriod: Math.round(trainingPeriodPoints),
    fourWeekAvg: Math.floor(isFinite(multiPeriodAvg) ? multiPeriodAvg : 0),
    additionalPay: { amount: pay.amount, qualifiedText: pay.text },
    payProgress,
    multiplier: { value: mult.value, badgeText: mult.badgeText },
    multProgress,
    earnings: {
      amount: earningsInt,
      text: `You are earning an incremental $${earningsInt} this week!`
    },
    pointsBreakdown: {
      period: periodBreakdown,
      multiPeriod: multiPeriodBreakdown,
      combinedBatchPeriod: combinedBatchPeriod
    }
  };
}

function QA_additionalPay_(points) {
  var sorted = QA_PAY_THRESHOLDS.slice().sort(function(a, b) { return b.points - a.points; });
  for (var i = 0; i < sorted.length; i++) {
    if (points >= sorted[i].points) {
      return { amount: sorted[i].amount, text: 'Qualified for $' + sorted[i].amount + ' Additional Earnings' };
    }
  }
  return { amount: 0, text: 'Qualified for $0 Additional Earning' };
}

function QA_consistencyMultiplier_(avg) {
  var a = Number(avg);
  var safe = isFinite(a) ? a : 0;
  var sorted = QA_MULTIPLIER_THRESHOLDS.slice().sort(function(a, b) { return b.avgMin - a.avgMin; });
  for (var i = 0; i < sorted.length; i++) {
    if (safe >= sorted[i].avgMin) {
      return { value: sorted[i].value, badgeText: sorted[i].badgeLabel };
    }
  }
  return { value: 0, badgeText: 'No badge available right now' };
}

function QA_progressToNextPayTier_(points) {
  var sorted = QA_PAY_THRESHOLDS.slice().sort(function(a, b) { return a.points - b.points; });
  var top = sorted[sorted.length - 1];
  if (points >= top.points) {
    return { barPct: 100, nextTierText: 'Keep up the great work!' };
  }
  var base = 0, next = sorted[0].points, nextAmount = sorted[0].amount;
  for (var i = 0; i < sorted.length; i++) {
    if (points < sorted[i].points) {
      next = sorted[i].points;
      nextAmount = sorted[i].amount;
      base = i > 0 ? sorted[i - 1].points : 0;
      break;
    }
  }
  var span = next - base;
  var progressed = span ? Math.max(0, points - base) : 0;
  var pct = span ? Math.min(100, Math.round((progressed / span) * 100)) : 0;
  var remaining = Math.max(0, next - points);
  return {
    barPct: pct,
    nextTierText: remaining.toLocaleString() + ' points to $' + nextAmount + ' Additional Earnings'
  };
}

function QA_progressToNextMultiplierTier_(avg) {
  var a = Number(avg);
  var safeAvg = isFinite(a) ? a : 0;
  var avgInt = Math.floor(safeAvg);
  var sorted = QA_MULTIPLIER_THRESHOLDS.slice().sort(function(a, b) { return b.avgMin - a.avgMin; });
  if (safeAvg >= sorted[0].avgMin) {
    return { avgInt: avgInt, nextTierText: '0 points till ' + sorted[0].badgeLabel.toLowerCase() };
  }
  var goal = sorted[0].avgMin, label = sorted[0].badgeLabel.toLowerCase();
  for (var i = sorted.length - 1; i >= 0; i--) {
    if (safeAvg < sorted[i].avgMin) {
      goal = sorted[i].avgMin;
      label = sorted[i].badgeLabel.toLowerCase();
      break;
    }
  }
  var remaining = Math.max(0, goal - avgInt);
  return { avgInt: avgInt, nextTierText: remaining.toLocaleString() + ' points till ' + label };
}

function QA_toDateStrInTz_(date, tz) {
  return Utilities.formatDate(date, tz, 'yyyy-MM-dd');
}

/**
 * Periods are 7-day windows starting on QA_PERIOD_START_DATE (in tz).
 * Current period = the period that contains today.
 * Multiplier average:
 * - Forward (within first 4 periods since start): average of periods 0..current (1 to 4 weeks of data).
 * - Backward (once today is past 4 weeks from start): average of ongoing week + past 3 weeks (4 periods).
 */
function QA_getPeriodBounds_(date, tz) {
  var startY = Number(Utilities.formatDate(QA_PERIOD_START_DATE, tz, 'yyyy'));
  var startM = Number(Utilities.formatDate(QA_PERIOD_START_DATE, tz, 'MM')) - 1;
  var startD = Number(Utilities.formatDate(QA_PERIOD_START_DATE, tz, 'dd'));
  var refY = Number(Utilities.formatDate(date, tz, 'yyyy'));
  var refM = Number(Utilities.formatDate(date, tz, 'MM')) - 1;
  var refD = Number(Utilities.formatDate(date, tz, 'dd'));

  var startAtMidnight = new Date(startY, startM, startD, 0, 0, 0, 0);
  var refAtMidnight = new Date(refY, refM, refD, 0, 0, 0, 0);
  var daysSinceStart = Math.round((refAtMidnight - startAtMidnight) / (24 * 60 * 60 * 1000));
  if (daysSinceStart < 0) daysSinceStart = 0;

  var periodIndex = Math.floor(daysSinceStart / QA_PERIOD_DAYS);
  var periodStartDate = QA_addDays_(new Date(startAtMidnight.getTime()), periodIndex * QA_PERIOD_DAYS);
  var periodEndDate = QA_addDays_(new Date(periodStartDate.getTime()), QA_PERIOD_DAYS - 1);
  periodEndDate.setHours(23, 59, 59, 999);

  var periodStartStr = QA_toDateStrInTz_(periodStartDate, tz);
  var periodEndStr = QA_toDateStrInTz_(periodEndDate, tz);

  var multiStartDate;
  var numPeriodsForAvg;

  if (periodIndex < QA_MULTIPLIER_NUM_PERIODS) {
    // Forward: from start through current period (1 to 4 periods)
    multiStartDate = new Date(startAtMidnight.getTime());
    numPeriodsForAvg = periodIndex + 1;
  } else {
    // Backward: ongoing week + past 3 weeks (4 periods)
    multiStartDate = QA_addDays_(new Date(periodStartDate.getTime()), -(QA_MULTIPLIER_NUM_PERIODS - 1) * QA_PERIOD_DAYS);
    numPeriodsForAvg = QA_MULTIPLIER_NUM_PERIODS;
  }

  var multiPeriodStartStr = QA_toDateStrInTz_(multiStartDate, tz);
  return {
    periodStartStr: periodStartStr,
    periodEndStr: periodEndStr,
    multiPeriodStartStr: multiPeriodStartStr,
    multiPeriodEndStr: periodEndStr,
    periodStart: periodStartDate,
    periodEnd: periodEndDate,
    numPeriodsForAvg: numPeriodsForAvg
  };
}

function QA_getPeriodIndex_(date, tz) {
  var startY = Number(Utilities.formatDate(QA_PERIOD_START_DATE, tz, 'yyyy'));
  var startM = Number(Utilities.formatDate(QA_PERIOD_START_DATE, tz, 'MM')) - 1;
  var startD = Number(Utilities.formatDate(QA_PERIOD_START_DATE, tz, 'dd'));
  var refY = Number(Utilities.formatDate(date, tz, 'yyyy'));
  var refM = Number(Utilities.formatDate(date, tz, 'MM')) - 1;
  var refD = Number(Utilities.formatDate(date, tz, 'dd'));
  var startAtMidnight = new Date(startY, startM, startD, 0, 0, 0, 0);
  var refAtMidnight = new Date(refY, refM, refD, 0, 0, 0, 0);
  var daysSinceStart = Math.round((refAtMidnight - startAtMidnight) / (24 * 60 * 60 * 1000));
  if (daysSinceStart < 0) daysSinceStart = 0;
  return Math.floor(daysSinceStart / QA_PERIOD_DAYS);
}

function QA_getPeriodBoundsByIndex_(periodIndex, tz) {
  var startY = Number(Utilities.formatDate(QA_PERIOD_START_DATE, tz, 'yyyy'));
  var startM = Number(Utilities.formatDate(QA_PERIOD_START_DATE, tz, 'MM')) - 1;
  var startD = Number(Utilities.formatDate(QA_PERIOD_START_DATE, tz, 'dd'));
  var startAtMidnight = new Date(startY, startM, startD, 0, 0, 0, 0);
  var periodStartDate = QA_addDays_(new Date(startAtMidnight.getTime()), periodIndex * QA_PERIOD_DAYS);
  var periodEndDate = QA_addDays_(new Date(periodStartDate.getTime()), QA_PERIOD_DAYS - 1);
  periodEndDate.setHours(23, 59, 59, 999);
  var numPeriodsForAvg = periodIndex < QA_MULTIPLIER_NUM_PERIODS ? periodIndex + 1 : QA_MULTIPLIER_NUM_PERIODS;
  return {
    periodStartStr: QA_toDateStrInTz_(periodStartDate, tz),
    periodEndStr: QA_toDateStrInTz_(periodEndDate, tz),
    periodStart: periodStartDate,
    periodEnd: periodEndDate,
    numPeriodsForAvg: numPeriodsForAvg
  };
}

function QA_isDateInRange_(ts, startStr, endStr, tz) {
  var dayStr = QA_toDateStrInTz_(ts, tz);
  return dayStr >= startStr && dayStr <= endStr;
}

function QA_addDays_(dt, days) {
  var d = new Date(dt);
  d.setDate(d.getDate() + days);
  return d;
}
