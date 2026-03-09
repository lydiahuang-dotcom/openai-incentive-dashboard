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

const QA_PILOT_ACCESS_EMAILS = [
  'lydia.huang@invisible.email',
  'michael.hernandez@invisible.email',
  'rana.traboulsi@invisible.email'
];

function QA_normalizeEmail_(str) {
  const s = String(str || '').trim();
  const match = s.match(/\<([^\>]+)\>/);
  const emailOnly = match ? match[1].trim() : s;
  return emailOnly.toLowerCase();
}

/** Web app entry for QA — use this as the deployment "entry point" for the QA app. */
function doGetQA() {
  return HtmlService.createHtmlOutputFromFile('QA_Index')
    .setTitle('QA Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/** Main data endpoint for QA dashboard. Access = QA_PILOT_ACCESS_EMAILS OR anyone in col Q, U, or AM.
 *  When deployed "Execute as: User", the viewer must have at least Viewer access to the spreadsheet, or the script cannot read it. */
function getQADashboardData() {
  var email = '';
  try {
    const raw = Session.getActiveUser().getEmail();
    email = raw ? QA_normalizeEmail_(raw) : '';
    if (!email) { Logger.log('QA getDashboardData: no email'); return { noAccess: true }; }

    const pilotSet = new Set(QA_PILOT_ACCESS_EMAILS.map(function(e) { return QA_normalizeEmail_(e); }).filter(Boolean));

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(QA_SHEET_NAME);
    if (!sh) { Logger.log('QA getDashboardData: sheet not found ' + QA_SHEET_NAME); return { error: 'Sheet "' + QA_SHEET_NAME + '" not found' }; }

    const lastRow = sh.getLastRow();
    if (lastRow < 2) { Logger.log('QA getDashboardData: no data rows'); return { noAccess: true }; }

    // Build set of emails that appear in col Q, U, or AM (any row) — grant access to pilot + anyone in sheet
    const colQ = sh.getRange(2, QA_COL_Q_EMAIL, lastRow, QA_COL_Q_EMAIL).getValues();
    const colU = sh.getRange(2, QA_COL_U_EMAIL, lastRow, QA_COL_U_EMAIL).getValues();
    const colAM = sh.getRange(2, QA_COL_AM_EMAIL, lastRow, QA_COL_AM_EMAIL).getValues();
    const sheetEmails = new Set();
    for (var i = 0; i < colQ.length; i++) {
      var e = QA_normalizeEmail_(colQ[i][0]);
      if (e) sheetEmails.add(e);
      e = QA_normalizeEmail_(colU[i][0]);
      if (e) sheetEmails.add(e);
      e = QA_normalizeEmail_(colAM[i][0]);
      if (e) sheetEmails.add(e);
    }
    var allowed = pilotSet.has(email) || sheetEmails.has(email);
    if (!allowed) { Logger.log('QA getDashboardData: access denied ' + email); return { noAccess: true }; }

    Logger.log('QA getDashboardData: access ok, loading data for ' + email + ', lastRow=' + lastRow);
    var result = QA_getDashboardDataForEmail_(email);
    Logger.log('QA getDashboardData: success for ' + email);
    return result;
  } catch (err) {
    Logger.log('QA getDashboardData ERROR for ' + email + ': ' + (err && err.message ? err.message : err));
    return { error: String(err && err.message ? err.message : err) };
  }
}

function QA_getDashboardDataForEmail_(email) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(QA_SHEET_NAME);
  if (!sh) throw new Error('Sheet "' + QA_SHEET_NAME + '" not found');

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

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { noAccess: true };

  const maxCol = Math.max(QA_COL_P_TS, QA_COL_T_TS, QA_COL_AL_TS, QA_COL_Q_EMAIL, QA_COL_U_EMAIL, QA_COL_AM_EMAIL, QA_COL_L_COMPLET);
  const values = sh.getRange(2, 1, lastRow, maxCol).getValues();

  let periodPoints = 0;
  let multiPeriodTotalPoints = 0;
  function makeBreakdown() {
    return { Eval_M: 0, Eval_Q: 0, RSHF_M: 0, RSHF_Q: 0, RSHF_AJ: 0, HLRM_M: 0, HLRM_Q: 0, HLRM_AJ: 0, Categories_M: 0, Categories_Q: 0, Categories_AJ: 0, Out_R: 0, Out_AK: 0, Out_N_only: 0, Out_N_and_AK: 0 };
  }
  var periodBreakdown = makeBreakdown();
  var multiPeriodBreakdown = makeBreakdown();

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
        if (agentP === email) {
          var ptsM = QA_POINTS_EVAL_M * completion;
          if (QA_isDateInRange_(tsO, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsM; periodBreakdown.Eval_M += ptsM; }
          if (QA_isDateInRange_(tsO, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsM; multiPeriodBreakdown.Eval_M += ptsM; }
        }
      }
      if (qTrue && tsS instanceof Date) {
        agentT = QA_normalizeEmail_(row[QA_COL_U_EMAIL - 1]);
        if (agentT === email) {
          var ptsQ = QA_POINTS_EVAL_Q * completion;
          if (QA_isDateInRange_(tsS, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsQ; periodBreakdown.Eval_Q += ptsQ; }
          if (QA_isDateInRange_(tsS, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsQ; multiPeriodBreakdown.Eval_Q += ptsQ; }
        }
      }
    }

    if (isRSHF) {
      if (mTrue && tsO instanceof Date) {
        agentP = QA_normalizeEmail_(row[QA_COL_Q_EMAIL - 1]);
        if (agentP === email) {
          var ptsM = QA_POINTS_RSHF_M * completion;
          if (QA_isDateInRange_(tsO, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsM; periodBreakdown.RSHF_M += ptsM; }
          if (QA_isDateInRange_(tsO, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsM; multiPeriodBreakdown.RSHF_M += ptsM; }
        }
      }
      if (qTrue && tsS instanceof Date) {
        agentT = QA_normalizeEmail_(row[QA_COL_U_EMAIL - 1]);
        if (agentT === email) {
          var ptsQ = QA_POINTS_RSHF_Q * completion;
          if (QA_isDateInRange_(tsS, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsQ; periodBreakdown.RSHF_Q += ptsQ; }
          if (QA_isDateInRange_(tsS, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsQ; multiPeriodBreakdown.RSHF_Q += ptsQ; }
        }
      }
      if (ajTrue && tsAJ instanceof Date) {
        agentAL = QA_normalizeEmail_(row[QA_COL_AM_EMAIL - 1]);
        if (agentAL === email) {
          var ptsAJ = QA_POINTS_RSHF_AJ * completion;
          if (QA_isDateInRange_(tsAJ, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsAJ; periodBreakdown.RSHF_AJ += ptsAJ; }
          if (QA_isDateInRange_(tsAJ, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsAJ; multiPeriodBreakdown.RSHF_AJ += ptsAJ; }
        }
      }
    }

    if (isHLRM) {
      if (mTrue && tsO instanceof Date) {
        agentP = QA_normalizeEmail_(row[QA_COL_Q_EMAIL - 1]);
        if (agentP === email) {
          var ptsM = QA_POINTS_HLRM_M * completion;
          if (QA_isDateInRange_(tsO, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsM; periodBreakdown.HLRM_M += ptsM; }
          if (QA_isDateInRange_(tsO, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsM; multiPeriodBreakdown.HLRM_M += ptsM; }
        }
      }
      if (qTrue && tsS instanceof Date) {
        agentT = QA_normalizeEmail_(row[QA_COL_U_EMAIL - 1]);
        if (agentT === email) {
          var ptsQ = QA_POINTS_HLRM_Q * completion;
          if (QA_isDateInRange_(tsS, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsQ; periodBreakdown.HLRM_Q += ptsQ; }
          if (QA_isDateInRange_(tsS, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsQ; multiPeriodBreakdown.HLRM_Q += ptsQ; }
        }
      }
      if (ajTrue && tsAJ instanceof Date) {
        agentAL = QA_normalizeEmail_(row[QA_COL_AM_EMAIL - 1]);
        if (agentAL === email) {
          var ptsAJ = QA_POINTS_HLRM_AJ * completion;
          if (QA_isDateInRange_(tsAJ, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsAJ; periodBreakdown.HLRM_AJ += ptsAJ; }
          if (QA_isDateInRange_(tsAJ, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsAJ; multiPeriodBreakdown.HLRM_AJ += ptsAJ; }
        }
      }
    }

    if (isCategories) {
      if (mTrue && tsO instanceof Date) {
        agentP = QA_normalizeEmail_(row[QA_COL_Q_EMAIL - 1]);
        if (agentP === email) {
          var ptsM = QA_POINTS_CATEGORIES_M * completion;
          if (QA_isDateInRange_(tsO, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsM; periodBreakdown.Categories_M += ptsM; }
          if (QA_isDateInRange_(tsO, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsM; multiPeriodBreakdown.Categories_M += ptsM; }
        }
      }
      if (qTrue && tsS instanceof Date) {
        agentT = QA_normalizeEmail_(row[QA_COL_U_EMAIL - 1]);
        if (agentT === email) {
          var ptsQ = QA_POINTS_CATEGORIES_Q * completion;
          if (QA_isDateInRange_(tsS, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsQ; periodBreakdown.Categories_Q += ptsQ; }
          if (QA_isDateInRange_(tsS, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsQ; multiPeriodBreakdown.Categories_Q += ptsQ; }
        }
      }
      if (ajTrue && tsAJ instanceof Date) {
        agentAL = QA_normalizeEmail_(row[QA_COL_AM_EMAIL - 1]);
        if (agentAL === email) {
          var ptsAJ = QA_POINTS_CATEGORIES_AJ * completion;
          if (QA_isDateInRange_(tsAJ, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsAJ; periodBreakdown.Categories_AJ += ptsAJ; }
          if (QA_isDateInRange_(tsAJ, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsAJ; multiPeriodBreakdown.Categories_AJ += ptsAJ; }
        }
      }
    }

    if (isOut) {
      if (qTrue && tsS instanceof Date) {
        agentT = QA_normalizeEmail_(row[QA_COL_U_EMAIL - 1]);
        if (agentT === email) {
          var ptsR = QA_POINTS_OUT_R * completion;
          if (QA_isDateInRange_(tsS, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsR; periodBreakdown.Out_R += ptsR; }
          if (QA_isDateInRange_(tsS, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsR; multiPeriodBreakdown.Out_R += ptsR; }
        }
      }
      if (ajTrue && tsAJ instanceof Date) {
        agentAL = QA_normalizeEmail_(row[QA_COL_AM_EMAIL - 1]);
        if (agentAL === email) {
          var ptsAK = QA_POINTS_OUT_AK * completion;
          if (QA_isDateInRange_(tsAJ, periodStartStr, periodEndStr, QA_TZ)) { periodPoints += ptsAK; periodBreakdown.Out_AK += ptsAK; }
          if (QA_isDateInRange_(tsAJ, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) { multiPeriodTotalPoints += ptsAK; multiPeriodBreakdown.Out_AK += ptsAK; }
        }
      }
      if (mTrue && tsO instanceof Date) {
        agentP = QA_normalizeEmail_(row[QA_COL_Q_EMAIL - 1]);
        if (agentP === email) {
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


  var multiPeriodAvg = numPeriodsForAvg > 0 ? multiPeriodTotalPoints / numPeriodsForAvg : 0;
  return QA_buildResponse_(email, periodPoints, multiPeriodAvg, periodRangeText, periodBreakdown, multiPeriodBreakdown);
}

function QA_buildResponse_(email, periodPoints, multiPeriodAvg, periodRangeText, periodBreakdown, multiPeriodBreakdown) {
  periodBreakdown = periodBreakdown || {};
  multiPeriodBreakdown = multiPeriodBreakdown || {};
  var pay = QA_additionalPay_(periodPoints);
  var payProgress = QA_progressToNextPayTier_(periodPoints);
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
    weekCompletions: Math.round(periodPoints || 0),
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
      multiPeriod: multiPeriodBreakdown
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

function QA_isDateInRange_(ts, startStr, endStr, tz) {
  var dayStr = QA_toDateStrInTz_(ts, tz);
  return dayStr >= startStr && dayStr <= endStr;
}

function QA_addDays_(dt, days) {
  var d = new Date(dt);
  d.setDate(d.getDate() + days);
  return d;
}
