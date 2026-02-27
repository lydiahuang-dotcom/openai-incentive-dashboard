/** QA Dashboard CONFIG — all names prefixed with QA_ to avoid clashes with Code.js in same project **/
const QA_SHEET_NAME = 'Gold_Testing';
const QA_TZ = 'America/New_York';

// Column indices (1-based)
const QA_COL_H_TYPE = 8;    // H — type: Eval, RSHF, hlrm, categories (case-insensitive)
const QA_COL_M_FLAG = 13;    // M — if true, add points to agent in P
const QA_COL_O_TS = 15;      // O — timestamp for M / row
const QA_COL_P_EMAIL = 16;   // P — agent email (points when M true)
const QA_COL_Q_FLAG = 17;    // Q — if true, add points to agent in T
const QA_COL_S_TS = 19;      // S — timestamp for Q
const QA_COL_T_EMAIL = 20;   // T — agent email (points when Q true)
const QA_COL_AJ_FLAG = 36;   // AJ — if true (RSHF/hlrm only), add points to agent in AL
const QA_COL_AK_TS = 37;     // AK — timestamp for AJ (fallback: O)
const QA_COL_AL_EMAIL = 38;  // AL — agent email (points when AJ true)

// Points by type (Col H) and flag. All case-insensitive.
// Eval: M true → 20k to P, Q true → 5k to T
// RSHF: M true → 30k to P, Q true → 15k to T, AJ true → 10k to AL
// hlrm: M true → 60k to P, Q true → 20k to T, AJ true → 60k to AL
// categories: M true → 20k to P, Q true → 10k to T, AJ true → 5k to AL
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

/** Main data endpoint for QA dashboard. Access = only QA_PILOT_ACCESS_EMAILS (sheet P/T/AL does not grant access). */
function getQADashboardData() {
  const raw = Session.getActiveUser().getEmail();
  const email = raw ? QA_normalizeEmail_(raw) : '';
  if (!email) return { noAccess: true };

  const pilotSet = new Set(QA_PILOT_ACCESS_EMAILS.map(function(e) { return QA_normalizeEmail_(e); }).filter(Boolean));
  if (!pilotSet.has(email)) return { noAccess: true };

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(QA_SHEET_NAME);
  if (!sh) throw new Error('Sheet "' + QA_SHEET_NAME + '" not found');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { noAccess: true };

  return QA_getDashboardDataForEmail_(email);
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

  const numRows = lastRow - 1;
  const maxCol = Math.max(QA_COL_O_TS, QA_COL_S_TS, QA_COL_AK_TS, QA_COL_P_EMAIL, QA_COL_T_EMAIL, QA_COL_AL_EMAIL);
  const values = sh.getRange(2, 1, numRows, maxCol).getValues();

  let periodPoints = 0;
  let multiPeriodTotalPoints = 0;

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var typeRaw = String(row[QA_COL_H_TYPE - 1] || '').trim().toLowerCase();
    if (!typeRaw) continue; // Col H empty → 0 points for this row
    var isEval = typeRaw.indexOf('eval') !== -1;
    var isRSHF = typeRaw.indexOf('rshf') !== -1;
    var isHLRM = typeRaw.indexOf('hlrm') !== -1;
    var isCategories = typeRaw.indexOf('categories') !== -1;

    var mTrue = row[QA_COL_M_FLAG - 1] === true || String(row[QA_COL_M_FLAG - 1] || '').trim().toUpperCase() === 'TRUE';
    var qTrue = row[QA_COL_Q_FLAG - 1] === true || String(row[QA_COL_Q_FLAG - 1] || '').trim().toUpperCase() === 'TRUE';
    var ajTrue = row[QA_COL_AJ_FLAG - 1] === true || String(row[QA_COL_AJ_FLAG - 1] || '').trim().toUpperCase() === 'TRUE';

    var tsO = row[QA_COL_O_TS - 1];
    var tsS = row[QA_COL_S_TS - 1];
    var tsAJ = row[QA_COL_AK_TS - 1] instanceof Date ? row[QA_COL_AK_TS - 1] : tsO;

    var agentP, agentT, agentAL;

    if (isEval) {
      if (mTrue && tsO instanceof Date) {
        agentP = QA_normalizeEmail_(row[QA_COL_P_EMAIL - 1]);
        if (agentP === email) {
          if (QA_isDateInRange_(tsO, periodStartStr, periodEndStr, QA_TZ)) periodPoints += QA_POINTS_EVAL_M;
          if (QA_isDateInRange_(tsO, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) multiPeriodTotalPoints += QA_POINTS_EVAL_M;
        }
      }
      if (qTrue && tsS instanceof Date) {
        agentT = QA_normalizeEmail_(row[QA_COL_T_EMAIL - 1]);
        if (agentT === email) {
          if (QA_isDateInRange_(tsS, periodStartStr, periodEndStr, QA_TZ)) periodPoints += QA_POINTS_EVAL_Q;
          if (QA_isDateInRange_(tsS, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) multiPeriodTotalPoints += QA_POINTS_EVAL_Q;
        }
      }
    }

    if (isRSHF) {
      if (mTrue && tsO instanceof Date) {
        agentP = QA_normalizeEmail_(row[QA_COL_P_EMAIL - 1]);
        if (agentP === email) {
          if (QA_isDateInRange_(tsO, periodStartStr, periodEndStr, QA_TZ)) periodPoints += QA_POINTS_RSHF_M;
          if (QA_isDateInRange_(tsO, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) multiPeriodTotalPoints += QA_POINTS_RSHF_M;
        }
      }
      if (qTrue && tsS instanceof Date) {
        agentT = QA_normalizeEmail_(row[QA_COL_T_EMAIL - 1]);
        if (agentT === email) {
          if (QA_isDateInRange_(tsS, periodStartStr, periodEndStr, QA_TZ)) periodPoints += QA_POINTS_RSHF_Q;
          if (QA_isDateInRange_(tsS, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) multiPeriodTotalPoints += QA_POINTS_RSHF_Q;
        }
      }
      if (ajTrue && tsAJ instanceof Date) {
        agentAL = QA_normalizeEmail_(row[QA_COL_AL_EMAIL - 1]);
        if (agentAL === email) {
          if (QA_isDateInRange_(tsAJ, periodStartStr, periodEndStr, QA_TZ)) periodPoints += QA_POINTS_RSHF_AJ;
          if (QA_isDateInRange_(tsAJ, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) multiPeriodTotalPoints += QA_POINTS_RSHF_AJ;
        }
      }
    }

    if (isHLRM) {
      if (mTrue && tsO instanceof Date) {
        agentP = QA_normalizeEmail_(row[QA_COL_P_EMAIL - 1]);
        if (agentP === email) {
          if (QA_isDateInRange_(tsO, periodStartStr, periodEndStr, QA_TZ)) periodPoints += QA_POINTS_HLRM_M;
          if (QA_isDateInRange_(tsO, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) multiPeriodTotalPoints += QA_POINTS_HLRM_M;
        }
      }
      if (qTrue && tsS instanceof Date) {
        agentT = QA_normalizeEmail_(row[QA_COL_T_EMAIL - 1]);
        if (agentT === email) {
          if (QA_isDateInRange_(tsS, periodStartStr, periodEndStr, QA_TZ)) periodPoints += QA_POINTS_HLRM_Q;
          if (QA_isDateInRange_(tsS, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) multiPeriodTotalPoints += QA_POINTS_HLRM_Q;
        }
      }
      if (ajTrue && tsAJ instanceof Date) {
        agentAL = QA_normalizeEmail_(row[QA_COL_AL_EMAIL - 1]);
        if (agentAL === email) {
          if (QA_isDateInRange_(tsAJ, periodStartStr, periodEndStr, QA_TZ)) periodPoints += QA_POINTS_HLRM_AJ;
          if (QA_isDateInRange_(tsAJ, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) multiPeriodTotalPoints += QA_POINTS_HLRM_AJ;
        }
      }
    }

    if (isCategories) {
      if (mTrue && tsO instanceof Date) {
        agentP = QA_normalizeEmail_(row[QA_COL_P_EMAIL - 1]);
        if (agentP === email) {
          if (QA_isDateInRange_(tsO, periodStartStr, periodEndStr, QA_TZ)) periodPoints += QA_POINTS_CATEGORIES_M;
          if (QA_isDateInRange_(tsO, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) multiPeriodTotalPoints += QA_POINTS_CATEGORIES_M;
        }
      }
      if (qTrue && tsS instanceof Date) {
        agentT = QA_normalizeEmail_(row[QA_COL_T_EMAIL - 1]);
        if (agentT === email) {
          if (QA_isDateInRange_(tsS, periodStartStr, periodEndStr, QA_TZ)) periodPoints += QA_POINTS_CATEGORIES_Q;
          if (QA_isDateInRange_(tsS, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) multiPeriodTotalPoints += QA_POINTS_CATEGORIES_Q;
        }
      }
      if (ajTrue && tsAJ instanceof Date) {
        agentAL = QA_normalizeEmail_(row[QA_COL_AL_EMAIL - 1]);
        if (agentAL === email) {
          if (QA_isDateInRange_(tsAJ, periodStartStr, periodEndStr, QA_TZ)) periodPoints += QA_POINTS_CATEGORIES_AJ;
          if (QA_isDateInRange_(tsAJ, multiPeriodStartStr, multiPeriodEndStr, QA_TZ)) multiPeriodTotalPoints += QA_POINTS_CATEGORIES_AJ;
        }
      }
    }
  }


  var multiPeriodAvg = numPeriodsForAvg > 0 ? multiPeriodTotalPoints / numPeriodsForAvg : 0;
  return QA_buildResponse_(email, periodPoints, multiPeriodAvg, periodRangeText);
}

function QA_buildResponse_(email, periodPoints, multiPeriodAvg, periodRangeText) {
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
