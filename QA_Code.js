/** QA Dashboard CONFIG — all names prefixed with QA_ to avoid clashes with Code.js in same project **/
const QA_SHEET_NAME = 'Gold_Testing';
const QA_TZ = 'America/New_York';

// Column indices (1-based)
const QA_COL_M_FLAG = 13;   // M — scenario 1 qualified
const QA_COL_O_TS = 15;     // O — scenario 1 timestamp
const QA_COL_P_EMAIL = 16;  // P — scenario 1 agent email
const QA_COL_Q_FLAG = 17;   // Q — scenario 2 qualified
const QA_COL_S_TS = 19;     // S — scenario 2 timestamp
const QA_COL_T_EMAIL = 20;  // T — scenario 2 agent email

const QA_POINTS_SCENARIO1 = 1500;
const QA_POINTS_SCENARIO2 = 1000;

const QA_COMPETITION_START_DATE = new Date(2026, 0, 31);
const QA_PERIOD_DAYS = 9;
const QA_MULTIPLIER_NUM_PERIODS = 4;

const QA_PAY_THRESHOLDS = [
  { points: 120000, amount: 75 },
  { points: 155000, amount: 90 },
  { points: 190000, amount: 100 }
];

const QA_MULTIPLIER_THRESHOLDS = [
  { avgMin: 120000, value: 1, badgeLabel: '1x Multiplier Active' },
  { avgMin: 155000, value: 1.1, badgeLabel: '1.1x Multiplier Active' },
  { avgMin: 190000, value: 1.25, badgeLabel: '1.25x Multiplier Active' }
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

/** Main data endpoint for QA dashboard. Access = pilot list OR email in Col P or Col T. */
function getQADashboardData() {
  const raw = Session.getActiveUser().getEmail();
  const email = raw ? QA_normalizeEmail_(raw) : '';
  if (!email) return { noAccess: true };

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(QA_SHEET_NAME);
  if (!sh) throw new Error('Sheet "' + QA_SHEET_NAME + '" not found');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { noAccess: true };

  const pilotSet = new Set(QA_PILOT_ACCESS_EMAILS.map(function(e) { return QA_normalizeEmail_(e); }).filter(Boolean));
  if (pilotSet.has(email)) return QA_getDashboardDataForEmail_(email);

  const numRows = lastRow - 1;
  const maxCol = Math.max(QA_COL_P_EMAIL, QA_COL_T_EMAIL);
  const values = sh.getRange(2, 1, numRows, maxCol).getValues();
  const sheetEmails = new Set();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const pEmail = QA_normalizeEmail_(row[QA_COL_P_EMAIL - 1]);
    const tEmail = QA_normalizeEmail_(row[QA_COL_T_EMAIL - 1]);
    if (pEmail) sheetEmails.add(pEmail);
    if (tEmail) sheetEmails.add(tEmail);
  }
  if (!sheetEmails.has(email)) return { noAccess: true };

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
  const periodRangeText =
    Utilities.formatDate(periodStart, QA_TZ, 'M/d/yyyy') + ' - ' + Utilities.formatDate(periodEnd, QA_TZ, 'M/d/yyyy');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { noAccess: true };

  const numRows = lastRow - 1;
  const maxCol = Math.max(QA_COL_O_TS, QA_COL_S_TS, QA_COL_P_EMAIL, QA_COL_T_EMAIL);
  const values = sh.getRange(2, 1, numRows, maxCol).getValues();

  const numPeriods = QA_MULTIPLIER_NUM_PERIODS;
  let periodPoints = 0;
  let multiPeriodTotalPoints = 0;

  for (var i = 0; i < values.length; i++) {
    var row = values[i];

    var mTrue = row[QA_COL_M_FLAG - 1] === true || String(row[QA_COL_M_FLAG - 1] || '').trim().toUpperCase() === 'TRUE';
    if (mTrue) {
      var ts1 = row[QA_COL_O_TS - 1];
      if (ts1 instanceof Date) {
        var inCurrent = QA_isDateInRange_(ts1, periodStartStr, periodEndStr, QA_TZ);
        var inMulti = QA_isDateInRange_(ts1, multiPeriodStartStr, multiPeriodEndStr, QA_TZ);
        var agentP = QA_normalizeEmail_(row[QA_COL_P_EMAIL - 1]);
        if (agentP === email) {
          if (inCurrent) periodPoints += QA_POINTS_SCENARIO1;
          if (inMulti) multiPeriodTotalPoints += QA_POINTS_SCENARIO1;
        }
      }
    }

    var qTrue = row[QA_COL_Q_FLAG - 1] === true || String(row[QA_COL_Q_FLAG - 1] || '').trim().toUpperCase() === 'TRUE';
    if (qTrue) {
      var ts2 = row[QA_COL_S_TS - 1];
      if (ts2 instanceof Date) {
        inCurrent = QA_isDateInRange_(ts2, periodStartStr, periodEndStr, QA_TZ);
        inMulti = QA_isDateInRange_(ts2, multiPeriodStartStr, multiPeriodEndStr, QA_TZ);
        var agentT = QA_normalizeEmail_(row[QA_COL_T_EMAIL - 1]);
        if (agentT === email) {
          if (inCurrent) periodPoints += QA_POINTS_SCENARIO2;
          if (inMulti) multiPeriodTotalPoints += QA_POINTS_SCENARIO2;
        }
      }
    }
  }

  var multiPeriodAvg = multiPeriodTotalPoints / numPeriods;
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

function QA_getPeriodBounds_(date, tz) {
  var startY = Number(Utilities.formatDate(QA_COMPETITION_START_DATE, tz, 'yyyy'));
  var startM = Number(Utilities.formatDate(QA_COMPETITION_START_DATE, tz, 'MM')) - 1;
  var startD = Number(Utilities.formatDate(QA_COMPETITION_START_DATE, tz, 'dd'));
  var refY = Number(Utilities.formatDate(date, tz, 'yyyy'));
  var refM = Number(Utilities.formatDate(date, tz, 'MM')) - 1;
  var refD = Number(Utilities.formatDate(date, tz, 'dd'));
  var startAtNoon = new Date(startY, startM, startD, 12, 0, 0, 0);
  var refAtNoon = new Date(refY, refM, refD, 12, 0, 0, 0);
  var daysSinceStart = Math.floor((refAtNoon - startAtNoon) / (24 * 60 * 60 * 1000));
  var periodIndex = Math.max(0, Math.floor(daysSinceStart / QA_PERIOD_DAYS));
  var periodStartDate = QA_addDays_(new Date(startAtNoon.getTime()), periodIndex * QA_PERIOD_DAYS);
  var periodEndDate = QA_addDays_(new Date(startAtNoon.getTime()), periodIndex * QA_PERIOD_DAYS + (QA_PERIOD_DAYS - 1));
  var periodStartStr = QA_toDateStrInTz_(periodStartDate, tz);
  var periodEndStr = QA_toDateStrInTz_(periodEndDate, tz);
  var multiStartDate = QA_addDays_(new Date(startAtNoon.getTime()), (periodIndex - (QA_MULTIPLIER_NUM_PERIODS - 1)) * QA_PERIOD_DAYS);
  var multiPeriodStartStr = QA_toDateStrInTz_(multiStartDate, tz);
  return {
    periodStartStr: periodStartStr,
    periodEndStr: periodEndStr,
    multiPeriodStartStr: multiPeriodStartStr,
    multiPeriodEndStr: periodEndStr,
    periodStart: periodStartDate,
    periodEnd: periodEndDate
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
