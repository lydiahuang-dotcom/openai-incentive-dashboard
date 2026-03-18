/** CONFIG **/
const SHEET_NAME = '[Image] Prod';
const TZ = 'America/New_York';


// Column indices (1-based)
const COL_TIMESTAMP = 7; // G
const COL_EMAIL     = 8; // H
const COL_H_TYPE    = 9; // I — type: RSHF, Evals, HLRM, categories (case-insensitive)
const COL_STATUS    = 10; // J
const COL_COMPLET   = 12; // L — value (multiplier for points)
const COL_ERROR     = 36; // AJ — error count for quality eligibility
const ALLOWED_STATUSES = ['Task Submitted', 'Revised'];

// Points per row = basePoints * value in col L (case-insensitive match on col I)
const POINTS_RSHF = 75000;
const POINTS_EVALS = 20000;
const POINTS_HLRM = 50000;
const POINTS_CATEGORIES = 30000;
const POINTS_MULTI_OUT = 35000;


// ——— Competition parameters (adjust these when rules change) ———
/** First date of the first competition period (midnight in TZ). Format: year, month-1, day */
const COMPETITION_START_DATE = new Date(2026, 0, 31);  // 1/31 → period 0: 1/31–2/8, period 1: 2/9–2/17, … 
/** Length of one period in days (e.g. 9 = 9-day window; on day 10 a new period starts). */
const PERIOD_DAYS = 9;
/** Number of periods used for consistency multiplier (trailing + current). */
const MULTIPLIER_NUM_PERIODS = 4;


/**
* Pay tiers: current 9-day period total points → dollar amount.
* 0.9M → $50, 1.8M → $70, 2.7M → $80
*/
const PAY_THRESHOLDS = [
  { points: 900000, amount: 50 },
  { points: 1800000, amount: 70 },
  { points: 2700000, amount: 80 }
];

/**
* Multiplier tiers: (trailing 3 + current) period average points → multiplier value.
* 0.9M → 1.0, 1.8M → 1.1, 2.7M → 1.25
*/
const MULTIPLIER_THRESHOLDS = [
  { avgMin: 900000, value: 1, badgeLabel: '1x Multiplier Active' },
  { avgMin: 1800000, value: 1.1, badgeLabel: '1.1x Multiplier Active' },
  { avgMin: 2700000, value: 1.25, badgeLabel: '1.25x Multiplier Active' }
];


// ——— Pilot launch: only these emails can access the dashboard. Remove after pilot and uncomment sheet-based access below. ———
const PILOT_ACCESS_EMAILS = [
 'lydia.huang@invisible.email',
 'michael.hernandez@invisible.email',
 'rana.traboulsi@invisible.email'
];

/** Managers who can export the full report (CSV). */
const MANAGER_EMAILS = [
  'lydia.huang@invisible.email',
  'michael.hernandez@invisible.email',
  'rana.traboulsi@invisible.email'
];


/** Normalize email for comparison: strip "Name <email>" to just the address, trim, lowercase. */
function normalizeEmail_(str) {
 const s = String(str || '').trim();
 const match = s.match(/\<([^\>]+)\>/);
 const emailOnly = match ? match[1].trim() : s;
 return emailOnly.toLowerCase();
}


/** Web app entry. Use ?app=qa in the URL for the QA dashboard; otherwise Trainer dashboard. */
function doGet(e) {
  e = e || {};
  var params = e.parameter || {};
  var app = params.app;
  if (Array.isArray(app)) app = app[0];
  app = String(app || '').trim().toLowerCase();
  if (app === 'qa') {
    return HtmlService.createHtmlOutputFromFile('QA_Index')
      .setTitle('QA Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Trainer Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}


/** True if the current user (Session.getActiveUser().getEmail()) is in MANAGER_EMAILS. Use when deployed "Execute as: User". */
function isManager() {
  try {
    const raw = Session.getActiveUser().getEmail();
    const email = raw ? normalizeEmail_(raw) : '';
    if (!email) return false;
    const list = MANAGER_EMAILS.map(e => normalizeEmail_(e)).filter(Boolean);
    return list.indexOf(email) !== -1;
  } catch (e) {
    return false;
  }
}

/** Main data endpoint for the UI. When deployed "Execute as: User", the viewer must have at least Viewer access to the spreadsheet. */
function getDashboardData() {
  var email = '';
  try {
    const raw = Session.getActiveUser().getEmail();
    email = raw ? normalizeEmail_(raw) : '';
    if (!email) { Logger.log('Trainer getDashboardData: no email'); return { noAccess: true }; }

    Logger.log('Trainer getDashboardData: loading for ' + email);
    var result = getDashboardDataForEmail_(email);
    Logger.log('Trainer getDashboardData: success for ' + email);
    return result;
  } catch (err) {
    Logger.log('Trainer getDashboardData ERROR for ' + email + ': ' + (err && err.message ? err.message : err));
    return { error: String(err && err.message ? err.message : err) };
  }
}

/** Load sheet and build dashboard response. Access: pilot list OR any email appearing in col H. */
function getDashboardDataForEmail_(email) {
  const pilotAllowed = PILOT_ACCESS_EMAILS.map(e => normalizeEmail_(e)).filter(Boolean);

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return { error: 'Sheet "' + SHEET_NAME + '" not found' };

  const now = new Date();
  const bounds = getPeriodBounds_(now, TZ);
  const {
    periodStartStr,
    periodEndStr,
    multiPeriodStartStr,
    multiPeriodEndStr,
    periodStart,
    periodEnd
  } = bounds;
  const periodRangeText =
    `${Utilities.formatDate(periodStart, TZ, 'M/d/yyyy')} - ${Utilities.formatDate(periodEnd, TZ, 'M/d/yyyy')}`;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return { noAccess: true };
  }

  // Pull only the 6 columns we use: G(7), H(8), I(9), J(10), L(12), AJ(36) — two ranges
  const valuesMain = sh.getRange(2, 7, lastRow, 12).getValues();   // cols G–L → indices 0–5
  const valuesError = sh.getRange(2, 36, lastRow, 36).getValues(); // col AJ

  // Emails that appear in col H (any row) — grant access to pilot list + anyone in sheet
  const sheetEmails = new Set();
  for (let i = 0; i < valuesMain.length; i++) {
    const e = String(valuesMain[i][1] || '').trim().toLowerCase();
    if (e) sheetEmails.add(e);
  }
  const allowed = pilotAllowed.includes(email) || sheetEmails.has(email);
  if (!allowed) return { noAccess: true };

  const numPeriods = MULTIPLIER_NUM_PERIODS;
  // Per-email valid pool: only tasks under the error ceiling count toward points and eligibility
  const validPoolC = Object.create(null);
  const validPoolE = Object.create(null);
  const periodTotalsByEmail = Object.create(null);
  const multiPeriodPointsByEmail = Object.create(null);

  for (let i = 0; i < valuesMain.length; i++) {
    const main = valuesMain[i];
    const rowEmail = String(main[1] || '').trim().toLowerCase(); // col H
    if (!rowEmail) continue;

    const ts = main[0]; // col G
    if (!(ts instanceof Date)) continue;

    const st = String(main[3] || '').trim(); // col J
    if (!ALLOWED_STATUSES || !ALLOWED_STATUSES.length || !ALLOWED_STATUSES.includes(st)) continue;

    const kVal = Number(main[5] || 0); // col L
    if (!isFinite(kVal)) continue;

    const hRaw = String(main[2] || '').trim().toLowerCase(); // col I
    let basePoints = 0;
    if (hRaw.indexOf('rshf') !== -1) basePoints = POINTS_RSHF;
    else if (hRaw.indexOf('evals') !== -1 || hRaw.indexOf('eval') !== -1 ||hRaw.indexOf('ema') !== -1 || (hRaw.indexOf('hrm') !== -1 && hRaw.indexOf('hlrm') === -1)) basePoints = POINTS_EVALS;
    else if (hRaw.indexOf('hlrm') !== -1) basePoints = POINTS_HLRM;
    else if (hRaw.indexOf('categories') !== -1) basePoints = POINTS_CATEGORIES;
    else if (hRaw.indexOf('multi-out') !== -1) basePoints = POINTS_MULTI_OUT;
    if (basePoints === 0) continue;

    const pts = basePoints * kVal;
    const c = kVal;
    const errVal = isFinite(Number(valuesError[i][0])) ? Number(valuesError[i][0]) : 0; // col AJ

    // Format date once per row, then reuse for both range checks (avoids 2x Utilities.formatDate)
    const dayStr = toDateStrInTz_(ts, TZ);
    const inCurrentPeriod = dayStr >= periodStartStr && dayStr <= periodEndStr;
    const inMultiPeriod = dayStr >= multiPeriodStartStr && dayStr <= multiPeriodEndStr;

    if (!inMultiPeriod) continue;

    const curC = validPoolC[rowEmail] || 0;
    const curE = validPoolE[rowEmail] || 0;
    const newC = curC + c;
    const newE = curE + errVal;
    const underCeiling = (newC > 8 && newE <= 0.125 * newC) || (newC <= 8 && newE <= 1);

    if (underCeiling) {
      validPoolC[rowEmail] = newC;
      validPoolE[rowEmail] = newE;
      multiPeriodPointsByEmail[rowEmail] = (multiPeriodPointsByEmail[rowEmail] || 0) + pts;
      if (inCurrentPeriod) {
        periodTotalsByEmail[rowEmail] = (periodTotalsByEmail[rowEmail] || 0) + pts;
      }
    }
  }

  const periodPoints = periodTotalsByEmail[email] || 0;
  const multiPeriodTotalPoints = multiPeriodPointsByEmail[email] || 0;
  const multiPeriodAvg = multiPeriodTotalPoints / numPeriods;
  const vC = validPoolC[email] || 0;
  const vE = validPoolE[email] || 0;
  const qualityEligible = (vC > 8 && vE <= 0.125 * vC) || (vC <= 8 && vE <= 1);

  const fromSheet = Object.keys(periodTotalsByEmail).map(e => ({
    email: e,
    weeklyCompletions: periodTotalsByEmail[e] || 0
  }));
  const allAccessEmails = [...new Set([...pilotAllowed, ...sheetEmails])];
  const accessOnlyNoData = allAccessEmails
    .filter(e => !(e in periodTotalsByEmail))
    .map(e => ({ email: e, weeklyCompletions: 0 }));
  const allPeriodRows = fromSheet.concat(accessOnlyNoData);

  const leaderboard = buildLeaderboard_(allPeriodRows, email);

  return buildResponse_(email, periodPoints, multiPeriodAvg, periodRangeText, leaderboard, qualityEligible);
}


/**
 * Load sheet once and build period/multi-period aggregates for current + last 2 periods (manager report).
 * Returns { allEmails, periods: [{ periodRangeText, periodTotalsByEmail, multiPeriodPointsByEmail, validPoolC, validPoolE }] }.
 */
function getAllAgentsDataForReport_() {
  const pilotAllowed = PILOT_ACCESS_EMAILS.map(e => normalizeEmail_(e)).filter(Boolean);

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return null;

  const now = new Date();
  const currentPeriodIndex = getPeriodIndex_(now, TZ);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    const periods = buildReportPeriodsTrainer_([], currentPeriodIndex, pilotAllowed);
    return { allEmails: [...pilotAllowed], periods };
  }

  const valuesMain = sh.getRange(2, 7, lastRow, 12).getValues();
  const valuesError = sh.getRange(2, 36, lastRow, 36).getValues();

  const sheetEmails = new Set();
  const parsedRows = [];

  for (let i = 0; i < valuesMain.length; i++) {
    const main = valuesMain[i];
    const rowEmail = String(main[1] || '').trim().toLowerCase();
    if (!rowEmail) continue;

    const ts = main[0];
    if (!(ts instanceof Date)) continue;

    const st = String(main[3] || '').trim();
    if (!ALLOWED_STATUSES || !ALLOWED_STATUSES.length || !ALLOWED_STATUSES.includes(st)) continue;

    const kVal = Number(main[5] || 0);
    if (!isFinite(kVal)) continue;

    const hRaw = String(main[2] || '').trim().toLowerCase();
    let basePoints = 0;
    if (hRaw.indexOf('rshf') !== -1) basePoints = POINTS_RSHF;
    else if (hRaw.indexOf('evals') !== -1 || hRaw.indexOf('eval') !== -1 || hRaw.indexOf('ema') !== -1 || (hRaw.indexOf('hrm') !== -1 && hRaw.indexOf('hlrm') === -1)) basePoints = POINTS_EVALS;
    else if (hRaw.indexOf('hlrm') !== -1) basePoints = POINTS_HLRM;
    else if (hRaw.indexOf('categories') !== -1) basePoints = POINTS_CATEGORIES;
    else if (hRaw.indexOf('multi-out') !== -1) basePoints = POINTS_MULTI_OUT;
    if (basePoints === 0) continue;

    const pts = basePoints * kVal;
    const c = kVal;
    const errVal = isFinite(Number(valuesError[i][0])) ? Number(valuesError[i][0]) : 0;
    const dayStr = toDateStrInTz_(ts, TZ);

    sheetEmails.add(rowEmail);
    parsedRows.push({ email: rowEmail, dayStr, ts: ts.getTime(), pts, c, errVal });
  }

  const allEmails = [...new Set([...pilotAllowed, ...sheetEmails])];
  const periods = buildReportPeriodsTrainer_(parsedRows, currentPeriodIndex, pilotAllowed);
  return { allEmails, periods };
}

/** Build report data for current and last 2 periods (Trainer). */
function buildReportPeriodsTrainer_(parsedRows, currentPeriodIndex, pilotAllowed) {
  const numPeriods = MULTIPLIER_NUM_PERIODS;
  const reportIndexes = [];
  for (let off = -2; off <= 0; off++) {
    const p = currentPeriodIndex + off;
    if (p >= 0) reportIndexes.push(p);
  }
  if (reportIndexes.length === 0) reportIndexes.push(0);

  const result = [];
  for (let idx = 0; idx < reportIndexes.length; idx++) {
    const periodIndex = reportIndexes[idx];
    const b = getPeriodBoundsByIndex_(periodIndex, TZ);
    const periodStartStr = b.periodStartStr;
    const periodEndStr = b.periodEndStr;
    const multiPeriodStartStr = b.multiPeriodStartStr;
    const multiPeriodEndStr = b.multiPeriodEndStr;

    const inMulti = parsedRows.filter(function(r) {
      return r.dayStr >= multiPeriodStartStr && r.dayStr <= multiPeriodEndStr;
    });
    // Process in same order as dashboard (sheet order), not by timestamp, so report matches individual view
    // inMulti preserves parsedRows order (sheet order)

    const validPoolC = Object.create(null);
    const validPoolE = Object.create(null);
    const periodTotalsByEmail = Object.create(null);
    const multiPeriodPointsByEmail = Object.create(null);

    for (let i = 0; i < inMulti.length; i++) {
      const r = inMulti[i];
      const curC = validPoolC[r.email] || 0;
      const curE = validPoolE[r.email] || 0;
      const newC = curC + r.c;
      const newE = curE + r.errVal;
      const underCeiling = (newC > 8 && newE <= 0.125 * newC) || (newC <= 8 && newE <= 1);

      if (underCeiling) {
        validPoolC[r.email] = newC;
        validPoolE[r.email] = newE;
        multiPeriodPointsByEmail[r.email] = (multiPeriodPointsByEmail[r.email] || 0) + r.pts;
        if (r.dayStr >= periodStartStr && r.dayStr <= periodEndStr) {
          periodTotalsByEmail[r.email] = (periodTotalsByEmail[r.email] || 0) + r.pts;
        }
      }
    }

    const periodRangeText = Utilities.formatDate(b.periodStart, TZ, 'M/d/yyyy') + ' - ' + Utilities.formatDate(b.periodEnd, TZ, 'M/d/yyyy');
    result.push({
      periodRangeText,
      periodTotalsByEmail,
      multiPeriodPointsByEmail,
      validPoolC,
      validPoolE
    });
  }
  return result;
}


/** Escape a CSV field (wrap in quotes and double internal quotes if needed). */
function csvEscape_(val) {
  const s = String(val == null ? '' : val);
  if (s.indexOf('"') !== -1 || s.indexOf(',') !== -1 || s.indexOf('\n') !== -1) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}


/**
 * Manager-only: return full report as CSV string (current + last 2 periods).
 * Columns: Period, Agent email, Total points of period, Qualified incentive, Average points in past 4 periods, Qualified multiplier, Total payout.
 * Total payout = Qualified incentive * Qualified multiplier when quality-eligible; otherwise 0.
 */
function getManagerReport() {
  if (!isManager()) {
    return { error: 'Unauthorized' };
  }
  const data = getAllAgentsDataForReport_();
  if (!data) return { error: 'Sheet not found or no data' };

  const numPeriods = MULTIPLIER_NUM_PERIODS;
  const header = ['Period', 'Agent email', 'Total points of period', 'Qualified incentive', 'Average points in past 4 periods', 'Qualified multiplier', 'Total payout'];
  const rows = [header.map(c => csvEscape_(c)).join(',')];

  for (let p = 0; p < data.periods.length; p++) {
    const period = data.periods[p];
    for (let i = 0; i < data.allEmails.length; i++) {
      const email = data.allEmails[i];
      const periodPoints = period.periodTotalsByEmail[email] || 0;
      const multiTotal = period.multiPeriodPointsByEmail[email] || 0;
      const multiPeriodAvg = numPeriods > 0 ? multiTotal / numPeriods : 0;
      const pay = additionalPay_(periodPoints);
      const mult = consistencyMultiplier_(multiPeriodAvg);
      // Quality is applied per task (only qualifying tasks count toward points); do not zero out the person's payout.
      const totalPayout = pay.amount * (mult.value || 0);

      rows.push([
        csvEscape_(period.periodRangeText),
        csvEscape_(email),
        csvEscape_(Math.round(periodPoints)),
        csvEscape_(pay.amount),
        csvEscape_(Math.floor(multiPeriodAvg)),
        csvEscape_(mult.value),
        csvEscape_(Math.round(totalPayout))
      ].join(','));
    }
  }

  return { csv: rows.join('\r\n') };
}


/** Build all UI-facing values (tiers, progress, text). weekCompletions/fourWeekAvg hold points for front-end.
 * Quality is applied per task (only qualifying tasks count); we do not disqualify the whole person. */
function buildResponse_(email, periodPoints, multiPeriodAvg, periodRangeText, leaderboard, qualityEligible) {
  const pay = additionalPay_(periodPoints);
  const payProgress = progressToNextPayTier_(periodPoints);
  const mult = consistencyMultiplier_(multiPeriodAvg);
  const multProgress = progressToNextMultiplierTier_(multiPeriodAvg);

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
    earnings: { amount: earningsInt, text: `You are earning an incremental $${earningsInt} in this period!` },
    leaderboard
  };
}


/** Tiers: points → amount (PAY_THRESHOLDS) **/
function additionalPay_(points) {
  const sorted = PAY_THRESHOLDS.slice().sort((a, b) => b.points - a.points);
  for (let i = 0; i < sorted.length; i++) {
    if (points >= sorted[i].points) {
      return { amount: sorted[i].amount, text: `Qualified for $${sorted[i].amount} Additional Earnings` };
    }
  }
  return { amount: 0, text: 'Qualified for $0 Additional Earning' };
}


function consistencyMultiplier_(avg) {
  const a = Number(avg);
  const safe = isFinite(a) ? a : 0;
  const sorted = MULTIPLIER_THRESHOLDS.slice().sort((a, b) => b.avgMin - a.avgMin);
  for (let i = 0; i < sorted.length; i++) {
    if (safe >= sorted[i].avgMin) {
      return { value: sorted[i].value, badgeText: sorted[i].badgeLabel };
    }
  }
  return { value: 0, badgeText: 'No badge available right now' };
}


function progressToNextPayTier_(points) {
  const sorted = PAY_THRESHOLDS.slice().sort((a, b) => a.points - b.points);
  const top = sorted[sorted.length - 1];
  if (points >= top.points) {
    return { barPct: 100, nextTierText: 'Keep up the great work!' };
  }
  let base = 0, next = sorted[0].points, nextAmount = sorted[0].amount;
  for (let i = 0; i < sorted.length; i++) {
    if (points < sorted[i].points) {
      next = sorted[i].points;
      nextAmount = sorted[i].amount;
      base = i > 0 ? sorted[i - 1].points : 0;
      break;
    }
  }
  const span = next - base;
  const progressed = span ? Math.max(0, points - base) : 0;
  const pct = span ? Math.min(100, Math.round((progressed / span) * 100)) : 0;
  const remaining = Math.max(0, next - points);
  return {
    barPct: pct,
    nextTierText: `${remaining.toLocaleString()} points to $${nextAmount} Additional Earnings`
  };
}


function progressToNextMultiplierTier_(avg) {
  const a = Number(avg);
  const safeAvg = isFinite(a) ? a : 0;
  const avgInt = Math.floor(safeAvg);
  const sorted = MULTIPLIER_THRESHOLDS.slice().sort((a, b) => b.avgMin - a.avgMin);
  if (safeAvg >= sorted[0].avgMin) {
    return { avgInt, nextTierText: `0 points till ${sorted[0].badgeLabel.toLowerCase()}` };
  }
  let goal = sorted[0].avgMin, label = sorted[0].badgeLabel.toLowerCase();
  for (let i = sorted.length - 1; i >= 0; i--) {
    if (safeAvg < sorted[i].avgMin) {
      goal = sorted[i].avgMin;
      label = sorted[i].badgeLabel.toLowerCase();
      break;
    }
  }
  const remaining = Math.max(0, goal - avgInt);
  return { avgInt, nextTierText: `${remaining.toLocaleString()} points till ${label}` };
}





/** Leaderboard (shared ranks / competition ranking) **/
function buildLeaderboard_(rows, viewerEmail) {
 // Sort by completions desc so agents with 0 are last; tie-break by email
 const sorted = rows
   .slice()
   .sort((a, b) => (b.weeklyCompletions - a.weeklyCompletions) || a.email.localeCompare(b.email));


 // Competition ranking with ties:
 // scores: 100,90,90,80 => ranks: 1,2,2,4
 let prevScore = null;
 let currentRank = 0;


 const ranked = sorted.map((r, idx) => {
   const score = Math.round(r.weeklyCompletions || 0);
   if (prevScore === null || score !== prevScore) {
     currentRank = idx + 1;
     prevScore = score;
   }
   return {
     rank: currentRank,
     email: r.email,
     name: emailToName_(r.email),
     weeklyCompletions: score,
     isMe: r.email === viewerEmail
   };
 });


 const top10 = ranked.slice(0, 10);
 const topLeft = top10.slice(0, 5);
 const topRight = top10.slice(5, 10);


 // ✅ If viewer is in top 10, we don't need YOUR POSITION section
 const viewerInTop10 = top10.some(r => r.isMe);


 // Build neighbor rows only if NOT in top 10
 let yourRows = [];
 if (!viewerInTop10) {
   const meIndex = ranked.findIndex(r => r.email === viewerEmail);


   yourRows = buildNeighborRows_(ranked, meIndex).map(r => ({
     ...r,
     // ✅ Arrow text ONLY for "me"
     arrow: r.isMe ? '↑ Complete more tasks to move up!' : ''
   }));
 }


 return { topLeft, topRight, yourRows, viewerInTop10 };
}




// Neighbor rows by sorted order (prev/me/next)
function buildNeighborRows_(ranked, meIndex) {
 if (!ranked.length) return [];
 if (meIndex < 0) return ranked.slice(0, 3);


 let start = Math.max(0, meIndex - 1);
 let end = Math.min(ranked.length - 1, meIndex + 1);


 while ((end - start + 1) < 3) {
   if (start > 0) start--;
   else if (end < ranked.length - 1) end++;
   else break;
 }
 return ranked.slice(start, end + 1);
}


function emailToName_(email) {
 const user = String(email || '').split('@')[0] || '';
 const parts = user.split('.');
 const first = parts[0] ? cap_(parts[0]) : user;
 const lastInitial = parts[1] ? (parts[1].charAt(0).toUpperCase() + '.') : '';
 return lastInitial ? `${first} ${lastInitial}` : first;
}
function cap_(s) {
 s = String(s || '');
 return s ? s.charAt(0).toUpperCase() + s.slice(1).toLowerCase() : s;
}


/** Calendar day of date in TZ as "yyyy-MM-dd" for consistent comparison. */
function toDateStrInTz_(date, tz) {
 return Utilities.formatDate(date, tz, 'yyyy-MM-dd');
}


/**
* Period bounds from competition start: each period is PERIOD_DAYS long;
* on day (PERIOD_DAYS+1) a new period starts.
* Uses noon when adding days so the calendar day is correct in TZ regardless of server.
*/
function getPeriodBounds_(date, tz) {
 const startY = Number(Utilities.formatDate(COMPETITION_START_DATE, tz, 'yyyy'));
 const startM = Number(Utilities.formatDate(COMPETITION_START_DATE, tz, 'MM')) - 1;
 const startD = Number(Utilities.formatDate(COMPETITION_START_DATE, tz, 'dd'));
 const refY = Number(Utilities.formatDate(date, tz, 'yyyy'));
 const refM = Number(Utilities.formatDate(date, tz, 'MM')) - 1;
 const refD = Number(Utilities.formatDate(date, tz, 'dd'));


 const startAtNoon = new Date(startY, startM, startD, 12, 0, 0, 0);
 const refAtNoon = new Date(refY, refM, refD, 12, 0, 0, 0);
 const daysSinceStart = Math.floor((refAtNoon - startAtNoon) / (24 * 60 * 60 * 1000));
 const periodIndex = Math.max(0, Math.floor(daysSinceStart / PERIOD_DAYS));


 const periodStartDate = addDays_(new Date(startAtNoon.getTime()), periodIndex * PERIOD_DAYS);
 const periodEndDate = addDays_(new Date(startAtNoon.getTime()), periodIndex * PERIOD_DAYS + (PERIOD_DAYS - 1));


 const periodStartStr = toDateStrInTz_(periodStartDate, tz);
 const periodEndStr = toDateStrInTz_(periodEndDate, tz);


 const multiStartDate = addDays_(new Date(startAtNoon.getTime()), (periodIndex - (MULTIPLIER_NUM_PERIODS - 1)) * PERIOD_DAYS);
 const multiPeriodStartStr = toDateStrInTz_(multiStartDate, tz);


 return {
   periodStartStr,
   periodEndStr,
   multiPeriodStartStr,
   multiPeriodEndStr: periodEndStr,
   periodStart: periodStartDate,
   periodEnd: periodEndDate,
   periodIndex
 };
}

/** Return 0-based period index for a date in tz. */
function getPeriodIndex_(date, tz) {
  const startY = Number(Utilities.formatDate(COMPETITION_START_DATE, tz, 'yyyy'));
  const startM = Number(Utilities.formatDate(COMPETITION_START_DATE, tz, 'MM')) - 1;
  const startD = Number(Utilities.formatDate(COMPETITION_START_DATE, tz, 'dd'));
  const refY = Number(Utilities.formatDate(date, tz, 'yyyy'));
  const refM = Number(Utilities.formatDate(date, tz, 'MM')) - 1;
  const refD = Number(Utilities.formatDate(date, tz, 'dd'));
  const startAtNoon = new Date(startY, startM, startD, 12, 0, 0, 0);
  const refAtNoon = new Date(refY, refM, refD, 12, 0, 0, 0);
  const daysSinceStart = Math.floor((refAtNoon - startAtNoon) / (24 * 60 * 60 * 1000));
  return Math.max(0, Math.floor(daysSinceStart / PERIOD_DAYS));
}

/** Bounds for a given 0-based period index. */
function getPeriodBoundsByIndex_(periodIndex, tz) {
  const startY = Number(Utilities.formatDate(COMPETITION_START_DATE, tz, 'yyyy'));
  const startM = Number(Utilities.formatDate(COMPETITION_START_DATE, tz, 'MM')) - 1;
  const startD = Number(Utilities.formatDate(COMPETITION_START_DATE, tz, 'dd'));
  const startAtNoon = new Date(startY, startM, startD, 12, 0, 0, 0);
  const periodStartDate = addDays_(new Date(startAtNoon.getTime()), periodIndex * PERIOD_DAYS);
  const periodEndDate = addDays_(new Date(startAtNoon.getTime()), periodIndex * PERIOD_DAYS + (PERIOD_DAYS - 1));
  const multiStartDate = addDays_(new Date(startAtNoon.getTime()), (periodIndex - (MULTIPLIER_NUM_PERIODS - 1)) * PERIOD_DAYS);
  return {
    periodStartStr: toDateStrInTz_(periodStartDate, tz),
    periodEndStr: toDateStrInTz_(periodEndDate, tz),
    periodStart: periodStartDate,
    periodEnd: periodEndDate,
    multiPeriodStartStr: toDateStrInTz_(multiStartDate, tz),
    multiPeriodEndStr: toDateStrInTz_(periodEndDate, tz)
  };
}


/** True if the calendar day of ts in tz is between startStr and endStr (inclusive). */
function isDateInRange_(ts, startStr, endStr, tz) {
 const dayStr = toDateStrInTz_(ts, tz);
 return dayStr >= startStr && dayStr <= endStr;
}


function addDays_(dt, days) {
 const d = new Date(dt);
 d.setDate(d.getDate() + days);
 return d;
}



