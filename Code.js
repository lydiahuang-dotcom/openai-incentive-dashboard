/** CONFIG **/
const SHEET_NAME = 'Gold';
const TZ = 'America/New_York';


// Column indices (1-based)
const COL_TIMESTAMP = 6; // F
const COL_EMAIL     = 7; // G
const COL_STATUS    = 9; // I
const COL_COMPLET   = 11; // K
const COL_ERROR     = 35; // AI — error count for quality eligibility
const ALLOWED_STATUSES = ['Task Submitted'];


// ——— Competition parameters (adjust these when rules change) ———
/** First date of the first competition period (midnight in TZ). Format: year, month-1, day */
const COMPETITION_START_DATE = new Date(2026, 0, 31);  // 1/31 → period 0: 1/31–2/8, period 1: 2/9–2/17, … 
/** Length of one period in days (e.g. 9 = 9-day window; on day 10 a new period starts). */
const PERIOD_DAYS = 9;
/** Number of periods used for consistency multiplier (trailing + current). */
const MULTIPLIER_NUM_PERIODS = 4;


/**
* Additional pay tiers: completions in current period → dollar amount.
* Sorted ascending by completions; first matching tier wins.
* Example: 20 → $50, 40 → $70, 60 → $80
*/
const PAY_THRESHOLDS = [
 { completions: 20, amount: 50 },
 { completions: 40, amount: 70 },
 { completions: 60, amount: 80 }
];


/**
* Multiplier tiers: (trailing 3 + current) period average → multiplier value.
* Sorted ascending by average; first matching tier wins.
* Example: 20 → 1, 40 → 1.1, 60 → 1.25
*/
const MULTIPLIER_THRESHOLDS = [
 { avgMin: 20, value: 1, badgeLabel: '1x Multiplier Active' },
 { avgMin: 40, value: 1.1, badgeLabel: '1.1x Multiplier Active' },
 { avgMin: 60, value: 1.25, badgeLabel: '1.25x Multiplier Active' }
];


// ——— Pilot launch: only these emails can access the dashboard. Remove after pilot and uncomment sheet-based access below. ———
const PILOT_ACCESS_EMAILS = [
 'lydia.huang@invisible.email',
 'michael.hernandez@invisible.email',
 'rana.traboulsi@invisible.email',
 'adrien.damseaux@invisible.email',
 'ioana.maier@invisible.email',
 'shihab.uddin@invisible.email',
 'sylwia.wlodyga@invisible.email',
 'taeheetay.jin@invisible.email',
 'aida.durakovic@invisible.email',
 'kristen.sonntag@invisible.email',
 'philip.gordon@invisible.email',
 'sharon.mcallister@invisible.email',
 'krzysztof.doda@invisible.email',
 'ha.anh@invisible.email',
 'anna.novoseltseva@invisible.email',
 'mark.weaver@invisible.email',
 'andrew.vi@invisible.email',
 'kubra.koc@invisible.email'
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
  var params = (e && e.parameter) ? e.parameter : {};
  var app = String(params.app || '').trim().toLowerCase();
  if (app === 'qa') {
    return HtmlService.createHtmlOutputFromFile('QA_Index')
      .setTitle('QA Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Trainer Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}


/** Main data endpoint for the UI **/
function getDashboardData() {
  const raw = Session.getActiveUser().getEmail();
  const email = raw ? normalizeEmail_(raw) : '';
  if (!email) return { noAccess: true };

  const pilotAllowed = PILOT_ACCESS_EMAILS.map(e => normalizeEmail_(e)).filter(Boolean);
  if (!pilotAllowed.includes(email)) return { noAccess: true };

  return getDashboardDataForEmail_(email);
}

/** Load sheet and build dashboard response for a validated pilot email. */
function getDashboardDataForEmail_(email) {
  const pilotAllowed = PILOT_ACCESS_EMAILS.map(e => normalizeEmail_(e)).filter(Boolean);

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Sheet "${SHEET_NAME}" not found`);

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

  // Read columns once (include COL_ERROR for quality eligibility)
  const numRows = lastRow - 1;
  const values = sh
    .getRange(2, 1, numRows, Math.max(COL_COMPLET, COL_STATUS, COL_ERROR))
    .getValues();

  const numPeriods = MULTIPLIER_NUM_PERIODS;
  let periodCompletions = 0;
  let multiPeriodTotalCompletions = 0;
  let multiPeriodErrorCount = 0;

  // Leaderboard totals for ALL trainers for current period
  const periodTotalsByEmail = Object.create(null);

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    const rowEmail = String(row[COL_EMAIL - 1] || '').trim().toLowerCase();
    if (!rowEmail) continue;

    const ts = row[COL_TIMESTAMP - 1];
    if (!(ts instanceof Date)) continue;

    if (ALLOWED_STATUSES && ALLOWED_STATUSES.length) {
      const st = String(row[COL_STATUS - 1] || '').trim();
      if (!ALLOWED_STATUSES.includes(st)) continue;
    }

    const c = Number(row[COL_COMPLET - 1] || 0);
    if (!isFinite(c)) continue;

    const inCurrentPeriod = isDateInRange_(ts, periodStartStr, periodEndStr, TZ);
    const inMultiPeriod = isDateInRange_(ts, multiPeriodStartStr, multiPeriodEndStr, TZ);

    if (inCurrentPeriod) {
      periodTotalsByEmail[rowEmail] = (periodTotalsByEmail[rowEmail] || 0) + c;
    }

    if (rowEmail === email) {
      if (inCurrentPeriod) periodCompletions += c;
      if (inMultiPeriod) {
        multiPeriodTotalCompletions += c;
        const err = Number(row[COL_ERROR - 1] || 0);
        if (isFinite(err)) multiPeriodErrorCount += err;
      }
    }
  }

  // Payout eligibility: quality rule over past 3 periods + current
  const qualityEligible = (function () {
    if (multiPeriodTotalCompletions > 8) {
      return multiPeriodErrorCount <= 0.125 * multiPeriodTotalCompletions;
    }
    return multiPeriodErrorCount <= 1;
  })();

  const multiPeriodAvg = multiPeriodTotalCompletions / numPeriods;

  // Leaderboard = everyone in Gold (sheet) with current-period completions + access-list agents with no data
  const fromSheet = Object.keys(periodTotalsByEmail).map(e => ({
    email: e,
    weeklyCompletions: periodTotalsByEmail[e] || 0
  }));
  const allAccessEmails = pilotAllowed;
  const accessOnlyNoData = allAccessEmails
    .filter(e => !(e in periodTotalsByEmail))
    .map(e => ({ email: e, weeklyCompletions: 0 }));
  const allPeriodRows = fromSheet.concat(accessOnlyNoData);

  const leaderboard = buildLeaderboard_(allPeriodRows, email);

  return buildResponse_(email, periodCompletions, multiPeriodAvg, periodRangeText, leaderboard, qualityEligible);
}


/** Build all UI-facing values (tiers, progress, text). Keeps weekRangeText/weekCompletions/fourWeekAvg keys for front-end compatibility. */
function buildResponse_(email, periodCompletions, multiPeriodAvg, periodRangeText, leaderboard, qualityEligible) {
 const pay = additionalPay_(periodCompletions);
 const payProgress = progressToNextPayTier_(periodCompletions);


 const mult = consistencyMultiplier_(multiPeriodAvg);
 const multProgress = progressToNextMultiplierTier_(multiPeriodAvg);


 let earnings = 0;
 if (pay.amount > 0) {
   earnings = mult.value ? pay.amount * mult.value : pay.amount;
 }
 const earningsInt = Math.round(earnings);

 // Only show incremental message when explicitly eligible; otherwise always show ineligibility (handles undefined if old code calls without 6th arg)
 const earningsText = (qualityEligible === true)
   ? `You are earning an incremental $${earningsInt} this week!`
   : 'You are not eligible for the additional earnings due to quality issues.';

 return {
   email,
   weekRangeText: periodRangeText,
   weekCompletions: Math.round(periodCompletions || 0),
   fourWeekAvg: Math.floor(isFinite(multiPeriodAvg) ? multiPeriodAvg : 0),


   additionalPay: {
     amount: pay.amount,
     qualifiedText: pay.text
   },
   payProgress,


   multiplier: {
     value: mult.value,
     badgeText: mult.badgeText
   },
   multProgress,


   earnings: {
     amount: (qualityEligible === true) ? earningsInt : 0,
     text: earningsText
   },


   leaderboard
 };
}


/** Tiers (driven by PAY_THRESHOLDS and MULTIPLIER_THRESHOLDS config) **/
function additionalPay_(completions) {
 const sorted = PAY_THRESHOLDS.slice().sort((a, b) => b.completions - a.completions);
 for (let i = 0; i < sorted.length; i++) {
   if (completions >= sorted[i].completions) {
     const amt = sorted[i].amount;
     return { amount: amt, text: `Qualified for $${amt} Additional Earnings` };
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


function progressToNextPayTier_(completions) {
 const sorted = PAY_THRESHOLDS.slice().sort((a, b) => a.completions - b.completions);
 const top = sorted[sorted.length - 1];
 if (completions >= top.completions) {
   return { barPct: 100, nextTierText: 'Keep up the great work!' };
 }


 let base = 0, next = sorted[0].completions, nextAmount = sorted[0].amount;
 for (let i = 0; i < sorted.length; i++) {
   if (completions < sorted[i].completions) {
     next = sorted[i].completions;
     nextAmount = sorted[i].amount;
     base = i > 0 ? sorted[i - 1].completions : 0;
     break;
   }
 }


 const span = next - base;
 const progressed = span ? Math.max(0, completions - base) : 0;
 const pct = span ? Math.min(100, Math.round((progressed / span) * 100)) : 0;
 const remaining = Math.max(0, next - completions);


 return {
   barPct: pct,
   nextTierText: `${remaining} Completions to $${nextAmount} Additional Earnings`
 };
}


function progressToNextMultiplierTier_(avg) {
 const a = Number(avg);
 const safeAvg = isFinite(a) ? a : 0;
 const avgInt = Math.floor(safeAvg);


 const sorted = MULTIPLIER_THRESHOLDS.slice().sort((a, b) => b.avgMin - a.avgMin);
 if (safeAvg >= sorted[0].avgMin) {
   return {
     avgInt,
     nextTierText: `0 completion till ${sorted[0].badgeLabel.toLowerCase()}`
   };
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
 return {
   avgInt,
   nextTierText: `${remaining} completion(s) till ${label}`
 };
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
   periodEnd: periodEndDate
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


