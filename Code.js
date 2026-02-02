/** CONFIG **/
const SHEET_NAME = 'Gold_Testing';
const TZ = 'America/New_York';

// Column indices (1-based)
const COL_TIMESTAMP = 6; // F
const COL_EMAIL     = 7; // G
const COL_STATUS    = 9; // I
const COL_COMPLET   = 11; // K
const ALLOWED_STATUSES = ['Task Submitted'];

/** Web app entry **/
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Trainer Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/** Main data endpoint for the UI **/
function getDashboardData() {
  const email = (Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email) {
    throw new Error(
      'Unable to determine your email. This dashboard requires Google Workspace sign-in and correct web app deployment settings.'
    );
  }

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Sheet "${SHEET_NAME}" not found`);

  const now = new Date();
  const { weekStart, weekEnd } = getWeekBounds_(now, TZ);
  const weekRangeText =
    `${Utilities.formatDate(weekStart, TZ, 'M/d/yyyy')} - ${Utilities.formatDate(weekEnd, TZ, 'M/d/yyyy')}`;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    // No rows; treat as no access (or change if you prefer)
    return { noAccess: true };
  }

  // Access control: viewer email must exist in COL_EMAIL
  const emailValues = sh
    .getRange(2, COL_EMAIL, lastRow - 1, 1)
    .getValues()
    .flat()
    .map(e => String(e || '').trim().toLowerCase());

  if (!emailValues.includes(email)) {
    return { noAccess: true };
  }

  // Read columns once
  const numRows = lastRow - 1;
  const values = sh
    .getRange(2, 1, numRows, Math.max(COL_COMPLET, COL_STATUS))
    .getValues();

  // Past 4 FULL weeks (including current week)
  const fourWeekStart = addDays_(weekStart, -21);
  const fourWeekEnd   = weekEnd; 

  let weekCompletions = 0;
  let fourWeekTotalCompletions = 0;

  // Leaderboard totals for ALL trainers for current week
  const weeklyTotalsByEmail = Object.create(null);

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

    // Current week totals (all trainers)
    if (ts >= weekStart && ts <= weekEnd) {
      weeklyTotalsByEmail[rowEmail] = (weeklyTotalsByEmail[rowEmail] || 0) + c;
    }

    // Viewer-specific totals
    if (rowEmail === email) {
      if (ts >= weekStart && ts <= weekEnd) weekCompletions += c;
      if (ts >= fourWeekStart && ts <= fourWeekEnd) fourWeekTotalCompletions += c;
    }
  }

  const fourWeekAvg = fourWeekTotalCompletions / 4;

  const allWeeklyRows = Object.keys(weeklyTotalsByEmail).map(e => ({
    email: e,
    weeklyCompletions: weeklyTotalsByEmail[e] || 0
  }));

  const leaderboard = buildLeaderboard_(allWeeklyRows, email);

  return buildResponse_(email, weekCompletions, fourWeekAvg, weekRangeText, leaderboard);
}

/** Build all UI-facing values (tiers, progress, text) */
function buildResponse_(email, weekCompletions, fourWeekAvg, weekRangeText, leaderboard) {
  const pay = additionalPay_(weekCompletions);
  const payProgress = progressToNextPayTier_(weekCompletions);

  const mult = consistencyMultiplier_(fourWeekAvg);
  const multProgress = progressToNextMultiplierTier_(fourWeekAvg);

  let earnings = 0;
  if (pay.amount > 0) {
    earnings = mult.value ? pay.amount * mult.value : pay.amount;
  }
  const earningsInt = Math.round(earnings);

  return {
    email,
    weekRangeText,
    weekCompletions: Math.round(weekCompletions || 0),
    fourWeekAvg: Math.floor(isFinite(fourWeekAvg) ? fourWeekAvg : 0),

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
      amount: earningsInt,
      text: `You are earning an incremental $${earningsInt} this week!`
    },

    leaderboard
  };
}

/** Tiers **/
function additionalPay_(weekly) {
  if (weekly >= 40) return { amount: 80, text: 'Qualified for $80 Additional Earnings' };
  if (weekly >= 28) return { amount: 70, text: 'Qualified for $70 Additional Earnings' };
  if (weekly >= 16) return { amount: 50, text: 'Qualified for $50 Additional Earnings' };
  return { amount: 0, text: 'Qualified for $0 Additional Earning' };
}

function consistencyMultiplier_(avg) {
  const a = Number(avg);
  const safe = isFinite(a) ? a : 0;
  if (safe >= 40) return { value: 1.25, badgeText: '1.25x Multiplier Active' };
  if (safe >= 28) return { value: 1.10, badgeText: '1.1x Multiplier Active' };
  return { value: 0, badgeText: 'No badge available right now' };
}

function progressToNextPayTier_(weekly) {
  if (weekly >= 40) {
    return { barPct: 100, nextTierText: 'Keep up the great work!' };
  }

  let base = 0, next = 16;
  if (weekly >= 16 && weekly < 28) { base = 16; next = 28; }
  else if (weekly >= 28 && weekly < 40) { base = 28; next = 40; }

  const span = next - base;
  const progressed = Math.max(0, weekly - base);
  const pct = Math.min(100, Math.round((progressed / span) * 100));

  const remaining = Math.max(0, next - weekly);
  const nextPay = (next === 16) ? 50 : (next === 28) ? 70 : 80;

  return {
    barPct: pct,
    nextTierText: `${remaining} Completions to $${nextPay} Additional Earnings`
  };
}

function progressToNextMultiplierTier_(avg) {
  const a = Number(avg);
  const safeAvg = isFinite(a) ? a : 0;

  // Displayed 4-week average integer (always round down)
  const avgInt = Math.floor(safeAvg);

  // Pick the goal tier based on the REAL average (not the floored display)
  // This avoids showing "0 to 1.1x" when the true avg is 28.2, etc.
  if (safeAvg >= 40) {
    return {
      avgInt,
      nextTierText: '0 completion till 1.25x multiplier active'
    };
  }

  const goal = (safeAvg >= 28) ? 40 : 28;
  const label = (goal === 40) ? '1.25x' : '1.1x';

  // Remaining is derived from the displayed integer so they always add up
  const remaining = Math.max(0, goal - avgInt);

  return {
    avgInt,
    nextTierText: `${remaining} completion(s) till ${label} multiplier active`
  };
}


/** Leaderboard (shared ranks / competition ranking) **/
/** Leaderboard (shared ranks / competition ranking) **/
function buildLeaderboard_(rows, viewerEmail) {
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

/** Week bounds: Monday 00:00:00 to Sunday 23:59:59 in TZ */
function getWeekBounds_(date, tz) {
  const y = Number(Utilities.formatDate(date, tz, 'yyyy'));
  const m = Number(Utilities.formatDate(date, tz, 'MM')) - 1;
  const d = Number(Utilities.formatDate(date, tz, 'dd'));

  const local = new Date(y, m, d, 12, 0, 0);
  const day = local.getDay(); // 0 Sun..6 Sat
  const mondayOffset = (day + 6) % 7;

  const monday = new Date(local);
  monday.setDate(local.getDate() - mondayOffset);
  monday.setHours(0, 0, 0, 0);

  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  sunday.setHours(23, 59, 59, 999);

  return { weekStart: monday, weekEnd: sunday };
}

function addDays_(dt, days) {
  const d = new Date(dt);
  d.setDate(d.getDate() + days);
  return d;
}
