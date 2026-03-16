/**
 * Submission sheet (active spreadsheet): Col A = Timestamp, B = Email, D = task_id, G = Team, J = Batch, M = Status.
 * Config sheet (external): https://docs.google.com/spreadsheets/d/1ue7Ey5XfxB97t2O9TL70gvK8OIR0GXIxYTnvBu-i_Xk
 *   Sheet 1: Col A = batch name, Col D = end date (empty = batch still running).
 * Current batches = rows where Col D is empty. Closed = Col D not empty and today (ET) > Col D.
 * Batch winners: initial stars + for each closed batch, +1 star to team with most unique task_id (Col D) for that batch.
 */
var TIMEZONE = "America/New_York";
var CONFIG_SPREADSHEET_ID = "1ue7Ey5XfxB97t2O9TL70gvK8OIR0GXIxYTnvBu-i_Xk";
/** Col G team names (exactly as in sheet). */
var AC_TEAMS = ["Legion", "Squadron", "Titans", "Vanguard"];
/** Apply star logic only when batch finish date (Col D) is after this. Format yyyy-MM-dd. */
var STAR_LOGIC_CUTOFF_DATE = "2026-03-08";
/** Batch Winners Leaderboard: initial stars (batches finished on or before STAR_LOGIC_CUTOFF_DATE). */
var BATCH_WINNERS_STARS = { "Titans": 2, "Legion": 1, "Squadron": 1, "Vanguard": 0 };

/** Normalize Col G to one of AC_TEAMS (case-insensitive trim), or null. */
function normalizeTeam_(team) {
  if (team == null || team === "") return null;
  var s = String(team).trim();
  for (var i = 0; i < AC_TEAMS.length; i++) {
    if (AC_TEAMS[i].toLowerCase() === s.toLowerCase()) return AC_TEAMS[i];
  }
  return null;
}

function normalizeStatus_(s) {
  if (s == null) return "";
  return String(s).trim().toLowerCase().replace(/\s+/g, " ");
}

/** From timestamp (string "M/D/YYYY ..." or Date), return "yyyy-MM-dd" for comparison. */
function timestampToDateStrEST_(val) {
  if (val == null || val === "") return null;
  if (val instanceof Date) return Utilities.formatDate(val, TIMEZONE, "yyyy-MM-dd");
  var s = String(val).trim();
  var space = s.indexOf(" ");
  var datePart = space >= 0 ? s.substring(0, space) : s;
  var parts = datePart.split("/");
  if (parts.length !== 3) return null;
  var month = parseInt(parts[0], 10);
  var day = parseInt(parts[1], 10);
  var year = parseInt(parts[2], 10);
  if (isNaN(month) || isNaN(day) || isNaN(year)) return null;
  var pad = function(n) { return n < 10 ? "0" + n : String(n); };
  return year + "-" + pad(month) + "-" + pad(day);
}

/** Read Submission sheet in one range read (A:M). Col A=timestamp, B=email, D=task_id, G=team, J=batch, M=Status. */
function readSubmissionRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange("A2:M" + lastRow).getDisplayValues();
  return data.map(function(row, idx) {
    var tsRaw = row[0];
    return {
      sheetRow: idx + 2,
      taskId: row[3],
      timestamp: tsRaw !== "" && tsRaw != null ? String(tsRaw).trim() : null,
      email: row[1],
      team: row[6],
      batch: row[9],
      status: row[12]
    };
  });
}

/** Only "Submitted" (case-insensitive) counts. */
function isCountedStatus_(status) {
  return normalizeStatus_(status) === "submitted";
}

/** Row is in given batch: Col J = batchName, status "Submitted", team in AC_TEAMS. */
function rowMeetsBatch_(row, batchName) {
  if (batchName == null || batchName === "") return false;
  if (!row.timestamp || !row.team || !row.status) return false;
  var batchStr = row.batch != null ? String(row.batch).trim() : "";
  if (batchStr !== batchName) return false;
  if (!isCountedStatus_(row.status)) return false;
  return normalizeTeam_(row.team) !== null;
}

/** Parse config sheet (Sheet 1): Col A = batch name, Col D = end date (e.g. "3/7/2026 6:03:11").
 *  Star logic applies only when Col D finish date > STAR_LOGIC_CUTOFF_DATE (and batch closed: today > Col D).
 *  Return { currentBatches: [], closedBatches: [] }. */
function readConfigBatches_() {
  var out = { currentBatches: [], closedBatches: [] };
  try {
    var configSs = SpreadsheetApp.openById(CONFIG_SPREADSHEET_ID);
    var configSh = configSs.getSheetByName("Sheet1");
    if (!configSh) return out;
    var lastRow = configSh.getLastRow();
    if (lastRow < 2) return out;
    var data = configSh.getRange("A2:D" + lastRow).getDisplayValues();
    var todayStr = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd");
    for (var i = 0; i < data.length; i++) {
      var batchName = data[i][0] != null ? String(data[i][0]).trim() : "";
      if (!batchName) continue;
      var colD = data[i][3];
      var colDEmpty = colD == null || String(colD).trim() === "";
      if (colDEmpty) {
        out.currentBatches.push(batchName);
      } else {
        var endDate = parseConfigDate_(colD);
        if (endDate && todayStr > endDate && endDate > STAR_LOGIC_CUTOFF_DATE) {
          out.closedBatches.push(batchName);
        }
      }
    }
  } catch (e) {}
  return out;
}

/** Parse config Col D to "yyyy-MM-dd" for comparison. Handles "3/7/2026 6:03:11" (date with time). */
function parseConfigDate_(val) {
  if (val == null || val === "") return null;
  if (val instanceof Date) return Utilities.formatDate(val, TIMEZONE, "yyyy-MM-dd");
  var s = String(val).trim();
  var datePart = s.indexOf(" ") >= 0 ? s.substring(0, s.indexOf(" ")).trim() : s;
  var parts = datePart.split("/");
  if (parts.length !== 3) return null;
  var m = parseInt(parts[0], 10), d = parseInt(parts[1], 10), y = parseInt(parts[2], 10);
  if (isNaN(m) || isNaN(d) || isNaN(y)) return null;
  var pad = function(n) { return n < 10 ? "0" + n : String(n); };
  return y + "-" + pad(m) + "-" + pad(d);
}

/** From qualifying rows, count unique Col D (task_id) per team. Returns { "Legion": n, "Squadron": n, "Titans": n, "Vanguard": n }. */
function countUniqueTaskIdByTeam_(rows) {
  var sets = {};
  AC_TEAMS.forEach(function(t) { sets[t] = {}; });
  rows.forEach(function(row) {
    var teamKey = normalizeTeam_(row.team);
    if (teamKey && sets[teamKey] && row.taskId != null && row.taskId !== "") {
      sets[teamKey][String(row.taskId)] = true;
    }
  });
  var out = {};
  AC_TEAMS.forEach(function(t) { out[t] = Object.keys(sets[t]).length; });
  return out;
}

/** Leaderboard for current batch: unique task_id count per team. Returns [["Team", "Submitted"], [name, count], ...]. */
function GET_TEAM_LEADERBOARD() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Submission sheet");
  if (!sheet) return [["Team", "Submitted"]];
  var rows = readSubmissionRows_(sheet);
  var qualified = rows.filter(function(row) { return rowMeetsBatch_(row); });
  var stats = countUniqueTaskIdByTeam_(qualified);

  var results = [["Team", "Submitted"]];
  AC_TEAMS.forEach(function(name) {
    results.push([name, stats[name] || 0]);
  });
  return results;
}

/** Min timestamp (ms since epoch) among rows where Col J = CURRENT_BATCH and status = Submitted. For "fighting for" timer. */
function GET_BATCH_START_TIMESTAMP() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Submission sheet");
  if (!sheet) return null;
  var rows = readSubmissionRows_(sheet);
  var qualified = rows.filter(function(row) { return rowMeetsBatch_(row) && row.timestamp; });
  if (qualified.length === 0) return null;
  var minMs = null;
  qualified.forEach(function(row) {
    var d = parseTimestampToDate_(row.timestamp);
    if (d && !isNaN(d.getTime())) {
      var ms = d.getTime();
      if (minMs === null || ms < minMs) minMs = ms;
    }
  });
  return minMs;
}

/** Parse Col A as Eastern only. Timer = current Eastern time − min(Col A) in Eastern. Strip " (EST)" / " (EDT)" then parse. */
function parseTimestampToDate_(tsStr) {
  if (!tsStr || typeof tsStr !== "string") return null;
  var s = String(tsStr).trim()
    .replace(/\s*\(EST\)\s*$/i, "").replace(/\s*\(EDT\)\s*$/i, "")
    .replace(/\s*\(ET\)\s*$/i, "").trim();
  if (!s) return null;
  var tz = "America/New_York";
  var formats = [
    "M/d/yyyy H:mm:ss",
    "M/d/yyyy h:mm:ss a",
    "M/d/yyyy h:mm a"
  ];
  for (var i = 0; i < formats.length; i++) {
    try {
      var d = Utilities.parseDate(s, tz, formats[i]);
      if (d && !isNaN(d.getTime())) return d;
    } catch (e) {}
  }
  return null;
}

/** First name + last initial from Col B (email). */
function emailToDisplayName_(email) {
  if (!email || typeof email !== "string") return "—";
  var s = String(email).trim().toLowerCase();
  var at = s.indexOf("@");
  if (at <= 0) return "—";
  var local = s.substring(0, at);
  var dot = local.indexOf(".");
  if (dot <= 0) return local.charAt(0).toUpperCase() + local.substring(1) + ".";
  var first = local.substring(0, dot).trim();
  var last = local.substring(dot + 1).trim();
  var firstCap = first.length ? first.charAt(0).toUpperCase() + first.substring(1) : "";
  var lastInitial = last.length ? last.charAt(0).toUpperCase() + "." : "";
  return (firstCap + " " + lastInitial).trim() || "—";
}

/** Top 3 performers per team in current batch: unique task_id (Col D) per (team, email). Display = first + last initial from Col B. */
function GET_TOP_PERFORMERS_CURRENT_BATCH() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Submission sheet");
  if (!sheet) return {};
  var rows = readSubmissionRows_(sheet);
  var qualified = rows.filter(function(row) { return rowMeetsBatch_(row); });
  return buildTopPerformersFromRows_(qualified);
}

/** Build top performers object from already-filtered rows (avoids re-reading sheet). */
function buildTopPerformersFromRows_(qualified) {
  var byTeamEmail = {};
  AC_TEAMS.forEach(function(t) { byTeamEmail[t] = {}; });
  qualified.forEach(function(row) {
    var teamKey = normalizeTeam_(row.team);
    if (!teamKey || !byTeamEmail[teamKey]) return;
    var e = String(row.email || "").trim().toLowerCase();
    if (!e) return;
    if (!byTeamEmail[teamKey][e]) byTeamEmail[teamKey][e] = {};
    if (row.taskId != null && row.taskId !== "") byTeamEmail[teamKey][e][String(row.taskId)] = true;
  });
  var out = {};
  AC_TEAMS.forEach(function(name) {
    var list = [];
    for (var em in byTeamEmail[name]) list.push({ email: em, count: Object.keys(byTeamEmail[name][em]).length });
    list.sort(function(a, b) { return b.count - a.count; });
    out[name] = list.slice(0, 3).map(function(x) { return { name: emailToDisplayName_(x.email), count: x.count }; });
  });
  return out;
}

/** Single call: read config + Submission once; return per-batch data + batch winners. */
function GET_AC_DASHBOARD_DATA() {
  var config = readConfigBatches_();
  var currentBatchNames = config.currentBatches || [];
  var closedBatchNames = config.closedBatches || [];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Submission sheet");
  var rows = [];
  if (sheet) rows = readSubmissionRows_(sheet);

  var currentBatchesData = [];
  for (var b = 0; b < currentBatchNames.length; b++) {
    var batchName = currentBatchNames[b];
    var qualified = rows.filter(function(row) { return rowMeetsBatch_(row, batchName); });
    var stats = countUniqueTaskIdByTeam_(qualified);
    var leaderboard = [["Team", "Submitted"]];
    AC_TEAMS.forEach(function(name) {
      leaderboard.push([name, stats[name] || 0]);
    });
    var minMs = null;
    qualified.forEach(function(row) {
      if (!row.timestamp) return;
      var d = parseTimestampToDate_(row.timestamp);
      if (d && !isNaN(d.getTime())) {
        var ms = d.getTime();
        if (minMs === null || ms < minMs) minMs = ms;
      }
    });
    currentBatchesData.push({
      batchName: batchName,
      leaderboard: leaderboard,
      topPerformers: buildTopPerformersFromRows_(qualified),
      batchStartMs: minMs
    });
  }

  var batchWinners = {};
  AC_TEAMS.forEach(function(t) { batchWinners[t] = BATCH_WINNERS_STARS[t] || 0; });
  for (var c = 0; c < closedBatchNames.length; c++) {
    var closedBatch = closedBatchNames[c];
    var closedRows = rows.filter(function(row) {
      if (!isCountedStatus_(row.status)) return false;
      var batchStr = row.batch != null ? String(row.batch).trim() : "";
      if (batchStr !== closedBatch) return false;
      return normalizeTeam_(row.team) !== null;
    });
    var closedStats = countUniqueTaskIdByTeam_(closedRows);
    var maxCount = 0, winnerTeam = null;
    AC_TEAMS.forEach(function(name) {
      var n = closedStats[name] || 0;
      if (n > maxCount) { maxCount = n; winnerTeam = name; }
    });
    if (winnerTeam) batchWinners[winnerTeam] = (batchWinners[winnerTeam] || 0) + 1;
  }

  return {
    currentBatchesData: currentBatchesData,
    batchWinners: batchWinners
  };
}

function doGet() {
  var timestamp = Utilities.formatDate(new Date(), "America/New_York", "h:mm a 'ET'");
  var html = HtmlService.createTemplateFromFile('AC_front_end');
  html.lastUpdated = timestamp;
  return html.evaluate()
    .setTitle('AC Team Leaderboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
