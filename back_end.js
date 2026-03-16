/**
 * Production sheet: Col C = unique count, F = Timestamp, G = Email, H = Batch, I = Status, K = Team (1-4 or empty).
 * We read by column letter, filter rows that meet the requirement, then count unique Col C per team.
 */
var BATCH_FILTER = "rm";
var TIMEZONE = "America/New_York";
var TEAM_NAMES = { 1: "Fenrir", 2: "Titan", 3: "Pegasus", 4: "Sphinx" };
var GOAL = 14;

function teamToId_(team) {
  if (team == null || team === "") return null;
  var n = parseInt(team, 10);
  return (n >= 1 && n <= 4) ? n : null;
}

function normalizeStatus_(s) {
  if (s == null) return "";
  return String(s).trim().toLowerCase().replace(/\s+/g, " ");
}

/** From timestamp (string "M/D/YYYY ..." or Date), return "yyyy-MM-dd" for comparison. Row string is already EST — use the date in the string as-is, no timezone conversion. */
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

/** Read Production sheet. Col F read as display values so date string parses as US M/D/YYYY. */
function readProductionRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var c = sheet.getRange("C2:C" + lastRow).getValues();
  var fRange = sheet.getRange("F2:F" + lastRow);
  var f = fRange.getDisplayValues();
  var g = sheet.getRange("G2:G" + lastRow).getValues();
  var h = sheet.getRange("H2:H" + lastRow).getValues();
  var iCol = sheet.getRange("I2:I" + lastRow).getValues();
  var k = sheet.getRange("K2:K" + lastRow).getValues();
  return c.map(function(_, idx) {
    var tsRaw = f[idx][0];
    return { sheetRow: idx + 2, colC: c[idx][0], timestamp: tsRaw !== "" && tsRaw != null ? String(tsRaw).trim() : null, email: g[idx][0], batch: h[idx][0], status: iCol[idx][0], team: k[idx][0] };
  });
}

/** Col I statuses that count: "task submitted" or "revised" (case-insensitive). */
function isCountedStatus_(status) {
  var n = normalizeStatus_(status);
  return n === "task submitted" || n === "revised";
}

/** Row meets: has timestamp, team, status; date = dateStr (EST); batch contains "rm"; status = "task submitted" or "revised"; team 1-4. Uses row F as-is (already EST), no timezone conversion. */
function rowMeets_(row, dateStr) {
  if (!row.timestamp || !row.team || !row.status) return false;
  try {
    var rowDateEST = timestampToDateStrEST_(row.timestamp);
    if (!rowDateEST) return false;
    var okBatch = String(row.batch).toLowerCase().includes(BATCH_FILTER);
    var okStatus = isCountedStatus_(row.status);
    var teamKey = teamToId_(row.team);
    return rowDateEST === dateStr && okBatch && okStatus && teamKey !== null;
  } catch (e) { return false; }
}

/** From qualifying rows, count unique colC per team id (1-4). Returns { 1: n, 2: n, 3: n, 4: n }. */
function countUniqueColCByTeam_(rows) {
  var sets = { 1: {}, 2: {}, 3: {}, 4: {} };
  rows.forEach(function(row) {
    var teamKey = teamToId_(row.team);
    if (teamKey && sets[teamKey] !== undefined && row.colC != null && row.colC !== "") {
      sets[teamKey][String(row.colC)] = true;
    }
  });
  return { 1: Object.keys(sets[1]).length, 2: Object.keys(sets[2]).length, 3: Object.keys(sets[3]).length, 4: Object.keys(sets[4]).length };
}

function GET_TEAM_LEADERBOARD() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Production");
  if (!sheet) return [["Team", "Submitted", "Progress %"]];
  var todayEST = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd");
  var rows = readProductionRows_(sheet);
  var qualified = rows.filter(function(row) { return rowMeets_(row, todayEST); });
  var stats = countUniqueColCByTeam_(qualified);

  var results = [["Team", "Submitted", "Progress %"]];
  [1, 2, 3, 4].forEach(function(id) {
    results.push([TEAM_NAMES[id] || ("Team " + id), stats[id], stats[id] / GOAL]);
  });
  return results;
}

  /** Debug: same logic as GET_TEAM_LEADERBOARD. Run from Apps Script (Run > DEBUG_COUNTED_ROWS_BY_GROUP). */
  function DEBUG_COUNTED_ROWS_BY_GROUP() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Production");
    if (!sheet) return { error: "Sheet 'Production' not found" };
    var todayEST = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd");
    var rows = readProductionRows_(sheet);
    var qualified = rows.filter(function(row) { return rowMeets_(row, todayEST); });
    var counts = countUniqueColCByTeam_(qualified);
    var byTeam = { "Fenrir": [], "Titan": [], "Pegasus": [], "Sphinx": [] };
    qualified.forEach(function(row) {
      var name = TEAM_NAMES[teamToId_(row.team)];
      if (name) byTeam[name].push({ sheetRow: row.sheetRow, colC: row.colC, timestamp: String(row.timestamp), email: String(row.email || ""), batch: String(row.batch || ""), status: String(row.status || ""), team: row.team });
    });
    var countObj = { "Fenrir": counts[1], "Titan": counts[2], "Pegasus": counts[3], "Sphinx": counts[4] };
    Logger.log("--- DEBUG_COUNTED_ROWS_BY_GROUP --- todayEST = " + todayEST + " counts = " + JSON.stringify(countObj));
    Logger.log(JSON.stringify({ todayEST: todayEST, counts: countObj, byTeam: byTeam }, null, 2));
    return { todayEST: todayEST, counts: countObj, byTeam: byTeam };
  }

  /**
   * Run this from Apps Script (Run > DEBUG_WHY_ROWS_NOT_COUNTED) to see exactly why each row
   * is INCLUDED or EXCLUDED. Check View > Logs or Executions to see the output.
   * Only logs rows that have batch containing "rm" so you can spot date/team/status issues.
   */
  function DEBUG_WHY_ROWS_NOT_COUNTED() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Production");
    if (!sheet) { Logger.log("ERROR: Sheet 'Production' not found"); return; }
    var todayEST = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd");
    var rows = readProductionRows_(sheet);
    Logger.log("========== DEBUG: Why rows are included or excluded ==========");
    Logger.log("todayEST (date we count as 'today') = " + todayEST);
    Logger.log("Total data rows read: " + rows.length);
    var included = 0;
    var excludedReasons = {};
    rows.forEach(function(row) {
      var batchStr = String(row.batch || "").toLowerCase();
      var hasRm = batchStr.indexOf(BATCH_FILTER) !== -1;
      if (!hasRm) return;
      var statusNorm = normalizeStatus_(row.status);
      var hasCountedStatus = isCountedStatus_(row.status);
      var teamKey = teamToId_(row.team);
      var hasTeam = teamKey !== null;
      var reason = null;
      var rowDateEST = null;
      if (!row.timestamp) reason = "no timestamp";
      else if (!row.status) reason = "no status";
      else if (!hasTeam) reason = "team empty or not 1-4 (raw team='" + JSON.stringify(row.team) + "')";
      else if (!hasCountedStatus) reason = "status not 'task submitted' or 'revised' (normalized='" + statusNorm + "')";
      else {
        rowDateEST = timestampToDateStrEST_(row.timestamp);
        if (!rowDateEST) reason = "timestamp could not be parsed";
        else if (rowDateEST !== todayEST) reason = "date mismatch (rowDateEST=" + rowDateEST + ", todayEST=" + todayEST + ")";
      }
      if (reason) {
        excludedReasons[reason] = (excludedReasons[reason] || 0) + 1;
        Logger.log("EXCLUDED sheetRow=" + row.sheetRow + " | F='" + String(row.timestamp).substring(0, 25) + "' | team=" + row.team + " | " + reason);
      } else {
        included++;
        Logger.log("INCLUDED  sheetRow=" + row.sheetRow + " | F='" + String(row.timestamp).substring(0, 25) + "' | team=" + TEAM_NAMES[teamKey] + " | colC=" + row.colC);
      }
    });
    Logger.log("--- Summary: INCLUDED=" + included + " | EXCLUDED by reason: " + JSON.stringify(excludedReasons));
    var qualified = rows.filter(function(row) { return rowMeets_(row, todayEST); });
    var counts = countUniqueColCByTeam_(qualified);
    Logger.log("Final unique Col C counts: " + JSON.stringify({ 1: counts[1], 2: counts[2], 3: counts[3], 4: counts[4] }));
  }

  /** Last 7 Days Winner: per day, which team had most unique Col C (same requirement). Returns [{ dateStr, dateDisplay, teamName, count, isToday }, ...]. */
  function GET_LAST_7_DAYS_WINNERS() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Production");
    if (!sheet) return [];
    var rows = readProductionRows_(sheet);
    var dates = [6, 5, 4, 3, 2, 1, 0].map(function(k) {
      return Utilities.formatDate(new Date(new Date().getTime() - k * 24 * 60 * 60 * 1000), TIMEZONE, "yyyy-MM-dd");
    });
    return dates.map(function(dateStr, di) {
      var qualified = rows.filter(function(row) { return rowMeets_(row, dateStr); });
      var stats = countUniqueColCByTeam_(qualified);
      var maxCount = 0, winnerId = null;
      [1, 2, 3, 4].forEach(function(id) {
        if (stats[id] > maxCount) { maxCount = stats[id]; winnerId = id; }
      });
      var teamName = winnerId ? TEAM_NAMES[winnerId] : "—";
      var dateDisplay = Utilities.formatDate(new Date(dateStr + "T12:00:00"), TIMEZONE, "MMM d");
      return { dateStr: dateStr, dateDisplay: dateDisplay, teamName: teamName, count: maxCount || 0, isToday: di === dates.length - 1 };
    });
  }

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

  /** Row in last 7 days and meets requirement. Uses row F date as-is (already EST). */
  function rowInLast7AndMeets_(row, dateStrs) {
    if (!row.timestamp || !row.team || !row.status) return false;
    try {
      var rowDateEST = timestampToDateStrEST_(row.timestamp);
      if (!rowDateEST || dateStrs.indexOf(rowDateEST) === -1) return false;
      return rowMeets_(row, rowDateEST);
    } catch (e) { return false; }
  }

  /** Top 3 performers per team in last 7 days: unique Col C count per (team, email). Returns { "Fenrir": [{ name, count }, ...], ... }. */
  function GET_TOP_PERFORMERS_LAST_7_DAYS() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Production");
    if (!sheet) return {};
    var dateStrs = [6, 5, 4, 3, 2, 1, 0].map(function(k) {
      return Utilities.formatDate(new Date(new Date().getTime() - k * 24 * 60 * 60 * 1000), TIMEZONE, "yyyy-MM-dd");
    });
    var rows = readProductionRows_(sheet);
    var qualified = rows.filter(function(row) { return rowInLast7AndMeets_(row, dateStrs); });
    var byTeamEmail = { 1: {}, 2: {}, 3: {}, 4: {} };
    qualified.forEach(function(row) {
      var teamKey = teamToId_(row.team);
      if (!teamKey || !byTeamEmail[teamKey]) return;
      var e = String(row.email || "").trim().toLowerCase();
      if (!e) return;
      if (!byTeamEmail[teamKey][e]) byTeamEmail[teamKey][e] = {};
      if (row.colC != null && row.colC !== "") byTeamEmail[teamKey][e][String(row.colC)] = true;
    });
    var out = {};
    [1, 2, 3, 4].forEach(function(id) {
      var name = TEAM_NAMES[id];
      var list = [];
      for (var em in byTeamEmail[id]) list.push({ email: em, count: Object.keys(byTeamEmail[id][em]).length });
      list.sort(function(a, b) { return b.count - a.count; });
      out[name] = list.slice(0, 3).map(function(x) { return { name: emailToDisplayName_(x.email), count: x.count }; });
    });
    return out;
  }

  function doGet() {
    const timestamp = Utilities.formatDate(new Date(), "America/New_York", "h:mm a 'ET'");
    
    // Create the template
    const html = HtmlService.createTemplateFromFile('front_end');
    html.lastUpdated = timestamp; // Pass the time to the HTML
    
    return html.evaluate()
        .setTitle('Team Leaderboard')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }