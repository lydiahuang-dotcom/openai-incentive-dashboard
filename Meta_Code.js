/**
 * Deploy with Meta_Code.js + Meta_Index.html bound to your spreadsheet.
 * Aggregates Hubstaff_Export, Metabase_Expert_Country_List, Airtable_Expert_Country_List, Agent_Geo_Pay_RateCard.
 */

/** Emails allowed to use this web app (deploy “Execute as: User accessing the web app”). */
var META_MANAGER_EMAILS = [
  'lydia.huang@invisible.email',
  'rebecca.harrison@invisible.email',
  'beatrice.tomasello@invisible.email'
];

/** From the Google Sheet URL: .../d/<ID>/edit — required when users open the /exec web app (getActive() is not reliable there). */
var META_SPREADSHEET_ID = '';

function META_getSpreadsheet_() {
  var id = String(META_SPREADSHEET_ID || '').trim();
  if (!id) {
    try {
      id = String(PropertiesService.getScriptProperties().getProperty('META_SPREADSHEET_ID') || '').trim();
    } catch (e) {
      id = '';
    }
  }
  if (id) {
    return SpreadsheetApp.openById(id);
  }
  try {
    return SpreadsheetApp.getActive();
  } catch (e) {
    throw new Error(
      'Set META_SPREADSHEET_ID to your sheet ID from the URL, or open the script from the spreadsheet. Web apps need the ID.'
    );
  }
}

/** Web app entry — standalone deployment only. */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Meta_Index')
    .setTitle('Hubstaff Meta — Geo pay')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/** Normalize email: strip "Name <email>", trim, lowercase. */
function META_normalizeEmail_(str) {
  var s = String(str || '').trim();
  var match = s.match(/<([^>]+)>/);
  var emailOnly = match ? match[1].trim() : s;
  return emailOnly.toLowerCase();
}

function META_isManager() {
  try {
    var raw = Session.getActiveUser().getEmail();
    var email = raw ? META_normalizeEmail_(raw) : '';
    if (!email) return false;
    var list = META_MANAGER_EMAILS.map(function (e) {
      return META_normalizeEmail_(e);
    }).filter(Boolean);
    return list.indexOf(email) !== -1;
  } catch (e) {
    return false;
  }
}

var META_SHEET_HUBSTAFF = 'Hubstaff_Export';
var META_SHEET_EXPERTS = 'Metabase_Expert_Country_List';
/** Col B = email, col D = country — second source; wins on conflict with Metabase. */
var META_SHEET_AIRTABLE_EXPERTS = 'Airtable_Expert_Country_List';
var META_SHEET_RATE_CARD = 'Agent_Geo_Pay_RateCard';

/** Rate type → [country col, rate col] (1-based). */
var META_GEO_COLS = {
  1: { country: 1, rate: 4 },   // A, D — INV Advanced Geo
  2: { country: 6, rate: 9 },   // F, I — INV Expert Geo
  3: { country: 11, rate: 12 }, // K, L — ML/ENG Geo
  4: { country: 16, rate: 17 }, // P, Q — Coding Expert Geo
  5: { country: 21, rate: 24 }  // U, X — INV Generalist Geo
};

var META_RATE_LABELS = [
  '',
  'INV Advanced Geo Rates',
  'INV Expert Geo Rates',
  'ML/ENG Geo Rates',
  'Coding Expert Geo Rates',
  'INV Generalist Geo Rates',
  'Custom'
];

function metaNormalizeNamePart_(s) {
  return String(s || '')
    .trim()
    .replace(/\s+/g, ' ');
}

function metaEmailToFirstLast_(email) {
  var e = META_normalizeEmail_(email);
  var at = e.indexOf('@');
  if (at <= 0) return { first: '', last: '' };
  var local = e.substring(0, at);
  var parts = local.split('.');
  if (parts.length < 2) {
    return { first: metaNormalizeNamePart_(parts[0] || ''), last: '' };
  }
  return {
    first: metaNormalizeNamePart_(parts[0]),
    last: metaNormalizeNamePart_(parts.slice(1).join(' '))
  };
}

/** Match Metabase_Expert_Country_List: trim, case-insensitive. */
function metaExpertsKey_(first, last) {
  return String(first || '').trim().toLowerCase() + '|' + String(last || '').trim().toLowerCase();
}

function metaNormalizeCountry_(s) {
  return String(s || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

/** Hubstaff column H "tracked" is in seconds → hours. */
function metaTrackedToHours_(val) {
  var n = parseFloat(val);
  if (isNaN(n) || n <= 0) return 0;
  return n / 3600;
}

function metaParseMoney_(cell) {
  if (cell == null || cell === '') return null;
  if (typeof cell === 'number' && !isNaN(cell)) return cell;
  var s = String(cell).replace(/[$,\s]/g, '');
  var n = parseFloat(s);
  return isNaN(n) ? null : n;
}

/** Split "Project | Sub" — if no pipe, sub is empty and full string is project. */
function metaParseProject_(raw) {
  var s = String(raw || '').trim();
  var pipe = s.indexOf('|');
  if (pipe === -1) return { projectName: s, subProject: '' };
  return {
    projectName: s.substring(0, pipe).trim(),
    subProject: s.substring(pipe + 1).trim()
  };
}

/**
 * Canonical project key for all aggregations and pay: text after " - " when present
 * (e.g. "Redwood - ATP Evals" → "ATP Evals"); otherwise the trimmed segment before "|".
 */
function metaProjectKey_(projectNameBeforePipe) {
  var s = String(projectNameBeforePipe || '').trim() || '(no project)';
  var marker = ' - ';
  var idx = s.indexOf(marker);
  if (idx === -1) return s;
  var rest = s.substring(idx + marker.length).trim();
  return rest || s;
}

/** Single label for all meeting + training phrasing (ampersand / and / plural variants). */
var META_MEETINGS_TRAININGS_LABEL = 'Meetings & Trainings';

/**
 * Collapse "Meetings & Training", "Meetings and training", typos, etc. into one bucket.
 */
function metaCanonicalMeetingsLabel_(sub) {
  var t = String(sub || '').trim();
  if (!t) return t;
  var compact = t
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[–—]/g, '-');
  var hasMeet = /\bmeetings?\b/.test(compact);
  var hasTrain = /\btrain(ing|ings)?\b/.test(compact);
  if (hasMeet && hasTrain) return META_MEETINGS_TRAININGS_LABEL;
  return t;
}

/**
 * Canonical sub key: placeholder subs → suffix from project when "A - B" pattern; else use same label as project (metaProjectKey_).
 * Meeting/training phrasing → META_MEETINGS_TRAININGS_LABEL.
 */
function metaSubKey_(subRaw, projectNameBeforePipe) {
  var proj = String(projectNameBeforePipe || '').trim();
  var projectKey = metaProjectKey_(proj);
  var sub = String(subRaw || '').trim();
  var isDashPlaceholder =
    sub === '' ||
    sub === '-' ||
    sub === '—' ||
    sub === '\u2013' ||
    sub === '\u2014';

  if (isDashPlaceholder) {
    sub = projectKey;
  }

  sub = metaCanonicalMeetingsLabel_(sub);

  if (sub === '—') sub = projectKey;

  return sub || projectKey;
}

/** Apply once per Hubstaff row before any hour aggregation. */
function metaCanonicalizeHubstaffRow_(parsed) {
  var proj = String(parsed.projectName || '').trim() || '(no project)';
  var projectKey = metaProjectKey_(proj);
  var subKey = metaSubKey_(parsed.subProject, proj);
  return { projectKey: projectKey, subKey: subKey };
}

function metaReadExpertsMap_(ss) {
  var sh = ss.getSheetByName(META_SHEET_EXPERTS);
  if (!sh) return null;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return {};
  var b = sh.getRange(2, 2, lastRow, 2).getValues();
  var c = sh.getRange(2, 3, lastRow, 3).getValues();
  var f = sh.getRange(2, 6, lastRow, 6).getValues();
  var map = {};
  for (var i = 0; i < b.length; i++) {
    var first = metaNormalizeNamePart_(b[i][0]);
    var last = metaNormalizeNamePart_(c[i][0]);
    if (!first || !last) continue;
    var key = metaExpertsKey_(first, last);
    var country = String(f[i][0] || '').trim();
    map[key] = country;
  }
  return map;
}

/**
 * Normalized email (col B) → country (col D). Missing sheet → {}.
 */
function metaReadAirtableEmailCountryMap_(ss) {
  var sh = ss.getSheetByName(META_SHEET_AIRTABLE_EXPERTS);
  if (!sh) return {};
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return {};
  var emails = sh.getRange(2, 2, lastRow, 2).getValues();
  var countries = sh.getRange(2, 4, lastRow, 4).getValues();
  var map = {};
  for (var i = 0; i < emails.length; i++) {
    var em = META_normalizeEmail_(emails[i][0]);
    if (!em) continue;
    var ctry = String(countries[i][0] || '').trim();
    if (ctry) map[em] = ctry;
  }
  return map;
}

/**
 * Metabase (first|last) first; Airtable (email) second. If both set and differ, Airtable wins.
 */
function metaMergeExpertCountries_(countryFromMeta, countryFromAirtable) {
  var m = String(countryFromMeta || '').trim();
  var a = String(countryFromAirtable || '').trim();
  if (!a) return m;
  if (!m) return a;
  if (metaNormalizeCountry_(m) === metaNormalizeCountry_(a)) return m;
  return a;
}

function metaReadRateCardMatrix_(ss) {
  var sh = ss.getSheetByName(META_SHEET_RATE_CARD);
  if (!sh) return null;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  return sh.getRange(2, 1, lastRow, 24).getValues();
}

function metaFindRateForCountry_(matrix, rateType, countryRaw) {
  var cols = META_GEO_COLS[rateType];
  if (!cols || !matrix || !matrix.length) return null;
  var want = metaNormalizeCountry_(countryRaw);
  if (!want) return null;
  var cIdx = cols.country - 1;
  var rIdx = cols.rate - 1;
  for (var i = 0; i < matrix.length; i++) {
    var row = matrix[i];
    if (!row) continue;
    var cellCountry = metaNormalizeCountry_(row[cIdx]);
    if (cellCountry && cellCountry === want) {
      return metaParseMoney_(row[rIdx]);
    }
  }
  return null;
}

/**
 * @returns {{
 *   error?: string,
 *   hubstaffRows: Array,
 *   expertsMap: Object,
 *   rateMatrix: Array,
 *   byUser: Array,
 *   projectTotals: Array,
 *   subProjectTotals: Array,
 *   uniqueProjects: Array,
 *   agentsByProject: Object,
 *   expertResolution: Array,
 *   sampleRows: Array
 * }}
 */
function metaAggregateCore_(ss) {
  var hub = ss.getSheetByName(META_SHEET_HUBSTAFF);
  if (!hub) return { error: 'Sheet "' + META_SHEET_HUBSTAFF + '" not found' };

  var expertsMap = metaReadExpertsMap_(ss);
  if (expertsMap === null) return { error: 'Sheet "' + META_SHEET_EXPERTS + '" not found' };

  var rateMatrix = metaReadRateCardMatrix_(ss);
  if (rateMatrix === null) return { error: 'Sheet "' + META_SHEET_RATE_CARD + '" not found' };

  var airtableByEmail = metaReadAirtableEmailCountryMap_(ss);

  var lastRow = hub.getLastRow();
  var hubstaffDataRows = lastRow >= 2 ? lastRow - 1 : 0;
  if (lastRow < 2) {
    return {
      hubstaffRows: [],
      hubstaffDataRows: hubstaffDataRows,
      expertsMap: expertsMap,
      airtableByEmail: airtableByEmail,
      rateMatrix: rateMatrix,
      byUser: [],
      projectTotals: [],
      subProjectTotals: [],
      uniqueProjects: [],
      agentsByProject: {},
      expertResolution: [],
      sampleRows: []
    };
  }

  // B..Q contiguous: B=0, H=6, P=14, Q=15
  var grid = hub.getRange(2, 2, lastRow, 17).getValues();

  var tripleMap = {};
  var projHours = {};
  var subKeyHours = {};
  var agentsByProject = {};

  function addTriple(email, projKey, subKey, hrs) {
    var ek = META_normalizeEmail_(email);
    if (!ek) return;
    var pk = String(projKey || '').trim() || '(no project)';
    var sk = String(subKey || '').trim() || pk;
    var tkey = ek + '\t' + pk + '\t' + sk;
    if (!tripleMap[tkey]) tripleMap[tkey] = 0;
    tripleMap[tkey] += hrs;

    if (!projHours[pk]) projHours[pk] = 0;
    projHours[pk] += hrs;

    var subAggKey = pk + '\t' + sk;
    if (!subKeyHours[subAggKey]) subKeyHours[subAggKey] = 0;
    subKeyHours[subAggKey] += hrs;

    if (!agentsByProject[pk]) agentsByProject[pk] = {};
    if (!agentsByProject[pk][ek]) agentsByProject[pk][ek] = 0;
    agentsByProject[pk][ek] += hrs;
  }

  var sampleRows = [];
  var i;
  for (i = 0; i < grid.length; i++) {
    var row = grid[i];
    var dateRaw = row[0];
    var trackedRaw = row[6];
    var emailRaw = row[14];
    var projectRaw = row[15];

    var hrs = metaTrackedToHours_(trackedRaw);
    var parsed = metaParseProject_(projectRaw);
    var canon = metaCanonicalizeHubstaffRow_(parsed);
    var em = META_normalizeEmail_(emailRaw);

    if (sampleRows.length < 300) {
      sampleRows.push({
        sheetRow: i + 2,
        date: dateRaw instanceof Date ? dateRaw.toISOString() : String(dateRaw || ''),
        email: em,
        trackedRaw: trackedRaw,
        hours: Math.round(hrs * 10000) / 10000,
        projectName: canon.projectKey,
        subProject: canon.subKey,
        rawProject: String(projectRaw || '')
      });
    }

    if (!em || hrs <= 0) continue;
    addTriple(em, canon.projectKey, canon.subKey, hrs);
  }

  var byUserMap = {};
  for (var tk in tripleMap) {
    if (!Object.prototype.hasOwnProperty.call(tripleMap, tk)) continue;
    var parts = tk.split('\t');
    var em = parts[0];
    var pk = parts[1];
    var sk = parts[2];
    if (!byUserMap[em]) byUserMap[em] = [];
    byUserMap[em].push({
      projectName: pk,
      subProject: sk,
      hours: Math.round(tripleMap[tk] * 10000) / 10000
    });
  }

  var byUser = [];
  var emails = Object.keys(byUserMap).sort();
  for (i = 0; i < emails.length; i++) {
    var eml = emails[i];
    byUser.push({ email: eml, projects: byUserMap[eml] });
  }

  var projectTotals = [];
  var pkeys = Object.keys(projHours).sort();
  for (i = 0; i < pkeys.length; i++) {
    var p = pkeys[i];
    projectTotals.push({ projectName: p, hours: Math.round(projHours[p] * 10000) / 10000 });
  }

  var subProjectTotals = [];
  var sks = Object.keys(subKeyHours).sort();
  for (i = 0; i < sks.length; i++) {
    var sk0 = sks[i];
    var bits = sk0.split('\t');
    subProjectTotals.push({
      projectName: bits[0],
      subProject: bits[1] != null && bits[1] !== '' ? bits[1] : bits[0],
      hours: Math.round(subKeyHours[sk0] * 10000) / 10000
    });
  }

  var uniqueProjects = pkeys.slice();

  var expertResolution = [];
  var seenE = {};
  for (i = 0; i < emails.length; i++) {
    var em2 = emails[i];
    if (seenE[em2]) continue;
    seenE[em2] = true;
    var fl = metaEmailToFirstLast_(em2);
    var key = metaExpertsKey_(fl.first, fl.last);
    var countryMeta = expertsMap[key] || '';
    var countryAir = airtableByEmail[em2] || '';
    var country = metaMergeExpertCountries_(countryMeta, countryAir);
    expertResolution.push({
      email: em2,
      firstName: fl.first,
      lastName: fl.last,
      country: country || '',
      status: country ? 'matched' : 'no_expert_row'
    });
  }
  expertResolution.sort(function (a, b) {
    if (a.status !== b.status) return a.status === 'no_expert_row' ? -1 : 1;
    return a.email.localeCompare(b.email);
  });

  return {
    hubstaffRows: grid,
    hubstaffDataRows: hubstaffDataRows,
    expertsMap: expertsMap,
    airtableByEmail: airtableByEmail,
    rateMatrix: rateMatrix,
    byUser: byUser,
    projectTotals: projectTotals,
    subProjectTotals: subProjectTotals,
    uniqueProjects: uniqueProjects,
    agentsByProject: agentsByProject,
    expertResolution: expertResolution,
    sampleRows: sampleRows
  };
}

/**
 * Main load for Meta_Index. Managers only.
 */
function getMetaDashboardData() {
  var email = '';
  try {
    try {
      var raw = Session.getActiveUser().getEmail();
      email = raw ? META_normalizeEmail_(raw) : '';
    } catch (e0) {
      email = '';
    }
    if (!email) {
      return { noAccess: true, noAccessReason: 'no_session_email' };
    }
    if (!META_isManager()) {
      return {
        noAccess: true,
        noAccessReason: 'not_on_list',
        recognizedEmail: email
      };
    }

    var ss = META_getSpreadsheet_();
    var agg = metaAggregateCore_(ss);
    if (agg.error) return { error: agg.error };

    var resList = agg.expertResolution || [];
    var matched = 0;
    for (var ri = 0; ri < resList.length; ri++) {
      if (resList[ri].status === 'matched') matched++;
    }

    return {
      viewerEmail: email,
      rateLabels: META_RATE_LABELS,
      projectTotals: agg.projectTotals,
      subProjectTotals: agg.subProjectTotals,
      uniqueProjects: agg.uniqueProjects,
      agentsByProject: agg.agentsByProject,
      expertResolution: agg.expertResolution,
      hubstaffDataRows: agg.hubstaffDataRows != null ? agg.hubstaffDataRows : 0,
      uniqueAgents: resList.length,
      uniqueAgentsWithCountry: matched
    };
  } catch (err) {
    Logger.log('getMetaDashboardData ERROR ' + email + ': ' + (err && err.message ? err.message : err));
    return { error: String(err && err.message ? err.message : err) };
  }
}

/**
 * @param {Object} projectRates — projectName → { rateType: 1-6, customHourly?: number }
 * @param {Object} manualAgentRates — email → { hourlyRate: number } (applied when geo lookup fails)
 */
function metaCalculatePay(projectRates, manualAgentRates) {
  try {
    if (!META_isManager()) return { error: 'Access denied' };
    projectRates = projectRates || {};
    manualAgentRates = manualAgentRates || {};

    var ss = META_getSpreadsheet_();
    var agg = metaAggregateCore_(ss);
    if (agg.error) return { error: agg.error };

    var expertsMap = agg.expertsMap;
    var airtableByEmail = agg.airtableByEmail || {};
    var rateMatrix = agg.rateMatrix;

    var manualNorm = {};
    for (var mk in manualAgentRates) {
      if (Object.prototype.hasOwnProperty.call(manualAgentRates, mk)) {
        manualNorm[META_normalizeEmail_(mk)] = manualAgentRates[mk];
      }
    }

    function countryForEmail(em) {
      var emn = META_normalizeEmail_(em);
      var fl = metaEmailToFirstLast_(em);
      var key = metaExpertsKey_(fl.first, fl.last);
      return metaMergeExpertCountries_(expertsMap[key] || '', airtableByEmail[emn] || '');
    }

    var totalPay = 0;
    var byProject = [];
    var issues = [];
    var traceRows = [];
    var projects = Object.keys(agg.agentsByProject);

    for (var pi = 0; pi < projects.length; pi++) {
      var projName = projects[pi];
      var choice = projectRates[projName];
      if (!choice || !choice.rateType) {
        issues.push({
          projectName: projName,
          email: '',
          reason: 'No rate card selected for this project',
          hours: 0
        });
        continue;
      }

      var rateType = parseInt(choice.rateType, 10);
      var customHourly = parseFloat(choice.customHourly);
      var agentMap = agg.agentsByProject[projName];
      var emails = Object.keys(agentMap);
      var projPay = 0;

      if (rateType === 6) {
        if (isNaN(customHourly) || customHourly < 0) {
          issues.push({
            projectName: projName,
            email: '',
            reason: 'Custom rate missing or invalid',
            hours: 0
          });
          continue;
        }
        var sumH = 0;
        for (var ei = 0; ei < emails.length; ei++) {
          sumH += agentMap[emails[ei]];
        }
        projPay = sumH * customHourly;
        totalPay += projPay;
        traceRows.push({
          projectName: projName,
          email: '(all agents)',
          hours: Math.round(sumH * 10000) / 10000,
          rateType: 6,
          rateTypeLabel: META_RATE_LABELS[6],
          hourlyRate: customHourly,
          pay: Math.round(projPay * 100) / 100,
          note: 'Custom — flat hourly × combined hours'
        });
        byProject.push({
          projectName: projName,
          rateType: 6,
          rateLabel: META_RATE_LABELS[6],
          pay: Math.round(projPay * 100) / 100,
          detail: traceRows.filter(function (r) {
            return r.projectName === projName;
          })
        });
        continue;
      }

      if (rateType < 1 || rateType > 5) {
        issues.push({
          projectName: projName,
          email: '',
          reason: 'Invalid rate type',
          hours: 0
        });
        continue;
      }

      for (var ej = 0; ej < emails.length; ej++) {
        var em = emails[ej];
        var hrs = agentMap[em];
        var manualHr = manualNorm[em];
        var manualNum =
          manualHr && manualHr.hourlyRate != null && !isNaN(parseFloat(manualHr.hourlyRate))
            ? parseFloat(manualHr.hourlyRate)
            : null;

        var ctry = countryForEmail(em);
        var tableRate = ctry ? metaFindRateForCountry_(rateMatrix, rateType, ctry) : null;
        var geoFailed = !ctry || tableRate == null;

        var hourly = null;
        var note = '';
        if (manualNum != null && geoFailed) {
          hourly = manualNum;
          note = 'Manual override';
        } else if (tableRate != null) {
          hourly = tableRate;
          note = 'Geo table - ' + ctry;
        }

        if (hourly == null) {
          if (!ctry) {
            issues.push({
              projectName: projName,
              email: em,
              reason: 'No country in Metabase or Airtable expert lists for this email — add manual $/hr',
              hours: hrs
            });
          } else {
            issues.push({
              projectName: projName,
              email: em,
              reason: 'No rate for country "' + ctry + '" in ' + META_RATE_LABELS[rateType] + ' — add manual $/hr',
              hours: hrs
            });
          }
          continue;
        }

        var pay = hrs * hourly;
        projPay += pay;
        traceRows.push({
          projectName: projName,
          email: em,
          hours: Math.round(hrs * 10000) / 10000,
          rateType: rateType,
          rateTypeLabel: META_RATE_LABELS[rateType],
          hourlyRate: Math.round(hourly * 10000) / 10000,
          pay: Math.round(pay * 100) / 100,
          note: note
        });
      }

      totalPay += projPay;
      byProject.push({
        projectName: projName,
        rateType: rateType,
        rateLabel: META_RATE_LABELS[rateType],
        pay: Math.round(projPay * 100) / 100,
        detail: traceRows.filter(function (r) {
          return r.projectName === projName && r.rateType !== 6;
        })
      });
    }

    return {
      totalPay: Math.round(totalPay * 100) / 100,
      byProject: byProject,
      issues: issues,
      traceRows: traceRows
    };
  } catch (err) {
    return { error: String(err && err.message ? err.message : err) };
  }
}
