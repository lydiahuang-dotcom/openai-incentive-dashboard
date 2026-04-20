/**
 * Hubstaff Meta — Client Charge (standalone Apps Script).
 * Reads Hubstaff_Export; only rows where Col Q contains "Trainer".
 * Country: Metabase_Expert_Country_List (first|last) then Airtable_Expert_Country_List (email col B, country col D); Airtable wins on conflict.
 * Then Country_Clusters (cols A–D) → Client_Charge_RateCard cols I–L by tier row.
 */

/** Manager access: same as Meta — use META_MANAGER_EMAILS + META_isManager() in Meta_Code.js (one list). */

/** Served from Meta_Index.html tabs alongside agent pay; no separate doGet. */

function CC_normalizeEmail_(str) {
  var s = String(str || '').trim();
  var match = s.match(/<([^>]+)>/);
  var emailOnly = match ? match[1].trim() : s;
  return emailOnly.toLowerCase();
}

var CC_SHEET_HUBSTAFF = 'Hubstaff_Export';
var CC_SHEET_EXPERTS = 'Metabase_Expert_Country_List';
var CC_SHEET_CLUSTERS = 'Country_Clusters';
var CC_SHEET_CLIENT_CARD = 'Client_Charge_RateCard';

/** Tier rows on Client_Charge_RateCard */
var CC_ROW_GENERALIST = 10;
var CC_ROW_ADVANCED = 11;
var CC_ROW_EXPERT = 13;
var CC_ROW_CODING_ML = 14;

/** Rate columns I–L (1-based) */
var CC_RATE_COL_START = 9;

function ccNormalizeNamePart_(s) {
  return String(s || '')
    .trim()
    .replace(/\s+/g, ' ');
}

function ccEmailToFirstLast_(email) {
  var e = CC_normalizeEmail_(email);
  var at = e.indexOf('@');
  if (at <= 0) return { first: '', last: '' };
  var local = e.substring(0, at);
  var parts = local.split('.');
  if (parts.length < 2) {
    return { first: ccNormalizeNamePart_(parts[0] || ''), last: '' };
  }
  return {
    first: ccNormalizeNamePart_(parts[0]),
    last: ccNormalizeNamePart_(parts.slice(1).join(' '))
  };
}

function ccExpertsKey_(first, last) {
  return String(first || '').trim().toLowerCase() + '|' + String(last || '').trim().toLowerCase();
}

function ccNormalizeCountry_(s) {
  return String(s || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function ccTrackedToHours_(val) {
  var n = parseFloat(val);
  if (isNaN(n) || n <= 0) return 0;
  return n / 3600;
}

function ccParseMoney_(cell) {
  if (cell == null || cell === '') return null;
  if (typeof cell === 'number' && !isNaN(cell)) return cell;
  var s = String(cell).replace(/[$,\s]/g, '');
  var n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function ccReadExpertsMap_(ss) {
  var sh = ss.getSheetByName(CC_SHEET_EXPERTS);
  if (!sh) return null;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return {};
  var b = sh.getRange(2, 2, lastRow, 2).getValues();
  var c = sh.getRange(2, 3, lastRow, 3).getValues();
  var f = sh.getRange(2, 6, lastRow, 6).getValues();
  var map = {};
  for (var i = 0; i < b.length; i++) {
    var first = ccNormalizeNamePart_(b[i][0]);
    var last = ccNormalizeNamePart_(c[i][0]);
    if (!first || !last) continue;
    var key = ccExpertsKey_(first, last);
    map[key] = String(f[i][0] || '').trim();
  }
  return map;
}

/**
 * Which cluster column (0=A…3=D) contains this country in Country_Clusters.
 */
function ccFindClusterColumnIndex_(countryRaw, ss) {
  var sh = ss.getSheetByName(CC_SHEET_CLUSTERS);
  if (!sh) return -1;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return -1;
  var want = ccNormalizeCountry_(countryRaw);
  if (!want) return -1;
  var data = sh.getRange(2, 1, lastRow, 4).getValues();
  for (var r = 0; r < data.length; r++) {
    for (var c = 0; c < 4; c++) {
      if (data[r][c] == null || data[r][c] === '') continue;
      if (ccNormalizeCountry_(data[r][c]) === want) return c;
    }
  }
  return -1;
}

function ccTierRowForChargeType_(chargeType) {
  var t = String(chargeType || '').toLowerCase();
  if (t === 'generalist') return CC_ROW_GENERALIST;
  if (t === 'advanced') return CC_ROW_ADVANCED;
  if (t === 'expert') return CC_ROW_EXPERT;
  if (t === 'coding_ml' || t === 'coding' || t === 'ml_eng') return CC_ROW_CODING_ML;
  return 0;
}

function ccGetClientChargeRate_(ss, tierRow, clusterIdx) {
  var sh = ss.getSheetByName(CC_SHEET_CLIENT_CARD);
  if (!sh) return null;
  var col = CC_RATE_COL_START + clusterIdx;
  if (clusterIdx < 0 || clusterIdx > 3) return null;
  var v = sh.getRange(tierRow, col, tierRow, col).getValue();
  return ccParseMoney_(v);
}

var CC_CHARGE_LABELS = {
  generalist: 'Generalist',
  advanced: 'Advanced',
  expert: 'Expert',
  coding_ml: 'Coding / ML·ENG',
  custom: 'Custom'
};

/** Cols I–L on Client_Charge_RateCard — cluster A–D. */
function ccGeoTierLabelForCluster_(clusterIdx) {
  if (clusterIdx === 0) return 'Low cost geo';
  if (clusterIdx === 1) return 'Mid cost geo';
  if (clusterIdx === 2) return 'High cost geo';
  if (clusterIdx === 3) return 'USA & Canada';
  return '';
}

/**
 * Billable charge only; skips non-billable trace rows.
 * Custom / discount flat $/hr → customCharge; rate-card geo → I–L buckets; manual unrouted → other.
 */
function ccAggregateChargesByProject_(traceRows) {
  var byName = {};
  var i;
  for (i = 0; i < traceRows.length; i++) {
    var tr = traceRows[i];
    if (String(tr.chargeType || '').toLowerCase() === 'nonbillable') continue;
    var pay = parseFloat(tr.pay);
    if (isNaN(pay) || pay <= 0) continue;
    var pn = String(tr.projectName || '');
    if (!byName[pn]) {
      byName[pn] = {
        projectName: pn,
        totalCharge: 0,
        lowCostGeo: 0,
        midCostGeo: 0,
        highCostGeo: 0,
        usaCanada: 0,
        customCharge: 0,
        other: 0
      };
    }
    var b = byName[pn];
    b.totalCharge += pay;
    var ct = String(tr.chargeType || '').toLowerCase();
    if (ct === 'custom' || ct === 'discount') {
      b.customCharge += pay;
      continue;
    }
    var g = tr.geoClusterIdx;
    if (g === 0) b.lowCostGeo += pay;
    else if (g === 1) b.midCostGeo += pay;
    else if (g === 2) b.highCostGeo += pay;
    else if (g === 3) b.usaCanada += pay;
    else b.other += pay;
  }

  var rows = [];
  for (var k in byName) {
    if (Object.prototype.hasOwnProperty.call(byName, k)) {
      var row = byName[k];
      row.totalCharge = Math.round(row.totalCharge * 100) / 100;
      row.lowCostGeo = Math.round(row.lowCostGeo * 100) / 100;
      row.midCostGeo = Math.round(row.midCostGeo * 100) / 100;
      row.highCostGeo = Math.round(row.highCostGeo * 100) / 100;
      row.usaCanada = Math.round(row.usaCanada * 100) / 100;
      row.customCharge = Math.round(row.customCharge * 100) / 100;
      row.other = Math.round(row.other * 100) / 100;
      rows.push(row);
    }
  }
  rows.sort(function (a, b2) {
    if (b2.totalCharge !== a.totalCharge) return b2.totalCharge - a.totalCharge;
    return String(a.projectName).localeCompare(String(b2.projectName));
  });
  return rows;
}

/**
 * Same Hubstaff row shape as Agent pay: col P project|sub, col H tracked, col Q type.
 * Uses Meta_Code canonical project/sub keys (all timer rows, not Trainer-only).
 */
function ccAggregateForClientCharge_(ss) {
  var hub = ss.getSheetByName(CC_SHEET_HUBSTAFF);
  if (!hub) return { error: 'Sheet "' + CC_SHEET_HUBSTAFF + '" not found' };

  var expertsMap = ccReadExpertsMap_(ss);
  if (expertsMap === null) return { error: 'Sheet "' + CC_SHEET_EXPERTS + '" not found' };

  var airtableByEmail = metaReadAirtableEmailCountryMap_(ss);

  var lastRow = hub.getLastRow();
  var totalExportRows = lastRow >= 2 ? lastRow - 1 : 0;

  var projHours = {};
  var subKeyHours = {};
  var agentsBySubKey = {};
  var rowsWithHours = 0;

  if (lastRow < 2) {
    return {
      expertsMap: expertsMap,
      airtableByEmail: airtableByEmail,
      subProjectTotals: [],
      projectTotals: [],
      uniqueProjects: [],
      agentsBySubKey: {},
      expertResolution: [],
      hubstaffExportRows: 0,
      hubstaffDataRows: 0,
      hubstaffRowsCounted: 0,
      totalExportRows: 0,
      chargeLabels: CC_CHARGE_LABELS
    };
  }

  var grid = hub.getRange(2, 2, lastRow, 17).getValues();
  var emailsSeen = {};
  var i;

  for (i = 0; i < grid.length; i++) {
    var trackedRaw = grid[i][6];
    var emailRaw = grid[i][14];
    var projectRaw = grid[i][15];
    var hrs = metaTrackedToHours_(trackedRaw);
    var parsed = metaParseProject_(projectRaw);
    var canon = metaCanonicalizeHubstaffRow_(parsed);
    var em = META_normalizeEmail_(emailRaw);

    if (!em || hrs <= 0) continue;

    rowsWithHours++;
    var pk = canon.projectKey;
    var sk = canon.subKey;
    var aggKey = pk + '\t' + sk;

    if (!subKeyHours[aggKey]) subKeyHours[aggKey] = 0;
    subKeyHours[aggKey] += hrs;

    if (!projHours[pk]) projHours[pk] = 0;
    projHours[pk] += hrs;

    if (!agentsBySubKey[aggKey]) agentsBySubKey[aggKey] = {};
    if (!agentsBySubKey[aggKey][em]) agentsBySubKey[aggKey][em] = 0;
    agentsBySubKey[aggKey][em] += hrs;

    emailsSeen[em] = true;
  }

  var emails = Object.keys(emailsSeen).sort();
  var expertResolution = [];
  for (i = 0; i < emails.length; i++) {
    var em2 = emails[i];
    var fl = ccEmailToFirstLast_(em2);
    var key = ccExpertsKey_(fl.first, fl.last);
    var countryMeta = expertsMap[key] || '';
    var countryAir = airtableByEmail[em2] || '';
    var country = metaMergeExpertCountries_(countryMeta, countryAir);
    expertResolution.push({
      email: em2,
      firstName: fl.first,
      lastName: fl.last,
      country: country,
      status: country ? 'matched' : 'no_expert_row'
    });
  }
  expertResolution.sort(function (a, b) {
    if (a.status !== b.status) return a.status === 'no_expert_row' ? -1 : 1;
    return a.email.localeCompare(b.email);
  });

  var subProjectTotals = [];
  var sks = Object.keys(subKeyHours).sort();
  for (i = 0; i < sks.length; i++) {
    var ak = sks[i];
    var tabI = ak.indexOf('\t');
    var pk2 = tabI === -1 ? ak : ak.substring(0, tabI);
    var sk2 = tabI === -1 ? '' : ak.substring(tabI + 1);
    subProjectTotals.push({
      projectName: pk2,
      subProject: sk2,
      hours: Math.round(subKeyHours[ak] * 10000) / 10000,
      rowKey: ak
    });
  }

  var projectTotals = [];
  var pkeys = Object.keys(projHours).sort();
  for (i = 0; i < pkeys.length; i++) {
    var pp = pkeys[i];
    projectTotals.push({
      projectName: pp,
      hours: Math.round(projHours[pp] * 10000) / 10000
    });
  }

  return {
    expertsMap: expertsMap,
    airtableByEmail: airtableByEmail,
    subProjectTotals: subProjectTotals,
    projectTotals: projectTotals,
    uniqueProjects: pkeys,
    agentsBySubKey: agentsBySubKey,
    expertResolution: expertResolution,
    hubstaffExportRows: totalExportRows,
    hubstaffDataRows: totalExportRows,
    hubstaffRowsCounted: rowsWithHours,
    totalExportRows: totalExportRows,
    chargeLabels: CC_CHARGE_LABELS
  };
}

function getClientChargeDashboardData() {
  var email = '';
  try {
    try {
      var raw = Session.getActiveUser().getEmail();
      email = raw ? CC_normalizeEmail_(raw) : '';
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
    var agg = ccAggregateForClientCharge_(ss);
    if (agg.error) return { error: agg.error };

    var clustersSh = ss.getSheetByName(CC_SHEET_CLUSTERS);
    var cardSh = ss.getSheetByName(CC_SHEET_CLIENT_CARD);
    if (!clustersSh) return { error: 'Sheet "' + CC_SHEET_CLUSTERS + '" not found' };
    if (!cardSh) return { error: 'Sheet "' + CC_SHEET_CLIENT_CARD + '" not found' };

    var resList = agg.expertResolution || [];
    var matched = 0;
    for (var ri = 0; ri < resList.length; ri++) {
      if (resList[ri].status === 'matched') matched++;
    }

    return {
      viewerEmail: email,
      subProjectTotals: agg.subProjectTotals,
      projectTotals: agg.projectTotals,
      uniqueProjects: agg.uniqueProjects || [],
      expertResolution: agg.expertResolution,
      hubstaffExportRows: agg.hubstaffExportRows,
      hubstaffDataRows: agg.hubstaffDataRows != null ? agg.hubstaffDataRows : agg.hubstaffExportRows,
      hubstaffRowsCounted: agg.hubstaffRowsCounted,
      uniqueAgents: resList.length,
      uniqueAgentsWithCountry: matched,
      chargeOptions: [
        { id: 'generalist', label: CC_CHARGE_LABELS.generalist },
        { id: 'advanced', label: CC_CHARGE_LABELS.advanced },
        { id: 'expert', label: CC_CHARGE_LABELS.expert },
        { id: 'coding_ml', label: CC_CHARGE_LABELS.coding_ml },
        { id: 'custom', label: CC_CHARGE_LABELS.custom }
      ]
    };
  } catch (err) {
    Logger.log('getClientChargeDashboardData ERROR: ' + (err && err.message ? err.message : err));
    return { error: String(err && err.message ? err.message : err) };
  }
}

/**
 * @param {Object} rowCharges — rowKey (project\\t sub) → { billable, chargeType?, customHourly? }
 */
function clientChargeCalculatePay(rowCharges, manualAgentRates) {
  try {
    if (!META_isManager()) return { error: 'Access denied' };
    rowCharges = rowCharges || {};
    manualAgentRates = manualAgentRates || {};

    var ss = META_getSpreadsheet_();
    var agg = ccAggregateForClientCharge_(ss);
    if (agg.error) return { error: agg.error };

    var expertsMap = agg.expertsMap;
    var airtableByEmail = agg.airtableByEmail || {};
    var agentsBySubKey = agg.agentsBySubKey || {};
    var manualNorm = {};
    for (var mk in manualAgentRates) {
      if (Object.prototype.hasOwnProperty.call(manualAgentRates, mk)) {
        manualNorm[CC_normalizeEmail_(mk)] = manualAgentRates[mk];
      }
    }

    function countryForEmail(em) {
      var emn = META_normalizeEmail_(em);
      var fl = ccEmailToFirstLast_(em);
      return metaMergeExpertCountries_(
        expertsMap[ccExpertsKey_(fl.first, fl.last)] || '',
        airtableByEmail[emn] || ''
      );
    }

    var totalCharge = 0;
    var traceRows = [];
    var issues = [];
    var rowKeys = Object.keys(agentsBySubKey);

    for (var ri = 0; ri < rowKeys.length; ri++) {
      var rowKey = rowKeys[ri];
      var choice = rowCharges[rowKey];
      var agentMap = agentsBySubKey[rowKey];
      var tabI = rowKey.indexOf('\t');
      var projName = tabI === -1 ? rowKey : rowKey.substring(0, tabI);
      var subName = tabI === -1 ? '' : rowKey.substring(tabI + 1);

      if (!choice || !choice.billable) {
        issues.push({
          projectName: projName,
          subProject: subName,
          email: '',
          reason: 'Select Billable, Non-billable, or Discount for each row',
          hours: 0
        });
        continue;
      }

      var bill = String(choice.billable).toLowerCase();
      if (bill === 'nonbillable') {
        var sumNb = 0;
        var emNb = Object.keys(agentMap);
        for (var x = 0; x < emNb.length; x++) sumNb += agentMap[emNb[x]];
        traceRows.push({
          projectName: projName,
          subProject: subName,
          email: '—',
          hours: Math.round(sumNb * 10000) / 10000,
          chargeType: 'nonbillable',
          chargeTypeLabel: 'Non-billable',
          hourlyRate: 0,
          pay: 0,
          note: 'Excluded from client charge',
          geoClusterIdx: -1,
          geoTierLabel: ''
        });
        continue;
      }

      if (bill !== 'billable' && bill !== 'discount') {
        issues.push({
          projectName: projName,
          subProject: subName,
          email: '',
          reason: 'Invalid billable value',
          hours: 0
        });
        continue;
      }

      if (bill === 'discount' && String(choice.chargeType || '').toLowerCase() !== 'custom') {
        issues.push({
          projectName: projName,
          subProject: subName,
          email: '',
          reason: 'Discount rows must use Custom client rate',
          hours: 0
        });
        continue;
      }

      if (!choice.chargeType) {
        issues.push({
          projectName: projName,
          subProject: subName,
          email: '',
          reason: 'Select a client charge tier for this row',
          hours: 0
        });
        continue;
      }

      var chargeType = String(choice.chargeType).toLowerCase();
      var customHourly = parseFloat(choice.customHourly);
      var agentEmails = Object.keys(agentMap);
      var rowTotal = 0;

      if (chargeType === 'custom') {
        if (isNaN(customHourly) || customHourly < 0) {
          issues.push({
            projectName: projName,
            subProject: subName,
            email: '',
            reason: 'Custom client rate missing or invalid',
            hours: 0
          });
          continue;
        }
        var sumH = 0;
        for (var ei = 0; ei < agentEmails.length; ei++) {
          sumH += agentMap[agentEmails[ei]];
        }
        rowTotal = sumH * customHourly;
        totalCharge += rowTotal;
        var isDiscount = bill === 'discount';
        traceRows.push({
          projectName: projName,
          subProject: subName,
          email: '(all agents)',
          hours: Math.round(sumH * 10000) / 10000,
          chargeType: isDiscount ? 'discount' : 'custom',
          chargeTypeLabel: isDiscount ? 'Discount (custom)' : CC_CHARGE_LABELS.custom,
          hourlyRate: customHourly,
          pay: Math.round(rowTotal * 100) / 100,
          note: isDiscount
            ? 'Discount — custom $/hr × hours on this project × sub row'
            : 'Custom — flat rate × hours on this project × sub row',
          geoClusterIdx: -1,
          geoTierLabel: isDiscount ? 'Discount' : 'Custom'
        });
        continue;
      }

      var tierRow = ccTierRowForChargeType_(chargeType);
      if (!tierRow) {
        issues.push({
          projectName: projName,
          subProject: subName,
          email: '',
          reason: 'Invalid charge tier',
          hours: 0
        });
        continue;
      }

      var tierLabel = CC_CHARGE_LABELS[chargeType] || chargeType;

      for (var ej = 0; ej < agentEmails.length; ej++) {
        var em = agentEmails[ej];
        var hrs = agentMap[em];
        var manualHr = manualNorm[em];
        var manualNum =
          manualHr && manualHr.hourlyRate != null && !isNaN(parseFloat(manualHr.hourlyRate))
            ? parseFloat(manualHr.hourlyRate)
            : null;

        var ctry = countryForEmail(em);
        var clusterIdx = ctry ? ccFindClusterColumnIndex_(ctry, ss) : -1;
        var clusterLetter = clusterIdx >= 0 && clusterIdx <= 3 ? ['A', 'B', 'C', 'D'][clusterIdx] : '';

        var cardRate =
          ctry && clusterIdx >= 0 && clusterIdx <= 3
            ? ccGetClientChargeRate_(ss, tierRow, clusterIdx)
            : null;
        var geoFailed = !ctry || clusterIdx < 0 || cardRate == null;

        var rate = null;
        var note = '';

        if (manualNum != null && geoFailed) {
          rate = manualNum;
          if (!ctry) note = 'Manual override (no country in Metabase or Airtable expert lists)';
          else if (clusterIdx < 0) note = 'Manual override (country not in Country_Clusters A–D)';
          else note = 'Manual override (no rate in Client_Charge_RateCard)';
        } else if (cardRate != null) {
          rate = cardRate;
          note = 'Client charge — ' + ctry + ' — ' + tierLabel;
        } else {
          if (!ctry) {
            issues.push({
              projectName: projName,
              subProject: subName,
              email: em,
              reason: 'No country in Metabase or Airtable expert lists — add manual client $/hr',
              hours: hrs
            });
          } else if (clusterIdx < 0) {
            issues.push({
              projectName: projName,
              subProject: subName,
              email: em,
              reason: 'Country "' + ctry + '" not in Country_Clusters (cols A–D) — add manual $/hr',
              hours: hrs
            });
          } else {
            issues.push({
              projectName: projName,
              subProject: subName,
              email: em,
              reason: 'No rate in card row ' + tierRow + ' col ' + clusterLetter + ' — add manual $/hr',
              hours: hrs
            });
          }
          continue;
        }

        var pay = hrs * rate;
        rowTotal += pay;
        var geoIdx = clusterIdx >= 0 && clusterIdx <= 3 ? clusterIdx : -1;
        var geoLbl = geoIdx >= 0 ? ccGeoTierLabelForCluster_(geoIdx) : 'Other';
        traceRows.push({
          projectName: projName,
          subProject: subName,
          email: em,
          hours: Math.round(hrs * 10000) / 10000,
          chargeType: chargeType,
          chargeTypeLabel: tierLabel,
          hourlyRate: Math.round(rate * 10000) / 10000,
          pay: Math.round(pay * 100) / 100,
          note: note,
          geoClusterIdx: geoIdx,
          geoTierLabel: geoLbl
        });
      }

      totalCharge += rowTotal;
    }

    return {
      totalCharge: Math.round(totalCharge * 100) / 100,
      traceRows: traceRows,
      issues: issues,
      projectChargeRows: ccAggregateChargesByProject_(traceRows)
    };
  } catch (err) {
    return { error: String(err && err.message ? err.message : err) };
  }
}
