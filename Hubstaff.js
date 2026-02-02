const AIRTABLE_TOKEN = PropertiesService.getScriptProperties().getProperty('AT_Hubstaff_Token');
const BASE_ID = PropertiesService.getScriptProperties().getProperty('HS_Base_ID');
const TABLE_NAME = 'Data';
const VIEW_NAME = '[Image Prod] INV - Bea';
const CREATED_TIME_FIELD = 'Created time';
const TARGET_SHEET_NAME = 'Hubstaff';
const STATE_CELL = 'AB1';
/********************************/

/**
 * Convert Airtable field values into safe strings for Sheets
 */
function normalizeValue(value) {
  if (value === null || value === undefined) return '';

  // Arrays (linked records, multi-select, lookups, collaborators)
  if (Array.isArray(value)) {
    return value
      .map(v =>
        typeof v === 'object' && v !== null
          ? (v.name || v.id || JSON.stringify(v))
          : String(v)
      )
      .join(', ');
  }

  // Objects (attachments, rollups, etc.)
  if (typeof value === 'object') {
    // Attachments → list filenames
    if (value.url || value.filename) {
      return value.filename || value.url;
    }
    return JSON.stringify(value);
  }

  // Primitives
  return String(value);
}

function syncAirtableIncremental() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get or create target sheet
  let sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TARGET_SHEET_NAME);
  }

  // Read last sync timestamp
  const lastSyncValue = sheet.getRange(STATE_CELL).getValue();
  const isFirstRun = !lastSyncValue;
  const lastSyncISO = lastSyncValue
    ? new Date(lastSyncValue).toISOString()
    : null;

  let allRecords = [];
  let offset;

  do {
    let url = `https://api.airtable.com/v0/${BASE_ID}/${encodeURIComponent(TABLE_NAME)}?pageSize=100`;

    if (VIEW_NAME) {
      url += `&view=${encodeURIComponent(VIEW_NAME)}`;
    }

    if (lastSyncISO) {
      const formula = `IS_AFTER({${CREATED_TIME_FIELD}}, "${lastSyncISO}")`;
      url += `&filterByFormula=${encodeURIComponent(formula)}`;
    }

    url += `&sort[0][field]=${encodeURIComponent(CREATED_TIME_FIELD)}`;
    url += `&sort[0][direction]=asc`;

    if (offset) {
      url += `&offset=${offset}`;
    }

    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: `Bearer ${AIRTABLE_TOKEN}` }
    });

    const data = JSON.parse(response.getContentText());
    allRecords = allRecords.concat(data.records);
    offset = data.offset;

  } while (offset);

  if (allRecords.length === 0) return;

  // FIRST RUN: clear + write headers
  let headers;
  if (isFirstRun) {
    sheet.clearContents();
    headers = Object.keys(allRecords[0].fields);
  if (!headers.includes(CREATED_TIME_FIELD)) {
  headers.push(CREATED_TIME_FIELD);
}
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
  }

  // Build rows safely (string-normalized)
 const rows = allRecords.map(record =>
  headers.map(h => {
    if (h === CREATED_TIME_FIELD) {
      // Use Airtable system timestamp
      return record.createdTime
        ? new Date(record.createdTime).toISOString()
        : '';
    }
    return normalizeValue(record.fields[h]);
  })
  );

  // Batch append (FAST)
  const startRow = sheet.getLastRow() + 1;
  sheet
    .getRange(startRow, 1, rows.length, headers.length)
    .setValues(rows);

  // Store newest created time
const newestCreatedTime =
  allRecords[allRecords.length - 1].createdTime;

if (newestCreatedTime) {
  sheet.getRange(STATE_CELL).setValue(new Date(newestCreatedTime));
  }
}