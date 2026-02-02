/************ CONFIG ************/
const ROSTER_TOKEN = PropertiesService.getScriptProperties().getProperty('AT_Roster_Token');
const AT2_BASE_ID = PropertiesService.getScriptProperties().getProperty('Roster_Base_ID');
const AT2_TABLE_NAME = 'Agent Roster';
const AT2_VIEW_NAME = '[Image] Prod for Tracker Pull'; // optional
const AT2_TARGET_SHEET_NAME = 'Roster'; // different tab
const AT2_ID_COLUMN_HEADER = '_airtable_record_id'; // unique helper column

/****************************************************
 * UTILITIES
 ****************************************************/

/**
 * Normalize Airtable values into safe strings
 */
function normalizeValue(value) {
  if (value === null || value === undefined) return '';

  if (Array.isArray(value)) {
    return value
      .map(v =>
        typeof v === 'object' && v !== null
          ? (v.name || v.id || JSON.stringify(v))
          : String(v)
      )
      .join(', ');
  }

  if (typeof value === 'object') {
    if (value.filename) return value.filename;
    return JSON.stringify(value);
  }

  return String(value);
}


/****************************************************
 * MAIN SYNC FUNCTION
 ****************************************************/

function ATRosterPullOverwrite() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(AT2_TARGET_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(AT2_TARGET_SHEET_NAME);
  }

  /************ FETCH ALL AIRTABLE RECORDS ************/
  let allRecords = [];
  let offset;

  do {
  let url =
    `https://api.airtable.com/v0/${AT2_BASE_ID}/${encodeURIComponent(AT2_TABLE_NAME)}` +
    `?pageSize=100` +
    `&cellFormat=string` +
    `&timeZone=UTC` +
    `&userLocale=en-US`;


    if (AT2_VIEW_NAME) {
      url += `&view=${encodeURIComponent(AT2_VIEW_NAME)}`;
    }

    if (offset) {
      url += `&offset=${offset}`;
    }

    const response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: `Bearer ${ROSTER_TOKEN}`
      },
      muteHttpExceptions: true
    });

    const data = JSON.parse(response.getContentText());

    if (!data.records) {
      throw new Error('Airtable API error: ' + response.getContentText());
    }

    allRecords = allRecords.concat(data.records);
    offset = data.offset;

  } while (offset);

  /************ RESET SHEET ************/
  sheet.clearContents();

  if (allRecords.length === 0) return;

  /************ BUILD HEADERS FROM ALL FIELDS + FILTER ALLOWED COLUMNS ************/
  const fieldSet = new Set();
  allRecords.forEach(record => {
    Object.keys(record.fields || {}).forEach(field => fieldSet.add(field));
  });
  
  const ALLOWED_FIELDS = [
  'Agent Email',
  'Status',
  'P: Static FTE',
  'Programmer',
  'Programming Languages',
  'P: Resident Country',
  'Geographical Region',
  'Role',
  'PQ: Primary Campaign',
  'Historical Campaign List'
];

/************ BUILD HEADERS FROM FILTERED FIELDS ************/
const headers = ALLOWED_FIELDS.slice();
headers.push(AT2_ID_COLUMN_HEADER);

sheet
  .getRange(1, 1, 1, headers.length)
  .setValues([headers]);


  /************ BUILD DATA ROWS ************/
  const rows = allRecords.map(record =>
    headers.map(header => {
      if (header === AT2_ID_COLUMN_HEADER) return record.id;
      return normalizeValue(record.fields[header]);
    })
  );

  /************ WRITE ALL DATA ************/
  sheet
    .getRange(2, 1, rows.length, headers.length)
    .setValues(rows);
}
