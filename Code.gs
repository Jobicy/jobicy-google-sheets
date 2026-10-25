/** Bound Google Apps Script. All writes target the generated "Jobicy Jobs" tab. */
const JOBICY_HEADERS = [
  'ID', 'Title', 'Company', 'Location', 'Type', 'Published at',
  'Salary min', 'Salary max', 'Currency', 'Period', 'Jobicy URL',
  'Status', 'Last seen in feed', 'Status checked at'
];
const JOBICY_API = 'https://jobicy.com/api/jobs';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Jobicy')
    .addItem('Synchronize jobs', 'syncJobicyJobs')
    .addItem('Install hourly synchronization', 'installJobicyTrigger')
    .addItem('Remove synchronization trigger', 'removeJobicyTrigger')
    .addToUi();
}

function installJobicyTrigger() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('Open the bound spreadsheet before installing the trigger.');
  PropertiesService.getScriptProperties().setProperty('JOBICY_SPREADSHEET_ID', spreadsheet.getId());
  removeJobicyTrigger();
  ScriptApp.newTrigger('syncJobicyJobs').timeBased().everyHours(1).create();
}

function removeJobicyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'syncJobicyJobs') ScriptApp.deleteTrigger(trigger);
  });
}

function syncJobicyJobs() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error('A synchronization is already running.');
  try {
    const properties = PropertiesService.getScriptProperties();
    const lastSync = Number(properties.getProperty('JOBICY_LAST_SUCCESS_MS') || 0);
    if (Date.now() - lastSync < 3600000) {
      throw new Error('Wait at least one hour between successful feed synchronizations.');
    }
    const spreadsheetId = properties.getProperty('JOBICY_SPREADSHEET_ID');
    const spreadsheet = spreadsheetId
      ? SpreadsheetApp.openById(spreadsheetId)
      : SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) throw new Error('Open the bound spreadsheet or configure JOBICY_SPREADSHEET_ID.');
    const existingSheet = spreadsheet.getSheetByName('Jobicy Jobs');
    const existing = existingSheet && existingSheet.getLastRow()
      ? existingSheet.getDataRange().getValues() : [];
    if (existing.length && JSON.stringify(existing[0]) !== JSON.stringify(JOBICY_HEADERS)) {
      throw new Error('The generated tab has unexpected columns. Rename it before starting a fresh import.');
    }
    const filters = {};
    ['geo', 'industry', 'tag'].forEach(function (key) {
      const value = (properties.getProperty('JOBICY_' + key.toUpperCase()) || '').trim();
      if (value) filters[key] = value;
    });
    if (filters.tag && (filters.tag.length < 3 || filters.tag.length > 50)) {
      throw new Error('JOBICY_TAG must contain 3–50 characters.');
    }
    const started = Date.now();
    const jobs = jobicyFetchAll_(filters, started);
    const now = new Date().toISOString();
    const rows = jobicyMergeRows_(existing.slice(1), jobs, now);
    if (rows.length > 10000) throw new Error('10,000-row safeguard reached. Archive the generated tab and start a fresh import.');
    for (let offset = 0; offset < rows.length; offset += 100) {
      jobicyCheckTime_(started);
      const batch = rows.slice(offset, offset + 100);
      const statuses = jobicyFetchJson_(JOBICY_API + '/status?ids=' + batch.map(function (row) { return row[0]; }).join(','));
      const statusMap = jobicyValidateStatuses_(statuses, batch.map(function (row) { return row[0]; }));
      batch.forEach(function (row) {
        row[11] = statusMap.get(String(row[0]));
        row[13] = statuses.checkedAt || now;
      });
    }
    // Only write after every page and status batch succeeds.
    const values = [JOBICY_HEADERS].concat(rows.map(function (row) { return row.map(jobicySafeCell_); }));
    const sheet = existingSheet || spreadsheet.insertSheet('Jobicy Jobs');
    if (sheet.getMaxRows() < values.length) sheet.insertRowsAfter(sheet.getMaxRows(), values.length - sheet.getMaxRows());
    if (sheet.getMaxColumns() < JOBICY_HEADERS.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), JOBICY_HEADERS.length - sheet.getMaxColumns());
    sheet.getRange(1, 1, values.length, JOBICY_HEADERS.length).setNumberFormat('@').setValues(values);
    sheet.setFrozenRows(1);
    SpreadsheetApp.flush();
    properties.setProperty('JOBICY_LAST_SUCCESS_MS', String(Date.now()));
    console.log('Synchronized ' + jobs.length + ' feed jobs; retained ' + rows.length + ' records. Jobs powered by https://jobicy.com/');
    return { feedJobs: jobs.length, storedJobs: rows.length };
  } finally {
    lock.releaseLock();
  }
}

function jobicyFetchAll_(filters, started) {
  const jobs = new Map();
  const cursors = new Set();
  let cursor = null;
  for (let page = 0; page < 100; page++) {
    jobicyCheckTime_(started);
    const parameters = Object.assign({ count: 100 }, filters);
    if (cursor) parameters.cursor = cursor;
    const query = Object.keys(parameters).map(function (key) {
      return encodeURIComponent(key) + '=' + encodeURIComponent(parameters[key]);
    }).join('&');
    const payload = jobicyFetchJson_(JOBICY_API + '?' + query);
    if (!Array.isArray(payload.jobs) ||
        !(payload.nextCursor === null || (typeof payload.nextCursor === 'string' && payload.nextCursor.length > 0)) ||
        typeof payload.hasMore !== 'boolean' || payload.hasMore !== (payload.nextCursor !== null)) {
      throw new Error('Invalid jobs response or pagination metadata.');
    }
    payload.jobs.forEach(function (job) {
      if (job && jobicyValidId_(job.id) && typeof job.url === 'string' && /^https:\/\/jobicy\.com\/[^\s]*$/.test(job.url)) {
        jobs.set(String(job.id), job);
      }
    });
    cursor = payload.nextCursor;
    if (cursor === null) return Array.from(jobs.values());
    if (cursors.has(cursor)) throw new Error('Repeated pagination cursor.');
    cursors.add(cursor);
  }
  throw new Error('100-page safeguard reached. The generated tab has not been updated.');
}

function jobicyFetchJson_(url) {
  const response = UrlFetchApp.fetch(url, {
    method: 'get', headers: { Accept: 'application/json' },
    muteHttpExceptions: true, followRedirects: false, validateHttpsCertificates: true
  });
  const status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    const headers = response.getAllHeaders();
    throw new Error('Jobicy HTTP ' + status + '. Retry on a later run. Retry-After: ' + (headers['Retry-After'] || headers['retry-after'] || 'unspecified'));
  }
  let data;
  try { data = JSON.parse(response.getContentText()); }
  catch (_error) { throw new Error('Jobicy returned invalid JSON.'); }
  if (!data || typeof data !== 'object' || data.success === false) throw new Error('Jobicy reported an API error.');
  return data;
}

function jobicyMergeRows_(previous, jobs, now) {
  const rows = new Map();
  previous.forEach(function (row) {
    if (!jobicyValidId_(row[0]) || row.length !== JOBICY_HEADERS.length) throw new Error('Invalid stored row. Restore the generated tab before syncing.');
    rows.set(String(row[0]), row.slice());
  });
  jobs.forEach(function (job) {
    const old = rows.get(String(job.id));
    rows.set(String(job.id), [String(job.id), job.jobTitle || '', job.companyName || '', job.jobGeo || '',
      Array.isArray(job.jobType) ? job.jobType.join(', ') : '', job.pubDate || '',
      job.salaryMin == null ? '' : job.salaryMin, job.salaryMax == null ? '' : job.salaryMax,
      job.salaryCurrency || '', job.salaryPeriod || '', job.url, old ? old[11] : 'unknown', now, old ? old[13] : '']);
  });
  return Array.from(rows.values());
}

function jobicyValidateStatuses_(payload, ids) {
  if (payload.success !== true || payload.count !== ids.length || !Array.isArray(payload.jobs)) throw new Error('Invalid status response.');
  const requested = new Set(ids.map(String));
  const result = new Map();
  payload.jobs.forEach(function (item) {
    const id = String(item && item.id);
    if (!item || !requested.has(id) || result.has(id) || ['active', 'closed', 'unknown'].indexOf(item.status) < 0) throw new Error('Invalid status record.');
    result.set(id, item.status);
  });
  if (result.size !== requested.size) throw new Error('Incomplete status response.');
  return result;
}

function jobicyValidId_(value) {
  const text = String(value);
  return /^[1-9][0-9]*$/.test(text) && Number.isSafeInteger(Number(text));
}

function jobicySafeCell_(value) {
  const text = value == null ? '' : String(value);
  return /^[\s]*[=+\-@]/.test(text) ? "'" + text : text;
}

function jobicyCheckTime_(started) {
  if (Date.now() - started > 240000) throw new Error('Execution-time safeguard reached. Try narrower filters. The generated tab has not been updated.');
}
