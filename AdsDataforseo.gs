/**
 * PAID ADS VIA DATAFORSEO - drop-in replacement for the Zenserp script (Multi Functions)
 *
 * Same job, same schedule, same output as Zenserp. Only the provider changes:
 *   - every keyword in 'Final keywords', mobile and desktop, Auckland, twice daily (7am / 7pm)
 *   - DataForSEO SERP live/advanced; every 'paid' item is one row, tagged top_ads / bottom_ads
 *   - same 8 columns and the same 'No Ads Found' / 'Network Error' / 'HTTP xxx' /
 *     'Rate Limited (429)' rows as Zenserp, so Combined Data needs no changes
 *     (DataForSEO account errors are written as 'API error <code>: <message>')
 *   - one keyword per API call (as DataForSEO requires for live calls)
 *   - each keyword is looked at samplesPerKeyword (3) times per device, a few seconds apart,
 *     because Google does not show ads on every page load. An ad seen in ANY look is written;
 *     'No Ads Found' only when every look succeeded and none had ads
 *   - batches of 25 keywords, mobile -> desktop, watchdog, retries and sheet lock
 *
 * CHECK ONE KEYWORD: set testKeyword below and run testAdsDfsKeyword(). It logs each look and
 * the rows the script would write. Writes nothing.
 *
 * HOW TO USE
 *   Test:     outputSheets = DFS tabs (default). Run startAdsDfsNow(), then compareAdsParity().
 *   Go live:  set outputSheets to the Zen tab names, run removeLegacyZenserpTriggers(),
 *             then setupAdsDfsSchedule().
 *
 * CREDENTIALS: Script Properties DFS_LOGIN / DFS_PASSWORD if set, otherwise config.gs.
 */

var ADS_DFS = {
  inputSheetName : 'Final keywords',
  firstRow       : 2,
  lastRow        : 1001,        // stops earlier if the sheet has fewer keywords
  batchSize      : 25,          // keywords per batch (one call takes ~8s on average, up to ~16s)
  samplesPerKeyword: 3,         // looks per keyword per device; ads rotate between page loads
  parallelRequests: 5,          // keywords fetched at the same time (DataForSEO allows up to 30)
  liveUrl        : 'https://api.dataforseo.com/v3/serp/google/organic/live/advanced',
  locationCode   : 1011036,     // Auckland, New Zealand
  languageCode   : 'en',
  seDomain       : 'google.co.nz',
  depth          : 10,

  maxExecutionTimeMs : 270000,  // stop starting new calls after 4.5 min (limit is 6 min; a live call can take up to 2 min)
  fullAutomationHours: [7, 19], // 7am and 7pm (script time zone)
  maxRunHours        : 8,
  maxBatchRetries    : 2,
  maxRequestRetries  : 2,       // in-batch retries for a failed keyword
  lockTimeoutMs      : 30000,
  rate429DelayMs     : 5000,

  // Test: DFS tabs. Go live: 'AdsResultsZenMobile' / 'AdsResultsZenDesktop'
  outputSheets: { mobile: 'AdsResultsDFSMobile', desktop: 'AdsResultsDFSDesktop' },

  costPerLiveCallUsd : 0.002,   // measured on this account, 2026-09-30

  testKeyword        : 'car insurance nz'   // used by testAdsDfsKeyword()
};

var ADS_DFS_HEADER = ['Timestamp', 'Keyword', 'Ad Type', 'Ad Position',
                      'Ad Title', 'Displayed Link', 'Ad Link', 'Ad Snippet'];

// Ad Title values for non-ad rows (same wording as the Zenserp script)
var ADS_DFS_NO_ADS  = 'No Ads Found';
var ADS_DFS_NET_ERR = 'Network Error';
var ADS_DFS_429     = 'Rate Limited (429)';

var ADS_DFS_HANDLER  = 'runAdsDfsAutomation';
var ADS_DFS_STARTER  = 'startAdsDfsAutomation';
var ADS_DFS_WATCHDOG = 'checkAdsDfsAbandonedFlags';


// ==========================================================
// SETUP / TEAR-DOWN
// ==========================================================

function setupAdsDfsSchedule() {
  removeAdsDfsSchedules();

  ADS_DFS.fullAutomationHours.forEach(function (hour) {
    ScriptApp.newTrigger(ADS_DFS_STARTER).timeBased().everyDays(1).atHour(hour).nearMinute(0).create();
    Logger.log('✓ DFS ads run scheduled daily at ' + hour + ':00');
  });
  ScriptApp.newTrigger(ADS_DFS_WATCHDOG).timeBased().everyHours(1).create();
  Logger.log('✓ Watchdog scheduled hourly');
  adsDfsCostEstimate();
}

function removeAdsDfsSchedules() {
  var deleted = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var h = t.getHandlerFunction();
    if (h === ADS_DFS_HANDLER || h === ADS_DFS_STARTER || h === ADS_DFS_WATCHDOG) {
      ScriptApp.deleteTrigger(t);
      deleted++;
    }
  });
  PropertiesService.getScriptProperties().setProperty('adsDfsRunning', 'false');
  Logger.log('Removed ' + deleted + ' DFS ads trigger(s).');
}

/** Go-live step: deletes every trigger for the old Zenserp handlers. */
function removeLegacyZenserpTriggers() {
  var legacy = ['runZenserpAutomation', 'startZenserpAutomation', 'checkAbandonedFlags',
                'getZenserpAdsMobile', 'getZenserpAdsDesktop'];
  var deleted = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (legacy.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
      deleted++;
    }
  });
  PropertiesService.getScriptProperties().setProperty('zenserpRunning', 'false');
  Logger.log('Removed ' + deleted + ' legacy Zenserp trigger(s).');
}

function adsDfsCostEstimate() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADS_DFS.inputSheetName);
  var keywords = Math.max(0, adsDfsLastRow_(sheet) - ADS_DFS.firstRow + 1);
  var callsPerRun = keywords * 2 * ADS_DFS.samplesPerKeyword;
  var perRun = callsPerRun * ADS_DFS.costPerLiveCallUsd;
  Logger.log(keywords + ' keywords × 2 devices × ' + ADS_DFS.samplesPerKeyword + ' looks = ' + callsPerRun + ' calls per run');
  Logger.log('Per run: $' + perRun.toFixed(2) + ' | per day: $' +
             (perRun * ADS_DFS.fullAutomationHours.length).toFixed(2));
}


// ==========================================================
// START / STOP / WATCHDOG
// ==========================================================

function startAdsDfsNow() {
  PropertiesService.getScriptProperties().setProperty('adsDfsRunning', 'false');
  deleteAdsDfsBatchTriggers_();
  startAdsDfsAutomation();
}

function startAdsDfsAutomation() {
  var props = PropertiesService.getScriptProperties();

  if (props.getProperty('adsDfsRunning') === 'true') {
    var startDate = new Date(props.getProperty('adsDfsStartTime') || '');
    if (!isNaN(startDate.getTime()) && (new Date() - startDate) / 3600000 <= ADS_DFS.maxRunHours) {
      Logger.log('DFS ads run already in progress. Skipping.');
      return;
    }
    Logger.log('⚠️ Stale DFS ads run. Auto-recovering.');
    deleteAdsDfsBatchTriggers_();
  }

  props.setProperty('adsDfsRunning', 'true');
  props.setProperty('adsDfsDevice', 'mobile');
  props.setProperty('adsDfsBatchIndex', '0');
  props.setProperty('adsDfsOffset', '0');
  props.setProperty('adsDfsStartTime', new Date().toISOString());
  props.setProperty('adsDfsRetryCount', '0');
  deleteAdsDfsBatchTriggers_();

  Logger.log('================ STARTING DFS ADS RUN ================');

  try {
    ScriptApp.newTrigger(ADS_DFS_HANDLER).timeBased().after(60 * 1000).create();
  } catch (e) {
    props.setProperty('adsDfsRunning', 'false');
    props.setProperty('adsDfsStartupFailure', new Date().toISOString() + ': ' + e.message);
    throw e;
  }
}

function stopAdsDfsAutomation() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('adsDfsRunning', 'false');
  props.setProperty('adsDfsStoppedReason', 'manual_' + new Date().toISOString());
  deleteAdsDfsBatchTriggers_();
  Logger.log('DFS ads automation stopped.');
}

function checkAdsDfsAbandonedFlags() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('adsDfsRunning') !== 'true') return;

  var startDate = new Date(props.getProperty('adsDfsStartTime') || '');
  var hours = isNaN(startDate.getTime()) ? null : (new Date() - startDate) / 3600000;

  if (hours === null || hours > ADS_DFS.maxRunHours) {
    Logger.log('⚠️ [WATCHDOG] Clearing stuck DFS ads run.');
    props.setProperty('adsDfsRunning', 'false');
    props.setProperty('adsDfsStoppedReason', 'watchdog_' + new Date().toISOString());
    deleteAdsDfsBatchTriggers_();
  }
}

function deleteAdsDfsBatchTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === ADS_DFS_HANDLER) ScriptApp.deleteTrigger(t);
  });
}

function scheduleNextAdsDfs_(delayMinutes) {
  deleteAdsDfsBatchTriggers_();
  try {
    ScriptApp.newTrigger(ADS_DFS_HANDLER).timeBased().after(delayMinutes * 60 * 1000).create();
  } catch (e) {
    Logger.log('⚠️ Trigger creation failed: ' + e.message);
    Utilities.sleep(5000);
    try {
      ScriptApp.newTrigger(ADS_DFS_HANDLER).timeBased().after((delayMinutes + 1) * 60 * 1000).create();
    } catch (e2) {
      PropertiesService.getScriptProperties().setProperty('adsDfsTriggerFailed', new Date().toISOString());
      Logger.log('✗ Trigger creation failed again: ' + e2.message);
    }
  }
}


// ==========================================================
// ORCHESTRATOR
// ==========================================================

function runAdsDfsAutomation() {
  var executionStart = new Date();
  var props = PropertiesService.getScriptProperties();

  checkAdsDfsAbandonedFlags();
  if (props.getProperty('adsDfsRunning') !== 'true') {
    deleteAdsDfsBatchTriggers_();
    return;
  }

  // Safety net: if Google kills this execution before it finishes, this trigger resumes the
  // run from the last saved keyword. It is replaced by the normal trigger at the end.
  scheduleNextAdsDfs_(7);

  var batches  = adsDfsBuildBatches_();
  var device   = props.getProperty('adsDfsDevice') || 'mobile';
  var batchIdx = parseInt(props.getProperty('adsDfsBatchIndex') || '0', 10);
  var offset   = parseInt(props.getProperty('adsDfsOffset') || '0', 10);

  if (batchIdx >= batches.length) {
    Logger.log('✓ All DFS ads batches complete.');
    props.setProperty('adsDfsRunning', 'false');
    props.setProperty('adsDfsCompletedAt', new Date().toISOString());
    deleteAdsDfsBatchTriggers_();
    return;
  }

  var batch = batches[batchIdx];
  Logger.log('DFS ' + device.toUpperCase() + ' batch ' + (batchIdx + 1) + '/' + batches.length +
             ' (rows ' + batch.start + '-' + batch.end + ')');

  var result;
  try {
    result = getAdsDfsBatch_(device, batch.start, batch.end, batchIdx === 0 && offset === 0, offset, executionStart);
    props.setProperty('adsDfsRetryCount', '0');
  } catch (e) {
    var retryCount = parseInt(props.getProperty('adsDfsRetryCount') || '0', 10);
    if (retryCount >= ADS_DFS.maxBatchRetries) {
      Logger.log('✗ Batch ' + batchIdx + ' (' + device + ') failed ' + (retryCount + 1) + ' times (' + e.message + '). Skipping.');
      props.setProperty('adsDfsLastBatch', device + '_' + batchIdx + '_skipped');
      props.setProperty('adsDfsRetryCount', '0');
      adsDfsAdvance_(props, device, batchIdx);
      scheduleNextAdsDfs_(1);
      return;
    }
    Logger.log('⚠️ Batch error: ' + e.message + ' (retry ' + (retryCount + 1) + '/' + ADS_DFS.maxBatchRetries + ')');
    props.setProperty('adsDfsRetryCount', String(retryCount + 1));
    scheduleNextAdsDfs_(2);
    return;
  }

  if (!result.complete) {
    // Hit the time guard mid-batch: continue from where it stopped
    props.setProperty('adsDfsOffset', String(result.nextOffset));
    scheduleNextAdsDfs_(1);
    return;
  }

  props.setProperty('adsDfsLastBatch', device + '_' + batchIdx + '_ok');
  props.setProperty('adsDfsLastBatchTime', new Date().toISOString());
  adsDfsAdvance_(props, device, batchIdx);
  scheduleNextAdsDfs_(1);
}

/** Batches of 50 rows, only as far as the keywords go. */
function adsDfsBuildBatches_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADS_DFS.inputSheetName);
  var last = Math.min(ADS_DFS.lastRow, adsDfsLastRow_(sheet));
  var batches = [];
  for (var s = ADS_DFS.firstRow; s <= last; s += ADS_DFS.batchSize) {
    batches.push({ start: s, end: Math.min(s + ADS_DFS.batchSize - 1, last) });
  }
  return batches;
}

/** Last row with a keyword in column A (ignores formulas returning ''). */
function adsDfsLastRow_(sheet) {
  if (!sheet || sheet.getLastRow() < ADS_DFS.firstRow) return ADS_DFS.firstRow - 1;
  var vals = sheet.getRange(ADS_DFS.firstRow, 1, sheet.getLastRow() - ADS_DFS.firstRow + 1, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).trim() !== '') return ADS_DFS.firstRow + i;
  }
  return ADS_DFS.firstRow - 1;
}

/** mobile -> desktop for the same batch, then on to the next batch. */
function adsDfsAdvance_(props, device, batchIdx) {
  props.setProperty('adsDfsOffset', '0');
  if (device === 'mobile') {
    props.setProperty('adsDfsDevice', 'desktop');
  } else {
    props.setProperty('adsDfsDevice', 'mobile');
    props.setProperty('adsDfsBatchIndex', String(batchIdx + 1));
  }
}


// ==========================================================
// FETCH ONE BATCH
// ==========================================================

function getAdsDfsBatch_(device, startRow, endRow, clearSheet, offset, executionStart) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inputSheet = ss.getSheetByName(ADS_DFS.inputSheetName);
  if (!inputSheet) throw new Error('Input sheet not found: ' + ADS_DFS.inputSheetName);

  var outName = ADS_DFS.outputSheets[device];
  var outSheet = ss.getSheetByName(outName) || ss.insertSheet(outName);
  var lock = LockService.getScriptLock();
  var props = PropertiesService.getScriptProperties();

  if (clearSheet) adsDfsWrite_(lock, outSheet, [ADS_DFS_HEADER], true);

  var keywords = inputSheet.getRange(startRow, 1, endRow - startRow + 1, 1).getValues()
                           .map(function (r) { return String(r[0] || '').trim(); });
  var headers = adsDfsHeaders_();
  var pos = offset;

  while (pos < keywords.length) {
    if (new Date() - executionStart > ADS_DFS.maxExecutionTimeMs) {
      Logger.log('⚠️ Time guard hit at keyword ' + pos + '/' + keywords.length + '. Will continue.');
      return { complete: false, nextOffset: pos };
    }
    var chunk = keywords.slice(pos, pos + Math.max(1, ADS_DFS.parallelRequests));
    var rows = adsDfsFetchChunk_(chunk, device, headers);
    if (rows.length) adsDfsWrite_(lock, outSheet, rows, false);
    pos += chunk.length;
    props.setProperty('adsDfsOffset', String(pos));   // progress survives a killed execution
  }

  Logger.log('✓ ' + device + ' rows ' + startRow + '-' + endRow + ' done in ' +
             ((new Date() - executionStart) / 1000).toFixed(1) + 's');
  return { complete: true, nextOffset: 0 };
}

/**
 * Looks at each keyword in the chunk samplesPerKeyword times (the looks for one keyword are
 * sequential, a few seconds apart) and returns the combined sheet rows.
 */
function adsDfsFetchChunk_(chunk, device, headers) {
  var looks = chunk.map(function () { return []; });
  for (var n = 0; n < ADS_DFS.samplesPerKeyword; n++) {
    adsDfsFetchOnce_(chunk, device, headers).forEach(function (r, i) { if (r) looks[i].push(r); });
  }
  var rows = [];
  chunk.forEach(function (kw, i) {
    if (!kw) return;
    adsDfsCombine_(kw, looks[i], device).forEach(function (row) { rows.push(row); });
  });
  return rows;
}

/** One look at every keyword in the chunk, in parallel, with retries. Returns parsed results. */
function adsDfsFetchOnce_(chunk, device, headers) {
  var results = new Array(chunk.length);
  var pending = [];
  chunk.forEach(function (kw, i) { if (kw) pending.push(i); });

  for (var attempt = 0; attempt <= ADS_DFS.maxRequestRetries && pending.length; attempt++) {
    if (attempt > 0) Utilities.sleep(ADS_DFS.rate429DelayMs);

    var requests = pending.map(function (i) { return adsDfsRequest_(chunk[i], device, headers); });
    var responses;
    try {
      responses = UrlFetchApp.fetchAll(requests);
    } catch (e) {
      Logger.log('Fetch failed: ' + e.message);
      pending.forEach(function (i) { results[i] = { retry: true, reason: ADS_DFS_NET_ERR }; });
      continue;
    }

    var stillPending = [];
    responses.forEach(function (resp, j) {
      var i = pending[j];
      results[i] = adsDfsParseResponse_(chunk[i], resp.getResponseCode(), resp.getContentText(), adsDfsNow_());
      if (results[i].retry) stillPending.push(i);
    });
    pending = stillPending;
  }
  return results;
}

/**
 * Pure function. Merges the looks for one keyword into sheet rows:
 *   - every ad seen in any successful look (one row per ad type + advertiser, best position)
 *   - otherwise 'No Ads Found' if every look succeeded
 *   - otherwise an error row, so a failed look can never be read as 0 competitors
 */
function adsDfsCombine_(keyword, looks, device) {
  var ok = looks.filter(function (r) { return r && !r.retry && !r.error; });
  var best = {};
  ok.forEach(function (r) {
    r.rows.forEach(function (row) {
      if (!row[2]) return;                                  // the 'No Ads Found' row
      var key = row[2] + '|' + adsDfsHost_(row[5] || row[6]);
      if (!best[key] || row[3] < best[key][3]) best[key] = row;
    });
  });
  var ads = Object.keys(best).map(function (k) { return best[k]; });
  if (ads.length) {
    return ads.sort(function (a, b) {
      return a[2] === b[2] ? a[3] - b[3] : (a[2] === 'top_ads' ? -1 : 1);
    });
  }
  if (looks.length && ok.length === looks.length) {
    return [ok[0].rows[0]];                                 // 'No Ads Found'
  }
  var failed = looks.filter(function (r) { return !r || r.retry || r.error; })[0] || {};
  var reason = failed.retry ? failed.reason : (failed.rows ? failed.rows[0][4] : ADS_DFS_NET_ERR);
  if (typeof Logger !== 'undefined') Logger.log('✗ "' + keyword + '" (' + device + '): ' + reason);
  return [[adsDfsNow_(), keyword, '', '', reason || ADS_DFS_NET_ERR, '', '', '']];
}

/** Host of a URL or displayed link, lowercase, without www. */
function adsDfsHost_(s) {
  var m = String(s || '').trim().toLowerCase().match(/^(?:[a-z]+:\/\/)?([^\/\s›?#:]+)/);
  return m ? m[1].replace(/^www\./, '') : '';
}

function adsDfsNow_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

/** One live task per request - DataForSEO allows only one task per live call. */
function adsDfsRequest_(keyword, device, headers) {
  return {
    url: ADS_DFS.liveUrl,
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: headers,
    payload: JSON.stringify([{
      keyword       : keyword,
      location_code : ADS_DFS.locationCode,
      language_code : ADS_DFS.languageCode,
      device        : device,
      se_domain     : ADS_DFS.seDomain,
      depth         : ADS_DFS.depth
    }])
  };
}

/**
 * Diagnostic: looks at ADS_DFS.testKeyword exactly as a run would (samplesPerKeyword looks per
 * device), logs what DataForSEO returned for each look, then the rows that would be written.
 * Writes nothing.
 */
function testAdsDfsKeyword() {
  var headers = adsDfsHeaders_();
  var kw = ADS_DFS.testKeyword;
  Logger.log('Keyword: "' + kw + '" | location_code ' + ADS_DFS.locationCode);
  ['mobile', 'desktop'].forEach(function (device) {
    var looks = [];
    for (var n = 1; n <= ADS_DFS.samplesPerKeyword; n++) {
      var resp = UrlFetchApp.fetchAll([adsDfsRequest_(kw, device, headers)])[0];
      var body = resp.getContentText();
      var info = '';
      try {
        var task = JSON.parse(body).tasks[0];
        var res = task.result && task.result[0];
        info = 'status ' + task.status_code + ' | types: ' + (res ? (res.item_types || []).join(',') : '-') +
               ' | check_url: ' + (res ? res.check_url : '-');
      } catch (e) { info = 'unparseable response: ' + body.slice(0, 200); }
      var parsed = adsDfsParseResponse_(kw, resp.getResponseCode(), body, adsDfsNow_());
      looks.push(parsed);
      var ads = parsed.rows ? parsed.rows.filter(function (r) { return r[2]; }) : [];
      Logger.log('--- ' + device + ' look ' + n + ': ' +
                 (ads.length ? ads.map(function (r) { return adsDfsHost_(r[5]); }).join(', ')
                             : (parsed.retry ? parsed.reason : parsed.rows[0][4])));
      Logger.log('    ' + info);
    }
    var rows = adsDfsCombine_(kw, looks, device);
    Logger.log('=== ' + device + ' WRITES ' + rows.length + ' row(s):');
    rows.forEach(function (r) { Logger.log('    ' + r.slice(2, 6).join(' | ')); });
  });
}

// DataForSEO status codes worth retrying (docs: appendix/errors)
//   40101 search engine error, 40103 task failed - resubmit, 40202 rate limit per minute,
//   40209 too many simultaneous requests, 5xxxx internal / timeout / service unavailable
var ADS_DFS_RETRY_CODES = [40101, 40103, 40202, 40209];

/**
 * Pure function (no Apps Script services). One row per paid ad, like Zenserp:
 * Ad Type is top_ads or bottom_ads (bottom = after the first organic result).
 * Returns {rows:[...]} or {retry:true, reason} for a transient failure.
 * DataForSEO returns HTTP 200 for almost everything; the real status is status_code
 * at the top level and on each task.
 */
function adsDfsParseResponse_(keyword, httpCode, bodyText, ts) {
  var errRow = function (text) { return { error: true, rows: [[ts, keyword, '', '', text, '', '', '']] }; };

  if (httpCode === 429) return { retry: true, reason: ADS_DFS_429 };
  if (httpCode >= 500)  return { retry: true, reason: 'HTTP ' + httpCode };

  var data;
  try { data = JSON.parse(bodyText); }
  catch (e) { return httpCode === 200 ? { retry: true, reason: ADS_DFS_NET_ERR } : errRow('HTTP ' + httpCode); }

  // Top-level status (e.g. 40100 bad login, 40200 / 40210 no balance, 40203 daily cost limit hit)
  var top = data && data.status_code;
  var task = data && data.tasks && data.tasks[0];
  var code = (top && top !== 20000) ? top : (task ? task.status_code : null);
  var msg  = (top && top !== 20000) ? data.status_message : (task ? task.status_message : '');

  if (code === null) return { retry: true, reason: ADS_DFS_NET_ERR };
  if (code !== 20000) {
    var text = 'API error ' + code + (msg ? ': ' + msg : '');
    if (code === 40202 || code === 40209) return { retry: true, reason: ADS_DFS_429 };
    if (code >= 50000 || ADS_DFS_RETRY_CODES.indexOf(code) !== -1) return { retry: true, reason: text };
    return errRow(text);   // e.g. bad login, no funds, cost limit, invalid field - retrying won't help
  }

  var result = task.result && task.result[0];
  var items = (result && Array.isArray(result.items)) ? result.items : null;
  // No result or an empty page is a failed fetch, not "no ads"
  if (!items || items.length === 0) return { retry: true, reason: ADS_DFS_NET_ERR };

  var firstOrganic = Infinity;
  items.forEach(function (it) {
    if (it && it.type === 'organic' && it.rank_absolute < firstOrganic) firstOrganic = it.rank_absolute;
  });

  var counts = { top_ads: 0, bottom_ads: 0 };
  var rows = [];
  items.forEach(function (ad) {
    if (!ad || ad.type !== 'paid') return;
    var block = ad.rank_absolute > firstOrganic ? 'bottom_ads' : 'top_ads';
    counts[block]++;
    rows.push([
      ts, keyword, block, counts[block],
      ad.title || '',
      ad.breadcrumb || ad.domain || '',
      ad.url || '',
      ad.description || ''
    ]);
  });

  if (!rows.length) rows.push([ts, keyword, '', '', ADS_DFS_NO_ADS, '', '', '']);
  return { rows: rows };
}

function adsDfsHeaders_() {
  var props = PropertiesService.getScriptProperties();
  var login = props.getProperty('DFS_LOGIN')    || DFS_LOGIN;
  var pass  = props.getProperty('DFS_PASSWORD') || DFS_PASSWORD;
  return { 'Authorization': 'Basic ' + Utilities.base64Encode(login + ':' + pass) };
}

function adsDfsWrite_(lock, sheet, rows, clear) {
  var got = false;
  try {
    got = lock.tryLock(ADS_DFS.lockTimeoutMs);
    if (!got) throw new Error('Could not acquire sheet lock');
    if (clear) {
      sheet.clear();
      sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    } else if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
  } finally {
    if (got) lock.releaseLock();
  }
}


// ==========================================================
// TEST ONLY: compare the DFS tabs with the Zen tabs
// ==========================================================

/**
 * Writes a 'Parity Check' tab: per keyword, the number of non-Trade Me ads Zenserp and
 * DataForSEO each found (mobile, desktop) and the latest timestamp on each side.
 */
function compareAdsParity() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var kwSheet = ss.getSheetByName(ADS_DFS.inputSheetName);
  var last = adsDfsLastRow_(kwSheet);
  var keywords = kwSheet.getRange(ADS_DFS.firstRow, 1, last - ADS_DFS.firstRow + 1, 1).getValues()
                        .map(function (r) { return String(r[0] || '').trim(); }).filter(String);

  var zm = adsDfsCount_(ss.getSheetByName('AdsResultsZenMobile'));
  var dm = adsDfsCount_(ss.getSheetByName('AdsResultsDFSMobile'));
  var zd = adsDfsCount_(ss.getSheetByName('AdsResultsZenDesktop'));
  var dd = adsDfsCount_(ss.getSheetByName('AdsResultsDFSDesktop'));

  var out = [['Keyword', 'Zen mobile', 'DFS mobile', 'Zen desktop', 'DFS desktop',
              'Zen last run', 'DFS last run', 'Match (0 vs >0)']];
  var match = 0, diff = 0;
  keywords.forEach(function (kw) {
    var v = [zm[kw], dm[kw], zd[kw], dd[kw]].map(function (x) { return x ? x.n : 'n/a'; });
    var same = (v[0] > 0) === (v[1] > 0) && (v[2] > 0) === (v[3] > 0);
    if (same) match++; else diff++;
    out.push([kw, v[0], v[1], v[2], v[3],
              (zm[kw] || zd[kw] || {}).ts || '', (dm[kw] || dd[kw] || {}).ts || '',
              same ? 'yes' : 'NO']);
  });

  var sheet = ss.getSheetByName('Parity Check') || ss.insertSheet('Parity Check');
  sheet.clear();
  sheet.getRange(1, 1, out.length, out[0].length).setValues(out);
  sheet.setFrozenRows(1);
  Logger.log('Parity: ' + match + ' match, ' + diff + ' differ, of ' + keywords.length);
}

/** keyword -> {n: ads not on trademe.co.nz, ts: latest timestamp} */
function adsDfsCount_(sheet) {
  var map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues().forEach(function (r) {
    var kw = String(r[1] || '').trim();
    if (!kw) return;
    var e = map[kw] || (map[kw] = { n: 0, ts: '' });
    var ts = r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : String(r[0]);
    if (ts > e.ts) e.ts = ts;
    if (!r[2]) return;   // No Ads Found / error rows have no Ad Type
    if (!/trademe\.co\.nz/i.test(String(r[5]) + ' ' + String(r[6]))) e.n++;
  });
  return map;
}


// ==========================================================
// STATUS
// ==========================================================

function checkAdsDfsStatus() {
  var props = PropertiesService.getScriptProperties();
  Logger.log('================ DFS ADS STATUS ================');
  Logger.log('Running: ' + (props.getProperty('adsDfsRunning') === 'true' ? 'YES' : 'NO'));
  Logger.log('Output tabs: ' + ADS_DFS.outputSheets.mobile + ' / ' + ADS_DFS.outputSheets.desktop);
  ['adsDfsDevice', 'adsDfsBatchIndex', 'adsDfsOffset', 'adsDfsLastBatch', 'adsDfsLastBatchTime',
   'adsDfsCompletedAt', 'adsDfsStoppedReason', 'adsDfsStartupFailure', 'adsDfsTriggerFailed']
    .forEach(function (k) {
      var v = props.getProperty(k);
      if (v) Logger.log(k + ': ' + v);
    });
  var handlers = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  Logger.log('Daily starters: ' + handlers.filter(function (h) { return h === ADS_DFS_STARTER; }).length +
             ' (expected ' + ADS_DFS.fullAutomationHours.length + ')');
}
