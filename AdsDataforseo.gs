/**
 * PAID ADS VIA DATAFORSEO - SHADOW RUN (replaces Zenserp once parity is proven)
 *
 * Same batch/trigger/watchdog architecture as Multi Functions (Zenserp), but:
 *   - calls DataForSEO SERP live/advanced and keeps 'paid' ads and 'shopping' ads
 *   - DOUBLE-CHECK: every keyword/device is fetched twice (samplesPerKeyword); an advertiser
 *     seen in either snapshot counts, and "No Ads Found" requires BOTH snapshots to succeed
 *     with no ads. Any failure is written as an error row, never as 0 competitors
 *   - one row per advertiser (deduped by domain), Displayed Link = ad domain
 *   - writes to SHADOW tabs (AdsResultsDFSMobile / AdsResultsDFSDesktop) using the
 *     same 8-column schema as the Zenserp tabs, so Combined Data can be re-pointed
 *     later without changing formulas
 *   - uses its own trigger handlers and script properties (adsDfs*), so it never
 *     touches the running Zenserp automation
 *   - resumes mid-batch: progress is saved after every chunk, so the 330s guard
 *     never silently drops keywords
 *
 * HOW TO USE
 *   1. Set ADS_DFS.sampleEndRow (e.g. 101 = first 100 keywords) for the parity test.
 *   2. Run adsDfsCostEstimate() and check it fits under the DataForSEO daily cap.
 *   3. Run startAdsDfsNow() for one shadow run, or setupAdsDfsShadowSchedule() for 7am/7pm.
 *   4. When the run has completed, run compareAdsParity() and review 'Parity Check'.
 *
 * CREDENTIALS: read from Script Properties (DFS_LOGIN / DFS_PASSWORD) if set,
 * otherwise falls back to the constants in config.gs.
 */

var ADS_DFS = {
  inputSheetName : 'Final keywords',
  firstRow       : 2,
  lastRow        : 1001,        // full run covers rows 2-1001 (1,000 keywords)
  sampleEndRow   : 101,         // parity test: set to a row number to cap the run; null = full run
  batchSize      : 50,
  chunkSize      : 10,          // keywords per fetchAll call (× samplesPerKeyword requests)
  samplesPerKeyword: 2,         // double-check: independent SERP snapshots per keyword/device
  liveUrl        : 'https://api.dataforseo.com/v3/serp/google/organic/live/advanced',
  locationCode   : 2554,        // New Zealand
  languageCode   : 'en',
  seDomain       : 'google.co.nz',
  depth          : 10,

  maxExecutionTimeMs : 330000,  // Apps Script limit is 360s
  fullAutomationHours: [7, 19], // 7am and 7pm (script time zone)
  maxRunHours        : 8,
  maxBatchRetries    : 2,
  maxRequestRetries  : 2,       // in-batch retries for a failed keyword
  lockTimeoutMs      : 30000,
  rate429DelayMs     : 5000,
  delayBetweenChunksMs: 200,

  outputSheets: { mobile: 'AdsResultsDFSMobile', desktop: 'AdsResultsDFSDesktop' },

  costPerLiveCallUsd : 0.002,   // measured on this account, 2026-09-30
  dailyCapUsd        : 40       // set to the DataForSEO daily limit on the account
};

var ADS_DFS_HEADER = ['Timestamp', 'Keyword', 'Ad Type', 'Ad Position',
                      'Ad Title', 'Displayed Link', 'Ad Link', 'Ad Snippet'];

// Sentinel values in the Ad Title column (same as the Zenserp script where they overlap)
var ADS_DFS_NO_ADS   = 'No Ads Found';
var ADS_DFS_NET_ERR  = 'Network Error';
var ADS_DFS_SKIPPED  = 'Skipped (invalid keyword)';
var ADS_DFS_EMPTY    = 'Empty SERP (check)';

var ADS_DFS_HANDLER  = 'runAdsDfsAutomation';
var ADS_DFS_STARTER  = 'startAdsDfsAutomation';
var ADS_DFS_WATCHDOG = 'checkAdsDfsAbandonedFlags';


// ==========================================================
// SETUP / TEAR-DOWN
// ==========================================================

function setupAdsDfsShadowSchedule() {
  removeAdsDfsSchedules();

  ADS_DFS.fullAutomationHours.forEach(function (hour) {
    ScriptApp.newTrigger(ADS_DFS_STARTER).timeBased().everyDays(1).atHour(hour).nearMinute(0).create();
    Logger.log('✓ DFS ads run scheduled daily at ' + hour + ':00');
  });
  ScriptApp.newTrigger(ADS_DFS_WATCHDOG).timeBased().everyHours(1).create();
  Logger.log('✓ Watchdog scheduled hourly');
  adsDfsCostEstimate();
}

/** Removes ONLY the DFS shadow triggers. Zenserp triggers are left alone. */
function removeAdsDfsSchedules() {
  var deleted = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var h = t.getHandlerFunction();
    if (h === ADS_DFS_HANDLER || h === ADS_DFS_STARTER || h === ADS_DFS_WATCHDOG) {
      ScriptApp.deleteTrigger(t);
      deleted++;
    }
  });
  var props = PropertiesService.getScriptProperties();
  props.setProperty('adsDfsRunning', 'false');
  Logger.log('Removed ' + deleted + ' DFS ads trigger(s).');
}

/**
 * CUT-OVER ONLY: run this when the parity check has passed to stop Zenserp.
 * Deletes every trigger for the old Zenserp handlers and clears their state.
 */
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
  var lastRow = ADS_DFS.sampleEndRow || ADS_DFS.lastRow;
  var keywords = lastRow - ADS_DFS.firstRow + 1;
  var runsPerDay = ADS_DFS.fullAutomationHours.length;
  var callsPerRun = keywords * 2 * ADS_DFS.samplesPerKeyword;
  var perDay = callsPerRun * runsPerDay * ADS_DFS.costPerLiveCallUsd;
  Logger.log('================ COST ESTIMATE (assumes $' + ADS_DFS.costPerLiveCallUsd + ' per live call) ================');
  Logger.log(keywords + ' keywords × 2 devices × ' + ADS_DFS.samplesPerKeyword + ' snapshots = ' + callsPerRun + ' calls per run');
  Logger.log('Per run: $' + (callsPerRun * ADS_DFS.costPerLiveCallUsd).toFixed(2));
  Logger.log('Per day (' + runsPerDay + ' runs): $' + perDay.toFixed(2) + ' of the $' + ADS_DFS.dailyCapUsd + ' daily cap');
  Logger.log('This is ADDITIONAL to organic + search volume spend.');
  return perDay;
}


// ==========================================================
// START / STOP / WATCHDOG
// ==========================================================

function startAdsDfsNow() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('adsDfsRunning', 'false');
  deleteAdsDfsBatchTriggers_();
  startAdsDfsAutomation();
}

function startAdsDfsAutomation() {
  var props = PropertiesService.getScriptProperties();

  if (props.getProperty('adsDfsRunning') === 'true') {
    var startTime = props.getProperty('adsDfsStartTime');
    var startDate = startTime ? new Date(startTime) : null;
    if (startDate && !isNaN(startDate.getTime()) &&
        (new Date() - startDate) / 3600000 <= ADS_DFS.maxRunHours) {
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
  Logger.log('Batches per device: ' + adsDfsBuildBatches_().length);

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

  var startTime = props.getProperty('adsDfsStartTime');
  var startDate = startTime ? new Date(startTime) : null;
  var hours = (startDate && !isNaN(startDate.getTime())) ? (new Date() - startDate) / 3600000 : null;

  if (hours === null || hours > ADS_DFS.maxRunHours) {
    Logger.log('⚠️ [WATCHDOG] Clearing stuck DFS ads run (' + (hours === null ? 'no valid startTime' : hours.toFixed(1) + 'h') + ').');
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
  delayMinutes = delayMinutes || 1;
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

function adsDfsBuildBatches_() {
  var last = ADS_DFS.sampleEndRow ? Math.min(ADS_DFS.sampleEndRow, ADS_DFS.lastRow) : ADS_DFS.lastRow;
  var batches = [];
  for (var s = ADS_DFS.firstRow; s <= last; s += ADS_DFS.batchSize) {
    batches.push({ start: s, end: Math.min(s + ADS_DFS.batchSize - 1, last) });
  }
  return batches;
}

function runAdsDfsAutomation() {
  var executionStart = new Date();
  var props = PropertiesService.getScriptProperties();

  checkAdsDfsAbandonedFlags();   // guard against orphaned state on every entry

  if (props.getProperty('adsDfsRunning') !== 'true') {
    deleteAdsDfsBatchTriggers_();
    return;
  }

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
             ' (rows ' + batch.start + '-' + batch.end + '), offset ' + offset);

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
    // Hit the time guard mid-batch: resume from the saved offset
    props.setProperty('adsDfsOffset', String(result.nextOffset));
    scheduleNextAdsDfs_(1);
    return;
  }

  props.setProperty('adsDfsLastBatch', device + '_' + batchIdx + '_ok');
  props.setProperty('adsDfsLastBatchTime', new Date().toISOString());
  adsDfsAdvance_(props, device, batchIdx);
  scheduleNextAdsDfs_(1);
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
// FETCH ONE BATCH (resumable)
// ==========================================================

function getAdsDfsBatch_(device, startRow, endRow, clearSheet, offset, executionStart) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inputSheet = ss.getSheetByName(ADS_DFS.inputSheetName);
  if (!inputSheet) throw new Error('Input sheet not found: ' + ADS_DFS.inputSheetName);

  var outName = ADS_DFS.outputSheets[device];
  var outSheet = ss.getSheetByName(outName) || ss.insertSheet(outName);
  var lock = LockService.getScriptLock();

  if (clearSheet) {
    adsDfsWrite_(lock, outSheet, [ADS_DFS_HEADER], true);
  }

  var lastInputRow = inputSheet.getLastRow();
  if (startRow > lastInputRow) return { complete: true, nextOffset: 0 };
  var actualEnd = Math.min(endRow, lastInputRow);

  // Keep blanks in place so offsets stay stable between executions
  var raw = inputSheet.getRange(startRow, 1, actualEnd - startRow + 1, 1).getValues()
                      .map(function (r) { return String(r[0] || '').trim(); });

  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var pos = offset;

  while (pos < raw.length) {
    if (new Date() - executionStart > ADS_DFS.maxExecutionTimeMs) {
      Logger.log('⚠️ Time guard hit at offset ' + pos + '/' + raw.length + '. Will resume.');
      return { complete: false, nextOffset: pos };
    }

    var chunk = raw.slice(pos, pos + ADS_DFS.chunkSize);
    var rows = adsDfsFetchChunk_(chunk, device, ts);
    if (rows.length) adsDfsWrite_(lock, outSheet, rows, false);
    pos += chunk.length;

    if (pos < raw.length) Utilities.sleep(ADS_DFS.delayBetweenChunksMs);
  }

  Logger.log('✓ ' + device + ' rows ' + startRow + '-' + actualEnd + ' done in ' +
             ((new Date() - executionStart) / 1000).toFixed(1) + 's');
  return { complete: true, nextOffset: 0 };
}

/**
 * Fetches a chunk of keywords in parallel; returns sheet rows.
 * Each keyword is fetched ADS_DFS.samplesPerKeyword times (independent SERP snapshots),
 * because ads rotate between page loads and one snapshot can miss a bidding competitor.
 */
function adsDfsFetchChunk_(chunk, device, ts) {
  var headers = adsDfsHeaders_();
  var samples = ADS_DFS.samplesPerKeyword;
  var results = chunk.map(function () { return []; });   // results[i][s] = parsed sample
  var rows = [];

  var pending = [];   // jobs: {i: keyword index, s: sample index}
  chunk.forEach(function (kw, i) {
    if (!kw) return;                                                     // blank cell
    if (!adsDfsIsValidKeyword_(kw)) {
      Logger.log('Skipped invalid keyword: "' + kw + '"');
      rows.push([ts, kw, '', '', ADS_DFS_SKIPPED, '', '', '']);
      return;
    }
    for (var s = 0; s < samples; s++) pending.push({ i: i, s: s });
  });

  for (var attempt = 0; attempt <= ADS_DFS.maxRequestRetries && pending.length; attempt++) {
    if (attempt > 0) Utilities.sleep(ADS_DFS.rate429DelayMs / 2);

    var requests = pending.map(function (job) {
      return {
        url: ADS_DFS.liveUrl,
        method: 'post',
        muteHttpExceptions: true,
        headers: headers,
        payload: JSON.stringify([{
          keyword       : chunk[job.i],
          location_code : ADS_DFS.locationCode,
          language_code : ADS_DFS.languageCode,
          device        : device,
          se_domain     : ADS_DFS.seDomain,
          depth         : ADS_DFS.depth
        }])
      };
    });

    var responses;
    try {
      responses = UrlFetchApp.fetchAll(requests);
    } catch (e) {
      // whole call failed; every pending sample stays retryable
      pending.forEach(function (job) { results[job.i][job.s] = { status: 'retry', reason: ADS_DFS_NET_ERR }; });
      continue;
    }

    var stillPending = [];
    var saw429 = false;
    responses.forEach(function (resp, j) {
      var job = pending[j];
      var parsed = adsDfsParseResponse_(resp.getResponseCode(), resp.getContentText());
      results[job.i][job.s] = parsed;
      if (parsed.status === 'retry') {
        stillPending.push(job);
        if (parsed.reason === 'HTTP 429') saw429 = true;
      }
    });
    if (saw429) Utilities.sleep(ADS_DFS.rate429DelayMs);
    pending = stillPending;
  }

  var disagreed = 0;
  chunk.forEach(function (kw, i) {
    if (!kw || !results[i].length) return;
    var combined = adsDfsCombineSamples_(kw, results[i], ts);
    if (combined.disagreed) disagreed++;
    combined.rows.forEach(function (row) { rows.push(row); });
  });
  if (disagreed) Logger.log('Double-check: snapshots disagreed on ' + disagreed + ' keyword(s) in this chunk');
  return rows;
}

/**
 * Pure function: turns one API response into
 *   {status:'ok', ads:[...]}      SERP fetched (ads may be empty = genuinely no ads)
 *   {status:'empty'}              SERP fetched but contained no elements at all (suspicious)
 *   {status:'error', reason}      definitive failure, retrying won't help
 *   {status:'retry', reason}      transient failure
 * Counts top-level 'paid' ads and 'shopping' (paid product listing) ads. Sitelinks nested
 * inside a paid ad are ignored - they are not separate advertisers.
 */
function adsDfsParseResponse_(httpCode, bodyText) {
  if (httpCode === 429) return { status: 'retry', reason: 'HTTP 429' };
  if (httpCode >= 500)  return { status: 'retry', reason: 'HTTP ' + httpCode };
  if (httpCode !== 200) return { status: 'error', reason: 'HTTP ' + httpCode };

  var data;
  try { data = JSON.parse(bodyText); }
  catch (e) { return { status: 'retry', reason: ADS_DFS_NET_ERR }; }

  var task = data && data.tasks && data.tasks[0];
  if (!task) return { status: 'retry', reason: ADS_DFS_NET_ERR };

  if (task.status_code !== 20000) {
    // 5xxxx = DataForSEO-side/transient; 4xxxx = our request is wrong, retrying won't help
    if (task.status_code >= 50000) return { status: 'retry', reason: 'Task error ' + task.status_code };
    return { status: 'error', reason: 'Task error ' + task.status_code };
  }

  var result = task.result && task.result[0];
  if (!result) return { status: 'retry', reason: ADS_DFS_NET_ERR };

  var items = Array.isArray(result.items) ? result.items : [];
  if (items.length === 0) return { status: 'empty' };

  var ads = [];
  items.forEach(function (it) {
    if (!it) return;
    if (it.type === 'paid') {
      ads.push({
        type: 'paid', pos: it.rank_group, title: it.title || '',
        domain: it.domain || adsDfsHost_(it.url), url: it.url || '',
        snippet: it.description || ''
      });
    } else if (it.type === 'shopping' && Array.isArray(it.items)) {
      it.items.forEach(function (el, k) {
        var host = adsDfsHost_(el.url);
        // Google redirect links don't name the merchant; fall back to the source name
        if (!host || /(^|\.)google\./.test(host)) host = String(el.source || host || '').toLowerCase();
        ads.push({
          type: 'shopping', pos: k + 1, title: el.title || '',
          domain: host, url: el.url || '', snippet: el.description || ''
        });
      });
    }
  });
  return { status: 'ok', ads: ads };
}

/**
 * Merges the snapshots for one keyword into sheet rows.
 *   - an ad seen in ANY successful snapshot counts (presence is proven even if another failed)
 *   - "No Ads Found" only when EVERY snapshot succeeded and none had ads
 *   - anything else is an error row, so a failure can never be read as 0 competitors
 * One row per advertiser (deduped by domain + ad type), best position kept.
 */
function adsDfsCombineSamples_(keyword, samples, ts) {
  var ok = samples.filter(function (r) { return r && r.status === 'ok'; });
  var byKey = {};
  var perSample = ok.map(function (r) {
    var set = {};
    r.ads.forEach(function (ad) {
      var key = ad.type + '|' + adsDfsNormDomain_(ad.domain);
      set[key] = true;
      if (!byKey[key] || (ad.pos && ad.pos < byKey[key].pos)) byKey[key] = ad;
    });
    return Object.keys(set).sort().join(',');
  });
  var disagreed = perSample.some(function (x) { return x !== perSample[0]; });
  var keys = Object.keys(byKey);

  if (keys.length) {
    return {
      disagreed: disagreed,
      rows: keys.map(function (k) {
        var ad = byKey[k];
        return [ts, keyword, ad.type, ad.pos || '', ad.title, ad.domain, ad.url, ad.snippet];
      })
    };
  }

  if (ok.length === samples.length) {
    return { disagreed: false, rows: [[ts, keyword, '', '', ADS_DFS_NO_ADS, '', '', '']] };
  }

  var failed = samples.filter(function (r) { return !r || r.status !== 'ok'; })[0] || {};
  var reason = failed.status === 'empty' ? ADS_DFS_EMPTY : (failed.reason || ADS_DFS_NET_ERR);
  return { disagreed: false, rows: [[ts, keyword, '', '', reason, '', '', '']] };
}

/** Host of a URL or display string: lowercase, no protocol/path. '' if none. */
function adsDfsHost_(s) {
  var m = String(s || '').trim().toLowerCase().match(/^(?:[a-z]+:\/\/)?([^\/\s›?#:]+)/);
  return m ? m[1] : '';
}

function adsDfsNormDomain_(d) {
  return adsDfsHost_(d).replace(/^www\./, '');
}

/** trademe.co.nz and any subdomain of it. */
function adsDfsIsTradeMe_(d) {
  var h = adsDfsNormDomain_(d);
  return h === 'trademe.co.nz' || /\.trademe\.co\.nz$/.test(h);
}

/** Same rules DataForSEO batches silently fail on (from SearchVolDataforseo.gs) */
function adsDfsIsValidKeyword_(k) {
  if (!k) return false;
  if (/[,\[\]|"()]/.test(k)) return false;
  if (k.split(/\s+/).length > 9) return false;
  if (k.length > 80) return false;
  if (/https?:\/\/|www\./i.test(k)) return false;
  return true;
}

function adsDfsHeaders_() {
  var props = PropertiesService.getScriptProperties();
  var login = props.getProperty('DFS_LOGIN')    || DFS_LOGIN;
  var pass  = props.getProperty('DFS_PASSWORD') || DFS_PASSWORD;
  return {
    'Authorization': 'Basic ' + Utilities.base64Encode(login + ':' + pass),
    'Content-Type' : 'application/json'
  };
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
      var next = sheet.getLastRow() + 1;
      sheet.getRange(next, 1, rows.length, rows[0].length).setValues(rows);
    }
  } finally {
    if (got) lock.releaseLock();
  }
}


// ==========================================================
// PARITY CHECK: Zenserp tabs vs DFS shadow tabs
// ==========================================================

/**
 * Builds a 'Parity Check' tab: per keyword, competitor (non-trademe.co.nz) ad counts from
 * Zenserp vs DataForSEO for mobile and desktop, plus the paused/enabled decision each would
 * produce under the Combined Data rule (#1 organic AND 0 mobile AND 0 desktop competitors).
 * Rows flagged DIFFERENT are the ones to investigate. Keywords with an error sentinel on
 * either side are marked UNKNOWN rather than treated as 0 competitors.
 */
function compareAdsParity() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var last = ADS_DFS.sampleEndRow || ADS_DFS.lastRow;
  var kwSheet = ss.getSheetByName(ADS_DFS.inputSheetName);
  var keywords = kwSheet.getRange(ADS_DFS.firstRow, 1, last - ADS_DFS.firstRow + 1, 1)
                        .getValues().map(function (r) { return String(r[0] || '').trim(); })
                        .filter(String);

  var zenM = adsDfsCountByKeyword_(ss.getSheetByName('AdsResultsZenMobile'));
  var zenD = adsDfsCountByKeyword_(ss.getSheetByName('AdsResultsZenDesktop'));
  var dfsM = adsDfsCountByKeyword_(ss.getSheetByName(ADS_DFS.outputSheets.mobile));
  var dfsD = adsDfsCountByKeyword_(ss.getSheetByName(ADS_DFS.outputSheets.desktop));
  var org  = adsDfsBestTradeMeOrganic_(ss.getSheetByName('Organic'));

  var out = [['Keyword', 'TM organic pos',
              'Zen mobile', 'DFS mobile', 'Zen desktop', 'DFS desktop',
              'Zen decision', 'DFS decision', 'Result']];
  var same = 0, diff = 0, unknown = 0;

  keywords.forEach(function (kw) {
    var zm = zenM[kw], dm = dfsM[kw], zd = zenD[kw], dd = dfsD[kw];
    var pos = org[kw];
    var bad = [zm, dm, zd, dd].some(function (x) { return x === undefined || x.err; });
    var decide = function (m, d) {
      if (m === undefined || d === undefined || m.err || d.err) return 'unknown';
      return (pos === 1 && m.n === 0 && d.n === 0) ? 'paused' : 'enabled';
    };
    var zDec = decide(zm, zd), dDec = decide(dm, dd);
    var result;
    if (bad || zDec === 'unknown' || dDec === 'unknown') { result = 'UNKNOWN'; unknown++; }
    else if (zDec === dDec) { result = 'same'; same++; }
    else { result = 'DIFFERENT'; diff++; }

    var n = function (x) { return x === undefined ? 'n/a' : (x.err ? 'ERR' : x.n); };
    out.push([kw, pos === undefined ? '' : pos, n(zm), n(dm), n(zd), n(dd), zDec, dDec, result]);
  });

  var sheet = ss.getSheetByName('Parity Check') || ss.insertSheet('Parity Check');
  sheet.clear();
  sheet.getRange(1, 1, out.length, out[0].length).setValues(out);
  sheet.setFrozenRows(1);
  Logger.log('Parity: ' + same + ' same, ' + diff + ' DIFFERENT, ' + unknown + ' unknown, of ' + keywords.length);
}

/**
 * keyword -> {n: distinct non-Trade Me advertiser domains, err: true if an error/skip row was
 * seen and no advertiser was found}. Works on both the Zenserp and the DFS tabs.
 */
function adsDfsCountByKeyword_(sheet) {
  var map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  var errTitles = [ADS_DFS_NET_ERR, ADS_DFS_SKIPPED, ADS_DFS_EMPTY, 'Rate Limited (429)'];
  var vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  vals.forEach(function (r) {
    var kw = String(r[1] || '').trim();
    if (!kw) return;
    var title = String(r[4] || '');
    var entry = map[kw] || (map[kw] = { domains: {}, hadErr: false });
    if (title === ADS_DFS_NO_ADS) return;
    if (errTitles.indexOf(title) !== -1 || /^(HTTP |Task error)/.test(title)) { entry.hadErr = true; return; }
    // Prefer the displayed domain; Zenserp ad links can be Google redirects
    var host = adsDfsNormDomain_(r[5]) || adsDfsNormDomain_(r[6]);
    if (host && !adsDfsIsTradeMe_(host)) entry.domains[host] = true;
  });
  Object.keys(map).forEach(function (kw) {
    var n = Object.keys(map[kw].domains).length;
    map[kw] = { n: n, err: map[kw].hadErr && n === 0 };
  });
  return map;
}

function adsDfsBestTradeMeOrganic_(sheet) {
  var best = {};
  if (!sheet || sheet.getLastRow() < 2) return best;
  var vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  vals.forEach(function (r) {
    var kw = String(r[0] || '').trim();
    var p = Number(r[1]);
    if (!kw || isNaN(p) || !adsDfsIsTradeMe_(r[2])) return;
    if (best[kw] === undefined || p < best[kw]) best[kw] = p;
  });
  return best;
}


// ==========================================================
// STATUS
// ==========================================================

function checkAdsDfsStatus() {
  var props = PropertiesService.getScriptProperties();
  var running = props.getProperty('adsDfsRunning') === 'true';
  var batches = adsDfsBuildBatches_().length;
  Logger.log('================ DFS ADS STATUS ================');
  Logger.log('Running: ' + (running ? 'YES' : 'NO'));
  if (running) {
    Logger.log('Progress: ' + props.getProperty('adsDfsDevice') + ' batch ' +
               (parseInt(props.getProperty('adsDfsBatchIndex') || '0', 10) + 1) + '/' + batches +
               ', offset ' + props.getProperty('adsDfsOffset'));
  }
  ['adsDfsLastBatch', 'adsDfsLastBatchTime', 'adsDfsCompletedAt', 'adsDfsStoppedReason',
   'adsDfsStartupFailure', 'adsDfsTriggerFailed'].forEach(function (k) {
    var v = props.getProperty(k);
    if (v) Logger.log(k + ': ' + v);
  });
  var t = ScriptApp.getProjectTriggers().map(function (x) { return x.getHandlerFunction(); });
  Logger.log('Daily starters: ' + t.filter(function (h) { return h === ADS_DFS_STARTER; }).length +
             ' (expected ' + ADS_DFS.fullAutomationHours.length + ')');
}
