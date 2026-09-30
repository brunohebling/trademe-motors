/* =================================================================== */
/*   SEARCH-VOLUME SNAPSHOT – GOOGLE ADS KEYWORD PLANNER (LIVE)        */
/*   Uses:  https://api.dataforseo.com/v3/keywords_data/google_ads/search_volume/live                                         */
/* =================================================================== */
function fetchSearchVolume() {

  const ss        = SpreadsheetApp.getActive();
  const kwSheet   = ss.getSheetByName('GSC Filtered Data');
  const volSheet  = ss.getSheetByName('SearchVolume') || ss.insertSheet('SearchVolume');

  // overwrite previous run
  volSheet.clear();
  volSheet.appendRow(['Keyword','Search Volume (avg monthly)','CPC (USD)','Competition','Timestamp']);

  const authHdr = {
    'Authorization': 'Basic ' + Utilities.base64Encode(DFS_LOGIN + ':' + DFS_PASSWORD),
    'Content-Type' : 'application/json'
  };

  // All keywords from GSC Filtered Data, col A, from row 2 down
  const lastRow = kwSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No keywords found in GSC Filtered Data.');
    return;
  }

const allKeywords = kwSheet
    .getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat()
    .map(k => (k || '').toString().trim())
    .filter(k => k !== '')
    .filter(k => !k.includes(',') && !k.includes('[') && !k.includes(']') && !k.includes('|') && !k.includes('"') && !k.includes('(') && !k.includes(')'))
    .filter(k => k.split(' ').length <= 9)
    .filter(k => !k.includes('http://') && !k.includes('https://') && !k.includes('www.'))  // ← filter URLs
    .filter(k => k.length <= 80)  // ← DataForSEO's keyword character limit

  Logger.log('Total input keywords: ' + allKeywords.length);

  const ts   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const rows = [];

  /* --- DataForSEO location code for New Zealand (Google Ads) ---------- */
  const LOCATION_CODE_NZ = 2554;     // (country level)

  /* batch in chunks of 500 keywords (= API limit per live call) */
  const CHUNK = 500;
  for (let i = 0; i < allKeywords.length; i += CHUNK) {

    const chunk      = allKeywords.slice(i, i + CHUNK);
    const chunkIndex = (i / CHUNK) + 1;
    Logger.log('Processing chunk ' + chunkIndex + ' with ' + chunk.length + ' keywords');

    const payload = JSON.stringify([{
      keywords          : chunk,
      location_code     : LOCATION_CODE_NZ,
      language_code     : LANGUAGE_CODE,   // your global "en"
      include_serp_info : false,           // we only need volume/CPC
      priority          : 2
    }]);

    const resp = UrlFetchApp.fetch(
      'https://api.dataforseo.com/v3/keywords_data/google_ads/search_volume/live',
      { method:'post', payload:payload, headers:authHdr, muteHttpExceptions:true }
    );

    const code = resp.getResponseCode();
    Logger.log('Chunk ' + chunkIndex + ' HTTP status: ' + code);

if (code === 200) {
  const data = JSON.parse(resp.getContentText());
  const taskStatus = data.tasks?.[0]?.status_code;

  Logger.log('Chunk ' + chunkIndex + ' task status: ' + taskStatus);

  if (chunkIndex === 1) {
    Logger.log('CHUNK 1 FULL RESPONSE: ' + resp.getContentText().substring(0, 3000));
  }

  if (taskStatus !== 20000) {
    Logger.log('Chunk ' + chunkIndex + ' task error: ' + data.tasks?.[0]?.status_message);
  }

  const kwInfo = taskStatus === 20000 && data.tasks[0].result
                 ? data.tasks[0].result
                 : [];

  Logger.log('Chunk ' + chunkIndex + ' returned ' + kwInfo.length + ' result items');

  // Build a map: keyword (lowercase) → item  ← keep everything from here down as-is

      // Build a map: keyword (lowercase) → item
      const resultByKeyword = {};
      kwInfo.forEach(item => {
        if (item && item.keyword) {
          resultByKeyword[item.keyword.toLowerCase()] = item;
        }
      });

      // For every keyword in the original chunk, write a row (even if DFS had no data)
      chunk.forEach(kw => {
        const item = resultByKeyword[kw.toLowerCase()] || {};

        rows.push([
          kw,
          item.search_volume || '',
          item.cpc           || '',
          item.competition   || '',
          ts
        ]);
      });

    } else {
      Logger.log('Error for chunk ' + chunkIndex + ': HTTP ' + code + ' - ' + resp.getContentText());
      // visible error row so you know a whole chunk died
      rows.push(['CHUNK_' + chunkIndex, 'ERR ' + code, '', '', ts]);
    }

    Utilities.sleep(500);   // tiny pause – keep well below DFS live QPS limits
  }

  Logger.log('Total rows to write into SearchVolume: ' + rows.length);

  if (rows.length) {
    volSheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }
}
