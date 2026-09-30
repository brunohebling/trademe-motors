/** @OnlyCurrentDoc */

// custom menu function
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Save Paused Keywords Data', 'savePausedKeywordsData')
      .addItem('Save A1 and B1 Data', 'saveA1B1Data')
      .addToUi();
}

// function to save data for paused keywords
function savePausedKeywordsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historicalSheet = ss.getSheetByName("Historical Data");
  const keywords_paused = getPausedKeywords();
  const date = new Date();
  const keywordsArray = keywords_paused.split(', '); // Split the keywords string into an array
  
  // Find the next available row
  let startRow = 6; // Start populating from row 6
  let lastRow = historicalSheet.getLastRow();
  let rowToWrite = lastRow > startRow ? lastRow + 1 : startRow;

  // Append each keyword in a separate row starting from B6
  keywordsArray.forEach(keyword => {
    historicalSheet.getRange("B" + rowToWrite).setValue(date);
    historicalSheet.getRange("C" + rowToWrite).setValue(keyword);
    rowToWrite++;
  });
}

// function to save data for A1 and B1
function saveA1B1Data() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pausedKeywordSheet = ss.getSheetByName("Paused Keyword Count");
  const keywords_paused = pausedKeywordSheet.getRange("A1").getValue();
  const date = pausedKeywordSheet.getRange("B1").getValue();
  pausedKeywordSheet.appendRow([keywords_paused, date]);
}

// function to get paused keywords
function getPausedKeywords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Combined Data");
  const lastRow = dataSheet.getLastRow();
  const keywords = dataSheet.getRange("A2:A" + lastRow).getValues(); // Assuming keywords are in column A
  const recommendations = dataSheet.getRange("E2:E" + lastRow).getValues(); // Assuming recommendations are in column E
  const pausedKeywords = [];
  
  for (let i = 0; i < keywords.length; i++) {
    if (recommendations[i][0].toLowerCase() === "paused") { // Assuming "Paused" is written in lowercase
      pausedKeywords.push(keywords[i][0]);
    }
  }
  
  return pausedKeywords.join(', '); // Join keywords into a single string
}

/* ❶  Organic rank for a batch of keywords */
function DFS_ORGANIC_POSITIONS(keywords, domain) {
  const creds = Utilities.base64Encode('bruno.hebling@trademe.co.nz:ed50c1ebdf956254');   // <-- replace
  const tasks = keywords.filter(k => k)
    .map(k => ({keyword: k,
                location_code: 2840,
                language_code: 'en',
                device: 'desktop'}));
  const res = UrlFetchApp.fetch(
      'https://api.dataforseo.com/v3/serp/google/organic/live',
      {method: 'post',
       contentType: 'application/json',
       headers: {Authorization: 'Basic ' + creds},
       payload: JSON.stringify(tasks),
       muteHttpExceptions: true});
  const json = JSON.parse(res)
               .tasks.map(t => {
                 const hit = t.result[0].items
                              .find(i => i.domain.endsWith(domain));
                 return [t.keyword,
                         hit ? hit.rank_group : '>100'];
               });
  return json;
}
