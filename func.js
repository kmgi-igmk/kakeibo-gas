function doGet(e) {
  const pagename = (e.parameter.page || 'index');
  return HtmlService.createTemplateFromFile(pagename)
                    .evaluate()
                    .setTitle('ためため家計簿')
                    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getGASUrl() {
  return ScriptApp.getService().getUrl();
}

function getSheetData(sheetName, year, rowIdx, colIdx, rowNum, colNum) {
  let sheet = getSheet(sheetName, year);
  return sheet.getRange(rowIdx, colIdx, rowNum, colNum).getValues();
}

/*
* Tableに追加された複数のデータをまとめてSpreadSheetに登録する
*/
function batchInsert(x, year) {
  let logLevel = 'INFO';
  let logging = `Registered. data=[${x.toString()}]`;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5 * 1000)
    let sheet = getSheet('receipts', year);
    let lastRow = sheet.getLastRow();
    sheet.getRange(`${++lastRow}`, 1, x.length, x[0].length).setValues(x);
  } catch (e) {
    logLevel = 'ERROR';
    logging = e.message;
  } finally {
    lock.releaseLock();
    dumpLog(year, logLevel, logging);
  }
}

// 以下private function
function getSsId(year) {
  const json = getJson('id.json');
  return json['ssId'][0][year];
}

function getSheet(name, year) {
  let ss = SpreadsheetApp.openById(getSsId(year));
  return ss.getSheetByName(name);
}

function getJson(fname) {
  var fileIT = DriveApp.getFilesByName(fname).next();
  var textdata = fileIT.getBlob().getDataAsString('utf8');
  return JSON.parse(textdata);
}

function dumpLog(year, level, logging) {
  let sheet = getSheet('log', year);
  sheet.appendRow([
    new Date(),
    level,
    logging
  ]);
}
