var SHEET_NAME = 'PV情報';

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    var headers = data.headers;
    var values = data.values;

    if (!headers || !values || headers.length === 0) {
      throw new Error('headers/values が空です');
    }

    var numCols = headers.length;

    // Always overwrite row 1 with the latest header set
    if (sheet.getLastColumn() < numCols) {
      sheet.insertColumnsAfter(
        Math.max(sheet.getLastColumn(), 1),
        numCols - Math.max(sheet.getLastColumn(), 1)
      );
    }
    sheet.getRange(1, 1, 1, numCols).setValues([headers]);
    sheet.getRange(1, 1, 1, numCols)
      .setFontWeight('bold')
      .setBackground('#1B4332')
      .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);

    // Upsert by ANDPADシステムID (column 4 in the 134-column layout)
    var andpadColIndex = headers.indexOf('ANDPADシステムID');
    var andpadId = andpadColIndex >= 0 ? String(values[andpadColIndex]) : '';
    var updateRow = -1;

    if (andpadId && sheet.getLastRow() > 1) {
      var colValues = sheet.getRange(2, andpadColIndex + 1, sheet.getLastRow() - 1, 1).getValues();
      for (var r = 0; r < colValues.length; r++) {
        if (String(colValues[r][0]) === andpadId) {
          updateRow = r + 2;
          break;
        }
      }
    }

    // Pad values to match header count
    while (values.length < numCols) {
      values.push('');
    }

    if (updateRow > 0) {
      sheet.getRange(updateRow, 1, 1, numCols).setValues([values]);
    } else {
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, numCols).setValues([values]);
    }

    sheet.autoResizeColumns(1, Math.min(numCols, 20));

    return ContentService.createTextOutput(
      JSON.stringify({ status: 'ok', message: updateRow > 0 ? '更新しました' : '新規登録しました' })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var id = e && e.parameter ? e.parameter.id : '';

  if (!id) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'ok', message: 'PV情報収集ツール GAS API' })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet || sheet.getLastRow() < 2) {
      throw new Error('データが見つかりません');
    }

    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var andpadColIndex = headers.indexOf('ANDPADシステムID');

    if (andpadColIndex < 0) {
      throw new Error('ANDPADシステムID列が見つかりません');
    }

    var colValues = sheet.getRange(2, andpadColIndex + 1, sheet.getLastRow() - 1, 1).getValues();
    var rowIndex = -1;
    for (var r = 0; r < colValues.length; r++) {
      if (String(colValues[r][0]) === String(id)) {
        rowIndex = r + 2;
        break;
      }
    }

    if (rowIndex < 0) {
      throw new Error('該当するANDPADシステムIDのデータが見つかりません');
    }

    var rowValues = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
    var data = {};
    for (var c = 0; c < headers.length; c++) {
      if (headers[c] === '') continue;
      data[headers[c]] = rowValues[c];
    }

    return ContentService.createTextOutput(
      JSON.stringify({ status: 'success', data: data })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
