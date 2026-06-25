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

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#1B4332')
        .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    } else {
      var existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var existingSet = {};
      for (var i = 0; i < existingHeaders.length; i++) {
        existingSet[existingHeaders[i]] = i;
      }
      var newHeaders = [];
      for (var j = 0; j < headers.length; j++) {
        if (!(headers[j] in existingSet)) {
          newHeaders.push(headers[j]);
        }
      }
      if (newHeaders.length > 0) {
        var col = sheet.getLastColumn() + 1;
        for (var k = 0; k < newHeaders.length; k++) {
          sheet.getRange(1, col + k).setValue(newHeaders[k]);
        }
        existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        for (var m = 0; m < existingHeaders.length; m++) {
          existingSet[existingHeaders[m]] = m;
        }
      }
      var finalHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var reordered = [];
      for (var n = 0; n < finalHeaders.length; n++) {
        var idx = headers.indexOf(finalHeaders[n]);
        reordered.push(idx >= 0 ? values[idx] : '');
      }
      values = reordered;
      headers = finalHeaders;
    }

    var existingRows = sheet.getLastRow() > 1
      ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
      : [];
    var andpadId = data.andpadId || '';
    var facilityName = data.facilityName || '';
    var updateRow = -1;

    if (andpadId) {
      var andpadCol = headers.indexOf('ANDPADシステムID');
      if (andpadCol >= 0 && sheet.getLastRow() > 1) {
        var colValues = sheet.getRange(2, andpadCol + 1, sheet.getLastRow() - 1, 1).getValues();
        for (var r = 0; r < colValues.length; r++) {
          if (String(colValues[r][0]) === String(andpadId)) {
            updateRow = r + 2;
            break;
          }
        }
      }
    }

    if (updateRow > 0) {
      sheet.getRange(updateRow, 1, 1, values.length).setValues([values]);
    } else {
      sheet.appendRow(values);
    }

    sheet.autoResizeColumns(1, headers.length > 20 ? 20 : headers.length);

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

function doGet() {
  return ContentService.createTextOutput(
    JSON.stringify({ status: 'ok', message: 'PV情報収集ツール GAS API' })
  ).setMimeType(ContentService.MimeType.JSON);
}
