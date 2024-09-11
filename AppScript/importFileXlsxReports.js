var XLSX = {};

function importXLSXToAppSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TASK REQUEST PRODUCT');
  var lastRow = sheet.getLastRow();
  Logger.log('Hàng cuối cùng: ' + lastRow);

  if (lastRow === 0) {
    Logger.log('Không có dữ liệu trong bảng.');
    return;
  }

  var fileFound = false;
  var fileName = '';

  for (var rowIndex = 1; rowIndex <= lastRow; rowIndex++) {
    var cellValue = sheet.getRange(rowIndex, 15).getValue();
    if (cellValue && cellValue.endsWith('.xlsx')) {
      fileName = cellValue;
      fileFound = true;
      Logger.log('Tên file tìm thấy: ' + fileName);
      break;
    }
  }

  if (!fileFound) {
    Logger.log('Không tìm thấy file .xlsx trong cột 15.');
    return;
  }

  if (fileName.startsWith('FILE CSV/')) {
    fileName = fileName.replace('FILE CSV/', '');
  }

  Logger.log('Tên file sau khi loại bỏ đường dẫn: ' + fileName);

  var folderId = '16Y3ZO3WGH0cvkrwR6v6RdWRY2q6YN8PE';
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByName(fileName);

  if (files.hasNext()) {
    var file = files.next();
    Logger.log('Đã tìm thấy file: ' + file.getName());

    var blob = file.getBlob();
    
    var newSpreadsheet = SpreadsheetApp.create(fileName + ' - Copy');
    var newSheet = newSpreadsheet.getActiveSheet();
    
    var fileId = file.getId();
    var fileUrl = 'https://www.googleapis.com/drive/v3/files/' + fileId + '/copy';
    var options = {
      'method': 'POST',
      'headers': {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify({
        'name': fileName + ' - Copy',
        'mimeType': MimeType.GOOGLE_SHEETS
      })
    };
    
    try {
      var response = UrlFetchApp.fetch(fileUrl, options);
      var jsonResponse = JSON.parse(response.getContentText());
      var newFileId = jsonResponse.id;
      Logger.log('ID file mới: ' + newFileId);

      var tempSpreadsheet = SpreadsheetApp.openById(newFileId);
      var tempSheet = tempSpreadsheet.getSheets()[0];
      var data = tempSheet.getDataRange().getValues();

      for (var i = 1; i < data.length; i++) {
        var row = data[i];

        var taskRequestProduct = (typeof row[1] === 'string') ? row[1].trim() : '';
        // var product = (typeof row[2] === 'string') ? row[2].trim() : '';
        // var productName = (typeof row[3] === 'string') ? row[3].trim() : '';
        // var productType = (typeof row[4] === 'string') ? row[4].trim() : '';
        var description = (typeof row[5] === 'string') ? row[5].trim() : '';
        var score = (typeof row[10] === 'number') ? row[10] : Number(row[10]);
        // var startTime = new Date(row[14]); 
        // var endTime = new Date(row[15]); 

        // var service = (typeof row[16] === 'string') ? row[16].trim() : '';
        // var serviceDelivery = (typeof row[17] === 'string') ? row[17].trim() : '';
        // var taskRequest = (typeof row[18] === 'string') ? row[18].trim() : '';

        if (!taskRequestProduct) {
          Logger.log('Dữ liệu không hợp lệ tại hàng ' + (i+1) + ': TASK REQUEST PRODUCT không được để trống.');
          continue;
        }

        var payload = {
          "Action": "Add",
          "Properties": {
            "Locale": "en-US"
          },
          "Rows": [
            {
              "TASK REQUEST PRODUCT": taskRequestProduct,
              // "PRODUCT": product,
              // "PRODUCT NAME": productName,
              // "PRODUCT TYPE": productType,
              "DESCRIPTION": description,
              "SCORE": score,
              // "SERVICE": service,
              // "SERVICE DELIVERY": serviceDelivery,
              "FILE CSV": fileName,
              // "START TIME": startTime,
              // "END TIME": endTime
            }
          ]
        };

        var apiUrl = "https://api.appsheet.com/api/v2/apps/a8cc83c0-c3bf-45e3-8739-8621466bf243/tables/REPORTS/Action";
        var apiOptions = {
          "method": "post",
          "contentType": "application/json",
          "headers": {
            "ApplicationAccessKey": "V2-AhEY7-x4cEJ-7Sjqa-srG7U-z7lqQ-hu2S4-EE48p-1aM36"
          },
          "payload": JSON.stringify(payload),
          "muteHttpExceptions": true
        };

        try {
          var apiResponse = UrlFetchApp.fetch(apiUrl, apiOptions);
          Logger.log('Phản hồi từ API: ' + apiResponse.getContentText());
        } catch (e) {
          Logger.log('Lỗi khi gửi yêu cầu: ' + e.message);
        }
      }

      DriveApp.getFileById(newFileId).setTrashed(true);
    } catch (e) {
      Logger.log('Lỗi khi chuyển đổi và xử lý file: ' + e.message);
    }
  } else {
    Logger.log("Không tìm thấy file với tên: " + fileName);
  }
}