var XLSX = {};
//FILE REPORT XLSX
//idapp 43704351-3fa5-4dac-8f80-9c2a55a7d5da
// key V2-YXz9z-peiTp-blDqN-DtkZ5-dH0fX-3YssU-8x0ZU-c8vqG
function importXLSXToAppSheet(fileName) {
  Logger.log("Tên file nhận được từ cột FILE CSV: " + fileName);
  if (fileName.startsWith('FILE REPORT XLSX/')) {
    fileName = fileName.replace('FILE REPORT XLSX/', '');
  }
  Logger.log('Tên file sau khi loại bỏ đường dẫn: ' + fileName);
  var folderId = '16Y3ZO3WGH0cvkrwR6v6RdWRY2q6YN8PE';
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByName(fileName);

  if (files.hasNext()) {
    var file = files.next();
    Logger.log('Đã tìm thấy file: ' + file.getName());

    var fileId = file.getId();
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
      var fileUrl = 'https://www.googleapis.com/drive/v3/files/' + fileId + '/copy';
      var response = UrlFetchApp.fetch(fileUrl, options);
      var jsonResponse = JSON.parse(response.getContentText());
      var newFileId = jsonResponse.id;
      Logger.log('ID file mới: ' + newFileId);

      var tempSpreadsheet = SpreadsheetApp.openById(newFileId);
      var tempSheet = tempSpreadsheet.getSheets()[0];
      var data = tempSheet.getDataRange().getValues();

      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var taskRequestProduct = row[0] ? row[0].toString().trim() : '';
        var description = row[1] ? row[1].toString().trim() : '';
        var reportsCategory = row[2] ? row[2].toString().trim() : '';
        var healthCheckScore = row[3] ? row[3].toString().trim() : '';
        var startTime = (row[4] instanceof Date) ? row[4].toISOString() : '';
        var endTime = (row[5] instanceof Date) ? row[5].toISOString() : '';

        if (!taskRequestProduct) {
          Logger.log('Dữ liệu không hợp lệ tại hàng ' + (i + 1) + ': TASK REQUEST PRODUCT không được để trống.');
          continue;
        }

        var payload = {
          "Action": "Add",
          "Properties": {
            "Locale": "en-US"
          },
          "Rows": [{
            "TASK REQUEST PRODUCT": taskRequestProduct,
            "DESCRIPTION": description,
            "REPORT CATEGORY": reportsCategory,
            "HEALTH CHECK SCORE": healthCheckScore,
            "STARTTIME": startTime,
            "ENDTIME": endTime
          }]
        };

        var apiUrl = "https://api.appsheet.com/api/v2/apps/43704351-3fa5-4dac-8f80-9c2a55a7d5da/tables/REPORTS/Action";
        var apiOptions = {
          "method": "post",
          "contentType": "application/json",
          "headers": {
            "ApplicationAccessKey": "V2-YXz9z-peiTp-blDqN-DtkZ5-dH0fX-3YssU-8x0ZU-c8vqG"
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
      Logger.log('Lỗi khi xử lý file: ' + e.message);
    }
  } else {
    Logger.log("Không tìm thấy file với tên: " + fileName);
  }
}
