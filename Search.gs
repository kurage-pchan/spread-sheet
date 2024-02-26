function searchAndCopyRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName("計算用シート");
  var sheet2 = ss.getSheetByName("Output");
  
  // シート1のデータを取得
  var data = sheet1.getRange("A10:Z").getValues();
  
  // シート2のセルの値を取得
  var B6 = sheet2.getRange("B6").getValue();
  var C6 = sheet2.getRange("C6").getValue();
  var D6 = sheet2.getRange("D6").getValue();
  var E6 = sheet2.getRange("E6").getValue();
  var F6 = sheet2.getRange("F6").getValue();

  var targetlist = [B6,C6,D6,E6,F6];
  
  // シート2の5行目以降をクリア
  var lastRow = sheet2.getLastRow();
  if (lastRow >= 10) {
    sheet2.getRange("A10:Z" + lastRow).clearContent();
  }

  // targetlistが空の場合、シート1のデータをシート2にペーストして処理を終了
  if (targetlist.every(item => item === "")) {
    sheet2.getRange("A10").offset(0, 0, data.length, data[0].length).setValues(data);
    return;
  }

  // 検索
  var targetRows = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i].indexOf(B6) !== -1) {
      targetRows.push(data[i]);
    }
  }

  // シート2の10行目以降にペースト
  if (targetRows.length > 0) {
    sheet2.getRange("A10").offset(0, 0, targetRows.length, targetRows[0].length).setValues(targetRows);
  }

  for (var i = 1; i < targetlist.length; i++) {
    var seachAgent = targetlist[i];

    if (targetlist[i] != "") {
      // 更新したデータを取得
      var data = sheet2.getRange("A10:Z" + sheet2.getLastRow()).getValues();
    
      // シート2の10行目以降をクリア
      var lastRow = sheet2.getLastRow();
      if (lastRow >= 10) {
        sheet2.getRange("A10:Z" + lastRow).clearContent();
      }
    
      // 検索
      var targetRows = [];
      for (var j = 0; j < data.length; j++) {
        if (data[j].indexOf(seachAgent) !== -1) {
          targetRows.push(data[j]);
        }
      }
    
      // シート2の5行目以降にコピー
      if (targetRows.length > 0) {
        sheet2.getRange("A10").offset(0, 0, targetRows.length, targetRows[0].length).setValues(targetRows);
      }
    }
  }
}
