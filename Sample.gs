function OrderCHECK() {
  // スプレッドシートの取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 注文データが入っているシート名
  var sheetName = "リスト(raw)";  // 既存リストも変える必要あり
  var dataSheet = spreadsheet.getSheetByName(sheetName);
  
  // 受注データの抽出結果を貼り付けるシート名
  var resultSheetName = "顧客別注文データ";  // 新しく作成する必要あり
  var resultSheet = spreadsheet.getSheetByName(resultSheetName);
  
  // データ範囲と値を取得
  var dataRange = dataSheet.getDataRange();
  var data = dataRange.getValues();
  
  // ヘッダー行を格納する配列
  var result = [];
  
  // データを検査して結果に追加
  for (var i = 1; i < data.length; i++) {
    var name = data[i][2];
    name = name.replace(/[\s\t\n]/g, "");
    
    var status = data[i][7];
    var cancelStatus = "ｷｬﾝｾﾙ済み（予約可能枠）です";  // キャンセルステータス文言

    if (status == "受取待ちです") {
      // 2023/10/28：同じデータがある場合は個数を加算する処理追加
      var existingIndex = result.findIndex(function(row) {
        return (
          row[2].replace(/[\s\t\n]/g,"") == name && // 名前
          row[3] == data[i][3] && // 受取日
          row[4] == data[i][4] && // 受取時間
          row[5] == data[i][5] && // 店舗
          row[6] == data[i][6]    // 電話番号
        );
      });

    if (existingIndex !== -1) {
      for (var k = 10; k < data[i].length; k++) {
        // 2023/11/17：「-」が入っていたら除外する処理に変更
        var valueToAdd = data[i][k];
        if (valueToAdd !== "-" && !isNaN(Number(valueToAdd))) {
          result[existingIndex][k] = result[existingIndex][k] !== "-" ? Number(result[existingIndex][k]) + Number(valueToAdd) : Number(valueToAdd);
        }
      }
    } else {
      result.push(data[i]);
    }
  }
  
  if (status == cancelStatus) {
    var existingIndex = result.findIndex(function(row) {
      return (
        row[2].replace(/[\s\t\n]/g,"") == name && // 名前
        row[3] == data[i][3] && // 受取日
        row[4] == data[i][4] && // 受取時間
        row[5] == data[i][5] && // 店舗
        row[6] == data[i][6]    // 電話番号
      );
    });

    if (existingIndex !== -1) {
      for (var k = 10; k < data[i].length; k++) {
        // 2023/11/17：「-」が入っていたら除外する処理に変更
        var valueToAdd = data[i][k];
        if (valueToAdd !== "-" && !isNaN(Number(valueToAdd))) {
          result[existingIndex][k] = result[existingIndex][k] !== "-" ? Number(result[existingIndex][k]) + Number(valueToAdd) : Number(valueToAdd);
        }
      }
    } else {
      result.push(data[i]);
    }
  }
}
  
  // 結果をクリア（見出しは残す）
  if (result.length > 0) {
    resultSheet.getRange(2, 1, result.length, result[0].length).clear();
  }

  // 結果を結果シートに貼り付け
  resultSheet.getRange(2, 1, result.length, result[0].length).setValues(result);

  // 空白を削除（半角・全角）
  resultSheet.getDataRange().createTextFinder(" ").replaceAllWith("");
  resultSheet.getDataRange().createTextFinder("　").replaceAllWith(""); // 全角スペース

  // 関数を追加する処理
  var lastRowWithData = resultSheet.getLastRow();

  // データがない場合は処理しない
  if (lastRowWithData > 1) {
    for (var row = 2; row <= lastRowWithData; row++) {
      var sumFormula = "=SUM(M" + row + ":DJ" + row + ")";
      resultSheet.getRange(row, 12).setFormula(sumFormula);

      // 商品名参照関数
      var indexMatchFormula = '=ARRAYFORMULA(TEXTJOIN(CHAR(10), TRUE, IF($M' + row + ':$DJ' + row + '>0, VLOOKUP($M$1:$DJ$1, \'商品マスタ\'!A:B, 1, FALSE) & " " & VLOOKUP($M$1:$DJ$1, \'商品マスタ\'!A:B, 2, FALSE) & " x " & $M' + row + ':$DJ' + row + ' & "個", "")))';
      resultSheet.getRange(row, 9).setFormula(indexMatchFormula);
    }
  }
}
