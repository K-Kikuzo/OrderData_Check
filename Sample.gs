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
    var status = data[i][7];
    var cancelStatus = "ｷｬﾝｾﾙ済み（予約可能枠）です";  // キャンセルステータス文言

    if (status == "受取待ちです") {
      // 2023/10/28：同じデータがある場合は個数を加算する処理追加
      var existingIndex = result.findIndex(function(row) {
        return (
          row[2] == data[i][2] && // 名前
          row[3] == data[i][3] && // 受取日
          row[4] == data[i][4] && // 受取時間
          row[6] == data[i][6]    // 電話番号
        );
      });

      if (existingIndex !== -1) {
        for (var k = 11; k < data[i].length; k++) {
          result[existingIndex][k] += data[i][k];
        }
      } else {
        result.push(data[i]);
      }
    } else if (status == cancelStatus) {
      // キャンセルの場合、対応する「受取待ちです」の行を検索し、計算を行う（キーは名前・受取日・受取時間・電話番号の4つとする）
      for (var j = i - 1; j >= 0; j--) {
        if (
          data[j][2] == data[i][2] && // 名前
          data[j][3] == data[i][3] && // 受取日
          data[j][4] == data[i][4] && // 受取時間
          data[j][6] == data[i][6]    // 電話番号
        ) {
          // 受注数とキャンセル個数の計算を行う
          for (var k = 11; k < data[j].length; k++) {
            data[j][k] += data[i][k];
          }
          break;
        }
      }
    }
  }
  
  // 結果をクリア（見出しは残す）
  if (result.length > 0) {
    resultSheet.getRange(2, 1, result.length, result[0].length).clear();
  }

  // 結果を結果シートに貼り付け
  resultSheet.getRange(2, 1, result.length, result[0].length).setValues(result);

  // SUM関数を追加する処理
  var lastRowWithData = resultSheet.getLastRow();

  // データがない場合は処理しない
  if (lastRowWithData > 1) {
    for (var row = 2; row <= lastRowWithData; row++) {
      var sumFormula = "=SUM(L" + row + ":AQ" + row + ")";
      resultSheet.getRange(row, 11).setFormula(sumFormula);
    }
  }
}
