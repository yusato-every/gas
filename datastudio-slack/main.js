// スプレッドシートに登録されているdatastudioをslack連携する
function main() {
  try {
    // 登録リストを管理しているスプレッドシートIDを入れる
    var sheetId = "";
    var sheet = SpreadsheetApp.openById(sheetId).getSheetByName("連携対象");
    // .getActiveSheet();
  } catch (e) {
    Logger.log(e);
  }

  // ログインIDの入っている行が対象
  var range = sheet.getRange(1, 2, sheet.getLastRow()).getValues();
  Logger.log(range.length);

  for (var i = 2; i <= range.length + 1; i++) {
    var title = sheet.getRange(i, 1).getValue();
    var data_url = sheet.getRange(i, 2).getValue();

    Logger.log(title, data_url);
    if (title === "" || data_url === "") {
      continue;
    }

    postKpi(title, data_url);
  }
}
