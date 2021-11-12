// タブにパスワードジェネレータを表示してボタンひとつでランダムなパスワードを作成してくれるやつ
var ss = SpreadsheetApp.getActiveSpreadsheet();

function onOpen() {
  var menus = [
    { name: "8文字で生成", functionName: "at8" },
    { name: "12文字で生成", functionName: "at12" },
  ];
  ss.addMenu("🔑パスワードジェネレーター🔑", menus);
}

function at8() {
  setPassList(8);
}
function at12() {
  setPassList(12);
}

function setPassList(len) {
  try {
    // 対象のスプレッドシートIDを入れる
    var sheetId = "";
    var sheet =
      SpreadsheetApp.openById(sheetId).getSheetByName(
        "アカウント一括入稿シート"
      );
    //.getActiveSheet()
  } catch (e) {
    Logger.log(e);
  }

  // ログインIDの入っている行が対象
  var range = sheet.getRange(1, 6, sheet.getLastRow()).getValues();
  Logger.log(range.length);

  var headerLen = 0;
  // A列にNoと書いてある行までをheaderとする
  for (let i = 1; ; i++) {
    var headerMark = sheet.getRange(i, 1).getValue();
    if (headerMark === "No") {
      headerLen = i;
      break;
    }
  }

  // header行はスキップしてパスワードを挿入する
  for (var i = headerLen + 1; i <= range.length + 1; i++) {
    var loginId = sheet.getRange(i, 6).getValue();
    Logger.log(loginId);
    if (loginId === "") {
      continue;
    }

    var pass = generatePass(len);
    sheet.getRange(i, 7).setValue(pass);
  }
}
// generatePass 英数字小文字大文字を含むパスワードを生成する
function generatePass(len) {
  //英数字を用意する
  var lowercaseLetters = "abcdefghijklmnopqrstuvwxyz";
  var uppercaseLetters = lowercaseLetters.toUpperCase();
  var numbers = "0123456789";

  // それぞれ何文字使うかランダムで決める
  var lowercaseLen = Math.ceil(Math.random() * (len - 2));
  var uppercaseLen = Math.ceil(Math.random() * (len - lowercaseLen - 1));
  var numberLen = len - lowercaseLen - uppercaseLen;

  // それぞれの長さ分ランダムに使う英数字を決める
  var password = [];
  for (var i = 0; i < lowercaseLen; i++) {
    password.push(
      lowercaseLetters.charAt(
        Math.floor(Math.random() * lowercaseLetters.length)
      )
    );
  }

  for (var i = 0; i < uppercaseLen; i++) {
    password.push(
      uppercaseLetters.charAt(
        Math.floor(Math.random() * uppercaseLetters.length)
      )
    );
  }
  for (var i = 0; i < numberLen; i++) {
    password.push(numbers.charAt(Math.floor(Math.random() * numbers.length)));
  }

  // ランダムに並び替え
  for (let i = password.length - 1; i >= 0; i--) {
    const randomNumber = Math.floor(Math.random() * (i + 1));
    [password[i], password[randomNumber]] = [
      password[randomNumber],
      password[i],
    ];
  }

  return password.join("");
}
