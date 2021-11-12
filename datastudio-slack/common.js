// Datastudioから配信された最新のメールに添付されている画像を取得
function searchFromEmailAttachments(subject) {
  var Threads = GmailApp.search(
    'from:"data-studio-noreply@google.com" subject:' + subject,
    0,
    1
  );
  var messages = Threads[0].getMessages();
  var message = messages[messages.length - 1];

  // メールから画像を取得
  var attachments = message.getAttachments();
  //var pictures = []
  for (attachment of attachments) {
    if (attachment.getName().indexOf(".jpg") > 0) {
      //pictures.push(attachment);
      return attachment;
    }
  }
  //return pictures;
}

function sendSlack(file, title, text, data_url, channel) {
  var url = "https://slack.com/api/files.upload";
  // 使用するbotのtokenを入れる
  var token = "";

  var data = {
    token: token,
    channels: channel,
    title: title,
    initial_comment: "*▼" + text + "*\n" + data_url,
    file: file,
    file_type: "jpg",
  };

  var option = {
    method: "POST",
    payload: data,
  };

  return UrlFetchApp.fetch(url, option);
}
