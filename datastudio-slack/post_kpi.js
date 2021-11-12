function postKpi(title, data_url) {
  var picture = searchFromEmailAttachments(title);

  var now = new Date();
  var date = Utilities.formatDate(
    new Date(now.getFullYear(), now.getMonth(), now.getDate()),
    "Asia/Tokyo",
    "YYYY-MM-dd"
  );

  // 投稿先のチャンネル名を入れる
  var channel = "";
  var title, text, data_url;

  text = date;
  sendSlack(picture, title, text, data_url, channel);
}
