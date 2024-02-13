function sendNotification(e) {
  // 設定
  var url = "https://ptb.discord.com/api/webhooks/"; // DiscordのWebhook URLをここに入力してください

  // スプレッドシートの情報を取得
  var form = FormApp.getActiveForm();
  var responseSheetId = form.getDestinationId();
  var responseSheet = SpreadsheetApp.openById(responseSheetId);
  var sheet = responseSheet.getActiveSheet();

  var responses = e.response.getItemResponses();
  var targetTeam = responses[0].getResponse();
  var teamResponse = responses[1].getResponse();
  
  // 通し番号のカウント
  var lastRow = sheet.getLastRow();
  var counter = lastRow - 1;

  sheet.getRange("D1").setValue("通し番号");
  sheet.getRange("D"+lastRow).setValue("#"+counter);

  // Discordに送信するメッセージの設定
  var colorCode = parseInt("219ddd", 16);
  var embeds = [
    {
      "title": `${targetTeam}`, 
      "color": colorCode,
      "fields": [
        {
          "name" : "",
          "value": `${teamResponse}`,
          "inline": false
        },
        {
          "name" : "",
          "value": "#"+counter,
          "inline": false
        }
      ],
    }
  ]
  //Discordにメッセージを送信する
  sendToDiscord(url, embeds);
}


function sendToDiscord(url, embeds) {
  var jsonData = {
  "embeds" : embeds
  };
  var payload = JSON.stringify(jsonData);
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': payload,
  };
   
  UrlFetchApp.fetch(url, options);
}
