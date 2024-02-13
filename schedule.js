function sendNotification() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastExecutionDate = scriptProperties.getProperty('lastExecutionDate');
  var currentDate = new Date().toLocaleDateString();
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = activeSpreadsheet.getSheetByName("毎日のスケジュール");
  var archiveSheet = activeSpreadsheet.getSheetByName("アーカイブ");
  var range = activeSheet.getRange("C5:I52");
  var sourceColumn = "C";
  var column = 3; //C
  var targetColumn = "I";
  var startRow = 5;
  var endRow = 52;

  //各種設定項目の読み込み
  var botName = activeSheet.getRange("J6").getValue().toString();
  var botIcon = activeSheet.getRange("J8").getValue().toString();
  var botUrl = activeSheet.getRange("J10").getValue().toString();

  var advanceNotice = activeSheet.getRange("J12").getValue();

  var sendTitle = activeSheet.getRange("J14").getValue().toString();
  var sendS = activeSheet.getRange("J16").getValue().toString();
  var sendE = activeSheet.getRange("J18").getValue().toString();

  var userIDs = activeSheet.getRange("J24").getValue().toString().split(",");
  var roleIDs = activeSheet.getRange("J26").getValue().toString().split(",");

  var values = range.getValues();
  var fontWeights = range.getFontWeights();
  var fontFamilies = range.getFontFamilies();

  var sourceRange = activeSheet.getRange(sourceColumn + startRow + ":" + targetColumn + endRow);
  var sourceFontWeights = sourceRange.getFontWeights();
  var sourceFontFamilies = sourceRange.getFontFamilies();
  var sourceValues = sourceRange.getValues();
  var sourceFontLines = sourceRange.getFontLines();

  var source2Range = activeSheet.getRange(sourceColumn + startRow + ":" + sourceColumn + endRow);
  var source2FontWeights = source2Range.getFontWeights();
  var source2FontFamilies = source2Range.getFontFamilies();
  var source2Values = source2Range.getValues();
  var source2FontLines = source2Range.getFontLines();
  
  if (lastExecutionDate == currentDate) {
    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length - 1; j++) {
        values[i][j] = values[i][j + 1].toString();
        fontWeights[i][j] = fontWeights[i][j + 1];
        fontFamilies[i][j] = fontFamilies[i][j + 1];
        sourceFontLines[i][j] = sourceFontLines[i][j + 1];
      }
      values[i][values[i].length - 1] = "";
      fontWeights[i][values[i].length - 1] = fontWeights[i][values[i].length];
      fontFamilies[i][values[i].length - 1] = fontFamilies[i][values[i].length];
      sourceFontLines[i][values[i].length - 1] = sourceFontLines[i][values[i].length];
    }

    range.setValues(values);
    range.setFontWeights(fontWeights);
    range.setFontFamilies(fontFamilies);
    range.setFontLines(sourceFontLines);

    try {
      var targetRangeValues = [];
      var targetRangeFontStyles = [];
      var targetRangeFontFamilies = [];
      var targetRangeFontLines = [];

      for (var i = 0; i < sourceValues.length; i++) {
        var rowValues = source2Values[i];
        var rowFontWeights = source2FontWeights[i];
        var rowFontFamilies = source2FontFamilies[i];

        for (var j = 0; j < rowValues.length; j++) {
          var value = rowValues[j] !== "" ? rowValues[j] : "";
          var fontWeight = rowFontWeights[j] === "bold" ? "bold" : "normal";
          var fontFamily = rowFontFamilies[j];
          var fontLine = "none";

          if (fontWeight !== "bold"){
            value = "";
          }

          targetRangeValues.push([value]);
          targetRangeFontStyles.push([fontWeight]);
          targetRangeFontFamilies.push([fontFamily]);
          targetRangeFontLines.push([fontLine]);
        }
      }

      var targetRange = activeSheet.getRange(targetColumn + startRow + ":" + targetColumn + endRow);
      targetRange.clearContent();
      targetRange.setValues(targetRangeValues);
      targetRange.setFontWeights(targetRangeFontStyles);
      targetRange.setFontFamilies(targetRangeFontFamilies);
      targetRange.setFontLines(targetRangeFontLines);
    }
    catch(e){
      Logger.log(e);
    }

    scriptProperties.setProperty('lastExecutionDate', currentDate);
  }

  var nowTriggerText = setTrigger();
  var nowRow = findRowsWithText(activeSheet, 2, 5, 52, nowTriggerText);

  var cell = activeSheet.getRange(nowRow, column);
  var cellval = cell.getValue().toString();
  var cellFontLine = cell.getFontLine();

  var cellbuf = cellval.split(",");

  if (cellval !== "" && cellFontLine !== "line-through")
  {
    cell.setFontLine('line-through');

    if (cellbuf.length < 2){
      cellbuf = [
        cellbuf[0],
        ""
      ]
    }
    
    var insertText = `${currentDate} ${nowTriggerText.slice(-8,-3)} ${cellbuf[0]} ${cellbuf[1]}`;
    shiftAndInsertText(archiveSheet, insertText);

    var users = [];
    var roles = [];

    for(var i = 0; i < userIDs.length; i++){
      users.push(userIDs[i]);
    }
    for(var i = 0; i < roleIDs.length; i++){
      roles.push(roleIDs[i]);
    } 

    var textMessage = "";

    var mUser = users.map(userId => `<@${userId}>`).join(' ');
    if (mUser != '<@>') {
      textMessage += mUser;
    }
    var mRole = roles.map(roleId => `<@&${roleId}>`).join(' ');
    if (mRole != '<@&>') {
      textMessage += mRole;
    }
    
    // Discordに送信するメッセージの設定
    var colorCode = parseInt("219ddd", 16);
    var embeds = [
      { 
        "title": sendTitle,
        "color": colorCode,
        "fields": [
          {
            "name" : `${sendS}「${cellbuf[0]}」${sendE}`,
            "value": `${cellbuf[1]}`,
            "inline": false
          },
          {
            "name": "実行時間",
            "value": `${nowTriggerText.slice(-8,-3)}`,
            "inline": false
          }
        ],
      }
    ]
    
    //Discordにメッセージを送信する
    sendToDiscord(botUrl, textMessage, embeds, botName, botIcon);
  }
}

function sendToDiscord(url, content, embeds, name, icon) {
  var jsonData = {
    "username": name,
    "avatar_url": icon,
    "content": content,
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

function findRowsWithText(activeSheet, column, startRow, endRow, searchText) {
  var data = activeSheet.getRange(startRow, column, endRow - startRow + 1).getValues();
  var rowsWithText = [];

  for (var i = 0; i < data.length; i++) {
    var cellValue = data[i][0].toString();
    if (cellValue.includes(searchText)) {
      var rowIndex = startRow + i;
      rowsWithText.push(rowIndex);
    }
  }

  return rowsWithText;
}

function setTrigger() {
  var now = new Date();
  var minutes = now.getMinutes();
  var roundedMinutes = Math.floor(minutes / 30) * 30;
  var nextTrigger = new Date(now.getFullYear(), now.getMonth(), now.getDate(), now.getHours(), roundedMinutes + 30, 0);
  var delayMilliseconds = nextTrigger.getTime() - now.getTime();
  var formattedTime = now.getHours().toString().padStart(2, '0') + ":" + ("0" + roundedMinutes).slice(-2) + ":00" ;
  var triggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  ScriptApp.newTrigger("sendNotification").timeBased().after(delayMilliseconds).create();

  return formattedTime;
}

function shiftAndInsertText(sheet, insertText) {
  var lastRow = sheet.getLastRow();
  
  // 行数が上限に達している場合、最後の行を削除
  if (lastRow >= 100) {
    sheet.deleteRow(100);
    lastRow--;
  }
  
  if (lastRow != 0){
    var rangeToMove = sheet.getRange(1, 1, lastRow, 1);
    var destinationRange = sheet.getRange(2, 1, lastRow, 1);
    rangeToMove.copyTo(destinationRange);
  }
  
  sheet.getRange(1, 1).setValue(insertText);
}
