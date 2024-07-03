function getUserName(userId) {
        
  const url = "https://api.line.me/v2/bot/profile/" + userId;
  const response = UrlFetchApp.fetch(url, {
    "headers" : {
      "Authorization" : "Bearer " + "g1uBnwKvSelVhIrXVW3fMmK2b8h75ZVQRheZkf2gVuFP/TxBmFjpnXde4ijmhTgPSeWk7mFrQlZzAMpilUJVCPQhN4qMTZMxoTfbeGBLYsWCoDerMTLO49tikkwGmkgdUJehVJPXmzL1I+kseaI+tQdB04t89/1O/w1cDnyilFU="
      }
    });

  return JSON.parse(response.getContentText()).displayName;
}

function changeGeneration(){

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("develop");
  var currentIndex = sheet.getRange(2, 8).getValue();

  for(var i = 2 ; i < sheet.getLastRow() + 1 ; i++){

    if(sheet.getRange(i, 6).getValue() == "学年区分" + currentIndex){

      sheet.getRange(i, 1, 1, 6).clearContent();
    }
  }

  currentIndex--;
  if(currentIndex == 0){

    currentIndex = 3;
  }

  sheet.getRange(2, 8).setValue(currentIndex);


  sheet.getRange(2, 1, sheet.getLastRow()+1, 6).sort({column: 6, ascending: false});
}

function doPost(e){

  let data = JSON.parse(e.postData.contents);
  let events = data.events;

  for(let i = 0; i < events.length; i++){
    let event = events[i];

    if(event.type == "unfollow"){

      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("develop");  // シートを取得
      const searchRange = sheet.getRange("A2:A");
      const searchString = event.source.userId;

      const textObject = searchRange.createTextFinder(searchString).matchEntireCell(true);
      const result = textObject.findAll();
      const id_row = result[0].getRow();

      sheet.getRange(id_row, 1, 1, 6).clearContent();
      sheet.getRange(2, 1, sheet.getLastRow() + 1, 6).sort({column: 6, ascending: false});
    }
    if(event.type == 'follow'){

        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("develop");  // シートを取得
        var lastRow = sheet.getLastRow();  // 最終行を取得
        sheet.getRange(lastRow + 1,1).setValue(event.source.userId);  // A列目にユーザID記入
        sheet.getDataRange().removeDuplicates([1]);  // ユーザIDの重複を削除

        const searchRange = sheet.getRange("A2:A");
        const searchString = event.source.userId;

        const textObject = searchRange.createTextFinder(searchString).matchEntireCell(true);
        const result = textObject.findAll();
        const id_row = result[0].getRow();

        sheet.getRange(id_row, 2).setValue(0);
        sheet.getRange(id_row, 4).setValue(0);
        //sheet.getRange(id_row, 5).setValue(getUserName(event.source.userId));
        sheet.getRange(id_row, 6).clearContent();

        var gradeIndex = sheet.getRange(2, 8).getValue();

        var gradeIndex_first = gradeIndex % 3 + 1;
        var gradeIndex_second = (gradeIndex + 1) % 3 + 1;


        let contents = {
          replyToken: event.replyToken, 
          messages: [
            {
              type: 'template',
              altText: '学年確認',
              template: {
                type: 'confirm',
                text: 'あなたの学年を教えてください',
                actions: [
                    {
                      type: 'message',
                      label: '１年生',
                      text: '学年区分' + gradeIndex_first
                    },
                    {
                      type: 'message',
                      label: '２年生',
                      text: '学年区分' + gradeIndex_second
                    },
                    
                ]
              }
            }
          ]
        };

        reply(contents);
    }
    if(event.type == 'message'){

      if(event.message.type == 'text'){ // 受信したのが普通のテキストメッセージだったとき

        let return_text = "";
        let contents;
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("develop"); 
        const searchRange = sheet.getRange("A2:A");
        const searchString = event.source.userId;

        const textObject = searchRange.createTextFinder(searchString).matchEntireCell(true);
        const result = textObject.findAll();
        const id_row = result[0].getRow();
        var absent_flag = sheet.getRange(id_row, 2).getValue();
        var groupsep_flag = sheet.getRange(id_row, 4).getValue();

        var gradeClassBoolean = sheet.getRange(id_row, 6).isBlank();
        var nameFilledBoolean = sheet.getRange(id_row, 5).isBlank();
        
        if(event.message.text == "キャンセル"){

          if(absent_flag == 1){

            sheet.getRange(id_row, 2).setValue(0);
            absent_flag = 0;
            return_text = "欠席連絡をキャンセルしました\n";
          }
          else if(groupsep_flag == 1 || groupsep_flag == 2){

            sheet.getRange(id_row, 4).setValue(0);
            groupsep_flag = 0;
            return_text = "グループ作成をキャンセルしました\n";
          }
        }
        if(event.message.text == "続行"){

          if(groupsep_flag == 2){

            var sheet_groups = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Groups");
            var num = sheet_groups.getRange("A1").getValue();

            sheet.getRange(id_row, 4).setValue(0);
            groupsep_flag = 0;
            groupManage(num);


            //表のヘッダーとデータ範囲を取得 ※１
            var header = sheet_groups.getRange(1, 1, 1, num).getDisplayValues();
            var values = sheet_groups.getRange(2, 1, sheet_groups.getLastRow() - 1, num).getDisplayValues();

            //データテーブルを作成
            var table = Charts.newDataTable();

            //ヘッダー行を入力
            for(let i=0; i<header[0].length; i++) {
              table.addColumn(Charts.ColumnType.STRING, header[0][i]);
            }

            //データ範囲を入力
            for(var j=0; j<values.length; j++) {
              table.addRow(values[j]);
            }

            //表グラフを作成＆画像化
            const blob = Charts.newTableChart()
            .setDataTable(table.build())
            .setDimensions(400, 150 + (sheet_groups.getLastRow() - 1) * 20)
            .build()
            .getAs('image/png');

            //画像をシートに挿入
            //sheet_groups.insertImage(blob, 'A1');

            var imageFile = DriveApp.createFile(blob).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            var imageUrl = imageFile.getDownloadUrl();

            return_text = "グループが作成されました";

            // 送信するデータをオブジェクトとして作成する
            contents = {
              replyToken: event.replyToken, // event.replyToken は受信したメッセージに含まれる応答トークン
              messages: [
                { 
                  type: 'text', 
                  text:  return_text 
                },
                {
                  type: 'image',
                  originalContentUrl: imageUrl,
                  previewImageUrl: imageUrl
                }
              ],
            };
            reply(contents);

            imageFile.setTrashed(true);
          }

          if(absent_flag == 1){

            sheet.getRange(id_row, 2).setValue(0);
            absent_flag = 0;

            sheet.getRange(id_row, 3).clearContent();
            return_text = "欠席連絡を削除しました";
          }
        }
        if(gradeClassBoolean){

          sheet.getRange(id_row, 6).setValue(event.message.text);

          sheet.getRange(2, 1, sheet.getLastRow() + 1, 6).sort({column: 6, ascending: false});
          return_text = "学年区分を設定しました\n次に氏名を教えてください";
        }
        if(!gradeClassBoolean && nameFilledBoolean){

          sheet.getRange(id_row, 5).setValue(event.message.text);
          return_text = event.message.text + "　さんで登録しました\n変更する場合は管理者に連絡してください";
        }
        if(absent_flag == 1){
          
          if(event.message.text == "欠席連絡"){

            absent_flag = 0;
          }
          else if(event.message.text == "グループ作成"){

            absent_flag = 0;
            return_text = "欠席連絡をキャンセルしました\n";
          }
          else{

            sheet.getRange(id_row, 3).setValue(event.message.text);
            sheet.getRange(id_row, 2).setValue(0);
            return_text = "欠席理由を「" + event.message.text +  "」で登録しました";
          }
        }
        if(groupsep_flag == 1){

          var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("develop");
          var sheet_attendance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance");
          var sheet_groups = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Groups"); 

          var num = event.message.text;
          if(isFinite(num) && num >= 2 ){
            sheet_groups.getRange("A1").setValue(event.message.text);

            for(let i = 1 ; i < sheet_attendance.getLastColumn() + 1; i++){

              for(let j = 1 ; j < sheet_attendance.getLastRow() + 1; j++){

                return_text += sheet_attendance.getRange(j, i).getValue();
                return_text += "\n";
              }

              return_text += "\n";
            }

            contents = {
              replyToken: event.replyToken, 
              messages: [
                {
                  type: 'text',
                  text: return_text
                },
                {
                  type: 'template',
                  altText: '確認',
                  template: {
                    type: 'confirm',
                    text: '以上のメンバで' + num + '個のグループを作成します\nよろしいですか？',
                    actions: [
                        {
                            type: 'message',
                            label: '続行',
                           text: '続行'
                        },
                        {
                            type: 'message',
                            label: 'キャンセル',
                            text: 'キャンセル'
                        }
                    ]
                  }
                }
              ]
            };

            sheet.getRange(id_row, 4).setValue(2);
            reply(contents);
          }
          else if(event.message.text == "キャンセル"){

            groupsep_flag = 0;
            return_text = "グループ作成をキャンセルしました";
          }
          else{

            sheet.getRange(id_row, 4).setValue(1);
            return_text = "無効な値です\nもう一度入力してください";
          }

          
        }

        if(event.message.text == "欠席連絡"){

          if(absent_flag != 1){

            sheet.getRange(id_row, 2).setValue(1);
            if(!sheet.getRange(id_row, 3).isBlank()){


              var reason = sheet.getRange(id_row, 3).getValue();
              contents = {
                replyToken: event.replyToken, 
                messages: [
                  {
                    type: 'template',
                    altText: '確認',
                    template: {
                      type: 'confirm',
                      text: '既に欠席連絡が\n「' + reason +'」\nで登録されています\n続行すると登録が削除されます',
                      actions: [
                          {
                              type: 'message',
                              label: '続行',
                             text: '続行'
                          },
                          {
                              type: 'message',
                              label: 'キャンセル',
                              text: 'キャンセル'
                          }
                      ]
                    }
                  }
                ]
              };

              reply(contents);
            }
            else{

              return_text = "欠席理由を教えてください\n理由の送信を以って欠席連絡が確定します";
            }
          }
        }
        if(event.message.text == "グループ作成"){

          var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("develop");
          var sheet_attendance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance");
          var sheet_groups = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Groups"); 
          
          var index1 = 2;
          var index2 = 2;
          var index3 = 2;

          sheet_attendance.clear();

          sheet_attendance.getRange(1, 1).setValue("学年区分1");
          sheet_attendance.getRange(1, 2).setValue("学年区分2");
          sheet_attendance.getRange(1, 3).setValue("学年区分3");

          for(let i = 2 ; i < sheet.getLastRow() + 1; i++){

            if(sheet.getRange(i, 3).isBlank()){

              if(sheet.getRange(i, 6).getValue() == "学年区分1"){

                sheet_attendance.getRange(index1++, 1).setValue(sheet.getRange(i, 5).getValue());
              }
              else if(sheet.getRange(i, 6).getValue() == "学年区分2"){

                sheet_attendance.getRange(index2++, 2).setValue(sheet.getRange(i, 5).getValue());
              }
              else if(sheet.getRange(i, 6).getValue() == "学年区分3"){

                sheet_attendance.getRange(index3++, 3).setValue(sheet.getRange(i, 5).getValue());
              }
            }
          }

          for(let i = 1 ; i < sheet_attendance.getLastColumn() + 1; i++){

            for(let j = 1 ; j < sheet_attendance.getLastRow() + 1; j++){

              return_text += sheet_attendance.getRange(j, i).getValue();
              return_text += "\n";
            }

            return_text += "\n";
          }

          sheet.getRange(id_row, 4).setValue(1);
          return_text += "上のメンバが出席予定です\n何グループ作成しますか？\n（半角数字）";
        }
        if(event.message.text == "dev"){

          if(event.message.text == "Uf369a26de05cb0a3a4cce3ff739ffe44"){

            return_text = "開発者向けリンク\n\nhttps://docs.google.com/spreadsheets/d/1OJUXJJY2hItbeNKzHD4qj1bWGCFhADtFJEzAFpGdTBM/edit?usp=sharing";
          }
          else{

            return_text = "実行権限がありません\nこの機能は開発者用です";
          }
        }
        if(event.message.text == "世代交代"){

          if(event.source.userId == sheet.getRange(2, 9).getValue()){

            if("学年区分" + sheet.getRange(2, 8).getValue() != sheet.getRange(id_row, 6).getValue()){

              changeGeneration();
              return_text = "世代を交代しました";
            }
            else{

              return_text = "あなたは最高学年なので実行できません\n委員長を交代して次期委員長に実行させてください";
            }
          }
          else{

            return_text = "この機能は委員長のみに権限があります";
          }
        }
        if(event.message.text == "欠席状況確認"){

          if(absent_flag == 1){

            absent_flag = 0;
            return_text = "欠席連絡をキャンセルしました\n";
          }
          else if(groupsep_flag != 0){

            groupsep_flag = 0;
            return_text = "グループ作成をキャンセルしました\n\n";
          }

          for(let i = 2; i < sheet.getLastRow() + 1 ; i++){
            
            if(!sheet.getRange(i, 3).isBlank()){

              return_text += sheet.getRange(i, 5).getValue() + "：「" + sheet.getRange(i, 3).getValue() + " 」\n\n";
            }
          }
        }
        if(event.message.text == "委員長交代"){

          var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("develop");
          if(sheet.getRange(2, 9).getValue() == event.source.userId){

            code = Math.floor(Math.random() * (9999 - 1000) + 1000);

            sheet.getRange(3, 9).setValue(code);
            return_text = "委員長を交代するには以下の認証コードを次期委員長が送信してください\n\n認証コード : " + code;
          }
          else{

            return_text = "この機能は委員長のみに権限があります";
          }
        }
        if(event.message.text == sheet.getRange(3, 9).getValue()){

          sheet.getRange(2, 9).setValue(event.source.userId);
          sheet.getRange(3, 9).clearContent();

          return_text = "委員長を交代しました";
        }
        
        // 送信するデータをオブジェクトとして作成する
        contents = {
          replyToken: event.replyToken, // event.replyToken は受信したメッセージに含まれる応答トークン
          messages: [{ type: 'text', text:  return_text }],
        };
        reply(contents);
      }
    }
  }
}

function reply(contents){
  let channelAccessToken = "g1uBnwKvSelVhIrXVW3fMmK2b8h75ZVQRheZkf2gVuFP/TxBmFjpnXde4ijmhTgPSeWk7mFrQlZzAMpilUJVCPQhN4qMTZMxoTfbeGBLYsWCoDerMTLO49tikkwGmkgdUJehVJPXmzL1I+kseaI+tQdB04t89/1O/w1cDnyilFU=";
  let replyUrl = "https://api.line.me/v2/bot/message/reply"; // LINE にデータを送り返すときに使う URL
  let options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + channelAccessToken
    },
    payload: JSON.stringify(contents) // リクエストボディは payload に入れる
  };
  UrlFetchApp.fetch(replyUrl, options);
}

function groupManage(numGroups){

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Attendance');
  var outputSheet = ss.getSheetByName('Groups'); // 結果出力用のシート名

  // コミュニティメンバーのリストを取得
  var community_1 = sheet.getRange('A2:A').getValues().flat().filter(String);
  var community_2 = sheet.getRange('B2:B').getValues().flat().filter(String);
  var community_3 = sheet.getRange('C2:C').getValues().flat().filter(String);

  // 結果シートをクリア
  outputSheet.clear();

  var lastIdx = 0;
  lastIdx = makeGroups(community_1, numGroups, outputSheet, lastIdx);
  lastIdx = makeGroups(community_2, numGroups, outputSheet, lastIdx);
  makeGroups(community_3, numGroups, outputSheet, lastIdx);
}

function makeGroups(community, numGroups, outputSheet, lastIndex) {

  // メンバーをシャッフル
  community.sort(function() { return 0.5 - Math.random(); });

  // 結果シートの最後の行を見つける
  var lastRow = outputSheet.getLastRow();
  var startRow = lastRow + 1;

  if (lastRow == 0) {
    // グループラベルを追加

    for(let index = 0; index < numGroups; index++){

      outputSheet.getRange(startRow, index + 1).setValue('Group ' + (index + 1));
    }
    startRow++;
  }

  // グループを結果シートに追記
  console.log(community);

  if(lastIndex != 0){

    startRow = lastRow;
    row--;
  }

  var row = 0, line = lastIndex;
  for(var index = 0; index < community.length ; index++){

    var value = community[index];
    outputSheet.getRange(row + startRow, line + 1).setValue(value);
    console.log(row + startRow, line + 1, value);
    lastIndex = ++line;

    if(line > numGroups - 1){

      row++;
      line = 0;
    }
  }

  console.log(lastIndex)
  return lastIndex % numGroups;
}
