function doPost(e){

  let data = JSON.parse(e.postData.contents);
  let events = data.events;

  for(let i = 0; i < events.length; i++){
    let event = events[i];

    if(event.type == 'follow'){

        var sp_data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("develop");  // シートを取得
        var lastRow = sp_data.getLastRow();  // 最終行を取得
        sp_data.getRange(lastRow + 1,1).setValue(event.source.userId);  // A列目にユーザID記入
        sp_data.getDataRange().removeDuplicates([1]);  // ユーザIDの重複を削除

        var lastRow_abf = sp_data.getLastRow();
        if(sp_data.getRange(lastRow_abf, 2).isBlank()){

          sp_data.getRange(lastRow_abf, 2).setValue(0);
          sp_data.getRange(lastRow_abf, 4).setValue(0);
        }
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
        
        if(event.message.text == "キャンセル"){

          if(absent_flag == 1){

            sheet.getRange(id_row, 2).setValue(0);
            absent_flag = 0;
            return_text = "欠席連絡をキャンセルしました\n";
          }
          else if(groupsep_flag == 1){

            sheet.getRange(id_row, 4).setValue(0);
            groupsep_flag = 0;
            return_text = "グループ作成をキャンセルしました\n";
          }
        }
        if(event.message.text == "続行"){

          if(groupsep_flag == 1){

            sheet.getRange(id_row, 4).setValue(0);
            groupsep_flag = 0;
            main();

            return_text = "グループが作成されました";
          }
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
          
          let j = 1;
          for(let i = 1; i < sheet.getLastRow() + 1 ; i++){

            if(sheet.getRange(i, 3).isBlank()){
            
              sheet_attendance.getRange(j++, 5).setValue(sheet.getRange(i, 5).getValue());
            }
          }

          sheet.getRange(id_row, 4).setValue(0);
        }
        if(event.message.text == "欠席連絡"){

          if(absent_flag != 1){

            sheet.getRange(id_row, 2).setValue(1);
            return_text = "欠席理由を教えてください\n理由の送信を以って欠席連絡が確定します";
          }
        }
        if(event.message.text == "グループ作成"){
          
          sheet.getRange(id_row, 4).setValue(1);

          var sheet_community = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Community"); 
          var sheet_attendance = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance");

          for(let i = 1 ; i < sheet_community.getLastColumn() + 1; i++){
            for(let j = 1 ; j < sheet_community.getLastRow() + 1; j++){

              sheet_attendance.getRange(j, i).setValue(sheet_community.getRange(j, i).getValue());
            }
          }

          for(let i = 1 ; i < sheet_community.getLastColumn() + 1; i++){

            for(let j = 1 ; j < sheet_community.getLastRow() + 1; j++){

              return_text += sheet_community.getRange(j, i).getValue();
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
                  text: '以上のメンバでグループ作成を行います\nよろしいですか？',
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
        if(event.message.text == "dev"){

          return_text = "開発者向けリンク\n\nhttps://docs.google.com/spreadsheets/d/1OJUXJJY2hItbeNKzHD4qj1bWGCFhADtFJEzAFpGdTBM/edit?usp=sharing";
        }
        if(event.message.text == "gas"){

          return_text = "https://script.google.com/home/projects/1jZP8-ORDK_xLEzfolVtjtG-crqLXhxqsR8BasivLhKeJumfW02Emsjfs/edit";
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

function main(){

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Attendance');    // コミュニティメンバーのシート名
  var outputSheet = ss.getSheetByName('Groups'); // 結果出力用のシート名

  // コミュニティメンバーのリストを取得
  var community_1 = sheet.getRange('A2:A').getValues().flat().filter(String);
  var community_2 = sheet.getRange('B2:B').getValues().flat().filter(String);
  var community_3 = sheet.getRange('C2:C').getValues().flat().filter(String);
  
  // 欠席者のリストを取得
  var absentees = [];//sheet.getRange('D2:D').getValues().flat().filter(String);

  // グループ数を取得
  var numGroups = 3;//sheet.getRange('F1').getValue();

  // 結果シートをクリア
  outputSheet.clear();

  var lastIdx = 0;
  lastIdx = makeGroups(community_1, absentees, numGroups, outputSheet, lastIdx);
  lastIdx = makeGroups(community_2, absentees, numGroups, outputSheet, lastIdx);
  makeGroups(community_3, absentees, numGroups, outputSheet, lastIdx);
}

function makeGroups(community, absentees, numGroups, outputSheet, lastIndex) {

    // 欠席者を除外
  var presentMembers = community.filter(function(member) {
    return absentees.indexOf(member) === -1;
  });

  // メンバーをシャッフル
  presentMembers.sort(function() { return 0.5 - Math.random(); });

  // グループを初期化
  //var groups = Array.from({ length: numGroups }, () => []);

  // メンバーを均等にグループに分ける
  //var groupIndex = 0;
  //presentMembers.forEach(function(member) {
  //  groups[groupIndex].push(member);
  //  groupIndex = (groupIndex + 1) % numGroups;
  //});

  // 結果シートの最後の行を見つける
  var lastRow = outputSheet.getLastRow();
  var startRow = lastRow + 1;

  if (lastRow == 0) {
    // グループラベルを追加
    //presentMembers.forEach(function(presentMembers, index) {
    //  outputSheet.getRange(startRow, index + 1).setValue('Group ' + (index + 1));
    //});

    for(let index = 0; index < numGroups + 1; index++){

      outputSheet.getRange(startRow, index + 1).setValue('Group ' + (index + 1));
    }
    startRow++;
  }

  // グループを結果シートに追記
  var col = lastIndex;
  //if(col != 0){

  //  startRow = lastRow;
  //}

  console.log(presentMembers);
  for (var i = 0; i <= presentMembers.length / (numGroups + 2); i++) {
    for (var j = col; j < numGroups + col + 1; j++) {

      if(j > numGroups + 1){

        i++;
      }
      outputSheet.getRange(i + startRow, j + 1).setValue(presentMembers[i * (numGroups + col + 1) + j - col]);
      console.log(i + startRow, j, presentMembers[i * (numGroups + col +  1) + j - col]);
      lastIndex = j -col;
    }
  }

  return lastIndex % numGroups;
}
