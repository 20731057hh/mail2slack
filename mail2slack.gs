var warningStatement = "トリガー起動中のため編集禁止です。トリガーを停止してください。"

/////////////////////////////////////////////////////////////////////////////////////////////////
//【関数名】
//カスタムトリガー登録関数
//【引数】
//※引数なし
/////////////////////////////////////////////////////////////////////////////////////////////////
function Mail_forward_setTrigger() {
  //実行中かをチェックする
  var triggers = ScriptApp.getProjectTriggers();
  for(var i=0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == "Mail_forward") {
      Browser.msgBox( "すでに実行中のため、処理を終了します。");
      return;
    }
  }

  //設定内容を取得
  var bk = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = bk.getSheetByName('Gmail転送Bot');
  var NameRanges = bk.getNamedRanges();
  var ToName = "";
  var SlackToken = "";
  var Nnum = 0;
  for(var NRange in NameRanges){
    var tmp = NameRanges[NRange].getName();
    if(tmp == "ToName"){
      ToName = NameRanges[NRange].getRange().getValue();
      if(ToName == "" || (ToName[0] != "@" && ToName[0] != "#")){
        Browser.msgBox( "Slack送信先が不正のため、処理を終了します。");
        return;
      };
      Nnum += 1;
    }
    else if(tmp == "URLRegs"){
      Nnum += 1;
    }
    else if(tmp == "MailTitles"){
      Nnum += 1;
    }
    else if(tmp == "SlackToken"){
      var reg = /^[0-9a-zA-Z]*-[0-9a-zA-Z]*-[0-9a-zA-Z]*-[0-9a-zA-Z]*-[0-9a-zA-Z]*/;
      SlackToken = NameRanges[NRange].getRange().getValue();
      if (SlackToken.match(reg) == null){
        Browser.msgBox( "Slackトークンが不正のため、処理を終了します。");
        return;
      }
      SlackToken = NameRanges[NRange].getRange().getValue();
      Nnum += 1;
    };
  };
  if(Nnum != 4){
    Browser.msgBox( "「名前付き範囲」が設定されていないため、処理を終了します。");
    return;
  };

  //Slack転送トリガーの登録
  var setTriggerParam_Mail_forward = {
    ver: 2,
    setMinutes: 1,
    setFunc: "Mail_forward"
  };
  setTriggerEM(setTriggerParam_Mail_forward);

  //トリガー停止トリガーの登録
  var date = new Date();
  var setTriggerParam_deleteTrigger = {
    ver: 3,
    setDate: new Date(date.setDate(date.getDate() + 7)),
    setFunc: "Mail_forward_deleteTrigger"
  };
  Logger.log(setTriggerParam_deleteTrigger.setDate);
  setTriggerDodate(setTriggerParam_deleteTrigger);

  for(var NRange in NameRanges){
    var tmp = NameRanges[NRange].getName();
    if(tmp == "ToName"){
      NameRanges[NRange].getRange().setBackground("#808080")
      var protection = NameRanges[NRange].getRange().protect();
      protection.setDescription(warningStatement);
      protection.setWarningOnly(true);
    }
    else if(tmp == "URLRegs"){
      NameRanges[NRange].getRange().setBackground("#808080")
      var protection = NameRanges[NRange].getRange().protect();
      protection.setDescription(warningStatement);
      protection.setWarningOnly(true);
    }
    else if(tmp == "MailTitles"){
      NameRanges[NRange].getRange().setBackground("#808080")
      var protection = NameRanges[NRange].getRange().protect();
      protection.setDescription(warningStatement);
      protection.setWarningOnly(true);
    }
    else if(tmp == "SlackToken"){
      NameRanges[NRange].getRange().setBackground("#808080")
      var protection = NameRanges[NRange].getRange().protect();
      protection.setDescription(warningStatement);
      protection.setWarningOnly(true);
    };
  };

  //開始Slackの送信
  var attachments = JSON.stringify([{
    title: "★認証メールのSlack転送を開始します★ ※1週間後に自動停止します。",
    title_link: bk.getUrl(),
    text: "※クリックするとツールのスプレッドシートが表示されます。",
    color: "#7CFC00"
  }]);
  var payload = {
    "channel": ToName,
    "attachments": attachments,
    "username": "仮URLメールのSlack転送Bot",
    "parse": 'full',
    "icon_emoji": ':email:'
  };
  var option = {
    "method": "POST", //POST送信
    "payload": payload //POSTデータ
  };
  UrlFetchApp.fetch("https://slack.com/api/chat.postMessage?token=" + SlackToken, option);

  Browser.msgBox( "認証メールのSlack転送を開始します。※1週間後に自動停止します。");
}

/////////////////////////////////////////////////////////////////////////////////////////////////
//【関数名】
//カスタムトリガーの停止関数
//【引数】
//※引数なし
/////////////////////////////////////////////////////////////////////////////////////////////////
function Mail_forward_deleteTrigger() {
  //登録したカスタムトリガーの削除
  deleteTrigger("Mail_forward");
  deleteTrigger("Mail_forward_deleteTrigger");

  //設定内容を取得
  var bk = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = bk.getSheetByName('Gmail転送Bot');
  var NameRanges = bk.getNamedRanges();
  var ToName = "";
  var SlackToken = "";
  var Nnum = 0;
  for(var NRange in NameRanges){
    var tmp = NameRanges[NRange].getName();
    if(tmp == "ToName"){
      ToName = NameRanges[NRange].getRange().getValue();
      Nnum += 1;
    }
    else if(tmp == "URLRegs"){
      Nnum += 1;
    }
    else if(tmp == "MailTitles"){
      Nnum += 1;
    }
    else if(tmp == "SlackToken"){
      SlackToken = NameRanges[NRange].getRange().getValue();
      Nnum += 1;
    };
  };
  if(Nnum != 4){
    Browser.msgBox( "編集制限の解除に失敗しました。シート保護の範囲を確認してください。");
    return;
  };

  //実行中の制限を解除
  var Protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  var endReng = null;
  for(var Protection in Protections){
    if(Protections[Protection].getDescription() == warningStatement){
      Protections[Protection].getRange().setBackground(null);
      Protections[Protection].remove();
    }else if(Protections[Protection].getDescription() == "END"){
      endReng = Protections[Protection].getRange();
    }
  };
  if(endReng != null){
    endReng.setBackground("#b2b2b2")
  }

  //停止Slackの送信
  var attachments = JSON.stringify([{
    title: "★認証メールのSlack転送を停止しました★",
    title_link: bk.getUrl(),
    text: "※クリックするとツールのスプレッドシートが表示されます。",
    color: "#7CFC00"
  }]);
  var payload = {
    "channel": ToName,
    "attachments": attachments,
    "username": "仮URLメールのSlack転送Bot",
    "parse": 'full',
    "icon_emoji": ':email:'
  };
  var option = {
    "method": "POST", //POST送信
    "payload": payload //POSTデータ
  };
  UrlFetchApp.fetch("https://slack.com/api/chat.postMessage?token=" + SlackToken, option);
}

/////////////////////////////////////////////////////////////////////////////////////////////////
//【関数名】
//Slackへのメール転送関数
//【引数】
//※トリガー実施のため引数無し
/////////////////////////////////////////////////////////////////////////////////////////////////
function Mail_forward() {

  //設定内容を取得
  var bk = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = bk.getSheetByName('Gmail転送Bot');
  var NameRanges = bk.getNamedRanges();
  var ToName = "";
  var SlackToken = "";
  var URLReg = "";
  var MailTitle = ""
  var URLRegs = "";
  var MailTitles = "";
  var Nnum = 0;
  
  for(var NRange in NameRanges){
    var tmp = NameRanges[NRange].getName();
    if(tmp == "ToName"){
      ToName = NameRanges[NRange].getRange().getValue();
      Nnum += 1;
    }
    else if(tmp == "URLRegs"){
      URLRegs = NameRanges[NRange].getRange().getValues();
      Nnum += 1;
    }
    else if(tmp == "MailTitles"){
      MailTitles = NameRanges[NRange].getRange().getValues();
      Nnum += 1;
    }
    else if(tmp == "SlackToken"){
      SlackToken = NameRanges[NRange].getRange().getValues();
      Nnum += 1;
    };
  }
  if(Nnum != 4){
    Browser.msgBox( "「名前付き範囲」が設定されていないため、処理を終了します。");
    Mail_forward_deleteTrigger();
    return;
  }
  for( var URLReg in URLRegs){
    if(URLRegs[URLReg] == "" || URLRegs[URLReg] == "END"){ break; }
    //指定条件でメールを検索
    var threads = GmailApp.search("subject:("+MailTitles[URLReg]+") label:unread ");
    //スレッドからメールを取得する　→二次元配列で格納
    var msgs = GmailApp.getMessagesForThreads(threads);
    //対象メッセージ取得
    for (var i = 0; i < threads.length; i++) { //スレッドの数だけ
      for (var j = 0; j < msgs[i].length; j++) { //メールの数だけ
        if(msgs[i][j].isUnread()){
          Logger.log("i:["+i+"]j:["+j+"]===============================================");
          postSlack(SlackToken, msgs[i][j].getSubject(),msgs[i][j].getPlainBody(),ToName,URLRegs[URLReg])
          msgs[i][j].markRead();//既読に変更
        }
      }
    }
  }
}

/////////////////////////////////////////////////////////////////////////////////////////////////
//【関数名】
//カスタムトリガー登録関数（毎分間隔）
//【引数】
//var setTriggerParam = {
//  ver: 2,
//  setMinutes: 実行分,
//  setFunc: 実行関数
//  };
/////////////////////////////////////////////////////////////////////////////////////////////////
function setTriggerEM(p) {
  ScriptApp.newTrigger(p.setFunc).timeBased().everyMinutes(p.setMinutes).create();
}

/////////////////////////////////////////////////////////////////////////////////////////////////
//【関数名】
//カスタムトリガー登録関数 (日時指定)
//【引数】
//var setTriggerParam = {
//  ver: 3,
//  setDate: 実行日時,
//  setFunc: 実行関数
//  };
/////////////////////////////////////////////////////////////////////////////////////////////////
function setTriggerDodate(p) {
  ScriptApp.newTrigger(p.setFunc).timeBased().at(p.setDate).create();
}

/////////////////////////////////////////////////////////////////////////////////////////////////
//【関数名】
//トリガー削除関数
//【引数】
//func: 削除対象関数名
/////////////////////////////////////////////////////////////////////////////////////////////////
function deleteTrigger(func) {
  var triggers = ScriptApp.getProjectTriggers();
  for(var i=0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == func) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/////////////////////////////////////////////////////////////////////////////////////////////////
//【関数名】
//Slack送信関数
//【引数】
//subject：メールタイトル
//body：メール本文
//ToName：Slack送信先
//URLReg：認証URLの正規表現
/////////////////////////////////////////////////////////////////////////////////////////////////
function postSlack(SlackToken, subject, body, ToName, URLReg)  {
  var rep; //検索した文字列を格納する変数

  //設定内容を取得
  var bk = SpreadsheetApp.getActiveSpreadsheet();
  //認証URLを取得
  var regexp = new RegExp(URLReg);
  rep = body.match(regexp);
  var slackTitle = "";
  var url = "";
  if(rep == null){
    slackTitle = "！！！認証URLの取得に失敗！！！";
    url = bk.getUrl();
  }else{
    slackTitle = "https://chart.apis.google.com/chart?cht=qr&chs=500x500&chl=" + rep[0].replace(/\r\n/,"");
    url = "https://chart.apis.google.com/chart?cht=qr&chs=500x500&chl=" + rep[0].replace(/\r\n/,"");
  }

  var attachments = JSON.stringify([{
    title: subject,
    text: "\n```" + body + "```",
    mrkdwn_in: ["text", "pretext"],
    color: "#7CFC00"
  }]);
  var payload = {
    "channel": ToName,
    "attachments": attachments,
    "username": 'Gmail転送BOT',
    "parse": 'full',
    "icon_emoji": ':email:'
  };
  var option = {
    "method": "POST", //POST送信
    "payload": payload //POSTデータ
  };
  var attachments2 = JSON.stringify([{
    title: slackTitle,
    title_link: url,
    text: "※URLをクリックするとブラウザでQRコードが表示されます。",
    color: "#7CFC00"
  }]);
  var payload2 = {
    "channel": ToName,
    "attachments": attachments2,
    "username": "認証用URLのQRコード",
    "parse": 'full',
    "icon_emoji": ':email:'
  };
  var option2 = {
    "method": "POST", //POST送信
    "payload": payload2 //POSTデータ
  };
  UrlFetchApp.fetch("https://slack.com/api/chat.postMessage?token=" + SlackToken, option);
  UrlFetchApp.fetch("https://slack.com/api/chat.postMessage?token=" + SlackToken, option2);
}