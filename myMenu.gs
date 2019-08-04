function onOpen(){

  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  //メニュー配列
  var myMenu=[
    {name: "シート作成", functionName: "makeSheet"},
    {name: "開始", functionName: "Mail_forward_setTrigger"},
    {name: "停止", functionName: "Mail_forward_deleteTrigger"}
  ];

  sheet.addMenu("mail2Slack",myMenu); //メニューを追加

}