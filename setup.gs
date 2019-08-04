function makeSheet() {
  var res = Browser.msgBox("表示中のシートを編集しますが宜しいでしょうか？", "キャンセルで処理を中断します。", Browser.Buttons.OK_CANCEL);
  if (res == 'cancel') {
    return;
  }
  var sheet = SpreadsheetApp.getActive();
  
  sheet.getActiveSheet().setName('Gmail転送Bot');
  sheet.getRange('B2').setValue('Token');
  sheet.getRange('C2').setValue('※送信先のSlackトークン');
  sheet.getRange('B3').setValue('送信先');
  sheet.getRange('C3').setValue('※Slackの@ユーザID or #チャンネル名');
  sheet.getRange('B5').setValue('URL正規表現');
  sheet.getRange('C5').setValue('件名');
  sheet.getRange('B11:C11').setValue('END');
  sheet.setNamedRange('SlackToken', sheet.getRange('C2'));
  sheet.setNamedRange('ToName', sheet.getRange('C3'));
  sheet.setNamedRange('URLRegs', sheet.getRange('B6:B11'));
  sheet.setNamedRange('MailTitles', sheet.getRange('C6:C11'));
  sheet.getRangeList(['B2:C3', 'B5:C11']).activate()
  .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRangeList(['B2:B3', 'B5:C5']).activate()
  .setBackground('#c9daf8');
  sheet.getRange('B11:C11').activate();
  sheet.getActiveRangeList().setBackground('#cccccc');
  sheet.getRange('B4').setValue('※空白行かENDで終端を判定');
  sheet.getActiveSheet().setColumnWidth(1, 50);
  sheet.getActiveSheet().setColumnWidth(2, 150);
  sheet.getActiveSheet().setColumnWidth(3, 250);
};