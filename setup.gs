/////////////////////////////////////////////////////////////////////////////////////////////////
//【関数名】
//シート作成関数
//【引数】
//※引数なし
/////////////////////////////////////////////////////////////////////////////////////////////////
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
  sheet.getRange('B20:C20').setValue('END');
  sheet.setNamedRange('SlackToken', sheet.getRange('C2'));
  sheet.setNamedRange('ToName', sheet.getRange('C3'));
  sheet.setNamedRange('URLRegs', sheet.getRange('B6:B20'));
  sheet.setNamedRange('MailTitles', sheet.getRange('C6:C20'));
  sheet.setNamedRange('END', sheet.getRange('B20:C20'));
  sheet.getRange('B20:C20').protect().setWarningOnly(true).setDescription("END");
  sheet.getRangeList(['B2:C3', 'B5:C20'])
       .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRangeList(['B2:B3', 'B5:C5']).setBackground('#c9daf8');
  sheet.getRange('B20:C20').setBackground('#b2b2b2');
  sheet.getRange('B4').setValue('※空白行かENDで終端を判定');
  sheet.getActiveSheet().setColumnWidth(1, 50);
  sheet.getActiveSheet().setColumnWidth(2, 150);
  sheet.getActiveSheet().setColumnWidth(3, 250);
};