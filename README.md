# mail2slack

## 概要

特定の件名を検知してslackにメール本文を転送する。
本文内のURLを正規表現で取得し、QRコードの転送も行う。
※QRコード化はGoogleのAPIを使用

## 事前準備

ForkやGitからソースコードをご自身のスクリプトエディタに移行できる方は、「使い方」からお読みください。

【前提条件】
* GitHubのアカウントでログイン済みであること
* Google Apps Script GitHub アシスタントを導入済みであること ※導入方法は[こちら](https://qiita.com/20731057hh/items/7f76f9e53e9da5c85ae9#%E6%A6%82%E8%A6%81)

【セットアップ】
1. 本リポジトリ（mail2slack)をForkする  
  画像右端にあるForkボタンから自分のリポジトリにmail2slackを取得できます。  
    <img src="https://github.com/20731057hh/mail2slack/blob/imags/fork.png" width="500">
1. スプレッドシートを新規作成する
1. スクリプトエディタからmail2slackをcloneする
1. スプレッドシートを再読み込みする
1. メニューバーに「mail2slack」が表示されていれば成功  
    <img src="https://github.com/20731057hh/mail2slack/blob/imags/menue.png" width="200">
1. 実行時にgoogleの承認が発生するので承認する
1. シート「Gmail転送Bot」が作成されたら成功

## 使い方
【前提条件】
* 「mail2slack」のメニューが表示されていること

【シート作成】
1. 「mail2slack」のメニューから「シート作成」を実行する  
    <img src="https://github.com/20731057hh/mail2slack/blob/imags/sheetCriate.png" width="200">
1. 初回実行時のみGoogle承認が表示されるので承認する
1. シート「Gmail転送Bot」が作成されたら成功  
    <img src="https://github.com/20731057hh/mail2slack/blob/imags/sheet.png" width="150">

【設定方法】
1. SlackのTokenを設定する  
　　<img src="https://github.com/20731057hh/mail2slack/blob/imags/token.png" width="300">
   * SlackのTokenの取得方法は[こちら](https://qiita.com/ykhirao/items/0d6b9f4a0cc626884dbb)
   * Slackアカウントの権限によってはTokenが取得できない場合もあります
1. Slackの送信先を設定する  
　　<img src="https://github.com/20731057hh/mail2slack/blob/imags/send.png" width="300">
   * 個人のDMに送信する場合は「@ + slackのユーザID(メールアドレスの@より前)」で設定してください
   * チャンネルに送信する場合は、#も含めてチャンネル名を設定してください
   * プライベートチャンネルの場合も#が必要です
1. QRコード化したいURLの正規表現を設定する  
　　<img src="https://github.com/20731057hh/mail2slack/blob/imags/url.png" width="150">
   * 複数の条件に一致する場合は、最初の１つ目が対象となります
   * 検索条件でQRコード化されるだけなので、URL以外の文字列でもQRコード化されます
1. 検索するメールの件名を設定する  
　　<img src="https://github.com/20731057hh/mail2slack/blob/imags/title.png" width="200">
   * 入力した件名に部分一致するタイトルのメールかつ、未読のメールがSlack送信対象となります
   * メールを転送したら既読に更新します
   * 受診時ではなく、未読メールを定期的にチェックしているので対象メールを未読にするとSlackに転送されます
   
【起動】
1. 「mail2slack」のメニューから「開始」を実行する  
   * Bot起動中は設定値の変更はできません
1. 送信先のSlackにて転送開始通知が受診できていればOK  
　　<img src="https://github.com/20731057hh/mail2slack/blob/imags/start.png" width="400">
   * 受診できていない場合は、Torkenや送信先の指定に誤りがないか確認してください
