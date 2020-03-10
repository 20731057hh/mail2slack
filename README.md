# mail2slack

## 概要

特定の件名を検知してslackにメール本文を転送する。
本文内のURLを正規表現で取得し、QRコードの転送も行う。
※QRコードはGoogleのAPIを使用

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
1. 「mail2slack」のメニューが表示されるので、「シート作成」を実行する  
    <img src="https://github.com/20731057hh/mail2slack/blob/imags/sheetCriate.png" width="200">
1. 実行時にgoogleの承認が発生するので承認する
1. シート「Gmail転送Bot」が作成されたら成功


