Google Carendar Library for VBA

今のところ、Google カレンダーに予定を追加するクラスと、デモ用のコードしかありません。
発展途上です。

■ デモ

Excelファイルにモジュールをインポートしてください。

GoogleCalendar.cls
Module1.bas

GoogleCalendar.cls のコメントを参照して、参照設定してください。

セルに情報を入力してください。

A1：ユーザー名（メールアドレス）
A2：パスワード
A3：entry.xmlの内容

Alt + F8 を押して、test を実行してください。
2011年4月30日（土）21:00-22:00 にサンプルの予定が追加されます。

■ 動作確認

Excel 2007 on Windows Vista

■ 文字コード

GoogleCalendar.cls Shift_JIS
Module1.bas Shift_JIS
entry.xml UTF-8

■ 参考にしたサイト

Google Calendar APIの基礎
http://www.rcdtokyo.com/ucb/contents/i000815.php

Google Calendar Data API Developer's Guide: Protocolを日本語訳しました
http://d.hatena.ne.jp/shingotada/20070516

WinHTTP ライブラリで Web スクレイピング(1)～ GET 編～
http://www.f3.dion.ne.jp/~element/msaccess/AcTipsWinHTTP1.html

WinHttpRequest Object
http://msdn.microsoft.com/en-us/library/aa384106(v=VS.85).aspx