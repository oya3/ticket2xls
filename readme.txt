機能：
redmine のチケットに記載した特定の文字列を抽出しxlsにする

使い方：
perl ticket2xls.pl "redmineサイトアドレス" "REST API キー" チケット番号 出力ファイル(xls)

仕様：
対象チケットの説明(description)に記述された特定の文字列のみ抽出する
xls ファイル出力時は、ms excel は必要ない。
perl がインストール済みであること。
ppm で spreadsheet 関連モジュールがインストール済みであること。

特定文字のフォーマット
・行の先頭から '@[xxx]@' で記述された xxx がキーとなる
・次のキーが登場するまでの文字列は全て検出したキーに持ち物となる。
・出力は各キー別にセルを分けて出力する。

動作確認：
windows 7(32,64)環境のみ

同梱内容：
ticket2xls.pl
ツール本体

get.bat
動作サンプルバッチファイル


下記、チケットの説明欄の記載方法と出力サンプル
-----------------------
説明
@[test]@
test

@[test2]@
テスト２

@[テスト３]@
test3
test3-1

出力サンプル
----------------------
 |A   |B       |C       |...
1|test|test2   |テスト３|...
2|test|テスト２|test3   |
 |    |        |test31  |...
