◆目的
たくさんSSIDの異なるWi-Fiルーター(AP)があって、動作確認するためのプログラム

◆実行内容
targetlist.csvを読み込み、動作確認したいSSIDとPWを取得する。
PCの電波受信状況を確認し、電波が強いものから接続をする。
PINGの送信、http接続を行う。
状況をlog.txtに書き込み、結果をdata.csvに書き込む
二回目以降は、targetlist.csvではなく、data.csvから読み込む。

◆実行方法
startup.batをダブルクリックして実行する。

◆todo
http接続で受け取る側のサーバーでログを受け取って、スマホで確認できるように。
