# GameTweetWatcher

**NOTE: This script was created for internal investigation project only, and so it won't make sense for anyone who don't relate to the project.**

**注意: このスクリプトは、社内の調査プロジェクトのためだけに作られたものなので、プロジェクトに関係のない人には機能しません。**

# 使用方法

Developer Account の取得、アプリの登録、OAUth1 ライブラリの組み込み等の詳細な手順は Qiita の下記の記事が参考になるかと思います。

これらの記事、特に 1 つ目と 2 つ目の記事をお読みいただき、全体の流れを掴んだ後、手順にしたがって設定していただくとスムーズです。

* [Twitter API と GAS を用いた情報収集 - 1. Twitter Developer Account 編](https://qiita.com/anti-digital/items/5f085d8d7361785f7def)
* [Twitter API と GAS を用いた情報収集 - 2. GAS+OAuth 編](https://qiita.com/anti-digital/items/acbd70b3ecedc6ff0b38)
* [Twitter API と GAS を用いた情報収集 - 3. Google スプレッドシートとの連携 編](https://qiita.com/anti-digital/items/f7d6de42974066ad1f25)

## 1. Google の My Drive の任意の場所に GAS を作成し、下記のコードを貼り付けます

https://github.com/anti-digital-tech/GameTweetWatcher/blob/master/src.gs/Code.gs

プロジェクト自体は [clasp](https://github.com/google/clasp) で開発しています。
ブラウザの GAS エディタを使わず clasp + TypeScript を使う場合には、普通に git clone してお使いください。
その場合 src.gs フォルダ以下は不要です。

## 2. 上記スクリプトの [OAuth1 ライブラリ](https://github.com/googleworkspace/apps-script-oauth1) を追加します

具体的な手順は、次のようになります。

2.(a) スクリプトエディタ上で、左側にある "Libraries" の横にある "+" を押します

2.(b) "Add a Library" ダイアログが表示されるので、そこの Script ID に 次の ID を入力します

`1CXDCY5sqT9ph64fFwSzVtXnbjpSfWdRymafDrtIZ7Z_hwysTY7IIhi7s`

2.(c) "Look up" します

2.(d) "Add" ボタンを押して追加します

<img alt="OAuth1 ライブラリの追加" width=400 src="https://camo.qiitausercontent.com/e05cf9534b3a3e8eda2a60e64f979c3990d7fd36/68747470733a2f2f71696974612d696d6167652d73746f72652e73332e61702d6e6f727468656173742d312e616d617a6f6e6177732e636f6d2f302f313835393935362f31656635636132642d663165612d643066652d366333342d3637343964633233323064312e706e67">

より詳細な手順は、Qiita の記事 : [Twitter API と GAS を用いた情報収集 - 2. GAS+OAuth 編](https://qiita.com/anti-digital/items/acbd70b3ecedc6ff0b38) または、[OAuth1 ライブラリ](https://github.com/googleworkspace/apps-script-oauth1) の "Set up" を参照してください。

## 3. Twitter Developer Account を取得し、アプリを登録し、App Key と Secret 値を入手します

[Twitter Developer Platform](https://developer.twitter.com/en/apply-for-access) から Twitter Developer Account を取得した後、
[Twitter Developer Portal](https://developer.twitter.com/en/portal) でアプリを登録し、App Key と Secret 値を入手します。

<img alt="Twitter Developer Portal でのアプリの登録とKeyの生成" width=400 src="https://camo.qiitausercontent.com/487cc2c3bb7f8629542b9a4d563f6debb204ba67/68747470733a2f2f71696974612d696d6167652d73746f72652e73332e61702d6e6f727468656173742d312e616d617a6f6e6177732e636f6d2f302f313835393935362f62343239393261622d396439632d393637642d303864332d6539646236343934383930372e706e67">

## 4. 1. のスクリプトのコールバックを 3. で登録したアプリの Callback URLs として登録します

スクリプトのコールバックは、スクリプトの ID から、次のように決まります。

`https://script.google.com/macros/d/{Script ID}/usercallback`

## 5. 下記 Google スプレッドシートをコピーして、Google の My Drive の任意の場所に置きます

https://docs.google.com/spreadsheets/d/1xiovn8szDPkuN6_QCCQQ0tACCwUMLMFwfZM_Y9cka2E/edit?usp=sharing

## 6. Google の My Drive の任意の場所に、メディア保存用、バックアップ用、履歴保存用のフォルダを 3 つそれぞれ作ります

これら、それぞれの ID をスクリプトに指定するので、控えておきます。

## 7. 1. のスクリプトに、各 Key、ID を記述します

```JavaScript
// ID of the Target Google Spreadsheet (Book)
const VAL_ID_TARGET_BOOK              = '{5. の Google スプレッドシートの ID}';
// ID of the Google Drive where the images will be placed
const VAL_ID_GDRIVE_FOLDER_MEDIA      = '{6. のメディア保存用の Google ドライブ上のフォルダの ID}';
// ID of the Google Drive where backup files will be placed
const VAL_ID_GDRIVE_FOLDER_BACKUP     = '{6. のバックアップ用の Google ドライブ上のフォルダの ID}';
// ID of the Google Drive where history files will be placed
const VAL_ID_GDRIVE_FOLDER_HISTORY    = '{6. の履歴保存用の Google ドライブ上のフォルダの ID}';
// Key and Secret to access Twitter APIs
const VAL_CONSUMER_API_KEY            = '{3. の Twitter Developer Portal で取得した API Key}';
const VAL_CONSUMER_API_SECRET         = '{3. の Twitter Developer Portal で取得した API Secret Key}';
```

## 8. 1. のスクリプトの getOAuthURL() を実行します

セキュリティー関連の警告ダイアログが出ますので、許可します。

## 9. ログに吐き出された URL をブラウザで開き、Twitter との連携を許可します

以上で Twitter API を呼ぶ準備が整います。

## 10. main() 関数を定期的に実行するようにスケジュールする

**main()** 関数を呼ぶことにより、5. のスプレッドシートに、キーワードで指定したツイートの検索結果が書き留められていきます。

また、**backup()** 関数はバックアップ処理を行う関数なので、これもスケジュールして、週いち程度で呼ぶように設定します。
