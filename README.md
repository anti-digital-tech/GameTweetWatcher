# GameTweetWatcher

**NOTE: This script was created for internal investigation project only, and so it won't make sense for anyone who don't relate to the project.**

**注意: このスクリプトは、社内の調査プロジェクトのためだけに作られたものなので、プロジェクトに関係のない人には機能しません。**

# 使用方法

Developer Account の取得、アプリの登録、OAUth1 ライブラリの組み込み等の詳細な手順は Qiita の下記の記事が参考になるかと思います。

* https://qiita.com/anti-digital/items/5f085d8d7361785f7def
* https://qiita.com/anti-digital/items/acbd70b3ecedc6ff0b38
* https://qiita.com/anti-digital/items/f7d6de42974066ad1f25

(1) Google の My Drive の任意の場所に GAS を作成し、下記のコードを貼り付けます

https://github.com/anti-digital-tech/GameTweetWatcher/blob/master/src.gs/Code.gs

([clasp](https://github.com/google/clasp) を使う場合には、[TypeScript 版](https://github.com/anti-digital-tech/GameTweetWatcher/blob/master/src/Code.ts) を使用するのが良いと思います)

(2) 上記スクリプトの [OAuth1 ライブラリ](https://github.com/googleworkspace/apps-script-oauth1) を追加します

(3) Twitter Developer Account を取得し、アプリを登録し、App Key と Secret 値を入手します

[Twitter Developer Platform](https://developer.twitter.com/en/apply-for-access) から Twitter Developer Account を取得した後、
[Twitter Developer Portal](https://developer.twitter.com/en/portal) でアプリを登録し、App Key と Secret 値を入手します。

(4) (1) のスクリプトのコールバックを (3) で登録したアプリの Callback URLs として登録します

スクリプトのコールバックは、スクリプトの ID から、次のように決まります。
`https://script.google.com/macros/d/{Script ID}/usercallback`

(5) 下記 Google スプレッドシートをコピーして、Google の My Drive の任意の場所に置きます

https://docs.google.com/spreadsheets/d/1xiovn8szDPkuN6_QCCQQ0tACCwUMLMFwfZM_Y9cka2E/edit?usp=sharing

(6) Google の My Drive の任意の場所に、メディア保存用、バックアップ用、履歴保存用のフォルダを 3 つそれぞれ作ります

これら、それぞれの ID を GAS で使用するので控えておきます。

(7) (1) のスクリプトに、各 Key, ID を記述します

```JavaScript
// ID of the Target Google Spreadsheet (Book)
const VAL_ID_TARGET_BOOK              = '{(5) の Google スプレッドシートの ID}';
// ID of the Google Drive where the images will be placed
const VAL_ID_GDRIVE_FOLDER_MEDIA      = '{(6) のメディア保存用の Google ドライブ上のフォルダの ID}';
// ID of the Google Drive where backup files will be placed
const VAL_ID_GDRIVE_FOLDER_BACKUP     = '{(6) のバックアップ用の Google ドライブ上のフォルダの ID}';
// ID of the Google Drive where history files will be placed
const VAL_ID_GDRIVE_FOLDER_HISTORY    = '{(6) の履歴保存用の Google ドライブ上のフォルダの ID}';
// Key and Secret to access Twitter APIs
const VAL_CONSUMER_API_KEY            = '{(3) の Twitter Developer Portal で取得した API Key}';
const VAL_CONSUMER_API_SECRET         = '{(3) の Twitter Developer Portal で取得した API Secret Key}';
```

(8) (1) のスクリプトの **getOAuthURL()** を実行します

セキュリティー関連の警告ダイアログが出ますので、許可します。

(9) ログに吐き出された URL をブラウザで開き、Twitter 連携させます

以上で Twitter API を呼ぶ準備ができます。

(10) **main()** 関数を定期的に実行するようにスケジュールする


