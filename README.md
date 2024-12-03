# NET不動産登記情報取得pythonスクリプト

[登記情報提供サービス](https://www1.touki.or.jp/gateway.html) の「不動産請求情報」ページをスクレーピングするためのPythonスクリプトです。スクレーピングに「selenium」、結果ファイルのExcel出力に「openpyxl」を利用しています。

## 目次

- [exe化対応](#exe化対応)
- [インストール](#インストール)
- [入力画面実行](#入力画面実行)
- [実行](#実行)
- [ライセンス](#ライセンス)

## exe化対応

前提：ローカルPC(Win10)のみ対応

1. **exe作成**

   ```cmd
   pyinstaller main.py --onefile --clean
   cd dist
   rename main.exe touki-scapy.exe
   
   ```

2. **その他の更新内容**

   - dist/Condition.xlsx を追加
   - dist/outputにtemplate.xlsx を追加
   - requirements.txt から streamlit を削除

3. **インストール方法**

   コンソール画面を開いて、以下のコマンドを投入してみてください。

   ```cmd
   git clone https://github.com/katsu-yoshimu/touki-scrapy.git
   ```

   「touki-scrapy」の下に「dist」ディレクトリが作成されます。
   「dist」ディレクトリには以下の構成でファイルがあります。

   ```memo
   dist
    +- touki-scrapy.exe
    +- Condition.xlsx
    +- output
      +- template.xlsx
   ```

   参照： [github distディレクトリ構成](https://github.com/katsu-yoshimu/touki-scrapy/tree/main/dist)

4. **実行方法**

   「Condition.xlsx」を開いて「ID番号」、「パスワード」を記入して保存してください。
   その他の条件は適当なものに変更してください。

   その後「touki-scrapy.exe」をダブルクリックすると収集処理が実行できます。
   コンソール画面が自動的に開きます。
   その後はpythonスクリプトを実行したものと同様な動きになります。

   ※2024/11/29時点で強制ログインの別パターンは未対応です。

5. **補足:**

   - [登記情報提供サービス](https://www1.touki.or.jp/gateway.html) サイトへのアクセス集中を回避するため、アクセス間隔は約3秒（正確には「3秒±20%」）以上となるように待機時間を設けています。なお、「3秒」は「selenimuContorller.py」の「INTERVAL_TIME = 3」で設定しており変更は容易です。

   - 想定外エラーが発生した場合、処理中断します。

   - 想定外エラーの対処は、基本的に「スクリプトの再実行」をする運用をお願いします。
   この運用が困難な場合は、想定外エラーの原因調査の上、スクリプトでの対応の可否を判断し、ご相談の上で、スクリプト修正を実施します。

   - 想定外エラーの原因調査のため、想定外エラー発生時点の **「トレースバック情報」「HTMLソース」** をログ出力、**「HTMLブラウザ画面のスナップショット」** をoutputディレクトリに出力します。

   - 「地番・家屋番号」の検索時において、頻度が高く「再実行」を促されるページされ検索できないめ、検索の再実行します。ただし、5回連続して検索できない場合は、処理中断します。なお、「5回」は「toukiController.py」の「CHIBAN_RETRY_OUT_COUNT = 5」で設定しており変更は容易です。

   - 「ただいま混み合っております。」ページが表示されることがあります。該当ページに「戻るボタンで戻れない場合は、ログアウトしてください。」との記載があり、「戻る」ボタンをクリックしても戻れなかったため、現時点（2024/11/13）は想定外エラー扱いとしています。

   - 「Server disconected」ページが表示されることがあります。実行マシンが不安定な場合に発生します。Python再実行するほか回復方法がないため、「スクリプトの再実行」をする運用をお願いします。

   - スクリプト実行中のブラウザは手動操作しないでください。このブラウザを非表示にすることも可能です。「selenimuContorller.py」のコメントアウトしている「# options.add_argument('--headless')」を「options.add_argument('--headless')」と変更して有効化してください。

## インストール

***2024/11/29時点にて設定ファイルのExcel対応およびexe配布としたため、本章は対応終了しました。***

前提：ローカルPC(Win10)に **git、ptyhon3.12** がインストール済

1. **ローカルPCにリポジトリのクーロン作成:**

   ```cmd
   cd %適当なディレクトリ%
   git clone https://github.com/katsu-yoshimu/touki-scrapy.git
   ```

2. **ローカルPCに仮想完了作成と仮想環境アクティベート:**

   ```cmd
   python -m venv venv
   venv\Scripts\activate
   ```

3. **ローカルPCに必要なPythonパッケージをインストール:**

   ```cmd
   cd touki-scrapy
   pip install -r requirements.txt
   ```

## 入力画面実行

***2024/11/29時点にて設定ファイルのExcel対応およびexe配布としたため、本章は対応終了しました。***

1. **実行:**

   ```cmd
   streamlit run inputCondition.py
   ```

   ```cmd
   You can now view your Streamlit app in your browser.

   Local URL: http://localhost:8501
   Network URL: http://192.168.1.3:8501
   ```

   と表示された後、ブラウザに入力画面が表示されます。入力画面の収集条件を入力して「収集実行」をクリックしてください。

2. **実行結果:**

   「output」ディレクトリにExcelファイルを作成します。
   以下は2024/11/13 9:18～9:28に実行した結果の例。

   ```cmd
   dir output
    -a----        2024/11/13      9:20          59620 output_20241113_091828.xlsx
   ```

3. **入力画面停止:**

   1.で実行したコマンド入力画面で「ctrl＋C」を入力してください。

   ```cmd
   Stopping...
   ```

   と表示された後、コマンド入力状態となります。

## 実行

***2024/11/29時点にて設定ファイルのExcel対応およびexe配布としたため、本章は対応終了しました。***

1. **設定ファイル更新:**

   ```notepad.exe
   notepad.exe config.json
   ```

   ```config.json
    {
        "user_id" : "ユーザIDで書き替えてください",
        "password" : "パスワードで書き替えてください",
        "conditions_list" : [
            ["土地", "32", "松江市東奥谷町", "380", "389", ["全部事項"]],
            ["土地", "32", "松江市東奥谷町", "380", "389", ["土地所在図/地積測量図"]],
            ["建物", "32", "松江市東奥谷町", "380", "389", ["全部事項"]],
            ["建物", "32", "松江市東奥谷町", "380", "389", ["建物図面/各階平面図"]]
        ]
    }
   ```

    - 収集条件リスト=conditions_list は収集条件を複数指定可能です。
    - 収集条件の各項目の設定内容は以下の通りです。
        - 収集条件[0]：種別："土地" or "建物"
        - 収集条件[1]：都道府県番号："01"：北海道 ～ "47"：沖縄
        - 収集条件[2]：市町村名：例）"松江市東奥谷町"
        - 収集条件[3]：地番・家屋番号の検索From：例）"380"
        - 収集条件[4]：地番・家屋番号の検索To：例）"389"
        - 収集条件[5]：請求種別："全部事項" or "土地所在図/地積測量図" or "建物図面/各階平面図"

2. **スクリプト実行:**

   ```cmd
   python main.py > log_yyyymmdd_hhmm.txt 2>&1
   ```

   「log_yyyymmdd_hhmm.txt」を「log_%date:~0,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%.txt」と書き替えると現在時刻のファイル名になります。

3. **実行結果:**

   condig.json の conditions_list のリスト数のExcelファイルを作成します。
以下は2024/11/13 9:18～9:28に実行した結果の例。

   ```cmd
   dir output
    -a----        2024/11/13      9:20          59620 output_20241113_091828.xlsx
    -a----        2024/11/13      9:27          65568 output_20241113_092015.xlsx
    -a----        2024/11/13      9:28          58876 output_20241113_092715.xlsx
    -a----        2024/11/13      9:32          60729 output_20241113_092839.xlsx
   ```

## ライセンス

ライセンスは Apache2 License に準拠します。
