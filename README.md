# Google Apps Script(GAS)で日本株価(ETF 含む)、投資信託価格の取得

Google スプレッドシートを活用し、米株(GOOGLEFINACE で株価を取得できます)、日本株、投資信託の価格(GAS で作成した関数`STOCKPRICEJP`で取得します、株価は２０分ディレイです)を取得します。

特に、日本株と投資信託は Parser ライブラリを使用します。
使用にあたっては Parser ライブラリの追加を行う必要があります。(以下に記述してありますのでその手順にしたがって行ってください。)

## 作成の仕方

- Google のアカウントで chrome でログインし空のスプレッドシートを作成する

## Parser を追加

- スプレッドシートの`拡張機能`から`Apps Script`をクリックします。(Google はこのアプリを検証していません。というエラーメッセージが出ますが、`高度な`をクリックし`プロジェクトへ(危険)`のリンクをクリックしてください。そして、アクセスの`許可`をします。)
- ライブラリに`Parser`を追加します。
  スクリプト ID : 1Mc8BthYthXx6CoIz90-JiSzSafVnT6U3t0z_W3hLTAX5ek4w0G_EIrNw

  <img width="516" alt="スクリーンショット 2024-02-13 17 09 45" src="https://github.com/nuovotaka/NISA_GAS-2024/assets/11598404/b3a4985d-424c-4f7e-92e3-f4cb2ca72960">

- スクリプト ID を追加して検索をします。

  <img width="517" alt="スクリーンショット 2024-02-13 17 10 01" src="https://github.com/nuovotaka/NISA_GAS-2024/assets/11598404/b7d59582-9612-450a-be23-660d93d2316f">

- スクリプトの内容が表示されたら`追加`をクリック

## コードを追加

- `コード.gs`にプログラムを追加

  <img width="572" alt="スクリーンショット 2024-02-13 17 09 00" src="https://github.com/nuovotaka/NISA_GAS-2024/assets/11598404/0a94be08-c589-45e3-a409-aae36099be87">

- 今あるコードを全て削除します
- getStockPrice-gas.gs のコードをコピー＆ペーストする

### スプレッドシートタブにてスプレッドシートの作成

シートを１つ追加します。
次にシートの名前を変更してください。

- `シート1` -> `株価`
- `シート2` -> `表`

#### 株価シートを作成

1 行目に A 列から順に以下を記入

- 銘柄名
- 取引所コード
- 証券コード
- 株価(円)
- 株価(＄)
- 保有数(株、口数)
- 購入価格(1 単元)
- コストベース
- 時価評価額
- 備考

証券コード、株価(円)、時価評価額はプログラム上で各銘柄のセルの位置から値を取得していますので日本株及び投資信託のセルの位置はプログラムのコードと合わせてください。

<img width="1147" alt="スクリーンショット 2024-02-14 21 04 24" src="https://github.com/nuovotaka/NISA_GAS-2024/assets/11598404/e553388d-eca7-4c07-a7b8-2d8d29ecd316">

#### 銘柄名や証券コード、購入価格(1 単元)を入力する

米国株の場合は、ティッカーシンボルというのが決まっているのでそれを入れる
グーグルやアップルは、`GOOG`,`AAPL`などです

#### 数式を入力する

1. 米国株がある場合は株価(＄)を取得する式を入力する
   米国株の取得はスプレッドシートの株価(＄)の各セルで`=GOOGLEFINANCE(証券コードのセル位置)`を入力します。

（例）`=GOOGLEFINANCE(C2)`

C2 の位置が米国の証券コードになっていないとエラーとなります

取引所コードを指定する場合は
（例）`=GOOGLEFINANCE(B2&":"&C2)`

B2 に値を入れない場合でもエラーにはなりませんが google が推測した最適解を提供してくれます。

2. 日本株、等身の価格(円)を取得する式を入力する
   日本株(ETF 含む)と投信の株価(円)の各セルで`=STOCKPRICEJP(取引所コードのセルの位置,証券コードのセルの位置)`を入力します。

   (例) `=STOCKPRICEJP(B3,C3)`

   B3 のセルの値には`JP`or`TOSHIN`を入力してください。
   C3 が証券コードになります。
   取引所コードのセルの値が`JP`の場合は、日本株の証券コードが必要になります。
   投資信託の証券コードは ISIN コードになります。

3. 株価(円)
   米国株の円ベースの株価を取得するには先ずドル円を取得する必要があります。
   株価一覧の下にドル円取得のためのコードを入力する

```
=GOOGLEFINANCE("USDJPY")
```

4. コストベース
   購入価格(1 単元)\*保有数
   購入価格(1 単元)は円ベースの方が後々計算しやすいので円ベースで入れる
   (例)

```
=G2*F2
```

5. 時価評価額
   株価(円)\*保有数
   (例)

```
=D2*F2
```

#### 表シートを作成

1 行目は最初に更新日時が来るのでそのあとは銘柄名を横並びに記載してください。

- 更新日時
- 1 つ目の銘柄名
- 2 つ目の銘柄名
- 3 つ目の銘柄名
- 4 つ目の銘柄名
- 5 つ目の銘柄名
- 6 つ目の銘柄名

今回はサンプルで以下のようになっています

- Google(GOOG)
- バンガード・S&P 500(VOO)
- S&P 500 ETF(1655)
- ソフトバンク(9984)
- eMAXIS Slim S&P500
- eMAXIS Slim 全世界株式

<img width="883" alt="スクリーンショット 2024-02-13 20 10 41" src="https://github.com/nuovotaka/NISA_GAS-2024/assets/11598404/0bdfd8cb-f314-4514-b77b-624f250f4066">

### Apps Script タブにて

#### 日本株の銘柄数や投資信託の数によりコードを一部変更する必要あり

- `updateStockPriceList`のコードを変更します(KEY,VALUE の形式となっていますコード上で`eDataCell`と`eColumn`の株価を取得する`KEY`は同じにしてください。)

#### 表作成のタイマーを設定

時計マークの`トリガー`にて行います。

- トリガーを追加
- 実行する関数を選択にて`updateStockPriceList`を選択
- イベントのソースで選択にて`時間主動型`を選択
- 時間ベースのトリガーのタイプを選択にて`日付ベースのタイマー`を選択
- 時刻を選択にて`午前０時〜１時`を選択
- 保存

時刻選択はなるべく一日に１回程度に留めておかないと運用する上での上限値がありますのでそれに抵触する恐れがあります。
また、日本株取得時にサーバーへの負荷となりますのでご注意ください。

## 参考

後は、コストベースと時価評価額の値を使ってグラフなど作成されてはいかがでしょうか？
それは各自行ってください。
尚、appsscript.json は`Parser`を追加すると作られるものなので気にしなくても大丈夫です。
