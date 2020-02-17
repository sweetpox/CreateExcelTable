# Dictionaryを使ってDataTableをExcelのように扱う

ExcelやCSVを読み込み、DataTableで使いたい値をインデックスで指定する際に、いつもExcelの画面を見ながら「Aが0、Bが1、Cが2...」と数えながらインデックスを確認していました。
その時に、Excelみたいに使いたい値を「H5」のようにセル位置で指定すれば値を取得できないかと考え、Dictionaryを使ってDataTableをExcelのように扱う方法を考えました。

# CreateExcelTableについて

DataTableをDictionaryに変換し、Excelのように **「A1」** や **「B2」** とセル位置を指定してデータを取り出すことが出来るようにしています。

### メリット

 - 取得したい値をExcelのセル位置で指定することができるのでインデックス値の計算が不要
 - `変数名（"セル位置"）`で値を取得できるため、入力が短くてすむ

### デメリット

 - 繰り返し処理を行うときに、cnt, iなどのカウンター用の変数を用意して`列 + cnt.ToString`と指定するため、いろいろとめんどくさい
 - DataTableをDictionaryに移しているので、データ量が膨大だと無駄に処理時間が増えてしまう

## CreateExcelTableの使い方

ダウンロードが完了したら、任意のプロジェクトフォルダ内に **「CreateExcelTable.xmal」** ファイルを配置してください。
配置が完了したら、UiPathでプロジェクトを開いてください。

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/556204/9b8b4ab3-f516-bfd0-c364-a52a9d0cb009.png)

### Main実行

**「CreateExcelTable.xmal」** をダブルクリックして、「デバッグ用」を選択し、[Ctrl + D]を押してコメントアウトしてください。

![commentout.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/556204/9b8f1cbe-1fbf-3a77-7401-007bae8f0ffd.png)

**「Main.xmal」** で、`ワークフローファイルを呼び出し`を配置してください。
配置が完了したら、[...]をクリックして、**「CreateExcelTable.xmal」** を選択して開くをクリックしてください。

![call_workflow.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/556204/12026bc5-f4ba-bc86-bdc7-3183be8e049e.png)

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/556204/0848eab6-6c68-e54a-5cfc-b2cf033ca4e5.png)

`ワークフローファイルを呼び出し`の「引数をインポート」をクリックしてください。

![argument_import.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/556204/d91a3baf-deef-84df-a2db-61479bf379d6.png)

値の欄で[Ctrl + K]を押して、任意の名前で変数を作成してください。
プロジェクトファイルで実際に利用をする場合は、短い変数名を使用すると参照するときに楽だと思います。

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/556204/83dc0dd4-5508-22bc-d78b-4f388a62758c.png)

`et("B3")`のように`変数名（"セル位置"）`で実際のデータを取得することができます。
※"b3"のように小文字で入力すると、エラーになってしまうので注意してください。

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/556204/154eacdf-c061-95ee-dc03-370909244dea.png)


### デバッグ実行

**「CreateExcelTable.xmal」** をダブルクリックし、実行を押してください。
「デバッグ用」をコメントアウトしている場合は、[Ctrl + E]を押して、コメントアウトを解除してください。

![main.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/556204/c93534bb-d805-ee84-ceb1-30d27d35bba4.png)

## CreateExcelTableの実行手順

### １．ファイル選択

実行すると以下のポップアップが表示されるので、OKをクリックします。

![process1.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/556204/5eeb5795-6a23-2760-f3ef-38d473eb546b.png)

ファイル選択画面で、ExcelファイルかCSVファイルを選択してください。
`※それ以外のファイルについては対応していません。`

![process2.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/556204/0be4f751-75b3-1695-86a6-bca23bb5e376.png)

### ２．ヘッダー列作成選択

ヘッダー列を作成するか確認するポップアップが表示されます。
作成する場合は **「はい」**、しない場合は **「いいえ」** を選択してください。

![process3.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/556204/bb58b599-3033-0770-8839-cc7334f1d653.png)

以下の画像のデータを読み込んだときの、ヘッダー列ありとなしの出力結果は下段に記載しているので、参照してください。

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/556204/abc2835c-c59a-17d9-8dd4-4d46f2a25481.png)


#### ヘッダー列あり

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/556204/5128a972-3446-fe6c-e8fb-c1ece210fb22.png)

#### ヘッダー列なし

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/556204/2a992939-4f77-ceb3-c326-2f754e07b7c5.png)

### ３．Excel選択時の設定

CSVファイルを選択した場合は、「２．ヘッダー列作成選択」で終了です。
Excelファイルを選択した場合は、以下のように **「シート名」** と **「範囲読み込み位置」** がデフォルト設定で良いか確認するポップアップが表示されます。

![process4.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/556204/78c20f05-2230-9ba6-f6c9-65469a16a4cb.png)

**「いいえ」** を選択した場合は以下のように入力ダイアログが表示されるので、**「シート名」** と **「範囲読み込み位置」** を入力してください。

![process5.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/556204/e744d769-13cd-a12a-6256-223306154aed.png)

![process6.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/556204/b94e4455-deaa-c0cb-d0f9-b8f5456cf62f.png)
