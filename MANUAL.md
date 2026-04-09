# Hint Book Formatter 利用マニュアル

このマニュアルは、Hint Book Formatter を使ってスプレッドシートの内容を印刷用レイアウトに変換し、PDF として保存するための手順をまとめたものです。

## 1. このツールでできること

- Google Sheets のデータを読み込む
- Apps Script 経由で書式付きデータを読み込む
- 印刷用の見開きレイアウトを確認する
- 本にした時の並び順プレビューを確認する
- ブラウザの印刷機能から PDF を出力する

## 2. 画面の見方

画面は大きく 2 つのエリアに分かれています。

- 左側: 入力と操作
- 右側: プレビュー

左側には次の入力欄とボタンがあります。

- `Google Sheets URL or ID`
- `Load spreadsheet`
- `Use sample CSV`
- `Download JSON`
- `Export PDF`
- `Apps Script Web App URL`
- `Load Apps Script`
- `Use Apps Script sample`

## 3. 使い方の基本

### 3-1. まずはどちらで読み込むか決める

データの読み込み方法は 2 通りあります。

- 手軽に使う: `Google Sheets URL or ID`
- 書式や SIDE 定義も使う: `Apps Script Web App URL`

通常は、見た目の再現度を上げたい場合は `Apps Script Web App URL` を使ってください。

### 3-2. Google Sheets から読み込む

1. `Google Sheets URL or ID` に対象のスプレッドシート URL または ID を貼る
2. `Load spreadsheet` を押す
3. 右側のプレビューが更新される

注意:

- シートがブラウザから読める状態である必要があります
- この方法では、Apps Script 経由の共有設定や SIDE 定義は使えません

### 3-3. Apps Script から読み込む

1. `Apps Script Web App URL` に Web アプリの `/exec` URL を貼る
2. `Load Apps Script` を押す
3. 右側のプレビューが更新される

この方法では、Apps Script 側で返している以下の情報も使えます。

- `rows`
- `settings`
- `sideDefinitions`

### 3-4. プレビューの切り替え

右側のプレビュー上部には 2 つの切り替えボタンがあります。

- `印刷用`: 印刷シートとしての見え方
- `本にした`: 冊子順の見え方

用途に応じて切り替えて確認してください。

## 4. PDF の作り方

1. データを読み込む
2. 右側プレビューで内容を確認する
3. `Export PDF` を押す
4. ブラウザの印刷ダイアログで `Save as PDF` を選ぶ
5. 保存する

注意:

- PDF はアプリが直接保存するのではなく、ブラウザの印刷機能を使います
- 画像が多い場合は、読み込みに少し時間がかかることがあります

## 5. スプレッドシートの基本列

最低限、次の列を使います。

- `order`
- `page_no`
- `step`
- `side`
- `body`
- `image`

主な意味:

- `order`: 実際の並び順
- `page_no`: 表示するページ番号
- `step`: 上部の見出し
- `side`: SIDE 表示用の値または ID
- `body`: 本文
- `image`: 画像 URL または Google Drive 共有リンク

追加で使える列:

- `image_2`, `image_3`, `image_4`
- `image_position`
- `image_width`
- `image_height`
- `image_align`
- `image_fit`

## 6. BODY 内の画像記法

`body` の中に次のような記法を書くと、本文の途中に画像を差し込めます。

- `{{image}}`
- `{{image:2}}`
- `{{images:1,2,3}}`

例:

```text
説明文
{{image}}
続きの説明文
```

複数画像の例:

```text
{{images:1,2 width=160px align=center}}
```

## 7. Apps Script を使う場合の準備

### 7-1. Web アプリ URL の確認

使う URL は Apps Script の `/exec` URL です。

注意:

- `/dev` ではなく `/exec` を使う
- コードを更新したら再デプロイする
- 古い URL を貼らない

### 7-2. Apps Script の公開設定

Apps Script の Web アプリでは、少なくともブラウザからアクセスできる設定が必要です。

推奨例:

- Execute as: `Me`
- Who has access: `Anyone` または `Anyone in your organization`

### 7-3. side シートについて

SIDE を ID で管理する場合は、Apps Script 側で `side` シートを読みます。

最低限あるとよい列:

- `id`
- `text`

任意で使える列:

- `height`
- `background_color`
- `text_color`
- `font_family`
- `font_size`

### 7-4. settings シートについて

共有フォントや倍率を使う場合は `settings` シートを作成します。

例:

```text
step_font_family,MS Mincho
body_font_family,Yu Gothic
side_font_family,MS Mincho
page_no_font_family,Arial
step_font_scale,4
body_font_scale,2.5
```

## 8. よくあるつまずき

### 8-1. ページが真っ白になる

確認ポイント:

- GitHub Pages の公開元が正しいか
- 最新のデプロイが反映されているか
- URL が古くないか

### 8-2. Apps Script が効かない

確認ポイント:

- URL が `/exec` か
- Web アプリを再デプロイしたか
- ブラウザでその URL を直接開くと JSON が見えるか
- `Load Apps Script` 後の `Status` にエラーが出ていないか

### 8-3. SIDE 表示が数字のままになる

確認ポイント:

- `side` シートがあるか
- `id` と `text` 列があるか
- Apps Script が `sideDefinitions` を返しているか

### 8-4. 日本語が文字化けする

確認ポイント:

- スプレッドシート側の内容が正しく入力されているか
- CSV ではなく Apps Script の JSON を見ているか
- Google Sheets 側で意図しない変換が起きていないか

## 9. 運用のおすすめ

- 日常利用は `Apps Script Web App URL` を使う
- レイアウト確認後に `Export PDF` で出力する
- Apps Script を更新した時は、必ず再デプロイして最新の `/exec` URL を確認する
- シート構成を変える時は、`side` と `settings` の影響もあわせて確認する

## 10. 問題が起きた時に共有してほしい情報

不具合連絡の時は、次の情報があると切り分けが早くなります。

- 使用した URL の種類
  - Google Sheets
  - Apps Script
- 画面下の `Status` メッセージ
- ブラウザの Network タブで見えたステータスコード
- どのページがどう崩れているかのスクリーンショット

## 11. 管理者向けメモ

Apps Script の設定手順は [apps-script/README.md](C:/Users/shira/hintbookmaker/apps-script/README.md) を参照してください。
