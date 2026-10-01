# レシート → スプレッドシート

レシート画像を Claude で読み取り、**日付・店名・購入品目(数量・単価・金額)** を Excel (.xlsx) にまとめます。

## セットアップ

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY=sk-ant-...
```

## 使い方

```bash
# フォルダ内の画像 (jpg / png / gif / webp) をまとめて処理
python receipt_to_xlsx.py receipts/ -o receipts.xlsx

# ファイルを個別に指定
python receipt_to_xlsx.py IMG_0001.jpg IMG_0002.png
```

## 品名をシンプルにする

品名は「明治おいしい牛乳1000ml」→「牛乳」のように、ブランド・容量などを省いた代表品名にまとめます
(レシート上の表記は「レシート表記」列に残ります)。

使いたい代表品名のリストがある場合は、1列目に品名を並べた CSV / TXT を `--names` で渡すと、
すべての品名をそのリストのどれか(当てはまらなければ「その他」)に揃えます。
Google スプレッドシートからは「ファイル → ダウンロード → CSV」で書き出せます。

```bash
python receipt_to_xlsx.py receipts/ --names names.csv
```

## 出力

| シート | 内容 |
| --- | --- |
| 明細 | 1行 = 1品目。日付 / 店名 / 品名 / 数量 / 単価 / 金額 / レシート表記 / ファイル |
| レシート一覧 | 1行 = 1枚。日付 / 店名 / 品目数 / 品目合計 / レシート合計 / 差額 / ファイル |

「差額」が 0 でない行は、読み取り漏れや税の扱いの差がある可能性があるので元画像を確認してください。
大きな画像は送信前に自動で縮小します。
