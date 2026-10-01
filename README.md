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

## 出力

| シート | 内容 |
| --- | --- |
| 明細 | 1行 = 1品目。日付 / 店名 / 品目 / 数量 / 単価 / 金額 / ファイル |
| レシート一覧 | 1行 = 1枚。日付 / 店名 / 品目数 / 品目合計 / レシート合計 / 差額 / ファイル |

「差額」が 0 でない行は、読み取り漏れや税の扱いの差がある可能性があるので元画像を確認してください。
大きな画像は送信前に自動で縮小します。
