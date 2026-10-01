"""レシート画像を読み取り、日付・店名・購入品名・金額をExcelにまとめるツール.

1レシート=1行で、購入品名は「雑貨(トイレットペーパー)」「タクシー代」のような
シンプルな代表名にまとめる。

使い方:
    python receipt_to_xlsx.py receipts/ -o receipts.xlsx
    python receipt_to_xlsx.py img1.jpg img2.png -o out.xlsx
    python receipt_to_xlsx.py receipts/ --names names.csv   # 区分リストを差し替える
    python receipt_to_xlsx.py --from-json data/receipts.json  # 読み取り済みデータからExcelだけ作る
"""

from __future__ import annotations

import argparse
import base64
import csv
import io
import json
import sys
from collections import defaultdict
from pathlib import Path

import anthropic
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image, ImageOps

MODEL = "claude-opus-5-5"
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".webp"}
MAX_IMAGE_BYTES = 4_500_000  # API上限 5MB に余裕を持たせる
MAX_LONG_EDGE = 2400

# 参考スプレッドシート「レシートデータのエクセル・CSV化」で使われている区分
DEFAULT_CATEGORIES = [
    "タクシー代", "交通費", "交通費チャージ", "飲食代", "飲料", "食品", "雑貨",
    "事務用品", "PC用品", "工具", "医薬品", "書籍", "花代", "園芸用品", "送料",
    "宅急便運賃", "通信費", "ユニフォームレンタル等", "メンテナンス料金", "清掃代",
    "ゴミ処理券", "ゴミ処理手数料", "会費", "部会費", "行政証明書発行手数料",
]
HEADER_WORDS = {"区分", "品名", "品目", "代表品名", "購入品名", "商品名"}

RECEIPT_SCHEMA = {
    "type": "object",
    "properties": {
        "date": {"type": "string", "description": "購入日 (YYYY-MM-DD)。読み取れない場合は空文字"},
        "store_name": {"type": "string", "description": "店名。読み取れない場合は空文字"},
        "purchase_name": {"type": "string", "description": "シンプルな購入品名"},
        "total": {"type": "number", "description": "合計金額 (円、税込)"},
        "items": {
            "type": "array",
            "description": "レシートに載っている商品 (確認用)",
            "items": {
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "レシート上の表記"},
                    "amount": {"type": "number", "description": "金額 (円、値引き後)"},
                },
                "required": ["name", "amount"],
                "additionalProperties": False,
            },
        },
    },
    "required": ["date", "store_name", "purchase_name", "total", "items"],
    "additionalProperties": False,
}

# スキャン1枚に複数のレシートが貼られていることがあるので、配列で受け取る
PAGE_SCHEMA = {
    "type": "object",
    "properties": {"receipts": {"type": "array", "items": RECEIPT_SCHEMA}},
    "required": ["receipts"],
    "additionalProperties": False,
}

PROMPT_TEMPLATE = """この画像に写っているレシート・領収書をすべて読み取り、1枚ごとに
購入日・店名・購入品名・合計金額を抽出してください。

- 1枚の画像に複数のレシートが貼られていることがある。写っているものはすべて receipts に入れる。
- 画像が上下逆さま・横向きの場合もあるので、向きを補正して読む。
- 同じ取引のレシートと領収書が両方ある場合は1件にまとめる(例: ダイソーの領収証と付属のレシート)。

- date: 西暦 YYYY-MM-DD 形式に変換する（和暦や「25/10/01」のような表記も変換）。
- store_name: 店名。チェーン店は支店名まで（例: 「セブン-イレブン 上野桜木2丁目店」）。
  タクシーは会社名（例: 「日本交通」「国際自動車 (km)」）。
- total: 税込の支払合計。
- items: 商品行をレシートの表記どおりに書き起こす。小計・合計・税・お預り・お釣り・
  ポイントの行は含めず、値引きは直前の商品の金額に反映する。
- purchase_name: レシート全体を、次の区分リストのいずれかを使ってシンプルに表す。
  {categories}
  - 中身を示すと分かりやすい区分は「区分（主な品目）」とし、品目は1〜2個の一般名詞で短く書く。
    ブランド名・容量・型番は書かない。品目が多いときは「等」を付ける。
  - タクシー代・飲食代・花代など、区分だけで分かるものは区分のみにする。
  - どれにも当てはまらない場合は内容を表す短い名前を付け、判別できなければ「不明」とする。
  例: 「タクシー代」「飲食代」「雑貨（トイレットペーパー）」「雑貨（洗剤・ゴミ袋）」
      「事務用品（輪ゴム）」「事務用品（ファイル等）」「飲料（Y1000）」「食品（ケーキ類）」
      「PC用品（SDカード）」「書籍（雑誌代）」「通信費（切手代）」「交通費チャージ」
  括弧は全角「（）」を使う。"""


def load_categories(path: Path) -> list[str]:
    """区分リストを読み込む (CSV/TXT の1列目、空行・重複・見出しは除く)."""
    names: list[str] = []
    with path.open(encoding="utf-8-sig", newline="") as f:
        for row in csv.reader(f):
            name = row[0].strip() if row else ""
            if name and name not in names and name not in HEADER_WORDS:
                names.append(name)
    return names


def build_prompt(categories: list[str]) -> str:
    return PROMPT_TEMPLATE.format(categories="、".join(categories))


def load_image(path: Path) -> tuple[str, str]:
    """画像を(必要なら縮小して)base64化し、(media_type, data) を返す."""
    raw = path.read_bytes()
    media_type = {
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".png": "image/png",
        ".gif": "image/gif",
        ".webp": "image/webp",
    }[path.suffix.lower()]

    img = Image.open(io.BytesIO(raw))
    if len(raw) > MAX_IMAGE_BYTES or max(img.size) > MAX_LONG_EDGE:
        img = ImageOps.exif_transpose(img).convert("RGB")
        img.thumbnail((MAX_LONG_EDGE, MAX_LONG_EDGE))
        quality = 90
        while True:
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=quality)
            if buf.tell() <= MAX_IMAGE_BYTES or quality <= 50:
                break
            quality -= 10
        raw, media_type = buf.getvalue(), "image/jpeg"

    return media_type, base64.standard_b64encode(raw).decode("utf-8")


def extract_receipts(client: anthropic.Anthropic, path: Path, prompt: str) -> list[dict]:
    media_type, data = load_image(path)
    response = client.beta.messages.create(
        model=MODEL,
        max_tokens=16000,
        betas=["server-side-fallback-2026-07-01"],
        fallbacks="default",
        output_config={
            "effort": "medium",
            "format": {"type": "json_schema", "schema": PAGE_SCHEMA},
        },
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": data}},
                {"type": "text", "text": prompt},
            ],
        }],
    )
    if response.stop_reason == "refusal":
        raise RuntimeError("モデルが応答を拒否しました")
    if response.stop_reason == "max_tokens":
        raise RuntimeError("出力が途中で切れました")
    text = next(b.text for b in response.content if b.type == "text")
    return json.loads(text)["receipts"]


def collect_images(inputs: list[str]) -> list[Path]:
    paths: list[Path] = []
    for arg in inputs:
        p = Path(arg)
        if p.is_dir():
            paths.extend(sorted(f for f in p.iterdir() if f.suffix.lower() in IMAGE_EXTS))
        elif p.suffix.lower() in IMAGE_EXTS:
            paths.append(p)
        else:
            print(f"スキップ (対応していない形式): {p}", file=sys.stderr)
    return paths


def style_sheet(ws, widths: list[int], money_cols: list[int]) -> None:
    for i, width in enumerate(widths, start=1):
        cell = ws.cell(row=1, column=i)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(i)].width = width
    for col in money_cols:
        for (cell,) in ws.iter_rows(min_row=2, min_col=col, max_col=col):
            cell.number_format = "#,##0"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def slash_date(date: str) -> str:
    return date.replace("-", "/")


def write_receipt_sheet(ws, rows: list[tuple[Path, dict]]) -> None:
    ws.append(["日付", "店名", "購入品名", "金額"])
    for _, r in rows:
        ws.append([slash_date(r["date"]), r["store_name"], r["purchase_name"], r["total"]])
    style_sheet(ws, [12, 30, 30, 10], money_cols=[4])


def write_workbook(results: list[tuple[Path, dict]], out: Path) -> None:
    results = sorted(results, key=lambda r: (r[1]["date"] or "9999", r[1]["store_name"]))
    wb = Workbook()

    write_receipt_sheet(wb.active, results)
    wb.active.title = "Master"

    # 月別シート (年が1つだけなら「1月」、複数あれば「2025年1月」)
    by_month: dict[str, list[tuple[Path, dict]]] = defaultdict(list)
    for path, r in results:
        by_month[r["date"][:7] if len(r["date"]) >= 7 else ""].append((path, r))
    years = {ym[:4] for ym in by_month if ym}
    for ym, rows in by_month.items():
        if not ym:
            title = "日付不明"
        elif len(years) == 1:
            title = f"{int(ym[5:7])}月"
        else:
            title = f"{ym[:4]}年{int(ym[5:7])}月"
        write_receipt_sheet(wb.create_sheet(title), rows)

    # 確認用: 読み取った商品行と、商品合計とレシート合計の差額
    detail = wb.create_sheet("明細(確認用)")
    detail.append(["日付", "店名", "購入品名", "レシート表記", "金額", "ファイル"])
    for path, r in results:
        for item in r["items"]:
            detail.append([slash_date(r["date"]), r["store_name"], r["purchase_name"],
                           item["name"], item["amount"], path.name])
        diff = r["total"] - sum(item["amount"] for item in r["items"])
        if r["items"] and diff:
            detail.append([slash_date(r["date"]), r["store_name"], r["purchase_name"],
                           "(合計との差額: 税・端数など)", diff, path.name])
    style_sheet(detail, [12, 30, 30, 36, 10, 24], money_cols=[5])

    wb.save(out)


def load_json(path: Path) -> list[tuple[Path, dict]]:
    """読み取り済みデータ (レシートの配列、各要素に file を含む) を読み込む."""
    records = json.loads(path.read_text(encoding="utf-8"))
    return [(Path(r.get("file", "")), r) for r in records]


def main() -> int:
    parser = argparse.ArgumentParser(description="レシート画像をExcelにまとめます")
    parser.add_argument("inputs", nargs="*", help="画像ファイルまたはフォルダ")
    parser.add_argument("-o", "--output", default="receipts.xlsx", help="出力ファイル名")
    parser.add_argument("--names", type=Path,
                        help="購入品名の区分リスト (CSV/TXT、1列目を使用)。省略時は既定の区分")
    parser.add_argument("--from-json", type=Path,
                        help="API を使わず、読み取り済みデータ (JSON) から Excel を作る")
    args = parser.parse_args()

    if args.from_json:
        results = load_json(args.from_json)
        write_workbook(results, Path(args.output))
        print(f"{len(results)} 件のレシートを {args.output} に保存しました")
        return 0

    categories = load_categories(args.names) if args.names else DEFAULT_CATEGORIES
    prompt = build_prompt(categories)

    images = collect_images(args.inputs)
    if not images:
        print("画像が見つかりませんでした", file=sys.stderr)
        return 1

    client = anthropic.Anthropic()
    results: list[tuple[Path, dict]] = []
    for path in images:
        print(f"読み取り中: {path}")
        try:
            receipts = extract_receipts(client, path, prompt)
            print(f"  {len(receipts)} 件")
            results.extend((path, r) for r in receipts)
        except (anthropic.APIConnectionError, anthropic.APIStatusError, RuntimeError) as e:
            print(f"  失敗: {e}", file=sys.stderr)

    if not results:
        return 1
    write_workbook(results, Path(args.output))
    print(f"{len(images)} 枚の画像から {len(results)} 件のレシートを {args.output} に保存しました")
    return 0


if __name__ == "__main__":
    sys.exit(main())
