#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Jリーグチケットサイトから、J1全クラブの試合ごとの席種価格・完売状況を取得し
CSVに履歴として記録するスクリプト。

【重要な前提と制約】
- クラブの「試合一覧」ページ (https://www.jleague-ticket.jp/club/{code}/) は
  ボット対策(Cloudflare等)で機械的なアクセスが弾かれるケースが確認されています。
  そのため、まずは「対象試合のURLを直接指定する」運用を基本にしています。
  (例: https://www.jleague-ticket.jp/sales/perform/2626898/001)
- 個別試合ページ (/sales/perform/xxxxxxx/001) は通常のHTTPリクエストで取得できることを確認済みです。
- 「完売」表示の実際のHTML構造は、完売試合の実例で未確認のため、
  detect_status() 内のロジックは暫定です。完売中の試合URLが見つかったら
  そのページのHTMLを見て、SOLD_OUT_PATTERNS を調整してください。
"""

import csv
import os
import re
import time
from datetime import datetime, timezone, timedelta

import pandas as pd
import requests
from bs4 import BeautifulSoup

JST = timezone(timedelta(hours=9))

# ── J1全20クラブのクラブコード ──────────────────────────────
J1_CLUBS = {
    "ka": "鹿島アントラーズ",
    "mh": "水戸ホーリーホック",  # ※J2実際は水戸だが取得元リストのまま記載。必要に応じて削除可
    "ur": "浦和レッズ",
    "je": "ジェフ千葉",
    "kr": "柏レイソル",
    "to": "FC東京",
    "vn": "東京ヴェルディ",
    "mz": "FC町田ゼルビア",
    "kf": "川崎フロンターレ",
    "ym": "横浜F・マリノス",
    "ss": "清水エスパルス",
    "ng": "名古屋グランパス",
    "ks": "京都サンガF.C.",
    "go": "ガンバ大阪",
    "co": "セレッソ大阪",
    "vi": "ヴィッセル神戸",
    "fo": "ファジアーノ岡山",
    "sh": "サンフレッチェ広島",
    "af": "アビスパ福岡",
    "vv": "V・ファーレン長崎",
}

# ── 監視したい試合URLをここに追加していく運用 ───────────────
# 自動でクラブページから拾えない場合の暫定リスト。
# 例: 東京V対柏 (2026/08/14)
TARGET_MATCH_URLS = [
    "https://www.jleague-ticket.jp/sales/perform/2626898/001",
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
}

OUTPUT_CSV = "ticket_prices.csv"
OUTPUT_XLSX = "ticket_prices.xlsx"

SOLD_OUT_PATTERNS = ["完売", "SOLD OUT", "販売終了"]
PRICE_PATTERN = re.compile(r"([\d,]+)円\s*[~〜～]\s*([\d,]+)円\s*/\s*枚")


def fetch(url: str) -> str | None:
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        resp.raise_for_status()
        return resp.text
    except requests.RequestException as e:
        print(f"[WARN] fetch failed: {url} ({e})")
        return None


def extract_match_meta(html: str, url: str) -> dict:
    soup = BeautifulSoup(html, "html.parser")
    title = soup.title.string.strip() if soup.title and soup.title.string else ""
    # 例: "東京ヴェルディ対柏レイソル　明治安田Ｊ１リーグ(2026/08/14) | Ｊリーグチケット"
    date_match = re.search(r"\((\d{4}/\d{2}/\d{2})\)", title)
    match_date = date_match.group(1) if date_match else ""
    card = title.split("(")[0].split("|")[0].strip() if title else ""
    perform_id_match = re.search(r"/perform/(\d+)/", url)
    perform_id = perform_id_match.group(1) if perform_id_match else ""
    return {"perform_id": perform_id, "card": card, "match_date": match_date, "url": url}


def extract_seat_blocks(html: str) -> list[dict]:
    """
    ページ全体のテキストから席種ブロックを抽出する。
    各ブロックは概ね次のテキスト構造:
        <席種名見出し>
        <価格帯>円～<価格帯>円/枚
        発売 情報
        ...(発売期間などが続く)
        <席種名> 選択する   ← リンクテキスト。ブロックの終端目印
    """
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text("\n")
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    seats = []
    current_name = None
    for i, line in enumerate(lines):
        price_m = PRICE_PATTERN.search(line)
        if price_m and current_name:
            is_dynamic = False
            # 直前～直後数行に《変動》があれば変動価格対象
            window = lines[max(0, i - 2): i + 8]
            if any("変動" in w for w in window):
                is_dynamic = True
            status = "販売中"
            if any(p in w for w in window for p in SOLD_OUT_PATTERNS):
                status = "完売"
            seats.append({
                "seat_type": current_name,
                "price_min": price_m.group(1).replace(",", ""),
                "price_max": price_m.group(2).replace(",", ""),
                "dynamic": is_dynamic,
                "status": status,
            })
            current_name = None
        elif not price_m and 2 <= len(line) <= 30 and not line.startswith("■") \
                and "円" not in line and "選択する" not in line and "発売" not in line:
            # 席種名候補(短い行、価格や発売情報でないもの)
            current_name = line
    return seats


def check_match(url: str) -> list[dict]:
    html = fetch(url)
    if not html:
        return []
    meta = extract_match_meta(html, url)
    seats = extract_seat_blocks(html)
    now = datetime.now(JST).strftime("%Y-%m-%d %H:%M")
    rows = []
    for seat in seats:
        rows.append({
            "checked_at": now,
            "card": meta["card"],
            "match_date": meta["match_date"],
            "perform_id": meta["perform_id"],
            "url": meta["url"],
            **seat,
        })
    return rows


def append_csv(rows: list[dict]):
    if not rows:
        return
    file_exists = os.path.isfile(OUTPUT_CSV)
    with open(OUTPUT_CSV, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        if not file_exists:
            writer.writeheader()
        writer.writerows(rows)


def safe_sheet_name(name: str) -> str:
    """Excelのシート名で使えない文字を除去し31文字に収める"""
    cleaned = re.sub(r'[\\/*?:\[\]]', '_', str(name)).strip()
    return (cleaned or "sheet")[:31]


def build_price_display(df: pd.DataFrame) -> pd.DataFrame:
    def fmt(row):
        base = f"{int(row['price_min']):,}円〜{int(row['price_max']):,}円"
        if str(row.get("dynamic")).lower() in ("true", "1"):
            base += "[変動]"
        if str(row.get("status")) == "完売":
            base += "(完売)"
        return base

    df = df.copy()
    df["price_display"] = df.apply(fmt, axis=1)
    return df


def build_pivots(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """試合(card)ごとに 行=checked_at, 列=seat_type のピボット表を作る"""
    pivots = {}
    for card, group in df.groupby("card", dropna=False):
        pivot = group.pivot_table(
            index="checked_at",
            columns="seat_type",
            values="price_display",
            aggfunc="first",
        )
        pivots[safe_sheet_name(card)] = pivot
    return pivots


def export_pivot_xlsx(pivots: dict[str, pd.DataFrame]):
    if not pivots:
        return
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        for sheet_name, pivot in pivots.items():
            pivot.to_excel(writer, sheet_name=sheet_name)
    print(f"[INFO] {OUTPUT_XLSX} を出力しました")


def export_google_sheets(pivots: dict[str, pd.DataFrame]):
    """
    環境変数 GOOGLE_SHEET_ID と GOOGLE_CREDENTIALS_PATH が設定されている場合のみ、
    Googleスプレッドシートへ書き出す。未設定なら黙ってスキップする。
    """
    sheet_id = os.environ.get("GOOGLE_SHEET_ID")
    creds_path = os.environ.get("GOOGLE_CREDENTIALS_PATH")

    if not sheet_id or not creds_path or not os.path.isfile(creds_path):
        print("[INFO] Google Sheets連携は未設定のためスキップします")
        return

    try:
        import gspread
        from google.oauth2.service_account import Credentials
        from gspread_dataframe import set_with_dataframe
    except ImportError:
        print("[WARN] gspread / gspread-dataframe が未インストールのためスキップします")
        return

    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(sheet_id)

    for sheet_name, pivot in pivots.items():
        try:
            ws = sh.worksheet(sheet_name)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=sheet_name, rows=200, cols=50)
        ws.clear()
        set_with_dataframe(ws, pivot.reset_index())

    print(f"[INFO] Googleスプレッドシート({sheet_id})に書き出しました")


def main():
    all_rows = []
    for url in TARGET_MATCH_URLS:
        print(f"[INFO] checking {url}")
        rows = check_match(url)
        all_rows.extend(rows)
        time.sleep(1.5)  # サイト負荷軽減のためのウェイト

    append_csv(all_rows)
    print(f"[INFO] {len(all_rows)}件を {OUTPUT_CSV} に追記しました")

    if not os.path.isfile(OUTPUT_CSV):
        return

    full_df = pd.read_csv(OUTPUT_CSV, encoding="utf-8-sig")
    if full_df.empty:
        return
    full_df = build_price_display(full_df)
    pivots = build_pivots(full_df)

    export_pivot_xlsx(pivots)
    export_google_sheets(pivots)


if __name__ == "__main__":
    main()
