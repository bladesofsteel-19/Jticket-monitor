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
- 完売判定は「選択する」リンクが表示されているかどうかで行っています。
  実際の完売中の試合ページ(2625117番)で、完売中の席種には「選択する」が
  表示されないことを確認済みです。
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

# ── クラブのフルネーム → 略称(jleague-ticket.jpの表記に準拠) ──
CLUB_ABBR = {
    "鹿島アントラーズ": "鹿島",
    "水戸ホーリーホック": "水戸",
    "浦和レッズ": "浦和",
    "ジェフユナイテッド千葉": "千葉",
    "ジェフ千葉": "千葉",
    "柏レイソル": "柏",
    "ＦＣ東京": "FC東京",
    "FC東京": "FC東京",
    "東京ヴェルディ": "東京V",
    "ＦＣ町田ゼルビア": "町田",
    "FC町田ゼルビア": "町田",
    "川崎フロンターレ": "川崎",
    "横浜Ｆ・マリノス": "横浜FM",
    "横浜F・マリノス": "横浜FM",
    "清水エスパルス": "清水",
    "名古屋グランパス": "名古屋",
    "京都サンガF.C.": "京都",
    "ガンバ大阪": "G大阪",
    "セレッソ大阪": "C大阪",
    "ヴィッセル神戸": "神戸",
    "ファジアーノ岡山": "岡山",
    "サンフレッチェ広島": "広島",
    "アビスパ福岡": "福岡",
    "Ｖ・ファーレン長崎": "長崎",
    "V・ファーレン長崎": "長崎",
}


def abbreviate_card(text: str) -> str:
    """カード名に含まれるクラブのフルネームを略称に置き換える(マッチしなければそのまま)"""
    for full, abbr in CLUB_ABBR.items():
        text = text.replace(full, abbr)
    return text

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

# ── 監視したい試合URL(フォールバック用) ───────────────
# 通常はGoogleスプレッドシートの「対象試合」シートから読み込む。
# Google Sheets未設定の場合や、そちらにURLが無い場合に、この内蔵リストが使われる。
TARGET_MATCH_URLS = [
    "https://www.jleague-ticket.jp/sales/perform/2626898/001",
    "https://www.jleague-ticket.jp/sales/perform/2625117/001",
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
TARGETS_SHEET_NAME = "対象試合"

SOLD_OUT_PATTERNS = ["完売", "SOLD OUT", "販売終了"]
RANGE_PRICE_PATTERN = re.compile(r"([\d,]+)円\s*[~〜～]\s*([\d,]+)円\s*/\s*枚")
SINGLE_PRICE_PATTERN = re.compile(r"基本価格[：:]\s*([\d,]+)円\s*/\s*枚")
SELECT_LINK_TEXT = "選択する"


def get_gspread_client():
    """
    環境変数からGoogleスプレッドシートに接続する。
    設定が無い/ライブラリが無い場合は (None, None) を返す(呼び出し側で黙ってスキップする)。
    """
    sheet_id = os.environ.get("GOOGLE_SHEET_ID")
    creds_path = os.environ.get("GOOGLE_CREDENTIALS_PATH")

    if not sheet_id or not creds_path or not os.path.isfile(creds_path):
        return None, None

    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        print("[WARN] gspread が未インストールのためスキップします")
        return None, None

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
    gc = gspread.authorize(creds)

    try:
        sh = gc.open_by_key(sheet_id)
    except PermissionError as e:
        original = e.__cause__
        if original is not None and hasattr(original, "response"):
            print(f"[ERROR] Google Sheets APIエラー詳細: {original.response.text}")
        else:
            print(f"[ERROR] 元の例外: {repr(original)}")
        raise

    return gc, sh


def load_target_urls() -> list[tuple[str, str]]:
    """
    Googleスプレッドシート内の「対象試合」シートから (URL, B列のラベル) を読み込む。
    A列: 試合ページのURL / B列: シート名に使いたいチーム名など(任意)
    シートが無ければ、案内文付きで自動作成する。
    Google Sheets未設定/接続失敗時は、コード内蔵の TARGET_MATCH_URLS にフォールバックする。
    """
    try:
        gc, sh = get_gspread_client()
    except Exception as e:
        print(f"[WARN] Google Sheetsへの接続に失敗したため、内蔵リストを使います: {e}")
        gc, sh = None, None

    fallback = [(u, "") for u in TARGET_MATCH_URLS]

    if gc is None or sh is None:
        print("[INFO] Google Sheets未設定のため、コード内蔵のTARGET_MATCH_URLSを使います")
        return fallback

    import gspread

    try:
        ws = sh.worksheet(TARGETS_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=TARGETS_SHEET_NAME, rows=100, cols=2)
        ws.update([[
            "ここに試合ページのURLを1行に1つずつ貼ってください(例: https://www.jleague-ticket.jp/sales/perform/xxxxxxx/001)",
            "シート名にしたいチーム名(例: 東京V-柏)",
        ]])
        print(f"[INFO] 「{TARGETS_SHEET_NAME}」シートが無かったので新規作成しました。URLを貼ってから再実行してください")
        return []

    rows = ws.get_all_values()
    pairs = []
    for row in rows:
        url = row[0].strip() if len(row) > 0 else ""
        label = row[1].strip() if len(row) > 1 else ""
        if url.startswith("http"):
            pairs.append((url, label))

    if not pairs:
        print(f"[INFO] 「{TARGETS_SHEET_NAME}」シートにURLが見つからないため、内蔵リストを使います")
        return fallback

    print(f"[INFO] 「{TARGETS_SHEET_NAME}」シートから{len(pairs)}件のURLを読み込みました")
    return pairs


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
    date_match = re.search(r"\((\d{4})/(\d{2})/(\d{2})\)", title)
    match_date = date_match.group(0)[1:-1] if date_match else ""
    match_mmdd = f"{int(date_match.group(2))}{date_match.group(3)}" if date_match else ""

    raw_card = title.split("(")[0].split("|")[0].strip() if title else ""
    # "明治安田Ｊ１リーグ" 等のリーグ名表記を除去
    raw_card = re.sub(r"明治安田.{0,6}リーグ", "", raw_card).strip()
    card = abbreviate_card(raw_card)

    perform_id_match = re.search(r"/perform/(\d+)/", url)
    perform_id = perform_id_match.group(1) if perform_id_match else ""
    return {
        "perform_id": perform_id,
        "card": card,
        "match_date": match_date,
        "match_mmdd": match_mmdd,
        "url": url,
    }


def extract_seat_blocks(html: str) -> list[dict]:
    """
    ページ全体のテキストから席種ブロックを抽出する。
    各ブロックは概ね次のテキスト構造:
        <席種名見出し>
        基本価格：<価格>円/枚  または  <最小>円～<最大>円/枚
        発売 情報
        ...(発売期間などが続く)
        <席種名> [選択する]   ← 「選択する」が無い場合は完売
    """
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text("\n")
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    seats = []
    current_name = None
    block_start = 0
    BLOCK_NUMBER_RE = re.compile(r"^\d+$")

    for i, line in enumerate(lines):
        range_m = RANGE_PRICE_PATTERN.search(line)
        single_m = None if range_m else SINGLE_PRICE_PATTERN.search(line)

        if (range_m or single_m) and current_name:
            if range_m:
                price_min = range_m.group(1).replace(",", "")
                price_max = range_m.group(2).replace(",", "")
            else:
                price_min = price_max = single_m.group(1).replace(",", "")

            near_window = lines[max(0, i - 2): i + 8]
            is_dynamic = any("変動" in w for w in near_window)

            # このブロックの範囲は「次の番号行(次の席種の開始)」の手前まで
            block_end = len(lines)
            for j in range(i + 1, len(lines)):
                if BLOCK_NUMBER_RE.match(lines[j]):
                    block_end = j
                    break
            block_window = lines[i:block_end]

            if any(p in w for w in block_window for p in SOLD_OUT_PATTERNS):
                status = "完売"
            elif not any(SELECT_LINK_TEXT in w for w in block_window):
                # 「選択する」リンクが表示されていない = 完売(実例で確認済み)
                status = "完売"
            else:
                status = "販売中"

            seats.append({
                "seat_type": current_name,
                "price_min": price_min,
                "price_max": price_max,
                "dynamic": is_dynamic,
                "status": status,
            })
            current_name = None
        elif not range_m and not single_m and 2 <= len(line) <= 30 and not line.startswith("■") \
                and "円" not in line and "選択する" not in line and "発売" not in line \
                and "基本価格" not in line:
            # 席種名候補(短い行、価格や発売情報でないもの)
            current_name = line
    return seats


def check_match(url: str, label: str = "") -> list[dict]:
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
            "sheet_label": label,
            "match_date": meta["match_date"],
            "match_mmdd": meta["match_mmdd"],
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
        base = f"{int(row['price_max']):,}"
        if str(row.get("dynamic")).lower() in ("true", "1"):
            base += "[変動]"
        status = str(row.get("status", ""))
        if status == "完売":
            base += " 完売"
        elif status.startswith("要確認"):
            base += " 要確認"
        return base

    df = df.copy()
    df["price_display"] = df.apply(fmt, axis=1)
    return df


def build_pivots(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """試合(perform_id)ごとに 行=checked_at, 列=seat_type のピボット表を作る。
    シート名は「対象試合」シートB列のラベル(例: 東京V-柏)を優先し、
    無ければページから読み取ったcardを使う。日付(mmdd)を末尾に付与する。"""
    pivots = {}
    for perform_id, group in df.groupby("perform_id", dropna=False):
        label = ""
        if "sheet_label" in group.columns:
            non_empty = [v for v in group["sheet_label"] if str(v).strip()]
            label = str(non_empty[0]).strip() if non_empty else ""
        card = label if label else str(group["card"].iloc[0])

        mmdd = str(group["match_mmdd"].iloc[0]) if "match_mmdd" in group.columns else ""
        sheet_label = f"{card}_{mmdd}" if mmdd else card
        pivot = group.pivot_table(
            index="checked_at",
            columns="seat_type",
            values="price_display",
            aggfunc="first",
        )
        pivots[safe_sheet_name(sheet_label)] = pivot
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
    try:
        import gspread
        from gspread_dataframe import set_with_dataframe
    except ImportError:
        print("[WARN] gspread / gspread-dataframe が未インストールのためスキップします")
        return

    gc, sh = get_gspread_client()
    if gc is None or sh is None:
        print("[INFO] Google Sheets連携は未設定のためスキップします")
        return

    for sheet_name, pivot in pivots.items():
        try:
            ws = sh.worksheet(sheet_name)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=sheet_name, rows=200, cols=50)
        ws.clear()
        set_with_dataframe(ws, pivot.reset_index())

    print(f"[INFO] Googleスプレッドシートに書き出しました")


def main():
    targets = load_target_urls()

    all_rows = []
    for url, label in targets:
        print(f"[INFO] checking {url}")
        rows = check_match(url, label)
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
