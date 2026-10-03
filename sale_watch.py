#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
各クラブの公式サイトから「今後の試合のチケット一般発売日」を取得し、
Googleスプレッドシートの「発売予定」シートに一覧として書き出すスクリプト。

monitor.py(価格チェック)とは別の目的・別のスケジュールで動かす想定。
Google Sheetsまわりの接続処理は monitor.py のものをそのまま再利用する。

【対応クラブ(現時点)】
- FC町田ゼルビア: 専用の「チケット販売スケジュール」ページ(テキストブロック形式)
- 横浜F・マリノス: 専用の「販売スケジュール」ページ(HTML表)

他のクラブは、それぞれ公式サイトの形式を個別に確認しながら追加していく。
"""

import io
import re
from datetime import datetime, timezone, timedelta

import pandas as pd
import requests
from bs4 import BeautifulSoup

from monitor import get_gspread_client, retry_on_transient_error, HEADERS, TARGETS_SHEET_NAME, CLUB_ABBR

JST = timezone(timedelta(hours=9))
SALE_SHEET_NAME = "発売予定"


def fetch(url: str) -> str | None:
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        resp.raise_for_status()
        return resp.text
    except requests.RequestException as e:
        print(f"[WARN] fetch failed: {url} ({e})")
        return None


# ── FC町田ゼルビア ──────────────────────────────────────────
def parse_zelvia(html: str) -> list[dict]:
    """
    zelvia.co.jp/stadium-ticket/schedule/ を解析する。
    ブロック形式: 【第N節】MM月DD日（曜）HH:MM〜 相手名 / スタジアム 会場名 / FC先行販売 日付 / 一般販売 日付
    """
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text("\n")

    # 「【」区切りでブロックに分割(先頭の「【」は除去されるので、以後は「【」で始まる形に戻す)
    blocks = ["【" + b for b in text.split("【")[1:]]

    rows = []
    for block in blocks:
        m = re.search(
            r"【(?P<section>[^】]+)】\s*"
            r"(?P<month>\d{1,2})月(?P<day>\d{1,2})日"
            r"（[^）]+）\s*(?P<time>[\d:]+|未定)\s*〜\s*"
            r"(?P<opponent>\S+?)\s*\n"
            r"スタジアム\s*\n?\s*(?P<venue>\S+)\s*\n"
            r".*?"
            r"一般販売\s*\n\s*(?P<general>[^\n〜]+)\s*〜",
            block,
            re.S,
        )
        if not m:
            continue
        rows.append({
            "club": "FC町田ゼルビア",
            "section": m.group("section").strip(),
            "match_date": f"{int(m.group('month'))}/{int(m.group('day'))}",
            "opponent": m.group("opponent").strip(),
            "venue": m.group("venue").strip(),
            "general_sale": m.group("general").strip(),
        })
    return rows


# ── 横浜F・マリノス ─────────────────────────────────────────
def parse_marinos(html: str) -> list[dict]:
    """
    f-marinos.com/ticket/schedule のHTML表を解析する。
    「一般販売」列は他の優先販売列と違い、開始日時のみ(終了日時が無い)ことを目印に判別する。
    """
    try:
        tables = pd.read_html(io.StringIO(html))
    except ValueError:
        return []
    if not tables:
        return []

    df = tables[0]
    rows = []

    single_date_re = re.compile(r"^\d{1,2}/\d{1,2}\([^)]+\)\s*\d{1,2}:\d{2}\s*[~〜～]?\s*$")
    range_date_re = re.compile(r"[~〜～]")

    for _, row in df.iterrows():
        cells = [str(c).strip() for c in row.tolist()]
        if not cells or cells[0].lower() == "nan":
            continue

        match_cell = cells[0]
        opponent_cell = cells[1] if len(cells) > 1 else ""

        match_date_m = re.search(r"(\d{1,2}/\d{1,2})", match_cell)
        opponent_m = re.match(r"([^\[]+)", opponent_cell)
        if not match_date_m or not opponent_m:
            continue

        # 「一般販売」列 = 単一日時(終了日時が無い)セルのうち、一番後ろにあるもの
        general_sale = ""
        for cell in cells[2:]:
            if single_date_re.match(cell):
                general_sale = cell
            elif range_date_re.search(cell):
                continue

        if not general_sale:
            continue

        rows.append({
            "club": "横浜F・マリノス",
            "section": match_cell.strip(),
            "match_date": match_date_m.group(1),
            "opponent": opponent_m.group(1).strip(),
            "venue": "",
            "general_sale": general_sale.strip(),
        })
    return rows


# ── 清水エスパルス ──────────────────────────────────────────
def parse_spulse(html: str) -> list[dict]:
    """
    s-pulse.co.jp/tickets/schedule を解析する。
    ブロック形式: ### 大会名・節 / M.D 曜日 時刻 K.O. 会場 / VS / 相手名 / (各種先行) / 一般販売 日付 時刻
    発売情報がまだ無い試合は「情報掲載までお待ち下さい。」となっており、その場合はスキップする。
    """
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text("\n")

    blocks = [b for b in re.split(r"(?:^|\n)###\s*", text) if b.strip()]

    rows = []
    for block in blocks:
        m = re.match(
            r"(?P<section>[^\n]+)\n+"
            r"(?P<month>\d{1,2})\.(?P<day>\d{1,2})\s+\S+\n+"
            r"(?:[\d:]+\s*K\.O\.\s*)?(?P<venue>[^\n]+)\n+"
            r"VS\n+"
            r"(?P<opponent>[^\n]+)\n+"
            r"(?P<rest>.*)",
            block,
            re.S,
        )
        if not m:
            continue

        rest = m.group("rest")
        general_m = re.search(
            r"一般販売\s*\n\s*(?P<gm>\d{1,2})/(?P<gd>\d{1,2})\([^)]*\)\s*\n?\s*(?P<gt>[\d:]+)",
            rest,
        )
        if not general_m:
            continue  # 「情報掲載までお待ち下さい」等、まだ発売日未定の試合はスキップ

        rows.append({
            "club": "清水エスパルス",
            "section": m.group("section").strip(),
            "match_date": f"{int(m.group('month'))}/{int(m.group('day'))}",
            "opponent": m.group("opponent").strip(),
            "venue": m.group("venue").strip(),
            "general_sale": f"{int(general_m.group('gm'))}/{int(general_m.group('gd'))} {general_m.group('gt')}",
        })
    return rows


# 既知のクラブ名(フルネーム)一覧。相手チーム名の特定に使う
ALL_CLUB_FULL_NAMES = sorted(set(CLUB_ABBR.keys()), key=len, reverse=True)


def find_opponent_name(text: str) -> str | None:
    for name in ALL_CLUB_FULL_NAMES:
        if name in text:
            return name
    return None


# ── 京都サンガF.C. ──────────────────────────────────────────
def parse_sanga(html: str) -> list[dict]:
    """
    sanga-fc.jp/ticket/schedule を解析する。
    「第N節」を区切りにブロック化し、各ブロック内の「一般販売」行から日付を取得する。
    対戦相手はブロック先頭付近に含まれるクラブ名(既知のクラブ名リスト)で特定する。
    販売日程がまだ無い試合(「試合開催日決定後、...」等)はスキップする。
    """
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text("\n")

    blocks = re.split(r"第(\d+節)", text)
    rows = []
    # re.splitで奇数インデックスに節番号、偶数インデックスにその前後のテキストが入る
    for i in range(1, len(blocks), 2):
        section = blocks[i]
        body = blocks[i + 1] if i + 1 < len(blocks) else ""
        head = body[:200]  # 日付・対戦相手はブロック冒頭付近にあるはず

        date_m = re.search(r"(\d{1,2})\.(\d{1,2})", head)
        opponent = find_opponent_name(head)
        general_m = re.search(
            r"一般販売\s*\n?\s*(\d{1,2})月(\d{1,2})日[^\d〜]*〜?\s*\n?\s*([\d:]+)",
            body,
            re.S,
        )

        if not (date_m and opponent and general_m):
            continue  # 対戦相手未定・販売日程未定の試合はスキップ

        rows.append({
            "club": "京都サンガF.C.",
            "section": f"第{section}",
            "match_date": f"{int(date_m.group(1))}/{int(date_m.group(2))}",
            "opponent": opponent,
            "venue": "",
            "general_sale": f"{int(general_m.group(1))}/{int(general_m.group(2))} {general_m.group(3)}",
        })
    return rows


# ── 対象クラブ一覧 ──────────────────────────────────────────
SALE_SOURCES = [
    {
        "club": "FC町田ゼルビア",
        "url": "https://www.zelvia.co.jp/stadium-ticket/schedule/",
        "parser": parse_zelvia,
    },
    {
        "club": "横浜F・マリノス",
        "url": "https://www.f-marinos.com/ticket/schedule",
        "parser": parse_marinos,
    },
    {
        "club": "清水エスパルス",
        "url": "https://www.s-pulse.co.jp/tickets/schedule",
        "parser": parse_spulse,
    },
    {
        "club": "京都サンガF.C.",
        "url": "https://www.sanga-fc.jp/ticket/schedule",
        "parser": parse_sanga,
    },
]


def collect_all() -> list[dict]:
    all_rows = []
    for source in SALE_SOURCES:
        print(f"[INFO] checking {source['club']} ({source['url']})")
        html = fetch(source["url"])
        if not html:
            continue
        rows = source["parser"](html)
        print(f"[INFO] {source['club']}: {len(rows)}件取得")
        all_rows.extend(rows)
    return all_rows


def export_to_sheet(rows: list[dict]):
    if not rows:
        print("[INFO] 書き出すデータがありません")
        return

    gc, sh = get_gspread_client()
    if gc is None or sh is None:
        print("[INFO] Google Sheets未設定のためスキップします")
        return

    import gspread

    try:
        ws = retry_on_transient_error(sh.worksheet, SALE_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SALE_SHEET_NAME, rows=200, cols=10)

    df = pd.DataFrame(rows)
    now = datetime.now(JST).strftime("%Y-%m-%d %H:%M")
    df.insert(0, "確認日時", now)

    retry_on_transient_error(ws.clear)
    from gspread_dataframe import set_with_dataframe
    retry_on_transient_error(set_with_dataframe, ws, df)
    print(f"[INFO] 「{SALE_SHEET_NAME}」シートに{len(df)}件書き出しました")

    # 「対象試合」シートの右隣に配置する
    try:
        all_ws = retry_on_transient_error(sh.worksheets)
        others = [w for w in all_ws if w.title != SALE_SHEET_NAME]
        names = [w.title for w in others]

        if TARGETS_SHEET_NAME in names:
            idx = names.index(TARGETS_SHEET_NAME)
            ordered = others[: idx + 1] + [w for w in all_ws if w.title == SALE_SHEET_NAME] + others[idx + 1:]
        else:
            ordered = [w for w in all_ws if w.title == SALE_SHEET_NAME] + others

        retry_on_transient_error(sh.reorder_worksheets, ordered)
        print(f"[INFO] 「{SALE_SHEET_NAME}」シートを「{TARGETS_SHEET_NAME}」の右隣に配置しました")
    except Exception as e:
        print(f"[WARN] シートの並び替えに失敗しました: {e}")


def main():
    rows = collect_all()
    export_to_sheet(rows)


if __name__ == "__main__":
    main()
