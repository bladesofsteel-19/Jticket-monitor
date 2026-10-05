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
import os
import re
import time
from datetime import datetime, timezone, timedelta

import pandas as pd
import requests
from bs4 import BeautifulSoup

from monitor import get_gspread_client, retry_on_transient_error, HEADERS, TARGETS_SHEET_NAME, CLUB_ABBR

JST = timezone(timedelta(hours=9))
SALE_SHEET_NAME = "発売予定"


def text_with_img_alts(html: str) -> str:
    """
    HTMLをテキスト化する際、<img alt="..."> の内容もテキストとして含める。
    クラブ名がロゴ画像だけで表示され、文字としては存在しないページ対策。
    """
    soup = BeautifulSoup(html, "html.parser")
    for img in soup.find_all("img"):
        alt = img.get("alt", "").strip()
        if alt:
            img.replace_with(alt)
    return soup.get_text("\n")


def dedupe_name(s: str) -> str:
    """
    ロゴ画像のalt属性と、直後のテキストでチーム名が重複しているケースを1つにまとめる。
    例: 「浦和レッズ浦和レッズ」「浦和レッズ 浦和レッズ」→「浦和レッズ」
    """
    s = s.strip()
    n = len(s)
    if n % 2 == 0 and n > 0 and s[: n // 2] == s[n // 2:]:
        return s[: n // 2]
    parts = s.split()
    if len(parts) == 2 and parts[0] == parts[1]:
        return parts[0]
    return s


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
    text = text_with_img_alts(html)

    # 「【」区切りでブロックに分割(先頭の「【」は除去されるので、以後は「【」で始まる形に戻す)
    blocks = ["【" + b for b in text.split("【")[1:]]

    rows = []
    for block in blocks:
        m = re.search(
            r"【(?P<section>[^】]+)】\s*"
            r"(?P<month>\d{1,2})月(?P<day>\d{1,2})日"
            r"（[^）]+）\s*(?P<time>[\d:]+|未定)\s*〜\s*"
            r"(?P<opponent>[^\n]+?)\s*\n"
            r".*?スタジアム\s*\n?\s*(?P<venue>\S+)\s*\n"
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
            "opponent": dedupe_name(m.group("opponent")),
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
    「M.D 曜日(英字3文字)」という日付表記を目印にブロックを区切る
    (見出し記号に頼らず、実際に表示されている日付バッジのテキストを基準にする)。
    発売情報がまだ無い試合は「情報掲載までお待ち下さい。」となっており、その場合はスキップする。
    """
    text = text_with_img_alts(html)

    anchors = list(re.finditer(r"(\d{1,2})\.(\d{1,2})\s+(SAT|SUN|MON|TUE|WED|THU|FRI)", text))
    rows = []
    for idx, am in enumerate(anchors):
        start = am.start()
        end = anchors[idx + 1].start() if idx + 1 < len(anchors) else len(text)
        block = text[start:end]

        prev_end = anchors[idx - 1].end() if idx > 0 else 0
        pre_text = text[prev_end:start]
        section_m = re.search(r"第\d+節", pre_text)
        if not section_m:
            section_m = re.search(r"天皇杯[^\n]*", pre_text)
        section = section_m.group(0).strip() if section_m else "大会不明"

        vs_m = re.search(r"K\.O\.\s*([^\n]+)\n+VS\n+([^\n]+)", block)
        if not vs_m:
            continue
        venue, opponent = vs_m.group(1).strip(), vs_m.group(2).strip()

        general_m = re.search(
            r"一般販売\s*\n\s*(\d{1,2})/(\d{1,2})\([^)]*\)\s*\n?\s*([\d:]+)",
            block,
        )
        if not general_m:
            continue  # 「情報掲載までお待ち下さい」等、まだ発売日未定の試合はスキップ

        rows.append({
            "club": "清水エスパルス",
            "section": section,
            "match_date": f"{int(am.group(1))}/{int(am.group(2))}",
            "opponent": opponent,
            "venue": venue,
            "general_sale": f"{int(general_m.group(1))}/{int(general_m.group(2))} {general_m.group(3)}",
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
    text = text_with_img_alts(html)

    blocks = re.split(r"第(\d+節)", text)
    rows = []
    # re.splitで奇数インデックスに節番号、偶数インデックスにその前後のテキストが入る
    for i in range(1, len(blocks), 2):
        section = blocks[i]
        body = blocks[i + 1] if i + 1 < len(blocks) else ""
        head = body[:200]  # 日付・対戦相手はブロック冒頭付近にあるはず

        date_m = re.search(r"(\d{1,2})\.(\d{1,2})", head)
        opponent = find_opponent_name(head)

        # 「一般販売」の直後、多少余計な文字(罫線・「販売中」等)が挟まっても
        # 最初に現れる日付+時刻パターンを拾う
        general_m = None
        gm_idx = body.rfind("一般販売")
        if gm_idx != -1:
            window = body[gm_idx: gm_idx + 150]
            general_m = re.search(r"(\d{1,2})月(\d{1,2})日[^\d]*?(\d{1,2}:\d{2})", window)

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
        "expect_marker": "一般販売",
    },
    {
        "club": "横浜F・マリノス",
        "url": "https://www.f-marinos.com/ticket/schedule",
        "parser": parse_marinos,
        "expect_marker": "一般販売",
    },
    # 清水エスパルス・京都サンガF.C.は、サイト側のボット対策により
    # GitHub Actionsからの取得が(Playwrightを使っても)できなかったため、
    # 一旦対象から外している。パーサー自体(parse_spulse / parse_sanga)は残してあるので、
    # 将来別の取得方法が見つかれば再度有効化できる。
]


DEBUG_TEXT_DUMP = os.environ.get("SALE_WATCH_DEBUG") == "1"


def fetch_with_marker_retry(source: dict, max_attempts: int = 3, delay: int = 8) -> str | None:
    """
    取得したHTMLに、本来あるはずの目印文字列(expect_marker)が含まれているか確認する。
    無ければ、ボット対策等で内容が間引かれて返された可能性があるとみなし、
    少し待って再取得する。
    """
    marker = source.get("expect_marker")
    for attempt in range(1, max_attempts + 1):
        html = fetch(source["url"])
        if not html:
            if attempt < max_attempts:
                time.sleep(delay)
                continue
            return None

        if not marker or marker in html:
            return html

        print(f"[WARN] {source['club']}: 目印「{marker}」が見つかりません(内容が間引かれた可能性)。"
              f"{delay}秒待って再取得します({attempt}/{max_attempts})")
        if attempt < max_attempts:
            time.sleep(delay)

    return html  # 最終的に目印が無くても、最後に取得できた内容をそのまま返す(呼び出し側で0件になるだけ)


def collect_all() -> list[dict]:
    all_rows = []
    for source in SALE_SOURCES:
        print(f"[INFO] checking {source['club']} ({source['url']})")
        html = fetch_with_marker_retry(source)
        if not html:
            continue

        if DEBUG_TEXT_DUMP:
            sample = text_with_img_alts(html)[:1500]
            print(f"[DEBUG] ---- {source['club']} の実際のテキスト(先頭1500文字) ----")
            print(repr(sample))
            print("[DEBUG] ---- ここまで ----")

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
