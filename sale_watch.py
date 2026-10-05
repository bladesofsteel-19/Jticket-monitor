#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
各クラブの公式サイトから「今後の試合のチケット一般発売日」を取得し、
Googleスプレッドシートの「発売予定」シートに一覧として書き出すスクリプト。

monitor.py(価格チェック)とは別の目的・別のスケジュールで動かす想定。
Google Sheetsまわりの接続処理は monitor.py のものをそのまま再利用する。

【対応クラブ(現時点)】
- 1ページ解析方式(parser): FC町田ゼルビア / 横浜F・マリノス / ガンバ大阪 / セレッソ大阪 /
  アビスパ福岡 / 浦和レッズ / 名古屋グランパス
- 記事巡回方式(collector): 川崎フロンターレ(Playwright必須) / 東京ヴェルディ(requestsのみ)
- 1ページ解析方式(GAS経由の取得にも対応): 清水エスパルス / 京都サンガF.C.
- 1ページ解析方式: ヴィッセル神戸

他のクラブは、それぞれ公式サイトの形式を個別に確認しながら追加していく。
"""

import io
import os
import re
import time
import unicodedata
from datetime import date, datetime, timezone, timedelta
from urllib.parse import urljoin

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


def decode_response(resp: requests.Response) -> str:
    """
    レスポンス本文を正しい文字コードで文字列にする。
    サーバーが Content-Type に charset を付けていない場合(例: 名古屋の発売予定ページは
    「text/html;」だけ)、requests は規格上の既定値 ISO-8859-1 で解釈してしまい、
    日本語が「ç\x99ºå£²...」のように文字化けする。その場合は UTF-8 を優先して試し、
    だめなら本文から推定した文字コード(Shift_JIS等)で読む。
    """
    content_type = resp.headers.get("Content-Type", "").lower()
    if "charset=" in content_type:
        return resp.text
    try:
        return resp.content.decode("utf-8")
    except UnicodeDecodeError:
        resp.encoding = resp.apparent_encoding or "utf-8"
        return resp.text


def fetch(url: str) -> str | None:
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        resp.raise_for_status()
        return decode_response(resp)
    except requests.RequestException as e:
        status = getattr(getattr(e, "response", None), "status_code", None)
        print(f"[WARN] fetch failed: {url} (status={status}, {type(e).__name__}: {e})")
        return None


def fetch_with_playwright(url: str, click_category: bool = False) -> str | None:
    """
    JavaScriptで後から内容が表示されるページ用に、実際のブラウザ(Chromium)で取得する。
    川崎フロンターレ(frontale.co.jp)専用。
    click_category=True の場合のみ、「チケット」カテゴリーのリンクをクリックする
    (記事ページ自体にも「チケット」という文字があり、誤ってクリックして
    記事から離脱してしまうのを防ぐため、一覧ページ取得時だけに限定する)。
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("[WARN] playwrightが未インストールのためスキップします")
        return None

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            context = browser.new_context(
                user_agent=HEADERS.get("User-Agent"),
                locale="ja-JP",
            )
            page = context.new_page()
            page.goto(url, wait_until="networkidle", timeout=30000)

            if click_category:
                # 「チケット」カテゴリーのリンク/ボタンがあれば、実際にクリックして
                # 一覧を絞り込む(URL直接アクセスだけでは反映されないサイト対策)
                try:
                    category_link = page.get_by_text("チケット", exact=True).first
                    if category_link.count() > 0:
                        category_link.click(timeout=5000)
                        page.wait_for_timeout(3000)
                except Exception:
                    pass

            # 遅延読み込み(スクロールで初めて表示される)対策として、
            # 下までスクロールしてから少し待つ
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            page.wait_for_timeout(3000)
            content = page.content()
            browser.close()
            return content
    except Exception as e:
        print(f"[WARN] playwrightでの取得に失敗しました: {url} ({e})")
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
            "general_sale": format_sale_datetime(m.group("general")),
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
            "venue": re.sub(r"\s+", "", (re.search(r"\[([^\]]+)\]", opponent_cell) or [None, ""])[1]),
            "general_sale": format_sale_datetime(general_sale),
        })
    return rows


# ── 清水エスパルス ──────────────────────────────────────────
def parse_spulse(html: str) -> list[dict]:
    """
    s-pulse.co.jp/tickets/schedule を解析する。
    1試合ごとに次の並びで書かれている(全角英数字はNFKCで半角にしてから扱う):
        ロゴ:明治安田J1リーグ 第2節      ← 見出し(大会名・節)
        8.15 SAT                         ← 日付バッジ(これを区切りにする)
        18:30 K.O. IAIスタジアム日本平
        VS
        ロゴ:横浜F・マリノス              ← エンブレム画像のalt
        横浜F・マリノス
        ... 一般販売 / 7/30(木) / 10:00
    発売情報がまだ無い試合は「情報掲載までお待ち下さい。」となっており、その場合はスキップする。
    """
    text = nfkc(text_with_img_alts(html))

    anchors = list(re.finditer(r"(\d{1,2})\.(\d{1,2})\s+(SAT|SUN|MON|TUE|WED|THU|FRI)", text))
    rows = []
    for idx, am in enumerate(anchors):
        end_pos = anchors[idx + 1].start() if idx + 1 < len(anchors) else len(text)
        block = text[am.end():end_pos]

        # 見出し(大会名・節)は日付バッジの直前の行
        prev_end = anchors[idx - 1].end() if idx > 0 else 0
        pre_lines = [ln.strip() for ln in text[prev_end:am.start()].splitlines() if ln.strip()]
        section = extract_section(" ".join(pre_lines[-2:])) if pre_lines else ""

        # 「K.O. 会場名 ... VS ... 相手名」。VSの後はエンブレムのalt(「ロゴ:相手名」)と相手名が続く
        vs_m = re.search(r"K\.O\.\s*([^\n]*)\n\s*VS\s*\n\s*([^\n]+)", block)
        if not vs_m:
            continue
        venue = vs_m.group(1).strip()
        opponent = re.sub(r"^ロゴ\s*[:：]\s*", "", vs_m.group(2).strip())

        general_m = re.search(
            r"一般販売\s*(\d{1,2})/(\d{1,2})\s*\([^)]*\)\s*(\d{1,2}:\d{2})",
            block,
        )
        if not general_m:
            continue  # 「情報掲載までお待ち下さい」等、まだ発売日未定の試合はスキップ

        rows.append({
            "club": "清水エスパルス",
            "section": section,
            "match_date": f"{int(am.group(1))}/{int(am.group(2))}",
            "opponent": normalize_opponent(opponent),
            "venue": venue,
            "general_sale": f"{int(general_m.group(1))}/{int(general_m.group(2))} {general_m.group(3)}",
        })
    return rows


# 既知のクラブ名(フルネーム)一覧。相手チーム名の特定に使う
ALL_CLUB_FULL_NAMES = sorted(set(CLUB_ABBR.keys()), key=len, reverse=True)


def find_opponent_name(text: str) -> str | None:
    """
    テキスト中に含まれる既知のクラブ名(フルネーム)を返す。
    「ＦＣ町田ゼルビア」のような全角表記でも一致するよう、両方をNFKCで揃えて比較する。
    """
    t = nfkc(text)
    for name in ALL_CLUB_FULL_NAMES:
        if nfkc(name) in t:
            return name
    return None


# ── 京都サンガF.C. ──────────────────────────────────────────
def parse_sanga(html: str) -> list[dict]:
    """
    sanga-fc.jp/ticket/schedule を解析する。
    画面上は試合名をクリックすると販売日程が開く作りだが、これは表示/非表示の切り替えだけで、
    全試合の販売日程の表は最初からHTMLに含まれている(クリック操作は不要)。
    1試合ごとの並び(NFKC後):
        明治安田J1リーグ                     ← 大会名(ACLは「AFCチャンピオンズリーグElite / リーグステージ」)
        第9節 10.10 (土) 19:00 FC町田ゼルビア
        受付・販売種別 | ... | 一般販売 (販売中) | 9月26日(土) 12:00～
    「第N節」を区切りにし、各区間の中の「一般販売」の日時を拾う。
    販売日程がまだ無い試合(「試合開催日決定後、...」や表が無いもの)はスキップする。
    """
    text = nfkc(text_with_img_alts(html))
    # 実際のHTMLはインデント用の空白が非常に多く(「第4節」の後に500文字以上の空白が続く)、
    # 「第N節の直後150文字」に日付や相手名が収まらなかったため、空白と空行を詰めてから扱う
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" *\n[\s]*", "\n", text)

    anchors = list(re.finditer(r"第\s*(\d+)\s*節", text))
    rows = []
    for idx, am in enumerate(anchors):
        end_pos = anchors[idx + 1].start() if idx + 1 < len(anchors) else len(text)
        body = text[am.end():end_pos]
        head = body[:150]  # 日付・キックオフ・対戦相手は「第N節」の直後にある

        # 大会名は「第N節」の直前の1〜2行
        prev_end = anchors[idx - 1].end() if idx > 0 else 0
        pre_lines = [ln.strip() for ln in text[prev_end:am.start()].splitlines() if ln.strip()]
        section = extract_section(" ".join(pre_lines[-2:] + [am.group(0)]))

        date_m = re.search(r"(\d{1,2})\.(\d{1,2})", head)
        opponent = find_opponent_name(head)
        if not opponent:
            # ACLの海外クラブ等、既知のクラブ名リストに無い相手は「キックオフ時刻(または未定)の後ろ」を相手名とみなす
            opp_m = re.search(r"(?:\d{1,2}:\d{2}|キックオフ未定)\s*([^\n|]+)", head)
            opponent = opp_m.group(1).strip() if opp_m else None

        general_m = None
        gm_idx = body.find("一般販売")
        if gm_idx != -1:
            window = body[gm_idx: gm_idx + 100]
            general_m = re.search(r"(\d{1,2})月(\d{1,2})日[^\d]*?(\d{1,2}:\d{2})", window)

        if not (date_m and opponent and general_m):
            continue  # 対戦相手未定・販売日程未定の試合はスキップ

        rows.append({
            "club": "京都サンガF.C.",
            "section": section,
            # 日程未確定の試合は「2.13(土) or 2.14(日)」のように候補日が2つあるが、1つ目を使う
            "match_date": f"{int(date_m.group(1))}/{int(date_m.group(2))}",
            "opponent": normalize_opponent(opponent),
            "venue": "",
            "general_sale": f"{int(general_m.group(1))}/{int(general_m.group(2))} {general_m.group(3)}",
        })
    return rows


# ── ガンバ大阪 ──────────────────────────────────────────
def parse_gamba(html: str) -> list[dict]:
    """
    gamba-osaka.net/ticket/schedule/ を解析する。
    「HOME」を区切りにブロック化し、各ブロック内の日付・対戦相手(vs. 〜)・
    一般販売の行から情報を取得する。一般販売が「未定」の試合はスキップする。
    """
    text = text_with_img_alts(html)

    parts = text.split("HOME")
    rows = []
    for i in range(1, len(parts)):
        block = parts[i]
        # 大会名・節(「明治安田J1リーグ 第12節」等)は「HOME」の直前の行に書かれている
        prev_lines = [ln.strip() for ln in parts[i - 1].splitlines() if ln.strip()]
        section = extract_section(prev_lines[-1]) if prev_lines else ""

        # 日付候補は複数ある(FC先行販売の日付も同じ形式のため)。
        # 「日付(曜)→時刻→＠スタジアム」の並びが直後に続くものだけを、本当の試合日として扱う。
        date_m = re.search(
            r"(\d{1,2})\.(\d{1,2})\s*\([^)]*\)\s*\n\s*[\d:]+\s*\n\s*＠\s*([^\n]+)",
            block,
        )

        opponent_m = re.search(r"vs\.\s*([^\n]+)", block)
        if not (date_m and opponent_m):
            continue

        # U-21(セカンドチーム)の試合は対象外。見出し(「U-21 Jリーグ …」)か相手名(「U-21 清水エスパルス」)で判定
        if section.startswith("U-21") or re.search(r"U-?21", nfkc(opponent_m.group(1))):
            continue

        general_m = re.search(r"一般販売\s*\n\s*([^\n未]+)", block)
        if not general_m:
            continue  # 「未定」等、まだ発売日未定の試合はスキップ

        rows.append({
            "club": "ガンバ大阪",
            "section": section,
            "match_date": f"{int(date_m.group(1))}/{int(date_m.group(2))}",
            "opponent": dedupe_name(opponent_m.group(1)),
            "venue": date_m.group(3).strip(),  # 「＠パナスタ」の部分
            "general_sale": format_sale_datetime(general_m.group(1)),
        })
    return rows


# ── セレッソ大阪 ──────────────────────────────────────────
def parse_cerezo(html: str) -> list[dict]:
    """
    cerezo.jp/ticket/ の「販売スケジュール」部分(表の下にあるテキスト形式の
    繰り返しブロック)を解析する。
    ブロック形式: 第N節 M/D（曜）HH:MM vs 相手 / 会場：... / 一般販売 ・企画チケット販売 / 日付
    """
    text = text_with_img_alts(html)

    blocks = re.split(r"第(\d+)節\s+", text)
    rows = []
    for i in range(1, len(blocks), 2):
        section = blocks[i]
        body = blocks[i + 1] if i + 1 < len(blocks) else ""
        head = body[:120]

        m = re.match(
            r"(\d{1,2})/(\d{1,2})（[^）]+）\s*(\d{1,2}:\d{2})\s*vs\s*(\S+)",
            head,
        )
        if not m:
            continue

        gm_idx = body.rfind("一般販売")
        if gm_idx == -1:
            continue
        window = body[gm_idx: gm_idx + 100]
        general_m = re.search(r"(\d{1,2})/(\d{1,2})\([^)]*\)\s*([\d:]+)", window)
        if not general_m:
            continue

        rows.append({
            "club": "セレッソ大阪",
            "section": f"第{section}節",
            "match_date": f"{int(m.group(1))}/{int(m.group(2))}",
            "opponent": dedupe_name(m.group(4)),
            "venue": (re.search(r"会場\s*[:：]\s*([^\n]+)", body[:400]) or [None, ""])[1].strip(),
            "general_sale": f"{int(general_m.group(1))}/{int(general_m.group(2))} {general_m.group(3)}",
        })
    return rows


# ── アビスパ福岡 ──────────────────────────────────────────
def parse_avispa(html: str) -> list[dict]:
    """
    avispa.co.jp/match/series/2026-27 を解析する。
    試合一覧(ホーム/アウェイ混在)から「HOME」のブロックだけを対象にする。
    一般発売日は「一般 M.D[曜]HH:MM〜」という形式(ピリオド区切り)。
    """
    text = text_with_img_alts(html)

    blocks = re.split(r"\n(HOME|AWAY)\s", text)
    rows = []
    for i in range(1, len(blocks), 2):
        if blocks[i] != "HOME":
            continue
        chunk = blocks[i + 1] if i + 1 < len(blocks) else ""

        date_m = re.search(r"(\d{1,2})/(\d{1,2})", chunk[:200])
        opponent = find_opponent_name(chunk[:300])
        general_m = re.search(
            r"一般\s*(\d{1,2})\.(\d{1,2})\[[^\]]*\]\s*(\d{1,2}:\d{2})", chunk
        )

        if not (date_m and opponent and general_m):
            continue  # 発売日未定の試合はスキップ

        rows.append({
            "club": "アビスパ福岡",
            # ブロック冒頭は「会場名 → 明治安田J1リーグ 第9節 → 日付」の順
            "section": extract_section(chunk[:200]),
            "match_date": f"{int(date_m.group(1))}/{int(date_m.group(2))}",
            "opponent": opponent,
            "venue": chunk.strip().splitlines()[0].strip() if chunk.strip() else "",  # 「HOME ベスト電器スタジアム」
            "general_sale": f"{int(general_m.group(1))}/{int(general_m.group(2))} {general_m.group(3)}",
        })
    return rows


# ── 浦和レッズ ──────────────────────────────────────────
def parse_urawa(html: str) -> list[dict]:
    """
    urawa-reds.co.jp/ticket/saleperiod.html のHTML表を解析する。
    列: 試合日時・会場 / 大会・節 / 対戦相手 / シーズンチケット / 各種先行販売 / 一般販売 / リンク
    """
    try:
        tables = pd.read_html(io.StringIO(html))
    except ValueError:
        return []
    if not tables:
        return []

    df = tables[0]
    rows = []
    for _, row in df.iterrows():
        cells = [str(c).strip() for c in row.tolist()]
        if len(cells) < 6 or cells[0].lower() == "nan":
            continue

        match_cell = cells[0]
        opponent_cell = cells[2]
        general_cell = cells[5]

        date_m = re.search(r"(\d{1,2})/(\d{1,2})", match_cell)
        general_m = re.search(r"(\d{1,2})/(\d{1,2})\([^)]*\)\s*([\d:]+)", general_cell)
        if not (date_m and general_m and opponent_cell):
            continue

        rows.append({
            "club": "浦和レッズ",
            "section": extract_section(cells[1]),  # 「大会・節」列
            "match_date": f"{int(date_m.group(1))}/{int(date_m.group(2))}",
            "opponent": dedupe_name(opponent_cell),
            # 「10/21(水) 19:30キックオフ 埼玉スタジアム2002」の「キックオフ」以降
            "venue": (re.search(r"キックオフ\s*(.+)$", nfkc(match_cell)) or [None, ""])[1].strip(),
            "general_sale": f"{int(general_m.group(1))}/{int(general_m.group(2))} {general_m.group(3)}",
        })
    return rows


# ── ヴィッセル神戸 ──────────────────────────────────────────
# 公式サイトの注記:「販売開始時間はロイヤル、レギュラー、一般販売は10:00～」
VISSEL_DEFAULT_SALE_TIME = "10:00"


def parse_vissel(html: str) -> list[dict]:
    """
    vissel-kobe.co.jp/ticket/schedule/ を解析する。
    1試合ごとの並び(NFKC・空白整理後):
        明治安田J1リーグ        ← 大会ロゴのalt(ACLは「ACL Elite」)
        第12節                  ← ACLは「MD2」
        10/24 (土)              ← 試合日(次の行がキックオフ時刻)。これを区切りにする
        15:00
        町田 町田               ← エンブレムのalt + クラブ名(略称)
        ノエスタ                ← スタジアム
        ロイヤル: 販売中 ... 一般: 販売中   または   一般: 10/14(水)
    一般販売の欄には日付しか無いので、時刻は公式サイト記載の10:00を補う。
    すでに発売済みの試合は日付が消えて「販売中」になるため、その場合は「販売中」と出力する。
    """
    text = nfkc(text_with_img_alts(html))
    # このサイトは改行が「\r\n」で、インデントの空白も非常に多いので、改行と空白を揃えてから扱う
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" *\n\s*", "\n", text)

    # 試合日 = 「日付(曜)」の次の行がキックオフ時刻のもの(販売日の「10/10(土)」と区別するため)
    anchors = list(re.finditer(
        r"(\d{1,2})/(\d{1,2})\s*\([^)]*\)\s*\n(?:\d{1,2}:\d{2}|未定|-+)", text))
    if not anchors and DEBUG_TEXT_DUMP:
        idx = text.find("対戦相手")
        print("[DEBUG] ヴィッセル神戸: 試合日が見つかりません。整形後のテキスト:")
        print(repr(text[max(0, idx):max(0, idx) + 1500]))

    rows = []
    for i, am in enumerate(anchors):
        prev_end = anchors[i - 1].end() if i > 0 else 0
        next_start = anchors[i + 1].start() if i + 1 < len(anchors) else len(text)

        # 大会名・節は試合日の直前の2行
        pre_lines = [ln.strip() for ln in text[prev_end:am.start()].splitlines() if ln.strip()]
        section = extract_section(" ".join(pre_lines[-2:])) if pre_lines else ""

        after = text[am.end():next_start]
        # 試合日・キックオフの後ろ: 相手名(エンブレムのaltと重複することがある) → スタジアム → 販売日程
        after_lines = []
        for ln in after.splitlines():
            ln = ln.strip()
            if ln and (not after_lines or after_lines[-1] != ln):
                after_lines.append(ln)
        if len(after_lines) < 2:
            continue
        opponent = normalize_opponent(dedupe_name(after_lines[0]))
        venue = after_lines[1]

        general_m = re.search(r"一般\s*[:：]\s*([^\n]+)", after)
        if not general_m:
            continue
        general_raw = general_m.group(1).strip()
        date_m = re.search(r"(\d{1,2})/(\d{1,2})", general_raw)
        if date_m:
            general = f"{int(date_m.group(1))}/{int(date_m.group(2))} {VISSEL_DEFAULT_SALE_TIME}"
        elif "販売中" in general_raw:
            general = "販売中"
        else:
            continue  # 「未定」等

        rows.append({
            "club": "ヴィッセル神戸",
            "section": section,
            "match_date": f"{int(am.group(1))}/{int(am.group(2))}",
            "opponent": opponent,
            "venue": venue,
            "general_sale": general,
        })
    return rows


# ── 川崎フロンターレ ──────────────────────────────────────────
TICKET_TITLE_PHRASE = "「チケット販売」のお知らせ"


def collect_frontale() -> list[dict]:
    """
    frontale.co.jp は専用の発売日一覧ページがJSで空のままだったため、
    ニュース一覧(または記事内の「関連するお知らせ」)から
    「M/D 相手「チケット販売」のお知らせ」という記事を見つけて個別に読む。
    """
    seed_urls = [
        "https://www.frontale.co.jp/info/ticket/",
        "https://www.frontale.co.jp/info/index.html",
    ]
    list_html = None
    used_seed = None
    for seed in seed_urls:
        list_html = fetch_with_playwright(seed, click_category=True)
        if list_html:
            used_seed = seed
            break
    if not list_html:
        print("[WARN] 川崎フロンターレ: ニュース一覧の取得に失敗しました")
        return []

    if DEBUG_TEXT_DUMP:
        print(f"[DEBUG] 川崎フロンターレ: 一覧ページ取得成功 ({used_seed})")

    soup = BeautifulSoup(list_html, "html.parser")
    links = []
    all_titles_sample = []
    for a in soup.find_all("a", href=True):
        title = a.get_text(strip=True)
        if title and len(all_titles_sample) < 20:
            all_titles_sample.append(title)
        if TICKET_TITLE_PHRASE in title:
            href = a["href"]
            full_url = href if href.startswith("http") else f"https://www.frontale.co.jp{href}"
            if full_url not in links:
                links.append(full_url)
    links = links[:10]  # 直近いくつかだけ

    if DEBUG_TEXT_DUMP:
        print(f"[DEBUG] 川崎フロンターレ: リンク候補{len(links)}件見つかりました: {links}")
        print(f"[DEBUG] 川崎フロンターレ: ページ内のリンクテキスト例(先頭20件): {all_titles_sample}")

    rows = []
    for url in links:
        html = fetch_with_playwright(url)
        if not html:
            continue
        text = text_with_img_alts(html)

        match_m = re.search(
            r"(\d{1,2})月(\d{1,2})日[^\n]*?第(\d+)節\s*(\S+?)戦",
            text,
        )
        general_m = re.search(
            r"一般[：:]\s*(\d{1,2})月(\d{1,2})日[^\d]*?(\d{1,2}:\d{2})",
            text,
        )
        if not (match_m and general_m):
            if DEBUG_TEXT_DUMP:
                print(f"[DEBUG] 川崎フロンターレ: {url} の解析に失敗(match_m={bool(match_m)}, general_m={bool(general_m)})")
                print(repr(text[:1500]))
            continue

        rows.append({
            "club": "川崎フロンターレ",
            "section": f"第{int(match_m.group(3))}節",
            "match_date": f"{int(match_m.group(1))}/{int(match_m.group(2))}",
            "opponent": dedupe_name(match_m.group(4)),
            "venue": "",
            "general_sale": f"{int(general_m.group(1))}/{int(general_m.group(2))} {general_m.group(3)}",
        })
        time.sleep(1)

    return rows


# ── 名古屋グランパス・東京ヴェルディ共通の補助関数 ─────────────────────
def nfkc(s) -> str:
    """全角数字・全角括弧・全角コロン等を半角に揃える(「10：00」→「10:00」など)"""
    return unicodedata.normalize("NFKC", str(s)).strip()


FREE_SALE_DT_RE = re.compile(
    r"(\d{1,2})(?:/|\.|月)(\d{1,2})日?\s*(?:\([^)]*\))?\s*(?:(\d{1,2}):(\d{2}))?"
)


def format_sale_datetime(raw: str) -> str:
    """
    クラブごとにバラバラな発売日時の書き方を、他クラブと同じ「M/D H:MM」に揃える。
      「7月24日（金）12:00」(町田) → 「7/24 12:00」
      「9.19(土)10:00 ～」(ガンバ) → 「9/19 10:00」
    日付が読み取れなければ元の文字列をそのまま返す(情報を落とさないため)。
    """
    t = nfkc(raw)
    m = FREE_SALE_DT_RE.search(t)
    if not m:
        return raw.strip()
    md = f"{int(m.group(1))}/{int(m.group(2))}"
    return f"{md} {int(m.group(3))}:{m.group(4)}" if m.group(3) else md


def extract_section(text: str) -> str:
    """
    大会名・節の表記を「C列(section)」用の短い形に揃える。
      「明治安田J1リーグ 第12節」→「第12節」
      「U-21 Jリーグ 交流戦ラウンド第4節」→「U-21 第4節」
      「天皇杯 JFA 第106回全日本サッカー選手権大会 3回戦」→「天皇杯 3回戦」
      「Jリーグ YBCルヴァンカップ 1stラウンド2回戦」→「ルヴァンカップ 2回戦」
      「AFCチャンピオンズリーグElite MD3」→「ACLE MD3」
    該当しなければ空文字。
    """
    t = nfkc(text)
    round_m = re.search(r"(\d+回戦|準々決勝|準決勝|決勝|プレーオフ)", t)
    rnd = f" {round_m.group(1)}" if round_m else ""
    if "天皇杯" in t:
        return f"天皇杯{rnd}"
    if "ルヴァン" in t:
        return f"ルヴァンカップ{rnd}"
    if re.search(r"AFC|ACL", t):
        name = "ACL2" if re.search(r"Two|ACL\s*2", t) else "ACLE"
        md = re.search(r"MD\s*(\d+)", t)
        if md:
            return f"{name} MD{md.group(1)}"
        sec = re.search(r"第\s*(\d+)\s*節", t)  # 京都の表記「リーグステージ 第2節」
        return f"{name} 第{int(sec.group(1))}節" if sec else f"{name}{rnd}"
    m = re.search(r"第\s*(\d+)\s*節", t)
    if m:
        prefix = "U-21 " if re.search(r"U-?21", t) else ""
        return f"{prefix}第{int(m.group(1))}節"
    return ""


def _flat_col(c) -> str:
    """表の列名を比較しやすい形にする(<br>由来の空白を除去、複数段ヘッダーは連結)"""
    if isinstance(c, tuple):
        c = " ".join(str(x) for x in c)
    return re.sub(r"\s+", "", nfkc(c))


def table_header_candidates(df: pd.DataFrame):
    """
    表の見出し候補を返す。見出しが<th>でなく<td>で書かれている表では、
    pandasが見出し行を通常のデータ行として扱うため、1行目を見出しとみなす候補も返す。
    戻り値: (列名リスト, データ部分のDataFrame) の候補を順に yield
    """
    yield [_flat_col(c) for c in df.columns], df
    if len(df) > 0:
        yield [_flat_col(c) for c in df.iloc[0].tolist()], df.iloc[1:]


def read_tables(html: str) -> list[pd.DataFrame]:
    """
    pd.read_html の安全版。表が1つも無いページでは空リストを返す。
    flavor="lxml" を明示しているのは、lxmlで表が見つからなかった場合に
    pandasが bs4+html5lib へ自動で切り替えようとし、html5lib未インストールで
    ImportError になって全体が止まるのを防ぐため。
    """
    try:
        return pd.read_html(io.StringIO(html), flavor="lxml")
    except (ValueError, ImportError):
        return []


def find_col(cols: list[str], *keywords: str) -> int | None:
    """列名リストから、いずれかのキーワードを含む最初の列番号を返す"""
    for i, c in enumerate(cols):
        if any(k in c for k in keywords):
            return i
    return None


MATCH_DATE_RE = re.compile(r"(\d{1,2})(?:/|月)(\d{1,2})")
SALE_DT_RE = re.compile(
    r"(\d{1,2})(?:/|月)(\d{1,2})日?\s*\([^)]*\)\s*(?:(\d{1,2}):(\d{2}))?"
)


def parse_sale_datetime(cell: str, default_time: str | None = None) -> str | None:
    """
    「9/19（土） 10：00〜」「7月24日(金)12:00～」等を「9/19 10:00」形式にする。
    時刻が書かれていない場合は default_time を補う(無ければ日付のみ)。
    「─」等、日付が無いセルは None。
    """
    m = SALE_DT_RE.search(nfkc(cell))
    if not m:
        return None
    md = f"{int(m.group(1))}/{int(m.group(2))}"
    if m.group(3):
        return f"{md} {int(m.group(3))}:{m.group(4)}"
    return f"{md} {default_time}" if default_time else md


def is_past_match(month: int, day: int, today: date | None = None) -> bool:
    """
    年の書かれていない「M/D」が、今日より前の試合かどうかを判定する。
    シーズンが年をまたぐため、6か月以上離れた月は前年/翌年の試合とみなす。
    """
    today = today or datetime.now(JST).date()
    year = today.year
    if month < today.month - 6:
        year += 1
    elif month > today.month + 6:
        year -= 1
    try:
        return date(year, month, day) < today
    except ValueError:
        return False


def _build_abbr_to_full() -> dict[str, str]:
    """CLUB_ABBR(フルネーム→略称)を逆引きできるようにする。「C大阪」→「セレッソ大阪」"""
    table = {}
    for full, abbr in CLUB_ABBR.items():
        if isinstance(abbr, str) and abbr.strip():
            table[nfkc(abbr)] = full
    return table


ABBR_TO_FULL = _build_abbr_to_full()


def normalize_opponent(name: str) -> str:
    """「C大阪戦」→「セレッソ大阪」。略称が変換表に無ければ「戦」を外しただけで返す"""
    n = nfkc(name)
    n = re.sub(r"^(東京ヴェルディ|名古屋グランパス)?\s*(vs\.?|VS\.?)\s*", "", n)
    n = re.sub(r"戦$", "", n).strip()
    return ABBR_TO_FULL.get(n, dedupe_name(n))


# ── 名古屋グランパス ──────────────────────────────────────────
def parse_grampus(html: str) -> list[dict]:
    """
    nagoya-grampus.jp/ticket/schedule/ のHTML表を解析する。
    列: 大会 / 節 / 日にち / K.O. / 対戦相手 / 会場 / プラチナ優先 / ファンクラブ超優先 /
        ファンクラブ優先 / 一般販売 / 駐車場販売 / ファンクラブ特典招待券 / 観戦様式
    列の並びが変わっても動くよう、列番号ではなく列名(見出し)で位置を特定する。
    """
    tables = read_tables(html)

    rows = []
    for raw_df in tables:
        for cols, df in table_header_candidates(raw_df):
            i_gen = find_col(cols, "一般販売")
            i_opp = find_col(cols, "対戦相手")
            i_date = find_col(cols, "日にち", "開催日")
            if None not in (i_gen, i_opp, i_date):
                break
        else:
            continue  # 発売スケジュール以外の表
        i_comp = find_col(cols, "大会")
        i_sec = find_col(cols, "節")
        i_venue = find_col(cols, "会場")

        for _, r in df.iterrows():
            cells = [nfkc(c) for c in r.tolist()]
            date_m = MATCH_DATE_RE.search(cells[i_date])
            general = parse_sale_datetime(cells[i_gen])
            opponent = cells[i_opp]
            if not (date_m and general and opponent and opponent.lower() != "nan"):
                continue  # 「─」等、一般販売日が未定の試合はスキップ

            comp = cells[i_comp] if i_comp is not None else ""
            sec = cells[i_sec] if i_sec is not None else ""
            section = comp
            if re.fullmatch(r"\d+", sec):
                if comp.upper() == "J1":
                    section = f"第{int(sec)}節"
                elif comp.upper() == "YBC":
                    section = f"ルヴァンカップ 第{int(sec)}節"
                else:
                    section = f"{comp} 第{int(sec)}節".strip()

            venue = cells[i_venue] if i_venue is not None else ""
            rows.append({
                "club": "名古屋グランパス",
                "section": section,
                "match_date": f"{int(date_m.group(1))}/{int(date_m.group(2))}",
                "opponent": normalize_opponent(opponent),
                "venue": "" if venue.lower() == "nan" else venue,
                "general_sale": general,
            })
    return rows


# ── 東京ヴェルディ ──────────────────────────────────────────
VERDY_BASE = "https://www.verdy.co.jp"
VERDY_LIST_URL = "https://www.verdy.co.jp/news/tag/top?p={page}"
VERDY_TITLE_KEYWORD = "チケット販売"
# U-21やベレーザ(女子)等、トップチームのJ1以外の試合の告知は除外する(NFKC正規化後の表記で判定)
VERDY_EXCLUDE_WORDS = ("U-21", "U21", "ベレーザ", "ユース", "ジュニア")
# 公式サイトの記載:「販売開始初日の販売開始時間は、会員割引・一般販売ともに12:00～」
VERDY_DEFAULT_SALE_TIME = "12:00"
# カップ戦の勝ち上がり次第でホーム開催かどうかが決まる「仮の」販売スケジュールの告知を見分ける文言。
# 例:「ホームゲーム開催となった場合のチケット販売スケジュール」「アウェイゲーム開催となった場合、販売は実施しません」
# 開催が確定すると、クラブは改めて販売概要の記事を出すので、そちらを拾えば足りる。
# (「試合が中止となった場合」等の通常の注意書きに反応しないよう、開催・勝ち上がりに関する言い回しに限定している)
VERDY_CONDITIONAL_RE = re.compile(
    r"(ホームゲーム|ホーム|アウェイゲーム|アウェイ)開催となった場合"
    r"|勝ち上がった場合"
    r"|結果をもとに、?改めてお知らせ"
)


def parse_verdy_article(html: str) -> list[dict]:
    """
    verdy.co.jp の「〇月ホームゲームチケット販売について」記事内の販売スケジュール表を解析する。
    列: (節) / 開催日(または日時・開催日時) / キックオフ / 対戦相手(または対戦カード) / 会員販売 / 一般販売
    記事によって列構成が違うため、列名で位置を特定する。
    同じ記事内の価格表にも「一般販売」列があるが、「対戦」列が無いので除外される。
    """
    tables = read_tables(html)

    rows = []
    for raw_df in tables:
        for cols, df in table_header_candidates(raw_df):
            i_gen = find_col(cols, "一般販売")
            i_opp = find_col(cols, "対戦相手", "対戦カード")
            i_date = find_col(cols, "開催日", "日時", "日にち")
            if None not in (i_gen, i_opp, i_date):
                break
        else:
            continue  # 価格表など、販売スケジュール以外の表
        i_sec = find_col(cols, "節")

        for _, r in df.iterrows():
            cells = [nfkc(c) for c in r.tolist()]
            date_m = MATCH_DATE_RE.search(cells[i_date])
            general = parse_sale_datetime(cells[i_gen], default_time=VERDY_DEFAULT_SALE_TIME)
            opponent = cells[i_opp]
            if not (date_m and general and opponent and opponent.lower() != "nan"):
                continue

            sec = cells[i_sec] if i_sec is not None else ""
            rows.append({
                "club": "東京ヴェルディ",
                "section": f"第{sec}節" if re.fullmatch(r"\d+", sec) else "",
                "match_date": f"{int(date_m.group(1))}/{int(date_m.group(2))}",
                "opponent": normalize_opponent(opponent),
                "venue": "",
                "general_sale": general,
            })
    return rows


def collect_verdy(max_pages: int = 3, max_articles: int = 8) -> list[dict]:
    """
    verdy.co.jp は販売日程ページの表がJSで後から読み込まれるため、
    ニュース一覧(ヴェルディ=トップチームのタグ)から「チケット販売」を含む記事を探して個別に読む。
    一覧・記事とも素のHTMLに内容が含まれているので、Playwrightは使わず requests で取得する。
    """
    links = []
    for page in range(1, max_pages + 1):
        list_html = fetch(VERDY_LIST_URL.format(page=page))
        if not list_html:
            continue
        soup = BeautifulSoup(list_html, "html.parser")
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if not re.search(r"/news/\d+/?$", href.split("?")[0]):
                continue
            title = nfkc(a.get_text(" ", strip=True))
            if VERDY_TITLE_KEYWORD not in title:
                continue
            if any(w in title for w in VERDY_EXCLUDE_WORDS):
                continue
            full_url = urljoin(VERDY_BASE, href)
            if full_url not in links:
                links.append(full_url)
        if len(links) >= max_articles:
            break
        time.sleep(1)
    links = links[:max_articles]  # 一覧は新しい順なので、先頭ほど新しい記事

    if DEBUG_TEXT_DUMP:
        print(f"[DEBUG] 東京ヴェルディ: リンク候補{len(links)}件: {links}")

    rows = []
    seen = set()
    for url in links:
        html = fetch(url)
        if not html:
            continue
        if VERDY_CONDITIONAL_RE.search(nfkc(text_with_img_alts(html))):
            print(f"[INFO] 東京ヴェルディ: {url} は勝ち上がり次第の仮スケジュールのためスキップします")
            time.sleep(1)
            continue

        try:
            article_rows = parse_verdy_article(html)
        except Exception as e:
            print(f"[WARN] 東京ヴェルディ: {url} の解析中にエラー({e})")
            article_rows = []
        if not article_rows and DEBUG_TEXT_DUMP:
            print(f"[DEBUG] 東京ヴェルディ: {url} から販売スケジュール表を読み取れませんでした")
            print(repr(text_with_img_alts(html)[:1500]))

        for row in article_rows:
            # 同じ試合が複数の記事(更新版など)に載っている場合は、新しい記事の内容を優先する
            key = (row["match_date"], row["opponent"])
            if key in seen:
                continue
            month, day = map(int, row["match_date"].split("/"))
            if is_past_match(month, day):
                continue  # 過去の記事に載っている、既に終わった試合は除外
            seen.add(key)
            rows.append(row)
        time.sleep(1)

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
    {
        "club": "ガンバ大阪",
        "url": "https://www.gamba-osaka.net/ticket/schedule/",
        "parser": parse_gamba,
        "expect_marker": "一般販売",
    },
    {
        "club": "セレッソ大阪",
        "url": "https://www.cerezo.jp/ticket/",
        "parser": parse_cerezo,
        "expect_marker": "一般販売",
    },
    {
        "club": "アビスパ福岡",
        "url": "https://www.avispa.co.jp/match/series/2026-27",
        "parser": parse_avispa,
        "expect_marker": "一般",
    },
    {
        "club": "浦和レッズ",
        "url": "https://www.urawa-reds.co.jp/ticket/saleperiod.html",
        "parser": parse_urawa,
        "expect_marker": "一般販売",
    },
    {
        "club": "名古屋グランパス",
        "url": "https://nagoya-grampus.jp/ticket/schedule/",
        "parser": parse_grampus,
        # 見出しが「一般<br>販売」で「一般販売」が連続した文字列にならないため、表タイトルを目印にする
        "expect_marker": "チケット販売スケジュール",
    },
    {
        "club": "川崎フロンターレ",
        "collector": collect_frontale,
    },
    {
        "club": "東京ヴェルディ",
        "collector": collect_verdy,
    },
    {
        "club": "清水エスパルス",
        "url": "https://www.s-pulse.co.jp/tickets/schedule",
        "parser": parse_spulse,
        "expect_marker": "一般販売",
        # GitHub Actions(クラウドのIP)からのアクセスが拒否される場合に、
        # Google Apps Script 経由の取得に切り替える(SALE_WATCH_PROXY_URL 設定時のみ)
        "proxy_fallback": True,
    },
    {
        "club": "京都サンガF.C.",
        "url": "https://www.sanga-fc.jp/ticket/schedule",
        "parser": parse_sanga,
        "expect_marker": "一般販売",
        "proxy_fallback": True,
    },
    {
        "club": "ヴィッセル神戸",
        "url": "https://www.vissel-kobe.co.jp/ticket/schedule/",
        "parser": parse_vissel,
        # 「チケット販売スケジュール」はメニューにもあり、DEBUG出力がメニュー部分になってしまうため、表の見出しを目印にする
        "expect_marker": "対戦相手",
    },
    # (旧メモ)清水エスパルス・京都サンガF.C.は、サイト側のボット対策により
    # GitHub Actionsからの取得が(Playwrightを使っても)できなかったため、
    # 一旦対象から外している。パーサー自体(parse_spulse / parse_sanga)は残してあるので、
    # 将来別の取得方法が見つかれば再度有効化できる。
]


DEBUG_TEXT_DUMP = os.environ.get("SALE_WATCH_DEBUG") == "1"


# Google Apps Script(GAS)で作った取得用Webアプリ。GitHub Actions のIPが拒否されるサイト向け。
# 未設定なら使わない。GitHub の Secrets に登録し、ワークフローの env で渡す。
PROXY_URL = os.environ.get("SALE_WATCH_PROXY_URL", "").strip()
PROXY_KEY = os.environ.get("SALE_WATCH_PROXY_KEY", "").strip()


def fetch_via_proxy(url: str) -> str | None:
    """GAS経由でページを取得する(Googleのサーバーから取りに行くので、IP単位の拒否を回避できることがある)"""
    if not PROXY_URL:
        return None
    try:
        resp = requests.get(PROXY_URL, params={"url": url, "key": PROXY_KEY}, timeout=60)
        resp.raise_for_status()
        text = decode_response(resp)
        if text.startswith("ERROR:"):
            print(f"[WARN] プロキシ側でエラー: {url} ({text[:200]})")
            return None
        return text
    except requests.RequestException as e:
        print(f"[WARN] プロキシ経由の取得に失敗: {url} ({e})")
        return None


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

    # 直接取得でうまくいかなかったサイトは、設定があればGAS経由で取り直す
    if source.get("proxy_fallback") and PROXY_URL:
        print(f"[INFO] {source['club']}: Google Apps Script 経由で取得を試みます")
        proxied = fetch_via_proxy(source["url"])
        if proxied and (not marker or marker in proxied):
            return proxied
        print(f"[WARN] {source['club']}: プロキシ経由でも目印「{marker}」が見つかりませんでした")

    return html  # 最終的に目印が無くても、最後に取得できた内容をそのまま返す(呼び出し側で0件になるだけ)


# ── 会場名・節の補完 ─────────────────────────────────────────
# 販売ページに会場(や節)が載っていないクラブは、Football LAB(データスタジアム運営)の
# J1日程表から、試合日をキーに補う。カップ戦・ACLは載っていないので空欄のまま。
FOOTBALL_LAB_CODES = {
    "東京ヴェルディ": "tk-v",
    "京都サンガF.C.": "kyot",
    "川崎フロンターレ": "ka-f",
}
FOOTBALL_LAB_URL = "https://www.football-lab.jp/{code}/match?year={year}"

# 略称で書かれている会場名を正式名称に揃える(載っていない略称はそのまま表示される)
VENUE_FULL_NAMES = {
    "味スタ": "味の素スタジアム",
    "MUFG国立": "MUFGスタジアム(国立競技場)",
    "U等々力": "Uvanceとどろきスタジアム by Fujitsu",
    "サンガS": "サンガスタジアム by KYOCERA",
    "パナスタ": "パナソニックスタジアム吹田",
    "吹田S": "市立吹田サッカースタジアム",
    "ノエスタ": "ノエビアスタジアム神戸",
    "御崎公園": "御崎公園球技場",  # ACLでの名称(ノエビアスタジアム神戸と同じ会場)
}


def fetch_footballlab_lookup(code: str) -> dict[str, dict]:
    """{"11/21": {"section": "第15節", "venue": "味スタ"}, ...} を返す。取得できなければ空の辞書"""
    today = datetime.now(JST).date()
    season_year = today.year if today.month >= 7 else today.year - 1  # 2026/27シーズン → 2026
    html = fetch(FOOTBALL_LAB_URL.format(code=code, year=season_year))
    if not html:
        return {}

    # pd.read_html だと「10.10」が数値の10.1になってしまうため、セルを文字列のまま読む
    soup = BeautifulSoup(html, "html.parser")
    lookup = {}
    for tr in soup.find_all("tr"):
        cells = [nfkc(td.get_text(" ", strip=True)) for td in tr.find_all(["td", "th"])]
        if len(cells) < 4 or not re.fullmatch(r"\d+", cells[0]):
            continue
        dm = re.fullmatch(r"(\d{1,2})\.(\d{1,2})", cells[1])
        if not dm:
            continue
        # 列: 節 / 開催日 / (曜) / 相手 / [スコア(試合後のみ)] / H or A / 会場 / ...
        venue = ""
        for i in range(3, len(cells) - 1):
            if cells[i] in ("H", "A"):
                venue = cells[i + 1]
                break
        lookup[f"{int(dm.group(1))}/{int(dm.group(2))}"] = {
            "section": f"第{int(cells[0])}節",
            "venue": venue,
        }
    return lookup


def fill_missing_from_footballlab(rows: list[dict]) -> None:
    """会場・節が空欄の行を Football LAB の日程表で補う(該当クラブのみ、1クラブ1回だけ取得)"""
    cache: dict[str, dict] = {}
    for r in rows:
        code = FOOTBALL_LAB_CODES.get(r["club"])
        if not code or (r.get("venue") and r.get("section")):
            continue
        if code not in cache:
            cache[code] = fetch_footballlab_lookup(code)
            if DEBUG_TEXT_DUMP:
                print(f"[DEBUG] {r['club']}: Football LAB の日程 {len(cache[code])}件")
        info = cache[code].get(r["match_date"])
        if not info:
            continue  # カップ戦・ACLなど、J1日程に無い試合は空欄のまま
        if not r.get("section"):
            r["section"] = info["section"]
        if not r.get("venue"):
            r["venue"] = info["venue"]


def normalize_venues(rows: list[dict]) -> None:
    """
    会場名を正式名称に揃える。メインのスプレッドシートの「スタジアム略称」シート(表記→正式名称)を優先し、
    シートが無い・見つからない会場は、コード内の VENUE_FULL_NAMES で変換する。
    """
    table = {}
    resolve = None
    try:
        from price_history import load_venue_table, resolve_venue
        gc, sh = get_gspread_client()
        if sh is not None:
            table = load_venue_table(sh)
            resolve = resolve_venue
    except Exception as e:
        print(f"[WARN] 「スタジアム略称」シートの読み込みに失敗したため、コード内の変換表を使います ({e})")

    for r in rows:
        v = nfkc(r.get("venue", "")) if r.get("venue") else ""
        if v and table and resolve:
            hit = resolve(v, table)
            if hit and hit[1]:
                r["venue"] = hit[1]
                continue
        r["venue"] = VENUE_FULL_NAMES.get(v, v)


def collect_all() -> list[dict]:
    all_rows = []
    for source in SALE_SOURCES:
        # 1クラブで想定外のエラーが起きても、他のクラブの取得とシート書き出しは続ける
        try:
            all_rows.extend(collect_one(source))
        except Exception as e:
            print(f"[WARN] {source['club']}: 取得中にエラーが発生したためスキップします ({type(e).__name__}: {e})")

    try:
        fill_missing_from_footballlab(all_rows)
    except Exception as e:
        print(f"[WARN] 会場・節の補完に失敗しました ({type(e).__name__}: {e})")
    normalize_venues(all_rows)
    return all_rows


def collect_one(source: dict) -> list[dict]:
    """1クラブ分の発売予定を取得する"""
    if "collector" in source:
        # 1ページでは完結しない(ニュース記事を複数巡回する)クラブ用
        print(f"[INFO] checking {source['club']} (collector)")
        rows = source["collector"]()
        print(f"[INFO] {source['club']}: {len(rows)}件取得")
        return rows

    print(f"[INFO] checking {source['club']} ({source['url']})")
    html = fetch_with_marker_retry(source)
    if not html:
        return []

    if DEBUG_TEXT_DUMP:
        full_text = text_with_img_alts(html)
        marker = source.get("expect_marker")
        idx = full_text.find(marker) if marker else -1
        if idx != -1:
            start = max(0, idx - 400)
            sample = full_text[start: start + 1800]
            label = f"「{marker}」の周辺"
        else:
            sample = full_text[:1500]
            label = "先頭1500文字(目印が見つからなかったため)"
        print(f"[DEBUG] ---- {source['club']} の実際のテキスト({label}) ----")
        print(repr(sample))
        print("[DEBUG] ---- ここまで ----")

    rows = source["parser"](html)
    print(f"[INFO] {source['club']}: {len(rows)}件取得")
    return rows


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


# ── チケットサイトの試合ページURLの自動取得 ─────────────────────────
# jleague-ticket.jp のクラブ別ページ(/club/{code}/)には、そのクラブの試合ページへのリンクが並んでいる。
# 発売予定の各試合について、試合日が一致するリンクを探し、試合ページのタイトルで
# 「ホームクラブ・試合日」を確かめてから採用する(アウェイ側の試合や別日の試合を取り違えないため)。
JLT_BASE = "https://www.jleague-ticket.jp"
PERFORM_LINK_RE = re.compile(r"/sales/perform/(\d+)/(\d+)")


def _key(s) -> str:
    return re.sub(r"\s+", "", nfkc(s))


def club_code_for(club: str) -> str | None:
    from monitor import J1_CLUBS
    for code, full in J1_CLUBS.items():
        if _key(full) == _key(club):
            return code
    return None


def collect_club_page_links(code: str) -> dict[str, set[str]]:
    """{試合ページURL: {リンク文字列に含まれる 'M/D', ...}} を返す"""
    html = fetch(f"{JLT_BASE}/club/{code}/")
    if not html:
        return {}
    soup = BeautifulSoup(html, "html.parser")
    links: dict[str, set[str]] = {}
    for a in soup.find_all("a", href=True):
        m = PERFORM_LINK_RE.search(a["href"])
        if not m:
            continue
        url = f"{JLT_BASE}/sales/perform/{m.group(1)}/{m.group(2)}"
        text = nfkc(a.get_text(" ", strip=True))
        dm = re.search(r"(\d{1,2})/(\d{1,2})", text)
        if dm:
            links.setdefault(url, set()).add(f"{int(dm.group(1))}/{int(dm.group(2))}")
    return links


def link_ticket_urls(rows: list[dict], clubs: set[str] | None = None) -> None:
    """
    発売予定の各行に、チケットサイトの試合ページURL(ticket_url)を付ける。
    clubs を渡した場合は、そのクラブ(照合キー)だけを対象にする(価格履歴を記録するクラブに絞るため)。
    """
    from monitor import extract_match_meta

    for r in rows:
        r.setdefault("ticket_url", "")

    by_club: dict[str, list[dict]] = {}
    for r in rows:
        by_club.setdefault(r["club"], []).append(r)

    for club, club_rows in by_club.items():
        if clubs is not None and _key(club) not in clubs:
            continue
        code = club_code_for(club)
        if not code:
            continue
        links = collect_club_page_links(code)
        need = {r["match_date"]: r for r in club_rows}
        found = 0
        for url, mds in links.items():
            if not any(md in need and not need[md]["ticket_url"] for md in mds):
                continue
            page = fetch(url)
            time.sleep(1)
            if not page:
                continue
            meta = extract_match_meta(page, url)
            home = nfkc(meta.get("raw_card", "")).split("対")[0].strip()
            dm = re.search(r"\d{4}/(\d{1,2})/(\d{1,2})", meta.get("match_date", ""))
            if not dm or _key(home) != _key(club):
                continue  # アウェイ側の試合など
            md = f"{int(dm.group(1))}/{int(dm.group(2))}"
            if md in need and not need[md]["ticket_url"]:
                need[md]["ticket_url"] = url
                found += 1
        print(f"[INFO] {club}: チケットサイトの試合ページを{found}件見つけました(リンク候補{len(links)}件)")
        time.sleep(2)


def history_clubs() -> set[str] | None:
    """「履歴ファイル」シートに登録されているクラブ(照合キー)。読めなければ None(=全クラブ対象)"""
    try:
        from price_history import load_history_files
        gc, sh = get_gspread_client()
        if sh is None:
            return None
        files = load_history_files(sh)
        return {club for (_season, club) in files} or None
    except Exception as e:
        print(f"[WARN] 「履歴ファイル」シートを読めませんでした ({e})")
        return None


def main():
    rows = collect_all()
    try:
        link_ticket_urls(rows, history_clubs())
    except Exception as e:
        print(f"[WARN] 試合ページURLの取得に失敗しました ({type(e).__name__}: {e})")
        for r in rows:
            r.setdefault("ticket_url", "")
    # 見つけた試合ページURLは「発売予定」シートの ticket_url 列に書くだけで、「対象試合」シートには入れない。
    # 価格の取得とクラブ別ファイルへの記録は、monitor.py(price_history.py)がこの列を読んで行う。
    export_to_sheet(rows)


if __name__ == "__main__":
    main()
