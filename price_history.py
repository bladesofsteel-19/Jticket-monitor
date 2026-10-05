#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
価格履歴をクラブ別のスプレッドシートに記録するスクリプト(monitor.py とは独立して単体で実行する)。
    実行: python price_history.py

【仕組み】
- 記録する試合は「発売予定」シートの ticket_url 列(sale_watch.py が自動で見つけた試合ページURL)だけ。
  「対象試合」シート(手動入力)とメインのスプレッドシートの試合別シートは monitor.py の担当で、
  このスクリプトは読みも書きもしない(手動と自動は完全に別系統)。
- 試合ページのタイトルからホームクラブと試合日を読み取り、「発売予定」シートの行と照合する。
- 一般発売日〜試合日の期間内の試合だけ、「履歴ファイル」シートで指定したクラブ別ファイルに記録する。
  (「履歴ファイル」にクラブの行が無ければ記録しない。神戸はこれで対象外になる)
- 1試合=1シート。シート名は「MMDD 相手略称_スタジアム略称」(例: 1021 C大阪_豊田ス)。
- 1行目は見出し(A列「記録日時」、B列以降が席種)。席種の列はチケットサイトの表示順で並べ、
  途中で増えた席種は、サイト上で1つ前にある席種の右隣に列を挿入する。既存の列は動かさない。
- 前回の記録から価格か販売状況が変わったときだけ1行追記する(容量節約と、変化の時点を見やすくするため)。

【メインのスプレッドシートに必要なシート】
- 発売予定     : sale_watch.py が書き出す(club / match_date / opponent / venue / general_sale 列を使う)
- 履歴ファイル : 見出し「シーズン」「クラブ名」「スプレッドシートURL」
- スタジアム略称: 見出し「表記」「略称」「正式名称」
"""

import re
import time
import unicodedata
from datetime import date, datetime

from monitor import (
    JST,
    CLUB_ABBR,
    get_gspread_client,
    retry_on_transient_error,
)

SALE_SHEET_NAME = "発売予定"
HISTORY_FILES_SHEET_NAME = "履歴ファイル"
VENUE_SHEET_NAME = "スタジアム略称"

HEADER_FIRST = "記録日時"
NEW_SHEET_ROWS = 100
NEW_SHEET_COLS = 30
DEFAULT_SHEET_TITLES = {"シート1", "Sheet1"}


# ── 共通の小道具 ─────────────────────────────────────────────
def nfkc(s) -> str:
    return unicodedata.normalize("NFKC", str(s or "")).strip()


def norm_key(s) -> str:
    """照合用のキー。全角/半角・括弧の全角/半角・空白の有無を揃える"""
    return re.sub(r"\s+", "", nfkc(s))


def season_label(d: date) -> str:
    """2026/27シーズン(2026年7月〜2027年6月)の試合 → '26-27'"""
    start = d.year if d.month >= 7 else d.year - 1
    return f"{start % 100:02d}-{(start + 1) % 100:02d}"


def parse_full_date(text: str) -> date | None:
    m = re.search(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", str(text))
    if not m:
        return None
    try:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def parse_md_near(text: str, ref: date) -> date | None:
    """
    年の無い「M/D」を、基準日(ref=試合日)の前後半年の範囲で解釈する。
    発売日は試合日より前なので、試合日より後になってしまう場合は前年とみなす。
    """
    # スプレッドシートが日付として解釈して「2026/10/30」の形になっていても読めるよう、年の部分を外す
    m = re.search(r"(\d{1,2})/(\d{1,2})", re.sub(r"\d{4}[/-]", "", str(text)))
    if not m:
        return None
    month, day = int(m.group(1)), int(m.group(2))
    for year in (ref.year, ref.year - 1):
        try:
            d = date(year, month, day)
        except ValueError:
            continue
        if d <= ref:
            return d
    return None


# ── メインのスプレッドシートから設定を読む ─────────────────────
def load_records(sh, title: str) -> list[dict]:
    import gspread
    try:
        ws = retry_on_transient_error(sh.worksheet, title)
    except gspread.WorksheetNotFound:
        print(f"[WARN] 「{title}」シートが見つかりません")
        return []
    return retry_on_transient_error(ws.get_all_records)


def load_history_files(sh) -> dict[tuple[str, str], str]:
    """{(シーズン, クラブ名の照合キー): スプレッドシートURL}"""
    result = {}
    for r in load_records(sh, HISTORY_FILES_SHEET_NAME):
        season = nfkc(r.get("シーズン", ""))
        club = norm_key(r.get("クラブ名", ""))
        url = str(r.get("スプレッドシートURL", "")).strip()
        if season and club and url.startswith("http"):
            result[(season, club)] = url
    return result


def load_venue_table(sh) -> dict[str, tuple[str, str]]:
    """{表記の照合キー: (略称, 正式名称)}"""
    table = {}
    for r in load_records(sh, VENUE_SHEET_NAME):
        key = norm_key(r.get("表記", ""))
        short = nfkc(r.get("略称", ""))
        official = nfkc(r.get("正式名称", ""))
        if key and (short or official):
            table[key] = (short, official)
    return table


def resolve_venue(name: str, table: dict[str, tuple[str, str]]) -> tuple[str, str] | None:
    """
    会場名から (略称, 正式名称) を探す。
    完全一致が無ければ、表記が会場名に含まれるものを探し、複数あれば最も長い表記を優先する。
    見つからなければ None。
    """
    key = norm_key(name)
    if not key:
        return None
    if key in table:
        return table[key]
    hits = [k for k in table if k and k in key]
    if hits:
        return table[max(hits, key=len)]
    return None


def load_sale_rows(sh) -> dict[tuple[str, str], dict]:
    """「発売予定」シートを {(クラブ名の照合キー, 'M/D'): 行} にする"""
    result = {}
    for r in load_records(sh, SALE_SHEET_NAME):
        club = norm_key(r.get("club", ""))
        # 「10/21」のほか、スプレッドシートが日付に変換した「2026/10/21」の形でも読めるようにする
        m = re.search(r"(\d{1,2})/(\d{1,2})", re.sub(r"\d{4}[/-]", "", nfkc(r.get("match_date", ""))))
        if club and m:
            result[(club, f"{int(m.group(1))}/{int(m.group(2))}")] = r
    return result


# ── 試合ごとのシート ─────────────────────────────────────────
def short_club(name: str, abbr_map: dict) -> str:
    """クラブのフルネームを略称に(変換表に無ければそのまま)"""
    mapping = dict(CLUB_ABBR)
    mapping.update(abbr_map or {})
    key = norm_key(name)
    for full, abbr in mapping.items():
        if norm_key(full) == key:
            return abbr
    return nfkc(name)


def sheet_prefix(match_day: date, opponent_short: str) -> str:
    return f"{match_day.month:02d}{match_day.day:02d} {opponent_short}"


def sheet_sort_key(title: str):
    """シーズン内の日付順(7月始まり)。試合シート以外は先頭に置く"""
    m = re.match(r"^(\d{2})(\d{2}) ", title)
    if not m:
        return (0, 0, 0)
    month, day = int(m.group(1)), int(m.group(2))
    return (1, month + 12 if month < 7 else month, day)


def seat_display(seat: dict) -> str:
    """セルに入れる値。価格(変動制で幅があれば「最低～最高」)+ 完売なら「 完売」"""
    try:
        pmin, pmax = int(seat["price_min"]), int(seat["price_max"])
        price = f"{pmax:,}" if pmin == pmax else f"{pmin:,}～{pmax:,}"
    except (KeyError, TypeError, ValueError):
        price = ""
    if seat.get("status") == "完売":
        return f"{price} 完売".strip()
    return price


def merge_columns(headers: list[str], page_order: list[str]) -> list[tuple[int, str]]:
    """
    既存の見出し(A列を除く席種)に無い席種を、サイト上の並び順に従って挿入する位置を決める。
    戻り値: [(挿入後の見出しリスト上の位置, 席種名), ...](挿入はこの順に行う)
    - サイト上で1つ前にある既存の席種の右隣に入れる
    - 前に既存の席種が無ければ、サイト上で後ろにある既存の席種の左隣に入れる
    - どちらも無ければ右端
    """
    current = list(headers)
    inserts = []
    for i, seat in enumerate(page_order):
        if seat in current:
            continue
        pos = None
        for prev in reversed(page_order[:i]):
            if prev in current:
                pos = current.index(prev) + 1
                break
        if pos is None:
            for nxt in page_order[i + 1:]:
                if nxt in current:
                    pos = current.index(nxt)
                    break
        if pos is None:
            pos = len(current)
        current.insert(pos, seat)
        inserts.append((pos, seat))
    return inserts


def build_updated_table(values: list[list[str]], seats: list[dict], checked_at: str) -> list[list[str]] | None:
    """
    シートの現在の内容(values)に今回の結果を反映した表を返す。
    前回の行から変化が無ければ None(書き込み不要)。
    """
    page_order = []
    current_values = {}
    for s in seats:
        name = s["seat_type"]
        if name not in current_values:
            page_order.append(name)
        current_values[name] = seat_display(s)

    if values and values[0] and values[0][0] == HEADER_FIRST:
        header = list(values[0])
        data = [list(r) for r in values[1:] if any(c.strip() for c in r)]
    else:
        header = [HEADER_FIRST]
        data = []

    width = len(header)
    data = [r + [""] * (width - len(r)) for r in data]

    seat_headers = header[1:]
    for pos, seat in merge_columns(seat_headers, page_order):
        seat_headers.insert(pos, seat)
        for r in data:
            r.insert(pos + 1, "")  # A列(記録日時)のぶん+1
    header = [HEADER_FIRST] + seat_headers

    new_row = [checked_at] + [current_values.get(h, "") for h in seat_headers]

    if data:
        last = data[-1]
        if last[1:] == new_row[1:]:
            return None

    return [header] + data + [new_row]


def write_match_sheet(sh, name_to_ws: dict, title_prefix: str, title: str,
                      seats: list[dict], checked_at: str) -> tuple[bool, bool]:
    """
    1試合分を書き込む。戻り値: (シートを新規作成したか, 行を追記したか)
    既存シートはシート名の先頭(「MMDD 相手」)で探すので、後から会場名が変わっても同じシートに書く。
    """
    ws = None
    for t, w in name_to_ws.items():
        if t == title_prefix or t.startswith(title_prefix + "_"):
            ws = w
            break

    created = False
    if ws is None:
        ws = retry_on_transient_error(sh.add_worksheet, title=title, rows=NEW_SHEET_ROWS, cols=NEW_SHEET_COLS)
        name_to_ws[title] = ws
        created = True
        values = []
    else:
        values = retry_on_transient_error(ws.get_all_values)

    table = build_updated_table(values, seats, checked_at)
    if table is None:
        return created, False

    need_rows, need_cols = len(table), max(len(r) for r in table)
    if need_rows > ws.row_count or need_cols > ws.col_count:
        retry_on_transient_error(
            ws.resize,
            rows=max(ws.row_count, need_rows + 50),
            cols=max(ws.col_count, need_cols + 5),
        )
    retry_on_transient_error(ws.update, values=table, range_name="A1")
    return created, True


def tidy_spreadsheet(sh):
    """試合シートを日付順に並べ、空の初期シート(「シート1」)を消す"""
    all_ws = retry_on_transient_error(sh.worksheets)
    match_ws = [w for w in all_ws if sheet_sort_key(w.title)[0] == 1]
    if match_ws:
        for w in all_ws:
            if w.title in DEFAULT_SHEET_TITLES:
                vals = retry_on_transient_error(w.get_all_values)
                if not any(any(c.strip() for c in r) for r in vals):
                    retry_on_transient_error(sh.del_worksheet, w)
        all_ws = retry_on_transient_error(sh.worksheets)
    ordered = sorted(all_ws, key=lambda w: sheet_sort_key(w.title))
    if [w.id for w in ordered] != [w.id for w in all_ws]:
        retry_on_transient_error(sh.reorder_worksheets, ordered)


# ── 入口 ───────────────────────────────────────────────────
def record_price_history(all_rows: list[dict], abbr_map: dict | None = None):
    """
    monitor.py の check_match() の結果(全試合分)を受け取り、対象の試合だけクラブ別ファイルに記録する。
    """
    if not all_rows:
        return
    gc, sh = get_gspread_client()
    if gc is None or sh is None:
        print("[INFO] Google Sheets未設定のため価格履歴の記録をスキップします")
        return

    history_files = load_history_files(sh)
    if not history_files:
        print(f"[INFO] 「{HISTORY_FILES_SHEET_NAME}」シートに登録が無いため価格履歴の記録をスキップします")
        return
    venue_table = load_venue_table(sh)
    sale_rows = load_sale_rows(sh)
    today = datetime.now(JST).date()

    # 試合(perform_id)ごとにまとめる
    by_match: dict[str, list[dict]] = {}
    for r in all_rows:
        by_match.setdefault(str(r.get("perform_id") or r.get("url")), []).append(r)

    opened: dict[str, tuple] = {}   # URL → (spreadsheet, {title: worksheet})
    touched_files = set()
    unknown_venues = set()
    written = skipped = 0

    for key, rows in by_match.items():
        head = rows[0]
        match_day = parse_full_date(head.get("match_date", ""))
        raw_card = nfkc(head.get("raw_card", ""))
        if not match_day or "対" not in raw_card:
            print(f"[WARN] 価格履歴: 試合日かカードが読み取れないためスキップ ({head.get('url')})")
            continue
        home, away = [x.strip() for x in raw_card.split("対", 1)]

        url = history_files.get((season_label(match_day), norm_key(home)))
        if not url:
            continue  # 履歴ファイル未登録のクラブ(神戸など)やアウェイ側のクラブ

        md = f"{match_day.month}/{match_day.day}"
        sale = sale_rows.get((norm_key(home), md))

        # 記録期間: 一般発売日〜試合日(発売予定シートに行が無い試合は、試合日までずっと記録する)
        if today > match_day:
            continue
        if sale:
            sale_day = parse_md_near(sale.get("general_sale", ""), match_day)
            if sale_day and today < sale_day:
                skipped += 1
                continue

        opponent = sale.get("opponent") if sale else away
        opp_short = short_club(opponent, abbr_map or {})
        prefix = sheet_prefix(match_day, opp_short)

        venue_name = nfkc(sale.get("venue", "")) if sale else ""
        venue_short = ""
        if venue_name:
            hit = resolve_venue(venue_name, venue_table)
            if hit and hit[0]:
                venue_short = hit[0]
            else:
                venue_short = venue_name
                unknown_venues.add(venue_name)
        title = f"{prefix}_{venue_short}" if venue_short else prefix

        seats = [r for r in rows if r.get("seat_type")]
        if not seats:
            print(f"[WARN] 価格履歴: 席種が取得できなかったためスキップ ({title})")
            continue

        try:
            if url not in opened:
                book = retry_on_transient_error(gc.open_by_url, url)
                ws_list = retry_on_transient_error(book.worksheets)
                opened[url] = (book, {w.title: w for w in ws_list})
            book, name_to_ws = opened[url]
            created, appended = write_match_sheet(
                book, name_to_ws, prefix, title, seats, head.get("checked_at", "")
            )
            if created or appended:
                touched_files.add(url)
            if appended:
                written += 1
                print(f"[INFO] 価格履歴: {home} {title} に追記しました")
        except Exception as e:
            print(f"[WARN] 価格履歴: {home} {title} の書き込みに失敗しました ({type(e).__name__}: {e})")
        time.sleep(1.2)  # 1分あたりの読み書き回数の上限に当たらないよう間隔を空ける

    for url in touched_files:
        try:
            tidy_spreadsheet(opened[url][0])
        except Exception as e:
            print(f"[WARN] 価格履歴: シートの並び替えに失敗しました ({e})")

    for v in sorted(unknown_venues):
        print(f"[WARN] スタジアム略称が未登録: {v}(「{VENUE_SHEET_NAME}」シートに追加すると次回から略称になります)")
    print(f"[INFO] 価格履歴: {written}試合を追記しました(発売前でスキップ: {skipped}試合)")


def infer_upcoming_date(md_text: str, today: date) -> date | None:
    """年の無い「M/D」を、今日以降の直近の日付として解釈する(過ぎた試合は None)"""
    m = re.search(r"(\d{1,2})/(\d{1,2})", re.sub(r"\d{4}[/-]", "", nfkc(md_text)))
    if not m:
        return None
    month, day = int(m.group(1)), int(m.group(2))
    for year in (today.year, today.year + 1):
        try:
            d = date(year, month, day)
        except ValueError:
            continue
        if d >= today and (d - today).days <= 330:
            return d
    return None


def collect_auto_rows(sh) -> list[dict]:
    """
    「発売予定」シートの ticket_url 列にある試合のうち、記録期間内(一般発売日〜試合日)で、
    ホームクラブが履歴ファイルに登録されているものだけを取得する。
    """
    from monitor import check_match

    history_files = load_history_files(sh)
    today = datetime.now(JST).date()
    rows: list[dict] = []
    seen = set()
    fetched = 0
    for rec in load_records(sh, SALE_SHEET_NAME):
        url = str(rec.get("ticket_url", "")).strip()
        if not url.startswith("http") or url in seen:
            continue
        seen.add(url)
        match_day = infer_upcoming_date(rec.get("match_date", ""), today)
        if not match_day:
            continue  # 終わった試合
        if (season_label(match_day), norm_key(rec.get("club", ""))) not in history_files:
            continue  # 履歴ファイル未登録のクラブ
        sale_day = parse_md_near(rec.get("general_sale", ""), match_day)
        if sale_day and today < sale_day:
            continue  # 一般発売前(発売日になったら取得を始める)
        result = check_match(url)
        fetched += 1
        if not result:
            print(f"[WARN] 価格履歴: 試合ページを取得できませんでした ({url})")
        rows.extend(result)
        time.sleep(1.5)  # サイト負荷軽減
    print(f"[INFO] 価格履歴: 発売予定シートのURLから{fetched}試合を取得しました")
    return rows


def main():
    from monitor import load_club_abbr_map

    gc, sh = get_gspread_client()
    if gc is None or sh is None:
        print("[INFO] Google Sheets未設定のため価格履歴の記録をスキップします")
        return
    rows = collect_auto_rows(sh)
    abbr_map = load_club_abbr_map()  # 相手チームの略称(「チーム名」シート)
    record_price_history(rows, abbr_map)


if __name__ == "__main__":
    main()
