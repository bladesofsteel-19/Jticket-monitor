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
- 取得するたびに1日1行追記する(価格が前日と同じでも記録する)。同じ日に2回以上実行した場合は、その日の行を上書きする。

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

HEADER_FIRST = "記録日時"        # 旧形式(移行用)
SALE_LABEL_PREFIX = "発売日："   # 新形式のA1セル
HOLIDAY_SHEET_NAME = "祝日"      # 条件付き書式(日曜・祝日=赤)が参照する、非表示の祝日一覧
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
    """
    セルに入れる値(メインのスプレッドシートと同じ表記)。
      大人価格(一番高い価格)、変動価格なら「 [変動]」、完売なら「 完売」を付ける。
      例: 「6,500」「4,800 [変動]」「3,600 完売」
    """
    try:
        price = f"{int(seat['price_max']):,}"
    except (KeyError, TypeError, ValueError):
        price = ""
    if str(seat.get("dynamic")).lower() in ("true", "1"):
        price += " [変動]"
    if seat.get("status") == "完売":
        price += " 完売"
    return price.strip()


def to_serial(d: date) -> int:
    """日付をスプレッドシートの日付シリアル値に(1899/12/30 起点)"""
    return (d - date(1899, 12, 30)).days


def from_cell_date(v, match_day: date) -> int | None:
    """
    A列の値をシリアル値に。表示値の「10/05」(年なし)は、試合日以前の直近の日付として年を補う。
    旧形式の「2026-10-05 06:00」やシリアル値にも対応。
    """
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return int(v)
    d = parse_full_date(str(v)) or parse_md_near(str(v), match_day)
    return to_serial(d) if d else None


PRICE_DISPLAY_RE = re.compile(r"^([\d,]+)\s*(.*)$")


def price_cell(display: str) -> dict:
    """
    「4,000 [変動] 完売」のような表示用の文字列を、数値+表示形式のセルにする。
    セルの中身は数値(4000)のまま、表示形式で「4,000 [変動] 完売」と見せるので、
    右寄せ・カンマ区切りになり、計算や条件付き書式でもそのまま数値として扱える。
    """
    m = PRICE_DISPLAY_RE.match(display or "")
    if not m:
        return {"userEnteredValue": {"stringValue": display}} if display else {}
    number = int(m.group(1).replace(",", ""))
    suffix = m.group(2).strip().replace('"', "")
    pattern = "#,##0" + (f'" {suffix}"' if suffix else "")
    return {"userEnteredValue": {"numberValue": number},
            "userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": pattern}}}


def table_to_update_request(sheet_id: int, table: list[list]) -> dict:
    """表全体を1回で書き込むリクエスト(1行目=文字、A列=日付、価格=数値+表示形式)"""
    width = max(len(r) for r in table)
    rows = []
    for ri, r in enumerate(table):
        cells = []
        for ci in range(width):
            v = r[ci] if ci < len(r) else ""
            if ri == 0:
                cells.append({"userEnteredValue": {"stringValue": str(v)}} if v != "" else {})
            elif ci == 0 and isinstance(v, int):
                cells.append({"userEnteredValue": {"numberValue": v},
                              "userEnteredFormat": {"numberFormat": {"type": "DATE", "pattern": "MM/dd"}}})
            else:
                cells.append(price_cell(str(v)))
        rows.append({"values": cells})
    return {"updateCells": {
        "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": len(table),
                  "startColumnIndex": 0, "endColumnIndex": width},
        "rows": rows,
        "fields": "userEnteredValue,userEnteredFormat.numberFormat"}}


def sale_label(sale_day: date | None) -> str:
    return f"{SALE_LABEL_PREFIX}{sale_day:%m/%d}" if sale_day else f"{SALE_LABEL_PREFIX}不明"


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


def build_updated_table(values: list[list], seats: list[dict], today: date, label: str,
                        match_day: date) -> tuple[list[list], bool] | tuple[None, bool]:
    """
    シートの現在の内容(values: 画面に表示されている値)に今回の結果を反映した表を返す。
    価格は表示の文字列(「4,000 [変動]」)のまま比べ、書き込むときに数値+表示形式に変換する。
    戻り値: (新しい表, 旧形式から移行したか)。書き込みが不要なら (None, False)。
    価格が前日と同じでも、毎日1行追加する。同じ日に2回以上実行した場合は、その日の行を上書きする。
    表の形:
      1行目 = [「発売日：MM/DD」, 席種1, 席種2, ...]
      2行目〜 = [日付(シリアル値), 価格, 価格, ...]
    同じ日に2回実行された場合は、その日の行を上書きする(1日1行)。
    """
    page_order = []
    current_values = {}
    for s_ in seats:
        name = s_["seat_type"]
        if name not in current_values:
            page_order.append(name)
        current_values[name] = seat_display(s_)

    migrated = False
    first = str(values[0][0]).strip() if values and values[0] else ""
    if first == HEADER_FIRST or first.startswith(SALE_LABEL_PREFIX):
        header = [str(c).strip() for c in values[0]]
        raw_data = values[1:]
        migrated = first == HEADER_FIRST
    else:
        header = [""]
        raw_data = []

    width = len(header)
    data = []
    for r in raw_data:
        r = list(r) + [""] * (width - len(r))
        if not any(str(c).strip() for c in r):
            continue
        serial = from_cell_date(r[0], match_day)
        cells = [str(c).strip() for c in r[1:width]]
        if migrated:
            # 旧形式の「6,800～7,800」は最高価格だけにする
            cells = [c.split("～")[-1].strip() if "～" in c else c for c in cells]
        data.append([serial if serial is not None else r[0]] + cells)

    seat_headers = header[1:]
    for pos, seat in merge_columns(seat_headers, page_order):
        seat_headers.insert(pos, seat)
        for r in data:
            r.insert(pos + 1, "")  # A列のぶん+1
    header = [label] + seat_headers

    today_serial = to_serial(today)
    new_row = [today_serial] + [current_values.get(h, "") for h in seat_headers]

    if data and data[-1][0] == today_serial:
        # 同じ日の2回目以降は、その日の行を上書きする
        if (not migrated and data[-1][1:] == new_row[1:]
                and header == [str(c).strip() for c in values[0]]):
            return None, False  # 同じ日・同じ内容なら書き込み不要
        data[-1] = new_row
        return [header] + data, migrated

    return [header] + data + [new_row], migrated


def hex_color(h: str) -> dict:
    h = h.lstrip("#")
    return {"red": int(h[0:2], 16) / 255, "green": int(h[2:4], 16) / 255, "blue": int(h[4:6], 16) / 255}


def price_diff_formula(op: str) -> str:
    """1つ上の行(前回の記録)との価格差で判定する式。価格は数値なので、そのまま引き算で比べる"""
    return f"=AND(ISNUMBER(B3),ISNUMBER(B2),{op.format(d='(B3-B2)')})"


def formatting_requests(sheet_id: int) -> list[dict]:
    """日付の表示形式・色分けと、価格の上下の色分けを設定するリクエスト(シート作成時・移行時に1回だけ)"""
    date_range = {"sheetId": sheet_id, "startRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 1}  # A2:A
    price_range = {"sheetId": sheet_id, "startRowIndex": 2, "startColumnIndex": 1}                   # B3:最後

    def rule(rng, formula, text=None, bg=None, bold=False):
        fmt = {}
        if text:
            fmt["textFormat"] = {"foregroundColor": hex_color(text), "bold": bold}
        if bg:
            fmt["backgroundColor"] = hex_color(bg)
        return {"ranges": [rng], "booleanRule": {
            "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": formula}]},
            "format": fmt}}

    rules = [
        # 日付: 日曜・祝日=赤、土曜=青の背景色(祝日の土曜は赤を優先)
        rule(date_range, f'=AND($A2<>"",OR(WEEKDAY($A2)=1,COUNTIF(INDIRECT("{HOLIDAY_SHEET_NAME}!A:A"),$A2)>0))',
             bg="#F8CBAD"),
        rule(date_range, '=AND($A2<>"",WEEKDAY($A2)=7)', bg="#BDD7EE"),
        # 価格: 前回より500円以上上がった=濃い赤、500円未満上がった=薄い赤、下がった場合は青
        rule(price_range, price_diff_formula("{d}>=500"), text="#FFFFFF", bg="#C00000", bold=True),
        rule(price_range, price_diff_formula("AND({d}>0,{d}<500)"), bg="#F4CCCC"),
        rule(price_range, price_diff_formula("{d}<=-500"), text="#FFFFFF", bg="#1155CC", bold=True),
        rule(price_range, price_diff_formula("AND({d}<0,{d}>-500)"), bg="#CFE2F3"),
    ]
    reqs = [{"repeatCell": {
        "range": date_range,
        "cell": {"userEnteredFormat": {"numberFormat": {"type": "DATE", "pattern": "MM/dd"}}},
        "fields": "userEnteredFormat.numberFormat"}}]
    reqs += [{"addConditionalFormatRule": {"rule": r, "index": i}} for i, r in enumerate(rules)]
    return reqs


def clear_conditional_formats_requests(sh, sheet_id: int) -> list[dict]:
    """移行時に、既存の条件付き書式を消すリクエスト(重複して増えないように)"""
    meta = retry_on_transient_error(sh.fetch_sheet_metadata)
    for sheet in meta.get("sheets", []):
        if sheet["properties"]["sheetId"] == sheet_id:
            n = len(sheet.get("conditionalFormats", []))
            return [{"deleteConditionalFormatRule": {"sheetId": sheet_id, "index": 0}} for _ in range(n)]
    return []


def read_displayed(ws) -> list[list]:
    """画面に表示されている値で読む(価格の「[変動]」「完売」は表示形式に入っているため)"""
    return retry_on_transient_error(ws.get_all_values) or []


def write_match_sheet(sh, name_to_ws: dict, title_prefix: str, title: str,
                      seats: list[dict], today: date, label: str, match_day: date) -> tuple[bool, bool]:
    """
    1試合分を書き込む。戻り値: (シートを新規作成したか, 行を追記・更新したか)
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
        values = read_displayed(ws)

    table, migrated = build_updated_table(values, seats, today, label, match_day)
    if table is None:
        return created, False

    need_rows, need_cols = len(table), max(len(r) for r in table)
    if need_rows > ws.row_count or need_cols > ws.col_count:
        retry_on_transient_error(
            ws.resize,
            rows=max(ws.row_count, need_rows + 50),
            cols=max(ws.col_count, need_cols + 5),
        )
    if migrated:
        retry_on_transient_error(ws.clear)
    reqs = [table_to_update_request(ws.id, table)]
    if created or migrated:
        if migrated:
            reqs += clear_conditional_formats_requests(sh, ws.id)
        reqs += formatting_requests(ws.id)
    retry_on_transient_error(sh.batch_update, {"requests": reqs})
    return created, True


def _rule_signature(rule: dict) -> tuple:
    """条件付き書式の比較用(APIは0の色成分を省略して返すので、丸めて揃える)"""
    br = rule.get("booleanRule", {})
    formula = br.get("condition", {}).get("values", [{}])[0].get("userEnteredValue", "")
    fmt = br.get("format", {})

    def color(c):
        c = c or {}
        return tuple(round(c.get(k, 0), 2) for k in ("red", "green", "blue"))

    text = fmt.get("textFormat", {})
    return (formula, color(fmt.get("backgroundColor")),
            color(text.get("foregroundColor")) if "foregroundColor" in text else None,
            bool(text.get("bold", False)))


def match_day_from_title(title: str, today: date) -> date | None:
    """シート名「1021 C大阪_豊田ス」の MMDD から、今日に最も近い年の日付を求める"""
    m = re.match(r"^(\d{2})(\d{2}) ", title)
    if not m:
        return None
    cands = []
    for y in (today.year - 1, today.year, today.year + 1):
        try:
            cands.append(date(y, int(m.group(1)), int(m.group(2))))
        except ValueError:
            pass
    return min(cands, key=lambda d: abs((d - today).days)) if cands else None


def rewrite_values_requests(ws, today: date) -> list[dict]:
    """
    既存シートの値を今の形式(日付=日付、価格=数値+表示形式)で書き直すリクエスト。
    以前の版で価格を文字列として書いていたシートを、数値に変換するために使う。
    """
    values = read_displayed(ws)
    if not values or not str(values[0][0]).startswith(SALE_LABEL_PREFIX):
        return []
    match_day = match_day_from_title(ws.title, today)
    if not match_day:
        return []
    width = len(values[0])
    table = [[str(c) for c in values[0]]]
    for r in values[1:]:
        r = list(r) + [""] * (width - len(r))
        if not any(str(c).strip() for c in r):
            continue
        serial = from_cell_date(r[0], match_day)
        table.append([serial if serial is not None else r[0]] + [str(c).strip() for c in r[1:width]])
    return [table_to_update_request(ws.id, table)]


def refresh_formatting(sh, name_to_ws: dict, today: date):
    """
    ファイル内の試合シートの条件付き書式が今の設定と違っていれば付け直す
    (色や判定方法を変えたときに、既存のシートにも反映させるため)。メタデータの取得は1ファイル1回。
    付け直すシートは、値も今の形式で書き直す(価格の文字列→数値の変換もここで行われる)。
    """
    meta = retry_on_transient_error(sh.fetch_sheet_metadata)
    ws_by_id = {w.id: w for w in name_to_ws.values()}
    reqs = []
    fixed = 0
    for sheet in meta.get("sheets", []):
        props = sheet["properties"]
        if sheet_sort_key(props.get("title", ""))[0] != 1:
            continue
        sid = props["sheetId"]
        expected = [r["addConditionalFormatRule"]["rule"] for r in formatting_requests(sid)
                    if "addConditionalFormatRule" in r]
        current = sheet.get("conditionalFormats", [])
        if [_rule_signature(r) for r in current] == [_rule_signature(r) for r in expected]:
            continue
        if sid in ws_by_id:
            reqs += rewrite_values_requests(ws_by_id[sid], today)
        reqs += [{"deleteConditionalFormatRule": {"sheetId": sid, "index": 0}} for _ in current]
        reqs += formatting_requests(sid)
        fixed += 1
    if reqs:
        retry_on_transient_error(sh.batch_update, {"requests": reqs})
        print(f"[INFO] 価格履歴: {fixed}シートの条件付き書式を更新しました")


def japanese_holidays(years: list[int]) -> list[date]:
    try:
        import jpholiday
    except ImportError:
        print("[WARN] jpholiday が未インストールのため、祝日の色分けは行いません(土日のみ)")
        return []
    days = []
    for y in years:
        days += [d for d, _name in jpholiday.year_holidays(y)]
    return sorted(days)


def ensure_holiday_sheet(sh, holidays: list[date]):
    """条件付き書式が参照する非表示の「祝日」シートを用意する(中身は毎回最新にする)"""
    import gspread
    try:
        ws = retry_on_transient_error(sh.worksheet, HOLIDAY_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = retry_on_transient_error(sh.add_worksheet, title=HOLIDAY_SHEET_NAME, rows=100, cols=2)
        retry_on_transient_error(sh.batch_update, {"requests": [{"updateSheetProperties": {
            "properties": {"sheetId": ws.id, "hidden": True}, "fields": "hidden"}}]})
    rows = [[to_serial(d)] for d in holidays] or [[""]]
    if len(rows) > ws.row_count:
        retry_on_transient_error(ws.resize, rows=len(rows) + 10, cols=2)
    retry_on_transient_error(ws.clear)
    retry_on_transient_error(ws.update, values=rows, range_name="A1")


def tidy_spreadsheet(sh, holidays: list[date]):
    """祝日シートの更新、試合シートの日付順の並べ替え、空の初期シート(「シート1」)の削除"""
    ensure_holiday_sheet(sh, holidays)
    all_ws = retry_on_transient_error(sh.worksheets)
    match_ws = [w for w in all_ws if sheet_sort_key(w.title)[0] == 1]
    if match_ws:
        for w in all_ws:
            if w.title in DEFAULT_SHEET_TITLES:
                vals = retry_on_transient_error(w.get_all_values)
                if not any(any(str(c).strip() for c in r) for r in vals):
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
        sale_day = parse_md_near(sale.get("general_sale", ""), match_day) if sale else None
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
                book, name_to_ws, prefix, title, seats, today, sale_label(sale_day), match_day
            )
            if created or appended:
                touched_files.add(url)
            if appended:
                written += 1
                print(f"[INFO] 価格履歴: {home} {title} に追記しました")
        except Exception as e:
            print(f"[WARN] 価格履歴: {home} {title} の書き込みに失敗しました ({type(e).__name__}: {e})")
        time.sleep(1.2)  # 1分あたりの読み書き回数の上限に当たらないよう間隔を空ける

    holidays = japanese_holidays([today.year - 1, today.year, today.year + 1]) if touched_files else []
    for url in touched_files:
        try:
            tidy_spreadsheet(opened[url][0], holidays)
        except Exception as e:
            print(f"[WARN] 価格履歴: シートの並び替えに失敗しました ({e})")
    for url, (book, names) in opened.items():
        try:
            refresh_formatting(book, names, today)
        except Exception as e:
            print(f"[WARN] 価格履歴: 条件付き書式の更新に失敗しました ({e})")

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


def clear_past_ticket_urls(sh, today: date) -> None:
    """
    「発売予定」シートで、試合日を過ぎた(=試合日の翌日以降の)試合の ticket_url を空欄にする。
    sale_watch.py の次の実行(週1回)でシートは作り直されるが、それまでの間も古いURLを残さないため。
    """
    import gspread
    try:
        ws = retry_on_transient_error(sh.worksheet, SALE_SHEET_NAME)
    except gspread.WorksheetNotFound:
        return
    values = retry_on_transient_error(ws.get_all_values)
    if not values or "ticket_url" not in values[0] or "match_date" not in values[0]:
        return
    col_url = values[0].index("ticket_url")
    col_date = values[0].index("match_date")
    cells = []
    for i, row in enumerate(values[1:], start=2):
        url = row[col_url].strip() if col_url < len(row) else ""
        md = row[col_date] if col_date < len(row) else ""
        if url and infer_upcoming_date(md, today) is None:
            cells.append(gspread.Cell(row=i, col=col_url + 1, value=""))
    if cells:
        retry_on_transient_error(ws.update_cells, cells)
        print(f"[INFO] 「{SALE_SHEET_NAME}」シート: 試合日を過ぎた{len(cells)}試合のURLを消しました")


def main():
    from monitor import load_club_abbr_map

    gc, sh = get_gspread_client()
    if gc is None or sh is None:
        print("[INFO] Google Sheets未設定のため価格履歴の記録をスキップします")
        return
    try:
        clear_past_ticket_urls(sh, datetime.now(JST).date())
    except Exception as e:
        print(f"[WARN] 発売予定シートの古いURLの削除に失敗しました ({e})")
    rows = collect_auto_rows(sh)
    abbr_map = load_club_abbr_map()  # 相手チームの略称(「チーム名」シート)
    record_price_history(rows, abbr_map)


if __name__ == "__main__":
    main()
