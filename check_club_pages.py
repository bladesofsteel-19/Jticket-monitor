#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
jleague-ticket.jp のクラブ別ページ(/club/{code}/)に GitHub Actions から
アクセスできるかを確かめるための、一度きりの診断スクリプト。

各クラブについて、次をログに出す:
  - requests で取得したときの HTTP ステータス、ページの長さ、試合ページへのリンク(/sales/perform/…)の数
  - (Playwright が入っていれば)実ブラウザで取得したときの同じ情報
リンクが見つかったクラブは、先頭3件のリンクとその文字列も表示する。
"""

import re
import time

import requests

from monitor import HEADERS

CLUBS = {
    "ur": "浦和レッズ",
    "vn": "東京ヴェルディ",
    "mz": "FC町田ゼルビア",
    "kf": "川崎フロンターレ",
    "ym": "横浜F・マリノス",
    "ss": "清水エスパルス",
    "ng": "名古屋グランパス",
    "ks": "京都サンガF.C.",
    "go": "ガンバ大阪",
    "co": "セレッソ大阪",
    "af": "アビスパ福岡",
}
PERFORM_RE = re.compile(r'href="([^"]*/sales/perform/\d+/\d+[^"]*)"[^>]*>(.*?)</a>', re.S)


def summarize(label: str, status, html: str):
    links = PERFORM_RE.findall(html or "")
    print(f"  [{label}] status={status} 長さ={len(html or '')} 試合リンク={len(links)}件")
    title = re.search(r"<title>(.*?)</title>", html or "", re.S)
    print(f"    title: {title.group(1).strip()[:80] if title else '(なし)'}")
    for href, text in links[:3]:
        text = re.sub(r"<[^>]+>|\s+", " ", text).strip()
        print(f"    - {href}  {text[:60]}")


def try_requests(url: str):
    try:
        resp = requests.get(url, headers=HEADERS, timeout=20)
        resp.encoding = resp.apparent_encoding or "utf-8"
        summarize("requests", resp.status_code, resp.text)
    except requests.RequestException as e:
        print(f"  [requests] エラー: {type(e).__name__}: {e}")


def try_playwright(url: str, browser):
    try:
        page = browser.new_page(user_agent=HEADERS["User-Agent"], locale="ja-JP")
        resp = page.goto(url, wait_until="networkidle", timeout=40000)
        page.wait_for_timeout(3000)
        summarize("playwright", resp.status if resp else None, page.content())
        page.close()
    except Exception as e:
        print(f"  [playwright] エラー: {type(e).__name__}: {e}")


def main():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        sync_playwright = None
        print("[INFO] playwright が無いので requests だけで確認します")

    if sync_playwright:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            for code, name in CLUBS.items():
                url = f"https://www.jleague-ticket.jp/club/{code}/"
                print(f"== {name} {url}")
                try_requests(url)
                try_playwright(url, browser)
                time.sleep(2)
            browser.close()
    else:
        for code, name in CLUBS.items():
            url = f"https://www.jleague-ticket.jp/club/{code}/"
            print(f"== {name} {url}")
            try_requests(url)
            time.sleep(2)


if __name__ == "__main__":
    main()
