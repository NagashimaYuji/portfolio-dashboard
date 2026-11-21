# -*- coding: utf-8 -*-
"""
update_stock.py
毎週土曜日に実行: 株価・為替を取得してExcelを更新する

データソース:
  ひふみ投信  : Yahoo Finance Japan  https://finance.yahoo.co.jp/quote/9C31108A
  VTI/VOO/QQQ : Bloomberg Japan (403 の場合 yfinance にフォールバック)
  その他      : yfinance
"""

import os
import sys
import shutil
import re
from datetime import date

# Windows cp932 コンソール対策
if sys.stdout.encoding and sys.stdout.encoding.lower() in ('cp932', 'shift_jis', 'shift-jis'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import requests
from bs4 import BeautifulSoup
import yfinance as yf
import openpyxl

BASE_DIR = r'C:\Users\yna78\OneDrive - Corazon\Yuji Private\資産ポートフォリオ戦略'

def find_latest_excel(base_dir: str) -> str:
    """BASE_DIR 内の最新 資産クラス整理_YYYYMMDD.xlsx を返す"""
    import glob
    files = glob.glob(os.path.join(base_dir, '資産クラス整理_????????.xlsx'))
    if not files:
        raise FileNotFoundError(f'資産クラス整理_YYYYMMDD.xlsx が見つかりません: {base_dir}')
    return max(files)  # 辞書順最大 = 最新日付

def find_latest_sheet(wb) -> str:
    """直近の「資産クラスYYYYMMDD」シートを返す（今日のシートは除く）"""
    today_str = date.today().strftime('%Y%m%d')
    candidates = sorted(
        [s for s in wb.sheetnames
         if s.startswith('資産クラス') and s[5:].isdigit()
         and len(s[5:]) == 8 and s[5:] != today_str],
        reverse=True
    )
    if not candidates:
        raise RuntimeError('コピー元となる資産クラスシートが見つかりません')
    return candidates[0]

SCRAPE_HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/124.0.0.0 Safari/537.36'
    ),
    'Accept-Language': 'ja-JP,ja;q=0.9,en-US;q=0.8',
}

# ──────────────────────────────────────────────
# 銘柄 → 書き込みセルの対応は「セクション + 銘柄名 (F列) のパターン」で動的に解決する。
#   ・行を追加/削除してもシートを走査して該当行を探すため行番号の固定は不要
#   ・該当セクションで銘柄名がマッチし、G列(時価単価)が直接値の行すべてに書き込む
#     (=G87 のような参照行は自動反映されるので触らない)
#   ・銘柄名は NFKC 正規化 + 大文字化 + 空白/記号除去してからマッチ (２２５→225 等)
# ──────────────────────────────────────────────
import unicodedata

def norm_name(s):
    s = unicodedata.normalize('NFKC', str(s or '')).upper()
    return re.sub(r'[\s・．.、,（）()]', '', s)

# yfinance ティッカー: (ticker, section, 銘柄名パターン)
YF_TICKERS = [
    # 投資信託 ETF (Bloomberg 優先 → yfinance フォールバック)
    ('VTI',    'F', r'^VTI$'),
    ('VOO',    'F', r'^VOO$'),
    ('QQQ',    'F', r'^QQQ$'),
    # 日本株 (証券コードでマッチ。9024/7014 は旧シートの誤記)
    ('8001.T', 'G', r'8001'),            # 伊藤忠商事
    ('8316.T', 'G', r'8316'),            # 三井住友FG
    ('8306.T', 'G', r'8306'),            # 三菱UFJ
    ('9023.T', 'G', r'902[34]'),         # 東京地下鉄
    ('7013.T', 'G', r'701[34]'),         # IHI
    ('8303.T', 'G', r'8303'),            # SBI新生銀行
    ('2501.T', 'G', r'2501'),            # サッポロHD
    ('285A.T', 'G', r'285A'),            # キオクシア
    ('7011.T', 'G', r'7011'),            # 三菱重工
    ('446A.T', 'G', r'446A|ノースサンド'),  # ノースサンド (2026-11 追加)
    # 米国株
    ('GOOG',   'H', r'^GOOGL?$'),
    ('ISRG',   'H', r'^ISRG$'),
    ('AMZN',   'H', r'^AMZN$'),
    ('MSFT',   'H', r'^MSFT$'),
    ('NVDA',   'H', r'^NVDA$'),
    ('META',   'H', r'^META$'),
    ('GS',     'H', r'^GOLDMANSACH?S$'),  # Goldman Sachs (Sacs 誤記含む)
    ('JPM',    'H', r'^JPMORGAN$'),
    ('LLY',    'H', r'^LLY$'),
    ('PLTR',   'H', r'^PLTR$'),
    ('AAPL',   'H', r'^AA?PP?L$'),        # APPL 誤記含む
    ('TSLA',   'H', r'^TSLA$'),
    ('EW',     'H', r'^EW$'),
    ('MS',     'H', r'^MORGANSTANLEY$'),
    # PAWN は上場廃止のため除外
]
BLOOMBERG_TICKERS = {'VTI', 'VOO', 'QQQ'}

# Yahoo!ファイナンス (投信): (表示名, section, 銘柄名パターン, ファンドコード, 基準価額の除数)
#   基準価額は「円/1万口」で取得される。シートの単位に合わせて除数で換算する。
YAHOO_JP_FUNDS = [
    ('ひふみ投信',              'F', r'ひふみ',                          '9C31108A', 10000),  # シートは 円/口
    ('ストックインデックスF225', 'F', r'スト[ッツ]クインデ[ッツ]クス.*225', '04311859', 1),      # シートは 円/万口
]

def resolve_rows(ws, section, pattern):
    """
    section 内で銘柄名(F列)が pattern にマッチし、G列が数式でない行のリストを返す。
    F列が =F117 のような参照なら参照先の名前で判定する。
    """
    from person_summary import scan_sections
    sections, _ = scan_sections(ws)
    if section not in sections:
        return []
    rows = []
    for r in sections[section]['rows']:
        name = ws.cell(row=r, column=6).value
        for _ in range(4):   # 参照チェーン解決
            m = re.fullmatch(r'=\$?([A-Z]{1,2})\$?(\d+)', str(name or '').strip())
            if not m:
                break
            from openpyxl.utils import column_index_from_string
            name = ws.cell(row=int(m.group(2)), column=column_index_from_string(m.group(1))).value
        if not re.search(pattern, norm_name(name)):
            continue
        g = ws.cell(row=r, column=7).value
        if isinstance(g, str) and g.strip().startswith('='):
            continue   # 参照行 (=G87 等) は自動反映
        rows.append(r)
    return rows

BLOOMBERG_ETFS = {
    'VTI': 'https://www.bloomberg.com/jp/quote/VTI:US',
    'VOO': 'https://www.bloomberg.com/jp/quote/VOO:US',
    'QQQ': 'https://www.bloomberg.com/jp/quote/QQQ:US',
}


def get_price_yfinance(ticker: str) -> float | None:
    try:
        hist = yf.Ticker(ticker).history(period='5d')
        if not hist.empty:
            # 当日分が未確定で NaN のことがあるため NaN を除外して直近終値を取る
            closes = hist['Close'].dropna()
            if not closes.empty:
                price = round(float(closes.iloc[-1]), 4)
                if price == price and price > 0:   # NaN / 0 ガード
                    return price
            print(f'  [WARN] yfinance {ticker}: 有効な終値なし (全て NaN)')
    except Exception as e:
        print(f'  [WARN] yfinance {ticker}: {e}')
    return None


def get_yahoo_jp_fund_price(code: str, divisor: float = 1) -> float | None:
    """Yahoo!ファイナンス (投信) から基準価額 (円/1万口) を取得し divisor で割って返す"""
    url = f'https://finance.yahoo.co.jp/quote/{code}'
    try:
        r = requests.get(url, headers=SCRAPE_HEADERS, timeout=15)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, 'html.parser')
        el = soup.select_one(
            'span.StyledNumber__value__3rXW, '
            'span[class*="PriceBoard__price"], '
            'span[class*="StyledNumber--vertical"]'
        )
        if el:
            raw = el.get_text().strip().replace(',', '')
            nav_per_10k = float(raw)          # 円/1万口
            return round(nav_per_10k / divisor, 4)
    except Exception as e:
        print(f'  [WARN] Yahoo!ファイナンス {code} スクレイプ失敗: {e}')
    return None


def get_bloomberg_etf_price(ticker: str) -> float | None:
    """Bloomberg Japan からETF価格を取得。失敗時は yfinance にフォールバック"""
    url = BLOOMBERG_ETFS[ticker]
    try:
        r = requests.get(url, headers=SCRAPE_HEADERS, timeout=15)
        if r.status_code == 403:
            raise ValueError('Bloomberg 403 - フォールバック')
        r.raise_for_status()
        soup = BeautifulSoup(r.text, 'html.parser')
        # Bloomberg price selectors (layout は変わる可能性あり)
        for sel in ['[class*="priceText"]', '[class*="price__"]', '[data-field="regularMarketPrice"]']:
            el = soup.select_one(sel)
            if el:
                raw = re.sub(r'[^\d.]', '', el.get_text())
                if raw:
                    return round(float(raw), 4)
    except Exception as e:
        print(f'  [INFO] Bloomberg {ticker} ({url}): {e} → yfinance にフォールバック')
    return get_price_yfinance(ticker)


def copy_worksheet(wb, source_name: str, dest_name: str):
    if dest_name in wb.sheetnames:
        del wb[dest_name]
    ws = wb.copy_worksheet(wb[source_name])
    ws.title = dest_name
    return ws


def main():
    today        = date.today()
    today_str    = today.strftime('%Y%m%d')
    new_sheet    = f'資産クラス{today_str}'
    update_label = f'更新日: {today.strftime("%Y/%m/%d")}'

    if len(sys.argv) > 1:
        # 引数でファイル指定: そのファイルを in-place 更新 (最新ファイルが Excel でロック中のとき等)
        src_excel = out_excel = os.path.abspath(sys.argv[1])
    else:
        # ソース: 最新の 資産クラス整理_YYYYMMDD.xlsx を自動検出 / 出力: 今日付けの新ファイル
        src_excel  = find_latest_excel(BASE_DIR)
        out_excel  = os.path.join(BASE_DIR, f'資産クラス整理_{today_str}.xlsx')

    print(f'=== update_stock.py 実行日: {today} ===')
    print(f'ソースファイル: {src_excel}')
    print(f'出力ファイル  : {out_excel}')

    if not os.path.exists(src_excel):
        print(f'[ERROR] ファイルが見つかりません: {src_excel}')
        sys.exit(1)

    # 同日付きファイルが既に存在すれば警告
    if os.path.exists(out_excel) and out_excel != src_excel:
        print(f'[INFO] 既存の出力ファイルを上書きします: {out_excel}')

    tmp_path = out_excel + '.tmp.xlsx'
    shutil.copy2(src_excel, tmp_path)

    wb = openpyxl.load_workbook(tmp_path)

    from person_summary import (write_person_summary, scan_sections,
                                find_fx_cell, fix_subtotals, fill_missing_row_formulas)

    if new_sheet in wb.sheetnames:
        # 同日に再実行: 既存シートを壊さず (手で追加した行を残して) 価格だけ更新
        ws = wb[new_sheet]
        print(f'シート {new_sheet} は既に存在 → 価格のみ上書き更新 (手動追加行は保持)')
    else:
        # コピー元: 直近の日付シートを自動選択
        source_sheet = find_latest_sheet(wb)
        print(f'ソースシート  : {source_sheet}  →  {new_sheet}')
        print(f'\nシートをコピー: {source_sheet} → {new_sheet}')
        ws = copy_worksheet(wb, source_sheet, new_sheet)

    ws.cell(row=1, column=1).value = update_label
    print(f'A1 に "{update_label}" を記入')

    # 小計 SUM 範囲の自動補正 (手で行を追加したときの取りこぼし)
    for coord, old, new in fix_subtotals(ws):
        print(f'  [FIX] 小計範囲を修正: {coord}  {old} → {new}')
    # 手入力行の未入力数式 (時価総額/購入総額/含み損益/為替参照) を補完
    for coord, f in fill_missing_row_formulas(ws):
        print(f'  [FILL] {coord} = {f}')

    print('\n--- 価格取得・更新 ---')
    updated_rows = set()

    # 直近為替 (USD/JPY): '直近為替' ヘッダ直下のセルを動的に特定
    fx_row, fx_col = find_fx_cell(ws)
    fx = get_price_yfinance('USDJPY=X')
    if fx is not None:
        ws.cell(row=fx_row, column=fx_col).value = fx
        print(f'  USDJPY: {fx}  → {ws.cell(row=fx_row, column=fx_col).coordinate}')
    else:
        print('  USDJPY: 取得失敗 → 既存値を保持')

    # 投資信託 (Yahoo!ファイナンス)
    for label, sec, pat, code, divisor in YAHOO_JP_FUNDS:
        rows = resolve_rows(ws, sec, pat)
        if not rows:
            print(f'  {label}: シートに該当行なし → スキップ')
            continue
        price = get_yahoo_jp_fund_price(code, divisor)
        if price is None:
            print(f'  {label}: 取得失敗 → スキップ (要手動更新)')
            continue
        for r in rows:
            ws.cell(row=r, column=7).value = price
        updated_rows.update(rows)
        print(f'  {label} (Yahoo {code}): {price}  → G{",G".join(map(str, rows))}')

    # ETF / 株式 (Bloomberg → yfinance / yfinance)
    for ticker, sec, pat in YF_TICKERS:
        rows = resolve_rows(ws, sec, pat)
        if not rows:
            print(f'  {ticker}: シートに該当行なし → スキップ')
            continue
        price = get_bloomberg_etf_price(ticker) if ticker in BLOOMBERG_TICKERS else get_price_yfinance(ticker)
        if price is None or price != price:   # None / NaN は既存値を保持
            print(f'  {ticker}: データ取得失敗 → スキップ (既存値を保持)')
            continue
        for r in rows:
            ws.cell(row=r, column=7).value = price
        updated_rows.update(rows)
        print(f'  {ticker}: {price}  → G{",G".join(map(str, rows))}')

    # 自動更新できなかった行 (G列が直接値なのにどの価格ソースにも該当しない) を列挙
    sections, _ = scan_sections(ws)
    manual = []
    for sec in ('F', 'G', 'H'):
        for r in sections.get(sec, {}).get('rows', []):
            g = ws.cell(row=r, column=7).value
            if r in updated_rows or (isinstance(g, str) and g.strip().startswith('=')):
                continue
            if g is None or isinstance(g, (int, float)):
                manual.append(f'{ws.cell(row=r, column=6).value} (G{r})')
    if manual:
        print('\n  [要手動更新] 価格ソース未登録: ' + ' / '.join(manual))

    # 名前別 資産クラス集計表 (SUMIF 数式) を書き込み
    info = write_person_summary(ws)
    print(f'\n名前別 資産クラス集計を書き込み (row{info["table_top"]}〜)')

    # 保存 → 今日付けの新ファイルとして出力
    wb.save(tmp_path)

    try:
        os.replace(tmp_path, out_excel)
        print(f'\n[OK] 保存完了: {out_excel}')
    except (PermissionError, OSError) as e:
        fallback = os.path.join(BASE_DIR, f'資産クラス整理_{today_str}_tmp.xlsx')
        shutil.move(tmp_path, fallback)
        print(f'\n[WARNING] 保存失敗 ({e}) → 別名で保存: {fallback}')
    print(f'  新しいシート: {new_sheet}')


if __name__ == '__main__':
    main()
