# -*- coding: utf-8 -*-
"""
transfer_to_summary.py
最新の資産クラスシートの各セクション小計を読み取り、
資産推移纏め シートに本日付け列として追記（既存なら上書き）する

【読み取り方針】
  各セクションの「小計」行のキャッシュ値を直接使用する。
  小計が見つからない/読めない場合のみベース値で補完。

  A 円建て現預金    : 資産クラスシート row26  col6  (小計)
  B 外貨現預金      : 資産クラスシート row36  col6  (小計)
  C 円建て社債      : 資産クラスシート row40  col7  (1行のみ、直接値)
  D US$建て社債     : 資産クラスシート row54  col8  (小計)
  E Private Credit  : ベース値 × (USDJPY_新 / USDJPY_旧) で為替調整
  F 投資信託        : 資産クラスシート row68  col8  (小計)
  G 日本株          : 資産クラスシート row83  col8  (小計)
  H 米国株          : 資産クラスシート row123 col8  (小計)
  REIT              : 0 (固定)
  I 事業投資（優先株）: 資産クラスシート row134 col8 (小計)
  J 自宅不動産      : 資産クラスシート row137 col8  (直接値)
"""

import os, sys, shutil
from datetime import date, datetime

if sys.stdout.encoding and sys.stdout.encoding.lower() in ('cp932', 'shift_jis', 'shift-jis'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import openpyxl

BASE_DIR   = r'C:\Users\yna78\OneDrive - Corazon\Yuji Private\資産ポートフォリオ戦略'
EXCEL_FILE = os.path.join(BASE_DIR, '資産クラス整理_20260415.xlsx')

TS_SHEET = '資産推移纏め'

TARGET_DATE = date.today()
SN_NEW = '資産クラス' + TARGET_DATE.strftime('%Y%m%d')

print('=== transfer_to_summary.py ===')
print(f'対象日  : {TARGET_DATE}')
print(f'新シート: {SN_NEW}')

# ── ファイルをコピーして読み込み ─────────────────────────
tmp = EXCEL_FILE + '.tmp.xlsx'
shutil.copy2(EXCEL_FILE, tmp)
wb = openpyxl.load_workbook(tmp, data_only=True)

if SN_NEW not in wb.sheetnames:
    print(f'[ERROR] シート "{SN_NEW}" が見つかりません。利用可能: {wb.sheetnames}')
    os.remove(tmp)
    sys.exit(1)

ws_new = wb[SN_NEW]
ws_ts  = wb[TS_SHEET]

def fvn(r, c):
    """新シートのセル値を float で返す (None または変換不可 → None)"""
    v = ws_new.cell(row=r, column=c).value
    try:
        return float(v) if v is not None else None
    except (TypeError, ValueError):
        return None

# ── 資産推移纏め の行マップ構築 ───────────────────────────
ts_row_map = {}
for r in range(2, ws_ts.max_row + 1):
    v = ws_ts.cell(row=r, column=2).value
    if v:
        ts_row_map[str(v).strip()] = r

print(f'\n資産推移纏め 行マップ:')
for k, v in ts_row_map.items():
    print(f'  row{v}: "{k}"')

# ── 書き込み列を決定 (今日の日付列 or 末尾+1) ─────────────
target_col = None
for c in range(3, ws_ts.max_column + 2):
    cv = ws_ts.cell(row=2, column=c).value
    if cv is None:
        target_col = c
        break
    if isinstance(cv, datetime) and cv.date() == TARGET_DATE:
        target_col = c
        print(f'  既存の {TARGET_DATE} 列を上書き: col={target_col}')
        break
if target_col is None:
    target_col = ws_ts.max_column + 1

# ベース列 = 書き込み列の1つ前 (前回日付の正しいデータ)
base_col = target_col - 1
print(f'書き込み列={target_col}  ベース列={base_col}')

def base_val(label):
    r = ts_row_map.get(label)
    if r:
        v = ws_ts.cell(row=r, column=base_col).value
        return float(v) if v is not None else 0.0
    return 0.0

# ── USDJPY 取得 ────────────────────────────────────────────
# row29 col8 = 新シートの直近為替レート
USDJPY_NEW = fvn(29, 8)
if not USDJPY_NEW or USDJPY_NEW < 50:
    print(f'  [WARN] USDJPY 読み取り失敗 (値={USDJPY_NEW}), 143.0 を使用')
    USDJPY_NEW = 143.0

# 前回の USDJPY はベース列の Private Credit 行から逆算 or 固定値
# 資産推移纏めには為替レートが入っていないので、ベース値から推定
# 前回シート名から読む
SN_PREV_CANDIDATES = [s for s in wb.sheetnames if s.startswith('資産クラス') and s != SN_NEW]
USDJPY_PREV = None
if SN_PREV_CANDIDATES:
    sn_prev = sorted(SN_PREV_CANDIDATES)[-1]  # 最新の前回シート
    ws_prev = wb[sn_prev]
    v_prev = ws_prev.cell(row=29, column=8).value
    try:
        USDJPY_PREV = float(v_prev) if v_prev is not None else None
    except (TypeError, ValueError):
        pass
    print(f'前回シート: {sn_prev}  USDJPY_PREV={USDJPY_PREV}')
if not USDJPY_PREV or USDJPY_PREV < 50:
    USDJPY_PREV = 159.039
    print(f'  [WARN] USDJPY_PREV 取得失敗, {USDJPY_PREV} を使用')

print(f'USDJPY: 前回={USDJPY_PREV:.3f}  新={USDJPY_NEW:.3f}')

# ═══════════════════════════════════════════════════════════
# 各セクション 小計読み取り
# ═══════════════════════════════════════════════════════════

totals = {}

# ── A: 円建て現預金 (row26 col6 小計) ─────────────────────
sec_a = fvn(26, 6)
if sec_a is None or sec_a == 0:
    # フォールバック: 範囲スキャン
    sec_a = sum(fvn(r, 6) or 0 for r in range(6, 28) if fvn(r, 6) is not None)
if sec_a is None or sec_a == 0:
    sec_a = base_val('円建て現預金')
    print(f'A 円建て現預金: 小計読み取り失敗 → ベース値 {sec_a:,.0f}')
else:
    print(f'A 円建て現預金: {sec_a:,.0f}  (row26 col6 小計)')
totals['円建て現預金'] = sec_a

# ── B: 外貨現預金 (row36 col6 小計) ───────────────────────
sec_b = fvn(36, 6)
if sec_b is None or sec_b == 0:
    sec_b = base_val('外貨現預金')
    print(f'B 外貨現預金: 小計読み取り失敗 → ベース値 {sec_b:,.0f}')
else:
    print(f'B 外貨現預金: {sec_b:,.0f}  (row36 col6 小計)')
totals['外貨現預金'] = sec_b

# ── C: 円建て社債 (row40 col7 直接値、1行のみ) ────────────
sec_c = fvn(40, 7)
if sec_c is None or sec_c == 0:
    # row38-45 スキャン
    for r in range(38, 46):
        if str(ws_new.cell(row=r, column=3).value or '').strip() == '小計':
            break
        v = fvn(r, 7)
        if v:
            sec_c = v
            break
if sec_c is None or sec_c == 0:
    sec_c = base_val('円建て社債')
    print(f'C 円建て社債: 読み取り失敗 → ベース値 {sec_c:,.0f}')
else:
    print(f'C 円建て社債: {sec_c:,.0f}  (row40 col7)')
totals['円建て社債'] = sec_c

# ── D: US$建て社債 (row54 col8 小計) ─────────────────────
sec_d = fvn(54, 8)
# D の label キーを BASE から探す
d_label = next((k for k in ts_row_map if 'US' in k and '社債' in k), 'US＄建て社債')
if sec_d is None or sec_d == 0:
    sec_d = base_val(d_label) * (USDJPY_NEW / USDJPY_PREV)
    print(f'D US$建て社債: 小計読み取り失敗 → 為替調整 {sec_d:,.0f}')
else:
    print(f'D US$建て社債: {sec_d:,.0f}  (row54 col8 小計)')
totals[d_label] = sec_d

# ── E: Private Credit (row56 col8 小計) ───────────────────
e_label = next((k for k in ts_row_map if 'Private' in k or 'Credit' in k), 'Private Credit')
sec_e = fvn(56, 8)
if sec_e is None or sec_e == 0:
    e_base = base_val(e_label)
    sec_e = e_base * (USDJPY_NEW / USDJPY_PREV) if e_base else 0.0
    print(f'E {e_label}: {sec_e:,.0f}  (小計読み取り失敗 → ベース {e_base:,.0f} × 為替調整)')
else:
    print(f'E {e_label}: {sec_e:,.0f}  (row56 col8 小計)')
totals[e_label] = sec_e

# ── F: 投資信託 (row68 col8 小計) ─────────────────────────
sec_f = fvn(68, 8)
if sec_f is None or sec_f == 0:
    sec_f = base_val('投資信託')
    print(f'F 投資信託: 小計読み取り失敗 → ベース値 {sec_f:,.0f}')
else:
    print(f'F 投資信託: {sec_f:,.0f}  (row68 col8 小計)')
totals['投資信託'] = sec_f

# ── G: 日本株 (row83 col8 小計) ───────────────────────────
sec_g = fvn(83, 8)
if sec_g is None or sec_g == 0:
    sec_g = base_val('日本株')
    print(f'G 日本株: 小計読み取り失敗 → ベース値 {sec_g:,.0f}')
else:
    print(f'G 日本株: {sec_g:,.0f}  (row83 col8 小計)')
totals['日本株'] = sec_g

# ── H: 米国株 (row123 col8 小計) ──────────────────────────
sec_h = fvn(123, 8)
if sec_h is None or sec_h == 0:
    sec_h = base_val('米国株')
    print(f'H 米国株: 小計読み取り失敗 → ベース値 {sec_h:,.0f}')
else:
    print(f'H 米国株: {sec_h:,.0f}  (row123 col8 小計)')
totals['米国株'] = sec_h

# ── REIT: 0 固定 ───────────────────────────────────────────
reit_label = next((k for k in ts_row_map if 'REIT' in k), 'REIT')
totals[reit_label] = 0.0
print(f'REIT: 0  (固定)')

# ── I: 事業投資（優先株）(row134 col8 小計) ────────────────
sec_i = fvn(134, 8)
i_label = next((k for k in ts_row_map if '事業' in k or '優先株' in k), '事業投資（優先株）')
if sec_i is None or sec_i == 0:
    sec_i = base_val(i_label)
    print(f'I 事業投資: 小計読み取り失敗 → ベース値 {sec_i:,.0f}')
else:
    print(f'I 事業投資: {sec_i:,.0f}  (row134 col8 小計)')
totals[i_label] = sec_i

# ── J: 自宅不動産 (row137 col8 直接値) ────────────────────
sec_j = fvn(137, 8)
j_label = next((k for k in ts_row_map if '不動産' in k), '自宅不動産')
if sec_j is None or sec_j == 0:
    sec_j = base_val(j_label)
    print(f'J 自宅不動産: 読み取り失敗 → ベース値 {sec_j:,.0f}')
else:
    print(f'J 自宅不動産: {sec_j:,.0f}  (row137 col8)')
totals[j_label] = sec_j

# ── 合計計算 ─────────────────────────────────────────────
excl_keys = {'不動産', 'Total', '合計'}
total_excl = sum(v for k, v in totals.items()
                 if not any(x in k for x in excl_keys))
total_incl = total_excl + sec_j

print(f'\n=== 合計サマリー ===')
for k, v in totals.items():
    print(f'  {k}: {v:,.0f}')
print(f'  Total(不動産除く): {total_excl:,.0f}')
print(f'  Total(不動産含む): {total_incl:,.0f}')

# ═══════════════════════════════════════════════════════════
# 資産推移纏め に書き込む (data_only=False で再ロード)
# ═══════════════════════════════════════════════════════════
wb2 = openpyxl.load_workbook(tmp)   # 数式付きで再ロード
ws_ts2 = wb2[TS_SHEET]

print(f'\n書き込み: 資産推移纏め col={target_col}  ({TARGET_DATE})')

# 日付行
ws_ts2.cell(row=2, column=target_col).value = datetime(
    TARGET_DATE.year, TARGET_DATE.month, TARGET_DATE.day)

# Total 行の値マップ
total_label_excl = next((k for k in ts_row_map if 'Total' in k and ('除く' in k or 'excl' in k.lower())), None)
total_label_incl = next((k for k in ts_row_map if 'Total' in k and '含む' in k), None)

# totals に Total 行を追加
if total_label_excl:
    totals[total_label_excl] = round(total_excl)
if total_label_incl:
    totals[total_label_incl] = round(total_incl)

written = []
for r in range(2, ws_ts2.max_row + 1):
    lbl = ws_ts2.cell(row=r, column=2).value
    if not lbl:
        continue
    lbl_s = str(lbl).strip()
    if lbl_s in totals:
        val = round(totals[lbl_s])
        ws_ts2.cell(row=r, column=target_col).value = val
        written.append((r, lbl_s, val))

print(f'書き込んだ行:')
for r, lbl, v in written:
    print(f'  row{r} "{lbl}": {v:,.0f}')

# 保存
wb2.save(tmp)
try:
    os.replace(tmp, EXCEL_FILE)
    print(f'\n[OK] 保存完了: {EXCEL_FILE}')
except (PermissionError, OSError):
    bk = os.path.join(BASE_DIR, f'資産クラス整理_{TARGET_DATE.strftime("%Y%m%d")}_updated.xlsx')
    shutil.move(tmp, bk)
    print(f'\n[WARNING] 元ファイルが開かれているため別名保存: {bk}')

print('完了!')
