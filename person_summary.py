# -*- coding: utf-8 -*-
"""
person_summary.py
資産クラスシートに「名前別 資産クラス集計」表を書き込む。

  ・資産クラス(A〜J) × 名前(Yuji / Fumi / Mina / Akatsuki) を SUMIF 数式で集計
    → Excel で開けば自動計算され、次回以降のシートコピーにも引き継がれる
  ・名前の表記ゆれは SUMIF のワイルドカードで吸収 (Akatsuki/AKATSUKI/Akasuki → "Aka*")
  ・「未分類(差)」列 = クラス小計 − 名前別合計 (0 なら集計漏れなし)

update_stock.py から write_person_summary(ws) として呼ばれる。
単体実行すると最新ファイルの最新シートに書き込み、Python で検算結果を表示する。
"""

import os, re, sys, glob
from datetime import datetime

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import column_index_from_string, get_column_letter

# (表示名, SUMIF 条件)
PERSONS = [
    ('Yuji',     'Yuji*'),
    ('Fumi',     'Fumi*'),
    ('Mina',     'Mina*'),
    ('Akatsuki', 'Aka*'),
]

# 各セクションの「時価(円)」列: A/B=F列, C=G列, D〜J=H列
VALUE_COL = {'A': 6, 'B': 6, 'C': 7}
DEFAULT_VALUE_COL = 8

# Summary 表 (row144〜) の各クラス行: 時価合計は G 列
SUMMARY_HEADER_LABEL = 'Summary'

# 名前別集計表の配置
TABLE_TOP  = 160          # タイトル行
TABLE_LEFT = 5            # E 列
TABLE_ROWS = 18           # 書き込み・クリア対象の行数
TABLE_COLS = 9            # E〜M

HDR_FILL = PatternFill('solid', fgColor='1F3864')
HDR_FONT = Font(color='FFFFFF', bold=True, size=10)
TOT_FILL = PatternFill('solid', fgColor='D9E1F2')
THIN = Side(style='thin', color='B0B0B0')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _is_subtotal_row(ws, r):
    return any(str(ws.cell(row=r, column=c).value or '').strip() == '小計' for c in (3, 7))


def scan_sections(ws):
    """
    {section: {'label': str, 'rows': [data rows]}} を返す。
    データ行 = D列に名前がある行 (ヘッダ行 '名前' は除く)。
    セクション見出し行自体にデータがある場合 (E: Private Credit) も含める。
    """
    sections, order, cur = {}, [], None
    for r in range(1, ws.max_row + 1):
        b = str(ws.cell(row=r, column=2).value or '').strip()
        # 総合計行 / Summary 表に到達したら終了 (J セクションには小計行がない)
        if (str(ws.cell(row=r, column=3).value or '').strip() == '総合計'
                or str(ws.cell(row=r, column=5).value or '').strip() == SUMMARY_HEADER_LABEL):
            break
        if re.fullmatch(r'[A-J]', b):
            cur = b
            sections[cur] = {'label': str(ws.cell(row=r, column=3).value or '').strip(), 'rows': []}
            order.append(cur)
        elif cur is None:
            continue
        elif _is_subtotal_row(ws, r):
            cur = None
            continue
        d = ws.cell(row=r, column=4).value
        if d is not None and str(d).strip() not in ('', '名前'):
            sections[cur]['rows'].append(r)
    return sections, order


def _blocks(rows):
    """連続する行番号をブロックにまとめる: [127,130,131] → [(127,127),(130,131)]"""
    out = []
    for r in sorted(rows):
        if out and r == out[-1][1] + 1:
            out[-1][1] = r
        else:
            out.append([r, r])
    return [tuple(b) for b in out]


def _sumif_formula(rows, vcol, crit):
    vl = get_column_letter(vcol)
    parts = [f'SUMIF(D{a}:D{b},"{crit}",{vl}{a}:{vl}{b})' for a, b in _blocks(rows)]
    return '=' + '+'.join(parts) if parts else '=0'


def find_summary_rows(ws):
    """Summary 表の {class_code: row} と A〜I / A〜J 合計行を返す"""
    hdr = None
    for r in range(100, min(ws.max_row, 300) + 1):
        if str(ws.cell(row=r, column=5).value or '').strip() == SUMMARY_HEADER_LABEL:
            hdr = r
            break
    rows = {}
    if hdr:
        for r in range(hdr + 1, hdr + 20):
            code = str(ws.cell(row=r, column=5).value or '').strip()
            if re.fullmatch(r'[A-J]', code):
                rows[code] = r
    return rows


def write_person_summary(ws):
    """ws (資産クラスシート) に名前別集計表を書き込む (既存領域は上書き)"""
    sections, order = scan_sections(ws)
    sum_rows = find_summary_rows(ws)

    # 既存の表領域をクリア (再実行時の残骸を消す)
    for r in range(TABLE_TOP, TABLE_TOP + TABLE_ROWS):
        for c in range(TABLE_LEFT, TABLE_LEFT + TABLE_COLS):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.fill = PatternFill()
            cell.border = Border()
            cell.font = Font()

    E, F = TABLE_LEFT, TABLE_LEFT + 1
    pc0 = TABLE_LEFT + 2                          # 最初の名前列 (G)
    c_tot = pc0 + len(PERSONS)                    # 合計列 (K)
    c_unk = c_tot + 1                             # 未分類(差) 列 (L)
    L = get_column_letter

    ws.cell(row=TABLE_TOP, column=E, value='名前別 資産クラス集計（時価・円）').font = Font(bold=True, size=12, color='1F3864')

    hr = TABLE_TOP + 1
    heads = ['', '資産クラス'] + [p for p, _ in PERSONS] + ['合計', '未分類(差)']
    for i, h in enumerate(heads):
        c = ws.cell(row=hr, column=E + i, value=h or None)
        c.fill, c.font, c.border = HDR_FILL, HDR_FONT, BORDER
        c.alignment = Alignment(horizontal='center', vertical='center')

    r = hr + 1
    first_cls_row = r
    cls_row_of = {}
    for code in [k for k in 'ABCDEFGHIJ' if k in sections]:
        sec = sections[code]
        vcol = VALUE_COL.get(code, DEFAULT_VALUE_COL)
        ws.cell(row=r, column=E, value=code).border = BORDER
        ws.cell(row=r, column=F, value=sec['label']).border = BORDER
        for i, (_, crit) in enumerate(PERSONS):
            c = ws.cell(row=r, column=pc0 + i, value=_sumif_formula(sec['rows'], vcol, crit))
            c.number_format, c.border = '#,##0', BORDER
        c = ws.cell(row=r, column=c_tot, value=f'=SUM({L(pc0)}{r}:{L(c_tot - 1)}{r})')
        c.number_format, c.border = '#,##0', BORDER
        c.font = Font(bold=True)
        # 未分類(差) = Summary 表のクラス時価合計 (G列) − 名前別合計
        if code in sum_rows:
            c = ws.cell(row=r, column=c_unk, value=f'=G{sum_rows[code]}-{L(c_tot)}{r}')
        else:
            c = ws.cell(row=r, column=c_unk, value=None)
        c.number_format, c.border = '#,##0', BORDER
        cls_row_of[code] = r
        r += 1
    last_cls_row = r - 1

    # 合計行: 流動資産 (A〜I) / 資産合計 (A〜J) / 構成比
    liquid_rows = [cls_row_of[k] for k in 'ABCDEFGHI' if k in cls_row_of]
    total_specs = [
        ('流動資産（A〜I）', liquid_rows),
        ('資産合計（A〜J）', list(range(first_cls_row, last_cls_row + 1))),
    ]
    tot_row_of = {}
    for label, rows in total_specs:
        ws.cell(row=r, column=E).fill = TOT_FILL
        ws.cell(row=r, column=E).border = BORDER
        c = ws.cell(row=r, column=F, value=label)
        c.font, c.fill, c.border = Font(bold=True), TOT_FILL, BORDER
        for col in range(pc0, c_unk + 1):
            refs = ','.join(f'{L(col)}{x}' for x in rows)
            c = ws.cell(row=r, column=col, value=f'=SUM({refs})')
            c.number_format, c.font, c.fill, c.border = '#,##0', Font(bold=True), TOT_FILL, BORDER
        tot_row_of[label] = r
        r += 1

    # 構成比 (資産合計に対する各人の比率)
    aj = tot_row_of['資産合計（A〜J）']
    ws.cell(row=r, column=E).border = BORDER
    c = ws.cell(row=r, column=F, value='構成比（資産合計）')
    c.font, c.border = Font(bold=True), BORDER
    for col in range(pc0, c_tot + 1):
        c = ws.cell(row=r, column=col, value=f'=IF(${L(c_tot)}${aj}=0,"",{L(col)}{aj}/${L(c_tot)}${aj})')
        c.number_format, c.border = '0.0%', BORDER
    r += 1
    li = tot_row_of['流動資産（A〜I）']
    ws.cell(row=r, column=E).border = BORDER
    c = ws.cell(row=r, column=F, value='構成比（流動資産）')
    c.font, c.border = Font(bold=True), BORDER
    for col in range(pc0, c_tot + 1):
        c = ws.cell(row=r, column=col, value=f'=IF(${L(c_tot)}${li}=0,"",{L(col)}{li}/${L(c_tot)}${li})')
        c.number_format, c.border = '0.0%', BORDER

    return {'sections': sections, 'cls_row_of': cls_row_of, 'tot_row_of': tot_row_of}


# ══════════════════════════════════════════════════════════════
# 検算用: 数式を Python で評価して名前別の値を算出
# ══════════════════════════════════════════════════════════════
def make_evaluator(ws_f, ws_c):
    from openpyxl.utils.datetime import to_excel
    memo = {}

    def ev(r, c, depth=12):
        key = (r, c)
        if key in memo:
            return memo[key]
        if depth <= 0:
            return None
        v = ws_c.cell(row=r, column=c).value if ws_c is not None else None
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            memo[key] = float(v)
            return memo[key]
        v = ws_f.cell(row=r, column=c).value
        res = None
        if v is None:
            res = None
        elif isinstance(v, datetime):
            try:
                res = float(to_excel(v))
            except Exception:
                res = None
        elif isinstance(v, (int, float)):
            res = float(v)
        else:
            s = str(v).strip()
            if not s.startswith('='):
                try:
                    res = float(s)
                except ValueError:
                    res = None
            else:
                expr = s[1:].replace('$', '')
                if '!' in expr:
                    res = None
                else:
                    def _sum(m):
                        tot = 0.0
                        for arg in m.group(1).split(','):
                            arg = arg.strip()
                            mr = re.fullmatch(r'([A-Z]{1,3})(\d+):([A-Z]{1,3})(\d+)', arg)
                            if mr:
                                c1, c2 = column_index_from_string(mr.group(1)), column_index_from_string(mr.group(3))
                                for rr in range(int(mr.group(2)), int(mr.group(4)) + 1):
                                    for cc in range(c1, c2 + 1):
                                        x = ev(rr, cc, depth - 1)
                                        tot += x or 0.0
                            else:
                                ms = re.fullmatch(r'([A-Z]{1,3})(\d+)', arg)
                                if ms:
                                    x = ev(int(ms.group(2)), column_index_from_string(ms.group(1)), depth - 1)
                                    tot += x or 0.0
                        return f'({tot})'
                    expr = re.sub(r'SUM\(([^()]*)\)', _sum, expr, flags=re.IGNORECASE)

                    def _ref(m):
                        x = ev(int(m.group(2)), column_index_from_string(m.group(1)), depth - 1)
                        return f'({x})' if x is not None else '(0)'
                    expr = re.sub(r'\b([A-Z]{1,3})(\d+)\b', _ref, expr)
                    if re.fullmatch(r'[\d\s\+\-\*\/\(\)\.eE]+', expr):
                        try:
                            res = float(eval(expr))  # noqa: S307 — 数字と演算子のみ
                        except Exception:
                            res = None
        memo[key] = res
        return res
    return ev


def compute_person_values(ws_f, ws_c=None):
    """名前別 × クラス別の時価を Python で算出 (検算・レポート用)"""
    sections, _ = scan_sections(ws_f)
    ev = make_evaluator(ws_f, ws_c)
    result = {}
    for code in [k for k in 'ABCDEFGHIJ' if k in sections]:
        sec = sections[code]
        vcol = VALUE_COL.get(code, DEFAULT_VALUE_COL)
        per = {p: 0.0 for p, _ in PERSONS}
        unmatched = 0.0
        for r in sec['rows']:
            name = str(ws_f.cell(row=r, column=4).value or '').strip().lower()
            val = ev(r, vcol) or 0.0
            hit = False
            for p, crit in PERSONS:
                if name.startswith(crit.rstrip('*').lower()):
                    per[p] += val
                    hit = True
                    break
            if not hit:
                unmatched += val
        result[code] = {'label': sec['label'], 'per': per, 'unmatched': unmatched}
    return result


def _find_latest_excel(base_dir):
    files = glob.glob(os.path.join(base_dir, '資産クラス整理_????????.xlsx'))
    if not files:
        raise FileNotFoundError('資産クラス整理_YYYYMMDD.xlsx が見つかりません')
    return max(files)


if __name__ == '__main__':
    import openpyxl
    if sys.stdout.encoding and sys.stdout.encoding.lower() in ('cp932', 'shift_jis', 'shift-jis'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    path = _find_latest_excel(BASE_DIR)
    wb = openpyxl.load_workbook(path)
    sn = sorted([s for s in wb.sheetnames
                 if s.startswith('資産クラス') and s[5:].isdigit() and len(s[5:]) == 8])[-1]
    ws = wb[sn]
    write_person_summary(ws)
    wb.save(path)
    print(f'[OK] 名前別集計を書き込み: {os.path.basename(path)} / {sn} (row{TABLE_TOP}〜)')
