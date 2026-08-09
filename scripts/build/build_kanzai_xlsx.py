"""管財事件の xlsx 書式一式を case.yaml から生成する builder。

対象（--doc で選択。いずれも裁判所配布ブックのコピーに値を書き、数式・書式を保持する）:
  shisan-fusai     資産及び負債一覧表（0208/0106）
  saikensha-kanzai 債権者一覧表8シート（0209-0216/0107-0114）
  hikazei-check    被課税公租公課チェック表（0217/0115）
  shisan-kanzai    財産目録17/16シート（0218-0234/0116-0131）
  lease            リース物件一覧表（0235=0132 共通）
  sosho-shobun     訴訟・処分行為一覧表（0236-0237/0133-0134）

列位置はシートのヘッダ行の文言から動的に解決する（版・自然人/法人の列ズレに追従）。
チェック欄は「□」セルを「☑」に置換。自由財産拡張申立欄は「■」を書く。
行数が足りないときは直前行の書式をコピーして挿入する。
"""
import argparse
import json
import sys
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
from scripts.build import xlsxlib
from scripts.build.xlsxlib import XlsxFiller


# ---------- 共通ヘルパ ----------

def norm(s):
    return "".join(str(s).split()) if s is not None else ""


def header_map(ws, header_row=1):
    """ヘッダ行の文言（空白除去）→ 列レターの辞書。"""
    out = {}
    for c in ws[header_row]:
        if c.value is not None:
            out[norm(c.value)] = c.column_letter
    return out


def find_col(hmap, *keywords):
    """キーワード（部分一致・先勝ち）で列レターを引く。"""
    for kw in keywords:
        for label, col in hmap.items():
            if kw in label:
                return col
    return None


def ensure_rows(ws, first_data_row, count, numbered_rows):
    """データ行が足りなければ最終データ行の書式をコピーして挿入し、番号列を振り直す。"""
    import copy as _copy
    have = numbered_rows
    if count <= have:
        return
    last = first_data_row + have - 1
    add = count - have
    ws.insert_rows(last + 1, add)
    for i in range(1, add + 1):
        for c in ws[last]:
            dst = ws.cell(row=last + i, column=c.column)
            if c.has_style:
                dst._style = _copy.copy(c._style)
        ws.cell(row=last + i, column=1, value=have + i)


def creditor_class(kind):
    """creditors.kind → 債権者一覧表のシート分類番号（1〜7）。"""
    if kind in ("貸金", "クレジット", "立替金", "保証", "求償"):
        return 1  # 借入金
    return {"手形小切手": 2, "買掛金": 3, "リース": 4, "労働債権": 5,
            "公租公課": 7}.get(kind, 6)  # 家賃・養育費・その他 → 6


CLASS_LABELS = {1: "借入金", 2: "手形・小切手債権", 3: "買掛金", 4: "リース債権",
                5: "労働債権", 6: "その他", 7: "公租公課"}

# 優先的破産債権・財団債権に集計する分類
PRIORITY_CLASSES = {5, 7}


def labor_total(c):
    lab = c.get("labor") or {}
    vals = [lab.get("salary"), lab.get("retirement_allowance"),
            lab.get("dismissal_notice_allowance")]
    known = [v for v in vals if isinstance(v, int)]
    if known:
        return sum(known)
    return c.get("balance")


def creditor_amount(c):
    if c.get("kind") == "労働債権":
        return labor_total(c)
    return c.get("balance")


# ---------- 資産の区分集計（0208・0218 総括表と共通） ----------

def asset_categories(case, corporate=False):
    """財産目録の区分順に (区分名, 名目額, 回収見込額, 備考) を返す。

    recovery が無い項目は名目額と同額。退職金は 1/8、不動産は担保控除、
    回収見込みのない貸付金は 0 の既定を適用する。
    """
    a = (case.get("assets") or {})

    def tot(items, value_key="value", recovery_default=None):
        nominal = recovery = 0
        for it in items or []:
            v = it.get(value_key)
            if not isinstance(v, int):
                for alt in ("balance", "amount", "surrender_value", "deposit",
                            "face_value", "total"):
                    if isinstance(it.get(alt), int):
                        v = it[alt]
                        break
            if not isinstance(v, int):
                continue
            nominal += v
            r = it.get("recovery")
            if not isinstance(r, int):
                r = recovery_default(it, v) if recovery_default else v
            recovery += r
        return nominal, recovery

    cash = a.get("cash") or 0
    dep_n, dep_r = tot((a.get("deposits") or []) + (a.get("tsumitate") or []))
    ins_n, ins_r = tot(a.get("insurance"), "surrender_value")
    veh_n, veh_r = tot(a.get("vehicles"))
    shk_n, shk_r = tot(a.get("shikikin"), "deposit",
                       lambda it, v: it["refund"] if isinstance(it.get("refund"), int) else v)
    ret = a.get("retirement_expectation") or 0
    ret_r = ret // 8  # 8分の1相当額（大阪地裁運用）
    ovp_n, ovp_r = tot(a.get("overpayment"), "amount")
    re_n, re_r = tot(a.get("real_estate"),
                     recovery_default=lambda it, v: max(0, v - (it.get("mortgage") or 0)))
    sec_n, sec_r = tot(a.get("securities"))
    nr_n, nr_r = tot(a.get("notes_receivable"), "face_value")
    ar_n, ar_r = tot(a.get("accounts_receivable"), "amount")
    ln_n, ln_r = tot(a.get("loans_receivable"), "amount",
                     lambda it, v: v if it.get("recoverable") else 0)
    inv_n, inv_r = tot(a.get("inventory"))
    mac_n, mac_r = tot(a.get("machinery"))
    fix_n, fix_r = tot(a.get("fixtures"))
    oth_n, oth_r = tot(a.get("other"))
    upc_n, upc_r = tot(a.get("upcoming"), "amount")

    rows = [
        ("現金", cash, cash),
        ("預貯金・積立金" if not corporate else "預貯金", dep_n, dep_r),
        ("保険解約返戻金", ins_n, ins_r),
        ("自動車", veh_n, veh_r),
        ("賃借保証金・敷金", shk_n, shk_r),
    ]
    if not corporate:
        rows += [("退職金", ret, ret_r)]
    rows += [
        ("過払金", ovp_n, ovp_r),
        ("不動産", re_n, re_r),
        ("有価証券", sec_n, sec_r),
        ("受取手形・小切手", nr_n, nr_r),
        ("売掛金", ar_n, ar_r),
        ("貸付金", ln_n, ln_r),
        ("在庫商品", inv_n, inv_r),
        ("機械・工具類", mac_n, mac_r),
        ("什器備品", fix_n, fix_r),
        ("その他", oth_n + upc_n, oth_r + upc_r),
    ]
    return rows


# ---------- 0208 資産及び負債一覧表 ----------

def build_shisan_fusai(case, vdir, output):
    filler = XlsxFiller(vdir / "template.xlsx")
    ws = filler.wb.worksheets[0]
    name = ((case.get("meta") or {}).get("applicant") or {}).get("name", "")
    corporate = bool((case.get("corporation") or {}).get("name"))
    if corporate:
        name = case["corporation"]["name"]
    filler.write(ws.title, "J1", name, "debtor")

    # 資産（B列の科目ラベルに行を合わせて C/D を書く）
    cats = dict((n, (v, r)) for n, v, r in asset_categories(case, corporate))
    label_rows = {}
    for row in ws.iter_rows(min_row=5, max_row=25, min_col=2, max_col=2):
        for c in row:
            if c.value:
                label_rows[norm(c.value)] = c.row
    total_n = total_r = 0
    for label, (v, r) in cats.items():
        rr = label_rows.get(norm(label))
        if rr is None:
            continue
        if v:
            filler.write(ws.title, f"C{rr}", v, "asset")
        if v or r:
            filler.write(ws.title, f"D{rr}", r, "asset")
        total_n += v
        total_r += r

    # 負債（I列の科目ラベルに J を書く）
    debt = {}
    for c in case.get("creditors") or []:
        amt = creditor_amount(c)
        if not isinstance(amt, int):
            continue
        debt[creditor_class(c.get("kind"))] = debt.get(creditor_class(c.get("kind")), 0) + amt
    debt_rows = {"借入金": debt.get(1), "手形・小切手債権": debt.get(2),
                 "買掛金": debt.get(3), "リース債権": debt.get(4),
                 "労働債権": debt.get(5), "公租公課": debt.get(7)}
    sonota_ippan = debt.get(6)
    ilabel_rows = {}
    for row in ws.iter_rows(min_row=5, max_row=20, min_col=9, max_col=9):
        for c in row:
            if c.value:
                ilabel_rows.setdefault(norm(c.value), []).append(c.row)
    for label, v in debt_rows.items():
        if v is None:
            continue
        rr = ilabel_rows.get(norm(label))
        if rr:
            filler.write(ws.title, f"J{rr[0]}", v, "debt")
    if sonota_ippan is not None:
        rr = ilabel_rows.get(norm("その他の債権"))
        if rr:  # 1つ目＝一般破産債権のその他
            filler.write(ws.title, f"J{rr[0]}", sonota_ippan, "debt")
    # 債権者総数（J列の「名」ラベルのセルに「N名」）
    n_cred = len(case.get("creditors") or [])
    for row in ws.iter_rows(min_row=21, max_row=24, min_col=10, max_col=10):
        for c in row:
            if norm(c.value) == "名":
                filler.write(ws.title, c.coordinate, f"{n_cred}名", "count")
    filler.wb.save(output)
    return {"total_assets": total_n, "total_recovery": total_r,
            "total_debt": sum(v for v in debt.values()),
            "creditor_count": n_cred, "warnings": filler.warnings}


# ---------- 0209-0216 債権者一覧表（管財） ----------

def build_saikensha_kanzai(case, vdir, output):
    filler = XlsxFiller(vdir / "template.xlsx")
    wb = filler.wb
    creditors = case.get("creditors") or []
    for i, c in enumerate(creditors):
        c.setdefault("no", i + 1)

    # 総括シート（1枚目）
    ws = wb.worksheets[0]
    hm = header_map(ws)
    cols = {"name": find_col(hm, "債権者名"), "zip": find_col(hm, "〒"),
            "addr": find_col(hm, "住所"), "tel": find_col(hm, "TEL", "ＴＥＬ"),
            "kind": find_col(hm, "債権の種類"), "note": find_col(hm, "備考")}
    ensure_rows(ws, 2, len(creditors), 10)
    total = 0
    total_priority = 0
    unknown = 0
    for i, c in enumerate(creditors):
        r = 2 + i
        cls = creditor_class(c.get("kind"))
        filler.write(ws.title, f"{cols['name']}{r}", c.get("name", ""), "sokatsu")
        if c.get("zip"):
            filler.write(ws.title, f"{cols['zip']}{r}", c["zip"], "sokatsu")
        if c.get("address"):
            filler.write(ws.title, f"{cols['addr']}{r}", c["address"], "sokatsu")
        tel = c.get("tel", "")
        if c.get("fax"):
            tel = f"{tel}\n{c['fax']}" if tel else c["fax"]
        if tel:
            filler.write(ws.title, f"{cols['tel']}{r}", tel, "sokatsu")
        filler.write(ws.title, f"{cols['kind']}{r}", cls, "sokatsu")
        note = c.get("note", "")
        if cls == 6:
            note = (c.get("kind", "") + ("　" + note if note else "")) or note
        if note:
            filler.write(ws.title, f"{cols['note']}{r}", note, "sokatsu")
        amt = creditor_amount(c)
        if isinstance(amt, int):
            total += amt
            if cls in PRIORITY_CLASSES:
                total_priority += amt
        else:
            unknown += 1
    # 債権者数・債務総額（ラベル行の空セル）
    foot = 2 + max(len(creditors), 10) + 2  # 案内表示行の下＝債権者数の行
    for row in ws.iter_rows(min_row=foot, max_row=foot + 4):
        labels = {norm(c.value): c for c in row if c.value}
        if "債権者数" in labels:
            rr = row[0].row
            filler.write(ws.title, f"C{rr}", len(creditors), "sokatsu")
            filler.write(ws.title, f"E{rr}", total, "sokatsu")
            filler.write(ws.title, f"E{rr + 1}", total_priority, "sokatsu")
            break

    # 分類別シート
    by_class = {}
    for c in creditors:
        by_class.setdefault(creditor_class(c.get("kind")), []).append(c)
    for si, ws in enumerate(wb.worksheets[1:], start=1):
        cls = si  # 2枚目=1借入金 … 8枚目=7公租公課
        items = by_class.get(cls, [])
        if not items:
            continue
        hm = header_map(ws)
        get = lambda *kw: find_col(hm, *kw)
        # 番号入りの既存行数を数える
        numbered = sum(1 for row in ws.iter_rows(min_row=2, min_col=1, max_col=1)
                       for c in row if isinstance(c.value, int))
        ensure_rows(ws, 2, len(items), numbered or 10)
        for i, c in enumerate(items):
            r = 2 + i
            def W(col, val):
                if col and val not in (None, ""):
                    filler.write(ws.title, f"{col}{r}", val, f"sheet{cls}")
            if cls == 7:  # 公租公課: 税目・所轄・年度
                W(get("税目"), c.get("tax_kind") or c.get("name"))
                W(get("所轄"), c.get("tax_office") or c.get("name"))
                W(get("〒"), c.get("zip"))
                W(get("住所"), c.get("address"))
                W(get("TEL", "ＴＥＬ"), c.get("tel"))
                W(get("金額"), c.get("balance"))
                W(get("年度"), c.get("tax_year"))
                W(get("備考"), c.get("note"))
                continue
            W(get("債権者名", "債権者"), c.get("name"))
            W(get("〒"), c.get("zip"))
            W(get("住所"), c.get("address"))
            tel = c.get("tel", "")
            if c.get("fax"):
                tel = f"{tel}\n{c['fax']}" if tel else c["fax"]
            W(get("TEL", "ＴＥＬ"), tel)
            if cls == 5:  # 労働債権
                lab = c.get("labor") or {}
                W(get("給料未払期間"), lab.get("period"))
                W(get("①給料", "①"), lab.get("salary"))
                W(get("②退職手当", "②"), lab.get("retirement_allowance"))
                W(get("③解雇予告手当", "③"), lab.get("dismissal_notice_allowance"))
                W(get("備考"), c.get("note"))
                continue
            W(get("金額"), c.get("balance"))
            if cls == 1:
                W(get("借入日"), c.get("origin_date"))
                W(get("最後の弁済日"), c.get("last_payment"))
                W(get("使途"), c.get("use"))
                for key, flag in (("別除権", c.get("secured")),
                                  ("保証人", c.get("guarantor"))):
                    col = get(key)
                    if col and flag:
                        cell = ws[f"{col}{r}"]
                        if cell.value == "□":
                            cell.value = "☑"
                        else:
                            filler.write(ws.title, f"{col}{r}", "☑", "sheet1")
            if cls == 2:
                tg = c.get("tegata") or {}
                W(get("手形小切手番号"), tg.get("number"))
                W(get("裏書"), tg.get("endorsement"))
                W(get("支払期日"), tg.get("due"))
            if cls == 4:
                ls = c.get("lease") or {}
                W(get("リース物件"), ls.get("item"))
            if cls == 6:
                W(get("債権の種類"), c.get("kind"))
            W(get("備考"), c.get("note"))
    filler.wb.save(output)
    return {"creditor_count": len(creditors), "total_debt": total,
            "total_priority": total_priority, "unknown_balance": unknown,
            "warnings": filler.warnings}


# ---------- 0217 被課税公租公課チェック表 ----------

# 税目名 → 書式の税目ラベルに現れる別表記（法人税目・徴収区分の違いを吸収）
TAX_ALIASES = {
    "法人住民税": ("法人府民税", "法人等市民税", "法人市民税"),
    "住民税": ("特別徴収府県民税", "府県民税・市町村民税"),
    "社会保険料": ("健康保険料", "厚生年金保険料"),
    "事業税": ("法人事業税", "個人事業税"),
}


def build_hikazei_check(case, vdir, output):
    filler = XlsxFiller(vdir / "template.xlsx")
    ws = filler.wb.worksheets[0]
    kouso = [c for c in case.get("creditors") or [] if c.get("kind") == "公租公課"]
    # 税目ラベル（B列）→ 行。チェックは A列 の □ を ☑ に
    checked = 0
    for row in ws.iter_rows(min_row=3):
        a = ws.cell(row=row[0].row, column=1)
        b = ws.cell(row=row[0].row, column=2)
        if a.value != "□" or not b.value:
            continue
        label = norm(b.value)
        for c in kouso:
            tk = norm(c.get("tax_kind") or "")
            if tk and (tk in label or label.startswith(tk[:3])
                       or any(al in label for al in TAX_ALIASES.get(tk, ()))):
                a.value = "☑"
                checked += 1
                if c.get("tax_year"):
                    filler.write(ws.title, f"G{a.row}", c["tax_year"], "check")
                if isinstance(c.get("balance"), int):
                    filler.write(ws.title, f"H{a.row}", c["balance"], "check")
                if c.get("tax_office"):
                    cur = b.value if False else ws.cell(row=a.row, column=3).value
                    if isinstance(cur, str) and "（　　　　）" in cur:
                        ws.cell(row=a.row, column=3).value = cur.replace(
                            "（　　　　）", f"（{c['tax_office'].replace('税務署','').replace('市税事務所','')}）", 1)
                break
    filler.wb.save(output)
    return {"checked": checked, "kouso_count": len(kouso), "warnings": filler.warnings}


# ---------- 0218-0234 財産目録（管財） ----------

def build_shisan_kanzai(case, vdir, output):
    filler = XlsxFiller(vdir / "template.xlsx")
    wb = filler.wb
    a = case.get("assets") or {}
    corporate = bool((case.get("corporation") or {}).get("name"))

    def sheet_by(kw):
        for ws in wb.worksheets:
            if kw in norm(ws.title):
                return ws
        return None

    def fill_sheet(kw, items, mapping, header_row=1, jiyuzaisan_key="jiyuzaisan"):
        """mapping: [(ヘッダキーワード, 値関数)]"""
        ws = sheet_by(kw)
        if ws is None or not items:
            return
        hm = header_map(ws, header_row)
        numbered = sum(1 for row in ws.iter_rows(min_row=header_row + 1, min_col=1, max_col=1)
                       for c in row if isinstance(c.value, int))
        ensure_rows(ws, header_row + 1, len(items), numbered or 5)
        for i, it in enumerate(items):
            r = header_row + 1 + i
            for kws, fn in mapping:
                col = find_col(hm, *kws) if isinstance(kws, tuple) else find_col(hm, kws)
                v = fn(it)
                if col and v not in (None, ""):
                    filler.write(ws.title, f"{col}{r}", v, kw)
            jcol = find_col(hm, "自由財産")
            if jcol and it.get(jiyuzaisan_key):
                filler.write(ws.title, f"{jcol}{r}", "■", kw)

    rec = lambda it, *alts: it.get("recovery") if isinstance(it.get("recovery"), int) else next(
        (it.get(k) for k in alts if isinstance(it.get(k), int)), None)

    fill_sheet("預貯金", (a.get("deposits") or []) + (a.get("tsumitate") or []), [
        (("金融機関",), lambda it: it.get("bank") or it.get("kind")),
        (("支店名",), lambda it: it.get("branch")),
        (("種類",), lambda it: it.get("account_type") or it.get("kind")),
        (("口座番号",), lambda it: it.get("account_no")),
        (("残高",), lambda it: it.get("balance") if it.get("balance") is not None else it.get("total")),
        (("回収見込額",), lambda it: rec(it, "balance", "total")),
        (("備考",), lambda it: it.get("note") or it.get("as_of")),
    ])
    fill_sheet("保険", a.get("insurance") or [], [
        (("保険会社名",), lambda it: it.get("company")),
        (("保険種類",), lambda it: it.get("kind")),
        (("証券番号",), lambda it: it.get("policy_no")),
        (("名目額",), lambda it: it.get("surrender_value")),
        (("回収見込額",), lambda it: rec(it, "surrender_value")),
        (("備考",), lambda it: it.get("note")),
    ])
    fill_sheet("自動車", a.get("vehicles") or [], [
        (("車名",), lambda it: it.get("name")),
        (("初度登録年",), lambda it: it.get("first_registration")),
        (("登録番号",), lambda it: it.get("registration_no")),
        (("保管場所",), lambda it: it.get("location")),
        (("簿価",), lambda it: it.get("value")),
        (("回収見込額",), lambda it: rec(it, "value")),
        (("備考",), lambda it: it.get("note")),
    ])
    fill_sheet("敷金", a.get("shikikin") or [], [
        (("賃借物件",), lambda it: it.get("property")),
        (("差入額",), lambda it: it.get("deposit")),
        (("契約上の返戻金",), lambda it: it.get("refund")),
        (("滞納額",), lambda it: it.get("arrears")),
        (("原状回復費用",), lambda it: it.get("restore_cost")),
        (("回収見込額",), lambda it: rec(it, "refund", "deposit")),
        (("備考",), lambda it: it.get("note")),
    ])
    ret = a.get("retirement_expectation")
    if ret and not corporate:
        cur = ((case.get("career") or [{}])[0] or {})
        fill_sheet("退職金", [{"employer": cur.get("employer"), "start": cur.get("start"),
                              "amount": ret, "eighth": ret // 8}], [
            (("雇用主",), lambda it: it.get("employer")),
            (("勤務開始日",), lambda it: it.get("start")),
            (("支給見込額",), lambda it: it.get("amount")),
            (("1/8相当額", "８分の１", "8分の1"), lambda it: it.get("eighth")),
        ])
    fill_sheet("過払金", a.get("overpayment") or [], [
        (("相手方",), lambda it: it.get("against")),
        (("額面額",), lambda it: it.get("amount")),
        (("合意額",), lambda it: it.get("agreed") or it.get("amount")),
        (("回収費用控除後残金",), lambda it: rec(it, "amount")),
        (("備考",), lambda it: it.get("note")),
    ])
    fill_sheet("不動産", a.get("real_estate") or [], [
        (("種類",), lambda it: it.get("kind")),
        (("所在地",), lambda it: it.get("address")),
        (("地番",), lambda it: it.get("lot_no")),
        (("評価額",), lambda it: it.get("value")),
        (("被担保債権額",), lambda it: it.get("mortgage")),
        (("回収見込額",), lambda it: it.get("recovery") if isinstance(it.get("recovery"), int)
            else (max(0, (it.get("value") or 0) - (it.get("mortgage") or 0)) or None)),
        (("備考",), lambda it: it.get("note")),
    ])
    fill_sheet("有価証券", a.get("securities") or [], [
        (("財産の内容",), lambda it: it.get("name")),
        (("数量",), lambda it: it.get("quantity")),
        (("証券番号",), lambda it: it.get("cert_no")),
        (("所在場所", "証券等"), lambda it: it.get("location")),
        (("簿価",), lambda it: it.get("value")),
        (("回収見込額",), lambda it: rec(it, "value")),
        (("備考",), lambda it: it.get("note")),
    ])
    fill_sheet("手形・小切手", a.get("notes_receivable") or [], [
        (("振出人",), lambda it: it.get("drawer")),
        (("〒",), lambda it: it.get("zip")),
        (("住所",), lambda it: it.get("address")),
        (("手形小切手番号",), lambda it: it.get("number")),
        (("裏書人",), lambda it: it.get("endorser")),
        (("支払期日",), lambda it: it.get("due")),
        (("額面",), lambda it: it.get("face_value")),
        (("回収見込額",), lambda it: rec(it, "face_value")),
    ])
    fill_sheet("売掛金", a.get("accounts_receivable") or [], [
        (("債務者",), lambda it: it.get("debtor")),
        (("〒",), lambda it: it.get("zip")),
        (("住所",), lambda it: it.get("address")),
        (("金額",), lambda it: it.get("amount")),
        (("回収見込額",), lambda it: rec(it, "amount")),
        (("備考",), lambda it: it.get("note")),
    ])
    fill_sheet("貸付金", a.get("loans_receivable") or [], [
        (("債務者",), lambda it: it.get("debtor")),
        (("金額",), lambda it: it.get("amount")),
        (("回収見込額",), lambda it: it.get("recovery") if isinstance(it.get("recovery"), int)
            else (it.get("amount") if it.get("recoverable") else 0)),
        (("備考",), lambda it: it.get("reason") or it.get("note")),
    ])
    fill_sheet("在庫商品", a.get("inventory") or [], [
        (("品名",), lambda it: it.get("name")),
        (("個数",), lambda it: it.get("quantity")),
        (("所在場所",), lambda it: it.get("location")),
        (("簿価",), lambda it: it.get("value")),
        (("回収見込額",), lambda it: rec(it, "value")),
        (("備考",), lambda it: it.get("note")),
    ])
    fill_sheet("機械・工具", a.get("machinery") or [], [
        (("名称",), lambda it: it.get("name")),
        (("個数",), lambda it: it.get("quantity")),
        (("所在場所",), lambda it: it.get("location")),
        (("簿価",), lambda it: it.get("value")),
        (("回収見込額",), lambda it: rec(it, "value")),
        (("備考",), lambda it: it.get("note")),
    ])
    fill_sheet("什器備品", a.get("fixtures") or [], [
        (("品名",), lambda it: it.get("name")),
        (("個数",), lambda it: it.get("quantity")),
        (("所在場所",), lambda it: it.get("location")),
        (("簿価",), lambda it: it.get("value")),
        (("回収見込額",), lambda it: rec(it, "value")),
        (("備考",), lambda it: it.get("note")),
    ])
    fill_sheet("その他", a.get("other") or [], [
        (("財産の種類",), lambda it: it.get("name")),
        (("数量",), lambda it: it.get("quantity")),
        (("所在場所",), lambda it: it.get("location")),
        (("簿価",), lambda it: it.get("value")),
        (("回収見込額",), lambda it: rec(it, "value")),
        (("備考",), lambda it: it.get("note")),
    ])
    # 処分済財産（ヘッダは4行目）
    ws = sheet_by("処分済財産")
    disposed = a.get("disposed") or []
    if ws is not None and disposed:
        hm = header_map(ws, 4)
        numbered = sum(1 for row in ws.iter_rows(min_row=5, min_col=1, max_col=1)
                       for c in row if isinstance(c.value, int))
        ensure_rows(ws, 5, len(disposed), numbered or 5)
        for i, it in enumerate(disposed):
            r = 5 + i
            for kw, key in (("財産の種類", "name"), ("処分日", "date"),
                            ("処分価格", "price"), ("使途", "use")):
                col = find_col(hm, kw)
                if col and it.get(key) not in (None, ""):
                    filler.write(ws.title, f"{col}{r}", it[key], "disposed")
            for kw, key in (("架空計上", "fictitious"), ("処分", "disposed"),
                            ("決算書", "in_financials")):
                col = find_col(hm, kw)
                if col and it.get(key):
                    cell = ws[f"{col}{r}"]
                    cell.value = "☑" if cell.value in ("□", None) else cell.value

    # 総括表（B列ラベル合わせで C/D を書く。F=自由財産の額は■対象＋現金から集計）
    def jiyu_sum(items, value_keys=("recovery", "value", "surrender_value", "deposit",
                                    "balance", "amount", "total")):
        s = 0
        for it in items or []:
            if not it.get("jiyuzaisan"):
                continue
            for k in value_keys:
                if isinstance(it.get(k), int):
                    s += it[k]
                    break
        return s

    jiyu = {
        "現金": a.get("cash") or 0,  # 現金は本来的自由財産（99万円まで）
        "預貯金・積立金": jiyu_sum((a.get("deposits") or []) + (a.get("tsumitate") or [])),
        "保険解約返戻金": jiyu_sum(a.get("insurance")),
        "自動車": jiyu_sum(a.get("vehicles")),
        "賃借保証金・敷金": jiyu_sum(a.get("shikikin")),
        "過払金": jiyu_sum(a.get("overpayment")),
        "有価証券": jiyu_sum(a.get("securities")),
        "売掛金": jiyu_sum(a.get("accounts_receivable")),
        "貸付金": jiyu_sum(a.get("loans_receivable")),
        "什器備品": jiyu_sum(a.get("fixtures")),
        "その他": jiyu_sum(a.get("other")),
    }
    ws = wb.worksheets[0]
    cats = asset_categories(case, corporate)
    hm = header_map(ws)
    jcol = find_col(hm, "自由財産")
    label_rows = {}
    for row in ws.iter_rows(min_row=2, max_row=20, min_col=2, max_col=2):
        for c in row:
            if c.value:
                label_rows[norm(c.value)] = c.row
    total_n = total_r = 0
    for label, v, r in cats:
        rr = label_rows.get(norm(label))
        total_n += v
        total_r += r
        if rr is None:
            continue
        if v:
            filler.write(ws.title, f"C{rr}", v, "sokatsu")
        if v or r:
            filler.write(ws.title, f"D{rr}", r, "sokatsu")
        jv = jiyu.get(label)
        if jcol and jv:
            filler.write(ws.title, f"{jcol}{rr}", jv, "sokatsu")
    filler.wb.save(output)
    return {"total_assets": total_n, "total_recovery": total_r,
            "jiyuzaisan_total": sum(jiyu.values()),
            "warnings": filler.warnings}


# ---------- 0235 リース物件一覧表 ----------

def build_lease(case, vdir, output):
    filler = XlsxFiller(vdir / "template.xlsx")
    ws = filler.wb.worksheets[0]
    items = [c for c in case.get("creditors") or [] if c.get("kind") == "リース"]
    hm = header_map(ws)
    numbered = sum(1 for row in ws.iter_rows(min_row=2, min_col=1, max_col=1)
                   for c in row if isinstance(c.value, int))
    ensure_rows(ws, 2, len(items), numbered or 10)
    for i, c in enumerate(items):
        r = 2 + i
        ls = c.get("lease") or {}
        for kw, val in (("債権者名", c.get("name")), ("リース物件", ls.get("item")),
                        ("所在地", ls.get("location"))):
            col = find_col(hm, kw)
            if col and val:
                filler.write(ws.title, f"{col}{r}", val, "lease")
        for kw, flag in (("契約書等", ls.get("contract_doc")), ("返還済", ls.get("returned"))):
            col = find_col(hm, kw)
            if col and flag:
                cell = ws[f"{col}{r}"]
                cell.value = "☑" if cell.value in ("□", None) else cell.value
    filler.wb.save(output)
    return {"lease_count": len(items), "warnings": filler.warnings}


# ---------- 0236-0237 訴訟・処分行為一覧表 ----------

def build_sosho_shobun(case, vdir, output):
    filler = XlsxFiller(vdir / "template.xlsx")
    wb = filler.wb
    lawsuits = case.get("lawsuits") or []
    disposals = case.get("pre_bankruptcy_disposals") or []
    ws = wb.worksheets[0]
    hm = header_map(ws)
    numbered = sum(1 for row in ws.iter_rows(min_row=2, min_col=1, max_col=1)
                   for c in row if isinstance(c.value, int))
    ensure_rows(ws, 2, len(lawsuits), numbered or 10)
    for i, it in enumerate(lawsuits):
        r = 2 + i
        for kw, key in (("事件の種類", "kind"), ("相手方", "counterparty"),
                        ("係属裁判所", "court"), ("事件番号", "case_no"), ("備考", "note")):
            col = find_col(hm, kw)
            if col and it.get(key):
                filler.write(ws.title, f"{col}{r}", it[key], "sosho")
    ws = wb.worksheets[1]
    hm = header_map(ws)
    numbered = sum(1 for row in ws.iter_rows(min_row=2, min_col=1, max_col=1)
                   for c in row if isinstance(c.value, int))
    ensure_rows(ws, 2, len(disposals), numbered or 9)
    for i, it in enumerate(disposals):
        r = 2 + i
        for kw, key in (("相手方", "counterparty"), ("行為時期", "date"),
                        ("行為類型", "kind"), ("処分価格", "price"),
                        (("参考となるべき事項"), "note")):
            col = find_col(hm, kw if isinstance(kw, str) else kw)
            if col and it.get(key) not in (None, ""):
                filler.write(ws.title, f"{col}{r}", it[key], "shobun")
    filler.wb.save(output)
    return {"lawsuit_count": len(lawsuits), "disposal_count": len(disposals),
            "warnings": filler.warnings}


DOCS = {
    "shisan-fusai": ("shisan-fusai-ichiran", "資産及び負債一覧表", build_shisan_fusai),
    "saikensha-kanzai": ("saikensha-ichiran-kanzai", "債権者一覧表（管財）", build_saikensha_kanzai),
    "hikazei-check": ("hikazei-kouso-check", "被課税公租公課チェック表", build_hikazei_check),
    "shisan-kanzai": ("shisan-mokuroku-kanzai", "財産目録（管財）", build_shisan_kanzai),
    "lease": ("lease-bukken-ichiran", "リース物件一覧表", build_lease),
    "sosho-shobun": ("sosho-shobun-ichiran", "訴訟・処分行為一覧表", build_sosho_shobun),
}


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--doc", required=True, choices=sorted(DOCS))
    ap.add_argument("--case", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--court", default="osaka")
    args = ap.parse_args()

    case = yaml.safe_load(open(args.case, encoding="utf-8"))
    form, name, fn = DOCS[args.doc]
    vdir = xlsxlib.resolve_form(args.court, form, proc=xlsxlib.proc_of_case(case))
    if vdir is None:
        print(xlsxlib.registration_guidance(args.court, form, name))
        sys.exit(3)
    stats = fn(case, vdir, args.output)
    warnings = stats.pop("warnings", [])
    Path(args.output + ".meta.json").write_text(
        json.dumps({"doc": args.doc, **stats}, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"出力: {args.output}")
    print("  " + " / ".join(f"{k}={v}" for k, v in stats.items()))
    for w in warnings:
        print(f"  警告: {w}")


if __name__ == "__main__":
    main()
