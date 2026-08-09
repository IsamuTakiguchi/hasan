#!/usr/bin/env python3
"""債権者一覧表を case.yaml から生成する。

- 裁判所書式が登録済み（courts/<court>/forms/saikensha-ichiran/）なら fillmap 駆動で
  テンプレートに記入する。
- 未登録なら --generic で事務所内ドラフト様式（xlsx 新規生成）を出力する。

並び順: 一般債権者を case.yaml の記載順で先に、公租公課を末尾に置き、番号を振り直す。
出力と同時に <output>.meta.json（負債総額等）を書き、crosscheck_outputs.py が
書類間突合に使う。

使い方:
  python3 scripts/build/build_saikensha_ichiran.py --case cases/<id>/case.yaml \
      --output cases/<id>/output/債権者一覧表_<姓>.xlsx [--court osaka] [--generic]
"""
import argparse
import json
import sys
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent))
import xlsxlib  # noqa: E402

import openpyxl  # noqa: E402


def order_creditors(creditors):
    general = [c for c in creditors if c.get("kind") != "公租公課"]
    kouso = [c for c in creditors if c.get("kind") == "公租公課"]
    ordered = general + kouso
    for i, c in enumerate(ordered, 1):
        c = dict(c)
        c["no"] = i
        yield c


def umu(v):
    return {True: "有", False: "無"}.get(v, "")


def build_generic(case, out):
    meta = case.get("meta", {})
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "債権者一覧表"
    ws["A1"] = "債権者一覧表"
    ws["A1"].font = xlsxlib.FONT_TITLE
    ws["A2"] = f"申立人: {meta.get('applicant', {}).get('name', '')}"
    ws["A2"].font = xlsxlib.FONT
    ws["G2"] = f"基準日: {meta.get('junin_tsuchi_date', '')}（受任通知発送日）"
    ws["G2"].font = xlsxlib.FONT

    headers = ["番号", "債権者名", "住所", "債権の種類", "借入開始", "使途",
               "当初借入額(円)", "現在残高(円)", "最終弁済", "保証人", "別除権", "備考"]
    widths = [5, 24, 28, 10, 10, 22, 13, 13, 10, 7, 7, 18]
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        ws.cell(row=4, column=i, value=h)
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    xlsxlib.style_header(ws, 4, len(headers))

    row = 5
    total = 0
    unknown = 0
    for c in order_creditors(case.get("creditors", [])):
        vals = [c.get("no"), c.get("name", ""), c.get("address", ""), c.get("kind", ""),
                c.get("origin_date", ""), c.get("use", ""), c.get("principal"),
                c.get("balance"), c.get("last_payment", ""), umu(c.get("guarantor")),
                "有" if c.get("secured") else "", c.get("note", "")]
        for i, v in enumerate(vals, 1):
            ws.cell(row=row, column=i, value=v)
        xlsxlib.style_data(ws, row, len(headers))
        for col in ("G", "H"):
            ws[f"{col}{row}"].number_format = "#,##0"
        if isinstance(c.get("balance"), int):
            total += c["balance"]
        else:
            unknown += 1
        row += 1

    ws.cell(row=row, column=2, value="合計")
    ws.cell(row=row, column=8, value=f"=SUM(H5:H{row - 1})")
    xlsxlib.style_data(ws, row, len(headers))
    ws[f"H{row}"].number_format = "#,##0"
    ws[f"B{row}"].font = openpyxl.styles.Font(name="游ゴシック", size=10, bold=True)
    xlsxlib.draft_note(ws, row + 2)
    if unknown:
        xlsxlib.draft_note(ws, row + 3, f"※ 残高不明の債権者が {unknown} 名ある（合計に含まれない）。")

    Path(out).parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return {"total_debt": total, "creditor_count": row - 5, "unknown_balance": unknown}


USE_LITERALS = {
    "住宅ローン": "□住宅ローン", "購入": "□購入", "生活費": "□生活費",
    "返済": "□返済", "飲食交際遊興費": "□飲食交際遊興費", "保証": "□保証", "その他": "□その他",
}
USE_ROW = {"住宅ローン": "use_r0", "購入": "use_r0", "生活費": "use_r0",
           "返済": "use_r1", "飲食交際遊興費": "use_r1", "保証": "use_r2", "その他": "use_r2"}


def classify_use(c):
    """case.yaml の債権者から書式の使途チェック（複数可）を決める。判定できなければ その他。"""
    use = c.get("use", "") or ""
    kind = c.get("kind", "") or ""
    hits = []
    if "住宅ローン" in use:
        hits.append("住宅ローン")
    if kind in ("クレジット", "立替金") and ("生活費" in use):
        hits.append("生活費")
    elif kind in ("クレジット", "立替金"):
        hits.append("購入")
    elif "生活費" in use:
        hits.append("生活費")
    if any(k in use for k in ("飲食", "交際", "遊興")):
        hits.append("飲食交際遊興費")
    if "借換" in use or "返済のため" in use:
        hits.append("返済")
    if kind in ("保証", "求償") or "保証" in use:
        hits.append("保証")
    if not hits:
        hits.append("その他")
    return hits


def split_address(addr):
    """住所文字列から (〒表記, 残り住所) を切り出す。〒が無ければ ("", addr)。"""
    import re
    m = re.match(r"^(〒?\s*\d{3}-?\d{4})\s*(.*)$", addr or "")
    if m:
        z = m.group(1)
        return (z if z.startswith("〒") else "〒" + z), m.group(2)
    return "", addr or ""


def build_from_template(case, vdir, out):
    """B1105（3行/債権者×16枠）への記入。公租公課は B1106 に回すため除外する。"""
    fillmap = yaml.safe_load((vdir / "fillmap.yaml").read_text(encoding="utf-8"))
    general = [c for c in case.get("creditors", []) if c.get("kind") != "公租公課"]
    kouso_count = len(case.get("creditors", [])) - len(general)

    rows = []
    total = 0
    housing_total = 0
    hosho_total = 0
    for c in general:
        uses = classify_use(c)
        postal, address = split_address(c.get("address", ""))
        note_parts = [c.get("note", "")] if c.get("note") else []
        if ("保証" in uses or "その他" in uses) and c.get("use"):
            note_parts.append(c["use"])
        row = {"name": c.get("name", ""), "postal": postal, "address": address,
               "balance": c.get("balance"), "date_from": c.get("origin_date", ""),
               "date_to": "", "note": "　".join(p for p in note_parts if p)}
        for u in uses:
            row[USE_ROW[u]] = USE_LITERALS[u]
        if c.get("balance") is not None or c.get("sources_ok", True):
            row["chosahyo"] = "□"  # 残高資料がある前提（intakeで残高の出所必須のため）
        rows.append(row)
        b = c.get("balance")
        if isinstance(b, int):
            total += b
            if "住宅ローン" in uses:
                housing_total += b
            if "保証" in uses:
                hosho_total += b

    values = {"fields": {
        "creditors": rows,
        "creditor_count": f"債権者数　{len(rows)}　名",
        "housing_loan_total": housing_total,
        "hosho_total": hosho_total,
    }}
    filler = xlsxlib.XlsxFiller(vdir / "template.xlsx")
    filler.apply(fillmap, values)
    filler.save(out)
    for w in filler.warnings:
        print(f"警告: {w}")
    if kouso_count:
        print(f"注記: 公租公課 {kouso_count} 件はこの書式に載せていない（B1106 公租公課用一覧表で出力する）")
    return {"total_debt": total, "creditor_count": len(rows),
            "unknown_balance": sum(1 for r in rows if r["balance"] is None),
            "kouso_excluded": kouso_count}


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--case", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--court", default="osaka")
    ap.add_argument("--generic", action="store_true", help="裁判所書式を使わず事務所内ドラフト様式で出力")
    args = ap.parse_args()

    case = yaml.safe_load(open(args.case, encoding="utf-8"))
    vdir = None if args.generic else xlsxlib.resolve_form(args.court, "saikensha-ichiran")
    if vdir is None and not args.generic:
        print(xlsxlib.registration_guidance(args.court, "saikensha-ichiran", "債権者一覧表"))
        sys.exit(3)

    stats = build_from_template(case, vdir, args.output) if vdir else build_generic(case, args.output)
    meta_path = Path(args.output + ".meta.json")
    meta_path.write_text(json.dumps({"doc": "saikensha-ichiran", **stats}, ensure_ascii=False, indent=1),
                         encoding="utf-8")
    print(f"出力: {args.output}")
    print(f"債権者 {stats['creditor_count']} 名 / 負債総額 {stats['total_debt']:,} 円"
          + (f" / 残高不明 {stats['unknown_balance']} 名" if stats['unknown_balance'] else ""))


if __name__ == "__main__":
    main()
