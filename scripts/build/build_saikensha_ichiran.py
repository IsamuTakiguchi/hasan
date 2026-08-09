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


def build_from_template(case, vdir, out):
    fillmap = yaml.safe_load((vdir / "fillmap.yaml").read_text(encoding="utf-8"))
    meta = case.get("meta", {})
    rows = []
    total = 0
    for c in order_creditors(case.get("creditors", [])):
        rows.append({"no": c.get("no"), "name": c.get("name", ""), "address": c.get("address", ""),
                     "kind": c.get("kind", ""), "origin_date": c.get("origin_date", ""),
                     "use": c.get("use", ""), "principal": c.get("principal"),
                     "balance": c.get("balance"), "last_payment": c.get("last_payment", ""),
                     "guarantor": umu(c.get("guarantor")), "note": c.get("note", "")})
        if isinstance(c.get("balance"), int):
            total += c["balance"]
    values = {"fields": {
        "applicant_name": meta.get("applicant", {}).get("name", ""),
        "creditors": rows,
    }}
    filler = xlsxlib.XlsxFiller(vdir / "template.xlsx")
    filler.apply(fillmap, values)
    filler.save(out)
    for w in filler.warnings:
        print(f"警告: {w}")
    return {"total_debt": total, "creditor_count": len(rows),
            "unknown_balance": sum(1 for r in rows if r["balance"] is None)}


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
