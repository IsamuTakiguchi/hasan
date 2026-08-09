#!/usr/bin/env python3
"""資産目録を case.yaml から生成する。

裁判所書式（courts/<court>/forms/shisan-mokuroku/）が登録済みなら fillmap 駆動、
未登録なら --generic で事務所内ドラフト様式を出力する。
<output>.meta.json に資産総額を書き出し、crosscheck_outputs.py が突合に使う。

使い方:
  python3 scripts/build/build_shisan_mokuroku.py --case cases/<id>/case.yaml \
      --output cases/<id>/output/資産目録_<姓>.xlsx [--court osaka] [--generic]
"""
import argparse
import json
import sys
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent))
import xlsxlib  # noqa: E402

import openpyxl  # noqa: E402


def collect_items(assets):
    """(区分, 名称・明細, 評価額, 備考) のリストと合計を返す。"""
    items = []

    def add(cat, name, value, note=""):
        items.append({"cat": cat, "name": name, "value": value, "note": note})

    if assets.get("cash") is not None:
        add("現金", "手持ち現金", assets["cash"])
    for d in assets.get("deposits", []) or []:
        name = f"{d.get('bank', '')}{d.get('branch', '') and ' ' + d['branch']} {d.get('account_type', '')} {d.get('account_no', '')}".strip()
        add("預貯金", name, d.get("balance"), d.get("as_of", "") and f"基準日 {d['as_of']}")
    for x in assets.get("insurance", []) or []:
        add("保険", f"{x.get('company', '')} {x.get('kind', '')}".strip(), x.get("surrender_value"), "解約返戻金見込")
    for x in assets.get("vehicles", []) or []:
        note = ""
        if x.get("loan_balance"):
            note = f"所有権留保ローン残 {x['loan_balance']:,}円"
        add("自動車", x.get("name", ""), x.get("value"), note)
    for x in assets.get("real_estate", []) or []:
        note = f"被担保債権残 {x['mortgage']:,}円" if x.get("mortgage") else ""
        add("不動産", f"{x.get('kind', '')} {x.get('address', '')}".strip(), x.get("value"), note)
    if assets.get("retirement_expectation") is not None:
        add("退職金", "退職金支給見込額", assets["retirement_expectation"])
    for x in assets.get("overpayment", []) or []:
        add("過払金等", x.get("against", ""), x.get("amount"))
    for x in assets.get("other", []) or []:
        add("その他", x.get("name", ""), x.get("value"), x.get("note", ""))

    total = sum(x["value"] for x in items if isinstance(x["value"], int))
    return items, total


def build_generic(case, out):
    meta = case.get("meta", {})
    items, total = collect_items(case.get("assets", {}) or {})
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "資産目録"
    ws["A1"] = "資産目録"
    ws["A1"].font = xlsxlib.FONT_TITLE
    ws["A2"] = f"申立人: {meta.get('applicant', {}).get('name', '')}"
    ws["A2"].font = xlsxlib.FONT

    headers = ["区分", "名称・明細", "評価額(円)", "備考"]
    for i, (h, w) in enumerate(zip(headers, [10, 40, 14, 30]), 1):
        ws.cell(row=4, column=i, value=h)
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    xlsxlib.style_header(ws, 4, len(headers))

    row = 5
    for it in items:
        ws.cell(row=row, column=1, value=it["cat"])
        ws.cell(row=row, column=2, value=it["name"])
        ws.cell(row=row, column=3, value=it["value"])
        ws.cell(row=row, column=4, value=it["note"])
        xlsxlib.style_data(ws, row, len(headers))
        ws[f"C{row}"].number_format = "#,##0"
        row += 1

    ws.cell(row=row, column=2, value="合計")
    ws.cell(row=row, column=3, value=f"=SUM(C5:C{row - 1})")
    xlsxlib.style_data(ws, row, len(headers))
    ws[f"C{row}"].number_format = "#,##0"
    ws[f"B{row}"].font = openpyxl.styles.Font(name="游ゴシック", size=10, bold=True)
    xlsxlib.draft_note(ws, row + 2)

    Path(out).parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return {"total_assets": total, "item_count": len(items)}


def build_from_template(case, vdir, out):
    fillmap = yaml.safe_load((vdir / "fillmap.yaml").read_text(encoding="utf-8"))
    items, total = collect_items(case.get("assets", {}) or {})
    values = {"fields": {
        "applicant_name": case.get("meta", {}).get("applicant", {}).get("name", ""),
        "assets": items,
    }}
    filler = xlsxlib.XlsxFiller(vdir / "template.xlsx")
    filler.apply(fillmap, values)
    filler.save(out)
    for w in filler.warnings:
        print(f"警告: {w}")
    return {"total_assets": total, "item_count": len(items)}


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--case", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--court", default="osaka")
    ap.add_argument("--generic", action="store_true")
    args = ap.parse_args()

    case = yaml.safe_load(open(args.case, encoding="utf-8"))
    vdir = None if args.generic else xlsxlib.resolve_form(args.court, "shisan-mokuroku")
    if vdir is None and not args.generic:
        print(xlsxlib.registration_guidance(args.court, "shisan-mokuroku", "資産目録"))
        sys.exit(3)

    stats = build_from_template(case, vdir, args.output) if vdir else build_generic(case, args.output)
    Path(args.output + ".meta.json").write_text(
        json.dumps({"doc": "shisan-mokuroku", **stats}, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"出力: {args.output}")
    print(f"資産 {stats['item_count']} 項目 / 総額 {stats['total_assets']:,} 円")


if __name__ == "__main__":
    main()
